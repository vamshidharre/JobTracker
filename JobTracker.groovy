// ============================================================
// JOB APPLICATION TRACKER — Google Apps Script  v2.1
// Paste into Extensions > Apps Script in your Google Sheet
// Improvements: auto date extraction · improved German classification · smarter quota management
// ============================================================

const SHEET_NAME      = "Applications";
const DASHBOARD_NAME  = "Dashboard";
const LABEL_NAME      = "JobTracker";
const MAX_THREADS     = 20;          // reduced to 20 to stay within daily Gmail quotas
const MAX_PROCESS     = 15;          // reduced to 15 threads per run (each costs ~2-3 API calls)
const QUOTA_PROP_KEY  = "DAILY_GMAIL_CALLS";
const MAX_DAILY_CALLS = 400;         // Conservative daily limit to avoid quota errors

// ── Column positions (1-based) ──────────────────────────────
const COL = {
  THREAD_ID    : 1,
  COMPANY      : 2,
  JOB_TITLE    : 3,
  HR_NAME      : 4,
  HR_EMAIL     : 5,
  HR_PHONE     : 6,
  STATUS       : 7,
  DATE_APPLIED : 8,
  EMAIL_DATE   : 9,
  SUBJECT      : 10,
  GMAIL_LINK   : 11,
  NOTES        : 12,
  ATS_PLATFORM : 13,   // NEW: detected ATS (Greenhouse, Workday, etc.)
  FINGERPRINT  : 14,   // NEW: fuzzy-dedup key
};

// ── Status priority list ────────────────────────────────────
const STATUS_ORDER = [
  "Applied",
  "Application Received",
  "Follow-up Required",
  "Interview Invitation",
  "Interview Scheduled",
  "Technical Assessment",
  "Offer",
  "Rejected",
  "Withdrawn",
];

// ── Status colours ──────────────────────────────────────────
const STATUS_COLORS = {
  "Offer"                 : "#d4edda",
  "Interview Scheduled"   : "#cce5ff",
  "Interview Invitation"  : "#d1ecf1",
  "Technical Assessment"  : "#e2d9f3",
  "Application Received"  : "#fff3cd",
  "Applied"               : "#f8f9fa",
  "Rejected"              : "#f8d7da",
  "Follow-up Required"    : "#ffeeba",
  "Withdrawn"             : "#e2e3e5",
};

// ── Classification rules (EN + DE, ordered by priority) ────
const RULES = [
  {
    status: "Rejected",
    keywords: [
      "unfortunately","regret to inform","not moving forward","not selected",
      "other candidates","decided to proceed with other","position has been filled",
      "you have not been selected","we will not be moving forward","not be progressing",
      "we have decided not to proceed","unsuccessful","did not match","not the right fit",
      "leider","absage","nicht weiterkommen","nicht ausgewählt",
      "haben uns für andere kandidaten entschieden","stelle besetzt",
      "können wir deine bewerbung leider nicht weiter berücksichtigen",
      "wir können ihnen leider keine zusage","bedauerlicherweise","nicht berücksichtigen",
    ],
  },
  {
    status: "Offer",
    keywords: [
      "offer letter","job offer","we'd like to offer","we would like to offer",
      "pleased to offer","formal offer","compensation package","employment offer",
      "extend an offer","offer of employment","accept our offer",
      "angebot","arbeitsvertrag","herzlichen glückwunsch","vertragsangebot",
      "wir freuen uns, ihnen ein angebot","wir möchten ihnen ein angebot",
    ],
  },
  {
    status: "Interview Scheduled",
    keywords: [
      "your interview is scheduled","interview confirmation","calendar invite",
      "zoom link","teams meeting","meeting confirmed","google meet",
      "webex","interview details","ihr gespräch ist bestätigt",
      "terminbestätigung","besprechungsdetails",
    ],
  },
  {
    status: "Technical Assessment",
    keywords: [
      "technical assessment","coding challenge","take-home","homework assignment",
      "hackerrank","codility","testgorilla","assessment link","online test",
      "technische aufgabe","programmieraufgabe","testaufgabe",
    ],
  },
  {
    status: "Follow-up Required",
    keywords: [
      "following up","just checking in","wanted to follow up",
      "any update on my application","checking on the status",
      "additional information required","additional documents required",
      "please provide","need from you","we need","require from you",
      "fragebogen ausfüllen","gehaltsvorstellung mitteilen","fehlende unterlagen",
      "weitere unterlagen","zusätzliche unterlagen","noch benötigen wir",
      "bitte senden sie uns","bitte übermitteln sie","bitte reichen sie nach",
      "noch fehlende","ergänzende unterlagen einreichen","dokumente nachreichen",
      "questionnaire","salary expectation required","please send us","please submit",
      "missing documents","documents are missing","please complete",
      "complete your application","additional steps required","action required",
      "information is missing","upload additional","provide additional",
    ],
  },
  {
    status: "Interview Invitation",
    keywords: [
      "invite you to interview","schedule an interview","schedule a call",
      "phone screen","we'd like to invite","would you be available",
      "let's connect","vorstellungsgespräch","einladung zum gespräch",
      "telefoninterview","wir würden uns freuen","können wir einen termin",
      "wir laden sie ein","terminvorschlag",
    ],
  },
  {
    status: "Application Received",
    keywords: [
      "application received","thank you for applying","thank you for your application",
      "we have received your application","successfully submitted",
      "application submitted","your cv has been received","received your cv",
      "bewerbungseingang","eingang deiner bewerbung","eingang ihrer bewerbung",
      "wir haben deine bewerbung erhalten","wir haben ihre bewerbung erhalten",
      "unterlagen haben wir erhalten","die unterlagen haben wir erhalten",
      "bestätigung deiner bewerbung","bestätigung ihrer bewerbung",
      "vielen dank für deine bewerbung","vielen dank für ihre bewerbung",
      "vielen dank für die bewerbung","danke für die bewerbung",
      "bewerbung wird eingehend geprüft","wir melden uns zeitnah",
    ],
  },
];

// ── Known ATS platforms ─────────────────────────────────────
const ATS_MAP = {
  "greenhouse.io"     : "Greenhouse",
  "lever.co"          : "Lever",
  "workday.com"       : "Workday",
  "myworkdayjobs.com" : "Workday",
  "smartrecruiters.com": "SmartRecruiters",
  "taleo.net"         : "Taleo",
  "icims.com"         : "iCIMS",
  "recruitee.com"     : "Recruitee",
  "jobvite.com"       : "Jobvite",
  "personio.de"       : "Personio",
  "softgarden.de"     : "Softgarden",
  "stepstone.de"      : "StepStone",
  "xing.com"          : "XING",
  "linkedin.com"      : "LinkedIn",
};

// ── Extraction patterns ─────────────────────────────────────
const JOB_PATTERNS = [
  // "Position: Senior Engineer" / "Role: Data Analyst"
  /(?:position|role|job\s*title|stelle|position\s*titel)[:\s]+([^\n,|]{3,70})/i,
  // "applying for the Senior Engineer role"
  /applying\s+for\s+(?:the\s+)?([^\n,|]{3,70}?)(?:\s+(?:position|role|at|@))/i,
  // "bewirbt sich als / für die Stelle als"
  /(?:als|für\s+(?:die\s+)?stelle\s+(?:als\s+)?|stelle:)\s+([^\n,|]{3,70})/i,
  // Subject: "Re: Application for Senior Engineer – Acme"
  /(?:application\s+for|bewerbung\s+(?:als|für|auf))[:\s–-]+([^\n,|–-]{3,70})/i,
  // "the [Job Title] position/opportunity"
  /(?:the\s+)([A-Z][a-zA-ZÄÖÜäöü\s\-/]{3,50})\s+(?:position|role|opportunity|stelle)/,
];

const COMPANY_PATTERNS = [
  // "team at Acme Corp"
  /(?:team\s+at|working\s+at|join|joining)\s+([A-Z][a-zA-Z0-9\s&.,\-]{2,50}?)(?:\s|,|\.|\n|$)/i,
  // "at / with / bei Acme – we are"
  /(?:\bat\b|\bwith\b|\bbei\b|\bvon\b)\s+([A-Z][a-zA-Z0-9\s&.,\-]{2,50})(?:\s+(?:we\b|is\b|has\b|are\b|team\b|AG\b|GmbH\b|SE\b))/,
  // "Acme Corp is hiring / is looking"
  /^([A-Z][a-zA-Z0-9\s&.,\-]{2,50})\s+(?:is\s+hiring|is\s+looking|has\s+reviewed)/m,
  // German: "Acme GmbH / AG / SE / KG"
  /([A-ZÄÖÜ][a-zA-ZÄÖÜäöü0-9\s&.,\-]{2,50}(?:\s(?:GmbH|AG|SE|KG|e\.V\.|Inc\.|Ltd\.|Corp\.)))/,
];

const PHONE_RE    = /(?:\+?[\d][\d\s\-().]{7,18}\d)/g;
const HR_NAME_RE  = /(?:regards|sincerely|best|cheers|thanks|mit\s+freundlichen\s+grüßen|viele\s+grüße|herzliche\s+grüße)[,.\s]*\r?\n+\s*([A-ZÄÖÜ][a-zA-ZÄÖÜäöü\-]{1,30}(?:\s[A-ZÄÖÜ][a-zA-ZÄÖÜäöü\-]{1,30}){0,2})/i;


// ============================================================
// QUOTA MANAGEMENT
// ============================================================
function checkAndUpdateQuota(callsNeeded) {
  const props = PropertiesService.getScriptProperties();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  const lastDate = props.getProperty("QUOTA_DATE");
  let dailyCalls = parseInt(props.getProperty(QUOTA_PROP_KEY) || "0");

  // Reset counter if it's a new day
  if (lastDate !== today) {
    dailyCalls = 0;
    props.setProperty("QUOTA_DATE", today);
  }

  // Check if we have quota available
  if (dailyCalls + callsNeeded > MAX_DAILY_CALLS) {
    Logger.log(`⛔ Daily quota limit reached: ${dailyCalls}/${MAX_DAILY_CALLS} calls used. Skipping this run.`);
    return false;
  }

  // Update the counter
  dailyCalls += callsNeeded;
  props.setProperty(QUOTA_PROP_KEY, dailyCalls.toString());
  Logger.log(`📊 Quota usage: ${dailyCalls}/${MAX_DAILY_CALLS} calls today`);

  return true;
}


// ============================================================
// MAIN ENTRY — triggered every 30 minutes
// ============================================================
function runTracker() {
  // Estimate API calls needed: 1 search + (MAX_PROCESS × 3 calls per thread)
  const estimatedCalls = 1 + (MAX_PROCESS * 3);

  // Check quota before proceeding
  if (!checkAndUpdateQuota(estimatedCalls)) {
    return; // Skip this run if quota exceeded
  }

  const sheet    = getOrCreateSheet();
  const existing = getExistingIndex(sheet);   // { threadId: rowNum } + { fingerprint: rowNum }
  const threads  = getJobEmailThreads();

  let added = 0, updated = 0, dupeSkipped = 0, skipped = 0;
  const processed = [];   // successfully processed threads (for labelling)

  for (let i = 0; i < threads.length; i++) {
    // Per-run safety cap to stay within quotas
    if (added + updated + dupeSkipped >= MAX_PROCESS) {
      skipped = threads.length - i;
      Logger.log(`⚠️ Reached per-run cap (${MAX_PROCESS}). ${skipped} threads deferred to next run.`);
      break;
    }

    const thread = threads[i];
    const id     = thread.getId();

    // Skip threads we already have in the sheet (belt-and-suspenders with -label: query)
    if (existing.byThreadId[id]) {
      // Still label it so the search excludes it next time
      processed.push(thread);
      continue;
    }

    try {
      const messages = thread.getMessages();
      const latest   = messages[messages.length - 1];
      const body     = latest.getPlainBody();
      const subject  = latest.getSubject();
      const combined = subject + "\n" + body;
      const status   = classifyEmail(combined);
      const data     = extractData(latest, thread, status);
      const fp       = data.fingerprint;

      if (fp && existing.byFingerprint[fp]) {
        // Different thread but same company+title — likely duplicate application channel
        const targetRow = existing.byFingerprint[fp];
        updateStatusIfNewer(sheet, targetRow, status, data);
        appendDuplicateNote(sheet, targetRow, id, data);
        dupeSkipped++;
      } else {
        appendRow(sheet, data);
        const newRow = sheet.getLastRow();
        existing.byThreadId[id] = newRow;
        if (fp) existing.byFingerprint[fp] = newRow;
        added++;
      }
      processed.push(thread);
    } catch (e) {
      // Catch quota errors gracefully and stop processing
      if (e.message && e.message.includes("Service invoked too many times")) {
        Logger.log(`⛔ Quota limit hit after processing ${i} threads. Stopping early.`);
        break;
      }
      Logger.log(`⚠️ Error processing thread ${id}: ${e.message}`);
    }
  }

  Logger.log(`Run complete — ${added} added, ${updated} updated, ${dupeSkipped} duplicates merged, ${skipped} deferred.`);

  // Only label threads that were successfully processed
  if (processed.length > 0) {
    labelProcessedThreads(processed);
  }

  // Only refresh dashboard if something changed
  if (added > 0 || updated > 0 || dupeSkipped > 0) {
    refreshDashboard();
  }
}


// ============================================================
// EMAIL SEARCH
// ============================================================
function getJobEmailThreads() {
  const query = [
    "(",
    "subject:(application OR applied OR interview OR rejection OR offer OR assessment)",
    "OR subject:(bewerbung OR vorstellungsgespräch OR absage OR angebot OR einladung)",
    "OR from:(greenhouse.io OR lever.co OR workday.com OR personio.de OR softgarden.de)",
    ")",
    "newer_than:30d",             // reduced from 90d — old threads are already tracked
    `-label:${LABEL_NAME}`,       // ← KEY FIX: skip threads already processed & labelled
  ].join(" ");
  return GmailApp.search(query, 0, MAX_THREADS);
}


// ============================================================
// CLASSIFICATION
// ============================================================
function classifyEmail(text) {
  const lower = text.toLowerCase();
  for (const rule of RULES) {
    if (rule.keywords.some(kw => lower.includes(kw))) return rule.status;
  }
  return "Applied";
}

function isNewerStatus(current, incoming) {
  // "Rejected" and "Withdrawn" are terminal — never downgrade from Offer
  const ci = STATUS_ORDER.indexOf(current);
  const ii = STATUS_ORDER.indexOf(incoming);
  if (ci === -1) return true;
  if (ii === -1) return false;
  // Special: don't overwrite Offer with Rejected (could be a different role email)
  if (current === "Offer" && incoming === "Rejected") return false;
  return ii > ci;
}


// ============================================================
// DATA EXTRACTION
// ============================================================
function extractData(message, thread, status) {
  const subject  = message.getSubject();
  const body     = message.getPlainBody();
  const from     = message.getFrom();
  const date     = message.getDate();
  const threadId = thread.getId();

  const hrEmail   = extractEmail(from);
  const hrName    = extractHrName(from, body);
  const phone     = extractPhone(body);
  const ats       = detectAts(hrEmail, body);
  const company   = extractCompany(subject, body, hrEmail, ats);
  const jobTitle  = extractJobTitle(subject, body);
  const emailDate = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
  const fp        = buildFingerprint(company, jobTitle);

  // Try to extract date applied from email content or use first message date in thread
  const dateApplied = extractDateApplied(body, subject, thread);

  return {
    threadId,
    company,
    jobTitle,
    hrName,
    hrEmail,
    hrPhone     : phone,
    status,
    dateApplied,
    emailDate,
    subject,
    gmailLink   : `https://mail.google.com/mail/u/0/#inbox/${threadId}`,
    notes       : "",
    atsPlatform : ats,
    fingerprint : fp,
  };
}

// ── Helpers ─────────────────────────────────────────────────

function extractEmail(from) {
  const m = from.match(/<([^>]+)>/);
  return m ? m[1].toLowerCase().trim() : from.toLowerCase().trim();
}

function extractDateApplied(body, subject, thread) {
  const searchText = subject + "\n" + body.substring(0, 1000);

  // Patterns to extract application date from email text
  // EN: "You applied on March 15, 2024" / "submitted on 15/03/2024"
  // DE: "Sie haben sich am 15.03.2024 beworben" / "eingegangen am 15.03.2024"
  const datePatterns = [
    // ISO format: 2024-03-15
    /(?:applied|submitted|beworben|eingegangen)[\s\w]*(?:am|on|:)?\s*(\d{4}[-./]\d{1,2}[-./]\d{1,2})/i,
    // European format: 15.03.2024 or 15/03/2024
    /(?:applied|submitted|beworben|eingegangen)[\s\w]*(?:am|on|:)?\s*(\d{1,2}[./-]\d{1,2}[./-]\d{4})/i,
    // Date in "Application received on..." context
    /(?:received|eingegangen)[\s\w]*(?:am|on|:)?\s*(\d{1,2}[./-]\d{1,2}[./-]\d{4})/i,
    /(?:received|eingegangen)[\s\w]*(?:am|on|:)?\s*(\d{4}[-./]\d{1,2}[-./]\d{1,2})/i,
  ];

  for (const pattern of datePatterns) {
    const match = searchText.match(pattern);
    if (match) {
      try {
        const dateStr = match[1];
        const parsedDate = parseFlexibleDate(dateStr);
        if (parsedDate) return parsedDate;
      } catch (e) {
        // continue to next pattern
      }
    }
  }

  // Fallback: Use the first message date in the thread (likely when user applied)
  try {
    const messages = thread.getMessages();
    if (messages.length > 0) {
      const firstDate = messages[0].getDate();
      return Utilities.formatDate(firstDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }
  } catch (e) {
    Logger.log(`Warning: Could not get thread messages for date extraction: ${e.message}`);
  }

  return ""; // Return empty if extraction fails
}

// Parse dates in various formats: 2024-03-15, 15.03.2024, 15/03/2024, etc.
function parseFlexibleDate(dateStr) {
  if (!dateStr) return "";

  // Normalize separators to hyphens
  const normalized = dateStr.replace(/[./]/g, "-");
  const parts = normalized.split("-");

  if (parts.length !== 3) return "";

  let year, month, day;

  // Detect format: YYYY-MM-DD vs DD-MM-YYYY
  if (parts[0].length === 4) {
    // ISO format: YYYY-MM-DD
    year = parseInt(parts[0]);
    month = parseInt(parts[1]);
    day = parseInt(parts[2]);
  } else {
    // European format: DD-MM-YYYY
    day = parseInt(parts[0]);
    month = parseInt(parts[1]);
    year = parseInt(parts[2]);
  }

  // Validate
  if (year < 2000 || year > 2100) return "";
  if (month < 1 || month > 12) return "";
  if (day < 1 || day > 31) return "";

  // Create date and format as YYYY-MM-DD
  try {
    const date = new Date(year, month - 1, day);
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
  } catch (e) {
    return "";
  }
}


function extractHrName(from, body) {
  // Clean display name from "Name <email>" format
  const fromName = from.replace(/<[^>]+>/, "").replace(/"/g, "").trim();
  // Filter out ATS system names (e.g. "Greenhouse", "Workday Notifications")
  const isSystemSender = /notification|noreply|no-reply|donotreply|system|alert|team|recruiting/i.test(fromName);
  if (fromName && fromName.length > 1 && !isSystemSender) return fromName;
  // Fallback: signature block
  const m = body.match(HR_NAME_RE);
  return m ? m[1].trim() : "";
}

function extractPhone(body) {
  const matches = body.match(PHONE_RE);
  if (!matches) return "";
  // Return the longest match (most likely a full number)
  return matches.sort((a, b) => b.length - a.length)[0].trim();
}

function detectAts(hrEmail, body) {
  for (const [domain, name] of Object.entries(ATS_MAP)) {
    if (hrEmail.includes(domain) || body.toLowerCase().includes(domain)) return name;
  }
  return "";
}

function extractCompany(subject, body, hrEmail, ats) {
  // 1. Use HR email domain (skip known ATS / free email domains)
  const skipDomains = ["gmail","yahoo","hotmail","outlook","gmx","web.de","t-online",
                       "greenhouse","lever","workday","smartrecruiters","taleo","icims",
                       "recruitee","jobvite","personio","softgarden","stepstone","linkedin","xing"];
  if (hrEmail) {
    const domain = hrEmail.split("@")[1] || "";
    const domainBase = domain.split(".")[0].toLowerCase();
    if (domain && !skipDomains.some(s => domain.includes(s))) {
      return capitalize(domainBase);
    }
  }
  // 2. Regex patterns in subject + first 600 chars of body
  const searchText = subject + "\n" + body.substring(0, 600);
  for (const pat of COMPANY_PATTERNS) {
    const m = searchText.match(pat);
    if (m) return m[1].trim().replace(/\s+/g, " ");
  }
  // 3. If ATS detected, subject sometimes has "CompanyName via Greenhouse"
  if (ats) {
    const viaMatch = subject.match(/^([A-Z][a-zA-Z0-9\s&.,\-]{2,50})\s+via\s+/i);
    if (viaMatch) return viaMatch[1].trim();
  }
  return "";
}

function extractJobTitle(subject, body) {
  const searchText = subject + "\n" + body.substring(0, 800);
  for (const pat of JOB_PATTERNS) {
    const m = searchText.match(pat);
    if (m) {
      return m[1]
        .trim()
        .replace(/\s+/g, " ")
        .replace(/[–\-]\s*$/, "")  // strip trailing dash
        .substring(0, 80);
    }
  }
  // Fallback: clean up subject line
  return subject
    .replace(/^re:\s*/i, "")
    .replace(/^fwd?:\s*/i, "")
    .replace(/application\s*(for)?\s*/i, "")
    .replace(/bewerbung\s*(als|für|auf)?\s*/i, "")
    .trim()
    .substring(0, 80);
}

// Build a normalised fingerprint for fuzzy dedup: "acme|senior engineer"
function buildFingerprint(company, jobTitle) {
  if (!company && !jobTitle) return "";
  const norm = s => s.toLowerCase()
    .replace(/\b(the|a|an|der|die|das|ein|eine)\b/g, "")
    .replace(/[^a-z0-9äöüß]+/g, "")
    .trim();
  return `${norm(company)}|${norm(jobTitle)}`;
}

function capitalize(str) {
  return str.charAt(0).toUpperCase() + str.slice(1);
}


// ============================================================
// SHEET HELPERS
// ============================================================
function getOrCreateSheet() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const headers = [
      "Thread ID","Company","Job Title","HR Name","HR Email",
      "HR Phone","Status","Date Applied","Email Date","Subject",
      "Gmail Link","Notes","ATS Platform","Fingerprint",
    ];
    const hdr = sheet.getRange(1, 1, 1, headers.length);
    hdr.setValues([headers]);
    hdr.setFontWeight("bold").setBackground("#1a1a2e").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(COL.GMAIL_LINK,   200);
    sheet.setColumnWidth(COL.SUBJECT,      260);
    sheet.setColumnWidth(COL.COMPANY,      160);
    sheet.setColumnWidth(COL.JOB_TITLE,    220);
    sheet.setColumnWidth(COL.FINGERPRINT,  1);    // hide visually
    sheet.setColumnWidth(COL.THREAD_ID,    1);
  }
  return sheet;
}

// Returns { byThreadId: {id: row}, byFingerprint: {fp: row} }
function getExistingIndex(sheet) {
  const data = sheet.getDataRange().getValues();
  const idx  = { byThreadId: {}, byFingerprint: {} };
  for (let i = 1; i < data.length; i++) {
    const row = i + 1;
    const tid = data[i][COL.THREAD_ID   - 1];
    const fp  = data[i][COL.FINGERPRINT - 1];
    if (tid) idx.byThreadId[tid]  = row;
    if (fp)  idx.byFingerprint[fp] = row;
  }
  return idx;
}

function appendRow(sheet, d) {
  sheet.appendRow([
    d.threadId, d.company, d.jobTitle, d.hrName, d.hrEmail,
    d.hrPhone, d.status, d.dateApplied, d.emailDate,
    d.subject, d.gmailLink, d.notes, d.atsPlatform, d.fingerprint,
  ]);
  colorStatusCell(sheet, sheet.getLastRow(), d.status);
}

function updateStatusIfNewer(sheet, rowNum, newStatus, d) {
  const current = sheet.getRange(rowNum, COL.STATUS).getValue();
  if (!isNewerStatus(current, newStatus)) return;

  sheet.getRange(rowNum, COL.STATUS).setValue(newStatus);
  sheet.getRange(rowNum, COL.EMAIL_DATE).setValue(d.emailDate);
  colorStatusCell(sheet, rowNum, newStatus);

  // Update Date Applied if it's currently empty and we have a value
  const existingDateApplied = sheet.getRange(rowNum, COL.DATE_APPLIED).getValue();
  if (!existingDateApplied && d.dateApplied) {
    sheet.getRange(rowNum, COL.DATE_APPLIED).setValue(d.dateApplied);
  }

  const existingNotes = sheet.getRange(rowNum, COL.NOTES).getValue();
  sheet.getRange(rowNum, COL.NOTES).setValue(
    (existingNotes ? existingNotes + " | " : "") +
    `→ ${newStatus} on ${d.emailDate}`
  );
}

function appendDuplicateNote(sheet, rowNum, newThreadId, d) {
  const existingNotes = sheet.getRange(rowNum, COL.NOTES).getValue();
  sheet.getRange(rowNum, COL.NOTES).setValue(
    (existingNotes ? existingNotes + " | " : "") +
    `Duplicate thread ${newThreadId} via ${d.atsPlatform || "unknown"} on ${d.emailDate}`
  );
}

function colorStatusCell(sheet, row, status) {
  const color = STATUS_COLORS[status] || "#ffffff";
  sheet.getRange(row, COL.STATUS).setBackground(color).setFontWeight("bold");
}

function labelProcessedThreads(threads) {
  let label = GmailApp.getUserLabelByName(LABEL_NAME);
  if (!label) label = GmailApp.createLabel(LABEL_NAME);
  threads.forEach(t => t.addLabel(label));
}


// ============================================================
// DASHBOARD SHEET
// ============================================================
function refreshDashboard() {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet  = ss.getSheetByName(SHEET_NAME);
  if (!dataSheet) return;

  let dash = ss.getSheetByName(DASHBOARD_NAME);
  if (!dash) {
    dash = ss.insertSheet(DASHBOARD_NAME, 0);   // first tab
  }
  dash.clearContents();
  dash.clearFormats();

  const allData = dataSheet.getDataRange().getValues().slice(1);
  const today   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  // ── Aggregate ───────────────────────────────────────────
  const statusCounts   = {};
  const atsCounts      = {};
  const recentActivity = [];

  STATUS_ORDER.forEach(s => statusCounts[s] = 0);

  allData.forEach(row => {
    const status    = row[COL.STATUS      - 1] || "Applied";
    const ats       = row[COL.ATS_PLATFORM - 1] || "Direct";
    const emailDate = row[COL.EMAIL_DATE  - 1];
    statusCounts[status] = (statusCounts[status] || 0) + 1;
    atsCounts[ats]       = (atsCounts[ats]       || 0) + 1;
    if (emailDate === today) recentActivity.push(row);
  });

  const total      = allData.length;
  const interviews = allData.filter(r => (r[COL.STATUS-1]||"").includes("Interview")).length;
  const offers     = statusCounts["Offer"]    || 0;
  const rejected   = statusCounts["Rejected"] || 0;
  const responseRate = total > 0 ? Math.round(((total - (statusCounts["Applied"] || 0)) / total) * 100) : 0;

  // ── Layout ──────────────────────────────────────────────
  // Title
  dash.setColumnWidth(1, 200);
  dash.setColumnWidth(2, 140);
  dash.setColumnWidth(3, 140);
  dash.setColumnWidth(4, 140);
  dash.setColumnWidth(5, 180);

  setCell(dash, 1, 1, "📋 JOB APPLICATION TRACKER", "#1a1a2e", "#ffffff", 16, true, 5);
  setCell(dash, 2, 1, `Last updated: ${today}`, "#f0f4ff", "#555", 10, false, 5);

  // KPI row
  setCell(dash, 4, 1, "TOTAL APPLICATIONS", "#1a1a2e", "#7dd3fc", 9, true);
  setCell(dash, 4, 2, "INTERVIEWS",         "#1a1a2e", "#7dd3fc", 9, true);
  setCell(dash, 4, 3, "OFFERS",             "#1a1a2e", "#7dd3fc", 9, true);
  setCell(dash, 4, 4, "REJECTIONS",         "#1a1a2e", "#7dd3fc", 9, true);
  setCell(dash, 4, 5, "RESPONSE RATE",      "#1a1a2e", "#7dd3fc", 9, true);

  setCell(dash, 5, 1, total,          "#0f172a", "#ffffff", 28, true);
  setCell(dash, 5, 2, interviews,     "#0f172a", "#86efac", 28, true);
  setCell(dash, 5, 3, offers,         "#0f172a", "#fde68a", 28, true);
  setCell(dash, 5, 4, rejected,       "#0f172a", "#fca5a5", 28, true);
  setCell(dash, 5, 5, responseRate + "%", "#0f172a", "#c4b5fd", 28, true);

  dash.setRowHeight(5, 60);

  // Status breakdown
  setCell(dash, 7, 1, "STATUS BREAKDOWN", "#1a1a2e", "#ffffff", 11, true, 2);
  let sRow = 8;
  setCell(dash, sRow, 1, "Status",  "#334155", "#94a3b8", 9, true);
  setCell(dash, sRow, 2, "Count",   "#334155", "#94a3b8", 9, true);
  setCell(dash, sRow, 3, "Bar",     "#334155", "#94a3b8", 9, true);
  sRow++;
  STATUS_ORDER.forEach(status => {
    const count = statusCounts[status] || 0;
    if (count === 0) return;
    const bar   = "█".repeat(Math.min(count, 30));
    setCell(dash, sRow, 1, status, STATUS_COLORS[status] || "#f8f9fa", "#1a1a2e", 10, false);
    setCell(dash, sRow, 2, count,  STATUS_COLORS[status] || "#f8f9fa", "#1a1a2e", 10, true);
    setCell(dash, sRow, 3, bar,    STATUS_COLORS[status] || "#f8f9fa", "#475569", 8, false);
    sRow++;
  });

  // ATS breakdown
  setCell(dash, 7, 5, "SOURCE / ATS PLATFORM", "#1a1a2e", "#ffffff", 11, true);
  let aRow = 8;
  setCell(dash, aRow, 5, "Platform", "#334155", "#94a3b8", 9, true);
  aRow++;
  Object.entries(atsCounts)
    .sort((a, b) => b[1] - a[1])
    .forEach(([ats, cnt]) => {
      setCell(dash, aRow, 5, `${ats}: ${cnt}`, "#1e293b", "#e2e8f0", 10, false);
      aRow++;
    });

  // Today's activity
  const actStart = Math.max(sRow, aRow) + 2;
  setCell(dash, actStart, 1, `TODAY'S ACTIVITY (${today})`, "#1a1a2e", "#ffffff", 11, true, 5);
  if (recentActivity.length === 0) {
    setCell(dash, actStart + 1, 1, "No new emails processed today.", "#f8fafc", "#64748b", 10, false, 5);
  } else {
    let tRow = actStart + 1;
    setCell(dash, tRow, 1, "Company",   "#334155", "#94a3b8", 9, true);
    setCell(dash, tRow, 2, "Job Title", "#334155", "#94a3b8", 9, true);
    setCell(dash, tRow, 3, "Status",    "#334155", "#94a3b8", 9, true);
    setCell(dash, tRow, 4, "Email",     "#334155", "#94a3b8", 9, true);
    tRow++;
    recentActivity.forEach(r => {
      const st = r[COL.STATUS - 1] || "Applied";
      setCell(dash, tRow, 1, r[COL.COMPANY   - 1], "#f8fafc", "#1e293b", 10);
      setCell(dash, tRow, 2, r[COL.JOB_TITLE - 1], "#f8fafc", "#1e293b", 10);
      setCell(dash, tRow, 3, st,                    STATUS_COLORS[st] || "#f8fafc", "#1e293b", 10, true);
      setCell(dash, tRow, 4, r[COL.HR_EMAIL  - 1], "#f8fafc", "#64748b", 9);
      tRow++;
    });
  }

  Logger.log("Dashboard refreshed.");
}

// Helper: write a cell with background + font styling, optional merge
function setCell(sheet, row, col, value, bg, fg, fontSize, bold, mergeCount) {
  const cell = sheet.getRange(row, col);
  cell.setValue(value)
      .setBackground(bg || "#ffffff")
      .setFontColor(fg || "#000000")
      .setFontSize(fontSize || 10)
      .setFontWeight(bold ? "bold" : "normal")
      .setVerticalAlignment("middle")
      .setWrap(false);
  if (mergeCount && mergeCount > 1) {
    sheet.getRange(row, col, 1, mergeCount).merge();
  }
}


// ============================================================
// TRIGGER SETUP
// ============================================================
function installTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "runTracker")
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger("runTracker")
    .timeBased()
    .everyHours(1)          // changed to 1 hour (from 30 min) to reduce quota usage
    .create();

  Logger.log("✅ Trigger installed — runTracker fires every hour.");
}

function installDailySummaryTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "sendDailySummary")
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger("sendDailySummary")
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  Logger.log("✅ Daily summary trigger installed — fires at 8am every day.");
}


// ============================================================
// DAILY SUMMARY EMAIL
// ============================================================
function sendDailySummary() {
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues().slice(1);
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const todayRows = data.filter(r => r[COL.EMAIL_DATE - 1] === today);

  if (todayRows.length === 0) return;

  const counts = {};
  todayRows.forEach(r => {
    const s = r[COL.STATUS - 1] || "Unknown";
    counts[s] = (counts[s] || 0) + 1;
  });

  const lines      = Object.entries(counts).map(([s, n]) => `  • ${s}: ${n}`).join("\n");
  const total      = data.length;
  const interviews = data.filter(r => (r[COL.STATUS-1]||"").includes("Interview")).length;
  const offers     = data.filter(r => r[COL.STATUS-1] === "Offer").length;
  const rejected   = data.filter(r => r[COL.STATUS-1] === "Rejected").length;

  const body = `
Job Application Tracker — Daily Summary (${today})
${"─".repeat(50)}

New emails processed today: ${todayRows.length}
${lines}

OVERALL STATS
  • Total applications tracked : ${total}
  • Interviews                 : ${interviews}
  • Rejections                 : ${rejected}
  • Offers                     : ${offers}

View your sheet → ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}
  `.trim();

  GmailApp.sendEmail(
    Session.getActiveUser().getEmail(),
    `📋 Job Tracker — ${today} (${todayRows.length} new)`,
    body
  );
  Logger.log("Daily summary sent.");
}