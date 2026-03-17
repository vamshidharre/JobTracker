# 📋 Job Application Tracker — Google Apps Script

> Automatically track every job application directly from your Gmail into a fully organised Google Sheet — with a live dashboard, analytics, and auto-reply drafts. 100% free, runs inside your own Google account.

---

## 🚀 What It Does

Stop manually logging every application. This script scans your Gmail every 15 minutes, extracts key details from job-related emails, and keeps a live spreadsheet up to date — including status changes, HR contact info, and interview progress.

Supports both **English and German** emails out of the box.

---

## ✨ Features

### 📥 Email Processing
- Automatically scans Gmail every 15 minutes
- Checks the last 90 days, up to 150 threads per run
- Filters by job-related keywords in English and German
- Labels all processed threads with `JobTracker` inside Gmail

### 🧠 Smart Classification
9 status levels tracked in priority order:

| Status | Description |
|---|---|
| Applied | Initial outreach sent |
| Application Received | Confirmation email from company |
| Follow-up Required | You sent a follow-up |
| Interview Invitation | Company invited you to interview |
| Interview Scheduled | Interview time confirmed |
| Technical Assessment | Coding challenge / take-home sent |
| Offer | Job offer received |
| Rejected | Application unsuccessful |
| Withdrawn | You withdrew your application |

> Status only ever moves **forward** — it will never downgrade (e.g. an Offer will never be overwritten by a Rejection email from a different role).

### 🔍 Data Extraction
Automatically pulls from each email:
- Company name (from domain, body text, German legal forms like GmbH / AG)
- Job title (from subject line and body, EN + DE patterns)
- HR name (filters out ATS bots and noreply senders)
- HR email address and phone number
- Direct Gmail link to the thread

### 🔁 Smart Duplicate Detection
- **Thread ID matching** — recognises the same Gmail thread on repeat runs
- **Fuzzy fingerprint matching** — detects the same job applied via different channels (e.g. LinkedIn and company website) and merges them into one row instead of creating duplicates

### 📝 Auto-Reply Drafts
When an interview invite, scheduled interview, technical assessment, or offer arrives, the script automatically creates a polished Gmail draft ready for you to review and send:
- Detects email language (EN or DE) automatically
- Addresses the recruiter by name when available
- Includes company name and job title in the reply
- Checks for existing drafts to avoid duplicates
- Toggle off anytime: set `AUTO_DRAFT_ENABLED = false`

### 📋 Dashboard Tab
A live summary sheet that refreshes every 15 minutes:
- **KPI row** — Total Applications, Interviews, Offers, Rejections, Response Rate
- **Status breakdown** with inline bar charts
- **ATS / source platform** breakdown
- **Today's activity** table

### 📊 Analytics Tab
A dedicated analytics sheet with:
- **Conversion funnel** — Applied → Received → Interview → Offer, with percentage rates
- **Weekly volume chart** — last 12 weeks of activity with colour-coded trend bars
- **Day-of-week heatmap** — shows which days you receive responses most
- **Platform performance table** — per ATS: applications, interviews, interview rate, and offers

### 📧 Daily Summary Email
- Sent to your own Gmail at 8am every day
- Shows new emails processed, status breakdown, and overall stats
- Includes a direct link back to your sheet

---

## 🛠️ Setup

### Step 1 — Create a Google Sheet
Open [Google Sheets](https://sheets.google.com) and create a new blank spreadsheet.

### Step 2 — Open Apps Script
Inside the sheet, go to **Extensions → Apps Script**.

### Step 3 — Paste the Script
Delete any existing code and paste the full contents of `JobTracker.gs` into the editor.

### Step 4 — Install Triggers
In the Apps Script editor, run these two functions **once** by selecting them from the dropdown and clicking ▶ Run:

```
installTrigger()             → starts the 15-minute email scan
installDailySummaryTrigger() → enables the 8am daily summary email
```

### Step 5 — Authorise Permissions
Google will ask you to grant permissions. The script needs access to:
- **Gmail** — to read job-related emails and create draft replies
- **Google Sheets** — to write data into your spreadsheet

> ⚠️ All data stays entirely within your own Google account. Nothing is sent to any external server.

### Step 6 — Run Manually (First Time)
Select `runTracker` from the dropdown and click ▶ Run to do an immediate first scan. After that, it runs automatically every 15 minutes.

---

## ⚙️ Configuration

All settings are at the top of the script file:

| Constant | Default | Description |
|---|---|---|
| `SHEET_NAME` | `"Applications"` | Name of the data sheet tab |
| `DASHBOARD_NAME` | `"Dashboard"` | Name of the dashboard tab |
| `ANALYTICS_NAME` | `"Analytics"` | Name of the analytics tab |
| `LABEL_NAME` | `"JobTracker"` | Gmail label applied to processed threads |
| `MAX_THREADS` | `150` | Max emails scanned per run |
| `AUTO_DRAFT_ENABLED` | `true` | Set to `false` to disable auto-drafts |
| `DRAFT_TRIGGER_STATUSES` | Interview, Offer, Assessment | Which statuses trigger a draft |

### Scan more history
Change `newer_than:90d` in `getJobEmailThreads()` to e.g. `newer_than:180d`.

### One-time full mailbox scan
Temporarily remove `newer_than:90d` and set `MAX_THREADS = 500`, run once manually, then restore the defaults.

> ⚠️ Google Apps Script has a 6-minute execution limit per run. Very large scans may time out.

---

## 🗂️ Sheet Structure

The **Applications** sheet contains these columns:

| Column | Description |
|---|---|
| Thread ID | Internal Gmail thread ID |
| Company | Extracted company name |
| Job Title | Extracted job title |
| HR Name | Recruiter / sender name |
| HR Email | Recruiter email address |
| HR Phone | Phone number if found in signature |
| Status | Current application status (colour-coded) |
| Date Applied | Fill in manually |
| Email Date | Date of the most recent relevant email |
| Subject | Email subject line |
| Gmail Link | Direct link to the Gmail thread |
| Notes | Auto-updated status history + duplicate notes |
| ATS Platform | Detected platform (Greenhouse, Workday, etc.) |
| Fingerprint | Internal dedup key (company + title) |

---

## 🌍 Supported Languages

| Feature | English | German |
|---|---|---|
| Status classification | ✅ | ✅ |
| Company extraction | ✅ | ✅ (GmbH, AG, SE, KG) |
| Job title extraction | ✅ | ✅ |
| HR name extraction | ✅ | ✅ |
| Auto-reply drafts | ✅ | ✅ |

---

## 🏢 Detected ATS Platforms

Greenhouse · Lever · Workday · SmartRecruiters · Taleo · iCIMS · Recruitee · Jobvite · Personio · Softgarden · StepStone · XING · LinkedIn

---

## 🔒 Privacy & Security

- The script runs entirely inside **your own Google account**
- No data is sent to any external server or third party
- No API keys or paid services required
- You can view all permissions granted under **Extensions → Apps Script → Project Settings**
- Deleting the script removes all automation immediately

---

## 🐛 Troubleshooting

**No rows appearing after first run**
Make sure your Gmail has emails matching the search keywords. Try running `runTracker()` manually and check the Apps Script **Logs** (View → Logs) for output.

**Wrong company or job title extracted**
The extraction relies on email formatting. You can manually edit any cell — the script will not overwrite fields you've filled in manually (except Status and Email Date on updates).

**Draft created in wrong language**
The language detector needs at least 2 German signal words to switch to DE. You can manually delete the draft and reply however you prefer.

**Trigger not running**
Go to **Extensions → Apps Script → Triggers** (clock icon) and confirm the trigger exists. If not, run `installTrigger()` again.

**Execution time exceeded**
Reduce `MAX_THREADS` or narrow `newer_than:Xd` to a shorter window.

---

## 📄 License

Free to use and modify for personal use. If you share or build on this publicly, a credit or link back is appreciated.

---

## 🙌 Credits

Built with Google Apps Script · Gmail API · Google Sheets API

If this helped your job search, consider sharing it with someone who needs it. ⭐
#   J o b T r a c k e r  
 