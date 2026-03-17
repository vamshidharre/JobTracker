# 📋 Job Application Tracker
### Powered by Google Apps Script · Gmail · Google Sheets

![Version](https://img.shields.io/badge/version-3.0-blue)
![License](https://img.shields.io/badge/license-Free-green)
![Languages](https://img.shields.io/badge/languages-EN%20%7C%20DE-orange)
![Platform](https://img.shields.io/badge/platform-Google%20Workspace-red)

> Automatically scan your Gmail, extract job application details, track status changes, generate analytics, and draft recruiter replies — all inside a free Google Sheet.

---

## 📌 Table of Contents

- [Overview](#-overview)
- [Features](#-features)
- [How It Works](#-how-it-works)
- [Setup Guide](#-setup-guide)
- [Configuration](#-configuration)
- [Sheet Structure](#-sheet-structure)
- [Supported Languages](#-supported-languages)
- [ATS Platforms](#-detected-ats-platforms)
- [Privacy & Security](#-privacy--security)
- [Troubleshooting](#-troubleshooting)
- [License](#-license)

---

## 🧩 Overview

Manually tracking job applications across spreadsheets, sticky notes, and browser tabs is exhausting. This script does it all automatically.

Every 15 minutes it scans your Gmail, identifies job-related emails, and logs them into a structured Google Sheet — complete with company name, job title, HR contact, and application status. It also builds a live Dashboard, an Analytics tab, and auto-drafts reply emails when an interview or offer arrives.

**No paid tools. No third-party apps. Everything stays in your Google account.**

---

## ✨ Features

| Category | What It Does |
|---|---|
| 📥 Email Scanning | Scans Gmail every 15 min, last 90 days, up to 150 threads |
| 🧠 Classification | 9 status levels, EN + DE keywords, status only moves forward |
| 🔍 Data Extraction | Company, job title, HR name, email, phone, Gmail link |
| 🔁 Deduplication | Thread ID + fuzzy fingerprint matching to prevent duplicates |
| 📝 Auto-Drafts | Bilingual reply drafts for interviews, assessments, and offers |
| 📋 Dashboard | Live KPIs, status breakdown, ATS sources, today's activity |
| 📊 Analytics | Funnel rates, weekly trends, day-of-week heatmap, platform stats |
| 📧 Daily Email | 8am summary with new activity and overall stats |

---

## 🔄 How It Works

```
Gmail Inbox
    │
    ▼
Gmail Search Query (job keywords + last 90 days)
    │
    ▼
For each thread:
    ├─ Already tracked?        ──► Update status if progressed
    ├─ Same job, new channel?  ──► Merge into existing row
    └─ New application?        ──► Append new row
    │
    ▼
Auto-draft reply (if Interview / Offer / Assessment)
    │
    ▼
Refresh Dashboard + Analytics tabs
```

---

## 🛠️ Setup Guide

### Prerequisites
- A Google account with Gmail
- Google Sheets (free)
- No coding experience needed

---

### Step 1 — Create a New Google Sheet

Go to [sheets.google.com](https://sheets.google.com) and create a blank spreadsheet. Give it any name you like.

---

### Step 2 — Open Apps Script

Inside the sheet click **Extensions → Apps Script**. A new editor tab will open.

---

### Step 3 — Paste the Script

Delete all existing code in the editor, paste the full contents of `JobTracker.gs`, and click 💾 **Save**.

---

### Step 4 — Install Triggers

In the Apps Script editor, run each of these functions **once** by selecting from the dropdown and clicking ▶ **Run**:

```js
installTrigger()              // Starts the 15-minute Gmail scan
installDailySummaryTrigger()  // Enables the 8am daily summary email
```

---

### Step 5 — Authorise Permissions

Google will prompt you to grant access. Click through and allow:

| Permission | Why It's Needed |
|---|---|
| Gmail (read) | To scan job-related emails |
| Gmail (compose) | To create auto-reply drafts |
| Google Sheets | To write data into your spreadsheet |

> 🔒 All permissions are scoped to your own account. Nothing leaves Google's ecosystem.

---

### Step 6 — Run the First Scan

Select `runTracker` from the function dropdown and click ▶ **Run**. This triggers an immediate first scan. From this point it runs automatically every 15 minutes.

Check **View → Logs** in the editor to see what was found.

---

## ⚙️ Configuration

All settings are at the top of `JobTracker.gs`:

| Constant | Default | Description |
|---|---|---|
| `SHEET_NAME` | `"Applications"` | Name of the main data tab |
| `DASHBOARD_NAME` | `"Dashboard"` | Name of the dashboard tab |
| `ANALYTICS_NAME` | `"Analytics"` | Name of the analytics tab |
| `LABEL_NAME` | `"JobTracker"` | Gmail label for processed threads |
| `MAX_THREADS` | `150` | Max threads scanned per run |
| `AUTO_DRAFT_ENABLED` | `true` | Set to `false` to disable auto-drafts |
| `DRAFT_TRIGGER_STATUSES` | Interview, Offer, Assessment | Statuses that trigger a draft reply |

### Common Tweaks

**Scan more history**
```js
// In getJobEmailThreads() — change 90 to any number of days
"newer_than:180d"
```

**Scan entire mailbox (one-time)**
```js
// Temporarily remove newer_than and increase thread limit
MAX_THREADS = 500
// Remove "newer_than:90d" from the query
// Run once manually, then restore defaults
```

> ⚠️ Apps Script has a **6-minute execution limit** per run. Large scans may time out — increase gradually.

**Change summary email time**
```js
// In installDailySummaryTrigger()
.atHour(8)  // Change to any hour in 24h format
```

---

## 🗂️ Sheet Structure

The **Applications** tab contains 14 columns:

| # | Column | Auto-filled? | Description |
|---|---|---|---|
| 1 | Thread ID | ✅ | Internal Gmail thread identifier |
| 2 | Company | ✅ | Extracted company name |
| 3 | Job Title | ✅ | Extracted job title |
| 4 | HR Name | ✅ | Recruiter or sender name |
| 5 | HR Email | ✅ | Recruiter email address |
| 6 | HR Phone | ✅ | Phone number from email signature |
| 7 | Status | ✅ | Colour-coded application status |
| 8 | Date Applied | ✏️ Manual | Fill in the date you submitted |
| 9 | Email Date | ✅ | Date of the most recent relevant email |
| 10 | Subject | ✅ | Email subject line |
| 11 | Gmail Link | ✅ | Direct link to the Gmail thread |
| 12 | Notes | ✅ | Auto-appended status history and dedup notes |
| 13 | ATS Platform | ✅ | Detected platform (Greenhouse, Workday, etc.) |
| 14 | Fingerprint | ✅ | Internal dedup key — company + job title normalised |

### Status Colour Coding

| Status | Colour |
|---|---|
| Offer | 🟢 Green |
| Interview Scheduled | 🔵 Blue |
| Interview Invitation | 🩵 Light Blue |
| Technical Assessment | 🟣 Purple |
| Application Received | 🟡 Yellow |
| Applied | ⚪ Light Grey |
| Rejected | 🔴 Light Red |
| Follow-up Required | 🟠 Amber |
| Withdrawn | ⬜ Grey |

---

## 🌍 Supported Languages

| Feature | English | German |
|---|---|---|
| Status classification | ✅ | ✅ |
| Company name extraction | ✅ | ✅ incl. GmbH, AG, SE, KG |
| Job title extraction | ✅ | ✅ |
| HR name extraction | ✅ | ✅ |
| Auto-reply drafts | ✅ | ✅ auto-detected |
| Daily summary email | ✅ | — |

Language is auto-detected per email. No manual setup required.

---

## 🏢 Detected ATS Platforms

| Platform | Domain Detected |
|---|---|
| Greenhouse | greenhouse.io |
| Lever | lever.co |
| Workday | workday.com / myworkdayjobs.com |
| SmartRecruiters | smartrecruiters.com |
| Taleo | taleo.net |
| iCIMS | icims.com |
| Recruitee | recruitee.com |
| Jobvite | jobvite.com |
| Personio | personio.de |
| Softgarden | softgarden.de |
| StepStone | stepstone.de |
| XING | xing.com |
| LinkedIn | linkedin.com |

---

## 🔒 Privacy & Security

| Topic | Detail |
|---|---|
| Data location | Stays entirely within your Google account |
| External servers | None — no data is sent outside Google |
| API keys | Not required |
| Third-party services | Not used |
| Permissions | Viewable under Extensions → Apps Script → Project Settings |
| Removal | Deleting the script stops all automation immediately |

---

## 🐛 Troubleshooting

| Problem | Solution |
|---|---|
| No rows appearing after first run | Run `runTracker()` manually and check **View → Logs**. Make sure your Gmail has emails matching job keywords. |
| Wrong company or job title | Edit the cell manually — the script won't overwrite manual edits except Status and Email Date. |
| Draft created in wrong language | The detector needs 2+ German signal words. Delete the draft and reply manually if needed. |
| Trigger not firing | Go to **Extensions → Apps Script → Triggers** (clock icon). If missing, run `installTrigger()` again. |
| Execution timed out | Reduce `MAX_THREADS` or shorten the `newer_than:Xd` window. |
| Duplicate rows appearing | Check that the Fingerprint column is populated. If company or job title couldn't be extracted, dedup won't activate. |
| Draft re-created every run | The script checks by thread ID. If you delete a draft, a new one may appear on the next run for that thread. |

---

## 📄 License

Free to use and modify for personal use.
If you share this publicly or build on it, a link back to this repo is appreciated.

---

## 🙌 Credits

Built with **Google Apps Script** · **Gmail API** · **Google Sheets API**

If this helped your job search, consider giving the repo a ⭐ and sharing it with someone who's job hunting.
