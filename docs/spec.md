# App spec

A Google Sheet holds the data, Apps Script serves a mobile web app, and only the accounts listed in Settings can use it.

## Cases tab: one row per case

| Col | Field | Input | Required | Notes |
|---|---|---|---|---|
| A | Case ID | Auto (C-0001) | Auto | Never reused, even after a delete |
| B | Date of surgery | Date | Yes | Not in the future |
| C | Patient name | Text | Yes | |
| D | Age | Number (years, decimals allowed) | Yes | 0–120 |
| E | Gender | Male / Female / Other | Yes | |
| F | IP / UHID no. | Text | No | |
| G | Hospital | Dropdown + add new | Yes | Defaults to last used |
| H | Surgeon | Dropdown + add new | Yes | |
| I | Case type | Dropdown | Yes | |
| J | Anaesthesia type | Dropdown | No | |
| K | Case description | Dropdown + add new | No | Learns new entries |
| L | Elective / Emergency | Dropdown | Yes | Defaults to Elective |
| M | Yojana / Private | Two-way toggle | Yes | Defaults to Private |
| N | Risk notes | Tags, stored as `a; b; c` | No | Builds a vocabulary for suggestions |
| O | Duration (min) | Number | No | |
| P | Fee charged (₹) | Number | Yes | Total case fee; suggests last fee for hospital + case type |
| Q | Shared case | Yes / blank | No | Case shared with another anaesthesiologist |
| R | Role | Took over / Handed over | If shared | |
| S | Co-anaesthesiologist | Dropdown + add new | If shared | Learns new names |
| T | Collected by | Me / Colleague / Separately | If shared | Who the hospital pays |
| U | My share (₹) | Number | No | Blank = split not agreed yet |
| V | Amount received (₹) | Number | When paid | From the hospital, or from the colleague if they collected |
| W | Balance (₹) | Formula | Auto | Expected − V; expected = P, or U when the colleague collects / paid separately |
| X | Payment status | Formula | Auto | Paid / Partially paid / Pending / Split pending; "Written off" is a manual value |
| Y | Payment date | Date | When paid | |
| Z | Payment mode | Cash / UPI / NEFT-RTGS / Cheque | When paid | |
| AA | Settled with colleague on | Date | No | When I collected and paid the colleague their part |
| AB | Bill / ref no. | Text | No | |
| AC | Notes | Text | No | |
| AD–AE | Created at, Updated at | Auto | Auto | |

**Other tabs.** *Lists* holds dropdown values and vocabularies. *Trash* has the same columns plus Deleted at. *Settings* holds key/value configuration (see README).

## Screens

| Screen | Purpose |
|---|---|
| Home (first screen) | Greeting, waiting to be paid, billed and collected this month, Log a case, gentle reminders, recent cases |
| New / Edit case | Entry form with one-tap gender, urgency and hospital chips; edit mode adds Delete. Saving a new case returns to Home; saving an edit returns to Records |
| Records | Search; period (including custom dates), hospital and status filters; Mark paid; Select → settle or delete in bulk; Export |
| Insights | Period chips (including custom), headline tiles with change vs the previous period, charts and filters; tapping a chart opens the matching cases; Export |
| More | Appearance (Auto / Light / Dark, Sui on/off), Manage lists, Colleague balances, Trash, Send dues statement now, Back up now, Refresh |

Navigation: Home · Records · **+** (new case) · Insights · More.

## Rules

| Rule | Detail |
|---|---|
| Shared case earnings | My earnings = full fee for normal cases, my share for shared cases; undecided splits are excluded and shown as Split pending |
| I owe colleague | I collected: colleague share × (received ÷ fee), until marked settled |
| Colleague owes me | Colleague collected: my share − received |
| Partial payment | Mark paid adds the amount to what was received; status follows the balance |
| Bulk settle | Each selected unpaid case becomes fully paid with one date and mode; paid and written-off cases are skipped |
| Delete | Moves rows to Trash; typed DELETE above the threshold; Undo for N seconds; restore any time before the purge |
| Purge | Nightly trigger removes Trash rows older than the retention days |
| Backup | Monthly trigger plus a button: full workbook as .xlsx in Drive, newest 12 kept |
| Dues statement | 1st of each month (~9 am) plus a button: PDF of what each hospital owes (colleague-collected cases excluded), emailed with the PDF attached; optional per-hospital PDFs; optional WhatsApp totals via CallMeBot (no patient names) |
| Drive layout | `Case Tracker Backups/` (flat), `Case Tracker Exports/<year>/`, `Case Tracker Statements/<year>/` (a dated subfolder per send only when per-hospital PDFs are on) |
| Comparison deltas | This month vs the same days last month; last month vs the month before; this FY vs the same point last FY; last FY vs the FY before; custom vs the same number of days just before; all time has none |
| Dues ageing | Outstanding by hospital in 0–30 / 31–60 / 61–90 / 90+ day buckets (bucket edges from Settings); colleague-collected cases excluded |
| Exports | Cases: the filtered rows. Summary: one tab each for overview, month, hospital, case type, colleague, dues ageing. FY summary for CA: income by month and hospital, shared cases, notes; no patient names. Files go to `Case Tracker Exports/<year>/` |
| Appearance | Theme and Sui are stored per device in the browser; nothing is written to the Sheet |
| Access | `checkAccess_()` allows only the Settings emails; the Sheet's own sharing is the second lock |
| Concurrency | All writes take a script lock |

## Server functions (`Code.gs`)

| Function | Called from |
|---|---|
| `getInitData`, `saveCase`, `markPaid`, `setSplit`, `settleColleague`, `deleteCases`, `restoreCases`, `getTrash`, `purgeTrash`, `listUpdate`, `exportCases`, `exportTables`, `backupNow`, `getBackupInfo`, `sendStatementNow` | The app |
| `setupSheet`, `installTriggers`, `testWhatsApp` | Run from the editor |
| `migrateAddAgeGender`, `migrateAddCaseFields`, `migrateAddSharedCases` | Only for sheets created before those columns existed |
| `purgeOldTrash`, `monthlyBackup`, `monthlyStatement` | Triggers |
