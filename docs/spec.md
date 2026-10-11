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
| P | Fee charged (₹) | Number | Yes | Suggests last fee for hospital + case type |
| Q | Amount received (₹) | Number | When paid | Adds up across partial payments |
| R | Balance (₹) | Formula | Auto | P − Q |
| S | Payment status | Formula | Auto | Paid / Partially paid / Pending; "Written off" is a manual value |
| T | Payment date | Date | When paid | |
| U | Payment mode | Cash / UPI / NEFT-RTGS / Cheque | When paid | |
| V | Bill / ref no. | Text | No | |
| W | Notes | Text | No | |
| X–Y | Created at, Updated at | Auto | Auto | |

**Other tabs.** *Lists* holds dropdown values and vocabularies. *Trash* has the same columns plus Deleted at. *Settings* holds key/value configuration (see README).

## Screens

| Screen | Purpose |
|---|---|
| Add / Edit case | Entry form; edit mode adds Delete |
| Records | Search and filters; Mark paid; Select → settle or delete in bulk; Export |
| Dashboard | KPI tiles, charts and filters; tapping a chart opens the matching cases |
| More | Manage lists, Trash, Back up now, Refresh |

## Rules

| Rule | Detail |
|---|---|
| Partial payment | Mark paid adds the amount to what was received; status follows the balance |
| Bulk settle | Each selected unpaid case becomes fully paid with one date and mode; paid and written-off cases are skipped |
| Delete | Moves rows to Trash; typed DELETE above the threshold; Undo for N seconds; restore any time before the purge |
| Purge | Nightly trigger removes Trash rows older than the retention days |
| Backup | Monthly trigger plus a button: full workbook as .xlsx in Drive, newest 12 kept |
| Access | `checkAccess_()` allows only the Settings emails; the Sheet's own sharing is the second lock |
| Concurrency | All writes take a script lock |

## Server functions (`Code.gs`)

| Function | Called from |
|---|---|
| `getInitData`, `saveCase`, `markPaid`, `deleteCases`, `restoreCases`, `getTrash`, `purgeTrash`, `listUpdate`, `exportCases`, `backupNow`, `getBackupInfo` | The app |
| `setupSheet`, `installTriggers` | Run once from the editor |
| `migrateAddAgeGender`, `migrateAddCaseFields` | Only for sheets created before those columns existed |
| `purgeOldTrash`, `monthlyBackup` | Triggers |
