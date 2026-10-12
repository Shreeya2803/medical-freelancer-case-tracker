# Medical Freelancer Case Tracker

A free, phone-first app for doctors who work freelance across several hospitals: log each case in under 30 seconds, track who has paid and who still owes, and see earnings at a glance. It replaces the paper ledger.

Built for a freelance anaesthesiologist; it suits any doctor billing per case across institutions.

**Cost: ₹0.** It runs entirely on a Google account: Google Sheets is the database, Google Apps Script is the server and the app, and the data can be downloaded as Excel at any time.

<p>
  <img src="docs/screenshots/01-home.png" width="24%" alt="Home">
  <img src="docs/screenshots/02-add-case.png" width="24%" alt="New case">
  <img src="docs/screenshots/04-insights.png" width="24%" alt="Insights">
  <img src="docs/screenshots/08-records-custom-dates.png" width="24%" alt="Records with custom dates">
</p>
<p>
  <img src="docs/screenshots/05-insights-dues.png" width="24%" alt="Money owed by hospitals, with ageing">
  <img src="docs/screenshots/07-export.png" width="24%" alt="Export options">
  <img src="docs/screenshots/10-all-clear.png" width="24%" alt="All clear moment">
  <img src="docs/screenshots/11-shared-case.png" width="24%" alt="Shared case entry">
</p>
<p>
  <img src="docs/screenshots/13-home-dark.png" width="24%" alt="Home, dark mode">
  <img src="docs/screenshots/14-insights-dark.png" width="24%" alt="Insights, dark mode">
  <img src="docs/screenshots/15-add-dark.png" width="24%" alt="New case, dark mode">
  <img src="docs/screenshots/12-colleague-balances.png" width="24%" alt="Colleague balances">
</p>

*Screenshots use made-up demo data.*

## Features

| Area | What it does |
|---|---|
| **Home** | Greeting, total waiting to be paid (tap to see those cases), billed and collected this month, one-tap **Log a case**, gentle reminders (hospitals with old dues, splits still to agree) and recent cases. |
| **Add case** | Date, patient name, age, gender, IP/UHID, hospital, surgeon, case type, anaesthesia type, case description, elective/emergency, Yojana/private, risk notes, fee. Gender and elective/emergency are one-tap buttons; the five most-used hospitals are chips, with the rest under *More*. Dropdowns learn new hospitals, surgeons and descriptions as they are typed. Suggests the last fee charged for the same hospital and case type. |
| **Shared cases** | For cases taken over from, or handed over to, another anaesthesiologist. Record who collects from the hospital and your share; leave the share blank until the split is agreed ("Split pending"). An optional ledger (Settings: *Track colleague settlements*) shows who owes whom per colleague, with one-tap settle. Earnings everywhere count only your share. |
| **Risk notes** | Tag-style input. Every note typed becomes vocabulary and is suggested again when its letters are typed, matching anywhere in the word. |
| **Records** | Search; filters by period (including **custom dates**), hospital and status; edit; one-tap **Mark paid** (full or partial); multi-select to settle a hospital's pending cases in one go; **Export** of the filtered view to Excel. |
| **Insights** | Period chips (this month, last month, this FY, last FY, all time, **custom**). Four headline tiles (billed, collected, outstanding, over 60 days) with ▲/▼ against the previous comparable period. Money owed by each hospital with ageing bars; billed vs collected by month; where earnings come from (hospital, case type or surgeon, with case count and average fee); case mix (shared vs independent, elective/emergency, Yojana/private, age groups); colleagues. Filters for hospital, case type, surgeon and more; tapping a bar opens those cases. |
| **Exports** | From Insights: **Cases** (every case in view), **Summary** (by month, hospital, case type, colleague and dues ageing, one tab each), **FY summary for your CA** (income by month and hospital, no patient names), and the **Dues statement** PDF. Excel files are filed in Drive by year. |
| **Look and feel** | "Ivory & Sage" theme with a matching **dark mode** (Auto / Light / Dark under More). **Sui**, a small syringe mascot, appears only at moments: the Home greeting, after a case is saved, when a hospital's dues are cleared, and on empty screens. It can be switched off. Both settings are saved per device. |
| **Safe deletes** | Select and delete any filtered set of cases. They go to Trash for 30 days, with Undo, Restore and a typed confirmation for large deletes. |
| **Dues statement** | On the 1st of each month (or on demand from More), a PDF of what each hospital owes is emailed and saved to Drive. Optional: one ready-to-forward PDF per hospital, and a totals-only WhatsApp summary via the free [CallMeBot](https://www.callmebot.com/blog/free-api-whatsapp-messages/) personal API. |
| **Backups** | One-tap and automatic monthly Excel backups to Google Drive (newest 12 kept). |
| **Access** | Only the Google accounts listed in Settings can open the app or the data. |

## How it works

```
iPhone / Android browser  ──►  Google Apps Script web app  ──►  Google Sheet (Cases · Lists · Trash · Settings)
   (home-screen shortcut)        Code.gs + Index.html               └─► Google Drive: Excel exports, monthly backups, dues statements
```

| Layer | Tool |
|---|---|
| Database | Google Sheets ([template](template/case_tracker_backend.xlsx)) |
| Server | Google Apps Script ([`src/Code.gs`](src/Code.gs)) |
| App (all screens) | One HTML file served by Apps Script ([`src/Index.html`](src/Index.html)), no frameworks |
| Scheduled jobs | Apps Script triggers: nightly Trash clean-up, monthly backup, monthly dues statement |

## Setup (about 20 minutes)

1. **Create the database.** Upload [`template/case_tracker_backend.xlsx`](template/case_tracker_backend.xlsx) to Google Drive and open it with Google Sheets. It is saved as a Google Sheet. Delete example row 2 on the Cases tab.
2. **Add the code.** In the Sheet, open **Extensions → Apps Script**. Replace `Code.gs` with [`src/Code.gs`](src/Code.gs). Add an HTML file named `Index` and paste in [`src/Index.html`](src/Index.html). In Project Settings, set the time zone.
3. **Run setup once.** In the function dropdown, run `setupSheet` (dropdowns and settings rows), then `installTriggers` (nightly Trash clean-up, monthly backup and monthly dues statement). Approve the permissions: *Advanced → Go to project → Allow*.
4. **Fill the Settings tab.** Enter the doctor name and credentials (shown at the top of every screen), the owner and editor emails (only these accounts get in), the statement email, and the Hospitals and Surgeons on the Lists tab.
5. **Deploy.** Go to **Deploy → New deployment → Web app**, with *Execute as: User accessing the web app* and *Who has access: Anyone with Google account*. Copy the `/exec` URL.
6. **Install on the phone.** Open the URL in Safari or Chrome, sign in, and tap *Share → Add to Home Screen*.

To update later, paste the new code, then go to **Deploy → Manage deployments → Edit → New version**. The URL stays the same.

**Upgrading from v1.x to v2.0:** replace both `Code.gs` and `Index.html` and deploy a new version. The Sheet does not change and no migration is needed. The first Summary or FY export may ask to approve Drive access again.

> **iPhone tip:** if Safari shows *"Sorry, unable to open the file at present"*, more than one Google account is signed in. Sign out of all accounts and sign in with just the one listed in Settings.

<details>
<summary>Using Google's <code>clasp</code> command-line tool instead of copy-paste</summary>

```bash
npm install -g @google/clasp
clasp login
clasp clone <your-script-id> --rootDir src   # Script ID: Apps Script → Project Settings
clasp push
```
</details>

## Settings (Settings tab)

| Setting | Default | Effect |
|---|---|---|
| Financial year start month | 4 (April) | "This FY" / "Last FY" filters |
| Ageing buckets | 30 / 60 / 90 days | Ageing-of-dues chart |
| Overdue threshold | 60 days | Overdue KPI |
| Trash retention | 30 days | Nightly purge |
| Bulk-delete typed confirmation above | 20 cases | Requires typing DELETE |
| Undo banner | 10 seconds | Undo window after a delete |
| Owner / editor email | blank | Once filled, only these accounts can use the app |
| Backup folder name | Case Tracker Backups | Drive folder for backups |
| Drive folders | — | Exports and statements are filed by year: `Case Tracker Exports/2026/`, `Case Tracker Statements/2026/` |
| Doctor name / Credentials | blank | Header on every screen |
| Monthly dues statement | On | Off = only when "Send dues statement now" is tapped |
| Statement email | blank | Recipients of the PDF statement |
| Statement WhatsApp number / CallMeBot API key | blank | Number that activated CallMeBot, with country code, and its key |
| Per-hospital PDFs | Off | On = also one ready-to-forward PDF per hospital with each statement |
| Payment details (for hospital statements) | blank | UPI ID / bank details printed on each hospital's statement |
| Track colleague settlements | Yes | No = only record your final share on shared cases |

## Data model

One row per case on the **Cases** tab. Balance and payment status are sheet formulas. Full column list: [docs/spec.md](docs/spec.md).

## Privacy

Patient data lives only in the owner's Google Sheet, never in this repository. The app runs as the signed-in user, so it can only read what that account has been shared. Keep the Sheet private, share it only with the accounts listed in Settings, and turn on 2-step verification.

## Licence

[MIT](LICENSE)
