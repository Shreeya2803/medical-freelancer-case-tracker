/**
 * Anaesthesia Case & Earnings Tracker — server code
 * Paste into Extensions → Apps Script of the case_tracker_backend Google Sheet.
 */

const SHEET_ID = ''; // Leave blank when this script is opened from the Sheet (Extensions → Apps Script)
const CASES = 'Cases', LISTS = 'Lists', TRASH = 'Trash', SETTINGS = 'Settings';
const NCOLS = 25;
const COL = {
  id: 1, date: 2, patient: 3, age: 4, gender: 5, uhid: 6, hospital: 7, surgeon: 8, caseType: 9,
  anaesthesia: 10, caseDesc: 11, urgency: 12, scheme: 13, riskNotes: 14, duration: 15, fee: 16,
  received: 17, balance: 18, status: 19, payDate: 20, payMode: 21, billNo: 22, notes: 23,
  created: 24, updated: 25
};
const GENDERS = ['Male', 'Female', 'Other'];
const SCHEMES = ['Private', 'Yojana'];
// Lists tab columns
const LCOL = { hospitals: 1, surgeons: 2, caseTypes: 3, anaesthesia: 4, urgency: 5, payModes: 6, statuses: 7,
  genders: 8, caseDescs: 9, riskNotes: 10, schemes: 11 };
const NLISTS = 11;
const RISK_SEP = '; ';
const A1 = c => String.fromCharCode(64 + c);   // column number → letter (A–Z)
const BAL_F = r => {
  const K = A1(COL.fee), L = A1(COL.received);
  return `=IF(${K}${r}="","",${K}${r}-N(${L}${r}))`;
};
const STATUS_F = r => {
  const K = A1(COL.fee), L = A1(COL.received), M = A1(COL.balance);
  return `=IF(${K}${r}="","",IF(${M}${r}<=0,"Paid",IF(N(${L}${r})>0,"Partially paid","Pending")))`;
};

/* ---------- Entry point ---------- */

function doGet() {
  checkAccess_();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Case Tracker')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .addMetaTag('apple-mobile-web-app-capable', 'yes');
}

/* ---------- One-time setup: run from the editor (select setupSheet → Run) ---------- */

// Rebuilds the Sheet's dropdowns on the Cases tab from the Lists tab.
// Needed because some dropdowns don't survive the Excel → Google Sheets conversion.
function setupSheet() {
  const ss = ss_();
  const cases = ss.getSheetByName(CASES), lists = ss.getSheetByName(LISTS);
  const rows = Math.max(cases.getMaxRows() - 1, 1000);
  const map = [
    [COL.hospital, 1], [COL.surgeon, 2], [COL.caseType, 3], [COL.anaesthesia, 4],
    [COL.urgency, 5], [COL.payMode, 6], [COL.gender, 8], [COL.caseDesc, LCOL.caseDescs], [COL.scheme, LCOL.schemes]
  ];
  map.forEach(([caseCol, listCol]) => {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(lists.getRange(2, listCol, 200, 1), true)
      .setAllowInvalid(false)
      .build();
    cases.getRange(2, caseCol, rows, 1).setDataValidation(rule);
  });
  ss.setSpreadsheetTimeZone('Asia/Kolkata');
  // Make sure the header settings exist on the Settings tab (shown at the top of every screen).
  const st = ss.getSheetByName(SETTINGS);
  const have = st.getRange(1, 1, Math.max(st.getLastRow(), 1), 1).getValues().map(r => String(r[0]).trim());
  [['Doctor name', '', 'Shown at the top of every screen, e.g. Dr. A. Sharma'],
   ['Credentials', '', 'Shown under the name, e.g. MBBS, MD · Anaesthesiologist']].forEach(row => {
    if (!have.includes(row[0])) st.appendRow(row);
  });
  SpreadsheetApp.flush();
  Logger.log('Dropdowns set on the Cases tab; Settings checked.');
}

// Run ONCE (select migrateAddAgeGender → Run) to add Age and Gender columns after Patient name.
// Sheets shifts every formula, dropdown and existing row automatically. Safe to run twice.
function migrateAddAgeGender() {
  const ss = ss_();
  [CASES, TRASH].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (String(sh.getRange(1, COL.age).getValue()).trim() === 'Age') return;   // already done
    sh.insertColumnsAfter(COL.patient, 2);
    sh.getRange(1, COL.patient).copyTo(sh.getRange(1, COL.age, 1, 2), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    sh.getRange(1, COL.age, 1, 2).setValues([['Age', 'Gender']]);
    sh.getRange(2, COL.age, sh.getMaxRows() - 1, 1).setNumberFormat('0.#');
    sh.setColumnWidth(COL.age, 60);
    sh.setColumnWidth(COL.gender, 80);
  });
  const lists = ss.getSheetByName(LISTS);
  if (String(lists.getRange(1, 8).getValue()).trim() !== 'Gender') {
    lists.getRange(1, 7).copyTo(lists.getRange(1, 8), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    lists.getRange(1, 8, 1 + GENDERS.length, 1).setValues([['Gender']].concat(GENDERS.map(g => [g])));
  }
  setupSheet();
  Logger.log('Age and Gender columns added.');
}

// Run ONCE (select migrateAddCaseFields → Run) to add Case description, Scheme (Yojana/Private) and Risk notes.
// Run migrateAddAgeGender first if you haven't. Safe to run twice.
function migrateAddCaseFields() {
  const ss = ss_();
  [CASES, TRASH].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (String(sh.getRange(1, COL.age).getValue()).trim() !== 'Age') throw new Error('Run migrateAddAgeGender first.');
    if (String(sh.getRange(1, COL.caseDesc).getValue()).trim() === 'Case description') return;   // already done
    sh.insertColumnsAfter(COL.anaesthesia, 1);                 // → Case description
    sh.insertColumnsAfter(COL.urgency, 2);                     // → Scheme, Risk notes
    [[COL.caseDesc, 'Case description', 200], [COL.scheme, 'Yojana / Private', 100], [COL.riskNotes, 'Risk notes', 240]]
      .forEach(([c, title, w]) => {
        sh.getRange(1, COL.patient).copyTo(sh.getRange(1, c), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
        sh.getRange(1, c).setValue(title);
        sh.setColumnWidth(c, w);
      });
  });
  const lists = ss.getSheetByName(LISTS);
  [[LCOL.caseDescs, 'Case descriptions', []], [LCOL.riskNotes, 'Risk notes', []], [LCOL.schemes, 'Yojana / Private', SCHEMES]]
    .forEach(([c, title, vals]) => {
      if (String(lists.getRange(1, c).getValue()).trim() === title) return;
      lists.getRange(1, 7).copyTo(lists.getRange(1, c), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      lists.getRange(1, c).setValue(title);
      if (vals.length) lists.getRange(2, c, vals.length, 1).setValues(vals.map(v => [v]));
    });
  setupSheet();
  Logger.log('Case description, Yojana / Private and Risk notes columns added.');
}

/* ---------- Called from the app ---------- */

function getInitData() {
  checkAccess_();
  const m = getSettings_().map;
  return {
    lists: getLists_(),
    records: getRecords_(),
    lastHospital: PropertiesService.getUserProperties().getProperty('lastHospital') || '',
    fyStartMonth: Number(m['Financial year start month']) || 4,
    ageing: [
      Number(m['Ageing bucket 1 upper (days)']) || 30,
      Number(m['Ageing bucket 2 upper (days)']) || 60,
      Number(m['Ageing bucket 3 upper (days)']) || 90
    ],
    overdueDays: Number(m['Overdue threshold (days)']) || 60,
    bulkConfirm: Number(m['Bulk-delete typed confirmation above']) || 20,
    undoSeconds: Number(m['Undo banner (seconds)']) || 10,
    trashDays: Number(m['Trash retention (days)']) || 30,
    doctorName: String(m['Doctor name'] || '').trim(),
    credentials: String(m['Credentials'] || '').trim()
  };
}

function saveCase(c) {
  checkAccess_();
  const err = validate_(c);
  if (err) throw new Error(err);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let id;
  try {
    const sh = ss_().getSheetByName(CASES);
    const now = new Date();
    let row, created;
    if (c.id) {
      row = findRow_(sh, c.id);
      if (!row) throw new Error('Case ' + c.id + ' not found. It may have been deleted.');
      id = c.id;
      created = sh.getRange(row, COL.created).getValue() || now;
    } else {
      row = firstEmptyRow_(sh);
      id = nextId_();
      created = now;
    }
    const received = Number(c.received) || 0;

    // Add any new hospital/surgeon to Lists FIRST, so the Sheet's dropdown rules accept it.
    const desc = String(c.caseDesc || '').trim();
    const risks = cleanRisks_(c.riskNotes);
    addToList_(LCOL.hospitals, c.hospital.trim());
    addToList_(LCOL.surgeons, c.surgeon.trim());
    if (desc) addToList_(LCOL.caseDescs, desc);
    risks.forEach(r => addToList_(LCOL.riskNotes, r));
    SpreadsheetApp.flush();

    sh.getRange(row, 1, 1, COL.received).setValues([[
      id, toDate_(c.date), String(c.patient).trim(), num_(c.age), c.gender || '', c.uhid || '',
      c.hospital.trim(), c.surgeon.trim(), c.caseType, c.anaesthesia || '', desc, c.urgency || 'Elective',
      c.scheme, risks.join(RISK_SEP), num_(c.duration), Number(c.fee), received
    ]]);
    sh.getRange(row, COL.payDate, 1, 6).setValues([[
      received > 0 ? toDate_(c.payDate) : '', received > 0 ? (c.payMode || '') : '',
      c.billNo || '', c.notes || '', created, now
    ]]);
    ensureFormulas_(sh, row);

    PropertiesService.getUserProperties().setProperty('lastHospital', c.hospital.trim());
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  const data = getInitData();
  data.savedId = id;
  return data;
}

/**
 * Record a payment.
 * p = { ids: ['C-0001', ...], date: 'yyyy-mm-dd', mode: 'UPI',
 *       amount: 2500 }   // amount only for a single case (adds to what was received)
 * Without amount, each case is settled in full (received = fee).
 * Paid and Written-off cases are skipped.
 */
function markPaid(p) {
  checkAccess_();
  if (!p || !p.ids || !p.ids.length) throw new Error('No cases selected.');
  if (!p.date) throw new Error('Pick a payment date.');
  if (!p.mode) throw new Error('Pick a payment mode.');
  const today = Utilities.formatDate(new Date(), ss_().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  if (p.date > today) throw new Error('Payment date cannot be in the future.');
  const single = p.amount !== undefined && p.amount !== null && p.amount !== '';
  if (single && p.ids.length !== 1) throw new Error('An amount can only be given for one case.');
  if (single && !(Number(p.amount) > 0)) throw new Error('Amount must be more than 0.');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let updated = 0;
  try {
    const sh = ss_().getSheetByName(CASES);
    const last = sh.getLastRow();
    if (last < 2) throw new Error('No cases found.');
    const n = last - 1;
    const ids = sh.getRange(2, COL.id, n, 1).getValues().map(r => String(r[0]).trim());
    // Columns K (fee) through T (updated at): edit values, keep formulas as formulas.
    const rng = sh.getRange(2, COL.fee, n, COL.updated - COL.fee + 1);
    const vals = rng.getValues(), forms = rng.getFormulas();
    const k = c => c - COL.fee;  // column → index in this block
    const want = new Set(p.ids), now = new Date();
    const dirty = new Set();

    ids.forEach((id, i) => {
      if (!want.has(id)) return;
      const v = vals[i];
      const fee = Number(v[k(COL.fee)]) || 0, got = Number(v[k(COL.received)]) || 0;
      if (String(v[k(COL.status)]) === 'Written off' || fee - got <= 0) return;
      v[k(COL.received)] = single ? got + Number(p.amount) : fee;
      v[k(COL.payDate)] = toDate_(p.date);
      v[k(COL.payMode)] = p.mode;
      v[k(COL.updated)] = now;
      dirty.add(i);
      updated++;
    });

    if (updated) {
      const first = Math.min(...dirty), lastI = Math.max(...dirty);
      const out = [];
      for (let i = first; i <= lastI; i++) out.push(vals[i].map((x, j) => forms[i][j] || x));
      sh.getRange(2 + first, COL.fee, out.length, out[0].length).setValues(out);
      SpreadsheetApp.flush();
    }
  } finally {
    lock.releaseLock();
  }
  const data = getInitData();
  data.updated = updated;
  return data;
}

/* ---------- Delete, Trash, Restore ---------- */


// Moves cases to the Trash tab. Their rows on Cases are emptied (formulas kept) and reused later.
function deleteCases(ids) {
  checkAccess_();
  if (!ids || !ids.length) throw new Error('No cases selected.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let moved = [];
  try {
    const ss = ss_(), sh = ss.getSheetByName(CASES), trash = ss.getSheetByName(TRASH);
    const last = sh.getLastRow();
    if (last < 2) throw new Error('No cases found.');
    const n = last - 1;
    const rng = sh.getRange(2, 1, n, NCOLS);
    const vals = rng.getValues(), forms = rng.getFormulas();
    const want = new Set(ids), now = new Date(), dirty = [];
    vals.forEach((v, i) => {
      if (!want.has(String(v[0]).trim())) return;
      moved.push(v.concat([now]));          // values as shown, incl. balance + status
      dirty.push(i);
      const r = i + 2;
      vals[i] = v.map(() => '');
      forms[i] = forms[i].map(() => '');
      forms[i][COL.balance - 1] = BAL_F(r);
      forms[i][COL.status - 1] = STATUS_F(r);
    });
    if (!moved.length) throw new Error('Those cases were not found. Refresh and try again.');
    trash.getRange(Math.max(trash.getLastRow(), 1) + 1, 1, moved.length, NCOLS + 1).setValues(moved);
    const first = Math.min(...dirty), lastI = Math.max(...dirty), out = [];
    for (let i = first; i <= lastI; i++) out.push(vals[i].map((x, j) => forms[i][j] || x));
    sh.getRange(2 + first, 1, out.length, NCOLS).setValues(out);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  const data = getInitData();
  data.deleted = moved.map(v => String(v[0]));
  return data;
}

// Puts cases from Trash back onto Cases.
function restoreCases(ids) {
  checkAccess_();
  if (!ids || !ids.length) throw new Error('Nothing to restore.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let restored = 0;
  try {
    const ss = ss_(), sh = ss.getSheetByName(CASES), trash = ss.getSheetByName(TRASH);
    const tLast = trash.getLastRow();
    if (tLast < 2) throw new Error('Trash is empty.');
    const tv = trash.getRange(2, 1, tLast - 1, NCOLS + 1).getValues();
    const want = new Set(ids), keep = [];
    tv.forEach(v => {
      const id = String(v[0]).trim();
      if (!want.has(id)) { keep.push(v); return; }
      addToList_(1, String(v[COL.hospital - 1]).trim());
      addToList_(2, String(v[COL.surgeon - 1]).trim());
      if (String(v[COL.caseDesc - 1]).trim()) addToList_(LCOL.caseDescs, String(v[COL.caseDesc - 1]).trim());
      SpreadsheetApp.flush();
      const row = findRow_(sh, id) || firstEmptyRow_(sh);
      sh.getRange(row, 1, 1, COL.received).setValues([v.slice(0, COL.received)]);
      sh.getRange(row, COL.payDate, 1, 6).setValues([v.slice(COL.payDate - 1, NCOLS - 1).concat([new Date()])]);
      sh.getRange(row, COL.balance).setFormula(BAL_F(row));
      if (String(v[COL.status - 1]) === 'Written off') sh.getRange(row, COL.status).setValue('Written off');
      else sh.getRange(row, COL.status).setFormula(STATUS_F(row));
      restored++;
    });
    rewriteTrash_(trash, keep);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  const data = getInitData();
  data.restored = restored;
  return data;
}

function getTrash() {
  checkAccess_();
  const ss = ss_(), trash = ss.getSheetByName(TRASH), tz = ss.getSpreadsheetTimeZone();
  const days = trashDays_();
  const last = trash.getLastRow();
  if (last < 2) return { items: [], days };
  const fmt = d => (d instanceof Date ? Utilities.formatDate(d, tz, 'yyyy-MM-dd') : String(d || ''));
  const now = Date.now();
  const items = trash.getRange(2, 1, last - 1, NCOLS + 1).getValues()
    .filter(v => String(v[0]).trim())
    .map(v => {
      const del = v[NCOLS] instanceof Date ? v[NCOLS] : new Date(v[NCOLS]);
      return {
        id: String(v[0]).trim(), date: fmt(v[1]), patient: String(v[COL.patient - 1]),
        hospital: String(v[COL.hospital - 1]), caseType: String(v[COL.caseType - 1]),
        fee: Number(v[COL.fee - 1]) || 0, status: String(v[COL.status - 1]),
        deletedOn: fmt(del), daysLeft: Math.max(0, days - Math.floor((now - del.getTime()) / 864e5))
      };
    })
    .sort((a, b) => b.deletedOn.localeCompare(a.deletedOn));
  return { items, days };
}

// Permanently removes cases from Trash. ids = null empties the whole Trash.
function purgeTrash(ids) {
  checkAccess_();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const trash = ss_().getSheetByName(TRASH);
    const last = trash.getLastRow();
    if (last >= 2) {
      const want = ids ? new Set(ids) : null;
      const keep = want ? trash.getRange(2, 1, last - 1, NCOLS + 1).getValues()
        .filter(v => String(v[0]).trim() && !want.has(String(v[0]).trim())) : [];
      rewriteTrash_(trash, keep);
    }
  } finally {
    lock.releaseLock();
  }
  return getTrash();
}

// Runs nightly (see installTriggers): removes Trash rows older than the retention days.
function purgeOldTrash() {
  const trash = ss_().getSheetByName(TRASH);
  const last = trash.getLastRow();
  if (last < 2) return;
  const cutoff = Date.now() - trashDays_() * 864e5;
  const rows = trash.getRange(2, 1, last - 1, NCOLS + 1).getValues();
  const keep = rows.filter(v => {
    if (!String(v[0]).trim()) return false;
    const d = v[NCOLS] instanceof Date ? v[NCOLS] : new Date(v[NCOLS]);
    return isNaN(d) || d.getTime() >= cutoff;
  });
  if (keep.length !== rows.length) rewriteTrash_(trash, keep);
}

// Run ONCE from the editor (select installTriggers → Run). Re-run after the Sheet changes owner.
function installTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(t => ['purgeOldTrash', 'monthlyBackup'].includes(t.getHandlerFunction()))
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('purgeOldTrash').timeBased().everyDays(1).atHour(2).create();
  ScriptApp.newTrigger('monthlyBackup').timeBased().onMonthDay(1).atHour(3).create();
  Logger.log('Scheduled: nightly Trash clean-up (~2 am) and a backup on the 1st of each month (~3 am).');
}

function rewriteTrash_(trash, keep) {
  const last = trash.getLastRow();
  if (last >= 2) trash.getRange(2, 1, last - 1, NCOLS + 1).clearContent();
  if (keep.length) trash.getRange(2, 1, keep.length, NCOLS + 1).setValues(keep);
}

function trashDays_() {
  return Number(getSettings_().map['Trash retention (days)']) || 30;
}


/* ---------- Lists, Export, Backup ---------- */

const LIST_COLS = { hospitals: LCOL.hospitals, surgeons: LCOL.surgeons, caseTypes: LCOL.caseTypes,
  caseDescs: LCOL.caseDescs, riskNotes: LCOL.riskNotes };

// op = 'add' | 'remove'; which = 'hospitals' | 'surgeons' | 'caseTypes'
// Removing only hides a name from the dropdowns; existing cases keep it.
function listUpdate(op, which, value) {
  checkAccess_();
  const col = LIST_COLS[which];
  if (!col) throw new Error('Unknown list.');
  value = String(value || '').trim();
  if (!value) throw new Error('Enter a name.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (op === 'add') {
      addToList_(col, value);
    } else if (op === 'remove') {
      const sh = ss_().getSheetByName(LISTS);
      const n = Math.max(sh.getLastRow() - 1, 1);
      const rng = sh.getRange(2, col, n, 1);
      const keep = rng.getValues().map(r => String(r[0]).trim()).filter(v => v && v !== value);
      rng.clearContent();
      if (keep.length) sh.getRange(2, col, keep.length, 1).setValues(keep.map(v => [v]));
    } else {
      throw new Error('Unknown action.');
    }
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return getLists_();
}

// Saves the given cases as an Excel file in Drive ("Case Tracker Exports") and returns its link.
function exportCases(ids) {
  checkAccess_();
  if (!ids || !ids.length) throw new Error('No cases to export.');
  const ss = ss_(), sh = ss.getSheetByName(CASES), tz = ss.getSpreadsheetTimeZone();
  const last = sh.getLastRow();
  const all = sh.getRange(1, 1, last, NCOLS).getValues();
  const want = new Set(ids);
  const rows = all.slice(1).filter(r => want.has(String(r[0]).trim()))
    .sort((a, b) => (b[COL.date - 1] || 0) - (a[COL.date - 1] || 0));
  if (!rows.length) throw new Error('Those cases were not found. Refresh and try again.');
  const name = 'Cases export ' + Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HHmm');
  const file = writeXlsx_(name, [all[0]].concat(rows), folder_('Case Tracker Exports'));
  return { url: file.getUrl(), name: file.getName(), count: rows.length };
}

// Full copy of the workbook as .xlsx in the backup folder. Keeps the latest 12.
function backupNow() {
  checkAccess_();
  return backup_();
}
function monthlyBackup() { backup_(); }

function backup_() {
  const ss = ss_(), tz = ss.getSpreadsheetTimeZone();
  const folderName = String(getSettings_().map['Backup folder name'] || 'Case Tracker Backups').trim();
  const folder = folder_(folderName);
  const name = 'Case Tracker backup ' + Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const blob = xlsxBlob_(ss.getId(), name);
  const file = folder.createFile(blob);
  // keep the newest 12 backups
  const files = [];
  const it = folder.getFiles();
  while (it.hasNext()) { const f = it.next(); if (f.getName().indexOf('Case Tracker backup') === 0) files.push(f); }
  files.sort((a, b) => b.getDateCreated() - a.getDateCreated()).slice(12).forEach(f => f.setTrashed(true));
  const when = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  PropertiesService.getScriptProperties().setProperty('lastBackup', when);
  return { url: file.getUrl(), name: file.getName(), folder: folderName, date: when };
}

function getBackupInfo() {
  checkAccess_();
  return { last: PropertiesService.getScriptProperties().getProperty('lastBackup') || '' };
}

function folder_(name) {
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

function xlsxBlob_(spreadsheetId, name) {
  const url = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export?format=xlsx';
  const res = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } });
  return res.getBlob().setName(name + '.xlsx');
}

// Writes rows to a temporary Google Sheet, converts it to .xlsx in `folder`, then deletes the temp sheet.
function writeXlsx_(name, rows, folder) {
  const tmp = SpreadsheetApp.create(name);
  const sh = tmp.getSheets()[0].setName('Cases');
  sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sh.getRange(1, 1, 1, rows[0].length).setFontWeight('bold').setBackground('#1f4e78').setFontColor('#ffffff');
  if (rows.length > 1) {
    sh.getRange(2, COL.date, rows.length - 1, 1).setNumberFormat('dd-mmm-yyyy');
    sh.getRange(2, COL.payDate, rows.length - 1, 1).setNumberFormat('dd-mmm-yyyy');
    sh.getRange(2, COL.fee, rows.length - 1, 3).setNumberFormat('[>=10000000]##\\,##\\,##\\,##0;[>=100000]##\\,##\\,##0;##,##0');
    sh.getRange(2, COL.created, rows.length - 1, 2).setNumberFormat('dd-mmm-yyyy hh:mm');
  }
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, rows[0].length);
  SpreadsheetApp.flush();
  const file = folder.createFile(xlsxBlob_(tmp.getId(), name));
  DriveApp.getFileById(tmp.getId()).setTrashed(true);
  return file;
}

/* ---------- Reads ---------- */

function getLists_() {
  const sh = ss_().getSheetByName(LISTS);
  const n = Math.max(sh.getLastRow() - 1, 1);
  const v = sh.getRange(2, 1, n, NLISTS).getValues();
  const col = i => v.map(r => String(r[i]).trim()).filter(Boolean);
  return {
    hospitals: col(0), surgeons: col(1), caseTypes: col(2), anaesthesia: col(3),
    urgency: col(4), payModes: col(5), statuses: col(6),
    genders: col(7).length ? col(7) : GENDERS,
    caseDescs: col(8), riskNotes: col(9),
    schemes: col(10).length ? col(10) : SCHEMES
  };
}

function getRecords_() {
  const ss = ss_();
  const sh = ss.getSheetByName(CASES);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const tz = ss.getSpreadsheetTimeZone();
  const fmt = d => (d instanceof Date ? Utilities.formatDate(d, tz, 'yyyy-MM-dd') : String(d || ''));
  return sh.getRange(2, 1, last - 1, NCOLS).getValues()
    .map((r, i) => ({
      row: i + 2,
      id: String(r[COL.id - 1]).trim(),
      date: fmt(r[COL.date - 1]),
      patient: String(r[COL.patient - 1]),
      age: r[COL.age - 1] === '' ? '' : Number(r[COL.age - 1]),
      gender: String(r[COL.gender - 1]),
      uhid: String(r[COL.uhid - 1]),
      hospital: String(r[COL.hospital - 1]),
      surgeon: String(r[COL.surgeon - 1]),
      caseType: String(r[COL.caseType - 1]),
      anaesthesia: String(r[COL.anaesthesia - 1]),
      caseDesc: String(r[COL.caseDesc - 1]),
      urgency: String(r[COL.urgency - 1]),
      scheme: String(r[COL.scheme - 1]),
      riskNotes: String(r[COL.riskNotes - 1]).split(';').map(x => x.trim()).filter(Boolean),
      duration: r[COL.duration - 1] === '' ? '' : Number(r[COL.duration - 1]),
      fee: Number(r[COL.fee - 1]) || 0,
      received: Number(r[COL.received - 1]) || 0,
      balance: Number(r[COL.balance - 1]) || 0,
      status: String(r[COL.status - 1]),
      payDate: fmt(r[COL.payDate - 1]),
      payMode: String(r[COL.payMode - 1]),
      billNo: String(r[COL.billNo - 1]),
      notes: String(r[COL.notes - 1])
    }))
    .filter(x => x.id);
}

function getSettings_() {
  const v = ss_().getSheetByName(SETTINGS).getDataRange().getValues();
  const map = {};
  v.slice(1).forEach(r => { map[String(r[0]).trim()] = r[1]; });
  const emails = [map['Owner email (Dad)'], map['Editor email (you)']]
    .map(e => String(e || '').trim().toLowerCase())
    .filter(Boolean);
  return { map, emails };
}

/* ---------- Helpers ---------- */

function ss_() {
  return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

// Extra lock on top of Sheet sharing: once Settings has emails, only those accounts get in.
function checkAccess_() {
  const allowed = getSettings_().emails;
  if (!allowed.length) return;
  const me = (Session.getActiveUser().getEmail() || '').toLowerCase();
  if (!allowed.includes(me)) throw new Error('Access denied for ' + (me || 'this account') + '.');
}

function validate_(c) {
  if (!c) return 'Nothing to save.';
  const missing = [];
  if (!c.date) missing.push('date');
  if (!String(c.patient || '').trim()) missing.push('patient name');
  if (c.age === '' || c.age === undefined || c.age === null) missing.push('age');
  if (!c.gender) missing.push('gender');
  if (!SCHEMES.includes(c.scheme)) missing.push('Yojana / Private');
  if (!String(c.hospital || '').trim()) missing.push('hospital');
  if (!String(c.surgeon || '').trim()) missing.push('surgeon');
  if (!c.caseType) missing.push('case type');
  if (missing.length) return 'Please fill: ' + missing.join(', ') + '.';

  const today = Utilities.formatDate(new Date(), ss_().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  if (c.date > today) return 'Surgery date cannot be in the future.';
  if (!(Number(c.fee) > 0)) return 'Fee must be more than 0.';
  if (!(Number(c.age) >= 0 && Number(c.age) <= 120)) return 'Age must be between 0 and 120.';
  const received = Number(c.received) || 0;
  if (received < 0) return 'Amount received cannot be negative.';
  if (received > 0 && !c.payMode) return 'Pick a payment mode.';
  if (received > 0 && !c.payDate) return 'Pick a payment date.';
  return '';
}

// Risk notes arrive as an array (or "a; b" text); trimmed, no duplicates, no ';' inside a note.
function cleanRisks_(x) {
  const arr = Array.isArray(x) ? x : String(x || '').split(';');
  const seen = new Set(), out = [];
  arr.map(v => String(v).replace(/;/g, ',').trim()).filter(Boolean).forEach(v => {
    if (!seen.has(v.toLowerCase())) { seen.add(v.toLowerCase()); out.push(v); }
  });
  return out;
}

function toDate_(s) {
  if (!s) return '';
  const p = String(s).split('-').map(Number);
  return new Date(p[0], p[1] - 1, p[2]);
}

function num_(v) {
  return v === '' || v === null || v === undefined ? '' : Number(v);
}

function firstEmptyRow_(sh) {
  const n = Math.max(sh.getLastRow(), 2);
  const ids = sh.getRange(2, 1, n - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (!String(ids[i][0]).trim()) return i + 2;
  return n + 1;
}

function findRow_(sh, id) {
  const n = sh.getLastRow();
  if (n < 2) return 0;
  const ids = sh.getRange(2, 1, n - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]).trim() === id) return i + 2;
  return 0;
}

// Next ID looks at Cases, Trash and the last ID issued, so a deleted case's ID is never reused.
function nextId_() {
  let max = 0;
  [CASES, TRASH].forEach(name => {
    const sh = ss_().getSheetByName(name);
    const n = sh.getLastRow();
    if (n < 2) return;
    sh.getRange(2, 1, n - 1, 1).getValues().forEach(r => {
      const m = String(r[0]).match(/^C-(\d+)$/);
      if (m) max = Math.max(max, Number(m[1]));
    });
  });
  // Also remember the highest ID ever issued, so IDs purged from Trash are never reused.
  const props = PropertiesService.getScriptProperties();
  max = Math.max(max, Number(props.getProperty('maxId')) || 0);
  props.setProperty('maxId', String(max + 1));
  return 'C-' + String(max + 1).padStart(4, '0');
}

// Rows past 1001 have no formulas yet; add them. A manual "Written off" status is left alone.
function ensureFormulas_(sh, row) {
  const bal = sh.getRange(row, COL.balance);
  if (!bal.getFormula()) bal.setFormula(BAL_F(row));
  const st = sh.getRange(row, COL.status);
  if (!st.getFormula() && st.getValue() !== 'Written off') {
    st.setFormula(STATUS_F(row));
  }
}

// Adds a new hospital or surgeon to the Lists tab if it is not there yet.
function addToList_(col, value) {
  if (!value) return;
  const sh = ss_().getSheetByName(LISTS);
  const n = Math.max(sh.getLastRow(), 2);
  const v = sh.getRange(2, col, n - 1, 1).getValues().map(r => String(r[0]).trim());
  if (v.some(x => x.toLowerCase() === value.toLowerCase())) return;
  const empty = v.findIndex(x => !x);
  sh.getRange(empty === -1 ? n + 1 : empty + 2, col).setValue(value);
}
