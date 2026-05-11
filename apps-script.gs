/**
 * DANISH DIARY — APPS SCRIPT BACKEND (v2)
 * ============================================
 * Personal khata / receivables tracker with multi-currency support.
 *
 * Sheets used (auto-created/migrated on first call):
 *   People: personId, name, phone, currency, notes, archived, createdAt
 *   Ledger: entryId, personId, date, type, amount, category,
 *           description, dueDate, deletedAt, createdAt
 *
 * Auth model: a single shared token. Anyone with the deployed URL +
 * token can read/write. Fine for personal use behind an unguessable
 * Vercel URL.
 *
 * The script must be bound to a Google Sheet (Extensions → Apps
 * Script from inside a Sheet). Standalone scripts can't use
 * getActiveSpreadsheet().
 */

const SECRET_TOKEN  = 'diary-7Nq2Pk5Hm8Rt3vXz9wL';
const PEOPLE_SHEET  = 'People';
const LEDGER_SHEET  = 'Ledger';
const CACHE_TTL_S   = 20;
const CACHE_KEY     = 'diary_dashboard_v2';

const PEOPLE_HEADERS = ['personId', 'name', 'phone', 'currency', 'notes', 'archived', 'createdAt'];
const LEDGER_HEADERS = ['entryId', 'personId', 'date', 'type', 'amount', 'category', 'description', 'dueDate', 'deletedAt', 'createdAt'];

const DEFAULT_CURRENCY = 'PKR';
const ALLOWED_CURRENCIES = ['PKR','USD','EUR','GBP','AED','SAR','INR','CAD','AUD','JPY','CNY','TRY','BDT','LKR'];
const ALLOWED_CATEGORIES = ['loan','advance','business','personal','refund','goods','services','other'];

// ============================================
// DISPATCH
// ============================================

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'health';
  if (action === 'health') {
    return jsonResponse({ ok: true, status: 'Danish Diary backend is live', version: 2, timestamp: new Date().toISOString() });
  }
  try {
    if (!authOk_(e && e.parameter && e.parameter.token)) {
      return jsonResponse({ ok: false, error: 'Unauthorized' });
    }
    if (action === 'dashboard')  return jsonResponse({ ok: true, data: getDashboardCached_() });
    if (action === 'personDetail') return jsonResponse({ ok: true, data: getPersonDetail_(e.parameter.personId) });
    if (action === 'ledger')     return jsonResponse({ ok: true, data: getFullLedger_() });
    return jsonResponse({ ok: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonResponse({ ok: false, error: 'Server error: ' + (err.message || String(err)) });
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (!authOk_(data.token)) return jsonResponse({ ok: false, error: 'Unauthorized' });
    const action = data.action;
    if (!action) return jsonResponse({ ok: false, error: 'Missing action' });
    if (action === 'addPerson')     return handleAddPerson_(data);
    if (action === 'updatePerson')  return handleUpdatePerson_(data);
    if (action === 'archivePerson') return handleArchivePerson_(data);
    if (action === 'addEntry')      return handleAddEntry_(data);
    if (action === 'updateEntry')   return handleUpdateEntry_(data);
    if (action === 'deleteEntry')   return handleDeleteEntry_(data);
    if (action === 'restoreEntry')  return handleRestoreEntry_(data);
    if (action === 'purgeEntry')    return handlePurgeEntry_(data);
    return jsonResponse({ ok: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonResponse({ ok: false, error: 'Server error: ' + (err.message || String(err)) });
  }
}

function authOk_(token) { return token === SECRET_TOKEN; }

// ============================================
// SHEETS — ensure they exist with current headers; migrate if needed
// ============================================

function getSheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Script is not bound to a spreadsheet. Open the target Google Sheet → Extensions → Apps Script, paste this code there, and re-deploy.');
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    return sheet;
  }
  // Migrate: append any new headers as columns at the end.
  const lastCol = sheet.getLastColumn();
  const existing = sheet.getRange(1, 1, 1, Math.max(lastCol, 1)).getValues()[0].map(String);
  let added = false;
  for (let i = 0; i < headers.length; i++) {
    if (existing[i] !== headers[i]) {
      // Either missing or wrong name at this position
      if (i < lastCol && existing[i] && existing[i] !== headers[i]) {
        // Header at this position differs — leave as-is? Safer: insert column.
        // To stay backward-compatible we'll just write the expected name.
        sheet.getRange(1, i + 1).setValue(headers[i]);
        added = true;
      } else {
        sheet.getRange(1, i + 1).setValue(headers[i]);
        added = true;
      }
    }
  }
  if (added) sheet.setFrozenRows(1);
  return sheet;
}

function getPeopleSheet_() { return getSheet_(PEOPLE_SHEET, PEOPLE_HEADERS); }
function getLedgerSheet_() { return getSheet_(LEDGER_SHEET, LEDGER_HEADERS); }

function colIdx_(header, name) {
  const i = header.indexOf(name);
  return i < 0 ? -1 : i;
}

// ============================================
// READS
// ============================================

function readPeople_() {
  const sheet = getPeopleSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const width = PEOPLE_HEADERS.length;
  const values = sheet.getRange(2, 1, last - 1, width).getValues();
  const H = PEOPLE_HEADERS;
  const idx = {
    personId: colIdx_(H, 'personId'),
    name: colIdx_(H, 'name'),
    phone: colIdx_(H, 'phone'),
    currency: colIdx_(H, 'currency'),
    notes: colIdx_(H, 'notes'),
    archived: colIdx_(H, 'archived'),
    createdAt: colIdx_(H, 'createdAt')
  };
  return values
    .map((r, i) => ({
      row: i + 2,
      personId: String(r[idx.personId] || ''),
      name: String(r[idx.name] || ''),
      phone: String(r[idx.phone] || ''),
      currency: normCurrency_(r[idx.currency]),
      notes: String(r[idx.notes] || ''),
      archived: toBool_(r[idx.archived]),
      createdAt: r[idx.createdAt]
    }))
    .filter((p) => p.personId);
}

function readLedger_(includeDeleted) {
  const sheet = getLedgerSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const width = LEDGER_HEADERS.length;
  const values = sheet.getRange(2, 1, last - 1, width).getValues();
  const H = LEDGER_HEADERS;
  const idx = {
    entryId: colIdx_(H, 'entryId'),
    personId: colIdx_(H, 'personId'),
    date: colIdx_(H, 'date'),
    type: colIdx_(H, 'type'),
    amount: colIdx_(H, 'amount'),
    category: colIdx_(H, 'category'),
    description: colIdx_(H, 'description'),
    dueDate: colIdx_(H, 'dueDate'),
    deletedAt: colIdx_(H, 'deletedAt'),
    createdAt: colIdx_(H, 'createdAt')
  };
  const out = values
    .map((r, i) => ({
      row: i + 2,
      entryId: String(r[idx.entryId] || ''),
      personId: String(r[idx.personId] || ''),
      date: r[idx.date],
      type: String(r[idx.type] || '').toLowerCase(),
      amount: parseFloat(r[idx.amount]) || 0,
      category: normCategory_(r[idx.category]),
      description: String(r[idx.description] || ''),
      dueDate: r[idx.dueDate] instanceof Date ? r[idx.dueDate] : null,
      deletedAt: r[idx.deletedAt] instanceof Date ? r[idx.deletedAt] : null,
      createdAt: r[idx.createdAt]
    }))
    .filter((e) => e.entryId && e.personId);
  return includeDeleted ? out : out.filter((e) => !e.deletedAt);
}

// ============================================
// DASHBOARD
// ============================================

function getDashboardCached_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CACHE_KEY);
  if (cached) return JSON.parse(cached);
  const data = buildDashboard_();
  try { cache.put(CACHE_KEY, JSON.stringify(data), CACHE_TTL_S); } catch (e) {}
  return data;
}

function buildDashboard_() {
  const people = readPeople_();
  const ledger = readLedger_();
  const tz = Session.getScriptTimeZone();
  const today = startOfDay_(new Date());

  const byPerson = {};
  for (const e of ledger) {
    if (!byPerson[e.personId]) byPerson[e.personId] = [];
    byPerson[e.personId].push(e);
  }

  const computed = people.map((p) => {
    const entries = (byPerson[p.personId] || []).slice().sort(byDateAsc_);
    let totalCharged = 0, totalRepaid = 0;
    let lastActivity = null, lastRepayment = null, firstCharge = null;
    let overdueCount = 0, overdueAmount = 0;
    for (const e of entries) {
      if (e.type === 'charge') {
        totalCharged += e.amount;
        if (!firstCharge && e.date instanceof Date) firstCharge = e.date;
        if (e.dueDate instanceof Date && e.dueDate.getTime() < today.getTime()) {
          overdueCount += 1;
          overdueAmount += e.amount;
        }
      } else if (e.type === 'repayment') {
        totalRepaid += e.amount;
        if (e.date instanceof Date && (!lastRepayment || e.date.getTime() > lastRepayment.getTime())) {
          lastRepayment = e.date;
        }
      }
      if (e.date instanceof Date && (!lastActivity || e.date.getTime() > lastActivity.getTime())) {
        lastActivity = e.date;
      }
    }
    const balance = round2_(totalCharged - totalRepaid);
    const anchor = lastRepayment || firstCharge;
    const daysSincePayment = anchor ? daysBetween_(anchor, today) : 0;
    return {
      personId: p.personId,
      name: p.name,
      phone: p.phone,
      currency: p.currency,
      notes: p.notes,
      archived: p.archived,
      totalCharged: round2_(totalCharged),
      totalRepaid: round2_(totalRepaid),
      balance: balance,
      entryCount: entries.length,
      overdueCount: overdueCount,
      overdueAmount: round2_(Math.max(0, overdueAmount - totalRepaid > 0 ? overdueAmount : overdueAmount)),
      lastActivity: lastActivity ? fmtDate_(lastActivity, tz) : '',
      lastRepayment: lastRepayment ? fmtDate_(lastRepayment, tz) : '',
      firstCharge: firstCharge ? fmtDate_(firstCharge, tz) : '',
      daysSincePayment: daysSincePayment
    };
  });

  computed.sort((a, b) => {
    if (a.archived !== b.archived) return a.archived ? 1 : -1;
    const aOwes = a.balance > 0, bOwes = b.balance > 0;
    if (aOwes !== bOwes) return aOwes ? -1 : 1;
    if (a.daysSincePayment !== b.daysSincePayment) return b.daysSincePayment - a.daysSincePayment;
    return (a.name || '').localeCompare(b.name || '');
  });

  const outstandingByCurrency = {};
  const peopleByCurrency = {};
  const aging = { fresh: 0, week: 0, month: 0, twoMonth: 0, old: 0 };
  let totalOverdueAcrossPeople = 0;
  for (const p of computed) {
    if (p.archived) continue;
    if (p.balance > 0) {
      outstandingByCurrency[p.currency] = round2_((outstandingByCurrency[p.currency] || 0) + p.balance);
      peopleByCurrency[p.currency] = (peopleByCurrency[p.currency] || 0) + 1;
      const d = p.daysSincePayment;
      if (d <= 7)       aging.fresh += 1;
      else if (d <= 30) aging.week += 1;
      else if (d <= 60) aging.month += 1;
      else if (d <= 90) aging.twoMonth += 1;
      else              aging.old += 1;
    }
    if (p.overdueCount > 0) totalOverdueAcrossPeople += 1;
  }

  const peopleById = {};
  for (const p of computed) peopleById[p.personId] = p;
  const recent = ledger
    .slice()
    .sort((a, b) => {
      const ta = a.date instanceof Date ? a.date.getTime() : 0;
      const tb = b.date instanceof Date ? b.date.getTime() : 0;
      if (tb !== ta) return tb - ta;
      const ca = a.createdAt instanceof Date ? a.createdAt.getTime() : 0;
      const cb = b.createdAt instanceof Date ? b.createdAt.getTime() : 0;
      return cb - ca;
    })
    .slice(0, 50)
    .map((e) => ({
      entryId: e.entryId,
      personId: e.personId,
      personName: (peopleById[e.personId] && peopleById[e.personId].name) || '(unknown)',
      currency: (peopleById[e.personId] && peopleById[e.personId].currency) || DEFAULT_CURRENCY,
      date: e.date instanceof Date ? fmtDate_(e.date, tz) : '',
      type: e.type,
      amount: round2_(e.amount),
      category: e.category,
      description: e.description,
      dueDate: e.dueDate ? fmtDate_(e.dueDate, tz) : ''
    }));

  return {
    summary: {
      outstandingByCurrency: outstandingByCurrency,
      peopleByCurrency: peopleByCurrency,
      peopleWithBalance: Object.values(peopleByCurrency).reduce((a, b) => a + b, 0),
      totalPeople: computed.filter((p) => !p.archived).length,
      overduePeople: totalOverdueAcrossPeople,
      aging: aging
    },
    people: computed,
    recent: recent
  };
}

function getPersonDetail_(personId) {
  const tz = Session.getScriptTimeZone();
  const people = readPeople_();
  const person = people.find((p) => p.personId === personId);
  if (!person) throw new Error('Person not found');
  const ledger = readLedger_().filter((e) => e.personId === personId);
  ledger.sort(byDateDesc_);
  let totalCharged = 0, totalRepaid = 0;
  for (const e of ledger) {
    if (e.type === 'charge') totalCharged += e.amount;
    else if (e.type === 'repayment') totalRepaid += e.amount;
  }
  return {
    personId: person.personId,
    name: person.name,
    phone: person.phone,
    currency: person.currency,
    notes: person.notes,
    archived: person.archived,
    totalCharged: round2_(totalCharged),
    totalRepaid: round2_(totalRepaid),
    balance: round2_(totalCharged - totalRepaid),
    entries: ledger.map((e) => ({
      entryId: e.entryId,
      date: e.date instanceof Date ? fmtDate_(e.date, tz) : '',
      type: e.type,
      amount: round2_(e.amount),
      category: e.category,
      description: e.description,
      dueDate: e.dueDate ? fmtDate_(e.dueDate, tz) : ''
    }))
  };
}

function getFullLedger_() {
  const tz = Session.getScriptTimeZone();
  const people = readPeople_();
  const peopleById = {};
  for (const p of people) peopleById[p.personId] = p;
  const ledger = readLedger_(true);
  return ledger.map((e) => ({
    entryId: e.entryId,
    personId: e.personId,
    personName: (peopleById[e.personId] && peopleById[e.personId].name) || '(unknown)',
    currency: (peopleById[e.personId] && peopleById[e.personId].currency) || DEFAULT_CURRENCY,
    date: e.date instanceof Date ? fmtDate_(e.date, tz) : '',
    type: e.type,
    amount: round2_(e.amount),
    category: e.category,
    description: e.description,
    dueDate: e.dueDate ? fmtDate_(e.dueDate, tz) : '',
    deletedAt: e.deletedAt ? Utilities.formatDate(e.deletedAt, tz, 'yyyy-MM-dd HH:mm') : ''
  }));
}

// ============================================
// WRITES — people
// ============================================

function handleAddPerson_(data) {
  const name = String(data.name || '').trim();
  if (!name) return jsonResponse({ ok: false, error: 'Name required' });
  const phone = normalizePhone_(data.phone);
  const currency = normCurrency_(data.currency);
  const notes = String(data.notes || '').trim();
  const sheet = getPeopleSheet_();
  const personId = nextPersonId_(sheet);
  // Match PEOPLE_HEADERS column order.
  sheet.appendRow([personId, name, phone, currency, notes, false, new Date()]);
  bustCache_();
  return jsonResponse({ ok: true, personId: personId });
}

function handleUpdatePerson_(data) {
  const personId = String(data.personId || '');
  if (!personId) return jsonResponse({ ok: false, error: 'personId required' });
  const sheet = getPeopleSheet_();
  const row = findRowByValue_(sheet, PEOPLE_HEADERS.indexOf('personId') + 1, personId);
  if (!row) return jsonResponse({ ok: false, error: 'Person not found' });
  const set = (col, val) => sheet.getRange(row, PEOPLE_HEADERS.indexOf(col) + 1).setValue(val);
  if (data.name !== undefined)     set('name', String(data.name).trim());
  if (data.phone !== undefined)    set('phone', normalizePhone_(data.phone));
  if (data.currency !== undefined) set('currency', normCurrency_(data.currency));
  if (data.notes !== undefined)    set('notes', String(data.notes).trim());
  bustCache_();
  return jsonResponse({ ok: true });
}

function handleArchivePerson_(data) {
  const personId = String(data.personId || '');
  if (!personId) return jsonResponse({ ok: false, error: 'personId required' });
  const sheet = getPeopleSheet_();
  const row = findRowByValue_(sheet, PEOPLE_HEADERS.indexOf('personId') + 1, personId);
  if (!row) return jsonResponse({ ok: false, error: 'Person not found' });
  const archived = data.archived === false ? false : true;
  sheet.getRange(row, PEOPLE_HEADERS.indexOf('archived') + 1).setValue(archived);
  bustCache_();
  return jsonResponse({ ok: true });
}

// ============================================
// WRITES — ledger entries
// ============================================

function handleAddEntry_(data) {
  const personId = String(data.personId || '');
  if (!personId) return jsonResponse({ ok: false, error: 'personId required' });
  const type = String(data.type || '').toLowerCase();
  if (type !== 'charge' && type !== 'repayment') {
    return jsonResponse({ ok: false, error: "type must be 'charge' or 'repayment'" });
  }
  const amount = parseFloat(data.amount);
  if (isNaN(amount) || amount <= 0) {
    return jsonResponse({ ok: false, error: 'Amount must be a positive number' });
  }
  const date = parseLocalDate_(data.date);
  if (!date) return jsonResponse({ ok: false, error: 'Invalid date (expected YYYY-MM-DD)' });
  const dueDate = data.dueDate ? parseLocalDate_(data.dueDate) : null;
  const category = normCategory_(data.category);
  const description = String(data.description || '').trim();

  const people = readPeople_();
  if (!people.find((p) => p.personId === personId)) {
    return jsonResponse({ ok: false, error: 'Person not found' });
  }

  const sheet = getLedgerSheet_();
  const entryId = nextEntryId_(sheet);
  // Match LEDGER_HEADERS column order.
  sheet.appendRow([entryId, personId, date, type, amount, category, description, dueDate || '', '', new Date()]);

  const newRow = sheet.getLastRow();
  const dateCol = LEDGER_HEADERS.indexOf('date') + 1;
  const amtCol = LEDGER_HEADERS.indexOf('amount') + 1;
  const dueCol = LEDGER_HEADERS.indexOf('dueDate') + 1;
  const createdCol = LEDGER_HEADERS.indexOf('createdAt') + 1;
  sheet.getRange(newRow, dateCol).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(newRow, amtCol).setNumberFormat('#,##0.00');
  sheet.getRange(newRow, createdCol).setNumberFormat('yyyy-mm-dd hh:mm');
  if (dueDate) sheet.getRange(newRow, dueCol).setNumberFormat('yyyy-mm-dd');

  bustCache_();
  return jsonResponse({ ok: true, entryId: entryId });
}

function handleUpdateEntry_(data) {
  const entryId = String(data.entryId || '');
  if (!entryId) return jsonResponse({ ok: false, error: 'entryId required' });
  const sheet = getLedgerSheet_();
  const row = findRowByValue_(sheet, LEDGER_HEADERS.indexOf('entryId') + 1, entryId);
  if (!row) return jsonResponse({ ok: false, error: 'Entry not found' });
  const set = (col, val) => sheet.getRange(row, LEDGER_HEADERS.indexOf(col) + 1).setValue(val);
  if (data.date !== undefined) {
    const d = parseLocalDate_(data.date);
    if (!d) return jsonResponse({ ok: false, error: 'Invalid date' });
    set('date', d);
  }
  if (data.amount !== undefined) {
    const a = parseFloat(data.amount);
    if (isNaN(a) || a <= 0) return jsonResponse({ ok: false, error: 'Invalid amount' });
    set('amount', a);
  }
  if (data.category !== undefined)    set('category', normCategory_(data.category));
  if (data.description !== undefined) set('description', String(data.description).trim());
  if (data.dueDate !== undefined)     set('dueDate', data.dueDate ? parseLocalDate_(data.dueDate) : '');
  bustCache_();
  return jsonResponse({ ok: true });
}

function handleDeleteEntry_(data) {
  // Soft delete — set deletedAt timestamp.
  const entryId = String(data.entryId || '');
  if (!entryId) return jsonResponse({ ok: false, error: 'entryId required' });
  const sheet = getLedgerSheet_();
  const row = findRowByValue_(sheet, LEDGER_HEADERS.indexOf('entryId') + 1, entryId);
  if (!row) return jsonResponse({ ok: false, error: 'Entry not found' });
  sheet.getRange(row, LEDGER_HEADERS.indexOf('deletedAt') + 1).setValue(new Date());
  bustCache_();
  return jsonResponse({ ok: true });
}

function handleRestoreEntry_(data) {
  const entryId = String(data.entryId || '');
  if (!entryId) return jsonResponse({ ok: false, error: 'entryId required' });
  const sheet = getLedgerSheet_();
  const row = findRowByValue_(sheet, LEDGER_HEADERS.indexOf('entryId') + 1, entryId);
  if (!row) return jsonResponse({ ok: false, error: 'Entry not found' });
  sheet.getRange(row, LEDGER_HEADERS.indexOf('deletedAt') + 1).setValue('');
  bustCache_();
  return jsonResponse({ ok: true });
}

function handlePurgeEntry_(data) {
  // Hard delete (only for already soft-deleted entries).
  const entryId = String(data.entryId || '');
  if (!entryId) return jsonResponse({ ok: false, error: 'entryId required' });
  const sheet = getLedgerSheet_();
  const row = findRowByValue_(sheet, LEDGER_HEADERS.indexOf('entryId') + 1, entryId);
  if (!row) return jsonResponse({ ok: false, error: 'Entry not found' });
  sheet.deleteRow(row);
  bustCache_();
  return jsonResponse({ ok: true });
}

// ============================================
// IDs + HELPERS
// ============================================

function nextPersonId_(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return 'p001';
  const col = PEOPLE_HEADERS.indexOf('personId') + 1;
  const values = sheet.getRange(2, col, last - 1, 1).getValues();
  let maxN = 0;
  for (const r of values) {
    const m = String(r[0] || '').match(/^p(\d+)$/i);
    if (m) maxN = Math.max(maxN, parseInt(m[1], 10));
  }
  return 'p' + String(maxN + 1).padStart(3, '0');
}

function nextEntryId_(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return 'e00001';
  const col = LEDGER_HEADERS.indexOf('entryId') + 1;
  const values = sheet.getRange(2, col, last - 1, 1).getValues();
  let maxN = 0;
  for (const r of values) {
    const m = String(r[0] || '').match(/^e(\d+)$/i);
    if (m) maxN = Math.max(maxN, parseInt(m[1], 10));
  }
  return 'e' + String(maxN + 1).padStart(5, '0');
}

function findRowByValue_(sheet, col, target) {
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const values = sheet.getRange(2, col, last - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(target)) return i + 2;
  }
  return 0;
}

function parseLocalDate_(s) {
  if (!s) return null;
  if (s instanceof Date) return s;
  const parts = String(s).split('-');
  if (parts.length !== 3) return null;
  const d = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
  return isNaN(d.getTime()) ? null : d;
}

function fmtDate_(d, tz) { return Utilities.formatDate(d, tz, 'yyyy-MM-dd'); }
function startOfDay_(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function daysBetween_(a, b) {
  const ms = startOfDay_(b).getTime() - startOfDay_(a).getTime();
  return Math.floor(ms / (1000 * 60 * 60 * 24));
}
function byDateAsc_(a, b) {
  const ta = a.date instanceof Date ? a.date.getTime() : 0;
  const tb = b.date instanceof Date ? b.date.getTime() : 0;
  return ta - tb;
}
function byDateDesc_(a, b) { return byDateAsc_(b, a); }
function round2_(n) { return Math.round((Number(n) || 0) * 100) / 100; }
function toBool_(v) {
  if (v === true) return true;
  if (v === false) return false;
  const s = String(v || '').toLowerCase();
  return s === 'true' || s === '1' || s === 'yes';
}
function normalizePhone_(p) { return String(p || '').replace(/[^\d]/g, ''); }
function normCurrency_(c) {
  const s = String(c || '').toUpperCase().trim();
  if (ALLOWED_CURRENCIES.indexOf(s) >= 0) return s;
  return DEFAULT_CURRENCY;
}
function normCategory_(c) {
  const s = String(c || '').toLowerCase().trim();
  if (ALLOWED_CATEGORIES.indexOf(s) >= 0) return s;
  return '';
}
function bustCache_() {
  try { CacheService.getScriptCache().remove(CACHE_KEY); } catch (e) {}
}
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
