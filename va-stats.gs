// ============================================================
// VA Weekly Stats — Google Apps Script
// ============================================================
// INSTALL:
//  1. Extensions > Apps Script in the weekly report spreadsheet
//  2. Delete all existing code, paste this file, Save
//  3. Run setupWeeklyTrigger() ONCE → schedules auto-run every Monday 3 PM
//  4. Run refreshAll() now to populate immediately
//
// CREATES / UPDATES (weekly report spreadsheet only):
//  "current week leads"  — raw lead rows for prev Sun–Sat
//  "current week stats"  — per-account summary with rates, rating, notes
//
// MTD VA Stats and Follow-Up Adherence live in a separate script.
// ============================================================

// ── IDs & SHEET NAMES ───────────────────────────────────────
const SOURCE_SS_ID = '1Kd0smsW26IJR1H5kJ0MP0kH8jPOTazo2_TbLEkgmbRs';
const MAIN_SHEETS  = ['ROOF, MAIN', 'HVAC, MAIN', 'GUTTER, Main', 'WINDOWS, MAIN'];

// ── SOURCE COLUMN INDICES (0-based, A=0 … O=14) ─────────────
const COL_VA         = 0;   // A
const COL_DATE_IN    = 1;   // B — date lead came in
const COL_SUBACCOUNT = 2;   // C
const COL_LEAD_NAME  = 3;   // D
const COL_STATUS     = 6;   // G — status
const COL_SCHED_DATE = 7;   // H
const COL_DISP       = 8;   // I — disposition / cancellation reason
const COL_DATE_CONF  = 9;   // J — date appt confirmed
const COL_FU1        = 10;  // K–O follow-up notes (5 slots)
const COL_FU5        = 14;

// ── STATUS BUCKETS ───────────────────────────────────────────
function getStatusBucket(status) {
  const s = status.toLowerCase().trim();
  if (s === 'confirmed' || s === 'manual booked')           return 'booked';
  if (s.includes('cancelled') || s.includes('invalid'))    return 'cancelled';
  if (['not booked','unconfirmed','not responding',
       'call back'].includes(s))                            return 'notResponding';
  if (s === 'client handles' || s === 'satellite quote')    return 'clientHandles';
  if (s === 'reschedule needed')                            return 'reschedule';
  return 'other';
}

const BOOKING_KPI = 0.45;
const EXCLUDE_CLIENT_HANDLES_ACCTS = ['Outstanding Roofing', 'Good Guy Roofing'];

// ============================================================
// ENTRY POINT
// ============================================================

function refreshAll() {
  const today    = new Date();
  const prevWeek = getPrevWeekRange(today);

  Logger.log('Loading leads from source sheets…');
  const { leads, rawRows } = loadAllLeads();
  Logger.log('Total rows loaded: ' + leads.length);

  const weekLeads   = leads.filter(l => inRange(l.dateIn, prevWeek) || inRange(l.dateConf, prevWeek));
  const weekRawRows = rawRows.filter((_, i) => {
    const l = leads[i];
    return inRange(l.dateIn, prevWeek) || inRange(l.dateConf, prevWeek);
  });

  writeCurrentWeekLeads(weekLeads, weekRawRows, prevWeek, today);
  updateCurrentWeekStats(weekLeads, prevWeek, today);

  Logger.log('Done!');
  SpreadsheetApp.getUi().alert('Done! Weekly sheets have been refreshed.');
}

// ============================================================
// DATE HELPERS
// ============================================================

function getPrevWeekRange(today) {
  const d      = today.getDay();
  const endD   = new Date(today); endD.setDate(today.getDate() - d - 1); endD.setHours(23,59,59,999);
  const startD = new Date(endD);  startD.setDate(endD.getDate() - 6);    startD.setHours(0,0,0,0);
  return { start: startD, end: endD };
}

function parseDate(v) {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  const d = new Date(v); return isNaN(d.getTime()) ? null : d;
}

function inRange(d, r) { return !!d && d >= r.start && d <= r.end; }
function pct(n, d)     { return d ? Math.round(n / d * 100) + '%' : '0%'; }

function fmtDate(d) {
  if (!d || isNaN(d.getTime())) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'MMM d, yyyy');
}
function fmtShort(d) {
  if (!d || isNaN(d.getTime())) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'M/d/yyyy');
}

// ============================================================
// DATA LOADING
// ============================================================

function loadAllLeads() {
  const ss      = SpreadsheetApp.openById(SOURCE_SS_ID);
  const leads   = [];
  const rawRows = [];

  MAIN_SHEETS.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) { Logger.log('WARNING: sheet not found — "' + name + '"'); return; }
    const last = sheet.getLastRow();
    if (last < 2) return;

    const data = sheet.getRange(2, 1, last - 1, 15).getValues();
    data.forEach(row => {
      const va = String(row[COL_VA] || '').trim();
      if (!va || va.toLowerCase() === 'va') return;

      const status   = String(row[COL_STATUS] || '').trim();
      const disp     = String(row[COL_DISP]   || '').trim().toLowerCase();
      const bucket   = getStatusBucket(status);

      leads.push({
        va,
        status: status.toLowerCase().trim(),
        disp,
        bucket,
        isCancelledInvalid: bucket === 'cancelled',
        dateIn:     parseDate(row[COL_DATE_IN]),
        dateConf:   parseDate(row[COL_DATE_CONF]),
        schedDate:  parseDate(row[COL_SCHED_DATE]),
        subaccount: String(row[COL_SUBACCOUNT] || '').trim(),
        leadName:   String(row[COL_LEAD_NAME]  || '').trim(),
        followUps:  Array.from({length:5}, (_,k) => String(row[COL_FU1+k]||'').trim()),
        sheet: name
      });
      rawRows.push(row);
    });
  });

  return { leads, rawRows };
}

function getAccountList() {
  try {
    const ss   = SpreadsheetApp.openById(SOURCE_SS_ID);
    const dash = ss.getSheetByName('DASHBOARD');
    if (!dash) return null;
    const last = dash.getLastRow();
    if (last < 7) return null;
    const accCol = dash.getRange(7, 2, last - 6, 1).getValues();
    const vaCol  = dash.getRange(7, 4, last - 6, 1).getValues();
    const list   = [];
    accCol.forEach((r, i) => {
      const acc = String(r[0]          || '').trim();
      const va  = String(vaCol[i][0]   || '').trim();
      if (acc) list.push({ va, subaccount: acc });
    });
    return list.length ? list : null;
  } catch(e) {
    Logger.log('Could not read DASHBOARD: ' + e.message);
    return null;
  }
}

// ============================================================
// SHARED HELPERS
// ============================================================

function resetSheet(destSS, name) {
  const ex = destSS.getSheetByName(name);
  if (ex) destSS.deleteSheet(ex);
  return destSS.insertSheet(name);
}

function writeGrid(sheet, data, W) {
  if (!data.length) return;
  const padded = data.map(r => {
    const p = r.map(v => v == null ? '' : v);
    while (p.length < W) p.push('');
    return p.slice(0, W);
  });
  sheet.getRange(1, 1, padded.length, W).setValues(padded);
}

function styleRow(sheet, row, W, bg, fg, bold, sz) {
  const r = sheet.getRange(row, 1, 1, W);
  if (bg)   r.setBackground(bg);
  if (fg)   r.setFontColor(fg);
  if (bold) r.setFontWeight('bold');
  if (sz)   r.setFontSize(sz);
}

// ============================================================
// SHEET A: "current week leads"
// ============================================================

function writeCurrentWeekLeads(weekLeads, weekRawRows, weekRange, today) {
  const destSS = SpreadsheetApp.getActiveSpreadsheet();
  const sheet  = resetSheet(destSS, 'current week leads');
  const label  = fmtDate(weekRange.start) + ' – ' + fmtDate(weekRange.end);
  const W = 15;

  const HEADER = [
    'VA','Date','Sub-account','Lead Name','Address',
    'Distance & Drive Time','Status','Scheduled Date',
    'Disposition','Date Confirmed',
    '1st Call - Follow up','2nd Call - Follow up',
    '3rd Call - Follow up','4th Call - Follow up','5th Call - Follow up'
  ];

  const titleRow = ['CURRENT WEEK LEADS — ' + label + ' — Generated ' + fmtDate(today)];
  titleRow.push(...Array(W - 1).fill(''));

  const OUT = [titleRow, HEADER];
  weekRawRows.forEach(row => {
    OUT.push(Array.from({length: W}, (_, i) => {
      const v = row[i];
      if (v instanceof Date && !isNaN(v.getTime())) return fmtShort(v);
      return v == null ? '' : v;
    }));
  });

  writeGrid(sheet, OUT, W);
  styleRow(sheet, 1, W, '#1F3864', '#FFFFFF', true, 11);
  styleRow(sheet, 2, W, '#2F5496', '#FFFFFF', true, 10);
  sheet.setFrozenRows(2);
  sheet.autoResizeColumns(1, W);
  Logger.log('current week leads: ' + weekLeads.length + ' rows written');
}

// ============================================================
// SHEET B: "current week stats"
// ============================================================

const CW_W           = 21;
const CW_VA          = 1;
const CW_SUBACCT     = 2;
const CW_BOOKED      = 3;
const CW_CANCELLED   = 4;
const CW_NOTRESP     = 5;
const CW_CLIENT      = 6;
const CW_RESCHED     = 7;
const CW_TOTAL       = 8;
const CW_BOOK_RATE   = 9;
const CW_CANCEL_RATE = 10;
const CW_NR_RATE     = 11;
const CW_CLIENT_RATE = 12;
const CW_SUBACCT2    = 13;
const CW_OSA         = 14;
const CW_FAKE        = 15;
const CW_NOTINT      = 16;
const CW_UNQUAL      = 17;
const CW_NUMISSUE    = 18;
const CW_TOTAL2      = 19;
const CW_RATING      = 20;
const CW_NOTES       = 21;

function updateCurrentWeekStats(weekLeads, weekRange, today) {
  const destSS = SpreadsheetApp.getActiveSpreadsheet();
  const sheet  = resetSheet(destSS, 'current week stats');
  const label  = fmtDate(weekRange.start) + ' – ' + fmtDate(weekRange.end);

  const acctMap = {};
  weekLeads.forEach(l => {
    const key = l.subaccount || '(blank)';
    if (!acctMap[key]) acctMap[key] = {
      va: l.va, subaccount: key,
      booked:0, cancelled:0, notResponding:0,
      clientHandles:0, reschedule:0, other:0, total:0,
      osa:0, fakeLead:0, notInterested:0,
      unqualified:0, numIssue:0, dupLead:0
    };
    const s = acctMap[key];
    s.total++;
    s[l.bucket] = (s[l.bucket] || 0) + 1;
    const d = l.disp;
    if      (d === 'osa')                        s.osa++;
    else if (d === 'fake lead' || d === 'fake')  s.fakeLead++;
    else if (d === 'not interested')             s.notInterested++;
    else if (d === 'unqualified')                s.unqualified++;
    else if (d === '# issue' || d === '#issue')  s.numIssue++;
    else if (d === 'dup lead')                   s.dupLead++;
  });

  const dashList     = getAccountList();
  const dashVALookup = {};
  if (dashList) dashList.forEach(x => { dashVALookup[x.subaccount] = x.va; });

  let acctKeys;
  if (dashList) {
    acctKeys = dashList.map(x => x.subaccount);
    dashList.forEach(x => {
      if (acctMap[x.subaccount] && !acctMap[x.subaccount].va)
        acctMap[x.subaccount].va = x.va;
    });
  } else {
    acctKeys = Object.keys(acctMap).sort();
  }

  const titleRow = Array(CW_W).fill('');
  titleRow[0] = 'WEEKLY LEAD REPORT — ' + label + ' — Generated ' + fmtDate(today);

  const subRow = Array(CW_W).fill('');
  subRow[0] = 'Leads included: Col B (date entered) OR Col J (appt confirmed) falls in prev Sun–Sat.';
  subRow[CW_SUBACCT2 - 1] = 'Disposition breakdown — rate of total leads per account.';

  const grpRow = Array(CW_W).fill('');
  grpRow[CW_VA - 1]     = '← MAIN STATS';
  grpRow[CW_SUBACCT2-1] = '← CANCELLED / INVALID BREAKDOWN';
  grpRow[CW_NOTES - 1]  = 'NOTES →';

  const hdrRow = [
    'VA','Sub-account',
    'Confirmed /\nManual Booked','Cancelled /\nInvalid','Not\nResponding',
    'Client\nHandles','Needs\nRescheduling','Total',
    'Booking\nRate','Cancelled/\nInvalid Rate','No Response\nRate','Client Handles\nRate',
    'Sub-account','Out of SA\n(OSA)','Fake\nLead','Not\nInterested',
    'UNQUALIFIED','Number\nIssue','Total','Overall\nRating','Notes'
  ];

  const dataRows = acctKeys.map(key => {
    const s = acctMap[key] || { va: dashVALookup[key] || '', subaccount: key,
      booked:0, cancelled:0, notResponding:0, clientHandles:0, reschedule:0, total:0,
      osa:0, fakeLead:0, notInterested:0, unqualified:0, numIssue:0, dupLead:0 };
    if (!s.va) s.va = dashVALookup[key] || '';

    const noResp        = (s.notResponding || 0) + (s.reschedule || 0);
    const total         = s.total || 0;
    const isSpecialAcct = EXCLUDE_CLIENT_HANDLES_ACCTS.includes(key);
    const bookingDenom  = isSpecialAcct ? Math.max(0, total - (s.clientHandles || 0)) : total;
    const { rating, notes } = scoreAccount(s, noResp, bookingDenom);

    return [
      s.va, key,
      s.booked       || 0,
      s.cancelled    || 0,
      noResp,
      s.clientHandles|| 0,
      s.reschedule   || 0,
      total,
      pct(s.booked       || 0, bookingDenom),
      pct(s.cancelled    || 0, total),
      pct(noResp,            total),
      pct(s.clientHandles|| 0, total),
      key,
      pct(s.osa          || 0, total),
      pct(s.fakeLead     || 0, total),
      pct(s.notInterested|| 0, total),
      pct(s.unqualified  || 0, total),
      pct(s.numIssue     || 0, total),
      total,
      rating,
      notes
    ];
  });

  const gt = Array(CW_W).fill('');
  gt[0] = 'TOTAL';
  [2,3,4,5,6,7,18].forEach(i => { gt[i] = dataRows.reduce((a,r) => a + (Number(r[i])||0), 0); });
  gt[8]  = pct(gt[2], gt[7]);
  gt[9]  = pct(gt[3], gt[7]);
  gt[10] = pct(gt[4], gt[7]);
  gt[11] = pct(gt[5], gt[7]);
  gt[12] = 'TOTAL';

  const OUT = [titleRow, subRow, grpRow, hdrRow, gt, ...dataRows];
  writeGrid(sheet, OUT, CW_W);

  styleRow(sheet, 1, CW_W, '#1F3864', '#FFFFFF', true, 11);
  styleRow(sheet, 2, CW_W, '#1F3864', '#CCDDFF', false, 9);
  styleRow(sheet, 3, CW_W, '#F2F2F2', '#666666', false, 8);
  sheet.getRange(3, CW_SUBACCT2, 1, 7).setBackground('#FFF8E7');
  sheet.getRange(3, CW_NOTES,    1, 1).setBackground('#EAF4EA');
  sheet.getRange(4, 1,           1, 12).setBackground('#2F5496').setFontColor('#FFFFFF').setFontWeight('bold');
  sheet.getRange(4, CW_SUBACCT2, 1, 7).setBackground('#BF8F00').setFontColor('#FFFFFF').setFontWeight('bold');
  sheet.getRange(4, CW_RATING,   1, 1).setBackground('#375623').setFontColor('#FFFFFF').setFontWeight('bold');
  sheet.getRange(4, CW_NOTES,    1, 1).setBackground('#375623').setFontColor('#FFFFFF').setFontWeight('bold');
  styleRow(sheet, 5, CW_W, '#F2F2F2', '#000000', true, 10);

  const dataStart = 6;
  for (let i = 0; i < dataRows.length; i++) {
    const row    = dataStart + i;
    const r      = dataRows[i];
    const bkPct  = parseInt(r[8]);
    const rating = r[19];

    if (!isNaN(bkPct) && r[7] > 0) {
      const bc = sheet.getRange(row, 9);
      if      (bkPct >= 50) bc.setBackground('#C6EFCE').setFontColor('#276221');
      else if (bkPct >= 30) bc.setBackground('#FFEB9C').setFontColor('#9C5700');
      else                  bc.setBackground('#FFC7CE').setFontColor('#9C0006');
    }

    const rc = sheet.getRange(row, 20);
    if      (rating === 3) rc.setBackground('#C6EFCE').setFontColor('#276221').setFontWeight('bold');
    else if (rating === 2) rc.setBackground('#FFEB9C').setFontColor('#9C5700').setFontWeight('bold');
    else if (rating === 1) rc.setBackground('#FFC7CE').setFontColor('#9C0006').setFontWeight('bold');
  }

  sheet.getRange(4, 1, 1, CW_W).setWrap(true);
  sheet.getRange(dataStart, CW_NOTES, dataRows.length, 1).setWrap(true);
  sheet.setColumnWidth(CW_VA,       70);
  sheet.setColumnWidth(CW_SUBACCT,  160);
  sheet.setColumnWidth(CW_SUBACCT2, 160);
  sheet.setColumnWidth(CW_NOTES,    350);
  [3,4,5,6,7,8,9,10,11,12,14,15,16,17,18,19,20].forEach(c => sheet.setColumnWidth(c, 72));
  sheet.setRowHeight(4, 50);
  sheet.setFrozenRows(4);
  sheet.setFrozenColumns(2);

  Logger.log('current week stats: ' + dataRows.length + ' accounts');
}

// ── Rating + Notes logic ─────────────────────────────────────
function scoreAccount(s, noResp, bookingDenom) {
  const total  = s.total     || 0;
  const booked = s.booked    || 0;
  const cancel = s.cancelled || 0;
  const denom  = (bookingDenom != null ? bookingDenom : total) || 0;

  if (total === 0) return { rating: 'N/A', notes: 'No lead came in.' };

  const bookPct   = denom ? booked / denom : 0;
  const cancelPct = total ? cancel / total : 0;
  const nrPct     = total ? noResp / total : 0;
  const osaPct    = (s.osa || 0) / total;

  let rating;
  if      (bookPct >= 0.50 && total >= 7) rating = 3;
  else if (bookPct >= 0.50)               rating = 2;
  else if (bookPct >= 0.30)               rating = 2;
  else                                    rating = 1;

  if (cancelPct > 0.45 && rating > 1) rating--;
  if (nrPct     > 0.50 && rating > 1) rating--;

  const parts = [];
  const br = Math.round(bookPct   * 100);
  const cr = Math.round(cancelPct * 100);
  const nr = Math.round(nrPct     * 100);
  const or = Math.round(osaPct    * 100);

  if      (bookPct >= 0.50) parts.push('Good lead flow with a high booked rate at ' + br + '%.');
  else if (bookPct >= 0.30) parts.push('Moderate booking rate at ' + br + '%.');
  else if (total > 0)       parts.push('Low booking rate at ' + br + '%.');

  if (osaPct    > 0)    parts.push('OSA and no-response are both at ' + or + '%.');
  if (cancelPct > 0.35) parts.push('High cancellation rate (' + cr + '%).');
  if (nrPct     > 0.40) parts.push('High no-response rate (' + nr + '%) — review follow-up cadence.');
  if (bookPct >= 0.50 && total < 7)
                        parts.push('Low volume (' + total + ' leads) — good rate but confirm trend with more data.');
  if (booked === 0 && total >= 5)
                        parts.push('Zero bookings — urgent review needed.');
  if (s.unqualified > 0) parts.push(s.unqualified + ' unqualified lead(s).');
  if (s.dupLead     > 0) parts.push(s.dupLead     + ' dup lead(s).');

  return { rating, notes: parts.join(' ') };
}

// ============================================================
// TRIGGER
// ============================================================

function setupWeeklyTrigger() {
  _deleteTriggersFor('refreshAll');
  ScriptApp.newTrigger('refreshAll')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(15)
    .create();
  SpreadsheetApp.getUi().alert(
    '✅ Weekly trigger set!\n\n' +
    'Weekly sheets refresh every Monday at 3 PM (script timezone).\n\n' +
    'Confirm timezone is America/New_York in Apps Script → Project Settings → Time zone.'
  );
}

function _deleteTriggersFor(fnName) {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === fnName) ScriptApp.deleteTrigger(t);
  });
}

// ============================================================
// CUSTOM MENU
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 VA Stats')
    .addItem('🔄 Refresh Weekly Sheets Now', 'refreshAll')
    .addSeparator()
    .addItem('📅 Setup Monday Auto-Run (run once)', 'setupWeeklyTrigger')
    .addToUi();
}
