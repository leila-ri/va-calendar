// ============================================================
// VA Weekly Stats Automation — Google Apps Script
// ============================================================
// INSTALL:
//  1. Extensions > Apps Script in the weekly report spreadsheet
//  2. Delete all existing code, paste this file, Save
//  3. Run setupWeeklyTrigger() ONCE  → Monday 3 PM auto-run for weekly sheets
//  4. Run setupRealtimeTrigger() ONCE → every-minute auto-run for MTD + Follow-Up
//  5. Run refreshAll() now to populate everything immediately
//
// CREATES / UPDATES:
//  "current week leads"   — raw lead rows for prev Sun–Sat (feeds formulas)
//  "Current Week Stats"   — per-account summary with rates, rating, notes
//  "MTD VA Stats"         — month-to-date per VA, booking rate KPI  [LIVE]
//  "Follow-Up Adherence"  — 5-day cadence compliance per VA          [LIVE]
//
// TRIGGERS:
//  refreshLiveSheets()  — every 1 minute (MTD + Follow-Up only)
//  refreshAll()         — every Monday 3 PM (all four sheets)
// ============================================================

// ── IDs & SHEET NAMES ───────────────────────────────────────
const SOURCE_SS_ID     = '1Kd0smsW26IJR1H5kJ0MP0kH8jPOTazo2_TbLEkgmbRs';
const MAIN_SHEETS      = ['ROOF, MAIN', 'HVAC, MAIN', 'GUTTER, Main', 'WINDOWS, MAIN'];
const SCORECARD_SS_ID  = '1z6mLBx2gfoqFpSyjgg4C_69oB3eOT3lBmRynKWIwybk';
const SCORECARD_SHEET  = 'VA Team Scorecard';
const SC_VA_ROW        = 4;   // row that holds VA names
const SC_BOOKING_ROW   = 7;   // row to write MTD booking rate
const SC_FFUP_ROW      = 8;   // row to write follow-up adherence %

// ── SOURCE COLUMN INDICES (0-based, A=0 … O=14) ─────────────
const COL_VA         = 0;   // A
const COL_DATE_IN    = 1;   // B — date lead came in
const COL_SUBACCOUNT = 2;   // C — sub-account  ← formulas reference col C
const COL_LEAD_NAME  = 3;   // D
const COL_STATUS     = 6;   // G — status       ← formulas reference col G
const COL_SCHED_DATE = 7;   // H
const COL_DISP       = 8;   // I — disposition  ← cancellation reason
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

const CANCEL_REASONS = ['not interested','unqualified','osa','fake lead','dup lead','# issue'];
const BOOKING_KPI    = 0.45;
const EXCLUDE_CLIENT_HANDLES_ACCTS = ['Outstanding Roofing', 'Good Guy Roofing'];

// ── Trigger handler-function name used to identify the live trigger ──
const LIVE_TRIGGER_FN    = 'refreshLiveSheets';
const WEEKLY_TRIGGER_FN  = 'refreshAll';

// ============================================================
// ENTRY POINTS
// ============================================================

// Full refresh — weekly report + live sheets (called by Monday trigger & menu)
function refreshAll() {
  const today    = new Date();
  const prevWeek = getPrevWeekRange(today);
  const mtdRange = getMTDRange(today);

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
  updateMTDStats(leads, mtdRange, today);
  updateFollowUpAdherence(leads, mtdRange, today);
  updateScorecard(leads, mtdRange, today);

  Logger.log('Done!');
  SpreadsheetApp.getUi().alert('Done! All sheets have been refreshed.');
}

// Live refresh — MTD VA Stats + Follow-Up Adherence only (called every minute)
function refreshLiveSheets() {
  const today    = new Date();
  const mtdRange = getMTDRange(today);

  const { leads } = loadAllLeads();
  updateMTDStats(leads, mtdRange, today);
  updateFollowUpAdherence(leads, mtdRange, today);
  updateScorecard(leads, mtdRange, today);

  Logger.log('Live refresh done — ' + today.toISOString());
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

function getMTDRange(today) {
  return {
    start: new Date(today.getFullYear(), today.getMonth(), 1, 0, 0, 0, 0),
    end:   new Date(today.getFullYear(), today.getMonth(), today.getDate(), 23, 59, 59, 999)
  };
}

function parseDate(v) {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  const d = new Date(v); return isNaN(d.getTime()) ? null : d;
}

function inRange(d, r) { return !!d && d >= r.start && d <= r.end; }
function pct(n, d)     { return d ? Math.round(n / d * 100) + '%' : '0%'; }
function pctN(n, d)    { return d ? n / d : 0; }

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

      const status       = String(row[COL_STATUS]    || '').trim();
      const disp         = String(row[COL_DISP]      || '').trim().toLowerCase();
      const dateIn       = parseDate(row[COL_DATE_IN]);
      const dateConf     = parseDate(row[COL_DATE_CONF]);
      const bucket       = getStatusBucket(status);

      leads.push({
        va,
        status: status.toLowerCase().trim(),
        disp,
        bucket,
        isCancelledInvalid: bucket === 'cancelled',
        dateIn,
        dateConf,
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
      const acc = String(r[0]             || '').trim();
      const va  = String(vaCol[i][0] || '').trim();
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
// SHEET B: "Current Week Stats"
// ============================================================

const CW_W          = 21;
const CW_VA         = 1;
const CW_SUBACCT    = 2;
const CW_BOOKED     = 3;
const CW_CANCELLED  = 4;
const CW_NOTRESP    = 5;
const CW_CLIENT     = 6;
const CW_RESCHED    = 7;
const CW_TOTAL      = 8;
const CW_BOOK_RATE  = 9;
const CW_CANCEL_RATE= 10;
const CW_NR_RATE    = 11;
const CW_CLIENT_RATE= 12;
const CW_SUBACCT2   = 13;
const CW_OSA        = 14;
const CW_FAKE       = 15;
const CW_NOTINT     = 16;
const CW_UNQUAL     = 17;
const CW_NUMISSUE   = 18;
const CW_TOTAL2     = 19;
const CW_RATING     = 20;
const CW_NOTES      = 21;

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

    const noResp       = (s.notResponding || 0) + (s.reschedule || 0);
    const total        = s.total || 0;
    const isSpecialAcct= EXCLUDE_CLIENT_HANDLES_ACCTS.includes(key);
    const bookingDenom = isSpecialAcct ? Math.max(0, total - (s.clientHandles || 0)) : total;
    const { rating, notes } = scoreAccount(s, noResp, bookingDenom);

    return [
      s.va, key,
      s.booked      || 0,
      s.cancelled   || 0,
      noResp,
      s.clientHandles||0,
      s.reschedule  || 0,
      total,
      pct(s.booked       ||0, bookingDenom),
      pct(s.cancelled    ||0, total),
      pct(noResp,            total),
      pct(s.clientHandles||0, total),
      key,
      pct(s.osa          ||0, total),
      pct(s.fakeLead     ||0, total),
      pct(s.notInterested||0, total),
      pct(s.unqualified  ||0, total),
      pct(s.numIssue     ||0, total),
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
  sheet.setColumnWidth(CW_VA,      70);
  sheet.setColumnWidth(CW_SUBACCT, 160);
  sheet.setColumnWidth(CW_SUBACCT2,160);
  sheet.setColumnWidth(CW_NOTES,   350);
  [3,4,5,6,7,8,9,10,11,12,14,15,16,17,18,19,20].forEach(c => sheet.setColumnWidth(c, 72));
  sheet.setRowHeight(4, 50);
  sheet.setFrozenRows(4);
  sheet.setFrozenColumns(2);

  Logger.log('current week stats: ' + dataRows.length + ' accounts');
}

function scoreAccount(s, noResp, bookingDenom) {
  const total  = s.total  || 0;
  const booked = s.booked || 0;
  const cancel = s.cancelled || 0;
  const denom  = (bookingDenom != null ? bookingDenom : total) || 0;

  if (total === 0) return { rating: 'N/A', notes: 'No lead came in.' };

  const bookPct   = denom ? booked / denom  : 0;
  const cancelPct = total ? cancel / total  : 0;
  const nrPct     = total ? noResp / total  : 0;
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
// SHEET C: MTD VA STATS  [LIVE]
// ============================================================

function updateMTDStats(allLeads, mtdRange, today) {
  const destSS    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet     = resetSheet(destSS, 'MTD VA Stats');
  const monthLabel= Utilities.formatDate(today, Session.getScriptTimeZone(), 'MMMM yyyy');

  const mtdLeads = allLeads.filter(l =>
    inRange(l.dateIn, mtdRange) || inRange(l.dateConf, mtdRange)
  );

  const vaMap = {};
  mtdLeads.forEach(l => {
    if (!vaMap[l.va]) vaMap[l.va] = {
      total:0, booked:0, cancelled:0, notResponding:0,
      clientHandles:0, reschedule:0, other:0,
      osa:0, unqualified:0, dupLead:0, numIssue:0, fake:0, notInt:0
    };
    const s = vaMap[l.va];
    s.total++;
    s[l.bucket] = (s[l.bucket] || 0) + 1;
    if (l.isCancelledInvalid) {
      const d = l.disp;
      if      (d === 'not interested')            s.notInt++;
      else if (d === 'unqualified')               s.unqualified++;
      else if (d === 'osa')                       s.osa++;
      else if (d === 'fake lead' || d === 'fake') s.fake++;
      else if (d === 'dup lead')                  s.dupLead++;
      else if (d === '# issue' || d === '#issue') s.numIssue++;
    }
  });

  const EXCL_DISP = new Set(['osa','unqualified','dup lead','# issue','#issue']);
  mtdLeads.forEach(l => {
    if (!vaMap[l.va]) return;
    const isSpecialAcct = EXCLUDE_CLIENT_HANDLES_ACCTS.includes(l.subaccount);
    const excluded = (l.bucket === 'cancelled') ||
                     EXCL_DISP.has(l.disp) ||
                     (isSpecialAcct && l.bucket === 'clientHandles');
    if (!vaMap[l.va]._qual) vaMap[l.va]._qual = 0;
    if (!excluded) vaMap[l.va]._qual++;
    if (isSpecialAcct && l.bucket === 'clientHandles') vaMap[l.va].total--;
  });

  const COLS = ['VA','Total Leads','Qualified Leads','Confirmed/Manual Booked',
                'Booking Rate','KPI (45%+)',
                'Cancelled/Invalid','OSA','Unqualified','Dup Lead','# Issue',
                'Fake Lead','Not Int.','Not Responding','Client Handles','Reschedule'];
  const W = COLS.length;

  const rows = Object.entries(vaMap)
    .sort((a,b) => {
      const ra = a[1]._qual ? a[1].booked/a[1]._qual : 0;
      const rb = b[1]._qual ? b[1].booked/b[1]._qual : 0;
      return rb - ra;
    })
    .map(([va, s]) => {
      const qual = s._qual || 0;
      const rate = qual ? s.booked/qual : null;
      const kpi  = rate === null ? '—' : rate >= BOOKING_KPI ? '✓ Met' : '✗ Below';
      return [
        va, s.total, qual, s.booked,
        qual ? pct(s.booked, qual) : 'N/A', kpi,
        s.cancelled||0, s.osa||0, s.unqualified||0, s.dupLead||0, s.numIssue||0,
        s.fake||0, s.notInt||0,
        (s.notResponding||0)+(s.reschedule||0), s.clientHandles||0, s.reschedule||0
      ];
    });

  const gt = Array(W).fill(0); gt[0] = 'TOTAL';
  rows.forEach(r => [1,2,3,6,7,8,9,10,11,12,13,14,15].forEach(i => gt[i] += Number(r[i])||0));
  gt[4] = gt[2] ? pct(gt[3], gt[2]) : 'N/A';
  gt[5] = gt[2] ? (gt[3]/gt[2] >= BOOKING_KPI ? '✓ Met' : '✗ Below') : '—';

  const lastRun = 'Last updated: ' + Utilities.formatDate(today, Session.getScriptTimeZone(), 'MMM d, yyyy h:mm a');
  const OUT = [];
  OUT.push(['MTD VA STATS — ' + monthLabel + ' — ' + lastRun]);
  OUT.push(['Booking Rate = Confirmed+Manual Booked ÷ Qualified Leads (excl. Cancelled/Invalid + OSA/Unqualified/Dup Lead/# Issue).']);
  OUT.push(['KPI: 45%+  |  L1→L2 threshold: 25%+ consistently  |  Rating highlighted green/red']);
  OUT.push(Array(W).fill(''));
  OUT.push(COLS);
  OUT.push(gt);
  rows.forEach(r => OUT.push(r));

  writeGrid(sheet, OUT, W);
  styleRow(sheet, 1, W, '#1F3864', '#FFFFFF', true, 11);
  styleRow(sheet, 2, W, '#1F3864', '#FFFFFF', false, 9);
  styleRow(sheet, 3, W, '#1F3864', '#FFFFFF', false, 9);
  styleRow(sheet, 5, W, '#2F5496', '#FFFFFF', true, 10);
  styleRow(sheet, 6, W, '#F2F2F2', '#000000', true, 10);

  const ds = 7;
  for (let i = 0; i < rows.length; i++) {
    const kpi   = rows[i][5];  // '✓ Met' / '✗ Below' / '—' — use computed value, not getValue()
    const kCell = sheet.getRange(ds+i, 6);
    const rCell = sheet.getRange(ds+i, 5);
    if (kpi === '✓ Met') {
      kCell.setBackground('#C6EFCE').setFontColor('#276221');
      rCell.setBackground('#C6EFCE').setFontColor('#276221');
    } else if (kpi === '✗ Below') {
      kCell.setBackground('#FFC7CE').setFontColor('#9C0006');
      rCell.setBackground('#FFC7CE').setFontColor('#9C0006');
    }
  }
  sheet.autoResizeColumns(1, W);
  sheet.setFrozenRows(5);
  Logger.log('MTD VA Stats: ' + Object.keys(vaMap).length + ' VAs');
}

// ============================================================
// SHEET D: FOLLOW-UP ADHERENCE  [LIVE]
// ============================================================

function updateFollowUpAdherence(allLeads, mtdRange, today) {
  const destSS    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet     = resetSheet(destSS, 'Follow-Up Adherence');
  const monthLabel= Utilities.formatDate(today, Session.getScriptTimeZone(), 'MMMM yyyy');

  const EXCL_DISP = new Set(['osa','unqualified','dup lead','# issue','#issue']);
  const activeLeads = allLeads.filter(l => {
    if (l.bucket === 'booked')    return false;
    if (l.bucket === 'cancelled') return false;
    if (EXCL_DISP.has(l.disp))   return false;
    return inRange(l.dateIn, mtdRange) || inRange(l.dateConf, mtdRange);
  });

  const vaFU = {};
  activeLeads.forEach(l => {
    if (!vaFU[l.va]) vaFU[l.va] = { count:0, exp:0, filled:[0,0,0,0,0], missed:[] };
    const s = vaFU[l.va];
    s.count++;
    const days   = l.dateIn ? Math.floor((today - l.dateIn)/86400000) : 0;
    const expect = Math.min(5, Math.max(1, days + 1));
    s.exp += expect;
    let actual = 0; const missDays = [];
    for (let k = 0; k < 5; k++) {
      if (l.followUps[k]) { actual++; s.filled[k]++; }
      else if (k < expect) missDays.push('Day '+(k+1));
    }
    if (actual < expect) s.missed.push({
      sub: l.subaccount, name: l.leadName,
      date: fmtShort(l.dateIn), status: l.status,
      expect, actual, missing: expect-actual, days: missDays.join(', ')
    });
  });

  const COLS = ['VA','Active Leads','Slots Expected',
                'Day 1','Day 2','Day 3','Day 4','Day 5',
                'Total Filled','Adherence %','Score /15 pts','Leads w/ Gaps'];
  const W = COLS.length;

  const sumRows = Object.entries(vaFU)
    .sort((a,b) => {
      const fa = a[1].exp ? a[1].filled.reduce((x,y)=>x+y,0)/a[1].exp : 1;
      const fb = b[1].exp ? b[1].filled.reduce((x,y)=>x+y,0)/b[1].exp : 1;
      return fb - fa;
    })
    .map(([va, s]) => {
      const tot   = s.filled.reduce((a,b)=>a+b,0);
      const adh   = s.exp ? Math.round(tot/s.exp*100) : 100;
      const score = (adh/100*15).toFixed(1);
      return [va, s.count, s.exp,
              s.filled[0],s.filled[1],s.filled[2],s.filled[3],s.filled[4],
              tot, adh+'%', score+'/15', s.missed.length];
    });

  const DET     = ['VA','Sub-account','Lead Name','Date','Status','Expected','Actual','Missing','Missing Days'];
  const detRows = [];
  Object.entries(vaFU).sort((a,b)=>a[0].localeCompare(b[0])).forEach(([va,s]) =>
    s.missed.forEach(m => detRows.push([va,m.sub,m.name,m.date,m.status,m.expect,m.actual,m.missing,m.days]))
  );

  const totalGaps = sumRows.reduce((a,r)=>a+(Number(r[11])||0),0);
  const lastRun   = 'Last updated: ' + Utilities.formatDate(today, Session.getScriptTimeZone(), 'MMM d, yyyy h:mm a');
  const OUT = [];
  OUT.push(['FOLLOW-UP ADHERENCE — ' + monthLabel + ' — ' + lastRun]);
  OUT.push(['Tracks 5-day cadence (Col K–O) for all active non-booked MTD leads. Score = 15% of monthly scorecard.']);
  OUT.push(['Standard: full 5-day cadence per lead. Deduction: −3 pts per lead dropped before Day 5.']);
  OUT.push(Array(W).fill(''));
  OUT.push(COLS);
  sumRows.forEach(r => OUT.push(r));
  OUT.push(Array(W).fill(''));
  OUT.push(['LEADS WITH MISSING FOLLOW-UPS — ' + totalGaps]);
  OUT.push(DET.concat(Array(W-DET.length).fill('')));
  detRows.forEach(r => OUT.push(r.concat(Array(W-r.length).fill(''))));

  writeGrid(sheet, OUT, W);
  styleRow(sheet, 1, W, '#1F3864', '#FFFFFF', true, 11);
  styleRow(sheet, 2, W, '#1F3864', '#FFFFFF', false, 9);
  styleRow(sheet, 3, W, '#1F3864', '#FFFFFF', false, 9);
  styleRow(sheet, 5, W, '#2F5496', '#FFFFFF', true, 10);

  const ds = 6;
  for (let i = 0; i < sumRows.length; i++) {
    const p    = parseInt(sumRows[i][9]);  // Adherence % at index 9 — use computed value, not getValue()
    const cell = sheet.getRange(ds+i, 10);
    if (!isNaN(p)) {
      if      (p >= 90) cell.setBackground('#C6EFCE').setFontColor('#276221');
      else if (p >= 70) cell.setBackground('#FFEB9C').setFontColor('#9C5700');
      else              cell.setBackground('#FFC7CE').setFontColor('#9C0006');
    }
  }
  sheet.autoResizeColumns(1, W);
  sheet.setFrozenRows(5);
  Logger.log('Follow-Up Adherence: ' + Object.keys(vaFU).length + ' VAs');
}

// ============================================================
// SCORECARD WRITE  — pushes MTD booking rate + follow-up adherence
// into the "VA Team Scorecard" sheet (separate spreadsheet).
// Row 4 = VA names; Row 7 = booking rate; Row 8 = follow-up %.
// Values written as decimals (0.52) so % cell formatting shows correctly.
// ============================================================

function updateScorecard(allLeads, mtdRange, today) {
  try {
    const ss    = SpreadsheetApp.openById(SCORECARD_SS_ID);
    const sheet = ss.getSheetByName(SCORECARD_SHEET);
    if (!sheet) { Logger.log('Scorecard sheet not found: ' + SCORECARD_SHEET); return; }

    // ── MTD leads ──────────────────────────────────────────
    const mtdLeads = allLeads.filter(l =>
      inRange(l.dateIn, mtdRange) || inRange(l.dateConf, mtdRange)
    );

    // ── Booking rate per VA (mirrors updateMTDStats logic) ─
    const EXCL_DISP = new Set(['osa','unqualified','dup lead','# issue','#issue']);
    const vaBook = {};
    mtdLeads.forEach(l => {
      if (!vaBook[l.va]) vaBook[l.va] = { booked: 0, qual: 0 };
      if (l.bucket === 'booked') vaBook[l.va].booked++;
      const isSpecial  = EXCLUDE_CLIENT_HANDLES_ACCTS.includes(l.subaccount);
      const excluded   = (l.bucket === 'cancelled') ||
                         EXCL_DISP.has(l.disp) ||
                         (isSpecial && l.bucket === 'clientHandles');
      if (!excluded) vaBook[l.va].qual++;
    });

    // ── Follow-up adherence per VA (mirrors updateFollowUpAdherence) ─
    const activeLeads = mtdLeads.filter(l =>
      l.bucket !== 'booked' && l.bucket !== 'cancelled' && !EXCL_DISP.has(l.disp)
    );
    const vaFU = {};
    activeLeads.forEach(l => {
      if (!vaFU[l.va]) vaFU[l.va] = { filled: 0, exp: 0 };
      const days   = l.dateIn ? Math.floor((today - l.dateIn) / 86400000) : 0;
      const expect = Math.min(5, Math.max(1, days + 1));
      vaFU[l.va].exp += expect;
      for (let k = 0; k < expect; k++) {
        if (l.followUps[k]) vaFU[l.va].filled++;
      }
    });

    // ── Read VA names from row 4 ───────────────────────────
    const lastCol = sheet.getLastColumn();
    if (lastCol < 1) return;
    const vaRow = sheet.getRange(SC_VA_ROW, 1, 1, lastCol).getValues()[0];

    // ── Write per-column ───────────────────────────────────
    vaRow.forEach((cell, i) => {
      const vaName = String(cell).trim();
      if (!vaName) return;
      const col = i + 1;

      const bk = vaBook[vaName];
      if (bk && bk.qual > 0) {
        sheet.getRange(SC_BOOKING_ROW, col)
          .setValue(bk.booked / bk.qual)
          .setNumberFormat('0%');
      }

      const fu = vaFU[vaName];
      if (fu && fu.exp > 0) {
        sheet.getRange(SC_FFUP_ROW, col)
          .setValue(fu.filled / fu.exp)
          .setNumberFormat('0%');
      }
    });

    Logger.log('Scorecard updated — ' + today.toISOString());
  } catch(e) {
    Logger.log('updateScorecard error: ' + e.message);
  }
}

// ============================================================
// TRIGGERS
// ============================================================

// Run once — sets up every-1-minute live refresh for MTD + Follow-Up only.
// Does NOT touch the Monday weekly trigger.
function setupRealtimeTrigger() {
  _deleteTriggersFor(LIVE_TRIGGER_FN);
  ScriptApp.newTrigger(LIVE_TRIGGER_FN)
    .timeBased()
    .everyMinutes(1)
    .create();
  SpreadsheetApp.getUi().alert(
    '✅ Live trigger set!\n\n' +
    '"MTD VA Stats" and "Follow-Up Adherence" will now refresh every minute automatically.\n\n' +
    'To stop it, use the menu → Stop Live Auto-Refresh.'
  );
}

// Removes the live trigger without touching the Monday weekly trigger.
function stopRealtimeTrigger() {
  const removed = _deleteTriggersFor(LIVE_TRIGGER_FN);
  SpreadsheetApp.getUi().alert(
    removed > 0
      ? '⏹ Live auto-refresh stopped. MTD and Follow-Up sheets will no longer update automatically.'
      : 'No live trigger was running.'
  );
}

// Run once — sets the Monday 3 PM weekly refresh.
// Does NOT touch the live trigger.
function setupWeeklyTrigger() {
  _deleteTriggersFor(WEEKLY_TRIGGER_FN);
  ScriptApp.newTrigger(WEEKLY_TRIGGER_FN)
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(15)
    .create();
  SpreadsheetApp.getUi().alert(
    '✅ Weekly trigger set!\n\n' +
    'All four sheets will auto-refresh every Monday at 3 PM (script timezone).\n\n' +
    'Tip: confirm the timezone is America/New_York in Apps Script → Project Settings → Time zone.'
  );
}

// Deletes all triggers whose handler matches fnName; returns count removed.
function _deleteTriggersFor(fnName) {
  let count = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === fnName) {
      ScriptApp.deleteTrigger(t);
      count++;
    }
  });
  return count;
}

// ============================================================
// CUSTOM MENU
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 VA Stats')
    // ── Manual refresh ──────────────────────────────────────
    .addItem('🔄 Refresh All Sheets Now',             'refreshAll')
    .addItem('⚡ Refresh Live Sheets Now (MTD + F/U)', 'refreshLiveSheets')
    .addSeparator()
    // ── Live (MTD + Follow-Up) trigger ──────────────────────
    .addItem('▶ Start Live Auto-Refresh (every 1 min)', 'setupRealtimeTrigger')
    .addItem('⏹ Stop Live Auto-Refresh',                'stopRealtimeTrigger')
    .addSeparator()
    // ── Weekly trigger ──────────────────────────────────────
    .addItem('📅 Setup Monday Weekly Auto-Run (run once)', 'setupWeeklyTrigger')
    .addToUi();
}
