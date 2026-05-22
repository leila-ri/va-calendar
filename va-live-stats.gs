// ============================================================
// VA Live Stats — Booking Rate + Follow-Up Adherence
// ============================================================
// INSTALL (separate GAS project — paste into the VA Team Scorecard):
//  1. Extensions > Apps Script → paste this file, Save
//  2. Run setupLiveTrigger() ONCE → auto-refreshes every minute
//  3. Run refreshLiveStats() now to populate immediately
//
// CREATES / UPDATES (in this spreadsheet):
//  "MTD Booking Rate"     — per-VA booking stats for the current month
//  "Follow-Up Adherence"  — per-VA cadence compliance, date-based
//
// ALSO WRITES into "VA Team Scorecard" tab (same spreadsheet):
//  Row 7  — MTD Booking Rate %
//  Row 8  — Follow-Up Adherence %
//  (VA names are read from Row 4 to find each VA's column)
//
// FOLLOW-UP DATE LOGIC:
//  Working days = Monday – Saturday. Sundays and US federal holidays
//  are skipped. Each slot (K–O) has a specific due date:
//    Slot K due: lead date + 1 working day
//    Slot L due: lead date + 2 working days
//    Slot M due: lead date + 3 working days
//    Slot N due: lead date + 4 working days
//    Slot O due: lead date + 5 working days
//  A slot is only expected if its due date ≤ today.
// ============================================================

// ── CONFIG ──────────────────────────────────────────────────
const SOURCE_SS_ID    = '1Kd0smsW26IJR1H5kJ0MP0kH8jPOTazo2_TbLEkgmbRs';
const SCORECARD_SS_ID = '1z6mLBx2gfoqFpSyjgg4C_69oB3eOT3lBmRynKWIwybk';
const SCORECARD_SHEET = 'VA Team Scorecard';
const SC_VA_ROW       = 4;
const SC_BOOKING_ROW  = 7;
const SC_FFUP_ROW     = 8;

const MAIN_SHEETS = ['ROOF, MAIN', 'HVAC, MAIN', 'GUTTER, Main', 'WINDOWS, MAIN'];
const BOOKING_KPI = 0.45;
const EXCLUDE_CLIENT_HANDLES_ACCTS = ['Outstanding Roofing', 'Good Guy Roofing'];

// Source column indices (0-based)
const COL_VA         = 0;   // A
const COL_DATE_IN    = 1;   // B — date lead came in
const COL_SUBACCOUNT = 2;   // C
const COL_LEAD_NAME  = 3;   // D
const COL_STATUS     = 6;   // G
const COL_DISP       = 8;   // I — disposition
const COL_DATE_CONF  = 9;   // J — date confirmed
const COL_FU1        = 10;  // K–O follow-up slots (5 total)

// ============================================================
// ENTRY POINT
// ============================================================

function refreshLiveStats() {
  const today    = new Date();
  const mtdRange = getMTDRange(today);
  const leads    = loadLeads();

  writeBookingRateSheet(leads, mtdRange, today);
  writeFollowUpSheet(leads, mtdRange, today);
  writeToScorecard(leads, mtdRange, today);

  Logger.log('Live stats updated — ' + today.toISOString());
}

// ============================================================
// SHEET 1: MTD BOOKING RATE
// ============================================================
//
// FORMULA:  Booking Rate = Booked ÷ Qualified Leads
//
// BOOKED:  status = "Confirmed" or "Manual Booked"
//
// QUALIFIED = all MTD leads minus:
//   • Cancelled / Invalid bucket
//   • Disposition = OSA, Unqualified, Dup Lead, # Issue
//   • Client Handles on Outstanding Roofing + Good Guy Roofing only
//
// KPI THRESHOLD: 45%+
//
// ============================================================

function writeBookingRateSheet(leads, mtdRange, today) {
  const destSS    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet     = resetSheet_(destSS, 'MTD Booking Rate');
  const monthLabel= fmtMonth_(today);
  const lastRun   = fmtTs_(today);

  const EXCL_DISP = new Set(['osa', 'unqualified', 'dup lead', '# issue', '#issue']);

  // ── Aggregate per VA ────────────────────────────────────
  const vaMap = {};
  leads
    .filter(l => inRange_(l.dateIn, mtdRange) || inRange_(l.dateConf, mtdRange))
    .forEach(l => {
      if (!vaMap[l.va]) vaMap[l.va] = {
        total: 0, booked: 0, qualified: 0,
        cancelled: 0, exclDisp: 0, exclCH: 0
      };
      const m = vaMap[l.va];
      m.total++;

      if (l.bucket === 'booked') m.booked++;

      const isSpecial  = EXCLUDE_CLIENT_HANDLES_ACCTS.includes(l.subaccount);
      const isCancelled= l.bucket === 'cancelled';
      const isExclDisp = EXCL_DISP.has(l.disp);
      const isExclCH   = isSpecial && l.bucket === 'clientHandles';

      if (isCancelled) m.cancelled++;
      else if (isExclDisp) m.exclDisp++;
      else if (isExclCH)   m.exclCH++;

      if (!isCancelled && !isExclDisp && !isExclCH) m.qualified++;
    });

  const COLS = [
    'VA', 'Total Leads', 'Qualified Leads', 'Booked',
    'Booking Rate', 'KPI (45%+)',
    'Cancelled / Invalid', 'Excluded Dispositions',
    'Client Handles (excl.)'
  ];
  const W = COLS.length;

  const rows = Object.entries(vaMap)
    .sort((a, b) => {
      const ra = a[1].qualified ? a[1].booked / a[1].qualified : 0;
      const rb = b[1].qualified ? b[1].booked / b[1].qualified : 0;
      return rb - ra;
    })
    .map(([va, m]) => {
      const rate = m.qualified > 0 ? m.booked / m.qualified : null;
      const kpi  = rate === null ? '—' : rate >= BOOKING_KPI ? '✓ Met' : '✗ Below';
      return [
        va, m.total, m.qualified, m.booked,
        rate !== null ? pct_(m.booked, m.qualified) : 'N/A',
        kpi,
        m.cancelled, m.exclDisp, m.exclCH
      ];
    });

  // Grand totals
  const gt = Array(W).fill(''); gt[0] = 'TOTAL';
  [1,2,3,6,7,8].forEach(i => { gt[i] = rows.reduce((s,r) => s + (Number(r[i])||0), 0); });
  gt[4] = gt[2] > 0 ? pct_(gt[3], gt[2]) : 'N/A';
  gt[5] = gt[2] > 0 ? (gt[3]/gt[2] >= BOOKING_KPI ? '✓ Met' : '✗ Below') : '—';

  const OUT = [
    ['MTD BOOKING RATE — ' + monthLabel + ' — ' + lastRun],
    ['Booking Rate = Booked ÷ Qualified Leads.  Excluded: Cancelled/Invalid, OSA, Unqualified, Dup Lead, # Issue,  Client Handles (Outstanding Roofing + Good Guy Roofing only).'],
    Array(W).fill(''),
    COLS,
    gt,
    ...rows
  ];
  writeGrid_(sheet, OUT, W);

  // Formatting
  styleRow_(sheet, 1, W, '#1F3864', '#FFFFFF', true, 11);
  styleRow_(sheet, 2, W, '#1F3864', '#CCDDFF', false, 9);
  styleRow_(sheet, 4, W, '#2F5496', '#FFFFFF', true, 10);
  styleRow_(sheet, 5, W, '#F2F2F2', '#000000', true, 10);

  const ds = 6;
  for (let i = 0; i < rows.length; i++) {
    const kpi  = rows[i][5];
    const rate = rows[i][4];
    const rateCell = sheet.getRange(ds + i, 5);
    const kpiCell  = sheet.getRange(ds + i, 6);
    if (kpi === '✓ Met') {
      rateCell.setBackground('#C6EFCE').setFontColor('#276221');
      kpiCell.setBackground('#C6EFCE').setFontColor('#276221').setFontWeight('bold');
    } else if (kpi === '✗ Below') {
      rateCell.setBackground('#FFC7CE').setFontColor('#9C0006');
      kpiCell.setBackground('#FFC7CE').setFontColor('#9C0006').setFontWeight('bold');
    }
  }

  sheet.setFrozenRows(4);
  sheet.autoResizeColumns(1, W);
  Logger.log('MTD Booking Rate sheet: ' + rows.length + ' VAs');
}

// ============================================================
// SHEET 2: FOLLOW-UP ADHERENCE
// ============================================================
//
// WHICH LEADS ARE TRACKED:
//   Active MTD leads — not booked, not cancelled, not bad dispositions.
//
// DUE DATES (working days = Mon–Sat, skip Sundays + US holidays):
//   Slot K (Day 1) due: lead date + 1 working day
//   Slot L (Day 2) due: lead date + 2 working days
//   Slot M (Day 3) due: lead date + 3 working days
//   Slot N (Day 4) due: lead date + 4 working days
//   Slot O (Day 5) due: lead date + 5 working days
//
//   Example — lead date 5/18 (Mon), checking on 5/22 (Fri):
//     Day 1 due 5/19 ✓ expected    Day 2 due 5/20 ✓ expected
//     Day 3 due 5/21 ✓ expected    Day 4 due 5/22 ✓ expected (today)
//     Day 5 due 5/23  — not yet due
//   → 4 slots expected; any unfilled = a miss.
//
// ============================================================

function writeFollowUpSheet(leads, mtdRange, today) {
  const destSS    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet     = resetSheet_(destSS, 'Follow-Up Adherence');
  const monthLabel= fmtMonth_(today);
  const lastRun   = fmtTs_(today);
  const todayMid  = midnight_(today);

  const EXCL_DISP = new Set(['osa', 'unqualified', 'dup lead', '# issue', '#issue']);

  const activeLeads = leads.filter(l => {
    if (l.bucket === 'booked')    return false;
    if (l.bucket === 'cancelled') return false;
    if (EXCL_DISP.has(l.disp))   return false;
    return inRange_(l.dateIn, mtdRange) || inRange_(l.dateConf, mtdRange);
  });

  // ── Per-VA aggregation ──────────────────────────────────
  const vaMap = {};
  const missedRows = [];  // for the detail section

  activeLeads.forEach(l => {
    if (!l.dateIn) return;
    if (!vaMap[l.va]) vaMap[l.va] = { count: 0, expected: 0, filled: 0 };
    const m = vaMap[l.va];
    m.count++;

    for (let slot = 1; slot <= 5; slot++) {
      const dueDate = addWorkingDays_(l.dateIn, slot);
      if (midnight_(dueDate) > todayMid) break; // not yet due — and neither are later slots

      m.expected++;
      if (l.followUps[slot - 1]) {
        m.filled++;
      } else {
        // Record this specific miss
        missedRows.push({
          va:       l.va,
          sub:      l.subaccount,
          name:     l.leadName,
          leadDate: fmtShort_(l.dateIn),
          slot:     'Day ' + slot,
          dueDate:  fmtShort_(dueDate),
          status:   l.status
        });
      }
    }
  });

  // ── Summary rows ────────────────────────────────────────
  const SCOLS = [
    'VA', 'Active Leads', 'Expected Slots', 'Filled Slots',
    'Adherence %', 'Score /15 pts', 'Missed Slots'
  ];
  const SW = SCOLS.length;

  const sumRows = Object.entries(vaMap)
    .sort((a, b) => {
      const fa = a[1].expected ? a[1].filled / a[1].expected : 1;
      const fb = b[1].expected ? b[1].filled / b[1].expected : 1;
      return fb - fa;
    })
    .map(([va, m]) => {
      const adh   = m.expected ? Math.round(m.filled / m.expected * 100) : 100;
      const score = (adh / 100 * 15).toFixed(1);
      const missed= m.expected - m.filled;
      return [va, m.count, m.expected, m.filled, adh + '%', score + '/15', missed];
    });

  // ── Detail rows ─────────────────────────────────────────
  const DCOLS = ['VA', 'Sub-Account', 'Lead Name', 'Lead Date', 'Slot', 'Due Date', 'Status'];
  const DW    = DCOLS.length;

  const totalMissed = sumRows.reduce((s, r) => s + (Number(r[6]) || 0), 0);

  // Use the wider of the two sections for writeGrid_ width
  const W = Math.max(SW, DW);

  const OUT = [];
  OUT.push(['FOLLOW-UP ADHERENCE — ' + monthLabel + ' — ' + lastRun]);
  OUT.push([
    'Working days = Mon–Sat. Sundays and US federal holidays are skipped. ' +
    'Each slot (K–O) is due one additional working day after the lead date. ' +
    'A slot is only counted as expected once its due date has passed.'
  ]);
  OUT.push(Array(W).fill(''));
  OUT.push(SCOLS.concat(Array(W - SW).fill('')));
  sumRows.forEach(r => OUT.push(r.concat(Array(W - SW).fill(''))));
  OUT.push(Array(W).fill(''));
  OUT.push(['MISSED FOLLOW-UPS — ' + totalMissed + ' total'].concat(Array(W - 1).fill('')));
  OUT.push(DCOLS.concat(Array(W - DW).fill('')));
  missedRows
    .sort((a, b) => a.va.localeCompare(b.va) || a.leadDate.localeCompare(b.leadDate))
    .forEach(r => OUT.push(
      [r.va, r.sub, r.name, r.leadDate, r.slot, r.dueDate, r.status]
      .concat(Array(W - DW).fill(''))
    ));

  writeGrid_(sheet, OUT, W);

  // Formatting
  styleRow_(sheet, 1, W, '#1F3864', '#FFFFFF', true, 11);
  styleRow_(sheet, 2, W, '#1F3864', '#CCDDFF', false, 9);
  styleRow_(sheet, 4, W, '#2F5496', '#FFFFFF', true, 10); // summary header

  const ds = 5;
  for (let i = 0; i < sumRows.length; i++) {
    const p    = parseInt(sumRows[i][4]);
    const cell = sheet.getRange(ds + i, 5);
    if (!isNaN(p)) {
      if      (p >= 90) cell.setBackground('#C6EFCE').setFontColor('#276221');
      else if (p >= 70) cell.setBackground('#FFEB9C').setFontColor('#9C5700');
      else              cell.setBackground('#FFC7CE').setFontColor('#9C0006');
    }
  }

  const missHeaderRow = 5 + sumRows.length + 2; // blank + "MISSED..." row
  const detailHdrRow  = missHeaderRow + 1;
  styleRow_(sheet, missHeaderRow, W, '#E8F0FE', '#1F3864', true, 10);
  styleRow_(sheet, detailHdrRow,  W, '#4A86E8', '#FFFFFF', true, 10);

  sheet.setFrozenRows(4);
  sheet.autoResizeColumns(1, W);
  Logger.log('Follow-Up Adherence sheet: ' + sumRows.length + ' VAs, ' + totalMissed + ' missed slots');
}

// ============================================================
// SCORECARD ROW WRITE  (rows 7 & 8 of "VA Team Scorecard")
// ============================================================

function writeToScorecard(leads, mtdRange, today) {
  try {
    const ss    = SpreadsheetApp.openById(SCORECARD_SS_ID);
    const sheet = ss.getSheetByName(SCORECARD_SHEET);
    if (!sheet) { Logger.log('Scorecard sheet not found: ' + SCORECARD_SHEET); return; }

    const lastCol = sheet.getLastColumn();
    if (lastCol < 1) return;

    const EXCL_DISP = new Set(['osa', 'unqualified', 'dup lead', '# issue', '#issue']);
    const todayMid  = midnight_(today);

    // Booking rate per VA
    const vaBook = {};
    leads
      .filter(l => inRange_(l.dateIn, mtdRange) || inRange_(l.dateConf, mtdRange))
      .forEach(l => {
        if (!vaBook[l.va]) vaBook[l.va] = { booked: 0, qualified: 0 };
        if (l.bucket === 'booked') vaBook[l.va].booked++;
        const isSpecial = EXCLUDE_CLIENT_HANDLES_ACCTS.includes(l.subaccount);
        const excluded  = (l.bucket === 'cancelled') || EXCL_DISP.has(l.disp) ||
                          (isSpecial && l.bucket === 'clientHandles');
        if (!excluded) vaBook[l.va].qualified++;
      });

    // Follow-up adherence per VA (date-based)
    const vaFU = {};
    leads
      .filter(l => {
        if (l.bucket === 'booked' || l.bucket === 'cancelled') return false;
        if (EXCL_DISP.has(l.disp)) return false;
        return inRange_(l.dateIn, mtdRange) || inRange_(l.dateConf, mtdRange);
      })
      .forEach(l => {
        if (!l.dateIn) return;
        if (!vaFU[l.va]) vaFU[l.va] = { filled: 0, expected: 0 };
        for (let slot = 1; slot <= 5; slot++) {
          if (midnight_(addWorkingDays_(l.dateIn, slot)) > todayMid) break;
          vaFU[l.va].expected++;
          if (l.followUps[slot - 1]) vaFU[l.va].filled++;
        }
      });

    // Write per VA column
    const vaRow = sheet.getRange(SC_VA_ROW, 1, 1, lastCol).getValues()[0];
    vaRow.forEach((cell, i) => {
      const vaName = String(cell).trim();
      if (!vaName) return;
      const col = i + 1;
      const bk  = vaBook[vaName];
      if (bk && bk.qualified > 0) {
        sheet.getRange(SC_BOOKING_ROW, col).setValue(bk.booked / bk.qualified).setNumberFormat('0%');
      }
      const fu = vaFU[vaName];
      if (fu && fu.expected > 0) {
        sheet.getRange(SC_FFUP_ROW, col).setValue(fu.filled / fu.expected).setNumberFormat('0%');
      }
    });

    Logger.log('Scorecard rows updated');
  } catch(e) {
    Logger.log('writeToScorecard error: ' + e.message);
  }
}

// ============================================================
// WORKING-DAY HELPERS
// ============================================================

// Cache holiday sets per year for the duration of a single execution.
const holidayCache_ = {};

// Returns the date that is n working days (Mon–Sat, excl. US holidays) after startDate.
function addWorkingDays_(startDate, n) {
  const d = new Date(startDate.getTime());
  let added = 0;
  while (added < n) {
    d.setDate(d.getDate() + 1);
    if (!isNonWorkingDay_(d)) added++;
  }
  return d;
}

function isNonWorkingDay_(date) {
  if (date.getDay() === 0) return true; // Sunday
  return isUSHoliday_(date);
}

function isUSHoliday_(date) {
  const y = date.getFullYear();
  if (!holidayCache_[y]) holidayCache_[y] = buildHolidaySet_(y);
  const key = y + '-' + (date.getMonth() + 1) + '-' + date.getDate();
  return holidayCache_[y].has(key);
}

function buildHolidaySet_(year) {
  const keys = new Set();
  const add  = d => keys.add(d.getFullYear() + '-' + (d.getMonth()+1) + '-' + d.getDate());

  // Fixed-date federal holidays
  add(new Date(year,  0,  1));  // New Year's Day
  add(new Date(year,  5, 19));  // Juneteenth
  add(new Date(year,  6,  4));  // Independence Day
  add(new Date(year, 10, 11));  // Veterans Day
  add(new Date(year, 11, 25));  // Christmas Day

  // Floating federal holidays
  add(nthWeekday_(year, 0, 1, 3));   // MLK Day:        3rd Mon of Jan
  add(nthWeekday_(year, 1, 1, 3));   // Presidents Day: 3rd Mon of Feb
  add(lastWeekday_(year, 4, 1));     // Memorial Day:   last Mon of May
  add(nthWeekday_(year, 8, 1, 1));   // Labor Day:      1st Mon of Sep
  add(nthWeekday_(year, 9, 1, 2));   // Columbus Day:   2nd Mon of Oct
  add(nthWeekday_(year, 10, 4, 4));  // Thanksgiving:   4th Thu of Nov

  return keys;
}

// Returns the nth occurrence of dayOfWeek (0=Sun…6=Sat) in month0-indexed month.
function nthWeekday_(year, month0, dayOfWeek, n) {
  const d = new Date(year, month0, 1);
  while (d.getDay() !== dayOfWeek) d.setDate(d.getDate() + 1);
  d.setDate(d.getDate() + (n - 1) * 7);
  return d;
}

// Returns the last occurrence of dayOfWeek in month0-indexed month.
function lastWeekday_(year, month0, dayOfWeek) {
  const d = new Date(year, month0 + 1, 0); // last day of month
  while (d.getDay() !== dayOfWeek) d.setDate(d.getDate() - 1);
  return d;
}

// Returns a Date at midnight (time zeroed) for date comparison.
function midnight_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

// ============================================================
// DATE HELPERS
// ============================================================

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

function inRange_(d, r) { return !!d && d >= r.start && d <= r.end; }
function pct_(n, d)     { return d ? Math.round(n / d * 100) + '%' : '0%'; }

function fmtShort_(d) {
  if (!d || isNaN(d.getTime())) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'M/d/yyyy');
}
function fmtMonth_(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'MMMM yyyy');
}
function fmtTs_(d) {
  return 'Last updated: ' + Utilities.formatDate(d, Session.getScriptTimeZone(), 'MMM d, yyyy h:mm a');
}

// ============================================================
// DATA LOADING
// ============================================================

function loadLeads() {
  const ss    = SpreadsheetApp.openById(SOURCE_SS_ID);
  const leads = [];

  MAIN_SHEETS.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) { Logger.log('WARNING: sheet not found — "' + name + '"'); return; }
    const last = sheet.getLastRow();
    if (last < 2) return;

    sheet.getRange(2, 1, last - 1, 15).getValues().forEach(row => {
      const va = String(row[COL_VA] || '').trim();
      if (!va || va.toLowerCase() === 'va') return;

      leads.push({
        va,
        bucket:     getStatusBucket_(String(row[COL_STATUS] || '')),
        disp:       String(row[COL_DISP]        || '').trim().toLowerCase(),
        status:     String(row[COL_STATUS]       || '').trim(),
        dateIn:     parseDate(row[COL_DATE_IN]),
        dateConf:   parseDate(row[COL_DATE_CONF]),
        subaccount: String(row[COL_SUBACCOUNT]   || '').trim(),
        leadName:   String(row[COL_LEAD_NAME]    || '').trim(),
        followUps:  Array.from({length: 5}, (_, k) => String(row[COL_FU1 + k] || '').trim())
      });
    });
  });

  return leads;
}

function getStatusBucket_(status) {
  const s = status.toLowerCase().trim();
  if (s === 'confirmed' || s === 'manual booked')        return 'booked';
  if (s.includes('cancelled') || s.includes('invalid')) return 'cancelled';
  if (['not booked','unconfirmed','not responding','call back'].includes(s))
                                                         return 'notResponding';
  if (s === 'client handles' || s === 'satellite quote') return 'clientHandles';
  if (s === 'reschedule needed')                         return 'reschedule';
  return 'other';
}

// ============================================================
// SHEET HELPERS
// ============================================================

function resetSheet_(ss, name) {
  const ex = ss.getSheetByName(name);
  if (ex) ss.deleteSheet(ex);
  return ss.insertSheet(name);
}

function writeGrid_(sheet, data, W) {
  if (!data.length) return;
  const padded = data.map(r => {
    const p = r.map(v => v == null ? '' : v);
    while (p.length < W) p.push('');
    return p.slice(0, W);
  });
  sheet.getRange(1, 1, padded.length, W).setValues(padded);
}

function styleRow_(sheet, row, W, bg, fg, bold, sz) {
  const r = sheet.getRange(row, 1, 1, W);
  if (bg)   r.setBackground(bg);
  if (fg)   r.setFontColor(fg);
  if (bold) r.setFontWeight('bold');
  if (sz)   r.setFontSize(sz);
}

// ============================================================
// TRIGGER
// ============================================================

function setupLiveTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'refreshLiveStats') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('refreshLiveStats').timeBased().everyMinutes(1).create();
  SpreadsheetApp.getUi().alert(
    '✅ Live trigger set!\n\nBoth sheets + scorecard rows update every minute.\n\nTo stop: menu → Stop Auto-Refresh.'
  );
}

function stopLiveTrigger() {
  let n = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'refreshLiveStats') { ScriptApp.deleteTrigger(t); n++; }
  });
  SpreadsheetApp.getUi().alert(n > 0 ? '⏹ Live trigger stopped.' : 'No live trigger was running.');
}

// ============================================================
// CUSTOM MENU
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 VA Live Stats')
    .addItem('⚡ Refresh Now',                    'refreshLiveStats')
    .addSeparator()
    .addItem('▶ Start Auto-Refresh (every 1 min)', 'setupLiveTrigger')
    .addItem('⏹ Stop Auto-Refresh',                'stopLiveTrigger')
    .addToUi();
}
