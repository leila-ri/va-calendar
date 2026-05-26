// ============================================================
// VA Live Stats — Booking Rate + Follow-Up Adherence
// ============================================================
// INSTALL (separate GAS project — paste into the VA Team Scorecard):
//  1. Extensions > Apps Script → paste this file, Save
//  2. Run setupLiveTrigger() ONCE → auto-refreshes every minute
//  3. Run refreshLiveStats() now to populate immediately
//
// CREATES / UPDATES (in this spreadsheet):
//  "MTD Booking Rate"     — per-VA booking stats + score for the current month
//  "Follow-Up Adherence"  — per-VA cadence compliance, date-based
//
// FOLLOW-UP DATE LOGIC:
//  Working days = Monday–Saturday.  Sundays + US federal holidays skipped.
//  Slot K due = lead date + 1 working day
//  Slot L due = lead date + 2 working days  … through Slot O at +5.
//  A slot is only expected once its due date has passed (≤ today).
//
// FOLLOW-UP FOR RESOLVED / SPECIAL-STATUS LEADS:
//  Leads with these statuses were already resolved but may have had
//  follow-up work done before the VA stopped:
//    Spanish lead, Cancelled/Invalid, OSA, Unqualified, # Issue,
//    Dup Lead, Fake Lead, Satellite Quote, Client Handles
//  Rule: if the VA DID start following up (≥ 1 filled slot) and then
//  STOPPED (empty expected slots after the last filled slot), those
//  gaps count against adherence.  Leads with zero filled slots are
//  skipped — the VA categorised them on the spot with no follow-up due.
// ============================================================

// ── CONFIG ──────────────────────────────────────────────────
const SOURCE_SS_ID         = '1Kd0smsW26IJR1H5kJ0MP0kH8jPOTazo2_TbLEkgmbRs';
const SCORECARD_SS_ID      = '1z6mLBx2gfoqFpSyjgg4C_69oB3eOT3lBmRynKWIwybk';
const SCORECARD_SHEET      = 'VA Team Scorecard';
const SC_VA_ROW            = 4;
const SC_BOOKING_SCORE_ROW = 6;

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
  const mtdRange = getMTDRange_(today);
  const leads    = loadLeads_();

  writeBookingRateSheet_(leads, mtdRange, today);
  writeFollowUpSheet_(leads, mtdRange, today);

  Logger.log('Live stats updated — ' + today.toISOString());
}

// ============================================================
// BOOKING SCORE  (0 – 10)
// ============================================================
//  45%+  → 10   (at/above KPI)
//  40–44 →  9
//  35–39 →  8
//  30–34 →  7
//  25–29 →  6
//  20–24 →  5
//  15–19 →  4
//  10–14 →  3
//   5–9  →  2
//   1–4  →  1
//   0%   →  0
//  no data → '—'

function bookingScore_(rate) {
  if (rate === null || rate === undefined) return '—';
  const p = rate * 100;
  if (p >= 45) return 10;
  if (p >= 40) return 9;
  if (p >= 35) return 8;
  if (p >= 30) return 7;
  if (p >= 25) return 6;
  if (p >= 20) return 5;
  if (p >= 15) return 4;
  if (p >= 10) return 3;
  if (p >= 5)  return 2;
  if (p >  0)  return 1;
  return 0;
}

// ============================================================
// FOLLOW-UP SCORE  (0 – 10)
// ============================================================
//  90%+  → 10
//  85–89 →  9
//  80–84 →  8
//  75–79 →  7
//  70–74 →  6
//  65–69 →  5
//  60–64 →  4
//  55–59 →  3
//  50–54 →  2
//  40–49 →  1
//  < 40% →  0
//  no data → '—'

function followUpScore_(adherence) {
  if (adherence === null || adherence === undefined) return '—';
  const p = adherence * 100;
  if (p >= 90) return 10;
  if (p >= 85) return 9;
  if (p >= 80) return 8;
  if (p >= 75) return 7;
  if (p >= 70) return 6;
  if (p >= 65) return 5;
  if (p >= 60) return 4;
  if (p >= 55) return 3;
  if (p >= 50) return 2;
  if (p >= 40) return 1;
  return 0;
}

// ============================================================
// SHEET 1: MTD BOOKING RATE
// ============================================================

function writeBookingRateSheet_(leads, mtdRange, today) {
  const destSS    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet     = resetSheet_(destSS, 'MTD Booking Rate');
  const monthLabel= fmtMonth_(today);
  const lastRun   = fmtTs_(today);

  const EXCL_DISP = new Set(['osa','unqualified','dup lead','# issue','#issue']);

  // ── Aggregate per VA ────────────────────────────────────
  const vaMap = {};
  leads
    .filter(l => inRange_(l.dateIn, mtdRange) || inRange_(l.dateConf, mtdRange))
    .forEach(l => {
      if (!vaMap[l.va]) vaMap[l.va] = {
        total:0, booked:0, qualified:0,
        cancelled:0, exclDisp:0, exclCH:0
      };
      const m = vaMap[l.va];
      m.total++;
      if (l.bucket === 'booked') m.booked++;

      const isSpecial   = EXCLUDE_CLIENT_HANDLES_ACCTS.includes(l.subaccount);
      const isCancelled = l.bucket === 'cancelled';
      const isExclDisp  = EXCL_DISP.has(l.disp);
      const isExclCH    = isSpecial && l.bucket === 'clientHandles';

      if (isCancelled) m.cancelled++;
      else if (isExclDisp) m.exclDisp++;
      else if (isExclCH)   m.exclCH++;

      if (!isCancelled && !isExclDisp && !isExclCH) m.qualified++;
    });

  const COLS = [
    'VA','Total Leads','Qualified Leads','Booked',
    'Booking Rate','KPI (45%+)',
    'Cancelled / Invalid','Excluded Dispositions','Client Handles (excl.)'
  ];
  const W = COLS.length; // 9 cols (A–I); rows 1–2 merged to A:J

  const vaScores = {}; // va → score, used for scorecard write + row coloring

  const rows = Object.entries(vaMap)
    .sort((a,b) => {
      const ra = a[1].qualified ? a[1].booked/a[1].qualified : 0;
      const rb = b[1].qualified ? b[1].booked/b[1].qualified : 0;
      return rb - ra;
    })
    .map(([va, m]) => {
      const rate  = m.qualified > 0 ? m.booked/m.qualified : null;
      const score = bookingScore_(rate);
      const kpi   = rate === null ? '—' : rate >= BOOKING_KPI ? '✓ Met' : '✗ Below';
      vaScores[va] = score;
      return [
        va, m.total, m.qualified, m.booked,
        rate !== null ? pct_(m.booked, m.qualified) : 'N/A',
        kpi,
        m.cancelled, m.exclDisp, m.exclCH
      ];
    });

  // Grand totals
  const gt = Array(W).fill(''); gt[0] = 'TOTAL';
  [1,2,3,6,7,8].forEach(i => { gt[i] = rows.reduce((s,r)=>s+(Number(r[i])||0),0); });
  gt[4] = gt[2] > 0 ? pct_(gt[3], gt[2]) : 'N/A';
  gt[5] = gt[2] > 0 ? (gt[3]/gt[2] >= BOOKING_KPI ? '✓ Met' : '✗ Below') : '—';

  const OUT = [
    ['MTD BOOKING RATE — ' + monthLabel + ' — ' + lastRun],
    [
      'Booking Rate = Booked ÷ Qualified Leads.  ' +
      'Excluded from denominator: Cancelled/Invalid · OSA · Unqualified · Dup Lead · # Issue · ' +
      'Client Handles (Outstanding Roofing + Good Guy Roofing only).  ' +
      'Score: 45%+=10 · 40%=9 · 35%=8 · 30%=7 · 25%=6 · 20%=5 · 15%=4 · 10%=3 · 5%=2 · 1%=1 · 0%=0'
    ],
    Array(W).fill(''),
    COLS,
    gt,
    ...rows
  ];
  writeGrid_(sheet, OUT, W);

  // Formatting
  styleRow_(sheet, 1, W, '#1F3864','#FFFFFF', true,  11);
  styleRow_(sheet, 2, W, '#1F3864','#CCDDFF', false,  9);
  styleRow_(sheet, 4, W, '#2F5496','#FFFFFF', true,  10);
  styleRow_(sheet, 5, W, '#F2F2F2','#000000', true,  10);

  // Merge title rows A:J before per-row coloring
  sheet.getRange(1, 1, 1, 10).merge();
  sheet.getRange(2, 1, 1, 10).merge();

  const ds = 6;
  for (let i = 0; i < rows.length; i++) {
    const score    = vaScores[rows[i][0]];  // score kept for color logic
    const kpi      = rows[i][5];            // KPI now at index 5
    const rateCell = sheet.getRange(ds+i, 5);
    const kpiCell  = sheet.getRange(ds+i, 6);

    if (kpi === '✓ Met') {
      rateCell.setBackground('#C6EFCE').setFontColor('#276221');
      kpiCell.setBackground('#C6EFCE').setFontColor('#276221').setFontWeight('bold');
    } else if (kpi === '✗ Below') {
      const bg = (typeof score === 'number' && score >= 6) ? '#FFEB9C' : '#FFC7CE';
      const fg = (typeof score === 'number' && score >= 6) ? '#9C5700' : '#9C0006';
      rateCell.setBackground(bg).setFontColor(fg);
      kpiCell.setBackground(bg).setFontColor(fg).setFontWeight('bold');
    }
  }

  sheet.setFrozenRows(4);
  sheet.autoResizeColumns(1, W);
  sheet.setColumnWidth(1, 80);
  sheet.getRange(1, 1, 1, 1).setWrap(true);
  sheet.getRange(2, 1, 1, 1).setWrap(true);

  // Write booking scores to VA Team Scorecard row 6
  writeBookingScoresToScorecard_(vaScores);

  Logger.log('MTD Booking Rate sheet: ' + rows.length + ' VAs');
}

// ============================================================
// SHEET 2: FOLLOW-UP ADHERENCE  (date-based, stop detection)
// ============================================================

// Returns true if this lead has one of the "resolved/special" statuses.
function isSpecialStatusLead_(l) {
  const st = l.status.toLowerCase();
  if (st.includes('spanish'))          return true;
  if (st.includes('cancelled'))        return true;
  if (st.includes('invalid'))          return true;
  if (st === 'satellite quote')        return true;
  if (st === 'client handles')         return true;

  const d = l.disp;
  if (d === 'osa')                     return true;
  if (d === 'unqualified')             return true;
  if (d === '# issue' || d === '#issue') return true;
  if (d === 'dup lead')                return true;
  if (d === 'fake lead' || d === 'fake') return true;
  if (d === 'not interested')          return true;

  return false;
}

// 0-based index of the last filled follow-up slot, or -1 if none filled.
function lastFilledIdx_(lead) {
  let last = -1;
  for (let k = 0; k < 5; k++) if (lead.followUps[k]) last = k;
  return last;
}

function writeFollowUpSheet_(leads, mtdRange, today) {
  const destSS    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet     = resetSheet_(destSS, 'Follow-Up Adherence');
  const monthLabel= fmtMonth_(today);
  const lastRun   = fmtTs_(today);
  const todayMid  = midnight_(today);

  const vaMap     = {};
  const missedRows= [];

  leads.forEach(l => {
    if (l.bucket === 'booked') return;  // booked = no more follow-up needed
    if (!inRange_(l.dateIn, mtdRange) && !inRange_(l.dateConf, mtdRange)) return;
    if (!l.dateIn) return;

    const special   = isSpecialStatusLead_(l);
    const lastFilled= lastFilledIdx_(l);

    if (special) {
      // Only penalise if follow-up was started and then stopped.
      // Leads with 0 filled slots were categorised on the spot — skip.
      if (lastFilled < 0) return;

      // For these leads, only count expected slots AFTER the last filled one.
      // Any empty slot after that point = the VA stopped = a miss.
      if (!vaMap[l.va]) vaMap[l.va] = { count:0, expected:0, filled:0 };
      vaMap[l.va].count++;

      for (let slot = lastFilled + 2; slot <= 5; slot++) {   // slots are 1-indexed
        const dueDate = addWorkingDays_(l.dateIn, slot);
        if (midnight_(dueDate) > todayMid) break;            // not yet due

        vaMap[l.va].expected++;
        if (l.followUps[slot - 1]) {
          vaMap[l.va].filled++;
        } else {
          missedRows.push({
            va:       l.va,
            sub:      l.subaccount,
            name:     l.leadName,
            leadDate: fmtShort_(l.dateIn),
            slot:     'Day ' + slot,
            dueDate:  fmtShort_(dueDate),
            status:   l.status,
            note:     'Stopped after Day ' + (lastFilled + 1)
          });
        }
      }
    } else {
      // Standard active lead — every expected slot must be filled.
      if (!vaMap[l.va]) vaMap[l.va] = { count:0, expected:0, filled:0 };
      vaMap[l.va].count++;

      for (let slot = 1; slot <= 5; slot++) {
        const dueDate = addWorkingDays_(l.dateIn, slot);
        if (midnight_(dueDate) > todayMid) break;

        vaMap[l.va].expected++;
        if (l.followUps[slot - 1]) {
          vaMap[l.va].filled++;
        } else {
          missedRows.push({
            va:       l.va,
            sub:      l.subaccount,
            name:     l.leadName,
            leadDate: fmtShort_(l.dateIn),
            slot:     'Day ' + slot,
            dueDate:  fmtShort_(dueDate),
            status:   l.status,
            note:     ''
          });
        }
      }
    }
  });

  // ── Summary rows ────────────────────────────────────────
  const SCOLS = ['VA','Active Leads','Expected Slots','Filled Slots',
                 'Adherence %','Score (0–10)','Missed Slots'];
  const SW = SCOLS.length;

  const sumRows = Object.entries(vaMap)
    .sort((a,b) => {
      const fa = a[1].expected ? a[1].filled/a[1].expected : 1;
      const fb = b[1].expected ? b[1].filled/b[1].expected : 1;
      return fb - fa;
    })
    .map(([va, m]) => {
      const adh   = m.expected ? m.filled/m.expected : 1;
      const score = followUpScore_(adh);
      const missed= m.expected - m.filled;
      return [va, m.count, m.expected, m.filled, pct_(m.filled, m.expected || 1), score, missed];
    });

  // ── Detail rows ─────────────────────────────────────────
  const DCOLS = ['VA','Sub-Account','Lead Name','Lead Date',
                 'Slot','Due Date','Lead Status','Note'];
  const DW    = DCOLS.length;
  const W     = Math.max(SW, DW);

  const totalMissed = sumRows.reduce((s,r)=>s+(Number(r[6])||0),0);

  const OUT = [];
  OUT.push(['FOLLOW-UP ADHERENCE — ' + monthLabel + ' — ' + lastRun]);
  OUT.push([
    'Working days = Mon–Sat (Sundays + US holidays skipped).  ' +
    'Standard active leads: every expected slot must be filled.  ' +
    'Resolved/special-status leads (Spanish, Cancelled, OSA, Unqualified, # Issue, Dup Lead, ' +
    'Fake Lead, Satellite Quote, Client Handles): penalised only if follow-up started then stopped ' +
    '(slots after last filled entry count as missed).'
  ]);
  OUT.push(Array(W).fill(''));
  OUT.push(SCOLS.concat(Array(W-SW).fill('')));
  sumRows.forEach(r => OUT.push(r.concat(Array(W-SW).fill(''))));
  OUT.push(Array(W).fill(''));
  OUT.push(['MISSED FOLLOW-UPS — ' + totalMissed + ' total'].concat(Array(W-1).fill('')));
  OUT.push(DCOLS.concat(Array(W-DW).fill('')));
  missedRows
    .sort((a,b) => a.va.localeCompare(b.va) || a.leadDate.localeCompare(b.leadDate) || a.slot.localeCompare(b.slot))
    .forEach(r => OUT.push(
      [r.va, r.sub, r.name, r.leadDate, r.slot, r.dueDate, r.status, r.note]
      .concat(Array(W-DW).fill(''))
    ));

  writeGrid_(sheet, OUT, W);

  // Formatting
  styleRow_(sheet, 1, W, '#1F3864','#FFFFFF', true,  11);
  styleRow_(sheet, 2, W, '#1F3864','#CCDDFF', false,  9);
  styleRow_(sheet, 4, W, '#2F5496','#FFFFFF', true,  10);

  const ds = 5;
  for (let i = 0; i < sumRows.length; i++) {
    const adh   = parseInt(sumRows[i][4]);   // adherence % string → number
    const score = sumRows[i][5];             // Score (0–10)
    const adhCell   = sheet.getRange(ds+i, 5);
    const scoreCell = sheet.getRange(ds+i, 6);

    let bg, fg;
    if (!isNaN(adh)) {
      if      (adh >= 90) { bg = '#C6EFCE'; fg = '#276221'; }
      else if (adh >= 70) { bg = '#FFEB9C'; fg = '#9C5700'; }
      else                { bg = '#FFC7CE'; fg = '#9C0006'; }
      adhCell.setBackground(bg).setFontColor(fg);
    }
    if (typeof score === 'number') {
      const sbg = score === 10 ? '#C6EFCE' : score >= 6 ? '#FFEB9C' : '#FFC7CE';
      const sfg = score === 10 ? '#276221' : score >= 6 ? '#9C5700' : '#9C0006';
      scoreCell.setBackground(sbg).setFontColor(sfg).setFontWeight('bold');
    }
  }

  const missHeaderRow = ds + sumRows.length + 1;
  const detailHdrRow  = missHeaderRow + 1;
  styleRow_(sheet, missHeaderRow, W, '#E8F0FE','#1F3864', true,  10);
  styleRow_(sheet, detailHdrRow,  W, '#4A86E8','#FFFFFF', true,  10);

  // Merge title rows A:J
  sheet.getRange(1, 1, 1, 10).merge();
  sheet.getRange(2, 1, 1, 10).merge();

  sheet.setFrozenRows(4);
  sheet.autoResizeColumns(1, W);
  sheet.setColumnWidth(1, 80);
  sheet.getRange(1, 1, 1, 1).setWrap(true);
  sheet.getRange(2, 1, 1, 1).setWrap(true);
  Logger.log('Follow-Up Adherence: ' + sumRows.length + ' VAs, ' + totalMissed + ' missed slots');
}

// ============================================================
// SCORECARD: BOOKING SCORE → ROW 6
// ============================================================

function writeBookingScoresToScorecard_(vaScores) {
  try {
    const ss    = SpreadsheetApp.openById(SCORECARD_SS_ID);
    const sheet = ss.getSheetByName(SCORECARD_SHEET);
    if (!sheet) { Logger.log('Scorecard sheet not found'); return; }
    const lastCol = sheet.getLastColumn();
    if (lastCol < 1) return;
    const vaRow = sheet.getRange(SC_VA_ROW, 1, 1, lastCol).getValues()[0];
    vaRow.forEach((cell, i) => {
      const vaName = String(cell).trim();
      if (!vaName) return;
      const score = vaScores[vaName];
      if (typeof score === 'number') {
        sheet.getRange(SC_BOOKING_SCORE_ROW, i + 1).setValue(score);
      }
    });
    Logger.log('Booking scores written to scorecard row ' + SC_BOOKING_SCORE_ROW);
  } catch(e) {
    Logger.log('writeBookingScoresToScorecard_ error: ' + e.message);
  }
}

// ============================================================
// WORKING-DAY HELPERS
// ============================================================

const holidayCache_ = {};

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
  if (date.getDay() === 0) return true;  // Sunday
  return isUSHoliday_(date);
}

function isUSHoliday_(date) {
  const y = date.getFullYear();
  if (!holidayCache_[y]) holidayCache_[y] = buildHolidaySet_(y);
  return holidayCache_[y].has(y + '-' + (date.getMonth()+1) + '-' + date.getDate());
}

function buildHolidaySet_(year) {
  const keys = new Set();
  const add  = d => keys.add(d.getFullYear()+'-'+(d.getMonth()+1)+'-'+d.getDate());

  // Fixed-date federal holidays
  add(new Date(year,  0,  1));  // New Year's Day
  add(new Date(year,  5, 19));  // Juneteenth
  add(new Date(year,  6,  4));  // Independence Day
  add(new Date(year, 10, 11));  // Veterans Day
  add(new Date(year, 11, 25));  // Christmas

  // Floating federal holidays
  add(nthWeekday_(year,  0, 1, 3));  // MLK Day:        3rd Mon of Jan
  add(nthWeekday_(year,  1, 1, 3));  // Presidents Day: 3rd Mon of Feb
  add(lastWeekday_(year, 4, 1));     // Memorial Day:   last Mon of May
  add(nthWeekday_(year,  8, 1, 1));  // Labor Day:      1st Mon of Sep
  add(nthWeekday_(year,  9, 1, 2));  // Columbus Day:   2nd Mon of Oct
  add(nthWeekday_(year, 10, 4, 4));  // Thanksgiving:   4th Thu of Nov

  return keys;
}

function nthWeekday_(year, month0, dow, n) {
  const d = new Date(year, month0, 1);
  while (d.getDay() !== dow) d.setDate(d.getDate() + 1);
  d.setDate(d.getDate() + (n-1) * 7);
  return d;
}

function lastWeekday_(year, month0, dow) {
  const d = new Date(year, month0+1, 0);
  while (d.getDay() !== dow) d.setDate(d.getDate() - 1);
  return d;
}

function midnight_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

// ============================================================
// DATE / FORMAT HELPERS
// ============================================================

function getMTDRange_(today) {
  return {
    start: new Date(today.getFullYear(), today.getMonth(), 1, 0, 0, 0, 0),
    end:   new Date(today.getFullYear(), today.getMonth(), today.getDate(), 23, 59, 59, 999)
  };
}

function parseDate_(v) {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  const d = new Date(v); return isNaN(d.getTime()) ? null : d;
}

function inRange_(d, r)  { return !!d && d >= r.start && d <= r.end; }
function pct_(n, d)      { return d ? Math.round(n/d*100)+'%' : '0%'; }

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

function loadLeads_() {
  const ss    = SpreadsheetApp.openById(SOURCE_SS_ID);
  const leads = [];

  MAIN_SHEETS.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) { Logger.log('WARNING: sheet not found — "' + name + '"'); return; }
    const last = sheet.getLastRow();
    if (last < 2) return;

    sheet.getRange(2, 1, last-1, 15).getValues().forEach(row => {
      const va = String(row[COL_VA] || '').trim();
      if (!va || va.toLowerCase() === 'va') return;

      leads.push({
        va,
        bucket:     getStatusBucket_(String(row[COL_STATUS] || '')),
        status:     String(row[COL_STATUS]    || '').trim(),
        disp:       String(row[COL_DISP]      || '').trim().toLowerCase(),
        dateIn:     parseDate_(row[COL_DATE_IN]),
        dateConf:   parseDate_(row[COL_DATE_CONF]),
        subaccount: String(row[COL_SUBACCOUNT]|| '').trim(),
        leadName:   String(row[COL_LEAD_NAME] || '').trim(),
        followUps:  Array.from({length:5}, (_,k) => String(row[COL_FU1+k]||'').trim())
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
    '✅ Live trigger set!\n\nSheets update every minute.\n\nTo stop: menu → Stop Auto-Refresh.'
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
