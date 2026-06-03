// ============================================================
// VA Live Stats — Booking Rate + Follow-Up Adherence
// ============================================================
// INSTALL (separate GAS project — paste into the VA Team Scorecard):
//  1. Extensions > Apps Script → paste this file, Save
//  2. Run setupLiveTrigger() ONCE → auto-refreshes every 5 minutes
//  3. Run refreshLiveStats() now to populate immediately
//
// CREATES / UPDATES (in this spreadsheet):
//  "MTD Booking Rate"     — per-VA booking stats + score for the current month
//  "Follow-Up Adherence"  — per-VA daily cadence compliance
//
// FOLLOW-UP LOGIC (day-based):
//  Every working day (Mon–Sat, no US holidays) from lead date must be
//  "covered" by: (a) a call made that day, OR (b) a callback window
//  parsed from the note text (tomorrow / next week / June 15 / etc.).
//  Un-covered working days = missed days.
//
//  Special-status leads (Cancelled, Client Handles, OSA, etc.):
//    Cutoff = last call date — stopping is correct, gaps before that are not.
//    Zero-entry leads = resolved on the spot, skipped entirely.
//
//  Active/Other leads: cutoff = today.
//
//  🚩 Late First Contact = first follow-up note is 2+ working days after lead date.
// ============================================================

// ── CONFIG ──────────────────────────────────────────────────
const SOURCE_SS_ID         = '1Kd0smsW26IJR1H5kJ0MP0kH8jPOTazo2_TbLEkgmbRs';
const SCORECARD_SS_ID      = '1z6mLBx2gfoqFpSyjgg4C_69oB3eOT3lBmRynKWIwybk';
const SCORECARD_SHEET      = 'VA Team Scorecard';
const SC_VA_ROW            = 4;
const SC_BOOKING_SCORE_ROW = 6;
const SC_START_COL         = 3;  // VA names (row 4) and scores start at column C

const MAIN_SHEETS = ['ROOF, MAIN', 'HVAC, MAIN', 'GUTTER, Main', 'WINDOWS, MAIN'];
const BOOKING_KPI = 0.45;
const EXCLUDE_CLIENT_HANDLES_ACCTS = ['Outstanding Roofing', 'Good Guy Roofing'];

// Off-day calendar + Sunday-coverage spreadsheet
const COVERAGE_SS_ID = '1RmLtprnhJxhY7asBbUSPuiahD4H8vbz_dIb42Y__wr0';
const CALENDAR_API   = 'https://script.google.com/macros/s/AKfycbwQdE3oeV_UWrUr4i5yYo9T_kzTRHKGjx2erFl7OI27d5La1BmUEoRImUE-INyWvVTuJg/exec';

// Flex VAs who started mid-period — only treated as flex on/after this date.
// Add a new entry here whenever a VA transitions from core → flex.
const FLEX_VA_EFFECTIVE = {
  'jessica': new Date(2026, 5, 5)   // June 5, 2026
};

// Source column indices (0-based)
const COL_VA         = 0;   // A — VA who handled the lead
const COL_DATE_IN    = 1;   // B — date lead came in
const COL_SUBACCOUNT = 2;   // C
const COL_LEAD_NAME  = 3;   // D
const COL_STATUS     = 6;   // G
const COL_DISP       = 8;   // I — disposition
const COL_DATE_CONF  = 9;   // J — date confirmed
const COL_FU1        = 10;  // K–O follow-up note slots (5 total, indices 10–14)

// ============================================================
// ENTRY POINT
// ============================================================

function refreshLiveStats() {
  const today    = new Date();
  const leads    = loadLeads_();
  const offDays  = loadOffDays_();
  const sunCov   = loadSundayCoverage_();
  const vaAsmt   = loadVAAssignments_();

  const currRange = getMTDRange_(today);
  const prevRange = getPrevMonthRange_(today);

  const currMo = fmtMonthShort_(currRange.start);  // e.g. "June"
  const prevMo = fmtMonthShort_(prevRange.start);  // e.g. "May"

  writeBookingRateSheet_(leads, currRange, today, 'MTD Booking Rate');
  writeFollowUpSheet_(leads, currRange, today, offDays, sunCov, vaAsmt, 'Follow-Up Adherence');

  // Previous month sheets are written once and then left alone — they are final
  // the moment the month ends, so we never overwrite them on subsequent runs.
  const destSS    = SpreadsheetApp.getActiveSpreadsheet();
  const prevBRTab = prevMo + ' Booking Rate';
  const prevFUTab = prevMo + ' Follow-Up Adherence';
  if (!destSS.getSheetByName(prevBRTab)) {
    writeBookingRateSheet_(leads, prevRange, today, prevBRTab);
  }
  if (!destSS.getSheetByName(prevFUTab)) {
    writeFollowUpSheet_(leads, prevRange, today, offDays, sunCov, vaAsmt, prevFUTab);
  }

  Logger.log('Live stats updated — ' + today.toISOString());
}

// ============================================================
// BOOKING SCORE  (0 – 10)
// ============================================================

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

function writeBookingRateSheet_(leads, mtdRange, today, tabName) {
  const destSS    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet     = resetSheet_(destSS, tabName || 'MTD Booking Rate');
  const monthLabel= fmtMonth_(mtdRange.start);
  const lastRun   = fmtTs_(today);

  const EXCL_DISP = new Set(['osa','unqualified','dup lead','# issue','#issue']);

  const vaMap = {};
  leads
    .filter(l => inRange_(l.dateIn, mtdRange))
    .forEach(l => {
      if (!vaMap[l.va]) vaMap[l.va] = {
        total:0, booked:0, qualified:0,
        cancelled:0, exclDisp:0, exclCH:0, exclConfOther:0
      };
      const m = vaMap[l.va];

      // Col I doubles as "Who confirmed?" for booked leads.
      // If confirmed by a different VA, exclude from this VA's totals entirely.
      const isConfByOther = l.bucket === 'booked' && !!l.disp &&
                            l.disp !== l.va.toLowerCase();
      if (isConfByOther) { m.exclConfOther++; return; }

      m.total++;
      // Booked = confirmed status AND the confirmation date (col J) falls in this month.
      // A lead confirmed in a different month is not counted as booked here.
      if (l.bucket === 'booked' && inRange_(l.dateConf, mtdRange)) m.booked++;

      const isSpecial   = EXCLUDE_CLIENT_HANDLES_ACCTS.includes(l.subaccount);
      const isCancelled = l.bucket === 'cancelled';
      const isExclDisp  = EXCL_DISP.has(l.disp) || l.status.toLowerCase().includes('spanish');
      const isExclCH    = isSpecial && l.bucket === 'clientHandles';

      if (isCancelled) m.cancelled++;
      else if (isExclDisp) m.exclDisp++;
      else if (isExclCH)   m.exclCH++;

      if (!isCancelled && !isExclDisp && !isExclCH) m.qualified++;
    });

  const COLS = [
    'VA','Total Leads','Qualified Leads','Booked',
    'Booking Rate','Score (0–10)','KPI (45%+)',
    'Cancelled / Invalid','Excluded Dispositions','Client Handles (excl.)',
    'Conf. by Other (excl.)'
  ];
  const W = COLS.length;

  const vaScores = {};

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
        score, kpi,
        m.cancelled, m.exclDisp, m.exclCH, m.exclConfOther
      ];
    });

  const gt = Array(W).fill(''); gt[0] = 'TOTAL';
  [1,2,3,7,8,9,10].forEach(i => { gt[i] = rows.reduce((s,r)=>s+(Number(r[i])||0),0); });
  gt[4] = gt[2] > 0 ? pct_(gt[3], gt[2]) : 'N/A';
  gt[5] = bookingScore_(gt[2] > 0 ? gt[3]/gt[2] : null);
  gt[6] = gt[2] > 0 ? (gt[3]/gt[2] >= BOOKING_KPI ? '✓ Met' : '✗ Below') : '—';

  const OUT = [
    ['MTD BOOKING RATE — ' + monthLabel + ' — ' + lastRun],
    [
      'Booking Rate = Booked ÷ Qualified Leads.  ' +
      'Excluded: Cancelled/Invalid · OSA · Unqualified · Dup Lead · # Issue · ' +
      'Client Handles (Outstanding Roofing + Good Guy Roofing only).  ' +
      'Score: 45%+=10 · 40%=9 · 35%=8 · 30%=7 · 25%=6 · 20%=5 · 15%=4 · 10%=3 · 5%=2 · 1%=1 · 0%=0'
    ],
    Array(W).fill(''),
    COLS,
    gt,
    ...rows
  ];
  writeGrid_(sheet, OUT, W);

  styleRow_(sheet, 1, W, '#1F3864','#FFFFFF', true,  11);
  styleRow_(sheet, 2, W, '#1F3864','#CCDDFF', false,  9);
  styleRow_(sheet, 4, W, '#2F5496','#FFFFFF', true,  10);
  styleRow_(sheet, 5, W, '#F2F2F2','#000000', true,  10);

  sheet.getRange(1, 1, 1, W).merge();
  sheet.getRange(2, 1, 1, W).merge();

  const ds = 6;
  for (let i = 0; i < rows.length; i++) {
    const score     = rows[i][5];
    const kpi       = rows[i][6];
    const rateCell  = sheet.getRange(ds+i, 5);
    const scoreCell = sheet.getRange(ds+i, 6);
    const kpiCell   = sheet.getRange(ds+i, 7);

    if (kpi === '✓ Met') {
      rateCell.setBackground('#C6EFCE').setFontColor('#276221');
      scoreCell.setBackground('#C6EFCE').setFontColor('#276221').setFontWeight('bold');
      kpiCell.setBackground('#C6EFCE').setFontColor('#276221').setFontWeight('bold');
    } else if (kpi === '✗ Below') {
      const bg = (typeof score === 'number' && score >= 6) ? '#FFEB9C' : '#FFC7CE';
      const fg = (typeof score === 'number' && score >= 6) ? '#9C5700' : '#9C0006';
      rateCell.setBackground(bg).setFontColor(fg);
      scoreCell.setBackground(bg).setFontColor(fg).setFontWeight('bold');
      kpiCell.setBackground(bg).setFontColor(fg).setFontWeight('bold');
    }
  }

  sheet.setFrozenRows(4);
  sheet.autoResizeColumns(1, W);
  sheet.setColumnWidth(1, 80);
  sheet.getRange(1, 1, 1, 1).setWrap(true);
  sheet.getRange(2, 1, 1, 1).setWrap(true);

  writeBookingScoresToScorecard_(vaScores);
  Logger.log('MTD Booking Rate sheet: ' + rows.length + ' VAs');
}

// ============================================================
// SHEET 2: FOLLOW-UP ADHERENCE  (day-based, callback-window aware)
// ============================================================

function isSpecialStatusLead_(l) {
  const st = l.status.toLowerCase();
  if (st.includes('spanish'))          return true;
  if (st.includes('cancelled'))        return true;
  if (st.includes('invalid'))          return true;
  if (st === 'satellite quote')        return true;
  if (st === 'client handles')         return true;
  if (st.includes('reschedule'))       return true;
  if (st.includes('call back') || st.includes('callback')) return true;

  const d = l.disp;
  if (d === 'osa')                     return true;
  if (d === 'unqualified')             return true;
  if (d === '# issue' || d === '#issue') return true;
  if (d === 'dup lead')                return true;
  if (d === 'fake lead' || d === 'fake') return true;

  return false;
}

// ── Follow-up note parsers ───────────────────────────────────

const MONTH_MAP_ = {
  jan:0, january:0, feb:1, february:1, mar:2, march:2,
  apr:3, april:3, may:4, jun:5, june:5, jul:6, july:6,
  aug:7, august:7, sep:8, sept:8, september:8,
  oct:9, october:9, nov:10, november:10, dec:11, december:11
};

function dateKey_(d) {
  return d.getFullYear() + '-' + (d.getMonth()+1) + '-' + d.getDate();
}

// Pull the leading MM/DD or M/D date from a note string
function parseLeadingDate_(text, leadDate) {
  const m = text.trim().match(/^(\d{1,2})[\/\-]{1,2}(\d{1,2})/);
  if (!m) return null;
  const mo = parseInt(m[1]) - 1, day = parseInt(m[2]);
  if (mo < 0 || mo > 11 || day < 1 || day > 31) return null;
  let yr = leadDate.getFullYear();
  let dt = new Date(yr, mo, day);
  // If parsed date is >30 days before lead date it likely rolled to next year
  if (midnight_(dt) < new Date(midnight_(leadDate).getTime() - 30*24*60*60*1000)) {
    dt = new Date(yr + 1, mo, day);
  }
  return dt;
}

// Find a specific calendar date referenced after a callback keyword
function parseSpecificCallbackDate_(text, callDate) {
  const t = text.toLowerCase();

  // "10th of June"
  let m = t.match(/(\d{1,2})(?:st|nd|rd|th)\s+of\s+([a-z]+)/);
  if (m) {
    const mn = MONTH_MAP_[m[2]];
    if (mn !== undefined) return bestFutureDate_(callDate, mn, parseInt(m[1]));
  }

  // "June 15" / "May 5"
  for (const [name, mn] of Object.entries(MONTH_MAP_)) {
    const re = new RegExp('\\b' + name + '\\s+(\\d{1,2})', 'i');
    m = t.match(re);
    if (m) return bestFutureDate_(callDate, mn, parseInt(m[1]));
  }

  // "the 30th" / "on the 30th"
  m = t.match(/\bthe\s+(\d{1,2})(?:st|nd|rd|th)/);
  if (m) {
    const day = parseInt(m[1]);
    let d = new Date(callDate.getFullYear(), callDate.getMonth(), day);
    if (midnight_(d) <= midnight_(callDate)) {
      d = new Date(callDate.getFullYear(), callDate.getMonth() + 1, day);
    }
    return d;
  }

  // Inline MM/DD not at the very start (that's the call date)
  for (const dm of text.matchAll(/(?<!\d)(\d{1,2})\/(\d{1,2})(?!\d)/g)) {
    if (dm.index < 5) continue;
    const mo = parseInt(dm[1]) - 1, day = parseInt(dm[2]);
    if (mo >= 0 && mo <= 11 && day >= 1 && day <= 31) {
      return bestFutureDate_(callDate, mo, day);
    }
  }

  return null;
}

function bestFutureDate_(afterDate, month0, day) {
  let yr = afterDate.getFullYear();
  let d = new Date(yr, month0, day);
  if (midnight_(d) <= midnight_(afterDate)) d = new Date(yr + 1, month0, day);
  return d;
}

// Returns {start, end} (both midnight Dates) or null if no callback detected
function parseCallbackWindow_(text, callDate) {
  if (!/callback|call\s*back|\bcb\b|will\s+call|requesting?\s+(?:a\s+)?call(?:back)?|lead\s+will\s+call|call(?:ing)?\s+(?:back|later|tomorrow)/i.test(text)) {
    return null;
  }

  const t  = text.toLowerCase();
  const cd = midnight_(callDate);

  // Same-day / "today" / "later today" / "at [time]"
  if (/\btoday\b/.test(t) || /later\s+(?:this\s+(?:morning|afternoon|evening)|today)/.test(t) || /\bat\s+\d/.test(t)) {
    return { start: cd, end: midnight_(addWorkingDays_(callDate, 1)) };
  }

  // "tomorrow"
  if (/\btomorrow\b/.test(t)) {
    const d = midnight_(addWorkingDays_(callDate, 1));
    return { start: d, end: d };
  }

  // "next week" / "nextweek"
  if (/next\s*(?:week|wk)/.test(t)) return nextWeekWindow_(callDate);

  // "within a week" / "within the week"
  if (/within\s+(?:a|the)\s+week/.test(t)) {
    return {
      start: midnight_(addWorkingDays_(callDate, 1)),
      end:   midnight_(addWorkingDays_(callDate, 5))
    };
  }

  // "1st week of [month]" / "first week of [month]"
  const fwMatch = t.match(/(?:1st|first)\s+week\s+of\s+([a-z]+)/);
  if (fwMatch) {
    const mn = MONTH_MAP_[fwMatch[1]];
    if (mn !== undefined) {
      let yr = callDate.getFullYear();
      let firstMon = new Date(yr, mn, 1);
      if (midnight_(firstMon) <= cd) firstMon = new Date(++yr, mn, 1);
      while (firstMon.getDay() !== 1) firstMon.setDate(firstMon.getDate() + 1);
      const lastDay = new Date(firstMon);
      lastDay.setDate(lastDay.getDate() + 5);
      return { start: midnight_(firstMon), end: midnight_(lastDay) };
    }
  }

  // "next [weekday]"
  const DOW = ['sunday','monday','tuesday','wednesday','thursday','friday','saturday'];
  const nextDow = t.match(/next\s+(monday|tuesday|wednesday|thursday|friday|saturday|sunday)/);
  if (nextDow) {
    const target = DOW.indexOf(nextDow[1]);
    const d = new Date(callDate);
    d.setDate(d.getDate() + 1);
    while (d.getDay() !== target) d.setDate(d.getDate() + 1);
    return { start: midnight_(d), end: midnight_(d) };
  }

  // "later" (generic — no specific day)
  if (/\blater\b/.test(t)) {
    return { start: cd, end: midnight_(addWorkingDays_(callDate, 1)) };
  }

  // Specific date (June 15, 10th of June, the 30th, inline MM/DD)
  const specific = parseSpecificCallbackDate_(text, callDate);
  if (specific) return { start: midnight_(specific), end: midnight_(specific) };

  // Callback signal present but no recognisable time → assume within 3 working days
  return {
    start: midnight_(addWorkingDays_(callDate, 1)),
    end:   midnight_(addWorkingDays_(callDate, 3))
  };
}

function nextWeekWindow_(fromDate) {
  const d = new Date(fromDate);
  d.setDate(d.getDate() + 1);
  while (d.getDay() !== 1) d.setDate(d.getDate() + 1); // land on Monday
  const e = new Date(d);
  e.setDate(e.getDate() + 5); // through Saturday
  return { start: midnight_(d), end: midnight_(e) };
}

// Parse all follow-up cells → [{callDate, cbWindow}]
function parseFollowUpEntries_(followUps, leadDate) {
  const entries = [];
  let prev = leadDate;
  for (let i = 0; i < followUps.length; i++) {
    const raw = followUps[i];
    if (!raw) continue;
    let callDate = parseLeadingDate_(raw, leadDate);
    if (!callDate) {
      // First actual entry with no date → assume called on lead date (never late)
      callDate = entries.length === 0 ? leadDate : addWorkingDays_(prev, 1);
    }
    callDate = midnight_(callDate);
    entries.push({ callDate, cbWindow: parseCallbackWindow_(raw, callDate) });
    prev = callDate;
  }
  return entries;
}

// Map of dateKey → 'call' | 'window' for every covered day
function buildCoverageMap_(entries) {
  const map = new Map();
  for (const { callDate, cbWindow } of entries) {
    map.set(dateKey_(callDate), 'call');
    if (cbWindow) {
      const d = new Date(cbWindow.start);
      while (d <= cbWindow.end) {
        const k = dateKey_(d);
        if (!map.has(k)) map.set(k, 'window');
        d.setDate(d.getDate() + 1);
      }
    }
  }
  return map;
}

// ── Main sheet writer ────────────────────────────────────────

function writeFollowUpSheet_(leads, mtdRange, today, offDays, sunCov, vaAssignments, tabName) {
  offDays       = offDays       || new Map();
  sunCov        = sunCov        || new Map();
  vaAssignments = vaAssignments || new Map();

  const destSS    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet     = resetSheet_(destSS, tabName || 'Follow-Up Adherence');
  const monthLabel= fmtMonth_(mtdRange.start);
  const lastRun   = fmtTs_(today);
  // For past months, cap at end of range (not actual today) so leads on the
  // last day of the month are fully evaluated rather than cut off early.
  const rangeMid  = midnight_(mtdRange.end);
  const todayMid  = rangeMid.getTime() < midnight_(today).getTime() ? rangeMid : midnight_(today);

  const vaMap     = {};
  const missedRows= [];

  leads.forEach(l => {
    if (l.bucket === 'booked') return;
    if (!inRange_(l.dateIn, mtdRange)) return;
    if (!l.dateIn) return;

    // Default TRUE — if account not listed, assume Sunday coverage applies
    const hasSunCov = sunCov.get(l.subaccount.toLowerCase()) !== false;
    const special   = isSpecialStatusLead_(l);
    const entries   = parseFollowUpEntries_(l.followUps, l.dateIn);

    // Special lead with zero calls = resolved on the spot, nothing to check
    if (special && entries.length === 0) return;

    // Cutoff: special leads stop at last call (or end of its callback window)
    let cutoff;
    if (special) {
      const last   = entries[entries.length - 1];
      const winEnd = last.cbWindow ? last.cbWindow.end : null;
      cutoff = (winEnd && winEnd > last.callDate) ? winEnd : last.callDate;
    } else {
      // Active leads: max 5 expected days (1 first contact + 4 follow-ups)
      const maxCutoff = midnight_(addExpectedDays_(l.dateIn, 4, hasSunCov));
      cutoff = maxCutoff < todayMid ? maxCutoff : todayMid;
    }

    const coverage = buildCoverageMap_(entries);

    if (!vaMap[l.va]) vaMap[l.va] = { count:0, expected:0, filled:0 };
    vaMap[l.va].count++;

    // Flag if first contact is 2+ expected days after lead date
    const firstCall  = entries.length > 0 ? entries[0].callDate : null;
    const maxOkFirst = midnight_(addExpectedDays_(l.dateIn, 1, hasSunCov));
    const lateFirst  = !!firstCall && firstCall > maxOkFirst;

    // Walk every expected day from lead date to cutoff.
    // Responsibility shifts on flex-off days (→ core VA) and core holidays (→ flex VA).
    const d = new Date(midnight_(l.dateIn));
    while (d <= cutoff) {
      if (!isExpectedNonWorkingDay_(d, hasSunCov)) {
        const respVA = getResponsibleVA_(new Date(d), l.va, l.subaccount, offDays, vaAssignments);
        if (!vaMap[respVA]) vaMap[respVA] = { count:0, expected:0, filled:0 };
        vaMap[respVA].expected++;

        const key = dateKey_(d);
        if (coverage.has(key)) {
          vaMap[respVA].filled++;
        } else {
          const isLateGap = lateFirst && !!firstCall && d < firstCall;
          let note = isLateGap ? '🚩 Late First Contact' : '';
          if (respVA !== l.va) {
            const shifted = isUSHoliday_(new Date(d))
              ? '🎌 Core holiday — Flex on duty'
              : '📅 Flex off — Core on duty';
            note = note ? note + ' · ' + shifted : shifted;
          }
          missedRows.push({
            va:       respVA,
            sub:      l.subaccount,
            name:     l.leadName,
            leadDate: fmtShort_(l.dateIn),
            missDate: fmtShort_(new Date(d)),
            status:   l.status,
            note,
            lateFlag: isLateGap
          });
        }
      }
      d.setDate(d.getDate() + 1);
    }
  });

  // ── Summary ──────────────────────────────────────────────
  const SCOLS = ['VA','Active Leads','Expected Days','Covered Days',
                 'Adherence %','Score (0–10)','Missed Days'];
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

  // ── Missed detail ─────────────────────────────────────────
  const DCOLS = ['VA','Sub-Account','Lead Name','Lead Date',
                 'Missed Date','Lead Status','Note'];
  const DW    = DCOLS.length;
  const W     = Math.max(SW, DW, 10);  // minimum 10 cols to match MTD Booking Rate

  const totalMissed = sumRows.reduce((s,r) => s+(Number(r[6])||0), 0);

  const sortedMissed = [...missedRows].sort((a,b) =>
    a.va.localeCompare(b.va) ||
    a.leadDate.localeCompare(b.leadDate) ||
    a.missDate.localeCompare(b.missDate)
  );

  const OUT = [];
  OUT.push(['FOLLOW-UP ADHERENCE — ' + monthLabel + ' — ' + lastRun]);
  OUT.push([
    'Every working day (Mon–Sat, no holidays) must be covered by a call or callback window. ' +
    'Callback windows parsed from notes (tomorrow / next week / specific date / later). ' +
    'Special-status leads: counted up to last call date only. ' +
    '🚩 Late First Contact = first follow-up was 2+ working days after lead arrived.'
  ]);
  OUT.push(Array(W).fill(''));
  OUT.push(SCOLS.concat(Array(W - SW).fill('')));
  sumRows.forEach(r => OUT.push(r.concat(Array(W - SW).fill(''))));
  OUT.push(Array(W).fill(''));
  OUT.push(['MISSED DAYS — ' + totalMissed + ' total'].concat(Array(W - 1).fill('')));
  OUT.push(DCOLS.concat(Array(W - DW).fill('')));
  sortedMissed.forEach(r => OUT.push(
    [r.va, r.sub, r.name, r.leadDate, r.missDate, r.status, r.note]
    .concat(Array(W - DW).fill(''))
  ));

  writeGrid_(sheet, OUT, W);

  // Header / summary formatting
  styleRow_(sheet, 1, W, '#1F3864','#FFFFFF', true,  11);
  styleRow_(sheet, 2, W, '#1F3864','#CCDDFF', false,  9);
  styleRow_(sheet, 4, W, '#2F5496','#FFFFFF', true,  10);

  const ds = 5;
  for (let i = 0; i < sumRows.length; i++) {
    const adh       = parseInt(sumRows[i][4]);
    const score     = sumRows[i][5];
    const adhCell   = sheet.getRange(ds + i, 5);
    const scoreCell = sheet.getRange(ds + i, 6);
    if (!isNaN(adh)) {
      let bg, fg;
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
  const detailStart   = detailHdrRow + 1;
  styleRow_(sheet, missHeaderRow, W, '#E8F0FE','#1F3864', true, 10);
  styleRow_(sheet, detailHdrRow,  W, '#4A86E8','#FFFFFF', true, 10);

  // Orange highlight for Late First Contact rows
  for (let i = 0; i < sortedMissed.length; i++) {
    if (sortedMissed[i].lateFlag) {
      sheet.getRange(detailStart + i, 1, 1, DW)
        .setBackground('#FF9900')
        .setFontColor('#FFFFFF')
        .setFontWeight('bold');
    }
  }

  sheet.getRange(1, 1, 1, 10).merge();
  sheet.getRange(2, 1, 1, 10).merge();
  sheet.setFrozenRows(4);
  sheet.autoResizeColumns(1, W);
  sheet.setColumnWidth(1, 80);   // VA
  sheet.setColumnWidth(2, 130);  // Sub-Account / Active Leads
  sheet.setColumnWidth(3, 150);  // Lead Name / Expected Days
  sheet.getRange(1, 1, 1, 1).setWrap(true);
  sheet.getRange(2, 1, 1, 1).setWrap(true);
  Logger.log('Follow-Up Adherence: ' + sumRows.length + ' VAs, ' + totalMissed + ' missed days');
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
    if (lastCol < SC_START_COL) return;

    const numCols = lastCol - SC_START_COL + 1;
    const vaRow   = sheet.getRange(SC_VA_ROW, SC_START_COL, 1, numCols).getValues()[0];

    Logger.log('Scorecard VA names (row 4 from col C): ' + JSON.stringify(vaRow));
    Logger.log('Source VA scores: ' + JSON.stringify(vaScores));

    vaRow.forEach((cell, i) => {
      const vaName = String(cell).trim();
      if (!vaName) return;
      const score = vaScores[vaName];
      if (typeof score === 'number') {
        sheet.getRange(SC_BOOKING_SCORE_ROW, SC_START_COL + i).setValue(score);
        Logger.log('Wrote score ' + score + ' for "' + vaName + '" at col ' + (SC_START_COL + i));
      } else {
        Logger.log('No score match for scorecard name: "' + vaName + '"');
      }
    });
    Logger.log('Booking scores written to scorecard row ' + SC_BOOKING_SCORE_ROW + ' from col C');
  } catch(e) {
    Logger.log('writeBookingScoresToScorecard_ error: ' + e.message);
  }
}

// ============================================================
// OFF-DAY CALENDAR  (flex VA off days from GAS web app)
// ============================================================

function loadOffDays_() {
  // Returns Map<"YYYY-MM-DD", Set<vaNameLowerCase>>
  const map = new Map();
  try {
    const resp = UrlFetchApp.fetch(
      CALENDAR_API + '?action=getSharedEntries',
      {muteHttpExceptions: true, followRedirects: true}
    );
    if (resp.getResponseCode() !== 200) {
      Logger.log('loadOffDays_ HTTP ' + resp.getResponseCode());
      return map;
    }
    const data    = JSON.parse(resp.getContentText());
    const entries = data.entries || {};
    for (const [dateKey, items] of Object.entries(entries)) {
      for (const e of items) {
        if (e.type === 'flex-off' && e.name) {
          if (!map.has(dateKey)) map.set(dateKey, new Set());
          map.get(dateKey).add(e.name.toLowerCase().trim());
        }
      }
    }
    Logger.log('loadOffDays_: ' + map.size + ' dates with off entries');
  } catch(e) {
    Logger.log('loadOffDays_ error: ' + e.message);
  }
  return map;
}

// ============================================================
// SUNDAY COVERAGE  (per-account, from coverage spreadsheet)
// ============================================================

function loadSundayCoverage_() {
  // Returns Map<subaccountLowerCase, boolean>
  // true  = has Sunday coverage (row contains ✅)
  // false = no Sunday coverage  (row contains ❌, or row has no emoji)
  // Accounts NOT in the map → default is TRUE (see call site: !==false)
  const map = new Map();
  try {
    const ss = SpreadsheetApp.openById(COVERAGE_SS_ID);
    ss.getSheets().forEach(sheet => {
      const data = sheet.getDataRange().getValues();
      let curAcct = null;
      data.forEach(row => {
        const cells   = row.map(c => String(c || '').trim());
        const rowText = cells.join(' ');
        const c0      = cells[0];
        if (!c0) return;

        const isSunRow = /sunday/i.test(rowText);
        if (!isSunRow) {
          if (c0.length > 1) curAcct = c0;
          return;
        }
        if (!curAcct) return;

        // Only ✅ emoji means yes; ❌ or no emoji means no
        const hasCov = /✅/.test(rowText) && !/❌/.test(rowText);

        map.set(curAcct.toLowerCase(), hasCov);
        Logger.log('Sunday coverage: "' + curAcct + '" = ' + (hasCov ? 'yes ✅' : 'no ❌'));
      });
    });
    Logger.log('loadSundayCoverage_: ' + map.size + ' accounts loaded');
  } catch(e) {
    Logger.log('loadSundayCoverage_ error: ' + e.message);
  }
  return map;
}

// ============================================================
// VA RESPONSIBILITY ATTRIBUTION
// ============================================================

// Returns the VA who should be held responsible for a given missed/expected day.
// Rules:
//   1. US holiday (core VA off) → flex VA on duty (if present and effective)
//   2. Flex VA off day          → core VA on duty
//   3. Otherwise                → defaultVA (col A)
// Sub-account assignments come from the "VA Assignments" tab in COVERAGE_SS_ID.
function getResponsibleVA_(date, defaultVA, subaccount, offDays, vaAssignments) {
  const asmt   = vaAssignments.get(subaccount.toLowerCase());
  const coreVA = asmt && asmt.coreVA ? asmt.coreVA : defaultVA;
  let   flexVA = asmt && asmt.flexVA ? asmt.flexVA : null;

  // Respect effective-date for mid-period flex transitions (e.g. Jessica from June 5)
  if (flexVA) {
    const effDate = FLEX_VA_EFFECTIVE[flexVA.toLowerCase()];
    if (effDate && midnight_(date) < midnight_(effDate)) flexVA = null;
  }

  const dateK   = dateKeyPadded_(date);
  const flexOff = offDays.get(dateK) || new Set();

  if (isUSHoliday_(date)) {
    // Core VA is off — flex VA covers (if present and effective)
    if (flexVA && !flexOff.has(flexVA.toLowerCase())) return flexVA;
  }

  if (flexVA && flexOff.has(flexVA.toLowerCase())) {
    // Flex VA is off — core VA covers
    return coreVA;
  }

  return defaultVA;
}

// ============================================================
// VA ASSIGNMENTS  (sub-account → core VA + flex VA)
// ============================================================

// Reads the "VA Assignments" tab from the coverage spreadsheet.
// Tab format (row 1 = headers skipped):
//   Col A: Sub-Account  |  Col B: Core VA  |  Col C: Flex VA
// Update this tab whenever coverage assignments change.
function loadVAAssignments_() {
  const map = new Map();
  try {
    const ss    = SpreadsheetApp.openById(COVERAGE_SS_ID);
    const sheet = ss.getSheetByName('VA Assignments');
    if (!sheet) {
      Logger.log('loadVAAssignments_: "VA Assignments" tab not found — all days use col-A VA');
      return map;
    }
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const sub    = String(rows[i][0] || '').trim();
      const coreVA = String(rows[i][1] || '').trim();
      const flexVA = String(rows[i][2] || '').trim();
      if (sub) map.set(sub.toLowerCase(), { coreVA, flexVA });
    }
    Logger.log('loadVAAssignments_: ' + map.size + ' accounts loaded');
  } catch(e) {
    Logger.log('loadVAAssignments_ error: ' + e.message);
  }
  return map;
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

// Sunday-coverage-aware non-working day check (used in per-lead day walks)
function isExpectedNonWorkingDay_(date, hasSunCov) {
  const dow = date.getDay();
  if (dow === 0) return !hasSunCov;  // Sunday: skip unless account has coverage
  return isUSHoliday_(date);
}

// Sunday-coverage-aware "add N expected days" (for 5-day cap + late-first-contact)
function addExpectedDays_(startDate, n, hasSunCov) {
  const d = new Date(startDate.getTime());
  let added = 0;
  while (added < n) {
    d.setDate(d.getDate() + 1);
    if (!isExpectedNonWorkingDay_(d, hasSunCov)) added++;
  }
  return d;
}

// Zero-padded date key for calendar API lookups: "YYYY-MM-DD"
function dateKeyPadded_(d) {
  return d.getFullYear() + '-'
    + String(d.getMonth() + 1).padStart(2, '0') + '-'
    + String(d.getDate()).padStart(2, '0');
}


function isNonWorkingDay_(date) {
  if (date.getDay() === 0) return true;
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

  add(new Date(year,  0,  1));  // New Year's Day
  add(new Date(year,  5, 19));  // Juneteenth
  add(new Date(year,  6,  4));  // Independence Day
  add(new Date(year, 10, 11));  // Veterans Day
  add(new Date(year, 11, 25));  // Christmas

  add(nthWeekday_(year,  0, 1, 3));  // MLK Day
  add(nthWeekday_(year,  1, 1, 3));  // Presidents Day
  add(lastWeekday_(year, 4, 1));     // Memorial Day
  add(nthWeekday_(year,  8, 1, 1));  // Labor Day
  add(nthWeekday_(year,  9, 1, 2));  // Columbus Day
  add(nthWeekday_(year, 10, 4, 4));  // Thanksgiving

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

function getPrevMonthRange_(today) {
  // Full previous calendar month (1st → last day)
  const prev = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  return {
    start: new Date(prev.getFullYear(), prev.getMonth(), 1, 0, 0, 0, 0),
    end:   new Date(prev.getFullYear(), prev.getMonth() + 1, 0, 23, 59, 59, 999)
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
function fmtMonthShort_(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'MMMM');
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

    sheet.getRange(2, 1, last-1, 16).getValues().forEach(row => {
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
  ScriptApp.newTrigger('refreshLiveStats').timeBased().everyMinutes(5).create();
  SpreadsheetApp.getUi().alert(
    '✅ Live trigger set!\n\nSheets update every 5 minutes.\n\nTo stop: menu → Stop Auto-Refresh.'
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
    .addItem('▶ Start Auto-Refresh (every 5 min)', 'setupLiveTrigger')
    .addItem('⏹ Stop Auto-Refresh',                'stopLiveTrigger')
    .addToUi();
}
