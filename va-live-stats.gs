// ============================================================
// VA Live Stats — Booking Rate + Follow-Up Adherence
// ============================================================
// INSTALL (separate GAS project):
//  1. Open the VA Team Scorecard spreadsheet
//  2. Extensions > Apps Script → paste this file, Save
//  3. Run setupLiveTrigger() ONCE → auto-refreshes every minute
//  4. Run refreshLiveStats() now to populate immediately
//
// WHAT IT DOES:
//  Reads all leads from the source spreadsheet, computes two metrics
//  per VA for the current month, and writes them into the
//  "VA Team Scorecard" sheet:
//    Row 4  — VA names (read-only, used to find each VA's column)
//    Row 7  — MTD Booking Rate        (written as %, e.g. 52%)
//    Row 8  — Follow-Up Adherence %   (written as %, e.g. 87%)
// ============================================================

// ── CONFIG ──────────────────────────────────────────────────
const SOURCE_SS_ID    = '1Kd0smsW26IJR1H5kJ0MP0kH8jPOTazo2_TbLEkgmbRs';
const SCORECARD_SS_ID = '1z6mLBx2gfoqFpSyjgg4C_69oB3eOT3lBmRynKWIwybk';
const SCORECARD_SHEET = 'VA Team Scorecard';
const SC_VA_ROW       = 4;   // row that holds VA names (read-only)
const SC_BOOKING_ROW  = 7;   // row to write MTD booking rate
const SC_FFUP_ROW     = 8;   // row to write follow-up adherence %

const MAIN_SHEETS     = ['ROOF, MAIN', 'HVAC, MAIN', 'GUTTER, Main', 'WINDOWS, MAIN'];

// Source column indices (0-based)
const COL_VA        = 0;   // A — VA name
const COL_DATE_IN   = 1;   // B — date lead came in
const COL_SUBACCOUNT= 2;   // C — sub-account
const COL_STATUS    = 6;   // G — status
const COL_DISP      = 8;   // I — disposition
const COL_DATE_CONF = 9;   // J — date confirmed
const COL_FU1       = 10;  // K–O — follow-up slots (5 total)

// Accounts where "Client Handles" is excluded from the VA's booking rate.
const EXCLUDE_CLIENT_HANDLES_ACCTS = ['Outstanding Roofing', 'Good Guy Roofing'];

// ============================================================
// ENTRY POINT
// ============================================================

function refreshLiveStats() {
  const today    = new Date();
  const mtdRange = getMTDRange(today);
  const leads    = loadLeads();

  const bookingRates  = computeBookingRates(leads, mtdRange);
  const adherenceRates= computeFollowUpAdherence(leads, mtdRange, today);

  writeToScorecard(bookingRates, adherenceRates);
  Logger.log('Live stats updated — ' + today.toISOString());
}

// ============================================================
// METRIC 1 — BOOKING RATE
// ============================================================
//
// FORMULA:  Booking Rate = Booked ÷ Qualified Leads
//
// BOOKED leads:
//   Status is "Confirmed" OR "Manual Booked"
//
// QUALIFIED leads = all MTD leads MINUS the following exclusions:
//   • Bucket = Cancelled/Invalid  (status contains "cancelled" or "invalid")
//   • Disposition = OSA, Unqualified, Dup Lead, # Issue
//     → these are leads that never should have been worked (bad data / wrong area)
//   • For Outstanding Roofing and Good Guy Roofing ONLY:
//     Bucket = Client Handles (status = "Client Handles" or "Satellite Quote")
//     → on these two accounts the homeowner manages their own booking, so
//       the VA can't be held responsible for those outcomes
//
// WHY these exclusions?
//   The KPI measures how well the VA converts workable leads.
//   Leads that are invalid, out-of-area, or handled by the client
//   aren't in the VA's control, so including them would unfairly
//   lower their rate.
//
// KPI THRESHOLD: 45%+
//
// ============================================================

function computeBookingRates(leads, mtdRange) {
  const EXCL_DISP = new Set(['osa', 'unqualified', 'dup lead', '# issue', '#issue']);

  const vaMap = {};

  leads
    .filter(l => inRange(l.dateIn, mtdRange) || inRange(l.dateConf, mtdRange))
    .forEach(l => {
      if (!vaMap[l.va]) vaMap[l.va] = { booked: 0, qualified: 0 };
      const m = vaMap[l.va];

      // Count booked
      if (l.bucket === 'booked') m.booked++;

      // Determine if this lead counts toward qualified denominator
      const isSpecialAcct = EXCLUDE_CLIENT_HANDLES_ACCTS.includes(l.subaccount);
      const excluded = (l.bucket === 'cancelled')             // invalid / cancelled
                    || EXCL_DISP.has(l.disp)                  // bad-data dispositions
                    || (isSpecialAcct && l.bucket === 'clientHandles'); // special acct CH
      if (!excluded) m.qualified++;
    });

  // Return map:  { 'VA Name': 0.52, ... }
  const result = {};
  Object.entries(vaMap).forEach(([va, m]) => {
    result[va] = m.qualified > 0 ? m.booked / m.qualified : null;
  });
  return result;
}

// ============================================================
// METRIC 2 — FOLLOW-UP ADHERENCE %
// ============================================================
//
// FORMULA:  Adherence % = Total Filled Slots ÷ Total Expected Slots
//
// WHICH LEADS ARE TRACKED:
//   Active MTD leads only — meaning leads where:
//     • Bucket is NOT booked (already converted — no follow-up needed)
//     • Bucket is NOT cancelled/invalid
//     • Disposition is NOT OSA, Unqualified, Dup Lead, # Issue
//
// HOW EXPECTED SLOTS ARE COUNTED:
//   Each lead gets a 5-slot follow-up cadence (columns K through O).
//   How many of those 5 slots are "expected" depends on how many days
//   have passed since the lead came in:
//
//     Days since lead came in  →  Expected follow-up slots
//     0 days                   →  1  (Day 1 should be filled)
//     1 day                    →  2  (Day 1 + Day 2)
//     2 days                   →  3
//     3 days                   →  4
//     4+ days                  →  5  (all slots expected)
//
//   Formula: expect = min(5, days_since_lead + 1)
//   Minimum is always 1 (even same-day leads need a Day 1 note).
//
// HOW FILLED SLOTS ARE COUNTED:
//   A slot is "filled" if the corresponding column (K, L, M, N, or O)
//   contains any non-empty text.
//
// EXAMPLE:
//   VA has 3 active leads:
//     Lead A — 4 days old, 5 slots expected, 5 filled  → 5/5
//     Lead B — 2 days old, 3 slots expected, 2 filled  → 2/3
//     Lead C — 0 days old, 1 slot expected,  0 filled  → 0/1
//   Total filled = 7, total expected = 9
//   Adherence = 7/9 = 78%
//
// ============================================================

function computeFollowUpAdherence(leads, mtdRange, today) {
  const EXCL_DISP = new Set(['osa', 'unqualified', 'dup lead', '# issue', '#issue']);

  const vaMap = {};

  leads
    .filter(l => {
      if (l.bucket === 'booked')    return false;  // already converted
      if (l.bucket === 'cancelled') return false;  // invalid/cancelled
      if (EXCL_DISP.has(l.disp))   return false;  // bad-data dispositions
      return inRange(l.dateIn, mtdRange) || inRange(l.dateConf, mtdRange);
    })
    .forEach(l => {
      if (!vaMap[l.va]) vaMap[l.va] = { filled: 0, expected: 0 };
      const m = vaMap[l.va];

      const daysSinceIn = l.dateIn ? Math.floor((today - l.dateIn) / 86400000) : 0;
      const expect      = Math.min(5, Math.max(1, daysSinceIn + 1));
      m.expected += expect;

      // Count how many of the expected slots are actually filled
      for (let k = 0; k < expect; k++) {
        if (l.followUps[k]) m.filled++;
      }
    });

  // Return map:  { 'VA Name': 0.78, ... }
  const result = {};
  Object.entries(vaMap).forEach(([va, m]) => {
    result[va] = m.expected > 0 ? m.filled / m.expected : null;
  });
  return result;
}

// ============================================================
// SCORECARD WRITE
// ============================================================

function writeToScorecard(bookingRates, adherenceRates) {
  try {
    const ss    = SpreadsheetApp.openById(SCORECARD_SS_ID);
    const sheet = ss.getSheetByName(SCORECARD_SHEET);
    if (!sheet) { Logger.log('Sheet not found: ' + SCORECARD_SHEET); return; }

    const lastCol = sheet.getLastColumn();
    if (lastCol < 1) return;

    // Read VA names from row 4
    const vaRow = sheet.getRange(SC_VA_ROW, 1, 1, lastCol).getValues()[0];

    vaRow.forEach((cell, i) => {
      const vaName = String(cell).trim();
      if (!vaName) return;
      const col = i + 1;

      const br = bookingRates[vaName];
      if (br !== null && br !== undefined) {
        sheet.getRange(SC_BOOKING_ROW, col).setValue(br).setNumberFormat('0%');
      }

      const ar = adherenceRates[vaName];
      if (ar !== null && ar !== undefined) {
        sheet.getRange(SC_FFUP_ROW, col).setValue(ar).setNumberFormat('0%');
      }
    });

    Logger.log('Scorecard written — booking: ' + Object.keys(bookingRates).length +
               ' VAs, adherence: ' + Object.keys(adherenceRates).length + ' VAs');
  } catch(e) {
    Logger.log('writeToScorecard error: ' + e.message);
  }
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

      const status = String(row[COL_STATUS] || '').trim();
      const disp   = String(row[COL_DISP]   || '').trim().toLowerCase();

      leads.push({
        va,
        bucket:     getStatusBucket(status),
        disp,
        dateIn:     parseDate(row[COL_DATE_IN]),
        dateConf:   parseDate(row[COL_DATE_CONF]),
        subaccount: String(row[COL_SUBACCOUNT] || '').trim(),
        followUps:  Array.from({length:5}, (_,k) => String(row[COL_FU1+k]||'').trim())
      });
    });
  });

  return leads;
}

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

function inRange(d, r) { return !!d && d >= r.start && d <= r.end; }

// ============================================================
// TRIGGER
// ============================================================

function setupLiveTrigger() {
  // Remove any existing live trigger first
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'refreshLiveStats') ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger('refreshLiveStats')
    .timeBased()
    .everyMinutes(1)
    .create();

  SpreadsheetApp.getUi().alert(
    '✅ Live trigger set!\n\n' +
    'Booking Rate (row 7) and Follow-Up Adherence (row 8) will update every minute.\n\n' +
    'To stop: run stopLiveTrigger() from the menu.'
  );
}

function stopLiveTrigger() {
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'refreshLiveStats') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  SpreadsheetApp.getUi().alert(
    removed > 0 ? '⏹ Live trigger stopped.' : 'No live trigger was running.'
  );
}

// ============================================================
// CUSTOM MENU
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 VA Live Stats')
    .addItem('⚡ Refresh Now',                   'refreshLiveStats')
    .addSeparator()
    .addItem('▶ Start Auto-Refresh (every 1 min)', 'setupLiveTrigger')
    .addItem('⏹ Stop Auto-Refresh',               'stopLiveTrigger')
    .addToUi();
}
