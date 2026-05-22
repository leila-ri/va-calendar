// ============================================================
//  LAUNCH & FLAG TRACKER — Google Apps Script
//  Monitors Slack #internal-team for:
//    • "Launched / Launch / Live" → "Launched Accounts" sheet
//    • "red" / "orange"          → "Flagged Accounts" sheet
// ============================================================

// Set these in Apps Script: File → Project properties → Script properties
// SLACK_TOKEN  → your Slack bot token (xoxb-...)
// CHANNEL_ID   → the channel ID to monitor (e.g. C08JV24HW5B)
var SLACK_TOKEN  = PropertiesService.getScriptProperties().getProperty("SLACK_TOKEN")  || "";
var CHANNEL_ID   = PropertiesService.getScriptProperties().getProperty("CHANNEL_ID")   || "";

// ── Sheet names ─────────────────────────────────────────────
var SHEET_NAME      = "Launched Accounts";
var FLAG_SHEET_NAME = "Flagged Accounts";

// ── Keywords ────────────────────────────────────────────────
var KEYWORDS = ["launched", "launch update", "is now live", "is live",
                "went live", "now live", "has launched", "has been launched"];

var FLAG_KEYWORDS = ["red", "orange"];

// ── Add custom menu when spreadsheet opens ──────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🚀 Launch Tracker")
    .addItem("Sync Now",               "manualSync")
    .addItem("Debug Connection",       "debugSync")
    .addItem("Full Reset & Reimport",  "fullReset")
    .addItem("Initial Setup",          "setupLaunchTracker")
    .addSeparator()
    .addItem("Setup Flag Tracker",     "setupFlaggedTracker")
    .addItem("Sync Flagged Accounts",  "syncFlaggedPosts")
    .addToUi();
}

// ── One-time setup: Launch Tracker ──────────────────────────
function setupLaunchTracker() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);

    var headers = ["Company Name", "Date Launched", "VA", "Posted By", "Slack Message", "Notes"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    var hdr = sheet.getRange(1, 1, 1, headers.length);
    hdr.setFontWeight("bold");
    hdr.setBackground("#1264A3");   // Slack blue
    hdr.setFontColor("#FFFFFF");
    sheet.setFrozenRows(1);

    sheet.setColumnWidth(1, 230);
    sheet.setColumnWidth(2, 120);
    sheet.setColumnWidth(3, 120);
    sheet.setColumnWidth(4, 140);
    sheet.setColumnWidth(5, 420);
    sheet.setColumnWidth(6, 200);
  }

  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "syncLaunchPosts") {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("syncLaunchPosts")
    .timeBased()
    .everyMinutes(30)
    .create();

  PropertiesService.getScriptProperties().deleteProperty("lastProcessedTs");
  syncLaunchPosts();

  // Also set up the flag tracker on first run
  setupFlaggedTracker();

  SpreadsheetApp.getUi().alert(
    "✅ Launch Tracker is live!\n\n" +
    "All historical launches have been imported.\n" +
    "New launches will sync automatically every 30 minutes.\n\n" +
    "You can also use  🚀 Launch Tracker → Sync Now  at any time."
  );
}

// ── One-time setup: Flag Tracker ─────────────────────────────
function setupFlaggedTracker() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(FLAG_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(FLAG_SHEET_NAME);

    var headers = ["Company Name", "Date Flagged", "Flag Color", "VA", "Posted By", "Slack Message", "Notes"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    var hdr = sheet.getRange(1, 1, 1, headers.length);
    hdr.setFontWeight("bold");
    hdr.setBackground("#B22222");   // dark red
    hdr.setFontColor("#FFFFFF");
    sheet.setFrozenRows(1);

    sheet.setColumnWidth(1, 230);
    sheet.setColumnWidth(2, 120);
    sheet.setColumnWidth(3, 110);
    sheet.setColumnWidth(4, 120);
    sheet.setColumnWidth(5, 140);
    sheet.setColumnWidth(6, 420);
    sheet.setColumnWidth(7, 200);
  }

  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "syncFlaggedPosts") {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("syncFlaggedPosts")
    .timeBased()
    .everyMinutes(30)
    .create();

  PropertiesService.getScriptProperties().deleteProperty("lastFlaggedTs");
  syncFlaggedPosts();
}

// ── Manual sync from menu (both sheets) ─────────────────────
function manualSync() {
  syncLaunchPosts();
  syncFlaggedPosts();
  SpreadsheetApp.getUi().alert("✅ Sync complete!");
}

// ── Debug: test connection and show what's being fetched ─────
function debugSync() {
  var props = PropertiesService.getScriptProperties();
  var lastTs = props.getProperty("lastProcessedTs") || "0";

  var testRes  = UrlFetchApp.fetch(
    "https://slack.com/api/conversations.history?channel=" + CHANNEL_ID + "&limit=1",
    { headers: { "Authorization": "Bearer " + SLACK_TOKEN }, muteHttpExceptions: true }
  );
  var testData = JSON.parse(testRes.getContentText());

  if (!testData.ok) {
    SpreadsheetApp.getUi().alert("❌ Slack API error: " + testData.error + "\n\nCheck token and scopes.");
    return;
  }

  var messages = fetchSlackMessages("0");
  var matched  = messages.filter(function(m) {
    if (!m.text) return false;
    var lower = m.text.toLowerCase();
    return KEYWORDS.some(function(kw) { return lower.indexOf(kw) !== -1; });
  });
  var flagged  = messages.filter(function(m) {
    if (!m.text) return false;
    var lower = m.text.toLowerCase();
    return FLAG_KEYWORDS.some(function(kw) { return lower.indexOf(kw) !== -1; });
  });

  SpreadsheetApp.getUi().alert(
    "✅ Slack connected!\n\n" +
    "Total messages fetched: "              + messages.length + "\n" +
    "Messages matching launch keywords: "   + matched.length  + "\n" +
    "Messages matching flag keywords: "     + flagged.length  + "\n" +
    "Stored lastTs (launches): "            + lastTs          + "\n\n" +
    (matched.length > 0
      ? "Sample launch match:\n" + matched[0].text.substring(0, 150)
      : "⚠️ No messages matched launch keywords: " + KEYWORDS.join(", "))
  );
}

// ── Full reset: clears both sheets and re-imports everything ─
function fullReset() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty("lastProcessedTs");
  props.deleteProperty("lastFlaggedTs");

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var launchSheet = ss.getSheetByName(SHEET_NAME);
  if (launchSheet && launchSheet.getLastRow() > 1) {
    launchSheet.getRange(2, 1, launchSheet.getLastRow() - 1, 6).clearContent();
  }

  var flagSheet = ss.getSheetByName(FLAG_SHEET_NAME);
  if (flagSheet && flagSheet.getLastRow() > 1) {
    flagSheet.getRange(2, 1, flagSheet.getLastRow() - 1, 7).clearContent();
  }

  syncLaunchPosts();
  syncFlaggedPosts();
  SpreadsheetApp.getUi().alert("✅ Full reset complete!");
}

// ── Core sync: launch posts → "Launched Accounts" ───────────
function syncLaunchPosts() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { setupLaunchTracker(); return; }

  var props  = PropertiesService.getScriptProperties();
  var lastTs = props.getProperty("lastProcessedTs") || "0";

  var companyData  = getCompanyData(ss);
  var companyNames = companyData.names;
  var vaMap        = companyData.vaMap;

  var messages = fetchSlackMessages(lastTs);
  if (!messages.length) { Logger.log("No new messages."); return; }

  messages.reverse();

  var newRows   = [];
  var latestTs  = parseFloat(lastTs);
  var userCache = {};

  messages.forEach(function(msg) {
    if (!msg.text) return;

    var lowerText = msg.text.toLowerCase();
    var hit = KEYWORDS.some(function(kw) { return lowerText.indexOf(kw) !== -1; });
    if (!hit) return;

    var ts = parseFloat(msg.ts);
    if (ts > latestTs) latestTs = ts;

    var date     = Utilities.formatDate(new Date(ts * 1000),
                                        Session.getScriptTimeZone(), "MM/dd/yyyy");
    var company  = findCompanyInText(msg.text, companyNames)
                   || extractCompanyFromMessage(msg.text)
                   || "";
    if (!company) return;

    var va       = vaMap[company] || "";
    var postedBy = resolveUser(msg.user, userCache);
    var preview  = msg.text.replace(/\n+/g, " ").substring(0, 400);

    newRows.push([company, date, va, postedBy, preview, ""]);
  });

  if (newRows.length) {
    var lastRow = Math.max(sheet.getLastRow(), 1);
    sheet.getRange(lastRow + 1, 1, newRows.length, 6).setValues(newRows);

    for (var r = 0; r < newRows.length; r++) {
      var rowNum = lastRow + 1 + r;
      var bg = (rowNum % 2 === 0) ? "#EAF4FB" : "#FFFFFF";
      sheet.getRange(rowNum, 1, 1, 6).setBackground(bg);
    }

    Logger.log("Added " + newRows.length + " launch entries.");
  }

  if (latestTs > parseFloat(lastTs)) {
    props.setProperty("lastProcessedTs", latestTs.toString());
  }
}

// ── Core sync: flagged posts → "Flagged Accounts" ────────────
function syncFlaggedPosts() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(FLAG_SHEET_NAME);
  if (!sheet) { setupFlaggedTracker(); return; }

  var props  = PropertiesService.getScriptProperties();
  var lastTs = props.getProperty("lastFlaggedTs") || "0";

  var companyData  = getCompanyData(ss);
  var companyNames = companyData.names;
  var vaMap        = companyData.vaMap;

  var messages = fetchSlackMessages(lastTs);
  if (!messages.length) { Logger.log("No new messages for flags."); return; }

  messages.reverse();

  var newRows   = [];
  var latestTs  = parseFloat(lastTs);
  var userCache = {};

  messages.forEach(function(msg) {
    if (!msg.text) return;

    var lowerText = msg.text.toLowerCase();
    var hit = FLAG_KEYWORDS.some(function(kw) { return lowerText.indexOf(kw) !== -1; });
    if (!hit) return;

    var ts = parseFloat(msg.ts);
    if (ts > latestTs) latestTs = ts;

    var date      = Utilities.formatDate(new Date(ts * 1000),
                                         Session.getScriptTimeZone(), "MM/dd/yyyy");
    var company   = findCompanyInText(msg.text, companyNames)
                    || extractCompanyFromMessage(msg.text)
                    || "";
    if (!company) return;

    var flagColor = detectFlagColor(msg.text);
    var va        = vaMap[company] || "";
    var postedBy  = resolveUser(msg.user, userCache);
    var preview   = msg.text.replace(/\n+/g, " ").substring(0, 400);

    newRows.push([company, date, flagColor, va, postedBy, preview, ""]);
  });

  if (newRows.length) {
    var lastRow = Math.max(sheet.getLastRow(), 1);
    sheet.getRange(lastRow + 1, 1, newRows.length, 7).setValues(newRows);

    for (var r = 0; r < newRows.length; r++) {
      var rowNum = lastRow + 1 + r;
      var bg = (rowNum % 2 === 0) ? "#FDECEA" : "#FFFFFF";   // light red tint on alternating rows
      sheet.getRange(rowNum, 1, 1, 7).setBackground(bg);

      // Color the "Flag Color" cell (column 3) to match the flag
      var color = newRows[r][2];
      var cellBg = color === "Red"    ? "#FF4D4D"
                 : color === "Orange" ? "#FFA500"
                 :                      "#FF8C00";   // Red & Orange → dark orange
      sheet.getRange(rowNum, 3).setBackground(cellBg).setFontColor("#FFFFFF").setFontWeight("bold");
    }

    Logger.log("Added " + newRows.length + " flagged entries.");
  }

  if (latestTs > parseFloat(lastTs)) {
    props.setProperty("lastFlaggedTs", latestTs.toString());
  }
}

// ── Detect flag color from message text ─────────────────────
function detectFlagColor(text) {
  var lower    = text.toLowerCase();
  var hasRed   = lower.indexOf("red")    !== -1;
  var hasOrange= lower.indexOf("orange") !== -1;
  if (hasRed && hasOrange) return "Red & Orange";
  if (hasRed)   return "Red";
  if (hasOrange) return "Orange";
  return "";
}

// ── Fetch all Slack messages since `oldest` (paginated) ──────
function fetchSlackMessages(oldest) {
  var all    = [];
  var cursor = null;

  for (var page = 0; page < 50; page++) {
    var params = { channel: CHANNEL_ID, limit: 200 };
    if (oldest && oldest !== "0") params.oldest = oldest;
    if (cursor)                    params.cursor  = cursor;

    var qs  = Object.keys(params)
                .map(function(k) { return k + "=" + encodeURIComponent(params[k]); })
                .join("&");

    var res  = UrlFetchApp.fetch("https://slack.com/api/conversations.history?" + qs, {
      headers: { "Authorization": "Bearer " + SLACK_TOKEN },
      muteHttpExceptions: true
    });
    var data = JSON.parse(res.getContentText());

    if (!data.ok) { Logger.log("Slack error: " + data.error); break; }

    all    = all.concat(data.messages || []);
    cursor = (data.response_metadata || {}).next_cursor || null;
    if (!cursor) break;

    Utilities.sleep(800);
  }

  return all;
}

// ── Get company names + VA map from the first sheet ──────────
function getCompanyData(ss) {
  var data   = ss.getSheets()[0].getDataRange().getValues();
  var header = data[0];
  var subCol = 1;
  var vaCol  = 2;

  for (var j = 0; j < header.length; j++) {
    var h = header[j].toString().toLowerCase();
    if (h.indexOf("sub-account") !== -1 || h.indexOf("sub account") !== -1) subCol = j;
    if (h === "va") vaCol = j;
  }

  var names = [];
  var vaMap = {};

  for (var i = 1; i < data.length; i++) {
    var name = data[i][subCol] ? data[i][subCol].toString().trim() : "";
    var va   = data[i][vaCol]  ? data[i][vaCol].toString().trim()  : "";
    if (name && name.toLowerCase() !== "inactive") {
      names.push(name);
      if (va) vaMap[name] = va;
    }
  }

  names.sort(function(a, b) { return b.length - a.length; });
  return { names: names, vaMap: vaMap };
}

// ── Strip Slack formatting from message text ─────────────────
function cleanSlackText(text) {
  return text
    .replace(/<@[A-Z0-9]+>/g, "")
    .replace(/<#[A-Z0-9]+\|[^>]*>/g, "")
    .replace(/<#[A-Z0-9]+>/g, "")
    .replace(/:[a-z0-9_\-+]+:/g, "")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g,  "<")
    .replace(/&gt;/g,  ">")
    .replace(/\*/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

// ── Strip common announcement prefixes ───────────────────────
function trimAnnouncementPrefix(name) {
  return name
    .replace(/^(hey[\s,!]+|hi[\s,!]+|attention\s+team[\s,!]*)/i, "")
    .replace(/^(fyi[\s,:]+|heads\s+up[\s,:]+|update[\s,:]+)/i, "")
    .replace(/^(good\s+(morning|afternoon|evening)[\s,!]+)/i, "")
    .replace(/\s+/g, " ")
    .trim();
}

// ── Validate an extracted company name candidate ──────────────
function isValidCompanyName(name) {
  if (!name || name.length < 4 || name.length > 90) return false;
  if (!/[a-zA-Z]/.test(name)) return false;
  if (name.split(/\s+/).length > 7) return false;
  if (/^(today|yesterday|a\s+few|already|as\s+of|this\s+morning|now|just|at\s+\d)/i.test(name)) return false;
  if (/^\d+(:\d+)?\s*(am|pm|a|p)?\s*$/i.test(name)) return false;
  if (/\b(ahead\s+of|understand|getting|because|although|however|since|during|please|available)\b/i.test(name)) return false;
  return true;
}

var GENERIC_WORDS = /^(roofing|construction|contracting|contractors|contractor|exterior|exteriors|home|solutions|solution|services|service|group|company|co|inc|llc|and|of|the|a|an|systems|system|restoration|remodeling|windows|gutters|gutter|heating|cooling|air|hvac|general|building|builders|builder|management|enterprises|enterprise|properties|property|professionals|pros|residential|commercial|repair|repairs|supply|supplies|climate|weather|energy|power|plus|all|top|best|pro|elite|premium|quality|first|one|new|old|east|west|north|south)$/i;

function getKeyWords(name) {
  return name
    .replace(/[&+\/\\]/g, " ")
    .split(/[\s\-_.()'*"]+/)
    .filter(function(w) { return w.length >= 2 && !GENERIC_WORDS.test(w); });
}

// ── Find a known company name inside a Slack message ─────────
function findCompanyInText(text, names) {
  if (!text || !names.length) return null;
  var cleaned = cleanSlackText(text).toLowerCase();

  for (var i = 0; i < names.length; i++) {
    if (cleaned.indexOf(names[i].toLowerCase()) !== -1) return names[i];
  }

  var bestName  = null;
  var bestScore = 0;
  var THRESHOLD = 0.6;

  for (var i = 0; i < names.length; i++) {
    var keys = getKeyWords(names[i]);
    if (!keys.length) continue;

    var hits = keys.filter(function(w) {
      return cleaned.indexOf(w.toLowerCase()) !== -1;
    }).length;

    var score = hits / keys.length;
    if (score > bestScore && score >= THRESHOLD) {
      bestScore = score;
      bestName  = names[i];
    }
  }

  return bestName;
}

// ── Strip "is now live" etc. from end of a candidate ─────────
function trimLaunchSuffix(name) {
  return name
    .replace(/\s+is\s+now\s+(live|launched).*/i, "")
    .replace(/\s+is\s+(live|launched).*/i, "")
    .replace(/\s+went\s+live.*/i, "")
    .replace(/\s+now\s+live.*/i, "")
    .replace(/\s+has\s+(been\s+)?launched.*/i, "")
    .replace(/[\s!*:.\-–—]+$/, "")
    .trim();
}

// ── Try regex patterns to extract company name ───────────────
function extractCompanyFromMessage(rawText) {
  if (!rawText) return null;

  var text = cleanSlackText(rawText);

  var patterns = [
    /launch\s+(?:update|overview)\s*[:\s\-–—]+\s*([A-Z][^\n]{3,80})/im,
    /\bon\s+([A-Z][^:\n]{3,70}):/i,
    /^\s*([A-Z][^\n]{3,80}?)\s+is\s+now\s+(?:live|launched)/im,
    /^\s*([A-Z][^\n]{3,80}?)\s+is\s+live/im,
    /^\s*([A-Z][^\n]{3,80}?)\s+went\s+live/im,
    /^\s*([A-Z][^\n]{3,80}?)\s+now\s+live/im,
    /^\s*([A-Z][^\n]{3,80}?)\s+has\s+(?:been\s+)?launched/im,
    /launched[:\s\-–—]+([A-Z][^\n!]{3,80})/i,
  ];

  for (var i = 0; i < patterns.length; i++) {
    var m = text.match(patterns[i]);
    if (m && m[1]) {
      var candidate = trimAnnouncementPrefix(trimLaunchSuffix(m[1]));
      if (isValidCompanyName(candidate)) return candidate;
    }
  }
  return null;
}

// ── Resolve Slack user ID → display name (cached) ────────────
function resolveUser(userId, cache) {
  if (!userId) return "";
  if (cache[userId]) return cache[userId];

  try {
    var res  = UrlFetchApp.fetch(
      "https://slack.com/api/users.info?user=" + encodeURIComponent(userId),
      { headers: { "Authorization": "Bearer " + SLACK_TOKEN }, muteHttpExceptions: true }
    );
    var data = JSON.parse(res.getContentText());
    if (data.ok && data.user) {
      var name = data.user.profile.display_name || data.user.real_name || userId;
      cache[userId] = name;
      return name;
    }
  } catch(e) { /* ignore */ }

  cache[userId] = userId;
  return userId;
}
