// JVALU ICV Tracker — Apps Script V4
// Updated columns: added expiryDate, responsible

var SHEET_NAME = "Clients";
var DOCS_SHEET = "Documents";

// ── Column order for the Clients sheet ──────────────────────────────────
var CLIENT_HEADERS = [
  "id", "company", "tradeLicense", "contact", "phone", "email",
  "financialYear", "expiryDate", "responsible", "stage", "addedDate", "notes",
  "prevICV", "prevIcvFy", "prevIcvScore", "hasEmirati", "icvScore"
];

function doGet(e) {
  var output;
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var action = e.parameter.action;

    if (action === "getAll") {
      output = JSON.stringify({ success: true, clients: getAllClients(ss) });
    }
    else if (action === "saveClient") {
      var client = JSON.parse(decodeURIComponent(e.parameter.client));
      saveClient(ss, client);
      output = JSON.stringify({ success: true });
    }
    else if (action === "saveDocs") {
      var clientId = decodeURIComponent(e.parameter.clientId);
      var documents = JSON.parse(decodeURIComponent(e.parameter.documents));
      saveDocs(ss, clientId, documents);
      output = JSON.stringify({ success: true });
    }
    else if (action === "deleteClient") {
      var id = decodeURIComponent(e.parameter.id);
      deleteClient(ss, id);
      output = JSON.stringify({ success: true });
    }
    else {
      output = JSON.stringify({ success: false, error: "Unknown action: " + action });
    }
  } catch(err) {
    output = JSON.stringify({ success: false, error: err.message + " | " + err.stack });
  }

  return ContentService.createTextOutput(output)
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Safely stringify a cell value — dates become yyyy-MM-dd so they work
//    directly in <input type="date"> fields on the front end ─────────────
function formatCellVal(v) {
  if (v === undefined || v === null) return "";
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v)) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(v);
}

// ── Get or create a sheet with the given headers ─────────────────────────
function getOrCreateSheet(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sh;
}

// ── Migrate existing Clients sheet — add any missing new columns ─────────
function migrateClientsSheet(sh) {
  var existing = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  CLIENT_HEADERS.forEach(function(col) {
    if (existing.indexOf(col) === -1) {
      var newCol = sh.getLastColumn() + 1;
      sh.getRange(1, newCol).setValue(col);
      existing.push(col); // keep local list in sync
    }
  });
}

// ── Read all clients + their documents ───────────────────────────────────
function getAllClients(ss) {
  var cSh = getOrCreateSheet(ss, SHEET_NAME, CLIENT_HEADERS);
  var dSh = getOrCreateSheet(ss, DOCS_SHEET, ["clientId","docId","received","date","remarks"]);

  migrateClientsSheet(cSh);

  var cData = cSh.getDataRange().getValues();
  var dData = dSh.getDataRange().getValues();

  if (cData.length <= 1) return [];

  var cHeaders = cData[0];
  var dHeaders = dData[0];

  // Build documents map keyed by clientId
  var docsMap = {};
  for (var i = 1; i < dData.length; i++) {
    var row = dData[i];
    var cid = String(row[dHeaders.indexOf("clientId")]);
    var did = parseInt(row[dHeaders.indexOf("docId")]);
    if (!docsMap[cid]) docsMap[cid] = {};
    var rv = row[dHeaders.indexOf("received")];
    docsMap[cid][did] = {
      received: rv === true || String(rv).toUpperCase() === "TRUE",
      date:     formatCellVal(row[dHeaders.indexOf("date")]),
      remarks:  row[dHeaders.indexOf("remarks")] ? String(row[dHeaders.indexOf("remarks")]) : ""
    };
  }

  // Build clients array
  var clients = [];
  for (var j = 1; j < cData.length; j++) {
    var crow = cData[j];
    if (!crow[0]) continue;
    var client = {};
    for (var k = 0; k < cHeaders.length; k++) {
      client[cHeaders[k]] = formatCellVal(crow[k]);
    }
    client.documents = docsMap[String(client.id)] || {};
    clients.push(client);
  }
  return clients;
}

// ── Save (insert or update) a single client ──────────────────────────────
function saveClient(ss, client) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sh = getOrCreateSheet(ss, SHEET_NAME, CLIENT_HEADERS);
    migrateClientsSheet(sh);

    // Read actual current headers (after migration)
    var currentHeaders = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];

    var data = sh.getDataRange().getValues();
    var rowIdx = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(client.id)) { rowIdx = i + 1; break; }
    }

    var row = currentHeaders.map(function(h) {
      return client[h] !== undefined ? client[h] : "";
    });

    if (rowIdx > 0) {
      sh.getRange(rowIdx, 1, 1, row.length).setValues([row]);
    } else {
      sh.appendRow(row);
    }
  } finally {
    lock.releaseLock();
  }
}

// ── Save documents (delete old rows, insert fresh) ────────────────────────
function saveDocs(ss, clientId, documents) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sh = getOrCreateSheet(ss, DOCS_SHEET, ["clientId","docId","received","date","remarks"]);
    var data = sh.getDataRange().getValues();

    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]) === String(clientId)) sh.deleteRow(i + 1);
    }

    var rows = [];
    var keys = Object.keys(documents);
    for (var j = 0; j < keys.length; j++) {
      var doc = documents[keys[j]];
      rows.push([
        String(clientId), parseInt(keys[j]),
        doc.received ? true : false,
        doc.date    || "",
        doc.remarks || ""
      ]);
    }
    if (rows.length > 0) {
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, 5).setValues(rows);
    }
  } finally {
    lock.releaseLock();
  }
}

// ── Delete a client and all their documents ───────────────────────────────
function deleteClient(ss, id) {
  var cSh = ss.getSheetByName(SHEET_NAME);
  var dSh = ss.getSheetByName(DOCS_SHEET);

  if (cSh) {
    var cData = cSh.getDataRange().getValues();
    for (var i = cData.length - 1; i >= 1; i--) {
      if (String(cData[i][0]) === String(id)) { cSh.deleteRow(i + 1); break; }
    }
  }
  if (dSh) {
    var dData = dSh.getDataRange().getValues();
    for (var i = dData.length - 1; i >= 1; i--) {
      if (String(dData[i][0]) === String(id)) dSh.deleteRow(i + 1);
    }
  }
}

// ══════════════════════════════════════════════════════════════════════════
// AUTOMATED BACKUPS
//
// ONE-TIME SETUP (do this once, manually, from the Apps Script editor):
//   1. Open this script in Apps Script (Extensions → Apps Script).
//   2. In the function dropdown at the top, select "setupWeeklyBackup".
//   3. Click "Run". The first run will ask you to authorize Drive access —
//      approve it (it's your own script, on your own spreadsheet).
//   4. Done. A "JVALU Backups" folder will appear in your Google Drive, and
//      a dated copy of the whole spreadsheet will be saved there every
//      Monday at ~2am automatically, going forward. The 12 most recent
//      weekly backups are kept; older ones are auto-deleted so Drive
//      doesn't fill up.
//
// You never need to touch this again after the one-time setup — it keeps
// running even through future code redeployments, since triggers are
// stored separately from the code itself.
// ══════════════════════════════════════════════════════════════════════════

var BACKUP_FOLDER_NAME = "JVALU Backups";
var BACKUPS_TO_KEEP = 12; // ~3 months of weekly backups

function setupWeeklyBackup() {
  // Remove any existing backup triggers first, so re-running this doesn't
  // create duplicates.
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "createWeeklyBackup") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger("createWeeklyBackup")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(2)
    .create();

  // Also create an immediate backup right now, so you have one straight away.
  createWeeklyBackup();

  Logger.log("Weekly backup trigger installed. A backup was also just created now.");
}

function createWeeklyBackup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceFile = DriveApp.getFileById(ss.getId());

  var folder = getOrCreateBackupFolder_();
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  var backupName = ss.getName() + " — backup " + stamp;

  sourceFile.makeCopy(backupName, folder);

  pruneOldBackups_(folder);
}

function getOrCreateBackupFolder_() {
  var folders = DriveApp.getFoldersByName(BACKUP_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(BACKUP_FOLDER_NAME);
}

function pruneOldBackups_(folder) {
  var files = [];
  var it = folder.getFiles();
  while (it.hasNext()) files.push(it.next());

  if (files.length <= BACKUPS_TO_KEEP) return;

  files.sort(function(a, b) {
    return b.getDateCreated().getTime() - a.getDateCreated().getTime();
  });

  for (var i = BACKUPS_TO_KEEP; i < files.length; i++) {
    files[i].setTrashed(true);
  }
}
