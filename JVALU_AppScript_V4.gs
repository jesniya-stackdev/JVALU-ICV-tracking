// JVALU ICV Tracker — Apps Script V4
// New fields: responsible, expiryDate, prevICV, prevIcvFy, prevIcvScore

var SHEET_NAME = "Clients";
var DOCS_SHEET = "Documents";

var CLIENT_HEADERS = [
  "id","company","tradeLicense","contact","phone","email",
  "responsible","financialYear","expiryDate","stage","addedDate","notes",
  "prevICV","prevIcvFy","prevIcvScore"
];

function doGet(e) {
  var output;
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var action = e.parameter.action;

    if (action === "getAll") {
      output = JSON.stringify({ success: true, clients: getAllClients(ss) });
    } else if (action === "saveClient") {
      saveClient(ss, JSON.parse(decodeURIComponent(e.parameter.client)));
      output = JSON.stringify({ success: true });
    } else if (action === "saveDocs") {
      saveDocs(ss, decodeURIComponent(e.parameter.clientId),
               JSON.parse(decodeURIComponent(e.parameter.documents)));
      output = JSON.stringify({ success: true });
    } else if (action === "deleteClient") {
      deleteClient(ss, decodeURIComponent(e.parameter.id));
      output = JSON.stringify({ success: true });
    } else {
      output = JSON.stringify({ success: false, error: "Unknown action: " + action });
    }
  } catch(err) {
    output = JSON.stringify({ success: false, error: err.message });
  }
  return ContentService.createTextOutput(output)
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sh;
}

function getAllClients(ss) {
  var cSh = getOrCreateSheet(ss, SHEET_NAME, CLIENT_HEADERS);
  var dSh = getOrCreateSheet(ss, DOCS_SHEET, ["clientId","docId","received","date","remarks"]);

  var cData = cSh.getDataRange().getValues();
  var dData = dSh.getDataRange().getValues();
  if (cData.length <= 1) return [];

  var cHeaders = cData[0];
  var dHeaders = dData[0];

  // Build docs map
  var docsMap = {};
  for (var i = 1; i < dData.length; i++) {
    var row = dData[i];
    var cid = String(row[dHeaders.indexOf("clientId")]);
    var did = parseInt(row[dHeaders.indexOf("docId")]);
    if (!docsMap[cid]) docsMap[cid] = {};
    var rv = row[dHeaders.indexOf("received")];
    docsMap[cid][did] = {
      received: rv === true || String(rv).toUpperCase() === "TRUE",
      date:    row[dHeaders.indexOf("date")]    ? String(row[dHeaders.indexOf("date")])    : "",
      remarks: row[dHeaders.indexOf("remarks")] ? String(row[dHeaders.indexOf("remarks")]) : ""
    };
  }

  var clients = [];
  for (var j = 1; j < cData.length; j++) {
    var crow = cData[j];
    if (!crow[0]) continue;
    var client = {};
    for (var k = 0; k < cHeaders.length; k++) {
      client[cHeaders[k]] = (crow[k] !== undefined && crow[k] !== null) ? String(crow[k]) : "";
    }
    client.prevICV = client.prevICV === "true" || client.prevICV === "TRUE";
    client.documents = docsMap[String(client.id)] || {};
    clients.push(client);
  }
  return clients;
}

function saveClient(ss, client) {
  var sh = getOrCreateSheet(ss, SHEET_NAME, CLIENT_HEADERS);
  var data = sh.getDataRange().getValues();
  var rowIdx = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(client.id)) { rowIdx = i + 1; break; }
  }
  var row = CLIENT_HEADERS.map(function(h) {
    return client[h] !== undefined ? client[h] : "";
  });
  if (rowIdx > 0) {
    sh.getRange(rowIdx, 1, 1, row.length).setValues([row]);
  } else {
    sh.appendRow(row);
  }
}

function saveDocs(ss, clientId, documents) {
  var sh = getOrCreateSheet(ss, DOCS_SHEET, ["clientId","docId","received","date","remarks"]);
  var data = sh.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(clientId)) sh.deleteRow(i + 1);
  }
  var rows = [];
  var keys = Object.keys(documents);
  for (var j = 0; j < keys.length; j++) {
    var doc = documents[keys[j]];
    rows.push([String(clientId), parseInt(keys[j]), doc.received ? true : false,
               doc.date || "", doc.remarks || ""]);
  }
  if (rows.length > 0) {
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, 5).setValues(rows);
  }
}

function deleteClient(ss, id) {
  var cSh = ss.getSheetByName(SHEET_NAME);
  var dSh = ss.getSheetByName(DOCS_SHEET);
  if (cSh) {
    var cd = cSh.getDataRange().getValues();
    for (var i = cd.length - 1; i >= 1; i--) {
      if (String(cd[i][0]) === String(id)) { cSh.deleteRow(i + 1); break; }
    }
  }
  if (dSh) {
    var dd = dSh.getDataRange().getValues();
    for (var i = dd.length - 1; i >= 1; i--) {
      if (String(dd[i][0]) === String(id)) dSh.deleteRow(i + 1);
    }
  }
}
