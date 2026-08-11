// Google Apps Script backend for the badge-calculator customer list.
// Setup instructions: see SYNC-SETUP.md in the repo root.
//
// Data lives in a sheet called "Orders" (auto-created on first request).
// Order images are uploaded to a Drive folder called "BadgeCalcOrderImages"
// and only the resulting share link is stored in the sheet (image data URLs
// are far too large for a spreadsheet cell).

var SHEET_NAME = 'Orders';
var IMAGE_FOLDER_NAME = 'BadgeCalcOrderImages';
var HEADERS = ['id', 'seq', 'createdAt', 'branch', 'branchLabel', 'customerName', 'phone', 'address', 'items', 'total', 'imageUrl', 'status', 'updatedAt'];

function doGet(e) {
  return jsonResponse({ ok: true, message: 'Badge Calculator Orders API' });
}

function doPost(e) {
  var body = {};
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse({ error: 'invalid request body' });
  }

  var sheet = getSheet();
  var action = body.action;
  var result;
  try {
    if (action === 'list') result = listOrders(sheet);
    else if (action === 'create') result = createOrder(sheet, body.order || {});
    else if (action === 'update') result = updateOrder(sheet, body.id, body.patch || {});
    else if (action === 'delete') result = deleteOrder(sheet, body.id);
    else if (action === 'clearAll') result = clearAll(sheet);
    else result = { error: 'unknown action: ' + action };
  } catch (err) {
    result = { error: String(err) };
  }
  return jsonResponse(result);
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
  }
  return sheet;
}

function getImagesFolder() {
  var folders = DriveApp.getFoldersByName(IMAGE_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(IMAGE_FOLDER_NAME);
}

function saveImage(dataUrl, filename) {
  if (!dataUrl) return '';
  var match = String(dataUrl).match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.*)$/);
  if (!match) return '';
  var contentType = match[1];
  var bytes = Utilities.base64Decode(match[2]);
  var blob = Utilities.newBlob(bytes, contentType, filename);
  var file = getImagesFolder().createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/uc?export=view&id=' + file.getId();
}

function nextSeq() {
  var props = PropertiesService.getScriptProperties();
  var next = Number(props.getProperty('orderSeq') || '0') + 1;
  props.setProperty('orderSeq', String(next));
  return next;
}

function rowToOrder(row) {
  var obj = {};
  for (var i = 0; i < HEADERS.length; i++) obj[HEADERS[i]] = row[i];
  obj.items = obj.items ? JSON.parse(obj.items) : [];
  obj.total = Number(obj.total) || 0;
  obj.seq = Number(obj.seq) || 0;
  obj.image = obj.imageUrl || '';
  delete obj.imageUrl;
  return obj;
}

function listOrders(sheet) {
  var data = sheet.getDataRange().getValues();
  var rows = data.slice(1).filter(function (r) { return r[0]; });
  return { ok: true, orders: rows.map(rowToOrder) };
}

function findRowById(sheet, id) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === id) return i + 1; // 1-indexed sheet row
  }
  return -1;
}

function createOrder(sheet, order) {
  var seq = nextSeq();
  var id = order.id || Utilities.getUuid();
  var createdAt = new Date().toISOString();
  var imageUrl = saveImage(order.image, 'order-' + seq + '.png');
  sheet.appendRow([
    id, seq, createdAt,
    order.branch || '', order.branchLabel || '',
    order.customerName || '', order.phone || '', order.address || '',
    JSON.stringify(order.items || []), order.total || 0, imageUrl,
    order.status || 'in_progress', createdAt
  ]);
  return { ok: true, id: id, seq: seq, imageUrl: imageUrl };
}

function updateOrder(sheet, id, patch) {
  var rowNum = findRowById(sheet, id);
  if (rowNum === -1) return { error: 'not found' };
  for (var i = 0; i < HEADERS.length; i++) {
    var key = HEADERS[i];
    if (Object.prototype.hasOwnProperty.call(patch, key)) {
      var val = patch[key];
      if (key === 'items') val = JSON.stringify(val);
      sheet.getRange(rowNum, i + 1).setValue(val);
    }
  }
  sheet.getRange(rowNum, HEADERS.indexOf('updatedAt') + 1).setValue(new Date().toISOString());
  return { ok: true };
}

function deleteOrder(sheet, id) {
  var rowNum = findRowById(sheet, id);
  if (rowNum === -1) return { error: 'not found' };
  sheet.deleteRow(rowNum);
  return { ok: true };
}

function clearAll(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
  return { ok: true };
}
