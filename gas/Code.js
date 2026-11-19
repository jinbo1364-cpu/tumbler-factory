function doGet() {
  return ContentService.createTextOutput('\uAD8C\uD55C \uC2B9\uC778\uC774 \uB05D\uB0AC\uC2B5\uB2C8\uB2E4. \uC774 \uCC3D\uC744 \uB2EB\uACE0 \uD130\uBBF8\uB110\uB85C \uB3CC\uC544\uAC00\uC138\uC694.');
}

function doPost(event) {
  try {
    const request = JSON.parse(event.postData.contents);
    let data;
    if (request.action === 'setup') {
      data = setupStorage();
    } else {
      const spreadsheet = getConfiguredSpreadsheet();
      if (request.action === 'read') data = readSheets(spreadsheet, request.tabs);
      else if (request.action === 'append') data = appendRows(spreadsheet, request.rows);
      else if (request.action === 'update') data = updateRow(spreadsheet, request);
      else if (request.action === 'savePhoto') data = savePhoto(request);
      else if (request.action === 'getPhoto') data = getPhoto(request.id);
      else throw new Error('\uC9C0\uC6D0\uD558\uC9C0 \uC54A\uB294 \uC77C\uC785\uB2C8\uB2E4.');
    }
    return jsonOutput({ ok: true, data: data });
  } catch (error) {
    const message = error && error.userMessage
      ? error.userMessage
      : '\uC54C \uC218 \uC5C6\uB294 \uC624\uB958\uAC00 \uB0AC\uC2B5\uB2C8\uB2E4.';
    return jsonOutput({ ok: false, error: message });
  }
}

function setupStorage() {
  const properties = PropertiesService.getScriptProperties();
  const spreadsheet = getOrCreateSpreadsheet(properties);
  const folder = getOrCreatePhotoFolder(properties);

  ensureSheetHeaders(spreadsheet, 'inspections', [
    'inspection_id',
    'product_id',
    'inspected_at',
    'revision',
    'weight_g',
    'memo',
    'decision',
    'content_version',
  ]);
  ensureSheetHeaders(spreadsheet, 'images', ['inspection_id', 'position', 'image_url']);
  removeEmptyDefaultSheet(spreadsheet);

  return { sheetUrl: spreadsheet.getUrl(), folderUrl: folder.getUrl() };
}

function getOrCreateSpreadsheet(properties) {
  const rememberedId = properties.getProperty('SHEET_ID');
  if (rememberedId) {
    try {
      return SpreadsheetApp.openById(rememberedId);
    } catch (error) {
      // Recreate the spreadsheet if the remembered file no longer exists.
    }
  }

  const spreadsheet = SpreadsheetApp.create('\uD140\uBE14\uB7EC \uACF5\uC7A5 \uD488\uC9C8\uAD00\uB9AC \uB370\uC774\uD130\uBCA0\uC774\uC2A4');
  properties.setProperty('SHEET_ID', spreadsheet.getId());
  return spreadsheet;
}

function getOrCreatePhotoFolder(properties) {
  const rememberedId = properties.getProperty('FOLDER_ID');
  if (rememberedId) {
    try {
      return DriveApp.getFolderById(rememberedId);
    } catch (error) {
      // Recreate the folder if the remembered folder no longer exists.
    }
  }

  const folder = DriveApp.createFolder('\uD140\uBE14\uB7EC \uACF5\uC7A5 \uD488\uC9C8\uAC80\uC0AC \uC0AC\uC9C4');
  properties.setProperty('FOLDER_ID', folder.getId());
  return folder;
}

function ensureSheetHeaders(spreadsheet, name, requiredHeaders) {
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) sheet = spreadsheet.insertSheet(name);

  const lastColumn = sheet.getLastColumn();
  const existingHeaders = lastColumn > 0
    ? sheet.getRange(1, 1, 1, lastColumn).getValues()[0]
    : [];
  const missingHeaders = requiredHeaders.filter((header) => !existingHeaders.includes(header));
  if (missingHeaders.length > 0) {
    sheet.getRange(1, lastColumn + 1, 1, missingHeaders.length).setValues([missingHeaders]);
  }
}

function removeEmptyDefaultSheet(spreadsheet) {
  spreadsheet.getSheets().forEach((sheet) => {
    const isDefaultSheet = sheet.getName() === 'Sheet1' || sheet.getName() === '\uC2DC\uD2B8\uD2B8';
    const isEmpty = sheet.getLastRow() === 0 && sheet.getLastColumn() === 0;
    if (isDefaultSheet && isEmpty) spreadsheet.deleteSheet(sheet);
  });
}

function getConfiguredSpreadsheet() {
  const properties = PropertiesService.getScriptProperties();
  const spreadsheetId = properties.getProperty('SHEET_ID');
  const folderId = properties.getProperty('FOLDER_ID');
  if (!spreadsheetId || !folderId) {
    throw createUserError('\uC2DC\uD2B8\uC640 \uC0AC\uC9C4 \uD3F4\uB354\uAC00 \uC544\uC9C1 \uC5C6\uC2B5\uB2C8\uB2E4. npm.cmd run gas:init\uC744 \uC2E4\uD589\uD558\uC138\uC694.');
  }

  try {
    DriveApp.getFolderById(folderId);
    return SpreadsheetApp.openById(spreadsheetId);
  } catch (error) {
    throw createUserError('\uC2DC\uD2B8\uC640 \uC0AC\uC9C4 \uD3F4\uB354\uAC00 \uC544\uC9C1 \uC5C6\uC2B5\uB2C8\uB2E4. npm.cmd run gas:init\uC744 \uC2E4\uD589\uD558\uC138\uC694.');
  }
}

function readSheets(spreadsheet, tabNames) {
  if (!Array.isArray(tabNames)) throw new Error('\uD0ED \uC774\uB984 \uBAA9\uB85D\uC774 \uD544\uC694\uD569\uB2C8\uB2E4.');
  const result = {};
  tabNames.forEach((tabName) => {
    const sheet = spreadsheet.getSheetByName(tabName);
    result[tabName] = sheet ? readSheetRows(sheet) : [];
  });
  return result;
}

function readSheetRows(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn === 0) return [];

  const values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
  const headers = values[0];
  return values.slice(1)
    .filter((row) => row.some((value) => value !== '' && value !== null))
    .map((row) => {
      const item = {};
      headers.forEach((header, index) => {
        if (header !== '' && header !== null) item[header] = readCellValue(row[index]);
      });
      return item;
    });
}

function appendRows(spreadsheet, rowsByTab) {
  if (!rowsByTab || typeof rowsByTab !== 'object' || Array.isArray(rowsByTab)) {
    throw new Error('\uB354\uD560 \uC904\uC744 \uD655\uC778\uD558\uC138\uC694.');
  }

  let added = 0;
  Object.keys(rowsByTab).forEach((tabName) => {
    const rows = rowsByTab[tabName];
    if (!Array.isArray(rows)) throw new Error('\uB354\uD560 \uC904\uC744 \uD655\uC778\uD558\uC138\uC694.');
    const sheet = spreadsheet.getSheetByName(tabName);
    if (!sheet) throw new Error('\uC2DC\uD2B8 \uD0ED\uC744 \uCC3E\uC744 \uC218 \uC5C6\uC2B5\uB2C8\uB2E4.');
    if (rows.length === 0) return;

    const headers = getSheetHeaders(sheet);
    if (headers.length === 0) throw new Error('\uC2DC\uD2B8 \uBA38\uB9AC\uAE00\uC744 \uCC3E\uC744 \uC218 \uC5C6\uC2B5\uB2C8\uB2E4.');
    const values = rows.map((row) => headers.map((header) => toSheetValue(row[header])));
    sheet.getRange(sheet.getLastRow() + 1, 1, values.length, headers.length).setValues(values);
    added += values.length;
  });
  return { added: added };
}

function updateRow(spreadsheet, request) {
  if (!request || typeof request.tab !== 'string') {
    throw createUserError('\uC218\uC815\uD560 \uAE30\uB85D\uC774 \uC5C6\uC2B5\uB2C8\uB2E4.');
  }
  const sheet = spreadsheet.getSheetByName(request.tab);
  if (!sheet || typeof request.key !== 'string' || !request.changes || typeof request.changes !== 'object' || Array.isArray(request.changes)) {
    throw createUserError('\uC218\uC815\uD560 \uAE30\uB85D\uC774 \uC5C6\uC2B5\uB2C8\uB2E4.');
  }

  const headers = getSheetHeaders(sheet);
  const keyColumn = headers.indexOf(request.key);
  if (keyColumn < 0 || sheet.getLastRow() < 2) throw createUserError('\uC218\uC815\uD560 \uAE30\uB85D\uC774 \uC5C6\uC2B5\uB2C8\uB2E4.');

  const keyValues = sheet.getRange(2, keyColumn + 1, sheet.getLastRow() - 1, 1).getValues();
  const matchIndex = keyValues.findIndex((row) => valuesEqual(readCellValue(row[0]), request.value));
  if (matchIndex < 0) throw createUserError('\uC218\uC815\uD560 \uAE30\uB85D\uC774 \uC5C6\uC2B5\uB2C8\uB2E4.');

  const rowNumber = matchIndex + 2;
  Object.keys(request.changes).forEach((header) => {
    const column = headers.indexOf(header);
    if (column >= 0) sheet.getRange(rowNumber, column + 1).setValue(toSheetValue(request.changes[header]));
  });
  return { updated: 1 };
}

function getSheetHeaders(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn === 0) return [];
  return sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
}

function toSheetValue(value) {
  if (value === null || value === undefined || value === '') return '';
  if (typeof value === 'number') return value;
  return "'" + String(value);
}

function readCellValue(value) {
  if (value instanceof Date) return Utilities.formatDate(value, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
  if (typeof value === 'string' && value.charAt(0) === "'") return value.slice(1);
  return value;
}

function valuesEqual(left, right) {
  if (left instanceof Date) left = readCellValue(left);
  if (right instanceof Date) right = readCellValue(right);
  return left === right;
}

function createUserError(message) {
  const error = new Error(message);
  error.userMessage = message;
  return error;
}

function savePhoto(request) {
  if (typeof request.name !== 'string' || typeof request.mimeType !== 'string' || typeof request.base64 !== 'string') {
    throw new Error('\uC0AC\uC9C4 \uC815\uBCF4\uB97C \uD655\uC778\uD558\uC138\uC694.');
  }

  const properties = PropertiesService.getScriptProperties();
  const folder = DriveApp.getFolderById(properties.getProperty('FOLDER_ID'));
  const bytes = Utilities.base64Decode(request.base64);
  const blob = Utilities.newBlob(bytes, request.mimeType, request.name);
  const file = folder.createFile(blob);
  return { id: file.getId() };
}

function getPhoto(fileId) {
  if (typeof fileId !== 'string' || fileId === '') {
    throw createUserError('\uB4F1\uB85D\uB41C \uC0AC\uC9C4\uC774 \uC5C6\uC2B5\uB2C8\uB2E4.');
  }

  const properties = PropertiesService.getScriptProperties();
  const folder = DriveApp.getFolderById(properties.getProperty('FOLDER_ID'));
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    if (file.getId() === fileId) {
      return {
        mimeType: file.getMimeType(),
        base64: Utilities.base64Encode(file.getBlob().getBytes()),
      };
    }
  }
  throw createUserError('\uB4F1\uB85D\uB41C \uC0AC\uC9C4\uC774 \uC5C6\uC2B5\uB2C8\uB2E4.');
}

function jsonOutput(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}