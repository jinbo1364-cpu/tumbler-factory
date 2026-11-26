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
      else if (request.action === 'saveGeminiResult') data = saveGeminiResult(spreadsheet, request.record);
      else if (request.action === 'listGeminiResults') data = listGeminiResults(spreadsheet, request.inspection_id);
      else if (request.action === 'runGemini') data = runGemini(request);
      else if (request.action === 'setOnePortKey') data = setOnePortKey(request.key);
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
  ensureSheetHeaders(spreadsheet, 'gemini_results', [
    'result_id',
    'inspection_id',
    'created_at',
    'mode',
    'revision',
    'instruction_name',
    'instruction_text',
    'product_id',
    'weight_g',
    'images',
    'result',
  ]);
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

function saveGeminiResult(spreadsheet, record) {
  if (!record || typeof record !== 'object' || Array.isArray(record)) {
    throw createUserError('\uD310\uB3C5 \uACB0\uACFC\uB97C \uD655\uC778\uD558\uC138\uC694.');
  }
  if (typeof record.images !== 'string' || typeof record.result !== 'string') {
    throw createUserError('\uD310\uB3C5 \uACB0\uACFC\uC758 \uC0AC\uC9C4\uACFC \uD56D\uBAA9\uC740 JSON \uAE00\uC774\uC5B4\uC57C \uD569\uB2C8\uB2E4.');
  }
  try {
    JSON.parse(record.images);
    JSON.parse(record.result);
  } catch (error) {
    throw createUserError('\uD310\uB3C5 \uACB0\uACFC\uC758 \uC0AC\uC9C4\uACFC \uD56D\uBAA9\uC744 \uC77D\uC744 \uC218 \uC5C6\uC2B5\uB2C8\uB2E4.');
  }

  const appended = appendRows(spreadsheet, { gemini_results: [record] });
  return { saved: appended.added };
}

function listGeminiResults(spreadsheet, inspectionId) {
  if (inspectionId !== undefined && typeof inspectionId !== 'string') {
    throw createUserError('\uAC80\uC0AC\uBC88\uD638\uB97C \uD655\uC778\uD558\uC138\uC694.');
  }

  const sheet = spreadsheet.getSheetByName('gemini_results');
  if (!sheet) return { results: [] };
  const records = readSheetRows(sheet);
  return {
    results: inspectionId === undefined
      ? records
      : records.filter((record) => record.inspection_id === inspectionId),
  };
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

function runGemini(request) {
  if (
    !request ||
    typeof request.instruction !== 'string' ||
    typeof request.product_name !== 'string' ||
    typeof request.weight !== 'string' ||
    typeof request.revision !== 'string' ||
    typeof request.standard !== 'string' ||
    !Array.isArray(request.imageIds) ||
    request.imageIds.length === 0
  ) {
    throw createUserError('\uD310\uB3C5\uC5D0 \uD544\uC694\uD55C \uAC80\uC0AC \uC815\uBCF4\uC640 \uC0AC\uC9C4\uC744 \uD655\uC778\uD558\uC138\uC694.');
  }

  const apiKey = PropertiesService.getScriptProperties().getProperty('ONEPORT_API_KEY');
  if (!apiKey) throw createUserError('OnePort \uD0A4\uAC00 \uC5C6\uC2B5\uB2C8\uB2E4.');

  const content = [{
    type: 'text',
    text: [
      request.instruction,
      `\uAC80\uC0AC \uC815\uBCF4: \uC81C\uD488 ${request.product_name || '\uC81C\uD488 \uBBF8\uC9C0\uC815'}, \uC911\uB7C9 ${request.weight || '\uBBF8\uCE21\uC815'}, \uC801\uC6A9 \uAC80\uC0AC\uAE30\uC900 ${request.revision}.`,
      `\uAC80\uC0AC\uAE30\uC900 \uC6D0\uBB38:\n${request.standard}`,
      '\uC544\uB798 6\uAC1C \uD56D\uBAA9\uC758 JSON\uC73C\uB85C\uB9CC \uB2F5\uD574 \uC8FC\uC138\uC694: \uD310\uC815, \uBD88\uB7C9\uC720\uD615, \uAD00\uCC30\u00B7\uC218\uCE58 \uADFC\uAC70, \uC801\uC6A9 \uAE30\uC900\uACFC \uCD9C\uCC98, \uAD8C\uACE0 \uC870\uCE58, \uC0AC\uB78C \uC2B9\uC778 \uD544\uC694 \uC5EC\uBD80. \uD310\uC815\uC740 OK, NG, HOLD \uC911 \uD558\uB098\uB85C \uC791\uC131\uD574 \uC8FC\uC138\uC694.',
    ].join('\n\n'),
  }];
  let totalImageBytes = 0;

  request.imageIds.forEach((imageId) => {
    if (typeof imageId !== 'string' || imageId === '') {
      throw createUserError('\uB4F1\uB85D\uB41C \uC0AC\uC9C4\uC774 \uC5C6\uC2B5\uB2C8\uB2E4.');
    }

    let file;
    try {
      file = DriveApp.getFileById(imageId);
    } catch (error) {
      throw createUserError('\uB4F1\uB85D\uB41C \uC0AC\uC9C4\uC774 \uC5C6\uC2B5\uB2C8\uB2E4.');
    }
    const blob = file.getBlob();
    const bytes = blob.getBytes();
    const mimeType = blob.getContentType() || file.getMimeType();
    if (!mimeType || mimeType.indexOf('image/') !== 0) {
      throw createUserError('\uB4F1\uB85D\uB41C \uC0AC\uC9C4\uC774 \uC5C6\uC2B5\uB2C8\uB2E4.');
    }
    totalImageBytes += bytes.length;
    if (totalImageBytes > 3 * 1024 * 1024) {
      throw createUserError('\uC0AC\uC9C4\uC774 \uB108\uBB34 \uCEE4\uC11C \uD310\uB3C5\uD560 \uC218 \uC5C6\uC2B5\uB2C8\uB2E4.');
    }
    content.push({
      type: 'image_url',
      image_url: { url: `data:${mimeType};base64,${Utilities.base64Encode(bytes)}` },
    });
  });

  let response;
  try {
    response = UrlFetchApp.fetch('https://oneport.kr/v1/chat/completions', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: `Bearer ${apiKey}` },
      payload: JSON.stringify({
        model: 'gemini-3.8-flash',
        messages: [{ role: 'user', content: content }],
        response_format: { type: 'json_object' },
        max_tokens: 8192,
      }),
      muteHttpExceptions: true,
      timeoutSeconds: 120,
    });
  } catch (error) {
    throw createUserError('OnePort\uC5D0 \uC5F0\uACB0\uD558\uC9C0 \uBABB\uD588\uC2B5\uB2C8\uB2E4. 2\uBD84\uCBE4 \uB4A4 \uB2E4\uC2DC \uB20C\uB7EC \uC8FC\uC138\uC694.');
  }

  const statusCode = response.getResponseCode();
  let responseBody = {};
  try {
    responseBody = JSON.parse(response.getContentText());
  } catch (error) {
    if (statusCode < 200 || statusCode >= 300) {
      throw createUserError(getOnePortErrorMessage(statusCode, {}, apiKey));
    }
    throw createUserError('AI \uC751\uB2F5 \uD615\uC2DD\uC774 \uC62C\uBC14\uB974\uC9C0 \uC54A\uC544 \uD310\uB3C5\uD558\uC9C0 \uBABB\uD588\uC2B5\uB2C8\uB2E4.');
  }

  if (statusCode < 200 || statusCode >= 300) {
    throw createUserError(getOnePortErrorMessage(statusCode, responseBody, apiKey));
  }

  const contentText = responseBody && responseBody.choices && responseBody.choices[0] &&
    responseBody.choices[0].message && responseBody.choices[0].message.content;
  if (typeof contentText !== 'string') {
    throw createUserError('AI \uC751\uB2F5 \uD615\uC2DD\uC774 \uC62C\uBC14\uB974\uC9C0 \uC54A\uC544 \uD310\uB3C5\uD558\uC9C0 \uBABB\uD588\uC2B5\uB2C8\uB2E4.');
  }

  const jsonText = contentText.trim()
    .replace(/^```(?:json)?\s*/i, '')
    .replace(/\s*```$/, '')
    .trim();
  let parsedResult;
  try {
    parsedResult = JSON.parse(jsonText);
  } catch (error) {
    throw createUserError('AI \uC751\uB2F5 \uD615\uC2DD\uC774 \uC62C\uBC14\uB974\uC9C0 \uC54A\uC544 \uD310\uB3C5\uD558\uC9C0 \uBABB\uD588\uC2B5\uB2C8\uB2E4.');
  }

  const result = normalizeGeminiResult(parsedResult);
  if (!result) throw createUserError('AI \uC751\uB2F5 \uD615\uC2DD\uC774 \uC62C\uBC14\uB974\uC9C0 \uC54A\uC544 \uD310\uB3C5\uD558\uC9C0 \uBABB\uD588\uC2B5\uB2C8\uB2E4.');
  return result;
}

function normalizeGeminiResult(value) {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return null;
  const fields = [
    '\uD310\uC815',
    '\uBD88\uB7C9\uC720\uD615',
    '\uAD00\uCC30\u00B7\uC218\uCE58 \uADFC\uAC70',
    '\uC801\uC6A9 \uAE30\uC900\uACFC \uCD9C\uCC98',
    '\uAD8C\uACE0 \uC870\uCE58',
    '\uC0AC\uB78C \uC2B9\uC778 \uD544\uC694 \uC5EC\uBD80',
  ];
  const result = {};
  for (let index = 0; index < fields.length; index += 1) {
    const field = fields[index];
    if (typeof value[field] !== 'string') return null;
    result[field] = value[field];
  }
  if (!['OK', 'NG', 'HOLD'].includes(result['\uD310\uC815'])) return null;
  return result;
}

function getOnePortErrorMessage(statusCode, responseBody, apiKey) {
  const error = responseBody && responseBody.error && typeof responseBody.error === 'object'
    ? responseBody.error
    : {};
  const safeMessage = sanitizeOnePortText(error.message || '\uC694\uCCAD\uC5D0 \uC2E4\uD328\uD588\uC2B5\uB2C8\uB2E4.', apiKey);
  if (statusCode === 401 && error.reason === 'key_revoked') return '\uC774 OnePort \uD0A4\uB294 \uAFBC\uC838 \uC788\uC2B5\uB2C8\uB2E4(\uCF54\uB4DC 401). \uB2E4\uB978 \uD0A4\uB85C \uBC14\uAFD4 \uC8FC\uC138\uC694.';
  if (statusCode === 401) return 'OnePort \uD0A4\uAC00 \uB9DE\uC9C0 \uC54A\uC2B5\uB2C8\uB2E4(\uCF54\uB4DC 401). \uD0A4\uB97C \uB2E4\uC2DC \uB123\uC5B4 \uC8FC\uC138\uC694.';
  if (statusCode === 402) return 'OnePort \uC794\uC561\uC774 \uBD80\uC871\uD574 \uD310\uB3C5\uD558\uC9C0 \uBABB\uD588\uC2B5\uB2C8\uB2E4(\uCF54\uB4DC 402). \uB2E4\uB978 \uD0A4\uB85C \uBC14\uAFD4 \uC8FC\uC138\uC694.';
  if (statusCode === 403) return '\uC774 OnePort \uD0A4\uB85C\uB294 \uD310\uB3C5 \uBAA8\uB378\uC744 \uC4F8 \uC218 \uC5C6\uC2B5\uB2C8\uB2E4(\uCF54\uB4DC 403). \uB2E4\uB978 \uD0A4\uB85C \uBC14\uAFD4 \uC8FC\uC138\uC694.';
  if (statusCode === 404) return 'OnePort\uC5D0\uC11C \uD310\uB3C5 \uBAA8\uB378\uC744 \uCC3E\uC9C0 \uBABB\uD588\uC2B5\uB2C8\uB2E4(\uCF54\uB4DC 404). \uBAA8\uB378 \uC774\uB984\uC744 \uD655\uC778\uD574 \uC8FC\uC138\uC694.';
  if (statusCode === 413) return '\uC0AC\uC9C4\uC774 \uB108\uBB34 \uCEE4\uC11C \uD310\uB3C5\uD560 \uC218 \uC5C6\uC2B5\uB2C8\uB2E4(\uCF54\uB4DC 413).';
  if (statusCode === 429 && (error.code === 'key_limit_exceeded' || error.type === 'key_limit_exceeded')) {
    return 'OnePort \uD0A4\uC758 \uC0AC\uC6A9 \uD55C\uB3C4\uB97C \uB2E4 \uC37C\uC2B5\uB2C8\uB2E4(\uCF54\uB4DC 429). \uB2E4\uB978 \uD0A4\uB85C \uBC14\uAFD4 \uC8FC\uC138\uC694.';
  }
  if ([429, 500, 502, 503, 504].includes(statusCode)) {
    return `AI \uD310\uB3C5 \uC11C\uBC84\uAC00 \uBD90\uBE55\uB2C8\uB2E4(\uCF54\uB4DC ${statusCode}). 2\uBD84\uCBE4 \uB4A4 \uB2E4\uC2DC \uB20C\uB7EC \uC8FC\uC138\uC694.`;
  }
  return `OnePort \uC624\uB958(\uCF54\uB4DC ${statusCode}): ${safeMessage}`;
}

function sanitizeOnePortText(value, apiKey) {
  let text = String(value);
  if (apiKey) text = text.split(apiKey).join('[\uAC00\uB9BC]');
  return text.replace(/Bearer\s+[^\s"'`,}]+/gi, 'Bearer [\uAC00\uB9BC]');
}

function setOnePortKey(key) {
  if (typeof key !== 'string' || key === '') {
    throw createUserError('OnePort \uD0A4\uAC00 .env\uC5D0 \uC5C6\uC2B5\uB2C8\uB2E4.');
  }
  PropertiesService.getScriptProperties().setProperty('ONEPORT_API_KEY', key);
  return { saved: true };
}

function jsonOutput(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}