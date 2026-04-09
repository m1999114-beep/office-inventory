// ============================================
// 庫存管理系統 API
// ============================================

var SPREADSHEET_ID = '1f7th15c2hTuzUdHLWrjH8cfDMhN3Crhc-bpQlDgUCZ8';
var STOCK_SHEET_NAME = 'stock';
var LOG_SHEET_NAME = 'logs';
var GEMINI_MODEL = 'gemini-1.5-flash';
var SCRIPT_LOCK_WAIT_MS = 30000;

var COLUMNS = {
  id: 0,
  name: 1,
  spec: 2,
  unit: 3,
  warehouseQty: 4,
  siteQty: 5,
  safeQty: 6
};

function doGet(e) {
  try {
    var stock = getStockData();
    return jsonResponse({
      status: 'success',
      stock: stock,
      count: stock.length,
      timestamp: new Date().toISOString()
    });
  } catch (err) {
    return jsonResponse({
      status: 'error',
      msg: err.message
    });
  }
}

function doPost(e) {
  try {
    var payload = parsePostPayload_(e);
    var action = normalizeRequiredString_(payload.action, 'action');
    var person = normalizeOptionalString_(payload.person, '未填');
    var result;

    if (action === 'warehouse_audit') {
      result = warehouseAudit(payload.id, payload.qty, person);
    } else if (action === 'audit') {
      result = siteAudit(payload.id, payload.qty, person);
    } else if (action === 'move_to_site') {
      result = moveToSite(payload.id, payload.qty, person);
    } else if (action === 'gemini_ocr') {
      result = callGeminiAPI(payload.image);
    } else if (action === 'gemini_parse_text') {
      result = parseTextRemoveWithAI(payload.text, payload.stockList);
    } else {
      result = {
        status: 'error',
        msg: '未知的操作: ' + action
      };
    }

    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({
      status: 'error',
      msg: err.message
    });
  }
}

function getStockData() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getRequiredSheet_(ss, STOCK_SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  var stock = [];
  var i;

  for (i = 1; i < data.length; i++) {
    if (isStockRowEmpty_(data[i])) continue;
    stock.push(mapStockRow_(data[i]));
  }

  return stock;
}

function warehouseAudit(id, newQty, person) {
  var normalizedId = normalizeRequiredString_(id, 'id');
  var normalizedQty = normalizeNonNegativeQty_(newQty, 'qty');

  return withStockLock_(function (ss, sheet, data) {
    var item = findStockItemById_(data, normalizedId);
    if (!item) return { status: 'error', msg: '找不到物品 ID: ' + normalizedId };

    sheet.getRange(item.rowIndex, COLUMNS.warehouseQty + 1).setValue(normalizedQty);

    logAction_(ss, {
      timestamp: new Date(),
      action: '倉庫盤點',
      itemId: item.id,
      itemName: item.name,
      oldQty: item.warehouseQty,
      newQty: normalizedQty,
      diff: roundQty_(normalizedQty - item.warehouseQty),
      person: person
    });

    return {
      status: 'success',
      msg: '倉庫盤點完成',
      oldQty: item.warehouseQty,
      newQty: normalizedQty
    };
  });
}

function siteAudit(id, newQty, person) {
  var normalizedId = normalizeRequiredString_(id, 'id');
  var normalizedQty = normalizeNonNegativeQty_(newQty, 'qty');

  return withStockLock_(function (ss, sheet, data) {
    var item = findStockItemById_(data, normalizedId);
    if (!item) return { status: 'error', msg: '找不到物品 ID: ' + normalizedId };

    sheet.getRange(item.rowIndex, COLUMNS.siteQty + 1).setValue(normalizedQty);

    logAction_(ss, {
      timestamp: new Date(),
      action: '工地盤點',
      itemId: item.id,
      itemName: item.name,
      oldQty: item.siteQty,
      newQty: normalizedQty,
      diff: roundQty_(normalizedQty - item.siteQty),
      person: person
    });

    return {
      status: 'success',
      msg: '工地盤點完成',
      oldQty: item.siteQty,
      newQty: normalizedQty
    };
  });
}

function moveToSite(id, moveQty, person) {
  var normalizedId = normalizeRequiredString_(id, 'id');
  var normalizedQty = normalizePositiveQty_(moveQty, 'qty');

  return withStockLock_(function (ss, sheet, data) {
    var item = findStockItemById_(data, normalizedId);
    if (!item) return { status: 'error', msg: '找不到物品 ID: ' + normalizedId };
    if (normalizedQty > item.warehouseQty) {
      return {
        status: 'error',
        msg: '倉庫數量不足，目前只有 ' + item.warehouseQty
      };
    }

    var newWarehouseQty = roundQty_(item.warehouseQty - normalizedQty);
    var newSiteQty = roundQty_(item.siteQty + normalizedQty);

    sheet.getRange(item.rowIndex, COLUMNS.warehouseQty + 1, 1, 2)
      .setValues([[newWarehouseQty, newSiteQty]]);

    logAction_(ss, {
      timestamp: new Date(),
      action: '移至工地',
      itemId: item.id,
      itemName: item.name,
      oldQty: item.warehouseQty + '→' + item.siteQty,
      newQty: newWarehouseQty + '→' + newSiteQty,
      diff: normalizedQty,
      person: person
    });

    return {
      status: 'success',
      msg: '已移動 ' + normalizedQty + ' 至工地',
      warehouse: { old: item.warehouseQty, new: newWarehouseQty },
      site: { old: item.siteQty, new: newSiteQty }
    };
  });
}

function parseTextRemoveWithAI(text, stockList) {
  var inputText = normalizeRequiredString_(text, 'text');
  var normalizedStockList = normalizeStockList_(stockList);

  if (!normalizedStockList.length) {
    normalizedStockList = getStockData().map(function (item) {
      return {
        id: normalizeOptionalString_(item.id, ''),
        name: normalizeOptionalString_(item.name, ''),
        spec: normalizeOptionalString_(item.spec, '')
      };
    });
  }

  try {
    var prompt = [
      '你是庫存清單配對助手。',
      '請根據 input 與 stockList，逐行找出最適合的單一物料。',
      '只能使用 stockList 內既有資料作為 matchedId、matchedName、matchedSpec。',
      '找不到時 matchedId、matchedName、matchedSpec 都請填空字串。',
      'qty 只保留數字，可為小數。',
      '回傳純 JSON 陣列，格式為：',
      '[{"originalInput":"原始輸入","qty":1,"matchedId":"ID","matchedName":"品名","matchedSpec":"規格"}]',
      '請勿回傳 markdown 或多餘說明。',
      'stockList=' + JSON.stringify(normalizedStockList),
      'input=' + inputText
    ].join('\n');

    var responseText = runGeminiRequest_([{ text: prompt }]);
    var parsed = normalizeTextMatchResults_(parseJsonValue_(responseText), normalizedStockList);
    return {
      status: 'success',
      items: parsed,
      source: 'gemini'
    };
  } catch (err) {
    Logger.log('gemini_parse_text fallback: ' + err.message);
    return {
      status: 'success',
      items: fallbackParseTextRemove_(inputText, normalizedStockList),
      source: 'fallback'
    };
  }
}

function callGeminiAPI(base64Image) {
  try {
    var imageData = parseBase64Image_(base64Image);
    var prompt = [
      '你是一個專業的物料入庫助手。',
      '請辨識圖片中的表格內容，找出所有的品名、規格和數量。',
      '回傳純 JSON 陣列，格式為：',
      '[{"name":"品名","spec":"規格","qty":數字}]',
      '若看不清楚或沒有內容，回傳 []。',
      '請勿回傳 markdown。'
    ].join('\n');

    var responseText = runGeminiRequest_([
      { text: prompt },
      { inline_data: { mime_type: imageData.mimeType, data: imageData.data } }
    ]);

    return {
      status: 'success',
      items: normalizeOcrItems_(parseJsonValue_(responseText))
    };
  } catch (err) {
    return {
      status: 'error',
      msg: 'AI Request Failed: ' + err.message
    };
  }
}

function logAction_(ss, log) {
  try {
    var logSheet = ss.getSheetByName(LOG_SHEET_NAME);
    if (!logSheet) {
      logSheet = ss.insertSheet(LOG_SHEET_NAME);
      logSheet.appendRow([
        '時間', '操作', '物品ID', '品名',
        '原數量', '新數量', '差異', '操作人員'
      ]);
      logSheet.getRange(1, 1, 1, 8)
        .setFontWeight('bold')
        .setBackground('#e0e0e0');
    }

    logSheet.appendRow([
      log.timestamp,
      log.action,
      log.itemId,
      log.itemName,
      log.oldQty,
      log.newQty,
      log.diff,
      log.person
    ]);
  } catch (err) {
    Logger.log('記錄失敗: ' + err.message);
  }
}

function withStockLock_(callback) {
  var lock = LockService.getScriptLock();
  lock.waitLock(SCRIPT_LOCK_WAIT_MS);
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = getRequiredSheet_(ss, STOCK_SHEET_NAME);
    var data = sheet.getDataRange().getValues();
    return callback(ss, sheet, data);
  } finally {
    lock.releaseLock();
  }
}

function findStockItemById_(data, id) {
  var targetId = String(id);
  var i;

  for (i = 1; i < data.length; i++) {
    if (String(data[i][COLUMNS.id]) === targetId) {
      return {
        rowIndex: i + 1,
        id: data[i][COLUMNS.id],
        name: normalizeOptionalString_(data[i][COLUMNS.name], ''),
        warehouseQty: roundQty_(data[i][COLUMNS.warehouseQty]),
        siteQty: roundQty_(data[i][COLUMNS.siteQty])
      };
    }
  }

  return null;
}

function getRequiredSheet_(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('找不到工作表: ' + sheetName);
  return sheet;
}

function parsePostPayload_(e) {
  if (!e || !e.postData || !e.postData.contents) {
    throw new Error('缺少請求內容');
  }

  try {
    return JSON.parse(e.postData.contents);
  } catch (err) {
    throw new Error('無效的 JSON 格式');
  }
}

function parseBase64Image_(base64Image) {
  var raw = normalizeRequiredString_(base64Image, 'image');
  var match = raw.match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/);

  if (match) {
    return {
      mimeType: match[1],
      data: match[2]
    };
  }

  return {
    mimeType: 'image/jpeg',
    data: raw
  };
}

function runGeminiRequest_(parts) {
  var apiKey = getGeminiApiKey_();
  var apiUrl = 'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL + ':generateContent?key=' + apiKey;
  var payload = {
    contents: [{ parts: parts }],
    generationConfig: {
      responseMimeType: 'application/json',
      temperature: 0.1
    }
  };
  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  var response = UrlFetchApp.fetch(apiUrl, options);
  var json = JSON.parse(response.getContentText());

  if (json.error) throw new Error(json.error.message);
  if (!json.candidates || !json.candidates.length) throw new Error('AI 沒有回傳結果');
  if (!json.candidates[0].content || !json.candidates[0].content.parts) throw new Error('AI 回傳內容為空');

  var text = '';
  var responseParts = json.candidates[0].content.parts;
  var i;

  for (i = 0; i < responseParts.length; i++) {
    if (responseParts[i].text) text += responseParts[i].text;
  }

  if (!text) throw new Error('AI 回傳內容為空');
  return text;
}

function getGeminiApiKey_() {
  var key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) {
    throw new Error('尚未設定 GEMINI_API_KEY，請到 Apps Script 的 Script Properties 新增');
  }
  return key;
}

function normalizeOcrItems_(value) {
  var items = asArray_(value);
  var normalized = [];
  var i;

  for (i = 0; i < items.length; i++) {
    normalized.push({
      name: normalizeOptionalString_(items[i].name, ''),
      spec: normalizeOptionalString_(items[i].spec, ''),
      qty: roundQty_(items[i].qty)
    });
  }

  return normalized;
}

function normalizeTextMatchResults_(value, stockList) {
  var items = asArray_(value);
  var normalized = [];
  var i;

  for (i = 0; i < items.length; i++) {
    var item = items[i] || {};
    var matched = null;
    if (item.matchedId) matched = findStockListById_(stockList, item.matchedId);
    if (!matched && item.matchedName) matched = findStockListByNameSpec_(stockList, item.matchedName, item.matchedSpec);

    normalized.push({
      originalInput: normalizeOptionalString_(item.originalInput, ''),
      qty: normalizeParsedQty_(item.qty),
      matchedId: matched ? matched.id : normalizeOptionalString_(item.matchedId, ''),
      matchedName: matched ? matched.name : normalizeOptionalString_(item.matchedName, ''),
      matchedSpec: matched ? matched.spec : normalizeOptionalString_(item.matchedSpec, '')
    });
  }

  return normalized;
}

function fallbackParseTextRemove_(text, stockList) {
  var lines = String(text).split(/\r?\n/);
  var results = [];
  var i;

  for (i = 0; i < lines.length; i++) {
    var rawLine = normalizeOptionalString_(lines[i], '');
    if (!rawLine) continue;

    var parsedLine = parseTextRemoveLine_(rawLine);
    var match = findBestStockMatch_(parsedLine.itemText, stockList);

    results.push({
      originalInput: rawLine,
      qty: parsedLine.qty,
      matchedId: match ? match.id : '',
      matchedName: match ? match.name : '',
      matchedSpec: match ? match.spec : ''
    });
  }

  return results;
}

function parseTextRemoveLine_(line) {
  var cleaned = normalizeOptionalString_(line, '').replace(/^(?:[-\u2022]\s*|\d+\.\s+)/, '');
  var qty = 1;
  var itemText = cleaned;
  var qtyMatch = cleaned.match(/(?:[\*xX×])\s*(\d+(?:\.\d+)?)\s*$/);

  if (!qtyMatch) qtyMatch = cleaned.match(/\s+(\d+(?:\.\d+)?)\s*$/);

  if (qtyMatch) {
    qty = normalizeParsedQty_(qtyMatch[1]);
    itemText = cleaned.substring(0, qtyMatch.index);
  }

  return {
    itemText: normalizeOptionalString_(itemText, ''),
    qty: qty
  };
}

function findBestStockMatch_(rawText, stockList) {
  var query = normalizeMatchText_(rawText);
  var queryTokens = tokenizeMatchText_(rawText);
  var best = null;
  var bestScore = 0;
  var i;

  if (!query) return null;

  for (i = 0; i < stockList.length; i++) {
    var item = stockList[i];
    var name = normalizeMatchText_(item.name);
    var spec = normalizeMatchText_(item.spec);
    var combined = name + spec;
    var score = 0;
    var j;

    if (!combined) continue;
    if (query === combined || query === name || (spec && query === spec)) score += 100;
    if (query.indexOf(name) >= 0 || name.indexOf(query) >= 0) score += 60;
    if (spec && (query.indexOf(spec) >= 0 || spec.indexOf(query) >= 0)) score += 40;

    for (j = 0; j < queryTokens.length; j++) {
      if (queryTokens[j].length < 2) continue;
      if (combined.indexOf(queryTokens[j]) >= 0) score += 12;
    }

    if (score > bestScore) {
      bestScore = score;
      best = item;
    }
  }

  return bestScore >= 40 ? best : null;
}

function normalizeStockList_(stockList) {
  if (!stockList || Object.prototype.toString.call(stockList) !== '[object Array]') return [];

  var normalized = [];
  var i;

  for (i = 0; i < stockList.length; i++) {
    normalized.push({
      id: normalizeOptionalString_(stockList[i].id, ''),
      name: normalizeOptionalString_(stockList[i].name, ''),
      spec: normalizeOptionalString_(stockList[i].spec, '')
    });
  }

  return normalized;
}

function findStockListById_(stockList, id) {
  var target = normalizeOptionalString_(id, '');
  var i;

  for (i = 0; i < stockList.length; i++) {
    if (normalizeOptionalString_(stockList[i].id, '') === target) return stockList[i];
  }

  return null;
}

function findStockListByNameSpec_(stockList, name, spec) {
  var targetName = normalizeOptionalString_(name, '');
  var targetSpec = normalizeOptionalString_(spec, '');
  var i;

  for (i = 0; i < stockList.length; i++) {
    if (normalizeOptionalString_(stockList[i].name, '') === targetName &&
        normalizeOptionalString_(stockList[i].spec, '') === targetSpec) {
      return stockList[i];
    }
  }

  return null;
}

function normalizeMatchText_(text) {
  return normalizeOptionalString_(text, '')
    .toLowerCase()
    .replace(/[（()）［］\[\]{}]/g, '')
    .replace(/[\s\-_.,，。+]/g, '')
    .replace(/[＊*x×]/g, '');
}

function tokenizeMatchText_(text) {
  return normalizeOptionalString_(text, '')
    .toLowerCase()
    .replace(/[（()）［］\[\]{}]/g, ' ')
    .split(/[\s\-_.,，。+*xX×]+/)
    .filter(function (token) { return !!token; });
}

function mapStockRow_(row) {
  return {
    id: row[COLUMNS.id],
    name: normalizeOptionalString_(row[COLUMNS.name], ''),
    spec: normalizeOptionalString_(row[COLUMNS.spec], ''),
    unit: normalizeOptionalString_(row[COLUMNS.unit], '個'),
    warehouseQty: roundQty_(row[COLUMNS.warehouseQty]),
    siteQty: roundQty_(row[COLUMNS.siteQty]),
    safeQty: roundQty_(row[COLUMNS.safeQty])
  };
}

function isStockRowEmpty_(row) {
  return !row[COLUMNS.id] && !row[COLUMNS.name];
}

function parseJsonValue_(text) {
  var cleaned = normalizeOptionalString_(text, '')
    .replace(/```json/g, '')
    .replace(/```/g, '')
    .trim();

  try {
    return JSON.parse(cleaned);
  } catch (err) {
    var start = cleaned.indexOf('[');
    var end = cleaned.lastIndexOf(']');
    if (start >= 0 && end > start) return JSON.parse(cleaned.substring(start, end + 1));

    start = cleaned.indexOf('{');
    end = cleaned.lastIndexOf('}');
    if (start >= 0 && end > start) return JSON.parse(cleaned.substring(start, end + 1));

    throw new Error('AI JSON 解析失敗');
  }
}

function asArray_(value) {
  if (Object.prototype.toString.call(value) === '[object Array]') return value;
  if (value && Object.prototype.toString.call(value.items) === '[object Array]') return value.items;
  throw new Error('AI 回傳格式不正確');
}

function normalizeRequiredString_(value, fieldName) {
  var text = normalizeOptionalString_(value, '');
  if (!text) throw new Error('缺少必要參數 (' + fieldName + ')');
  return text;
}

function normalizeOptionalString_(value, fallback) {
  if (value === undefined || value === null) return fallback || '';
  var text = String(value).trim();
  return text === '' ? (fallback || '') : text;
}

function normalizeNonNegativeQty_(value, fieldName) {
  var qty = Number(value);
  if (!isFinite(qty)) throw new Error(fieldName + ' 必須是數字');
  if (qty < 0) throw new Error(fieldName + ' 不能小於 0');
  return roundQty_(qty);
}

function normalizePositiveQty_(value, fieldName) {
  var qty = normalizeNonNegativeQty_(value, fieldName);
  if (qty <= 0) throw new Error(fieldName + ' 必須大於 0');
  return qty;
}

function normalizeParsedQty_(value) {
  var qty = Number(value);
  if (!isFinite(qty) || qty <= 0) return 1;
  return roundQty_(qty);
}

function roundQty_(value) {
  var qty = Number(value);
  if (!isFinite(qty)) return 0;
  return Math.round(qty * 1000) / 1000;
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function testGetStock() {
  var result = doGet({});
  Logger.log(result.getContent());
}
