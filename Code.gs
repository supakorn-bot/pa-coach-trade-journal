/**
 * PA Coach Web App v3.5 - Google Sheet Sync + Drive Images
 * Workflow: GitHub Pages index.html -> Apps Script Web App -> Google Sheet + Google Drive
 * v3.5 จุดสำคัญ: setupDatabase ไม่ล้างข้อมูลเดิม, upsert กัน ID ซ้ำ, รองรับ headers เพิ่ม, เก็บรูปลง Drive
 */

const PA_COACH_CONFIG = {
  sheetName: 'Trades',
  folderName: 'PA_Coach_Trade_Images',
  makeDriveImagesPublic: true,
};

const TRADE_HEADERS = [
  'id', 'updatedAt', 'createdAt', 'tradeDate', 'tradeTime', 'symbol', 'market',
  'direction', 'status', 'strategyName', 'entry', 'sl', 'tp1', 'tp2', 'tp3',
  'maxRR', 'resultR', 'decision', 'json'
];

function setupDatabase() {
  const ss = getSpreadsheet_();
  const sheet = getTradeSheet_();
  ensureHeaders_(sheet);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, Math.min(TRADE_HEADERS.length, sheet.getLastColumn()));
  const folder = getImageFolder_();
  return jsonLog_({
    ok: true,
    message: 'setupDatabase completed - existing data kept',
    version: 'v3.5',
    spreadsheetUrl: ss.getUrl(),
    sheetName: sheet.getName(),
    folderUrl: folder.getUrl(),
  });
}

function resetDatabaseDanger() {
  const sheet = getTradeSheet_();
  sheet.clear();
  sheet.getRange(1, 1, 1, TRADE_HEADERS.length).setValues([TRADE_HEADERS]);
  sheet.setFrozenRows(1);
  return jsonLog_({ ok: true, message: 'resetDatabaseDanger completed', version: 'v3.5' });
}

function doGet(e) {
  return jsonOut_({ ok: true, message: 'PA Coach Sync API is running', version: 'v3.5' });
}

function doPost(e) {
  try {
    const raw = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
    const req = JSON.parse(raw);
    validateSecret_(req.secret);

    const action = req.action || '';
    if (action === 'ping') return jsonOut_({ ok: true, message: 'pong', version: 'v3.5', now: new Date().toISOString() });
    if (action === 'listTrades') {
      const trades = listTrades_();
      return jsonOut_({ ok: true, count: trades.length, trades, version: 'v3.5' });
    }
    if (action === 'upsertTrades') {
      const trades = Array.isArray(req.trades) ? req.trades : [];
      const mode = String(req.mode || 'upsert').toLowerCase();
      const saved = upsertTrades_(trades, mode);
      return jsonOut_({ ok: true, count: saved.length, trades: saved, version: 'v3.5' });
    }
    if (action === 'deleteTrade') {
      deleteTrade_(req.id);
      return jsonOut_({ ok: true, message: 'deleted', id: req.id, version: 'v3.5' });
    }
    if (action === 'telegramAlert') {
      return handleTelegramAlert_(req);
    }
    if (action === 'clearTrades') {
      clearTrades_();
      return jsonOut_({ ok: true, message: 'cleared', version: 'v3.5' });
    }
    if (action === 'backupSnapshot') {
      const file = backupSnapshot_();
      return jsonOut_({ ok: true, fileUrl: file.getUrl(), version: 'v3.5' });
    }
    throw new Error('Unknown action: ' + action);
  } catch (err) {
    return jsonOut_({ ok: false, error: err.message || String(err), version: 'v3.5' });
  }
}

function getSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('PA_COACH_SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) throw new Error('ไม่พบ Spreadsheet: ให้เปิด Apps Script จาก Google Sheet หรือใส่ PA_COACH_SPREADSHEET_ID ใน Script properties');
  return active;
}

function getTradeSheet_() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(PA_COACH_CONFIG.sheetName);
  if (!sheet) sheet = ss.insertSheet(PA_COACH_CONFIG.sheetName);
  ensureHeaders_(sheet);
  return sheet;
}

function ensureHeaders_(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const firstRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0].filter(String);
  if (!firstRow.length || firstRow[0] !== 'id') {
    sheet.getRange(1, 1, 1, TRADE_HEADERS.length).setValues([TRADE_HEADERS]);
    sheet.setFrozenRows(1);
    return;
  }
  const existing = new Set(firstRow);
  const missing = TRADE_HEADERS.filter(h => !existing.has(h));
  if (missing.length) {
    sheet.getRange(1, firstRow.length + 1, 1, missing.length).setValues([missing]);
  }
  sheet.setFrozenRows(1);
}

function getHeaderMap_(sheet) {
  ensureHeaders_(sheet);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => { if (h) map[String(h)] = i; });
  return { headers, map };
}

function getImageFolder_() {
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty('PA_COACH_DRIVE_FOLDER_ID');
  if (savedId) {
    try { return DriveApp.getFolderById(savedId); } catch (e) {}
  }
  const folders = DriveApp.getFoldersByName(PA_COACH_CONFIG.folderName);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(PA_COACH_CONFIG.folderName);
  props.setProperty('PA_COACH_DRIVE_FOLDER_ID', folder.getId());
  return folder;
}

function validateSecret_(incomingSecret) {
  const props = PropertiesService.getScriptProperties();
  const expected = props.getProperty('PA_COACH_SECRET');
  if (!expected) throw new Error('ยังไม่ได้ตั้งค่า PA_COACH_SECRET ใน Script properties');
  if (String(incomingSecret || '') !== expected) throw new Error('API Secret ไม่ถูกต้อง');
}

function listTrades_() {
  const sheet = getTradeSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const { headers, map } = getHeaderMap_(sheet);
  const values = sheet.getRange(2, 1, last - 1, headers.length).getValues();
  const trades = values.map(row => rowToTrade_(row, map)).filter(Boolean);
  return dedupeTrades_(trades);
}

function rowToTrade_(row, map) {
  const jsonIndex = map.json;
  if (jsonIndex !== undefined && row[jsonIndex]) {
    try { return JSON.parse(row[jsonIndex]); } catch (e) {}
  }
  const t = {};
  Object.keys(map).forEach(k => t[k] = row[map[k]]);
  return t.id ? t : null;
}

function upsertTrades_(inputTrades, mode) {
  const sheet = getTradeSheet_();
  ensureHeaders_(sheet);
  if (mode === 'replace') clearTrades_();

  const trades = dedupeTrades_((inputTrades || [])
    .filter(t => t && typeof t === 'object')
    .map((t, i) => normalizeTrade_(t, i)));

  const map = getExistingRowMap_();
  const saved = [];
  trades.forEach(trade => {
    const processed = processTradeImages_(trade);
    const row = tradeToRow_(processed);
    const rowNumber = map[processed.id];
    if (rowNumber) {
      sheet.getRange(rowNumber, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }
    saved.push(processed);
  });
  consolidateDuplicateRows_();
  return dedupeTrades_(saved);
}

function clearTrades_() {
  const sheet = getTradeSheet_();
  const last = sheet.getLastRow();
  if (last >= 2) sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).clearContent();
}

function deleteTrade_(id) {
  if (!id) throw new Error('missing id');
  const rows = getRowsById_(id);
  rows.reverse().forEach(r => getTradeSheet_().deleteRow(r));
}

function getExistingRowMap_() {
  const sheet = getTradeSheet_();
  const last = sheet.getLastRow();
  const map = {};
  if (last < 2) return map;
  const ids = sheet.getRange(2, 1, last - 1, 1).getValues().flat();
  ids.forEach((id, i) => { if (id && !map[String(id)]) map[String(id)] = i + 2; });
  return map;
}

function getRowsById_(id) {
  const sheet = getTradeSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const ids = sheet.getRange(2, 1, last - 1, 1).getValues().flat();
  const rows = [];
  ids.forEach((x, i) => { if (String(x) === String(id)) rows.push(i + 2); });
  return rows;
}

function consolidateDuplicateRows_() {
  const sheet = getTradeSheet_();
  const last = sheet.getLastRow();
  if (last < 3) return;
  const ids = sheet.getRange(2, 1, last - 1, 1).getValues().flat();
  const seen = {};
  const deleteRows = [];
  ids.forEach((id, i) => {
    if (!id) return;
    const rowNum = i + 2;
    if (seen[id]) deleteRows.push(rowNum);
    else seen[id] = rowNum;
  });
  deleteRows.reverse().forEach(r => sheet.deleteRow(r));
}

function normalizeTrade_(trade, i) {
  const copy = JSON.parse(JSON.stringify(trade));
  copy.id = copy.id || Utilities.getUuid();
  copy.updatedAt = new Date().toISOString();
  copy.createdAt = copy.createdAt || copy.updatedAt;
  copy.symbol = copy.symbol || 'XAUUSD';
  copy.appVersion = copy.appVersion || 'v3.5';
  return copy;
}

function tradeToRow_(trade) {
  const sheet = getTradeSheet_();
  const { headers } = getHeaderMap_(sheet);
  const base = {
    id: trade.id || '',
    updatedAt: trade.updatedAt || trade.createdAt || '',
    createdAt: trade.createdAt || '',
    tradeDate: trade.tradeDate || '',
    tradeTime: trade.tradeTime || '',
    symbol: trade.symbol || '',
    market: trade.market || '',
    direction: trade.direction || '',
    status: trade.status || '',
    strategyName: trade.strategyName || trade.strategy || '',
    entry: trade.entry === null || trade.entry === undefined ? '' : trade.entry,
    sl: trade.sl === null || trade.sl === undefined ? '' : trade.sl,
    tp1: trade.tp1 === null || trade.tp1 === undefined ? '' : trade.tp1,
    tp2: trade.tp2 === null || trade.tp2 === undefined ? '' : trade.tp2,
    tp3: trade.tp3 === null || trade.tp3 === undefined ? '' : trade.tp3,
    maxRR: trade.maxRR === null || trade.maxRR === undefined ? '' : trade.maxRR,
    resultR: trade.resultR === null || trade.resultR === undefined ? '' : trade.resultR,
    decision: trade.decision || '',
    json: JSON.stringify(trade),
  };
  return headers.map(h => base[h] !== undefined ? base[h] : '');
}

function dedupeTrades_(trades) {
  const out = {};
  (trades || []).forEach(t => {
    if (!t) return;
    const id = String(t.id || Utilities.getUuid());
    t.id = id;
    const old = out[id];
    if (!old || String(t.updatedAt || t.createdAt || '') > String(old.updatedAt || old.createdAt || '')) out[id] = t;
  });
  return Object.keys(out).map(k => out[k]).sort((a,b) => String(b.updatedAt || b.createdAt || '').localeCompare(String(a.updatedAt || a.createdAt || '')));
}

function processTradeImages_(trade) {
  const t = JSON.parse(JSON.stringify(trade));
  ['chartW1', 'chartD1', 'chartH4', 'chartM15'].forEach(field => {
    if (t[field]) t[field] = saveDataUrlImageIfNeeded_(t[field], t, field);
  });
  if (Array.isArray(t.resultImages)) {
    t.resultImages = t.resultImages.map((src, i) => saveDataUrlImageIfNeeded_(src, t, 'result_' + (i + 1)));
  }
  return t;
}

function saveDataUrlImageIfNeeded_(src, trade, label) {
  if (!src || typeof src !== 'string') return src;
  if (!src.startsWith('data:image/')) return src;
  const match = src.match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/);
  if (!match) return src;
  const mime = match[1];
  const base64 = match[2];
  const ext = mime.indexOf('png') > -1 ? 'png' : (mime.indexOf('webp') > -1 ? 'webp' : 'jpg');
  const bytes = Utilities.base64Decode(base64);
  const safeSymbol = String(trade.symbol || 'TRADE').replace(/[^a-zA-Z0-9_-]/g, '_');
  const safeLabel = String(label || 'chart').replace(/[^a-zA-Z0-9_-]/g, '_');
  const filename = safeSymbol + '_' + trade.id + '_' + safeLabel + '_' + Date.now() + '.' + ext;
  const blob = Utilities.newBlob(bytes, mime, filename);
  const file = getImageFolder_().createFile(blob);
  if (PA_COACH_CONFIG.makeDriveImagesPublic) {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  }
  return 'https://drive.google.com/uc?export=view&id=' + file.getId();
}

function backupSnapshot_() {
  const trades = listTrades_();
  const blob = Utilities.newBlob(JSON.stringify({ version: 'v3.5', exportedAt: new Date().toISOString(), trades }, null, 2), 'application/json', 'pa_coach_backup_' + Date.now() + '.json');
  return getImageFolder_().createFile(blob);
}


function htmlEscape_(value) {
  return String(value === null || value === undefined ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function sendTelegramMessage_(text) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('TELEGRAM_BOT_TOKEN');
  const chatId = props.getProperty('TELEGRAM_CHAT_ID');

  if (!token) throw new Error('ยังไม่ได้ตั้งค่า TELEGRAM_BOT_TOKEN ใน Script properties');
  if (!chatId) throw new Error('ยังไม่ได้ตั้งค่า TELEGRAM_CHAT_ID ใน Script properties');

  const url = 'https://api.telegram.org/bot' + token + '/sendMessage';
  const payload = {
    chat_id: chatId,
    text: String(text || '').slice(0, 3900),
    parse_mode: 'HTML',
    disable_web_page_preview: true,
  };

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    payload,
    muteHttpExceptions: true,
  });

  const code = res.getResponseCode();
  const body = res.getContentText();

  if (code < 200 || code >= 300) {
    throw new Error('Telegram ส่งไม่สำเร็จ: ' + body);
  }

  try {
    return JSON.parse(body);
  } catch (e) {
    return { raw: body };
  }
}

function buildTelegramAlertText_(trade) {
  const t = trade || {};

  const blocks = Array.isArray(t.blocks) && t.blocks.length
    ? t.blocks.map(x => '• ' + htmlEscape_(x)).join('\n')
    : '-';

  const supportZone = htmlEscape_(t.supportZone || '-');
  const resistanceZone = htmlEscape_(t.resistanceZone || '-');

  return [
    '🚨 <b>PA Coach Alert</b>',
    '',
    '<b>Symbol:</b> ' + htmlEscape_(t.symbol || '-'),
    '<b>Market:</b> ' + htmlEscape_(t.market || '-'),
    '<b>Direction:</b> ' + htmlEscape_(t.direction || 'WAIT'),
    '<b>Status:</b> ' + htmlEscape_(t.status || 'PLAN'),
    '<b>Strategy:</b> ' + htmlEscape_(t.strategyName || t.strategy || '-'),
    '',
    '<b>Decision:</b> ' + htmlEscape_(t.decision || '-'),
    '',
    '<b>Entry:</b> ' + htmlEscape_(t.entry ?? '-'),
    '<b>SL:</b> ' + htmlEscape_(t.sl ?? '-'),
    '<b>TP1:</b> ' + htmlEscape_(t.tp1 ?? '-'),
    '<b>TP2:</b> ' + htmlEscape_(t.tp2 ?? '-'),
    '<b>TP3:</b> ' + htmlEscape_(t.tp3 ?? '-'),
    '<b>Max RR:</b> ' + htmlEscape_(t.maxRR ?? '-'),
    '',
    '<b>Support Zone:</b> ' + supportZone,
    '<b>Resistance Zone:</b> ' + resistanceZone,
    '',
    '<b>เหตุผลเข้า / ไม่เข้า:</b>',
    htmlEscape_(t.entryReason || '-'),
    '',
    '<b>สิ่งที่ยังบล็อกอยู่:</b>',
    blocks,
    '',
    '<b>Risk Note:</b>',
    htmlEscape_(t.riskNote || '-'),
    '',
    '⚠️ ไม่ใช่สัญญาณการันตีกำไร ให้เช็กแท่งปิด ข่าว SL และ RR ก่อนเข้าออเดอร์'
  ].join('\n');
}

function handleTelegramAlert_(req) {
  const text = req.text || buildTelegramAlertText_(req.trade || {});
  if (!text) throw new Error('ไม่มีข้อความสำหรับส่ง Telegram');

  const telegramResult = sendTelegramMessage_(text);

  return jsonOut_({
    ok: true,
    message: 'telegram alert sent',
    telegram: telegramResult,
    version: 'v3.5',
  });
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function jsonLog_(obj) {
  Logger.log(JSON.stringify(obj, null, 2));
  return obj;
}
