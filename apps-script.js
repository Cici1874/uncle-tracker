// ══════════════════════════════════════════════════════
// Phil の 家务督察站 — Google Apps Script Backend
// ══════════════════════════════════════════════════════
// 部署步骤：
// 1. 创建一个新的 Google Sheet
// 2. 创建两个 sheet 页：
//    - "Records"（列：taskId, taskName, date, photoUrl, freq）
//    - "Flowers"（列：type, date）
// 3. 打开 Extensions > Apps Script
// 4. 粘贴这段代码
// 5. 点击 Deploy > New Deployment
//    - 类型选 "Web app"
//    - Execute as: Me
//    - Who has access: Anyone
// 6. 复制部署 URL，替换两个 HTML 文件里的 YOUR_APPS_SCRIPT_URL_HERE
//
// ⚠️ 重要：每次修改代码后必须创建 **新的 deployment version**
//    否则 live URL 还是旧代码！
// ══════════════════════════════════════════════════════

function doGet(request) {
  var action = request.parameter.action;

  // 设置 CORS headers
  var output;

  if (action === 'addRecord') {
    output = addRecord(request.parameter);
  } else if (action === 'addFlower') {
    output = addFlower(request.parameter);
  } else if (action === 'getRecords') {
    output = getRecords();
  } else if (action === 'getFlowers') {
    output = getFlowers();
  } else {
    output = ContentService.createTextOutput(
      JSON.stringify({ error: 'Unknown action' })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  return output;
}

// ── 写入打卡记录 ──
function addRecord(params) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Records');

  // 防止重复写入：检查同一个 taskId 在5分钟内是否已经有记录
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    var lastData = sheet.getRange(lastRow, 1, 1, 5).getValues()[0];
    var lastTaskId = lastData[0];
    var lastDate = new Date(lastData[2]);
    var now = new Date(params.date);
    if (lastTaskId === params.taskId && (now - lastDate) < 5 * 60 * 1000) {
      return ContentService.createTextOutput(
        JSON.stringify({ status: 'duplicate_skipped' })
      ).setMimeType(ContentService.MimeType.JSON);
    }
  }

  sheet.appendRow([
    params.taskId,
    params.taskName,
    params.date,
    params.photoUrl || '',
    params.freq,
  ]);

  // 设置文本格式防止日期被自动转换
  var newLastRow = sheet.getLastRow();
  sheet.getRange(newLastRow, 3).setNumberFormat('@'); // date column

  return ContentService.createTextOutput(
    JSON.stringify({ status: 'ok' })
  ).setMimeType(ContentService.MimeType.JSON);
}

// ── 写入鲜花记录 ──
function addFlower(params) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Flowers');

  sheet.appendRow([
    params.type,
    params.date,
  ]);

  var newLastRow = sheet.getLastRow();
  sheet.getRange(newLastRow, 2).setNumberFormat('@');

  return ContentService.createTextOutput(
    JSON.stringify({ status: 'ok' })
  ).setMimeType(ContentService.MimeType.JSON);
}

// ── 读取所有打卡记录 ──
function getRecords() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Records');
  var lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return ContentService.createTextOutput(
      JSON.stringify([])
    ).setMimeType(ContentService.MimeType.JSON);
  }

  var data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  var records = data.map(function(row) {
    return {
      taskId: row[0],
      taskName: row[1],
      date: formatDateValue(row[2]),
      photoUrl: row[3],
      freq: row[4],
    };
  });

  return ContentService.createTextOutput(
    JSON.stringify(records)
  ).setMimeType(ContentService.MimeType.JSON);
}

// ── 读取鲜花记录 ──
function getFlowers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Flowers');
  var lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return ContentService.createTextOutput(
      JSON.stringify([])
    ).setMimeType(ContentService.MimeType.JSON);
  }

  var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  var flowers = data.map(function(row) {
    return {
      type: row[0],
      date: formatDateValue(row[1]),
    };
  });

  return ContentService.createTextOutput(
    JSON.stringify(flowers)
  ).setMimeType(ContentService.MimeType.JSON);
}

// ── 日期格式化助手 ──
// Google Sheets 有时会把日期字符串自动转成 Date 对象
// 这个函数统一处理两种情况
function formatDateValue(val) {
  if (val instanceof Date) {
    return val.toISOString();
  }
  // 如果是字符串，原样返回
  return String(val);
}
