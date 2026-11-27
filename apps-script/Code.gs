/**
 * ระบบบันทึกโภชนบำบัดผู้ป่วยนอก — Google Apps Script
 * เก็บข้อมูลเป็นแถวใน Google Sheet (1 แถว = 1 ผู้ป่วย)
 * ผูกกับ Google Sheet: สร้างจาก Extensions > Apps Script ภายในชีตนั้น
 */
var SHEET_NAME = 'Records';
var HEADERS = ['id', 'วันที่', 'HN', 'ชื่อ-สกุล', 'แผนก', 'อายุ', 'เพศ', 'BMI', 'พลังงานเป้าหมาย (kcal)', 'บันทึกเมื่อ', 'json'];

function doGet(e) {
  // โหมดเชื่อมข้อมูลจากหน้าเว็บ GitHub Pages: ?action=records
  if (e && e.parameter && e.parameter.action === 'records') {
    var cb = e.parameter.callback;
    if (cb && /^[A-Za-z0-9_.]+$/.test(cb)) {
      return ContentService.createTextOutput(cb + '(' + getRecords() + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(getRecords()).setMimeType(ContentService.MimeType.JSON);
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ระบบบันทึกโภชนบำบัดผู้ป่วยนอก')
    .setFaviconUrl('https://raw.githubusercontent.com/piyanart502-dev/OPD/main/icons/favicon-32.png')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** รับข้อมูลจากหน้าเว็บ GitHub Pages (ส่งเป็น JSON ทั้งชุด) */
function doPost(e) {
  var out;
  try {
    saveRecords(e.postData.contents);
    out = { ok: true };
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#E6F4F1');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** ส่งข้อมูลทั้งหมดกลับไปให้หน้าเว็บ (เป็นข้อความ JSON) */
function getRecords() {
  var sh = getSheet_();
  var last = sh.getLastRow();
  if (last < 2) return '[]';
  var col = HEADERS.indexOf('json') + 1;
  var rows = sh.getRange(2, col, last - 1, 1).getValues();
  var out = [];
  rows.forEach(function (r) {
    if (r[0]) { try { out.push(JSON.parse(r[0])); } catch (e) {} }
  });
  return JSON.stringify(out);
}

/** รับข้อมูลทั้งหมดจากหน้าเว็บ แล้วเขียนทับลงชีต */
function saveRecords(jsonText) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var list = JSON.parse(jsonText || '[]');
    var sh = getSheet_();
    var last = sh.getLastRow();
    if (last > 1) sh.getRange(2, 1, last - 1, HEADERS.length).clearContent();
    if (list.length) {
      var rows = list.map(function (r) {
        return [
          r.id || '', r.date || '', "'" + (r.hn || ''), r.name || '', r.dept || '',
          r.age || '', r.gender || '',
          (r.weight && r.height) ? Math.round(r.weight / Math.pow(r.height / 100, 2) * 10) / 10 : '',
          r.energyTotal || '', r.savedAt || '', JSON.stringify(r)
        ];
      });
      sh.getRange(2, 1, rows.length, HEADERS.length).setValues(rows);
    }
    return true;
  } finally {
    lock.releaseLock();
  }
}
