// Nhận chỉ số điện nước từ điện thoại -> ghi vào Google Sheet (tab ChiSo)
// Cách cài: xem HUONG_DAN.txt

var TAB_CS = 'ChiSo';
var TAB_CH = 'CauHinh';

function caiDat() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cs = ss.getSheetByName(TAB_CS) || ss.insertSheet(TAB_CS);
  if (cs.getLastRow() === 0) {
    cs.appendRow(['KY', 'SAP', 'LOAI', 'CHISOMOI', 'THOIGIAN', 'GHICHU', 'NGUOIGHI', 'MAGHI']);
    cs.setFrozenRows(1);
  }
  cs.getRange('A:A').setNumberFormat('@');   // giữ nguyên "12/2026", không bị đổi thành ngày
  cs.getRange('B:C').setNumberFormat('@');
  cs.getRange('E:E').setNumberFormat('@');
  var ch = ss.getSheetByName(TAB_CH) || ss.insertSheet(TAB_CH);
  if (ch.getRange('A1').getValue() === '') {
    ch.getRange('A1').setValue('Kỳ đang ghi (mm/yyyy)');
    ch.getRange('B1').setNumberFormat('@').setValue('12/2026');
    ch.getRange('A2').setValue('Đổi ô B1 khi sang kỳ mới. Mọi chỉ số gửi lên sẽ được gắn kỳ này.');
  }
}

function ky_() {
  var ch = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB_CH);
  return String(ch.getRange('B1').getDisplayValue()).trim();
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return json_({ ok: true, ky: ky_() });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var d = JSON.parse(e.postData.contents);
    var sap = String(d.sap || '').trim().toUpperCase();
    var loai = String(d.loai || '').trim().toUpperCase().charAt(0);
    var moi = String(d.moi == null ? '' : d.moi).trim().replace(',', '.');
    var id = String(d.id || '').trim();
    if (!/^[A-Z0-9\-]{1,12}$/.test(sap)) return json_({ ok: false, loi: 'Mã sạp không hợp lệ' });
    if (loai !== 'D' && loai !== 'N') return json_({ ok: false, loi: 'Loại phải là D hoặc N' });
    if (!/^\d{1,9}(\.\d{1,3})?$/.test(moi)) return json_({ ok: false, loi: 'Chỉ số không hợp lệ' });
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB_CS);
    var last = sh.getLastRow();
    if (id && last > 1) {                       // chống ghi trùng khi điện thoại gửi lại
      var from = Math.max(2, last - 1500);
      var ids = sh.getRange(from, 8, last - from + 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === id) return json_({ ok: true, trung: true, ky: ky_() });
    }
    var gc = String(d.ghichu || '').replace(/[\r\n,]+/g, ' ').substring(0, 200);
    var ng = String(d.nguoi || '').replace(/[\r\n,]+/g, ' ').substring(0, 40);
    var tg = Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm:ss');
    sh.appendRow([ky_(), sap, loai, moi, tg, gc, ng, id]);
    return json_({ ok: true, ky: ky_() });
  } catch (err) {
    return json_({ ok: false, loi: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}
