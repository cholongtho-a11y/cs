// Ghi chỉ số điện nước chợ Long Thọ - máy chủ nhận số (Google Apps Script) - PHIÊN BẢN 2
// Cài đặt: xem HUONG_DAN_V2.txt
//
// Các tab trong Google Sheet:
//   ChiSo      : số liệu điện thoại gửi lên (chỉ máy chủ ghi)
//   DauKy      : chỉ số cũ đầu kỳ (dán từ file Excel xuất ra, 1 lần mỗi kỳ)
//   NhanVien   : TEN | PIN | HOATDONG (Có/Không)
//   CauHinh    : B1 = kỳ ghi (tháng tiêu thụ, mm/yyyy), B2 = MỞ/ĐÓNG, B3 = khoá cho Excel tự tải, B4 = ngưỡng cảnh báo tăng (lần)

var TAB_CS = 'ChiSo', TAB_DK = 'DauKy', TAB_NV = 'NhanVien', TAB_CH = 'CauHinh';
var CS_HEAD = ['KY', 'SAP', 'LOAI', 'CHISOMOI', 'THOIGIAN', 'GHICHU', 'NGUOIGHI', 'MAGHI', 'TRANGTHAI', 'QUAYVONG'];
var DK_HEAD = ['KY', 'SAP', 'LOAI', 'MAKH', 'TENKH', 'KHUVUC', 'CHISOCU', 'TB3', 'DACO'];

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function tab_(n) { return ss_().getSheetByName(n); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

// ---------- CÀI ĐẶT (chạy 1 lần) ----------
function caiDat() {
  var s = ss_();
  var cs = tab_(TAB_CS) || s.insertSheet(TAB_CS);
  if (cs.getLastRow() === 0) cs.appendRow(CS_HEAD); else cs.getRange(1, 1, 1, CS_HEAD.length).setValues([CS_HEAD]);
  cs.setFrozenRows(1);
  cs.getRange('A:C').setNumberFormat('@'); cs.getRange('E:E').setNumberFormat('@'); cs.getRange('H:J').setNumberFormat('@');
  var dk = tab_(TAB_DK) || s.insertSheet(TAB_DK);
  if (dk.getLastRow() === 0) dk.appendRow(DK_HEAD);
  dk.setFrozenRows(1);
  dk.getRange('A:C').setNumberFormat('@');
  var nv = tab_(TAB_NV) || s.insertSheet(TAB_NV);
  if (nv.getLastRow() === 0) { nv.appendRow(['TEN', 'PIN', 'HOATDONG']); nv.getRange('B:B').setNumberFormat('@'); }
  nv.setFrozenRows(1);
  var ch = tab_(TAB_CH) || s.insertSheet(TAB_CH);
  var lab = [['Kỳ đang ghi (tháng tiêu thụ, mm/yyyy)'], ['Trạng thái đợt ghi (MỞ / ĐÓNG)'], ['Khoá cho Excel tự tải (không đưa cho nhân viên)'], ['Ngưỡng cảnh báo tăng đột biến (lần so với TB 3 kỳ)']];
  ch.getRange('A1:A4').setValues(lab);
  ch.getRange('B1:B4').setNumberFormat('@');
  if (ch.getRange('B1').getDisplayValue() === '') ch.getRange('B1').setValue('11/2026');
  ch.getRange('B2').setValue('ĐÓNG');
  if (ch.getRange('B3').getDisplayValue() === '') ch.getRange('B3').setValue(taoMa_(24));
  if (ch.getRange('B4').getDisplayValue() === '') ch.getRange('B4').setValue('2');
  ch.getRange('A6').setValue('Ghi chú: đổi B1 sang kỳ mới; B2 chỉ nhập đúng chữ MỞ khi bắt đầu đợt ghi và ĐÓNG khi kết thúc.');
  ch.setColumnWidth(1, 380);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Chỉ số điện nước')
    .addItem('Mở đợt ghi', 'moDot').addItem('Đóng đợt ghi', 'dongDot')
    .addSeparator().addItem('Tạo PIN cho nhân viên chưa có PIN', 'taoPin')
    .addItem('Đổi khoá Excel tự tải', 'doiKhoaExcel')
    .addItem('Cài đặt lần đầu / sửa cấu trúc các tab', 'caiDat').addToUi();
}
function moDot() { tab_(TAB_CH).getRange('B2').setValue('MỞ'); SpreadsheetApp.getActive().toast('Đã MỞ đợt ghi kỳ ' + kyHienTai_()); }
function dongDot() { tab_(TAB_CH).getRange('B2').setValue('ĐÓNG'); SpreadsheetApp.getActive().toast('Đã ĐÓNG đợt ghi'); }
function doiKhoaExcel() { tab_(TAB_CH).getRange('B3').setValue(taoMa_(24)); SpreadsheetApp.getActive().toast('Đã đổi khoá. Cập nhật khoá mới trong phần mềm Excel (CAI_DAT: GC_AUTO_KHOA).'); }
function taoPin() {
  var nv = tab_(TAB_NV), n = nv.getLastRow(), dung = {}, i, p;
  if (n < 2) return;
  var v = nv.getRange(2, 1, n - 1, 3).getValues();
  for (i = 0; i < v.length; i++) if (String(v[i][1]).trim() !== '') dung[String(v[i][1]).trim()] = true;
  for (i = 0; i < v.length; i++) {
    if (String(v[i][0]).trim() !== '' && String(v[i][1]).trim() === '') {
      do { p = String(Math.floor(100000 + Math.random() * 900000)); } while (dung[p]);
      dung[p] = true; v[i][1] = p; if (String(v[i][2]).trim() === '') v[i][2] = 'Có';
    }
  }
  nv.getRange(2, 1, v.length, 3).setValues(v);
}
function taoMa_(n) { var c = 'abcdefghijkmnpqrstuvwxyz23456789', s = '', i; for (i = 0; i < n; i++) s += c.charAt(Math.floor(Math.random() * c.length)); return s; }

// ---------- TIỆN ÍCH ----------
function cau_(o) { var ch = tab_(TAB_CH); return String(ch.getRange(o).getDisplayValue()).trim(); }
function kyHienTai_() { return cau_('B1'); }
function dotMo_() { return cau_('B2').toUpperCase().replace('Ở', 'O') === 'MO'; }
function nguong_() { var x = parseFloat(String(cau_('B4')).replace(',', '.')); return (x > 0) ? x : 2; }
function chuanSap_(s) { return String(s == null ? '' : s).toUpperCase().replace(/\s+/g, '').replace(/[^A-Z0-9\-]/g, '').substring(0, 14); }
function chuanLoai_(l) { l = String(l == null ? '' : l).trim().toUpperCase(); if (l === 'Đ') l = 'D'; return l.charAt(0); }
function so_(x) { var t = String(x == null ? '' : x).trim().replace(',', '.'); return /^\d{1,9}(\.\d{1,3})?$/.test(t) ? parseFloat(t) : null; }

function nhanVien_(pin) {
  pin = String(pin == null ? '' : pin).trim();
  if (!/^\d{4,8}$/.test(pin)) return null;
  var nv = tab_(TAB_NV), n = nv.getLastRow(), v, i;
  if (n < 2) return null;
  v = nv.getRange(2, 1, n - 1, 3).getDisplayValues();
  for (i = 0; i < v.length; i++) {
    if (String(v[i][1]).trim() === pin && String(v[i][2]).trim().toLowerCase() !== 'không' && String(v[i][2]).trim().toLowerCase() !== 'khong') return String(v[i][0]).trim();
  }
  return null;
}

function dauKy_() {
  var dk = tab_(TAB_DK), n = dk.getLastRow(), v, i, r, out = [], kyDK = '';
  if (n < 2) return { ds: out, ky: '' };
  v = dk.getRange(2, 1, n - 1, 9).getDisplayValues();
  for (i = 0; i < v.length; i++) {
    r = v[i];
    if (String(r[1]).trim() === '') continue;
    if (!kyDK) kyDK = String(r[0]).trim();
    out.push({ sap: chuanSap_(r[1]), loai: chuanLoai_(r[2]), makh: r[3], ten: r[4], kv: r[5], cu: r[6] === '' ? null : parseFloat(String(r[6]).replace(',', '.')), tb3: r[7] === '' ? 0 : parseFloat(String(r[7]).replace(',', '.')) });
  }
  return { ds: out, ky: kyDK };
}

function daGhi_(ky) {
  var cs = tab_(TAB_CS), n = cs.getLastRow(), v, i, m = {};
  if (n < 2) return m;
  v = cs.getRange(2, 1, n - 1, 9).getDisplayValues();
  for (i = 0; i < v.length; i++) if (String(v[i][0]).trim() === ky && String(v[i][8]).trim() !== 'CHUA_KHOP') m[chuanSap_(v[i][1]) + '|' + chuanLoai_(v[i][2])] = v[i][3];
  return m;
}

// Chống dò PIN: 10 lần sai trong 10 phút thì khoá 10 phút
function khoa_() { var c = CacheService.getScriptCache(); return parseInt(c.get('sai') || '0', 10) >= 10; }
function sai_() { var c = CacheService.getScriptCache(), n = parseInt(c.get('sai') || '0', 10) + 1; c.put('sai', String(n), 600); }

// ---------- GIAO TIẾP ----------
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.a === 'xuat') return xuat_(p);
  return json_({ ok: true });
}

function xuat_(p) {
  var khoa = cau_('B3');
  if (!khoa || String(p.k || '') !== khoa) return ContentService.createTextOutput('LOI: sai khoa').setMimeType(ContentService.MimeType.TEXT);
  var ky = String(p.ky || kyHienTai_()).trim(), cs = tab_(TAB_CS), n = cs.getLastRow(), v, i, j, lines = [CS_HEAD.join(',')], f;
  if (n >= 2) {
    v = cs.getRange(2, 1, n - 1, CS_HEAD.length).getDisplayValues();
    for (i = 0; i < v.length; i++) {
      if (String(v[i][0]).trim() !== ky) continue;
      f = [];
      for (j = 0; j < CS_HEAD.length; j++) f.push('"' + String(v[i][j]).replace(/"/g, '""') + '"');
      lines.push(f.join(','));
    }
  }
  return ContentService.createTextOutput(lines.join('\n')).setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var raw = String(e.postData.contents), d = JSON.parse(raw.substring(Math.max(raw.indexOf('{'), 0))), a = String(d.a || 'ghi');
    if (a === 'dauky') {
      var kh = cau_('B3');
      if (!kh || String(d.k || '') !== kh) return json_({ ok: false, loi: 'Sai khoa' });
      var rows = Utilities.parseCsv(String(d.csv || '').replace(/^\uFEFF/, ''));
      if (rows.length < 2) return json_({ ok: false, loi: 'File rong' });
      var tk = tab_(TAB_DK), w = 9, out = [], q;
      for (q = 1; q < rows.length; q++) { while (rows[q].length < w) rows[q].push(''); if (String(rows[q][1]).trim() !== '') out.push(rows[q].slice(0, w)); }
      tk.getRange(2, 1, Math.max(tk.getMaxRows() - 1, 1), w).clearContent();
      tk.getRange(2, 1, out.length, w).setNumberFormat('@').setValues(out);
      return json_({ ok: true, n: out.length });
    }
    if (khoa_()) return json_({ ok: false, loi: 'Tạm khoá do nhập sai PIN nhiều lần. Thử lại sau 10 phút.' });
    var ten = nhanVien_(d.pin);
    if (!ten) { sai_(); return json_({ ok: false, loi: 'PIN không đúng', pin: false }); }
    var ky = kyHienTai_(), mo = dotMo_();
    if (a === 'dn') return json_({ ok: true, ten: ten, ky: ky, mo: mo });
    if (a === 'ds') {
      var dk = dauKy_(), dg = daGhi_(ky), ds = [], i, x;
      for (i = 0; i < dk.ds.length; i++) {
        x = dk.ds[i];
        ds.push([x.sap, x.loai, x.ten, x.kv, x.cu, x.tb3, dg[x.sap + '|' + x.loai] === undefined ? '' : dg[x.sap + '|' + x.loai]]);
      }
      return json_({ ok: true, ten: ten, ky: ky, mo: mo, kyDauKy: dk.ky, ngay: nguong_(), ds: ds });
    }
    if (a === 'ghi') return ghi_(d, ten, ky, mo);
    return json_({ ok: false, loi: 'Lệnh không hợp lệ' });
  } catch (err) {
    return json_({ ok: false, loi: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function ghi_(d, ten, ky, mo) {
  if (!mo) return json_({ ok: false, loi: 'Đợt ghi đang ĐÓNG. Liên hệ quản lý để mở đợt.', dong: true });
  var sap = chuanSap_(d.sap), loai = chuanLoai_(d.loai), moi = so_(d.moi), id = String(d.id || '').trim();
  var gc = String(d.ghichu || '').replace(/[\r\n,]+/g, ' ').substring(0, 200);
  var qv = d.qv ? 'Có' : '';
  if (!sap) return json_({ ok: false, loi: 'Thiếu mã sạp' });
  if (loai !== 'D' && loai !== 'N') return json_({ ok: false, loi: 'Loại phải là Điện hoặc Nước' });
  if (moi === null) return json_({ ok: false, loi: 'Chỉ số không hợp lệ' });
  var cs = tab_(TAB_CS), last = cs.getLastRow(), i, ids;
  if (id && last > 1) {
    var from = Math.max(2, last - 2000);
    ids = cs.getRange(from, 8, last - from + 1, 1).getValues();
    for (i = 0; i < ids.length; i++) if (String(ids[i][0]) === id) return json_({ ok: true, trung: true });
  }
  var dk = dauKy_(), khop = null;
  for (i = 0; i < dk.ds.length; i++) if (dk.ds[i].sap === sap && dk.ds[i].loai === loai) { khop = dk.ds[i]; break; }
  var tt = khop ? 'KHOP' : 'CHUA_KHOP';
  if (!khop) {
    if (gc === '') return json_({ ok: false, loi: 'Mã sạp không có trong danh sách. Hãy ghi vị trí đồng hồ vào ô ghi chú để nhân viên văn phòng đối chiếu.' });
  } else if (khop.cu !== null) {
    if (moi < khop.cu && !d.qv) return json_({ ok: false, loi: 'Số mới (' + moi + ') nhỏ hơn số cũ (' + khop.cu + '). Nếu đồng hồ quay vòng hoặc mới thay, hãy tích ô tương ứng.' });
    var tieuThu = moi - khop.cu;
    if (!d.qv && khop.tb3 > 0 && tieuThu > nguong_() * khop.tb3 && gc === '') return json_({ ok: false, loi: 'Tiêu thụ ' + tieuThu + ' cao hơn ' + nguong_() + ' lần trung bình (' + khop.tb3 + '). Hãy kiểm tra lại số và ghi lý do.' });
  }
  var tg = Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm:ss');
  cs.appendRow([ky, sap, loai, String(moi), tg, gc, ten, id, tt, qv]);
  return json_({ ok: true, tt: tt });
}
