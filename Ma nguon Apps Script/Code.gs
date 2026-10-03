/**
 * APP SỔ THEO DÕI HỌC SINH – BẢN THỬ NGHIỆM (Bước 3)
 * Google Apps Script gắn với 1 file Google Sheets (file này là nơi lưu toàn bộ dữ liệu).
 *
 * Phạm vi bản thử: Trang chủ, Học sinh (tải danh sách), Điểm danh,
 * Tải dữ liệu học kì (Sổ báo giảng + Kho thư viện nhận xét, ghép tiết với bài), Sổ theo dõi.
 *
 * Quy ước: mọi ngày lưu dạng chữ dd/MM/yyyy; mã định danh lưu dạng chữ.
 */

// ============================================================================ CẤU HÌNH BẢNG
var TZ = 'Asia/Ho_Chi_Minh';
var BANG = {
  CaiDat:       ['Khoa', 'GiaTri'],
  HocKy:        ['MaHK', 'NamHoc', 'HocKy', 'NgayBatDau', 'NgayDayCuoi', 'NgayTaiSBG', 'SoTiet'],
  MonHoc:       ['TenMon', 'TenSheetBieuMau', 'CoDiemKTDK', 'TenTrongSBG', 'DangDay', 'ThuTu'],
  HocSinh:      ['MaDinhDanh', 'HoTen', 'NgaySinh', 'GioiTinh', 'DanToc', 'TrangThai', 'SDT', 'SoBuoiTuan',
                 'Lop', 'NamHoc', 'Truong', 'NgayChuyenDi', 'ThuTu'],
  TuanHoc:      ['MaHK', 'Tuan', 'TuNgay', 'DenNgay', 'ChuDe'],
  LichBaoGiang: ['MaHK', 'Ngay', 'Thu', 'Tuan', 'Buoi', 'Tiet', 'MonSBG', 'Mon', 'TenBaiDay', 'Bai', 'CachGhep'],
  ThuVien:      ['Khoi', 'BoSach', 'HocKy', 'Mon', 'Tuan', 'ChuDe', 'Bai', 'TrangSGV', 'HoatDong', 'CanCu',
                 'MucDo', 'Mau', 'NoiDung', 'PCNL'],
  KhoThuVien:   ['Khoi', 'BoSach', 'HocKy', 'Mon', 'SoCau', 'SoBai', 'NgayTai', 'TenFile'],
  NhanXetChung: ['Nhom', 'MucDo', 'Mau', 'NoiDung', 'PCNL'],
  DiemDanh:     ['Ngay', 'MaHK', 'MaDinhDanh', 'TrangThai', 'Buoi', 'GhiChu', 'CapNhat'],
  SoTheoDoi:    ['ID', 'ThoiGian', 'Ngay', 'MaHK', 'Tuan', 'Thang', 'Tiet', 'MaDinhDanh', 'Mon', 'Bai',
                 'HoatDong', 'MucDo', 'NoiDung', 'GhiChu'],
  // Nhận xét tháng cô đã SỬA (bản app tự ghép không lưu – luôn tính lại từ sổ theo dõi)
  NhanXetThang: ['NamHoc', 'Thang', 'MaDinhDanh', 'Sheet', 'NXMon', 'NLChung', 'NLDacThu', 'PhamChat', 'CapNhat'],
  // Đánh giá định kỳ cô đã chốt/sửa: Phan = tên môn (Muc T/H/C, Diem, NhanXet) hoặc 'NLPC' (15 mức dạng JSON + 3 nhận xét)
  DanhGiaKy:    ['NamHoc', 'Ky', 'MaDinhDanh', 'Phan', 'Muc', 'Diem', 'NhanXet', 'MucNLPC', 'NXNLChung', 'NXNLDacThu', 'NXPhamChat', 'CapNhat'],
  BaiKiemTra:   ['NamHoc', 'Ky', 'MaDinhDanh', 'Mon', 'FileId', 'Url', 'DiemAI', 'CapNhat'],
  PhienBanNX:   ['NamHoc', 'Ky', 'MaDinhDanh', 'Phan', 'Truong', 'Truoc', 'Sau', 'Lo', 'Luc'],
  MinhChung:    ['MaMC', 'NamHoc', 'MaDinhDanh', 'Ten', 'Loai', 'SoTrang', 'NgayTai', 'GhiChu', 'ThuMucId'],
  MinhChung_Trang: ['MaMC', 'TrangSo', 'FileId', 'Url', 'KichThuoc', 'AnhNho'],
  PhongTrao:    ['MaPT', 'NamHoc', 'Ten', 'Loai', 'Cap', 'Ngay', 'GhiChu', 'CapNhat'],
  ThamGia:      ['MaPT', 'MaDinhDanh', 'KetQua', 'GhiChu', 'CapNhat'],
  // Đánh giá thường xuyên hằng tháng (cách mới): thư viện tháng, mức cô chọn + phần cô sửa, nộp vở hằng ngày
  ThuVienThang: ['Khoi', 'BoSach', 'HocKy', 'Thang', 'Mon', 'TieuChi', 'MucDo', 'Mau', 'NoiDung'],
  DanhGiaThang: ['NamHoc', 'Thang', 'MaDinhDanh', 'Phan', 'MucKT', 'MucKN', 'Muc', 'NhanXet', 'CapNhat'],
  NopVo:        ['Ngay', 'MaDinhDanh', 'Mon', 'TrangThai', 'GhiChu', 'CapNhat'],
  // Thư viện tháng bản 2 (theo ý cô): mỗi tháng 2 nội dung Kiến thức / Kỹ năng, mỗi nội dung các yêu cầu cần đạt × 4 mức × 3 câu
  ThuVienYeuCau: ['Khoi', 'BoSach', 'HocKy', 'Thang', 'Mon', 'NoiDung', 'MaYC', 'YeuCau', 'CanCu', 'MucDo', 'Mau', 'CauNX'],
  MauCuaCo:     ['Mon', 'NoiDung', 'YeuCau', 'MucDo', 'CauNX', 'CapNhat'],
  // Nhóm học sinh cần quan tâm: Nhom = các nhóm cách nhau ';', GhiChu 1 dòng, BoGoiY = gợi ý cô đã bỏ qua ("nhóm|yyyy-mm")
  NhomHS:       ['MaDinhDanh', 'Nhom', 'GhiChu', 'BoGoiY', 'CapNhat']
};
var MUC4 = ['Hoàn thành xuất sắc', 'Hoàn thành tốt', 'Hoàn thành', 'Chưa hoàn thành'];
var MUC_PC = ['Tốt', 'Đạt', 'Cần cố gắng'];

// Thư viện tạm: dùng cho môn chưa có thư viện theo bài (4 mức × 4 câu chung)
var THU_VIEN_TAM = {
  'Hoàn thành xuất sắc': ['Nắm vững bài, làm rất tốt.', 'Tích cực, sáng tạo trong giờ học.', 'Vận dụng linh hoạt kiến thức.', 'Trình bày rõ ràng, tự tin.'],
  'Hoàn thành tốt': ['Hoàn thành tốt nội dung bài học.', 'Làm đúng các bài tập.', 'Hiểu bài, tham gia tích cực.', 'Thực hiện tốt yêu cầu của bài.'],
  'Hoàn thành': ['Hoàn thành nội dung bài học.', 'Làm được bài khi có gợi ý.', 'Còn sai sót nhỏ.', 'Cần cẩn thận hơn.'],
  'Chưa hoàn thành': ['Chưa hoàn thành bài học.', 'Chưa nắm được nội dung bài.', 'Cần cố gắng nhiều hơn.', 'Cần ôn lại bài ở nhà.']
};

// ============================================================================ WEB APP
function doGet() {
  // Index.html là 1 file đã gộp sẵn mọi thứ (không còn lệnh include) → mở thẳng, không cần xử lý mẫu
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Sổ theo dõi học sinh')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}
function include(name) { return HtmlService.createHtmlOutputFromFile(name).getContent(); }

// Đường link của app (bản đã triển khai) – dùng cho menu trong Google Trang tính và mã QR trong Cài đặt
function layDuongLink() {
  var url = ScriptApp.getService().getUrl() || '';
  return url.replace(/\/dev$/, '/exec');            // luôn đưa link bản chính thức, không đưa link bản thử của người lập trình
}
// Menu "📒 Sổ theo dõi" xuất hiện mỗi khi mở file Google Trang tính trên máy tính
function onOpen() {
  SpreadsheetApp.getUi().createMenu('📒 Sổ theo dõi')
    .addItem('Mở app', 'moAppTuMenu')
    .addItem('Lấy đường link / mã QR để mở trên điện thoại', 'moAppTuMenu')
    .addSeparator()
    .addItem('Làm mới dữ liệu (sau khi sửa trực tiếp trên Trang tính)', 'lamMoiDuLieu')
    .addToUi();
}
function moAppTuMenu() {
  var url = layDuongLink();
  var html = url
    ? '<div style="font-family:Arial;font-size:14px">' +
      '<p><a href="' + url + '" target="_blank" style="font-size:18px;font-weight:bold">▶ Mở app Sổ theo dõi</a></p>' +
      '<p>Mở trên điện thoại: quét mã QR này bằng camera.</p><div id="qr"></div>' +
      '<p style="word-break:break-all;color:#555">' + url + '</p></div>' +
      // mã QR tạo ngay trên trình duyệt, không gửi đường link ra dịch vụ bên ngoài
      '<script src="https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js"></script>' +
      '<script>new QRCode(document.getElementById("qr"), {text: ' + JSON.stringify(url) + ', width: 200, height: 200});</script>'
    : '<p style="font-family:Arial">App chưa được triển khai. Vào Tiện ích mở rộng → Apps Script → Triển khai → Tùy chọn triển khai mới → Ứng dụng web.</p>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(340).setHeight(420), 'Sổ theo dõi học sinh');
}

// Chạy 1 lần sau khi dán mã: tạo các bảng, môn học mặc định, cài đặt mặc định.
function caiDatLanDau() {
  var ss = SpreadsheetApp.getActive();
  Object.keys(BANG).forEach(function (ten) {
    var sh = ss.getSheetByName(ten) || ss.insertSheet(ten);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, BANG[ten].length).setValues([BANG[ten]]).setFontWeight('bold').setBackground('#D9E2F3');
      sh.setFrozenRows(1);
    }
    sh.getRange(1, 1, sh.getMaxRows(), BANG[ten].length).setNumberFormat('@');   // tất cả dạng chữ
  });
  var md = ss.getSheetByName('Sheet1') || ss.getSheetByName('Trang tính1');
  if (md && md.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(md);
  if (doc_('MonHoc').length === 0) {
    ghiTatCa_('MonHoc', [
      ['Tiếng Việt', 'Tiếng việt', 'Có', 'TIẾNG VIỆT; CỦNG CỐ TIẾNG VIỆT', 'Có', '1'],
      ['Toán', 'Toán', 'Có', 'TOÁN; CỦNG CỐ TOÁN', 'Có', '2'],
      ['Hoạt động trải nghiệm', 'Hoạt động trải nghiệm', 'Không', 'HĐTN', 'Có', '3'],
      ['Đạo đức', 'Đạo đức', 'Không', 'ĐẠO ĐỨC', 'Có', '4']
    ]);
  }
  var cd = caiDat_();
  var macDinh = { BoSach: 'Kết nối tri thức với cuộc sống', NgayNhacThang: '25', SoNgayNhacHK: '14' };
  Object.keys(macDinh).forEach(function (k) { if (!cd[k]) luuCaiDat_(k, macDinh[k]); });
  return 'Đã tạo xong các bảng dữ liệu.';
}

// ============================================================================ TIỆN ÍCH BẢNG
function sh_(ten) {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(ten);
  if (!sh && BANG[ten] && ss.getSheetByName('CaiDat')) {   // bảng mới thêm ở bản cập nhật: tự tạo, không cần chạy lại caiDatLanDau
    sh = ss.insertSheet(ten);
    sh.getRange(1, 1, 1, BANG[ten].length).setValues([BANG[ten]]).setFontWeight('bold').setBackground('#D9E2F3');
    sh.setFrozenRows(1);
    sh.getRange(1, 1, sh.getMaxRows(), BANG[ten].length).setNumberFormat('@');
  }
  if (!sh) throw new Error('Chưa có bảng "' + ten + '". Hãy chạy hàm caiDatLanDau trước.');
  return sh;
}
// ---------------------------------------------------------------------------- Đọc / ghi bảng (có nhớ tạm để nhanh hơn)
// 1) Trong 1 lần gọi máy chủ: mỗi bảng chỉ đọc Trang tính 1 lần (_NHO), ghi thì xoá phần nhớ của bảng đó.
// 2) Bảng lớn, ít đổi (BANG_DEM): nhớ thêm 6 giờ trong CacheService của Google; mỗi lần app ghi bảng đó thì tăng "phiên bản"
//    nên bản nhớ cũ không bao giờ được dùng lại. Sửa TRỰC TIẾP trên Trang tính thì dùng menu "📒 Sổ theo dõi → Làm mới dữ liệu".
var _NHO = {}, _SHEET = {};
var BANG_DEM = { CaiDat: 1, MonHoc: 1, HocKy: 1, TuanHoc: 1, LichBaoGiang: 1, ThuVien: 1, KhoThuVien: 1, NhanXetChung: 1, ThuVienThang: 1, ThuVienYeuCau: 1, MauCuaCo: 1, NhomHS: 1 };
function sheetNho_(ten) { return _SHEET[ten] || (_SHEET[ten] = sh_(ten)); }
function docTrangTinh_(ten) {     // mảng các dòng (mảng giá trị), không kể dòng tiêu đề
  var sh = sheetNho_(ten), n = sh.getLastRow();
  return n < 2 ? [] : sh.getRange(2, 1, n - 1, BANG[ten].length).getDisplayValues();
}
function phienMoi_() { return Date.now() + '_' + Math.random().toString(36).slice(2, 8); }   // không trùng kể cả khi ghi 2 lần trong 1 mili-giây
var _PB = null;                   // phiên bản các bảng, đọc 1 lần mỗi lần gọi máy chủ
function phienBan_(ten) { if (!_PB) _PB = thuocTinh_().getProperties(); return _PB['pb_' + ten] || '0'; }
function docDem_(ten) {
  var c = CacheService.getScriptCache(), k = 'b_' + ten + '_' + phienBan_(ten);
  try {
    var soPhan = c.get(k + '_n');
    if (soPhan) {
      var khoa = []; for (var i = 0; i < +soPhan; i++) khoa.push(k + '_' + i);
      var m = c.getAll(khoa), s = '';
      for (i = 0; i < khoa.length; i++) { if (m[khoa[i]] == null) { s = null; break; } s += m[khoa[i]]; }
      if (s != null) return JSON.parse(s);
    }
  } catch (e) {}
  var rows = docTrangTinh_(ten);
  try {        // mỗi phần ≤ 30 000 ký tự (giới hạn 100 KB/ô nhớ, chữ Việt tới 3 byte/ký tự)
    var json = JSON.stringify(rows), n = Math.ceil(json.length / 30000), phan = {};
    if (n <= 200) { for (var j = 0; j < n; j++) phan[k + '_' + j] = json.substr(j * 30000, 30000); phan[k + '_n'] = String(n); c.putAll(phan, 21600); }
  } catch (e) {}
  return rows;
}
function boNho_(ten) {            // gọi sau mỗi lần ghi bảng
  delete _NHO[ten];
  if (ten === 'SoTheoDoi') delete _NHO._monBaiSTD;
  if (BANG_DEM[ten]) { var v = phienMoi_(); thuocTinh_().setProperty('pb_' + ten, v); if (_PB) _PB['pb_' + ten] = v; }
}
function lamMoiDuLieu() {         // menu: sau khi sửa trực tiếp trên Trang tính
  _NHO = {}; _PB = null; Object.keys(BANG_DEM).forEach(function (t) { thuocTinh_().setProperty('pb_' + t, phienMoi_()); });
  try { SpreadsheetApp.getActive().toast('App sẽ đọc lại dữ liệu mới từ Trang tính.', 'Đã làm mới', 5); } catch (e) {}
}
function doc_(ten) {            // đọc bảng thành mảng đối tượng {cột: giá trị} (bản sao, sửa thoải mái)
  if (!_NHO[ten]) _NHO[ten] = BANG_DEM[ten] ? docDem_(ten) : docTrangTinh_(ten);
  var cot = BANG[ten];
  return _NHO[ten].map(function (r) { var o = {}; cot.forEach(function (c, i) { o[c] = r[i] == null ? '' : r[i]; }); return o; });
}
function hang_(ten, o) { return BANG[ten].map(function (c) { return o[c] == null ? '' : String(o[c]); }); }
function duDong_(sh, cuoi) {      // bảo đảm Trang tính có đủ số dòng
  if (sh.getMaxRows() < cuoi) sh.insertRowsAfter(sh.getMaxRows(), cuoi - sh.getMaxRows());
}
function ghiTatCa_(ten, rows) {  // ghi đè toàn bộ dữ liệu (giữ dòng tiêu đề)
  var sh = sheetNho_(ten), w = BANG[ten].length, n = sh.getLastRow();
  if (n > 1) sh.getRange(2, 1, n - 1, w).clearContent();
  if (rows.length) { duDong_(sh, rows.length + 1); sh.getRange(2, 1, rows.length, w).setNumberFormat('@').setValues(rows); }
  boNho_(ten);
}
function themDong_(ten, rows) {
  if (!rows.length) return;
  var sh = sheetNho_(ten), w = BANG[ten].length, dau = sh.getLastRow() + 1;
  duDong_(sh, dau + rows.length - 1);
  sh.getRange(dau, 1, rows.length, w).setNumberFormat('@').setValues(rows);
  boNho_(ten);
}
// Giữ các dòng thoả giuLai(o), thay phần còn lại bằng moi. Nhanh: nếu các dòng bị thay nằm liền nhau thì chỉ ghi
// đúng đoạn đó (VD lưu điểm danh 1 ngày, sửa 1 ô đánh giá, hoàn tác nhận xét vừa lưu); còn lại mới ghi lại cả bảng.
function thayTheTheo_(ten, giuLai, moi) {
  var tat = doc_(ten), bo = [];
  tat.forEach(function (o, i) { if (!giuLai(o)) bo.push(i); });
  if (!bo.length) return themDong_(ten, moi);
  var sh = sheetNho_(ten), w = BANG[ten].length, dau = bo[0] + 2, lienTuc = bo[bo.length - 1] - bo[0] + 1 === bo.length;
  if (lienTuc && (moi.length === bo.length || bo[bo.length - 1] === tat.length - 1)) {
    if (moi.length) { duDong_(sh, dau + moi.length - 1); sh.getRange(dau, 1, moi.length, w).setNumberFormat('@').setValues(moi); }
    if (moi.length < bo.length) sh.getRange(dau + moi.length, 1, bo.length - moi.length, w).clearContent();
    return boNho_(ten);
  }
  if (lienTuc && !moi.length) {     // xoá 1 đoạn giữa bảng
    try { sh.deleteRows(dau, bo.length); return boNho_(ten); } catch (e) {}
  }
  ghiTatCa_(ten, tat.filter(giuLai).map(function (o) { return hang_(ten, o); }).concat(moi));
}
function caiDat_() { var o = {}; doc_('CaiDat').forEach(function (r) { o[r.Khoa] = r.GiaTri; }); return o; }
function luuCaiDat_(k, v) {
  var ds = doc_('CaiDat'), co = false;
  ds.forEach(function (r) { if (r.Khoa === k) { r.GiaTri = v; co = true; } });
  if (!co) ds.push({ Khoa: k, GiaTri: v });
  ghiTatCa_('CaiDat', ds.map(function (r) { return [r.Khoa, r.GiaTri]; }));
}
function voiKhoa_(fn) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  _NHO = {};                       // trong khoá: luôn đọc dữ liệu mới nhất (lần gọi khác có thể vừa ghi)
  try { return fn(); } finally { lock.releaseLock(); }
}

// ============================================================================ NGÀY THÁNG
function homNay_() { return Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy'); }
function key_(dmy) { var p = String(dmy).split('/'); return (+p[2] || 0) * 10000 + (+p[1]) * 100 + (+p[0]); }
function toDate_(dmy) { var p = String(dmy).split('/'); return new Date(+p[2], +p[1] - 1, +p[0]); }
function soNgay_(a, b) { return Math.round((toDate_(b) - toDate_(a)) / 86400000); }
var THU = ['Chủ nhật', 'Thứ Hai', 'Thứ Ba', 'Thứ Tư', 'Thứ Năm', 'Thứ Sáu', 'Thứ Bảy'];

// ============================================================================ HỌC KÌ
function hocKyHienTai_() {
  var ma = caiDat_().MaHKHienTai; if (!ma) return null;
  return doc_('HocKy').filter(function (h) { return h.MaHK === ma; })[0] || null;
}
function hkSau_(hk, namHoc) {    // học kì kế tiếp
  if (!hk) return { namHoc: namHoc, hocKy: 'I' };
  if (hk.HocKy === 'I') return { namHoc: hk.NamHoc, hocKy: 'II' };
  var y = /(\d{4})-(\d{4})/.exec(hk.NamHoc); return { namHoc: y ? (+y[1] + 1) + '-' + (+y[2] + 1) : '', hocKy: 'I' };
}
function monDangDay_() { return doc_('MonHoc').filter(function (m) { return m.DangDay === 'Có'; }); }
function khoThieu_(khoi, boSach, hocKy) {       // các môn đang dạy chưa có thư viện cho học kì này
  var kho = doc_('KhoThuVien');
  return monDangDay_().map(function (m) { return m.TenMon; }).filter(function (mon) {
    return !kho.some(function (k) { return k.Khoi === khoi && k.BoSach === boSach && k.HocKy === hocKy && k.Mon === mon; });
  });
}

// ============================================================================ TRANG CHỦ
function layTrangChu() {
  var cd = caiDat_(), hk = hocKyHienTai_(), nay = homNay_(), d = toDate_(nay);
  var out = {
    homNay: nay, thu: THU[d.getDay()], thang: d.getMonth() + 1, truong: cd.Truong || '', lop: cd.Lop || '',
    namHoc: cd.NamHoc || '', hocKy: hk ? hk.HocKy : '', tuan: '', chuDe: '', lich: [], nhac: [], siSo: 0
  };
  out.siSo = hocSinhDangHoc_(nay).length;
  out.giaoDien = giaoDien_();
  if (!cd.NamHoc || !out.siSo) out.nhac.push({ muc: 'do', text: 'Chưa có danh sách học sinh. Vào Học sinh → Tải danh sách.', toi: 'hocsinh' });
  if (!hk) {
    out.nhac.push({ muc: 'do', text: 'Chưa có dữ liệu học kì. Vào Cài đặt → Tải dữ liệu học kì.', toi: 'caidat' });
    return out;
  }
  var tuan = doc_('TuanHoc').filter(function (t) { return t.MaHK === hk.MaHK && key_(t.TuNgay) <= key_(nay) + 1 && key_(nay) <= key_(t.DenNgay) + 2; })[0];
  if (tuan) { out.tuan = tuan.Tuan; out.chuDe = tuan.ChuDe; }
  out.lich = doc_('LichBaoGiang').filter(function (r) { return r.MaHK === hk.MaHK && r.Ngay === nay; })
    .sort(function (a, b) { return +a.Tiet - +b.Tiet; })
    .map(function (r) { return { tiet: r.Tiet, buoi: r.Buoi, monSBG: r.MonSBG, mon: r.Mon, ten: r.TenBaiDay, bai: r.Bai }; });

  // Nhắc: thư viện còn thiếu cho học kì hiện tại
  var thieu = khoThieu_(cd.Khoi, cd.BoSach, hk.HocKy);
  if (thieu.length) out.nhac.push({ muc: 'vang', text: 'Đang dùng thư viện tạm cho: ' + thieu.join(', ') + ' (HK ' + hk.HocKy + '). Tải thư viện khi có.', toi: 'caidat' });
  // Nhắc: sắp hết học kì / đã hết học kì
  var conLai = soNgay_(nay, hk.NgayDayCuoi), soNgayNhac = +(cd.SoNgayNhacHK || 14), sau = hkSau_(hk, cd.NamHoc);
  var thieuSau = khoThieu_(cd.Khoi, cd.BoSach, sau.hocKy);
  var chuanBi = 'Sổ báo giảng HK ' + sau.hocKy + (thieuSau.length ? ' và thư viện nhận xét: ' + thieuSau.join(', ') : '');
  if (conLai < 0 && hk.HocKy === 'II') out.nhac.push({ muc: 'do', text: 'Năm học ' + hk.NamHoc + ' đã kết thúc (' + hk.NgayDayCuoi + '). Cô xuất đủ CK2 môn học, năng lực – phẩm chất CK2 và Nhận xét GVCN; sau đó Cài đặt → Năm học → Bắt đầu năm học mới.', toi: 'caidat' });
  else if (conLai < 0) out.nhac.push({ muc: 'do', text: 'Học kì ' + hk.HocKy + ' đã kết thúc (' + hk.NgayDayCuoi + '). Cô tải ' + chuanBi + ' để bắt đầu học kì mới.', toi: 'caidat' });
  else if (conLai === 0) out.nhac.push({ muc: 'do', text: 'Hôm nay là ngày dạy cuối HK ' + hk.HocKy + '. Cô hoàn thành đánh giá cuối kì và chuẩn bị ' + chuanBi + '.', toi: 'caidat' });
  else if (conLai <= soNgayNhac) out.nhac.push({ muc: 'vang', text: 'Còn ' + conLai + ' ngày nữa kết thúc HK ' + hk.HocKy + ' (ngày dạy cuối ' + hk.NgayDayCuoi + '). Cô chuẩn bị ' + chuanBi + '.', toi: 'caidat' });
  // Nhắc: đánh giá định kỳ, nhận xét GVCN (các mục này chỉ hiện quanh thời điểm cần làm)
  out.giaoDien.kyNhac.forEach(function (k) { out.nhac.push({ muc: 'vang', text: 'Đến kì đánh giá ' + TEN_KY[k] + ' (' + k + '): vào Đánh giá → Định kỳ → chọn ' + k + ' → Tổng hợp.', toi: 'danhgia' }); });
  if (out.giaoDien.hienGVCN && out.giaoDien.kyNhac.indexOf('CK2') >= 0) out.nhac.push({ muc: 'vang', text: 'Cuối năm: viết Nhận xét của GVCN – vào Đánh giá → Định kỳ → Nhận xét GVCN cuối năm.', toi: 'danhgia' });
  // Nhắc: nhận xét tháng
  if (d.getDate() >= +(cd.NgayNhacThang || 25)) out.nhac.push({ muc: 'xanh', text: 'Đến kì đánh giá thường xuyên tháng ' + (d.getMonth() + 1) + ': vào Đánh giá → Thường xuyên (hằng tháng) → chọn mức Toán, Tiếng Việt cho từng em → xuất file mẫu đánh giá thường xuyên.', toi: 'danhgia' });
  return out;
}

// ============================================================================ CÀI ĐẶT
function layCaiDat() {
  var cd = caiDat_(), hk = hocKyHienTai_();
  return {
    caiDat: cd, monHoc: doc_('MonHoc'), hocKy: doc_('HocKy'), hienTai: hk ? hk.MaHK : '',
    kho: doc_('KhoThuVien'), khoThang: khoThang_(), soNhanXetChung: doc_('NhanXetChung').length,
    hkSau: hkSau_(hk, cd.NamHoc)
  };
}
function luuCaiDatChung(o) {
  return voiKhoa_(function () {
    ['Truong', 'Lop', 'Khoi', 'NamHoc', 'BoSach', 'NgayNhacThang', 'SoNgayNhacHK'].forEach(function (k) { if (o[k] != null) luuCaiDat_(k, String(o[k]).trim()); });
    return 'Đã lưu cài đặt.';
  });
}
function luuMonHoc(ds) {
  return voiKhoa_(function () {
    ghiTatCa_('MonHoc', ds.filter(function (m) { return m.TenMon; }).map(function (m, i) { m.ThuTu = i + 1; return hang_('MonHoc', m); }));
    return 'Đã lưu danh sách môn học.';
  });
}
function chonHocKy(maHK) { luuCaiDat_('MaHKHienTai', maHK); return 'Đã chuyển sang học kì ' + maHK; }

// ============================================================================ HỌC SINH
function hocSinhDangHoc_(ngay) {
  return doc_('HocSinh').filter(function (h) {
    return h.TrangThai !== 'Đã chuyển đi' || (h.NgayChuyenDi && key_(ngay) < key_(h.NgayChuyenDi));
  }).sort(function (a, b) { return +a.ThuTu - +b.ThuTu; });
}
function layHocSinh() {
  var tat = doc_('HocSinh');
  return {
    dangHoc: tat.filter(function (h) { return h.TrangThai !== 'Đã chuyển đi'; }).sort(function (a, b) { return +a.ThuTu - +b.ThuTu; }),
    chuyenDi: tat.filter(function (h) { return h.TrangThai === 'Đã chuyển đi'; })
  };
}
// info: {truong, lop, khoi, namHoc}; ds: [{maDinhDanh, hoTen, ...}] (đã đọc ở trình duyệt)
function nhapDanhSachHocSinh(info, ds) {
  return voiKhoa_(function () {
    if (!ds || !ds.length) throw new Error('File không có học sinh nào – không nhập để tránh xoá nhầm cả lớp.');
    var cd = caiDat_(), nay = homNay_();
    if (cd.NamHoc && info.namHoc && cd.NamHoc !== info.namHoc)
      throw new Error('File là năm học ' + info.namHoc + ' nhưng app đang ở năm học ' + cd.NamHoc + '. Muốn sang năm học mới: Cài đặt → Năm học → Bắt đầu năm học mới.');
    var cu = {}; doc_('HocSinh').forEach(function (h) { cu[h.MaDinhDanh] = h; });
    var dem = { moi: 0, capNhat: 0, chuyenDi: 0, boQua: 0 }, moiDs = [];
    var chuyenDi = function (t) { return /chuyển đi|thôi học|nghỉ học/i.test(t || ''); };
    ds.forEach(function (x, i) {
      var h = cu[x.maDinhDanh];
      if (!h) {
        if (chuyenDi(x.trangThai)) { dem.boQua++; return; }      // chưa từng có trong lớp: bỏ qua
        dem.moi++;
        h = { MaDinhDanh: x.maDinhDanh };
      } else if (chuyenDi(x.trangThai) && h.TrangThai !== 'Đã chuyển đi') {
        dem.chuyenDi++;
      } else dem.capNhat++;
      h.HoTen = x.hoTen; h.NgaySinh = x.ngaySinh; h.GioiTinh = x.gioiTinh; h.DanToc = x.danToc; h.SDT = x.sdt; h.SoBuoiTuan = x.soBuoi;
      h.Lop = info.lop; h.NamHoc = info.namHoc; h.Truong = info.truong; h.ThuTu = i + 1;
      if (chuyenDi(x.trangThai)) { if (h.TrangThai !== 'Đã chuyển đi') h.NgayChuyenDi = nay; h.TrangThai = 'Đã chuyển đi'; }
      else { h.TrangThai = x.trangThai || 'Đang học'; h.NgayChuyenDi = ''; }
      cu[x.maDinhDanh] = h; moiDs.push(x.maDinhDanh);
    });
    // Học sinh có trong app nhưng không còn trong file: coi như chuyển đi (giữ dữ liệu)
    Object.keys(cu).forEach(function (ma) {
      if (moiDs.indexOf(ma) < 0 && cu[ma].TrangThai !== 'Đã chuyển đi') { cu[ma].TrangThai = 'Đã chuyển đi'; cu[ma].NgayChuyenDi = nay; dem.chuyenDi++; }
    });
    ghiTatCa_('HocSinh', Object.keys(cu).map(function (ma) { return hang_('HocSinh', cu[ma]); }));
    if (info.truong) luuCaiDat_('Truong', info.truong);
    if (info.lop) luuCaiDat_('Lop', info.lop);
    if (info.khoi) luuCaiDat_('Khoi', info.khoi);
    if (info.namHoc) luuCaiDat_('NamHoc', info.namHoc);
    return 'Đã nhập: ' + dem.moi + ' học sinh mới, cập nhật ' + dem.capNhat + ', chuyển đi ' + dem.chuyenDi +
      (dem.boQua ? ', bỏ qua ' + dem.boQua + ' em đã chuyển đi trước khi nhập' : '') + '.';
  });
}

// ============================================================================ ĐIỂM DANH
function layDiemDanh(ngay) {
  ngay = ngay || homNay_();
  var hk = hocKyHienTai_(), co = {};
  doc_('DiemDanh').forEach(function (r) { if (r.Ngay === ngay) co[r.MaDinhDanh] = r; });
  return {
    ngay: ngay, maHK: hk ? hk.MaHK : '',
    hocSinh: hocSinhDangHoc_(ngay).map(function (h) {
      var r = co[h.MaDinhDanh] || {};
      return { ma: h.MaDinhDanh, ten: h.HoTen, trangThai: r.TrangThai || 'Có mặt', buoi: r.Buoi || 'Cả ngày', ghiChu: r.GhiChu || '' };
    }),
    daLuu: Object.keys(co).length > 0
  };
}
function luuDiemDanh(ngay, ds) {
  return voiKhoa_(function () {
    var hk = hocKyHienTai_(), luc = Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm');
    var moi = ds.map(function (x) { return hang_('DiemDanh', { Ngay: ngay, MaHK: hk ? hk.MaHK : '', MaDinhDanh: x.ma, TrangThai: x.trangThai, Buoi: x.buoi, GhiChu: x.ghiChu, CapNhat: luc }); });
    thayTheTheo_('DiemDanh', function (r) { return r.Ngay !== ngay; }, moi);
    var vang = ds.filter(function (x) { return x.trangThai !== 'Có mặt'; }).length;
    return 'Đã lưu điểm danh ' + ngay + ': có mặt ' + (ds.length - vang) + ', vắng ' + vang + '.';
  });
}

// ============================================================================ TẢI DỮ LIỆU HỌC KÌ
// sbg: {tiet:[...], tuan:[...]} đã đọc ở trình duyệt
function nhapSoBaoGiang(namHoc, hocKy, tenFile, sbg) {
  return voiKhoa_(function () {
    if (!sbg.tiet.length) throw new Error('Không đọc được tiết dạy nào trong file.');
    var ma = namHoc + '-HK' + hocKy, monMap = banDoMonSBG_();
    var ngays = sbg.tiet.map(function (t) { return t.ngay; }).sort(function (a, b) { return key_(a) - key_(b); });
    thayTheTheo_('HocKy', function (h) { return h.MaHK !== ma; }, [hang_('HocKy', {
      MaHK: ma, NamHoc: namHoc, HocKy: hocKy, NgayBatDau: ngays[0], NgayDayCuoi: ngays[ngays.length - 1],
      NgayTaiSBG: homNay_() + ' (' + tenFile + ')', SoTiet: sbg.tiet.length })]);
    thayTheTheo_('TuanHoc', function (t) { return t.MaHK !== ma; }, sbg.tuan.map(function (w) {
      return hang_('TuanHoc', { MaHK: ma, Tuan: w.tuan, TuNgay: w.tu, DenNgay: w.den, ChuDe: w.chuDe }); }));
    thayTheTheo_('LichBaoGiang', function (t) { return t.MaHK !== ma; }, sbg.tiet.map(function (t) {
      return hang_('LichBaoGiang', { MaHK: ma, Ngay: t.ngay, Thu: t.thu, Tuan: t.tuan, Buoi: t.buoi, Tiet: t.tiet,
        MonSBG: t.mon, Mon: monMap[t.mon] || '', TenBaiDay: t.bai, Bai: '', CachGhep: '' }); }));
    luuCaiDat_('MaHKHienTai', ma);
    var kq = ghepTietHocKy_(ma);
    return 'Đã nhập Sổ báo giảng HK ' + hocKy + ' (' + namHoc + '): ' + sbg.tiet.length + ' tiết, ' + sbg.tuan.length +
      ' tuần, từ ' + ngays[0] + ' đến ' + ngays[ngays.length - 1] + '. ' + kq;
  });
}
function banDoMonSBG_() {       // 'CỦNG CỐ TOÁN' -> 'Toán'
  var m = {};
  doc_('MonHoc').forEach(function (r) {
    String(r.TenTrongSBG || '').split(';').forEach(function (x) { x = x.trim().toUpperCase(); if (x) m[x] = r.TenMon; });
  });
  return m;
}
// tv: {thongTin, mon:[{ten, dong:[...]}], chung:[...]} đã đọc ở trình duyệt
function nhapThuVien(tenFile, tv) {
  return voiKhoa_(function () {
    var tt = tv.thongTin, khoi = String(tt.Khoi), bs = tt.BoSach, hk = tt.HocKy, nay = homNay_(), msg = [];
    tv.mon.forEach(function (m) {
      var cung = function (r) { return r.Khoi === khoi && r.BoSach === bs && r.HocKy === hk && r.Mon === m.ten; };
      thayTheTheo_('ThuVien', function (r) { return !cung(r); }, m.dong.map(function (d) {
        return hang_('ThuVien', { Khoi: khoi, BoSach: bs, HocKy: hk, Mon: m.ten, Tuan: d.tuan, ChuDe: d.chuDe, Bai: d.bai, TrangSGV: d.trangSGV,
          HoatDong: d.hoatDong, CanCu: d.canCu, MucDo: d.mucDo, Mau: d.mau, NoiDung: d.noiDung, PCNL: d.pcnl }); }));
      var soBai = {}; m.dong.forEach(function (d) { soBai[d.bai] = 1; });
      thayTheTheo_('KhoThuVien', function (r) { return !cung(r); }, [hang_('KhoThuVien', {
        Khoi: khoi, BoSach: bs, HocKy: hk, Mon: m.ten, SoCau: m.dong.length, SoBai: Object.keys(soBai).length, NgayTai: nay, TenFile: tenFile })]);
      msg.push(m.ten + ' (' + m.dong.length + ' câu)');
    });
    if (tv.chung && tv.chung.length) {
      ghiTatCa_('NhanXetChung', tv.chung.map(function (c) { return hang_('NhanXetChung', { Nhom: c.nhom, MucDo: c.mucDo, Mau: c.mau, NoiDung: c.noiDung, PCNL: c.pcnl }); }));
      msg.push('Nhận xét chung (' + tv.chung.length + ' câu)');
    }
    var cd = caiDat_(); if (!cd.BoSach) luuCaiDat_('BoSach', bs);
    var ketQua = 'Đã đưa vào Kho thư viện – Lớp ' + khoi + ', HK ' + hk + ': ' + msg.join('; ') + '.';
    var hkHT = hocKyHienTai_();
    if (hkHT && hkHT.HocKy === hk) ketQua += ' ' + ghepTietHocKy_(hkHT.MaHK);
    return ketQua;
  });
}

// ---------------------------------------------------------------------------- GHÉP TIẾT VỚI BÀI
function chuanHoa_(s) {
  return String(s || '').toLowerCase().normalize('NFC').replace(/\(.*?\)/g, ' ').replace(/[^0-9a-zà-ỹđ\s]/g, ' ').replace(/\s+/g, ' ').trim();
}
var TU_BO = { 'bài': 1, 'tiết': 1, 'đọc': 1, 'viết': 1, 'và': 1, 'luyện': 1, 'tập': 1, 'các': 1, 'của': 1, 'em': 1, 'một': 1, 'số': 0 };
function tu_(s) { return chuanHoa_(s).split(' ').filter(function (w) { return w.length > 1 && !TU_BO[w]; }); }
function trongTuan_(chuoiTuan, t) {          // '5–6' chứa 6 ?
  var m = String(chuoiTuan).match(/\d+/g); if (!m) return false;
  var a = +m[0], b = +(m[1] || m[0]); return t >= a && t <= b;
}
// Hàm thuần (kiểm tra được ngoài Google): tiets [{tuan, ten}] theo thứ tự dạy; bais [{bai, tuan, canCu}] theo thứ tự thư viện
function ghepTietVoiBai_(tiets, bais) {
  var cur = null, out = [];
  var soBai = function (b) { var m = /^Bài\s+(\d+)/i.exec(b.bai); return m ? +m[1] : null; };
  tiets.forEach(function (t) {
    var ten = chuanHoa_(t.ten), kq = null, cach = '';
    var m = /bài\s*(\d+)\s*[:.]/i.exec(t.ten);                          // 1) có ghi "BÀI 7:"
    if (m) kq = bais.filter(function (b) { return soBai(b) === +m[1]; })[0];
    if (kq) cach = 'số bài';
    if (!kq) {                                                          // 2) trùng nội dung tiết trong cột Căn cứ
      kq = bais.filter(function (b) { return String(b.canCu).split(' | ').some(function (x) { return chuanHoa_(x) === ten; }); })[0];
      if (kq) cach = 'khớp sổ báo giảng';
    }
    var ungVien = bais.filter(function (b) { return trongTuan_(b.tuan, t.tuan); });
    if (!kq) {                                                          // 3) tên bài nằm trọn trong nội dung tiết (lấy tên dài nhất)
      var gan = bais.filter(function (b) { return trongTuan_(b.tuan, t.tuan) || trongTuan_(b.tuan, t.tuan - 1) || trongTuan_(b.tuan, t.tuan + 1); });
      var tenDem = ' ' + ten + ' ', iCur0 = cur ? bais.indexOf(cur) : -1, diemMax = -1;
      gan.forEach(function (b) {        // ưu tiên: tên dài hơn > thuộc đúng tuần > nằm sau bài đang dạy
        var tb = chuanHoa_(b.bai.replace(/^Bài\s+\d+\s*:\s*/i, ''));
        if (!tb || tenDem.indexOf(' ' + tb + ' ') < 0) return;
        var d = tb.length * 10 + (ungVien.indexOf(b) >= 0 ? 5 : 0) + (bais.indexOf(b) > iCur0 ? 2 : 0);
        if (d > diemMax) { diemMax = d; kq = b; }
      });
      if (kq) cach = 'chứa tên bài';
    }
    if (!kq && ungVien.length) {                                        // 4) giống tên bài nhất trong tuần
      var tt = tu_(t.ten), best = null, diem = 0;
      ungVien.forEach(function (b) {
        var tb = tu_(b.bai.replace(/^Bài\s+\d+\s*:\s*/i, '')), d = 0;
        tb.forEach(function (w) { if (tt.indexOf(w) >= 0) d++; });
        var tiLe = tb.length ? d / tb.length : 0;
        if (d >= 2 && tiLe >= 0.5 && tiLe > diem) { diem = tiLe; best = b; }
      });
      if (best) { kq = best; cach = 'giống tên bài'; }
    }
    // 5) tiết chỉ ghi chung chung ("Luyện tập"…): tiếp bài đang dạy, trừ khi tuần này thư viện có bài khác mà bài đang dạy không thuộc tuần này
    if (!kq && cur && (ungVien.indexOf(cur) >= 0 || (!ungVien.length && trongTuan_(cur.tuan, t.tuan - 1)))) { kq = cur; cach = 'tiếp bài trước'; }
    if (!kq && ungVien.length) {                                        // 5) đoán: bài đầu tiên của tuần chưa dạy qua
      var iCur = cur ? bais.indexOf(cur) : -1;
      kq = ungVien.filter(function (b) { return bais.indexOf(b) > iCur; })[0] || ungVien[0]; cach = 'đoán';
    }
    if (kq) cur = kq;
    out.push(kq ? { bai: kq.bai, cach: cach } : { bai: '', cach: '' });
  });
  return out;
}
function ghepTietHocKy_(maHK) {
  var cd = caiDat_(), hk = doc_('HocKy').filter(function (h) { return h.MaHK === maHK; })[0];
  if (!hk) return '';
  var tv = doc_('ThuVien'), lich = doc_('LichBaoGiang'), dem = { tong: 0, ghep: 0, doan: 0, thieu: 0 };
  monDangDay_().forEach(function (mon) {
    var bais = [], seen = {};
    tv.forEach(function (r) {   // gom cột Căn cứ của mọi hoạt động trong bài (mỗi hoạt động ứng với 1 nội dung tiết)
      if (r.Khoi !== cd.Khoi || r.BoSach !== cd.BoSach || r.HocKy !== hk.HocKy || r.Mon !== mon.TenMon) return;
      if (!seen[r.Bai]) { seen[r.Bai] = { bai: r.Bai, tuan: r.Tuan, canCu: '' }; bais.push(seen[r.Bai]); }
      if (r.CanCu && seen[r.Bai].canCu.indexOf(r.CanCu) < 0) seen[r.Bai].canCu += (seen[r.Bai].canCu ? ' | ' : '') + r.CanCu;
    });
    var cua = lich.filter(function (r) { return r.MaHK === maHK && r.Mon === mon.TenMon && r.CachGhep !== 'ghép tay'; })
      .sort(function (a, b) { return key_(a.Ngay) - key_(b.Ngay) || +a.Tiet - +b.Tiet; });
    if (!bais.length) { cua.forEach(function (r) { r.Bai = ''; r.CachGhep = 'thư viện tạm'; }); return; }
    // các tiết "Củng cố" không làm lệch thứ tự bài: ghép riêng, chỉ nhận kết quả khi khớp chắc chắn
    var chinh = cua.filter(function (r) { return !/CỦNG CỐ/.test(r.MonSBG); }), cungCo = cua.filter(function (r) { return /CỦNG CỐ/.test(r.MonSBG); });
    var kq = ghepTietVoiBai_(chinh.map(function (r) { return { tuan: +r.Tuan, ten: r.TenBaiDay }; }), bais);
    chinh.forEach(function (r, i) { r.Bai = kq[i].bai; r.CachGhep = kq[i].cach; });
    cungCo.forEach(function (r) {   // củng cố: lấy bài của tiết chính gần nhất trước đó trong cùng tuần
      var truoc = chinh.filter(function (c) { return key_(c.Ngay) * 100 + +c.Tiet <= key_(r.Ngay) * 100 + +r.Tiet && c.Tuan === r.Tuan; }).pop();
      r.Bai = truoc ? truoc.Bai : ''; r.CachGhep = truoc ? 'củng cố' : '';
    });
    cua.forEach(function (r) { dem.tong++; if (!r.Bai) dem.thieu++; else if (r.CachGhep === 'đoán') dem.doan++; else dem.ghep++; });
  });
  ghiTatCa_('LichBaoGiang', lich.map(function (r) { return hang_('LichBaoGiang', r); }));
  if (!dem.tong) return '';
  return 'Ghép tiết với bài: ' + dem.ghep + ' tiết tự động' + (dem.doan ? ', ' + dem.doan + ' tiết app đoán (cô nên kiểm tra)' : '') +
    (dem.thieu ? ', ' + dem.thieu + ' tiết chưa ghép' : '') + '.';
}
function layGhepTiet(mon) {
  var hk = hocKyHienTai_(), cd = caiDat_(); if (!hk) return { tiet: [], bai: [] };
  var bai = [], seen = {};
  doc_('ThuVien').forEach(function (r) {
    if (r.Khoi === cd.Khoi && r.BoSach === cd.BoSach && r.HocKy === hk.HocKy && r.Mon === mon && !seen[r.Bai]) { seen[r.Bai] = 1; bai.push(r.Bai); }
  });
  var tiet = doc_('LichBaoGiang').filter(function (r) { return r.MaHK === hk.MaHK && r.Mon === mon; })
    .sort(function (a, b) { return key_(a.Ngay) - key_(b.Ngay) || +a.Tiet - +b.Tiet; })
    .map(function (r) { return { ngay: r.Ngay, tiet: r.Tiet, tuan: r.Tuan, monSBG: r.MonSBG, ten: r.TenBaiDay, bai: r.Bai, cach: r.CachGhep }; });
  return { tiet: tiet, bai: bai };
}
function luuGhepTay(mon, ds) {    // ds: [{ngay, tiet, bai}]
  return voiKhoa_(function () {
    var hk = hocKyHienTai_(), m = {};
    ds.forEach(function (x) { m[x.ngay + '#' + x.tiet] = x.bai; });
    var lich = doc_('LichBaoGiang');
    lich.forEach(function (r) {
      var k = r.Ngay + '#' + r.Tiet;
      if (r.MaHK === hk.MaHK && r.Mon === mon && m.hasOwnProperty(k) && m[k] !== r.Bai) { r.Bai = m[k]; r.CachGhep = 'ghép tay'; }
    });
    ghiTatCa_('LichBaoGiang', lich.map(function (r) { return hang_('LichBaoGiang', r); }));
    return 'Đã lưu ghép bài cho môn ' + mon + '.';
  });
}

// ============================================================================ SỔ THEO DÕI
function layMonDangDay() { return monDangDay_(); }
function layTietTheoNgay(ngay) {
  var hk = hocKyHienTai_(); if (!hk) return [];
  var monCo = {}; monDangDay_().forEach(function (m) { monCo[m.TenMon] = 1; });
  return doc_('LichBaoGiang').filter(function (r) { return r.MaHK === hk.MaHK && r.Ngay === ngay && monCo[r.Mon]; })
    .sort(function (a, b) { return +a.Tiet - +b.Tiet; })
    .map(function (r) { return { tiet: r.Tiet, buoi: r.Buoi, mon: r.Mon, monSBG: r.MonSBG, ten: r.TenBaiDay, bai: r.Bai, tuan: r.Tuan }; });
}
function layDanhSachBai(mon) {
  var hk = hocKyHienTai_(), cd = caiDat_(); if (!hk) return [];
  var out = [], seen = {};
  doc_('ThuVien').forEach(function (r) {
    if (r.Khoi === cd.Khoi && r.BoSach === cd.BoSach && r.HocKy === hk.HocKy && r.Mon === mon && !seen[r.Bai]) { seen[r.Bai] = 1; out.push({ bai: r.Bai, tuan: r.Tuan }); }
  });
  return out;
}
// Bộ nhận xét của 1 bài: {hoatDong:[{ten, muc:{mức:[câu]}}], tam:bool}
function layBoNhanXet(mon, bai) {
  var bt = baiThang_(bai); if (bt) return boThang_(mon, bt.thang, bt.noiDung);
  var hk = hocKyHienTai_(), cd = caiDat_(), hd = {}, thuTu = [];
  if (hk && bai) doc_('ThuVien').forEach(function (r) {
    if (r.Khoi === cd.Khoi && r.BoSach === cd.BoSach && r.HocKy === hk.HocKy && r.Mon === mon && r.Bai === bai && r.MucDo) {
      if (!hd[r.HoatDong]) { hd[r.HoatDong] = {}; thuTu.push(r.HoatDong); }
      (hd[r.HoatDong][r.MucDo] = hd[r.HoatDong][r.MucDo] || []).push(r.NoiDung);
    }
  });
  if (!thuTu.length) return { hoatDong: [{ ten: 'Học tập chung', muc: THU_VIEN_TAM }], tam: true, muc: MUC4 };
  return { hoatDong: thuTu.map(function (t) { return { ten: t, muc: hd[t] }; }), tam: false, muc: MUC4 };
}
function layNhanXetChung() {
  var nhom = {}, thuTu = [];
  doc_('NhanXetChung').forEach(function (r) {
    if (!nhom[r.Nhom]) { nhom[r.Nhom] = { ten: r.Nhom, pcnl: r.PCNL, muc: {} }; thuTu.push(r.Nhom); }
    (nhom[r.Nhom].muc[r.MucDo] = nhom[r.Nhom].muc[r.MucDo] || []).push(r.NoiDung);
  });
  return { nhom: thuTu.map(function (t) { return nhom[t]; }), muc: MUC_PC };
}
// Học sinh + nhận xét đã ghi cho (môn, bài) để hiện dấu trên từng em
// Sổ theo dõi lớn dần cả năm: chỉ đọc 2 cột Môn, Bài để tìm đoạn dòng của bài, rồi đọc riêng đoạn đó
function docSTDTheoBai_(mon, bai) {
  var sh = sheetNho_('SoTheoDoi'), n = sh.getLastRow(), cot = BANG.SoTheoDoi, cMon = cot.indexOf('Mon');
  if (n < 2) return [];
  var mb = _NHO._monBaiSTD || (_NHO._monBaiSTD = sh.getRange(2, cMon + 1, n - 1, 2).getDisplayValues()), dau = -1, cuoi = -1;
  mb.forEach(function (r, i) { if (r[0] === mon && r[1] === bai) { if (dau < 0) dau = i; cuoi = i; } });
  if (dau < 0) return [];
  return sh.getRange(dau + 2, 1, cuoi - dau + 1, cot.length).getDisplayValues().map(function (r) {
    var o = {}; cot.forEach(function (c, i) { o[c] = r[i]; }); return o;
  });
}
function layHocSinhChoBai(ngay, mon, bai) {
  var ds = hocSinhDangHoc_(ngay || homNay_()), da = {};
  (_NHO.SoTheoDoi ? doc_('SoTheoDoi') : docSTDTheoBai_(mon, bai)).forEach(function (r) {
    if (r.Mon === mon && r.Bai === bai) (da[r.MaDinhDanh] = da[r.MaDinhDanh] || []).push({ id: r.ID, hd: r.HoatDong, muc: r.MucDo, nd: r.NoiDung, gc: r.GhiChu, ngay: r.Ngay, luc: r.ThoiGian });
  });
  return ds.map(function (h) { return { ma: h.MaDinhDanh, ten: h.HoTen, da: da[h.MaDinhDanh] || [] }; });
}
// x: {ngay, tiet, mon, bai, hoatDong, mucDo, noiDung, ghiChu, hocSinh:[mã]}
function luuNhanXet(x) {
  return voiKhoa_(function () {
    if (!x.hocSinh || !x.hocSinh.length) throw new Error('Chưa chọn học sinh.');
    if (!x.mucDo) throw new Error('Chưa chọn mức.');
    var hk = hocKyHienTai_(), ngay = x.ngay || homNay_(), luc = Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm:ss');
    var tuan = hk ? (doc_('TuanHoc').filter(function (t) { return t.MaHK === hk.MaHK && key_(t.TuNgay) <= key_(ngay) && key_(ngay) <= key_(t.DenNgay) + 2; })[0] || {}).Tuan : '';
    var rows = x.hocSinh.map(function (ma) {
      return hang_('SoTheoDoi', { ID: Utilities.getUuid().slice(0, 8), ThoiGian: luc, Ngay: ngay, MaHK: hk ? hk.MaHK : '', Tuan: tuan || '',
        Thang: +ngay.split('/')[1], Tiet: x.tiet || '', MaDinhDanh: ma, Mon: x.mon, Bai: x.bai || '', HoatDong: x.hoatDong,
        MucDo: x.mucDo, NoiDung: x.noiDung || '', GhiChu: x.ghiChu || '' });
    });
    themDong_('SoTheoDoi', rows);
    // trả về mã từng dòng để trình duyệt có thể "Hoàn tác" hoặc xoá lẻ
    return { msg: 'Đã lưu nhận xét cho ' + rows.length + ' học sinh.', ids: rows.map(function (r) { return r[0]; }), luc: luc };
  });
}
function xoaNhanXet(id) { return xoaNhieuNhanXet([id]); }
function xoaNhieuNhanXet(ids) {
  return voiKhoa_(function () {
    var bo = {}; ids.forEach(function (i) { bo[i] = 1; });
    var truoc = doc_('SoTheoDoi').length;
    thayTheTheo_('SoTheoDoi', function (r) { return !bo[r.ID]; }, []);
    var n = truoc - doc_('SoTheoDoi').length;
    return 'Đã xoá ' + n + ' nhận xét.';
  });
}
// ============================================================================ NHÓM HỌC SINH CẦN QUAN TÂM (theo ý cô: gọn, 3 nhóm sẵn)
var NHOM_HS = ['Bồi dưỡng', 'Cần rèn thêm', 'Cần chú ý'];
function layNhomHS() {
  var ds = {};
  doc_('NhomHS').forEach(function (r) {
    var n = String(r.Nhom || '').split(';').filter(function (x) { return x; });
    if (n.length || r.GhiChu) ds[r.MaDinhDanh] = { nhom: n, gc: r.GhiChu };
  });
  return { nhom: NHOM_HS, ds: ds };
}
function suaNhomHS_(ma, fn) {
  var luc = Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm');
  var cu = doc_('NhomHS').filter(function (r) { return r.MaDinhDanh === ma; })[0] || { MaDinhDanh: ma, Nhom: '', GhiChu: '', BoGoiY: '' };
  fn(cu); cu.CapNhat = luc;
  thayTheTheo_('NhomHS', function (r) { return r.MaDinhDanh !== ma; }, [hang_('NhomHS', cu)]);
}
function luuNhomHS(ma, nhom, gc) {
  return voiKhoa_(function () {
    suaNhomHS_(ma, function (r) { r.Nhom = (nhom || []).filter(function (x) { return NHOM_HS.indexOf(x) >= 0; }).join(';'); r.GhiChu = String(gc || '').trim().slice(0, 150); });
    return layNhomHS();
  });
}
function thangKhoa_() { var p = homNay_().split('/'); return p[2] + '-' + p[1]; }
function boQuaGoiYNhom(ma, nhom) {
  return voiKhoa_(function () {
    suaNhomHS_(ma, function (r) { var b = String(r.BoGoiY || '').split(' ').filter(function (x) { return x && x.split('|')[1] === thangKhoa_(); });
      b.push(nhom.replace(/ /g, '_') + '|' + thangKhoa_()); r.BoGoiY = b.join(' '); });
    return 'Đã bỏ qua gợi ý (tháng sau app xem lại).';
  });
}
// App gợi ý – cô xác nhận mới thêm vào nhóm. Căn cứ: đánh giá tháng gần nhất, vắng 30 ngày qua, nộp vở 30 ngày qua.
function layGoiYNhom() {
  var hn = homNay_(), k = key_(hn), th = +hn.split('/')[1], truoc = th === 1 ? 12 : th - 1, t = null;
  [th, truoc].some(function (x) {
    var d = duLieuThang_(x);
    var co = Object.keys(d.du).some(function (ma) { return ['Tiếng Việt', 'Toán'].some(function (m) { var o = d.du[ma][m]; return o && (o.mucKT || o.mucKN); }); });
    if (co) t = d; return co;
  });
  var tu = toDate_(hn); tu.setDate(tu.getDate() - 30); var k30 = key_(Utilities.formatDate(tu, TZ, 'dd/MM/yyyy'));
  var vang = {};
  doc_('DiemDanh').forEach(function (r) {
    if (r.TrangThai === 'Có mặt' || key_(r.Ngay) < k30 || key_(r.Ngay) > k) return;
    vang[r.MaDinhDanh] = (vang[r.MaDinhDanh] || 0) + (r.Buoi === 'Sáng' || r.Buoi === 'Chiều' ? 0.5 : 1);
  });
  var vo = demVo_(function (ngay) { var x = key_(ngay); return x >= k30 && x <= k; }).vo;
  var nhom = layNhomHS().ds, bo = {}, kh = thangKhoa_();
  doc_('NhomHS').forEach(function (r) { String(r.BoGoiY || '').split(' ').forEach(function (x) { var p = x.split('|'); if (p[1] === kh) bo[r.MaDinhDanh + '|' + p[0].replace(/_/g, ' ')] = 1; }); });
  var THAP = { CHT: 1, CCG: 1 }, ten = { kt: 'kiến thức', kn: 'kỹ năng' }, out = [];
  hocSinhDangHoc_(hn).forEach(function (h) {
    var ma = h.MaDinhDanh, ly = { 'Bồi dưỡng': [], 'Cần rèn thêm': [], 'Cần chú ý': [] };
    if (t) ['Tiếng Việt', 'Toán'].forEach(function (m) {
      var o = (t.du[ma] || {})[m]; if (!o) return;
      if (o.mucKT === 'HTXS' && o.mucKN === 'HTXS') ly['Bồi dưỡng'].push(m + ' tháng ' + t.thang + ': HTXS cả kiến thức và kỹ năng');
      var yeu = [['kt', o.mucKT], ['kn', o.mucKN]].filter(function (x) { return THAP[x[1]]; });
      if (yeu.length) ly['Cần rèn thêm'].push(m + ' tháng ' + t.thang + ': ' + yeu.map(function (x) { return ten[x[0]] + ' ' + x[1]; }).join(', '));
    });
    if (t) [['NLC', 'năng lực chung'], ['PC', 'phẩm chất']].forEach(function (p) { var o = (t.du[ma] || {})[p[0]]; if (o && o.muc === 'Cần cố gắng') ly['Cần chú ý'].push('Tháng ' + t.thang + ': ' + p[1] + ' cần cố gắng'); });
    MON_VO.forEach(function (m) {
      var v = (vo[ma] || {})[m]; if (!v) return;
      if (v.E >= 3) ly['Cần rèn thêm'].push('Vở ' + m + ': chưa đúng – chưa đẹp ' + v.E + ' lần (30 ngày qua)');
      if (v.N + v.cu >= 3) ly['Cần chú ý'].push('Chưa nộp vở ' + m + ' ' + (v.N + v.cu) + ' lần (30 ngày qua)');
    });
    if ((vang[ma] || 0) >= 3) ly['Cần chú ý'].push('Vắng ' + vang[ma] + ' buổi trong 30 ngày qua');
    NHOM_HS.forEach(function (n) {
      if (!ly[n].length || bo[ma + '|' + n] || ((nhom[ma] || {}).nhom || []).indexOf(n) >= 0) return;
      out.push({ ma: ma, ten: h.HoTen, nhom: n, lyDo: ly[n] });
    });
  });
  return { goiY: out, thang: t ? t.thang : '' };
}

function layHoSo(ma) {
  var h = doc_('HocSinh').filter(function (x) { return x.MaDinhDanh === ma; })[0];
  var nx = doc_('SoTheoDoi').filter(function (r) { return r.MaDinhDanh === ma; })
    .sort(function (a, b) { return key_(b.Ngay) - key_(a.Ngay) || (a.ThoiGian < b.ThoiGian ? 1 : -1); });
  var dd = doc_('DiemDanh').filter(function (r) { return r.MaDinhDanh === ma && r.TrangThai !== 'Có mặt'; });
  return { hs: h, nhanXet: nx, vang: dd, phongTrao: phongTraoCuaHS_(caiDat_().NamHoc || '')[ma] || [] };
}

// ============================================================================ NHẬN XÉT THÁNG (biểu mẫu đánh giá thường xuyên)
// App GHÉP sẵn từ các câu cô đã chọn trong sổ theo dõi của tháng (không dùng AI). Cô sửa, rồi xuất vào file mẫu của trường.
var GIOI_HAN_THANG = 250;
var TEN_SHEET_TONG_HOP = 'Sổ tổng hợp các môn';
var DIEM_MUC = { 'Hoàn thành xuất sắc': 4, 'Hoàn thành tốt': 3, 'Hoàn thành': 2, 'Chưa hoàn thành': 1, 'Tốt': 3, 'Đạt': 2, 'Cần cố gắng': 1 };
// Phân loại cột "PC/NL liên quan" của thư viện vào 3 nhóm của biểu mẫu
var NL_CHUNG = [['tự chủ', 'Tự chủ và tự học'], ['tự học', 'Tự chủ và tự học'], ['giao tiếp và hợp tác', 'Giao tiếp và hợp tác'],
                ['hợp tác', 'Giao tiếp và hợp tác'], ['giải quyết vấn đề và sáng tạo', 'Giải quyết vấn đề và sáng tạo'], ['sáng tạo', 'Giải quyết vấn đề và sáng tạo']];
var NL_DAC_THU = [['ngôn ngữ', 'Ngôn ngữ'], ['văn học', 'Ngôn ngữ'], ['tính toán', 'Tính toán'], ['khoa học', 'Khoa học'],
                  ['công nghệ', 'Công nghệ'], ['tin học', 'Tin học'], ['thẩm mĩ', 'Thẩm mĩ'], ['thể chất', 'Thể chất']];
var PHAM_CHAT = [['yêu nước', 'Yêu nước'], ['nhân ái', 'Nhân ái'], ['chăm chỉ', 'Chăm chỉ'], ['trung thực', 'Trung thực'], ['trách nhiệm', 'Trách nhiệm']];
function phanLoaiPCNL_(pcnl) {
  var kq = { chung: [], dacThu: [], pc: [] };
  var them = function (ds, ten) { if (ds.indexOf(ten) < 0) ds.push(ten); };
  String(pcnl || '').split(';').forEach(function (phan) {
    var x = phan.toLowerCase().trim(); if (!x) return;
    if (x.indexOf('toán') >= 0) { them(kq.dacThu, 'Tính toán'); return; }   // năng lực toán học (kể cả "GQVĐ toán học") là năng lực đặc thù
    var tim = function (bang, ds) { for (var i = 0; i < bang.length; i++) if (x.indexOf(bang[i][0]) >= 0) { them(ds, bang[i][1]); return true; } return false; };
    if (!tim(NL_DAC_THU, kq.dacThu) && !tim(NL_CHUNG, kq.chung)) tim(PHAM_CHAT, kq.pc);
  });
  return kq;
}
function tachCau_(nd) {
  return String(nd || '').split(/(?<=[.!?])\s+/).map(function (c) { return c.trim(); }).filter(Boolean)
    .map(function (c) { return /[.!?]$/.test(c) ? c : c + '.'; });
}
function noiCau_(ds, gioiHan) {        // nối các câu, không vượt giới hạn ký tự
  var out = '';
  ds.forEach(function (c) { var thu = out ? out + ' ' + c : c; if (thu.length <= gioiHan) out = thu; });
  return out;
}
// Câu "chung chung" (xuất hiện ở từ 3 bài trở lên trong thư viện, VD "Làm đúng các bài tập.") xếp sau câu nêu nội dung cụ thể của bài.
var _CAU_CHUNG = {};
function chonCau_(ghi) {               // ghi: [{muc, nd}] theo thứ tự thời gian → {tot:[], can:[], tb, tiLeYeu}
  var dem = {}, tong = 0, yeu = 0;
  ghi.forEach(function (g, i) {
    var d = DIEM_MUC[g.muc] || 2; tong += d; if (d <= 2) yeu++;
    tachCau_(g.nd).forEach(function (c) {
      var k = dem[c] || (dem[c] = { cau: c, n: 0, d: d, cuoi: 0 });
      k.n++; k.d = d; k.cuoi = i;
    });
  });
  var ds = Object.keys(dem).map(function (k) { return dem[k]; }).sort(function (a, b) {
    return (_CAU_CHUNG[a.cau] ? 1 : 0) - (_CAU_CHUNG[b.cau] ? 1 : 0) || b.n - a.n || b.cuoi - a.cuoi;
  });
  return { tot: ds.filter(function (x) { return x.d >= 3; }).map(function (x) { return x.cau; }),
           can: ds.filter(function (x) { return x.d <= 1; }).concat(ds.filter(function (x) { return x.d === 2; })),
           tb: ghi.length ? tong / ghi.length : 0, tiLeYeu: ghi.length ? yeu / ghi.length : 0 };
}
// Nhận xét môn học trong tháng: nêu điểm mạnh trước (tối đa 2 câu); chỉ thêm câu "cần cố gắng"
// khi phần chưa tốt chiếm từ 1/4 số lần ghi trở lên hoặc lặp lại – tránh 1 lần ghi yếu làm lệch cả tháng.
function ghepNhanXetMon_(ghi, gioiHan) {
  if (!ghi.length) return '';
  var c = chonCau_(ghi), ds;
  var can = c.can.filter(function (x) { return c.tiLeYeu >= 0.25 || x.n >= 2; }).map(function (x) { return x.cau; });
  if (c.tb >= 2.5) ds = c.tot.slice(0, 2).concat(can.slice(0, 1));
  else ds = c.can.map(function (x) { return x.cau; }).slice(0, 2).concat(c.tot.slice(0, 1));
  if (!ds.length) ds = c.tot.concat(c.can.map(function (x) { return x.cau; })).slice(0, 2);
  return noiCau_(ds, gioiHan);
}
function cauTieuBieu_(c) {             // 1 câu đại diện: tháng tốt lấy câu tốt, tháng yếu lấy câu cần cố gắng
  var can = c.can.map(function (x) { return x.cau; });
  return (c.tb >= 2.5 ? c.tot : can)[0] || c.tot[0] || can[0] || '';
}
function tuMuc_(tb) { return tb >= 2.75 ? 'Tốt' : (tb >= 1.75 ? 'Đạt' : 'Cần cố gắng'); }
function chuThuong_(c) { return c.charAt(0).toLowerCase() + c.slice(1); }
// Nhận xét năng lực/phẩm chất: "Ngôn ngữ: Tốt – đọc đúng, rõ ràng." nối nhiều mục
function ghepNangLuc_(theoMuc, thuTu, gioiHan, tienTo) {
  var phan = [];
  thuTu.forEach(function (ten) {
    var ghi = theoMuc[ten]; if (!ghi || !ghi.length) return;
    var c = chonCau_(ghi), cau = cauTieuBieu_(c);
    phan.push((tienTo || '') + (tienTo ? chuThuong_(ten) : ten) + ': ' + tuMuc_(c.tb) + (cau ? ' – ' + chuThuong_(cau) : '.'));
  });
  return noiCau_(phan, gioiHan);
}
var THU_TU_CHUNG = ['Tự chủ và tự học', 'Giao tiếp và hợp tác', 'Giải quyết vấn đề và sáng tạo'];
var THU_TU_DAC_THU = ['Ngôn ngữ', 'Tính toán', 'Khoa học', 'Công nghệ', 'Tin học', 'Thẩm mĩ', 'Thể chất'];
var THU_TU_PC = ['Yêu nước', 'Nhân ái', 'Chăm chỉ', 'Trung thực', 'Trách nhiệm'];
// ghi: các dòng sổ theo dõi (đã gắn .pl = phanLoaiPCNL_) của 1 học sinh trong phạm vi cần tổng hợp
function nhanXetNangLuc_(ghi, gioiHan) {
  var nhom = { chung: {}, dacThu: {}, pc: {} };
  ghi.forEach(function (g) {
    ['chung', 'dacThu', 'pc'].forEach(function (k) { g.pl[k].forEach(function (ten) { (nhom[k][ten] = nhom[k][ten] || []).push(g); }); });
  });
  return { nlc: ghepNangLuc_(nhom.chung, THU_TU_CHUNG, gioiHan), nld: ghepNangLuc_(nhom.dacThu, THU_TU_DAC_THU, gioiHan, 'Năng lực '),
           pc: ghepNangLuc_(nhom.pc, THU_TU_PC, gioiHan) };
}
function namCuaThang_(namHoc, thang) { var y = /(\d{4})-(\d{4})/.exec(namHoc || ''); return y ? (+thang >= 8 ? +y[1] : +y[2]) : new Date().getFullYear(); }

function layBieuMauThang(thang) {
  thang = +thang;
  var cd = caiDat_(), namHoc = cd.NamHoc || '', nam = namCuaThang_(namHoc, thang);
  var dauThang = '01/' + ('0' + thang).slice(-2) + '/' + nam;
  var mon = monDangDay_();
  // PC/NL của từng hoạt động (theo thư viện) và của từng nhóm nhận xét chung
  var pcnl = {};
  doc_('ThuVien').forEach(function (r) { if (r.Khoi === cd.Khoi) pcnl[r.Mon + '|' + r.Bai + '|' + r.HoatDong] = r.PCNL; });
  doc_('NhanXetChung').forEach(function (r) { pcnl['Chung||' + r.Nhom] = r.PCNL; });
  var baiCuaCau = {}; _CAU_CHUNG = {};
  doc_('ThuVien').forEach(function (r) {
    if (r.Khoi !== cd.Khoi) return;
    tachCau_(r.NoiDung).forEach(function (c) { var k = r.Mon + '|' + c; (baiCuaCau[k] = baiCuaCau[k] || {})[r.Bai] = 1; });
  });
  Object.keys(baiCuaCau).forEach(function (k) { if (Object.keys(baiCuaCau[k]).length >= 3) _CAU_CHUNG[k.split('|').slice(1).join('|')] = 1; });
  var hs = doc_('HocSinh').filter(function (h) {
    return h.TrangThai !== 'Đã chuyển đi' || (h.NgayChuyenDi && key_(h.NgayChuyenDi) > key_(dauThang));
  }).sort(function (a, b) { return +a.ThuTu - +b.ThuTu; });
  var ghiTheoHS = {};
  doc_('SoTheoDoi').forEach(function (r) {
    if (+r.Thang !== thang || String(r.MaHK).indexOf(namHoc) !== 0) return;
    var pl = phanLoaiPCNL_(pcnl[r.Mon + '|' + r.Bai + '|' + r.HoatDong] || '');
    (ghiTheoHS[r.MaDinhDanh] = ghiTheoHS[r.MaDinhDanh] || []).push({ muc: r.MucDo, nd: r.NoiDung, mon: r.Mon, pl: pl, ngay: r.Ngay });
  });
  var daSua = {};
  doc_('NhanXetThang').forEach(function (r) { if (r.NamHoc === namHoc && +r.Thang === thang) daSua[r.MaDinhDanh + '|' + r.Sheet] = r; });
  var sheets = [{ ten: TEN_SHEET_TONG_HOP, mon: '' }].concat(mon.map(function (m) { return { ten: m.TenSheetBieuMau || m.TenMon, mon: m.TenMon }; }));
  var du = {};
  hs.forEach(function (h) {
    var ghi = (ghiTheoHS[h.MaDinhDanh] || []).sort(function (a, b) { return key_(a.ngay) - key_(b.ngay); });
    du[h.MaDinhDanh] = {};
    sheets.forEach(function (s) {
      var o;
      if (!s.mon) {     // Sổ tổng hợp: mỗi môn 1 câu tiêu biểu; năng lực, phẩm chất từ mọi ghi nhận (kể cả nhận xét chung)
        var phan = mon.map(function (m) {
          var g = ghi.filter(function (x) { return x.mon === m.TenMon; }); if (!g.length) return '';
          var c = chonCau_(g), cau = cauTieuBieu_(c);
          return cau ? m.TenMon + ': ' + chuThuong_(cau) : '';
        }).filter(Boolean);
        var nl = nhanXetNangLuc_(ghi, GIOI_HAN_THANG);
        o = { mon: noiCau_(phan, GIOI_HAN_THANG), nlc: nl.nlc, nld: nl.nld, pc: nl.pc, soGhi: ghi.length };
      } else {
        var gm = ghi.filter(function (x) { return x.mon === s.mon; });
        var nl2 = nhanXetNangLuc_(gm, GIOI_HAN_THANG);
        o = { mon: ghepNhanXetMon_(gm, GIOI_HAN_THANG), nlc: nl2.nlc, nld: nl2.nld, pc: nl2.pc, soGhi: gm.length };
      }
      o.goc = { mon: o.mon, nlc: o.nlc, nld: o.nld, pc: o.pc };       // bản app ghép (để "↺ Lấy lại bản app")
      var sua = daSua[h.MaDinhDanh + '|' + s.ten];
      if (sua) { o.mon = sua.NXMon; o.nlc = sua.NLChung; o.nld = sua.NLDacThu; o.pc = sua.PhamChat; o.sua = true; o.luc = sua.CapNhat; }
      du[h.MaDinhDanh][s.ten] = o;
    });
  });
  return { namHoc: namHoc, thang: thang, nam: nam, gioiHan: GIOI_HAN_THANG, sheets: sheets,
           hs: hs.map(function (h, i) { return { ma: h.MaDinhDanh, ten: h.HoTen, ns: h.NgaySinh, stt: i + 1 }; }), du: du };
}
// Lưu phần cô sửa của 1 em ở 1 sheet (x: {ma, sheet, mon, nlc, nld, pc})
function luuNhanXetThang(thang, x) {
  return voiKhoa_(function () {
    var namHoc = caiDat_().NamHoc || '', luc = Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm');
    var cung = function (r) { return r.NamHoc === namHoc && +r.Thang === +thang && r.MaDinhDanh === x.ma && r.Sheet === x.sheet; };
    thayTheTheo_('NhanXetThang', function (r) { return !cung(r); }, [hang_('NhanXetThang', {
      NamHoc: namHoc, Thang: thang, MaDinhDanh: x.ma, Sheet: x.sheet, NXMon: x.mon, NLChung: x.nlc, NLDacThu: x.nld, PhamChat: x.pc, CapNhat: luc })]);
    return luc;
  });
}
function boSuaThang(thang, ma, sheet) {
  return voiKhoa_(function () {
    var namHoc = caiDat_().NamHoc || '';
    thayTheTheo_('NhanXetThang', function (r) { return !(r.NamHoc === namHoc && +r.Thang === +thang && r.MaDinhDanh === ma && r.Sheet === sheet); }, []);
    return 'Đã lấy lại bản app ghép.';
  });
}
// Lưu bản sao file biểu mẫu đã xuất vào Google Drive của cô: App Sổ theo dõi / Biểu mẫu đã xuất / <năm học>
function luuFileXuat(tenFile, base64) {
  var namHoc = caiDat_().NamHoc || 'Chưa rõ năm học';
  var thuMuc = function (cha, ten) {
    var it = cha ? cha.getFoldersByName(ten) : DriveApp.getFoldersByName(ten);
    return it.hasNext() ? it.next() : (cha ? cha.createFolder(ten) : DriveApp.createFolder(ten));
  };
  var dich = thuMuc(thuMuc(thuMuc(null, 'App Sổ theo dõi'), 'Biểu mẫu đã xuất'), namHoc);
  var f = dich.createFile(Utilities.newBlob(Utilities.base64Decode(base64), MimeType.MICROSOFT_EXCEL, tenFile));
  return { url: f.getUrl(), ten: tenFile, thuMuc: dich.getUrl() };
}

// ============================================================================ ĐÁNH GIÁ ĐỊNH KỲ (GK1, CK1, GK2, CK2)
// Gợi ý mức môn học (T/H/C) và 15 mức năng lực – phẩm chất (T/Đ/C) từ sổ theo dõi; ghép nhận xét ≤ 500 ký tự.
// Cô chốt mức, nhập điểm KTĐK (cuối kì), sửa nhận xét, rồi xuất vào 2 file mẫu của Bộ.
var GIOI_HAN_KY = 500;
var DS_NLPC = THU_TU_CHUNG.concat(THU_TU_DAC_THU, THU_TU_PC);       // đúng thứ tự cột E..S của file năng lực – phẩm chất
// Phạm vi ngày của 1 kì: GK = đầu học kì → hết tuần ôn tập giữa kì; CK = cả học kì
function phamViKy_(ky) {
  var cd = caiDat_(), hocKy = /1$/.test(ky) ? 'I' : 'II', namHoc = cd.NamHoc || '';
  var hk = doc_('HocKy').filter(function (h) { return h.NamHoc === namHoc && h.HocKy === hocKy; })[0];
  if (!namHoc) throw new Error('Chưa đặt năm học – vào Cài đặt → Thông tin chung (hoặc tải danh sách học sinh) trước.');
  if (!hk) throw new Error('Chưa có Sổ báo giảng học kì ' + hocKy + ' của năm học ' + namHoc + ' – cần tải ở Cài đặt trước khi đánh giá ' + ky + '.');
  var den = hk.NgayDayCuoi;
  if (/^GK/.test(ky)) {
    var tuanOn = doc_('LichBaoGiang').filter(function (r) { return r.MaHK === hk.MaHK && /giữa (học )?k[ìỳ]/i.test(r.TenBaiDay); }).map(function (r) { return +r.Tuan; });
    var dsTuan = doc_('TuanHoc').filter(function (t) { return t.MaHK === hk.MaHK; }).sort(function (a, b) { return +a.Tuan - +b.Tuan; });
    var tuanGiua = tuanOn.length ? Math.max.apply(null, tuanOn) : (dsTuan.length ? +dsTuan[Math.ceil(dsTuan.length / 2) - 1].Tuan : 0);
    var t = dsTuan.filter(function (x) { return +x.Tuan === tuanGiua; })[0];
    if (t) den = t.DenNgay;
  }
  return { maHK: hk.MaHK, hocKy: hocKy, namHoc: namHoc, tu: hk.NgayBatDau, den: den, cuoiKy: /^CK/.test(ky) };
}
function mucMon_(tb) { return !tb ? '' : (tb >= 2.8 ? 'T' : (tb >= 1.8 ? 'H' : 'C')); }
function mucNLPC_(tb) { return !tb ? '' : (tb >= 2.75 ? 'T' : (tb >= 1.75 ? 'Đ' : 'C')); }
function tbDiem_(ghi) { if (!ghi.length) return 0; var s = 0; ghi.forEach(function (g) { s += DIEM_MUC[g.muc] || 2; }); return s / ghi.length; }
// Nhận xét môn cả kì: tối đa 3 câu mạnh + 1 câu cần cố gắng (cùng quy tắc với nhận xét tháng)
function ghepNhanXetKy_(ghi, gioiHan) {
  if (!ghi.length) return '';
  var c = chonCau_(ghi);
  var can = c.can.filter(function (x) { return c.tiLeYeu >= 0.25 || x.n >= 2; }).map(function (x) { return x.cau; });
  var ds = c.tb >= 2.5 ? c.tot.slice(0, 3).concat(can.slice(0, 1)) : c.can.map(function (x) { return x.cau; }).slice(0, 2).concat(c.tot.slice(0, 2));
  return noiCau_(ds, gioiHan);
}
function layDanhGiaKy(ky) {
  if (ky === 'GVCN') return layNhanXetGVCN_();
  var pv = phamViKy_(ky), cd = caiDat_(), mon = monDangDay_();
  var pcnl = {}, baiCuaCau = {}; _CAU_CHUNG = {};
  doc_('ThuVien').forEach(function (r) {
    if (r.Khoi !== cd.Khoi) return;
    pcnl[r.Mon + '|' + r.Bai + '|' + r.HoatDong] = r.PCNL;
    tachCau_(r.NoiDung).forEach(function (c) { var k = r.Mon + '|' + c; (baiCuaCau[k] = baiCuaCau[k] || {})[r.Bai] = 1; });
  });
  Object.keys(baiCuaCau).forEach(function (k) { if (Object.keys(baiCuaCau[k]).length >= 3) _CAU_CHUNG[k.split('|').slice(1).join('|')] = 1; });
  doc_('NhanXetChung').forEach(function (r) { pcnl['Chung||' + r.Nhom] = r.PCNL; });
  var hs = doc_('HocSinh').filter(function (h) {
    return h.TrangThai !== 'Đã chuyển đi' || (h.NgayChuyenDi && key_(h.NgayChuyenDi) > key_(pv.den));
  }).sort(function (a, b) { return +a.ThuTu - +b.ThuTu; });
  var gy = goiYTuThang_(pv, hs, mon);           // có đánh giá tháng trong kì → gợi ý từ các tháng; không thì dùng sổ theo dõi theo bài
  var ghiTheoHS = {};
  if (!gy) doc_('SoTheoDoi').forEach(function (r) {
    if (r.MaHK !== pv.maHK || key_(r.Ngay) < key_(pv.tu) || key_(r.Ngay) > key_(pv.den)) return;
    (ghiTheoHS[r.MaDinhDanh] = ghiTheoHS[r.MaDinhDanh] || []).push({ muc: r.MucDo, nd: r.NoiDung, mon: r.Mon, ngay: r.Ngay,
      pl: phanLoaiPCNL_(pcnl[r.Mon + '|' + r.Bai + '|' + r.HoatDong] || '') });
  });
  var anhBai = {};
  doc_('BaiKiemTra').forEach(function (r) { if (r.NamHoc === pv.namHoc && r.Ky === ky) anhBai[r.MaDinhDanh + '|' + r.Mon] = r.Url; });
  var daSua = {};
  doc_('DanhGiaKy').forEach(function (r) { if (r.NamHoc === pv.namHoc && r.Ky === ky) daSua[r.MaDinhDanh + '|' + r.Phan] = r; });
  var du = {};
  hs.forEach(function (h) {
    var ghi = (ghiTheoHS[h.MaDinhDanh] || []).sort(function (a, b) { return key_(a.ngay) - key_(b.ngay); });
    var o = { mon: {}, nlpc: null };
    mon.forEach(function (m) {
      var gm = ghi.filter(function (x) { return x.mon === m.TenMon; });
      var g = gy && gy.du[h.MaDinhDanh].mon[m.TenMon];
      var x = g ? { mucGoiY: g.mucGoiY, nx: g.nx, diem: '', soGhi: g.soGhi, anh: anhBai[h.MaDinhDanh + '|' + m.TenMon] || '' }
        : { mucGoiY: mucMon_(tbDiem_(gm)), nx: ghepNhanXetKy_(gm, GIOI_HAN_KY), diem: '', soGhi: gm.length, anh: anhBai[h.MaDinhDanh + '|' + m.TenMon] || '' };
      x.muc = x.mucGoiY; x.goc = { muc: x.muc, nx: x.nx, diem: '' };
      var s = daSua[h.MaDinhDanh + '|' + m.TenMon];
      if (s) { x.muc = s.Muc; x.diem = s.Diem; x.nx = s.NhanXet; x.sua = true; }
      o.mon[m.TenMon] = x;
    });
    // 15 mức năng lực – phẩm chất + 3 nhận xét
    var theo = {}; ghi.forEach(function (g) { ['chung', 'dacThu', 'pc'].forEach(function (k) { g.pl[k].forEach(function (t) { (theo[t] = theo[t] || []).push(g); }); }); });
    var gp = gy && gy.du[h.MaDinhDanh].nlpc;
    var mucGoiY = {}; DS_NLPC.forEach(function (t) { mucGoiY[t] = gp ? gp.mucGoiY[t] : mucNLPC_(tbDiem_(theo[t] || [])); });
    var nl = gp ? gp : nhanXetNangLuc_(ghi, GIOI_HAN_KY);
    var p = { mucGoiY: mucGoiY, muc: JSON.parse(JSON.stringify(mucGoiY)), nlc: nl.nlc, nld: nl.nld, pc: nl.pc, soGhi: gp ? gp.soGhi : ghi.length };
    p.goc = { muc: JSON.parse(JSON.stringify(mucGoiY)), nlc: p.nlc, nld: p.nld, pc: p.pc };
    var sp = daSua[h.MaDinhDanh + '|NLPC'];
    if (sp) { try { p.muc = JSON.parse(sp.MucNLPC || '{}'); } catch (e) {} p.nlc = sp.NXNLChung; p.nld = sp.NXNLDacThu; p.pc = sp.NXPhamChat; p.sua = true; }
    o.nlpc = p;
    du[h.MaDinhDanh] = o;
  });
  var lichSu = doc_('PhienBanNX').filter(function (r) { return r.NamHoc === pv.namHoc && r.Ky === ky; })
    .map(function (r) { return { ma: r.MaDinhDanh, phan: r.Phan, truong: r.Truong, lo: r.Lo }; });
  return { ky: ky, tu: pv.tu, den: pv.den, cuoiKy: pv.cuoiKy, namHoc: pv.namHoc, gioiHan: GIOI_HAN_KY, dsNLPC: DS_NLPC, lichSu: lichSu,
           nguon: gy ? 'thang' : 'bai', thangDung: gy ? gy.thang : [],
           nhomNLPC: { chung: THU_TU_CHUNG, dacThu: THU_TU_DAC_THU, pc: THU_TU_PC },
           mon: mon.map(function (m) { return { ten: m.TenMon, sheet: m.TenSheetBieuMau || m.TenMon, coDiem: m.CoDiemKTDK === 'Có' }; }),
           hs: hs.map(function (h, i) { return { ma: h.MaDinhDanh, ten: h.HoTen, ns: h.NgaySinh, stt: i + 1 }; }), du: du };
}
// x: {ma, phan: 'Tên môn' | 'NLPC', muc, diem, nx}  hoặc  {ma, phan:'NLPC', mucNLPC:{…}, nlc, nld, pc}
function luuDanhGiaKy(ky, x) {
  return voiKhoa_(function () {
    var namHoc = caiDat_().NamHoc || '', luc = Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm');
    if (x.diem !== '' && x.diem != null && !/^(10|[1-9])$/.test(String(x.diem))) throw new Error('Điểm KTĐK phải là số nguyên từ 1 đến 10.');
    var cung = function (r) { return r.NamHoc === namHoc && r.Ky === ky && r.MaDinhDanh === x.ma && r.Phan === x.phan; };
    thayTheTheo_('DanhGiaKy', function (r) { return !cung(r); }, [hang_('DanhGiaKy', {
      NamHoc: namHoc, Ky: ky, MaDinhDanh: x.ma, Phan: x.phan, Muc: x.muc || '', Diem: x.diem || '', NhanXet: x.nx || '',
      MucNLPC: x.mucNLPC ? JSON.stringify(x.mucNLPC) : '', NXNLChung: x.nlc || '', NXNLDacThu: x.nld || '', NXPhamChat: x.pc || '', CapNhat: luc })]);
    return luc;
  });
}
function boSuaKy(ky, ma, phan) {
  return voiKhoa_(function () {
    var namHoc = caiDat_().NamHoc || '';
    thayTheTheo_('DanhGiaKy', function (r) { return !(r.NamHoc === namHoc && r.Ky === ky && r.MaDinhDanh === ma && r.Phan === phan); }, []);
    return 'Đã lấy lại bản app gợi ý.';
  });
}

// ============================================================================ TRỢ LÝ AI (GEMINI) – TUỲ CHỌN
// Khoá API lưu trong Script Properties (không nằm trong Trang tính, app không bao giờ gửi khoá ra giao diện).
// Chỉ dùng khoá của dự án Google Cloud ĐÃ BẬT THANH TOÁN (gói trả phí): gói miễn phí có thể dùng dữ liệu để cải thiện sản phẩm.
// Không gửi tên, ngày sinh, mã định danh học sinh cho AI.
var MODEL_GEMINI_MAC_DINH = 'gemini-2.5-flash';
function thuocTinh_() { return PropertiesService.getScriptProperties(); }
function layCaiDatAI() {
  var k = thuocTinh_().getProperty('GEMINI_API_KEY') || '';
  return { coKhoa: !!k, duoi: k ? k.slice(-4) : '', model: thuocTinh_().getProperty('GEMINI_MODEL') || MODEL_GEMINI_MAC_DINH };
}
function luuCaiDatAI(o) {
  var p = thuocTinh_();
  if (o.xoaKhoa) p.deleteProperty('GEMINI_API_KEY');
  else if (o.khoa) {
    var k = String(o.khoa).trim();
    if (!/^[\w-]{20,}$/.test(k)) throw new Error('Khoá API không đúng dạng (thường bắt đầu bằng "AIza…", không có dấu cách).');
    p.setProperty('GEMINI_API_KEY', k);
  }
  if (o.model != null) { var m = String(o.model).trim(); if (m) p.setProperty('GEMINI_MODEL', m); else p.deleteProperty('GEMINI_MODEL'); }
  return layCaiDatAI();
}
// Gọi Gemini; parts theo đúng định dạng API. Trả về chữ của câu trả lời.
function goiGemini_(parts, json) {
  var cd = layCaiDatAI(), khoa = thuocTinh_().getProperty('GEMINI_API_KEY');
  if (!khoa) throw new Error('Chưa cài khoá AI Gemini (Cài đặt → Trợ lý AI).');
  var r = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(cd.model) + ':generateContent', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true, headers: { 'x-goog-api-key': khoa },
    payload: JSON.stringify({ contents: [{ role: 'user', parts: parts }],
      generationConfig: json ? { temperature: 0, responseMimeType: 'application/json' } : { temperature: 0.4 } })
  });
  var ma = r.getResponseCode(), body = {};
  try { body = JSON.parse(r.getContentText()); } catch (e) {}
  if (ma !== 200) {
    var tb = (body.error && body.error.message) || ('mã lỗi ' + ma);
    if (ma === 400 && /API key/i.test(tb)) throw new Error('Khoá AI không hợp lệ – kiểm tra lại ở Cài đặt → Trợ lý AI.');
    if (ma === 403) throw new Error('Khoá AI không có quyền dùng Gemini (403). Kiểm tra dự án Google Cloud của khoá.');
    if (ma === 404) throw new Error('Không có mẫu AI "' + cd.model + '". Sửa tên mẫu ở Cài đặt → Trợ lý AI.');
    if (ma === 429) throw new Error('AI đang quá tải hoặc hết hạn mức – thử lại sau ít phút.');
    throw new Error('AI báo lỗi: ' + tb);
  }
  var c = body.candidates && body.candidates[0], ps = (c && c.content && c.content.parts) || [];
  var text = ps.filter(function (x) { return x.text && !x.thought; }).map(function (x) { return x.text; }).join('');
  if (!text) throw new Error('AI không trả lời (' + ((c && c.finishReason) || 'không rõ lí do') + ').');
  return text;
}
function thuKetNoiAI() {
  var t = goiGemini_([{ text: 'Trả lời đúng một từ: OK' }], false);
  return 'Kết nối được với AI (' + layCaiDatAI().model + '): ' + t.trim().slice(0, 30);
}
// Đọc điểm từ ảnh ô "Điểm" (chỉ phần ảnh đã cắt, không có tên học sinh). Trả {diem: 1..10 | null, chacChan: 0..1}
function docDiemAnh(b64) {
  var text = goiGemini_([
    { inline_data: { mime_type: 'image/jpeg', data: b64 } },
    { text: 'Ảnh là ô "Điểm" trên bài kiểm tra của học sinh tiểu học Việt Nam. Giáo viên viết tay điểm bằng số (thường mực đỏ), ' +
            'có thể kèm điểm viết bằng chữ (VD "Chín", "Mười"). Hãy đọc điểm. Điểm hợp lệ là số nguyên từ 1 đến 10. ' +
            'Nếu ảnh không có điểm, mờ, hoặc có nhiều số mâu thuẫn thì diem = null. ' +
            'Chỉ trả về JSON: {"diem": <số nguyên 1-10 hoặc null>, "chacChan": <số 0-1>, "ghiChu": "<rất ngắn, nếu có>"}' }
  ], true);
  var o = {};
  try { o = JSON.parse(text); } catch (e) { var m = /\{[\s\S]*\}/.exec(text); try { o = m ? JSON.parse(m[0]) : {}; } catch (e2) {} }
  var d = Number(o.diem);
  return { diem: (d >= 1 && d <= 10 && Math.floor(d) === d) ? d : null, chacChan: Math.max(0, Math.min(1, Number(o.chacChan) || 0)), ghiChu: String(o.ghiChu || '').slice(0, 100) };
}
// Lưu ảnh bài kiểm tra vào Drive: App Sổ theo dõi / Bài kiểm tra / <năm học> / <Kì> - <Môn>. Chụp lại thì ảnh cũ vào Thùng rác.
function luuAnhBaiKT(ky, ma, mon, b64, diemAI) {
  return voiKhoa_(function () {
    var namHoc = caiDat_().NamHoc || 'Chưa rõ năm học';
    var hs = doc_('HocSinh').filter(function (h) { return h.MaDinhDanh === ma; })[0];
    if (!hs) throw new Error('Không tìm thấy học sinh.');
    var thuMuc = function (cha, ten) {
      var it = cha ? cha.getFoldersByName(ten) : DriveApp.getFoldersByName(ten);
      return it.hasNext() ? it.next() : (cha ? cha.createFolder(ten) : DriveApp.createFolder(ten));
    };
    var dich = thuMuc(thuMuc(thuMuc(thuMuc(null, 'App Sổ theo dõi'), 'Bài kiểm tra'), namHoc), ky + ' - ' + mon);
    var cung = function (r) { return r.NamHoc === namHoc && r.Ky === ky && r.MaDinhDanh === ma && r.Mon === mon; };
    doc_('BaiKiemTra').filter(cung).forEach(function (r) { try { DriveApp.getFileById(r.FileId).setTrashed(true); } catch (e) {} });
    var ten = ('0' + (hs.ThuTu || '')).slice(-2) + '. ' + hs.HoTen + ' - ' + mon + ' ' + ky + '.jpg';
    var f = dich.createFile(Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', ten));
    var luc = Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm');
    thayTheTheo_('BaiKiemTra', function (r) { return !cung(r); }, [hang_('BaiKiemTra', {
      NamHoc: namHoc, Ky: ky, MaDinhDanh: ma, Mon: mon, FileId: f.getId(), Url: f.getUrl(), DiemAI: diemAI == null ? '' : diemAI, CapNhat: luc })]);
    return { url: f.getUrl() };
  });
}
function goiGeminiJson_(parts, nhietDo) {   // như goiGemini_ nhưng trả JSON và cho chọn độ sáng tạo
  var cd = layCaiDatAI(), khoa = thuocTinh_().getProperty('GEMINI_API_KEY');
  if (!khoa) throw new Error('Chưa cài khoá AI Gemini (Cài đặt → Trợ lý AI).');
  var r = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(cd.model) + ':generateContent', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true, headers: { 'x-goog-api-key': khoa },
    payload: JSON.stringify({ contents: [{ role: 'user', parts: parts }], generationConfig: { temperature: nhietDo, responseMimeType: 'application/json' } })
  });
  var ma = r.getResponseCode(), body = {};
  try { body = JSON.parse(r.getContentText()); } catch (e) {}
  if (ma !== 200) {
    var tb = (body.error && body.error.message) || ('mã lỗi ' + ma);
    if (ma === 400 && /API key/i.test(tb)) throw new Error('Khoá AI không hợp lệ – kiểm tra lại ở Cài đặt → Trợ lý AI.');
    if (ma === 429) throw new Error('AI đang quá tải hoặc hết hạn mức – thử lại sau ít phút.');
    throw new Error('AI báo lỗi: ' + tb);
  }
  var c = body.candidates && body.candidates[0], ps = (c && c.content && c.content.parts) || [];
  var text = ps.filter(function (x) { return x.text && !x.thought; }).map(function (x) { return x.text; }).join('');
  try { return JSON.parse(text); } catch (e) { var m = /\[[\s\S]*\]/.exec(text); if (m) return JSON.parse(m[0]); }
  throw new Error('AI trả lời không đúng định dạng – thử lại.');
}

// ============================================================================ VIẾT LẠI NHẬN XÉT CUỐI KÌ BẰNG AI (tuỳ chọn)
// Chỉ cho CK1, CK2. Gửi cho AI: mã tạm (1, 2, 3…), tên môn/nhóm năng lực, mức, điểm, nội dung nhận xét đã bỏ tên học sinh.
// Không gửi họ tên, ngày sinh, mã định danh. Cô xem trước rồi mới dùng; mọi thay đổi ghi vào PhienBanNX để hoàn tác (xoá khi xuất file).
var COT_TRUONG = { nx: 'NhanXet', nlc: 'NXNLChung', nld: 'NXNLDacThu', pc: 'NXPhamChat' };
function catGioiHan_(t, n) {
  t = String(t || '').replace(/\s*\n+\s*/g, ' ').replace(/\s{2,}/g, ' ').trim();
  if (t.length <= n) return t;
  var cat = t.slice(0, n), i = Math.max(cat.lastIndexOf('. '), cat.lastIndexOf('! '));
  return i > n * 0.5 ? cat.slice(0, i + 1) : cat.replace(/\s+\S*$/, '').replace(/[,;:\s]+$/, '') + '.';
}
// ds: [{k, phan, muc, diem, goc}] (tối đa 10 mục); lan: số lần viết lại (lần sau sáng tạo hơn) → [{k, ban}]
function vietLaiAI(ky, ds, lan) {
  if (!/^CK/.test(ky) && ky !== 'GVCN') throw new Error('Viết lại bằng AI chỉ dùng cho nhận xét cuối kì (CK1, CK2) và nhận xét GVCN cuối năm.');
  var gvcn = ky === 'GVCN', gioiHan = gvcn ? GIOI_HAN_GVCN : GIOI_HAN_KY;
  if (!ds || !ds.length) return [];
  if (ds.length > 10) throw new Error('Mỗi lần gửi tối đa 10 nhận xét.');
  var gui = ds.map(function (x) {
    return { k: String(x.k), phan: String(x.phan || ''), muc: String(x.muc || ''), diem: x.diem ? String(x.diem) : undefined,
             goc: String(x.goc || '').replace(/\d{6,}/g, '').slice(0, 1000) };   // bỏ mọi dãy số dài (phòng lẫn mã định danh)
  });
  var nhac = gvcn ? nhacGVCN_(gui) : 'Bạn là giáo viên chủ nhiệm lớp 2 ở Việt Nam, viết nhận xét đánh giá định kỳ cuối học kì theo Thông tư 27/2020/TT-BGDĐT. ' +
    'Viết lại từng nhận xét dưới đây cho mạch lạc, tự nhiên, gần gũi, mang tính động viên.\n' +
    'Quy tắc:\n' +
    '- Giữ đúng các ý trong bản gốc (điều làm tốt, điều cần cố gắng); KHÔNG bịa thêm thành tích, lỗi hay chi tiết mới. Được gộp ý trùng, lược tên bài học quá chi tiết.\n' +
    '- Phù hợp mức đánh giá. Môn học: T = Hoàn thành tốt, H = Hoàn thành, C = Chưa hoàn thành. Năng lực, phẩm chất: T = Tốt, Đ = Đạt, C = Cần cố gắng.\n' +
    '- Nêu điều làm tốt trước; nếu bản gốc có điều cần cố gắng thì nêu sau, diễn đạt nhẹ nhàng, cụ thể.\n' +
    '- Không nêu tên học sinh, không xưng "con". Có thể bắt đầu bằng "Em".\n' +
    '- Mỗi nhận xét tối đa ' + GIOI_HAN_KY + ' ký tự (nên 200–400), một đoạn văn, không gạch đầu dòng, không xuống dòng. Tiếng Việt chuẩn, đúng chính tả.\n' +
    'Chỉ trả về JSON: [{"k": "<mã của mục>", "ban": "<nhận xét mới>"}], đủ tất cả các mục, đúng mã.\n\nCác mục:\n' + JSON.stringify(gui);
  var kq = goiGeminiJson_([{ text: nhac }], Math.min(1, 0.5 + 0.2 * (+lan || 0)));
  if (!Array.isArray(kq)) kq = kq && (kq.items || kq.ketQua) || [];
  var co = {}; gui.forEach(function (x) { co[x.k] = 1; });
  return kq.filter(function (x) { return x && co[String(x.k)] && x.ban; }).map(function (x) {
    var b = catGioiHan_(x.ban, gioiHan); return { k: String(x.k), ban: gvcn ? catTheoChu_(b, GIOI_HAN_CHU_GVCN) : b };
  });
}
// Dùng bản AI cho nhiều mục cùng lúc. ds: [{x: (như luuDanhGiaKy), ls: [{truong, truoc, sau}]}]; lo: mã lần viết lại
function apDungAI(ky, ds, lo) {
  return voiKhoa_(function () {
    var namHoc = caiDat_().NamHoc || '', luc = Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm:ss'), khoa = {}, moi = [], ls = [];
    ds.forEach(function (d) {
      var x = d.x;
      khoa[x.ma + '|' + x.phan] = 1;
      moi.push(hang_('DanhGiaKy', { NamHoc: namHoc, Ky: ky, MaDinhDanh: x.ma, Phan: x.phan, Muc: x.muc || '', Diem: x.diem || '', NhanXet: x.nx || '',
        MucNLPC: x.mucNLPC ? JSON.stringify(x.mucNLPC) : '', NXNLChung: x.nlc || '', NXNLDacThu: x.nld || '', NXPhamChat: x.pc || '', CapNhat: luc }));
      (d.ls || []).forEach(function (l) {
        ls.push(hang_('PhienBanNX', { NamHoc: namHoc, Ky: ky, MaDinhDanh: x.ma, Phan: x.phan, Truong: l.truong, Truoc: l.truoc, Sau: l.sau, Lo: lo, Luc: luc }));
      });
    });
    thayTheTheo_('DanhGiaKy', function (r) { return !(r.NamHoc === namHoc && r.Ky === ky && khoa[r.MaDinhDanh + '|' + r.Phan]); }, moi);
    themDong_('PhienBanNX', ls);
    return { so: ls.length, lo: lo };
  });
}
// Hoàn tác: yc = {ma, phan, truong} (1 ô, lần gần nhất) hoặc {lo} (cả lần viết lại cả lớp). yc.ep: bỏ qua kiểm tra "cô đã sửa sau đó".
function hoanTacAI(ky, yc) {
  return voiKhoa_(function () {
    var namHoc = caiDat_().NamHoc || '', ls = doc_('PhienBanNX'), chon = [];
    var cungKy = function (r) { return r.NamHoc === namHoc && r.Ky === ky; };
    if (yc.lo) { for (var i = ls.length - 1; i >= 0; i--) if (cungKy(ls[i]) && ls[i].Lo === yc.lo) chon.push(i); }
    else for (var j = ls.length - 1; j >= 0; j--) { var r = ls[j]; if (cungKy(r) && r.MaDinhDanh === yc.ma && r.Phan === yc.phan && r.Truong === yc.truong) { chon.push(j); break; } }
    if (!chon.length) throw new Error('Không còn gì để hoàn tác.');
    var dg = doc_('DanhGiaKy'), viTri = {};
    dg.forEach(function (r, i) { if (cungKy(r)) viTri[r.MaDinhDanh + '|' + r.Phan] = i; });
    var doi = [], boQua = [], xoa = {};
    chon.forEach(function (i) {     // từ mới nhất về cũ nhất
      var l = ls[i], k = viTri[l.MaDinhDanh + '|' + l.Phan], cot = COT_TRUONG[l.Truong];
      if (k == null || !cot) { xoa[i] = 1; return; }
      if (dg[k][cot] !== l.Sau && !yc.ep) { boQua.push({ ma: l.MaDinhDanh, phan: l.Phan, truong: l.Truong }); return; }
      dg[k][cot] = l.Truoc; xoa[i] = 1;
      doi.push({ ma: l.MaDinhDanh, phan: l.Phan, truong: l.Truong, nd: l.Truoc });
    });
    if (doi.length) ghiTatCa_('DanhGiaKy', dg.map(function (r) { return hang_('DanhGiaKy', r); }));
    if (Object.keys(xoa).length) ghiTatCa_('PhienBanNX', ls.filter(function (r, i) { return !xoa[i]; }).map(function (r) { return hang_('PhienBanNX', r); }));
    return { doi: doi, boQua: boQua };
  });
}
// Sau khi xuất file: xoá lịch sử phiên bản của phần đã xuất (loai 'mon' hoặc 'nlpc')
function xoaLichSuAI(ky, loai) {
  return voiKhoa_(function () {
    var namHoc = caiDat_().NamHoc || '';
    thayTheTheo_('PhienBanNX', function (r) { return !(r.NamHoc === namHoc && r.Ky === ky && ((r.Phan === 'NLPC') === (loai === 'nlpc'))); }, []);
    return 'ok';
  });
}

// ============================================================================ NHẬN XÉT CỦA GVCN CUỐI NĂM (≤ 200 chữ, ≤ 900 ký tự – file "nhan_xet_GVCN_lop_")
// Bản nháp ghép theo quy tắc từ kết quả đánh giá cuối năm (CK2; chưa có thì tạm CK1). Lưu phần cô sửa vào DanhGiaKy (Ky = Phan = 'GVCN')
// để dùng chung cơ chế lưu, "Viết lại bằng AI", hoàn tác và lịch sử phiên bản với đánh giá định kỳ.
var GIOI_HAN_GVCN = 900, GIOI_HAN_CHU_GVCN = 200;
var CUM_PC = { 'Yêu nước': 'yêu quê hương, đất nước', 'Nhân ái': 'biết yêu thương, giúp đỡ bạn bè', 'Chăm chỉ': 'chăm chỉ',
               'Trung thực': 'trung thực', 'Trách nhiệm': 'có trách nhiệm với việc được giao' };
function demChu_(t) { t = String(t || '').trim(); return t ? t.split(/\s+/).length : 0; }
function noiDs_(ds) { return ds.length < 2 ? ds.join('') : ds.slice(0, -1).join(', ') + ' và ' + ds[ds.length - 1]; }
function catTheoChu_(t, soChu) {         // giữ trọn câu, không quá soChu chữ
  if (demChu_(t) <= soChu) return t;
  var out = '';
  tachCau_(t).forEach(function (c) { var thu = out ? out + ' ' + c : c; if (demChu_(thu) <= soChu) out = thu; });
  return out || t.split(/\s+/).slice(0, soChu).join(' ');
}
function tomTatNam_(d, nguon) {
  var mon = nguon.mon.map(function (m) { var x = d.mon[m.ten]; return m.ten + ' ' + (x.muc || '?') + (x.diem ? ' (' + x.diem + ' điểm)' : ''); }).join(', ');
  var dem = { T: 0, 'Đ': 0, C: 0 }; nguon.dsNLPC.forEach(function (t) { var m = d.nlpc.muc[t]; if (dem[m] != null) dem[m]++; });
  return 'Môn học: ' + mon + '. Năng lực, phẩm chất: Tốt ' + dem.T + ', Đạt ' + dem['Đ'] + ', Cần cố gắng ' + dem.C + '.';
}
function ghepNhanXetGVCN_(d, nguon, lopSau, thanhTich) {
  var theo = { T: [], H: [], C: [] }, cau = [];
  nguon.mon.forEach(function (m) { var x = d.mon[m.ten]; if (theo[x.muc]) theo[x.muc].push(m.ten); });
  var phan = [];
  if (theo.T.length) phan.push('hoàn thành tốt ' + (theo.T.length > 1 ? 'các môn ' : 'môn ') + noiDs_(theo.T));
  if (theo.H.length) phan.push('hoàn thành ' + (theo.H.length > 1 ? 'các môn ' : 'môn ') + noiDs_(theo.H));
  if (theo.C.length) phan.push('chưa hoàn thành ' + (theo.C.length > 1 ? 'các môn ' : 'môn ') + noiDs_(theo.C));
  if (phan.length) cau.push('Năm học qua, em ' + phan.join('; ') + '.');
  theo.T.slice(0, 2).forEach(function (m) {                // ý cụ thể từ nhận xét môn: "Kể chuyện…" → "Em kể chuyện…"
    var c = tachCau_(d.mon[m].nx || '')[0];
    if (c) cau.push(/^(Em|Con)\s/.test(c) || /^["“(]/.test(c) ? c : 'Em ' + chuThuong_(c));
  });
  var p = d.nlpc, tot = nguon.dsNLPC.filter(function (t) { return p.muc[t] === 'T'; }), can = nguon.dsNLPC.filter(function (t) { return p.muc[t] === 'C'; });
  var pcTot = tot.filter(function (t) { return CUM_PC[t]; }).map(function (t) { return CUM_PC[t]; });
  var nlTot = tot.filter(function (t) { return !CUM_PC[t]; }).map(chuThuong_);
  cau = cau.concat(cauThanhTich_(thanhTich));                // phong trào, cuộc thi, giải thưởng trong năm
  if (pcTot.length) cau.push('Em ' + noiDs_(pcTot) + '.');
  if (nlTot.length) cau.push('Em thể hiện tốt năng lực ' + noiDs_(nlTot) + '.');
  if (!pcTot.length && !nlTot.length) cau.push('Em ngoan, thực hiện tốt nội quy lớp học.');
  if (can.length) cau.push('Em cần cố gắng hơn về ' + noiDs_(can.map(chuThuong_)) + '.');
  else if (theo.C.length) cau.push('Em cần cố gắng hơn trong học tập.');
  cau.push('Cô mong em tiếp tục phát huy' + (can.length || theo.H.length || theo.C.length ? ' điểm mạnh và cố gắng hơn nữa' : '') + (lopSau ? ' ở lớp ' + lopSau : '') + '.');
  return catTheoChu_(noiCau_(cau, GIOI_HAN_GVCN), GIOI_HAN_CHU_GVCN);
}
function layNhanXetGVCN_() {
  var cd = caiDat_(), namHoc = cd.NamHoc || '', nguon = null, kyNguon = '';
  try { nguon = layDanhGiaKy('CK2'); kyNguon = 'CK2'; }
  catch (e) {
    try { nguon = layDanhGiaKy('CK1'); kyNguon = 'CK1'; }
    catch (e2) { throw new Error('Chưa có dữ liệu cuối kì để soạn nhận xét GVCN (cần Sổ báo giảng và đánh giá cuối kì).'); }
  }
  var lopSau = +cd.Khoi ? (+cd.Khoi + 1) : '';
  var thanhTich = phongTraoCuaHS_(namHoc);
  var hsTat = {}; doc_('HocSinh').forEach(function (h) { hsTat[h.MaDinhDanh] = h; });
  var daSua = {};
  doc_('DanhGiaKy').forEach(function (r) { if (r.NamHoc === namHoc && r.Ky === 'GVCN' && r.Phan === 'GVCN') daSua[r.MaDinhDanh] = r; });
  var du = {};
  nguon.hs.forEach(function (h) {
    var d = nguon.du[h.ma], goc = ghepNhanXetGVCN_(d, nguon, lopSau, thanhTich[h.ma]);
    var soGhi = 0; nguon.mon.forEach(function (m) { soGhi += d.mon[m.ten].soGhi || 0; });
    var tt = thanhTich[h.ma] || [];
    var x = { muc: '', diem: '', soGhi: soGhi, nx: goc, goc: { nx: goc }, tomTat: tomTatNam_(d, nguon) +
      (tt.length ? ' Phong trào, cuộc thi: ' + tt.map(function (p) { return p.ten + ' (' + p.kq + (p.cap && p.cap !== 'Lớp' ? ', cấp ' + p.cap.toLowerCase() : '') + ')'; }).join('; ') + '.' : '') };
    var s = daSua[h.ma]; if (s) { x.nx = s.NhanXet; x.sua = true; }
    du[h.ma] = { mon: { GVCN: x } };
  });
  var lichSu = doc_('PhienBanNX').filter(function (r) { return r.NamHoc === namHoc && r.Ky === 'GVCN'; })
    .map(function (r) { return { ma: r.MaDinhDanh, phan: r.Phan, truong: r.Truong, lo: r.Lo }; });
  return { ky: 'GVCN', gvcn: true, kyNguon: kyNguon, tu: nguon.tu, den: nguon.den, cuoiKy: true, namHoc: namHoc,
           gioiHan: GIOI_HAN_GVCN, gioiHanChu: GIOI_HAN_CHU_GVCN, dsNLPC: nguon.dsNLPC, nhomNLPC: nguon.nhomNLPC, lichSu: lichSu,
           mon: [{ ten: 'GVCN', sheet: 'Sheet1', coDiem: false }],
           hs: nguon.hs.map(function (h) { return { ma: h.ma, ten: h.ten, ns: h.ns, stt: h.stt, gt: (hsTat[h.ma] || {}).GioiTinh || '' }; }), du: du };
}
function nhacGVCN_(gui) {
  var khoi = caiDat_().Khoi || '2';
  return 'Bạn là giáo viên chủ nhiệm lớp ' + khoi + ' ở Việt Nam, viết "Nhận xét của giáo viên chủ nhiệm" cuối năm học cho từng học sinh (theo Thông tư 27/2020/TT-BGDĐT). ' +
    'Viết lại từng nhận xét dưới đây cho mạch lạc, tự nhiên, ấm áp, mang tính động viên.\n' +
    'Quy tắc:\n' +
    '- Giữ đúng các ý trong bản gốc và phần "muc" (kết quả học tập các môn, năng lực, phẩm chất); KHÔNG bịa thêm thành tích, lỗi hay chi tiết mới.\n' +
    '- Bố cục: kết quả học tập nổi bật; năng lực, phẩm chất nổi bật; điều cần cố gắng (nếu bản gốc có); một câu động viên ngắn.\n' +
    '- Không nêu tên học sinh, không xưng "con". Có thể bắt đầu bằng "Em".\n' +
    '- Tối đa ' + GIOI_HAN_CHU_GVCN + ' chữ và ' + GIOI_HAN_GVCN + ' ký tự (nên 60–120 chữ), một đoạn văn, không gạch đầu dòng, không xuống dòng. Tiếng Việt chuẩn, đúng chính tả.\n' +
    'Chỉ trả về JSON: [{"k": "<mã của mục>", "ban": "<nhận xét mới>"}], đủ tất cả các mục, đúng mã.\n\nCác mục:\n' + JSON.stringify(gui);
}

// ============================================================================ MINH CHỨNG (ẢNH TÀI LIỆU CỦA HỌC SINH)
// 1 tài liệu = nhiều trang ảnh. Ảnh lưu ở Drive: App Sổ theo dõi / Minh chứng / <năm học> / <STT. Họ tên>.
// Bảng MinhChung (1 dòng/tài liệu), MinhChung_Trang (1 dòng/trang, kèm ảnh thu nhỏ để hiện nhanh trong app).
var LOAI_MC = ['Bài kiểm tra', 'Bài làm, vở', 'Sản phẩm (vẽ, thủ công…)', 'Giấy khen, thành tích', 'Khác'];
function thuMucCon_(cha, ten) {
  var it = cha ? cha.getFoldersByName(ten) : DriveApp.getFoldersByName(ten);
  return it.hasNext() ? it.next() : (cha ? cha.createFolder(ten) : DriveApp.createFolder(ten));
}
function timThuMuc_(duongDan) {          // tìm (không tạo) thư mục theo đường dẫn từ gốc Drive; không có → null
  var cha = null;
  for (var i = 0; i < duongDan.length; i++) {
    var it = cha ? cha.getFoldersByName(duongDan[i]) : DriveApp.getFoldersByName(duongDan[i]);
    if (!it.hasNext()) return null;
    cha = it.next();
  }
  return cha;
}
function layMinhChung(ma) {
  var trang = {};
  doc_('MinhChung_Trang').forEach(function (t) { (trang[t.MaMC] = trang[t.MaMC] || []).push({ so: +t.TrangSo, fileId: t.FileId, url: t.Url, nho: t.AnhNho, kb: Math.round((+t.KichThuoc || 0) / 1024) }); });
  var namHoc = caiDat_().NamHoc || '';
  var hs = doc_('HocSinh').filter(function (h) { return h.MaDinhDanh === ma; })[0] || {};
  return { loai: LOAI_MC, namHoc: namHoc, tenHS: hs.HoTen || '', ds: doc_('MinhChung').filter(function (m) { return m.MaDinhDanh === ma; }).map(function (m) {
    return { maMC: m.MaMC, namHoc: m.NamHoc, ten: m.Ten, loai: m.Loai, ngay: m.NgayTai, ghiChu: m.GhiChu,
             trang: (trang[m.MaMC] || []).sort(function (a, b) { return a.so - b.so; }) };
  }).reverse() };
}
function taoMinhChung(ma, info) {
  return voiKhoa_(function () {
    var hs = doc_('HocSinh').filter(function (h) { return h.MaDinhDanh === ma; })[0];
    if (!hs) throw new Error('Không tìm thấy học sinh.');
    var namHoc = caiDat_().NamHoc || 'Chưa rõ năm học';
    var thu = thuMucCon_(thuMucCon_(thuMucCon_(thuMucCon_(null, 'App Sổ theo dõi'), 'Minh chứng'), namHoc), ('0' + (hs.ThuTu || '')).slice(-2) + '. ' + hs.HoTen);
    var maMC = 'MC' + Utilities.getUuid().replace(/-/g, '').slice(0, 10);
    themDong_('MinhChung', [hang_('MinhChung', { MaMC: maMC, NamHoc: namHoc, MaDinhDanh: ma, Ten: String(info.ten || 'Tài liệu').slice(0, 100),
      Loai: info.loai || 'Khác', SoTrang: 0, NgayTai: Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy'), GhiChu: String(info.ghiChu || '').slice(0, 300), ThuMucId: thu.getId() })]);
    return maMC;
  });
}
// Thêm 1 trang (ảnh JPEG base64) vào cuối tài liệu; anhNho: ảnh thu nhỏ (dataURL) để hiện danh sách
function themTrangMC(maMC, b64, anhNho) {
  var mc = doc_('MinhChung').filter(function (m) { return m.MaMC === maMC; })[0];
  if (!mc) throw new Error('Không tìm thấy tài liệu.');
  var so = 1 + doc_('MinhChung_Trang').filter(function (t) { return t.MaMC === maMC; }).reduce(function (m, t) { return Math.max(m, +t.TrangSo); }, 0);
  var thu = mc.ThuMucId ? DriveApp.getFolderById(mc.ThuMucId) : thuMucCon_(null, 'App Sổ theo dõi');
  var bytes = Utilities.base64Decode(b64);
  var f = thu.createFile(Utilities.newBlob(bytes, 'image/jpeg', mc.Ten + ' (' + mc.NgayTai.replace(/\//g, '-') + ') - trang ' + so + '.jpg'));
  return voiKhoa_(function () {
    themDong_('MinhChung_Trang', [hang_('MinhChung_Trang', { MaMC: maMC, TrangSo: so, FileId: f.getId(), Url: f.getUrl(), KichThuoc: bytes.length,
      AnhNho: String(anhNho || '').length < 40000 ? anhNho : '' })]);
    capNhatSoTrang_(maMC);
    return { so: so };
  });
}
function capNhatSoTrang_(maMC) {
  var n = doc_('MinhChung_Trang').filter(function (t) { return t.MaMC === maMC; }).length;
  var ds = doc_('MinhChung'); ds.forEach(function (m) { if (m.MaMC === maMC) m.SoTrang = n; });
  ghiTatCa_('MinhChung', ds.map(function (m) { return hang_('MinhChung', m); }));
}
function layAnhTrang(fileId) { return Utilities.base64Encode(DriveApp.getFileById(fileId).getBlob().getBytes()); }
function layAnhMC(maMC) {          // tất cả trang (để gộp PDF trên máy)
  return doc_('MinhChung_Trang').filter(function (t) { return t.MaMC === maMC && t.FileId; })
    .sort(function (a, b) { return +a.TrangSo - +b.TrangSo; }).map(function (t) { return layAnhTrang(t.FileId); });
}
function luuPDFMinhChung(maMC, b64) {   // bản PDF gộp → lưu cạnh ảnh trong Drive (dự phòng khi không tải về được)
  var mc = doc_('MinhChung').filter(function (m) { return m.MaMC === maMC; })[0];
  if (!mc) throw new Error('Không tìm thấy tài liệu.');
  var thu = mc.ThuMucId ? DriveApp.getFolderById(mc.ThuMucId) : thuMucCon_(null, 'App Sổ theo dõi');
  var f = thu.createFile(Utilities.newBlob(Utilities.base64Decode(b64), 'application/pdf', mc.Ten + ' (' + mc.NgayTai.replace(/\//g, '-') + ').pdf'));
  return { url: f.getUrl() };
}
function xoaTrangMC(maMC, so) {
  return voiKhoa_(function () {
    var ds = doc_('MinhChung_Trang'), con = [], k = 0;
    ds.filter(function (t) { return t.MaMC === maMC; }).sort(function (a, b) { return +a.TrangSo - +b.TrangSo; }).forEach(function (t) {
      if (+t.TrangSo === +so) { try { DriveApp.getFileById(t.FileId).setTrashed(true); } catch (e) {} return; }
      t.TrangSo = ++k; con.push(t);
    });
    ghiTatCa_('MinhChung_Trang', ds.filter(function (t) { return t.MaMC !== maMC; }).concat(con).map(function (t) { return hang_('MinhChung_Trang', t); }));
    capNhatSoTrang_(maMC);
    return 'Đã xoá trang ' + so + ' (ảnh vào Thùng rác Drive).';
  });
}
function xoaMinhChung(maMC) {
  return voiKhoa_(function () {
    doc_('MinhChung_Trang').forEach(function (t) { if (t.MaMC === maMC && t.FileId) { try { DriveApp.getFileById(t.FileId).setTrashed(true); } catch (e) {} } });
    thayTheTheo_('MinhChung_Trang', function (t) { return t.MaMC !== maMC; }, []);
    thayTheTheo_('MinhChung', function (m) { return m.MaMC !== maMC; }, []);
    return 'Đã xoá tài liệu (ảnh vào Thùng rác Drive, khôi phục được trong 30 ngày).';
  });
}

// ============================================================================ DUNG LƯỢNG ẢNH + XOÁ ẢNH NĂM CŨ
function layDungLuong() {
  var namHoc = caiDat_().NamHoc || '', theo = {};
  var lay = function (n) { return theo[n] = theo[n] || { namHoc: n, taiLieu: 0, trang: 0, byte: 0, anhKT: 0, daXoa: 0 }; };
  var namCuaMC = {};
  doc_('MinhChung').forEach(function (m) { namCuaMC[m.MaMC] = m.NamHoc; lay(m.NamHoc).taiLieu++; });
  doc_('MinhChung_Trang').forEach(function (t) { var o = lay(namCuaMC[t.MaMC] || '?'); o.trang++; o.byte += +t.KichThuoc || 0; });
  doc_('BaiKiemTra').forEach(function (b) { if (b.FileId) { lay(b.NamHoc).anhKT++; lay(b.NamHoc).byte += 250 * 1024; } });   // ảnh bài KT ~250 KB
  var ds = Object.keys(theo).map(function (k) { var o = theo[k]; o.mb = Math.round(o.byte / 1048576 * 10) / 10; o.dangDung = k === namHoc; return o; })
    .sort(function (a, b) { return a.namHoc < b.namHoc ? 1 : -1; });
  return { namHoc: namHoc, ds: ds };
}
function xoaAnhNamHoc(namHoc, goLai) {
  if (goLai !== namHoc) throw new Error('Gõ lại đúng năm học để xác nhận.');
  return voiKhoa_(function () {
    if (namHoc === caiDat_().NamHoc) throw new Error('Không xoá ảnh của năm học đang dùng.');
    var luc = Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy');
    var boRac = function (id) { try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {} };
    // Cả thư mục năm học vào Thùng rác (nhanh); không thấy thư mục (cô đã chuyển chỗ) thì bỏ từng ảnh theo mã file
    var tmMC = timThuMuc_(['App Sổ theo dõi', 'Minh chứng', namHoc]), tmKT = timThuMuc_(['App Sổ theo dõi', 'Bài kiểm tra', namHoc]);
    var maMC = {}; doc_('MinhChung').forEach(function (m) { if (m.NamHoc === namHoc) maMC[m.MaMC] = 1; });
    var trangNam = doc_('MinhChung_Trang').filter(function (t) { return maMC[t.MaMC]; }), soTrang = trangNam.length;
    if (tmMC) tmMC.setTrashed(true); else trangNam.forEach(function (t) { if (t.FileId) boRac(t.FileId); });
    if (tmKT) tmKT.setTrashed(true); else doc_('BaiKiemTra').forEach(function (b) { if (b.NamHoc === namHoc && b.FileId) boRac(b.FileId); });
    thayTheTheo_('MinhChung_Trang', function (t) { return !maMC[t.MaMC]; }, []);
    ghiTatCa_('MinhChung', doc_('MinhChung').map(function (m) {
      if (maMC[m.MaMC]) { m.SoTrang = 0; m.GhiChu = (m.GhiChu ? m.GhiChu + ' ' : '') + '(đã xoá ảnh ' + luc + ')'; }
      return hang_('MinhChung', m);
    }));
    var soKT = 0;
    ghiTatCa_('BaiKiemTra', doc_('BaiKiemTra').map(function (b) { if (b.NamHoc === namHoc && b.FileId) { b.FileId = ''; b.Url = ''; soKT++; } return hang_('BaiKiemTra', b); }));
    return 'Đã xoá ảnh năm học ' + namHoc + ': ' + soTrang + ' trang minh chứng, ' + soKT + ' ảnh bài kiểm tra – ảnh đã vào Thùng rác Google Drive, khôi phục được trong 30 ngày. Nhận xét, kết quả vẫn giữ.';
  });
}

// ============================================================================ KẾT THÚC NĂM HỌC → NĂM HỌC MỚI
// Sao lưu nguyên file dữ liệu thành "Lưu trữ <năm học> - Lớp …" (thư mục App Sổ theo dõi / Lưu trữ), rồi làm trống
// các bảng theo năm. Giữ: cài đặt, môn học, kho thư viện, nhận xét chung; ảnh minh chứng (xoá riêng ở Dung lượng ảnh).
var BANG_THEO_NAM = ['HocSinh', 'DiemDanh', 'SoTheoDoi', 'NhanXetThang', 'DanhGiaKy', 'PhienBanNX', 'HocKy', 'TuanHoc', 'LichBaoGiang', 'PhongTrao', 'ThamGia', 'DanhGiaThang', 'NopVo', 'NhomHS'];
function namSau_(n) { var y = /(\d{4})-(\d{4})/.exec(n || ''); return y ? (+y[1] + 1) + '-' + (+y[2] + 1) : ''; }
function layNamHoc() {
  var cd = caiDat_(), namHoc = cd.NamHoc || '';
  var hs = doc_('HocSinh').filter(function (h) { return h.TrangThai !== 'Đã chuyển đi'; });
  var dem = function (ky, laNL) {
    var co = {}; doc_('DanhGiaKy').forEach(function (r) { if (r.NamHoc === namHoc && r.Ky === ky && (r.Phan === 'NLPC') === laNL) co[r.MaDinhDanh] = 1; });
    return hs.filter(function (h) { return co[h.MaDinhDanh]; }).length;
  };
  var hk2 = doc_('HocKy').filter(function (h) { return h.NamHoc === namHoc && h.HocKy === 'II'; })[0];
  var luuTru = []; try { luuTru = JSON.parse(cd.LuuTru || '[]'); } catch (e) {}
  return { namHoc: namHoc, lop: cd.Lop || '', siSo: hs.length, ck2: dem('CK2', false), nlpcCK2: dem('CK2', true), gvcn: dem('GVCN', false),
           hk2: hk2 ? { den: hk2.NgayDayCuoi, xong: soNgay_(homNay_(), hk2.NgayDayCuoi) < 0 } : null, goiY: namSau_(namHoc), luuTru: luuTru };
}
function batDauNamHocMoi(namMoi, goLai) {
  namMoi = String(namMoi || '').trim();
  if (!/^\d{4}-\d{4}$/.test(namMoi) || +namMoi.slice(5) !== +namMoi.slice(0, 4) + 1) throw new Error('Năm học mới phải có dạng 2027-2028.');
  if (goLai !== namMoi) throw new Error('Gõ lại đúng năm học mới để xác nhận.');
  return voiKhoa_(function () {
    var cd = caiDat_(), namCu = cd.NamHoc || '';
    if (namCu && namMoi <= namCu) throw new Error('Năm học mới phải sau năm học đang dùng (' + namCu + ').');
    var ss = SpreadsheetApp.getActive();
    var ban = DriveApp.getFileById(ss.getId()).makeCopy('Lưu trữ ' + (namCu || 'năm cũ') + ' - Lớp ' + (cd.Lop || '') + ' - Sổ theo dõi',
      thuMucCon_(thuMucCon_(null, 'App Sổ theo dõi'), 'Lưu trữ'));
    var luuTru = []; try { luuTru = JSON.parse(cd.LuuTru || '[]'); } catch (e) {}
    luuTru.push({ namHoc: namCu, lop: cd.Lop || '', url: ban.getUrl(), ngay: Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy') });
    BANG_THEO_NAM.forEach(function (t) { ghiTatCa_(t, []); });
    luuCaiDat_('LuuTru', JSON.stringify(luuTru));
    luuCaiDat_('NamHoc', namMoi);
    luuCaiDat_('MaHKHienTai', '');
    return { url: ban.getUrl(), namCu: namCu, namMoi: namMoi };
  });
}

// ============================================================================ TỔNG HỢP ĐIỂM DANH (tuần / tháng / học kì / cả năm / tuỳ chọn)
// Vắng cả ngày = 1 ngày, vắng buổi sáng hoặc chiều = 0,5 ngày. Chỉ tính các ngày cô đã lưu điểm danh.
function phamViDiemDanh_(yc) {
  var cd = caiDat_(), namHoc = cd.NamHoc || '', hk = doc_('HocKy').filter(function (h) { return h.NamHoc === namHoc; });
  var hkI = hk.filter(function (h) { return h.HocKy === 'I'; })[0], hkII = hk.filter(function (h) { return h.HocKy === 'II'; })[0];
  var f = function (d) { return Utilities.formatDate(d, TZ, 'dd/MM/yyyy'); };
  if (yc.kieu === 'tuan') {
    var d = toDate_(yc.ngay || homNay_()), thu = (d.getDay() + 6) % 7;      // thứ Hai = 0
    var dau = new Date(d.getFullYear(), d.getMonth(), d.getDate() - thu);
    return { tu: f(dau), den: f(new Date(dau.getFullYear(), dau.getMonth(), dau.getDate() + 6)), ten: 'Tuần từ ' + f(dau) };
  }
  if (yc.kieu === 'thang') {
    var t = +yc.thang, nam = t >= 8 ? +namHoc.slice(0, 4) : +namHoc.slice(5);
    if (!nam) throw new Error('Chưa đặt năm học ở Cài đặt.');
    return { tu: f(new Date(nam, t - 1, 1)), den: f(new Date(nam, t, 0)), ten: 'Tháng ' + t + '/' + nam };
  }
  if (yc.kieu === 'hk1' || yc.kieu === 'hk2') {
    var h = yc.kieu === 'hk1' ? hkI : hkII;
    if (!h) throw new Error('Chưa có Sổ báo giảng học kì ' + (yc.kieu === 'hk1' ? 'I' : 'II') + ' nên chưa biết ngày bắt đầu, kết thúc học kì.');
    return { tu: h.NgayBatDau, den: h.NgayDayCuoi, ten: 'Học kì ' + h.HocKy + ' năm học ' + namHoc };
  }
  if (yc.kieu === 'nam') {
    if (!hkI) throw new Error('Chưa có Sổ báo giảng học kì I nên chưa biết ngày bắt đầu năm học.');
    return { tu: hkI.NgayBatDau, den: hkII ? hkII.NgayDayCuoi : homNay_(), ten: 'Năm học ' + namHoc };
  }
  if (!yc.tu || !yc.den || key_(yc.tu) > key_(yc.den)) throw new Error('Chọn "từ ngày" trước "đến ngày".');
  return { tu: yc.tu, den: yc.den, ten: 'Từ ' + yc.tu + ' đến ' + yc.den };
}
function layTongHopDiemDanh(yc) {
  var pv = phamViDiemDanh_(yc), a = key_(pv.tu), b = key_(pv.den), cd = caiDat_();
  var trong = function (n) { var k = key_(n); return k >= a && k <= b; };
  var dd = doc_('DiemDanh').filter(function (r) { return trong(r.Ngay); });
  var ngayDD = {}; dd.forEach(function (r) { ngayDD[r.Ngay] = 1; });
  var ngayDay = {}; doc_('LichBaoGiang').forEach(function (r) { if (trong(r.Ngay) && key_(r.Ngay) <= key_(homNay_())) ngayDay[r.Ngay] = 1; });
  var chuaDD = Object.keys(ngayDay).filter(function (n) { return !ngayDD[n]; }).sort(function (x, y) { return key_(x) - key_(y); });
  // Học sinh: đang học + em chuyển đi sau ngày đầu của khoảng (vẫn có điểm danh trong khoảng)
  var hs = doc_('HocSinh').filter(function (h) { return h.TrangThai !== 'Đã chuyển đi' || (h.NgayChuyenDi && key_(h.NgayChuyenDi) >= a); })
    .sort(function (x, y) { return +x.ThuTu - +y.ThuTu; });
  var theo = {};
  dd.forEach(function (r) {
    if (r.TrangThai === 'Có mặt') return;
    var o = theo[r.MaDinhDanh] = theo[r.MaDinhDanh] || { P: 0, KP: 0, ngay: [] };
    var so = r.Buoi === 'Sáng' || r.Buoi === 'Chiều' ? 0.5 : 1, p = r.TrangThai === 'Vắng có phép';
    o[p ? 'P' : 'KP'] += so;
    o.ngay.push({ ngay: r.Ngay, loai: p ? 'P' : 'KP', buoi: r.Buoi || 'Cả ngày', ghiChu: r.GhiChu || '' });
  });
  var soNgay = Object.keys(ngayDD).length;
  return { ten: pv.ten, tu: pv.tu, den: pv.den, soNgay: soNgay, chuaDD: chuaDD, lop: cd.Lop || '', truong: cd.Truong || '', namHoc: cd.NamHoc || '',
    hs: hs.map(function (h, i) {
      var o = theo[h.MaDinhDanh] || { P: 0, KP: 0, ngay: [] };
      o.ngay.sort(function (x, y) { return key_(x.ngay) - key_(y.ngay); });
      return { stt: i + 1, ma: h.MaDinhDanh, ten: h.HoTen, chuyenDi: h.TrangThai === 'Đã chuyển đi', P: o.P, KP: o.KP,
               chuyenCan: soNgay ? Math.round((1 - (o.P + o.KP) / soNgay) * 1000) / 10 : null, ngay: o.ngay };
    }) };
}

// ============================================================================ PHONG TRÀO – CUỘC THI
// Ghi theo từng phong trào / cuộc thi: thông tin chung (PhongTrao) + các em tham gia và kết quả (ThamGia).
// Dùng trong hồ sơ học sinh, bản nháp nhận xét GVCN cuối năm, và xuất Excel báo cáo nhà trường.
var LOAI_PT = ['Phong trào', 'Cuộc thi', 'Hoạt động ngoại khoá', 'Khác'];
var CAP_PT = ['Lớp', 'Trường', 'Phường/Xã', 'Quận/Huyện', 'Thành phố/Tỉnh', 'Toàn quốc'];
var KET_QUA_PT = ['Tham gia', 'Giải Nhất', 'Giải Nhì', 'Giải Ba', 'Giải Khuyến khích', 'Huy chương Vàng', 'Huy chương Bạc', 'Huy chương Đồng', 'Đạt', 'Được khen'];
function laGiai_(kq) { return !!kq && kq !== 'Tham gia'; }
function layPhongTrao() {
  var namHoc = caiDat_().NamHoc || '', tg = {};
  doc_('ThamGia').forEach(function (r) { var o = tg[r.MaPT] = tg[r.MaPT] || { so: 0, giai: 0 }; o.so++; if (laGiai_(r.KetQua)) o.giai++; });
  return { loai: LOAI_PT, cap: CAP_PT, ketQua: KET_QUA_PT, namHoc: namHoc,
    ds: doc_('PhongTrao').filter(function (p) { return p.NamHoc === namHoc; })
      .sort(function (a, b) { return key_(b.Ngay) - key_(a.Ngay); })
      .map(function (p) { var o = tg[p.MaPT] || { so: 0, giai: 0 }; return { maPT: p.MaPT, ten: p.Ten, loai: p.Loai, cap: p.Cap, ngay: p.Ngay, ghiChu: p.GhiChu, so: o.so, giai: o.giai }; }) };
}
function layMotPhongTrao(maPT) {
  var p = maPT === 'moi' ? { MaPT: '', Ten: '', Loai: 'Cuộc thi', Cap: 'Trường', Ngay: homNay_(), GhiChu: '' }      // mục mới
    : doc_('PhongTrao').filter(function (x) { return x.MaPT === maPT; })[0];
  if (!p) throw new Error('Không tìm thấy phong trào / cuộc thi này.');
  var tg = {}; doc_('ThamGia').forEach(function (r) { if (r.MaPT === maPT) tg[r.MaDinhDanh] = r; });
  var hs = doc_('HocSinh').filter(function (h) { return h.TrangThai !== 'Đã chuyển đi' || tg[h.MaDinhDanh]; })
    .sort(function (a, b) { return +a.ThuTu - +b.ThuTu; });
  return { pt: { maPT: p.MaPT, ten: p.Ten, loai: p.Loai, cap: p.Cap, ngay: p.Ngay, ghiChu: p.GhiChu },
    hs: hs.map(function (h) { var r = tg[h.MaDinhDanh]; return { ma: h.MaDinhDanh, ten: h.HoTen, chuyenDi: h.TrangThai === 'Đã chuyển đi', kq: r ? r.KetQua : '', gc: r ? r.GhiChu : '' }; }) };
}
// info: {maPT?, ten, loai, cap, ngay (dd/MM/yyyy), ghiChu} → maPT
function luuPhongTrao(info) {
  var ten = String(info.ten || '').trim();
  if (!ten) throw new Error('Cô ghi tên phong trào / cuộc thi.');
  return voiKhoa_(function () {
    var namHoc = caiDat_().NamHoc || '', maPT = info.maPT || ('PT' + Utilities.getUuid().replace(/-/g, '').slice(0, 10));
    var luc = Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm');
    thayTheTheo_('PhongTrao', function (r) { return r.MaPT !== maPT; }, [hang_('PhongTrao', {
      MaPT: maPT, NamHoc: namHoc, Ten: ten.slice(0, 150), Loai: info.loai || 'Phong trào', Cap: info.cap || 'Lớp',
      Ngay: info.ngay || homNay_(), GhiChu: String(info.ghiChu || '').slice(0, 300), CapNhat: luc })]);
    return maPT;
  });
}
// ds: [{ma, kq, gc}] – danh sách đầy đủ các em tham gia (em không có trong ds = không tham gia)
function luuThamGia(maPT, ds) {
  return voiKhoa_(function () {
    var luc = Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm');
    thayTheTheo_('ThamGia', function (r) { return r.MaPT !== maPT; }, ds.map(function (x) {
      return hang_('ThamGia', { MaPT: maPT, MaDinhDanh: x.ma, KetQua: x.kq || 'Tham gia', GhiChu: String(x.gc || '').slice(0, 200), CapNhat: luc });
    }));
    var giai = ds.filter(function (x) { return laGiai_(x.kq); }).length;
    return 'Đã lưu: ' + ds.length + ' em tham gia' + (giai ? ', ' + giai + ' em đạt giải / được khen' : '') + '.';
  });
}
function xoaPhongTrao(maPT) {
  return voiKhoa_(function () {
    thayTheTheo_('ThamGia', function (r) { return r.MaPT !== maPT; }, []);
    thayTheTheo_('PhongTrao', function (r) { return r.MaPT !== maPT; }, []);
    return 'Đã xoá phong trào / cuộc thi.';
  });
}
// Phong trào của 1 em trong năm học (hồ sơ, nhận xét GVCN)
function phongTraoCuaHS_(namHoc) {
  var pt = {}; doc_('PhongTrao').forEach(function (p) { if (p.NamHoc === namHoc) pt[p.MaPT] = p; });
  var theo = {};
  doc_('ThamGia').forEach(function (r) {
    var p = pt[r.MaPT]; if (!p) return;
    (theo[r.MaDinhDanh] = theo[r.MaDinhDanh] || []).push({ ten: p.Ten, loai: p.Loai, cap: p.Cap, ngay: p.Ngay, kq: r.KetQua, gc: r.GhiChu });
  });
  Object.keys(theo).forEach(function (m) { theo[m].sort(function (a, b) { return key_(a.ngay) - key_(b.ngay); }); });
  return theo;
}
// "cuộc thi Viết chữ đẹp cấp trường" (không lặp chữ "cuộc thi" nếu tên đã có)
function tenPT_(x) {
  var tien = x.loai === 'Cuộc thi' ? 'cuộc thi ' : x.loai === 'Phong trào' ? 'phong trào ' : '';
  var ten = x.ten.trim();
  if (tien && ten.toLowerCase().indexOf(tien.trim()) === 0) { tien = ''; ten = chuThuong_(ten); }
  return tien + ten + (x.cap && x.cap !== 'Lớp' ? ' cấp ' + x.cap.toLowerCase() : '');
}
// Câu thành tích cho nhận xét GVCN: giải/khen trước (tối đa 3), rồi các phong trào tham gia (tối đa 3)
function cauThanhTich_(ds) {
  if (!ds || !ds.length) return [];
  var giai = ds.filter(function (x) { return laGiai_(x.kq); }), thamGia = ds.filter(function (x) { return !laGiai_(x.kq); }), cau = [];
  if (giai.length) cau.push('Em ' + noiDs_(giai.slice(0, 3).map(function (x) {
    var kq = x.kq === 'Được khen' ? 'được khen trong' : x.kq === 'Đạt' ? 'đạt' : 'đạt ' + chuThuong_(x.kq);
    return kq + ' ' + tenPT_(x);
  })) + '.');
  if (thamGia.length) cau.push('Em tích cực tham gia ' + noiDs_(thamGia.slice(0, 3).map(tenPT_)) + '.');
  return cau;
}
// Báo cáo cả năm: mỗi dòng 1 em tham gia 1 phong trào (để xuất Excel)
function layBaoCaoPhongTrao() {
  var cd = caiDat_(), namHoc = cd.NamHoc || '', hs = {};
  doc_('HocSinh').forEach(function (h) { hs[h.MaDinhDanh] = h; });
  var pt = doc_('PhongTrao').filter(function (p) { return p.NamHoc === namHoc; }).sort(function (a, b) { return key_(a.Ngay) - key_(b.Ngay); });
  var tg = {}; doc_('ThamGia').forEach(function (r) { (tg[r.MaPT] = tg[r.MaPT] || []).push(r); });
  var dong = [];
  pt.forEach(function (p) {
    (tg[p.MaPT] || []).sort(function (a, b) { return (+(hs[a.MaDinhDanh] || {}).ThuTu || 0) - (+(hs[b.MaDinhDanh] || {}).ThuTu || 0); }).forEach(function (r) {
      var h = hs[r.MaDinhDanh] || {};
      dong.push({ ngay: p.Ngay, ten: p.Ten, loai: p.Loai, cap: p.Cap, stt: h.ThuTu || '', ma: r.MaDinhDanh, hoTen: h.HoTen || '', ns: h.NgaySinh || '', kq: r.KetQua, gc: r.GhiChu });
    });
  });
  return { namHoc: namHoc, lop: cd.Lop || '', truong: cd.Truong || '', soPT: pt.length, dong: dong };
}

// ============================================================================ GIAO DIỆN GỌN: CHỨC NĂNG BẬT/TẮT + HIỆN THEO THỜI ĐIỂM
// Cô bật dần các chức năng nâng cao ở Cài đặt → Chức năng. Đánh giá định kỳ / Nhận xét GVCN / Năm học mới chỉ hiện quanh
// thời điểm cần làm (trừ khi cô chọn "Luôn hiện"). Trả về cùng Trang chủ để không thêm lần gọi máy chủ.
var CHUC_NANG_MAC_DINH = { theoBai: false, tongHopDD: true, phongTrao: true, minhChung: false, chupBai: false, ai: false, thoiDiem: 'tuDong' };
var TEN_KY = { GK1: 'giữa học kì I', CK1: 'cuối học kì I', GK2: 'giữa học kì II', CK2: 'cuối năm' };
function chucNang_() {
  var o = {}; try { o = JSON.parse(caiDat_().ChucNang || '{}'); } catch (e) {}
  var cn = {}; Object.keys(CHUC_NANG_MAC_DINH).forEach(function (k) { cn[k] = o[k] == null ? CHUC_NANG_MAC_DINH[k] : o[k]; });
  if (o.ai == null && thuocTinh_().getProperty('GEMINI_API_KEY')) cn.ai = true;     // đã cài khoá AI từ trước thì bật sẵn
  return cn;
}
function luuChucNang(o) {
  var cn = {}; Object.keys(CHUC_NANG_MAC_DINH).forEach(function (k) { if (o[k] != null) cn[k] = k === 'thoiDiem' ? (o[k] === 'luon' ? 'luon' : 'tuDong') : !!o[k]; });
  voiKhoa_(function () { luuCaiDat_('ChucNang', JSON.stringify(cn)); });
  _NHO = {};
  return giaoDien_();
}
function giaoDien_() {
  var cn = chucNang_(), nay = homNay_(), namHoc = caiDat_().NamHoc || '';
  var lech = function (moc) { return soNgay_(moc, nay); };            // số ngày từ mốc đến hôm nay (âm: chưa tới)
  var ky = [];
  ['GK1', 'CK1', 'GK2', 'CK2'].forEach(function (k) { try { ky.push({ ky: k, den: phamViKy_(k).den }); } catch (e) {} });
  // Định kỳ hiện từ 14 ngày trước hết kì đánh giá đến 30 ngày sau (cuối năm: 60 ngày); GVCN quanh cuối năm
  var kyHien = ky.filter(function (x) { var d = lech(x.den); return d >= -14 && d <= (x.ky === 'CK2' ? 60 : 30); }).map(function (x) { return x.ky; });
  var kyNhac = ky.filter(function (x) { var d = lech(x.den); return d >= -14 && d <= 7; }).map(function (x) { return x.ky; });
  var ck2 = ky.filter(function (x) { return x.ky === 'CK2'; })[0];
  var hienGVCN = !!ck2 && lech(ck2.den) >= -21 && lech(ck2.den) <= 60;
  var hk2 = doc_('HocKy').filter(function (h) { return h.NamHoc === namHoc && h.HocKy === 'II'; })[0];
  var hienNamHoc = !namHoc || (!!hk2 && lech(hk2.NgayDayCuoi) >= -7);
  if (cn.thoiDiem === 'luon') { kyHien = ['GK1', 'CK1', 'GK2', 'CK2']; hienGVCN = true; hienNamHoc = true; }
  return { chucNang: cn, kyHien: kyHien, kyNhac: kyNhac, hienGVCN: hienGVCN, hienNamHoc: hienNamHoc };
}

// ============================================================================ TẢI TRƯỚC (cho việc hằng ngày nhanh)
// Mở 1 bài: bộ nhận xét + học sinh và các nhận xét đã ghi – 1 lần gọi thay cho 2
function layMoBai(ngay, mon, bai, baiHS) { return { bo: layBoNhanXet(mon, bai), hs: layHocSinhChoBai(ngay, mon, baiHS || bai) }; }
// Khi cô đang xem Trang chủ: tải sẵn các tiết hôm nay + điểm danh hôm nay, trong 1 lần gọi
function layTaiTruoc(ngay, ds) {
  var out = { ngay: ngay, bai: {}, dd: layDiemDanh(ngay) };
  (ds || []).slice(0, 8).forEach(function (x) { out.bai[x.mon + '|' + (x.bai || '') + '|' + (x.baiHS || '')] = layMoBai(ngay, x.mon, x.bai, x.baiHS); });
  return out;
}

// ============================================================================ ĐÁNH GIÁ THƯỜNG XUYÊN HẰNG THÁNG (cách mới – theo quyết định của cô, 03/10/2026)
// Toán, Tiếng Việt: mỗi tháng 1 lần, 2 tiêu chí "Hiểu kiến thức", "Kỹ năng môn học" × 4 mức HTXS/HTT/HT/CCG (thư viện tháng).
// HĐTN, Đạo đức: cô tự nhận xét (có câu gợi ý theo mức). NL đặc thù: app tự suy từ mức Toán, TV. NL chung, phẩm chất: cô chọn mức.
// Máy chủ chỉ lưu MỨC cô chọn và phần cô SỬA (DanhGiaThang); câu nhận xét được ghép trên máy theo thư viện.
var MUC_THANG = ['HTXS', 'HTT', 'HT', 'CCG'];
var MON_VO = ['Tiếng Việt', 'Toán'];
function hocKyCuaThang_(t) { t = +t; return (t >= 8 || t === 1) ? 'I' : 'II'; }
function nhapThuVienThang(tenFile, tv) {
  return voiKhoa_(function () {
    var tt = tv.thongTin, khoi = String(tt.Khoi), bs = tt.BoSach, hk = tt.HocKy, msg = [];
    if (tv.v2) {
      tv.mon.forEach(function (m) {
        var thang = {}; m.dong.forEach(function (d) { thang[d.thang] = 1; });
        thayTheTheo_('ThuVienYeuCau', function (r) { return !(r.Khoi === khoi && r.BoSach === bs && r.HocKy === hk && r.Mon === m.ten && thang[r.Thang]); }, m.dong.map(function (d) {
          return hang_('ThuVienYeuCau', { Khoi: khoi, BoSach: bs, HocKy: hk, Thang: d.thang, Mon: m.ten, NoiDung: d.noiDung, MaYC: d.ma, YeuCau: d.yeuCau,
            CanCu: d.canCu, MucDo: d.mucDo, Mau: d.mau, CauNX: d.cau });
        }));
        var yc = {}; m.dong.forEach(function (d) { yc[d.thang + d.ma] = 1; });
        msg.push(m.ten + ' tháng ' + Object.keys(thang).join(', ') + ' (' + Object.keys(yc).length + ' yêu cầu, ' + m.dong.length + ' câu)');
      });
      var cd2 = caiDat_(); if (!cd2.BoSach) luuCaiDat_('BoSach', bs);
      return 'Đã đưa vào Kho thư viện THÁNG (yêu cầu kiến thức, kỹ năng) – Lớp ' + khoi + ', HK ' + hk + ': ' + msg.join('; ') + '.';
    }
    tv.mon.forEach(function (m) {
      thayTheTheo_('ThuVienThang', function (r) { return !(r.Khoi === khoi && r.BoSach === bs && r.HocKy === hk && r.Mon === m.ten); }, m.dong.map(function (d) {
        return hang_('ThuVienThang', { Khoi: khoi, BoSach: bs, HocKy: hk, Thang: d.thang, Mon: m.ten, TieuChi: d.tieuChi, MucDo: d.mucDo, Mau: d.mau, NoiDung: d.noiDung });
      }));
      msg.push(m.ten + ' (' + m.dong.length + ' câu)');
    });
    var cd = caiDat_(); if (!cd.BoSach) luuCaiDat_('BoSach', bs);
    return 'Đã đưa vào Kho thư viện THÁNG – Lớp ' + khoi + ', HK ' + hk + ': ' + msg.join('; ') + '.';
  });
}
function khoThang_() {          // tóm tắt kho thư viện tháng để hiện ở Cài đặt
  var o = {};
  doc_('ThuVienThang').forEach(function (r) {
    var k = r.Khoi + '|' + r.BoSach + '|' + r.HocKy + '|' + r.Mon;
    var x = o[k] = o[k] || { khoi: r.Khoi, boSach: r.BoSach, hocKy: r.HocKy, mon: r.Mon, soCau: 0, thang: {} };
    x.soCau++; x.thang[r.Thang] = 1;
  });
  doc_('ThuVienYeuCau').forEach(function (r) {
    var k = 'v2|' + r.Khoi + '|' + r.BoSach + '|' + r.HocKy + '|' + r.Mon;
    var x = o[k] = o[k] || { khoi: r.Khoi, boSach: r.BoSach, hocKy: r.HocKy, mon: r.Mon, soCau: 0, thang: {}, yc: {}, v2: true };
    x.soCau++; x.thang[r.Thang] = 1; x.yc[r.Thang + r.MaYC] = 1;
  });
  return Object.keys(o).map(function (k) { var x = o[k]; x.thang = Object.keys(x.thang); if (x.yc) x.soYC = Object.keys(x.yc).length; delete x.yc; return x; });
}

// ---------------------------------------------------------------------------- GHI ĐÁNH GIÁ THÁNG theo yêu cầu cần đạt (giống màn ghi theo bài)
// "Bài" ảo: 'Tháng 9 – Kiến thức' / 'Tháng 9 – Kỹ năng'; "Hoạt động" = yêu cầu cần đạt; ghi vào Sổ theo dõi như nhận xét theo bài.
var ND_THANG = ['Kiến thức', 'Kỹ năng'];
var MUC_DAY_DU = { HTXS: 'Hoàn thành xuất sắc', HTT: 'Hoàn thành tốt', HT: 'Hoàn thành', CHT: 'Chưa hoàn thành' };
function tenBaiThang_(thang, nd) { return 'Tháng ' + (+thang) + ' – ' + nd; }
function baiThang_(bai) {
  var m = /^Tháng (\d+) – (Kiến thức|Kỹ năng)$/.exec(String(bai || '')); return m ? { thang: +m[1], noiDung: m[2] } : null;
}
function thuVienYC_(mon, thang, nd) {          // các dòng thư viện của 1 (môn, tháng, nội dung) cho lớp đang dạy
  var cd = caiDat_(), hk = hocKyCuaThang_(thang);
  return doc_('ThuVienYeuCau').filter(function (r) {
    return String(r.Khoi) === String(cd.Khoi) && r.BoSach === cd.BoSach && r.HocKy === hk && +r.Thang === +thang && r.Mon === mon && (!nd || r.NoiDung === nd);
  });
}
function boThang_(mon, thang, nd) {
  var yc = {}, thuTu = [];
  thuVienYC_(mon, thang, nd).forEach(function (r) {
    var x = yc[r.YeuCau]; if (!x) { x = yc[r.YeuCau] = { ten: r.YeuCau, ma: r.MaYC, canCu: r.CanCu, muc: {}, mau: {} }; thuTu.push(r.YeuCau); }
    var m = MUC_DAY_DU[r.MucDo] || r.MucDo; (x.muc[m] = x.muc[m] || []).push(r.CauNX);
  });
  doc_('MauCuaCo').forEach(function (r) {
    if (r.Mon === mon && r.NoiDung === nd && yc[r.YeuCau]) (yc[r.YeuCau].mau[r.MucDo] = yc[r.YeuCau].mau[r.MucDo] || []).push(r.CauNX);
  });
  return { hoatDong: thuTu.map(function (t) { return yc[t]; }), tam: false, muc: MUC4, thang: { thang: +thang, noiDung: nd } };
}
// Các tháng đã có thư viện yêu cầu, theo môn – để hiện màn chọn
function layChonThang() {
  var cd = caiDat_(), o = {};
  doc_('ThuVienYeuCau').forEach(function (r) {
    if (String(r.Khoi) !== String(cd.Khoi) || r.BoSach !== cd.BoSach) return;
    var k = r.Mon + '|' + r.Thang + '|' + r.NoiDung; o[k] = o[k] || {}; o[k][r.MaYC] = 1;
  });
  var ds = Object.keys(o).map(function (k) { var p = k.split('|'); return { mon: p[0], thang: +p[1], noiDung: p[2], soYC: Object.keys(o[k]).length }; });
  var h = homNay_().split('/'), t = +h[1];
  if (+h[0] <= 7) t = t === 1 ? 12 : t - 1;       // tuần đầu tháng: thường cô đang chốt tháng trước
  return { ds: ds, thangNay: t };
}
// Ghi nhiều em, mỗi em 1 mức + 1 câu riêng (nút "Cả lớp: HTT"), trong 1 lần khoá
function luuNhanXetNhieu(x, ds) {
  return voiKhoa_(function () {
    if (!ds || !ds.length) throw new Error('Không có học sinh nào cần ghi.');
    var hk = hocKyHienTai_(), ngay = x.ngay || homNay_(), luc = Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm:ss');
    var tuan = hk ? (doc_('TuanHoc').filter(function (t) { return t.MaHK === hk.MaHK && key_(t.TuNgay) <= key_(ngay) && key_(ngay) <= key_(t.DenNgay) + 2; })[0] || {}).Tuan : '';
    var rows = ds.map(function (d) {
      return hang_('SoTheoDoi', { ID: Utilities.getUuid().slice(0, 8), ThoiGian: luc, Ngay: ngay, MaHK: hk ? hk.MaHK : '', Tuan: tuan || '',
        Thang: +ngay.split('/')[1], Tiet: '', MaDinhDanh: d.ma, Mon: x.mon, Bai: x.bai || '', HoatDong: x.hoatDong,
        MucDo: d.mucDo, NoiDung: d.noiDung || '', GhiChu: '' });
    });
    themDong_('SoTheoDoi', rows);
    return { msg: 'Đã ghi cho ' + rows.length + ' học sinh.', ids: rows.map(function (r) { return r[0]; }), luc: luc };
  });
}
// Mẫu câu riêng của cô cho 1 yêu cầu + mức (dùng lại các tháng sau nếu cùng yêu cầu)
function luuMauCuaCo(m) {
  return voiKhoa_(function () {
    var cau = String(m.cauNX || '').trim(); if (!cau) throw new Error('Chưa có câu để lưu thành mẫu.');
    if (!/[.!?]$/.test(cau)) cau += '.';
    var trung = doc_('MauCuaCo').some(function (r) { return r.Mon === m.mon && r.NoiDung === m.noiDung && r.YeuCau === m.yeuCau && r.MucDo === m.mucDo && r.CauNX === cau; });
    if (!trung) themDong_('MauCuaCo', [hang_('MauCuaCo', { Mon: m.mon, NoiDung: m.noiDung, YeuCau: m.yeuCau, MucDo: m.mucDo, CauNX: cau,
      CapNhat: Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm') })]);
    return cau;
  });
}
function xoaMauCuaCo(m) {
  return voiKhoa_(function () {
    thayTheTheo_('MauCuaCo', function (r) { return !(r.Mon === m.mon && r.NoiDung === m.noiDung && r.YeuCau === m.yeuCau && r.MucDo === m.mucDo && r.CauNX === m.cauNX); }, []);
    return 'Đã xoá mẫu.';
  });
}
// ---------------------------------------------------------------------------- TỔNG HỢP GHI THEO YÊU CẦU CẦN ĐẠT (bước 3)
var MA_MUC_YC = { 'Hoàn thành xuất sắc': 'HTXS', 'Hoàn thành tốt': 'HTT', 'Hoàn thành': 'HT', 'Chưa hoàn thành': 'CHT' };
var DIEM_YC = { HTXS: 4, HTT: 3, HT: 2, CHT: 1 };
// {mon: {Tiếng Việt: {kt: [{ma, ten}], kn: [...]}}, ghi: {maHS: {môn: {kt: {tên yêu cầu: {m: 'HTT', c: 'câu…'}}, kn: {…}}}}}
// Môn được tính theo yêu cầu khi tháng đó có thư viện yêu cầu VÀ (đã có lần ghi hoặc không có thư viện tháng bản cũ)
function ycThang_(thang, libCu) {
  var cd = caiDat_(), hk = hocKyCuaThang_(thang), ds = {}, ghi = {}, mon = {};
  doc_('ThuVienYeuCau').forEach(function (r) {
    if (String(r.Khoi) !== String(cd.Khoi) || r.BoSach !== cd.BoSach || r.HocKy !== hk || +r.Thang !== +thang) return;
    var a = ds[r.Mon] = ds[r.Mon] || { kt: [], kn: [], co: {} }, l = r.NoiDung === 'Kỹ năng' ? 'kn' : 'kt';
    if (!a.co[l + r.YeuCau]) { a.co[l + r.YeuCau] = 1; a[l].push({ ma: r.MaYC, ten: r.YeuCau }); }
  });
  Object.keys(ds).forEach(function (m) {
    var coGhi = false;
    ND_THANG.forEach(function (nd) {
      var l = nd === 'Kỹ năng' ? 'kn' : 'kt', bai = tenBaiThang_(thang, nd);
      (_NHO.SoTheoDoi ? doc_('SoTheoDoi') : docSTDTheoBai_(m, bai)).forEach(function (r) {
        if (r.Mon !== m || r.Bai !== bai) return;
        var g = (ghi[r.MaDinhDanh] = ghi[r.MaDinhDanh] || {})[m] = (ghi[r.MaDinhDanh] || {})[m] || { kt: {}, kn: {} };
        g[l][r.HoatDong] = { m: MA_MUC_YC[r.MucDo] || r.MucDo, c: [r.NoiDung, r.GhiChu].filter(function (x) { return x; }).join(' ') };   // lần ghi sau đè lần trước
        coGhi = true;
      });
    });
    if (coGhi || !libCu[m]) mon[m] = { kt: ds[m].kt, kn: ds[m].kn };
  });
  return { mon: mon, ghi: ghi };
}
function mucYC_(g) {          // trung bình mức các yêu cầu đã ghi → HTXS / HTT / HT / CHT
  var d = Object.keys(g || {}).map(function (k) { return DIEM_YC[g[k].m] || 0; }).filter(function (x) { return x; });
  if (!d.length) return '';
  var tb = d.reduce(function (a, b) { return a + b; }, 0) / d.length;
  return tb >= 3.5 ? 'HTXS' : tb >= 2.5 ? 'HTT' : tb >= 1.5 ? 'HT' : 'CHT';
}
// Ghép nhận xét môn từ các yêu cầu: điều làm tốt (HTXS, HTT) trước, điều cần cố gắng (HT, CHT) sau, câu vở cuối.
// Không vượt n ký tự; ưu tiên giữ ít nhất 1 điều làm tốt + 1 điều cần cố gắng + câu vở. (Giống hệt ghepYC trên máy.)
function ghepYC_(g, yc, vo, n) {
  g = g || { kt: {}, kn: {} };
  var lay = function (l, tot) {
    return yc[l].map(function (r) { return g[l][r.ten]; }).filter(function (x) { return x && x.c && ((x.m === 'HTXS' || x.m === 'HTT') === tot); })
      .map(function (x) { var c = String(x.c).trim(); return /[.!?]$/.test(c) ? c : c + '.'; });
  };
  var xen = function (a, b) { var o = []; for (var i = 0; i < Math.max(a.length, b.length); i++) { if (i < a.length) o.push(a[i]); if (i < b.length) o.push(b[i]); } return o; };
  var bo = function (ds) { var o = []; ds.forEach(function (c) { if (o.indexOf(c) < 0) o.push(c); }); return o; };
  var S = bo(xen(lay('kt', true), lay('kn', true))), L = bo(xen(lay('kt', false), lay('kn', false)));
  var uuTien = [], chon = {}, dai = 0;
  for (var i = 0; i < Math.max(S.length, L.length); i++) { if (i < S.length) uuTien.push('S' + i); if (i < L.length) uuTien.push('L' + i); if (i === 0 && vo) uuTien.push('V'); }
  if (!S.length && !L.length && vo) uuTien.push('V');
  var cau = function (k) { return k === 'V' ? vo : (k[0] === 'S' ? S : L)[+k.slice(1)]; };
  uuTien.forEach(function (k) { var c = cau(k), them = (dai ? 1 : 0) + c.length; if (dai + them <= n) { chon[k] = 1; dai += them; } });
  return S.filter(function (c, i) { return chon['S' + i]; }).concat(L.filter(function (c, i) { return chon['L' + i]; }), chon.V ? [vo] : []).join(' ');
}
// Các tháng đã có lần ghi theo yêu cầu (Sổ theo dõi chỉ giữ năm học hiện tại)
function thangCoGhiYC_() {
  var sh = sheetNho_('SoTheoDoi'), n = sh.getLastRow(), cMon = BANG.SoTheoDoi.indexOf('Mon'), o = {};
  if (n < 2) return o;
  var mb = _NHO._monBaiSTD || (_NHO._monBaiSTD = sh.getRange(2, cMon + 1, n - 1, 2).getDisplayValues());
  mb.forEach(function (r) { var b = baiThang_(r[1]); if (b) o[b.thang] = 1; });
  return o;
}
function duLieuThang_(thang) {
  thang = +thang;
  var cd = caiDat_(), namHoc = cd.NamHoc || '', hk = hocKyCuaThang_(thang), nam = namCuaThang_(thang);
  var hs = doc_('HocSinh').filter(function (h) { return h.TrangThai !== 'Đã chuyển đi'; }).sort(function (a, b) { return +a.ThuTu - +b.ThuTu; });
  var mon = monDangDay_(), lib = {};
  doc_('ThuVienThang').forEach(function (r) {
    if (String(r.Khoi) !== String(cd.Khoi) || r.BoSach !== cd.BoSach || r.HocKy !== hk || +r.Thang !== thang) return;
    var a = lib[r.Mon] = lib[r.Mon] || {}, b = a[r.TieuChi] = a[r.TieuChi] || {};
    (b[r.MucDo] = b[r.MucDo] || []).push(r.NoiDung);
  });
  var yc = ycThang_(thang, lib);                 // ghi theo yêu cầu cần đạt (Sổ theo dõi → Đánh giá tháng)
  var nxChung = { nlc: {}, pc: {} };            // câu năng lực chung, phẩm chất (bảng Nhận xét chung) theo mức Tốt/Đạt/Cần cố gắng
  doc_('NhanXetChung').forEach(function (r) {
    var pl = phanLoaiPCNL_(r.PCNL), dich = pl.pc.length ? nxChung.pc : pl.chung.length ? nxChung.nlc : null;
    if (dich) (dich[r.MucDo] = dich[r.MucDo] || []).push(r.NoiDung);
  });
  var du = {};
  doc_('DanhGiaThang').forEach(function (r) {
    if (r.NamHoc !== namHoc || +r.Thang !== thang) return;
    (du[r.MaDinhDanh] = du[r.MaDinhDanh] || {})[r.Phan] = { mucKT: r.MucKT, mucKN: r.MucKN, muc: r.Muc, nx: r.NhanXet };
  });
  var dv = demVo_(function (ngay) { var p = String(ngay).split('/'); return +p[1] === thang && +p[2] === nam; }), vo = dv.vo;
  // Môn ghi theo yêu cầu: mức Kiến thức / Kỹ năng luôn tính từ các yêu cầu (phần nhận xét cô sửa vẫn giữ)
  Object.keys(yc.mon).forEach(function (m) {
    hs.forEach(function (h) {
      var g = (yc.ghi[h.MaDinhDanh] || {})[m] || { kt: {}, kn: {} }, d = du[h.MaDinhDanh] = du[h.MaDinhDanh] || {};
      var o = d[m] = d[m] || { mucKT: '', mucKN: '', muc: '', nx: '' };
      o.mucKT = mucYC_(g.kt); o.mucKN = mucYC_(g.kn);
    });
  });
  return { thang: thang, nam: nam, namHoc: namHoc, hocKy: hk, gioiHan: GIOI_HAN_THANG, lib: lib, nxChung: nxChung, du: du, vo: vo, soNgayVo: dv.soNgay, soNgayVoCu: dv.soNgayCu, mucVo: MUC_VO,
           yc: yc.mon, ghi: yc.ghi,
           hs: hs.map(function (h, i) { return { ma: h.MaDinhDanh, ten: h.HoTen, stt: i + 1 }; }),
           mon: mon.map(function (m) { return { ten: m.TenMon, sheet: m.TenSheetBieuMau || m.TenMon, tieuChi: !!lib[m.TenMon] || !!yc.mon[m.TenMon], yc: !!yc.mon[m.TenMon] }; }),
           sheets: [{ ten: TEN_SHEET_TONG_HOP, mon: '' }].concat(mon.map(function (m) { return { ten: m.TenSheetBieuMau || m.TenMon, mon: m.TenMon }; })) };
}
// ds: [{ma, phan, mucKT, mucKN, muc, nx}] – lưu 1 em hoặc cả lớp (chọn nhanh) trong 1 lần
function luuDanhGiaThang(thang, ds) {
  return voiKhoa_(function () {
    var namHoc = caiDat_().NamHoc || '', luc = Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm'), khoa = {};
    ds.forEach(function (x) { khoa[x.ma + '|' + x.phan] = 1; });
    thayTheTheo_('DanhGiaThang', function (r) { return !(r.NamHoc === namHoc && +r.Thang === +thang && khoa[r.MaDinhDanh + '|' + r.Phan]); }, ds.map(function (x) {
      return hang_('DanhGiaThang', { NamHoc: namHoc, Thang: thang, MaDinhDanh: x.ma, Phan: x.phan, MucKT: x.mucKT || '', MucKN: x.mucKN || '', Muc: x.muc || '', NhanXet: x.nx || '', CapNhat: luc });
    }));
    return luc;
  });
}

// ============================================================================ NỘP VỞ HẰNG NGÀY (Tiếng Việt, Toán)
// Nộp vở (theo ý cô, 03/10/2026): 4 mức chất lượng (mức 3 có 2 câu) + "Chưa nộp". Mã: A–E, N.
// Lưu gọn: mỗi ngày mỗi môn 1 dòng (MaDinhDanh '*', GhiChu = "mã:A mã:B …"). Dữ liệu cách cũ (mỗi em 1 dòng
// Chưa nộp / Chưa làm xong / Nộp muộn + dòng '*' môn '*') vẫn đọc được.
var MUC_VO = [['A', 'Nhanh, đẹp, đúng'], ['B', 'Đẹp, đúng'], ['C', 'Đúng, biết trình bày'], ['D', 'Đúng, chưa biết trình bày'], ['E', 'Chưa đúng – chưa đẹp'], ['N', 'Chưa nộp']];
var CAU_VO = { A: 'Bài làm trong vở nhanh, đẹp, đúng.', B: 'Bài làm trong vở đẹp, đúng.', C: 'Bài làm trong vở đúng, biết trình bày.',
               D: 'Bài làm trong vở đúng nhưng chưa biết trình bày.', E: 'Bài làm trong vở chưa đúng, chưa đẹp, cần cố gắng hơn.' };
function giaiMaVo_(s) { var o = {}; String(s || '').split(/\s+/).forEach(function (x) { var p = x.split(':'); if (p.length === 2 && p[1]) o[p[0]] = p[1]; }); return o; }
// Đếm các lần trong những ngày thoả loc(ngay): {vo: {ma: {môn: {A.., N, cu: số lần chưa xong/muộn cách cũ}}}, soNgay, soNgayCu}
function demVo_(loc) {
  var vo = {}, ngay = {}, ngayCu = {};
  var o = function (ma, m) { var a = vo[ma] = vo[ma] || {}; return a[m] = a[m] || { A: 0, B: 0, C: 0, D: 0, E: 0, N: 0, cu: 0 }; };
  doc_('NopVo').forEach(function (r) {
    if (!loc(r.Ngay)) return;
    if (r.MaDinhDanh === '*') {
      if (r.Mon === '*') { ngayCu[r.Ngay] = 1; return; }
      ngay[r.Ngay] = 1; var g = giaiMaVo_(r.GhiChu);
      Object.keys(g).forEach(function (ma) { var x = o(ma, r.Mon); if (x[g[ma]] != null) x[g[ma]]++; });
      return;
    }
    var x = o(r.MaDinhDanh, r.Mon);                 // cách cũ
    if (r.TrangThai === 'Chưa nộp') x.N++; else x.cu++;
  });
  return { vo: vo, soNgay: Object.keys(ngay).length + Object.keys(ngayCu).length, soNgayCu: Object.keys(ngayCu).length };
}
// Câu nhận xét vở từ số lần: mức gặp nhiều nhất (bằng nhau thì lấy mức tốt hơn) + nhắc nộp vở. Giống hệt cauVoDem trên máy.
function cauVoDem_(v, soNgayCu) {
  if (!v) return soNgayCu ? 'Nộp vở đầy đủ, đúng hạn.' : '';
  var tot = '', n = 0;
  ['A', 'B', 'C', 'D', 'E'].forEach(function (k) { if (v[k] > n) { n = v[k]; tot = k; } });
  var thieu = v.N + v.cu, nhac = !thieu ? '' : v.cu ? (thieu <= 2 ? 'Đôi khi còn chưa nộp vở, chưa làm xong bài.' : 'Cần chú ý nộp vở đầy đủ, làm xong bài đúng hạn.')
    : (thieu <= 2 ? 'Đôi khi còn chưa nộp vở.' : 'Cần chú ý nộp vở đầy đủ, đúng hạn.');
  if (!tot) return nhac || (soNgayCu ? 'Nộp vở đầy đủ, đúng hạn.' : '');
  return CAU_VO[tot] + (nhac ? ' ' + nhac : '');
}
function cong_(a, b) { if (!b) return a; a = a || { A: 0, B: 0, C: 0, D: 0, E: 0, N: 0, cu: 0 }; Object.keys(b).forEach(function (k) { a[k] = (a[k] || 0) + b[k]; }); return a; }
function layNopVo(ngay) {
  ngay = ngay || homNay_();
  var k = key_(ngay), hom = {}, daLuu = {}, cuHom = null, truoc = {}, ngayTruoc = {};
  doc_('NopVo').forEach(function (r) {
    if (r.MaDinhDanh !== '*') { if (r.Ngay === ngay) (cuHom = cuHom || {})[r.MaDinhDanh + '|' + r.Mon] = r.TrangThai; return; }
    if (r.Mon === '*') { if (r.Ngay === ngay) cuHom = cuHom || {}; return; }
    var g = giaiMaVo_(r.GhiChu);
    if (r.Ngay === ngay) { hom[r.Mon] = g; daLuu[r.Mon] = true; return; }
    if (key_(r.Ngay) > k) return;
    Object.keys(g).forEach(function (ma) {            // mặc định: mức của lần chấm gần nhất trước đó (bỏ qua "Chưa nộp")
      var kk = ma + '|' + r.Mon; if (g[ma] === 'N' || (truoc[kk] && truoc[kk].k > key_(r.Ngay))) return;
      truoc[kk] = { k: key_(r.Ngay), m: g[ma] }; if (!ngayTruoc[r.Mon] || key_(ngayTruoc[r.Mon]) < key_(r.Ngay)) ngayTruoc[r.Mon] = r.Ngay;
    });
  });
  return { ngay: ngay, muc: MUC_VO, mon: MON_VO, daLuu: daLuu, ngayTruoc: ngayTruoc,
    hs: hocSinhDangHoc_(ngay).map(function (h) {
      var tt = {}, ma = h.MaDinhDanh;
      MON_VO.forEach(function (m) {
        if (daLuu[m]) tt[m] = hom[m][ma] || '';
        else if (cuHom && cuHom[ma + '|' + m] === 'Chưa nộp') tt[m] = 'N';
        else tt[m] = (truoc[ma + '|' + m] || {}).m || '';
      });
      return { ma: ma, ten: h.HoTen, tt: tt };
    }) };
}
// ds: [{ma, tt: {môn: mã mức}}] – mỗi môn 1 dòng; em chưa chọn mức thì không ghi
function luuNopVo(ngay, ds) {
  return voiKhoa_(function () {
    var luc = Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm'), moi = [], dem = {};
    MON_VO.forEach(function (m) {
      var ma = ds.filter(function (x) { return x.tt[m]; }).map(function (x) { return x.ma + ':' + x.tt[m]; });
      dem[m] = { co: ma.length, chua: ds.filter(function (x) { return x.tt[m] === 'N'; }).length };
      if (ma.length) moi.push(hang_('NopVo', { Ngay: ngay, MaDinhDanh: '*', Mon: m, TrangThai: 'Đã lưu', GhiChu: ma.join(' '), CapNhat: luc }));
    });
    thayTheTheo_('NopVo', function (r) { return r.Ngay !== ngay; }, moi);
    return 'Đã lưu nộp vở ' + ngay + ': ' + MON_VO.map(function (m) { return m + ' ' + dem[m].co + ' em' + (dem[m].chua ? ' (' + dem[m].chua + ' chưa nộp)' : ''); }).join('; ') + '.';
  });
}

// ============================================================================ GỢI Ý ĐỊNH KỲ TỪ ĐÁNH GIÁ HẰNG THÁNG (đợt 2)
// Giữa kì / cuối kì: mức môn học (T/H/C), 15 mục năng lực – phẩm chất và nhận xét lấy từ các tháng trong kì.
// Câu gợi ý dùng chung với màn đánh giá tháng (gửi xuống máy) để nhận xét kì ghép đúng như cô thấy ở từng tháng.
var GOI_Y_MON_THANG = {
  'Hoạt động trải nghiệm': {
    HTXS: ['Tích cực, sáng tạo trong các hoạt động trải nghiệm.', 'Tham gia hoạt động rất hào hứng, biết chia sẻ cảm nghĩ.', 'Vận dụng tốt điều đã học vào cuộc sống hằng ngày.', 'Chủ động, hợp tác tốt với bạn trong hoạt động.'],
    HTT: ['Tham gia tích cực các hoạt động trải nghiệm.', 'Biết chia sẻ, hợp tác với bạn khi hoạt động.', 'Thực hiện tốt các nhiệm vụ được giao.', 'Biết vận dụng điều đã học vào cuộc sống.'],
    HT: ['Tham gia đầy đủ các hoạt động trải nghiệm.', 'Thực hiện được nhiệm vụ khi có hướng dẫn.', 'Đôi khi còn rụt rè khi tham gia hoạt động.', 'Hoàn thành các hoạt động theo yêu cầu.'],
    CCG: ['Cần tích cực hơn khi tham gia hoạt động.', 'Cần mạnh dạn chia sẻ với thầy cô, bạn bè.', 'Cần hợp tác với bạn tốt hơn.', 'Cần chú ý thực hiện nhiệm vụ được giao.'] },
  'Đạo đức': {
    HTXS: ['Hiểu rõ và thực hiện rất tốt các chuẩn mực đã học.', 'Biết nhắc nhở bạn cùng thực hiện hành vi đúng.', 'Luôn lễ phép, kính trọng thầy cô, yêu quý bạn bè.', 'Vận dụng tốt bài học vào việc làm hằng ngày.'],
    HTT: ['Hiểu và thực hiện tốt các hành vi đã học.', 'Lễ phép, biết kính trọng thầy cô.', 'Biết yêu quý, giúp đỡ bạn bè.', 'Biết làm những việc phù hợp với bài học.'],
    HT: ['Biết các hành vi đạo đức đã học.', 'Thực hiện được hành vi đúng khi được nhắc nhở.', 'Biết lễ phép với thầy cô, bạn bè.', 'Hoàn thành các bài học theo yêu cầu.'],
    CCG: ['Cần thực hiện tốt hơn các hành vi đã học.', 'Cần lễ phép hơn với thầy cô, bạn bè.', 'Cần chú ý lắng nghe và làm theo bài học.', 'Cần cố gắng tự giác làm việc tốt.'] }
};
var CAU_NLD_THANG = {
  'Ngôn ngữ': { 'Tốt': ['Năng lực ngôn ngữ tốt: đọc, viết, nói rõ ràng.', 'Đọc, viết, diễn đạt tốt.'], 'Đạt': ['Năng lực ngôn ngữ đạt yêu cầu.', 'Đọc, viết đạt yêu cầu.'],
                'Cần cố gắng': ['Cần rèn thêm kĩ năng đọc, viết.', 'Cần luyện đọc, viết nhiều hơn.'] },
  'Tính toán': { 'Tốt': ['Năng lực tính toán tốt: tính nhanh, chính xác.', 'Tính toán nhanh, cẩn thận.'], 'Đạt': ['Năng lực tính toán đạt yêu cầu.', 'Tính toán đạt yêu cầu.'],
                 'Cần cố gắng': ['Cần rèn thêm kĩ năng tính toán.', 'Cần luyện tính toán cẩn thận hơn.'] }
};
var DIEM_THANG = { HTXS: 4, HTT: 3, HT: 2, CCG: 1, CHT: 1, 'Tốt': 3, 'Đạt': 2, 'Cần cố gắng': 1 };
var TC_KT_ = 'Hiểu kiến thức', TC_KN_ = 'Kỹ năng môn học';
function layDanhGiaThang(thang) {
  var d = duLieuThang_(thang); d.goiY = GOI_Y_MON_THANG; d.cauNLD = CAU_NLD_THANG; return d;
}
// Chọn 1 trong các mẫu câu – mỗi em 1 mẫu cố định (giống hệt trên máy)
function chon4_(ds, ma, k) { if (!ds || !ds.length) return ''; var n = parseInt(String(ma).slice(-4), 10) || 0; return ds[(n + k) % ds.length]; }
function noiCauRong_(ds, n) { return noiCau_(ds.filter(function (x) { return x; }), n); }
function cauVoThang_(t, ma, mon) {
  if (!t.soNgayVo || MON_VO.indexOf(mon) < 0) return '';
  return cauVoDem_((t.vo[ma] || {})[mon], t.soNgayVoCu);
}
function mucNLDThang_(rec) {
  var m = [rec.mucKT, rec.mucKN].filter(function (x) { return x; });
  if (!m.length) return '';
  if (m.indexOf('CCG') >= 0 || m.indexOf('CHT') >= 0) return 'Cần cố gắng';
  return m.every(function (x) { return x === 'HTXS' || x === 'HTT'; }) ? 'Tốt' : 'Đạt';
}
// Nhận xét 1 ô của 1 tháng đúng như màn đánh giá tháng hiện: phần cô sửa, hoặc bản app ghép
function nxThang_(t, ma, phan) {
  var du = t.du[ma] || {}, rec = du[phan] || {}, g = t.gioiHan;
  if (rec.nx) return rec.nx;
  if (phan === 'NLD') {
    var a = mucNLDThang_(du['Tiếng Việt'] || {}), b = mucNLDThang_(du['Toán'] || {});
    return noiCauRong_([a ? chon4_(CAU_NLD_THANG['Ngôn ngữ'][a], ma, 0) : '', b ? chon4_(CAU_NLD_THANG['Tính toán'][b], ma, 1) : ''], g);
  }
  if (phan === 'NLC' || phan === 'PC') {
    if (!rec.muc) return '';
    var ds = (phan === 'NLC' ? t.nxChung.nlc : t.nxChung.pc)[rec.muc] || [], c1 = chon4_(ds, ma, 0), c2 = chon4_(ds, ma, 2);
    return noiCauRong_([c1, c2 !== c1 ? c2 : ''], g);
  }
  if (t.yc && t.yc[phan]) {
    if (!rec.mucKT && !rec.mucKN) return '';
    return ghepYC_((t.ghi[ma] || {})[phan], t.yc[phan], cauVoThang_(t, ma, phan), g);
  }
  var lib = t.lib[phan];
  if (lib) {
    if (!rec.mucKT && !rec.mucKN) return '';
    return noiCauRong_([rec.mucKT ? chon4_((lib[TC_KT_] || {})[rec.mucKT], ma, 0) : '', rec.mucKN ? chon4_((lib[TC_KN_] || {})[rec.mucKN], ma, 1) : '', cauVoThang_(t, ma, phan)], g);
  }
  return rec.muc ? noiCauRong_([chon4_((GOI_Y_MON_THANG[phan] || {})[rec.muc], ma, 0), cauVoThang_(t, ma, phan)], g) : '';
}
function thangTrongKy_(pv) {
  var a = toDate_(pv.tu), b = toDate_(pv.den), ds = [], d = new Date(a.getFullYear(), a.getMonth(), 1);
  while (d <= b) { ds.push(d.getMonth() + 1); d = new Date(d.getFullYear(), d.getMonth() + 1, 1); }
  return ds;
}
function laCauVo_(c) { return /nộp vở|làm xong bài|trong vở/i.test(c); }
// Gộp câu nhận xét các tháng (tối đa 3 tháng gần nhất, bỏ trùng, bỏ câu nộp vở) + 1 câu nộp vở cả kì
function ghepNhieuThang_(T, ma, phan, voCaKy) {
  var cau = [];
  T.slice(-3).forEach(function (t) { tachCau_(nxThang_(t, ma, phan)).forEach(function (c) { if (!laCauVo_(c) && cau.indexOf(c) < 0) cau.push(c); }); });
  return noiCau_(voCaKy ? cau.concat([voCaKy]) : cau, GIOI_HAN_KY);
}
// null nếu kì này chưa có dữ liệu đánh giá tháng nào (khi đó dùng cách cũ – sổ theo dõi theo bài)
function goiYTuThang_(pv, hs, mon) {
  var coDL = thangCoGhiYC_(); doc_('DanhGiaThang').forEach(function (r) { if (r.NamHoc === pv.namHoc) coDL[+r.Thang] = 1; });
  var thang = thangTrongKy_(pv).filter(function (t) { return coDL[t]; });
  if (!thang.length) return null;
  var T = thang.map(duLieuThang_), tb = function (a) { return a.length ? a.reduce(function (x, y) { return x + y; }, 0) / a.length : 0; };
  var mucBon = function (x) { return !x ? '' : (x >= 2.8 ? 'T' : (x >= 1.8 ? 'Đ' : 'C')); };      // thang 4 mức (Toán, TV) → T/Đ/C
  var out = { thang: thang, du: {} };
  hs.forEach(function (h) {
    var ma = h.MaDinhDanh, o = { mon: {}, nlpc: null };
    mon.forEach(function (m) {
      var diem = [], soThang = 0, vKy = null, ngayVo = 0, ngayCu = 0;
      T.forEach(function (t) {
        var rec = (t.du[ma] || {})[m.TenMon] || {}, ds = (t.lib[m.TenMon] || t.yc[m.TenMon] ? [rec.mucKT, rec.mucKN] : [rec.muc]).filter(function (x) { return x; });
        if (ds.length) soThang++;
        ds.forEach(function (x) { diem.push(DIEM_THANG[x] || 0); });
        if (MON_VO.indexOf(m.TenMon) >= 0) { ngayVo += t.soNgayVo; ngayCu += t.soNgayVoCu; vKy = cong_(vKy, (t.vo[ma] || {})[m.TenMon]); }
      });
      var voKy = MON_VO.indexOf(m.TenMon) >= 0 && ngayVo ? cauVoDem_(vKy, ngayCu) : '';
      o.mon[m.TenMon] = { mucGoiY: mucMon_(tb(diem)), nx: ghepNhieuThang_(T, ma, m.TenMon, voKy), soGhi: soThang };
    });
    var dTV = [], dToan = [], dNLC = [], dPC = [], soThangNL = 0;
    T.forEach(function (t) {
      var d = t.du[ma] || {}, co = false;
      ['mucKT', 'mucKN'].forEach(function (k) {
        if ((d['Tiếng Việt'] || {})[k]) { dTV.push(DIEM_THANG[d['Tiếng Việt'][k]]); co = true; }
        if ((d['Toán'] || {})[k]) { dToan.push(DIEM_THANG[d['Toán'][k]]); co = true; }
      });
      if ((d.NLC || {}).muc) { dNLC.push(DIEM_THANG[d.NLC.muc]); co = true; }
      if ((d.PC || {}).muc) { dPC.push(DIEM_THANG[d.PC.muc]); co = true; }
      if (co) soThangNL++;
    });
    var mucGoiY = {}; DS_NLPC.forEach(function (t) { mucGoiY[t] = ''; });       // khoa học, công nghệ, tin học, thẩm mĩ, thể chất: cô chọn
    mucGoiY['Ngôn ngữ'] = mucBon(tb(dTV)); mucGoiY['Tính toán'] = mucBon(tb(dToan));
    THU_TU_CHUNG.forEach(function (t) { mucGoiY[t] = mucNLPC_(tb(dNLC)); });
    THU_TU_PC.forEach(function (t) { mucGoiY[t] = mucNLPC_(tb(dPC)); });
    o.nlpc = { mucGoiY: mucGoiY, nlc: ghepNhieuThang_(T, ma, 'NLC'), nld: ghepNhieuThang_(T, ma, 'NLD'), pc: ghepNhieuThang_(T, ma, 'PC'), soGhi: soThangNL };
    out.du[ma] = o;
  });
  return out;
}
