// ============================================================================
// ĐỌC FILE TRÊN TRÌNH DUYỆT (không cần quyền đặc biệt của Google)
//   - Sổ báo giảng (.docx)       : dùng JSZip để mở file, đọc word/document.xml
//   - Danh sách học sinh (.xlsx) : dùng SheetJS (XLSX)
//   - Thư viện nhận xét (.xlsx)  : dùng SheetJS (XLSX)
// Các hàm ở đây là hàm thuần (không gọi máy chủ) để kiểm tra được bằng Node.
// ============================================================================

var Parsers = (function () {
  function xmlText(xml) {
    return xml.replace(/<\/w:p>/g, ' / ').replace(/<[^>]+>/g, ' ')
      .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'")
      .replace(/\s+/g, ' ').replace(/^[\s\/]+|[\s\/]+$/g, '');
  }
  function pad2(n) { n = String(n); return n.length < 2 ? '0' + n : n; }
  function namHocYears(namHoc) {           // '2026-2027' -> [2026, 2027]
    var m = /(\d{4})\s*-\s*(\d{4})/.exec(namHoc || '');
    return m ? [+m[1], +m[2]] : null;
  }

  // ---------------------------------------------------------------- Sổ báo giảng
  // xml: nội dung word/document.xml ; namHoc: '2026-2027' (để suy ra năm khi ô ngày không ghi năm)
  function parseSoBaoGiangXml(xml, namHoc) {
    var years = namHocYears(namHoc);
    var parts = xml.match(/<w:p[ >][\s\S]*?<\/w:p>|<w:tr[ >][\s\S]*?<\/w:tr>/g) || [];
    var rows = [], week = null, chuDe = '', day = null, buoi = null, canhBao = [];
    var chuDeTheoTuan = {}, ngayTieuDe = {};
    function tieuDeTuan(t) {            // đọc "CHỦ ĐỀ …" và "(Từ ngày … đến ngày …)" ở tiêu đề tuần
      t = t.normalize ? t.normalize('NFC') : t;
      var mc = /(CHỦ ĐỀ\s*\d*\s*[.:]?\s*[^]*?)(?:\s+THỨ(?=\s|$)|$)/.exec(t);
      if (mc && !chuDeTheoTuan[week]) chuDeTheoTuan[week] = mc[1].replace(/\s+/g, ' ').replace(/[\s\/]+$/, '').trim();
      var md = /Từ ngày\s*([\d\/ ]+?)\s*đến ngày\s*([\d\/ ]+)/i.exec(t);
      if (md && !ngayTieuDe[week]) ngayTieuDe[week] = [md[1].replace(/\s/g, ''), md[2].replace(/\s/g, '')];
    }
    parts.forEach(function (part) {
      var t = xmlText(part);
      var mw = /U[ẦẤA]N\s*:?\s*(\d+)/.exec(t);
      if (part.indexOf('<w:tr') === 0) {
        var cells = (part.match(/<w:tc>[\s\S]*?<\/w:tc>/g) || []).map(xmlText);
        if (mw && /TIẾT/i.test(t)) { week = +mw[1]; tieuDeTuan(t); }
        if (cells.length < 5) return;
        if (/\d{1,2}\/\d{1,2}/.test(cells[0])) day = cells[0];
        var b = cells[1].replace(/[\s\/]/g, '').toUpperCase();
        if (b === 'SÁNG' || b === 'CHIỀU') buoi = b === 'SÁNG' ? 'Sáng' : 'Chiều';
        var tiet = cells[2].replace(/[\s\/]/g, '');
        if (!/^\d{1,2}$/.test(tiet) || !day) return;
        var d = /(\d{1,2})\/(\d{1,2})(?:\D+(\d{4}))?/.exec(day);
        if (!d) return;
        var mm = +d[2], yy = d[3] ? +d[3] : (years ? (mm >= 8 ? years[0] : years[1]) : null);
        var thu = day.split('/')[0].replace(/\d+$/, '').trim();
        rows.push({
          tuan: week, ngay: pad2(d[1]) + '/' + pad2(mm) + '/' + (yy || ''), thu: thu, buoi: buoi, tiet: +tiet,
          mon: cells[3].replace(/\s*\/\s*/g, ' ').trim().toUpperCase(),
          bai: cells[4].replace(/\s*\/\s*/g, ' ').replace(/\s+/g, ' ').trim()
        });
      } else {
        if (mw) week = +mw[1];
        if (week != null) tieuDeTuan(t);
      }
    });
    // Danh sách tuần: lấy ngày thực tế trong bảng (đáng tin hơn ngày ghi ở tiêu đề tuần)
    var tuan = {};
    rows.forEach(function (r) {
      var k = key(r.ngay);
      if (!tuan[r.tuan]) tuan[r.tuan] = { tuan: r.tuan, tu: r.ngay, den: r.ngay, kTu: k, kDen: k, chuDe: chuDeTheoTuan[r.tuan] || '' };
      var w = tuan[r.tuan];
      if (k < w.kTu) { w.kTu = k; w.tu = r.ngay; }
      if (k > w.kDen) { w.kDen = k; w.den = r.ngay; }
    });
    var dsTuan = Object.keys(tuan).map(function (k) { return tuan[k]; }).sort(function (a, b) { return a.tuan - b.tuan; });
    for (var i = 1; i < dsTuan.length; i++) {
      if (dsTuan[i].tuan !== dsTuan[i - 1].tuan + 1) canhBao.push('Thiếu tuần giữa tuần ' + dsTuan[i - 1].tuan + ' và ' + dsTuan[i].tuan + '.');
      if (dsTuan[i].kTu <= dsTuan[i - 1].kDen) canhBao.push('Ngày của tuần ' + dsTuan[i].tuan + ' chồng lên tuần trước.');
    }
    if (rows.some(function (r) { return !/\d{4}$/.test(r.ngay); })) canhBao.push('Có ngày không xác định được năm – kiểm tra năm học.');
    Object.keys(ngayTieuDe).forEach(function (w) {
      var t = tuan[w]; if (!t) return;
      var lech = function (ghi, thuc, choPhep) {   // choPhep: số ngày được lệch (tiêu đề hay ghi tới thứ Bảy)
        var m = /(\d{1,2})\/(\d{1,2})/.exec(ghi); if (!m) return false;
        var p = thuc.split('/'); return +m[2] !== +p[1] || Math.abs(+m[1] - +p[0]) > choPhep;
      };
      if (lech(ngayTieuDe[w][0], t.tu, 0) || lech(ngayTieuDe[w][1], t.den, 1))
        canhBao.push('Tuần ' + w + ': tiêu đề ghi "' + ngayTieuDe[w].join(' – ') + '" nhưng trong bảng là ' + t.tu + ' – ' + t.den + ' (app dùng ngày trong bảng).');
    });
    dsTuan.forEach(function (w) { delete w.kTu; delete w.kDen; });
    return { tiet: rows, tuan: dsTuan, canhBao: canhBao };
  }
  function key(dmy) { var p = String(dmy).split('/'); return (+p[2] || 0) * 10000 + (+p[1]) * 100 + (+p[0]); }

  // ---------------------------------------------------------------- Danh sách học sinh
  // aoa: mảng các dòng (mỗi dòng là mảng ô) của sheet đầu tiên
  function parseDanhSachHocSinh(aoa) {
    var info = { truong: '', lop: '', namHoc: '' }, hdr = -1;
    for (var i = 0; i < Math.min(aoa.length, 30); i++) {
      var seen = {}, line = (aoa[i] || []).map(function (v) { return v == null ? '' : String(v).trim(); })
        .filter(function (v) { if (!v || seen[v]) return false; seen[v] = 1; return true; }).join(' ');
      var ml = /DANH SÁCH HỌC SINH LỚP\s*([^\s]+)/i.exec(line); if (ml) info.lop = ml[1];
      var mn = /Năm học:?\s*(\d{4}\s*-\s*\d{4})/i.exec(line); if (mn) info.namHoc = mn[1].replace(/\s/g, '');
      if (/TIỂU HỌC|TRƯỜNG/i.test(line) && !info.truong && !/DANH SÁCH/i.test(line)) info.truong = line.trim();
      if (/định danh/i.test(line) && hdr < 0) hdr = i;
    }
    if (hdr < 0) throw new Error('Không tìm thấy dòng tiêu đề có cột "Mã định danh".');
    var H = aoa[hdr].map(function (v) { return String(v || '').toLowerCase().trim(); });
    function col(re) { for (var j = 0; j < H.length; j++) if (re.test(H[j])) return j; return -1; }
    var c = { ma: col(/định danh|mã học sinh/), ten: col(/họ.*tên/), ns: col(/ngày sinh/), gt: col(/giới tính/), dt: col(/dân tộc/),
              tt: col(/trạng thái/), sdt: col(/sđt|điện thoại/), sb: col(/số buổi/) };
    var hs = [];
    for (var r = hdr + 1; r < aoa.length; r++) {
      var row = aoa[r] || [], ma = row[c.ma] == null ? '' : String(row[c.ma]).trim();
      if (!/^\d{6,}$/.test(ma)) continue;
      var g = function (k) { return c[k] >= 0 && row[c[k]] != null ? String(row[c[k]]).trim() : ''; };
      hs.push({ maDinhDanh: ma, hoTen: g('ten'), ngaySinh: g('ns'), gioiTinh: g('gt'), danToc: g('dt'), trangThai: g('tt') || 'Đang học', sdt: g('sdt'), soBuoi: g('sb') });
    }
    info.khoi = (/^(\d)/.exec(info.lop) || [])[1] || '';
    return { info: info, hocSinh: hs };
  }
  function laChuyenDi(trangThai) { return /chuyển đi|thôi học|nghỉ học/i.test(trangThai || ''); }

  // ---------------------------------------------------------------- Thư viện nhận xét
  // sheets: { tenSheet: aoa } ; trả về {thongTin, mon: [{ten, dong: [...]}], chung: [...]}
  // Loại file thư viện theo sheet ThongTin: 'ThuVienNhanXet' (theo bài) hoặc 'ThuVienThang' (theo tháng)
  function loaiThuVien(sheets) {
    var loai = ''; (sheets['ThongTin'] || []).forEach(function (r) { if (r && String(r[0]).trim() === 'Loai') loai = String(r[1] || '').trim(); });
    return loai;
  }
  // Thư viện THEO THÁNG: mỗi sheet môn có cột Tháng – Tiêu chí – Mức độ ("HTXS – Hoàn thành xuất sắc") – Mẫu – Nội dung nhận xét
  function parseThuVienThang(sheets) {
    var tt = {};
    (sheets['ThongTin'] || []).forEach(function (r) { if (r && r[0]) tt[String(r[0]).trim()] = r[1] == null ? '' : String(r[1]).trim(); });
    if (tt.Loai !== 'ThuVienThang') throw new Error('File không phải thư viện nhận xét THEO THÁNG (sheet ThongTin, Loai = ThuVienThang).');
    if (String(tt.PhienBanDinhDang) === '2') return parseThuVienYeuCau(sheets, tt);
    var out = { thongTin: tt, mon: [] };
    Object.keys(sheets).forEach(function (name) {
      var aoa = sheets[name]; if (!aoa || !aoa.length) return;
      var H = aoa[0].map(function (v) { return String(v || '').trim(); }), ix = function (t) { return H.indexOf(t); };
      if (ix('Tháng') < 0 || ix('Tiêu chí') < 0 || ix('Mức độ') < 0 || ix('Nội dung nhận xét') < 0) return;
      var dong = [];
      for (var k = 1; k < aoa.length; k++) {
        var x = aoa[k]; if (!x || !x[ix('Nội dung nhận xét')]) continue;
        dong.push({ thang: s(x[ix('Tháng')]), tieuChi: s(x[ix('Tiêu chí')]), mucDo: s(x[ix('Mức độ')]).split(/\s+[–-]\s+/)[0].trim(),
                    mau: s(x[ix('Mẫu')]), noiDung: s(x[ix('Nội dung nhận xét')]) });
      }
      if (dong.length) out.mon.push({ ten: name.split('–')[0].trim(), dong: dong });
    });
    if (!out.mon.length) throw new Error('Không thấy sheet môn nào có các cột Tháng – Tiêu chí – Mức độ – Nội dung nhận xét.');
    return out;
  }
  // Bản 2 (theo ý cô): sheet môn có cột Tháng – Nội dung (Kiến thức/Kỹ năng) – Mã – Yêu cầu cần đạt – Mức độ – Mẫu – Nội dung nhận xét;
  // căn cứ của từng yêu cầu lấy ở sheet "Yêu cầu theo tháng"
  function parseThuVienYeuCau(sheets, tt) {
    var so = function (v) { return (/(\d+)/.exec(s(v)) || [])[1] || ''; }, canCu = {};
    var yc = sheets['Yêu cầu theo tháng'];
    if (yc && yc.length) {
      var H0 = yc[0].map(function (v) { return String(v || '').trim(); }), j = function (t) { for (var i = 0; i < H0.length; i++) if (H0[i].indexOf(t) === 0) return i; return -1; };
      for (var r = 1; r < yc.length; r++) { var y = yc[r]; if (!y) continue; canCu[s(y[j('Môn')]) + '|' + so(y[j('Tháng')]) + '|' + s(y[j('Mã')])] = s(y[j('Căn cứ')]); }
    }
    var out = { thongTin: tt, mon: [], v2: true };
    Object.keys(sheets).forEach(function (name) {
      var aoa = sheets[name]; if (!aoa || !aoa.length || name === 'Yêu cầu theo tháng') return;
      var H = aoa[0].map(function (v) { return String(v || '').trim(); }), ix = function (t) { return H.indexOf(t); };
      if (ix('Tháng') < 0 || ix('Nội dung') < 0 || ix('Yêu cầu cần đạt') < 0 || ix('Mức độ') < 0 || ix('Nội dung nhận xét') < 0) return;
      var ten = name.split('–')[0].trim(), dong = [];
      for (var k = 1; k < aoa.length; k++) {
        var x = aoa[k]; if (!x || !s(x[ix('Nội dung nhận xét')]) || !s(x[ix('Yêu cầu cần đạt')])) continue;
        var th = so(x[ix('Tháng')]), ma = s(x[ix('Mã')]);
        dong.push({ thang: th, noiDung: /kỹ năng/i.test(s(x[ix('Nội dung')])) ? 'Kỹ năng' : 'Kiến thức', ma: ma, yeuCau: s(x[ix('Yêu cầu cần đạt')]),
                    canCu: canCu[ten + '|' + th + '|' + ma] || '', mucDo: s(x[ix('Mức độ')]).split(/\s+[–-]\s+/)[0].trim().toUpperCase(),
                    mau: s(x[ix('Mẫu')]), cau: s(x[ix('Nội dung nhận xét')]) });
      }
      if (dong.length) out.mon.push({ ten: ten, dong: dong });
    });
    if (!out.mon.length) throw new Error('Không thấy sheet môn nào có các cột Tháng – Nội dung – Yêu cầu cần đạt – Mức độ – Nội dung nhận xét.');
    var sai = [];
    out.mon.forEach(function (m) { m.dong.forEach(function (d) { if (['HTXS', 'HTT', 'HT', 'CHT'].indexOf(d.mucDo) < 0) sai.push(m.ten + ' – ' + d.yeuCau + ': mức "' + d.mucDo + '"'); }); });
    if (sai.length) throw new Error('Có mức độ không đúng (chỉ dùng HTXS, HTT, HT, CHT): ' + sai.slice(0, 3).join('; ') + (sai.length > 3 ? '…' : ''));
    return out;
  }
  function parseThuVien(sheets) {
    var tt = {};
    (sheets['ThongTin'] || []).forEach(function (r) { if (r && r[0]) tt[String(r[0]).trim()] = r[1] == null ? '' : String(r[1]).trim(); });
    if (tt.Loai !== 'ThuVienNhanXet') throw new Error('File không có sheet "ThongTin" của thư viện nhận xét (hãy tạo file bằng skill tao-thu-vien-nhan-xet).');
    var out = { thongTin: tt, mon: [], chung: [] };
    Object.keys(sheets).forEach(function (name) {
      var aoa = sheets[name]; if (!aoa || !aoa.length) return;
      var H = aoa[0].map(function (v) { return String(v || '').trim(); });
      var ix = function (t) { return H.indexOf(t); };
      if (name === 'Nhận xét chung') {
        for (var i = 1; i < aoa.length; i++) {
          var r = aoa[i]; if (!r || !r[ix('Nội dung nhận xét')]) continue;
          out.chung.push({ nhom: s(r[ix('Nhóm')]), mucDo: s(r[ix('Mức độ')]), mau: s(r[ix('Mẫu')]), noiDung: s(r[ix('Nội dung nhận xét')]), pcnl: s(r[ix('PC/NL liên quan')]) });
        }
        return;
      }
      if (ix('Hoạt động') < 0 || ix('Nội dung nhận xét') < 0 || ix('Bài') < 0) return;
      var dong = [];
      for (var k = 1; k < aoa.length; k++) {
        var x = aoa[k]; if (!x || !x[ix('Nội dung nhận xét')]) continue;
        dong.push({ tuan: s(x[ix('Tuần')]), chuDe: s(x[ix('Chủ đề / Chủ điểm')] != null ? x[ix('Chủ đề / Chủ điểm')] : x[ix('Chủ điểm')]),
                    bai: s(x[ix('Bài')]), trangSGV: s(x[ix('Trang SGV')]), hoatDong: s(x[ix('Hoạt động')]),
                    canCu: s(x[ix('Căn cứ (Sổ báo giảng / YCCĐ SGV)')]), mucDo: s(x[ix('Mức độ')]), mau: s(x[ix('Mẫu')]),
                    noiDung: s(x[ix('Nội dung nhận xét')]), pcnl: s(x[ix('PC/NL liên quan')]) });
      }
      var tenMon = s(aoa[1] && aoa[1][ix('Môn')]) || name.split('–')[0].trim();
      if (dong.length) out.mon.push({ ten: tenMon, dong: dong });
    });
    if (!out.mon.length) throw new Error('Không thấy sheet môn học nào đúng mẫu cột.');
    return out;
  }
  function s(v) { return v == null ? '' : String(v).trim(); }

  return { parseSoBaoGiangXml: parseSoBaoGiangXml, parseDanhSachHocSinh: parseDanhSachHocSinh,
           parseThuVien: parseThuVien, parseThuVienThang: parseThuVienThang, loaiThuVien: loaiThuVien, laChuyenDi: laChuyenDi, key: key };
})();

// ============================================================================
// ĐIỀN DỮ LIỆU VÀO FILE BIỂU MẪU CỦA TRƯỜNG (.xlsx) – GIỮ NGUYÊN ĐỊNH DẠNG
// Cách làm: mở file như 1 tệp nén (JSZip), sửa trực tiếp nội dung XML của từng sheet:
// chỉ thay giá trị của các ô cần điền (giữ nguyên kiểu định dạng s="…" của ô), không đụng phần khác.
// Ghép dòng theo Mã định danh (cột B). Hàm thuần, kiểm tra được bằng Node.
// ============================================================================
var BieuMau = (function () {
  function giaiMa(s) { return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&'); }
  function maHoa(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
  function boBOM(s) { return s.charCodeAt(0) === 0xFEFF ? s.slice(1) : s; }
  function tienTo(xml) { var m = /<(\w+:)?worksheet[\s>]/.exec(xml); return m && m[1] ? m[1] : ''; }
  function chuSoCot(c) { var n = 0; for (var i = 0; i < c.length; i++) n = n * 26 + (c.charCodeAt(i) - 64); return n; }

  // Danh sách sheet: [{ten, duongDan}]
  function dsSheet(workbookXml, relsXml) {
    var rels = {};
    (relsXml.match(/<(?:\w+:)?Relationship\b[^>]*>/g) || []).forEach(function (r) {
      var id = /\bId="([^"]+)"/.exec(r), t = /\bTarget="([^"]+)"/.exec(r);
      if (id && t) rels[id[1]] = t[1].replace(/^\/?xl\//, '');
    });
    return (workbookXml.match(/<(?:\w+:)?sheet\b[^>]*>/g) || []).map(function (s) {
      var ten = /\bname="([^"]*)"/.exec(s), id = /\br:id="([^"]+)"/.exec(s);
      return { ten: giaiMa(ten ? ten[1] : ''), duongDan: 'xl/' + (rels[id ? id[1] : ''] || '') };
    });
  }
  function sharedStrings(xml) {
    if (!xml) return [];
    return (xml.match(/<(?:\w+:)?si\b[\s\S]*?<\/(?:\w+:)?si>/g) || []).map(function (si) {
      return giaiMa((si.match(/<(?:\w+:)?t\b[^>]*>[\s\S]*?<\/(?:\w+:)?t>/g) || []).map(function (t) { return t.replace(/<[^>]+>/g, ''); }).join(''));
    });
  }
  function giaTriO(cXml, sst) {
    var v = /<(?:\w+:)?v>([\s\S]*?)<\/(?:\w+:)?v>/.exec(cXml);
    if (/\bt="s"/.test(cXml)) return v ? (sst[+v[1]] || '') : '';
    if (/\bt="inlineStr"/.test(cXml)) return giaiMa((cXml.match(/<(?:\w+:)?t\b[^>]*>[\s\S]*?<\/(?:\w+:)?t>/g) || []).map(function (t) { return t.replace(/<[^>]+>/g, ''); }).join(''));
    return v ? giaiMa(v[1]) : '';
  }
  // Ô chữ ghi theo kiểu "shared string" (t="s") giống hệt file mẫu, vì nhiều hệ thống nhập liệu chỉ đọc kiểu này.
  // themChuoi(text) trả về số thứ tự chuỗi trong bảng sharedStrings; không có bảng thì dùng kiểu chữ trực tiếp.
  var _themChuoi = null;
  function oMoi(p, ref, s, text) {
    var dau = '<' + p + 'c r="' + ref + '"' + (s ? ' s="' + s + '"' : '');
    if (_themChuoi) return dau + ' t="s"><' + p + 'v>' + _themChuoi(String(text)) + '</' + p + 'v></' + p + 'c>';
    return dau + ' t="inlineStr"><' + p + 'is><' + p + 't xml:space="preserve">' + maHoa(text) + '</' + p + 't></' + p + 'is></' + p + 'c>';
  }
  // Ghi các ô vào 1 dòng: giaTri = {G: 'nội dung', E: '9', …}. Ô có sẵn: giữ kiểu s; ô chưa có: chèn đúng thứ tự cột.
  function ghiDong(rowXml, soDong, giaTri, p, sMacDinh) {
    var cells = rowXml.match(/<(?:\w+:)?c\b[^>]*?(?:\/>|>[\s\S]*?<\/(?:\w+:)?c>)/g) || [];
    var ds = cells.map(function (c) { var r = /\br="([A-Z]+)\d+"/.exec(c); return { cot: r ? r[1] : '', xml: c }; });
    Object.keys(giaTri).forEach(function (cot) {
      if (giaTri[cot] == null) return;
      var i = ds.map(function (x) { return x.cot; }).indexOf(cot);
      if (i >= 0) {
        var s = /\bs="(\d+)"/.exec(ds[i].xml);
        ds[i].xml = oMoi(p, cot + soDong, s ? s[1] : '', giaTri[cot]);
      } else {
        var o = { cot: cot, xml: oMoi(p, cot + soDong, sMacDinh, giaTri[cot]) }, k = 0;
        while (k < ds.length && chuSoCot(ds[k].cot) < chuSoCot(cot)) k++;
        ds.splice(k, 0, o);
      }
    });
    var mo = /^<(?:\w+:)?row\b[^>]*?(?:\/>|>)/.exec(rowXml)[0];
    if (/\/>$/.test(mo)) mo = mo.replace(/\s*\/>$/, '>');
    return mo + ds.map(function (x) { return x.xml; }).join('') + '</' + p + 'row>';
  }
  // Điền 1 sheet. duLieu: { maDinhDanh: {B:…, C:…, G:…} }; cotMa: cột chứa mã (B); dongDau: dòng dữ liệu đầu tiên (3)
  // themNeuThieu: học sinh có trong duLieu mà không có trong file thì thêm dòng mới ở cuối (dùng giá trị duLieu[ma].dongMoi)
  function dienSheet(xml, sst, duLieu, tuyChon) {
    xml = boBOM(xml);
    var p = tienTo(xml), cotMa = tuyChon.cotMa || 'B', dongDau = tuyChon.dongDau || 3;
    var daDien = {}, khongCoTrongApp = [], soDongCuoi = dongDau - 1, sMau = '';
    xml = xml.replace(/<(?:\w+:)?row\b[^>]*?(?:\/>|>[\s\S]*?<\/(?:\w+:)?row>)/g, function (row) {
      var n = +(/\br="(\d+)"/.exec(row) || [0, 0])[1];
      soDongCuoi = Math.max(soDongCuoi, n);
      if (n < dongDau) return row;
      var oMa = (row.match(new RegExp('<(?:\\w+:)?c\\b[^>]*\\br="' + cotMa + n + '"[^>]*?(?:\\/>|>[\\s\\S]*?<\\/(?:\\w+:)?c>)')) || [])[0];
      var ma = oMa ? giaTriO(oMa, sst).trim() : '';
      if (!sMau && oMa) { var sm = /\bs="(\d+)"/.exec(oMa); if (sm) sMau = sm[1]; }
      if (!ma) return row;
      if (!duLieu[ma]) { khongCoTrongApp.push(ma); return row; }
      daDien[ma] = 1;
      return ghiDong(row, n, duLieu[ma].o, p, sMau);
    });
    var them = [];
    if (tuyChon.themNeuThieu) {
      Object.keys(duLieu).forEach(function (ma) {
        if (daDien[ma] || !duLieu[ma].dongMoi) return;
        soDongCuoi++;
        var gt = {}; Object.keys(duLieu[ma].dongMoi).forEach(function (k) { gt[k] = duLieu[ma].dongMoi[k]; });
        Object.keys(duLieu[ma].o).forEach(function (k) { gt[k] = duLieu[ma].o[k]; });
        them.push(ghiDong('<' + p + 'row r="' + soDongCuoi + '">', soDongCuoi, gt, p, sMau));
        daDien[ma] = 1;
      });
      if (them.length) {
        if (new RegExp('<' + p + 'sheetData\\s*/>').test(xml)) xml = xml.replace(new RegExp('<' + p + 'sheetData\\s*/>'), '<' + p + 'sheetData>' + them.join('') + '</' + p + 'sheetData>');
        else xml = xml.replace(new RegExp('</' + p + 'sheetData>'), them.join('') + '</' + p + 'sheetData>');
      }
    }
    // Ô có nhiều dòng chữ: bỏ chiều cao dòng cố định không cần thiết thì Excel tự giãn – giữ nguyên, không đổi.
    xml = xml.replace(new RegExp('<' + p + 'dimension ref="[^"]*"\\s*/>'), '<' + p + 'dimension ref="A1:' + (tuyChon.cotCuoi || 'M') + Math.max(soDongCuoi, 2) + '"/>');
    return { xml: xml, daDien: Object.keys(daDien), khongCoTrongApp: khongCoTrongApp, themMoi: them.length };
  }
  // zip: đối tượng JSZip đã mở file mẫu. bang: { 'Tên sheet': duLieu }. Trả về báo cáo; zip được sửa trực tiếp.
  function dienFile(zip, bang, tuyChon) {
    return Promise.all([zip.file('xl/workbook.xml').async('string'), zip.file('xl/_rels/workbook.xml.rels').async('string'),
                        zip.file('xl/sharedStrings.xml') ? zip.file('xl/sharedStrings.xml').async('string') : Promise.resolve('')])
      .then(function (k) {
        var sheets = dsSheet(boBOM(k[0]), boBOM(k[1])), sstXml = boBOM(k[2]), sst = sharedStrings(sstXml), baoCao = { sheet: {}, thieuSheet: [] };
        var chuoiMoi = [], viTri = {};
        sst.forEach(function (t, i) { if (!(t in viTri)) viTri[t] = i; });
        _themChuoi = sstXml ? function (t) {
          if (t in viTri) return viTri[t];
          viTri[t] = sst.length + chuoiMoi.length; chuoiMoi.push(t); return viTri[t];
        } : null;
        var viec = Object.keys(bang).map(function (ten) {
          var sh = sheets.filter(function (x) { return x.ten.trim().toLowerCase() === ten.trim().toLowerCase(); })[0];
          if (!sh || !zip.file(sh.duongDan)) { baoCao.thieuSheet.push(ten); return Promise.resolve(); }
          return zip.file(sh.duongDan).async('string').then(function (xml) {
            var kq = dienSheet(xml, sst, bang[ten], tuyChon || {});
            zip.file(sh.duongDan, kq.xml);
            baoCao.sheet[ten] = { daDien: kq.daDien.length, khongCoTrongApp: kq.khongCoTrongApp, themMoi: kq.themMoi };
          });
        });
        return Promise.all(viec).then(function () {
          _themChuoi = null;
          if (chuoiMoi.length) {      // nối các chuỗi mới vào cuối bảng sharedStrings, cập nhật số đếm
            var p = (/<(\w+:)?sst[\s>]/.exec(sstXml) || [])[1] || '';
            var them = chuoiMoi.map(function (t) { return '<' + p + 'si><' + p + 't xml:space="preserve">' + maHoa(t) + '</' + p + 't></' + p + 'si>'; }).join('');
            var tong = sst.length + chuoiMoi.length;
            sstXml = sstXml.replace(new RegExp('</' + p + 'sst>\\s*$'), them + '</' + p + 'sst>')
              .replace(/(<(?:\w+:)?sst\b[^>]*?\scount=")\d+"/, '$1' + tong + '"')
              .replace(/(<(?:\w+:)?sst\b[^>]*?\suniqueCount=")\d+"/, '$1' + tong + '"');
            zip.file('xl/sharedStrings.xml', sstXml);
          }
          baoCao.dsSheet = sheets.map(function (x) { return x.ten; }); return baoCao;
        });
      });
  }
  // Đọc nhanh: lấy các mã định danh có trong 1 sheet (để kiểm tra file mẫu đúng lớp)
  function docMa(zip, tenSheet) {
    return Promise.all([zip.file('xl/workbook.xml').async('string'), zip.file('xl/_rels/workbook.xml.rels').async('string'),
                        zip.file('xl/sharedStrings.xml') ? zip.file('xl/sharedStrings.xml').async('string') : Promise.resolve('')])
      .then(function (k) {
        var sheets = dsSheet(boBOM(k[0]), boBOM(k[1])), sst = sharedStrings(boBOM(k[2]));
        var sh = sheets.filter(function (x) { return !tenSheet || x.ten === tenSheet; })[0];
        return zip.file(sh.duongDan).async('string').then(function (xml) {
          var ma = []; (boBOM(xml).match(/<(?:\w+:)?c\b[^>]*\br="B\d+"[^>]*?(?:\/>|>[\s\S]*?<\/(?:\w+:)?c>)/g) || []).forEach(function (c) {
            var v = giaTriO(c, sst).trim(); if (/^\d{6,}$/.test(v)) ma.push(v);
          });
          var oE = /<(?:\w+:)?c\b[^>]*\br="E3"[^>]*?(?:\/>|>[\s\S]*?<\/(?:\w+:)?c>)/.exec(boBOM(xml));   // cột Tháng của dòng đầu
          return { sheets: sheets.map(function (x) { return x.ten; }), ma: ma, thang: oE ? giaTriO(oE[0], sst).trim() : '' };
        });
      });
  }
  // Nhận diện loại biểu mẫu qua chữ ở các dòng tiêu đề (1–7) của từng sheet, để không điền nhầm file.
  // loai: 'thang' | 'mon' | 'nlpc' | 'gvcn' → { ok, sheets: [tên sheet đúng mẫu], chiHuongDan }
  var DAU_HIEU = {
    thang: ['Tháng', 'Nhận xét năng lực chung'],
    mon: ['Mức đạt được', 'Thời điểm đánh giá'],
    nlpc: ['Tự chủ và tự học', 'Thời điểm đánh giá'],
    gvcn: ['Nhận xét của GVCN']
  };
  function kiemTraMau(zip, loai) {
    var dauHieu = DAU_HIEU[loai];
    return Promise.all([zip.file('xl/workbook.xml').async('string'), zip.file('xl/_rels/workbook.xml.rels').async('string'),
                        zip.file('xl/sharedStrings.xml') ? zip.file('xl/sharedStrings.xml').async('string') : Promise.resolve('')])
      .then(function (k) {
        var sheets = dsSheet(boBOM(k[0]), boBOM(k[1])), sst = sharedStrings(boBOM(k[2]));
        var tenDung = [], conLai = sheets.filter(function (s) { return !/huongdan|gioi_tinh/i.test(s.ten); });
        return Promise.all(conLai.map(function (sh) {
          if (!zip.file(sh.duongDan)) return null;
          return zip.file(sh.duongDan).async('string').then(function (xml) {
            var chu = (boBOM(xml).match(/<(?:\w+:)?row\b[^>]*\br="[1-7]"[^>]*>[\s\S]*?<\/(?:\w+:)?row>/g) || []).map(function (row) {
              return (row.match(/<(?:\w+:)?c\b[^>]*?(?:\/>|>[\s\S]*?<\/(?:\w+:)?c>)/g) || []).map(function (c) { return giaTriO(c, sst); }).join(' | ');
            }).join(' | ').toLowerCase();
            if (dauHieu.every(function (d) { return chu.indexOf(d.toLowerCase()) >= 0; })) tenDung.push(sh.ten);
          });
        })).then(function () { return { ok: tenDung.length > 0, sheets: tenDung, chiHuongDan: !conLai.length }; });
      });
  }
  return { dienFile: dienFile, dienSheet: dienSheet, docMa: docMa, dsSheet: dsSheet, sharedStrings: sharedStrings, kiemTraMau: kiemTraMau };
})();
if (typeof module !== 'undefined') { module.exports = Parsers; module.exports.BieuMau = BieuMau; }
