// ============================================================================ DÙNG CHUNG máy chủ + trình duyệt
// build.py ghép file này vào cuối Code.gs và vào Index.html, nên 2 bên suy ra giống hệt nhau. Chỉ dùng JavaScript thuần.
//
// SUY NĂNG LỰC – PHẨM CHẤT TỪ KẾT QUẢ CÁC MÔN (Thông tư 27/2020: giáo viên căn cứ biểu hiện trong quá trình học tập,
// rèn luyện; phẩm chất, năng lực chung hình thành qua mọi môn và hoạt động; năng lực đặc thù chủ yếu qua môn tương ứng).
// Mỗi mục có các "căn cứ" (+1 tốt / 0 bình thường / -1 cần cố gắng) lấy từ dữ liệu của tháng:
//   Tiếng Việt, Toán (mức Kiến thức, Kỹ năng và từng yêu cầu), HĐTN, Đạo đức, nộp vở, chuyên cần (vắng không phép), phong trào.
// Mức: Tốt = mọi căn cứ đều tốt (ít nhất 2 căn cứ); Cần cố gắng = căn cứ cần cố gắng nhiều hơn căn cứ tốt; còn lại Đạt.
// Khoa học, Công nghệ, Tin học, Thẩm mĩ, Thể chất (môn do giáo viên khác dạy / lớp 2 chưa có môn riêng): mặc định Đạt.
var NLPC_CHUNG = ['Tự chủ và tự học', 'Giao tiếp và hợp tác', 'Giải quyết vấn đề và sáng tạo'];
var NLPC_DAC_THU = ['Ngôn ngữ', 'Tính toán', 'Khoa học', 'Công nghệ', 'Tin học', 'Thẩm mĩ', 'Thể chất'];
var NLPC_PC = ['Yêu nước', 'Nhân ái', 'Chăm chỉ', 'Trung thực', 'Trách nhiệm'];
var NLPC_MAC_DINH = { 'Khoa học': 1, 'Công nghệ': 1, 'Tin học': 1, 'Thẩm mĩ': 1, 'Thể chất': 1 };
var MON_HDTN = 'Hoạt động trải nghiệm', MON_DD = 'Đạo đức';

function diemMucNL(m) { return m === 'HTXS' || m === 'HTT' ? 1 : m === 'HT' ? 0 : m === 'CHT' || m === 'CCG' ? -1 : null; }
function tenMucNgan(m) { return { HTXS: 'HTXS', HTT: 'HTT', HT: 'HT', CHT: 'CHT', CCG: 'CHT' }[m] || m; }

// Mức của yêu cầu đầu tiên có tên khớp mẫu (VD "Kể chuyện, nói và nghe") trong lần ghi của tháng
function mucYeuCauNL(t, ma, mon, mau) {
  var g = ((t.ghi || {})[ma] || {})[mon]; if (!g) return '';
  var kq = '';
  ['kt', 'kn'].forEach(function (l) { Object.keys(g[l] || {}).forEach(function (ten) { if (!kq && mau.test(ten)) kq = g[l][ten].m; }); });
  return kq;
}

// Kết quả: { muc: {mục: 'Tốt' | 'Đạt' | 'Cần cố gắng' | ''}, canCu: {mục: [{d: +1/0/-1, t: 'chữ'}]} }
function nlpcTuThang(t, ma) {
  var du = (t.du || {})[ma] || {}, o = function (p) { return du[p] || {}; };
  var tv = o('Tiếng Việt'), toan = o('Toán'), hd = o(MON_HDTN).muc || '', dd = o(MON_DD).muc || '';
  var c = function (d, chu, loai) { return d === null || d === undefined ? null : { d: d, t: chu, k: loai || '' }; };
  var mon = function (ten, m, phan) { return m ? c(diemMucNL(m), ten + (phan ? ' ' + phan : '') + ' ' + tenMucNgan(m)) : null; };
  // nộp vở: cộng Tiếng Việt + Toán
  var v = null;
  ['Tiếng Việt', 'Toán'].forEach(function (m) { var x = ((t.vo || {})[ma] || {})[m]; if (!x) return; v = v || { A: 0, B: 0, C: 0, D: 0, E: 0, N: 0, cu: 0 }; Object.keys(x).forEach(function (k) { v[k] = (v[k] || 0) + x[k]; }); });
  var voCL = null, voNop = null;
  if (v) {
    var tot = '', n = 0; ['A', 'B', 'C', 'D', 'E'].forEach(function (k) { if (v[k] > n) { n = v[k]; tot = k; } });
    var tenVo = { A: 'nhanh, đẹp, đúng', B: 'đẹp, đúng', C: 'đúng, biết trình bày', D: 'đúng, chưa biết trình bày', E: 'chưa đúng – chưa đẹp' };
    if (tot) voCL = c(tot === 'A' || tot === 'B' ? 1 : tot === 'C' ? 0 : -1, 'Vở thường ' + tenVo[tot], 'vo');
    var thieu = v.N + v.cu; voNop = c(!thieu ? 1 : thieu <= 2 ? 0 : -1, thieu ? 'Chưa nộp vở ' + thieu + ' lần' : 'Nộp vở đầy đủ', 'nop');
  } else if (t.soNgayVoCu) voNop = c(1, 'Nộp vở đầy đủ', 'nop');
  var kp = (t.vangKP || {})[ma] || 0;
  var cc = t.coDiemDanh ? c(!kp ? 1 : kp <= 2 ? 0 : -1, kp ? 'Vắng không phép ' + kp + ' buổi' : 'Đi học đều', 'cc') : null;
  var pt = (t.pt || {})[ma] || 0, ptc = pt ? c(1, 'Tham gia ' + pt + ' phong trào, cuộc thi') : null;
  var dTB = [tv.mucKT, tv.mucKN, toan.mucKT, toan.mucKN].map(diemMucNL).filter(function (x) { return x !== null; });
  var tb = dTB.length ? dTB.reduce(function (a, b) { return a + b; }, 0) / dTB.length : null;
  var hoc = tb === null ? null : c(tb >= 0.5 ? 1 : tb > -0.5 ? 0 : -1, 'Kết quả học tập chung ' + (tb >= 0.5 ? 'tốt' : tb > -0.5 ? 'đạt' : 'còn hạn chế'));
  var ycTV = function (mau, chu) { var m = mucYeuCauNL(t, ma, 'Tiếng Việt', mau); return m ? c(diemMucNL(m), chu + ' ' + tenMucNgan(m)) : null; };
  var ycToan = function (mau, chu) { var m = mucYeuCauNL(t, ma, 'Toán', mau); return m ? c(diemMucNL(m), chu + ' ' + tenMucNgan(m)) : null; };
  var HD = mon('HĐTN', hd), DD = mon('Đạo đức', dd);
  var canCu = {
    'Tự chủ và tự học': [voCL, voNop, mon('Tiếng Việt', tv.mucKN, 'kỹ năng'), mon('Toán', toan.mucKN, 'kỹ năng'), HD],
    'Giao tiếp và hợp tác': [ycTV(/kể chuyện|nói và nghe|nói theo tình huống/i, 'Kể chuyện, nói và nghe') || mon('Tiếng Việt', tv.mucKN, 'kỹ năng'), HD],
    'Giải quyết vấn đề và sáng tạo': [ycToan(/giải.*toán|bài toán/i, 'Giải toán có lời văn') || mon('Toán', toan.mucKN, 'kỹ năng'), ycTV(/đoạn văn|viết đoạn/i, 'Viết đoạn văn'), HD],
    'Ngôn ngữ': [mon('Tiếng Việt', tv.mucKT, 'kiến thức'), mon('Tiếng Việt', tv.mucKN, 'kỹ năng')],
    'Tính toán': [mon('Toán', toan.mucKT, 'kiến thức'), mon('Toán', toan.mucKN, 'kỹ năng')],
    'Yêu nước': [DD, HD, ptc],
    'Nhân ái': [DD, HD],
    'Chăm chỉ': [voNop, cc, voCL, hoc],
    'Trung thực': [DD],
    'Trách nhiệm': [voNop, cc, HD, ptc]
  };
  var coDuLieu = false, kq = { muc: {}, canCu: {} }, soNguon = {};
  Object.keys(canCu).forEach(function (k) { soNguon[k] = canCu[k].length; canCu[k] = canCu[k].filter(function (x) { return x; }); if (canCu[k].length) coDuLieu = true; });
  NLPC_CHUNG.concat(NLPC_DAC_THU, NLPC_PC).forEach(function (k) {
    if (NLPC_MAC_DINH[k]) { kq.muc[k] = coDuLieu ? 'Đạt' : ''; kq.canCu[k] = [{ d: 0, t: 'Môn do giáo viên khác dạy / lớp 2 chưa có môn riêng – mặc định Đạt, cô sửa nếu cần' }]; return; }
    var ds = canCu[k], tot = 0, yeu = 0;
    ds.forEach(function (x) { if (x.d > 0) tot++; else if (x.d < 0) yeu++; });
    kq.canCu[k] = ds;
    // Tốt cần "biểu hiện rõ và thường xuyên": mọi căn cứ đều tốt và có ít nhất 2 căn cứ (mục chỉ có 1 nguồn như Trung thực thì 1 là đủ)
    kq.muc[k] = !ds.length ? '' : (tot === ds.length && (ds.length >= 2 || soNguon[k] === 1) ? 'Tốt' : yeu > tot ? 'Cần cố gắng' : 'Đạt');
  });
  return kq;
}

// Câu nhận xét theo mục × mức (2 câu mỗi mức để các em không trùng nhau hoàn toàn)
var CAU_NLPC = {
  'Tự chủ và tự học': { 'Tốt': ['Tự giác học tập, tự hoàn thành bài trong vở.', 'Có ý thức tự học, tự làm bài đầy đủ.'], 'Đạt': ['Biết tự học, tự làm bài khi được nhắc.', 'Bước đầu biết tự học, tự làm bài.'], 'Cần cố gắng': ['Cần tự giác hơn trong học tập, làm bài.', 'Cần tự làm bài đầy đủ hơn.'] },
  'Giao tiếp và hợp tác': { 'Tốt': ['Mạnh dạn giao tiếp, hợp tác tốt với bạn.', 'Nói rõ ràng, biết lắng nghe và hợp tác với bạn.'], 'Đạt': ['Biết giao tiếp, hợp tác với bạn.', 'Biết trao đổi, làm việc cùng bạn.'], 'Cần cố gắng': ['Cần mạnh dạn hơn khi giao tiếp, hợp tác.', 'Cần tích cực trao đổi, hợp tác với bạn.'] },
  'Giải quyết vấn đề và sáng tạo': { 'Tốt': ['Biết giải quyết vấn đề, có ý tưởng sáng tạo.', 'Tự tìm cách giải bài, vận dụng linh hoạt.'], 'Đạt': ['Bước đầu biết giải quyết vấn đề đơn giản.', 'Giải quyết được vấn đề khi có gợi ý.'], 'Cần cố gắng': ['Cần tích cực suy nghĩ để giải quyết vấn đề.', 'Cần mạnh dạn tìm cách giải bài.'] },
  'Ngôn ngữ': { 'Tốt': ['Năng lực ngôn ngữ tốt: đọc, viết, nói rõ ràng.', 'Đọc, viết, diễn đạt tốt.'], 'Đạt': ['Năng lực ngôn ngữ đạt yêu cầu.', 'Đọc, viết đạt yêu cầu.'], 'Cần cố gắng': ['Cần rèn thêm kĩ năng đọc, viết.', 'Cần luyện đọc, viết nhiều hơn.'] },
  'Tính toán': { 'Tốt': ['Năng lực tính toán tốt: tính nhanh, chính xác.', 'Tính toán nhanh, cẩn thận.'], 'Đạt': ['Năng lực tính toán đạt yêu cầu.', 'Tính toán đạt yêu cầu.'], 'Cần cố gắng': ['Cần rèn thêm kĩ năng tính toán.', 'Cần luyện tính toán cẩn thận hơn.'] },
  'Yêu nước': { 'Tốt': ['Yêu quê hương, tích cực tham gia phong trào.', 'Yêu trường lớp, quê hương, nhiệt tình tham gia hoạt động.'], 'Đạt': ['Biết yêu quê hương, trường lớp.', 'Tham gia các hoạt động của lớp.'], 'Cần cố gắng': ['Cần tích cực hơn trong các phong trào.', 'Cần tham gia hoạt động chung nhiều hơn.'] },
  'Nhân ái': { 'Tốt': ['Yêu thương, quan tâm, giúp đỡ bạn bè.', 'Biết chia sẻ, giúp đỡ mọi người.'], 'Đạt': ['Biết quan tâm đến bạn bè.', 'Hòa nhã với bạn bè.'], 'Cần cố gắng': ['Cần quan tâm, giúp đỡ bạn nhiều hơn.', 'Cần hòa nhã, chia sẻ với bạn.'] },
  'Chăm chỉ': { 'Tốt': ['Chăm chỉ, đi học đều, làm bài đầy đủ.', 'Chăm học, chuẩn bị bài chu đáo.'], 'Đạt': ['Đi học đều, làm bài tương đối đầy đủ.', 'Có ý thức học tập.'], 'Cần cố gắng': ['Cần chăm chỉ hơn, làm bài đầy đủ.', 'Cần đi học đều và làm bài đầy đủ hơn.'] },
  'Trung thực': { 'Tốt': ['Thật thà, mạnh dạn nhận lỗi.', 'Trung thực trong học tập và sinh hoạt.'], 'Đạt': ['Thật thà trong học tập.', 'Biết nói thật.'], 'Cần cố gắng': ['Cần thật thà hơn khi nhận lỗi.', 'Cần trung thực hơn trong học tập.'] },
  'Trách nhiệm': { 'Tốt': ['Có trách nhiệm, nộp vở đúng hạn, giữ gìn đồ dùng.', 'Hoàn thành tốt việc được giao.'], 'Đạt': ['Biết giữ gìn đồ dùng, làm việc được giao.', 'Làm được việc được giao.'], 'Cần cố gắng': ['Cần có trách nhiệm hơn với việc được giao.', 'Cần cố gắng hoàn thành việc được giao.'] }
};
// Mục "Cần cố gắng" vì vở / chuyên cần thì nói đúng điều đó
var CAU_YEU_NL = { nop: 'Cần nộp vở đầy đủ, đúng hạn.', cc: 'Cần đi học đều, đúng giờ.', vo: 'Cần làm bài trong vở cẩn thận, đúng hơn.' };
function chonCauNL(ds, ma, k) { if (!ds || !ds.length) return ''; var n = parseInt(String(ma).slice(-4), 10) || 0; return ds[(n + k) % ds.length]; }
// Nhận xét 1 ô (phan: 'NLC' | 'NLD' | 'PC'): điều tốt trước, điều cần cố gắng sau; mục chỉ Đạt thì nêu khi không có gì nổi bật. Không quá n ký tự.
function cauNLPC(kq, phan, ma, n) {
  var ds = phan === 'NLC' ? NLPC_CHUNG : phan === 'PC' ? NLPC_PC : ['Ngôn ngữ', 'Tính toán'];
  var lay = function (muc) {
    return ds.filter(function (k) { return kq.muc[k] === muc; }).map(function (k) {
      if (muc === 'Cần cố gắng' && phan === 'PC') {   // phẩm chất yếu vì vở / chuyên cần → câu nói đúng điều đó
        var y = (kq.canCu[k] || []).filter(function (x) { return x.d < 0 && CAU_YEU_NL[x.k]; })[0];
        if (y) return CAU_YEU_NL[y.k];
      }
      return chonCauNL((CAU_NLPC[k] || {})[muc], ma, ds.indexOf(k));
    });
  };
  var tot = lay('Tốt'), can = lay('Cần cố gắng'), dat = lay('Đạt');
  var cau = tot.concat(can); if (!cau.length) cau = dat.slice(0, 2); else if (cau.length < 2) cau = cau.concat(dat.slice(0, 1));
  var out = '', da = {};
  cau.forEach(function (x) { if (!x || da[x]) return; da[x] = 1; var th = out ? out + ' ' + x : x; if (th.length <= n) out = th; });
  return out;
}


// ============================================================================ NHẬN XÉT TỪ BÀI KIỂM TRA (điểm từng câu × đề)
// de: [{cau, toiDa, muc, kiNang}] · cau: {câu: điểm em đạt}. Trọn điểm → làm tốt; dưới nửa số điểm → cần rèn thêm; còn lại → còn sai sót nhỏ.
function soVN(x) { return String(x).replace('.', ','); }
function nhanXetBKT(de, cau, diem, tenBai, n) {
  if (!de || !de.length || !cau) return '';
  var tot = [], vua = [], yeu = [], m3 = { tot: 0, yeu: 0, co: 0 };
  de.forEach(function (c) {
    var td = Number(String(c.toiDa).replace(',', '.')), d = cau[c.cau];
    if (!(td > 0) || d === undefined || d === null || d === '' || !c.kiNang) return;
    d = Number(String(d).replace(',', '.')); if (isNaN(d)) return;
    var kn = String(c.kiNang).trim().replace(/\.$/, ''); kn = kn.charAt(0).toLowerCase() + kn.slice(1);
    var r = d / td;
    if (r >= 0.999) tot.push(kn); else if (r < 0.5) yeu.push(kn); else vua.push(kn);
    if (String(c.muc) === '3') { m3.co++; if (r >= 0.999) m3.tot++; else if (r < 0.5) m3.yeu++; }
  });
  if (!tot.length && !vua.length && !yeu.length) return '';
  var bo = function (ds) { var o = []; ds.forEach(function (x) { if (o.indexOf(x) < 0) o.push(x); }); return o; };
  tot = bo(tot); vua = bo(vua); yeu = bo(yeu);
  var cau3 = [];
  if (diem) cau3.push('Bài kiểm tra ' + tenBai + ' đạt ' + diem + ' điểm.');
  if (tot.length) cau3.push('Làm tốt: ' + tot.slice(0, 4).join('; ') + '.');
  if (vua.length) cau3.push('Còn sai sót nhỏ ở: ' + vua.slice(0, 2).join('; ') + '.');
  if (yeu.length) cau3.push('Cần rèn thêm: ' + yeu.slice(0, 3).join('; ') + '.');
  if (m3.co && m3.tot === m3.co && yeu.length + vua.length) cau3.push('Làm được cả câu vận dụng (mức 3).');
  else if (m3.co && m3.yeu === m3.co && tot.length) cau3.push('Câu vận dụng (mức 3) còn lúng túng.');
  var out = '';
  cau3.forEach(function (x) { var th = out ? out + ' ' + x : x; if (th.length <= (n || 300)) out = th; });
  return out;
}
