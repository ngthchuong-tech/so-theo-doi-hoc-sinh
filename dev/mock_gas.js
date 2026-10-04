// Giả lập Google Apps Script để chạy thử app trên máy (KHÔNG dùng khi chạy thật).
// Dữ liệu lưu tạm trong localStorage của trình duyệt. Thêm ?ngay=2026-10-02 vào địa chỉ để giả lập "hôm nay".
(function () {
  var q = /[?&]ngay=(\d{4})-(\d{2})-(\d{2})/.exec(location.search);
  var RealDate = Date, fixed = q ? new RealDate(+q[1], +q[2] - 1, +q[3], 9, 0, 0).getTime() : RealDate.now();
  function FakeDate() {
    var a = Array.prototype.slice.call(arguments);
    if (!a.length) return new RealDate(fixed + (RealDate.now() - startReal));
    return new (Function.prototype.bind.apply(RealDate, [null].concat(a)))();
  }
  var startReal = RealDate.now();
  FakeDate.prototype = RealDate.prototype; FakeDate.now = function () { return fixed; }; FakeDate.UTC = RealDate.UTC; FakeDate.parse = RealDate.parse;
  window.Date = FakeDate;

  var KEY = 'mock_gas_sheets_v1', data;
  try { data = JSON.parse(localStorage.getItem(KEY) || '{}'); } catch (e) { data = {}; }
  function save() { try { localStorage.setItem(KEY, JSON.stringify(data)); } catch (e) {} }
  function Sheet(name) { this.name = name; if (!data[name]) data[name] = { rows: [], max: 1000 }; }
  Sheet.prototype = {
    getName: function () { return this.name; },
    _d: function () { return data[this.name]; },
    getLastRow: function () { var r = this._d().rows; for (var i = r.length - 1; i >= 0; i--) if (r[i] && r[i].some(function (v) { return v !== '' && v != null; })) return i + 1; return 0; },
    getMaxRows: function () { return Math.max(this._d().max, this._d().rows.length); },
    getLastColumn: function () { var h = this._d().rows[0] || []; for (var i = h.length - 1; i >= 0; i--) if (h[i] !== '' && h[i] != null) return i + 1; return 0; },
    insertRowsAfter: function (after, n) { this._d().max += n; save(); },
    deleteRows: function (row, n) { ghiDem('ghi', this.name, n); this._d().rows.splice(row - 1, n); save(); },
    setFrozenRows: function () {},
    getRange: function (row, col, nr, nc) { return new Range(this, row, col, nr || 1, nc || 1); }
  };
  function Range(sh, row, col, nr, nc) { this.sh = sh; this.row = row; this.col = col; this.nr = nr; this.nc = nc; }
  // Bộ đếm đọc/ghi Trang tính cho mỗi lần gọi máy chủ (ước lượng tốc độ trên Google thật): window._doTocDo
  var dem = null;
  function ghiDem(loai, sh, so) { if (!dem) return; dem[loai]++; dem[loai + 'Dong'] += so; dem.bang[sh] = (dem.bang[sh] || 0) + 1; }
  Range.prototype = {
    getValues: function () {
      ghiDem('doc', this.sh.name, this.nr);
      var rows = this.sh._d().rows, out = [];
      for (var i = 0; i < this.nr; i++) { var r = rows[this.row - 1 + i] || [], o = []; for (var j = 0; j < this.nc; j++) { var v = r[this.col - 1 + j]; o.push(v == null ? '' : v); } out.push(o); }
      return out;
    },
    getDisplayValues: function () { return this.getValues().map(function (r) { return r.map(function (v) { return String(v); }); }); },
    setValues: function (vals) {
      ghiDem('ghi', this.sh.name, vals.length);
      var rows = this.sh._d().rows;
      for (var i = 0; i < vals.length; i++) { var k = this.row - 1 + i; rows[k] = rows[k] || []; for (var j = 0; j < vals[i].length; j++) rows[k][this.col - 1 + j] = vals[i][j] == null ? '' : String(vals[i][j]); }
      save(); return this;
    },
    clearContent: function () { var rows = this.sh._d().rows; for (var i = 0; i < this.nr; i++) { var r = rows[this.row - 1 + i]; if (r) for (var j = 0; j < this.nc; j++) r[this.col - 1 + j] = ''; } save(); return this; },
    setNumberFormat: function () { return this; }, setFontWeight: function () { return this; }, setBackground: function () { return this; }
  };
  window.SpreadsheetApp = {
    getActive: function () {
      return {
        getSheetByName: function (n) { return data[n] ? new Sheet(n) : null; },
        insertSheet: function (n) { var s = new Sheet(n); save(); return s; },
        deleteSheet: function (s) { delete data[s.name]; save(); },
        getSheets: function () { return Object.keys(data).map(function (n) { return new Sheet(n); }); },
        getId: function () { return 'FILE-DU-LIEU-GIA-LAP'; }
      };
    }
  };
  function p2(n) { return ('0' + n).slice(-2); }
  window.Utilities = {
    formatDate: function (d, tz, f) {
      return f.replace('dd', p2(d.getDate())).replace('MM', p2(d.getMonth() + 1)).replace('yyyy', d.getFullYear())
        .replace('HH', p2(d.getHours())).replace('mm', p2(d.getMinutes())).replace('ss', p2(d.getSeconds()));
    },
    getUuid: function () { return Math.random().toString(16).slice(2) + Math.random().toString(16).slice(2); },
    base64Decode: function (s) { return s; },          // giả lập: "bytes" chính là chuỗi base64
    base64Encode: function (s) { return s; },
    newBlob: function (data, type, ten) { return { ten: ten, type: type, size: data.length, data: data }; }
  };
  // Lịch Google giả lập (nhắc việc): giữ sự kiện trong bộ nhớ, xem bằng window._lich
  window._lich = {};
  window.CalendarApp = { getDefaultCalendar: function () { return {
    createEvent: function (ten, bd, kt, o) { var id = 'ev' + Math.random().toString(16).slice(2, 10), ev = { ten: ten, bd: bd, nhac: [] };
      window._lich[id] = ev;
      return { getId: function () { return id; }, removeAllReminders: function () { ev.nhac = []; }, addPopupReminder: function (p) { ev.nhac.push(p); } }; },
    getEventById: function (id) { return window._lich[id] ? { deleteEvent: function () { delete window._lich[id]; } } : null; }
  }; } };
  // Google Drive giả lập: ghi nhận tên file đã lưu; nội dung giữ trong bộ nhớ (mất khi tải lại trang)
  window._driveFiles = {};
  function thuMucGia(ten) {
    return { getId: function () { return 'TM-' + ten; }, setTrashed: function () { window._thuMucDaXoa = (window._thuMucDaXoa || []).concat([ten]); },
      getUrl: function () { return 'https://drive.google.com/drive/folders/GIA-LAP'; },
      getFoldersByName: function (t) { return { hasNext: function () { return false; } }; },
      createFolder: function (t) { return thuMucGia(t); },
      createFile: function (b) { window._fileDaLuu = (window._fileDaLuu || []).concat([b.ten]); var id = 'GIA-LAP-' + Math.random().toString(36).slice(2, 8); window._driveFiles[id] = b.data;
        return { getId: function () { return id; }, getUrl: function () { return 'https://drive.google.com/file/d/' + id + '/view'; } }; } };
  }
  window.DriveApp = { getFoldersByName: function () { return { hasNext: function () { return false; } }; }, createFolder: thuMucGia,
    getFolderById: function (id) { return thuMucGia(String(id).replace(/^TM-/, '')); },
    getFileById: function (id) { return { setTrashed: function () { window._fileDaXoa = (window._fileDaXoa || []).concat([id]); },
      getBlob: function () { return { getBytes: function () { if (!(id in window._driveFiles)) throw new Error('[giả lập] ảnh không còn trong bộ nhớ (đã tải lại trang)'); return window._driveFiles[id]; } }; },
      makeCopy: function (ten) { window._banSao = (window._banSao || []).concat([ten]); return { getUrl: function () { return 'https://docs.google.com/spreadsheets/d/BAN-SAO-GIA-LAP'; } }; } }; } };
  // CacheService giả lập: bộ nhớ trong trang (mất khi tải lại trang, giống bộ nhớ đệm hết hạn)
  var _cache = {};
  window.CacheService = { getScriptCache: function () { return {
    get: function (k) { return k in _cache ? _cache[k] : null; },
    getAll: function (ks) { var o = {}; ks.forEach(function (k) { if (k in _cache) o[k] = _cache[k]; }); return o; },
    putAll: function (o) { Object.keys(o).forEach(function (k) { if (String(o[k]).length > 100000) throw new Error('Giá trị quá 100KB'); _cache[k] = o[k]; }); },
    put: function (k, v) { _cache[k] = v; }, remove: function (k) { delete _cache[k]; } }; } };
  // Script Properties giả lập (khoá AI) – lưu riêng trong localStorage của máy
  window.PropertiesService = { getScriptProperties: function () {
    var K = 'gia_lap_script_properties', doc = function () { try { return JSON.parse(localStorage.getItem(K) || '{}'); } catch (e) { return {}; } };
    return { getProperty: function (k) { return doc()[k] || null; }, getProperties: function () { return doc(); },
      setProperty: function (k, v) { var o = doc(); o[k] = v; localStorage.setItem(K, JSON.stringify(o)); },
      deleteProperty: function (k) { var o = doc(); delete o[k]; localStorage.setItem(K, JSON.stringify(o)); } };
  } };
  // Gemini giả lập: không gọi ra Internet. Ghi lại yêu cầu (để kiểm tra không gửi tên học sinh) và trả điểm giả.
  // Đặt window.diemGiaLap = 7 (hoặc null) để chọn điểm AI "đọc" được; khoá "SAI…" để thử lỗi khoá.
  window.UrlFetchApp = { fetch: function (url, o) {
    window._yeuCauAI = (window._yeuCauAI || []).concat([{ url: url, payload: o.payload }]);
    var khoa = (o.headers || {})['x-goog-api-key'] || '';
    var tl = function (ma, obj) { return { getResponseCode: function () { return ma; }, getContentText: function () { return JSON.stringify(obj); } }; };
    if (/^SAI/.test(khoa)) return tl(400, { error: { message: 'API key not valid. Please pass a valid API key.' } });
    if (window.aiLoiGiaLap) return tl(429, { error: { message: 'Resource exhausted' } });
    var nhacViet = JSON.parse(o.payload).contents[0].parts.map(function (p) { return p.text || ''; }).join('');
    if (/Viết lại từng nhận xét/.test(nhacViet)) {           // viết lại nhận xét: trả bản "AI" giả cho từng mục
      var muc = JSON.parse(nhacViet.split('Các mục:\n')[1]);
      var kq = muc.map(function (m) { var g = m.goc.replace(/\.$/, '');
        return { k: m.k, ban: 'Em ' + g.charAt(0).toLowerCase() + g.slice(1) + '. Cô mong em tiếp tục cố gắng nhé (bản AI ' + (Math.random() * 100 | 0) + ').' }; });
      return tl(200, { candidates: [{ content: { parts: [{ text: JSON.stringify(kq) }] }, finishReason: 'STOP' }] });
    }
    var hoiDiem = /inline_data/.test(o.payload);
    var d = window.diemGiaLap !== undefined ? window.diemGiaLap : 5 + Math.floor(Math.random() * 6);
    var text = hoiDiem ? JSON.stringify({ diem: d, chacChan: d == null ? 0 : 0.92, ghiChu: '' }) : 'OK';
    return tl(200, { candidates: [{ content: { parts: [{ text: text }] }, finishReason: 'STOP' }] });
  } };
  window.MimeType = { MICROSOFT_EXCEL: 'application/vnd.ms-excel' };
  window.ScriptApp ={ getService: function () { return { getUrl: function () { return 'https://script.google.com/macros/s/VI-DU-DUONG-LINK/exec'; } }; } };
  window.LockService ={ getScriptLock: function () { return { waitLock: function () {}, releaseLock: function () {} }; } };
  function runner(ok, fail) {
    return new Proxy({}, {
      get: function (_, name) {
        if (name === 'withSuccessHandler') return function (f) { return runner(f, fail); };
        if (name === 'withFailureHandler') return function (f) { return runner(ok, f); };
        return function () {
          var args = JSON.parse(JSON.stringify(Array.prototype.slice.call(arguments)));
          setTimeout(function () {
            // Như Google thật: mỗi lần gọi máy chủ là 1 lần chạy mới, biến toàn cục (bộ nhớ tạm) bắt đầu trống
            if ('_NHO' in window) { window._NHO = {}; window._SHEET = {}; window._PB = null; }
            dem = { doc: 0, docDong: 0, ghi: 0, ghiDong: 0, bang: {} }; var t0 = performance.now();
            try { var r = window[name].apply(null, args); }
            catch (e) { dem = null; console.error(e); fail && fail(e); return; }
            var d = dem; dem = null; d.ham = name; d.ms = Math.round(performance.now() - t0);
            (window._doTocDo = window._doTocDo || []).push(d);
            ok && ok(r === undefined ? undefined : JSON.parse(JSON.stringify(r)));
          }, 120);
        };
      }
    });
  }
  // google.script.history: trên Google thật, lịch sử của web app gắn với thanh địa chỉ trình duyệt
  var hist = {
    push: function (state, params, hash) { history.pushState(state, '', location.pathname + location.search + '#' + (hash || '')); },
    replace: function (state, params, hash) { history.replaceState(state, '', location.pathname + location.search + '#' + (hash || '')); },
    setChangeHandler: function (fn) { window.addEventListener('popstate', function (e) { fn({ state: e.state, location: {} }); }); }
  };
  window.google = { script: { run: runner(null, null), history: hist } };
  // Trên Google, app nằm trong khung cách ly: lệnh history.back()/go() từ trong app KHÔNG có tác dụng.
  // Giả lập đúng như vậy để không bị "chạy thử thì được, lên Google thì hỏng". Nút Back của điện thoại: gọi nutBackDienThoai().
  var backThat = history.back.bind(history);
  history.back = function () { console.warn('[giả lập Google] history.back() bị chặn – app phải tự quay lại'); };
  history.go = function () { console.warn('[giả lập Google] history.go() bị chặn'); };
  window.nutBackDienThoai = function () { backThat(); };
  window.xoaDuLieuGiaLap = function () { localStorage.removeItem(KEY); location.reload(); };
  // Lần đầu: tạo bảng như khi chạy caiDatLanDau trên Google
  window.addEventListener('DOMContentLoaded', function () { if (!data.CaiDat && window.caiDatLanDau) caiDatLanDau(); });
})();
