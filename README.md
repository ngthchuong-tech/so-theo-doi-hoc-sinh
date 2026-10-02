# Sổ theo dõi học sinh – app hỗ trợ giáo viên chủ nhiệm tiểu học

Web app trên **Google Apps Script + Google Trang tính + Google Drive** cho 1 giáo viên chủ nhiệm lớp tiểu học
(lớp 2, bộ sách Kết nối tri thức). Dữ liệu học sinh nằm trong tài khoản Google của giáo viên; app chỉ mở được bằng chính tài khoản đó.

## Chức năng chính
- Trang chủ: ngày, tuần học, lịch dạy hôm nay (từ Sổ báo giảng), thẻ nhắc việc.
- Điểm danh hằng ngày; tổng hợp theo tuần / tháng / học kì / năm, xuất Excel.
- Sổ theo dõi: ghi nhận xét theo bài học bằng thư viện câu mẫu (4 mức), phẩm chất – năng lực.
- Đánh giá: nhận xét tháng, định kỳ GK1 / CK1 / GK2 / CK2 (môn học T/H/C, điểm KTĐK, 15 mục năng lực – phẩm chất),
  nhận xét GVCN cuối năm; điền thẳng vào file mẫu tải từ hệ thống của Bộ GD&ĐT.
- Tuỳ chọn AI (Gemini, gói trả phí): đọc điểm từ ảnh bài kiểm tra; "Viết lại bằng AI" nhận xét cuối kì / GVCN (xem trước, hoàn tác).
- Minh chứng: chụp tài liệu nhiều trang, gộp PDF; dung lượng ảnh, xoá ảnh năm cũ; chuyển năm học mới (sao lưu năm cũ).

## Cấu trúc thư mục
| Thư mục / file | Nội dung |
|---|---|
| `src/` | Mã nguồn: `Code.gs` (máy chủ), `Index.html`, `Client.html`, `Styles.html`, `DanhGia.html`, `HuongDan.html`, `parsers.js` |
| `build.py` | Gộp mã nguồn → `Ma nguon Apps Script/` (2 file để dán vào Apps Script) và `xem-thu/` (chạy thử trên máy) |
| `Ma nguon Apps Script/` | Bản để cài: `Code.gs`, `Index.html`, `appsscript.json` |
| `dev/mock_gas.js` | Giả lập Google (Trang tính, Drive, Gemini…) để chạy thử trên máy, dữ liệu lưu trong trình duyệt |
| `Cập nhật …txt`, `*.docx` | Ghi chú từng lần cập nhật, hướng dẫn cài đặt, phiếu thử |

## Sửa và cài đặt
1. Sửa trong `src/`, rồi chạy `python build.py`.
2. Chạy thử trên máy: mở `xem-thu/ban-phat-hanh.html` qua một máy chủ web tĩnh (VD `python -m http.server 8765 --directory xem-thu`);
   thêm `?ngay=2026-11-10` vào đường link để giả lập ngày.
3. Cài lên Google: mở Apps Script gắn với file Trang tính dữ liệu → dán `Ma nguon Apps Script/Code.gs` và `Index.html`
   → Lưu → Triển khai → Quản lý các bản triển khai → Phiên bản mới (quyền truy cập: "Chỉ mình tôi").

## Dữ liệu và bảo mật
- Kho mã này **không chứa dữ liệu học sinh** (`.gitignore` chặn file Excel, ảnh, PDF). Không chép dữ liệu lớp vào đây.
- Khoá AI lưu trong Script Properties của Apps Script, không nằm trong mã hay Trang tính.
- Không gửi họ tên, ngày sinh, mã định danh học sinh cho AI.
