"""Đóng gói mã nguồn thành các file để dán vào Apps Script (thư mục 'Ma nguon Apps Script'),
và tạo trang xem thử chạy được trên máy (thư mục 'xem-thu', dùng dữ liệu giả lập, không cần Google).

Chạy: python build.py
"""
import os, re, shutil

GOC = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(GOC, 'src')
DIST = os.path.join(GOC, 'Ma nguon Apps Script')
XEM = os.path.join(GOC, 'xem-thu')

def doc(p):
    return open(os.path.join(SRC, p), encoding='utf8').read()

def ghi(thu_muc, ten, nd):
    os.makedirs(thu_muc, exist_ok=True)
    open(os.path.join(thu_muc, ten), 'w', encoding='utf8', newline='\n').write(nd)

# 1) Bộ file cho Apps Script
if os.path.isdir(DIST):
    shutil.rmtree(DIST)
# Gộp toàn bộ giao diện (định dạng + đọc file + xử lý) vào MỘT file Index.html,
# để khi cài/cập nhật chỉ phải dán 2 file: Code.gs và Index.html (dán nhiều file dễ nhầm).
mot_file = doc('Index.html')
mot_file = mot_file.replace("<?!= include('Styles'); ?>", doc('Styles.html'))
mot_file = mot_file.replace("<?!= include('HuongDan'); ?>", doc('HuongDan.html'))
mot_file = mot_file.replace("<?!= include('DanhGia'); ?>", doc('DanhGia.html'))
mot_file = mot_file.replace("<?!= include('Parsers'); ?>", '<script>\n' + doc('parsers.js') + '\n</script>')
mot_file = mot_file.replace("<?!= include('Client'); ?>", doc('Client.html'))
assert '<?' not in mot_file, 'Index.html sau khi gộp vẫn còn lệnh include'
ghi(DIST, 'Code.gs', doc('Code.gs'))
ghi(DIST, 'Index.html', mot_file)
ghi(DIST, 'appsscript.json', doc('appsscript.json'))

# 2) Trang xem thử trên máy: thay các lệnh include bằng nội dung, chèn bộ giả lập Google
html = doc('Index.html')
html = html.replace("<?!= include('Styles'); ?>", doc('Styles.html'))
html = html.replace("<?!= include('HuongDan'); ?>", doc('HuongDan.html'))
html = html.replace("<?!= include('DanhGia'); ?>", doc('DanhGia.html'))
gia_lap = '<script src="mock_gas.js"></script>\n<script src="Code.gs.js"></script>\n<script>\n' + doc('parsers.js') + '\n</script>\n'
html = html.replace("<?!= include('Parsers'); ?>", gia_lap)
html = html.replace("<?!= include('Client'); ?>", doc('Client.html'))
ghi(XEM, 'index.html', html)
ghi(XEM, 'Code.gs.js', doc('Code.gs'))
shutil.copy(os.path.join(GOC, 'dev', 'mock_gas.js'), os.path.join(XEM, 'mock_gas.js'))
# 3) Chạy thử CHÍNH file Index.html sẽ dán lên Google (chỉ chèn thêm bộ giả lập vào đầu trang)
ghi(XEM, 'ban-phat-hanh.html', mot_file.replace('<head>', '<head>\n<script src="mock_gas.js"></script>\n<script src="Code.gs.js"></script>', 1))
print('Đã tạo:', DIST, 'và', XEM)
