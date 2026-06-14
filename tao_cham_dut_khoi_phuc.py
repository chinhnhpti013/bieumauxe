"""
Tạo mẫu biểu Chấm dứt - Khôi phục bảo hiểm xe cơ giới PTI.

Cách dùng:
  1. Điền các trường NHẬP TAY trong phần "CẤU HÌNH NHẬP TAY" bên dưới
  2. Chạy: python tao_cham_dut_khoi_phuc.py
  3. File xuất: output/Cham-dut-khoi-phuc.docx

Nguồn dữ liệu tự động:
  - bien_so_xe, chu_xe, so_gcn_bh, phi_bh → đọc từ input/thong_tin_giam_dinh_xe.xlsx
  - nam_hien_tai → tự tính từ ngày hệ thống

Nguồn dữ liệu nhập tay (phần bên dưới):
  - gcn_bh_ngay_cap: ngày cấp giấy chứng nhận BH
  - ngay/gio thời hạn BH VCX ô tô: cuối hạn BH (thường = gcn_bh_den_ngay)
  - Chứng từ doanh thu thực thu lần 1: ngày/tháng/năm và số tiền
  - Hạn thanh toán đợt 1: ngày/tháng/năm
"""

import zipfile, re, os, openpyxl
from datetime import datetime

# ════════════════════════════════════════════════════════════════════════
# CẤU HÌNH NHẬP TAY — Chỉnh sửa các giá trị bên dưới trước khi chạy
# ════════════════════════════════════════════════════════════════════════

NHAP_TAY = {
    # Ngày cấp GCN BH (DD/MM/YYYY)
    'gcn_bh_ngay_cap': '05/11/2025',

    # Thời hạn BH VCX ô tô (thường = gcn_bh_den_ngay trong Excel)
    'ngay_thoi_han_bao_hiem_vcx_oto': '05/11/2026',   # DD/MM/YYYY
    'gio_phut_thoi_han_bao_hiem_vcx_oto': '00:00',     # HH:MM

    # Chứng từ doanh thu thực thu lần 1
    'ngay_chung_tu_doanh_thu_thuc_thu_lan_1':  '05',        # ngày
    'thang_chung_tu_doanh_thu_thuc_thu_lan_1': '11',        # tháng
    'nam_chung_tu_doanh_thu_thuc_thu_lan_1':   '2025',      # năm
    'tien_han_thanh_toan_dong_1': '8,220,000',               # số tiền đợt 1 (đồng)

    # Hạn thanh toán đợt 1
    'ngay_han_thanh_toan_dong_1':  '05',        # ngày
    'thang_han_thanh_toan_dong_1': '11',        # tháng
    'nam_han_thanh_toan_dong_1':   '2025',      # năm
}

# ════════════════════════════════════════════════════════════════════════


# ── 1. Đọc dữ liệu chung từ Excel ───────────────────────────────────
wb = openpyxl.load_workbook('input/thong_tin_giam_dinh_xe.xlsx')
ws_info = wb['Thông tin']

info = {}
for row in ws_info.iter_rows(values_only=True):
    if row[0] and str(row[0]).startswith('{') and row[2] is not None:
        key = str(row[0]).strip('{}')
        info[key] = str(row[2])

# Ánh xạ các trường chung vào placeholder riêng của mẫu này
info['phi_bao_hiem_da_VAT'] = info.get('phi_bh', '')
info['nam_hien_tai'] = str(datetime.now().year)

# Tổ hợp ngày/tháng/năm (cho các trường dạng "ngày DD tháng MM năm YYYY")
def tong_hop_ngay(prefix, d):
    ngay  = d.get(f'ngay_{prefix}',  '')
    thang = d.get(f'thang_{prefix}', '')
    nam   = d.get(f'nam_{prefix}',   '')
    return f'ngày {ngay} tháng {thang} năm {nam}' if ngay else ''

# Nạp dữ liệu nhập tay vào info
info.update(NHAP_TAY)

# Tổng hợp ngày đầy đủ
info['ngay_thang_nam_chung_tu_doanh_thu_thuc_thu_lan_1'] = tong_hop_ngay(
    'chung_tu_doanh_thu_thuc_thu_lan_1', NHAP_TAY)
info['ngay_thang_nam_han_thanh_toan_dong_1'] = tong_hop_ngay(
    'han_thanh_toan_dong_1', NHAP_TAY)

# Alias placeholder có khoảng trắng bên trong ({gcn_bh_ ngay_cap})
info['gcn_bh_ ngay_cap'] = NHAP_TAY.get('gcn_bh_ngay_cap', '')


# ── 2. Hàm fix split-placeholder trong XML ───────────────────────────
def merge_split_placeholders(xml):
    def process_para(para):
        t_re = re.compile(r'(<w:t(?=[>\s])[^>]*>)(.*?)(</w:t>)', re.DOTALL)
        matches = list(t_re.finditer(para))
        if not matches:
            return para
        texts = [m.group(2) for m in matches]
        combined = ''.join(texts)
        if '{' not in combined:
            return para
        char_map = []
        for idx, t in enumerate(texts):
            char_map.extend([idx] * len(t))
        placeholders_found = list(re.compile(r'\{[^{}]+\}').finditer(combined))
        if not placeholders_found:
            return para
        texts = list(texts)
        for ph in placeholders_found:
            s, e = ph.start(), ph.end() - 1
            if s >= len(char_map) or e >= len(char_map):
                continue
            si, ei = char_map[s], char_map[e]
            if si == ei:
                continue
            offset_in_si = s - sum(len(texts[k]) for k in range(si))
            offset_in_ei_end = e - sum(len(texts[k]) for k in range(ei)) + 1
            prefix_txt = texts[si][:offset_in_si]
            suffix_txt = texts[ei][offset_in_ei_end:]
            texts[si] = prefix_txt + ph.group(0)
            texts[ei] = suffix_txt
            for k in range(si+1, ei):
                texts[k] = ''
        result = para
        for i in range(len(matches)-1, -1, -1):
            m = matches[i]
            if ' ' in texts[i] and 'xml:space' not in m.group(1):
                new_t = '<w:t xml:space="preserve">' + texts[i] + m.group(3)
            else:
                new_t = m.group(1) + texts[i] + m.group(3)
            result = result[:m.start()] + new_t + result[m.end():]
        return result

    out = []
    pos = 0
    for pm in re.finditer(r'<w:p[ >]', xml):
        p_start = pm.start()
        p_end_idx = xml.find('</w:p>', p_start)
        if p_end_idx == -1:
            break
        p_end = p_end_idx + len('</w:p>')
        out.append(xml[pos:p_start])
        out.append(process_para(xml[p_start:p_end]))
        pos = p_end
    out.append(xml[pos:])
    return ''.join(out)


def apply_replacements(xml, data):
    for key, val in data.items():
        xml = xml.replace('{' + key + '}', val)
    return xml


# ── 3. Tạo file output ───────────────────────────────────────────────
src = 'assets/Cham-dut-khoi-phuc.docx'
os.makedirs('output', exist_ok=True)
dst = 'output/Cham-dut-khoi-phuc.docx'

with zipfile.ZipFile(src, 'r') as zin:
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data_bytes = zin.read(item.filename)
            if item.filename.startswith('word/') and item.filename.endswith('.xml'):
                xml = data_bytes.decode('utf-8')
                xml = merge_split_placeholders(xml)
                xml = apply_replacements(xml, info)
                data_bytes = xml.encode('utf-8')
            zout.writestr(item, data_bytes)

print(f'Tạo xong: {dst}')

# Kiểm tra placeholder còn sót
with zipfile.ZipFile(dst, 'r') as z:
    xml_check = z.read('word/document.xml').decode('utf-8')
remaining = re.findall(r'\{[a-zA-ZÀ-ỹ _][^{}]{1,60}\}', xml_check)
remaining = [p for p in remaining if not p.startswith('{28A0')]
if remaining:
    print('⚠  Placeholder còn sót:', sorted(set(remaining)))
else:
    print('✓  Tất cả placeholder đã được thay thế!')

# In tóm tắt dữ liệu đã điền
print()
print('─── Dữ liệu đã điền ───────────────────────────────────')
print(f"  Biển số xe    : {info.get('bien_so_xe', '')}")
print(f"  Chủ xe        : {info.get('chu_xe', '')}")
print(f"  Số GCN BH     : {info.get('so_gcn_bh', '')}")
print(f"  Phí BH        : {info.get('phi_bao_hiem_da_VAT', '')} đồng")
print(f"  Ngày cấp GCN  : {NHAP_TAY['gcn_bh_ngay_cap']}")
print(f"  Thời hạn BH   : {NHAP_TAY['ngay_thoi_han_bao_hiem_vcx_oto']} {NHAP_TAY['gio_phut_thoi_han_bao_hiem_vcx_oto']}")
print(f"  Chứng từ lần 1: {info['ngay_thang_nam_chung_tu_doanh_thu_thuc_thu_lan_1']}")
print(f"  Tiền đợt 1    : {NHAP_TAY['tien_han_thanh_toan_dong_1']} đồng")
print(f"  Hạn TT đợt 1  : {info['ngay_thang_nam_han_thanh_toan_dong_1']}")
