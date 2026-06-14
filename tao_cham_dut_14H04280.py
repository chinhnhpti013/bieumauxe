"""
Chấm dứt - Khôi phục BH xe 14H-042.80
Nguồn: GCN BH số 013OTTN+250004310 + Phiếu xác minh phí
Hồ sơ: Công ty TNHH Phúc Xuyên — KIA Carnival Luxury 2.2D 8 chỗ
"""

import zipfile, re, os
from datetime import datetime

# ════════════════════════════════════════════════════════════════════════
# DỮ LIỆU — trích từ GCN BH (5.pdf) và Phiếu xác minh phí (6.pdf)
# ════════════════════════════════════════════════════════════════════════

info = {
    # ── Thông tin xe (GCN BH trang 1) ──────────────────────────────
    'bien_so_xe':   '14H-042.80',
    'so_khung':     'PH601440',
    'so_may':       'PC321432',
    'hang_xe':      'KIA',
    'dong_xe':      'CARNIVAL LUXURY 2.2D 8 CHO',
    'so_cho_ngoi':  '8',
    'nam_sx':       '2023',

    # ── Chủ xe (GCN BH trang 1) ─────────────────────────────────────
    'chu_xe':           'CÔNG TY TNHH PHÚC XUYÊN',
    'dia_chi_chu_xe':   'Tổ 7, Khu I, Phường Yên Thanh, Uông Bí, Quảng Ninh',
    'dien_thoai_chu_xe':'0203.663366',

    # ── Hợp đồng bảo hiểm (GCN BH + Phiếu XMP) ─────────────────────
    'so_gcn_bh':       '013OTTN+250004310',
    'so_hop_dong':     '0000277/HD/013-013/XO/2025',
    'gcn_bh_tu_ngay':  '28/12/2025',
    'gcn_bh_den_ngay': '28/12/2026',
    'gia_tri_xe':      '970,000,000',
    'phi_bh':          '14,841,000',
    'dk_bs':           'BS02, BS05, BS06',

    # ── Ánh xạ riêng mẫu Chấm dứt ──────────────────────────────────
    'phi_bao_hiem_da_VAT': '14,841,000',
    'nam_hien_tai':        str(datetime.now().year),

    # Ngày cấp GCN (Phiếu XMP: "Ngày cấp: 25/12/2025")
    'gcn_bh_ngay_cap':   '25/12/2025',
    'gcn_bh_ ngay_cap':  '25/12/2025',   # alias (khoảng trắng trong XML)

    # Thời hạn BH VCX ô tô (GCN BH trang 2: "09:25 phút, ngày 28/12/2026")
    'ngay_thoi_han_bao_hiem_vcx_oto':      '28/12/2026',
    'gio_phut_thoi_han_bao_hiem_vcx_oto':  '09:25',

    # Chứng từ Doanh thu Thực Thu lần 1
    # (Phiếu XMP mục III: ngày chứng từ 11/02/2026, số tiền 18,551,250)
    'ngay_chung_tu_doanh_thu_thuc_thu_lan_1':  '11',
    'thang_chung_tu_doanh_thu_thuc_thu_lan_1': '02',
    'nam_chung_tu_doanh_thu_thuc_thu_lan_1':   '2026',
    'tien_han_thanh_toan_dong_1':              '18,551,250',

    # Hạn thanh toán đợt 1
    # (Phiếu XMP mục I: Kỳ 1 chậm nhất ngày 24/01/2026)
    'ngay_han_thanh_toan_dong_1':  '24',
    'thang_han_thanh_toan_dong_1': '01',
    'nam_han_thanh_toan_dong_1':   '2026',

    # Alias khớp placeholder trong mẫu docx (tên ngắn hơn)
    'gcn_bh_gio_phut':                      '09:25',
    'ngay_chung_tu_doanh_thu_thuc_thu':     '11',
    'thang_chung_tu_doanh_thu_thuc_thu':    '02',
    'nam_chung_tu_doanh_thu_thuc_thu':      '2026',
    'ngay_han_thanh_toan':                  '24',
    'thang_han_thanh_toan':                 '01',
    'nam_han_thanh_toan':                   '2026',
    'tien_han_thanh_toan':                  '18,551,250',
    'ngay_thang_nam_han_thanh_toan':        'ngày 24 tháng 01 năm 2026',
    # Ngày nộp phí kỳ tiếp theo (kỳ 2: hạn 24/02/2026)
    'ke_tiep_ngay_nop_phi':                 '24/02/2026',
}

# Tổng hợp chuỗi ngày đầy đủ
info['ngay_thang_nam_chung_tu_doanh_thu_thuc_thu_lan_1'] = (
    f"ngày {info['ngay_chung_tu_doanh_thu_thuc_thu_lan_1']} "
    f"tháng {info['thang_chung_tu_doanh_thu_thuc_thu_lan_1']} "
    f"năm {info['nam_chung_tu_doanh_thu_thuc_thu_lan_1']}"
)
info['ngay_thang_nam_han_thanh_toan_dong_1'] = (
    f"ngày {info['ngay_han_thanh_toan_dong_1']} "
    f"tháng {info['thang_han_thanh_toan_dong_1']} "
    f"năm {info['nam_han_thanh_toan_dong_1']}"
)

# ════════════════════════════════════════════════════════════════════════


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
            texts[si] = texts[si][:offset_in_si] + ph.group(0)
            texts[ei] = texts[ei][offset_in_ei_end:]
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


# ── Tạo file output ──────────────────────────────────────────────────
src = 'assets/Cham-dut-khoi-phuc.docx'
os.makedirs('output', exist_ok=True)
dst = 'output/Cham-dut-khoi-phuc_14H04280.docx'

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

print()
print('─── Tóm tắt dữ liệu đã điền ──────────────────────────────')
print(f"  Xe            : {info['bien_so_xe']} — {info['hang_xe']} {info['dong_xe']}")
print(f"  Chủ xe        : {info['chu_xe']}")
print(f"  Số GCN BH     : {info['so_gcn_bh']}")
print(f"  Phí BH        : {info['phi_bao_hiem_da_VAT']} đồng (đã VAT)")
print(f"  Ngày cấp GCN  : {info['gcn_bh_ngay_cap']}")
print(f"  Thời hạn BH   : {info['ngay_thoi_han_bao_hiem_vcx_oto']} lúc {info['gio_phut_thoi_han_bao_hiem_vcx_oto']}")
print(f"  Chứng từ TT1  : {info['ngay_thang_nam_chung_tu_doanh_thu_thuc_thu_lan_1']}")
print(f"  Tiền TT đợt 1 : {info['tien_han_thanh_toan_dong_1']} VND")
print(f"  Hạn TT đợt 1  : {info['ngay_thang_nam_han_thanh_toan_dong_1']}")
