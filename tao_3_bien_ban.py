import zipfile, re, os, openpyxl

# ── 1. Đọc dữ liệu Excel ──────────────────────────────────────────────
wb = openpyxl.load_workbook('docs/thong_tin_giam_dinh_xe_filled.xlsx')
ws_info = wb['Thông tin']
ws_pt   = wb['Phụ tùng']

info = {}
for row in ws_info.iter_rows(values_only=True):
    if row[0] and str(row[0]).startswith('{') and row[2] is not None:
        key = str(row[0]).strip('{}')
        info[key] = str(row[2])

gdv_map = {
    'CHINH05':  {'ten': 'Nguyễn Hồng Chinh',  'sdt': ''},
    'TUYENLM':  {'ten': 'Lương Minh Tuyến',    'sdt': ''},
    'DUYNT':    {'ten': 'Nguyễn Thế Duy',      'sdt': ''},
    'SONTT':    {'ten': 'Trần Thanh Sơn',       'sdt': ''},
    'VIETNT05': {'ten': 'Nguyễn Tiến Việt',    'sdt': ''},
    'HUONGNV':  {'ten': 'Nguyễn Văn Hướng',    'sdt': ''},
    'TUNGHX':   {'ten': 'Hoàng Xuân Tùng',     'sdt': ''},
}

ma_gdv = info.get('ma_giam_dinh_vien', '')
if ma_gdv in gdv_map:
    g = gdv_map[ma_gdv]
    info['ma_giam_dinh_vien'] = g['ten']
    info['SĐT'] = g['sdt']
else:
    info['SĐT'] = ''

# Đọc phụ tùng
phu_tung_all = []
for row in ws_pt.iter_rows(min_row=2, values_only=True):
    if row[0] and 'Phương án hợp lệ' not in str(row[0]):
        phu_tung_all.append({'ten': str(row[0]), 'phuong_an': str(row[1]) if row[1] else '', 'sl': row[2] or 1})

# Tất cả phụ tùng (cho các biên bản cần danh sách đầy đủ)
for i in range(1, 15):
    info[f'ten_phu_tung_{i}'] = phu_tung_all[i-1]['ten'] if i-1 < len(phu_tung_all) else ''

# Phụ tùng THU HỒI (cho Vật tư thu hồi)
pt_thu_hoi = [p for p in phu_tung_all if 'thu hồi' in p['phuong_an'].lower() and 'không' not in p['phuong_an'].lower()]
for i in range(1, 14):
    info[f'ten_pt_thu_hoi_{i}'] = pt_thu_hoi[i-1]['ten'] if i-1 < len(pt_thu_hoi) else ''

# Alias cho Biên bản nghiệm thu
info['ten_lai_xe'] = info.get('lai_xe', '')
info['ten_gara_sua_chua'] = info.get('ten_gara', '')
info['Dien_bien_tai_nan'] = info.get('dien_bien_tai_nan', '')


# ── 2. Chuyển số tiền sang chữ ────────────────────────────────────────
def so_thanh_chu(so_str):
    """Chuyển số tiền VND sang chữ tiếng Việt."""
    so_str = so_str.replace(',', '').replace('.', '').strip()
    try:
        n = int(so_str)
    except ValueError:
        return so_str

    don_vi = ['', 'nghìn', 'triệu', 'tỷ']
    chu_so = ['không', 'một', 'hai', 'ba', 'bốn', 'năm', 'sáu', 'bảy', 'tám', 'chín']
    chu_chuc = ['', 'mười', 'hai mươi', 'ba mươi', 'bốn mươi', 'năm mươi',
                'sáu mươi', 'bảy mươi', 'tám mươi', 'chín mươi']

    def ba_chu_so(x):
        tram = x // 100
        chuc = (x % 100) // 10
        dv = x % 10
        parts = []
        if tram > 0:
            parts.append(chu_so[tram] + ' trăm')
        if chuc > 1:
            s = chu_chuc[chuc]
            if dv > 0:
                s += ' ' + chu_so[dv]
            parts.append(s)
        elif chuc == 1:
            s = 'mười'
            if dv > 0:
                s += ' ' + chu_so[dv]
            parts.append(s)
        elif dv > 0:
            if tram > 0 or chuc > 0:
                parts.append('lẻ ' + chu_so[dv])
            else:
                parts.append(chu_so[dv])
        return ' '.join(parts)

    if n == 0:
        return 'không đồng'

    groups = []
    temp = n
    while temp > 0:
        groups.append(temp % 1000)
        temp //= 1000

    parts = []
    for i in range(len(groups) - 1, -1, -1):
        g = groups[i]
        if g > 0:
            s = ba_chu_so(g)
            if don_vi[i]:
                s += ' ' + don_vi[i]
            parts.append(s)

    result = ' '.join(parts)
    # Viết hoa chữ đầu
    result = result[0].upper() + result[1:] + ' đồng'
    return result

tien_so = info.get('tien_tt', '0')
info['tien_tt_chu'] = so_thanh_chu(tien_so)


# ── 3. Hàm fix split-placeholder trong XML ───────────────────────────
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

            prefix = texts[si][:offset_in_si]
            suffix = texts[ei][offset_in_ei_end:]

            texts[si] = prefix + ph.group(0)
            texts[ei] = suffix
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


def make_docx(src, dst, data):
    os.makedirs('output', exist_ok=True)
    with zipfile.ZipFile(src, 'r') as zin:
        with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data_bytes = zin.read(item.filename)
                if item.filename.startswith('word/') and item.filename.endswith('.xml'):
                    xml = data_bytes.decode('utf-8')
                    xml = merge_split_placeholders(xml)
                    xml = apply_replacements(xml, data)
                    data_bytes = xml.encode('utf-8')
                zout.writestr(item, data_bytes)

    # Kiểm tra placeholder còn sót
    with zipfile.ZipFile(dst, 'r') as z:
        xml_check = z.read('word/document.xml').decode('utf-8')
    remaining = re.findall(r'\{[a-zA-ZÀ-ỹ_][^{}]{1,50}\}', xml_check)
    remaining = [p for p in remaining if not p.startswith('{28A0')]
    print(f'Tạo xong: {dst}')
    if remaining:
        print(f'  ⚠ Placeholder còn sót: {sorted(set(remaining))}')
    else:
        print(f'  ✓ Tất cả placeholder đã được thay thế!')


# ── 4. Tạo 3 file output ──────────────────────────────────────────────
make_docx(
    'assets/Xác-nhận-bồi-thường.docx',
    'output/Xac-nhan-boi-thuong.docx',
    info
)

make_docx(
    'assets/Vật-tư-thu-hồi.docx',
    'output/Vat-tu-thu-hoi.docx',
    info
)

make_docx(
    'assets/Biên-bản-nghiệm-thu.docx',
    'output/Bien-ban-nghiem-thu.docx',
    info
)

print()
print(f'Tiền bồi thường: {tien_so} → {info["tien_tt_chu"]}')
if pt_thu_hoi:
    print(f'Phụ tùng thu hồi: {[p["ten"] for p in pt_thu_hoi]}')
else:
    print('Không có phụ tùng thu hồi (Vật tư thu hồi để trống)')
