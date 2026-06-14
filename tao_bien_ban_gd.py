import zipfile, re, os, openpyxl

# ── 1. Đọc dữ liệu Excel ──────────────────────────────────────────────
wb = openpyxl.load_workbook('input/thong_tin_giam_dinh_xe.xlsx')
ws_info = wb['Thông tin']
ws_pt   = wb['Phụ tùng']
ws_gdv  = wb['GĐV']

info = {}
for row in ws_info.iter_rows(values_only=True):
    if row[0] and str(row[0]).startswith('{') and row[2] is not None:
        key = str(row[0]).strip('{}')
        info[key] = str(row[2])

gdv_map = {}
for row in ws_gdv.iter_rows(min_row=2, values_only=True):
    if row[0] and row[1]:
        gdv_map[str(row[0])] = {'ten': str(row[1]), 'sdt': str(row[2]) if row[2] else ''}

ma_gdv = info.get('ma_giam_dinh_vien', '')
if ma_gdv in gdv_map:
    g = gdv_map[ma_gdv]
    info['ma_giam_dinh_vien'] = g['ten']
    info['SĐT'] = g['sdt']
else:
    info['SĐT'] = ''

phu_tung = []
for row in ws_pt.iter_rows(min_row=2, values_only=True):
    if row[0] and 'Phương án hợp lệ' not in str(row[0]):
        phu_tung.append(str(row[0]))

for i in range(1, 14):
    info[f'ten_phu_tung_{i}'] = phu_tung[i-1] if i-1 < len(phu_tung) else ''

# Alias: template dùng {Dien_bien_tai_nan} (chữ hoa D)
info['Dien_bien_tai_nan'] = info.get('dien_bien_tai_nan', '')


# ── 2. Hàm fix split-placeholder trong XML ───────────────────────────
def merge_split_placeholders(xml):
    """
    Gộp các <w:r> liên tiếp trong cùng <w:p> mà khi nối text lại
    tạo thành {placeholder} hoàn chỉnh.
    """
    def process_para(para):
        # Chỉ match <w:t> hoặc <w:t ...>, KHÔNG match <w:tab> hay <w:tabs>
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

        texts = list(texts)  # mutable copy
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


# ── 3. Tạo file output ───────────────────────────────────────────────
src = 'assets/Biên-bản-giám-định.docx'
os.makedirs('output', exist_ok=True)
dst = 'output/Bien-ban-giam-dinh.docx'

with zipfile.ZipFile(src, 'r') as zin:
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data_bytes = zin.read(item.filename)
            if item.filename in ('word/document.xml', 'word/header1.xml',
                                  'word/footer1.xml', 'word/footer2.xml'):
                xml = data_bytes.decode('utf-8')
                xml = merge_split_placeholders(xml)
                xml = apply_replacements(xml, info)
                data_bytes = xml.encode('utf-8')
            zout.writestr(item, data_bytes)

print('Tạo xong:', dst)

# Kiểm tra placeholder còn sót
with zipfile.ZipFile(dst, 'r') as z:
    xml_check = z.read('word/document.xml').decode('utf-8')
remaining = re.findall(r'\{[a-zA-ZÀ-ỹ_][^{}]{1,50}\}', xml_check)
remaining = [p for p in remaining if not p.startswith('{28A0')]
if remaining:
    print('Placeholder còn sót:', sorted(set(remaining)))
else:
    print('Tất cả placeholder đã được thay thế!')
