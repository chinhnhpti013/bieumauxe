# -*- coding: utf-8 -*-
import zipfile, re, os, openpyxl

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

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
    os.makedirs(os.path.dirname(dst), exist_ok=True)
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
    with zipfile.ZipFile(dst, 'r') as z:
        xml_check = z.read('word/document.xml').decode('utf-8')
    remaining = re.findall(r'\{[a-zA-ZÀ-ỹ_][^{}]{1,50}\}', xml_check)
    remaining = [p for p in remaining if not p.startswith('{28A0')]
    print(f'Tạo xong: {dst}')
    if remaining:
        print(f'  ⚠ Placeholder còn sót: {sorted(set(remaining))}')

# ── Đọc dữ liệu Excel ────────────────────────────────────────────────
wb = openpyxl.load_workbook(os.path.join(BASE_DIR, 'input', 'thong_tin_giam_dinh_xe.xlsx'))
ws_info = wb['Thông tin']

info = {}
for row in ws_info.iter_rows(values_only=True):
    if row[0] and str(row[0]).startswith('{') and row[2] is not None:
        key = str(row[0]).strip('{}')
        info[key] = str(row[2]) if str(row[2]) != 'None' else ''

gdv_map = {
    'CHINH05':  {'ten': 'Nguyễn Hồng Chinh',  'sdt': '0903 210 598'},
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

info['Dien_bien_tai_nan'] = info.get('dien_bien_tai_nan', '')

# ── Sinh file ─────────────────────────────────────────────────────────
make_docx(
    os.path.join(BASE_DIR, 'assets', 'TBTN-YCBT.docx'),
    os.path.join(BASE_DIR, 'output', 'TBTN-YCBT.docx'),
    info
)
