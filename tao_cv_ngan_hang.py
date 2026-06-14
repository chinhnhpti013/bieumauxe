# -*- coding: utf-8 -*-
import zipfile, re, os, openpyxl

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ASSET  = os.path.join(BASE_DIR, 'assets', 'CV-Ngân hàng.docx')
OUTPUT = os.path.join(BASE_DIR, 'output', 'CV-Ngan-hang.docx')

os.makedirs(os.path.dirname(OUTPUT), exist_ok=True)


def so_sang_chu(so_str):
    try:
        so = int(str(so_str).replace(",", "").replace(".", "").strip())
    except (ValueError, AttributeError):
        return str(so_str)
    don_vi = ["", "nghìn", "triệu", "tỷ"]
    chu_so = ["không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín"]

    def doc_3(n):
        if n == 0:
            return ""
        tram = n // 100
        chuc = (n % 100) // 10
        dv = n % 10
        r = ""
        if tram > 0:
            r += chu_so[tram] + " trăm"
        if chuc == 0:
            if dv > 0 and tram > 0:
                r += " lẻ " + chu_so[dv]
            elif dv > 0:
                r += chu_so[dv]
        elif chuc == 1:
            r += " mười"
            if dv == 5: r += " lăm"
            elif dv > 0: r += " " + chu_so[dv]
        else:
            r += " " + chu_so[chuc] + " mươi"
            if dv == 1: r += " mốt"
            elif dv == 5: r += " lăm"
            elif dv > 0: r += " " + chu_so[dv]
        return r.strip()

    if so == 0:
        return "Không đồng"
    groups = []
    tmp = so
    while tmp > 0:
        groups.append(tmp % 1000)
        tmp //= 1000
    parts = []
    for i in range(len(groups) - 1, -1, -1):
        g = groups[i]
        if g > 0:
            txt = doc_3(g)
            if don_vi[i]:
                txt += " " + don_vi[i]
            parts.append(txt)
    result = " ".join(parts).strip()
    return result[0].upper() + result[1:] + " đồng"


def merge_split_placeholders(xml):
    def process_para(m):
        para = m.group(0)
        t_pat = re.compile(r'<w:t(?=[>\s])[^>]*>(.*?)</w:t>', re.DOTALL)
        ts = list(t_pat.finditer(para))
        if not ts:
            return para
        full = "".join(x.group(1) for x in ts)
        if "{" not in full or "}" not in full:
            return para
        has_split = any(("{" in x.group(1) and "}" not in x.group(1)) or
                        ("}" in x.group(1) and "{" not in x.group(1)) for x in ts)
        if not has_split:
            return para
        new_para = para
        first = ts[0]
        safe = full.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        new_t = re.sub(r'(<w:t(?=[>\s])[^>]*>)(.*?)(</w:t>)',
                       lambda mm: mm.group(1) + safe + mm.group(3),
                       first.group(0), count=1, flags=re.DOTALL)
        new_para = new_para.replace(first.group(0), new_t, 1)
        for t in ts[1:]:
            cleared = re.sub(r'(<w:t(?=[>\s])[^>]*>)(.*?)(</w:t>)',
                             lambda mm: mm.group(1) + mm.group(3),
                             t.group(0), count=1, flags=re.DOTALL)
            new_para = new_para.replace(t.group(0), cleared, 1)
        return new_para
    return re.sub(r'<w:p[ >].*?</w:p>', process_para, xml, flags=re.DOTALL)


def apply_replacements(xml, data):
    for k, v in data.items():
        xml = xml.replace("{" + k + "}", v)
    return xml


# ── Đọc dữ liệu từ Excel ─────────────────────────────────────────────
excel_path = os.path.join(BASE_DIR, 'input', 'thong_tin_giam_dinh_xe.xlsx')
wb = openpyxl.load_workbook(excel_path)
ws_info = wb['Thông tin']

info = {}
for row in ws_info.iter_rows(values_only=True):
    if row[0] and str(row[0]).startswith('{') and row[2] is not None:
        key = str(row[0]).strip('{}')
        val = str(row[2]) if str(row[2]) != 'None' else ''
        info[key] = val

# Alias
info['ten_chu_xe'] = info.get('chu_xe', '')

# Số tiền thiệt hại: ưu tiên tien_tt
so_tien_so = info.get('tien_tt', '0').replace(' ', '')
info['so_tien_thiet_hai_so'] = so_tien_so
info['so_tien_thiet_hai_chu'] = so_sang_chu(so_tien_so)

print("Số tiền bằng chữ:", info['so_tien_thiet_hai_chu'])

# ── Sinh file docx ────────────────────────────────────────────────────
with zipfile.ZipFile(ASSET, 'r') as zin:
    all_files = zin.namelist()
    file_contents = {f: zin.read(f) for f in all_files}

xml = file_contents['word/document.xml'].decode('utf-8')
xml = merge_split_placeholders(xml)
xml = apply_replacements(xml, info)

import xml.etree.ElementTree as ET
try:
    ET.fromstring(xml)
    print("XML hợp lệ.")
except ET.ParseError as e:
    print("LỖI XML:", e)
    raise

file_contents['word/document.xml'] = xml.encode('utf-8')

with zipfile.ZipFile(OUTPUT, 'w', zipfile.ZIP_DEFLATED) as zout:
    for fname, content in file_contents.items():
        zout.writestr(fname, content)

print(f"Đã tạo: {OUTPUT}")
