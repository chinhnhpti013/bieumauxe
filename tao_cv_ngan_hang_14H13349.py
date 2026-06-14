# -*- coding: utf-8 -*-
# CV Ngân hàng — xe 14H-133.49 (HDBank Quảng Ninh)
import zipfile, re, os

ASSET = r"d:\MCP\Claude Code\mau bieu giam dinh pti\assets\CV-Ngân hàng.docx"
OUTPUT = r"d:\MCP\Claude Code\mau bieu giam dinh pti\output\CV-Ngan-hang-14H-133.49.docx"

os.makedirs(os.path.dirname(OUTPUT), exist_ok=True)

def so_sang_chu(so_str):
    so = int(so_str.replace(",", "").replace(".", ""))
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
        new_t = re.sub(r'(<w:t(?=[>\s])[^>]*>)(.*?)(</w:t>)',
                       lambda mm: mm.group(1) + full + mm.group(3),
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
        v_safe = v.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        xml = xml.replace("{" + k + "}", v_safe)
    return xml


TIEN_SO = "4.558.334"
TIEN_CHU = so_sang_chu(TIEN_SO)

data = {
    "ten_ngan_hang": "Ngân hàng TMCP Phát triển Thành phố Hồ Chí Minh (HDBank) - Chi nhánh Quảng Ninh",
    "dia_chi_ngan_hang": "Số 9A-10A, Đường 25/4, Phường Hồng Gai, Tỉnh Quảng Ninh",
    "bien_so_xe": "14H-133.49",
    "chu_xe": "NGUYỄN NGỌC ANH",
    "ten_chu_xe": "NGUYỄN NGỌC ANH",
    "dia_chi_chu_xe": "Thanh Sơn, Uông Bí, Quảng Ninh",
    "so_khung": "RLLVFPNT6TH737098",
    "so_gcn_bh": "/ACDT+I013TN+260000156",
    "gcn_bh_tu_ngay": "16/04/2026",
    "gcn_bh_den_ngay": "16/04/2027",
    "ngay_tai_nan": "10/06/2026",
    "lai_xe": "Nguyễn Ngọc Anh",
    "dia_diem_tai_nan": "Số nhà 40 đập tràn nhà máy điện Uông Bí, Khu Trưng Vương, Phường Vàng Danh, Quảng Ninh",
    "so_tien_thiet_hai_so": TIEN_SO,
    "so_tien_thiet_hai_chu": TIEN_CHU,
}

print("Số tiền bằng chữ:", TIEN_CHU)

with zipfile.ZipFile(ASSET, 'r') as zin:
    all_files = zin.namelist()
    file_contents = {f: zin.read(f) for f in all_files}

xml = file_contents['word/document.xml'].decode('utf-8')
xml = merge_split_placeholders(xml)
xml = apply_replacements(xml, data)
# Mẫu gốc đã có sẵn chữ "Ngân hàng" trước {ten_ngan_hang} ở dòng Kính gửi
xml = xml.replace("Ngân hàng Ngân hàng TMCP", "Ngân hàng TMCP")

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
