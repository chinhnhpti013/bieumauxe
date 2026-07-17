# -*- coding: utf-8 -*-
"""Mẫu B, F, H — Vật tư thu hồi, Xác nhận bồi thường, Biên bản nghiệm thu."""
from pti_common import (load_data, danh_sach_phu_tung, phu_tung_thu_hoi,
                        so_thanh_chu, render, bao_cao_placeholder_con_sot)

info, phu_tung = load_data()

danh_sach_phu_tung(info, phu_tung, n=14)

# Vật tư thu hồi chỉ liệt kê phụ tùng phương án "Thay thế có thu hồi"
pt_thu_hoi = phu_tung_thu_hoi(phu_tung)
danh_sach_phu_tung(info, pt_thu_hoi, prefix='ten_pt_thu_hoi', n=13)

info['tien_tt_chu'] = so_thanh_chu(info.get('tien_tt', '0'))

for asset, out in [
    ('Xác-nhận-bồi-thường.docx', 'Xac-nhan-boi-thuong.docx'),
    ('Vật-tư-thu-hồi.docx',      'Vat-tu-thu-hoi.docx'),
    ('Biên-bản-nghiệm-thu.docx', 'Bien-ban-nghiem-thu.docx'),
]:
    dst = render(asset, out, info)
    print('Tạo xong:', dst)
    bao_cao_placeholder_con_sot(dst)

print()
print(f'Tiền bồi thường: {info.get("tien_tt", "")} → {info["tien_tt_chu"]}')
if pt_thu_hoi:
    print(f'Phụ tùng thu hồi: {[p["ten"] for p in pt_thu_hoi]}')
else:
    print('Không có phụ tùng thu hồi (Vật tư thu hồi để trống)')
