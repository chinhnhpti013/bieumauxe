# -*- coding: utf-8 -*-
"""Mẫu B, F, H — Vật tư thu hồi, Xác nhận bồi thường, Biên bản nghiệm thu."""
from pti_common import (load_data, danh_sach_phu_tung, phu_tung_thu_hoi,
                        so_thanh_chu, render, bao_cao_placeholder_con_sot)

info, phu_tung = load_data()

# Chỉ mẫu "Vật tư thu hồi" dùng {sttN}/{ten_phu_tung_N}; hai mẫu còn lại không có
# placeholder nào của bảng hạng mục nên dùng chung info là an toàn.
# Bảng này chỉ liệt kê hạng mục phương án "Thay thế có thu hồi", nhưng STT vẫn giữ
# số thứ tự gốc trên tab Phụ tùng ở Bước 2.
pt_thu_hoi = phu_tung_thu_hoi(phu_tung)
danh_sach_phu_tung(info, pt_thu_hoi, n=14, stt=True)
if len(pt_thu_hoi) > 14:
    print(f'  ⚠ Có {len(pt_thu_hoi)} vật tư thu hồi nhưng mẫu chỉ in được 14 dòng đầu.')

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
