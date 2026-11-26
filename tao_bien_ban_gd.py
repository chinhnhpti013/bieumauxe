# -*- coding: utf-8 -*-
"""Mẫu A — Biên bản giám định."""
from pti_common import (load_data, bang_bien_ban_gd, render,
                        bao_cao_placeholder_con_sot)

info, phu_tung = load_data()

# Mẫu có 17 dòng: STT, tên hạng mục, mức độ hư hỏng và các ô đánh dấu 'x'
bang_bien_ban_gd(info, phu_tung, n=17)
if len(phu_tung) > 17:
    print(f'  ⚠ Có {len(phu_tung)} hạng mục nhưng mẫu chỉ in được 17 dòng đầu.')

dst = render('Biên-bản-giám-định.docx', 'Bien-ban-giam-dinh.docx', info)
print('Tạo xong:', dst)
bao_cao_placeholder_con_sot(dst)
