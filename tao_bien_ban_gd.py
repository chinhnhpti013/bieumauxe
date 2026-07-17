# -*- coding: utf-8 -*-
"""Mẫu A — Biên bản giám định."""
from pti_common import (load_data, danh_sach_phu_tung, render,
                        bao_cao_placeholder_con_sot)

info, phu_tung = load_data()
danh_sach_phu_tung(info, phu_tung, n=13)

dst = render('Biên-bản-giám-định.docx', 'Bien-ban-giam-dinh.docx', info)
print('Tạo xong:', dst)
bao_cao_placeholder_con_sot(dst)
