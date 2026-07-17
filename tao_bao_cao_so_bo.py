# -*- coding: utf-8 -*-
"""Mẫu C — Báo cáo sơ bộ giám định."""
from pti_common import (load_data, danh_sach_phu_tung, render,
                        bao_cao_placeholder_con_sot)

info, phu_tung = load_data()
danh_sach_phu_tung(info, phu_tung, n=14)

dst = render('Báo-cáo-sơ-bộ-giám-định.docx', 'Bao-cao-so-bo.docx', info)
print('Tạo xong:', dst)
bao_cao_placeholder_con_sot(dst)
