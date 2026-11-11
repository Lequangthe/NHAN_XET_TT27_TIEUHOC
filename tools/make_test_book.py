#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
make_test_book.py  -  tao workbook test SoDiem_Test.xlsx bang openpyxl.

Dung o buoc -Test cua build.ps1 (tranh loi ghi mang qua Excel COM).

Cau truc moi sheet (giong so diem chi tiet cua vnEdu/SMIS):
    row 1 : TRUONG TIEU HOC PHU THUAN
    row 2 : SO DIEM CHI TIET
    row 3 : Mon: <ten mon> - <dot danh gia>
    row 4 : STT | Ho va ten | Ngay sinh | Nhan xet | KTDK | XL
    row 5-9 : 5 hoc sinh

Sheet name chua du Lop + Dot + Mon de add-in tu nhan dien:
    "Toan - Lop 3 - Giua ky 1" ...
O nhan xet cua hoc sinh cuoi (dong 9) duoc dien san chu "GIU NGUYEN"
de kiem tra cai dat GhiDe=0 (khong ghi de o da co chu).

Cach dung:
    python make_test_book.py <duong_dan.xlsx>
"""
from __future__ import annotations

import io
import json
import os
import sys

# (ten sheet, mon, dot, ma mon, ma dot)
SHEETS = [
    ("Toán - Lớp 3 - Giữa kỳ 1", "Toán", "Giữa kỳ 1", "Toan", "GK1"),
    ("Toán - Lớp 3 - Cuối kỳ 1", "Toán", "Cuối kỳ 1", "Toan", "HK1"),
    ("Tiếng Việt - Lớp 3 - Giữa kỳ 2", "Tiếng Việt", "Giữa kỳ 2", "TiengViet", "GK2"),
    ("TNXH - Lớp 3 - Cuối kỳ 2", "Tự nhiên và Xã hội", "Cuối kỳ 2", "TNXH", "CK2"),
]

HDR = ["STT", "Họ và tên", "Ngày sinh", "Nhận xét", "KTĐK", "XL"]
MARK = [10, 9, 8, 4, 7]
MARKER = "GIỮ NGUYÊN"
FIRST_DATA_ROW = 5


def make(path: str) -> str:
    from openpyxl import Workbook
    from datetime import date

    wb = Workbook()
    first = True
    for title, mon, dot, mon_code, dot_code in SHEETS:
        if first:
            ws = wb.active
            first = False
        else:
            ws = wb.create_sheet()
        ws.title = title[:31]

        ws["A1"] = "TRƯỜNG TIỂU HỌC PHÚ THUẬN"
        ws["A2"] = "SỔ ĐIỂM CHI TIẾT"
        ws["A3"] = "Môn: %s - %s - Lớp 3" % (mon, dot)

        for c, h in enumerate(HDR, start=1):
            ws.cell(row=4, column=c, value=h)

        for i, m in enumerate(MARK, start=1):
            r = FIRST_DATA_ROW + i - 1
            ws.cell(row=r, column=1, value=i)
            ws.cell(row=r, column=2, value="Học sinh số %d" % i)
            ws.cell(row=r, column=3, value=date(2017, 1, i))
            ws.cell(row=r, column=3).number_format = "dd/mm/yyyy"
            ws.cell(row=r, column=5, value=m)

        # o nhan xet cuoi cung: da co chu -> phai duoc giu nguyen khi GhiDe=0
        ws.cell(row=FIRST_DATA_ROW + len(MARK) - 1, column=4, value=MARKER)

        ws.column_dimensions["B"].width = 22
        ws.column_dimensions["D"].width = 60

    wb.save(path)

    # metadata cho build.ps1 (giu nguyen tieng Viet qua JSON UTF-8)
    meta = {
        "sheets": [
            {"ten": t[:31], "mon": m, "dot": d, "monCode": mc, "dotCode": dc}
            for t, m, d, mc, dc in SHEETS
        ],
        "nHs": len(MARK),
        "firstRow": FIRST_DATA_ROW,
        "lastRow": FIRST_DATA_ROW + len(MARK) - 1,
        "noteCol": 4,
        "maxCol": 6,
        "marker": MARKER,
    }
    meta_path = os.path.splitext(path)[0] + ".sheets.json"
    with io.open(meta_path, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)

    return path


def main(argv):
    if len(argv) < 2:
        print("dung: make_test_book.py <duong_dan.xlsx>")
        return 1
    out = make(argv[1])
    print("OK  test book -> %s (%d sheet)" % (out, len(SHEETS)))
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
