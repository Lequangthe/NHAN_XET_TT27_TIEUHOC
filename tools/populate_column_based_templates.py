import json
import os
import shutil
import ftfy
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

RAW_JSON_PATH = r"D:\01_CHUYEN_MON\AI_AGENT\BK\_DULIEUTHO\MON_HOC_GD_NLPC\mon_hoc_va_hdgd.json"

MAP_SUBJECTS = {
    'Tieng_Viet': 'TiengViet',
    'Toan': 'Toan',
    'Tu_Nhien_Xa_Hoi': 'TNXH',
    'Dao_Duc': 'DaoDuc',
    'Hoat_Dong_Trai_Nghiem': 'HDTN',
    'Am_Nhac': 'AmNhac',
    'My_Thuat': 'MyThuat',
    'Co_so': 'GDTC',
    'Gd_The_Chat': 'GDTC',
    'The_Chat': 'GDTC',
    'Khoa_Hoc': 'KhoaHoc',
    'Lich_Su_Dia_Ly': 'LSDL',
    'Tieng_Anh': 'TiengAnh',
    'Tin_Hoc': 'TinHoc',
    'Cong_Nghe': 'CongNghe'
}

def clean_vietnamese_text(text):
    if not isinstance(text, str):
        return str(text)
    txt = text.strip()
    fixed = ftfy.fix_text(txt)
    if len(fixed) > 0 and fixed[0].islower():
        fixed = fixed[0].upper() + fixed[1:]
    return fixed.strip()

def create_column_template(file_path, sub_name, list_t, list_h, list_c):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Mẫu Nhận Xét"

    font_header = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
    fill_header_t = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_header_h = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
    fill_header_c = PatternFill(start_color="C55A11", end_color="C55A11", fill_type="solid")

    font_data = Font(name="Times New Roman", size=11)

    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    headers = [
        "TỐT (Điểm 9 - 10 / Hoàn thành tốt)",
        "HOÀN THÀNH (Điểm 5 - 8 / Hoàn thành)",
        "CẦN CỐ GẮNG (Điểm < 5 / Chưa hoàn thành)"
    ]

    ws.append(headers)

    ws.cell(row=1, column=1).font = font_header
    ws.cell(row=1, column=1).fill = fill_header_t
    ws.cell(row=1, column=1).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.cell(row=1, column=2).font = font_header
    ws.cell(row=1, column=2).fill = fill_header_h
    ws.cell(row=1, column=2).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.cell(row=1, column=3).font = font_header
    ws.cell(row=1, column=3).fill = fill_header_c
    ws.cell(row=1, column=3).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.column_dimensions['A'].width = 55
    ws.column_dimensions['B'].width = 55
    ws.column_dimensions['C'].width = 55

    max_len = max(len(list_t), len(list_h), len(list_c), 1)

    for r in range(max_len):
        row_num = r + 2
        val_t = list_t[r] if r < len(list_t) else ""
        val_h = list_h[r] if r < len(list_h) else ""
        val_c = list_c[r] if r < len(list_c) else ""

        c1 = ws.cell(row=row_num, column=1, value=val_t)
        c2 = ws.cell(row=row_num, column=2, value=val_h)
        c3 = ws.cell(row=row_num, column=3, value=val_c)

        for c in (c1, c2, c3):
            c.font = font_data
            c.border = thin_border
            c.alignment = Alignment(vertical="top", wrap_text=True)

    os.makedirs(os.path.dirname(file_path), exist_ok=True)
    wb.save(file_path)

def extract_comments_from_raw(data, json_grade_key, raw_subj_key):
    g_data = data.get(json_grade_key, {})
    sub_data = g_data.get(raw_subj_key, {})

    list_t = []
    list_h = []
    list_c = []

    if isinstance(sub_data, dict):
        for raw_c in sub_data.get('T', []):
            txt = clean_vietnamese_text(raw_c) if isinstance(raw_c, str) else clean_vietnamese_text(raw_c.get('nhan_xet', raw_c.get('diem_manh', '')))
            if txt and txt not in list_t:
                list_t.append(txt)

        for raw_c in sub_data.get('H', []):
            txt = clean_vietnamese_text(raw_c) if isinstance(raw_c, str) else clean_vietnamese_text(raw_c.get('nhan_xet', raw_c.get('diem_manh', '')))
            if txt and txt not in list_h:
                list_h.append(txt)

        for raw_c in sub_data.get('C', []):
            txt = clean_vietnamese_text(raw_c) if isinstance(raw_c, str) else clean_vietnamese_text(raw_c.get('nhan_xet', raw_c.get('diem_manh', '')))
            if txt and txt not in list_c:
                list_c.append(txt)

    if not list_t:
        list_t = ["Hoàn thành xuất sắc các nội dung bài học, tiếp thu bài nhanh và tự tin phát biểu."]
    if not list_h:
        list_h = ["Hoàn thành tốt các bài tập trên lớp, chấp hành tốt yêu cầu của giáo viên."]
    if not list_c:
        list_c = ["Cần chú ý tập trung hơn trong giờ học và rèn luyện thêm ở nhà."]

    return list_t, list_h, list_c

def main():
    print("Loading master raw JSON data...")
    with open(RAW_JSON_PATH, "r", encoding="utf-8-sig") as f:
        data = json.load(f)

    dest_base = r"D:\GG\A\nx_tieuhoc\data\Mau_Nhan_Xet"

    # Wipe old destination directory completely to ensure no old files remain
    if os.path.exists(dest_base):
        shutil.rmtree(dest_base)

    count_files = 0
    total_comments = 0

    grades_map = {
        'lop_1': 'Lop_1',
        'lop_2': 'Lop_2',
        'lop_3': 'Lop_3',
        'lop_4': 'Lop_4',
        'lop_5': 'Lop_5'
    }

    for json_grade_key, folder_grade_name in grades_map.items():
        grade_data = data.get(json_grade_key, {})

        for raw_subj_key in grade_data.keys():
            file_subj_name = MAP_SUBJECTS.get(raw_subj_key, raw_subj_key)
            list_t, list_h, list_c = extract_comments_from_raw(data, json_grade_key, raw_subj_key)

            for hk in ['HK1', 'HK2']:
                out_path = os.path.join(dest_base, folder_grade_name, hk, f"{file_subj_name}.xlsx")
                create_column_template(out_path, file_subj_name, list_t, list_h, list_c)
                count_files += 1
                total_comments += (len(list_t) + len(list_h) + len(list_c))

    print(f"Generated {count_files} column-based template files with {total_comments} cleaned Vietnamese comments!")

    # Sync across all target locations
    targets = [
        r"C:\NX_MON\NXTiengViet_data\Mau_Nhan_Xet",
        r"C:\Users\Administrator\Downloads\NXTiengViet_data\Mau_Nhan_Xet",
        r"D:\GG\A\nx_tieuhoc\dist\NXTiengViet_data\Mau_Nhan_Xet",
        r"D:\01_CHUYEN_MON\AI_AGENT\BK\_cong_cu\NXTiengViet_data\Mau_Nhan_Xet",
        r"D:\01_CHUYEN_MON\AI_AGENT\BK\_cong_cu\nx_tieuhoc\data\Mau_Nhan_Xet"
    ]

    for t in targets:
        if os.path.exists(t):
            shutil.rmtree(t)
        shutil.copytree(dest_base, t)
        print(f"Synchronized templates to: {t}")

if __name__ == "__main__":
    main()
