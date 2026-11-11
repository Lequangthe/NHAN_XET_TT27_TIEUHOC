import json
import os
import shutil
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

LEVEL_LABEL_MAP = {
    'T': ['9-10', 'Tốt'],
    'H': ['7-8', '5-6', 'Hoàn thành'],
    'C': ['<5', 'Chưa hoàn thành']
}

def create_excel_template(file_path, comment_rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Mẫu Nhận Xét"

    font_header = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
    fill_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    font_data = Font(name="Times New Roman", size=11)

    thin_border = Border(
        left=Side(style='thin', color='A6A6A6'),
        right=Side(style='thin', color='A6A6A6'),
        top=Side(style='thin', color='A6A6A6'),
        bottom=Side(style='thin', color='A6A6A6')
    )

    ws.append(["Mức Điểm / Xếp Loại", "Nội Dung Nhận Xét Mẫu"])

    ws.cell(row=1, column=1).font = font_header
    ws.cell(row=1, column=1).fill = fill_header
    ws.cell(row=1, column=2).font = font_header
    ws.cell(row=1, column=2).fill = fill_header

    ws.column_dimensions['A'].width = 24
    ws.column_dimensions['B'].width = 85

    r_idx = 2
    for lvl_tag, text in comment_rows:
        ws.cell(row=r_idx, column=1, value=lvl_tag).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=r_idx, column=2, value=text).alignment = Alignment(vertical="center", wrap_text=True)

        ws.cell(row=r_idx, column=1).font = font_data
        ws.cell(row=r_idx, column=2).font = font_data
        ws.cell(row=r_idx, column=1).border = thin_border
        ws.cell(row=r_idx, column=2).border = thin_border
        r_idx += 1

    os.makedirs(os.path.dirname(file_path), exist_ok=True)
    wb.save(file_path)

def main():
    print("Loading master raw JSON data...")
    with open(RAW_JSON_PATH, "r", encoding="utf-8-sig") as f:
        data = json.load(f)

    dest_base = r"D:\GG\A\nx_tieuhoc\data\Mau_Nhan_Xet"

    count_files = 0
    total_comments_written = 0

    grades_map = {
        'lop_1': 'Lop_1',
        'lop_2': 'Lop_2',
        'lop_3': 'Lop_3',
        'lop_4': 'Lop_4',
        'lop_5': 'Lop_5'
    }

    for json_grade_key, folder_grade_name in grades_map.items():
        grade_data = data.get(json_grade_key, {})

        for raw_subj_key, level_dict in grade_data.items():
            file_subj_name = MAP_SUBJECTS.get(raw_subj_key, raw_subj_key)

            comment_rows = []

            for lvl_key in ['T', 'H', 'C']:
                c_list = level_dict.get(lvl_key, [])
                tags = LEVEL_LABEL_MAP.get(lvl_key, [lvl_key])

                for idx, c_text in enumerate(c_list):
                    tag = tags[idx % len(tags)]
                    if isinstance(c_text, str) and c_text.strip():
                        comment_rows.append((tag, c_text.strip()))
                    elif isinstance(c_text, dict):
                        txt = c_text.get('nhan_xet', c_text.get('diem_manh', ''))
                        if txt:
                            comment_rows.append((tag, txt.strip()))

            # If no comments found, fallback 1 sample line
            if not comment_rows:
                comment_rows = [
                    ('Tốt', 'Hoàn thành xuất sắc các nội dung môn học, tiếp thu bài nhanh.'),
                    ('Hoàn thành', 'Hoàn thành tốt các nhiệm vụ học tập môn học trên lớp.'),
                    ('Chưa hoàn thành', 'Cần rèn luyện và cố gắng nhiều hơn trong học tập.')
                ]

            # Populate for both HK1 and HK2
            for hk in ['HK1', 'HK2']:
                out_path = os.path.join(dest_base, folder_grade_name, hk, f"{file_subj_name}.xlsx")
                create_excel_template(out_path, comment_rows)
                count_files += 1
                total_comments_written += len(comment_rows)

    print(f"Generated {count_files} template files with {total_comments_written} authentic comments!")

    # Synchronize to dist, Downloads, and BK
    dist_dir = r"D:\GG\A\nx_tieuhoc\dist\NXTiengViet_data\Mau_Nhan_Xet"
    downloads_dir = r"C:\Users\Administrator\Downloads\NXTiengViet_data\Mau_Nhan_Xet"
    bk_dir = r"D:\01_CHUYEN_MON\AI_AGENT\BK\_cong_cu\NXTiengViet_data\Mau_Nhan_Xet"

    for target in [dist_dir, downloads_dir, bk_dir]:
        if os.path.exists(target):
            shutil.rmtree(target)
        shutil.copytree(dest_base, target)
        print(f"Synchronized templates to: {target}")

if __name__ == "__main__":
    main()
