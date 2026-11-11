import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# Bank of primary school TT27 comments
DATA_BANK = {
    'TiengViet': [
        ('9-10', 'Đọc to, rõ ràng, diễn cảm tốt. Viết chữ đẹp, đúng chính tả, vốn từ phong phú.'),
        ('9-10', 'Nắm vững kiến thức bài học, trả lời câu hỏi lưu khoát, chữ viết sạch đẹp.'),
        ('9-10', 'Kỹ năng đọc hiểu tốt, biết diễn đạt ý kiến cá nhân mạch lạc, tự tin.'),
        ('9-10', 'Viết câu đúng cú pháp, diễn đạt trôi chảy, viết đoạn văn sinh động.'),
        ('7-8', 'Đọc bài trôi chảy, nắm được nội dung bài đọc, trả lời được các câu hỏi.'),
        ('7-8', 'Hiểu bài tốt, hoàn thành các bài tập trên lớp. Cần cẩn thận hơn nét chữ.'),
        ('7-8', 'Đọc bài rõ ràng, viết bài đầy đủ. Cần phát biểu xây dựng bài nhiều hơn.'),
        ('5-6', 'Đọc bài rõ ràng, viết chữ khá sạch sẽ. Cần chú ý cách đặt câu.'),
        ('5-6', 'Nắm được nội dung bài đọc cơ bản. Cần chú ý tốc độ viết và chính tả.'),
        ('<5', 'Đọc bài còn chậm. Cần rèn luyện thêm kỹ năng đọc và viết ở nhà.'),
        ('Tốt', 'Hăng hái phát biểu, đọc diễn cảm tốt, tích cực chia sẻ bài viết hay với các bạn.'),
        ('Hoàn thành', 'Hoàn thành tốt các nhiệm vụ học tập môn Tiếng Việt trên lớp.'),
        ('Chưa hoàn thành', 'Cần nỗ lực nhiều hơn trong kỹ năng đọc hiểu và tập làm văn.')
    ],
    'Toan': [
        ('9-10', 'Tính toán nhanh, chính xác, tư duy toán học tốt. Giải toán có lời văn thành thạo.'),
        ('9-10', 'Nắm vững các phép tính, kỹ năng làm bài cẩn thận, độ chính xác cao.'),
        ('9-10', 'Tư duy logic tốt, hoàn thành xuất sắc các bài tập toán trên lớp.'),
        ('9-10', 'Thực hiện thành thạo các phép tính, trình bày bài giải sạch đẹp, rõ ràng.'),
        ('7-8', 'Thực hiện tốt các phép tính cơ bản. Cần rèn thêm kỹ năng giải toán có lời văn.'),
        ('7-8', 'Hoàn thành các bài tập trên lớp. Cần cẩn thận hơn khi tính toán số lớn.'),
        ('7-8', 'Nắm được kiến thức cơ bản. Trình bày bài giải rõ ràng, sạch sẽ.'),
        ('5-6', 'Thực hiện được các phép tính cơ bản. Cần tính toán cẩn thận hơn.'),
        ('5-6', 'Biết cách làm bài tập toán. Cần tập trung hơn khi làm bài.'),
        ('<5', 'Tính toán còn chậm và chưa chính xác. Cần rèn thêm các phép tính cơ bản.'),
        ('Tốt', 'Tư duy nhanh nhạy, tính toán chính xác, tự tin giải các bài toán nâng cao.'),
        ('Hoàn thành', 'Hoàn thành đầy đủ các bài tập toán theo yêu cầu môn học.'),
        ('Chưa hoàn thành', 'Cần tăng cường rèn luyện kỹ năng tính toán cơ bản hằng ngày.')
    ],
    'TNXH': [
        ('Tốt', 'Hăng hái phát biểu, tích cực tham gia thảo luận nhóm, nắm chắc bài học.'),
        ('Tốt', 'Có hiểu biết tốt về thế giới tự nhiên và xã hội, liên hệ thực tế giỏi.'),
        ('Tốt', 'Nắm vững nội dung bài học, có ý thức giữ gìn vệ sinh và bảo vệ môi trường.'),
        ('Hoàn thành', 'Nắm được các kiến thức cơ bản của môn học. Cần mạnh dạn phát biểu hơn.'),
        ('Hoàn thành', 'Hoàn thành nhiệm vụ học tập trên lớp. Cần chú ý liên hệ kiến thức thực tế.'),
        ('Hoàn thành', 'Hiểu bài và tham gia học tập tích cực. Lắng nghe thầy cô giảng bài.'),
        ('Chưa hoàn thành', 'Cần chú ý tập trung hơn trong giờ học Tự nhiên và Xã hội.')
    ],
    'DaoDuc': [
        ('Tốt', 'Ngoan ngoãn, lễ phép với thầy cô, hòa đồng với bạn bè. Thực hiện tốt quy định.'),
        ('Tốt', 'Có hành vi đạo đức tốt, biết giúp đỡ bạn bè, trung thực và trách nhiệm.'),
        ('Tốt', 'Luôn tôn trọng người lớn, biết vâng lời, có ý thức kỷ luật cao.'),
        ('Hoàn thành', 'Lễ phép, ngoan ngoãn. Cần chủ động giúp đỡ bạn bè nhiều hơn.'),
        ('Hoàn thành', 'Chấp hành tốt nội quy trường lớp. Cần phát huy tinh thần tự giác.'),
        ('Hoàn thành', 'Hiểu được các bài học đạo đức, thể hiện ngoan ngoãn trong lớp.'),
        ('Chưa hoàn thành', 'Cần rèn luyện thêm ý thức tự giác và chấp hành nội quy.')
    ],
    'HDTN': [
        ('Tốt', 'Tích cực tham gia các hoạt động tập thể, năng nổ, hợp tác tốt với bạn bè.'),
        ('Tốt', 'Có kỹ năng sống tốt, tự giác hoàn thành nhiệm vụ được giao, gương mẫu.'),
        ('Tốt', 'Nhiệt tình, sáng tạo trong các hoạt động trải nghiệm, có tinh thần đồng đội.'),
        ('Hoàn thành', 'Tham gia đầy đủ các hoạt động trải nghiệm. Cần chủ động hơn.'),
        ('Hoàn thành', 'Hoàn thành tốt nhiệm vụ được giao. Hòa đồng với các bạn trong nhóm.'),
        ('Hoàn thành', 'Có cố gắng tham gia các hoạt động chung của lớp và trường.'),
        ('Chưa hoàn thành', 'Cần hăng hái và tự tin hơn khi tham gia các hoạt động trải nghiệm.')
    ]
}

def create_template_file(file_path, sub_name, rows):
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

    headers = ["Mức Điểm / Xếp Loại", "Nội Dung Nhận Xét Mẫu"]
    ws.append(headers)

    ws.cell(row=1, column=1).font = font_header
    ws.cell(row=1, column=1).fill = fill_header
    ws.cell(row=1, column=2).font = font_header
    ws.cell(row=1, column=2).fill = fill_header

    ws.column_dimensions['A'].width = 24
    ws.column_dimensions['B'].width = 80

    for r_idx, (level, text) in enumerate(rows, 2):
        ws.cell(row=r_idx, column=1, value=level).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=r_idx, column=2, value=text).alignment = Alignment(vertical="center", wrap_text=True)

        ws.cell(row=r_idx, column=1).font = font_data
        ws.cell(row=r_idx, column=2).font = font_data
        ws.cell(row=r_idx, column=1).border = thin_border
        ws.cell(row=r_idx, column=2).border = thin_border

    os.makedirs(os.path.dirname(file_path), exist_ok=True)
    wb.save(file_path)

def main():
    base_dirs = [
        r"D:\GG\A\nx_tieuhoc\data\Mau_Nhan_Xet",
        r"D:\GG\A\nx_tieuhoc\dist\NXTiengViet_data\Mau_Nhan_Xet"
    ]

    classes = [f"Lop_{i}" for i in range(1, 6)]
    hks = ["HK1", "HK2"]

    count = 0
    for b_dir in base_dirs:
        for c in classes:
            for hk in hks:
                for sub_key, rows in DATA_BANK.items():
                    fp = os.path.join(b_dir, c, hk, f"{sub_key}.xlsx")
                    create_template_file(fp, sub_key, rows)
                    count += 1

    print(f"Successfully generated {count} Excel template .xlsx files across all classes and semesters!")

if __name__ == "__main__":
    main()
