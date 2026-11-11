import re

test_titles = [
    'SỔ ĐÁNH GIÁ ĐỊNH KỲ LỚP 3A6',
    'BẢNG TỔNG HỢP MÔN TOÁN LỚP 5B HKII',
    'MÔN TIẾNG VIỆT 1/2 HỌC KỲ I',
    'BẢNG ĐIỂM LỚP 4C CẢ NĂM',
    'Sheet1',
    'Lop_2_HK2_Toan'
]

def detect_class(text):
    text_u = text.upper()
    m = re.search(r'(?:LỚP|LOP|KHỐI|KHOI)\s*([1-5])|\b([1-5])[A-Z0-9/\.\-_]', text_u)
    if m:
        g = m.group(1) or m.group(2)
        return f'Lop_{g}'
    return ''

def detect_semester(text):
    text_u = text.upper()
    if any(k in text_u for k in ['HK2', 'HK 2', 'HKII', 'HỌC KỲ 2', 'HỌC KỲ II', 'HOC KY 2', 'CUỐI KỲ 2', 'GIỮA KỲ 2']):
        return 'HK2'
    return 'HK1'

for t in test_titles:
    print(f'Text: "{t}" -> Class: {detect_class(t)} | Semester: {detect_semester(t)}')
