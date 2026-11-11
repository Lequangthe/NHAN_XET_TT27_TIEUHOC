# NX Tiểu học (NXTiengViet) — add-in Excel điền nhận xét theo TT27

Add-in Excel dành cho **cấp Tiểu học**: đọc sổ điểm (vnEdu / SMIS / file Excel tự làm), nhận diện **Lớp – Đợt đánh giá – Môn**, rồi điền câu nhận xét vào đúng cột cho cả lớp.

- 4 đợt đánh giá: **Giữa kỳ 1 (GK1) – Cuối kỳ 1 (HK1) – Giữa kỳ 2 (GK2) – Cuối kỳ 2 (CK2)**
- 5 khối lớp, mỗi khối có bộ môn riêng, mỗi môn 3 mức: **TỐT / HOÀN THÀNH / CẦN CỐ GẮNG**
- Ngân hàng câu nằm trong file Excel mẫu → giáo viên tự sửa bằng Excel, không cần sửa mã

> File cần copy mang theo (giữ nguyên tên, nằm cùng thư mục):
> - `NXTiengViet.xlam` — add-in
> - `NXTiengViet_data\` — thư mục dữ liệu (nằm cạnh file `.xlam`)

## 1. Cài đặt

1. Mở Excel.
2. **File → Options → Add-ins → Manage: Excel Add-ins → Go… → Browse…** → chọn `NXTiengViet.xlam` → **OK**.
3. Tab **"NHẬN XÉT TIỂU HỌC (TT27)"** xuất hiện trên Ribbon (sau tab Review).

Nếu chưa thấy tab: đóng rồi mở lại Excel (ribbon nhúng sẵn trong file).

> Chỉ cần bật "Trust access to the VBA project object model" khi **build lại** add-in, không cần khi dùng.

## 2. Cấu trúc dữ liệu mẫu

```
<Thu muc cai add-in>\
  NXTiengViet.xlam            # add-in
  NXTiengViet_data\
    thongdiep.txt             # toàn bộ chuỗi giao diện (UTF-8)
    caidat.ini                # cài đặt, tự tạo khi chạy lần đầu
    nhat_ky.txt               # nhật ký các lần điền
    backup\                   # bản sao lưu sổ điểm trước khi ghi
    Mau_Nhan_Xet\
      Lop_1\ Lop_2\ Lop_3\ Lop_4\ Lop_5\
        GK1\ HK1\ GK2\ CK2\
          <Mon>.xlsx          # vd Toan.xlsx, TiengViet.xlsx, TNXH.xlsx ...
```

Mỗi file mẫu `<Mon>.xlsx`:

| Cột A | Cột B | Cột C |
|---|---|---|
| TỐT | HOÀN THÀNH | CẦN CỐ GẮNG |

- Dòng 1 là **tiêu đề 3 mức**, câu nhận xét bắt đầu từ **dòng 2**.
- Mỗi dòng 1 câu (một số câu có thể gồm 2 ý, vẫn nằm trong 1 ô).
- Bộ môn theo khối:
  - **Lớp 1–2–3**: AmNhac, DaoDuc, GDTC, HDTN, MyThuat, TiengViet, TNXH, Toan
  - **Lớp 4–5**: CongNghe, DaoDuc, GDTC, HDTN, KhoaHoc, LSDL, TiengAnh, TiengViet, TinHoc, Toan
- Tổng dữ liệu chuẩn: **176 file** (5 lớp × 4 đợt × số môn).

Danh sách lớp/môn **không hard-code**: add-in quét thư mục nên chỉ cần thêm/bớt file là có ngay (bấm **Nạp lại Mẫu**).

## 3. Cách dùng

1. Mở file sổ điểm, bấm vào **sheet của môn cần điền**.
2. Bấm **Chọn Lớp & Môn** → chọn lần lượt:
   - **1. Khối lớp** (Lop_1 … Lop_5)
   - **2. Đợt đánh giá** (GK1 / HK1 / GK2 / CK2)
   - **3. Môn học** (chỉ liệt kê môn có file mẫu ở lớp + đợt đó)
3. (Tuỳ chọn) Bấm **Nhận diện từ sheet** để add-in tự đọc tên sheet + tiêu đề và chọn đúng Lớp/Đợt/Môn.
4. (Tuỳ chọn) Bấm **Chế độ điền** để chọn cách xác định mức đạt.
5. Bấm **Xem trước** → add-in tạo sheet `NX_XEMTRUOC` liệt kê câu sẽ điền, **chưa ghi gì**.
6. Bấm **Điền nhận xét** → hộp thoại xác nhận → add-in:
   - sao lưu sổ điểm vào `NXTiengViet_data\backup\`,
   - ghi câu vào **cột Nhận xét** của **sheet đang mở**,
   - bỏ qua ô đã có chữ (mặc định `GhiDe=0`),
   - ghi nhật ký vào `nhat_ky.txt`.
7. Bấm **Hoàn tác (Undo)** nếu muốn trả lại nguyên trạng cột nhận xét.

> Add-in chỉ điền **sheet đang mở** (không điền toàn bộ workbook) — đúng với cách sổ điểm vnEdu/SMIS tách mỗi môn một sheet.

## 4. Chế độ điền (mức đạt)

| Mã | Chế độ | Cách xác định mức |
|---|---|---|
| 1 | Theo cột Điểm / Xếp loại | đọc cột xếp loại, nếu trống thì đọc cột điểm |
| 2 | Xoay vòng ngẫu nhiên | xoay TỐT & HOÀN THÀNH, không cần cột điểm |
| 3 | Tất cả mức TỐT | mọi học sinh = T |
| 4 | Tất cả mức HOÀN THÀNH | mọi học sinh = H |

Với chế độ 1, thứ tự dò:

1. Cột **Xếp loại** (`XL`, `Xếp loại`, `Mức đạt`, `Học lực`…): `T/HTT/HTTT…` → T, `H/HT/Đ…` → H, `C/CHT…` → C.
2. Không có xếp loại → cột điểm **KT** (`KTĐK`, `KT giữa kỳ`, `KT cuối kỳ`…): `>= 9` → T, `>= 5` → H, `< 5` → C.
3. Không có cả hai → dùng `MucMacDinh` trong `caidat.ini` (mặc định `H`).

Mức nào thiếu câu ở môn đang chọn, add-in tự lấy theo thứ tự: **đúng môn → đợt khác của cùng môn → môn khác đã nạp**.

## 5. Cài đặt `caidat.ini`

| Khoá | Mặc định | Ý nghĩa |
|---|---|---|
| `GhiDe` | `0` | `0` = bỏ qua ô đã có chữ; `1` = ghi đè |
| `CotNhanXet` | `auto` | `auto`/`last` = cột cuối (hoặc cột đúng đợt), `first` = cột đầu, `all` = mọi cột nhận xét, `<số>`/`<chữ cái>` = chỉ định cột |
| `CotDiemManual` | (trống) | chỉ định cột điểm, trống = tự dò |
| `CotNhanXetManual` | (trống) | chỉ định cột nhận xét, ưu tiên hơn `CotNhanXet` |
| `HangDauTienManual` | (trống) | hàng học sinh đầu tiên, trống = tự dò |
| `MucMacDinh` | `H` | mức dùng khi không đọc được điểm/xếp loại |
| `CheDoDienManual` | `1` | chế độ điền 1–4 |
| `LopSelected` / `DotSelected` / `MonSelected` | `Lop_3` / `HK1` / `Toan` | lựa chọn hiện hành |
| `ThuMucMau` | (trống) | thư mục mẫu tự chọn (nút "Chọn thư mục Mẫu") |
| `CauHinhVersion` | `3` | phiên bản cấu hình, add-in tự nâng cấp file cũ |

Sửa `caidat.ini` bằng Notepad (lưu **UTF-8**), sau đó bấm **Nạp lại Mẫu** hoặc mở lại Excel.

## 6. Chuỗi giao diện `thongdiep.txt`

Toàn bộ chữ hiển thị (hộp thoại, tiêu đề, hướng dẫn) nằm trong `NXTiengViet_data\thongdiep.txt` (UTF-8), dạng `khoá = giá trị`, `\n` = xuống dòng.

- Muốn đổi câu chữ: sửa file này, **không cần build lại** add-in.
- Thiếu khoá nào thì add-in dùng chuỗi mặc định trong mã.

## 7. Danh mục nút Ribbon

| Nhóm | Nút | Chức năng |
|---|---|---|
| ⚡ Điền Nhận Xét | **Điền nhận xét** | Điền thật cho sheet đang mở |
| | **Chế độ điền** | Chọn chế độ 1–4 |
| | **Chọn Lớp & Môn** | Chọn Lớp → Đợt → Môn |
| | **Nhận diện từ sheet** | Tự đoán Lớp/Đợt/Môn từ sheet đang mở |
| | **Chọn cột & dòng** | Chỉ định cột điểm, cột nhận xét, hàng đầu |
| | **Hoàn tác (Undo)** | Trả lại dữ liệu trước khi điền |
| | **Xem trước** | Tạo sheet `NX_XEMTRUOC`, không ghi |
| 📊 Mẫu Excel (.xlsx) | **Mở mẫu môn hiện tại** | Mở file `<Mon>.xlsx` đang chọn |
| | **Mở thư mục Mẫu** | Mở `Mau_Nhan_Xet` |
| | **Chọn thư mục Mẫu** | Trỏ tới thư mục mẫu ở ổ khác |
| | **Nạp lại Mẫu** | Xoá cache, đọc lại câu từ file mẫu |
| ❓ Trợ Giúp | **Hướng dẫn** | Hướng dẫn sử dụng |
| | **Thông tin** | Phiên bản, đường dẫn dữ liệu, mẫu đang chọn |

## 8. Build lại (khi sửa mã)

Cần Excel 2016/365 (16.x) + Python 3.13 + `openpyxl`.

```powershell
# build + đóng gói
powershell -NoProfile -ExecutionPolicy Bypass -File .\build.ps1

# build + chạy test thật (mở Excel COM, điền thử 4 đợt)
powershell -NoProfile -ExecutionPolicy Bypass -File .\build.ps1 -Test

# build từ bộ dữ liệu khác (mặc định lấy .\data)
powershell -NoProfile -ExecutionPolicy Bypass -File .\build.ps1 -DataSource "D:\...\NX_MON\NXTiengViet_data"

# deploy: sao lưu bản cũ trong NX_MON rồi copy .xlam + thongdiep.txt + Mau_Nhan_Xet
powershell -NoProfile -ExecutionPolicy Bypass -File .\build.ps1 -Deploy
```

`-Test` sẽ:

1. tạo sổ điểm giả 4 sheet (`Toán – Lớp 3 – Giữa kỳ 1`, `Toán – Lớp 3 – Cuối kỳ 1`, `Tiếng Việt – Lớp 3 – Giữa kỳ 2`, `TNXH – Lớp 3 – Cuối kỳ 2`) bằng `openpyxl`;
2. nạp add-in thật, chạy **Xem trước** rồi **Điền** từng sheet;
3. kiểm tra: báo cáo `OK Lop_3 <đợt> <môn> write=4 skip=1`, ô đã có chữ được giữ nguyên, các cột khác không đổi, câu không trùng, có `nhat_ky.txt`;
4. bật `GhiDe=1` → kiểm tra ghi đè toàn bộ (`write=5 skip=0`) rồi trả lại `GhiDe=0`;
5. kiểm tra `thongdiep.txt` nạp đúng UTF-8 (có dấu tiếng Việt).

Script build tự bật `AccessVBOM` lúc đóng gói rồi **khôi phục** trạng thái cũ; file cũ trong `dist` được đưa vào **Thùng rác** chứ không xoá vĩnh viễn.

## 9. Lưu ý kỹ thuật

- **Mã VBA chỉ dùng ASCII.** Máy dùng ANSI 1252 nên chữ có dấu viết trực tiếp trong `.bas` bị hỏng khi biên dịch. Mọi chuỗi tiếng Việt phải đọc từ `thongdiep.txt` (UTF-8) qua `NX_T()` / `NX_Fmt()`.
- **`build.ps1` cũng phải là ASCII** (hoặc UTF-8 **có BOM**). PowerShell 5.1 đọc file không BOM theo ANSI 1252 → chuỗi tiếng Việt vỡ và script báo lỗi cú pháp. Vì vậy tên sheet test được sinh từ `tools\make_test_book.py` ra file `.sheets.json` (UTF-8) thay vì viết cứng trong `.ps1`.
- **Ribbon nhúng trong file**: `customUI/customUI.xml` + relationship, `onAction` không kèm đường dẫn.
- **Ngân hàng mẫu nạp lười (lazy)**: chỉ mở file `.xlsx` của môn đang dùng, không mở cả 176 file mỗi lần chạy.
- **Fallback đợt** khi lớp/môn thiếu đúng đợt: GK1→HK1→GK2→CK2; HK1→GK1→CK2→GK2; GK2→CK2→HK1→GK1; CK2→GK2→HK1→GK1. Thư mục cũ `HK2` được hiểu là `CK2`, `CK1` là `HK1`.
- **Không lặp câu**: mỗi mức được xáo trộn (Fisher–Yates) và rút tuần tự, đồng thời tránh lặp câu vừa dùng ở ô liền trước.
- Module: `modUtil` (chuẩn hoá chuỗi/mảng), `modApp` (trạng thái ứng dụng), `modPath` (đường dẫn, 4 đợt, UTF-8, INI, thông điệp), `modBank` (ngân hàng câu lazy), `modSheet` (dò layout/đợt/lớp), `modFill` (luồng điền, backup, undo, nhật ký), `modMain` (nút Ribbon), `ThisWorkbook` (khởi tạo dữ liệu lần đầu).
- Không dùng `DecimalSeparator`, mọi module đều `Option Explicit`, không có UserForm.

## 10. Cấu trúc dự án

```
nx_tieuhoc\
  build.ps1                  # đóng gói .xlam + chèn ribbon + test + deploy
  src\                       # 7 module .bas + ThisWorkbook.cls
  customUI\customUI.xml      # định nghĩa Ribbon
  tools\inject_customui.py   # chèn customUI vào .xlam
  tools\make_test_book.py    # tạo sổ điểm test + .sheets.json
  data\Mau_Nhan_Xet\         # 176 file mẫu (nguồn)
  data\thongdiep.txt         # chuỗi giao diện UTF-8 (nguồn)
  dist\NXTiengViet.xlam      # bản phát hành
  dist\NXTiengViet_data\     # dữ liệu kèm theo bản phát hành
  _backup_20261108\          # bản sao lưu trước khi nâng cấp
```
