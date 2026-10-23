# DỰ ÁN GREENTECH – NHẬP MÔN CNTT

## 1. Tổng quan

### 1.1. Mục đích

Nội dung giúp sinh viên:

- Hiểu mô hình xử lý thông tin IPO và tư duy hệ thống DIKW để tổ chức dữ liệu khoa học.
- Thiết kế kiến trúc lưu trữ dự án theo mô hình cây thư mục chuẩn hóa.
- Sử dụng Terminal, Shell Script và toán tử Pipe để xử lý tệp nhật ký dung lượng lớn.
- Xây dựng tài liệu kỹ thuật chuyên nghiệp bằng Markdown và Word Styles.
- Vận dụng tư duy Database Mindset và hàm VLOOKUP để đối soát dữ liệu trong Excel.

### 1.2. Công nghệ sử dụng

- Command Line Interface (CMD / Terminal / Shell).
- Shell Scripting (`.bat` / `.sh`).
- Markdown.
- Microsoft Word (Heading Styles, Auto TOC).
- Microsoft Excel (VLOOKUP, TRIM, AutoFilter).
- Git và GitHub.

# PHẦN 1. TƯ DUY HỆ THỐNG VÀ KIẾN TRÚC LƯU TRỮ

## 2. Mô hình xử lý thông tin IPO và DIKW

### 2.1. Mục tiêu

Sinh viên hiểu được:

- Vì sao hệ thống hạ tầng cần được nhìn nhận theo luồng xử lý thông tin.
- Mô hình IPO là gì và cách áp dụng vào bài toán quản lý log.
- Mô hình DIKW là gì và cách chuyển dữ liệu thô thành quyết định vận hành.

### 2.2. Mô hình IPO (Input – Process – Output)

- **Input:** Tệp `log-he-thong.log` phát sinh liên tục từ máy chủ.
- **Process:** Script `auto_report.bat` lọc các dòng chứa `CRITICAL`, `ERROR` hoặc `TIMEOUT` bằng Pipe.
- **Output:** Tệp `bao-cao-su-co.txt` trong thư mục `backup/` để phục vụ báo cáo.

### 2.3. Mô hình DIKW (Data – Information – Knowledge – Wisdom)

- **Data (Dữ liệu):** Hàng triệu dòng log thô, chưa được phân tích.
- **Information (Thông tin):** Các dòng log đã lọc theo từ khóa như `ERROR 404`, `TIMEOUT`.
- **Knowledge (Tri thức):** Nhận biết các mẫu lỗi lặp lại và xác định trạm có tần suất sự cố cao.
- **Wisdom (Trí tuệ):** Đưa ra quyết định vận hành như nâng cấp thiết bị hoặc điều chỉnh cấu hình mạng.

## 3. Kiến trúc cây thư mục 4 trụ cột

### 3.1. Khái niệm

Một dự án hạ tầng chuyên nghiệp cần tách biệt bốn nhóm nội dung, tránh trộn lẫn mã nguồn, dữ liệu và tài liệu trong cùng một thư mục.

- `docs/` – **Tài liệu:** README, đặc tả kỹ thuật.
- `scripts/` – **Xử lý:** Các tệp tự động hóa `.bat` và `.sh`.
- `data/` – **Dữ liệu:** Log hệ thống, bảng đối soát thiết bị.
- `backup/` – **Sao lưu:** Các bản sao báo cáo phục vụ truy vết.

### 3.2. Thiết kế cây thư mục

```text
greentech-system/
├── docs/
│   ├── README.md
│   └── srs-he-thong.docx
├── scripts/
│   └── auto_report.bat
├── data/
│   ├── log-he-thong.log
│   └── quan-ly-thiet-bi.xlsx
├── backup/
│   └── bao-cao-su-co.txt
└── .env
```

### 3.3. Quy tắc đặt tên

- Sử dụng `kebab-case` cho tên thư mục và tệp, ví dụ: `log-he-thong.log`.
- Không sử dụng tiếng Việt có dấu.
- Không sử dụng khoảng trắng trong tên thư mục và tệp.
- Tệp cấu hình nhạy cảm `.env` cần được bảo vệ và không được commit lên GitHub công khai.

## 4. Thiết kế cơ sở dữ liệu Excel (Asset Database)

### 4.1. Khái niệm

Bảng tính Excel cần có trường định danh duy nhất ở bảng danh mục và trường tham chiếu ở bảng còn lại để hạn chế trùng lặp dữ liệu.

- **Bảng cha:** `danh-muc-linh-kien` – lưu thông tin gốc của từng loại linh kiện.
- **Bảng con:** `trang-thai-tram` – lưu trạng thái từng trạm và mã linh kiện được sử dụng.

### 4.2. Thiết kế bảng

**Sheet `danh-muc-linh-kien`**

| Trường | Ý nghĩa |
|---|---|
| `ma_linh_kien` | Khóa chính |
| `ten_thiet_bi` | Tên thiết bị |
| `don_gia` | Đơn giá |
| `loai_storage` | Loại lưu trữ SSD/HDD |

**Sheet `trang-thai-tram`**

| Trường | Ý nghĩa |
|---|---|
| `ma_tram` | Mã trạm |
| `ma_linh_kien` | Khóa ngoại |
| `dung_luong_ram` | Dung lượng RAM |
| `ngay_lap_dat` | Ngày lắp đặt |

**Quan hệ dữ liệu:**

`trang-thai-tram.ma_linh_kien` → `danh-muc-linh-kien.ma_linh_kien`

### 4.3. Quy tắc chính

- Khóa ngoại `ma_linh_kien` đặt ở bảng phía nhiều là `trang-thai-tram`.
- Không để trống `ma_linh_kien` khi trạm đã được lắp thiết bị.
- Vùng dữ liệu cần liền mạch, không có hàng trống hoặc ô gộp (Merge Cells).
- Mã linh kiện cần được chuẩn hóa, loại bỏ khoảng trắng thừa trước khi đối chiếu.

# PHẦN 2. XÂY DỰNG CÁC MODULE TỰ ĐỘNG HÓA

## 5. Bối cảnh

Xây dựng hệ thống quản trị hạ tầng gồm bốn module chức năng, tương ứng với các trụ cột của cây thư mục.

| Module | Công cụ chính | Thư mục liên quan |
|---|---|---|
| Khởi tạo cấu trúc dự án | Terminal / CMD | Toàn bộ cây thư mục |
| Robot lọc log tự động | Shell Script, Pipe, `findstr` / `grep` | `scripts/`, `data/`, `backup/` |
| Tài liệu hướng dẫn vận hành | Markdown, Word Styles | `docs/` |
| Đối soát tài nguyên | Excel VLOOKUP, TRIM | `data/` |

## 6. Module 1: Khởi tạo cấu trúc dự án

### 6.1. Chức năng

Tạo cây thư mục dự án bằng các lệnh CMD, hạn chế thao tác thủ công và bảo đảm cấu trúc thống nhất.

### 6.2. Quy trình xử lý

1. Xác định vị trí cần tạo thư mục gốc (Root).
2. Sử dụng lệnh `mkdir` để tạo các thư mục con.
3. Kiểm tra cây thư mục bằng lệnh `tree`.
4. Di chuyển vào thư mục gốc bằng đường dẫn tương đối để tạo tệp `.env`.
5. Kiểm tra `.env` và bảo đảm tệp này không bị đẩy lên GitHub công khai.

### 6.3. Lệnh mẫu trên Windows CMD

```bat
mkdir greentech-system
mkdir greentech-system\docs greentech-system\scripts greentech-system\data greentech-system\backup

cd greentech-system
type nul > .env

tree /F
```

**Giải thích:**

- `mkdir`: Tạo thư mục.
- `cd`: Di chuyển đến thư mục cần làm việc.
- `type nul > .env`: Tạo tệp `.env` rỗng nếu tệp chưa tồn tại.
- `tree /F`: Hiển thị cây thư mục cùng các tệp.

### 6.4. Lưu ý

- Không dùng tiếng Việt có dấu hoặc khoảng trắng trong tên thư mục.
- Tránh cố định đường dẫn tuyệt đối như `C:\Users\Ten` để tăng khả năng chạy trên máy khác.
- Tệp `.env` không tự động trở thành tệp ẩn trên Windows chỉ vì tên bắt đầu bằng dấu chấm. Có thể đánh dấu ẩn bằng lệnh:

```bat
attrib +h .env
```

- Thêm `.env` vào `.gitignore` để tránh đưa tệp lên Git:

```gitignore
.env
```

## 7. Module 2: Robot lọc log tự động

**Script:** `scripts/auto_report.bat`

### 7.1. Chức năng

Tự động đọc tệp `log-he-thong.log`, có thể lên đến 500 MB, lọc các dòng nghiêm trọng và nối kết quả vào báo cáo mà không cần mở toàn bộ tệp bằng trình soạn thảo.

### 7.2. Quy trình xử lý

1. In tiêu đề báo cáo và thời gian chạy.
2. Dùng `findstr` trên Windows hoặc `grep` trên Linux/macOS để tìm các dòng chứa từ khóa cần thiết.
3. Dùng toán tử Pipe (`|`) để chuyển dữ liệu từ lệnh đọc sang lệnh lọc.
4. Dùng toán tử `>>` để nối kết quả vào cuối báo cáo, tránh ghi đè dữ liệu cũ.
5. Lưu báo cáo trong thư mục `backup/`.

### 7.3. Script mẫu `auto_report.bat`

```bat
@echo off

echo ===== BAO CAO SU CO HE THONG ===== >> backup\bao-cao-su-co.txt
echo Thoi gian chay: %date% %time% >> backup\bao-cao-su-co.txt

type data\log-he-thong.log | findstr /I "CRITICAL ERROR TIMEOUT" >> backup\bao-cao-su-co.txt

echo Hoan tat loc log.
```

**Lưu ý:** Cần chạy script từ thư mục gốc `greentech-system` để các đường dẫn tương đối hoạt động đúng.

### 7.4. Output mẫu `backup/bao-cao-su-co.txt`

```text
===== BAO CAO SU CO HE THONG =====
Thoi gian chay: Mon 08/03/2026 10:00:00

[10:00:12] ERROR 404 - tram-cam-bien-27 khong phan hoi
[10:00:45] CRITICAL - mat ket noi may chu trung tam
[10:01:03] TIMEOUT - tram-cam-bien-58 qua thoi gian cho
```

### 7.5. Lưu ý và bẫy thường gặp

- Dùng `>` thay vì `>>` sẽ ghi đè nội dung cũ trong tệp đích.
- Không nên mở tệp log lớn bằng trình soạn thảo thông thường nếu máy không đủ tài nguyên.
- Với tệp 500 MB, có thể dùng `findstr` trực tiếp trên tệp để tránh tạo thêm luồng trung gian bằng `type`:

```bat
findstr /I "CRITICAL ERROR TIMEOUT" data\log-he-thong.log >> backup\bao-cao-su-co.txt
```

- Cần kiểm tra đường dẫn và thẩm định (Audit) trước khi chạy lệnh xóa do AI đề xuất. Không chạy `del /f /s /q` khi chưa xác nhận chính xác phạm vi tác động.

## 8. Module 3: Tài liệu hướng dẫn vận hành

### 8.1. Chức năng

Chuẩn hóa tài liệu nội bộ để nhân viên mới có thể vận hành hệ thống mà không cần hỏi lại kỹ sư phụ trách.

### 8.2. Quy trình xử lý README.md

1. Đặt tên hệ thống ở cấp Heading 1 bằng cú pháp `#`.
2. Viết hướng dẫn chạy script trong khối code.
3. Tạo checklist kiểm tra thiết bị bằng Task List.
4. Trình bày thông số phần cứng bằng bảng Markdown.

### 8.3. README.md mẫu

````markdown
# He thong quan tri ha tang GreenTech K26

## Cach chay script loc log

Mo CMD tai thu muc goc du an va chay:

```bat
scripts\auto_report.bat
```

## Checklist kiem tra thiet bi

- [ ] Kiem tra ket noi tram cam bien
- [ ] Kiem tra dung luong RAM con trong
- [ ] Sao luu bao cao vao backup/

## Thong so tram

| Tram | RAM | Ngay lap dat |
|---|---|---|
| tram-01 | 4GB | 2026-01-10 |
````

### 8.4. Quy trình xử lý SRS Document bằng Word

1. Gán Heading Styles (Heading 1, Heading 2, Heading 3) cho các tiêu đề.
2. Chèn mục lục tự động (Auto TOC) ngay sau trang bìa.
3. Kiểm tra mục lục và cập nhật số trang trước khi xuất bản.
4. Kiểm tra định dạng, đánh số mục và tính nhất quán của tài liệu.

## 9. Module 4: Đối soát tài nguyên bằng Excel

### 9.1. Chức năng

Tự động điền tên thiết bị vào bảng trạng thái trạm dựa trên mã linh kiện, hạn chế sai lệch do nhập liệu thủ công.

### 9.2. Quy trình xử lý

1. Làm sạch dữ liệu bằng hàm `TRIM` để loại bỏ khoảng trắng thừa.
2. Xác định vùng tra cứu trong bảng `danh-muc-linh-kien`.
3. Cố định vùng dữ liệu bằng dấu `$` để công thức không bị lệch khi sao chép.
4. Dùng `VLOOKUP` với tham số `FALSE` để tìm kết quả khớp chính xác.
5. Kiểm tra lỗi `#N/A` và xử lý mã không tồn tại hoặc chưa được chuẩn hóa.
6. Dùng AutoFilter để kiểm tra dữ liệu theo từng trạm.

### 9.3. Công thức mẫu

**Loại bỏ khoảng trắng thừa:**

```excel
=TRIM(A2)
```

**Tra cứu tên thiết bị theo mã linh kiện:**

```excel
=VLOOKUP(TRIM(B2),'danh-muc-linh-kien'!$A$2:$D$100,2,FALSE)
```

**Giải thích:**

- `TRIM(B2)`: Loại bỏ khoảng trắng thừa ở mã cần tra cứu.
- `'danh-muc-linh-kien'!$A$2:$D$100`: Vùng dữ liệu tra cứu.
- `2`: Trả về giá trị ở cột thứ hai trong vùng tra cứu.
- `FALSE`: Yêu cầu khớp chính xác.

**Lưu ý:** Công thức trên giả định mã linh kiện nằm ở cột A và tên thiết bị nằm ở cột B của sheet `danh-muc-linh-kien`.

### 9.4. Lưu ý và bẫy thường gặp

- Hàng trống hoặc ô Merge Cell có thể gây khó khăn khi lọc và xử lý dữ liệu.
- Nếu không cố định vùng tra cứu bằng `$`, vùng dữ liệu có thể bị dịch chuyển khi sao chép công thức.
- Hàm `TRIM` giúp xử lý khoảng trắng thông thường nhưng không loại bỏ mọi ký tự trắng đặc biệt.
- Nếu xuất hiện lỗi `#N/A`, cần kiểm tra mã linh kiện, dữ liệu nguồn và định dạng của các ô.

## 10. Các bẫy dữ liệu cần xử lý

| Bẫy | Rủi ro nếu không xử lý | Cách phòng tránh |
|---|---|---|
| Tên thư mục có dấu hoặc khoảng trắng | Script và đường dẫn có thể gặp lỗi | Đặt tên theo `kebab-case`, không dấu |
| Lệnh xóa nguy hiểm (`del /f /s /q`) | Xóa nhầm dữ liệu dự án | Thẩm định đường dẫn và phạm vi trước khi thực thi |
| Hàng trống hoặc Merge Cell trong Excel | AutoFilter và việc xử lý dữ liệu có thể sai lệch | Chuẩn hóa vùng dữ liệu liền mạch |
| Khoảng trắng thừa trong mã linh kiện | VLOOKUP trả về lỗi `#N/A` | Dùng `TRIM` và kiểm tra dữ liệu |
| Dùng `>` thay vì `>>` khi ghi báo cáo | Báo cáo cũ bị ghi đè | Dùng `>>` khi cần nối thêm dữ liệu |
| Tính sai đơn vị tốc độ mạng | Ước lượng sai thời gian tải log | Quy đổi Mbps sang MB/s bằng cách chia cho 8 |

---


