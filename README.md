# PDF Steel Schedule Extractor

Công cụ trích xuất bảng thép (steel schedule) từ file PDF và xuất ra Excel.

## Tính năng

- Quét vùng góc dưới-bên trái của bản vẽ PDF
- Xuất tất cả bảng được phát hiện trong vùng quét ra file Excel
- Chế độ debug vẽ các khung bảng được phát hiện lên file PNG
- Dễ dàng cấu hình thông qua biến CONFIG
- Hỗ trợ:
  - Xử lý một file PDF đơn lẻ
  - Xử lý toàn bộ thư mục chứa PDF

## Yêu cầu hệ thống

- Python 3.11+ (pandas 3.x yêu cầu)
- Các thư viện Python (xem `requirements.txt`):
  - pdfplumber
  - pandas
  - openpyxl (ghi file Excel)
  - PyMuPDF
  - Pillow (PIL)

## Cài đặt

```bash
pip install -r requirements.txt
```

## Cấu trúc thư mục

```text
Steel-Schedule-Finder/
├── Schedule.py          # File chính của chương trình
├── requirements.txt     # Danh sách thư viện
├── README.md            # Hướng dẫn sử dụng
├── context/             # Tài liệu kế hoạch / ghi chú (plan.md)
├── tests/               # Test tự động (unittest)
├── input/               # Thư mục chứa file PDF đầu vào (không đẩy lên git)
└── Run_YYYYMMDD_HHMMSS/ # Kết quả mỗi lần chạy, nằm trong "run_root" (không đẩy lên git)
    ├── output.xlsx      # File Excel kết quả
    └── debug/           # Các file debug PNG
```

## Cách sử dụng

### 1. Cấu hình

Mở file `Schedule.py` và điều chỉnh các tham số trong biến `CONFIG`:

```python
CONFIG = {
    # Chế độ input: "single" hoặc "folder"
    "input_mode": "folder",
    
    # Đường dẫn file PDF đơn lẻ (khi input_mode = "single")
    "input_pdf": r"D:\Working\Draft\input\SUB1-[0].pdf",
    
    # Đường dẫn thư mục chứa PDF (khi input_mode = "folder")
    "input_folder": r"D:\Working\Draft\input",
    
    # Tên file Excel đầu ra (lưu trong thư mục Run_* của mỗi lần chạy)
    "output_excel": "output.xlsx",
    
    # Thư mục gốc lưu các lần chạy
    "run_root": r"D:\Working\Draft",
    
    # Chế độ debug (ảnh debug lưu trong Run_*/debug/)
    "debug": True,
    "export_debug_images": True,
    
    # Vùng quét trên trang PDF (tỷ lệ 0.00-1.00)
    "search_region": {
        "left": 0.00,      # Từ trái: 0%
        "right": 0.40,     # Đến phải: 40% chiều rộng
        "top": 0.60,       # Từ trên: 60% chiều cao
        "bottom": 1.00,    # Đến dưới: 100% chiều cao
    },
    
    # Cài đặt trích xuất bảng
    "table_settings": {
        "vertical_strategy": "lines",
        "horizontal_strategy": "lines",
        "snap_tolerance": 5,
        "join_tolerance": 5,
        "intersection_tolerance": 8,
        "edge_min_length": 10,
    },
    
    # Bộ lọc cơ bản
    "min_rows": 3,      # Số dòng tối thiểu
    "min_cols": 2,      # Số cột tối thiểu
    "min_cells": 8,     # Số ô tối thiểu
    
    # Tên sheet trong Excel
    "sheet_name": "Steel Schedule",
    
    # Cài đặt chấm điểm bảng
    "scoring": {
        "left_weight": 40,        # Ưu tiên bảng sát góc trái
        "bottom_weight": 40,      # Ưu tiên bảng sát đáy
        "aspect_weight": 30,      # Ưu tiên bảng rộng
        "cell_weight": 20,        # Ưu tiên nhiều ô
        "text_weight": 20,        # Ưu tiên nhiều text
        "height_penalty": 50,     # Phạt bảng quá cao
        "min_score": 80,          # Điểm tối thiểu để giữ bảng
        "height_threshold": 0.20, # Ngưỡng phạt chiều cao (tỷ lệ chiều cao trang)
        "vertical_threshold": 1.2, # Ngưỡng phạt bảng dọc (tỷ lệ rộng/cao)
        "vertical_penalty": 100,  # Điểm phạt bảng dọc
    },
}
```

### 2. Điều chỉnh vùng quét (search_region)

Vùng quét được xác định bằng tỷ lệ (0.00-1.00) so với chiều rộng/chiều cao trang:

```python
"search_region": {
    "left": 0.00,    # 0% từ trái
    "right": 0.40,   # 40% chiều rộng
    "top": 0.60,     # 60% từ trên
    "bottom": 1.00,  # 100% chiều cao (đáy trang)
}
```

**Ví dụ:**

- `left: 0.00, right: 0.40` → Quét 40% chiều rộng từ trái sang
- `top: 0.60, bottom: 1.00` → Quét 40% chiều cao từ giữa xuống đáy

### 3. Chạy chương trình

```bash
python Schedule.py
```

### 4. Xem kết quả

- **File Excel**: Lưu trong thư mục `Run_YYYYMMDD_HHMMSS/output.xlsx`
- **File debug**: Lưu trong thư mục `Run_YYYYMMDD_HHMMSS/debug/` (nếu `export_debug_images = True`)

## Chế độ Debug

Khi bật chế độ debug (`debug: True`), chương trình sẽ:

1. In ra console các thông tin chi tiết về quá trình xử lý
2. Tạo các file PNG trong thư mục `debug/` với:
   - **Vùng đỏ**: Vùng quét được định nghĩa trong `search_region`
   - **Vùng xanh**: Các bảng được phát hiện
   - **Chữ T1, T2, ...**: Số thứ tự bảng

## Tùy chỉnh

### Điều chỉnh độ nhạy phát hiện bảng

Nếu bảng bị bỏ sót hoặc phát hiện sai, điều chỉnh các tham số trong `table_settings`:

```python
"table_settings": {
    "snap_tolerance": 5,           # Tăng nếu các đường kẻ không nối lại
    "join_tolerance": 5,           # Tăng nếu các đoạn đường kẻ bị tách
    "intersection_tolerance": 8,   # Tăng nếu giao điểm không được phát hiện
    "edge_min_length": 10,         # Tăng nếu phát hiện nhiều đường kẻ nhỏ không cần thiết
}
```

### Lọc bảng không mong muốn

Nếu xuất hiện nhiều bảng nhỏ không cần thiết, tăng các ngưỡng lọc:

```python
"min_rows": 3,      # Tăng lên 4-5
"min_cols": 2,      # Tăng lên 3-4
"min_cells": 8,     # Tăng lên 10-15
```

Hoặc tăng điểm tối thiểu:

```python
"min_score": 80,    # Tăng lên 100-120
```

### Điều chỉnh trọng số chấm điểm

Nếu bảng steel schedule không được ưu tiên đúng, điều chỉnh các trọng số trong `scoring`:

```python
"scoring": {
    "left_weight": 40,        # Tăng nếu bảng nằm sát góc trái
    "bottom_weight": 40,      # Tăng nếu bảng nằm sát đáy
    "aspect_weight": 30,      # Tăng nếu bảng rộng hơn là cao
    "cell_weight": 20,        # Tăng nếu bảng có nhiều ô
    "text_weight": 20,        # Tăng nếu bảng có nhiều text
}
```

## Xử lý sự cố

### Không tìm thấy bảng nào

1. Kiểm tra lại vùng quét `search_region` có đúng không
2. Xem file debug PNG để xem vùng quét và các bảng được phát hiện
3. Tăng `min_score` hoặc điều chỉnh `table_settings`

### Phát hiện nhiều bảng nhỏ không cần thiết

1. Tăng `min_rows`, `min_cols`, `min_cells`
2. Tăng `min_score`
3. Tăng `edge_min_length` trong `table_settings`

### Bảng bị cắt xén hoặc thiếu dữ liệu

1. Mở rộng vùng quét `search_region`
2. Giảm `min_rows`, `min_cols`, `min_cells`
3. Giảm `min_score`

## Chạy test

```bash
python -m unittest discover -s tests -t .
```

Test "smoke" dùng file `input/TRUSS LOT 73.pdf`. File này không nằm trong git; nếu không có, test đó sẽ được bỏ qua.

## Lưu ý

- Các thư mục `input/` và `Run_*/`, cùng mọi file `*.pdf`, không được đẩy lên git (đã cấu hình trong `.gitignore`)
- Trang PDF bị xoay (rotated) có thể làm khung debug bị lệch vị trí; kết quả Excel không bị ảnh hưởng nhiều nhưng nên kiểm tra lại bằng ảnh debug
- Mỗi lần chạy chương trình sẽ tạo một thư mục `Run_YYYYMMDD_HHMMSS` mới
- File Excel kết quả được lưu trong thư mục run tương ứng
- Nên bật chế độ debug khi điều chỉnh cấu hình để dễ dàng kiểm tra

## Ví dụ cấu hình

### Xử lý một file PDF đơn lẻ

```python
CONFIG = {
    "input_mode": "single",
    "input_pdf": r"D:\Working\Draft\input\drawing.pdf",
    "output_excel": "output.xlsx",
    # ... các cài đặt khác
}
```

### Xử lý toàn bộ thư mục PDF

```python
CONFIG = {
    "input_mode": "folder",
    "input_folder": r"D:\Working\Draft\input",
    "output_excel": "output.xlsx",
    # ... các cài đặt khác
}
```
