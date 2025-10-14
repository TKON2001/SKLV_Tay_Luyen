# Auto Tẩy Luyện Tool v1.0

Công cụ tự động hóa cho game với các tính năng:
- **Chụp ảnh màn hình**: Chụp lại các khu vực cụ thể trên màn hình để phân tích chỉ số
- **Nhận dạng ký tự quang học (OCR)**: Đọc các con số từ hình ảnh đã chụp
- **Điều khiển chuột**: Tự động nhấp vào các nút "Tẩy Luyện" và "Khóa" chỉ số
- **Giao diện đồ họa**: Cung cấp giao diện thân thiện để thiết lập và điều khiển

## Yêu cầu hệ thống

- Windows 10/11
- Python 3.7 trở lên
- Game chạy trên trình giả lập (LDPlayer, BlueStacks, Nox, v.v.)

## Bước 1: Cài đặt Tesseract-OCR

1. Truy cập trang [Tesseract at UB Mannheim](https://github.com/UB-Mannheim/tesseract/wiki)
2. Tải về file cài đặt cho Windows (ví dụ: `tesseract-ocr-w64-setup-v5.x.x.exe`)
3. **QUAN TRỌNG**: Trong quá trình cài đặt, hãy chọn "Add Tesseract to system PATH"
4. Ghi nhớ đường dẫn cài đặt (thường là `C:\Program Files\Tesseract-OCR`)

## Bước 2: Cài đặt các thư viện Python

Mở Command Prompt (CMD) hoặc PowerShell và chạy:

```bash
pip install -r requirements.txt
```

Hoặc cài đặt từng thư viện:

```bash
pip install pillow
pip install pyautogui
pip install pytesseract
pip install pygetwindow
pip install keyboard
```

## Bước 3: Chạy ứng dụng

```bash
python auto_tay_luyen.py
```

## Hướng dẫn sử dụng

### 1. Chọn Cửa Sổ Game
- Nhấn nút "Chọn Cửa Sổ"
- Nhập một phần tên của trình giả lập (ví dụ: LDPlayer, BlueStacks)
- Chọn cửa sổ từ danh sách

### 2. Thiết Lập Tọa Độ

#### Nút Tẩy Luyện:
- Nhấn "Thiết lập" bên cạnh "Nút Tẩy Luyện"
- Di chuyển chuột đến vị trí nút Tẩy Luyện trong game
- Nhấn phím **F8** để ghi nhận tọa độ

#### Các Chỉ Số (lặp lại cho cả 4 chỉ số):

**Đặt vùng đọc:**
- Nhấn "Đặt vùng"
- Di chuyển chuột đến góc trên-trái của vùng hiển thị số
- Nhấn **F8**
- Di chuyển chuột đến góc dưới-phải
- Nhấn **F8** lần nữa

**Đặt nút khóa:**
- Nhấn "Đặt nút"
- Di chuyển chuột đến nút khóa tương ứng
- Nhấn **F8**

### 3. Nhập Chỉ Số Mong Muốn
- Nhập con số tối thiểu bạn muốn đạt được cho mỗi chỉ số
- Ví dụ: nếu muốn Thân +2500 trở lên, nhập 2500

### 4. Điều Khiển
- **Bắt đầu**: Nhấn nút "Bắt Đầu" hoặc phím **F5**
- **Dừng lại**: Nhấn nút "Dừng Lại" hoặc phím **F6**
- **Test OCR**: Nhấn "Test OCR" để kiểm tra khả năng đọc chỉ số

## Tính năng

- ✅ **Tự động lưu cấu hình**: Tọa độ và thiết lập được lưu tự động
- ✅ **Hotkey**: Sử dụng F5/F6 để điều khiển nhanh
- ✅ **Log chi tiết**: Theo dõi quá trình hoạt động
- ✅ **Test OCR**: Kiểm tra độ chính xác đọc chỉ số
- ✅ **Giao diện thân thiện**: Dễ sử dụng và thiết lập

## Lưu ý quan trọng

### Độ chính xác OCR
- OCR không phải lúc nào cũng chính xác 100%
- Font chữ, màu nền, và độ phân giải có thể ảnh hưởng kết quả
- Sử dụng "Test OCR" để kiểm tra và điều chỉnh vùng đọc

### Tọa độ
- Đảm bảo cửa sổ game không bị di chuyển sau khi thiết lập
- Nếu thay đổi độ phân giải, cần thiết lập lại tọa độ

### Rủi ro
- Việc sử dụng công cụ tự động có thể vi phạm điều khoản dịch vụ
- Sử dụng có trách nhiệm và tự chịu rủi ro

## Khắc phục sự cố

### Lỗi "Không tìm thấy Tesseract-OCR"
1. Kiểm tra Tesseract đã được cài đặt chưa
2. Thêm Tesseract vào system PATH
3. Hoặc sửa đường dẫn trong code:
```python
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
```

### OCR đọc sai
1. Sử dụng "Test OCR" để kiểm tra
2. Điều chỉnh vùng đọc cho chính xác hơn
3. Kiểm tra ảnh debug được lưu (`debug_stat_X.png`)

### Cửa sổ không được chọn
1. Đảm bảo trình giả lập đang chạy
2. Thử nhập tên chính xác hơn
3. Kiểm tra cửa sổ có bị ẩn không

## Cấu trúc file

```
tool_tay/
├── auto_tay_luyen.py          # File chính của ứng dụng
├── requirements.txt           # Danh sách thư viện cần thiết
├── README.md                 # Hướng dẫn này
├── config_tay_luyen.json     # File cấu hình (tự tạo)
└── debug_stat_X.png          # Ảnh debug OCR (tự tạo)
```

## Tùy chỉnh

Bạn có thể tùy chỉnh các tham số trong code:
- Thời gian chờ giữa các lần click (`time.sleep`)
- Cấu hình OCR (`ocr_config`)
- Xử lý ảnh (`process_image_for_ocr`)

## Phiên bản

- **v1.0**: Phiên bản đầu tiên với đầy đủ tính năng cơ bản

## Hỗ trợ

Nếu gặp vấn đề, hãy kiểm tra:
1. File log trong ứng dụng
2. Ảnh debug OCR
3. Cấu hình tọa độ
4. Phiên bản Tesseract và Python
#
