# Serial To WinForms - Hướng dẫn sử dụng

## File EXE đã build
File: `dist/SerialToWinForms.exe` (~24 MB)

## Cài đặt trên máy mới

### Bước 1: Copy các file cần thiết
Copy các file sau sang máy mới:
- `SerialToWinForms.exe` (từ thư mục dist)
- `config.json` (nếu muốn giữ cấu hình cũ)

### Bước 2: Không cần cài đặt gì thêm!
File EXE đã bao gồm tất cả thư viện cần thiết:
- ✅ Python runtime
- ✅ pyserial (serial communication)
- ✅ pywinauto (UI automation)
- ✅ pystray (system tray)
- ✅ PIL/Pillow (icons)
- ✅ tkinter (GUI)
- ✅ Tất cả dependencies khác

### Bước 3: Chạy chương trình
1. Double-click `SerialToWinForms.exe`
2. Cửa sổ GUI sẽ mở ra
3. Cấu hình các thông số:
   - **COM Port**: Port nhận dữ liệu serial (VD: COM8)
   - **Baudrate**: Tốc độ truyền (VD: 9600)
   - **Target App**: Tên cửa sổ ứng dụng đích (VD: Shop-Flow System From Indonesia(Pack))
   - **Textbox ID**: Auto ID của textbox đích (VD: GIFTBOX_AUTO)
4. Click **Save Config** để lưu cấu hình
5. Click **Start** để bắt đầu

## Các tính năng

### Status Monitoring
- 🔴 **Red dot**: Disconnected
- 🟢 **Green dot**: Connected
- Hiển thị trạng thái Serial Port và Shop-Flow connection

### System Tray
- Click nút X → ứng dụng ẩn xuống system tray (không thoát)
- Click icon tray → hiện lại cửa sổ
- Right-click icon tray:
  - **Show**: Hiện cửa sổ
  - **Hide**: Ẩn xuống tray
  - **Start/Stop**: Bật/tắt xử lý
  - **Exit**: Thoát hoàn toàn

### Activity Log
- Hiển thị toàn bộ hoạt động real-time
- Màu sắc:
  - Đen: Thông tin thường
  - Xanh lá: Thành công
  - Đỏ: Lỗi
  - Cam: Cảnh báo

### Counters
- **Data Received**: Tổng số lần nhận dữ liệu từ serial
- **Success**: Số lần gửi thành công vào Shop-Flow
- **Errors**: Số lỗi xảy ra

## Yêu cầu hệ thống

- Windows 7/8/10/11 (64-bit)
- Không cần cài Python
- Không cần cài thêm thư viện
- Cần có Shop-Flow app đang chạy (nếu muốn gửi dữ liệu)
- Cần có COM port (thật hoặc virtual) để nhận dữ liệu

## Khắc phục sự cố

### Lỗi: "Port COM8 not found"
- Kiểm tra COM port có tồn tại không (Device Manager)
- Thử đổi sang port khác trong Settings

### Lỗi: "Target window not found"
- Đảm bảo Shop-Flow app đang chạy
- Kiểm tra tên cửa sổ trong Target App có đúng không
- Tên cửa sổ phải khớp chính xác (có thể xem trong Task Manager)

### Lỗi: "Textbox not found"
- Kiểm tra Textbox ID có đúng không
- Có thể dùng tool như UISpy để tìm Auto ID chính xác

## File log

Chương trình tự động tạo log file trong thư mục `log/`:
- Format: `YYYY-MM-DD.txt`
- VD: `log/2025-11-14.txt`

Nếu có lỗi, kiểm tra file log để xem chi tiết.

## Liên hệ & hỗ trợ

Nếu gặp vấn đề, hãy kiểm tra:
1. File log trong thư mục `log/`
2. Config file `config.json` có đúng format không
3. Shop-Flow app có đang chạy không
4. COM port có hoạt động không

---
**Built with PyInstaller - Standalone executable, no dependencies required!**
