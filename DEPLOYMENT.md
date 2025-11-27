# 📦 Serial To WinForms - Deployment Guide

## 🚀 Quick Deployment

### Files cần thiết:
```
C:\Serial_to_MES\
├── SerialToWinForms.exe    (File chính)
├── config.json             (Cấu hình serial & WinForms - tùy chọn)
├── settings.json           (Cấu hình nâng cao - tùy chọn)
└── version.txt             (File version hiện tại)
```

### Cách deploy:

1. **Copy file exe**
   ```
   Copy: dist\SerialToWinForms.exe
   Đến:  C:\Serial_to_MES\SerialToWinForms.exe
   ```

2. **Tạo file config.json** (nếu chưa có):
   ```json
   {
       "port": "COM10",
       "baudrate": 9600,
       "target_app_title": "Shop-Flow System From Vietnam(Pack)",
       "textbox_auto_id": "GIFTBOX_AUTO",
       "backend": "win32"
   }
   ```

3. **Tạo file settings.json** (tùy chọn - dùng mặc định nếu không có):
   ```json
   {
       "program_directory": "C:\\Serial_to_MES",
       "ftp_server": "10.62.102.5",
       "ftp_user": "update",
       "ftp_password": "update",
       "ftp_directory": "KhanhDQ/Update_Program/Serial_to_MES/",
       "max_log_lines": 50,
       "idle_timeout_minutes": 30,
       "max_consecutive_errors": 10,
       "connection_grace_period": 5,
       "max_disconnect_tolerance": 20
   }
   ```

4. **Tạo file version.txt**:
   ```
   1.0.0
   ```

5. **Chạy ứng dụng**:
   - Double-click `SerialToWinForms.exe`
   - Hoặc tạo shortcut trên Desktop
   - Hoặc thêm vào Startup folder để chạy cùng Windows

---

## 🔧 Build lại từ source

### Yêu cầu:
- Python 3.13+
- Virtual environment đã cài đặt packages

### Các bước:

1. **Activate virtual environment**:
   ```powershell
   .\.venv\Scripts\Activate.ps1
   ```

2. **Chạy build script**:
   ```powershell
   python build_exe.py
   ```

3. **Tìm file exe**:
   ```
   dist\SerialToWinForms.exe
   ```

---

## ⚙️ Cấu hình sau khi cài đặt

### Qua giao diện:
1. Mở ứng dụng
2. Menu **File → Settings**
3. Cấu hình 2 tabs:
   - **📦 Update Settings**: FTP server, program directory
   - **📊 Monitoring Settings**: Timeout, error tolerance, etc.
4. Click **💾 Save Settings**

### Hoặc chỉnh sửa trực tiếp file JSON ở trên

---

## 🔄 Update chương trình

### Tự động (qua FTP):
- Chương trình tự check update từ FTP server
- Download và cài đặt tự động khi có version mới

### Thủ công:
1. Build version mới
2. Copy `SerialToWinForms.exe` mới
3. Overwrite file cũ
4. Giữ nguyên `config.json` và `settings.json`

---

## 📝 Notes

### Các tính năng chính:
- ✅ Giao diện GUI hiện đại với Menu bar
- ✅ Settings dialog đẹp mắt với 2 tabs
- ✅ About dialog chuyên nghiệp
- ✅ Tất cả tham số có thể cấu hình qua GUI
- ✅ Lưu settings vào file JSON
- ✅ System tray support
- ✅ Auto-update support (qua FTP)
- ✅ Logging chi tiết

### Troubleshooting:
- **Không kết nối được COM port**: Check COM port number và baudrate
- **Không tìm thấy Shop-Flow**: Check "Target App Title" trong settings
- **FTP update không hoạt động**: Check FTP settings trong Menu → Settings

---

## 👨‍💻 Developer Info

**Developed by**: KhanhIT - IT Team  
**Company**: ITM Semiconductor  
**Version**: 1.0.0  
**Date**: November 2025  

---

## 📞 Support

Nếu gặp vấn đề, liên hệ IT Team để được hỗ trợ.
