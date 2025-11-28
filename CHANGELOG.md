# Changelog

## Version 1.0.3 - 2025-11-28

### 🐛 Bug Fixes
- **Fixed settings/config file save location for .exe version**
  - Trước đây: Khi chạy file .exe, các file `settings.json` và `config.json` được lưu vào thư mục temp của PyInstaller (không tìm thấy được)
  - Bây giờ: Files được lưu vào **cùng thư mục với file .exe**
  - Thêm function `get_app_directory()` để phát hiện đúng đường dẫn cho cả .py và .exe

### 📝 Technical Details
```python
def get_app_directory(self):
    """Get application directory (works for both .py and .exe)"""
    if getattr(sys, 'frozen', False):
        # Running as compiled exe
        return os.path.dirname(sys.executable)
    else:
        # Running as script
        return os.path.dirname(os.path.abspath(__file__))
```

### 📂 File Structure khi chạy .exe
```
C:\Serial_to_MES\
├── SerialToWinForms.exe    ← File executable
├── config.json             ← Tự động tạo khi nhấn "Save Config"
├── settings.json           ← Tự động tạo khi nhấn "Save Settings"
└── version.txt             ← File version
```

---

## Version 1.0.2 - 2025-11-27

### ✨ New Features
- **Beautiful UI redesign**
  - Settings Dialog với 2 tabs đẹp mắt
  - About Dialog chuyên nghiệp
  - Hover effects trên buttons
  - Icons cho mỗi setting

### ⚙️ Settings Management
- Tất cả hardcoded values giờ có thể cấu hình qua GUI
- Menu bar với File → Settings
- Tab 1: Update Settings (FTP, Program Directory)
- Tab 2: Monitoring Settings (Timeouts, Tolerances)
- Test FTP Connection button

---

## Version 1.0.1 - 2025-11-26

### 🎯 Initial Release
- Serial to WinForms communication
- Config management
- System tray support
- Activity logging
