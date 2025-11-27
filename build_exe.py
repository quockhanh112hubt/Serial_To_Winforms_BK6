"""
Build script for Serial To WinForms application
Creates a standalone .exe file using PyInstaller
"""

import PyInstaller.__main__
import os

# Get the current directory
current_dir = os.path.dirname(os.path.abspath(__file__))

# Build parameters
PyInstaller.__main__.run([
    'serial_to_winforms_gui.py',  # Main script
    '--name=SerialToWinForms',    # Output name
    '--onefile',                   # Single executable file
    '--windowed',                  # No console window
    '--icon=app_icon.ico',        # Application icon
    '--add-data=app_icon.ico;.',  # Include icon in bundle
    '--hidden-import=PIL._tkinter_finder',  # Include PIL dependencies
    '--hidden-import=pystray._win32',       # Include pystray dependencies
    '--hidden-import=serial.tools.list_ports',  # Include pyserial tools
    '--clean',                     # Clean PyInstaller cache
    '--noconfirm',                # Replace output without asking
])

print("\n" + "="*60)
print("✅ Build completed successfully!")
print("="*60)
print(f"\n📦 Executable file: {os.path.join(current_dir, 'dist', 'SerialToWinForms.exe')}")
print("\n📝 Instructions:")
print("1. Copy 'SerialToWinForms.exe' from 'dist' folder")
print("2. Create 'config.json' in the same folder (or use existing)")
print("3. Create 'settings.json' in the same folder (optional)")
print("4. Run SerialToWinForms.exe")
print("\n" + "="*60)
