"""
Build script for update_script.exe
"""

import PyInstaller.__main__
import os

# Get the current directory
current_dir = os.path.dirname(os.path.abspath(__file__))

# Build parameters
PyInstaller.__main__.run([
    'update_script.py',            # Main script
    '--name=update_script',        # Output name
    '--onefile',                   # Single executable file
    '--windowed',                  # No console window
    '--icon=app_icon.ico',        # Application icon
    '--clean',                     # Clean PyInstaller cache
    '--noconfirm',                # Replace output without asking
])

print("\n" + "="*60)
print("✅ update_script.exe build completed!")
print("="*60)
print(f"\n📦 Executable: {os.path.join(current_dir, 'dist', 'update_script.exe')}")
print("\n" + "="*60)
