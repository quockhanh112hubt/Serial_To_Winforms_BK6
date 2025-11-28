"""
Build ALL executables for Serial To WinForms
- SerialToWinForms.exe (main app)
- update_script.exe (updater)
"""

import subprocess
import sys
import os

print("="*70)
print("🔨 Building Serial To WinForms - ALL EXECUTABLES")
print("="*70)

# Build main app
print("\n📦 Step 1/2: Building SerialToWinForms.exe...")
print("-"*70)
result1 = subprocess.run([sys.executable, "build_exe.py"])

if result1.returncode != 0:
    print("\n❌ Failed to build SerialToWinForms.exe")
    sys.exit(1)

# Build update script
print("\n📦 Step 2/2: Building update_script.exe...")
print("-"*70)
result2 = subprocess.run([sys.executable, "build_update_script.py"])

if result2.returncode != 0:
    print("\n❌ Failed to build update_script.exe")
    sys.exit(1)

# Success
print("\n" + "="*70)
print("✅ ALL BUILDS COMPLETED SUCCESSFULLY!")
print("="*70)
print("\n📂 Output files:")
print("   1. dist/SerialToWinForms.exe  (Main application)")
print("   2. dist/update_script.exe      (Update tool)")
print("\n📝 Deployment:")
print("   Copy both files to: C:\\Serial_to_MES\\")
print("\n" + "="*70)
