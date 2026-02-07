# Build Script for Windows Executable
# Creates a standalone .exe file for Malayalam Church Songs PPT Generator

import PyInstaller.__main__
import os
import sys
import shutil
import subprocess
from pathlib import Path

# Ensure we're in the correct directory
script_dir = Path(__file__).parent.absolute()
os.chdir(script_dir)

print(f"Working directory: {os.getcwd()}")
print()

# Check if required files exist
if not os.path.exists('gui_ppt_generator.py'):
    print("❌ ERROR: gui_ppt_generator.py not found!")
    print("   Please make sure you're in the Church_Songs directory.")
    sys.exit(1)

if not os.path.exists('generate_hcs_ppt.py'):
    print("❌ ERROR: generate_hcs_ppt.py not found!")
    print("   This file is required for the executable to work.")
    sys.exit(1)

# Clean previous builds
if os.path.exists('build'):
    print("🧹 Cleaning previous build directory...")
    shutil.rmtree('build')
if os.path.exists('dist'):
    print("🧹 Cleaning previous dist directory...")
    shutil.rmtree('dist')

# Ensure the old exe is not running (prevents WinError 5)
try:
    subprocess.run(
        ["taskkill", "/F", "/IM", "Church_Songs_Generator.exe"],
        check=False,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
except Exception:
    pass

print("🔨 Building Windows Executable...")
print("="*60)
print()

# Build arguments
PyInstaller.__main__.run([
    'gui_ppt_generator.py',           # Main script
    '--onefile',                       # Single executable
    '--windowed',                      # No console window
    '--name=Church_Songs_Generator',   # Executable name
    '--icon=NONE',                     # No icon (can add later)
    '--add-data=generate_hcs_ppt.py;.',  # Include main generator
    '--add-data=images;images',        # Include images folder (QR + Holy Communion)
    '--clean',                         # Clean cache
    '--noconfirm',                     # Overwrite without asking
])

print("\n" + "="*60)
print("✅ Build complete!")
print("="*60)

exe_path = 'dist/Church_Songs_Generator.exe'
if os.path.exists(exe_path):
    file_size_mb = os.path.getsize(exe_path) / (1024*1024)
    print(f"\n📦 Executable created: {exe_path}")
    print(f"   File size: {file_size_mb:.1f} MB")
    print("\n💡 Distribution Instructions:")
    print("   1. Copy 'dist/Church_Songs_Generator.exe' to any Windows computer")
    print("   2. No Python installation needed!")
    print("   3. Just double-click to run")
    print("\n⚠️  NOTE: Users need to have their source PPT files ready")
    print("   (Will be prompted on first run)")
else:
    print("\n❌ ERROR: Executable was not created!")
    print("   Check the output above for errors.")
    sys.exit(1)
