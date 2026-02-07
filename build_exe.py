# Build Script for Windows Executable
# Creates a standalone .exe file for Malayalam Church Songs PPT Generator

import PyInstaller.__main__
import os
import shutil
from pathlib import Path

# Clean previous builds
if os.path.exists('build'):
    shutil.rmtree('build')
if os.path.exists('dist'):
    shutil.rmtree('dist')

print("🔨 Building Windows Executable...")
print("="*60)

# Build arguments
PyInstaller.__main__.run([
    'gui_ppt_generator.py',           # Main script
    '--onefile',                       # Single executable
    '--windowed',                      # No console window
    '--name=Church_Songs_Generator',   # Executable name
    '--icon=NONE',                     # No icon (can add later)
    '--add-data=generate_hcs_ppt.py;.',  # Include main generator
    '--clean',                         # Clean cache
    '--noconfirm',                     # Overwrite without asking
])

print("\n" + "="*60)
print("✅ Build complete!")
print("="*60)
print(f"\n📦 Executable created: dist/Church_Songs_Generator.exe")
print(f"   File size: {os.path.getsize('dist/Church_Songs_Generator.exe') / (1024*1024):.1f} MB")
print("\n💡 Distribution Instructions:")
print("   1. Copy 'dist/Church_Songs_Generator.exe' to any Windows computer")
print("   2. No Python installation needed!")
print("   3. Just double-click to run")
print("\n⚠️  NOTE: Users need to have their source PPT files ready")
print("   (Will be prompted on first run)")
