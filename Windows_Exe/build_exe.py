# Build Script for Windows Executable
# Creates a standalone .exe file for Malayalam Church Songs PPT Generator

import PyInstaller.__main__
import os
import sys
import shutil
import subprocess
from pathlib import Path

# Configuration
GIT_REPO_URL = "https://github.com/jthampan/Malayalam_Church_Songs.git"
GIT_BRANCH = "main"

# Ensure we're in the correct directory
script_dir = Path(__file__).parent.absolute()
parent_dir = script_dir.parent
os.chdir(script_dir)

print(f"Working directory: {os.getcwd()}")
print()

# Check if required files exist
if not os.path.exists('gui_ppt_generator.py'):
    print("❌ ERROR: gui_ppt_generator.py not found!")
    print("   Please make sure you're in the Windows_Exe directory.")
    sys.exit(1)

malayalam_script = parent_dir / "Malayalam" / "generate_malayalam_hcs_ppt.py"
if not malayalam_script.exists():
    print(f"❌ ERROR: generate_malayalam_hcs_ppt.py not found at {malayalam_script}!")
    print("   This file is required for the executable to work.")
    sys.exit(1)

images_dir = parent_dir / "images"
if not images_dir.exists():
    print(f"⚠️  WARNING: images folder not found at {images_dir}")
    print("   QR code and Holy Communion images may not be available.")

# Download onedrive_git_local if not present
onedrive_git_local = parent_dir / "onedrive_git_local"
print("\n📥 Checking for onedrive_git_local folder...")
if onedrive_git_local.exists() and (onedrive_git_local / "Holy Communion Services - Slides").exists():
    print(f"✅ Found existing onedrive_git_local at {onedrive_git_local}")
    # Also check for Hymns_malayalam_KK.pptx in Malayalam HCS folder
    hymns_file = onedrive_git_local / "Holy Communion Services - Slides" / "Malayalam HCS" / "Hymns_malayalam_KK.pptx"
    if not hymns_file.exists():
        print(f"⚠️  Hymns_malayalam_KK.pptx not found in Malayalam HCS folder")
        print(f"   Expected at: {hymns_file}")
else:
    print(f"⚠️  onedrive_git_local not found or incomplete")
    print(f"📥 Downloading from git repository...")
    print(f"🔗 Repository: {GIT_REPO_URL}")
    
    # Check if git is available
    try:
        subprocess.run(["git", "--version"], capture_output=True, check=True)
    except (subprocess.CalledProcessError, FileNotFoundError):
        print("\n❌ ERROR: Git not installed!")
        print("   Please install git or manually create onedrive_git_local folder")
        print("   Get git from: https://git-scm.com/downloads")
        sys.exit(1)
    
    # Clean up old incomplete download
    if onedrive_git_local.exists():
        shutil.rmtree(onedrive_git_local, ignore_errors=True)
    
    # Create temp directory for sparse checkout
    temp_clone = parent_dir / ".temp_build_clone"
    if temp_clone.exists():
        shutil.rmtree(temp_clone, ignore_errors=True)
    
    try:
        print("⏳ Cloning repository (this may take a few minutes)...")
        temp_clone.mkdir(exist_ok=True)
        
        # Initialize git with sparse checkout
        subprocess.run(["git", "init"], cwd=temp_clone, check=True, capture_output=True)
        subprocess.run(["git", "remote", "add", "origin", GIT_REPO_URL], cwd=temp_clone, check=True, capture_output=True)
        subprocess.run(["git", "config", "core.sparseCheckout", "true"], cwd=temp_clone, check=True, capture_output=True)
        
        # Only checkout onedrive_git_local folder
        sparse_file = temp_clone / ".git" / "info" / "sparse-checkout"
        sparse_file.parent.mkdir(parents=True, exist_ok=True)
        sparse_file.write_text("onedrive_git_local/\n")
        
        # Pull the folder
        result = subprocess.run(
            ["git", "pull", "origin", GIT_BRANCH],
            cwd=temp_clone,
            capture_output=True,
            timeout=600  # 10 minute timeout
        )
        
        if result.returncode != 0:
            error_msg = result.stderr.decode('utf-8', errors='ignore').lower()
            if 'authentication' in error_msg or '403' in error_msg or '401' in error_msg:
                print("\n❌ ERROR: Authentication required for private repository!")
                print("   Options:")
                print("   1. Make the repository public on GitHub")
                print("   2. Use Personal Access Token in GIT_REPO_URL")
                print("      Format: https://TOKEN@github.com/username/repo.git")
                print(f"   3. Manually create '{onedrive_git_local}' folder with PPT files")
            else:
                print(f"\n❌ ERROR: Git clone failed: {result.stderr.decode('utf-8', errors='ignore')}")
            shutil.rmtree(temp_clone, ignore_errors=True)
            sys.exit(1)
        
        # Move the downloaded folder
        downloaded = temp_clone / "onedrive_git_local"
        if downloaded.exists() and (downloaded / "Holy Communion Services - Slides").exists():
            shutil.move(str(downloaded), str(onedrive_git_local))
            print(f"✅ Successfully downloaded onedrive_git_local")
            
            # Verify Hymns file is present
            hymns_file = onedrive_git_local / "Holy Communion Services - Slides" / "Malayalam HCS" / "Hymns_malayalam_KK.pptx"
            if hymns_file.exists():
                print(f"✅ Hymns_malayalam_KK.pptx found in Malayalam HCS folder")
            else:
                print(f"⚠️  WARNING: Hymns_malayalam_KK.pptx not found")
                print(f"   Expected at: {hymns_file}")
                print(f"   Please ensure it's in the repository under:")
                print(f"   onedrive_git_local/Holy Communion Services - Slides/Malayalam HCS/")
            
            shutil.rmtree(temp_clone, ignore_errors=True)
        else:
            print("\n❌ ERROR: onedrive_git_local folder not found in repository!")
            print(f"   Please ensure the repository contains 'onedrive_git_local' folder")
            shutil.rmtree(temp_clone, ignore_errors=True)
            sys.exit(1)
            
    except subprocess.TimeoutExpired:
        print("\n❌ ERROR: Download timed out - repository may be too large or network is slow")
        print(f"   Please manually create '{onedrive_git_local}' folder with PPT files")
        shutil.rmtree(temp_clone, ignore_errors=True)
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ ERROR: Failed to download: {e}")
        print(f"   Please manually create '{onedrive_git_local}' folder with PPT files")
        shutil.rmtree(temp_clone, ignore_errors=True)
        sys.exit(1)

print()

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

# Build arguments - include the Malayalam generator from parent/Malayalam folder
PyInstaller.__main__.run([
    'gui_ppt_generator.py',           # Main script
    '--onefile',                       # Single executable
    '--windowed',                      # No console window
    '--name=Church_Songs_Generator',   # Executable name
    '--icon=NONE',                     # No icon (can add later)
    f'--add-data={malayalam_script};.',  # Include Malayalam generator
    f'--add-data={images_dir};images' if images_dir.exists() else '--',  # Include images folder
    f'--add-data={onedrive_git_local};onedrive_git_local' if onedrive_git_local.exists() else '--',  # Include PPT files
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
    
    # Check what was included
    included_items = []
    if malayalam_script.exists():
        included_items.append("✅ Malayalam generator script")
    if images_dir.exists():
        included_items.append("✅ Images folder (QR code, Holy Communion)")
    if onedrive_git_local.exists():
        included_items.append("✅ onedrive_git_local (all PPT files)")
        # Check for Hymns file
        hymns_file = onedrive_git_local / "Holy Communion Services - Slides" / "Malayalam HCS" / "Hymns_malayalam_KK.pptx"
        if hymns_file.exists():
            included_items.append("✅ Hymns_malayalam_KK.pptx (Malayalam hymns)")
        else:
            included_items.append("⚠️  Hymns_malayalam_KK.pptx (missing - please add to repo)")
    
    print("\n📋 Included in executable:")
    for item in included_items:
        print(f"   {item}")
    
    print("\n💡 Distribution Instructions:")
    print("   1. Copy 'dist/Church_Songs_Generator.exe' to any Windows computer")
    print("   2. No Python installation needed!")
    print("   3. No internet needed - all PPT files are included!")
    print("   4. Just double-click to run")
    print("\n📁 Users can optionally:")
    print("   - Browse to their own OneDrive folder for latest files")
    print("   - Exe works standalone with included PPT files")
else:
    print("\n❌ ERROR: Executable was not created!")
    print("   Check the output above for errors.")
    sys.exit(1)
