# Malayalam Church Songs - PPT Generator

Automated PowerPoint generation tool for Mar Thoma Syrian Church Malayalam Holy Communion Services.

## 🚀 **Quick Start for Users**

### Windows Executable 🖥️
**Standalone Windows app** - No browser, no internet needed:

📥 **[Download Windows .exe](WINDOWS_EXE_GUIDE.md)** - Double-click and go!

📖 **[Windows Guide](WINDOWS_EXE_GUIDE.md)** - Installation and usage instructions

---

## 📁 Folder Structure

```
Malayalam_Church_Songs/
├── Malayalam/                         # Malayalam scripts and files
│   ├── generate_malayalam_hcs_ppt.py # Main Malayalam PPT generator
│   ├── extract_malayalam_hymns.py    # Hymn extraction to Excel
│   ├── test_all_hymns.py             # Batch hymn testing
│   ├── Hymns_malayalam_KK.pptx       # Malayalam hymns source
│   └── kk_hymn_mapping.json          # Hymn number to title mapping
├── English/                           # English scripts (future)
│   └── extract_english_hymns.py
├── Windows_Exe/                       # Windows executable files
│   ├── gui_ppt_generator.py          # GUI application
│   ├── build_exe.py                  # Build script
│   └── build_windows_exe.bat         # Windows build batch
├── images/                            # Holy Communion images, QR codes
└── onedrive_git_local/                # All PPT files (downloaded from git at build time)
    └── Holy Communion Services - Slides/
        └── Malayalam HCS/
            ├── 2024 - Mal/
            ├── 2025-Mal/
            ├── 2026- Mal/
            └── Hymns_malayalam_KK.pptx  # Malayalam hymns master file
```

---

## Features

- 🎵 **Automated Slide Generation**: Generate complete service presentations from hymn numbers
- 📝 **Summary Slide**: Auto-generates summary with extracted 2-3 word titles
- 🔍 **Smart Search**: Finds hymns across multiple source PPT files
- 🖼️ **Image Handling**: Adds Holy Communion images, QR codes for Offertory
- 📊 **Hymn Extraction**: Extract all hymns from existing PPTs to Excel report
- 🎯 **Multi-Section Support**: Opening, Thanksgiving, Offertory, Confession, Communion, Closing
- 📂 **Flexible Paths**: User-provided folder or automatic fallback to onedrive_git_local

## Scripts

### 1. `generate_malayalam_hcs_ppt.py` (in Malayalam folder)
Main script to generate Malayalam Holy Communion Service presentations.

**Usage:**
```bash
cd Malayalam
python3 generate_malayalam_hcs_ppt.py --batch service_file.txt "Output_Name.pptx"
```

**Service File Format:**
```
# Date: 8 Feb 2026
313|Opening|
314|B/A|
315|Offertory|
316|Confession|
331|Communion|
343|Communion|
|Closing|yeshuveppole aakuvaan
```

**Path Resolution:**
1. Uses user-provided source folder (if specified in GUI)
2. Falls back to onedrive_git_local folder (bundled with exe or downloaded from git)

**Features:**
- Supports hymns with numbers (e.g., `313`) and without numbers (title-only)
- Automatically extracts 2-3 word titles from Manglish lyrics
- Malayalam-only service generation
- Filters out image-only slides
- Creates proper title slides for each section

### 2. `extract_malayalam_hymns.py` (in Malayalam folder)
Extract all hymn information from existing PPT files to Excel report with 3 tabs.

**Usage:**
```bash
cd Malayalam
python3 extract_malayalam_hymns.py
```

**Output:**
- Creates `malayalam_hymns_report.xlsx` with 3 tabs:
  - **Tab 1**: Sort by Hymn Number
  - **Tab 2**: Sort by Filename
  - **Tab 3**: KK Hymn Mapping (from kk_hymn_mapping.json)
- Lists all hymns found with hymn numbers, titles, and source files
- Supports both numbered hymns and title-only entries

### 3. `test_all_hymns.py` (in Malayalam folder)
Batch testing tool to verify all hymns can be generated.

**Usage:**
```bash
cd Malayalam
python3 test_all_hymns.py
```

**Features:**
- Reads hymns from `malayalam_hymns_report.xlsx`
- Tests each unique hymn number and title-only entry
- Creates individual test PPTs for each hymn
- Reports success/failure for all hymns

## Windows Executable Build

The Windows executable is built from files in the `Windows_Exe/` folder:

**To build:**
```bash
cd Windows_Exe
python build_exe.py
```

The executable will be created in `Windows_Exe/dist/Church_Songs_Generator.exe`

**What's included:**
- GUI application (`gui_ppt_generator.py`)
- Malayalam generator (`generate_malayalam_hcs_ppt.py` from Malayalam folder)
- Images folder (Holy Communion, QR codes)

## Requirements

```bash
pip install python-pptx pandas openpyxl
```

## Path Resolution (Important!)

### For Windows Executable Users
The executable includes all PPT files at build time - **no download needed!**

- ✅ Fully self-contained - works offline immediately
- ✅ No git or internet required on user machines
- ✅ Just run the exe and generate presentations

Users can optionally browse to their own OneDrive folder for latest files.

### For Script Users (Running Python Directly)
The script automatically downloads and caches PPT files:

1. **First Run**: Automatically downloads `onedrive_git_local` folder from git repository
2. **Subsequent Runs**: Uses the cached local copy (no download needed)
3. **Manual Setup** (optional): You can manually download and place files in `onedrive_git_local` folder

### Configuration (for script users and builders)
Update `GIT_REPO_URL` in `Malayalam/generate_malayalam_hcs_ppt.py` with your repository URL:

**For public repositories:**
```python
GIT_REPO_URL = "https://github.com/yourusername/Malayalam_Church_Songs.git"
GIT_BRANCH = "main"
```

**For private repositories (requires authentication):**
```python
# Option 1: Use Personal Access Token (PAT)
GIT_REPO_URL = "https://YOUR_TOKEN@github.com/yourusername/Malayalam_Church_Songs.git"
GIT_BRANCH = "main"

# Option 2: Disable auto-download (manual setup only)
GIT_REPO_URL = None
```

**To create a GitHub Personal Access Token:**
1. Go to GitHub → Settings → Developer settings → Personal access tokens
2. Generate new token (classic) with `repo` scope
3. Copy the token and use it in the URL above

### Search Order
The application searches for PPT files in this order:

1. **User-provided folder** (via GUI Browse button or command-line working directory)
2. **onedrive_git_local** (bundled in exe or downloaded from git)
3. **BASE_DIR** (script location - for development)

## Directory Structure

```
Church_Songs/
├── Malayalam/
│   ├── generate_malayalam_hcs_ppt.py
│   ├── extract_malayalam_hymns.py
│   ├── test_all_hymns.py
│   ├── Hymns_malayalam_KK.pptx
│   └── kk_hymn_mapping.json
├── English/
│   └── extract_english_hymns.py
├── Windows_Exe/
│   ├── gui_ppt_generator.py
│   ├── build_exe.py
│   └── build_windows_exe.bat
├── images/                       # Holy Communion images, QR codes
├── onedrive_git_local/           # All PPT files (in git repo)
│   └── Holy Communion Services - Slides/
│       ├── Malayalam HCS/
│       │   ├── 2024 - Mal/
│       │   ├── 2025-Mal/
│       │   ├── 2026- Mal/
│       │   └── Hymns_malayalam_KK.pptx  # Malayalam hymns master file
│       └── English HCS/
│           ├── 2024 - Eng HCS/
│           ├── 2025 - Eng HCS/
│           └── 2026 - Eng HCS/
└── test_hymns_output/            # Generated test files
```

## Configuration

Edit the paths in `generate_hcs_ppt.py`:
- `TEMPLATE_PPT`: Template presentation file
- `SEARCH_DIRS`: Directories to search for source PPTs
- `HOLY_COMMUNION_IMAGE`: Path to HC image
- `QR_CODE_IMAGE`: Path to offertory QR code

## Recent Updates

### Latest Features (Feb 2026)
- ✅ Support for hymns without hymn numbers (title-only entries)
- ✅ Fixed 2-digit hymn title extraction (17-91)
- ✅ Image-only slide filtering (skip HC intro images from source)
- ✅ Title slides show "Hymn - [Title]" for title-only songs
- ✅ Summary displays titles even when hymn number is empty
- ✅ Fixed BytesIO image copying

### Title Extraction
- Extracts first 2-3 words from Manglish lyrics
- Handles vertical tab characters (`\x0b`)
- Filters out corrupted Malayalam text
- Updates summary slide automatically

## Example Output

**Summary Slide:**
```
Opening:       313 manassode shaapa-maraththil thoongiya
B/A:           314 ponneshu nararh thiru
Offertory:     315 sneha virunna-nubhavippan (2)
Confession:    316 mannaaye bhujikka! jeeva
Communion:     331 swargga raaja puthrare
               343 Yaahenna Daivam Ennidayanaho
Closing:       yeshuveppole aakuvaan
```

---

## 🏗️ Building Windows Executable (For Developers)

Want to create a standalone .exe file for church members who prefer a desktop app?

### Prerequisites
- Python 3.8 or higher
- Windows OS (for building Windows executables)

### Build Steps

#### Option 1: Automated Build (Recommended)
```bash
# Clone the repository
git clone https://github.com/jthampan/Malayalam_Church_Songs.git
cd Malayalam_Church_Songs

# Run the Windows batch file
build_windows_exe.bat
```
The executable will be created in `dist/Church_Songs_Generator.exe`

#### Option 2: Manual Build
```bash
# Install dependencies
pip install pyinstaller python-pptx pandas openpyxl

# Run the build script
python build_exe.py
```

### What Gets Built
- **File:** `Church_Songs_Generator.exe`
- **Size:** ~15-20 MB (includes Python runtime and all dependencies)
- **Portable:** No installation required, runs on any Windows computer
- **Features:**
  - Simple GUI with browse buttons
  - One-time source folder setup (remembered)
  - Real-time progress and output log
  - Auto-saves to Desktop
  - Opens PowerPoint when done

### Distribution
1. Upload `dist/Church_Songs_Generator.exe` to GitHub Releases
2. Church members download the .exe
3. Double-click to run - no Python needed!
4. See [WINDOWS_EXE_GUIDE.md](WINDOWS_EXE_GUIDE.md) for user instructions

### GUI Features
- **Source PPT Folder**: One-time setup, path is saved
- **Service File Selection**: File browser
- **Generate Button**: Large, prominent, easy to click
- **Progress Bar**: Visual feedback during generation
- **Output Log**: Shows all messages and errors
- **Auto-open**: Option to open PowerPoint immediately

---

## License

For Mar Thoma Syrian Church use.

## Author

Joby Thampan
