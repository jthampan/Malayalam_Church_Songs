# Malayalam Church Songs - PPT Generator

Automated PowerPoint generation tool for Mar Thoma Syrian Church Malayalam Holy Communion Services.

## 🚀 **Quick Start for Users**

### Windows Executable 🖥️
**Standalone Windows app** - No browser, no internet needed:

📥 **[Download Windows .exe](WINDOWS_EXE_GUIDE.md)** - Double-click and go!

📖 **[Windows Guide](WINDOWS_EXE_GUIDE.md)** - Installation and usage instructions

---

## Features

- 🎵 **Automated Slide Generation**: Generate complete service presentations from hymn numbers
- 📝 **Summary Slide**: Auto-generates summary with extracted 2-3 word titles
- 🔍 **Smart Search**: Finds hymns across multiple source PPT files
- 🖼️ **Image Handling**: Adds Holy Communion images, QR codes for Offertory
- 📊 **Hymn Extraction**: Extract all hymns from existing PPTs to Excel report
- 🎯 **Multi-Section Support**: Opening, Thanksgiving, Offertory, Confession, Communion, Closing

## Scripts

### 1. `generate_hcs_ppt.py`
Main script to generate Holy Communion Service presentations.

**Usage:**
```bash
python3 generate_hcs_ppt.py --batch service_file.txt "Output_Name.pptx"
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

**Features:**
- Supports hymns with numbers (e.g., `313`) and without numbers (title-only)
- Automatically extracts 2-3 word titles from Manglish lyrics
- Handles both Malayalam and English services
- Filters out image-only slides
- Creates proper title slides for each section

### 2. `extract_malayalam_hymns.py`
Extract all hymn information from existing PPT files to Excel report.

**Usage:**
```bash
python3 extract_malayalam_hymns.py
```

**Output:**
- Creates `malayalam_hymns_report.xlsx`
- Lists all hymns found with hymn numbers, titles, and source files
- Supports both numbered hymns and title-only entries

### 3. `test_all_hymns.py`
Batch testing tool to verify all hymns can be generated.

**Usage:**
```bash
python3 test_all_hymns.py
```

**Features:**
- Reads hymns from `malayalam_hymns_report.xlsx`
- Tests each unique hymn number and title-only entry
- Creates individual test PPTs for each hymn
- Reports success/failure for all hymns

## Requirements

```bash
pip install python-pptx pandas openpyxl
```

## Directory Structure

```
Church_Songs/
├── generate_hcs_ppt.py          # Main PPT generator
├── extract_malayalam_hymns.py   # Hymn extraction to Excel
├── test_all_hymns.py             # Batch hymn testing
├── service_8feb2026.txt          # Sample service file
├── images/                       # Holy Communion images, QR codes
├── OneDrive_2026-02-05/          # Source PPT files
│   └── Holy Communion Services - Slides/
│       └── Malayalam HCS/
│           ├── 2024 - Mal/
│           ├── 2025-Mal/
│           └── 2026- Mal/
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
