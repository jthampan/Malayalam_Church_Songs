# Windows Executable Version - User Guide

## 🖥️ Malayalam Church Songs PPT Generator for Windows

**Super Simple Windows Application - No Browser, No Internet Needed!**

---

## ✨ What You Get

A standalone Windows application (`.exe` file) that runs on any Windows computer without installing Python or any other software.

---

## 📥 Installation

### Option 1: Download Pre-Built Executable
1. Download `Church_Songs_Generator.exe` from the GitHub releases page
2. Save it anywhere on your computer (Desktop is fine)
3. **That's it!** No installation needed.

### Option 2: Build It Yourself
If you have Python installed:
```bash
# Clone the repository
git clone https://github.com/jthampan/Malayalam_Church_Songs.git
cd Malayalam_Church_Songs

# Run the build script
build_windows_exe.bat
```
The executable will be created in the `dist` folder.

---

## 🚀 How to Use (First Time)

### Step 1: Double-Click the Executable
- Find `Church_Songs_Generator.exe`
- Double-click to open
- A window will appear with the PPT generator

### Step 2: One-Time Setup (Source PPT Folder)
1. Click **"Browse..."** next to "Source PPT Folder"
2. Navigate to the folder containing your hymn PowerPoint files
   - Example: `OneDrive/Holy Communion Services - Slides/Malayalam HCS/2026- Mal`
3. Select the folder
4. Click "Select Folder"

**✅ Done!** This folder will be remembered for next time.

---

## 🎉 How to Use (Every Time)

### Step 1: Select Service File
1. Click **"Browse..."** next to "Service File"
2. Select your service text file (e.g., `service_8feb2026.txt`)

### Step 2: Generate PowerPoint
1. Click the big green button: **"🎵 GENERATE POWERPOINT"**
2. Wait 10-30 seconds while it generates
3. A success message will appear
4. Click "Yes" to open your PowerPoint immediately
5. **Your PPT is saved to your Desktop!**

That's it! 🎊

---

## 📄 Service File Format

Create a text file (use Notepad) with this format:

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

**Rules:**
- First line: `# Date: DD MMM YYYY`
- Each hymn: `hymn_number|Section_Label|optional_title`
- Use `|` (pipe character) as separator
- For songs without hymn numbers, leave first part empty: `|Closing|song title`

**Supported Sections:**
- Opening
- B/A (Thanksgiving)
- Offertory
- Confession
- Communion (can have multiple)
- Closing
- Dedication

---

## ❓ Frequently Asked Questions

### Q: Do I need internet to use this?
**A:** No! Once you have the .exe file and your source PPT files, you can use it completely offline.

### Q: Do I need to install Python?
**A:** No! The executable includes everything needed.

### Q: Where does it save the PowerPoint?
**A:** On your Desktop with a filename like: `HCS_Malayalam_07_Feb_2026.pptx`

### Q: Can I use this on Mac or Linux?
**A:** This .exe is Windows-only.

### Q: What if I get a "Windows protected your PC" warning?
**A:** This is normal for unsigned executables. Click "More info" then "Run anyway". The software is safe!

### Q: Can I share this .exe with other church members?
**A:** Yes! Just send them the .exe file. They'll need to do the one-time setup with their source PPT folder.

### Q: What if my source PPT files are on OneDrive?
**A:** Make sure OneDrive is synced to your computer, then select the local OneDrive folder during setup.

### Q: "Song not found" error?
**A:** 
- Check that the hymn number exists in your source PPT files
- Make sure you selected the correct source folder
- Verify the hymn number in your service file is correct

### Q: Can I generate multiple PPTs without restarting?
**A:** Yes! Just select a different service file and click Generate again.

---

## 🛠️ Troubleshooting

### Problem: Application won't start
- **Solution:** You may need to install [Microsoft Visual C++ Redistributable](https://aka.ms/vs/17/release/vc_redist.x64.exe)

### Problem: "Generator script not found" error
- **Solution:** Make sure `generate_hcs_ppt.py` is in the same folder as the .exe

### Problem: Output log shows "File not found" errors
- **Solution:** 
  1. Check your source PPT folder path is correct
  2. Ensure PPT files are named correctly (e.g., `313.pptx` or `Hymn 313.pptx`)
  3. Make sure OneDrive files are downloaded (not cloud-only)

### Problem: PowerPoint doesn't open automatically
- **Solution:** The file is still saved to your Desktop. Navigate there and open it manually.

---

## 📧 Support

**Need help?** Contact: joby.thampan@gmail.com

**GitHub Repository:** https://github.com/jthampan/Malayalam_Church_Songs

**Video Tutorial:** [Coming soon]

---

## 🔄 Updates

To get the latest version:
1. Download the new .exe file from GitHub
2. Replace your old .exe file
3. Your settings (source folder) will be remembered

---

## 🎯 Quick Video Guide (Coming Soon)

Watch a 2-minute video showing:
1. First-time setup
2. Generating your first PPT
3. Tips and tricks

---

**Enjoy the simplified PPT generation! 🎵**
