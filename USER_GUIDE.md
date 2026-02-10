# 🎵 Malayalam Church Songs - PPT Generator (For Users)

## ✨ **Super Simple Windows App**

Generate your Holy Communion Service PowerPoint in a few clicks.

---

## 🚀 **How to Use**

### **Step 1: Open the App**
- Run `Church_Songs_Generator.exe`

### **Step 2: Setup Source Folder (Optional)**

**Good News: No setup needed!** 🎉

The exe includes all PPT files - it works offline immediately with no downloads or configuration.

**Optional: Use Your Own Files**
1. Click **Browse...** next to **Source Folder**
2. Select your OneDrive folder with latest PPT files
3. The app will use your files instead of the bundled ones

### **What's Included:**
- ✅ All Malayalam PPT files (bundled at build time)
- ✅ Works completely offline
- ✅ No internet or git needed
- ✅ No passwords required

### **Step 3: Enter the Service List**
Use this format inside the Service Songs List box:

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

**Format:** `hymn_number|Section|optional_title`

**Sections you can use:**
- Opening
- B/A (Thanksgiving Prayers)
- Offertory
- Confession
- Communion (you can have multiple)
- Closing
- Dedication

### **Step 4: Generate**
Click **Generate PowerPoint**. The PPT is saved to your Desktop.

---

## 📁 **Setting Up Offline Mode**

**For Exe Users: Nothing to do!** ✅

The executable includes all PPT files - it works offline out of the box.

**For Script Users (running Python directly):**
- First run: Automatically downloads PPT files from git
- Subsequent runs: Uses cached files
- Manual setup: Create `onedrive_git_local` folder with PPT files

**For Builders (creating new exe):**
Edit `Windows_Exe/build_exe.py` to configure git repository:
```python
GIT_REPO_URL = "https://github.com/yourusername/Malayalam_Church_Songs.git"
GIT_BRANCH = "main"
```

For private repos, use Personal Access Token:
```python
GIT_REPO_URL = "https://YOUR_TOKEN@github.com/yourusername/Malayalam_Church_Songs.git"
```

---

## ❓ **FAQ**

**Q: Do I need internet?**  
A: No! The exe includes all PPT files. Works completely offline.

**Q: What if I don't provide a source folder?**  
A: The exe uses bundled PPT files automatically.

**Q: How big is the exe file?**  
A: Approximately 50-100 MB (includes all PPT files, images, and Python runtime).

**Q: What if a hymn is not found?**  
A: It will add a title slide only for that hymn.

**Q: Can I use hymns without numbers?**  
A: Yes. Example: `|Closing|yeshuveppole aakuvaan`

**Q: Can I use my OneDrive synced folder?**  
A: Yes! Just browse to your OneDrive folder to use the latest files.

**Q: How do I update the PPT files in the exe?**  
A: Rebuild the exe - it will download the latest files from git during build.

---

## 📞 **Need Help?**

Contact: **joby.thampan@gmail.com**

Repository: https://github.com/jthampan/Malayalam_Church_Songs

---

**Made with ❤️ for Mar Thoma Syrian Church, Singapore**
