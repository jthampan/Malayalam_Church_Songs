# 🎵 Malayalam Church Songs - PPT Generator (For Users)

## ✨ **Super Simple - No Installation Needed!**

Generate your Holy Communion Service PowerPoint in 3 clicks:

---

## 🚀 **How to Use (For Non-Technical Users)**

### **Step 1: Open Google Colab**
Click this link to open the generator:

👉 **[OPEN PPT GENERATOR](https://colab.research.google.com/github/jthampan/Malayalam_Church_Songs/blob/main/Malayalam_Church_Songs_Generator.ipynb)**

*(Opens in Google Colab - works on any device with internet)*

---

### **Step 2: Run Each Cell**
You'll see numbered sections. Just click the **▶️ Play button** on each cell from top to bottom:

1. **Cell 1**: Install software (30 seconds)
2. **Cell 2**: Setup folders (5 seconds)  
3. **Cell 3**: Upload your service file
4. **Cell 4**: Generate PowerPoint (10-30 seconds)
5. **Cell 5**: Download your PPT!

---

### **Step 3: Prepare Your Service File**

Create a text file (in Notepad/TextEdit) with this format:

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

---

## 📋 **Complete Example Service File**

Save this as `service_8feb2026.txt`:

```
# Date: 8 Feb 2026
# Service: Holy Communion

313|Opening|
314|B/A|
315|Offertory|
316|Confession|
331|Communion|
343|Communion|
353|Communion|
354|Communion|
|Closing|yeshuveppole aakuvaan
```

---

## ❓ **FAQ**

**Q: Do I need to install Python?**  
A: No! Everything runs in the cloud (Google Colab).

**Q: Does it work on my phone/tablet?**  
A: Yes! Works on any device with a web browser.

**Q: What if a hymn is not found?**  
A: The system will add a title slide but no content slides. You may need to add that hymn manually.

**Q: Can I use hymns without numbers?**  
A: Yes! Use format: `|Section|song title name`  
Example: `|Closing|yeshuveppole aakuvaan`

**Q: How do I add the Holy Communion image?**  
A: Currently you need to upload it manually in Step 2 to the `images/` folder. Or the system will skip it if not found.

**Q: Where are the source PPT files?**  
A: You need to upload your existing PPT files with hymns in Step 2, or contact the administrator to get them pre-loaded.

---

## 🎯 **Quick Start (Summary)**

1. Click the [Colab link](https://colab.research.google.com/github/jthampan/Malayalam_Church_Songs/blob/main/Malayalam_Church_Songs_Generator.ipynb)
2. Press ▶️ on each cell (5 cells total)
3. Upload your service.txt file when prompted
4. Download your generated PowerPoint!

**That's it!** 🎉

---

## 📞 **Need Help?**

Contact: **joby.thampan@gmail.com**

Repository: https://github.com/jthampan/Malayalam_Church_Songs

---

**Made with ❤️ for Mar Thoma Syrian Church, Singapore**
