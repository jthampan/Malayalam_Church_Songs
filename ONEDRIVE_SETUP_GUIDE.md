# How to Setup OneDrive with the Windows App

There are **2 easy ways** to use your OneDrive files with the Church Songs Generator:

---

## ✅ Method 1: OneDrive Desktop Sync (Recommended)

**This is the easiest and best method!**

### Step 1: Make Sure OneDrive is Installed
- OneDrive should already be installed on Windows 10/11
- Look for the **cloud icon** in your taskbar (bottom-right, near the clock)

### Step 2: Sign In to OneDrive
1. Click the OneDrive cloud icon
2. Sign in with your Microsoft account
3. Choose folders to sync (make sure your Church folder is included)

### Step 3: Wait for Files to Sync
- Your files will automatically download to:
  ```
  C:\Users\YourName\OneDrive\Documents\Church\Services\...
  ```
- You can see sync progress by clicking the OneDrive icon

### Step 4: Use in the App
1. Open Church_Songs_Generator.exe
2. Click "Browse" next to "Source PPT Folder"
3. Navigate to: `C:\Users\YourName\OneDrive\Documents\...`
4. Select your hymn folder
5. Done! ✅

### Benefits:
- ✅ Files automatically stay updated
- ✅ Works offline (files are on your PC)
- ✅ One-time setup
- ✅ Any changes in OneDrive appear locally

---

## 📥 Method 2: Manual Download (If OneDrive Sync Isn't Available)

Use this if you can't install OneDrive Desktop or prefer not to sync.

### Step 1: Download from OneDrive Web
1. Open your OneDrive link in a web browser
2. Navigate to your hymn PowerPoint folder
3. Select all files (Ctrl+A)
4. Click the **"Download"** button at the top
5. Browser will download a ZIP file

### Step 2: Extract Files
1. Find the downloaded ZIP file (usually in Downloads folder)
2. Right-click → "Extract All"
3. Choose a location (e.g., Desktop or Documents)
4. Extract the files

### Step 3: Use in the App
1. Open Church_Songs_Generator.exe
2. Click "Browse" next to "Source PPT Folder"
3. Navigate to where you extracted the files
4. Select that folder
5. Done! ✅

### Note:
- ⚠️ Files won't auto-update with this method
- You'll need to re-download if PPTs are updated in OneDrive
- Good for occasional use

---

## 🔗 About the "Sync from OneDrive" Button in the App

**Why doesn't automatic sync work?**

Microsoft OneDrive requires:
- Microsoft account authentication (OAuth login)
- App registration with Microsoft
- Special permissions

This is complex for a simple standalone app. The two methods above are:
- ✅ Simpler for users
- ✅ More reliable
- ✅ Work offline
- ✅ No authentication needed

---

## 💡 Which Method Should I Use?

| Situation | Recommended Method |
|-----------|-------------------|
| You have Windows 10/11 | **Method 1** (OneDrive Desktop) |
| You generate PPTs frequently | **Method 1** (OneDrive Desktop) |
| Files change often | **Method 1** (OneDrive Desktop) |
| One-time use | **Method 2** (Manual Download) |
| No internet at home | **Method 2** (Download once) |
| Corporate computer (can't install OneDrive) | **Method 2** (Manual Download) |

---

## ❓ Troubleshooting

### "OneDrive folder is empty"
- Right-click the OneDrive icon → Settings
- Go to "Account" tab
- Click "Choose folders"
- Make sure your Church folder is checked

### "Files show cloud icon"
- OneDrive is set to "Files On-Demand"
- Right-click the folder → "Always keep on this device"
- Files will download

### "Can't find OneDrive folder"
- Open File Explorer
- Look for "OneDrive" in the left sidebar
- Or go to: `C:\Users\YourName\OneDrive`

### "OneDrive not installed"
- Download from: https://www.microsoft.com/en-us/microsoft-365/onedrive/download
- Or use Method 2 (Manual Download)

---

## 📧 Still Need Help?

Contact: joby.thampan@gmail.com

---

**Bottom Line:**  
Most users should use **Method 1** (OneDrive Desktop Sync) - it's automatic, keeps files updated, and works offline!
