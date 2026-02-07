#!/usr/bin/env python3
"""
Malayalam Church Songs PPT Generator - GUI Version
Simple Windows application
"""

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import subprocess
import sys
import os
from pathlib import Path
from datetime import datetime
import threading

class PPTGeneratorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Malayalam Church Songs - PPT Generator")
        self.root.geometry("900x600")  # Wide enough for long paths
        self.root.resizable(False, False)
        
        # Variables
        self.service_file = tk.StringVar()
        self.source_folder = tk.StringVar()
        self.language = tk.StringVar(value="Malayalam")  # Default to Malayalam
        self.settings_file = Path.home() / ".church_ppt_settings.txt"
        
        # Load saved settings
        self.load_settings()
        
        # Create UI
        self.create_ui()
        
    def create_ui(self):
        # Header
        header_frame = tk.Frame(self.root, bg="#2c3e50", height=80)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        title = tk.Label(
            header_frame,
            text="🎵 Malayalam Church Songs",
            font=("Arial", 18, "bold"),
            bg="#2c3e50",
            fg="white"
        )
        title.pack(pady=10)
        
        subtitle = tk.Label(
            header_frame,
            text="PowerPoint Generator for Holy Communion Services",
            font=("Arial", 10),
            bg="#2c3e50",
            fg="#ecf0f1"
        )
        subtitle.pack()
        
        # Main content
        main_frame = tk.Frame(self.root, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Section 1: Source PPT Folder (One-time setup)
        setup_label = tk.Label(
            main_frame,
            text="⚙️ One-Time Setup (First Time Only)",
            font=("Arial", 12, "bold")
        )
        setup_label.grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=(0, 10))
        
        # Language selection
        lang_label = tk.Label(main_frame, text="Language:", font=("Arial", 10))
        lang_label.grid(row=0, column=3, sticky=tk.W, pady=5, padx=(20, 0))
        
        lang_combo = ttk.Combobox(
            main_frame,
            textvariable=self.language,
            values=["Malayalam", "English"],
            state="readonly",
            width=12
        )
        lang_combo.grid(row=0, column=4, pady=5)
        lang_combo.current(0)
        
        # Option 1: Browse local folder
        source_label = tk.Label(main_frame, text="Option 1 - Local Folder:", font=("Arial", 10))
        source_label.grid(row=1, column=0, sticky=tk.W, pady=5)
        
        source_entry = tk.Entry(main_frame, textvariable=self.source_folder, width=40, state="readonly")
        source_entry.grid(row=1, column=1, padx=5, pady=5)
        
        source_btn = tk.Button(
            main_frame,
            text="Browse...",
            command=self.select_source_folder,
            bg="#3498db",
            fg="white",
            font=("Arial", 9, "bold"),
            padx=10
        )
        source_btn.grid(row=1, column=2, pady=5)
        
        # Option 2: OneDrive link
        onedrive_label = tk.Label(main_frame, text="Option 2 - OneDrive Link:", font=("Arial", 10))
        onedrive_label.grid(row=2, column=0, sticky=tk.W, pady=5)
        
        self.onedrive_link = tk.StringVar()
        onedrive_entry = tk.Entry(main_frame, textvariable=self.onedrive_link, width=40)
        onedrive_entry.grid(row=2, column=1, padx=5, pady=5)
        
        onedrive_btn = tk.Button(
            main_frame,
            text="Help / Open Link",
            command=self.sync_from_onedrive,
            bg="#9b59b6",
            fg="white",
            font=("Arial", 9, "bold"),
            padx=10
        )
        onedrive_btn.grid(row=2, column=2, pady=5)
        
        # Help text
        help_label = tk.Label(
            main_frame,
            text="💡 Best: Use OneDrive Desktop sync, then browse to C:\\Users\\...\\OneDrive\\... folder",
            font=("Arial", 8),
            fg="#7f8c8d"
        )
        help_label.grid(row=3, column=0, columnspan=5, sticky=tk.W, pady=(0, 5))
        
        # Separator
        separator = ttk.Separator(main_frame, orient=tk.HORIZONTAL)
        separator.grid(row=4, column=0, columnspan=5, sticky="ew", pady=20)
        
        # Section 2: Generate PPT (Every time)
        generate_label = tk.Label(
            main_frame,
            text="🎉 Generate PowerPoint (Every Time)",
            font=("Arial", 12, "bold")
        )
        generate_label.grid(row=5, column=0, columnspan=5, sticky=tk.W, pady=(0, 10))
        
        service_label = tk.Label(main_frame, text="Service File:", font=("Arial", 10))
        service_label.grid(row=6, column=0, sticky=tk.W, pady=5)
        
        service_entry = tk.Entry(main_frame, textvariable=self.service_file, width=40, state="readonly")
        service_entry.grid(row=6, column=1, padx=5, pady=5)
        
        service_btn = tk.Button(
            main_frame,
            text="Browse...",
            command=self.select_service_file,
            bg="#3498db",
            fg="white",
            font=("Arial", 9, "bold"),
            padx=10
        )
        service_btn.grid(row=6, column=2, pady=5)
        
        # Generate button (big and prominent)
        self.generate_btn = tk.Button(
            main_frame,
            text="🎵 GENERATE POWERPOINT",
            command=self.generate_ppt,
            bg="#27ae60",
            fg="white",
            font=("Arial", 14, "bold"),
            height=2,
            cursor="hand2"
        )
        self.generate_btn.grid(row=7, column=0, columnspan=5, pady=20, sticky="ew")
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=8, column=0, columnspan=5, sticky="ew", pady=5)
        
        # Output log
        log_label = tk.Label(main_frame, text="Output:", font=("Arial", 10, "bold"))
        log_label.grid(row=9, column=0, columnspan=5, sticky=tk.W, pady=(10, 5))
        
        self.log_text = scrolledtext.ScrolledText(
            main_frame,
            height=12,
            width=105,
            state="disabled",
            wrap=tk.CHAR,
            font=("Consolas", 8)
        )
        self.log_text.grid(row=10, column=0, columnspan=5, pady=5, sticky="ew")
        
        # Footer with help
        footer_frame = tk.Frame(self.root, bg="#ecf0f1", height=40)
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM)
        footer_frame.pack_propagate(False)
        
        help_text = tk.Label(
            footer_frame,
            text="Need help? Contact: joby.thampan@gmail.com",
            font=("Arial", 8),
            bg="#ecf0f1",
            fg="#7f8c8d"
        )
        help_text.pack(pady=10)
        
    def load_settings(self):
        """Load previously saved source folder"""
        if self.settings_file.exists():
            try:
                with open(self.settings_file, 'r') as f:
                    saved_folder = f.read().strip()
                    if os.path.isdir(saved_folder):
                        self.source_folder.set(saved_folder)
                        self.log("✅ Loaded saved source folder: " + saved_folder)
            except:
                pass
                
    def save_settings(self):
        """Save source folder for next time"""
        try:
            with open(self.settings_file, 'w') as f:
                f.write(self.source_folder.get())
        except:
            pass
            
    def find_language_folder(self, base_folder):
        """Intelligently find the correct language-specific PPT folder.
        
        Searches for standard structure:
        - Holy Communion Services - Slides/Malayalam HCS/
        - Holy Communion Services - Slides/English HCS/
        
        Returns the language root folder to search ALL year subfolders recursively.
        """
        language = self.language.get()
        base_path = Path(base_folder)
        
        # Check if this is the parent folder with standard structure
        hcs_folder = base_path / "Holy Communion Services - Slides"
        if hcs_folder.exists():
            if language == "Malayalam":
                lang_folder = hcs_folder / "Malayalam HCS"
            else:
                lang_folder = hcs_folder / "English HCS"
            
            if lang_folder.exists():
                # Return the language folder root to search ALL year folders
                ppt_files = list(lang_folder.glob("**/*.ppt*"))  # Recursive search
                if ppt_files:
                    self.log(f"📁 Auto-detected: {lang_folder.relative_to(base_path)}")
                    self.log(f"   Will search all year folders (2024, 2025, 2026, etc.)")
                    return str(lang_folder)
        
        # Check if user selected the HCS folder directly
        if base_path.name == "Holy Communion Services - Slides":
            if language == "Malayalam":
                lang_folder = base_path / "Malayalam HCS"
            else:
                lang_folder = base_path / "English HCS"
            
            if lang_folder.exists():
                ppt_files = list(lang_folder.glob("**/*.ppt*"))
                if ppt_files:
                    self.log(f"📁 Will search all subfolders in: {lang_folder.name}")
                    return str(lang_folder)
        
        # Check if user selected the language folder directly (Malayalam HCS or English HCS)
        if "Malayalam HCS" in str(base_path) or "English HCS" in str(base_path):
            # Return this folder - will search all year subfolders
            return str(base_path)
        
        # User selected a specific folder - use as-is
        return base_folder
    
    def select_source_folder(self):
        """Select folder containing source PPT files"""
        folder = filedialog.askdirectory(
            title="Select folder with your hymn PPT files (or parent OneDrive folder)",
            initialdir=str(Path.home())
        )
        if folder:
            # Try to intelligently find the language-specific folder
            detected_folder = self.find_language_folder(folder)
            
            # Validate folder has PPT files
            ppt_files = list(Path(detected_folder).glob("*.ppt*"))
            
            # Also search recursively if directly in detected folder not found
            if len(ppt_files) == 0:
                ppt_files = list(Path(detected_folder).glob("**/*.ppt*"))
            
            if len(ppt_files) == 0:
                response = messagebox.askyesno(
                    "No PowerPoint Files Found",
                    f"Warning: No PowerPoint files found in:\n\n{detected_folder}\n\n"
                    f"Selected folder: {folder}\n\n"
                    "This folder appears to be empty or doesn't contain .pptx/.ppt files.\n\n"
                    "Common issues:\n"
                    "• Wrong folder selected\n"
                    f"• No {self.language.get()} PPT files in this location\n"
                    "• OneDrive files are cloud-only (not downloaded)\n"
                    "• Files have different extensions\n\n"
                    "If using OneDrive:\n"
                    "→ Right-click folder → 'Always keep on this device'\n"
                    "→ Wait for sync to complete\n\n"
                    "Do you want to use this folder anyway?"
                )
                if not response:
                    return
            else:
                # Check for potential OneDrive cloud-only files
                if "onedrive" in detected_folder.lower():
                    sample_file = ppt_files[0]
                    file_size = os.path.getsize(sample_file)
                    if file_size < 1000:  # Suspiciously small - might be placeholder
                        messagebox.showwarning(
                            "OneDrive Sync Warning",
                            f"Warning: Files may not be fully downloaded!\n\n"
                            f"Sample file size: {file_size} bytes (very small)\n\n"
                            "OneDrive may be showing cloud-only placeholders.\n\n"
                            "To fix:\n"
                            "1. Right-click the folder\n"
                            "2. Select 'Always keep on this device'\n"
                            "3. Wait for OneDrive to download all files\n"
                            "4. Try again\n\n"
                            "Folder will be saved, but generation may fail until files are downloaded."
                        )
            
            # Save the detected folder (not the original selection)
            self.source_folder.set(detected_folder)
            self.save_settings()
            self.log(f"✅ Source folder set: {detected_folder}")
            self.log(f"   Found {len(ppt_files)} PowerPoint files (searching all subfolders)")
            
            if folder != detected_folder:
                self.log(f"   (Auto-detected from: {folder})")
            
            if len(ppt_files) > 0:
                match_msg = ""
                if folder != detected_folder:
                    match_msg = f"\n✨ Smart Detection:\nYou selected: {Path(folder).name}\nUsing {self.language.get()} folder: {Path(detected_folder).name}\nSearching all year subfolders (2024, 2025, 2026, etc.)\n"
                
                messagebox.showinfo(
                    "Setup Complete",
                    f"Source folder saved!\n\n"
                    f"Found {len(ppt_files)} PowerPoint files across all subfolders.{match_msg}\n"
                    "You won't need to select this again.\n\n"
                    "Now you can generate presentations by selecting a service file and clicking Generate."
                )
    
    def sync_from_onedrive(self):
        """Download files from OneDrive link"""
        onedrive_link = self.onedrive_link.get().strip()
        
        if not onedrive_link:
            # Show help dialog even without link
            messagebox.showinfo(
                "OneDrive Setup Help",
                "🌐 How to Use OneDrive Files:\n\n"
                "METHOD 1 (Recommended):\n"
                "• Make sure OneDrive Desktop is installed\n"
                "• Sign in and sync your files\n"
                "• Use 'Browse' → Select OneDrive folder\n"
                "• Files stay updated automatically!\n\n"
                "METHOD 2:\n"
                "• Paste your OneDrive link above\n"
                "• Click this button to open in browser\n"
                "• Download files manually\n"
                "• Use 'Browse' → Select downloaded folder\n\n"
                "See ONEDRIVE_SETUP_GUIDE.md for detailed instructions."
            )
            return
        
        # If link is provided, offer to open it
        response = messagebox.askyesno(
            "Open OneDrive Link",
            "📎 OneDrive Link Detected!\n\n"
            f"{onedrive_link}\n\n"
            "I'll open this link in your browser.\n\n"
            "Then:\n"
            "1. Select all files\n"
            "2. Click 'Download'\n"
            "3. Extract the downloaded ZIP\n"
            "4. Use 'Browse' button to select that folder\n\n"
            "💡 TIP: Use OneDrive Desktop sync instead for automatic updates!\n\n"
            "Open link in browser now?"
        )
        
        if response:
            import webbrowser
            webbrowser.open(onedrive_link)
            self.log("🌐 Opened OneDrive link in browser")
            self.log("📥 Download files, then use 'Browse' button to select folder")
            self.log("")
            self.log("💡 For automatic sync, use OneDrive Desktop (see help)")
    
    def select_service_file(self):
        """Select service text file"""
        file = filedialog.askopenfilename(
            title="Select your service file",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
            initialdir=str(Path.home() / "Documents")
        )
        if file:
            self.service_file.set(file)
            self.log(f"✅ Service file selected: {os.path.basename(file)}")
    
    def log(self, message):
        """Add message to log"""
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")
        self.root.update()
    
    def generate_ppt(self):
        """Generate PowerPoint in a background thread"""
        # Validate inputs
        if not self.source_folder.get():
            messagebox.showerror(
                "Missing Source Folder",
                "Please select your source PPT folder first (One-Time Setup section)."
            )
            return
        
        # Validate source folder exists
        if not os.path.exists(self.source_folder.get()):
            messagebox.showerror(
                "Source Folder Not Found",
                f"The source folder no longer exists:\n\n{self.source_folder.get()}\n\n"
                "This can happen if:\n"
                "• OneDrive folder was moved or deleted\n"
                "• OneDrive is not synced\n"
                "• External drive was disconnected\n\n"
                "Please:\n"
                "1. Check OneDrive is synced and folder exists\n"
                "2. Use 'Browse' button to select the correct folder"
            )
            return
        
        # Validate source folder has PPT files
        ppt_files = list(Path(self.source_folder.get()).glob("**/*.ppt*"))  # Recursive search
        if len(ppt_files) == 0:
            messagebox.showerror(
                "No PowerPoint Files Found",
                f"No PowerPoint files found in:\n\n{self.source_folder.get()}\n\n"
                "Please make sure:\n"
                "• You selected the correct folder with hymn PPT files\n"
                "• Files have .pptx or .ppt extension\n"
                "• OneDrive files are downloaded (not cloud-only)\n\n"
                "If using OneDrive:\n"
                "→ Right-click folder → 'Always keep on this device'"
            )
            return
            
        if not self.service_file.get():
            messagebox.showerror(
                "Missing Service File",
                "Please select your service text file."
            )
            return
        
        # Validate service file exists
        if not os.path.exists(self.service_file.get()):
            messagebox.showerror(
                "Service File Not Found",
                f"The service file no longer exists:\n\n{self.service_file.get()}\n\n"
                "Please select the file again using the 'Browse' button."
            )
            return
        
        # Run in background thread to prevent UI freezing
        thread = threading.Thread(target=self._generate_ppt_thread, daemon=True)
        thread.start()
    
    def _generate_ppt_thread(self):
        """Background thread for generation"""
        try:
            # Disable button and start progress
            self.generate_btn.config(state="disabled", text="🔄 GENERATING...")
            self.progress.start(10)
            
            self.log("\n" + "="*70)
            self.log("🎵 Starting PowerPoint generation...")
            self.log("="*70 + "\n")
            
            # Prepare output filename
            date_str = datetime.now().strftime("%d_%b_%Y")
            output_file = str(Path.home() / "Desktop" / f"HCS_Malayalam_{date_str}.pptx")
            
            # Find generate_hcs_ppt.py
            script_path = self._find_generator_script()
            if not script_path:
                raise Exception("Generator script not found. Please ensure generate_hcs_ppt.py is in the same folder as this program.")
            
            self.log(f"📄 Service file: {os.path.basename(self.service_file.get())}")
            self.log(f"📁 Source folder: {self.source_folder.get()}")
            self.log(f"🌐 Language: {self.language.get()}")
            self.log(f"💾 Output: {output_file}\n")
            
            # Create a temporary batch file that includes language directive
            temp_batch_file = Path.home() / "Desktop" / ".temp_service.txt"
            with open(temp_batch_file, 'w', encoding='utf-8') as temp_f:
                # Add language directive
                temp_f.write(f"# Language: {self.language.get()}\n")
                # Copy original service file content
                with open(self.service_file.get(), 'r', encoding='utf-8') as orig_f:
                    temp_f.write(orig_f.read())
            
            # Run the generator with timeout
            try:
                result = subprocess.run(
                    [sys.executable, script_path, "--batch", str(temp_batch_file), output_file],
                    capture_output=True,
                    text=True,
                    cwd=self.source_folder.get(),  # Run from source folder to find PPTs
                    timeout=300  # 5 minute timeout
                )
            except subprocess.TimeoutExpired:
                self.log("\n" + "="*70)
                self.log("⏱️ TIMEOUT ERROR")
                self.log("="*70)
                self.log("\nGeneration took too long (>5 minutes).")
                self.log("\nPossible causes:")
                self.log("  • Very large PPT files")
                self.log("  • Network drive is slow")
                self.log("  • OneDrive is not fully synced")
                self.log("\nPlease check:")
                self.log("  1. OneDrive files are fully downloaded")
                self.log("  2. Source folder is on local drive (not network)")
                self.log("  3. PowerPoint files are not corrupted")
                
                # Clean up temporary file
                try:
                    if temp_batch_file.exists():
                        temp_batch_file.unlink()
                except:
                    pass
                
                raise Exception("Generation timed out after 5 minutes")
            finally:
                # Clean up temporary file
                try:
                    if temp_batch_file.exists():
                        temp_batch_file.unlink()
                except:
                    pass
            
            # Show output
            if result.stdout:
                self.log(result.stdout)
            
            if result.stderr:
                self.log("\n⚠️ Warnings:")
                self.log(result.stderr)
            
            # Check success
            if result.returncode == 0 and os.path.exists(output_file):
                file_size = os.path.getsize(output_file) / 1024
                self.log("\n" + "="*70)
                self.log("✅ SUCCESS!")
                self.log("="*70)
                self.log(f"\n📥 Your PowerPoint is ready: {output_file}")
                self.log(f"   File size: {file_size:.1f} KB")
                
                # Ask to open file
                if messagebox.askyesno(
                    "Success!",
                    f"PowerPoint generated successfully!\n\n"
                    f"Saved to: {output_file}\n\n"
                    f"Would you like to open it now?"
                ):
                    os.startfile(output_file)  # Windows-specific
                    
            else:
                self.log("\n" + "="*70)
                self.log("❌ ERROR")
                self.log("="*70)
                self.log("\nGeneration failed. Please check the messages above.")
                
                # Provide helpful error analysis
                error_hints = []
                
                if result.stderr and "not found" in result.stderr.lower():
                    error_hints.append("• Some hymn numbers were not found in your PPT files")
                    error_hints.append("• Check that the hymn numbers in your service file are correct")
                    error_hints.append("• Verify all required PPT files are in the source folder")
                
                if result.stderr and ("permission" in result.stderr.lower() or "access" in result.stderr.lower()):
                    error_hints.append("• Permission denied accessing files")
                    error_hints.append("• Check OneDrive files are not locked or in use")
                    error_hints.append("• Close any open PowerPoint files")
                
                if not os.path.exists(output_file):
                    error_hints.append("• Output file was not created")
                    error_hints.append("• Check you have write permission to Desktop")
                    error_hints.append("• Verify source PPT files are readable")
                
                # Check for OneDrive sync issues
                source_path = Path(self.source_folder.get())
                if "onedrive" in str(source_path).lower():
                    # Check if files might be cloud-only
                    ppt_files = list(source_path.glob("*.ppt*"))
                    if ppt_files:
                        sample_file = ppt_files[0]
                        # Check if file size is suspiciously small (might be cloud placeholder)
                        if os.path.getsize(sample_file) < 1000:
                            error_hints.append("• OneDrive files may not be fully downloaded!")
                            error_hints.append("→ Right-click source folder → 'Always keep on this device'")
                            error_hints.append("→ Wait for OneDrive to finish syncing")
                
                if error_hints:
                    self.log("\n💡 Possible issues:")
                    for hint in error_hints:
                        self.log(hint)
                
                messagebox.showerror(
                    "Generation Failed",
                    "Failed to generate PowerPoint.\n\n"
                    "Please check the output log for details.\n\n"
                    "Common issues:\n"
                    "• Hymn numbers not found in PPT files\n"
                    "• OneDrive files not fully downloaded\n"
                    "• Source folder is inaccessible\n"
                    "• Service file format is incorrect"
                )
                
        except Exception as e:
            self.log(f"\n❌ ERROR: {str(e)}")
            
            # Provide context-specific error messages
            error_msg = str(e)
            help_text = "\n\nPlease check:\n"
            
            if "timeout" in error_msg.lower():
                help_text += "• OneDrive files are fully downloaded\n"
                help_text += "• Source folder is on local drive\n"
                help_text += "• Network connection is stable"
            elif "not found" in error_msg.lower():
                help_text += "• Source folder exists and is accessible\n"
                help_text += "• OneDrive is synced properly\n"
                help_text += "• Files are not cloud-only"
            elif "permission" in error_msg.lower():
                help_text += "• You have access to the OneDrive folder\n"
                help_text += "• Files are not locked by another program\n"
                help_text += "• Desktop is writable"
            else:
                help_text += "• Source folder has PPT files\n"
                help_text += "• Service file format is correct\n"
                help_text += "• OneDrive is properly synced"
            
            messagebox.showerror(
                "Error", 
                f"An error occurred:\n\n{error_msg}{help_text}"
            )
            
        finally:
            # Re-enable button and stop progress
            self.progress.stop()
            self.generate_btn.config(state="normal", text="🎵 GENERATE POWERPOINT")
    
    def _find_generator_script(self):
        """Find generate_hcs_ppt.py in various locations"""
        # Check current directory
        script_name = "generate_hcs_ppt.py"
        
        # Same directory as this script
        current_dir = Path(__file__).parent if hasattr(sys, '_MEIPASS') else Path.cwd()
        if (current_dir / script_name).exists():
            return str(current_dir / script_name)
        
        # PyInstaller temp folder
        if hasattr(sys, '_MEIPASS'):
            temp_path = Path(sys._MEIPASS) / script_name
            if temp_path.exists():
                return str(temp_path)
        
        # Working directory
        if (Path.cwd() / script_name).exists():
            return str(Path.cwd() / script_name)
        
        return None

def main():
    root = tk.Tk()
    app = PPTGeneratorGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
