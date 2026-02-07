#!/usr/bin/env python3
"""
Malayalam Church Songs PPT Generator - GUI Version
Simple Windows application for non-technical users
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
        self.root.geometry("700x600")
        self.root.resizable(False, False)
        
        # Variables
        self.service_file = tk.StringVar()
        self.source_folder = tk.StringVar()
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
        help_label.grid(row=3, column=0, columnspan=3, sticky=tk.W, pady=(0, 5))
        
        help_label.grid(row=3, column=0, columnspan=3, sticky=tk.W, pady=(0, 5))
        
        # Separator
        separator = ttk.Separator(main_frame, orient=tk.HORIZONTAL)
        separator.grid(row=4, column=0, columnspan=3, sticky="ew", pady=20)
        
        # Section 2: Generate PPT (Every time)
        generate_label = tk.Label(
            main_frame,
            text="🎉 Generate PowerPoint (Every Time)",
            font=("Arial", 12, "bold")
        )
        generate_label.grid(row=5, column=0, columnspan=3, sticky=tk.W, pady=(0, 10))
        
        service_label = tk.Label(main_frame, text="Service File:", font=("Arial", 10))
        service_label.grid(row=6, column=0, sticky=tk.W, pady=5)
        
        service_entry = tk.Entry(main_frame, textvariable=self.service_file, width=40, state="readonly")
        service_entry.grid(row=6, column=1, padx=5, pady=5)
        
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
        self.generate_btn.grid(row=7, column=0, columnspan=3, pady=20, sticky="ew")
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=8, column=0, columnspan=3, sticky="ew", pady=5)
        
        # Output log
        log_label = tk.Label(main_frame, text="Output:", font=("Arial", 10, "bold"))
        log_label.grid(row=9, column=0, columnspan=3, sticky=tk.W, pady=(10, 5))
        
        self.log_text = scrolledtext.ScrolledText(
            main_frame,
            height=12,
            width=70,
            state="disabled",
            font=("Consolas", 9)
        )
        self.log_text.grid(row=10, column=0, columnspan=3, pady=5)
        
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
            
    def select_source_folder(self):
        """Select folder containing source PPT files"""
        folder = filedialog.askdirectory(
            title="Select folder with your hymn PPT files",
            initialdir=str(Path.home())
        )
        if folder:
            self.source_folder.set(folder)
            self.save_settings()
            self.log(f"✅ Source folder set: {folder}")
            messagebox.showinfo(
                "Setup Complete",
                "Source folder saved! You won't need to select this again.\n\n"
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
            
        if not self.service_file.get():
            messagebox.showerror(
                "Missing Service File",
                "Please select your service text file."
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
            self.log(f"💾 Output: {output_file}\n")
            
            # Run the generator
            result = subprocess.run(
                [sys.executable, script_path, "--batch", self.service_file.get(), output_file],
                capture_output=True,
                text=True,
                cwd=self.source_folder.get()  # Run from source folder to find PPTs
            )
            
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
                messagebox.showerror(
                    "Generation Failed",
                    "Failed to generate PowerPoint. Please check the output log for details."
                )
                
        except Exception as e:
            self.log(f"\n❌ ERROR: {str(e)}")
            messagebox.showerror("Error", f"An error occurred:\n\n{str(e)}")
            
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
