#!/usr/bin/env python3
"""
Malayalam Church Songs PPT Generator - GUI Version
Simple Windows application
"""

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import sys
import os
import json
from pathlib import Path
from contextlib import redirect_stdout, redirect_stderr
from io import StringIO
from datetime import datetime
import threading
import re

class PPTGeneratorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Malayalam Church Songs - PPT Generator")
        self.root.geometry("980x640")  # Larger to fit output box comfortably
        self.root.resizable(False, False)
        
        # Variables
        self.source_folder = tk.StringVar()
        self.language = tk.StringVar(value="Malayalam")  # Default to Malayalam
        self.ppt_count = tk.StringVar(value="No folder selected")
        self.settings_file = Path.home() / ".church_ppt_settings.txt"
        self._is_loading = True  # Flag to track initial load
        self.default_service_text = (
            "# Format: hymn_num|label|title_hint\n"
            "# Example:\n"
            "# Date: 8 Feb 2026\n"
            "91|Opening|\n"
            "110|ThanksGiving|\n"
            "420|Offertory|\n"
            "Message\n"
            "211|Confession|\n"
            "313|Communion|\n"
            "427|Closing|\n"
        )
        
        # Add trace to auto-update PPT count when folder path changes
        self.source_folder.trace_add('write', lambda *args: self.update_ppt_count())
        
        # Create UI first
        self.create_ui()
        
        # Load saved settings (after UI is created so logging works)
        self.load_settings()
        
        # Mark loading complete
        self._is_loading = False
    
    def create_ui(self):
        # Header
        # Main content
        main_frame = tk.Frame(self.root, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Section 1: HCS PPT Generator
        setup_label = tk.Label(
            main_frame,
            text="⚙️ HCS PPT Generator",
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
        
        # Source folder
        source_label = tk.Label(main_frame, text="Source Folder:", font=("Arial", 10))
        source_label.grid(row=1, column=0, sticky=tk.W, pady=5)
        
        source_entry = tk.Entry(main_frame, textvariable=self.source_folder, width=40)
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
        
        # PPT count display
        ppt_count_label = tk.Label(
            main_frame,
            textvariable=self.ppt_count,
            font=("Arial", 9),
            fg="#27ae60",
            anchor=tk.W
        )
        ppt_count_label.grid(row=1, column=3, columnspan=2, sticky=tk.W, padx=(10, 0))
        
        # Help text
        help_label = tk.Label(
            main_frame,
            text="💡 Tip: Use OneDrive Desktop sync, then browse to C:\\Users\\...\\OneDrive\\... folder",
            font=("Arial", 8),
            fg="#7f8c8d"
        )
        help_label.grid(row=2, column=0, columnspan=5, sticky=tk.W, pady=(0, 5))
        
        # Separator
        separator = ttk.Separator(main_frame, orient=tk.HORIZONTAL)
        separator.grid(row=3, column=0, columnspan=5, sticky="ew", pady=20)
        
        # Section 2: Service Songs List
        generate_label = tk.Label(
            main_frame,
            text="🎉 Service Songs List",
            font=("Arial", 12, "bold")
        )
        generate_label.grid(row=4, column=0, columnspan=5, sticky=tk.W, pady=(0, 10))
        
        self.service_text = scrolledtext.ScrolledText(
            main_frame,
            height=7,
            width=72,
            wrap=tk.WORD,
            font=("Consolas", 9)
        )
        self.service_text.grid(row=5, column=0, columnspan=5, padx=5, pady=5, sticky="ew")
        self.service_text.insert("1.0", self.default_service_text)
        self.service_text.edit_modified(False)
        self.service_text.bind("<<Modified>>", self._on_service_text_change)
        
        # Generate button (small and centered)
        self.generate_btn = tk.Button(
            main_frame,
            text="🎵 GENERATE POWERPOINT",
            command=self.generate_ppt,
            bg="#27ae60",
            fg="white",
            font=("Arial", 11, "bold"),
            width=26,
            height=1,
            cursor="hand2"
        )
        self.generate_btn.grid(row=6, column=0, columnspan=5, pady=14)
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=7, column=0, columnspan=5, sticky="ew", pady=5)
        
        # Output log
        log_label = tk.Label(main_frame, text="Output:", font=("Arial", 10, "bold"))
        log_label.grid(row=8, column=0, columnspan=5, sticky=tk.W, pady=(10, 5))
        
        self.log_text = scrolledtext.ScrolledText(
            main_frame,
            height=9,
            width=105,
            state="disabled",
            wrap=tk.CHAR,
            font=("Consolas", 8)
        )
        self.log_text.grid(row=9, column=0, columnspan=5, pady=(5, 10), sticky="ew")
        
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
                    raw = f.read().strip()
                if raw.startswith("{"):
                    data = json.loads(raw)
                    saved_folder = data.get("source_folder", "").strip()
                    saved_text = data.get("service_text", "").strip()
                    saved_language = data.get("language", "Malayalam").strip()
                    if saved_folder and os.path.isdir(saved_folder):
                        self.source_folder.set(saved_folder)
                    if saved_text:
                        self.service_text.delete("1.0", tk.END)
                        self.service_text.insert("1.0", saved_text)
                        self.service_text.edit_modified(False)
                    if saved_language in ("Malayalam", "English"):
                        self.language.set(saved_language)
                else:
                    saved_folder = raw
                    if os.path.isdir(saved_folder):
                        self.source_folder.set(saved_folder)
                        # PPT count will be updated automatically by trace callback
            except:
                pass
                
    def save_settings(self):
        """Save source folder for next time"""
        try:
            data = {
                "source_folder": self.source_folder.get().strip(),
                "service_text": self.get_service_text().strip(),
                "language": self.language.get().strip(),
            }
            with open(self.settings_file, 'w') as f:
                f.write(json.dumps(data))
        except:
            pass

    def _on_service_text_change(self, event=None):
        if self._is_loading:
            self.service_text.edit_modified(False)
            return
        if self.service_text.edit_modified():
            self.service_text.edit_modified(False)
            self.save_settings()
    
    def update_ppt_count(self):
        """Check the current source folder path and update PPT count display"""
        folder_path = self.source_folder.get().strip()
        
        if not folder_path:
            self.ppt_count.set("No folder selected")
            return
        
        if not os.path.isdir(folder_path):
            self.ppt_count.set("✗ Invalid folder path")
            return
        
        # Try to intelligently find the language-specific folder
        detected_folder = self.find_language_folder(folder_path)
        
        # Count PPT files recursively
        try:
            ppt_files = list(Path(detected_folder).glob("**/*.ppt*"))
            
            if len(ppt_files) > 0:
                self.ppt_count.set(f"✓ Found {len(ppt_files)} PPT files (all subfolders)")
                # Log only on initial load
                if self._is_loading and hasattr(self, 'log_text'):
                    self.log(f"✅ Loaded saved source folder: {detected_folder}")
                    if folder_path != detected_folder:
                        self.log(f"   (Auto-detected from: {folder_path})")
            else:
                self.ppt_count.set("⚠ No PPT files found")
        except Exception as e:
            self.ppt_count.set("✗ Error scanning folder")
            
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
                    self.ppt_count.set("✗ Folder not selected")
                    return
                else:
                    self.ppt_count.set("⚠ No PPT files found")
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
            # Setting source_folder will trigger auto-update of PPT count via trace callback
            self.source_folder.set(detected_folder)
            self.save_settings()
            
            # Log additional info
            if folder != detected_folder:
                self.log(f"   Will search all year folders (2024, 2025, 2026, etc.)")
            self.log("")
    
    def get_service_text(self):
        return self.service_text.get("1.0", tk.END).strip()

    def _normalize_service_text(self, raw_text):
        lines = raw_text.splitlines()
        normalized_lines = []
        total_songs = 0
        communion_songs = 0

        for line in lines:
            cleaned = line.strip()
            if not cleaned:
                continue
            if cleaned.startswith("#"):
                normalized_lines.append(cleaned)
                continue

            if cleaned.lower() == "message":
                normalized_lines.append("Message")
                total_songs += 1
                continue

            parts = cleaned.split("|")
            if len(parts) >= 2:
                hymn_num = parts[0].strip()
                label = parts[1].strip()
                title_hint = parts[2].strip() if len(parts) > 2 else ""
                normalized_lines.append(f"{hymn_num}|{label}|{title_hint}")
                total_songs += 1
                if label.lower() in ("communion", "holy communion"):
                    communion_songs += 1
            else:
                normalized_lines.append(cleaned)

        return "\n".join(normalized_lines), total_songs, communion_songs

    def _extract_service_date(self, service_text):
        date_pattern = re.compile(r"^#\s*Date\s*:\s*(.+)$", re.IGNORECASE)
        for line in service_text.splitlines():
            match = date_pattern.match(line.strip())
            if match:
                return match.group(1).strip()
        return ""

    def _format_date_for_filename(self, date_text):
        if not date_text:
            return ""
        formats = [
            "%d %b %Y",
            "%d %B %Y",
            "%d-%b-%Y",
            "%d-%B-%Y",
            "%d/%m/%Y",
            "%d-%m-%Y",
        ]
        for fmt in formats:
            try:
                parsed = datetime.strptime(date_text, fmt)
                return parsed.strftime("%d_%b_%Y")
            except ValueError:
                continue
        return ""
    
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
            
        service_text = self.get_service_text()
        if not service_text:
            messagebox.showerror(
                "Missing Service List",
                "Please enter the service list in the text box."
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
            service_text_raw = self.get_service_text()
            service_date_text = self._extract_service_date(service_text_raw)
            date_str = self._format_date_for_filename(service_date_text)
            if not date_str:
                date_str = datetime.now().strftime("%d_%b_%Y")
            output_file = str(Path.home() / "Desktop" / f"HCS_Malayalam_{date_str}.pptx")
            if os.path.exists(output_file):
                try:
                    os.remove(output_file)
                    self.log(f"🧹 Removed existing output: {output_file}")
                except Exception as e:
                    raise Exception(
                        f"Cannot remove existing output file:\n{output_file}\n\n"
                        f"Error: {str(e)}\n\n"
                        "Please close the PPT if it is open and try again."
                    )
            
            # Find generate_hcs_ppt.py
            script_path = self._find_generator_script()
            if not script_path:
                raise Exception("Generator script not found. Please ensure generate_hcs_ppt.py is in the same folder as this program.")
            
            # Get full paths
            service_text_raw = self.get_service_text()
            service_text, total_songs, communion_songs = self._normalize_service_text(service_text_raw)
            
            # Debug logging
            self.log("📄 Service list: (input box)")
            self.log(f"🧮 Parsed songs: {total_songs} (Communion: {communion_songs})")
            self.log(f"📁 Source folder: {self.source_folder.get()}")
            self.log(f"🌐 Language: {self.language.get()}")
            self.log(f"💾 Output: {output_file}\n")
            
            # Create a temporary batch file that includes language directive
            desktop_path = Path.home() / "Desktop"
            
            # Ensure Desktop folder exists
            if not desktop_path.exists():
                try:
                    desktop_path.mkdir(parents=True, exist_ok=True)
                    self.log(f"📁 Created Desktop folder: {desktop_path}")
                except Exception as e:
                    raise Exception(
                        f"Cannot create Desktop folder: {desktop_path}\n\n"
                        f"Error: {str(e)}\n\n"
                        f"Please ensure you have write permissions to your home directory."
                    )
            
            temp_batch_file = desktop_path / ".temp_service.txt"
            self.log(f"📝 Creating temp file: {temp_batch_file}")
            
            try:
                # Write to temp file
                self.log(f"💾 Writing temp file...")
                with open(temp_batch_file, 'w', encoding='utf-8') as temp_f:
                    # Add language directive
                    temp_f.write(f"# Language: {self.language.get()}\n")
                    temp_f.write(service_text)
                    
                self.log(f"✅ Temp file created successfully\n")
                    
            except Exception as e:
                # Clean up temp file if it was created
                try:
                    if temp_batch_file.exists():
                        temp_batch_file.unlink()
                        self.log(f"🗑️ Cleaned up temp file")
                except:
                    pass
                raise e
            
            # Run the generator in-process to avoid relaunching the GUI exe
            generator_stdout = ""
            generator_stderr = ""
            generator_error = None

            try:
                try:
                    import generate_hcs_ppt
                except Exception as e:
                    raise Exception(
                        "Cannot load generate_hcs_ppt.py. Please ensure it is packaged with the app.\n\n"
                        f"Error: {str(e)}"
                    )

                original_argv = sys.argv[:]
                original_cwd = os.getcwd()
                stdout_buf = StringIO()
                stderr_buf = StringIO()

                try:
                    sys.argv = [script_path, "--batch", str(temp_batch_file), output_file]
                    os.chdir(self.source_folder.get())

                    with redirect_stdout(stdout_buf), redirect_stderr(stderr_buf):
                        generate_hcs_ppt.main()
                finally:
                    sys.argv = original_argv
                    os.chdir(original_cwd)

                generator_stdout = stdout_buf.getvalue().strip()
                generator_stderr = stderr_buf.getvalue().strip()
            except Exception as e:
                generator_error = str(e)
            finally:
                # Clean up temporary file
                try:
                    if temp_batch_file.exists():
                        temp_batch_file.unlink()
                except:
                    pass

            # Show output
            if generator_stdout:
                self.log(generator_stdout)
            if generator_stderr:
                self.log("\n⚠️ Warnings:")
                self.log(generator_stderr)
            if generator_error:
                self.log("\n❌ ERROR:")
                self.log(generator_error)

            # Check success
            if os.path.exists(output_file):
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
                
                error_text = "\n".join(filter(None, [generator_stderr, generator_error]))

                if error_text and "not found" in error_text.lower():
                    error_hints.append("• Some hymn numbers were not found in your PPT files")
                    error_hints.append("• Check that the hymn numbers in your service list are correct")
                    error_hints.append("• Verify all required PPT files are in the source folder")
                
                if error_text and ("permission" in error_text.lower() or "access" in error_text.lower()):
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
                    "• Service list format is incorrect"
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
                help_text += "• Service list format is correct\n"
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
