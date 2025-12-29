import customtkinter as ctk
import win32api
import win32print
import os
import time
import threading
from tkinter import filedialog, messagebox
from datetime import datetime

# Set appearance mode and color theme
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class MultiBatchPrinter(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Window configuration
        self.title("📄 Multi-Format Batch Printer Pro")
        self.geometry("1100x700")
        self.minsize(1000, 650)
        
        # Supported file formats
        self.supported_formats = {
            'pdf': '📕 PDF Files',
            'docx': '📘 Word Documents (DOCX)',
            'doc': '📘 Word Documents (DOC)',
            'xlsx': '📗 Excel Spreadsheets (XLSX)',
            'xls': '📗 Excel Spreadsheets (XLS)',
            'pptx': '📙 PowerPoint (PPTX)',
            'ppt': '📙 PowerPoint (PPT)',
            'txt': '📄 Text Files',
            'rtf': '📝 Rich Text Format',
            'odt': '📋 OpenDocument Text',
            'ods': '📊 OpenDocument Spreadsheet'
        }
        
        # Variables
        self.folder_path = ctk.StringVar(value="No folder selected")
        self.wait_time = ctk.IntVar(value=5)
        self.sort_ascending = ctk.BooleanVar(value=True)  # New variable for sort order
        self.selected_formats = {}  # Track which formats are selected
        self.print_files = []
        self.is_printing = False
        
        # Initialize all formats as selected
        for fmt in self.supported_formats.keys():
            self.selected_formats[fmt] = ctk.BooleanVar(value=True)
        
        # Create UI
        self.create_header()
        self.create_two_column_layout()
        self.create_status_bar()
        
        # Get default printer on startup
        self.update_printer_info()
    
    def create_header(self):
        header_frame = ctk.CTkFrame(self, corner_radius=10)
        header_frame.pack(fill="x", padx=20, pady=(15, 10))

        header_frame.grid_columnconfigure(0, weight=1)
        header_frame.grid_columnconfigure(1, weight=0)

        # LEFT: Title
        title_frame = ctk.CTkFrame(header_frame, fg_color="transparent")
        title_frame.grid(row=0, column=0, sticky="w", padx=15, pady=15)

        ctk.CTkLabel(
            title_frame,
            text="📄 Multi-Format Batch Printer Pro",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(anchor="w")

        ctk.CTkLabel(
            title_frame,
            text="Print PDF, Word, Excel, PowerPoint and more!",
            font=ctk.CTkFont(size=12),
            text_color="gray"
        ).pack(anchor="w")

        # RIGHT: ACTION BUTTONS
        action_frame = ctk.CTkFrame(header_frame, fg_color="transparent")
        action_frame.grid(row=0, column=1, sticky="e", padx=15)

        self.scan_btn = ctk.CTkButton(
            action_frame,
            text="🔍 Scan",
            command=self.scan_folder,
            width=110,
            height=36
        )
        self.scan_btn.pack(side="left", padx=5)

        self.print_btn = ctk.CTkButton(
            action_frame,
            text="🖨️ Print",
            command=self.start_printing,
            width=110,
            height=36,
            fg_color="#2fa572",
            hover_color="#28865e"
        )
        self.print_btn.pack(side="left", padx=5)
        self.print_btn.configure(state="disabled")

        self.stop_btn = ctk.CTkButton(
            action_frame,
            text="⏹ Stop",
            command=self.stop_printing,
            width=110,
            height=36,
            fg_color="#c42b1c",
            hover_color="#a21a0e"
        )
        self.stop_btn.pack(side="left", padx=5)
        self.stop_btn.configure(state="disabled")

        
    def create_two_column_layout(self):
        """Create two-column layout"""
        # Main container for columns
        columns_frame = ctk.CTkFrame(self, fg_color="transparent")
        columns_frame.pack(fill="both", expand=True, padx=15, pady=10)
        
        # LEFT COLUMN
        left_column = ctk.CTkFrame(columns_frame, fg_color="transparent")
        left_column.pack(side="left", fill="both", expand=True, padx=(0, 10))
        
        self.create_folder_section(left_column)
        self.create_format_selection(left_column)
        self.create_settings_section(left_column)
        
        
        # RIGHT COLUMN
        right_column = ctk.CTkFrame(columns_frame, fg_color="transparent")
        right_column.pack(side="right", fill="both", expand=True, padx=(10, 0))
        
        self.create_file_list_section(right_column)
        self.create_console_section(right_column)
    
    def create_folder_section(self, parent):
        """Create folder selection section"""
        folder_frame = ctk.CTkFrame(parent, corner_radius=10)
        folder_frame.pack(fill="x", pady=(0, 10))
        
        label = ctk.CTkLabel(
            folder_frame,
            text="📁 Select Folder:",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        label.pack(anchor="w", padx=15, pady=(15, 5))
        
        path_frame = ctk.CTkFrame(folder_frame, fg_color="transparent")
        path_frame.pack(fill="x", padx=15, pady=(0, 10))
        
        self.path_label = ctk.CTkLabel(
            path_frame,
            textvariable=self.folder_path,
            font=ctk.CTkFont(size=11),
            anchor="w",
            wraplength=400
        )
        self.path_label.pack(fill="x", pady=(0, 5))
        
        browse_btn = ctk.CTkButton(
            folder_frame,
            text="📂 Browse Folder",
            command=self.browse_folder,
            height=36,
            font=ctk.CTkFont(size=13, weight="bold")
        )
        browse_btn.pack(fill="x", padx=15, pady=(0, 15))
    
    def create_format_selection(self, parent):
        """Create file format selection section"""
        format_frame = ctk.CTkFrame(parent, corner_radius=10)
        format_frame.pack(fill="x", pady=(0, 10))
        
        # Header with select all/none buttons
        header = ctk.CTkFrame(format_frame, fg_color="transparent")
        header.pack(fill="x", padx=15, pady=(15, 5))
        
        label = ctk.CTkLabel(
            header,
            text="📋 File Formats:",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        label.pack(side="left")
        
        btn_frame = ctk.CTkFrame(header, fg_color="transparent")
        btn_frame.pack(side="right")
        
        select_all_btn = ctk.CTkButton(
            btn_frame,
            text="All",
            command=self.select_all_formats,
            width=50,
            height=24,
            font=ctk.CTkFont(size=10)
        )
        select_all_btn.pack(side="left", padx=2)
        
        select_none_btn = ctk.CTkButton(
            btn_frame,
            text="None",
            command=self.select_no_formats,
            width=50,
            height=24,
            font=ctk.CTkFont(size=10)
        )
        select_none_btn.pack(side="left", padx=2)
        
        # Scrollable frame for checkboxes
        scroll_frame = ctk.CTkScrollableFrame(format_frame, height=150)
        scroll_frame.pack(fill="x", padx=15, pady=(5, 15))
        
        # Create checkboxes for each format
        self.format_checkboxes = {}
        for fmt, description in self.supported_formats.items():
            cb = ctk.CTkCheckBox(
                scroll_frame,
                text=description,
                variable=self.selected_formats[fmt],
                font=ctk.CTkFont(size=11),
                command=self.on_format_change
            )
            cb.pack(anchor="w", pady=2)
            self.format_checkboxes[fmt] = cb
    
    def create_settings_section(self, parent):
        """Create settings section"""
        settings_frame = ctk.CTkFrame(parent, corner_radius=10)
        settings_frame.pack(fill="x", pady=(0, 10))
        
        # Title
        label = ctk.CTkLabel(
            settings_frame,
            text="⚙️ Settings:",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        label.pack(anchor="w", padx=15, pady=(15, 10))
        
        # Sort order button
        sort_frame = ctk.CTkFrame(settings_frame, fg_color="transparent")
        sort_frame.pack(fill="x", padx=15, pady=(0, 10))
        
        sort_label = ctk.CTkLabel(
            sort_frame,
            text="Print order:",
            font=ctk.CTkFont(size=12)
        )
        sort_label.pack(side="left")
        
        self.sort_btn = ctk.CTkButton(
            sort_frame,
            text="🔼 Ascending (A→Z)",
            command=self.toggle_sort_order,
            width=150,
            height=32,
            font=ctk.CTkFont(size=11)
        )
        self.sort_btn.pack(side="right")
        
        # Wait time slider
        wait_label = ctk.CTkLabel(
            settings_frame,
            text="Wait time between prints:",
            font=ctk.CTkFont(size=12)
        )
        wait_label.pack(anchor="w", padx=15, pady=(0, 5))
        
        self.wait_value_label = ctk.CTkLabel(
            settings_frame,
            text="5 seconds",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#1f6aa5"
        )
        self.wait_value_label.pack(anchor="w", padx=15, pady=(0, 5))
        
        slider = ctk.CTkSlider(
            settings_frame,
            from_=3,
            to=15,
            number_of_steps=12,
            variable=self.wait_time,
            command=self.update_wait_label
        )
        slider.pack(fill="x", padx=15, pady=(0, 15))
        
        # Printer info
        self.printer_label = ctk.CTkLabel(
            settings_frame,
            text="🖨️ Printer: Loading...",
            font=ctk.CTkFont(size=11),
            anchor="w",
            wraplength=400
        )
        self.printer_label.pack(anchor="w", padx=15, pady=(0, 15))
    
    def toggle_sort_order(self):
        """Toggle between ascending and descending sort order"""
        self.sort_ascending.set(not self.sort_ascending.get())
        
        if self.sort_ascending.get():
            self.sort_btn.configure(text="🔼 Ascending (A→Z)")
            self.log("Print order: Ascending (A→Z)")
        else:
            self.sort_btn.configure(text="🔽 Descending (Z→A)")
            self.log("Print order: Descending (Z→A)")
        
        # Re-scan if folder is already selected
        if self.folder_path.get() != "No folder selected":
            self.scan_folder()
    
    def create_file_list_section(self, parent):
        """Create file list section"""
        list_frame = ctk.CTkFrame(parent, corner_radius=10)
        list_frame.pack(fill="both", expand=True, pady=(0, 10))
        
        # Header
        header = ctk.CTkFrame(list_frame, fg_color="transparent")
        header.pack(fill="x", padx=15, pady=(15, 10))
        
        label = ctk.CTkLabel(
            header,
            text="📋 Files to Print:",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        label.pack(side="left")
        
        self.file_count_label = ctk.CTkLabel(
            header,
            text="0 files",
            font=ctk.CTkFont(size=12),
            text_color="gray"
        )
        self.file_count_label.pack(side="right")
        
        # Scrollable text box
        self.file_textbox = ctk.CTkTextbox(
            list_frame,
            height=250,
            font=ctk.CTkFont(family="Consolas", size=10),
            wrap="none"
        )
        self.file_textbox.pack(fill="both", expand=True, padx=15, pady=(0, 15))
    
    def create_console_section(self, parent):
        """Create console output section"""
        console_frame = ctk.CTkFrame(parent, corner_radius=10)
        console_frame.pack(fill="both", expand=True)
        
        label = ctk.CTkLabel(
            console_frame,
            text="📟 Console Output:",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        label.pack(anchor="w", padx=15, pady=(15, 10))
        
        self.console = ctk.CTkTextbox(
            console_frame,
            height=180,
            font=ctk.CTkFont(family="Consolas", size=10),
            wrap="word"
        )
        self.console.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        self.log("Ready to print multiple file formats.")
    
    def create_status_bar(self):
        """Create status bar"""
        status_frame = ctk.CTkFrame(self, corner_radius=10)
        status_frame.pack(fill="x", padx=20, pady=(0, 20))
        
        self.status_bar = ctk.CTkLabel(
            status_frame,
            text="⚡ Status: Ready",
            font=ctk.CTkFont(size=12, weight="bold"),
            anchor="w"
        )
        self.status_bar.pack(fill="x", padx=15, pady=10)
    
    def select_all_formats(self):
        """Select all file formats"""
        for var in self.selected_formats.values():
            var.set(True)
        self.log("All file formats selected")
    
    def select_no_formats(self):
        """Deselect all file formats"""
        for var in self.selected_formats.values():
            var.set(False)
        self.log("All file formats deselected")
    
    def on_format_change(self):
        """Called when format selection changes"""
        selected_count = sum(1 for var in self.selected_formats.values() if var.get())
        if selected_count == 0:
            self.log("Warning: No file formats selected", "warning")
    
    def browse_folder(self):
        """Open folder browser dialog"""
        folder = filedialog.askdirectory(title="Select folder containing files")
        if folder:
            self.folder_path.set(folder)
            self.log(f"Folder selected: {folder}")
            self.scan_folder()
    
    def update_wait_label(self, value):
        """Update wait time label"""
        seconds = int(float(value))
        self.wait_value_label.configure(text=f"{seconds} seconds")
    
    def update_printer_info(self):
        """Update printer information"""
        try:
            printer_name = win32print.GetDefaultPrinter()
            self.printer_label.configure(text=f"🖨️ Printer: {printer_name}")
            self.log(f"Default printer: {printer_name}")
        except Exception as e:
            self.printer_label.configure(text=f"🖨️ Printer: Error - {str(e)}")
            self.log(f"Error getting printer: {str(e)}", "error")
    
    def get_file_icon(self, extension):
        """Get icon for file type"""
        icons = {
            'pdf': '📕',
            'docx': '📘', 'doc': '📘',
            'xlsx': '📗', 'xls': '📗',
            'pptx': '📙', 'ppt': '📙',
            'txt': '📄',
            'rtf': '📝',
            'odt': '📋',
            'ods': '📊'
        }
        return icons.get(extension, '📄')
    
    def scan_folder(self):
        """Scan folder for printable files"""
        folder = self.folder_path.get()
        
        if folder == "No folder selected":
            messagebox.showwarning("Warning", "Please select a folder first!")
            return
        
        if not os.path.exists(folder):
            messagebox.showerror("Error", "Folder does not exist!")
            return
        
        # Check if at least one format is selected
        selected = [fmt for fmt, var in self.selected_formats.items() if var.get()]
        if not selected:
            messagebox.showwarning("Warning", "Please select at least one file format!")
            return
        
        sort_order = "ascending" if self.sort_ascending.get() else "descending"
        self.log(f"Scanning folder for {', '.join(selected)} files... (Order: {sort_order})")
        self.file_textbox.delete("1.0", "end")
        
        # Get files
        try:
            all_files = os.listdir(folder)
            self.print_files = []
            
            for filename in all_files:
                ext = filename.split('.')[-1].lower()
                if ext in selected:
                    self.print_files.append(filename)
            
            # Sort based on selected order
            self.print_files.sort(reverse=not self.sort_ascending.get())
            
            if not self.print_files:
                self.log("No matching files found in folder.", "warning")
                self.file_count_label.configure(text="0 files")
                self.print_btn.configure(state="disabled")
                return
            
            # Display files grouped by type
            file_types = {}
            for filename in self.print_files:
                ext = filename.split('.')[-1].lower()
                if ext not in file_types:
                    file_types[ext] = []
                file_types[ext].append(filename)
            
            total_count = 0
            for ext in sorted(file_types.keys()):
                files = file_types[ext]
                icon = self.get_file_icon(ext)
                
                self.file_textbox.insert("end", f"\n{icon} {ext.upper()} Files ({len(files)}):\n")
                self.file_textbox.insert("end", "─" * 40 + "\n")
                
                for i, filename in enumerate(files, 1):
                    full_path = os.path.join(folder, filename)
                    file_size = os.path.getsize(full_path)
                    size_kb = file_size / 1024
                    
                    total_count += 1
                    self.file_textbox.insert("end", f"{total_count}. {filename}\n")
                    self.file_textbox.insert("end", f"   Size: {size_kb:.1f} KB\n\n")
            
            count = len(self.print_files)
            self.file_count_label.configure(text=f"{count} file{'s' if count > 1 else ''}")
            self.log(f"Found {count} file(s) to print", "success")
            self.print_btn.configure(state="normal")
            
        except Exception as e:
            self.log(f"Error scanning folder: {str(e)}", "error")
            messagebox.showerror("Error", f"Failed to scan folder:\n{str(e)}")
    
    def start_printing(self):
        """Start printing in background thread"""
        if not self.print_files:
            messagebox.showwarning("Warning", "No files to print!")
            return
        
        # Confirm
        order_text = "A→Z (Ascending)" if self.sort_ascending.get() else "Z→A (Descending)"
        result = messagebox.askyesno(
            "Confirm",
            f"Print {len(self.print_files)} file(s)?\n\nOrder: {order_text}\nWait time: {self.wait_time.get()} seconds"
        )
        
        if not result:
            return
        
        # Disable buttons
        self.is_printing = True
        self.scan_btn.configure(state="disabled")
        self.print_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.sort_btn.configure(state="disabled")
        
        # Start printing thread
        thread = threading.Thread(target=self.print_worker, daemon=True)
        thread.start()
    
    def print_worker(self):
        """Worker thread for printing"""
        folder = self.folder_path.get()
        wait_time = self.wait_time.get()
        
        try:
            printer_name = win32print.GetDefaultPrinter()
        except Exception as e:
            self.log(f"Error: Cannot get printer - {str(e)}", "error")
            self.after(0, self.printing_complete)
            return
        
        order_text = "ascending (A→Z)" if self.sort_ascending.get() else "descending (Z→A)"
        self.log(f"Starting batch print... ({len(self.print_files)} files, {order_text})", "info")
        
        success_count = 0
        fail_count = 0
        
        for i, filename in enumerate(self.print_files, 1):
            if not self.is_printing:
                self.log("Printing stopped by user.", "warning")
                break
            
            full_path = os.path.join(folder, filename)
            full_path = os.path.normpath(full_path)
            
            ext = filename.split('.')[-1].lower()
            icon = self.get_file_icon(ext)
            
            self.log(f"[{i}/{len(self.print_files)}] {icon} Printing: {filename}")
            self.update_status(f"Printing {i}/{len(self.print_files)}: {filename}")
            
            try:
                win32api.ShellExecute(
                    0,
                    "print",
                    full_path,
                    f'/d:"{printer_name}"',
                    ".",
                    0
                )
                
                self.log(f"✓ Sent to printer successfully", "success")
                success_count += 1
                
                # Wait before next print
                if i < len(self.print_files) and self.is_printing:
                    for sec in range(wait_time, 0, -1):
                        if not self.is_printing:
                            break
                        self.update_status(f"Waiting {sec} seconds...")
                        time.sleep(1)
                
            except Exception as e:
                self.log(f"✗ Error: {str(e)}", "error")
                fail_count += 1
        
        # Summary
        self.log("=" * 50)
        self.log(f"Batch printing completed!", "info")
        self.log(f"Success: {success_count} | Failed: {fail_count}")
        self.log("=" * 50)
        
        self.after(0, self.printing_complete)
        
        # Show completion message
        if success_count > 0:
            self.after(0, lambda: messagebox.showinfo(
                "Complete",
                f"Printing completed!\n\nSuccess: {success_count}\nFailed: {fail_count}"
            ))
    
    def stop_printing(self):
        """Stop printing process"""
        self.is_printing = False
        self.log("Stopping printing...", "warning")
    
    def printing_complete(self):
        """Reset UI after printing"""
        self.is_printing = False
        self.scan_btn.configure(state="normal")
        self.print_btn.configure(state="normal")
        self.stop_btn.configure(state="disabled")
        self.sort_btn.configure(state="normal")
        self.update_status("Ready")
    
    def update_status(self, message):
        """Update status bar (thread-safe)"""
        self.after(0, lambda: self.status_bar.configure(text=f"⚡ Status: {message}"))
    
    def log(self, message, level="info"):
        """Add message to console (thread-safe)"""
        def _log():
            timestamp = datetime.now().strftime("%H:%M:%S")
            
            # Color based on level
            if level == "error":
                prefix = "❌"
            elif level == "success":
                prefix = "✓"
            elif level == "warning":
                prefix = "⚠️"
            else:
                prefix = "ℹ️"
            
            formatted = f"[{timestamp}] {prefix} {message}\n"
            self.console.insert("end", formatted)
            self.console.see("end")
        
        self.after(0, _log)


if __name__ == "__main__":
    app = MultiBatchPrinter()
    app.mainloop()