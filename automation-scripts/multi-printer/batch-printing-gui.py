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


class PDFBatchPrinter(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Window configuration
        self.title("📄 PDF Batch Printer Pro")
        self.geometry("1100x700")
        self.minsize(1000, 650)
        
        # Variables
        self.folder_path = ctk.StringVar(value="No folder selected")
        self.wait_time = ctk.IntVar(value=5)
        self.pdf_files = []
        self.is_printing = False
        
        # Create UI
        self.create_header()
        self.create_two_column_layout()
        self.create_status_bar()
        
        # Get default printer on startup
        self.update_printer_info()
    
    def create_header(self):
        """Create header with title and description"""
        header_frame = ctk.CTkFrame(self, corner_radius=10)
        header_frame.pack(fill="x", padx=20, pady=(20, 10))
        
        title = ctk.CTkLabel(
            header_frame,
            text="📄 PDF Batch Printer Pro",
            font=ctk.CTkFont(size=28, weight="bold")
        )
        title.pack(pady=(15, 5))
        
        subtitle = ctk.CTkLabel(
            header_frame,
            text="Print multiple PDF files with ease - Two Column Layout",
            font=ctk.CTkFont(size=14),
            text_color="gray"
        )
        subtitle.pack(pady=(0, 15))
    
    def create_two_column_layout(self):
        """Create two-column layout"""
        # Main container for columns
        columns_frame = ctk.CTkFrame(self, fg_color="transparent")
        columns_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # LEFT COLUMN
        left_column = ctk.CTkFrame(columns_frame, fg_color="transparent")
        left_column.pack(side="left", fill="both", expand=True, padx=(0, 10))
        
        self.create_folder_section(left_column)
        self.create_settings_section(left_column)
        self.create_control_buttons(left_column)
        
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
    
    def create_control_buttons(self, parent):
        """Create control buttons"""
        button_frame = ctk.CTkFrame(parent, corner_radius=10)
        button_frame.pack(fill="x", pady=(0, 10))
        
        label = ctk.CTkLabel(
            button_frame,
            text="🎮 Controls:",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        label.pack(anchor="w", padx=15, pady=(15, 10))
        
        self.scan_btn = ctk.CTkButton(
            button_frame,
            text="🔍 Scan Folder",
            command=self.scan_folder,
            height=40,
            font=ctk.CTkFont(size=13, weight="bold")
        )
        self.scan_btn.pack(fill="x", padx=15, pady=(0, 8))
        
        self.print_btn = ctk.CTkButton(
            button_frame,
            text="🖨️ Start Printing",
            command=self.start_printing,
            height=40,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color="#2fa572",
            hover_color="#28865e"
        )
        self.print_btn.pack(fill="x", padx=15, pady=(0, 8))
        self.print_btn.configure(state="disabled")
        
        self.stop_btn = ctk.CTkButton(
            button_frame,
            text="⏹️ Stop Printing",
            command=self.stop_printing,
            height=40,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color="#c42b1c",
            hover_color="#a21a0e"
        )
        self.stop_btn.pack(fill="x", padx=15, pady=(0, 15))
        self.stop_btn.configure(state="disabled")
    
    def create_file_list_section(self, parent):
        """Create file list section"""
        list_frame = ctk.CTkFrame(parent, corner_radius=10)
        list_frame.pack(fill="both", expand=True, pady=(0, 10))
        
        # Header
        header = ctk.CTkFrame(list_frame, fg_color="transparent")
        header.pack(fill="x", padx=15, pady=(15, 10))
        
        label = ctk.CTkLabel(
            header,
            text="📋 PDF Files:",
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
        self.log("Ready to print PDF files.")
    
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
    
    def browse_folder(self):
        """Open folder browser dialog"""
        folder = filedialog.askdirectory(title="Select folder containing PDF files")
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
    
    def scan_folder(self):
        """Scan folder for PDF files"""
        folder = self.folder_path.get()
        
        if folder == "No folder selected":
            messagebox.showwarning("Warning", "Please select a folder first!")
            return
        
        if not os.path.exists(folder):
            messagebox.showerror("Error", "Folder does not exist!")
            return
        
        self.log("Scanning folder for PDF files...")
        self.file_textbox.delete("1.0", "end")
        
        # Get PDF files
        try:
            all_files = os.listdir(folder)
            self.pdf_files = sorted([f for f in all_files if f.lower().endswith(".pdf")])
            
            if not self.pdf_files:
                self.log("No PDF files found in folder.", "warning")
                self.file_count_label.configure(text="0 files")
                self.print_btn.configure(state="disabled")
                return
            
            # Display files
            for i, filename in enumerate(self.pdf_files, 1):
                full_path = os.path.join(folder, filename)
                file_size = os.path.getsize(full_path)
                size_kb = file_size / 1024
                
                self.file_textbox.insert("end", f"{i}. {filename}\n")
                self.file_textbox.insert("end", f"   Size: {size_kb:.1f} KB\n\n")
            
            count = len(self.pdf_files)
            self.file_count_label.configure(text=f"{count} file{'s' if count > 1 else ''}")
            self.log(f"Found {count} PDF file(s)", "success")
            self.print_btn.configure(state="normal")
            
        except Exception as e:
            self.log(f"Error scanning folder: {str(e)}", "error")
            messagebox.showerror("Error", f"Failed to scan folder:\n{str(e)}")
    
    def start_printing(self):
        """Start printing in background thread"""
        if not self.pdf_files:
            messagebox.showwarning("Warning", "No PDF files to print!")
            return
        
        # Confirm
        result = messagebox.askyesno(
            "Confirm",
            f"Print {len(self.pdf_files)} PDF file(s)?\n\nWait time: {self.wait_time.get()} seconds"
        )
        
        if not result:
            return
        
        # Disable buttons
        self.is_printing = True
        self.scan_btn.configure(state="disabled")
        self.print_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        
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
        
        self.log(f"Starting batch print... ({len(self.pdf_files)} files)", "info")
        
        success_count = 0
        fail_count = 0
        
        for i, filename in enumerate(self.pdf_files, 1):
            if not self.is_printing:
                self.log("Printing stopped by user.", "warning")
                break
            
            full_path = os.path.join(folder, filename)
            full_path = os.path.normpath(full_path)
            
            self.log(f"[{i}/{len(self.pdf_files)}] Printing: {filename}")
            self.update_status(f"Printing {i}/{len(self.pdf_files)}: {filename}")
            
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
                if i < len(self.pdf_files) and self.is_printing:
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
    app = PDFBatchPrinter()
    app.mainloop()