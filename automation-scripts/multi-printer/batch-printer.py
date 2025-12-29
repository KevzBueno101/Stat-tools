import customtkinter as ctk
import os
import win32print
import win32api
from datetime import datetime
import threading

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


class MultiBatchPrinter(ctk.CTk):

    def __init__(self):
        super().__init__()

        self.title("Multi-Format Batch Printer Pro")
        self.geometry("1100x800")
        self.minsize(900, 600)

        # ===== SCROLLABLE ROOT =====
        self.scroll = ctk.CTkScrollableFrame(self)
        self.scroll.pack(fill="both", expand=True)

        # Build UI
        self.create_header(self.scroll)
        self.create_two_column_layout(self.scroll)
        self.create_status_bar(self.scroll)

        # Printer info
        self.update_printer_info()

    # =====================================================
    # HEADER
    # =====================================================
    def create_header(self, parent):
        header_frame = ctk.CTkFrame(parent, corner_radius=10)
        header_frame.pack(fill="x", padx=20, pady=(20, 10))

        title = ctk.CTkLabel(
            header_frame,
            text="📄 Multi-Format Batch Printer Pro",
            font=ctk.CTkFont(size=28, weight="bold")
        )
        title.pack(anchor="w", padx=20, pady=(15, 5))

        subtitle = ctk.CTkLabel(
            header_frame,
            text="Batch printing utility for PDF, DOCX, XLSX, TXT files",
            font=ctk.CTkFont(size=14)
        )
        subtitle.pack(anchor="w", padx=20, pady=(0, 15))

    # =====================================================
    # MAIN CONTENT
    # =====================================================
    def create_two_column_layout(self, parent):
        container = ctk.CTkFrame(parent, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=20, pady=10)

        container.grid_columnconfigure(0, weight=1)
        container.grid_columnconfigure(1, weight=1)

        # LEFT PANEL
        left = ctk.CTkFrame(container, corner_radius=10)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        ctk.CTkLabel(
            left, text="📂 File Selection",
            font=ctk.CTkFont(size=18, weight="bold")
        ).pack(anchor="w", padx=15, pady=10)

        self.folder_label = ctk.CTkLabel(left, text="No folder selected")
        self.folder_label.pack(anchor="w", padx=15, pady=5)

        ctk.CTkButton(
            left, text="Select Folder",
            command=self.select_folder
        ).pack(padx=15, pady=10)

        ctk.CTkButton(
            left, text="Start Printing",
            fg_color="#2a9d8f",
            command=self.start_printing
        ).pack(padx=15, pady=20)

        # RIGHT PANEL
        right = ctk.CTkFrame(container, corner_radius=10)
        right.grid(row=0, column=1, sticky="nsew", padx=(10, 0))

        ctk.CTkLabel(
            right, text="🖨 Printer Console",
            font=ctk.CTkFont(size=18, weight="bold")
        ).pack(anchor="w", padx=15, pady=10)

        self.console = ctk.CTkTextbox(
            right, height=300, font=("Consolas", 11)
        )
        self.console.pack(fill="both", expand=True, padx=15, pady=10)

    # =====================================================
    # STATUS BAR
    # =====================================================
    def create_status_bar(self, parent):
        status = ctk.CTkFrame(parent, corner_radius=10)
        status.pack(fill="x", padx=20, pady=(10, 20))

        self.printer_label = ctk.CTkLabel(
            status, text="Printer: Detecting..."
        )
        self.printer_label.pack(anchor="w", padx=15, pady=10)

    # =====================================================
    # LOGIC
    # =====================================================
    def update_printer_info(self):
        try:
            printer = win32print.GetDefaultPrinter()
            self.printer_label.configure(text=f"Printer: {printer}")
        except:
            self.printer_label.configure(text="Printer: Not detected")

    def select_folder(self):
        from tkinter import filedialog
        self.folder = filedialog.askdirectory()
        if self.folder:
            self.folder_label.configure(text=self.folder)
            self.log(f"Folder selected: {self.folder}")

    def start_printing(self):
        if not hasattr(self, "folder"):
            self.log("No folder selected", "warning")
            return

        threading.Thread(target=self.batch_print, daemon=True).start()

    def batch_print(self):
        for file in os.listdir(self.folder):
            path = os.path.join(self.folder, file)
            if os.path.isfile(path):
                try:
                    win32api.ShellExecute(
                        0, "print", path, None, ".", 0
                    )
                    self.log(f"Printing: {file}", "success")
                except Exception as e:
                    self.log(f"Failed: {file} → {e}", "error")

    def log(self, message, level="info"):
        def _log():
            timestamp = datetime.now().strftime("%H:%M:%S")
            prefix = {
                "success": "✅",
                "error": "❌",
                "warning": "⚠️"
            }.get(level, "ℹ️")

            self.console.insert(
                "end", f"[{timestamp}] {prefix} {message}\n"
            )
            self.console.see("end")

        self.after(0, _log)


if __name__ == "__main__":
    app = MultiBatchPrinter()
    app.mainloop()
