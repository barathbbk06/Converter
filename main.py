"""
Local Word <-> PDF Converter
GUI Application for Windows (Windows 10/11)
------------------------------------------
A reliable, 100% offline document converter built with Tkinter.
Supports Word (.docx, .doc) -> PDF and PDF -> Word (.docx).
"""

from typing import Optional
import os
import sys
import threading
import subprocess
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Import conversion engine
from converter import (
    convert_word_to_pdf,
    convert_pdf_to_word,
    get_unique_filepath,
    get_default_output_path,
    is_ms_word_available,
    get_libreoffice_path,
    ConversionError
)

# Attempt High-DPI awareness on Windows for crisp rendering
try:
    import ctypes
    # Shcore\SetProcessDpiAwareness(1) or User32\SetProcessDPIAware()
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        ctypes.windll.user32.SetProcessDPIAware()
except Exception:
    pass


class ConflictDialog(tk.Toplevel):
    """
    Custom modal dialog shown when the target output file already exists.
    Asks the user whether to Overwrite, Create Unique Filename, or Cancel.
    """
    def __init__(self, parent, filename: str, unique_filename: str):
        super().__init__(parent)
        self.result = "cancel"  # default if closed

        self.title("File Already Exists")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        # Modern styling
        self.configure(bg="#F8FAFC")

        main_frame = ttk.Frame(self, padding=20)
        main_frame.pack(fill="both", expand=True)

        # Header title
        title_label = ttk.Label(
            main_frame,
            text="Output File Already Exists",
            font=("Segoe UI", 12, "bold"),
            foreground="#0F172A"
        )
        title_label.pack(anchor="w", pady=(0, 8))

        # Description
        desc_text = (
            f"The file '{filename}' already exists in the destination folder.\n\n"
            f"Choose an option:\n"
            f"• Overwrite: Replace the existing file.\n"
            f"• Create Unique Name: Save as '{unique_filename}'."
        )
        msg_label = ttk.Label(
            main_frame,
            text=desc_text,
            font=("Segoe UI", 9),
            foreground="#334155",
            wraplength=380,
            justify="left"
        )
        msg_label.pack(anchor="w", pady=(0, 16))

        # Buttons
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill="x", pady=(5, 0))

        btn_unique = tk.Button(
            btn_frame,
            text="Create Unique Name",
            font=("Segoe UI", 9, "bold"),
            bg="#2563EB",
            fg="white",
            activebackground="#1D4ED8",
            activeforeground="white",
            relief="flat",
            padx=12,
            pady=6,
            cursor="hand2",
            command=self._choose_unique
        )
        btn_unique.pack(side="left", padx=(0, 6))

        btn_overwrite = tk.Button(
            btn_frame,
            text="Overwrite",
            font=("Segoe UI", 9),
            bg="#E2E8F0",
            fg="#0F172A",
            activebackground="#CBD5E1",
            relief="flat",
            padx=12,
            pady=6,
            cursor="hand2",
            command=self._choose_overwrite
        )
        btn_overwrite.pack(side="left", padx=(0, 6))

        btn_cancel = tk.Button(
            btn_frame,
            text="Cancel",
            font=("Segoe UI", 9),
            bg="#F1F5F9",
            fg="#64748B",
            activebackground="#E2E8F0",
            relief="flat",
            padx=12,
            pady=6,
            cursor="hand2",
            command=self._choose_cancel
        )
        btn_cancel.pack(side="left")

        # Center dialog relative to parent
        self.update_idletasks()
        pw = parent.winfo_width()
        ph = parent.winfo_height()
        px = parent.winfo_rootx()
        py = parent.winfo_rooty()
        w = self.winfo_width()
        h = self.winfo_height()
        x = px + max(0, (pw - w) // 2)
        y = py + max(0, (ph - h) // 2)
        self.geometry(f"+{x}+{y}")

        self.protocol("WM_DELETE_WINDOW", self._choose_cancel)
        self.wait_window(self)

    def _choose_unique(self):
        self.result = "unique"
        self.destroy()

    def _choose_overwrite(self):
        self.result = "overwrite"
        self.destroy()

    def _choose_cancel(self):
        self.result = "cancel"
        self.destroy()


class WordPdfConverterApp(tk.Tk):
    """Main Tkinter Application for Local Word <-> PDF Converter."""

    MODE_WORD_TO_PDF = "Word to PDF (.docx → .pdf)"
    MODE_PDF_TO_WORD = "PDF to Word (.pdf → .docx)"

    def __init__(self):
        super().__init__()

        self.title("Local Word ↔ PDF Converter (100% Offline)")
        self.geometry("740x820")
        self.minsize(680, 780)
        self.configure(bg="#F8FAFC")

        # State Variables
        self.conversion_mode = tk.StringVar(value=self.MODE_WORD_TO_PDF)
        self.input_file_path = tk.StringVar()
        self.dest_folder_path = tk.StringVar()
        self.status_message = tk.StringVar(value="Ready. Select an input document to begin.")
        self.last_output_path = ""
        self.is_converting = False

        # Configure TTK Styles
        self._setup_styles()

        # Build UI Components
        self._build_ui()

        # Initialize environment status
        self._detect_system_tools()

    def _setup_styles(self):
        style = ttk.Style(self)
        style.theme_use("clam")

        # Global frame styling
        style.configure("TFrame", background="#F8FAFC")
        style.configure("Card.TFrame", background="#FFFFFF", relief="flat")

        # Labels
        style.configure("Header.TLabel", background="#F8FAFC", foreground="#0F172A", font=("Segoe UI", 16, "bold"))
        style.configure("SubHeader.TLabel", background="#F8FAFC", foreground="#475569", font=("Segoe UI", 9))
        style.configure("Section.TLabel", background="#FFFFFF", foreground="#0F172A", font=("Segoe UI", 10, "bold"))
        style.configure("Regular.TLabel", background="#FFFFFF", foreground="#334155", font=("Segoe UI", 9))
        style.configure("Muted.TLabel", background="#FFFFFF", foreground="#64748B", font=("Segoe UI", 8))
        style.configure("Success.TLabel", background="#ECFDF5", foreground="#065F46", font=("Segoe UI", 9, "bold"))
        style.configure("Error.TLabel", background="#FEF2F2", foreground="#991B1B", font=("Segoe UI", 9, "bold"))

        # Combobox
        style.configure("TCombobox", font=("Segoe UI", 9), padding=5)

        # Progressbar
        style.configure("TProgressbar", thickness=8)

    def _build_ui(self):
        container = ttk.Frame(self, padding=(24, 16, 24, 16))
        container.pack(fill="both", expand=True)

        # 1. Header with Offline Badge
        header_frame = ttk.Frame(container)
        header_frame.pack(fill="x", pady=(0, 14))

        title_lbl = ttk.Label(header_frame, text="Local Word ↔ PDF Converter", style="Header.TLabel")
        title_lbl.pack(anchor="w")

        sub_frame = ttk.Frame(header_frame)
        sub_frame.pack(fill="x", pady=(2, 0))

        sub_lbl = ttk.Label(
            sub_frame,
            text="Fast, private & offline conversion directly on your Windows PC",
            style="SubHeader.TLabel"
        )
        sub_lbl.pack(side="left")

        # Badge: 100% Offline
        badge = tk.Label(
            sub_frame,
            text="● 100% Offline & Private",
            font=("Segoe UI", 8, "bold"),
            bg="#DCFCE7",
            fg="#166534",
            padx=8,
            pady=2,
            relief="flat"
        )
        badge.pack(side="right")

        # 2. Card: Conversion Mode Selection
        card_mode = self._create_card(container)
        card_mode.pack(fill="x", pady=(0, 12))

        ttk.Label(card_mode, text="1. Conversion Type", style="Section.TLabel").pack(anchor="w", pady=(0, 6))

        mode_row = ttk.Frame(card_mode, style="Card.TFrame")
        mode_row.pack(fill="x")

        self.mode_combo = ttk.Combobox(
            mode_row,
            textvariable=self.conversion_mode,
            values=[self.MODE_WORD_TO_PDF, self.MODE_PDF_TO_WORD],
            state="readonly",
            font=("Segoe UI", 9),
            width=35
        )
        self.mode_combo.pack(side="left", padx=(0, 10))
        self.mode_combo.bind("<<ComboboxSelected>>", self._on_mode_changed)

        self.mode_hint = ttk.Label(
            mode_row,
            text="Select conversion direction or choose a file to auto-detect",
            style="Muted.TLabel"
        )
        self.mode_hint.pack(side="left", fill="x")

        # 3. Card: Input File Selection
        card_input = self._create_card(container)
        card_input.pack(fill="x", pady=(0, 12))

        ttk.Label(card_input, text="2. Input File", style="Section.TLabel").pack(anchor="w", pady=(0, 6))

        input_entry_frame = ttk.Frame(card_input, style="Card.TFrame")
        input_entry_frame.pack(fill="x", pady=(0, 4))

        self.input_entry = ttk.Entry(
            input_entry_frame,
            textvariable=self.input_file_path,
            font=("Segoe UI", 9)
        )
        self.input_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.input_file_path.trace_add("write", self._on_input_path_typed)

        browse_input_btn = tk.Button(
            input_entry_frame,
            text="Browse Input File...",
            font=("Segoe UI", 9, "bold"),
            bg="#E2E8F0",
            fg="#0F172A",
            activebackground="#CBD5E1",
            relief="flat",
            padx=12,
            pady=4,
            cursor="hand2",
            command=self._browse_input_file
        )
        browse_input_btn.pack(side="right")

        self.file_details_lbl = ttk.Label(
            card_input,
            text="No file selected. Supports .docx, .doc, and .pdf",
            style="Muted.TLabel"
        )
        self.file_details_lbl.pack(anchor="w")

        # 4. Card: Destination Folder Selection
        card_dest = self._create_card(container)
        card_dest.pack(fill="x", pady=(0, 12))

        dest_header_frame = ttk.Frame(card_dest, style="Card.TFrame")
        dest_header_frame.pack(fill="x", pady=(0, 6))

        ttk.Label(dest_header_frame, text="3. Destination Folder", style="Section.TLabel").pack(side="left")

        # Quick action: "Same as Input Folder"
        same_folder_btn = tk.Button(
            dest_header_frame,
            text="Use Input File's Folder",
            font=("Segoe UI", 8),
            bg="#F1F5F9",
            fg="#2563EB",
            relief="flat",
            cursor="hand2",
            command=self._use_input_folder
        )
        same_folder_btn.pack(side="right")

        dest_entry_frame = ttk.Frame(card_dest, style="Card.TFrame")
        dest_entry_frame.pack(fill="x", pady=(0, 4))

        self.dest_entry = ttk.Entry(
            dest_entry_frame,
            textvariable=self.dest_folder_path,
            font=("Segoe UI", 9)
        )
        self.dest_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))

        browse_dest_btn = tk.Button(
            dest_entry_frame,
            text="Choose Destination...",
            font=("Segoe UI", 9, "bold"),
            bg="#E2E8F0",
            fg="#0F172A",
            activebackground="#CBD5E1",
            relief="flat",
            padx=12,
            pady=4,
            cursor="hand2",
            command=self._browse_dest_folder
        )
        browse_dest_btn.pack(side="right")

        self.dest_details_lbl = ttk.Label(
            card_dest,
            text="Files will be saved in this directory.",
            style="Muted.TLabel"
        )
        self.dest_details_lbl.pack(anchor="w")

        # 5. Card: Conversion Action & Progress
        card_action = self._create_card(container)
        card_action.pack(fill="x", pady=(0, 12))

        action_top = ttk.Frame(card_action, style="Card.TFrame")
        action_top.pack(fill="x", pady=(0, 8))

        self.convert_btn = tk.Button(
            action_top,
            text="▶  Convert Document Now",
            font=("Segoe UI", 10, "bold"),
            bg="#2563EB",
            fg="white",
            activebackground="#1D4ED8",
            activeforeground="white",
            relief="flat",
            padx=20,
            pady=8,
            cursor="hand2",
            command=self._start_conversion
        )
        self.convert_btn.pack(side="left")

        # Status Label
        self.status_lbl = ttk.Label(
            action_top,
            textvariable=self.status_message,
            style="Regular.TLabel"
        )
        self.status_lbl.pack(side="left", padx=(14, 0), fill="x", expand=True)

        # Progress bar
        self.progress_bar = ttk.Progressbar(card_action, mode="indeterminate")
        self.progress_bar.pack(fill="x", pady=(2, 0))

        # 6. Card: Results Banner & Action Links (Initially hidden or neutral)
        self.card_result = ttk.Frame(container, padding=12)
        self.card_result.configure(style="TFrame")
        self.card_result.pack(fill="x", pady=(0, 8))

        self.result_box = tk.Frame(self.card_result, bg="#F1F5F9", relief="flat", bd=1, padx=12, pady=10)
        self.result_box.pack(fill="x")

        self.result_title = tk.Label(
            self.result_box,
            text="Result Status",
            font=("Segoe UI", 9, "bold"),
            bg="#F1F5F9",
            fg="#475569"
        )
        self.result_title.pack(anchor="w")

        self.result_path_text = tk.Text(
            self.result_box,
            height=2,
            font=("Consolas", 8),
            bg="#FFFFFF",
            fg="#0F172A",
            relief="solid",
            bd=1,
            wrap="char"
        )
        self.result_path_text.pack(fill="x", pady=(4, 6))
        self.result_path_text.insert("1.0", "No conversion performed yet.")
        self.result_path_text.config(state="disabled")

        self.result_buttons_frame = tk.Frame(self.result_box, bg="#F1F5F9")
        self.result_buttons_frame.pack(fill="x")

        self.open_file_btn = tk.Button(
            self.result_buttons_frame,
            text="Open Converted File",
            font=("Segoe UI", 8, "bold"),
            bg="#E2E8F0",
            fg="#0F172A",
            relief="flat",
            padx=10,
            pady=3,
            cursor="hand2",
            state="disabled",
            command=self._open_converted_file
        )
        self.open_file_btn.pack(side="left", padx=(0, 6))

        self.open_folder_btn = tk.Button(
            self.result_buttons_frame,
            text="Open Destination Folder",
            font=("Segoe UI", 8),
            bg="#E2E8F0",
            fg="#0F172A",
            relief="flat",
            padx=10,
            pady=3,
            cursor="hand2",
            state="disabled",
            command=self._open_destination_folder
        )
        self.open_folder_btn.pack(side="left", padx=(0, 6))

        self.copy_path_btn = tk.Button(
            self.result_buttons_frame,
            text="Copy Path",
            font=("Segoe UI", 8),
            bg="#E2E8F0",
            fg="#0F172A",
            relief="flat",
            padx=10,
            pady=3,
            cursor="hand2",
            state="disabled",
            command=self._copy_output_path
        )
        self.copy_path_btn.pack(side="left")

        # 7. Environmental Info Footer
        footer_frame = ttk.Frame(container)
        footer_frame.pack(fill="x", side="bottom")

        self.env_lbl = ttk.Label(
            footer_frame,
            text="Checking local conversion backends...",
            style="Muted.TLabel"
        )
        self.env_lbl.pack(side="left")

        info_note_btn = tk.Button(
            footer_frame,
            text="ℹ Scanned PDF & Layout Notes",
            font=("Segoe UI", 8),
            bg="#F8FAFC",
            fg="#2563EB",
            relief="flat",
            cursor="hand2",
            command=self._show_scanned_pdf_info
        )
        info_note_btn.pack(side="right")

    def _create_card(self, parent) -> tk.Frame:
        """Helper to create a modern padded card with clean border."""
        card = tk.Frame(
            parent,
            bg="#FFFFFF",
            highlightbackground="#CBD5E1",
            highlightthickness=1,
            padx=14,
            pady=12
        )
        return card

    def _detect_system_tools(self):
        """Detect available backends in background so startup remains instant."""
        def check():
            word_ok, word_info = is_ms_word_available()
            lo_path = get_libreoffice_path()

            status_parts = []
            if word_ok:
                status_parts.append(f"Word Engine: {word_info}")
            elif lo_path:
                status_parts.append("Word Engine: LibreOffice Headless")
            else:
                status_parts.append("Word Engine: Neither Word nor LibreOffice found")

            status_parts.append("PDF Engine: pdf2docx (Local)")

            text = " | ".join(status_parts)
            self.after(0, lambda: self.env_lbl.config(text=text))

        threading.Thread(target=check, daemon=True).start()

    def _on_mode_changed(self, event=None):
        mode = self.conversion_mode.get()
        if mode == self.MODE_WORD_TO_PDF:
            self.mode_hint.config(text="Converting Word document to PDF format (.pdf)")
        else:
            self.mode_hint.config(text="Extracting PDF content to editable Word format (.docx)")
        self._validate_input_ui()

    def _on_input_path_typed(self, *args):
        path = self.input_file_path.get().strip().strip('"')
        if not path:
            self.file_details_lbl.config(text="No file selected. Supports .docx, .doc, and .pdf")
            return

        if os.path.isfile(path):
            size_kb = os.path.getsize(path) / 1024
            size_str = f"{size_kb:.1f} KB" if size_kb < 1024 else f"{size_kb/1024:.2f} MB"
            ext = Path(path).suffix.lower()

            # Auto-detect conversion mode if clear
            if ext in [".docx", ".doc"] and self.conversion_mode.get() != self.MODE_WORD_TO_PDF:
                self.conversion_mode.set(self.MODE_WORD_TO_PDF)
                self.mode_hint.config(text="Auto-detected Word document. Mode set to Word -> PDF.")
            elif ext == ".pdf" and self.conversion_mode.get() != self.MODE_PDF_TO_WORD:
                self.conversion_mode.set(self.MODE_PDF_TO_WORD)
                self.mode_hint.config(text="Auto-detected PDF document. Mode set to PDF -> Word.")

            self.file_details_lbl.config(text=f"✓ Valid file: {Path(path).name} ({size_str})")

            # Default destination directory to the input directory if not set yet
            if not self.dest_folder_path.get().strip():
                self.dest_folder_path.set(os.path.dirname(os.path.abspath(path)))
        else:
            self.file_details_lbl.config(text="File path does not exist or is inaccessible.")

    def _browse_input_file(self):
        curr_mode = self.conversion_mode.get()
        if curr_mode == self.MODE_WORD_TO_PDF:
            filetypes = [
                ("Word Documents (*.docx;*.doc)", "*.docx;*.doc"),
                ("PDF Documents (*.pdf)", "*.pdf"),
                ("All Files (*.*)", "*.*")
            ]
        else:
            filetypes = [
                ("PDF Documents (*.pdf)", "*.pdf"),
                ("Word Documents (*.docx;*.doc)", "*.docx;*.doc"),
                ("All Files (*.*)", "*.*")
            ]

        chosen = filedialog.askopenfilename(
            title="Select Input File to Convert",
            filetypes=filetypes
        )
        if chosen:
            self.input_file_path.set(os.path.abspath(chosen))
            # Auto-fill destination folder if empty
            if not self.dest_folder_path.get().strip():
                self.dest_folder_path.set(os.path.dirname(os.path.abspath(chosen)))

    def _browse_dest_folder(self):
        initial = self.dest_folder_path.get().strip() or os.path.expanduser("~")
        chosen = filedialog.askdirectory(
            title="Select Destination Folder",
            initialdir=initial
        )
        if chosen:
            self.dest_folder_path.set(os.path.abspath(chosen))

    def _use_input_folder(self):
        in_path = self.input_file_path.get().strip().strip('"')
        if in_path:
            parent = os.path.dirname(os.path.abspath(in_path))
            self.dest_folder_path.set(parent)
        else:
            messagebox.showinfo("Notice", "Please select an input file first.")

    def _validate_input_ui(self) -> bool:
        in_path = self.input_file_path.get().strip().strip('"')
        if not in_path:
            return False
        return os.path.isfile(in_path)

    def _start_conversion(self):
        if self.is_converting:
            return

        in_path = self.input_file_path.get().strip().strip('"')
        dest_folder = self.dest_folder_path.get().strip().strip('"')
        mode = self.conversion_mode.get()

        # Validation 1: Input file
        if not in_path:
            messagebox.showerror("Missing Input File", "Please select or enter the path of an input file.")
            return

        if not os.path.isfile(in_path):
            messagebox.showerror("File Not Found", f"The input file does not exist:\n\n{in_path}")
            return

        ext = Path(in_path).suffix.lower()
        if mode == self.MODE_WORD_TO_PDF and ext not in [".docx", ".doc"]:
            messagebox.showerror(
                "Invalid Format",
                f"Selected mode is 'Word to PDF', but the chosen file has extension '{ext}'.\n"
                "Please select a .docx or .doc file, or switch conversion mode."
            )
            return

        if mode == self.MODE_PDF_TO_WORD and ext != ".pdf":
            messagebox.showerror(
                "Invalid Format",
                f"Selected mode is 'PDF to Word', but the chosen file has extension '{ext}'.\n"
                "Please select a .pdf file, or switch conversion mode."
            )
            return

        # Validation 2: Destination Folder
        if not dest_folder:
            # Default to input file's folder
            dest_folder = os.path.dirname(os.path.abspath(in_path))
            self.dest_folder_path.set(dest_folder)

        try:
            os.makedirs(dest_folder, exist_ok=True)
        except Exception as e:
            messagebox.showerror("Folder Error", f"Cannot create destination folder:\n\n{e}")
            return

        # Target output filename calculation
        target_ext = "pdf" if mode == self.MODE_WORD_TO_PDF else "docx"
        target_path = get_default_output_path(in_path, dest_folder, target_ext)

        # Check for destination file collision
        final_output_path = target_path
        if os.path.exists(target_path):
            unique_candidate = get_unique_filepath(target_path)
            dlg = ConflictDialog(self, os.path.basename(target_path), os.path.basename(unique_candidate))
            if dlg.result == "cancel":
                self.status_message.set("Conversion canceled by user.")
                return
            elif dlg.result == "unique":
                final_output_path = unique_candidate
            elif dlg.result == "overwrite":
                final_output_path = target_path

        # Start Conversion Thread
        self.is_converting = True
        self.convert_btn.config(state="disabled", text="⏳ Converting...")
        self.progress_bar.start(10)
        self.status_message.set("Conversion in progress. Please wait...")

        # Clear previous result UI
        self._update_result_ui_pending()

        thread = threading.Thread(
            target=self._run_conversion_worker,
            args=(in_path, final_output_path, mode),
            daemon=True
        )
        thread.start()

    def _run_conversion_worker(self, in_path: str, out_path: str, mode: str):
        error_msg = None
        try:
            if mode == self.MODE_WORD_TO_PDF:
                convert_word_to_pdf(in_path, out_path)
            else:
                convert_pdf_to_word(in_path, out_path)
        except Exception as e:
            error_msg = str(e)

        self.after(0, lambda: self._on_conversion_finished(out_path, error_msg))

    def _on_conversion_finished(self, out_path: str, error_msg: Optional[str]):
        self.is_converting = False
        self.progress_bar.stop()
        self.convert_btn.config(state="normal", text="▶  Convert Document Now")

        if error_msg is None and os.path.isfile(out_path):
            # Success
            self.last_output_path = out_path
            self.status_message.set("✓ Conversion completed successfully!")
            self._update_result_ui_success(out_path)
            messagebox.showinfo("Success", f"Document converted successfully!\n\nSaved to:\n{out_path}")
        else:
            # Error
            self.last_output_path = ""
            self.status_message.set("✗ Conversion failed.")
            self._update_result_ui_error(error_msg or "Unknown error occurred.")
            messagebox.showerror("Conversion Failed", f"An error occurred during conversion:\n\n{error_msg}")

    def _update_result_ui_pending(self):
        self.result_box.configure(bg="#F1F5F9")
        self.result_title.configure(text="Status: Converting...", bg="#F1F5F9", fg="#475569")
        self.result_path_text.config(state="normal")
        self.result_path_text.delete("1.0", tk.END)
        self.result_path_text.insert("1.0", "Processing document...")
        self.result_path_text.config(state="disabled")
        self.open_file_btn.config(state="disabled")
        self.open_folder_btn.config(state="disabled")
        self.copy_path_btn.config(state="disabled")

    def _update_result_ui_success(self, out_path: str):
        self.result_box.configure(bg="#ECFDF5")
        self.result_title.configure(text="✓ Conversion Succeeded", bg="#ECFDF5", fg="#065F46")
        self.result_path_text.config(state="normal")
        self.result_path_text.delete("1.0", tk.END)
        self.result_path_text.insert("1.0", out_path)
        self.result_path_text.config(state="disabled")
        self.open_file_btn.config(state="normal")
        self.open_folder_btn.config(state="normal")
        self.copy_path_btn.config(state="normal")

    def _update_result_ui_error(self, err_text: str):
        self.result_box.configure(bg="#FEF2F2")
        self.result_title.configure(text="✗ Conversion Error", bg="#FEF2F2", fg="#991B1B")
        self.result_path_text.config(state="normal")
        self.result_path_text.delete("1.0", tk.END)
        self.result_path_text.insert("1.0", err_text)
        self.result_path_text.config(state="disabled")
        self.open_file_btn.config(state="disabled")
        self.open_folder_btn.config(state="disabled")
        self.copy_path_btn.config(state="disabled")

    def _open_converted_file(self):
        if self.last_output_path and os.path.isfile(self.last_output_path):
            try:
                os.startfile(self.last_output_path)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to open file:\n{e}")

    def _open_destination_folder(self):
        if self.last_output_path and os.path.isfile(self.last_output_path):
            folder = os.path.dirname(self.last_output_path)
            try:
                # Open Explorer and select the converted file
                subprocess.run(["explorer", f"/select,{self.last_output_path}"])
            except Exception:
                try:
                    os.startfile(folder)
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to open folder:\n{e}")
        else:
            dest = self.dest_folder_path.get().strip()
            if dest and os.path.isdir(dest):
                os.startfile(dest)

    def _copy_output_path(self):
        if self.last_output_path:
            self.clipboard_clear()
            self.clipboard_append(self.last_output_path)
            self.status_message.set("✓ Output path copied to clipboard!")

    def _show_scanned_pdf_info(self):
        info_text = (
            "About Document Formats & Scanned PDFs\n"
            "------------------------------------\n\n"
            "1. Word to PDF (.docx -> .pdf):\n"
            "• Uses your local Microsoft Word installation (via pywin32) for 100% native layout fidelity.\n"
            "• If Microsoft Word is not installed, LibreOffice headless mode is used as a fallback.\n\n"
            "2. PDF to Word (.pdf -> .docx):\n"
            "• Converts vector text, paragraphs, tables, images, and alignments into editable Word elements.\n\n"
            "3. Scanned PDFs (Important Limitation):\n"
            "• Scanned PDFs consist of raw photographic/scanned images of pages without an underlying text layer.\n"
            "• When converted, the scanned images are preserved inside the Word document. To extract editable text from scanned images, optical character recognition (OCR) software like Tesseract is required.\n\n"
            "4. Privacy & Offline Guarantee:\n"
            "• All operations execute 100% locally on your PC. No files, metadata, or telemetry are ever sent over the network."
        )
        messagebox.showinfo("Document Conversion & Scanned PDF Info", info_text)


def main():
    app = WordPdfConverterApp()
    app.mainloop()


if __name__ == "__main__":
    main()
