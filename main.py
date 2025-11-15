import os
import string
import tempfile
import threading
import tkinter as tk
from collections import defaultdict
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import List, Tuple

import openpyxl

from core.excel_io import (
    write_rules_sheet_openpyxl,
    write_summary_section_openpyxl,
    write_transactions_sheet_openpyxl,
)
from core.translations import get_translation
from core.utils import decrypt_pdf, load_rules, pdf_to_text
from core.parsers import BaseParser, Transaction, registry


def process_pdf_to_excel(
    pdf_path: str,
    parser_instance: BaseParser,
    rules: List[Tuple[str, str]],
    output_path: str,
    existing_excel: str = None,
    sheet_name: str = "Tranzactii",
    language: str = "en",
) -> Tuple[bool, str]:
    """Process PDF and create/update Excel file with transactions"""
    transactions = parser_instance.parse_pdf(pdf_path)

    if not transactions:
        return False, get_translation("no_transactions_found", language)

    # Parsed transactions are Transaction objects and will be used directly.
    if existing_excel and os.path.exists(existing_excel):
        workbook = openpyxl.open(existing_excel)
    else:
        workbook = openpyxl.Workbook()

    rate_buckets, cheltuieli, rate_noi = compute_summary(transactions)

    columns = parser_instance.get_columns(language)
    write_transactions_sheet_openpyxl(
        workbook, sheet_name, columns, transactions, rules, language
    )

    if rules:
        write_rules_sheet_openpyxl(workbook, rules, language)

    # summary: write to the transactions worksheet
    trans_ws = workbook[sheet_name]
    write_summary_section_openpyxl(
        trans_ws,
        [
            {"months": k + 1, "sum": v}
            for k, v in sorted(rate_buckets.items(), key=lambda x: x[0])
        ],
        len(columns) + 3,
        language,
    )

    workbook.save(output_path)
    return True, get_translation("successfully_created", language).format(output_path)


def compute_summary(transactions: List[Transaction]) -> Tuple[dict, float, float]:
    """Compute summary from a list of Transaction objects or mapping-like records.

    Returns (rate_buckets, cheltuieli, rate_noi)
    """
    rate_buckets: dict[int, float] = defaultdict(int)
    cheltuieli = 0.0
    rate_noi = 0.0

    for tx in transactions:
        amount = tx.amount

        if not tx.installment:
            cheltuieli += amount
            continue

        rata_nr = tx.installment
        rata_total = tx.installment_count

        rate_buckets[rata_total - rata_nr] += amount

        if rata_nr == 1:
            total_tr = tx.total_transaction
            try:
                rate_noi += float(total_tr or 0)
            except Exception:
                pass

    return rate_buckets, cheltuieli, rate_noi


class ParserGUI:
    """Tkinter GUI for PDF parser application"""

    __readable_pdf_path: Path = None

    def __init__(self):
        self.root = tk.Tk()
        self.root.title(get_translation("app_title", "en"))
        self.root.geometry("600x500")

        # Initialize variables
        self.__init_variables()

        # Available parsers
        self.parsers = registry.get_parsers()

        self.__setup_ui()

    def __init_variables(self):
        """Initialize all GUI variables"""
        self.pdf_path = tk.StringVar()
        self.excel_path = tk.StringVar()
        self.sheet_name_var = tk.StringVar(value="Tranzactii")
        self.output_path = tk.StringVar()
        self.selected_parser = tk.StringVar()
        self.language_var = tk.StringVar(value="ro")
        # whether the user wants to select an existing Excel file (otherwise they pick an output path)
        self.use_existing_excel = tk.BooleanVar(value=True)
        # rule editing removed; rules come from the workbook (Rules sheet) or rules.csv fallback

    def update_ui_language(self):
        """Update UI elements with current language"""
        current_lang = self.language_var.get()
        self.root.title(get_translation("app_title", current_lang))
        # Walk the widget tree and update any widgets that declare a trans_key or trans_heading
        for widget in self.root.winfo_children():
            self.__update_widget_text(widget, current_lang)

    def __update_widget_text(self, widget, language):
        """Recursively update widget text based on language"""
        # First update this widget if it has an explicit translation key
        try:
            if hasattr(widget, "trans_key"):
                widget.config(text=get_translation(widget.trans_key, language))
        except Exception:
            pass

        # Update treeview headings if provided
        if isinstance(widget, ttk.Treeview) and hasattr(widget, "trans_heading"):
            try:
                for col, key in getattr(widget, "trans_heading").items():
                    widget.heading(col, text=get_translation(key, language))
            except Exception:
                pass

        # Recurse into children
        if hasattr(widget, "winfo_children"):
            for child in widget.winfo_children():
                self.__update_widget_text(child, language)

    def __setup_ui(self):
        """Setup the user interface"""
        self.root.tk.call("source", "./forest-theme/forest-light.tcl")
        ttk.Style().theme_use("forest-light")

        # Main frame (card)
        main_frame = ttk.Frame(self.root, style="Card", padding=16)
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

        row = 0

        # Header
        header = ttk.Label(
            main_frame,
            text=get_translation("app_title", self.language_var.get()),
        )
        header.trans_key = "app_title"
        header.grid(
            row=row, column=0, columnspan=3, sticky=(tk.W), pady=(0, 10), padx=(10, 10)
        )
        # row += 1

        # # Language Selection
        # ttk.Label(main_frame, text="Language:").grid(
        #     row=row, column=0, sticky=tk.W, pady=4
        # )
        language_combo = ttk.Combobox(
            main_frame,
            textvariable=self.language_var,
            values=["en", "ro"],
            state="readonly",
        )
        language_combo.grid(row=row, column=1, sticky=(tk.E))
        # language_combo.pack(side=tk.BOTTOM)
        language_combo.bind("<<ComboboxSelected>>", lambda e: self.update_ui_language())
        row += 1

        # Parser Selection
        parser_label = ttk.Label(
            main_frame, text=get_translation("parser", self.language_var.get())
        )
        parser_label.trans_key = "parser"
        parser_label.grid(row=row, column=0, sticky=tk.W, pady=5)
        parser_combo = ttk.Combobox(
            main_frame,
            textvariable=self.selected_parser,
            values=list(self.parsers.keys()),
            state="readonly",
        )
        parser_combo.grid(row=row, column=1, sticky=(tk.W, tk.E), pady=5)
        if self.parsers:
            parser_combo.set(list(self.parsers.keys())[0])
        row += 1

        # PDF File Selection
        pdf_label = ttk.Label(
            main_frame, text=get_translation("pdf_file", self.language_var.get())
        )
        pdf_label.trans_key = "pdf_file"
        pdf_label.grid(row=row, column=0, sticky=tk.W, pady=4)
        pdf_frame = ttk.Frame(main_frame)
        pdf_frame.grid(row=row, column=1, sticky=(tk.W, tk.E), pady=5)
        pdf_frame.columnconfigure(0, weight=1)
        ttk.Entry(pdf_frame, textvariable=self.pdf_path, state="readonly").grid(
            row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5)
        )
        ttk.Button(
            pdf_frame,
            text=get_translation("browse", self.language_var.get()),
            command=self.browse_pdf,
        ).grid(row=0, column=1)
        row += 1

        # Toggle: user chooses whether to use an existing Excel file or create a new one
        cb = ttk.Checkbutton(
            main_frame,
            text=get_translation("select_existing_excel", self.language_var.get()),
            variable=self.use_existing_excel,
            command=self._update_output_visibility,
        )
        cb.grid(row=row, column=0, sticky=tk.W, pady=4)
        row += 1

        # Excel File Selection (Optional) - user can also specify an output file to start from scratch
        excel_label = ttk.Label(
            main_frame, text=get_translation("excel_file", self.language_var.get())
        )
        excel_label.trans_key = "excel_file"
        excel_label.grid(row=row, column=0, sticky=tk.W, pady=4)
        self.excel_label = excel_label

        excel_frame = ttk.Frame(main_frame)
        self.excel_frame = excel_frame
        excel_frame.grid(row=row, column=1, sticky=(tk.W, tk.E), pady=5)
        excel_frame.columnconfigure(0, weight=1)
        ttk.Entry(excel_frame, textvariable=self.excel_path, state="readonly").grid(
            row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5)
        )
        ttk.Button(
            excel_frame,
            text=get_translation("browse", self.language_var.get()),
            command=self.browse_excel,
        ).grid(row=0, column=1)
        ttk.Button(
            excel_frame,
            text=get_translation("remove_selected", self.language_var.get()),
            command=self.clear_excel,
        ).grid(row=0, column=2, padx=(5, 0))
        row += 1

        # Output file selection (when not using an existing workbook)
        self.output_label = ttk.Label(
            main_frame, text=get_translation("output_file", self.language_var.get())
        )
        self.output_label.trans_key = "output_file"
        self.output_label.grid(row=row, column=0, sticky=tk.W, pady=4)

        self.output_frame = ttk.Frame(main_frame)
        self.output_frame.grid(row=row, column=1, sticky=(tk.W, tk.E), pady=5)
        self.output_frame.columnconfigure(0, weight=1)
        ttk.Entry(self.output_frame, textvariable=self.output_path).grid(
            row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5)
        )
        ttk.Button(
            self.output_frame,
            text=get_translation("browse", self.language_var.get()),
            command=self.browse_output,
        ).grid(row=0, column=1)

        # visibility is toggled by _update_output_visibility
        row += 1
        # initialize visibility based on the checkbox
        self._update_output_visibility()

        # Rules UI removed: rules are read automatically from the workbook 'Rules' sheet or rules.csv
        # (no in-GUI editing)
        row += 1
        # Output file widgets are handled by _update_output_visibility() elsewhere
        sep = ttk.Separator(main_frame, orient="horizontal")
        sep.grid(row=row, column=0, columnspan=3, sticky=(tk.E, tk.W), pady=(8, 8))
        row += 1
        lbl_sheet = ttk.Label(
            main_frame, text=get_translation("sheet_name", self.language_var.get())
        )
        lbl_sheet.trans_key = "sheet_name"
        lbl_sheet.grid(row=row, column=0, sticky=tk.W, pady=4)
        sheet_frame = ttk.Frame(main_frame)
        sheet_frame.grid(row=row, column=1, sticky=(tk.W, tk.E), pady=5)
        sheet_frame.columnconfigure(0, weight=1)
        ttk.Entry(sheet_frame, textvariable=self.sheet_name_var).grid(
            row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5)
        )
        row += 1
        # Clear-existing option removed
        row += 1
        self.process_btn = ttk.Button(
            main_frame,
            text=get_translation("process_pdf", self.language_var.get()),
            command=self.process_pdf,
        )
        self.process_btn.trans_key = "process_pdf"
        self.process_btn.grid(
            row=row, column=0, columnspan=2, pady=14, ipadx=20, ipady=6
        )
        row += 1
        lbl_status = ttk.Label(
            main_frame, text=get_translation("status", self.language_var.get())
        )
        lbl_status.trans_key = "status"
        lbl_status.grid(row=row, column=0, sticky=(tk.W, tk.N), pady=5)
        row += 1
        text_frame = ttk.Frame(main_frame)
        text_frame.grid(
            row=row, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5
        )
        text_frame.columnconfigure(0, weight=1)
        text_frame.rowconfigure(0, weight=1)
        self.status_text = tk.Text(
            text_frame, height=15, wrap=tk.WORD, font=("Consolas", 10)
        )
        scrollbar = ttk.Scrollbar(
            text_frame, orient=tk.VERTICAL, command=self.status_text.yview
        )
        self.status_text.configure(yscrollcommand=scrollbar.set)
        self.status_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.status_text.config(state="disabled")
        main_frame.rowconfigure(row, weight=1)

    def browse_pdf(self):
        """Browse for PDF file"""
        filename = filedialog.askopenfilename(
            title="Select PDF File",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if filename:

            def background_proc():
                self.pdf_path.set(filename)
                pdf_path = Path(filename)
                self.sheet_name_var.set(
                    pdf_path.name.replace(".pdf", "").strip(string.digits)
                )
                # Auto-detect parser
                detected_parser = registry.auto_detect_parser(filename)
                if detected_parser:
                    self.selected_parser.set(detected_parser)
                    self.log_message(
                        f"{get_translation('auto_detected_parser', self.language_var.get())} {detected_parser}"
                    )
                else:
                    self.log_message(
                        get_translation(
                            "could_not_auto_detect", self.language_var.get()
                        )
                    )

            threading.Thread(target=background_proc).start()

    def browse_excel(self):
        """Browse for existing Excel file"""
        filename = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if filename:
            self.excel_path.set(filename)
            # Auto-set output path to the same file
            self.output_path.set(filename)
            self._update_output_visibility()

    def _update_output_visibility(self):
        """Show/hide the existing-file vs output-file widgets based on the checkbox state.

        If the user chooses to use an existing Excel file, show `excel_frame` and hide
        the output file widgets. Otherwise show the output widgets and hide the
        excel chooser.
        """
        use_existing = bool(self.use_existing_excel.get())

        if use_existing:
            to_remove = ["output_label", "output_frame"]
            to_show = ["excel_frame", "excel_label"]
        else:
            to_show = ["output_label", "output_frame"]
            to_remove = ["excel_frame", "excel_label"]

        for i in to_remove:
            try:
                getattr(self, i).grid_remove()
            except Exception:
                pass

        for i in to_show:
            try:
                getattr(self, i).grid()
            except Exception:
                pass

    def clear_excel(self):
        """Clear Excel file selection"""
        self.excel_path.set("")
        self.output_path.set("")
        self._update_output_visibility()

    def browse_output(self):
        """Browse for output file location"""
        filename = filedialog.asksaveasfilename(
            title="Save Output As",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if filename:
            self.output_path.set(filename)

    def log_message(self, message):
        """Add message to status log"""
        # enable temporarily to insert message
        self.status_text.config(state="normal")
        self.status_text.insert(tk.END, f"{message}\n")
        self.status_text.see(tk.END)
        self.status_text.config(state="disabled")
        self.root.update_idletasks()

    def ask_password(self, title: str = None, prompt: str = None) -> str | None:
        """Show modal password dialog and return entered password or None."""
        title = title or get_translation("password", self.language_var.get())
        prompt = prompt or get_translation("enter_password", self.language_var.get())

        result = {"pwd": None}

        dlg = tk.Toplevel(self.root)
        dlg.title(title)
        dlg.transient(self.root)
        dlg.resizable(False, False)
        dlg.grab_set()

        frm = ttk.Frame(dlg, padding=12)
        frm.grid(row=0, column=0, sticky="nsew")

        ttk.Label(frm, text=prompt).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 8)
        )

        pwd_var = tk.StringVar()
        entry = ttk.Entry(frm, textvariable=pwd_var, show="*")
        entry.grid(row=1, column=0, columnspan=2, sticky="we", pady=(0, 8))
        entry.focus_set()

        def on_ok():
            result["pwd"] = pwd_var.get()
            dlg.destroy()

        def on_cancel():
            dlg.destroy()

        ok_btn = ttk.Button(
            frm, text=get_translation("ok", self.language_var.get()), command=on_ok
        )
        ok_btn.grid(row=2, column=0, sticky="e", padx=(0, 6))
        cancel_btn = ttk.Button(
            frm,
            text=get_translation("cancel", self.language_var.get()),
            command=on_cancel,
        )
        cancel_btn.grid(row=2, column=1, sticky="w")

        dlg.wait_window()
        return result["pwd"]

    def process_pdf(self):
        """Process the PDF file"""
        # Validate inputs
        if not self.pdf_path.get():
            messagebox.showerror(
                get_translation("error", self.language_var.get()),
                get_translation("please_select_pdf", self.language_var.get()),
            )
            return

        if not self.selected_parser.get():
            messagebox.showerror(
                get_translation("error", self.language_var.get()),
                get_translation("please_select_parser", self.language_var.get()),
            )
            return

        # Validate according to user choice
        if self.use_existing_excel.get():
            if not self.excel_path.get():
                messagebox.showerror(
                    get_translation("error", self.language_var.get()),
                    get_translation("please_select_excel", self.language_var.get()),
                )
                return
        else:
            if not self.output_path.get():
                messagebox.showerror(
                    get_translation("error", self.language_var.get()),
                    get_translation("please_specify_output", self.language_var.get()),
                )
                return

        # Disable process button
        self.process_btn.config(state="disabled")
        self.log_message(
            get_translation("starting_processing", self.language_var.get())
        )

        # Run processing in separate thread
        # thread = threading.Thread(target=self._process_pdf_thread)
        self._process_pdf_thread()
        # thread.daemon = True
        # thread.start()

    def _process_pdf_thread(self):
        """Process PDF in separate thread"""
        delete_temp_file = False
        try:
            pdf_path = self.pdf_path.get()
            try:
                result = pdf_to_text(pdf_path)
            except:
                print("Could not parse; using decryptor")
                fd, tmp_fpath = tempfile.mkstemp(suffix=".pdf")
                os.close(fd)
                pwd = self.ask_password(
                    prompt=get_translation("pdf_password", self.language_var.get())
                )
                if pwd:
                    decrypt_pdf(pdf_path, tmp_fpath, pwd)
                    pdf_path = tmp_fpath
                    delete_temp_file = True
                else:
                    self.log_message(
                        get_translation("password_cancelled", self.language_var.get())
                    )
            # Get parser instance
            parser_instance = registry.create_parser(self.selected_parser.get())

            rules_path = (
                Path(__file__).parent
                / "data"
                / "rules"
                / f"{self.language_var.get()}.csv"
            )
            rules = load_rules(rules_path)

            # Process PDF
            existing_excel = self.excel_path.get() if self.excel_path.get() else None
            success, message = process_pdf_to_excel(
                pdf_path,
                parser_instance,
                rules,
                self.output_path.get(),
                existing_excel,
                sheet_name=self.sheet_name_var.get(),
                language=self.language_var.get(),
            )

            if success:
                self.log_message(f"SUCCESS: {message}")
                messagebox.showinfo(
                    get_translation("success", self.language_var.get()),
                    f"{get_translation('pdf_processed_successfully', self.language_var.get())}\n\n{message}",
                )
            else:
                self.log_message(f"ERROR: {message}")
                messagebox.showerror(
                    get_translation("error", self.language_var.get()),
                    f"{get_translation('failed_to_process_pdf', self.language_var.get())}\n\n{message}",
                )
            if delete_temp_file:
                os.remove(tmp_fpath)
        except Exception as e:
            error_msg = get_translation(
                "unexpected_error", self.language_var.get()
            ).format(str(e))
            self.log_message(f"ERROR: {error_msg}")
            messagebox.showerror(
                get_translation("error", self.language_var.get()), error_msg
            )
        finally:
            # Re-enable process button
            self.root.after(0, lambda: self.process_btn.config(state="normal"))

    def run(self):
        """Start the GUI application"""
        self.root.mainloop()


if __name__ == "__main__":
    app = ParserGUI()
    app.run()
