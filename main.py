import os
from pathlib import Path
import threading
import tkinter as tk
from collections import defaultdict
from tkinter import filedialog, messagebox, ttk
from typing import List, Tuple

import openpyxl

from lib.excel_io import write_rules_sheet_openpyxl
from lib.excel_io import (
    write_summary_section_openpyxl,
    write_transactions_sheet_openpyxl,
)

from lib.translations import get_translation
from parsers import BaseParser, Transaction, registry


def load_rules(path: str) -> List[Tuple[str, str]]:
    """Load rules from CSV or TXT file"""
    loaded: List[Tuple[str, str]] = []
    with open(path, "r") as f:
        for line in f.readlines():
            elements = line.split(",")
            if len(elements) < 2:
                raise ValueError("Bad rules file")
            loaded.append((elements[0], elements[1]))
    return loaded


def process_pdf_to_excel(
    pdf_path: str,
    parser_instance: BaseParser,
    rules: List[Tuple[str, str]],
    output_path: str,
    existing_excel: str = None,
    sheet_name: str = "Tranzactii",
    use_rules: bool = True,
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

    if use_rules and rules:
        write_rules_sheet_openpyxl(workbook, rules, language)

    # summary: write to the transactions worksheet
    trans_ws = workbook[sheet_name]
    write_summary_section_openpyxl(
        trans_ws,
        [{"months": k, "sum": v} for k, v in rate_buckets.items()],
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

    def __init__(self):
        self.root = tk.Tk()
        self.root.title(get_translation("app_title", "en"))
        self.root.geometry("600x500")

        # Initialize variables
        self._init_variables()

        # Available parsers
        self.parsers = registry.get_parsers()

        self.setup_ui()

    def _init_variables(self):
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
            self._update_widget_text(widget, current_lang)

    def _update_widget_text(self, widget, language):
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
                self._update_widget_text(child, language)

    def setup_ui(self):
        """Setup the user interface"""
        # Apply theme and refined styles
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        # Color palette
        bg = "#f4f7fb"  # overall background
        card_bg = "#ffffff"
        accent = "#0a66c2"
        accent_dark = "#084f92"

        # Root bg
        try:
            self.root.configure(background=bg)
        except Exception:
            pass

        # Card frame style
        style.configure("Card.TFrame", background=card_bg, relief="flat")
        style.configure("TFrame", background=bg)

        # Header
        style.configure("Header.TLabel", font=("Segoe UI", 16, "bold"), background=card_bg, foreground=accent_dark)

        # Labels and entries
        style.configure("TLabel", font=("Segoe UI", 10), background=card_bg)
        style.configure("TEntry", fieldbackground="#fbfdff")

        # Buttons
        style.configure("TButton", font=("Segoe UI", 10))
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"), foreground="#ffffff", background=accent)
        style.map("Accent.TButton",
                  foreground=[('active', '#ffffff')],
                  background=[('active', accent_dark), ('!disabled', accent)])

        # Process button gets an accent style
        style.configure("Process.TButton", parent="Accent.TButton")

        # Status text styling (use direct tk config later)

        # Main frame (card)
        main_frame = ttk.Frame(self.root, padding=16, style="Card.TFrame")
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
            style="Header.TLabel",
        )
        header.trans_key = "app_title"
        header.grid(row=row, column=0, columnspan=3, sticky=(tk.W), pady=(0, 10))
        row += 1

        # Language Selection
        ttk.Label(main_frame, text="Language:").grid(
            row=row, column=0, sticky=tk.W, pady=4
        )
        language_combo = ttk.Combobox(
            main_frame,
            textvariable=self.language_var,
            values=["en", "ro"],
            state="readonly",
        )
        language_combo.grid(row=row, column=1, sticky=(tk.W, tk.E), pady=5)
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
            style="Process.TButton",
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
                # Auto-detect parser
                detected_parser = registry.auto_detect_parser(filename)
                if detected_parser:
                    self.selected_parser.set(detected_parser)
                    self.log_message(
                        f"{get_translation('auto_detected_parser', self.language_var.get())} {detected_parser}"
                    )
                else:
                    self.log_message(
                        get_translation("could_not_auto_detect", self.language_var.get())
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
            # No rules UI to populate; rules will be read from workbook (or rules.csv) at processing time

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
        thread = threading.Thread(target=self._process_pdf_thread)
        thread.daemon = True
        thread.start()

    def _process_pdf_thread(self):
        """Process PDF in separate thread"""
        try:
            # Get parser instance
            parser_instance = registry.create_parser(self.selected_parser.get())

            # Load rules from the workbook (preferred) or rules.csv fallback
            rules = []
            wb_path = self.excel_path.get() if self.excel_path.get() else None
            if wb_path and os.path.exists(wb_path):
                try:
                    import openpyxl

                    wb = openpyxl.load_workbook(wb_path, data_only=True)
                    if "Rules" in wb.sheetnames:
                        ws = wb["Rules"]
                        for row in ws.iter_rows(min_row=2, values_only=True):
                            if not row:
                                continue
                            p = row[0] or ""
                            c = row[1] or ""
                            if p and c:
                                rules.append((str(p), str(c)))
                except Exception:
                    # fallback to rules.csv below
                    pass

            if not rules:
                # fallback to rules.csv in current directory
                try:
                    rules_path = Path(__file__).parent / "rules.csv"
                    rules = load_rules(rules_path)
                except Exception:
                    rules = []

            use_rules_flag = bool(rules)
            self.log_message(
                get_translation("loaded_rules", self.language_var.get()).format(
                    len(rules), use_rules_flag, False
                )
            )

            # Process PDF
            existing_excel = self.excel_path.get() if self.excel_path.get() else None
            success, message = process_pdf_to_excel(
                self.pdf_path.get(),
                parser_instance,
                rules,
                self.output_path.get(),
                existing_excel,
                sheet_name=self.sheet_name_var.get(),
                # clear_existing option removed
                use_rules=use_rules_flag,
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
