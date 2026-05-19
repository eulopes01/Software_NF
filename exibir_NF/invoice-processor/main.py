"""
main.py — Invoice Processor — Graphical Interface
---------------------------------------------------
Entry point of the application. Provides a Tkinter-based GUI that:
    1. Allows the user to select multiple PDF invoices at once
    2. Processes each PDF through the full pipeline:
       pdf_reader → spreadsheet_writer → reconciler
    3. Displays a table with all extracted transactions
    4. Shows the reconciliation result (✅ MATCH or ❌ DIVERGENCE)
       for each person processed

Run:
    python main.py

Dependencies:
    pip install pdfplumber openpyxl
    (Tkinter is included in the Python standard library)
"""

import logging
import sys
import re
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent / "src"))

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Optional

# ─────────────────────────────────────────────
# Logging setup
# ─────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
    handlers=[
        logging.FileHandler("logs/processing.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
logger = logging.getLogger(__name__)

# ─────────────────────────────────────────────
# Internal imports
# ─────────────────────────────────────────────

try:
    from pdf_reader import read_pdf, Transaction, extract_itau_by_cardholder, extract_itau_empresas_by_cardholder, CardHolder, BANK_ITAU, BANK_ITAU_EMPRESAS
    from spreadsheet_writer import write_transactions
    from reconciler import reconcile, read_pdf_total
    from utils import load_config, format_currency, sanitize_text
    from pdf_unlocker import is_pdf_protected, unlock_pdf_context
    from spreadsheet_creator_dialog import open_creator_dialog
except ImportError as exc:
    # Show a user-friendly error if a module is missing
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror(
        "Import Error",
        f"Failed to load a required module:\n\n{exc}\n\n"
        "Make sure all source files are in the 'src/' directory\n"
        "and the virtual environment is activated."
    )
    sys.exit(1)

# ─────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────

APP_TITLE           = "Invoice Processor"
APP_VERSION         = "1.0.0"
WINDOW_MIN_WIDTH    = 960
WINDOW_MIN_HEIGHT   = 640
WINDOW_START_WIDTH  = 1100
WINDOW_START_HEIGHT = 720

COLOR_BG            = "#F5F5F5"
COLOR_HEADER        = "#1F4E79"
COLOR_HEADER_FG     = "#FFFFFF"
COLOR_MATCH         = "#1E7E34"
COLOR_DIVERGENCE    = "#C82333"
COLOR_ACCENT        = "#2E75B6"
COLOR_ROW_ODD       = "#FFFFFF"
COLOR_ROW_EVEN      = "#EBF3FB"
COLOR_BTN_PRIMARY   = "#2E75B6"
COLOR_BTN_DANGER    = "#C82333"
COLOR_BTN_FG        = "#FFFFFF"

FONT_TITLE          = ("Arial", 16, "bold")
FONT_SUBTITLE       = ("Arial", 11)
FONT_LABEL          = ("Arial", 10)
FONT_LABEL_BOLD     = ("Arial", 10, "bold")
FONT_TABLE          = ("Arial", 9)
FONT_TABLE_HEADER   = ("Arial", 9, "bold")
FONT_STATUS         = ("Arial", 10, "bold")

TABLE_COLUMNS = [
    ("Person",      160),
    ("Date",         70),
    ("Description", 280),
    ("Amount",      100),
    ("Bank",         80),
    ("Status",      100),
]

SPREADSHEET_DIR     = Path("spreadsheets")
PDF_DIR             = Path("pdfs")


# ─────────────────────────────────────────────
# Function 1 — build_main_window
# ─────────────────────────────────────────────

def build_main_window(root: tk.Tk) -> dict:
    """
    Builds and configures the main application window and all
    its widgets: header, PDF selection panel, action buttons,
    transactions table, summary panel, and status bar.

    Returns a dictionary of widget references so that other
    functions can update them without using global variables.
    The dictionary keys are descriptive widget names.

    Args:
        root: The root Tk window to build widgets into.

    Returns:
        A dict of widget references keyed by name.
    """
    root.title(f"{APP_TITLE} v{APP_VERSION}")
    root.minsize(WINDOW_MIN_WIDTH, WINDOW_MIN_HEIGHT)
    root.geometry(f"{WINDOW_START_WIDTH}x{WINDOW_START_HEIGHT}")
    root.configure(bg=COLOR_BG)
    root.resizable(True, True)

    widgets = {}

    # ── Header ────────────────────────────────
    header_frame = tk.Frame(root, bg=COLOR_HEADER, pady=12)
    header_frame.pack(fill="x")

    tk.Label(
        header_frame, text=f"  {APP_TITLE}",
        font=FONT_TITLE, bg=COLOR_HEADER, fg=COLOR_HEADER_FG,
    ).pack(side="left")

    tk.Label(
        header_frame, text=f"v{APP_VERSION}  ",
        font=FONT_SUBTITLE, bg=COLOR_HEADER, fg="#B0C4DE",
    ).pack(side="right", anchor="s", pady=4)

    # ── PDF selection panel ───────────────────
    selection_frame = tk.LabelFrame(
        root, text=" PDF Invoices ", font=FONT_LABEL_BOLD,
        bg=COLOR_BG, padx=10, pady=8,
    )
    selection_frame.pack(fill="x", padx=16, pady=(12, 4))

    pdf_listbox_frame = tk.Frame(selection_frame, bg=COLOR_BG)
    pdf_listbox_frame.pack(fill="x")

    pdf_scrollbar = tk.Scrollbar(pdf_listbox_frame, orient="vertical")
    pdf_listbox = tk.Listbox(
        pdf_listbox_frame, height=4, font=FONT_LABEL,
        yscrollcommand=pdf_scrollbar.set,
        selectmode="extended", bg="#FFFFFF",
    )
    pdf_scrollbar.config(command=pdf_listbox.yview)
    pdf_scrollbar.pack(side="right", fill="y")
    pdf_listbox.pack(fill="x", expand=True)
    widgets["pdf_listbox"] = pdf_listbox

    # ── Buttons row ───────────────────────────
    btn_frame = tk.Frame(root, bg=COLOR_BG, pady=6)
    btn_frame.pack(fill="x", padx=16)

    btn_select = _make_button(
        btn_frame, "📂  Select PDFs", COLOR_BTN_PRIMARY,
        lambda: select_pdf_files(widgets),
    )
    btn_select.pack(side="left", padx=(0, 8))
    widgets["btn_select"] = btn_select

    btn_new_sheet = _make_button(
        btn_frame, "📊  New Spreadsheet", "#17A589",
        lambda: open_creator_dialog(root, output_dir=SPREADSHEET_DIR),
    )
    btn_new_sheet.pack(side="left", padx=(0, 8))
    widgets["btn_new_sheet"] = btn_new_sheet

    btn_clear = _make_button(
        btn_frame, "🗑  Clear list", "#6C757D",
        lambda: clear_pdf_list(widgets),
    )
    btn_clear.pack(side="left", padx=(0, 8))
    widgets["btn_clear"] = btn_clear

    btn_process = _make_button(
        btn_frame, "▶  Process invoices", COLOR_ACCENT,
        lambda: start_processing_thread(widgets),
    )
    btn_process.pack(side="left")
    widgets["btn_process"] = btn_process

    btn_clear_table = _make_button(
        btn_frame, "🔄  Clear results", "#6C757D",
        lambda: clear_results(widgets),
    )
    btn_clear_table.pack(side="right")
    widgets["btn_clear_table"] = btn_clear_table

    # ── Password frame ────────────────────────
    password_frame = tk.Frame(root, bg=COLOR_BG, pady=2)
    password_frame.pack(fill="x", padx=16)

    tk.Label(
        password_frame, text="🔒  PDF Password (if protected):",
        font=FONT_LABEL_BOLD, bg=COLOR_BG,
    ).pack(side="left", padx=(0, 8))

    password_var = tk.StringVar()
    password_entry = tk.Entry(
        password_frame, textvariable=password_var,
        font=FONT_LABEL, width=24, show="●",
    )
    password_entry.pack(side="left")
    widgets["password_var"] = password_var

    tk.Label(
        password_frame,
        text="Leave blank if PDF has no password.",
        font=("Arial", 9), bg=COLOR_BG, fg="#6C757D",
    ).pack(side="left", padx=(8, 0))

    # ── Transactions table ────────────────────
    table_frame = tk.LabelFrame(
        root, text=" Extracted Transactions ", font=FONT_LABEL_BOLD,
        bg=COLOR_BG, padx=6, pady=6,
    )
    table_frame.pack(fill="both", expand=True, padx=16, pady=(4, 4))

    table_scroll_y = ttk.Scrollbar(table_frame, orient="vertical")
    table_scroll_x = ttk.Scrollbar(table_frame, orient="horizontal")

    tree = ttk.Treeview(
        table_frame,
        columns=[col for col, _ in TABLE_COLUMNS],
        show="headings",
        yscrollcommand=table_scroll_y.set,
        xscrollcommand=table_scroll_x.set,
        selectmode="browse",
    )
    table_scroll_y.config(command=tree.yview)
    table_scroll_x.config(command=tree.xview)
    table_scroll_y.pack(side="right", fill="y")
    table_scroll_x.pack(side="bottom", fill="x")
    tree.pack(fill="both", expand=True)

    for col_name, col_width in TABLE_COLUMNS:
        tree.heading(col_name, text=col_name, anchor="w")
        tree.column(col_name, width=col_width, minwidth=60, anchor="w")

    tree.tag_configure("match",      background="#D4EDDA", foreground=COLOR_MATCH)
    tree.tag_configure("divergence", background="#F8D7DA", foreground=COLOR_DIVERGENCE)
    tree.tag_configure("odd",        background=COLOR_ROW_ODD)
    tree.tag_configure("even",       background=COLOR_ROW_EVEN)
    widgets["tree"] = tree

    # ── Summary panel ─────────────────────────
    summary_frame = tk.Frame(root, bg=COLOR_BG, pady=4)
    summary_frame.pack(fill="x", padx=16)

    widgets["summary_labels"] = {}
    widgets["summary_frame"] = summary_frame

    # ── Status bar ────────────────────────────
    status_bar = tk.Label(
        root, text="Ready — select PDF invoices to begin.",
        font=FONT_LABEL, bg=COLOR_HEADER, fg=COLOR_HEADER_FG,
        anchor="w", padx=10, pady=4,
    )
    status_bar.pack(fill="x", side="bottom")
    widgets["status_bar"] = status_bar

    # ── Progress bar ──────────────────────────
    progress = ttk.Progressbar(root, mode="indeterminate")
    widgets["progress"] = progress

    return widgets


# ─────────────────────────────────────────────
# Function 2 — select_pdf_files
# ─────────────────────────────────────────────

def select_pdf_files(widgets: dict) -> None:
    """
    Opens a native file dialog allowing the user to select
    multiple PDF invoice files simultaneously.

    Selected files are validated (must end in .pdf) and added
    to the PDF listbox. Duplicate paths are ignored — the same
    file cannot be added twice to prevent double-processing.

    The function does nothing if the user cancels the dialog.

    Security: Only files with the .pdf extension are accepted.
    The file dialog is restricted to PDF files via the filetypes
    filter, providing both UX convenience and a first layer of
    input validation.

    Args:
        widgets: Dictionary of widget references from build_main_window().

    Returns:
        None. Updates the pdf_listbox widget directly.
    """
    file_paths = filedialog.askopenfilenames(
        title="Select PDF Invoice Files",
        filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        initialdir=str(PDF_DIR) if PDF_DIR.exists() else str(Path.home()),
    )

    if not file_paths:
        logger.debug("File dialog cancelled by user.")
        return

    listbox = widgets["pdf_listbox"]
    existing_paths = set(listbox.get(0, "end"))
    added_count = 0
    rejected_count = 0

    for raw_path in file_paths:
        path = Path(raw_path)

        # Validate extension — reject non-PDF files
        if path.suffix.lower() != ".pdf":
            logger.warning(f"Rejected non-PDF file: '{path.name}'")
            rejected_count += 1
            continue

        # Reject duplicate entries
        if str(path) in existing_paths:
            logger.debug(f"Duplicate skipped: '{path.name}'")
            continue

        listbox.insert("end", str(path))
        existing_paths.add(str(path))
        added_count += 1

    total_in_list = listbox.size()
    status_msg = f"{added_count} PDF(s) added — {total_in_list} total in list."

    if rejected_count > 0:
        status_msg += f" {rejected_count} non-PDF file(s) rejected."

    _set_status(widgets, status_msg)
    logger.info(f"PDF selection: added={added_count}, rejected={rejected_count}, total={total_in_list}")


# ─────────────────────────────────────────────
# Function 3 — start_processing_thread
# ─────────────────────────────────────────────

def start_processing_thread(widgets: dict) -> None:
    """
    Validates that PDFs are selected and starts the processing
    pipeline in a background thread to keep the UI responsive.

    Running the pipeline in a separate thread prevents the
    Tkinter main loop from freezing while PDF extraction,
    spreadsheet writing, and reconciliation are in progress.

    The processing button is disabled during execution and
    re-enabled when the thread finishes (success or error).
    A progress bar is shown during processing.

    Args:
        widgets: Dictionary of widget references from build_main_window().

    Returns:
        None. Spawns a daemon thread that calls process_all_invoices().
    """
    listbox = widgets["pdf_listbox"]
    pdf_paths = list(listbox.get(0, "end"))

    if not pdf_paths:
        messagebox.showwarning(
            "No PDFs Selected",
            "Please select at least one PDF invoice before processing.",
        )
        return

    # Validate that the spreadsheets directory exists
    if not SPREADSHEET_DIR.exists():
        answer = messagebox.askyesno(
            "Directory Not Found",
            f"The spreadsheets directory was not found:\n'{SPREADSHEET_DIR}'\n\n"
            "Would you like to create it now?",
        )
        if answer:
            SPREADSHEET_DIR.mkdir(parents=True, exist_ok=True)
            logger.info(f"Created spreadsheets directory: {SPREADSHEET_DIR}")
        else:
            return

    # Disable UI controls during processing
    widgets["btn_process"].config(state="disabled")
    widgets["btn_select"].config(state="disabled")
    widgets["btn_clear"].config(state="disabled")
    widgets["progress"].pack(fill="x", padx=16, pady=2)
    widgets["progress"].start(10)
    _set_status(widgets, f"Processing {len(pdf_paths)} invoice(s)... please wait.")

    def run():
        try:
            process_all_invoices(widgets, pdf_paths)
        except Exception as exc:
            logger.error(f"Unexpected error during processing: {exc}", exc_info=True)
            widgets["tree"].after(0, lambda: messagebox.showerror(
                "Processing Error",
                f"An unexpected error occurred:\n\n{exc}\n\n"
                "Check the logs/processing.log file for details."
            ))
        finally:
            widgets["tree"].after(0, lambda: _restore_ui_after_processing(widgets))

    thread = threading.Thread(target=run, daemon=True)
    thread.start()


# ─────────────────────────────────────────────
# Function 4 — process_all_invoices
# ─────────────────────────────────────────────

def process_all_invoices(widgets: dict, pdf_paths: list[str]) -> None:
    """
    Runs the full processing pipeline for each selected PDF:
        1. read_pdf()          — extract transactions and total
        2. resolve spreadsheet — find or prompt for the .xlsx file
        3. write_transactions() — insert rows into the spreadsheet
        4. reconcile()         — compare PDF total vs spreadsheet total
        5. populate_table()    — display results in the UI table

    Each PDF is processed independently. Errors in one PDF do not
    stop the processing of the remaining files.

    The function runs in a background thread (started by
    start_processing_thread) and must never call Tkinter widget
    methods directly — all UI updates are scheduled via .after(0, ...).

    Args:
        widgets:   Dictionary of widget references.
        pdf_paths: List of absolute PDF file path strings.

    Returns:
        None. Updates the UI table via populate_table().
    """
    config = load_config()
    all_rows: list[dict] = []
    summary_results: list = []

    for pdf_path in pdf_paths:
        path = Path(pdf_path)
        logger.info(f"Processing: {path.name}")

        try:
            # Step 1 — Unlock if protected, then read PDF
            password = widgets.get("password_var", tk.StringVar()).get().strip()
            protected = is_pdf_protected(pdf_path)

            if protected and not password:
                logger.warning(f"'{path.name}' is password-protected but no password was provided.")
                _set_status(widgets, f"🔒 '{path.name}' is protected — enter the password and try again.")
                continue

            # Step 2 — Extract raw text
            if protected:
                with unlock_pdf_context(pdf_path, password) as unlocked_path:
                    transactions, raw_total = read_pdf(str(unlocked_path))
                    raw_text = _extract_raw_text_from_pdf(str(unlocked_path))
            else:
                transactions, raw_total = read_pdf(pdf_path)
                raw_text = _extract_raw_text_from_pdf(pdf_path)

            validated_total = read_pdf_total(raw_total, path.name)
            sheet_name = _resolve_sheet_name(config)

            # Step 3 — Extract by cardholder if Itaú, else single person
            from pdf_reader import detect_bank, BANK_ITAU
            bank = detect_bank(raw_text)

            if bank == BANK_ITAU:
                cardholders, _ = extract_itau_by_cardholder(raw_text, path.name)
            elif bank == BANK_ITAU_EMPRESAS:
                cardholders, _ = extract_itau_empresas_by_cardholder(raw_text, path.name)
            else:
                # For BB or unknown: treat as single cardholder
                person_name = _resolve_person_name(path)
                from pdf_reader import CardHolder
                cardholders = [CardHolder(
                    name=person_name,
                    card_last_digits="0000",
                    transactions=transactions,
                    bank=bank,
                )]

            # Step 4 — Process each cardholder separately
            for cardholder in cardholders:
                person_name = _format_cardholder_name(cardholder.name)
                card_digits = cardholder.card_last_digits

                spreadsheet_path = _resolve_spreadsheet_path_by_card(
                    person_name, card_digits, sheet_name, widgets
                )

                if not spreadsheet_path:
                    logger.warning(f"No spreadsheet for '{person_name}' (...{card_digits}). Skipping.")
                    continue

                write_transactions(
                    spreadsheet_path=str(spreadsheet_path),
                    sheet_name=sheet_name,
                    transactions=cardholder.transactions,
                )

                result = reconcile(
                    person_name=f"{person_name} (...{card_digits})",
                    sheet_name=sheet_name,
                    spreadsheet_path=str(spreadsheet_path),
                    pdf_total=validated_total,
                )

                status_label = "✅ MATCH" if result.status == "MATCH" else "❌ DIVERGENCE"
                for transaction in cardholder.transactions:
                    all_rows.append({
                        "person":      f"{person_name} (...{card_digits})",
                        "date":        transaction.date,
                        "description": transaction.description,
                        "amount":      format_currency(transaction.amount),
                        "bank":        transaction.bank.upper(),
                        "status":      status_label,
                        "match":       result.status == "MATCH",
                    })

                summary_results.append(result)

        except FileNotFoundError as exc:
            logger.error(f"File not found during processing of '{path.name}': {exc}")
            _set_status(widgets, f"❌ File not found: {path.name}")
        except ValueError as exc:
            logger.error(f"Validation error processing '{path.name}': {exc}")
            _set_status(widgets, f"❌ Error in '{path.name}': {exc}")
        except Exception as exc:
            logger.error(f"Unexpected error processing '{path.name}': {exc}", exc_info=True)
            _set_status(widgets, f"❌ Unexpected error in '{path.name}'.")

    # Schedule UI update on the main thread
    widgets["tree"].after(0, lambda: populate_table(widgets, all_rows, summary_results))


# ─────────────────────────────────────────────
# Function 5 — populate_table
# ─────────────────────────────────────────────

def populate_table(
    widgets: dict,
    rows: list[dict],
    summary_results: list,
) -> None:
    """
    Clears the transactions table and populates it with the
    extracted transaction rows and reconciliation results.

    Each row is color-coded:
        - Green background → ✅ MATCH
        - Red background   → ❌ DIVERGENCE
        - Alternating white/light-blue for readability

    After populating the table, a summary panel is rendered
    below the table showing one result badge per person:
        ✅ Aline Maris — MATCH (R$ 1.337,61)
        ❌ Carlos Belruss — DIVERGENCE (diff: R$ 50,00)

    This function must only be called from the main Tkinter
    thread (scheduled via .after() from the worker thread).

    Args:
        widgets:         Dictionary of widget references.
        rows:            List of row dicts from process_all_invoices().
        summary_results: List of ReconciliationResult objects.

    Returns:
        None. Updates tree and summary widgets directly.
    """
    tree = widgets["tree"]

    # Clear existing rows
    for item in tree.get_children():
        tree.delete(item)

    # Clear summary panel
    for label in widgets.get("summary_labels", {}).values():
        label.destroy()
    widgets["summary_labels"] = {}

    if not rows:
        _set_status(widgets, "No transactions extracted. Check logs for details.")
        return

    # Insert rows with alternating colors and status tags
    for index, row in enumerate(rows):
        is_match = row.get("match", True)
        row_tag = "match" if is_match else "divergence"

        tree.insert(
            "", "end",
            values=(
                sanitize_text(row["person"],      max_length=50),
                sanitize_text(row["date"],        max_length=10),
                sanitize_text(row["description"], max_length=100),
                sanitize_text(row["amount"],      max_length=20),
                sanitize_text(row["bank"],        max_length=20),
                row["status"],
            ),
            tags=(row_tag,),
        )

    # Render summary badges per person
    summary_frame = widgets["summary_frame"]
    match_count = sum(1 for r in summary_results if r.status == "MATCH")
    divergence_count = len(summary_results) - match_count

    for result in summary_results:
        is_match = result.status == "MATCH"
        icon = "✅" if is_match else "❌"
        color = COLOR_MATCH if is_match else COLOR_DIVERGENCE
        diff_str = format_currency(result.difference)
        label_text = (
            f"{icon}  {result.person_name}  |  "
            f"PDF: {format_currency(result.pdf_total)}  |  "
            f"Sheet: {format_currency(result.spreadsheet_total)}  |  "
            f"Diff: {diff_str}"
        )
        lbl = tk.Label(
            summary_frame, text=label_text,
            font=FONT_STATUS, bg=COLOR_BG, fg=color, anchor="w",
        )
        lbl.pack(fill="x", pady=1)
        widgets["summary_labels"][result.person_name] = lbl

    total_transactions = len(rows)
    status_msg = (
        f"Done — {total_transactions} transaction(s) processed | "
        f"{match_count} ✅ match | {divergence_count} ❌ divergence"
    )
    _set_status(widgets, status_msg)
    logger.info(status_msg)


# ─────────────────────────────────────────────
# Private helpers
# ─────────────────────────────────────────────

def clear_pdf_list(widgets: dict) -> None:
    """Clears all entries from the PDF listbox."""
    widgets["pdf_listbox"].delete(0, "end")
    _set_status(widgets, "PDF list cleared.")


def clear_results(widgets: dict) -> None:
    """Clears the transactions table and summary panel."""
    tree = widgets["tree"]
    for item in tree.get_children():
        tree.delete(item)
    for label in widgets.get("summary_labels", {}).values():
        label.destroy()
    widgets["summary_labels"] = {}
    _set_status(widgets, "Results cleared.")


def _set_status(widgets: dict, message: str) -> None:
    """Updates the status bar text safely from any thread."""
    def update():
        widgets["status_bar"].config(text=f"  {message}")
    try:
        widgets["status_bar"].after(0, update)
    except Exception:
        pass


def _restore_ui_after_processing(widgets: dict) -> None:
    """Re-enables UI controls after the processing thread finishes."""
    widgets["btn_process"].config(state="normal")
    widgets["btn_select"].config(state="normal")
    widgets["btn_clear"].config(state="normal")
    widgets["progress"].stop()
    widgets["progress"].pack_forget()


def _make_button(parent, text: str, color: str, command) -> tk.Button:
    """Creates a styled button with consistent appearance."""
    return tk.Button(
        parent, text=text, command=command,
        bg=color, fg=COLOR_BTN_FG,
        font=FONT_LABEL_BOLD, relief="flat",
        padx=14, pady=6, cursor="hand2",
        activebackground=color, activeforeground=COLOR_BTN_FG,
    )


def _resolve_person_name(pdf_path: Path) -> str:
    """
    Infers the person's name from the PDF file name.

    Convention: the PDF file should be named after the person,
    e.g. 'Aline_Maris_maio2025.pdf' → 'Aline Maris'.
    Underscores are replaced with spaces and the month/year
    suffix (if present) is stripped from the END of the name only.

    Strategy:
        1. Remove trailing 4-digit year (e.g. _2025)
        2. Remove trailing month name + optional year (e.g. _maio2025)
        3. Remove trailing MM_YYYY pattern (e.g. _05_2025)
        4. Replace underscores/hyphens with spaces and strip

    Falls back to the raw file stem if parsing fails.
    """
    stem = pdf_path.stem

    # Step 1 — Remove trailing _MM_YYYY or _YYYY patterns
    clean = re.sub(r"[_\-]\d{2}[_\-]\d{4}$", "", stem)
    clean = re.sub(r"[_\-]\d{4}$", "", clean)

    # Step 2 — Full month names first (avoids "mai" matching inside "Maris")
    clean = re.sub(
        r"[_\-](janeiro|fevereiro|marco|abril|maio|junho|julho|agosto|"
        r"setembro|outubro|novembro|dezembro)\d*$",
        "", clean, flags=re.IGNORECASE,
    )
    # Short abbreviations — "mai" excluded because full "maio" covers it
    clean = re.sub(
        r"[_\-](jan|fev|mar|abr|jun|jul|ago|set|out|nov|dez)\d*$",
        "", clean, flags=re.IGNORECASE,
    )

    # Step 3 — Replace separators with spaces
    clean = clean.replace("_", " ").replace("-", " ").strip()

    # Step 4 — Collapse multiple spaces
    clean = re.sub(r"\s+", " ", clean).strip()

    result = clean if clean else stem
    logger.debug(f"_resolve_person_name: '{pdf_path.name}' → '{result}'")
    return result


def _resolve_spreadsheet_path(person_name: str, widgets: dict) -> Optional[Path]:
    """
    Searches the spreadsheets directory for a .xlsx file whose
    name contains the person's name (case-insensitive).

    If no match is found, prompts the user to manually select
    the correct spreadsheet file via a file dialog.

    Returns the Path if found or selected, or None if the user cancels.
    """
    name_normalized = person_name.lower().replace(" ", "_")

    # Search for matching spreadsheet
    for xlsx_file in SPREADSHEET_DIR.glob("*.xlsx"):
        if name_normalized in xlsx_file.stem.lower():
            logger.info(f"Spreadsheet matched: '{xlsx_file.name}' for '{person_name}'")
            return xlsx_file

    # Prompt user to select manually
    logger.warning(f"No spreadsheet auto-matched for '{person_name}'.")
    selected = filedialog.askopenfilename(
        title=f"Select spreadsheet for: {person_name}",
        filetypes=[("Excel files", "*.xlsx")],
        initialdir=str(SPREADSHEET_DIR) if SPREADSHEET_DIR.exists() else str(Path.home()),
    )

    return Path(selected) if selected else None


def _resolve_sheet_name(config: dict) -> str:
    """
    Resolves the target sheet (tab) name from config.json.

    Priority:
        1. Uses 'reference_month' key from config.json if present
           Example: "MAIO 2025"
        2. Falls back to the current month/year if not set

    To change the active month, edit config.json:
        "reference_month": "JUNHO 2025"

    Returns a string like 'MAIO 2025'.
    """
    from datetime import datetime

    reference_month = config.get("reference_month", "").strip()

    if reference_month:
        normalized = reference_month.upper()
        logger.debug(f"Sheet name from config: '{normalized}'")
        return normalized

    # Fallback to current month if not set in config
    now = datetime.now()
    months_pt = [
        "JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL",
        "MAIO", "JUNHO", "JULHO", "AGOSTO",
        "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO",
    ]
    month_name = months_pt[now.month - 1]
    fallback = f"{month_name} {now.year}"
    logger.warning(
        f"'reference_month' not set in config.json. "
        f"Using current month as fallback: '{fallback}'. "
        "Add 'reference_month' to config.json to avoid this warning."
    )
    return fallback




def _extract_raw_text_from_pdf(pdf_path: str) -> str:
    """
    Extracts raw text from a PDF file using pdfplumber.
    Returns concatenated text from all pages.
    Falls back to empty string if extraction fails.
    """
    try:
        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            return "\n".join(
                page.extract_text() for page in pdf.pages
                if page.extract_text()
            )
    except Exception as exc:
        logger.warning(f"Could not extract raw text from '{pdf_path}': {exc}")
        return ""


def _format_cardholder_name(raw_name: str) -> str:
    """
    Formats a cardholder name from the PDF (ALL CAPS) to Title Case
    for use in spreadsheet file names and display.

    Example: "CLAYTON PIRES DOS SANTOS" → "Clayton Pires Dos Santos"
    """
    if not raw_name:
        return "Unknown"
    return raw_name.strip().title()


def _resolve_spreadsheet_path_by_card(
    person_name: str,
    card_digits: str,
    sheet_name: str,
    widgets: dict,
) -> Optional[Path]:
    """
    Finds or creates a spreadsheet for a specific cardholder + card number.

    Strategy:
        1. Build the expected filename deterministically
        2. If file exists → use it (never create duplicate)
        3. If file does not exist → create it with the exact expected name

    Expected filename:
        Card_{Name}_{CardDigits}_{Month}_{Year}.xlsx
        Example: Card_Clayton_Pires_Dos_Santos_3172_MAIO_2026.xlsx
    """
    # Build the expected filename deterministically
    safe_name  = person_name.replace(" ", "_")
    safe_month = sheet_name.replace(" ", "_").upper()
    expected_name = f"Card_{safe_name}_{card_digits}_{safe_month}.xlsx"
    expected_path = SPREADSHEET_DIR / expected_name

    # If file already exists — use it directly, never create duplicate
    if expected_path.exists():
        logger.info(f"Spreadsheet found: '{expected_name}'")
        return expected_path

    # File does not exist — create it with the exact expected name
    logger.info(
        f"Spreadsheet not found: '{expected_name}'. "
        "Creating automatically."
    )

    try:
        from generate_spreadsheet import generate_person_spreadsheet

        person_name_with_card = f"{person_name} {card_digits}"

        file_path = generate_person_spreadsheet(
            person_name=person_name_with_card,
            reference_month=sheet_name,
            due_date="",
            output_dir=SPREADSHEET_DIR,
        )

        logger.info(f"Auto-created: '{file_path.name}'")
        return file_path

    except Exception as exc:
        logger.error(
            f"Failed to create spreadsheet for "
            f"'{person_name}' (...{card_digits}): {exc}"
        )
        return None

# ─────────────────────────────────────────────
# Entry point
# ─────────────────────────────────────────────

if __name__ == "__main__":
    Path("logs").mkdir(parents=True, exist_ok=True)

    root = tk.Tk()
    widgets = build_main_window(root)
    widgets['root'] = root

    logger.info(f"{APP_TITLE} v{APP_VERSION} started.")

    root.mainloop()