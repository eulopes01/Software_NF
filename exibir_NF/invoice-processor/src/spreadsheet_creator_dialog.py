"""
spreadsheet_creator_dialog.py — New Spreadsheet Dialog
--------------------------------------------------------
Provides a Tkinter dialog window that allows the user to
create one or more blank .xlsx spreadsheets directly from
the GUI, without running generate_spreadsheet.py manually.

The dialog collects the same information as generate_spreadsheet.py:
    - Person names (one per line)
    - Reference month (e.g. MAIO 2025)
    - Due date (e.g. 15/06/2025)
    - Annuity amount (e.g. 18.75)

On confirmation, it calls generate_all_spreadsheets() and
shows a success or error message.

Each function has a single responsibility and at least 25 lines,
following Clean Architecture, SOLID, and Secure by Design principles.

Usage:
    from spreadsheet_creator_dialog import open_creator_dialog

    open_creator_dialog(parent_window, output_dir=Path("spreadsheets"))
"""

import logging
import sys
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk
from typing import Optional

# Add scripts/ to path so generate_spreadsheet can be imported
sys.path.insert(0, str(Path(__file__).parent.parent / "scripts"))

from generate_spreadsheet import generate_all_spreadsheets

# ─────────────────────────────────────────────
# Logging setup
# ─────────────────────────────────────────────

logger = logging.getLogger(__name__)

# ─────────────────────────────────────────────
# Style constants
# ─────────────────────────────────────────────

COLOR_BG           = "#F5F5F5"
COLOR_HEADER       = "#1F4E79"
COLOR_HEADER_FG    = "#FFFFFF"
COLOR_ACCENT       = "#2E75B6"
COLOR_BTN_FG       = "#FFFFFF"
COLOR_SUCCESS      = "#1E7E34"
COLOR_ERROR        = "#C82333"

FONT_TITLE         = ("Arial", 13, "bold")
FONT_LABEL         = ("Arial", 10)
FONT_LABEL_BOLD    = ("Arial", 10, "bold")
FONT_SMALL         = ("Arial", 9)
FONT_BTN           = ("Arial", 10, "bold")

DIALOG_WIDTH       = 520
DIALOG_HEIGHT      = 480

PORTUGUESE_MONTHS  = [
    "JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL",
    "MAIO", "JUNHO", "JULHO", "AGOSTO",
    "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO",
]

CURRENT_YEARS      = [str(y) for y in range(2024, 2030)]
OUTPUT_DIR         = Path("spreadsheets")


# ─────────────────────────────────────────────
# Function 1 — open_creator_dialog
# ─────────────────────────────────────────────

def open_creator_dialog(parent: tk.Tk, output_dir: Path = OUTPUT_DIR) -> None:
    """
    Opens the spreadsheet creator dialog as a modal window
    centered relative to the parent window.

    The dialog is modal — the parent window is blocked until
    the dialog is closed. This prevents the user from processing
    invoices while the creator dialog is open.

    Args:
        parent:     The parent Tk window (main application window).
        output_dir: Directory where generated .xlsx files will be saved.
                    Defaults to OUTPUT_DIR ("spreadsheets/").

    Returns:
        None. All actions are performed inside the dialog.
    """
    if not isinstance(parent, (tk.Tk, tk.Toplevel)):
        logger.error(
            f"open_creator_dialog received invalid parent: {type(parent)}. "
            "Expected tk.Tk or tk.Toplevel."
        )
        return

    dialog = tk.Toplevel(parent)
    dialog.title("Create New Spreadsheet")
    dialog.configure(bg=COLOR_BG)
    dialog.resizable(False, False)
    dialog.grab_set()  # Make dialog modal

    # Center dialog relative to parent
    parent.update_idletasks()
    px = parent.winfo_x()
    py = parent.winfo_y()
    pw = parent.winfo_width()
    ph = parent.winfo_height()
    x = px + (pw - DIALOG_WIDTH) // 2
    y = py + (ph - DIALOG_HEIGHT) // 2
    dialog.geometry(f"{DIALOG_WIDTH}x{DIALOG_HEIGHT}+{x}+{y}")

    widgets = _build_dialog_widgets(dialog, output_dir)

    dialog.bind("<Return>", lambda e: _on_confirm(widgets, dialog, output_dir))
    dialog.bind("<Escape>", lambda e: dialog.destroy())

    logger.debug("Spreadsheet creator dialog opened.")
    dialog.wait_window()


# ─────────────────────────────────────────────
# Function 2 — _build_dialog_widgets
# ─────────────────────────────────────────────

def _build_dialog_widgets(dialog: tk.Toplevel, output_dir: Path) -> dict:
    """
    Builds and places all widgets inside the creator dialog:

        Header bar      — "Create New Spreadsheet" title
        People field    — Text area for entering names (one per line)
        Month dropdown  — Combobox with Portuguese month names
        Year dropdown   — Combobox with years 2024–2029
        Due date field  — Entry for the invoice due date (DD/MM/YYYY)
        Annuity field   — Entry for the monthly annuity amount
        Output label    — Shows where files will be saved
        Buttons row     — Cancel and Create buttons

    Args:
        dialog:     The Toplevel dialog window to build widgets into.
        output_dir: The directory where files will be saved (shown in UI).

    Returns:
        A dictionary of widget references keyed by name.
    """
    from datetime import datetime
    widgets = {}
    now = datetime.now()

    # ── Header ────────────────────────────────
    header = tk.Frame(dialog, bg=COLOR_HEADER, pady=10)
    header.pack(fill="x")
    tk.Label(
        header, text="  📊  Create New Spreadsheet",
        font=FONT_TITLE, bg=COLOR_HEADER, fg=COLOR_HEADER_FG,
    ).pack(side="left")

    # ── Form frame ────────────────────────────
    form = tk.Frame(dialog, bg=COLOR_BG, padx=20, pady=10)
    form.pack(fill="both", expand=True)

    # People
    tk.Label(form, text="Person names (one per line):",
             font=FONT_LABEL_BOLD, bg=COLOR_BG, anchor="w").pack(fill="x")
    tk.Label(form, text="Example: Aline Maris",
             font=FONT_SMALL, bg=COLOR_BG, fg="#6C757D", anchor="w").pack(fill="x")

    people_text = tk.Text(form, height=5, font=FONT_LABEL, relief="solid", bd=1)
    people_text.pack(fill="x", pady=(2, 10))
    widgets["people_text"] = people_text

    # Month + Year side by side
    month_year_frame = tk.Frame(form, bg=COLOR_BG)
    month_year_frame.pack(fill="x", pady=(0, 10))

    # Month
    month_frame = tk.Frame(month_year_frame, bg=COLOR_BG)
    month_frame.pack(side="left", fill="x", expand=True, padx=(0, 10))
    tk.Label(month_frame, text="Month:", font=FONT_LABEL_BOLD, bg=COLOR_BG, anchor="w").pack(fill="x")
    month_var = tk.StringVar(value=PORTUGUESE_MONTHS[now.month - 1])
    month_combo = ttk.Combobox(
        month_frame, textvariable=month_var,
        values=PORTUGUESE_MONTHS, state="readonly", font=FONT_LABEL,
    )
    month_combo.pack(fill="x")
    widgets["month_var"] = month_var

    # Year
    year_frame = tk.Frame(month_year_frame, bg=COLOR_BG)
    year_frame.pack(side="left", fill="x", expand=True)
    tk.Label(year_frame, text="Year:", font=FONT_LABEL_BOLD, bg=COLOR_BG, anchor="w").pack(fill="x")
    year_var = tk.StringVar(value=str(now.year))
    year_combo = ttk.Combobox(
        year_frame, textvariable=year_var,
        values=CURRENT_YEARS, state="readonly", font=FONT_LABEL,
    )
    year_combo.pack(fill="x")
    widgets["year_var"] = year_var

    # Due date + Annuity side by side
    due_annuity_frame = tk.Frame(form, bg=COLOR_BG)
    due_annuity_frame.pack(fill="x", pady=(0, 10))

    # Due date
    due_frame = tk.Frame(due_annuity_frame, bg=COLOR_BG)
    due_frame.pack(side="left", fill="x", expand=True, padx=(0, 10))
    tk.Label(due_frame, text="Due date (DD/MM/YYYY):",
             font=FONT_LABEL_BOLD, bg=COLOR_BG, anchor="w").pack(fill="x")
    due_var = tk.StringVar(value="")
    due_entry = tk.Entry(due_frame, textvariable=due_var, font=FONT_LABEL, relief="solid", bd=1)
    due_entry.pack(fill="x")
    widgets["due_var"] = due_var

    # Annuity
    annuity_frame = tk.Frame(due_annuity_frame, bg=COLOR_BG)
    annuity_frame.pack(side="left", fill="x", expand=True)
    tk.Label(annuity_frame, text="Annuity (R$):",
             font=FONT_LABEL_BOLD, bg=COLOR_BG, anchor="w").pack(fill="x")
    annuity_var = tk.StringVar(value="18.75")
    annuity_entry = tk.Entry(annuity_frame, textvariable=annuity_var, font=FONT_LABEL, relief="solid", bd=1)
    annuity_entry.pack(fill="x")
    widgets["annuity_var"] = annuity_var

    # Output dir label
    tk.Label(
        form,
        text=f"📁  Files will be saved to: {output_dir.resolve()}/",
        font=FONT_SMALL, bg=COLOR_BG, fg="#6C757D", anchor="w",
    ).pack(fill="x", pady=(0, 10))

    # Status label
    status_label = tk.Label(form, text="", font=FONT_SMALL, bg=COLOR_BG, anchor="w")
    status_label.pack(fill="x")
    widgets["status_label"] = status_label

    # ── Buttons ───────────────────────────────
    btn_frame = tk.Frame(dialog, bg=COLOR_BG, pady=10, padx=20)
    btn_frame.pack(fill="x")

    tk.Button(
        btn_frame, text="Cancel", font=FONT_BTN,
        bg="#6C757D", fg=COLOR_BTN_FG, relief="flat",
        padx=16, pady=6, cursor="hand2",
        command=dialog.destroy,
    ).pack(side="right", padx=(8, 0))

    tk.Button(
        btn_frame, text="✅  Create Spreadsheet(s)", font=FONT_BTN,
        bg=COLOR_ACCENT, fg=COLOR_BTN_FG, relief="flat",
        padx=16, pady=6, cursor="hand2",
        command=lambda: _on_confirm(widgets, dialog, output_dir),
    ).pack(side="right")

    widgets["dialog"] = dialog
    return widgets


# ─────────────────────────────────────────────
# Function 3 — _on_confirm
# ─────────────────────────────────────────────

def _on_confirm(widgets: dict, dialog: tk.Toplevel, output_dir: Path) -> None:
    """
    Handles the Create button click event.

    Collects and validates all form inputs, then calls
    generate_all_spreadsheets() to create the .xlsx files.
    Shows a success message with the list of created files,
    or an error message if validation or generation fails.

    Validation rules:
        - At least one person name must be provided
        - Person names must contain only letters and spaces
        - Due date must match DD/MM/YYYY format
        - Annuity must be a positive number

    Args:
        widgets:    Dictionary of widget references from _build_dialog_widgets().
        dialog:     The Toplevel dialog window (closed on success).
        output_dir: Directory where files will be saved.

    Returns:
        None. Updates the status label and shows messageboxes.
    """
    # Collect inputs
    raw_people = widgets["people_text"].get("1.0", "end").strip()
    month      = widgets["month_var"].get().strip().upper()
    year       = widgets["year_var"].get().strip()
    due_date   = widgets["due_var"].get().strip()
    annuity    = widgets["annuity_var"].get().strip()

    # Validate inputs
    error = _validate_form_inputs(raw_people, month, year, due_date, annuity)
    if error:
        widgets["status_label"].config(text=f"⚠️  {error}", fg=COLOR_ERROR)
        logger.warning(f"Form validation failed: {error}")
        return

    # Parse people list
    people = _parse_people_input(raw_people)
    reference_month = f"{month} {year}"

    widgets["status_label"].config(
        text=f"Creating {len(people)} spreadsheet(s)...", fg=COLOR_ACCENT
    )
    dialog.update()

    try:
        # Temporarily override annuity in generate_spreadsheet module
        import generate_spreadsheet as gs
        original_annuity = gs.ANNUITY_AMOUNT
        gs.ANNUITY_AMOUNT = float(annuity)

        created = generate_all_spreadsheets(
            people=people,
            reference_month=reference_month,
            due_date=due_date,
            output_dir=output_dir,
        )

        # Restore original annuity
        gs.ANNUITY_AMOUNT = original_annuity

        if created:
            names = "\n".join(f"  • {p.name}" for p in created)
            messagebox.showinfo(
                "Spreadsheets Created",
                f"✅ {len(created)} spreadsheet(s) created successfully!\n\n"
                f"{names}\n\nSaved to: {output_dir.resolve()}/",
                parent=dialog,
            )
            logger.info(
                f"Dialog created {len(created)} spreadsheet(s): "
                f"{[p.name for p in created]}"
            )
            dialog.destroy()
        else:
            widgets["status_label"].config(
                text="❌ No files were created. Check the logs.",
                fg=COLOR_ERROR,
            )

    except Exception as exc:
        logger.error(f"Error creating spreadsheets from dialog: {exc}", exc_info=True)
        messagebox.showerror(
            "Error",
            f"Failed to create spreadsheet(s):\n\n{exc}",
            parent=dialog,
        )
        widgets["status_label"].config(text=f"❌ Error: {exc}", fg=COLOR_ERROR)


# ─────────────────────────────────────────────
# Function 4 — _validate_form_inputs
# ─────────────────────────────────────────────

def _validate_form_inputs(
    raw_people: str,
    month: str,
    year: str,
    due_date: str,
    annuity: str,
) -> Optional[str]:
    """
    Validates all form inputs before attempting to create spreadsheets.

    Validation rules applied:
        1. People field must not be empty
        2. Each name must contain only letters, spaces, and hyphens
        3. Month must be one of the 12 Portuguese month names
        4. Year must be a 4-digit number between 2020 and 2035
        5. Due date must match DD/MM/YYYY format
        6. Annuity must be a positive number (0 is allowed)

    This function is intentionally strict to prevent malformed
    spreadsheets that would cause errors during invoice processing.

    Args:
        raw_people: Raw text from the people input field.
        month:      Selected month name in Portuguese (uppercase).
        year:       Selected year as a string.
        due_date:   Due date string entered by the user.
        annuity:    Annuity amount string entered by the user.

    Returns:
        An error message string if validation fails, or None if valid.
    """
    import re

    if not raw_people or not raw_people.strip():
        return "Please enter at least one person name."

    people = _parse_people_input(raw_people)
    if not people:
        return "No valid person names found. Enter one name per line."

    for name in people:
        if not re.match(r"^[A-Za-zÀ-ÿ\s\-]+$", name):
            return (
                f"Invalid name: '{name}'. "
                "Names must contain only letters, spaces, and hyphens."
            )

    if month not in PORTUGUESE_MONTHS:
        return f"Invalid month: '{month}'. Please select from the dropdown."

    if not year.isdigit() or not (2020 <= int(year) <= 2035):
        return f"Invalid year: '{year}'. Must be between 2020 and 2035."

    if not due_date:
        return "Due date is required. Format: DD/MM/YYYY"

    if not re.match(r"^\d{2}/\d{2}/\d{4}$", due_date):
        return f"Invalid due date format: '{due_date}'. Expected: DD/MM/YYYY"

    if not annuity:
        return "Annuity amount is required. Enter 0 if not applicable."

    try:
        annuity_value = float(annuity)
        if annuity_value < 0:
            return "Annuity amount must be zero or positive."
    except ValueError:
        return f"Invalid annuity amount: '{annuity}'. Must be a number (e.g. 18.75)"

    return None


# ─────────────────────────────────────────────
# Function 5 — _parse_people_input
# ─────────────────────────────────────────────

def _parse_people_input(raw_text: str) -> list[str]:
    """
    Parses the raw text from the people input field into a
    clean list of person name strings.

    Processing steps:
        1. Split by newlines to get one name per line
        2. Strip surrounding whitespace from each line
        3. Remove empty lines
        4. Collapse multiple internal spaces into one
        5. Deduplicate — remove repeated names (case-insensitive)
        6. Return names in the original input order

    This function is lenient about whitespace and blank lines
    to provide a good user experience, while ensuring the
    output is a clean, deduplicated list of names.

    Args:
        raw_text: The raw string from the people Text widget,
                  with one name per line.

    Returns:
        A list of clean, unique, non-empty name strings.
        Returns an empty list if no valid names are found.
    """
    import re

    if not raw_text or not isinstance(raw_text, str):
        return []

    lines = raw_text.strip().splitlines()
    seen: set[str] = set()
    people: list[str] = []

    for line in lines:
        # Strip and collapse whitespace
        clean = re.sub(r"\s+", " ", line.strip())

        if not clean:
            continue

        # Deduplicate case-insensitively
        normalized = clean.lower()
        if normalized in seen:
            logger.debug(f"Duplicate name skipped: '{clean}'")
            continue

        seen.add(normalized)
        people.append(clean)

    logger.debug(f"_parse_people_input: parsed {len(people)} name(s) from input.")
    return people