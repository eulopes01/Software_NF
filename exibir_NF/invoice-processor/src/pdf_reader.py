"""
pdf_reader.py — Phase 2: PDF Invoice Reader
--------------------------------------------
Responsible for reading, detecting, extracting and normalizing
credit card invoice data from Itaú and Banco do Brasil PDF files.

Each function has a single responsibility and at least 25 lines,
following Clean Architecture, SOLID, and Secure by Design principles.

Dependencies:
    pip install pdfplumber

Usage:
    from pdf_reader import read_pdf

    transactions, total = read_pdf("pdfs/invoice_john.pdf")
"""

import re
import logging
from pathlib import Path
from dataclasses import dataclass, field
from typing import Optional

import pdfplumber

# ─────────────────────────────────────────────
# Logging setup
# ─────────────────────────────────────────────

logger = logging.getLogger(__name__)

# ─────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────

BANK_ITAU          = "itau"
BANK_BANCO_BRASIL  = "banco_brasil"
BANK_UNKNOWN       = "unknown"

ITAU_IDENTIFIER       = "itaucard"
BB_IDENTIFIER         = "banco do brasil"

ITAU_SECTION_MARKER   = "lançamentos nacionais"
ITAU_TOTAL_MARKER     = "total desta fatura"
ITAU_END_MARKERS      = [
    "crédito do cartão final",
    "débito do cartão final",
    "total de créditos",
    "total de débitos",
]

# Matches: DD/MM  DESCRIPTION  VALUE
# Example: 03/10  MERCADOPAGO*MLIVRE  71,24
ITAU_TRANSACTION_PATTERN = re.compile(
    r"^(\d{2}/\d{2})\s+(.+?)\s+([-]?\d{1,3}(?:\.\d{3})*,\d{2})$"
)

# Matches: Total desta fatura   1.337,61
ITAU_TOTAL_PATTERN = re.compile(
    r"total desta fatura\s+([-]?\d{1,3}(?:\.\d{3})*,\d{2})",
    re.IGNORECASE,
)

MAX_FILE_SIZE_MB = 20
ALLOWED_EXTENSION = ".pdf"

# ─────────────────────────────────────────────
# Data model
# ─────────────────────────────────────────────

@dataclass
class Transaction:
    """Represents a single credit card transaction extracted from the invoice."""
    date: str            # Format: DD/MM (from the invoice)
    description: str     # Merchant or transaction description
    amount: float        # Transaction value in BRL (negative = credit)
    bank: str            # Source bank identifier


# ─────────────────────────────────────────────
# Function 1 — read_pdf
# ─────────────────────────────────────────────

def read_pdf(file_path: str) -> tuple[list[Transaction], float]:
    """
    Entry point for the PDF reading pipeline.

    Validates the file, detects the bank, delegates extraction,
    normalizes the results and returns structured transaction data.

    Args:
        file_path: Absolute or relative path to the PDF invoice file.

    Returns:
        A tuple of (list of Transaction, invoice total as float).

    Raises:
        ValueError: If the file is invalid, unsupported, or unreadable.
        FileNotFoundError: If the file does not exist at the given path.
    """
    path = Path(file_path)

    logger.info(f"Starting PDF read pipeline for: {path.name}")

    _validate_file(path)

    raw_text = _extract_raw_text(path)

    if not raw_text or not raw_text.strip():
        raise ValueError(
            f"No readable text found in '{path.name}'. "
            "The file may be scanned (image-based) or corrupted."
        )

    bank = detect_bank(raw_text)

    logger.info(f"Detected bank: {bank} for file: {path.name}")

    if bank == BANK_ITAU:
        transactions, total = extract_itau(raw_text, path.name)
    elif bank == BANK_BANCO_BRASIL:
        transactions, total = extract_banco_brasil(raw_text, path.name)
    else:
        raise ValueError(
            f"Unrecognized bank format in '{path.name}'. "
            f"Supported banks: Itaú, Banco do Brasil."
        )

    normalized = [normalize_transaction(t) for t in transactions]

    logger.info(
        f"Extraction complete: {len(normalized)} transactions | "
        f"Total: R$ {total:.2f} | File: {path.name}"
    )

    return normalized, total


# ─────────────────────────────────────────────
# Function 2 — detect_bank
# ─────────────────────────────────────────────

def detect_bank(raw_text: str) -> str:
    """
    Identifies which bank issued the invoice based on keywords
    found in the extracted PDF text.

    The detection is case-insensitive and searches the first
    500 characters of the document, where bank identifiers
    are consistently present in both Itaú and BB invoices.

    Args:
        raw_text: Full plain text extracted from the PDF pages.

    Returns:
        One of the bank constants: BANK_ITAU, BANK_BANCO_BRASIL,
        or BANK_UNKNOWN if detection fails.
    """
    if not raw_text or not isinstance(raw_text, str):
        logger.warning("detect_bank received empty or invalid text.")
        return BANK_UNKNOWN

    # Sample the beginning of the document for faster and safer detection.
    # Both banks place their identifier in the first visible block.
    sample = raw_text[:500].lower().strip()

    if not sample:
        logger.warning("PDF text sample is empty after stripping.")
        return BANK_UNKNOWN

    if ITAU_IDENTIFIER in sample:
        logger.debug("Bank detected: Itaú (matched 'itaucard')")
        return BANK_ITAU

    if BB_IDENTIFIER in sample:
        logger.debug("Bank detected: Banco do Brasil (matched 'banco do brasil')")
        return BANK_BANCO_BRASIL

    # If the identifier is not in the first 500 chars, widen the search
    # to account for different PDF layouts with headers or images.
    full_lower = raw_text.lower()

    if ITAU_IDENTIFIER in full_lower:
        logger.debug("Bank detected (full scan): Itaú")
        return BANK_ITAU

    if BB_IDENTIFIER in full_lower:
        logger.debug("Bank detected (full scan): Banco do Brasil")
        return BANK_BANCO_BRASIL

    logger.warning(f"Could not detect bank. Text sample: {sample[:100]!r}")
    return BANK_UNKNOWN


# ─────────────────────────────────────────────
# Function 3 — extract_itau
# ─────────────────────────────────────────────

def extract_itau(raw_text: str, file_name: str = "") -> tuple[list[Transaction], float]:
    """
    Extracts all national transactions and the invoice total
    from an Itaú credit card invoice PDF text.

    Itaú invoice structure (per page):
      - "Lançamentos nacionais"
      - "HOLDER NAME (card_last_digits)"
      - "DATA  MOVIMENTAÇÃO  VALOR EM R$"
      - Rows: DD/MM  DESCRIPTION  VALUE
      - "Crédito do cartão final (XXXX)  X,XX"
      - "Débito do cartão final (XXXX)  X,XX"

    The extractor handles multiple cardholders in the same invoice,
    negative values (credits/reversals), and parceled purchases.
    Skips summary lines, header lines, and non-transaction content.

    Args:
        raw_text: Full plain text extracted from the PDF.
        file_name: Original file name used for logging purposes only.

    Returns:
        A tuple of (list of Transaction, invoice total as float).
    """
    transactions: list[Transaction] = []
    invoice_total: float = 0.0
    inside_section = False

    lines = raw_text.splitlines()

    for line in lines:
        stripped = line.strip()

        if not stripped:
            continue

        # Detect total invoice value
        total_match = ITAU_TOTAL_PATTERN.search(stripped)
        if total_match:
            invoice_total = _parse_brazilian_float(total_match.group(1))
            logger.debug(f"Invoice total found: R$ {invoice_total:.2f}")
            continue

        # Enter transaction section
        lower = stripped.lower()
        if ITAU_SECTION_MARKER in lower:
            inside_section = True
            continue

        # Exit transaction section on summary lines
        if inside_section and any(marker in lower for marker in ITAU_END_MARKERS):
            inside_section = False
            continue

        if not inside_section:
            continue

        # Skip section headers and column headers
        if lower.startswith("data") or lower.startswith("clayton") or lower.startswith("titular"):
            continue

        # Attempt to match a transaction line
        match = ITAU_TRANSACTION_PATTERN.match(stripped)
        if not match:
            logger.debug(f"[itau] Skipped non-transaction line: {stripped!r}")
            continue

        date, description, raw_value = match.groups()
        amount = _parse_brazilian_float(raw_value)

        transactions.append(Transaction(
            date=date,
            description=description.strip(),
            amount=amount,
            bank=BANK_ITAU,
        ))

    logger.info(
        f"[itau] Extracted {len(transactions)} transactions "
        f"| Total: R$ {invoice_total:.2f} | File: {file_name}"
    )

    return transactions, invoice_total


# ─────────────────────────────────────────────
# Function 4 — extract_banco_brasil
# ─────────────────────────────────────────────

def extract_banco_brasil(raw_text: str, file_name: str = "") -> tuple[list[Transaction], float]:
    """
    Extracts all transactions and the invoice total from a
    Banco do Brasil (BB) credit card invoice PDF text.

    BB invoice structure:
      - Header with "Banco do Brasil" branding
      - "LANÇAMENTOS" or "Compras e Saques" section
      - Rows: DD/MM/YYYY  DESCRIPTION  VALUE
      - "Total da fatura" or "Total a pagar" at the end

    Note: BB PDF layouts may vary depending on the card type
    (Ourocard, Visa, Mastercard). This extractor targets the
    most common layout. If parsing fails, the function returns
    an empty list with total=0.0 and logs a warning.

    Args:
        raw_text: Full plain text extracted from the PDF.
        file_name: Original file name used for logging purposes only.

    Returns:
        A tuple of (list of Transaction, invoice total as float).
    """
    transactions: list[Transaction] = []
    invoice_total: float = 0.0

    # BB uses DD/MM/YYYY in transaction lines
    bb_transaction_pattern = re.compile(
        r"^(\d{2}/\d{2}/\d{4})\s+(.+?)\s+([-]?\d{1,3}(?:\.\d{3})*,\d{2})$"
    )
    bb_total_pattern = re.compile(
        r"total\s+(?:da\s+fatura|a\s+pagar)\s+([-]?\d{1,3}(?:\.\d{3})*,\d{2})",
        re.IGNORECASE,
    )

    bb_section_markers = ["lançamentos", "compras e saques", "extrato"]
    bb_end_markers     = ["total da fatura", "total a pagar", "encargos"]

    inside_section = False
    lines = raw_text.splitlines()

    for line in lines:
        stripped = line.strip()

        if not stripped:
            continue

        lower = stripped.lower()

        # Detect invoice total
        total_match = bb_total_pattern.search(stripped)
        if total_match:
            invoice_total = _parse_brazilian_float(total_match.group(1))
            logger.debug(f"[bb] Invoice total found: R$ {invoice_total:.2f}")
            continue

        # Enter transaction section
        if any(marker in lower for marker in bb_section_markers):
            inside_section = True
            continue

        # Exit on summary markers (but capture total first above)
        if inside_section and any(marker in lower for marker in bb_end_markers):
            inside_section = False
            continue

        if not inside_section:
            continue

        match = bb_transaction_pattern.match(stripped)
        if not match:
            logger.debug(f"[bb] Skipped non-transaction line: {stripped!r}")
            continue

        date_raw, description, raw_value = match.groups()
        # Normalize BB date DD/MM/YYYY → DD/MM to match Itaú format
        date = date_raw[:5]
        amount = _parse_brazilian_float(raw_value)

        transactions.append(Transaction(
            date=date,
            description=description.strip(),
            amount=amount,
            bank=BANK_BANCO_BRASIL,
        ))

    if not transactions:
        logger.warning(
            f"[bb] No transactions extracted from '{file_name}'. "
            "Layout may differ from expected. Manual review recommended."
        )

    logger.info(
        f"[bb] Extracted {len(transactions)} transactions "
        f"| Total: R$ {invoice_total:.2f} | File: {file_name}"
    )

    return transactions, invoice_total


# ─────────────────────────────────────────────
# Function 5 — normalize_transaction
# ─────────────────────────────────────────────

def normalize_transaction(transaction: Transaction) -> Transaction:
    """
    Sanitizes and standardizes a Transaction object to ensure
    consistent, safe data before it is written to the spreadsheet.

    Normalization rules applied:
      - Description: stripped of extra whitespace, truncated to 100 chars,
        control characters removed, safe for Excel cell insertion.
      - Date: validated as DD/MM pattern; set to "00/00" if invalid.
      - Amount: validated as float; set to 0.0 if corrupted.
      - Bank: lowercased and stripped.

    This function is the last line of defense before data reaches
    the spreadsheet writer. It must never raise — it logs warnings
    and returns a sanitized version of the input instead.

    Args:
        transaction: A raw Transaction object as returned by extract_*.

    Returns:
        A new Transaction object with normalized, safe field values.
    """
    if not isinstance(transaction, Transaction):
        logger.error(
            f"normalize_transaction received non-Transaction input: "
            f"{type(transaction)}. Returning empty transaction."
        )
        return Transaction(date="00/00", description="INVALID", amount=0.0, bank="unknown")

    # Sanitize description
    raw_desc = transaction.description or ""
    # Remove control characters (ASCII 0–31 except tab) — prevents Excel injection
    clean_desc = re.sub(r"[\x00-\x08\x0b-\x1f]", "", raw_desc)
    # Collapse multiple whitespace into single space
    clean_desc = re.sub(r"\s+", " ", clean_desc).strip()
    # Prevent Excel formula injection (=, +, -, @ at start of cell)
    if clean_desc and clean_desc[0] in ("=", "+", "-", "@", "|", "%"):
        clean_desc = "'" + clean_desc
    # Truncate to safe Excel cell length
    clean_desc = clean_desc[:100]

    # Validate date format DD/MM
    raw_date = transaction.date or ""
    if not re.match(r"^\d{2}/\d{2}$", raw_date):
        logger.warning(f"Invalid date format '{raw_date}' — setting to '00/00'.")
        raw_date = "00/00"

    # Validate amount
    raw_amount = transaction.amount
    if not isinstance(raw_amount, (int, float)) or raw_amount != raw_amount:
        # NaN check: float NaN != float NaN
        logger.warning(f"Invalid amount '{raw_amount}' — setting to 0.0.")
        raw_amount = 0.0

    # Sanitize bank name
    clean_bank = str(transaction.bank or "unknown").lower().strip()[:50]

    return Transaction(
        date=raw_date,
        description=clean_desc,
        amount=float(raw_amount),
        bank=clean_bank,
    )


# ─────────────────────────────────────────────
# Private helpers
# ─────────────────────────────────────────────

def _validate_file(path: Path) -> None:
    """
    Validates that the given path points to a real, readable,
    non-empty PDF file within the allowed size limit.

    Raises:
        FileNotFoundError: If the file does not exist.
        ValueError: If the file is not a PDF, is empty, or exceeds size limit.
    """
    if not path.exists():
        raise FileNotFoundError(f"Invoice file not found: '{path}'")

    if not path.is_file():
        raise ValueError(f"Path is not a file: '{path}'")

    if path.suffix.lower() != ALLOWED_EXTENSION:
        raise ValueError(
            f"Invalid file type '{path.suffix}'. "
            f"Only '{ALLOWED_EXTENSION}' files are supported."
        )

    file_size_mb = path.stat().st_size / (1024 * 1024)
    if file_size_mb > MAX_FILE_SIZE_MB:
        raise ValueError(
            f"File '{path.name}' is too large ({file_size_mb:.1f} MB). "
            f"Maximum allowed size is {MAX_FILE_SIZE_MB} MB."
        )

    if path.stat().st_size == 0:
        raise ValueError(f"File '{path.name}' is empty.")

    logger.debug(f"File validation passed: {path.name} ({file_size_mb:.2f} MB)")


def _extract_raw_text(path: Path) -> str:
    """
    Opens a PDF file using pdfplumber and extracts all text
    from all pages, concatenated with newline separators.

    Raises:
        ValueError: If pdfplumber fails to open or read the file.
    """
    try:
        all_pages_text = []
        with pdfplumber.open(str(path)) as pdf:
            if not pdf.pages:
                raise ValueError(f"PDF '{path.name}' has no readable pages.")
            for page_number, page in enumerate(pdf.pages, start=1):
                page_text = page.extract_text()
                if page_text:
                    all_pages_text.append(page_text)
                else:
                    logger.debug(f"Page {page_number} of '{path.name}' returned no text.")
        return "\n".join(all_pages_text)
    except Exception as exc:
        raise ValueError(
            f"Failed to read PDF '{path.name}': {exc}"
        ) from exc


def _parse_brazilian_float(value_str: str) -> float:
    """
    Converts a Brazilian-formatted currency string to a Python float.

    Brazilian format: 1.337,61 → 1337.61
    Handles negative values:  -1.598,21 → -1598.21

    Args:
        value_str: Raw currency string from the PDF.

    Returns:
        Float value. Returns 0.0 if parsing fails.
    """
    if not value_str or not isinstance(value_str, str):
        logger.warning(f"_parse_brazilian_float received invalid input: {value_str!r}")
        return 0.0
    try:
        cleaned = value_str.strip().replace(".", "").replace(",", ".")
        return float(cleaned)
    except ValueError:
        logger.warning(f"Could not parse currency value: {value_str!r}")
        return 0.0