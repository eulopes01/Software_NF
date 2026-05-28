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

try:
    from pdf_ocr import extract_text_from_scanned_pdf, is_scanned_pdf, check_tesseract_installation
    TESSERACT_AVAILABLE = check_tesseract_installation()["available"]
except ImportError:
    TESSERACT_AVAILABLE = False
    logger = logging.getLogger(__name__)
    logger.warning("pdf_ocr module not found — OCR fallback disabled.")

# ─────────────────────────────────────────────
# Logging setup
# ─────────────────────────────────────────────

logger = logging.getLogger(__name__)

# ─────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────

BANK_ITAU          = "itau"
BANK_ITAU_EMPRESAS = "itau_empresas"
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
# Also matches lines where date appears after other text:
# Example: Financiamento da Fatura 05/10 POSTO AUSTRAL 10,66
ITAU_TRANSACTION_PATTERN = re.compile(
    r"^(\d{2}/\d{2})\s+(.+?)\s+([-]?\d{1,3}(?:\.\d{3})*,\d{2})$"
)

# Fallback pattern for lines where date is not at the start
ITAU_TRANSACTION_PATTERN_INLINE = re.compile(
    r"(?:^|\s)(\d{2}/\d{2})\s+([A-Z][A-Z\s\*\.0-9/]+?)\s+([-]?\d{1,3}(?:\.\d{3})*,\d{2})$"
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

@dataclass
class CardHolder:
    """Represents a cardholder section extracted from the invoice."""
    name: str                        # Full name of the cardholder
    card_last_digits: str            # Last 4 digits of the card
    transactions: list               # List of Transaction objects
    bank: str                        # Source bank identifier

# Pattern to detect cardholder name and card digits
# Matches: "CLAYTON PIRES DOS SANTOS (3172)"
# Also matches with leading date text: "06/10/14 e 05/11/14 CLAYTON PIRES DOS SANTOS (5885)"
# ─────────────────────────────────────────────
# Itaú Empresas constants
# ─────────────────────────────────────────────

# Matches: "Lançamentos no cartão (final 3172)"
ITAU_EMP_SECTION_START  = "lançamentos no cartão (final"
ITAU_EMP_SECTION_END    = "lançamentos no cartão (final"

# Matches: "NOME DA PESSOA (final 7221)"
ITAU_EMP_HOLDER_PATTERN = re.compile(
    r"^([A-ZÁÉÍÓÚÀÃÕÂÊÎÔÛÇ][A-ZÁÉÍÓÚÀÃÕÂÊÎÔÛÇ\s]+?)\s*\(final\s+(\d{4})\)",
    re.IGNORECASE,
)

# Matches transaction line: "09/10 MERCADOLIVRE*24PRO07/12"
ITAU_EMP_TX_DATE_DESC = re.compile(
    r"^(\d{2}/\d{2})\s+(.+?)(?:\s+\d{2}/\d{2})?$"
)

# Matches value line: "206,35" or "1.226,00"
ITAU_EMP_VALUE_PATTERN = re.compile(
    r"^([-]?\d{1,3}(?:\.\d{3})*,\d{2})$"
)

# Itaú Empresas total pattern
ITAU_EMP_TOTAL_PATTERN = re.compile(
    r"total\s+desta\s+fatura\s+([-]?\d{1,3}(?:\.\d{3})*,\d{2})",
    re.IGNORECASE,
)

ITAU_EMP_IDENTIFIER    = "itaú empresas"

ITAU_EMP_SKIP_MARKERS  = [
    "lançamentos internacionais",
    "lançamentos: produtos e serviços",
    "data estabelecimento",
    "data produtos",
    "valor em r$",
    "limite de gastos",
    "limite retirada",
    "dólar de conversão",
    "total transações",
    "repasse de iof",
    "total lançamentos",
]

ITAU_EMP_CARDHOLDER_PATTERN = re.compile(
    r"^([A-ZÁÉÍÓÚÀÃÕÂÊÎÔÛÇ][A-ZÁÉÍÓÚÀÃÕÂÊÎÔÛÇ\s]+?)\s*\(final\s+(\d{4})\)",
    re.IGNORECASE,
)

ITAU_CARDHOLDER_PATTERN = re.compile(
    r"([A-ZÁÉÍÓÚÀÃÕÂÊÎÔÛÇ][A-ZÁÉÍÓÚÀÃÕÂÊÎÔÛÇ\s]+?)\s*\((\d{4})\)\s*$"
)


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
    elif bank == BANK_ITAU_EMPRESAS:
        cardholders, total = extract_itau_empresas_by_cardholder(raw_text, path.name)
        transactions = [t for ch in cardholders for t in ch.transactions]
    elif bank == BANK_BANCO_BRASIL:
        transactions, total = extract_banco_brasil(raw_text, path.name)
    else:
        raise ValueError(
            f"Unrecognized bank format in '{path.name}'. "
            f"Supported banks: Itaú, Itaú Empresas, Banco do Brasil."
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

    Detection order (important — Itaú Empresas before generic Itaú):
        1. Itaú Empresas  → "itauempresas", "cartaoitauempresas"
        2. Itaú personal  → "itau"
        3. Banco do Brasil → "banco do brasil"

    Uses accent normalization (NFKD) for robust matching of
    Portuguese text that may have or lack accents depending on
    the PDF extraction quality.

    Args:
        raw_text: Full plain text extracted from the PDF pages.

    Returns:
        One of: BANK_ITAU_EMPRESAS, BANK_ITAU, BANK_BANCO_BRASIL,
        or BANK_UNKNOWN if detection fails.
    """
    import unicodedata

    if not raw_text or not isinstance(raw_text, str):
        logger.warning("detect_bank received empty or invalid text.")
        return BANK_UNKNOWN

    def normalize(text: str) -> str:
        """Remove accents and lowercase for robust matching."""
        return unicodedata.normalize("NFKD", text.lower()).encode("ascii", "ignore").decode("ascii")

    sample = normalize(raw_text[:1000])
    full   = normalize(raw_text)

    if not sample:
        logger.warning("PDF text sample is empty after stripping.")
        return BANK_UNKNOWN

    # Itaú Empresas — checked BEFORE generic Itaú to avoid false positives
    itau_empresas_signals = [
        "itau empresas",
        "cartaoitauempresas",
        "itauempresas",
        "cartao itau empresas",
        "lancamentos no cartao (final",
        "lançamentos no cartão (final",
    ]
    for signal in itau_empresas_signals:
        normalized_signal = normalize(signal)
        if normalized_signal in sample or normalized_signal in full:
            logger.debug(f"Bank detected: Itaú Empresas (matched '{signal}')")
            return BANK_ITAU_EMPRESAS

    # Generic Itaú (personal cards)
    if "itau" in sample or "itau" in full:
        logger.debug("Bank detected: Itaú")
        return BANK_ITAU

    # Banco do Brasil
    if "banco do brasil" in sample or "banco do brasil" in full:
        logger.debug("Bank detected: Banco do Brasil")
        return BANK_BANCO_BRASIL

    logger.warning(f"Could not detect bank. Text sample: {raw_text[:100]!r}")
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

        # Attempt to match a transaction line (primary pattern)
        match = ITAU_TRANSACTION_PATTERN.match(stripped)

        # Try fallback pattern for lines with leading text before the date
        if not match:
            match = ITAU_TRANSACTION_PATTERN_INLINE.search(stripped)

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
# Function 3b — extract_itau_by_cardholder
# ─────────────────────────────────────────────

def extract_itau_by_cardholder(raw_text: str, file_name: str = "") -> tuple[list, float]:
    """
    Extracts Itaú invoice transactions grouped by cardholder.

    Each 'Lançamentos nacionais' section belongs to one cardholder
    with a specific card number (last 4 digits). This function
    identifies each section, extracts the cardholder name and card
    digits, and groups transactions accordingly.

    Args:
        raw_text: Full plain text extracted from the PDF.
        file_name: Original file name for logging purposes.

    Returns:
        A tuple of (list of CardHolder, invoice total as float).
    """
    cardholders: list[CardHolder] = []
    invoice_total: float = 0.0
    lines = raw_text.splitlines()

    total_match = ITAU_TOTAL_PATTERN.search(raw_text)
    if total_match:
        invoice_total = _parse_brazilian_float(total_match.group(1))

    i = 0
    while i < len(lines):
        line = lines[i].strip()

        # Detect start of a cardholder section
        if ITAU_SECTION_MARKER in line.lower():
            # Look at the next non-empty line for the cardholder name
            j = i + 1
            holder_name = ""
            card_digits = ""

            while j < len(lines) and j < i + 5:
                candidate = lines[j].strip()
                match = ITAU_CARDHOLDER_PATTERN.search(candidate)
                if match:
                    holder_name = match.group(1).strip()
                    card_digits = match.group(2).strip()
                    break
                j += 1

            if not holder_name:
                i += 1
                continue

            # Extract transactions for this cardholder
            transactions = []
            k = j + 1
            while k < len(lines):
                tx_line = lines[k].strip()
                lower = tx_line.lower()

                # Stop at end markers or next section
                if any(marker in lower for marker in ITAU_END_MARKERS):
                    break
                if ITAU_SECTION_MARKER in lower:
                    break

                # Skip headers
                if lower.startswith("data") or lower.startswith("titular"):
                    k += 1
                    continue

                # Try to match transaction
                tx_match = ITAU_TRANSACTION_PATTERN.match(tx_line)
                if not tx_match:
                    tx_match = ITAU_TRANSACTION_PATTERN_INLINE.search(tx_line)

                if tx_match:
                    date, description, raw_value = tx_match.groups()
                    amount = _parse_brazilian_float(raw_value)
                    transactions.append(Transaction(
                        date=date,
                        description=description.strip(),
                        amount=amount,
                        bank=BANK_ITAU,
                    ))
                k += 1

            # Normalize transactions
            normalized = [normalize_transaction(t) for t in transactions]

            cardholders.append(CardHolder(
                name=holder_name,
                card_last_digits=card_digits,
                transactions=normalized,
                bank=BANK_ITAU,
            ))

            logger.info(
                f"[itau] Cardholder: '{holder_name}' (card ...{card_digits}) "
                f"→ {len(normalized)} transactions"
            )
            i = k
            continue

        i += 1

    logger.info(
        f"[itau] Total cardholders: {len(cardholders)} "
        f"| Invoice total: R$ {invoice_total:.2f} | File: {file_name}"
    )
    return cardholders, invoice_total



# ─────────────────────────────────────────────
# Function 3c — extract_itau_empresas_by_cardholder
# ─────────────────────────────────────────────

def extract_itau_empresas_by_cardholder(raw_text: str, file_name: str = "") -> tuple[list, float]:
    """
    Extracts Itaú Empresas invoice transactions grouped by cardholder.

    Itaú Empresas (qpdf-decrypted) layout — text is compressed without spaces:
        Total:         "LTotaldoslançamentosatuais 32.287,55"
        Cardholder:    "NOMESOBRENOME(finalXXXX)"
        Transactions:  "DD/MM DESCRIPTION VALUE" on same line
        Two per line:  "DD/MM DESC VAL DD/MM DESC VAL"

    Args:
        raw_text: Full plain text extracted from the PDF.
        file_name: Original file name for logging purposes.

    Returns:
        A tuple of (list of CardHolder, invoice total as float).
    """
    import unicodedata

    def norm(text: str) -> str:
        return unicodedata.normalize("NFKD", text.lower()).encode("ascii", "ignore").decode("ascii")

    cardholders: list = []
    invoice_total: float = 0.0

    # Find invoice total
    total_pats = [
        re.compile(r"[Ll]?[Tt]otaldosl[aã]nçamentosatuais\s+([\d.,]+)", re.IGNORECASE),
        re.compile(r"=?[Tt]otaldestafatura\s+([\d.,]+)", re.IGNORECASE),
        re.compile(r"[Ll]{1,2}[aã]nçamentosatuais\s+([\d.,]+)", re.IGNORECASE),
    ]
    for pat in total_pats:
        m = pat.search(raw_text)
        if m:
            invoice_total = _parse_brazilian_float(m.group(1))
            logger.debug(f"[itau_empresas] Invoice total: R$ {invoice_total:.2f}")
            break

    # Cardholder pattern: NAME(finalXXXX)
    holder_pat = re.compile(
        r"([A-ZÀ-Ü]{3,}(?:\s+[A-ZÀ-Ü]{2,})*)\s*\(final\s*(\d{4})\)",
        re.IGNORECASE,
    )

    # Transaction: DD/MM DESCRIPTION VALUE (optional suffix CL/CT)
    tx_pat = re.compile(
        r"^(\d{2}/\d{2})\s+(.+?)\s+([-]?\d{1,3}(?:[.]\d{3})*,\d{2})\s*(?:CL|CT|DIF)?$"
    )

    # Two transactions on same line
    multi_pat = re.compile(
        r"(\d{2}/\d{2})\s+([A-Z][^\d\n]+?)\s+([-]?\d{1,3}(?:[.]\d{3})*,\d{2})"
    )

    # Skip markers (normalized)
    skip = [
        "limitedegastos", "limiteretirada", "dataestabelecimento",
        "data estabelecimento", "valoremr", "lancamentosproduto",
        "lançamentosproduto", "anuidade", "estorno", "novotetodejuros",
        "jurosdacompra", "cetdacompra", "encargos", "limitemaximo",
        "lançamentosnocar", "lancamentosnocar", "previsaodo",
        "seguefatura", "pagamentominimo", "parcelasfixas",
    ]

    lines = raw_text.splitlines()
    i = 0

    while i < len(lines):
        line      = lines[i].strip()
        line_norm = norm(line)

        # Skip section total lines (not cardholders)
        if re.search(r"[Ll]ançamentosnocar|[Ll]ancamentosnocar", line, re.IGNORECASE):
            i += 1
            continue

        holder_match = holder_pat.search(line)
        if holder_match:
            holder_name = holder_match.group(1).strip()
            card_digits = holder_match.group(2).strip()
            transactions = []

            # Check if cardholder line itself contains a transaction
            # Example: "CAMILAMHSANTOS(final0667) 01/05 STARLINKINTERNET 211,97"
            inline_tx = re.search(
                r"\)\s+(\d{2}/\d{2})\s+(.+?)\s+([-]?\d{1,3}(?:[.]\d{3})*,\d{2})\s*$",
                line
            )
            if inline_tx:
                amount = _parse_brazilian_float(inline_tx.group(3))
                transactions.append(Transaction(
                    date=inline_tx.group(1),
                    description=inline_tx.group(2).strip(),
                    amount=amount,
                    bank=BANK_ITAU_EMPRESAS,
                ))

            j = i + 1

            while j < len(lines):
                tx_line      = lines[j].strip()
                tx_norm      = norm(tx_line)

                # Stop at next cardholder section
                if holder_pat.search(tx_line) and j > i + 1:
                    break

                # Stop at products/services section
                if "lancamentosproduto" in tx_norm or "lançamentosproduto" in tx_norm:
                    break

                # Skip metadata lines
                if any(mk in tx_norm for mk in skip):
                    j += 1
                    continue

                # Try single transaction match
                tx_m = tx_pat.match(tx_line)
                if tx_m:
                    amount = _parse_brazilian_float(tx_m.group(3))
                    transactions.append(Transaction(
                        date=tx_m.group(1),
                        description=tx_m.group(2).strip(),
                        amount=amount,
                        bank=BANK_ITAU_EMPRESAS,
                    ))
                    j += 1
                    continue

                # Try multiple transactions on same line
                multi_matches = multi_pat.findall(tx_line)
                if len(multi_matches) >= 2:
                    for date, desc, raw_val in multi_matches:
                        amount = _parse_brazilian_float(raw_val)
                        transactions.append(Transaction(
                            date=date,
                            description=desc.strip(),
                            amount=amount,
                            bank=BANK_ITAU_EMPRESAS,
                        ))
                    j += 1
                    continue

                j += 1

            normalized = [normalize_transaction(t) for t in transactions]

            cardholders.append(CardHolder(
                name=holder_name,
                card_last_digits=card_digits,
                transactions=normalized,
                bank=BANK_ITAU_EMPRESAS,
            ))

            logger.info(
                f"[itau_empresas] Cardholder: '{holder_name}' "
                f"(card ...{card_digits}) -> {len(normalized)} transactions"
            )
            i = j
            continue

        i += 1

    logger.info(
        f"[itau_empresas] Total cardholders: {len(cardholders)} "
        f"| Invoice total: R$ {invoice_total:.2f} | File: {file_name}"
    )
    return cardholders, invoice_total



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

        full_text = "\n".join(all_pages_text)

        # Fallback to OCR if no text was extracted (scanned PDF)
        if not full_text.strip() and TESSERACT_AVAILABLE:
            logger.info(
                f"No text extracted from '{path.name}' via pdfplumber. "
                "Falling back to OCR (Tesseract)..."
            )
            full_text = extract_text_from_scanned_pdf(str(path))

        elif not full_text.strip() and not TESSERACT_AVAILABLE:
            raise ValueError(
                f"No readable text found in '{path.name}'. "
                "The file appears to be scanned (image-based). "
                "Install Tesseract to enable OCR: "
                "sudo apt install tesseract-ocr tesseract-ocr-por tesseract-ocr-eng"
            )

        return full_text

    except ValueError:
        raise
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