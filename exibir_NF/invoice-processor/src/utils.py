"""
utils.py — Shared Utility Functions
-------------------------------------
Provides reusable helper functions used across all modules:
pdf_reader, spreadsheet_writer, reconciler, and main.

Each function has a single responsibility and at least 25 lines,
following Clean Architecture, SOLID, and Secure by Design principles.

No external dependencies — uses only Python standard library.

Usage:
    from utils import format_currency, format_date, validate_file,
                      sanitize_text, load_config
"""

import json
import logging
import os
import re
from datetime import datetime
from pathlib import Path
from typing import Any, Optional

# ─────────────────────────────────────────────
# Logging setup
# ─────────────────────────────────────────────

logger = logging.getLogger(__name__)

# ─────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────

# Supported file extensions per type
ALLOWED_PDF_EXTENSION   = ".pdf"
ALLOWED_XLSX_EXTENSION  = ".xlsx"

# File size limits
MAX_PDF_SIZE_MB         = 20
MAX_XLSX_SIZE_MB        = 20

# Text sanitization limits
MAX_TEXT_LENGTH         = 500
MAX_CELL_LENGTH         = 100

# Date formats
INVOICE_DATE_FORMAT     = "%d/%m"           # DD/MM  — from PDF invoices
FULL_DATE_FORMAT        = "%d/%m/%Y"        # DD/MM/YYYY — full date
ISO_DATE_FORMAT         = "%Y-%m-%d"        # ISO 8601

# Brazilian currency pattern: 1.234,56 or -1.234,56
BRAZILIAN_CURRENCY_RE   = re.compile(
    r"^-?\d{1,3}(?:\.\d{3})*,\d{2}$"
)

# Characters that trigger Excel formula injection
EXCEL_INJECTION_CHARS   = ("=", "+", "-", "@", "|", "%", "\t", "\r")

# Default config file path
DEFAULT_CONFIG_PATH     = Path("config.json")

# Default configuration values used when config.json is absent
DEFAULT_CONFIG: dict[str, Any] = {
    "log_level":         "INFO",
    "log_file":          "logs/processing.log",
    "spreadsheet_dir":   "spreadsheets",
    "pdf_dir":           "pdfs",
    "template_path":     "templates/card_template.xlsx",
    "payment_method":    "A VISTA",
    "reconcile_tolerance": 0.01,
    "backup_enabled":    True,
    "max_transactions":  500,
    "reference_month":   "",
}


# ─────────────────────────────────────────────
# Function 1 — format_currency
# ─────────────────────────────────────────────

def format_currency(value: float, symbol: str = "R$") -> str:
    """
    Formats a float value as a Brazilian Real currency string.

    Uses Brazilian number formatting conventions:
        - Thousands separator: period (.)
        - Decimal separator: comma (,)
        - Prefix: currency symbol followed by a space

    Examples:
        format_currency(1337.61)     → "R$ 1.337,61"
        format_currency(-50.00)      → "R$ -50,00"
        format_currency(0.0)         → "R$ 0,00"
        format_currency(1000000.99)  → "R$ 1.000.000,99"

    Args:
        value:  The numeric value to format. Must be int or float.
        symbol: The currency symbol prefix. Defaults to "R$".

    Returns:
        A formatted currency string in Brazilian format.

    Raises:
        TypeError: If value is not a numeric type.
    """
    if not isinstance(value, (int, float)):
        raise TypeError(
            f"format_currency expects int or float. "
            f"Received: {type(value).__name__} ({value!r})"
        )

    if value != value:
        # NaN check
        logger.warning("format_currency received NaN — returning 'R$ 0,00'")
        return f"{symbol} 0,00"

    is_negative = value < 0
    abs_value = abs(value)

    # Format with 2 decimal places using standard float formatting
    # then convert to Brazilian notation
    formatted = f"{abs_value:,.2f}"               # → "1,337.61"
    formatted = formatted.replace(",", "X")        # → "1X337.61"
    formatted = formatted.replace(".", ",")        # → "1X337,61"
    formatted = formatted.replace("X", ".")        # → "1.337,61"

    sign = "-" if is_negative else ""

    result = f"{symbol} {sign}{formatted}"
    logger.debug(f"format_currency: {value} → '{result}'")
    return result


# ─────────────────────────────────────────────
# Function 2 — format_date
# ─────────────────────────────────────────────

def format_date(
    date_str: str,
    input_format: str = INVOICE_DATE_FORMAT,
    output_format: str = FULL_DATE_FORMAT,
    reference_year: Optional[int] = None,
) -> str:
    """
    Parses and reformats a date string from one format to another.

    Primarily used to convert the DD/MM format found in Itaú invoices
    to DD/MM/YYYY by injecting the current or reference year.

    Examples:
        format_date("03/10")                      → "03/10/2025"
        format_date("03/10/2025", "%d/%m/%Y")     → "03/10/2025"
        format_date("2025-10-03", "%Y-%m-%d",
                    output_format="%d/%m/%Y")      → "03/10/2025"

    Args:
        date_str:        The raw date string to convert.
        input_format:    strptime format of the input string.
                         Defaults to "%d/%m" (Itaú invoice format).
        output_format:   strftime format of the output string.
                         Defaults to "%d/%m/%Y".
        reference_year:  Year to inject when input_format has no year.
                         Defaults to the current year if not provided.

    Returns:
        The reformatted date string, or the original string if parsing fails.
    """
    if not date_str or not isinstance(date_str, str):
        logger.warning(f"format_date received invalid input: {date_str!r}")
        return date_str or ""

    cleaned = date_str.strip()

    if not cleaned:
        logger.warning("format_date received empty string after stripping.")
        return ""

    # Inject reference year when input format has no year component
    parse_str = cleaned
    parse_fmt = input_format

    if "%Y" not in input_format and "%y" not in input_format:
        year = reference_year or datetime.now().year
        parse_str = f"{cleaned}/{year}"
        parse_fmt = f"{input_format}/%Y"

    try:
        parsed_date = datetime.strptime(parse_str, parse_fmt)
        result = parsed_date.strftime(output_format)
        logger.debug(f"format_date: '{date_str}' → '{result}'")
        return result
    except ValueError as exc:
        logger.warning(
            f"format_date could not parse '{date_str}' "
            f"with format '{input_format}': {exc}. "
            "Returning original string."
        )
        return date_str


# ─────────────────────────────────────────────
# Function 3 — validate_file
# ─────────────────────────────────────────────

def validate_file(
    file_path: str,
    allowed_extension: str,
    max_size_mb: float = MAX_PDF_SIZE_MB,
) -> Path:
    """
    Validates that a file path points to a real, readable, non-empty
    file of the correct type within the allowed size limit.

    Performs the following checks in order:
        1. Path is not empty or None
        2. Path exists on disk
        3. Path points to a file (not a directory)
        4. File extension matches the allowed extension
        5. File size is within the allowed limit
        6. File is not empty (size > 0 bytes)

    Used as the first line of defense in pdf_reader and
    spreadsheet_writer before attempting to open any file.

    Args:
        file_path:         Absolute or relative path to the file.
        allowed_extension: Required file extension (e.g. ".pdf", ".xlsx").
                           Case-insensitive comparison is applied.
        max_size_mb:       Maximum allowed file size in megabytes.
                           Defaults to MAX_PDF_SIZE_MB (20 MB).

    Returns:
        A resolved Path object if all validations pass.

    Raises:
        ValueError: If file_path is empty or None.
        FileNotFoundError: If the file does not exist.
        ValueError: If any validation check fails.
    """
    if not file_path or not isinstance(file_path, str):
        raise ValueError(
            f"file_path must be a non-empty string. "
            f"Received: {file_path!r}"
        )

    path = Path(file_path.strip())

    if not path.exists():
        raise FileNotFoundError(
            f"File not found: '{path}'. "
            "Please check the path and try again."
        )

    if not path.is_file():
        raise ValueError(
            f"Path is not a file: '{path}'. "
            "Directories and symlinks are not accepted."
        )

    if path.suffix.lower() != allowed_extension.lower():
        raise ValueError(
            f"Invalid file type '{path.suffix}' for '{path.name}'. "
            f"Expected: '{allowed_extension}'."
        )

    file_size_mb = path.stat().st_size / (1024 * 1024)
    if file_size_mb > max_size_mb:
        raise ValueError(
            f"File '{path.name}' is too large: {file_size_mb:.2f} MB. "
            f"Maximum allowed: {max_size_mb} MB."
        )

    if path.stat().st_size == 0:
        raise ValueError(
            f"File '{path.name}' is empty (0 bytes). "
            "Please provide a valid non-empty file."
        )

    logger.debug(
        f"File validated: '{path.name}' "
        f"({file_size_mb:.3f} MB, extension='{path.suffix}')"
    )

    return path.resolve()


# ─────────────────────────────────────────────
# Function 4 — sanitize_text
# ─────────────────────────────────────────────

def sanitize_text(
    text: str,
    max_length: int = MAX_CELL_LENGTH,
    allow_newlines: bool = False,
) -> str:
    """
    Sanitizes a raw text string for safe use in spreadsheet cells
    and log entries.

    The following transformations are applied in order:
        1. Convert to string and strip surrounding whitespace
        2. Remove ASCII control characters (0x00–0x1F) except:
           - Tab (0x09) is always removed
           - Newline (0x0A) is kept only if allow_newlines=True
        3. Collapse multiple consecutive whitespace into single space
        4. Prevent Excel formula injection by prefixing dangerous
           leading characters (=, +, -, @, |, %) with a single quote
        5. Truncate to max_length characters

    This function is critical for security: without it, a malicious
    PDF could inject formulas into Excel cells (e.g. a description
    starting with "=CMD|' /C calc'!A0" would execute on open).

    Args:
        text:           The raw string to sanitize.
        max_length:     Maximum allowed character length after sanitization.
                        Defaults to MAX_CELL_LENGTH (100).
        allow_newlines: If True, preserves newline characters (0x0A).
                        Defaults to False (strip all control chars).

    Returns:
        A sanitized, safe string ready for spreadsheet or log insertion.
    """
    if text is None:
        return ""

    if not isinstance(text, str):
        text = str(text)

    # Strip surrounding whitespace
    result = text.strip()

    if not result:
        return ""

    # Remove control characters
    if allow_newlines:
        # Keep newlines (0x0A), remove everything else in 0x00–0x1F
        result = re.sub(r"[\x00-\x09\x0b-\x1f\x7f]", "", result)
    else:
        # Remove all control characters including newlines
        result = re.sub(r"[\x00-\x1f\x7f]", "", result)

    # Collapse multiple whitespace into a single space
    result = re.sub(r"\s+", " ", result).strip()

    # Prevent Excel formula injection
    # Leading dangerous characters get a single-quote prefix
    if result and result[0] in EXCEL_INJECTION_CHARS:
        result = "'" + result
        logger.warning(
            f"sanitize_text: Excel injection character detected and escaped. "
            f"Original start: {text[:10]!r}"
        )

    # Truncate to safe length
    if len(result) > max_length:
        logger.debug(
            f"sanitize_text: text truncated from {len(result)} "
            f"to {max_length} chars."
        )
        result = result[:max_length]

    return result


# ─────────────────────────────────────────────
# Function 5 — load_config
# ─────────────────────────────────────────────

def load_config(config_path: Optional[str] = None) -> dict[str, Any]:
    """
    Loads the application configuration from a JSON file.

    If the config file does not exist at the given path, the
    function falls back to DEFAULT_CONFIG and logs a warning.
    This ensures the application always has a valid configuration
    to run with, even on first run before config.json is created.

    The loaded config is merged with DEFAULT_CONFIG so that any
    missing keys in the user's file are filled with safe defaults.
    This makes the config file forward-compatible with new settings.

    Security rules enforced:
        - Config file must be a .json file
        - Config file must not exceed 1 MB (prevents DoS via huge config)
        - Only known keys from DEFAULT_CONFIG are returned
          (unknown keys are ignored to prevent injection via config)

    Args:
        config_path: Path to the JSON config file.
                     Defaults to DEFAULT_CONFIG_PATH ("config.json").

    Returns:
        A dictionary containing the merged application configuration.
        Always returns a complete config (never raises on missing file).
    """
    path = Path(config_path) if config_path else DEFAULT_CONFIG_PATH

    # Return defaults immediately if config file does not exist
    if not path.exists():
        logger.warning(
            f"Config file not found at '{path}'. "
            "Using default configuration. "
            "Create 'config.json' to customize settings."
        )
        return dict(DEFAULT_CONFIG)

    # Security: only accept .json files
    if path.suffix.lower() != ".json":
        logger.error(
            f"Config file must be a .json file. "
            f"Received: '{path.suffix}'. Using defaults."
        )
        return dict(DEFAULT_CONFIG)

    # Security: reject oversized config files
    max_config_size_bytes = 1024 * 1024  # 1 MB
    if path.stat().st_size > max_config_size_bytes:
        logger.error(
            f"Config file '{path.name}' exceeds 1 MB size limit. "
            "Using defaults."
        )
        return dict(DEFAULT_CONFIG)

    try:
        with open(path, "r", encoding="utf-8") as config_file:
            raw_config = json.load(config_file)
    except json.JSONDecodeError as exc:
        logger.error(
            f"Config file '{path.name}' contains invalid JSON: {exc}. "
            "Using defaults."
        )
        return dict(DEFAULT_CONFIG)
    except Exception as exc:
        logger.error(
            f"Failed to read config file '{path.name}': {exc}. "
            "Using defaults."
        )
        return dict(DEFAULT_CONFIG)

    if not isinstance(raw_config, dict):
        logger.error(
            f"Config file '{path.name}' must contain a JSON object. "
            "Using defaults."
        )
        return dict(DEFAULT_CONFIG)

    # Merge: start with defaults, override only known keys
    merged = dict(DEFAULT_CONFIG)
    unknown_keys = []

    for key, value in raw_config.items():
        if key in DEFAULT_CONFIG:
            merged[key] = value
        else:
            unknown_keys.append(key)

    if unknown_keys:
        logger.warning(
            f"Config file contains unknown keys (ignored): {unknown_keys}. "
            "Only known configuration keys are accepted."
        )

    logger.info(f"Configuration loaded from '{path.name}'.")
    return merged