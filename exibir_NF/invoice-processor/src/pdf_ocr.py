"""
pdf_ocr.py — OCR for Scanned PDF Invoices
-------------------------------------------
Responsible for detecting scanned (image-based) PDF files,
converting each page to an image, running OCR using Tesseract,
and returning the extracted text in the same format expected
by pdf_reader.py.

This module is the fallback when pdfplumber cannot extract
text from a PDF (i.e. the PDF is a scanned image).

Pipeline integration:
    pdf_reader.py calls extract_text_from_scanned_pdf()
    when the normal text extraction returns empty content.

Each function has a single responsibility and at least 25 lines,
following Clean Architecture, SOLID, and Secure by Design principles.

Dependencies:
    pip install pytesseract pdf2image pillow
    sudo apt install tesseract-ocr tesseract-ocr-por tesseract-ocr-eng

Usage:
    from pdf_ocr import extract_text_from_scanned_pdf, is_scanned_pdf

    if is_scanned_pdf("pdfs/invoice.pdf"):
        text = extract_text_from_scanned_pdf("pdfs/invoice.pdf")
"""

import logging
import shutil
import tempfile
from pathlib import Path
from typing import Optional

# ─────────────────────────────────────────────
# Logging setup
# ─────────────────────────────────────────────

logger = logging.getLogger(__name__)

# ─────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────

ALLOWED_EXTENSION       = ".pdf"
MAX_FILE_SIZE_MB        = 20

# Tesseract language config — Portuguese + English
# '+' means use both language models simultaneously
TESSERACT_LANG          = "por+eng"

# Tesseract OCR configuration for invoice documents:
# --oem 3  → use LSTM neural net OCR engine (most accurate)
# --psm 6  → assume a uniform block of text (best for invoices)
TESSERACT_CONFIG        = "--oem 3 --psm 6"

# Image resolution for PDF-to-image conversion
# 300 DPI is the minimum recommended for accurate OCR
# Higher = better accuracy but slower processing
IMAGE_DPI               = 300

# Minimum text length to consider OCR successful
MIN_TEXT_LENGTH         = 50

# ─────────────────────────────────────────────
# Function 1 — is_scanned_pdf
# ─────────────────────────────────────────────

def is_scanned_pdf(file_path: str) -> bool:
    """
    Determines whether a PDF is scanned (image-based) by attempting
    to extract text using pdfplumber. If the extracted text is empty
    or below the minimum length threshold, the PDF is considered
    scanned and requires OCR processing.

    This function is used by pdf_reader.py to decide whether to use
    the normal text extraction pipeline or the OCR fallback.

    Args:
        file_path: Absolute or relative path to the PDF file.

    Returns:
        True if the PDF appears to be scanned (no extractable text).
        False if the PDF contains selectable text.

    Raises:
        FileNotFoundError: If the file does not exist.
        ValueError: If the file is not a valid PDF.
    """
    import pdfplumber

    path = Path(file_path)

    if not path.exists():
        raise FileNotFoundError(
            f"PDF file not found: '{path}'. "
            "Please check the path and try again."
        )

    if path.suffix.lower() != ALLOWED_EXTENSION:
        raise ValueError(
            f"Invalid file type '{path.suffix}'. "
            f"Only '{ALLOWED_EXTENSION}' files are supported."
        )

    file_size_mb = path.stat().st_size / (1024 * 1024)
    if file_size_mb > MAX_FILE_SIZE_MB:
        raise ValueError(
            f"File '{path.name}' is too large ({file_size_mb:.1f} MB). "
            f"Maximum allowed: {MAX_FILE_SIZE_MB} MB."
        )

    try:
        with pdfplumber.open(str(path)) as pdf:
            total_text = ""
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    total_text += page_text

        is_scanned = len(total_text.strip()) < MIN_TEXT_LENGTH

        logger.debug(
            f"is_scanned_pdf: '{path.name}' → "
            f"{'scanned' if is_scanned else 'text-based'} "
            f"(extracted {len(total_text.strip())} chars)"
        )
        return is_scanned

    except Exception as exc:
        logger.warning(
            f"Could not determine if '{path.name}' is scanned: {exc}. "
            "Assuming scanned (will attempt OCR)."
        )
        return True


# ─────────────────────────────────────────────
# Function 2 — convert_pdf_to_images
# ─────────────────────────────────────────────

def convert_pdf_to_images(
    file_path: str,
    dpi: int = IMAGE_DPI,
) -> list:
    """
    Converts each page of a PDF file into a PIL Image object
    suitable for OCR processing by Tesseract.

    Uses pdf2image (which wraps poppler) to render PDF pages
    at the specified DPI. Higher DPI produces clearer images
    and more accurate OCR results, at the cost of more memory
    and processing time.

    The recommended DPI for invoice OCR is 300. Values below
    200 DPI often produce unreliable results for financial data
    (amounts, dates) due to small font sizes.

    Args:
        file_path: Absolute or relative path to the PDF file.
        dpi:       Resolution for image rendering in dots per inch.
                   Defaults to IMAGE_DPI (300).

    Returns:
        A list of PIL Image objects, one per PDF page.

    Raises:
        FileNotFoundError: If the PDF file does not exist.
        ValueError: If conversion fails or produces no images.
    """
    from pdf2image import convert_from_path
    from pdf2image.exceptions import PDFInfoNotInstalledError, PDFPageCountError

    path = Path(file_path)

    if not path.exists():
        raise FileNotFoundError(f"PDF file not found: '{path}'")

    if dpi < 150:
        logger.warning(
            f"DPI value {dpi} is below the recommended minimum of 150. "
            "OCR accuracy may be poor. Consider using at least 300 DPI."
        )

    if dpi > 600:
        logger.warning(
            f"DPI value {dpi} is very high. "
            "This may cause slow processing and high memory usage."
        )

    try:
        images = convert_from_path(
            str(path),
            dpi=dpi,
            fmt="PNG",
        )

        if not images:
            raise ValueError(
                f"pdf2image returned no images for '{path.name}'. "
                "The PDF may be empty or corrupted."
            )

        logger.info(
            f"PDF converted to {len(images)} image(s) "
            f"at {dpi} DPI: '{path.name}'"
        )
        return images

    except PDFInfoNotInstalledError:
        raise ValueError(
            "poppler is not installed. "
            "Please install it with: sudo apt install poppler-utils"
        )
    except PDFPageCountError as exc:
        raise ValueError(
            f"Could not read page count from '{path.name}': {exc}. "
            "The PDF may be corrupted."
        ) from exc
    except Exception as exc:
        raise ValueError(
            f"Failed to convert '{path.name}' to images: {exc}"
        ) from exc


# ─────────────────────────────────────────────
# Function 3 — run_ocr_on_image
# ─────────────────────────────────────────────

def run_ocr_on_image(
    image,
    language: str = TESSERACT_LANG,
    config: str = TESSERACT_CONFIG,
) -> str:
    """
    Runs Tesseract OCR on a single PIL Image and returns
    the extracted text as a string.

    The language parameter supports multiple languages
    simultaneously using '+' as separator (e.g. "por+eng").
    This is important for Brazilian bank invoices that may
    contain English terms mixed with Portuguese text.

    The OCR config "--oem 3 --psm 6" is optimized for
    invoice documents with uniform text blocks. Different
    PSM values may work better for other document types:
        --psm 4  → single column of text
        --psm 11 → sparse text (receipts with wide spacing)

    Args:
        image:    A PIL Image object to run OCR on.
        language: Tesseract language code(s). Default: "por+eng".
        config:   Tesseract configuration string.

    Returns:
        Extracted text as a string. Returns empty string if
        OCR produces no output or fails.
    """
    try:
        import pytesseract
    except ImportError:
        raise ValueError(
            "pytesseract is not installed. "
            "Please install it with: pip install pytesseract"
        )

    if image is None:
        logger.warning("run_ocr_on_image received None image — returning empty string.")
        return ""

    try:
        text = pytesseract.image_to_string(
            image,
            lang=language,
            config=config,
        )

        if not text or not text.strip():
            logger.warning(
                "OCR returned empty text for this page. "
                "The image may be too dark, blurry, or low resolution."
            )
            return ""

        clean_text = text.strip()
        logger.debug(
            f"OCR extracted {len(clean_text)} characters "
            f"from image (lang='{language}')"
        )
        return clean_text

    except Exception as exc:
        logger.error(
            f"OCR failed on image: {exc}. "
            "Check that Tesseract is installed and the language packs are available."
        )
        return ""


# ─────────────────────────────────────────────
# Function 4 — extract_text_from_scanned_pdf
# ─────────────────────────────────────────────

def extract_text_from_scanned_pdf(
    file_path: str,
    dpi: int = IMAGE_DPI,
    language: str = TESSERACT_LANG,
) -> str:
    """
    Full OCR pipeline for a scanned PDF invoice.

    Converts each page to an image and runs Tesseract OCR,
    concatenating the text from all pages with newline separators.
    The result is in the same format as pdfplumber's extract_text(),
    making it a drop-in replacement in pdf_reader.py.

    If OCR produces insufficient text (below MIN_TEXT_LENGTH),
    a warning is logged and the partial result is still returned
    so the caller can decide how to handle it.

    Args:
        file_path: Absolute or relative path to the scanned PDF.
        dpi:       Resolution for image rendering. Default: 300.
        language:  Tesseract language code(s). Default: "por+eng".

    Returns:
        Full extracted text from all pages as a single string,
        with pages separated by newlines.

    Raises:
        FileNotFoundError: If the PDF does not exist.
        ValueError: If conversion or OCR fails completely.
    """
    path = Path(file_path)

    logger.info(
        f"Starting OCR pipeline for: '{path.name}' "
        f"(DPI={dpi}, lang='{language}')"
    )

    images = convert_pdf_to_images(file_path, dpi=dpi)

    if not images:
        raise ValueError(
            f"No images were generated from '{path.name}'. "
            "Cannot perform OCR on empty page list."
        )

    all_pages_text: list[str] = []

    for page_number, image in enumerate(images, start=1):
        logger.debug(f"Running OCR on page {page_number}/{len(images)}...")
        page_text = run_ocr_on_image(image, language=language)

        if page_text:
            all_pages_text.append(page_text)
            logger.debug(
                f"Page {page_number}: extracted {len(page_text)} characters."
            )
        else:
            logger.warning(f"Page {page_number}: OCR returned no text.")

    full_text = "\n".join(all_pages_text)

    if len(full_text.strip()) < MIN_TEXT_LENGTH:
        logger.warning(
            f"OCR result for '{path.name}' is very short "
            f"({len(full_text.strip())} chars). "
            "The scan quality may be too low for reliable extraction. "
            "Consider rescanning at higher resolution (300+ DPI)."
        )
    else:
        logger.info(
            f"OCR complete: '{path.name}' → "
            f"{len(all_pages_text)} page(s), "
            f"{len(full_text.strip())} characters extracted."
        )

    return full_text


# ─────────────────────────────────────────────
# Function 5 — check_tesseract_installation
# ─────────────────────────────────────────────

def check_tesseract_installation() -> dict:
    """
    Verifies that Tesseract OCR and the required language packs
    are correctly installed and accessible on the system.

    This function is called at application startup to provide
    a clear error message if Tesseract is missing, rather than
    failing silently during invoice processing.

    Checks performed:
        1. pytesseract Python package is importable
        2. Tesseract binary is found on the system PATH
        3. Tesseract version is readable
        4. Portuguese language pack (por) is installed
        5. English language pack (eng) is installed

    Returns:
        A dictionary with keys:
            "available":  bool — True if fully ready for OCR
            "version":    str  — Tesseract version string or error
            "languages":  list — installed language codes
            "has_por":    bool — Portuguese pack installed
            "has_eng":    bool — English pack installed
            "error":      str  — error message if not available

    Never raises — returns availability=False with error details instead.
    """
    result = {
        "available": False,
        "version":   "unknown",
        "languages": [],
        "has_por":   False,
        "has_eng":   False,
        "error":     "",
    }

    try:
        import pytesseract
    except ImportError:
        result["error"] = (
            "pytesseract is not installed. "
            "Run: pip install pytesseract"
        )
        logger.error(result["error"])
        return result

    try:
        version = pytesseract.get_tesseract_version()
        result["version"] = str(version)
    except Exception as exc:
        result["error"] = (
            f"Tesseract binary not found: {exc}. "
            "Run: sudo apt install tesseract-ocr"
        )
        logger.error(result["error"])
        return result

    try:
        languages = pytesseract.get_languages(config="")
        result["languages"] = languages
        result["has_por"] = "por" in languages
        result["has_eng"] = "eng" in languages
    except Exception as exc:
        logger.warning(f"Could not retrieve Tesseract language list: {exc}")

    missing_langs = []
    if not result["has_por"]:
        missing_langs.append("tesseract-ocr-por")
    if not result["has_eng"]:
        missing_langs.append("tesseract-ocr-eng")

    if missing_langs:
        result["error"] = (
            f"Missing language packs: {', '.join(missing_langs)}. "
            f"Run: sudo apt install {' '.join(missing_langs)}"
        )
        logger.warning(result["error"])
        result["available"] = False
        return result

    result["available"] = True
    logger.info(
        f"Tesseract ready: version={result['version']} "
        f"| languages={result['languages']}"
    )
    return result