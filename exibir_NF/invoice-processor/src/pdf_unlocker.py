"""
pdf_unlocker.py — PDF Password Removal
----------------------------------------
Responsible for detecting password-protected PDF files,
validating the provided password, and producing an unlocked
copy ready for text extraction by pdf_reader.py.

The original password-protected file is never modified.
A temporary unlocked copy is created, used for processing,
and deleted automatically after the pipeline finishes.

Each function has a single responsibility and at least 25 lines,
following Clean Architecture, SOLID, and Secure by Design principles.

Dependencies:
    pip install pypdf

Usage:
    from pdf_unlocker import unlock_pdf, is_pdf_protected

    if is_pdf_protected("pdfs/invoice.pdf"):
        unlocked_path = unlock_pdf("pdfs/invoice.pdf", password="12345")
        # process unlocked_path ...
        # cleanup is handled automatically via context manager
"""

import logging
import shutil
from contextlib import contextmanager
from pathlib import Path
from typing import Generator, Optional

from pypdf import PdfReader, PdfWriter

# ─────────────────────────────────────────────
# Logging setup
# ─────────────────────────────────────────────

logger = logging.getLogger(__name__)

# ─────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────

ALLOWED_EXTENSION       = ".pdf"
MAX_FILE_SIZE_MB        = 20
UNLOCKED_SUFFIX         = "_unlocked"
MAX_PASSWORD_LENGTH     = 128

# ─────────────────────────────────────────────
# Function 1 — is_pdf_protected
# ─────────────────────────────────────────────

def is_pdf_protected(file_path: str) -> bool:
    """
    Checks whether a PDF file is password-protected.

    Opens the PDF using pypdf and checks the is_encrypted
    property. If the file is encrypted but can be opened
    with an empty password (some PDFs have owner-only
    restrictions), it is still considered protected since
    text extraction may fail without proper decryption.

    Args:
        file_path: Absolute or relative path to the PDF file.

    Returns:
        True if the PDF is password-protected, False otherwise.

    Raises:
        FileNotFoundError: If the file does not exist.
        ValueError: If the file is not a valid PDF or is too large.
    """
    path = Path(file_path)

    if not path.exists():
        raise FileNotFoundError(
            f"PDF file not found: '{path}'. "
            "Please check the path and try again."
        )

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
            f"Maximum allowed: {MAX_FILE_SIZE_MB} MB."
        )

    try:
        reader = PdfReader(str(path))
        is_protected = reader.is_encrypted
        logger.debug(
            f"Protection check: '{path.name}' → "
            f"{'protected' if is_protected else 'not protected'}"
        )
        return is_protected
    except Exception as exc:
        logger.warning(
            f"Could not check encryption status of '{path.name}': {exc}. "
            "Assuming not protected."
        )
        return False


# ─────────────────────────────────────────────
# Function 2 — validate_password
# ─────────────────────────────────────────────

def validate_password(file_path: str, password: str) -> bool:
    """
    Attempts to open a password-protected PDF with the given
    password to verify it is correct before processing.

    This validation step prevents partial processing failures
    where the pipeline starts but fails mid-way due to a wrong
    password. It is always called before unlock_pdf().

    Security note: the password is never logged, stored, or
    included in any error message. Only a boolean result is
    returned to avoid credential exposure.

    Args:
        file_path: Path to the password-protected PDF file.
        password:  The password string to test.

    Returns:
        True if the password is correct, False otherwise.

    Raises:
        FileNotFoundError: If the file does not exist.
        ValueError: If inputs are invalid.
    """
    path = Path(file_path)

    if not path.exists():
        raise FileNotFoundError(f"PDF file not found: '{path}'")

    if not password or not isinstance(password, str):
        logger.warning("validate_password received empty or invalid password.")
        return False

    if len(password) > MAX_PASSWORD_LENGTH:
        logger.warning(
            f"Password exceeds maximum length of {MAX_PASSWORD_LENGTH} characters. "
            "Rejecting without attempting decryption."
        )
        return False

    try:
        reader = PdfReader(str(path))

        if not reader.is_encrypted:
            logger.debug(f"'{path.name}' is not encrypted — password not required.")
            return True

        # decrypt() returns 0 (wrong), 1 (user password), or 2 (owner password)
        result = reader.decrypt(password)
        is_valid = result > 0

        if is_valid:
            logger.debug(f"Password validated successfully for '{path.name}'.")
        else:
            logger.warning(f"Invalid password provided for '{path.name}'.")

        return is_valid

    except Exception as exc:
        logger.error(
            f"Error validating password for '{path.name}': {exc}"
        )
        return False


# ─────────────────────────────────────────────
# Function 3 — unlock_pdf
# ─────────────────────────────────────────────

def unlock_pdf(file_path: str, password: str) -> Path:
    """
    Decrypts a password-protected PDF and saves an unlocked
    copy to the same directory as the original file.

    The unlocked copy is named:
        original_name_unlocked.pdf

    The original file is never modified. The unlocked copy
    should be deleted after processing using cleanup_unlocked_pdf()
    or the unlock_pdf_context() context manager.

    Args:
        file_path: Path to the password-protected PDF file.
        password:  The correct password for the PDF.

    Returns:
        Path to the unlocked PDF copy.

    Raises:
        FileNotFoundError: If the original file does not exist.
        ValueError: If the password is wrong or decryption fails.
    """
    path = Path(file_path)

    if not path.exists():
        raise FileNotFoundError(f"PDF file not found: '{path}'")

    if not password or not isinstance(password, str):
        raise ValueError(
            "Password must be a non-empty string. "
            "Cannot unlock PDF without a valid password."
        )

    # Validate password before attempting to write
    if not validate_password(file_path, password):
        raise ValueError(
            f"Incorrect password for '{path.name}'. "
            "Please check the password and try again."
        )

    unlocked_path = path.parent / f"{path.stem}{UNLOCKED_SUFFIX}{path.suffix}"

    try:
        reader = PdfReader(str(path))
        reader.decrypt(password)

        writer = PdfWriter()

        for page_number, page in enumerate(reader.pages):
            writer.add_page(page)
            logger.debug(f"Copied page {page_number + 1} from '{path.name}'")

        with open(unlocked_path, "wb") as output_file:
            writer.write(output_file)

        logger.info(
            f"PDF unlocked successfully: '{path.name}' → '{unlocked_path.name}'"
        )
        return unlocked_path

    except ValueError:
        # Re-raise ValueError (wrong password) without wrapping
        raise
    except Exception as exc:
        # Clean up partial output if it was created
        if unlocked_path.exists():
            unlocked_path.unlink()
        raise ValueError(
            f"Failed to unlock '{path.name}': {exc}"
        ) from exc


# ─────────────────────────────────────────────
# Function 4 — cleanup_unlocked_pdf
# ─────────────────────────────────────────────

def cleanup_unlocked_pdf(unlocked_path: Path) -> None:
    """
    Deletes the temporary unlocked PDF copy after processing
    is complete.

    This function must always be called after unlock_pdf()
    to ensure temporary files do not accumulate on disk.
    It is safe to call even if the file no longer exists
    (e.g. if cleanup was already performed).

    Security note: temporary unlocked PDFs should never
    remain on disk longer than necessary, as they contain
    sensitive financial data without password protection.

    Args:
        unlocked_path: Path to the unlocked PDF file to delete.

    Returns:
        None. Logs a warning if deletion fails but does not raise.
    """
    if unlocked_path is None:
        logger.debug("cleanup_unlocked_pdf called with None — nothing to clean up.")
        return

    if not isinstance(unlocked_path, Path):
        unlocked_path = Path(unlocked_path)

    if not unlocked_path.exists():
        logger.debug(
            f"cleanup_unlocked_pdf: '{unlocked_path.name}' "
            "does not exist — already cleaned up."
        )
        return

    # Safety check: only delete files with the unlocked suffix
    # to prevent accidental deletion of original files
    if UNLOCKED_SUFFIX not in unlocked_path.stem:
        logger.error(
            f"cleanup_unlocked_pdf refused to delete '{unlocked_path.name}': "
            f"file does not contain the expected suffix '{UNLOCKED_SUFFIX}'. "
            "Only temporary unlocked copies should be deleted."
        )
        return

    try:
        unlocked_path.unlink()
        logger.info(f"Temporary unlocked PDF deleted: '{unlocked_path.name}'")
    except Exception as exc:
        logger.warning(
            f"Failed to delete temporary unlocked PDF '{unlocked_path.name}': {exc}. "
            "Please delete it manually to avoid sensitive data exposure."
        )


# ─────────────────────────────────────────────
# Function 5 — unlock_pdf_context
# ─────────────────────────────────────────────

@contextmanager
def unlock_pdf_context(
    file_path: str,
    password: str,
) -> Generator[Path, None, None]:
    """
    Context manager that unlocks a password-protected PDF,
    yields the path to the unlocked copy, and automatically
    deletes the temporary file when the block exits — even
    if an exception occurs inside the block.

    This is the recommended way to use the unlock feature,
    as it guarantees cleanup without requiring the caller
    to explicitly call cleanup_unlocked_pdf().

    Usage:
        with unlock_pdf_context("pdfs/invoice.pdf", password="12345") as unlocked:
            transactions, total = read_pdf(str(unlocked))

    Args:
        file_path: Path to the password-protected PDF file.
        password:  The correct password for the PDF.

    Yields:
        Path to the temporary unlocked PDF file.

    Raises:
        FileNotFoundError: If the original file does not exist.
        ValueError: If the password is wrong or decryption fails.
    """
    unlocked_path: Optional[Path] = None

    if not file_path or not isinstance(file_path, str):
        raise ValueError(
            f"file_path must be a non-empty string. "
            f"Received: {file_path!r}"
        )

    if not password or not isinstance(password, str):
        raise ValueError(
            "password must be a non-empty string. "
            "Cannot unlock PDF without a valid password."
        )

    logger.debug(
        f"unlock_pdf_context: entering context for '{Path(file_path).name}'"
    )

    try:
        unlocked_path = unlock_pdf(file_path, password)
        yield unlocked_path
    finally:
        if unlocked_path is not None:
            cleanup_unlocked_pdf(unlocked_path)
            logger.debug(
                f"unlock_pdf_context: cleanup complete for '{Path(file_path).name}'"
            )