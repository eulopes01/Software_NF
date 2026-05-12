"""
tests/test_utils.py — Unit Tests for utils.py
-----------------------------------------------
Tests all 5 public functions in utils.py:
    - format_currency()
    - format_date()
    - validate_file()
    - sanitize_text()
    - load_config()

Run:
    python -m pytest tests/test_utils.py -v
"""

import json
import sys
import tempfile
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent.parent / "src"))
sys.path.insert(0, str(Path(__file__).parent.parent))

from utils import (
    format_currency,
    format_date,
    load_config,
    sanitize_text,
    validate_file,
)


# ─────────────────────────────────────────────
# format_currency
# ─────────────────────────────────────────────

class TestFormatCurrency:

    def test_positive_value(self):
        assert format_currency(1337.61) == "R$ 1.337,61"

    def test_zero_value(self):
        assert format_currency(0.0) == "R$ 0,00"

    def test_negative_value(self):
        assert format_currency(-50.00) == "R$ -50,00"

    def test_large_value(self):
        assert format_currency(1000000.99) == "R$ 1.000.000,99"

    def test_custom_symbol(self):
        result = format_currency(100.00, symbol="USD")
        assert result == "USD 100,00"

    def test_integer_input(self):
        assert format_currency(100) == "R$ 100,00"

    def test_invalid_type_raises(self):
        with pytest.raises(TypeError):
            format_currency("not a number")

    def test_invalid_none_raises(self):
        with pytest.raises(TypeError):
            format_currency(None)

    def test_small_decimal(self):
        assert format_currency(0.01) == "R$ 0,01"

    def test_thousand_separator(self):
        result = format_currency(1000.00)
        assert "1.000" in result


# ─────────────────────────────────────────────
# format_date
# ─────────────────────────────────────────────

class TestFormatDate:

    def test_dd_mm_to_dd_mm_yyyy(self):
        result = format_date("03/10", reference_year=2025)
        assert result == "03/10/2025"

    def test_uses_current_year_when_no_reference(self):
        from datetime import datetime
        result = format_date("03/10")
        year = str(datetime.now().year)
        assert year in result

    def test_full_date_conversion(self):
        result = format_date("2025-10-03", "%Y-%m-%d", "%d/%m/%Y")
        assert result == "03/10/2025"

    def test_invalid_date_returns_original(self):
        result = format_date("99/99")
        assert result == "99/99"

    def test_empty_string_returns_empty(self):
        result = format_date("")
        assert result == ""

    def test_none_returns_empty(self):
        result = format_date(None)
        assert result == ""

    def test_reference_year_injected(self):
        result = format_date("15/06", reference_year=2024)
        assert "2024" in result

    def test_already_full_date_passthrough(self):
        result = format_date("15/06/2025", "%d/%m/%Y", "%d/%m/%Y")
        assert result == "15/06/2025"


# ─────────────────────────────────────────────
# validate_file
# ─────────────────────────────────────────────

class TestValidateFile:

    def test_valid_pdf_file(self, tmp_path):
        pdf = tmp_path / "test.pdf"
        pdf.write_bytes(b"%PDF-1.4 fake content")
        result = validate_file(str(pdf), ".pdf")
        assert result == pdf.resolve()

    def test_valid_xlsx_file(self, tmp_path):
        xlsx = tmp_path / "test.xlsx"
        xlsx.write_bytes(b"PK fake xlsx content")
        result = validate_file(str(xlsx), ".xlsx")
        assert result == xlsx.resolve()

    def test_file_not_found_raises(self):
        with pytest.raises(FileNotFoundError):
            validate_file("/nonexistent/path/file.pdf", ".pdf")

    def test_wrong_extension_raises(self, tmp_path):
        txt = tmp_path / "test.txt"
        txt.write_text("content")
        with pytest.raises(ValueError, match="Invalid file type"):
            validate_file(str(txt), ".pdf")

    def test_empty_file_raises(self, tmp_path):
        empty = tmp_path / "empty.pdf"
        empty.write_bytes(b"")
        with pytest.raises(ValueError, match="empty"):
            validate_file(str(empty), ".pdf")

    def test_empty_path_raises(self):
        with pytest.raises(ValueError):
            validate_file("", ".pdf")

    def test_none_path_raises(self):
        with pytest.raises(ValueError):
            validate_file(None, ".pdf")

    def test_oversized_file_raises(self, tmp_path):
        big = tmp_path / "big.pdf"
        big.write_bytes(b"x" * (21 * 1024 * 1024))  # 21 MB
        with pytest.raises(ValueError, match="too large"):
            validate_file(str(big), ".pdf", max_size_mb=20)


# ─────────────────────────────────────────────
# sanitize_text
# ─────────────────────────────────────────────

class TestSanitizeText:

    def test_strips_whitespace(self):
        assert sanitize_text("  hello  ") == "hello"

    def test_removes_control_characters(self):
        result = sanitize_text("hello\x00world\x1f!")
        assert "\x00" not in result
        assert "\x1f" not in result

    def test_prevents_excel_injection_equals(self):
        result = sanitize_text("=SUM(A1:A10)")
        assert result.startswith("'")

    def test_prevents_excel_injection_plus(self):
        result = sanitize_text("+cmd|calc")
        assert result.startswith("'")

    def test_prevents_excel_injection_at(self):
        result = sanitize_text("@malicious")
        assert result.startswith("'")

    def test_normal_text_unchanged(self):
        result = sanitize_text("MERCADOPAGO*MLIVRE")
        assert result == "MERCADOPAGO*MLIVRE"

    def test_truncates_to_max_length(self):
        long_text = "A" * 200
        result = sanitize_text(long_text, max_length=100)
        assert len(result) <= 100

    def test_none_returns_empty(self):
        assert sanitize_text(None) == ""

    def test_collapses_multiple_spaces(self):
        result = sanitize_text("hello   world")
        assert result == "hello world"

    def test_non_string_converted(self):
        result = sanitize_text(12345)
        assert result == "12345"


# ─────────────────────────────────────────────
# load_config
# ─────────────────────────────────────────────

class TestLoadConfig:

    def test_returns_defaults_when_no_file(self, tmp_path):
        config = load_config(str(tmp_path / "nonexistent.json"))
        assert config["payment_method"] == "A VISTA"
        assert config["backup_enabled"] is True
        assert config["max_transactions"] == 500

    def test_loads_valid_config(self, tmp_path):
        config_file = tmp_path / "config.json"
        config_file.write_text(json.dumps({"log_level": "DEBUG"}))
        config = load_config(str(config_file))
        assert config["log_level"] == "DEBUG"

    def test_merges_missing_keys_with_defaults(self, tmp_path):
        config_file = tmp_path / "config.json"
        config_file.write_text(json.dumps({"log_level": "WARNING"}))
        config = load_config(str(config_file))
        assert "payment_method" in config
        assert "backup_enabled" in config

    def test_ignores_unknown_keys(self, tmp_path):
        config_file = tmp_path / "config.json"
        config_file.write_text(json.dumps({"unknown_key": "value"}))
        config = load_config(str(config_file))
        assert "unknown_key" not in config

    def test_rejects_invalid_json(self, tmp_path):
        config_file = tmp_path / "config.json"
        config_file.write_text("{invalid json}")
        config = load_config(str(config_file))
        assert config["log_level"] == "INFO"

    def test_rejects_non_json_extension(self, tmp_path):
        txt_file = tmp_path / "config.txt"
        txt_file.write_text("{}")
        config = load_config(str(txt_file))
        assert config["log_level"] == "INFO"

    def test_rejects_oversized_config(self, tmp_path):
        big_config = tmp_path / "config.json"
        big_config.write_bytes(b"x" * (1024 * 1024 + 1))
        config = load_config(str(big_config))
        assert config["log_level"] == "INFO"