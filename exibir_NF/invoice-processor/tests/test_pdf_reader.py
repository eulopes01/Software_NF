"""
tests/test_pdf_reader.py — Unit Tests for pdf_reader.py
---------------------------------------------------------
Tests all 5 public functions in pdf_reader.py:
    - read_pdf()
    - detect_bank()
    - extract_itau()
    - extract_banco_brasil()
    - normalize_transaction()

Run:
    python -m pytest tests/test_pdf_reader.py -v
"""

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent.parent / "src"))
sys.path.insert(0, str(Path(__file__).parent.parent))

from pdf_reader import (
    BANK_BANCO_BRASIL,
    BANK_ITAU,
    BANK_UNKNOWN,
    Transaction,
    detect_bank,
    extract_banco_brasil,
    extract_itau,
    normalize_transaction,
    read_pdf,
)


# ─────────────────────────────────────────────
# detect_bank
# ─────────────────────────────────────────────

class TestDetectBank:

    def test_detects_itau(self):
        text = "ITAUCARD MULTIPLO MASTERCARD INTERNATIONAL - final 3172"
        assert detect_bank(text) == BANK_ITAU

    def test_detects_itau_lowercase(self):
        text = "itaucard visa platinum"
        assert detect_bank(text) == BANK_ITAU

    def test_detects_banco_brasil(self):
        text = "Banco do Brasil - Fatura do Cartão"
        assert detect_bank(text) == BANK_BANCO_BRASIL

    def test_unknown_bank(self):
        text = "Some unknown bank statement"
        assert detect_bank(text) == BANK_UNKNOWN

    def test_empty_string_returns_unknown(self):
        assert detect_bank("") == BANK_UNKNOWN

    def test_none_returns_unknown(self):
        assert detect_bank(None) == BANK_UNKNOWN

    def test_itau_in_full_text_scan(self):
        # itaucard not in first 500 chars but present later
        text = ("x" * 600) + "itaucard"
        assert detect_bank(text) == BANK_ITAU

    def test_bb_in_full_text_scan(self):
        text = ("x" * 600) + "banco do brasil"
        assert detect_bank(text) == BANK_BANCO_BRASIL


# ─────────────────────────────────────────────
# extract_itau
# ─────────────────────────────────────────────

class TestExtractItau:

    SAMPLE_TEXT = """
ITAUCARD MULTIPLO MASTERCARD
Titular: JOHN DOE
Resumo da fatura em R$
Total desta fatura 1.337,61
Lançamentos nacionais
JOHN DOE (3172)
DATA MOVIMENTAÇÃO VALOR EM R$
03/10 MERCADOPAGO*MLIVRE 71,24
05/10 POSTO AUSTRAL 10,66
12/10 DROGARIAS PACHECO 01/03 43,34
20/10 CASAS BAHIACOM -1.598,21
Crédito do cartão final (3172) -1.598,28
Débito do cartão final (3172) 2.847,97
"""

    def test_extracts_transactions(self):
        transactions, _ = extract_itau(self.SAMPLE_TEXT)
        assert len(transactions) == 4

    def test_extracts_correct_total(self):
        _, total = extract_itau(self.SAMPLE_TEXT)
        assert total == 1337.61

    def test_transaction_fields(self):
        transactions, _ = extract_itau(self.SAMPLE_TEXT)
        first = transactions[0]
        assert first.date == "03/10"
        assert "MERCADOPAGO" in first.description
        assert first.amount == 71.24
        assert first.bank == BANK_ITAU

    def test_negative_value_extracted(self):
        transactions, _ = extract_itau(self.SAMPLE_TEXT)
        amounts = [t.amount for t in transactions]
        assert -1598.21 in amounts

    def test_empty_text_returns_empty(self):
        transactions, total = extract_itau("")
        assert transactions == []
        assert total == 0.0

    def test_skips_summary_lines(self):
        transactions, _ = extract_itau(self.SAMPLE_TEXT)
        descriptions = [t.description for t in transactions]
        assert not any("Crédito do cartão" in d for d in descriptions)
        assert not any("Débito do cartão" in d for d in descriptions)

    def test_bank_field_set_correctly(self):
        transactions, _ = extract_itau(self.SAMPLE_TEXT)
        assert all(t.bank == BANK_ITAU for t in transactions)


# ─────────────────────────────────────────────
# extract_banco_brasil
# ─────────────────────────────────────────────

class TestExtractBancoBrasil:

    SAMPLE_TEXT = """
Banco do Brasil - Ourocard
Total da fatura 850,00
lançamentos
10/05/2025 SUPERMERCADO ABC 150,00
15/05/2025 FARMACIA XYZ 45,50
20/05/2025 POSTO COMBUSTIVEL -20,00
total da fatura 850,00
"""

    def test_extracts_transactions(self):
        transactions, _ = extract_banco_brasil(self.SAMPLE_TEXT)
        assert len(transactions) == 3

    def test_extracts_total(self):
        _, total = extract_banco_brasil(self.SAMPLE_TEXT)
        assert total == 850.00

    def test_date_normalized_to_dd_mm(self):
        transactions, _ = extract_banco_brasil(self.SAMPLE_TEXT)
        for t in transactions:
            assert len(t.date) == 5
            assert t.date[2] == "/"

    def test_bank_field_set_correctly(self):
        transactions, _ = extract_banco_brasil(self.SAMPLE_TEXT)
        assert all(t.bank == BANK_BANCO_BRASIL for t in transactions)

    def test_empty_text_returns_empty(self):
        transactions, total = extract_banco_brasil("")
        assert transactions == []
        assert total == 0.0

    def test_negative_amount_extracted(self):
        transactions, _ = extract_banco_brasil(self.SAMPLE_TEXT)
        amounts = [t.amount for t in transactions]
        assert -20.00 in amounts


# ─────────────────────────────────────────────
# normalize_transaction
# ─────────────────────────────────────────────

class TestNormalizeTransaction:

    def test_normal_transaction_unchanged(self):
        t = Transaction(date="03/10", description="MERCADOPAGO", amount=71.24, bank=BANK_ITAU)
        result = normalize_transaction(t)
        assert result.date == "03/10"
        assert result.description == "MERCADOPAGO"
        assert result.amount == 71.24

    def test_removes_control_characters(self):
        t = Transaction(date="03/10", description="SHOP\x00\x1fSTORE", amount=10.0, bank=BANK_ITAU)
        result = normalize_transaction(t)
        assert "\x00" not in result.description
        assert "\x1f" not in result.description

    def test_prevents_excel_injection(self):
        t = Transaction(date="03/10", description="=MALICIOUS()", amount=10.0, bank=BANK_ITAU)
        result = normalize_transaction(t)
        assert result.description.startswith("'")

    def test_truncates_long_description(self):
        t = Transaction(date="03/10", description="A" * 200, amount=10.0, bank=BANK_ITAU)
        result = normalize_transaction(t)
        assert len(result.description) <= 100

    def test_invalid_date_set_to_default(self):
        t = Transaction(date="99-99-9999", description="SHOP", amount=10.0, bank=BANK_ITAU)
        result = normalize_transaction(t)
        assert result.date == "00/00"

    def test_invalid_amount_set_to_zero(self):
        t = Transaction(date="03/10", description="SHOP", amount=float("nan"), bank=BANK_ITAU)
        result = normalize_transaction(t)
        assert result.amount == 0.0

    def test_non_transaction_input_returns_invalid(self):
        result = normalize_transaction("not a transaction")
        assert result.description == "INVALID"
        assert result.amount == 0.0

    def test_bank_lowercased(self):
        t = Transaction(date="03/10", description="SHOP", amount=10.0, bank="ITAU")
        result = normalize_transaction(t)
        assert result.bank == "itau"

    def test_collapses_whitespace_in_description(self):
        t = Transaction(date="03/10", description="SHOP   NAME", amount=10.0, bank=BANK_ITAU)
        result = normalize_transaction(t)
        assert result.description == "SHOP NAME"


# ─────────────────────────────────────────────
# read_pdf — file validation
# ─────────────────────────────────────────────

class TestReadPdf:

    def test_file_not_found_raises(self):
        with pytest.raises(FileNotFoundError):
            read_pdf("/nonexistent/path/invoice.pdf")

    def test_wrong_extension_raises(self, tmp_path):
        txt = tmp_path / "invoice.txt"
        txt.write_text("not a pdf")
        with pytest.raises(ValueError, match="Invalid file type"):
            read_pdf(str(txt))

    def test_empty_file_raises(self, tmp_path):
        empty = tmp_path / "empty.pdf"
        empty.write_bytes(b"")
        with pytest.raises(ValueError):
            read_pdf(str(empty))

    def test_oversized_file_raises(self, tmp_path):
        big = tmp_path / "big.pdf"
        big.write_bytes(b"%PDF " + b"x" * (21 * 1024 * 1024))
        with pytest.raises(ValueError, match="too large"):
            read_pdf(str(big))