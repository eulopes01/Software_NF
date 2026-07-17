from typing import List
from PyPDF2 import PdfMerger
import os
from domain.expense import ExpenseRecord
from utils.validators import is_valid_pdf_path

class SecurePdfMergerService:
    def merge_chronologically(self, records: List[ExpenseRecord], output_path: str) -> bool:
        if not records:
            return False

        # NOVO: Filtra para manter apenas as despesas que possuem um PDF anexado
        records_with_pdf = [r for r in records if r.pdf_path]
        
        if not records_with_pdf:
            return False # Retorna falso se não houver nenhum PDF para juntar

        sorted_records = sorted(records_with_pdf, key=lambda r: r.date)
        merger = PdfMerger()

        try:
            self._append_pdfs(merger, sorted_records)
            merger.write(output_path)
            return True
        except Exception as e:
            raise RuntimeError("An error occurred while merging the PDF files.") from e
        finally:
            merger.close()

    def _append_pdfs(self, merger: PdfMerger, records: List[ExpenseRecord]) -> None:
        for record in records:
            if not is_valid_pdf_path(record.pdf_path):
                raise FileNotFoundError(f"Invalid or missing PDF: {record.pdf_path}")
            merger.append(record.pdf_path)