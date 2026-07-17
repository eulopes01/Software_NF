from dataclasses import dataclass
from datetime import datetime

@dataclass
class ExpenseRecord:
    date: datetime
    vendor: str          # NOVO CAMPO: Fornecedor
    description: str
    value: float
    pdf_path: str