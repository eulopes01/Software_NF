import os
from datetime import datetime

def is_valid_pdf_path(file_path: str) -> bool:
    if not file_path or not isinstance(file_path, str):
        return False
    if not os.path.isfile(file_path):
        return False
    if not file_path.lower().endswith('.pdf'):
        return False
    return True

def parse_date(date_string: str) -> datetime:
    # Extrai apenas os números que o usuário digitou (ignora barras se ele digitar por costume)
    digits_only = ''.join(filter(str.isdigit, date_string))
    current_year = datetime.now().year

    if not digits_only:
        raise ValueError("Date cannot be empty.")

    # Se o usuário digitou apenas 4 números (ex: 1002)
    if len(digits_only) == 4:
        day, month = digits_only[:2], digits_only[2:4]
        full_date_str = f"{day}/{month}/{current_year}"
        
    # Se o usuário digitou os 8 números (ex: 10022026)
    elif len(digits_only) == 8:
        day, month, year = digits_only[:2], digits_only[2:4], digits_only[4:]
        full_date_str = f"{day}/{month}/{year}"
        
    else:
        raise ValueError("Invalid format. Just type DDMM (e.g., 1002).")

    try:
        # Verifica se é uma data real (ex: bloqueia 32/13/2026)
        return datetime.strptime(full_date_str, "%d/%m/%Y")
    except ValueError:
        raise ValueError("Invalid date. Check the day and month.")

def parse_value(value_string: str) -> float:
    try:
        clean_value = value_string.replace('R$', '').replace(',', '.').strip()
        return float(clean_value)
    except ValueError:
        raise ValueError("Invalid value. Please enter a valid number.")