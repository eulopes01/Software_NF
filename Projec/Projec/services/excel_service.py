import openpyxl
from typing import List
from domain.expense import ExpenseRecord

class ExcelAutomatorService:
    def update_spreadsheet(self, records: List[ExpenseRecord], excel_path: str) -> bool:
        try:
            # Carrega a planilha sem alterar a formatação existente
            wb = openpyxl.load_workbook(excel_path)
            
            # Trabalha na aba (planilha) que estiver ativa/aberta por padrão
            sheet = wb.active 

            # ==================================================
            # PASSO 1: A Borracha (Limpeza de Dados Antigos)
            # ==================================================
            # Pega o número da última linha que tem algum dado na planilha
            max_row = sheet.max_row
            
            # Passa apagando os dados antigos a partir da linha 4, 
            # mexendo APENAS nas colunas que nós usamos.
            if max_row >= 4:
                for row in range(4, max_row + 1):
                    sheet[f"A{row}"] = None
                    sheet[f"C{row}"] = None
                    sheet[f"D{row}"] = None
                    sheet[f"E{row}"] = None

            # ==================================================
            # PASSO 2: Inserção da Nova Lista
            # ==================================================
            # Sempre recomeça a preencher da linha 4 limpa
            current_row = 4
            for record in records:
                sheet[f"A{current_row}"] = record.date.strftime("%d/%m/%Y")
                sheet[f"C{current_row}"] = record.vendor
                sheet[f"D{current_row}"] = record.description
                sheet[f"E{current_row}"] = record.value 
                current_row += 1

            # Salva o arquivo sobrescrevendo de forma segura
            wb.save(excel_path)
            return True
            
        except Exception as e:
            print(f"Erro crítico no motor de Excel: {e}")
            return False