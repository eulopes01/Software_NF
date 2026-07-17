import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import os
from typing import List
from domain.expense import ExpenseRecord
from services.pdf_service import SecurePdfMergerService
from utils.validators import parse_date, parse_value, is_valid_pdf_path

class ExpenseApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Expense PDF Merger")
        self.records: List[ExpenseRecord] = []
        self.pdf_service = SecurePdfMergerService()
        
        self.vendors_file = "vendors.json"
        self.saved_vendors = self._load_vendors()
        
        self._build_ui()

    def _load_vendors(self) -> List[str]:
        if os.path.exists(self.vendors_file):
            try:
                with open(self.vendors_file, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                return []
        return []

    def _save_vendor(self, vendor: str) -> None:
        if vendor and vendor not in self.saved_vendors:
            self.saved_vendors.append(vendor)
            self.saved_vendors.sort()
            try:
                with open(self.vendors_file, "w", encoding="utf-8") as f:
                    json.dump(self.saved_vendors, f)
                self.vendor_combobox['values'] = self.saved_vendors
            except Exception:
                pass 

    def _format_currency(self, value: float) -> str:
        standard_format = f"{value:,.2f}"
        translation_table = str.maketrans(',.', '.,')
        return standard_format.translate(translation_table)

    def _build_ui(self) -> None:
        self._build_input_frame()
        self._build_treeview()
        self._build_footer_frame()

    def _build_input_frame(self) -> None:
        frame = tk.Frame(self.root)
        frame.pack(pady=10, padx=10, fill=tk.X)

        tk.Label(frame, text="Date:").grid(row=0, column=0)
        self.date_entry = tk.Entry(frame, width=8)
        self.date_entry.grid(row=0, column=1, padx=5)

        tk.Label(frame, text="Vendor:").grid(row=0, column=2)
        self.vendor_combobox = ttk.Combobox(frame, values=self.saved_vendors, width=15)
        self.vendor_combobox.grid(row=0, column=3, padx=5)

        tk.Label(frame, text="Description:").grid(row=0, column=4)
        self.desc_entry = tk.Entry(frame, width=20)
        self.desc_entry.grid(row=0, column=5, padx=5)

        tk.Label(frame, text="Value (R$):").grid(row=0, column=6)
        self.val_entry = tk.Entry(frame, width=8)
        self.val_entry.grid(row=0, column=7, padx=5)

        self.pdf_path_var = tk.StringVar()
        tk.Button(frame, text="Attach PDF", command=self._select_pdf).grid(row=0, column=8, padx=5)
        
        tk.Button(frame, text="Add Expense", command=self._add_expense).grid(row=0, column=9, padx=5)

    def _build_treeview(self) -> None:
        columns = ("Date", "Vendor", "Description", "Value", "PDF Path")
        self.tree = ttk.Treeview(self.root, columns=columns, show="headings")
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120)
        self.tree.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

    def _build_footer_frame(self) -> None:
        footer_frame = tk.Frame(self.root)
        footer_frame.pack(pady=10, padx=10, fill=tk.X)

        self.total_var = tk.StringVar()
        self.total_var.set("Total: R$ 0,00")
        total_label = tk.Label(footer_frame, textvariable=self.total_var, font=("Arial", 12, "bold"))
        total_label.pack(side=tk.LEFT)

        # NOVO: Botões de Edição e Remoção na interface
        tk.Button(footer_frame, text="Edit Selected", command=self._edit_expense).pack(side=tk.LEFT, padx=15)
        tk.Button(footer_frame, text="Delete Selected", fg="red", command=self._delete_expense).pack(side=tk.LEFT)

        tk.Button(footer_frame, text="Merge PDFs", bg="green", fg="white", 
                  command=self._merge_pdfs).pack(side=tk.RIGHT)

    def _select_pdf(self) -> None:
        file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if file_path:
            self.pdf_path_var.set(file_path)
            messagebox.showinfo("Success", "PDF attached temporarily.")

    def _add_expense(self) -> None:
        try:
            date_obj = parse_date(self.date_entry.get())
            vendor = self.vendor_combobox.get().strip()  
            desc = self.desc_entry.get().strip()
            value_float = parse_value(self.val_entry.get())
            pdf_path = self.pdf_path_var.get()

            if not desc or not vendor:
                raise ValueError("Vendor and Description are required.")
                
            if pdf_path and not is_valid_pdf_path(pdf_path):
                raise ValueError("The attached file is not a valid PDF.")

            self._save_vendor(vendor)

            record = ExpenseRecord(date_obj, vendor, desc, value_float, pdf_path)
            self.records.append(record)
            
            # Centralizamos a atualização da tela na função _refresh_ui (DRY)
            self._refresh_ui()
            self._clear_inputs()
        except ValueError as e:
            messagebox.showerror("Input Error", str(e))

    def _edit_expense(self) -> None:
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select an expense to edit.")
            return

        # Pega o ID exato da linha e remove a nota da memória
        idx = int(selected[0])
        record = self.records.pop(idx)

        # Devolve os dados para os campos de preenchimento lá em cima
        self._clear_inputs()
        self.date_entry.insert(0, record.date.strftime("%d/%m/%Y"))
        self.vendor_combobox.set(record.vendor)
        self.desc_entry.insert(0, record.description)
        self.val_entry.insert(0, str(record.value))
        
        if record.pdf_path:
            self.pdf_path_var.set(record.pdf_path)

        # Atualiza a tela (a nota sumirá da tabela até você clicar em Add de novo)
        self._refresh_ui()

    def _delete_expense(self) -> None:
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select an expense to delete.")
            return

        if messagebox.askyesno("Confirm", "Are you sure you want to delete this expense?"):
            idx = int(selected[0])
            self.records.pop(idx)
            self._refresh_ui()

    def _refresh_ui(self) -> None:
        # 1. Ordena a lista
        self.records.sort(key=lambda r: r.date)
        
        # 2. Limpa a tabela
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        # 3. Reconstrói a tabela ancorando o iid (identificador) ao índice da memória
        for idx, rec in enumerate(self.records):
            date_str = rec.date.strftime("%d/%m/%Y")
            formatted_value = self._format_currency(rec.value)
            display_pdf = rec.pdf_path if rec.pdf_path else "No PDF"
            self.tree.insert("", tk.END, iid=str(idx), values=(date_str, rec.vendor, rec.description, f"R$ {formatted_value}", display_pdf))
            
        # 4. Atualiza o Total
        total = sum(record.value for record in self.records)
        formatted_total = self._format_currency(total)
        self.total_var.set(f"Total: R$ {formatted_total}")

    def _clear_inputs(self) -> None:
        self.date_entry.delete(0, tk.END)
        self.vendor_combobox.set("") 
        self.desc_entry.delete(0, tk.END)
        self.val_entry.delete(0, tk.END)
        self.pdf_path_var.set("")

    def _merge_pdfs(self) -> None:
        if not self.records:
            messagebox.showwarning("Warning", "No expenses registered.")
            return

        has_pdfs = any(r.pdf_path for r in self.records)
        if not has_pdfs:
            messagebox.showwarning("Warning", "None of the registered expenses have a PDF attached to merge.")
            return

        output_path = filedialog.asksaveasfilename(defaultextension=".pdf", 
                                                   filetypes=[("PDF Files", "*.pdf")])
        if not output_path:
            return

        try:
            success = self.pdf_service.merge_chronologically(self.records, output_path)
            if success:
                messagebox.showinfo("Success", "PDFs merged successfully in chronological order!")
            else:
                messagebox.showwarning("Warning", "Failed to merge. No valid PDFs found.")
        except Exception:
            messagebox.showerror("Error", "Failed to merge PDFs securely. Check if files exist and are valid.")