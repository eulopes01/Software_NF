import customtkinter as ctk
from tkinter import ttk, filedialog, messagebox
import json
import os
from datetime import datetime
from typing import List
from domain.expense import ExpenseRecord
from services.pdf_service import SecurePdfMergerService
from utils.validators import parse_date, parse_value, is_valid_pdf_path

class ExpenseApp:
    def __init__(self, root: ctk.CTk):
        self.root = root
        self.root.title("Expense PDF Merger")
        self.pdf_service = SecurePdfMergerService()
        
        # Define o diretório de dados na pasta AppData (oculta e segura)
        # O os.getenv('APPDATA') aponta para C:\Users\SeuUsuario\AppData\Roaming
        appdata_path = os.getenv('APPDATA')
        self.data_dir = os.path.join(appdata_path, 'ExpenseApp')
        
        if not os.path.exists(self.data_dir):
            os.makedirs(self.data_dir)
            
        self.vendors_file = os.path.join(self.data_dir, "vendors.json")
        self.expenses_file = os.path.join(self.data_dir, "expenses.json")
        
        self.saved_vendors = self._load_vendors()
        self.records: List[ExpenseRecord] = self._load_expenses()
        
        self._style_treeview()
        self._build_ui()
        
        if self.records:
            self._refresh_ui()
            
    def _style_treeview(self) -> None:
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview",
                        background="#2b2b2b",
                        foreground="white",
                        rowheight=28,
                        fieldbackground="#2b2b2b",
                        bordercolor="#343638",
                        borderwidth=0)
        style.map('Treeview', background=[('selected', '#1f538d')])
        style.configure("Treeview.Heading",
                        background="#565b5e",
                        foreground="white",
                        relief="flat",
                        font=("Arial", 10, "bold"))
        style.map("Treeview.Heading", background=[('active', '#343638')])

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
                self.vendor_combobox.configure(values=self.saved_vendors)
            except Exception:
                pass 

    def _load_expenses(self) -> List[ExpenseRecord]:
        if os.path.exists(self.expenses_file):
            try:
                with open(self.expenses_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    loaded_records = []
                    for item in data:
                        date_obj = datetime.strptime(item["date"], "%d/%m/%Y")
                        loaded_records.append(ExpenseRecord(
                            date=date_obj,
                            vendor=item.get("vendor", ""),
                            description=item.get("description", ""),
                            value=item.get("value", 0.0),
                            pdf_path=item.get("pdf_path", "")
                        ))
                    return loaded_records
            except Exception:
                return [] 
        return []

    def _save_expenses(self) -> None:
        try:
            data = []
            for r in self.records:
                data.append({
                    "date": r.date.strftime("%d/%m/%Y"),
                    "vendor": r.vendor,
                    "description": r.description,
                    "value": r.value,
                    "pdf_path": r.pdf_path
                })
            with open(self.expenses_file, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=4) 
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
        frame = ctk.CTkFrame(self.root, corner_radius=10)
        frame.pack(pady=15, padx=15, fill="x")

        ctk.CTkLabel(frame, text="Date:", font=("Arial", 12, "bold")).grid(row=0, column=0, padx=(10, 2), pady=10)
        self.date_entry = ctk.CTkEntry(frame, width=80, placeholder_text="DD/MM")
        self.date_entry.grid(row=0, column=1, padx=5, pady=10)

        ctk.CTkLabel(frame, text="Vendor:", font=("Arial", 12, "bold")).grid(row=0, column=2, padx=(10, 2), pady=10)
        self.vendor_combobox = ctk.CTkComboBox(frame, values=self.saved_vendors, width=140)
        self.vendor_combobox.grid(row=0, column=3, padx=5, pady=10)

        ctk.CTkLabel(frame, text="Description:", font=("Arial", 12, "bold")).grid(row=0, column=4, padx=(10, 2), pady=10)
        self.desc_entry = ctk.CTkEntry(frame, width=180)
        self.desc_entry.grid(row=0, column=5, padx=5, pady=10)

        ctk.CTkLabel(frame, text="Value (R$):", font=("Arial", 12, "bold")).grid(row=0, column=6, padx=(10, 2), pady=10)
        self.val_entry = ctk.CTkEntry(frame, width=80)
        self.val_entry.grid(row=0, column=7, padx=5, pady=10)

        self.pdf_path_var = ctk.StringVar()
        ctk.CTkButton(frame, text="Attach PDF", width=90, fg_color="#4A4A4A", hover_color="#333333", command=self._select_pdf).grid(row=0, column=8, padx=5, pady=10)
        
        self.clear_pdf_btn = ctk.CTkButton(frame, text="Clear PDF", width=80, fg_color="#C21807", hover_color="#960018", command=self._clear_pdf)
        self.clear_pdf_btn.grid(row=0, column=9, padx=5, pady=10)
        self.clear_pdf_btn.grid_remove() 
        
        ctk.CTkButton(frame, text="Add Expense", width=100, command=self._add_expense).grid(row=0, column=10, padx=(5, 10), pady=10)

    def _build_treeview(self) -> None:
        columns = ("Date", "Vendor", "Description", "Value", "PDF Path")
        self.tree = ttk.Treeview(self.root, columns=columns, show="headings")
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120)
            
        # NOVO: Configurando a 'tag' visual para as linhas sem PDF
        self.tree.tag_configure("no_pdf", foreground="#ff6666") # Vermelho claro otimizado para modo escuro
            
        self.tree.pack(pady=5, padx=15, fill="both", expand=True)

    def _build_footer_frame(self) -> None:
        footer_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        footer_frame.pack(pady=10, padx=15, fill="x")

        self.total_var = ctk.StringVar()
        self.total_var.set("Total: R$ 0,00")
        total_label = ctk.CTkLabel(footer_frame, textvariable=self.total_var, font=("Arial", 18, "bold"), text_color="#28a745")
        total_label.pack(side="left")

        ctk.CTkButton(footer_frame, text="Edit Selected", width=110, fg_color="#5a6268", hover_color="#41474b", command=self._edit_expense).pack(side="left", padx=15)
        ctk.CTkButton(footer_frame, text="Delete Selected", width=110, fg_color="#C21807", hover_color="#960018", command=self._delete_expense).pack(side="left")
        ctk.CTkButton(footer_frame, text="Delete All", width=100, fg_color="#8B0000", hover_color="#5c0000", font=("Arial", 12, "bold"), command=self._delete_all_expenses).pack(side="left", padx=15)

        ctk.CTkButton(footer_frame, text="Merge PDFs", width=130, height=35, fg_color="#28a745", hover_color="#218838", font=("Arial", 14, "bold"), command=self._merge_pdfs).pack(side="right")

    def _select_pdf(self) -> None:
        file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if file_path:
            self.pdf_path_var.set(file_path)
            messagebox.showinfo("Success", "PDF attached temporarily.")

    def _clear_pdf(self) -> None:
        self.pdf_path_var.set("")
        messagebox.showinfo("Info", "PDF attachment removed.")
        self.clear_pdf_btn.grid_remove()

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
            
            self._refresh_ui()
            self._clear_inputs()
        except ValueError as e:
            messagebox.showerror("Input Error", str(e))

    def _edit_expense(self) -> None:
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select an expense to edit.")
            return

        idx = int(selected[0])
        record = self.records.pop(idx)

        self._clear_inputs()
        self.date_entry.insert(0, record.date.strftime("%d/%m/%Y"))
        self.vendor_combobox.set(record.vendor)
        self.desc_entry.insert(0, record.description)
        self.val_entry.insert(0, str(record.value))
        
        if record.pdf_path:
            self.pdf_path_var.set(record.pdf_path)

        self.clear_pdf_btn.grid()
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

    def _delete_all_expenses(self) -> None:
        if not self.records:
            messagebox.showinfo("Info", "There are no expenses to delete.")
            return
            
        if messagebox.askyesno("Confirm Delete All", "Are you sure you want to delete ALL expenses? This action cannot be undone."):
            self.records.clear()
            self._clear_inputs() 
            self._refresh_ui()   

    def _refresh_ui(self) -> None:
        self.records.sort(key=lambda r: r.date)
        
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        for idx, rec in enumerate(self.records):
            date_str = rec.date.strftime("%d/%m/%Y")
            formatted_value = self._format_currency(rec.value)
            
            # NOVO: Lógica para aplicar a tag vermelha caso não haja PDF
            row_tags = ()
            if rec.pdf_path:
                display_pdf = rec.pdf_path
            else:
                display_pdf = "No PDF"
                row_tags = ("no_pdf",) # Vincula a tag vermelha apenas a esta linha
                
            self.tree.insert("", "end", iid=str(idx), values=(date_str, rec.vendor, rec.description, f"R$ {formatted_value}", display_pdf), tags=row_tags)
            
        total = sum(record.value for record in self.records)
        formatted_total = self._format_currency(total)
        self.total_var.set(f"Total: R$ {formatted_total}")
        
        self._save_expenses()

    def _clear_inputs(self) -> None:
        self.date_entry.delete(0, "end")
        self.vendor_combobox.set("") 
        self.desc_entry.delete(0, "end")
        self.val_entry.delete(0, "end")
        self.pdf_path_var.set("")
        
        self.clear_pdf_btn.grid_remove()

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