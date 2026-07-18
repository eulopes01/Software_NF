import customtkinter as ctk
from ui.app_window import ExpenseApp

def main() -> None:
    # Configuração do visual moderno (Dark Mode nativo)
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")
    
    root = ctk.CTk()
    root.geometry("1000x550") # Um pouco mais largo para acomodar o novo visual
    
    app = ExpenseApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()