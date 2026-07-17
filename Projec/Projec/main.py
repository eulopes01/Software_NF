import tkinter as tk
from ui.app_window import ExpenseApp

def main() -> None:
    root = tk.Tk()
    root.geometry("800x400")
    # Inicia a aplicação instanciando a View principal
    app = ExpenseApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()