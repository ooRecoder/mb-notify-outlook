import tkinter as tk
from tkinter import filedialog

def get():
    """Abre um popup para o usuário selecionar o arquivo da planilha (.xlsx)."""
    root = tk.Tk()
    root.withdraw()  # Oculta a janela principal do Tkinter
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione a planilha",
        filetypes=[("Planilhas Excel", "*.xlsx *.xls")],
        defaultextension=".xlsx"
    )
    root.destroy()
    return caminho_arquivo