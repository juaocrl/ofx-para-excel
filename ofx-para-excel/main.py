import tkinter as tk
from tkinter import filedialog
from ofxparse import OfxParser
import pandas as pd

def ofx_to_dataframe(ofx_file):
    with open(ofx_file, 'rb') as f:
        ofx = OfxParser.parse(f)

    transactions = ofx.account.statement.transactions

    df = pd.DataFrame([t.__dict__ for t in transactions])

    return df

def main():
    root = tk.Tk()
    root.withdraw()  # Ocultar a janela principal do tkinter

    # Abrir uma caixa de diálogo para selecionar o arquivo OFX
    ofx_file = filedialog.askopenfilename(title="Selecione o arquivo OFX", filetypes=[("Arquivos OFX", "*.ofx")])

    if not ofx_file:
        print("Nenhum arquivo selecionado. Encerrando o programa.")
        return

    # Converter o arquivo OFX em um DataFrame
    df = ofx_to_dataframe(ofx_file)

    # Abrir uma caixa de diálogo para selecionar onde salvar a planilha
    save_file = filedialog.asksaveasfilename(title="Salvar a planilha como", defaultextension=".xlsx", filetypes=[("Arquivos Excel", "*.xlsx")])

    if not save_file:
        print("Nenhum local de salvamento selecionado. Encerrando o programa.")
        return

    # Exportar o DataFrame para o local de salvamento escolhido (Excel)
    df.to_excel(save_file, index=False)
    print(f"Arquivo Excel salvo em {save_file}")

if __name__ == "__main__":
    main()