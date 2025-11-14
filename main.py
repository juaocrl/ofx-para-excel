import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import ImageTk, Image
from ofxparse import OfxParser
import pandas as pd
from datetime import datetime

def ofx_to_dataframe(ofx_file, banco, nome_banco):
    with open(ofx_file, 'rb') as f:
        ofx = OfxParser.parse(f)

    transactions = ofx.account.statement.transactions

    data = []
    for t in transactions:
        
        debito = banco if t.amount >= -1 else ""
        credito = banco if t.amount <= -1 else ""
        tipocred = ()
        if t.amount <= -1 :
            tipocred= "deb"
        else: 
            tipocred= "cred"
        
        excel_date = t.date.toordinal() - datetime(1899, 12, 30).toordinal()
        
        data.append({
            "debito": debito,
            "credito": credito,
            "data": excel_date,
            "valor": str(t.amount).replace(".", ","),
            "codigo do historico": "",
            "n.documento": "",
            "nome do emitente": (tipocred + " C/C " + nome_banco).upper(),
            "complemento do historico": t.memo.upper(),
            "historico total": (tipocred + " c/c " + nome_banco + " " + t.memo).upper(),
        })

    df = pd.DataFrame(data)
    return df

def process_ofx_files():
    banco = entry_numero_banco.get()
    nome_banco = entry_nome_banco.get()
    
    if not banco or not nome_banco:
        messagebox.showwarning("Entrada inválida", "Por favor, preencha todos os campos.")
        return

    ofx_files = filedialog.askopenfilenames(title="Selecione os arquivos OFX", filetypes=[("Arquivos OFX", "*.ofx")])
    if not ofx_files:
        messagebox.showinfo("Nenhum arquivo", "Nenhum arquivo selecionado. Encerrando o programa.")
        return

    dfs = []
    for ofx_file in ofx_files:
        df = ofx_to_dataframe(ofx_file, banco, nome_banco)
        dfs.append(df)

    combined_df = pd.concat(dfs, ignore_index=True)
    combined_df = combined_df[["debito", "credito", "data", "valor", "codigo do historico", "n.documento", "nome do emitente", "complemento do historico", "historico total"]]

    save_file = filedialog.asksaveasfilename(title="Salvar a planilha como", defaultextension=".xlsx", filetypes=[("Arquivos Excel", "*.xlsx")])
    if not save_file:
        messagebox.showinfo("Nenhum local de salvamento", "Nenhum local de salvamento selecionado. Encerrando o programa.")
        return

    try:
        combined_df.to_excel(save_file, index=False)
        messagebox.showinfo("Sucesso", f"Arquivo Excel salvo em {save_file}")
    except PermissionError:
        messagebox.showerror("Erro de Permissão", f"Não foi possível salvar o arquivo em {save_file}. Tente um local diferente.")

def main():
    tela = tk.Tk()
    tela.title("Conversor OFX para Excel")
    tela.resizable(False, False)
    tela.iconbitmap("logo_cadasto.ico")

    # Layout
    tk.Label(tela, text="Número do Banco:").grid(row=0, column=0, padx=10, pady=10)
    global entry_numero_banco
    entry_numero_banco = tk.Entry(tela)
    entry_numero_banco.grid(row=0, column=1, padx=10, pady=10)

    tk.Label(tela, text="Nome do Banco:").grid(row=1, column=0, padx=10, pady=10)
    global entry_nome_banco
    entry_nome_banco = tk.Entry(tela)
    entry_nome_banco.grid(row=1, column=1, padx=10, pady=10)

    btn_processar = tk.Button(tela, text="Selecionar e Processar Arquivos OFX", command=process_ofx_files)
    btn_processar.grid(row=2, column=0, columnspan=2, padx=10, pady=20)

    tela.mainloop()

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        messagebox.showerror("Erro", f"Não foi possível rodar: {e}")
