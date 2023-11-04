import tkinter as tk
from tkinter import filedialog
from ofxparse import OfxParser
import pandas as pd
import os

def ofx_to_dataframe(ofx_file, banco, nome_banco):
    with open(ofx_file, 'rb') as f:
        ofx = OfxParser.parse(f)

    transactions = ofx.account.statement.transactions

    # Criar o DataFrame com as colunas desejadas e preencher os valores
    data = []
    for t in transactions:
        debito = banco if t.type == "credit" else ""
        credito = banco if t.type == "debit" else ""
        data.append({
            "debito": debito,
            "credito": credito,
            "data": t.date.strftime("%d/%m/%Y"),  # Formatar a data como dd/mm/aaaa
            "valor": str(t.amount).replace(".", ","),  # Substituir ponto por vírgula
            "codigo do historico": "",
            "n.documento": "",
            "nome do emitente": (t.type + " C/C " + nome_banco).upper(),  # Converter para caixa alta
            "complemento do historico": t.memo.upper(),  # Converter para caixa alta
            "historico total": (t.type + " " + nome_banco + " " + t.memo).upper(),  # Converter para caixa alta
        })

    df = pd.DataFrame(data)

    return df

def main():
    root = tk.Tk()
    root.withdraw()  # Ocultar a janela principal do tkinter
    banco = input("Insira o número do banco cadastrado no plano de contas: ")
    nome_banco = input("Insira o nome do banco: ")
    
    # Abrir uma caixa de diálogo para selecionar os arquivos OFX
    ofx_files = filedialog.askopenfilenames(title="Selecione os arquivos OFX", filetypes=[("Arquivos OFX", "*.ofx")])

    if not ofx_files:
        print("Nenhum arquivo selecionado. Encerrando o programa.")
        return

    # Inicializar uma lista para armazenar os DataFrames de cada arquivo
    dfs = []
    
    for ofx_file in ofx_files:
        # Converter o arquivo OFX em um DataFrame
        df = ofx_to_dataframe(ofx_file, banco, nome_banco)
        dfs.append(df)

    # Concatenar todos os DataFrames em um único DataFrame
    combined_df = pd.concat(dfs, ignore_index=True)

    # Renomear as colunas conforme necessário
    combined_df = combined_df[["debito", "credito", "data", "valor", "codigo do historico", "n.documento", "nome do emitente", "complemento do historico", "historico total"]]

    # Abrir uma caixa de diálogo para selecionar onde salvar a planilha
    save_file = filedialog.asksaveasfilename(title="Salvar a planilha como", defaultextension=".xlsx", filetypes=[("Arquivos Excel", "*.xlsx")])

    if not save_file:
        print("Nenhum local de salvamento selecionado. Encerrando o programa.")
        return

    # Exportar o DataFrame modificado para o local de salvamento escolhido (Excel)
    combined_df.to_excel(save_file, index=False)
    print(f"Arquivo Excel salvo em {save_file}")

if __name__ == "__main__":
    main()
