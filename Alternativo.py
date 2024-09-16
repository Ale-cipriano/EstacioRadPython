import tkinter as tk
from tkinter import scrolledtext
import pandas as pd

# Carregar a planilha Excel usando pandas
tabela = pd.read_excel("GESTAO_DE_EXAMES_PERIODICOS.xlsx")

# Converter apenas colunas que contenham datas para o formato adequado (dd-mm-aaaa)
for coluna in tabela.columns:
    if pd.api.types.is_datetime64_any_dtype(tabela[coluna]):
        tabela[coluna] = pd.to_datetime(tabela[coluna], errors='coerce').dt.strftime('%d-%m-%Y')  # Formata para 'dd-mm-aaaa'

# Função para exibir a tabela na interface tkinter
def exibir_tabela():
    # Cria uma janela principal
    janela = tk.Tk()
    janela.title("Exibição da Tabela")

    # Cria um widget de texto com barra de rolagem para exibir a tabela
    txt_tabela = scrolledtext.ScrolledText(janela, width=300, height=80, wrap=tk.WORD)
    txt_tabela.pack(padx=10, pady=10)

    # Formatar os dados do DataFrame como string e inseri-los no widget
    tabela_str = tabela.to_string(index=False)  # Converte DataFrame para string sem os índices
    txt_tabela.insert(tk.END, tabela_str)

    # Torna o widget de texto somente leitura
    txt_tabela.config(state=tk.DISABLED)

    # Inicia o loop principal da interface
    janela.mainloop()

# Exibir a tabela na interface
exibir_tabela()