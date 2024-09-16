import tkinter as tk
from tkinter import scrolledtext
from openpyxl import load_workbook
from prettytable import PrettyTable
import datetime

# Carregar a planilha Excel
tabela = load_workbook("GESTAO_DE_EXAMES_PERIODICOS.xlsx", data_only=True)  # 'data_only' exibe valores ao invés de fórmulas

# Obter a aba ativa (a única aba)
aba_ativa = tabela.active

# Criar uma tabela utilizando PrettyTable para formatação
tabela_formatada = PrettyTable()

# Definir os cabeçalhos da tabela com base na 4ª linha da planilha
cabecalhos = [celula.value if celula.value is not None else f"Coluna_{i}" for i, celula in enumerate(aba_ativa[4], start=1)]

# Garantir que os cabeçalhos são únicos
cabecalhos_unicos = []
for i, cabecalho in enumerate(cabecalhos):
    if cabecalho in cabecalhos_unicos:
        cabecalhos_unicos.append(f"{cabecalho}_{i+1}")  # Adiciona índice para evitar duplicados
    else:
        cabecalhos_unicos.append(cabecalho)

# Adicionar os cabeçalhos à tabela
tabela_formatada.field_names = cabecalhos_unicos

# Função para formatar datas no padrão dd-mm-aaaa
def formatar_data(valor):
    if isinstance(valor, datetime.datetime):
        return valor.strftime('%d-%m-%Y')  # Formata data no padrão dd-mm-aaaa
    return valor

# Lista para armazenar as linhas formatadas
linhas_formatadas = []

# Preencher a tabela com as demais linhas (a partir da 5ª linha)
for linha in aba_ativa.iter_rows(min_row=5, values_only=True):
    linha_formatada = [formatar_data(celula) for celula in linha]  # Formatar datas corretamente
    if any(celula is not None for celula in linha_formatada):
        linhas_formatadas.append(linha_formatada)

# Filtrar colunas
colunas_validas = []
for i in range(len(cabecalhos_unicos)):
    if any(linha[i] is not None for linha in linhas_formatadas):
        colunas_validas.append(i)

# Criar uma nova tabela sem colunas vazias
tabela_sem_colunas_vazias = PrettyTable()
tabela_sem_colunas_vazias.field_names = [cabecalhos_unicos[i] for i in colunas_validas]

# Adicionar as linhas à tabela sem colunas vazias
for linha in linhas_formatadas:
    linha_filtrada = [linha[i] for i in colunas_validas]
    tabela_sem_colunas_vazias.add_row(linha_filtrada)

# Função para exibir a tabela na interface tkinter
def exibir_tabela():
    # Cria uma janela principal
    janela = tk.Tk()
    janela.title("Exibição da Tabela")

    # Cria um widget de texto com barra de rolagem para exibir a tabela
    txt_tabela = scrolledtext.ScrolledText(janela, width=300, height=90, wrap=tk.WORD)
    txt_tabela.pack(padx=10, pady=10)

    # Insere a tabela formatada no widget de texto
    txt_tabela.insert(tk.END, str(tabela_sem_colunas_vazias))

    # Torna o widget de texto somente leitura
    txt_tabela.config(state=tk.DISABLED)

    # Inicia o loop principal da interface
    janela.mainloop()

# Exibir a tabela na interface
exibir_tabela()