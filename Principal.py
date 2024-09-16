from openpyxl import load_workbook
from prettytable import PrettyTable

# Carregar a planilha Excel
tabela = load_workbook("GESTAO_DE_EXAMES_PERIODICOS.xlsx")

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

tabela_formatada.field_names = cabecalhos_unicos  # Define os nomes das colunas

# Preencher a tabela com as demais linhas (a partir da 5ª linha)
for linha in aba_ativa.iter_rows(min_row=5, values_only=True):
    tabela_formatada.add_row(linha)

# Exibir a tabela formatada
print(tabela_formatada)