import pandas as pd

# Carregar o arquivo Excel
df = pd.read_excel('Resultado1.xlsx', sheet_name='Sheet1')

# Excluir as linhas em branco na coluna "ColunaAlvo"
df = df.dropna(subset=['Número'])

# Salvar o DataFrame em um novo arquivo Excel
df.to_excel('Resultado2.xlsx', index=False)