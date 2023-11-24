import pandas as pd

# Leia a planilha do Excel
df = pd.read_excel('PlanilhaJetSender.xlsx', usecols=['Negócio - Título', 'Negócio - Data do Acidente', 'Negócio - Veículo 3º', 'Pessoa - Telefone'])

df[['Negócio ', 'Título']] = df['Negócio - Título'].str.split('- ', n=1, expand=True)

df['Pessoa - Telefone'] = df['Pessoa - Telefone'].str.split(', ')
df = df.explode('Pessoa - Telefone')

# Edite a coluna 'Título' para deixar somente a primeira letra maiúscula e manter o primeiro e último nome
df['Título'] = df['Título'].str.title()
df['Título'] = df['Título'].str.split(' ')
df['Título'] = df['Título'].str[0] + ' ' + df['Título'].str[-1]


# Use a função melt para transformar as colunas de telefone em linhas
df = df.melt(id_vars=['Título', 'Negócio - Data do Acidente', 'Negócio - Veículo 3º', 'Pessoa - Telefone'], var_name='Telefone', value_name='Número')

df = df.dropna(subset=['Título', 'Negócio - Data do Acidente', 'Negócio - Veículo 3º', 'Pessoa - Telefone'])

df = df.drop(['Telefone', 'Número'], axis=1)

# Salve o resultado em um novo arquivo ou faça o que precisar com os dados
df.to_excel('Resultado1.xlsx', index=False)