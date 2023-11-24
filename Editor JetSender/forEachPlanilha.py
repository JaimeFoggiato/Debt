import pandas as pd

# Leia a planilha do Excel e selecione apenas as colunas desejadas
df = pd.read_excel('PlanilhaJetSender.xlsx', usecols=['Negócio - Título', 'Negócio - Data do Acidente', 'Negócio - Veículo 3º', 'Pessoa - Telefone'])

# Divida a coluna 'Negócio - Título' em 'Negócio' e 'Título'
df[['Negócio', 'Título']] = df['Negócio - Título'].str.split('-', n=1, expand=True)

# Separe a coluna 'Pessoa - Telefone' em várias linhas
df['Pessoa - Telefone'] = df['Pessoa - Telefone'].str.split(', ')
df = df.explode('Pessoa - Telefone')

# Renomeie as colunas 'var_name' e 'value_name'
df = df.rename(columns={'var_name': 'Telefone', 'value_name': 'Número'})

# Edite a coluna 'Título' para deixar somente a primeira letra maiúscula e manter o primeiro e último nome
df['Título'] = df['Título'].str.title()
df['Título'] = df['Título'].str.split(' ')
df['Título'] = df['Título'].str[0] + ' ' + df['Título'].str[-1]

# Edite a coluna 'Número' para adicionar '55' na frente, remover hífens, parenteses e espaços e transformar em número
df['Pessoa - Telefone'] = '55' + df['Pessoa - Telefone'].str.replace('[-() ]', '', regex=True).astype(int).astype(str)

# Use a função melt para transformar as colunas de telefone em linhas
df = df.melt(id_vars=['Título', 'Negócio - Data do Acidente', 'Negócio - Veículo 3º', 'Pessoa - Telefone'], var_name='Telefone', value_name='Número')

# Remova as linhas com valores em branco na coluna 'Número'
df = df.dropna(subset=['Pessoa - Telefone'])

# Remova as colunas indesejadas
df = df.drop(['Telefone', 'Número'], axis=1)

# Salve o resultado em um novo arquivo ou faça o que precisar com os dados
df.to_excel('Resultado1.xlsx', index=False)