# importar bibliotecas

import pandas as pd

# importar a base de dados

tabela_vendas = pd.read_excel('Vendas.xlsx')

# visualizar a base de dados
pd.set_option('display.max_columns', None)

# print(tabela_vendas)
print('\nTabela de Vendas: ')
print(tabela_vendas.head())
print('-' * 50)

# faturamento por loja
    # tabela_vendas[['ID Loja', 'Valor Final']]

    # Criar uma lista com cada uma das lojas e do lado a soma do faturamento
        # tabela_vendas.groupby('ID Loja').sum()

    # tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()

faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print('\nTabela de Faturamento: ')
print(faturamento)
print('-' * 50)


# quantidade de produtos vendidos por loja

quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print('\nTabela de Quantidades: ')
print(quantidade)
print('-' * 50)

# ticket médio por produto em cada loja
    # usar .to_frame() para retirar o flot64 em baixo e transformar em uma tabela
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
print('\nTabela de Ticket Médio: ')
print(ticket_medio)
print('-' * 50)

# enviar um e-amil com o relatório

