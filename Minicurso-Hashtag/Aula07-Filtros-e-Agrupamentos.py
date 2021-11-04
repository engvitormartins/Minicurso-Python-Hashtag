# importar bibliotecas

import pandas as pd

# importar a base de dados

tabela_vendas = pd.read_excel('Vendas.xlsx')

# visualizar a base de dados
pd.set_option('display.max_columns', None)

# print(tabela_vendas)
print(tabela_vendas.head())

# faturamento por loja
    # tabela_vendas[['ID Loja', 'Valor Final']]

    # Criar uma lista com cada uma das lojas e do lado a soma do faturamento
        # tabela_vendas.groupby('ID Loja').sum()

    # tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()

faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()

print(faturamento)


# quantidade de produtos vendidos por loja


# ticket médio por produto em cada loja

# enviar um e-amil com o relatório