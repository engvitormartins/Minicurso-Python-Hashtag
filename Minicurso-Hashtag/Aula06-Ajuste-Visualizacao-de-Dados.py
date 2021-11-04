# importar bibliotecas

import pandas as pd

# importar a base de dados

tabela_vendas = pd.read_excel('Vendas.xlsx')

# visualizar a base de dados
pd.set_option('display.max_columns', None)

print(tabela_vendas)

# faturamento por loja

# quantidade de produtos vendidos por loja

# ticket médio por produto em cada loja

# enviar um e-amil com o relatório