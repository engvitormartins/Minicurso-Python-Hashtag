# importar bibliotecas

import pandas as pd
import numpy as np
import win32com.client as win32
import xlsxwriter




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
#print('\nTabela de Faturamento: ')
#print(faturamento)
#print('-' * 50)


# quantidade de produtos vendidos por loja

quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
#print('\nTabela de Quantidades: ')
#print(quantidade)
#print('-' * 50)

# ticket médio por produto em cada loja
    # usar .to_frame() para retirar o flot64 em baixo e transformar em uma tabela
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
#print('\nTabela de Ticket Médio: ')
#print(ticket_medio)
#print('-' * 50)

# enviar um e-amil com o relatório



tabela_tratada = pd.DataFrame()
tabela_tratada['FATURAMENTO'] = faturamento
tabela_tratada['QTDD'] = quantidade
tabela_tratada['TICKET MEDIO'] = ticket_medio
# tabela_tratada.to_excel('Dados-tratado.xlsx')

#print('\nTabela Tratada: ')
#print(tabela_tratada)
#print('-' * 50)

pivot = pd.pivot_table(tabela_vendas, index= ['ID Loja'], values='Valor Final', aggfunc='sum')
pivot.style.format({'Valor Final':'R${:,.2f}'})


pivot.to_excel('Dados-Pivot.xlsx')

print('\nTabel Pivot: ')
print(pivot)
print('-' * 50)



outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'vitormartinshp@gmail.com'
mail.Subject = 'Relatório de Vendas por Loja'
#mail.Attachments.Add('c:\\Users\\T-Gamer\\OneDrive\\01_Github_engvitormartins\\MeusProjetos'
#                     '\\Minicurso-Python-Hashtag\\Minicurso-Hashtag\\Dados-tratado.xlsx')

mail.HTMLBody = f'''


<p>Prezados,</p>

<p>Segue o Relatório de Vendas por cada Loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{quantidade.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição.</p>



<p>Att.,</p>
<p>Vitor Martins</p>
'''

#mail.Send()


# print('Email Enviado com sucesso!')
