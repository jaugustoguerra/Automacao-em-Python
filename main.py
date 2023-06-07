import pandas as pd
from pandas import DataFrame
import win32com.client as win32

tabela_vendas: DataFrame = pd.read_excel('Vendas.xlsx')
print(tabela_vendas)
print('-' * 50)

faturamento = tabela_vendas [['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)
print('-' * 50)

quantidade = tabela_vendas [['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)
print('-' * 50)

ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0:'Ticket Médio'})
print(ticket_medio)
print('-' * 50)

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'splgugulive1@gmail.com'
mail.Subject = 'Relatório de Vendas'
mail.HTMLBody = f'''
<p>Prezados</p>

<p>Segue o Relatório de Vendas por cada Loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final' :'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{quantidade.to_html()}

<p>Ticket Médio por Loja:</p>
{ticket_medio.to_html(formatters={'Valor Final' :'R#{:,.2f}'.format})}

<p>Qualquer dúvida estou a disposição.</p>

<p>Att.,</p>
<p>José Augusto Guerra.</p>
'''

mail.Send()