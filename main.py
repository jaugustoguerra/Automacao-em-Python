import pandas as pd
from pandas import DataFrame
import win32com.client as win32

def ler_tabela_vendas(filepath: str) -> DataFrame:
    try:
        return pd.read_excel(filepath)
    except Exception as e:
        print(f"Erro ao ler o arquivo Excel: {e}")
        return pd.DataFrame()

def calcular_faturamento(tabela_vendas: DataFrame) -> DataFrame:
    return tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()

def calcular_quantidade(tabela_vendas: DataFrame) -> DataFrame:
    return tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()

def calcular_ticket_medio(faturamento: DataFrame, quantidade: DataFrame) -> DataFrame:
    ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
    return ticket_medio.rename(columns={0: 'Ticket Médio'})

def enviar_email(faturamento: DataFrame, quantidade: DataFrame, ticket_medio: DataFrame, destinatario: str):
    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = destinatario
        mail.Subject = 'Relatório de Vendas'
        mail.HTMLBody = f'''
        <p>Prezados</p>

        <p>Segue o Relatório de Vendas por cada Loja.</p>

        <p>Faturamento:</p>
        {faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

        <p>Quantidade Vendida:</p>
        {quantidade.to_html()}

        <p>Ticket Médio por Loja:</p>
        {ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

        <p>Qualquer dúvida estou a disposição.</p>

        <p>Att.,</p>
        <p>José Augusto Guerra.</p>
        '''
        mail.Send()
        print("E-mail enviado com sucesso!")
    except Exception as e:
        print(f"Erro ao enviar o e-mail: {e}")

def main():
    filepath = r'C:\Users\Augus\Documents\Dev\Python\Automacao-em-Python\Vendas.xlsx'
    destinatario = 'lineh_kta@hotmail.com'

    tabela_vendas = ler_tabela_vendas(filepath)
    if tabela_vendas.empty:
        return

    print(tabela_vendas)
    print('-' * 50)

    faturamento = calcular_faturamento(tabela_vendas)
    print(faturamento)
    print('-' * 50)

    quantidade = calcular_quantidade(tabela_vendas)
    print(quantidade)
    print('-' * 50)

    ticket_medio = calcular_ticket_medio(faturamento, quantidade)
    print(ticket_medio)
    print('-' * 50)

    enviar_email(faturamento, quantidade, ticket_medio, destinatario)

if __name__ == "__main__":
    main()