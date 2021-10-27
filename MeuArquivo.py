# import pandas as pd -- pip install pandas
import pandas as pd
# import pywin32 as win32 -- pip install pywin32
import win32com.client as win32

# instalar -- pip install openpyxl , para ler a tabela
# instalar pywin32

# Importar a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')


# Visualizar a base de dados
pd.set_option('display.max_columns', None)  # para aparecer todas as colunas
print(tabela_vendas)


# Faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby(
    'ID Loja').sum()
print(faturamento)

# Quantidade de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)

print('-' * 50)

# Ticket médio por produto em cada loja "em medio quanto custou um produto que a loja vendeu"
# to_frame trasforme os dados em tabela
ticket_medio = (faturamento['Valor Final'] /
                quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(
    columns={0: 'Ticket Médio'})  # Muda o nome da coluna
print(ticket_medio)

# Enviar um email com o relatório "Tem que ter o outlook configurado no email"
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'binglassyela@gmail.com'
mail.Subject = 'Relatório de Vendas por Loja'
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
<p>Bruna</p>

'''

mail.Send()

print('Email Enviado')
