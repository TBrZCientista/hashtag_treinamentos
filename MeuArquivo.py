#importar base
#visualizar a base
#calculo de faturamento por loja
#quantidade de produtos vendidos por loja
#ticket médio por produto por loja
#enviar e-mail com relatório

import pandas as pd
import win32com.client as win32

tabela_de_vendas = pd.read_excel("Vendas.xlsx")

pd.set_option("display.max_columns",None)

faturamento = tabela_de_vendas[['ID Loja','Valor Final']].groupby('ID Loja').sum()

#o processo acima, filtrou as colunas ID LOja e Valor final, e depois agrupou as lojas e somou seus valores de venda

print(faturamento)

qtd_prod_vend_loja = tabela_de_vendas[["ID Loja","Quantidade"]].groupby('ID Loja').sum()

#o processo acima, filtrou as culonas ID Loja e Quantidade, e depois agrupou as lojas e a quantidade dos produtos vendidos

print(qtd_prod_vend_loja)

ticket_medio = (faturamento['Valor Final'] / qtd_prod_vend_loja['Quantidade']).to_frame()

#como não podemos fazer operações diretamente entre as tabelas, podemos selecionar dentro das tabelas
#as colunas que queremos que sejam operadas. No final foi utilizado o to_frame para colocar os dados
#numa tabela, pois o retorno sem ele são só um conjunto de dados.

print(ticket_medio)

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'felipefreitasdaschagas@hotmail.com'
mail.Subject = 'Relatório de vendas por Loja'
mail.HTMLBody = '''
Prezados,
Segue o relatório de vendas por cada Loja.

Faturamento:
{}

Quantidade Vendida:
{}

Ticket Médio dos produtos em cada Loja:
{}

Qualquer dúvida estou a disposição.

At.te:
Freitas
'''

mail.Send()