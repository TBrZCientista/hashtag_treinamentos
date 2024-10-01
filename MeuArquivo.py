#importar base
#visualizar a base
#calculo de faturamento por loja
#quantidade de produtos vendidos por loja
#ticket médio por produto por loja
#enviar e-mail com relatório

import pandas as pd

tabela_de_vendas = pd.read_excel("Vendas.xlsx")

print(tabela_de_vendas)