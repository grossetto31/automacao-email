# %%
# importar bibliotecas
import pandas as pd
import win32com.client as win32

# %%
# importar a base de dados
tabela_vendas = pd.read_excel("Vendas.xlsx")

# %%
# visualizar a base de dados
pd.set_option('display.max_columns', None)
print(tabela_vendas)
print("-" * 50)

# %%
# faturamento por loja
faturamento = tabela_vendas[["ID Loja", "Valor Final"]].groupby(
    "ID Loja").sum()
print(faturamento)
print("-" * 50)

# %%
# quantidade de produtos vendidos por loja
produtos_vendidos = tabela_vendas[[
    "ID Loja", "Quantidade"]].groupby("ID Loja").sum()
print(produtos_vendidos)
print("-" * 50)

# %%
# ticket medio por produto em cada loja
ticket_medio = (faturamento["Valor Final"] /
                produtos_vendidos["Quantidade"]).to_frame()
ticket_medio = ticket_medio.rename(columns={0: "Ticket Médio"})
print(ticket_medio)
print("-" * 50)

# %%
# enviar email com o relatorio
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = "gabriel.rossetto@geobiogas.tech"
mail.Subject = "Relatório de Vendas"
mail.HTMLBody = f'''


<p>Prezados,</p>

<p>Segue o Relatório de Vendas por cada Loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={"Valor Final": "R${:,.2f}".format})}

<p>Quantidade Vendida:</p>
{produtos_vendidos.to_html(formatters={"Quantidade": "{:,}".format})}

<p>Ticket Médio dos Produtos em Cada Loja:</p>
{ticket_medio.to_html(formatters={"Ticket Médio": "R${:,.2f}".format})}

<p>Qualquer dúvida estou à disposição.</p>

<p>Att.,</p>
<p>Gabriel</p>

'''

mail.Send()

print("Email Enviado!")

# %%
