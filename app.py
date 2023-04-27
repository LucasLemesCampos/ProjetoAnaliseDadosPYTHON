import pandas as pd
import smtplib
import email.message

# Importar a base de Dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

# Visualizar a base de Dados
#pd.set_option('display.max_columns', None)
#print(tabela_vendas)

# Faturamento por Loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()

# Quantidade de Produtos Vendidos por Loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()

# Ticket medio por produto em cada loja
#to_frame() - Para Converter para uma tabela
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})

# Enviar um Email com Relatorio

corpo_email = f"""
<p>Bom dia</p>

<p>Segue Relatorio de Venda por cada loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{quantidade.to_html()}

<p>Ticket Medio por Produto:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer Duvida Estou a Disposição.</p>

<p>Att</p>
<p>Lucas</p>
"""

msg = email.message.Message()
msg['Subject'] = "Relatorio Mês Abril"
msg['From'] = 'lucascampos.mbl@gmail.com'
msg['To'] = 'lucascampos.mbl@gmail.com'
password = 'zwhvktvpwrqtcnjp'
msg.add_header('Content-Type', 'text/html')
msg.set_payload(corpo_email )

s = smtplib.SMTP('smtp.gmail.com: 587')
s.starttls()
# Login Credentials for sending the mail
s.login(msg['From'], password)
s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
print('Email enviado')


