#INSTALANDO BIBLIOTECA PANDAS QUE SERVE PARA LER TABELAS DO EXCEL, WIN32 PARA ENIVO DE EMAIL
import pandas as pd
import win32com.client as Win32
 
print('-' * 50)
#IMPORTAR A BASE DE DADOS
tabela_vendas = pd.read_excel('Vendas.xlsx')

print('-' * 50)
#VISUALIZAR A BASE DE DADOS 
pd.set_option('display.max_columns',None)
print(tabela_vendas)

print('-' * 50)
#FATURAMENTO POR LOJA
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

print('-' * 50)
#QUANTIDADE  DE PRODUTOS VENDIDOS POR LOJA
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)

print('-' * 50)
#TICKET MÉDIO POR PRODUTO EM CADA LOJA 
ticket_médio = (faturamento ['Valor Final'] / quantidade ['Quantidade']).to_frame()
ticket_médio = ticket_médio.rename(columns={0: 'Ticket Médio'})
print(ticket_médio)

#ENVIAR UM EMAIL COM RELATORIO, USANDO F NA FRENTE DAS ASPAS PODE USAR AS VARIAVEIS DENTRO DAS CHAVES

outlook = Win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'medeirosmary94@gmail.com'
mail.subject = 'RELATÓRIO DE VENDA POR LOJA'
mail.HTMLBody = f''' PREZADOS,
SEGUE O RELATORIO DE VENDAS POR CADA LOJA.
<p>FATURAMENTO</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>QUANTIDADE VENDIDA:</p>
{quantidade.to_html()}

<p> TICKET MEDIO DOS PRODUTOS EM CADA LOJA:<p/>
{ticket_médio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

Qualquer duvida estou a disposição.
Att..
Mary Medeiros

'''
mail.Send()
print ('EMAIL ENVIADO COM SUCESSO!')