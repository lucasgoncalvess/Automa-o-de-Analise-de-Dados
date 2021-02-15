

import pandas as pd

df = pd.read_excel(r'/content/drive/MyDrive/Colab Notebooks/Projetos Intensivão/Aula 1/Vendas.xlsx') #caminho do aqruivo vendas
display(df)

"""Tratamaneto de Dados
> Calculo do Faturamento




"""

#faturamento
faturamento = df[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
#pondo em ordem decrescente
faturamento = faturamento.sort_values(by='Valor Final', ascending=False) #false indica que o rankeamento é de forma decrescente
display(faturamento)

""">Quantidade vendida por loja"""

#quantidade
quantidade = df[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
#rankeando
quantidade = quantidade.sort_values(by='ID Loja', ascending=False)
display(quantidade)

""">Ticket Médio: faturamento dividido pela quantidade"""

#ticket médio
ticket_medio = (faturamento['Valor Final']/ quantidade['Quantidade']).to_frame()

#display(ticket_medio)

#renomear a coluna ticket médio
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'}) # a coluna que possui valor 0 passa a receber Ticket Médio
#display(ticket_medio)

#rankeando
ticket_medio = ticket_medio.sort_values(by='Ticket Médio', ascending=False)
display(ticket_medio)

"""Função para enviar e-mail"""

# importando bibliotecas
import smtplib
import email.message

#criando função para enviar e-mail
def enviar_email(resumo_loja, loja):
  server = smtplib.SMTP('smtp.gmail.com:587')  
  corpo_email = f"""
  <p>Ola Loja,</p>
  {resumo_loja.to_html()}
  <p>Obrigado pela atenção</p>"""
    
  msg = email.message.Message()
  msg['Subject'] = f'RESUMO LOJAS - Loja: {loja}'
    
  msg['From'] = 'from@gmail.com'
  msg['To'] = 'to@gmail.com'
  password = "senha"
  msg.add_header('Content-Type', 'text/html')
  msg.set_payload(corpo_email )
    
  s = smtplib.SMTP('smtp.gmail.com: 587')
  s.starttls()
  # Login Credentials for sending the mail
  s.login(msg['From'], password)
  s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
  print('Email enviado')

"""Criando o relatório por Loja"""

#variavel loja
lojas = df['ID Loja'].unique()

#utilizando o for
for loja in lojas:
    tabela_loja = df.loc[df['ID Loja'] == loja, ['ID Loja', 'Quantidade', 'Valor Final']]
    resumo_loja = tabela_loja.groupby('ID Loja').sum()
    resumo_loja['Ticket Médio'] = resumo_loja['Valor Final'] / resumo_loja['Quantidade']
    enviar_email(resumo_loja, loja) #envia email

"""E-mail para a diretoria"""

tabela_diretoria = faturamento.join(quantidade).join(ticket_medio)
enviar_email(tabela_diretoria, "Todas as Lojas")
