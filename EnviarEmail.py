import smtplib
import pyodbc
import pandas as pd
import time
import re

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from email.mime.base import MIMEBase
from email import encoders

regex = '^[a-z0-9]+[\._]?[a-z0-9]+[@]\w+[.]\w{2,3}$'
def check(email):   
    if(re.search(regex,email)):  
        print("Email Válido")      
    else:  
        print("Email Inválido")  


#CONEXAO COM BANCO DE DADOS
print('')
print('Iniciando...')
#tempo.sleep(2)
print('Conectando ao banco de dados')
dados_conexao = (
   'DRIVER={SQL SERVER Native CLient 11.0};SERVER=192.168.1.239;DATABASE=sapiens;UID=sapiens;PWD=sapiens'
)
conexao = pyodbc.connect(dados_conexao)
#tempo.sleep(4)
print("Conexão Bem Sucedida")
print('')

#FAZENDO O SELECT NO BANCO PARA ACHAR OS CLIENTES
Estado = input('Digite o estado: ')

#Uf = ['AC','AL','AP','AM','BA','CE','DF','ES','GO','MA','MT','MS','MG','PA','PB','PR','PE','PI','RJ','RN','RS','RO','RR','SC','SP','SE','TO']

if len(Estado) != 2:
    print('Uf não é valida. Verifique!')
else:
    ImgNet =  input('Digite o Nome da Imagem: ')

    tituloEml = input('Digite o Titulo do e-mail: ')

    sqlEXEC = f'''SELECT  NOMCLI,INTNET
                    FROM e085cli 
                   WHERE codcli not in (1,2,3,4,5,6) 
                     AND tipcli = 'J' AND SitCli = 'A' 
                     AND sigufs = '{Estado}' '''
    SQL2 = f'SELECT NOMCLI,INTNET FROM E085CLI WHERE CODCLI IN (11300,7139,6846)'  #7139
    #dataSQL = pd.read_sql(sqlEXEC,conexao)
    #dados = pd.DataFrame(dataSQL)

    #Lista = list(dados['INTNET'])

    # reading the spreadsheet
    email_list = pd.read_sql(sqlEXEC,conexao)
    print('Total de e-mail para enviar')
    print(len(email_list))
    print('\n iniciando ....')

    # getting the names and the emails
    #names = email_list['NOMCLI']
    emails = email_list['INTNET']

    # iterate through the records
    for i in range(len(emails)):
        # for every record get the name and the email addresses
        eml = emails[i]
        print(eml)
    
        #1 - Startar o Servidor SMTP
        host = "smtp.clarice.com.br"
        port = "587"
        login = "comercial1@clarice.com.br"
        senha = "cla5800"
    
        server  = smtplib.SMTP(host,port)
        server.ehlo()
        server.starttls()
        server.login(login,senha)
    
        #print('Chegou aqui 1.0')
        #2 - CONSTRUIR O EMAIL TIPO MIME

        corpo = f"""
        <p>Olá amigo cliente, tudo bem?.</p>
        
        <p style="color:red;"><h2>O dia das mães está chegando e VOCÊ NÃO VAI DEIXAR PARA COMPRAR NA ULTIMA HORA NÃO É?</h2></p>
        
        <p><h3>Pensando nisso elaboramos uma oferta SUPER ESPECIAL, confira na imagem a baixo.<h3></p>
        
        <img src='http://ftp.clarice.com.br/webimagens/{ImgNet} '>
        
        <p><h3>QUER SABER MAIS DESSA SUPER OFERTA? FALE DIRETO COM UM DE NOSSOS REPRESENTANTES:</h3></p>
        <p><h2>Você pode clicar no link: <a href="https://wa.me/message/SQ7LDPXE6XQ6H1">OfertaClarice</a></h2></p>  
        <p><h3>Ou responder esse e-mail, se preferir ligue para: 49 9149-5447</h3></p>
        """
        email_msg = MIMEMultipart()
        email_msg['From'] = login
        email_msg['To']   = eml
        email_msg['Subject'] = tituloEml
        email_msg.attach(MIMEText(corpo,'html'))
    
        #Enviar uma arquivo como Anexo.
        #attchment = open(ImgNet,'rb')
        #att = MIMEBase('application', 'octet-stream')
        #att.set_payload(attchment.read())
        #encoders.encode_base64(att)
        #att.add_header('Content-Disposition',f'attachment; filename = Combo_600X600.jpg')
        #attchment.close()
        #email_msg.attach(att)
    
        #3 - ENVIAR O EMAIL tipo MIME no Servidor SMTP
        try:
            server.sendmail(email_msg['From'],email_msg['To'],email_msg.as_string())
        except:
            print('E-mail não existe.')

        server.quit()
    
        print('Email enviado com Sucesso!')
        time.sleep(2)

    time.sleep(3)
    print('\n\nTodos os e-mails foram enviados com sucesso.')
