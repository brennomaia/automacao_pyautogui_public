import pyautogui
from pyperclip import copy as copia
import time
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


pyautogui.PAUSE = 1
print('Aguarde estamos preparando os dados do Faturamento...')

# Passo 1 - Abrir o sistema(google drive)
link = '<ENDERECO_DRIVE>'
copia(r'LOCAL_CHROME.EXE')
pyautogui.hotkey('win', 'r')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')

copia(link)
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')

# Passo 2 - Baixar a base de dados
time.sleep(2)
pyautogui.click(x=1727, y=285)
pyautogui.click(x=2758, y=187)
time.sleep(2)
pyautogui.click(x=2609, y=621)


# Passo 3 Calcular Faturamento
tabela = pd.read_excel('<LOCAL_ARQUIVO>')
faturamento = tabela['Valor Final'].sum()
quantidade = tabela['Quantidade'].sum()
print('Pronto!')

# Passo 4 Enviando Email.
email_destino = input('Digite seu email para receber o faturamento: ')
host = 'smtp.gmail.com'
port = 587
user = '<CONTA_EMAIL>'
password = '<SENHA>'

server = smtplib.SMTP(host, port)
server.ehlo()
server.starttls()
server.login(user, password)

# Criando mensagem
message = f"""Prezados(as).
O faturamento do mês de Dezembro foi de R${faturamento:.2f}
A quantidade de {int(quantidade)} produtos vendidos!

Parabens, pelo resultado!!

Att, 
Estudante Pythonico """
email_msg = MIMEMultipart()
email_msg['From'] = user
email_msg['To'] = email_destino
email_msg['Subject'] = 'Faturamento Mensal'

email_msg.attach(MIMEText(message, 'plain'))
print('Enviando Email...')
server.sendmail(email_msg['From'], email_msg['To'], email_msg.as_string())
print('Email enviada!')
server.quit()






