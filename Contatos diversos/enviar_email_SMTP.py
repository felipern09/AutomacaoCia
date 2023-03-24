import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from dados import host, port, k1, em_rem
from openpyxl import load_workbook as l_w
# simple code to send e-mails through smtplib

s = smtplib.SMTP(host=host, port=port)
s.starttls()
s.login(em_rem, k1)

wb = l_w('Nomes e e-mails.xlsx')
sh = wb['Dados']
x = 1
while x <= len(sh['A']):
    msg = MIMEMultipart()
    message = f'''
    Olá, {str(sh[f'A{x}'].value).title().split(sep=' ')[0]}!\n
    \n
    Para repor o encontro com colaboradores novatos cancelado no dia 21/03, abrimos novo horário hoje:\n
    24/03:
    14h às 15h30 - Sala 3.
    \n
    Atenciosamente,\n
    Felipe Rodrigues
    '''
    # parameters of the message
    msg['From'] = em_rem
    msg['To'] = str(sh[f'B{x}'].value).lower()
    msg['Subject'] = "Reposição Encontro Cinthia Guimarães"
    msg.attach(MIMEText(message, 'plain', _charset='utf-8'))
    s.send_message(msg)
    del msg
    x += 1
s.quit()
