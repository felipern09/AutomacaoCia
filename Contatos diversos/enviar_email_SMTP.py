import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

s = smtplib.SMTP(host=usuario.servidor, port=usuario.porta)
s.starttls()
s.login(email_remetente, senha)

msg = MIMEMultipart()
message = f'''
        Olá, {str(personal.nome).title().split(sep=' ')[0]}!\n
        \n
        Seguem dados para o pagamento:\n
        \n
        PIX: 03732305000186
        Banco: Itaú
        Agência: 6205
        C/C: 01588-3\n
        Assim que o pagamento for feito, favor responder esse e-mail com o comprovante bancário.
        \n
        Atenciosamente,\n
        Marcelo Gonçalves
        '''
# parameters of the message
msg['From'] = email_remetente
msg['To'] = 'felipe.rodrigues@ciaathletica.com.br'
msg['Subject'] = "Aulas de personal - Cia Athletica"
msg.attach(MIMEText(message, 'plain', _charset='utf-8'))

# Anexo PNG
arquivo_png = usuario.assinatura
with open(arquivo_png, 'rb') as img_file:
    imagem = MIMEImage(img_file.read())
    msg.attach(imagem)

s.send_message(msg)
del msg
s.quit()
