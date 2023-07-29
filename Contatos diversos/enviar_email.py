from src.models.models import Colaborador, engine
from sqlalchemy.orm import sessionmaker
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from src.models.dados_servd import em_rem, k1, host, port


# # code to send e-mails through win32com.client libary
matriculas = input('Digite as matriculas a serem comunicadas separando-as por vírgula: ')
mensagem = input('Mensagem: ')
mat = matriculas.split(',')
Sessions = sessionmaker(bind=engine)
session = Sessions()

for item in mat:
    pessoa = session.query(Colaborador).filter_by(matricula=item).first()
    if pessoa:
        email_remetente = em_rem
        senha = k1
        # set up smtp connection
        s = smtplib.SMTP(host=host, port=port)
        s.starttls()
        s.login(email_remetente, senha)

        # send e-mail to employee with a pdf file so he/she can go to bank to open an account
        msg = MIMEMultipart('alternative')
        msg['From'] = email_remetente
        msg['To'] = pessoa.email
        msg['Subject'] = "Carta para Abertura de conta"
        text = MIMEText(f'''Olá, {str(pessoa.nome).title().split(" ")[0]}!<br><br>
        {mensagem}.<br>
        Atenciosamente,<br>
        <img src="cid:image1">''', 'html')

        # set up the parameters of the message
        msg.attach(text)
        image = MIMEImage(
            open(r'C:\Users\Felipe Rodrigues\PycharmProjects\AutomacaoCia\Admissao\static\assinatura.png', 'rb').read())
        image.add_header('Content-ID', '<image1>')
        msg.attach(image)
        s.sendmail(email_remetente, pessoa.email, msg.as_string())
        del msg
        s.quit()
