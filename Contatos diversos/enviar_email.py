import win32com.client as win32
import os


# simple code to send e-mails through win32com.client libary
os.system('taskkill /im outlook.exe /f')
outlook = win32.Dispatch('outlook.application')

email = outlook.CreateItem(0)
email.to = 'felipe.rodrigues@ciaathletica.com.br'
email.Subject = 'Teste código e-mail'
email.HTMLBody = r'''
<p>Oi,</p>
<p></p>
<p>Esse é um teste de e-mail pelo código python.</p>
<p><img src="\\Qnapcia\rh\01 - RH\01 - Administração.Controles\08 - Dados, Documentos e Declarações Diversas\
Logo Cia\Assinatura.png"></p>
'''
email.Send()
