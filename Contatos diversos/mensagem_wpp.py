import urllib
from openpyxl import load_workbook as l_w
# simple code to send whatsapp messages through browser
wb = l_w("AV.xlsm")
sh = wb['Planilha1']
x=1
while x<139:
    pessoa = str(sh[f"B{x}"].value).split(' ')[0]
    email = str(sh[f"D{x}"].value)
    numero = str(sh[f"H{x}"].value)
    url = str(sh[f"F{x}"].value)
    mensagem = f'Oi {pessoa}, te enviei por e-mail(no {email}) o resultado da primeira etapa da sua avaliação de desempenho e o link da pesquisa sobre a avaliação. Antes da segunda etapa, precisamos que responda a pesquisa. Ok? Se puder responder agora, é bem rápido, dura no máximo 5 minutos. Segue o link: {url}'
    texto = urllib.parse.quote(mensagem)
    cel = urllib.parse.quote(numero)
    link = f'https://web.whatsapp.com/send?phone={cel}&text={texto}'
    if numero:
        print(link)
        print(pessoa)
        print(email)
        print(numero)
    x=x+1