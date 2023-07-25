import urllib
from urllib import parse
from openpyxl import load_workbook as l_w


# simple code to send whatsapp messages through browser
wb = l_w("AV.xlsm")
sh = wb['Planilha1']

for x in range(2, len(sh['A'])):
    pessoa = str(sh[f"B{x}"].value).split(' ')[0]
    email = str(sh[f"D{x}"].value)
    numero = str(sh[f"H{x}"].value)
    url = str(sh[f"F{x}"].value)
    mensagem = f'Oi {pessoa}, te enviei por e-mail(no {email}) o resultado da primeira etapa da sua avaliação de ' \
               f'desempenho e o link da pesquisa sobre a avaliação. Antes da segunda etapa, precisamos que ' \
               f'responda a pesquisa. Ok? Se puder responder agora, é bem rápido, dura no máximo 5 minutos. ' \
               f'Segue o link: {url}'
    texto = urllib.parse.quote(mensagem)
    cel = urllib.parse.quote(numero)
    link = f'https://web.whatsapp.com/send?phone={cel}&text={texto}'
    if numero:
        print(link)
        print(pessoa)
        print(email)
        print(numero)
    x += 1
