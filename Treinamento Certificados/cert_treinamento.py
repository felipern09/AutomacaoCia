import win32com.client as win32
from docx2pdf import convert
from openpyxl import load_workbook
import docx
import os
# Subistituir nome nos modellos de certificados e salvar como em uma pasta da área de trabalho
outlook = win32.Dispatch('outlook.application')
wb = load_workbook("Treinamento.xlsx")
pasta = 'C:\\Users\\Felipe Rodrigues\\PycharmProjects\\AutomacaoCia\\Treinamento Certificados\\Certificados'
cterr1 = 'C:\\Users\\Felipe Rodrigues\\PycharmProjects\\AutomacaoCia\\Treinamento Certificados\\TreinamentoTerrestre1.docx'
cterr2 = 'C:\\Users\\Felipe Rodrigues\\PycharmProjects\\AutomacaoCia\\Treinamento Certificados\\TreinamentoTerrestre2.docx'
caqua1 = 'C:\\Users\\Felipe Rodrigues\\PycharmProjects\\AutomacaoCia\\Treinamento Certificados\\TreinamentoAquat1.docx'
caqua2 = 'C:\\Users\\Felipe Rodrigues\\PycharmProjects\\AutomacaoCia\\Treinamento Certificados\\TreinamentoAquat2.docx'
# Terrestre 1
# x=6
# while x<=6:
#     t1 = docx.Document(cterr1)
#     sh = wb['TreinamentoTerrestre1']
#     nome = str(sh[f'B{x}'].value)
#     endeletr = str(sh[f'C{x}'].value)
#     doc = t1
#     for p in doc.paragraphs:
#         if '#nome' in p.text:
#             inline = p.runs
#         # Loop added to work with runs (strings with same style)
#             for i in range(len(inline)):
#              if '#nome' in inline[i].text:
#                 text = inline[i].text.replace('#nome', nome)
#                 inline[i].text = text
#     doc.save(pasta+f'\\{nome} PST1.docx')
#     convert(pasta+f'\\{nome} PST1.docx', pasta+f'\\{nome} PST1.pdf')
#     os.remove(pasta+f'\\{nome} PST1.docx')

    # email = outlook.CreateItem(0)
    # email.to = endeletr
    # email.Subject = 'Certificado Curso Primeiros Socorros'
    # email.HTMLBody = f'''
    # <p>Boa tarde,</p>
    # <p></p>
    # <p>Segue certificado do curso de primeiros socorros.</p>
    # <p></p>
    # <p>Atenciosamente,</p>
    # <p><img src="\\\Qnapcia\\rh\\01 - RH\\01 - Administração.Controles\\08 - Dados, Documentos e Declarações Diversas\\Logo Cia\\Assinatura.png"></p>
    # '''
    # anexo = pasta+f'\\{nome} PST1.pdf'
    # email.Attachments.Add(anexo)
    # email.Send()
    # x=x+1

# #Terrestre2
# x=2
# while x<50:
#     t2 = docx.Document(cterr2)
#     sh = wb['TreinamentoTerrestre2']
#     nome = str(sh[f'B{x}'].value)
#     endeletr = str(sh[f'C{x}'].value)
#     doc = t2
#     for p in doc.paragraphs:
#         if '#nome' in p.text:
#             inline = p.runs
#         # Loop added to work with runs (strings with same style)
#             for i in range(len(inline)):
#              if '#nome' in inline[i].text:
#                 text = inline[i].text.replace('#nome', nome)
#                 inline[i].text = text
#     doc.save(pasta+f'\\{nome} PST2.docx')
#     convert(pasta+f'\\{nome} PST2.docx', pasta+f'\\{nome} PST2.pdf')
#     os.remove(pasta+f'\\{nome} PST2.docx')
#
#     email = outlook.CreateItem(0)
#     email.to = endeletr
#     email.Subject = 'Certificado Curso Primeiros Socorros'
#     email.HTMLBody = f'''
#         <p>Boa tarde,</p>
#         <p></p>
#         <p>Segue certificado do curso de primeiros socorros.</p>
#         <p></p>
#         <p>Atenciosamente,</p>
#         <p><img src="\\\Qnapcia\\rh\\01 - RH\\01 - Administração.Controles\\08 - Dados, Documentos e Declarações Diversas\\Logo Cia\\Assinatura.png"></p>
#         '''
#     anexo = pasta+f'\\{nome} PST2.pdf'
#     email.Attachments.Add(anexo)
#     email.Send()
#     x=x+1
#Aquatico1
x=2
while x<=2:
    a1 = docx.Document(caqua1)
    sh = wb['TreinamentoAquat1']
    nome = str(sh[f'B{x}'].value)
    endeletr = str(sh[f'C{x}'].value)
    doc = a1
    for p in doc.paragraphs:
        if '#nome' in p.text:
            inline = p.runs
        # Loop added to work with runs (strings with same style)
            for i in range(len(inline)):
             if '#nome' in inline[i].text:
                text = inline[i].text.replace('#nome', nome)
                inline[i].text = text
    doc.save(pasta+f'\\{nome} PSA1.docx')
    convert(pasta+f'\\{nome} PSA1.docx', pasta+f'\\{nome} PSA1.pdf')
    os.remove(pasta+f'\\{nome} PSA1.docx')

    email = outlook.CreateItem(0)
    email.to = endeletr
    email.Subject = 'Certificado Curso Primeiros Socorros - Aquático'
    email.HTMLBody = f'''
            <p>Boa tarde,</p>
            <p></p>
            <p>Segue certificado do curso de primeiros socorros.</p>
            <p></p>
            <p>Atenciosamente,</p>
            <p><img src="\\\Qnapcia\\rh\\01 - RH\\01 - Administração.Controles\\08 - Dados, Documentos e Declarações Diversas\\Logo Cia\\Assinatura.png"></p>
            '''
    anexo = pasta+f'\\{nome} PSA1.pdf'
    email.Attachments.Add(anexo)
    email.Send()
    x=x+1
# #Aquatico2
# x=2
# while x<22:
#     a2 = docx.Document(caqua2)
#     sh = wb['TreinamentoAquat2']
#     nome = str(sh[f'B{x}'].value)
#     endeletr = str(sh[f'C{x}'].value)
#     doc = a2
#     for p in doc.paragraphs:
#         if '#nome' in p.text:
#             inline = p.runs
#         # Loop added to work with runs (strings with same style)
#             for i in range(len(inline)):
#              if '#nome' in inline[i].text:
#                 text = inline[i].text.replace('#nome', nome)
#                 inline[i].text = text
#     doc.save(pasta+f'\\{nome} PSA2.docx')
#     convert(pasta+f'\\{nome} PSA2.docx', pasta+f'\\{nome} PSA2.pdf')
#     os.remove(pasta+f'\\{nome} PSA2.docx')
#
#     email = outlook.CreateItem(0)
#     email.to = endeletr
#     email.Subject = 'Certificado Curso Primeiros Socorros - Aquático'
#     email.HTMLBody = f'''
#                 <p>Boa tarde,</p>
#                 <p></p>
#                 <p>Segue certificado do curso de primeiros socorros.</p>
#                 <p></p>
#                 <p>Atenciosamente,</p>
#                 <p><img src="\\\Qnapcia\\rh\\01 - RH\\01 - Administração.Controles\\08 - Dados, Documentos e Declarações Diversas\\Logo Cia\\Assinatura.png"></p>
#                 '''
#     anexo = pasta+f'\\{nome} PSA2.pdf'
#     email.Attachments.Add(anexo)
#     email.Send()
#     x=x+1
