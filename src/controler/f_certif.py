import docx
from docx2pdf import convert
import locale
import os
import tkinter.filedialog
from tkinter import messagebox
import tkinter.filedialog

locale.setlocale(locale.LC_MONETARY, 'pt_BR.UTF-8')
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
locale.setlocale(locale.LC_MONETARY, 'pt_BR.UTF-8')


def emitir_certificados(pst: int, psa: int, nome: str, data: str, horas: int, participantes: list):
    mesext = {'01': 'JAN', '02': 'FEV', '03': 'MAR', '04': 'ABR', '05': 'MAI', '06': 'JUN',
              '07': 'JUL', '08': 'AGO', '09': 'SET', '10': 'OUT', '11': 'NOV', '12': 'DEZ'}
    dia, mes, ano = data.split('/')
    if pst != 0:
        nome = 'Primeiros Socorros - Terrestre'
        modelo = os.path.relpath(rf'C:\Users\{os.getlogin()}\PycharmProjects\AutomacaoCia\src\models\static\files\certificados\Treinamento.docx')
    else:
        if psa != 0:
            nome = 'Primeiros Socorros - Aquático'
            modelo = os.path.relpath(rf'C:\Users\{os.getlogin()}\PycharmProjects\AutomacaoCia\src\models\static\files\certificados\Treinamento.docx')
        else:
            modelo = os.path.relpath(rf'C:\Users\{os.getlogin()}\PycharmProjects\AutomacaoCia\src\models\static\files\certificados\TreinamentoGeral.docx')

    def extenso(datacompleta: str):
        diaex, mesex, anoex = datacompleta.split('/')
        mesextens = {'01': 'janeiro', '02': 'fevereiro', '03': 'março', '04': 'abril', '05': 'maio', '06': 'junho',
                     '07': 'julho', '08': 'agosto', '09': 'setembro', '10': 'outubro', '11': 'novembro',
                     '12': 'dezembro'}
        return f'{diaex} de {mesextens[mesex]} de {anoex}.'

    def exthoras(hr: int):
        horasext = {'1': 'uma', '2': 'duas', '3': 'três', '4': 'quatro', '5': 'cinco', '6': 'seis', '7': 'sete',
                    '8': 'oito', '9': 'nove', '10': 'dez', '11': 'onze', '12': 'doze', '13': 'treze', '14': 'quatorze',
                    '15': 'quinze'}
        return horasext[str(hr)]

    for item in participantes:
        doc = docx.Document(modelo)
        for parag in doc.paragraphs:
            if '#nome' in parag.text:
                inline = parag.runs
                for i in range(len(inline)):
                    if '#nome' in inline[i].text:
                        text = inline[i].text.replace('#nome', item.title())
                        inline[i].text = text
            if '#treinamento' in parag.text:
                inline = parag.runs
                for i in range(len(inline)):
                    if '#treinamento' in inline[i].text:
                        text = inline[i].text.replace('#treinamento', nome.title())
                        inline[i].text = text
            if '#data' in parag.text:
                inline = parag.runs
                for i in range(len(inline)):
                    if '#data' in inline[i].text:
                        text = inline[i].text.replace('#data', data)
                        inline[i].text = text
            if '#duracao' in parag.text:
                inline = parag.runs
                for i in range(len(inline)):
                    if '#duracao' in inline[i].text:
                        text = inline[i].text.replace('#duracao', str(horas))
                        inline[i].text = text
            if '#hrsexten' in parag.text:
                inline = parag.runs
                for i in range(len(inline)):
                    if '#hrsexten' in inline[i].text:
                        text = inline[i].text.replace('#hrsexten', exthoras(horas))
                        inline[i].text = text
            if '#extens' in parag.text:
                inline = parag.runs
                for i in range(len(inline)):
                    if '#extens' in inline[i].text:
                        text = inline[i].text.replace('#extens', extenso(data))
                        inline[i].text = text
        caminho = rf'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\13 - Treinamentos\{ano}\{mes} - {mesext[mes]}\{nome}\Certificados'
        if os.path.exists(caminho):
            doc.save(caminho + f'\\{item} - {nome}.docx')
            convert(caminho + f'\\{item} - {nome}.docx', caminho + f'\\{item} - {nome}.pdf')
            os.remove(caminho + f'\\{item} - {nome}.docx')
        else:
            os.makedirs(caminho)
            os.makedirs(caminho.replace('Certificados', 'Lista de presença'))
            os.makedirs(caminho.replace('Certificados', 'Material'))
            doc.save(caminho + f'\\{item} - {nome}.docx')
            convert(caminho + f'\\{item} - {nome}.docx', caminho + f'\\{item} - {nome}.pdf')
            os.remove(caminho + f'\\{item} - {nome}.docx')

    tkinter.messagebox.showinfo(
        title='Certificados ok!',
        message='Certificados gerados com sucesso!'
    )
