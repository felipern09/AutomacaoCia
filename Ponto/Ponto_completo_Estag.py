from openpyxl import load_workbook as l_w
from openpyxl.styles import NamedStyle
import pandas as pd
from datetime import datetime as dt
import win32com.client as client
from PIL import ImageGrab
import os

# Ler arquivo txt dos registror rejeitados
geral = pd.read_csv(
    r'C:\Users\Felipe Rodrigues\Desktop\Relatorios Ponto\Rej - 16-03-2023.txt', sep=' ', header=None,
    encoding='iso8859-1'
)
geral = geral.rename(columns={0: 'matricula', 16: 'data', 17: 'dia', 18: 'hora'})
geral = geral[geral.dia != 'Batida']
geral = geral[geral.matricula < 9999]
mat = []
mat_unicas = []
for matricula in geral['matricula']:
    mat.append(matricula)
    mat_unicas = list(set(mat))
for matr in mat_unicas:
    geral2 = geral[geral.matricula == matr]
    geral2 = geral2.set_index('matricula')
    geral2 = geral2.dropna(axis='columns')
    # print(geral2)
    geral2 = geral2.drop(geral2.iloc[:, [2, 3, 4]], axis=1)
    geral2['dia'] = pd.to_datetime(geral2['dia'], format='%d/%m/%Y')
    geral2['dia'] = geral2['dia'].apply(lambda y: dt.strftime(y, '%d/%m/%Y'))
    geral2['hora'] = geral2['hora'].apply(lambda x: f'{x}:00')
    geral2['hora'] = pd.to_timedelta(geral2['hora'])
    # geral2 = geral2.set_index('dia')
    # geral2 = geral2.groupby('dia')
    geral2.to_excel(f'../Ponto/xls/Ponto Estágio - {matr}.xlsx')
    wb = l_w(f'../Ponto/xls/Ponto Estágio - {matr}.xlsx')
    sh = wb['Sheet1']
    sh['A1'].value = 'Matrícula'
    sh['B1'].value = 'Data'
    sh['C1'].value = 'Entrada 1'
    sh['D1'].value = 'Saída 1'
    sh['E1'].value = 'Entrada 2'
    sh['F1'].value = 'Saída 2'
    sh['G1'].value = 'Entrada 3'
    sh['H1'].value = 'Saída 3'
    x = 2
    for row in sh:
        if sh[f'B{x}'].value == sh[f'B{x-1}'].value:
            if sh[f'D{x-1}'].value is None:
                sh[f'D{x - 1}'].value = sh[f'C{x}'].value
                sh.delete_rows(x, 1)
            else:
                if sh[f'E{x - 1}'].value is None:
                    sh[f'E{x - 1}'].value = sh[f'C{x}'].value
                    sh.delete_rows(x, 1)
                else:
                    if sh[f'F{x - 1}'].value is None:
                        sh[f'F{x - 1}'].value = sh[f'C{x}'].value
                        sh.delete_rows(x, 1)
                    else:
                        if sh[f'G{x - 1}'].value is None:
                            sh[f'G{x - 1}'].value = sh[f'C{x}'].value
                            sh.delete_rows(x, 1)
                        else:
                            if sh[f'H{x - 1}'].value is None:
                                sh[f'H{x - 1}'].value = sh[f'C{x}'].value
                                sh.delete_rows(x, 1)
    for row in sh:
        if sh[f'B{x}'].value == sh[f'B{x - 1}'].value:
            if sh[f'D{x - 1}'].value is None:
                sh[f'D{x - 1}'].value = sh[f'C{x}'].value
                sh.delete_rows(x, 1)
            else:
                if sh[f'E{x - 1}'].value is None:
                    sh[f'E{x - 1}'].value = sh[f'C{x}'].value
                    sh.delete_rows(x, 1)
                else:
                    if sh[f'F{x - 1}'].value is None:
                        sh[f'F{x - 1}'].value = sh[f'C{x}'].value
                        sh.delete_rows(x, 1)
                    else:
                        if sh[f'G{x - 1}'].value is None:
                            sh[f'G{x - 1}'].value = sh[f'C{x}'].value
                            sh.delete_rows(x, 1)
                        else:
                            if sh[f'H{x - 1}'].value is None:
                                sh[f'H{x - 1}'].value = sh[f'C{x}'].value
                                sh.delete_rows(x, 1)
        for row in sh:
            if sh[f'B{x}'].value == sh[f'B{x - 1}'].value:
                if sh[f'D{x - 1}'].value is None:
                    sh[f'D{x - 1}'].value = sh[f'C{x}'].value
                    sh.delete_rows(x, 1)
                else:
                    if sh[f'E{x - 1}'].value is None:
                        sh[f'E{x - 1}'].value = sh[f'C{x}'].value
                        sh.delete_rows(x, 1)
                    else:
                        if sh[f'F{x - 1}'].value is None:
                            sh[f'F{x - 1}'].value = sh[f'C{x}'].value
                            sh.delete_rows(x, 1)
                        else:
                            if sh[f'G{x - 1}'].value is None:
                                sh[f'G{x - 1}'].value = sh[f'C{x}'].value
                                sh.delete_rows(x, 1)
                            else:
                                if sh[f'H{x - 1}'].value is None:
                                    sh[f'H{x - 1}'].value = sh[f'C{x}'].value
                                    sh.delete_rows(x, 1)
        x += 1

    estilo_data = NamedStyle(name='data', number_format='DD/MM/YYYY')
    estilo_hora = NamedStyle(name='hora', number_format='HH:MM:SS')
    for cell in sh['B']:
        sh[f'B{int(cell.row)}'].style = estilo_data
    for item in sh['C']:
        sh[f'C{int(item.row)}'].style = estilo_hora
    for item in sh['D']:
        sh[f'D{int(item.row)}'].style = estilo_hora
    for item in sh['E']:
        sh[f'E{int(item.row)}'].style = estilo_hora
    for item in sh['F']:
        sh[f'F{int(item.row)}'].style = estilo_hora
    for item in sh['G']:
        sh[f'G{int(item.row)}'].style = estilo_hora
    for item in sh['H']:
        sh[f'H{int(item.row)}'].style = estilo_hora
    sh.column_dimensions['A'].width = 11
    sh.column_dimensions['B'].width = 11
    sh.column_dimensions['C'].width = 11
    sh.column_dimensions['D'].width = 11
    sh.column_dimensions['E'].width = 11
    sh.column_dimensions['F'].width = 11
    sh.column_dimensions['G'].width = 11
    sh.column_dimensions['H'].width = 11
    wb2 = l_w(f'../Ponto/xls/zzBase.xlsx')
    ws2 = wb2.active
    for row in ws2.rows:
        for cell in row:
            if cell.value == matr:
                nome = ws2.cell(row=cell.row, column=3).value
                linha = int((len(sh['A']) / 2) + 1)
                wb.save(f'../Ponto/xls/Ponto Estágio - {nome}.xlsx')
                os.remove(f'../Ponto/xls/Ponto Estágio - {matr}.xlsx')
                # t.sleep(2)
                # inserir código de envio de email aqui
                excel = client.Dispatch('Excel.Application')
                plan = excel.Workbooks.Open(
                    r'C:\Users\Felipe Rodrigues\PycharmProjects\AutomacaoCia\Ponto\xls\Ponto Estágio - {}.xlsx'
                    .format(nome)
                )
                folha = plan.Sheets['Sheet1']
                excel.visible = 0
                copyrange = folha.Range(f'B1:F{linha}')
                copyrange.CopyPicture(Format=2)
                ImageGrab.grabclipboard().save(f'Ponto {str(nome).title().split(" ")[0]}.png')
                excel.Quit()
                html_body = f'''
                <p>Olá {str(nome).title().split(" ")[0]},</p>
                <p> Segue em anexo seu relatório de ponto do mês Fevereiro/2023.</p><br>
                Atenciosamente,<br>
                Felipe Rodrigues,<br>
                '''
                outlook = client.Dispatch('Outlook.Application')
                message = outlook.CreateItem(0)
                message.To = 'felipe.rodrigs09@gmail.com'
                message.Subject = 'Relatório de Ponto'
                message.HTMLBody = html_body
                message.Attachments.Add(
                    f'C:\\Users\\Felipe Rodrigues\\PycharmProjects\\AutomacaoCia\\Ponto\\'
                    f'Ponto {str(nome).title().split(" ")[0]}.png'
                )
                message.Send()
# # procurar na planilha base.xlsx nome e-mail matricula no ponto e horarios de entrada e saida
