from openpyxl import load_workbook as l_w
import pyautogui as pa
import time as t
wb = l_w('Contracheque.xlsx')
sh = wb['Planilha1']
x=2

while x <= 53:
    competencia = str(sh[f'A{x}'].value).replace('/','')
    pagamento = str(sh[f'B{x}'].value).replace('/','')
    de = str(sh[f'C{x}'].value)
    ate = str(sh[f'D{x}'].value)
    caminho = f'C:\\Users\\RH\\PycharmProjects\\AutomacaoCia\\Emissao de contracheques\\Contracheque {competencia}.pdf'
    pa.click(-816,515), t.sleep(0.5), pa.write(competencia), pa.press('tab'), pa.write(pagamento), pa.press('tab'), pa.write(de), pa.press('tab'), pa.write(ate)
    pa.click(-787,731), t.sleep(4),
    pa.hotkey('ctrl', 's'), pa.write(caminho), t.sleep(0.5), pa.press('enter'), t.sleep(0.5), pa.click(-33,132)
    x += 1
