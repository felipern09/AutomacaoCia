import pyautogui as pa
from openpyxl import load_workbook as l_w
import time as t
import pyperclip as pp

wb = l_w('PlanPonto.xlsx')
sh = wb['Planilha1']
x = 2
while x <= 93:
    matricula = str(sh[f'A{x}'].value)
    nome = str(sh[f'B{x}'].value)
    pis = str(sh[f'C{x}'].value)
    horario = '1'
    funcao = str(sh[f'D{x}'].value)
    depto = str(sh[f'E{x}'].value)
    admiss = str(sh[f'F{x}'].value)
    # # clicar incluir
    pa.click(-1507,147), t.sleep(3)
    pa.write(matricula), pa.press('tab')
    pp.copy(nome), pa.hotkey('ctrl','v'), pa.press('tab')
    pa.write(pis), pa.press('tab', 4)
    pa.write(horario), pa.press('tab')
    pp.copy(funcao), pa.hotkey('ctrl','v'), pa.press('tab')
    pp.copy(depto), pa.hotkey('ctrl','v'), pa.press('tab')
    pa.write(admiss)
    # click concluir
    pa.click(-1492,617)
    t.sleep(2)
    x += 1