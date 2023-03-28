import pyautogui as pa
import time as t
from openpyxl import load_workbook as l_w
pa.hotkey('alt', 'tab'), pa.press('a'), t.sleep(2)
wb = l_w("Lancamentos.xlsx")

# # lançamento de faltas
sh = wb['Faltas']
x = 2
while x <= 5:
    mat = str(sh[f'A{x}'].value)
    rub = str(sh[f'C{x}'].value)
    hr = str(sh[f'D{x}'].value)
    pa.write(mat), t.sleep(2), pa.press('enter', 2), t.sleep(1.5), pa.press('i'), t.sleep(0.5), pa.write(rub)
    t.sleep(0.5), pa.press('enter'), t.sleep(0.5), pa.write(hr), t.sleep(0.5), pa.press('enter', 63)
    x += 1

# deletar férias antigas
sh = wb['DeletarFerias']
x = 2
while x <= 11:
    rub = ['1006', '1007', '1010', '1011', '1012', '1037']
    mat = str(sh[f'A{x}'].value)
    pa.write(mat), pa.press('enter', 2), pa.press('d')
    for r in rub:
        pa.write(r), pa.press('enter'), pa.press('left'), pa.press('enter')
    pa.press('enter', 2)
    x += 1
#


# lançamento de horistas
sh = wb['Horistas']
x = 2
while x <= 113:

    mat = str(sh[f'A{x}'].value)
    rub = str(sh[f'C{x}'].value)
    hr = str(sh[f'D{x}'].value)
    pa.write(mat), pa.press('enter', 2), pa.press('a'), pa.write(rub)
    pa.press('enter'), pa.write(hr), pa.press('enter')
    pa.press('enter'), pa.press('enter')
    x += 1

# lançamento de comissões
sh = wb['Comissoes']
x = 2
while x <= 15:
    mat = str(sh[f'A{x}'].value)
    rub = str(sh[f'C{x}'].value)
    hr = str(sh[f'D{x}'].value)
    pa.write(mat), pa.press('enter', 2), pa.press('i'), pa.write(rub)
    pa.press('enter'), pa.write(hr), pa.press('enter')
    pa.press('enter'), pa.press('enter')
    x += 1

# lançamento de DSR
sh = wb['DSR']
x = 2
while x <= 62:
    mat = str(sh[f'A{x}'].value)
    rub = '27'
    pa.write(mat), pa.press('enter', 2), pa.press('i'), pa.write(rub)
    pa.press('enter', 2), pa.press('enter', 2)
    x += 1

# # Lançamento de adiantamento
sh = wb['Adiantamento']
x = 2
while x <= 15:
    mat = str(sh[f'A{x}'].value)
    rub = '81'
    hr = str(sh[f'D{x}'].value)
    pa.write(mat), pa.press('enter', 2), pa.press('i'), pa.write(rub)
    pa.press('enter'), pa.write(hr), pa.press('enter')
    pa.press('enter'), pa.press('enter')
    x += 1

# lançamento de desconto de VT
sh = wb['DescontoVT']
x = 2
while x <= 30:
    mat = str(sh[f'A{x}'].value)
    rub = '80'
    hr = str(sh[f'D{x}'].value)
    pa.write(mat), pa.press('enter', 2), pa.press('i'), pa.write(rub)
    pa.press('enter'), pa.press('enter')
    pa.press('enter'), pa.press('enter')
    x += 1

# lançamento de plano de saúde
sh = wb['Plano']
x = 2
while x <= 5:
    mat = str(sh[f'A{x}'].value)
    rub = str(sh[f'C{x}'].value)
    hr = str(sh[f'D{x}'].value)
    sq = str(sh[f'E{x}'].value)
    pa.write(mat), pa.press('enter', 2), pa.press('i'), pa.write(rub)
    pa.press('enter'), pa.write(sq), pa.press('enter'), pa.write(hr)
    pa.press('enter', 3),
    x += 1
