import pyautogui as pa
import time as t
from openpyxl import load_workbook as l_w
pa.hotkey('alt', 'tab'), pa.press('a'), t.sleep(2)
wb = l_w("Lancamentos.xlsx")

# # lançamento de faltas
sh = wb['Faltas']
x = 2
while x <= len(sh['A']):
    mat = str(sh[f'A{x}'].value)
    rub = str(sh[f'C{x}'].value)
    hr = str(sh[f'D{x}'].value)
    pa.write(mat), t.sleep(0.5), t.sleep(2), t.sleep(0.5), pa.press('enter', 2), t.sleep(0.5), t.sleep(1.5), t.sleep(0.5), pa.press('i'), t.sleep(0.5), t.sleep(0.5), t.sleep(0.5), pa.write(rub)
    t.sleep(0.5), t.sleep(0.5), pa.press('enter'), t.sleep(0.5), t.sleep(0.5), t.sleep(0.5), pa.write(hr), t.sleep(0.5), t.sleep(0.5), t.sleep(0.5), pa.press('enter', 65)
    x += 1

# deletar férias antigas
sh = wb['DeletarFerias']
x = 2
while x <= len(sh['A']):
    rub = ['1006', '1007', '1010', '1011', '1012', '1037']
    mat = str(sh[f'A{x}'].value)
    pa.write(mat), t.sleep(0.5), pa.press('enter', 2), t.sleep(1), pa.press('d')
    for r in rub:
        pa.write(r), t.sleep(0.8), pa.press('enter'), t.sleep(0.5), pa.press('left'), t.sleep(0.5), pa.press('enter')
    pa.press('enter', 2)
    x += 1

# lançamento de horistas
sh = wb['Horistas']
x = 2
while x <= len(sh['A']):
    mat = str(sh[f'A{x}'].value)
    rub = str(sh[f'C{x}'].value)
    hr = str(sh[f'D{x}'].value)
    obshr = str(sh[f'E{x}'].value)
    pa.write(mat), t.sleep(0.5), pa.press('enter', 2), t.sleep(2.3)
    try:
        pa.center(pa.locateOnScreen('dsr.png'))
        dsrlancado = True
    except:
        try:
            pa.center(pa.locateOnScreen('dsr2.png'))
            dsrlancado = True
        except:
            dsrlancado = False
    pa.press('a'), t.sleep(0.5), pa.write(rub)
    pa.press('enter'), t.sleep(0.5), pa.write(hr), t.sleep(0.5), pa.press('enter', 2)
    if dsrlancado:
        pass
    else:
        if obshr != 'HORA AULA ESTÁGIO 5.10':
            pa.press('i'), t.sleep(0.5), pa.write('27'), t.sleep(0.5)
            pa.press('enter', 3), t.sleep(0.5)
    pa.press('enter')
    x += 1

# lançamento de comissões
sh = wb['Comissoes']
x = 2
while x <= len(sh['A']):
    mat = str(sh[f'A{x}'].value)
    rub = str(sh[f'C{x}'].value)
    hr = str(sh[f'D{x}'].value)
    pa.write(mat), t.sleep(0.5), pa.press('enter', 2), t.sleep(0.5), pa.press('i'), t.sleep(0.5), pa.write(rub)
    pa.press('enter'), t.sleep(0.5), pa.write(hr), t.sleep(0.5), pa.press('enter')
    pa.press('enter'), t.sleep(0.5), pa.press('enter')
    x += 1

# # Lançamento de adiantamento
sh = wb['Adiantamento']
x = 2
while x <= len(sh['A']):
    mat = str(sh[f'A{x}'].value)
    rub = '81'
    hr = str(sh[f'D{x}'].value)
    pa.write(mat), t.sleep(0.5), pa.press('enter', 2), t.sleep(0.5), pa.press('i'), t.sleep(0.5), pa.write(rub)
    pa.press('enter'), t.sleep(0.5), pa.write(hr), t.sleep(0.5), pa.press('enter')
    pa.press('enter'), t.sleep(0.5), pa.press('enter')
    x += 1

# lançamento de desconto de VT
sh = wb['DescontoVT']
x = 2
while x <= len(sh['A']):
    mat = str(sh[f'A{x}'].value)
    rub = '80'
    hr = str(sh[f'D{x}'].value)
    pa.write(mat), t.sleep(0.5), pa.press('enter', 2), t.sleep(0.5), pa.press('i'), t.sleep(0.5), pa.write(rub)
    pa.press('enter'), t.sleep(0.5), pa.press('enter')
    pa.press('enter'), t.sleep(0.5), pa.press('enter')
    x += 1

# lançamento de plano de saúde
sh = wb['Plano']
x = 2
while x <= len(sh['A']):
    mat = str(sh[f'A{x}'].value)
    rub = str(sh[f'C{x}'].value)
    hr = str(sh[f'D{x}'].value)
    sq = str(sh[f'E{x}'].value)
    pa.write(mat), t.sleep(0.5), pa.press('enter', 2), t.sleep(0.5), pa.press('i'), t.sleep(0.5), pa.write(rub)
    pa.press('enter'), t.sleep(0.5), pa.write(sq), t.sleep(0.5), pa.press('enter'), t.sleep(0.5), pa.write(hr)
    pa.press('enter', 3), t.sleep(0.5),
    x += 1
