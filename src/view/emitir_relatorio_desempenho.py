# Under development


# import os
# import time as t
# import pyautogui as pa
# import pyscreenshot as ps
# from openpyxl import load_workbook as l_w
#
#
# # code to save images of performance evaluation of employee
# wb = l_w("../models/static/files/Resultados Individuais AV.xlsm", read_only=False)
# sh = wb['Nota Geral']
# pa.hotkey('alt', 'tab')
# t.sleep(0.5)
# linha = 106
# click = 265
# while sh[f"K{linha}"].value is not None:
#     pa.click(1402, click), t.sleep(0.2)
#     nome = str(sh[f"K{linha}"].value)
#     pasta = f'\\\Qnapcia\\rh\\01 - RH\\01 - Administração.Controles\\02 - Funcionários, Departamentos e Férias\\' \
#             f'000 - Pastas Funcionais\\00 - ATIVOS\\{nome}\\Diversos\\AV Desemp {nome}.jpg'
#
#     def main():
#         image = ps.grab(bbox=(38, 228, 1042, 770))
#         image.save(f'AV Desemp {nome}.jpg', 'jpeg')
#     if __name__ == '__main__':
#         main()
#     os.rename(f'AV Desemp {nome}.jpg', pasta)
#     linha = linha + 1
#     click = click + 26
#     if click > 981:
#         break
