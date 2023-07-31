import tkinter as tk
from tkcalendar import DateEntry
from datetime import datetime
import tkinter.filedialog
from openpyxl import load_workbook as l_w
import tkinter.filedialog
from tkinter import ttk
from tkinter import *


class MainApplication(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Cálculo de folha - Cia BSB")
        self.geometry('361x250')
        self.img = PhotoImage(file='../models/static/imgs/Icone.png')
        self.iconphoto(False, self.img)
        self.columnconfigure(0, weight=5)
        self.rowconfigure(0, weight=5)
        for child in self.winfo_children():
            child.grid_configure(padx=1, pady=3)
        self.notebook = ttk.Notebook(self)

        self.Frame1 = Frame1(self.notebook)
        self.Frame2 = Frame2(self.notebook)

        self.notebook.add(self.Frame1, text='Gerar Folha')
        self.notebook.add(self.Frame2, text='Lançar Folha no Dexion')

        self.notebook.pack()


class Frame1(ttk.Frame):
    def __init__(self, container):
        super().__init__()
        self.competencia = IntVar()
        self.competencias = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]
        # aparecer dropdown com nomes da plan
        self.labelcomp = ttk.Label(self, width=60, text="Escolha a competência da folha: ")
        self.labelcomp.grid(column=1, row=1, padx=25, pady=50, sticky=W)
        self.combocomp = ttk.Combobox(self, values=self.competencias, textvariable=self.competencia, width=15)
        self.combocomp.grid(column=1, row=1, padx=205, pady=50, sticky=W)
        # gerar folha da competencia selecionada
        self.botaogerar = ttk.Button(self, width=20, text="Gerar folha",
                                     command=lambda: [])
        self.botaogerar.grid(column=1, row=3, padx=190, pady=1, sticky=W)


class Frame2(ttk.Frame):
    def __init__(self, container):
        super().__init__()
        self.competencia = IntVar()
        self.competencias = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]
        # aparecer dropdown com nomes da plan
        self.labelcomp = ttk.Label(self, width=60, text="Escolha a competência da folha: ")
        self.labelcomp.grid(column=1, row=1, padx=25, pady=50, sticky=W)
        self.combocomp = ttk.Combobox(self, values=self.competencias, textvariable=self.competencia, width=15)
        self.combocomp.grid(column=1, row=1, padx=205, pady=50, sticky=W)
        # gerar folha da competencia selecionada
        self.botaolancar = ttk.Button(self, width=20, text="Lançar folha no Dexion",
                                      command=lambda: [])
        self.botaolancar.grid(column=1, row=3, padx=190, pady=1, sticky=W)


if __name__ == '__main__':
    app = MainApplication()
    app.mainloop()
