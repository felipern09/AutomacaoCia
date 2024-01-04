import datetime
import tkinter as tk
from src.controler.f_folha import confirma_grade, lancar_folha_no_dexion, previa_folha
from tkcalendar import DateEntry
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
        self.Frame3 = Frame3(self.notebook)

        self.notebook.add(self.Frame1, text='Gerar Folha')
        self.notebook.add(self.Frame2, text='Lançar Folha no Dexion')
        self.notebook.add(self.Frame3, text='Enviar Prévia da Folha')

        self.notebook.pack()


class Frame1(ttk.Frame):
    def __init__(self, container):
        super().__init__()
        self.hoje = datetime.datetime.today()
        # aparecer dropdown com nomes da plan
        self.labelcomp = ttk.Label(self, width=60, text="Escolha a competência da folha: ")
        self.labelcomp.grid(column=1, row=1, padx=25, pady=50, sticky=W)
        self.entrydtfolha = DateEntry(self, selectmode='day', year=self.hoje.year, month=self.hoje.month,
                                        day=self.hoje.day, locale='pt_BR')
        self.entrydtfolha.grid(column=1, row=1, padx=205, pady=50, sticky=W)
        # gerar folha da competencia selecionada
        self.botaogerar = ttk.Button(self, width=20, text="Gerar folha",
                                     command=lambda: [confirma_grade(self.entrydtfolha.get())])
        self.botaogerar.grid(column=1, row=3, padx=190, pady=1, sticky=W)


class Frame2(ttk.Frame):
    def __init__(self, container):
        super().__init__()
        self.competencia = IntVar()
        self.competencias = list(range(1, 13))
        # aparecer dropdown com nomes da plan
        self.labelcomp = ttk.Label(self, width=60, text="Escolha a competência da folha: ")
        self.labelcomp.grid(column=1, row=1, padx=25, pady=50, sticky=W)
        self.combocomp = ttk.Combobox(self, values=self.competencias, textvariable=self.competencia, width=15)
        self.combocomp.grid(column=1, row=1, padx=205, pady=50, sticky=W)
        # gerar folha da competencia selecionada
        self.botaolancar = ttk.Button(self, width=20, text="Lançar folha no Dexion",
                                      command=lambda: [lancar_folha_no_dexion(self.competencia.get())])
        self.botaolancar.grid(column=1, row=3, padx=190, pady=1, sticky=W)


class Frame3(ttk.Frame):
    def __init__(self, container):
        super().__init__()
        self.competencia = IntVar()
        self.competencias = list(range(1, 13))
        # aparecer dropdown com nomes da plan
        self.labelcomp = ttk.Label(self, width=60, text='Para enviar prévia da folha siga os passos:')
        self.labelcomp.grid(column=1, row=1, padx=25, pady=20, sticky=W)
        self.label1 = ttk.Label(self, width=60, text='1º - Gere os contracheques da competência no Dexion.')
        self.label1.grid(column=1, row=2, padx=25, pady=2, sticky=W)
        self.label2 = ttk.Label(self, width=60, text='2º - Clique em enviar prévias.')
        self.label2.grid(column=1, row=3, padx=25, pady=2, sticky=W)
        # gerar folha da competencia selecionada
        self.botaolancar = ttk.Button(self, width=20, text='Enviar prévias da folha', command=lambda: [previa_folha()])
        self.botaolancar.grid(column=1, row=4, padx=190, pady=1, sticky=W)


if __name__ == '__main__':
    app = MainApplication()
    app.mainloop()
