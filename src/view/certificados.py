import tkinter as tk
from tkcalendar import DateEntry
from datetime import datetime
from tkinter import ttk
from tkinter import *


class MainApplication(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Certificados - Cia BSB")
        self.geometry('661x150')
        self.img = PhotoImage(file='../models/static/imgs/Icone.png')
        self.iconphoto(False, self.img)
        self.columnconfigure(0, weight=5)
        self.rowconfigure(0, weight=5)
        for child in self.winfo_children():
            child.grid_configure(padx=1, pady=3)
        self.notebook = ttk.Notebook(self)
        self.Frame1 = Frame1(self.notebook)
        self.notebook.add(self.Frame1, text='Emissão de Certificados')
        self.notebook.pack()


class Frame1(ttk.Frame):
    def __init__(self, container):
        super().__init__()
        self.hoje = datetime.today()
        self.horas = IntVar()
        self.hrs = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15]
        # definir nome do treinamento
        self.labelnome = ttk.Label(self, width=120, text="Digite o nome do treinamento:")
        self.labelnome.grid(column=1, row=10, padx=25, pady=1, sticky=W)
        self.entrymatr = ttk.Entry(self, width=100)
        self.entrymatr.grid(column=1, row=11, padx=25, pady=1, sticky=W)
        # definir data do trinamento
        self.labeladmiss = ttk.Label(self, width=60, text="Data do treinamento:")
        self.labeladmiss.grid(column=1, row=12, padx=25, pady=1, sticky=W)
        self.entryadmiss = DateEntry(self, selectmode='day', year=self.hoje.year, month=self.hoje.month,
                                     day=self.hoje.day, locale='pt_BR')
        self.entryadmiss.grid(column=1, row=12, padx=165, pady=1, sticky=W)
        # definir horas de duração
        self.labeldurac = ttk.Label(self, width=60, text='Duração em horas:')
        self.labeldurac.grid(column=1, row=13, padx=25, pady=1, sticky=W)
        self.combodur = ttk.Combobox(self, width=12, textvariable=self.horas, values=self.hrs)
        self.combodur.grid(column=1, row=13, padx=165, pady=1, sticky=W)
        # selecionar funcionário que participou
        self.botaocadastrar = ttk.Button(self, width=20, text="Emitir certificados", command=lambda: [])
        self.botaocadastrar.grid(column=1, row=28, padx=520, pady=1, sticky=W)


if __name__ == '__main__':
    app = MainApplication()
    app.mainloop()
