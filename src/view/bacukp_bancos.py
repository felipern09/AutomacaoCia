import tkinter as tk
from datetime import datetime
import tkinter.filedialog
from src.controler.f_folha import backup_bancos
from tkinter import ttk
from tkinter import *


class MainApplication(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Backup Banco de Dados")
        self.geometry('480x180')
        self.img = PhotoImage(file='../models/static/imgs/Icone.png')
        self.iconphoto(False, self.img)
        self.columnconfigure(0, weight=5)
        self.rowconfigure(0, weight=5)
        for child in self.winfo_children():
            child.grid_configure(padx=1, pady=3)
        self.notebook = ttk.Notebook(self)

        self.Frame1 = BkUp(self.notebook)

        self.notebook.add(self.Frame1, text='Backup')

        self.notebook.pack()


class BkUp(ttk.Frame):
    def __init__(self, container):
        super().__init__()

        self.hoje = datetime.today()
        self.geral = StringVar()
        self.caminhoapp = StringVar()
        self.caminhoaut = StringVar()
        self.labelapp = ttk.Label(self, width=40, text="Escolher arquivo appCia.")
        self.labelapp.grid(column=1, row=1, padx=25, pady=1, sticky=W)
        self.botaoapp = ttk.Button(self, text="Escolha o arquivo", command=self.selecionar_appcia)
        self.botaoapp.grid(column=1, row=1, padx=350, pady=1, sticky=W)
        self.labelaut = ttk.Label(self, width=40, text="Escolher arquivo AutumoaçãoCia.")
        self.labelaut.grid(column=1, row=2, padx=25, pady=1, sticky=W)
        self.botaoaut = ttk.Button(self, text="Escolha o arquivo", command=self.selecionar_automacao)
        self.botaoaut.grid(column=1, row=2, padx=350, pady=1, sticky=W)
        self.botaobaixar = ttk.Button(self, width=20, text="AppCia -> Automação",
                                      command=lambda: [backup_bancos(self.caminhoapp.get(), self.caminhoaut.get(), 1)])
        self.botaobaixar.grid(column=1, row=3, padx=320, pady=40, sticky=W)
        self.botaoupload = ttk.Button(self, width=20, text="Automação -> AppCia",
                                      command=lambda: [backup_bancos(self.caminhoapp.get(), self.caminhoaut.get(), 2)])
        self.botaoupload.grid(column=1, row=3, padx=140, pady=40, sticky=W)

    def selecionar_appcia(self):
        try:
            caminhoplan = tkinter.filedialog.askopenfilename(title='Arquivo DB appCia')
            self.caminhoapp.set(str(caminhoplan))
        except ValueError:
            pass

    def selecionar_automacao(self):
        try:
            caminhoplan = tkinter.filedialog.askopenfilename(title='Arquivo DB AutomaçãoCia')
            self.caminhoaut.set(str(caminhoplan))
        except ValueError:
            pass


if __name__ == '__main__':
    app = MainApplication()
    app.mainloop()
