import tkinter as tk
from src.controler.f_folha import confirma_grade, lancar_folha_no_dexion
from tkinter import ttk
from tkinter import *
import sys
import os


class MainApplication(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Cia BSB - Atividades RH")
        self.geometry('600x450')
        self.img = PhotoImage(file='../models/static/imgs/Icone.png')
        self.iconphoto(False, self.img)
        self.columnconfigure(0, weight=5)
        self.rowconfigure(0, weight=5)
        for child in self.winfo_children():
            child.grid_configure(padx=1, pady=3)
        self.notebook = ttk.Notebook(self)

        self.Frame1 = Frame1(self.notebook)

        self.notebook.add(self.Frame1, text='Atalhos')

        self.notebook.pack()


def admiss():
    os.system('admissao_e_desligamento.py')


def altera_folha():
    os.system('alteracoes_folha.py')


def cert():
    os.system('certificados.py')


def contatos():
    os.system('contatos_diversos.py')


def contracheques():
    os.system('emitir_contrachques.py')


def lanca_folha():
    os.system('gerar_lancar_folha.py')


def rel_ponto():
    os.system('ponto.py')


def atest():
    os.system('registrar_atestados.py')


def adiant():
    os.system('solicitar_adiatementos.py')


def pgto():
    os.system('solicitar_pagamento.py')


def vt():
    os.system('solicitar_vt.py')


def unif():
    os.system('uniformes.py')


class Frame1(ttk.Frame):
    def __init__(self, container):
        super().__init__()
        # gerar folha da competencia selecionada
        self.botaoadmiss = ttk.Button(self, text="\n  Admissão e Desligamento  \n",
                                     command=lambda: [admiss()])
        self.botaoadmiss.grid(column=1, row=1, padx=20, pady=20, sticky=W)
        self.botaofolha = ttk.Button(self, width=20, text="\n  Alterações Folha  \n",
                                     command=lambda: [altera_folha()])
        self.botaofolha.grid(column=2, row=1, padx=20, pady=20, sticky=W)
        self.botaogerar = ttk.Button(self, width=20, text="\n  Gerar e Lançar Folha  \n",
                                     command=lambda: [lanca_folha()])
        self.botaogerar.grid(column=3, row=1, padx=20, pady=20, sticky=W)
        self.botaogerar = ttk.Button(self, width=20, text="\n  Emitir Holerites  \n",
                                     command=lambda: [contracheques()])
        self.botaogerar.grid(column=1, row=2, padx=20, pady=20, sticky=W)

        self.botaogerar = ttk.Button(self, width=20, text="\n  Emitir Rel. Ponto  \n",
                                     command=lambda: [rel_ponto()])
        self.botaogerar.grid(column=2, row=2, padx=20, pady=20, sticky=W)

        self.botaogerar = ttk.Button(self, width=20, text="\n  Solicitar Pagamento  \n",
                                     command=lambda: [pgto()])
        self.botaogerar.grid(column=3, row=2, padx=20, pady=20, sticky=W)
        self.botaogerar = ttk.Button(self, width=20, text="\n  Certificados  \n",
                                     command=lambda: [cert()])
        self.botaogerar.grid(column=1, row=3, padx=20, pady=20, sticky=W)
        self.botaogerar = ttk.Button(self, width=20, text="\n  Contatos diversos  \n",
                                     command=lambda: [contatos()])
        self.botaogerar.grid(column=2, row=3, padx=20, pady=20, sticky=W)
        self.botaogerar = ttk.Button(self, width=20, text="\n  Atestados  \n",
                                     command=lambda: [atest()])
        self.botaogerar.grid(column=3, row=3, padx=20, pady=20, sticky=W)
        self.botaogerar = ttk.Button(self, width=20, text="\n  Adiantamentos  \n",
                                     command=lambda: [adiant()])
        self.botaogerar.grid(column=1, row=4, padx=20, pady=20, sticky=W)
        self.botaogerar = ttk.Button(self, width=20, text="\n  VT  \n",
                                     command=lambda: [vt()])
        self.botaogerar.grid(column=2, row=4, padx=20, pady=20, sticky=W)
        self.botaogerar = ttk.Button(self, width=20, text="\n  Uniformes  \n",
                                     command=lambda: [unif()])
        self.botaogerar.grid(column=3, row=4, padx=20, pady=20, sticky=W)






if __name__ == '__main__':
    app = MainApplication()
    app.mainloop()
