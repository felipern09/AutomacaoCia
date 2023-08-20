import datetime
import tkinter as tk
from src.controler.funcoes import salvar_holerites
from tkinter import ttk
from tkinter import *
from src.models.models import Colaborador, engine
from sqlalchemy.orm import sessionmaker


class MainApplication(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Salvar e Enviar Contracheques")
        self.geometry('361x250')
        self.img = PhotoImage(file='../models/static/imgs/Icone.png')
        self.iconphoto(False, self.img)
        self.columnconfigure(0, weight=5)
        self.rowconfigure(0, weight=5)
        for child in self.winfo_children():
            child.grid_configure(padx=1, pady=3)
        self.notebook = ttk.Notebook(self)
        self.Frame1 = Salvar(self.notebook)
        self.notebook.add(self.Frame1, text='Contracheques')
        self.notebook.pack()


class Salvar(ttk.Frame):
    def __init__(self, container):
        super().__init__()
        self.hoje = datetime.datetime.today()
        self.anos = list(range(2000, self.hoje.year+1))
        self.meses = list(range(1, 13))
        # aparecer dropdown com nomes da plan
        self.labelcomp = ttk.Label(self, width=60, text="Salvar contracheques nas pastas pessoais.")
        self.labelcomp.grid(column=1, row=1, padx=10, pady=2, sticky=W)
        # gerar folha da competencia selecionada
        self.botaogerar = ttk.Button(self, text="Salvar",
                                     command=lambda: [salvar_holerites()])
        self.botaogerar.grid(column=1, row=2, padx=230, pady=1, sticky=W)
        sessions = sessionmaker(bind=engine)
        session = sessions()
        self.grupo = []
        pessoas = session.query(Colaborador).filter_by(desligamento=None).all()
        for pess in pessoas:
            if pess.nome != '':
                self.grupo.append(pess.nome)
        self.nomes = list(sorted(set(filter(None, self.grupo))))
        # aparecer dropdown com nomes da plan
        self.labelem = ttk.Label(self, text='Enviar contracheques por e-mail:')
        self.labelem.grid(column=1, row=4, padx=10, pady=8, sticky=W)
        # Nome
        self.labelnome = ttk.Label(self, width=60, text='Nome:')
        self.labelnome.grid(column=1, row=5, padx=10, pady=2, sticky=W)
        self.combonome = ttk.Combobox(self, width=40, values=self.nomes)
        self.combonome.grid(column=1, row=5, padx=55, pady=2, sticky=W)
        # Ano
        self.labelano = ttk.Label(self, width=60, text='Ano - ')
        self.labelano.grid(column=1, row=6, padx=10, pady=2, sticky=W)
        self.labelde = ttk.Label(self, width=60, text='de:')
        self.labelde.grid(column=1, row=6, padx=45, pady=2, sticky=W)
        self.combodeano = ttk.Combobox(self, width=8, values=self.anos)
        self.combodeano.grid(column=1, row=6, padx=75, pady=2, sticky=W)

        self.labelate = ttk.Label(self, width=60, text='até:')
        self.labelate.grid(column=1, row=6, padx=165, pady=2, sticky=W)
        self.comboateano = ttk.Combobox(self, width=8, values=self.anos)
        self.comboateano.grid(column=1, row=6, padx=195, pady=2, sticky=W)

        # Meses
        self.labelmes = ttk.Label(self, width=60, text='Meses - ')
        self.labelmes.grid(column=1, row=7, padx=10, pady=2, sticky=W)
        self.labeldemes = ttk.Label(self, width=60, text='de:')
        self.labeldemes.grid(column=1, row=7, padx=65, pady=2, sticky=W)
        self.combodemes = ttk.Combobox(self, width=8, values=self.meses)
        self.combodemes.grid(column=1, row=7, padx=95, pady=2, sticky=W)

        self.labelatemes = ttk.Label(self, width=60, text='até:')
        self.labelatemes.grid(column=1, row=7, padx=185, pady=2, sticky=W)
        self.comboatemes = ttk.Combobox(self, width=8, values=self.meses)
        self.comboatemes.grid(column=1, row=7, padx=215, pady=2, sticky=W)

        # Enviar
        self.botaolancar = ttk.Button(self, width=20, text="Enviar",
                                      command=lambda: [])
        self.botaolancar.grid(column=1, row=8, padx=190, pady=1, sticky=W)


if __name__ == '__main__':
    app = MainApplication()
    app.mainloop()
