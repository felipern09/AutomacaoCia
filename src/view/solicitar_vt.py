from src.controler.f_vt import incluir_vt, retirar_vt, gerar_vt
import tkinter as tk
from tkinter import ttk
from tkinter import *
from src.models.models import Colaborador, engine
from sqlalchemy.orm import sessionmaker


class MainApplication(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Vale transporte - Cia BSB")
        self.geometry('480x200')
        self.img = PhotoImage(file='../models/static/imgs/Icone.png')
        self.iconphoto(False, self.img)
        self.columnconfigure(0, weight=5)
        self.rowconfigure(0, weight=5)
        for child in self.winfo_children():
            child.grid_configure(padx=1, pady=3)
        self.notebook = ttk.Notebook(self)

        self.Frame1 = CadastrarVale(self.notebook)
        self.Frame2 = RetirarVale(self.notebook)
        self.Frame3 = GerarVT(self.notebook)

        self.notebook.add(self.Frame1, text='Incluir na lista de VT')
        self.notebook.add(self.Frame2, text='Retirar da lista de VT')
        self.notebook.add(self.Frame3, text='Gerar VT')

        self.notebook.pack()


class CadastrarVale(ttk.Frame):
    def __init__(self, container):
        super().__init__()
        sessions = sessionmaker(bind=engine)
        session = sessions()
        self.grupo = []
        pessoas = session.query(Colaborador).filter_by(desligamento=None).filter(Colaborador.ag.isnot(None)).filter(Colaborador.ag.isnot('None')).all()
        for pess in pessoas:
            if pess.nome != '':
                self.grupo.append(pess.nome)
        pessoas2 = session.query(Colaborador).filter_by(desligamento='None').filter(Colaborador.ag.isnot(None)).filter(Colaborador.ag.isnot('None')).all()
        for pess in pessoas2:
            if pess.nome != '':
                self.grupo.append(pess.nome)
        self.nomes = list(sorted(set(filter(None, self.grupo))))
        self.nome1 = StringVar()
        self.labelnome = ttk.Label(self, width=20, text="Nome")
        self.labelnome.grid(column=1, row=1, padx=25, pady=15, sticky=W)
        self.combonome1 = ttk.Combobox(self, values=self.nomes, textvariable=self.nome1, width=50)
        self.combonome1.grid(column=1, row=2, padx=25, pady=1, sticky=W)
        # checkbutton para indicar onde foi feito
        self.vt = IntVar()
        self.editar = ttk.Radiobutton(self, text='BRB', value=1, variable=self.vt)
        self.editar.grid(column=1, row=26, padx=26, pady=1, sticky=W)
        self.onde = ttk.Radiobutton(self, text='Valecard', value=2, variable=self.vt)
        self.onde.grid(column=1, row=27, padx=26, pady=1, sticky=W)
        self.onde = ttk.Radiobutton(self, text='Goiás', value=3, variable=self.vt)
        self.onde.grid(column=1, row=28, padx=26, pady=1, sticky=W)
        self.botaocadastrar = ttk.Button(self, text='Registrar pessoa na lista de VT',
                                         command=lambda: [
                                             incluir_vt(self.combonome1.get(), self.vt.get())])
        self.botaocadastrar.grid(column=1, row=29, padx=290, pady=1, sticky=W)


class RetirarVale(ttk.Frame):
    def __init__(self, container):
        super().__init__()
        sessions = sessionmaker(bind=engine)
        session = sessions()
        self.grupo = []
        pessoas = session.query(Colaborador).filter_by(desligamento=None).filter(Colaborador.ag.isnot(None)).filter(Colaborador.ag.isnot('None')).all()
        for pess in pessoas:
            if pess.nome != '':
                self.grupo.append(pess.nome)
        pessoas2 = session.query(Colaborador).filter_by(desligamento='None').filter(Colaborador.ag.isnot(None)).filter(Colaborador.ag.isnot('None')).all()
        for pess in pessoas2:
            if pess.nome != '':
                self.grupo.append(pess.nome)
        self.nomes = list(sorted(set(filter(None, self.grupo))))
        self.nome1 = StringVar()
        self.labelnome = ttk.Label(self, width=20, text="Nome")
        self.labelnome.grid(column=1, row=1, padx=25, pady=15, sticky=W)
        self.combonome1 = ttk.Combobox(self, values=self.nomes, textvariable=self.nome1, width=50)
        self.combonome1.grid(column=1, row=2, padx=25, pady=1, sticky=W)
        # checkbutton para indicar qual VT
        self.vt = IntVar()
        self.editar = ttk.Radiobutton(self, text='BRB', value=1, variable=self.vt)
        self.editar.grid(column=1, row=26, padx=26, pady=1, sticky=W)
        self.onde = ttk.Radiobutton(self, text='Valecard', value=2, variable=self.vt)
        self.onde.grid(column=1, row=27, padx=26, pady=1, sticky=W)
        self.onde = ttk.Radiobutton(self, text='Goiás', value=3, variable=self.vt)
        self.onde.grid(column=1, row=28, padx=26, pady=1, sticky=W)
        self.botaocadastrar = ttk.Button(self, text='Retirar pessoa na lista de VT',
                                         command=lambda: [retirar_vt(self.combonome1.get(), self.vt.get())])
        self.botaocadastrar.grid(column=1, row=29, padx=290, pady=1, sticky=W)


class GerarVT(ttk.Frame):
    def __init__(self, container):
        super().__init__()
        self.tipos = ['BRB', 'Valecard', 'Goiás']
        self.tipovt = StringVar()
        self.labelnome = ttk.Label(self, text="Gerar pedido de VT para o tipo:")
        self.labelnome.grid(column=1, row=1, padx=25, pady=15, sticky=W)
        self.combonome1 = ttk.Combobox(self, values=self.tipos, textvariable=self.tipovt, width=50)
        self.combonome1.grid(column=1, row=2, padx=25, pady=1, sticky=W)
        self.botaocadastrar = ttk.Button(self, text='Gerar pedido de VT',
                                         command=lambda: [gerar_vt(self.tipovt.get())])
        self.botaocadastrar.grid(column=1, row=29, padx=290, pady=15, sticky=W)


if __name__ == '__main__':
    app = MainApplication()
    app.mainloop()
