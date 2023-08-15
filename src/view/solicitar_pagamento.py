import tkinter as tk
from tkcalendar import DateEntry
from tkinter import ttk
from datetime import datetime
from tkinter import *
from src.models.models import Colaborador, engine
from sqlalchemy.orm import sessionmaker
from src.controler.funcoes import confirmar_pagamento, gerar_planilha_pgto_itau


class MainApplication(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Solicitar pagamento")
        self.geometry('661x440')
        self.img = PhotoImage(file='../models/static/imgs/Icone.png')
        self.iconphoto(False, self.img)
        self.notebook = ttk.Notebook(self)
        self.Frame1 = Pgto(self.notebook)
        self.Frame2 = PgtoPorAqr(self.notebook)
        self.notebook.add(self.Frame1, text='Gerar Pedido Pgto')
        self.notebook.add(self.Frame2, text='Gerar Pedido Por Arquivo')
        self.notebook.pack()


class Pgto(ttk.Frame):
    def __init__(self, container):
        super().__init__()
        self.tipos = ['Salário', 'Férias', 'Vale Transporte', 'Vale Alimentação', 'Comissão',
                      '13º salário', 'Bolsa Estágio', 'Bônus', 'Adiantamento Salarial',
                      'Rescisão', 'Bolsa Auxílio', 'Pensão Alimentícia', 'Pgto em C/C',
                      'Remuneração']
        sessions = sessionmaker(bind=engine)
        session = sessions()
        self.grupo = []
        pessoas = session.query(Colaborador).filter_by(desligamento=None).all()
        for pess in pessoas:
            if pess.nome != '':
                self.grupo.append(pess.nome)
        pessoas2 = session.query(Colaborador).filter_by(desligamento='None').all()
        for pess in pessoas2:
            if pess.nome != '':
                self.grupo.append(pess.nome)
        self.nomes = list(sorted(set(filter(None, self.grupo))))
        self.hoje = datetime.today()
        self.nome1 = StringVar()
        self.nome2 = StringVar()
        self.nome3 = StringVar()
        self.nome4 = StringVar()
        self.nome5 = StringVar()
        self.nome6 = StringVar()
        self.nome7 = StringVar()
        self.nome8 = StringVar()
        self.nome9 = StringVar()
        self.nome10 = StringVar()
        self.nome11 = StringVar()
        self.nome12 = StringVar()
        self.nome13 = StringVar()
        self.nome14 = StringVar()
        self.nome15 = StringVar()
        self.tipo1 = StringVar()
        self.tipo2 = StringVar()
        self.tipo3 = StringVar()
        self.tipo4 = StringVar()
        self.tipo5 = StringVar()
        self.tipo6 = StringVar()
        self.tipo7 = StringVar()
        self.tipo8 = StringVar()
        self.tipo9 = StringVar()
        self.tipo10 = StringVar()
        self.tipo11 = StringVar()
        self.tipo12 = StringVar()
        self.tipo13 = StringVar()
        self.tipo14 = StringVar()
        self.tipo15 = StringVar()
        self.cargo = StringVar()
        self.departamento = StringVar()
        self.tipocontr = StringVar()
        self.nomesplan = []
        self.labelnome = ttk.Label(self, width=20, text="Nome")
        self.labelnome.grid(column=1, row=1, padx=25, pady=1, sticky=W)
        self.combonome1 = ttk.Combobox(self, values=self.nomes, textvariable=self.nome1, width=50)
        self.combonome1.grid(column=1, row=2, padx=25, pady=1, sticky=W)
        self.combonome2 = ttk.Combobox(self, values=self.nomes, textvariable=self.nome2, width=50)
        self.combonome2.grid(column=1, row=3, padx=25, pady=1, sticky=W)
        self.combonome3 = ttk.Combobox(self, values=self.nomes, textvariable=self.nome3, width=50)
        self.combonome3.grid(column=1, row=4, padx=25, pady=1, sticky=W)
        self.combonome4 = ttk.Combobox(self, values=self.nomes, textvariable=self.nome4, width=50)
        self.combonome4.grid(column=1, row=5, padx=25, pady=1, sticky=W)
        self.combonome5 = ttk.Combobox(self, values=self.nomes, textvariable=self.nome5, width=50)
        self.combonome5.grid(column=1, row=6, padx=25, pady=1, sticky=W)
        self.combonome6 = ttk.Combobox(self, values=self.nomes, textvariable=self.nome6, width=50)
        self.combonome6.grid(column=1, row=7, padx=25, pady=1, sticky=W)
        self.combonome7 = ttk.Combobox(self, values=self.nomes, textvariable=self.nome7, width=50)
        self.combonome7.grid(column=1, row=8, padx=25, pady=1, sticky=W)
        self.combonome8 = ttk.Combobox(self, values=self.nomes, textvariable=self.nome8, width=50)
        self.combonome8.grid(column=1, row=9, padx=25, pady=1, sticky=W)
        self.combonome9 = ttk.Combobox(self, values=self.nomes, textvariable=self.nome9, width=50)
        self.combonome9.grid(column=1, row=10, padx=25, pady=1, sticky=W)
        self.combonome10 = ttk.Combobox(self, values=self.nomes, textvariable=self.nome10, width=50)
        self.combonome10.grid(column=1, row=11, padx=25, pady=1, sticky=W)
        self.combonome11 = ttk.Combobox(self, values=self.nomes, textvariable=self.nome11, width=50)
        self.combonome11.grid(column=1, row=12, padx=25, pady=1, sticky=W)
        self.combonome12 = ttk.Combobox(self, values=self.nomes, textvariable=self.nome12, width=50)
        self.combonome12.grid(column=1, row=13, padx=25, pady=1, sticky=W)
        self.combonome13 = ttk.Combobox(self, values=self.nomes, textvariable=self.nome13, width=50)
        self.combonome13.grid(column=1, row=14, padx=25, pady=1, sticky=W)
        self.combonome14 = ttk.Combobox(self, values=self.nomes, textvariable=self.nome14, width=50)
        self.combonome14.grid(column=1, row=15, padx=25, pady=1, sticky=W)
        self.combonome15 = ttk.Combobox(self, values=self.nomes, textvariable=self.nome15, width=50)
        self.combonome15.grid(column=1, row=16, padx=25, pady=1, sticky=W)
        # tipo
        self.labeltipo = ttk.Label(self, width=20, text="Tipo")
        self.labeltipo.grid(column=1, row=1, padx=350, pady=1, sticky=W)
        self.combotipo1 = ttk.Combobox(self, values=self.tipos, textvariable=self.tipo1, width=25)
        self.combotipo1.grid(column=1, row=2, padx=350, pady=1, sticky=W)
        self.combotipo2 = ttk.Combobox(self, values=self.tipos, textvariable=self.tipo2, width=25)
        self.combotipo2.grid(column=1, row=3, padx=350, pady=1, sticky=W)
        self.combotipo3 = ttk.Combobox(self, values=self.tipos, textvariable=self.tipo3, width=25)
        self.combotipo3.grid(column=1, row=4, padx=350, pady=1, sticky=W)
        self.combotipo4 = ttk.Combobox(self, values=self.tipos, textvariable=self.tipo4, width=25)
        self.combotipo4.grid(column=1, row=5, padx=350, pady=1, sticky=W)
        self.combotipo5 = ttk.Combobox(self, values=self.tipos, textvariable=self.tipo5, width=25)
        self.combotipo5.grid(column=1, row=6, padx=350, pady=1, sticky=W)
        self.combotipo6 = ttk.Combobox(self, values=self.tipos, textvariable=self.tipo6, width=25)
        self.combotipo6.grid(column=1, row=7, padx=350, pady=1, sticky=W)
        self.combotipo7 = ttk.Combobox(self, values=self.tipos, textvariable=self.tipo7, width=25)
        self.combotipo7.grid(column=1, row=8, padx=350, pady=1, sticky=W)
        self.combotipo8 = ttk.Combobox(self, values=self.tipos, textvariable=self.tipo8, width=25)
        self.combotipo8.grid(column=1, row=9, padx=350, pady=1, sticky=W)
        self.combotipo9 = ttk.Combobox(self, values=self.tipos, textvariable=self.tipo9, width=25)
        self.combotipo9.grid(column=1, row=10, padx=350, pady=1, sticky=W)
        self.combotipo10 = ttk.Combobox(self, values=self.tipos, textvariable=self.tipo10, width=25)
        self.combotipo10.grid(column=1, row=11, padx=350, pady=1, sticky=W)
        self.combotipo11 = ttk.Combobox(self, values=self.tipos, textvariable=self.tipo11, width=25)
        self.combotipo11.grid(column=1, row=12, padx=350, pady=1, sticky=W)
        self.combotipo12 = ttk.Combobox(self, values=self.tipos, textvariable=self.tipo12, width=25)
        self.combotipo12.grid(column=1, row=13, padx=350, pady=1, sticky=W)
        self.combotipo13 = ttk.Combobox(self, values=self.tipos, textvariable=self.tipo13, width=25)
        self.combotipo13.grid(column=1, row=14, padx=350, pady=1, sticky=W)
        self.combotipo14 = ttk.Combobox(self, values=self.tipos, textvariable=self.tipo14, width=25)
        self.combotipo14.grid(column=1, row=15, padx=350, pady=1, sticky=W)
        self.combotipo15 = ttk.Combobox(self, values=self.tipos, textvariable=self.tipo15, width=25)
        self.combotipo15.grid(column=1, row=16, padx=350, pady=1, sticky=W)
        # valor
        self.labelvalor = ttk.Label(self, width=20, text="Valor")
        self.labelvalor.grid(column=1, row=1, padx=525, pady=1, sticky=W)
        self.entryvalor1 = ttk.Entry(self, width=20)
        self.entryvalor1.grid(column=1, row=2, padx=525, pady=1, sticky=W)
        self.entryvalor2 = ttk.Entry(self, width=20)
        self.entryvalor2.grid(column=1, row=3, padx=525, pady=1, sticky=W)
        self.entryvalor3 = ttk.Entry(self, width=20)
        self.entryvalor3.grid(column=1, row=4, padx=525, pady=1, sticky=W)
        self.entryvalor4 = ttk.Entry(self, width=20)
        self.entryvalor4.grid(column=1, row=5, padx=525, pady=1, sticky=W)
        self.entryvalor5 = ttk.Entry(self, width=20)
        self.entryvalor5.grid(column=1, row=6, padx=525, pady=1, sticky=W)
        self.entryvalor6 = ttk.Entry(self, width=20)
        self.entryvalor6.grid(column=1, row=7, padx=525, pady=1, sticky=W)
        self.entryvalor7 = ttk.Entry(self, width=20)
        self.entryvalor7.grid(column=1, row=8, padx=525, pady=1, sticky=W)
        self.entryvalor8 = ttk.Entry(self, width=20)
        self.entryvalor8.grid(column=1, row=9, padx=525, pady=1, sticky=W)
        self.entryvalor9 = ttk.Entry(self, width=20)
        self.entryvalor9.grid(column=1, row=10, padx=525, pady=1, sticky=W)
        self.entryvalor10 = ttk.Entry(self, width=20)
        self.entryvalor10.grid(column=1, row=11, padx=525, pady=1, sticky=W)
        self.entryvalor11 = ttk.Entry(self, width=20)
        self.entryvalor11.grid(column=1, row=12, padx=525, pady=1, sticky=W)
        self.entryvalor12 = ttk.Entry(self, width=20)
        self.entryvalor12.grid(column=1, row=13, padx=525, pady=1, sticky=W)
        self.entryvalor13 = ttk.Entry(self, width=20)
        self.entryvalor13.grid(column=1, row=14, padx=525, pady=1, sticky=W)
        self.entryvalor14 = ttk.Entry(self, width=20)
        self.entryvalor14.grid(column=1, row=15, padx=525, pady=1, sticky=W)
        self.entryvalor15 = ttk.Entry(self, width=20)
        self.entryvalor15.grid(column=1, row=16, padx=525, pady=1, sticky=W)
        self.data = DateEntry(self, selectmode='day', year=self.hoje.year, month=self.hoje.month, day=self.hoje.day, locale='pt_BR')
        self.data.grid(column=1, row=28, padx=25, pady=1, sticky=W)
        self.botao = ttk.Button(self, text="Criar Planilha",command=lambda: [gerar_planilha_pgto_itau(
            self.nome1.get(), self.nome2.get(), self.nome3.get(), self.nome4.get(), self.nome5.get(), self.nome6.get(), self.nome7.get(), self.nome8.get(),
            self.nome9.get(),self.nome10.get(), self.nome11.get(), self.nome12.get(), self.nome13.get(), self.nome14.get(), self.nome15.get(), self.tipo1.get(),
            self.tipo2.get(), self.tipo3.get(), self.tipo4.get(), self.tipo5.get(), self.tipo6.get(), self.tipo7.get(), self.tipo8.get(), self.tipo9.get(),
            self.tipo10.get(),self.tipo11.get(), self.tipo12.get(), self.tipo13.get(), self.tipo14.get(), self.tipo15.get(), self.entryvalor1.get(),
            self.entryvalor2.get(),self.entryvalor3.get(), self.entryvalor4.get(), self.entryvalor5.get(),
            self.entryvalor6.get(),self.entryvalor7.get(), self.entryvalor8.get(), self.entryvalor9.get(),
            self.entryvalor10.get(),self.entryvalor11.get(), self.entryvalor12.get(), self.entryvalor13.get(),
            self.entryvalor14.get(),self.entryvalor15.get(), self.data.get()
        )])
        self.botao.grid(column=1, row=28, padx=430, pady=1, sticky=W)
        self.botao = ttk.Button(self, text="Cria capa e envia e-mail",
                                command=lambda: [confirmar_pagamento(self.tipo1.get(),
            self.tipo2.get(), self.tipo3.get(), self.tipo4.get(), self.tipo5.get(), self.tipo6.get(), self.tipo7.get(), self.tipo8.get(), self.tipo9.get(),
            self.tipo10.get(),self.tipo11.get(), self.tipo12.get(), self.tipo13.get(), self.tipo14.get(), self.tipo15.get(), self.entryvalor1.get(),
            self.entryvalor2.get(),self.entryvalor3.get(), self.entryvalor4.get(), self.entryvalor5.get(),
            self.entryvalor6.get(),self.entryvalor7.get(), self.entryvalor8.get(), self.entryvalor9.get(),
            self.entryvalor10.get(),self.entryvalor11.get(), self.entryvalor12.get(), self.entryvalor13.get(),
            self.entryvalor14.get(),self.entryvalor15.get(), self.data.get())])
        self.botao.grid(column=1, row=28, padx=515, pady=1, sticky=W)


class PgtoPorAqr(ttk.Frame):
    def __init__(self, container):
        super().__init__()
        self.labelnome = ttk.Label(self, width=20, text="Escolha o arquivo")
        self.labelnome.grid(column=1, row=1, padx=25, pady=1, sticky=W)
        self.botao = ttk.Button(self, text="Gerar Pedido",command=lambda: [])
        self.botao.grid(column=1, row=28, padx=430, pady=1, sticky=W)


if __name__ == '__main__':
    app = MainApplication()
    app.mainloop()
