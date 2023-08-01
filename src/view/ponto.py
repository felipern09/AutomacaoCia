# ponto principal -> gerar relatórios de pontos em pdf de acorodo com a data escolhida
# ponto completo -> gerar relatórios de ponto em excel a partir do xlsx geral do programa secullum
# ponto estagiários -> gerar somente relatórios de estagiários para envio para líder (de acordo com data selecionada)
#  relatorio de atrasos -> atraves do relatorio geral salvo em xlsx gerar comparatvo de horários registrados
# com horarios de cadastro e informar se houve atraso superior a 10 min para cada registro
# cadastro de funcionário no programa secullum
from src.controler.funcoes import gerar_relatorios_ponto_pdf
import tkinter as tk
from tkcalendar import DateEntry
from datetime import datetime
import tkinter.filedialog
from openpyxl import load_workbook as l_w
import tkinter.filedialog
from tkinter import ttk
from tkinter import *

# Under development.


class MainApplication(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Relatórios de Ponto - Cia BSB")
        self.geometry('661x350')
        self.img = PhotoImage(file='../models/static/imgs/Icone.png')
        self.iconphoto(False, self.img)
        self.columnconfigure(0, weight=5)
        self.rowconfigure(0, weight=5)
        for child in self.winfo_children():
            child.grid_configure(padx=1, pady=3)
        self.notebook = ttk.Notebook(self)

        self.Frame1 = Frame1(self.notebook)
        self.Frame2 = Frame2(self.notebook)

        self.notebook.add(self.Frame1, text='Gerar relação de pontos')
        self.notebook.add(self.Frame2, text='Gerar conferência de horários')

        self.notebook.pack()


class Frame1(ttk.Frame):
    def __init__(self, container):
        super().__init__()

        self.hoje = datetime.today()
        self.geral = StringVar()
        self.caminho = StringVar()
        self.labelescolh = ttk.Label(self, width=40, text="Escolha o arquivo AFD: ")
        self.labelescolh.grid(column=1, row=1, padx=25, pady=1, sticky=W)
        self.botaoescolha = ttk.Button(self, text="Selecionar AFD", command=self.selecionar_funcionario)
        self.botaoescolha.grid(column=1, row=1, padx=350, pady=1, sticky=W)
        self.nome = StringVar()
        self.horario = StringVar()
        self.cargo = StringVar()
        self.departamento = StringVar()
        self.tipocontr = StringVar()
        self.nomesplan = []
        # definir data inicial
        self.labelinicial = ttk.Label(self, width=20, text="Data inicial:")
        self.labelinicial.grid(column=1, row=12, padx=25, pady=1, sticky=W)
        self.entryinicial = DateEntry(self, selectmode='day', year=self.hoje.year, month=self.hoje.month,
                                     day=self.hoje.day, locale='pt_BR')
        self.entryinicial.grid(column=1, row=12, padx=125, pady=1, sticky=W)
        # definir data final
        self.labelfinal = ttk.Label(self, width=20, text="Data final:")
        self.labelfinal.grid(column=1, row=13, padx=25, pady=1, sticky=W)
        self.entryfinal = DateEntry(self, selectmode='day', year=self.hoje.year, month=self.hoje.month,
                                     day=self.hoje.day, locale='pt_BR')
        self.entryfinal.grid(column=1, row=13, padx=125, pady=1, sticky=W)

        self.estag = IntVar()
        self.editar = ttk.Checkbutton(self, text='Apenas estagiários.', variable=self.estag)
        self.editar.grid(column=1, row=26, padx=26, pady=1, sticky=W)

        def carregarfunc(local):
            planwb = l_w(local)
            plansh = planwb['Respostas ao formulário 1']
            lista = []
            for x, pessoa in enumerate(plansh):
                lista.append(f'{x + 1} - {pessoa[2].value}')

        self.botaocadastrar = ttk.Button(self, width=20, text="Gerar Relatórios",
                                         command=lambda: [
                                             gerar_relatorios_ponto_pdf(self.caminho.get(),
                                                                        self.entryinicial.get(),
                                                                        self.entryfinal.get())
                                         ])
        self.botaocadastrar.grid(column=1, row=28, padx=520, pady=1, sticky=W)

    def selecionar_funcionario(self):
        try:
            caminhoplan = tkinter.filedialog.askopenfilename(title='Arquivo AFD')
            self.caminho.set(str(caminhoplan))
        except ValueError:
            pass


class Frame2(ttk.Frame):
    def __init__(self, container):
        super().__init__()
        self.hoje = datetime.today()
        self.geral = StringVar()
        self.caminho = StringVar()
        self.labelescolh = ttk.Label(self, width=40, text="Escolha o arquivo AFD: ")
        self.labelescolh.grid(column=1, row=1, padx=25, pady=1, sticky=W)
        self.botaoescolha = ttk.Button(self, text="Selecionar AFD", command=self.selecionar_funcionario)
        self.botaoescolha.grid(column=1, row=1, padx=350, pady=1, sticky=W)
        self.nome = StringVar()
        self.horario = StringVar()
        self.cargo = StringVar()
        self.departamento = StringVar()
        self.tipocontr = StringVar()
        self.nomesplan = []
        # definir data inicial
        self.labelinicial = ttk.Label(self, width=20, text="Data inicial:")
        self.labelinicial.grid(column=1, row=12, padx=25, pady=1, sticky=W)
        self.entryinicial = DateEntry(self, selectmode='day', year=self.hoje.year, month=self.hoje.month,
                                      day=self.hoje.day, locale='pt_BR')
        self.entryinicial.grid(column=1, row=12, padx=125, pady=1, sticky=W)
        # definir data final
        self.labelfinal = ttk.Label(self, width=20, text="Data final:")
        self.labelfinal.grid(column=1, row=13, padx=25, pady=1, sticky=W)
        self.entryfinal = DateEntry(self, selectmode='day', year=self.hoje.year, month=self.hoje.month,
                                    day=self.hoje.day, locale='pt_BR')
        self.entryfinal.grid(column=1, row=13, padx=125, pady=1, sticky=W)

        self.estag = IntVar()
        self.editar = ttk.Checkbutton(self, text='Apenas estagiários.', variable=self.estag)
        self.editar.grid(column=1, row=26, padx=26, pady=1, sticky=W)

        def carregarfunc(local):
            planwb = l_w(local)
            plansh = planwb['Respostas ao formulário 1']
            lista = []
            for x, pessoa in enumerate(plansh):
                lista.append(f'{x + 1} - {pessoa[2].value}')

        self.botaocadastrar = ttk.Button(self, width=20, text="Gerar Relatórios",
                                         command=lambda: [])
        self.botaocadastrar.grid(column=1, row=28, padx=520, pady=1, sticky=W)

    def selecionar_funcionario(self):
        try:
            caminhoplan = tkinter.filedialog.askopenfilename(title='Planilha Funcionários')
            self.caminho.set(str(caminhoplan))
        except ValueError:
            pass


if __name__ == '__main__':
    app = MainApplication()
    app.mainloop()
