# ponto principal -> gerar relatórios de pontos em pdf de acorodo com a data escolhida
# ponto completo -> gerar relatórios de ponto em excel a partir do xlsx geral do programa secullum
# ponto estagiários -> gerar somente relatórios de estagiários para envio para líder (de acordo com data selecionada)
#  relatorio de atrasos -> atraves do relatorio geral salvo em xlsx gerar comparatvo de horários registrados
# com horarios de cadastro e informar se houve atraso superior a 10 min para cada registro
# cadastro de funcionário no programa secullum
from src.controler.f_ponto import gerar_relatorios_ponto_pdf, cadastrar_no_ponto
import tkinter as tk
from tkcalendar import DateEntry
from datetime import datetime
import tkinter.filedialog
from openpyxl import load_workbook as l_w
import tkinter.filedialog
from tkinter import ttk
from tkinter import *
from src.models.models import Colaborador, engine
from sqlalchemy.orm import sessionmaker

# Under development.


class MainApplication(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Relatórios de Ponto - Cia BSB")
        self.geometry('432x280')
        self.img = PhotoImage(file='../models/static/imgs/Icone.png')
        self.iconphoto(False, self.img)
        self.columnconfigure(0, weight=5)
        self.rowconfigure(0, weight=5)
        for child in self.winfo_children():
            child.grid_configure(padx=1, pady=3)
        self.notebook = ttk.Notebook(self)

        self.Frame1 = GerarRelatorio(self.notebook)
        self.Frame2 = ConferirAtrasos(self.notebook)
        self.Frame3 = CadastrarNoRelogio(self.notebook)

        self.notebook.add(self.Frame1, text='Gerar relação de pontos')
        self.notebook.add(self.Frame2, text='Gerar conferência de horários')
        self.notebook.add(self.Frame3, text='Cadastrar no Relógio')

        self.notebook.pack()


class GerarRelatorio(ttk.Frame):
    def __init__(self, container):
        super().__init__()

        self.hoje = datetime.today()
        self.geral = StringVar()
        self.caminho = StringVar()
        self.labelescolh = ttk.Label(self, width=40, text="Escolha o arquivo AFD: ")
        self.labelescolh.grid(column=1, row=1, padx=25, pady=1, sticky=W)
        self.botaoescolha = ttk.Button(self, text="Selecionar AFD", command=self.selecionar_funcionario)
        self.botaoescolha.grid(column=1, row=1, padx=165, pady=1, sticky=W)
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
        self.botaocadastrar = ttk.Button(self, width=20, text="Gerar Relatórios",
                                         command=lambda: [
                                             gerar_relatorios_ponto_pdf(self.caminho.get(),
                                                                        self.entryinicial.get(),
                                                                        self.entryfinal.get(),
                                                                        self.estag.get())
                                         ])
        self.botaocadastrar.grid(column=1, row=28, padx=165, pady=1, sticky=W)

    def selecionar_funcionario(self):
        try:
            caminhoplan = tkinter.filedialog.askopenfilename(title='Arquivo AFD')
            self.caminho.set(str(caminhoplan))
        except ValueError:
            pass


class ConferirAtrasos(ttk.Frame):
    def __init__(self, container):
        super().__init__()
        self.hoje = datetime.today()
        self.geral = StringVar()
        self.caminho = StringVar()
        self.labelescolh = ttk.Label(self, width=40, text="Escolha o arquivo AFD: ")
        self.labelescolh.grid(column=1, row=1, padx=25, pady=1, sticky=W)
        self.botaoescolha = ttk.Button(self, text="Selecionar AFD", command=self.selecionar_funcionario)
        self.botaoescolha.grid(column=1, row=1, padx=165, pady=1, sticky=W)
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
        self.botaocadastrar.grid(column=1, row=28, padx=165, pady=1, sticky=W)

    def selecionar_funcionario(self):
        try:
            caminhoplan = tkinter.filedialog.askopenfilename(title='Planilha Funcionários')
            self.caminho.set(str(caminhoplan))
        except ValueError:
            pass


class CadastrarNoRelogio(ttk.Frame):
    def __init__(self, container):
        super().__init__()
        sessions = sessionmaker(bind=engine)
        session = sessions()
        self.grupo = []
        pessoas = session.query(Colaborador).filter_by(desligamento=None).filter(Colaborador.ag.isnot(None)).filter(Colaborador.ag.isnot('None')).all()
        for pess in pessoas:
            if pess.nome != '':
                self.grupo.append(pess.nome)
        self.nomes = list(sorted(set(filter(None, self.grupo))))
        self.hoje = datetime.today()
        self.geral = StringVar()
        self.caminho = StringVar()
        self.labelescolh = ttk.Label(self, width=90, text='Abra o programa Gerenciado iDx Class e escolha o nome do colaborador')
        self.labelescolh.grid(column=1, row=1, padx=25, pady=1, sticky=W)
        self.nome = StringVar()
        self.horario = StringVar()
        self.cargo = StringVar()
        self.departamento = StringVar()
        self.tipocontr = StringVar()
        self.nomesplan = []
        # definir data inicial
        self.labelinicial = ttk.Label(self, width=20, text="Nome:")
        self.labelinicial.grid(column=1, row=12, padx=25, pady=5, sticky=W)
        self.combonome = ttk.Combobox(self, width=45, values=self.nomes)
        self.combonome.grid(column=1, row=12, padx=80, pady=5, sticky=W)
        self.labelmatricula = ttk.Label(self, width=20, text="Matrícula: ")
        self.labelmatricula.grid(column=1, row=13, padx=25, pady=5, sticky=W)
        self.labelpis = ttk.Label(self, width=20, text="Pis: ")
        self.labelpis.grid(column=1, row=14, padx=25, pady=5, sticky=W)
        self.labelchk = ttk.Label(self, width=20, text="")
        self.labelchk.grid(column=1, row=15, padx=25, pady=5, sticky=W)
        self.entrymatr = ttk.Entry(self, width=30)

        def mostrar_pis_mat(event):
            nome = event.widget.get()
            pessoa = session.query(Colaborador).filter_by(nome=nome).order_by(Colaborador.matricula.desc()).first()
            if 'ESTAG' in pessoa.cargo:
                self.labelmatricula.config(text=f"Matrícula: {pessoa.matricula}")
            else:
                self.labelmatricula.config(text=f"Matrícula: {pessoa.matricula}")
                self.labelpis.config(text=f"Pis: {pessoa.pis}")

        self.combonome.bind("<<ComboboxSelected>>", mostrar_pis_mat)
        # checkbutton para indicar onde foi feito

        def chkbt():
            self.labelchk.config(text='Matricula no ponto: ')
            self.entrymatr.grid_configure(column=1, row=14, padx=25, pady=5, sticky=W)

        self.alterarmtr = IntVar()
        self.alterar = ttk.Checkbutton(self, text='Alterar Matr ponto', variable=self.alterarmtr, command=chkbt)
        self.alterar.grid(column=1, row=26, padx=226, pady=1, sticky=W)
        self.botaocadastrar = ttk.Button(self, text="Cadastrar no Ponto",
                                         command=lambda: [cadastrar_no_ponto(self.combonome.get(),
                                                                             self.alterarmtr.get(),
                                                                             self.entrymatr.get())])
        self.botaocadastrar.grid(column=1, row=28, padx=255, pady=8, sticky=W)


# implementar pesquisa do funcionário direto no banco em vez da planilha base
# 	criar campo 'ponto' no db para salvar a matrícula do ponto eletrônico

if __name__ == '__main__':
    app = MainApplication()
    app.mainloop()
