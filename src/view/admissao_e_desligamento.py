import tkinter as tk
from tkcalendar import DateEntry
from datetime import datetime
import tkinter.filedialog
from src.controler.funcoes import cadastro_funcionario, salvar_docs_funcionarios, enviar_emails_funcionario, \
    cadastro_estagiario, cadastrar_autonomo, validar_pis, desligar_pessoa
from openpyxl import load_workbook as l_w
from src.models.listas import horarios, cargos, departamentos, tipodecontrato
import tkinter.filedialog
from tkinter import ttk
from tkinter import *
from src.models.models import Colaborador, engine
from sqlalchemy.orm import sessionmaker


class MainApplication(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Atividades DP - Cia BSB")
        self.geometry('661x550')
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
        self.Frame4 = Frame4(self.notebook)
        self.Frame5 = Frame5(self.notebook)

        self.notebook.add(self.Frame1, text='Cadastrar Funcionário')
        self.notebook.add(self.Frame2, text='Cadastrar Estagiário')
        self.notebook.add(self.Frame3, text='Cadastrar Autônomo')
        self.notebook.add(self.Frame4, text='Desligamento')
        self.notebook.add(self.Frame5, text='Atualizar Banco de Dados')

        self.notebook.pack()


class Frame1(ttk.Frame):
    def __init__(self, container):
        super().__init__()

        self.hoje = datetime.today()
        self.geral = StringVar()
        self.caminho = StringVar()
        self.labelescolh = ttk.Label(self, width=40, text="Escolher planilha de novos funcionários")
        self.labelescolh.grid(column=1, row=1, padx=25, pady=1, sticky=W)
        self.botaoescolha = ttk.Button(self, text="Escolha a planilha", command=self.selecionar_funcionario)
        self.botaoescolha.grid(column=1, row=1, padx=350, pady=1, sticky=W)
        self.nome = StringVar()
        self.horario = StringVar()
        self.cargo = StringVar()
        self.departamento = StringVar()
        self.tipocontr = StringVar()
        self.nomesplan = []
        # aparecer dropdown com nomes da plan
        self.labelnome = ttk.Label(self, width=20, text="Nome:")
        self.labelnome.grid(column=1, row=10, padx=25, pady=1, sticky=W)
        self.combonome = ttk.Combobox(self, values=self.nomesplan, textvariable=self.nome, width=50)
        self.combonome.grid(column=1, row=10, padx=125, pady=1, sticky=W)
        # aparecer entry para preencher matricula
        self.labelmatr = ttk.Label(self, width=20, text="Matrícula:")
        self.labelmatr.grid(column=1, row=11, padx=25, pady=1, sticky=W)
        self.entrymatr = ttk.Entry(self, width=20)
        self.entrymatr.grid(column=1, row=11, padx=125, pady=1, sticky=W)
        # aparecer entry para preencher admissao
        self.labeladmiss = ttk.Label(self, width=20, text="Admissão:")
        self.labeladmiss.grid(column=1, row=12, padx=25, pady=1, sticky=W)
        self.entryadmiss = DateEntry(self, selectmode='day', year=self.hoje.year, month=self.hoje.month,
                                     day=self.hoje.day, locale='pt_BR')
        self.entryadmiss.grid(column=1, row=12, padx=125, pady=1, sticky=W)
        # aparecer horario preenchido e dropdown para escolher horario
        self.labelhor = ttk.Label(self, width=55, text="Horário preenchido: ")
        self.labelhor.grid(column=1, row=14, padx=25, pady=1, sticky=W)
        self.combohor = ttk.Combobox(self, values=horarios, textvariable=self.horario, width=50)
        self.combohor.grid(column=1, row=15, padx=125, pady=1, sticky=W)
        # aparecer entry para preencher salario
        self.labelsal = ttk.Label(self, width=20, text="Salário:")
        self.labelsal.grid(column=1, row=16, padx=25, pady=1, sticky=W)
        self.entrysal = ttk.Entry(self, width=20)
        self.entrysal.grid(column=1, row=16, padx=125, pady=1, sticky=W)
        # aparecer dropdown para escolher cargo
        self.labelcargo = ttk.Label(self, width=20, text="Cargo")
        self.labelcargo.grid(column=1, row=18, padx=25, pady=1, sticky=W)
        self.combocargo = ttk.Combobox(self, values=cargos, textvariable=self.cargo, width=50)
        self.combocargo.grid(column=1, row=18, padx=125, pady=1, sticky=W)
        # aparecer dropdown para escolher depto
        self.labeldepto = ttk.Label(self, width=20, text="Departamento:")
        self.labeldepto.grid(column=1, row=19, padx=25, pady=1, sticky=W)
        self.combodepto = ttk.Combobox(self, values=departamentos, textvariable=self.departamento, width=50)
        self.combodepto.grid(column=1, row=19, padx=125, pady=1, sticky=W)
        # aparecer dropdown para escolher tipo_contr
        self.labelcontr = ttk.Label(self, width=20, text="Tipo de contrato:")
        self.labelcontr.grid(column=1, row=21, padx=25, pady=1, sticky=W)
        self.combocontr = ttk.Combobox(self, values=tipodecontrato, textvariable=self.tipocontr, width=50)
        self.combocontr.grid(column=1, row=21, padx=125, pady=1, sticky=W)
        self.hrs = StringVar()
        self.hrm = StringVar()
        # aparecer entry para preencher hrsem
        self.labelhrsem = ttk.Label(self, width=20, text="Hrs Sem.:")
        self.labelhrsem.grid(column=1, row=24, padx=25, pady=1, sticky=W)
        self.entryhrsem = ttk.Entry(self, width=20, textvariable=self.hrs)
        self.entryhrsem.grid(column=1, row=24, padx=125, pady=1, sticky=W)
        # aparecer entry para preencher hrmens
        self.labelhrmen = ttk.Label(self, width=20, text="Hrs Mens.:")
        self.labelhrmen.grid(column=1, row=25, padx=25, pady=1, sticky=W)
        self.entryhrmen = ttk.Entry(self, width=20, textvariable=self.hrm)
        self.entryhrmen.grid(column=1, row=25, padx=125, pady=1, sticky=W)
        self.agencia = StringVar()
        self.conta = StringVar()
        self.digito = StringVar()
        # aparecer entry para agencia
        self.labelag = ttk.Label(self, width=20, text="Agência:")
        self.labelag.grid(column=1, row=24, padx=260, pady=1, sticky=W)
        self.entryag = ttk.Entry(self, width=20, textvariable=self.agencia)
        self.entryag.grid(column=1, row=24, padx=320, pady=1, sticky=W)
        # aparecer entry para conta
        self.labelcc = ttk.Label(self, width=20, text="Conta:")
        self.labelcc.grid(column=1, row=25, padx=260, pady=1, sticky=W)
        self.entrycc = ttk.Entry(self, width=20, textvariable=self.conta)
        self.entrycc.grid(column=1, row=25, padx=320, pady=1, sticky=W)
        # aparecer entry para ditigo
        self.labeldig = ttk.Label(self, width=20, text="Dígito:")
        self.labeldig.grid(column=1, row=26, padx=260, pady=1, sticky=W)
        self.entrydig = ttk.Entry(self, width=20, textvariable=self.digito)
        self.entrydig.grid(column=1, row=26, padx=320, pady=1, sticky=W)
        # checkbutton para indicar onde foi feito
        self.edicao = IntVar()
        self.editar = ttk.Checkbutton(self, text='Editar cadastro feito manualmente.', variable=self.edicao)
        self.editar.grid(column=1, row=26, padx=26, pady=1, sticky=W)
        self.feitonde = IntVar()
        self.onde = ttk.Checkbutton(self, text='Cadastro realizado fora da Cia.', variable=self.feitonde)
        self.onde.grid(column=1, row=27, padx=26, pady=1, sticky=W)

        def mostrar_horario(event):
            nome = event.widget.get()
            num, name = nome.split(' - ')
            linha = int(num)
            planwb = l_w(self.caminho.get())
            plansh = planwb['Respostas ao formulário 1']
            self.labelhor.config(text='Horário preenchido: ' + str(plansh[f'AI{linha}'].value))

        self.combonome.bind("<<ComboboxSelected>>", mostrar_horario)

        def carregarfunc(local):
            planwb = l_w(local)
            plansh = planwb['Respostas ao formulário 1']
            lista = []
            for x, pessoa in enumerate(plansh):
                lista.append(f'{x + 1} - {pessoa[2].value}')
            self.combonome.config(values=lista)

        self.botaocarregar = ttk.Button(self, text="Carregar planilha",
                                        command=lambda: [carregarfunc(self.caminho.get())])
        self.botaocarregar.grid(column=1, row=9, padx=350, pady=25, sticky=W)
        self.botaocadastrar = ttk.Button(self, width=20, text="Cadastrar no Dexion",
                                         command=lambda: [cadastro_funcionario(self.caminho.get(), self.edicao.get(),
                                                                               self.feitonde.get(),
                                                                               self.combonome.get(),
                                                                               self.entrymatr.get(),
                                                                               self.entryadmiss.get(),
                                                                               self.combohor.get(),
                                                                               self.entrysal.get(),
                                                                               self.combocargo.get(),
                                                                               self.combodepto.get(),
                                                                               self.combocontr.get(),
                                                                               self.hrs.get(),
                                                                               self.hrm.get(),
                                                                               self.agencia.get(),
                                                                               self.conta.get(),
                                                                               self.digito.get())])
        self.botaocadastrar.grid(column=1, row=28, padx=520, pady=1, sticky=W)
        self.botaosalvar = ttk.Button(self, width=20, text="Salvar Docs",
                                      command=lambda: [salvar_docs_funcionarios(self.entrymatr.get())])
        self.botaosalvar.grid(column=1, row=29, padx=520, pady=1, sticky=W)
        self.botaoenviaemail = ttk.Button(self, width=20, text="Enviar e-mails",
                                          command=lambda: [enviar_emails_funcionario(self.entrymatr.get())])
        self.botaoenviaemail.grid(column=1, row=30, padx=520, pady=1, sticky=W)

    def selecionar_funcionario(self):
        try:
            caminhoplan = tkinter.filedialog.askopenfilename(title='Planilha Funcionários')
            self.caminho.set(str(caminhoplan))
        except ValueError:
            pass


class Frame2(ttk.Frame):
    def __init__(self, container):
        super().__init__()
        self.hoje = datetime.today()
        self.caminhoest = StringVar()
        self.nomeest = StringVar()
        self.horarioest = StringVar()
        self.cargoest = StringVar()
        self.departamentoest = StringVar()
        self.tipocontrest = StringVar()
        self.nomesplanest = []
        self.labelescolhest = ttk.Label(self, width=40, text="Escolher planilha de novos estagiários")
        self.labelescolhest.grid(column=1, row=2, padx=25, pady=1, sticky=W)
        self.botaoescolhest = ttk.Button(self, text="Escolha a planilha", command=self.selecionarest)
        self.botaoescolhest.grid(column=1, row=2, padx=350, pady=1, sticky=W)
        self.labelnomest = ttk.Label(self, width=20, text="Nome:")
        self.labelnomest.grid(column=1, row=10, padx=25, pady=1, sticky=W)
        self.combonomest = ttk.Combobox(self, values=self.nomesplanest, textvariable=self.nomeest, width=50)
        self.combonomest.grid(column=1, row=10, padx=125, pady=1, sticky=W)
        # aparecer entry para preencher matricula
        self.labelmatrest = ttk.Label(self, width=20, text="Matrícula:")
        self.labelmatrest.grid(column=1, row=11, padx=25, pady=1, sticky=W)
        self.entrymatrest = ttk.Entry(self, width=20)
        self.entrymatrest.grid(column=1, row=11, padx=125, pady=1, sticky=W)
        # aparecer entry para preencher admissao
        self.labeladmissest = ttk.Label(self, width=20, text="Admissão:")
        self.labeladmissest.grid(column=1, row=12, padx=25, pady=1, sticky=W)
        self.entryadmissest = DateEntry(self, selectmode='day', year=self.hoje.year, month=self.hoje.month,
                                        day=self.hoje.day, locale='pt_BR')
        self.entryadmissest.grid(column=1, row=12, padx=125, pady=1, sticky=W)
        # aparecer combo departamento
        self.labeldeptoest = ttk.Label(self, width=20, text="Departamento:")
        self.labeldeptoest.grid(column=1, row=19, padx=25, pady=1, sticky=W)
        self.combodeptoest = ttk.Combobox(self, values=departamentos, textvariable=self.departamentoest, width=50)
        self.combodeptoest.grid(column=1, row=19, padx=125, pady=1, sticky=W)
        self.agenciaest = StringVar()
        self.contaest = StringVar()
        self.digitoest = StringVar()
        # aparecer entry para agencia
        self.labelagest = ttk.Label(self, width=20, text="Agência:")
        self.labelagest.grid(column=1, row=24, padx=260, pady=1, sticky=W)
        self.entryagest = ttk.Entry(self, width=20, textvariable=self.agenciaest)
        self.entryagest.grid(column=1, row=24, padx=320, pady=1, sticky=W)
        # aparecer entry para conta
        self.labelccest = ttk.Label(self, width=20, text="Conta:")
        self.labelccest.grid(column=1, row=25, padx=260, pady=1, sticky=W)
        self.entryccest = ttk.Entry(self, width=20, textvariable=self.contaest)
        self.entryccest.grid(column=1, row=25, padx=320, pady=1, sticky=W)
        # aparecer entry para ditigo
        self.labeldigest = ttk.Label(self, width=20, text="Dígito:")
        self.labeldigest.grid(column=1, row=26, padx=260, pady=1, sticky=W)
        self.entrydigest = ttk.Entry(self, width=20, textvariable=self.digitoest)
        self.entrydigest.grid(column=1, row=26, padx=320, pady=1, sticky=W)
        self.solicitarest = IntVar()
        self.solictest = ttk.Checkbutton(self, text='Apenas solicitar contrato.', variable=self.solicitarest)
        self.solictest.grid(column=1, row=25, padx=26, pady=1, sticky=W)
        self.edicaoest = IntVar()
        self.editarest = ttk.Checkbutton(self, text='Editar cadastro feito manualmente.', variable=self.edicaoest)
        self.editarest.grid(column=1, row=26, padx=26, pady=1, sticky=W)
        self.feitondeest = IntVar()
        self.ondeest = ttk.Checkbutton(self, text='Cadastro realizado fora da Cia.', variable=self.feitondeest)
        self.ondeest.grid(column=1, row=27, padx=26, pady=1, sticky=W)
        self.cargoest = StringVar()
        self.botaocadastrarest = ttk.Button(self, width=20, text="Cadastrar Estagiário",
                                            command=lambda: [
                                                cadastro_estagiario(
                                                    self.solicitarest.get(), self.caminhoest.get(),
                                                    self.edicaoest.get(), self.feitondeest.get(),
                                                    self.combonomest.get(),
                                                    self.entrymatrest.get(), self.entryadmissest.get(),
                                                    '', self.combodeptoest.get(),
                                                    '', '', '',
                                                    self.agenciaest.get(),
                                                    self.contaest.get(),
                                                    self.digitoest.get()
                                                )
                                            ]
                                            )
        self.botaocadastrarest.grid(column=1, row=28, padx=520, pady=1, sticky=W)

        def carregarest(local):
            planwb = l_w(local)
            plansh = planwb['Respostas ao formulário 1']
            lista = []
            for x, pessoa in enumerate(plansh):
                lista.append(f'{x + 1} - {pessoa[2].value}')
            self.combonomest.config(values=lista)

        self.botaocarregest = ttk.Button(self, text="Carregar planilha",
                                         command=lambda: [carregarest(self.caminhoest.get())])
        self.botaocarregest.grid(column=1, row=4, padx=350, pady=25, sticky=W)

    def selecionarest(self):
        try:
            caminhoplanest = tkinter.filedialog.askopenfilename(title='Planilha Estagiários')
            self.caminhoest.set(str(caminhoplanest))
        except ValueError:
            pass


class Frame3(ttk.Frame):
    def __init__(self, container):
        super().__init__()
        self.hoje = datetime.today()
        self.caminhoaut = StringVar()
        self.horarioaut = StringVar()
        self.cargoaut = StringVar()
        self.departamentoaut = StringVar()
        self.tipocontraut = StringVar()
        self.nomesplanaut = []
        self.labelescolh = ttk.Label(self, width=40, text="Escolher planilha de autônomos")
        self.labelescolh.grid(column=1, row=1, padx=25, pady=1, sticky=W)
        self.botaoescolh = ttk.Button(self, text="Escolha a planilha", command=self.selecionaraut)
        self.botaoescolh.grid(column=1, row=1, padx=350, pady=1, sticky=W)
        self.nomeaut = StringVar()
        self.cargo = StringVar()
        self.departamento = StringVar()
        self.nomesplanaut = []
        # aparecer dropdown com nomes da plan
        self.labelnomeaut = ttk.Label(self, width=20, text="Nome:")
        self.labelnomeaut.grid(column=1, row=10, padx=25, pady=1, sticky=W)
        self.combonomeaut = ttk.Combobox(self, values=self.nomesplanaut, textvariable=self.nomeaut, width=50)
        self.combonomeaut.grid(column=1, row=10, padx=125, pady=1, sticky=W)
        # aparecer entry para preencher matricula
        self.labelmatraut = ttk.Label(self, width=20, text="Matrícula:")
        self.labelmatraut.grid(column=1, row=11, padx=25, pady=1, sticky=W)
        self.entrymatraut = ttk.Entry(self, width=20)
        self.entrymatraut.grid(column=1, row=11, padx=125, pady=1, sticky=W)
        # aparecer entry para preencher admissao
        self.labeladmissaut = ttk.Label(self, width=20, text="Admissão:")
        self.labeladmissaut.grid(column=1, row=12, padx=25, pady=1, sticky=W)
        self.entryadmissaut = DateEntry(self, selectmode='day', year=self.hoje.year, month=self.hoje.month,
                                        day=self.hoje.day, locale='pt_BR')
        self.entryadmissaut.grid(column=1, row=12, padx=125, pady=1, sticky=W)
        # aparecer dropdown para escolher cargo
        self.labelcargoaut = ttk.Label(self, width=20, text="Cargo")
        self.labelcargoaut.grid(column=1, row=18, padx=25, pady=1, sticky=W)
        self.combocargoaut = ttk.Combobox(self, values=cargos, textvariable=self.cargo, width=50)
        self.combocargoaut.grid(column=1, row=18, padx=125, pady=1, sticky=W)
        # aparecer dropdown para escolher depto
        self.labeldeptoaut = ttk.Label(self, width=20, text="Departamento:")
        self.labeldeptoaut.grid(column=1, row=19, padx=25, pady=1, sticky=W)
        self.combodeptoaut = ttk.Combobox(self, values=departamentos, textvariable=self.departamento, width=50)
        self.combodeptoaut.grid(column=1, row=19, padx=125, pady=1, sticky=W)
        self.feitondeaut = IntVar()
        self.ondeaut = ttk.Checkbutton(self, text='Cadastro realizado fora da Cia.', variable=self.feitondeaut)
        self.ondeaut.grid(column=1, row=27, padx=26, pady=1, sticky=W)

        def carregaraut(local):
            planwb = l_w(local)
            plansh = planwb['Respostas ao formulário 1']
            lista = []
            for x, pessoa in enumerate(plansh):
                lista.append(f'{x + 1} - {pessoa[2].value}')
            self.combonomeaut.config(values=lista)

        self.botaocarregar = ttk.Button(self, text="Carregar planilha",
                                        command=lambda: [carregaraut(self.caminhoaut.get())])
        self.botaocarregar.grid(column=1, row=9, padx=350, pady=25, sticky=W)
        self.botaovalidarpis = ttk.Button(self, width=20, text="Validar PIS",
                                          command=lambda: [validar_pis(self.caminhoaut.get(), self.combonomeaut.get())])
        self.botaovalidarpis.grid(column=1, row=10, padx=520, pady=1, sticky=W)

        self.botaocarregar = ttk.Button(self, width=20, text="Cadastrar autônomo",
                                        command=lambda: [
                                            cadastrar_autonomo(self.caminhoaut.get(), self.combonomeaut.get(),
                                                               self.entrymatraut.get(), self.entryadmissaut.get(),
                                                               self.combocargoaut.get(),
                                                               self.combodeptoaut.get(), self.feitondeaut.get())])
        self.botaocarregar.grid(column=1, row=28, padx=520, pady=1, sticky=W)

    def selecionaraut(self):
        try:
            caminhoplanaut = tkinter.filedialog.askopenfilename(title='Planilha Autônomos')
            self.caminhoaut.set(str(caminhoplanaut))
        except ValueError:
            pass


class Frame4(ttk.Frame):
    def __init__(self, container):
        super().__init__()
        self.hoje = datetime.today()
        # definir funcionários ativos
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
        self.ativos = list(sorted(set(filter(None, self.grupo))))
        # campo nome
        self.labelnm = ttk.Label(self, text='Nome: ', width=25)
        self.labelnm.grid(pady=10, padx=15, column=1, row=1, sticky=W)
        self.combon = ttk.Combobox(self, values=self.ativos, width=40)
        self.combon.grid(pady=1, padx=140, column=1, row=1, sticky=W)
        # campo data
        self.labeldt = ttk.Label(self, text='Data do desligamento: ', width=25)
        self.labeldt.grid(pady=1, padx=15, column=1, row=2, sticky=W)
        self.dtentry = DateEntry(self, selectmode='day', year=self.hoje.year, month=self.hoje.month, day=self.hoje.day,
                                 locale='pt_BR')
        self.dtentry.grid(pady=1, padx=140, column=1, row=2, sticky=W)
        # campo tipo do desligamento
        self.labeltipo = ttk.Label(self, text='Tipo do desligamento:', width=25)
        self.labeltipo.grid(pady=1, padx=15, column=1, row=3, sticky=W)
        # radio buttons
        self.tipo = IntVar()
        self.estdem = ttk.Radiobutton(self, text='Estagiário', variable=self.tipo, value=1)
        self.estdem.grid(pady=5, padx=15, column=1, row=4, sticky=W)
        self.fpedav = ttk.Radiobutton(self, text='Funcionário (à pedido COM aviso)', variable=self.tipo, value=2)
        self.fpedav.grid(pady=5, padx=220, column=1, row=4, sticky=W)
        self.fpedsav = ttk.Radiobutton(self, text='Funcionário (à pedido SEM aviso)', variable=self.tipo, value=3)
        self.fpedsav.grid(pady=5, padx=15, column=1, row=5, sticky=W)
        self.fdemac = ttk.Radiobutton(self, text='Funcionário demitido por acordo', variable=self.tipo, value=4)
        self.fdemac.grid(pady=5, padx=220, column=1, row=5, sticky=W)
        self.fdemsav = ttk.Radiobutton(self, text='Funcionário demitido SEM aviso', variable=self.tipo, value=5)
        self.fdemsav.grid(pady=5, padx=15, column=1, row=6, sticky=W)
        self.fdemav = ttk.Radiobutton(self, text='Funcionário demitido COM aviso', variable=self.tipo, value=6)
        self.fdemav.grid(pady=5, padx=220, column=1, row=6, sticky=W)

        # button registrar desligamento
        self.btdesligar = ttk.Button(self, text="Registrar desligamento",
                                     command=lambda: [
                                         desligar_pessoa(self.combon.get(), self.dtentry.get(), self.tipo.get())
                                     ])
        self.btdesligar.grid(column=1, row=7, padx=480, pady=1, sticky=W)


class Frame5(ttk.Frame):
    def __init__(self, container):
        super().__init__()
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
        self.ativos = list(sorted(set(filter(None, self.grupo))))
        # choose employee
        self.labelnom = ttk.Label(self, text='Escolha o colaborador a ser atualizado: ')
        self.labelnom.grid(column=1, row=1, padx=10, pady=5, sticky=W)
        self.combnom = ttk.Combobox(self, values=self.ativos, width=40)
        self.combnom.grid(column=2, row=1, padx=10, pady=5, sticky=W)
        # add label to choose wich data the user want to update
        self.labelescol = ttk.Label(self, text='Escolha a informação que deseja atualizar:')
        self.labelescol.grid(column=1, row=2, padx=10, pady=15, sticky=W)
        # radio buttons
        self.tipo = IntVar()
        self.nom = ttk.Radiobutton(self, text='Nome', variable=self.tipo, value=1)
        self.nom.grid(pady=5, padx=15, column=1, row=3, sticky=W)
        self.crg = ttk.Radiobutton(self, text='Cargo', variable=self.tipo, value=2)
        self.crg.grid(pady=5, padx=10, column=2, row=3, sticky=W)
        self.dept = ttk.Radiobutton(self, text='Departamento', variable=self.tipo, value=3)
        self.dept.grid(pady=5, padx=15, column=1, row=4, sticky=W)
        self.conta = ttk.Radiobutton(self, text='Conta Bancária', variable=self.tipo, value=4)
        self.conta.grid(pady=5, padx=10, column=2, row=4, sticky=W)
        self.bttatualizar = ttk.Button(self, text='Atualizar cadastro', command=lambda: [])
        self.bttatualizar.grid(pady=5, padx=220, column=2, row=10, sticky=W)


if __name__ == '__main__':
    app = MainApplication()
    app.mainloop()
