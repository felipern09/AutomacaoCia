import datetime
import tkinter as tk
from src.controler.funcoes import lancar_ferias, lancar_atestado, lancar_desligamento, lancar_substit, \
    lancar_hrscomple, lancar_faltas, lancar_escala, lancar_novaaula, salvar_banco_aulas, inativar_aulas
from tkinter import ttk
from tkinter import *
from src.models.models import Colaborador, engine
from src.models.modelsfolha import Aulas, enginefolha
from sqlalchemy.orm import sessionmaker
from tkcalendar import DateEntry


class MainApplication(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Cálculo de folha - Cia BSB")
        self.geometry('550x300')
        self.img = PhotoImage(file='../models/static/imgs/Icone.png')
        self.iconphoto(False, self.img)
        self.columnconfigure(0, weight=5)
        self.rowconfigure(0, weight=5)
        for child in self.winfo_children():
            child.grid_configure(padx=1, pady=3)
        self.notebook = ttk.Notebook(self)

        self.Frame1 = Complementares(self.notebook)
        self.Frame2 = Faltas(self.notebook)
        self.Frame3 = Substituicoes(self.notebook)
        self.Frame4 = Ferias(self.notebook)
        self.Frame5 = Deslig(self.notebook)
        self.Frame6 = Atestados(self.notebook)
        self.Frame7 = Escala(self.notebook)
        self.Frame8 = NovaAula(self.notebook)
        self.Frame9 = InativarAula(self.notebook)

        self.notebook.add(self.Frame1, text='Hr Compl')
        self.notebook.add(self.Frame2, text='Faltas')
        self.notebook.add(self.Frame3, text='Substituições')
        self.notebook.add(self.Frame4, text='Férias')
        self.notebook.add(self.Frame5, text='Desligamentos')
        self.notebook.add(self.Frame6, text='Atestados')
        self.notebook.add(self.Frame7, text='Esacala')
        self.notebook.add(self.Frame8, text='Nova Aula')
        self.notebook.add(self.Frame9, text='Inativar Aula')
        self.notebook.pack()


class Complementares(ttk.Frame):
    def __init__(self, container):
        super().__init__()
        self.hoje = datetime.datetime.today()
        sessions = sessionmaker(bind=engine)
        session = sessions()
        self.grupo = []
        pessoas = session.query(Colaborador).filter_by(desligamento=None).all()
        for pess in pessoas:
            if pess.nome != '':
                self.grupo.append(pess.nome)
        self.nomes = list(sorted(set(filter(None, self.grupo))))

        sessionsaula = sessionmaker(bind=enginefolha)
        sessionaula = sessionsaula()
        self.grupodept = []
        aulas = sessionaula.query(Aulas).all()
        for aul in aulas:
            if aul.departamento != '':
                self.grupodept.append(aul.departamento)
        self.deptoaulas = list(sorted(set(filter(None, self.grupodept))))
        self.grupoaulas = []
        aulas = sessionaula.query(Aulas).all()
        for aula in aulas:
            if aula.nome != '':
                self.grupoaulas.append(aula.nome)
        self.aaulas = list(sorted(set(filter(None, self.grupoaulas))))

        # nome
        self.labelnome = ttk.Label(self, width=20, text='Nome:')
        self.labelnome.grid(column=1, row=1, padx=5, pady=2, sticky=W)
        # combo nomes
        self.combonome = ttk.Combobox(self, width=45, values=self.nomes)
        self.combonome.grid(column=1, row=2, padx=25, pady=2, sticky=W)
        # Depto
        self.labeldepto = ttk.Label(self, width=20, text='Departamento:')
        self.labeldepto.grid(column=1, row=3, padx=5, pady=2, sticky=W)
        # combo Deptos
        self.combodepto = ttk.Combobox(self, width=45, values=self.deptoaulas)
        self.combodepto.grid(column=1, row=4, padx=25, pady=2, sticky=W)
        # Aula
        self.labelaula = ttk.Label(self, width=20, text='Aula:')
        self.labelaula.grid(column=1, row=5, padx=5, pady=2, sticky=W)
        # combo Aulas
        self.comboaula = ttk.Combobox(self, width=45, values=self.aaulas)
        self.comboaula.grid(column=1, row=6, padx=25, pady=2, sticky=W)
        # Data
        self.labeldt = ttk.Label(self, width=20, text='Data:')
        self.labeldt.grid(column=1, row=7, padx=5, pady=2, sticky=W)
        self.entrydt = DateEntry(self, selectmode='day', year=self.hoje.year, month=self.hoje.month,
                                     day=self.hoje.day, locale='pt_BR')
        self.entrydt.grid(column=1, row=7, padx=45, pady=2, sticky=W)
        # Horas
        self.labelhr = ttk.Label(self, width=20, text='Horas:')
        self.labelhr.grid(column=1, row=8, padx=5, pady=2, sticky=W)
        # entry hr
        self.entryhr = ttk.Entry(self, width=20)
        self.entryhr.grid(column=1, row=8, padx=45, pady=2, sticky=W)
        # gerar folha da competencia selecionada
        self.botaogerar = ttk.Button(self, width=20, text="Lançar Hr compl.", command=lambda: [
            lancar_hrscomple(self.combonome.get(), self.combodepto.get(), self.comboaula.get(),
                             self.entrydt.get(), float(str(self.entryhr.get()).replace(',','.')))])
        self.botaogerar.grid(column=1, row=30, padx=190, pady=1, sticky=W)


class Faltas(ttk.Frame):
    def __init__(self, container):
        super().__init__()
        self.hoje = datetime.datetime.today()
        sessions = sessionmaker(bind=engine)
        session = sessions()
        self.grupo = []
        pessoas = session.query(Colaborador).filter_by(desligamento=None).all()
        for pess in pessoas:
            if pess.nome != '':
                self.grupo.append(pess.nome)
        self.nomes = list(sorted(set(filter(None, self.grupo))))

        sessionsaula = sessionmaker(bind=enginefolha)
        sessionaula = sessionsaula()
        self.grupodept = []
        aulas = sessionaula.query(Aulas).all()
        for aul in aulas:
            if aul.departamento != '':
                self.grupodept.append(aul.departamento)
        self.deptoaulas = list(sorted(set(filter(None, self.grupodept))))

        # nome
        self.labelnome = ttk.Label(self, width=20, text='Nome:')
        self.labelnome.grid(column=1, row=1, padx=5, pady=2, sticky=W)
        # combo nomes
        self.combonome = ttk.Combobox(self, width=45, values=self.nomes)
        self.combonome.grid(column=1, row=2, padx=25, pady=2, sticky=W)
        # Depto
        self.labeldepto = ttk.Label(self, width=20, text='Departamento:')
        self.labeldepto.grid(column=1, row=3, padx=5, pady=2, sticky=W)
        # combo Deptos
        self.combodepto = ttk.Combobox(self, width=45, values=self.deptoaulas)
        self.combodepto.grid(column=1, row=4, padx=25, pady=2, sticky=W)
        # Data
        self.labeldt = ttk.Label(self, width=20, text='Data:')
        self.labeldt.grid(column=1, row=7, padx=5, pady=2, sticky=W)
        self.entrydt = DateEntry(self, selectmode='day', year=self.hoje.year, month=self.hoje.month,
                                     day=self.hoje.day, locale='pt_BR')
        self.entrydt.grid(column=1, row=7, padx=45, pady=2, sticky=W)
        # Horas
        self.labelhr = ttk.Label(self, width=20, text='Horas:')
        self.labelhr.grid(column=1, row=8, padx=5, pady=2, sticky=W)
        # entry hr
        self.entryhr = ttk.Entry(self, width=20)
        self.entryhr.grid(column=1, row=8, padx=45, pady=2, sticky=W)
        # gerar folha da competencia selecionada
        self.botaogerar = ttk.Button(self, width=20, text="Lançar Falta", command=lambda: [
            lancar_faltas(self.combonome.get(), self.combodepto.get(), self.entrydt.get(),
                          float(str(self.entryhr.get()).replace(',','.')))])
        self.botaogerar.grid(column=1, row=30, padx=190, pady=1, sticky=W)


class Substituicoes(ttk.Frame):
    def __init__(self, container):
        super().__init__()
        self.hoje = datetime.datetime.today()
        sessions = sessionmaker(bind=engine)
        session = sessions()
        self.grupo = []
        pessoas = session.query(Colaborador).filter_by(desligamento=None).all()
        for pess in pessoas:
            if pess.nome != '':
                self.grupo.append(pess.nome)
        self.nomes = list(sorted(set(filter(None, self.grupo))))

        sessionsaula = sessionmaker(bind=enginefolha)
        sessionaula = sessionsaula()
        self.grupodept = []
        aulas = sessionaula.query(Aulas).all()
        for aul in aulas:
            if aul.departamento != '':
                self.grupodept.append(aul.departamento)
        self.deptoaulas = list(sorted(set(filter(None, self.grupodept))))
        self.grupoaulas = []
        aulas = sessionaula.query(Aulas).all()
        for aula in aulas:
            if aula.nome != '':
                self.grupoaulas.append(aula.nome)
        self.aaulas = list(sorted(set(filter(None, self.grupoaulas))))
        # nome
        self.labelsubstituto = ttk.Label(self, width=20, text='Substituto:')
        self.labelsubstituto.grid(column=1, row=3, padx=5, pady=2, sticky=W)
        # combo nomes
        self.combosubstituto = ttk.Combobox(self, width=45, values=self.nomes)
        self.combosubstituto.grid(column=1, row=4, padx=25, pady=2, sticky=W)
        # nome
        self.labelnome = ttk.Label(self, width=20, text='Substituído:')
        self.labelnome.grid(column=1, row=1, padx=5, pady=2, sticky=W)
        # combo nomes
        self.combonome = ttk.Combobox(self, width=45, values=self.nomes)
        self.combonome.grid(column=1, row=2, padx=25, pady=2, sticky=W)
        # Depto
        self.labeldepto = ttk.Label(self, width=20, text='Departamento:')
        self.labeldepto.grid(column=1, row=5, padx=5, pady=2, sticky=W)
        # combo Deptos
        self.combodepto = ttk.Combobox(self, width=45, values=self.deptoaulas)
        self.combodepto.grid(column=1, row=6, padx=25, pady=2, sticky=W)
        # Aula
        self.labelaula = ttk.Label(self, width=20, text='Aula:')
        self.labelaula.grid(column=1, row=7, padx=5, pady=2, sticky=W)
        # combo Aulas
        self.comboaula = ttk.Combobox(self, width=45, values=self.aaulas)
        self.comboaula.grid(column=1, row=8, padx=25, pady=2, sticky=W)
        # Data
        self.labeldt = ttk.Label(self, width=20, text='Data:')
        self.labeldt.grid(column=1, row=9, padx=5, pady=2, sticky=W)
        self.entrydt = DateEntry(self, selectmode='day', year=self.hoje.year, month=self.hoje.month,
                                     day=self.hoje.day, locale='pt_BR')
        self.entrydt.grid(column=1, row=9, padx=45, pady=2, sticky=W)
        # Horas
        self.labelhr = ttk.Label(self, width=20, text='Horas:')
        self.labelhr.grid(column=1, row=10, padx=5, pady=2, sticky=W)
        # entry hr
        self.entryhr = ttk.Entry(self, width=20)
        self.entryhr.grid(column=1, row=10, padx=45, pady=2, sticky=W)
        # gerar folha da competencia selecionada
        self.botaogerar = ttk.Button(self, width=20, text="Lançar Substit", command=lambda: [
            lancar_substit(self.combonome.get(), self.combosubstituto.get(),
                           self.combodepto.get(), self.comboaula.get(), self.entrydt.get(),
                           float(str(self.entryhr.get()).replace(',','.')))])
        self.botaogerar.grid(column=1, row=30, padx=190, pady=1, sticky=W)


class Ferias(ttk.Frame):
    def __init__(self, container):
        super().__init__()
        self.hoje = datetime.datetime.today()
        sessions = sessionmaker(bind=engine)
        session = sessions()
        self.grupo = []
        pessoas = session.query(Colaborador).filter_by(desligamento=None).all()
        for pess in pessoas:
            if pess.nome != '':
                self.grupo.append(pess.nome)
        self.nomes = list(sorted(set(filter(None, self.grupo))))
        sessionsaula = sessionmaker(bind=enginefolha)
        sessionaula = sessionsaula()
        self.grupodept = []
        aulas = sessionaula.query(Aulas).all()
        for aul in aulas:
            if aul.departamento != '':
                self.grupodept.append(aul.departamento)
        self.deptoaulas = list(sorted(set(filter(None, self.grupodept))))
        self.grupoaulas = []
        # nome
        self.labelnome = ttk.Label(self, width=20, text='Nome:')
        self.labelnome.grid(column=1, row=1, padx=5, pady=2, sticky=W)
        # combo nomes
        self.combonome = ttk.Combobox(self, width=45, values=self.nomes)
        self.combonome.grid(column=1, row=2, padx=25, pady=2, sticky=W)
        # Depto
        self.labeldepto = ttk.Label(self, width=20, text='Departamento:')
        self.labeldepto.grid(column=1, row=5, padx=5, pady=2, sticky=W)
        # combo Deptos
        self.combodepto = ttk.Combobox(self, width=45, values=self.deptoaulas)
        self.combodepto.grid(column=1, row=6, padx=25, pady=2, sticky=W)

        # Data
        self.labeldt = ttk.Label(self, width=20, text='Início:')
        self.labeldt.grid(column=1, row=9, padx=5, pady=2, sticky=W)
        self.entrydti = DateEntry(self, selectmode='day', year=self.hoje.year, month=self.hoje.month,
                                     day=self.hoje.day, locale='pt_BR')
        self.entrydti.grid(column=1, row=9, padx=45, pady=2, sticky=W)
        # Data
        self.labeldt = ttk.Label(self, width=20, text='Fim:')
        self.labeldt.grid(column=1, row=10, padx=5, pady=2, sticky=W)
        self.entrydtf = DateEntry(self, selectmode='day', year=self.hoje.year, month=self.hoje.month,
                                     day=self.hoje.day, locale='pt_BR')
        self.entrydtf.grid(column=1, row=10, padx=45, pady=2, sticky=W)
        # gerar folha da competencia selecionada
        self.botaogerar = ttk.Button(self, width=20, text="Lançar Férias", command=lambda: [
            lancar_ferias(self.combonome.get(), self.combodepto.get(), self.entrydti.get(), self.entrydtf.get())])
        self.botaogerar.grid(column=1, row=30, padx=190, pady=1, sticky=W)


class Deslig(ttk.Frame):
    def __init__(self, container):
        super().__init__()
        self.hoje = datetime.datetime.today()
        sessions = sessionmaker(bind=engine)
        session = sessions()
        self.grupo = []
        pessoas = session.query(Colaborador).all()
        for pess in pessoas:
            if pess.nome != '':
                self.grupo.append(pess.nome)
        self.nomes = list(sorted(set(filter(None, self.grupo))))
        sessionsaula = sessionmaker(bind=enginefolha)
        sessionaula = sessionsaula()
        self.grupodept = []
        aulas = sessionaula.query(Aulas).all()
        for aul in aulas:
            if aul.departamento != '':
                self.grupodept.append(aul.departamento)
        self.deptoaulas = list(sorted(set(filter(None, self.grupodept))))
        self.grupoaulas = []
        # nome
        self.labelnome = ttk.Label(self, width=20, text='Nome:')
        self.labelnome.grid(column=1, row=1, padx=5, pady=2, sticky=W)
        # combo nomes
        self.combonome = ttk.Combobox(self, width=45, values=self.nomes)
        self.combonome.grid(column=1, row=2, padx=25, pady=2, sticky=W)
        # Depto
        self.labeldepto = ttk.Label(self, width=20, text='Departamento:')
        self.labeldepto.grid(column=1, row=5, padx=5, pady=2, sticky=W)
        # combo Deptos
        self.combodepto = ttk.Combobox(self, width=45, values=self.deptoaulas)
        self.combodepto.grid(column=1, row=6, padx=25, pady=2, sticky=W)

        # Data
        self.labeldt = ttk.Label(self, width=20, text='Data:')
        self.labeldt.grid(column=1, row=9, padx=5, pady=2, sticky=W)
        self.entrydt = DateEntry(self, selectmode='day', year=self.hoje.year, month=self.hoje.month,
                                     day=self.hoje.day, locale='pt_BR')
        self.entrydt.grid(column=1, row=9, padx=45, pady=2, sticky=W)

        # gerar folha da competencia selecionada
        self.botaogerar = ttk.Button(self, width=20, text="Lançar Deslig", command=lambda: [
            lancar_desligamento(self.combonome.get(), self.combodepto.get(), self.entrydt.get())])
        self.botaogerar.grid(column=1, row=30, padx=190, pady=1, sticky=W)


class Atestados(ttk.Frame):
    def __init__(self, container):
        super().__init__()
        self.hoje = datetime.datetime.today()
        sessions = sessionmaker(bind=engine)
        session = sessions()
        self.grupo = []
        pessoas = session.query(Colaborador).filter_by(desligamento=None).all()
        for pess in pessoas:
            if pess.nome != '':
                self.grupo.append(pess.nome)
        self.nomes = list(sorted(set(filter(None, self.grupo))))
        sessionsaula = sessionmaker(bind=enginefolha)
        sessionaula = sessionsaula()
        self.grupodept = []
        aulas = sessionaula.query(Aulas).all()
        for aul in aulas:
            if aul.departamento != '':
                self.grupodept.append(aul.departamento)
        self.deptoaulas = list(sorted(set(filter(None, self.grupodept))))
        self.grupoaulas = []
        # nome
        self.labelnome = ttk.Label(self, width=20, text='Nome:')
        self.labelnome.grid(column=1, row=1, padx=5, pady=2, sticky=W)
        # combo nomes
        self.combonome = ttk.Combobox(self, width=45, values=self.nomes)
        self.combonome.grid(column=1, row=2, padx=25, pady=2, sticky=W)
        # Depto
        self.labeldepto = ttk.Label(self, width=20, text='Departamento:')
        self.labeldepto.grid(column=1, row=5, padx=5, pady=2, sticky=W)
        # combo Deptos
        self.combodepto = ttk.Combobox(self, width=45, values=self.deptoaulas)
        self.combodepto.grid(column=1, row=6, padx=25, pady=2, sticky=W)

        # Data
        self.labeldt = ttk.Label(self, width=20, text='Data:')
        self.labeldt.grid(column=1, row=9, padx=5, pady=2, sticky=W)
        self.entrydt = DateEntry(self, selectmode='day', year=self.hoje.year, month=self.hoje.month,
                                     day=self.hoje.day, locale='pt_BR')
        self.entrydt.grid(column=1, row=9, padx=45, pady=2, sticky=W)

        # gerar folha da competencia selecionada
        self.botaogerar = ttk.Button(self, width=20, text="Lançar Atestado", command=lambda: [
            lancar_atestado(self.combonome.get(), self.combodepto.get(), self.entrydt.get())])
        self.botaogerar.grid(column=1, row=30, padx=190, pady=1, sticky=W)


class Escala(ttk.Frame):
    def __init__(self, container):
        super().__init__()
        self.hoje = datetime.datetime.today()
        sessions = sessionmaker(bind=engine)
        session = sessions()
        self.grupo = []
        pessoas = session.query(Colaborador).filter_by(desligamento=None).all()
        for pess in pessoas:
            if pess.nome != '':
                self.grupo.append(pess.nome)
        self.nomes = list(sorted(set(filter(None, self.grupo))))

        sessionsaula = sessionmaker(bind=enginefolha)
        sessionaula = sessionsaula()
        self.grupodept = []
        aulas = sessionaula.query(Aulas).all()
        for aul in aulas:
            if aul.departamento != '':
                self.grupodept.append(aul.departamento)
        self.deptoaulas = list(sorted(set(filter(None, self.grupodept))))
        self.grupoaulas = []
        aulas = sessionaula.query(Aulas).all()
        for aula in aulas:
            if aula.nome != '':
                self.grupoaulas.append(aula.nome)
        self.aaulas = list(sorted(set(filter(None, self.grupoaulas))))

        # nome
        self.labelnome = ttk.Label(self, width=20, text='Nome:')
        self.labelnome.grid(column=1, row=1, padx=5, pady=2, sticky=W)
        # combo nomes
        self.combonome = ttk.Combobox(self, width=45, values=self.nomes)
        self.combonome.grid(column=1, row=2, padx=25, pady=2, sticky=W)
        # Depto
        self.labeldepto = ttk.Label(self, width=20, text='Departamento:')
        self.labeldepto.grid(column=1, row=3, padx=5, pady=2, sticky=W)
        # combo Deptos
        self.combodepto = ttk.Combobox(self, width=45, values=self.deptoaulas)
        self.combodepto.grid(column=1, row=4, padx=25, pady=2, sticky=W)
        # Aula
        self.labelaula = ttk.Label(self, width=20, text='Aula:')
        self.labelaula.grid(column=1, row=5, padx=5, pady=2, sticky=W)
        # combo Aulas
        self.comboaula = ttk.Combobox(self, width=45, values=self.aaulas)
        self.comboaula.grid(column=1, row=6, padx=25, pady=2, sticky=W)
        # Data
        self.labeldt = ttk.Label(self, width=20, text='Data:')
        self.labeldt.grid(column=1, row=7, padx=5, pady=2, sticky=W)
        self.entrydt = DateEntry(self, selectmode='day', year=self.hoje.year, month=self.hoje.month,
                                     day=self.hoje.day, locale='pt_BR')
        self.entrydt.grid(column=1, row=7, padx=45, pady=2, sticky=W)
        # Horas
        self.labelhr = ttk.Label(self, width=20, text='Horas:')
        self.labelhr.grid(column=1, row=8, padx=5, pady=2, sticky=W)
        # entry hr
        self.entryhr = ttk.Entry(self, width=20)
        self.entryhr.grid(column=1, row=8, padx=45, pady=2, sticky=W)
        # gerar folha da competencia selecionada
        self.botaogerar = ttk.Button(self, width=20, text="Lançar Escala.", command=lambda: [
            lancar_escala(self.combonome.get(), self.combodepto.get(), self.comboaula.get(),
                          self.entrydt.get(), float(str(self.entryhr.get()).replace(',','.')))])
        self.botaogerar.grid(column=1, row=30, padx=190, pady=1, sticky=W)


class NovaAula(ttk.Frame):
    def __init__(self, container):
        super().__init__()
        self.hoje = datetime.datetime.today()
        sessions = sessionmaker(bind=engine)
        session = sessions()
        self.grupo = []
        pessoas = session.query(Colaborador).filter_by(desligamento=None).all()
        for pess in pessoas:
            if pess.nome != '':
                self.grupo.append(pess.nome)
        self.nomes = list(sorted(set(filter(None, self.grupo))))
        self.dias = ['Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado', 'Domingo']
        sessionsaula = sessionmaker(bind=enginefolha)
        sessionaula = sessionsaula()
        self.grupodept = []
        aulas = sessionaula.query(Aulas).all()
        for aul in aulas:
            if aul.departamento != '':
                self.grupodept.append(aul.departamento)
        self.deptoaulas = list(sorted(set(filter(None, self.grupodept))))
        self.grupoaulas = []
        aulas = sessionaula.query(Aulas).all()
        for aula in aulas:
            if aula.nome != '':
                self.grupoaulas.append(aula.nome)
        self.aaulas = list(sorted(set(filter(None, self.grupoaulas))))

        # nome
        self.labelnome = ttk.Label(self, width=20, text='Nome:')
        self.labelnome.grid(column=1, row=1, padx=5, pady=2, sticky=W)
        # combo nomes
        self.combonome = ttk.Combobox(self, width=45, values=self.nomes)
        self.combonome.grid(column=1, row=2, padx=25, pady=2, sticky=W)
        # Depto
        self.labeldepto = ttk.Label(self, width=20, text='Departamento:')
        self.labeldepto.grid(column=1, row=3, padx=5, pady=2, sticky=W)
        # combo Deptos
        self.combodepto = ttk.Combobox(self, width=45, values=self.deptoaulas)
        self.combodepto.grid(column=1, row=4, padx=25, pady=2, sticky=W)
        # Aula
        self.labelaula = ttk.Label(self, width=20, text='Aula:')
        self.labelaula.grid(column=1, row=5, padx=5, pady=2, sticky=W)
        # combo Aulas
        self.comboaula = ttk.Combobox(self, width=45, values=self.aaulas)
        self.comboaula.grid(column=1, row=6, padx=25, pady=2, sticky=W)
        # Dia Semana
        self.labeldias = ttk.Label(self, width=20, text='Dia da Semana:')
        self.labeldias.grid(column=1, row=7, padx=5, pady=2, sticky=W)
        # combo Dias
        self.combodias = ttk.Combobox(self, width=17, values=self.dias)
        self.combodias.grid(column=1, row=7, padx=95, pady=2, sticky=W)
        # Inicio
        self.labeliniciohr = ttk.Label(self, width=20, text='Início:')
        self.labeliniciohr.grid(column=1, row=8, padx=5, pady=2, sticky=W)
        # entry inicio
        self.entryiniciohr = ttk.Entry(self, width=20)
        self.entryiniciohr.grid(column=1, row=8, padx=95, pady=2, sticky=W)
        # Fim
        self.labelfimhr = ttk.Label(self, width=20, text='Fim:')
        self.labelfimhr.grid(column=1, row=9, padx=5, pady=2, sticky=W)
        # entry fim
        self.entryfimhr = ttk.Entry(self, width=20)
        self.entryfimhr.grid(column=1, row=9, padx=95, pady=2, sticky=W)
        # Valor
        self.labelvalor = ttk.Label(self, width=20, text='Valor:')
        self.labelvalor.grid(column=1, row=10, padx=5, pady=2, sticky=W)
        # entry Valor
        self.entryvalor = ttk.Entry(self, width=20)
        self.entryvalor.grid(column=1, row=10, padx=95, pady=2, sticky=W)
        # gerar folha da competencia selecionada
        self.botaogerar = ttk.Button(self, width=20, text="Lançar Nova Aula", command=lambda: [
            lancar_novaaula(self.combonome.get(), self.combodepto.get(), self.comboaula.get(),
                            self.combodias.get(), self.entryiniciohr.get(), self.entryfimhr.get(),
                            self.entryvalor.get())])
        self.botaogerar.grid(column=1, row=30, padx=190, pady=1, sticky=W)
        self.botaosalvar = ttk.Button(self, width=20, text="Salvar Banco de Aulas", command=lambda: [
            salvar_banco_aulas()])
        self.botaosalvar.grid(column=1, row=30, padx=350, pady=1, sticky=W)


class InativarAula(ttk.Frame):
    def __init__(self, container):
        super().__init__()
        self.hoje = datetime.datetime.today()
        sessions = sessionmaker(bind=engine)
        session = sessions()
        self.grupo = []
        pessoas = session.query(Colaborador).filter_by(desligamento=None).all()
        for pess in pessoas:
            if pess.nome != '':
                self.grupo.append(pess.nome)
        self.nomes = list(sorted(set(filter(None, self.grupo))))
        self.canvas = Canvas(self)
        self.canvas.pack(side=LEFT, fill=BOTH, expand=1)
        self.barraroll = ttk.Scrollbar(self, orient=VERTICAL, command=self.canvas.yview)
        self.barraroll.pack(side=LEFT, fill=Y)
        self.canvas.config(yscrollcommand=self.barraroll.set)
        self.canvas.bind('<Configure>', lambda e: self.canvas.config(scrollregion=self.canvas.bbox('all')))
        self.canvframe = Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.canvframe, anchor='nw')
        self.aulas_para_alterar = []

        def seleciona_aula(event):
            widget = event.widget
            numero = int(str(widget.cget('text')).split(' ')[0])
            if numero in self.aulas_para_alterar:
                self.aulas_para_alterar.remove(numero)
            else:
                self.aulas_para_alterar.append(numero)
            self.aulas_para_alterar.sort()

        def mostrar_aulas(event):
            nome = event.widget.get()
            sessions = sessionmaker(enginefolha)
            session = sessions()
            aulas = session.query(Aulas).filter_by(professor=nome).filter_by(status='Ativa').all()
            try:
                self.item.grid_remove()
            except AttributeError:
                pass
            for i, aula in enumerate(aulas):
                texto = f'{aula.numero} - {aula.departamento} - {aula.nome}: {aula.diadasemana} de {datetime.datetime.strftime(datetime.datetime.strptime(aula.inicio, "%H:%M:%S"), "%H:%M")} às {datetime.datetime.strftime(datetime.datetime.strptime(aula.fim, "%H:%M:%S"), "%H:%M")}'
                var_name = f'var_{i}'
                value = IntVar()
                globals()[var_name] = value
                self.item = tk.Checkbutton(self.canvframe, text=texto, variable=globals()[var_name])
                self.item.grid(column=1, row=i+3, padx=25, pady=2, sticky=W)
                self.item.bind('<Button-1>', seleciona_aula)
        # nome
        self.labelnome = ttk.Label(self.canvframe, width=20, text='Nome:')
        self.labelnome.grid(column=1, row=1, padx=5, pady=2, sticky=W)
        # combo nomes
        self.combonome = ttk.Combobox(self.canvframe, width=45, values=self.nomes)
        self.combonome.grid(column=1, row=2, padx=25, pady=2, sticky=W)
        self.combonome.bind('<<ComboboxSelected>>', mostrar_aulas)
        # gerar folha da competencia selecionada
        self.botaogerar = ttk.Button(self.canvframe, width=20, text="Inativar Aula", command=lambda: [
            inativar_aulas(self.aulas_para_alterar)
        ])
        self.botaogerar.grid(column=1, row=190, padx=190, pady=1, sticky=W)
        self.botaosalvar = ttk.Button(self.canvframe, width=20, text="Salvar Banco de Aulas", command=lambda: [
            salvar_banco_aulas()])
        self.botaosalvar.grid(column=1, row=190, padx=350, pady=1, sticky=W)


if __name__ == '__main__':
    app = MainApplication()
    app.mainloop()
