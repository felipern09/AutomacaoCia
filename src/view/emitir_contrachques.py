import datetime
import tkinter as tk
from src.controler.f_contrach import salvar_holerites, incluir_grade_email_holerite
import tkinter.filedialog
from tkinter import ttk
from tkinter import *
from src.models.models import Colaborador, engine
from sqlalchemy.orm import sessionmaker
from tkcalendar import DateEntry


class MainApplication(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Salvar e Enviar Contracheques")
        self.geometry('440x250')
        self.img = PhotoImage(file='../models/static/imgs/Icone.png')
        self.iconphoto(False, self.img)
        self.columnconfigure(0, weight=5)
        self.rowconfigure(0, weight=5)
        for child in self.winfo_children():
            child.grid_configure(padx=1, pady=3)
        self.notebook = ttk.Notebook(self)
        self.Frame1 = Salvar(self.notebook)
        self.Frame2 = AnexarGrade(self.notebook)
        self.Frame3 = EnviarEmail(self.notebook)
        self.notebook.add(self.Frame1, text='Salvar Contracheques')
        self.notebook.add(self.Frame2, text='Anexar Grades')
        self.notebook.add(self.Frame3, text='Enviar por e-mail')
        self.notebook.pack()


class Salvar(ttk.Frame):
    def __init__(self, container):
        super().__init__()
        # descrição funçoes para usuraio
        self.labeldescr = ttk.Label(self, width=100, text="A função de emissão de contracheque deve obedecer a "
                                                          "seguinte sequência:")
        self.labeldescr.grid(column=1, row=1, padx=10, pady=5, sticky=W)
        self.labeldescr1 = ttk.Label(self, width=100,
                                    text="1º - Gerar Contracheques no Dexion (procedimentos automat.)"
                                         "\n2º - Salvar Contracheques (botão 'Salvar' dessa aba)"
                                         "\n3º - Anexar Grades (aba seguinte)"
                                         "\n4º - Enviar Contracheques Pelo Dexion")
        self.labeldescr1.grid(column=1, row=2, padx=10, pady=8, sticky=W)
        # aparecer dropdown com nomes da plan
        self.labelcomp = ttk.Label(self, width=60, text="Salvar contracheques nas pastas pessoais.")
        self.labelcomp.grid(column=1, row=10, padx=10, pady=5, sticky=W)
        # gerar folha da competencia selecionada
        self.botaogerar = ttk.Button(self, text="Salvar", command=lambda: [salvar_holerites()])
        self.botaogerar.grid(column=1, row=11, padx=230, pady=1, sticky=W)


class AnexarGrade(ttk.Frame):
    def __init__(self, container):
        super().__init__()
        self.hoje = datetime.datetime.today()
        self.comp = list(range(1, 13))
        col2 = 120
        # aparecer dropdown com nomes da plan
        self.labelcomp = ttk.Label(self, width=60, text="Anexar grades aos arquivos '.zip' do dexion.")
        self.labelcomp.grid(column=1, row=1, padx=10, pady=1, sticky=W)
        # selecionar arquivo da grade
        self.labelarq = ttk.Label(self, width=60, text="Arquivo da grade:")
        self.labelarq.grid(column=1, row=2, padx=10, pady=2, sticky=W)
        self.folha = StringVar()
        self.btfolha = ttk.Button(self, text="Escolha a planilha", command=self.selecionar_folha)
        self.btfolha.grid(column=1, row=2, padx=col2, pady=2, sticky=W)
        # selecionar competencia
        self.labelcompet = ttk.Label(self, width=40, text="Competência:")
        self.labelcompet.grid(column=1, row=3, padx=10, pady=2, sticky=W)
        self.competencia = DateEntry(self, selectmode='day', year=self.hoje.year, month=self.hoje.month,
                                     day=self.hoje.day, locale='pt_BR')
        self.competencia.grid(column=1, row=3, padx=col2, pady=2, sticky=W)
        # selecionar data do pagamento
        self.labeldt = ttk.Label(self, width=60, text="Data do pgto:")
        self.labeldt.grid(column=1, row=4, padx=10, pady=2, sticky=W)
        self.pagamento = DateEntry(self, selectmode='day', year=self.hoje.year, month=self.hoje.month,
                                     day=self.hoje.day, locale='pt_BR')
        self.pagamento.grid(column=1, row=4, padx=col2, pady=2, sticky=W)

        # gerar folha da competencia selecionada
        self.botaogerar = ttk.Button(self, text="Anexar", command=lambda: [
            incluir_grade_email_holerite(self.folha.get(), self.competencia.get(), self.pagamento.get())])
        self.botaogerar.grid(column=1, row=11, padx=210, pady=1, sticky=W)

    def selecionar_folha(self):
        try:
            caminhoplan = tkinter.filedialog.askopenfilename(title='Planilha Folha')
            self.folha.set(str(caminhoplan))
        except ValueError:
            pass


class EnviarEmail(ttk.Frame):
    def __init__(self, container):
        super().__init__()
        self.hoje = datetime.datetime.today()
        self.anos = list(range(2000, self.hoje.year+1))
        self.meses = list(range(1, 13))
        sessions = sessionmaker(bind=engine)
        session = sessions()
        self.grupo = []
        pessoas = session.query(Colaborador).filter_by(desligamento=None).filter(Colaborador.ag.isnot(None)).filter(Colaborador.ag.isnot('None')).all()
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
        self.botaolancar = ttk.Button(self, text="Enviar", command=lambda: [])
        self.botaolancar.grid(column=1, row=8, padx=190, pady=1, sticky=W)


if __name__ == '__main__':
    app = MainApplication()
    app.mainloop()
