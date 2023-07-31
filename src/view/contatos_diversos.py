import tkinter as tk
from tkcalendar import DateEntry
from datetime import datetime
import tkinter.filedialog
from src.controler.funcoes import cadastro_funcionario, salvar_docs_funcionarios, enviar_emails_funcionario, \
    cadastro_estagiario, cadastrar_autonomo, validar_pis
from openpyxl import load_workbook as l_w
from src.models.listas import horarios, cargos, departamentos, tipodecontrato
import tkinter.filedialog
from tkinter import ttk, scrolledtext
from tkinter import *
from tkinter.scrolledtext import ScrolledText


class MainApplication(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Contatos Diversos - Cia BSB")
        self.geometry('700x520')
        self.img = PhotoImage(file='../models/static/imgs/Icone.png')
        self.iconphoto(False, self.img)
        self.columnconfigure(0, weight=5)
        self.rowconfigure(0, weight=5)
        for child in self.winfo_children():
            child.grid_configure(padx=1, pady=3)
        self.notebook = ttk.Notebook(self)

        self.Frame1 = Frame1(self.notebook)
        self.Frame2 = Frame2(self.notebook)

        self.notebook.add(self.Frame1, text='Enviar e-mails')
        self.notebook.add(self.Frame2, text='Enviar Whatsapp')

        self.notebook.pack()


class Frame1(ttk.Frame):
    def __init__(self, container):
        super().__init__()

        self.hoje = datetime.today()
        self.geral = StringVar()
        self.caminho = StringVar()
        self.labelesreva = ttk.Label(self, width=90, text="Escreva abaixo a mensagem a ser enviada por e-mail: ")
        self.labelesreva.grid(column=1, row=1, padx=25, pady=1, sticky=W)
        self.email = StringVar()
        self.entryemail = ScrolledText(self, wrap=tk.WORD)
        self.entryemail.grid(column=1, row=11, padx=25, pady=5, sticky=W)
        self.botaocadastrar = ttk.Button(self, width=20, text="Enviar e-mail", command=lambda: [])
        self.botaocadastrar.grid(column=1, row=28, padx=520, pady=1, sticky=W)


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


if __name__ == '__main__':
    app = MainApplication()
    app.mainloop()
