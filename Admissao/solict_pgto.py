import tkinter as tk
import num2words as nw
from tkcalendar import DateEntry
from tkinter import ttk, messagebox
from datetime import datetime
import tkinter.filedialog
from tkinter import *
from models import Colaborador, engine
from sqlalchemy.orm import sessionmaker
from openpyxl import load_workbook as l_w
import config

a = 0.0
b = ''
c = ''


class MainApplication(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Solicitar pagamento")
        self.geometry('661x440')
        self.img = PhotoImage(file='./static/icone.png')
        self.iconphoto(False, self.img)
        self.notebook = ttk.Notebook(self)
        self.Frame1 = Pgto(self.notebook)
        self.Frame2 = PgtoFin(self.notebook)
        self.notebook.add(self.Frame1, text='Gerar Plan Itau')
        self.notebook.add(self.Frame2, text='Enviar Pedido ao Financeiro')
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
        pessoas = session.query(Colaborador).filter_by(desligamento='None').all()
        self.grupo = []
        for pessoa in pessoas:
            self.grupo.append(pessoa.nome)
            self.grupo.sort()
        pessoas2 = session.query(Colaborador).filter_by(desligamento=None).all()
        for pess in pessoas2:
            self.grupo.append(pess.nome)
            self.grupo.sort()
        self.nomes = self.grupo
        print(self.nomes)
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
        self.data.grid(column=1, row=28, padx=225, pady=1, sticky=W)
        self.botao = ttk.Button(self, width=35, text="Criar Planilha",command=lambda: [fazplanilha(
            self.nome1.get(), self.nome2.get(), self.nome3.get(), self.nome4.get(), self.nome5.get(), self.nome6.get(), self.nome7.get(), self.nome8.get(),
            self.nome9.get(),self.nome10.get(), self.nome11.get(), self.nome12.get(), self.nome13.get(), self.nome14.get(), self.nome15.get(), self.tipo1.get(),
            self.tipo2.get(), self.tipo3.get(), self.tipo4.get(), self.tipo5.get(), self.tipo6.get(), self.tipo7.get(), self.tipo8.get(), self.tipo9.get(),
            self.tipo10.get(),self.tipo11.get(), self.tipo12.get(), self.tipo13.get(), self.tipo14.get(), self.tipo15.get(), self.entryvalor1.get(),
            self.entryvalor2.get(),self.entryvalor3.get(), self.entryvalor4.get(), self.entryvalor5.get(),
            self.entryvalor6.get(),self.entryvalor7.get(), self.entryvalor8.get(), self.entryvalor9.get(),
            self.entryvalor10.get(),self.entryvalor11.get(), self.entryvalor12.get(), self.entryvalor13.get(),
            self.entryvalor14.get(),self.entryvalor15.get(), self.data.get()
        )])
        self.botao.grid(column=1, row=28, padx=380, pady=1, sticky=W)


class PgtoFin(ttk.Frame):
    def __init__(self, container):
        super().__init__()
        self.hoje = datetime.today()
        self.nome = StringVar()
        self.horario = StringVar()
        self.cargo = StringVar()
        self.departamento = StringVar()
        self.tipocontr = StringVar()
        self.nomesplan = []
        # aparecer entry para preencher matricula
        self.labelvalor = ttk.Label(self, width=20, text="Valor do pagamento:")
        self.labelvalor.grid(column=1, row=11, padx=25, pady=1, sticky=W)
        self.entryvalor = ttk.Entry(self, width=20)
        self.entryvalor.grid(column=1, row=11, padx=225, pady=1, sticky=W)
        # aparecer horario preenchido e dropdown para escolher horario
        self.labelquantos = ttk.Label(self, width=55, text="Quantos tipos de pgto: ")
        self.labelquantos.grid(column=1, row=14, padx=25, pady=1, sticky=W)
        self.comboquantos = ttk.Combobox(self, values=['1', '2', '3', '4', '5', '6'], textvariable=self.horario, width=50)
        self.comboquantos.grid(column=1, row=15, padx=225, pady=1, sticky=W)
        # aparecer entry para preencher salario
        self.labelsal = ttk.Label(self, width=20, text="Tipos: (colocar checkbox)")
        self.labelsal.grid(column=1, row=16, padx=25, pady=1, sticky=W)
        self.entrysal = ttk.Entry(self, width=20)
        self.entrysal.grid(column=1, row=16, padx=225, pady=1, sticky=W)
        # aparecer dropdown para escolher depto
        self.labelcalendario = ttk.Label(self, width=20, text="Data:")
        self.labelcalendario.grid(column=1, row=19, padx=25, pady=1, sticky=W)
        self.calendario = DateEntry(self, selectmode='day', year=self.hoje.year, month=self.hoje.month, day=self.hoje.day, locale='pt_BR')
        self.calendario.grid(column=1, row=19, padx=225, pady=1, sticky=W)
        # aparecer dropdown para escolher tipo_contr
        self.labelcompetencia = ttk.Label(self, width=20, text="Competencia:")
        self.labelcompetencia.grid(column=1, row=21, padx=25, pady=1, sticky=W)
        self.combocompetencia = ttk.Combobox(self, values=['mespassado', 'mesatual', 'mesquevem'], textvariable=self.tipocontr, width=50)
        self.combocompetencia.grid(column=1, row=21, padx=225, pady=1, sticky=W)
        self.edicao = IntVar()
        self.editar = ttk.Checkbutton(self, text='Editar cadastro feito manualmente.', variable=self.edicao)
        self.editar.grid(column=1, row=26, padx=26, pady=1, sticky=W)
        self.feitonde = IntVar()
        self.onde = ttk.Checkbutton(self, text='Cadastro realizado fora da Cia.', variable=self.feitonde)
        self.onde.grid(column=1, row=27, padx=26, pady=1, sticky=W)
        self.botao = ttk.Button(self, width=35, text="Cria capa e envia e-mail",
                   command=lambda: [confirmapgto()])
        self.botao.grid(column=1, row=28, padx=380, pady=1, sticky=W)


def confirmapgto(valor='10,00', tipo1='adiantamento', tipo2='0', tipo3='0', tipo4='0', tipo5='0', tipo6='0',
                 dia='03/07/2023', competencia='07/2023'):
    msg_box = tkinter.messagebox.askquestion('Confirma pagamento',
                                             'Tem certeza que deseja enviar o pagamento ao financeiro?\n'
                                             f'Valor: R$ {valor}\n'
                                             f'Data: {dia}\n'
                                             f'Tipo: {tipo1}\n'
                                             f'Competência: {competencia}\n',
                                             icon='warning')
    if msg_box == 'yes':
        tkinter.messagebox.showinfo('Pagamento enviado!', 'Pagamento enviado ao financeiro com sucesso!')
    else:
        tkinter.messagebox.showinfo('Editar dados', 'Pagamento não enviado. Edite os dados e tente novamente.')


def excrevervelor(total):
    # transformar algarismos do total em número por extenso e com reais e centavos
    reais, centavos = str(format(total, '.2f')).split('.')
    if int(reais) == 1:
        strReal = 'real'
    else:
        strReal = 'reais'
    if int(centavos) == 1:
        strCentavo = 'centavo'
    else:
        strCentavo = 'centavos'
    if int(reais) == 0:
        extenso = f'{nw.num2words(centavos, lang="pt_BR").capitalize()} {strCentavo}.'
    else:
        if int(centavos) == 0:
            extenso = f'{nw.num2words(reais, lang="pt_BR").capitalize()} {strReal}.'
        else:
            extenso = f'{nw.num2words(reais, lang="pt_BR").capitalize()} {strReal} e {nw.num2words(centavos, lang="pt_BR")} {strCentavo}.'
    return extenso


def fazplanilha(nome1, nome2, nome3, nome4, nome5, nome6, nome7, nome8, nome9, nome10, nome11, nome12, nome13, nome14, nome15,
                tipo1, tipo2, tipo3, tipo4, tipo5, tipo6, tipo7, tipo8, tipo9, tipo10, tipo11, tipo12, tipo13, tipo14, tipo15,
                val1, val2, val3, val4, val5, val6, val7, val8, val9, val10, val11, val12, val13, val14, val15, data):
    if val1 != '':
        valor1 = float(val1.replace(',','.'))
    else:
        valor1 = ''

    if val2 != '':
        valor2 = float(val2.replace(',','.'))
    else:
        valor2 = ''

    if val3 != '':
        valor3 = float(val3.replace(',','.'))
    else:
        valor3 = ''

    if val4 != '':
        valor4 = float(val4.replace(',','.'))
    else:
        valor4 = ''

    if val5 != '':
        valor5 = float(val5.replace(',','.'))
    else:
        valor5 = ''

    if val6 != '':
        valor6 = float(val6.replace(',','.'))
    else:
        valor6 = ''

    if val7 != '':
        valor7 = float(val7.replace(',','.'))
    else:
        valor7 = ''

    if val8 != '':
        valor8 = float(val8.replace(',','.'))
    else:
        valor8 = ''

    if val9 != '':
        valor9 = float(val9.replace(',','.'))
    else:
        valor9 = ''

    if val10 != '':
        valor10 = float(val10.replace(',','.'))
    else:
        valor10 = ''

    if val11 != '':
        valor11 = float(val11.replace(',','.'))
    else:
        valor11 = ''

    if val12 != '':
        valor12 = float(val12.replace(',','.'))
    else:
        valor12 = ''

    if val13 != '':
        valor13 = float(val13.replace(',','.'))
    else:
        valor13 = ''

    if val14 != '':
        valor14 = float(val14.replace(',','.'))
    else:
        valor14 = ''

    if val15 != '':
        valor15 = float(val15.replace(',','.'))
    else:
        valor15 = ''

    sessions = sessionmaker(bind=engine)
    session = sessions()
    dia = data.replace('/','.')
    tipos = {'':'','Salário': '1', 'Férias': '2', 'Vale Transporte': '3', 'Vale Alimentação': '4', 'Comissão': '5',
             '13º salário': '6', 'Bolsa Estágio': '7', 'Bônus': '8', 'Adiantamento Salarial': '9',
             'Rescisão': '10', 'Bolsa Auxílio': '11', 'Pensão Alimentícia': '12', 'Pgto em C/C': '13',
             'Remuneração': '14'}
    wb = l_w('Planilha Itau.xlsx', read_only=False)
    sh = wb['Planilha1']
    pessoa1 = session.query(Colaborador).filter_by(nome=nome1).first()
    pessoa2 = session.query(Colaborador).filter_by(nome=nome2).first()
    pessoa3 = session.query(Colaborador).filter_by(nome=nome3).first()
    pessoa4 = session.query(Colaborador).filter_by(nome=nome4).first()
    pessoa5 = session.query(Colaborador).filter_by(nome=nome5).first()
    pessoa6 = session.query(Colaborador).filter_by(nome=nome6).first()
    pessoa7 = session.query(Colaborador).filter_by(nome=nome7).first()
    pessoa8 = session.query(Colaborador).filter_by(nome=nome8).first()
    pessoa9 = session.query(Colaborador).filter_by(nome=nome9).first()
    pessoa10 = session.query(Colaborador).filter_by(nome=nome10).first()
    pessoa11 = session.query(Colaborador).filter_by(nome=nome11).first()
    pessoa12 = session.query(Colaborador).filter_by(nome=nome12).first()
    pessoa13 = session.query(Colaborador).filter_by(nome=nome13).first()
    pessoa14 = session.query(Colaborador).filter_by(nome=nome14).first()
    pessoa15 = session.query(Colaborador).filter_by(nome=nome15).first()

    sh['A1'].value = pessoa1.ag
    sh['A2'].value = pessoa2.ag
    sh['A3'].value = pessoa3.ag
    sh['A4'].value = pessoa4.ag
    sh['A5'].value = pessoa5.ag
    sh['A6'].value = pessoa6.ag
    sh['A7'].value = pessoa7.ag
    sh['A8'].value = pessoa8.ag
    sh['A9'].value = pessoa9.ag
    sh['A10'].value = pessoa10.ag
    sh['A11'].value = pessoa11.ag
    sh['A12'].value = pessoa12.ag
    sh['A13'].value = pessoa13.ag
    sh['A14'].value = pessoa14.ag
    sh['A15'].value = pessoa15.ag

    sh['B1'].value = pessoa1.conta
    sh['B2'].value = pessoa2.conta
    sh['B3'].value = pessoa3.conta
    sh['B4'].value = pessoa4.conta
    sh['B5'].value = pessoa5.conta
    sh['B6'].value = pessoa6.conta
    sh['B7'].value = pessoa7.conta
    sh['B8'].value = pessoa8.conta
    sh['B9'].value = pessoa9.conta
    sh['B10'].value = pessoa10.conta
    sh['B11'].value = pessoa11.conta
    sh['B12'].value = pessoa12.conta
    sh['B13'].value = pessoa13.conta
    sh['B14'].value = pessoa14.conta
    sh['B15'].value = pessoa15.conta

    sh['C1'].value = pessoa1.cdigito
    sh['C2'].value = pessoa2.cdigito
    sh['C3'].value = pessoa3.cdigito
    sh['C4'].value = pessoa4.cdigito
    sh['C5'].value = pessoa5.cdigito
    sh['C6'].value = pessoa6.cdigito
    sh['C7'].value = pessoa7.cdigito
    sh['C8'].value = pessoa8.cdigito
    sh['C9'].value = pessoa9.cdigito
    sh['C10'].value = pessoa10.cdigito
    sh['C11'].value = pessoa11.cdigito
    sh['C12'].value = pessoa12.cdigito
    sh['C13'].value = pessoa13.cdigito
    sh['C14'].value = pessoa14.cdigito
    sh['C15'].value = pessoa15.cdigito

    sh['D1'].value = pessoa1.nome
    sh['D2'].value = pessoa2.nome
    sh['D3'].value = pessoa3.nome
    sh['D4'].value = pessoa4.nome
    sh['D5'].value = pessoa5.nome
    sh['D6'].value = pessoa6.nome
    sh['D7'].value = pessoa7.nome
    sh['D8'].value = pessoa8.nome
    sh['D9'].value = pessoa9.nome
    sh['D10'].value = pessoa10.nome
    sh['D11'].value = pessoa11.nome
    sh['D12'].value = pessoa12.nome
    sh['D13'].value = pessoa13.nome
    sh['D14'].value = pessoa14.nome
    sh['D15'].value = pessoa15.nome

    sh['E1'].value = pessoa1.cpf
    sh['E2'].value = pessoa2.cpf
    sh['E3'].value = pessoa3.cpf
    sh['E4'].value = pessoa4.cpf
    sh['E5'].value = pessoa5.cpf
    sh['E6'].value = pessoa6.cpf
    sh['E7'].value = pessoa7.cpf
    sh['E8'].value = pessoa8.cpf
    sh['E9'].value = pessoa9.cpf
    sh['E10'].value = pessoa10.cpf
    sh['E11'].value = pessoa11.cpf
    sh['E12'].value = pessoa12.cpf
    sh['E13'].value = pessoa13.cpf
    sh['E14'].value = pessoa14.cpf
    sh['E15'].value = pessoa15.cpf

    sh['F1'].value = tipos[tipo1]
    sh['F2'].value = tipos[tipo2]
    sh['F3'].value = tipos[tipo3]
    sh['F4'].value = tipos[tipo4]
    sh['F5'].value = tipos[tipo5]
    sh['F6'].value = tipos[tipo6]
    sh['F7'].value = tipos[tipo7]
    sh['F8'].value = tipos[tipo8]
    sh['F9'].value = tipos[tipo9]
    sh['F10'].value = tipos[tipo10]
    sh['F11'].value = tipos[tipo11]
    sh['F12'].value = tipos[tipo12]
    sh['F13'].value = tipos[tipo13]
    sh['F14'].value = tipos[tipo14]
    sh['F15'].value = tipos[tipo15]

    sh['G1'].value = valor1
    sh['G2'].value = valor2
    sh['G3'].value = valor3
    sh['G4'].value = valor4
    sh['G5'].value = valor5
    sh['G6'].value = valor6
    sh['G7'].value = valor7
    sh['G8'].value = valor8
    sh['G9'].value = valor9
    sh['G10'].value = valor10
    sh['G11'].value = valor11
    sh['G12'].value = valor12
    sh['G13'].value = valor13
    sh['G14'].value = valor14
    sh['G15'].value = valor15

    sh.column_dimensions['A'].width = 6
    sh.column_dimensions['B'].width = 8
    sh.column_dimensions['C'].width = 4
    sh.column_dimensions['D'].width = 45
    sh.column_dimensions['E'].width = 20
    sh.column_dimensions['F'].width = 4
    sh.column_dimensions['G'].width = 16
    wb.save(f'Pagamento Itau {dia}.xlsx')
    config.x = 20


if __name__ == '__main__':
    app = MainApplication()
    app.mainloop()


# alterar o modelo de pagamento inserindo dados armazenados
# salvar como o modelo na pasta do mês e do tipo de pagamento solicitaçao de pgto-> Mes ->arquivos
# transformar modelo salvo na pasta "arquivos" mes para pdf e salvar na pasta definitiva solicitaçao de pgto-> Mes ->Tipo pgto-> pgto x dia y
# enviar o arquivo definitivo para o financeiro com corpo do e-mail definindo tipo pgto, valor total e data
