from openpyxl import load_workbook as l_w
from src.controler.f_uniformes import gerar_recibo_uniformes
import tkinter.filedialog
from tkinter import ttk
from tkinter import *
import tkinter.filedialog

root = Tk()
root.title("Recibos uniformes - Cia BSB")
img = PhotoImage(file='../models/static/imgs/Icone.png')
root.iconphoto(False, img)
root.geometry('480x350')
root.columnconfigure(0, weight=5)
root.rowconfigure(0, weight=5)

# implementar forma de generalizar controle
# 		o usuario deve apaenas fornecer os dados do uniforme selecionado e quem recebeu e o sistema controlar o estoque

my_notebook = ttk.Notebook(root)
my_notebook.pack()

geral = StringVar()
caminho = StringVar()
caminhoest = StringVar()
caminhoaut = StringVar()


def selecionarfunc():
    try:
        caminhoplan = tkinter.filedialog.askopenfilename(title='Planilha Funcionários')
        caminho.set(str(caminhoplan))
    except ValueError:
        pass


funcionario = Frame(my_notebook, width=10, height=20)
ttk.Label(funcionario, width=40, text="Escolher planilha relatório de uniformes").grid(column=1, row=1, padx=25, pady=1,
                                                                                       sticky=W)
ttk.Button(funcionario, text="Escolha a planilha", command=selecionarfunc).grid(column=1, row=1, padx=350, pady=1,
                                                                                sticky=W)

nome = StringVar()
cargo = StringVar()
cpf = StringVar()
tamanho = StringVar()
tamanho2 = StringVar()
genero = StringVar()
nomesplan = []
tamanhos = []
# aparecer dropdown com nomes da plan
labelnome = ttk.Label(funcionario, width=20, text="Nome:")
labelnome.grid(column=1, row=10, padx=25, pady=1, sticky=W)
combonome = ttk.Combobox(funcionario, values=nomesplan, textvariable=nome, width=38)
combonome.grid(column=1, row=10, padx=125, pady=1, sticky=W)
labelcargo = ttk.Label(funcionario, width=40, text="Cargo:")
labelcargo.grid(column=1, row=11, padx=25, pady=1, sticky=W)
labelcpf = ttk.Label(funcionario, width=20, text="CPF:")
labelcpf.grid(column=1, row=12, padx=25, pady=1, sticky=W)
labelgenero = ttk.Label(funcionario, width=20, text="Tipo:")
labelgenero.grid(column=1, row=13, padx=25, pady=1, sticky=W)

# aparecer entry para preencher tamanho
labeltam = ttk.Label(funcionario, width=20, text="Tamanho:")
labeltam.grid(column=1, row=14, padx=25, pady=1, sticky=W)
combotam = ttk.Combobox(funcionario, values=tamanhos, textvariable=tamanho, width=10)
combotam.grid(column=1, row=14, padx=125, pady=1, sticky=W)
maisum = IntVar()


def adicionartamanho():
    labeltam.config(text='Tamanho 1:')
    labeltam1 = ttk.Label(funcionario, width=20, text='Tamanho 2:')
    labeltam1.grid(column=1, row=15, padx=25, pady=1, sticky=W)
    combotam1 = ttk.Combobox(funcionario, values=tamanhos, textvariable=tamanho2, width=10)
    combotam1.grid(column=1, row=15, padx=125, pady=1, sticky=W)
    if labelgenero['text'] == 'Tipo: Feminino':
        lista_tamanhos = ['P', 'M', 'G', 'GG']
        combotam1.config(values=lista_tamanhos)
    if labelgenero['text'] == 'Tipo: Masculino':
        lista_tamanhos = ['M', 'G', 'XG', 'GG', 'XGG']
        combotam1.config(values=lista_tamanhos)


onde = ttk.Checkbutton(
    funcionario, text='Mais de um tamanho.', variable=maisum, onvalue=1, offvalue=0, command=adicionartamanho
)
onde.grid(column=1, row=17, padx=25, pady=1, sticky=W)


def mostravalores(event):
    nome = event.widget.get()
    num, name = nome.split(' - ')
    linha = int(num)
    planwb = l_w(caminho.get())
    plansh = planwb['Nomes']
    labelcargo.config(text=f'Cargo: {str(plansh[f"B{linha}"].value)}')
    labelcpf.config(text=f'CPF: {str(plansh[f"C{linha}"].value)}')
    labelgenero.config(text=f'Tipo: {str(plansh[f"D{linha}"].value)}')
    if labelgenero['text'] == 'Tipo: Feminino':
        lista_tamanhos = ['P', 'M', 'G', 'GG']
        combotam.config(values=lista_tamanhos)
    if labelgenero['text'] == 'Tipo: Masculino':
        lista_tamanhos = ['M', 'G', 'XG', 'GG', 'XGG']
        combotam.config(values=lista_tamanhos)


combonome.bind("<<ComboboxSelected>>", mostravalores)


def carregarfunc(local):
    planwb = l_w(local)
    plansh = planwb['Nomes']
    lista = []
    for x, pessoa in enumerate(plansh):
        lista.append(f'{x + 1} - {pessoa[0].value}')
    combonome.config(values=lista)


ttk.Button(funcionario, text="Carregar planilha", command=lambda: [carregarfunc(caminho.get())]).grid(column=1, row=9,
                                                                                                      padx=350, pady=25,
                                                                                                      sticky=W)
ttk.Button(funcionario, width=20, text="Gerar recibo",
           command=lambda: [
               gerar_recibo_uniformes(caminho.get(), nome.get(), str(labelcargo['text']), str(labelcpf['text']),
                                      str(labelgenero['text']), tamanho.get(), tamanho2.get())
           ]).grid(column=1, row=28, padx=325, pady=1, sticky=W)
funcionario.pack(fill='both', expand=0)
my_notebook.add(funcionario, text='Gerar Recibo de uniforme')
root.mainloop()
