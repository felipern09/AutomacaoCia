from datetime import datetime
from src.controler.funcoes import emitir_certificados
import tkinter as tk
from tkcalendar import DateEntry
from tkinter import ttk
from tkinter import *
from src.models.models import Colaborador, engine
from sqlalchemy.orm import sessionmaker


class MainApplication(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Certificados - Cia BSB")
        self.geometry('661x300')
        self.img = PhotoImage(file='../models/static/imgs/Icone.png')
        self.iconphoto(False, self.img)
        self.columnconfigure(0, weight=5)
        self.rowconfigure(0, weight=5)
        for child in self.winfo_children():
            child.grid_configure(padx=1, pady=3)
        self.notebook = ttk.Notebook(self)
        self.Frame1 = Frame1(self.notebook)
        self.notebook.add(self.Frame1, text='Emissão de Certificados')
        self.notebook.pack(fill=BOTH)


class Frame1(ttk.Frame):
    def __init__(self, container):
        super().__init__()
        self.hoje = datetime.today()
        self.horas = IntVar()
        self.hrs = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15]
        # lista de nomes de funcionários com checkbox
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
        self.canvas = Canvas(self)
        self.canvas.pack(side=LEFT, fill=BOTH, expand=1)
        self.barraroll = ttk.Scrollbar(self, orient=VERTICAL, command=self.canvas.yview)
        self.barraroll.pack(side=LEFT, fill=Y)
        self.canvas.config(yscrollcommand=self.barraroll.set)
        self.canvas.bind('<Configure>', lambda e: self.canvas.config(scrollregion=self.canvas.bbox('all')))

        self.canvframe = Frame(self.canvas)
        self.canvas.create_window((0,0), window=self.canvframe, anchor='nw')

        #   loop for pessoa.nome com pesq em db
        #       i, enumerate(lista de nomes)
        #           self.label[i], row=i, self.checkbox[i]
        # definir nome do treinamento
        self.labelnome = ttk.Label(self.canvframe, width=120, text="Digite o nome do treinamento:")
        self.labelnome.grid(column=1, row=10, padx=25, pady=1, sticky=W)
        self.entrynome = ttk.Entry(self.canvframe, width=100)
        self.entrynome.grid(column=1, row=11, padx=25, pady=1, sticky=W)
        # definir data do trinamento
        self.labelcert = ttk.Label(self.canvframe, width=60, text="Data do treinamento:")
        self.labelcert.grid(column=1, row=12, padx=25, pady=1, sticky=W)
        self.entrycert = DateEntry(self.canvframe, selectmode='day', year=self.hoje.year, month=self.hoje.month,
                                     day=self.hoje.day, locale='pt_BR')
        self.entrycert.grid(column=1, row=12, padx=165, pady=1, sticky=W)
        # definir horas de duração
        self.labeldurac = ttk.Label(self.canvframe, width=60, text='Duração em horas:')
        self.labeldurac.grid(column=1, row=13, padx=25, pady=1, sticky=W)
        self.combodur = ttk.Combobox(self.canvframe, width=12, textvariable=self.horas, values=self.hrs)
        self.combodur.grid(column=1, row=13, padx=165, pady=1, sticky=W)
        self.labelparticp = ttk.Label(self.canvframe, width=60, text='Participantes:')
        self.labelparticp.grid(column=1, row=14, padx=25, pady=5, sticky=W)
        self.participantes = []

        def configparticip(event):
            widget = event.widget
            nome = widget.cget('text')
            if nome in self.participantes:
                self.participantes.remove(nome)
            else:
                self.participantes.append(nome)
            self.participantes.sort()

        for i, item in enumerate(self.nomes):
            var_name = f'var_{i}'
            value = IntVar()
            globals()[var_name] = value
            self.item = tk.Checkbutton(self.canvframe, text=item, variable=globals()[var_name])
            self.item.grid(column=1, row=i+16, padx=25, pady=1, sticky=W)
            self.item.bind('<Button-1>', configparticip)
        self.pst = IntVar(value=0)
        self.primsocrrt = tk.Checkbutton(self.canvframe, text='PS Terrestre', variable=self.pst)
        self.primsocrrt.grid(column=1, row=206, padx=500, pady=4, sticky=W)
        self.psa = IntVar(value=0)
        self.primsocrra = tk.Checkbutton(self.canvframe, text='PS Aquático', variable=self.psa)
        self.primsocrra.grid(column=1, row=207, padx=500, pady=4, sticky=W)
        self.botaocadastrar = ttk.Button(self.canvframe, width=20, text="Emitir certificados",
                                         command=lambda: [
                                             emitir_certificados(self.pst.get(), self.psa.get(),
                                                                 self.entrynome.get(),
                                                                 self.entrycert.get(),
                                                                 self.horas.get(),
                                                                 self.participantes)
                                         ])
        self.botaocadastrar.grid(column=1, row=208, padx=500, pady=1, sticky=W)


# implementar forma de aparecer lista de nomes com checkbox para adicionar esses nomes em uma lista como parametro da função

if __name__ == '__main__':
    app = MainApplication()
    app.mainloop()
