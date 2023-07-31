import tkinter as tk
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
        self.labelesreva = ttk.Label(self, width=90, text="Escreva abaixo a mensagem a ser enviada por whatsapp: ")
        self.labelesreva.grid(column=1, row=1, padx=25, pady=1, sticky=W)
        self.wpp = StringVar()
        self.entryemail = ScrolledText(self, wrap=tk.WORD)
        self.entryemail.grid(column=1, row=11, padx=25, pady=5, sticky=W)
        self.botaocadastrar = ttk.Button(self, width=20, text="Enviar whatsapp", command=lambda: [])
        self.botaocadastrar.grid(column=1, row=28, padx=520, pady=1, sticky=W)


if __name__ == '__main__':
    app = MainApplication()
    app.mainloop()
