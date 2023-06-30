import pyautogui as pa
import pyperclip as pp
from datetime import datetime
import time as t
from models import Colaborador, engine
from sqlalchemy.orm import sessionmaker
from openpyxl import load_workbook as l_w
from listas import horarios, cargos, departamentos, tipodecontrato, municipios
import os
import tkinter.filedialog
from tkinter import ttk, messagebox
from tkinter import *
import docx
import docx2pdf
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email import encoders
from dados_servd import em_rem, em_ti, em_if, k1
from difflib import SequenceMatcher


# this code automates the task process necessary for hiring a company employee. Made in tkinter, it was developed for
# the desktop to be integrated with other HR programs in the company
# it automates, through different processes, admissions of employees, interns or freelancers.
# each one with its specificity

root = Tk()
root.title("Atividades DP - Cia BSB")
img = PhotoImage(file='./static/icone.png')
root.iconphoto(False, img)
root.geometry('661x550')
root.columnconfigure(0, weight=5)
root.rowconfigure(0, weight=5)

for child in root.winfo_children():
    child.grid_configure(padx=1, pady=3)

my_notebook = ttk.Notebook(root)
my_notebook.pack()

menubar = Menu(root)
filemenu = Menu(menubar, tearoff=0)
filemenu.add_command(label="Configurações", command='')
filemenu.add_separator()
filemenu.add_command(label="Solicitar Suporte", command='')
filemenu.add_command(label="Sair", command=root.quit)
menubar.add_cascade(label="Configurações", menu=filemenu)
geral = StringVar()

Sessions = sessionmaker(bind=engine)
session = Sessions()
caminho = StringVar()


def cadastro_funcionario(caminho='', editar=0, ondestou=0, nome='', matricula='', admissao='',
                         horario='', salario='', cargo='', depto='', tipo_contr='',
                         hrsem='', hrmens=''):
    if caminho == '' or nome == '' or matricula == '' or admissao == '' or horario == '' or salario == '' or \
            cargo == '' or depto == '' or tipo_contr == '' or hrsem == '' or hrmens == '':
        tkinter.messagebox.showinfo(
            title='Erro de preenchimento',
            message='Preencha todos os campos antes de cadastrar o funcionário!'
        )
    else:
        wb = l_w(caminho, read_only=False)
        sh = wb['Respostas ao formulário 1']
        num, name = nome.strip().split(' - ')
        linha = int(num)

        # search for the highest compatibility between the city filled in the form and the cities in the lists to define codmunnas value
        pn = l_w('Cadastro de Funcionário (respostas).xlsx', read_only=False)
        sn = pn['Respostas ao formulário 1']
        x = 2
        while x <= len(sn['L']):
            est = str(sn[f'AJ{x}'].value)
            cidade = str(sn[f'L{x}'].value).title()
            lista = []
            dicion = {}
            for cid in municipios[est]:
                dicion[SequenceMatcher(None, cidade, cid).ratio()] = cid
                lista.append(SequenceMatcher(None, cidade, cid).ratio())
            codmunnas = municipios[str(sh[f'AJ{linha}'].value).upper().strip()][dicion[max(lista)]]

        # search for the highest compatibility between the city filled in the form and the cities in the lists to define codmunend value
        x = 2
        while x <= len(sn['L']):
            est = str(sn[f'AJ{x}'].value)
            cidade = str(sn[f'L{x}'].value).title()
            listaend = []
            dicionend = {}
            for cid in municipios[est]:
                dicionend[SequenceMatcher(None, cidade, cid).ratio()] = cid
                listaend.append(SequenceMatcher(None, cidade, cid).ratio())
            codmunend = municipios[str(sh[f'T{linha}'].value).upper().strip()][dicionend[max(listaend)]]

        lotacao = {'UNIDADE PARK SUL - QUALQUER DEPARTAMENTO':'0013','KIDS':'0010','MUSCULAÇÃO':'0007',
                 'ESPORTES E LUTAS':'0008','CROSSFIT':'0012','GINÁSTICA':'0006','GESTANTES':'0008','RECEPÇÃO':'0003',
                 'FINANCEIRO':'0001','TI':'0001','MARKETING':'0001','MANUTENÇÃO':'0004'}
        if editar == 0:
            if ondestou == 0:
                # Cadastro iniciado na Cia
                if linha:
                    pess = Colaborador(matricula=matricula, nome=name.upper(), admiss=admissao,
                                       nascimento=str(sh[f'D{linha}'].value),
                                       pis=str(int(sh[f'Y{linha}'].value)).zfill(11),
                                       cpf=str(int(sh[f'V{linha}'].value)).zfill(11),
                                       rg=str(int(sh[f'W{linha}'].value)),
                                       emissor=str(sh[f'X{linha}'].value), email=str(sh[f'B{linha}'].value),
                                       genero=str(sh[f'E{linha}'].value),
                                       estado_civil=str(sh[f'F{linha}'].value), cor=str(sh[f'G{linha}'].value),
                                       instru=str(sh[f'J{linha}'].value),
                                       nacional=str(sh[f'K{linha}'].value),
                                       cod_municipionas=codmunnas,                                       
                                       cid_nas=str(sh[f'L{linha}'].value), uf_nas=str(sh[f'AJ{linha}'].value),
                                       pai=str(sh[f'M{linha}'].value).upper(),
                                       mae=str(sh[f'N{linha}'].value).upper(), endereco=str(sh[f'O{linha}'].value),
                                       num=str(int(sh[f'P{linha}'].value)),
                                       bairro=str(sh[f'Q{linha}'].value), cep=str(int(sh[f'R{linha}'].value)),
                                       cidade=str(sh[f'S{linha}'].value),
                                       uf=str(sh[f'T{linha}'].value),
                                       cod_municipioend=codmunend,
                                       tel=str(int(sh[f'U{linha}'].value)),
                                       tit_eleit=str(sh[f'Z{linha}'].value), zona_eleit=str(sh[f'AA{linha}'].value),
                                       sec_eleit=str(sh[f'AB{linha}'].value),
                                       ctps=str(int(sh[f'AC{linha}'].value)), serie_ctps=str(sh[f'AD{linha}'].value),
                                       uf_ctps=str(sh[f'AE{linha}'].value),
                                       emiss_ctps=str(sh[f'AF{linha}'].value), depto=depto,
                                       cargo=cargo,
                                       horario=horario, salario=salario, tipo_contr=tipo_contr, hr_sem=hrsem,
                                       hr_mens=hrmens)
                    session.add(pess)
                    session.commit()
                    pessoa = session.query(Colaborador).filter_by(matricula=matricula).first()
                    pa.click(pa.center(pa.locateOnScreen('./static/Dexion.png')))
                    pa.press('alt'), pa.press('a'), pa.press('t'), t.sleep(2), pa.press(
                        'i'), t.sleep(5), pa.write(str(pessoa.matricula)), pa.press('enter'), t.sleep(20)
                    pp.copy(pessoa.nome), pa.hotkey('ctrl', 'v'), pa.press('tab'), pa.write(pessoa.cpf)
                    t.sleep(1), pa.press('tab', 3), pa.write(pessoa.genero), pa.press('tab'), pa.write(pessoa.cor)
                    t.sleep(1), pa.press('tab', 2), pa.write(pessoa.instru)
                    t.sleep(1), pa.press('tab'), pa.write(pessoa.estado_civil)
                    if str(pessoa.estado_civil) == '2 - Casado(a)':
                        pa.press('tab', 6)
                    else:
                        pa.press('tab', 5)
                    pa.write(datetime.strftime(datetime.strptime(pessoa.nascimento, '%Y-%m-%d %H:%M:%S'), '%d%m%Y'))
                    t.sleep(1), pa.press('tab'), pp.copy(pessoa.cid_nas), pa.hotkey('ctrl', 'v'), pa.press('tab')
                    t.sleep(1), pa.write(pessoa.uf_nas), pa.press('tab'), pa.write(pessoa.cod_municipionas), pa.press(
                        'tab')
                    t.sleep(1), pp.copy(pessoa.pai), pa.hotkey('ctrl', 'v'), pa.press('tab'), pa.write('105'), pa.press(
                        'tab')
                    t.sleep(1), pp.copy(pessoa.mae), pa.hotkey('ctrl', 'v'), pa.press('tab'), pa.write('105')
                    # #clique em documentos
                    pa.click(pa.center(pa.locateOnScreen('./static/Documentos.png')))
                    pa.press('tab'), pa.write(str(pessoa.rg)), pa.press('tab'), pa.write(
                        pessoa.emissor), pa.press('tab', 3), pa.write(pessoa.cod_municipioend),
                    pa.press('tab'), pa.write(pessoa.pis), pa.press('enter')
                    pa.press('tab'), pa.write(pessoa.tit_eleit), pa.press('tab'), pa.write(pessoa.zona_eleit), pa.press(
                        'tab'), pa.write(
                        pessoa.sec_eleit), pa.press('tab'), pa.write(pessoa.ctps)
                    pa.press('tab'), pa.write(pessoa.serie_ctps), pa.press('tab'), pa.write(pessoa.uf_ctps), pa.press(
                        'tab'), pa.write(datetime.strftime(datetime.strptime(pessoa.emiss_ctps, '%Y-%m-%d %H:%M:%S'),
                                                           '%d%m%Y')), pa.press('tab')
                    # #clique em endereço
                    pa.click(pa.center(pa.locateOnScreen('./static/Endereco.png')))
                    pa.press('tab', 2), pp.copy(pessoa.endereco), pa.hotkey('ctrl', 'v'), pa.press(
                        'tab'), pa.write(pessoa.num), pa.press('tab', 2)
                    pp.copy(pessoa.bairro), pa.hotkey('ctrl', 'v'), pa.press('tab'), pp.copy(pessoa.cidade), pa.hotkey(
                        'ctrl', 'v'), pa.press('tab'), pa.write(pessoa.uf)
                    pa.press('tab'), pa.write(pessoa.cep), pa.press('tab'), pa.write(pessoa.cod_municipioend), pa.press(
                        'tab'), pa.write(str(pessoa.tel)), pa.press('tab', 2), pa.write(pessoa.email)
                    # #clique em dados contratuais
                    pa.click(pa.center(pa.locateOnScreen('./static/Contratuais.png')))
                    pa.press('tab'), pa.write(str(pessoa.admiss).replace('/', '')), t.sleep(1)
                    pa.press('tab', 10), pa.write('2')
                    # #clique em Contrato de Experiência
                    try:
                        pa.click(pa.center(pa.locateOnScreen('./static/Experiencia.png')))
                    except:
                        t.sleep(5)
                        pa.click(pa.center(pa.locateOnScreen('./static/Experiencia.png')))
                    pa.press('tab'), pa.write('45'), pa.press('tab'), pa.write('45'), pa.press(
                        'tab'), pa.press(
                        'space'), pa.press('tab', 2), pa.write('003')
                    pa.press('tab'), pa.write(str(pessoa.matricula))
                    # #clique em Outros
                    try:
                        pa.click(pa.center(pa.locateOnScreen('./static/Outros.png')))
                    except:
                        t.sleep(5)
                        pa.click(pa.center(pa.locateOnScreen('./static/Outros.png')))
                    t.sleep(2), pa.write('CARGO GERAL')
                    # #clique em lupa de descrição de cargos
                    pa.click(pa.center(pa.locateOnScreen('./static/Lupa.png')))
                    t.sleep(2), pp.copy(pessoa.cargo), pa.hotkey('ctrl', 'v'), t.sleep(1.5), pa.press('enter', 2)
                    t.sleep(1), pa.press('tab'), pa.write(pessoa.salario), pa.press('tab')
                    if str(pessoa.tipo_contr) == 'Horista':
                        pa.press('1')
                    else:
                        pa.press('5')
                    pa.press('tab', 2), t.sleep(1), pa.write(str(pessoa.hr_sem)), pa.press('tab'), t.sleep(1)
                    pa.write(str(pessoa.hr_mens))
                    pa.press('tab', 5), pa.write('00395419000190'), pa.press('tab', 2), pa.write('5')
                    # #clique em eventos trabalhistas
                    pa.click(pa.center(pa.locateOnScreen('./static/EVTrab.png')))
                    t.sleep(1)
                    # #clique em lotação
                    pa.click(pa.center(pa.locateOnScreen('./static/Lotacoes.png')))
                    pa.press('tab'), pa.press('tab'), pa.write('i'), pa.write(str(pessoa.admiss).replace('/', ''))
                    t.sleep(1), pa.press('enter'), t.sleep(1)
                    pp.copy(lotacao[f'{pessoa.depto}']), pa.hotkey('ctrl', 'v'), pa.press('enter'), pa.write('f')
                    pa.press('tab'), pa.write('4')
                    pa.press('tab', 6), pa.write('i'), t.sleep(2), pa.press('tab'), pa.write(pessoa.horario)
                    t.sleep(3), pa.press('tab', 3), pa.press('enter'), t.sleep(3)
                    # #clique em cancelar novo registro de horario
                    pa.click(pa.center(pa.locateOnScreen('./static/Cancelarhor.png'))), t.sleep(2.5)
                    # #clique em salvar lotação
                    pa.click(pa.center(pa.locateOnScreen('./static/Salvarlot.png'))), t.sleep(1)
                    # #clique em fechar lotação
                    try:
                        pa.click(pa.center(pa.locateOnScreen('./static/Fecharlot.png'))), t.sleep(1)
                    except:
                        t.sleep(4)
                        pa.click(pa.center(pa.locateOnScreen('./static/Fecharlot.png')))
                    # #clique em Compatibilidade
                    pa.click(pa.center(pa.locateOnScreen('./static/Compatibilidade.png'))), t.sleep(1)
                    # #clique em Compatibilidade de novo
                    pa.click(pa.center(pa.locateOnScreen('./static/Compatibilidade2.png'))), t.sleep(1)
                    # #clique em CAGED
                    pa.click(pa.center(pa.locateOnScreen('./static/CAGED.png')))
                    pa.press('tab', 2), pa.write('20'), pa.press('tab'), pa.write('1'), t.sleep(0.5)
                    # #clique em RAIS
                    pa.click(pa.center(pa.locateOnScreen('./static/RAIS.png')))
                    pp.copy(pessoa.nome), pa.hotkey('ctrl', 'v'), pa.press('tab', 2), pa.write('2')
                    pa.press('tab'), pa.write('10')
                    # #clique em Salvar
                    pa.click(pa.center(pa.locateOnScreen('./static/Salvarcadastro.png'))), t.sleep(10)
                    # #clique em fechar novo cadastro
                    pa.click(pa.center(pa.locateOnScreen('./static/Fecharnovo1.png'))), t.sleep(2)
                    # #clique em fechar trabalhadores
                    pa.click(pa.center(pa.locateOnScreen('./static/Fechartrab1.png'))), t.sleep(0.5)
                    os.makedirs(r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e '
                                r'Férias\000 - Pastas Funcionais\00 - ATIVOS\{}'.format(pessoa.nome))
                    os.makedirs(r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e '
                                r'Férias\000 - Pastas Funcionais\00 - ATIVOS\{}\Atestados'.format(pessoa.nome))
                    os.makedirs(r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e '
                                r'Férias\000 - Pastas Funcionais\00 - ATIVOS\{}\Diversos'.format(pessoa.nome))
                    os.makedirs(r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e '
                                r'Férias\000 - Pastas Funcionais\00 - ATIVOS\{}\Contratuais'.format(pessoa.nome))
                    os.makedirs(r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e '
                                r'Férias\000 - Pastas Funcionais\00 - ATIVOS\{}\Férias'.format(pessoa.nome))
                    os.makedirs(r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e '
                                r'Férias\000 - Pastas Funcionais\00 - ATIVOS\{}\Pontos'.format(pessoa.nome))
                    os.makedirs(r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e '
                                r'Férias\000 - Pastas Funcionais\00 - ATIVOS\{}\Recibos'.format(pessoa.nome))
                    os.makedirs(r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e '
                                r'Férias\000 - Pastas Funcionais\00 - ATIVOS\{}\Rescisão'.format(pessoa.nome))
                    tkinter.messagebox.showinfo(
                        title='Cadastro ok!',
                        message='Cadastro realizado com sucesso!'
                    )
            else:
                # Cadastro iniciado em casa
                wb = l_w(caminho, read_only=False)
                sh = wb['Respostas ao formulário 1']
                num, name = nome.strip().split(' - ')
                linha = int(num)
                if linha:
                    pess = Colaborador(matricula=matricula, nome=name.upper(), admiss=admissao,
                                       nascimento=str(sh[f'D{linha}'].value),
                                       pis=str(int(sh[f'Y{linha}'].value)).zfill(11), cpf=str(int(sh[f'V{linha}'].value)).zfill(11),
                                       rg=str(int(sh[f'W{linha}'].value)),
                                       emissor=str(sh[f'X{linha}'].value), email=str(sh[f'B{linha}'].value),
                                       genero=str(sh[f'E{linha}'].value),
                                       estado_civil=str(sh[f'F{linha}'].value), cor=str(sh[f'G{linha}'].value),
                                       instru=str(sh[f'J{linha}'].value),
                                       nacional=str(sh[f'K{linha}'].value),
                                       cod_municipionas=codmunnas,
                                       cid_nas=str(sh[f'L{linha}'].value), uf_nas=str(sh[f'AJ{linha}'].value),
                                       pai=str(sh[f'M{linha}'].value).upper(),
                                       mae=str(sh[f'N{linha}'].value).upper(), endereco=str(sh[f'O{linha}'].value),
                                       num=str(int(sh[f'P{linha}'].value)),
                                       bairro=str(sh[f'Q{linha}'].value), cep=str(int(sh[f'R{linha}'].value)),
                                       cidade=str(sh[f'S{linha}'].value),
                                       uf=str(sh[f'T{linha}'].value),
                                       cod_municipioend=codmunend,
                                       tel=str(int(sh[f'U{linha}'].value)),
                                       tit_eleit=str(sh[f'Z{linha}'].value), zona_eleit=str(sh[f'AA{linha}'].value),
                                       sec_eleit=str(sh[f'AB{linha}'].value),
                                       ctps=str(int(sh[f'AC{linha}'].value)), serie_ctps=str(sh[f'AD{linha}'].value),
                                       uf_ctps=str(sh[f'AE{linha}'].value),
                                       emiss_ctps=str(sh[f'AF{linha}'].value), depto=depto,
                                       cargo=cargo,
                                       horario=horario, salario=salario, tipo_contr=tipo_contr, hr_sem=hrsem, hr_mens=hrmens)
                    session.add(pess)
                    session.commit()
                    pessoa = session.query(Colaborador).filter_by(matricula=matricula).first()
                    pa.click(pa.center(pa.locateOnScreen('./static/Dexion.png')))
                    pa.press('alt'), pa.press('a'), pa.press('t'), t.sleep(2), pa.press(
                        'i'), t.sleep(5), pa.write(str(pessoa.matricula)), pa.press('enter'), t.sleep(20)
                    pp.copy(pessoa.nome), pa.hotkey('ctrl', 'v'), pa.press('tab'), pa.write(pessoa.cpf)
                    t.sleep(1), pa.press('tab', 3), pa.write(pessoa.genero), pa.press('tab'), pa.write(pessoa.cor)
                    t.sleep(1), pa.press('tab', 2), pa.write(pessoa.instru)
                    t.sleep(1), pa.press('tab'), pa.write(pessoa.estado_civil)
                    if str(pessoa.estado_civil) == '2 - Casado(a)':
                        pa.press('tab', 6)
                    else:
                        pa.press('tab', 5)
                    pa.write(datetime.strftime(datetime.strptime(pessoa.nascimento, '%Y-%m-%d %H:%M:%S'), '%d%m%Y'))
                    t.sleep(1), pa.press('tab'), pp.copy(pessoa.cid_nas), pa.hotkey('ctrl', 'v'), pa.press('tab')
                    t.sleep(1), pa.write(pessoa.uf_nas), pa.press('tab'), pa.write(pessoa.cod_municipionas), pa.press('tab')
                    t.sleep(1), pp.copy(pessoa.pai), pa.hotkey('ctrl', 'v'), pa.press('tab'), pa.write('105'), pa.press('tab')
                    t.sleep(1), pp.copy(pessoa.mae), pa.hotkey('ctrl', 'v'), pa.press('tab'), pa.write('105')
                    # #clique em documentos
                    pa.click(pa.center(pa.locateOnScreen('./static/Documentos.png')))
                    pa.press('tab'), pa.write(str(pessoa.rg)), pa.press('tab'), pa.write(
                        pessoa.emissor), pa.press('tab', 3), pa.write(pessoa.cod_municipioend),
                    pa.press('tab'), pa.write(pessoa.pis), pa.press('enter')
                    pa.press('tab'), pa.write(pessoa.tit_eleit), pa.press('tab'), pa.write(pessoa.zona_eleit), pa.press('tab'), pa.write(
                        pessoa.sec_eleit), pa.press('tab'), pa.write(pessoa.ctps)
                    pa.press('tab'), pa.write(pessoa.serie_ctps), pa.press('tab'), pa.write(pessoa.uf_ctps), pa.press(
                        'tab'), pa.write(datetime.strftime(datetime.strptime(pessoa.emiss_ctps, '%Y-%m-%d %H:%M:%S'), '%d%m%Y')), pa.press('tab')
                    # #clique em endereço
                    pa.click(pa.center(pa.locateOnScreen('./static/Endereco.png')))
                    pa.press('tab', 2), pp.copy(pessoa.endereco), pa.hotkey('ctrl', 'v'), pa.press(
                        'tab'), pa.write(pessoa.num), pa.press('tab', 2)
                    pp.copy(pessoa.bairro), pa.hotkey('ctrl', 'v'), pa.press('tab'), pp.copy(pessoa.cidade), pa.hotkey(
                        'ctrl', 'v'), pa.press('tab'), pa.write(pessoa.uf)
                    pa.press('tab'), pa.write(pessoa.cep), pa.press('tab'), pa.write(pessoa.cod_municipioend), pa.press(
                        'tab'), pa.write(str(pessoa.tel)), pa.press('tab', 2), pa.write(pessoa.email)
                    # #clique em dados contratuais
                    pa.click(pa.center(pa.locateOnScreen('./static/Contratuais.png')))
                    pa.press('tab'), pa.write(str(pessoa.admiss).replace('/', '')), t.sleep(1)
                    pa.press('tab', 10), pa.write('2')
                    # #clique em Contrato de Experiência
                    try:
                        pa.click(pa.center(pa.locateOnScreen('./static/Experiencia.png')))
                    except:
                        t.sleep(5)
                        pa.click(pa.center(pa.locateOnScreen('./static/Experiencia.png')))
                    pa.press('tab'), pa.write('45'), pa.press('tab'), pa.write('45'), pa.press(
                        'tab'), pa.press(
                        'space'), pa.press('tab', 2), pa.write('003')
                    pa.press('tab'), pa.write(str(pessoa.matricula))
                    # #clique em Outros
                    try:
                        pa.click(pa.center(pa.locateOnScreen('./static/Outros.png')))
                    except:
                        t.sleep(5)
                        pa.click(pa.center(pa.locateOnScreen('./static/Outros.png')))
                    t.sleep(2), pa.write('CARGO GERAL')
                    # #clique em lupa de descrição de cargos
                    pa.click(pa.center(pa.locateOnScreen('./static/Lupa.png')))
                    t.sleep(2), pp.copy(pessoa.cargo), pa.hotkey('ctrl', 'v'), t.sleep(1.5), pa.press('enter',2)
                    t.sleep(1), pa.press('tab'), pa.write(pessoa.salario), pa.press('tab')
                    if str(pessoa.tipo_contr) == 'Horista':
                        pa.press('1')
                    else:
                        pa.press('5')
                    pa.press('tab', 2), t.sleep(1), pa.write(str(pessoa.hr_sem)), pa.press('tab'), t.sleep(1)
                    pa.write(str(pessoa.hr_mens))
                    pa.press('tab', 5), pa.write('00395419000190'), pa.press('tab', 2), pa.write('5')
                    # #clique em eventos trabalhistas
                    pa.click(pa.center(pa.locateOnScreen('./static/EVTrab.png')))
                    t.sleep(1)
                    # #clique em lotação
                    pa.click(pa.center(pa.locateOnScreen('./static/Lotacoes.png')))
                    pa.press('tab'), pa.press('tab'), pa.write('i'), pa.write(str(pessoa.admiss).replace('/',''))
                    t.sleep(1), pa.press('enter'), t.sleep(1)
                    pp.copy(lotacao[f'{pessoa.depto}']), pa.hotkey('ctrl', 'v'), pa.press('enter'), pa.write('f')
                    pa.press('tab'), pa.write('4')
                    pa.press('tab', 6), pa.write('i'), t.sleep(2), pa.press('tab'), pa.write(pessoa.horario)
                    t.sleep(3), pa.press('tab', 3), pa.press('enter'), t.sleep(3)
                    # #clique em cancelar novo registro de horario
                    pa.click(pa.center(pa.locateOnScreen('./static/Cancelarhor.png'))), t.sleep(2.5)
                    # #clique em salvar lotação
                    pa.click(pa.center(pa.locateOnScreen('./static/Salvarlot.png'))), t.sleep(1)
                    # #clique em fechar lotação
                    try:
                        pa.click(pa.center(pa.locateOnScreen('./static/Fecharlot.png'))), t.sleep(1)
                    except:
                        t.sleep(4)
                        pa.click(pa.center(pa.locateOnScreen('./static/Fecharlot.png')))
                    # #clique em Compatibilidade
                    pa.click(pa.center(pa.locateOnScreen('./static/Compatibilidade.png'))), t.sleep(1)
                    # #clique em Compatibilidade de novo
                    pa.click(pa.center(pa.locateOnScreen('./static/Compatibilidade2.png'))), t.sleep(1)
                    # #clique em CAGED
                    pa.click(pa.center(pa.locateOnScreen('./static/CAGED.png')))
                    pa.press('tab', 2), pa.write('20'), pa.press('tab'), pa.write('1'), t.sleep(0.5)
                    # #clique em RAIS
                    pa.click(pa.center(pa.locateOnScreen('./static/RAIS.png')))
                    pp.copy(pessoa.nome), pa.hotkey('ctrl', 'v'), pa.press('tab', 2), pa.write('2')
                    pa.press('tab'), pa.write('10')
                    # #clique em Salvar
                    pa.click(pa.center(pa.locateOnScreen('./static/Salvarcadastro.png'))), t.sleep(10)
                    # #clique em fechar novo cadastro
                    pa.click(pa.center(pa.locateOnScreen('./static/Fecharnovo1.png'))), t.sleep(2)
                    # #clique em fechar trabalhadores
                    pa.click(pa.center(pa.locateOnScreen('./static/Fechartrab1.png'))), t.sleep(0.5)

                    os.makedirs(r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e '
                                r'Férias\000 - Pastas Funcionais\00 - ATIVOS\{}'.format(pessoa.nome))
                    os.makedirs(r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e '
                                r'Férias\000 - Pastas Funcionais\00 - ATIVOS\{}\Atestados'.format(pessoa.nome))
                    os.makedirs(r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e '
                                r'Férias\000 - Pastas Funcionais\00 - ATIVOS\{}\Diversos'.format(pessoa.nome))
                    os.makedirs(r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e '
                                r'Férias\000 - Pastas Funcionais\00 - ATIVOS\{}\Contratuais'.format(pessoa.nome))
                    os.makedirs(r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e '
                                r'Férias\000 - Pastas Funcionais\00 - ATIVOS\{}\Férias'.format(pessoa.nome))
                    os.makedirs(r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e '
                                r'Férias\000 - Pastas Funcionais\00 - ATIVOS\{}\Pontos'.format(pessoa.nome))
                    os.makedirs(r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e '
                                r'Férias\000 - Pastas Funcionais\00 - ATIVOS\{}\Recibos'.format(pessoa.nome))
                    os.makedirs(r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e '
                                r'Férias\000 - Pastas Funcionais\00 - ATIVOS\{}\Rescisão'.format(pessoa.nome))
                    tkinter.messagebox.showinfo(
                        title='Cadastro ok!',
                        message='Cadastro realizado com sucesso!'
                    )

        else:
            if ondestou == 0:
                # Cadastro EDITADO na Cia
                wb = l_w(caminho, read_only=False)
                num, name = nome.strip().split(' - ')
                linha = int(num)
                if linha:
                    pessoa = session.query(Colaborador).filter_by(matricula=matricula).first()
                    pa.click(pa.center(pa.locateOnScreen('./static/Dexion.png')))
                    pa.press('alt'), pa.press('a'), pa.press('t'), t.sleep(2), pa.press(
                        'a'), t.sleep(5), pa.write(str(pessoa.matricula)), pa.press('enter'), t.sleep(15)
                    pp.copy(pessoa.nome), pa.hotkey('ctrl', 'v'), pa.press('tab'), pa.write(pessoa.cpf)
                    pa.press('tab', 3), pa.write(pessoa.genero), pa.press('tab'), pa.write(pessoa.cor)
                    pa.press('tab', 2), pa.write(pessoa.instru)
                    pa.press('tab'), pa.write(pessoa.estado_civil)
                    if str(pessoa.estado_civil) == '2 - Casado(a)':
                        pa.press('tab', 6)
                    else:
                        pa.press('tab', 5)
                    pa.write(datetime.strftime(datetime.strptime(pessoa.nascimento, '%Y-%m-%d %H:%M:%S'), '%d%m%Y'))
                    pa.press('tab'), pp.copy(pessoa.cid_nas), pa.hotkey('ctrl', 'v'), pa.press('tab')
                    pa.write(pessoa.uf_nas), pa.press('tab'), pa.write(pessoa.cod_municipionas), pa.press('tab')
                    pp.copy(pessoa.pai), pa.hotkey('ctrl', 'v'), pa.press('tab'), pa.write('105'), pa.press('tab')
                    pp.copy(pessoa.mae), pa.hotkey('ctrl', 'v'), pa.press('tab'), pa.write('105')
                    # #clique em documentos
                    pa.click(pa.center(pa.locateOnScreen('./static/Documentos.png')))
                    pa.press('tab'), pa.write(str(pessoa.rg)), pa.press('tab'), pa.write(
                        pessoa.emissor), pa.press('tab', 3), pa.write(pessoa.cod_municipioend),
                    pa.press('tab'), pa.write(pessoa.pis), pa.press('enter')
                    pa.press('tab'), pa.write(pessoa.tit_eleit), pa.press('tab'), pa.write(pessoa.zona_eleit), pa.press('tab'), pa.write(
                        pessoa.sec_eleit), pa.press('tab'), pa.write(pessoa.ctps)
                    pa.press('tab'), pa.write(pessoa.serie_ctps), pa.press('tab'), pa.write(pessoa.uf_ctps), pa.press(
                        'tab'), pa.write(datetime.strftime(datetime.strptime(pessoa.emiss_ctps, '%Y-%m-%d %H:%M:%S'), '%d%m%Y')), pa.press('tab')
                    # #clique em endereço
                    pa.click(pa.center(pa.locateOnScreen('./static/Endereco.png')))
                    pa.press('tab', 2), pp.copy(pessoa.endereco), pa.hotkey('ctrl', 'v'), pa.press(
                        'tab'), pa.write(pessoa.num), pa.press('tab', 2)
                    pp.copy(pessoa.bairro), pa.hotkey('ctrl', 'v'), pa.press('tab'), pp.copy(pessoa.cidade), pa.hotkey(
                        'ctrl', 'v'), pa.press('tab'), pa.write(pessoa.uf)
                    pa.press('tab'), pa.write(pessoa.cep), pa.press('tab'), pa.write(pessoa.cod_municipioend), pa.press(
                        'tab'), pa.write(str(pessoa.tel)), pa.press('tab', 2), pa.write(pessoa.email)
                    # #clique em dados contratuais
                    pa.click(pa.center(pa.locateOnScreen('./static/Contratuais.png')))
                    pa.press('tab'), pa.write(str(pessoa.admiss).replace('/', '')), t.sleep(1)
                    pa.press('tab', 10), pa.write('2')
                    # #clique em Contrato de Experiência
                    try:
                        pa.click(pa.center(pa.locateOnScreen('./static/Experiencia.png')))
                    except:
                        t.sleep(5)
                        pa.click(pa.center(pa.locateOnScreen('./static/Experiencia.png')))
                    pa.press('tab'), pa.write('45'), pa.press('tab'), pa.write('45'), pa.press(
                        'tab'), pa.press(
                        'space'), pa.press('tab', 2), pa.write('003')
                    pa.press('tab'), pa.write(str(pessoa.matricula))
                    # #clique em Outros
                    try:
                        pa.click(pa.center(pa.locateOnScreen('./static/Outros.png')))
                    except:
                        t.sleep(5)
                        pa.click(pa.center(pa.locateOnScreen('./static/Outros.png')))
                    t.sleep(2), pa.write('CARGO GERAL')
                    # #clique em lupa de descrição de cargos
                    pa.click(pa.center(pa.locateOnScreen('./static/Lupa.png')))
                    t.sleep(2), pp.copy(pessoa.cargo), pa.hotkey('ctrl', 'v'), t.sleep(1.5), pa.press('enter',2)
                    t.sleep(1), pa.press('tab'), pa.write(pessoa.salario), pa.press('tab')
                    if str(pessoa.tipo_contr) == 'Horista':
                        pa.press('1')
                    else:
                        pa.press('5')
                    pa.press('tab', 2), t.sleep(1), pa.write(str(pessoa.hr_sem)), pa.press('tab'), t.sleep(1)
                    pa.write(str(pessoa.hr_mens))
                    pa.press('tab', 5), pa.write('00395419000190'), pa.press('tab', 2), pa.write('5')
                    # #clique em Compatibilidade
                    pa.click(pa.center(pa.locateOnScreen('./static/Compatibilidade1.png'))), t.sleep(1)
                    # #clique em Compatibilidade de novo
                    pa.click(pa.center(pa.locateOnScreen('./static/Compatibilidade2.png'))), t.sleep(1)
                    # #clique em CAGED
                    pa.click(pa.center(pa.locateOnScreen('./static/CAGED.png')))
                    pa.press('tab', 2), pa.write('20'), pa.press('tab'), pa.write('1'), t.sleep(0.5)
                    # #clique em RAIS
                    pa.click(pa.center(pa.locateOnScreen('./static/RAIS.png')))
                    pp.copy(pessoa.nome), pa.hotkey('ctrl', 'v'), pa.press('tab', 2), pa.write('2')
                    pa.press('tab'), pa.write('10')
                    # #clique em Salvar
                    pa.click(pa.center(pa.locateOnScreen('./static/Salvarcadastro.png'))), t.sleep(10)
                    # #clique em fechar novo cadastro
                    pa.click(pa.center(pa.locateOnScreen('./static/Fecharnovo1.png'))), t.sleep(2)
                    # #clique em fechar trabalhadores
                    pa.click(pa.center(pa.locateOnScreen('./static/Fechartrab1.png'))), t.sleep(0.5)

                    os.makedirs(r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e '
                                r'Férias\000 - Pastas Funcionais\00 - ATIVOS\{}'.format(pessoa.nome))
                    os.makedirs(r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e '
                                r'Férias\000 - Pastas Funcionais\00 - ATIVOS\{}\Atestados'.format(pessoa.nome))
                    os.makedirs(r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e '
                                r'Férias\000 - Pastas Funcionais\00 - ATIVOS\{}\Diversos'.format(pessoa.nome))
                    os.makedirs(r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e '
                                r'Férias\000 - Pastas Funcionais\00 - ATIVOS\{}\Contratuais'.format(pessoa.nome))
                    os.makedirs(r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e '
                                r'Férias\000 - Pastas Funcionais\00 - ATIVOS\{}\Férias'.format(pessoa.nome))
                    os.makedirs(r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e '
                                r'Férias\000 - Pastas Funcionais\00 - ATIVOS\{}\Pontos'.format(pessoa.nome))
                    os.makedirs(r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e '
                                r'Férias\000 - Pastas Funcionais\00 - ATIVOS\{}\Recibos'.format(pessoa.nome))
                    os.makedirs(r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e '
                                r'Férias\000 - Pastas Funcionais\00 - ATIVOS\{}\Rescisão'.format(pessoa.nome))
                    tkinter.messagebox.showinfo(title='Cadastro ok!',
                                                message='Cadastro editado com sucesso!')
                else:
                    # Cadastro EDITADO em casa
                    wb = l_w(caminho, read_only=False)
                    num, name = nome.strip().split(' - ')
                    linha = int(num)
                    if linha:
                        pessoa = session.query(Colaborador).filter_by(matricula=matricula).first()
                        pa.click(pa.center(pa.locateOnScreen('./static/Dexion.png')))
                        pa.press('alt'), pa.press('a'), pa.press('t'), t.sleep(2), pa.press(
                            'a'), t.sleep(5), pa.write(str(pessoa.matricula)), pa.press('enter'), t.sleep(20)
                        pp.copy(pessoa.nome), pa.hotkey('ctrl', 'v'), pa.press('tab'), pa.write(pessoa.cpf)
                        t.sleep(1), pa.press('tab', 3), pa.write(pessoa.genero), pa.press('tab'), pa.write(pessoa.cor)
                        t.sleep(1), pa.press('tab', 2), pa.write(pessoa.instru)
                        t.sleep(1), pa.press('tab'), pa.write(pessoa.estado_civil)
                        if str(pessoa.estado_civil) == '2 - Casado(a)':
                            pa.press('tab', 6)
                        else:
                            pa.press('tab', 5)
                        pa.write(datetime.strftime(datetime.strptime(pessoa.nascimento, '%Y-%m-%d %H:%M:%S'), '%d%m%Y'))
                        t.sleep(1), pa.press('tab'), pp.copy(pessoa.cid_nas), pa.hotkey('ctrl', 'v'), pa.press('tab')
                        t.sleep(1), pa.write(pessoa.uf_nas), pa.press('tab'), pa.write(
                            pessoa.cod_municipionas), pa.press('tab')
                        t.sleep(1), pp.copy(pessoa.pai), pa.hotkey('ctrl', 'v'), pa.press('tab'), pa.write(
                            '105'), pa.press('tab')
                        t.sleep(1), pp.copy(pessoa.mae), pa.hotkey('ctrl', 'v'), pa.press('tab'), pa.write('105')
                        # #clique em documentos
                        pa.click(pa.center(pa.locateOnScreen('./static/Documentos.png')))
                        pa.press('tab'), pa.write(str(pessoa.rg)), pa.press('tab'), pa.write(
                            pessoa.emissor), pa.press('tab', 3), pa.write(pessoa.cod_municipioend),
                        pa.press('tab'), pa.write(pessoa.pis), pa.press('enter')
                        pa.press('tab'), pa.write(pessoa.tit_eleit), pa.press('tab'), pa.write(
                            pessoa.zona_eleit), pa.press('tab'), pa.write(
                            pessoa.sec_eleit), pa.press('tab'), pa.write(pessoa.ctps)
                        pa.press('tab'), pa.write(pessoa.serie_ctps), pa.press('tab'), pa.write(
                            pessoa.uf_ctps), pa.press(
                            'tab'), pa.write(
                            datetime.strftime(datetime.strptime(pessoa.emiss_ctps, '%Y-%m-%d %H:%M:%S'),
                                              '%d%m%Y')), pa.press('tab')
                        # #clique em endereço
                        pa.click(pa.center(pa.locateOnScreen('./static/Endereco.png')))
                        pa.press('tab', 2), pp.copy(pessoa.endereco), pa.hotkey('ctrl', 'v'), pa.press(
                            'tab'), pa.write(pessoa.num), pa.press('tab', 2)
                        pp.copy(pessoa.bairro), pa.hotkey('ctrl', 'v'), pa.press('tab'), pp.copy(
                            pessoa.cidade), pa.hotkey(
                            'ctrl', 'v'), pa.press('tab'), pa.write(pessoa.uf)
                        pa.press('tab'), pa.write(pessoa.cep), pa.press('tab'), pa.write(
                            pessoa.cod_municipioend), pa.press(
                            'tab'), pa.write(str(pessoa.tel)), pa.press('tab', 2), pa.write(pessoa.email)
                        # #clique em dados contratuais
                        pa.click(pa.center(pa.locateOnScreen('./static/Contratuais.png')))
                        pa.press('tab'), pa.write(str(pessoa.admiss).replace('/', '')), t.sleep(1)
                        pa.press('tab', 10), pa.write('2')
                        # #clique em Contrato de Experiência
                        try:
                            pa.click(pa.center(pa.locateOnScreen('./static/Experiencia.png')))
                        except:
                            t.sleep(5)
                            pa.click(pa.center(pa.locateOnScreen('./static/Experiencia.png')))
                        pa.press('tab'), pa.write('45'), pa.press('tab'), pa.write('45'), pa.press(
                            'tab'), pa.press(
                            'space'), pa.press('tab', 2), pa.write('003')
                        pa.press('tab'), pa.write(str(pessoa.matricula))
                        # #clique em Outros
                        try:
                            pa.click(pa.center(pa.locateOnScreen('./static/Outros.png')))
                        except:
                            t.sleep(5)
                            pa.click(pa.center(pa.locateOnScreen('./static/Outros.png')))
                        t.sleep(2), pa.write('CARGO GERAL')
                        # #clique em lupa de descrição de cargos
                        pa.click(pa.center(pa.locateOnScreen('./static/Lupa.png')))
                        t.sleep(2), pp.copy(pessoa.cargo), pa.hotkey('ctrl', 'v'), t.sleep(1.5), pa.press('enter', 2)
                        t.sleep(1), pa.press('tab'), pa.write(pessoa.salario), pa.press('tab')
                        if str(pessoa.tipo_contr) == 'Horista':
                            pa.press('1')
                        else:
                            pa.press('5')
                        pa.press('tab', 2), t.sleep(1), pa.write(str(pessoa.hr_sem)), pa.press('tab'), t.sleep(1)
                        pa.write(str(pessoa.hr_mens))
                        pa.press('tab', 5), pa.write('00395419000190'), pa.press('tab', 2), pa.write('5')
                        # #clique em Compatibilidade
                        pa.click(pa.center(pa.locateOnScreen('./static/Compatibilidade.png'))), t.sleep(1)
                        # #clique em Compatibilidade de novo
                        pa.click(pa.center(pa.locateOnScreen('./static/Compatibilidade2.png'))), t.sleep(1)
                        # #clique em CAGED
                        pa.click(pa.center(pa.locateOnScreen('./static/CAGED.png')))
                        pa.press('tab', 2), pa.write('20'), pa.press('tab'), pa.write('1'), t.sleep(0.5)
                        # #clique em RAIS
                        pa.click(pa.center(pa.locateOnScreen('./static/RAIS.png')))
                        pp.copy(pessoa.nome), pa.hotkey('ctrl', 'v'), pa.press('tab', 2), pa.write('2')
                        pa.press('tab'), pa.write('10')
                        # #clique em Salvar
                        pa.click(pa.center(pa.locateOnScreen('./static/Salvarcadastro.png'))), t.sleep(10)
                        # #clique em fechar novo cadastro
                        pa.click(pa.center(pa.locateOnScreen('./static/Fecharnovo1.png'))), t.sleep(2)
                        # #clique em fechar trabalhadores
                        pa.click(pa.center(pa.locateOnScreen('./static/Fechartrab1.png'))), t.sleep(0.5)

                        os.makedirs(
                            r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e '
                            r'Férias\000 - Pastas Funcionais\00 - ATIVOS\{}'.format(pessoa.nome))
                        os.makedirs(
                            r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e '
                            r'Férias\000 - Pastas Funcionais\00 - ATIVOS\{}\Atestados'.format(pessoa.nome))
                        os.makedirs(
                            r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e '
                            r'Férias\000 - Pastas Funcionais\00 - ATIVOS\{}\Diversos'.format(pessoa.nome))
                        os.makedirs(
                            r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e '
                            r'Férias\000 - Pastas Funcionais\00 - ATIVOS\{}\Contratuais'.format(pessoa.nome))
                        os.makedirs(
                            r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e '
                            r'Férias\000 - Pastas Funcionais\00 - ATIVOS\{}\Férias'.format(pessoa.nome))
                        os.makedirs(
                            r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e '
                            r'Férias\000 - Pastas Funcionais\00 - ATIVOS\{}\Pontos'.format(pessoa.nome))
                        os.makedirs(
                            r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e '
                            r'Férias\000 - Pastas Funcionais\00 - ATIVOS\{}\Recibos'.format(pessoa.nome))
                        os.makedirs(
                            r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e '
                            r'Férias\000 - Pastas Funcionais\00 - ATIVOS\{}\Rescisão'.format(pessoa.nome))
                        tkinter.messagebox.showinfo(
                            title='Cadastro ok!',
                            message='Cadastro editado com sucesso!'
                        )


def salvadocsfunc(matricula):
    pessoa = session.query(Colaborador).filter_by(matricula=matricula).first()
    pa.click(pa.center(pa.locateOnScreen('./static/Dexion.png')))
    p_pessoa = r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e Férias' \
               r'\000 - Pastas Funcionais\00 - ATIVOS\{}'.format(pessoa.nome)
    p_atestado = r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e Férias' \
                 r'\000 - Pastas Funcionais\00 - ATIVOS\{}\Atestado'.format(pessoa.nome)
    p_contr = r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e Férias' \
              r'\000 - Pastas Funcionais\00 - ATIVOS\{}\Contratuais'.format(pessoa.nome)
    p_diversos = r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e Férias' \
                 r'\000 - Pastas Funcionais\00 - ATIVOS\{}\Diversos'.format(pessoa.nome)
    p_ferias = r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e Férias' \
               r'\000 - Pastas Funcionais\00 - ATIVOS\{}\Férias'.format(pessoa.nome)
    p_ponto = r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e Férias' \
              r'\000 - Pastas Funcionais\00 - ATIVOS\{}\Pontos'.format(pessoa.nome)
    p_rec = r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e Férias' \
            r'\000 - Pastas Funcionais\00 - ATIVOS\{}\Recibos'.format(pessoa.nome)
    p_rescisao = r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e Férias' \
                 r'\000 - Pastas Funcionais\00 - ATIVOS\{}\Rescisão'.format(pessoa.nome)
    p_ac = r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e Férias' \
           r'\000 - Pastas Funcionais\00 - ATIVOS\Modelo\AC Modelo.docx'
    p_abconta = r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e Férias' \
                r'\000 - Pastas Funcionais\00 - ATIVOS\Modelo\Abertura Conta MODELO.docx'
    p_fcadas = r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e Férias' \
               r'\000 - Pastas Funcionais\00 - ATIVOS\Modelo\Ficha Cadastral MODELO.docx'
    p_recibos = r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e Férias' \
                r'\000 - Pastas Funcionais\00 - ATIVOS\Modelo\Recibo Crachá e Uniformes MODELO.docx'
    p_recibovt = r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e Férias' \
                 r'\000 - Pastas Funcionais\00 - ATIVOS\Modelo\Recibo VT MODELO.docx'
    p_codetic = r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e Férias' \
                r'\000 - Pastas Funcionais\00 - ATIVOS\Modelo\Cod Etica MODELO.docx'
    ps_contr = r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e Férias' \
               r'\000 - Pastas Funcionais\00 - ATIVOS\{}\Contratuais\Contrato.pdf'.format(pessoa.nome)
    ps_acordo = r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e Férias' \
                r'\000 - Pastas Funcionais\00 - ATIVOS\{}\Contratuais\Acordo Banco de Horas.pdf'.format(pessoa.nome)
    ps_recctps = r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e Férias' \
                 r'\000 - Pastas Funcionais\00 - ATIVOS\{}\Contratuais\Recibo de Entrega e Dev CTPS.pdf'.format(
        pessoa.nome)
    ps_anotctps = r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e Férias' \
                  r'\000 - Pastas Funcionais\00 - ATIVOS\{}\Contratuais\Anotacoes CTPS.pdf'.format(pessoa.nome)
    ps_termovt = r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e Férias' \
                 r'\000 - Pastas Funcionais\00 - ATIVOS\{}\Contratuais\Termo Opcao VT.pdf'.format(pessoa.nome)
    ps_contrato = r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e Férias' \
                  r'\000 - Pastas Funcionais\00 - ATIVOS\{}\Contratuais\Contrato de Trabalho.pdf'.format(pessoa.nome)
    ps_ficha = r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e Férias' \
               r'\000 - Pastas Funcionais\00 - ATIVOS\{}\Contratuais\Folha de Registro.pdf'.format(pessoa.nome)
    try:
        os.mkdir(p_pessoa)
        os.mkdir(p_atestado)
        os.mkdir(p_contr)
        os.mkdir(p_diversos)
        os.mkdir(p_ferias)
        os.mkdir(p_ponto)
        os.mkdir(p_rec)
        os.mkdir(p_rescisao)
    except:
        pass

    lotacao = {
        'Unidade Park Sul - Qualquer Departamento': ['0013', 'Thais Feitosa', 'thais.morais@ciaathletica.com.br',
                                                     'Líder Park Sul'],
        'Kids': ['0010', 'Cindy Stefanie', 'cindy.neves@ciaathletica.com.br', 'Líder Kids'],
        'Musculação': ['0007', 'Thaís Feitosa', 'thais.morais@ciaathletica.com.br', 'Líder Musculação'],
        'Esportes e Lutas': ['0008', 'Morgana Rossini', 'morganalourenco@yahoo.com.br', 'Líder Natação'],
        'Crossfit': ['0012', 'Guilherme Salles', 'gmoreirasalles@gmail.com', 'Líder Crossfit'],
        'Ginástica': ['0006', 'Morgana Rossini', 'morganalourenco@yahoo.com.br', 'Líder Ginástica'],
        'Gestantes': ['0006', 'Filipe Feijó', 'filipe.feijo@ciaathletica.com.br', 'Líder Ginástica'],
        'Recepção': ['0003', 'Paulo Renato', 'paulo.simoes@ciaathletica.com.br', 'Gerente Vendas'],
        'Administrativo': ['0001', 'Felipe Rodrigues', 'felipe.rodrigues@ciaathletica.com.br', 'Gerente RH'],
        'Manutenção': ['0004', 'José Aparecido', 'aparecido.grota@ciaathletica.com.br', 'Gerente Manutenção'],
    }

    abert_c = docx.Document(p_abconta)
    ac = docx.Document(p_ac)
    fch_c = docx.Document(p_fcadas)
    recibos = docx.Document(p_recibos)
    recibovt = docx.Document(p_recibovt)
    codetic = docx.Document(p_codetic)

    # # imprimir recibo entrega e devolução de ctps
    pa.press('alt'), pa.press('r'), pa.press('a'), pa.press('r'), pa.press('e'), pa.press('tab'), pa.write(str(
        pessoa.matricula))
    pa.press('tab', 3), pa.write(str(pessoa.admiss).replace('/', '')), pa.press('tab'), t.sleep(0.5), pa.press('space')
    t.sleep(0.5), pa.press('tab'), pa.write(str(pessoa.admiss).replace('/', '')), pa.press('tab', 2)
    t.sleep(1), pa.press('enter'), t.sleep(2)

    # # clique no endereço de salvamento do recibo
    pa.click(pa.center(pa.locateOnScreen('./static/salvar.png'))), t.sleep(1)
    pp.copy(ps_recctps), pa.hotkey('ctrl', 'v'), t.sleep(0.5)
    pa.press('tab', 2), t.sleep(0.5), pa.press('enter')
    t.sleep(5)
    # # clique para fechar recibo ctps
    pa.click(pa.center(pa.locateOnScreen('./static/fechar_janela.png'))), t.sleep(0.5)
    pa.click(pa.center(pa.locateOnScreen('./static/fechar_janela.png')))

    # # Imprimir Acordo de Banco de horas
    pa.press('alt'), pa.press('r'), pa.press('z'), pa.press('d'), pa.press('d')
    pa.write("(matricula = '00{}')".format(str(pessoa.matricula))), t.sleep(1), pa.press('tab'), pa.write('2')
    pa.press('tab'), pa.press('enter'), t.sleep(10)
    pa.click(pa.center(pa.locateOnScreen('./static/salvar.png')))
    t.sleep(1), pp.copy(ps_acordo)
    pa.hotkey('ctrl', 'v'), t.sleep(1), pa.press('enter'), t.sleep(15)
    # # clique para fechar acordo
    pa.click(pa.center(pa.locateOnScreen('./static/fechar_janela.png'))), t.sleep(0.5)
    pa.click(pa.center(pa.locateOnScreen('./static/fechar_janela.png')))

    # # imprimir Anotações em CTPS
    pa.press('alt'), pa.press('r'), pa.press('a'), pa.press('c'), pa.press('e'), pa.press('tab')
    pa.write(str(pessoa.matricula))
    pa.press('tab', 2), pa.write(str(pessoa.admiss).replace('/', '')), pa.press('tab')
    pa.write(str(pessoa.admiss).replace('/', '')), pa.press('tab', 4), pa.press('space')
    pa.press('tab'), pa.press('enter'), t.sleep(1.5)
    pa.click(pa.center(pa.locateOnScreen('./static/salvar.png'))), t.sleep(1)
    pp.copy(ps_anotctps), pa.hotkey('ctrl', 'v'), t.sleep(0.5), pa.press('enter')
    t.sleep(2), pa.click(pa.center(pa.locateOnScreen('./static/fechar_janela.png'))), t.sleep(0.5)
    pa.click(pa.center(pa.locateOnScreen('./static/fechar_janela.png')))

    # # imprimir Termo VT
    pa.press('alt'), pa.press('r'), pa.press('a'), pa.press('v'), pa.press('e'), pa.press('tab')
    pa.write(str(pessoa.matricula)), pa.press('tab', 2), pa.write(str(pessoa.admiss).replace('/', ''))
    pa.press('tab'), pa.write('d'), pa.press('tab', 4), pa.press('space')
    pa.press('tab', 6), pa.press('enter'), t.sleep(1.5)
    pa.click(pa.center(pa.locateOnScreen('./static/salvar.png')))
    pp.copy(ps_termovt), pa.hotkey('ctrl', 'v'), t.sleep(0.5), pa.press('enter')
    t.sleep(2), pa.click(pa.center(pa.locateOnScreen('./static/fechar_janela.png'))), t.sleep(0.5)
    pa.click(pa.center(pa.locateOnScreen('./static/fechar_janela.png')))

    # # Imprimir Contrato
    pa.press('alt'), pa.press('r'), pa.press('z'), pa.press('d')
    if pessoa.tipo_contr == 'Horista':
        pa.press('c')
    else:
        pa.press('o')

    pa.write("(matricula = '00{}')".format(str(pessoa.matricula))), t.sleep(1), pa.press('tab'), pa.write('2')
    pa.press('tab'), pa.press('enter'), t.sleep(5)
    pa.click(pa.center(pa.locateOnScreen('./static/salvar.png')))
    pp.copy(ps_contrato), pa.hotkey('ctrl', 'v'), t.sleep(0.5), pa.press('enter')
    t.sleep(10), pa.click(pa.center(pa.locateOnScreen('./static/fechar_janela.png'))), t.sleep(0.5)
    pa.click(pa.center(pa.locateOnScreen('./static/fechar_janela.png')))

    # # Imprimir Folha de rosto de Cadastro
    pa.press('alt'), pa.press('r'), pa.press('i'), pa.press('o'), pa.press('r'), pa.press('e'), pa.press('tab')
    pa.write(str(pessoa.matricula)), pa.press('tab', 2), pa.write(str(pessoa.admiss).replace('/', '')), pa.press('tab',
                                                                                                                 2)
    pa.press('enter'), t.sleep(3)
    pa.click(pa.center(pa.locateOnScreen('./static/salvar.png')))
    pp.copy(ps_ficha), pa.hotkey('ctrl', 'v'), t.sleep(0.5), pa.press('enter')
    t.sleep(3), pa.click(pa.center(pa.locateOnScreen('./static/fechar_janela.png'))), t.sleep(0.5)
    pa.click(pa.center(pa.locateOnScreen('./static/fechar_janela.png')))

    # # Alterar AC e Salvar na pasta
    ac.paragraphs[1].text = str(ac.paragraphs[1].text).replace('#gerente', lotacao[str(pessoa.depto).title()][1])
    ac.paragraphs[2].text = str(ac.paragraphs[2].text).replace('#nome_completo', pessoa.nome)
    ac.paragraphs[3].text = str(ac.paragraphs[3].text).replace('#cargo', pessoa.cargo)
    ac.paragraphs[11].text = str(ac.paragraphs[11].text).replace('#salario', pessoa.salario)
    ac.save(p_contr + '\\AC.docx')
    docx2pdf.convert(p_contr + '\\AC.docx', p_contr + '\\AC.pdf')
    os.remove(p_contr + '\\AC.docx')

    # # Alterar Abertura de Conta e salvar na pasta
    abert_c.paragraphs[14].text = str(abert_c.paragraphs[14].text).replace('#nome_completo', pessoa.nome).replace(
        '#rg', pessoa.rg).replace(
        '#cpf', pessoa.cpf).replace('#endereco', pessoa.endereco).replace('#cep', pessoa.cep).replace('#bairro',
                                                                                                      pessoa.bairro).replace(
        '#cargo', pessoa.cargo).replace('#data', pessoa.admiss)
    abert_c.save(p_contr + '\\Abertura Conta.docx')
    docx2pdf.convert(p_contr + '\\Abertura Conta.docx', p_contr + '\\Abertura Conta.pdf')
    os.remove(p_contr + '\\Abertura Conta.docx')

    # # Alterar Ficha cadastral e salvar na pasta
    fch_c.paragraphs[34].text = str(fch_c.paragraphs[34].text).replace('#gerente#',
                                                                       lotacao[str(pessoa.depto).title()][1])
    fch_c.paragraphs[9].text = str(fch_c.paragraphs[9].text).replace('#nome_completo', pessoa.nome)
    fch_c.paragraphs[21].text = str(fch_c.paragraphs[21].text).replace('#cargo', pessoa.cargo).replace('#depart',
                                                                                                       str(pessoa.depto).title())
    fch_c.paragraphs[19].text = str(fch_c.paragraphs[19].text).replace('#end_eletr', pessoa.email)
    fch_c.paragraphs[17].text = str(fch_c.paragraphs[17].text).replace('#mae#', pessoa.mae)
    fch_c.paragraphs[16].text = str(fch_c.paragraphs[16].text).replace('#pai#', pessoa.pai)
    fch_c.paragraphs[15].text = str(fch_c.paragraphs[15].text).replace('#ident', pessoa.rg).replace('#cpf#',
                                                                                                    pessoa.cpf)
    fch_c.paragraphs[13].text = str(fch_c.paragraphs[13].text).replace('#telefone', pessoa.tel)
    fch_c.paragraphs[12].text = str(fch_c.paragraphs[12].text).replace('#codigo', pessoa.cep).replace('#cid', pessoa.cidade).replace(
        '#uf',
        pessoa.uf)
    fch_c.paragraphs[11].text = str(fch_c.paragraphs[11].text).replace('#local', pessoa.endereco).replace('#qd', pessoa.bairro)
    fch_c.paragraphs[10].text = str(fch_c.paragraphs[10].text).replace('#nasc', datetime.strftime(datetime.strptime(pessoa.nascimento,'%Y-%m-%d %H:%M:%S'), '%d/%m/%Y')).replace('#gen',
                                                                                                           pessoa.genero).replace(
        '#est_civ', str(pessoa.estado_civil).replace('1 - ','').replace('2 - ','').replace('3 - ','').replace('4 - ',''))
    fch_c.save(p_contr + '\\Ficha Cadastral.docx')
    docx2pdf.convert(p_contr + '\\Ficha Cadastral.docx', p_contr + '\\Ficha Cadastral.pdf')
    os.remove(p_contr + '\\Ficha Cadastral.docx')

    # # Alterar Recibos e salvar na pasta
    recibos.paragraphs[4].text = str(recibos.paragraphs[4].text).replace('#nome_completo', pessoa.nome)
    recibos.paragraphs[12].text = str(recibos.paragraphs[12].text).replace('#nome_completo', pessoa.nome)
    recibos.paragraphs[19].text = str(recibos.paragraphs[19].text).replace('#nome_completo', pessoa.nome)
    recibos.paragraphs[27].text = str(recibos.paragraphs[27].text).replace('#nome_completo', pessoa.nome)
    recibos.paragraphs[40].text = str(recibos.paragraphs[40].text).replace('#nome_completo', pessoa.nome)
    recibos.paragraphs[48].text = str(recibos.paragraphs[48].text).replace('#nome_completo', pessoa.nome)
    recibos.save(p_contr + '\\Recibos.docx')
    docx2pdf.convert(p_contr + '\\Recibos.docx', p_contr + '\\Recibos.pdf')
    os.remove(p_contr + '\\Recibos.docx')

    # # Alterar Código de Ética e Salvar na pasta
    codetic.paragraphs[534].text = str(codetic.paragraphs[534].text).replace('#nome_completo', pessoa.nome)
    codetic.paragraphs[535].text = str(codetic.paragraphs[535].text).replace('#func', pessoa.cargo)
    codetic.paragraphs[537].text = str(codetic.paragraphs[537].text).replace('#nome_completo', pessoa.nome)
    codetic.paragraphs[541].text = str(codetic.paragraphs[541].text).replace('#admiss', pessoa.admiss)
    codetic.save(p_contr + '\\Cod Etica.docx')
    docx2pdf.convert(p_contr + '\\Cod Etica.docx', p_contr + '\\Cod Etica.pdf')
    os.remove(p_contr + '\\Cod Etica.docx')
    tkinter.messagebox.showinfo(
        title='Documentos ok!',
        message='Documentos salvos com sucesso!'
    )


def enviaemailsfunc(matricula):
    pessoa = session.query(Colaborador).filter_by(matricula=matricula).first()
    p_contr = r'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\02 - Funcionários, Departamentos e Férias' \
              r'\000 - Pastas Funcionais\00 - ATIVOS\{}\Contratuais'.format(pessoa.nome)
    email_remetente = em_rem
    senha = k1
    # set up smtp connection
    s = smtplib.SMTP(host='smtp.office365.com', port=587)
    s.starttls()
    s.login(email_remetente, senha)

    # send e-mail to employee with a pdf file so he/she can go to bank to open an account
    msg = MIMEMultipart('alternative')
    msg['From'] = email_remetente
    msg['To'] = pessoa.email
    msg['Subject'] = "Carta para Abertura de conta"
    arquivo = p_contr + '\\Abertura Conta.pdf'
    text = MIMEText(f'''Olá, {str(pessoa.nome).title().split(" ")[0]}!<br><br>
    Segue sua carta para abertura de conta bancária no Itaú.<br>
    Você deve abrir a conta antes de iniciar seu contrato de trabalho. Ok?<br><br>
    Atenciosamente,<br>
    <img src="cid:image1">''', 'html')
    
    # set up the parameters of the message
    msg.attach(text)
    image = MIMEImage(
        open(r'C:\Users\Felipe Rodrigues\PycharmProjects\AutomacaoCia\Admissao\static\assinatura.png', 'rb').read())
    image.add_header('Content-ID', '<image1>')
    msg.attach(image)
    # attach pdf file
    part = MIMEBase('application', "octet-stream")
    part.set_payload(open(arquivo, "rb").read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment',
                    filename=f'Carta Banco {str(pessoa.nome).title().split(" ")[0]}.pdf')
    msg.attach(part)
    s.sendmail(email_remetente, pessoa.email, msg.as_string())
    del msg

    # send e-mail to coworker asking to register the ner employee
    msg = MIMEMultipart('alternative')
    arquivo = p_contr + '\\Ficha Cadastral.pdf'
    text = MIMEText(f'''Oi, Wallace!<br><br>Segue a ficha cadastral do(a) {pessoa.nome}.<br><br>Abs.,<br><img src="cid:image1">''', 'html')
    msg.attach(text)
    image = MIMEImage(
        open(r'C:\Users\Felipe Rodrigues\PycharmProjects\AutomacaoCia\Admissao\static\assinatura.png', 'rb').read())
    image.add_header('Content-ID', '<image1>')
    msg.attach(image)
    # set up the parameters of the message
    msg['From'] = email_remetente
    msg['To'] = em_ti
    msg['Subject'] = f"Ficha Cadastral {str(pessoa.nome).title().split(' ')[0]}"
    # attach pdf file
    part = MIMEBase('application', "octet-stream")
    part.set_payload(open(arquivo, "rb").read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment',
                    filename=f'Ficha Cadastral {str(pessoa.nome).title().split(" ")[0]}.pdf')
    msg.attach(part)
    s.sendmail(email_remetente, em_ti, msg.as_string())
    del msg
    s.quit()
    tkinter.messagebox.showinfo(
        title='E-mails ok!',
        message='E-mails enviados com sucesso'
    )


caminhoest = StringVar()


def cadastro_estagiario(solicitar_contr=0, caminho='', editar=0, ondestou=0, nome='', matricula='', admissao='',
                         cargo='', depto='', tipo_contr='Horista',
                         hrsem='25', hrmens='100'):
    pa.FAILSAFE = False
    salario = 5.10
    if solicitar_contr == 1:
        hoje = datetime.today()
        wb = l_w(caminho)
        sh = wb['Respostas ao formulário 1']
        num, name = nome.strip().split(' - ')
        linha = int(num)
        lotacao = {
            'Unidade Park Sul - qualquer departamento': ['0013', 'Thais Feitosa', 'thais.morais@ciaathletica.com.br',
                                                         'Líder Park Sul'],
            'Kids': ['0010', 'Cindy Stefanie', 'cindy.neves@ciaathletica.com.br', 'Líder Kids'],
            'Musculação': ['0007', 'Thais Feitosa', 'thais.morais@ciaathletica.com.br', 'Líder Musculação'],
            'Esportes e Lutas': ['0008', 'Morgana Rossini', 'morganalourenco@yahoo.com.br', 'Líder Natação'],
            'Crossfit': ['0012', 'Guilherme Salles', 'gmoreirasalles@gmail.com', 'Líder Crossfit'],
            'Ginástica': ['0006', 'Morgana Rossini', 'morganalourenco@yahoo.com.br', 'Líder Ginástica'],
            'Gestantes': ['0006', 'Filipe Feijó', 'filipe.feijo@ciaathletica.com.br', 'Gerente Técnico'],
            'Recepção': ['0003', 'Paulo Renato', 'paulo.simoes@ciaathletica.com.br', 'Gerente Vendas'],
            'Administrativo': ['0001', 'Felipe Rodrigues', 'felipe.rodrigues@ciaathletica.com.br', 'Gerente RH'],
            'Manutenção': ['0004', 'José Aparecido', 'aparecido.grota@ciaathletica.com.br', 'Gerente Manutenção'],
        }
        cadastro = {'nome': str(sh[f"C{linha}"].value).title().strip(), 'nasc_ed': sh[f"D{linha}"].value,
                    'genero': str(sh[f"E{linha}"].value), 'est_civ': str(sh[f"F{linha}"].value),
                    'pai': str(sh[f"M{linha}"].value), 'mae': str(sh[f"N{linha}"].value), 'end': str(sh[f"O{linha}"].value),
                    'num': str(sh[f"P{linha}"].value), 'bairro': str(sh[f"Q{linha}"].value),
                    'cep': str(sh[f"R{linha}"].value).replace('.', '').replace('-', ''),
                    'cid_end': str(sh[f"S{linha}"].value), 'uf_end': str(sh[f"T{linha}"].value),
                    'tel': str(sh[f"U{linha}"].value).replace('(', '').replace(')', '').replace('-', '').replace(' ', ''),
                    'mun_end': str(sh[f"AP{linha}"].value),
                    'cpf': str(sh[f"V{linha}"].value).strip().replace('.', '').replace('-', '').replace(' ', '').zfill(11),
                    'rg': str(sh[f"W{linha}"].value).strip().replace('.', '').replace('-', '').replace(' ', ''),
                    'emissor': str(sh[f"X{linha}"].value),
                    'lotacao': str(lotacao[f'{sh[f"AG{linha}"].value}'][0]).zfill(4),
                    'cargo': str(sh[f"AH{linha}"].value), 'horario': str(sh[f"AI{linha}"].value),
                    'email': str(sh[f"B{linha}"].value).strip(),
                    'admissao_ed': str(sh[f"AL{linha}"].value),
                    'faculdade': str(sh[f"AV{linha}"].value), 'semestre': str(sh[f"AS{linha}"].value),
                    'turno': str(sh[f"AT{linha}"].value), 'conclusao': str(sh[f"AU{linha}"].value), 'salario': str(sh[f"AM{linha}"].value),
                    'hrsemanais': str(sh[f"AQ{linha}"].value), 'hrmensais': str(sh[f"AR{linha}"].value)}
        email_remetente = em_rem
        senha = k1
        lot = lotacao[f'{sh[f"AG{linha}"].value}']
        pasta = r'\192.168.0.250'
        modelo = f'\\{pasta}\\rh\\01 - RH\\01 - Administração.Controles\\02 - Funcionários, Departamentos e Férias\\000 - Pastas Funcionais\\00 - ATIVOS\\0 - Estagiários\\Modelo'
        os.makedirs(
            f'\\{pasta}\\rh\\01 - RH\\01 - Administração.Controles\\02 - Funcionários, Departamentos e Férias\\000 - Pastas Funcionais\\00 - ATIVOS\\0 - Estagiários\\0 - Ainda nao iniciaram\\{str(cadastro["nome"]).upper()}')
        os.makedirs(
            f'\\{pasta}\\rh\\01 - RH\\01 - Administração.Controles\\02 - Funcionários, Departamentos e Férias\\000 - Pastas Funcionais\\00 - ATIVOS\\0 - Estagiários\\0 - Ainda nao iniciaram\\{str(cadastro["nome"]).upper()}\\Atestados')
        os.makedirs(
            f'\\{pasta}\\rh\\01 - RH\\01 - Administração.Controles\\02 - Funcionários, Departamentos e Férias\\000 - Pastas Funcionais\\00 - ATIVOS\\0 - Estagiários\\0 - Ainda nao iniciaram\\{str(cadastro["nome"]).upper()}\\Diversos')
        os.makedirs(
            f'\\{pasta}\\rh\\01 - RH\\01 - Administração.Controles\\02 - Funcionários, Departamentos e Férias\\000 - Pastas Funcionais\\00 - ATIVOS\\0 - Estagiários\\0 - Ainda nao iniciaram\\{str(cadastro["nome"]).upper()}\\Contratuais')
        os.makedirs(
            f'\\{pasta}\\rh\\01 - RH\\01 - Administração.Controles\\02 - Funcionários, Departamentos e Férias\\000 - Pastas Funcionais\\00 - ATIVOS\\0 - Estagiários\\0 - Ainda nao iniciaram\\{str(cadastro["nome"]).upper()}\\Ferias')
        os.makedirs(
            f'\\{pasta}\\rh\\01 - RH\\01 - Administração.Controles\\02 - Funcionários, Departamentos e Férias\\000 - Pastas Funcionais\\00 - ATIVOS\\0 - Estagiários\\0 - Ainda nao iniciaram\\{str(cadastro["nome"]).upper()}\\Ponto')
        os.makedirs(
            f'\\{pasta}\\rh\\01 - RH\\01 - Administração.Controles\\02 - Funcionários, Departamentos e Férias\\000 - Pastas Funcionais\\00 - ATIVOS\\0 - Estagiários\\0 - Ainda nao iniciaram\\{str(cadastro["nome"]).upper()}\\Recibo')
        os.makedirs(
            f'\\{pasta}\\rh\\01 - RH\\01 - Administração.Controles\\02 - Funcionários, Departamentos e Férias\\000 - Pastas Funcionais\\00 - ATIVOS\\0 - Estagiários\\0 - Ainda nao iniciaram\\{str(cadastro["nome"]).upper()}\\Rescisao')
        pasta_contratuais = f'\\{pasta}\\rh\\01 - RH\\01 - Administração.Controles\\02 - Funcionários, Departamentos e Férias\\000 - Pastas Funcionais\\00 - ATIVOS\\0 - Estagiários\\0 - Ainda nao iniciaram\\{str(cadastro["nome"]).upper()}\\Contratuais'

        # change tree docx model files with intern data and save pdfs files
        solicitacao = docx.Document(modelo + r'\Solicitacao MODELO - Copia.docx')
        solicitacao.tables[0].rows[4].cells[0].paragraphs[0].text = str(
            solicitacao.tables[0].rows[4].cells[0].paragraphs[0].text).replace('#supervisor_estagio', f'{lot[1]}')
        solicitacao.tables[0].rows[5].cells[1].paragraphs[0].text = str(
            solicitacao.tables[0].rows[5].cells[1].paragraphs[0].text).replace('#cargo', f'{lot[3]}')
        solicitacao.tables[0].rows[6].cells[0].paragraphs[0].text = str(
            solicitacao.tables[0].rows[6].cells[0].paragraphs[0].text).replace('#email_supervisor', f'{lot[2]}')
        solicitacao.tables[0].rows[9].cells[0].paragraphs[0].text = str(
            solicitacao.tables[0].rows[9].cells[0].paragraphs[0].text).replace('#horario', cadastro['horario'])
        solicitacao.tables[0].rows[14].cells[0].paragraphs[0].text = str(
            solicitacao.tables[0].rows[14].cells[0].paragraphs[0].text).replace('#nome_completo', cadastro['nome'])
        solicitacao.tables[0].rows[15].cells[0].paragraphs[0].text = str(
            solicitacao.tables[0].rows[15].cells[0].paragraphs[0].text).replace('#nasc',
                                                                                datetime.strftime(cadastro['nasc_ed'],
                                                                                                  '%d/%m/%Y')
                                                                                ).replace('#rg', cadastro['rg']).replace('#cpf', cadastro['cpf'])
        solicitacao.tables[0].rows[16].cells[0].paragraphs[0].text = str(
            solicitacao.tables[0].rows[16].cells[0].paragraphs[0].text).replace('#sexo', cadastro['genero'])
        solicitacao.tables[0].rows[17].cells[0].paragraphs[0].text = str(
            solicitacao.tables[0].rows[17].cells[0].paragraphs[0].text).replace('#endereco', cadastro['end'])
        solicitacao.tables[0].rows[18].cells[0].paragraphs[0].text = str(
            solicitacao.tables[0].rows[18].cells[0].paragraphs[0].text).replace('#cep', cadastro['cep']).replace(
            '#bairro', cadastro['bairro'])
        solicitacao.tables[0].rows[19].cells[0].paragraphs[0].text = str(
            solicitacao.tables[0].rows[19].cells[0].paragraphs[0].text).replace('#telefone', cadastro['tel'])
        solicitacao.tables[0].rows[20].cells[0].paragraphs[0].text = str(
            solicitacao.tables[0].rows[20].cells[0].paragraphs[0].text).replace('#end_eletr', cadastro['email'])
        solicitacao.tables[0].rows[22].cells[0].paragraphs[0].text = str(
            solicitacao.tables[0].rows[22].cells[0].paragraphs[0].text).replace('#semestre', cadastro['semestre'])
        solicitacao.tables[0].rows[23].cells[0].paragraphs[0].text = str(
            solicitacao.tables[0].rows[23].cells[0].paragraphs[0].text).replace('#turno', cadastro['turno']).replace(
            '#ano_concl', cadastro['conclusao'])
        solicitacao.tables[0].rows[24].cells[0].paragraphs[0].text = str(
            solicitacao.tables[0].rows[24].cells[0].paragraphs[0].text).replace('#faculdade', cadastro['faculdade'])
        solicitacao.save(pasta_contratuais + f'\\Pedido TCE {str(cadastro["nome"]).split(" ")[0]}.docx')
        docx2pdf.convert(pasta_contratuais + f'\\Pedido TCE {str(cadastro["nome"]).split(" ")[0]}.docx',
                         pasta_contratuais + f'\\Pedido TCE {str(cadastro["nome"]).split(" ")[0]}.pdf')
        os.remove(pasta_contratuais + f'\\Pedido TCE {str(cadastro["nome"]).split(" ")[0]}.docx')

        ficha_cadastral = docx.Document(modelo + r'\Ficha Cadastral MODELO - Copia.docx')
        ficha_cadastral.tables[1].rows[0].cells[0].paragraphs[0].text = str(
            ficha_cadastral.tables[1].rows[0].cells[0].paragraphs[0].text).replace('#nome_completo', cadastro['nome'])
        ficha_cadastral.tables[1].rows[1].cells[0].paragraphs[0].text = str(
            ficha_cadastral.tables[1].rows[1].cells[0].paragraphs[0].text).replace('#nasc', datetime.strftime(
            cadastro['nasc_ed'], '%d/%m/%Y'))
        ficha_cadastral.tables[1].rows[1].cells[2].paragraphs[0].text = str(
            ficha_cadastral.tables[1].rows[1].cells[2].paragraphs[0].text).replace('#gen', cadastro['genero'])
        ficha_cadastral.tables[1].rows[1].cells[4].paragraphs[0].text = str(
            ficha_cadastral.tables[1].rows[1].cells[4].paragraphs[0].text).replace('#est_civ', cadastro['est_civ'])
        ficha_cadastral.tables[1].rows[2].cells[0].paragraphs[0].text = str(
            ficha_cadastral.tables[1].rows[2].cells[0].paragraphs[0].text).replace('#local', cadastro['end'])
        ficha_cadastral.tables[1].rows[2].cells[4].paragraphs[0].text = str(
            ficha_cadastral.tables[1].rows[2].cells[4].paragraphs[0].text).replace('#qd', cadastro['bairro'])
        ficha_cadastral.tables[1].rows[2].cells[7].paragraphs[0].text = str(
            ficha_cadastral.tables[1].rows[2].cells[7].paragraphs[0].text).replace('#codigo', cadastro['cep'])
        ficha_cadastral.tables[1].rows[4].cells[1].paragraphs[0].text = str(
            ficha_cadastral.tables[1].rows[4].cells[1].paragraphs[0].text).replace('#telefone', cadastro['tel'])
        ficha_cadastral.tables[1].rows[4].cells[5].paragraphs[0].text = str(
            ficha_cadastral.tables[1].rows[4].cells[5].paragraphs[0].text).replace('#ident', cadastro['rg'])
        ficha_cadastral.tables[1].rows[5].cells[1].paragraphs[0].text = str(
            ficha_cadastral.tables[1].rows[5].cells[1].paragraphs[0].text).replace('#cpf#', cadastro['cpf'])
        ficha_cadastral.tables[1].rows[6].cells[3].paragraphs[0].text = str(
            ficha_cadastral.tables[1].rows[6].cells[3].paragraphs[0].text).replace('#pai#', cadastro['pai'])
        ficha_cadastral.tables[1].rows[7].cells[1].paragraphs[0].text = str(
            ficha_cadastral.tables[1].rows[7].cells[1].paragraphs[0].text).replace('#mae#', cadastro['mae'])
        ficha_cadastral.tables[1].rows[8].cells[0].paragraphs[0].text = str(
            ficha_cadastral.tables[1].rows[8].cells[0].paragraphs[0].text).replace('#end_eletr', cadastro['email'])
        ficha_cadastral.tables[1].rows[8].cells[1].paragraphs[0].text = str(
            ficha_cadastral.tables[1].rows[8].cells[1].paragraphs[0].text).replace('#depart', str(sh["AG3"].value))
        ficha_cadastral.save(pasta_contratuais + f'\\Ficha Cadastral {str(cadastro["nome"]).split(" ")[0]}.docx')
        docx2pdf.convert(pasta_contratuais + f'\\Ficha Cadastral {str(cadastro["nome"]).split(" ")[0]}.docx',
                         pasta_contratuais + f'\\Ficha Cadastral {str(cadastro["nome"]).split(" ")[0]}.pdf')
        os.remove(pasta_contratuais + f'\\Ficha Cadastral {str(cadastro["nome"]).split(" ")[0]}.docx')

        carta_banco = docx.Document(
            modelo + r'\Abertura Conta MODELO.docx')
        carta_banco.paragraphs[14].text = str(carta_banco.paragraphs[14].text).replace('#nome_completo',
                                                                                       cadastro['nome']
                                                                                       ).replace('#rg',cadastro['rg']
                                                                                                 ).replace(
            '#cpf', cadastro['cpf']).replace('#endereço', cadastro['end']).replace('#cep', cadastro['cep']).replace(
            '#bairro', cadastro['bairro']).replace('#desde#', datetime.strftime(hoje, '%d/%m/%Y'))
        carta_banco.save(pasta_contratuais + f'\\Carta Banco {str(cadastro["nome"]).split(" ")[0]}.docx')
        docx2pdf.convert(pasta_contratuais + f'\\Carta Banco {str(cadastro["nome"]).split(" ")[0]}.docx',
                         pasta_contratuais + f'\\Carta Banco {str(cadastro["nome"]).split(" ")[0]}.pdf')
        os.remove(pasta_contratuais + f'\\Carta Banco {str(cadastro["nome"]).split(" ")[0]}.docx')

        # set up smtpp connection
        s = smtplib.SMTP(host='smtp.office365.com', port=587)
        s.starttls()
        s.login(email_remetente, senha)

        # send e-mail to intern with a pdf file so he/she can go to bank to open an account
        msg = MIMEMultipart('alternative')
        arquivo = pasta_contratuais + f'\\Carta Banco {str(cadastro["nome"]).split(" ")[0]}.pdf'
        text = MIMEText(f'''Olá, {str(cadastro["nome"]).split(" ")[0]}!<br><br>
        Segue sua carta para abertura de conta bancária no Itaú.<br>
        Você deve abrir a conta antes de iniciar os trabalhos no estágio. Ok?<br>
        Você já pode buscar seu contrato no IF. Será necessário levar uma declaração de matrícula do seu curso.<br><br>
        Atenciosamente,<br>
        <img src="cid:image1">''', 'html')
        msg.attach(text)
        image = MIMEImage(
            open(r'C:\Users\Felipe Rodrigues\PycharmProjects\AutomacaoCia\Admissao\static\assinatura.png', 'rb').read())
        image.add_header('Content-ID', '<image1>')
        msg.attach(image)
        # set up the parameters of the message
        msg['From'] = email_remetente
        msg['To'] = cadastro['email']
        msg['Subject'] = "Carta para Abertura de conta"
        # attach pdf file
        part = MIMEBase('application', "octet-stream")
        part.set_payload(open(arquivo, "rb").read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment',
                        filename=f'Carta Banco {str(cadastro["nome"]).split(" ")[0]}.pdf')
        msg.attach(part)
        s.sendmail(email_remetente, cadastro['email'], msg.as_string())
        del msg

        # send e-mail to coworker asking to register the intern
        msg = MIMEMultipart('alternative')
        arquivo = pasta_contratuais + f'\\Ficha Cadastral {str(cadastro["nome"]).split(" ")[0]}.pdf'
        text = MIMEText(f'''Oi, Wallace!<br><br>
        Segue a ficha cadastral do(a) estagiário(a) {cadastro["nome"]}.<br><br>
        Abs.,<br>
        <img src="cid:image1">''', 'html')
        msg.attach(text)
        image = MIMEImage(
            open(r'C:\Users\Felipe Rodrigues\PycharmProjects\AutomacaoCia\Admissao\static\assinatura.png', 'rb').read())
        # Define the image's ID as referenced in the HTML body above
        image.add_header('Content-ID', '<image1>')
        msg.attach(image)
        # set up the parameters of the message
        msg['From'] = email_remetente
        msg['To'] = em_ti
        msg['Subject'] = f"Ficha Cadastral {str(cadastro['nome']).split(' ')[0]}"
        # attach pdf file
        part = MIMEBase('application', "octet-stream")
        part.set_payload(open(arquivo, "rb").read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment',
                        filename=f'Ficha Cadastral {str(cadastro["nome"]).split(" ")[0]}.pdf')
        msg.attach(part)
        s.sendmail(email_remetente, em_ti, msg.as_string())
        del msg

        # send document asking for the intern contract
        msg = MIMEMultipart('alternative')
        arquivo = pasta_contratuais + f'\\Pedido TCE {str(cadastro["nome"]).split(" ")[0]}.pdf'
        text = MIMEText(f'''Olá!\n\nSegue pedido de TCE do(a) estagiário(a) {cadastro["nome"]}.\n\nAtenciosamente,<br><img src="cid:image1">''', 'html')
        msg.attach(text)
        image = MIMEImage(
            open(r'C:\Users\Felipe Rodrigues\PycharmProjects\AutomacaoCia\Admissao\static\assinatura.png', 'rb').read())
        # Define the image's ID as referenced in the HTML body above
        image.add_header('Content-ID', '<image1>')
        msg.attach(image)
        # set up the parameters of the message
        msg['From'] = email_remetente
        msg['To'] = em_if
        msg['Subject'] = f"Pedido TCE {str(cadastro['nome']).split(' ')[0]}"
        # attach pdf file
        part = MIMEBase('application', "octet-stream")
        part.set_payload(open(arquivo, "rb").read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment',
                        filename=f'Pedido TCE {str(cadastro["nome"]).split(" ")[0]}.pdf')
        msg.attach(part)
        s.sendmail(email_remetente, em_if, msg.as_string())
        del msg
        s.quit()
        tkinter.messagebox.showinfo(
            title='E-mails ok!',
            message='E-mails enviados com sucesso'
        )

    else:
        if editar == 0:
            if ondestou == 0:
                # Cadastro iniciado na Cia
                wb = l_w(caminho)
                sh = wb['Respostas ao formulário 1']
                num, name = nome.strip().split(' - ')
                linha = int(num)
                lotacao = {'UNIDADE PARK SUL - QUALQUER DEPARTAMENTO': '0013', 'KIDS': '0010', 'MUSCULAÇÃO': '0007',
                           'ESPORTES E LUTAS': '0008', 'CROSSFIT': '0012', 'GINÁSTICA': '0006', 'GESTANTES': '0008',
                           'RECEPÇÃO': '0003',
                           'FINANCEIRO': '0001', 'TI': '0001', 'MARKETING': '0001', 'MANUTENÇÃO': '0004'}
                if str(sh[f'E{linha}'].value) == 'Masculino':
                    cargo = 'ESTAGIARIO'
                else:
                    cargo = 'ESTAGIARIA'

                estag = session.query(Colaborador).filter_by(matricula=matricula).first()
                if estag:
                    pass
                else:
                    estag_cadastrado = Colaborador(
                        matricula=matricula, nome=name.upper(), admiss=admissao,
                        nascimento=str(sh[f'D{linha}'].value),
                        cpf=str(sh[f'V{linha}'].value).replace('.','').replace('-','').zfill(11),
                        rg=str(int(sh[f'W{linha}'].value)),
                        emissor='SSP/DF', email=str(sh[f'B{linha}'].value),
                        genero=str(sh[f'E{linha}'].value),
                        estado_civil=str(sh[f'F{linha}'].value), cor='9',
                        instru='08 - Educação Superior Incompleta',
                        nacional='Brasileiro(a)',
                        pai=str(sh[f'M{linha}'].value).upper(),
                        mae=str(sh[f'N{linha}'].value).upper(), endereco=str(sh[f'O{linha}'].value),
                        num='1',
                        bairro=str(sh[f'Q{linha}'].value), cep=str(sh[f'R{linha}'].value).replace('.','').replace('-',''),
                        cidade='Brasília', cid_nas='Brasília - DF',
                        uf='DF',
                        cod_municipioend=municipios['DF']['Brasília'],
                        tel=str(sh[f'U{linha}'].value).replace('(','').replace(')','').replace('.','').replace('-',''),
                        depto=depto, cargo=cargo,
                        horario=str(sh[f'AI{linha}'].value), salario=salario, tipo_contr=tipo_contr,
                        hr_sem='25', hr_mens='100',
                        est_semestre=str(sh[f'AS{linha}'].value),
                        est_turno=str(sh[f'AT{linha}'].value),
                        est_prev_conclu=str(sh[f'AU{linha}'].value),
                        est_faculdade=str(sh[f'AV{linha}'].value),
                        est_endfacul='End',
                        est_numendfacul='1',
                        est_bairroendfacul='Bairro'
                    )
                    session.add(estag_cadastrado)
                    session.commit()
                pasta = r'\192.168.0.250'
                try:
                    os.makedirs(
                        f'\\{pasta}\\rh\\01 - RH\\01 - Administração.Controles\\02 - Funcionários, Departamentos e Férias\\000 - Pastas Funcionais\\00 - ATIVOS\\0 - Estagiários\\0 - Ainda nao iniciaram\\{estag.nome}')
                    os.makedirs(
                        f'\\{pasta}\\rh\\01 - RH\\01 - Administração.Controles\\02 - Funcionários, Departamentos e Férias\\000 - Pastas Funcionais\\00 - ATIVOS\\0 - Estagiários\\0 - Ainda nao iniciaram\\{estag.nome}\\Atestados')
                    os.makedirs(
                        f'\\{pasta}\\rh\\01 - RH\\01 - Administração.Controles\\02 - Funcionários, Departamentos e Férias\\000 - Pastas Funcionais\\00 - ATIVOS\\0 - Estagiários\\0 - Ainda nao iniciaram\\{estag.nome}\\Diversos')
                    os.makedirs(
                        f'\\{pasta}\\rh\\01 - RH\\01 - Administração.Controles\\02 - Funcionários, Departamentos e Férias\\000 - Pastas Funcionais\\00 - ATIVOS\\0 - Estagiários\\0 - Ainda nao iniciaram\\{estag.nome}\\Contratuais')
                    os.makedirs(
                        f'\\{pasta}\\rh\\01 - RH\\01 - Administração.Controles\\02 - Funcionários, Departamentos e Férias\\000 - Pastas Funcionais\\00 - ATIVOS\\0 - Estagiários\\0 - Ainda nao iniciaram\\{estag.nome}\\Ferias')
                    os.makedirs(
                        f'\\{pasta}\\rh\\01 - RH\\01 - Administração.Controles\\02 - Funcionários, Departamentos e Férias\\000 - Pastas Funcionais\\00 - ATIVOS\\0 - Estagiários\\0 - Ainda nao iniciaram\\{estag.nome}\\Ponto')
                    os.makedirs(
                        f'\\{pasta}\\rh\\01 - RH\\01 - Administração.Controles\\02 - Funcionários, Departamentos e Férias\\000 - Pastas Funcionais\\00 - ATIVOS\\0 - Estagiários\\0 - Ainda nao iniciaram\\{estag.nome}\\Recibo')
                    os.makedirs(
                        f'\\{pasta}\\rh\\01 - RH\\01 - Administração.Controles\\02 - Funcionários, Departamentos e Férias\\000 - Pastas Funcionais\\00 - ATIVOS\\0 - Estagiários\\0 - Ainda nao iniciaram\\{estag.nome}\\Rescisao')
                except:
                    pass
                # abrir cadastro no dexion e atualizar informações campo a campo
                pessoa = session.query(Colaborador).filter_by(matricula=matricula).first()
                pastapessoa = f'\\{pasta}\\rh\\01 - RH\\01 - Administração.Controles\\02 - Funcionários, Departamentos e Férias\\000 - Pastas Funcionais\\00 - ATIVOS\\0 - Estagiários\\0 - Ainda nao iniciaram\\{pessoa.nome}'
                pa.click(pa.center(pa.locateOnScreen('./static/Dexion.png')))
                pa.press('alt'), pa.press('a'), pa.press('t'), t.sleep(2), pa.press(
                    'i'), t.sleep(5), pa.write(str(pessoa.matricula)), pa.press('enter'), t.sleep(20)
                pp.copy(pessoa.nome), pa.hotkey('ctrl', 'v'), pa.press('tab'), pa.write(pessoa.cpf)
                t.sleep(1), pa.press('tab', 3), pa.write(pessoa.genero), pa.press('tab'), pa.write(pessoa.cor)
                t.sleep(1), pa.press('tab', 2), pa.write(pessoa.instru)
                t.sleep(1), pa.press('tab')
                if str(pessoa.estado_civil) == 'Casado(a)':
                    pa.write('2')
                    pa.press('tab', 6)
                else:
                    pa.write('1')
                    pa.press('tab', 5)
                pa.write(datetime.strftime(datetime.strptime(pessoa.nascimento, '%Y-%m-%d %H:%M:%S'), '%d%m%Y'))
                t.sleep(1), pa.press('tab'), pp.copy(pessoa.cid_nas), pa.hotkey('ctrl', 'v'), pa.press('tab')
                t.sleep(1), pa.write(pessoa.uf), pa.press('tab'), pa.write(pessoa.cod_municipioend), pa.press('tab')
                t.sleep(1), pp.copy(pessoa.pai), pa.hotkey('ctrl', 'v'), pa.press('tab'), pa.write('105'), pa.press('tab')
                t.sleep(1), pp.copy(pessoa.mae), pa.hotkey('ctrl', 'v'), pa.press('tab'), pa.write('105')
                # # clique em documentos
                pa.click(pa.center(pa.locateOnScreen('./static/Documentos.png')))
                pa.press('tab'), pa.write(str(pessoa.rg)), pa.press('tab'), pa.write(
                    pessoa.emissor), pa.press('tab', 3), pa.write(pessoa.cod_municipioend)
                # #clique em endereço
                pa.click(pa.center(pa.locateOnScreen('./static/Endereco.png')))
                pa.press('tab', 2), pp.copy(pessoa.endereco), pa.hotkey('ctrl', 'v'), pa.press(
                    'tab'), pa.write(pessoa.num), pa.press('tab', 2)
                pp.copy(pessoa.bairro), pa.hotkey('ctrl', 'v'), pa.press('tab'), pp.copy(pessoa.cidade), pa.hotkey(
                    'ctrl', 'v'), pa.press('tab'), pa.write(pessoa.uf)
                pa.press('tab'), pa.write(pessoa.cep), pa.press('tab'), pa.write(pessoa.cod_municipioend), pa.press(
                    'tab'), pa.write(str(pessoa.tel)), pa.press('tab', 2), pa.write(pessoa.email)
                # #clique em dados contratuais
                pa.click(pa.center(pa.locateOnScreen('./static/Contratuais.png')))
                pa.press('tab'), pa.write(str(pessoa.admiss).replace('/', '')), t.sleep(1)
                pa.press('tab'), pa.write('9')
                pa.press('tab', 7), pa.write('n'), pa.press('tab'), pa.write('4')
                pa.press('tab'), pa.write('Ed. Fisica')
                pa.press('tab', 2), pa.write(str(int(str(pessoa.admiss).replace('/', ''))+2).zfill(8))
                # #clique em instituição de ensino
                try:
                    pa.click(pa.center(pa.locateOnScreen('./static/faculdade.png')))
                except:
                    t.sleep(3)
                    pa.click(pa.center(pa.locateOnScreen('./static/faculdade.png')))
                pa.press('tab'), pp.copy(pessoa.est_faculdade), pa.hotkey('ctrl', 'v')
                pa.press('tab'), pp.copy(pessoa.est_endfacul), pa.hotkey('ctrl', 'v')
                pa.press('tab'), pp.copy(pessoa.est_numendfacul), pa.hotkey('ctrl', 'v')
                pa.press('tab'), pp.copy(pessoa.est_bairroendfacul), pa.hotkey('ctrl', 'v')
                # #clique em Outros
                try:
                    pa.click(pa.center(pa.locateOnScreen('./static/Outros.png')))
                except:
                    t.sleep(5)
                    pa.click(pa.center(pa.locateOnScreen('./static/Outros.png')))
                t.sleep(2), pa.write('CARGO GERAL')
                pa.press('tab'), pa.write(pessoa.cargo)
                t.sleep(1), pa.press('tab'), pa.write(pessoa.salario), pa.press('tab')
                if str(pessoa.tipo_contr) == 'Horista':
                    pa.press('1')
                else:
                    pa.press('5')
                pa.press('tab', 2), t.sleep(1), pa.write(str(pessoa.hr_sem)), pa.press('tab'), t.sleep(1)
                pa.write(str(pessoa.hr_mens))
                # #clique em eventos trabalhistas
                pa.click(pa.center(pa.locateOnScreen('./static/EVTrab.png')))
                t.sleep(1)
                # #clique em lotação
                pa.click(pa.center(pa.locateOnScreen('./static/Lotacoes.png')))
                pa.press('tab'), pa.press('tab'), pa.write('i'), pa.write(str(pessoa.admiss).replace('/', ''))
                t.sleep(1), pa.press('enter'), t.sleep(1)
                pp.copy(lotacao[f'{pessoa.depto}']), pa.hotkey('ctrl', 'v'), pa.press('enter'), pa.write('f')
                pa.press('tab'), pa.write('4')
                # #clique em salvar lotação
                pa.click(pa.center(pa.locateOnScreen('./static/Salvarbtn.png'))), t.sleep(1)
                # #clique em fechar lotação
                try:
                    pa.click(pa.center(pa.locateOnScreen('./static/Fecharlot.png'))), t.sleep(1)
                except:
                    t.sleep(4)
                    pa.click(pa.center(pa.locateOnScreen('./static/Fecharlot.png')))
                # #clique em Compatibilidade
                pa.click(pa.center(pa.locateOnScreen('./static/Compatibilidade3.png'))), t.sleep(1)
                # #clique em Compatibilidade de novo
                pa.click(pa.center(pa.locateOnScreen('./static/Compatibilidade2.png'))), t.sleep(1)
                pa.press('tab', 2), pa.write('9')
                # #clique em Salvar
                pa.click(pa.center(pa.locateOnScreen('./static/Salvarcadastro.png'))), t.sleep(10)
                # #clique em fechar novo cadastro
                pa.click(pa.center(pa.locateOnScreen('./static/Fecharnovo1.png'))), t.sleep(2)
                # #clique em fechar trabalhadores
                pa.click(pa.center(pa.locateOnScreen('./static/Fechartrab1.png'))), t.sleep(0.5)
                os.rename(pastapessoa,
                          f'\\{pasta}\\rh\\01 - RH\\01 - Administração.Controles\\02 - Funcionários, Departamentos e Férias\\000 - Pastas Funcionais\\00 - ATIVOS\\0 - Estagiários\\{pessoa.nome}')
                tkinter.messagebox.showinfo(
                    title='Cadastro ok!',
                    message='Cadastro realizado com sucesso!'
                )
            else:
                # Cadastro iniciado em casa
                wb = l_w(caminho)
                sh = wb['Respostas ao formulário 1']
                num, name = nome.strip().split(' - ')
                linha = int(num)
                lotacao = {'UNIDADE PARK SUL - QUALQUER DEPARTAMENTO': '0013', 'KIDS': '0010', 'MUSCULAÇÃO': '0007',
                           'ESPORTES E LUTAS': '0008', 'CROSSFIT': '0012', 'GINÁSTICA': '0006', 'GESTANTES': '0008',
                           'RECEPÇÃO': '0003',
                           'FINANCEIRO': '0001', 'TI': '0001', 'MARKETING': '0001', 'MANUTENÇÃO': '0004'}
                if str(sh[f'E{linha}'].value) == 'Masculino':
                    cargo = 'ESTAGIARIO'
                else:
                    cargo = 'ESTAGIARIA'

                estag = session.query(Colaborador).filter_by(matricula=matricula).first()
                if estag:
                    pass
                else:
                    estag_cadastrado = Colaborador(
                        matricula=matricula, nome=name.upper(), admiss=admissao,
                        nascimento=str(sh[f'D{linha}'].value),
                        cpf=str(int(sh[f'V{linha}'].value)).zfill(11),
                        rg=str(int(sh[f'W{linha}'].value)),
                        emissor='SSP/DF', email=str(sh[f'B{linha}'].value),
                        genero=str(sh[f'E{linha}'].value),
                        estado_civil=str(sh[f'F{linha}'].value), cor='9',
                        instru='08 - Educação Superior Incompleta',
                        nacional='Brasileiro(a)',
                        pai=str(sh[f'M{linha}'].value).upper(),
                        mae=str(sh[f'N{linha}'].value).upper(), endereco=str(sh[f'O{linha}'].value),
                        num='1',
                        bairro=str(sh[f'Q{linha}'].value), cep=str(sh[f'R{linha}'].value).replace('.','').replace('-',''),
                        cidade='Brasília', cid_nas='Brasília - DF',
                        uf='DF',
                        cod_municipioend=municipios['DF']['Brasília'],
                        tel=str(sh[f'U{linha}'].value).replace('(','').replace(')','').replace('.','').replace('-',''),
                        depto=depto, cargo=cargo,
                        horario=str(sh[f'AI{linha}'].value), salario=salario, tipo_contr=tipo_contr,
                        hr_sem='25', hr_mens='100',
                        est_semestre=str(sh[f'AS{linha}'].value),
                        est_turno=str(sh[f'AT{linha}'].value),
                        est_prev_conclu=str(sh[f'AU{linha}'].value),
                        est_faculdade=str(sh[f'AV{linha}'].value),
                        est_endfacul='End',
                        est_numendfacul='1',
                        est_bairroendfacul='Bairro'
                    )
                    session.add(estag_cadastrado)
                    session.commit()
                pasta = r'\192.168.0.250'
                try:
                    os.makedirs(
                        f'\\{pasta}\\rh\\01 - RH\\01 - Administração.Controles\\02 - Funcionários, Departamentos e Férias\\000 - Pastas Funcionais\\00 - ATIVOS\\0 - Estagiários\\0 - Ainda nao iniciaram\\{estag.nome}')
                    os.makedirs(
                        f'\\{pasta}\\rh\\01 - RH\\01 - Administração.Controles\\02 - Funcionários, Departamentos e Férias\\000 - Pastas Funcionais\\00 - ATIVOS\\0 - Estagiários\\0 - Ainda nao iniciaram\\{estag.nome}\\Atestados')
                    os.makedirs(
                        f'\\{pasta}\\rh\\01 - RH\\01 - Administração.Controles\\02 - Funcionários, Departamentos e Férias\\000 - Pastas Funcionais\\00 - ATIVOS\\0 - Estagiários\\0 - Ainda nao iniciaram\\{estag.nome}\\Diversos')
                    os.makedirs(
                        f'\\{pasta}\\rh\\01 - RH\\01 - Administração.Controles\\02 - Funcionários, Departamentos e Férias\\000 - Pastas Funcionais\\00 - ATIVOS\\0 - Estagiários\\0 - Ainda nao iniciaram\\{estag.nome}\\Contratuais')
                    os.makedirs(
                        f'\\{pasta}\\rh\\01 - RH\\01 - Administração.Controles\\02 - Funcionários, Departamentos e Férias\\000 - Pastas Funcionais\\00 - ATIVOS\\0 - Estagiários\\0 - Ainda nao iniciaram\\{estag.nome}\\Ferias')
                    os.makedirs(
                        f'\\{pasta}\\rh\\01 - RH\\01 - Administração.Controles\\02 - Funcionários, Departamentos e Férias\\000 - Pastas Funcionais\\00 - ATIVOS\\0 - Estagiários\\0 - Ainda nao iniciaram\\{estag.nome}\\Ponto')
                    os.makedirs(
                        f'\\{pasta}\\rh\\01 - RH\\01 - Administração.Controles\\02 - Funcionários, Departamentos e Férias\\000 - Pastas Funcionais\\00 - ATIVOS\\0 - Estagiários\\0 - Ainda nao iniciaram\\{estag.nome}\\Recibo')
                    os.makedirs(
                        f'\\{pasta}\\rh\\01 - RH\\01 - Administração.Controles\\02 - Funcionários, Departamentos e Férias\\000 - Pastas Funcionais\\00 - ATIVOS\\0 - Estagiários\\0 - Ainda nao iniciaram\\{estag.nome}\\Rescisao')
                except:
                    pass
                # abrir cadastro no dexion e atualizar informações campo a campo
                pessoa = session.query(Colaborador).filter_by(matricula=matricula).first()
                pastapessoa = f'\\{pasta}\\rh\\01 - RH\\01 - Administração.Controles\\02 - Funcionários, Departamentos e Férias\\000 - Pastas Funcionais\\00 - ATIVOS\\0 - Estagiários\\0 - Ainda nao iniciaram\\{pessoa.nome}'
                pa.click(pa.center(pa.locateOnScreen('./static/Dexion.png')))
                pa.press('alt'), pa.press('a'), pa.press('t'), t.sleep(2), pa.press(
                    'i'), t.sleep(5), pa.write(str(pessoa.matricula)), pa.press('enter'), t.sleep(20)
                pp.copy(pessoa.nome), pa.hotkey('ctrl', 'v'), pa.press('tab'), pa.write(pessoa.cpf)
                t.sleep(1), pa.press('tab', 3), pa.write(pessoa.genero), pa.press('tab'), pa.write(pessoa.cor)
                t.sleep(1), pa.press('tab', 2), pa.write(pessoa.instru)
                t.sleep(1), pa.press('tab')
                if str(pessoa.estado_civil) == 'Casado(a)':
                    pa.write('2')
                    pa.press('tab', 6)
                else:
                    pa.write('1')
                    pa.press('tab', 5)
                pa.write(datetime.strftime(datetime.strptime(pessoa.nascimento, '%Y-%m-%d %H:%M:%S'), '%d%m%Y'))
                t.sleep(1), pa.press('tab'), pp.copy(pessoa.cid_nas), pa.hotkey('ctrl', 'v'), pa.press('tab')
                t.sleep(1), pa.write(pessoa.uf), pa.press('tab'), pa.write(pessoa.cod_municipioend), pa.press('tab')
                t.sleep(1), pp.copy(pessoa.pai), pa.hotkey('ctrl', 'v'), pa.press('tab'), pa.write('105'), pa.press('tab')
                t.sleep(1), pp.copy(pessoa.mae), pa.hotkey('ctrl', 'v'), pa.press('tab'), pa.write('105')
                # # clique em documentos
                pa.click(pa.center(pa.locateOnScreen('./static/Documentos.png')))
                pa.press('tab'), pa.write(str(pessoa.rg)), pa.press('tab'), pa.write(
                    pessoa.emissor), pa.press('tab', 3), pa.write(pessoa.cod_municipioend)
                # #clique em endereço
                pa.click(pa.center(pa.locateOnScreen('./static/Endereco.png')))
                pa.press('tab', 2), pp.copy(pessoa.endereco), pa.hotkey('ctrl', 'v'), pa.press(
                    'tab'), pa.write(pessoa.num), pa.press('tab', 2)
                pp.copy(pessoa.bairro), pa.hotkey('ctrl', 'v'), pa.press('tab'), pp.copy(pessoa.cidade), pa.hotkey(
                    'ctrl', 'v'), pa.press('tab'), pa.write(pessoa.uf)
                pa.press('tab'), pa.write(pessoa.cep), pa.press('tab'), pa.write(pessoa.cod_municipioend), pa.press(
                    'tab'), pa.write(str(pessoa.tel)), pa.press('tab', 2), pa.write(pessoa.email)
                # #clique em dados contratuais
                pa.click(pa.center(pa.locateOnScreen('./static/Contratuais.png')))
                pa.press('tab'), pa.write(str(pessoa.admiss).replace('/', '')), t.sleep(1)
                pa.press('tab'), pa.write('9')
                pa.press('tab', 7), pa.write('n'), pa.press('tab'), pa.write('4')
                pa.press('tab'), pa.write('Ed. Fisica')
                pa.press('tab', 2), pa.write(str(int(str(pessoa.admiss).replace('/', ''))+2).zfill(8))
                # #clique em instituição de ensino
                try:
                    pa.click(pa.center(pa.locateOnScreen('./static/faculdade.png')))
                except:
                    t.sleep(3)
                    pa.click(pa.center(pa.locateOnScreen('./static/faculdade.png')))
                pa.press('tab'), pp.copy(pessoa.est_faculdade), pa.hotkey('ctrl', 'v')
                pa.press('tab'), pp.copy(pessoa.est_endfacul), pa.hotkey('ctrl', 'v')
                pa.press('tab'), pp.copy(pessoa.est_numendfacul), pa.hotkey('ctrl', 'v')
                pa.press('tab'), pp.copy(pessoa.est_bairroendfacul), pa.hotkey('ctrl', 'v')
                # #clique em Outros
                try:
                    pa.click(pa.center(pa.locateOnScreen('./static/Outros.png')))
                except:
                    t.sleep(5)
                    pa.click(pa.center(pa.locateOnScreen('./static/Outros.png')))
                t.sleep(2), pa.write('CARGO GERAL')
                pa.press('tab'), pa.write(pessoa.cargo)
                t.sleep(1), pa.press('tab'), pa.write(pessoa.salario), pa.press('tab')
                if str(pessoa.tipo_contr) == 'Horista':
                    pa.press('1')
                else:
                    pa.press('5')
                pa.press('tab', 2), t.sleep(1), pa.write(str(pessoa.hr_sem)), pa.press('tab'), t.sleep(1)
                pa.write(str(pessoa.hr_mens))
                # #clique em eventos trabalhistas
                pa.click(pa.center(pa.locateOnScreen('./static/EVTrab.png')))
                t.sleep(1)
                # #clique em lotação
                pa.click(pa.center(pa.locateOnScreen('./static/Lotacoes.png')))
                pa.press('tab'), pa.press('tab'), pa.write('i'), pa.write(str(pessoa.admiss).replace('/', ''))
                t.sleep(1), pa.press('enter'), t.sleep(1)
                pp.copy(lotacao[f'{pessoa.depto}']), pa.hotkey('ctrl', 'v'), pa.press('enter'), pa.write('f')
                pa.press('tab'), pa.write('4')
                # #clique em salvar lotação
                pa.click(pa.center(pa.locateOnScreen('./static/Salvarbtn.png'))), t.sleep(1)
                # #clique em fechar lotação
                try:
                    pa.click(pa.center(pa.locateOnScreen('./static/Fecharlot.png'))), t.sleep(1)
                except:
                    t.sleep(4)
                    pa.click(pa.center(pa.locateOnScreen('./static/Fecharlot.png')))
                # #clique em Compatibilidade
                pa.click(pa.center(pa.locateOnScreen('./static/Compatibilidade3.png'))), t.sleep(1)
                # #clique em Compatibilidade de novo
                pa.click(pa.center(pa.locateOnScreen('./static/Compatibilidade2.png'))), t.sleep(1)
                pa.press('tab', 2), pa.write('9')
                # #clique em Salvar
                pa.click(pa.center(pa.locateOnScreen('./static/Salvarcadastro.png'))), t.sleep(10)
                # #clique em fechar novo cadastro
                pa.click(pa.center(pa.locateOnScreen('./static/Fecharnovo1.png'))), t.sleep(2)
                # #clique em fechar trabalhadores
                pa.click(pa.center(pa.locateOnScreen('./static/Fechartrab1.png'))), t.sleep(0.5)
                os.rename(pastapessoa,
                          f'\\{pasta}\\rh\\01 - RH\\01 - Administração.Controles\\02 - Funcionários, Departamentos e Férias\\000 - Pastas Funcionais\\00 - ATIVOS\\0 - Estagiários\\{pessoa.nome}')
                tkinter.messagebox.showinfo(
                    title='Cadastro ok!',
                    message='Cadastro realizado com sucesso!'
                )
        else:
            if ondestou == 0:
                # Editando o cadastro na Cia
                hoje = datetime.today()
                wb = l_w(caminho)
                sh = wb['Respostas ao formulário 1']
                num, name = nome.strip().split(' - ')
                linha = int(num)
                lotacao = {
                    'Unidade Park Sul - qualquer departamento': ['0013', 'Thais Feitosa',
                                                                 'thais.morais@ciaathletica.com.br',
                                                                 'Líder Park Sul'],
                    'Kids': ['0010', 'Cindy Stefanie', 'cindy.neves@ciaathletica.com.br', 'Líder Kids'],
                    'Musculação': ['0007', 'Aline Kanyó', 'aline.kanyo@soucia.com.br', 'Líder Musculação'],
                    'Esportes e Lutas': ['0008', 'Morgana Rossini', 'morganalourenco@yahoo.com.br', 'Líder Natação'],
                    'Crossfit': ['0012', 'Guilherme Salles', 'gmoreirasalles@gmail.com', 'Líder Crossfit'],
                    'Ginástica': ['0006', 'Hugo Albuquerque', 'hugo.albuquerque@ciaathletica.com.br',
                                  'Líder Ginástica'],
                    'Gestantes': ['0006', 'Hugo Albuquerque', 'hugo.albuquerque@ciaathletica.com.br',
                                  'Líder Ginástica'],
                    'Recepção': ['0003', 'Paulo Renato', 'paulo.simoes@ciaathletica.com.br', 'Gerente Vendas'],
                    'Administrativo': ['0001', 'Felipe Rodrigues', 'felipe.rodrigues@ciaathletica.com.br',
                                       'Gerente RH'],
                    'Manutenção': ['0004', 'José Aparecido', 'aparecido.grota@ciaathletica.com.br',
                                   'Gerente Manutenção'],
                }
                cadastro = {'nome': str(sh[f"C{linha}"].value).title().strip(), 'nasc_ed': sh[f"D{linha}"].value,
                            'genero': str(sh[f"E{linha}"].value), 'est_civ': str(sh[f"F{linha}"].value),
                            'pai': str(sh[f"M{linha}"].value), 'mae': str(sh[f"N{linha}"].value),
                            'end': str(sh[f"O{linha}"].value),
                            'num': str(sh[f"P{linha}"].value), 'bairro': str(sh[f"Q{linha}"].value),
                            'cep': str(sh[f"R{linha}"].value).replace('.', '').replace('-', ''),
                            'cid_end': str(sh[f"S{linha}"].value), 'uf_end': str(sh[f"T{linha}"].value),
                            'tel': str(sh[f"U{linha}"].value).replace('(', '').replace(')', '').replace('-',
                                                                                                        '').replace(' ',
                                                                                                                    ''),
                            'mun_end': str(sh[f"AP{linha}"].value),
                            'cpf': str(sh[f"V{linha}"].value).strip().replace('.', '').replace('-', '').replace(' ',
                                                                                                                '').zfill(
                                11),
                            'rg': str(sh[f"W{linha}"].value).strip().replace('.', '').replace('-', '').replace(' ', ''),
                            'emissor': str(sh[f"X{linha}"].value),
                            'lotacao': str(lotacao[f'{sh[f"AG{linha}"].value}'][0]).zfill(4),
                            'cargo': str(sh[f"AH{linha}"].value), 'horario': str(sh[f"AI{linha}"].value),
                            'email': str(sh[f"B{linha}"].value).strip(),
                            'admissao_ed': str(sh[f"AL{linha}"].value),
                            'faculdade': str(sh[f"AV{linha}"].value), 'semestre': str(sh[f"AS{linha}"].value),
                            'turno': str(sh[f"AT{linha}"].value), 'conclusao': str(sh[f"AU{linha}"].value),
                            'salario': str(sh[f"AM{linha}"].value),
                            'hrsemanais': str(sh[f"AQ{linha}"].value), 'hrmensais': str(sh[f"AR{linha}"].value)}
            else:
                # Editando o cadastro em Casa
                wb = l_w(caminho)
                sh = wb['Respostas ao formulário 1']
                num, name = nome.strip().split(' - ')
                linha = int(num)
                lotacao = {'UNIDADE PARK SUL - QUALQUER DEPARTAMENTO': '0013', 'KIDS': '0010', 'MUSCULAÇÃO': '0007',
                           'ESPORTES E LUTAS': '0008', 'CROSSFIT': '0012', 'GINÁSTICA': '0006', 'GESTANTES': '0008',
                           'RECEPÇÃO': '0003',
                           'FINANCEIRO': '0001', 'TI': '0001', 'MARKETING': '0001', 'MANUTENÇÃO': '0004'}
                if str(sh[f'E{linha}'].value) == 'Masculino':
                    cargo = 'ESTAGIARIO'
                else:
                    cargo = 'ESTAGIARIA'

                estag = session.query(Colaborador).filter_by(matricula=matricula).first()
                if estag:
                    pass
                else:
                    estag_cadastrado = Colaborador(
                        matricula=matricula, nome=name.upper(), admiss=admissao,
                        nascimento=str(sh[f'D{linha}'].value),
                        cpf=str(int(sh[f'V{linha}'].value)).zfill(11),
                        rg=str(int(sh[f'W{linha}'].value)),
                        emissor='SSP/DF', email=str(sh[f'B{linha}'].value),
                        genero=str(sh[f'E{linha}'].value),
                        estado_civil=str(sh[f'F{linha}'].value), cor='9',
                        instru='08 - Educação Superior Incompleta',
                        nacional='Brasileiro(a)',
                        pai=str(sh[f'M{linha}'].value).upper(),
                        mae=str(sh[f'N{linha}'].value).upper(), endereco=str(sh[f'O{linha}'].value),
                        num='1',
                        bairro=str(sh[f'Q{linha}'].value), cep=str(sh[f'R{linha}'].value).replace('.','').replace('-',''),
                        cidade='Brasília', cid_nas='Brasília - DF',
                        uf='DF',
                        cod_municipioend=municipios['DF']['Brasília'],
                        tel=str(sh[f'U{linha}'].value).replace('(','').replace(')','').replace('.','').replace('-',''),
                        depto=depto, cargo=cargo,
                        horario=str(sh[f'AI{linha}'].value), salario=salario, tipo_contr=tipo_contr, 
                        hr_sem='25', hr_mens='100',
                        est_semestre=str(sh[f'AS{linha}'].value),
                        est_turno=str(sh[f'AT{linha}'].value),
                        est_prev_conclu=str(sh[f'AU{linha}'].value),
                        est_faculdade=str(sh[f'AV{linha}'].value),
                        est_endfacul='End',
                        est_numendfacul='1',
                        est_bairroendfacul='Bairro'
                    )
                    session.add(estag_cadastrado)
                    session.commit()
                pasta = r'\192.168.0.250'
                try:
                    os.makedirs(
                        f'\\{pasta}\\rh\\01 - RH\\01 - Administração.Controles\\02 - Funcionários, Departamentos e Férias\\000 - Pastas Funcionais\\00 - ATIVOS\\0 - Estagiários\\0 - Ainda nao iniciaram\\{estag.nome}')
                    os.makedirs(
                        f'\\{pasta}\\rh\\01 - RH\\01 - Administração.Controles\\02 - Funcionários, Departamentos e Férias\\000 - Pastas Funcionais\\00 - ATIVOS\\0 - Estagiários\\0 - Ainda nao iniciaram\\{estag.nome}\\Atestados')
                    os.makedirs(
                        f'\\{pasta}\\rh\\01 - RH\\01 - Administração.Controles\\02 - Funcionários, Departamentos e Férias\\000 - Pastas Funcionais\\00 - ATIVOS\\0 - Estagiários\\0 - Ainda nao iniciaram\\{estag.nome}\\Diversos')
                    os.makedirs(
                        f'\\{pasta}\\rh\\01 - RH\\01 - Administração.Controles\\02 - Funcionários, Departamentos e Férias\\000 - Pastas Funcionais\\00 - ATIVOS\\0 - Estagiários\\0 - Ainda nao iniciaram\\{estag.nome}\\Contratuais')
                    os.makedirs(
                        f'\\{pasta}\\rh\\01 - RH\\01 - Administração.Controles\\02 - Funcionários, Departamentos e Férias\\000 - Pastas Funcionais\\00 - ATIVOS\\0 - Estagiários\\0 - Ainda nao iniciaram\\{estag.nome}\\Ferias')
                    os.makedirs(
                        f'\\{pasta}\\rh\\01 - RH\\01 - Administração.Controles\\02 - Funcionários, Departamentos e Férias\\000 - Pastas Funcionais\\00 - ATIVOS\\0 - Estagiários\\0 - Ainda nao iniciaram\\{estag.nome}\\Ponto')
                    os.makedirs(
                        f'\\{pasta}\\rh\\01 - RH\\01 - Administração.Controles\\02 - Funcionários, Departamentos e Férias\\000 - Pastas Funcionais\\00 - ATIVOS\\0 - Estagiários\\0 - Ainda nao iniciaram\\{estag.nome}\\Recibo')
                    os.makedirs(
                        f'\\{pasta}\\rh\\01 - RH\\01 - Administração.Controles\\02 - Funcionários, Departamentos e Férias\\000 - Pastas Funcionais\\00 - ATIVOS\\0 - Estagiários\\0 - Ainda nao iniciaram\\{estag.nome}\\Rescisao')
                except:
                    pass
                # abrir cadastro no dexion e atualizar informações campo a campo
                pessoa = session.query(Colaborador).filter_by(matricula=matricula).first()
                pa.click(pa.center(pa.locateOnScreen('./static/Dexion.png')))
                pa.press('alt'), pa.press('a'), pa.press('t'), t.sleep(2), pa.press(
                    'a'), t.sleep(5), pa.write(str(pessoa.matricula)), pa.press('enter'), t.sleep(20)
                pp.copy(pessoa.nome), pa.hotkey('ctrl', 'v'), pa.press('tab'), pa.write(pessoa.cpf)
                t.sleep(1), pa.press('tab', 3), pa.write(pessoa.genero), pa.press('tab'), pa.write(pessoa.cor)
                t.sleep(1), pa.press('tab', 2), pa.write(pessoa.instru)
                t.sleep(1), pa.press('tab')
                if str(pessoa.estado_civil) == 'Casado(a)':
                    pa.write('2')
                    pa.press('tab', 6)
                else:
                    pa.write('1')
                    pa.press('tab', 5)
                pa.write(datetime.strftime(datetime.strptime(pessoa.nascimento, '%Y-%m-%d %H:%M:%S'), '%d%m%Y'))
                t.sleep(1), pa.press('tab'), pp.copy(pessoa.cid_nas), pa.hotkey('ctrl', 'v'), pa.press('tab')
                t.sleep(1), pa.write(pessoa.uf), pa.press('tab'), pa.write(pessoa.cod_municipioend), pa.press('tab')
                t.sleep(1), pp.copy(pessoa.pai), pa.hotkey('ctrl', 'v'), pa.press('tab'), pa.write('105'), pa.press('tab')
                t.sleep(1), pp.copy(pessoa.mae), pa.hotkey('ctrl', 'v'), pa.press('tab'), pa.write('105')
                # # clique em documentos
                pa.click(pa.center(pa.locateOnScreen('./static/Documentos.png')))
                pa.press('tab'), pa.write(str(pessoa.rg)), pa.press('tab'), pa.write(
                    pessoa.emissor), pa.press('tab', 3), pa.write(pessoa.cod_municipioend)
                # #clique em endereço
                pa.click(pa.center(pa.locateOnScreen('./static/Endereco.png')))
                pa.press('tab', 2), pp.copy(pessoa.endereco), pa.hotkey('ctrl', 'v'), pa.press(
                    'tab'), pa.write(pessoa.num), pa.press('tab', 2)
                pp.copy(pessoa.bairro), pa.hotkey('ctrl', 'v'), pa.press('tab'), pp.copy(pessoa.cidade), pa.hotkey(
                    'ctrl', 'v'), pa.press('tab'), pa.write(pessoa.uf)
                pa.press('tab'), pa.write(pessoa.cep), pa.press('tab'), pa.write(pessoa.cod_municipioend), pa.press(
                    'tab'), pa.write(str(pessoa.tel)), pa.press('tab', 2), pa.write(pessoa.email)
                # #clique em dados contratuais
                pa.click(pa.center(pa.locateOnScreen('./static/Contratuais.png')))
                pa.press('tab'), pa.write(str(pessoa.admiss).replace('/', '')), t.sleep(1)
                pa.press('tab'), pa.write('9')
                pa.press('tab', 7), pa.write('n'), pa.press('tab'), pa.write('4')
                pa.press('tab'), pa.write('Ed. Fisica')
                pa.press('tab', 2), pa.write(str(int(str(pessoa.admiss).replace('/', ''))+2).zfill(8))
                # #clique em instituição de ensino
                pa.click(pa.center(pa.locateOnScreen('./static/faculdade.png')))
                pa.press('tab'), pp.copy(pessoa.est_faculdade), pa.hotkey('ctrl', 'v')
                pa.press('tab'), pp.copy(pessoa.est_endfacul), pa.hotkey('ctrl', 'v')
                pa.press('tab'), pp.copy(pessoa.est_numendfacul), pa.hotkey('ctrl', 'v')
                pa.press('tab'), pp.copy(pessoa.est_bairroendfacul), pa.hotkey('ctrl', 'v')
                # #clique em Outros
                try:
                    pa.click(pa.center(pa.locateOnScreen('./static/Outros.png')))
                except:
                    t.sleep(5)
                    pa.click(pa.center(pa.locateOnScreen('./static/Outros.png')))
                t.sleep(2), pa.write('CARGO GERAL')
                pa.press('tab'), pa.write(pessoa.cargo)
                t.sleep(1), pa.press('tab'), pa.write(pessoa.salario), pa.press('tab')
                if str(pessoa.tipo_contr) == 'Horista':
                    pa.press('1')
                else:
                    pa.press('5')
                pa.press('tab', 2), t.sleep(1), pa.write(str(pessoa.hr_sem)), pa.press('tab'), t.sleep(1)
                pa.write(str(pessoa.hr_mens))
                # #clique em Compatibilidade
                pa.click(pa.center(pa.locateOnScreen('./static/Compatibilidade3.png'))), t.sleep(1)
                # #clique em Compatibilidade de novo
                pa.click(pa.center(pa.locateOnScreen('./static/Compatibilidade2.png'))), t.sleep(1)
                pa.press('tab', 2), pa.write('9')
                # #clique em Salvar
                pa.click(pa.center(pa.locateOnScreen('./static/Salvarcadastro.png'))), t.sleep(10)
                # #clique em fechar novo cadastro
                pa.click(pa.center(pa.locateOnScreen('./static/Fecharnovo1.png'))), t.sleep(2)
                # #clique em fechar trabalhadores
                pa.click(pa.center(pa.locateOnScreen('./static/Fechartrab1.png'))), t.sleep(0.5)
                tkinter.messagebox.showinfo(
                    title='Cadastro ok!',
                    message='Cadastro editado com sucesso!'
                )


caminhoaut = StringVar()


def cadastrar_autonomo(caminhoaut, nomeaut, matriculaaut, admissaoaut, cargoaut, deptoaut, ondeaut):
    # Cadastro iniciado em casa
    wb = l_w(caminhoaut)
    sh = wb['Respostas ao formulário 1']
    num, name = nomeaut.strip().split(' - ')
    linha = int(num)
    lotacao = {'UNIDADE PARK SUL - QUALQUER DEPARTAMENTO': '0013', 'KIDS': '0010', 'MUSCULAÇÃO': '0007',
               'ESPORTES E LUTAS': '0008', 'CROSSFIT': '0012', 'GINÁSTICA': '0006', 'GESTANTES': '0008',
               'RECEPÇÃO': '0003',
               'FINANCEIRO': '0001', 'TI': '0001', 'MARKETING': '0001', 'MANUTENÇÃO': '0004'}
    aut = session.query(Colaborador).filter_by(matricula=matriculaaut).first()
    if aut:
        pass
    else:
        aut_cadastrado = Colaborador(
            matricula=matriculaaut, nome=name.upper(), admiss=admissaoaut,
            nascimento=str(sh[f'D{linha}'].value), pis=str(sh[f'S{linha}'].value).replace('.','').replace('-','').zfill(11),
            cpf=str(int(sh[f'P{linha}'].value)).zfill(11),
            rg=str(int(sh[f'Q{linha}'].value)),
            emissor='SSP/DF', email=str(sh[f'B{linha}'].value),
            genero=str(sh[f'E{linha}'].value), cor='9',
            instru=str(sh[f'F{linha}'].value),
            nacional='Brasileiro(a)', estado_civil='Solteiro(a)',
            endereco=str(sh[f'I{linha}'].value),
            num=str(sh[f'J{linha}'].value),
            bairro=str(sh[f'K{linha}'].value), cep=str(sh[f'L{linha}'].value).replace('.', '').replace('-', ''),
            cidade=str(sh[f'M{linha}'].value), cid_nas='Brasília - DF', uf='DF',
            cod_municipioend=municipios['DF']['Brasília'],
            tel=str(sh[f'O{linha}'].value).replace('(', '').replace(')', '').replace('.', '').replace('-', ''),
            depto=deptoaut, cargo=cargoaut,
        )
        session.add(aut_cadastrado)
        session.commit()
    # abrir cadastro no dexion e atualizar informações campo a campo
    pessoa = session.query(Colaborador).filter_by(matricula=matriculaaut).first()
    pa.click(pa.center(pa.locateOnScreen('./static/Dexion.png')))
    pa.press('alt'), pa.press('a'), pa.press('t'), t.sleep(2), pa.press(
        'i'), t.sleep(5), pa.write(str(pessoa.matricula)), pa.press('enter'), t.sleep(20)
    pp.copy(pessoa.nome), pa.hotkey('ctrl', 'v'), pa.press('tab'), pa.write(pessoa.cpf)
    t.sleep(1), pa.press('tab', 3), pa.write(pessoa.genero), pa.press('tab'), pa.write(pessoa.cor)
    t.sleep(1), pa.press('tab', 2), pa.write(pessoa.instru)
    t.sleep(1), pa.press('tab')
    if str(pessoa.estado_civil) == 'Casado(a)':
        pa.write('2')
        pa.press('tab', 6)
    else:
        pa.write('1')
        pa.press('tab', 5)
    pa.write(datetime.strftime(datetime.strptime(pessoa.nascimento, '%Y-%m-%d %H:%M:%S'), '%d%m%Y'))
    t.sleep(1), pa.press('tab'), pp.copy(pessoa.cid_nas), pa.hotkey('ctrl', 'v')
    # # clique em documentos
    pa.click(pa.center(pa.locateOnScreen('./static/Documentos.png')))
    pa.press('tab'), pa.write(str(pessoa.rg)), pa.press('tab'), pa.write(
        pessoa.emissor), pa.press('tab', 3), pa.write(pessoa.cod_municipioend), pa.press('tab')
    pa.write(pessoa.pis)
    # #clique em endereço
    pa.click(pa.center(pa.locateOnScreen('./static/Endereco.png')))
    pa.press('tab', 2), pp.copy(pessoa.endereco), pa.hotkey('ctrl', 'v'), pa.press(
        'tab'), pa.write(pessoa.num), pa.press('tab', 2)
    pp.copy(pessoa.bairro), pa.hotkey('ctrl', 'v'), pa.press('tab'), pp.copy(pessoa.cidade), pa.hotkey(
        'ctrl', 'v'), pa.press('tab'), pa.write(pessoa.uf)
    pa.press('tab'), pa.write(pessoa.cep), pa.press('tab'), pa.write(pessoa.cod_municipioend), pa.press(
        'tab'), pa.write(str(pessoa.tel)), pa.press('tab', 2), pa.write(pessoa.email)
    # #clique em dados contratuais
    pa.click(pa.center(pa.locateOnScreen('./static/Contratuais.png')))
    pa.press('tab'), pa.write(str(pessoa.admiss).replace('/', '')), t.sleep(1)
    pa.press('tab'), pa.write('7')
    # #clique em Outros
    try:
        pa.click(pa.center(pa.locateOnScreen('./static/Outros.png')))
    except:
        t.sleep(5)
        pa.click(pa.center(pa.locateOnScreen('./static/Outros.png')))
    t.sleep(2), pa.write('CARGO GERAL')
    pa.press('tab'), pa.write(pessoa.cargo)
    # #clique em eventos trabalhistas
    pa.click(pa.center(pa.locateOnScreen('./static/EVTrab.png')))
    t.sleep(1)
    # #clique em lotação
    pa.click(pa.center(pa.locateOnScreen('./static/Lotacoes.png')))
    pa.press('tab'), pa.press('tab'), pa.write('i'), pa.write(str(pessoa.admiss).replace('/', ''))
    t.sleep(1), pa.press('enter'), t.sleep(1)
    pp.copy(lotacao[f'{pessoa.depto}']), pa.hotkey('ctrl', 'v'), pa.press('enter'), pa.write('3')
    # #clique em salvar lotação
    pa.click(pa.center(pa.locateOnScreen('./static/Salvarbtn.png'))), t.sleep(1)
    # #clique em fechar lotação
    try:
        pa.click(pa.center(pa.locateOnScreen('./static/Fecharlot.png'))), t.sleep(1)
    except:
        t.sleep(4)
        pa.click(pa.center(pa.locateOnScreen('./static/Fecharlot.png')))
    # #clique em Compatibilidade
    pa.click(pa.center(pa.locateOnScreen('./static/Compatibilidade3.png'))), t.sleep(1)
    # #clique em Compatibilidade de novo
    pa.click(pa.center(pa.locateOnScreen('./static/Compatibilidade2.png'))), t.sleep(1)
    pa.press('tab', 2), pa.write('13')
    # #clique em Salvar
    pa.click(pa.center(pa.locateOnScreen('./static/Salvarcadastro.png'))), t.sleep(10)
    # #clique em fechar novo cadastro
    pa.click(pa.center(pa.locateOnScreen('./static/Fecharnovo1.png'))), t.sleep(2)
    # #clique em fechar trabalhadores
    pa.click(pa.center(pa.locateOnScreen('./static/Fechartrab1.png'))), t.sleep(0.5)
    tkinter.messagebox.showinfo(
        title='Cadastro ok!',
        message='Cadastro realizado com sucesso!'
    )


def validarpis(local, nome):
    wb = l_w(local, read_only=False)
    sh = wb['Respostas ao formulário 1']
    num, name = nome.strip().split(' - ')
    x = int(num)
    pis = str(sh[f'S{x}'].value).replace('-','').replace('.','').zfill(11)
    v1 = int(pis[0]) * 3
    v2 = int(pis[1]) * 2
    v3 = int(pis[2]) * 9
    v4 = int(pis[3]) * 8
    v5 = int(pis[4]) * 7
    v6 = int(pis[5]) * 6
    v7 = int(pis[6]) * 5
    v8 = int(pis[7]) * 4
    v9 = int(pis[8]) * 3
    v10 = int(pis[9]) * 2
    d = int(pis[10])
    soma = v1 + v2 + v3 + v4 + v5 + v6 + v7 + v8 + v9 + v10
    divisao = soma % 11
    resultado = 11 - divisao
    if resultado != d:
        if resultado == 10 & d == 0 | resultado == 11 & d == 0:
            pass
        else:
            tkinter.messagebox.showinfo(
                title='Erro!',
                message='PIS inválido!'
            )
    else:
        tkinter.messagebox.showinfo(
            title='Ok!',
            message='PIS ok!'
        )


def desligar(estag, func, apedido, acordo, mandado, comaviso, semaviso):
    pass


def selecionarfunc():
    try:
        caminhoplan = tkinter.filedialog.askopenfilename(title='Planilha Funcionários')
        caminho.set(str(caminhoplan))
    except ValueError:
        pass


def selecionarest():
    try:
        caminhoplanest = tkinter.filedialog.askopenfilename(title='Planilha Estagiários')
        caminhoest.set(str(caminhoplanest))
    except ValueError:
        pass


def selecionaraut():
    try:
        caminhoplanaut = tkinter.filedialog.askopenfilename(title='Planilha Autônomos')
        caminhoaut.set(str(caminhoplanaut))
    except ValueError:
        pass


funcionario = Frame(my_notebook, width=10, height=20)
ttk.Label(funcionario, width=40, text="Escolher planilha de novos funcionários").grid(column=1, row=1, padx=25, pady=1, sticky=W)
ttk.Button(funcionario, text="Escolha a planilha", command=selecionarfunc).grid(column=1, row=1, padx=350, pady=1, sticky=W)
nome = StringVar()
horario = StringVar()
cargo = StringVar()
departamento = StringVar()
tipocontr = StringVar()
nomesplan = []
# aparecer dropdown com nomes da plan
labelnome = ttk.Label(funcionario, width=20, text="Nome:")
labelnome.grid(column=1, row=10, padx=25, pady=1, sticky=W)
combonome = ttk.Combobox(funcionario, values=nomesplan, textvariable=nome, width=50)
combonome.grid(column=1, row=10, padx=125, pady=1, sticky=W)
# aparecer entry para preencher matricula
labelmatr = ttk.Label(funcionario, width=20, text="Matrícula:")
labelmatr.grid(column=1, row=11, padx=25, pady=1, sticky=W)
entrymatr = ttk.Entry(funcionario, width=20)
entrymatr.grid(column=1, row=11, padx=125, pady=1, sticky=W)
# aparecer entry para preencher admissao
labeladmiss = ttk.Label(funcionario, width=20, text="Admissão:")
labeladmiss.grid(column=1, row=12, padx=25, pady=1, sticky=W)
entryadmiss = ttk.Entry(funcionario, width=20)
entryadmiss.grid(column=1, row=12, padx=125, pady=1, sticky=W)
# aparecer horario preenchido e dropdown para escolher horario
labelhor = ttk.Label(funcionario, width=55, text="Horário preenchido: ")
labelhor.grid(column=1, row=14, padx=25, pady=1, sticky=W)
combohor = ttk.Combobox(funcionario, values=horarios, textvariable=horario, width=50)
combohor.grid(column=1, row=15, padx=125, pady=1, sticky=W)
# aparecer entry para preencher salario
labelsal = ttk.Label(funcionario, width=20, text="Salário:")
labelsal.grid(column=1, row=16, padx=25, pady=1, sticky=W)
entrysal = ttk.Entry(funcionario, width=20)
entrysal.grid(column=1, row=16, padx=125, pady=1, sticky=W)
# aparecer dropdown para escolher cargo
labelcargo = ttk.Label(funcionario, width=20, text="Cargo")
labelcargo.grid(column=1, row=18, padx=25, pady=1, sticky=W)
combocargo = ttk.Combobox(funcionario, values=cargos, textvariable=cargo, width=50)
combocargo.grid(column=1, row=18, padx=125, pady=1, sticky=W)
# aparecer dropdown para escolher depto
labeldepto = ttk.Label(funcionario, width=20, text="Departamento:")
labeldepto.grid(column=1, row=19, padx=25, pady=1, sticky=W)
combodepto = ttk.Combobox(funcionario, values=departamentos, textvariable=departamento, width=50)
combodepto.grid(column=1, row=19, padx=125, pady=1, sticky=W)
# aparecer dropdown para escolher tipo_contr
labelcontr = ttk.Label(funcionario, width=20, text="Tipo de contrato:")
labelcontr.grid(column=1, row=21, padx=25, pady=1, sticky=W)
combocontr = ttk.Combobox(funcionario, values=tipodecontrato, textvariable=tipocontr, width=50)
combocontr.grid(column=1, row=21, padx=125, pady=1, sticky=W)
hrs = StringVar()
hrm = StringVar()
# aparecer entry para preencher hrsem
labelhrsem = ttk.Label(funcionario, width=20, text="Hrs Sem.:")
labelhrsem.grid(column=1, row=24, padx=25, pady=1, sticky=W)
entryhrsem = ttk.Entry(funcionario, width=20, textvariable=hrs)
entryhrsem.grid(column=1, row=24, padx=125, pady=1, sticky=W)
# aparecer entry para preencher hrmens
labelhrmen = ttk.Label(funcionario, width=20, text="Hrs Mens.:")
labelhrmen.grid(column=1, row=25, padx=25, pady=1, sticky=W)
entryhrmen = ttk.Entry(funcionario, width=20, textvariable=hrm)
entryhrmen.grid(column=1, row=25, padx=125, pady=1, sticky=W)
edicao = IntVar()
editar = ttk.Checkbutton(funcionario, text='Editar cadastro feito manualmente.', variable=edicao)
editar.grid(column=1, row=26, padx=26, pady=1, sticky=W)
feitonde = IntVar()
onde = ttk.Checkbutton(funcionario, text='Cadastro realizado fora da Cia.', variable=feitonde)
onde.grid(column=1, row=27, padx=26, pady=1, sticky=W)


def mostrarhorario(event):
    nome = event.widget.get()
    num, name = nome.split(' - ')
    linha = int(num)
    planwb = l_w(caminho.get())
    plansh = planwb['Respostas ao formulário 1']
    labelhor.config(text='Horário preenchido: ' + plansh[f'AI{linha+1}'].value)


combonome.bind("<<ComboboxSelected>>", mostrarhorario)


def carregarfunc(local):
    planwb = l_w(local)
    plansh = planwb['Respostas ao formulário 1']
    lista = []
    for x, pessoa in enumerate(plansh):
        lista.append(f'{x+1} - {pessoa[2].value}')
    combonome.config(values=lista)


ttk.Button(funcionario, text="Carregar planilha", command=lambda: [carregarfunc(caminho.get())]).grid(column=1, row=9, padx=350, pady=25, sticky=W)
ttk.Button(funcionario, width=20, text="Cadastrar no Dexion",
           command=lambda: [cadastro_funcionario(caminho.get(),edicao.get(),feitonde.get(),combonome.get(),
                                                 entrymatr.get(), entryadmiss.get(), combohor.get(),entrysal.get(),
                                                 combocargo.get(), combodepto.get(),combocontr.get(), hrs.get(),
                                                 hrm.get())]).grid(column=1, row=28, padx=520, pady=1, sticky=W)
ttk.Button(funcionario, width=20, text="Salvar Docs",
           command=lambda: [salvadocsfunc(entrymatr.get())]).grid(column=1, row=29, padx=520, pady=1, sticky=W)
ttk.Button(funcionario, width=20, text="Enviar e-mails",
           command=lambda: [enviaemailsfunc(entrymatr.get())]).grid(column=1, row=30, padx=520, pady=1, sticky=W)
funcionario.pack(fill='both', expand=0)

estagiario = Frame(my_notebook, width=60, height=50)
ttk.Label(estagiario, width=40, text="Escolher planilha de novos estagiários").grid(column=1, row=2, padx=25, pady=1, sticky=W)
ttk.Button(estagiario, text="Escolha a planilha", command=selecionarest).grid(column=1, row=2, padx=350, pady=1, sticky=W)
labelnomest = ttk.Label(estagiario, width=20, text="Nome:")
labelnomest.grid(column=1, row=10, padx=25, pady=1, sticky=W)
combonomest = ttk.Combobox(estagiario, values=nomesplan, textvariable=nome, width=50)
combonomest.grid(column=1, row=10, padx=125, pady=1, sticky=W)
# aparecer entry para preencher matricula
labelmatrest = ttk.Label(estagiario, width=20, text="Matrícula:")
labelmatrest.grid(column=1, row=11, padx=25, pady=1, sticky=W)
entrymatrest = ttk.Entry(estagiario, width=20)
entrymatrest.grid(column=1, row=11, padx=125, pady=1, sticky=W)
# aparecer entry para preencher admissao
labeladmissest = ttk.Label(estagiario, width=20, text="Admissão:")
labeladmissest.grid(column=1, row=12, padx=25, pady=1, sticky=W)
entryadmissest = ttk.Entry(estagiario, width=20)
entryadmissest.grid(column=1, row=12, padx=125, pady=1, sticky=W)
# aparecer dropdown para escolher depto
labeldeptoest = ttk.Label(estagiario, width=20, text="Departamento:")
labeldeptoest.grid(column=1, row=19, padx=25, pady=1, sticky=W)
combodeptoest = ttk.Combobox(estagiario, values=departamentos, textvariable=departamento, width=50)
combodeptoest.grid(column=1, row=19, padx=125, pady=1, sticky=W)
solicitarest = IntVar()
solictest = ttk.Checkbutton(estagiario, text='Apenas solicitar contrato.', variable=solicitarest)
solictest.grid(column=1, row=25, padx=26, pady=1, sticky=W)
edicaoest = IntVar()
editarest = ttk.Checkbutton(estagiario, text='Editar cadastro feito manualmente.', variable=edicaoest)
editarest.grid(column=1, row=26, padx=26, pady=1, sticky=W)
feitondeest = IntVar()
ondeest = ttk.Checkbutton(estagiario, text='Cadastro realizado fora da Cia.', variable=feitondeest)
ondeest.grid(column=1, row=27, padx=26, pady=1, sticky=W)
cargoest = StringVar()

ttk.Button(estagiario, width=20, text="Cadastrar Funcionário",
           command=lambda: [cadastro_estagiario(solicitarest.get(), caminhoest.get(),edicaoest.get(),feitondeest.get(),combonomest.get(),
                                                 entrymatrest.get(), entryadmissest.get(), combocargo.get(), combodeptoest.get(),
                                                 combocontr.get(), hrs.get(),
                                                 hrm.get())]).grid(column=1, row=28, padx=520, pady=1, sticky=W)


def carregarest(local):
    planwb = l_w(local)
    plansh = planwb['Respostas ao formulário 1']
    lista = []
    for x, pessoa in enumerate(plansh):
        lista.append(f'{x+1} - {pessoa[2].value}')
    combonomest.config(values=lista)


ttk.Button(estagiario, text="Carregar planilha", command=lambda: [carregarest(caminhoest.get())]).grid(column=1, row=4, padx=350, pady=25, sticky=W)
estagiario.pack(fill='both', expand=0)


def carregaraut(local):
    planwb = l_w(local)
    plansh = planwb['Respostas ao formulário 1']
    lista = []
    for x, pessoa in enumerate(plansh):
        lista.append(f'{x+1} - {pessoa[2].value}')
    combonomeaut.config(values=lista)


autonomo = Frame(my_notebook, width=660, height=550)
ttk.Label(autonomo, width=40, text="Escolher planilha de autônomos").grid(column=1, row=1, padx=25, pady=1, sticky=W)
ttk.Button(autonomo, text="Escolha a planilha", command=selecionaraut).grid(column=1, row=1, padx=350, pady=1, sticky=W)
nomeaut = StringVar()
cargo = StringVar()
departamento = StringVar()
nomesplanaut = []
# aparecer dropdown com nomes da plan
labelnomeaut = ttk.Label(autonomo, width=20, text="Nome:")
labelnomeaut.grid(column=1, row=10, padx=25, pady=1, sticky=W)
combonomeaut = ttk.Combobox(autonomo, values=nomesplan, textvariable=nome, width=50)
combonomeaut.grid(column=1, row=10, padx=125, pady=1, sticky=W)
# aparecer entry para preencher matricula
labelmatraut = ttk.Label(autonomo, width=20, text="Matrícula:")
labelmatraut.grid(column=1, row=11, padx=25, pady=1, sticky=W)
entrymatraut = ttk.Entry(autonomo, width=20)
entrymatraut.grid(column=1, row=11, padx=125, pady=1, sticky=W)
# aparecer entry para preencher admissao
labeladmissaut = ttk.Label(autonomo, width=20, text="Admissão:")
labeladmissaut.grid(column=1, row=12, padx=25, pady=1, sticky=W)
entryadmissaut = ttk.Entry(autonomo, width=20)
entryadmissaut.grid(column=1, row=12, padx=125, pady=1, sticky=W)
# aparecer dropdown para escolher cargo
labelcargoaut = ttk.Label(autonomo, width=20, text="Cargo")
labelcargoaut.grid(column=1, row=18, padx=25, pady=1, sticky=W)
combocargoaut = ttk.Combobox(autonomo, values=cargos, textvariable=cargo, width=50)
combocargoaut.grid(column=1, row=18, padx=125, pady=1, sticky=W)
# aparecer dropdown para escolher depto
labeldeptoaut = ttk.Label(autonomo, width=20, text="Departamento:")
labeldeptoaut.grid(column=1, row=19, padx=25, pady=1, sticky=W)
combodeptoaut = ttk.Combobox(autonomo, values=departamentos, textvariable=departamento, width=50)
combodeptoaut.grid(column=1, row=19, padx=125, pady=1, sticky=W)
feitondeaut = IntVar()
ondeaut = ttk.Checkbutton(autonomo, text='Cadastro realizado fora da Cia.', variable=feitonde)
ondeaut.grid(column=1, row=27, padx=26, pady=1, sticky=W)
ttk.Button(autonomo, text="Carregar planilha", command=lambda: [carregaraut(caminhoaut.get())]).grid(column=1, row=9, padx=350, pady=25, sticky=W)
ttk.Button(autonomo, width=20, text="Validar PIS",
           command=lambda: [validarpis(caminhoaut.get(),combonomeaut.get())]).grid(column=1, row=10, padx=520, pady=1, sticky=W)

ttk.Button(autonomo, width=20, text="Cadastrar autônomo",
           command=lambda: [cadastrar_autonomo(caminhoaut.get(),combonomeaut.get(),
                                               entrymatraut.get(), entryadmissaut.get(), combocargoaut.get(),
                                               combodeptoaut.get(), feitondeaut.get())]).grid(column=1, row=28, padx=520, pady=1, sticky=W)
autonomo.pack(fill='both', expand=1)


def enviarcontracheque():
    pass


contracheque = Frame(my_notebook, width=660, height=550)
ttk.Label(contracheque, width=20, text="Escolher planilha de autônomos").grid(column=1, row=3, padx=25, pady=1, sticky=W)
ttk.Button(contracheque, text="Planilha Autônomos", command=selecionaraut).grid(column=1, row=3, padx=20, pady=1, sticky=E)
ttk.Button(contracheque, text="Carregar planilha", command=lambda: [enviarcontracheque]).grid(column=1, row=9, padx=20, pady=25, sticky=E)
contracheque.pack(fill='both', expand=1)


def enviarmsg():
    pass


mensagem = Frame(my_notebook, width=660, height=550)
ttk.Label(mensagem, width=20, text="Escolher planilha de autônomos").grid(column=1, row=3, padx=25, pady=1, sticky=W)
ttk.Button(mensagem, text="Planilha Autônomos", command=selecionaraut).grid(column=1, row=3, padx=20, pady=1, sticky=E)
ttk.Button(mensagem, text="Carregar planilha", command=lambda: [enviarmsg]).grid(column=1, row=9, padx=20, pady=25, sticky=E)
mensagem.pack(fill='both', expand=1)


def enviarmsgferias():
    pass


ferias = Frame(my_notebook, width=660, height=550)
ttk.Label(ferias, width=20, text="Escolher planilha de autônomos").grid(column=1, row=3, padx=25, pady=1, sticky=W)
ttk.Button(ferias, text="Planilha Autônomos", command=selecionaraut).grid(column=1, row=3, padx=20, pady=1, sticky=E)
ttk.Button(ferias, text="Carregar planilha", command=lambda: [enviarmsgferias]).grid(column=1, row=9, padx=20, pady=25, sticky=E)
ferias.pack(fill='both', expand=1)

desl = session.query(Colaborador).filter_by(desligamento=None).all()
ativos = []
for p in desl:
    ativos.append(str(p.nome).title())
adesligar = sorted(ativos)
desligado = StringVar()
desligamento = Frame(my_notebook, width=660, height=550)
labelnomdeslig = ttk.Label(desligamento, width=20, text="Nome:")
labelnomdeslig.grid(column=1, row=10, padx=25, pady=1, sticky=W)
combonomdeslig = ttk.Combobox(desligamento, values=adesligar, textvariable=desligado, width=50)
combonomdeslig.grid(column=1, row=10, padx=125, pady=1, sticky=W)
solicitardeslig = IntVar()
solictdeslig = ttk.Checkbutton(desligamento, text='Apenas solicitar contrato.', variable=solicitardeslig)
solictdeslig.grid(column=1, row=25, padx=26, pady=1, sticky=W)
edicaodeslig = IntVar()
editardeslig = ttk.Checkbutton(desligamento, text='Editar cadastro feito manualmente.', variable=edicaodeslig)
editardeslig.grid(column=1, row=26, padx=26, pady=1, sticky=W)
feitondedeslig = IntVar()
ondedeslig = ttk.Checkbutton(desligamento, text='Cadastro realizado fora da Cia.', variable=feitondedeslig)
ondedeslig.grid(column=1, row=27, padx=26, pady=1, sticky=W)
cargodeslig = StringVar()

ttk.Button(desligamento, width=20, text="Realizar desligamento",
           command=lambda: []).grid(column=1, row=28, padx=520, pady=1, sticky=W)
desligamento.pack(fill='both', expand=1)

my_notebook.add(funcionario, text='Cadastrar Funcionário')
my_notebook.add(estagiario, text='Cadastrar Estagiário')
my_notebook.add(autonomo, text='Cadastrar Autônomo')
my_notebook.add(contracheque, text='Enviar contracheque')
my_notebook.add(mensagem, text='Mensagem')
my_notebook.add(ferias, text='Férias')
my_notebook.add(desligamento, text='Desligamento')

root.config(menu=menubar)
root.mainloop()
