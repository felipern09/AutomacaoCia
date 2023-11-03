from datetime import datetime as dt, timedelta as td
import datetime
from dateutil.relativedelta import relativedelta
import holidays
import locale
from openpyxl import load_workbook as l_w
from openpyxl.styles import PatternFill, Font
import openpyxl.utils.cell
import os
import pandas as pd
import pyautogui as pa
from src.models.modelsfolha import Aula, Folha, Aulas, Faltas, Ferias, Hrcomplement, Atestado, Desligados, \
    Escala, Substituicao, enginefolha
from src.models.models import Colaborador, engine
from sqlalchemy.orm import sessionmaker
import tkinter.filedialog
from tkinter import messagebox
import tkinter.filedialog
import time as t

locale.setlocale(locale.LC_MONETARY, 'pt_BR.UTF-8')
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
locale.setlocale(locale.LC_MONETARY, 'pt_BR.UTF-8')

atestado = PatternFill(start_color='A9D08E',
                       end_color='A9D08E',
                       fill_type='solid')
falta = PatternFill(start_color='FF0000',
                    end_color='FF0000',
                    fill_type='solid')
ferias = PatternFill(start_color='9BC2E6',
                     end_color='9BC2E6',
                     fill_type='solid')
feriado = PatternFill(start_color='F4B084',
                      end_color='F4B084',
                      fill_type='solid')
fds = PatternFill(start_color='BFBFBF',
                  end_color='BFBFBF',
                  fill_type='solid')
deslig = PatternFill(start_color='454545',
                     end_color='454545',
                     fill_type='solid')
subst = PatternFill(start_color='FFFF00',
                    end_color='FFFF00',
                    fill_type='solid')
comple = PatternFill(start_color='FFC000',
                     end_color='FFC000',
                     fill_type='solid')


def confirma_folha(comp: int):
    """
    Confirm that the user wnats to proceed with procedures to post values of payroll in external aplication.
    :param comp: Month reference to payroll calculate.
    :return: Call function lancar_folha_no_dexion().
    """
    resp = messagebox.askyesno(title='Tem certeza?',
                               message=f'Tem certeza que deseja lançar a folha do mês {comp} no Dexion?')
    if resp:
        lancar_folha_no_dexion(comp)


def confirma_grade(comp: int):
    """
    Confirm that the user wnats to proceed with procedures to register the payroll.
    :param comp: Month reference to payroll calculate.
    :return: Call function salvar_planilha_soma_final().
    """
    r = messagebox.askyesno(title='Tem certeza?',
                            message=f'Tem certeza que deseja gerar a folha do mês {comp}?\n'
                                    f'Essa ação irá sobrepor qualquer arquivo de folha já salvo na pasta dessa competência.')
    if r:
        salvar_planilha_soma_final(comp)


def lancar_folha_no_dexion(competencia):
    """
    Register values of payroll in external aplication through PyAutogui package.
    :param competencia: Month reference to payroll calculate.
    :return: External aplication with payroll values registred.
    """
    while 1 < 2:
        if pa.locateOnScreen('../models/static/imgs/Dexion.png'):
            pa.click(pa.center(pa.locateOnScreen('../models/static/imgs/Dexion.png')))
            break
        else:
            t.sleep(5)
    pa.press('a'), t.sleep(2)
    hj = dt.today()
    mes = str(competencia).zfill(2)
    ano = hj.year
    mesext = {'01': 'JAN', '02': 'FEV', '03': 'MAR', '04': 'ABR', '05': 'MAI', '06': 'JUN',
              '07': 'JUL', '08': 'AGO', '09': 'SET', '10': 'OUT', '11': 'NOV', '12': 'DEZ'}
    folhagrd = os.path.abspath(rf"\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\04 - Folha de Pgto\{ano}\{mes} - {mesext[mes]}\Grades e Comissões\Lancamentos.xlsx")
    wb = l_w(folhagrd, read_only=False)

    # lançamento de faltas
    sh = wb['Faltas']
    x = 2
    while x <= len(sh['A']):
        mat = str(sh[f'A{x}'].value)
        rub = str(sh[f'C{x}'].value)
        hr = str(sh[f'D{x}'].value)
        pa.write(mat), t.sleep(3), pa.press('enter', 2), t.sleep(2.5), pa.press('i'), t.sleep(2.5), pa.write(rub)
        t.sleep(1), pa.press('enter'), t.sleep(1.5), pa.write(hr), t.sleep(1.5), pa.press('enter', 65), t.sleep(4)
        x += 1

    # deletar férias antigas
    sh = wb['DeletarFerias']
    x = 2
    while x <= len(sh['A']):
        if sh[f'A{x}'].value is not None:
            rub = ['1006', '1007', '1010', '1011', '1012', '1037']
            mat = str(sh[f'A{x}'].value)
            pa.write(mat), t.sleep(0.5), pa.press('enter', 2), t.sleep(1), pa.press('d')
            for r in rub:
                pa.write(r), t.sleep(0.8), pa.press('enter'), t.sleep(0.5), pa.press('left'), t.sleep(0.5), pa.press(
                    'enter')
            pa.press('enter', 2)
            x += 1

    # lançamento de horistas
    sh = wb['Horistas']
    x = 2
    while x <= len(sh['A']):
        if sh[f'A{x}'].value is not None:
            mat = str(sh[f'A{x}'].value)
            rub = str(sh[f'C{x}'].value)
            hr = str(sh[f'D{x}'].value)
            obshr = str(sh[f'E{x}'].value)
            pa.write(mat), t.sleep(0.5), pa.press('enter', 2), t.sleep(2.3)
            pa.press('a'), t.sleep(0.5), pa.write(rub)
            pa.press('enter'), t.sleep(0.5), pa.write(hr), t.sleep(0.5), pa.press('enter', 2)
            pa.press('enter')
            x += 1
    # lançamento de dias dobrados — estágio
    sh = wb['Compl Est']
    x = 2
    while x <= len(sh['A']):
        if sh[f'A{x}'].value is not None:
            mat = str(sh[f'A{x}'].value)
            rub = str(sh[f'C{x}'].value)
            dias = str(sh[f'D{x}'].value)
            pa.write(mat), t.sleep(0.5), pa.press('enter', 2), t.sleep(0.5), pa.press('i'), t.sleep(0.5), pa.write(rub)
            pa.press('enter'), t.sleep(0.5), pa.write(dias), t.sleep(0.5), pa.press('enter')
            pa.press('enter'), t.sleep(0.5), pa.press('enter')
            x += 1

    # lançamento de comissões
    sh = wb['Comissoes']
    x = 2
    while x <= len(sh['A']):
        if sh[f'A{x}'].value is not None:
            mat = str(sh[f'A{x}'].value)
            rub = str(sh[f'C{x}'].value)
            hr = str(sh[f'D{x}'].value)
            pa.write(mat), t.sleep(0.5), pa.press('enter', 2), t.sleep(0.5), pa.press('i'), t.sleep(0.5), pa.write(rub)
            pa.press('enter'), t.sleep(0.5), pa.write(hr), t.sleep(0.5), pa.press('enter')
            pa.press('enter'), t.sleep(0.5), pa.press('enter')
            x += 1

    # # Lançamento de adiantamento
    sh = wb['Adiantamento']
    x = 2
    while x <= len(sh['A']):
        if sh[f'A{x}'].value is not None:
            mat = str(sh[f'A{x}'].value)
            rub = '81'
            hr = str(sh[f'D{x}'].value)
            pa.write(mat), t.sleep(0.5), pa.press('enter', 2), t.sleep(0.5), pa.press('i'), t.sleep(0.5), pa.write(rub)
            pa.press('enter'), t.sleep(0.5), pa.write(hr), t.sleep(0.5), pa.press('enter')
            pa.press('enter'), t.sleep(0.5), pa.press('enter')
            x += 1

    # lançamento de desconto de VT
    sh = wb['DescontoVT']
    x = 2
    while x <= len(sh['A']):
        if sh[f'A{x}'].value is not None:
            mat = str(sh[f'A{x}'].value)
            rub = '80'
            pa.write(mat), t.sleep(0.5), pa.press('enter', 2), t.sleep(0.5), pa.press('i'), t.sleep(0.5), pa.write(rub)
            pa.press('enter'), t.sleep(0.5), pa.press('enter')
            pa.press('enter'), t.sleep(0.5), pa.press('enter')
            x += 1

    # lançamento de plano de saúde
    sh = wb['Plano']
    x = 2
    while x <= len(sh['A']):
        if sh[f'A{x}'].value is not None:
            mat = str(sh[f'A{x}'].value)
            rub = str(sh[f'C{x}'].value)
            hr = str(sh[f'D{x}'].value)
            sq = str(sh[f'E{x}'].value)
            pa.write(mat), t.sleep(0.5), pa.press('enter', 2), t.sleep(0.5), pa.press('i'), t.sleep(0.5), pa.write(rub)
            pa.press('enter'), t.sleep(0.5), pa.write(sq), t.sleep(0.5), pa.press('enter'), t.sleep(0.5), pa.write(hr)
            pa.press('enter', 3), t.sleep(0.5),
            x += 1
    while 1 < 2:
        if pa.locateOnScreen('../models/static/imgs/pyt.png'):
            pa.click(pa.center(pa.locateOnScreen('../models/static/imgs/pyt.png')))
            break
        else:
            t.sleep(5)
    tkinter.messagebox.showinfo(
        title='Folha ok!',
        message=f'Folha do mês {competencia} lançada no Dexion com sucesso!'
    )


def somar_aulas_da_grade(diasem: str, inic: datetime.datetime, fim: datetime.datetime, competencia: int, iniciograd: str, fimgrad: str) -> float:
    """
    Sum classes that have status 'Active' at the period selected for the payroll.
    :param diasem: Weekday.
    :param inic: Class start time.
    :param fim: End time of the class.
    :param competencia: Month of the payroll.
    :param iniciograd: Start day for the payroll.
    :param fimgrad: End day of the payroll.
    :return: Sum of all classes times for each day of the week.
    """
    horas = fim - inic
    hr, minut, seg = str(horas).split(':')
    igrad = dt.strptime(iniciograd, '%d/%m/%Y')
    if fimgrad is not None:
        fgrad = dt.strptime(fimgrad, '%d/%m/%Y')
    else:
        fgrad = ''
    somadia = int(hr) + int(minut) / 60 + int(seg) / 60 * 60
    somadia = round(somadia, 2)
    soma = 0
    comp = dt(day=1, month=competencia, year=dt.today().year)
    inicio = dt(day=21, month=(comp - relativedelta(months=1)).month, year=(comp - relativedelta(months=1)).year)
    fechamento = dt(day=20, month=comp.month, year=comp.year)

    def intervalo(inicio, fechamento):
        for n in range(int((fechamento - inicio).days) + 1):
            yield inicio + td(n)

    if fgrad != '':
        if diasem == 'Segunda':
            for dia in intervalo(inicio, fechamento):
                if dia.weekday() == 0 and igrad <= dia <= fgrad:
                    soma += somadia
        if diasem == 'Terça':
            for dia in intervalo(inicio, fechamento):
                if dia.weekday() == 1 and igrad <= dia <= fgrad:
                    soma += somadia
        if diasem == 'Quarta':
            for dia in intervalo(inicio, fechamento):
                if dia.weekday() == 2 and igrad <= dia <= fgrad:
                    soma += somadia
        if diasem == 'Quinta':
            for dia in intervalo(inicio, fechamento):
                if dia.weekday() == 3 and igrad <= dia <= fgrad:
                    soma += somadia
        if diasem == 'Sexta':
            for dia in intervalo(inicio, fechamento):
                if dia.weekday() == 4 and igrad <= dia <= fgrad:
                    soma += somadia
        if diasem == 'Sábado':
            for dia in intervalo(inicio, fechamento):
                if dia.weekday() == 5 and igrad <= dia <= fgrad:
                    soma += somadia
        if diasem == 'Domingo':
            for dia in intervalo(inicio, fechamento):
                if dia.weekday() == 6 and igrad <= dia <= fgrad:
                    soma += somadia
    else:
        if diasem == 'Segunda':
            for dia in intervalo(inicio, fechamento):
                if dia.weekday() == 0 and dia >= igrad:
                    soma += somadia
        if diasem == 'Terça':
            for dia in intervalo(inicio, fechamento):
                if dia.weekday() == 1 and dia >= igrad:
                    soma += somadia
        if diasem == 'Quarta':
            for dia in intervalo(inicio, fechamento):
                if dia.weekday() == 2 and dia >= igrad:
                    soma += somadia
        if diasem == 'Quinta':
            for dia in intervalo(inicio, fechamento):
                if dia.weekday() == 3 and dia >= igrad:
                    soma += somadia
        if diasem == 'Sexta':
            for dia in intervalo(inicio, fechamento):
                if dia.weekday() == 4 and dia >= igrad:
                    soma += somadia
        if diasem == 'Sábado':
            for dia in intervalo(inicio, fechamento):
                if dia.weekday() == 5 and dia >= igrad:
                    soma += somadia
        if diasem == 'Domingo':
            for dia in intervalo(inicio, fechamento):
                if dia.weekday() == 6 and dia >= igrad:
                    soma += somadia
    return round(soma, 2)


def listar_aulas_ativas(compet) -> list:
    """
    List all classes with 'Active' status.
    :return: List of active classes.
    """
    sessions = sessionmaker(bind=enginefolha)
    session = sessions()
    sessionscol = sessionmaker(bind=engine)
    sessioncol = sessionscol()
    competencia = dt(day=10, month=compet, year=dt.today().year)
    inicio = dt(day=21, month=(competencia - relativedelta(months=1)).month,
                year=(competencia - relativedelta(months=1)).year)

    aulasativasdb = session.query(Aulas).filter_by(status='Ativa').all()
    for aulat in aulasativasdb:
        pessoa = sessioncol.query(Colaborador).filter_by(nome=aulat.professor).order_by(Colaborador.matricula.desc()).first()
        if pessoa:
            if pessoa.desligamento is not None:
                if dt.strptime(pessoa.desligamento, '%d/%m/%Y') >= inicio:
                    pass
                else:
                    aulat.status = 'Inativa'
                    session.commit()

    aulasativasdb = session.query(Aulas).filter_by(status='Ativa').all()
    aula = []
    for i, a in enumerate(aulasativasdb):
        aula.append(i)
        aula[i] = Aula(a.nome, a.professor, a.departamento, a.diadasemana, a.inicio, a.fim, a.valor, a.iniciograde,
                       a.fimgrade)
        yield aula[i]
    session.close()
    return aula


def listar_departamentos_ativos() -> list:
    """
    List all departments with classes with "Active" status.
    :return: List of active departments.
    """
    sessions = sessionmaker(bind=enginefolha)
    session = sessions()
    aulasativasdb = session.query(Aulas).filter_by(status='Ativa').all()
    departamentos = []
    for i, a in enumerate(aulasativasdb):
        departamentos.append(a.departamento)
        departamentos = list(set(departamentos))
    session.close()
    return departamentos


def listar_professores_ativos() -> list:
    """
    List of all teachers who have active classes.
    :return: List of active teachers.
    """
    sessions = sessionmaker(bind=enginefolha)
    session = sessions()
    aulasativasdb = session.query(Aulas).filter_by(status='Ativa').all()
    professores = []
    for i, a in enumerate(aulasativasdb):
        professores.append(a.professor)
        professores = list(set(professores))
    session.close()
    return professores


def calcular_total_monetario_folha(compet: int) -> float:
    """
    Sum the total payroll payment.
    :param compet: Month of payroll
    :return: Payment amount.
    """
    somatorio = 0
    for al in listar_aulas_ativas(compet):
        somatorio += somar_aulas_da_grade(al.dia, al.inicio, al.fim, compet, al.iniciograde, al.fimgrade) * float(
            str(al.valor).replace(',', '.')) * al.dsr
    return round(somatorio, 2)


def somar_horas_professor(folha: Folha, prof: str, depto: str, nome: str, compet: int) -> float:
    """
    Sum total of hours for each teacher.
    :param folha: Payroll reference.
    :param prof: Teacher.
    :param depto: Department
    :param nome: Name of the class.
    :param compet: Month of payroll
    :return: The amount of hours for each teacher on this payroll.
    """
    somahoras = 0
    for aula in folha.aulas:
        if aula.professor == prof and aula.departamento == depto and aula.nome == nome:
            somahoras += somar_aulas_da_grade(aula.dia, aula.inicio, aula.fim, compet, aula.iniciograde, aula.fimgrade)
    return somahoras


def consultar_faltas(comp) -> dict:
    sessions = sessionmaker(bind=enginefolha)
    session = sessions()
    falt = session.query(Faltas).all()
    inicio = dt(day=21, month=(dt(day=1, month=comp, year=dt.today().year) - relativedelta(months=1)).month,
                year=dt.today().year)
    fim = dt(day=20, month=comp, year=dt.today().year)
    dic = {}
    for f in falt:
        if inicio <= dt.strptime(f.data, '%d/%m/%Y') <= fim:
            if f.professor in dic:
                d2 = {f.professor: {f.data: {f.departamento: f.horas}}}
                dic[f.professor] = {**dic[f.professor], **d2[f.professor]}
            else:
                d2 = {f.professor: {f.data: {f.departamento: f.horas}}}
                dic = {**dic, **d2}
    session.close()
    return dic


def consultar_ferias(comp) -> dict:
    sessions = sessionmaker(bind=enginefolha)
    session = sessions()
    fer = session.query(Ferias).all()
    inicio = dt(day=21, month=(dt(day=1, month=comp, year=dt.today().year) - relativedelta(months=1)).month,
                year=dt.today().year)
    fim = dt(day=20, month=comp, year=dt.today().year)
    dic = {}
    for f in fer:
        if inicio <= dt.strptime(f.inicio, '%d/%m/%Y') <= fim:
            if f.professor in dic:
                d2 = {f.professor: {f.departamento: {f.inicio: f.fim}}}
                dic[f.professor] = {**dic[f.professor], **d2[f.professor]}
            else:
                d2 = {f.professor: {f.departamento: {f.inicio: f.fim}}}
                dic = {**dic, **d2}
    session.close()
    return dic


def consultar_atestados(comp) -> dict:
    sessions = sessionmaker(bind=enginefolha)
    session = sessions()
    ates = session.query(Atestado).all()
    inicio = dt(day=21, month=(dt(day=1, month=comp, year=dt.today().year) - relativedelta(months=1)).month,
                year=dt.today().year)
    fim = dt(day=20, month=comp, year=dt.today().year)
    dic = {}
    for a in ates:
        if inicio <= dt.strptime(a.data, '%d/%m/%Y') <= fim:
            if a.professor in dic:
                d2 = {a.professor: {a.data: a.departamento}}
                dic[a.professor] = {**dic[a.professor], **d2[a.professor]}
            else:
                d2 = {a.professor: {a.data: a.departamento}}
                dic = {**dic, **d2}
    session.close()
    return dic


def listar_feriados(comp: int) -> list:
    inicio = dt(day=21, month=(dt(day=1, month=comp, year=dt.today().year) - relativedelta(months=1)).month,
                year=dt.today().year)
    fim = dt(day=20, month=comp, year=dt.today().year)
    # Get the Bank Holidays for the given country
    feriados = holidays.country_holidays('BR')
    # Create a list of dates between the start and end date
    intervalo_datas = pd.date_range(inicio, fim)
    # Filter the dates to only include Bank Holidays
    feriados_nacionais = [date for date in intervalo_datas if date in feriados]
    return feriados_nacionais


def consultar_substituicoes(comp: int) -> dict:
    sessions = sessionmaker(bind=enginefolha)
    session = sessions()
    subst = session.query(Substituicao).all()
    inicio = dt(day=21, month=(dt(day=1, month=comp, year=dt.today().year) - relativedelta(months=1)).month,
                year=dt.today().year)
    fim = dt(day=20, month=comp, year=dt.today().year)
    dic = {}
    for s in subst:
        if inicio <= dt.strptime(s.data, '%d/%m/%Y') <= fim:
            if s.professorsubst in dic:
                dic2 = {s.professorsubst: {s.substituto: {s.departamento: {s.data: s.horas}}}}
                dic[s.professorsubst] = {**dic[s.professorsubst], **dic2[s.professorsubst]}
            else:
                d2 = {s.professorsubst: {s.substituto: {s.departamento: {s.data: s.horas}}}}
                dic = {**dic, **d2}
    session.close()
    return dic


def consultar_desligamentos(comp: int) -> dict:
    sessions = sessionmaker(bind=enginefolha)
    session = sessions()
    desl = session.query(Desligados).all()
    inicio = dt(day=21, month=(dt(day=1, month=comp, year=dt.today().year) - relativedelta(months=1)).month,
                year=dt.today().year)
    fim = dt(day=20, month=comp, year=dt.today().year)
    dic = {}
    for d in desl:
        if inicio <= dt.strptime(d.datadesligamento, '%d/%m/%Y') <= fim:
            if d.professor in dic:
                d2 = {d.professor: {d.departamento: d.datadesligamento}}
                dic[d.professor] = {**dic[d.professor], **d2[d.professor]}
            else:
                d2 = {d.professor: {d.departamento: d.datadesligamento}}
                dic = {**dic, **d2}
    session.close()
    return dic


def consultar_escalas(comp: int) -> dict:
    sessions = sessionmaker(bind=enginefolha)
    session = sessions()
    esc = session.query(Escala).all()
    inicio = dt(day=21, month=(dt(day=1, month=comp, year=dt.today().year) - relativedelta(months=1)).month,
                year=dt.today().year)
    fim = dt(day=20, month=comp, year=dt.today().year)
    dic = {}
    for e in esc:
        if inicio <= dt.strptime(e.data, '%d/%m/%Y') <= fim:
            if e.professor in dic:
                d2 = {e.professor: {e.data: {e.departamento: e.horas}}}
                dic[e.professor] = {**dic[e.professor], **d2[e.professor]}
            else:
                d2 = {e.professor: {e.data: {e.departamento: e.horas}}}
                dic = {**dic, **d2}
    session.close()
    return dic


def consultar_horas_complementares(comp: int) -> dict:
    sessions = sessionmaker(bind=enginefolha)
    session = sessions()
    hrsc = session.query(Hrcomplement).all()
    inicio = dt(day=21, month=(dt(day=1, month=comp, year=dt.today().year) - relativedelta(months=1)).month,
                year=dt.today().year)
    fim = dt(day=20, month=comp, year=dt.today().year)
    dic = {}
    for h in hrsc:
        if inicio <= dt.strptime(h.data, '%d/%m/%Y') <= fim:
            if h.professor in dic:
                d2 = {h.professor: {h.data: {h.departamento: h.horas}}}
                dic[h.professor] = {**dic[h.professor], **d2[h.professor]}
            else:
                d2 = {h.professor: {h.data: {h.departamento: h.horas}}}
                dic = {**dic, **d2}
    session.close()
    return dic


def salvar_planilha_grade_horaria(dic: dict, comp: int):
    hj = dt.today()
    mes = str(comp).zfill(2)
    ano = hj.year
    mesext = {'01': 'JAN', '02': 'FEV', '03': 'MAR', '04': 'ABR', '05': 'MAI', '06': 'JUN',
              '07': 'JUL', '08': 'AGO', '09': 'SET', '10': 'OUT', '11': 'NOV', '12': 'DEZ'}
    pasta_pgto = rf'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\04 - Folha de Pgto\{ano}\{mes} - {mesext[mes]}\Grades e Comissões'

    grd = os.path.relpath(rf'C:\Users\{os.getlogin()}\PycharmProjects\AutomacaoCia\src\models\static\files\Grade.xlsx')
    grade = l_w(grd, read_only=False)
    plan1 = grade['Planilha1']
    flt = consultar_faltas(comp)
    subs = consultar_substituicoes(comp)
    dslg = consultar_desligamentos(comp)
    fer = consultar_ferias(comp)
    complem = consultar_horas_complementares(comp)
    atest = consultar_atestados(comp)
    feriad = listar_feriados(comp)
    escal = consultar_escalas(comp)
    competencia = dt(day=10, month=comp, year=dt.today().year)
    inicio = dt(day=21, month=(competencia - relativedelta(months=1)).month,
                year=(competencia - relativedelta(months=1)).year)
    fechamento = dt(day=20, month=competencia.month, year=competencia.year)
    plan1['A1'].value = 'Folha'
    plan1['B1'].value = f'{fechamento.month} de {fechamento.year}'

    def intervalo(inicio, fechamento):
        for n in range(int((fechamento - inicio).days) + 1):
            yield dt.strftime(inicio + td(n), '%d/%m/%Y')
    # inserir data no cabeçalho da grade
    col = 3
    for item in list(intervalo(inicio, fechamento)):
        plan1.cell(column=col, row=3, value=dt.strftime(dt.strptime(item, '%d/%m/%Y'), '%a'))
        plan1.cell(column=col, row=4, value=dt.strftime(dt.strptime(item, '%d/%m/%Y'), '%d/%m'))
        col += 1
    # indesir total na coluna ao final das datas
    plan1.cell(column=col, row=3, value='Total')
    # formatar coloração de fds
    for itens in plan1.iter_cols(min_row=3, min_col=3, max_row=3, max_col=35):
        for cell in itens:
            if cell.value != 'Total':
                if cell.value == 'sáb' or cell.value == 'dom':
                    letras = openpyxl.utils.cell.get_column_letter(cell.column)
                    for numero in range(3, 150):
                        plan1[f'{letras}{numero}'].fill = fds
    # separar cada prof de cada depto
    musculacao = []
    ginastica = []
    esportes = []
    kids = []
    cross = []
    for i in dic:
        for sub in dic[i]:
            if dic[i][sub] == {}:
                pass
            else:
                if sub == 'Musculação':
                    musculacao.append(i)
                    musculacao.sort()
                if sub == 'Ginástica':
                    ginastica.append(i)
                    ginastica.sort()
                if sub == 'Esportes':
                    esportes.append(i)
                    esportes.sort()
                if sub == 'Kids':
                    kids.append(i)
                    kids.sort()
                if sub == 'Cross Cia':
                    cross.append(i)
                    cross.sort()

    plan1['A5'].value = 'Musculação'
    novalinha = 6
    for i in musculacao:
        for row in plan1.iter_cols(min_row=3, min_col=3, max_row=3, max_col=35):
            for cell in row:
                if cell.value != 'Total':
                    if cell.value == 'seg':
                        plan1.cell(column=cell.column, row=novalinha, value=somar_aulas_de_segunda(i, 'Musculação'))
                    if cell.value == 'ter':
                        plan1.cell(column=cell.column, row=novalinha, value=somar_aulas_de_terca(i, 'Musculação'))
                    if cell.value == 'qua':
                        plan1.cell(column=cell.column, row=novalinha, value=somar_aulas_de_quarta(i, 'Musculação'))
                    if cell.value == 'qui':
                        plan1.cell(column=cell.column, row=novalinha, value=somar_aulas_de_quinta(i, 'Musculação'))
                    if cell.value == 'sex':
                        plan1.cell(column=cell.column, row=novalinha, value=somar_aulas_de_sexta(i, 'Musculação'))
                    if cell.value == 'sáb':
                        plan1.cell(column=cell.column, row=novalinha, value=somar_aulas_de_sabado(i, 'Musculação'))
                    if cell.value == 'dom':
                        plan1.cell(column=cell.column, row=novalinha, value=somar_aulas_de_domingo(i, 'Musculação'))
                else:
                    letra = openpyxl.utils.cell.get_column_letter(cell.column - 1)
                    plan1.cell(column=cell.column, row=novalinha, value=f'=SUM(C{novalinha}:{letra}{novalinha})')
                # aplica cor de falta
                for nome in flt:
                    for dia in flt[nome]:
                        for depart in flt[nome][dia]:
                            if depart == 'Musculação' and nome == i:
                                if plan1.cell(column=cell.column, row=cell.row + 1).value == dt.strftime(
                                        dt.strptime(dia, '%d/%m/%Y'), '%d/%m'):
                                    plan1.cell(column=cell.column, row=novalinha).fill = falta
                # aplica alterações de substituição
                # {s.professorsubst: {s.substituto: {s.departamento: {s.data: s.horas}}}}
                # {s.professorsubst: {s.data: {s.substituto: {s.departamento: s.horas}}}}
                for nome in subs:
                    for substituto in subs[nome]:
                        for depart in subs[nome][substituto]:
                            for dia in subs[nome][substituto][depart]:
                                if depart == 'Musculação' and nome == i:
                                    if plan1.cell(column=cell.column, row=cell.row + 1).value == dt.strftime(
                                            dt.strptime(dia, '%d/%m/%Y'), '%d/%m'):
                                        plan1.cell(column=cell.column, row=novalinha).fill = falta
                                if depart == 'Musculação' and substituto == i:
                                    if plan1.cell(column=cell.column, row=cell.row + 1).value == dt.strftime(
                                            dt.strptime(dia, '%d/%m/%Y'), '%d/%m'):
                                        plan1.cell(column=cell.column, row=novalinha).value = plan1.cell(
                                            column=cell.column, row=novalinha).value + float(
                                            str(subs[nome][substituto][depart][dia]).replace(',', '.'))
                                        plan1.cell(column=cell.column, row=novalinha).fill = subst

                # aplica alterações de desligamento
                # {d.professor: {d.departamento: d.datadesligamento}}
                # conferir se tem outras aulas ativas ou foi desligado de tudo
                # se desligado de tudo, alterar status das aulas para inativas
                for nome in dslg:
                    for depart in dslg[nome]:
                        if depart == 'Musculação' and nome == i and dt.strptime(dia, '%d/%m/%Y') <= fechamento:
                            if plan1.cell(column=cell.column, row=cell.row + 1).value is not None:
                                if dt.strptime(dia, '%d/%m/%Y') <= dt(day=int(
                                        str(plan1.cell(column=cell.column, row=cell.row + 1).value).split('/')[0]),
                                                                      month=int(str(plan1.cell(column=cell.column,
                                                                                               row=cell.row + 1).value).split(
                                                                              '/')[1]),
                                                                      year=dt.today().year) <= fechamento:
                                    plan1.cell(column=cell.column, row=novalinha).value = 0
                                    plan1.cell(column=cell.column, row=novalinha).fill = deslig
                                    plan1.cell(column=cell.column, row=novalinha).font = Font(color='FFFFFF')

                # aplica talterações de férias
                # # {f.professor: {f.departamento: {f.inicio: f.fim}}}
                for nome in fer:
                    for depart in fer[nome]:
                        for inic in fer[nome][depart]:
                            if depart == 'Musculação' and nome == i and dt.strptime(inic, '%d/%m/%Y') <= fechamento:
                                if plan1.cell(column=cell.column, row=cell.row + 1).value is not None:
                                    if dt.strptime(inic, '%d/%m/%Y') <= dt(day=int(
                                            str(plan1.cell(column=cell.column, row=cell.row + 1).value).split('/')[0]),
                                                                           month=int(str(plan1.cell(column=cell.column,
                                                                                                    row=cell.row + 1).value).split(
                                                                                   '/')[1]),
                                                                           year=dt.today().year) <= dt.strptime(
                                            fer[nome][depart][inic], '%d/%m/%Y'):
                                        plan1.cell(column=cell.column, row=novalinha).fill = ferias
                                        plan1.cell(column=cell.column, row=novalinha).value = 0

                # aplica alterações de horas complementares
                # {h.professor: {h.data: {h.departamento: h.horas}}}
                for nome in complem:
                    for dia in complem[nome]:
                        for depart in complem[nome][dia]:
                            if depart == 'Musculação' and nome == i:
                                if plan1.cell(column=cell.column, row=cell.row + 1).value == dt.strftime(
                                        dt.strptime(dia, '%d/%m/%Y'), '%d/%m'):
                                    plan1.cell(column=cell.column, row=novalinha).value = plan1.cell(column=cell.column,
                                                                                                     row=novalinha).value + float(
                                        str(complem[nome][dia][depart]).replace(',', '.'))
                                    plan1.cell(column=cell.column, row=novalinha).fill = comple

                # aplica alterações de atestados
                # {a.professor: {a.data: a.departamento}}
                for nome in atest:
                    for d in atest[nome]:
                        if atest[nome][d] == 'Musculação' and nome == i:
                            if plan1.cell(column=cell.column, row=cell.row + 1).value == dt.strftime(
                                    dt.strptime(d, '%d/%m/%Y'), '%d/%m'):
                                plan1.cell(column=cell.column, row=novalinha).fill = atestado

                # aplica alterações de feriado
                # [datas de feriado formato dt]
                for dia in feriad:
                    if plan1.cell(column=cell.column, row=cell.row + 1).value == dt.strftime(dia, '%d/%m'):
                        plan1.cell(column=cell.column, row=novalinha).fill = feriado
                        plan1.cell(column=cell.column, row=novalinha).value = 0

                # aplica alterações de escala
                # {e.professor: {e.data: {e.departamento: e.horas}}}
                for nome in escal:
                    for dia in escal[nome]:
                        for depart in escal[nome][dia]:
                            if depart == 'Musculação' and nome == i:
                                if plan1.cell(column=cell.column, row=cell.row + 1).value == dt.strftime(
                                        dt.strptime(dia, '%d/%m/%Y'), '%d/%m'):
                                    plan1.cell(column=cell.column, row=novalinha).value = plan1.cell(column=cell.column,
                                                                                                     row=novalinha).value + float(
                                        str(escal[nome][dia][depart]).replace(',', '.'))
        plan1.cell(column=2, row=novalinha, value=i)
        novalinha += 1

    plan1[f'A{novalinha}'].value = 'Ginástica'
    novalinha += 1
    for i in ginastica:
        for row in plan1.iter_cols(min_row=3, min_col=3, max_row=3, max_col=35):
            for cell in row:
                if cell.value != 'Total':
                    if cell.value == 'seg':
                        plan1.cell(column=cell.column, row=novalinha, value=somar_aulas_de_segunda(i, 'Ginástica'))
                    if cell.value == 'ter':
                        plan1.cell(column=cell.column, row=novalinha, value=somar_aulas_de_terca(i, 'Ginástica'))
                    if cell.value == 'qua':
                        plan1.cell(column=cell.column, row=novalinha, value=somar_aulas_de_quarta(i, 'Ginástica'))
                    if cell.value == 'qui':
                        plan1.cell(column=cell.column, row=novalinha, value=somar_aulas_de_quinta(i, 'Ginástica'))
                    if cell.value == 'sex':
                        plan1.cell(column=cell.column, row=novalinha, value=somar_aulas_de_sexta(i, 'Ginástica'))
                    if cell.value == 'sáb':
                        plan1.cell(column=cell.column, row=novalinha, value=somar_aulas_de_sabado(i, 'Ginástica'))
                    if cell.value == 'dom':
                        plan1.cell(column=cell.column, row=novalinha, value=somar_aulas_de_domingo(i, 'Ginástica'))
                else:
                    letra = openpyxl.utils.cell.get_column_letter(cell.column - 1)
                    plan1.cell(column=cell.column, row=novalinha, value=f'=SUM(C{novalinha}:{letra}{novalinha})')
                # aplica cor de falta
                for nome in flt:
                    for dia in flt[nome]:
                        for depart in flt[nome][dia]:
                            if depart == 'Ginástica' and nome == i:
                                if plan1.cell(column=cell.column, row=cell.row + 1).value == dt.strftime(
                                        dt.strptime(dia, '%d/%m/%Y'), '%d/%m'):
                                    plan1.cell(column=cell.column, row=novalinha).fill = falta
                # aplica alterações de substituição
                # {s.professorsubst: {s.substituto: {s.departamento: {s.data: s.horas}}}}
                for nome in subs:
                    for substituto in subs[nome]:
                        for depart in subs[nome][substituto]:
                            for dia in subs[nome][substituto][depart]:
                                if depart == 'Ginástica' and nome == i:
                                    if plan1.cell(column=cell.column, row=cell.row + 1).value == dt.strftime(
                                            dt.strptime(dia, '%d/%m/%Y'), '%d/%m'):
                                        plan1.cell(column=cell.column, row=novalinha).fill = falta
                                if depart == 'Ginástica' and substituto == i:
                                    if plan1.cell(column=cell.column, row=cell.row + 1).value == dt.strftime(
                                            dt.strptime(dia, '%d/%m/%Y'), '%d/%m'):
                                        plan1.cell(column=cell.column, row=novalinha).value = plan1.cell(
                                            column=cell.column, row=novalinha).value + float(
                                            str(subs[nome][substituto][depart][dia]).replace(',', '.'))
                                        plan1.cell(column=cell.column, row=novalinha).fill = subst

                # aplica alterações de desligamento
                # {d.professor: {d.departamento: d.datadesligamento}}
                # conferir se tem outras aulas ativas ou foi desligado de tudo
                # se desligado de tudo, alterar status das aulas para inativas
                for nome in dslg:
                    for depart in dslg[nome]:
                        if depart == 'Ginástica' and nome == i and dt.strptime(dia, '%d/%m/%Y') <= fechamento:
                            if plan1.cell(column=cell.column, row=cell.row + 1).value is not None:
                                if dt.strptime(dia, '%d/%m/%Y') <= dt(day=int(
                                        str(plan1.cell(column=cell.column, row=cell.row + 1).value).split('/')[0]),
                                                                      month=int(str(plan1.cell(column=cell.column,
                                                                                               row=cell.row + 1).value).split(
                                                                              '/')[1]),
                                                                      year=dt.today().year) <= fechamento:
                                    plan1.cell(column=cell.column, row=novalinha).value = 0
                                    plan1.cell(column=cell.column, row=novalinha).fill = deslig
                                    plan1.cell(column=cell.column, row=novalinha).font = Font(color='FFFFFF')

                # aplica talterações de férias
                # # {f.professor: {f.departamento: {f.inicio: f.fim}}}
                for nome in fer:
                    for depart in fer[nome]:
                        for inic in fer[nome][depart]:
                            if depart == 'Ginástica' and nome == i and dt.strptime(inic, '%d/%m/%Y') <= fechamento:
                                if plan1.cell(column=cell.column, row=cell.row + 1).value is not None:
                                    if dt.strptime(inic, '%d/%m/%Y') <= dt(day=int(
                                            str(plan1.cell(column=cell.column, row=cell.row + 1).value).split('/')[0]),
                                                                           month=int(str(plan1.cell(column=cell.column,
                                                                                                    row=cell.row + 1).value).split(
                                                                                   '/')[1]),
                                                                           year=dt.today().year) <= dt.strptime(
                                            fer[nome][depart][inic], '%d/%m/%Y'):
                                        plan1.cell(column=cell.column, row=novalinha).fill = ferias
                                        plan1.cell(column=cell.column, row=novalinha).value = 0

                # aplica alterações de horas complementares
                # {h.professor: {h.data: {h.departamento: h.horas}}}
                for nome in complem:
                    for dia in complem[nome]:
                        for depart in complem[nome][dia]:
                            if depart == 'Ginástica' and nome == i:
                                if plan1.cell(column=cell.column, row=cell.row + 1).value == dt.strftime(
                                        dt.strptime(dia, '%d/%m/%Y'), '%d/%m'):
                                    plan1.cell(column=cell.column, row=novalinha).value = plan1.cell(column=cell.column,
                                                                                                     row=novalinha).value + float(
                                        str(complem[nome][dia][depart]).replace(',', '.'))
                                    plan1.cell(column=cell.column, row=novalinha).fill = comple

                # aplica alterações de atestados
                # {a.professor: {a.data: a.departamento}}
                for nome in atest:
                    for d in atest[nome]:
                        if atest[nome][d] == 'Ginástica' and nome == i:
                            if plan1.cell(column=cell.column, row=cell.row + 1).value == dt.strftime(
                                    dt.strptime(d, '%d/%m/%Y'), '%d/%m'):
                                plan1.cell(column=cell.column, row=novalinha).fill = atestado

                # aplica alterações de feriado
                # [datas de feriado formato dt]
                for dia in feriad:
                    if plan1.cell(column=cell.column, row=cell.row + 1).value == dt.strftime(dia, '%d/%m'):
                        plan1.cell(column=cell.column, row=novalinha).fill = feriado
                        plan1.cell(column=cell.column, row=novalinha).value = 0

                # aplica alterações de escala
                # {e.professor: {e.data: {e.departamento: e.horas}}}
                for nome in escal:
                    for dia in escal[nome]:
                        for depart in escal[nome][dia]:
                            if depart == 'Ginástica' and nome == i:
                                if plan1.cell(column=cell.column, row=cell.row + 1).value == dt.strftime(
                                        dt.strptime(dia, '%d/%m/%Y'), '%d/%m'):
                                    plan1.cell(column=cell.column, row=novalinha).value = plan1.cell(column=cell.column,
                                                                                                     row=novalinha).value + float(
                                        str(escal[nome][dia][depart]).replace(',', '.'))

        plan1.cell(column=2, row=novalinha, value=i)
        novalinha += 1

    plan1[f'A{novalinha}'].value = 'Kids'
    novalinha += 1
    for i in kids:
        for row in plan1.iter_cols(min_row=3, min_col=3, max_row=3, max_col=35):
            for cell in row:
                if cell.value != 'Total':
                    if cell.value == 'seg':
                        plan1.cell(column=cell.column, row=novalinha, value=somar_aulas_de_segunda(i, 'Kids'))
                    if cell.value == 'ter':
                        plan1.cell(column=cell.column, row=novalinha, value=somar_aulas_de_terca(i, 'Kids'))
                    if cell.value == 'qua':
                        plan1.cell(column=cell.column, row=novalinha, value=somar_aulas_de_quarta(i, 'Kids'))
                    if cell.value == 'qui':
                        plan1.cell(column=cell.column, row=novalinha, value=somar_aulas_de_quinta(i, 'Kids'))
                    if cell.value == 'sex':
                        plan1.cell(column=cell.column, row=novalinha, value=somar_aulas_de_sexta(i, 'Kids'))
                    if cell.value == 'sáb':
                        plan1.cell(column=cell.column, row=novalinha, value=somar_aulas_de_sabado(i, 'Kids'))
                    if cell.value == 'dom':
                        plan1.cell(column=cell.column, row=novalinha, value=somar_aulas_de_domingo(i, 'Kids'))
                else:
                    letra = openpyxl.utils.cell.get_column_letter(cell.column - 1)
                    plan1.cell(column=cell.column, row=novalinha, value=f'=SUM(C{novalinha}:{letra}{novalinha})')
                # aplica cor de falta
                for nome in flt:
                    for dia in flt[nome]:
                        for depart in flt[nome][dia]:
                            if depart == 'Kids' and nome == i:
                                if plan1.cell(column=cell.column, row=cell.row + 1).value == dt.strftime(
                                        dt.strptime(dia, '%d/%m/%Y'), '%d/%m'):
                                    plan1.cell(column=cell.column, row=novalinha).fill = falta
                # aplica alterações de substituição
                # {s.professorsubst: {s.substituto: {s.departamento: {s.data: s.horas}}}}
                for nome in subs:
                    for substituto in subs[nome]:
                        for depart in subs[nome][substituto]:
                            for dia in subs[nome][substituto][depart]:
                                if depart == 'Kids' and nome == i:
                                    if plan1.cell(column=cell.column, row=cell.row + 1).value == dt.strftime(
                                            dt.strptime(dia, '%d/%m/%Y'), '%d/%m'):
                                        plan1.cell(column=cell.column, row=novalinha).fill = falta
                                if depart == 'Kids' and substituto == i:
                                    if plan1.cell(column=cell.column, row=cell.row + 1).value == dt.strftime(
                                            dt.strptime(dia, '%d/%m/%Y'), '%d/%m'):
                                        plan1.cell(column=cell.column, row=novalinha).value = plan1.cell(
                                            column=cell.column, row=novalinha).value + float(
                                            str(subs[nome][substituto][depart][dia]).replace(',', '.'))
                                        plan1.cell(column=cell.column, row=novalinha).fill = subst

                # aplica alterações de desligamento
                # {d.professor: {d.departamento: d.datadesligamento}}
                # conferir se tem outras aulas ativas ou foi desligado de tudo
                # se desligado de tudo, alterar status das aulas para inativas
                for nome in dslg:
                    for depart in dslg[nome]:
                        if depart == 'Kids' and nome == i and dt.strptime(dslg[nome][depart], '%d/%m/%Y') <= fechamento:
                            if plan1.cell(column=cell.column, row=cell.row + 1).value is not None:
                                if dt.strptime(dslg[nome][depart], '%d/%m/%Y') <= dt(day=int(
                                        str(plan1.cell(column=cell.column, row=cell.row + 1).value).split('/')[0]),
                                                                                     month=int(str(plan1.cell(
                                                                                             column=cell.column,
                                                                                             row=cell.row + 1).value).split(
                                                                                             '/')[1]),
                                                                                     year=dt.today().year) <= fechamento:
                                    plan1.cell(column=cell.column, row=novalinha).value = 0
                                    plan1.cell(column=cell.column, row=novalinha).fill = deslig
                                    plan1.cell(column=cell.column, row=novalinha).font = Font(color='FFFFFF')

                # aplica talterações de férias
                # # {f.professor: {f.departamento: {f.inicio: f.fim}}}
                for nome in fer:
                    for depart in fer[nome]:
                        for inic in fer[nome][depart]:
                            if depart == 'Kids' and nome == i and dt.strptime(inic, '%d/%m/%Y') <= fechamento:
                                if plan1.cell(column=cell.column, row=cell.row + 1).value is not None:
                                    if dt.strptime(inic, '%d/%m/%Y') <= dt(day=int(
                                            str(plan1.cell(column=cell.column, row=cell.row + 1).value).split('/')[0]),
                                                                           month=int(str(plan1.cell(column=cell.column,
                                                                                                    row=cell.row + 1).value).split(
                                                                                   '/')[1]),
                                                                           year=dt.today().year) <= dt.strptime(
                                            fer[nome][depart][inic], '%d/%m/%Y'):
                                        plan1.cell(column=cell.column, row=novalinha).fill = ferias
                                        plan1.cell(column=cell.column, row=novalinha).value = 0

                # aplica alterações de horas complementares
                # {h.professor: {h.data: {h.departamento: h.horas}}}
                for nome in complem:
                    for dia in complem[nome]:
                        for depart in complem[nome][dia]:
                            if depart == 'Kids' and nome == i:
                                if plan1.cell(column=cell.column, row=cell.row + 1).value == dt.strftime(
                                        dt.strptime(dia, '%d/%m/%Y'), '%d/%m'):
                                    plan1.cell(column=cell.column, row=novalinha).value = plan1.cell(column=cell.column,
                                                                                                     row=novalinha).value + float(
                                        str(complem[nome][dia][depart]).replace(',', '.'))
                                    plan1.cell(column=cell.column, row=novalinha).fill = comple

                # aplica alterações de atestados
                # {a.professor: {a.data: a.departamento}}
                for nome in atest:
                    for d in atest[nome]:
                        if atest[nome][d] == 'Kids' and nome == i:
                            if plan1.cell(column=cell.column, row=cell.row + 1).value == dt.strftime(
                                    dt.strptime(d, '%d/%m/%Y'), '%d/%m'):
                                plan1.cell(column=cell.column, row=novalinha).fill = atestado

                # aplica alterações de feriado
                # [datas de feriado formato dt]
                for dia in feriad:
                    if plan1.cell(column=cell.column, row=cell.row + 1).value == dt.strftime(dia, '%d/%m'):
                        plan1.cell(column=cell.column, row=novalinha).fill = feriado
                        plan1.cell(column=cell.column, row=novalinha).value = 0

                # aplica alterações de escala
                # {e.professor: {e.data: {e.departamento: e.horas}}}
                for nome in escal:
                    for dia in escal[nome]:
                        for depart in escal[nome][dia]:
                            if depart == 'Kids' and nome == i:
                                if plan1.cell(column=cell.column, row=cell.row + 1).value == dt.strftime(
                                        dt.strptime(dia, '%d/%m/%Y'), '%d/%m'):
                                    plan1.cell(column=cell.column, row=novalinha).value = plan1.cell(column=cell.column,
                                                                                                     row=novalinha).value + float(
                                        str(escal[nome][dia][depart]).replace(',', '.'))
        plan1.cell(column=2, row=novalinha, value=i)
        novalinha += 1

    plan1[f'A{novalinha}'].value = 'Esportes'
    novalinha += 1
    for i in esportes:
        for row in plan1.iter_cols(min_row=3, min_col=3, max_row=3, max_col=35):
            for cell in row:
                if cell.value != 'Total':
                    if cell.value == 'seg':
                        plan1.cell(column=cell.column, row=novalinha, value=somar_aulas_de_segunda(i, 'Esportes'))
                    if cell.value == 'ter':
                        plan1.cell(column=cell.column, row=novalinha, value=somar_aulas_de_terca(i, 'Esportes'))
                    if cell.value == 'qua':
                        plan1.cell(column=cell.column, row=novalinha, value=somar_aulas_de_quarta(i, 'Esportes'))
                    if cell.value == 'qui':
                        plan1.cell(column=cell.column, row=novalinha, value=somar_aulas_de_quinta(i, 'Esportes'))
                    if cell.value == 'sex':
                        plan1.cell(column=cell.column, row=novalinha, value=somar_aulas_de_sexta(i, 'Esportes'))
                    if cell.value == 'sáb':
                        plan1.cell(column=cell.column, row=novalinha, value=somar_aulas_de_sabado(i, 'Esportes'))
                    if cell.value == 'dom':
                        plan1.cell(column=cell.column, row=novalinha, value=somar_aulas_de_domingo(i, 'Esportes'))
                else:
                    letra = openpyxl.utils.cell.get_column_letter(cell.column - 1)
                    plan1.cell(column=cell.column, row=novalinha, value=f'=SUM(C{novalinha}:{letra}{novalinha})')
                # aplica cor de falta
                for nome in flt:
                    for dia in flt[nome]:
                        for depart in flt[nome][dia]:
                            if depart == 'Esportes' and nome == i:
                                if plan1.cell(column=cell.column, row=cell.row + 1).value == dt.strftime(
                                        dt.strptime(dia, '%d/%m/%Y'), '%d/%m'):
                                    plan1.cell(column=cell.column, row=novalinha).fill = falta
                # aplica alterações de substituição
                # {s.professorsubst: {s.substituto: {s.departamento: {s.data: s.horas}}}}
                for nome in subs:
                    for substituto in subs[nome]:
                        for depart in subs[nome][substituto]:
                            for dia in subs[nome][substituto][depart]:
                                if depart == 'Esportes' and nome == i:
                                    if plan1.cell(column=cell.column, row=cell.row + 1).value == dt.strftime(
                                            dt.strptime(dia, '%d/%m/%Y'), '%d/%m'):
                                        plan1.cell(column=cell.column, row=novalinha).fill = falta
                                if depart == 'Esportes' and substituto == i:
                                    if plan1.cell(column=cell.column, row=cell.row + 1).value == dt.strftime(
                                            dt.strptime(dia, '%d/%m/%Y'), '%d/%m'):
                                        plan1.cell(column=cell.column, row=novalinha).value = plan1.cell(
                                            column=cell.column, row=novalinha).value + float(
                                            str(subs[nome][substituto][depart][dia]).replace(',', '.'))
                                        plan1.cell(column=cell.column, row=novalinha).fill = subst

                # aplica alterações de desligamento
                # {d.professor: {d.departamento: d.datadesligamento}}
                # conferir se tem outras aulas ativas ou foi desligado de tudo
                # se desligado de tudo, alterar status das aulas para inativas
                for nome in dslg:
                    for depart in dslg[nome]:
                        if depart == 'Esportes' and nome == i and dt.strptime(dia, '%d/%m/%Y') <= fechamento:
                            if plan1.cell(column=cell.column, row=cell.row + 1).value is not None:
                                if dt.strptime(dia, '%d/%m/%Y') <= dt(day=int(
                                        str(plan1.cell(column=cell.column, row=cell.row + 1).value).split('/')[0]),
                                                                      month=int(str(plan1.cell(column=cell.column,
                                                                                               row=cell.row + 1).value).split(
                                                                              '/')[1]),
                                                                      year=dt.today().year) <= fechamento:
                                    plan1.cell(column=cell.column, row=novalinha).value = 0
                                    plan1.cell(column=cell.column, row=novalinha).fill = deslig
                                    plan1.cell(column=cell.column, row=novalinha).font = Font(color='FFFFFF')

                # aplica talterações de férias
                # {f.professor: {f.departamento: {f.inicio: f.fim}}}
                for nome in fer:
                    for depart in fer[nome]:
                        for inic in fer[nome][depart]:
                            if depart == 'Esportes' and nome == i and dt.strptime(inic, '%d/%m/%Y') <= fechamento:
                                if plan1.cell(column=cell.column, row=cell.row + 1).value is not None:
                                    if dt.strptime(inic, '%d/%m/%Y') <= dt(day=int(
                                            str(plan1.cell(column=cell.column, row=cell.row + 1).value).split('/')[0]),
                                                                           month=int(str(plan1.cell(column=cell.column,
                                                                                                    row=cell.row + 1).value).split(
                                                                                   '/')[1]),
                                                                           year=dt.today().year) <= dt.strptime(
                                            fer[nome][depart][inic], '%d/%m/%Y'):
                                        plan1.cell(column=cell.column, row=novalinha).fill = ferias
                                        plan1.cell(column=cell.column, row=novalinha).value = 0

                # aplica alterações de horas complementares
                # {h.professor: {h.data: {h.departamento: h.horas}}}
                for nome in complem:
                    for dia in complem[nome]:
                        for depart in complem[nome][dia]:
                            if depart == 'Esportes' and nome == i:
                                if plan1.cell(column=cell.column, row=cell.row + 1).value == dt.strftime(
                                        dt.strptime(dia, '%d/%m/%Y'), '%d/%m'):
                                    plan1.cell(column=cell.column, row=novalinha).value = plan1.cell(column=cell.column,
                                                                                                     row=novalinha).value + float(
                                        str(complem[nome][dia][depart]).replace(',', '.'))
                                    plan1.cell(column=cell.column, row=novalinha).fill = comple

                # aplica alterações de atestados
                # {a.professor: {a.data: a.departamento}}
                for nome in atest:
                    for d in atest[nome]:
                        if atest[nome][d] == 'Esportes' and nome == i:
                            if plan1.cell(column=cell.column, row=cell.row + 1).value == dt.strftime(
                                    dt.strptime(d, '%d/%m/%Y'), '%d/%m'):
                                plan1.cell(column=cell.column, row=novalinha).fill = atestado

                # aplica alterações de feriado
                # [datas de feriado formato dt]
                for dia in feriad:
                    if plan1.cell(column=cell.column, row=cell.row + 1).value == dt.strftime(dia, '%d/%m'):
                        plan1.cell(column=cell.column, row=novalinha).fill = feriado
                        plan1.cell(column=cell.column, row=novalinha).value = 0

                # aplica alterações de escala
                # {e.professor: {e.data: {e.departamento: e.horas}}}
                for nome in escal:
                    for dia in escal[nome]:
                        for depart in escal[nome][dia]:
                            if depart == 'Esportes' and nome == i:
                                if plan1.cell(column=cell.column, row=cell.row + 1).value == dt.strftime(
                                        dt.strptime(dia, '%d/%m/%Y'), '%d/%m'):
                                    plan1.cell(column=cell.column, row=novalinha).value = plan1.cell(column=cell.column,
                                                                                                     row=novalinha).value + float(
                                        str(escal[nome][dia][depart]).replace(',', '.'))

        plan1.cell(column=2, row=novalinha, value=i)
        novalinha += 1

    plan1[f'A{novalinha}'].value = 'Cross Cia'
    novalinha += 1
    for i in cross:
        for row in plan1.iter_cols(min_row=3, min_col=3, max_row=3, max_col=35):
            for cell in row:
                if cell.value != 'Total':
                    if cell.value == 'seg':
                        plan1.cell(column=cell.column, row=novalinha, value=somar_aulas_de_segunda(i, 'Cross Cia'))
                    if cell.value == 'ter':
                        plan1.cell(column=cell.column, row=novalinha, value=somar_aulas_de_terca(i, 'Cross Cia'))
                    if cell.value == 'qua':
                        plan1.cell(column=cell.column, row=novalinha, value=somar_aulas_de_quarta(i, 'Cross Cia'))
                    if cell.value == 'qui':
                        plan1.cell(column=cell.column, row=novalinha, value=somar_aulas_de_quinta(i, 'Cross Cia'))
                    if cell.value == 'sex':
                        plan1.cell(column=cell.column, row=novalinha, value=somar_aulas_de_sexta(i, 'Cross Cia'))
                    if cell.value == 'sáb':
                        plan1.cell(column=cell.column, row=novalinha, value=somar_aulas_de_sabado(i, 'Cross Cia'))
                    if cell.value == 'dom':
                        plan1.cell(column=cell.column, row=novalinha, value=somar_aulas_de_domingo(i, 'Cross Cia'))
                else:
                    letra = openpyxl.utils.cell.get_column_letter(cell.column - 1)
                    plan1.cell(column=cell.column, row=novalinha, value=f'=SUM(C{novalinha}:{letra}{novalinha})')
                # aplica cor de falta
                for nome in flt:
                    for dia in flt[nome]:
                        for depart in flt[nome][dia]:
                            if depart == 'Cross Cia' and nome == i:
                                if plan1.cell(column=cell.column, row=cell.row + 1).value == dt.strftime(
                                        dt.strptime(dia, '%d/%m/%Y'), '%d/%m'):
                                    plan1.cell(column=cell.column, row=novalinha).fill = falta
                # aplica alterações de substituição
                # {s.professorsubst: {s.substituto: {s.departamento: {s.data: s.horas}}}}
                for nome in subs:
                    for substituto in subs[nome]:
                        for depart in subs[nome][substituto]:
                            for dia in subs[nome][substituto][depart]:
                                if depart == 'Cross Cia' and nome == i:
                                    if plan1.cell(column=cell.column, row=cell.row + 1).value == dt.strftime(
                                            dt.strptime(dia, '%d/%m/%Y'), '%d/%m'):
                                        plan1.cell(column=cell.column, row=novalinha).fill = falta
                                if depart == 'Cross Cia' and substituto == i:
                                    if plan1.cell(column=cell.column, row=cell.row + 1).value == dt.strftime(
                                            dt.strptime(dia, '%d/%m/%Y'), '%d/%m'):
                                        plan1.cell(column=cell.column, row=novalinha).value = plan1.cell(
                                            column=cell.column, row=novalinha).value + float(
                                            str(subs[nome][substituto][depart][dia]).replace(',', '.'))
                                        plan1.cell(column=cell.column, row=novalinha).fill = subst

                # aplica alterações de desligamento
                # {d.professor: {d.departamento: d.datadesligamento}}
                # conferir se tem outras aulas ativas ou foi desligado de tudo
                # se desligado de tudo, alterar status das aulas para inativas
                for nome in dslg:
                    for depart in dslg[nome]:
                        if depart == 'Cross Cia' and nome == i and dt.strptime(dia, '%d/%m/%Y') <= fechamento:
                            if plan1.cell(column=cell.column, row=cell.row + 1).value is not None:
                                if dt.strptime(dia, '%d/%m/%Y') <= dt(day=int(
                                        str(plan1.cell(column=cell.column, row=cell.row + 1).value).split('/')[0]),
                                                                      month=int(str(plan1.cell(column=cell.column,
                                                                                               row=cell.row + 1).value).split(
                                                                              '/')[1]),
                                                                      year=dt.today().year) <= fechamento:
                                    plan1.cell(column=cell.column, row=novalinha).value = 0
                                    plan1.cell(column=cell.column, row=novalinha).fill = deslig
                                    plan1.cell(column=cell.column, row=novalinha).font = Font(color='FFFFFF')

                # aplica talterações de férias
                # # {f.professor: {f.departamento: {f.inicio: f.fim}}}
                for nome in fer:
                    for depart in fer[nome]:
                        for inic in fer[nome][depart]:
                            if depart == 'Cross Cia' and nome == i and dt.strptime(inic, '%d/%m/%Y') <= fechamento:
                                if plan1.cell(column=cell.column, row=cell.row + 1).value is not None:
                                    if dt.strptime(inic, '%d/%m/%Y') <= dt(day=int(
                                            str(plan1.cell(column=cell.column, row=cell.row + 1).value).split('/')[0]),
                                                                           month=int(str(plan1.cell(column=cell.column,
                                                                                                    row=cell.row + 1).value).split(
                                                                                   '/')[1]),
                                                                           year=dt.today().year) <= dt.strptime(
                                            fer[nome][depart][inic], '%d/%m/%Y'):
                                        plan1.cell(column=cell.column, row=novalinha).fill = ferias
                                        plan1.cell(column=cell.column, row=novalinha).value = 0

                # aplica alterações de horas complementares
                # {h.professor: {h.data: {h.departamento: h.horas}}}
                for nome in complem:
                    for dia in complem[nome]:
                        for depart in complem[nome][dia]:
                            if depart == 'Cross Cia' and nome == i:
                                if plan1.cell(column=cell.column, row=cell.row + 1).value == dt.strftime(
                                        dt.strptime(dia, '%d/%m/%Y'), '%d/%m'):
                                    plan1.cell(column=cell.column, row=novalinha).value = plan1.cell(column=cell.column,
                                                                                                     row=novalinha).value + float(
                                        str(complem[nome][dia][depart]).replace(',', '.'))
                                    plan1.cell(column=cell.column, row=novalinha).fill = comple

                # aplica alterações de atestados
                # {a.professor: {a.data: a.departamento}}
                for nome in atest:
                    for d in atest[nome]:
                        if atest[nome][d] == 'Cross Cia' and nome == i:
                            if plan1.cell(column=cell.column, row=cell.row + 1).value == dt.strftime(
                                    dt.strptime(d, '%d/%m/%Y'), '%d/%m'):
                                plan1.cell(column=cell.column, row=novalinha).fill = atestado

                # aplica alterações de feriado
                # [datas de feriado formato dt]
                for dia in feriad:
                    if plan1.cell(column=cell.column, row=cell.row + 1).value == dt.strftime(dia, '%d/%m'):
                        plan1.cell(column=cell.column, row=novalinha).fill = feriado
                        plan1.cell(column=cell.column, row=novalinha).value = 0

                # aplica alterações de escala
                # {e.professor: {e.data: {e.departamento: e.horas}}}
                for nome in escal:
                    for dia in escal[nome]:
                        for depart in escal[nome][dia]:
                            if depart == 'Cross Cia' and nome == i:
                                if plan1.cell(column=cell.column, row=cell.row + 1).value == dt.strftime(
                                        dt.strptime(dia, '%d/%m/%Y'), '%d/%m'):
                                    plan1.cell(column=cell.column, row=novalinha).value = plan1.cell(column=cell.column,
                                                                                                     row=novalinha).value + float(
                                        str(escal[nome][dia][depart]).replace(',', '.'))

        plan1.cell(column=2, row=novalinha, value=i)
        novalinha += 1

    for i, coluna in enumerate(plan1.columns):
        max_length = 0
        column = coluna[0].column_letter
        for cell in coluna:
            try:
                if len(str(cell.value)) > max_length and len(str(cell.value)) != 18:
                    max_length = len(str(cell.value))
            except Exception:
                pass
        adjusted_width = max_length + 1
        plan1.column_dimensions[column].width = adjusted_width

    for row in plan1.iter_cols(min_row=3, min_col=3, max_row=3, max_col=35):
        for cell in row:
            if plan1.cell(column=cell.column, row=cell.row).value == 'Total':
                plan1.column_dimensions[openpyxl.utils.cell.get_column_letter(cell.column)].width = 8

    plan1['C1'].fill = atestado
    plan1['D1'].value = 'Atestado'
    plan1['C2'].fill = falta
    plan1['D2'].value = 'Falta'
    plan1['F1'].fill = ferias
    plan1['G1'].value = 'Férias'
    plan1['F2'].fill = feriado
    plan1['G2'].value = 'Feriado'
    plan1['I1'].fill = deslig
    plan1['J1'].value = 'Desligamento'
    plan1['I2'].fill = subst
    plan1['J2'].value = 'Substituiu'
    plan1['M1'].fill = comple
    plan1['N1'].value = 'Horas Complementares'
    grade.save(pasta_pgto + f'\\Grade {fechamento.month}-{fechamento.year}.xlsx')


def somar_aulas_de_segunda(nome: str, depto: str) -> float:
    sessions = sessionmaker(bind=enginefolha)
    session = sessions()
    aulasseg = session.query(Aulas).filter_by(professor=nome).filter_by(status='Ativa').filter_by(diadasemana='Segunda') \
        .filter_by(departamento=depto).all()
    somas = 0
    for aula in aulasseg:
        hr, minut, seg = str(dt.strptime(aula.fim, '%H:%M:%S') - dt.strptime(aula.inicio, '%H:%M:%S')).split(':')
        somadia = int(hr) + int(minut) / 60 + int(seg) / 60 * 60
        somadia = round(somadia, 2)
        somas += somadia
    session.close()
    return somas


def somar_aulas_de_terca(nome: str, depto: str) -> float:
    sessions = sessionmaker(bind=enginefolha)
    session = sessions()
    aulaster = session.query(Aulas).filter_by(professor=nome).filter_by(status='Ativa').filter_by(diadasemana='Terça') \
        .filter_by(departamento=depto).all()
    somas = 0
    for aula in aulaster:
        hr, minut, seg = str(dt.strptime(aula.fim, '%H:%M:%S') - dt.strptime(aula.inicio, '%H:%M:%S')).split(':')
        somadia = int(hr) + int(minut) / 60 + int(seg) / 60 * 60
        somadia = round(somadia, 2)
        somas += somadia
    session.close()
    return somas


def somar_aulas_de_quarta(nome: str, depto: str) -> float:
    sessions = sessionmaker(bind=enginefolha)
    session = sessions()
    aulasqua = session.query(Aulas).filter_by(professor=nome).filter_by(status='Ativa').filter_by(diadasemana='Quarta') \
        .filter_by(departamento=depto).all()
    somas = 0
    for aula in aulasqua:
        hr, minut, seg = str(dt.strptime(aula.fim, '%H:%M:%S') - dt.strptime(aula.inicio, '%H:%M:%S')).split(':')
        somadia = int(hr) + int(minut) / 60 + int(seg) / 60 * 60
        somadia = round(somadia, 2)
        somas += somadia
    session.close()
    return somas


def somar_aulas_de_quinta(nome: str, depto: str) -> float:
    sessions = sessionmaker(bind=enginefolha)
    session = sessions()
    aulasqui = session.query(Aulas).filter_by(professor=nome).filter_by(status='Ativa').filter_by(diadasemana='Quinta') \
        .filter_by(departamento=depto).all()
    somas = 0
    for aula in aulasqui:
        hr, minut, seg = str(dt.strptime(aula.fim, '%H:%M:%S') - dt.strptime(aula.inicio, '%H:%M:%S')).split(':')
        somadia = int(hr) + int(minut) / 60 + int(seg) / 60 * 60
        somadia = round(somadia, 2)
        somas += somadia
    session.close()
    return somas


def somar_aulas_de_sexta(nome: str, depto: str) -> float:
    sessions = sessionmaker(bind=enginefolha)
    session = sessions()
    aulassex = session.query(Aulas).filter_by(professor=nome).filter_by(status='Ativa').filter_by(diadasemana='Sexta') \
        .filter_by(departamento=depto).all()
    somas = 0
    for aula in aulassex:
        hr, minut, seg = str(dt.strptime(aula.fim, '%H:%M:%S') - dt.strptime(aula.inicio, '%H:%M:%S')).split(':')
        somadia = int(hr) + int(minut) / 60 + int(seg) / 60 * 60
        somadia = round(somadia, 2)
        somas += somadia
    session.close()
    return somas


def somar_aulas_de_sabado(nome: str, depto: str) -> float:
    sessions = sessionmaker(bind=enginefolha)
    session = sessions()
    aulassab = session.query(Aulas).filter_by(professor=nome).filter_by(status='Ativa').filter_by(diadasemana='Sábado') \
        .filter_by(departamento=depto).all()
    somas = 0
    for aula in aulassab:
        hr, minut, seg = str(dt.strptime(aula.fim, '%H:%M:%S') - dt.strptime(aula.inicio, '%H:%M:%S')).split(':')
        somadia = int(hr) + int(minut) / 60 + int(seg) / 60 * 60
        somadia = round(somadia, 2)
        somas += somadia
    session.close()
    return somas


def somar_aulas_de_domingo(nome: str, depto: str) -> float:
    sessions = sessionmaker(bind=enginefolha)
    session = sessions()
    aulasdom = session.query(Aulas).filter_by(professor=nome).filter_by(status='Ativa').filter_by(diadasemana='Domingo') \
        .filter_by(departamento=depto).all()
    somas = 0
    for aula in aulasdom:
        hr, minut, seg = str(dt.strptime(aula.fim, '%H:%M:%S') - dt.strptime(aula.inicio, '%H:%M:%S')).split(':')
        somadia = int(hr) + int(minut) / 60 + int(seg) / 60 * 60
        somadia = round(somadia, 2)
        somas += somadia
    session.close()
    return somas * 2


def salvar_planilha_soma_final(compet: int):
    hj = dt.today()
    mes = str(compet).zfill(2)
    ano = hj.year
    mesext = {'01': 'JAN', '02': 'FEV', '03': 'MAR', '04': 'ABR', '05': 'MAI', '06': 'JUN',
              '07': 'JUL', '08': 'AGO', '09': 'SET', '10': 'OUT', '11': 'NOV', '12': 'DEZ'}
    pasta_pgto = rf'\\192.168.0.250\rh\01 - RH\01 - Administração.Controles\04 - Folha de Pgto\{ano}\{mes} - {mesext[mes]}\Grades e Comissões'
    try:
        os.makedirs(pasta_pgto)
    except FileExistsError:
        pass
    sessions = sessionmaker(bind=enginefolha)
    session = sessions()
    folhadehoje = Folha(compet, list(listar_aulas_ativas(compet)), listar_departamentos_ativos())
    somaaulas = {}
    for i in listar_professores_ativos():
        somaaulas[i] = {}
        for d in listar_departamentos_ativos():
            somaaulas[i][d] = {}
    for aulas in listar_aulas_ativas(compet):
        somaaulas[aulas.professor][aulas.departamento][aulas.nome + f' ({aulas.valor})'] = round(
            somar_horas_professor(folhadehoje, aulas.professor, aulas.departamento, aulas.nome, compet), 2)
        dictchav = list(somaaulas.keys())
        dictchav.sort()
        somafinal = {i: somaaulas[i] for i in dictchav}
    arquivo = os.path.relpath(rf'C:\Users\{os.getlogin()}\PycharmProjects\AutomacaoCia\src\models\static\files\Somafinal.xlsx')
    plan = l_w(arquivo, read_only=False)
    folha = plan['Planilha1']
    folha['A1'].value = 'Matrícula'
    folha['B1'].value = 'Nome'
    folha['C1'].value = 'Aula e Valor'
    folha['D1'].value = 'Horas'
    x = 2
    for i in somafinal:
        matr = session.query(Aulas).filter_by(professor=str(i)).filter_by(status='Ativa').first()
        folha[f'A{x}'].value = int(matr.matrprof)
        folha[f'B{x}'].value = str(i)
        for sub in somafinal[i]:
            if somafinal[i][sub] == {}:
                pass
            else:
                for sub2 in somafinal[i][sub]:
                    folha[f'C{x}'].value = str(sub2)
                    folha[f'D{x}'].value = float(str(somafinal[i][sub][sub2]))
                    x += 1
    # folha['F1'].value = 'Total Bruto - Professores'
    # folha['G1'].value = locale.currency(totaldafolha(folhadehoje), grouping=True)
    plan.save(pasta_pgto + f'\\Somafinal mes {compet}.xlsx')
    salvar_planilha_grade_horaria(somafinal, compet)
    substitutos = {}
    complementares = {}
    feriasl = {}
    desligadosl = {}
    planilha = l_w(pasta_pgto + f'\\Grade {compet}-2023.xlsx')
    aba = planilha['Planilha1']
    for row in aba.iter_cols(min_row=3, min_col=3, max_row=115, max_col=35):
        for cell in row:
            if cell.fill == ferias:
                for i in range(1, 150):
                    if aba.cell(column=1, row=cell.row - i).value is not None:
                        depart = aba.cell(column=1, row=cell.row - i).value
                        break
                for r in aba.iter_cols(min_row=3, min_col=3, max_row=3, max_col=38):
                    for c in r:
                        if c.value == 'Total':
                            tt = c.column
                hrs = 0
                for m in range(3, tt):
                    hrs += aba.cell(column=m, row=cell.row).value
                feriasl[aba.cell(column=2, row=cell.row).value] = {depart: round(hrs, 2)}
            if cell.fill == deslig:
                for i in range(1, 150):
                    if aba.cell(column=1, row=cell.row - i).value is not None:
                        depart = aba.cell(column=1, row=cell.row - i).value
                        break
                for r in aba.iter_cols(min_row=3, min_col=3, max_row=3, max_col=38):
                    for c in r:
                        if c.value == 'Total':
                            tt = c.column
                hrs = 0
                for m in range(3, tt):
                    hrs += aba.cell(column=m, row=cell.row).value
                desligadosl[aba.cell(column=2, row=cell.row).value] = {depart: round(hrs, 2)}
            if cell.fill == subst:
                for i in range(1, 150):
                    if aba.cell(column=1, row=cell.row - i).value is not None:
                        depart = aba.cell(column=1, row=cell.row - i).value
                        break
                for r in aba.iter_cols(min_row=3, min_col=3, max_row=3, max_col=38):
                    for c in r:
                        if c.value == 'Total':
                            tt = c.column
                hrs = 0
                for m in range(3, tt):
                    hrs += aba.cell(column=m, row=cell.row).value
                substitutos[aba.cell(column=2, row=cell.row).value] = {depart: round(hrs, 2)}
            if cell.fill == comple:
                for i in range(1, 150):
                    if aba.cell(column=1, row=cell.row - i).value is not None:
                        depart = aba.cell(column=1, row=cell.row - i).value
                        break
                for r in aba.iter_cols(min_row=3, min_col=3, max_row=3, max_col=38):
                    for c in r:
                        if c.value == 'Total':
                            tt = c.column
                hrs = 0
                for m in range(3, tt):
                    hrs += aba.cell(column=m, row=cell.row).value
                complementares[aba.cell(column=2, row=cell.row).value] = {depart: round(hrs, 2)}

    planilha2 = l_w(pasta_pgto + f'\\Somafinal mes {compet}.xlsx', read_only=False)
    aba2 = planilha2['Planilha1']
    for pessoa in feriasl:
        for depart in feriasl[pessoa]:
            for cll in aba2['B']:
                if cll.value is not None:
                    if cll.value == pessoa:
                        aba2.cell(column=4, row=cll.row, value=float(feriasl[pessoa][depart]))
    for pessoa in desligadosl:
        for depart in desligadosl[pessoa]:
            for cell in aba2['B']:
                if cell.value is not None:
                    if cell.value == pessoa:
                        aba2.cell(column=4, row=cell.row).value = float(desligadosl[pessoa][depart])
    for pessoa in substitutos:
        for depart in substitutos[pessoa]:
            for cell in aba2['B']:
                if cell.value is not None:
                    if cell.value == pessoa:
                        aba2.cell(column=4, row=cell.row).value = float(substitutos[pessoa][depart])
    for pessoa in complementares:
        for depart in complementares[pessoa]:
            for cell in aba2['B']:
                if cell.value is not None:
                    if cell.value == pessoa:
                        aba2.cell(column=4, row=cell.row).value = float(complementares[pessoa][depart])
    for i, coluna in enumerate(aba2.columns):
        max_length = 0
        column = coluna[0].column_letter
        for cell in coluna:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except Exception:
                pass
        adjusted_width = max_length + 1
        aba2.column_dimensions[column].width = adjusted_width
    planilha2.save(pasta_pgto + f'\\Somafinal mes {compet}.xlsx')
    planilha3 = l_w(pasta_pgto + f'\\Grade {compet}-2023.xlsx', read_only=False)
    aba3 = planilha3['Planilha1']
    for row in aba3.iter_cols(min_row=3, min_col=3, max_row=120, max_col=35):
        for cell in row:
            if cell.value == 0:
                cell.value = ''

    planilha3.save(pasta_pgto + f'\\Grade {compet}-2023.xlsx')
    tkinter.messagebox.showinfo(
        title='Grade ok!',
        message=f'Grade do mês {compet} salva com sucesso!'
    )

    print('Férias \n', feriasl)
    print('Desligados \n', desligadosl)
    print('Substitutos \n', substitutos)
    print('Hrs Complementares \n', complementares)


def lancar_ferias(nome, depto, inicio, fim):
    sessions = sessionmaker(bind=enginefolha)
    session = sessions()
    sessionscol = sessionmaker(bind=engine)
    sessioncol = sessionscol()
    pessoa = sessioncol.query(Colaborador).filter_by(nome=nome).first()
    matricula = pessoa.matricula

    fer = Ferias(professor=nome, matrprof=matricula, departamento=depto, inicio=inicio, fim=fim)
    session.add(fer)
    session.commit()
    session.close()
    tkinter.messagebox.showinfo('Férias ok!', 'Férias lançadas com sucesso!')


def lancar_atestado(nome, depto, data):
    sessions = sessionmaker(bind=enginefolha)
    session = sessions()
    sessionscol = sessionmaker(bind=engine)
    sessioncol = sessionscol()

    professor = sessioncol.query(Colaborador).filter_by(nome=nome).first()
    matricula = professor.matricula

    atest = Atestado(professor=nome, matrprof=matricula, departamento=depto, data=data)
    session.add(atest)
    session.commit()
    session.close()
    tkinter.messagebox.showinfo('Atestado salvo!', 'Atestado salvo com sucesso!')


def lancar_desligamento(nome, depto, data):
    sessions = sessionmaker(bind=enginefolha)
    session = sessions()

    sessionscol = sessionmaker(bind=engine)
    sessioncol = sessionscol()

    professor = sessioncol.query(Colaborador).filter_by(nome=nome).first()
    matricula = professor.matricula

    deslig = Desligados(professor=nome, matrprof=matricula, departamento=depto, datadesligamento=data)
    session.add(deslig)
    session.commit()
    session.close()

    # se desligamento antes do iniício da folha: tornar inativas as aulas da pessoa
    tkinter.messagebox.showinfo('Desligamento ok!', 'Desligamento lançado com sucesso!')


def lancar_substit(substituido, substituto, departamento, aula, data, horas):
    sessionscol = sessionmaker(bind=engine)
    sessioncol = sessionscol()
    profsubstituido = sessioncol.query(Colaborador).filter_by(nome=substituido).first()
    matrsubsido = profsubstituido.matricula
    profsubstituto = sessioncol.query(Colaborador).filter_by(nome=substituto).first()
    matrsubstuto = profsubstituto.matricula
    sessions = sessionmaker(bind=enginefolha)
    session = sessions()
    subs = Substituicao(professorsubst=substituido, matrprof=matrsubsido, substituto=substituto,
                        matrsubstituto=matrsubstuto, departamento=departamento, aula=aula, data=data,
                        horas=horas)
    session.add(subs)
    session.commit()
    session.close()
    tkinter.messagebox.showinfo('Substituição ok!', 'Substituição lançada com sucesso!')


def lancar_hrscomple(nome, departamento, aula, data, horas):
    sessionscol = sessionmaker(bind=engine)
    sessioncol = sessionscol()
    pessoa = sessioncol.query(Colaborador).filter_by(nome=nome).first()
    matricula = pessoa.matricula
    sessions = sessionmaker(bind=enginefolha)
    session = sessions()
    hrcomp = Hrcomplement(professor=nome, matrprof=matricula, departamento=departamento, horas=horas, aula=aula, data=data)
    session.add(hrcomp)
    session.commit()
    session.close()
    tkinter.messagebox.showinfo('Horas Salvas!', 'Horas complementares salvas com sucesso!')


def lancar_faltas(nome, depto, data, hrs):
    sessions = sessionmaker(bind=enginefolha)
    session = sessions()
    sessionscol = sessionmaker(bind=engine)
    sessioncol = sessionscol()
    pessoa = sessioncol.query(Colaborador).filter_by(nome=nome).first()
    matricula = pessoa.matricula

    falt = Faltas(professor=nome, matrprof=matricula, departamento=depto, data=data, horas=hrs)
    session.add(falt)
    session.commit()
    session.close()
    tkinter.messagebox.showinfo('Falta Salva!', 'Falta salva com sucesso!')


def lancar_novaaula(nomeprof, depto, nomeaula, diasemana, inicio, fim, valor):
    sessions = sessionmaker(bind=enginefolha)
    session = sessions()
    inicio = inicio + ':00'
    fim = fim + ':00'
    sessionscol = sessionmaker(bind=engine)
    sessioncol = sessionscol()
    pessoa = sessioncol.query(Colaborador).filter_by(nome=nomeprof).first()
    matricula = pessoa.matricula
    hj = dt.today()
    aula = Aulas(nome=nomeaula, professor=nomeprof, departamento=depto, diadasemana=diasemana, inicio=inicio,
                 fim=fim, valor=valor, status='Ativa', iniciograde=dt.strftime(hj, '%d/%m/%Y'), matrprof=matricula)
    session.add(aula)
    session.commit()
    session.close()
    tkinter.messagebox.showinfo('Aula Salva!', 'Aula lançada com sucesso!')


def lancar_escala(nome, departamento, aula, data, horas):
    sessionscol = sessionmaker(bind=engine)
    sessioncol = sessionscol()
    pessoa = sessioncol.query(Colaborador).filter_by(nome=nome).first()
    matricula = pessoa.matricula
    sessions = sessionmaker(bind=enginefolha)
    session = sessions()
    esc = Escala(professor=nome, matrprof=matricula, departamento=departamento, horas=horas, aula=aula, data=data)
    session.add(esc)
    session.commit()
    session.close()
    tkinter.messagebox.showinfo('Escala Salva!', 'Escala salva com sucesso!')


def salvar_banco_aulas():
    sessions = sessionmaker(enginefolha)
    session = sessions()
    aula = session.query(Aulas).filter_by(status='Ativa').order_by(Aulas.professor).all()
    wb = l_w(rf'C:\Users\{os.getlogin()}\PycharmProjects\AutomacaoCia\src\models\static\files\Plan.xlsx', read_only=False)
    sh = wb['Planilha1']
    x = 2
    for a in aula:
        sh[f'A{x}'].value = a.numero
        sh[f'B{x}'].value = a.professor
        sh[f'C{x}'].value = a.nome
        sh[f'D{x}'].value = a.departamento
        sh[f'E{x}'].value = a.diadasemana
        sh[f'F{x}'].value = a.inicio
        sh[f'G{x}'].value = a.fim
        sh[f'H{x}'].value = a.valor
        sh[f'I{x}'].value = a.iniciograde
        sh[f'J{x}'].value = a.matrprof
        x += 1
    sh[f'A1'].value = 'Número'
    sh[f'B1'].value = 'Professor'
    sh[f'C1'].value = 'Nome da Aula'
    sh[f'D1'].value = 'Departamento'
    sh[f'E1'].value = 'Dia da semana'
    sh[f'F1'].value = 'Início'
    sh[f'G1'].value = 'Fim'
    sh[f'H1'].value = 'Valor'
    sh[f'I1'].value = 'Início Grade'
    sh[f'J1'].value = 'Matrícula'
    wb.save(rf'C:\Users\{os.getlogin()}\PycharmProjects\AutomacaoCia\src\models\static\files\CadastroAulas.xlsx')


def inativar_aulas(aulas: list):
    sessions = sessionmaker(enginefolha)
    session = sessions()
    for aula in aulas:
        a = session.query(Aulas).filter_by(numero=aula).first()
        a.status = 'Inativa'
        session.commit()
    tkinter.messagebox.showinfo(title='Aulas inativas!',message='Aulas inativas com sucesso!')
