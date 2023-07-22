from datetime import datetime as dt, timedelta as td
from dateutil.relativedelta import relativedelta
from modelsfolha import Aulas, engine
from sqlalchemy.orm import sessionmaker
from openpyxl import load_workbook as l_w
from openpyxl.styles import Color, PatternFill, Font, Border
import openpyxl.utils.cell
import locale
locale.setlocale(locale.LC_MONETARY, 'pt_BR.UTF-8')
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
Sessions = sessionmaker(bind=engine)
session = Sessions()
locale.setlocale(locale.LC_MONETARY, 'pt_BR.UTF-8')
data = int(input('Digite o mes da competência: '))


class Folha:
    def __init__(self, competencia, aulas, deptos):
        self.fechamento = dt(day=20, month=competencia, year=dt.today().year)
        self.compet = self.fechamento.month
        self.aulas = aulas
        self.dept = deptos


class Aula:
    def __init__(self, nome, prof, depart, diasem,inicio, fim, valorhr, iniciograde, fimgrade=''):
        self.nome = nome
        self.professor = prof
        self.departamento = depart
        self.dia = diasem
        self.inicio = dt.strptime(inicio, '%H:%M:%S')
        self.fim = dt.strptime(fim, '%H:%M:%S')
        self.valor = valorhr
        self.iniciograde = iniciograde
        self.fimgrade = fimgrade
        self.dsr = 1.1666


def somaaula(diasem, inic, fim, competencia, iniciograd, fimgrad):
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


def aulasativas():
    aulasativasdb = session.query(Aulas).filter_by(status='Ativa').all()
    aula = []
    for i, a in enumerate(aulasativasdb):
        aula.append(i)
        aula[i] = Aula(a.nome, a.professor, a.departamento, a.diadasemana, a.inicio, a.fim, a.valor, a.iniciograde, a.fimgrade)
        yield aula[i]


def deptosativos():
    aulasativasdb = session.query(Aulas).filter_by(status='Ativa').all()
    departamentos = []
    for i, a in enumerate(aulasativasdb):
        departamentos.append(a.departamento)
        departamentos = list(set(departamentos))
    return departamentos


def profsativos():
    aulasativasdb = session.query(Aulas).filter_by(status='Ativa').all()
    professores = []
    for i, a in enumerate(aulasativasdb):
        professores.append(a.professor)
        professores = list(set(professores))
    return professores


def totaldafolha(folha):
    somatorio = 0
    for al in list(aulasativas()):
        somatorio += somaaula(al.dia, al.inicio, al.fim, data, al.iniciograde, al.fimgrade) * float(str(al.valor).replace(',', '.')) * al.dsr
    return round(somatorio, 2)


def somaprof(folha, prof, depto, nome):
    somahoras = 0
    for aula in folha.aulas:
        if aula.professor == prof and aula.departamento == depto and aula.nome == nome:
            somahoras += somaaula(aula.dia, aula.inicio, aula.fim, data, aula.iniciograde, aula.fimgrade)
    return somahoras


def alteracoesfolha(dicionario):
    folha = dicionario
    hoje = dt.today()
    alt = l_w('AlteracoesGrade.xlsx', read_only=False)
    substituicoes = alt['Subs']
    faltas = alt['Falta']
    atestados = alt['Atestado']
    if hoje.day >= 21:
        inicio = dt(day=21, month=hoje.month, year=hoje.year)
        fechamento = dt(day=20, month=(hoje + relativedelta(months=1)).month,
                        year=(hoje + relativedelta(months=1)).year)
    if hoje.day < 21:
        inicio = dt(day=21, month=(hoje - relativedelta(months=1)).month,
                    year=(hoje - relativedelta(months=1)).year)
        fechamento = dt(day=20, month=hoje.month, year=hoje.year)
    
    return folha


def plandegrade(dic, comp):
    grade = l_w('Grade.xlsx', read_only=False)
    plan1 = grade['Planilha1']
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
    competencia = dt(day=1, month=comp, year=dt.today().year)
    inicio = dt(day=21, month=(competencia - relativedelta(months=1)).month, year=(competencia - relativedelta(months=1)).year)
    fechamento = dt(day=20, month=competencia.month, year=competencia.year)
    # primeira linha deve aparecer 'Folha' na coluna A1 e 'Julho' de '2023' na B1
    plan1['A1'].value = 'Folha'
    plan1['B1'].value = f'{fechamento.month} de {fechamento.year}'
    # na linha 3 a partir da célula C deve se iniciar os dias do intervalo de folha escritos como a inicial do dia da
    # semana

    def intervalo(inicio, fechamento):
        for n in range(int((fechamento - inicio).days) + 1):
            yield dt.strftime(inicio + td(n), '%d/%m')

    col = 3
    for item in list(intervalo(inicio, fechamento)):
        plan1.cell(column=col, row=3, value=dt.strftime(dt.strptime(item, '%d/%m'), '%a'))
        plan1.cell(column=col, row=4, value=item)
        col += 1
    plan1.cell(column=col, row=3, value='Total')

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

    plan1['A4'].value = 'Musculação'
    novalinha = 5
    for i in musculacao:
        for row in plan1.iter_cols(min_row=3, min_col=3, max_row=3, max_col=35):
            for cell in row:
                if cell.value != 'Total':
                    if cell.value == 'seg':
                        plan1.cell(column=cell.column, row=novalinha, value=somaseg(i, 'Musculação'))
                    if cell.value == 'ter':
                        plan1.cell(column=cell.column, row=novalinha, value=somater(i, 'Musculação'))
                    if cell.value == 'qua':
                        plan1.cell(column=cell.column, row=novalinha, value=somaqua(i, 'Musculação'))
                    if cell.value == 'qui':
                        plan1.cell(column=cell.column, row=novalinha, value=somaqui(i, 'Musculação'))
                    if cell.value == 'sex':
                        plan1.cell(column=cell.column, row=novalinha, value=somasex(i, 'Musculação'))
                    if cell.value == 'sáb':
                        plan1.cell(column=cell.column, row=novalinha, value=somasab(i, 'Musculação'))
                    if cell.value == 'dom':
                        plan1.cell(column=cell.column, row=novalinha, value=somadom(i, 'Musculação'))
                else:
                    letra = openpyxl.utils.cell.get_column_letter(cell.column - 1)
                    plan1.cell(column=cell.column, row=novalinha, value=f'=SUM(C{novalinha}:{letra}{novalinha})')
        plan1.cell(column=2, row=novalinha, value=i)
        novalinha = plan1.max_row + 1

    plan1[f'A{novalinha}'].value = 'Ginástica'
    novalinha = plan1.max_row + 1
    for i in ginastica:
        for row in plan1.iter_cols(min_row=3, min_col=3, max_row=3, max_col=35):
            for cell in row:
                if cell.value != 'Total':
                    if cell.value == 'seg':
                        plan1.cell(column=cell.column, row=novalinha, value=somaseg(i, 'Ginástica'))
                    if cell.value == 'ter':
                        plan1.cell(column=cell.column, row=novalinha, value=somater(i, 'Ginástica'))
                    if cell.value == 'qua':
                        plan1.cell(column=cell.column, row=novalinha, value=somaqua(i, 'Ginástica'))
                    if cell.value == 'qui':
                        plan1.cell(column=cell.column, row=novalinha, value=somaqui(i, 'Ginástica'))
                    if cell.value == 'sex':
                        plan1.cell(column=cell.column, row=novalinha, value=somasex(i, 'Ginástica'))
                    if cell.value == 'sáb':
                        plan1.cell(column=cell.column, row=novalinha, value=somasab(i, 'Ginástica'))
                    if cell.value == 'dom':
                        plan1.cell(column=cell.column, row=novalinha, value=somadom(i, 'Ginástica'))
                else:
                    letra = openpyxl.utils.cell.get_column_letter(cell.column - 1)
                    plan1.cell(column=cell.column, row=novalinha, value=f'=SUM(C{novalinha}:{letra}{novalinha})')
        plan1.cell(column=2, row=novalinha, value=i)
        novalinha = plan1.max_row + 1

    plan1[f'A{novalinha}'].value = 'Kids'
    novalinha = plan1.max_row + 1
    for i in kids:
        for row in plan1.iter_cols(min_row=3, min_col=3, max_row=3, max_col=35):
            for cell in row:
                if cell.value != 'Total':
                    if cell.value == 'seg':
                        plan1.cell(column=cell.column, row=novalinha, value=somaseg(i, 'Kids'))
                    if cell.value == 'ter':
                        plan1.cell(column=cell.column, row=novalinha, value=somater(i, 'Kids'))
                    if cell.value == 'qua':
                        plan1.cell(column=cell.column, row=novalinha, value=somaqua(i, 'Kids'))
                    if cell.value == 'qui':
                        plan1.cell(column=cell.column, row=novalinha, value=somaqui(i, 'Kids'))
                    if cell.value == 'sex':
                        plan1.cell(column=cell.column, row=novalinha, value=somasex(i, 'Kids'))
                    if cell.value == 'sáb':
                        plan1.cell(column=cell.column, row=novalinha, value=somasab(i, 'Kids'))
                    if cell.value == 'dom':
                        plan1.cell(column=cell.column, row=novalinha, value=somadom(i, 'Kids'))
                else:
                    letra = openpyxl.utils.cell.get_column_letter(cell.column - 1)
                    plan1.cell(column=cell.column, row=novalinha, value=f'=SUM(C{novalinha}:{letra}{novalinha})')
        plan1.cell(column=2, row=novalinha, value=i)
        novalinha = plan1.max_row + 1

    plan1[f'A{novalinha}'].value = 'Esportes'
    novalinha = plan1.max_row + 1
    for i in esportes:
        for row in plan1.iter_cols(min_row=3, min_col=3, max_row=3, max_col=35):
            for cell in row:
                if cell.value != 'Total':
                    if cell.value == 'seg':
                        plan1.cell(column=cell.column, row=novalinha, value=somaseg(i, 'Esportes'))
                    if cell.value == 'ter':
                        plan1.cell(column=cell.column, row=novalinha, value=somater(i, 'Esportes'))
                    if cell.value == 'qua':
                        plan1.cell(column=cell.column, row=novalinha, value=somaqua(i, 'Esportes'))
                    if cell.value == 'qui':
                        plan1.cell(column=cell.column, row=novalinha, value=somaqui(i, 'Esportes'))
                    if cell.value == 'sex':
                        plan1.cell(column=cell.column, row=novalinha, value=somasex(i, 'Esportes'))
                    if cell.value == 'sáb':
                        plan1.cell(column=cell.column, row=novalinha, value=somasab(i, 'Esportes'))
                    if cell.value == 'dom':
                        plan1.cell(column=cell.column, row=novalinha, value=somadom(i, 'Esportes'))
                else:
                    letra = openpyxl.utils.cell.get_column_letter(cell.column - 1)
                    plan1.cell(column=cell.column, row=novalinha, value=f'=SUM(C{novalinha}:{letra}{novalinha})')
        plan1.cell(column=2, row=novalinha, value=i)
        novalinha = plan1.max_row + 1

    plan1[f'A{novalinha}'].value = 'Cross Cia'
    novalinha = plan1.max_row + 1
    for i in cross:
        for row in plan1.iter_cols(min_row=3, min_col=3, max_row=3, max_col=35):
            for cell in row:
                if cell.value != 'Total':
                    if cell.value == 'seg':
                        plan1.cell(column=cell.column, row=novalinha, value=somaseg(i, 'Cross Cia'))
                    if cell.value == 'ter':
                        plan1.cell(column=cell.column, row=novalinha, value=somater(i, 'Cross Cia'))
                    if cell.value == 'qua':
                        plan1.cell(column=cell.column, row=novalinha, value=somaqua(i, 'Cross Cia'))
                    if cell.value == 'qui':
                        plan1.cell(column=cell.column, row=novalinha, value=somaqui(i, 'Cross Cia'))
                    if cell.value == 'sex':
                        plan1.cell(column=cell.column, row=novalinha, value=somasex(i, 'Cross Cia'))
                    if cell.value == 'sáb':
                        plan1.cell(column=cell.column, row=novalinha, value=somasab(i, 'Cross Cia'))
                    if cell.value == 'dom':
                        plan1.cell(column=cell.column, row=novalinha, value=somadom(i, 'Cross Cia'))
                else:
                    letra = openpyxl.utils.cell.get_column_letter(cell.column - 1)
                    plan1.cell(column=cell.column, row=novalinha, value=f'=SUM(C{novalinha}:{letra}{novalinha})')
        plan1.cell(column=2, row=novalinha, value=i)
        novalinha = plan1.max_row + 1

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
    for itens in plan1.iter_cols(min_row=3, min_col=3, max_row=3, max_col=35):
        for cell in itens:
            if cell.value != 'Total':
                if cell.value == 'sáb' or cell.value == 'dom':
                    letras = openpyxl.utils.cell.get_column_letter(cell.column)
                    for i in range(3,150):
                        plan1[f'{letras}{i}'].fill = fds

    plan1['C1'].fill = atestado
    plan1['D1'].value = 'Atestado'
    plan1['C2'].fill = falta
    plan1['D2'].value = 'Falta'
    plan1['F1'].fill = ferias
    plan1['G1'].value = 'Férias'
    plan1['F2'].fill = feriado
    plan1['G2'].value = 'Feriado'
    grade.save(f'Grade {fechamento.month}-{fechamento.year}.xlsx')


def ferias():
    pass


def atestado():
    pass


def falta(dias, aula):
    horas = aula.fim - aula.inicio
    hr, minut, seg = str(horas).split(':')
    somadia = int(hr) + int(minut) / 60 + int(seg) / 60 * 60
    somadia = round(somadia, 2)
    faltas = somadia * dias
    return round(faltas, 2)


def feriado():
    pass


def substituicaoacres(pessoa, inicio, fim):
    pass


def substituicaodesc():
    pass


def escala():
    pass


def somaseg(nome, depto):
    aulasseg = session.query(Aulas).filter_by(professor=nome).filter_by(diadasemana='Segunda')\
        .filter_by(departamento=depto).all()
    somas = 0
    for aula in aulasseg:
        hr, minut, seg = str(dt.strptime(aula.fim, '%H:%M:%S') - dt.strptime(aula.inicio, '%H:%M:%S')).split(':')
        somadia = int(hr) + int(minut) / 60 + int(seg) / 60 * 60
        somadia = round(somadia, 2)
        somas += somadia
    return somas


def somater(nome, depto):
    aulaster = session.query(Aulas).filter_by(professor=nome).filter_by(diadasemana='Terça')\
        .filter_by(departamento=depto).all()
    somas = 0
    for aula in aulaster:
        hr, minut, seg = str(dt.strptime(aula.fim, '%H:%M:%S') - dt.strptime(aula.inicio, '%H:%M:%S')).split(':')
        somadia = int(hr) + int(minut) / 60 + int(seg) / 60 * 60
        somadia = round(somadia, 2)
        somas += somadia
    return somas


def somaqua(nome, depto):
    aulasqua = session.query(Aulas).filter_by(professor=nome).filter_by(diadasemana='Quarta')\
        .filter_by(departamento=depto).all()
    somas = 0
    for aula in aulasqua:
        hr, minut, seg = str(dt.strptime(aula.fim, '%H:%M:%S') - dt.strptime(aula.inicio, '%H:%M:%S')).split(':')
        somadia = int(hr) + int(minut) / 60 + int(seg) / 60 * 60
        somadia = round(somadia, 2)
        somas += somadia
    return somas


def somaqui(nome, depto):
    aulasqui = session.query(Aulas).filter_by(professor=nome).filter_by(diadasemana='Quinta')\
        .filter_by(departamento=depto).all()
    somas = 0
    for aula in aulasqui:
        hr, minut, seg = str(dt.strptime(aula.fim, '%H:%M:%S') - dt.strptime(aula.inicio, '%H:%M:%S')).split(':')
        somadia = int(hr) + int(minut) / 60 + int(seg) / 60 * 60
        somadia = round(somadia, 2)
        somas += somadia
    return somas


def somasex(nome, depto):
    aulassex = session.query(Aulas).filter_by(professor=nome).filter_by(diadasemana='Sexta')\
        .filter_by(departamento=depto).all()
    somas = 0
    for aula in aulassex:
        hr, minut, seg = str(dt.strptime(aula.fim, '%H:%M:%S') - dt.strptime(aula.inicio, '%H:%M:%S')).split(':')
        somadia = int(hr) + int(minut) / 60 + int(seg) / 60 * 60
        somadia = round(somadia, 2)
        somas += somadia
    return somas


def somasab(nome, depto):
    aulassab = session.query(Aulas).filter_by(professor=nome).filter_by(diadasemana='Sábado')\
        .filter_by(departamento=depto).all()
    somas = 0
    for aula in aulassab:
        hr, minut, seg = str(dt.strptime(aula.fim, '%H:%M:%S') - dt.strptime(aula.inicio, '%H:%M:%S')).split(':')
        somadia = int(hr) + int(minut) / 60 + int(seg) / 60 * 60
        somadia = round(somadia, 2)
        somas += somadia
    return somas


def somadom(nome, depto):
    aulasdom = session.query(Aulas).filter_by(professor=nome).filter_by(diadasemana='Domingo')\
        .filter_by(departamento=depto).all()
    somas = 0
    for aula in aulasdom:
        hr, minut, seg = str(dt.strptime(aula.fim, '%H:%M:%S') - dt.strptime(aula.inicio, '%H:%M:%S')).split(':')
        somadia = int(hr) + int(minut) / 60 + int(seg) / 60 * 60
        somadia = round(somadia, 2)
        somas += somadia
    return somas


folhadehoje = Folha(8, list(aulasativas()), deptosativos())
somaaulas = {}
for i in profsativos():
    somaaulas[i] = {}
    for d in deptosativos():
        somaaulas[i][d] = {}
for aulas in aulasativas():
    somaaulas[aulas.professor][aulas.departamento][aulas.nome+f' ({aulas.valor})'] = round(somaprof(folhadehoje, aulas.professor, aulas.departamento, aulas.nome), 2)
    dictchav = list(somaaulas.keys())
    dictchav.sort()
    somafinal = {i: somaaulas[i] for i in dictchav}

plan = l_w('Somafinal.xlsx', read_only=False)
folha = plan['Planilha1']
folha['A1'].value = 'Matrícula'
folha['B1'].value = 'Nome'
folha['C1'].value = 'Aula e Valor'
folha['D1'].value = 'Horas'
x = 2
for i in somafinal:
    matr = session.query(Aulas).filter_by(professor=str(i)).first()
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
plan.save(f'Somafinal mes {data}.xlsx')
plandegrade(somafinal, data)
