import locale
from datetime import datetime as dt, timedelta as td
from dateutil.relativedelta import relativedelta
from modelsfolha import Aulas, engine
from sqlalchemy.orm import sessionmaker

data = dt.today()
Sessions = sessionmaker(bind=engine)
session = Sessions()
locale.setlocale(locale.LC_MONETARY, 'pt_BR.UTF-8')


class Folha:
    def __init__(self, competencia, aulas, deptos):
        if competencia.day >= 21:
            fechamento = dt(day=20, month=(competencia + relativedelta(months=1)).month,
                            year=(competencia + relativedelta(months=1)).year)
        if competencia.day < 21:
            fechamento = dt(day=20, month=competencia.month, year=competencia.year)
        self.compet = fechamento.month
        self.aulas = aulas
        self.dept = deptos


class Aula:
    def __init__(self, nome, prof, depart, diasem, inicio, fim, valorhr, iniciograde, fimgrade=''):
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


def somaaula(diasem, inic, fim, hoje, iniciograd, fimgrad):
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
    if hoje.day >= 21:
        inicio = dt(day=21, month=hoje.month, year=hoje.year)
        fechamento = dt(day=20, month=(hoje + relativedelta(months=1)).month,
                        year=(hoje + relativedelta(months=1)).year)
    if hoje.day < 21:
        inicio = dt(day=21, month=(hoje - relativedelta(months=1)).month,
                    year=(hoje - relativedelta(months=1)).year)
        fechamento = dt(day=20, month=hoje.month, year=hoje.year)

    def intervalo(inicio, fechamento):
        if hoje.day >= 21:
            inicio = dt(day=21, month=hoje.month, year=hoje.year)
            fechamento = dt(day=20, month=(hoje + relativedelta(months=1)).month,
                            year=(hoje + relativedelta(months=1)).year)
        if hoje.day < 21:
            inicio = dt(day=21, month=(hoje - relativedelta(months=1)).month,
                        year=(hoje - relativedelta(months=1)).year)
            fechamento = dt(day=20, month=hoje.month, year=hoje.year)

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


def falta(dias, aula):
    horas = aula.fim - aula.inicio
    hr, minut, seg = str(horas).split(':')
    somadia = int(hr) + int(minut) / 60 + int(seg) / 60 * 60
    somadia = round(somadia, 2)
    faltas = somadia * dias
    return round(faltas, 2)


def aulasativas():
    aulasativasdb = session.query(Aulas).filter_by(status='Ativa').all()
    aula = []
    for i, a in enumerate(aulasativasdb):
        aula.append(i)
        aula[i] = Aula(a.nome, a.professor, a.departamento, a.diadasemana, a.inicio, a.fim, a.valor, a.iniciograde,
                       a.fimgrade)
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


folhadehoje = Folha(data, list(aulasativas()), deptosativos())


def totaldafolha(folha):
    somatorio = 0
    for al in list(aulasativas()):
        somatorio += somaaula(al.dia, al.inicio, al.fim, data, al.iniciograde, al.fimgrade) * \
                     float(str(al.valor).replace(',', '.')) * al.dsr
    return round(somatorio, 2)


def somaprof(folha, prof, depto, nome):
    somahoras = 0
    for aula in folha.aulas:
        if aula.professor == prof and aula.departamento == depto and aula.nome == nome:
            somahoras += somaaula(aula.dia, aula.inicio, aula.fim, data, aula.iniciograde, aula.fimgrade)
    return somahoras


somaaulas = {}
for i in profsativos():
    somaaulas[i] = {}
    for d in deptosativos():
        somaaulas[i][d] = {}
for aulas in aulasativas():
    somaaulas[aulas.professor][aulas.departamento][aulas.nome+f'({aulas.valor})'] = \
        somaprof(folhadehoje, aulas.professor, aulas.departamento, aulas.nome)
print(somaaulas)
print(locale.currency(totaldafolha(folhadehoje), grouping=True))
