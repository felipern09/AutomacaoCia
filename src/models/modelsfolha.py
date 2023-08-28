from sqlalchemy import create_engine
from sqlalchemy import MetaData
from sqlalchemy.orm import declarative_base
from sqlalchemy import Column, Integer, String
from datetime import datetime as dt
import os
arqf = os.path.relpath(rf'C:\Users\{os.getlogin()}\PycharmProjects\AutomacaoCia\src\models\aulas.db')

enginefolha = create_engine('sqlite+pysqlite:///' + arqf, echo=True, future=True)
metadata_obj = MetaData()
Base = declarative_base()


class Aulas(Base):
    __tablename__ = "aulas"
    numero = Column(Integer, primary_key=True)
    nome = Column(String, nullable=False)
    professor = Column(String, nullable=False)
    matrprof = Column(Integer, nullable=False)
    departamento = Column(String, nullable=False)
    diadasemana = Column(String, nullable=False)
    inicio = Column(String, nullable=False)
    fim = Column(String, nullable=False)
    valor = Column(String, nullable=False)
    status = Column(String, nullable=False)
    iniciograde = Column(String, nullable=False)
    fimgrade = Column(String, nullable=True)


class Faltas(Base):
    __tablename__ = "faltas"
    numero = Column(Integer, primary_key=True)
    professor = Column(String, nullable=False)
    matrprof = Column(Integer, nullable=False)
    departamento = Column(String, nullable=False)
    data = Column(String, nullable=False)
    horas = Column(String, nullable=False)


class Ferias(Base):
    __tablename__ = 'ferias'
    numero = Column(Integer, primary_key=True)
    professor = Column(String, nullable=False)
    matrprof = Column(Integer, nullable=False)
    departamento = Column(String, nullable=False)
    inicio = Column(String, nullable=False)
    fim = Column(String, nullable=False)


class Atestado(Base):
    __tablename__ = 'atestado'
    numero = Column(Integer, primary_key=True)
    professor = Column(String, nullable=False)
    matrprof = Column(Integer, nullable=False)
    departamento = Column(String, nullable=False)
    data = Column(String, nullable=False)


class Substituicao(Base):
    __tablename__ = 'substituicao'
    numero = Column(Integer, primary_key=True)
    professorsubst = Column(String, nullable=False)
    matrprof = Column(Integer, nullable=False)
    substituto = Column(String, nullable=False)
    matrsubstituto = Column(String, nullable=False)
    departamento = Column(String, nullable=False)
    aula = Column(String, nullable=False)
    data = Column(String, nullable=False)
    horas = Column(String, nullable=False)


class Desligados(Base):
    __tablename__ = 'desligados'
    numero = Column(Integer, primary_key=True)
    professor = Column(String, nullable=False)
    matrprof = Column(Integer, nullable=False)
    departamento = Column(String, nullable=False)
    datadesligamento = Column(String, nullable=False)


class Escala(Base):
    __tablename__ = 'escala'
    numero = Column(Integer, primary_key=True)
    professor = Column(String, nullable=False)
    matrprof = Column(Integer, nullable=False)
    departamento = Column(String, nullable=False)
    aula = Column(String, nullable=False)
    horas = Column(String, nullable=False)
    data = Column(String, nullable=False)


class Hrcomplement(Base):
    __tablename__ = 'hrcomplementar'
    numero = Column(Integer, primary_key=True)
    professor = Column(String, nullable=False)
    matrprof = Column(Integer, nullable=False)
    departamento = Column(String, nullable=False)
    aula = Column(String, nullable=False)
    horas = Column(String, nullable=False)
    data = Column(String, nullable=False)


Base.metadata.create_all(enginefolha)


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

