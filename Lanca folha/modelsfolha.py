from sqlalchemy import create_engine
from sqlalchemy import MetaData
from sqlalchemy.orm import declarative_base
from sqlalchemy import Column, Integer, String

engine = create_engine("sqlite+pysqlite:///..\\Lanca folha\\aulas.db", echo=True, future=True)
metadata_obj = MetaData()
Base = declarative_base()


class Aulas(Base):
    __tablename__ = "colaborador"
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


Base.metadata.create_all(engine)
