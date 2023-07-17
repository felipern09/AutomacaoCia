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
    nome = Column(String, nullable=True)
    professor = Column(String, nullable=True)
    departamento = Column(String, nullable=True)
    diadasemana = Column(String, nullable=True)
    inicio = Column(String, nullable=True)
    fim = Column(String, nullable=True)
    valor = Column(String, nullable=True)
    status = Column(String, nullable=True)
    iniciograde = Column(String, nullable=True)
    fimgrade = Column(String, nullable=True)


Base.metadata.create_all(engine)
