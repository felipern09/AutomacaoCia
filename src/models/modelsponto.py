from sqlalchemy import create_engine
from sqlalchemy import MetaData
from sqlalchemy.orm import declarative_base
from sqlalchemy import Column, Integer, String
import os
arqp = os.path.relpath(rf'C:\Users\{os.getlogin()}\PycharmProjects\AutomacaoCia\src\models\baseponto.db')

engineponto = create_engine('sqlite+pysqlite:///' + arqp, echo=True, future=True)
metadata_obj = MetaData()
Base = declarative_base()


class BasePonto(Base):
    __tablename__ = "baseponto"
    numero = Column(Integer, primary_key=True)
    nome = Column(String, nullable=False)
    matricula = Column(Integer, nullable=False)
    pis = Column(String, nullable=True)
    matrponto = Column(String, nullable=False)
    email = Column(String, nullable=False)
    cargo = Column(String, nullable=False)
    departamento = Column(String, nullable=False)


Base.metadata.create_all(engineponto)
