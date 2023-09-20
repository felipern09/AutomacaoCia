from sqlalchemy import create_engine
from sqlalchemy import MetaData
from sqlalchemy.orm import declarative_base
from sqlalchemy import Column, Integer, String
import os
arqp = os.path.relpath(rf'C:\Users\{os.getlogin()}\PycharmProjects\AutomacaoCia\src\models\basevt.db')

enginevt = create_engine('sqlite+pysqlite:///' + arqp, echo=True, future=True)
metadata_obj = MetaData()
Base = declarative_base()


class BaseVT(Base):
    __tablename__ = "basevt"
    numero = Column(Integer, primary_key=True)
    nome = Column(String, nullable=False)
    matricula = Column(Integer, nullable=False)
    tipo_vt = Column(String, nullable=True)



Base.metadata.create_all(enginevt)
