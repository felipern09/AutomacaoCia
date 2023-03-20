from sqlalchemy import create_engine, text
from sqlalchemy import MetaData
from sqlalchemy.orm import declarative_base, sessionmaker
from sqlalchemy import Table, Column, Integer, Float, String, ForeignKey, Boolean, DateTime
from sqlalchemy.orm import relationship

engine = create_engine("sqlite+pysqlite:///colaboradores.db", echo=True, future=True)

metadata_obj = MetaData()
Base = declarative_base()


class Colaborador(Base):
    __tablename__ = "colaborador"
    matricula = Column(Integer, primary_key=True)
    nome = Column(String, nullable=True)
    admiss = Column(String, nullable=True)
    desligamento = Column(String, nullable=True)
    nascimento = Column(String, nullable=True)
    pis = Column(String, nullable=True)
    cpf = Column(String, nullable=True)
    rg = Column(String, nullable=True)
    emissor = Column(String, nullable=True)
    email = Column(String, nullable=True)
    genero = Column(String, nullable=True)
    estado_civil = Column(String, nullable=True)
    cor = Column(String, nullable=True)
    instru = Column(String, nullable=True)
    nacional = Column(String, nullable=True)
    cod_municipionas = Column(String, nullable=True)
    cid_nas = Column(String, nullable=True)
    uf_nas = Column(String, nullable=True)
    pai = Column(String, nullable=True)
    mae = Column(String, nullable=True)
    endereco = Column(String, nullable=True)
    num = Column(String, nullable=True)
    bairro = Column(String, nullable=True)
    cep = Column(String, nullable=True)
    cidade = Column(String, nullable=True)
    uf = Column(String, nullable=True)
    cod_municipioend = Column(String, nullable=True)
    tel = Column(String, nullable=True)
    tit_eleit = Column(String, nullable=True)
    zona_eleit = Column(String, nullable=True)
    sec_eleit = Column(String, nullable=True)
    ctps = Column(String, nullable=True)
    serie_ctps = Column(String, nullable=True)
    uf_ctps = Column(String, nullable=True)
    emiss_ctps = Column(String, nullable=True)
    depto = Column(String, nullable=True)
    cargo = Column(String, nullable=True)
    horario = Column(String, nullable=True)
    salario = Column(String, nullable=True)
    tipo_contr = Column(String, nullable=True)
    hr_sem = Column(Float, nullable=True)
    hr_mens = Column(Float, nullable=True)
    est_semestre = Column(String, nullable=True)
    est_turno = Column(String, nullable=True)
    est_prev_conclu = Column(String, nullable=True)
    est_faculdade = Column(String, nullable=True)
    est_endfacul = Column(String, nullable=True)
    est_numendfacul = Column(String, nullable=True)
    est_bairroendfacul = Column(String, nullable=True)


Base.metadata.create_all(engine)
