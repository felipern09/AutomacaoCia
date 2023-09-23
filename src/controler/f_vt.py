import locale
from src.models.modelsvt import enginevt, BaseVT
from src.models.models import Colaborador, engine
from sqlalchemy.orm import sessionmaker

locale.setlocale(locale.LC_MONETARY, 'pt_BR.UTF-8')
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
locale.setlocale(locale.LC_MONETARY, 'pt_BR.UTF-8')


def incluir_vt(nome: str, tipo: int):
    if tipo == 1:
        tp = 'BRB'
    elif tipo == 2:
        tp = 'Valecard'
    else:
        tp = 'Goiás'
    sessions = sessionmaker(engine)
    session = sessions()
    p = session.query(Colaborador).filter_by(nome=nome).order_by(Colaborador.matricula.desc()).first()
    matricula = p.matricula
    sessionsvt = sessionmaker(enginevt)
    sessionvt = sessionsvt()
    pessoa = BaseVT(nome=nome, matricula=matricula, tipo_vt=tp)
    sessionvt.add(pessoa)
    sessionvt.commit()


def retirar_vt(nome: str, tipo: int):
    if tipo == 1:
        tp = 'BRB'
    elif tipo == 2:
        tp = 'Valecard'
    else:
        tp = 'Goiás'
    sessionsvt = sessionmaker(enginevt)
    sessionvt = sessionsvt()
    pessoa = sessionvt.query(BaseVT).filter_by(nome=nome).filter_by(tipo_vt=tp).first()
    sessionvt.delete(pessoa)
    sessionvt.commit()


def gerar_vt(tipo: str):
    print(tipo)
