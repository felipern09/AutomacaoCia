import tkinter as tk
from tkcalendar import DateEntry
from datetime import datetime
import tkinter.filedialog
import pyautogui as pa
import pyperclip as pp
import time as t
from Admissao.models import Colaborador, engine
from sqlalchemy.orm import sessionmaker
from openpyxl import load_workbook as l_w
from listas import horarios, cargos, departamentos, tipodecontrato, municipios
import os
import tkinter.filedialog
from tkinter import ttk, messagebox
from tkinter import *
import docx
import docx2pdf
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email import encoders
from dados_servd import em_rem, em_ti, em_if, k1
from difflib import SequenceMatcher

matricula = int(input('Digite a matrícula desligada: '))
data_desligamento = int(input('Digite a matrícula desligada: '))
tipo_desligamento = int(input('Digite a matrícula desligada: '))
Sessions = sessionmaker(bind=engine)
session = Sessions()
pessoa = session.query(Colaborador).filter_by(matricula=matricula).first()
if pessoa:
    if pessoa.cargo == 'ESTAGIÁRIO' or pessoa.cargo == 'ESTAGIÁRIA':
        # mandar e-mail para if pedindo deligamento do TCE a partir da data de envio do e-mail, deve conter nome e cpf do estag
        # mandar e-mail para estag solicitando data para marcar devolução de uniformes, bts, e assinatura da rescisão
    else:
        # gerar docs de homologação no dexion: Rescisão 5 cópias, av prév, comprovantes recolhimento inss, carta preposto,
        # folha de registro, carta abono conduta, guia de seguro desemprego(?)
        if tipo_desligamento == 'Pedido':
            # e-mail informdando data de crédito na conta e solicitando data para marcar no sindicato e dev uniformes
        if tipo_desligamento == 'Demissao sem aviso':
            # e-mails com data do pgto, orientações do passo a passo, guias de orientação do FGTS e Seguro desemprego
            # explicar quanto saca do fgts
            # solicitar data para agendar no sindicato
        if tipo_desligamento == 'Demissão com aviso':
            # após gerada a rescisão e guia e-mail informanda dia do crédito em conta, guias de fgts e seguro
            #explicar quanto saca do fgts
            #e-mail marcando data para ir no sindicato, dev uniformes e bts
        if tipo_desligamento == 'Acordo':
            # após gerada a rescisão e guia e-mail informanda dia do crédito em conta, guias de fgts e seguro
            #explicar quanto saca do fgts
            #e-mail marcando data para ir no sindicato, dev uniformes e bts
    # e-mail para TI informando nome e CPF do funcionário/estagiário e solicitando o desligamento
    pessoa.desligamento = data_desligamento
    session.commit()

else:
    print('pessoa não cadastrada')
