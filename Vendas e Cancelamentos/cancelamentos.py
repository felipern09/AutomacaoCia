import pandas as pd

base = pd.read_excel('Cancelados 2021-2023.xlsx')
"""
 tipos de planos: ANUAL,ANUAL JR.,BI-ANUAL JR.,SEMESTRAL JR.,BI-ANUAL,ANUAL COLLEGE,ANUAL DIRETO,ANUAL DÉBITO PROGRAMADO
 ANUAL DÉBITO PROGRAMADO COLLEGE,ANUAL DÉBITO PROGRAMADO KIDS,ANUAL DÉBITO PROGRAMADO TEENS,ANUAL PLATINUM
 ANUAL PLATINUM DIRETO,MENSAL,MENSAL DCC,PLANO QUINZE MESES,PLANO TEMPO CERTO ANUAL,PLANO TEMPO CERTO JIU JITSU,
 SEMESTRAL,SEMESTRAL COLLEGE,SEMESTRAL PLATINUM,TRIMESTRAL,TRIMESTRAL JR.,TRIMESTRAL PLATINUM,BI-ANUAL PLATINUM
 Anual Jiu Jitsu,BI-ANUAL COLLEGE,ELO ANUAL DÉBITO PROGRAMADO COLLEGE,MENSAL CIA OUTDOOR 1X,MENSAL CIA OUTDOOR 2X
 MENSAL CIA OUTDOOR SÁBADO 1X, MENSAL JR,MENSAL PLATINUM, PLANO MENSAL COLLEGE
"""
# excluir planos que tiveram renovação
base = base.loc[base['Tem renovação'] == 'N']
# excluir planos gympass
base = base.loc[base['GymPass'] == 'N']
# confrontar com base de ativos e ver se algum desses fez novo plano (que não se classifica como renovação) -- excluir
# cruzar base de dados com turma/aula frequentada
# cruzar base de dados com baixa frequentcia
# cruzar com base de dados de ROs/Reclamações
total = base['Nome'].count()
# extrair da base de cancelados planos kids
basekids = base[base['Idade'] <= 10]
kids = basekids['Idade'].count()
# extratir da base de cancelados planos teens
baseteens = base.loc[(base['Idade'] >= 11) & (base['Idade'] <= 18)]
teens = baseteens['Idade'].count()
# extratir da base de cancelados planos college
basecollege = base.loc[(base['Idade'] >= 19) & (base['Idade'] <= 25)]
college = basecollege['Idade'].count()
# extrair da base planos adultos
baseadultos = base.loc[(base['Idade'] >= 26) & (base['Idade'] <= 64)]
adultos = baseadultos['Idade'].count()
# extrair da base planos platinum
baseplatinum = base[base['Idade'] >= 65]
platinum = baseplatinum['Idade'].count()
# extrair numero de cancelados
basecancelados = base[base['Tipo'] == 'Cancelados']
cancelados = basecancelados['Tipo'].count()
# extrair numero de encerrados
baseencerrados = base[base['Tipo'] == 'Ecerrados']
encerrados = baseencerrados['Tipo'].count()

print(
    f'Kids: {kids}\n'
    f'Teens: {teens}\n'
    f'College: {college}\n'
    f'Adultos: {adultos}\n'
    f'Platinum: {platinum}\n'
    f'TOTAL: {total}\n'
)