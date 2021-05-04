import csv
import sys
from matplotlib import pyplot as plt

sys.path.append('c:\\Projetos\\Python\\Banco_Dados')
from BancodeDados import SQLServer

def Calculacomplexidade(complexidade,comp_baixa,comp_media,comp_alta):
    if complexidade == "Baixa":
        return comp_baixa
    elif complexidade == "Média":
        return comp_media
    elif complexidade == "Alta":
        return comp_alta
    else:
        return comp_baixa

def DefinePeriodo(PRAZO_FINAL):
    if int(PRAZO_FINAL[8:10]) > 20:
        if (int(PRAZO_FINAL[5:7]) < 12):
            return PRAZO_FINAL[:4] + '-' + str('0' + str(int(PRAZO_FINAL[5:7]) + 1))[-2:]
        else:
            return str(int(PRAZO_FINAL[:4]) + 1) + '-' + str('01')
    return PRAZO_FINAL[:-3]

def CarregaCSV():
    planilha = csv.DictReader(open("Demandas_BRQ_de_2016_ate_202104.csv", encoding='utf-8'), delimiter=';')
    CAIXA = SQLServer('CAIXA')
    CAIXA.ExecutaComandoSQL("truncate table demandas_brq")

    for linha in planilha:
        sql="INSERT INTO Demandas_BRQ (ID, RESUMO, STATUS, QTDE, COMPLEXIDADE, DATA_CRIACAO, PRAZO_FINAL, PREPOSTO, SOLICITANTE, SERVICO, UST, GRUPO) values ("
        sql += str(linha["ID"]) + ",'"
        sql += str(linha["Resumo"])   + "','"
        sql += str(linha["Status"]) + "',"
        sql += str(linha["Quantidade"]) + ",'"
        sql += str(linha["Complexidade"]) + "','"
        sql += str(linha["Data de Criação"]) + "','"
        sql += str(linha["Prazo Final"]) + "','"
        sql += str(linha["Preposto 2605"]) + "','"
        sql += str(linha["Segmento Solicitante de Apoio a Ferramentas Rational"]) + "','"
        sql += str(linha["Serviço Especializado de Apoio a Ferramentas Rational"]) + "',"
        sql += str(linha["UST"]) + ",'"
        sql += str(linha["Grupo SIGCT"]) + "')"

        CAIXA.ExecutaComandoSQL(sql)

def DefineFerramenta(RESUMO):
    RESUMO = RESUMO.upper()
    if ('CCASE' in RESUMO) or ('CLEARCASE' in RESUMO) or ('CCRC' in RESUMO) or ('VIEW' in RESUMO) or ('VOB' in RESUMO):
        return 'CCASE'
    if 'GC' in RESUMO:
        return 'GC'
    elif 'RDNG' in RESUMO:
        return 'RDNG'
    elif ('RTC' in RESUMO) or ('WORKITEM' in RESUMO) or ('GID' in RESUMO):
        return 'RTC'
    elif ('RQM' in RESUMO) or ('QUALITY' in RESUMO):
        return 'RQM'
    elif 'RFT' in RESUMO:
        return 'RFT'
    elif 'CLM' in RESUMO:
        return 'CLM'
    elif ('TESTE' in RESUMO) or ('RTW' in RESUMO) or ('RPT' in RESUMO) or ('RIT' in RESUMO):
        return 'TESTE'
    else:
        return 'GERAL'

def AtualizaRegistros():
    CAIXA = SQLServer('CAIXA')
    cursor = CAIXA.ConsultaSQL("select D.ID, D.QTDE, D.COMPLEXIDADE,S.COMPLEXIDADE_BAIXA, S.COMPLEXIDADE_MEDIA,S.COMPLEXIDADE_ALTA, PRAZO_FINAL, RESUMO from Demandas_BRQ D,Servicos S where D.SERVICO = S.SERVICO ORDER BY D.PRAZO_FINAL")
    reg_demanda = cursor.fetchone()
    while reg_demanda:
        UST_TOTAL = 0
        QTDE = int(reg_demanda[1])

        #Definir a complexidade
        COMPLEXIDADE = Calculacomplexidade(reg_demanda[2].strip(),reg_demanda[3],reg_demanda[4],reg_demanda[5])
        UST_TOTAL = COMPLEXIDADE * QTDE

        #Atualizar campo UST
        CAIXA.ExecutaComandoSQL("UPDATE Demandas_BRQ set UST = " + str(UST_TOTAL) + " where ID = " + str(reg_demanda[0]))
        reg_demanda = cursor.fetchone()

    CAIXA = SQLServer('CAIXA')
    cursor = CAIXA.ConsultaSQL("select ID, RESUMO , PRAZO_FINAL from Demandas_BRQ")
    reg_demanda = cursor.fetchone()
    while reg_demanda:
        #Definir o período da demanda
        PERIODO = DefinePeriodo(str(reg_demanda[2])[:10])

        #Definir a ferramenta
        FERRAMENTA = DefineFerramenta(reg_demanda[1])

        #Atualizar campos 'PERIODO' e 'FERRAMENTA'
        CAIXA.ExecutaComandoSQL("UPDATE Demandas_BRQ set PERIODO = '" + str(PERIODO) + "' , FERRAMENTA = '" + str(FERRAMENTA) +  "' where ID = " + str(reg_demanda[0]))

        reg_demanda = cursor.fetchone()

CarregaCSV()
AtualizaRegistros()