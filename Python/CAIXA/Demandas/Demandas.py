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
    if (PRAZO_FINAL >= "2020-01-01") and (PRAZO_FINAL <= "2020-01-20"):
        return "2020-01"
    elif (PRAZO_FINAL >= "2020-01-21") and (PRAZO_FINAL <= "2020-02-20"):
        return "2020-02"
    elif (PRAZO_FINAL >= "2020-02-21") and (PRAZO_FINAL <= "2020-03-20"):
        return "2020-03"
    elif (PRAZO_FINAL >= "2020-03-21") and (PRAZO_FINAL <= "2020-04-20"):
        return "2020-04"
    elif (PRAZO_FINAL >= "2020-04-21") and (PRAZO_FINAL <= "2020-05-20"):
        return "2020-05"
    elif (PRAZO_FINAL >= "2020-05-21") and (PRAZO_FINAL <= "2020-06-20"):
        return "2020-06"
    elif (PRAZO_FINAL >= "2020-06-21") and (PRAZO_FINAL <= "2020-07-20"):
        return "2020-07"
    elif (PRAZO_FINAL >= "2020-07-21") and (PRAZO_FINAL <= "2020-08-20"):
        return "2020-08"
    elif (PRAZO_FINAL >= "2020-08-21") and (PRAZO_FINAL <= "2020-09-20"):
        return "2020-09"
    elif (PRAZO_FINAL >= "2020-09-21") and (PRAZO_FINAL <= "2020-10-20"):
        return "2020-10"
    elif (PRAZO_FINAL >= "2020-10-21") and (PRAZO_FINAL <= "2020-11-20"):
        return "2020-11"
    elif (PRAZO_FINAL >= "2020-11-21") and (PRAZO_FINAL <= "2020-12-20"):
        return "2020-12"


def CarregaCSV():
    planilha = csv.DictReader(open("Demandas_BRQ_01_10_2020.csv", encoding='utf-8'), delimiter=';')
    CAIXA = SQLServer('CAIXA')
    CAIXA.ExecutaComandoSQL("truncate table demandas_brq")

    for linha in planilha:
        sql="INSERT INTO Demandas_BRQ (ID, RESUMO, STATUS, QTDE, COMPLEXIDADE, DATA_CRIACAO, PRAZO_FINAL, SOLICITANTE, PREPOSTO, SERVICO, UST, GRUPO) values ("
        sql += str(linha["ID"]) + ",'"
        sql += str(linha["Resumo"])   + "','"
        sql += str(linha["Status"]) + "',"
        sql += str(linha["Quantidade"]) + ",'"
        sql += str(linha["Complexidade"]) + "','"
        sql += str(linha["Data de Criação"]) + "','"
        sql += str(linha["Prazo Final"]) + "','"
        sql += str(linha["Segmento Solicitante de Apoio a Ferramentas Rational"]) + "','"
        sql += str(linha["Preposto 2605"]) + "','"
        sql += str(linha["Serviço Especializado de Apoio a Ferramentas Rational"]) + "',"
        sql += str(linha["UST"]) + ",'"
        sql += str(linha["Grupo SIGCT"]) + "')"

        CAIXA.ExecutaComandoSQL(sql)

def DefineFerramenta(RESUMO):
    if 'CCASE' in RESUMO:
        return 'CCASE'
    elif 'RDNG' in RESUMO:
        return 'RDNG'
    elif 'RTC' in RESUMO:
        return 'RTC'
    elif 'RQM' in RESUMO:
        return 'RQM'
    elif 'TESTE' in RESUMO:
        return 'TESTE'
    elif 'RFT' in RESUMO:
        return 'RFT'
    elif 'CLM' in RESUMO:
        return 'CLM'


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

        #Definir o período da demanda
        PERIODO = DefinePeriodo(str(reg_demanda[6])[:10])

        #Definir a ferramenta
        FERRAMENTA = DefineFerramenta(reg_demanda[7])

        #Atualizar campo UST
        CAIXA.ExecutaComandoSQL("UPDATE Demandas_BRQ set UST = " + str(UST_TOTAL) + " , PERIODO = '" + str(PERIODO) + "', FERRAMENTA ='" + str(FERRAMENTA) + "' where ID = " + str(reg_demanda[0]))
        reg_demanda = cursor.fetchone()

        #Atualizar campo FERRAMENTA


def LeDados():
    CAIXA = SQLServer("CAIXA")
    cursor = CAIXA.ConsultaSQL("select PERIODO, sum(UST) as USTs from Demandas_BRQ where PERIODO is not null group by PERIODO ORDER BY USTs")
    dados = cursor.fetchone()
    Grupos=[]
    Valores=[]
    while dados:
        Grupos.append(str(dados[0]).strip())
        Valores.append(str(dados[1]).strip())
        dados = cursor.fetchone()


CarregaCSV()
AtualizaRegistros()
LeDados()