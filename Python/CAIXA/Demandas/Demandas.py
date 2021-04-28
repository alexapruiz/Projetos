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
    # PERIODOS 2016
    if (PRAZO_FINAL >= "2016-06-21") and (PRAZO_FINAL <= "2016-07-20"):
        return "2016-07"
    if (PRAZO_FINAL >= "2016-07-21") and (PRAZO_FINAL <= "2016-08-20"):
        return "2016-08"
    if (PRAZO_FINAL >= "2016-08-21") and (PRAZO_FINAL <= "2016-09-20"):
        return "2016-09"
    if (PRAZO_FINAL >= "2016-09-21") and (PRAZO_FINAL <= "2016-10-20"):
        return "2016-10"
    if (PRAZO_FINAL >= "2016-10-21") and (PRAZO_FINAL <= "2016-11-20"):
        return "2016-11"
    if (PRAZO_FINAL >= "2016-11-21") and (PRAZO_FINAL <= "2016-12-20"):
        return "2016-12"

    # PERIODOS 2017
    if (PRAZO_FINAL >= "2016-12-21") and (PRAZO_FINAL <= "2017-01-20"):
        return "2017-01"
    if (PRAZO_FINAL >= "2017-01-21") and (PRAZO_FINAL <= "2017-02-20"):
        return "2017-02"
    if (PRAZO_FINAL >= "2017-02-21") and (PRAZO_FINAL <= "2017-03-20"):
        return "2017-03"
    if (PRAZO_FINAL >= "2017-03-21") and (PRAZO_FINAL <= "2017-04-20"):
        return "2017-04"
    if (PRAZO_FINAL >= "2017-04-21") and (PRAZO_FINAL <= "2017-05-20"):
        return "2017-05"
    if (PRAZO_FINAL >= "2017-05-21") and (PRAZO_FINAL <= "2017-06-20"):
        return "2017-06"
    if (PRAZO_FINAL >= "2017-06-21") and (PRAZO_FINAL <= "2017-07-20"):
        return "2017-07"
    if (PRAZO_FINAL >= "2017-07-21") and (PRAZO_FINAL <= "2017-08-20"):
        return "2017-08"
    if (PRAZO_FINAL >= "2017-08-21") and (PRAZO_FINAL <= "2017-09-20"):
        return "2017-09"
    if (PRAZO_FINAL >= "2017-09-21") and (PRAZO_FINAL <= "2017-10-20"):
        return "2017-10"
    if (PRAZO_FINAL >= "2017-10-21") and (PRAZO_FINAL <= "2017-11-20"):
        return "2017-11"
    if (PRAZO_FINAL >= "2017-11-21") and (PRAZO_FINAL <= "2017-12-20"):
        return "2017-12"

    # PERIODOS 2018
    if (PRAZO_FINAL >= "2017-12-21") and (PRAZO_FINAL <= "2018-01-20"):
        return "2018-01"
    if (PRAZO_FINAL >= "2018-01-21") and (PRAZO_FINAL <= "2018-02-20"):
        return "2018-02"
    if (PRAZO_FINAL >= "2018-02-21") and (PRAZO_FINAL <= "2018-03-20"):
        return "2018-03"
    if (PRAZO_FINAL >= "2018-03-21") and (PRAZO_FINAL <= "2018-04-20"):
        return "2018-04"
    if (PRAZO_FINAL >= "2018-04-21") and (PRAZO_FINAL <= "2018-05-20"):
        return "2018-05"
    if (PRAZO_FINAL >= "2018-05-21") and (PRAZO_FINAL <= "2018-06-20"):
        return "2018-06"
    if (PRAZO_FINAL >= "2018-06-21") and (PRAZO_FINAL <= "2018-07-20"):
        return "2018-07"
    if (PRAZO_FINAL >= "2018-07-21") and (PRAZO_FINAL <= "2018-08-20"):
        return "2018-08"
    if (PRAZO_FINAL >= "2018-08-21") and (PRAZO_FINAL <= "2018-09-20"):
        return "2018-09"
    if (PRAZO_FINAL >= "2018-09-21") and (PRAZO_FINAL <= "2018-10-20"):
        return "2018-10"
    if (PRAZO_FINAL >= "2018-10-21") and (PRAZO_FINAL <= "2018-11-20"):
        return "2018-11"
    if (PRAZO_FINAL >= "2018-11-21") and (PRAZO_FINAL <= "2018-12-20"):
        return "2018-12"

    # PERIODOS 2019
    if (PRAZO_FINAL >= "2018-12-21") and (PRAZO_FINAL <= "2019-01-20"):
        return "2019-01"
    if (PRAZO_FINAL >= "2019-01-21") and (PRAZO_FINAL <= "2019-02-20"):
        return "2019-02"
    if (PRAZO_FINAL >= "2019-02-21") and (PRAZO_FINAL <= "2019-03-20"):
        return "2019-03"
    if (PRAZO_FINAL >= "2019-03-21") and (PRAZO_FINAL <= "2019-04-20"):
        return "2019-04"
    if (PRAZO_FINAL >= "2019-04-21") and (PRAZO_FINAL <= "2019-05-20"):
        return "2019-05"
    if (PRAZO_FINAL >= "2019-05-21") and (PRAZO_FINAL <= "2019-06-20"):
        return "2019-06"
    if (PRAZO_FINAL >= "2019-06-21") and (PRAZO_FINAL <= "2019-07-20"):
        return "2019-07"
    if (PRAZO_FINAL >= "2019-07-21") and (PRAZO_FINAL <= "2019-08-20"):
        return "2019-08"
    if (PRAZO_FINAL >= "2019-08-21") and (PRAZO_FINAL <= "2019-09-20"):
        return "2019-09"
    if (PRAZO_FINAL >= "2019-09-21") and (PRAZO_FINAL <= "2019-10-20"):
        return "2019-10"
    if (PRAZO_FINAL >= "2019-10-21") and (PRAZO_FINAL <= "2019-11-20"):
        return "2019-11"
    if (PRAZO_FINAL >= "2019-11-21") and (PRAZO_FINAL <= "2019-12-20"):
        return "2019-12"

    # PERIODOS 2020
    if (PRAZO_FINAL >= "2019-12-21") and (PRAZO_FINAL <= "2020-01-20"):
        return "2020-01"
    if (PRAZO_FINAL >= "2020-01-21") and (PRAZO_FINAL <= "2020-02-20"):
        return "2020-02"
    if (PRAZO_FINAL >= "2020-02-21") and (PRAZO_FINAL <= "2020-03-20"):
        return "2020-03"
    if (PRAZO_FINAL >= "2020-03-21") and (PRAZO_FINAL <= "2020-04-20"):
        return "2020-04"
    if (PRAZO_FINAL >= "2020-04-21") and (PRAZO_FINAL <= "2020-05-20"):
        return "2020-05"
    if (PRAZO_FINAL >= "2020-05-21") and (PRAZO_FINAL <= "2020-06-20"):
        return "2020-06"
    if (PRAZO_FINAL >= "2020-06-21") and (PRAZO_FINAL <= "2020-07-20"):
        return "2020-07"
    if (PRAZO_FINAL >= "2020-07-21") and (PRAZO_FINAL <= "2020-08-20"):
        return "2020-08"
    if (PRAZO_FINAL >= "2020-08-21") and (PRAZO_FINAL <= "2020-09-20"):
        return "2020-09"
    if (PRAZO_FINAL >= "2020-09-21") and (PRAZO_FINAL <= "2020-10-20"):
        return "2020-10"
    if (PRAZO_FINAL >= "2020-10-21") and (PRAZO_FINAL <= "2020-11-20"):
        return "2020-11"
    if (PRAZO_FINAL >= "2020-11-21") and (PRAZO_FINAL <= "2020-12-20"):
        return "2020-12"

    # PERIODOS 2021
    if (PRAZO_FINAL >= "2020-12-21") and (PRAZO_FINAL <= "2021-01-20"):
        return "2021-01"
    if (PRAZO_FINAL >= "2021-01-21") and (PRAZO_FINAL <= "2021-02-20"):
        return "2021-02"
    if (PRAZO_FINAL >= "2021-02-21") and (PRAZO_FINAL <= "2021-03-20"):
        return "2021-03"
    if (PRAZO_FINAL >= "2021-03-21") and (PRAZO_FINAL <= "2021-04-20"):
        return "2021-04"
    if (PRAZO_FINAL >= "2021-04-21") and (PRAZO_FINAL <= "2021-05-20"):
        return "2021-05"
    if (PRAZO_FINAL >= "2021-05-21") and (PRAZO_FINAL <= "2021-06-20"):
        return "2021-06"
    if (PRAZO_FINAL >= "2021-06-21") and (PRAZO_FINAL <= "2021-07-20"):
        return "2021-07"
    if (PRAZO_FINAL >= "2021-07-21") and (PRAZO_FINAL <= "2021-08-20"):
        return "2021-08"
    if (PRAZO_FINAL >= "2021-08-21") and (PRAZO_FINAL <= "2021-09-20"):
        return "2021-09"
    if (PRAZO_FINAL >= "2021-09-21") and (PRAZO_FINAL <= "2021-10-20"):
        return "2021-10"
    if (PRAZO_FINAL >= "2021-10-21") and (PRAZO_FINAL <= "2021-11-20"):
        return "2021-11"
    if (PRAZO_FINAL >= "2021-11-21") and (PRAZO_FINAL <= "2021-12-20"):
        return "2021-12"

def CarregaCSV():
    planilha = csv.DictReader(open("Demandas_BRQ_de_2016_ate_202104.csv", encoding='utf-8'), delimiter=';')
    CAIXA = SQLServer('CAIXA')
    CAIXA.ExecutaComandoSQL("truncate table demandas_brq")

    for linha in planilha:
        sql="INSERT INTO Demandas_BRQ (ID, RESUMO, STATUS, QTDE, COMPLEXIDADE, DATA_CRIACAO, PRAZO_FINAL, SOLICITANTE, SERVICO, UST, GRUPO) values ("
        sql += str(linha["ID"]) + ",'"
        sql += str(linha["Resumo"])   + "','"
        sql += str(linha["Status"]) + "',"
        sql += str(linha["Quantidade"]) + ",'"
        sql += str(linha["Complexidade"]) + "','"
        sql += str(linha["Data de Criação"]) + "','"
        sql += str(linha["Prazo Final"]) + "','"
        sql += str(linha["Segmento Solicitante de Apoio a Ferramentas Rational"]) + "','"
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