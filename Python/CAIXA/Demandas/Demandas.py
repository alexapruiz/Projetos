import csv
from BancoDados import BancoDeDados
from matplotlib import pyplot as plt

def calculacomplexidade(complexidade,comp_baixa,comp_media,comp_alta):
    if complexidade == "Baixa":
        return comp_baixa
    elif complexidade == "Média":
        return comp_media
    elif complexidade == "Alta":
        return comp_alta
    else:
        return 0

def DefinePeriodo(PRAZO_FINAL):
    if (PRAZO_FINAL >= "2020-01-01") and (PRAZO_FINAL <= "2020-01-20"):
        return "01/2020"
    elif (PRAZO_FINAL >= "2020-01-21") and (PRAZO_FINAL <= "2020-02-20"):
        return "02/2020"
    elif (PRAZO_FINAL >= "2020-02-21") and (PRAZO_FINAL <= "2020-03-20"):
        return "03/2020"
    elif (PRAZO_FINAL >= "2020-03-21") and (PRAZO_FINAL <= "2020-04-20"):
        return "04/2020"
    elif (PRAZO_FINAL >= "2020-04-21") and (PRAZO_FINAL <= "2020-05-20"):
        return "05/2020"
    elif (PRAZO_FINAL >= "2020-05-21") and (PRAZO_FINAL <= "2020-06-20"):
        return "06/2020"
    elif (PRAZO_FINAL >= "2020-06-21") and (PRAZO_FINAL <= "2020-07-20"):
        return "07/2020"
    elif (PRAZO_FINAL >= "2020-07-21") and (PRAZO_FINAL <= "2020-08-20"):
        return "08/2020"
    elif (PRAZO_FINAL >= "2020-08-21") and (PRAZO_FINAL <= "2020-09-20"):
        return "09/2020"
    elif (PRAZO_FINAL >= "2020-09-21") and (PRAZO_FINAL <= "2020-10-20"):
        return "10/2020"
    elif (PRAZO_FINAL >= "2020-10-21") and (PRAZO_FINAL <= "2020-11-20"):
        return "11/2020"
    elif (PRAZO_FINAL >= "2020-11-21") and (PRAZO_FINAL <= "2020-12-20"):
        return "12/2020"

def CarregaCSV():
    planilha = csv.DictReader(open("Demandas_BRQ_01_10_2020.csv", encoding='utf-8'), delimiter=';')
    BancoDeDados.ExecutaComandoSQL("delete from demandas_brq")

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

        BancoDeDados.ExecutaComandoSQL(sql)

def AtualizaRegistros():
    cursor = BancoDeDados.ConsultaSQL("select D.ID, D.QTDE, D.COMPLEXIDADE,S.COMPLEXIDADE_BAIXA, S.COMPLEXIDADE_MEDIA,S.COMPLEXIDADE_ALTA, PRAZO_FINAL from Demandas_BRQ D,Servicos S where D.SERVICO = S.SERVICO ORDER BY D.PRAZO_FINAL")
    reg_demanda = cursor.fetchone()
    while reg_demanda:
        UST_TOTAL = 0
        QTDE = int(reg_demanda[1])

        #Definir a complexidade
        COMPLEXIDADE = calculacomplexidade(reg_demanda[2].strip(),reg_demanda[3],reg_demanda[4],reg_demanda[5])
        UST_TOTAL = COMPLEXIDADE * QTDE

        #Definir o período da demanda
        PERIODO = DefinePeriodo(str(reg_demanda[6])[:10])

        #Atualizar campo UST
        BancoDeDados.ExecutaComandoSQL("UPDATE Demandas_BRQ set UST = " + str(UST_TOTAL) + " , PERIODO = '" + str(PERIODO) + "' where ID = " + str(reg_demanda[0]))
        reg_demanda = cursor.fetchone()

def CriaGrafico(Grupos,Valores):
    #Grafico de barras verticais
    #grupos = ['Produto A', 'Produto B', 'Produto C']
    #valores = [1, 10, 100]
    plt.bar(Grupos, Valores)
    plt.show()

    #Pizza
    #vendas = [3000, 2300, 1000, 500]
    #labels = ['E-commerce', 'Loja Física', 'e-mail', 'Marketplace']
    # define o nível de separabilidade entre as partes, ordem do vetor representa as partes
    #explode = (0.1, 0, 0, 0)
    # define o formato de visualização com saída em 1.1%%, sombras e a separação entre as partes
    #plt.pie(vendas, labels=labels, autopct='%1.1f%%', shadow=True, explode=explode)
    # inseri a legenda e a localização da legenda.
    #plt.legend(labels, loc=3)
    # define que o gráfico será plotado em circulo
    #plt.axis('equal')
    #plt.show()

def LeDados():
    cursor = BancoDeDados.ConsultaSQL("select PERIODO, sum(UST) as USTs from Demandas_BRQ where	PERIODO is not null group by PERIODO ORDER BY USTs")
    dados = cursor.fetchone()
    Grupos=[]
    Valores=[]
    while dados:
        Grupos.append(str(dados[0]).strip())
        Valores.append(str(dados[1]).strip())
        dados = cursor.fetchone()

    CriaGrafico(Grupos,Valores)

#CarregaCSV()
#AtualizaRegistros()
LeDados()
CriaGrafico()