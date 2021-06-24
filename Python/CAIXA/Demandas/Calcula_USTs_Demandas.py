import csv
import sys

sys.path.append('c:\\Projetos\\Python\\Banco_Dados')


servico_baixa = {'Apoio à Solução de Problemas Relacionados às Ferramentas': 4,
                 'Criação/Manutenção de área de projeto': 2,
                 'Mentoring': 2,
                 'Manutenção em permissões/perfis de usuário':1,
                 'Criação de área de projeto integrada/associação de áreas de projeto':1,
                 'Manutenção em Indicador / Relatório':8,
                 'Manutenção em Itens de Configuração (ex:Retirada de check-out)':2,
                 'Manutenção em painéis':2,
                 'Criação e Configuração de Projeto':4,
                 'Criação / Manutenção de Atributos':1,
                 'G1 - Capacitação':0,
                 'Criação / Manutenção de View':2,
                 'Criação / Manutenção de VOB':8,
                 'Criação / Manutenção de Artefatos':2,
                 'Criação / Manutenção de Trigger':16,
                 'Customização de modelos de script':4,
                 'Importação de itens de projeto': 4,
                 'Manutenção em Label': 4,
                 'Não Informado': 0
                 }


servico_media = {'Apoio à Solução de Problemas Relacionados às Ferramentas': 8,
                 'Criação/Manutenção de área de projeto': 4,
                 'Mentoring': 4,
                 'Manutenção em permissões/perfis de usuário': 2,
                 'Criação de área de projeto integrada/associação de áreas de projeto': 2,
                 'Manutenção em Indicador / Relatório': 16,
                 'Manutenção em Itens de Configuração (ex:Retirada de check-out)': 2,
                 'Manutenção em painéis': 4,
                 'Criação e Configuração de Projeto': 4,
                 'Criação / Manutenção de Atributos': 2,
                 'G1 - Capacitação': 0,
                 'Criação / Manutenção de View': 2,
                 'Criação / Manutenção de VOB': 8,
                 'Criação / Manutenção de Artefatos': 4,
                 'Criação / Manutenção de Trigger': 16,
                 'Customização de modelos de script': 4,
                 'Importação de itens de projeto': 4,
                 'Manutenção em Label': 4,
                 'Não Informado': 0
                 }

servico_alta = {'Apoio à Solução de Problemas Relacionados às Ferramentas': 12,
                 'Criação/Manutenção de área de projeto': 8,
                 'Mentoring': 8,
                 'Manutenção em permissões/perfis de usuário': 4,
                 'Criação de área de projeto integrada/associação de áreas de projeto': 4,
                 'Manutenção em Indicador / Relatório': 32,
                 'Manutenção em Itens de Configuração (ex:Retirada de check-out)': 2,
                 'Manutenção em painéis': 8,
                 'Criação e Configuração de Projeto': 4,
                 'Criação / Manutenção de Atributos': 4,
                 'G1 - Capacitação': 0,
                 'Criação / Manutenção de View': 2,
                 'Criação / Manutenção de VOB': 8,
                 'Criação / Manutenção de Artefatos': 8,
                 'Criação / Manutenção de Trigger': 16,
                 'Customização de modelos de script': 4,
                 'Importação de itens de projeto': 4,
                 'Manutenção em Label': 4,
                 'Não Informado': 0
                 }


def Calcula_USTs(SERVICO, COMPLEXIDADE, QUANTIDADE):
    #De acordo com a complexidade, buscar a quantidade de UST no dicionário correspondente
    if (COMPLEXIDADE == 'Baixa'):
        return int(servico_baixa[SERVICO]) * int(QUANTIDADE)
    elif (COMPLEXIDADE == 'Média'):
        return int(servico_media[SERVICO]) * int(QUANTIDADE)
    else:
        return int(servico_alta[SERVICO]) * int(QUANTIDADE)


#Abre a planilha original com as demandas do período
planilha = csv.DictReader(open("Demandas_2021_06.csv", encoding='utf-8'), delimiter=';')

#Prepara o arquivo de saida, com as USTs calculadas
arquivo_saida = open("Demandas_2021_06_saida.csv", "w", encoding="ANSI")

#Escreve a primeira linha, com o cabeçalho
linha_saida="ID;Resumo;Status;Quantidade;Complexidade;Data de Criação;Prazo Final;Serviço;UST;Grupo SIGCT" + "\n"
arquivo_saida.write(linha_saida)

for linha in planilha:
    linha_saida = ""
    linha_saida += str(linha["ID"]) + ";"
    linha_saida += str(linha["Resumo"])   + ";"
    linha_saida += str(linha["Status"]) + ";"
    linha_saida += str(linha["Quantidade"]) + ";"
    linha_saida += str(linha["Complexidade"]) + ";"
    linha_saida += str(linha["Data de Criação"]) + ";"
    linha_saida += str(linha["Prazo Final"]) + ";"
    linha_saida += str(linha["Serviço Especializado de Apoio a Ferramentas Rational"]) + ";"

    servico = linha["Serviço Especializado de Apoio a Ferramentas Rational"];
    complexidade = str(linha["Complexidade"]);
    qtde= str(linha["Quantidade"]);

    linha_saida += str(Calcula_USTs(servico,complexidade,qtde)) + ";"
    linha_saida += str(linha["Grupo SIGCT"])
    linha_saida += "\n"
    arquivo_saida.write(linha_saida)

arquivo_saida.close