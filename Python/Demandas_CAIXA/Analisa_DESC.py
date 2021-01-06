import pandas as pd

#Separando os tipos de grupos
grupos_padrao = ['CORPCAIXA\g7259csp','CORPCAIXA\g cs7266 cc_regiao_qualidade_sp' , 'CORPCAIXA\g cs7266 cc_metricas','CORPCAIXA\g cs7266 cc_regiao_ti_metricas','CORPCAIXA\g cs7266 cc_caixa_escrita','CORPCAIXA\g cs7266 cc_todos_escrita']
grupos_caixa = ['CORPCAIXA\g cs7266 cc_regiao_back_office','CORPCAIXA\g cs7266 cc_regiao_infraestrutura','CORPCAIXA\g cs7266 cc_regiao_captar','CORPCAIXA\g cs7266 cc_regiao_servicos_bancarios','CORPCAIXA\g cs7266 cc_regiao_canais','CORPCAIXA\G CS7266 CC_REGIAO_FINANCEIRO','CORPCAIXA\G CS7266 CC_REGIAO_MOBILIDADE','CORPCAIXA\g cs7266 cc_regiao_cartoes']
grupos_fabrica = ['CORPCAIXA\G CS7266 CC_REGIAO_RESOURCE','CORPCAIXA\G CS7266 CC_REGIAO_FIRST_DECISION','CORPCAIXA\G CS7266 CC_REGIAO_SPREAD','CORPCAIXA\g cs7266 cc_regiao_stefanini','CORPCAIXA\g cs7266 cc_regiao_msa_tty','CORPCAIXA\g cs7266 cc_regiao_global_web','CORPCAIXA\g cs7266 cc_regiao_foton','CORPCAIXA\G CS7266 CC_REGIAO_CAST','CORPCAIXA\G CS7266 CC_REGIAO_INDRA','CORPCAIXA\g cs7266 cc_regiao_cpm','CORPCAIXA\G CS7266 CC_REGIAO_TREE','CORPCAIXA\G CS7266 CC_REGIAO_TTY_SP','CORPCAIXA\G CS7266 CC_REGIAO_MAGNA','CORPCAIXA\g cs7266 cc_regiao_dba','CORPCAIXA\G CS7266 CC_REGIAO_MAPS']
outros_grupos = ['CORPCAIXA\g cs7266 pedep sp','CORPCAIXA\G DF7390 PRESTAR_SERVICO2','CORPCAIXA\G DF7390 FINANCIAMENTO_IMOBILIARIO','CORPCAIXA\G DF7390 NOVAS_TECNOLOGIAS','CORPCAIXA\g cs7266 sisag','CORPCAIXA\g cs7266 sisag_restrito','CORPCAIXA\g cr7265 suporte_caixa','CORPCAIXA\g cr7265 suporte_caixa','CORPCAIXA\g cs7266 sisag_restrito','CORPCAIXA\G DF7390 GESTAO_SUPORTE','CORPCAIXA\G CS7266 SISAG_RESTRITO_DIEBOLD','CORPCAIXA\G DF7390 CPM_BRAXIS']

#Abrindo os arquivos de entrada e saida
arquivo = open('c:\Projetos\_Arquivos\DESC_SP.txt')
arquivo_saida = open("saida.csv", "w")

saida = ''
GRUPO_CAIXA = ''
GRUPO_FABRICA = ''
OUTROS_GRUPOS = ''
for linha in arquivo:
    if linha.find('versioned') != -1:
        #Encontrou a primeira linha da VOB
        VOB = str(linha[24:-2])
    elif linha.find('FeatureLevel') != -1:
        #Encontrou a última linha da VOB, então precisa organizar a gravar a saída
        saida = VOB + ';' + GRUPO_CAIXA + ';' + GRUPO_FABRICA + ';' + OUTROS_GRUPOS
        arquivo_saida.write(saida)
        arquivo_saida.write('\n')
        GRUPO_CAIXA = ''
        GRUPO_FABRICA = ''
        OUTROS_GRUPOS = ''
    else:
        #Linhas do conteúdo da VOB
        if (linha.find('group') != -1) and (linha.find('Additional') != 2):
            GRUPO = str(linha[10:-1])
            if (GRUPO in grupos_caixa):
                GRUPO_CAIXA = GRUPO
            elif (GRUPO in grupos_fabrica):
                GRUPO_FABRICA = GRUPO
            elif (GRUPO in outros_grupos):
                OUTROS_GRUPOS = OUTROS_GRUPOS + ';' + GRUPO

#Fecha os arquivos
arquivo.close()
arquivo_saida.close()
print("Arquivo '" + arquivo_saida.name + "' gerado com sucesso")