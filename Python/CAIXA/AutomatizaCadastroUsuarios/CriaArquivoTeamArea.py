import os
import sys

sys.path.append('c:\\Projetos\\Python\\Banco_Dados')
from BancodeDados import SQLServer

DB_CAIXA = SQLServer('CAIXA')

def define_papel(argument):
    switcher = {
        "Arquitect Owner": "architect_owner",
        "Agente de Qualidade": "agente_qualidade",
        "Administrador": "Administrador",
        "Dono do Produto (Product Owner - PO)": "dono_produto",
        "Facilitador Ágil": "facilitador_agil",
        "Líder do Time (Squad Leader)": "lider_time",
        "Líder de Negócio": "lider_negocio",
        "Líder de TI": "lider_ti",
        "Líder de Solução": "lider_solucao",
        "Líder Técnico": "lider_tecnico",
        "Líder de Negócio (Business Owner - BO)": "lider_negocio",
        "Líder Ágil": "lider_agil",
        "Tester": "tester",
        "Time (Desenvolvedores)": "time",
        "UX Designer": "ux_designer"
    }
    return switcher.get(argument, "Papel não encontrado")


def CarregaPlanilhaparaBanco():
    import pandas as pd
    import sys
    from openpyxl import load_workbook

    sys.path.append('c:\\Projetos\\Python\\Banco_Dados')
    from BancodeDados import SQLServer

    DB_CAIXA = SQLServer('CAIXA')

    #Importa os dados da planilha
    arquivo = 'C:\\Projetos\\_Arquivos\\CAIXA\\Usuarios_SQUADS_05.03.2021.xlsx'

    #Limpa a tabela destino
    DB_CAIXA.ExecutaComandoSQL('Truncate Table Comunidade_Usuarios')

    #Abrir a planilha e gerar a tabela
    wb = load_workbook(arquivo)
    ws = wb['Planilha1']
    linha = 1
    for line in ws:
        if linha > 1:
            #Transformar a coluna papel, colocando apenas o ID
            ComandoSQL = "insert into Comunidade_Usuarios (COMUNIDADE, SQUAD, MATRICULA, PAPEL) values ('"
            ComandoSQL += str(line[0].value) + "','"
            ComandoSQL += str(line[1].value) + "','"
            ComandoSQL += str(line[2].value) + "','"
            ComandoSQL += str(define_papel(line[3].value)) + "')"
            DB_CAIXA.ExecutaComandoSQL(ComandoSQL)
        linha += 1
    print('Dados Importados com Sucesso!!!')


def CriaArquivosUsuariosComunidades():
    #SELECIONA AS COMUNIDADES
    cursor_comunidades = DB_CAIXA.ConsultaSQL('exec COMUNIDADE_SEL_DISTINCT')
    comunidades = cursor_comunidades.fetchone()
    while comunidades:
        # PARA CADA COMUNIDADE ENCONTRADA, PESQUISA OS SQUADS
        cursor_squads = DB_CAIXA.ConsultaSQL("exec COMUNIDADE_SEL_DISTINCT '" + str(comunidades[0] + "'"))
        squads = cursor_squads.fetchone()
        count_squad = 1
        while squads:
            # PARA CADA SQUAD, PESQUISA OS PAPEIS
            cursor_papeis = DB_CAIXA.ConsultaSQL("COMUNIDADE_SEL_DISTINCT '" + str(comunidades[0]) + "' , '" + str(squads[0] + "'"))
            papeis = cursor_papeis.fetchone()
            while papeis:
                cursor_matriculas = DB_CAIXA.ConsultaSQL("COMUNIDADE_SEL_DISTINCT '" + str(comunidades[0]) + "' , '" + str(squads[0] + "' , '" + str(papeis[0]) + "'"))
                matriculas = cursor_matriculas.fetchone()
                arquivo_matriculas = ''
                while matriculas:
                    arquivo_matriculas = arquivo_matriculas + matriculas[0] + ','
                    matriculas = cursor_matriculas.fetchone()

                nome_arquivo_saida = comunidades[0] + '_SQUAD' + str(count_squad) + '_' + papeis[0] + '.txt'
                arquivo_saida = open(os.getcwd() + '\\saida\\' + nome_arquivo_saida, "w", encoding="ANSI")
                arquivo_saida.write('rtc.usuarios=' + arquivo_matriculas[:-1] + '\n')
                arquivo_saida.write('rtc.roles=' + papeis[0]+ '\n')
                arquivo_saida.write('rtc.teamAreaNome=' + squads[0]+ '\n')
                arquivo_saida.write('rtc.arquivo_areas=' + comunidades[0]+ '\n')
                arquivo_saida.close()

                papeis = cursor_papeis.fetchone()

            count_squad += 1
            squads = cursor_squads.fetchone()

        comunidades = cursor_comunidades.fetchone()


CarregaPlanilhaparaBanco()
CriaArquivosUsuariosComunidades()
print('Arquivos criados com Sucesso!!!!')