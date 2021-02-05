import os
import sys

sys.path.append('c:\\Projetos\\Python\\Banco_Dados')
from BancodeDados import SQLServer

DB_CAIXA = SQLServer('CAIXA')

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