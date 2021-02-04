import os
import sys

sys.path.append('c:\\Projetos\\Python\\Banco_Dados')
from SQL_Server import SQLServer

#SELECIONA AS COMUNIDADES
cursor_comunidades = SQLServer.ConsultaSQL("SELECT DISTINCT(COMUNIDADE) FROM Comunidade_Usuarios ORDER BY COMUNIDADE")
comunidades = cursor_comunidades.fetchone()
while comunidades:
    # PARA CADA COMUNIDADE ENCONTRADA, PESQUISA OS SQUADS
    cursor_squads = SQLServer.ConsultaSQL("SELECT DISTINCT(SQUAD) as SQUAD FROM Comunidade_Usuarios WHERE COMUNIDADE = '" + str(comunidades[0]) + "' ORDER BY SQUAD")
    squads = cursor_squads.fetchone()
    count_squad = 1
    while squads:
        # PARA CADA SQUAD, PESQUISA OS PAPEIS
        cursor_papeis = SQLServer.ConsultaSQL("SELECT DISTINCT(PAPEL) as PAPEL FROM Comunidade_Usuarios WHERE COMUNIDADE = '" + str(comunidades[0]) + "' AND SQUAD = '" + str(squads[0] + "' ORDER BY PAPEL"))
        papeis = cursor_papeis.fetchone()
        while papeis:
            cursor_matriculas = SQLServer.ConsultaSQL("SELECT MATRICULA FROM Comunidade_Usuarios WHERE COMUNIDADE = '" + str(comunidades[0]) + "' AND SQUAD = '" + str(squads[0] + "' AND PAPEL = '" + str(papeis[0]) + "' ORDER BY MATRICULA"))
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