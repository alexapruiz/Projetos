import pyodbc
import os

conexao = pyodbc.connect('Driver={SQL Server Native Client 11.0};SERVER=NOVO\SQLEXPRESS;DATABASE=CAIXA;Trusted_Connection=yes')
conexao2 = pyodbc.connect('Driver={SQL Server Native Client 11.0};SERVER=NOVO\SQLEXPRESS;DATABASE=CAIXA;Trusted_Connection=yes')
conexao3 = pyodbc.connect('Driver={SQL Server Native Client 11.0};SERVER=NOVO\SQLEXPRESS;DATABASE=CAIXA;Trusted_Connection=yes')
conexao4 = pyodbc.connect('Driver={SQL Server Native Client 11.0};SERVER=NOVO\SQLEXPRESS;DATABASE=CAIXA;Trusted_Connection=yes')
comunidade = conexao.cursor()
squad = conexao2.cursor()
papel = conexao3.cursor()
matricula = conexao4.cursor()

#SELECIONA AS COMUNIDADES
comunidade.execute('SELECT DISTINCT(COMUNIDADE) FROM Comunidade_Usuarios ORDER BY COMUNIDADE')
for reg_comunidade in comunidade:
    # PARA CADA COMUNIDADE ENCONTRADA, PESQUISA OS SQUADS
    ComandoSQL = "SELECT DISTINCT(SQUAD) as SQUAD FROM Comunidade_Usuarios WHERE COMUNIDADE = '" + str(reg_comunidade[0]) + "' ORDER BY SQUAD"
    squad.execute(ComandoSQL)
    count_squad = 0
    for reg_squad in squad:
        # PARA CADA SQUAD, PESQUISA OS PAPEIS
        count_squad += 1
        ComandoSQL = "SELECT DISTINCT(PAPEL) as PAPEL FROM Comunidade_Usuarios WHERE COMUNIDADE = '" + str(reg_comunidade[0]) + "' AND SQUAD = '" + str(reg_squad[0] + "' ORDER BY PAPEL")
        papel.execute(ComandoSQL)
        for reg_papel in papel:
            ComandoSQL = "SELECT MATRICULA FROM Comunidade_Usuarios WHERE COMUNIDADE = '" + str(reg_comunidade[0]) + "' AND SQUAD = '" + str(reg_squad[0] + "' AND PAPEL = '" + reg_papel[0] + "' ORDER BY MATRICULA")
            matricula.execute(ComandoSQL)
            arquivo_matriculas = ''
            for reg_matricula in matricula:
                arquivo_matriculas = arquivo_matriculas + reg_matricula[0] + ','

            nome_arquivo_saida = reg_comunidade[0] + '_SQUAD' + str(count_squad) + '_' + reg_papel[0] + '.txt'
            arquivo_saida = open(os.getcwd() + '\\saida\\' + nome_arquivo_saida, "w")
            arquivo_saida.write('rtc.usuarios=' + arquivo_matriculas[:-1] + '\n')
            arquivo_saida.write('rtc.roles=' + reg_papel[0]+ '\n')
            arquivo_saida.write('rtc.teamAreaNome=' + reg_squad[0]+ '\n')
            arquivo_saida.write('rtc.arquivo_areas=' + reg_comunidade[0]+ '\n')
            arquivo_saida.close()