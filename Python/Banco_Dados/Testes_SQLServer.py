import os
import sys

sys.path.append('c:\\Projetos\\Python\\Banco_Dados')
from BancodeDados import SQLServer

DB_CAIXA = SQLServer('CAIXA')
cursor_comunidades = DB_CAIXA.ConsultaSQL('exec COMUNIDADE_SEL_DISTINCT')
comunidades = cursor_comunidades.fetchone()
x = 1
while comunidades:
    print(str(x) + ' - ' + str(comunidades[0]))
    comunidades = cursor_comunidades.fetchone()
    x += 1