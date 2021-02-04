import sys
import sqlite3
sys.path.append('c:\\Projetos\\Python\\Banco_Dados')
from BancodeDados import Banco_SQLite

if not (Banco_SQLite.ExecutaComandoSQL('Banco1.db',"update Clientes set Nome='Vitor' where Codigo = 1")):
    print('Erro')

if not (Banco_SQLite.ExecutaComandoSQL('Banco1.db',"update Clientes set Nome='Alex' where Codigo = 2")):
    print('Erro')

cursor_clientes = Banco_SQLite.ConsultaSQL('Banco1.db',"select * from Clientes")
for rec in cursor_clientes:
    print(rec)