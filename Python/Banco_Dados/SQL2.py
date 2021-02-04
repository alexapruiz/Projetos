import sys
import sqlite3
sys.path.append('c:\\Projetos\\Python\\Banco_Dados')
from BancodeDados import Banco_SQLite

cursor_clientes = Banco_SQLite.ConsultaSQL('C:\\Projetos\\Python\\Banco_Dados\\Banco1.db','select * from Clientes')
#clientes = cursor_clientes.fetchall()
cursor_clientes.fetchall()
for rec in cursor_clientes:
    print(rec)