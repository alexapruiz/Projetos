import sqlite3

# Cria uma conexão e um cursor
con = sqlite3.connect('c:\\Projetos\\Python\\Banco_Dados\\Banco1.db')
cur = con.cursor()

# Seleciona todos os registros
cur.execute('select * from Clientes')

# Recupera os resultados
recset = cur.fetchall()

# Mostra os registros da tabela emails
for rec in recset:
    print(rec)