import sqlite3

# Cria uma conexão e um cursor
con = sqlite3.connect('emails2.db')
cur = con.cursor()

#sql = 'create table emails (id integer primary key, nome varchar(100), email varchar(100))'
#cur.execute(sql)

sql = 'insert into emails values (null, ?, ?)'

recset = [('jane doe', 'jane@nowhere.org'),
 ('rock', 'rock@hardplace.com')]

# Insere os registros
#for rec in recset:
#    cur.execute(sql, rec)
#    # Confirma a transação
#    con.commit()

# Seleciona todos os registros
cur.execute('select * from emails')

# Recupera os resultados
recset = cur.fetchall()

# Mostra os registros da tabela emails
for rec in recset:
    print('%d: %s(%s)' % rec)
    print(rec)