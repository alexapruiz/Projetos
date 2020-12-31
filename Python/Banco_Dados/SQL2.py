import sqlite3

#Cria a conexão com o banco de dados
conn = sqlite3.connect('C:/Install/Ferramentas Data Science/SQLiteStudio/db1')

#Cria um cursor, para executar comandos no banco de dados
cur = conn.cursor()

#cur.execute("insert into CLIENTE2 (ID_CLIENTE, NOME_CLIENTE) values (7,'Lili')")
#conn.commit()

cur.execute("SELECT * from CLIENTE")
for linha in cur:
    print(linha)

cur.close()
conn.close()