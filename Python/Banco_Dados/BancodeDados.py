class SQLServer():
    import pyodbc

    StringConexao = "DRIVER=SQL Server;SERVER=NOVO\SQLEXPRESS;PORT=1433;DATABASE=CAIXA;Trustedconnection=yes"

    def ConsultaSQL(ComandoSQL) -> object:
        conn = pyodbc.connect(SQLServer.StringConexao)
        cursor = conn.cursor()
        cursor.execute(ComandoSQL)
        return cursor

    def ExecutaComandoSQL(ComandoSQL):
        conn = pyodbc.connect(BancoDeDados.StringConexao)
        cursor = conn.cursor()
        cursor.execute(ComandoSQL)
        conn.commit()
        return "Comando Executado com Sucesso!!!"

class Banco_SQLite():

    def ConsultaSQL(arquivo, ComandoSQL) -> object:
        import sqlite3

        conexao = sqlite3.connect(arquivo)
        cursor = conexao.cursor()
        cursor.execute(ComandoSQL)
        return cursor

    def ExecutaComandoSQL(arquivo, ComandoSQL) -> object:
        import sqlite3

        conexao = sqlite3.connect(arquivo)
        conexao.execute(ComandoSQL)
        conexao.commit()
        conexao.close()
        return True