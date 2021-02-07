class SQLServer():

    def __init__(self,DATABASE):
        self.StringConexao = 'DRIVER=SQL Server;SERVER=NOVO\SQLEXPRESS;PORT=1433;DATABASE=' + DATABASE + ';Trustedconnection=yes'

    def ConsultaSQL(self, ComandoSQL) -> object:
        import pyodbc
        conn = pyodbc.connect(self.StringConexao)
        cursor = conn.cursor()
        cursor.execute(ComandoSQL)
        return cursor

    def ExecutaComandoSQL(self, ComandoSQL):
        import pyodbc
        conn = pyodbc.connect(self.StringConexao)
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