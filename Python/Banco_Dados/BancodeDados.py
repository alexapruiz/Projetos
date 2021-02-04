import pyodbc

class SQLServer():
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
