import pyodbc

conexao = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=NOVO\SQLEXPRESS;DATABASE=CAIXA;Trusted_Connection=yes')
cursor = conexao.cursor()