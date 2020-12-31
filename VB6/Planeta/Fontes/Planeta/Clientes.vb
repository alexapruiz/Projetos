Imports MySql.Data.MySqlClient

Public Class Clientes
    Public Sub LeClientes(ByVal CodigoCliente As Integer, ByRef Rec As ADODB.Recordset)

        Dim conn As New ADODB.Connection
        Dim sSql As String

        'Prepara string para conexão com o banco de dados
        conn.ConnectionString = "driver=MySQL ODBC 5.1 Driver;database=planeta;uid=root;pwd=bia1701;"
        conn.Open()

        sSql = "select * from clientes"
        If CodigoCliente <> 0 Then
            sSql = sSql & " where codigo = " & CodigoCliente
        End If
        sSql = sSql & " order by codigo"

        Rec.Open(sSql, conn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)

    End Sub
End Class
