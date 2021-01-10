Public Class Item
    Public Sub ConsultaItem(ByVal Codigo As Integer, ByVal Tipo_Item As Integer, ByRef Rec As ADODB.Recordset)

        Dim conn As New ADODB.Connection
        Dim sSql As String

        'Prepara string para conexão com o banco de dados
        conn.ConnectionString = "driver=MySQL ODBC 5.1 Driver;database=planeta;uid=root;pwd=bia1701;"
        conn.Open()

        sSql = "select * from item "
        sSql = sSql & " where tipo = " & Tipo_Item
        If Codigo <> 0 Then
            sSql = sSql & " and codigo = " & Codigo
        End If
        sSql = sSql & " order by descricao"

        Rec.Open(sSql, conn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
    End Sub
End Class
