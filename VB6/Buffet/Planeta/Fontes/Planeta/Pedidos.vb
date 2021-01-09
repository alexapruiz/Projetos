Public Class Pedidos
    Public Sub ConsultaPedido(ByVal Codigo As String, ByRef Rec As ADODB.Recordset)

        Dim conn As New ADODB.Connection
        Dim sSql As String

        'Prepara string para conexão com o banco de dados
        conn.ConnectionString = "driver=MySQL ODBC 5.1 Driver;database=planeta;uid=root;pwd=bia1701;"
        conn.Open()

        If Codigo = -1 Then
            'Retornar o ultimo registro
            sSql = "select max(codigo) from pedido"
        Else
            sSql = "select * from pedido"
            If Codigo <> 0 Then
                sSql = sSql & " where codigo = " & Codigo
            End If
            sSql = sSql & " order by codigo"
        End If

        Rec.Open(sSql, conn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        'conn.Close()
    End Sub
    Public Sub Incluir(ByVal Data_Entrega As String, ByVal Hora_Entrega As String, ByVal cliente As Integer)

        Dim conn As New ADODB.Connection
        Dim sSql As String

        'Prepara string para conexão com o banco de dados
        conn.ConnectionString = "driver=MySQL ODBC 5.1 Driver;database=planeta;uid=root;pwd=bia1701;"
        conn.Open()

        sSql = "insert into pedido (data_entrega,hora_entrega,cliente) values ('#" & Data_Entrega & "#','" _
        & Hora_Entrega & "'," & cliente & ")"

        conn.Execute(sSql)
        conn.Close()
    End Sub
    Public Sub IncluirItens(ByVal CodigoPedido As Integer, ByVal CodigoItem As Integer, ByVal qtde As Integer, ByVal tema As String)

        Dim conn As New ADODB.Connection
        Dim sSql As String

        'Prepara string para conexão com o banco de dados
        conn.ConnectionString = "driver=MySQL ODBC 5.1 Driver;database=planeta;uid=root;pwd=bia1701;"
        conn.Open()

        sSql = "insert into item_pedido (codigo_pedido,codigo_item,qtde,tema) values ("
        sSql = sSql & CodigoPedido & ","
        sSql = sSql & CodigoItem & ","
        sSql = sSql & qtde & ",'"
        sSql = sSql & tema & "')"

        conn.Execute(sSql)
        conn.Close()
    End Sub
End Class