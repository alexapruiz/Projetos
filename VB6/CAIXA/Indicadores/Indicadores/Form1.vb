Public Class Form1

    Inherits System.Windows.Forms.Form

    'Public myCommand As New MySqlCommand
    'Public myAdapter As New MySqlDataAdapter
    'Public myData As New DataTable
    'Friend WithEvents CmdLimpaCampos As System.Windows.Forms.Button
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        'Atualiza campo COORDENACAO para a tabela DEMANDAS
        Call AtualizaCoordenacao("DEMANDAS")

        'Atualiza campo COORDENACAO para a tabela BACKLOG
        Call AtualizaCoordenacao("BACKLOG")

        MsgBox("Preparação executada com sucesso. Verifique o campo 'COORDENACAO' na base de dados", MsgBoxStyle.Information, AcceptButton)
    End Sub
    Private Sub AtualizaCoordenacao(ByVal TABELA As String)

        Dim conn As New ADODB.Connection
        Dim sSql As String

        'Prepara string para conexão com o banco de dados
        'conn.ConnectionString = "driver=MySQL ODBC 5.1 Driver;database=indicadores;uid=root;pwd=bia1701;"
        conn.ConnectionString = "driver={Microsoft Access Driver (*.mdb)};dbq=c:\alex\Indicadores.mdb"
        conn.Open()

        'Atualizando a COORDENAÇÃO - REDEASP05
        sSql = "UPDATE " & TABELA & " set COORDENACAO = 'REDEASP05'"
        sSql = sSql & " where supervisor IN ('REDEASP14','REDEASP63','REDEASP66', 'REDEASP77')"
        conn.Execute(sSql)

        sSql = "UPDATE " & TABELA & " set COORDENACAO = 'REDEASP05'"
        sSql = sSql & " where SUPERVISOR IN ('REDEASP69','REDEASP14','c296582')"
        conn.Execute(sSql)

        'Atualizando a COORDENAÇÃO - REDEASP53
        sSql = "UPDATE " & TABELA & " set COORDENACAO = 'REDEASP53'"
        sSql = sSql & " where(LiderProjeto = 'REDEASP28' Or Supervisor = 'REDEASP28' Or 'REDEASP89' Or Supervisor = 'REDEASP89')"
        conn.Execute(sSql)

        'Atualizando a COORDENAÇÃO - REDEASP59
        sSql = "UPDATE " & TABELA & " set COORDENACAO = 'REDEASP59'"
        sSql = sSql & " where supervisor IN ('REDEASP04','REDEASP43','REDEASP59','REDEASP70')"
        conn.Execute(sSql)

        sSql = "UPDATE " & TABELA & " set COORDENACAO = 'REDEASP59'"
        sSql = sSql & " where liderprojeto IN ('REDEASP23','REDEASP76','REDEASP57','REDEASP41','REDEASP39')"
        conn.Execute(sSql)

        sSql = "UPDATE " & TABELA & " set COORDENACAO = 'REDEASP59'"
        sSql = sSql & " where supervisor IN ('REDEASP23','REDEASP76','REDEASP57','REDEASP41','REDEASP39')"
        conn.Execute(sSql)

        'Atualizando a COORDENAÇÃO - REDEASP61
        sSql = "UPDATE " & TABELA & " set COORDENACAO = 'REDEASP61'"
        sSql = sSql & " where supervisor IN ('REDEASP03','REDEASP10','REDEASP12','REDEASP16','REDEASP64')"
        conn.Execute(sSql)

        sSql = "UPDATE " & TABELA & " set COORDENACAO = 'REDEASP61'"
        sSql = sSql & " where liderprojeto IN ('REDEASP73','REDEASP68')"
        conn.Execute(sSql)

        sSql = "UPDATE " & TABELA & " set COORDENACAO = 'REDEASP61'"
        sSql = sSql & " where supervisor IN ('REDEASP73','REDEASP68')"
        conn.Execute(sSql)

        'Atualizando a COORDENAÇÃO - REDEASP87
        sSql = "UPDATE " & TABELA & " set COORDENACAO = 'REDEASP87'"
        sSql = sSql & " where supervisor IN ('REDEASP25','REDEASP56')"
        conn.Execute(sSql)

        'Atualizando a COORDENAÇÃO - REDEASP93
        sSql = "UPDATE " & TABELA & " set COORDENACAO = 'REDEASP93'"
        sSql = sSql & " where supervisor IN ('REDEASP06','REDEASP54','REDEASP71','REDEASP21','c300518')"
        sSql = sSql & " and COORDENACAO <> 'REDEASP53'"
        conn.Execute(sSql)

        sSql = "UPDATE " & TABELA & " set COORDENACAO = 'REDEASP93'"
        sSql = sSql & " where liderprojeto IN ('REDEASP06','REDEASP54','REDEASP71','REDEASP21','c300518')"
        sSql = sSql & " and COORDENACAO <> 'REDEASP53'"
        conn.Execute(sSql)

        'Fechando a conexão
        conn.Close()
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Call ObtemIndicador1_1("DEMANDAS", "01/26/2010", "02/28/2010")

    End Sub
    Private Sub ObtemIndicador1_1(ByVal TABELA As String, ByVal Data1 As String, ByVal Data2 As String)

        Dim conn As New ADODB.Connection
        Dim Rec As New ADODB.Recordset
        Dim sSql As String

        'Prepara string para conexão com o banco de dados
        conn.ConnectionString = "driver={Microsoft Access Driver (*.mdb)};dbq=c:\alex\projetos\Indicadores.mdb"
        conn.Open()

        sSql = "SELECT Count(0) AS TOTAL_SIGTI, COORDENACAO "
        sSql = sSql & " FROM " & TABELA
        sSql = sSql & " WHERE DataSolicitacao Between #" & Data1 & "# And #" & data2 & "#"
        sSql = sSql & " AND DEMANDASIGTI IS NOT NULL"
        sSql = sSql & " GROUP BY DEMANDAS.coordenacao"
        sSql = sSql & " ORDER BY DEMANDAS.coordenacao;"

        Rec.Open(sSql, conn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)

    End Sub
End Class
