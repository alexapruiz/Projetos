VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsContasaPagar"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function Incluir(ByVal ID_CNT As Integer, ByVal ID_FOR As Integer, ByVal ID_PRD As Integer, ByVal QTDE As Integer, ByVal VALOR As Currency) As Integer

    Dim Rs As New ADODB.Recordset
    Dim sSql As String

    On Error GoTo ContasAPagar_Incluir_Erro

    sSql = "INSERT INTO CONTAS_A_PAGAR (ID_CNT,ID_FOR,ID_PRD,QTDE,VALOR) VALUES ("
    sSql = sSql & ID_CNT & ","
    sSql = sSql & ID_FOR & ","
    sSql = sSql & ID_PRD & ","
    sSql = sSql & QTDE & ","
    sSql = sSql & CInt(VALOR) & ")"
    
    Rs.Open sSql, Db, adOpenDynamic, adLockOptimistic

    Incluir = True

    Exit Function

ContasAPagar_Incluir_Erro:
    MsgBox "Erro ao Incluir os dados de Contas a Pagar", vbExclamation, "SGB"
    Incluir = False
End Function
Public Function Excluir(ByVal ID_CNT As Integer, ByVal Db As ADODB.Connection) As Integer

    Dim Rs As New ADODB.Recordset
    Dim sSql As String

    On Error GoTo ContasAPagar_Excluir_Erro

    sSql = "DELETE FROM CONTAS_A_PAGAR WHERE ID_CNT = " & ID_CNT

    Rs.Open sSql, Db, adOpenDynamic, adLockOptimistic

    Excluir = True

    Exit Function

ContasAPagar_Excluir_Erro:
    MsgBox "Erro ao Excluir os dados de Contas a Pagar", vbExclamation, "SGB"
    Excluir = False
End Function
Public Function IncluirEscala(ByVal ID_CNT As Integer, ByVal ID_COL As Integer, ByVal ID_FUNC As Integer) As Integer

    Dim Rs As New ADODB.Recordset
    Dim sSql As String

    On Error GoTo ContasAPagar_IncluirEscala_Erro

    sSql = "INSERT INTO ESCALA (ID_CNT,ID_COL,ID_FUNC) VALUES ("
    sSql = sSql & ID_CNT & ","
    sSql = sSql & ID_COL & ","
    sSql = sSql & ID_FUNC & ")"

    Rs.Open sSql, Db, adOpenDynamic, adLockOptimistic

    IncluirEscala = True

    Exit Function

ContasAPagar_IncluirEscala_Erro:
    MsgBox "Erro ao Incluir os dados da Escala", vbExclamation, "SGB"
    IncluirEscala = False
End Function
Public Function ExcluirEscala(ByVal ID_CNT As Integer) As Integer

    Dim Rs As New ADODB.Recordset
    Dim sSql As String

    On Error GoTo ContasAPagar_ExcluirEscala_Erro

    sSql = "DELETE FROM ESCALA WHERE ID_CNT = " & ID_CNT

    Rs.Open sSql, Db, adOpenDynamic, adLockOptimistic

    ExcluirEscala = True

    Exit Function

ContasAPagar_ExcluirEscala_Erro:
    MsgBox "Erro ao Excluir os dados de Contas a Pagar", vbExclamation, "SGB"
    ExcluirEscala = False
End Function
Public Function InserirFechamento(ByVal ID_FECHA As Integer, DATA_INI As String, DATA_FIM As String, ByVal INC_DESP As Integer) As Integer

    Dim Rs As New ADODB.Recordset
    Dim sSql As String

    On Error GoTo InserirFechamento_Erro

    sSql = "INSERT INTO FECHAMENTO (ID_FECHA,DATA_INI_VIG,DATA_FIM_VIG,INC_DESP) VALUES ("
    sSql = sSql & ID_FECHA & ",'"
    sSql = sSql & DATA_INI & "','"
    sSql = sSql & DATA_FIM & "','"
    sSql = sSql & INC_DESP & "')"

    Rs.Open sSql, Db, adOpenDynamic, adLockOptimistic

    InserirFechamento = True

    Exit Function

InserirFechamento_Erro:
    MsgBox "Erro ao Inserir Registro de Fechamento", vbExclamation, "SGB"
    InserirFechamento = False
End Function
Public Sub RegistrarPagtoFornecedores(DATA_INI As String, DATA_FIM As String, ID_FOR As String, Db As ADODB.Connection)

    Dim Rs As New ADODB.Recordset
    Dim sSql As String

    On Error GoTo RegistrarPagtoFornecedores_Erro

    sSql = "UPDATE CONTAS_A_PAGAR "
    sSql = sSql & " SET PAGO = 'S'"
    sSql = sSql & " WHERE ID_CNT IN ("
    sSql = sSql & " SELECT ID_CNT FROM CONTRATOS WHERE DATA_FESTA BETWEEN #" & DATA_INI
    sSql = sSql & "# AND #" & DATA_FIM & "#)"
    sSql = sSql & " AND ID_FOR = " & ID_FOR

    Rs.Open sSql, Db, adOpenDynamic, adLockOptimistic

    Exit Sub

RegistrarPagtoFornecedores_Erro:
    MsgBox "Erro ao Inserir Registro de Fechamento", vbExclamation, "SGB"
End Sub
Public Sub RegistrarPagtoFuncionarios(DATA_INI As String, DATA_FIM As String, ID_COL As String, Db As ADODB.Connection)

    Dim Rs As New ADODB.Recordset
    Dim sSql As String

    On Error GoTo RegistrarPagtoFuncionarios_Erro

    sSql = "UPDATE ESCALA "
    sSql = sSql & " SET PAGO = 'S'"
    sSql = sSql & " WHERE ID_CNT IN ("
    sSql = sSql & " SELECT ID_CNT FROM CONTRATOS WHERE DATA_FESTA BETWEEN #" & DATA_INI
    sSql = sSql & "# AND #" & DATA_FIM & "#)"
    sSql = sSql & " AND ID_COL = " & ID_COL

    Rs.Open sSql, Db, adOpenDynamic, adLockOptimistic

    Exit Sub

RegistrarPagtoFuncionarios_Erro:
    MsgBox "Erro ao atualizar dados de Pagamento de Funcionários", vbExclamation, "SGB"
End Sub

