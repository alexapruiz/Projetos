VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsParcela"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub Incluir(ByVal ID_CNT As Integer, ByVal ID_PAR As Integer, ByVal DATA_PAR As String, ByVal VALOR_PAR As Currency, ByVal NUM_DOC_PAR As String, ByVal FORMA_PAGTO As String, ByVal Db As ADODB.Connection)

    Dim sSql As String
    Dim Rs As New ADODB.Recordset

    sSql = "INSERT INTO PARCELA_CONTRATO (ID_CNT,ID_PAR,DATA_PAR,ID_FORMA,NUM_DOC_PAR,VALOR_PAR) VALUES ("
    sSql = sSql & ID_CNT & ","
    sSql = sSql & ID_PAR & ",'"
    sSql = sSql & DATA_PAR & "',"
    sSql = sSql & FORMA_PAGTO & ",'"
    sSql = sSql & NUM_DOC_PAR & "',"
    sSql = sSql & ConverteMoeda(VALOR_PAR) & ")"

    Rs.Open sSql, Db, adOpenDynamic, adLockOptimistic
End Sub
Public Sub Excluir(ByVal ID_CNT As Integer, ByVal Db As ADODB.Connection)

    Dim sSql As String
    Dim Rs As New ADODB.Recordset

    sSql = "DELETE FROM PARCELA_CONTRATO WHERE"
    sSql = sSql & " ID_CNT = " & ID_CNT

    Rs.Open sSql, Db, adOpenDynamic, adLockOptimistic
End Sub
