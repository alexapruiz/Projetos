VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDespesas"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function RegistrarDespesaPaga(ByVal ID_FECHA As Integer, ByVal ID_DESP As Integer, ByVal Db As ADODB.Connection) As Integer

    On Error GoTo RegistrarDespesaPaga_Erro

    Dim sSql As String
    Dim Rec As New ADODB.Recordset

    sSql = "INSERT INTO PAGTO_DESP_FIXA (ID_FECHA, ID_DESP, DATA_PAGTO) VALUES "
    sSql = sSql & " ( " & ID_FECHA & " , " & ID_DESP & " , " & Format(Now, "dd/mm/yyyy") & ")"
    
    Rec.Open sSql, Db, adOpenDynamic, adLockReadOnly

    Exit Function

RegistrarDespesaPaga_Erro:
    MsgBox "Erro ao registrar o Pagamento de Despesa Fixa", vbExclamation, "SGB"
End Function
