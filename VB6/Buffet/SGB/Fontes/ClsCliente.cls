VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsCliente"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function Incluir(ByVal NomeCliente As String, ByVal Endereco As String, ByVal CPF As String, ByVal Numero As String, ByVal CEP As String, ByVal Tel1 As String, ByVal Tel2 As String, ByVal Tel3 As String, ByVal Cidade As String, ByVal Bairro As String, ByVal RG As String, ByVal Db As adodb.Connection) As Integer

    Dim sSql As String
    Dim Rec As New adodb.Recordset
    Dim x As Long
    
    'Recuperar o último código de cliente cadastrado
    sSql = "select max(cod_cli) as CLIENTE from clientes"

    Rec.Open sSql, Db, adOpenDynamic, adLockReadOnly

    If Not Rec.EOF Then
        CodCliente = Rec("CLIENTE").Value
        CodCliente = Val(CodCliente & "") + 1
        Incluir = True
    Else
        CodCliente = 1
        Incluir = False
    End If
    Set Rec = Nothing

    sSql = "insert into CLIENTES (COD_CLI, NOM_CLI,CPF_CLI,END_CLI,NUM_CLI,CEP_CLI,TEL1_CLI,TEL2_CLI,TEL3_CLI,BAI_CLI,CID_CLI,RG_CLI) VALUES ("
    sSql = sSql & CodCliente & ",'"
    sSql = sSql & NomeCliente & "','"
    sSql = sSql & CPF & "','"
    sSql = sSql & Endereco & "',"
    sSql = sSql & Val(Numero) & ",'"
    sSql = sSql & CEP & "','"
    sSql = sSql & Tel1 & "','"
    sSql = sSql & Tel2 & "','"
    sSql = sSql & Tel3 & "','"
    sSql = sSql & Bairro & "','"
    sSql = sSql & Cidade & "','"
    sSql = sSql & RG

    sSql = sSql & "')"

    Rec.Open sSql, Db, adOpenDynamic, adLockReadOnly
    
    Incluir = CodCliente
End Function
Public Function Consultar(ByVal CodigoCliente As String, ByVal Db As adodb.Connection) As adodb.Recordset

    Dim sSql As String
    Dim Rec As New adodb.Recordset
    Dim x As Long
    
    'Recuperar o último código de cliente cadastrado
    sSql = "select * from clientes "
    sSql = sSql & " where COD_CLI = " & CodigoCliente

    Rec.Open sSql, Db, adOpenDynamic, adLockReadOnly

    If Not Rec.EOF Then
        Set Consultar = Rec
    End If
End Function

