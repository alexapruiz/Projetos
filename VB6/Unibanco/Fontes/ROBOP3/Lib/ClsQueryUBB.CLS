VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsQueryUBB"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim HwndLib             As Long

Dim m_DataSourceName    As String
Dim m_DatabaseName      As String
Dim m_LibPath           As String

Public DBUBB            As New ADOUbb.clsADOoUbbCrip
Public Property Get DataSourceName() As String
    DataSourceName = m_DataSourceName
End Property
Public Property Let DataSourceName(New_DataSourceName As String)
    m_DataSourceName = New_DataSourceName
End Property
Public Property Get DatabaseName() As String
    DatabaseName = m_DatabaseName
End Property
Public Property Let DatabaseName(New_DatabaseName As String)
    m_DatabaseName = New_DatabaseName
End Property
Public Property Get LibPath() As String
    LibPath = m_LibPath
End Property
Public Property Let LibPath(New_LibPath As String)
    m_LibPath = New_LibPath
End Property
Function getControle() As Recordset

On Error GoTo TrataErro

    Dim sStr As String
        
    sStr = ""
    sStr = sStr & "SELECT * FROM tfsctrla"
    
    If Not DBUBB.OpenRecordSetUBB(sStr, UbbStatic, getControle) Then
        GoTo TrataErro
    End If
    
    Exit Function
    
TrataErro:
  
    Err.Raise Err.Number, "Classe UBB - Query: getControle", DBUBB.GetLastErrorMsgUBB
    
End Function
Function getTabela(pTabela As String) As Recordset
On Error GoTo TrataErro

    Dim sStr As String
        
    sStr = "SELECT * FROM " & pTabela
    DBUBB.OpenRecordSetUBB sStr, UbbStatic, getTabela
    
    Exit Function
    
TrataErro:
  
    Err.Raise Err.Number, "Classe UBB - Query: getTabela", DBUBB.GetLastErrorMsgUBB
    
End Function
Function getPrzVenctoVrAlcada(AgenciaCentral As String) As Recordset
On Error GoTo TrataErro

    Dim sStr As String
        
    sStr = "SELECT agefsnudivc, agefsvllisu FROM tfsageng "
    sStr = sStr & " Where agefscdagbk = " & AgenciaCentral

    DBUBB.OpenRecordSetUBB sStr, UbbStatic, getPrzVenctoVrAlcada
    
    Exit Function
    
TrataErro:
  
    Err.Raise Err.Number, "Classe UBB - Query: getPrzVenctoVrAlcada", DBUBB.GetLastErrorMsgUBB
    
End Function
Function getTabelaAgeng() As Recordset

On Error GoTo TrataErro

    Dim sStr As String
        
    sStr = "SELECT "
    sStr = sStr & "Distinct (agefsnoagen), "
    sStr = sStr & "agefsestado, "
    sStr = sStr & "agefscdagen, "
    sStr = sStr & "agefsstmovi, "
    sStr = sStr & "agefsdtmvan, "
    sStr = sStr & "agefsdtmvat, "
    sStr = sStr & "agefsdtprox "
    sStr = sStr & "From tfsAgeng "

    DBUBB.OpenRecordSetUBB sStr, UbbStatic, getTabelaAgeng
    
    Exit Function
    
TrataErro:
  
    Err.Raise Err.Number, "Classe UBB - Query: getTabelaAgeng", DBUBB.GetLastErrorMsgUBB

End Function
Function getNsuTabelacaixa(pCaixa As String) As Recordset

On Error GoTo TrataErro

    Dim sStr As String
        
    sStr = "Select * From tfstcxag "
    sStr = sStr & " where tcxfsnucaix = " & pCaixa

    DBUBB.OpenRecordSetUBB sStr, UbbStatic, getNsuTabelacaixa
    
    Exit Function
    
TrataErro:
  
    Err.Raise Err.Number, "Classe UBB - Query: getNsuTabelacaixa", DBUBB.GetLastErrorMsgUBB

End Function
Function getAgenciaConta(ByVal pAgencia As String, ByVal pConta As String, Optional ByVal pTipoConta As String = "T") As Recordset
    On Error GoTo TrataErro

    Dim sStr As String
        
    sStr = "SELECT * FROM tfsconta "
    sStr = sStr & " where ctafscdagen = " & pAgencia
    sStr = sStr & " and ctafsnumcta = " & pConta
    
    If pTipoConta = "P" Or pTipoConta = "C" Then
        sStr = sStr & " and ctafstipcta = '" & pTipoConta & "'"
    End If

    DBUBB.OpenRecordSetUBB sStr, UbbStatic, getAgenciaConta
    
    Exit Function
    
TrataErro:
  
    Err.Raise Err.Number, "Classe UBB- Query: getAgenciaConta", DBUBB.GetLastErrorMsgUBB
    
End Function
Function ExecuteSQL(sStr As String) As Recordset

    On Error GoTo TrataErro

    If Not DBUBB.OpenRecordSetUBB(sStr, UbbStatic, ExecuteSQL) Then
        GoTo TrataErro
    End If
    
    Exit Function
    
TrataErro:
  
    Err.Raise Err.Number, "Classe UBB - Função: ExecuteSQL", DBUBB.GetLastErrorMsgUBB
    
End Function
Function getPracaCompensacao(pAgOrig As Integer) As Recordset

On Error GoTo TrataErro

    Dim sStr As String
        
    sStr = "SELECT * FROM tfsageng"
    sStr = sStr & " where agefscdagen = " & pAgOrig
   
    DBUBB.OpenRecordSetUBB sStr, UbbStatic, getPracaCompensacao
    
    Exit Function
    
TrataErro:
  
    Err.Raise Err.Number, "Classe UBB - Query: getPracaCompensacao", DBUBB.GetLastErrorMsgUBB

End Function
Public Sub SetConnection()

On Error GoTo TrataErro

    DBUBB.DatabaseName = Me.DatabaseName
    DBUBB.DataSourceName = Me.DataSourceName
    
    If Not DBUBB.OpenDataBaseUBB Then
        GoTo TrataErro
    End If
    
    Exit Sub
    
TrataErro:

    Err.Raise 998, "Abertura de Conexão - UBB", DBUBB.GetLastErrorMsgUBB
       
End Sub
Private Sub Class_Initialize()
    HwndLib = LoadLibrary(App.Comments & "\QxConnect_dll.dll")
End Sub
Private Sub Class_Terminate()
    FreeLibrary (HwndLib)
    DBUBB.CloseDataBaseUBB
End Sub
