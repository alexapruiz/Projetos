VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Recepcao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private m_Connection        As rdo.rdoConnection

Public Function SetConnection(ByRef pConnection As rdo.rdoConnection) As Boolean

    On Error GoTo Erro_Connection:
    
    SetConnection = False

    Set Geral.Banco = pConnection
    
    SetConnection = True
    
    Exit Function
    
Erro_Connection:
    
    

End Function

Public Function SetConnectionAgencia(ByRef pConnection As rdo.rdoConnection) As Boolean

    On Error GoTo Erro_Connection:
    
    SetConnectionAgencia = False

    Set Geral.BancoCaixa = pConnection
    
    SetConnectionAgencia = True
    
    Exit Function
    
Erro_Connection:

End Function


Public Function SetAgenciaApresentante(ByVal pAgenciaApresentante As String) As Boolean

    On Error GoTo Erro_AgenciaApresentante:
    
    SetAgenciaApresentante = False

    Geral.AgenciaApresentante = pAgenciaApresentante
    
    SetAgenciaApresentante = True
    
Erro_AgenciaApresentante:
    
    

End Function

Public Function SetDataProcessamento(ByRef pDataProcessamento As Long) As Boolean

    On Error GoTo Erro_SetDataProcessamento:
    
    SetDataProcessamento = False

    Geral.DataProcessamento = pDataProcessamento
    
    SetDataProcessamento = True
    
Erro_SetDataProcessamento:
    
    

End Function


Public Function SetUsuario(ByVal pUsuarioLogin As String) As Boolean

    On Error GoTo Erro_Usuario
    
    SetUsuario = False
    
    Geral.Usuario.Login = pUsuarioLogin
    
    SetUsuario = True
    
Erro_Usuario:

End Function

Public Sub showModal()

    If Geral.Banco Is Nothing Then
        MsgBox "Conexão não definida.", vbExclamation
        Exit Sub
    End If
    
    frmRecepcao.Show vbModal
End Sub


