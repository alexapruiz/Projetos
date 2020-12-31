VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RegistroOcorrencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Connection    As rdo.rdoConnection
Public Function SetAgenciaCentral(ByVal pAgenciaCentral As String) As Boolean

    On Error GoTo Erro_AgenciaCentral
    
    SetAgenciaCentral = False
    
    Geral.AgenciaCentral = pAgenciaCentral
    
    SetAgenciaCentral = True
    
Erro_AgenciaCentral:

End Function

Public Function SetAgenciaApresentante(ByVal pAgenciaApresentante As String) As Boolean

    On Error GoTo Erro_AgenciaApresentante
    
    SetAgenciaApresentante = False
    
    Geral.AgenciaApresentante = pAgenciaApresentante
    
    SetAgenciaApresentante = True
    
Erro_AgenciaApresentante:

End Function


Public Function SetAutentica(ByRef pAutentica As Object) As Boolean

    On Error GoTo Erro_Autentica
    
    SetAutentica = False
    
    Set Autentica = pAutentica
    
    SetAutentica = True
    
Erro_Autentica:
    

End Function

Public Function SetAutenticadora(ByVal pAutenticadora As Integer) As Boolean

    On Error GoTo Erro_Autenticadora
    
    SetAutenticadora = False
    
    Geral.Autenticadora = pAutenticadora
    
    SetAutenticadora = True
    
Erro_Autenticadora:
    

End Function


Public Function SetDataProcessamento(ByRef pDataProcessamento As Long) As Boolean

    On Error GoTo Erro_SetDataProcessamento:
    
    SetDataProcessamento = False

    Geral.DataProcessamento = pDataProcessamento
    
    SetDataProcessamento = True
    
Erro_SetDataProcessamento:
    
    

End Function



Public Function SetConnection(ByRef pConnection As rdo.rdoConnection) As Boolean

    On Error GoTo Erro_Connection:
    
    SetConnection = False

    Set Geral.Banco = pConnection
    
    SetConnection = True
    
    Exit Function
    
Erro_Connection:
    
    

End Function


Public Function SetConnectionAgencia(ByRef pConnection As rdo.rdoConnection) As Boolean

    On Error GoTo Erro_ConnectionAgencia:
    
    SetConnectionAgencia = False

    Set Geral.BancoCaixa = pConnection
    
    SetConnectionAgencia = True
    
    Exit Function
    
Erro_ConnectionAgencia:
    
    

End Function



Public Function SetUsuario(ByVal pUsuarioLogin As String) As Boolean

    On Error GoTo Erro_UsuarioNome
    
    SetUsuario = False
    
    Geral.Usuario.Login = pUsuarioLogin
    
    SetUsuario = True
    
Erro_UsuarioNome:


End Function


Public Sub ShowModal()

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Se o objeto banco não estiver definido, então não entra'
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Geral.Banco Is Nothing Then
        MsgBox "Conexão não definida.", vbExclamation
        Exit Sub
    End If
    
    ''''''''''''''''''''''''
    'Start da autenticadora'
    ''''''''''''''''''''''''
    Screen.MousePointer = vbHourglass
    If Not Autentica Is Nothing Then Autentica.Inicia
    Screen.MousePointer = vbDefault
    
    '''''''''''''''''
    'Abre form modal'
    '''''''''''''''''
    FrmRegOcorr.Show vbModal
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''
    'Fechou a tela, então finaliza a autenticadora'
    '''''''''''''''''''''''''''''''''''''''''''''''
    If Not Autentica Is Nothing Then Autentica.Finaliza

End Sub


