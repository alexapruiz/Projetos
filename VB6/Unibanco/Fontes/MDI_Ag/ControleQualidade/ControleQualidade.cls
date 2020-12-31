VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControleQualidade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SetConnection(ByRef pConnection As rdo.rdoConnection) As Boolean

    On Error GoTo Erro_Connection:
    
    SetConnection = False

    Set Geral.Banco = pConnection
    
    SetConnection = True
    
    Exit Function
    
Erro_Connection:
    
    

End Function


Public Function SetDataProcessamento(ByRef pDataProcessamento As Long) As Boolean

    On Error GoTo Erro_SetDataProcessamento:
    
    SetDataProcessamento = False

    Geral.DataProcessamento = pDataProcessamento
    
    SetDataProcessamento = True
    
Erro_SetDataProcessamento:
    
    

End Function




Public Function SetDriveCDR(ByVal pDrive As String) As Boolean

    On Error GoTo Erro_DriveCDR:
    SetDriveCDR = False
    
    
    Geral.CDR.Drive = pDrive
    
    SetDriveCDR = True
    
Erro_DriveCDR:
    
End Function

Public Function SetDirImagensCDR(ByVal pDirImagens As String) As Boolean

    On Error GoTo Erro_ImagensCDR:
    SetDirImagensCDR = False
    
    
    Geral.CDR.DiretorioImagens = pDirImagens
    
    SetDirImagensCDR = True
    
Erro_ImagensCDR:
    
End Function


'Public Function SetIdEnv_Mal(ByVal pIdEnv_Mal As String) As Boolean
'
'    On Error GoTo Erro_IdEnv_Mal
'
'    SetIdEnv_Mal = False
'
'
'
'    SetIdEnv_Mal = True
'
'Erro_IdEnv_Mal:
'
'
'
'End Function
'
Public Function SetIntervalo(ByVal pIntervalo As Integer) As Boolean

    On Error GoTo Erro_Intervalo:
    
    
    SetIntervalo = False
    
    
    Geral.Intervalo = pIntervalo
    
    SetIntervalo = True
    
Erro_Intervalo:

End Function

Public Function SetAtualizacao(ByVal pAtualizacao As Integer) As Boolean

    On Error GoTo Erro_Atualizacao:
    
    SetAtualizacao = False
    
    Geral.Atualizacao = pAtualizacao
    
    SetAtualizacao = True
    
Erro_Atualizacao:

End Function


Public Function SetAgenciaApresentante(ByVal pAgenciaApresentante As String) As Boolean

    On Error GoTo Erro_AgenciaApresentante
    
    SetAgenciaApresentante = False
    
    Geral.AgenciaApresentante = pAgenciaApresentante
    
    SetAgenciaApresentante = True
    
Erro_AgenciaApresentante:

End Function



Public Function SetDirImagens(ByVal pDirImagens As String) As Boolean

    On Error GoTo Erro_DirImagens:
    
    SetDirImagens = False
    
    Geral.DiretorioImagens = pDirImagens
    
    SetDirImagens = True
    
Erro_DirImagens:

End Function



Public Function SetUsuario(ByVal pUsuarioLogin As String) As Boolean

    On Error GoTo Erro_UsuarioNome
    
    SetUsuario = False
    
    Geral.Usuario.Login = pUsuarioLogin
    
    SetUsuario = True
    
Erro_UsuarioNome:


End Function



Public Sub ShowModal()

    If Geral.Banco Is Nothing Then
        MsgBox "Conexão não definida.", vbExclamation
        Exit Sub
    End If
    
    frmControleQualidade.Show vbModal

End Sub



