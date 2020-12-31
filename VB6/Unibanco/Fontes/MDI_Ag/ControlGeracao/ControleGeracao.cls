VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControleGeracao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SetDirImagensCDR(ByVal pDirImagens As String) As Boolean

    On Error GoTo Erro_ImagensCDR:
    SetDirImagensCDR = False
    
    
    Geral.CDR.DiretorioImagens = pDirImagens
    
    SetDirImagensCDR = True
    
Erro_ImagensCDR:
    
End Function



Public Function SetDirDadosCDR(ByVal pDirDados As String) As Boolean

    On Error GoTo Erro_DadosCDR:
    SetDirDadosCDR = False
    
    Geral.CDR.DiretorioDados = pDirDados
    
    SetDirDadosCDR = True
    
Erro_DadosCDR:
    
End Function




Public Function SetDriveCDR(ByVal pDrive As String) As Boolean

    On Error GoTo Erro_DriveCDR:
    SetDriveCDR = False
    
    
    Geral.CDR.Drive = pDrive
    
    SetDriveCDR = True
    
Erro_DriveCDR:
    
End Function


Public Function SetAgenciaApresentante(ByVal pAgenciaApresentante As String) As Boolean

    On Error GoTo Erro_AgenciaApresentante:
    
    SetAgenciaApresentante = False

    Geral.AgenciaApresentante = pAgenciaApresentante
    
    SetAgenciaApresentante = True
    
Erro_AgenciaApresentante:
    
    

End Function


Public Function SetConnection(ByVal pConnection As rdo.rdoConnection) As Boolean


    On Error GoTo Erro_SetConnection:
    
    SetConnection = False
    
    
    Set Geral.Banco = pConnection
    
    SetConnection = True
    
Erro_SetConnection:

End Function


Public Function SetDataProcessamento(ByVal pDataProcessamento As Long) As Boolean

    On Error GoTo Erro_DataProcessamento:
    
    SetDataProcessamento = False
    
    Geral.DataProcessamento = pDataProcessamento
    
    SetDataProcessamento = True
    
    
Erro_DataProcessamento:

End Function


Public Function SetDiretorioDados(ByVal pDirDados As String) As Boolean
    
    On Error GoTo Erro_SetDirDados:
    
    SetDiretorioDados = False
    
    Geral.DiretorioDados = pDirDados
    
    SetDiretorioDados = True
    
Erro_SetDirDados:
        
End Function

Public Function SetDiretorioImagens(ByVal pDirImagens As String) As Boolean
    
    On Error GoTo Erro_SetDirImagens:
    
    SetDiretorioImagens = False
    
    Geral.DiretorioImagens = pDirImagens
    
    SetDiretorioImagens = True
    
Erro_SetDirImagens:
        
End Function


Public Sub ShowModal()

    If Geral.Banco Is Nothing Then
        MsgBox "Conexão não definida.", vbExclamation
        Exit Sub
    End If

    frmControleGeracao.Show vbModal
End Sub


