VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Estatistica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Connection        As RDO.rdoConnection
Private m_Dataprocessamento As Long
Private m_Atualizacao       As Integer

Public Function SetConnection(ByRef pConnection As RDO.rdoConnection) As Boolean
    On Error GoTo Erro_Connection:
    SetConnection = False
    Set m_Connection = pConnection
    SetConnection = True
    Exit Function
Erro_Connection:
End Function

Public Function SetDataProcessamento(ByVal pDataProc As Long) As Boolean
    On Error GoTo Erro:
    SetDataProcessamento = False
    m_Dataprocessamento = pDataProc
    SetDataProcessamento = True
    Exit Function
Erro:
End Function

Public Function SetAtualizacao(ByVal pAtualizacao As Integer) As Boolean
    On Error GoTo Erro:
    SetAtualizacao = False
    m_Atualizacao = pAtualizacao
    SetAtualizacao = True
    Exit Function
Erro:
End Function

Public Sub ShowModal()
    Set frmEstatistica.m_Connection = m_Connection
    frmEstatistica.m_Atualizacao = m_Atualizacao
    frmEstatistica.m_Dataprocessamento = m_Dataprocessamento
    frmEstatistica.Show vbModal
End Sub


