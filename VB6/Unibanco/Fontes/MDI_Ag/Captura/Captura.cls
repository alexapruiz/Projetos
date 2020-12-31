VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Captura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Connection          As RDO.rdoConnection
Private m_Scanner             As Scanner
Private m_TipoScanner         As enumScanner
Private m_AgenciaApresentante As String
Private m_DataProcessamento   As Long
Private m_DiretorioDados      As String
Private m_DiretorioImagens    As String
Private m_DiretorioTrabalho   As String
Private m_Usuario             As String

Public Function SetScanner(ByRef pScanner As Object) As Boolean
    On Error GoTo Erro_Scanner:
    SetScanner = False
    Set m_Scanner = pScanner
    SetScanner = True
    Exit Function
Erro_Scanner:
End Function

Public Function SetConnection(ByRef pConnection As RDO.rdoConnection) As Boolean
    On Error GoTo Erro_Connection:
    SetConnection = False
    Set m_Connection = pConnection
    SetConnection = True
    Exit Function
Erro_Connection:
End Function

Public Function SetTipoScanner(ByVal pTipoScanner As Integer) As Boolean
    On Error GoTo Erro:
    SetTipoScanner = False
    m_TipoScanner = pTipoScanner
    SetTipoScanner = True
    Exit Function
Erro:
End Function

Public Function SetAgApresentante(ByVal pAgApre As String) As Boolean
    On Error GoTo Erro:
    SetAgApresentante = False
    m_AgenciaApresentante = pAgApre
    SetAgApresentante = True
    Exit Function
Erro:
End Function

Public Function SetDataProcessamento(ByVal pDataProc As Long) As Boolean
    On Error GoTo Erro:
    SetDataProcessamento = False
    m_DataProcessamento = pDataProc
    SetDataProcessamento = True
    Exit Function
Erro:
End Function

Public Function SetDirDados(ByVal pDirDados As String) As Boolean
    On Error GoTo Erro:
    SetDirDados = False
    m_DiretorioDados = pDirDados
    SetDirDados = True
    Exit Function
Erro:
End Function

Public Function SetDirTrabalho(ByVal pDirTrabalho As String) As Boolean
    On Error GoTo Erro:
    SetDirTrabalho = False
    m_DiretorioTrabalho = pDirTrabalho
    SetDirTrabalho = True
    Exit Function
Erro:
End Function

Public Function SetDirImagens(ByVal pDirImagens As String) As Boolean
    On Error GoTo Erro:
    SetDirImagens = False
    m_DiretorioImagens = pDirImagens
    SetDirImagens = True
    Exit Function
Erro:
End Function

Public Function SetUsuario(ByVal pUsuario As String) As Boolean
    On Error GoTo Erro:
    SetUsuario = False
    m_Usuario = pUsuario
    SetUsuario = True
    Exit Function
Erro:
End Function

Public Sub ShowModal()
    Set frmCaptura.m_Connection = m_Connection
    Set frmCaptura.m_Scanner = m_Scanner
    frmCaptura.m_TipoScanner = m_TipoScanner
    frmCaptura.m_AgenciaApresentante = m_AgenciaApresentante
    frmCaptura.m_DataProcessamento = m_DataProcessamento
    frmCaptura.m_DiretorioDados = m_DiretorioDados
    frmCaptura.m_DiretorioImagens = m_DiretorioImagens
    frmCaptura.m_DiretorioTrabalho = m_DiretorioTrabalho
    frmCaptura.m_Usuario = m_Usuario
    frmCaptura.Show vbModal
End Sub


