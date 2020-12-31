VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParametro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_PracaCompensacao          As String
Dim m_DataServer                As String
Dim m_TipoAgencia               As String
Dim m_ValorLimiteInferior       As Double
Dim m_AgenciaCentral            As Integer
Dim m_AgenciaSatelite           As Integer
Dim m_ValorCompensaEnvelope     As Double
Dim m_ValorCompensaMaloteVelho  As Double
Dim m_CompensaMaloteNovo        As Double
Public Property Get PracaCompensacao() As String
    PracaCompensacao = m_PracaCompensacao
End Property
Public Property Let PracaCompensacao(New_PracaCompensacao As String)
    m_PracaCompensacao = New_PracaCompensacao
End Property
Public Property Get DataServer() As String
    DataServer = m_DataServer
End Property
Public Property Let DataServer(ByVal New_DataServer As String)
    m_DataServer = Format$(New_DataServer, "000000")
End Property
Public Property Get TipoAgencia() As String
    TipoAgencia = m_TipoAgencia
End Property
Public Property Let TipoAgencia(ByVal New_TipoAgencia As String)
    m_TipoAgencia = New_TipoAgencia
End Property
Public Property Get ValorLimiteInferior() As Double
    ValorLimiteInferior = m_ValorLimiteInferior
End Property
Public Property Let ValorLimiteInferior(ByVal New_ValorLimiteInferior As Double)
    m_ValorLimiteInferior = New_ValorLimiteInferior
End Property
Public Property Get AgenciaCentral() As Integer
    AgenciaCentral = m_AgenciaCentral
End Property
Public Property Let AgenciaCentral(ByVal New_AgenciaCentral As Integer)
    m_AgenciaCentral = New_AgenciaCentral
End Property
Public Property Get AgenciaSatelite() As Integer
    AgenciaSatelite = m_AgenciaSatelite
End Property
Public Property Let AgenciaSatelite(ByVal New_AgenciaSatelite As Integer)
    m_AgenciaSatelite = New_AgenciaSatelite
End Property
Public Property Get ValorCompensaEnvelope() As Double
    ValorCompensaEnvelope = m_ValorCompensaEnvelope
End Property
Public Property Let ValorCompensaEnvelope(ByVal New_ValorCompensaEnvelope As Double)
    m_ValorCompensaEnvelope = New_ValorCompensaEnvelope
End Property
Public Property Get ValorCompensaMaloteVelho() As Double
    ValorCompensaMaloteVelho = m_ValorCompensaMaloteVelho
End Property
Public Property Let ValorCompensaMaloteVelho(ByVal New_ValorCompensaMaloteVelho As Double)
    m_ValorCompensaMaloteVelho = New_ValorCompensaMaloteVelho
End Property
Public Property Get CompensaMaloteNovo() As Double
    CompensaMaloteNovo = m_CompensaMaloteNovo
End Property
Public Property Let CompensaMaloteNovo(ByVal New_CompensaMaloteNovo As Double)
    m_CompensaMaloteNovo = New_CompensaMaloteNovo
End Property
