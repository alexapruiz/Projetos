VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_UsuarioAtual  As String
Dim m_CaixaIni      As Integer
Dim m_CaixaFim      As Integer
Dim m_Estacao       As String
Dim m_VersaoAtual   As Integer
Dim m_Caixa         As Integer
Dim m_BaseNSU       As Integer
Dim m_SDV           As String * 1
Dim m_CIF           As String * 6
Dim m_NSU           As Long
Dim m_NSU1          As Long
Dim m_NSU2          As Long
Dim m_NSU3          As Long
Public Property Get UsuarioAtual() As String
    UsuarioAtual = m_UsuarioAtual
End Property
Public Property Let UsuarioAtual(ByVal New_UsuarioAtual As String)
    m_UsuarioAtual = New_UsuarioAtual
End Property
Public Property Get CaixaIni() As Integer
    CaixaIni = m_CaixaIni
End Property
Public Property Let CaixaIni(ByVal New_CaixaIni As Integer)
    m_CaixaIni = New_CaixaIni
End Property
Public Property Get CaixaFim() As Integer
    CaixaFim = m_CaixaFim
End Property
Public Property Let CaixaFim(ByVal New_CaixaFim As Integer)
    m_CaixaFim = New_CaixaFim
End Property
Public Property Get Estacao() As String
    Estacao = m_Estacao
End Property
Public Property Let Estacao(ByVal New_Estacao As String)
    m_Estacao = Format(New_Estacao, "000")
End Property
Public Property Get VersaoAtual() As Integer
    VersaoAtual = m_VersaoAtual
End Property
Public Property Let VersaoAtual(ByVal New_VersaoAtual As Integer)
    m_VersaoAtual = New_VersaoAtual
End Property
Public Property Get Caixa() As Integer
    Caixa = m_Caixa
End Property
Public Property Let Caixa(ByVal New_Caixa As Integer)
    m_Caixa = New_Caixa
End Property
Public Property Get CIF() As String
    CIF = "999999"
End Property
Public Property Get SDV() As String
    SDV = m_SDV
End Property
Public Property Let SDV(ByVal New_SDV As String)
    m_SDV = New_SDV
End Property
Public Property Get NSU() As Long
    NSU = m_NSU
End Property
Public Property Let NSU(ByVal New_NSU As Long)
    m_NSU = New_NSU
End Property
Public Property Get NSU1() As Long
    NSU1 = m_NSU1
End Property
Public Property Let NSU1(ByVal New_NSU1 As Long)
    m_NSU1 = New_NSU1
End Property
Public Property Get NSU2() As Long
    NSU2 = m_NSU2
End Property
Public Property Let NSU2(ByVal New_NSU2 As Long)
    m_NSU2 = New_NSU2
End Property
Public Property Get NSU3() As Long
    NSU3 = m_NSU3
End Property
Public Property Let NSU3(ByVal New_NSU3 As Long)
    m_NSU3 = New_NSU3
End Property
Public Property Get BaseNSU() As Integer
    BaseNSU = m_BaseNSU
End Property
Public Property Let BaseNSU(ByVal New_BaseNSU As Integer)
    'If New_BaseNSU <= 0 Then
    '    Err.Raise 991, App.Title, "Numero sequencial NSU invalido"
    '    Exit Property
    'End If
    m_BaseNSU = New_BaseNSU
End Property
Sub AddNSU()
    Me.BaseNSU = Me.BaseNSU + 1
End Sub




