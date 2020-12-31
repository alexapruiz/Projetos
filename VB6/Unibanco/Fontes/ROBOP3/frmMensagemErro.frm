VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmMensagemErro 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3192
   ClientLeft      =   36
   ClientTop       =   36
   ClientWidth     =   7872
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3192
   ScaleWidth      =   7872
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Sair"
      Height          =   550
      Left            =   4524
      Picture         =   "frmMensagemErro.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2620
      Width           =   800
   End
   Begin VB.CommandButton cmdRepetir 
      Caption         =   "&Repetir"
      Height          =   550
      Left            =   3600
      Picture         =   "frmMensagemErro.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2620
      Width           =   800
   End
   Begin VB.CommandButton cmdContinuar 
      Caption         =   "&Continuar"
      Height          =   550
      Left            =   2592
      Picture         =   "frmMensagemErro.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2620
      Visible         =   0   'False
      Width           =   800
   End
   Begin RichTextLib.RichTextBox txtMensagem 
      Height          =   1995
      Left            =   75
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   585
      Width           =   7710
      _ExtentX        =   13610
      _ExtentY        =   3514
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMensagemErro.frx":0DE2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image3 
      Height          =   384
      Left            =   60
      Picture         =   "frmMensagemErro.frx":0E72
      Top             =   60
      Width           =   384
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   30
      Top             =   75
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Caixa Robô: Aviso de Falha no Sistema"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1650
      TabIndex        =   1
      Top             =   150
      Width           =   4440
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   525
      Left            =   30
      Shape           =   4  'Rounded Rectangle
      Top             =   30
      Width           =   7815
   End
End
Attribute VB_Name = "frmMensagemErro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ErroNumero          As String
Private ErroDescricao       As String
Private ErroDispositivo     As String
Private Linhas(6, 2)        As Long

Private m_Retorno       As enumRetornoMensagemErro
Public Function ShowModal(ByVal pDescricao As String, _
                          ByRef pErr As ErrObject, _
                          Optional ByVal ExibeCMDContinua As Boolean = False) As enumRetornoMensagemErro
                          
    Dim sStr            As String
    Dim ehDeadLock      As Integer
    
    ehDeadLock = InStr(1, UCase(pErr.Description), "DEADLOCKED", vbTextCompare)
    
    If ehDeadLock <> 0 Then
        MDIQuery.insLogErro Geral.DataProcessamento, Caixa.Estacao, pErr.Number, "Erro: " & ErroDescricao & "Módulo: " & pDescricao
        Espera (0.5)
        ShowModal = eContinuar
        Exit Function
    End If
        
    cmdContinuar.Visible = ExibeCMDContinua
    
    '''''''''''''''''''''''''''
    'Coloca a descricao normal'
    '''''''''''''''''''''''''''
    
    ErroNumero = Trim(pErr.Number)
    ErroDescricao = Replace(pErr.Description, "'", "|", 1, Len(pErr.Description))
    ErroDispositivo = pErr.Source
    
    Linhas(0, 0) = 0
    sStr = "Origem: " & vbCrLf
    Linhas(0, 1) = Len(sStr)
    
    Linhas(1, 0) = Linhas(0, 1)
    sStr = sStr & Trim(pDescricao) & vbCrLf
    Linhas(1, 1) = Len(sStr)
    
    Linhas(2, 0) = Linhas(1, 1)
    sStr = sStr & "Dispositivo: " & vbCrLf
    Linhas(2, 1) = Len(sStr)
    
    Linhas(3, 0) = Linhas(2, 1)
    sStr = sStr & ErroDispositivo & vbCrLf
    Linhas(3, 1) = Len(sStr)
    
    Linhas(4, 0) = Linhas(3, 1)
    sStr = sStr & "Descrição: " & vbCrLf
    Linhas(4, 1) = Len(sStr)
    
    Linhas(5, 0) = Linhas(4, 1)
    sStr = sStr & ErroDescricao
    Linhas(5, 1) = Len(sStr)
        
    txtMensagem.Text = sStr
                
    ''''''''''''''
    'Grava o Erro'
    ''''''''''''''
    On Error GoTo TrataErro
       
    MDIQuery.insLogErro Geral.DataProcessamento, Caixa.Estacao, pErr.Number, "Erro: " & ErroDescricao & "Módulo: " & pDescricao
    
    Me.Show vbModal
    ShowModal = m_Retorno
       
    Exit Function
    
TrataErro:
    MsgBox "Falha no Módulo de Tratamento de Erro", vbCritical + vbOKOnly

End Function
Private Sub cmdContinuar_Click()
    m_Retorno = eContinuar
    Unload Me
End Sub
Private Sub CmdRepetir_Click()
    m_Retorno = eRepetir
    Unload Me
End Sub
Private Sub cmdSair_Click()
    m_Retorno = eSair
    Unload Me
End Sub
Private Function TratarStringErro(ByVal pvsTexto As String) As String
    Dim i       As Long
    Dim sAux    As String
    
    sAux = ""
    
    For i = 1 To Len(pvsTexto)
        If Mid(pvsTexto, i, 1) <> "'" Then
            sAux = sAux & Mid(pvsTexto, i, 1)
        End If
    Next
    
    TratarStringErro = sAux
End Function
Private Sub Form_Activate()
    ''''''''''''''''''''''''
    'Centraliza a descrição'
    ''''''''''''''''''''''''
    FormataText txtMensagem, Linhas(0, 0), Linhas(0, 1), , vbCenter, 11, azul, True, True
    FormataText txtMensagem, Linhas(1, 0), Linhas(1, 1), , vbCenter, 11, PRETO, False, False
    FormataText txtMensagem, Linhas(2, 0), Linhas(2, 1), , vbCenter, 11, azul, True, True
    FormataText txtMensagem, Linhas(3, 0), Linhas(3, 1), , vbCenter, 11, PRETO, False, False
    FormataText txtMensagem, Linhas(4, 0), Linhas(4, 1), , vbCenter, 11, Vermelho, True, False
    FormataText txtMensagem, Linhas(5, 0), Linhas(5, 1), , vbCenter, 11, PRETO, False, False
 
    txtMensagem.SelStart = 0
    txtMensagem.SelLength = 0

End Sub
