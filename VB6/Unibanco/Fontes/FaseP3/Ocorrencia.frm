VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Ocorrencia 
   Caption         =   "Ocorrências de Processamento"
   ClientHeight    =   5244
   ClientLeft      =   -204
   ClientTop       =   1680
   ClientWidth     =   11232
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5244
   ScaleWidth      =   11232
   Begin VB.CommandButton cmdRemoverOcorrencia 
      Cancel          =   -1  'True
      Caption         =   "Remover Ocorrência"
      Height          =   324
      Left            =   4758
      TabIndex        =   1
      Top             =   4632
      Width           =   1716
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Fechar"
      Height          =   324
      Left            =   6558
      TabIndex        =   2
      Top             =   4632
      Width           =   1716
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&Confirmar"
      Default         =   -1  'True
      Height          =   324
      Left            =   2958
      TabIndex        =   0
      Top             =   4632
      Width           =   1716
   End
   Begin TabDlg.SSTab TabTipoOCorr 
      Height          =   5088
      Left            =   84
      TabIndex        =   3
      Top             =   84
      Width           =   11076
      _ExtentX        =   19537
      _ExtentY        =   8975
      _Version        =   393216
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   420
      TabCaption(0)   =   "&Envelope/Malote"
      TabPicture(0)   =   "Ocorrencia.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LstOcorrencias(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Depósito"
      TabPicture(1)   =   "Ocorrencia.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LstOcorrencias(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&Pagamento"
      TabPicture(2)   =   "Ocorrencia.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "LstOcorrencias(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Di&versos"
      TabPicture(3)   =   "Ocorrencia.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "LstOcorrencias(8)"
      Tab(3).Control(1)=   "LstOcorrencias(3)"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "&Aut. Débito"
      TabPicture(4)   =   "Ocorrencia.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "LstOcorrencias(4)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "&Transf. Valor"
      TabPicture(5)   =   "Ocorrencia.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "LstOcorrencias(5)"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "&Operacional"
      TabPicture(6)   =   "Ocorrencia.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "LstOcorrencias(6)"
      Tab(6).ControlCount=   1
      Begin VB.ListBox LstOcorrencias 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   3888
         Index           =   2
         ItemData        =   "Ocorrencia.frx":00C4
         Left            =   -74796
         List            =   "Ocorrencia.frx":00C6
         TabIndex        =   11
         Top             =   432
         Width           =   10728
      End
      Begin VB.ListBox LstOcorrencias 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   3888
         Index           =   6
         ItemData        =   "Ocorrencia.frx":00C8
         Left            =   -74796
         List            =   "Ocorrencia.frx":00CA
         TabIndex        =   10
         Top             =   432
         Width           =   10728
      End
      Begin VB.ListBox LstOcorrencias 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   3888
         Index           =   5
         ItemData        =   "Ocorrencia.frx":00CC
         Left            =   -74796
         List            =   "Ocorrencia.frx":00CE
         TabIndex        =   9
         Top             =   432
         Width           =   10728
      End
      Begin VB.ListBox LstOcorrencias 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   3888
         Index           =   4
         ItemData        =   "Ocorrencia.frx":00D0
         Left            =   -74796
         List            =   "Ocorrencia.frx":00D2
         TabIndex        =   8
         Top             =   432
         Width           =   10728
      End
      Begin VB.ListBox LstOcorrencias 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   3888
         Index           =   3
         ItemData        =   "Ocorrencia.frx":00D4
         Left            =   -74796
         List            =   "Ocorrencia.frx":00D6
         TabIndex        =   7
         Top             =   432
         Width           =   10728
      End
      Begin VB.ListBox LstOcorrencias 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   3888
         Index           =   8
         ItemData        =   "Ocorrencia.frx":00D8
         Left            =   -74796
         List            =   "Ocorrencia.frx":00DA
         TabIndex        =   6
         Top             =   432
         Width           =   10728
      End
      Begin VB.ListBox LstOcorrencias 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   3888
         Index           =   1
         ItemData        =   "Ocorrencia.frx":00DC
         Left            =   -74796
         List            =   "Ocorrencia.frx":00DE
         TabIndex        =   5
         Top             =   432
         Width           =   10728
      End
      Begin VB.ListBox LstOcorrencias 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   3888
         Index           =   0
         ItemData        =   "Ocorrencia.frx":00E0
         Left            =   204
         List            =   "Ocorrencia.frx":00E2
         TabIndex        =   4
         Top             =   432
         Width           =   10728
      End
   End
End
Attribute VB_Name = "Ocorrencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Variáveis do RDO
Private qryGetocorrencia    As rdoQuery

'Variáveis de Trabalho
Public Result               As Integer
Public CodOcorr             As Integer

'Variável utilizada no complemento da ocorrência
'***** (Esta variável não é limpa neste form) *****
Public m_Descricao          As String
'

Private Sub cmdCancelar_Click()
  Result = 0
  Me.Hide
End Sub


Private Sub CmdOK_Click()
    'Verificar qual o TAB Ativo
    'Verificar se foi selecionado um item
    If LstOcorrencias(TabTipoOCorr.Tab).ListIndex <> -1 Then
        CodOcorr = LstOcorrencias(TabTipoOCorr.Tab).ItemData(LstOcorrencias(TabTipoOCorr.Tab).ListIndex)
        Result = 1
        
'''        frmComplRegOcorr.m_Descricao = m_Descricao
'''        frmComplRegOcorr.Show vbModal, Me
'''        m_Descricao = frmComplRegOcorr.m_Descricao

        Me.Hide
    Else
        MsgBox "Selecione uma Ocorrência.", vbInformation, App.Title
        Exit Sub
    End If
End Sub

Private Sub cmdRemoverOcorrencia_Click()
    
    m_Descricao = ""
    
    CodOcorr = 0
    Result = 2
    Me.Hide
End Sub

Private Sub Form_Activate()
  TabTipoOCorr.Tab = 0
  LstOcorrencias(0).ListIndex = 0
  LstOcorrencias(0).SetFocus
  Result = 0
  
  'Centralizar  form
  Me.Top = (Screen.Height - Me.Height) / 2
  Me.Left = (Screen.Width - Me.Width) / 2
  
  'Se módulo chamador é CSP, esconder opção do botão de Remover ocorrência devido ao form CSP
  'já efetuar este evento com diferenças na remoção.
  If Me.cmdRemoverOcorrencia.Tag = "CSP" Then
    Me.cmdRemoverOcorrencia.Visible = False
    Me.CmdOK.Left = Me.CmdOK.Left + (Me.cmdRemoverOcorrencia.Width / 2)
    Me.CmdCancelar.Left = Me.CmdCancelar.Left - (Me.cmdRemoverOcorrencia.Width / 2)
  End If
  
End Sub


Private Sub Form_Load()

  Dim RsOcorrencia As rdoResultset
  
  'Preencher a Lista com as Ocorrencias da tabela OCORRENCIA
  Set qryGetocorrencia = Geral.Banco.CreateQuery("", "{call GetTodasOcorrencia }")

  Set RsOcorrencia = qryGetocorrencia.OpenResultset(rdOpenStatic, rdConcurReadOnly)
  If Not RsOcorrencia.EOF Then
    Do Until RsOcorrencia.EOF
      Select Case Val(RsOcorrencia!Ocorrencia)
        Case 1 To 100
          LstOcorrencias(0).AddItem Format(RsOcorrencia!Ocorrencia, "000") & " - " & RsOcorrencia!Descricao
          LstOcorrencias(0).ItemData(LstOcorrencias(0).NewIndex) = RsOcorrencia!Ocorrencia
        Case 101 To 200
          LstOcorrencias(1).AddItem Format(RsOcorrencia!Ocorrencia, "000") & " - " & RsOcorrencia!Descricao
          LstOcorrencias(1).ItemData(LstOcorrencias(1).NewIndex) = RsOcorrencia!Ocorrencia
        Case 201 To 300
          LstOcorrencias(2).AddItem Format(RsOcorrencia!Ocorrencia, "000") & " - " & RsOcorrencia!Descricao
          LstOcorrencias(2).ItemData(LstOcorrencias(2).NewIndex) = RsOcorrencia!Ocorrencia
        Case 301 To 400
          LstOcorrencias(3).AddItem Format(RsOcorrencia!Ocorrencia, "000") & " - " & RsOcorrencia!Descricao
          LstOcorrencias(3).ItemData(LstOcorrencias(3).NewIndex) = RsOcorrencia!Ocorrencia
        Case 401 To 500
          LstOcorrencias(4).AddItem Format(RsOcorrencia!Ocorrencia, "000") & " - " & RsOcorrencia!Descricao
          LstOcorrencias(4).ItemData(LstOcorrencias(4).NewIndex) = RsOcorrencia!Ocorrencia
        Case 501 To 599
          LstOcorrencias(5).AddItem Format(RsOcorrencia!Ocorrencia, "000") & " - " & RsOcorrencia!Descricao
          LstOcorrencias(5).ItemData(LstOcorrencias(5).NewIndex) = RsOcorrencia!Ocorrencia
      End Select

      RsOcorrencia.MoveNext
      DoEvents
    Loop

    'Preenchendo Ocorrência Fixa - 999
    LstOcorrencias(6).AddItem "999 - Erro Operacional"
    LstOcorrencias(6).ItemData(LstOcorrencias(6).NewIndex) = 999
  Else
    MsgBox "Nenhuma Ocorrência Cadastrada.", vbInformation, App.Title
    Result = 0
    Me.Hide
  End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Result = 0
End Sub

Private Sub LstOcorrencias_DblClick(Index As Integer)
  Call CmdOK_Click
End Sub
Private Sub LstOcorrencias_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    Call CmdOK_Click
  End If
End Sub
Private Sub TabTipoOCorr_Click(PreviousTab As Integer)
  LstOcorrencias(TabTipoOCorr.Tab).ListIndex = 0
  LstOcorrencias(TabTipoOCorr.Tab).SetFocus
End Sub

