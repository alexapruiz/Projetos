VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Controlador 
   Caption         =   "Serpente"
   ClientHeight    =   6180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6015
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6180
   ScaleWidth      =   6015
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Nível de Dificuldade"
      Height          =   1635
      Left            =   4230
      TabIndex        =   9
      Top             =   210
      Width           =   1695
      Begin VB.OptionButton OptImpos 
         Caption         =   "Muito Rapido"
         Height          =   225
         Left            =   150
         TabIndex        =   13
         Top             =   1200
         Width           =   1425
      End
      Begin VB.OptionButton OptRapido 
         Caption         =   "Rápido"
         Height          =   225
         Left            =   150
         TabIndex        =   12
         Top             =   900
         Width           =   1125
      End
      Begin VB.OptionButton OptNormal 
         Caption         =   "Normal"
         Height          =   225
         Left            =   150
         TabIndex        =   11
         Top             =   600
         Width           =   1125
      End
      Begin VB.OptionButton OptLento 
         Caption         =   "Lento"
         Height          =   225
         Left            =   150
         TabIndex        =   10
         Top             =   300
         Width           =   1125
      End
   End
   Begin Threed.SSCommand CmdEsquerda 
      Height          =   525
      Left            =   720
      TabIndex        =   1
      Top             =   5010
      Width           =   765
      _Version        =   65536
      _ExtentX        =   1349
      _ExtentY        =   926
      _StockProps     =   78
      Picture         =   "Serpente.frx":0000
   End
   Begin VB.PictureBox Picture1 
      Height          =   4065
      Left            =   120
      ScaleHeight     =   4005
      ScaleWidth      =   3975
      TabIndex        =   0
      Top             =   330
      Width           =   4035
      Begin Threed.SSPanel PnlSerpente 
         Height          =   400
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   400
         _Version        =   65536
         _ExtentX        =   706
         _ExtentY        =   706
         _StockProps     =   15
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Timer t 
      Interval        =   500
      Left            =   210
      Top             =   5070
   End
   Begin Threed.SSCommand CmdDescer 
      Height          =   525
      Left            =   1470
      TabIndex        =   2
      Top             =   5520
      Width           =   765
      _Version        =   65536
      _ExtentX        =   1349
      _ExtentY        =   926
      _StockProps     =   78
      Picture         =   "Serpente.frx":0452
   End
   Begin Threed.SSCommand CmdSubir 
      Height          =   525
      Left            =   1470
      TabIndex        =   3
      Top             =   4500
      Width           =   765
      _Version        =   65536
      _ExtentX        =   1349
      _ExtentY        =   926
      _StockProps     =   78
      Picture         =   "Serpente.frx":08A4
   End
   Begin Threed.SSCommand CmdDireita 
      Height          =   525
      Left            =   2220
      TabIndex        =   4
      Top             =   5010
      Width           =   765
      _Version        =   65536
      _ExtentX        =   1349
      _ExtentY        =   926
      _StockProps     =   78
      Picture         =   "Serpente.frx":0CF6
   End
   Begin Threed.SSCommand CmdPause 
      Height          =   525
      Left            =   1470
      TabIndex        =   8
      Top             =   5010
      Width           =   795
      _Version        =   65536
      _ExtentX        =   1402
      _ExtentY        =   926
      _StockProps     =   78
      Caption         =   "PAUSE"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
      Picture         =   "Serpente.frx":1148
   End
   Begin VB.Label LblPontos 
      Height          =   195
      Left            =   3600
      TabIndex        =   7
      Top             =   4500
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Pontos : "
      Height          =   195
      Left            =   2970
      TabIndex        =   6
      Top             =   4500
      Width           =   630
   End
End
Attribute VB_Name = "Controlador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CmdDescer_Click()

    S.DIRECAO = BAIXO
End Sub
Private Sub CmdDireita_Click()

    S.DIRECAO = DIREITA
End Sub
Private Sub CmdEsquerda_Click()

    S.DIRECAO = ESQUERDA
End Sub
Private Sub CmdPause_Click()

    t.Enabled = Not (t.Enabled)
End Sub
Private Sub CmdSubir_Click()

    S.DIRECAO = CIMA
End Sub
Private Sub CmdSubir_GotFocus()

    Me.SetFocus
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
    Case 37
        S.DIRECAO = ESQUERDA
    Case 38
        S.DIRECAO = CIMA
    Case 39
        S.DIRECAO = DIREITA
    Case 40
        S.DIRECAO = BAIXO
    End Select
End Sub
Private Sub Form_Load()

    S.DIRECAO = ESQUERDA
    Call S.Mover(45, PnlSerpente)
End Sub
Private Sub OptImpos_Click()

    t.Interval = 100
    Picture1.SetFocus
End Sub
Private Sub OptLento_Click()

    t.Interval = 3000
End Sub
Private Sub OptNormal_Click()

    t.Interval = 1000
End Sub
Private Sub OptRapido_Click()

    t.Interval = 500
End Sub
Private Sub t_Timer()

    Select Case S.DIRECAO
    Case ESQUERDA
        'Verificar os Limites
        If (S.Posicao Mod 10) = 1 Then
            'Bateu na parede
            t.Enabled = False
            MsgBox "Fim de Jogo"
            Exit Sub
        End If

        'Executar o movimento
        Call S.Mover(S.Posicao - 1, PnlSerpente)
    Case CIMA
        'Verificar os Limites
        Select Case S.Posicao
        Case 1 To 10
            'Bateu na parede
            t.Enabled = False
            MsgBox "Fim de Jogo"
            Exit Sub
        End Select

        'Executar o movimento
        Call S.Mover(S.Posicao - 10, PnlSerpente)
    Case DIREITA
        'Verificar os Limites
        If (S.Posicao Mod 10) = 0 Then
            'Bateu na parede
            t.Enabled = False
            MsgBox "Fim de Jogo"
            Exit Sub
        End If

        'Executar o movimento
        Call S.Mover(S.Posicao + 1, PnlSerpente)
    Case BAIXO
        'Verificar os Limites
        Select Case S.Posicao
        Case 91 To 100
            'Bateu na parede
            t.Enabled = False
            MsgBox "Fim de Jogo"
            Exit Sub
        End Select

        'Executar o movimento
        Call S.Mover(S.Posicao + 10, PnlSerpente)
    End Select

    LblPontos.Caption = Val(LblPontos.Caption) + 1
End Sub
