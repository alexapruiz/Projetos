VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2430
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   3570
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2430
   ScaleWidth      =   3570
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmd2 
      Height          =   405
      Left            =   1560
      TabIndex        =   8
      Top             =   600
      Width           =   435
   End
   Begin VB.CommandButton Cmd3 
      Height          =   405
      Left            =   2010
      TabIndex        =   7
      Top             =   600
      Width           =   435
   End
   Begin VB.CommandButton Cmd4 
      Height          =   405
      Left            =   1110
      TabIndex        =   6
      Top             =   1020
      Width           =   435
   End
   Begin VB.CommandButton Cmd5 
      Height          =   405
      Left            =   1560
      TabIndex        =   5
      Top             =   1020
      Width           =   435
   End
   Begin VB.CommandButton Cmd6 
      Height          =   405
      Left            =   2010
      TabIndex        =   4
      Top             =   1020
      Width           =   435
   End
   Begin VB.CommandButton Cmd7 
      Height          =   405
      Left            =   1110
      TabIndex        =   3
      Top             =   1440
      Width           =   435
   End
   Begin VB.CommandButton Cmd8 
      Height          =   405
      Left            =   1560
      TabIndex        =   2
      Top             =   1440
      Width           =   435
   End
   Begin VB.CommandButton Cmd9 
      Height          =   405
      Left            =   2010
      TabIndex        =   1
      Top             =   1440
      Width           =   435
   End
   Begin VB.CommandButton Cmd1 
      Height          =   405
      Left            =   1110
      TabIndex        =   0
      Top             =   600
      Width           =   435
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   60
      TabIndex        =   11
      Top             =   1950
      Width           =   3435
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Jogador 2 : O"
      Height          =   195
      Left            =   2520
      TabIndex        =   10
      Top             =   60
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Jogador 1 : X"
      Height          =   195
      Left            =   90
      TabIndex        =   9
      Top             =   60
      Width           =   945
   End
   Begin VB.Menu MnuNovoJogo 
      Caption         =   "Novo &Jogo"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public S As String
Public EMJOGO As Boolean

Private Sub Cmd1_Click()

    If Cmd1.Caption = "" Then
        'Casa ainda não selecionada -> Marcar
        Cmd1.Caption = Simbolo
    Else
        'Casa já selecionada -> Emitir mensagem
        MsgBox "Esta posição já foi selecionada !", vbInformation
    End If

    'Verificar se alguém venceu
    If VerificaGanhador("X") = True Then
        MsgBox "O vencedor foi o jogador 1"
    ElseIf VerificaGanhador("0") = True Then
        MsgBox "O vencedor foi o jogador 2"
    End If
End Sub
Public Function Simbolo() As String

    If S = "X" Then
        S = "0"
        Label3.Caption = "Jogador 1 !"
    Else
        S = "X"
        Label3.Caption = "Jogador 2 !"
    End If

    Simbolo = S
End Function
Private Sub Cmd2_Click()

    If Cmd2.Caption = "" Then
        'Casa ainda não selecionada -> Marcar
        Cmd2.Caption = Simbolo
    Else
        'Casa já selecionada -> Emitir mensagem
        MsgBox "Esta posição já foi selecionada !", vbInformation
    End If

    'Verificar se alguém venceu
    If VerificaGanhador("X") = True Then
        MsgBox "O vencedor foi o jogador 1"
    ElseIf VerificaGanhador("0") = True Then
        MsgBox "O vencedor foi o jogador 2"
    End If
End Sub
Private Sub Cmd3_Click()

    If Cmd3.Caption = "" Then
        'Casa ainda não selecionada -> Marcar
        Cmd3.Caption = Simbolo
    Else
        'Casa já selecionada -> Emitir mensagem
        MsgBox "Esta posição já foi selecionada !", vbInformation
    End If

    'Verificar se alguém venceu
    If VerificaGanhador("X") = True Then
        MsgBox "O vencedor foi o jogador 1"
    ElseIf VerificaGanhador("0") = True Then
        MsgBox "O vencedor foi o jogador 2"
    End If
End Sub
Private Sub Cmd4_Click()

    If Cmd4.Caption = "" Then
        'Casa ainda não selecionada -> Marcar
        Cmd4.Caption = Simbolo
    Else
        'Casa já selecionada -> Emitir mensagem
        MsgBox "Esta posição já foi selecionada !", vbInformation
    End If

    'Verificar se alguém venceu
    If VerificaGanhador("X") = True Then
        MsgBox "O vencedor foi o jogador 1"
    ElseIf VerificaGanhador("0") = True Then
        MsgBox "O vencedor foi o jogador 2"
    End If
End Sub
Private Sub Cmd5_Click()

    If Cmd5.Caption = "" Then
        'Casa ainda não selecionada -> Marcar
        Cmd5.Caption = Simbolo
    Else
        'Casa já selecionada -> Emitir mensagem
        MsgBox "Esta posição já foi selecionada !", vbInformation
    End If

    'Verificar se alguém venceu
    If VerificaGanhador("X") = True Then
        MsgBox "O vencedor foi o jogador 1"
    ElseIf VerificaGanhador("0") = True Then
        MsgBox "O vencedor foi o jogador 2"
    End If
End Sub
Private Sub Cmd6_Click()

    If Cmd6.Caption = "" Then
        'Casa ainda não selecionada -> Marcar
        Cmd6.Caption = Simbolo
    Else
        'Casa já selecionada -> Emitir mensagem
        MsgBox "Esta posição já foi selecionada !", vbInformation
    End If

    'Verificar se alguém venceu
    If VerificaGanhador("X") = True Then
        MsgBox "O vencedor foi o jogador 1"
    ElseIf VerificaGanhador("0") = True Then
        MsgBox "O vencedor foi o jogador 2"
    End If
End Sub
Private Sub Cmd7_Click()

    If Cmd7.Caption = "" Then
        'Casa ainda não selecionada -> Marcar
        Cmd7.Caption = Simbolo
    Else
        'Casa já selecionada -> Emitir mensagem
        MsgBox "Esta posição já foi selecionada !", vbInformation
    End If

    'Verificar se alguém venceu
    If VerificaGanhador("X") = True Then
        MsgBox "O vencedor foi o jogador 1"
    ElseIf VerificaGanhador("0") = True Then
        MsgBox "O vencedor foi o jogador 2"
    End If
End Sub
Private Sub Cmd8_Click()

    If Cmd8.Caption = "" Then
        'Casa ainda não selecionada -> Marcar
        Cmd8.Caption = Simbolo
    Else
        'Casa já selecionada -> Emitir mensagem
        MsgBox "Esta posição já foi selecionada !", vbInformation
    End If

    'Verificar se alguém venceu
    If VerificaGanhador("X") = True Then
        MsgBox "O vencedor foi o jogador 1"
    ElseIf VerificaGanhador("0") = True Then
        MsgBox "O vencedor foi o jogador 2"
    End If
End Sub
Private Sub Cmd9_Click()

    If Cmd9.Caption = "" Then
        'Casa ainda não selecionada -> Marcar
        Cmd9.Caption = Simbolo
    Else
        'Casa já selecionada -> Emitir mensagem
        MsgBox "Esta posição já foi selecionada !", vbInformation
    End If

    'Verificar se alguém venceu
    If VerificaGanhador("X") = True Then
        MsgBox "O vencedor foi o jogador 1"
    ElseIf VerificaGanhador("0") = True Then
        MsgBox "O vencedor foi o jogador 2"
    End If
End Sub
Private Sub Form_Load()

    S = "X"
    Call Simbolo
End Sub
Public Function VerificaGanhador(ByVal S As String) As Integer

    'Verificar todas as possibilidades
    If (Cmd1.Caption = S And Cmd2.Caption = S And Cmd3.Caption = S) Or _
        (Cmd1.Caption = S And Cmd4.Caption = S And Cmd7.Caption = S) Or _
        (Cmd1.Caption = S And Cmd5.Caption = S And Cmd9.Caption = S) Or _
        (Cmd2.Caption = S And Cmd5.Caption = S And Cmd8.Caption = S) Or _
        (Cmd3.Caption = S And Cmd5.Caption = S And Cmd7.Caption = S) Or _
        (Cmd3.Caption = S And Cmd6.Caption = S And Cmd9.Caption = S) Or _
        (Cmd4.Caption = S And Cmd5.Caption = S And Cmd6.Caption = S) Or _
        (Cmd7.Caption = S And Cmd8.Caption = S And Cmd9.Caption = S) Then

        Call ManipulaBotoes(False)
        VerificaGanhador = True
    End If
End Function
Public Sub ManipulaBotoes(ByVal Status As Boolean)

    Cmd1.Enabled = Status
    Cmd2.Enabled = Status
    Cmd3.Enabled = Status
    Cmd4.Enabled = Status
    Cmd5.Enabled = Status
    Cmd6.Enabled = Status
    Cmd7.Enabled = Status
    Cmd8.Enabled = Status
    Cmd9.Enabled = Status

    If Status = True Then
        Cmd1.Caption = ""
        Cmd2.Caption = ""
        Cmd3.Caption = ""
        Cmd4.Caption = ""
        Cmd5.Caption = ""
        Cmd6.Caption = ""
        Cmd7.Caption = ""
        Cmd8.Caption = ""
        Cmd9.Caption = ""
        
        S = "X"
        Call Simbolo
    End If
End Sub

Private Sub MnuNovoJogo_Click()

    Call ManipulaBotoes(True)

    S = "X"
    Call Simbolo
End Sub
