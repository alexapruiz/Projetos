VERSION 5.00
Begin VB.Form Login 
   Caption         =   "Login"
   ClientHeight    =   2595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3675
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2595
   ScaleWidth      =   3675
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   1845
      TabIndex        =   5
      Top             =   1845
      Width           =   1140
   End
   Begin VB.CommandButton CmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   405
      TabIndex        =   4
      Top             =   1845
      Width           =   1140
   End
   Begin VB.TextBox txtSenha 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1215
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1080
      Width           =   1500
   End
   Begin VB.TextBox txtUsuario 
      Height          =   285
      Left            =   1215
      TabIndex        =   1
      Top             =   630
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Senha"
      Height          =   195
      Left            =   615
      TabIndex        =   2
      Top             =   1125
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Usuário"
      Height          =   195
      Left            =   540
      TabIndex        =   0
      Top             =   675
      Width           =   540
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancelar_Click()

    End
End Sub

Private Sub CmdOK_Click()

    If UCase(Trim(txtUsuario.Text)) = "ALEX" Or UCase(Trim(txtUsuario.Text)) = "RUBIA" Then
        If UCase(Trim(txtSenha.Text)) = "BIA1701" Then
            Principal.Show 1
        Else
            MsgBox "Senha incorreta !", vbOKOnly, "SGB"
            'Call CriaLogUsuario
            Exit Sub
        End If
    Else
        MsgBox "Usuário não cadastrado !", vbOKOnly, "SGB"
        'Call CriaLogUsuario
        Exit Sub
    End If
End Sub

Private Sub Form_Load()

    If Command = "debug" Then
        Principal.Show 1
    End If
End Sub
