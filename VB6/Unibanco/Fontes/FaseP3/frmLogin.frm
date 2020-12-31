VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Desbloquear Aplicação"
   ClientHeight    =   1548
   ClientLeft      =   2832
   ClientTop       =   3480
   ClientWidth     =   3756
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFechar 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   348
      Left            =   1356
      TabIndex        =   2
      Top             =   1068
      Width           =   1044
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   135
      Width           =   2088
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   348
      Left            =   267
      TabIndex        =   1
      Top             =   1068
      Width           =   1044
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancelar"
      Height          =   348
      Left            =   2445
      TabIndex        =   3
      Top             =   1068
      Width           =   1044
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   525
      Width           =   2088
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Usuário :"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   5
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Senha :"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   6
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private iOperacao       As Integer
Private sUsuario        As String
Private sSenha          As String
Private strctUsuario    As TpUsuario
Friend Function Senha_Ok(ByRef pTPUsuario As TpUsuario) As Integer

    Senha_Ok = 0

    strctUsuario = pTPUsuario
    txtUserName.Text = strctUsuario.Usuario
    
    iOperacao = 0

    Me.Show vbModal
    
    Senha_Ok = iOperacao

    If iOperacao = 1 Then
        pTPUsuario.Usuario = sUsuario
        pTPUsuario.Senha = sSenha
    End If

End Function

Private Sub CmdCancel_Click()

    iOperacao = 0
    
    Unload Me

End Sub

Private Sub cmdFechar_Click()

    If MsgBox("Encerrar a Aplicação MDI ?", vbQuestion + vbYesNo) = vbYes Then
        iOperacao = 2
        Principal.mnuSair_Click (0)
    End If

End Sub

Private Sub CmdOK_Click()

    Dim sStr        As String

    sUsuario = Trim(txtUserName.Text)
    sSenha = txtPassword.Text
    
    If ((UCase(Decript(strctUsuario.Senha)) <> UCase(sSenha)) And UCase(Trim(strctUsuario.Usuario)) <> "DESENV") Or _
       ((UCase(Trim(strctUsuario.Usuario)) = "DESENV") And (UCase(strctUsuario.Senha) <> UCase(sSenha))) Then
               
               sStr = "Senha inválida." & Chr(10) & Chr(10)
        sStr = sStr & "Somente o usuário que bloqueou esta" & Chr(10)
        sStr = sStr & "aplicação poderá desbloqueá-la."
    
        MsgBox sStr, vbCritical
        SelecionarTexto txtPassword
        txtPassword.SetFocus
        
        Exit Sub
    End If
    
    iOperacao = 1
    
    Unload Me

End Sub

