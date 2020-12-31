VERSION 5.00
Begin VB.Form AlteraSenha 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alteração de Senha"
   ClientHeight    =   2568
   ClientLeft      =   2532
   ClientTop       =   3156
   ClientWidth     =   4476
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2568
   ScaleWidth      =   4476
   Begin VB.Frame Frame1 
      Height          =   2136
      Left            =   240
      TabIndex        =   8
      Top             =   192
      Width           =   4008
      Begin VB.TextBox txtSenhaAtual 
         ForeColor       =   &H00800000&
         Height          =   288
         IMEMode         =   3  'DISABLE
         Left            =   2160
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   288
         Width           =   1332
      End
      Begin VB.CommandButton cmdAlterar 
         Caption         =   "&Alterar"
         Height          =   324
         Left            =   720
         TabIndex        =   6
         Top             =   1584
         Width           =   1116
      End
      Begin VB.CommandButton cmdFechar 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   324
         Left            =   1968
         TabIndex        =   7
         Top             =   1584
         Width           =   1116
      End
      Begin VB.TextBox txtConfSenha 
         ForeColor       =   &H00800000&
         Height          =   288
         IMEMode         =   3  'DISABLE
         Left            =   2160
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   1056
         Width           =   1332
      End
      Begin VB.TextBox txtSenha 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   288
         IMEMode         =   3  'DISABLE
         Left            =   2172
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   720
         Width           =   1332
      End
      Begin VB.Label lblSenhaAtual 
         AutoSize        =   -1  'True
         Caption         =   "Senha Atual:"
         Height          =   192
         Left            =   384
         TabIndex        =   0
         Top             =   336
         Width           =   900
      End
      Begin VB.Label lblConfSenha 
         AutoSize        =   -1  'True
         Caption         =   "Confirmação da Senha:"
         Height          =   192
         Left            =   348
         TabIndex        =   4
         Top             =   1104
         Width           =   1680
      End
      Begin VB.Label lblSenha 
         AutoSize        =   -1  'True
         Caption         =   "Nova Senha:"
         Height          =   192
         Left            =   372
         TabIndex        =   2
         Top             =   768
         Width           =   936
      End
   End
End
Attribute VB_Name = "AlteraSenha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private qryAlteraSenhaUsuario   As rdoQuery
Private rsAlteraSenhaUsuario    As rdoResultset
Private bTrataSenhaExpirada          As Boolean

Private Function ValidaCampos() As Boolean

On Error GoTo Err_ValidaCampos
    
    ValidaCampos = False
    
    If Trim(txtSenhaAtual.Text) = "" Then
        MsgBox "O campo Senha Atual deve ser preenchido.", vbExclamation + vbOKOnly, App.Title
        txtSenhaAtual.SelStart = 0
        txtSenhaAtual.SelLength = Len(txtSenhaAtual.Text)
        txtSenhaAtual.SetFocus
        Exit Function
    
    ElseIf Trim(txtSenha.Text) = "" Then
        MsgBox "O campo Senha deve ser preenchido.", vbExclamation + vbOKOnly, App.Title
        txtSenha.SelStart = 0
        txtSenha.SelLength = Len(txtSenha.Text)
        txtSenha.SetFocus
        Exit Function
        
    ElseIf Trim(txtSenha.Text) <> Trim(txtConfSenha.Text) Then
        MsgBox "Confirmação da senha não confere.", vbExclamation + vbOKOnly, App.Title
        txtConfSenha.SelStart = 0
        txtConfSenha.SelLength = Len(txtConfSenha.Text)
        txtConfSenha.SetFocus
        Exit Function
    End If

    ValidaCampos = True
    Exit Function
    
Err_ValidaCampos:
    MsgBox Err.Description, vbCritical, Me.Caption
    
End Function

Private Sub cmdAlterar_Click()
    
Dim sysTime As SYSTEMTIME
Dim lDataSistema As Long

On Error GoTo ErroAlterar
    
    GetLocalTime sysTime

    lDataSistema = Val(Format(sysTime.wYear, "0000") & Format(sysTime.wMonth, "00") & Format(sysTime.wDay, "00"))
    
    
    If Not ValidaCampos Then
        Exit Sub
    End If
    
    qryAlteraSenhaUsuario.rdoParameters(1) = Geral.Usuario
    qryAlteraSenhaUsuario.rdoParameters(2) = Encript(Trim(txtSenha.Text))
    qryAlteraSenhaUsuario.Execute
    
    If qryAlteraSenhaUsuario.rdoParameters(0) <> 0 Then
        MsgBox "Erro na alteração da senha.", vbCritical + vbOKOnly, App.Title
        Exit Sub
    End If
    
    'Atualiza type global
    Geral.DataUltimaTrocaSenhaUsuario = lDataSistema
    
    MsgBox "Senha alterada com sucesso.", vbInformation + vbOKOnly, App.Title
    
    
    If bTrataSenhaExpirada Then
        Password.bSenhaAlterada = True
    End If
    
    Unload Me
    
    Exit Sub
    
ErroAlterar:
    Select Case TratamentoErro("Erro na alteração da senha.", Err, rdoErrors)
        Case vbCancel
            Unload Me
        Case vbRetry
            Resume
    End Select

End Sub
Private Sub CmdFechar_Click()
    Unload Me
End Sub
Private Sub Form_Activate()

Dim nForm As Form

    bTrataSenhaExpirada = False
    For Each nForm In Forms
        If nForm.Name = "Password" Then bTrataSenhaExpirada = True: Exit For
    Next

   'Centraliza form
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
   'Inclusão de chamada a rotina AtualizaAtividade
   Call AtualizaAtividade(23)

End Sub
Private Sub Form_Load()

    Set qryAlteraSenhaUsuario = Geral.Banco.CreateQuery("", "{? = call AlteraSenhaUsuario (?,?)}")
    
    'Desabilita nova senha do usuário
    lblSenha.Enabled = False
    lblConfSenha.Enabled = False
    txtSenha.Enabled = False
    txtConfSenha.Enabled = False
    txtSenha.BackColor = &HC0C0C0           'Cinza
    txtConfSenha.BackColor = &HC0C0C0       'Cinza
    cmdAlterar.Enabled = False
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    If Not (rsAlteraSenhaUsuario Is Nothing) Then Set rsAlteraSenhaUsuario = Nothing
    qryAlteraSenhaUsuario.Close
    
End Sub


Private Sub txtConfSenha_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        cmdAlterar.SetFocus
    End If

End Sub
Private Sub txtSenha_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        SendKeys ("{TAB}")
    End If
    
End Sub

Private Sub txtSenhaAtual_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo Err_txtSenhaAtual

    If KeyCode = vbKeyReturn Then
        ' ****************************************
        ' * Testa Digitação Obrigatória da Senha *
        ' ****************************************
        If Trim(txtSenhaAtual.Text) = "" Then
            Beep
            MsgBox "Digite a Senha !", vbExclamation + vbOKOnly, App.Title
            With txtSenhaAtual
                .SelStart = 0
                .SelLength = Len(Trim(.Text))
                .SetFocus
            End With
            Exit Sub
        End If
    
        ' ***********************************
        ' * Verificação do Login do Usuário *
        ' ***********************************
        Dim tbUsuario As rdoResultset
        Dim qryUsuario As New rdoQuery
        Dim eRetorno As Variant
        
        
        Set qryUsuario = Geral.Banco.CreateQuery("", "{call GetUsuario (?)}")
        With qryUsuario
            .rdoParameters(0).Value = Trim(Geral.Usuario)
            Set tbUsuario = .OpenResultset(rdConcurReadOnly)
        End With
        
        eRetorno = VerificaUsuario(tbUsuario, Geral.Usuario, txtSenhaAtual.Text, False)
    
        If eRetorno = eNAO_EXISTENTE Then
            Beep
            MsgBox "Problema na verificação da senha do usuário. Tente novamente !", vbExclamation + vbOKOnly, App.Title
            With txtSenhaAtual
                .SelStart = 0
                .SelLength = Len(Trim(.Text))
                .SetFocus
            End With
        ElseIf eRetorno = eSENHA_INCORRETA Then
            Beep
            MsgBox "Senha não Confere !", vbExclamation + vbOKOnly, App.Title
            With txtSenhaAtual
                .SelStart = 0
                .SelLength = Len(Trim(.Text))
                .SetFocus
            End With
        Else
            'Habilita nova senha do usuário
            lblSenha.Enabled = True
            lblConfSenha.Enabled = True
            txtSenha.Enabled = True
            txtConfSenha.Enabled = True
            txtSenha.BackColor = vbWhite
            txtConfSenha.BackColor = vbWhite
            cmdAlterar.Enabled = True
            
            lblSenhaAtual.Enabled = False
            txtSenhaAtual.Enabled = False

            
            txtSenha.SetFocus
        End If
    End If
    

Exit_txtSenhaAtual:
    If Not (tbUsuario Is Nothing) Then Set tbUsuario = Nothing
    If Not (qryUsuario Is Nothing) Then qryUsuario.Close
    Exit Sub
    
Err_txtSenhaAtual:
    Beep
    MsgBox Err.Description, vbCritical, App.Title
    GoTo Exit_txtSenhaAtual
    
End Sub
