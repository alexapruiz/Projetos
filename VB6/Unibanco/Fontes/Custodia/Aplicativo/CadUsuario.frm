VERSION 5.00
Begin VB.Form CadUsuario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Usuário"
   ClientHeight    =   4305
   ClientLeft      =   90
   ClientTop       =   1035
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   7860
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Permissões"
      Height          =   1872
      Left            =   3972
      TabIndex        =   18
      Top             =   2364
      Width           =   3792
      Begin VB.ListBox lstPermissao 
         Height          =   1548
         IntegralHeight  =   0   'False
         Left            =   144
         TabIndex        =   6
         Top             =   240
         Width           =   3504
      End
   End
   Begin VB.CommandButton cmdFechar 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   372
      Left            =   6264
      TabIndex        =   10
      Top             =   1536
      Width           =   1452
   End
   Begin VB.CommandButton cmdRemover 
      Caption         =   "&Remover"
      Height          =   372
      Left            =   6264
      TabIndex        =   9
      Top             =   1092
      Width           =   1452
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "A&lterar"
      Height          =   372
      Left            =   6264
      TabIndex        =   8
      Top             =   660
      Width           =   1452
   End
   Begin VB.CommandButton cmdAdicionar 
      Caption         =   "&Adicionar"
      Height          =   372
      Left            =   6264
      TabIndex        =   7
      Top             =   216
      Width           =   1452
   End
   Begin VB.Frame Frame2 
      Caption         =   "Grupos"
      Height          =   1872
      Left            =   84
      TabIndex        =   17
      Top             =   2364
      Width           =   3792
      Begin VB.ListBox lstGrupo 
         ForeColor       =   &H00800000&
         Height          =   1548
         IntegralHeight  =   0   'False
         ItemData        =   "CadUsuario.frx":0000
         Left            =   144
         List            =   "CadUsuario.frx":0002
         Style           =   1  'Checkbox
         TabIndex        =   5
         Top             =   216
         Width           =   3504
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1992
      Left            =   84
      TabIndex        =   11
      Top             =   96
      Width           =   6000
      Begin VB.ComboBox cmbLogin 
         ForeColor       =   &H00800000&
         Height          =   288
         Left            =   1212
         TabIndex        =   0
         Top             =   240
         Width           =   1752
      End
      Begin VB.TextBox txtCif 
         ForeColor       =   &H00800000&
         Height          =   288
         Left            =   1212
         MaxLength       =   9
         TabIndex        =   2
         Top             =   1048
         Width           =   1752
      End
      Begin VB.TextBox txtConfSenha 
         ForeColor       =   &H00800000&
         Height          =   288
         IMEMode         =   3  'DISABLE
         Left            =   4452
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1440
         Width           =   1332
      End
      Begin VB.TextBox txtSenha 
         ForeColor       =   &H00800000&
         Height          =   288
         IMEMode         =   3  'DISABLE
         Left            =   1212
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1440
         Width           =   1332
      End
      Begin VB.TextBox txtNome 
         ForeColor       =   &H00800000&
         Height          =   312
         Left            =   1212
         MaxLength       =   50
         TabIndex        =   1
         Top             =   632
         Width           =   4572
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Identificação"
         Height          =   192
         Left            =   180
         TabIndex        =   16
         Top             =   1068
         Width           =   912
      End
      Begin VB.Label Label4 
         Caption         =   "Confirmação da Senha:"
         Height          =   252
         Left            =   2652
         TabIndex        =   15
         Top             =   1500
         Width           =   1812
      End
      Begin VB.Label Label3 
         Caption         =   "Nome:"
         Height          =   192
         Left            =   180
         TabIndex        =   14
         Top             =   692
         Width           =   552
      End
      Begin VB.Label Label2 
         Caption         =   "Senha:"
         Height          =   252
         Left            =   180
         TabIndex        =   13
         Top             =   1458
         Width           =   552
      End
      Begin VB.Label Label1 
         Caption         =   "Login:"
         Height          =   252
         Left            =   180
         TabIndex        =   12
         Top             =   258
         Width           =   492
      End
   End
End
Attribute VB_Name = "CadUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsUsuarios                 As New ADODB.Recordset
Private rsGrupos                   As New ADODB.Recordset
Private RsGrupoUsuario             As New ADODB.Recordset
Private Proc_Selecionar            As New Custodia.Selecionar
Private Proc_Excluir               As New Custodia.Excluir
Private Proc_Atualizar             As New Custodia.Atualizar
Private Proc_Inserir               As New Custodia.Inserir
Private CodIdGrupo()               As Integer

Private Function ValidaCampos() As Boolean
    Dim Count As Integer

    If Trim(cmbLogin.Text) = "" Then
        MsgBox "O campo Login deve ser preenchido.", vbExclamation + vbOKOnly, Me.Caption
        cmbLogin.SelStart = 0
        cmbLogin.SelLength = Len(cmbLogin.Text)
        cmbLogin.SetFocus
        ValidaCampos = False
    ElseIf Trim(txtNome.Text) = "" Then
        MsgBox "O campo Nome deve ser preenchido.", vbExclamation + vbOKOnly, Me.Caption
        txtNome.SelStart = 0
        txtNome.SelLength = Len(txtNome.Text)
        txtNome.SetFocus
        ValidaCampos = False
    ElseIf Trim(txtCif.Text) = "" Then
        MsgBox "O campo CIF deve ser preenchido.", vbExclamation + vbOKOnly, Me.Caption
        txtCif.SelStart = 0
        txtCif.SelLength = Len(txtCif.Text)
        txtCif.SetFocus
        ValidaCampos = False
    ElseIf Trim(txtSenha.Text) = "" Then
        MsgBox "O campo Senha deve ser preenchido.", vbExclamation + vbOKOnly, Me.Caption
        txtSenha.SelStart = 0
        txtSenha.SelLength = Len(txtSenha.Text)
        txtSenha.SetFocus
        ValidaCampos = False
    ElseIf Trim(txtSenha.Text) <> Trim(txtConfSenha.Text) Then
        MsgBox "Confirmação da senha não confere.", vbExclamation + vbOKOnly, Me.Caption
        txtConfSenha.SelStart = 0
        txtConfSenha.SelLength = Len(txtConfSenha.Text)
        txtConfSenha.SetFocus
        ValidaCampos = False
    ElseIf lstGrupo.SelCount > 1 Then
        MsgBox "Selecione apenas um Grupo por usuário.", vbExclamation + vbOKOnly, Me.Caption
        lstGrupo.SetFocus
        ValidaCampos = False
        
    Else
        
        ValidaCampos = False
        For Count = 0 To lstGrupo.ListCount - 1
            If lstGrupo.Selected(Count) Then
                ValidaCampos = True
                Exit For
            End If
        Next
        If Not ValidaCampos Then
            MsgBox "Selecione pelo menos um grupo para o usuário.", vbExclamation + vbOKOnly, Me.Caption
            lstGrupo.SetFocus
        End If
    End If
End Function
Private Sub Limpar()

    cmbLogin.Text = ""
    txtNome.Text = ""
    txtCif.Text = ""
    txtSenha.Text = ""
    txtConfSenha.Text = ""
    
End Sub
Private Sub cmbLogin_Change()

     If Trim(cmbLogin.Text) <> "" Then
          cmbLogin_Click
     Else
          Form_Load
     End If
    
End Sub
Private Sub cmbLogin_Click()
     
Dim achou As Boolean
Dim Count As Integer

     achou = False
     If rsUsuarios.RecordCount > 0 Then
          rsUsuarios.MoveFirst
     End If
     
     While (Not rsUsuarios.EOF) And (Not achou)
          If UCase(Trim(cmbLogin.Text)) = UCase(Trim(rsUsuarios!Login)) Then
               achou = True
          Else
               rsUsuarios.MoveNext
          End If
     Wend
     
     For Count = 0 To lstGrupo.ListCount - 1
          lstGrupo.Selected(Count) = False
     Next
     
     txtNome.Text = ""
     txtCif.Text = ""
     txtSenha.Text = ""
     txtConfSenha.Text = ""
     cmdAdicionar.Enabled = True
     cmdAlterar.Enabled = True
     cmdRemover.Enabled = True
    
     If achou Then
          txtNome.Text = rsUsuarios!nome
          txtSenha.Text = Decript(rsUsuarios!Senha)
          txtConfSenha.Text = txtSenha.Text
          txtCif.Text = IIf(IsNull(rsUsuarios!Cif), "", rsUsuarios!Cif)
        
          Set RsGrupoUsuario = g_cMainConnection.Execute(Proc_Selecionar.GetGrupoUsuario(rsUsuarios!IdUsuario))

          While Not RsGrupoUsuario.EOF
               For Count = 0 To UBound(CodIdGrupo)
                    If CodIdGrupo(Count) = RsGrupoUsuario!IdGrupo Then
                         lstGrupo.Selected(Count) = True
                         Exit For
                    End If
               Next
               RsGrupoUsuario.MoveNext
          Wend
          
          'Posiciona no primeiro ítem do grupo pertencente ao usuário
          For Count = 0 To lstGrupo.ListCount - 1
               If lstGrupo.Selected(Count) Then
                    lstGrupo.ListIndex = Count
                    Exit For
               End If
          Next
          lstGrupo.Refresh
          
          RsGrupoUsuario.Close
          cmdAdicionar.Enabled = False
     Else
          cmdAlterar.Enabled = False
          cmdRemover.Enabled = False
     End If
     
End Sub
Private Sub cmbLogin_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        SendKeys ("{TAB}")
    End If

End Sub

Private Sub cmbLogin_KeyPress(KeyAscii As Integer)

If KeyAscii = Asc(" ") Then KeyAscii = 0: Exit Sub

If Len(cmbLogin.Text) >= 10 And KeyAscii <> 8 Then
     If cmbLogin.SelLength >= 1 Then Exit Sub
     KeyAscii = 0
End If

End Sub

Private Sub cmdAdicionar_Click()
    
Dim Count      As Integer
Dim sMsgErro   As String

    If Not ValidaCampos Then
        Exit Sub
    End If
    
On Error GoTo ErroAdicionar

     'Abre transação para inclusão de Usuário
     g_cMainConnection.BeginTrans
     
     sMsgErro = "Erro na inclusão do usuário."
     
     g_cMainConnection.Execute (Proc_Inserir.InsereUsuario( _
                                   Mid(Trim(cmbLogin.Text), 1, 10), _
                                   Trim(txtNome.Text), _
                                   Trim(txtCif.Text), _
                                   Encript(Trim(txtSenha.Text))))
    
        
     sMsgErro = "Erro na inserção do grupo-usuário."
     
     sMsgErro = "Erro na leitura dos usuários."
     Set rsUsuarios = g_cMainConnection.Execute(Proc_Selecionar.GetUsuario(Mid(Trim(cmbLogin.Text), 1, 10)))
     
     For Count = 0 To lstGrupo.ListCount - 1
          If lstGrupo.Selected(Count) Then
               g_cMainConnection.Execute (Proc_Inserir.InsereGrupoUsuario(CodIdGrupo(Count), rsUsuarios!IdUsuario))
          End If
     Next
        
     'Finaliza transação
     g_cMainConnection.CommitTrans
    
     On Error GoTo 0
     
     Form_Unload (0)
     Form_Load
     
     Exit Sub
    
ErroAdicionar:
     'Cancela transação
     g_cMainConnection.RollbackTrans
    
     Beep
     MsgBox sMsgErro, vbCritical, Me.Caption
     Unload Me

End Sub
Private Sub cmdAlterar_Click()
    
Dim Count      As Integer
Dim sMsgErro   As String

     On Error GoTo ErroAlterar
     
     If Not ValidaCampos Then
          Exit Sub
     End If
    
     'Abre transação para atualização de Usuário
     g_cMainConnection.BeginTrans
     
     sMsgErro = "Erro na alteração do usuário."
     
     g_cMainConnection.Execute (Proc_Atualizar.AtualizaUsuario( _
                                   Mid(Trim(cmbLogin.Text), 1, 10), _
                                   Trim(txtNome.Text), _
                                   Trim(txtCif.Text), _
                                   Encript(Trim(txtSenha.Text))))

     sMsgErro = "Erro na alteração do grupo-usuário."
     g_cMainConnection.Execute (Proc_Excluir.RemoveGrupoUsuario(rsUsuarios!IdUsuario))
     
     sMsgErro = "Erro na inserção do grupo-usuário."
     
     For Count = 0 To lstGrupo.ListCount - 1
          If lstGrupo.Selected(Count) Then
               g_cMainConnection.Execute (Proc_Inserir.InsereGrupoUsuario(CodIdGrupo(Count), rsUsuarios!IdUsuario))
          End If
     Next
     
     'Finaliza transação
     g_cMainConnection.CommitTrans
     
     On Error GoTo 0
     
     Form_Unload (0)
     Form_Load
     
     Exit Sub
     
ErroAlterar:
     
     'Cancela transação
     g_cMainConnection.RollbackTrans
    
     Beep
     MsgBox sMsgErro, vbCritical, Me.Caption
     Unload Me
     
End Sub
Private Sub CmdFechar_Click()

    Unload Me
    
End Sub
Private Sub cmdRemover_Click()
     
Dim Cancel As Integer
    
     If MsgBox("Confirma exclusão do usuário?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
          On Error GoTo ErroRemove
        
          'Abre transação para excluir registro das tabelas (GrupoUsuario e Usuario)
          g_cMainConnection.BeginTrans
          
          g_cMainConnection.Execute (Proc_Excluir.RemoveGrupoUsuario(rsUsuarios!IdUsuario))
          g_cMainConnection.Execute (Proc_Excluir.RemoveUsuario(rsUsuarios!IdUsuario))
          
          'Finaliza transação
          g_cMainConnection.CommitTrans
          
          Form_Unload (Cancel)
          Form_Load

          On Error GoTo 0
     End If
     Exit Sub
    
ErroRemove:
    
    'Cancela transação
    g_cMainConnection.RollbackTrans
    
    Beep
    MsgBox "Erro na exclusão do usuário.", vbCritical, Me.Caption
    Unload Me

End Sub
Private Sub Form_Activate()
     
    lstGrupo_Click
   
End Sub

Private Sub Form_Load()

Dim sMsgErro As String

     Limpar
     cmbLogin.Clear
     lstGrupo.Clear
     cmdAdicionar.Enabled = False
     cmdAlterar.Enabled = False
     cmdRemover.Enabled = False

On Error GoTo ErroLoad
     
     sMsgErro = "Erro na leitura dos usuários."
     Set rsUsuarios = g_cMainConnection.Execute(Proc_Selecionar.GetUsuario())
    
     While Not rsUsuarios.EOF
          cmbLogin.AddItem (rsUsuarios!Login)
          rsUsuarios.MoveNext
     Wend
    
     sMsgErro = "Erro na leitura dos grupos."
     Set rsGrupos = g_cMainConnection.Execute(Proc_Selecionar.GetTodosGrupos())
    
     If Not rsGrupos.EOF Then ReDim CodIdGrupo(rsGrupos.RecordCount - 1)
    
     While Not rsGrupos.EOF
          lstGrupo.AddItem Trim(rsGrupos!Descricao)
          CodIdGrupo(lstGrupo.NewIndex) = rsGrupos!IdGrupo
          rsGrupos.MoveNext
     Wend
     
     lstGrupo_Click
     
     On Error GoTo 0
     Exit Sub
    
ErroLoad:
     Beep
     MsgBox sMsgErro, vbCritical, Me.Caption
     Unload Me
     
End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    rsUsuarios.Close
    rsGrupos.Close

End Sub
Private Sub lstGrupo_Click()

     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     ' Quando for acrescido novo ítem, adicionar no form    '
     ' LOGIN os novos acessos por usuário                   '
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     lstPermissao.Clear
     Select Case CodIdGrupo(lstGrupo.ListIndex)
          
          Case Geral.GrupoUsuario.Supervisor
               lstPermissao.AddItem "Complementação"
               lstPermissao.AddItem "Prova Zero"

               lstPermissao.AddItem "Recep - Aviso de Diferença"
               lstPermissao.AddItem "Recep - Confirmação de Remessa"
               lstPermissao.AddItem "Recep - Rejeitados"
               lstPermissao.AddItem "Recep - Movimento de Data Boa"
               lstPermissao.AddItem "Recep - Baixa de Cheques"
               lstPermissao.AddItem "Recep - Tabela de Instruções"

               lstPermissao.AddItem "Ger - Movimento para VC"
               lstPermissao.AddItem "Ger - Arquivo CEL"
               lstPermissao.AddItem "Ger - Arquivo TER"
               lstPermissao.AddItem "Ger - Movimento de Rejeitados"

               lstPermissao.AddItem "Acompanhamento Produção"
               lstPermissao.AddItem "Supervisor"
               lstPermissao.AddItem "Parâmetros"
               lstPermissao.AddItem "Cadastro de Usuários"

               lstPermissao.AddItem "Consulta"
               lstPermissao.AddItem "Relatórios"
               lstPermissao.AddItem "Cheque Data-Boa"

          Case Geral.GrupoUsuario.Digitadores
               lstPermissao.AddItem "Complementação"
               lstPermissao.AddItem "Prova Zero"

               lstPermissao.AddItem "Recep - Aviso de Diferença"
               lstPermissao.AddItem "Recep - Confirmação de Remessa"
               lstPermissao.AddItem "Recep - Rejeitados"
               lstPermissao.AddItem "Recep - Movimento de Data Boa"
               lstPermissao.AddItem "Recep - Baixa de Cheques"
               lstPermissao.AddItem "Recep - Tabela de Instruções"

               lstPermissao.AddItem "Movimento para VC"
               lstPermissao.AddItem "Arquivo CEL"
               lstPermissao.AddItem "Arquivo TER"
               lstPermissao.AddItem "Movimento de Rejeitados"

               lstPermissao.AddItem "Consulta"
               lstPermissao.AddItem "Relatórios"
               lstPermissao.AddItem "Baixa Data-Boa"
     
     End Select
End Sub
Private Sub lstGrupo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys ("{TAB}")
    End If
End Sub
Private Sub lstPermissao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys ("{TAB}")
    End If
End Sub
Private Sub txtCif_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys ("{TAB}")
    End If
End Sub
Private Sub txtConfSenha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys ("{TAB}")
    End If
End Sub
Private Sub txtNome_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys ("{TAB}")
    End If
End Sub
Private Sub txtSenha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys ("{TAB}")
    End If
End Sub
