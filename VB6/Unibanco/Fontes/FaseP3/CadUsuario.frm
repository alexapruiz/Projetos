VERSION 5.00
Begin VB.Form CadUsuario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Usuário"
   ClientHeight    =   4860
   ClientLeft      =   36
   ClientTop       =   336
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   7320
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkAtivo 
      Caption         =   "Usuário Ativo"
      Height          =   264
      Left            =   108
      TabIndex        =   19
      Top             =   2052
      Visible         =   0   'False
      Width           =   1272
   End
   Begin VB.Frame Frame3 
      Caption         =   "Permissões"
      Height          =   2400
      Left            =   3660
      TabIndex        =   18
      Top             =   2364
      Width           =   3528
      Begin VB.ListBox lstPermissao 
         Height          =   2004
         IntegralHeight  =   0   'False
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3288
      End
   End
   Begin VB.CommandButton cmdFechar 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   372
      Left            =   5760
      TabIndex        =   10
      Top             =   1440
      Width           =   1452
   End
   Begin VB.CommandButton cmdRemover 
      Caption         =   "&Remover"
      Height          =   372
      Left            =   5760
      TabIndex        =   9
      Top             =   1000
      Width           =   1452
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "A&lterar"
      Height          =   372
      Left            =   5760
      TabIndex        =   8
      Top             =   560
      Width           =   1452
   End
   Begin VB.CommandButton cmdAdicionar 
      Caption         =   "&Adicionar"
      Height          =   372
      Left            =   5760
      TabIndex        =   7
      Top             =   120
      Width           =   1452
   End
   Begin VB.Frame Frame2 
      Caption         =   "Grupos"
      Height          =   2400
      Left            =   60
      TabIndex        =   17
      Top             =   2364
      Width           =   3528
      Begin VB.ListBox lstGrupo 
         ForeColor       =   &H00800000&
         Height          =   1992
         Left            =   132
         Style           =   1  'Checkbox
         TabIndex        =   5
         Top             =   240
         Width           =   3288
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1992
      Left            =   60
      TabIndex        =   11
      Top             =   0
      Width           =   5592
      Begin VB.ComboBox cmbLogin 
         ForeColor       =   &H00800000&
         Height          =   288
         Left            =   780
         TabIndex        =   0
         Tag             =   "Login"
         Top             =   240
         Width           =   1752
      End
      Begin VB.TextBox txtCif 
         ForeColor       =   &H00800000&
         Height          =   288
         Left            =   780
         TabIndex        =   2
         Tag             =   "CIF"
         Top             =   1048
         Width           =   1752
      End
      Begin VB.TextBox txtConfSenha 
         ForeColor       =   &H00800000&
         Height          =   288
         IMEMode         =   3  'DISABLE
         Left            =   4020
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   4
         Tag             =   "Senha"
         Top             =   1440
         Width           =   1332
      End
      Begin VB.TextBox txtSenha 
         ForeColor       =   &H00800000&
         Height          =   288
         IMEMode         =   3  'DISABLE
         Left            =   780
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   3
         Tag             =   "Senha"
         Top             =   1440
         Width           =   1332
      End
      Begin VB.TextBox txtNome 
         ForeColor       =   &H00800000&
         Height          =   312
         Left            =   780
         MaxLength       =   50
         TabIndex        =   1
         Tag             =   "Nome"
         Top             =   632
         Width           =   4572
      End
      Begin VB.Label Label5 
         Caption         =   "CIF:"
         Height          =   252
         Left            =   180
         TabIndex        =   16
         Top             =   1066
         Width           =   372
      End
      Begin VB.Label Label4 
         Caption         =   "Confirmação da Senha:"
         Height          =   252
         Left            =   2220
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
Private rsUsuarios               As rdoResultset
Private rsGrupos                 As rdoResultset
Private rsGrupoUsuario           As rdoResultset
Private rsGetDescModulo          As rdoResultset

Private qryInsereUsuario         As rdoQuery
Private qrySetUsuario            As rdoQuery
Private qryRemoveUsuario         As rdoQuery
Private qryInsereGrupoUsuario    As rdoQuery
Private qryRemoveAllGrupoUsuario As rdoQuery
Private qryGetGrupoUsuario       As rdoQuery
Private qryGetAllUsuarios        As rdoQuery
Private qryGetAllGrupos          As rdoQuery
Private qryGetDescModulo         As rdoQuery
Private qryGetUsuario            As rdoQuery
Private qryGetIdCampo            As rdoQuery
Private qryInsereLogUsuario      As rdoQuery

Private ParamDiasInativo         As Long   'Qtde de Dias que usuário pode ficar inativo segundo a tabela de parâmetros
Private QtdesDiasInativo         As Long   'Qtde de Dias que usuário está Inativo

Private m_Nome                   As String
Private m_CIF                    As String
Private m_Senha                  As String
Private Function ValidaCampos() As Boolean
    Dim Count As Integer
    
    If Trim(cmbLogin.Text) = "" Then
        MsgBox "O campo Login deve ser preenchido.", vbExclamation + vbOKOnly, App.Title
        cmbLogin.SelStart = 0
        cmbLogin.SelLength = Len(cmbLogin.Text)
        cmbLogin.SetFocus
        ValidaCampos = False
    ElseIf Trim(txtNome.Text) = "" Then
        MsgBox "O campo Nome deve ser preenchido.", vbExclamation + vbOKOnly, App.Title
        txtNome.SelStart = 0
        txtNome.SelLength = Len(txtNome.Text)
        txtNome.SetFocus
        ValidaCampos = False
    ElseIf Trim(txtSenha.Text) = "" Then
        MsgBox "O campo Senha deve ser preenchido.", vbExclamation + vbOKOnly, App.Title
        txtSenha.SelStart = 0
        txtSenha.SelLength = Len(txtSenha.Text)
        txtSenha.SetFocus
        ValidaCampos = False
    ElseIf Trim(txtSenha.Text) <> Trim(txtConfSenha.Text) Then
        MsgBox "Confirmação da senha não confere.", vbExclamation + vbOKOnly, App.Title
        txtConfSenha.SelStart = 0
        txtConfSenha.SelLength = Len(txtConfSenha.Text)
        txtConfSenha.SetFocus
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
            MsgBox "Selecione pelo menos um grupo para o usuário.", vbExclamation + vbOKOnly, App.Title
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
    End If
End Sub
Private Sub cmbLogin_Click()
    Dim achou As Boolean
    Dim Count As Integer

    achou = False
    If rsUsuarios.RowCount > 0 Then
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
        txtNome.Text = rsUsuarios!Nome
        txtSenha.Text = Decript(rsUsuarios!Senha)
        txtConfSenha.Text = Decript(rsUsuarios!Senha)
        txtCif.Text = IIf(IsNull(rsUsuarios!Cif), "", rsUsuarios!Cif)
        
        '''''''''''''''''''''''''''''''
        'Memoriza os campos do usuario'
        '''''''''''''''''''''''''''''''
        m_Nome = txtNome.Text
        m_Senha = txtSenha.Text
        m_CIF = txtCif.Text
        
        qryGetGrupoUsuario.rdoParameters(0) = rsUsuarios!idUsuario
        Set rsGrupoUsuario = qryGetGrupoUsuario.OpenResultset(rdOpenStatic, rdConcurReadOnly)
        While Not rsGrupoUsuario.EOF
            For Count = 0 To lstGrupo.ListCount - 1
                If Left(lstGrupo.List(Count), 3) = rsGrupoUsuario!IdGrupo Then
                    lstGrupo.Selected(Count) = True
                    lstGrupo.ListIndex = -1
                    lstGrupo.Refresh
                End If
            Next
            rsGrupoUsuario.MoveNext
        Wend

        '* Verifica se usuário está ativo ou inativo *'
        Call DataUltimoLogon
        
        rsGrupoUsuario.Close
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
Private Sub cmdAdicionar_Click()
    
    Dim Count                   As Integer
    Dim GrupoOk                 As Boolean
    Dim Cancel                  As Integer
    Dim idUsuario               As Long
    Dim sStrGrupoDepois         As String
    
    If Not ValidaCampos Then
        Exit Sub
    End If
    
    On Error GoTo ErroAdicionar
    rdoErrors.Clear
    
    '08/05/2001''''''''''''''''''''''''''''''''''
    'Não pode aceitar cadastro do Login "DESENV"'
    '''''''''''''''''''''''''''''''''''''''''''''
    If Trim(UCase(cmbLogin.Text)) = "DESENV" Then
        MsgBox "Não é permitido o cadastramento do login 'DESENV'", vbExclamation, Me.Caption
        Exit Sub
    End If
    
    qryInsereUsuario.rdoParameters(0).Direction = rdParamReturnValue
    qryInsereUsuario.rdoParameters(1) = Mid(Trim(cmbLogin.Text), 1, 10)
    qryInsereUsuario.rdoParameters(2) = Trim(txtNome.Text)
    qryInsereUsuario.rdoParameters(3) = Trim(txtCif.Text)
    qryInsereUsuario.rdoParameters(4) = Encript(Trim(txtSenha.Text))
    qryInsereUsuario.Execute
    
    If qryInsereUsuario.rdoParameters(0) <> 0 Then
        MsgBox "Erro na inserção do usuário.", vbCritical + vbOKOnly, App.Title
    Else
        GrupoOk = True
        For Count = 0 To lstGrupo.ListCount - 1
            If lstGrupo.Selected(Count) Then
                qryInsereGrupoUsuario.rdoParameters(1) = Mid(Trim(cmbLogin.Text), 1, 10)
                qryInsereGrupoUsuario.rdoParameters(2) = Left(lstGrupo.List(Count), 3)
                qryInsereGrupoUsuario.Execute
                If qryInsereGrupoUsuario.rdoParameters(0) <> 0 Then
                    MsgBox "Erro na inserção do grupo-usuário.", vbCritical + vbOKOnly, App.Title
                    GrupoOk = False
                    Exit For
                End If
            End If
        Next
        
        'Grava log de usuário
        idUsuario = GetUsuario(cmbLogin.Text)
        
        If idUsuario = 0 Then
            On Error Resume Next
            Exit Sub
        End If
        'Call GravaLog(idUsuario, 0, 240)
        
        sStrGrupoDepois = ResolveMapa(idUsuario, eRetornaMapa)
        
        '''''''''''''''''''
        'Insere LogUsuario'
        '''''''''''''''''''
        With qryInsereLogUsuario
            .rdoParameters(0) = idUsuario
            .rdoParameters(1) = 0
            .rdoParameters(2) = 240 'Cadastramento de usuário
            .rdoParameters(3) = Geral.idUsuario
            .rdoParameters(4) = ""
            .rdoParameters(5) = sStrGrupoDepois
            .Execute
        End With
        
        
        If GrupoOk Then
            Form_Unload (Cancel)
            Form_Load
        End If
        
        
    End If
    On Error Resume Next
    Exit Sub
    
ErroAdicionar:
    Select Case TratamentoErro("Erro na inserção de usuário.", Err, rdoErrors)
        Case vbCancel
            Unload Me
        Case vbRetry
            Resume
    End Select

End Sub
Private Sub cmdAlterar_Click()
    Dim Count               As Integer
    Dim GrupoOk             As Boolean
    Dim Cancel              As Integer
    Dim idUsuario           As Long
    Dim RsAux               As RDO.rdoResultset
    Dim sStrGrupoAntes      As String
    Dim sStrGrupoDepois     As String
    
    If Not ValidaCampos Then
        Exit Sub
    End If
    
    On Error GoTo ErroAlterar
    rdoErrors.Clear
    
    qrySetUsuario.rdoParameters(1) = Mid(Trim(cmbLogin.Text), 1, 10)
    qrySetUsuario.rdoParameters(2) = Trim(txtNome.Text)
    qrySetUsuario.rdoParameters(3) = Trim(txtCif.Text)
    qrySetUsuario.rdoParameters(4) = Encript(Trim(txtSenha.Text))
    qrySetUsuario.Execute
    
    ''''''''''''''''''''''''''''''
    'Compara os campos do usuario'
    ''''''''''''''''''''''''''''''
    If m_Nome <> txtNome.Text Then
        ''''''''''''''''''''''''''''''''''''''
        'O Nome do campo está no próprio Text'
        ''''''''''''''''''''''''''''''''''''''
        qryGetIdCampo.rdoParameters(0) = txtNome.Tag
        Set RsAux = qryGetIdCampo.OpenResultset(rdOpenStatic, rdConcurReadOnly)
    
        If RsAux.EOF Then
            MsgBox "Erro ao comparar os campos do usuário.", vbExclamation
        Else
            With qryInsereLogUsuario
                .rdoParameters(0) = rsUsuarios!idUsuario
                .rdoParameters(1) = RsAux!IdCampo
                .rdoParameters(2) = 241 'Alteração de usuário
                .rdoParameters(3) = Geral.idUsuario
                .rdoParameters(4) = ""
                .rdoParameters(5) = ""
                .Execute
            End With
        End If
    End If
    If m_CIF <> txtCif.Text Then
        ''''''''''''''''''''''''''''''''''''''
        'O Nome do campo está no próprio Text'
        ''''''''''''''''''''''''''''''''''''''
        qryGetIdCampo.rdoParameters(0) = txtCif.Tag
        Set RsAux = qryGetIdCampo.OpenResultset(rdOpenStatic, rdConcurReadOnly)
    
        If RsAux.EOF Then
            MsgBox "Erro ao comparar os campos do usuário.", vbExclamation
        Else
            With qryInsereLogUsuario
                .rdoParameters(0) = rsUsuarios!idUsuario
                .rdoParameters(1) = RsAux!IdCampo
                .rdoParameters(2) = 241 'Alteração de usuário
                .rdoParameters(3) = Geral.idUsuario
                .rdoParameters(4) = ""
                .rdoParameters(5) = ""
                .Execute
            End With
        End If
    End If
    If m_Senha <> txtSenha.Text Then
        ''''''''''''''''''''''''''''''''''''''
        'O Nome do campo está no próprio Text'
        ''''''''''''''''''''''''''''''''''''''
        qryGetIdCampo.rdoParameters(0) = txtSenha.Tag
        Set RsAux = qryGetIdCampo.OpenResultset(rdOpenStatic, rdConcurReadOnly)
    
        If RsAux.EOF Then
            MsgBox "Erro ao comparar os campos do usuário.", vbExclamation
        Else
            With qryInsereLogUsuario
                .rdoParameters(0) = rsUsuarios!idUsuario
                .rdoParameters(1) = RsAux!IdCampo
                .rdoParameters(2) = 241 'Alteração de usuário
                .rdoParameters(3) = Geral.idUsuario
                .rdoParameters(4) = ""
                .rdoParameters(5) = ""
                .Execute
            End With
        End If
    End If
    If ChkAtivo.Value = Checked And DataUltimoLogon = False Then
        Call AtualizaUltimoLogon
    End If
    
    If qrySetUsuario.rdoParameters(0) <> 0 Then
        MsgBox "Erro na alteração do usuário.", vbCritical + vbOKOnly, App.Title
    Else
        sStrGrupoAntes = ResolveMapa(rsUsuarios!idUsuario, eRetornaMapa)
        
        qryRemoveAllGrupoUsuario.rdoParameters(1) = rsUsuarios!idUsuario
        qryRemoveAllGrupoUsuario.Execute
        If qryRemoveAllGrupoUsuario.rdoParameters(0) <> 0 Then
            MsgBox "Erro na alteração do grupo-usuário.", vbCritical + vbOKOnly, App.Title
        Else
        
            GrupoOk = True
            For Count = 0 To lstGrupo.ListCount - 1
                If lstGrupo.Selected(Count) Then
                    qryInsereGrupoUsuario.rdoParameters(1) = Mid(Trim(cmbLogin.Text), 1, 10)
                    qryInsereGrupoUsuario.rdoParameters(2) = Left(lstGrupo.List(Count), 3)
                    qryInsereGrupoUsuario.Execute
                    If qryInsereGrupoUsuario.rdoParameters(0) <> 0 Then
                        MsgBox "Erro na inserção do grupo-usuário.", vbCritical + vbOKOnly, App.Title
                        GrupoOk = False
                        Exit For
                    End If
                End If
            Next
        
            sStrGrupoDepois = ResolveMapa(rsUsuarios!idUsuario, eRetornaMapa)
            If sStrGrupoAntes <> sStrGrupoDepois Then
                With qryInsereLogUsuario
                    .rdoParameters(0) = rsUsuarios!idUsuario
                    .rdoParameters(1) = 0
                    .rdoParameters(2) = 241 'Alteração de usuário
                    .rdoParameters(3) = Geral.idUsuario
                    .rdoParameters(4) = sStrGrupoAntes
                    .rdoParameters(5) = sStrGrupoDepois
                    .Execute
                End With
            End If
            
            'Grava log de usuário
            idUsuario = GetUsuario(cmbLogin.Text)
            If idUsuario = 0 Then
                On Error Resume Next
                Exit Sub
            End If
            'Call GravaLog(idUsuario, 0, 241)
            
            If GrupoOk Then
                Form_Unload (Cancel)
                Form_Load
            End If
        End If
    End If
    On Error Resume Next
    Exit Sub
ErroAlterar:
    Select Case TratamentoErro("Erro na atualização de usuário.", Err, rdoErrors)
        Case vbCancel
            Unload Me
        Case vbRetry
            Resume
    End Select

End Sub
Private Sub CmdFechar_Click()
    Unload Me
End Sub
Private Sub cmdRemover_Click()
    
    Dim Cancel      As Integer
    Dim idUsuario   As Long
    
    If MsgBox("Confirma exclusão do usuário?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        On Error GoTo ErroRemove
        rdoErrors.Clear
        
        'Obtem o IdUsuário em exclusâo
        idUsuario = GetUsuario(cmbLogin.Text)
        
        qryRemoveUsuario.rdoParameters(1) = rsUsuarios!idUsuario
        qryRemoveUsuario.Execute
        If qryRemoveUsuario.rdoParameters(0) <> 0 Then
            MsgBox "Erro na exclusão do usuário.", vbCritical + vbOKOnly, App.Title
        Else
            'Grava log de usuário
            If idUsuario = 0 Then
                On Error Resume Next
                Exit Sub
            End If
            'Call GravaLog(idUsuario, 0, 242)
            
            '''''''''''''''''''
            'Insere LogUsuario'
            '''''''''''''''''''
            With qryInsereLogUsuario
                .rdoParameters(0) = idUsuario
                .rdoParameters(1) = 0
                .rdoParameters(2) = 242 'Exclusão de usuário
                .rdoParameters(3) = Geral.idUsuario
                .rdoParameters(4) = ""
                .rdoParameters(5) = ""
                .Execute
            End With
            
            
        
            Form_Unload (Cancel)
            Form_Load
        End If
        On Error Resume Next
    End If
    Exit Sub
ErroRemove:
    Select Case TratamentoErro("Erro na exclusão de usuário.", Err, rdoErrors)
        Case vbCancel
            Unload Me
        Case vbRetry
            Resume
    End Select

End Sub
Private Sub Form_Activate()
   'Inclusão de chamada a rotina AtualizaAtividade
   Call AtualizaAtividade(23)
   lstGrupo_Click
End Sub
Private Sub Form_Load()
    Limpar
    cmbLogin.Clear
    lstGrupo.Clear
    cmdAdicionar.Enabled = False
    cmdAlterar.Enabled = False
    cmdRemover.Enabled = False
    ChkAtivo.Visible = False
    ChkAtivo.Value = Unchecked
            
    Set qryInsereUsuario = Geral.Banco.CreateQuery("", "{? = call InsereUsuario (?,?,?,?)}")
    Set qrySetUsuario = Geral.Banco.CreateQuery("", "{? = call AtualizaUsuario (?,?,?,?)}")
    Set qryRemoveUsuario = Geral.Banco.CreateQuery("", "{? = call RemoveUsuario (?)}")
    Set qryInsereGrupoUsuario = Geral.Banco.CreateQuery("", "{? = call InsereGrupoUsuario (?,?)}")
    Set qryRemoveAllGrupoUsuario = Geral.Banco.CreateQuery("", "{? = call RemoveAllGrupoUsuario (?)}")
    Set qryGetGrupoUsuario = Geral.Banco.CreateQuery("", "{call GetGrupoUsuario (?)}")
    Set qryGetDescModulo = Geral.Banco.CreateQuery("", "{call GetDescModulo (?)}")
    Set qryGetAllUsuarios = Geral.Banco.CreateQuery("", "{call GetAllUsuarios }")
    Set qryGetAllGrupos = Geral.Banco.CreateQuery("", "{call GetAllGrupos }")
    Set qryGetUsuario = Geral.Banco.CreateQuery("", "{call GetUsuario (?) }")
    Set qryGetIdCampo = Geral.Banco.CreateQuery("", "{call GetIdCampo (?) }")
    Set qryInsereLogUsuario = Geral.Banco.CreateQuery("", "{call InsereLogUsuario (?,?,?,?,?,?) }")
    
    On Error GoTo ErroUsuarios
    rdoErrors.Clear
    
    Set rsUsuarios = qryGetAllUsuarios.OpenResultset(rdOpenStatic, rdConcurReadOnly)
    While Not rsUsuarios.EOF
        cmbLogin.AddItem (rsUsuarios!Login)
        rsUsuarios.MoveNext
    Wend
    
    On Error GoTo ErroGrupos
    rdoErrors.Clear
    Set rsGrupos = qryGetAllGrupos.OpenResultset(rdOpenStatic, rdConcurReadOnly)
    While Not rsGrupos.EOF
        lstGrupo.AddItem (rsGrupos!IdGrupo & " - " & rsGrupos!Descricao)
        rsGrupos.MoveNext
    Wend
    On Error Resume Next
    Exit Sub
    
ErroUsuarios:
    Select Case TratamentoErro("Erro na leitura dos usuários.", Err, rdoErrors)
        Case vbCancel
            Unload Me
        Case vbRetry
            Resume
    End Select
    Exit Sub
ErroGrupos:
    Select Case TratamentoErro("Erro na leitura dos grupos.", Err, rdoErrors)
        Case vbCancel
            Unload Me
        Case vbRetry
            Resume
    End Select
End Sub
Private Sub Form_Unload(Cancel As Integer)
    rsUsuarios.Close
    rsGrupos.Close
    qryGetGrupoUsuario.Close
    qryGetAllUsuarios.Close
    qryGetAllGrupos.Close
    qryInsereGrupoUsuario.Close
    qryInsereUsuario.Close
    qryRemoveUsuario.Close
    qryRemoveAllGrupoUsuario.Close
    qryGetUsuario.Close
    
End Sub
Private Sub lstGrupo_Click()
    lstPermissao.Clear
    
    qryGetDescModulo.rdoParameters(0) = Left(lstGrupo.List(lstGrupo.ListIndex), 3)
        
    Set rsGetDescModulo = qryGetDescModulo.OpenResultset(rdOpenStatic, rdConcurReadOnly)
    
    Do Until rsGetDescModulo.EOF
       lstPermissao.AddItem rsGetDescModulo!Descricao
       rsGetDescModulo.MoveNext
    Loop
    
  '  Select Case UCase(Left(lstGrupo.List(lstGrupo.ListIndex), 3))
  '      Case "TER"
  '          lstPermissao.AddItem "Recepção"
  '          lstPermissao.AddItem "Cadastro"
  '          lstPermissao.AddItem "Complementação"
  '      Case "REC"
  '          lstPermissao.AddItem "Recepção"
  '          lstPermissao.AddItem "Cadastro"
  '          lstPermissao.AddItem "Registro de Ocorrência"
  '      Case "DIG"
  '          lstPermissao.AddItem "Captura"
  '          lstPermissao.AddItem "Controle de Qualidade"
  '          lstPermissao.AddItem "Recaptura"
  '          lstPermissao.AddItem "Complementação"
  '          lstPermissao.AddItem "Expedição"
  '          lstPermissao.AddItem "Confirmação de Agência / Conta"
  '      Case "PES"
  '          lstPermissao.AddItem "Consultas e Relatórios"
  '      Case "AUX"
  '          lstPermissao.AddItem "Recepção"
  '          lstPermissao.AddItem "Cadastro"
  '          lstPermissao.AddItem "Registro de Ocorrência"
  '          lstPermissao.AddItem "Captura"
  '          lstPermissao.AddItem "Controle de Qualidade"
  '          lstPermissao.AddItem "Recaptura"
  '          lstPermissao.AddItem "Complementação"
  '          lstPermissao.AddItem "Ilegíveis"
  '          lstPermissao.AddItem "Prova Zero"
  '          lstPermissao.AddItem "Expedição"
  '          lstPermissao.AddItem "Acompanhamento da Produção"
  '          lstPermissao.AddItem "Acompanhamento de Atividades"
  '          lstPermissao.AddItem "Acompanhamento de Recepção"
  '          lstPermissao.AddItem "Acompanhamento de Expedição"
  '          lstPermissao.AddItem "Acompanhamento de Usuarios"
  '          lstPermissao.AddItem "Auditoria"
  '          lstPermissao.AddItem "Finalização de Capas"
  '          lstPermissao.AddItem "Confirmação de Agência / Conta"
  '          lstPermissao.AddItem "Correção de Agência / Conta"
  '          lstPermissao.AddItem "Troca de Ordem"
  '          lstPermissao.AddItem "Consultas e Relatórios"
  '      Case "SUP", "SPT"
  '          lstPermissao.AddItem "Recepção"
  '          lstPermissao.AddItem "Cadastro"
  '          lstPermissao.AddItem "Registro de Ocorrência"
  '          lstPermissao.AddItem "Captura"
  '          lstPermissao.AddItem "Controle de Qualidade"
  '          lstPermissao.AddItem "Recaptura"
  '          lstPermissao.AddItem "Complementação"
  '          lstPermissao.AddItem "Ilegíveis"
  '          lstPermissao.AddItem "Prova Zero"
  '          lstPermissao.AddItem "Expedição"
  '          lstPermissao.AddItem "Acompanhamento da Produção"
  '          lstPermissao.AddItem "Acompanhamento de Atividades"
  '          lstPermissao.AddItem "Acompanhamento de Recepção"
  '          lstPermissao.AddItem "Acompanhamento de Expedição"
  '          lstPermissao.AddItem "Acompanhamento de Usuarios"
  '          lstPermissao.AddItem "Alçada"
  '          lstPermissao.AddItem "Vínculo"
  '          lstPermissao.AddItem "Auditoria"
  '          lstPermissao.AddItem "Finalização de Capas"
  '          lstPermissao.AddItem "Confirmação de Agência / Conta"
  '          lstPermissao.AddItem "Correção de Agência / Conta"
  '          lstPermissao.AddItem "Troca de Ordem"
  '          lstPermissao.AddItem "Exclusão"
  '          lstPermissao.AddItem "Parametros do Sistema"
  '          lstPermissao.AddItem "Cadastro de Usuários"
  '          lstPermissao.AddItem "Consultas e Relatórios"
  '      Case "LID"
  '          lstPermissao.AddItem "Recepção"
  '          lstPermissao.AddItem "Cadastro"
  '          lstPermissao.AddItem "Registro de Ocorrência"
  '          lstPermissao.AddItem "Captura"
  '          lstPermissao.AddItem "Controle de Qualidade"
  '          lstPermissao.AddItem "Recaptura"
  '          lstPermissao.AddItem "Complementação"
  '          lstPermissao.AddItem "Ilegíveis"
  '          lstPermissao.AddItem "Prova Zero"
  '          lstPermissao.AddItem "Expedição"
  '          lstPermissao.AddItem "Acompanhamento da Produção"
  '          lstPermissao.AddItem "Acompanhamento de Atividades"
  '          lstPermissao.AddItem "Acompanhamento de Recepção"
  '          lstPermissao.AddItem "Acompanhamento de Usuarios"
  '          lstPermissao.AddItem "Alçada"
  '          lstPermissao.AddItem "Vínculo"
  '          lstPermissao.AddItem "Auditoria"
  '          lstPermissao.AddItem "Confirmação de Agência / Conta"
  '          lstPermissao.AddItem "Troca de Ordem"
  '          lstPermissao.AddItem "Exclusão"
  '          lstPermissao.AddItem "Consultas e Relatórios"
  '  End Select
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
Function DataUltimoLogon() As Boolean
'* Verifica se usuario esta Ativo / Inativo e a qtde de dias que está Inativo *'

On Error GoTo TrataErro

    Dim qryDataUltimoLogon  As rdoQuery     'Traz a Qtde de Dias que Usuário não se loga no Sistema
    Dim rsDataUltimoLogon   As rdoResultset 'Recordset
    
    Set qryDataUltimoLogon = Geral.Banco.CreateQuery("", "{Call GetDataUltimoLogon(?)}")
    
    With qryDataUltimoLogon
        .rdoParameters(0) = cmbLogin.Text
        Set rsDataUltimoLogon = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With

    If Not rsDataUltimoLogon.EOF Then
        '* Qtde de dias que usuário está inativo
        QtdesDiasInativo = rsDataUltimoLogon!QtdeDiasInativo
        
        '*
            rsDataUltimoLogon.MoreResults
        '*
        
        '* Qtde de dias que usuário poderá ficar inativo
        ParamDiasInativo = rsDataUltimoLogon!Diasinativo
    End If

    ' Se a qtde de Dias Inativo for maior que a qtde máxima permitida de dias
    ' Inativo usuário não poderá se logar no sistema,  sem  a autorização  do
    ' Supervisor ou do Suporte.
    If QtdesDiasInativo > ParamDiasInativo Then
        ChkAtivo.Value = Unchecked
        ChkAtivo.Visible = True
        DataUltimoLogon = False
    Else
        ChkAtivo.Visible = False
        DataUltimoLogon = True
    End If
    
Exit Function

TrataErro:
    Select Case TratamentoErro("Não foi possível verificar se usuário está Ativo ou Inativo.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
End Select
End Function
Private Sub AtualizaUltimoLogon()
'* Atualiza Data/Hora do Ultimo Logon do Usuário *'

On Error GoTo TrataErro

    Dim qryAtualizaData     As rdoQuery     'Atualiza com a data atual o último logon do usuário se ele estiver ativo
    Dim rsAtualizaData      As rdoResultset 'Recordset
    
    Set qryAtualizaData = Geral.Banco.CreateQuery("", "{? =Call AtualizaDataUltimoLogon(?)}")

    With qryAtualizaData
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1) = cmbLogin.Text
        Set rsAtualizaData = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
        
        '* Se Parametro de Retorno for <> 0 ocorreu erro de atualização*'
        If .rdoParameters(0) <> 0 Then
            GoTo TrataErro
        End If
    End With
    
Exit Sub
TrataErro:
    Select Case TratamentoErro("Não foi possível atualizar última data de Logon.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
End Select
End Sub
Private Function GetUsuario(LoginName As String) As Long
    
    Dim RsGetUsuario    As rdoResultset
    
    On Error GoTo Err_GetUsuario
    
    GetUsuario = 0
    
    With qryGetUsuario
        .rdoParameters(0) = Mid(Trim(LoginName), 1, 10)
        Set RsGetUsuario = .OpenResultset(rdOpenStatic, rdConcurReadOnly)
    End With
    
    If RsGetUsuario.EOF Then GoTo Exit_GetUsuario

    GetUsuario = RsGetUsuario("IdUsuario")
    
Exit_GetUsuario:
    If Not (RsGetUsuario Is Nothing) Then Set RsGetUsuario = Nothing
    Exit Function
    
Err_GetUsuario:

    Beep
    MsgBox "Erro na leitura de dados do usuário.", vbCritical + vbOKOnly, App.Title
    GoTo Exit_GetUsuario

End Function
