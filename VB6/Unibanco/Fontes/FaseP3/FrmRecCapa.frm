VERSION 5.00
Begin VB.Form FrmRecCapa 
   Caption         =   "Cadastro de Capas"
   ClientHeight    =   3924
   ClientLeft      =   2016
   ClientTop       =   2712
   ClientWidth     =   7548
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3924
   ScaleWidth      =   7548
   Begin VB.Frame Frame4 
      Caption         =   "Quantidade"
      Height          =   1632
      Left            =   5664
      TabIndex        =   16
      Top             =   1440
      Width           =   1764
      Begin VB.Label Label4 
         Caption         =   "Recepcionada"
         Height          =   252
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1332
      End
      Begin VB.Label lblContaEnvMal 
         Alignment       =   2  'Center
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   852
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   1572
      End
   End
   Begin VB.CommandButton cmdFechar 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   372
      Left            =   4668
      TabIndex        =   7
      Top             =   3360
      Width           =   1512
   End
   Begin VB.CommandButton cmdEncerrar 
      Caption         =   "Encerrar &Agência"
      Height          =   372
      Left            =   2868
      TabIndex        =   6
      Top             =   3360
      Width           =   1512
   End
   Begin VB.CommandButton CmdRecepcionar 
      Caption         =   "Recepcionar"
      Height          =   372
      Left            =   1224
      TabIndex        =   5
      Top             =   3360
      Width           =   1512
   End
   Begin VB.Frame Frame3 
      Height          =   1632
      Left            =   120
      TabIndex        =   13
      Top             =   1440
      Width           =   5424
      Begin VB.TextBox txtNumMalote 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   372
         Left            =   2592
         MaxLength       =   12
         TabIndex        =   4
         Top             =   960
         Width           =   1812
      End
      Begin VB.TextBox txtCapa 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   372
         Left            =   2592
         MaxLength       =   14
         TabIndex        =   3
         Top             =   480
         Width           =   1812
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Numero do Malote :"
         Enabled         =   0   'False
         Height          =   192
         Left            =   888
         TabIndex        =   15
         Top             =   1080
         Width           =   1404
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Capa do Envelope :"
         Height          =   192
         Left            =   888
         TabIndex        =   14
         Top             =   600
         Width           =   1428
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo"
      Height          =   552
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   4212
      Begin VB.OptionButton optTipo 
         Caption         =   "Finin&vest"
         Height          =   192
         Index           =   2
         Left            =   3120
         TabIndex        =   19
         Top             =   240
         Width           =   924
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "&Envelope"
         Height          =   192
         Index           =   0
         Left            =   180
         TabIndex        =   0
         Top             =   240
         Width           =   1092
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "&Malote Empresa"
         Height          =   192
         Index           =   1
         Left            =   1356
         TabIndex        =   1
         Top             =   240
         Width           =   1452
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   432
      Left            =   4428
      ScaleHeight     =   384
      ScaleWidth      =   2964
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   120
      Width           =   3012
      Begin VB.Label lblTipo 
         Alignment       =   2  'Center
         Caption         =   "Envelope ou Malote"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   312
         Left            =   60
         TabIndex        =   11
         Top             =   0
         Width           =   2832
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Agência de Origem"
      Height          =   672
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   7320
      Begin VB.ComboBox cmbAgencia 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   336
         Left            =   504
         TabIndex        =   2
         Text            =   "0001"
         Top             =   240
         Width           =   1032
      End
      Begin VB.Label lblAgencia 
         Caption         =   "Nome da Agência de Origem"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   312
         Left            =   1968
         TabIndex        =   9
         Top             =   240
         Width           =   4692
      End
   End
End
Attribute VB_Name = "FrmRecCapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsAllAgenf          As rdoResultset
Private rsGetAgencia        As rdoResultset
Private rsAtualizaAgencia   As rdoResultset

Private qryGetAllAgenf      As rdoQuery
Private qryGetAgencia       As rdoQuery
Private qryInsereCapa       As rdoQuery
Private qryAtualizaAgencia  As rdoQuery

Private IdEnv_Mal           As String       'Identificação de Envelope / Malote

Private Agencia             As Integer      'Agência Atual
Private ContQtdeEnvMal      As Integer      'Conta quantidade de Envelopes
Private ContGravada         As Integer      'Conta quantidade gravada

Option Explicit
Private Sub cmbAgencia_Change()
    cmbAgencia_Click
End Sub
Private Sub cmbAgencia_Click()
    
    cmbAgencia.Text = Left(cmbAgencia.Text, 4)
    lblAgencia.Caption = ""
    
    If rsAllAgenf.RowCount > 0 Then
        rsAllAgenf.MoveFirst
    End If
    
    Do While Not rsAllAgenf.EOF
        If Format(cmbAgencia.Text, "0000") = Format(rsAllAgenf!agefscdagen, "0000") Then
            lblAgencia.Caption = rsAllAgenf!agefsnoagen
            Exit Do
        End If
        rsAllAgenf.MoveNext
    Loop

End Sub
Private Sub cmbAgencia_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys ("{TAB}")
    End If
End Sub
Private Sub cmbAgencia_LostFocus()
    
    If Len(Trim(cmbAgencia.Text)) = 0 Then Exit Sub
    
    If IsNumeric(cmbAgencia.Text) = False Then
        cmbAgencia.Text = ""
        cmbAgencia.SetFocus
    Else
        '* Verificação de Agencia e quantidade de capas *'
        If Agencia <> CInt(cmbAgencia.Text) Then
            ContQtdeEnvMal = 0
            lblContaEnvMal.Caption = "00"
            Agencia = cmbAgencia.Text
        End If
    End If
    
End Sub
Private Sub cmdEncerrar_Click()
    
    Dim StrMsg As String
    
    StrMsg = "Confirma encerramento da recepção da agência " & Agencia & "?"
    
    If MsgBox(StrMsg, vbQuestion + vbYesNo, App.Title) = vbYes Then
        LimpaCampos
        cmbAgencia.SetFocus
    Else
        txtCapa.SetFocus
    End If
    
End Sub
Private Sub CmdFechar_Click()
    Unload Me
End Sub
Private Sub SalvaReg()
    Dim IdCapa As Integer
    
    On Error GoTo ErroCapa
       
    rdoErrors.Clear
    
    CmdRecepcionar.SetFocus
    
    If CriticaAg = False Then Exit Sub
    
    If CriticaEnvMal = False Then Exit Sub
    
    'Identifica Envelope / Malote ou Fininvest
    
    If optTipo(0).Value Then        '*** Envelope ***
        IdEnv_Mal = "E"
    ElseIf optTipo(2).Value Then    '*** Fininvest ***
        IdEnv_Mal = "F"
    Else
        IdEnv_Mal = "M"             '*** Malote ***
    End If
    
    If VerificaAgencia = False Then Exit Sub
        
    With qryInsereCapa
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = 0                    ' IdLote
        .rdoParameters(3) = IdEnv_Mal
        .rdoParameters(4) = Val(txtCapa.Text)
        .rdoParameters(5) = IIf(optTipo(1).Value, Val(txtNumMalote.Text), 0)
        .rdoParameters(6) = Val(cmbAgencia.Text)
        .rdoParameters(7) = "0"                  ' Status
        .rdoParameters(8).Direction = rdParamOutput
        .rdoParameters(9) = "N"
        .Execute
        
        If .rdoParameters(0) = 1 Then

            MsgBox "Atenção! Não é possível Recepcionar este " & _
                Switch(optTipo(0).Value, "Envelope", optTipo(1).Value, "Malote", optTipo(2).Value, "Envelope Fininvest") & _
                ", pois já está recepcionado no sistema.", _
                vbExclamation + vbOKOnly, App.Title
            txtCapa.SetFocus
            Exit Sub
        ElseIf .rdoParameters(0) > 1 Then
            MsgBox "Erro na recepção de Envelope/Malote/Fininvest", vbCritical + vbOKOnly, App.Title
            Exit Sub
        End If
        
    End With
    
    'Gravar Log
    Call GravaLog(qryInsereCapa.rdoParameters(8).Value, 0, 20)
      
   'Cria Controle de Quantidade de Capas
    ContQtdeEnvMal = ContQtdeEnvMal + 1
       
      
    If ContQtdeEnvMal < 10 Then
        lblContaEnvMal = "0" & ContQtdeEnvMal
    Else
        lblContaEnvMal = ContQtdeEnvMal
    End If
    
    Call AtualizaInformacaoAgencia
    
    If optTipo(0).Value Then        '*** Envelope ***
        IdEnv_Mal = "E"
        txtCapa.Text = ""
    ElseIf optTipo(2).Value Then        '*** Fininvest ***
        IdEnv_Mal = "F"
        txtCapa.Text = ""
    Else                            '*** Malote ***
        IdEnv_Mal = "M"
        txtCapa.Text = ""
        txtNumMalote.Text = ""
    End If
    
    txtCapa.SetFocus

Exit Sub

ErroLog:
    Select Case TratamentoErro("Erro na atualização do Log de operação.", Err, rdoErrors)
        Case vbCancel
            Unload Me
        Case vbRetry
            Resume
    End Select
    Exit Sub
    
ErroCapa:
    Select Case TratamentoErro("Erro na recepção de Envelope/Malote.", Err, rdoErrors)
        Case vbCancel
            Unload Me
        Case vbRetry
            Resume
    End Select
End Sub
Private Sub CmdRecepcionar_Click()

    Call SalvaReg
    
End Sub
Private Sub Form_Activate()

   'Inclusão de chamada a rotina AtualizaAtividade
   Call AtualizaAtividade(3)
   
End Sub
Private Sub Form_Load()

    Set qryGetAllAgenf = Geral.Banco.CreateQuery("", "{ call GetAllAgenf }")
    Set qryGetAgencia = Geral.Banco.CreateQuery("", "{call GetAgencia (?,?,?) }")
    
    Set qryInsereCapa = Geral.Banco.CreateQuery("", "{ ? = call InsereCapa (?,?,?,?,?,?,?,?,?) }")
  
    Set qryAtualizaAgencia = Geral.Banco.CreateQuery("", "{call AtualizaAgencia(?,?,?)}")
    
    'Lê Tb. Agencia p/ preencher combo
    Set rsAllAgenf = qryGetAllAgenf.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    
    While Not rsAllAgenf.EOF
        cmbAgencia.AddItem (Format(rsAllAgenf!agefscdagen, "0000"))
        rsAllAgenf.MoveNext
    Wend

    optTipo(0).Value = True

End Sub
Private Sub Form_Unload(Cancel As Integer)
    rsAllAgenf.Close
    qryGetAllAgenf.Close
    qryInsereCapa.Close
End Sub

Private Sub optTipo_Click(Index As Integer)

    Call LimpaCampos
    
    lblAgencia.Caption = ""
    
    If optTipo(0).Value Then            '*** Envelope ***
        lblTipo.Caption = "Envelope"
        Label1.Caption = "Capa de Envelope"
        Label2.Enabled = False
        txtNumMalote.Enabled = False
        txtCapa.MaxLength = 10
    ElseIf optTipo(2).Value Then        '*** Envelope Fininvest ***
        lblTipo.Caption = "Envelope Fininvest"
        Label1.Caption = "Capa de Envelope"
        Label2.Enabled = False
        txtNumMalote.Enabled = False
        txtCapa.MaxLength = 10
    Else                                '*** Malote ***
        lblTipo.Caption = "Malote Empresa"
        Label1.Caption = "Capa de Malote"
        Label2.Enabled = True
        txtCapa.MaxLength = 14
        txtNumMalote.Enabled = True
    End If
    
    SendKeys ("{TAB}")

End Sub
Private Sub LimpaCampos()
    txtCapa.Text = ""
    txtNumMalote.Text = ""
    cmbAgencia.Text = ""
    ContQtdeEnvMal = 0
    lblContaEnvMal.Caption = "00"
End Sub
Private Sub optTipo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys ("{TAB}")
    End If
End Sub
Private Sub txtCapa_GotFocus()
    With txtCapa
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtcapa_KeyPress(KeyAscii As Integer)
    If txtNumMalote.Enabled Then
        If KeyAscii = vbKeyReturn Then
            SendKeys ("{TAB}")
        End If
    Else
        If KeyAscii = vbKeyReturn Then
            Call SalvaReg
        End If
    End If
End Sub
Private Function CriticaAg() As Boolean
'* Cria Função para validar Agência informada *'

    Dim RetAgencia As Integer

    CriticaAg = False
    
    If Trim(cmbAgencia.Text) = "" Then
        MsgBox "O campo Agência deve ser informado.", vbExclamation + vbOKOnly, App.Title
        cmbAgencia.SetFocus
        CriticaAg = False
        Exit Function
    Else
      'Validar Agencia
      RetAgencia = ValidaAgencia(Val(cmbAgencia.Text), "", False)

      'Verificar Retorno da Função
      Select Case RetAgencia
        Case 2
          'Agencia em Feriado
          MsgBox "A agência de origem está em feriado.", vbInformation + vbOKOnly, App.Title
          cmbAgencia.SetFocus
          CriticaAg = False
          Exit Function
        Case 3
          'Agencia Fechada
          MsgBox "A agência de origem está fechada.", vbInformation + vbOKOnly, App.Title
          cmbAgencia.SetFocus
          CriticaAg = False
          Exit Function
        Case 4
          'Agencia não Cadastrada
          MsgBox "A agência de origem não está cadastrada.", vbInformation + vbOKOnly, App.Title
          cmbAgencia.SetFocus
          CriticaAg = False
          Exit Function
      End Select
    End If
    
    CriticaAg = True
    
End Function
Private Function CriticaEnvMal() As Boolean
'* Função para Validação de Numero de Envelope / Malote *'

    Dim iLenMalote As Double
    
    CriticaEnvMal = False

    If Trim(txtCapa.Text) = "" Then
        MsgBox "O campo Capa deve ser informado.", vbExclamation + vbOKOnly, App.Title
        txtCapa.SetFocus
        CriticaEnvMal = False
        Exit Function
    ElseIf Not IsNumeric(txtCapa.Text) Or (Val(txtCapa.Text) < 0) Then
        MsgBox "O campo Capa está inválido.", vbExclamation + vbOKOnly, App.Title
        txtCapa.SetFocus
        CriticaEnvMal = False
        Exit Function
    ElseIf (Len(Trim(Val(txtCapa.Text))) < 2) Then
        MsgBox "O campo Capa deve ter no mínimo dois dígitos.", vbExclamation + vbOKOnly, App.Title
        txtCapa.SetFocus
        CriticaEnvMal = False
        Exit Function
    ElseIf optTipo(0).Value Then
            If Right(txtCapa.Text, 1) <> Modulo11UBB(Val(Left(txtCapa.Text, Len(txtCapa.Text) - 1))) Then
                If Right(txtCapa.Text, 1) <> Modulo11Simplificado(Val(Left(txtCapa.Text, Len(txtCapa.Text) - 1))) Then
                    If Right(txtCapa.Text, 1) <> Modulo11U(Val(Left(txtCapa.Text, Len(txtCapa.Text) - 1))) Then
                        MsgBox "O Dígito Verificador do campo Capa de Envelope não confere.", vbExclamation + vbOKOnly, App.Title
                        txtCapa.SetFocus
                        CriticaEnvMal = False
                        Exit Function
                        End If
                End If
            End If
    ElseIf optTipo(2).Value Then        '--- Consiste Envelope Fininvest ---
        If Len(Trim(txtCapa.Text)) <> 10 Then
            MsgBox "O campo Capa de Envelope (Fininvest) deve ter 10 dígitos.", vbExclamation + vbOKOnly, App.Title
            txtCapa.SetFocus
            CriticaEnvMal = False
            Exit Function
        End If
            
        If Left(txtCapa.Text, 2) <> "01" Or CLng(Mid(txtCapa.Text, 3, 7)) < 1 Then
            MsgBox "O campo Capa de Envelope (Fininvest) não é válido.", vbExclamation + vbOKOnly, App.Title
            txtCapa.SetFocus
            CriticaEnvMal = False
            Exit Function
        End If
            
        If Right(txtCapa.Text, 1) <> Modulo11Fininvest(Val(Left(txtCapa.Text, Len(txtCapa.Text) - 1))) Then
            MsgBox "O Dígito Verificador do campo Capa de Envelope (Fininvest) não confere.", vbExclamation + vbOKOnly, App.Title
            txtCapa.SetFocus
            CriticaEnvMal = False
            Exit Function
        End If
    
    ElseIf optTipo(1).Value And Len(Trim(txtCapa.Text)) <> 14 Then
        MsgBox "O campo Capa do Malote deve ter 14 dígitos.", vbExclamation + vbOKOnly, App.Title
        txtCapa.SetFocus
        CriticaEnvMal = False
        Exit Function
    ElseIf optTipo(1).Value And Left(Trim(txtCapa.Text), 4) <> "0600" Then
        MsgBox "O campo Capa do Malote não é válido.", vbExclamation + vbOKOnly, App.Title
        txtCapa.SetFocus
        CriticaEnvMal = False
        Exit Function
    ElseIf optTipo(1).Value And Trim(txtNumMalote.Text) = "" Then
        MsgBox "O campo Número do Malote deve ser informado.", vbExclamation + vbOKOnly, App.Title
        txtNumMalote.SetFocus
        CriticaEnvMal = False
        Exit Function
    ElseIf optTipo(1).Value And Not IsNumeric(txtNumMalote.Text) Then
        MsgBox "O campo Número do Malote está inválido.", vbExclamation + vbOKOnly, App.Title
        txtNumMalote.Text = ""
        txtNumMalote.SetFocus
        CriticaEnvMal = False
        Exit Function
    ElseIf optTipo(1).Value Then

        'Verifica Nr.Malote Novo ou Antigo
        iLenMalote = Len(Trim(txtNumMalote))
        If (Left(CStr(txtNumMalote), 2) = "09" And iLenMalote = 12) Or _
            (Left(CStr(txtNumMalote), 1) = "9" And iLenMalote = 11) Then

            txtNumMalote = Format(txtNumMalote, "000000000000")
            If Val(Mid(txtNumMalote, 3, 9)) < 1 Then
                MsgBox "Número do Malote Empresa Inválido !", vbExclamation + vbOKOnly, App.Title
                txtNumMalote.SetFocus
                CriticaEnvMal = False: Exit Function
            End If

            If Left(CStr(txtNumMalote), 2) <> "09" Then
                MsgBox "Número do Malote Empresa Inválido !", vbExclamation + vbOKOnly, App.Title
                txtNumMalote.SetFocus
                CriticaEnvMal = False: Exit Function
            End If
        Else
            If iLenMalote > 11 Then
                MsgBox "Número do Malote Empresa Inválido !", vbExclamation + vbOKOnly, App.Title
                txtNumMalote.SetFocus
                CriticaEnvMal = False: Exit Function
            End If
            txtNumMalote = Format(txtNumMalote, "00000000000")
        End If

        'Calcula Modulo 10 para Nr Malote antigo (11) ou Novo (12)posições
        If Not Modulo10(txtNumMalote, Len(txtNumMalote.Text)) Then
            MsgBox "O campo Número do Malote não é válido.", vbExclamation + vbOKOnly, App.Title
            txtNumMalote.SetFocus
            CriticaEnvMal = False
            Exit Function
        End If
    End If

    CriticaEnvMal = True

End Function
Private Function VerificaAgencia() As Boolean
'* Verifica se Agencia Atual ja foi gravada na Tabela Agencia *'

    VerificaAgencia = True

    With qryGetAgencia
        .rdoParameters(0) = Geral.DataProcessamento
        .rdoParameters(1) = Agencia
        .rdoParameters(2) = IIf(IdEnv_Mal = "M", "M", "E")  'Pesquisa Recepção de (M)alote ou (E)nvelope ou Envelope Fininvest
        Set rsGetAgencia = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If IsNull(rsGetAgencia!qtdGravada) Then
        MsgBox "Não Foi Recepcionado nenhum Lote com " & IIf(IdEnv_Mal = "M", "Malote", "Envelope") & " nesta agência", vbInformation + vbOKOnly
        
        LimpaCampos
        cmbAgencia.SetFocus
        VerificaAgencia = False
    Else
       ContGravada = rsGetAgencia!qtdGravada
    End If

End Function
Private Sub AtualizaInformacaoAgencia()
    With qryAtualizaAgencia
        .rdoParameters(0) = Geral.DataProcessamento
        .rdoParameters(1) = Agencia
        .rdoParameters(2) = IIf(IdEnv_Mal = "M", "M", "E")  'Atualiza Recepção de (M)alote ou (E)nvelope ou Envelope Fininvest
        Set rsAtualizaAgencia = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
End Sub
Private Sub txtNumMalote_GotFocus()
    With txtNumMalote
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtNumMalote_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call SalvaReg
    End If
End Sub
