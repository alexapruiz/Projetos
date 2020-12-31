VERSION 5.00
Begin VB.Form frmRecepcao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepção de Envelopes e Malotes"
   ClientHeight    =   2652
   ClientLeft      =   1392
   ClientTop       =   1380
   ClientWidth     =   6516
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2652
   ScaleWidth      =   6516
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFechar 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   372
      Left            =   3345
      TabIndex        =   3
      Top             =   2136
      Width           =   1512
   End
   Begin VB.CommandButton cmdRecepcionar 
      Caption         =   "&Recepcionar"
      Height          =   372
      Left            =   1659
      TabIndex        =   2
      Top             =   2136
      Width           =   1512
   End
   Begin VB.Frame Frame5 
      Caption         =   "Quantidade Recepcionada"
      Height          =   1176
      Left            =   4146
      TabIndex        =   6
      Top             =   804
      Width           =   2244
      Begin VB.Label lblQtdeRecepcionada 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   828
         Left            =   192
         TabIndex        =   9
         Top             =   240
         Width           =   1884
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1176
      Left            =   126
      TabIndex        =   5
      Top             =   804
      Width           =   3912
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
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   1680
         MaxLength       =   12
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   660
         Width           =   2076
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
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   1680
         MaxLength       =   14
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   240
         Width           =   2076
      End
      Begin VB.Label lblNumMalote 
         BackStyle       =   0  'Transparent
         Caption         =   "Número do Malote:"
         Height          =   192
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1512
      End
      Begin VB.Label lblCapa 
         BackStyle       =   0  'Transparent
         Caption         =   "Capa do Envelope:"
         Height          =   192
         Left            =   120
         TabIndex        =   7
         Top             =   300
         Width           =   1512
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Agência de Origem"
      Height          =   672
      Left            =   126
      TabIndex        =   4
      Top             =   84
      Width           =   6264
      Begin VB.Label lblAgencia 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
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
         Height          =   348
         Left            =   192
         TabIndex        =   11
         Top             =   222
         Width           =   828
      End
      Begin VB.Label lblNomeAgencia 
         BackStyle       =   0  'Transparent
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
         Left            =   1200
         TabIndex        =   10
         Top             =   240
         Width           =   4692
      End
   End
End
Attribute VB_Name = "frmRecepcao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private rsAllAgenf          As rdoResultset
Private qryGetAllAgenf      As rdoQuery
Private qryInsereCapa       As rdoQuery
Private qryInsereAgencia    As rdoQuery
Private qryIKInsereCapa     As rdoQuery
Private qryIKOCorrencia     As rdoQuery
Private CountEnvelopes      As Integer
Private agencia             As String
Private IdEnv_Mal           As String * 1
Private ChangeAgencia       As Boolean
Private ChangeTipo          As Boolean


Private Type tpCaixa
    CodigoTransacao         As String
    NumSeqServidor          As String
    Versao                  As String
    AgenciaCentralizadora   As String
    AgenciaSatelite         As String
    TipoAutorizacao         As String
    TipoTransacao           As String
    TipoMovimento           As String
    NumeroPeriferico        As String
    TipoTerminal            As String
    NSU                     As String
    NSUS                    As String
    Indicacao_Backoffice    As String
    Hora                    As String * 4
    Indicacao_Transacional  As String
    TipoAtendimento         As String
    TipoRepasse             As String
    CodigoEvento            As String
    NumeroCaixaExpresso     As String
    MatriculaSupervisor     As String
    Terminal                As String
End Type



Private IK                  As tpCaixa

Private Sub InsereInformacaoAgencia()
    On Error GoTo ErroAgencia
    rdoErrors.Clear
    With qryInsereAgencia
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = Val(agencia)
        .rdoParameters(3) = Val(txtLacre.Text)
        .rdoParameters(4) = Val(txtQtdeProtocolo.Text)
        .rdoParameters(5) = CountEnvelopes
        .rdoParameters(6) = txtIdentificador.Text
        .rdoParameters(7) = txtHora.Text
        .rdoParameters(8) = Format(Now, "hh:mm")
        '.rdoParameters(9) = IIf(optTipo(0).Value, "E", "M")
        .Execute
        If .rdoParameters(0) <> 0 Then
            MsgBox "Erro na atualização das informações da Agência de Origem.", vbCritical + vbOKOnly, App.Title
            Exit Sub
        End If
    End With
'    If Val(txtQtdeProtocolo.Text) <> CountEnvelopes Then
'        ImprimeOcorQtd
'    End If
    CountEnvelopes = 0

    Exit Sub
ErroAgencia:
    Select Case TratamentoErro(Geral.Banco, "Erro na atualização das informações da Agência de Origem.", Err, rdoErrors)
        Case vbCancel
            Unload Me
        Case vbRetry
            Resume
    End Select

End Sub

Private Function CriticaCampos() As Boolean

    Dim RetAgencia As Integer, iLenMalote As Integer

    CriticaCampos = False

    'Validar Agencia
    'RetAgencia = ValidaAgencia(Val(cmbAgencia.Text), "", False)

    'Verificar Retorno da Função
'    Select Case RetAgencia
'      Case 2
'        'Agencia em Feriado
'        MsgBox "A agência de origem está em feriado.", vbInformation + vbOKOnly, App.Title
'        cmbAgencia.SetFocus
'        CriticaCampos = False
'        Exit Function
'      Case 3
'        'Agencia Fechada
'        MsgBox "A agência de origem está fechada.", vbInformation + vbOKOnly, App.Title
'        cmbAgencia.SetFocus
'        CriticaCampos = False
'        Exit Function
'      Case 4
'        'Agencia não Cadastrada
'        MsgBox "A agência de origem não está cadastrada.", vbInformation + vbOKOnly, App.Title
'        cmbAgencia.SetFocus
'        CriticaCampos = False
'        Exit Function
'    End Select
    
    If Trim(txtCapa.Text) = "" Then
        MsgBox "O campo Capa deve ser informado.", vbExclamation + vbOKOnly, App.Title
        txtCapa.SetFocus
        Exit Function
    ElseIf Not IsNumeric(txtCapa.Text) Then
        MsgBox "O campo Capa está inválido.", vbExclamation + vbOKOnly, App.Title
        txtCapa.SetFocus
        Exit Function
    End If
    
    
    ''''''''''''''''''''''
    'Validação de Envelope
    ''''''''''''''''''''''
    If Len(txtCapa.Text) <= 8 Then
        If Right(txtCapa.Text, 1) <> Modulo11UBB(Val(Left(txtCapa.Text, Len(txtCapa.Text) - 1))) Then
            If Right(txtCapa.Text, 1) <> Modulo11Simplificado(Val(Left(txtCapa.Text, Len(txtCapa.Text) - 1))) Then
                If Right(txtCapa.Text, 1) <> Modulo11U(Val(Left(txtCapa.Text, Len(txtCapa.Text) - 1))) Then
                    MsgBox "O Dígito Verificador do campo Capa de Envelope não confere.", vbExclamation + vbOKOnly, App.Title
                    txtCapa.SetFocus
                    Exit Function
                End If
            End If
        End If
    Else
        If Len(txtCapa.Text) = 14 And txtNumMalote.Enabled = False Then
            txtNumMalote.Enabled = True
            Exit Function
        End If
        If Len(Trim(txtCapa.Text)) <> 14 Then
            MsgBox "O campo Capa do Malote deve ter 14 dígitos.", vbExclamation + vbOKOnly, App.Title
            txtCapa.SetFocus
            Exit Function
        ElseIf Left(Trim(txtCapa.Text), 4) <> "0600" Then
            MsgBox "O campo Capa do Malote deve não é válido.", vbExclamation + vbOKOnly, App.Title
            txtCapa.SetFocus
            Exit Function
        ElseIf Trim(txtNumMalote.Text) = "" Then
            MsgBox "O campo Número do Malote deve ser informado.", vbExclamation + vbOKOnly, App.Title
            txtNumMalote.SetFocus
            Exit Function
        ElseIf Not IsNumeric(txtNumMalote.Text) Then
            MsgBox "O campo Número do Malote está inválido.", vbExclamation + vbOKOnly, App.Title
            txtNumMalote.SetFocus
            Exit Function
        End If
        
        'Calculo
        
        iLenMalote = Len(Trim(txtNumMalote))
        If (Left(CStr(txtNumMalote), 2) = "09" And iLenMalote = 12) Or _
            (Left(CStr(txtNumMalote), 1) = "9" And iLenMalote = 11) Then

            txtNumMalote = Format(txtNumMalote, "000000000000")
            If Val(Mid(txtNumMalote, 3, 9)) < 1 Then
                MsgBox "Número do Malote Empresa Inválido !", vbExclamation + vbOKOnly, App.Title
                txtNumMalote.SetFocus
                Exit Function
            End If

            If Left(CStr(txtNumMalote), 2) <> "09" Then
                MsgBox "Número do Malote Empresa Inválido !", vbExclamation + vbOKOnly, App.Title
                txtNumMalote.SetFocus
                Exit Function
            End If
        Else
            If iLenMalote > 11 Then
                MsgBox "Número do Malote Empresa Inválido !", vbExclamation + vbOKOnly, App.Title
                txtNumMalote.SetFocus
                Exit Function
            End If
            txtNumMalote = Format(txtNumMalote, "00000000000")
        End If

        'Calcula Modulo 10 para Nr Malote antigo (11) ou Novo (12)posições
        If Not Modulo10(txtNumMalote, Len(txtNumMalote.Text)) Then
            MsgBox "O campo Número do Malote não é válido.", vbExclamation + vbOKOnly, App.Title
            txtNumMalote.SetFocus
            Exit Function
        End If
    End If
    
    CriticaCampos = True
    
End Function


Private Sub Limpar()
    CountEnvelopes = 0
    cmbAgencia.Text = ""
    txtLacre.Text = ""
    txtQtdeProtocolo.Text = ""
    txtIdentificador.Text = ""
    txtHora.Text = ""
    lblQtdeRecepcionada.Caption = Format(CountEnvelopes, "00")
    txtCapa.Text = ""
    txtNumMalote.Text = ""
End Sub


Private Sub cmdFechar_Click()
    Unload Me
End Sub

Private Sub cmdRecepcionar_Click()


    Dim IdCapa              As Integer
    Dim IdEnvMal            As String * 1
    Dim tb1                 As rdo.rdoResultset
    
    cmdRecepcionar.SetFocus
    
    If Not CriticaCampos Then
        Exit Sub
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''
    'Identifica se é Malote Empresa ou Envelope'
    ''''''''''''''''''''''''''''''''''''''''''''
    IdEnvMal = IIf(CBool(Len(Trim(txtCapa.Text)) > 8 And Trim(Left(txtCapa, 4)) = "0600"), "M", "E")

    On Error GoTo ErroCapa
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                   Entrando em comunicação com o IK                                        '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    '''''''''''''''''''''''''
    'Ler tabela de parametro'
    '''''''''''''''''''''''''
    With Geral.qryLeituraParametro
        Set tb1 = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    If tb1.EOF Then
        MsgBox "Não foi possível ler informações da tabela Parâmetro.", vbCritical + vbOKOnly, App.Title
        Geral.qryLeituraParametro.Close
        Unload Me
    End If
    
    

    ''''''''''''''''''''''''''''
    'Envia informação para o IK'
    ''''''''''''''''''''''''''''
'    With qryIKInsereCapa
'        .rdoParameters(0).Direction = rdParamReturnValue        '
'        .rdoParameters(1) = "IKRE"                              'Codigo da transação
'        .rdoParameters(2) = 0                                   'Numero seq. unico do servidor
'        .rdoParameters(3) = 509 'tb1!Versao                          'Versão do registro
'        .rdoParameters(4) = Geral.AgenciaApresentante           'Codigo da agencia centralizadora
'        .rdoParameters(5) = Geral.AgenciaApresentante           'Codigo da agencia satelite
'        .rdoParameters(6) = 0                                   'Tipo de autorizacao
'        .rdoParameters(7) = 0                                   'Tipo de transacao
'        .rdoParameters(8) = 1                                   'Tipo de Movimento
'        .rdoParameters(9) = 799                                 'Numero do periferico
'        .rdoParameters(10) = 1                                  'Tipo de terminal
'        .rdoParameters(11) = 0 ' NSU ???                        'Numero seq. unico do terminal
'        .rdoParameters(12) = 0                                  'Numero seq. unico substituto
'        .rdoParameters(13) = 0                                  'Indicacao de envio ao backoffice
'        .rdoParameters(14) = Format(Now, "HHMM")                'Hora de efetivação da transação
'        .rdoParameters(15) = " "                                'Indicação transacional
'        .rdoParameters(16) = IIf(IdEnvMal = "E", 6, 7)          'Tipo de atendimento
'        .rdoParameters(17) = 0                                  'Tipo de repasse
'        .rdoParameters(18) = 0                                  'Codigo de evento
'        .rdoParameters(19) = IIf(IdEnvMal = "E", "6", "7")      'Numero cx expresso (Env/Malote)
'        .rdoParameters(20) = Format(Now, "DDMMYYYY")            'Data da remessa
'        .rdoParameters(21) = 2                                  'Numero da remessa
'        .rdoParameters(22) = 2                                  'Tipo de remessa
'        .rdoParameters(23) = 1                                  'Quantidade de documentos
'        .rdoParameters(24) = 799                                'Numero terminal TAC
'        .Execute
'        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        'Tratamento de erro da procedure de inserção da capa ou malote'
'        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        If .rdoParameters(0) <> 0 Then
'            MsgBox "Não foi possível inserir esta capa no sistema IK.", vbCritical
'            Exit Sub
'        End If
'    End With
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                     Fim da comunicação com o IK                                        '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    rdoErrors.Clear
    With qryInsereCapa
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = 0                    ' IdLote
        .rdoParameters(3) = IdEnvMal
        .rdoParameters(4) = Val(txtCapa.Text)
        .rdoParameters(5) = IIf(IdEnvMal = "E", 0, Val(txtNumMalote.Text))
        .rdoParameters(6) = Val(lblAgencia)
        .rdoParameters(7) = "0"                  ' Status
        .rdoParameters(8).Direction = rdParamOutput
        .Execute
        If .rdoParameters(0) = 1 Then
            MsgBox "Atenção! Não é possível Recepcionar este " & _
                IIf(IdEnvMal = "E", "Envelope", "Malote") & _
                ", pois já está recepcionado no sistema.", _
                vbExclamation + vbOKOnly, App.Title
            txtCapa.SetFocus
            Exit Sub
        ElseIf .rdoParameters(0) > 1 Then
            MsgBox "Erro na recepção de Envelope/Malote", vbCritical + vbOKOnly, App.Title
            Exit Sub
        End If
    End With

    'Gravar Log
    Call GravaLog(Geral.Banco, Geral.DataProcessamento, qryInsereCapa.rdoParameters(8).Value, 0, Geral.Usuario.Login, 20)

    CountEnvelopes = CLng(CountEnvelopes) + 1
    agencia = lblAgencia.Caption
    lblQtdeRecepcionada.Caption = Format(CountEnvelopes, "00")
    
    txtCapa.Text = ""
    txtNumMalote.Text = ""
    
    txtCapa.SetFocus
    Exit Sub
ErroLog:
    Select Case TratamentoErro(Geral.Banco, "Erro na atualização do Log de operação.", Err, rdoErrors)
        Case vbCancel
            Unload Me
        Case vbRetry
            Resume
    End Select
    Exit Sub
ErroCapa:
    Select Case TratamentoErro(Geral.Banco, "Erro na recepção de Envelope/Malote.", Err, rdoErrors)
        Case vbCancel
            Unload Me
        Case vbRetry
            Resume
    End Select
End Sub

Private Sub Form_Load()

    Set qryGetAllAgenf = Geral.Banco.CreateQuery("", "{ call MDIAG_GetAllAgenf }")
    Set qryInsereCapa = Geral.Banco.CreateQuery("", "{ ? = call MDIAG_InsereCapa (?,?,?,?,?,?,?,?) }")
    Set qryInsereAgencia = Geral.Banco.CreateQuery("", "{? = call InsereAgencia(?,?,?,?,?,?,?,?,?)}")
    Set Geral.qryLeituraParametro = Geral.Banco.CreateQuery("", "{call MDIAG_LerParametro}")
    
    '''''''''''''''''''''''''
    'queries de acesso no IK'
    '''''''''''''''''''''''''
    Set qryIKInsereCapa = Geral.BancoCaixa.CreateQuery("", "{? = Call taenvicx(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)}")
    Set qryIKOCorrencia = Geral.BancoCaixa.CreateQuery("", "{? = Call tareocor(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)}")
    
    
    
    Set rsAllAgenf = qryGetAllAgenf.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
'    If Not rsAllAgenf.EOF Then
'    End If
    
    
    ''''''''''''''''''''''''''''''
    'Buscar a descrição da Agencia
    ''''''''''''''''''''''''''''''
    lblAgencia = Format(Geral.AgenciaApresentante, "0000")
    lblNomeAgencia = "Agencia "

    
    agencia = ""
    'IdEnv_Mal = IIf(optTipo(0).Value, "E", "M")
    CountEnvelopes = 0
    ChangeAgencia = True
    ChangeTipo = True
    txtCapa.Text = ""
    txtNumMalote.Text = ""
    lblQtdeRecepcionada.Caption = "00"
    'optTipo(0).Value = True
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim StrMsg As String
    
    If CountEnvelopes > 0 Then
'        If Val(txtQtdeProtocolo.Text) <> CountEnvelopes Then
'            StrMsg = "Atenção! A Quantidade do Protocolo não confere com Quantidade Recepcionada." & vbCr & _
'                     "Confirma encerramento da recepção da agência " & agencia & "?"
'        Else
'            StrMsg = "Confirma encerramento da recepção da agência " & agencia & "?"
'        End If
'        If MsgBox(StrMsg, vbQuestion + vbYesNo, App.Title) = vbYes Then
'            ' atualiza dados da agencia
'            InsereInformacaoAgencia
'            Cancel = 0
'        Else
'            Cancel = 1
'            txtCapa.SetFocus
'            Exit Sub
'        End If
    End If
    rsAllAgenf.Close
    qryGetAllAgenf.Close
    qryInsereCapa.Close
    qryInsereAgencia.Close
    
    qryIKInsereCapa.Close
    qryIKOCorrencia.Close
End Sub

Private Sub txtCapa_Change()

    ''''''''''''''''''''''''''''''''''''''''''''''
    'Se for malote empresa então habilita o texto'
    ''''''''''''''''''''''''''''''''''''''''''''''
    txtCapa.Text = Trim(txtCapa.Text)
    txtNumMalote.Enabled = CBool(Left(txtCapa.Text, 4) = "0600" And Len(txtCapa.Text) = 14)
    
    lblNumMalote.Enabled = txtNumMalote.Enabled
    
    
End Sub

Private Sub txtCapa_GotFocus()

    SelecionarTexto txtCapa
    
End Sub

Private Sub txtCapa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
    
        If Not (Left(txtCapa.Text, 4) = "0600" And Len(txtCapa.Text) = 14) Then
            cmdRecepcionar_Click
        Else
            SendKeys ("{TAB}")
        End If
        
    End If
End Sub

Private Sub txtCapa_KeyPress(KeyAscii As Integer)

    SoNumero KeyAscii

End Sub

Private Sub txtCapa_LostFocus()
'    With txtCapa
'        .Text = Trim(.Text)
'        If .Text <> "" Then
'            If Not IsNumeric(.Text) Then
'                MsgBox "O campo Capa deve ser um valor numérico.", vbExclamation + vbOKOnly, App.Title
'                .SetFocus
'            ElseIf Len(.Text) < 2 Then
'                MsgBox "O campo Capa deve ter no mínimo 2 dígitos.", vbExclamation + vbOKOnly, App.Title
'                .SetFocus
'            End If
'        End If
'    End With
End Sub

Private Sub txtHora_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys ("{TAB}")
    End If
End Sub
Private Sub txtNumMalote_GotFocus()

    SelecionarTexto txtNumMalote
    
End Sub

Private Sub txtNumMalote_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdRecepcionar_Click
    End If
End Sub

Private Sub txtNumMalote_KeyPress(KeyAscii As Integer)

    SoNumero KeyAscii

End Sub

Private Sub txtNumMalote_LostFocus()
    With txtNumMalote
        .Text = Trim(.Text)
        If .Text <> "" Then
            If Not IsNumeric(.Text) Then
                MsgBox "O campo Número do Malote deve ser um valor numérico.", vbExclamation + vbOKOnly, App.Title
                .SetFocus
            ElseIf Len(.Text) < 2 Then
                MsgBox "O campo Número do Malote deve ter no mínimo 2 dígitos.", vbExclamation + vbOKOnly, App.Title
                .SetFocus
            End If
        End If
    End With
End Sub
