VERSION 5.00
Begin VB.Form Recepcao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepção de Envelopes e Malotes"
   ClientHeight    =   4344
   ClientLeft      =   1392
   ClientTop       =   1356
   ClientWidth     =   6216
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4344
   ScaleWidth      =   6216
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdRecepcionar 
      Caption         =   "Recepcionar"
      Height          =   372
      Left            =   720
      TabIndex        =   6
      Top             =   3792
      Width           =   1512
   End
   Begin VB.CommandButton cmdEncerrar 
      Caption         =   "Encerrar &Agência"
      Height          =   372
      Left            =   2364
      TabIndex        =   7
      Top             =   3792
      Width           =   1512
   End
   Begin VB.CommandButton cmdFechar 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   372
      Left            =   4020
      TabIndex        =   8
      Top             =   3792
      Width           =   1512
   End
   Begin VB.Frame Frame3 
      Caption         =   "Dados"
      Height          =   2544
      Left            =   60
      TabIndex        =   11
      Top             =   1020
      Width           =   6072
      Begin VB.TextBox txtRemessa 
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
         Left            =   3180
         MaxLength       =   1
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   745
         Width           =   1452
      End
      Begin VB.TextBox TxtqtdeMal 
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
         Left            =   3180
         MaxLength       =   5
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1872
         Width           =   1452
      End
      Begin VB.TextBox txtqtdeEnv 
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
         Left            =   3180
         MaxLength       =   10
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1488
         Width           =   1452
      End
      Begin VB.TextBox txtLacre 
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
         Left            =   3180
         MaxLength       =   8
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   372
         Width           =   1452
      End
      Begin VB.TextBox txtIdentificador 
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
         Left            =   3180
         MaxLength       =   10
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1116
         Width           =   1452
      End
      Begin VB.Label Label2 
         Caption         =   "Remessa :"
         Height          =   192
         Left            =   960
         TabIndex        =   16
         Top             =   864
         Width           =   2352
      End
      Begin VB.Label Label6 
         Caption         =   "Quantidade Envelope :"
         Height          =   192
         Left            =   960
         TabIndex        =   15
         Top             =   1572
         Width           =   2232
      End
      Begin VB.Label Label5 
         Caption         =   "Quantidade Malote       :"
         Height          =   192
         Left            =   960
         TabIndex        =   14
         Top             =   1932
         Width           =   1872
      End
      Begin VB.Label Label3 
         Caption         =   "Identificador :"
         Height          =   192
         Left            =   960
         TabIndex        =   13
         Top             =   1200
         Width           =   1272
      End
      Begin VB.Label Label1 
         Caption         =   "Nº do Lacre :"
         Height          =   192
         Left            =   960
         TabIndex        =   12
         Top             =   480
         Width           =   972
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Agência de Origem"
      Height          =   672
      Left            =   60
      TabIndex        =   9
      Top             =   180
      Width           =   6072
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
         Left            =   120
         TabIndex        =   0
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
         Left            =   1200
         TabIndex        =   10
         Top             =   240
         Width           =   4692
      End
   End
End
Attribute VB_Name = "Recepcao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsAllAgenf          As rdoResultset
Private rsGetAgencia        As rdoResultset

Private qryGetAllAgenf      As rdoQuery
Private qryGetAgencia       As rdoQuery
Private qryInsereAgencia    As rdoQuery

Private Agencia             As String
Private IdEnv_Mal           As String * 1

Private ChangeAgencia       As Boolean
Private ChangeTipo          As Boolean


Private Sub InsereInformacaoAgenciaEnv()
'* Grava Dados Tabela Agencia *'

On Error GoTo ErroAgencia
    
     rdoErrors.Clear
     With qryInsereAgencia
          .rdoParameters(1) = Geral.DataProcessamento
          .rdoParameters(2) = Val(cmbAgencia)
          .rdoParameters(3) = Val(txtLacre.Text)
          .rdoParameters(4) = Val(txtqtdeEnv.Text)
          .rdoParameters(5) = txtIdentificador.Text
          .rdoParameters(6) = "E"
          .rdoParameters(7) = Val(txtRemessa.Text)
          .rdoParameters(8) = "R"
          .Execute
        
          If .rdoParameters(0) <> 0 Then
               MsgBox "Erro na atualização das informações da Agência de Origem.", vbCritical + vbOKOnly, App.Title
               Exit Sub
          End If
     End With

     Exit Sub

ErroAgencia:

     Select Case TratamentoErro("Erro na atualização das informações da Agência de Origem.", Err, rdoErrors)
        Case vbCancel
            Unload Me
        Case vbRetry
            Resume
    End Select

End Sub
Private Function CriticaCampos() As Boolean
   '* Faz a Consistência dos Campos *'
   
    Dim RetAgencia As Integer, iLenMalote As Integer

    CriticaCampos = True
   
    If Trim(txtLacre.Text) = "" Then
        MsgBox "O campo Nº do Lacre deve ser informado.", vbExclamation + vbOKOnly, App.Title
        txtLacre.SetFocus
        CriticaCampos = False
    ElseIf Not IsNumeric(txtLacre.Text) Then
        MsgBox "O Nº do Lacre está inválido.", vbExclamation + vbOKOnly, App.Title
        txtLacre.SetFocus
        CriticaCampos = False
    ElseIf Trim(txtqtdeEnv.Text) = "" And Trim(TxtqtdeMal.Text) = "" Then
        MsgBox "O campo Quantidade é obrigatório.", vbExclamation + vbOKOnly, App.Title
        txtqtdeEnv.SetFocus
        CriticaCampos = False
    End If
    
End Function
Private Sub Limpar()
 '* Limpa Objetos *'
 
    cmbAgencia.Text = ""
    txtLacre.Text = ""
    txtRemessa.Text = ""
    txtIdentificador.Text = ""
    txtqtdeEnv = ""
    TxtqtdeMal = ""
End Sub
Private Sub cmbAgencia_Change()
    cmbAgencia_Click
End Sub
Private Sub cmbAgencia_Click()
    If Not ChangeAgencia Then
        Exit Sub
    End If
    
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
    
    txtLacre.Enabled = True
    txtIdentificador.Enabled = True

End Sub
Private Sub cmbAgencia_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys ("{TAB}")
    End If
End Sub
Private Sub cmbAgencia_LostFocus()
    
    If LCase(ActiveControl.Name) = "cmdfechar" Then Exit Sub
    
    If Trim(cmbAgencia.Text) = "" Then
        MsgBox "O campo Agência deve ser informado.", vbExclamation + vbOKOnly, App.Title
        cmbAgencia.SetFocus
        Exit Sub
    Else
    
     'Valida Agencia
      RetAgencia = ValidaAgencia(Val(cmbAgencia.Text), "", False)

     'Verifica Retorno da Função
      Select Case RetAgencia
        Case 0
          Agencia = cmbAgencia.Text
          Exit Sub
        Case 2
         'Agencia em Feriado
          MsgBox "A agência de origem está em feriado.", vbInformation + vbOKOnly, App.Title
          cmbAgencia.SetFocus
        Case 3
         'Agencia Fechada
          MsgBox "A agência de origem está fechada.", vbInformation + vbOKOnly, App.Title
          cmbAgencia.SetFocus
        Case 4
         'Agencia não Cadastrada
          MsgBox "A agência de origem não está cadastrada.", vbInformation + vbOKOnly, App.Title
          cmbAgencia.SetFocus
      End Select
    End If
        
End Sub
Private Sub cmdEncerrar_Click()
     
    If Len(Trim(cmbAgencia)) = 0 Then
        cmbAgencia.SetFocus
        lblAgencia.Caption = "Nome da Agência de Origem"
        Exit Sub
    End If
    
    StrMsg = "Confirma encerramento da recepção da agência: " & Agencia & " ?"
    
    If MsgBox(StrMsg, vbQuestion + vbYesNo, App.Title) = vbYes Then
        Call Limpar
        lblAgencia.Caption = "Nome da Agência de Origem"
        Agencia = ""
        lblAgencia.Caption = ""
        cmbAgencia.SetFocus
    End If

End Sub
Private Sub CmdFechar_Click()

    Unload Me
    
End Sub
Private Sub SalvaRecepcaoEnv()

    Call InsereInformacaoAgenciaEnv
     
End Sub
Private Sub CmdRecepcionar_Click()

    If Not CriticaCampos Then Exit Sub

    If Len(Trim(txtqtdeEnv)) <> 0 Then
        SalvaRecepcaoEnv
    End If
            
    If Len(Trim(TxtqtdeMal)) <> 0 Then
        SalvaRecepcaoMal
    End If
    
    txtLacre.Text = ""
    txtRemessa.Text = ""
    txtIdentificador.Text = ""
    txtqtdeEnv.Text = ""
    TxtqtdeMal.Text = ""
    
    txtLacre.SetFocus
    
End Sub
Private Sub Form_Activate()

   'Inclusão de chamada a rotina AtualizaAtividade
   Call AtualizaAtividade(2)
   
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then
        cmdEncerrar_Click
    End If
End Sub
Private Sub Form_Load()
    
    Set qryGetAllAgenf = Geral.Banco.CreateQuery("", "{ call GetAllAgenf }")
    Set qryGetAgencia = Geral.Banco.CreateQuery("", "{call GetAgencia (?,?,?) }")
    Set qryInsereAgencia = Geral.Banco.CreateQuery("", "{? = call InsereAgencia(?,?,?,?,?,?,?,?)}")
    
    Set rsAllAgenf = qryGetAllAgenf.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    
    While Not rsAllAgenf.EOF
        cmbAgencia.AddItem (Format(rsAllAgenf!agefscdagen, "0000"))
        rsAllAgenf.MoveNext
    Wend
    
    Limpar
    Agencia = ""
    ChangeAgencia = True
    ChangeTipo = True
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    rsAllAgenf.Close
    qryGetAllAgenf.Close
    qryInsereAgencia.Close
End Sub

Private Sub txtIdentificador_GotFocus()
    With txtIdentificador
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtIdentificador_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys ("{TAB}")
    End If
End Sub
Private Sub txtLacre_GotFocus()
    With txtLacre
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtLacre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys ("{TAB}")
    End If
End Sub
Private Sub txtLacre_LostFocus()
    With txtLacre
        .Text = Trim(.Text)
        If .Text <> "" Then
            If Not IsNumeric(.Text) Then
                MsgBox "O campo Nº do Lacre deve ser um valor numérico.", vbExclamation + vbOKOnly, App.Title
                .SelStart = 0
                .SelLength = Len(.Text)
                .SetFocus
            End If
        End If
    End With
End Sub
Private Sub txtqtdeEnv_Change()
    If Len(Trim(txtqtdeEnv.Text)) = 0 Then Exit Sub
    If IsNumeric(txtqtdeEnv.Text) = False Then
        MsgBox "Valor inválido, digite novamente.", vbInformation + vbOKOnly
        txtqtdeEnv.SetFocus
    End If
End Sub
Private Sub txtqtdeEnv_GotFocus()
    With txtqtdeEnv
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtqtdeEnv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys ("{TAB}")
    End If
End Sub
Private Sub TxtqtdeMal_Change()
    If Len(Trim(TxtqtdeMal.Text)) = 0 Then Exit Sub
    If IsNumeric(TxtqtdeMal.Text) = False Then
        MsgBox "Valor inválido, digite novamente.", vbInformation + vbOKOnly
        TxtqtdeMal.SetFocus
    End If
End Sub
Private Sub TxtqtdeMal_GotFocus()
    With TxtqtdeMal
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub TxtqtdeMal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call CmdRecepcionar_Click
        Call cmdEncerrar_Click
    End If
End Sub
Private Sub InsereInformacaoAgenciaMal()
'* Grava Dados Tabela Agencia *'
    
On Error GoTo ErroAgencia
    
     rdoErrors.Clear
     With qryInsereAgencia
          .rdoParameters(1) = Geral.DataProcessamento
          .rdoParameters(2) = Val(cmbAgencia)
          .rdoParameters(3) = Val(txtLacre.Text)
          .rdoParameters(4) = Val(TxtqtdeMal.Text)
          .rdoParameters(5) = txtIdentificador.Text
          .rdoParameters(6) = "M"
          .rdoParameters(7) = Val(txtRemessa.Text)
          .rdoParameters(8) = "R"
          .Execute
        
          If .rdoParameters(0) <> 0 Then
               MsgBox "Erro na atualização das informações da Agência de Origem.", vbCritical + vbOKOnly, App.Title
               Exit Sub
          End If
     End With

     Exit Sub
    
ErroAgencia:
     
        Select Case TratamentoErro("Erro na atualização das informações da Agência de Origem.", Err, rdoErrors)
            Case vbCancel
                Unload Me
            Case vbRetry
                Resume
        End Select

End Sub
Private Sub SalvaRecepcaoMal()
    
     Call InsereInformacaoAgenciaMal
       
End Sub
Private Sub txtRemessa_GotFocus()
    With txtRemessa
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtRemessa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys ("{TAB}")
    End If
End Sub
Private Sub txtRemessa_LostFocus()
    If Len(Trim(txtRemessa.Text)) = 0 Then Exit Sub
    If IsNumeric(txtRemessa.Text) = False Then
        MsgBox "Valor inválido para este campo.", vbInformation + vbOKOnly
        txtRemessa.Text = ""
        txtRemessa.SetFocus
    End If
End Sub
