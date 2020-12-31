VERSION 5.00
Object = "{4A8E27C4-EACA-11D3-9FFC-00104BC8688C}#1.0#0"; "DateEdit.ocx"
Begin VB.Form Login 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2124
   ClientLeft      =   3348
   ClientTop       =   3252
   ClientWidth     =   4488
   FillColor       =   &H8000000F&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   7.8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Login.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   2124
   ScaleWidth      =   4488
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   372
      Left            =   2496
      TabIndex        =   4
      Top             =   1680
      Width           =   1332
   End
   Begin VB.CommandButton cmdConfirma 
      Caption         =   "&Ok"
      Height          =   372
      Left            =   672
      TabIndex        =   3
      Top             =   1680
      Width           =   1332
   End
   Begin VB.PictureBox Panel3D1 
      AutoSize        =   -1  'True
      Height          =   1452
      Left            =   120
      ScaleHeight     =   1404
      ScaleWidth      =   4164
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   4212
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   384
         Left            =   300
         Picture         =   "Login.frx":030A
         ScaleHeight     =   384
         ScaleWidth      =   384
         TabIndex        =   8
         Top             =   240
         Width           =   384
      End
      Begin VB.TextBox txtUsuario 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   384
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   0
         Top             =   60
         Width           =   1584
      End
      Begin VB.TextBox TxtSenha 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   384
         IMEMode         =   3  'DISABLE
         Left            =   2280
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   492
         Width           =   1584
      End
      Begin DATEEDITLib.DateEdit txtDataProcessamento 
         Height          =   384
         Left            =   2280
         TabIndex        =   2
         Top             =   900
         Width           =   1584
         _Version        =   65537
         _ExtentX        =   2794
         _ExtentY        =   677
         _StockProps     =   93
         ForeColor       =   8388608
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   10.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
      End
      Begin VB.Label Label3 
         Caption         =   "Data Movimento:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   132
         TabIndex        =   9
         Top             =   924
         Width           =   2028
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Senha:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1284
         TabIndex        =   7
         Top             =   504
         Width           =   864
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Usuário:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   336
         Left            =   1116
         TabIndex        =   6
         Top             =   120
         Width           =   1020
      End
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Cancelou As Boolean                              ' Indicação de Cancelamento
Public SenhaOk  As Boolean                              ' Indicação de Senha Digitada
Private rsUsuario        As New ADODB.Recordset        ' Resultset para Usuário
Private Proc_Selecionar  As New custodia.Selecionar    ' Define classe de seleção
Private RsGrupoUsuario       As New ADODB.Recordset
' ********************************************
' * Efetua a verificação do Login do Usuário *
' ********************************************
Private Sub cmdConfirma_Click()

Dim eRetorno        As enumRetornoUsuario
    
     Cancelou = False
     SenhaOk = False

     ' ******************************************
     ' * Testa Digitação Obrigatória do Usuário *
     ' ******************************************
     If Trim(txtUsuario.Text) = "" Then
          Beep
          MsgBox "Digite o nome de usuário !", vbExclamation + vbOKOnly, App.Title
          With txtUsuario
               .SelStart = 0
               .SelLength = Len(Trim(.Text))
               .SetFocus
          End With
          Exit Sub
     End If
     ' ****************************************
     ' * Testa Digitação Obrigatória da Senha *
     ' ****************************************
     If Trim(TxtSenha.Text) = "" Then
          Beep
          MsgBox "Digite a sua Senha !", vbExclamation + vbOKOnly, App.Title
          With TxtSenha
               .SelStart = 0
               .SelLength = Len(Trim(.Text))
               .SetFocus
          End With
          Exit Sub
     End If
     ' ****************************************
     ' * Testa Digitação Obrigatória da Data  *
     ' ****************************************
     If Trim(txtDataProcessamento.Text) = "" Then
          Beep
          MsgBox "Informe a data de movimento !", vbExclamation + vbOKOnly, App.Title
          
          With txtDataProcessamento
               .SetFocus
          End With
          
          Exit Sub
     End If
    
        
     'Verifica se a data do movimento é maior que a data do sistema
     If DataAAAAMMDD(Val(txtDataProcessamento.Text)) > Format(Date, "yyyymmdd") Then
          MsgBox "A data do Movimento não pode ser maior que a data do sistema!", vbExclamation + vbOKOnly, App.Title
          txtDataProcessamento.SetFocus
          Exit Sub
     End If

     If DataOk(Val(txtDataProcessamento.Text)) Then
'          If (Weekday(txtDataProcessamento.MaskText, vbSunday) = 1) Or _
'             (Weekday(txtDataProcessamento.MaskText, vbSunday) = 7) Then
'
'              MsgBox "Não é permitido data de final de semana.", vbInformation
'              txtDataProcessamento.SetFocus
'              Exit Sub
'          End If
          Geral.DataProcessamento = DataAAAAMMDD(Val(txtDataProcessamento.Text))
     Else
          MsgBox "A data informada não é válida!" & vbCr & "Obs.: O ano deve ser maior que 1997 e menor que 2051!", vbExclamation + vbOKOnly, App.Title
          txtDataProcessamento.SetFocus
          Exit Sub
     End If

    On Error GoTo ErroLogin

    ''''''''''''''''''''''''''
    ' Obtem dados do Usuário '
    ''''''''''''''''''''''''''
     Geral.UsuarioLogin = Left(Trim(txtUsuario.Text), 10)
     Set rsUsuario = g_cMainConnection.Execute(Proc_Selecionar.GetUsuario(Geral.UsuarioLogin))
     
     'Obtem o Nome do usuário
     If Not rsUsuario.EOF Then
          Geral.UsuarioNome = Trim(rsUsuario!nome)
     Else
          Geral.UsuarioNome = "Desenvolvimento"
     End If

    '***********************************
    '* Verificação do Login do Usuário *
    '***********************************
    eRetorno = VerificaUsuario(rsUsuario, txtUsuario.Text, TxtSenha.Text)
    
    If eRetorno = eSUPERVISOR Then
        Geral.UsuarioId = 0
        SenhaOk = True
    ElseIf eRetorno = eNAO_EXISTENTE Then
        Beep
        MsgBox "Usuário não Cadastrado !", vbExclamation + vbOKOnly, App.Title
        With txtUsuario
            .SelStart = 0
            .SelLength = Len(Trim(.Text))
            .SetFocus
        End With
    ElseIf eRetorno = eSENHA_INCORRETA Then
        Beep
        MsgBox "Senha não Confere !", vbExclamation + vbOKOnly, App.Title
        With TxtSenha
            .SelStart = 0
            .SelLength = Len(Trim(.Text))
            .SetFocus
        End With
    Else
        SenhaOk = True
            
        '''''''''''''''''''''''''''''''''''''''''''
        '*      Grava Informação de IdUsuario    *'
        '''''''''''''''''''''''''''''''''''''''''''
        Geral.UsuarioId = rsUsuario!IdUsuario
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        '* Verificação de Data de Processamento/Parâmetro *'
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        If GeraNovaDataProc(Geral.UsuarioId) = False Then
            Call cmdSair_Click
        End If
        
        ''''''''''''''''''''''''''''''''''''''''''''''''
        '* Monta habilitação de menu conforme Usuário *'
        ''''''''''''''''''''''''''''''''''''''''''''''''
        Call HabilitaMenu
        
    End If
            
    If SenhaOk Then
        Geral.UsuarioLogin = txtUsuario
        Me.Hide
    End If

    Exit Sub

ErroLogin:
     Beep
     MsgBox "Erro na conexão com Banco de Dados.", vbCritical, App.Title
     With txtUsuario
          .SelStart = 0
          .SelLength = Len(Trim(.Text))
          .SetFocus
     End With

End Sub
' ******************************
' * Cancela o Login no Sistema *
' ******************************
Private Sub cmdSair_Click()
    Cancelou = True
    Me.Hide
End Sub

Private Sub Form_Activate()
     
     'Centraliza o form
     Me.Top = (Screen.Height - Me.Height) / 2
     Me.Left = (Screen.Width - Me.Width) / 2
     
     Cancelou = True
End Sub

' **************************************
' * Carrega Módulo de Login no Sistema *
' **************************************
Private Sub Form_Load()
    txtDataProcessamento.Text = Format(DataDDMMAAAA(Geral.DataProcessamento), "00/00/0000")
    Cancelou = False
    SenhaOk = False
End Sub

Private Sub txtDataProcessamento_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If DataOk(Val(txtDataProcessamento.Text)) Then
            cmdConfirma_Click
        Else
            SendKeys "{TAB}"
        End If
    End If
End Sub

' ************************************
' * Ajustando Seleção do Campo Senha *
' ************************************
Private Sub txtSenha_GotFocus()
    With TxtSenha
        .SelStart = 0
        .SelLength = Len(Trim(.Text))
    End With
End Sub

Private Sub txtSenha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys ("{TAB}")
    End If
End Sub

' **************************************
' * Ajustando Seleção do Campo Usuario *
' **************************************
Private Sub txtUsuario_GotFocus()
    With txtUsuario
        .SelStart = 0
        .SelLength = Len(Trim(.Text))
    End With
End Sub
Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys ("{TAB}")
    End If
End Sub

Private Function VerificaUsuario(ByVal rsUsuario As Object, ByVal Usuario As String, ByVal Senha As String) As enumRetornoUsuario
    
    ' ***********************************
    ' * Verificação de Login do Usuário *
    ' ***********************************
    If UCase(Trim(Usuario)) = "DESENV" And UCase(Trim(Senha)) = Geral.SenhaDesenv Then
        '''''''''''''''''''''''''''''''
        'Retorno de usuario supervisor'
        '''''''''''''''''''''''''''''''
        VerificaUsuario = eSUPERVISOR
    ElseIf rsUsuario.EOF Then
        VerificaUsuario = eNAO_EXISTENTE
    ElseIf UCase(Decript(Trim(rsUsuario!Senha))) <> UCase(Trim(Senha)) Then
        VerificaUsuario = eSENHA_INCORRETA
    Else
        VerificaUsuario = eOK
    End If

End Function

Private Sub HabilitaMenu()

Dim ctlMenu As Control
Dim MenuName As String

On Error GoTo ErroHabilitaMenu

     For Each ctlMenu In Principal.Controls
          MenuName = UCase(ctlMenu.Name)
          
          If Left(MenuName, 3) = "MNU" Then
               If InStr("MNUSAIR*", MenuName) = 0 And ctlMenu.Caption <> "-" Then
                    ctlMenu.Enabled = False
               End If
          End If
     Next

     Set RsGrupoUsuario = g_cMainConnection.Execute(Proc_Selecionar.GetGrupoUsuario(Geral.UsuarioId))

     While Not RsGrupoUsuario.EOF
          With Principal
               Select Case RsGrupoUsuario!IdGrupo
                    Case Geral.GrupoUsuario.Supervisor
                              'Recepção
                              .MnuRecAvisoDiferenca.Enabled = True
                              .MnuRecepcao.Enabled = True
                              .MnuRecConfRemessa.Enabled = True
                              .MnuRecepcao.Enabled = True
                              .MnuRecDataBoa.Enabled = True
                              .MnuRecInstrucoes.Enabled = True
                              .MnuRecRejeitados.Enabled = True
                              .MnuRegraGP.Enabled = True

                              'Complementação
                              .MnuComplementacao.Enabled = True

                              'Prova Zero
                              .MnuProvaZero.Enabled = True

                              'Cheque Data Boa
                              .MnuDataBoa.Enabled = True
                              .MnuDataBoaCheques.Enabled = True
                              .MnuDataBoaFusao.Enabled = True

                              'Supervisão
                              .MnuSupervisao.Enabled = True
                              .MnuSupAcompProd.Enabled = True
                              .MnuSupSupervisor.Enabled = True
                              .MnuSupParametros.Enabled = True
                              .MnuSupCadUsuario.Enabled = True

                              'Geração
                              .MnuGeracao.Enabled = True
                              .MnuGerVC.Enabled = True
                              .MnuGerCEL.Enabled = False
                              .MnuGerCEL_Limite.Enabled = True
                              .MnuGerCEL_Superior.Enabled = True
                              .MnuGerCEL_Unibanco.Enabled = True
                              .MnuGerTer.Enabled = True
                              .MnuArqGerTer.Enabled = True
                              .MnuReGerter.Enabled = True
                              .MnuGerAvisoDiferenca.Enabled = True
                              .MnuRecBaixa.Enabled = True
                              .MnuGerRejeitados.Enabled = True
                              .mnuGerExportacao.Enabled = True
                              .mnuGerExportacaoBordero = True
                              .mnuGerExportacaoAlteracaoData = True
                              .mnuGerExportacaoChqBaixados = True
                              .mnuGerExportacaoDataBoa = True
                              
                              'Consulta
                              .MnuConsultas.Enabled = True
                              .MnuConsBorderoCheques = True
                              .mnuConsultaChequesBaixados = True
                              .MnuConsInstrucoes = True

                              'Relatórios
                              .MnuRelatorios.Enabled = True
                              .mnuBorderosConfirmacao.Enabled = True
                              .mnuBorderoTransmissao.Enabled = True
                              .mnuBorderosDatasChequesRejeitados.Enabled = True
                              .mnuBorderosConfirmados.Enabled = True
                              .mnuChequesBaixados.Enabled = True
                              .mnuChBxDataPro.Enabled = True
                              .mnuChBxGeral.Enabled = True
                              .MnuChequesDataBoa.Enabled = True
                              .MnuChequesPendenteFusao.Enabled = True
                              .mnuRelBordero.Enabled = True
                              .mnuRelAvisoDiferenca.Enabled = True
                              .mnuRelAvisoGerado.Enabled = True
                              .mnuRelAvisoRecebido.Enabled = True
                              

                    Case Geral.GrupoUsuario.Digitadores
                              'Recepção
                              .MnuRecAvisoDiferenca.Enabled = True
                              .MnuRecepcao.Enabled = True
                              .MnuRecConfRemessa.Enabled = True
                              .MnuRecepcao.Enabled = True
                              .MnuRecDataBoa.Enabled = True
                              .MnuRecInstrucoes.Enabled = True
                              .MnuRecRejeitados.Enabled = True
                              .MnuRegraGP.Enabled = True

                              'Complementação
                              .MnuComplementacao.Enabled = True

                              'Prova Zero
                              .MnuProvaZero.Enabled = True

                              'Cheque Data Boa
                              .MnuDataBoa.Enabled = True
                              .MnuDataBoaCheques.Enabled = True
                              .MnuDataBoaFusao.Enabled = True


                              'Geração
                              .MnuGeracao.Enabled = True
                              .MnuGerVC.Enabled = True
                              .MnuGerCEL.Enabled = False
                              .MnuGerCEL_Limite.Enabled = False
                              .MnuGerCEL_Superior.Enabled = False
                              .MnuGerCEL_Unibanco.Enabled = False
                              .MnuGerTer.Enabled = True
                              .MnuGerAvisoDiferenca.Enabled = True
                              .MnuRecBaixa.Enabled = True
                              .MnuGerRejeitados.Enabled = True

                              'Consulta
                              .MnuConsultas.Enabled = True
                              .MnuConsBorderoCheques = True
                              .mnuConsultaChequesBaixados = True
                              .MnuConsInstrucoes = True

                              'Relatórios
                              .MnuRelatorios.Enabled = True
                              .mnuBorderosConfirmacao.Enabled = True
                              .mnuBorderoTransmissao.Enabled = True
                              .mnuBorderosDatasChequesRejeitados.Enabled = True
                              .mnuBorderosConfirmados.Enabled = True
                              .mnuChequesBaixados.Enabled = True
                              .mnuChBxDataPro.Enabled = True
                              .mnuChBxGeral.Enabled = True
                              .MnuChequesDataBoa.Enabled = True
                              .MnuChequesPendenteFusao.Enabled = True
                              .mnuRelBordero.Enabled = True
                              .mnuRelAvisoDiferenca.Enabled = True
                              .mnuRelAvisoGerado.Enabled = True
                              .mnuRelAvisoRecebido.Enabled = True
                              
                              
                              
               End Select
          End With
          RsGrupoUsuario.MoveNext
     Wend
     
     Exit Sub
     
ErroHabilitaMenu:
     Beep
     MsgBox "Erro na conexão com Banco de Dados.", vbCritical, App.Title
     Unload Me
End Sub
