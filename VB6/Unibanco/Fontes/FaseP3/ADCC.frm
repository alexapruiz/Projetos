VERSION 5.00
Object = "{9CBA5D64-E3C8-11D3-9FFC-00104BC8688C}#1.0#0"; "CurrencyEdit.ocx"
Begin VB.Form ADCC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Autorização de Débito"
   ClientHeight    =   2100
   ClientLeft      =   1272
   ClientTop       =   1332
   ClientWidth     =   9384
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   9384
   Begin VB.Frame Frame1 
      Height          =   2040
      Left            =   48
      TabIndex        =   13
      Top             =   -24
      Width           =   9276
      Begin VB.CommandButton cmdConfirmar 
         Caption         =   "&Confirmar"
         Height          =   696
         Left            =   7548
         Picture         =   "ADCC.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   204
         Width           =   804
      End
      Begin VB.TextBox txtConta 
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
         Height          =   372
         Left            =   5268
         MaxLength       =   7
         TabIndex        =   4
         Top             =   1512
         Width           =   960
      End
      Begin VB.TextBox txtAgencia 
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
         Height          =   372
         Left            =   4509
         MaxLength       =   4
         TabIndex        =   3
         Top             =   1512
         Width           =   612
      End
      Begin VB.TextBox txtCMC72 
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
         Height          =   372
         Left            =   1359
         MaxLength       =   10
         TabIndex        =   1
         Top             =   1512
         Width           =   1308
      End
      Begin VB.TextBox txtCMC73 
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
         Height          =   372
         Left            =   2814
         MaxLength       =   12
         TabIndex        =   2
         Top             =   1512
         Width           =   1548
      End
      Begin VB.TextBox txtCMC71 
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
         Height          =   372
         Left            =   144
         MaxLength       =   8
         TabIndex        =   0
         Top             =   1512
         Width           =   1080
      End
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   696
         Left            =   8352
         Picture         =   "ADCC.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   204
         Width           =   816
      End
      Begin VB.CommandButton cmdFrenteVerso 
         Caption         =   "Fren/Ver"
         Height          =   696
         Left            =   6744
         Picture         =   "ADCC.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   204
         Width           =   804
      End
      Begin VB.CommandButton cmdInverteCor 
         Caption         =   "Inverter"
         Height          =   696
         Left            =   5928
         Picture         =   "ADCC.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   204
         Width           =   804
      End
      Begin VB.CommandButton cmdRotacao 
         Caption         =   "Rotação"
         Height          =   696
         Left            =   5112
         Picture         =   "ADCC.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   204
         Width           =   804
      End
      Begin VB.CommandButton cmdZoomMenos 
         Caption         =   "Zoom -"
         Height          =   696
         Left            =   4296
         Picture         =   "ADCC.frx":0F32
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   204
         Width           =   804
      End
      Begin VB.CommandButton cmdZoomMais 
         Caption         =   "Zoom +"
         Height          =   696
         Left            =   3480
         Picture         =   "ADCC.frx":123C
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   204
         Width           =   804
      End
      Begin CURRENCYEDITLib.CurrencyEdit TxtValor 
         Height          =   372
         Left            =   7116
         TabIndex        =   5
         Top             =   1512
         Width           =   2076
         _Version        =   65537
         _ExtentX        =   3662
         _ExtentY        =   656
         _StockProps     =   93
         ForeColor       =   8388608
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   11
         BackColor       =   -2147483643
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Valor"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   7164
         TabIndex        =   18
         Top             =   1248
         Width           =   468
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Conta"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   5316
         TabIndex        =   17
         Top             =   1248
         Width           =   888
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Agência"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4452
         TabIndex        =   16
         Top             =   1248
         Width           =   720
      End
      Begin VB.Label LblCMC7 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "CMC-7"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   168
         TabIndex        =   15
         Top             =   1236
         Width           =   636
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Autorização de Débito"
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   708
         TabIndex        =   14
         Top             =   408
         Width           =   1596
      End
      Begin VB.Image Image1 
         Height          =   384
         Left            =   180
         Picture         =   "ADCC.frx":1546
         Top             =   288
         Width           =   384
      End
   End
End
Attribute VB_Name = "ADCC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Declaração de Variáveis do RDO
Private qryRemoveTipoDocumento       As rdoQuery
Private qryGetADCCDuplicado          As rdoQuery
Private qryAtualizaDocumentoExcluido As rdoQuery
Private qryAtualizaADCC              As rdoQuery
Private qryGetADCC                   As rdoQuery
Private qryLeituraValorMaxADCC       As rdoQuery

'Declaração de Variáveis de trabalho
Private mForm As Form
Public Alterou As Boolean
Sub AjustesIniciais()

  'Setando as variáveis RDOQUERY
  Set qryGetADCCDuplicado = Geral.Banco.CreateQuery("", "{? = call GetADCCDuplicado (?,?,?)}")
  Set qryAtualizaDocumentoExcluido = Geral.Banco.CreateQuery("", "{call AtualizaDocumentoExcluido (?,?,?,?,?)}")
  Set qryAtualizaADCC = Geral.Banco.CreateQuery("", "{? = call AtualizaADCC (?,?,?,?,?,?,?,?)}")
  Set qryRemoveTipoDocumento = Geral.Banco.CreateQuery("", "{? = call RemoveTipoDocumento (?,?,?)}")
  Set qryLeituraValorMaxADCC = Geral.Banco.CreateQuery("", "{? = call LerParametro(?)}")
  
End Sub
Sub PesquisaADCC()

  On Error GoTo ERRO_PESQUISAADCC

  Dim sSql As String
  Dim RsADCC As rdoResultset
  Dim sCampo1 As String
  Dim sCampo2 As String
  Dim sCampo3 As String
  Dim svalor As String

  'Pesquisar a ADCC Atual e preencher os valores caso encontre
  sSql = Geral.DataProcessamento & " , " & Geral.Documento.IdDocto

  Set qryGetADCC = Geral.Banco.CreateQuery("", "{call GetADCC (" & sSql & ")}")

  Set RsADCC = qryGetADCC.OpenResultset(rdOpenStatic, rdConcurReadOnly)
  If Not RsADCC.EOF Then
    'Encontrou a Autorização -> Preencher os campos
    txtAgencia.Text = Format(RsADCC!Agencia, "0000")
    txtConta.Text = RsADCC!Conta
    TxtValor.Text = RsADCC!Valor * 100
  End If

  If Len(Trim(Geral.Documento.Leitura)) = 30 Then
    'Verificar se CMC7 está totalmente zerado
    If Geral.Documento.Leitura = String(30, "0") Or _
      Geral.Documento.Leitura = "409" & String(27, "0") Then
      txtAgencia.SetFocus
    End If
  End If

  txtCMC71.Text = Mid(Geral.Documento.Leitura, 1, 8)
  txtCMC72.Text = Mid(Geral.Documento.Leitura, 9, 10)
  txtCMC73.Text = Mid(Geral.Documento.Leitura, 19, 12)

  'Verificar codigo 409 e 256
  If Mid(Geral.Documento.Leitura, 1, 3) = "409" And Mid(Geral.Documento.Leitura, 9, 3) = "256" Then
    'Validar CMC7
    If TratarCamposCMC7(Geral.Documento.Leitura, sCampo1, sCampo2, sCampo3, svalor) Then
      txtAgencia.SetFocus
      Exit Sub
    End If
  End If

  txtCMC71.SetFocus

  Exit Sub

ERRO_PESQUISAADCC:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Selecionar Dados da Autorização.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Sub
Private Sub cmdConfirmar_Click()

'Valida preenchimento máximo de CMC7
   If VerificaPreenchimentoCMC7(Me) = False Then Exit Sub

   If SalvaADCC Then
      Alterou = True
      Me.Hide
   End If
End Sub
Private Sub cmdFrenteVerso_Click()

  mForm.cmdFrenteVerso_Click
End Sub
Private Sub cmdInverteCor_Click()

  mForm.cmdInverteCor_Click
End Sub
Private Sub cmdRotacao_Click()

  mForm.cmdRotacao_Click
End Sub
Private Sub CmdSair_Click()

  Alterou = False
  Me.Hide
End Sub
Private Sub cmdZoomMais_Click()

  mForm.cmdZoomMais_Click
End Sub
Private Sub cmdZoomMenos_Click()

  mForm.cmdZoomMenos_Click
End Sub
Function CamposOK() As Boolean

  'Primeiro Campo do CMC7
  If Len(Trim(txtCMC71.Text)) = 0 Then
    MsgBox "Informe o Primeiro Campo do CMC7.", vbInformation, App.Title
    CamposOK = False
    txtCMC71.SetFocus
    Exit Function
  End If

  'Segundo Campo do CMC7
  If Len(Trim(txtCMC72.Text)) = 0 Then
    MsgBox "Informe o Segundo Campo do CMC7.", vbInformation, App.Title
    CamposOK = False
    txtCMC72.SetFocus
    Exit Function
  End If

  'Terceiro Campo do CMC7
  If Len(Trim(txtCMC73.Text)) = 0 Then
    MsgBox "Informe o Terceiro Campo do CMC7.", vbInformation, App.Title
    CamposOK = False
    txtCMC73.SetFocus
    Exit Function
  End If

  'Agencia
  If Len(Trim(txtAgencia.Text)) = 0 Then
    MsgBox "Informe o Código da Agencia.", vbInformation, App.Title
    CamposOK = False
    txtAgencia.SetFocus
    Exit Function
  End If

  'Conta
  If Len(Trim(txtConta.Text)) = 0 Then
    MsgBox "Informe o Número da Conta.", vbInformation, App.Title
    CamposOK = False
    txtConta.SetFocus
    Exit Function
  End If

  'Valor da Autorização
  If Len(Trim(TxtValor.Text)) = 0 Or Val(TxtValor.Text) <= 0 Then
    MsgBox "Informe o Valor da Autorização.", vbInformation, App.Title
    CamposOK = False
    TxtValor.SetFocus
    Exit Function
  End If

  CamposOK = True
End Function
Function SalvaADCC() As Boolean

    On Error GoTo ERRO_SALVAADCC

    Dim vCMC7               As String
    Dim sCampo1             As String
    Dim sCampo2             As String
    Dim sCampo3             As String
    Dim svalor              As String
    Dim sTipo               As String
    Dim sTamanho            As Integer
    Dim strEncripta         As String
    SalvaADCC = False

    'Verificar se todos os campos estão preenchidos
    If CamposOK Then
        'Verificar se as tres primeiras posições do primeiro campo do CMC7 devem ser iguais à : 409
        If Mid(txtCMC71.Text, 1, 3) <> "409" And Val(txtCMC71.Text) <> 0 Then
            MsgBox "O CMC7 da Autorização não é válido.", vbInformation, App.Title
            txtCMC71.SetFocus
            Exit Function
        End If
    
        'Verificar se as tres primeiras posições do segundo campo do CMC7 devem ser iguais à : 256
        If Mid(txtCMC72.Text, 1, 3) <> "256" And Val(txtCMC72.Text) <> 0 Then
            MsgBox "O CMC7 da Autorização não é Válido.", vbInformation, App.Title
            txtCMC72.SetFocus
            Exit Function
        End If
    
        'Formatar o campo 'AGENCIA'
        txtAgencia.Text = Format(txtAgencia.Text, "0000")
        
        'Validar Agencia e Conta
        If Val(txtAgencia.Text) <> 0 And Val(txtConta.Text) <> 0 Then
            sTamanho = Len(Format(txtAgencia.Text, "0000") & Format(txtConta.Text, "0000000"))
            If Not Modulo10(Format(txtAgencia.Text, "0000") & Format(txtConta.Text, "0000000"), sTamanho) Then
                MsgBox "Agência e/ou Conta Inválidos!", vbInformation, App.Title
                txtAgencia.SetFocus
                Exit Function
            End If
        Else
            MsgBox "Agência e/ou Conta Inválidos!", vbInformation, App.Title
            txtAgencia.SetFocus
            Exit Function
        End If
    
        vCMC7 = txtCMC71.Text & txtCMC72.Text & txtCMC73.Text
    
        'Verificar se CMC7 está totalmente zerado
        If vCMC7 = String(30, "0") Or vCMC7 = "409" & String(27, "0") Then
            vCMC7 = "409" & String(27, "0")
        Else
            'Validar CMC7
            If Not TratarCamposCMC7(vCMC7, sCampo1, sCampo2, sCampo3, svalor) Then
                MsgBox "CMC7 Inválido.", vbInformation, App.Title
                'Verificar qual campo está zerado e posicionar o cursor
                If Val(sCampo1) = 0 Then
                    txtCMC71.SetFocus
                    Exit Function
                End If
                
                If Val(sCampo2) = 0 Then
                    txtCMC72.SetFocus
                    Exit Function
                End If
                
                If Val(sCampo3) = 0 Then
                    txtCMC73.SetFocus
                    Exit Function
                End If
            End If
        End If
    
        'Verificar se o valor da ADCC informada é maior que o limite permitido (Parâmetro)
        If CCur(TxtValor.Text) > CCur(Geral.ValorMaxADCC * 100) Then
            MsgBox "Valor informado na Autorização maior que o limite permitido.", vbInformation, App.Title
            TxtValor.SetFocus
            Exit Function
        End If
    
        'Verificar se o Documento pertence à outro Tipo
        If Geral.Documento.TipoDocto <> 4 And Geral.Documento.TipoDocto <> 0 Then
           With qryRemoveTipoDocumento
              .rdoParameters(1) = Geral.DataProcessamento     'Data Proc.
              .rdoParameters(2) = Geral.Documento.IdDocto     'IdDocto
              .rdoParameters(3) = Geral.Documento.TipoDocto   'Tipo do Documento
              .Execute
           End With
        End If
    
        'Atualiza campo Autenticação Digital
        strEncripta = G_EncriptaBO(4, CStr(Val(txtConta.Text)))
        If strEncripta = "" Then GoTo ERRO_SALVAADCC
        
        'Atualizar / Inserir Autorização (AtualizaADCC)
        With qryAtualizaADCC
           .rdoParameters(0).Direction = rdParamReturnValue
           .rdoParameters(1) = Geral.DataProcessamento                       'Data Proc.
           .rdoParameters(2) = Geral.Documento.IdDocto                       'IdDocto
           .rdoParameters(3) = vCMC7                                         'CMC7
           .rdoParameters(4) = txtAgencia.Text                               'Agencia
           .rdoParameters(5) = txtConta.Text                                 'Conta
           .rdoParameters(6) = Val(TxtValor.Text) / 100                      'Valor
           .rdoParameters(7) = 4                                             'TipoDocto
           .rdoParameters(8) = strEncripta                                   'Autenticacao Digital
           .Execute
        End With
    
        If qryAtualizaADCC(0).Value = 2 Then
           Geral.Documento.Status = "D"
        End If
    
        'Atualizar o Controle Global
        Geral.Documento.ValorTotal = Val(TxtValor.Text) / 100
        Geral.Documento.TipoDocto = 4
        Geral.Documento.Leitura = vCMC7
               
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Envia para confirmação somente se o usuario for terceiro e o documento ñ é duplicidade'
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If GrupoUsuario(Geral.Usuario, eG_TERCEIRO) And Geral.Documento.Status <> "D" Then
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Não faz nada caso não conseguiu atualizar o status do documento'
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If Not ConfirmaAgConta(Geral.Documento.IdDocto) Then
                MsgBox "Não foi possível enviar este documento para confirmação de Agência e Conta.", vbCritical
                Exit Function
            End If
            Geral.Documento.Status = "L"
        End If
       
        SalvaADCC = True
    End If

   Exit Function

ERRO_SALVAADCC:
   Screen.MousePointer = vbDefault
   Select Case TratamentoErro("Erro ao Atualizar Dados do ADCC.", Err, rdoErrors)
      Case vbCancel
      Case vbRetry
   End Select
End Function
Public Sub SetParent(ByRef aForm As Form)

  Set mForm = aForm
End Sub
Public Sub SetPosition(iLeft As Integer, iTop As Integer)

  Me.Left = iLeft
  Me.Top = iTop
End Sub
Private Sub Form_Activate()

  Call AjustesIniciais

  Call PesquisaADCC

  Screen.MousePointer = vbDefault
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

  Select Case KeyCode
    Case vbKeyAdd
      Call cmdZoomMais_Click
    Case vbKeySubtract
      Call cmdZoomMenos_Click
    Case vbKeyF10
      Call cmdInverteCor_Click
      KeyCode = 0
    Case vbKeyDivide
      Call cmdRotacao_Click
    Case vbKeyMultiply
      Call cmdConfirmar_Click
    Case vbKeyF11
      Call cmdFrenteVerso_Click
    Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight
      mForm.Form_KeyUp KeyCode, Shift
  End Select
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  Set qryGetADCCDuplicado = Nothing
  Set qryAtualizaDocumentoExcluido = Nothing
  Set qryAtualizaADCC = Nothing
  Set qryRemoveTipoDocumento = Nothing

End Sub

Private Sub txtAgencia_Change()
   If Len(Trim(txtAgencia.Text)) = txtAgencia.MaxLength Then
      SendKeys "{TAB}"
      DoEvents
   End If
End Sub
Private Sub txtAgencia_GotFocus()
   txtAgencia.SelStart = 0
   txtAgencia.SelLength = txtAgencia.MaxLength
End Sub
Private Sub txtAgencia_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      SendKeys "{TAB}"
      KeyAscii = 0
   ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
      KeyAscii = 0
   End If
End Sub

Private Sub txtAgencia_LostFocus()

    If Len(Trim(txtAgencia.Text)) = 0 Then Exit Sub

    'Valida Agencia
    If Not IsNumeric(txtAgencia) Then
        MsgBox "Número de Agência inválido, Redigite.", vbInformation, App.Title
        txtAgencia.Text = ""
        txtAgencia.SetFocus
        Exit Sub
    End If
End Sub
Private Sub txtCMC71_Change()

  If Len(Trim(txtCMC71.Text)) = txtCMC71.MaxLength Then
    SendKeys "{TAB}"
    DoEvents
  Else
    'Verificar se é a primeira posição e se é um zero
    If Left(txtCMC71.Text, 1) = "0" Then
      'Formatar os tres campos do CMC7 com zeros
      txtCMC71.Text = String(txtCMC71.MaxLength, "0")
      txtCMC72.Text = String(txtCMC72.MaxLength, "0")
      txtCMC73.Text = String(txtCMC73.MaxLength, "0")
      txtAgencia.SetFocus
    End If
  End If
End Sub

Private Sub txtCMC71_GotFocus()

  txtCMC71.SelStart = 0
  txtCMC71.SelLength = txtCMC71.MaxLength
End Sub


Private Sub txtCMC71_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub
Private Sub txtCMC72_Change()

  If Len(Trim(txtCMC72.Text)) = txtCMC72.MaxLength Then
    SendKeys "{TAB}"
    DoEvents
  End If
End Sub

Private Sub txtCMC72_GotFocus()

  txtCMC72.SelStart = 0
  txtCMC72.SelLength = txtCMC72.MaxLength
End Sub

Private Sub txtCMC72_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub
Private Sub txtCMC73_Change()

  If Len(Trim(txtCMC73.Text)) = txtCMC73.MaxLength Then
    SendKeys "{TAB}"
    DoEvents
  End If
End Sub

Private Sub txtCMC73_GotFocus()

  txtCMC73.SelStart = 0
  txtCMC73.SelLength = txtCMC73.MaxLength
End Sub

Private Sub txtCMC73_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub

Private Sub txtConta_Change()

   If Len(Trim(txtConta.Text)) = txtConta.MaxLength Then
      SendKeys "{TAB}"
      DoEvents
   End If
End Sub
Private Sub txtConta_GotFocus()

   txtConta.SelStart = 0
   txtConta.SelLength = txtConta.MaxLength
End Sub
Private Sub txtConta_KeyPress(KeyAscii As Integer)

   If KeyAscii = vbKeyReturn Then
      SendKeys "{TAB}"
      KeyAscii = 0
   ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
      KeyAscii = 0
   End If
End Sub
Private Sub txtConta_LostFocus()

    Dim sTamanho As String

    If Len(Trim(txtConta.Text)) = 0 Then Exit Sub

    'Validar Conta
    If Not IsNumeric(txtConta) Then
        MsgBox "Número de Conta inválido, Redigite.", vbInformation, App.Title
        txtConta.Text = ""
        txtConta.SetFocus
    End If

    If Val(txtAgencia.Text) <> 0 And Val(txtConta.Text) <> 0 Then
        sTamanho = Len(Format(txtAgencia.Text, "0000") & Format(txtConta.Text, "0000000"))
        If Not Modulo10(Format(txtAgencia.Text, "0000") & Format(txtConta.Text, "0000000"), sTamanho) Then
            MsgBox "Agência e/ou Conta Inválidos!", vbInformation, App.Title
            txtAgencia.SetFocus
            Exit Sub
        End If
    Else
        MsgBox "Agência e/ou Conta Inválidos!", vbInformation, App.Title
        txtAgencia.SetFocus
        Exit Sub
    End If
End Sub
Private Sub txtValor_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    Call cmdConfirmar_Click
  End If
End Sub


