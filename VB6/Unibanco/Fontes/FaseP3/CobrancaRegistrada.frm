VERSION 5.00
Object = "{9CBA5D64-E3C8-11D3-9FFC-00104BC8688C}#1.0#0"; "CurrencyEdit.ocx"
Object = "{4A8E27C4-EACA-11D3-9FFC-00104BC8688C}#1.0#0"; "DateEdit.ocx"
Begin VB.Form CobrancaRegistrada 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Complementação de Cobranças Registradas"
   ClientHeight    =   2328
   ClientLeft      =   1416
   ClientTop       =   1476
   ClientWidth     =   9300
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2328
   ScaleWidth      =   9300
   Begin VB.Frame Frame1 
      Height          =   2196
      Left            =   84
      TabIndex        =   13
      Top             =   60
      Width           =   9168
      Begin VB.CommandButton cmdZoomMais 
         Caption         =   "Zoom +"
         Height          =   696
         Left            =   3336
         Picture         =   "CobrancaRegistrada.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   276
         Width           =   804
      End
      Begin VB.CommandButton cmdZoomMenos 
         Caption         =   "Zoom -"
         Height          =   696
         Left            =   4152
         Picture         =   "CobrancaRegistrada.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   276
         Width           =   804
      End
      Begin VB.CommandButton cmdRotacao 
         Caption         =   "Rotação"
         Height          =   696
         Left            =   4968
         Picture         =   "CobrancaRegistrada.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   276
         Width           =   804
      End
      Begin VB.CommandButton cmdInverteCor 
         Caption         =   "Inverter"
         Height          =   696
         Left            =   5784
         Picture         =   "CobrancaRegistrada.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   276
         Width           =   804
      End
      Begin VB.CommandButton cmdFrenteVerso 
         Caption         =   "Fren/Ver"
         Height          =   696
         Left            =   6600
         Picture         =   "CobrancaRegistrada.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   276
         Width           =   804
      End
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   696
         Left            =   8208
         Picture         =   "CobrancaRegistrada.frx":0F32
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   276
         Width           =   816
      End
      Begin VB.CommandButton cmdConfirmar 
         Caption         =   "&Confirmar"
         Height          =   696
         Left            =   7404
         Picture         =   "CobrancaRegistrada.frx":123C
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   276
         Width           =   804
      End
      Begin DATEEDITLib.DateEdit TxtVencimento 
         Height          =   372
         Left            =   3753
         TabIndex        =   3
         Top             =   1668
         Width           =   1356
         _Version        =   65537
         _ExtentX        =   2392
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
         BackColor       =   -2147483643
      End
      Begin VB.TextBox txtCVT 
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
         Left            =   108
         MaxLength       =   5
         TabIndex        =   0
         Top             =   1668
         Width           =   732
      End
      Begin VB.TextBox txtNossoNumero 
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
         Left            =   1742
         MaxLength       =   15
         TabIndex        =   2
         Top             =   1668
         Width           =   1932
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
         Left            =   919
         MaxLength       =   4
         TabIndex        =   1
         Top             =   1668
         Width           =   744
      End
      Begin CURRENCYEDITLib.CurrencyEdit TxtValorBase 
         Height          =   372
         Left            =   5188
         TabIndex        =   4
         Top             =   1668
         Width           =   1896
         _Version        =   65537
         _ExtentX        =   3344
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
      Begin CURRENCYEDITLib.CurrencyEdit TxtValor 
         Height          =   372
         Left            =   7140
         TabIndex        =   5
         Top             =   1668
         Width           =   1896
         _Version        =   65537
         _ExtentX        =   3344
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
         Locked          =   -1  'True
      End
      Begin VB.Image Image1 
         Height          =   384
         Left            =   228
         Picture         =   "CobrancaRegistrada.frx":1546
         Top             =   396
         Width           =   384
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Cobrança Registrada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   792
         TabIndex        =   20
         Top             =   528
         Width           =   1884
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Vencimento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3888
         TabIndex        =   19
         Top             =   1392
         Width           =   1056
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "CVT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   156
         TabIndex        =   18
         Top             =   1392
         Width           =   396
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Nosso Número"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1908
         TabIndex        =   17
         Top             =   1392
         Width           =   1332
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Agência"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   996
         TabIndex        =   16
         Top             =   1392
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Base"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   5352
         TabIndex        =   15
         Top             =   1392
         Width           =   984
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Cobrado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   7248
         TabIndex        =   14
         Top             =   1392
         Width           =   1272
      End
   End
End
Attribute VB_Name = "CobrancaRegistrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Declaração de Variáveis do RDO
'Private qryGetCobrancaRegistradaDuplicada As rdoQuery
'Private qryAtualizaDocumentoExcluido As rdoQuery
Private qryAtualizaCobrancaRegistrada As rdoQuery
Private qryRemoveTipoDocumento As rdoQuery
Private qryGetCobrancaRegistrada As rdoQuery

'Declaração de Variáveis de trabalho
Private mForm As Form
Public Alterou As Boolean
Public AlteraValor As Boolean

Sub AjustesIniciais()

  On Error GoTo ERRO_AJUSTESINICIAIS

  'Setar os Objetos RDOQuery
'  Set qryGetCobrancaRegistradaDuplicada = Geral.Banco.CreateQuery("", "{? = call GetCobracaRegistradaDuplicada (?,?,?)}")
'  Set qryAtualizaDocumentoExcluido = Geral.Banco.CreateQuery("", "{call AtualizaDocumentoExcluido (?,?,?,?,?)}")
  Set qryAtualizaCobrancaRegistrada = Geral.Banco.CreateQuery("", "{call AtualizaCobrancaRegistrada (?,?,?,?,?,?,?,?,?)}")
  Set qryRemoveTipoDocumento = Geral.Banco.CreateQuery("", "{? = call RemoveTipoDocumento (?,?,?)}")

  Exit Sub

ERRO_AJUSTESINICIAIS:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Atualizar Dados da Cobrança Registrada.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Sub
Function SalvaCobrancaRegistrada()

  On Error GoTo ERRO_SALVACOBRANCAREG

  Dim RetAgencia As Integer
  Dim strEncripta   As String

  SalvaCobrancaRegistrada = False

  'Preencher o Campo 'Valor Cobrado'
  If Len(Trim(TxtValorBase.Text)) <> 0 Then
    TxtValor.Text = TxtValorBase.Text
  End If

  'Formatar o Campo 'Nosso Numero'
  txtNossoNumero.Text = Format(txtNossoNumero.Text, String(15, "0"))

  'Verificar se todos os campos estão preenchidos
  If CamposOK Then

    'Validar Campo NossoNumero
    If Not ValidaNossoNumero(txtNossoNumero.Text) Then
      MsgBox "O Campo 'Nosso Número' não é válido.", vbInformation, App.Title
      txtNossoNumero.SetFocus
      Exit Function
    End If

    'Se não existe informação da agência carregada, verificar (ValidaAgencia)
    If Not ValidaAgenciaPorDocto(Geral.Documento.Agencia, TxtVencimento.Text, True) Then
        TxtVencimento.SetFocus
        Exit Function
    End If
    
'    'Se não existe informação da agência carregada, verificar (ValidaAgencia)
'    RetAgencia = ValidaAgencia(Geral.Documento.Agencia, TxtVencimento.Text, True)
    
'    'Verificar Retorno da Função
'    Select Case RetAgencia
'        '08/05/2001''''''''''''''''''''''''''''''''''''''''
'        'Pode aceitar Cobrança Registrada vencida         '
'        'Comentado o "Case 1"                             '
'        '''''''''''''''''''''''''''''''''''''''''''''''''''
''      Case 1
''
''        'Documento Vencido
''        If Geral.capa.IdEnv_Mal = "E" Then
''            'Envelope -> Não Aceitar
''            MsgBox "Documento vencido não aceito na regra de caixa expresso.", vbInformation + vbOKOnly, App.Title
''            TxtVencimento.SetFocus
''            Exit Function
''        ElseIf Geral.capa.IdEnv_Mal = "M" Then
''            'Malote -> Pedir Confirmação
''            If MsgBox("Este documento pertence a um Malote e está vencido. Confirma ?", vbYesNo + vbInformation, App.Title) = vbNo Then
''              TxtVencimento.SetFocus
''              Exit Function
''            End If
''        Else
''            'Tipo Indefinido
''            MsgBox "Não foi possível definir se o documento pertence a um Envelope ou Malote " & Chr(13) & _
''            "para aplicar regra de validação de Data de Vencimento.", vbInformation + vbOKOnly, App.Title
''            Exit Function
''        End If
'      Case 2
'        'Agencia em Feriado
'        MsgBox "A agência de origem está em feriado.", vbInformation + vbOKOnly, App.Title
'        TxtVencimento.SetFocus
'        Exit Function
'      Case 3
'        'Agencia Fechada
'        MsgBox "A agência de origem está fechada.", vbInformation + vbOKOnly, App.Title
'        TxtVencimento.SetFocus
'        Exit Function
'      Case 4
'        'Agencia não Cadastrada
'        MsgBox "A agência de origem não está cadastrada.", vbInformation + vbOKOnly, App.Title
'        TxtVencimento.SetFocus
'        Exit Function
'    End Select

    'Verificar se o Documento pertence à outro Tipo
    If Geral.Documento.TipoDocto <> 13 And Geral.Documento.TipoDocto <> 0 Then
      With qryRemoveTipoDocumento
        .rdoParameters(1) = Geral.DataProcessamento     'Data Proc.
        .rdoParameters(2) = Geral.Documento.IdDocto     'IdDocto
        .rdoParameters(3) = Geral.Documento.TipoDocto   'Tipo do Documento
        .Execute
      End With
    End If

    'Atualiza campo Autenticação Digital
    strEncripta = G_EncriptaBO(13, CStr(Val(txtNossoNumero.Text)))
    If strEncripta = "" Then GoTo ERRO_SALVACOBRANCAREG

    'Atualizar / Inserir Cobranca Registrada
    With qryAtualizaCobrancaRegistrada
      .rdoParameters(0) = Geral.DataProcessamento                   'Data Proc.
      .rdoParameters(1) = Geral.Documento.IdDocto                   'IdDocto
      .rdoParameters(2) = txtCVT.Text                               'CVT
      .rdoParameters(3) = TxtVencimento.InverseText                 'Vencimento
      .rdoParameters(4) = txtAgencia.Text                           'Agencia
      .rdoParameters(5) = txtNossoNumero.Text                       'Nosso Numero
      .rdoParameters(6) = TxtValor.Text / 100                       'Valor Cobrado
      .rdoParameters(7) = 13                                        'TipoDocto
      .rdoParameters(8) = strEncripta                               'Autenticacao digital
      .Execute
    End With

    'Atualizar o Controle Global
    Geral.Documento.ValorTotal = Val(TxtValor.Text) / 100
    Geral.Documento.Leitura = ""
    Geral.Documento.TipoDocto = 13

    SalvaCobrancaRegistrada = True
  End If

  Exit Function

ERRO_SALVACOBRANCAREG:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Atualizar Dados da Cobrança Registrada.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
  Resume
End Function

Private Sub cmdConfirmar_Click()

  If SalvaCobrancaRegistrada Then
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


Private Sub Form_Activate()

  Call AjustesIniciais

  Call PesquisaCobrancaRegistrada
End Sub


Function CamposOK() As Boolean

  'CVT Preenchido
  If Len(Trim(txtCVT.Text)) = 0 Then
    MsgBox "Informe o Código do CVT.", vbInformation, App.Title
    CamposOK = False
    txtCVT.SetFocus
    Exit Function
  End If

  'Valida CVT
  If txtCVT.Text <> "55336" And txtCVT.Text <> "55360" And txtCVT.Text <> "55395" Then
    MsgBox "Código de CVT Inválido.", vbInformation, App.Title
    CamposOK = False
    txtCVT.SetFocus
    Exit Function
  End If

  'Vencimento
  If Len(Trim(TxtVencimento.Text)) = 0 Then
    MsgBox "Informe a Data de Vencimento do Documento.", vbInformation, App.Title
    CamposOK = False
    TxtVencimento.SetFocus
    Exit Function
  End If

  'Agencia
  If Len(Trim(txtAgencia.Text)) = 0 Then
    MsgBox "Informe o Código da Agência.", vbInformation, App.Title
    CamposOK = False
    txtAgencia.SetFocus
    Exit Function
  End If

  'Nosso Numero
  If Len(Trim(txtNossoNumero.Text)) = 0 Then
    MsgBox "Informe o Nosso Numero.", vbInformation, App.Title
    CamposOK = False
    txtNossoNumero.SetFocus
    Exit Function
  End If

  'Valor Base
  If Len(Trim(TxtValorBase.Text)) = 0 Then
    MsgBox "Informe o Valor Base.", vbInformation, App.Title
    CamposOK = False
    TxtValorBase.SetFocus
    Exit Function
  End If

  'Valor
  If Len(Trim(TxtValor.Text)) = 0 Then
    MsgBox "Informe o Valor Cobrado.", vbInformation, App.Title
    CamposOK = False
    TxtValor.SetFocus
    Exit Function
  End If

  CamposOK = True
End Function
Sub PesquisaCobrancaRegistrada()

  On Error GoTo ERRO_PESQUISACOBRANCAREG

  Dim sSql As String
  Dim RsCobrancaReg As rdoResultset

  'Pesquisar o Documento Atual e preencher os valores caso encontre
  sSql = Geral.DataProcessamento & " , " & Geral.Documento.IdDocto

  Set qryGetCobrancaRegistrada = Geral.Banco.CreateQuery("", "{call GetCobrancaRegistrada (" & sSql & ")}")

  Set RsCobrancaReg = qryGetCobrancaRegistrada.OpenResultset(rdOpenStatic, rdConcurReadOnly)
  If Not RsCobrancaReg.EOF Then
    'Encontrou o Documento -> Preencher os campos
    txtCVT.Text = RsCobrancaReg!CVT
    TxtVencimento.Text = Format(DataDDMMAAAA(RsCobrancaReg!vecto), "00000000")
    txtAgencia.Text = RsCobrancaReg!Agencia
    txtNossoNumero.Text = RsCobrancaReg!NossoNumero
    TxtValorBase.Text = RsCobrancaReg!ValorBase * 100
    TxtValor.Text = RsCobrancaReg!Valor * 100

    TxtVencimento.SetFocus
  Else
    txtCVT.SetFocus
  End If

  If AlteraValor = True Then
    'O Usuário só pode alterar os campos de valor
    txtCVT.Locked = True
    txtAgencia.Locked = True
    txtNossoNumero.Locked = True
    TxtVencimento.Locked = True

    TxtValorBase.SetFocus
  End If

  Screen.MousePointer = vbDefault

  Exit Sub

ERRO_PESQUISACOBRANCAREG:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Selecionar Dados da Cobrança Registrada.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Sub
Public Sub SetPosition(iLeft As Integer, iTop As Integer)

  Me.Left = iLeft
  Me.Top = iTop
End Sub
Public Sub SetParent(ByRef aForm As Form)

  Set mForm = aForm
End Sub
Public Function ValidaNossoNumero(ByVal pvsNossoNum As String) As Boolean

  On Error GoTo ERRO_VALIDANOSSONUMERO

  'Retorna True se a agencia for válida
  'MODULO 11 (2 BASE 9)
  'Esta rotina serve para conferir o Nosso Numero

  Dim soma As Integer, resto As Integer
  Dim digito_11 As Integer, p As Integer, peso As Integer
  Dim digito_rv As String
  Dim bOk As Boolean

  bOk = True          'default-True-nossonumero OK

  soma = 0
  resto = 0
  digito_11 = 0       'calculado pelo módulo 11
  digito_rv = ""      'caracter digitado pelo operador

  peso = 2            'começa multiplicar da direita para esquerda
  p = 13

  Do
      '*** Peso de 2 a 9 (multiplicação dos caracteres de 2 a 9) ***
      soma = soma + Mid(pvsNossoNum, p, 1) * peso
      p = p - 1       'ponteiro
      peso = peso + 1 'peso
      If (peso = 10) Then
         peso = 2
      End If
      If (p = 3) Then
         Exit Do
      End If
  Loop

  resto = soma Mod 11        'resto da divisão
  digito_11 = 11 - resto     'digito verificador

  '*** se o calculo for igual a 10 ou 11, muda-se para 0 ***
  If (digito_11 = 11) Or (digito_11 = 10) Then
     digito_11 = 0
  End If

  digito_rv = Mid(pvsNossoNum, 14, 1) 'digito verificador

  If CStr(digito_11) <> (digito_rv) Then
      bOk = False                     'digito não confere
      Exit Function
  End If

  soma = 0
  resto = 0
  digito_11 = 0           'calculado pelo módulo 11
  digito_rv = ""          'caracter digitado pelo operador
  peso = 2                'começa multiplicar da direita para esquerda
  p = 14

  Do
      '*** Peso de 2 a 9 (multiplicação dos caracteres de 2 a 9) ***
      soma = soma + Mid(pvsNossoNum, p, 1) * peso
      p = p - 1            'ponteiro
      peso = peso + 1      'peso
      If (peso = 10) Then
         peso = 2
      End If
      If (p = 2) Then
         Exit Do
      End If
  Loop

  resto = soma Mod 11        'resto da divisão
  digito_11 = 11 - resto     'digito verificador

  '*** se o calculo for igual a 10 ou 11, muda-se para 0 ***
  If (digito_11 = 11) Or (digito_11 = 10) Then
      digito_11 = 0
  End If

  digito_rv = Mid(pvsNossoNum, 15, 1) 'digito verificador

  If CStr(digito_11) <> (digito_rv) Then
      bOk = False                     'digito não confere
      Exit Function
  End If

  ValidaNossoNumero = bOk

  Exit Function

ERRO_VALIDANOSSONUMERO:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Validar Campo 'Nosso Numero'.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function
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
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub


Private Sub txtCVT_Change()

  If Len(Trim(txtCVT.Text)) = txtCVT.MaxLength Then
    SendKeys "{TAB}"
    DoEvents
  End If
End Sub

Private Sub txtCVT_GotFocus()

  txtCVT.SelStart = 0
  txtCVT.SelLength = txtCVT.MaxLength
End Sub

Private Sub txtCVT_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub

Private Sub txtCVT_LostFocus()

   If Len(Trim(txtCVT.Text)) > 0 And Val(txtCVT.Text) <> 0 Then
      'Valida CVT
      If txtCVT.Text <> "55336" And txtCVT.Text <> "55360" And txtCVT.Text <> "55395" Then
         MsgBox "Código de CVT Inválido.", vbInformation, App.Title
         txtCVT.Text = ""
         txtCVT.SetFocus
         Exit Sub
      End If
   End If
End Sub
Private Sub txtNossoNumero_Change()

   'If Len(Trim(txtNossoNumero.Text)) = txtNossoNumero.MaxLength Then
   '   SendKeys "{TAB}"
   '   DoEvents
   'End If
End Sub
Private Sub txtNossoNumero_GotFocus()

  txtNossoNumero.SelStart = 0
  txtNossoNumero.SelLength = txtNossoNumero.MaxLength
End Sub
Private Sub txtNossoNumero_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub
Private Sub txtNossoNumero_LostFocus()

   If Len(Trim(txtNossoNumero.Text)) > 0 Then

      'Formatar o Campo 'Nosso Numero'
      txtNossoNumero.Text = Format(txtNossoNumero.Text, String(15, "0"))

      'Validar Campo NossoNumero
      If Not ValidaNossoNumero(txtNossoNumero.Text) Then
         MsgBox "O Campo 'Nosso Número' não é válido.", vbInformation, App.Title
         txtNossoNumero.Text = ""
         txtNossoNumero.SetFocus
         Exit Sub
      End If
   End If
End Sub
Private Sub txtValor_GotFocus()

  TxtValor.SelStart = 0
  TxtValor.SelLength = Len(TxtValor.Text)
End Sub
Private Sub txtValor_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    Call cmdConfirmar_Click
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub
Private Sub TxtValorBase_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub

Private Sub TxtValorBase_LostFocus()

  If Len(Trim(TxtValorBase.Text)) <> 0 Then
    TxtValor.Text = TxtValorBase.Text
  End If
End Sub
Private Sub TxtVencimento_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf KeyAscii = vbKeySpace And TxtVencimento.Locked = False Then
    TxtVencimento.Text = Mid(Geral.DataProcessamento, 7, 2) & Mid(Geral.DataProcessamento, 5, 2) & Mid(Geral.DataProcessamento, 1, 4)
    KeyAscii = 0
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub
