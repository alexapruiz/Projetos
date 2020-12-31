VERSION 5.00
Object = "{9CBA5D64-E3C8-11D3-9FFC-00104BC8688C}#1.0#0"; "CurrencyEdit.ocx"
Object = "{4A8E27C4-EACA-11D3-9FFC-00104BC8688C}#1.0#0"; "DateEdit.ocx"
Begin VB.Form CobrancaEspecial 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Complementação de Cobranças Especiais"
   ClientHeight    =   2892
   ClientLeft      =   1296
   ClientTop       =   1308
   ClientWidth     =   8616
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2892
   ScaleWidth      =   8616
   Begin VB.Frame Frame1 
      Height          =   2844
      Left            =   84
      TabIndex        =   17
      Top             =   -12
      Width           =   8460
      Begin VB.TextBox TxtCedente 
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
         Left            =   1724
         MaxLength       =   7
         TabIndex        =   2
         Top             =   1488
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
         Left            =   964
         MaxLength       =   4
         TabIndex        =   1
         Top             =   1488
         Width           =   636
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
         Left            =   2808
         MaxLength       =   15
         TabIndex        =   3
         Top             =   1488
         Width           =   1932
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
         Top             =   1488
         Width           =   732
      End
      Begin VB.CommandButton cmdConfirmar 
         Caption         =   "&Confirmar"
         Height          =   696
         Left            =   6588
         Picture         =   "CobrancaEspecial.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   276
         Width           =   804
      End
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "Sair"
         Height          =   696
         Left            =   7404
         Picture         =   "CobrancaEspecial.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   276
         Width           =   816
      End
      Begin VB.CommandButton cmdFrenteVerso 
         Caption         =   "Fren/Ver"
         Height          =   696
         Left            =   5796
         Picture         =   "CobrancaEspecial.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   276
         Width           =   780
      End
      Begin VB.CommandButton cmdInverteCor 
         Caption         =   "Inverter"
         Height          =   696
         Left            =   4980
         Picture         =   "CobrancaEspecial.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   276
         Width           =   804
      End
      Begin VB.CommandButton cmdRotacao 
         Caption         =   "Rotação"
         Height          =   696
         Left            =   4164
         Picture         =   "CobrancaEspecial.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   276
         Width           =   804
      End
      Begin VB.CommandButton cmdZoomMenos 
         Caption         =   "Zoom -"
         Height          =   696
         Left            =   3348
         Picture         =   "CobrancaEspecial.frx":0F32
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   276
         Width           =   804
      End
      Begin VB.CommandButton cmdZoomMais 
         Caption         =   "Zoom +"
         Height          =   696
         Left            =   2532
         Picture         =   "CobrancaEspecial.frx":123C
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   276
         Width           =   804
      End
      Begin DATEEDITLib.DateEdit TxtVencimento 
         Height          =   372
         Left            =   4864
         TabIndex        =   4
         Top             =   1488
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
      Begin CURRENCYEDITLib.CurrencyEdit TxtValorBase 
         Height          =   372
         Left            =   6348
         TabIndex        =   5
         Top             =   1488
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
      Begin CURRENCYEDITLib.CurrencyEdit TxtAbatimento 
         Height          =   372
         Left            =   4268
         TabIndex        =   8
         Top             =   2328
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
      Begin CURRENCYEDITLib.CurrencyEdit txtDesconto 
         Height          =   372
         Left            =   2188
         TabIndex        =   7
         Top             =   2328
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
      Begin CURRENCYEDITLib.CurrencyEdit TxtJuros 
         Height          =   372
         Left            =   108
         TabIndex        =   6
         Top             =   2328
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
         Left            =   6348
         TabIndex        =   9
         Top             =   2328
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
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "( - ) Abatimento"
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
         Left            =   4320
         TabIndex        =   28
         Top             =   2064
         Width           =   1368
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "( - ) Desconto"
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
         Left            =   2220
         TabIndex        =   27
         Top             =   2064
         Width           =   1224
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "( + ) Juros/Mora"
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
         TabIndex        =   26
         Top             =   2064
         Width           =   1428
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Cedente"
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
         Left            =   1848
         TabIndex        =   25
         Top             =   1212
         Width           =   744
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
         Left            =   6480
         TabIndex        =   24
         Top             =   2064
         Width           =   1272
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
         Left            =   6540
         TabIndex        =   23
         Top             =   1212
         Width           =   984
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
         Left            =   960
         TabIndex        =   22
         Top             =   1212
         Width           =   720
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
         Left            =   2940
         TabIndex        =   21
         Top             =   1212
         Width           =   1332
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
         TabIndex        =   20
         Top             =   1212
         Width           =   396
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
         Left            =   5016
         TabIndex        =   19
         Top             =   1212
         Width           =   1056
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Cobrança Especial"
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   744
         TabIndex        =   18
         Top             =   528
         Width           =   1380
      End
      Begin VB.Image Image1 
         Height          =   384
         Left            =   228
         Picture         =   "CobrancaEspecial.frx":1546
         Top             =   396
         Width           =   384
      End
   End
End
Attribute VB_Name = "CobrancaEspecial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Declaração de Variáveis do RDO
'Private qryAtualizaDocumentoExcluido As rdoQuery
Private qryAtualizaCobrancaEspecial As rdoQuery
Private qryRemoveTipoDocumento As rdoQuery
Private qryGetCobrancaEspecial As rdoQuery

'Declaração de variáveis de trabalho
Public mForm As Form
Public Alterou As Boolean
Public AlteraValor As Boolean

Public Function ValidaNossoNumero_Esp() As Byte

  On Error GoTo ERRO_VALIDANOSSONUMERO_ESP

  Dim soma As Integer, resto As Integer
  Dim digito_11 As Integer, p As Integer, peso As Integer
  Dim digito_rv As String
  Dim bOk As Byte

  bOk = 1     'default - ok

  If (txtNossoNumero.Text = "000000000000000") And (txtCVT.Text = "77330") Then
    Select Case Val(txtAgencia)
      Case 98
        If (TxtCedente.Text <> "1189731") And (TxtCedente.Text <> "1187941") Then
          bOk = 2
          ValidaNossoNumero_Esp = bOk
          Exit Function
        End If
      Case 318
        If (TxtCedente.Text <> "1093794") And (TxtCedente.Text <> "1094701") Then
          bOk = 2
          ValidaNossoNumero_Esp = bOk
          Exit Function
        End If
      Case 926
        If (TxtCedente.Text <> "1010004") And (TxtCedente.Text <> "1015557") And (TxtCedente.Text <> "1014006") Then
          bOk = 2
          ValidaNossoNumero_Esp = bOk
          Exit Function
        End If
      Case Else
        bOk = 2
        ValidaNossoNumero_Esp = bOk
        Exit Function
    End Select
  End If

  If (txtAgencia.Text = "0098") And (txtCVT.Text = "77330") And (TxtCedente.Text = "1189731") Then
    soma = 0
    resto = 0
    digito_11 = 0       'calculado pelo módulo 11
    digito_rv = ""      'caracter digitado pelo operador

    peso = 2            'começa multiplicar da direita para esquerda
    p = 14
     
    Do
      '*** Peso de 2 a 9 (multiplicação dos caracteres de 2 a 9) ***
      soma = soma + Mid(txtNossoNumero.Text, p, 1) * peso
      p = p - 1            'ponteiro
      peso = peso + 1      'peso
      If (peso = 10) Then
        peso = 2
      End If
      If (p = 6) Then
        Exit Do
      End If
    Loop

    resto = soma Mod 11        'resto da divisão
    digito_11 = 11 - resto     'digito verificador

    '*** se o calculo for igual a 10 ou 11, muda-se para 0 ***
    If (digito_11 = 11) Or (digito_11 = 10) Then
      digito_11 = 0
    End If

    digito_rv = Mid(txtNossoNumero.Text, 15, 1) 'digito verificador

    If CStr(digito_11) <> (digito_rv) Then
      bOk = 3             'digito não confere
    End If

    ValidaNossoNumero_Esp = bOk    'retorna 1 ou 3
    Exit Function
  Else
    If (txtAgencia.Text = "0098") And (txtCVT.Text = "77330") And (TxtCedente.Text = "1187941") Then
      soma = 0
      resto = 0
      digito_11 = 0       'calculado pelo módulo 11
      digito_rv = ""      'caracter digitado pelo operador

      peso = 2            'começa multiplicar da direita para esquerda
      p = 14

      Do
        '*** Peso de 2 a 9 (multiplicação dos caracteres de 2 a 9) ***
        soma = soma + Mid(txtNossoNumero.Text, p, 1) * peso
        p = p - 1            'ponteiro
        peso = peso + 1      'peso
        If (peso = 10) Then
          peso = 2
        End If
        If (p = 7) Then
          Exit Do
        End If
      Loop

      resto = soma Mod 11     'resto da divisão
      digito_11 = 11 - resto  'digito verificador

      '*** se o calculo for igual a 10 ou 11, muda-se para 0 ***
      If (digito_11 = 11) Or (digito_11 = 10) Then
        digito_11 = 0
      End If

      digito_rv = Mid(txtNossoNumero.Text, 15, 1)     'digito verificador

      If CStr(digito_11) <> (digito_rv) Then
        bOk = 3             'digito não confere
      End If
      ValidaNossoNumero_Esp = bOk    'retorna 1 ou 3
      Exit Function
    End If
  End If

  bOk = 1             'default - ok
  soma = 0
  resto = 0
  digito_11 = 0       'calculado pelo módulo 11
  digito_rv = ""      'caracter digitado pelo operador

  peso = 2            'começa multiplicar da direita para esquerda
  p = 14

  Do
    '*** Peso de 2 a 9 (multiplicação dos caracteres de 2 a 9) ***
    soma = soma + Mid(txtNossoNumero.Text, p, 1) * peso
    p = p - 1            'ponteiro
    peso = peso + 1      'peso
    If (peso = 10) Then
      peso = 2
    End If
    If (p = 0) Then
      Exit Do
    End If
  Loop

  resto = soma Mod 11     'resto da divisão
  digito_11 = 11 - resto  'digito verificador

  '*** se o calculo for igual a 10 ou 11, muda-se para 0 ***
  If (digito_11 = 11) Or (digito_11 = 10) Then
    digito_11 = 0
  End If

  digito_rv = Mid(txtNossoNumero.Text, 15, 1)  'digito verificador

  If CStr(digito_11) <> (digito_rv) Then
    bOk = 3             'digito não confere
  End If

  ValidaNossoNumero_Esp = bOk    'retorna 1 ou 3

  Exit Function

ERRO_VALIDANOSSONUMERO_ESP:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Validar Campo 'Nosso Numero'.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function
Private Sub cmdConfirmar_Click()

  If SalvaCobrancaEspecial Then
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

  Call PesquisaCobrancaEspecial
End Sub
Private Function VerificarCodigoCedente() As Boolean

  VerificarCodigoCedente = False

  If (Val(TxtCedente.Text) >= 68000) And (Val(TxtCedente.Text) <= 68900) Then
    MsgBox "Codigo do Cedente não permitido.", vbInformation, App.Title
    TxtCedente.SetFocus
    Exit Function
  End If

  txtAgencia.Text = Format(txtAgencia.Text, "0000")

  TxtCedente.Text = Format(TxtCedente.Text, "0000000")

  'Calcula o digito verificador da agencia + codigo_cedente
  If Not Modulo10(txtAgencia.Text & TxtCedente.Text, 11) Then
    MsgBox "Número de Agência e Código Cedente não confere.", vbInformation, App.Title
    Exit Function
  End If

  VerificarCodigoCedente = True
End Function
Sub CalculaValorCobrado()

  On Error GoTo ERRO_CALCULAVALORCOBRADO

  Dim Valor As Currency

  'Verificar se foi informado o Valor Base
  If Val(TxtValorBase.Text) = 0 Then
    TxtValor.Text = ""
    Exit Sub
  End If

  'Verificar se foi informado Juros
  If Val(TxtJuros.Text) <> 0 Then
    Valor = Val(TxtValorBase.Text) + Val(TxtJuros.Text)
  Else
    Valor = TxtValorBase.Text
  End If

  'Verificar se foi informado Descontos
  If Val(txtDesconto.Text) <> 0 Then
    Valor = Valor - Val(txtDesconto.Text)
  End If

  'Verificar se foi informado Abatimentos
  If Val(TxtAbatimento.Text) <> 0 Then
    Valor = Valor - Val(TxtAbatimento.Text)
  End If

  'Transportar o Valor Final para a tela
  TxtValor.Text = Valor

  Exit Sub

ERRO_CALCULAVALORCOBRADO:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Calcular Valor Cobrado.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Sub
Function SalvaCobrancaEspecial()

  On Error GoTo ERRO_SALVACOBRANCAESP

  Dim RetAgencia As Integer
  Dim strEncripta   As String
  
  SalvaCobrancaEspecial = False

  'Formatar o Campo 'Nosso Numero'
  txtNossoNumero.Text = Format(txtNossoNumero.Text, String(15, "0"))

  'Verificar se todos os campos estão preenchidos
  If CamposOK Then

    'Validar Agencia
    If Not ValidaAgenciaPorDocto(Geral.Documento.Agencia, TxtVencimento.Text, True) Then
        TxtVencimento.SetFocus
        Exit Function
    End If
    
'    'Se não existe informação da agência carregada, verificar (ValidaAgencia)
'    RetAgencia = ValidaAgencia(Geral.Documento.Agencia, TxtVencimento.Text, True)
    
'    'Verificar Retorno da Função
'    Select Case RetAgencia
'
'        '08/05/2001''''''''''''''''''''''''''''''''''''''''
'        'Pode aceitar Cobrança Especial vencida           '
'        'Comentado o "Case 1"                             '
'        '''''''''''''''''''''''''''''''''''''''''''''''''''
'
''      Case 1
''        'Documento Vencido
''        If Geral.capa.IdEnv_Mal = "E" Then
''          'Envelope -> Não Aceitar
''          MsgBox "Documento vencido não aceito na regra de caixa expresso.", vbInformation + vbOKOnly, App.Title
''          TxtVencimento.SetFocus
''          Exit Function
''        ElseIf Geral.capa.IdEnv_Mal = "M" Then
''          'Malote -> Pedir Confirmação
''          If MsgBox("Este documento pertence a um Malote e está vencido. Confirma ?", vbYesNo + vbInformation, App.Title) = vbNo Then
''            TxtVencimento.SetFocus
''            Exit Function
''          End If
''        Else
''          'Tipo Indefinido
''          MsgBox "Não foi possível definir se o documento pertence a um Envelope ou Malote " & Chr(13) & _
''          "para aplicar regra de validação de Data de Vencimento.", vbInformation + vbOKOnly, App.Title
''          Exit Function
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

    'Validar Cedente
    If Not VerificarCodigoCedente Then Exit Function

    'Validar Campo NossoNumero
    If ValidaNossoNumero_Esp <> 1 Then
      MsgBox "O Campo 'Nosso Número' não é válido.", vbInformation, App.Title
      txtNossoNumero.SetFocus
      Exit Function
    End If

    'Verificar se o Documento pertence à outro Tipo
    If Geral.Documento.TipoDocto <> 14 And Geral.Documento.TipoDocto <> 0 Then
      With qryRemoveTipoDocumento
        .rdoParameters(1) = Geral.DataProcessamento     'Data Proc.
        .rdoParameters(2) = Geral.Documento.IdDocto     'IdDocto
        .rdoParameters(3) = Geral.Documento.TipoDocto   'Tipo do Documento
        .Execute
      End With
    End If

    'Atualiza campo Autenticação Digital
    strEncripta = G_EncriptaBO(14, CStr(Val(TxtCedente.Text)))
    If strEncripta = "" Then GoTo ERRO_SALVACOBRANCAESP

    'Atualizar / Inserir Cobranca Especial
    With qryAtualizaCobrancaEspecial
      .rdoParameters(0) = Geral.DataProcessamento                   'Data Proc.
      .rdoParameters(1) = Geral.Documento.IdDocto                   'IdDocto
      .rdoParameters(2) = txtCVT.Text                               'CVT
      .rdoParameters(3) = TxtVencimento.InverseText                 'Vencimento
      .rdoParameters(4) = txtAgencia.Text                           'Agencia
      .rdoParameters(5) = TxtCedente.Text                           'Cedente
      .rdoParameters(6) = txtNossoNumero.Text                       'Nosso Numero
      .rdoParameters(7) = CCur(TxtValorBase.Text / 100)             'Valor Base
      .rdoParameters(8) = Val(TxtJuros.Text) / 100                  'Juros
      .rdoParameters(9) = Val(txtDesconto.Text) / 100               'Descontos
      .rdoParameters(10) = Val(TxtAbatimento.Text) / 100            'Abatimentos
      .rdoParameters(11) = Val(TxtValor.Text) / 100                 'Valor Cobrado
      .rdoParameters(12) = 14                                       'TipoDocto
      .rdoParameters(13) = strEncripta                               'Autenticacao digital
      .Execute
    End With

    'Atualizar o Controle Global
    Geral.Documento.ValorTotal = Val(TxtValor.Text) / 100
    Geral.Documento.Leitura = ""
    Geral.Documento.TipoDocto = 14

    SalvaCobrancaEspecial = True
  End If

  Exit Function

ERRO_SALVACOBRANCAESP:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Atualizar Dados da Cobrança Especial.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function
Function CamposOK() As Boolean

  'CVT Preenchido
  If Len(Trim(txtCVT.Text)) = 0 Then
    MsgBox "Informe o Código do CVT.", vbInformation, App.Title
    CamposOK = False
    txtCVT.SetFocus
    Exit Function
  End If

  'Valida CVT
  If txtCVT.Text <> "77330" And txtCVT.Text <> "77437" And txtCVT.Text <> "77445" And txtCVT.Text <> "77453" Then
    MsgBox "Código de CVT Inválido.", vbInformation, App.Title
    CamposOK = False
    txtCVT.SetFocus
    Exit Function
  End If

  'Agencia
  If Len(Trim(txtAgencia.Text)) = 0 Then
    MsgBox "Informe o Código da Agência.", vbInformation, App.Title
    CamposOK = False
    txtAgencia.SetFocus
    Exit Function
  End If

  'Cedente
  If Len(Trim(TxtCedente.Text)) = 0 Then
    MsgBox "Informe o Código do Cedente.", vbInformation, App.Title
    CamposOK = False
    TxtCedente.SetFocus
    Exit Function
  End If

  'Nosso Numero
  If Len(Trim(txtNossoNumero.Text)) = 0 Then
    MsgBox "Informe o Nosso Numero.", vbInformation, App.Title
    CamposOK = False
    txtNossoNumero.SetFocus
    Exit Function
  End If

  'Vencimento
  If Len(Trim(TxtVencimento.Text)) = 0 Then
    MsgBox "Informe a Data de Vencimento do Documento.", vbInformation, App.Title
    CamposOK = False
    TxtVencimento.SetFocus
    Exit Function
  End If

  'Valor Base
  If Len(Trim(TxtValorBase.Text)) = 0 Then
    MsgBox "Informe o Valor Base.", vbInformation, App.Title
    CamposOK = False
    TxtValorBase.SetFocus
    Exit Function
  End If

  Call CalculaValorCobrado

  'Valor
  If Val(TxtValor.Text) <= 0 Then
    MsgBox "Valor Cobrado Inválido.", vbInformation, App.Title
    CamposOK = False
    TxtValorBase.SetFocus
    Exit Function
  End If

  CamposOK = True
End Function
Public Sub SetPosition(iLeft As Integer, iTop As Integer)

  Me.Left = iLeft
  Me.Top = iTop
End Sub
Public Sub SetParent(ByRef aForm As Form)

  Set mForm = aForm
End Sub
Sub PesquisaCobrancaEspecial()

  On Error GoTo ERRO_PESQUISACOBRANCAESPECIAL

  Dim sSql As String
  Dim RsCobrancaEsp As rdoResultset

  'Pesquisar o Documento Atual e preencher os valores caso encontre
  sSql = Geral.DataProcessamento & " , " & Geral.Documento.IdDocto

  Set qryGetCobrancaEspecial = Geral.Banco.CreateQuery("", "{call GetCobrancaEspecial (" & sSql & ")}")

  Set RsCobrancaEsp = qryGetCobrancaEspecial.OpenResultset(rdOpenStatic, rdConcurReadOnly)
  If Not RsCobrancaEsp.EOF Then
    'Encontrou o Documento -> Preencher os campos
    txtCVT.Text = RsCobrancaEsp!CVT
    TxtVencimento.Text = Format(DataDDMMAAAA(RsCobrancaEsp!vecto), "00000000")
    txtAgencia.Text = RsCobrancaEsp!Agencia
    TxtCedente.Text = RsCobrancaEsp!Cedente
    txtNossoNumero.Text = RsCobrancaEsp!NossoNumero
    TxtValorBase.Text = RsCobrancaEsp!ValorBase * 100
    TxtJuros.Text = RsCobrancaEsp!Juros * 100
    txtDesconto.Text = RsCobrancaEsp!Desconto * 100
    TxtAbatimento.Text = RsCobrancaEsp!Abatimento * 100
    TxtValor.Text = RsCobrancaEsp!Valor * 100

    TxtValorBase.SetFocus
  Else
    txtCVT.SetFocus
  End If

  If AlteraValor = True Then
    txtCVT.Locked = True
    TxtVencimento.Locked = True
    txtAgencia.Locked = True
    TxtCedente.Locked = True
    txtNossoNumero.Locked = True

    TxtValorBase.SetFocus
  End If

  Screen.MousePointer = vbDefault

  Exit Sub

ERRO_PESQUISACOBRANCAESPECIAL:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Selecionar Dados da Cobrança Especial.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Sub
Sub AjustesIniciais()

  On Error GoTo ERRO_AJUSTESINICIAIS

  'Setar os Objetos RDOQuery
'  Set qryAtualizaDocumentoExcluido = Geral.Banco.CreateQuery("", "{call AtualizaDocumentoExcluido (?,?,?,?,?)}")
  Set qryAtualizaCobrancaEspecial = Geral.Banco.CreateQuery("", "{call AtualizaCobrancaEspecial (?,?,?,?,?,?,?,?,?,?,?,?,?,?)}")
  Set qryRemoveTipoDocumento = Geral.Banco.CreateQuery("", "{? = call RemoveTipoDocumento (?,?,?)}")
  Set qryGetCobrancaEspecial = Geral.Banco.CreateQuery("", "{call GetCobrancaEspecial (?,?)}")

  Exit Sub

ERRO_AJUSTESINICIAIS:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Atualizar Dados da Cobrança Especial.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
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

Private Sub TxtAbatimento_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  End If
End Sub
Private Sub TxtAbatimento_LostFocus()

  'Calcula o Valor Cobrado
  Call CalculaValorCobrado
End Sub


Private Sub txtAgencia_Change()

  If Len(Trim(txtAgencia.Text)) = txtAgencia.MaxLength Then
    SendKeys "{TAB}"
    DoEvents
  End If
End Sub

Private Sub txtAgencia_GotFocus()

    SelecionarTexto txtAgencia

End Sub

Private Sub txtAgencia_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub

Private Sub TxtCedente_Change()

  If Len(Trim(TxtCedente.Text)) = TxtCedente.MaxLength Then
    SendKeys "{TAB}"
    DoEvents
  End If
End Sub

Private Sub TxtCedente_GotFocus()

    SelecionarTexto TxtCedente

End Sub

Private Sub TxtCedente_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub

Private Sub TxtCedente_Validate(Cancel As Boolean)

   If Len(Trim(TxtCedente.Text)) > 0 And Val(TxtCedente.Text) <> 0 Then
      'Validar Cedente
      If Not VerificarCodigoCedente Then
         Cancel = True
      End If
   End If
End Sub

Private Sub txtCVT_Change()

  If Len(Trim(txtCVT.Text)) = txtCVT.MaxLength Then
    SendKeys "{TAB}"
    DoEvents
  End If
End Sub

Private Sub txtCVT_GotFocus()

    SelecionarTexto txtCVT
    
End Sub


Private Sub txtCVT_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub

Private Sub txtDesconto_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtDesconto_LostFocus()

  'Calcula o Valor Cobrado
  Call CalculaValorCobrado
End Sub


Private Sub TxtJuros_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtJuros_LostFocus()

  'Calcula o Valor Cobrado
  Call CalculaValorCobrado
End Sub
Private Sub txtNossoNumero_Change()

  If Len(Trim(txtNossoNumero.Text)) = txtNossoNumero.MaxLength Then
    SendKeys "{TAB}", True
    DoEvents
  End If
End Sub

Private Sub txtNossoNumero_GotFocus()

    SelecionarTexto txtNossoNumero

End Sub

Private Sub txtNossoNumero_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub

Private Sub txtNossoNumero_Validate(Cancel As Boolean)

   If Len(Trim(txtNossoNumero.Text)) > 0 Then
      'Formatar o Campo 'Nosso Numero'
      txtNossoNumero.Text = Format(txtNossoNumero.Text, String(15, "0"))

      'Validar Campo NossoNumero
      If ValidaNossoNumero_Esp <> 1 Then
         MsgBox "O Campo 'Nosso Número' não é válido.", vbInformation, App.Title
         Cancel = True
      End If
   End If
End Sub
Private Sub txtValor_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    Call cmdConfirmar_Click
  End If
End Sub
Private Sub TxtValorBase_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub TxtValorBase_LostFocus()

  'Calcula o Valor Cobrado
  Call CalculaValorCobrado
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
