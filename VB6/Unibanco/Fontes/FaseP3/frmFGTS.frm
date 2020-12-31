VERSION 5.00
Object = "{9CBA5D64-E3C8-11D3-9FFC-00104BC8688C}#1.0#0"; "CurrencyEdit.ocx"
Object = "{4A8E27C4-EACA-11D3-9FFC-00104BC8688C}#1.0#0"; "DateEdit.ocx"
Begin VB.Form frmFGTS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Complementação de FGTS"
   ClientHeight    =   3336
   ClientLeft      =   1152
   ClientTop       =   2016
   ClientWidth     =   9336
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3336
   ScaleWidth      =   9336
   Begin DATEEDITLib.DateEdit dtData_Validade 
      Height          =   372
      Left            =   5352
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2076
      Width           =   1524
      _Version        =   65537
      _ExtentX        =   2688
      _ExtentY        =   656
      _StockProps     =   93
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      Locked          =   -1  'True
   End
   Begin VB.TextBox txtCodigo2 
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
      Left            =   1836
      MaxLength       =   12
      TabIndex        =   1
      Top             =   1248
      Width           =   1644
   End
   Begin VB.TextBox txtCodigo3 
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
      Left            =   3528
      MaxLength       =   12
      TabIndex        =   2
      Top             =   1248
      Width           =   1644
   End
   Begin VB.TextBox txtCodigo1 
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
      MaxLength       =   12
      TabIndex        =   0
      Top             =   1248
      Width           =   1644
   End
   Begin VB.TextBox txtCodigo4 
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
      Left            =   5220
      MaxLength       =   12
      TabIndex        =   3
      Top             =   1248
      Width           =   1644
   End
   Begin VB.TextBox txtCompetencia 
      BackColor       =   &H8000000F&
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
      Left            =   3864
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2076
      Width           =   1212
   End
   Begin VB.TextBox txtCGC_Tomador 
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
      Left            =   7056
      MaxLength       =   14
      TabIndex        =   8
      Top             =   2064
      Width           =   1968
   End
   Begin VB.TextBox txtCodigo_Recolhimento 
      BackColor       =   &H8000000F&
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
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2076
      Width           =   732
   End
   Begin VB.TextBox TxtCGC_Empresa 
      BackColor       =   &H8000000F&
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
      Left            =   1464
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2076
      Width           =   1776
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   696
      Left            =   8244
      Picture         =   "frmFGTS.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   144
      Width           =   816
   End
   Begin VB.CommandButton cmdFrenteVerso 
      Caption         =   "Fren/Ver"
      Height          =   696
      Left            =   6600
      Picture         =   "frmFGTS.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   144
      Width           =   816
   End
   Begin VB.CommandButton cmdInverteCor 
      Caption         =   "Inverter"
      Height          =   696
      Left            =   5772
      Picture         =   "frmFGTS.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   144
      Width           =   816
   End
   Begin VB.CommandButton cmdRotacao 
      Caption         =   "Rotação"
      Height          =   696
      Left            =   4944
      Picture         =   "frmFGTS.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   144
      Width           =   816
   End
   Begin VB.CommandButton cmdZoomMenos 
      Caption         =   "Zoom -"
      Height          =   696
      Left            =   4116
      Picture         =   "frmFGTS.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   144
      Width           =   816
   End
   Begin VB.CommandButton cmdZoomMais 
      Caption         =   "Zoom +"
      Height          =   696
      Left            =   3288
      Picture         =   "frmFGTS.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   144
      Width           =   816
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "&Confirmar"
      Height          =   696
      Left            =   7416
      Picture         =   "frmFGTS.frx":123C
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   144
      Width           =   816
   End
   Begin CURRENCYEDITLib.CurrencyEdit txtValorTotal 
      Height          =   372
      Left            =   7056
      TabIndex        =   12
      Top             =   2844
      Width           =   1932
      _Version        =   65537
      _ExtentX        =   3408
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
   Begin CURRENCYEDITLib.CurrencyEdit txtValorMulta 
      Height          =   372
      Left            =   3840
      TabIndex        =   11
      Top             =   2844
      Width           =   1740
      _Version        =   65537
      _ExtentX        =   3069
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
   Begin CURRENCYEDITLib.CurrencyEdit txtValorJAM 
      Height          =   372
      Left            =   1992
      TabIndex        =   10
      Top             =   2844
      Width           =   1740
      _Version        =   65537
      _ExtentX        =   3069
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
   Begin CURRENCYEDITLib.CurrencyEdit txtValorDeposito 
      Height          =   372
      Left            =   144
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2844
      Width           =   1740
      _Version        =   65537
      _ExtentX        =   3069
      _ExtentY        =   656
      _StockProps     =   93
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      Locked          =   -1  'True
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "(+) Multa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   3840
      TabIndex        =   30
      Top             =   2592
      Width           =   828
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "(+) JAM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   1992
      TabIndex        =   29
      Top             =   2592
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "Deposito"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   144
      TabIndex        =   28
      Top             =   2592
      Width           =   840
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "CNPJ/CEI Tomador"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   7056
      TabIndex        =   27
      Top             =   1824
      Width           =   1812
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Data de Validade"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   5328
      TabIndex        =   26
      Top             =   1824
      Width           =   1572
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "Cod. Recol."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   144
      TabIndex        =   25
      Top             =   1824
      Width           =   1092
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "Competência"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   3840
      TabIndex        =   24
      Top             =   1824
      Width           =   1224
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "(=) Total a recolher"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   7056
      TabIndex        =   23
      Top             =   2592
      Width           =   1800
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "CNPJ/CEI Empresa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   1464
      TabIndex        =   22
      Top             =   1824
      Width           =   1800
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "Linha Digitável do Código de Barras"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   144
      TabIndex        =   21
      Top             =   948
      Width           =   3396
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "FGTS"
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
      Left            =   720
      TabIndex        =   20
      Top             =   372
      Width           =   528
   End
   Begin VB.Image Image1 
      Height          =   384
      Left            =   216
      Picture         =   "frmFGTS.frx":1546
      Top             =   252
      Width           =   384
   End
End
Attribute VB_Name = "frmFGTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''
'Declaração de variáveis de trabalho
''''''''''''''''''''''''''''''''''''
Private mForm       As Form
Public Alterou      As Boolean
Public AlteraValor  As Boolean
Private m_bIsEvent  As Boolean

''''''''''''''''''''''''''''''''
'Declaração de Variáveis do RDO
''''''''''''''''''''''''''''''''
Private qryRemoveTipoDocumento  As rdoQuery
Private qryAtualizaFGTS         As rdoQuery
Private qryGetFGTS              As rdoQuery
Private bActivate               As Boolean
Private m_PrimeiraVez           As Boolean
Public Function CalculaCGC(ByVal CGC As String) As String
'* Calcule de DV *'

   Dim soma As Integer, resto As Integer
   Dim digito_11 As Integer, p As Integer, peso As Integer
   Dim digito_rv As String

   soma = 0
   resto = 0
   digito_11 = 0        'calculado pelo módulo 11
   digito_rv = ""       'caracter digitado pelo operador

   '*************************************************************
   'número do CGC: (13+2)                0 9 9.9 9 9.9 9 9/9 9 9 9 - D D
   '                                     x x x x x x x x x x x x x
   'multiplica da direita para esquerda: 6 5 4 3 2 9 8 7 6 5 4 3 2
   '*************************************************************
   
   peso = 2    'começa multiplicar da direita para esquerda
   p = 13      'tamanho do campo se o digito
   
   Do
      '*** Peso de 2 a 9 (multiplicação dos caracteres de 2 a 9) ***
      soma = soma + Mid(CGC, p, 1) * peso
      
      p = p - 1            'ponteiro
      peso = peso + 1      'peso
      If (peso = 10) Then
         peso = 2
      End If
      If (p = 0) Then
         Exit Do
      End If
   Loop
   
   resto = soma Mod 11        'resto da divisão
   digito_11 = 11 - resto     'digito verificador
   
   '*** se o calculo for igual a 10 ou 11, muda-se para 0 ***
   If (digito_11 = 11) Or (digito_11 = 10) Then
      digito_11 = 0
   End If

   Mid(CGC, Len(CGC) - 1, 1) = digito_11
   
   soma = 0
   resto = 0
   digito_11 = 0        'calculado pelo módulo 11
   digito_rv = ""       'caracter digitado pelo operador
   peso = 2             'começa multiplicar da direita para esquerda
   p = 14               'tamanho do campo (13) + 1º digito = 14
   
   Do
      '*** Peso de 2 a 9 (multiplicação dos caracteres de 2 a 9) ***
      soma = soma + Mid(CGC, p, 1) * peso
      p = p - 1            'ponteiro
      peso = peso + 1      'peso
      If (peso = 10) Then
         peso = 2
      End If
      If (p = 0) Then
         Exit Do
      End If
   Loop
   
   resto = soma Mod 11        'resto da divisão
   digito_11 = 11 - resto     'digito verificador
   
   '*** se o calculo for igual a 10 ou 11, muda-se para 0 ***
   If (digito_11 = 11) Or (digito_11 = 10) Then
      digito_11 = 0
   End If

   Mid(CGC, Len(CGC), 1) = digito_11

   CalculaCGC = CGC
    
End Function

Private Function GravaDocumento(ByVal psCodigoBarras As String) As Boolean

On Error GoTo ERRO_GravarDocumento:
    
    Dim strEncripta   As String
    
    GravaDocumento = False
    
    Screen.MousePointer = vbHourglass
    
    'Atualiza campo Autenticação Digital
    strEncripta = G_EncriptaBO(40, TxtCGC_Empresa.Text)
    If strEncripta = "" Then GoTo ERRO_GravarDocumento
    
    With qryAtualizaFGTS
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1) = Geral.DataProcessamento                 'Data Proc.
        .rdoParameters(2) = Geral.Documento.IdDocto                 'IdDocto
        .rdoParameters(3) = psCodigoBarras                          'Codigo de Barras
        .rdoParameters(4) = txtCodigo_Recolhimento                  'Codigo de recolhimento
        .rdoParameters(5) = TxtCGC_Empresa                          'CGC da Empresa
        .rdoParameters(6) = Format(txtCompetencia.Text, "yyyymm")   'Competencia
        .rdoParameters(7) = Format(Format(dtData_Validade.Text, "00/00/0000"), "yyyymmdd") 'Data de Validade
        .rdoParameters(8) = txtCGC_Tomador                          'CGC do Tomador
        .rdoParameters(9) = Val(txtValorDeposito.Text) / 100        'Valor JAM
        .rdoParameters(10) = Val(txtValorJAM.Text) / 100            'Valor JAM
        .rdoParameters(11) = Val(txtValorMulta.Text) / 100          'Valor da Multa
        .rdoParameters(12) = Val(txtValorTotal.Text) / 100          'Valor total
        .rdoParameters(13) = 40                                     'Tipo de documento
        .rdoParameters(14) = strEncripta                            'Autenticacao digital
        .Execute
    End With
    
    If qryAtualizaFGTS.rdoParameters(0) = 1 Then
        GoTo ERRO_GravarDocumento:
    End If
    If qryAtualizaFGTS.rdoParameters(0) = 2 Then
        'Documento Duplicado
        Geral.Documento.Status = "D"
        Geral.Capa.Duplicidade = "1"
    End If
    
    '''''''''''''''''''''''''''''
    'Atualizar o Controle Global
    '''''''''''''''''''''''''''''
    Geral.Documento.ValorTotal = Val(txtValorTotal.Text) / 100
    Geral.Documento.Leitura = psCodigoBarras
    Geral.Documento.TipoDocto = 40
    
    
    Screen.MousePointer = vbDefault
    GravaDocumento = True
    Exit Function
    
ERRO_GravarDocumento:
    Screen.MousePointer = vbDefault
    Call TratamentoErro("Erro ao Atualizar o documento.", Err, rdoErrors)

End Function
Private Function VerificaCodigo1() As Boolean

    On Error GoTo Err_VerificaCodigo1

    VerificaCodigo1 = False
    
    If Len(Trim(txtCodigo1.Text)) <> 12 Then Exit Function
    
    'Verifica se é mesmo um FGTS
    If (Mid(txtCodigo1, 1, 1) = "8" And Mid(txtCodigo1, 2, 1) = "5" And Mid(txtCodigo1, 3, 1) = "6") Or _
       (Mid(txtCodigo1, 1, 1) = "8" And Mid(txtCodigo1, 2, 1) = "5" And Mid(txtCodigo1, 3, 1) = "7") Then
        If Not Modulo10(txtCodigo1.Text, 12) Then Exit Function
    Else
        Exit Function
    End If
    
    VerificaCodigo1 = True
    
    Exit Function
    
Err_VerificaCodigo1:
    
    Call TratamentoErro("Erro ao verificar bloco 1.", Err, rdoErrors)
    
End Function

Private Function VerificaCodigo2() As Boolean

    On Error GoTo Err_VerificaCodigo2

    VerificaCodigo2 = False
    
    If Len(Trim(txtCodigo2.Text)) <> 12 Then Exit Function
    
    'Calculo do Modulo 10 para o campo 2
    If Not Modulo10(txtCodigo2.Text, 12) Then Exit Function
    
    VerificaCodigo2 = True
    
    Exit Function
    
Err_VerificaCodigo2:
    
    Call TratamentoErro("Erro ao verificar o bloco 2.", Err, rdoErrors)
    
End Function
Private Function VerificaCodigo3() As Boolean

    On Error GoTo Err_VerificaCodigo3

    VerificaCodigo3 = False
    
    If Len(Trim(txtCodigo3.Text)) <> 12 Then Exit Function
    
    'Calculo do Modulo 10 para o campo 3
    If Not Modulo10(txtCodigo3.Text, 12) Then Exit Function
    
    VerificaCodigo3 = True
    
    Exit Function
    
Err_VerificaCodigo3:

    Call TratamentoErro("Erro ao verificar bloco 3.", Err, rdoErrors)
    
End Function
Private Function VerificaCodigo4() As Boolean

    Dim sArrec As String
    
    On Error GoTo Err_VerificaCodigo4
    
    VerificaCodigo4 = False
    
    If Len(Trim(txtCodigo4.Text)) <> 12 Then Exit Function
    
    'Calculo do Modulo 10 para o campo 4
    If Not Modulo10(txtCodigo4.Text, 12) Then Exit Function
    
    'Só calcula este digito se não for o CDAE-RJ
    If ((Mid(txtCodigo1.Text, 1, 1) = "8") And (Mid(txtCodigo1.Text, 3, 1) = "6")) Then
    'Verifica se codigo de barras está batido atraves do 4º caracter
        sArrec = Mid(txtCodigo1.Text, 4, 1) + Mid(txtCodigo1.Text, 1, 3) + _
        Mid(txtCodigo1.Text, 5, 7) + Mid(txtCodigo2.Text, 1, 11) + _
        Mid(txtCodigo3.Text, 1, 11) + Mid(txtCodigo4.Text, 1, 11)
        'duvida
        If Not Modulo10Arrecadacao(sArrec, 44) Then Exit Function
     End If
    
    VerificaCodigo4 = True
    
    Exit Function
    
Err_VerificaCodigo4:
    
    Call TratamentoErro("Erro ao verificar o bloco 4.", Err, rdoErrors)

End Function
Function ValidaCodigoBarras() As Boolean
    On Error GoTo VALIDACODIGOBARRAS_ERRO


    Dim sCodigoBarras As String

    ValidaCodigoBarras = False

    'Verificar se o código de barras é válido
    If Not VerificaCodigo1 And Not bActivate Then
        MsgBox "Código de Barras Inválido.", vbInformation + vbOKOnly, App.Title
        txtCodigo1.SetFocus
        Exit Function
    End If

    If Not VerificaCodigo2 And Not bActivate Then
        MsgBox "Código de Barras Inválido.", vbInformation + vbOKOnly, App.Title
        txtCodigo2.SetFocus
        Exit Function
    End If

    If Not VerificaCodigo3 And Not bActivate Then
        MsgBox "Código de Barras Inválido.", vbInformation + vbOKOnly, App.Title
        txtCodigo3.SetFocus
        Exit Function
    End If

    If Not VerificaCodigo4 And Not bActivate Then
        MsgBox "Código de Barras Inválido.", vbInformation + vbOKOnly, App.Title
        txtCodigo4.SetFocus
        Exit Function
    End If
    
    ValidaCodigoBarras = True

    Exit Function

VALIDACODIGOBARRAS_ERRO:
    Screen.MousePointer = vbDefault
    Call TratamentoErro("Erro ao Validar o Código de Barras.", Err, rdoErrors)

End Function
Function SalvaFGTS() As Boolean

    On Error GoTo ERRO_SALVAFGTS
    
    Dim sCodigoBarras   As String
    Dim sCGC            As String
    Dim sCEI            As String
    Dim TipoDocto       As Integer
    
    SalvaFGTS = False
    
    Call CalculaValorTotal

    'Verificar se todos os campos estão preenchidos
    If Not ValidaCodigoBarras Then Exit Function

    'Verificar se CNPJ da Empresa é valido para o codigo de barras informado
    If Mid(txtCodigo3.Text, 11, 1) & Left(txtCodigo4.Text, 11) <> Left(TxtCGC_Empresa.Text, 12) Then
        MsgBox "CGC Empresa informado não é compatível com o código de barras.", vbInformation + vbOKOnly, App.Title
        TxtCGC_Empresa.SetFocus
        Exit Function
    End If
    
    If Not verificaCompetencia Then
        MsgBox "Data inválida.", vbCritical
        Exit Function
    End If
    
    If Not CamposOK Then Exit Function
    
    If Valida_CodRecol(Mid(txtCodigo2.Text, 5, 4)) = False Then Exit Function
    
    'Verifica Tipo de Inscrição da Empresa (Exceção de 1 à 2)
    If InStr("0179*0180*0181*0182", Mid(txtCodigo2.Text, 5, 4)) <> 0 Then
        If Mid(txtCodigo3.Text, 10, 1) = "1" Or Mid(txtCodigo3.Text, 10, 1) = "2" Then
            MsgBox "Empresa com tipo de inscrição inválido."
            Exit Function
        End If
    Else
        'Verifica Tipo de Inscrição da Empresa (Permitido Tipo de 1 à 9 com exceção do 3)
        If Mid(txtCodigo3.Text, 10, 1) = "0" Or Mid(txtCodigo3.Text, 10, 1) = "3" Then
            MsgBox "Empresa com tipo de inscrição inválido."
            Exit Function
        End If
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Obtem formatacao e verificacao do CGC e CEI da empresa'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sCGC = Format(Trim(TxtCGC_Empresa.Text), String(15, "0"))
    sCEI = Format(Trim(TxtCGC_Empresa.Text), String(12, "0"))
    If (Trim(sCGC) = "") Or (Trim(sCEI) = "") Then
        MsgBox "CGC/CEI Empresa náo é válido.", vbExclamation
        Exit Function
    End If
    
    If Not VerificaCGC(sCGC) Then
        If Not VerificaCEI(sCEI) Then
            MsgBox "CGC/CEI náo é válido.", vbExclamation
            TxtCGC_Empresa.SetFocus
            Exit Function
        End If
    End If
    
    ''''''''''''''''''''''''''''''''''''''
    'Verificação do CNPJ e CEI do tomador'
    ''''''''''''''''''''''''''''''''''''''
    If Mid(txtCodigo2.Text, 6, 3) = 181 Or Mid(txtCodigo2.Text, 6, 3) = 182 Then
        If Not Valida_Identificador(Format(txtCGC_Tomador.Text, String(16, "0"))) Then
            MsgBox "O IDENTIFICADOR não é válido!", vbExclamation
            txtCGC_Tomador.SetFocus
            Exit Function
        End If
    Else
        If Len(Trim(txtCGC_Tomador.Text)) >= 13 Then
            If Not VerificaCGC(Format(txtCGC_Tomador.Text, String(15, "0"))) Then
                MsgBox "O CNPJ do Tomador não é válido!", vbExclamation
                txtCGC_Tomador.SetFocus
                Exit Function
            End If
        Else
            If Mid(txtCodigo2.Text, 6, 3) <> 107 And Mid(txtCodigo2.Text, 6, 3) <> 112 And Mid(txtCodigo2.Text, 6, 3) <> 179 And Mid(txtCodigo2.Text, 6, 3) <> 180 Then
                If Not VerificaCEI(txtCGC_Tomador.Text) Then
                    MsgBox "O CEI Tomador não é válido!", vbExclamation
                    txtCGC_Tomador.SetFocus
                    Exit Function
                End If
            End If
        End If
    End If
    '''''''''''''''''''''''''''''''''
    'Obtem codigo de barras sem o dv'
    '''''''''''''''''''''''''''''''''
    sCodigoBarras = Mid(txtCodigo1.Text, 1, 11) & Mid(txtCodigo2.Text, 1, 11) & _
                    Mid(txtCodigo3.Text, 1, 11) & Mid(txtCodigo4.Text, 1, 11)

    ''''''''''''''''''''''''''''''''''''''''''''''''
    'Verificar se o Documento pertence à outro Tipo
    ''''''''''''''''''''''''''''''''''''''''''''''''
    Call VerificarDocumento

    ''''''''''''''''''''''''''
    'Atualizar / Inserir FGTS
    ''''''''''''''''''''''''''
    If Not GravaDocumento(sCodigoBarras) Then Exit Function

    SalvaFGTS = True

    '''''''''''''''''''''''''''''
    'Atualizar o Controle Global
    '''''''''''''''''''''''''''''
    'Geral.Documento.ValorTotal = Val(txtValor.Text) / 100
    'Geral.Documento.Leitura = sCodigoBarras

    Exit Function

ERRO_SALVAFGTS:
    Screen.MousePointer = vbDefault
    Call TratamentoErro("Erro ao Atualizar Dados do FGTS.", Err, rdoErrors)
End Function
Function CamposOK() As Boolean

    On Error GoTo Err_CamposOk
    ''''''''''''''''''''''''''''''''''''''
    'Se entrar em algum if, retorna false
    ''''''''''''''''''''''''''''''''''''''
    CamposOK = False
    '''''''''''''''''''''''''''''''''''
    'Primeiro Campo do Código de Barras
    '''''''''''''''''''''''''''''''''''
    If Len(Trim(txtCodigo1.Text)) = 0 Or Trim(txtCodigo1.Text) = "0" Then
        MsgBox "Informe o Primeiro Campo do Código de Barras.", vbInformation, App.Title
        txtCodigo1.SetFocus
        Exit Function
    End If
    '''''''''''''''''''''''''''''''''''
    'Segundo Campo do Código de Barras
    '''''''''''''''''''''''''''''''''''
    If Len(Trim(txtCodigo2.Text)) = 0 Or Trim(txtCodigo2.Text) = "0" Then
        MsgBox "Informe o Segundo Campo do Código de Barras.", vbInformation, App.Title
        txtCodigo2.SetFocus
        Exit Function
    End If
    '''''''''''''''''''''''''''''''''''
    'Terceiro Campo do Código de Barras
    '''''''''''''''''''''''''''''''''''
    If Len(Trim(txtCodigo3.Text)) = 0 Or Trim(txtCodigo3.Text) = "0" Then
        MsgBox "Informe o Terceiro Campo do Código de Barras.", vbInformation, App.Title
        txtCodigo3.SetFocus
        Exit Function
    End If
    ''''''''''''''''''''''''''''''''''
    'Quarto Campo do Código de Barras
    ''''''''''''''''''''''''''''''''''
    If Len(Trim(txtCodigo4.Text)) = 0 Or Trim(txtCodigo4.Text) = "0" Then
        MsgBox "Informe o Quarto Campo do Código de Barras.", vbInformation, App.Title
        txtCodigo4.SetFocus
        Exit Function
    End If
    
    ''''''''''''''''''''''''
    'Codigo de recolhimento
    ''''''''''''''''''''''''
    If Len(Trim(txtCodigo_Recolhimento.Text)) = 0 Or Trim(txtCodigo_Recolhimento.Text) = "0" Then
        MsgBox "Informe o Código de Recolhimento.", vbInformation
        txtCodigo_Recolhimento.SetFocus
        Exit Function
    End If
    ''''''''''''''''''''''''''''''
    ' Identificacao do Empregador
    ''''''''''''''''''''''''''''''
    If InStr("0179*0180*0181*0182", Mid(txtCodigo2.Text, 5, 4)) <> 0 Then
        If InStr("0*3*4*5*6*7*8*9", Mid(txtCodigo3.Text, 10, 1)) = 0 Then
            MsgBox "Código de Barras incorreto.", vbInformation
            txtCodigo3.SetFocus
            Exit Function
        End If
    End If
    
    ''''''''''''''''''''''''''
    'Código do CGC da Empresa
    ''''''''''''''''''''''''''
    If Len(Trim(TxtCGC_Empresa.Text)) = 0 Then
        MsgBox "Informe o CGC da Empresa.", vbInformation
        TxtCGC_Empresa.SetFocus
        Exit Function
    End If
    '''''''''''''''''''
    'Campo competencia
    '''''''''''''''''''
    If Len(Trim(txtCompetencia.Text)) = 0 Or Not _
        IsDate(Format("01" & Format(txtCompetencia.Text, "mmyyyy"), "00/00/0000")) Then
        
        MsgBox "Informe a Competência.", vbInformation
        txtCompetencia.SetFocus
        Exit Function
    End If
    ''''''''''''''''''''''''
    'Campo Data de Validade'
    ''''''''''''''''''''''''
    If Len(Trim(dtData_Validade.Text)) = 0 Then
        MsgBox "Informe a Data de Validade.", vbInformation
        dtData_Validade.SetFocus
        Exit Function
    End If
    ''''''''''''''''''''''''''''''''''''
    'Verifica documento se está vencido'
    ''''''''''''''''''''''''''''''''''''
    If Mid(txtCodigo2.Text, 6, 3) <> "112" And Mid(txtCodigo2.Text, 6, 3) <> "108" Then
        If Val(Geral.DataProcessamento) > Val(dtData_Validade.InverseText) Then
            MsgBox "Este documento não será aceito por estar fora do prazo de validade.", vbExclamation
            Exit Function
        End If
    End If
    
    ''''''''''''''''''''''''
    'Campo Valor do depósito
    ''''''''''''''''''''''''
    If Len(Trim(txtValorDeposito.Text)) = 0 Or Val(txtValorDeposito.Text) = 0 Then
        MsgBox "Informe o Valor do Depósito.", vbInformation
        txtValorDeposito.SetFocus
        Exit Function
    End If
    
    CamposOK = True
    
    Exit Function
    
Err_CamposOk:
    Call TratamentoErro("Erro ao validar campos.", Err, rdoErrors)
End Function
Sub PesquisaArrecFGTS()

    On Error GoTo ERRO_PESQUISAARRECFGTS
    
    Dim sSql As String
    Dim RsArrec As rdoResultset
    
    txtCodigo1.Text = Mid(Geral.Documento.Leitura, 1, 11)
    txtCodigo1.Text = txtCodigo1.Text & DV10(txtCodigo1.Text)
    txtCodigo2.Text = Mid(Geral.Documento.Leitura, 12, 11)
    txtCodigo2.Text = txtCodigo2.Text & DV10(txtCodigo2.Text)
    txtCodigo3.Text = Mid(Geral.Documento.Leitura, 23, 11)
    txtCodigo3.Text = txtCodigo3.Text & DV10(txtCodigo3.Text)
    txtCodigo4.Text = Mid(Geral.Documento.Leitura, 34, 11)
    txtCodigo4.Text = txtCodigo4.Text & DV10(txtCodigo4.Text)

    If Len(Trim(Geral.Documento.Leitura)) = 44 Then
        'Verificar se o DV está correto
        If VerificaCodigo1 Then
            If VerificaCodigo2 Then
                If VerificaCodigo3 Then
                    If VerificaCodigo4 Then
                        txtCodigo4.SelStart = 0
                        txtCodigo4.SelLength = Len(txtCodigo4.Text)
                        txtCodigo4.SetFocus
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If

    TxtValor.Text = ""
    txtCodigo1.SetFocus
    
    Exit Sub

ERRO_PESQUISAARRECFGTS:
    Screen.MousePointer = vbDefault
    Call TratamentoErro("Erro ao Selecionar Dados do FGTS.", Err, rdoErrors)
End Sub
Private Function verificaCompetencia() As Boolean

    Dim sDia    As String
    Dim sMes    As String
    Dim sAno    As String
    
    On Error GoTo Err_VerificaCompetencia
    
    verificaCompetencia = False

    If Trim(txtCompetencia.Text) = "" Then verificaCompetencia = True: Exit Function
    
    If Len(txtCompetencia.Text) <> 7 Then Exit Function
    
    sDia = "01"
    sMes = Left(txtCompetencia, 2)
    sAno = Right(txtCompetencia, 4)
    
    If Val(sMes) < 1 Or Val(sMes) > 12 Then Exit Function
    If Val(sAno) < 1900 Or Val(sAno) > 2100 Then Exit Function
    
    If Not IsDate(Format(sDia & sMes & sAno, "00/00/0000")) Then Exit Function
    
    
    
    'isto chama o evento change
    m_bIsEvent = True
    txtCompetencia.Text = sMes & "/" & sAno
    m_bIsEvent = False
    
    verificaCompetencia = True
    
    Exit Function
    
Err_VerificaCompetencia:
    
    Call TratamentoErro("Erro ao verificar a compentência.", Err, rdoErrors)

End Function
Private Sub VerificarDocumento()

    On Error GoTo Err_VerificarDocumento

    If Val(Geral.Documento.TipoDocto) <> 40 And Val(Geral.Documento.TipoDocto) <> 0 Then
        With qryRemoveTipoDocumento
            .rdoParameters(1) = Geral.DataProcessamento     'Data Proc.
            .rdoParameters(2) = Geral.Documento.IdDocto     'IdDocto
            .rdoParameters(3) = Geral.Documento.TipoDocto   'Tipo do Documento
            .Execute
        End With
    End If
    
    Exit Sub
    
Err_VerificarDocumento:
    Call TratamentoErro("Erro ao verificar documento.", Err, rdoErrors)

End Sub
Private Sub cmdConfirmar_Click()

    If SalvaFGTS Then

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
    If m_PrimeiraVez Then Exit Sub
    bActivate = True

    Call AjustesIniciais

    Call PesquisaFGTS

    Call CalculaValorTotal

    bActivate = False
    
    m_PrimeiraVez = True
        
End Sub
'
'Calcula competencia do codigo de barras
'
Private Function CalculaCompetencia() As String

    Dim sCodigo1        As String
    Dim sCodigo2        As String
    
    On Error GoTo Err_CalculaCompetencia

    CalculaCompetencia = ""
    
    If VerificaCodigo3 Then
    
        '''''''''''''''''''''''''''''''''''
        'Obtenção do codigo de competencia'
        '''''''''''''''''''''''''''''''''''
        sCodigo1 = Mid(txtCodigo3.Text, 4, 3)
                                   'Fixo "Regra"
        sCodigo2 = DateDiff("m", "31/12/1966", Format(Date, "dd/mm/yyyy"))
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Na maioria das vezes sCodigo1 será maior que sCodigo1,'
        'Mas pode acontecer de estar se pagando antecipadamente'
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'If Val(sCodigo2) >= Val(sCodigo1) Then
            CalculaCompetencia = Format(DateSerial(Year(Date), Month(Date) - (Val(sCodigo2) - Val(sCodigo1)), 1), "mm/yyyy")
        'End If
    
    End If
    
    Exit Function
    
Err_CalculaCompetencia:

    Call TratamentoErro("Erro ao calcular competência.", Err, rdoErrors)
End Function

'
'Calcula o valor do depósito de acordo com o codigo de barras
'
Private Function CalculaDeposito() As String

    Dim svalor  As String
    
    On Error GoTo Err_CalculaDeposito
    
    CalculaDeposito = ""
    
    If VerificaCodigo1 And VerificaCodigo2 Then
        
        svalor = Mid(txtCodigo1.Text, 5, 7)
        
        svalor = svalor & Mid(txtCodigo2.Text, 1, 4)
        
        CalculaDeposito = svalor
        
    End If
    
    Exit Function
    
Err_CalculaDeposito:
    Call TratamentoErro("Erro ao calcular o depósito.", Err, rdoErrors)
    
End Function

'
'Calcula a data de validade de acordo com o codigo de barras
'
Private Function CalculaValidade() As String

    Dim sDia        As String
    Dim sMes        As String
    Dim sAno        As String
    Dim sData       As String
    
    On Error GoTo Err_CalculaValidade
    
    CalculaValidade = ""

    ''''''''''''''''''''''''''''''''''''''''''''''''
    'verifica se os dois codigos de barras estão ok'
    ''''''''''''''''''''''''''''''''''''''''''''''''
    If VerificaCodigo2 And VerificaCodigo3 Then
        sDia = Mid(txtCodigo3.Text, 2, 2)
        sMes = Mid(txtCodigo2.Text, 11, 1) & Mid(txtCodigo3.Text, 1, 1)
        sAno = IIf(Mid(txtCodigo2.Text, 9, 2) > 50, Val(Mid(txtCodigo2.Text, 9, 2)) + 1900, Val(Mid(txtCodigo2.Text, 9, 2)) + 2000)

        sData = Format(sDia & "/" & sMes & "/" & sAno, "dd/mm/yyyy")

        If IsDate(sData) Then CalculaValidade = sData
        
    
    End If
    
    Exit Function
    
Err_CalculaValidade:
    Call TratamentoErro("Erro ao calcular a validade.", Err, rdoErrors)


End Function

Sub PesquisaFGTS()

    On Error GoTo ERRO_PesquisaFGTS
    
    Dim sSql      As String
    Dim RsFGTS    As rdoResultset
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Habilita para digitação somente os campos de valores'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If AlteraValor = True Then
        txtCodigo1.Locked = True
        txtCodigo2.Locked = True
        txtCodigo3.Locked = True
        txtCodigo4.Locked = True
        txtCGC_Tomador.Locked = True
    End If
    
    '''''''''''''''''''''''''''''''''''''''
    'Preenche os campos de codigo de barra
    '''''''''''''''''''''''''''''''''''''''
    txtCodigo1.Text = Mid(Geral.Documento.Leitura, 1, 11)
    txtCodigo1.Text = IIf(Val(txtCodigo1.Text) <> 0, txtCodigo1.Text & DV10(txtCodigo1.Text), "")
    txtCodigo2.Text = Mid(Geral.Documento.Leitura, 12, 11)
    txtCodigo2.Text = IIf(Val(txtCodigo2.Text) <> 0, txtCodigo2.Text & DV10(txtCodigo2.Text), "")
    txtCodigo3.Text = Mid(Geral.Documento.Leitura, 23, 11)
    txtCodigo3.Text = IIf(Val(txtCodigo3.Text) <> 0, txtCodigo3.Text & DV10(txtCodigo3.Text), "")
    txtCodigo4.Text = Mid(Geral.Documento.Leitura, 34, 11)
    txtCodigo4.Text = IIf(Val(txtCodigo4.Text) <> 0, txtCodigo4.Text & DV10(txtCodigo4.Text), "")
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Habilita ou Desabilita os campos do tomador, Juros e Multa'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Trim(txtCodigo2.Text) <> "" Then
        If Len(txtCodigo2.Text) > 8 Then
            Verifica_CodRecol (Mid(txtCodigo2.Text, 6, 3))
        End If
    End If
    
lbl_Continua:
    Screen.MousePointer = vbHourglass

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Pesquisar o FGTS Atual e preencher os valores caso encontre
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sSql = Geral.DataProcessamento & " , " & Geral.Documento.IdDocto

    Set qryGetFGTS = Geral.Banco.CreateQuery("", "{call GetFGTS (" & sSql & ")}")

    Set RsFGTS = qryGetFGTS.OpenResultset(rdOpenStatic, rdConcurReadOnly)
    
    If Not RsFGTS.EOF Then
        ''''''''''''''''''''''''''''''''''''''''''''''
        'Encontrou o Documento -> Preencher os campos
        ''''''''''''''''''''''''''''''''''''''''''''''
        txtCodigo_Recolhimento.Text = RsFGTS!CodRecolhimento
        TxtCGC_Empresa.Text = RsFGTS!CNPJCEI_Empresa
        txtCompetencia.Text = Format(Format(RsFGTS!Competencia, "0000/00"), "mm/yyyy")
        dtData_Validade.Text = Format(Format(RsFGTS!Validade, "0000/00/00"), "dd/mm/yyyy")
        txtCGC_Tomador.Text = RsFGTS!CNPJCEI_Tomador
        txtValorDeposito.Text = RsFGTS!Deposito * 100
        txtValorJAM.Text = RsFGTS!JAM * 100
        txtValorMulta.Text = RsFGTS!Multa * 100
    End If
    
    DoEvents
    If Len(Geral.Documento.Leitura) = 44 Then
        txtCodigo4.SetFocus
        txtCodigo4_LostFocus
        SendKeys "{TAB}"
    Else
        txtCodigo1.SetFocus
    End If


    Screen.MousePointer = vbDefault

    Exit Sub

ERRO_PesquisaFGTS:
    Screen.MousePointer = vbDefault
    Call TratamentoErro("Erro ao Selecionar Dados do FGTS.", Err, rdoErrors)
End Sub
Sub AjustesIniciais()

    On Error GoTo Err_AjustesIniciais

    '''''''''''''''''''''''''''''''
    'Setando as variáveis RDOQUERY
    '''''''''''''''''''''''''''''''
    Set qryAtualizaFGTS = Geral.Banco.CreateQuery("", "{? = call AtualizaFGTS (?,?,?,?,?,?,?,?,?,?,?,?,?,?)}")
    Set qryRemoveTipoDocumento = Geral.Banco.CreateQuery("", "{? = call RemoveTipoDocumento (?,?,?)}")
    
    Exit Sub
Err_AjustesIniciais:

    Call TratamentoErro("Erro ao iniciar.", Err, rdoErrors)
       
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub
Public Sub SetPosition(iLeft As Integer, iTop As Integer)
    Me.Left = iLeft
    Me.Top = iTop
End Sub
Public Sub SetParent(ByRef aForm As Form)
    Set mForm = aForm
End Sub
Sub CalculaValorTotal()

    On Error GoTo ERRO_CALCULAVALORTOTAL
    
    Dim Valor As Currency
    
    '--------------------   Calcular o Valor Total do Documento --------------------
    
    'Verificar se foi informado o Valor do INSS
    If Val(txtValorDeposito.Text) = 0 Then
        Valor = 0
    Else
        Valor = txtValorDeposito.Text
    End If
    
    'Verificar se foi informado 'Valor Outras Entidades'
    If Val(txtValorJAM.Text) <> 0 Then
        Valor = Valor + CCur(txtValorJAM.Text)
    End If
    
    'Verificar se foi informado Juros
    If Val(txtValorMulta.Text) <> 0 Then
        Valor = Valor + CCur(txtValorMulta.Text)
    End If
    
    'Transportar o Valor Final para a tela
    txtValorTotal.Text = Valor
    
    Exit Sub
    
ERRO_CALCULAVALORTOTAL:
    Screen.MousePointer = vbDefault
    Call TratamentoErro("Erro ao Calcular Valor Total do Documento.", Err, rdoErrors)
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
      Call mForm.Form_KeyUp(KeyCode, Shift)
  End Select
End Sub
Private Sub Form_Load()
    
    cmdSair.CausesValidation = False
    
    m_PrimeiraVez = False
    
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Set qryAtualizaFGTS = Nothing
    Set qryRemoveTipoDocumento = Nothing
    Set qryGetFGTS = Nothing
    m_PrimeiraVez = False
    
End Sub
Private Sub TxtCGC_Empresa_GotFocus()
    SelecionarTexto TxtCGC_Empresa
End Sub
Private Sub TxtCGC_Empresa_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub TxtCGC_Empresa_LostFocus()

    On Error GoTo Err_EmpresaLostFocus

    If Len(Trim(TxtCGC_Empresa.Text)) = 0 Then Exit Sub
    
    If LCase(Screen.ActiveControl.Name) = "cmdsair" Then Exit Sub

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Obtem formatacao e verificacao do CGC e CEI da empresa'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sCGC = Format(Trim(TxtCGC_Empresa.Text), String(15, "0"))
    sCEI = Format(Trim(TxtCGC_Empresa.Text), String(12, "0"))
    If (Trim(sCGC) = "") Or (Trim(sCEI) = "") Then
        MsgBox "CGC/CEI Empresa náo é válido.", vbExclamation
        Exit Sub
    End If
    
    If Not VerificaCGC(sCGC) Then
        If Not VerificaCEI(sCEI) Then
            MsgBox "CGC/CEI náo é válido.", vbExclamation
'            TxtCGC_Empresa.SetFocus
'            Exit Sub
        End If
    End If
    
    Exit Sub
Err_EmpresaLostFocus:
    Call TratamentoErro("Erro interno, txtCGC_Empresa_LostFocus().", Err, rdoErrors)
End Sub
Private Sub txtCGC_Tomador_Change()

    On Error GoTo Err_Tomador
    
    If Len(Trim(txtCGC_Tomador.Text)) = txtCGC_Tomador.MaxLength Then
        If Not IsNumeric(txtCGC_Tomador.Text) Then
            txtCGC_Tomador.Text = ""
            Exit Sub
        End If
'xxx
'        SendKeys "{TAB}"
'        DoEvents
    ElseIf Not IsNumeric(txtCGC_Tomador.Text) Then
        txtCGC_Tomador.Text = ""
    End If
    
    Exit Sub
Err_Tomador:
    Call TratamentoErro("Erro interno, txtCGC_Tomador_Change().", Err, rdoErrors)
End Sub
Private Sub txtCGC_Tomador_GotFocus()
    SelecionarTexto txtCGC_Tomador
End Sub
Private Sub txtCGC_Tomador_KeyPress(KeyAscii As Integer)
    SoNumero KeyAscii
End Sub
Private Sub txtCGC_Tomador_LostFocus()

Dim sIdent As String

    On Error GoTo Err_Tomador

    If LCase(Me.ActiveControl.Name) = "cmdsair" Then Exit Sub
    
    If Len(Trim(txtCGC_Tomador)) = 0 Then Exit Sub
    If txtCodigo_Recolhimento = "107" Then Exit Sub
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Obtem formatacao e verificacao do CGC e CEI da empresa'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sCGC = Format(Trim(txtCGC_Tomador.Text), String(15, "0"))
    sCEI = Format(Trim(txtCGC_Tomador.Text), String(12, "0"))
    sIdent = Format(Trim(txtCGC_Tomador.Text), String(16, "0"))
    
    If (Trim(sCGC) = "") Or (Trim(sCEI) = "" Or sIdent = "") Then
        MsgBox "CGC/CEI Tomador não é válido.", vbExclamation
        Exit Sub
    End If
    
    If Mid(txtCodigo2.Text, 5, 4) = "0181" Or Mid(txtCodigo2.Text, 5, 4) = "0182" Then
        If Not Valida_Identificador(sIdent) Then
            MsgBox "O IDENTIFICADOR não é válido.", vbExclamation
            txtCGC_Tomador.SetFocus
            Exit Sub
        End If
    Else
        If Not VerificaCGC(sCGC) Then
            If Not VerificaCEI(sCEI) Then
               MsgBox "CGC/CEI Tomador não é válido.", vbExclamation
               txtCGC_Tomador.SetFocus
               Exit Sub
            End If
        End If
    End If
    
    Exit Sub
    
Err_Tomador:
    Call TratamentoErro("Erro interno, txtCGC_Tomador_LostFocus().", Err, rdoErrors)
    
End Sub

Private Sub txtCodigo_Recolhimento_Change()

    On Error GoTo Err_CodigoRecolhimento

    If Len(Trim(txtCodigo_Recolhimento.Text)) = txtCodigo_Recolhimento.MaxLength Then
        If Not IsNumeric(txtCodigo_Recolhimento.Text) Then
            txtCodigo_Recolhimento.Text = ""
            Exit Sub
        End If
        SendKeys "{TAB}"
        DoEvents
    ElseIf Not IsNumeric(txtCodigo_Recolhimento.Text) Then
        txtCodigo_Recolhimento.Text = ""
    End If

    Exit Sub
Err_CodigoRecolhimento:
    Call TratamentoErro("Erro interno, txtCodigo_Recolhimento_Change().", Err, rdoErrors)

End Sub
Private Sub txtCodigo_Recolhimento_GotFocus()
    SelecionarTexto txtCodigo_Recolhimento
End Sub
Private Sub txtCodigo_Recolhimento_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub txtCodigo1_Change()

    On Error GoTo Err_Codigo1

    If Len(Trim(txtCodigo1.Text)) = txtCodigo1.MaxLength Then
        If Not IsNumeric(txtCodigo1.Text) Then
            txtCodigo1.Text = ""
            Exit Sub
        End If
        SendKeys "{TAB}"
        DoEvents
    ElseIf Not IsNumeric(txtCodigo1.Text) Then
        txtCodigo1.Text = ""
    End If
    
    Exit Sub
    
Err_Codigo1:
    Call TratamentoErro("Erro ao Calcular Valor Total do Documento.", Err, rdoErrors)
End Sub
Private Sub txtCodigo1_GotFocus()
    SelecionarTexto txtCodigo1
End Sub
Private Sub txtCodigo1_KeyPress(KeyAscii As Integer)

    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
        Exit Sub
    End If
    
End Sub
Private Sub txtCodigo1_LostFocus()

    On Error GoTo Err_Codigo1
  
    '* Verifica se usuário quer sair do FGTS *'
    If Screen.ActiveControl.Name = "cmdSair" Then
        CmdSair_Click
        Exit Sub
    End If
    
    Exit Sub
Err_Codigo1:
    Call TratamentoErro("Erro interno, txtCodigo1_LostFocus().", Err, rdoErrors)
End Sub
Private Sub txtCodigo1_Validate(Cancel As Boolean)

    On Error GoTo Err_Codigo1
        
    'Verificar se o código de barras é válido
    If Not VerificaCodigo1 And Not bActivate Then
        MsgBox "Código de Barras Inválido.", vbInformation, App.Title
        txtCodigo1.Text = ""
        Cancel = True
        SelecionarTexto txtCodigo1
    End If
    
    Exit Sub
    
Err_Codigo1:
    Call TratamentoErro("Erro interno, txtCodigo1_Validate().", Err, rdoErrors)
   
End Sub

Private Sub txtCodigo2_Change()
    If Len(Trim(txtCodigo2.Text)) = txtCodigo2.MaxLength Then
        If Not IsNumeric(txtCodigo2.Text) Then
            txtCodigo2.Text = ""
            Exit Sub
        End If
        SendKeys "{TAB}"
        DoEvents
    ElseIf Not IsNumeric(txtCodigo2.Text) Then
        txtCodigo2.Text = ""
    End If
End Sub
Private Sub txtCodigo2_GotFocus()
    SelecionarTexto txtCodigo2
End Sub
Private Sub txtCodigo2_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub txtCodigo2_LostFocus()
    
    On Error GoTo Err_Codigo2
    
    '* Verifica se usuário quer sair do FGTS *'
    If Screen.ActiveControl.Name = "cmdSair" Then
        CmdSair_Click
        Exit Sub
    End If
    Exit Sub
    
Err_Codigo2:
    
    Call TratamentoErro("Erro interno, txtCodigo2_LostFocus().", Err, rdoErrors)

End Sub
Private Sub txtCodigo2_Validate(Cancel As Boolean)

    'Verificar se o código de barras é válido
    If Not VerificaCodigo2 And Not bActivate Then
        MsgBox "Código de Barras Inválido.", vbInformation, App.Title
        txtCodigo2.Text = ""
        Cancel = True
        SelecionarTexto txtCodigo2
    End If
        
End Sub

Private Sub txtCodigo3_Change()

    On Error GoTo Err_Codigo3
    
    If Len(Trim(txtCodigo3.Text)) = txtCodigo3.MaxLength Then
        If Not IsNumeric(txtCodigo3.Text) Then
            txtCodigo3.Text = ""
            Exit Sub
        End If
        SendKeys "{TAB}"
        DoEvents
    ElseIf Not IsNumeric(txtCodigo3.Text) Then
        txtCodigo3.Text = ""
    End If
    
    Exit Sub
    
Err_Codigo3:
    Call TratamentoErro("Erro interno, txtCodigo3_Change().", Err, rdoErrors)
    
End Sub
Private Sub txtCodigo3_GotFocus()
    SelecionarTexto txtCodigo3
End Sub
Private Sub txtCodigo3_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub txtCodigo3_LostFocus()
    
    On Error GoTo Err_Codigo3

    '* Verifica se usuário quer sair do FGTS *'
    If Screen.ActiveControl.Name = "cmdSair" Then
        CmdSair_Click
        Exit Sub
    End If
    
    Exit Sub
Err_Codigo3:
    
    Call TratamentoErro("Erro interno, txtCodigo3_LostFocus().", Err, rdoErrors)
    
End Sub
Private Sub txtCodigo3_Validate(Cancel As Boolean)


    'Verificar se o código de barras é válido
    If Not VerificaCodigo3 And Not bActivate Then
        MsgBox "Código de Barras Inválido.", vbInformation, App.Title
        txtCodigo3.Text = ""
        Cancel = True
        SelecionarTexto txtCodigo3
    End If
   
End Sub
Private Sub txtCodigo4_Change()

    On Error GoTo Err_Codigo4

    If Len(Trim(txtCodigo4.Text)) = txtCodigo4.MaxLength Then

        If Not IsNumeric(txtCodigo4.Text) Then
            txtCodigo4.Text = ""
            Exit Sub
        End If
        txtCodigo4_KeyPress vbKeyReturn
        DoEvents
        SendKeys "{TAB}"

    ElseIf Not IsNumeric(txtCodigo4.Text) Then

        txtCodigo4.Text = ""

    End If
    
    Exit Sub
Err_Codigo4:
    Call TratamentoErro("Erro interno, txtCodigo4_Change().", Err, rdoErrors)
End Sub
Private Sub txtCodigo4_GotFocus()

    If bActivate = False Then SelecionarTexto txtCodigo4
    
End Sub
Private Sub txtCodigo4_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If VerificaCodigo2 Then
            Verifica_CodRecol Mid(txtCodigo2.Text, 6, 3)
        End If
    End If
    
    SoNumero KeyAscii
    
'    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> vbKeyBack Then
'        KeyAscii = 0
'    End If

    If Len(Trim(txtCodigo4.Text)) = 0 Then
        TxtCGC_Empresa.Text = ""
        txtCompetencia.Text = ""
        dtData_Validade.Text = ""
    End If
    
End Sub
Private Sub txtCodigo4_LostFocus()

    Dim CGCEmpresa As String
    
    On Error GoTo Err_Codigo4

    '* Verifica se usuário quer sair do FGTS *'
    If Screen.ActiveControl.Name = "cmdSair" Then
        CmdSair_Click
        Exit Sub
    End If
    
    If VerificaTxtVazio Then Exit Sub
        
    If ValidaCodigoBarras = True Then
        txtCodigo_Recolhimento = Mid$(txtCodigo3.Text, 7, 3)

        TxtCGC_Empresa.Text = Mid$(txtCodigo3, 11, 1) & Mid$(txtCodigo4, 1, 11)
        
        If Mid(txtCodigo3.Text, 10, 1) = "1" Or Mid(txtCodigo3.Text, 10, 1) = "4" Or _
           Mid(txtCodigo3.Text, 10, 1) = "6" Or Mid(txtCodigo3.Text, 10, 1) = "8" Then

            TxtCGC_Empresa.Text = Mid$(txtCodigo3, 11, 1) & Mid$(txtCodigo4, 1, 11) & "00"
            '* Cria Formatação para FGTS *'
            CGCEmpresa = (Format(TxtCGC_Empresa.Text, String(15, "0")))
            CGCEmpresa = CalculaCGC(CGCEmpresa)
            TxtCGC_Empresa.Text = Mid(CGCEmpresa, 2, 14)

        End If

        dtData_Validade.Text = CalculaValidade()
        txtCompetencia.Text = CalculaCompetencia()
        txtValorDeposito.Text = CalculaDeposito()
        If Trim(txtCodigo2.Text) <> "" Then
            If Len(txtCodigo2.Text) > 8 Then
                Verifica_CodRecol (Mid(txtCodigo2.Text, 6, 3))
            End If
        End If

        If bActivate = False Then CalculaValorTotal
   End If

   Exit Sub
   
Err_Codigo4:
   
   Call TratamentoErro("Erro interno, txtCodigo4_LostFocus().", Err, rdoErrors)
End Sub
Private Sub txtCodigo4_Validate(Cancel As Boolean)

    If Len(Trim(txtCodigo4)) <> 0 Then
        'Verificar se o código de barras é válido
        If Not VerificaCodigo4 And Not bActivate Then
            MsgBox "Código de Barras Inválido.", vbInformation, App.Title
            txtCodigo4.Text = ""
            Cancel = True
            SelecionarTexto txtCodigo4
        End If
    End If

End Sub

Private Sub txtCompetencia_GotFocus()
    SelecionarTexto txtCompetencia
End Sub

Private Sub txtCompetencia_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo Err_Competencia

    If Len(Trim(txtCompetencia.Text)) = 2 And KeyCode <> vbKeyDivide And _
       KeyCode <> vbKeyBack And KeyCode <> vbKeyDelete Then
        txtCompetencia.Text = txtCompetencia.Text & "/"
        txtCompetencia.SelStart = Len(txtCompetencia.Text)
        KeyCode = 0
    End If
    
    Exit Sub
    
Err_Competencia:
    
    Call TratamentoErro("Erro interno, txtCompetencia_KeyDown().", Err, rdoErrors)
End Sub
Private Sub txtCompetencia_KeyPress(KeyAscii As Integer)

    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If

End Sub
Private Sub txtCompetencia_LostFocus()

    If verificaCompetencia Then Exit Sub
    
ERRO_VERIFICACAO:
    MsgBox "Data inválida.", vbCritical
    txtCompetencia.Text = ""
    txtCompetencia.SetFocus

End Sub

Private Sub txtValorDeposito_LostFocus()
    Call CalculaValorTotal
End Sub

Private Sub txtValorJAM_GotFocus()
    SelecionarTexto txtValorJAM
End Sub

Private Sub txtValorJAM_LostFocus()
    Call CalculaValorTotal
End Sub
Private Sub txtValorMulta_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        Call cmdConfirmar_Click
'    End If
End Sub
Private Sub txtValorMulta_LostFocus()
    Call CalculaValorTotal
End Sub
Private Sub Verifica_CodRecol(TipoRecolhimento As Long)

    On Error GoTo Err_CodRecol

    'Verifica tipos de recolhimentos e dados a serem incluidos *'

    txtCGC_Tomador.MaxLength = 14   'Padrão para CNPJ/CEI
    Label10.Caption = "CNPJ/CEI Tomador"
    
    Select Case TipoRecolhimento
        Case 107:
            txtCGC_Tomador.Locked = True
            txtCGC_Tomador.TabStop = False
            txtCGC_Tomador.Text = ""
            txtValorMulta.Locked = True
            txtValorMulta.TabStop = False
            txtValorMulta.Text = ""
            txtValorJAM.Locked = True
            txtValorJAM.TabStop = False
            txtValorJAM.Text = ""

            txtCGC_Tomador.BackColor = G_ColorGray '&H80000000
            txtValorMulta.BackColor = G_ColorGray '&H80000000
            txtValorJAM.BackColor = G_ColorGray '&H80000000

        Case 108
            txtCGC_Tomador.Locked = False
            txtCGC_Tomador.TabStop = True
            txtValorMulta.Locked = False
            txtValorMulta.TabStop = True
            txtValorJAM.Locked = False
            txtValorJAM.TabStop = True

            txtCGC_Tomador.BackColor = G_ColorBackGround
            txtValorMulta.BackColor = G_ColorBackGround
            txtValorJAM.BackColor = G_ColorBackGround

        Case 111
            txtCGC_Tomador.Locked = False
            txtCGC_Tomador.TabStop = True
            txtValorMulta.Locked = True
            txtValorMulta.TabStop = False
            txtValorMulta.Text = ""
            txtValorJAM.Locked = True
            txtValorJAM.TabStop = False
            txtValorJAM.Text = ""

            txtCGC_Tomador.BackColor = G_ColorBackGround
            txtValorMulta.BackColor = G_ColorGray '&H80000000
            txtValorJAM.BackColor = G_ColorGray '&H80000000

        Case 112
            txtCGC_Tomador.Locked = True
            txtCGC_Tomador.TabStop = False
            txtCGC_Tomador.Text = ""
            txtValorMulta.Locked = False
            txtValorMulta.TabStop = True
            txtValorJAM.Locked = False
            txtValorJAM.TabStop = True

            txtCGC_Tomador.BackColor = G_ColorGray '&H80000000
            txtValorMulta.BackColor = G_ColorBackGround
            txtValorJAM.BackColor = G_ColorBackGround
        
        Case 179, 180
            txtCGC_Tomador.Locked = True
            txtCGC_Tomador.TabStop = False
            txtCGC_Tomador.Text = ""
            txtValorMulta.Locked = True
            txtValorMulta.TabStop = False
            txtValorJAM.Locked = True
            txtValorJAM.TabStop = False

            txtCGC_Tomador.BackColor = G_ColorGray '&H80000000
            txtValorMulta.BackColor = G_ColorGray '&H80000000
            txtValorJAM.BackColor = G_ColorGray '&H80000000
        
'        Case 180
'            txtCGC_Tomador.Locked = False
'            txtCGC_Tomador.TabStop = True
'            txtValorMulta.Locked = True
'            txtValorMulta.TabStop = False
'            txtValorJAM.Locked = True
'            txtValorJAM.TabStop = False
'
'            txtCGC_Tomador.BackColor = G_ColorBackGround
'            txtValorMulta.BackColor = G_ColorGray '&H80000000
'            txtValorJAM.BackColor = G_ColorGray '&H80000000
        
        Case 181, 182
            txtCGC_Tomador.Locked = False
            txtCGC_Tomador.TabStop = True
            Label10.Caption = "Identificador"
            txtCGC_Tomador.MaxLength = 16       'Padrão para Identificador
            txtValorMulta.Locked = True
            txtValorMulta.TabStop = False
            txtValorJAM.Locked = True
            txtValorJAM.TabStop = False

            txtCGC_Tomador.BackColor = G_ColorBackGround
            txtValorMulta.BackColor = G_ColorGray '&H80000000
            txtValorJAM.BackColor = G_ColorGray '&H80000000

    End Select
    
    Exit Sub
    
Err_CodRecol:
    Call TratamentoErro("Erro ao verificar o código de recolhimento.", Err, rdoErrors)
End Sub
Function Valida_CodRecol(CodRecol As Long) As Boolean

    On Error GoTo Err_CodRecol

    Select Case CodRecol
    Case 107, 179, 180
        Valida_CodRecol = True

    Case 108
        'Para o tipo de recolhimento 108 os campos obrigatórios são:
        'CNPJ do Tomador e Multa
        If Len(Trim(txtCGC_Tomador.Text)) = 0 Then
            MsgBox "Informe o CNPJ do Tomador.", vbInformation, App.Title
            txtCGC_Tomador.SetFocus
            Valida_CodRecol = False
            Exit Function
        Else
            Valida_CodRecol = True
        End If

        If Len(Trim(txtValorMulta.Text)) = 0 Then
            MsgBox "Informe o valor de Multa.", vbInformation, App.Title
            txtValorMulta.SetFocus
            Valida_CodRecol = False
            Exit Function
        Else
            Valida_CodRecol = True
        End If

    Case 111
        'Para o tipo de Recolhimento 111 os campos obrigatórios são:
        'CNPJ do Tomador
        If Len(Trim(txtCGC_Tomador.Text)) = 0 Then
            MsgBox "Informe o CGC do Tomador.", vbInformation, App.Title
            txtCGC_Tomador.SetFocus
            Valida_CodRecol = False
            Exit Function
        Else
            Valida_CodRecol = True
        End If

    Case 112
        'Para o tipo de Recolhimento 112 os campos obrigatórios são:
        'Valor de Multa
        If Len(Trim(txtValorMulta.Text)) = 0 Then
            MsgBox "Informe o valor de Multa.", vbInformation, App.Title
            txtValorMulta.SetFocus
            Valida_CodRecol = False
            Exit Function
        Else
            Valida_CodRecol = True
        End If

    'Novos Códigos
    Case 181 To 182
        If Len(Trim(txtCGC_Tomador.Text)) = 0 Then
            MsgBox "Informe o IDENTIFICADOR.", vbInformation, App.Title
            txtCGC_Tomador.SetFocus
            Valida_CodRecol = False
            Exit Function
        Else
            Valida_CodRecol = True
        End If
            
    Case Else
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Mesmo se o codigo não for aceito, verificar o prazo de validade. Assim'
        'não força o usuário fazer deste documento uma Arrecadação Convencional'
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        bMostrarMsg = False
        
        If Trim(dtData_Validade.Text) <> "" Then
            If Val(Format(Date, "yyyymmdd")) > Val(dtData_Validade.InverseText) Then
                MsgBox "Este documento não será aceito por estar fora do prazo de validade.", vbExclamation
            Else
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'Caso não esteja fora do prazo, simplesmente não      '
                'aceitar como FGTS e sim como Arrecadação Convencional'
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''
                bMostrarMsg = True
            End If
        End If
        
        If bMostrarMsg Then
            MsgBox "FGTS com este Código de Barras deve ser tratado como Arrecadação Convencional.", vbInformation + vbOKOnly, App.Title
            txtCodigo1.SetFocus
        End If
        
        Valida_CodRecol = False
        Exit Function
    End Select
    
    Exit Function

Err_CodRecol:
    Call TratamentoErro("Erro ao validar código de recolhimento.", Err, rdoErrors)
End Function

Private Sub txtValorTotal_GotFocus()
    SelecionarTexto txtValorTotal
End Sub

Private Sub txtValorTotal_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdConfirmar_Click
    End If
End Sub

Private Function Valida_Identificador(ByVal strIdent As String) As Boolean
   
   '--------- MODULO 11 (2 BASE 9) -----------------------------
   ' Esta rotina serve para conferir o Identificador: tam = 16
   '------------------------------------------------------------
   
   Dim soma As Integer, resto As Integer
   Dim digito_11 As Integer, p As Integer, peso As Integer
   Dim digito_rv As String
   Dim bOk As Boolean
   
   bOk = True           'default - OK
   
    '--- Verifica digitos do Identificador ---

    If strIdent = String(16, "0") Then
        bOk = False               'digito não confere
        Valida_Identificador = bOk
        Exit Function
    End If

   soma = 0
   resto = 0
   digito_11 = 0        'calculado pelo módulo 11
   digito_rv = ""       'caracter digitado pelo operador

   '**********************************************************************
   'número do Identificador: (14+2)      0 9 9 9 9 9 9 9 9 9 9 9 9 9 - D D
   '                                     x x x x x x x x x x x x x x
   'multiplica da direita para esquerda: 7 6 5 4 3 2 9 8 7 6 5 4 3 2
   '**********************************************************************
   
   peso = 2    'começa multiplicar da direita para esquerda
   p = 14      'tamanho do campo se o digito
   
   Do
      '*** Peso de 2 a 9 (multiplicação dos caracteres de 2 a 9) ***
      soma = soma + Mid(strIdent, p, 1) * peso
      p = p - 1            'ponteiro
      peso = peso + 1      'peso
      If (peso = 10) Then
         peso = 2
      End If
      If (p = 0) Then
         Exit Do
      End If
   Loop
   
   resto = soma Mod 11        'resto da divisão
   
   '--- Se o calculo for igual a (0) ou (1), muda-se para (0) senão subtrai-lo de (11) ---
    If (resto = 0) Or (resto = 1) Then
        digito_11 = 0
    Else
        digito_11 = (11 - resto)     'digito verificador
    End If

   digito_rv = Mid(strIdent, 15, 1)  '1º digito verificador
   
   If CStr(digito_11) <> (digito_rv) Then
      bOk = False                       'digito não confere
      Valida_Identificador = bOk
      Exit Function
   End If

   soma = 0
   resto = 0
   digito_11 = 0        'calculado pelo módulo 11
   digito_rv = ""       'caracter digitado pelo operador
   peso = 2             'começa multiplicar da direita para esquerda
   p = 15               'tamanho do campo (14) + 1º digito = 15
   
   '**********************************************************************
   'número do Identificador: (14+2)      0 9 9 9 9 9 9 9 9 9 9 9 9 9 - D D
   '                                     x x x x x x x x x x x x x x   x
   'multiplica da direita para esquerda: 8 7 6 5 4 3 2 9 8 7 6 5 4 3   2
   '**********************************************************************
   Do
      '*** Peso de 2 a 9 (multiplicação dos caracteres de 2 a 9) ***
      soma = soma + Mid(strIdent, p, 1) * peso
      p = p - 1            'ponteiro
      peso = peso + 1      'peso
      If (peso = 10) Then
         peso = 2
      End If
      If (p = 0) Then
         Exit Do
      End If
   Loop
   
   resto = soma Mod 11        'resto da divisão
   
   '--- Se o calculo for igual a (0) ou (1), muda-se para (0) senão subtrai-lo de (11) ---
    If (resto = 0) Or (resto = 1) Then
        digito_11 = 0
    Else
        digito_11 = (11 - resto)     'digito verificador
    End If
   

   digito_rv = Mid(strIdent, 16, 1)  '2º digito verificador
   
   If CStr(digito_11) <> (digito_rv) Then
      bOk = False               'digito não confere
   End If

   Valida_Identificador = bOk

End Function
