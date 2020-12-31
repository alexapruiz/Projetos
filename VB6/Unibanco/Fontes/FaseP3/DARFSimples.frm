VERSION 5.00
Object = "{9CBA5D64-E3C8-11D3-9FFC-00104BC8688C}#1.0#0"; "CurrencyEdit.ocx"
Object = "{4A8E27C4-EACA-11D3-9FFC-00104BC8688C}#1.0#0"; "DateEdit.ocx"
Begin VB.Form DARFSimples 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Complementação de DARF Símples"
   ClientHeight    =   2664
   ClientLeft      =   1188
   ClientTop       =   1320
   ClientWidth     =   7992
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2664
   ScaleWidth      =   7992
   Begin DATEEDITLib.DateEdit TxtPeriodo 
      Height          =   372
      Left            =   108
      TabIndex        =   0
      Top             =   1392
      Width           =   1296
      _Version        =   65537
      _ExtentX        =   2286
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
   Begin CURRENCYEDITLib.CurrencyEdit TxtReceita 
      Height          =   372
      Left            =   4476
      TabIndex        =   2
      Top             =   1392
      Width           =   1872
      _Version        =   65537
      _ExtentX        =   3302
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
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "&Confirmar"
      Height          =   696
      Left            =   6276
      Picture         =   "DARFSimples.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   132
      Width           =   816
   End
   Begin VB.CommandButton cmdZoomMais 
      Caption         =   "Zoom +"
      Height          =   696
      Left            =   2148
      Picture         =   "DARFSimples.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   132
      Width           =   816
   End
   Begin VB.CommandButton cmdZoomMenos 
      Caption         =   "Zoom -"
      Height          =   696
      Left            =   2976
      Picture         =   "DARFSimples.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   132
      Width           =   816
   End
   Begin VB.CommandButton cmdRotacao 
      Caption         =   "Rotação"
      Height          =   696
      Left            =   3804
      Picture         =   "DARFSimples.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   132
      Width           =   816
   End
   Begin VB.CommandButton cmdInverteCor 
      Caption         =   "Inverter"
      Height          =   696
      Left            =   4632
      Picture         =   "DARFSimples.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   132
      Width           =   816
   End
   Begin VB.CommandButton cmdFrenteVerso 
      Caption         =   "Fren/Ver"
      Height          =   696
      Left            =   5460
      Picture         =   "DARFSimples.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   132
      Width           =   816
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   696
      Left            =   7104
      Picture         =   "DARFSimples.frx":123C
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   132
      Width           =   816
   End
   Begin VB.TextBox TxtCGC 
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
      Left            =   1980
      MaxLength       =   15
      TabIndex        =   1
      Top             =   1392
      Width           =   1920
   End
   Begin CURRENCYEDITLib.CurrencyEdit TxtPrincipal 
      Height          =   372
      Left            =   72
      TabIndex        =   4
      Top             =   2184
      Width           =   1872
      _Version        =   65537
      _ExtentX        =   3302
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
   Begin CURRENCYEDITLib.CurrencyEdit TxtMulta 
      Height          =   372
      Left            =   2064
      TabIndex        =   5
      Top             =   2184
      Width           =   1872
      _Version        =   65537
      _ExtentX        =   3302
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
      Left            =   4056
      TabIndex        =   6
      Top             =   2184
      Width           =   1872
      _Version        =   65537
      _ExtentX        =   3302
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
   Begin CURRENCYEDITLib.CurrencyEdit TxtPercentual 
      Height          =   372
      Left            =   7020
      TabIndex        =   3
      Top             =   1392
      Width           =   888
      _Version        =   65537
      _ExtentX        =   1566
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
      MaxLength       =   4
      BackColor       =   -2147483643
   End
   Begin CURRENCYEDITLib.CurrencyEdit TxtValor 
      Height          =   372
      Left            =   6036
      TabIndex        =   7
      Top             =   2196
      Width           =   1872
      _Version        =   65537
      _ExtentX        =   3302
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
      Left            =   168
      Picture         =   "DARFSimples.frx":1546
      Top             =   252
      Width           =   384
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      Caption         =   "DARF Símples"
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
      Left            =   684
      TabIndex        =   23
      Top             =   396
      Width           =   1320
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Total"
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
      Left            =   6072
      TabIndex        =   22
      Top             =   1932
      Width           =   960
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "Juros "
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
      Left            =   4092
      TabIndex        =   21
      Top             =   1932
      Width           =   552
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "Multa"
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
      Left            =   2112
      TabIndex        =   20
      Top             =   1932
      Width           =   492
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "Principal"
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
      Left            =   132
      TabIndex        =   19
      Top             =   1932
      Width           =   780
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "CGC"
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
      Left            =   2016
      TabIndex        =   18
      Top             =   1140
      Width           =   444
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "Apuração"
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
      Left            =   120
      TabIndex        =   17
      Top             =   1128
      Width           =   852
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "Percentual"
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
      Left            =   6948
      TabIndex        =   16
      Top             =   1128
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "Receita Bruta"
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
      Left            =   4500
      TabIndex        =   15
      Top             =   1140
      Width           =   1224
   End
End
Attribute VB_Name = "DARFSimples"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Declaração de Variáveis do RDO
Private qryGetDARFSimples As rdoQuery
Private qryRemoveTipoDocumento As rdoQuery
Private qryAtualizaDARFSimples As rdoQuery

'Declaração de Variáveis de trabalho
Private mForm As Form
Public Alterou As Boolean
Public AlteraValor As Boolean

'Validação de informações do DARF SIMPLES
Private RegraValidaDARF() As tpValidacao
Private Type tpValidacao
    PercentualMinimo        As Currency
    PercentualMaximo        As Currency
    ValorMinimoDocumento    As Currency
End Type

Sub CalculaValorFinal()

  On Error GoTo ERRO_CALCULAVALORFINAL

  Dim Valor As Currency

  '----------------------------   Calcular o Valor Final do Documento --------------------------

  'Verificar se foi informado o Valor Principal
  If Val(TxtPrincipal.Text) = 0 Then
    Valor = 0
  Else
    Valor = TxtPrincipal.Text
  End If

  'Verificar se foi informado Multa
  If Val(TxtMulta.Text) <> 0 Then
    Valor = Valor + Val(TxtMulta.Text)
  End If

  'Verificar se foi informado Juros
  If Val(TxtJuros.Text) <> 0 Then
    Valor = Valor + Val(TxtJuros.Text)
  End If

  'Transportar o Valor Final para a tela
  TxtValor.Text = Valor

  Exit Sub

ERRO_CALCULAVALORFINAL:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Calcular Valor Final do Documento.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Sub
Sub CalculaValorPrincipal()

  On Error GoTo ERRO_CALCULAVALORPRINCIPAL

  Dim Valor As Currency

  '-------------------------- Calcular o Valor Principal do Documento --------------------------

  'Verificar se foi informado a Receita Bruta
  If Val(TxtReceita.Text) = 0 Then
    Exit Sub
  Else
    Valor = TxtReceita.Text / 100
  End If

  'Verificar se o valor da receita é maior que o permitido
  If Valor > 999999999.99 Then
    MsgBox "Valor de Receita não permitido.", vbInformation + vbOKOnly, App.Title
    Exit Sub
  End If

  'Verificar se foi informado o percentual
  If Val(TxtPercentual.Text) <> 0 Then
    'Valor = CLng(Format(Valor * (TxtPercentual.Text / 10000), ".00"))
    Valor = Format(Valor * (TxtPercentual.Text / 10000), ".00")
  Else
    Valor = 0
  End If

  Valor = Valor * 100

  'Transportar o Valor Final para a tela
  TxtPrincipal.Text = Valor

  Exit Sub

ERRO_CALCULAVALORPRINCIPAL:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Calcular Valor Final do Documento.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Sub
Private Sub cmdConfirmar_Click()
  If SalvaDARFSimples Then
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
  
    If Not CargaRegraValidaDARF Then
        Alterou = False
        Me.Hide
        Exit Sub
    End If
  
  Call AjustesIniciais
  
  Call PesquisaDARFSimples

End Sub
Public Sub SetPosition(iLeft As Integer, iTop As Integer)
  Me.Left = iLeft
  Me.Top = iTop
End Sub
Public Sub SetParent(ByRef aForm As Form)

  Set mForm = aForm
End Sub
Function SalvaDARFSimples() As Boolean

  On Error GoTo ERRO_SALVADARFSIMPLES

  Dim RetAgencia As Integer
  Dim strEncripta   As String
  
  SalvaDARFSimples = False

  'Call CalculaValorPrincipal
  Call CalculaValorFinal

  'Verificar se todos os campos estão preenchidos
  If CamposOK Then
    'Formatando o Campo CGC
    TxtCGC.Text = Format(TxtCGC.Text, String(15, "0"))

    'Validação de CGC
    If Mid(Format(TxtCGC.Text, String(15, "0")), 10, 4) <> "0001" Then
        MsgBox "Só são aceitos CGC's da Matriz.", vbInformation, App.Title
        TxtCGC.SetFocus
        Exit Function
    Else
      If Not VerificaCGC(TxtCGC.Text) Then
        MsgBox "CGC Inválido.", vbInformation, App.Title
        TxtCGC.SetFocus
        Exit Function
      End If
    End If

    'Verificar se a Agencia de Origem está OK
    If Not ValidaAgenciaPorDocto(Geral.Documento.Agencia, "", False) Then
        Exit Function
    End If

    'Verificar se o Documento pertence à outro Tipo
    If Geral.Documento.TipoDocto <> 17 And Geral.Documento.TipoDocto <> 0 Then
      With qryRemoveTipoDocumento
        .rdoParameters(1) = Geral.DataProcessamento     'Data Proc.
        .rdoParameters(2) = Geral.Documento.IdDocto     'IdDocto
        .rdoParameters(3) = Geral.Documento.TipoDocto   'Tipo do Documento
        .Execute
      End With
    End If

    'Atualiza campo Autenticação Digital
    strEncripta = G_EncriptaBO(17, CStr(Val(TxtCGC.Text)))
    If strEncripta = "" Then GoTo ERRO_SALVADARFSIMPLES

    'Atualizar / Inserir DARF Preto
    With qryAtualizaDARFSimples
      .rdoParameters(0) = Geral.DataProcessamento          'Data Proc.
      .rdoParameters(1) = Geral.Documento.IdDocto          'IdDocto
      .rdoParameters(2) = TxtPeriodo.InverseText           'PeriodoApuracao
      .rdoParameters(3) = TxtCGC.Text                      'CGC
      .rdoParameters(4) = "6106"                           'Receita
      .rdoParameters(5) = Val(TxtReceita.Text) / 100       'ReceitaBruta
      .rdoParameters(6) = Val(TxtPercentual.Text) / 100    'Percentual
      .rdoParameters(7) = Val(TxtPrincipal.Text) / 100     'ValorPrincipal
      .rdoParameters(8) = Val(TxtMulta.Text) / 100         'ValorMulta
      .rdoParameters(9) = Val(TxtJuros.Text) / 100         'Juros
      .rdoParameters(10) = Val(TxtValor.Text) / 100        'Valor
      .rdoParameters(11) = 17                              'TipoDocto
      .rdoParameters(12) = strEncripta                     'Autenticacao digital
      .Execute
    End With

    SalvaDARFSimples = True

    'Atualizar o Controle Global
    Geral.Documento.ValorTotal = Val(TxtValor.Text) / 100
    Geral.Documento.Leitura = ""
    Geral.Documento.TipoDocto = 17
  End If

  Exit Function

ERRO_SALVADARFSIMPLES:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Atualizar Dados do DARF Preto.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function
Sub PesquisaDARFSimples()

  On Error GoTo ERRO_PESQUISADARSIMPLES

  Dim sSql As String
  Dim RsDARFS As rdoResultset

  'Preencher os campos do DARF , caso encontre
  sSql = Geral.DataProcessamento & " , " & Geral.Documento.IdDocto

  Set qryGetDARFSimples = Geral.Banco.CreateQuery("", "{call GetDARFSimples (" & sSql & ")}")

  Set RsDARFS = qryGetDARFSimples.OpenResultset(rdOpenStatic, rdConcurReadOnly)
  If Not RsDARFS.EOF Then
    'Encontrou o DARF Preto -> Preencher os campos
    TxtPeriodo.Text = Mid(RsDARFS!PeriodoApuracao, 7, 2) & Mid(RsDARFS!PeriodoApuracao, 5, 2) & Mid(RsDARFS!PeriodoApuracao, 1, 4)
    TxtCGC.Text = RsDARFS!CGC
    TxtReceita.Text = Val(RsDARFS!ReceitaBruta * 100)
    TxtPercentual.Text = Val(RsDARFS!Percentual * 100)
    TxtPrincipal.Text = Val(RsDARFS!ValorPrincipal * 100)
    TxtMulta.Text = Val(RsDARFS!ValorMulta * 100)
    TxtJuros.Text = Val(RsDARFS!Juros * 100)
    TxtValor.Text = Val(RsDARFS!Valor * 100)

    'Posicionar o Foco no campo 'VALOR'
    TxtPrincipal.SetFocus
  Else
    'Posicionar o Foco no campo 'PERIODO APURAÇÃO'
    TxtPeriodo.SetFocus
  End If

  If AlteraValor = True Then
    'O Usuário só pode alterar os valores
    TxtPeriodo.Locked = True
    TxtCGC.Locked = True

    TxtPrincipal.SetFocus
  End If

  Screen.MousePointer = vbDefault

  Exit Sub

ERRO_PESQUISADARSIMPLES:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Selecionar Dados do DARF Simples.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Sub
Function CamposOK() As Boolean

    On Error GoTo ERRO_CAMPOSOK

    CamposOK = False

    'Período Apuração
    If Len(Trim(TxtPeriodo.Text)) = 0 Then
      MsgBox "Informe o Período de Apuração do Documento.", vbInformation, App.Title
      TxtPeriodo.SetFocus
      Exit Function
    End If

    'Ano Inicial e Final
    If Val(Mid(TxtPeriodo.Text, 5, 4)) > 2010 Or Val(Mid(TxtPeriodo.Text, 5, 4)) < 1980 Then
      MsgBox "O ano de apuração não pode ser menor que 1980 nem maior que 2010.", vbInformation, App.Title
      TxtPeriodo.SetFocus
      Exit Function
    End If
      
    'CGC
    If Len(Trim(TxtCGC.Text)) = 0 Then
      MsgBox "Informe o Número do CGC.", vbInformation, App.Title
      TxtCGC.SetFocus
      Exit Function
    End If
    
    'Receita Bruta
    If Len(Trim(TxtReceita.Text)) = 0 Then
      MsgBox "Informe o Valor da Receita Bruta.", vbInformation, App.Title
      TxtReceita.SetFocus
      Exit Function
    End If
    
    'Percentual
    If Len(Trim(TxtPercentual.Text)) = 0 Then
      MsgBox "Informe o Percentual.", vbInformation, App.Title
      TxtPercentual.SetFocus
      Exit Function
    End If
    
    'Validação do Percentual
    If CCur(TxtPercentual.Text / 100) < RegraValidaDARF(1).PercentualMinimo Or _
        CCur(TxtPercentual.Text / 100) > RegraValidaDARF(1).PercentualMaximo Then
      MsgBox "O Percentual deve estar entre " & _
            Trim(FormataValor(RegraValidaDARF(1).PercentualMinimo, 5)) & " e " & _
            Trim(FormataValor(RegraValidaDARF(1).PercentualMaximo, 5)) & " e " & _
            " .", vbInformation + vbOKOnly, App.Title
      TxtPercentual.SetFocus
      Exit Function
    End If
    
    'Principal
    If Val(TxtPrincipal.Text) = 0 Then
      MsgBox "Informe o Valor do Principal.", vbInformation, App.Title
      TxtPrincipal.SetFocus
      Exit Function
    End If
    
    Call CalculaValorFinal
    
    'Preenchimento do Valor
    If Len(Trim(TxtValor.Text)) = 0 Then
      MsgBox "Informe o Valor Final do Documento.", vbInformation, App.Title
      TxtValor.SetFocus
      Exit Function
    End If
    
    'Validação do Valor
    If (CCur(TxtValor.Text / 100) < RegraValidaDARF(1).ValorMinimoDocumento) And _
        (TxtPeriodo.Text <> "01011980" And TxtPeriodo.Text <> "08081980") Then
      MsgBox "Valor Arrecadado não pode ser inferior a R$ " & Trim(FormataValor(RegraValidaDARF(1).ValorMinimoDocumento, 10)) & " .", vbInformation + vbOKOnly, App.Title
      TxtReceita.SetFocus
      Exit Function
    End If

    '01.01.1980 ou 08.08.1980

    CamposOK = True

    Exit Function

ERRO_CAMPOSOK:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Verificar os valores dos campos.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function
Sub AjustesIniciais()

  'Setando as Variáveis do RDO
  Set qryGetDARFSimples = Geral.Banco.CreateQuery("", "{? = call GetDARFSimples (?,?)}")
  Set qryRemoveTipoDocumento = Geral.Banco.CreateQuery("", "{? = call RemoveTipoDocumento (?,?,?)}")
  Set qryAtualizaDARFSimples = Geral.Banco.CreateQuery("", "{call AtualizaDARFSimples (?,?,?,?,?,?,?,?,?,?,?,?,?)}")
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

  Set qryGetDARFSimples = Nothing
  Set qryRemoveTipoDocumento = Nothing
  Set qryAtualizaDARFSimples = Nothing
End Sub

Private Sub TxtCGC_Change()

  If Len(Trim(TxtCGC.Text)) = TxtCGC.MaxLength Then
    SendKeys "{TAB}"
    DoEvents
  End If
End Sub

Private Sub TxtCGC_GotFocus()

  TxtCGC.SelStart = 0
  TxtCGC.SelLength = Len(TxtCGC.Text)
End Sub


Private Sub TxtCGC_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub


Private Sub TxtCGC_Validate(Cancel As Boolean)

    'Formatando o Campo CGC
    TxtCGC.Text = Format(TxtCGC.Text, String(15, "0"))

    'Validação de CGC
    If Mid(Format(TxtCGC.Text, String(15, "0")), 10, 4) <> "0001" Then
        MsgBox "Só são aceitos CGC's da Matriz.", vbInformation, App.Title
        Cancel = True
    Else
      If Not VerificaCGC(TxtCGC.Text) Then
        MsgBox "CGC Inválido.", vbInformation, App.Title
        TxtCGC.SelStart = 0
        TxtCGC.SelLength = Len(TxtCGC.Text)
        Cancel = True
      End If
    End If
End Sub

Private Sub TxtJuros_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub


Private Sub txtJuros_LostFocus()

  Call CalculaValorFinal
End Sub


Private Sub txtMulta_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub


Private Sub txtMulta_LostFocus()

  Call CalculaValorFinal
End Sub

Private Sub TxtPercentual_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub
Private Sub TxtPercentual_LostFocus()

    'Validação do Receita
    If Len(Trim(TxtReceita.Text)) = 0 Then
        TxtReceita.SetFocus
        Exit Sub
    End If
    
    'Percentual
    If Len(Trim(TxtPercentual.Text)) = 0 Then
      Exit Sub
    End If

    'Validação do Percentual
    If CCur(TxtPercentual.Text / 100) < RegraValidaDARF(1).PercentualMinimo Or _
        CCur(TxtPercentual.Text / 100) > RegraValidaDARF(1).PercentualMaximo Then
      MsgBox "O Percentual deve estar entre " & _
            Trim(FormataValor(RegraValidaDARF(1).PercentualMinimo, 5)) & " e " & _
            Trim(FormataValor(RegraValidaDARF(1).PercentualMaximo, 5)) & _
            " .", vbInformation + vbOKOnly, App.Title
      TxtPercentual.SetFocus
      Exit Sub
    End If

    Call CalculaValorPrincipal
    Call CalculaValorFinal

    TxtPrincipal.SelStart = 0
    TxtPrincipal.SelLength = Len(TxtPrincipal.Text)
End Sub
Private Sub TxtPeriodo_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeySpace And TxtPeriodo.Locked = False Then
      KeyAscii = 0
      TxtPeriodo.Text = Mid(Geral.DataProcessamento, 7, 2) & Mid(Geral.DataProcessamento, 5, 2) & Mid(Geral.DataProcessamento, 1, 4)
      SendKeys "{TAB}"
  ElseIf KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub
Private Sub TxtPrincipal_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub
Private Sub TxtPrincipal_LostFocus()
  Call CalculaValorFinal
End Sub
Private Sub txtReceita_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub
Private Sub TxtReceita_LostFocus()
'* Valida Valor de Receita - Darf-Simples *'
    If Val(TxtReceita.Text) = 0 Then
        MsgBox "Informe o valor da Receita.", vbInformation, App.Title
        TxtReceita.SetFocus
    End If
End Sub
Private Sub txtValor_GotFocus()
  TxtValor.SelStart = 0
  TxtValor.SelLength = Len(TxtValor.Text)
End Sub
Private Sub txtValor_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    Call cmdConfirmar_Click
  End If
End Sub
Private Function CargaRegraValidaDARF() As Boolean

Dim rsGetRegra As rdoResultset
Dim qryGetRegra As rdoQuery

On Error GoTo Err_CargaRegraValidaDARF
    
    Erase RegraValidaDARF

    CargaRegraValidaDARF = False
    
    Set qryGetRegra = Geral.Banco.CreateQuery("", "{call GetRegraValidaDarfSimples }")
    
    Set rsGetRegra = qryGetRegra.OpenResultset(rdOpenStatic, rdConcurReadOnly)
    
    If Not rsGetRegra.EOF Then
        ReDim RegraValidaDARF(rsGetRegra.RowCount)
        
        While Not rsGetRegra.EOF
            RegraValidaDARF(rsGetRegra.AbsolutePosition).PercentualMinimo = rsGetRegra!PercentualMinimo
            RegraValidaDARF(rsGetRegra.AbsolutePosition).PercentualMaximo = rsGetRegra!PercentualMaximo
            RegraValidaDARF(rsGetRegra.AbsolutePosition).ValorMinimoDocumento = rsGetRegra!ValorMinimoDocumento

            rsGetRegra.MoveNext
        Wend
        CargaRegraValidaDARF = True
    Else
        MsgBox "Não há parâmetros de Regras para validação de DARF SIMPLES." & vbCrLf & vbCrLf & "Favor contatar o suporte.", vbCritical + vbOKOnly, App.Title
    End If
    
Exit_CargaRegraValidaDARF:
    If Not (rsGetRegra Is Nothing) Then rsGetRegra.Close
    qryGetRegra.Close
    
    Exit Function

Err_CargaRegraValidaDARF:
    Screen.MousePointer = vbDefault
    Beep
    Select Case TratamentoErro("Erro na carga das Regras de DARF SIMPLES.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    GoTo Exit_CargaRegraValidaDARF

End Function

