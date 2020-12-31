VERSION 5.00
Object = "{9CBA5D64-E3C8-11D3-9FFC-00104BC8688C}#1.0#0"; "CurrencyEdit.ocx"
Begin VB.Form GPS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Complementação de GPS"
   ClientHeight    =   2844
   ClientLeft      =   1296
   ClientTop       =   1320
   ClientWidth     =   7488
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2844
   ScaleWidth      =   7488
   Begin VB.TextBox TxtCompetencia 
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
      Left            =   1488
      MaxLength       =   7
      TabIndex        =   1
      Top             =   1404
      Width           =   1020
   End
   Begin VB.TextBox TxtCodPagto 
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
      Left            =   444
      MaxLength       =   4
      TabIndex        =   0
      Top             =   1404
      Width           =   600
   End
   Begin CURRENCYEDITLib.CurrencyEdit TxtOutrasEnt 
      Height          =   372
      Left            =   444
      TabIndex        =   4
      Top             =   2268
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
   Begin VB.CommandButton cmdZoomMais 
      Caption         =   "Zoom +"
      Height          =   696
      Left            =   1608
      Picture         =   "GPS.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   132
      Width           =   804
   End
   Begin VB.CommandButton cmdZoomMenos 
      Caption         =   "Zoom -"
      Height          =   696
      Left            =   2424
      Picture         =   "GPS.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   132
      Width           =   804
   End
   Begin VB.CommandButton cmdRotacao 
      Caption         =   "Rotação"
      Height          =   696
      Left            =   3240
      Picture         =   "GPS.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   132
      Width           =   804
   End
   Begin VB.CommandButton cmdInverteCor 
      Caption         =   "Inverter"
      Height          =   696
      Left            =   4056
      Picture         =   "GPS.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   132
      Width           =   804
   End
   Begin VB.CommandButton cmdFrenteVerso 
      Caption         =   "Fren/Ver"
      Height          =   696
      Left            =   4872
      Picture         =   "GPS.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   132
      Width           =   804
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   696
      Left            =   6480
      Picture         =   "GPS.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   132
      Width           =   816
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "&Confirmar"
      Height          =   696
      Left            =   5676
      Picture         =   "GPS.frx":123C
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   132
      Width           =   804
   End
   Begin VB.TextBox TxtIdentificador 
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
      Left            =   2952
      MaxLength       =   15
      TabIndex        =   2
      Top             =   1404
      Width           =   1920
   End
   Begin CURRENCYEDITLib.CurrencyEdit TxtJuros 
      Height          =   372
      Left            =   2916
      TabIndex        =   5
      Top             =   2268
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
   Begin CURRENCYEDITLib.CurrencyEdit TxtValorINSS 
      Height          =   372
      Left            =   5400
      TabIndex        =   3
      Top             =   1404
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
   Begin CURRENCYEDITLib.CurrencyEdit TxtTotal 
      Height          =   372
      Left            =   5400
      TabIndex        =   6
      Top             =   2268
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
      Left            =   192
      Picture         =   "GPS.frx":1546
      Top             =   312
      Width           =   384
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      Caption         =   "GPS"
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
      TabIndex        =   21
      Top             =   432
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Cód. Pagto"
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
      Left            =   252
      TabIndex        =   20
      Top             =   1092
      Width           =   972
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Competência"
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
      Left            =   1476
      TabIndex        =   19
      Top             =   1092
      Width           =   1176
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Identificador"
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
      Left            =   2976
      TabIndex        =   18
      Top             =   1092
      Width           =   1092
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor do INSS"
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
      Left            =   5472
      TabIndex        =   17
      Top             =   1092
      Width           =   1236
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Outras Entidades"
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
      Left            =   252
      TabIndex        =   16
      Top             =   1920
      Width           =   2064
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "At.Monet. / Juros/Multa"
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
      Left            =   2904
      TabIndex        =   15
      Top             =   1968
      Width           =   2052
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Arrecadado"
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
      Left            =   5376
      TabIndex        =   14
      Top             =   1980
      Width           =   1548
   End
End
Attribute VB_Name = "GPS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Declaração de Variáveis do RDO
Private qryRemoveTipoDocumento As rdoQuery
Private qryAtualizaGPS As rdoQuery
Private qryGetGps As rdoQuery
Private qryGetValidaGPS As rdoQuery

'Declaração de Variáveis de trabalho
Private mForm As Form
Public Alterou As Boolean
Public AlteraValor As Boolean
Private bActivate As Boolean

'Declaração de Variáveis de Flag
Private Flag_CodigoPagto As String * 1
Private Flag_DigitaValorINSS As String * 1
Private Flag_DigitaValorOutrasEntidades As String * 1
Private Flag_DataInicioVigencia As String * 1
Private Flag_DataFinalVigencia As String * 1
Private Flag_IndicMesCompetencia As String * 1
Private Flag_LimiteInicialCompetencia As String * 1
Private Flag_IndicRestricaoPagto As String * 1
Private Flag_CondPagtoNormal As String * 1
Private Flag_Datapagto13 As String * 1
Private Flag_CondPagto13 As String * 1

'Validação de informações do GPS
Private RegraValidaGPS() As tpValidacao
Private Type tpValidacao
    AnoInicial              As Long
    AnoFinal                As Long
    ValorMinimoDocumento    As Currency
End Type

'Declaração do Type
Private Pagto As TpPagto

'Declaração do Type do Tipo de Documento
Private Type TpPagto
  CodigoPagamento As Long
  TipoDocumento As String * 1
  DigitaValorINSS As String * 1
  DigitaOutrasEntidades As String * 1
  VerificaAtraso As String * 1
  DataInicial As Long
  DataFinal As Long
  IndicMesCompetencia As String * 1
  LimiteInicialCompetencia As Integer
  IndicRestricaoPagto As String * 1
  DataPagtoNormal As String * 2
  CondPagtoNormal As String * 1
  DataPagto13 As Integer
  CondPagto13 As String * 1
End Type
Private Function CalculaDataLimite(ByVal DataPagto As String) As Long

  Dim sDataLimite As String

  'Concatenar os campo para formar a data base
  sDataLimite = Format(Val(DataPagto), "00") & "/" & TxtCompetencia.Text

  'Acrescentar um mes à data base
  sDataLimite = Format(DateAdd("m", 1, sDataLimite), "dd/mm/yyyy")

  'Verificar se é Final de Semana
  If Weekday(sDataLimite) = vbSaturday Then
    'Data Limite = Sábado -> Acrescentar dois dias
    sDataLimite = Format(DateAdd("d", 2, sDataLimite), "dd/mm/yyyy")
  ElseIf Weekday(sDataLimite) = vbSunday Then
    'Data Limite = Domingo -> Acrescentar um dia
    sDataLimite = Format(DateAdd("d", 1, sDataLimite), "dd/mm/yyyy")
  End If

  CalculaDataLimite = Mid(sDataLimite, 7, 4) & Mid(sDataLimite, 4, 2) & Mid(sDataLimite, 1, 2)
End Function
Function ValidaAgenciaGPS(ByVal CodigoAgencia As Integer, ByVal sData As String) As Boolean

Dim RsAgenf As rdoResultset
Dim qryGetAgenf As rdoQuery
Dim intAgefsstmovi As Integer, lngAgefsdtmvan As Long

On Error GoTo ERRO_VALIDAAGENCIAGPS

    ValidaAgenciaGPS = False

    If Geral.Capa.agefsestado = "" Then
        'Verificar o Status da Agencia
        Set qryGetAgenf = Geral.Banco.CreateQuery("", "{call GetAgenf (" & CodigoAgencia & ")}")

        Set RsAgenf = qryGetAgenf.OpenResultset(rdOpenStatic, rdConcurReadOnly)
  
        If RsAgenf.EOF Then
            MsgBox "A Agência de Origem não está Cadastrada.", vbInformation, App.Title
            Exit Function
        End If
        intAgefsstmovi = RsAgenf!agefsstmovi
        lngAgefsdtmvan = RsAgenf!agefsdtmvan
    Else
        intAgefsstmovi = Geral.Capa.agefsstmovi
        lngAgefsdtmvan = Geral.Capa.agefsdtmvan
    End If
      

    'A Agencia está cadastrada -> Verificar o Status
    If intAgefsstmovi = 9 Then
      'Feriado
      MsgBox "A Agência de Origem está em Feriado.", vbInformation, App.Title
      Exit Function
    ElseIf intAgefsstmovi = 0 Then
      'Agencia Fechada
      MsgBox "A Agência de Origem está Fechada.", vbInformation, App.Title
      Exit Function
    ElseIf intAgefsstmovi = 2 Then
      'Agencia Aberta -> Verificar data do Movimento Anterior
      If sData < TransformaDataAAAAMMDD(lngAgefsdtmvan) Then
        'A Data é menor à data do Movimento Anterior -> Não Aceitar
        If MsgBox("Este Pagamento deveria ter sido efetuado até o dia " & Mid(sData, 7, 2) & "/" & Mid(sData, 5, 2) & "/" & Mid(sData, 1, 4) & ". Confirma ? ", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
          'Marcar Flag de erro
          Flag_IndicRestricaoPagto = "1"
        Else
          TxtCompetencia.SetFocus
          Exit Function
        End If
      End If
    End If

  ValidaAgenciaGPS = True

  Exit Function

ERRO_VALIDAAGENCIAGPS:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Validar a Agência.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function
Private Function DigitaOutrasEntidades() As Boolean

  On Error GoTo ERRO_DIGITAOUTRASENTIDADES

  DigitaOutrasEntidades = False

  'Marcar Flag de Erro
  Flag_DigitaValorOutrasEntidades = "0"

  'Verificar se o código de pagamento permite a digitação do valor de Outras Entidades
  If Val(Pagto.DigitaOutrasEntidades) = 1 Then
    'Obrigatório
    If Val(TxtOutrasEnt.Text) = 0 Then
      If MsgBox("Para este Código de Pagamento é obrigatório a digitação do Valor de Outras Entidades.Confirma ?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Function

      'Marcar Flag de Erro
      Flag_DigitaValorOutrasEntidades = "1"

    End If
  ElseIf Val(Pagto.DigitaOutrasEntidades) = 0 Then
    'Não Permitido
    If Val(TxtOutrasEnt.Text) <> 0 Then
      If MsgBox("Para este Código de Pagamento não é permitido informar o Valor de Outras Entidades.Confirma ?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then
        TxtOutrasEnt.SetFocus
        Exit Function
      End If

      'Marcar Flag de Erro
      Flag_DigitaValorOutrasEntidades = "1"

    End If
  ElseIf Val(Pagto.DigitaOutrasEntidades) <> 2 Then
    MsgBox "A Tabela 'ValidaGPS' possui um valor incorreto no campo 'DigitaOutrasEntidades' para o código de Pagamento : " & Pagto.CodigoPagamento & ". Informe o Suporte Informática.", vbInformation, App.Title
    Exit Function
  End If

  DigitaOutrasEntidades = True

  Exit Function

ERRO_DIGITAOUTRASENTIDADES:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Validar Valor de Outras Entidades.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function
Private Function DigitaValorINSS() As Boolean

  On Error GoTo ERRO_DIGITAVALORINSS

  DigitaValorINSS = False

  'Marcar Flag de Erro
  Flag_DigitaValorINSS = "0"

  'Verificar se o código de pagamento permite a digitação do valor do INSS
  If Val(Pagto.DigitaValorINSS) = 1 Then
    'Obrigatório
    If Val(TxtValorINSS.Text) = 0 Then
      If MsgBox("Para este Código de Pagamento é obrigatório a digitação do Valor do INSS.Confirma ?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then
        TxtValorINSS.SetFocus
        Exit Function
      End If

      'Marcar Flag de Erro
      Flag_DigitaValorINSS = "1"

    End If
  ElseIf Val(Pagto.DigitaValorINSS) = 0 Then
    'Não Permitido
    If Val(TxtValorINSS.Text) <> 0 Then
      If MsgBox("Para este Código de Pagamento não é permitido informar o Valor do INSS.Confirma ?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Function

      'Marcar Flag de Erro
      Flag_DigitaValorINSS = "1"

    End If
  ElseIf Val(Pagto.DigitaValorINSS) <> 2 Then
    MsgBox "A Tabela 'ValidaGPS' possui um valor incorreto no campo 'DigitaValorINSS' para o código de Pagamento : " & Pagto.CodigoPagamento & ". Informe o Suporte Informática.", vbInformation, App.Title
    Exit Function
  End If

  DigitaValorINSS = True

  Exit Function

ERRO_DIGITAVALORINSS:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Validar Digitação de Valor do INSS.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function
Function ValidaCodigoPagto(ByVal CodigoPagto As Long) As Boolean

  On Error GoTo ERRO_VALIDACODIGOPAGTO

  Dim RsValidaGPS As rdoResultset

  ValidaCodigoPagto = False

  Set qryGetValidaGPS = Geral.Banco.CreateQuery("", "{call GetValidaGPS (" & CodigoPagto & ")}")

  Set RsValidaGPS = qryGetValidaGPS.OpenResultset(rdOpenStatic, rdConcurReadOnly)
  If Not RsValidaGPS.EOF Then
    'Encontrou o Código de Pagamento -> Preencher o Type
    Pagto.CodigoPagamento = CodigoPagto
    Pagto.TipoDocumento = RsValidaGPS!TipoDocumento
    Pagto.DigitaValorINSS = RsValidaGPS!DigitaValorINSS
    Pagto.DigitaOutrasEntidades = RsValidaGPS!DigitaOutrasEntidades
    Pagto.VerificaAtraso = RsValidaGPS!VerificaAtraso
    Pagto.DataInicial = RsValidaGPS!DataInicial
    Pagto.DataFinal = RsValidaGPS!DataFinal
    Pagto.IndicMesCompetencia = RsValidaGPS!IndicMesCompetencia
    Pagto.LimiteInicialCompetencia = RsValidaGPS!LimiteInicialCompetencia
    Pagto.IndicRestricaoPagto = RsValidaGPS!IndicRestricaoPagto
    Pagto.DataPagtoNormal = RsValidaGPS!DataPagtoNormal
    Pagto.CondPagtoNormal = RsValidaGPS!CondPagtoNormal
    Pagto.DataPagto13 = RsValidaGPS!DataPagto13
    Pagto.CondPagto13 = RsValidaGPS!CondPagto13
  Else
    'Não encontrou o Código de Pagamento -> Limpar o type
    Call LimpaTypePagto
    Exit Function
  End If

  ValidaCodigoPagto = True

  Exit Function

ERRO_VALIDACODIGOPAGTO:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Validar Código de Pagamento.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function
Sub LimpaTypePagto()

  Pagto.CodigoPagamento = 0
  Pagto.TipoDocumento = ""
  Pagto.DigitaValorINSS = ""
  Pagto.DigitaOutrasEntidades = ""
  Pagto.VerificaAtraso = ""
  Pagto.DataInicial = 0
  Pagto.DataFinal = 0
  Pagto.IndicMesCompetencia = ""
  Pagto.LimiteInicialCompetencia = 0
  Pagto.IndicRestricaoPagto = ""
  Pagto.DataPagtoNormal = ""
  Pagto.CondPagtoNormal = ""
  Pagto.DataPagto13 = 0
  Pagto.CondPagto13 = ""
End Sub
Private Function ValidaIdentificador(Identificador As String, ByVal PossuiTipo As Boolean, ByRef sMens As String) As Boolean

  On Error GoTo ERRO_VALIDAIDENTIFICADOR

  Dim Ident As String

  sMens = ""

  ValidaIdentificador = False

  If PossuiTipo Then
    'Já Possui um tipo de documento definido -> Validar DV
    Select Case Pagto.TipoDocumento
      Case "1"
        'DEBCAD - CPF
        Identificador = Format(Identificador, String(14, "0"))
        If Left(Identificador, 5) <> "00000" Or Mid(Identificador, 6, 8) = "00000000" Then Exit Function
        
        ValidaIdentificador = VerificaMODULO11(Identificador)
        Exit Function

      Case "2"
        'Referencia
        Identificador = Format(Identificador, String(14, "0"))
        If Left(Identificador, 13) = "0000000000000" Then Exit Function

        ValidaIdentificador = VerificaMODULO11(Identificador)
        Exit Function

      Case "3"
        'CNPJ
        Identificador = Format(Identificador, String(15, "0"))
        If Left(Identificador, 8) = "00000000" Or Identificador = "111111111111111" Or _
            Identificador = "222222222222222" Or Identificador = "333333333333333" Or _
            Identificador = "444444444444444" Or Identificador = "555555555555555" Or _
            Identificador = "666666666666666" Or Identificador = "777777777777777" Or _
            Identificador = "888888888888888" Or Identificador = "999999999999999" Then Exit Function

        ValidaIdentificador = VerificaCGC(Identificador)
        Exit Function

      Case "5"
        'CNPJ Matriz
        Identificador = Format(Identificador, String(15, "0"))
        If Mid(Identificador, 10, 4) <> "0001" Then
          'Não Aceitar CNPJ da Matriz
          sMens = "Para este Código de Pagamento só é permitido CNPJ da Matriz."
          ValidaIdentificador = False
          Exit Function
        End If

        If Left(Identificador, 8) = "00000000" Or Identificador = "111111111111111" Or _
            Identificador = "222222222222222" Or Identificador = "333333333333333" Or _
            Identificador = "444444444444444" Or Identificador = "555555555555555" Or _
            Identificador = "666666666666666" Or Identificador = "777777777777777" Or _
            Identificador = "888888888888888" Or Identificador = "999999999999999" Then Exit Function

        ValidaIdentificador = VerificaCGC(Identificador)
        Exit Function

      Case "6"
        'CEI
        Identificador = Format(Identificador, String(12, "0"))

        ValidaIdentificador = VerificaCEI(Identificador)
        Exit Function

      Case "7"
        'NIT NR
        Identificador = Format(Identificador, String(11, "0"))

        If Identificador < 1000000000 Then Exit Function

        ValidaIdentificador = VerificaNIT(Identificador)
        Exit Function

      Case "8"
        'NB NR Beneficio
        Identificador = Format(Identificador, String(10, "0"))

        ValidaIdentificador = VerificaNB(Identificador)
        Exit Function

      Case "9"
        'NTC NR Titulo
        Identificador = Format(Identificador, String(14, "0"))

        ValidaIdentificador = VerificaMODULO11(Identificador)

        Exit Function

    End Select
  Else
    'Não possui um tipo de documento definido -> Verificar todos os DVs

    'DEBCAD - CPF
    Ident = Format(Identificador, String(14, "0"))
    If Left(Ident, 5) = "00000" Or Mid(Ident, 6, 8) <> "00000000" Then
      If VerificaMODULO11(Ident) Then
        ValidaIdentificador = True
        Exit Function
      End If
    End If

    'Referencia
    Ident = Format(Identificador, String(14, "0"))
    If Left(Ident, 13) <> "0000000000000" Then

      If VerificaMODULO11(Ident) Then
        ValidaIdentificador = True
        Exit Function
      End If

    End If

    'CNPJ e CNPJ Matriz
    Ident = Format(Identificador, String(15, "0"))

    If VerificaCGC(Ident) Then
      ValidaIdentificador = True
      Exit Function
    End If

    'CEI
    Ident = Format(Identificador, String(12, "0"))

    If VerificaCEI(Ident) Then
      ValidaIdentificador = True
      Exit Function
    End If

    'NIT NR
    Ident = Format(Identificador, String(11, "0"))

    If VerificaNIT(Ident) Then
      ValidaIdentificador = True
      Exit Function
    End If

    'NB NR Beneficio
    Ident = Format(Identificador, String(10, "0"))

    If VerificaNB(Ident) Then
      ValidaIdentificador = True
      Exit Function
    End If

    'NTC NR Titulo
    Ident = Format(Identificador, String(14, "0"))

    If VerificaMODULO11(Ident) Then
      ValidaIdentificador = True
    End If

    Exit Function

  End If

  ValidaIdentificador = True

  Exit Function

ERRO_VALIDAIDENTIFICADOR:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Validar o Identificador.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function
Private Function ValidaData(ByVal Mes As String, Ano As String) As Boolean

   ValidaData = False

   If Mes > 0 And Mes < 14 Then
      'Mes Válido -> Formatar e Verificar Ano
      Mes = Format(Mes, "00")
      If Ano < 1900 Or Ano > 2100 Then
         'Ano inválido -> Emitir Mensagem
         MsgBox "Ano de Competência Inválido.", vbInformation, App.Title
         TxtCompetencia.SetFocus
         Exit Function
      Else
         'Ano Válido -> Formatar e preencher campo formatado
         TxtCompetencia.Text = Mes & "/" & Format(Ano, "0000")
      End If
   Else
      'Mes Inválido -> Emitir Mensagem
      MsgBox "Mês de Competência Inválido.", vbInformation, App.Title
      TxtCompetencia.SetFocus
      Exit Function
   End If

   ValidaData = True
End Function
Private Function VerificaAtraso() As Boolean

   On Error GoTo ERRO_VERIFICAATRASO

   Dim DataLimite As Long

   VerificaAtraso = False

   If Val(Pagto.VerificaAtraso) = 1 And Mid(TxtCompetencia.Text, 1, 2) < 13 Then

      'Verificar se o mes de competencia é igual a 13
      If Val(Left(TxtCompetencia.Text, 2)) <> 13 Then
         'Verificar Atraso
         If Val(Pagto.DataPagtoNormal) > 0 Then
            'Verificar se a data do Movimento é maior que o próximo dia útil
            'após DataPagtoNormal, se este não for útil, do mes subsequente à competencia

            'Calcular Data Limite
            DataLimite = CalculaDataLimite(Pagto.DataPagtoNormal)

            If Not ValidaAgenciaGPS(Geral.Documento.Agencia, DataLimite) Then Exit Function

            'Verificar se a data do movimento é maior que a data limite
            If Geral.DataProcessamento > DataLimite Then
               If MsgBox("Para este Código de Pagamento não é permitido o pagamento de Documentos Atrasados. Confirma ? ", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
                  'O Usuário confirmou a ação -> Verificar se o Valor da Multa foi informado
                  If Val(TxtJuros.Text) = 0 Then
                     If MsgBox("É Obrigatório informar o valor da Multa para Documentos Atrasados. Confirma ? ", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then
                       TxtJuros.SetFocus
                        Exit Function
                     End If
                  End If
               Else
                  TxtCompetencia.SetFocus
               End If
            End If
         End If
      End If
   End If

   VerificaAtraso = True

   Exit Function

ERRO_VERIFICAATRASO:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Verificar Atraso.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function
Function VerificaCodigoPagto() As Boolean

  'Marcar Flag de Erro
  Flag_CodigoPagto = "0"

  'Verificar se o Código de Pagamento está cadastrado
  If Len(Trim(TxtCodPagto.Text)) > 0 And Not bActivate Then
    If Not ValidaCodigoPagto(Val(TxtCodPagto.Text)) Then
      'Emitir Mensagem informando que o código de pagamento não está cadastrado , se habilitado
      If TxtCodPagto.Locked = False Then
        If MsgBox("Código de Pagamento não cadastrado. Confirma ? ", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
          'Marcar Flag de Erro
          Flag_CodigoPagto = "1"
          VerificaCodigoPagto = True
        Else
          TxtCodPagto.SetFocus
          VerificaCodigoPagto = False
        End If
      Else
        Flag_CodigoPagto = "1"
        VerificaCodigoPagto = True
      End If
    Else
      VerificaCodigoPagto = True
    End If
  End If
End Function
Private Function VerificaCondPagtoNormal() As Boolean

  On Error GoTo ERRO_VERIFICACONDPAGTONORMAL

  VerificaCondPagtoNormal = False

  'Marcar Flag de Erro
  Flag_CondPagtoNormal = "0"

  'Se Mes de competencia = 13 -> Validar com função 'VerificaDataPagto13'
  If Val(Mid(TxtCompetencia.Text, 1, 2)) = 13 Then
    VerificaCondPagtoNormal = True
    Exit Function
  End If

  Select Case Pagto.CondPagtoNormal
    Case "0"
      'Não Validar

    Case "2"
      'Pode Antecipar
      'Verificar se o ano de competencia é menor ou igual ao movimento
      If Val(Mid(TxtCompetencia.Text, 4, 4)) < Val(Mid(Geral.DataProcessamento, 1, 4)) Then
        'Emitir Mensagem
        If MsgBox("Este Documento não pode ser pago com atraso. Confirma ?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
          'Marcar Flag de Erro
          Flag_CondPagtoNormal = "1"
        Else
          TxtCompetencia.SetFocus
          Exit Function
        End If
      ElseIf Val(Mid(TxtCompetencia.Text, 4, 4)) = Val(Mid(Geral.DataProcessamento, 1, 4)) Then
        'Verificar se o mes de competencia é menor que o mes do movimento
        If Val(Mid(TxtCompetencia.Text, 1, 2)) < Val(Mid(Geral.DataProcessamento, 3, 2)) Then
          'Mes de Competencia menor que mes do movimento -> Emitir Mensagem
          If MsgBox("Este Documento não pode ser pago com atraso. Confirma ?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
            'Marcar Flag de Erro
            Flag_CondPagtoNormal = "1"
          Else
            TxtCompetencia.SetFocus
            Exit Function
          End If
        End If
      End If

    Case "3"
      'Pode Postecipar
      'Verificar se o ano de competencia é maior que o ano do movimento
      If Val(Mid(TxtCompetencia.Text, 4, 4)) > Val(Mid(Geral.DataProcessamento, 1, 4)) Then
        'Emitir Mensagem
        If MsgBox("Este Documento não pode ser pago Antecipado. Confirma ?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
          'Marcar Flag de Erro
          Flag_CondPagtoNormal = "1"
        Else
          TxtCompetencia.SetFocus
          Exit Function
        End If
      ElseIf Val(Mid(TxtCompetencia.Text, 4, 4)) = Val(Mid(Geral.DataProcessamento, 1, 4)) Then
        'Verificar se o mes de competencia é maior que o mes do movimento
        If Val(Mid(TxtCompetencia.Text, 1, 2)) > Val(Mid(Geral.DataProcessamento, 5, 2)) Then
          'Mes de Competencia maior que mes do movimento -> Emitir Mensagem
          If MsgBox("Este Documento não pode ser pago Antecipado. Confirma ?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
            'Marcar Flag de Erro
            Flag_CondPagtoNormal = "1"
          Else
            TxtCompetencia.SetFocus
            Exit Function
          End If
        End If
      End If
  End Select

  VerificaCondPagtoNormal = True

  Exit Function

ERRO_VERIFICACONDPAGTONORMAL:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Verificar Condição de Pagamento Normal.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function
Private Function VerificaConfirmacao() As String

  On Error GoTo ERRO_VERIFICACONFIRMACAO

  'Verificar se pelo menos um flag está marcado com erro
  If Flag_CodigoPagto = "1" Or Flag_DigitaValorINSS = "1" Or _
      Flag_DigitaValorOutrasEntidades = "1" Or Flag_DataInicioVigencia = "1" Or _
      Flag_DataFinalVigencia = "1" Or Flag_IndicMesCompetencia = "1" Or _
      Flag_LimiteInicialCompetencia = "1" Or Flag_IndicRestricaoPagto = "1" Or _
      Flag_CondPagtoNormal = "1" Or Flag_Datapagto13 = "1" Or _
      Flag_CondPagto13 = "1" Then

    VerificaConfirmacao = "1"

  Else

    VerificaConfirmacao = "0"

  End If

  Exit Function

ERRO_VERIFICACONFIRMACAO:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Verificar Confirmação.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function
Private Function VerificaDataInicialVigencia() As Boolean

  On Error GoTo ERRO_VERIFICADATAINICIALVIGENCIA

  VerificaDataInicialVigencia = False

  'Marcar Flag de Erro
  Flag_DataInicioVigencia = "0"

  'Verificar se a data do movimento é menor que a Data Inicial de Vigencia
  If Geral.DataProcessamento < Pagto.DataInicial Then
    MsgBox "A Data do Movimento é menor que a Data Inicial de Vigência.", vbInformation + vbOKOnly, App.Title
    TxtCodPagto.SetFocus
    Exit Function
  End If

  VerificaDataInicialVigencia = True

  Exit Function

ERRO_VERIFICADATAINICIALVIGENCIA:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Verificar Data Inicial de Vigência.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function
Private Function VerificaDataFinalVigencia() As Boolean

  On Error GoTo ERRO_VERIFICADATAFINALVIGENCIA

  VerificaDataFinalVigencia = False

  'Marcar Flag de Erro
  Flag_DataFinalVigencia = "0"

  'Verificar se a data do movimento é maior que a Data Final de Vigencia
  If Geral.DataProcessamento > Pagto.DataFinal And Val(Pagto.DataFinal) <> 0 Then
    MsgBox "A Data do Movimento é maior que a Data Final de Vigência.", vbInformation + vbOKOnly, App.Title
    TxtCodPagto.SetFocus
    Exit Function
  End If

  VerificaDataFinalVigencia = True

  Exit Function

ERRO_VERIFICADATAFINALVIGENCIA:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Verificar Data Final de Vigência.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function
Private Function VerificaDataPagto13() As Boolean

  On Error GoTo ERRO_VERIFICADATAPAGTO13

  VerificaDataPagto13 = False

  If Val(Mid(TxtCompetencia.Text, 1, 2)) = 13 Then
    'O Mes de competencia é igual a 13 -> Verificar Campo 'Datapagto13'
    If Pagto.DataPagto13 = "2012" Then
      'Verificar se o ano de competecia é igual ao ano do movimento
      If Val(Mid(TxtCompetencia.Text, 4, 4)) = Val(Mid(Geral.DataProcessamento, 1, 4)) Then
        'Anos iguais -> Verificar Dia e Mes do Movimento
        If Val(Mid(Geral.DataProcessamento, 5, 4)) > 1220 Then
          'Já passou de 20/12 -> Verificar se pode Postecipar
          If Pagto.CondPagto13 <> "3" Then
            'Não pode postecipar -> Emitir Mensagem
            If MsgBox("Este Pagamento não pode ser efetuado com Atraso. Confirma ? ", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
              'Marcar Flag de Erro
              Flag_CondPagto13 = "1"
            Else
              TxtCompetencia.SetFocus
            End If
          End If
        End If
      ElseIf Val(Mid(TxtCompetencia.Text, 4, 4)) < Val(Mid(Geral.DataProcessamento, 1, 4)) Then
        'Ano de Competencia menor que ano do movimento -> Verificar se pode Postecipar
        If Pagto.CondPagto13 <> "3" Then
          'Não pode postecipar -> Emitir Mensagem
          If MsgBox("Este Pagamento não pode ser efetuado com Atraso. Confirma ? ", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
            'Marcar Flag de Erro
            Flag_CondPagto13 = "1"
          Else
            TxtCompetencia.SetFocus
          End If
        End If
      ElseIf Val(Mid(TxtCompetencia.Text, 4, 4)) > Val(Mid(Geral.DataProcessamento, 1, 4)) Then
        'Ano de Competencia maior que ano do movimento -> Verificar se pode Antecipar
        If Pagto.CondPagto13 <> "2" Then
          'Não pode Antecipar -> Emitir Mensagem
          If MsgBox("Este Pagamento não pode ser Antecipado. Confirma ? ", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
            'Marcar Flag de Erro
            Flag_CondPagto13 = "1"
          Else
            TxtCompetencia.SetFocus
          End If
        End If
      End If
    Else
      If MsgBox("Para este Código de Pagamento não pode ser Competência = 13. Confirma ? ", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
        'Marcar Flag de Erro
        Flag_CondPagto13 = "1"
      Else
        TxtCompetencia.SetFocus
        Exit Function
      End If
    End If
  End If

  VerificaDataPagto13 = True

  Exit Function

ERRO_VERIFICADATAPAGTO13:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Verificar Data de Pagamento para Décimo Terceiro.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function
Private Sub cmdConfirmar_Click()

  If SalvaGPS Then
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


Sub CalculaValorTotal()

  On Error GoTo ERRO_CALCULAVALORTOTAL

  Dim Valor As Currency

  '--------------------   Calcular o Valor Total do Documento --------------------

  'Verificar se foi informado o Valor do INSS
  If Val(TxtValorINSS.Text) = 0 Then
    Valor = 0
  Else
    Valor = TxtValorINSS.Text
  End If

  'Verificar se foi informado 'Valor Outras Entidades'
  If Val(TxtOutrasEnt.Text) <> 0 Then
    Valor = Valor + Val(TxtOutrasEnt.Text)
  End If

  'Verificar se foi informado Juros
  If Val(TxtJuros.Text) <> 0 Then
    Valor = Valor + Val(TxtJuros.Text)
  End If

  'Transportar o Valor Final para a tela
  TxtTotal.Text = Valor

  Exit Sub

ERRO_CALCULAVALORTOTAL:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Calcular Valor Total do Documento.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Sub
Function VerificaMODULO11(ByVal Ident As String) As Boolean
   
  On Error GoTo ERRO_VERIFICAMODULO11

  'Esta rotina serve para conferir o DEBCAD, NUM.TIT.COB. E REFERÊNCIA tamanho = 14

  Dim soma As Integer, resto As Integer
  Dim digito_11 As Integer, p As Integer, peso As Integer
  Dim digito_rv As String
  Dim bOk As Boolean

  bOk = True           'default - OK

  soma = 0
  resto = 0
  digito_11 = 0        'calculado pelo módulo 11
  digito_rv = ""       'caracter digitado pelo operador

  'número do DEBCAD: (13+1)             0 0 0 0 0 N N N N N N N N - D
  '                                     x x x x x x x x x x x x x
  'multiplica da direita para esquerda: 6 5 4 3 2 9 8 7 6 5 4 3 2

  peso = 2    'começa multiplicar da direita para esquerda
  p = 13      'tamanho do campo sem o digito

  Do
    '*** Peso de 2 a 9 (multiplicação dos caracteres de 2 a 9) ***
    soma = soma + Mid(Ident, p, 1) * peso
    p = p - 1            'ponteiro
    peso = peso + 1      'peso
    If (peso = 10) Then
      peso = 2
    End If
    If (p = 0) Then
      Exit Do
    End If
  Loop

  resto = soma Mod 11      'resto da divisão
  digito_11 = resto        'digito verificador
   
  '*** se o calculo for igual a 0 ou 1, muda-se para 0 ***
  If (digito_11 = 0) Or (digito_11 = 1) Then
    digito_11 = 0
  Else
    digito_11 = 11 - resto
  End If

  digito_rv = Mid(Ident, 14, 1)  'digito verificador

  If CStr(digito_11) <> (digito_rv) Then
    bOk = False               'digito não confere
    VerificaMODULO11 = bOk
    Exit Function
  End If

  VerificaMODULO11 = bOk

  Exit Function

ERRO_VERIFICAMODULO11:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Verificar Modulo 11.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function
Function VerificaNIT(ByVal NIT As String) As Boolean
      
  On Error GoTo ERRO_VERIFICANIT

  'Esta rotina serve para conferir o NIT: tam = 11

  Dim soma As Integer, resto As Integer
  Dim i As Integer
  Dim peso(10) As Integer
  Dim digito As Integer
  Dim digito_11 As Integer
  Dim digito_nit As String
  Dim bOk As Boolean

  bOk = True           'default - OK

  soma = 0
  resto = 0
  digito = 0           'calculado pelo módulo
  digito_nit = ""       'caracter digitado pelo operador

  'Número do NIT: (10+1)                 N N N N N N N N N N-D
  '                                      x x x x x x x x x x-x
  'Multiplica os números pelo valores :  3 2 9 8 7 6 5 4 3 2-x

  peso(1) = 3
  peso(2) = 2
  peso(3) = 9
  peso(4) = 8
  peso(5) = 7
  peso(6) = 6
  peso(7) = 5
  peso(8) = 4
  peso(9) = 3
  peso(10) = 2

  For i = 1 To 10
    soma = soma + Val(Mid(NIT, i, 1)) * peso(i)
  Next i

  resto = soma Mod 11                    'resto da divisão por 11
  digito_11 = 11 - resto                 'digito verificador

  'Se o calculo for igual a 10 ou 11, muda-se para 0 ***
  If (digito_11 = 11) Or (digito_11 = 10) Then
    digito_11 = 0
  End If

  digito_nit = Mid(NIT, 11, 1)           'digito verificador digitado

  If CStr(digito_11) <> (digito_nit) Then
    bOk = False                         'digito não confere
    VerificaNIT = bOk
    Exit Function
  End If

  VerificaNIT = bOk

  Exit Function

ERRO_VERIFICANIT:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Verificar NIT.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function
Function VerificaNB(ByVal NB As String) As Boolean

  On Error GoTo ERRO_VERIFICANB

  'Esta rotina serve para conferir o NB tamanho = 10

  Dim soma As Integer, resto As Integer
  Dim digito_10 As Integer, p As Integer, peso As Integer
  Dim digito_rv As String
  Dim digito_11 As String
  Dim bOk As Boolean

  bOk = True           'default - OK

  soma = 0
  resto = 0
  digito_11 = 0        'calculado pelo módulo 11
  digito_rv = ""       'caracter digitado pelo operador

  'número do DEBCAD: (9+1)              N N N N N N N N N - D
  '                                     x x x x x x x x x   x
  'multiplica da direita para esquerda: 2 9 8 7 6 5 4 3 2

  peso = 2    'começa multiplicar da direita para esquerda
  p = 9       'tamanho do campo sem o digito

  Do
    '*** Peso de 2 a 9 (multiplicação dos caracteres de 2 a 9) ***
    soma = soma + Mid(NB, p, 1) * peso
    p = p - 1            'ponteiro
    peso = peso + 1      'peso
    If (peso = 10) Then
      peso = 2
    End If
    If (p = 0) Then
      Exit Do
    End If
  Loop

  resto = soma Mod 11      'resto da divisão
  digito_10 = resto        'digito verificador

  '*** se o calculo for igual a 0 ou 1, muda-se para 0 ***
  If (digito_10 = 0) Or (digito_10 = 1) Then
    digito_10 = 0
  Else
    digito_10 = 11 - resto
  End If

  digito_rv = Mid(NB, 10, 1)  'digito verificador

  If CStr(digito_10) <> (digito_rv) Then
    bOk = False               'digito não confere
    VerificaNB = bOk
    Exit Function
  End If

  VerificaNB = bOk

  Exit Function

ERRO_VERIFICANB:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Verificar NB.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function
Public Sub SetPosition(iLeft As Integer, iTop As Integer)

  Me.Left = iLeft
  Me.Top = iTop
End Sub
Public Sub SetParent(ByRef aForm As Form)

  Set mForm = aForm
End Sub
Function SalvaGPS() As Boolean

On Error GoTo ERRO_SALVAGPS

  Dim PossuiTipo    As Boolean
  Dim Confirma      As String
  Dim strEncripta   As String

  SalvaGPS = False

  Call CalculaValorTotal

  'Verificar se todos os campos estão preenchidos
  If CamposOK Then

    'Verifica Código de Pagto zerado
    If Val(TxtCodPagto.Text) = 0 Then
      MsgBox "Não é permitido código de pagamento zerado.", vbInformation + vbOKOnly, App.Title
      TxtCodPagto.SetFocus
      Exit Function
    End If

    'Valida Campo 'DataInicial'
    If Trim(Flag_CodigoPagto) <> "1" Then
      If Not VerificaDataInicialVigencia Then Exit Function
    End If

    'Valida Campo 'DataFinal'
    If Trim(Flag_CodigoPagto) <> "1" Then
      If Not VerificaDataFinalVigencia Then Exit Function
    End If

    'Valida Campo 'DigitaValorINSS'
    If Trim(Flag_CodigoPagto) <> "1" Then
      If Not DigitaValorINSS Then Exit Function
    End If

    'Valida Campo 'DigitaOutrasEntidades'
    If Trim(Flag_CodigoPagto) <> "1" Then
      If Not DigitaOutrasEntidades Then Exit Function
    End If

    'Valida Campo 'VerificaAtraso'
    If Trim(Flag_CodigoPagto) <> "1" Then
      If Not VerificaAtraso Then Exit Function
    End If

    'Valida Campo 'DataPagto13'
    If Trim(Flag_CodigoPagto) <> "1" Then
      If Not VerificaDataPagto13 Then Exit Function
    End If

    'Valida Campo 'CondPagtoNormal'
    If Trim(Flag_CodigoPagto) <> "1" Then
      If Not VerificaCondPagtoNormal Then Exit Function
    End If

    Confirma = VerificaConfirmacao

    'Verificar se o Documento pertence à outro Tipo
    If Geral.Documento.TipoDocto <> 35 And Geral.Documento.TipoDocto <> 0 Then
      With qryRemoveTipoDocumento
        .rdoParameters(1) = Geral.DataProcessamento     'Data Proc.
        .rdoParameters(2) = Geral.Documento.IdDocto     'IdDocto
        .rdoParameters(3) = Geral.Documento.TipoDocto   'Tipo do Documento
        .Execute
      End With
    End If

    'Atualiza campo Autenticação Digital
    strEncripta = G_EncriptaBO(35, CStr(Val(TxtIdentificador.Text)))
    If strEncripta = "" Then GoTo ERRO_SALVAGPS

    'Atualizar / Inserir GPS
    With qryAtualizaGPS
      .rdoParameters(0) = Geral.DataProcessamento            'Data Proc.
      .rdoParameters(1) = Geral.Documento.IdDocto            'IdDocto
      .rdoParameters(2) = TxtCodPagto.Text                   'Codigo Pagamento
      .rdoParameters(3) = Mid(TxtCompetencia.Text, 4, 4) & _
                          Mid(TxtCompetencia.Text, 1, 2)     'Competencia
      .rdoParameters(4) = TxtIdentificador.Text              'Identificador
      .rdoParameters(5) = Val(TxtValorINSS.Text) / 100       'Valor INSS
      .rdoParameters(6) = Val(TxtOutrasEnt.Text) / 100       'Valor Outras Entidades
      .rdoParameters(7) = Val(TxtJuros.Text) / 100           'Valor Juros / Atual. Monetária
      .rdoParameters(8) = Val(TxtTotal.Text) / 100           'Valor
      .rdoParameters(9) = Confirma                           'Confirmação
      .rdoParameters(10) = "35"                              'Tipo de Documento
      .rdoParameters(11) = strEncripta                       'Autenticacao digital
      .Execute
    End With

    'Atualizar o Controle Global
    Geral.Documento.ValorTotal = Val(TxtTotal.Text) / 100
    Geral.Documento.Leitura = ""
    Geral.Documento.TipoDocto = 35

    SalvaGPS = True
  End If

  Exit Function

ERRO_SALVAGPS:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Atualizar Dados do GPS.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function
Function CamposOK() As Boolean

    CamposOK = False
    
    'Código de Pagamento
    If Len(Trim(TxtCodPagto.Text)) = 0 Then
      MsgBox "Informe o Código de Pagamento.", vbInformation, App.Title
      TxtCodPagto.SetFocus
      Exit Function
    End If
    
    'Competencia
    If Len(Trim(TxtCompetencia.Text)) = 0 Then
      MsgBox "Informe o Mês e Ano de Competencia.", vbInformation, App.Title
      TxtCompetencia.SetFocus
      Exit Function
    End If

    'Verifica se Ano da Competência está dentro da faixa permitida para pagto.
    If Val(Mid(TxtCompetencia.Text, 4, 4)) < RegraValidaGPS(1).AnoInicial Or Val(Mid(TxtCompetencia.Text, 4, 4)) > RegraValidaGPS(1).AnoFinal Then
        MsgBox "Ano de competência fora da faixa permitida para pagamento.", vbInformation + vbOKOnly, App.Title
        TxtCompetencia.SetFocus
        Exit Function
    End If

    'Identificador
    If Len(Trim(TxtIdentificador.Text)) = 0 Then
      MsgBox "Informe o Identificador.", vbInformation, App.Title
      TxtIdentificador.SetFocus
      Exit Function
    End If
    
    'Valor Arrecadado
    If Len(Trim(TxtTotal.Text)) = 0 Then
      MsgBox "Valor Arrecadado não pode ser inferior a R$ " & Trim(FormataValor(RegraValidaGPS(1).ValorMinimoDocumento, 10)) & " .", vbInformation + vbOKOnly, App.Title
      TxtValorINSS.SetFocus
      Exit Function
    End If
    
    If CCur(TxtTotal.Text / 100) < RegraValidaGPS(1).ValorMinimoDocumento Then
      MsgBox "Valor Arrecadado não pode ser inferior a R$ " & Trim(FormataValor(RegraValidaGPS(1).ValorMinimoDocumento, 10)) & " .", vbInformation + vbOKOnly, App.Title
      TxtValorINSS.SetFocus
      Exit Function
    End If
    
    CamposOK = True

End Function
Sub PesquisaGPS()

  On Error GoTo ERRO_PesquisaGPS

  Dim sSql As String
  Dim RsGps As rdoResultset

  'Pesquisar a ADCC Atual e preencher os valores caso encontre
  sSql = Geral.DataProcessamento & " , " & Geral.Documento.IdDocto

  Set qryGetGps = Geral.Banco.CreateQuery("", "{call GetGPS (" & sSql & ")}")

  Set RsGps = qryGetGps.OpenResultset(rdOpenStatic, rdConcurReadOnly)
  If Not RsGps.EOF Then
    'Encontrou o Documento -> Preencher os campos
    TxtCodPagto.Text = RsGps!CodigoPagamento
    TxtCompetencia.Text = Mid(RsGps!Competencia, 5, 2) & "/" & Mid(RsGps!Competencia, 1, 4)
    TxtIdentificador.Text = RsGps!Identificador
    TxtValorINSS.Text = RsGps!ValorINSS * 100
    TxtOutrasEnt.Text = RsGps!ValorEntidades * 100
    TxtJuros.Text = RsGps!Juros * 100
    TxtTotal.Text = RsGps!Valor * 100

    'Pesquisar na tabela ValidaGPS
    If Not ValidaCodigoPagto(TxtCodPagto.Text) Then
      Flag_CodigoPagto = "1"
    Else
      Flag_CodigoPagto = "0"
    End If

    TxtValorINSS.SetFocus
    DoEvents
  Else
    'Não Encontrou o Documento -> Posicionar no primeiro campo
    TxtCodPagto.SetFocus
  End If

  If AlteraValor = True Then
    TxtCodPagto.Locked = True
    TxtCompetencia.Locked = True
    TxtIdentificador.Locked = True

    TxtValorINSS.SetFocus
  End If

  Screen.MousePointer = vbDefault

  Exit Sub

ERRO_PesquisaGPS:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Selecionar Dados do GPS.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Sub
Sub AjustesIniciais()

  'Setando as variáveis RDOQUERY
  Set qryAtualizaGPS = Geral.Banco.CreateQuery("", "{call AtualizaGPS (?,?,?,?,?,?,?,?,?,?,?,?)}")
  Set qryRemoveTipoDocumento = Geral.Banco.CreateQuery("", "{? = call RemoveTipoDocumento (?,?,?)}")
  Set qryGetGps = Geral.Banco.CreateQuery("", "{call GetGPS (?,?)}")
  Set qryGetValidaGPS = Geral.Banco.CreateQuery("", "{call GetValidaGPS (?)}")
  
End Sub

Private Sub Form_Activate()

    bActivate = True
    
    If Not CargaRegraValidaGPS Then
        Alterou = False
        Me.Hide
        Exit Sub
    End If
    
    Call AjustesIniciais
    
    Call PesquisaGPS
    
    Call CalculaValorTotal
    
    bActivate = False

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

  Set qryAtualizaGPS = Nothing
  Set qryRemoveTipoDocumento = Nothing
  Set qryGetGps = Nothing
  Set qryGetValidaGPS = Nothing
  
End Sub
Private Sub TxtCodPagto_Change()

  If Len(Trim(TxtCodPagto.Text)) = TxtCodPagto.MaxLength Then
    SendKeys "{TAB}"
    DoEvents
  End If
End Sub
Private Sub TxtCodPagto_GotFocus()

  TxtCodPagto.SelStart = 0
  TxtCodPagto.SelLength = TxtCodPagto.MaxLength
End Sub

Private Sub TxtCodPagto_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub

Private Sub TxtCodPagto_LostFocus()

  If Val(TxtCodPagto.Text) <> 0 And Not bActivate Then
    Call VerificaCodigoPagto
  End If
End Sub


Private Sub txtCompetencia_GotFocus()

  TxtCompetencia.SelStart = 0
  TxtCompetencia.SelLength = TxtCompetencia.MaxLength
End Sub


Private Sub txtCompetencia_KeyPress(KeyAscii As Integer)

  If InStr(TxtCompetencia.Text, "/") = 0 And Len(Trim(TxtCompetencia.Text)) = 6 Then
    'Verificar se a tecla atual é um numero
    If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 And TxtCompetencia.SelLength = 0 Then
      KeyAscii = 0
    End If
  End If

  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then 'And KeyAscii <> 47
    KeyAscii = 0
  End If
End Sub
Private Sub txtCompetencia_LostFocus()

  On Error GoTo ERRO_COMPETENCIALOSTFOCUS

  Dim Data As String
  Dim Pos As Integer
  Dim AnoComp As Integer

  Dim Mes As String
  Dim Ano As String

  'Marcar Flag de Erro
  Flag_LimiteInicialCompetencia = "0"
  Flag_IndicMesCompetencia = "0"

  If Len(Trim(TxtCompetencia.Text)) = 0 Then Exit Sub
    'Valida Competência
    
    If Mid(TxtCompetencia.Text, 3, 1) = "/" Then
        If Not IsNumeric(Mid(TxtCompetencia.Text, 1, 2)) And IsNumeric(Mid(TxtCompetencia.Text, 4, 2)) Then
            MsgBox "Competência incorreta, Redigite!", vbInformation, App.Title
            TxtCompetencia.SetFocus
            Exit Sub
        End If
    Else
        If Not IsNumeric(Mid(TxtCompetencia.Text, 1, 2)) And IsNumeric(Mid(TxtCompetencia.Text, 3, 2)) Then
            MsgBox "Competência incorreta, Redigite!", vbInformation, App.Title
            TxtCompetencia.SetFocus
            Exit Sub
        End If
    End If
    
  With TxtCompetencia
    .Text = Trim(.Text)
    If .Text <> "" Then
      Pos = InStr(1, .Text, "/", vbTextCompare)
      If Pos > 0 Then
        'O Usuário digitou barra
        Mes = Mid(.Text, 1, Pos - 1)
        Ano = Mid(.Text, Pos + 1)
        If Not ValidaData(Mes, Ano) Then Exit Sub

        Data = Mid(.Text, 1, Pos - 1) & Mid(.Text, Pos + 1, Len(.Text) - Pos)
      Else
        'O Usuário não informou Data
        If Len(.Text) <> 6 Then
          MsgBox "O Campo Competência deve estar no formato MM/AAAA.", vbInformation, App.Title
          .SelStart = 0
          .SelLength = Len(.Text)
          .SetFocus
          Exit Sub
        Else
          Data = .Text
          Mes = Mid(.Text, 1, 2)
          Ano = Mid(.Text, 3)

          If Not ValidaData(Mes, Ano) Then Exit Sub

        End If
      End If

      If (Not IsNumeric(Data)) Or Val(Data) < 0 Or Mid(Data, 3) > 2100 Or Mid(Data, 1, 2) > 13 Then
        MsgBox "Competência Inválida.", vbInformation, App.Title
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
        Exit Sub
      Else
        .Text = Format(Data, "00/0000")
      End If
    End If
  End With

  'Verificar Campo IndicMesCompetencia
  Select Case Pagto.IndicMesCompetencia
  Case 1
    '1 à 12
    If Mid(TxtCompetencia.Text, 1, 2) > 12 Then
      If MsgBox("Mês de Competência não permitido. Confirma ?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
        'Marcar flag
        Flag_IndicMesCompetencia = "1"
        TxtIdentificador.SetFocus
      Else
        TxtCompetencia.SetFocus
      End If
    End If
  Case 2
    '3,6,9,12
    If Mid(TxtCompetencia.Text, 1, 2) <> 3 And Mid(TxtCompetencia.Text, 1, 2) <> 6 And Mid(TxtCompetencia.Text, 1, 2) <> 9 And Mid(TxtCompetencia.Text, 1, 2) <> 12 Then
      If MsgBox("Mês de Competência não permitido. Confirma ?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
        'Marcar flag
        Flag_IndicMesCompetencia = "1"
        TxtIdentificador.SetFocus
      Else
        TxtCompetencia.SetFocus
      End If
    End If
  Case 3
    '3,6,9,12,13
    If Mid(TxtCompetencia.Text, 1, 2) <> 3 And Mid(TxtCompetencia.Text, 1, 2) <> 6 And Mid(TxtCompetencia.Text, 1, 2) <> 9 And Mid(TxtCompetencia.Text, 1, 2) <> 12 And Mid(TxtCompetencia.Text, 1, 2) <> 13 Then
      If MsgBox("Mês de Competência não permitido. Confirma ?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
        'Marcar flag
        Flag_IndicMesCompetencia = "1"
        TxtIdentificador.SetFocus
      Else
        TxtCompetencia.SetFocus
      End If
    End If
  Case 4
    '1 à 13 - Todos os meses são permitidos
  End Select

  'Verificar se o Ano de Competencia é válido
  If Pagto.LimiteInicialCompetencia > Val(Ano) Then
    If MsgBox("O Ano de competência não é válido. Confirma ?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
      'Marcar Flag de Erro
      Flag_LimiteInicialCompetencia = "1"
    Else
      TxtCompetencia.SetFocus
    End If
  End If

  'Verificar Indicador de Restrição de Pagamento,
  Select Case Pagto.IndicRestricaoPagto
    Case "1"
      '-----------------------Competencia pode ser antecipada
      'Verificar se o ano do movimento é diferente ao ano da competencia
      If Val(Mid(Geral.DataProcessamento, 1, 4)) <> Ano Then
        If MsgBox("Este Pagamento só pode ser antecipado no próprio exercício. Confirma ?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
          'Marcar Flag de Erro
          Flag_IndicRestricaoPagto = "1"
          Exit Sub
        Else
          TxtCompetencia.SetFocus
        End If
      Else
        'Verificar se o mes de competencia é maior que o mes do movimento
        If Mes < Val(Mid(Geral.DataProcessamento, 5, 2)) Then
          If MsgBox("Mês de Competência deve ser maior que o mes do Movimento. Confirma ?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
            'Marcar Flag de Erro
            Flag_IndicRestricaoPagto = "1"
            Exit Sub
          Else
            TxtCompetencia.SetFocus
          End If
        End If
      End If
    Case "2"
      '-------------Competencia 1 e 2 do proximo exercicio podem ser antecipadas no mes 11 ou 12
      'Verificar se o mes de competencia é igual à 1 ou 2
      If Val(Mes) = 1 Or Val(Mes) = 2 Then
        'Verificar se o Ano de Competencia - 1 é igual ao ano do movimento
        If Val(Ano - 1) = Val(Mid(Geral.DataProcessamento, 1, 4)) Then
          'Verificar se o mes do movimento é igual à 11 ou 12
          If Val(Mid(Geral.DataProcessamento, 5, 2)) <> 11 And Val(Mid(Geral.DataProcessamento, 5, 2)) <> 12 Then
            If MsgBox("Este Pagamento só pode ser antecipado nos meses 11 ou 12. Confirma ?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
              'Marcar Flag de Erro
              Flag_IndicRestricaoPagto = "1"
              Exit Sub
            Else
              TxtCompetencia.SetFocus
            End If
          End If
        End If
      Else
        'Verificar se o ano de competencia é diferente ao ano do movimento
        If Val(Ano) <> Val(Mid(Geral.DataProcessamento, 1, 4)) Then
          If MsgBox("Este Pagamento só pode ser antecipado para mês de competência 1 ou 2. Confirma ?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
            'Marcar flag de Erro
            Flag_IndicRestricaoPagto = "1"
            Exit Sub
          Else
            TxtCompetencia.SetFocus
          End If
        End If
      End If

    Case "3"
      '----------------------Somente competencia 13 pode ser antecipada
      'Verificar se a competencia é igual à 13
      If Val(Mes) = 13 Then
        'Verificar se o ano da competencia é diferente ao ano do movimento
        If Val(Ano) <> Val(Mid(Geral.DataProcessamento, 1, 4)) Then
          If MsgBox("Este Pagamento só pode ser antecipado no próprio exercício. Confirma ?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
            'Marcar flag de erro
            Flag_IndicRestricaoPagto = "1"
            Exit Sub
          Else
            TxtCompetencia.SetFocus
          End If
        End If
      End If

    Case "4"
      '----------------------Não pode antecipar Competencia
      'Verificar se ano de competencia é maior ano de movimento
      If Val(Ano) > Val(Mid(Geral.DataProcessamento, 1, 4)) Then
        If MsgBox("Este Pagamento não pode ser antecipado. Confirma ?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
          'Marcar Flag de Erro
          Flag_IndicRestricaoPagto = "1"
          Exit Sub
        Else
          TxtCompetencia.SetFocus
        End If
      ElseIf Val(Ano) = Val(Mid(Geral.DataProcessamento, 1, 4)) Then
        'Verificar o mes de competencia
        If Val(Mes) > Val(Mid(Geral.DataProcessamento, 5, 2)) Then
          If MsgBox("Este Pagamento não pode ser antecipado. Confirma ?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
            'Marcar Flag de Erro
            Flag_IndicRestricaoPagto = "1"
            Exit Sub
          Else
            TxtCompetencia.SetFocus
          End If
        End If
      End If
  End Select

  Exit Sub

ERRO_COMPETENCIALOSTFOCUS:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Validar Competência.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Sub

Private Sub txtIdentificador_GotFocus()
  TxtIdentificador.SelStart = 0
  TxtIdentificador.SelLength = TxtIdentificador.MaxLength
End Sub
Private Sub TxtIdentificador_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub
Private Sub TxtIdentificador_LostFocus()

  Dim PossuiTipo As Boolean
  Dim sMens As String

  If Len(Trim(TxtIdentificador.Text)) <> 0 Then
    'Valida CGC / Identificador
    If Not IsNumeric(TxtIdentificador.Text) Then
        MsgBox "Identificador incorreto, Redigite!", vbInformation, App.Title
        TxtIdentificador.SetFocus
        Exit Sub
    End If

    'Valida o Identificador para cada tipo de documento
    If Val(Pagto.TipoDocumento) = 0 Then
      PossuiTipo = False
    Else
      PossuiTipo = True
    End If

    If Not ValidaIdentificador(TxtIdentificador, PossuiTipo, sMens) Then
      If Len(Trim(sMens)) = 0 Then
        MsgBox "Código Identificador Inválido.", vbInformation, App.Title
      Else
        MsgBox sMens, vbInformation, App.Title
      End If
      TxtIdentificador.Text = ""
      TxtIdentificador.SetFocus
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

  Call CalculaValorTotal
End Sub
Private Sub TxtOutrasEnt_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub

Private Sub TxtOutrasEnt_LostFocus()

  Call CalculaValorTotal
End Sub


Private Sub TxtTotal_GotFocus()

  TxtTotal.SelStart = 0
  TxtTotal.SelLength = Len(TxtTotal.Text)
End Sub

Private Sub TxtTotal_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    Call cmdConfirmar_Click
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub


Private Sub TxtValorINSS_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
    KeyAscii = 0
  End If
End Sub

Private Sub TxtValorINSS_LostFocus()

  Call CalculaValorTotal
End Sub

Private Function CargaRegraValidaGPS() As Boolean

Dim RsValidaGPS As rdoResultset
Dim qryGetRegraValidaGPS As rdoQuery

On Error GoTo Err_CargaRegraValidaGPS
    
    Erase RegraValidaGPS

    CargaRegraValidaGPS = False
    
    Set qryGetRegraValidaGPS = Geral.Banco.CreateQuery("", "{call GetRegraValidaGPS }")
    
    Set RsValidaGPS = qryGetRegraValidaGPS.OpenResultset(rdOpenStatic, rdConcurReadOnly)
    
    If Not RsValidaGPS.EOF Then
        ReDim RegraValidaGPS(RsValidaGPS.RowCount)
        
        While Not RsValidaGPS.EOF
            RegraValidaGPS(RsValidaGPS.AbsolutePosition).AnoInicial = RsValidaGPS!AnoInicial
            RegraValidaGPS(RsValidaGPS.AbsolutePosition).AnoFinal = RsValidaGPS!AnoFinal
            RegraValidaGPS(RsValidaGPS.AbsolutePosition).ValorMinimoDocumento = RsValidaGPS!ValorMinimoDocumento
            
            RsValidaGPS.MoveNext
        Wend
        CargaRegraValidaGPS = True
    Else
        MsgBox "Não há parâmetros de Regras para validação de G.P.S." & vbCrLf & vbCrLf & "Favor contatar o suporte.", vbCritical + vbOKOnly, App.Title
    End If
    
Exit_CargaRegraValidaGPS:
    If Not (RsValidaGPS Is Nothing) Then RsValidaGPS.Close
    qryGetRegraValidaGPS.Close
    
    Exit Function

Err_CargaRegraValidaGPS:
    Screen.MousePointer = vbDefault
    Beep
    Select Case TratamentoErro("Erro na carga das Regras de GPS.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    GoTo Exit_CargaRegraValidaGPS

End Function
