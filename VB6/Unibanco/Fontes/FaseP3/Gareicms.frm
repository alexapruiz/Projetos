VERSION 5.00
Object = "{9CBA5D64-E3C8-11D3-9FFC-00104BC8688C}#1.0#0"; "CurrencyEdit.ocx"
Object = "{4A8E27C4-EACA-11D3-9FFC-00104BC8688C}#1.0#0"; "DateEdit.ocx"
Begin VB.Form GareICMS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GARE -  (Guia de Arrecadação Estadual )"
   ClientHeight    =   3360
   ClientLeft      =   12
   ClientTop       =   3660
   ClientWidth     =   12180
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3360
   ScaleWidth      =   12180
   Begin VB.CommandButton cmdFrenteVerso 
      Caption         =   "Frente/Verso"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   9408
      Picture         =   "Gareicms.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   0
      Width           =   852
   End
   Begin VB.CommandButton cmdInverteCor 
      Caption         =   "Inverter Cor"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   8544
      Picture         =   "Gareicms.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   0
      Width           =   850
   End
   Begin VB.CommandButton cmdRotacao 
      Caption         =   "Rotação"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   7680
      Picture         =   "Gareicms.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   0
      Width           =   850
   End
   Begin VB.CommandButton cmdZoomMais 
      Caption         =   "Zoom +"
      Height          =   700
      Left            =   5952
      Picture         =   "Gareicms.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   0
      Width           =   850
   End
   Begin VB.CommandButton cmdZoomMenos 
      Caption         =   "Zoom -"
      Height          =   700
      Left            =   6816
      Picture         =   "Gareicms.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   0
      Width           =   850
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   700
      Left            =   11136
      Picture         =   "Gareicms.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   0
      Width           =   850
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "Confirmar"
      Height          =   700
      Left            =   10272
      Picture         =   "Gareicms.frx":123C
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   0
      Width           =   850
   End
   Begin VB.Frame fraGare 
      Height          =   2700
      Left            =   48
      TabIndex        =   33
      Top             =   672
      Width           =   12144
      Begin CURRENCYEDITLib.CurrencyEdit txtValor 
         Height          =   360
         Left            =   10176
         TabIndex        =   18
         Top             =   1152
         Width           =   1788
         _Version        =   65537
         _ExtentX        =   3154
         _ExtentY        =   635
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
      Begin CURRENCYEDITLib.CurrencyEdit txtHonorarios 
         Height          =   360
         Left            =   10176
         TabIndex        =   17
         Top             =   672
         Width           =   1788
         _Version        =   65537
         _ExtentX        =   3154
         _ExtentY        =   635
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
      Begin CURRENCYEDITLib.CurrencyEdit txtAcrescimo 
         Height          =   360
         Left            =   10176
         TabIndex        =   16
         Top             =   192
         Width           =   1788
         _Version        =   65537
         _ExtentX        =   3154
         _ExtentY        =   635
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
      Begin CURRENCYEDITLib.CurrencyEdit txtMulta 
         Height          =   360
         Left            =   7056
         TabIndex        =   15
         Top             =   2112
         Width           =   1788
         _Version        =   65537
         _ExtentX        =   3154
         _ExtentY        =   635
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
      Begin CURRENCYEDITLib.CurrencyEdit TxtJuros 
         Height          =   360
         Left            =   7056
         TabIndex        =   14
         Top             =   1632
         Width           =   1788
         _Version        =   65537
         _ExtentX        =   3154
         _ExtentY        =   635
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
      Begin CURRENCYEDITLib.CurrencyEdit txtValorReceita 
         Height          =   360
         Left            =   7056
         TabIndex        =   13
         Top             =   1152
         Width           =   1788
         _Version        =   65537
         _ExtentX        =   3154
         _ExtentY        =   635
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
      Begin DATEEDITLib.DateEdit datVencimento 
         Height          =   360
         Left            =   2340
         TabIndex        =   35
         Top             =   192
         Width           =   1308
         _Version        =   65537
         _ExtentX        =   2307
         _ExtentY        =   635
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
      Begin VB.TextBox txtInscEstadual 
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
         Left            =   2328
         MaxLength       =   12
         TabIndex        =   5
         Top             =   1152
         Width           =   1740
      End
      Begin VB.TextBox txtReceita 
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
         Left            =   2328
         MaxLength       =   4
         TabIndex        =   3
         Top             =   672
         Width           =   660
      End
      Begin VB.TextBox txtDividaAtiva 
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
         Left            =   2328
         MaxLength       =   13
         TabIndex        =   9
         Top             =   2112
         Width           =   1884
      End
      Begin VB.TextBox txtCGCCPF 
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
         Left            =   2328
         MaxLength       =   15
         TabIndex        =   7
         Top             =   1632
         Width           =   2190
      End
      Begin VB.TextBox txtReferencia 
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
         Left            =   7068
         MaxLength       =   7
         TabIndex        =   11
         Top             =   192
         Width           =   960
      End
      Begin VB.TextBox txtAIIM 
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
         Left            =   7056
         MaxLength       =   12
         TabIndex        =   12
         Top             =   672
         Width           =   1785
      End
      Begin VB.Label lblDividaAtiva 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "06-Inscrição na Dívida Ativa ou Nº da Etiqueta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   456
         Left            =   240
         TabIndex        =   8
         Top             =   2112
         Width           =   2088
      End
      Begin VB.Label lblReceita 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "03-Código de Receita "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   240
         TabIndex        =   2
         Top             =   744
         Width           =   1872
      End
      Begin VB.Label lblVencimento 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "02-Data de Vencimento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   228
         Left            =   240
         TabIndex        =   0
         Top             =   192
         Width           =   2028
      End
      Begin VB.Label lblInscrEstadual 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "04-Inscrição Estadual "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   240
         TabIndex        =   4
         Top             =   1212
         Width           =   1848
      End
      Begin VB.Label lblCGCCPF 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "05-CGC ou CPF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   252
         TabIndex        =   6
         Top             =   1716
         Width           =   1284
      End
      Begin VB.Label lblVencimentoFormato 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "(dd/mm/aaaa)"
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   240
         TabIndex        =   1
         Top             =   384
         Width           =   1080
      End
      Begin VB.Label lblReferencia 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "07-Referência (MMAAAA)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   4656
         TabIndex        =   10
         Top             =   240
         Width           =   2112
      End
      Begin VB.Label lblMulta 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "11-Multa de Mora ou Multa por Infração (Nominal ou Corrigida)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   588
         Left            =   4656
         TabIndex        =   23
         Top             =   2016
         Width           =   2388
      End
      Begin VB.Label lblJuros 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "10-Juros de Mora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   4644
         TabIndex        =   22
         Top             =   1716
         Width           =   1452
      End
      Begin VB.Label lblValorReceita 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "09-Valor da Receita (Nominal ou Corrigida)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   408
         Left            =   4656
         TabIndex        =   21
         Top             =   1152
         Width           =   2280
      End
      Begin VB.Label lblAIIM 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "08-Nº AIIM ou Nº Deicmeme ou Nº Parcelamento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   408
         Left            =   4656
         TabIndex        =   20
         Top             =   672
         Width           =   2316
      End
      Begin VB.Label lblAcrescimo 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "12-Acréscimo Financeiro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   8940
         TabIndex        =   24
         Top             =   192
         Width           =   1212
      End
      Begin VB.Label lblValor 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "14-Valor Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   8928
         TabIndex        =   26
         Top             =   1284
         Width           =   1188
      End
      Begin VB.Label lblHonorarios 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "13-Honorários Advocatícios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   408
         Left            =   8940
         TabIndex        =   25
         Top             =   672
         Width           =   1236
      End
   End
   Begin VB.Image Image1 
      Height          =   384
      Left            =   336
      Picture         =   "Gareicms.frx":1546
      Top             =   192
      Width           =   384
   End
   Begin VB.Label lblInformativo 
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      Caption         =   "Digitação de GARE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   192
      Left            =   1200
      TabIndex        =   34
      Top             =   300
      Width           =   1632
   End
End
Attribute VB_Name = "GareICMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Variavel de retorno informando se Alterou ou Cancelou
Public Alterou As Boolean

Public Erro As Byte
Public GrupoReceita As String
Public AlteraValor As Boolean
Private mForm As Form
Dim bAlterar As Boolean
Dim bAlterou As Boolean

Dim sPosicaoErro As String


'''''''''''''''''''''''''''''''''''''''''''''''
' Funções Para Importaçao das Tabelas do Gare '
'''''''''''''''''''''''''''''''''''''''''''''''

'Declaração do Type
Private RegGareCDarf As TFSCDARF

'Declaração do Type do Registro Gare
Private Type TFSCDARF
  CodigoPagamento                   As String * 4
  CodigoGrupo                       As String * 1
  Data_Inicio_Vigencia              As String * 8
  Data_Final_Vigencia               As String * 8
  Indicador_Excessao                As String * 1
  Indicador_Arrecadacao             As String * 1
  Tipo_Servico                      As String * 1
  Indicador_Autenticacao            As String * 1
  Indicador_Servico_Autenticacao    As String * 3
  Numero_Vias_Comprovante           As String * 2
  Valor                             As String * 12
  CrLf                              As String * 1
  
End Type

'Declaração do Type
Private RegGareCRGPS As TFSCRGPS

'Declaração do Type do Registro Gare
Private Type TFSCRGPS
  CodigoGrupoReceita                As String * 1
  IndicadorCotaIPVA                 As String * 1
  IndicadorVenctoNormal             As String * 1
  IndicadorInscEstadual             As String * 1
  IndicadorCampoDocto               As String * 1
  IndicadorInscAtiva                As String * 1
  IndicadorReferencia               As String * 1
  IndicadorNumParcelamento          As String * 1
  IndicadorValorReceita             As String * 1
  IndicadorValorJuro                As String * 1
  IndicadorValorMulta               As String * 1
  IndicadorAcresFinanceiro          As String * 1
  IndicadorHonoAdvogado             As String * 1
  CrLf                              As String * 1
End Type


Private Type tpModulo
    qryInserirGARE          As rdoQuery
    qryConvenioGare         As rdoQuery
    qryGetValidaGare        As rdoQuery
    qryGetGare              As rdoQuery
    qryGetAtivaCampoGare    As rdoQuery
    qryGetTipoCampoGare     As rdoQuery
    qryRemoveTipoDocumento  As rdoQuery
    
    qryInserirTFSCDARF      As rdoQuery
    qryInserirTFSCRGPS      As rdoQuery


End Type

'''''''''''''''''''''''''''''''''''''''''''
' Registro GareICMS                       '
'''''''''''''''''''''''''''''''''''''''''''

'Declaração do Type
Private RegGare As TpGare

'Declaração do Type do Registro Gare
Private Type TpGare
  CodigoReceita As Long
  TpInscricao   As String * 1  ' 0-Não Obrigatório 1-I.E  2-COD MUNICIPAL 3-NUM DECLARADO
  TpDocto       As String * 1  ' 0-Não Obrigatório 1-CNPJ 2-CPF 3-CNPJ OU CPF OU RENAVAM  4-PLACA 5-RENAVAM
  InscAtiva     As String * 1  ' 0-Não Obrigatório 1-INSC DIV ATIVA 2-NRO ETIQUETA 3-INSC. DIV ATIV / NRO ETIQUET 4-FAIXA IPVA OBRIGATÓRIO 5-FAIXA IPVA OCASIONAL
  Referencia    As String * 1  ' 0-Não Obrigatório 1-Obrigatório
  TpAIIM        As String * 1  ' 0-Não Obrigatório 1-NR AIIM 2-DI 3-DSI 4-DI OU DSI 5-NR PARCELAMENTO 6-EXERCICIO 7-NR NOTIFICACAO
  ValorReceita  As String * 1  ' 0-Não Obrigatório 1-Obrigatório 9-Ocasional
  JurosMora     As String * 1  ' 0-Não Obrigatório 1-Obrigatório 9-Ocasional
  Multa         As String * 1  ' 0-Não Obrigatório 1-Obrigatório 9-Ocasional
  Acrescimo     As String * 1  ' 0-Não Obrigatório 1-Obrigatório 9-Ocasional
  Honorarios    As String * 1  ' 0-Não Obrigatório 1-Obrigatório 9-Ocasional
End Type


Private Modulo As tpModulo
Private Function CalculaValorTotal() As Boolean
'* Calcula Valor total ao confimar preenchimento de dados*'
txtValor.Text = Val(Desformata_Valor(txtValorReceita.Text)) + Val(Desformata_Valor(TxtJuros.Text)) + Val(Desformata_Valor(txtMulta.Text)) + Val(Desformata_Valor(txtAcrescimo.Text)) + Val(Desformata_Valor(txtHonorarios.Text))
End Function
Private Function VerificarTudo() As Boolean
    
'Dim iErroData   As Integer
Dim sGruposAtivos As String
Dim lValidaGrupo As Boolean

    VerificarTudo = False
    
    If VerificaReceita = False Then
        Exit Function
    End If
    
    'Consiste data de vencto somente para receita dos grupos "B" e "F"
    If txtReceita.Tag = "B" Or txtReceita.Tag = "F" Then
        'Verifica se Feriado / Agência aberta e Data Válida
        If Len(datVencimento.Text) = 8 Then
            If Not ValidaAgenciaPorDocto(Geral.Capa.AgOrig, datVencimento.Text, True, True) Then
                datVencimento.SetFocus
                Exit Function
            End If
'            iErroData = ValidaAgencia(Geral.Capa.AgOrig, datVencimento.Text, True)
'            If iErroData <> 0 Then
'                Select Case iErroData
'                    Case 2 'Feriado
'                        MsgBox "A Agência de Origem está em Feriado.", vbInformation, App.Title
'                    Case 3 'Agência Fechada
'                        MsgBox "A Agência de Origem está Fechada.", vbInformation, App.Title
'                    Case 1 'Documento Vencido
'                        MsgBox "A Data de Vencimento deve ser maior que a Data do Movimento Anterior.", vbInformation, App.Title
'                End Select
'                If iErroData = 1 Or iErroData = 2 Or iErroData = 3 Then
'                    datVencimento_GotFocus
'                    datVencimento.SetFocus
'                    Exit Function
'                End If
'            End If
        Else
            'Verifica se feriado na agência
            If Not ValidaAgenciaPorDocto(Geral.Capa.AgOrig, "", False) Then
                datVencimento.SetFocus
                Exit Function
            End If
            
'            iErroData = ValidaAgencia(Geral.Capa.AgOrig, "", False)
'            Select Case iErroData
'                Case 2 'Feriado
'                    MsgBox "A Agência de Origem está em Feriado.", vbInformation, App.Title
'                Case 3 'Agência Fechada
'                    MsgBox "A Agência de Origem está Fechada.", vbInformation, App.Title
'            End Select
'            If iErroData = 2 Or iErroData = 3 Then
'                datVencimento_GotFocus
'                datVencimento.SetFocus
'                Exit Function
'            End If
        End If
    End If
    
    'Valida Inscricao Estadual
    'If InStr("ABCDEFN", GrupoReceita) Then
    If VerificaInscEstadual = False Then
        Exit Function
    End If
    'End If

    'Valida CGC/CPF
    'If InStr("CDN", GrupoReceita) Then
    If VerificarCGCCPF = False Then
        Exit Function
    End If
    'End If
    
    'Valida Divida Ativa
    If VerificarDividaAtiva = False Then
        Exit Function
    End If
    
    'Valida AIIM
    If VerificarAIIM = False Then
        Exit Function
    End If
    
    'Valida Acrescimo
    
    ' Carrega os grupos Ativos para Valor_Acrescimo_Financ
    ' Os grupos Ativos são  marcados como '9'
    ' na tabela TFSCRGPS
    
    sGruposAtivos = AtivaCampos("Valor_Acrescimo_Financ", "9")
    lValidaGrupo = IIf(InStr(sGruposAtivos, GrupoReceita) = 0, False, True)
           
    'If (GrupoReceita <> "B") And (GrupoReceita <> "F") Then
    If Not lValidaGrupo Then
        If Val(txtAcrescimo.Text) > 0 Then
            MsgBox "Para o grupo da receita digitado, Acréscimo não pode estar preenchido.", vbInformation + vbOKOnly, App.Title
        End If
        txtAcrescimo.Text = 0
    End If
    
    ' Carrega os grupos Ativos para Valor_Honor_Advoc
    ' Os grupos Ativos são  marcados como '9'
    ' na tabela TFSCRGPS
    
    sGruposAtivos = AtivaCampos("Valor_Honor_Advoc", "9")
    lValidaGrupo = IIf(InStr(sGruposAtivos, GrupoReceita) = 0, False, True)
    
    ' If (GrupoReceita <> "E") And (GrupoReceita <> "F") And (GrupoReceita <> "G") And (GrupoReceita <> "I") And (GrupoReceita <> "J") Then
    
    If Not lValidaGrupo Then
        If Val(txtHonorarios.Text) > 0 Then
            MsgBox "Para o grupo da receita digitado, Honorários não pode estar preenchido.", vbInformation + vbOKOnly, App.Title
        End If
        txtHonorarios.Text = 0
    End If

    'Verificar se o campo REFERENCIA está preenchido corretamente
    If txtReferencia.Locked = False Then
        If Len(Trim(txtReferencia.Text)) < 6 Then
            MsgBox "Referência Inválida.", vbInformation + vbOKOnly, App.Title
            Exit Function
        Else
            If Not VerificarReferencia Then Exit Function
        End If
    End If
    
    'Valida Valor Receita
    If Val(txtValorReceita.Text) = 0 Then
        'alterado em 28/04/2000 a pedidos de Selma
        'MsgBox "É obrigatório informar o Valor da Receita.", 16, "GARE"
        
        'txtValorReceita = ""
        'txtValorReceita.SetFocus
        'Exit Function
    End If
        
    If Val(txtValor.Text) = 0 Then
        MsgBox "Digite o Valor da Receita!", vbInformation + vbOKOnly, App.Title
        txtValorReceita.SetFocus
        Exit Function
    Else
        If (Val(txtValor.Text)) <> (Val(txtValorReceita.Text) + Val(TxtJuros.Text) + Val(txtMulta.Text) + Val(txtAcrescimo.Text) + Val(txtHonorarios.Text)) Then
            MsgBox "O Valor Total não confere com soma das parcelas. Verifique e confirme!", vbInformation + vbOKOnly, App.Title
            txtValor.SetFocus
            Exit Function
        End If
    End If
    
    VerificarTudo = True
End Function
Private Function VerificarAIIM() As Boolean
    
    VerificarAIIM = False
    
    If VerificaReceita = False Then
        Exit Function
    End If
    
    
    ''''''''''''''''''''''''''''''''
    ' Verifica AIIM                '
    ''''''''''''''''''''''''''''''''
    
    If Not ValidaAIIM Then
        MsgBox "Campo AIIM Inválido. Verifique e retorne.", vbInformation + vbOKOnly, App.Title
        txtAIIM_GotFocus
        txtAIIM.SetFocus
        Exit Function
    End If
    
   
    If Erro = 9 Then  ' Campo AIIM é Inativo

        MsgBox "Para o grupo da receita digitado, o campo AIIM não pode estar preenchido.", vbInformation + vbOKOnly, App.Title
        txtAIIM = "0000000000"
        Exit Function
    ElseIf Erro = 1 Then
        MsgBox "Campo AIIM incompleto. Verifique e retorne.", vbInformation + vbOKOnly, App.Title
        txtAIIM_GotFocus
        txtAIIM.SetFocus
        Exit Function
    End If
   
    
    VerificarAIIM = True
    
End Function
Private Function VerificarReferencia() As Boolean
    
    Dim ano_ref             As Integer  ' Ano de referencia
    Dim mes_ref             As Integer  ' mes de referencia
    
    Dim ano_movto           As Integer  ' ano de movimento
    Dim mes_movto           As Integer  ' mes de movimento
    
    Dim DiaMesAno_ref       As String     ' Mes/Ano de referencia
    Dim DiaMesAno_movto     As String     ' Mes/Ano de movimento
    
      
    VerificarReferencia = False
    
    If VerificaReceita = False Then
        Exit Function
    End If

    If Mid(txtReferencia.Text, 3, 1) <> "/" Then
        MsgBox "Mês/Ano de Referência Inválido. Verifique e retorne.", vbInformation + vbOKOnly, App.Title
        txtReferencia.SetFocus
        Exit Function
    End If
    
    If Len(Trim(txtReferencia.Text)) = 7 Then
        
        mes_ref = Val(Mid(txtReferencia.Text, 1, 2))
        ano_ref = Val(Mid(txtReferencia.Text, 4, 4))
    
    End If
    
    
    mes_movto = Val(Mid(Geral.DataProcessamento, 5, 2))
    ano_movto = Val(Mid(Geral.DataProcessamento, 1, 4))
    
    
    If ano_ref < 1980 Then ' Ver se ano é maior que 1980
                MsgBox "Ano de Referência Inválido. Verifique e retorne. ", vbInformation + vbOKOnly, App.Title
                txtReferencia_GotFocus
                txtReferencia.SetFocus
                Exit Function
                
    ElseIf mes_ref < 1 Then ' Ver se mes é menor que 01
                MsgBox "Mês de Referência Inválido. Verifique e retorne. ", vbInformation + vbOKOnly, App.Title
                txtReferencia_GotFocus
                txtReferencia.SetFocus
                Exit Function
                 
    ElseIf mes_ref > 12 Then ' Ver se mes é maior que 12
                MsgBox "Mês de Referência Inválido. Verifique e retorne. ", vbInformation + vbOKOnly, App.Title
                txtReferencia_GotFocus
                txtReferencia.SetFocus
                Exit Function
    End If
    
    
   
    ' Ajusta o Ano/Mes de Referencia e Ano/Mes de Movto para Verificacoes
    DiaMesAno_ref = "01" & "/" & CStr(mes_ref) & "/" & CStr(ano_ref)
    DiaMesAno_movto = DataDD_MM_AAAA(Geral.DataProcessamento)
    
    
    If Abs(DateDiff("yyyy", DiaMesAno_movto, DiaMesAno_ref)) > 7 Then
    
        MsgBox "Ano de Referência Inválido. Verifique e retorne. ", vbInformation + vbOKOnly, App.Title
        txtReferencia_GotFocus
        txtReferencia.SetFocus
        Exit Function

    End If
    
    
    If Val(txtReceita.Text) = 462 Or Val(txtReceita.Text) = 1466 Or Val(txtReceita.Text) = 1545 Then
    
        'Limite Inferior ( Inscrição Petrobras ) Fica a mesma regra até segunda ordem
        If Trim(txtInscEstadual.Text) = "108119504115" And _
              CDate(DiaMesAno_ref) > (CDate(DiaMesAno_movto) - 30) Then
              MsgBox "Mês/Ano de Referência Inválido. Verifique e retorne. ", vbInformation + vbOKOnly, App.Title
              txtReferencia_GotFocus
              txtReferencia.SetFocus
              Exit Function
        
        Else
            ' Limite Inferior
            If CDate(DiaMesAno_ref) > (CDate(DiaMesAno_movto) - 30) Then
                MsgBox "Mês/Ano de Referência Inválido. Verifique e retorne. ", vbInformation + vbOKOnly, App.Title
                txtReferencia_GotFocus
                txtReferencia.SetFocus
                Exit Function
            End If
            
        End If
        
        
    End If
    
    
    If Val(txtReceita.Text) = 607 Then
      
            ' Limite Inferior
            If CDate(DiaMesAno_ref) > (CDate(DiaMesAno_movto) + 180) Then
                MsgBox "Mês/Ano de Referência Inválido. Verifique e retorne. ", vbInformation + vbOKOnly, App.Title
                txtReferencia_GotFocus
                txtReferencia.SetFocus
                Exit Function
            End If
        
    End If
    
    If Val(txtReceita.Text) = 360 Then
    
        If InStr("1234", txtReferencia.Text) = 0 Or (Len(txtReferencia.Text) > 1) Then
            MsgBox "Cota do IPVA Inválida! Verifique e retorne. ", vbInformation + vbOKOnly, App.Title
            txtReferencia_GotFocus
            txtReferencia.SetFocus
            Exit Function
        End If
    
    End If
    
    
    VerificarReferencia = True
   
End Function

Private Function VerificarDividaAtiva() As Boolean
    VerificarDividaAtiva = False
        
    ''''''''''''''''''''''''''''''
    ' Verifica se digito confere '
    ''''''''''''''''''''''''''''''
    
    If Not ValidaInscDA Then
        MsgBox "Digito Divida Ativa não confere. Verifique e retorne.", vbInformation + vbOKOnly, App.Title
        txtDividaAtiva_GotFocus
        txtDividaAtiva.SetFocus
        Exit Function
    Else
        If Erro = 1 Then
            MsgBox "Campo Divida Ativa incompleto. Verifique e retorne.", vbInformation + vbOKOnly, App.Title
            txtDividaAtiva_GotFocus
            txtDividaAtiva.SetFocus
            Exit Function
        Else
            If Erro = 2 Then
                MsgBox "Código Divida Ativa Inválido. Verifique e retorne.", vbInformation + vbOKOnly, App.Title
                txtDividaAtiva_GotFocus
                txtDividaAtiva.SetFocus
                Exit Function
            End If
        End If
        VerificarDividaAtiva = True
    End If

End Function

Private Function VerificarCGCCPF() As Boolean
'Dim sGruposAtivos As String
'Dim sGruposInativos As String
'Dim sGruposDependentes As String
Dim sTipoCampo As String
'Dim lValidaGrupo As Boolean

    
    VerificarCGCCPF = False     'esta com erro
        
    If VerificaReceita = False Then
        Exit Function
    End If
        
        
    ' Verifica o Tipo do Campo
    sTipoCampo = TipoCampoGare("Campo_Docto", GrupoReceita)
        
        
        
    Select Case sTipoCampo
    
        Case 0
            ' Campo não é criticado
            VerificarCGCCPF = True
            Exit Function
        Case 1
            ' CNPJ Obrigatório
            If Val(txtCGCCPF) = 0 Then
                MsgBox "O preenchimento do CNPJ é obrigatório para o Código de Receita digitado.", vbInformation + vbOKOnly, App.Title
                txtCGCCPF_GotFocus
                txtCGCCPF.SetFocus
                Exit Function
            End If
        
        Case 2
            ' CPF Obrigatório
            If Val(txtCGCCPF) = 0 Then
                MsgBox "O preenchimento do CPF é obrigatório para o Código de Receita digitado.", vbInformation + vbOKOnly, App.Title
                txtCGCCPF_GotFocus
                txtCGCCPF.SetFocus
                Exit Function
            End If
        
        Case 3
            ' CNPJ ou CPF  (ou RENAVAM Obrigatório) O Zé Pereira Mandou Retirar a Critica do Renavam
            If Len(txtCGCCPF) > 11 Then
                txtCGCCPF = Format(txtCGCCPF, "000000000000000")    '15 posicoes
                If Not VerificaCGC(txtCGCCPF) Then
                    MsgBox "Digito do campo CNPJ/CPF não confere. Tente novamente.", vbInformation + vbOKOnly, App.Title
                    txtCGCCPF_GotFocus
                    txtCGCCPF.SetFocus
                    Exit Function
                Else
                    VerificarCGCCPF = True
                    Exit Function
                End If
            ElseIf Len(txtCGCCPF) = 11 Then
                txtCGCCPF = Format(txtCGCCPF, "00000000000")        '11 posicoes
                If Not VerificaCPF(txtCGCCPF) Then
                    MsgBox "Digito do campo CNPJ/CPF não confere. Tente novamente.", vbInformation + vbOKOnly, App.Title
                    txtCGCCPF_GotFocus
                    txtCGCCPF.SetFocus
                    Exit Function
                Else
                    VerificarCGCCPF = True
                    Exit Function
                End If
                
                
            Else
             
                MsgBox "Digito do campo CNPJ/CPF não confere. Tente novamente.", vbInformation + vbOKOnly, App.Title
                txtCGCCPF_GotFocus
                txtCGCCPF.SetFocus
                Exit Function
             
'            ElseIf Len(txtCGCCPF) = 9 Then
'                ' Critica o Renavam
'                If Not Val_Renavam(txtCGCCPF) Then
'                    MsgBox "Digito do Renavam não confere. Tente novamente.", vbInformation + vbOKOnly, App.Title
'                    txtCGCCPF_GotFocus
'                    txtCGCCPF.SetFocus
'                    Exit Function
'                Else
'                    VerificarCGCCPF = True
'                    Exit Function
'                End If
                
            End If
        
        Case 4
            ' Placa Obrigatório
            VerificarCGCCPF = True
            Exit Function
        
        Case 5
            ' Renavam Obrigatório
            If Not Val_Renavam(txtCGCCPF) Then
                MsgBox "Digito do Renavam não confere. Tente novamente.", vbInformation + vbOKOnly, App.Title
                txtCGCCPF_GotFocus
                txtCGCCPF.SetFocus
                Exit Function
            Else
                VerificarCGCCPF = True
                Exit Function
            End If
        
        Case 6
            'CNPJ ou CPF dependente (*) para os grupos marcados com
            ' 6 (I.E) ou (CNPJ-CPF) devem estar Preenchidos
            ' Se for dependente e InscEstadual estiver preenchido não critica o Campo
            
            
            If Val(txtInscEstadual) > 0 Then
            
                If Val(txtCGCCPF) > 0 Then
                    MsgBox "Para o código de receita digitado o campo Inscrição Estadual ou CGC/CPF deve estar preenchido.", vbInformation + vbOKOnly, App.Title
                    txtInscEstadual_GotFocus
                    txtInscEstadual.SetFocus
                    Exit Function
                Else
                    VerificarCGCCPF = True
                    Exit Function
                End If
            
            End If
            
            ' CNPJ ou CPF deve estar preenchido
            If Val(txtInscEstadual) = 0 And Val(txtCGCCPF) = 0 Then
                MsgBox "Para o código de receita digitado o campo Inscrição Estadual ou CGC/CPF deve estar preenchido.", vbInformation + vbOKOnly, App.Title
                txtInscEstadual_GotFocus
                txtInscEstadual.SetFocus
                Exit Function
            End If
            
            ' CNPJ ou CPF deve estar preenchido
            If Val(txtInscEstadual) = 0 And Val(txtCGCCPF) > 0 Then
            
                ' CNPJ ou CPF
                If Len(txtCGCCPF) > 11 Then
                    txtCGCCPF = Format(txtCGCCPF, "000000000000000")    '15 posicoes
                    If Not VerificaCGC(txtCGCCPF) Then
                        MsgBox "Digito do campo CNPJ/CPF não confere. Tente novamente.", vbInformation + vbOKOnly, App.Title
                        txtCGCCPF_GotFocus
                        txtCGCCPF.SetFocus
                        Exit Function
                    Else
                        VerificarCGCCPF = True
                        Exit Function
                    End If
                ElseIf Len(txtCGCCPF) = 11 Then
                    txtCGCCPF = Format(txtCGCCPF, "00000000000")        '11 posicoes
                    If Not VerificaCPF(txtCGCCPF) Then
                        MsgBox "Digito do campo CNPJ/CPF não confere. Tente novamente.", vbInformation + vbOKOnly, App.Title
                        txtCGCCPF_GotFocus
                        txtCGCCPF.SetFocus
                        Exit Function
                    Else
                        VerificarCGCCPF = True
                        Exit Function
                    End If
            
                Else
                    MsgBox "Digito do campo CNPJ/CPF não confere. Tente novamente.", vbInformation + vbOKOnly, App.Title
                    txtCGCCPF_GotFocus
                    txtCGCCPF.SetFocus
                    Exit Function
                
                End If
        
               
            End If
        
    
    End Select
    

End Function

Private Function VerificaInscEstadual() As Boolean

Dim sTipoCampo As String
   
    
    VerificaInscEstadual = False
        
    If Len(txtInscEstadual) = 0 Then
        txtInscEstadual = ""
    End If
    
        
        
        
    ' Inicia Validação da Inscrição estadual
    ' Ver Tipo do Campo do Saber se o Tipo do campo Dependente
    sTipoCampo = TipoCampoGare("Inscricao_Estadual", GrupoReceita)
    
    Select Case sTipoCampo
    
        Case 0
            ' Campo não está habilitado
            VerificaInscEstadual = True
            Exit Function
        Case 1
            ' Campo é obrigatório
            If Val(txtInscEstadual) = 0 Then
                MsgBox "Para o grupo da receita digitado, este campo  deve estar preenchido.", vbInformation + vbOKOnly, App.Title
                txtInscEstadual_GotFocus
                txtInscEstadual.SetFocus
                VerificaInscEstadual = False
                Exit Function
            ElseIf Len(txtInscEstadual.Text) <> 12 Then
                MsgBox "Inscrição Estadual inválida.", vbExclamation
                SelecionarTexto txtInscEstadual
                VerificaInscEstadual = False
                Exit Function
            End If
        Case 4
            ' Inscrição (Dependente): Inscrição ou CPF/CNPJ deve ser informado e
            ' será criticado após diditação do CPF/CNPJ
            
            If Val(txtCGCCPF) > 0 Then
                If Val(txtInscEstadual) > 0 Then
                    MsgBox "Para o grupo da receita digitado, o campo Inscrição Estadual não pode estar preenchido, pois já existe valor no campo CPF/CGC.", vbInformation + vbOKOnly, App.Title
                    VerificaInscEstadual = False
                    Exit Function
                End If
                VerificaInscEstadual = True
                Exit Function
            ElseIf (Val(txtCGCCPF) = 0) And (Val(txtInscEstadual) = 0) Then
                txtInscEstadual = ""
                VerificaInscEstadual = True
                Exit Function
            End If
    
    End Select
    
      
   
    If Not ValidaInscricao(sTipoCampo) Then
       
        If Erro = 1 Or Erro = 4 Then
            MsgBox "Campo Inscrição Estadual incompleto. Verifique e retorne.", vbInformation + vbOKOnly, App.Title
            txtInscEstadual_GotFocus
            txtInscEstadual.SetFocus
            Exit Function
        ElseIf Erro = 2 Then
            MsgBox "Código do Município Inválido! Verifique e retorne.", vbInformation + vbOKOnly, App.Title
            txtInscEstadual_GotFocus
            txtInscEstadual.SetFocus
            Exit Function
        ElseIf Erro = 3 Then
            MsgBox "Este número de Inscrição não é válido.", vbInformation + vbOKOnly, App.Title
            txtInscEstadual_GotFocus
            txtInscEstadual.SetFocus
            Exit Function
        ElseIf Erro = 5 Then
            MsgBox "Número da Declaração Inválido! Verifique e retorne.", vbInformation + vbOKOnly, App.Title
            txtInscEstadual_GotFocus
            txtInscEstadual.SetFocus
            Exit Function
        End If
   
   End If
        
   VerificaInscEstadual = True

End Function

'Private Function Old_VerificaInscEstadual() As Boolean
'    VerificaInscEstadual = False
'
'    If Len(txtInscEstadual) = 0 Then
'        txtInscEstadual = ""
'    End If
'
'    Select Case GrupoReceita
'
'        Case "A", "B", "E", "F", "H", "J"
'            If Val(txtInscEstadual) = 0 Then
'                MsgBox "Para o grupo da receita digitado, o campo Inscrição Estadual deve estar preenchido.", vbInformation + vbOKOnly, App.Title
'                txtInscEstadual_GotFocus
'                txtInscEstadual.SetFocus
'                VerificaInscEstadual = False
'                Exit Function
'            End If
'            If Len(txtInscEstadual.Text) <> 12 Then
'                MsgBox "Inscrição Estadual inválida.", vbExclamation
'                SelecionarTexto txtInscEstadual
'                VerificaInscEstadual = False
'                Exit Function
'            End If
'
'
'
'        Case "G", "I"
'            If Val(txtInscEstadual) > 0 Then
'                MsgBox "Para o grupo da receita digitado, Inscrição Estadual não pode estar preenchido.", vbInformation + vbOKOnly, App.Title
'            End If
'            txtInscEstadual = ""
'            VerificaInscEstadual = True
'            Exit Function
'        Case "C", "D", "N"
'            If Val(txtCGCCPF) > 0 Then
'                If Val(txtInscEstadual) > 0 Then
'                    MsgBox "Para o grupo da receita digitado, o campo Inscrição Estadual não pode estar preenchido, pois já existe valor no campo CPF/CGC.", vbInformation + vbOKOnly, App.Title
'                End If
'                txtInscEstadual = ""
'                VerificaInscEstadual = True
'                Exit Function
'            ElseIf (Val(txtCGCCPF) = 0) And (Val(txtInscEstadual) = 0) Then
'                txtInscEstadual = ""
'                VerificaInscEstadual = True
'                Exit Function
'            End If
'    End Select
'
'
'
'    If Not ValidaInscricao Then
'        MsgBox "Digito Inscrição Estadual não confere. Verifique e retorne.", vbInformation + vbOKOnly, App.Title
'        txtInscEstadual_GotFocus
'        txtInscEstadual.SetFocus
'        Exit Function
'    Else
'        If Erro = 1 Or Erro = 4 Then
'            MsgBox "Campo Inscrição Estadual incompleto. Verifique e retorne.", vbInformation + vbOKOnly, App.Title
'            txtInscEstadual_GotFocus
'            txtInscEstadual.SetFocus
'            Exit Function
'        ElseIf Erro = 2 Then
'            MsgBox "Código do Município Inválido! Verifique e retorne.", vbInformation + vbOKOnly, App.Title
'            txtInscEstadual_GotFocus
'            txtInscEstadual.SetFocus
'            Exit Function
'        ElseIf Erro = 3 Then
'            MsgBox "Este número de inscrição não é válido.", vbInformation + vbOKOnly, App.Title
'            txtInscEstadual_GotFocus
'            txtInscEstadual.SetFocus
'            Exit Function
'        End If
'    End If
'
'    VerificaInscEstadual = True
'
'End Function

Private Function VerificaReceita() As Boolean
    Dim tb As rdoResultset
    
    VerificaReceita = False
'    if Val(txtReceita) =
'    If Not ValidaData Then
'        Exit Function
'    End If
    
    If Len(txtReceita) = 0 Then
        MsgBox "Informe o código da receita.", vbInformation + vbOKOnly, App.Title
        txtReceita_GotFocus
        txtReceita.SetFocus
        Exit Function
    Else
        'Leda deixou regra fixa, por não saber se estes códigos constam ou não na
        'tabela de Gare. (23/03/2000)
        'Não são aceitos porque se refere a IPVA
        If Val(txtReceita) = 360 Or Val(txtReceita) = 358 Then
            MsgBox "Este código da Receita não é aceito pelo Unibanco.", vbCritical + vbOKOnly, App.Title
            txtReceita_GotFocus
            txtReceita.SetFocus
            Exit Function
        End If
        
        ' Validar código da receita digitado
        With Modulo.qryGetValidaGare
            .rdoParameters(0).Value = Val(txtReceita.Text)     'código da receita digitado
            Set tb = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
        End With
    
        If tb.EOF Then
            MsgBox "Código de receita inválido. Tente novamente.", vbCritical + vbOKOnly, App.Title
            txtReceita_GotFocus
            txtReceita.SetFocus
            Exit Function
        Else
            GrupoReceita = tb!CodigoGrupo
            MontaTela (tb!CodigoGrupo)
        End If

        'Valida Data de Vencimento somente para Código de Receita do Grupos (B) e (F) a (P)
        If tb!CodigoGrupo = "B" Or tb!CodigoGrupo = "F" Or tb!CodigoGrupo = "P" Then
            If Not ValidaData Then
                tb.Close
                Exit Function
            End If
        End If
        
        'Guarda no Tag a referencia
        txtReceita.Tag = tb!CodigoGrupo
        
        tb.Close
    End If
    
    VerificaReceita = True
                
End Function

Public Function ValidaData() As Boolean
    Dim sAno As String
    Dim iDia As String
    Dim iMes As String
    Dim iAno As String
    Dim iUltimoDia As String
    
    ValidaData = False
    
    If Not (txtReceita.Tag = "B" Or txtReceita.Tag = "F" Or txtReceita.Tag = "P") Then
        ValidaData = True
        Exit Function
    End If
    
    If Len(datVencimento.Text) < 8 Then
        MsgBox "Digite a data de vencimento no formato (dd/mm/aaaa)!", vbInformation + vbOKOnly, App.Title
        datVencimento.Text = ""
        datVencimento_GotFocus
        datVencimento.SetFocus
        Exit Function
    Else
        'alteração versão 3.3
        sAno = Right(datVencimento.Text, 4)
        If (sAno < 1950) Then
            MsgBox "O ano não pode ser menor do que 1950.", vbInformation + vbOKOnly, App.Title
            datVencimento.Text = ""
            datVencimento_GotFocus
            datVencimento.SetFocus
            Exit Function
        ElseIf (sAno > 2051) Then
            MsgBox "O ano não pode ser maior do que 2051.", vbInformation + vbOKOnly, App.Title
            datVencimento.Text = ""
            datVencimento_GotFocus
            datVencimento.SetFocus
            Exit Function
        End If
    End If
    ''''''''''''''''''''''''''''''''''
    ' Verifica se vencimento está OK '
    ''''''''''''''''''''''''''''''''''
    iDia = Left(datVencimento.Text, 2)
    iMes = Mid(datVencimento.Text, 3, 2)
    iAno = Right(datVencimento.Text, 4)

    Select Case Val(iMes)
        Case 1, 3, 5, 7, 8, 10, 12      ' 31 dias
            iUltimoDia = "31"
        Case 2                          ' 28/29 dias
            If Val(iAno) Mod 4 = 0 Then ' ano é bissexto
                iUltimoDia = "29"
            Else
                iUltimoDia = "28"
            End If
        Case 4, 6, 9, 11                ' 30 dias
            iUltimoDia = "30"
        Case Else
            MsgBox "Vencimento inválido. Digite novamente.", vbInformation + vbOKOnly, App.Title
            datVencimento.Text = ""
            datVencimento_GotFocus
            datVencimento.SetFocus
            Exit Function
    End Select
    
    If Val(iDia) < 1 Or Val(iDia) > Val(iUltimoDia) Then
        MsgBox "Vencimento inválido. Digite novamente.", vbInformation + vbOKOnly, App.Title
        datVencimento.Text = ""
        datVencimento_GotFocus
        datVencimento.SetFocus
        Exit Function
    End If
   
    ValidaData = True

End Function
Public Function ValidaAIIM() As Boolean
   
    Dim sTipoCampo As String   ' Tipo do Campo para formatacao
      
    Dim soma As Integer
    Dim resto As Integer
    Dim digito_11 As Integer
    Dim digito_rv As String
    Dim bOk As Byte
    
    bOk = True          'default - ok
   
    Erro = 0
    soma = 0
    resto = 0
    digito_11 = 0       'calculado pelo módulo 11
    digito_rv = ""      'caracter digitado pelo operador
      
    ' Carrega o tipo do campos Numero_Parcelamento
    
    sTipoCampo = TipoCampoGare("Numero_Parcelamento", GrupoReceita)
    
    Select Case sTipoCampo
     
        Case 0
            If Val(txtAIIM) <> 0 Then
                ' Número não deve estar preenchido
                Erro = 9
            End If
            Erro = 0
            bOk = True
            ValidaAIIM = bOk
            Exit Function
        
        Case 1
            
            If (GrupoReceita = "C") Then ' Grupo Eletrônico ou Manual
                
                txtAIIM = Format(txtAIIM, "000000000")
                
                If (Mid(Trim(txtAIIM), 1, 1) <> 0) Or (Mid(Trim(txtAIIM), 2, 1) <> 0) Then
                        
                    soma = soma + Mid(txtAIIM, 1, 1) * 9
                    soma = soma + Mid(txtAIIM, 2, 1) * 8
                    soma = soma + Mid(txtAIIM, 3, 1) * 7
                    soma = soma + Mid(txtAIIM, 4, 1) * 6
                    soma = soma + Mid(txtAIIM, 5, 1) * 5
                    soma = soma + Mid(txtAIIM, 6, 1) * 4
                    soma = soma + Mid(txtAIIM, 7, 1) * 3
                    soma = soma + Mid(txtAIIM, 8, 1) * 2
            
                    resto = soma Mod 11                'resto da divisão
                    digito_11 = Right(str(resto), 1)   'digito verificador
                    digito_rv = Mid(txtAIIM, 9, 1)     'digito verificador
                
                ElseIf (Mid(Trim(txtAIIM), 1, 1) = 0) And (Mid(Trim(txtAIIM), 2, 1) = 0) Then
                
                   ' O número do AIIM deve ser numerico <> de Zero
                   If Val(txtAIIM) = 0 Then ' Número não deve estar preenchido
                        Erro = 1
                        bOk = False
                        ValidaAIIM = bOk
                        Exit Function
                    Else
                        bOk = True
                        ValidaAIIM = bOk
                        Exit Function
                    End If
                
                End If
        
            End If
            
            
            If (GrupoReceita = "H") Or (GrupoReceita = "Q") Then  'AIIM GrupoReceita = "H" ou "Q"
                  
                  
                  ' Se GrupoReceita = "H" ou "Q"
                  
                  '''''''''''''''''''''
                  ' AIIM (eletronico) '
                  '''''''''''''''''''''
                  'NNNNNND
                  
                  If Len(txtAIIM) <> 7 Then
                      Erro = 1
                      ValidaAIIM = bOk
                      Exit Function
                  End If
                  
                  txtAIIM = Format(txtAIIM, "0000000")
            
                  soma = soma + Mid(txtAIIM, 1, 1) * 7
                  soma = soma + Mid(txtAIIM, 2, 1) * 6
                  soma = soma + Mid(txtAIIM, 3, 1) * 5
                  soma = soma + Mid(txtAIIM, 4, 1) * 4
                  soma = soma + Mid(txtAIIM, 5, 1) * 3
                  soma = soma + Mid(txtAIIM, 6, 1) * 2
            
                  resto = soma Mod 11                 'resto da divisão
                  digito_11 = Right(str(resto), 1)    'digito verificador
                  digito_rv = Mid(txtAIIM, 7, 1)      'digito verificador
            
            End If
            
            
            If (GrupoReceita = "M") Then  'AIIM GrupoReceita = "M"
            
                ' O número do AIIM deve ser numerico <> de Zero
                If Val(txtAIIM) = 0 Then ' Número não deve estar preenchido
                     Erro = 1
                     bOk = False
                     ValidaAIIM = bOk
                     Exit Function
                Else
                     bOk = True
                     ValidaAIIM = bOk
                     Exit Function
                End If
            
            
            End If
            
    
        Case 2
            
            ' Verifica de o Numero é  DI AANNNNNNND
            If Not Valida_DI(txtAIIM) Then
                Erro = 1
                bOk = False
            Else
                Erro = 0
                bOk = True
            End If
            
            ValidaAIIM = bOk
            Exit Function
        
        Case 3
        
            ' Verifica de o Numero é  DSI AANNNNNNND
            If Not Valida_DSI(txtAIIM) Then
                Erro = 1
                bOk = False
            Else
                Erro = 0
                bOk = True
            End If
            
            ValidaAIIM = bOk
            Exit Function
            
        Case 4
        
            ' Verifica de o Numero é DI ou DSI
            If Not Valida_DSI(txtAIIM) And Not Valida_DI(txtAIIM) Then
                Erro = 1
                bOk = False
            Else
                Erro = 0
                bOk = True
            End If
            
            ValidaAIIM = bOk
            Exit Function
            
        Case 5
            ' Se GrupoReceita = "B" Or GrupoReceita = "F" 'parcelamento NNNNNNNND
           
            
            txtAIIM = Format(txtAIIM, "000000000")
            
'            If Len(txtAIIM) <> 9 Then
'                Erro = 1
'                bOk = False
'                ValidaAIIM = bOk
'                Exit Function
'            End If
      
            soma = soma + Mid(txtAIIM, 1, 1) * 8
            soma = soma + Mid(txtAIIM, 2, 1) * 7
            soma = soma + Mid(txtAIIM, 3, 1) * 6
            soma = soma + Mid(txtAIIM, 4, 1) * 5
            soma = soma + Mid(txtAIIM, 5, 1) * 4
            soma = soma + Mid(txtAIIM, 6, 1) * 3
            soma = soma + Mid(txtAIIM, 7, 1) * 2
            soma = soma + Mid(txtAIIM, 8, 1) * 10
      
            resto = soma Mod 11        'resto da divisão
            digito_11 = 11 - resto     'digito verificador
      
            '*** se o calculo for igual a 10 ou 11, muda-se para 0 ***
            If (digito_11 = 10) Then
                digito_11 = 0
            End If
      
            If (digito_11 = 11) Then
                bOk = False     'digito não confere
            End If
   
            digito_rv = Mid(txtAIIM, 9, 1)    'digito verificador
            
        Case 6
            Erro = 0
            bOk = True
            ValidaAIIM = bOk
            Exit Function
        
        Case 7
            ' Verifica de o Numero é  DSI AANNNNNNND
            If Not Val_Notificacao(txtAIIM) Then
                Erro = 1
                bOk = False
            Else
                Erro = 0
                bOk = True
            End If
            
            ValidaAIIM = bOk
            Exit Function
        
        Case 8
            
            ' Verifica de o Numero é  DSI AANNNNNNND
            If Val(txtAIIM) = 0 Then ' Número é Ocasional
                Erro = 0
                bOk = True
                ValidaAIIM = bOk
                Exit Function
            End If
            
            ' Verifica de o Numero é  DSI AANNNNNNND
            If Not Val_Notificacao(txtAIIM) Then
                Erro = 1
                bOk = False
            Else
                Erro = 0
                bOk = True
            End If
            
            ValidaAIIM = bOk
            Exit Function
     
    End Select
    
         
    If CStr(digito_11) <> (digito_rv) Then
        bOk = False                  'digito não confere
        Erro = 1
    End If
    
    ValidaAIIM = bOk
   
End Function
Public Function ValidaInscDA() As Boolean
    
    Dim sGruposAtivos As String
    Dim sGruposInativos As String
    Dim sGruposDependentes As String
    Dim lValidaGrupo As Boolean
    
    Dim soma As Integer
    Dim resto As Integer
    Dim digito_11 As Integer
    Dim digito_rv As String
    Dim bOk As Byte
   
    bOk = True          'default - ok
   
    Erro = 0
    soma = 0
    resto = 0
    digito_11 = 0       'calculado pelo módulo 11
    digito_rv = ""      'caracter digitado pelo operador
      
    
    
    ' Carrega os grupos ativos para Inscricao_Ativa
    ' Os grupos ativos são  marcados como '3'
    ' na tabela TFSCRGPS
        
    sGruposAtivos = AtivaCampos("Inscricao_Ativa", "3")
    lValidaGrupo = IIf(InStr(sGruposAtivos, GrupoReceita) = 0, False, True)

    
    If Len(txtDividaAtiva) = 0 Then
        'If (GrupoReceita = "E") Or (GrupoReceita = "I") Or (GrupoReceita = "J") Then
        If lValidaGrupo Then
            Erro = 1
        End If
        txtDividaAtiva = "0000000000000"
        ValidaInscDA = bOk
        Exit Function
    Else
        ' If GrupoReceita <> "E" And GrupoReceita <> "I" And GrupoReceita <> "J" Then
        If Not lValidaGrupo Then
             If IsNull(txtDividaAtiva) Then
                txtDividaAtiva = "0000000000000"
                ValidaInscDA = bOk
                Exit Function
             ElseIf Val(txtDividaAtiva) > 0 Then
                MsgBox "Para o grupo da receita digitado, este campo não pode estar preenchido.", vbInformation + vbOKOnly, App.Title
             End If
            txtDividaAtiva = "0000000000000"
            ValidaInscDA = bOk
            Exit Function
        End If
        If Len(txtDividaAtiva) <= 9 Then
            txtDividaAtiva = Format(txtDividaAtiva, "000000000")
            If Mid$(txtDividaAtiva, 1, 9) = "000000000" Then
                bOk = False
                ValidaInscDA = bOk
                Exit Function
            End If
            soma = soma + Mid(txtDividaAtiva, 1, 1) * 1
            soma = soma + Mid(txtDividaAtiva, 2, 1) * 3
            soma = soma + Mid(txtDividaAtiva, 3, 1) * 4
            soma = soma + Mid(txtDividaAtiva, 4, 1) * 5
            soma = soma + Mid(txtDividaAtiva, 5, 1) * 6
            soma = soma + Mid(txtDividaAtiva, 6, 1) * 7
            soma = soma + Mid(txtDividaAtiva, 7, 1) * 8
            soma = soma + Mid(txtDividaAtiva, 8, 1) * 10
            resto = soma Mod 11                     'resto da divisão
            digito_11 = Right(str(resto), 1)        'digito verificador
            digito_rv = Mid(txtDividaAtiva, 9, 1)   'digito verificador
        Else
            If Len(txtDividaAtiva) < 11 Then
                Erro = 1
            Else
                txtDividaAtiva = Format$(txtDividaAtiva, "0000000000000")
                If Val(Mid(txtDividaAtiva, 1, 2)) = 0 Then
                    Erro = 2
                    ValidaInscDA = bOk
                    Exit Function
                End If
                soma = soma + Mid(txtDividaAtiva, 1, 1) * 1
                soma = soma + Mid(txtDividaAtiva, 2, 1) * 2
                soma = soma + Mid(txtDividaAtiva, 3, 1) * 3
                soma = soma + Mid(txtDividaAtiva, 4, 1) * 4
                soma = soma + Mid(txtDividaAtiva, 5, 1) * 5
                soma = soma + Mid(txtDividaAtiva, 6, 1) * 6
                soma = soma + Mid(txtDividaAtiva, 7, 1) * 7
                soma = soma + Mid(txtDividaAtiva, 8, 1) * 8
                soma = soma + Mid(txtDividaAtiva, 9, 1) * 9
                soma = soma + Mid(txtDividaAtiva, 10, 1) * 10
                soma = soma + Mid(txtDividaAtiva, 11, 1) * 11
                soma = soma + Mid(txtDividaAtiva, 12, 1) * 12
                resto = soma Mod 11                     'resto da divisão
                digito_11 = Right(str(resto), 1)        'digito verificador
                digito_rv = Mid(txtDividaAtiva, 13, 1)  'digito verificador
            End If
        End If
    End If
    If CStr(digito_11) <> (digito_rv) Then
        bOk = False             'digito não confere
    End If

    ValidaInscDA = bOk
   
End Function
Public Function ValidaInscricao(ByVal sTipoCampo As String) As Boolean
    
    Erro = 0  ' Retorna Codigo de Erro
   
    Select Case sTipoCampo
    
        Case 0
            ' Campo não está habilitado
             ValidaInscricao = True
             Exit Function
             
        Case 1
            ' Campo Obrigatório Inscrição Estadual
        
            If Val_Inscricao(txtInscEstadual) Then
                ValidaInscricao = True
                Exit Function
            Else
                ValidaInscricao = False
                Erro = 1
                Exit Function
               
            End If
            
        Case 2
            ' Valida Código do Municipio
            
            ' Critica os Códigos Fixos
            If ((Val(Mid(txtInscEstadual, 1, 3)) > 800) And (Val(Mid(txtInscEstadual, 1, 3)) < 900)) Or (Val(Mid(txtInscEstadual, 1, 3)) = 999) Then
                ValidaInscricao = False
                Erro = 2
                Exit Function
            End If
            
            If Val_Municipio(txtInscEstadual) Then
                ValidaInscricao = True
                Exit Function
            Else
                ValidaInscricao = False
                Erro = 2
                Exit Function
            End If
        Case 3
            ' Valida Número da Declaração
            
            If Val_Declaracao(txtInscEstadual) Then
                ValidaInscricao = True
                Exit Function
            Else
                ValidaInscricao = False
                Erro = 5
                Exit Function
            End If
            
        Case 4
            ' A Inscrição é dependente mas se está preenchida será validada
            If Val(txtInscEstadual) > 0 Then
                
                ' Verifica os Códigos Fixos
                If Not (Val(Mid(txtInscEstadual.Text, 1, 3)) >= 100 And Val(Mid(txtInscEstadual.Text, 1, 3)) <= 794) Or _
                   (Val(Mid(txtInscEstadual.Text, 1, 3)) >= 801 And Val(Mid(txtInscEstadual.Text, 1, 3)) <= 899) Or _
                   (Val(Mid(txtInscEstadual.Text, 1, 3)) = 999) Then
                   
                    ValidaInscricao = False
                    Erro = 1
                    Exit Function
                
                End If
                
                If Val_Inscricao(txtInscEstadual) Then
                    ValidaInscricao = True
                    Exit Function
                Else
                    ValidaInscricao = False
                    Erro = 1
                    Exit Function
                End If
                                
            End If

    End Select
   
End Function

Private Sub cmdConfirmar_Click()

    txtValor_KeyPress (vbKeyReturn)

End Sub

Private Sub cmdFrenteVerso_Click()

    mForm.cmdFrenteVerso_Click
    bAlterou = True
    
End Sub

Private Sub cmdInverteCor_Click()
    
    mForm.cmdInverteCor_Click
    
End Sub
Private Sub cmdRotacao_Click()
    
    mForm.cmdRotacao_Click
    bAlterou = True
    
End Sub
Private Sub CmdSair_Click()
    
    Alterou = False
    Me.Hide
    
End Sub
Private Sub cmdZoomMais_Click()
    
    mForm.cmdZoomMais_Click
    bAlterou = True
    
End Sub
Private Sub cmdZoomMenos_Click()
    
    mForm.cmdZoomMenos_Click
    bAlterou = True
    
End Sub

Private Sub Form_Activate()

'    Dim iErroData  As Integer

    'Verifica se feriado na agência
    If Not ValidaAgenciaPorDocto(Geral.Capa.AgOrig, "", False) Then
        CmdSair_Click
        Exit Sub
    End If
    
'    iErroData = ValidaAgencia(Geral.Capa.AgOrig, "", False)
'    If iErroData <> 0 Then
'        Select Case iErroData
'            Case 2 'Feriado
'                MsgBox "A Agência de Origem está em Feriado.", vbInformation, App.Title
'            Case 3 'Agência Fechada
'                MsgBox "A Agência de Origem está Fechada.", vbInformation, App.Title
'        End Select
'        If iErroData = 2 Or iErroData = 3 Then
'            CmdSair_Click
'            Exit Sub
'        End If
'    End If
    
    'Ler dados da tabela GARE, pode ou não existir registro
    If Not LerDadosGare() Then CmdSair_Click: Exit Sub
    
   
    bAlterou = False
    
    If Val(txtReceita.Text) <> 0 Then
        'Desabilita controles do documento conforme codigo da receita
        Call VerificaReceita
    Else
        'Desabilita controles do documento se cadastramento
        Call MontaTela("#")
    End If
    
    datVencimento.SetFocus
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

  Select Case KeyCode
    Case vbKeyAdd
      cmdZoomMais_Click
    Case vbKeySubtract
      cmdZoomMenos_Click
    Case vbKeyF10
      cmdInverteCor_Click
      KeyCode = 0
    Case vbKeyDivide
      cmdRotacao_Click
    Case vbKeyF11
      cmdFrenteVerso_Click
    Case vbKeyMultiply
        Call CalculaValorTotal
        Call cmdConfirmar_Click
    Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight
        mForm.Form_KeyUp KeyCode, Shift
  End Select
End Sub


Private Sub Form_Load()
    bAlterar = True
    
    Alterou = False
    
    With Modulo
        ' Cria a query para a gravação dos dados do GARE
        Set .qryInserirGARE = Geral.Banco.CreateQuery("", "{? = call InserirGARE (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)}")
    
    
        ' Cria a query para a gravação da Tabela Gare
        Set .qryInserirTFSCDARF = Geral.Banco.CreateQuery("", "{? = call InserirTFSCDARF (?,?,?,?,?,?,?,?,?,?,?)}")
    
        
         ' Cria a query para a gravação da Tabela Gare
        Set .qryInserirTFSCRGPS = Geral.Banco.CreateQuery("", "{? = call InserirTFSCRGPS (?,?,?,?,?,?,?,?,?,?,?,?,?)}")
        
        
        ' Cria a query para verficar convenio do Gare
        Set .qryConvenioGare = Geral.Banco.CreateQuery("", "{? = call ConvenioDarmGare(?)}")
    
        ' Cria a query para verificar código receita Gare        '
        'Set .qryGetValidaGare = Geral.Banco.CreateQuery("", "{call GetValidaGARE(?)}") ' Critica antiga
        'Parâmetro: (0)-código da Receita
        
        ' Nova critica baseada na tabela RegraMDI_Gare
        Set .qryGetValidaGare = Geral.Banco.CreateQuery("", "{call GetValidaGrupoGARE(?)}")
        
        
        ' Query para verificar quais campos sarão abilitados no Gare
        Set .qryGetAtivaCampoGare = Geral.Banco.CreateQuery("", "{call GetAtivaCampoGare(?,?)}")
        
        
        ' Query para verificar qual o tipo de campo no Gare
        Set .qryGetTipoCampoGare = Geral.Banco.CreateQuery("", "{call GetTipoCampoGare(?,?)}")
        
        
        
        ' Cria a query para ler dados do Gare/Icms
        Set .qryGetGare = Geral.Banco.CreateQuery("", "{? = call GetGare (?,?)}")
            'Parâmetro: (1)-Data Processamento (2)-IdDocto
            .qryGetGare.rdoParameters(0).Direction = rdParamReturnValue
        
        Set .qryRemoveTipoDocumento = Geral.Banco.CreateQuery("", "{? = call RemoveTipoDocumento (?,?,?)}")
    
        
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)

    With Modulo
        .qryInserirGARE.Close
        .qryConvenioGare.Close
        .qryGetValidaGare.Close
        .qryGetGare.Close
        .qryRemoveTipoDocumento.Close

    End With
    
End Sub

Private Sub datVencimento_GotFocus()

    'Se controle desativado, força foco no próximo controle habilitado
    If Not datVencimento.TabStop Then SendKeys "{TAB}": Exit Sub
    
    With datVencimento
        If Len(.Text) = 0 Then Exit Sub
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub txtAcrescimo_GotFocus()

    If Val(txtReceita.Text) = 0 Then
     MsgBox "Informe o Código da Receita.", vbInformation + vbOKOnly, App.Title
     txtReceita_GotFocus
     txtReceita.SetFocus
    End If

    With txtAcrescimo
        If Len(.Text) = 0 Then Exit Sub
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub txtAcrescimo_KeyPress(KeyAscii As Integer)

Dim sGruposAtivos As String
Dim lValidaGrupo As Boolean
    
    InibirTeclaAlfaValor KeyAscii
    If Not (KeyAscii = 0 Or KeyAscii = 13) Then bAlterou = True
    
    bAlterou = True
    
    If KeyAscii = 13 Then
        
        If VerificaReceita = False Then
            txtReceita_GotFocus
            txtReceita.SetFocus
            Exit Sub
        End If
    
    
        ' Carrega os grupos Ativos para Valor_Acrescimo_Financ
        ' Os grupos dependentes são  marcados como '9'
        ' na tabela TFSCRGPS
    
        sGruposAtivos = AtivaCampos("Valor_Acrescimo_Financ", "9")
        lValidaGrupo = IIf(InStr(sGruposAtivos, GrupoReceita) = 0, False, True)
           
    
        ' If (GrupoReceita <> "B") And (GrupoReceita <> "F") Then
        If Not lValidaGrupo Then
            If Val(Desformata_Valor(txtAcrescimo.Text)) > 0 Then
                MsgBox "Para o grupo da receita digitado, este campo não pode estar preenchido.", vbInformation + vbOKOnly, App.Title
            End If
            txtAcrescimo.Text = "0"
        End If
        txtValor.Text = Val(Desformata_Valor(txtValorReceita.Text)) + Val(Desformata_Valor(TxtJuros.Text)) + Val(Desformata_Valor(txtMulta.Text)) + Val(Desformata_Valor(txtAcrescimo.Text)) + Val(Desformata_Valor(txtHonorarios.Text))
        SendKeys "{TAB}"
    ElseIf (KeyAscii = 27) Then
         KeyAscii = 0
         CmdSair_Click
    End If

End Sub

Private Sub txtAcrescimo_LostFocus()
            
    txtValor.Text = Val(Desformata_Valor(txtValorReceita.Text)) + Val(Desformata_Valor(TxtJuros.Text)) + Val(Desformata_Valor(txtMulta.Text)) + Val(Desformata_Valor(txtAcrescimo.Text)) + Val(Desformata_Valor(txtHonorarios.Text))

End Sub

Private Sub txtAIIM_GotFocus()
    
    If Val(txtReceita.Text) = 0 Then
     MsgBox "Informe o Código da Receita.", vbInformation + vbOKOnly, App.Title
     txtReceita_GotFocus
     txtReceita.SetFocus
    End If
        
    With txtAIIM
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtAIIM_KeyPress(KeyAscii As Integer)
    
    'Verifica se Campo desabilitado (Locked)
    If Not txtAIIM.TabStop Then SendKeys "{TAB}": Exit Sub
    
    InibirTeclaAlfa KeyAscii
    If Not (KeyAscii = 0 Or KeyAscii = 13) Then bAlterou = True
   
    If KeyAscii = 13 Then
    
        'Valida AIIM
        If VerificarAIIM Then
            If txtValorReceita.Enabled Then
                txtValorReceita.SetFocus
            Else
                SendKeys "{TAB}"
            End If
        End If
        
    ElseIf (KeyAscii = 27) Then
         KeyAscii = 0
         CmdSair_Click
    End If

End Sub
Private Sub txtCGCCPF_GotFocus()
    
    If Val(txtReceita.Text) = 0 Then
     MsgBox "Informe o Código da Receita.", vbInformation + vbOKOnly, App.Title
     txtReceita_GotFocus
     txtReceita.SetFocus
    End If
    
    
    With txtCGCCPF
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtCGCCPF_KeyPress(KeyAscii As Integer)
    
    'Verifica se Campo desabilitado (Locked)
    If Not txtCGCCPF.TabStop Then SendKeys "{TAB}": Exit Sub
    
    InibirTeclaAlfa KeyAscii
    If Not (KeyAscii = 0 Or KeyAscii = 13) Then bAlterou = True

    If KeyAscii = 13 Then
    
        'Valida CGCCPF
        If VerificarCGCCPF Then
            If txtDividaAtiva.TabStop Then
                txtDividaAtiva_GotFocus
                txtDividaAtiva.SetFocus
            Else
                SendKeys "{TAB}"
            End If
        Else
            txtCGCCPF_GotFocus
            txtCGCCPF.SetFocus
        End If
    
    ElseIf (KeyAscii = 27) Then
         KeyAscii = 0
         CmdSair_Click
    End If

End Sub
Private Sub txtDividaAtiva_GotFocus()
    
    If Val(txtReceita.Text) = 0 Then
     MsgBox "Informe o Código da Receita.", vbInformation + vbOKOnly, App.Title
     txtReceita_GotFocus
     txtReceita.SetFocus
    End If
    
    With txtDividaAtiva
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtDividaAtiva_KeyPress(KeyAscii As Integer)
    
    'Verifica se Campo desabilitado (Locked)
    If Not txtDividaAtiva.TabStop Then SendKeys "{TAB}": Exit Sub
    
    InibirTeclaAlfa KeyAscii
    If Not (KeyAscii = 0 Or KeyAscii = 13) Then bAlterou = True
   
    If KeyAscii = 13 Then
                
        'Valida DividaAtiva
        If VerificarDividaAtiva = True Then
            If txtReferencia.TabStop Then
                txtReferencia_GotFocus
                txtReferencia.SetFocus
            Else
                SendKeys "{TAB}"
            End If
        End If
        
    ElseIf (KeyAscii = 27) Then
         KeyAscii = 0
         CmdSair_Click
    End If

End Sub

Private Sub txtHonorarios_GotFocus()

    If Val(txtReceita.Text) = 0 Then
     MsgBox "Informe o Código da Receita.", vbInformation + vbOKOnly, App.Title
     txtReceita_GotFocus
     txtReceita.SetFocus
    End If

    With txtHonorarios
        If Len(.Text) = 0 Then Exit Sub
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub txtHonorarios_KeyPress(KeyAscii As Integer)

    Dim sGruposAtivos As String
    Dim lValidaGrupo As Boolean
    
    InibirTeclaAlfaValor KeyAscii
    If Not (KeyAscii = 0 Or KeyAscii = 13) Then bAlterou = True
        
    If KeyAscii = 13 Then
        If VerificaReceita = False Then
            If txtReceita.TabStop Then
                txtReceita_GotFocus
                txtReceita.SetFocus
            Else
                SendKeys "{TAB}"
            End If
            Exit Sub
        End If
    
        ' Carrega os grupos Ativos para Valor_Honor_Advoc
        ' Os grupos dependentes são  marcados como '9'
        ' na tabela TFSCRGPS
    
        sGruposAtivos = AtivaCampos("Valor_Honor_Advoc", "9")
        lValidaGrupo = IIf(InStr(sGruposAtivos, GrupoReceita) = 0, False, True)
    
        ' If (GrupoReceita <> "E") And (GrupoReceita <> "F") And (GrupoReceita <> "G") And (GrupoReceita <> "I") And (GrupoReceita <> "J") Then
        If Not lValidaGrupo Then
            If Val(Desformata_Valor(txtHonorarios.Text)) > 0 Then
                MsgBox "Para o grupo da receita digitado, este campo não pode estar preenchido.", vbInformation + vbOKOnly, App.Title
            End If
            txtHonorarios.Text = "0"
        End If
        txtValor.Text = Val(Desformata_Valor(txtValorReceita.Text)) + Val(Desformata_Valor(TxtJuros.Text)) + Val(Desformata_Valor(txtMulta.Text)) + Val(Desformata_Valor(txtAcrescimo.Text)) + Val(Desformata_Valor(txtHonorarios.Text))
        txtValor.SetFocus
    ElseIf (KeyAscii = 27) Then
         KeyAscii = 0
         CmdSair_Click
    End If

End Sub

Private Sub txtHonorarios_LostFocus()
            
    txtValor.Text = Val(Desformata_Valor(txtValorReceita.Text)) + Val(Desformata_Valor(TxtJuros.Text)) + Val(Desformata_Valor(txtMulta.Text)) + Val(Desformata_Valor(txtAcrescimo.Text)) + Val(Desformata_Valor(txtHonorarios.Text))

End Sub

Private Sub txtInscEstadual_GotFocus()
    
    If Val(txtReceita.Text) = 0 Then
     MsgBox "Informe o Código da Receita.", vbInformation + vbOKOnly, App.Title
     txtReceita_GotFocus
     txtReceita.SetFocus
    End If
    
    
    With txtInscEstadual
        .SelStart = 0
        .SelLength = .MaxLength
    End With
    
End Sub

Private Sub txtInscEstadual_KeyPress(KeyAscii As Integer)
    
    'Verifica se Campo desabilitado (Locked)
    If Not txtInscEstadual.TabStop Then SendKeys "{TAB}": Exit Sub
    
    InibirTeclaAlfa KeyAscii
    If Not (KeyAscii = 0 Or KeyAscii = 13) Then bAlterou = True
    
    If KeyAscii = 13 Then
        'Valida Insc Estadual
        If VerificaInscEstadual = True Then
            If txtCGCCPF.TabStop Then
                txtCGCCPF_GotFocus
                txtCGCCPF.SetFocus
            Else
                SendKeys "{TAB}"
            End If
        Else
            txtInscEstadual_GotFocus
            txtInscEstadual.SetFocus
        End If
    ElseIf (KeyAscii = 27) Then
         KeyAscii = 0
         CmdSair_Click
    End If

End Sub

Private Sub txtJuros_GotFocus()
    
    If Val(txtReceita.Text) = 0 Then
     MsgBox "Informe o Código da Receita.", vbInformation + vbOKOnly, App.Title
     txtReceita_GotFocus
     txtReceita.SetFocus
    End If
    
    
    With TxtJuros
        If Len(.Text) = 0 Then Exit Sub
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub TxtJuros_KeyPress(KeyAscii As Integer)
    
    InibirTeclaAlfaValor KeyAscii
    If Not (KeyAscii = 0 Or KeyAscii = 13) Then bAlterou = True
        
    If KeyAscii = 13 Then
        txtValor.Text = Val(Desformata_Valor(txtValorReceita.Text)) + Val(Desformata_Valor(TxtJuros.Text)) + Val(Desformata_Valor(txtMulta.Text)) + Val(Desformata_Valor(txtAcrescimo.Text)) + Val(Desformata_Valor(txtHonorarios.Text))
        If txtMulta.Enabled Then
            txtMulta.SetFocus
        Else
            SendKeys "{TAB}"
        End If
    ElseIf (KeyAscii = 27) Then
         KeyAscii = 0
         CmdSair_Click
    End If

End Sub

Private Sub txtJuros_LostFocus()
            
    txtValor.Text = Val(Desformata_Valor(txtValorReceita.Text)) + Val(Desformata_Valor(TxtJuros.Text)) + Val(Desformata_Valor(txtMulta.Text)) + Val(Desformata_Valor(txtAcrescimo.Text)) + Val(Desformata_Valor(txtHonorarios.Text))

End Sub

Private Sub txtMulta_GotFocus()
        
    If Val(txtReceita.Text) = 0 Then
     MsgBox "Informe o Código da Receita.", vbInformation + vbOKOnly, App.Title
     txtReceita_GotFocus
     txtReceita.SetFocus
    End If
        
    With txtMulta
        If Len(.Text) = 0 Then Exit Sub
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub txtMulta_KeyPress(KeyAscii As Integer)
    
    InibirTeclaAlfaValor KeyAscii
    If Not (KeyAscii = 0 Or KeyAscii = 13) Then bAlterou = True
        
    If KeyAscii = 13 Then
        txtValor.Text = Val(Desformata_Valor(txtValorReceita.Text)) + Val(Desformata_Valor(TxtJuros.Text)) + Val(Desformata_Valor(txtMulta.Text)) + Val(Desformata_Valor(txtAcrescimo.Text)) + Val(Desformata_Valor(txtHonorarios.Text))
        If txtAcrescimo.Enabled Then
            txtAcrescimo.SetFocus
        Else
            SendKeys "{TAB}"
        End If
    ElseIf (KeyAscii = 27) Then
         KeyAscii = 0
         CmdSair_Click
    End If

End Sub

Private Sub txtMulta_LostFocus()
            
    txtValor.Text = Val(Desformata_Valor(txtValorReceita.Text)) + Val(Desformata_Valor(TxtJuros.Text)) + Val(Desformata_Valor(txtMulta.Text)) + Val(Desformata_Valor(txtAcrescimo.Text)) + Val(Desformata_Valor(txtHonorarios.Text))

End Sub

Private Sub txtReceita_GotFocus()
    
    With txtReceita
        .SelStart = 0
        .SelLength = .MaxLength
    End With
    
End Sub

Private Sub txtReceita_KeyPress(KeyAscii As Integer)
    
'Dim iErroData As Integer
    
    'Verifica se Campo desabilitado (Locked)
    If Not txtReceita.TabStop Then SendKeys "{TAB}": Exit Sub
    
    InibirTeclaAlfa KeyAscii
    If Not (KeyAscii = 0 Or KeyAscii = 13) Then bAlterou = True
    
    If KeyAscii = 13 Then
        'Valida Receita
        If VerificaReceita Then
            'Consiste data de vencto somente para receita dos grupos "B" e "F"
            If txtReceita.Tag = "B" Or txtReceita.Tag = "F" Then
                'Verifica se Feriado / Agência aberta e Data Válida
                If Len(datVencimento.Text) = 8 Then
                    If Not ValidaAgenciaPorDocto(Geral.Capa.AgOrig, datVencimento.Text, True, True) Then
                        datVencimento.SetFocus
                        Exit Sub
                    End If
                    
'                    iErroData = ValidaAgencia(Geral.Capa.AgOrig, datVencimento.Text, True)
'                    If iErroData <> 0 Then
'                        Select Case iErroData
'                            Case 2 'Feriado
'                                MsgBox "A Agência de Origem está em Feriado.", vbInformation, App.Title
'                            Case 3 'Agência Fechada
'                                MsgBox "A Agência de Origem está Fechada.", vbInformation, App.Title
'                            Case 1 'Documento Vencido
'                                MsgBox "A Data de Vencimento deve ser maior que a Data do Movimento Anterior.", vbInformation, App.Title
'                        End Select
'                        If iErroData = 1 Or iErroData = 2 Or iErroData = 3 Then
'                            datVencimento_GotFocus
'                            datVencimento.SetFocus
'                            Exit Sub
'                        End If
'                    End If
                Else
                    'Verifica se feriado na agência
                    If Not ValidaAgenciaPorDocto(Geral.Capa.AgOrig, "", False) Then
                        datVencimento.SetFocus
                        Exit Sub
                    End If
                    
'                    iErroData = ValidaAgencia(Geral.Capa.AgOrig, "", False)
'                    Select Case iErroData
'                        Case 2 'Feriado
'                            MsgBox "A Agência de Origem está em Feriado.", vbInformation, App.Title
'                        Case 3 'Agência Fechada
'                            MsgBox "A Agência de Origem está Fechada.", vbInformation, App.Title
'                    End Select
'                    If iErroData = 2 Or iErroData = 3 Then
'                        datVencimento_GotFocus
'                        datVencimento.SetFocus
'                        Exit Sub
'                    End If
                End If
            End If
        
            SendKeys "{TAB}"
        End If
    ElseIf (KeyAscii = 27) Then
         KeyAscii = 0
         CmdSair_Click
    End If

End Sub

Private Sub TxtReceita_LostFocus()
    
    If Val(txtReceita.Text) <> 0 Then
        VerificaReceita
    End If
    
End Sub

Private Sub txtReferencia_GotFocus()

    If Val(txtReceita.Text) = 0 Then
     MsgBox "Informe o Código da Receita.", vbInformation + vbOKOnly, App.Title
     txtReceita_GotFocus
     txtReceita.SetFocus
    End If

    With txtReferencia
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtReferencia_KeyPress(KeyAscii As Integer)

    'Verifica se Campo desabilitado (Locked)
    If Not txtReferencia.TabStop Then SendKeys "{TAB}": Exit Sub

    InibirTeclaAlfa KeyAscii
    If Not (KeyAscii = 0 Or KeyAscii = 13) Then bAlterou = True

    If KeyAscii = 13 Then
        'Valida Referencia
        If (Len(Trim(txtReferencia.Text)) = 6 Or Len(Trim(txtReferencia.Text)) = 7) Or (Len(Trim(txtReferencia.Text)) = 1) Then
            
            If Len(Trim(txtReferencia.Text)) = 6 Then
                txtReferencia.Text = Format(txtReferencia.Text, "@@/@@@@")
            End If
            
            If VerificarReferencia Then
                If txtAIIM.TabStop Then
                    txtAIIM_GotFocus
                    txtAIIM.SetFocus
                Else
                    SendKeys "{TAB}"
                End If
            End If
        Else
            MsgBox "Referência Inválida", vbInformation + vbOKOnly, App.Title
            txtReferencia.SetFocus
            Exit Sub
        End If
    ElseIf (KeyAscii = 27) Then
         KeyAscii = 0
         CmdSair_Click
    End If
End Sub
Public Function VerificaDataMMAAAA(ByVal pviData As String) As Boolean
    
    '''''''''''''''''''''''''''''''''''''''''''''
    ' Retorna True se a data é válida           '
    ' Data deve ser informada no formato MMAAAA '
    '''''''''''''''''''''''''''''''''''''''''''''
    
    Dim iMes As String
    Dim iAno As String
    Dim sData As String
    Dim bOk As Boolean
    
    bOk = True
    
    sData = pviData
    
    iMes = Mid(sData, 1, 2)
    iAno = Right(sData, 4)
    
    ''''''''''''''''''''''''''''''''''''''
    ' Verifica se mês está entre 01 e 12 '
    ''''''''''''''''''''''''''''''''''''''
    If Val(iMes) < 1 Or Val(iMes) > 12 Then
        bOk = False
    End If
    
    If Val(iAno) < 1950 Then
        bOk = False
    End If
        
    VerificaDataMMAAAA = bOk

End Function
Function VerificaConvenioGare() As Boolean
    
    VerificaConvenioGare = False    'default - false
    
    'verifica se a agencia processadora é de SP
    With Modulo.qryConvenioGare
        .rdoParameters(1) = Geral.AgenciaCentral
        .Execute
        If .rdoParameters(0) = 0 Then
            VerificaConvenioGare = True     'agencia é de SP
        Else
            VerificaConvenioGare = False    'agencia não é de SP
        End If
    End With

End Function
Private Sub txtReferencia_LostFocus()

    'Valida Referência do Gare

    If Len(Trim(txtReferencia)) = 0 Then Exit Sub
    
    If Len(txtReferencia.Text) = 6 Then
        txtReferencia.Text = Format(txtReferencia.Text, "@@/@@@@")
    End If

    
    If VerificarReferencia = True Then
        SelecionarTexto txtValorReceita
    End If
    
End Sub
Private Sub txtValor_GotFocus()
    
    If Val(txtReceita.Text) = 0 Then
     MsgBox "Informe o Código da Receita.", vbInformation + vbOKOnly, App.Title
     txtReceita_GotFocus
     txtReceita.SetFocus
    End If
    
    With Me.txtValor
        If Len(.Text) = 0 Then Exit Sub
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)

Dim sGruposAtivos   As String
Dim lValidaGrupo    As Boolean
Dim strEncripta     As String

On Error GoTo ERRO_TXTVALOR_GARE

    Dim dt_v As String
    Dim cod_m As String
    Dim cgc_cpf As String
    Dim insc_div As String
    Dim dt_ref As String
    Dim n_aiim As String
    Dim dValor As Double
    Dim bDuplicidade  As Boolean
    
    InibirTeclaAlfaValor KeyAscii
    If Not (KeyAscii = 0 Or KeyAscii = 13) Then bAlterou = True
        
    If KeyAscii = 13 Then
        
        'verificar se agencia é de SP para poder aceitar o docto Gare
        If Not VerificaConvenioGare Then
            MsgBox "Esta agência não aceita Gare. Favor verificar.", vbInformation, "Atenção"
            txtValor.SetFocus
            Exit Sub
        End If
                
        'Valida Todos os campos
        If VerificarTudo Then
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' So coleta data de vencimento quando grupo receita = B ou F '
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
                ' Carrega os grupos  para Vecto_Normal
                ' marcados como '1'
                ' na tabela TFSCRGPS
    
                sGruposAtivos = AtivaCampos("Vecto_Normal", "1")
                lValidaGrupo = IIf(InStr(sGruposAtivos, GrupoReceita) = 0, False, True)
                
                ' If (GrupoReceita = "B") Or (GrupoReceita = "F") Then
                If lValidaGrupo Then
                    dt_v = datVencimento.Text
                Else
                    dt_v = "00000000"
                End If
            
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' So coleta cod. municipio (insc.estadual) qdo grupo receita diferente de G e I '
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
                ' Carrega os grupos  para Inscricao_Estadual
                ' marcados como '0'
                ' na tabela TFSCRGPS
    
                sGruposAtivos = AtivaCampos("Inscricao_Estadual", "0")
                lValidaGrupo = IIf(InStr(sGruposAtivos, GrupoReceita) = 0, False, True)
                
                
                ' If (GrupoReceita = "G") Or (GrupoReceita = "I") Then
                If lValidaGrupo Then
                    cod_m = "0"
                Else
                    cod_m = txtInscEstadual
                End If
            
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' So coleta CGC/CPF qdo grupo receita diferente de A,B,E,F '
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
                ' Carrega os grupos  para Campo_Docto
                ' marcados como '0'
                ' na tabela TFSCRGPS
    
                sGruposAtivos = AtivaCampos("Campo_Docto", "0")
                lValidaGrupo = IIf(InStr(sGruposAtivos, GrupoReceita) = 0, False, True)
                
                
                
'                If (GrupoReceita = "A") Or (GrupoReceita = "B") Or _
'                   (GrupoReceita = "E") Or (GrupoReceita = "F") Then

                If lValidaGrupo Then
                    cgc_cpf = "0"
                Else
                    cgc_cpf = Format(txtCGCCPF, "000000000000000")
                End If
            
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' So coleta inscricao na divida qdo grupo receita = E,I,J,L,M '
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
                ' Carrega os grupos  para Inscricao_Ativa
                ' marcados como '3'
                ' na tabela TFSCRGPS
    
                sGruposAtivos = AtivaCampos("Inscricao_Ativa", "3")
                lValidaGrupo = IIf(InStr(sGruposAtivos, GrupoReceita) = 0, False, True)
                
                
'                If (GrupoReceita = "E") Or (GrupoReceita = "I") Or _
'                   (GrupoReceita = "J") Or (GrupoReceita = "L") Or (GrupoReceita = "M") Then
                
                If lValidaGrupo Then
                   insc_div = txtDividaAtiva
                Else
                   insc_div = "0"
                End If
            
                '''''''''''''''''''''''''''''''''''''''''''''''''''
                ' So coleta data referencia qdo grupo receita = A '
                '''''''''''''''''''''''''''''''''''''''''''''''''''
                
                ' Carrega os grupos  para Referencia
                ' marcados como '1'
                ' na tabela TFSCRGPS
    
                sGruposAtivos = AtivaCampos("Referencia", "1")
                lValidaGrupo = IIf(InStr(sGruposAtivos, GrupoReceita) = 0, False, True)
                
                
                
                ' If (GrupoReceita = "A") Then
                If lValidaGrupo Then
                
                    If VerificarReferencia() Then
                        
                        dt_ref = txtReferencia.Text
                        
                        ''''''''''''''''''''''''''''''''''''''
                        'Como não se acha o buf da referencia'
                        ''''''''''''''''''''''''''''''''''''''
                        If Not IsNumeric(Left(dt_ref, 2)) Then
                            MsgBox "Erro na data de referência.", vbExclamation
                            Exit Sub
                        End If
                        If Val(Left(dt_ref, 2)) < 1 Or Val(Left(dt_ref, 2)) > 12 Then
                            MsgBox "Erro na data de referência.", vbExclamation
                            Exit Sub
                        End If
                        
                        If Not IsNumeric(Right(dt_ref, 4)) Then
                            MsgBox "Erro na data de referência.", vbExclamation
                            Exit Sub
                        End If
                        If Val(Right(dt_ref, 4)) < 1950 Or Val(Right(dt_ref, 4)) > 2100 Then
                            MsgBox "Erro na data de referência.", vbExclamation
                            Exit Sub
                        End If
                    Else
                        MsgBox "Erro na data de referência.", vbExclamation
                        Exit Sub
                    End If
                Else
                   dt_ref = "000000"
                End If
            
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' So coleta AIIM qdo grupo receita diferente de A,D,E,G,I,J '
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
                
                ' Carrega os grupos  para Numero_Parcelamento
                ' marcados como '0'
                ' na tabela TFSCRGPS
    
                sGruposAtivos = AtivaCampos("Numero_Parcelamento", "0")
                lValidaGrupo = IIf(InStr(sGruposAtivos, GrupoReceita) = 0, False, True)
                
'                If (GrupoReceita = "A") Or (GrupoReceita = "D") Or (GrupoReceita = "E") Or _
'                   (GrupoReceita = "G") Or (GrupoReceita = "I") Or (GrupoReceita = "J") Then
                
                If lValidaGrupo Then
                    n_aiim = "0"
                Else
                    n_aiim = txtAIIM
                End If
            
                '***********************************
                '*** Insere dados na tabela Gare ***
                '***********************************
                sPosicaoErro = "InsGare"
                dValor = (Val(txtValor.Text) / 100)
                
                'Inicia Transação
                Geral.Banco.BeginTrans
                
                'Verificar se o Documento pertence à outro Tipo
                If Geral.Documento.TipoDocto <> 18 And Geral.Documento.TipoDocto <> 0 Then
                  With Modulo.qryRemoveTipoDocumento
                    .rdoParameters(1) = Geral.DataProcessamento     'Data Proc.
                    .rdoParameters(2) = Geral.Documento.IdDocto     'IdDocto
                    .rdoParameters(3) = Geral.Documento.TipoDocto   'Tipo do Documento
                    .Execute
                  End With
                End If
                
                'Atualiza campo Autenticação Digital
                If Val(cgc_cpf) <> 0 Then
                    strEncripta = G_EncriptaBO(etpdocGare, CStr(Val(cgc_cpf)))
                    If strEncripta = "" Then GoTo Exit_SalvaDados
                Else
                    strEncripta = G_EncriptaBO(etpdocGare, CStr(Val(cod_m)))
                    If strEncripta = "" Then GoTo Exit_SalvaDados
                End If
                
                With Modulo.qryInserirGARE
                    .rdoParameters(1) = Geral.DataProcessamento
                    .rdoParameters(2) = Geral.Documento.IdDocto
                    'Vencimento (aaaammdd)
                    .rdoParameters(3) = Val(Mid(dt_v, 5, 4) & Mid(dt_v, 3, 2) & Mid(dt_v, 1, 2))
                    'Código da Receita
                    .rdoParameters(4) = Val(txtReceita.Text)
                    'Inscrição Estadual
                    .rdoParameters(5) = Val(cod_m)
                    'CGC/CPF
                    .rdoParameters(6) = Val(cgc_cpf)
                    'Inscr. Dívida Ativa
                    .rdoParameters(7) = Val(insc_div)
                    'Data de Referência (aaaamm)
                    .rdoParameters(8) = Val(Right(dt_ref, 4) & Left(dt_ref, 2))
                    'Número do AIIM
                    .rdoParameters(9) = Val(n_aiim)
                    ' Autenticação Digital
                    .rdoParameters(10) = Null
                    'Vlr. da Receita
                    .rdoParameters(11) = (Val(txtValorReceita.Text) / 100)
                    'Vlr.do Juros
                    .rdoParameters(12) = (Val(TxtJuros.Text) / 100)
                    'Vlr. da Multa
                    .rdoParameters(13) = (Val(txtMulta.Text) / 100)
                    'Vlr. do Acréscimo
                    .rdoParameters(14) = (Val(txtAcrescimo.Text) / 100)
                    'Vlr. dos Honorários
                    .rdoParameters(15) = (Val(txtHonorarios.Text) / 100)
                    'Vlr. Total
                    .rdoParameters(16) = dValor
                    'Autenticacao digital
                    .rdoParameters(17) = strEncripta
                    .Execute
                
                    If .rdoParameters(0).Value <> 0 Then GoTo Exit_SalvaDados
                
                End With
                
                Geral.Documento.TipoDocto = etpdocGare
                Geral.Documento.ValorTotal = dValor
                Geral.Documento.Leitura = ""        'Não há necessidade de guardar campo leitura
            
                'Atualiza tabela Documento
                bDuplicidade = False    'Variavel somente com finalidade se atualizar campo Leitura
                If Not G_AtualizaCamposDocumento(bDuplicidade, Geral.Documento.IdDocto, , etpdocGare, , , , , dValor) Then
                    GoTo Exit_SalvaDados
                End If
                
                'Finalizou Transação
                Geral.Banco.CommitTrans
                
                Alterou = True
                Me.Hide
            End If
    ElseIf (KeyAscii = 27) Then
         KeyAscii = 0
         CmdSair_Click
    End If

    Exit Sub
    
Exit_SalvaDados:
    Alterou = False
    Geral.Banco.RollbackTrans
    MsgBox "Não foi possível Incluir/atualiza informações do GARE.", vbCritical + vbOKOnly, App.Title
'    cmdSair_Click
    Exit Sub
    
ERRO_TXTVALOR_GARE:
    Alterou = False
    Geral.Banco.RollbackTrans
    Select Case TratamentoErro("Não foi possível inserir o documento atual.(" & sPosicaoErro & ")", Err, rdoErrors)
        Case vbCancel
            CmdSair_Click
        Case vbRetry
    End Select

End Sub

Private Sub txtValorReceita_GotFocus()
    
    If Val(txtReceita.Text) = 0 Then
     MsgBox "Informe o Código da Receita.", vbInformation + vbOKOnly, App.Title
     txtReceita_GotFocus
     txtReceita.SetFocus
    End If
    
    With txtValorReceita
        If Len(.Text) = 0 Then Exit Sub
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Private Sub txtValorReceita_KeyPress(KeyAscii As Integer)
    
    InibirTeclaAlfaValor KeyAscii
    If Not (KeyAscii = 0 Or KeyAscii = 13) Then bAlterou = True
        
    If KeyAscii = 13 Then
        If Val(Desformata_Valor(txtValorReceita.Text)) <> 0 Then
            txtValor.Text = Val(Desformata_Valor(txtValorReceita.Text)) + Val(Desformata_Valor(TxtJuros.Text)) + Val(Desformata_Valor(txtMulta.Text)) + Val(Desformata_Valor(txtAcrescimo.Text)) + Val(Desformata_Valor(txtHonorarios.Text))
            If TxtJuros.Enabled Then
                TxtJuros.SetFocus
            Else
                SendKeys "{TAB}"
            End If
        Else
            'alterado em 28/04/2000 a pedidos de Selma
            TxtJuros.SetFocus
            'MsgBox "É obrigatório informar o Valor da Receita.", 16, "GARE"
            'txtValorReceita = ""
            'txtValorReceita.SetFocus
        End If
    ElseIf (KeyAscii = 27) Then
         KeyAscii = 0
         CmdSair_Click
    End If

End Sub
Private Sub datVencimento_KeyPress(KeyAscii As Integer)

    InibirTeclaAlfaCompensa KeyAscii
    If Not (KeyAscii = 0 Or KeyAscii = 13) Then bAlterou = True
    
    'Se teclado espaco, insere data de movimento na data de vencimento
    If KeyAscii = vbKeySpace And datVencimento.Locked = False Then
        datVencimento.Text = Right(Geral.DataProcessamento, 2) + Mid(Geral.DataProcessamento, 5, 2) + Left(Geral.DataProcessamento, 4)
    End If
    
    If KeyAscii = 13 Or (KeyAscii = 32) Then

        If KeyAscii = 32 Then
            If Len(datVencimento.Text) = 0 Then
                datVencimento.Text = Mid(Geral.DataProcessamento, 7, 2) & Mid(Geral.DataProcessamento, 5, 2) & Mid(Geral.DataProcessamento, 1, 4)
            End If
        End If

        'Verifica se Data Válida
        If Len(datVencimento.Text) <> 0 Then
        
            'Verifica se Data Válida
            If Not DataOk(datVencimento.Text) Then
                MsgBox "Data inválida. Verifique.", vbInformation + vbOKOnly, App.Title
                datVencimento_GotFocus
                datVencimento.SetFocus
                Exit Sub
            End If
            
            If ValidaData = True Then
                txtReceita_GotFocus
                txtReceita.SetFocus
            End If
        Else
            SendKeys "{TAB}"
        End If
        
    ElseIf (KeyAscii = 27) Then
        KeyAscii = 0
        CmdSair_Click
    End If

End Sub
Public Sub SetParent(ByRef aForm As Form)
  
  Set mForm = aForm

End Sub

Public Sub SetPosition(iLeft As Integer, iTop As Integer)

  Me.Left = iLeft
  Me.Top = iTop
  
End Sub

Private Function LerDadosGare() As Boolean

Dim rstModulo As rdoResultset

On Error GoTo Err_LerDadosGare
    
    LerDadosGare = False
    
    With Modulo.qryGetGare
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = Geral.Documento.IdDocto
        
        Set rstModulo = .OpenResultset(rdOpenStatic)
        
        If .rdoParameters(0).Value <> 0 Then
            MsgBox "Não foi possível ler dados referentes ao GARE.", vbInformation + vbOKOnly, App.Title
            Exit Function
        End If
        
        If Not rstModulo.EOF() Then
            'Atualiza dados de entrada
            If rstModulo!vecto <> 0 Then
                datVencimento.Text = Format(Right(rstModulo!vecto, 2) + Mid(rstModulo!vecto, 5, 2) + Left(rstModulo!vecto, 4), "00/00/0000")
            End If
            txtReceita.Text = rstModulo!Receita
            txtInscEstadual.Text = IIf(rstModulo!InscricaoEstadual = 0, "", rstModulo!InscricaoEstadual)
            txtCGCCPF.Text = IIf(rstModulo!CPFCGC = 0, "", rstModulo!CPFCGC)
            txtDividaAtiva.Text = IIf(rstModulo!DividaAtiva = 0, "", rstModulo!DividaAtiva)
            txtReferencia.Text = IIf(rstModulo!Referencia = 0, "", Right(CStr(rstModulo!Referencia), 2) + Left(CStr(rstModulo!Referencia), 4))
            txtAIIM.Text = IIf(rstModulo!AIIM = 0, "", rstModulo!AIIM)
            txtValorReceita.Text = (rstModulo!ValorReceita * 100)
            TxtJuros.Text = (rstModulo!Juros * 100)
            txtMulta.Text = (rstModulo!Multa * 100)
            txtAcrescimo.Text = (rstModulo!Acrescimo * 100)
            txtHonorarios.Text = (rstModulo!Honorarios * 100)
            txtValor.Text = (rstModulo!Valor * 100)
        Else
            'Atualiza Variáveis globais
            Geral.Documento.Leitura = ""
            Geral.Documento.ValorTotal = 0
            'Atualiza dados de entrada
            datVencimento.Text = ""
            txtReceita.Text = ""
            txtInscEstadual.Text = ""
            txtCGCCPF.Text = ""
            txtDividaAtiva.Text = ""
            txtReferencia.Text = ""
            txtAIIM.Text = ""
            txtValorReceita.Text = 0
            TxtJuros.Text = 0
            txtMulta.Text = 0
            txtAcrescimo.Text = 0
            txtHonorarios.Text = 0
            txtValor.Text = 0
        End If
    End With
    
    LerDadosGare = True

Exit_LerDadosGare:
    Set rstModulo = Nothing
    Exit Function
    
Err_LerDadosGare:

    Select Case TratamentoErro("Não foi possível ler dados referentes ao GARE !", Err, rdoErrors, False)
'        Case vbCancel
        Case vbRetry
            MsgBox "Não é possível repetir a operação!", vbInformation + vbOKOnly, App.Title
    End Select
    GoTo Exit_LerDadosGare

End Function

Private Sub MontaTela(sGrupo As String)

Dim rdoTipoCampo As rdoResultset
Dim sTipoCampo As String
Dim sGrupoAtual As String

    
    'Guarda Código do grupo atual
    sGrupoAtual = sGrupo
    
    'Se form chamador é diferente de Complementação, desabilita controle de entrada <> de Valores
    If AlteraValor Then
        datVencimento.TabStop = False
        datVencimento.Locked = True
        datVencimento.BackColor = G_ColorGray
        datVencimento.ForeColor = vbBlack
        
        txtReceita.TabStop = False
        txtReceita.Locked = True
        txtReceita.ForeColor = vbBlack
        txtReceita.BackColor = G_ColorGray
        ' sGrupo = "#"
    End If
    
     
    ' Ler os Campos e os Tipos
    With Modulo.qryGetTipoCampoGare
        .rdoParameters(0).Value = "*"  ' Todos os Campos
        .rdoParameters(1).Value = sGrupoAtual
        Set rdoTipoCampo = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
     

     ' Verifica se [ Inscricao_Estadual ] Campo deve ser ativado
     If Not rdoTipoCampo.EOF Then
        sTipoCampo = rdoTipoCampo!Inscricao_Estadual
     Else
        sTipoCampo = "0"
     End If
     
     txtInscEstadual.TabStop = IIf(sTipoCampo = "0", False, True)
     txtInscEstadual.Locked = (Not txtInscEstadual.TabStop)
     txtInscEstadual.BackColor = IIf(sTipoCampo <> "0", vbWhite, G_ColorGray)
     txtInscEstadual.ForeColor = IIf(sTipoCampo <> "0", G_ColorBlue, vbBlack)

         
    
     ' Verifica se [ Campo_Docto ] Campo deve ser ativado
     If Not rdoTipoCampo.EOF Then
        sTipoCampo = rdoTipoCampo!Campo_Docto
     Else
        sTipoCampo = "0"
     End If
     
    
     txtCGCCPF.TabStop = IIf(sTipoCampo = "0", False, True)
     txtCGCCPF.Locked = (Not txtCGCCPF.TabStop)
     txtCGCCPF.BackColor = IIf(sTipoCampo <> "0", vbWhite, G_ColorGray)
     txtCGCCPF.ForeColor = IIf(sTipoCampo <> "0", G_ColorBlue, vbBlack)
    
        
     ' Verifica se [ Inscricao_Ativa ] Campo deve ser ativado
     If Not rdoTipoCampo.EOF Then
        sTipoCampo = rdoTipoCampo!Inscricao_Ativa
     Else
        sTipoCampo = "0"
     End If
           
     txtDividaAtiva.TabStop = IIf(sTipoCampo = "0", False, True)
     txtDividaAtiva.Locked = (Not txtDividaAtiva.TabStop)
     txtDividaAtiva.BackColor = IIf(sTipoCampo <> "0", vbWhite, G_ColorGray)
     txtDividaAtiva.ForeColor = IIf(sTipoCampo <> "0", G_ColorBlue, vbBlack)
    
     ' Verifica se [ Referencia ] Campo deve ser ativado
     If Not rdoTipoCampo.EOF Then
         sTipoCampo = rdoTipoCampo!Referencia
     Else
        sTipoCampo = "0"
     End If
    
        
     txtReferencia.TabStop = IIf(sTipoCampo = "0", False, True)
     txtReferencia.Locked = (Not txtReferencia.TabStop)
     txtReferencia.BackColor = IIf(sTipoCampo <> "0", vbWhite, G_ColorGray)
     txtReferencia.ForeColor = IIf(sTipoCampo <> "0", G_ColorBlue, vbBlack)

    
     ' Verifica se [ Numero_Parcelamento ] Campo deve ser ativado
     If Not rdoTipoCampo.EOF Then
         sTipoCampo = rdoTipoCampo!Numero_Parcelamento
     Else
        sTipoCampo = "0"
     End If
    
     txtAIIM.TabStop = IIf(sTipoCampo = "0", False, True)
     txtAIIM.Locked = (Not txtAIIM.TabStop)
     txtAIIM.BackColor = IIf(sTipoCampo <> "0", vbWhite, G_ColorGray)
     txtAIIM.ForeColor = IIf(sTipoCampo <> "0", G_ColorBlue, vbBlack)
    
     ' Verifica se [ Valor_Receita ] Campo deve ser ativado
     If Not rdoTipoCampo.EOF Then
        sTipoCampo = rdoTipoCampo!Valor_Receita
     Else
        sTipoCampo = "0"
     End If
     
            
     txtValorReceita.Enabled = IIf(sTipoCampo = "0", False, True)
     txtValorReceita.BackColor = IIf(sTipoCampo <> "0", vbWhite, G_ColorGray)
     txtValorReceita.ForeColor = IIf(sTipoCampo <> "0", G_ColorBlue, vbBlack)
    
     ' Verifica se [ Valor_Juros ] Campo deve ser ativado
     If Not rdoTipoCampo.EOF Then
        sTipoCampo = rdoTipoCampo!Valor_Juros
     Else
        sTipoCampo = "0"
     End If
     
    
     TxtJuros.Enabled = IIf(sTipoCampo = "0", False, True)
     TxtJuros.BackColor = IIf(sTipoCampo <> "0", vbWhite, G_ColorGray)
     TxtJuros.ForeColor = IIf(sTipoCampo <> "0", G_ColorBlue, vbBlack)
    
     ' Verifica se [ Valor_Multa ] Campo deve ser ativado
     If Not rdoTipoCampo.EOF Then
        sTipoCampo = rdoTipoCampo!Valor_Multa
     Else
        sTipoCampo = "0"
     End If
     
    
     txtMulta.Enabled = IIf(sTipoCampo = "0", False, True)
     txtMulta.BackColor = IIf(sTipoCampo <> "0", vbWhite, G_ColorGray)
     txtMulta.ForeColor = IIf(sTipoCampo <> "0", G_ColorBlue, vbBlack)
     
     ' Verifica se [ Valor_Acrescimo_Financ ] Campo deve ser ativado
     If Not rdoTipoCampo.EOF Then
        sTipoCampo = rdoTipoCampo!Valor_Acrescimo_Financ
     Else
        sTipoCampo = "0"
     End If
     
    
     txtAcrescimo.Enabled = IIf(sTipoCampo = "0", False, True)
     txtAcrescimo.BackColor = IIf(sTipoCampo <> "0", vbWhite, G_ColorGray)
     txtAcrescimo.ForeColor = IIf(sTipoCampo <> "0", G_ColorBlue, vbBlack)
     
     ' Verifica se [ Valor_Honor_Advoc ] Campo deve ser ativado
     If Not rdoTipoCampo.EOF Then
        sTipoCampo = rdoTipoCampo!Valor_Honor_Advoc
     Else
        sTipoCampo = "0"
     End If
    
     txtHonorarios.Enabled = IIf(sTipoCampo = "0", False, True)
     txtHonorarios.BackColor = IIf(sTipoCampo <> "0", vbWhite, G_ColorGray)
     txtHonorarios.ForeColor = IIf(sTipoCampo <> "0", G_ColorBlue, vbBlack)
    
     txtValor.Enabled = True
     txtValor.BackColor = vbWhite
     txtValor.ForeColor = G_ColorBlue
    
     'Verifica se Opção de Desabilitar controles
     If sGrupo = "#" Then Exit Sub
    
     If AlteraValor Then
        If Not txtInscEstadual.TabStop And Val(txtInscEstadual.Text) = 0 Then txtInscEstadual.Text = ""
        If Not txtCGCCPF.TabStop And Val(txtCGCCPF.Text) = 0 Then txtCGCCPF.Text = ""
        If Not txtDividaAtiva.TabStop And Val(txtDividaAtiva.Text) = 0 Then txtDividaAtiva.Text = ""
        If Not txtReferencia.TabStop And Val(txtReferencia) = 0 Then txtReferencia.Text = ""
        If Not txtAIIM.TabStop And Val(txtAIIM) = 0 Then txtAIIM.Text = ""
        Exit Sub
     End If
    
     If Not txtInscEstadual.TabStop Then txtInscEstadual.Text = ""
     If Not txtCGCCPF.TabStop Then txtCGCCPF.Text = ""
     If Not txtDividaAtiva.TabStop Then txtDividaAtiva.Text = ""
     If Not txtReferencia.TabStop Then txtReferencia.Text = ""
     If Not txtAIIM.TabStop Then txtAIIM.Text = ""
    
     If Not txtValorReceita.Enabled Then txtValorReceita.Text = 0
     If Not TxtJuros.Enabled Then TxtJuros.Text = 0
     If Not txtMulta.Enabled Then txtMulta.Text = 0
     If Not txtAcrescimo.Enabled Then txtAcrescimo.Text = 0
     If Not txtHonorarios.Enabled Then txtHonorarios.Text = 0
     If Not txtValor.Enabled Then txtValor.Text = 0

End Sub

Private Sub txtValorReceita_LostFocus()
            
    txtValor.Text = Val(Desformata_Valor(txtValorReceita.Text)) + Val(Desformata_Valor(TxtJuros.Text)) + Val(Desformata_Valor(txtMulta.Text)) + Val(Desformata_Valor(txtAcrescimo.Text)) + Val(Desformata_Valor(txtHonorarios.Text))

End Sub

Private Function TipoCampoGare(ByVal sCampo As String, ByVal sGrupo As String) As String

Dim rdoTipoCampo As rdoResultset
Dim sTipoCampo As String

' Ler o Tipo do Campo

With Modulo.qryGetTipoCampoGare
    .rdoParameters(0).Value = Trim(sCampo)
    .rdoParameters(1).Value = sGrupo
    Set rdoTipoCampo = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
End With


sTipoCampo = rdoTipoCampo(0).Value

TipoCampoGare = sTipoCampo

End Function


Private Function AtivaCampos(ByVal sCampo As String, ByVal sTipo As String) As String

Dim CampoObrig As rdoResultset
Dim sGrupoNaoAbilitado As String

' Verifica se Campo é Obrigatório

        With Modulo.qryGetAtivaCampoGare
            .rdoParameters(0).Value = Trim(sCampo)
            .rdoParameters(1).Value = sTipo ' 0 - Não obrigatório  1-Obrigatório
            Set CampoObrig = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
        End With
    
    CampoObrig.MoveFirst
    While Not CampoObrig.EOF
      sGrupoNaoAbilitado = Trim(sGrupoNaoAbilitado) & CampoObrig!CodigoGrupo
      CampoObrig.MoveNext
    Wend
    
    AtivaCampos = sGrupoNaoAbilitado

End Function






Public Sub ImportaTFSCDARF()
 
    
    Dim RdoGare             As rdoResultset
    Dim DatFile             As Integer
    Dim lRetorno            As Long
    Dim Reg                 As String * 43
    Dim OffSet              As Long
    Dim sStr                As String
    Dim sTitulo             As String
    Dim sWhere              As String
    Dim strNomeArquivo      As String
    Dim lngRegs             As Long
    
 
    On Error GoTo ErroLeitura

    sTitulo = "Carga Tabela Gare"
     
     
     strNomeArquivo = Trim("M:\Arquivos\Ilco\TFSCDARF.DAT") ' Arquivo Carga Diaria Grupo Gare Aut. Digital
     
     If Not FileExist(strNomeArquivo) Then
        Beep
        MsgBox "Arquivo " & strNomeArquivo & " Não Encontrado", vbExclamation, sTitulo
        Exit Sub
     End If
     
     Screen.MousePointer = vbHourglass
     
     DatFile = FreeFile
          
     Open strNomeArquivo For Binary Access Read Lock Read Write As #DatFile
              
     OffSet = 1
      
     Get #DatFile, OffSet, Reg
          
     lngRegs = 0
          
                 
           ' Se arquivo foi lido ok
            If Len(Reg) < 14 Then
                MsgBox "Erro de Leitura", vbOKOnly + vbCritical, sTitulo
                GoTo ErroLeitura
            End If
              
                      
                     
          While Not EOF(DatFile)
          
              'Acumulador de registros lidos
              lngRegs = lngRegs + 1
              
              
              ' Atribui campos do arquivo
              RegGareCDarf.CodigoPagamento = Trim(Mid(Reg, 1, 4))
              RegGareCDarf.CodigoGrupo = Trim(Mid(Reg, 5, 1))
              RegGareCDarf.Data_Inicio_Vigencia = Trim(Mid(Reg, 6, 8))
              RegGareCDarf.Data_Final_Vigencia = Trim(Mid(Reg, 14, 8))
              RegGareCDarf.Indicador_Excessao = Trim(Mid(Reg, 22, 1))
              RegGareCDarf.Indicador_Arrecadacao = Trim(Mid(Reg, 23, 1))
              RegGareCDarf.Tipo_Servico = Trim(Mid(Reg, 24, 1))
              RegGareCDarf.Indicador_Autenticacao = Trim(Mid(Reg, 25, 1))
              RegGareCDarf.Indicador_Servico_Autenticacao = Trim(Mid(Reg, 26, 3))
              RegGareCDarf.Numero_Vias_Comprovante = Trim(Mid(Reg, 29, 2))
              RegGareCDarf.Valor = Trim(Mid(Reg, 31, 12))
                            
              ' Grava Registro na Tabela TFSCDARF
              
              ' Validar código da receita digitado
                With Modulo.qryInserirTFSCDARF
                    
                    
                    .rdoParameters(1).Value = CInt(RegGareCDarf.CodigoPagamento)
                    .rdoParameters(2).Value = RegGareCDarf.CodigoGrupo
                    .rdoParameters(3).Value = RegGareCDarf.Data_Inicio_Vigencia
                    .rdoParameters(4).Value = RegGareCDarf.Data_Final_Vigencia
                    .rdoParameters(5).Value = RegGareCDarf.Indicador_Excessao
                    .rdoParameters(6).Value = RegGareCDarf.Indicador_Arrecadacao
                    .rdoParameters(7).Value = RegGareCDarf.Tipo_Servico
                    .rdoParameters(8).Value = RegGareCDarf.Indicador_Autenticacao
                    .rdoParameters(9).Value = RegGareCDarf.Indicador_Servico_Autenticacao
                    .rdoParameters(10).Value = RegGareCDarf.Numero_Vias_Comprovante
                    .rdoParameters(11).Value = RegGareCDarf.Valor
                    Set RdoGare = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
                    
                End With
              
              
              
              OffSet = OffSet + Len(Reg)
              
              Get #DatFile, OffSet, Reg
              
                              
          Wend
      
          Close #DatFile
          
               
         
     Screen.MousePointer = vbDefault
     MsgBox vbCrLf & CStr(lngRegs) & " Registros(s) Foram Processados ", vbOKOnly + vbInformation, sTitulo
     
     
     
     Exit Sub
     
FimLeitura:
     
     Close #DatFile
          
     Screen.MousePointer = vbDefault
     Exit Sub
    
FimLeituraComErro:

     Close #DatFile
     GoTo FimLeitura
     
ErroLeitura:
     Beep
     Screen.MousePointer = vbDefault
     
     'Tratamento para diretório não localizado
     If Err.Number = 76 Then
          MsgBox "Diretório do Arquivo Não Localizado." & vbCrLf & vbCrLf & _
                    "Favor verificar diretório em parâmetros.", vbCritical, sTitulo
          GoTo FimLeitura
     End If
     If Err.Number = 53 Then
          MsgBox "Arquivo Não Localizado." & vbCrLf & vbCrLf & _
                    "Favor verificar e tentar novamente.", vbCritical, sTitulo
          GoTo FimLeitura
     End If
     If Err.Number = 55 Then
          MsgBox "Arquivo em utilização por outro usuário. Favor verificar!", vbCritical, sTitulo
          GoTo FimLeitura
     End If

    MsgBox "Erro na Leitura do Arquivo.", vbOKOnly + vbCritical, sTitulo

    GoTo FimLeitura

End Sub




''''''''''''''''''''''''''''''''''''''''''''''
' Leitura da Tabela TFSCRGPS.DAT
'''''''''''''''''''''''''''''''''''''''''''''

Public Sub ImportaTFSCRGPS()
 
    
    Dim RdoGare             As rdoResultset
    Dim DatFile             As Integer
    Dim lRetorno            As Long
    Dim Reg                 As String * 14
    Dim OffSet              As Long
    Dim sStr                As String
    Dim sTitulo             As String
    Dim sWhere              As String
    Dim strNomeArquivo      As String
    Dim lngRegs             As Long
    
 
    On Error GoTo ErroLeitura

    sTitulo = "Carga Tabela Gare"
     
     
     strNomeArquivo = Trim("M:\Arquivos\Ilco\TFSCRGPS.DAT") ' Arquivo Carga Diaria Grupo Gare Aut. Digital
     
     If Not FileExist(strNomeArquivo) Then
        Beep
        MsgBox "Arquivo " & strNomeArquivo & " Não Encontrado", vbExclamation, sTitulo
        Exit Sub
     End If
     
     Screen.MousePointer = vbHourglass
     
     DatFile = FreeFile
          
     Open strNomeArquivo For Binary Access Read Lock Read Write As #DatFile
              
     OffSet = 1
      
     ' Get #DatFile, OffSet, Reg
     Get #DatFile, OffSet, RegGareCRGPS
          
     lngRegs = 0
          
                 
           ' Se arquivo foi lido ok
            If Len(Reg) < 14 Then
                MsgBox "Erro de Leitura", vbOKOnly + vbCritical, sTitulo
                GoTo ErroLeitura
            End If
              
                      
            
          
          While Not EOF(DatFile)
          
              'Acumulador de registros lidos
              'lngRegs = lngRegs + 1
              
              
              ' Atribui campos do arquivo
'              RegGareCRGPS.CodigoGrupoReceita = Trim(Mid(Reg, 1, 1))
'              RegGareCRGPS.IndicadorCotaIPVA = Trim(Mid(Reg, 2, 1))
'              RegGareCRGPS.IndicadorVenctoNormal = Trim(Mid(Reg, 3, 1))
'              RegGareCRGPS.IndicadorInscEstadual = Trim(Mid(Reg, 4, 1))
'              RegGareCRGPS.IndicadorCampoDocto = Trim(Mid(Reg, 5, 1))
'              RegGareCRGPS.IndicadorInscAtiva = Trim(Mid(Reg, 6, 1))
'              RegGareCRGPS.IndicadorReferencia = Trim(Mid(Reg, 7, 1))
'              RegGareCRGPS.IndicadorNumParcelamento = Trim(Mid(Reg, 8, 1))
'              RegGareCRGPS.IndicadorValorReceita = Trim(Mid(Reg, 9, 1))
'              RegGareCRGPS.IndicadorValorJuro = Trim(Mid(Reg, 10, 1))
'              RegGareCRGPS.IndicadorValorMulta = Trim(Mid(Reg, 11, 1))
'              RegGareCRGPS.IndicadorAcresFinanceiro = Trim(Mid(Reg, 12, 1))
'              RegGareCRGPS.IndicadorHonoAdvogado = Trim(Mid(Reg, 13, 1))
              
                            
              ' Grava Registro na Tabela TFSCRGPS
              
              ' Validar código da receita digitado
                With Modulo.qryInserirTFSCRGPS
                    
                    
                    .rdoParameters(1).Value = RegGareCRGPS.CodigoGrupoReceita
                    .rdoParameters(2).Value = RegGareCRGPS.IndicadorCotaIPVA
                    .rdoParameters(3).Value = RegGareCRGPS.IndicadorVenctoNormal
                    .rdoParameters(4).Value = RegGareCRGPS.IndicadorInscEstadual
                    .rdoParameters(5).Value = RegGareCRGPS.IndicadorCampoDocto
                    .rdoParameters(6).Value = RegGareCRGPS.IndicadorInscAtiva
                    .rdoParameters(7).Value = RegGareCRGPS.IndicadorReferencia
                    .rdoParameters(8).Value = RegGareCRGPS.IndicadorNumParcelamento
                    .rdoParameters(9).Value = RegGareCRGPS.IndicadorValorReceita
                    .rdoParameters(10).Value = RegGareCRGPS.IndicadorValorJuro
                    .rdoParameters(11).Value = RegGareCRGPS.IndicadorValorMulta
                    .rdoParameters(12).Value = RegGareCRGPS.IndicadorAcresFinanceiro
                    .rdoParameters(13).Value = RegGareCRGPS.IndicadorHonoAdvogado
                    
                    
                    Set RdoGare = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
                    
                End With
              
              
              
              ' OffSet = OffSet + Len(Reg)
              OffSet = OffSet + 1
              
              ' Get #DatFile, OffSet, Reg
              Get #DatFile, OffSet, RegGareCRGPS
              
                              
          Wend
      
          Close #DatFile
          
               
         
     Screen.MousePointer = vbDefault
     MsgBox vbCrLf & CStr(lngRegs) & " Registros(s) Foram Processados ", vbOKOnly + vbInformation, sTitulo
     
     
     
     Exit Sub
     
FimLeitura:
     
     Close #DatFile
          
     Screen.MousePointer = vbDefault
     Exit Sub
    
FimLeituraComErro:

     Close #DatFile
     GoTo FimLeitura
     
ErroLeitura:
     Beep
     Screen.MousePointer = vbDefault
     
     'Tratamento para diretório não localizado
     If Err.Number = 76 Then
          MsgBox "Diretório do Arquivo Não Localizado." & vbCrLf & vbCrLf & _
                    "Favor verificar diretório em parâmetros.", vbCritical, sTitulo
          GoTo FimLeitura
     End If
     If Err.Number = 53 Then
          MsgBox "Arquivo Não Localizado." & vbCrLf & vbCrLf & _
                    "Favor verificar e tentar novamente.", vbCritical, sTitulo
          GoTo FimLeitura
     End If
     If Err.Number = 55 Then
          MsgBox "Arquivo em utilização por outro usuário. Favor verificar!", vbCritical, sTitulo
          GoTo FimLeitura
     End If

    MsgBox "Erro na Leitura do Arquivo.", vbOKOnly + vbCritical, sTitulo

    GoTo FimLeitura

End Sub


Private Function Valida_DSI(ByVal sNumDSI As String) As Boolean

Dim soma As Integer
Dim resto As Integer
Dim digito_11 As Integer
Dim digito_rv As String
Dim sNumeroDSI As String
    

soma = 0
resto = 0
digito_11 = 0       'calculado pelo módulo 11
digito_rv = ""      'caracter digitado pelo operador


    ' DSI AANNNNNNND
    sNumeroDSI = Format(sNumDSI, "0000000000")

    If Len(sNumeroDSI) <> 10 Then
        Valida_DSI = False
        Exit Function
    End If

    soma = 4 * 3
    soma = soma + Mid(sNumeroDSI, 1, 1) * 2
    soma = soma + Mid(sNumeroDSI, 2, 1) * 9
    soma = soma + Mid(sNumeroDSI, 3, 1) * 8
    soma = soma + Mid(sNumeroDSI, 4, 1) * 7
    soma = soma + Mid(sNumeroDSI, 5, 1) * 6
    soma = soma + Mid(sNumeroDSI, 6, 1) * 5
    soma = soma + Mid(sNumeroDSI, 7, 1) * 4
    soma = soma + Mid(sNumeroDSI, 8, 1) * 3
    soma = soma + Mid(sNumeroDSI, 9, 1) * 2

    resto = soma Mod 11        'resto da divisão
    digito_11 = 11 - resto     'digito verificador

    '*** se o calculo for igual a 10 ou 11, muda-se para 0 ***
    If (digito_11 >= 10) Then
         digito_11 = 0
    End If

    digito_rv = Mid(sNumeroDSI, 10, 1)    'digito verificador
    
    If CStr(digito_11) <> (digito_rv) Then
        Valida_DSI = False                  'digito não confere
    Else
        Valida_DSI = True                   'digito OK
    End If
    
    
End Function


Private Function Valida_DI(ByVal sNumDI As String) As Boolean

Dim soma As Integer
Dim resto As Integer
Dim digito_11 As Integer
Dim digito_rv As String
Dim sNumeroDI As String

soma = 0
resto = 0
digito_11 = 0       'calculado pelo módulo 11
digito_rv = ""      'caracter digitado pelo operador


    ' DSI AANNNNNNND
    sNumeroDI = Format(sNumDI, "0000000000")

    If Len(sNumeroDI) <> 10 Then
        Valida_DI = False
        Exit Function
    End If

    soma = 2 * 3
    soma = soma + Mid(sNumeroDI, 1, 1) * 2
    soma = soma + Mid(sNumeroDI, 2, 1) * 9
    soma = soma + Mid(sNumeroDI, 3, 1) * 8
    soma = soma + Mid(sNumeroDI, 4, 1) * 7
    soma = soma + Mid(sNumeroDI, 5, 1) * 6
    soma = soma + Mid(sNumeroDI, 6, 1) * 5
    soma = soma + Mid(sNumeroDI, 7, 1) * 4
    soma = soma + Mid(sNumeroDI, 8, 1) * 3
    soma = soma + Mid(sNumeroDI, 9, 1) * 2

    resto = soma Mod 11        'resto da divisão
    digito_11 = 11 - resto     'digito verificador

    '*** se o calculo for igual a 10 ou 11, muda-se para 0 ***
    If (digito_11 >= 10) Then
         digito_11 = 0
    End If

    digito_rv = Mid(sNumeroDI, 10, 1)    'digito verificador
    
    If CStr(digito_11) <> (digito_rv) Then
        Valida_DI = False                  'digito não confere
    Else
        Valida_DI = True                  'digito OK
    End If
    
    
End Function

Private Function Val_Notificacao(ByVal sNumNotifc As String) As Boolean

Dim soma As Integer
Dim resto As Integer
Dim digito_11 As Integer
Dim digito_rv As String
Dim sNumNotificacao As String
    
soma = 0
resto = 0
digito_11 = 0       'calculado pelo módulo 11
digito_rv = ""      'caracter digitado pelo operador


    ' DSI AANNNNNNND
    sNumNotificacao = Format(sNumNotifc, "000000000")

    If Len(sNumNotificacao) <> 9 Then
        Val_Notificacao = False
        Exit Function
    End If

    soma = soma + Mid(sNumNotificacao, 1, 1) * 8
    soma = soma + Mid(sNumNotificacao, 2, 1) * 7
    soma = soma + Mid(sNumNotificacao, 3, 1) * 6
    soma = soma + Mid(sNumNotificacao, 4, 1) * 5
    soma = soma + Mid(sNumNotificacao, 5, 1) * 4
    soma = soma + Mid(sNumNotificacao, 6, 1) * 3
    soma = soma + Mid(sNumNotificacao, 7, 1) * 2
    soma = soma + Mid(sNumNotificacao, 8, 1) * 10

    resto = soma Mod 11        'resto da divisão
    digito_11 = 11 - resto     'digito verificador

    '*** se o calculo for igual a 10 muda-se para 0 ***
    If (digito_11 = 10) Then
         digito_11 = 0
    End If

    digito_rv = Mid(sNumNotificacao, 9, 1)    'digito verificador
    
    If (CStr(digito_11) <> (digito_rv)) Or (CStr(digito_11) = "11") Then
        Val_Notificacao = False               'digito não confere
    Else
         Val_Notificacao = True               'digito OK
    End If
    
    
End Function




Private Function Val_Renavam(ByVal sNumRenavam As String) As Boolean

Dim soma As Integer
Dim resto As Integer
Dim digito_11 As Integer
Dim digito_rv As String
Dim strRenavam As String
    
soma = 0
resto = 0
digito_11 = 0       'calculado pelo módulo 11
digito_rv = ""      'caracter digitado pelo operador


    ' REVAVAM NNNNNNNND
    strRenavam = Format(sNumRenavam, "000000000")

    If Len(strRenavam) <> 9 Then
        Val_Renavam = False
        Exit Function
    End If

    soma = soma + Mid(strRenavam, 1, 1) * 9
    soma = soma + Mid(strRenavam, 2, 1) * 8
    soma = soma + Mid(strRenavam, 3, 1) * 7
    soma = soma + Mid(strRenavam, 4, 1) * 6
    soma = soma + Mid(strRenavam, 5, 1) * 5
    soma = soma + Mid(strRenavam, 6, 1) * 4
    soma = soma + Mid(strRenavam, 7, 1) * 3
    soma = soma + Mid(strRenavam, 8, 1) * 2

    resto = soma Mod 11        'resto da divisão
    digito_11 = 11 - resto     'digito verificador

    '*** se o calculo for igual a 10 muda-se para 0 ***
    If (digito_11 = 10) Or (digito_11 = 11) Then
         digito_11 = 0
    End If

    digito_rv = Mid(strRenavam, 9, 1)    'digito verificador
    
    If (CStr(digito_11) <> (digito_rv)) Then
        Val_Renavam = False               'digito não confere
    Else
        Val_Renavam = True               'digito OK
    End If
    
    
End Function






Private Function Val_Municipio(ByVal sMunicipio As String) As Boolean

Dim soma As Integer
Dim resto As Integer
Dim digito_11 As Integer
Dim digito_rv As String
Dim strMunicipio As String
    
soma = 0
resto = 0
digito_11 = 0       'calculado pelo módulo 11
digito_rv = ""      'caracter digitado pelo operador


    ' Municipio MMMD
    strMunicipio = Format(sMunicipio, "0000")

     
    If Val(strMunicipio) = 0 Then
        Val_Municipio = False
        Exit Function
    End If

    If Len(strMunicipio) <> 4 Then
        Val_Municipio = False
        Exit Function
    End If

    soma = soma + Mid(strMunicipio, 1, 1) * 4
    soma = soma + Mid(strMunicipio, 2, 1) * 3
    soma = soma + Mid(strMunicipio, 3, 1) * 2
    
    resto = soma Mod 11        'resto da divisão
   
    digito_11 = Mid(CStr(resto), Len(CStr(resto)), 1)    'digito verificador

    
    digito_rv = Mid(strMunicipio, 4, 1)    'digito verificador
    
    If (CStr(digito_11) <> (digito_rv)) Then
        Val_Municipio = False               'digito não confere
    Else
        Val_Municipio = True               'digito OK
    End If
    
    
End Function






Private Function Val_Declaracao(ByVal sDeclaracao As String) As Boolean

Dim soma As Integer
Dim resto As Integer
Dim digito_11 As Integer
Dim digito_rv As String
Dim strDeclaracao As String
    
soma = 0
resto = 0
digito_11 = 0       'calculado pelo módulo 11
digito_rv = ""      'caracter digitado pelo operador


   
    If Len(sDeclaracao) <> 9 Then
        Val_Declaracao = False
        Exit Function
    End If

     ' Municipio MMMD
    strDeclaracao = Format(sDeclaracao, "000000000")


    soma = soma + Mid(strDeclaracao, 1, 1) * 1
    soma = soma + Mid(strDeclaracao, 2, 1) * 3
    soma = soma + Mid(strDeclaracao, 3, 1) * 4
    soma = soma + Mid(strDeclaracao, 4, 1) * 5
    soma = soma + Mid(strDeclaracao, 5, 1) * 6
    soma = soma + Mid(strDeclaracao, 6, 1) * 7
    soma = soma + Mid(strDeclaracao, 7, 1) * 8
    soma = soma + Mid(strDeclaracao, 8, 1) * 10
    
    resto = soma Mod 11        'resto da divisão
   
    digito_11 = Mid(CStr(resto), Len(CStr(resto)), 1)    'digito verificador

    
    digito_rv = Mid(strDeclaracao, 9, 1)    'digito verificador
    
    If (CStr(digito_11) <> (digito_rv)) Then
        Val_Declaracao = False               'digito não confere
    Else
        Val_Declaracao = True               'digito OK
    End If
    
    
End Function
Private Function Val_Inscricao(ByVal sInscricao As String) As Boolean

    Dim soma As Integer
    Dim resto As Integer
    Dim digito_11 As Integer
    Dim digito_rv As String
    Dim strInscricao As String

    soma = 0
    resto = 0
    digito_11 = 0       'calculado pelo módulo 11
    digito_rv = ""      'caracter digitado pelo operador

    If Len(sInscricao) <> 12 Then
        Val_Inscricao = False
        Exit Function
    End If

    ' Formata Inscrição
    strInscricao = Format(sInscricao, "000000000000")

    If Mid(strInscricao, 1, 1) = "0" Then
         '*************************************************************
         ' Calculo do Primeiro Digito Nona Posição
         ' número da Inscricao: (8+1)            0 M M M S S S S D 0 0 0
         '                                       x x x x x x x x x
         ' multiplica da esquerda para direita:  1 3 4 5 6 7 8 10
         '*************************************************************
        
         soma = soma + Mid(txtInscEstadual, 1, 1) * 1
         soma = soma + Mid(txtInscEstadual, 2, 1) * 3
         soma = soma + Mid(txtInscEstadual, 3, 1) * 4
         soma = soma + Mid(txtInscEstadual, 4, 1) * 5
         soma = soma + Mid(txtInscEstadual, 5, 1) * 6
         soma = soma + Mid(txtInscEstadual, 6, 1) * 7
         soma = soma + Mid(txtInscEstadual, 7, 1) * 8
         soma = soma + Mid(txtInscEstadual, 8, 1) * 10
        
         resto = soma Mod 11                     'resto da divisão
         digito_11 = Right(str(resto), 1)        'digito verificador
         digito_rv = Mid(strInscricao, 9, 1)     'digito verificador informado
         
         If (CStr(digito_11) <> (digito_rv)) Then
            Val_Inscricao = False               'digito não confere
         Else
            Val_Inscricao = True                'digito OK
         End If
     
     ElseIf Mid(strInscricao, 1, 1) <> "0" Then
         '*************************************************************
         ' Calculo do Primeiro Digito Nona Posição
         ' número da Inscricao: (8+1)            0 M M M S S S S D 0 0 0
         '                                       x x x x x x x x x
         ' multiplica da esquerda para direita:  1 3 4 5 6 7 8 10
         '*************************************************************
        
         soma = soma + Mid(txtInscEstadual, 1, 1) * 1
         soma = soma + Mid(txtInscEstadual, 2, 1) * 3
         soma = soma + Mid(txtInscEstadual, 3, 1) * 4
         soma = soma + Mid(txtInscEstadual, 4, 1) * 5
         soma = soma + Mid(txtInscEstadual, 5, 1) * 6
         soma = soma + Mid(txtInscEstadual, 6, 1) * 7
         soma = soma + Mid(txtInscEstadual, 7, 1) * 8
         soma = soma + Mid(txtInscEstadual, 8, 1) * 10
        
         resto = soma Mod 11                     'resto da divisão
         digito_11 = Right(str(resto), 1)        'digito verificador
         digito_rv = Mid(strInscricao, 9, 1)     'digito verificador informado
         
         If (CStr(digito_11) <> (digito_rv)) Then
            Val_Inscricao = False               'digito não confere
            Exit Function
         Else
            Val_Inscricao = True                'digito OK
         End If
       
       
        '*************************************************************
        ' Calculo do Segundo Digito Décima Segunda Posição
        'número da inscricao: (11+1)          M M M  S S S S S D N N D
        '                                     x x x  x x x x x x x x
        'multiplica da direita para esquerda: 3 2 10 9 8 7 6 5 4 3 2
        '*************************************************************
        soma = 0
        resto = 0
        digito_11 = 0        'calculado pelo módulo 11
        digito_rv = ""       'caracter digitado pelo operador

        soma = soma + Mid(txtInscEstadual, 1, 1) * 3
        soma = soma + Mid(txtInscEstadual, 2, 1) * 2
        soma = soma + Mid(txtInscEstadual, 3, 1) * 10
        soma = soma + Mid(txtInscEstadual, 4, 1) * 9
        soma = soma + Mid(txtInscEstadual, 5, 1) * 8
        soma = soma + Mid(txtInscEstadual, 6, 1) * 7
        soma = soma + Mid(txtInscEstadual, 7, 1) * 6
        soma = soma + Mid(txtInscEstadual, 8, 1) * 5
        soma = soma + Mid(txtInscEstadual, 9, 1) * 4
        soma = soma + Mid(txtInscEstadual, 10, 1) * 3
        soma = soma + Mid(txtInscEstadual, 11, 1) * 2
        
        resto = soma Mod 11                     'resto da divisão
        digito_11 = Right(str(resto), 1)        'digito verificador
        digito_rv = Mid(strInscricao, 12, 1)    'digito verificador
        
        If (CStr(digito_11) <> (digito_rv)) Then
            Val_Inscricao = False               'digito não confere
        Else
            Val_Inscricao = True                'digito OK
        End If
     
     End If
End Function







'Public Function OLD_ValidaInscricao() As Byte
'
'
'    Dim sGruposAtivos As String
'    Dim lValidaGrupo As Boolean
'    Dim sTipoCampo As String
'
'
'    Dim soma As Integer
'    Dim resto As Integer
'    Dim digito_11 As Integer
'    Dim digito_rv As String
'    Dim bOk As Byte
'    Dim sStr            As String
'
'    bOk = True          'default - ok
'
'    Erro = 0
'    soma = 0
'    resto = 0
'    digito_11 = 0       'calculado pelo módulo 11
'    digito_rv = ""      'caracter digitado pelo operador
'
'    ' Carrega os grupos  para Inscricao_Estadual
'    ' marcados como '2'
'    ' na tabela TFSCRGPS
'
'    sGruposAtivos = AtivaCampos("Inscricao_Estadual", "2")
'    lValidaGrupo = IIf(InStr(sGruposAtivos, GrupoReceita) = 0, False, True)
'
'
'    sTipoCampo = TipoCampoGare("Inscricao_Estadual", GrupoReceita)
'
'
'
'
'    Select Case sTipoCampo
'
'        Case 0
'            ' Campo não está habilitado
'             ValidaInscricao = True
'             Exit Function
'        Case 2
'            ' Valida Código do Municipio
'            txtInscEstadual = Format(txtInscEstadual, "0000")
'
'            If Not Val_Municipio(txtInscEstadual) Then
'                ValidaInscricao = bOk
'                Erro = 2
'                Exit Function
'            Else
'                ValidaInscricao = bOk
'                Exit Function
'            End If
'        Case 3
'            ' Valida Número da Declaração
'            txtInscEstadual = Format(txtInscEstadual, "000000000")
'
'            If Not Val_Declaracao(txtInscEstadual) Then
'                ValidaInscricao = bOk
'                Erro = 5
'                Exit Function
'            Else
'                ValidaInscricao = bOk
'                Exit Function
'            End If
'
'    End Select
'
'
'
'    Select Case Len(txtInscEstadual)
''        Case 4
''            ' If GrupoReceita <> "J" And GrupoReceita <> "H" Then
''            If Not lValidaGrupo Then
''                ValidaInscricao = bOk
''                Erro = 3
''                Exit Function
''            End If
''            If ((Val(Mid(txtInscEstadual, 1, 3)) > 800) And (Val(Mid(txtInscEstadual, 1, 3)) < 900)) Or (Val(Mid(txtInscEstadual, 1, 3)) = 999) Then
''                Erro = 2
''            Else
''                soma = soma + Mid(txtInscEstadual, 1, 1) * 4
''                soma = soma + Mid(txtInscEstadual, 2, 1) * 3
''                soma = soma + Mid(txtInscEstadual, 3, 1) * 2
''                resto = soma Mod 11                     'resto da divisão
''                digito_11 = Right(str(resto), 1)        'digito verificador
''                digito_rv = Mid(txtInscEstadual, 4, 1)  'digito verificador
''                If digito_11 <> Val(digito_rv) Then
''                    bOk = False     'digito não confere
''                End If
''             End If
'        Case 8
'            txtInscEstadual = Format(txtInscEstadual, "000000000")
'
'        Case 12
'            'Verificando MMM
'            If Not (Val(Mid(txtInscEstadual.Text, 1, 3)) >= 100 And Val(Mid(txtInscEstadual.Text, 1, 3)) <= 794) Or _
'                   (Val(Mid(txtInscEstadual.Text, 1, 3)) >= 801 And Val(Mid(txtInscEstadual.Text, 1, 3)) <= 899) Or _
'                   (Val(Mid(txtInscEstadual.Text, 1, 3)) = 999) Then
'
'                Erro = 1
'
'            End If
'        Case Else
'
'            ' Carrega os grupos  para Inscricao_Estadual
'            ' marcados como '2'
'            ' na tabela TFSCRGPS
'
'            sGruposAtivos = AtivaCampos("Inscricao_Estadual", "4")
'            lValidaGrupo = IIf(InStr(sGruposAtivos, GrupoReceita) = 0, False, True)
'
'            ' If Val(txtInscEstadual) = 0 And Val(TxtCGCCPF) = 0 And (GrupoReceita = "C" Or GrupoReceita = "D" Or GrupoReceita = "N") Then
'            If Val(txtInscEstadual) = 0 And Val(txtCGCCPF) = 0 And lValidaGrupo Then
'                Erro = 4
'            Else
'                Erro = 1
'            End If
'
'    End Select
'
'    ''''''''''''''''''''''''''''''''''''''''
'    ' Caso tenha algum erro, sai da rotina '
'    ''''''''''''''''''''''''''''''''''''''''
'    If (Erro <> 0) Then
'        ValidaInscricao = bOk
'        Exit Function
'    End If
'
'    If Len(txtInscEstadual) = 4 Then
'        txtInscEstadual = Format(txtInscEstadual, "000000000000")
'        ValidaInscricao = bOk
'        Exit Function
'    End If
'
'    '*************************************************************
'    'número da Inscricao: (8+1)            0 M M M S S S S D 0 0 0
'    '                                      x x x x x x x x x
'    'multiplica da esquerda para direita:  1 3 4 5 6 7 8 10
'    '*************************************************************
'
'    soma = soma + Mid(txtInscEstadual, 1, 1) * 1
'    soma = soma + Mid(txtInscEstadual, 2, 1) * 3
'    soma = soma + Mid(txtInscEstadual, 3, 1) * 4
'    soma = soma + Mid(txtInscEstadual, 4, 1) * 5
'    soma = soma + Mid(txtInscEstadual, 5, 1) * 6
'    soma = soma + Mid(txtInscEstadual, 6, 1) * 7
'    soma = soma + Mid(txtInscEstadual, 7, 1) * 8
'    soma = soma + Mid(txtInscEstadual, 8, 1) * 10
'
'    resto = soma Mod 11                     'resto da divisão
'    digito_11 = Right(str(resto), 1)        'digito verificador
'    digito_rv = Mid(txtInscEstadual, 9, 1)  'digito verificador
'
'    If CStr(digito_11) <> (digito_rv) Then
'        bOk = False                         'digito não confere
'        ValidaInscricao = bOk
'        Exit Function
'    End If
'
'    If (Len(txtInscEstadual) = 12) And (Mid(txtInscEstadual, 1, 1) <> "0") Then
'
'        soma = 0
'        resto = 0
'        digito_11 = 0        'calculado pelo módulo 11
'        digito_rv = ""       'caracter digitado pelo operador
'
'        '*************************************************************
'        'número da inscricao: (11+1)          M M M  S S S S S D N N D
'        '                                     x x x  x x x x x x x x
'        'multiplica da direita para esquerda: 3 2 10 9 8 7 6 5 4 3 2
'        '*************************************************************
'
'        soma = soma + Mid(txtInscEstadual, 1, 1) * 3
'        soma = soma + Mid(txtInscEstadual, 2, 1) * 2
'        soma = soma + Mid(txtInscEstadual, 3, 1) * 10
'        soma = soma + Mid(txtInscEstadual, 4, 1) * 9
'        soma = soma + Mid(txtInscEstadual, 5, 1) * 8
'        soma = soma + Mid(txtInscEstadual, 6, 1) * 7
'        soma = soma + Mid(txtInscEstadual, 7, 1) * 6
'        soma = soma + Mid(txtInscEstadual, 8, 1) * 5
'        soma = soma + Mid(txtInscEstadual, 9, 1) * 4
'        soma = soma + Mid(txtInscEstadual, 10, 1) * 3
'        soma = soma + Mid(txtInscEstadual, 11, 1) * 2
'
'        resto = soma Mod 11                     'resto da divisão
'        digito_11 = Right(str(resto), 1)        'digito verificador
'        digito_rv = Mid(txtInscEstadual, 12, 1) 'digito verificador
'
'        If CStr(digito_11) <> (digito_rv) Then
'            bOk = False                         'digito não confere
'            ValidaInscricao = bOk
'            Exit Function
'        End If
'    Else
'        ValidaInscricao = bOk
'    End If
'
'    ValidaInscricao = bOk
'
'End Function


'Private Function Old_VerificarCGCCPF() As Boolean
'Dim sGruposAtivos As String
'Dim sGruposInativos As String
'Dim sGruposDependentes As String
'Dim sTipoCampo As String
'Dim lValidaGrupo As Boolean
'
'
'    VerificarCGCCPF = False     'esta com erro
'
'    If VerificaReceita = False Then
'        Exit Function
'    End If
'
'
'    ' Carrega os grupos dependentes para Campo_Docto
'    ' Os grupos dependentes são os grupos: "C", "D", "N" marcados como '6'
'    ' na tabela TFSCRGPS
'
'    sGruposDependentes = AtivaCampos("Campo_Docto", "6")
'    lValidaGrupo = IIf(InStr(sGruposDependentes, GrupoReceita) = 0, False, True)
'
'
'    ' Ver Tipo do Campo do Saber se é Dependente
'    sTipoCampo = TipoCampoGare("Campo_Docto", GrupoReceita)
'
'    ' Campo não está habilitado
'    If sTipoCampo = "0" Then
'        VerificarCGCCPF = True
'        Exit Function
'    End If
'
'    ' Se for dependente e InscEstadual estiver preenchido não critica o Campo
'    If sTipoCampo = "6" And (Val(txtInscEstadual)) > 0 Then
'
'        If Val(txtCGCCPF) > 0 Then
'            MsgBox "Para o código de receita digitado o campo Inscrição Estadual ou CGC/CPF deve estar preenchido.", vbInformation + vbOKOnly, App.Title
'            txtInscEstadual_GotFocus
'            txtInscEstadual.SetFocus
'            Exit Function
'        Else
'            VerificarCGCCPF = True
'            Exit Function
'        End If
'    End If
'
'    If Val(txtCGCCPF) = 0 Then
'        'If Val(txtInscEstadual) > 0 And (GrupoReceita = "C" Or GrupoReceita = "D" Or GrupoReceita = "N") Then
'        If Val(txtInscEstadual) > 0 And lValidaGrupo Then
'            VerificarCGCCPF = True
'            Exit Function
'        End If
'        ' If Val(txtInscEstadual) = 0 And ((GrupoReceita = "C" Or GrupoReceita = "D" Or GrupoReceita = "N")) Then
'        If Val(txtInscEstadual) = 0 And lValidaGrupo Then
'            MsgBox "Para o código de receita digitado o campo Inscrição Estadual ou CGC/CPF deve estar preenchido.", vbInformation + vbOKOnly, App.Title
'            txtInscEstadual_GotFocus
'            txtInscEstadual.SetFocus
'            Exit Function
'        End If
'
'
'        ' Carrega os grupos ativos para Campo_Docto
'        ' Os grupos ativos são  marcados como '3'
'        ' na tabela TFSCRGPS
'
'        sGruposAtivos = AtivaCampos("Campo_Docto", "3")
'        lValidaGrupo = IIf(InStr(sGruposAtivos, GrupoReceita) = 0, False, True)
'
'
'        'Select Case GrupoReceita
'        'Case "G", "H", "I", "J"
'
'         If lValidaGrupo Then
'
'                MsgBox "O preenchimento do campo CGC/CPF é obrigatório para o Código de Receita digitado.", vbInformation + vbOKOnly, App.Title
'                txtCGCCPF_GotFocus
'                txtCGCCPF.SetFocus
'                Exit Function
'         Else
'         'Case Else
'                If Val(txtCGCCPF) > 0 Then
'                    MsgBox "Para o grupo da receita digitado, o campo CGC/CPF não pode estar preenchido.", vbInformation + vbOKOnly, App.Title
'                    txtCGCCPF_GotFocus
'                    txtCGCCPF.SetFocus
'                    Exit Function
'                End If
'        End If
'        'End Select
'    Else
'        '''''''''''''''''''''''''''''''''''''''
'        ' Verifica se foi digitado CGC ou CPF '
'        '''''''''''''''''''''''''''''''''''''''
'
'        ' Carrega os grupos dependentes para Campo_Docto
'        ' Os grupos dependentes são marcados como '6'
'        ' na tabela TFSCRGPS
'
'        sGruposDependentes = AtivaCampos("Campo_Docto", "6")
'        lValidaGrupo = IIf(InStr(sGruposDependentes, GrupoReceita) = 0, False, True)
'
'        ' If Val(txtInscEstadual) > 0 And (GrupoReceita = "C" Or GrupoReceita = "D" Or GrupoReceita = "N") Then
'        If (Val(txtInscEstadual) > 0) And lValidaGrupo Then
'            MsgBox "Para o grupo da receita digitado, o campo CGC/CPF não pode estar preenchido, pois já existe valor no campo Inscrição Estadual.", vbInformation + vbOKOnly, App.Title
'            txtCGCCPF_GotFocus
'            txtCGCCPF.SetFocus
'            Exit Function
'        End If
'
'        ' Carrega os grupos inativos para Campo_Docto
'        ' Os grupos inativos são marcados como '0' Zero
'        ' na tabela TFSCRGPS
'
'        sGruposInativos = AtivaCampos("Campo_Docto", "0")
'        lValidaGrupo = IIf(InStr(sGruposInativos, GrupoReceita) = 0, False, True)
'
'        'If GrupoReceita = "A" Or GrupoReceita = "B" Or GrupoReceita = "E" Or GrupoReceita = "F" Then
'        If lValidaGrupo Then
'            MsgBox "Para o grupo da receita digitado, o campo CGC/CPF não pode estar preenchido.", vbInformation + vbOKOnly, App.Title
'            txtCGCCPF_GotFocus
'            txtCGCCPF.SetFocus
'            Exit Function
'        End If
'
'        If (Len(txtCGCCPF) = 9 And sTipoCampo = "3") Or (Len(txtCGCCPF) = 9 And sTipoCampo = "5") Then
'
'            If Not Val_Renavam(txtCGCCPF) Then
'                MsgBox "Digito do Renavam não confere. Tente novamente.", vbInformation + vbOKOnly, App.Title
'                txtCGCCPF_GotFocus
'                txtCGCCPF.SetFocus
'                Exit Function
'            Else
'                VerificarCGCCPF = True
'                Exit Function
'            End If
'
'        End If
'
'        If Len(txtCGCCPF) > 11 Then
'            txtCGCCPF = Format(txtCGCCPF, "000000000000000")    '15 posicoes
'            If Not VerificaCGC(txtCGCCPF) Then
'                MsgBox "Digito do campo CGC/CPF não confere. Tente novamente.", vbInformation + vbOKOnly, App.Title
'                txtCGCCPF_GotFocus
'                txtCGCCPF.SetFocus
'                Exit Function
'            End If
'        Else
'            txtCGCCPF = Format(txtCGCCPF, "00000000000")        '11 posicoes
'            If Not VerificaCPF(txtCGCCPF) Then
'                MsgBox "Digito do campo CGC/CPF não confere. Tente novamente.", vbInformation + vbOKOnly, App.Title
'                txtCGCCPF_GotFocus
'                txtCGCCPF.SetFocus
'                Exit Function
'            End If
'        End If
'    End If
'
'    VerificarCGCCPF = True      'esta OK
'
'End Function













'Private Sub BackUP_MontaTela(sGrupo As String)
'Dim sGrupoNaoAbilitado As String
'Dim sGrupoAtual As String
'
'
'    'Guarda Código do grupo atual
'    sGrupoAtual = sGrupo
'
'    'Se form chamador é diferente de Complementação, desabilita controle de entrada <> de Valores
'    If AlteraValor Then
'        datVencimento.TabStop = False
'        datVencimento.Locked = True
'        datVencimento.BackColor = G_ColorGray
'        datVencimento.ForeColor = vbBlack
'
'        txtReceita.TabStop = False
'        txtReceita.Locked = True
'        txtReceita.ForeColor = vbBlack
'        txtReceita.BackColor = G_ColorGray
'        ' sGrupo = "#"
'    End If
'
'
'
'
'
''    txtInscEstadual.TabStop = IIf(InStr("ABCDEFHJN", sGrupo) = 0, False, True)
''    txtInscEstadual.Locked = (Not txtInscEstadual.TabStop)
''    txtInscEstadual.BackColor = IIf(InStr("ABCDEFHJN", sGrupo) = 0, G_ColorGray, vbWhite)
''    txtInscEstadual.ForeColor = IIf(InStr("ABCDEFHJN", sGrupo) = 0, vbBlack, G_ColorBlue)
'
'     ' Verifica se Campo deve ser ativado
'     sGrupoNaoAbilitado = AtivaCampos("Inscricao_Estadual", "0")
'
'     txtInscEstadual.TabStop = IIf(InStr(sGrupoNaoAbilitado, sGrupo) = 0, True, False)
'     txtInscEstadual.Locked = (Not txtInscEstadual.TabStop)
'     txtInscEstadual.BackColor = IIf(InStr(sGrupoNaoAbilitado, sGrupo) = 0, vbWhite, G_ColorGray)
'     txtInscEstadual.ForeColor = IIf(InStr(sGrupoNaoAbilitado, sGrupo) = 0, G_ColorBlue, vbBlack)
'
'
''    txtCGCCPF.TabStop = IIf(InStr("CDGHIJN", sGrupo) = 0, False, True)
''    txtCGCCPF.Locked = (Not txtCGCCPF.TabStop)
''    txtCGCCPF.BackColor = IIf(InStr("CDGHIJN", sGrupo) = 0, G_ColorGray, vbWhite)
''    txtCGCCPF.ForeColor = IIf(InStr("CDGHIJN", sGrupo) = 0, vbBlack, G_ColorBlue)
'
'    ' Verifica se Campo deve ser ativado
'     sGrupoNaoAbilitado = AtivaCampos("Campo_Docto", "0")
'
'    txtCGCCPF.TabStop = IIf(InStr(sGrupoNaoAbilitado, sGrupo) = 0, True, False)
'    txtCGCCPF.Locked = (Not txtCGCCPF.TabStop)
'    txtCGCCPF.BackColor = IIf(InStr(sGrupoNaoAbilitado, sGrupo) = 0, vbWhite, G_ColorGray)
'    txtCGCCPF.ForeColor = IIf(InStr(sGrupoNaoAbilitado, sGrupo) = 0, G_ColorBlue, vbBlack)
'
'
''    txtDividaAtiva.TabStop = IIf(InStr("EIJ", sGrupo) = 0, False, True)
''    txtDividaAtiva.Locked = (Not txtDividaAtiva.TabStop)
''    txtDividaAtiva.BackColor = IIf(InStr("EIJ", sGrupo) = 0, G_ColorGray, vbWhite)
''    txtDividaAtiva.ForeColor = IIf(InStr("EIJ", sGrupo) = 0, vbBlack, G_ColorBlue)
'
'
'    ' Verifica se Campo deve ser ativado
'     sGrupoNaoAbilitado = AtivaCampos("Inscricao_Ativa", "0")
'
'    txtDividaAtiva.TabStop = IIf(InStr(sGrupoNaoAbilitado, sGrupo) = 0, True, False)
'    txtDividaAtiva.Locked = (Not txtDividaAtiva.TabStop)
'    txtDividaAtiva.BackColor = IIf(InStr(sGrupoNaoAbilitado, sGrupo) = 0, vbWhite, G_ColorGray)
'    txtDividaAtiva.ForeColor = IIf(InStr(sGrupoNaoAbilitado, sGrupo) = 0, G_ColorBlue, vbBlack)
'
''    txtReferencia.TabStop = IIf(InStr("A", sGrupo) = 0, False, True)
''    txtReferencia.Locked = (Not txtReferencia.TabStop)
''    txtReferencia.BackColor = IIf(InStr("A", sGrupo) = 0, G_ColorGray, vbWhite)
''    txtReferencia.ForeColor = IIf(InStr("A", sGrupo) = 0, vbBlack, G_ColorBlue)
''
'    ' Verifica se Campo deve ser ativado
'    sGrupoNaoAbilitado = AtivaCampos("Referencia", "0")
'
'    txtReferencia.TabStop = IIf(InStr(sGrupoNaoAbilitado, sGrupo) = 0, True, False)
'    txtReferencia.Locked = (Not txtReferencia.TabStop)
'    txtReferencia.BackColor = IIf(InStr(sGrupoNaoAbilitado, sGrupo) = 0, vbWhite, G_ColorGray)
'    txtReferencia.ForeColor = IIf(InStr(sGrupoNaoAbilitado, sGrupo) = 0, G_ColorBlue, vbBlack)
'
'
'
'
''    txtAIIM.TabStop = IIf(InStr("BCFHN", sGrupo) = 0, False, True)
''    txtAIIM.Locked = (Not txtAIIM.TabStop)
''    txtAIIM.BackColor = IIf(InStr("BCFHN", sGrupo) = 0, G_ColorGray, vbWhite)
''    txtAIIM.ForeColor = IIf(InStr("BCFHN", sGrupo) = 0, vbBlack, G_ColorBlue)
'
'    ' Verifica se Campo deve ser ativado
'    sGrupoNaoAbilitado = AtivaCampos("Numero_Parcelamento", "0")
'
'    txtAIIM.TabStop = IIf(InStr(sGrupoNaoAbilitado, sGrupo) = 0, True, False)
'    txtAIIM.Locked = (Not txtAIIM.TabStop)
'    txtAIIM.BackColor = IIf(InStr(sGrupoNaoAbilitado, sGrupo) = 0, vbWhite, G_ColorGray)
'    txtAIIM.ForeColor = IIf(InStr(sGrupoNaoAbilitado, sGrupo) = 0, G_ColorBlue, vbBlack)
'
'    'Para Valores não importa o form chamador, habilita conforme Código do Grupo
'    sGrupo = sGrupoAtual
'
''    txtValorReceita.Enabled = IIf(InStr("ABCDEFGHIJN", sGrupo) = 0, False, True)
''    txtValorReceita.BackColor = IIf(InStr("ABCDEFGHIJN", sGrupo) = 0, G_ColorGray, vbWhite)
''    txtValorReceita.ForeColor = IIf(InStr("ABCDEFGHIJN", sGrupo) = 0, vbBlack, G_ColorBlue)
''
'
'    ' Verifica se Campo deve ser ativado
'    sGrupoNaoAbilitado = AtivaCampos("Valor_Receita", "0")
'
'    txtValorReceita.Enabled = IIf(InStr(sGrupoNaoAbilitado, sGrupo) = 0, True, False)
'    txtValorReceita.BackColor = IIf(InStr(sGrupoNaoAbilitado, sGrupo) = 0, vbWhite, G_ColorGray)
'    txtValorReceita.ForeColor = IIf(InStr(sGrupoNaoAbilitado, sGrupo) = 0, G_ColorBlue, vbBlack)
'
'
'
''    TxtJuros.Enabled = IIf(InStr("ABCDEFGHIJN", sGrupo) = 0, False, True)
''    TxtJuros.BackColor = IIf(InStr("ABCDEFGHIJN", sGrupo) = 0, G_ColorGray, vbWhite)
''    TxtJuros.ForeColor = IIf(InStr("ABCDEFGHIJN", sGrupo) = 0, vbBlack, G_ColorBlue)
''
'    ' Verifica se Campo deve ser ativado
'    sGrupoNaoAbilitado = AtivaCampos("Valor_Juros", "0")
'
'    TxtJuros.Enabled = IIf(InStr(sGrupoNaoAbilitado, sGrupo) = 0, True, False)
'    TxtJuros.BackColor = IIf(InStr(sGrupoNaoAbilitado, sGrupo) = 0, vbWhite, G_ColorGray)
'    TxtJuros.ForeColor = IIf(InStr(sGrupoNaoAbilitado, sGrupo) = 0, G_ColorBlue, vbBlack)
'
'
'
''    txtMulta.Enabled = IIf(InStr("ABCDEFGHIJN", sGrupo) = 0, False, True)
''    txtMulta.BackColor = IIf(InStr("ABCDEFGHIJN", sGrupo) = 0, G_ColorGray, vbWhite)
''    txtMulta.ForeColor = IIf(InStr("ABCDEFGHIJN", sGrupo) = 0, vbBlack, G_ColorBlue)
''
'
'     ' Verifica se Campo deve ser ativado
'    sGrupoNaoAbilitado = AtivaCampos("Valor_Multa", "0")
'
'    txtMulta.Enabled = IIf(InStr(sGrupoNaoAbilitado, sGrupo) = 0, True, False)
'    txtMulta.BackColor = IIf(InStr(sGrupoNaoAbilitado, sGrupo) = 0, vbWhite, G_ColorGray)
'    txtMulta.ForeColor = IIf(InStr(sGrupoNaoAbilitado, sGrupo) = 0, G_ColorBlue, vbBlack)
'
''    txtAcrescimo.Enabled = IIf(InStr("BF", sGrupo) = 0, False, True)
''    txtAcrescimo.BackColor = IIf(InStr("BF", sGrupo) = 0, G_ColorGray, vbWhite)
''    txtAcrescimo.ForeColor = IIf(InStr("BF", sGrupo) = 0, vbBlack, G_ColorBlue)
''
'    ' Verifica se Campo deve ser ativado
'    sGrupoNaoAbilitado = AtivaCampos("Valor_Acrescimo_Financ", "0")
'
'    txtAcrescimo.Enabled = IIf(InStr(sGrupoNaoAbilitado, sGrupo) = 0, True, False)
'    txtAcrescimo.BackColor = IIf(InStr(sGrupoNaoAbilitado, sGrupo) = 0, vbWhite, G_ColorGray)
'    txtAcrescimo.ForeColor = IIf(InStr(sGrupoNaoAbilitado, sGrupo) = 0, G_ColorBlue, vbBlack)
'
'
''    txtHonorarios.Enabled = IIf(InStr("EFGIJ", sGrupo) = 0, False, True)
''    txtHonorarios.BackColor = IIf(InStr("EFGIJ", sGrupo) = 0, G_ColorGray, vbWhite)
''    txtHonorarios.ForeColor = IIf(InStr("EFGIJ", sGrupo) = 0, vbBlack, G_ColorBlue)
''
'    ' Verifica se Campo deve ser ativado
'    sGrupoNaoAbilitado = AtivaCampos("Valor_Honor_Advoc", "0")
'
'    txtHonorarios.Enabled = IIf(InStr(sGrupoNaoAbilitado, sGrupo) = 0, True, False)
'    txtHonorarios.BackColor = IIf(InStr(sGrupoNaoAbilitado, sGrupo) = 0, vbWhite, G_ColorGray)
'    txtHonorarios.ForeColor = IIf(InStr(sGrupoNaoAbilitado, sGrupo) = 0, G_ColorBlue, vbBlack)
'
'
''    txtValor.Enabled = IIf(InStr("ABCDEFGHIJN", sGrupo) = 0, False, True)
''    txtValor.BackColor = IIf(InStr("ABCDEFGHIJN", sGrupo) = 0, G_ColorGray, vbWhite)
''    txtValor.ForeColor = IIf(InStr("ABCDEFGHIJN", sGrupo) = 0, vbBlack, G_ColorBlue)
'
'    txtValor.Enabled = True
'    txtValor.BackColor = vbWhite
'    txtValor.ForeColor = G_ColorBlue
'
'
'
'
'    'Verifica se Opção de Desabilitar controles
'    If sGrupo = "#" Then Exit Sub
'
'    If AlteraValor Then
'        If Not txtInscEstadual.TabStop And Val(txtInscEstadual.Text) = 0 Then txtInscEstadual.Text = ""
'        If Not txtCGCCPF.TabStop And Val(txtCGCCPF.Text) = 0 Then txtCGCCPF.Text = ""
'        If Not txtDividaAtiva.TabStop And Val(txtDividaAtiva.Text) = 0 Then txtDividaAtiva.Text = ""
'        If Not txtReferencia.TabStop And Val(txtReferencia) = 0 Then txtReferencia.Text = ""
'        If Not txtAIIM.TabStop And Val(txtAIIM) = 0 Then txtAIIM.Text = ""
'        Exit Sub
'    End If
'
'    If Not txtInscEstadual.TabStop Then txtInscEstadual.Text = ""
'    If Not txtCGCCPF.TabStop Then txtCGCCPF.Text = ""
'    If Not txtDividaAtiva.TabStop Then txtDividaAtiva.Text = ""
'    If Not txtReferencia.TabStop Then txtReferencia.Text = ""
'    If Not txtAIIM.TabStop Then txtAIIM.Text = ""
'
'    If Not txtValorReceita.Enabled Then txtValorReceita.Text = 0
'    If Not TxtJuros.Enabled Then TxtJuros.Text = 0
'    If Not txtMulta.Enabled Then txtMulta.Text = 0
'    If Not txtAcrescimo.Enabled Then txtAcrescimo.Text = 0
'    If Not txtHonorarios.Enabled Then txtHonorarios.Text = 0
'    If Not txtValor.Enabled Then txtValor.Text = 0
'
'End Sub











'Private Function BackUP_VerificarReferencia() As Boolean
'
'    Dim sGruposAtivos As String
'    Dim lValidaGrupo As Boolean
'
'    Dim ano_ref As Integer
'    Dim mes_ref As Integer
'    Dim mes_for As String   'Formata Data no padrão MM/AAAA'
'    Dim ano_for As String   'Formata Data no padrão MM/AAAA'
'
'    VerificarReferencia = False
'
'    'Verifica txtReferencia
'        If Mid(txtReferencia, 3, 1) = "/" Then
'            mes_for = Mid(txtReferencia, 1, 2)
'            ano_for = Mid(txtReferencia, 4, 4)
'            mes_ref = Mid(txtReferencia, 1, 2)
'            ano_ref = Mid(txtReferencia, 4, 4)
'        Else
'            mes_for = Mid(txtReferencia, 1, 2)
'            ano_for = Mid(txtReferencia, 3, 4)
'            mes_ref = Mid(txtReferencia, 1, 2)
'            ano_ref = Mid(txtReferencia, 3, 4)
'        End If
'
'    If VerificaReceita = False Then
'        Exit Function
'    End If
'
'    'A2_OK-286
'    If (Val(txtReferencia) = 0) Then
'        txtReferencia = "000000"
'    End If
'
'
'
'    ' Carrega os grupos ativos para Referencia
'    ' Os grupos ativos são  marcados como '1'
'    ' na tabela TFSCRGPS
'
'    sGruposAtivos = AtivaCampos("Referencia", "1")
'    lValidaGrupo = IIf(InStr(sGruposAtivos, GrupoReceita) = 0, False, True)
'
'
'    ' If (GrupoReceita <> "A") Then
'    If Not lValidaGrupo Then
'        If Val(txtReferencia) > 0 Then
'            MsgBox "Para o grupo da receita digitado, referência não pode estar preenchido.", vbInformation + vbOKOnly, App.Title
'        End If
'        txtReferencia = "000000"
'        VerificarReferencia = True
'        Exit Function
'    Else
'        If (Len(txtReferencia) < 6) Then
'            MsgBox "Digite a referencia (MMAAAA)!", vbInformation + vbOKOnly, App.Title
'            txtReferencia_GotFocus
'            txtReferencia.SetFocus
'            Exit Function
'        Else
'            '''''''''''''''''''''''''''''''''''''''''
'            ' Verifica se Referencia é válida MMAAAA'
'            '''''''''''''''''''''''''''''''''''''''''
'            If Not VerificaDataMMAAAA(txtReferencia) Then
'                MsgBox "Referência inválida. Digite novamente (MMAAAA).", vbInformation + vbOKOnly, App.Title
'                txtReferencia_GotFocus
'                txtReferencia.SetFocus
'                Exit Function
'            Else
'                '''''''''''''''''''''''''''''''''''''''''''''''''
'                ' Verifica se o ano é maior que o permitido     '
'                '''''''''''''''''''''''''''''''''''''''''''''''''
'                If ano_ref <> Val(Mid(datVencimento.Text, 5, 4)) Then
'                    If ano_ref > 2100 Then
'                        MsgBox "O ano informado é maior que máximo permitido. ", vbInformation + vbOKOnly, App.Title
'                        txtReferencia_GotFocus
'                        txtReferencia.SetFocus
'                        Exit Function
'                    End If
'                '''''''''''''''''''''''''''''''''''''''''''''''''
'                ' Verifica se o ano difere do ano do vencimento '
'                '''''''''''''''''''''''''''''''''''''''''''''''''
'                    If ano_ref < 1980 Then
'                        MsgBox "Este Ano é inválido para a Data de Vencimento Informada. Verifique e retorne.", vbInformation + vbOKOnly, App.Title
'                        txtReferencia_GotFocus
'                        txtReferencia.SetFocus
'                        Exit Function
'                    Else
'                        If ((Val(Mid(datVencimento.Text, 5, 4)) - ano_ref) > 7) And ((Val(txtReceita) = 607) Or _
'                           (Val(txtReceita) = 462) Or (Val(txtReceita) = 1466) Or (Val(txtReceita) = 1545)) Then
'                            MsgBox "Este Ano é inválido para o Código de Receita e para a Data de Vencimento informada. Verifique e retorne.", vbInformation + vbOKOnly, App.Title
'                            txtReferencia_GotFocus
'                            txtReferencia.SetFocus
'                            Exit Function
'                        End If
'                    End If
'                End If
'
'                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                ' Verifica se o mês de referencia é aceito no mês do vencimento '
'                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''                If (Val(txtReceita) = 607) And ((mes_ref - (Val(Mid(datVencimento.Text, 3, 2)))) > 5) Then
''                    MsgBox "Mês inválido para o Código de Receita e para a Data de Vencimento informada. Verifique e retorne.", vbInformation + vbOKOnly, App.Title
''                    txtReferencia_GotFocus
''                    txtReferencia.SetFocus
''                    Exit Function
''                End If
'                'para estes códigos de receita, o mês de referencia não ser igual
'                'ou maior do que o mes da data de vencimento, somente poderá ser inferior.
'                If ((Val(txtReceita) = 462) Or (Val(txtReceita) = 1466) Or _
'                   (Val(txtReceita) = 1545)) And (mes_ref >= (Val(Mid(datVencimento.Text, 3, 2)))) And (ano_ref = Val(Mid(datVencimento.Text, 5, 4))) Then
'                    MsgBox "Mês inválido para o Código de Receita e para a Data de Vencimento informada. Verifique e retorne. ", vbInformation + vbOKOnly, App.Title
'                    txtReferencia_GotFocus
'                    txtReferencia.SetFocus
'                    Exit Function
'                End If
'            End If
'        End If
'        txtReferencia = mes_for & "/" & ano_for
'        VerificarReferencia = True
'    End If
'
'End Function
'
