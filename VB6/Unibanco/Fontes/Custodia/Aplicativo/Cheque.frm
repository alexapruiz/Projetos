VERSION 5.00
Object = "{9CBA5D64-E3C8-11D3-9FFC-00104BC8688C}#1.0#0"; "CurrencyEdit.ocx"
Object = "{ED123F48-E23F-11D4-B08D-00600899AB13}#1.0#0"; "UbbEdit.ocx"
Begin VB.Form Cheque 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Complementação de Cheques:"
   ClientHeight    =   3660
   ClientLeft      =   1755
   ClientTop       =   3090
   ClientWidth     =   8925
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   8925
   Begin VB.CommandButton CmdScanner 
      Caption         =   "&Leitora"
      Height          =   750
      Left            =   5000
      Picture         =   "Cheque.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   500
      Visible         =   0   'False
      Width           =   850
   End
   Begin VB.CommandButton CmdCMC7 
      Caption         =   "CMC&7"
      Height          =   750
      Left            =   5000
      Picture         =   "Cheque.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   500
      Width           =   850
   End
   Begin VB.CommandButton cmdLinha1 
      Caption         =   "&Linha 1"
      Height          =   750
      Left            =   5860
      Picture         =   "Cheque.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   500
      Width           =   850
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   750
      Left            =   7580
      Picture         =   "Cheque.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   500
      Width           =   850
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "&Confirmar"
      Height          =   750
      Left            =   6720
      Picture         =   "Cheque.frx":0D60
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   500
      Width           =   850
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      Height          =   444
      Left            =   36
      ScaleHeight     =   390
      ScaleWidth      =   8805
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   20
      Width           =   8868
      Begin VB.Image ImgCheque 
         Height          =   480
         Left            =   210
         Picture         =   "Cheque.frx":106A
         Top             =   0
         Width           =   480
      End
      Begin VB.Label LblEsc 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[ESC] - Cancela"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   192
         Left            =   4860
         TabIndex        =   27
         Top             =   108
         Visible         =   0   'False
         Width           =   1344
      End
      Begin VB.Label LblMsg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cheques:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   792
         TabIndex        =   28
         Top             =   36
         Width           =   1140
      End
      Begin VB.Image ImgScanner 
         Height          =   480
         Left            =   180
         Picture         =   "Cheque.frx":14AC
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.Frame Frame_Valor 
      Caption         =   "Valor (R$)"
      Height          =   624
      Left            =   5532
      TabIndex        =   20
      Top             =   1320
      Width           =   2028
      Begin CURRENCYEDITLib.CurrencyEdit Valor 
         Height          =   360
         Left            =   48
         TabIndex        =   3
         Top             =   204
         Width           =   1836
         _Version        =   65537
         _ExtentX        =   3238
         _ExtentY        =   635
         _StockProps     =   93
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Data de Depósito"
      Height          =   624
      Left            =   972
      TabIndex        =   19
      Top             =   552
      Width           =   2028
      Begin VB.ComboBox CboDtaDeposito 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         ItemData        =   "Cheque.frx":17B6
         Left            =   60
         List            =   "Cheque.frx":17B8
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   216
         Width           =   1836
      End
   End
   Begin UbbEdt.UbbEdit CMC7_Campo1 
      Height          =   624
      Left            =   972
      TabIndex        =   6
      Top             =   2892
      Width           =   1164
      _ExtentX        =   2090
      _ExtentY        =   1191
      TextColor       =   0
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FieldType       =   11
      Title           =   "Campo 1"
   End
   Begin UbbEdt.UbbEdit CMC7_Campo2 
      Height          =   624
      Left            =   2280
      TabIndex        =   7
      Top             =   2892
      Width           =   1380
      _ExtentX        =   2461
      _ExtentY        =   1191
      TextColor       =   -2147483640
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FieldType       =   12
      TextMaxNumChars =   10
      Title           =   "Campo 2"
   End
   Begin UbbEdt.UbbEdit CMC7_Campo3 
      Height          =   624
      Left            =   3804
      TabIndex        =   8
      Top             =   2904
      Width           =   1596
      _ExtentX        =   2831
      _ExtentY        =   1191
      TextColor       =   -2147483640
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FieldType       =   13
      TextMaxNumChars =   12
      Title           =   "Campo 3"
   End
   Begin UbbEdt.UbbEdit Linha1_Comp 
      Height          =   624
      Left            =   972
      TabIndex        =   10
      Top             =   2280
      Visible         =   0   'False
      Width           =   624
      _ExtentX        =   1191
      _ExtentY        =   1191
      TextColor       =   -2147483640
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FieldType       =   14
      TextMaxNumChars =   3
      Title           =   "Comp."
   End
   Begin UbbEdt.UbbEdit Linha1_Bco 
      Height          =   624
      Left            =   1620
      TabIndex        =   11
      Top             =   2280
      Visible         =   0   'False
      Width           =   624
      _ExtentX        =   1191
      _ExtentY        =   1191
      TextColor       =   -2147483640
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FieldType       =   3
      TextMaxNumChars =   3
      Title           =   "Banco"
   End
   Begin UbbEdt.UbbEdit Linha1_C1 
      Height          =   624
      Left            =   3108
      TabIndex        =   13
      Top             =   2280
      Visible         =   0   'False
      Width           =   408
      _ExtentX        =   820
      _ExtentY        =   1191
      TextColor       =   -2147483640
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FieldType       =   8
      TextMaxNumChars =   1
      Title           =   "C1"
   End
   Begin UbbEdt.UbbEdit Linha1_Conta 
      Height          =   624
      Left            =   3540
      TabIndex        =   14
      Top             =   2280
      Visible         =   0   'False
      Width           =   1380
      _ExtentX        =   2461
      _ExtentY        =   1191
      TextColor       =   -2147483640
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FieldType       =   15
      TextMaxNumChars =   10
      Title           =   "Conta"
   End
   Begin UbbEdt.UbbEdit Linha1_C2 
      Height          =   624
      Left            =   4992
      TabIndex        =   15
      Top             =   2280
      Visible         =   0   'False
      Width           =   408
      _ExtentX        =   820
      _ExtentY        =   1191
      TextColor       =   -2147483640
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FieldType       =   9
      TextMaxNumChars =   1
      Title           =   "C2"
   End
   Begin UbbEdt.UbbEdit Linha1_Cheque 
      Height          =   624
      Left            =   5472
      TabIndex        =   16
      Top             =   2280
      Visible         =   0   'False
      Width           =   948
      _ExtentX        =   1720
      _ExtentY        =   1191
      TextColor       =   -2147483640
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FieldType       =   17
      TextMaxNumChars =   6
      Title           =   "Cheque"
   End
   Begin UbbEdt.UbbEdit Linha1_C3 
      Height          =   624
      Left            =   6480
      TabIndex        =   17
      Top             =   2280
      Visible         =   0   'False
      Width           =   408
      _ExtentX        =   820
      _ExtentY        =   1191
      TextColor       =   -2147483640
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FieldType       =   10
      TextMaxNumChars =   1
      Title           =   "C3"
   End
   Begin UbbEdt.UbbEdit Linha1_Tipo 
      Height          =   624
      Left            =   6996
      TabIndex        =   18
      Top             =   2280
      Visible         =   0   'False
      Width           =   480
      _ExtentX        =   873
      _ExtentY        =   1191
      TextColor       =   -2147483640
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextMaxNumChars =   1
      Title           =   "Tipo"
   End
   Begin VB.Frame Frame_Conf 
      Caption         =   "Confirmação:"
      Height          =   624
      Left            =   3096
      TabIndex        =   5
      Top             =   1320
      Visible         =   0   'False
      Width           =   2310
      Begin VB.TextBox CNPJ_Conf 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   72
         MaxLength       =   14
         TabIndex        =   2
         Top             =   192
         Width           =   2130
      End
   End
   Begin VB.Frame Frame_CNPJ 
      Caption         =   "CNPJ / CPF"
      Height          =   624
      Left            =   948
      TabIndex        =   4
      Top             =   1320
      Width           =   2028
      Begin VB.TextBox CNPJ 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   72
         MaxLength       =   14
         TabIndex        =   1
         Top             =   216
         Width           =   1836
      End
   End
   Begin UbbEdt.UbbEdit Linha1_Ag 
      Height          =   624
      Left            =   2316
      TabIndex        =   12
      Top             =   2280
      Visible         =   0   'False
      Width           =   744
      _ExtentX        =   1349
      _ExtentY        =   1191
      TextColor       =   -2147483640
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FieldType       =   4
      TextMaxNumChars =   4
      Title           =   "Agência"
   End
   Begin UbbEdt.UBBValid UBBValid1 
      Left            =   8190
      Top             =   1545
      _ExtentX        =   794
      _ExtentY        =   820
      Banco           =   409
      Campo1          =   "CMC7_Campo1"
      Campo2          =   "CMC7_Campo2"
      Campo3          =   "CMC7_Campo3"
      Campo4          =   "Linha1_Comp"
      Campo5          =   "Linha1_Bco"
      Campo6          =   "Linha1_Ag"
      Campo7          =   "Linha1_C1"
      Campo8          =   "Linha1_Conta"
      Campo9          =   "Linha1_C2"
      Campo10         =   "Linha1_Cheque"
      Campo11         =   "Linha1_C3"
   End
   Begin VB.Label LblInfo 
      AutoSize        =   -1  'True
      Caption         =   "CMC7-Digitação."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   972
      TabIndex        =   9
      Top             =   2004
      Width           =   1404
   End
End
Attribute VB_Name = "Cheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '                            '* Type de Utilização de Banco *'                              '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 Private Type Procedures
     Inclusao            As New Custodia.Inserir    'Querys de Insert
     Alteracao           As New Custodia.Atualizar  'Querys de Update
     Deletacao           As New Custodia.Excluir    'Querys de Delete
     Selecao             As New Custodia.Selecionar 'Querys de Select
 End Type

 Private Procedures      As Procedures
    
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '                                  '* Variáveis Auxiliares *'                               '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 Dim Linha1              As Boolean    'Tipo  de Complementação por CMC-7 ou Linha 1
 Dim CPF_CNPJ            As Double     'Valor da 1ª digitação do CPF/CNPJ
 Dim bInclusao           As Boolean    'Tipo  de Tratamento Inclusão / Alteração

 Dim i_idbordero         As Long       'Identificação do Bordero
 Dim i_idcheque          As Long       'Identificação do Cheque
 Dim d_DtaDeposito       As String     'Data de Depósito formatada DD/MM/AAAA
 Dim d_DtaProcessamento  As Long       'Data de Processamento Formatada DD/MM/AAAA
 Dim sDtaDeposito        As Long       'Data de Depósito

 Dim TipoInscricao          As Integer
 Dim m_RetornoCheque        As enumRetornoModal
 Dim TeclouEsc              As Boolean
 Dim CMC7_TOP_CNPJ_BOTTON   As Boolean
 Dim Info                   As String
 
'Constantes de Pesquisa(Combo)
 Const CB_FINDSTRING = &H14C
 Const CB_ERR = -1
 Const CB_SETCURSEL = &H14E
 Const LenDifCmd_Fram = 50

'Conteúdo de Objetos
 Const L100Info = "CMC7 - Captura, Scanner (L100-Ativa)."
 Const LA93Info = "CMC7 - Captura, Scanner (LA93-Ativa)."
 Const NuloInfo = "CMC7 - Digitação."
 Const MsgInclusao = "Inclusão de Cheques:"
 Const MsgAlteracao = "Alteração de Cheques:"
 
'Posições de objetos
 Const FrmChequeHeight = 3200
 Const LabelInfo = 2000
 Const TOP_CNPJ = 1300
 Const BOT_CMC7 = 2200
 Const TOP_CMC7 = 1300
 Const BOT_CNPJ = 2200
Sub InclusaoCheque()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                           * Inclusão de Cheques Na Base de Dados *                        '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

Dim rsIncluiCheque As New ADODB.Recordset
Dim RsIncluiOcorr  As New ADODB.Recordset

    Set rsIncluiCheque = g_cMainConnection.Execute _
                        (Procedures.Inclusao.InsereCheque(d_DtaProcessamento _
                        , i_idbordero _
                        , sDtaDeposito _
                        , CStr(CMC7_Campo1) & CStr(CMC7_Campo2) & CStr(CMC7_Campo3) _
                        , CNPJ.Text _
                        , TipoInscricao _
                        , Format(Val(InserePonto(Valor.Text)), MASK_VALOR)))
                        
    ''''''''''''''''''''''''''''
    ' * Recupera IddoBordero * '
    ''''''''''''''''''''''''''''
    Call IdChequeIncluso

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' * Verifica se cheque possui erro se possuir grava flag de erro * '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Val(Mid(CMC7_Campo1, 4, 4)) = 0 And Val(Mid(CMC7_Campo3, 5, 7)) = 0 Then
        
        Set RsIncluiOcorr = g_cMainConnection.Execute _
                          (Procedures.Alteracao.AtualizaErroCheque(d_DtaProcessamento _
                         , i_idbordero _
                         , i_idcheque))
    End If
    
    m_RetornoCheque = eRetornoOK

Exit Sub

TrataErro:
    Call TratamentoErro("Falha na Incluisão Cheque", Err)
    m_RetornoCheque = eRetornoCancelar
    Unload Me
    
End Sub
Public Sub FormataDataProcessamento()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                         '* Formatação da Data de Processamento *'                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    d_DtaProcessamento = Mid(Geral.DataProcessamento, 1, 4) & _
                         Mid(Geral.DataProcessamento, 5, 2) & _
                         Mid(Geral.DataProcessamento, 7, 2)

End Sub
Sub AlteracaoCheque()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                           * Alteracao de Cheques Na Base de Dados *                       '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

Dim rsAlteraCheque As New ADODB.Recordset
Dim RsIncluiOcorr  As New ADODB.Recordset

        Set rsAlteraCheque = g_cMainConnection.Execute _
                            (Procedures.Alteracao.AtualizaCheque(d_DtaProcessamento _
                                                               , i_idbordero _
                                                               , i_idcheque _
                                                               , sDtaDeposito _
                                                               , CStr(CMC7_Campo1) & CStr(CMC7_Campo2) & CStr(CMC7_Campo3) _
                                                               , CNPJ.Text _
                                                               , TipoInscricao _
                                                               , Format(Val(InserePonto(Valor.Text)), MASK_VALOR)))

                                                               
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' * Verifica se cheque possui erro se possuir grava flag de erro * '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Val(Mid(CMC7_Campo1, 4, 4)) = 0 And Val(Mid(CMC7_Campo3, 5, 7)) = 0 Then
        
        Set RsIncluiOcorr = g_cMainConnection.Execute _
                          (Procedures.Alteracao.AtualizaErroCheque(d_DtaProcessamento _
                         , i_idbordero _
                         , i_idcheque))
                         
    End If
    
    m_RetornoCheque = eRetornoOK
    
Exit Sub

TrataErro:
    Call TratamentoErro("Falha na atualização do cheque.", Err)
    m_RetornoCheque = eRetornoCancelar
    Unload Me
                                                                       
End Sub
Sub IdChequeIncluso()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'               * Objetivo Recuperar o último Idcheque do Bordero atual *                   '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

Dim rsidChequeIncluso As New ADODB.Recordset
    
        Set rsidChequeIncluso = g_cMainConnection.Execute _
                               (Procedures.Selecao.GetMaxidCheque(d_DtaProcessamento _
                                                                , i_idbordero))

        If Not rsidChequeIncluso.EOF Then
            '''''''''''''''''''''''''
            ' * Atualiza Idcheque * '
            '''''''''''''''''''''''''
            i_idcheque = rsidChequeIncluso!MaxidCheque
        
        End If
        
Exit Sub
TrataErro:
    Call TratamentoErro("Falha ao recuperar último IdCheque do Bordero.", Err)
    m_RetornoCheque = eRetornoCancelar
    Unload Me
        
End Sub
Function Valida_CMC7() As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                              * Verifica se CMC7 esta Correto *                            '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

Dim Objetos As Control

   'Verifica Tipificação
    If Mid(CMC7_Campo2.Text, 10, 1) <> 5 Then
        MsgBox "Documento Inválido.", vbExclamation + vbOKOnly, App.Title
        Valor.SetFocus
        Exit Function
    End If
        
    Valida_CMC7 = False
        
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    ' * Verifica cor das Fontes dos campos de CMC-7 * '
    ' * Recupera somente objetos do tipo "UbbEdit"  * '
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    For Each Objetos In Me.Controls
        If TypeName(Objetos) = "UbbEdit" Then
            If Mid(Objetos.Name, 1, 4) = "CMC7" Then
                If Objetos.TextColor = &HFF& Then
                    MsgBox "CMC7 inválido.", vbExclamation + vbOKOnly, App.Title
                    CMC7_Campo1.Text = ""
                    CMC7_Campo2.Text = ""
                    CMC7_Campo3.Text = ""
                    CMC7_Campo1.SetFocus
                    Exit Function
                End If
            End If
        End If
    Next Objetos

    Valida_CMC7 = True

Exit Function
TrataErro:
    Call TratamentoErro("Falha ao validar CMC7.", Err)
    m_RetornoCheque = eRetornoCancelar
    Unload Me
    
End Function
Private Function VerDuplicidadeCMC7() As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                    * Verifica se Cmc7 já existe na Base de Dados *                        '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

    Dim rsVerCMC7 As New ADODB.Recordset
    Dim sCMC7     As String
    
    VerDuplicidadeCMC7 = False
    
        '''''''''''''''''''
        ' * Agrupa CMC7 * '
        '''''''''''''''''''
        sCMC7 = CStr(CMC7_Campo1) & CStr(CMC7_Campo2) & CStr(CMC7_Campo3)
                
        ''''''''''''''''''''''''''''''
        ' * Pesquisa para Inclusão * '
        ''''''''''''''''''''''''''''''
        If bInclusao Then
            Set rsVerCMC7 = g_cMainConnection.Execute(Procedures.Selecao.GetPesquisaCMC7 _
                                                     (d_DtaProcessamento _
                                                    , sCMC7))
    
        ''''''''''''''''''''''''''''''
        ' * Pesquisa para Alteração * '
        ''''''''''''''''''''''''''''''
        Else
            Set rsVerCMC7 = g_cMainConnection.Execute(Procedures.Selecao.GetPesquisaCMC7 _
                                                     (d_DtaProcessamento _
                                                    , sCMC7 _
                                                    , i_idcheque))
        End If
        
        If Not rsVerCMC7.EOF Then
            If rsVerCMC7!CMC7 = sCMC7 Then
                MsgBox "Cheque já cadastrado.", vbExclamation + vbOKOnly, App.Title
                
                CMC7_Campo1.Text = ""
                CMC7_Campo2.Text = ""
                CMC7_Campo3.Text = ""
                
                CNPJ.SetFocus
                
                Exit Function
            End If
        End If

    VerDuplicidadeCMC7 = True

Exit Function
TrataErro:
    Call TratamentoErro("Falha ao verificar Duplicidade de CMC7.", Err)
    m_RetornoCheque = eRetornoCancelar
    Unload Me

End Function
Function Valida_Linha1() As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                          * Verifica se Linha 1 esta Correta *                            '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

Dim Objetos As Control
        
    Valida_Linha1 = False
        
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    ' * Verifica cor das Fontes dos campos de Linha1 '
    ' * Recupera somente objetos do tipo "UbbEdit"   '
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    For Each Objetos In Me.Controls
        If TypeName(Objetos) = "UbbEdit" Then
            If Mid(Objetos.Name, 1, 6) = "Linha1" Then
                If Objetos.TextColor = &HFF& Then
                    MsgBox "Linha 1 inválida.", vbExclamation + vbOKOnly, App.Title
                    Objetos.SetFocus
                    Exit Function
                End If
            End If
        End If
    Next Objetos

    Valida_Linha1 = True

Exit Function
TrataErro:
    Call TratamentoErro("Falha na validação da Linha 1.", Err)
    m_RetornoCheque = eRetornoCancelar
    Unload Me
    
End Function
Function ValidaCampos() As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                      * Verifica se todos os campos estão preenchidos *                    '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

Dim Objetos     As Control
Dim sCPF_CNPJ   As String
Dim CMC7        As New CalculoCheque
Dim sCMC7       As String
Dim Tipo, Dig1, Dig2, Dig3 As Long

    ValidaCampos = False
    '''''''''''''''''''''''''''''''''''''
    ' * Valida Combo Data de Depósito * '
    '''''''''''''''''''''''''''''''''''''
    If Len(Trim(CboDtaDeposito.Text)) = 0 Then
        MsgBox "Data de Depósito é obrigatória.", vbExclamation + vbOKOnly, App.Title
        CboDtaDeposito.SetFocus
        Exit Function
    Else
        sDtaDeposito = Mid(CboDtaDeposito, 7, 4) & Mid(CboDtaDeposito, 4, 2) & Mid(CboDtaDeposito, 1, 2)
    End If
    
   'Revalidação/Validação de CGC_CPF
    If CNPJ.Tag = False Then
        If Len(Trim(CNPJ.Text)) = 0 Or Val(CNPJ.Text) = 0 Then
            CNPJ.SetFocus
            MsgBox "CNPJ / CGC é obrigatório.", vbExclamation + vbOKOnly, App.Title
            Exit Function
        ElseIf Len(Trim(CNPJ.Text)) = 11 Then
            
            If VerificaCPF(CNPJ.Text) = False Then
                CPF_CNPJ = CNPJ.Text
                CNPJ.Text = ""
                Frame_Conf.Caption = "Confirmação de CPF."
                Frame_Conf.Visible = True
                TipoInscricao = 1
                CNPJ_Conf.SetFocus
                Exit Function
            Else
                TipoInscricao = 1
            End If
        ElseIf Len(Trim(CNPJ.Text)) = 14 Then
        
            sCPF_CNPJ = Format(CNPJ.Text, String(15, "0"))
            
            If VerificaCGC(sCPF_CNPJ) = False Then
                CPF_CNPJ = CNPJ.Text
                CNPJ.Text = ""
                Frame_Conf.Caption = "Confirmação de CNPJ."
                Frame_Conf.Visible = True
                TipoInscricao = 2
                CNPJ_Conf.SetFocus
                Exit Function
            Else
                TipoInscricao = 2
            End If
        Else
            MsgBox "CNPJ/CPF inválido.", vbExclamation + vbOKOnly, App.Title
            CNPJ.SetFocus
            Exit Function
        End If
    End If

    ''''''''''''''''''''''''''
    ' * Validação de Valor * '
    ''''''''''''''''''''''''''
    If Val(Valor.Text) = 0 Then
        MsgBox "Campo Valor é obrigatório.", vbExclamation + vbOKOnly, App.Title
        Valor.SetFocus
        Exit Function
    End If
    
    If Linha1 Then
        ''''''''''''''''''''''''''''''''''''''''''''''''''
        ' * Valida preenchimento dos campos da Linha 1 * '
        ' * Recupera somente objetos do tipo "UbbEdit" * '
        ''''''''''''''''''''''''''''''''''''''''''''''''''
        For Each Objetos In Me.Controls
            If TypeName(Objetos) = "UbbEdit" Then
                If Mid(Objetos.Name, 1, 6) = "Linha1" Then
                    If Len(Trim(Objetos.Text)) = 0 Then
                        MsgBox Objetos.Title & " é obrigatório.", vbExclamation + vbOKOnly, App.Title
                        Objetos.SetFocus
                        Exit Function
                    End If
                End If
            End If
        Next Objetos
    
        ''''''''''''''''''''''''''''''''''''''''
        ' * Verifica se Linha 1 esta Correta * '
        ''''''''''''''''''''''''''''''''''''''''
        If Valida_Linha1 = False Then
            Exit Function
        End If
        
        ''''''''''''''''''''''''''''''''''
        ' * Transformar Linha1 em CMC7 * '
        ''''''''''''''''''''''''''''''''''
         CMC7.Comp = Linha1_Comp.Text
         CMC7.Banco = Linha1_Bco.Text
         CMC7.Agencia = Linha1_Ag.Text
         CMC7.Conta = Linha1_Conta.Text
         CMC7.NumeroCheque = Linha1_Cheque.Text
         CMC7.Tipificacao = Linha1_Tipo '"5"
         
         '''''''''''''''''''''''''''''''''
         ' * Calcula / Retorna do CMC7 * '
         '''''''''''''''''''''''''''''''''
        If CMC7.Calcula Then
            sCMC7 = (CMC7.CMC7)
            CMC7_Campo1.Text = Mid(sCMC7, 1, 8)
            CMC7_Campo2.Text = Mid(sCMC7, 9, 10)
            CMC7_Campo3.Text = Mid(sCMC7, 19, 12)
        End If
     
    Else
        
        For Each Objetos In Me.Controls
            If TypeName(Objetos) = "UbbEdit" Then
                If Mid(Objetos.Name, 1, 4) = "CMC7" Then
                    If Len(Trim(Objetos.Text)) = 0 Then
                        MsgBox Objetos.Title & " é obrigatório.", vbExclamation + vbOKOnly, App.Title
                        Objetos.SetFocus
                        Exit Function
                    End If
                End If
            End If
        Next Objetos
    End If
    
    '''''''''''''''''''''''''''''''''''''
    ' * Verifica se Cmc7 esta Correto * '
    '''''''''''''''''''''''''''''''''''''
    If Valida_CMC7 = False Then
        Exit Function
    End If
    
    ValidaCampos = True

Exit Function
TrataErro:
    Call TratamentoErro("Falha na validação dos campos.", Err)
    m_RetornoCheque = eRetornoCancelar
    Unload Me
    
End Function
Public Sub SetIdCheque(ByVal pIdCheque As Long)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                   * Objetivo Recuperar Informações: Idcheque(opcional) *                  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

    If pIdCheque = 0 Then
       '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       ' * True  - Se idCheque for =  0 se trata de Inclusão * '
       '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        bInclusao = True
    Else
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       ' * False - Se idCheque for <> 0 se trata de Alteração * '
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        bInclusao = False
    End If
    
    i_idcheque = pIdCheque
    
Exit Sub
TrataErro:
    Call TratamentoErro("Falha na função SetIdCheque", Err)
    m_RetornoCheque = eRetornoCancelar
    Unload Me
    
End Sub
Public Function ShowModal(ByRef pIdBordero As Long, _
                          ByRef pDataDeposito As Long, _
                       Optional ByRef pIdCheque As Double, _
                       Optional ByVal PosFormTop As Long) As enumRetornoModal

'*** Coleta de Parâmetros e Retorno de Parâmetros ***
On Error GoTo TrataErro

    i_idbordero = pIdBordero
    d_DtaDeposito = Mid(pDataDeposito, 7, 2) & "/" & Mid(pDataDeposito, 5, 2) & "/" & Mid(pDataDeposito, 1, 4)
    sDtaDeposito = pDataDeposito
    
   'Posiciona o Form em local desejado
    If PosFormTop <> 0 Then Me.Top = PosFormTop
    
   'Inicio
    Me.Show vbModal
        
   'Alimenta variaveis p/ retorno
    pDataDeposito = sDtaDeposito
    pIdCheque = i_idcheque
    ShowModal = m_RetornoCheque

Exit Function
TrataErro:
    Call TratamentoErro("Falha função ShowModal.", Err)
    m_RetornoCheque = eRetornoCancelar
    Unload Me

End Function
Private Sub TratamentoInclusao()
'*** Tratamento para Inclusão de Cheque ***
 On Error GoTo TrataErro
   
    CboDtaDeposito.AddItem d_DtaDeposito
    CboDtaDeposito.Text = d_DtaDeposito
    CboDtaDeposito.Locked = False

 Exit Sub
TrataErro:
    Call TratamentoErro("Falha ao tratar inclusão.", Err)
    m_RetornoCheque = eRetornoCancelar
    Unload Me
    
End Sub
Private Sub TratamentoAlteração()
'*** Alteração de Dados do Cheque ***
On Error GoTo TrataErro
    
    Dim rsDatasDeposito    As New ADODB.Recordset
    Dim RsDadosCheque      As New ADODB.Recordset

   
   'Datas de Depósito
    Set rsDatasDeposito = g_cMainConnection.Execute(Procedures.Selecao.GetDatasBordero _
                                                   (d_DtaProcessamento, i_idbordero))
    
    If Not rsDatasDeposito.EOF Then
       'Preenche Combo de Datas de Depósito
        Do While Not rsDatasDeposito.EOF
            CboDtaDeposito.AddItem Mid(rsDatasDeposito!DataDeposito, 7, 2) & "/" & Mid(rsDatasDeposito!DataDeposito, 5, 2) & "/" & Mid(rsDatasDeposito!DataDeposito, 1, 4)
            rsDatasDeposito.MoveNext
        Loop
    
    End If
    
   'Dados do Cheque
    Set RsDadosCheque = g_cMainConnection.Execute(Procedures.Selecao.GetDadosCheque _
                                                 (d_DtaProcessamento, i_idbordero, i_idcheque))
    
    If Not RsDadosCheque.EOF Then
       'Data de Depósito
        CboDtaDeposito.Text = (Mid(RsDadosCheque!DataDeposito, 7, 2) & "/" & Mid(RsDadosCheque!DataDeposito, 5, 2) & "/" & Mid(RsDadosCheque!DataDeposito, 1, 4))
        
       'CMC7
        CMC7_Campo1.Text = Mid(RsDadosCheque!CMC7, 1, 8)
        CMC7_Campo2.Text = Mid(RsDadosCheque!CMC7, 9, 10)
        CMC7_Campo3.Text = Mid(RsDadosCheque!CMC7, 19, 12)
        
       'Valor
        'Valor.Text = Format(RsDadosCheque!Valor, "#,##0.00")
        Valor.Text = RsDadosCheque!Valor * 100
        
       'CNPJ/CPF
        CNPJ.Text = RsDadosCheque!CNPJCPF
        
    End If
    
Exit Sub
TrataErro:
    Call TratamentoErro("Falha ao tratar alteração.", Err)
    m_RetornoCheque = eRetornoCancelar
    Unload Me
    
End Sub
Private Sub CboDtaDeposito_KeyPress(KeyAscii As Integer)
   'Somente valor Numérico * '
    If KeyAscii = vbKeyReturn Then
      SendKeys "{TAB}"
    ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack And KeyAscii <> 47 Then
      KeyAscii = 0
    End If
              
    If KeyAscii = 13 Then
        If Len(Trim(CboDtaDeposito.Text)) = 0 Then Exit Sub
        
       'Verifica se valor apresentado é uma data * '
        If IsDate(CboDtaDeposito.Text) = False Then
            MsgBox "Data Inválida.", vbCritical + vbOKOnly, App.Title
            Exit Sub
        End If
    End If
        
End Sub

Private Sub CmdCMC7_Click()

On Error GoTo TrataErro

    Dim Objetos  As Control

   'Se scanner conectado desabilita, senão sem efeito
    If Principal.Scanner.HABILITADO Then
        CmdCMC7.Visible = False
        CmdScanner.Visible = True
        
       'Desabilita leitura, enquanto estiver nesta seção
        Principal.Scanner.Tag = False
    End If
    
    LblInfo.Caption = NuloInfo
   
   'Recupera somente objetos do tipo "UbbEdit" * '
    For Each Objetos In Me.Controls
        If TypeName(Objetos) = "UbbEdit" Then
            If Mid(Objetos.Name, 1, 4) = "CMC7" Then
                Objetos.Visible = True
            Else
                Objetos.Visible = False
            End If
        End If
    Next Objetos
    
    Linha1 = False
    
   'Inibi Confirmação de CPF/CNPJ
    Frame_Conf.Visible = False
    
   'Valor
   If Principal.Scanner.Scanner = eLA93 Then
        CMC7_Campo1.SetFocus
   Else
        Valor.SetFocus
   End If
    
Exit Sub
TrataErro:
    Call TratamentoErro("Falha ao escolher digitação por Cmc7.", Err)
    m_RetornoCheque = eRetornoCancelar
    Unload Me
    
End Sub
Private Sub cmdConfirmar_Click()

On Error GoTo TrataErro

    If ValidaCampos Then
      
            ''''''''''''''''''''''''''''''''''''
            ' * Verifica Duplicidade de CMC7 * '
            ''''''''''''''''''''''''''''''''''''
            If VerDuplicidadeCMC7 = False Then Exit Sub

            If bInclusao Then
                '''''''''''''''''''''''''''''''''''''''''''''
                ' * Rotina de Inclusão  de Cheque na Base * '
                '''''''''''''''''''''''''''''''''''''''''''''
                Call InclusaoCheque
            Else
                '''''''''''''''''''''''''''''''''''''''''''''
                ' * Rotina de Alteração de Cheque na Base * '
                '''''''''''''''''''''''''''''''''''''''''''''
                Call AlteracaoCheque
            End If
        
    Else
        Exit Sub
    End If
        
    Unload Me
    
Exit Sub
TrataErro:
    Call TratamentoErro("Falha na confirmação do Cheque.", Err)
    m_RetornoCheque = eRetornoCancelar
    Unload Me
    
End Sub
Private Sub CmdConfirmar_GotFocus()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'* Chama o Botão de Confirmação uma vez que o objeto UBBEDIT não contempla o evento KEYPRESS *'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    cmdConfirmar_Click
End Sub
Private Sub CmdScanner_Click()
    Dim Objetos  As Control
    
    On Error GoTo TrataErro
    
    CmdCMC7.Visible = True
    CmdScanner.Visible = False
    
   'Reabilita Leitura
    Principal.Scanner.Tag = True
    LblInfo.Caption = IIf(Principal.Scanner.Scanner = eL100, L100Info, LA93Info)
   
   'Recupera somente objetos do tipo "UbbEdit" * '
    For Each Objetos In Me.Controls
        If TypeName(Objetos) = "UbbEdit" Then
            If Mid(Objetos.Name, 1, 4) = "CMC7" Then
                Objetos.Visible = True
            Else
                Objetos.Visible = False
            End If
        End If
    Next Objetos
    
    Linha1 = False
    
   'Inibi Confirmação de CPF/CNPJ
    Frame_Conf.Visible = False
    
   'Valor
    If Principal.Scanner.Scanner = eLA93 Then
        CMC7_Campo1.SetFocus
    Else
        Valor.SetFocus
    End If
    
Exit Sub

TrataErro:
    Call TratamentoErro("Falha ao escolher Captura c/ Scanner.", Err)
    m_RetornoCheque = eRetornoCancelar
    Unload Me
    
End Sub
Private Sub cmdLinha1_Click()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'* Habilita campos da Linha 1 e Desabilita Campos de CMC-7 *'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

Dim Objetos  As Control
       
    LblInfo.Caption = "Linha1 - Digitação"
    
   'Recupera somente objetos do tipo "UbbEdit"
    For Each Objetos In Me.Controls
        If TypeName(Objetos) = "UbbEdit" Then
            If Mid(Objetos.Name, 1, 6) = "Linha1" Then
                Objetos.Visible = True
                Objetos.Text = ""
            Else
                Objetos.Visible = False
            End If
        End If
    Next Objetos

    Linha1 = True
    
   'Inibição da Confirmação de CPF
    Frame_Conf.Visible = False
    
   'Valor
    Valor.SetFocus
    
Exit Sub
TrataErro:
    Call TratamentoErro("Erro ao escolher digitação por Linha 1.", Err)
    m_RetornoCheque = eRetornoCancelar
    Unload Me

End Sub
Private Sub cmdSair_Click()
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '                            * Sai da Complementação de Cheques *                           '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    m_RetornoCheque = eRetornoCancelar
    
   'Se scanner for a LA93 ejeta o Cheque
    If Principal.Scanner.Scanner = eLA93 Then
        Principal.Scanner.Eject
    End If
    
    Unload Me
End Sub
Private Sub Form_Activate()

    If CMC7_TOP_CNPJ_BOTTON Then
        CMC7_Campo1.SetFocus
    End If
    
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
   'Atualiza Flag se usuario teclou (esc) durante leitura de cheque(L100)
    If ActiveControl.Name = "CMC7_Campo1" And KeyAscii = vbKeyEscape Then
        TeclouEsc = True
    End If
End Sub
Private Sub Form_Load()
    On Error GoTo TrataErro
    Me.Height = FrmChequeHeight
    
   'Formata Data de Processamento
    Call FormataDataProcessamento
      
   'Valor Default Combo
    CboDtaDeposito.Clear
    
   'Default é CMC7
    Linha1 = False
    
   'flag da segunda digitação
    CNPJ.Tag = False
    
    If bInclusao Then
        Info = MsgInclusao
        Principal.Scanner.Tag = True
        Call TratamentoInclusao
    Else
        Info = MsgAlteracao
        Principal.Scanner.Tag = False
        Call TratamentoAlteração
    End If
    
    LblMsg.Caption = Info
    LblInfo.Top = LabelInfo
    
    If Principal.Scanner.HABILITADO And Principal.Scanner.Scanner = eLA93 Then
        CMC7_TOP_CNPJ_BOTTON = True
        TopPosition False
        If Not bInclusao Then
            CmdCMC7.Visible = False
            CmdScanner.Visible = True
            Principal.Scanner.Tag = False
            LblInfo.Caption = NuloInfo
            
        Else
            LblInfo.Caption = LA93Info
        End If
        
    ElseIf Principal.Scanner.HABILITADO And Principal.Scanner.Scanner = eL100 Then
        CMC7_TOP_CNPJ_BOTTON = False
        TopPosition True
        
        If Not bInclusao Then
            CmdCMC7.Visible = False
            CmdScanner.Visible = True
            Principal.Scanner.Tag = False
            LblInfo.Caption = NuloInfo
        Else
            LblInfo.Caption = L100Info
            SendKeys "{TAB}"
        End If
        
    Else
       'Padrões digitação de CMC7 ou Linha1
        CMC7_TOP_CNPJ_BOTTON = False
        
       'Principal.Scanner.Tag = False
        TopPosition True
        LblInfo.Caption = NuloInfo
        SendKeys "{TAB}"
               
    End If
    
Exit Sub

TrataErro:
    Call TratamentoErro("Erro ao Inicializar Formulário.", Err)
    m_RetornoCheque = eRetornoCancelar
    Unload Me
End Sub
Private Sub Cmc7_campo1_GotFocus()
    '''''''''''''''''''''''''''''''''
    '* Habilita/Desabilita Leitura *'
    '''''''''''''''''''''''''''''''''
On Error GoTo Erro:

Dim Ret As enumRetornoLeitura
TeclouEsc = False
    
If Not CMC7_TOP_CNPJ_BOTTON Then
    If CBool(Principal.Scanner.Tag) And _
       Principal.Scanner.HABILITADO And _
       Trim(Valor.Text) <> "" And _
       Trim(CNPJ.Text) <> "" Then
RetryL100:
        Msg True
        Ret = Principal.Scanner.Le()
        Msg False
        
        If Ret = eLeituraOK Then
        
            CMC7_Campo1.Text = Principal.Scanner.CMC7_Campo1
            CMC7_Campo2.Text = Principal.Scanner.CMC7_Campo2
            CMC7_Campo3.Text = Principal.Scanner.CMC7_Campo3
            
            cmdConfirmar.SetFocus
            
        ElseIf (Ret = eLeituraEsc And TeclouEsc) Or Ret = eTimeOut Then
            Valor.SetFocus
        ElseIf (Ret = eLeituraEsc And Not TeclouEsc) Then
            GoTo RetryL100
        ElseIf Ret = eLeituraFalha Then
            If MsgBox("Falha na Leitura... Tentar Novamente ? ", vbCritical + vbYesNo + vbApplicationModal, App.Title) = vbYes Then
                GoTo RetryL100
            Else
                CmdCMC7_Click
            End If
        End If
        
    End If
Else
    If CBool(Principal.Scanner.Tag) And _
       Principal.Scanner.HABILITADO Then
RetryLA97:
        Msg True
        Ret = Principal.Scanner.Le()
        Msg False
        
        If Ret = eLeituraOK Then
        
            CMC7_Campo1.Text = Principal.Scanner.CMC7_Campo1
            CMC7_Campo2.Text = Principal.Scanner.CMC7_Campo2
            CMC7_Campo3.Text = Principal.Scanner.CMC7_Campo3
            
            CNPJ.SetFocus
            
        ElseIf Ret = eLeituraEsc And TeclouEsc Then
            Valor.SetFocus
        ElseIf (Ret = eLeituraEsc And Not TeclouEsc) Then
            GoTo RetryLA97
        ElseIf Ret = eLeituraFalha Then
            If MsgBox("Falha na Leitura... Tentar Novamente ? ", vbCritical + vbYesNo + vbApplicationModal, App.Title) = vbYes Then
                GoTo RetryLA97
            Else
                CmdCMC7_Click
            End If
        ElseIf Ret = eLeituraFim Then
            MsgBox "Alimentador do Scanner está vazio !", vbOKOnly + vbExclamation, App.Title
            cmdSair.SetFocus
        End If
    End If

End If

If Ret = eErro Then
    If Not Principal.Scanner.Erro Is Nothing Then
      Err.Raise Principal.Scanner.Erro.Number, App.Title, Principal.Scanner.Erro.Description
    End If
End If

Exit Sub

Erro:
    If Err = Principal.Scanner.Erro Then
        Call TratamentoErro("Falha no Módulo de Leitura.", Principal.Scanner.Erro, False, True)
    Else
        Call TratamentoErro("Erro no Módulo de Leitura.", Err, False, False)
    End If
    
    m_RetornoCheque = eRetornoCancelar
    Unload Me

End Sub
Private Sub CNPJ_GotFocus()
   'Se usuario selecionou, desabilita segunda digitação
    If Frame_Conf.Visible Then
        CNPJ_Conf.Text = ""
        CPF_CNPJ = 0
        Frame_Conf.Visible = False
    End If
    
   'flag da segunda digitação
    CNPJ.Tag = False
    
    CNPJ.SelStart = 0
    CNPJ.SelLength = CNPJ.MaxLength
End Sub
Private Sub CNPJ_KeyDown(KeyCode As Integer, Shift As Integer)
   'Se (enter) seta novo objeto
    If KeyCode = 13 Then
        'Valor.SetFocus
        SendKeys "{Tab}"
    End If
End Sub
Private Sub CNPJ_KeyPress(KeyAscii As Integer)
   'Apenas digitacao de numeros
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And (KeyAscii <> vbKeyBack And KeyAscii <> 13) Then
        KeyAscii = 0
    End If
End Sub
Private Sub CNPJ_LostFocus()
    Dim sCPF_CNPJ   As String

    If ActiveControl.Name <> "cmdSair" And _
        ActiveControl.Name <> "cmdConfirmar" And _
        ActiveControl.Name <> "CmdCMC7" And _
        ActiveControl.Name <> "cmdLinha1" And _
        ActiveControl.Name <> "CmdScanner" And _
        Val(CNPJ.Text) <> 0 Then
        
        If Len(Trim(CNPJ.Text)) = 11 Then
            
            'sCPF_CNPJ = Format(CNPJ.Text, String(11, "0"))
            
            If VerificaCPF(CNPJ.Text) = False Then
                CPF_CNPJ = CNPJ.Text
                CNPJ.Text = ""
                Frame_Conf.Caption = "Confirmação de CPF."
                Frame_Conf.Visible = True
                TipoInscricao = 1
                CNPJ_Conf.SetFocus
                
                Exit Sub
            Else
                TipoInscricao = 1
            End If
        ElseIf Len(Trim(CNPJ.Text)) = 14 Then
            
            sCPF_CNPJ = Format(CNPJ.Text, String(15, "0"))
            
            If VerificaCGC(sCPF_CNPJ) = False Then
                CPF_CNPJ = CNPJ.Text
                CNPJ.Text = ""
                Frame_Conf.Caption = "Confirmação de CNPJ."
                Frame_Conf.Visible = True
                TipoInscricao = 2
                CNPJ_Conf.SetFocus
                Exit Sub
            Else
                TipoInscricao = 2
            End If
        Else
            MsgBox "CNPJ/CPF inválido.", vbExclamation + vbOKOnly, App.Title
            CNPJ.SetFocus
            Exit Sub
        End If
    End If
    
End Sub
Private Sub CNPJ_Conf_GotFocus()
   'Seleciona Texto
    CNPJ_Conf.SelStart = 0
    CNPJ_Conf.SelLength = CNPJ_Conf.MaxLength
End Sub
Private Sub CNPJ_Conf_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys "{TAB}"
    End If
End Sub
Private Sub CNPJ_Conf_KeyPress(KeyAscii As Integer)
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And (KeyAscii <> vbKeyBack And KeyAscii <> 13) Then
        KeyAscii = 0
    End If
End Sub
Private Sub CNPJ_Conf_LostFocus()
    If ActiveControl.Name <> "cmdSair" And ActiveControl.Name <> "CNPJ" Then
        ''''''''''''''''''''''''''''''''
        ' * Valida Confirmação de CPF *'
        ''''''''''''''''''''''''''''''''
        If Len(Trim(CNPJ_Conf.Text)) = 0 Then
            MsgBox "Informe a confirmação do CPF/CNPJ.", vbExclamation + vbOKOnly, App.Title
            CNPJ_Conf.SetFocus
            Exit Sub
        End If
        ''''''''''''''''''''''
        '   * Valida CPF *   '
        ''''''''''''''''''''''
        If CDbl(CNPJ_Conf) <> CPF_CNPJ Then
            MsgBox "CPF/CNPJ não confere com a primeira digitação.", vbExclamation + vbOKOnly, App.Title
            Frame_Conf.Visible = False
            CNPJ.Text = CPF_CNPJ
            CNPJ.Tag = True
            CNPJ_Conf.Text = ""
            CNPJ.SetFocus
        Else
            Frame_Conf.Visible = False
            CNPJ.Text = CPF_CNPJ
            CNPJ.Tag = True
            CNPJ_Conf.Text = ""
        End If
        
    End If
End Sub
Sub Msg(pStatus As Boolean)

    cmdSair.Cancel = Not CBool(pStatus)
    
   'Desabilita botões durante leitura
    CmdCMC7.Enabled = Not CBool(pStatus)
    cmdConfirmar.Enabled = Not CBool(pStatus)
    cmdSair.Enabled = Not CBool(pStatus)
    cmdLinha1.Enabled = Not CBool(pStatus)
    
    If pStatus Then
        Picture1.BackColor = &HC0C0FF    '&HFFFFC0
        LblMsg.Caption = "Insira Documento para Captura"
        LblEsc.Visible = True
        ImgCheque.Visible = False
        ImgScanner.Visible = True
        Screen.MousePointer = vbArrowHourglass
    Else
        Picture1.BackColor = &HC0C0C0
        LblMsg.Caption = Info
        LblEsc.Visible = False
        ImgCheque.Visible = True
        ImgScanner.Visible = False
        Screen.MousePointer = vbDefault
    End If
End Sub
Sub TopPosition(Posicao As Boolean)              'Sc Scanner 1 ou 2 (Diferenciação de leitura)
    
    CNPJ.TabIndex = IIf(Posicao, 1, 19)
    CNPJ_Conf.TabIndex = IIf(Posicao, 2, 20)
    Valor.TabIndex = IIf(Posicao, 3, 21)
    cmdConfirmar.TabIndex = 22
    cmdSair.TabIndex = 23
    cmdLinha1.TabIndex = 24
    CmdCMC7.TabIndex = 25
            
    CMC7_Campo1.Top = IIf(Posicao, BOT_CMC7, TOP_CMC7)
    CMC7_Campo2.Top = IIf(Posicao, BOT_CMC7, TOP_CMC7)
    CMC7_Campo3.Top = IIf(Posicao, BOT_CMC7, TOP_CMC7)
    
    Linha1_Comp.Top = IIf(Posicao, BOT_CMC7, TOP_CMC7)
    Linha1_Bco.Top = IIf(Posicao, BOT_CMC7, TOP_CMC7)
    Linha1_C1.Top = IIf(Posicao, BOT_CMC7, TOP_CMC7)
    Linha1_Ag.Top = IIf(Posicao, BOT_CMC7, TOP_CMC7)
    Linha1_Conta.Top = IIf(Posicao, BOT_CMC7, TOP_CMC7)
    Linha1_C2.Top = IIf(Posicao, BOT_CMC7, TOP_CMC7)
    Linha1_Cheque.Top = IIf(Posicao, BOT_CMC7, TOP_CMC7)
    Linha1_C3.Top = IIf(Posicao, BOT_CMC7, TOP_CMC7)
    Linha1_Tipo.Top = IIf(Posicao, BOT_CMC7, TOP_CMC7)
    
    Frame_CNPJ.Top = IIf(Posicao, TOP_CNPJ, BOT_CNPJ)
    Frame_Conf.Top = IIf(Posicao, TOP_CNPJ, BOT_CNPJ)
    Frame_Valor.Top = IIf(Posicao, TOP_CNPJ, BOT_CNPJ)
End Sub
Private Sub Valor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Not Principal.Scanner.HABILITADO Or Principal.Scanner.Scanner = eL100 Then
            If bInclusao Then
                If Linha1 Then
                    Linha1_Comp.SetFocus
                Else
                    CMC7_Campo1.SetFocus
                End If
            Else
                cmdConfirmar_Click
            End If
            
        ElseIf Principal.Scanner.Scanner = eLA93 Then
            cmdConfirmar_Click
        End If
    End If
End Sub
Private Sub valor_KeyPress(KeyAscii As Integer)
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And (KeyAscii <> vbKeyBack And KeyAscii <> 13) Then
        KeyAscii = 0
    End If
End Sub
