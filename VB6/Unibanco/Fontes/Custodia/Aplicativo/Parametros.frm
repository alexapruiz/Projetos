VERSION 5.00
Object = "{ED123F48-E23F-11D4-B08D-00600899AB13}#1.0#0"; "UbbEdit.ocx"
Begin VB.Form Parametros 
   Caption         =   "Sistema de Captura - Parâmetros"
   ClientHeight    =   6225
   ClientLeft      =   780
   ClientTop       =   1515
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6225
   ScaleWidth      =   9495
   Begin VB.CheckBox chkSoma 
      Caption         =   "Criticar  Somatória de Controles"
      Height          =   276
      Left            =   2028
      TabIndex        =   51
      Top             =   3696
      Width           =   2892
   End
   Begin VB.TextBox txtNomeTerceira 
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
      Left            =   6264
      MaxLength       =   40
      TabIndex        =   47
      Top             =   2160
      Width           =   2724
   End
   Begin VB.TextBox txtCidadeTerceira 
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
      Left            =   6396
      MaxLength       =   25
      TabIndex        =   27
      Top             =   2976
      Width           =   2604
   End
   Begin VB.TextBox TxtDirRecep 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   2244
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   966
      Width           =   6756
   End
   Begin VB.TextBox TxtDirTrans 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   2244
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   537
      Width           =   6756
   End
   Begin VB.Frame Frame1 
      Caption         =   "Scanner"
      Height          =   660
      Left            =   2004
      TabIndex        =   46
      Top             =   4080
      Width           =   5460
      Begin VB.OptionButton optcom 
         Caption         =   "COM2"
         Height          =   228
         Index           =   1
         Left            =   4416
         TabIndex        =   35
         Top             =   288
         Width           =   732
      End
      Begin VB.OptionButton optcom 
         Caption         =   "COM1"
         Height          =   228
         Index           =   0
         Left            =   3648
         TabIndex        =   34
         Top             =   288
         Width           =   732
      End
      Begin VB.ComboBox cboscanner 
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
         ItemData        =   "Parametros.frx":0000
         Left            =   144
         List            =   "Parametros.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   240
         Width           =   3204
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Opções do Arquivo CEL"
      Height          =   804
      Left            =   48
      TabIndex        =   45
      Top             =   4800
      Width           =   9396
      Begin VB.CheckBox ChkGerarCEL 
         Caption         =   "Gerar Arquivo"
         Height          =   228
         Left            =   432
         TabIndex        =   36
         Top             =   360
         Width           =   1620
      End
      Begin UbbEdt.UbbEdit TxtCompOrigem 
         Height          =   360
         Left            =   8280
         TabIndex        =   42
         Top             =   270
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   635
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
         TextMaxNumChars =   5
         Layout          =   0
         BorderStyle     =   0
      End
      Begin UbbEdt.UbbEdit TxtVersaoFinalCEL 
         Height          =   360
         Left            =   5955
         TabIndex        =   40
         Top             =   270
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   635
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
         TextMaxNumChars =   5
         Layout          =   0
         BorderStyle     =   0
      End
      Begin UbbEdt.UbbEdit TxtVersaoInicialCEL 
         Height          =   360
         Left            =   3750
         TabIndex        =   38
         Top             =   270
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   635
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
         TextMaxNumChars =   5
         Layout          =   0
         BorderStyle     =   0
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Versão Final"
         Height          =   192
         Left            =   4992
         TabIndex        =   39
         Top             =   336
         Width           =   912
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Versão Inicial"
         Height          =   192
         Left            =   2640
         TabIndex        =   37
         Top             =   336
         Width           =   972
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Comp. Origem"
         Height          =   192
         Left            =   7200
         TabIndex        =   41
         Top             =   336
         Width           =   1032
      End
   End
   Begin VB.CheckBox ChkHeaderAV 
      Caption         =   "Header AV"
      Height          =   276
      Left            =   6420
      TabIndex        =   32
      Top             =   3696
      Width           =   1068
   End
   Begin VB.CommandButton CmdBrowseRecep 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   9072
      TabIndex        =   11
      Top             =   984
      Width           =   372
   End
   Begin VB.CommandButton CmdBrowseTrans 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   9072
      TabIndex        =   8
      Top             =   552
      Width           =   372
   End
   Begin VB.CommandButton CmdConfirmar 
      Caption         =   "&Confirmar"
      Height          =   396
      Left            =   3132
      TabIndex        =   43
      Top             =   5712
      Width           =   1572
   End
   Begin VB.CommandButton CmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sai&r"
      Height          =   396
      Left            =   4788
      TabIndex        =   44
      Top             =   5712
      Width           =   1572
   End
   Begin UbbEdt.UbbEdit TxtQtdCheques 
      Height          =   360
      Left            =   2250
      TabIndex        =   1
      Top             =   120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   635
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
      TextMaxNumChars =   3
      Layout          =   0
      BorderStyle     =   0
   End
   Begin UbbEdt.UbbEdit TxtCPDOrigem 
      Height          =   360
      Left            =   2250
      TabIndex        =   17
      Top             =   1815
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   635
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
      FieldType       =   1
      TextMaxNumChars =   3
      Layout          =   0
      TextAlignment   =   0
      BorderStyle     =   0
   End
   Begin UbbEdt.UbbEdit TxtCodigoUSB 
      Height          =   360
      Left            =   2250
      TabIndex        =   13
      Top             =   1395
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   635
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
      TextMaxNumChars =   5
      Layout          =   0
      BorderStyle     =   0
   End
   Begin UbbEdt.UbbEdit TxtQtdDatas 
      Height          =   360
      Left            =   5160
      TabIndex        =   3
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   635
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
      TextMaxNumChars =   2
      Layout          =   0
      BorderStyle     =   0
   End
   Begin UbbEdt.UbbEdit TxtCodigoTerceira 
      Height          =   360
      Left            =   2250
      TabIndex        =   21
      Top             =   2235
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   635
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
      FieldType       =   1
      TextMaxNumChars =   4
      Layout          =   0
      TextAlignment   =   0
      BorderStyle     =   0
   End
   Begin UbbEdt.UbbEdit TxtCPDDestino 
      Height          =   360
      Left            =   8430
      TabIndex        =   19
      Top             =   1800
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   635
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
      FieldType       =   1
      TextMaxNumChars =   3
      Layout          =   0
      TextAlignment   =   0
      BorderStyle     =   0
   End
   Begin UbbEdt.UbbEdit TxtCNPJTerceira 
      Height          =   360
      Left            =   6975
      TabIndex        =   23
      Top             =   2565
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   635
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
      FieldType       =   1
      TextMaxNumChars =   14
      Layout          =   0
      TextAlignment   =   0
      BorderStyle     =   0
   End
   Begin UbbEdt.UbbEdit TxtValorChequeLimite 
      Height          =   360
      Left            =   2250
      TabIndex        =   25
      Top             =   2670
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   635
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
      FieldType       =   2
      TextMaxNumChars =   11
      Layout          =   0
      BorderStyle     =   0
      Title           =   "Valor (R$)"
   End
   Begin UbbEdt.UbbEdit TxtAgAcolhed 
      Height          =   585
      Left            =   8430
      TabIndex        =   15
      Top             =   1170
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   1032
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
      TextMaxNumChars =   4
      BorderStyle     =   0
      Title           =   ""
   End
   Begin UbbEdt.UbbEdit txtQtdDias 
      Height          =   360
      Left            =   8655
      TabIndex        =   5
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   635
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
      TextMaxNumChars =   2
      Layout          =   0
      BorderStyle     =   0
   End
   Begin UbbEdt.UbbEdit TxtCodAplicacao 
      Height          =   360
      Left            =   2250
      TabIndex        =   29
      Top             =   3075
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   635
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
      FieldType       =   1
      TextMaxNumChars =   3
      Layout          =   0
      TextAlignment   =   0
      BorderStyle     =   0
   End
   Begin UbbEdt.UbbEdit TxtUF_Terceira 
      Height          =   360
      Left            =   8565
      TabIndex        =   31
      Top             =   3420
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   635
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
      FieldType       =   1
      TextMaxNumChars =   2
      Layout          =   0
      TextAlignment   =   0
      BorderStyle     =   0
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Nome Terceira"
      Height          =   192
      Left            =   5040
      TabIndex        =   50
      Top             =   2244
      Width           =   1092
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "CNPJ Terceira"
      Height          =   192
      Left            =   7092
      TabIndex        =   49
      Top             =   2160
      Width           =   1068
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Nome Terceira"
      Height          =   192
      Left            =   7080
      TabIndex        =   48
      Top             =   2280
      Width           =   1092
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Cidade Terceira"
      Height          =   192
      Left            =   5052
      TabIndex        =   26
      Top             =   3096
      Width           =   1176
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "UF Terceira"
      Height          =   192
      Left            =   7644
      TabIndex        =   30
      Top             =   3480
      Width           =   864
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Código Aplicação"
      Height          =   192
      Left            =   828
      TabIndex        =   28
      Top             =   3144
      Width           =   1296
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Quant. Miníma Dias Data de Depósito "
      Height          =   192
      Left            =   5880
      TabIndex        =   4
      Top             =   216
      Width           =   2724
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Valor Cheque Limite"
      Height          =   192
      Left            =   720
      TabIndex        =   24
      Top             =   2736
      Width           =   1440
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "CNPJ Terceira"
      Height          =   192
      Left            =   5760
      TabIndex        =   22
      Top             =   2640
      Width           =   1068
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Quant. Cheques por Borderô"
      Height          =   192
      Left            =   120
      TabIndex        =   0
      Top             =   216
      Width           =   2040
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Quant. Datas por Borderô"
      Height          =   192
      Left            =   3276
      TabIndex        =   2
      Top             =   216
      Width           =   1824
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Agência Acolhedora"
      Height          =   192
      Left            =   6312
      TabIndex        =   14
      Top             =   1512
      Width           =   2040
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Código da Terceira"
      Height          =   192
      Left            =   120
      TabIndex        =   20
      Top             =   2304
      Width           =   2040
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "CPD Destino"
      Height          =   192
      Left            =   7440
      TabIndex        =   18
      Top             =   1920
      Width           =   924
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "CPD Origem"
      Height          =   192
      Left            =   120
      TabIndex        =   16
      Top             =   1872
      Width           =   2040
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Código USB"
      Height          =   192
      Left            =   1248
      TabIndex        =   12
      Top             =   1464
      Width           =   900
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Diretório de Recepção"
      Height          =   192
      Left            =   120
      TabIndex        =   9
      Top             =   1032
      Width           =   2040
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Diretório de Transmissão"
      Height          =   192
      Left            =   120
      TabIndex        =   6
      Top             =   624
      Width           =   2040
   End
End
Attribute VB_Name = "Parametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Proc    As New Custodia.Selecionar

Public Enum cam_BrowseForFolder
    camDefualtBrowse = 0
    camTheDesktop = 0
    camProgramsFolders = 2
    camControlPanel = 3
    camPrinters = 4
    camDocumentsFolder = 5
    camFavoritesFolder = 6
    camStartupFolder = 7
    camRecentFolder = 8
    camSendToFolder = 9
    camRecycleBin = 10
    camStartMenuFolder = 11
    camDesktopFolder = 16
    camMyComputer = 17
    camNetworkNeighborhood = 18
    camNetHoodFolder = 19
    camFontsFolder = 20
    camShellNewFolder = 21
End Enum

Dim nCOM               As Integer      'Variavel guardar porta do scanner p/ gravacao
Public Function BrowseForFolder(ByVal prmForm As Object, ByVal prmFolder As cam_BrowseForFolder) As String

    Dim bi                                  As BROWSEINFO
    Dim idl                                 As ITEMIDLIST
    Dim rtn                                 As Long
    Dim pidl                                As Long
    Dim path                                As String
    Dim Pos                                 As Integer
    Dim lresult                             As Long
    Dim x                                   As String
  
    bi.hOwner = prmForm.hwnd
    rtn& = SHGetSpecialFolderLocation(ByVal prmForm.hwnd, ByVal prmFolder, idl)
    
    bi.pidlRoot = idl.mkid.cb
    bi.lpszTitle = "Selecione a pasta"
    bi.ulFlags = BIF_RETURNONLYFSDIRS And BIF_DONTGOBELOWDOMAIN And BIF_STATUSTEXT _
        And BIF_RETURNFSANCESTORS And BIF_BROWSEFORCOMPUTER And BIF_BROWSEFORPRINTER
    
    pidl& = SHBrowseForFolder(bi) 'show the dialog box
    
    path$ = Space$(512) 'set the maximum returned path
    lresult = SHGetPathFromIDList(ByVal pidl&, ByVal path$)  'get the folder selected
    
    BrowseForFolder = ""
    If lresult Then 'if a folder was selected the
       Pos% = InStr(path$, Chr$(0)) 'extract the path
       BrowseForFolder = Left(path$, Pos - 1)
       'MsgBox "The folder you selected was:" + Chr$(10) + Chr$(10) + Left(path$, Pos - 1), vbInformation 'display the returned path
    End If

End Function
Sub PreencheCampos()
    On Error GoTo Erro:
    
    Dim sstr       As String
    Dim RsParametro As New ADODB.Recordset
    Dim RsSelecionar As New Custodia.Selecionar

    'Selecionando o registro da data
    Set RsParametro = g_cMainConnection.Execute(RsSelecionar.GetParametros(Geral.DataProcessamento))

    If Not RsParametro.EOF Then
        'Preencher os campos da tela
        TxtQtdCheques.Text = IIf(IsNull(RsParametro!QuantidadeCheques), 0, RsParametro!QuantidadeCheques)
        TxtQtdDatas.Text = IIf(IsNull(RsParametro!QuantidadeDatas), 0, RsParametro!QuantidadeDatas)
        txtQtdDias.Text = IIf(IsNull(RsParametro!QuantidadeMinimaDias), 0, RsParametro!QuantidadeMinimaDias)
        TxtDirTrans.Text = IIf(IsNull(RsParametro!DiretorioTransmissao), "", Trim(RsParametro!DiretorioTransmissao))
        TxtDirRecep.Text = IIf(IsNull(RsParametro!DiretorioRecepcao), "", Trim(RsParametro!DiretorioRecepcao))
        TxtCodigoUSB.Text = IIf(IsNull(RsParametro!Codigo_USB), 0, RsParametro!Codigo_USB)
        TxtAgAcolhed.Text = IIf(IsNull(RsParametro!CodigoAgAcolhed), 0, RsParametro!CodigoAgAcolhed)
        TxtCPDOrigem.Text = IIf(IsNull(RsParametro!CPD_Origem), "", RsParametro!CPD_Origem)
        TxtCPDDestino.Text = IIf(IsNull(RsParametro!CPD_Destino), "", RsParametro!CPD_Destino)
        TxtCodigoTerceira.Text = IIf(IsNull(RsParametro!Codigo_Terceira), "", RsParametro!Codigo_Terceira)
        txtNomeTerceira.Text = IIf(IsNull(RsParametro!Nome_Terceira), "", RsParametro!Nome_Terceira)
        TxtCNPJTerceira.Text = IIf(IsNull(RsParametro!CNPJ_Terceira), "", RsParametro!CNPJ_Terceira)
        txtCidadeTerceira.Text = RsParametro!Cidade_Terceira & ""
        TxtUF_Terceira.Text = IIf(IsNull(RsParametro!UF_Terceira), "", RsParametro!UF_Terceira)
        TxtValorChequeLimite.Text = IIf(IsNull(RsParametro!ValorChequeLimite), 0, Format(RsParametro!ValorChequeLimite, "#,##0.00"))
        TxtVersaoInicialCEL.Text = IIf(IsNull(RsParametro!Num_Versao_Inicial_CEL), 0, RsParametro!Num_Versao_Inicial_CEL)
        TxtVersaoFinalCEL.Text = IIf(IsNull(RsParametro!Num_Versao_Final_CEL), 0, RsParametro!Num_Versao_Final_CEL)
        TxtCompOrigem = IIf(IsNull(RsParametro!Comp_Origem_CEL), 0, RsParametro!Comp_Origem_CEL)
        TxtCodAplicacao.Text = IIf(IsNull(RsParametro!CodigoAplicacao), "", RsParametro!CodigoAplicacao)
        

        '* Scanner / Porta de Comunicação *'
        sstr = Trim(PegarOpcaoINI("Scanner", "Tipo", ""))
        
        If AchaScannerDLL(sstr) Then
            If sstr = "1" Or sstr = "2" Then
                cboscanner.ListIndex = CInt(sstr)
                sstr = Trim(PegarOpcaoINI("Scanner", "PortaCOM", ""))
                If sstr = "1" Or sstr = "2" Then
                    optcom.Item(IIf(sstr = 1, 0, 1)).Value = True
                    nCOM = CInt(sstr)
                Else
                    MsgBox "ATENÇÃO !!! - Parâmetro (PortaCom do arquivo .ini) é Inválida, Verifique...", vbInformation + vbOKOnly, App.Title & " - Custodia.ini"
                    cboscanner.ListIndex = 0
                    optcom.Item(0).Value = False
                    optcom.Item(1).Value = False
                End If
            Else
                If sstr <> "0" Then
                    MsgBox "ATENÇÃO !!! - Parâmetro (Scanner do arquivo .ini) é Inválido, Verifique...", vbInformation + vbOKOnly, App.Title & " - Custodia.ini"
                End If
                    cboscanner.ListIndex = 0
                    optcom.Item(0).Value = False
                    optcom.Item(1).Value = False
            End If
        Else
            cboscanner.ListIndex = 0
            optcom.Item(0).Value = False
            optcom.Item(1).Value = False
        End If

        '* HeaderAV *'
        If RsParametro!HeaderAV = True Then
            ChkHeaderAV.Value = 1
        Else
            ChkHeaderAV.Value = 0
        End If
        
        '* Critica de Somatória *'
        If RsParametro!CriticaSoma = True Then
            chkSoma.Value = 1
        Else
            chkSoma.Value = 0
        End If

        '* Gerar Arquivo CEL *'
        If RsParametro!GerarArquivo_CEL = True Then
            ChkGerarCEL.Value = 1
        Else
            ChkGerarCEL.Value = 0
        End If

    Else
        MsgBox "Não foi possível ler os parâmetros desta Data de Processamento.", vbInformation + vbOKOnly, App.Title
        Exit Sub
    End If

    Set RsParametro = Nothing
    Set RsSelecionar = Nothing
    
Exit Sub

Erro:
    Call TratamentoErro("Falha na Recuperação de parâmetros do sistema.", Err)
End Sub
Private Function VerificaCampos() As Boolean

    VerificaCampos = False

    'Verificar se todos os campos estão preenchidos

    'Quantidade de Cheques
    If Val(TxtQtdCheques.Text) = 0 Then
        MsgBox "Informe a Quantidade de Cheques.", vbInformation + vbOKOnly, App.Title
        TxtQtdCheques.SetFocus
        Exit Function
    End If

    'Quantidade de Datas
    If Val(TxtQtdDatas.Text) = 0 Then
        MsgBox "Informe a Quantidade de Datas.", vbInformation + vbOKOnly, App.Title
        TxtQtdDatas.SetFocus
        Exit Function
    End If

    'Diretório de Transmissão
    If Len(Trim(TxtDirTrans.Text)) = 0 Or Dir(TxtDirTrans.Text, vbDirectory) = "" Then
        MsgBox "Informe um Caminho válido para Transmissão de Arquivos .", vbInformation + vbOKOnly, App.Title
        TxtDirTrans.SetFocus
        Exit Function
    End If

    'Diretório de Recepção
    If Len(Trim(TxtDirRecep.Text)) = 0 Or Dir(TxtDirRecep.Text, vbDirectory) = "" Then
        MsgBox "Informe um Caminho válido para Recepção de Arquivos.", vbInformation + vbOKOnly, App.Title
        TxtDirRecep.SetFocus
        Exit Function
    End If

    'Codigo USB
    If IsEmpty(TxtCodigoUSB.Text) Then
        MsgBox "Informe o Código da USB.", vbInformation + vbOKOnly, App.Title
        TxtCodigoUSB.SetFocus
        Exit Function
    End If

    'Agencia Acolhedora
    If IsEmpty(TxtAgAcolhed.Text) Then
        MsgBox "Informe o Código da Agência Acolhedora.", vbInformation + vbOKOnly, App.Title
        TxtAgAcolhed.SetFocus
        Exit Function
    End If

    'CPD Origem
    If IsEmpty(TxtCPDOrigem.Text) Then
        MsgBox "Informe o Código do CPD de Origem.", vbInformation + vbOKOnly, App.Title
        TxtCPDOrigem.SetFocus
        Exit Function
    End If

    'CPD Destino
    If IsEmpty(TxtCPDDestino.Text) Then
        MsgBox "Informe o Código do CPD de Destino.", vbInformation + vbOKOnly, App.Title
        TxtCPDDestino.SetFocus
        Exit Function
    End If

    'Codigo da Terceira
    If IsEmpty(TxtCodigoTerceira.Text) Then
        MsgBox "Informe o Código da Terceira.", vbInformation + vbOKOnly, App.Title
        TxtCodigoTerceira.SetFocus
        Exit Function
    End If

    'UF da Terceira
    If Len(Trim(TxtUF_Terceira)) = 0 Then
        MsgBox "Informe a UF da Terceira.", vbInformation + vbOKOnly, App.Title
        TxtUF_Terceira.SetFocus
        Exit Function
    End If

    'Código da Aplicacao
    If Len(Trim(TxtCodAplicacao.Text)) = 0 Then
        MsgBox "Informe o Código da Aplicação.", vbInformation + vbOKOnly, App.Title
        TxtCodAplicacao.SetFocus
        Exit Function
    End If

    'CNPJ Terceira
    If Val(TxtCNPJTerceira.Text) = 0 Then
        MsgBox "Informe o CNPJ da Terceira.", vbInformation + vbOKOnly, App.Title
        TxtCNPJTerceira.SetFocus
        Exit Function
    End If
    
    'Cidade Terceira
    If Trim(txtCidadeTerceira.Text) = "" Then
        MsgBox "Informe a Cidade da Terceira.", vbInformation + vbOKOnly, App.Title
        txtCidadeTerceira.SetFocus
        Exit Function
    End If

    'Valor Limite Cheque
    If Val(TxtValorChequeLimite.Text) = 0 Then
        MsgBox "Informe o Valor para Cheque Limite.", vbInformation + vbOKOnly, App.Title
        TxtValorChequeLimite.SetFocus
        Exit Function
    End If

    'Versao Inicial - CEL
    If Val(TxtVersaoInicialCEL.Text) = 0 Then
        MsgBox "Informe o Valor Inicial para o Arquivo CEL.", vbInformation + vbOKOnly, App.Title
        TxtVersaoInicialCEL.SetFocus
        Exit Function
    End If
    
    'Versao Final - CEL
    If Val(TxtVersaoFinalCEL.Text) = 0 Then
        MsgBox "Informe o Valor Final para o Arquivo CEL.", vbInformation + vbOKOnly, App.Title
        TxtVersaoFinalCEL.SetFocus
        Exit Function
    End If
    
    'Comp. Origem - CEL
    If Val(TxtCompOrigem.Text) = 0 Then
        MsgBox "Informe o código da Câmara de Compensação da Origem.", vbInformation + vbOKOnly, App.Title
        TxtCompOrigem.SetFocus
        Exit Function
    End If

    VerificaCampos = True

End Function
Private Sub AtualizaParametros()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'              * Atualiza Informações de Parametros (T Y P E) - ON LINE *                   '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

    Dim RstAux As New ADODB.Recordset
    Dim sstr    As String

    Set RstAux = g_cMainConnection.Execute(Proc.GetParametros(Geral.DataProcessamento))
    
    If Not RstAux.EOF() Then
        
        With g_Parametros
            .QuantidadeCheques = RstAux!QuantidadeCheques & ""
            .QuantidadeDatas = RstAux!QuantidadeDatas & ""
            .DiretorioTransmissao = RstAux!DiretorioTransmissao & ""
            .DiretorioRecepcao = RstAux!DiretorioRecepcao & ""
            .Sequencia_Bordero = Val(RstAux!Seq_Bordero & "")
            .Gerar_Arquivo_CEL = CBool(RstAux!GerarArquivo_CEL)
            .Comp_Origem_CEL = Val(RstAux!Comp_Origem_CEL & "")
            .Numero_Lote_CEL = Val(RstAux!Num_Lote_CEL & "")
            .Numero_Versao_Inicial_CEL = Val(RstAux!Num_Versao_Inicial_CEL & "")
            .Numero_Versao_Final_CEL = Val(RstAux!Num_Versao_Final_CEL & "")
            .HeaderAV = CBool(RstAux!HeaderAV)
            .chkSoma = CBool(RstAux!CriticaSoma)
            .Codigo_USB = Val(RstAux!Codigo_USB & "")
            .CPD_Origem = RstAux!CPD_Origem & ""
            .CPD_Destino = RstAux!CPD_Destino & ""
            .Codigo_Terceira = RstAux!Codigo_Terceira & ""
            .CNPJ_Terceira = RstAux!CNPJ_Terceira & ""
            .UF_Terceira = RstAux!UF_Terceira & ""
            .Seq_Ocorrencia = IIf(IsNull(RstAux!Seq_Ocorrencia), 0, RstAux!Seq_Ocorrencia) & ""
            .CodigoAgAcolhed = IIf(IsNull(RstAux!CodigoAgAcolhed), 0, RstAux!CodigoAgAcolhed) & ""
            .Num_Remessa_TER = RstAux!Num_Remessa_TER & ""
            .ValorChequeLimite = IIf(IsNull(RstAux!ValorChequeLimite), 0, RstAux!ValorChequeLimite)
            .CodigoAplicacao = IIf(IsNull(RstAux!CodigoAplicacao), "", RstAux!CodigoAplicacao)
            .TMP_Pendente = IIf(IsNull(RstAux!TMP_Pendente), 0, RstAux!TMP_Pendente) & ""
            .QuantidadeMinimaDias = Val(txtQtdDias.Text)
            .Cidade_Terceira = txtCidadeTerceira.Text
            

            '''''''''''''''''''''''''''''''''''''''''''
            ' * Verifica .ini se estação usará L100 * '
            '''''''''''''''''''''''''''''''''''''''''''
            .Scanner = Trim(PegarOpcaoINI("Scanner", "Tipo", ""))
            .PortaCom = Trim(PegarOpcaoINI("Scanner", "PortaCOM", ""))
      
        End With
    End If
    RstAux.Close

Exit Sub
TrataErro:
    Call TratamentoErro("Erro ao atualizar Type de Parâmetros.", Err)
    Unload Me
    
End Sub
Private Sub cboscanner_Click()
    If cboscanner.ListIndex = 0 Then
        optcom(0).Value = False
        optcom(1).Value = False
        nCOM = 0
    Else
        optcom(0).Value = True
        nCOM = 1
    End If
End Sub

Private Sub CmdBrowseRecep_Click()

    Dim sDir        As String
    
    sDir = BrowseForFolder(Me, camMyComputer)
    
    If Trim(sDir) <> "" Then
        TxtDirRecep.Text = sDir
    End If
End Sub
Private Sub CmdBrowseTrans_Click()

    Dim sDir        As String
    
    sDir = BrowseForFolder(Me, camMyComputer)
    
    If Trim(sDir) <> "" Then
        TxtDirTrans.Text = sDir
    End If
End Sub
Private Sub cmdConfirmar_Click()
    Screen.MousePointer = vbHourglass
        If VerificaCampos Then
            Call GravaParametros
            Call AtualizaParametros
            Unload Me
        End If
    Screen.MousePointer = vbDefault
End Sub
Private Sub cmdSair_Click()
    Unload Me
End Sub
Private Sub Form_Activate()

    'Preencher os campos com os valores da base
    Call PreencheCampos
    
End Sub
Private Function GravaParametros() As Boolean
    
    On Error GoTo Erro:
    
    Dim RsAtualiza As New Custodia.Atualizar
    Dim RsParametros As New ADODB.Recordset
    
    GravaParametros = False

    'Atualizar a tabela de parametros
        Set RsParametros = g_cMainConnection.Execute(RsAtualiza.AtualizaParametros( _
                                                Geral.DataProcessamento, _
                                                Val(TxtQtdCheques.Text), _
                                                Val(TxtQtdDatas.Text), _
                                                Trim(TxtDirTrans.Text), _
                                                Trim(TxtDirRecep.Text), _
                                                Val(TxtCodigoUSB.Text), _
                                                Val(TxtAgAcolhed.Text), _
                                                Trim(TxtCPDOrigem.Text), _
                                                Trim(TxtCPDDestino.Text), _
                                                Trim(TxtCodigoTerceira.Text), _
                                                Trim(TxtCNPJTerceira.Text), _
                                                vbNull, _
                                                Trim(TxtUF_Terceira.Text), _
                                                Trim(TxtCodAplicacao.Text), _
                                                Format(Val(InserePonto(TxtValorChequeLimite.Text)), MASK_VALOR), _
                                                IIf(ChkHeaderAV.Value = 0, "No", "Yes"), _
                                                IIf(chkSoma.Value = 0, "No", "Yes"), _
                                                IIf(ChkGerarCEL.Value = 0, "No", "Yes"), _
                                                Val(TxtCompOrigem.Text), _
                                                Val(TxtVersaoInicialCEL.Text), _
                                                Val(TxtVersaoFinalCEL.Text), _
                                                Val(txtQtdDias.Text), _
                                                txtCidadeTerceira.Text, _
                                                txtNomeTerceira.Text))
                                                
    If cboscanner.ListIndex = 0 Then
       'não selecionado scanner
        GravarOpcaoINI "Scanner", "Tipo", 0
    Else
       'verifica se DLL referente ao scanner selecionada existe no dir. de sistema
        If AchaScannerDLL(cboscanner.ListIndex) Then
            If cboscanner.ListIndex = 1 Then
                GravarOpcaoINI "Scanner", "Tipo", 1
                GravarOpcaoINI "Scanner", "PortaCom", Val(nCOM)
            ElseIf cboscanner.ListIndex = 2 Then
               'Se scanner = LA93 verifica se driver instalado na estação
                If VerRegSerialLA93 = Trim(PegarOpcaoINI("Scanner", "Serial", "")) Then
                   'Se ha driver e DLL
                    GravarOpcaoINI "Scanner", "Tipo", 2
                    GravarOpcaoINI "Scanner", "PortaCom", Val(nCOM)
                Else
                   'Se driver não instalado
                    GravarOpcaoINI "Scanner", "Tipo", 0
                    Err.Raise 979, App.Title, "Driver( VipsDrv) do Scanner LA93 não Instalado." & vbCrLf _
                    & "Contate o Administrador do Sistema."
                End If
            End If
            
        Else
           'DLL não localizada
            GravarOpcaoINI "Scanner", "Tipo", 0
            Err.Raise 989, App.Title, "Não Localizado Arquivo (DLL) de Utilização do Scanner." & vbCrLf _
            & "Contate o Administrador do Sistema."
        End If
    End If
    
    GravaParametros = True
    Exit Function
    
Erro:
    Screen.MousePointer = vbDefault
    If Err.Number = 989 Or Err.Number = 979 Then
        Call TratamentoErro("Falha na Gravação dos parâmetros do Scanner." & vbCrLf & vbTab & "Verifique o problema em [Descrição], para continuar tecle [Sair]", Err, , True)
        Resume Next
    Else
        Call TratamentoErro("Falha na Gravação dos parâmetros do Sistema.", Err)
    End If

End Function
Private Sub Form_Load()
    Call PreencheCombo
End Sub
Private Sub PreencheCombo()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                         '* Preenche Combo com Valores Default *'                          '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    cboscanner.AddItem "Não possui Leitora"
    cboscanner.AddItem "L100"
    cboscanner.AddItem "LA93"
    

End Sub

Private Sub optcom_Click(Index As Integer)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                          '* Define Porta de Comunicação L100 *'                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If cboscanner.ListIndex > 0 Then
        If optcom(0).Value = True Then
            nCOM = 1
        End If
        If optcom(1).Value = True Then
            nCOM = 2
        End If
    Else
        optcom(0).Value = False
        optcom(1).Value = False
        nCOM = 0
    End If

End Sub

Private Sub txtCidadeTerceira_GotFocus()
     
     With txtCidadeTerceira
          .SelStart = 0
          .SelLength = .MaxLength
     End With

End Sub

Private Sub txtCidadeTerceira_KeyPress(KeyAscii As Integer)
     
     If KeyAscii = vbKeyEscape Then Exit Sub
     
     If KeyAscii = vbKeyReturn Then
          SendKeys "{TAB}"
          Exit Sub
     End If
     
     If KeyAscii = vbKeyBack Then Exit Sub
     With txtCidadeTerceira
          If Len(.Text) >= .MaxLength Then
               Beep
               KeyAscii = 0
               MsgBox "Número máximo permitido é de " & CStr(.MaxLength) & " caracteres", vbInformation, Me.Caption
               .SelStart = 0
               .SelLength = .MaxLength
               .SetFocus
               Exit Sub
          End If
     End With

End Sub



Private Sub txtNomeTerceira_GotFocus()
     With txtNomeTerceira
          .SelStart = 0
          .SelLength = .MaxLength
     End With
End Sub

Private Sub txtNomeTerceira_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyEscape Then Exit Sub
     
     If KeyAscii = vbKeyReturn Then
          SendKeys "{TAB}"
          Exit Sub
     End If
     
     If KeyAscii = vbKeyBack Then Exit Sub
     With txtNomeTerceira
          If Len(.Text) >= .MaxLength Then
               Beep
               KeyAscii = 0
               MsgBox "Número máximo permitido é de " & CStr(.MaxLength) & " caracteres", vbInformation, Me.Caption
               .SelStart = 0
               .SelLength = .MaxLength
               .SetFocus
               Exit Sub
          End If
     End With

End Sub
