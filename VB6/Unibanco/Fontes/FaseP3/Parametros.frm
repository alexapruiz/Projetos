VERSION 5.00
Object = "{9CBA5D64-E3C8-11D3-9FFC-00104BC8688C}#1.0#0"; "CURRENCYEDIT.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Parametros 
   Caption         =   "Parâmetros do Sistema"
   ClientHeight    =   5616
   ClientLeft      =   3204
   ClientTop       =   1704
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5616
   ScaleWidth      =   7800
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   372
      Left            =   4100
      TabIndex        =   90
      Top             =   5076
      Width           =   1452
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Confirmar"
      Height          =   372
      Left            =   2292
      TabIndex        =   89
      Top             =   5076
      Width           =   1452
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   4776
      Left            =   48
      TabIndex        =   0
      Top             =   48
      Width           =   7656
      _ExtentX        =   13504
      _ExtentY        =   8424
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   420
      TabCaption(0)   =   "Locais"
      TabPicture(0)   =   "Parametros.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame11"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Globais Envelope e Malote"
      TabPicture(1)   =   "Parametros.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(1)=   "Frame20"
      Tab(1).Control(2)=   "Frame21"
      Tab(1).Control(3)=   "Frame22"
      Tab(1).Control(4)=   "Frame6"
      Tab(1).Control(5)=   "Frame30"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Globais Envelope"
      TabPicture(2)   =   "Parametros.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame25"
      Tab(2).Control(1)=   "Frame13"
      Tab(2).Control(2)=   "Frame12"
      Tab(2).Control(3)=   "Frame10"
      Tab(2).Control(4)=   "Frame9"
      Tab(2).Control(5)=   "Frame8"
      Tab(2).Control(6)=   "Frame7"
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "Globais Malote"
      TabPicture(3)   =   "Parametros.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame14"
      Tab(3).Control(1)=   "Frame15"
      Tab(3).Control(2)=   "Frame16"
      Tab(3).Control(3)=   "Frame17"
      Tab(3).Control(4)=   "Frame18"
      Tab(3).Control(5)=   "Frame19"
      Tab(3).Control(6)=   "Frame23"
      Tab(3).Control(7)=   "Frame26"
      Tab(3).Control(8)=   "Frame27"
      Tab(3).Control(9)=   "fraAlcadaCoordMalote"
      Tab(3).ControlCount=   10
      TabCaption(4)   =   "Bloqueios/Autenticações"
      TabPicture(4)   =   "Parametros.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fraDiasTrocaSenha"
      Tab(4).Control(1)=   "Frame24"
      Tab(4).Control(2)=   "Frame29"
      Tab(4).Control(3)=   "Frame28"
      Tab(4).ControlCount=   4
      Begin VB.Frame fraDiasTrocaSenha 
         Caption         =   "Expirar senha"
         Height          =   732
         Left            =   -71016
         TabIndex        =   93
         Top             =   960
         Width           =   3432
         Begin VB.TextBox txtDiasTrocaSenha 
            Alignment       =   2  'Center
            ForeColor       =   &H00800000&
            Height          =   312
            Left            =   1800
            MaxLength       =   2
            TabIndex        =   94
            Tag             =   "QtdeDiasTrocaSenha"
            Top             =   252
            Width           =   816
         End
         Begin VB.Label lblDiasTrocadeSenha 
            AutoSize        =   -1  'True
            Caption         =   "Dias úteis"
            Height          =   192
            Left            =   840
            TabIndex        =   95
            Top             =   300
            Width           =   708
         End
      End
      Begin VB.Frame fraAlcadaCoordMalote 
         Caption         =   "Limite para Aprovação do Coordenador"
         Height          =   660
         Left            =   -70992
         TabIndex        =   91
         Top             =   1272
         Width           =   3252
         Begin CURRENCYEDITLib.CurrencyEdit txtAlcadaCoordMalote 
            Height          =   312
            Left            =   1380
            TabIndex        =   75
            Tag             =   "ValorAlcadaCoord_Mal"
            Top             =   240
            Width           =   1632
            _Version        =   65537
            _ExtentX        =   2879
            _ExtentY        =   550
            _StockProps     =   93
            ForeColor       =   8388608
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   11
            BackColor       =   -2147483643
         End
         Begin VB.Label lblAlcadaCoordMalote 
            Caption         =   "Alçada:"
            Height          =   252
            Left            =   180
            TabIndex        =   65
            Top             =   300
            Width           =   672
         End
      End
      Begin VB.Frame Frame30 
         Height          =   732
         Left            =   -74448
         TabIndex        =   20
         Top             =   2892
         Width           =   6552
         Begin VB.CheckBox CheckRecepcionaIK 
            Caption         =   "Efetuar Recepção de Capas no IK"
            Height          =   192
            Left            =   180
            TabIndex        =   92
            Top             =   312
            Width           =   2940
         End
      End
      Begin VB.Frame Frame24 
         Caption         =   "Autenticações por Grupo de Documento"
         Height          =   1332
         Left            =   -74736
         TabIndex        =   83
         Top             =   2640
         Width           =   7200
         Begin VB.ComboBox cmbGrupoDocumentos 
            Height          =   288
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   87
            Tag             =   "GrupoDocumentos"
            Top             =   600
            Width           =   3492
         End
         Begin VB.TextBox txtNr_Autenticacoes 
            ForeColor       =   &H00800000&
            Height          =   312
            Left            =   4380
            MaxLength       =   1
            TabIndex        =   88
            Tag             =   "Nr_Autenticacoes"
            Top             =   600
            Width           =   1512
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Nr. de Autenticações"
            Height          =   192
            Left            =   4380
            TabIndex        =   85
            Top             =   300
            Width           =   1488
         End
      End
      Begin VB.Frame Frame29 
         Caption         =   "Inativação automática de usuários"
         Height          =   732
         Left            =   -74688
         TabIndex        =   82
         Top             =   936
         Width           =   3432
         Begin VB.TextBox txtdiasinativo 
            Alignment       =   2  'Center
            ForeColor       =   &H00800000&
            Height          =   312
            Left            =   1656
            MaxLength       =   3
            TabIndex        =   86
            Tag             =   "DiasInativo"
            Top             =   252
            Width           =   816
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Dias"
            Height          =   192
            Left            =   888
            TabIndex        =   84
            Top             =   300
            Width           =   336
         End
      End
      Begin VB.Frame Frame28 
         Caption         =   "Tempo para bloqueio de estação"
         Height          =   732
         Left            =   -70956
         TabIndex        =   1
         Top             =   1944
         Visible         =   0   'False
         Width           =   3408
         Begin VB.TextBox txttmpbloqueio 
            ForeColor       =   &H00800000&
            Height          =   312
            Left            =   1416
            MaxLength       =   3
            TabIndex        =   3
            Tag             =   "TmpBloqueio"
            Top             =   276
            Width           =   1632
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "segundos"
            Height          =   192
            Left            =   264
            TabIndex        =   2
            Top             =   348
            Width           =   720
         End
      End
      Begin VB.Frame Frame6 
         Height          =   636
         Left            =   -74460
         TabIndex        =   18
         Top             =   1248
         Width           =   6576
         Begin CURRENCYEDITLib.CurrencyEdit txtInferior 
            Height          =   312
            Left            =   3900
            TabIndex        =   29
            Tag             =   "ValorInferior"
            Top             =   216
            Width           =   1512
            _Version        =   65537
            _ExtentX        =   2667
            _ExtentY        =   550
            _StockProps     =   93
            ForeColor       =   8388608
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   11
            BackColor       =   -2147483643
         End
         Begin VB.Label Label2 
            Caption         =   "Valor de Compesação Inferior (menores ou iguais):"
            Height          =   252
            Left            =   120
            TabIndex        =   24
            Top             =   276
            Width           =   3732
         End
      End
      Begin VB.Frame Frame22 
         Height          =   732
         Left            =   -74448
         TabIndex        =   19
         Top             =   2028
         Width           =   6552
         Begin CURRENCYEDITLib.CurrencyEdit txtAjusteContabil 
            Height          =   312
            Left            =   3900
            TabIndex        =   30
            Tag             =   "ValorAjusteContabil"
            Top             =   276
            Width           =   1512
            _Version        =   65537
            _ExtentX        =   2667
            _ExtentY        =   550
            _StockProps     =   93
            ForeColor       =   8388608
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   11
            BackColor       =   12632256
            Locked          =   -1  'True
         End
         Begin VB.Label Label16 
            Caption         =   "Valor de Ajuste Contábil (menores ou iguais):"
            Height          =   252
            Left            =   120
            TabIndex        =   25
            Top             =   300
            Width           =   3732
         End
      End
      Begin VB.Frame Frame27 
         Caption         =   "Conversão de ch. UBB para Compensação"
         Height          =   684
         Left            =   -74604
         TabIndex        =   61
         Top             =   3864
         Width           =   3300
         Begin CURRENCYEDITLib.CurrencyEdit txtValorCompensaNovo_Mal 
            Height          =   312
            Left            =   1320
            TabIndex        =   80
            Tag             =   "ValorCompensaNovo_Mal"
            Top             =   240
            Width           =   1632
            _Version        =   65537
            _ExtentX        =   2879
            _ExtentY        =   550
            _StockProps     =   93
            ForeColor       =   8388608
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   11
            BackColor       =   -2147483643
         End
         Begin VB.Label Label21 
            Caption         =   "Malote Novo Caixa Robô:"
            Height          =   468
            Left            =   180
            TabIndex        =   71
            Top             =   192
            Width           =   1080
         End
      End
      Begin VB.Frame Frame26 
         Caption         =   "Aprovação de ch. terceiro para Supervisor"
         Height          =   732
         Left            =   -70992
         TabIndex        =   60
         Top             =   3864
         Width           =   3300
         Begin CURRENCYEDITLib.CurrencyEdit txtAlcadaUBBMal 
            Height          =   312
            Left            =   1356
            TabIndex        =   81
            Tag             =   "ValorAlcadaOutros_Mal"
            Top             =   240
            Width           =   1632
            _Version        =   65537
            _ExtentX        =   2879
            _ExtentY        =   550
            _StockProps     =   93
            ForeColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   11
         End
         Begin VB.Label Label20 
            Caption         =   "Alçada :"
            Height          =   252
            Left            =   180
            TabIndex        =   70
            Top             =   300
            Width           =   972
         End
      End
      Begin VB.Frame Frame25 
         Caption         =   "Aprovação de cheque terceiro p/ Supervisor"
         Height          =   732
         Left            =   -74616
         TabIndex        =   38
         Top             =   3588
         Width           =   3396
         Begin CURRENCYEDITLib.CurrencyEdit txtAlcadaUBBEnv 
            Height          =   312
            Left            =   1320
            TabIndex        =   52
            Tag             =   "ValorAlcadaOutros_Env"
            Top             =   240
            Width           =   1632
            _Version        =   65537
            _ExtentX        =   2879
            _ExtentY        =   550
            _StockProps     =   93
            ForeColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   11
         End
         Begin VB.Label Label19 
            Caption         =   "Alçada :"
            Height          =   252
            Left            =   180
            TabIndex        =   45
            Top             =   300
            Width           =   972
         End
      End
      Begin VB.Frame Frame23 
         Caption         =   "Limite Máximo Dif. Malote Novo"
         Height          =   732
         Left            =   -74604
         TabIndex        =   59
         Top             =   3000
         Width           =   3300
         Begin CURRENCYEDITLib.CurrencyEdit TxtValorMaxDifLancto 
            Height          =   312
            Left            =   1320
            TabIndex        =   78
            Tag             =   "LimiteMaxDifLancto_Mal"
            Top             =   264
            Width           =   1632
            _Version        =   65537
            _ExtentX        =   2879
            _ExtentY        =   550
            _StockProps     =   93
            ForeColor       =   8388608
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   11
            BackColor       =   -2147483643
         End
         Begin VB.Label Label17 
            Caption         =   "Valor :"
            Height          =   252
            Left            =   204
            TabIndex        =   69
            Top             =   312
            Width           =   504
         End
      End
      Begin VB.Frame Frame21 
         Caption         =   "Autorização de Débito em C/C"
         Height          =   732
         Left            =   -70956
         TabIndex        =   22
         Top             =   3840
         Width           =   3060
         Begin CURRENCYEDITLib.CurrencyEdit TxtMaxADCC 
            Height          =   312
            Left            =   1308
            TabIndex        =   31
            Tag             =   "ValorMaxADCC"
            Top             =   240
            Width           =   1632
            _Version        =   65537
            _ExtentX        =   2879
            _ExtentY        =   550
            _StockProps     =   93
            ForeColor       =   8388608
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   11
            BackColor       =   12632256
            Locked          =   -1  'True
         End
         Begin VB.Label Label15 
            Caption         =   "Limite Máximo:"
            Height          =   252
            Left            =   168
            TabIndex        =   27
            Top             =   300
            Width           =   1092
         End
      End
      Begin VB.Frame Frame20 
         Height          =   732
         Left            =   -74460
         TabIndex        =   21
         Top             =   3840
         Width           =   3420
         Begin VB.CheckBox chkControleQualidade 
            Caption         =   "Sempre Realizar Controle de Qualidade"
            Height          =   204
            Left            =   156
            TabIndex        =   26
            Tag             =   "ControleQualidade"
            Top             =   312
            Width           =   3120
         End
      End
      Begin VB.Frame Frame19 
         Caption         =   "Conversão de ch. UBB para Compensação"
         Height          =   684
         Left            =   -74604
         TabIndex        =   57
         Top             =   2136
         Width           =   3300
         Begin CURRENCYEDITLib.CurrencyEdit txtCompensaMal 
            Height          =   312
            Left            =   1320
            TabIndex        =   76
            Tag             =   "ValorCompensa_Mal"
            Top             =   240
            Width           =   1632
            _Version        =   65537
            _ExtentX        =   2879
            _ExtentY        =   550
            _StockProps     =   93
            ForeColor       =   8388608
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   11
            BackColor       =   -2147483643
         End
         Begin VB.Label Label14 
            Caption         =   "Caixa Robô:"
            Height          =   252
            Left            =   180
            TabIndex        =   67
            Top             =   300
            Width           =   972
         End
      End
      Begin VB.Frame Frame18 
         Caption         =   "Prazo Máx. Vcto. Títulos sem consulta"
         Height          =   684
         Left            =   -70992
         TabIndex        =   58
         Top             =   3000
         Width           =   3300
         Begin VB.TextBox txtPrazoVctoMal 
            ForeColor       =   &H00800000&
            Height          =   312
            Left            =   1380
            MaxLength       =   3
            TabIndex        =   79
            Tag             =   "PrazoVencimento_Mal"
            Top             =   240
            Width           =   1632
         End
         Begin VB.Label Label13 
            Caption         =   "Qtd. Dias:"
            Height          =   252
            Left            =   180
            TabIndex        =   68
            Top             =   300
            Width           =   792
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "Geração Automática de Diferenças"
         Height          =   684
         Left            =   -74604
         TabIndex        =   55
         Top             =   1272
         Width           =   3300
         Begin CURRENCYEDITLib.CurrencyEdit txtAjusteAutoMal 
            Height          =   312
            Left            =   1320
            TabIndex        =   74
            Tag             =   "ValorAjusteAuto_Mal"
            Top             =   240
            Width           =   1632
            _Version        =   65537
            _ExtentX        =   2879
            _ExtentY        =   550
            _StockProps     =   93
            ForeColor       =   8388608
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   11
            BackColor       =   -2147483643
         End
         Begin VB.Label Label12 
            Caption         =   "Prova Zero:"
            Height          =   192
            Left            =   180
            TabIndex        =   64
            Top             =   300
            Width           =   912
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "Limite Máx. de diferença a débito"
         Height          =   684
         Left            =   -70992
         TabIndex        =   56
         Top             =   2136
         Width           =   3252
         Begin CURRENCYEDITLib.CurrencyEdit txtAjusteManualMal 
            Height          =   312
            Left            =   1380
            TabIndex        =   77
            Tag             =   "ValorAjusteVincManual_Mal"
            Top             =   240
            Width           =   1632
            _Version        =   65537
            _ExtentX        =   2879
            _ExtentY        =   550
            _StockProps     =   93
            ForeColor       =   8388608
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   11
            BackColor       =   -2147483643
         End
         Begin VB.Label Label11 
            Caption         =   "Vínculo Manual:"
            Height          =   252
            Left            =   180
            TabIndex        =   66
            Top             =   300
            Width           =   1212
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Aprovação de SAQUE p/ Supervisor"
         Height          =   660
         Left            =   -74604
         TabIndex        =   53
         Top             =   456
         Width           =   3300
         Begin CURRENCYEDITLib.CurrencyEdit txtAlcadaMal 
            Height          =   312
            Left            =   1320
            TabIndex        =   72
            TabStop         =   0   'False
            Tag             =   "ValorAlcada_Mal"
            Top             =   240
            Width           =   1632
            _Version        =   65537
            _ExtentX        =   2879
            _ExtentY        =   550
            _StockProps     =   93
            ForeColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   11
            BackColor       =   -2147483633
            Locked          =   -1  'True
         End
         Begin VB.Label Label10 
            Caption         =   "Alçada:"
            Height          =   252
            Left            =   180
            TabIndex        =   62
            Top             =   276
            Width           =   672
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Aprovação de DEPÓSITO p/ Supervisor"
         Height          =   660
         Left            =   -70992
         TabIndex        =   54
         Top             =   456
         Width           =   3252
         Begin CURRENCYEDITLib.CurrencyEdit txtAlcadaDepMal 
            Height          =   312
            Left            =   1380
            TabIndex        =   73
            Tag             =   "ValorAlcadaDep_Mal"
            Top             =   240
            Width           =   1632
            _Version        =   65537
            _ExtentX        =   2879
            _ExtentY        =   550
            _StockProps     =   93
            ForeColor       =   8388608
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   11
            BackColor       =   -2147483643
         End
         Begin VB.Label Label9 
            Caption         =   "Alçada:"
            Height          =   252
            Left            =   180
            TabIndex        =   63
            Top             =   300
            Width           =   672
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Conversão de ch. UBB para Compensação"
         Height          =   732
         Left            =   -74604
         TabIndex        =   36
         Top             =   2616
         Width           =   3396
         Begin CURRENCYEDITLib.CurrencyEdit txtCompensaEnv 
            Height          =   312
            Left            =   1320
            TabIndex        =   50
            Tag             =   "ValorCompensa_Env"
            Top             =   240
            Width           =   1632
            _Version        =   65537
            _ExtentX        =   2879
            _ExtentY        =   550
            _StockProps     =   93
            ForeColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   11
            BackColor       =   -2147483633
            Locked          =   -1  'True
         End
         Begin VB.Label Label7 
            Caption         =   "Caixa Robô:"
            Height          =   252
            Left            =   180
            TabIndex        =   43
            Top             =   300
            Width           =   972
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Prazo Máx. Vcto. Títulos sem consulta"
         Height          =   732
         Left            =   -70896
         TabIndex        =   37
         Top             =   2616
         Width           =   3180
         Begin VB.TextBox txtPrazoVctoEnv 
            ForeColor       =   &H00800000&
            Height          =   312
            Left            =   1380
            MaxLength       =   3
            TabIndex        =   51
            Tag             =   "PrazoVencimento_Env"
            Top             =   240
            Width           =   1632
         End
         Begin VB.Label Label8 
            Caption         =   "Qtd. Dias:"
            Height          =   252
            Left            =   180
            TabIndex        =   44
            Top             =   300
            Width           =   792
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Geração Automática de Diferenças"
         Height          =   732
         Left            =   -74604
         TabIndex        =   34
         Top             =   1608
         Width           =   3396
         Begin CURRENCYEDITLib.CurrencyEdit txtAjusteAutoEnv 
            Height          =   312
            Left            =   1320
            TabIndex        =   48
            Tag             =   "ValorAjusteAuto_Env"
            Top             =   240
            Width           =   1632
            _Version        =   65537
            _ExtentX        =   2879
            _ExtentY        =   550
            _StockProps     =   93
            ForeColor       =   8388608
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   11
            BackColor       =   -2147483643
         End
         Begin VB.Label Label5 
            Caption         =   "Prova Zero:"
            Height          =   192
            Left            =   180
            TabIndex        =   41
            Top             =   300
            Width           =   912
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Limite Máx. de diferença a débito"
         Height          =   732
         Left            =   -70896
         TabIndex        =   35
         Top             =   1608
         Visible         =   0   'False
         Width           =   3180
         Begin CURRENCYEDITLib.CurrencyEdit txtAjusteManualEnv 
            Height          =   312
            Left            =   1380
            TabIndex        =   49
            Tag             =   "ValorAjusteVincManual_Env"
            Top             =   240
            Width           =   1632
            _Version        =   65537
            _ExtentX        =   2879
            _ExtentY        =   550
            _StockProps     =   93
            ForeColor       =   8388608
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   11
            BackColor       =   -2147483643
         End
         Begin VB.Label Label6 
            Caption         =   "Vínculo Manual:"
            Height          =   252
            Left            =   180
            TabIndex        =   42
            Top             =   300
            Width           =   1212
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Aprovação de SAQUE p/ Supervisor"
         Height          =   732
         Left            =   -74604
         TabIndex        =   32
         Top             =   600
         Width           =   3396
         Begin CURRENCYEDITLib.CurrencyEdit txtAlcadaEnv 
            Height          =   312
            Left            =   1320
            TabIndex        =   46
            TabStop         =   0   'False
            Tag             =   "ValorAlcada_Env"
            Top             =   240
            Width           =   1632
            _Version        =   65537
            _ExtentX        =   2879
            _ExtentY        =   550
            _StockProps     =   93
            ForeColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   11
            BackColor       =   -2147483633
            Locked          =   -1  'True
         End
         Begin VB.Label Label3 
            Caption         =   "Alçada:"
            Height          =   252
            Left            =   180
            TabIndex        =   39
            Top             =   300
            Width           =   672
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Aprovação de DEPÓSITO p/ Supervisor"
         Height          =   732
         Left            =   -70896
         TabIndex        =   33
         Top             =   600
         Width           =   3180
         Begin CURRENCYEDITLib.CurrencyEdit txtAlcadaDepEnv 
            Height          =   312
            Left            =   1380
            TabIndex        =   47
            Tag             =   "ValorAlcadaDep_Env"
            Top             =   240
            Width           =   1632
            _Version        =   65537
            _ExtentX        =   2879
            _ExtentY        =   550
            _StockProps     =   93
            ForeColor       =   8388608
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   11
            BackColor       =   -2147483643
         End
         Begin VB.Label Label4 
            Caption         =   "Alçada:"
            Height          =   252
            Left            =   180
            TabIndex        =   40
            Top             =   300
            Width           =   672
         End
      End
      Begin VB.Frame Frame4 
         Height          =   612
         Left            =   -74460
         TabIndex        =   17
         Top             =   480
         Width           =   6576
         Begin VB.TextBox txtAgencia 
            BackColor       =   &H8000000F&
            ForeColor       =   &H00800000&
            Height          =   312
            Left            =   3900
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   28
            Tag             =   "AgenciaCentral"
            Top             =   216
            Width           =   1032
         End
         Begin VB.Label Label1 
            Caption         =   "Agência Processadora:"
            Height          =   252
            Left            =   1560
            TabIndex        =   23
            Top             =   276
            Width           =   1812
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Autenticadora"
         Height          =   732
         Left            =   3840
         TabIndex        =   8
         Top             =   3516
         Width           =   3252
         Begin VB.ComboBox cmbAutentica 
            ForeColor       =   &H00800000&
            Height          =   288
            ItemData        =   "Parametros.frx":008C
            Left            =   90
            List            =   "Parametros.frx":0099
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   288
            Width           =   3012
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Scanner"
         Height          =   732
         Left            =   540
         TabIndex        =   7
         Top             =   3516
         Width           =   3252
         Begin VB.ComboBox cmbScanner 
            ForeColor       =   &H00800000&
            Height          =   288
            ItemData        =   "Parametros.frx":00BE
            Left            =   120
            List            =   "Parametros.frx":00CB
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   288
            Width           =   3012
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Diretório de Dados"
         Height          =   732
         Left            =   540
         TabIndex        =   4
         Top             =   600
         Width           =   6552
         Begin VB.CommandButton cmdDados 
            Caption         =   "..."
            Height          =   372
            Left            =   6060
            TabIndex        =   10
            Top             =   240
            Width           =   372
         End
         Begin VB.TextBox txtDirDados 
            BackColor       =   &H8000000F&
            ForeColor       =   &H00800000&
            Height          =   312
            Left            =   60
            Locked          =   -1  'True
            TabIndex        =   9
            Tag             =   "Dir_Dados"
            Text            =   "Text1"
            Top             =   270
            Width           =   5892
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Diretório de Imagens"
         Height          =   732
         Left            =   540
         TabIndex        =   5
         Top             =   1572
         Width           =   6552
         Begin VB.TextBox txtDirImagens 
            BackColor       =   &H8000000F&
            ForeColor       =   &H00800000&
            Height          =   312
            Left            =   60
            Locked          =   -1  'True
            TabIndex        =   11
            Tag             =   "Dir_Imagens"
            Text            =   "Text1"
            Top             =   270
            Width           =   5892
         End
         Begin VB.CommandButton cmdImagens 
            Caption         =   "..."
            Height          =   372
            Left            =   6060
            TabIndex        =   12
            Top             =   240
            Width           =   372
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Diretório de Trabalho"
         Height          =   732
         Left            =   540
         TabIndex        =   6
         Top             =   2520
         Width           =   6552
         Begin VB.TextBox txtDirTrabalho 
            BackColor       =   &H8000000F&
            ForeColor       =   &H00800000&
            Height          =   312
            Left            =   60
            Locked          =   -1  'True
            TabIndex        =   13
            Tag             =   "Dir_Trabalho"
            Text            =   "Text1"
            Top             =   270
            Width           =   5892
         End
         Begin VB.CommandButton cmdTrabalho 
            Caption         =   "..."
            Height          =   372
            Left            =   6060
            TabIndex        =   14
            Top             =   240
            Width           =   372
         End
      End
   End
End
Attribute VB_Name = "Parametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ScannerInicio           As Integer

Private Type tpModulo
    qryLeituraParametro         As rdoQuery     ' Leitura da tabela parametro
    qryAtualizaParametro        As rdoQuery     ' Atualiza a tabela parametros
    qryAtualizaLogAlteracao     As rdoQuery     ' Atualiza a tabela de log de alteracoes
    qryGetUsuario               As rdoQuery
    qryGetGrupoDocumento        As rdoQuery     ' Leitura da tabela Grupo de Documentos
    qryAtualizaAutentGrpDocto   As rdoQuery     ' Atualiza autenticação na Tabela Grupo Documento
    rsUsuario                   As rdoResultset
    NrAutenticacoes()           As Integer      ' Matriz para cada Grupo de Documento contendo
                                                ' Nro. Autenticações(Nova) , Nro. Grupo e Nro. Autenticação(Atual)
End Type
Private Modulo                  As tpModulo
Private Fechou                  As Boolean
'Pega o campo correspondente
Private Function PegarDescricao(ByVal Ctrl_Name As String)

    Dim sRetorno    As String

    Ctrl_Name = LCase(Ctrl_Name)

    Select Case LCase(Ctrl_Name)
        Case "txtdirdados": sRetorno = Frame1.Caption
        Case "txtdirimagens": sRetorno = Frame2.Caption
        Case "txtdirtrabalho": sRetorno = Frame3.Caption
        Case "txtagencia": sRetorno = Label1.Caption
        Case "txtinferior": sRetorno = Label2.Caption
        Case "chkcontrolequalidade": sRetorno = chkControleQualidade.Caption
        Case "txtalcadaenv": sRetorno = Frame8.Caption
        Case "txtajusteautoenv": sRetorno = Frame10.Caption
        Case "txtcompensaenv": sRetorno = Frame13.Caption
        Case "txtalcadadepenv": sRetorno = Frame7.Caption
        Case "txtajustemanualenv": sRetorno = Frame9.Caption
        Case "txtprazovctoenv": sRetorno = Frame12.Caption
        Case "txtmaxadccenv": sRetorno = Frame21.Caption
        Case "txtalcadamal": sRetorno = Frame15.Caption
        Case "txtajusteautomal": sRetorno = Frame17.Caption
        Case "txtcompensamal": sRetorno = Frame19.Caption
        Case "txtalcadadepmal": sRetorno = Frame14.Caption
        Case "txtalcadacoordmalote": sRetorno = fraAlcadaCoordMalote.Caption
        Case "txtajustemanualmal": sRetorno = Frame16.Caption
        Case "txtprazovctomal": sRetorno = Frame18.Caption
        Case "txtmaxadcc": sRetorno = Label15.Caption
        Case "txtajustecontabil": sRetorno = Label16.Caption
        Case "txtvalormaxdiflancto": sRetorno = Frame23.Caption
        Case "txtnr_autenticacoes": sRetorno = Frame24.Caption
        Case "txtvalorcompensanovo_mal": sRetorno = Frame27.Caption
        Case "txttmpbloqueio": sRetorno = Frame28.Caption
        Case "txtdiasinativo": sRetorno = Frame29.Caption
       'Case "txthoralimreccapaik": sRetorno = Left(Label24.Caption, Len(Label24.Caption) - 1)
        Case "CheckRecepcionaIK": sRetorno = CheckRecepcionaIK.Caption
        Case "txtdiastrocasenha": sRetorno = fraDiasTrocaSenha.Caption
        
    End Select

    PegarDescricao = sRetorno

End Function
Private Function UsuarioSuporte(ByVal User As String) As Boolean
    UsuarioSuporte = GrupoUsuario(User, eG_SUPORTE)
End Function
Private Sub cmbGrupoDocumentos_Click()

If cmbGrupoDocumentos.ListIndex = -1 Then Exit Sub

txtNr_Autenticacoes = Modulo.NrAutenticacoes(cmbGrupoDocumentos.ListIndex, 0, 0)

End Sub

Private Sub cmdDados_Click()

    Dim sDir        As String
    
    sDir = BrowseForFolder(Me, camMyComputer)
    
    If Trim(sDir) <> "" Then
        txtDirDados.Text = sDir
    End If

End Sub
Private Sub cmdImagens_Click()

    Dim sDir        As String
    
    sDir = BrowseForFolder(Me, camMyComputer)
    
    If Trim(sDir) <> "" Then
        txtDirImagens.Text = sDir
    End If

End Sub
Private Sub cmdOk_Click()

DoEvents
On Error GoTo ERRO_CMDOK

    Dim iRet            As Long
    Dim ScannerOk       As Boolean
    Dim Parametro       As tpGlobais
    Dim tb1             As rdoResultset
    Dim NumBoxes        As Long
    Dim MaxDocBox       As Long
    Dim BoxDefault      As Long
    Dim Threshold       As Long
    Dim Compress        As Long
    Dim Ctrl            As Object
    Dim sDescricao      As String
    Dim bAlterar        As Boolean
    Dim sValorAtual     As String
    Dim sValorAntigo    As String
    Dim i               As Integer, bAlteradaAutenticacao As Boolean
    
    Screen.MousePointer = vbHourglass
    rdoErrors.Clear

      
  'Tratamento - Envelope
  'Verificar se o prazo máximo para vencimento de titulos sem consulta é maior que 2 -
   If Val(txtPrazoVctoEnv.Text) < 2 Or Val(txtPrazoVctoMal.Text) < 2 Then
        MsgBox "O Prazo máximo para vencimento não pode ser menor que 2.", vbInformation + vbOKOnly, App.Title
        If Val(txtPrazoVctoEnv.Text) < 2 Then
            txtPrazoVctoEnv.Text = "2"
            SSTab.Tab = 2
            Screen.MousePointer = vbDefault
            txtPrazoVctoEnv.SetFocus
            txtPrazoVctoEnv.SelStart = 0
            txtPrazoVctoEnv.SelLength = Len(txtPrazoVctoEnv)
        End If
        'Tratamento - Malote
        'Verificar se o prazo máximo para vencimento de titulos sem consulta é maior que 2
        If Val(txtPrazoVctoMal.Text) < 2 Then
            txtPrazoVctoMal.Text = "2"
            SSTab.Tab = 3
            Screen.MousePointer = vbDefault
            txtPrazoVctoMal.SetFocus
            txtPrazoVctoMal.SelStart = 0
            txtPrazoVctoMal.SelLength = Len(txtPrazoVctoMal)
        End If
        
        Exit Sub
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Só pode digitar 1 ou 2 no campo "Numero de Autenticacoes'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For i = 0 To UBound(Modulo.NrAutenticacoes, 1)
        
        'Verifica se houve alteração no Nr. de Autenticação (Nova com Atual)
        If Modulo.NrAutenticacoes(i, 0, 0) <> Modulo.NrAutenticacoes(i, 0, 1) Then bAlteradaAutenticacao = True
        
        If Modulo.NrAutenticacoes(i, 0, 0) < 0 Or Modulo.NrAutenticacoes(i, 0, 0) > 2 Then
            cmbGrupoDocumentos.ListIndex = i
            cmbGrupoDocumentos.SetFocus
            SSTab.Tab = SSTab.TabsPerRow - 1
            Screen.MousePointer = vbDefault
            
            MsgBox "O campo 'Número de Autenticações' deve estar entre 0 ou 2.", vbExclamation
            Exit Sub
        End If
    Next
    
    If Val(txtDiasTrocaSenha.Text) < 1 Then
        MsgBox "O número de dias para expirar a senha não pode ser menor que 1.", vbInformation, App.Title
        SSTab.Tab = 4
        Screen.MousePointer = vbDefault
        txtDiasTrocaSenha.SetFocus
        txtDiasTrocaSenha.SelStart = 0
        txtDiasTrocaSenha.SelLength = Len(txtDiasTrocaSenha)
        Exit Sub
    End If
    
    '''''''''''''''''''''''''''''''''''''
    'Leitura dos parametros sem alteracao
    '''''''''''''''''''''''''''''''''''''
    With Modulo.qryLeituraParametro
        .rdoParameters(0) = Geral.DataProcessamento
        Set tb1 = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    If tb1.EOF Then
        MsgBox "Não foi possível ler informações da tabela Parâmetro.", vbCritical + vbOKOnly, App.Title
        Modulo.qryLeituraParametro.Close
        GoTo ERRO_CMDOK
    End If

    With Parametro
        .DiretorioDados = txtDirDados
        .DiretorioImagens = txtDirImagens
        .DiretorioTrabalho = txtDirTrabalho
        .Scanner = cmbScanner.ListIndex
        .autenticadora = cmbAutentica.ListIndex
        If txtAlcadaCoordMalote.Text = "" Then
            .ValorAlcadaCoord_Mal = 0
        Else
            .ValorAlcadaCoord_Mal = txtAlcadaCoordMalote.Text
        End If
    End With

    If ChecarParametros(Parametro) Then
        GravarOpcaoINI "Diversos", "Scanner", CStr(cmbScanner.ListIndex)
        GravarOpcaoINI "Diversos", "Autenticadora", CStr(cmbAutentica.ListIndex)
    Else
        GoTo ERRO_CMDOK
    End If

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Atualiza LogAlteracao                                '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    With Modulo.qryAtualizaLogAlteracao
        For Each Ctrl In Me.Controls
            If ((TypeOf Ctrl Is TextBox) Or _
                (TypeOf Ctrl Is CurrencyEdit) Or _
                (TypeOf Ctrl Is DateEdit) Or _
                (TypeOf Ctrl Is CheckBox)) And _
                (Trim(Ctrl.Tag) <> "") And _
                Ctrl.Tag <> "Nr_Autenticacoes" Then

                'Verifica o tipo de controle
                If TypeOf Ctrl Is CurrencyEdit Then
                    sValorAtual = CStr(Val(InserePonto(Format(Ctrl.Text, "000"))))
                    If CStr(tb1(Ctrl.Tag)) <> CStr(sValorAtual) Then
                        sValorAntigo = Format(tb1(Ctrl.Tag), "#######0.00")
                        sDescricao = PegarDescricao(Ctrl.Name)
                        bAlterar = True
                    End If
                ElseIf TypeOf Ctrl Is DateEdit Then
                    sValorAtual = CStr(Ctrl.InverseText)
                    If CStr(tb1(Ctrl.Tag)) <> CStr(Val(sValorAtual)) Then
                        sValorAtual = Format(Format(CStr(Ctrl.InverseText), "0000/00/00"), "dd/mm/yyyy")
                        sValorAntigo = Format(Format(tb1(Ctrl.Tag), "0000/00/00"), "dd/mm/yyyy")
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        'Pega descrição do campo para gravar na tabela LogAlteracao'
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        sDescricao = PegarDescricao(Ctrl.Name)
                        bAlterar = True
                    End If
               
                ElseIf CStr(tb1(Ctrl.Tag) & "") <> CStr(Ctrl) Then
                    sValorAtual = Ctrl
                    sValorAntigo = CStr(tb1(Ctrl.Tag) & "")
                    sDescricao = PegarDescricao(Ctrl.Name)
                    bAlterar = True
                End If

                'Insere o log do parâmetro alterado
                If bAlterar Then
                    .rdoParameters(1) = Geral.DataProcessamento
                    .rdoParameters(2) = Geral.Usuario
                    .rdoParameters(3) = sDescricao
                    .rdoParameters(4) = CStr(sValorAntigo)
                    .rdoParameters(5) = sValorAtual
                    .Execute
                    bAlterar = False
                    If (.rdoParameters(0) <> 0) Then
                        GoTo ERRO_CMDOK
                    End If
                End If
            End If
        Next
        
        'Se houve alteração do Nr. de Autenticações, gerar log de ocorrência
        If bAlteradaAutenticacao Then
            sDescricao = PegarDescricao("txtNr_Autenticacoes")
            
            For i = 0 To UBound(Modulo.NrAutenticacoes, 1)
                'Verifica se houve alteração no Nr. de Autenticação (Nova com Atual)
                If Modulo.NrAutenticacoes(i, 0, 0) <> Modulo.NrAutenticacoes(i, 0, 1) Then
                    sValorAtual = CStr(Modulo.NrAutenticacoes(i, 0, 0))     ' Nro. Autent. (Nova)
                    sValorAntigo = CStr(Modulo.NrAutenticacoes(i, 0, 1))    ' Nro. Autent. (Atual)
                    
                    With Modulo.qryAtualizaLogAlteracao
                        .rdoParameters(1) = Geral.DataProcessamento
                        .rdoParameters(2) = Geral.Usuario
                        .rdoParameters(3) = Left(sDescricao & "(" & Trim(cmbGrupoDocumentos.List(i)) & ")", 100)
                        .rdoParameters(4) = sValorAntigo
                        .rdoParameters(5) = sValorAtual
                        .Execute
                        If (.rdoParameters(0) <> 0) Then
                            GoTo ERRO_CMDOK
                        End If
                    End With
                End If
            Next
        
        End If
        
    End With

    tb1.Close
    Modulo.qryAtualizaLogAlteracao.Close

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Atualiza Parametros                                  '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    With Modulo.qryAtualizaParametro
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = Val(txtAgencia.Text)
        .rdoParameters(3) = Val(InserePonto(txtInferior.Text))
        .rdoParameters(4) = chkControleQualidade.Value
        .rdoParameters(5) = Val(InserePonto(txtAlcadaEnv.Text))
        .rdoParameters(6) = Val(InserePonto(txtAlcadaMal.Text))
        .rdoParameters(7) = Val(InserePonto(txtAlcadaDepEnv.Text))
        .rdoParameters(8) = Val(InserePonto(txtAlcadaDepMal.Text))
        .rdoParameters(9) = Val(InserePonto(txtAjusteAutoEnv.Text))
        .rdoParameters(10) = Val(InserePonto(txtAjusteAutoMal.Text))
        .rdoParameters(11) = Val(InserePonto(txtAjusteManualEnv.Text))
        .rdoParameters(12) = Val(InserePonto(txtAjusteManualMal.Text))
        .rdoParameters(13) = Val(InserePonto(txtCompensaEnv.Text))
        .rdoParameters(14) = Val(InserePonto(txtCompensaMal.Text))
        .rdoParameters(15) = Val(InserePonto(TxtMaxADCC.Text))
        .rdoParameters(16) = Val(txtPrazoVctoEnv.Text)
        .rdoParameters(17) = Val(txtPrazoVctoMal.Text)
        .rdoParameters(18) = txtDirDados
        .rdoParameters(19) = txtDirImagens
        .rdoParameters(20) = txtDirTrabalho
        .rdoParameters(21) = Val(InserePonto(txtAjusteContabil.Text))
        .rdoParameters(22) = Val(InserePonto(TxtValorMaxDifLancto.Text))
        .rdoParameters(23) = Val(InserePonto(txtAlcadaUBBEnv.Text))
        .rdoParameters(24) = Val(InserePonto(txtAlcadaUBBMal.Text))
        .rdoParameters(25) = Val(InserePonto(txtValorCompensaNovo_Mal.Text))
        .rdoParameters(26) = Val(txttmpbloqueio.Text)
        .rdoParameters(27) = Val(txtdiasinativo.Text)
        .rdoParameters(28) = IIf(CheckRecepcionaIK, "S", "N")
        .rdoParameters(29) = Val(InserePonto(txtAlcadaCoordMalote.Text))
        .rdoParameters(30) = Val(txtDiasTrocaSenha.Text)
        
        .Execute
    End With
    
    If (Modulo.qryAtualizaParametro.rdoParameters(0) <> 0) Then
        GoTo ERRO_CMDOK
    End If
    
    'Atualiza type global
    Geral.QtdeDiasTrocaSenha = txtDiasTrocaSenha
    
    Modulo.qryAtualizaParametro.Close

    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Atualiza Autenticações na Tabela Grupo Documento '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    For i = 0 To UBound(Modulo.NrAutenticacoes, 1)
        ' Verifica se houve alteração de Autenticações (Atual com Nova)
        If Modulo.NrAutenticacoes(i, 0, 0) <> Modulo.NrAutenticacoes(i, 0, 1) Then
            With Modulo.qryAtualizaAutentGrpDocto
                .rdoParameters(0).Direction = rdParamReturnValue
                .rdoParameters(1) = Modulo.NrAutenticacoes(i, 1, 0) ' Grupo do Documento
                .rdoParameters(2) = Modulo.NrAutenticacoes(i, 0, 0) ' Nro de Autenticações(Nova)
                .Execute
                
                If .rdoParameters(0) <> 0 Then GoTo ERRO_CMDOK
                
            End With
        End If
    Next
    
    '''''''''''''''''''''''''''''''''''''''''
    With Geral
        .Scanner = Val(PegarOpcaoINI("Diversos", "Scanner", "0"))
        .autenticadora = Val(PegarOpcaoINI("Diversos", "Autenticadora", "0"))
        .VIPSDLL = Val(PegarOpcaoINI("Diversos", "VipsDll", "0"))
    End With
    
    NumBoxes = Val(PegarOpcaoINI("Diversos", "NumBoxes", "1"))
    MaxDocBox = Val(PegarOpcaoINI("Diversos", "MaxDocBox", "200"))
    BoxDefault = Val(PegarOpcaoINI("Diversos", "BoxDefault", "0"))
    Threshold = Val(PegarOpcaoINI("Diversos", "CutBords", "50"))
    Compress = Val(PegarOpcaoINI("Diversos", "Compress_JPG", "30"))
    
    With Modulo.qryLeituraParametro
        .rdoParameters(0) = Geral.DataProcessamento
        Set tb1 = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    If tb1.EOF Then
        MsgBox "Não foi possível ler informações da tabela Parâmetro.", vbCritical + vbOKOnly, App.Title
        Modulo.qryLeituraParametro.Close
        GoTo ERRO_CMDOK
    Else
        Geral.DiretorioDados = tb1!Dir_Dados & IIf(Right(tb1!Dir_Dados, 1) <> "\", "\", "")
        Geral.DiretorioImagens = tb1!Dir_Imagens & "\" & Geral.DataProcessamento & "\"
        Geral.DiretorioTrabalho = tb1!Dir_Trabalho & IIf(Right(tb1!Dir_Trabalho, 1) <> "\", "\", "")
        Geral.AgenciaCentral = str(tb1!AgenciaCentral)
        Geral.Intervalo = tb1!TM_Pendente
        Geral.Atualizacao = tb1!TM_Atualizacao
        Geral.DataFinalRegraAntiga_Mal = tb1!DataFinalRegraAntiga_Mal
        Geral.ValorMaxADCC = tb1!ValorMaxADCC
        Geral.ValorAlcadaCoord_Mal = tb1!ValorAlcadaCoord_Mal
    End If
    tb1.Close
    Modulo.qryLeituraParametro.Close
    
    Set Autentica = Nothing

    If Geral.autenticadora = 1 Then
        Set Autentica = New Autentica_IBM
    ElseIf Geral.autenticadora = 2 Then
        Set Autentica = New Autentica_Procomp
    End If
    
    If ScannerInicio = 1 Then
        Call FinalizaDLLsVIPS
        'If Geral.VIPSDLL = eDllProservi Then
        '    iRet = MC93_DeInit()
        'Else
        '    VIPS_Done
        'End If
    End If
    
    ScannerOk = False
  
    iRet = 1
    If Geral.Scanner = escnVIPS Then
        If Geral.VIPSDLL = eDllProservi Then
            iRet = MC93_SetImagem(3)
            If iRet = 1 Then
              iRet = MC93_SetLeitora(3)
              If iRet = 1 Then
                iRet = MC93_SetDPI(100)
                If iRet = 1 Then
                  iRet = MC93_SetAltura(420)
                  If iRet = 1 Then
                    iRet = MC93_SetComPort(1)
                    If iRet = 1 Then
                      iRet = MC93_SetImageDirectory(Geral.DiretorioImagens)
                      If iRet = 1 Then
                        iRet = MC93_CutBords(1)
                        If iRet = 1 Then
                          iRet = MC93_Init()
                          If iRet = 1 Then
                              ScannerOk = True
                          End If
                        End If
                      End If
                    End If
                  End If
                End If
              End If
            End If
            If Not ScannerOk Then
              MsgBox "Não foi possível inicializar a VIPS." & vbCr & "Erro: " & iRet, vbExclamation + vbOKOnly, App.Title
            End If
        
        ElseIf Geral.VIPSDLL = eDllNovaUBB Then 'VipsDll (Nova Versão)
            tSC_ParamDLL.BoxDefault = BoxDefault
            tSC_ParamDLL.MaxDocBox = MaxDocBox
            tSC_ParamDLL.NumBoxes = NumBoxes
            
            If InicializarVips Then
                ScannerOk = True
                bInicializou = True
            End If
            
        ' VipsDll do Unibanco
        Else
            VIPS_SetBoxes (NumBoxes)
            VIPS_SetMaxDocBox (MaxDocBox)
            VIPS_SetBoxDefault (BoxDefault)
            VIPS_SetCompress (Compress)
            VIPS_SetCutBords (Threshold)
            VIPS_SetCameraFile ("Doc100.cpf")
            VIPS_SetImageDirectory (Geral.DiretorioImagens)
            iRet = VIPS_Init()
            If iRet <> 0 Then
                MsgBox "Não foi possível inicializar a VIPS." & vbCr & "Erro: " & iRet, vbExclamation + vbOKOnly, App.Title
            Else
                ScannerOk = True
            End If
        End If

    ElseIf Geral.Scanner = escnCanonLS500 Then
      ' Inicializção da LS500 e Canon
      iRet = LS_ProcuraLS500(string1, string2, string3)
      If iRet = 0 Then
          MsgBox "Não foi possível localizar o scanner.", vbExclamation + vbOKOnly, App.Title
      Else
          iRet = LS_SetNumGauges(1)
          iRet = LS_Lapso(30)           'SCSI antiga/nova
          iRet = LS_SetSepara(0)        '1- separa
                                        '0- não separa
          iRet = LS_SetTimeOut(500)     '1/2 segundo
          iRet = LS_SetImage(3)         '(1) digitaliza só frente
                                        '(2) digitaliza só verso
                                        '(3) digitaliza frente e verso
          ScannerOk = True
      End If
    ElseIf Geral.Scanner = escnSemScanner Then
        If Geral.VIPSDLL = eDllNovaUBB Then
            Call FinalizaDLLsVIPS
        End If
    End If
    
    If ScannerOk Then
      ' habilitar menu no form principal
      Principal.mnuCapCaptura(5).Enabled = True
      Principal.MnuCapRecaptura(7).Enabled = True
    Else
      ' desabilitar menu no form principal
      Principal.mnuCapCaptura(5).Enabled = False
      Principal.MnuCapRecaptura(7).Enabled = False
    End If

    Screen.MousePointer = vbDefault
    Unload Me

    Exit Sub

ERRO_CMDOK:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Não foi possível atualizar os parâmetros.", Err, rdoErrors)
        Case vbCancel
            Unload Me
        Case vbRetry
            Resume
    End Select
End Sub
Private Sub CmdSair_Click()
    Unload Me
End Sub
Private Sub cmdTrabalho_Click()

    Dim sDir        As String
    
    sDir = BrowseForFolder(Me, camMyComputer)
    
    If Trim(sDir) <> "" Then
        txtDirTrabalho.Text = sDir
    End If

End Sub
Private Sub Form_Activate()
   'Inclusão de chamada a rotina AtualizaAtividade
    Call AtualizaAtividade(22)

    If Fechou Then
        Fechou = False
        Unload Me
    End If
    
   'Posicionamento da tela de parâmetros
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
End Sub
Private Sub Form_Load()

    On Error GoTo ERRO_LOAD

    Dim tb As rdoResultset
    Fechou = False

    SSTab.Tab = 0
    cmbScanner.ListIndex = Val(PegarOpcaoINI("Diversos", "Scanner", "0"))
    cmbAutentica.ListIndex = Val(PegarOpcaoINI("Diversos", "Autenticadora", "0"))
    ScannerInicio = cmbScanner.ListIndex

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Cria query para atualizar a tabela de log de parametros
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set Modulo.qryAtualizaLogAlteracao = Geral.Banco.CreateQuery("", "{? = call GravarLogAlteracao (?,?,?,?,?)}")

    '''''''''''''''''''''''''''''''''''''''''''''''''
    ' Cria query para atualizar a tabela parametros '
    '''''''''''''''''''''''''''''''''''''''''''''''''
    Set Modulo.qryAtualizaParametro = Geral.Banco.CreateQuery("", "{? = call GravarParametro (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)}")

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Cria query para atualizar a Tabela Grupo Documentos '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set Modulo.qryAtualizaAutentGrpDocto = Geral.Banco.CreateQuery("", "{? = call AlteraAutenticacaoGrupoDocto(?,?)}")
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Cria query para a leitura da tabela parametros '
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    Set Modulo.qryLeituraParametro = Geral.Banco.CreateQuery("", "{call LerParametro(?)}")

    Set Modulo.qryGetUsuario = Geral.Banco.CreateQuery("", "{call GetUsuario (?)}")

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Cria query para a leitura da tabela Grupo de Documentos '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set Modulo.qryGetGrupoDocumento = Geral.Banco.CreateQuery("", "{call GetGrupoDocumento (?)}")
    
    ''''''''''''''''''''''''''''''''
    ' Leitura da tabela parametros '
    ''''''''''''''''''''''''''''''''
    With Modulo.qryLeituraParametro
        .rdoParameters(0) = Geral.DataProcessamento
        Set tb = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With

    If tb.EOF Then
        MsgBox "Não foi localizado registro desta data na tabela Parametros.", vbExclamation + vbOKOnly, App.Title
        Unload Me
    Else
        txtAgencia.Text = tb!AgenciaCentral
        txtInferior.Text = RetiraPonto(Format(tb!ValorInferior, ".00"))
        chkControleQualidade.Value = tb!ControleQualidade
        txtAlcadaEnv.Text = RetiraPonto(Format(tb!ValorAlcada_Env, ".00"))
        txtAlcadaMal.Text = RetiraPonto(Format(tb!ValorAlcada_Mal, ".00"))
        txtAlcadaDepEnv.Text = RetiraPonto(Format(tb!ValorAlcadaDep_Env, ".00"))
        txtAlcadaDepMal.Text = RetiraPonto(Format(tb!ValorAlcadaDep_Mal, ".00"))
        txtAlcadaCoordMalote.Text = RetiraPonto(Format(tb!ValorAlcadaCoord_Mal, ".00"))
        txtAjusteAutoEnv.Text = RetiraPonto(Format(tb!ValorAjusteAuto_Env, ".00"))
        txtAjusteAutoMal.Text = RetiraPonto(Format(tb!ValorAjusteAuto_Mal, ".00"))
        txtAjusteManualEnv.Text = RetiraPonto(Format(tb!ValorAjusteVincManual_Env, ".00"))
        txtAjusteManualMal.Text = RetiraPonto(Format(tb!ValorAjusteVincManual_Mal, ".00"))
        txtCompensaEnv.Text = RetiraPonto(Format(tb!ValorCompensa_Env, ".00"))
        txtCompensaMal.Text = RetiraPonto(Format(tb!ValorCompensa_Mal, ".00"))
        TxtMaxADCC.Text = RetiraPonto(Format(tb!ValorMaxADCC, ".00"))
        txtPrazoVctoEnv.Text = tb!PrazoVencimento_Env
        txtPrazoVctoMal.Text = tb!PrazoVencimento_Mal
        txtDirDados = tb!Dir_Dados
        txtDirImagens = tb!Dir_Imagens
        txtDirTrabalho = tb!Dir_Trabalho
        TxtValorMaxDifLancto.Text = RetiraPonto(Format(tb!LimiteMaxDifLancto_Mal, ".00"))
        txtAjusteContabil.Text = RetiraPonto(Format(tb!ValorAjusteContabil, ".00"))
        txtAlcadaUBBEnv.Text = RetiraPonto(Format(tb!ValorAlcadaOutros_Env, ".00"))
        txtAlcadaUBBMal.Text = RetiraPonto(Format(tb!ValorAlcadaOutros_Mal, ".00"))
        txtValorCompensaNovo_Mal.Text = RetiraPonto(Format(tb!ValorCompensaNovo_Mal, ".00"))
        txttmpbloqueio = tb!tmpBloqueio
        txtdiasinativo = tb!Diasinativo
        txtDiasTrocaSenha = tb!QtdeDiasTrocaSenha
        
        If IsNull(tb!RecepcionaIK) Or tb!RecepcionaIK = "N" Then
            CheckRecepcionaIK.Value = 0
        Else
            CheckRecepcionaIK.Value = 1
        End If
        
        'txtHoraLimRecCapaIK.Text = IIf(IsNull(tb!HoraLimiteRec_IK), "     ", tb!HoraLimiteRec_IK)
        
        '''''''''''''''''''''''''''''''''''''''''''''''
        ' Carrega dados da tabela Grupo de Documentos '
        '''''''''''''''''''''''''''''''''''''''''''''''
        Set tb = Nothing
        With Modulo.qryGetGrupoDocumento
            .rdoParameters(0) = Null
            Set tb = .OpenResultset(rdOpenStatic, rdConcurReadOnly)
        End With
        
        If tb.EOF Then
            MsgBox "Não foi localizado registro na tabela Grupo de Documentos.", vbExclamation + vbOKOnly, App.Title
            Unload Me
        End If

        Erase Modulo.NrAutenticacoes
        ReDim Modulo.NrAutenticacoes(tb.RowCount - 1, 1, 1) 'Nr autenticação(Nova), NrGrupo, Nro. Autent.(Atual)
        
        Do Until tb.EOF
            cmbGrupoDocumentos.List(tb.AbsolutePosition - 1) = tb!Descricao
            Modulo.NrAutenticacoes(tb.AbsolutePosition - 1, 0, 0) = tb!Nr_Autenticacoes ' Nro de Autenticações(Nova)
            Modulo.NrAutenticacoes(tb.AbsolutePosition - 1, 1, 0) = tb!IdGrupo          ' Identifica Grupo Documento
            Modulo.NrAutenticacoes(tb.AbsolutePosition - 1, 0, 1) = tb!Nr_Autenticacoes ' Nro de Autenticações(Atual)
            tb.MoveNext
        Loop
        
    End If
    
    tb.Close

    If UsuarioSuporte(Geral.Usuario) Then
        cmdDados.Enabled = True
        cmdImagens.Enabled = True
        cmdTrabalho.Enabled = True
    Else
        cmdDados.Enabled = False
        cmdImagens.Enabled = False
        cmdTrabalho.Enabled = False
    End If
    
    Exit Sub
    
ERRO_LOAD:
    Select Case TratamentoErro("Não foi possível localizar os dados dos parâmetros.", Err, rdoErrors)
        Case vbCancel
            Fechou = True
            Unload Me
        Case vbRetry
            Resume
    End Select
End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    If Not Modulo.qryGetUsuario Is Nothing Then Modulo.qryGetUsuario.Close
    If Not Modulo.qryGetGrupoDocumento Is Nothing Then Modulo.qryGetGrupoDocumento.Close
    If Not Modulo.qryAtualizaAutentGrpDocto Is Nothing Then Modulo.qryAtualizaAutentGrpDocto.Close
    
    Erase Modulo.NrAutenticacoes
    
End Sub

Private Sub txtAgencia_LostFocus()
    With txtAgencia
        If .Text <> "" Then
            If Not IsNumeric(.Text) Then
                MsgBox "Código de agência inválido!", vbExclamation + vbOKOnly, App.Title
                SelecionarTexto txtAgencia
            End If
        End If
    End With
End Sub
Private Sub txtAjusteAutoEnv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys ("{TAB}")
    End If
End Sub
Private Sub txtAjusteAutoMal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys ("{TAB}")
    End If
End Sub
Private Sub txtAjusteManualEnv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys ("{TAB}")
    End If
End Sub
Private Sub txtAjusteManualMal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys ("{TAB}")
    End If
End Sub
Private Sub txtAlcadaCoordMalote_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        SendKeys ("{TAB}")
    End If

End Sub
Private Sub txtAlcadaDepEnv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys ("{TAB}")
    End If
End Sub
Private Sub txtAlcadaDepMal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys ("{TAB}")
    End If
End Sub
Private Sub txtAlcadaEnv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys ("{TAB}")
    End If
End Sub
Private Sub txtAlcadaMal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys ("{TAB}")
    End If
End Sub
Private Sub txtCompensaEnv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys ("{TAB}")
    End If
End Sub

Private Sub txtCompensaMal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys ("{TAB}")
    End If
End Sub
Private Sub txtdiasinativo_GotFocus()
   With txtdiasinativo
        .SelStart = 0
        .SelLength = (Len(txtdiasinativo))
   End With
End Sub

Private Sub txtDiasTrocaSenha_GotFocus()
   With txtDiasTrocaSenha
        .SelStart = 0
        .SelLength = (Len(txtDiasTrocaSenha))
   End With

End Sub

Private Sub txtNr_Autenticacoes_GotFocus()
    SelecionarTexto txtNr_Autenticacoes
End Sub
Private Sub txtNr_Autenticacoes_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys ("{TAB}")
    End If
End Sub
Private Sub txtNr_Autenticacoes_KeyPress(KeyAscii As Integer)
    If cmbGrupoDocumentos.ListIndex = -1 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    If KeyAscii = 13 Then KeyAscii = Asc(txtNr_Autenticacoes)
    SoNumero KeyAscii
    If Not (KeyAscii = 48 Xor KeyAscii = 49 Xor KeyAscii = 50 Xor KeyAscii = 8) Then
        MsgBox "Nro. de autenticações deve ser 1 ou 2", vbInformation, App.Title
        KeyAscii = 0
    End If
    
End Sub
Private Sub txtNr_Autenticacoes_LostFocus()
If Me.cmbGrupoDocumentos.ListIndex = -1 Then Exit Sub

    If Val(txtNr_Autenticacoes.Text) <> Modulo.NrAutenticacoes(cmbGrupoDocumentos.ListIndex, 0, 0) Then
        
        'Atualiza Nro Autenticações (Nova)
        Modulo.NrAutenticacoes(cmbGrupoDocumentos.ListIndex, 0, 0) = Val(txtNr_Autenticacoes.Text)
    
    End If

End Sub
Private Sub txtPrazoVctoEnv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys ("{TAB}")
    End If
End Sub
Private Sub txtPrazoVctoEnv_LostFocus()
    With txtPrazoVctoEnv
        If .Text <> "" Then
            If Not IsNumeric(.Text) Then
                MsgBox "Quantidade de dias deve ser um valor numérico!", vbExclamation + vbOKOnly, App.Title
                .SelStart = 0
                .SelLength = Len(Trim(.Text))
                .SetFocus
            End If
        End If
    End With
End Sub
Private Sub txtPrazoVctoMal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys ("{TAB}")
    End If
End Sub
Private Sub txtPrazoVctoMal_LostFocus()
    With txtPrazoVctoMal
        If .Text <> "" Then
            If Not IsNumeric(.Text) Then
                MsgBox "Quantidade de dias deve ser uma valor numérico!", vbExclamation + vbOKOnly, App.Title
                .SelStart = 0
                .SelLength = Len(Trim(.Text))
                .SetFocus
            End If
        End If
    End With
End Sub
Private Sub txttmpbloqueio_GotFocus()
    With txttmpbloqueio
        .SelStart = 0
        .SelLength = (Len(txttmpbloqueio))
    End With
End Sub
