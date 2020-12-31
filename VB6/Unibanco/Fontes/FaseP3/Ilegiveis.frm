VERSION 5.00
Object = "{00080003-B1BA-11CE-ABC6-F5B2E79D9E3F}#8.0#0"; "LTOCX80N.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Ilegiveis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ilegíveis"
   ClientHeight    =   8604
   ClientLeft      =   240
   ClientTop       =   492
   ClientWidth     =   11892
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8604
   ScaleWidth      =   11892
   Begin VB.CheckBox chkFiltro 
      Caption         =   "Filtrar capas"
      Height          =   204
      Left            =   120
      TabIndex        =   84
      Top             =   3468
      Width           =   1500
   End
   Begin VB.PictureBox Picture5 
      Height          =   252
      Left            =   120
      ScaleHeight     =   204
      ScaleWidth      =   3108
      TabIndex        =   81
      Top             =   384
      Width           =   3156
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Módulo anterior"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   192
         Left            =   1752
         TabIndex        =   83
         Top             =   0
         Width           =   1308
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Capa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   192
         Left            =   144
         TabIndex        =   82
         Top             =   0
         Width           =   456
      End
   End
   Begin VB.PictureBox frmLocalizar 
      Height          =   1272
      Left            =   4536
      ScaleHeight     =   1224
      ScaleWidth      =   2604
      TabIndex        =   73
      Top             =   1962
      Visible         =   0   'False
      Width           =   2652
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   300
         Left            =   1464
         TabIndex        =   77
         Top             =   816
         Width           =   972
      End
      Begin VB.CommandButton cmdProcurar 
         Caption         =   "&Localizar"
         Height          =   300
         Left            =   144
         TabIndex        =   75
         Top             =   816
         Width           =   972
      End
      Begin VB.TextBox txtNumEnvMal 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         MaxLength       =   18
         TabIndex        =   74
         Top             =   384
         Width           =   2304
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número Envelope/Malote"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   76
         Top             =   96
         Width           =   2232
      End
   End
   Begin VB.PictureBox PicTiposDoc 
      Height          =   4680
      Left            =   3336
      ScaleHeight     =   4632
      ScaleWidth      =   5124
      TabIndex        =   43
      Top             =   1968
      Visible         =   0   'False
      Width           =   5172
      Begin TabDlg.SSTab TabTipoDoc 
         Height          =   4224
         Left            =   276
         TabIndex        =   44
         Top             =   240
         Visible         =   0   'False
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   7451
         _Version        =   393216
         TabHeight       =   474
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Genéricos [ F6 ]"
         TabPicture(0)   =   "Ilegiveis.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "CmdFecharTiposDocto(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "frmTab(0)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Tributos [ F7 ]"
         TabPicture(1)   =   "Ilegiveis.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "CmdFecharTiposDocto(1)"
         Tab(1).Control(1)=   "frmTab(2)"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "Diversos [ F8 ]"
         TabPicture(2)   =   "Ilegiveis.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "frmTab(1)"
         Tab(2).Control(1)=   "CmdFecharTiposDocto(2)"
         Tab(2).ControlCount=   2
         Begin VB.Frame frmTab 
            Height          =   3384
            Index           =   0
            Left            =   140
            TabIndex        =   59
            Top             =   300
            Width           =   4270
            Begin VB.OptionButton OptGenericos 
               Caption         =   "(C) Lançamento Interno"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.6
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   11
               Left            =   180
               TabIndex        =   79
               Top             =   3072
               Width           =   2388
            End
            Begin VB.OptionButton OptGenericos 
               Caption         =   "(6) Capa de Envelope"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.6
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   5
               Left            =   180
               TabIndex        =   72
               Top             =   1518
               Width           =   3444
            End
            Begin VB.OptionButton OptGenericos 
               Caption         =   "(9) Capa de Malote"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.6
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   8
               Left            =   180
               TabIndex        =   71
               Top             =   2292
               Width           =   3336
            End
            Begin VB.OptionButton OptGenericos 
               Caption         =   "(B) Capa OCT"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.6
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   10
               Left            =   180
               TabIndex        =   68
               Top             =   2808
               Width           =   2388
            End
            Begin VB.OptionButton OptGenericos 
               Caption         =   "(A) OCT"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.6
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   9
               Left            =   180
               TabIndex        =   67
               Top             =   2550
               Width           =   1536
            End
            Begin VB.OptionButton OptGenericos 
               Caption         =   "(8) Cartão Crédito Avulso"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.6
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   7
               Left            =   180
               TabIndex        =   66
               Top             =   2034
               Width           =   3336
            End
            Begin VB.OptionButton OptGenericos 
               Caption         =   "(7) Autorização de Débito"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.6
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   6
               Left            =   180
               TabIndex        =   65
               Top             =   1776
               Width           =   2832
            End
            Begin VB.OptionButton OptGenericos 
               Caption         =   "(5) Cód. Barras com Valor de Ref."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.6
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   4
               Left            =   180
               TabIndex        =   64
               Top             =   1260
               Width           =   3912
            End
            Begin VB.OptionButton OptGenericos 
               Caption         =   "(4) Ficha de Compensação"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.6
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   3
               Left            =   180
               TabIndex        =   63
               Top             =   1002
               Width           =   3636
            End
            Begin VB.OptionButton OptGenericos 
               Caption         =   "(3) Concessionária"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.6
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   2
               Left            =   180
               TabIndex        =   62
               Top             =   744
               Width           =   2652
            End
            Begin VB.OptionButton OptGenericos 
               Caption         =   "(2) Depósito"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.6
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   1
               Left            =   180
               TabIndex        =   61
               Top             =   486
               Width           =   1464
            End
            Begin VB.OptionButton OptGenericos 
               Caption         =   "(1) Cheque"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.6
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   0
               Left            =   180
               TabIndex        =   60
               Top             =   228
               Width           =   1320
            End
         End
         Begin VB.Frame frmTab 
            Height          =   3144
            Index           =   2
            Left            =   -74860
            TabIndex        =   53
            Top             =   300
            Width           =   4270
            Begin VB.OptionButton OptTributos 
               Caption         =   "(7) FGTS - com código de barras"
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
               Height          =   252
               Index           =   6
               Left            =   144
               TabIndex        =   78
               Top             =   2160
               Width           =   3960
            End
            Begin VB.OptionButton OptTributos 
               Caption         =   "(3) FGTS"
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
               Height          =   252
               Index           =   2
               Left            =   144
               TabIndex        =   69
               Top             =   864
               Width           =   1944
            End
            Begin VB.OptionButton OptTributos 
               Caption         =   "(6) GPS"
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
               Height          =   300
               Index           =   5
               Left            =   144
               TabIndex        =   58
               Top             =   1800
               Width           =   1848
            End
            Begin VB.OptionButton OptTributos 
               Caption         =   "(1) DARF - Simples"
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
               Height          =   252
               Index           =   0
               Left            =   144
               TabIndex        =   57
               Top             =   240
               Width           =   2160
            End
            Begin VB.OptionButton OptTributos 
               Caption         =   "(2) DARF - Preto"
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
               Height          =   252
               Index           =   1
               Left            =   144
               TabIndex        =   56
               Top             =   552
               Width           =   1944
            End
            Begin VB.OptionButton OptTributos 
               Caption         =   "(4) GARE"
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
               Height          =   252
               Index           =   3
               Left            =   144
               TabIndex        =   55
               Top             =   1176
               Width           =   1188
            End
            Begin VB.OptionButton OptTributos 
               Caption         =   "(5) DARM"
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
               Height          =   252
               Index           =   4
               Left            =   144
               TabIndex        =   54
               Top             =   1488
               Width           =   1440
            End
         End
         Begin VB.Frame frmTab 
            Height          =   3144
            Index           =   1
            Left            =   -74860
            TabIndex        =   48
            Top             =   300
            Width           =   4270
            Begin VB.OptionButton OptDiversos 
               Caption         =   "(1) Título de Outros Bancos s/ cod. Barras"
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
               Height          =   252
               Index           =   0
               Left            =   120
               TabIndex        =   52
               Top             =   240
               Width           =   4032
            End
            Begin VB.OptionButton OptDiversos 
               Caption         =   "(2) Unicobrança Registrada"
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
               Height          =   252
               Index           =   1
               Left            =   120
               TabIndex        =   51
               Top             =   600
               Width           =   2880
            End
            Begin VB.OptionButton OptDiversos 
               Caption         =   "(3) Unicobrança Especial"
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
               Height          =   252
               Index           =   2
               Left            =   120
               TabIndex        =   50
               Top             =   960
               Width           =   2736
            End
            Begin VB.OptionButton OptDiversos 
               Caption         =   "(4) Arrecadação Convencional"
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
               Height          =   252
               Index           =   3
               Left            =   120
               TabIndex        =   49
               Top             =   1320
               Width           =   3012
            End
         End
         Begin VB.CommandButton CmdFecharTiposDocto 
            Caption         =   "&Fechar"
            Height          =   312
            Index           =   0
            Left            =   1740
            TabIndex        =   47
            Top             =   3804
            Width           =   1068
         End
         Begin VB.CommandButton CmdFecharTiposDocto 
            Caption         =   "&Fechar"
            Height          =   312
            Index           =   1
            Left            =   -73260
            TabIndex        =   46
            Top             =   3516
            Width           =   1068
         End
         Begin VB.CommandButton CmdFecharTiposDocto 
            Caption         =   "&Fechar"
            Height          =   312
            Index           =   2
            Left            =   -73260
            TabIndex        =   45
            Top             =   3516
            Width           =   1068
         End
      End
   End
   Begin VB.PictureBox FrmPesquisa 
      Height          =   1932
      Left            =   3060
      ScaleHeight     =   1884
      ScaleWidth      =   5724
      TabIndex        =   32
      Top             =   1968
      Visible         =   0   'False
      Width           =   5772
      Begin VB.CommandButton CmdFecharPesquisa 
         Caption         =   "&Fechar"
         Height          =   312
         Left            =   2184
         TabIndex        =   40
         Top             =   1464
         Width           =   1068
      End
      Begin ComctlLib.ProgressBar Progress 
         Height          =   300
         Left            =   336
         TabIndex        =   41
         Top             =   912
         Width           =   4932
         _ExtentX        =   8700
         _ExtentY        =   529
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Pesquisando por Documentos Ilegíveis. Aguarde ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   348
         TabIndex        =   42
         Top             =   576
         Width           =   4776
      End
   End
   Begin VB.Timer TmrAtualiza 
      Enabled         =   0   'False
      Interval        =   50000
      Left            =   9852
      Top             =   24
   End
   Begin VB.PictureBox Picture1 
      Height          =   264
      Left            =   120
      ScaleHeight     =   216
      ScaleWidth      =   1752
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   48
      Width           =   1800
      Begin VB.Label LblEnv_Mal 
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
         Height          =   228
         Left            =   108
         TabIndex        =   33
         Top             =   0
         Width           =   1296
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   252
      Left            =   3312
      ScaleHeight     =   204
      ScaleWidth      =   7032
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   384
      Width           =   7080
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Recap."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   192
         Left            =   2520
         TabIndex        =   80
         Top             =   0
         Width           =   612
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Duplic."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   192
         Left            =   1896
         TabIndex        =   39
         Top             =   0
         Width           =   588
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Nro."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   192
         Left            =   84
         TabIndex        =   30
         Top             =   0
         Width           =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Vínculo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   192
         Left            =   648
         TabIndex        =   29
         Top             =   0
         Width           =   624
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Ocorr."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   192
         Left            =   1356
         TabIndex        =   28
         Top             =   0
         Width           =   516
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Documento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   192
         Left            =   3312
         TabIndex        =   27
         Top             =   0
         Width           =   948
      End
      Begin VB.Label Label6 
         Caption         =   "Valor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   216
         Left            =   6276
         TabIndex        =   26
         Top             =   0
         Width           =   540
      End
   End
   Begin VB.PictureBox PctMalote 
      Height          =   264
      Left            =   3924
      ScaleHeight     =   216
      ScaleWidth      =   1176
      TabIndex        =   23
      Top             =   48
      Width           =   1224
      Begin VB.Label Label11 
         Caption         =   "Nro. Malote"
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
         Height          =   228
         Left            =   36
         TabIndex        =   24
         Top             =   0
         Width           =   1092
      End
   End
   Begin VB.PictureBox Picture4 
      Height          =   264
      Left            =   1980
      ScaleHeight     =   216
      ScaleWidth      =   528
      TabIndex        =   21
      Top             =   48
      Width           =   576
      Begin VB.Label Label12 
         Caption         =   "Lote"
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
         Height          =   228
         Left            =   24
         TabIndex        =   22
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.Frame FrmCmd 
      Height          =   4524
      Left            =   10404
      TabIndex        =   20
      Top             =   -72
      Width           =   1452
      Begin VB.CommandButton cmdEnviarCSP 
         Caption         =   "Enviar C&SP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   48
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1152
         Width           =   1392
      End
      Begin VB.CommandButton cmdComentario 
         Caption         =   "Co&mentário"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   36
         MaskColor       =   &H00404080&
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1812
         Width           =   1392
      End
      Begin VB.CommandButton CmdRecaptura 
         Caption         =   "Reca&ptura"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   36
         MaskColor       =   &H00404080&
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2148
         Width           =   1392
      End
      Begin VB.CommandButton CmdLocalizar 
         Caption         =   "&Localizar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   36
         MaskColor       =   &H00404080&
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2484
         Width           =   1392
      End
      Begin VB.CommandButton CmdTrocaOrdem 
         Caption         =   "&Troca de Ordem"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   36
         MaskColor       =   &H00404080&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1488
         Width           =   1392
      End
      Begin VB.CommandButton CmdCalculadora 
         Caption         =   "&Calculadora"
         Height          =   288
         Left            =   36
         TabIndex        =   1
         Top             =   478
         Width           =   1392
      End
      Begin VB.CommandButton CmdAtualizar 
         Caption         =   "&Atualizar"
         Height          =   288
         Left            =   36
         TabIndex        =   0
         Top             =   144
         Width           =   1392
      End
      Begin VB.CommandButton CmdTipoDocto 
         Caption         =   "Tipo &Docto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   36
         MaskColor       =   &H00404080&
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3156
         Width           =   1392
      End
      Begin VB.CommandButton cmdEnviarSupervisor 
         Caption         =   "&Reenvia Complem."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   36
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   812
         Width           =   1392
      End
      Begin VB.CommandButton CmdOcorrencia 
         Caption         =   "&Ocorrência"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   36
         MaskColor       =   &H00404080&
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2820
         Width           =   1392
      End
      Begin VB.CommandButton cmdExcluirCapa 
         Caption         =   "E&xcluir Capa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   36
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   3816
         Width           =   1392
      End
      Begin VB.CommandButton cmdEncerrar 
         Caption         =   "&Encerrar Capa"
         Height          =   288
         Left            =   36
         TabIndex        =   10
         Top             =   3492
         Width           =   1392
      End
      Begin VB.CommandButton cmdFechar 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   288
         Left            =   36
         TabIndex        =   12
         Top             =   4152
         Width           =   1392
      End
   End
   Begin VB.Frame FrmCmdImagem 
      Height          =   4080
      Left            =   10404
      TabIndex        =   19
      Top             =   4452
      Width           =   1452
      Begin VB.CommandButton cmdAuditoria 
         Caption         =   "A&uditoria"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   288
         Picture         =   "Ilegiveis.frx":0054
         Style           =   1  'Graphical
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   156
         Width           =   888
      End
      Begin VB.CommandButton cmdZoomMais 
         Caption         =   "Zoom +"
         Height          =   600
         Left            =   288
         Picture         =   "Ilegiveis.frx":01DE
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   852
         Width           =   888
      End
      Begin VB.CommandButton cmdZoomMenos 
         Caption         =   "Zoom -"
         Height          =   600
         Left            =   288
         Picture         =   "Ilegiveis.frx":04E8
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1476
         Width           =   888
      End
      Begin VB.CommandButton cmdRotacao 
         Caption         =   "Rotação"
         Height          =   600
         Left            =   288
         Picture         =   "Ilegiveis.frx":07F2
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2112
         Width           =   888
      End
      Begin VB.CommandButton cmdInverteCor 
         Caption         =   "Inverte cor"
         Height          =   600
         Left            =   288
         Picture         =   "Ilegiveis.frx":0AFC
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2736
         Width           =   888
      End
      Begin VB.CommandButton cmdFrenteVerso 
         Caption         =   "Fren/Ver"
         Height          =   600
         Left            =   288
         Picture         =   "Ilegiveis.frx":0E06
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   3372
         Width           =   888
      End
   End
   Begin VB.ListBox lstCapa 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2700
      ItemData        =   "Ilegiveis.frx":1110
      Left            =   120
      List            =   "Ilegiveis.frx":1112
      TabIndex        =   34
      Top             =   672
      Width           =   3156
   End
   Begin VB.Frame FrmImagem 
      Caption         =   "Imagem"
      Height          =   4500
      Left            =   120
      TabIndex        =   18
      Top             =   4032
      Width           =   10272
      Begin LeadLib.Lead Lead1 
         Height          =   4212
         Left            =   96
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   216
         Width           =   10068
         _Version        =   524288
         _ExtentX        =   17759
         _ExtentY        =   7429
         _StockProps     =   229
         BackColor       =   16777215
         BorderStyle     =   1
         ScaleHeight     =   349
         ScaleWidth      =   837
         DataField       =   ""
         BitmapDataPath  =   ""
         AnnDataPath     =   ""
      End
   End
   Begin VB.Timer TmrPesquisa 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9456
      Top             =   24
   End
   Begin MSFlexGridLib.MSFlexGrid GrdDocto 
      Height          =   2700
      Left            =   3264
      TabIndex        =   86
      Top             =   672
      Width           =   7152
      _ExtentX        =   12615
      _ExtentY        =   4763
      _Version        =   393216
      Rows            =   10
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      BackColorFixed  =   -2147483634
      BackColorBkg    =   -2147483634
      FocusRect       =   0
      GridLines       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblMotivoIlegivel 
      AutoSize        =   -1  'True
      Caption         =   "Ilegibilidade:"
      Height          =   192
      Left            =   2304
      TabIndex        =   88
      Top             =   3456
      Width           =   924
   End
   Begin VB.Label lblMotivoDoctoIlegivel 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ilegibilidade do documento:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   276
      Left            =   3312
      TabIndex        =   87
      Top             =   3408
      Width           =   7056
   End
   Begin VB.Label LblCapaDup 
      Caption         =   "Capa Duplicada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   288
      Left            =   6840
      TabIndex        =   70
      Top             =   60
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label lblOcorrencia 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ocorrência:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   276
      Left            =   120
      TabIndex        =   37
      Top             =   3792
      Width           =   10272
   End
   Begin VB.Label lblNumMalote 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   5232
      TabIndex        =   36
      Top             =   48
      Width           =   1500
   End
   Begin VB.Label lblLote 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   2616
      TabIndex        =   35
      Top             =   48
      Width           =   1224
   End
End
Attribute VB_Name = "Ilegiveis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Delaração dos Objetos RDO
Private qryGetCapa                      As rdoQuery
Private qryGetCapaDuplicada             As rdoQuery
Private qryGetDocumentos                As rdoQuery
Private qryAtualizaStatusCapa           As rdoQuery
Private qryGetOcorr                     As rdoQuery
Private qryGetMotivoIlegiveis           As rdoQuery
Private qryAtualizaOcorrencia           As rdoQuery
Private qryGetMotivoExclusao            As rdoQuery
Private qryInsereMotivoExclusao         As rdoQuery
Private qryRemoveMotivoExclusao         As rdoQuery
Private qryAtualizaValorDocumento       As rdoQuery
Private qryRemoveAjusteCapa             As rdoQuery
Private qryAtualizaOcorrenciaCapa       As rdoQuery
Private qryVA_GetDocumentosTransmitidos As rdoQuery
Private qryAtualizaCapa                 As rdoQuery
Private qryGetDocumentosCapa            As rdoQuery
Private qryRemoveDocumento              As rdoQuery
Private qryAtualizaStatusDocumentosCapa As rdoQuery
Private qryRemoveCapaRecepcionada       As rdoQuery

'Declaração de Variáveis
Private AlterouDocto                    As Boolean
Private PrimeiraVez                     As Boolean
Private bCapaDuplicada                  As Boolean
Private teclou                          As Boolean
Private IdSelecionado                   As Long
Private sTempo                          As Integer
Public sCapaOuDocumento                 As String

'Declaração dos Arrays
Private aDoc()                          As TDoc
Private aCapa()                         As TCapa

'Type de Capas
Private Type TCapa
  IdCapa            As Long
  IdLote            As Long
  IdEnv_Mal         As String * 1
  Capa              As String * 18
  NumMalote         As String * 11
  AgOrig            As Integer
  Status            As String * 1
  AlterouDocto      As Boolean
  Duplicidade       As Integer
  Comentario        As String * 60
  IdModuloAnterior  As Integer
End Type

'Type para Documentos
Private Type TDoc
  NrSeq             As Integer
  IdDocto           As Long
  IdCapa            As Long
  TipoDocto         As Integer
  CodMotivo         As Long             'Código do Motivo de Envio para Ilegíveis
  DscTipoDocto      As String * 18
  Duplicidade       As Boolean
  Ocorrencia        As String * 5
  RetornoTransacao  As Long
  Leitura           As String * 48
  Frente            As String * 20
  Verso             As String * 20
  Status            As String * 1
  Vinculo           As Long
  Valor             As Double
  Ordem             As String
End Type

Private Function RemoveCapaRecepcionada()
'Verifica se existe a mesma capa com status = 0 para capa não definida
'pela Vips, se Sim excluir a capa recepcionada somente
    
On Error GoTo TrataErro
    Dim sPosicaoErro As String
    sPosicaoErro = "RemoveCapa"
    
    'Remove Capa apenas Recepcionada (Status = 0)
    Set qryRemoveCapaRecepcionada = Geral.Banco.CreateQuery("", "{? = call RemoveCapaRecepcionada (?,?,?,?)}")
    'Parâmetros (1)-Data (2)-Capa (3)-AgOrig (4)-Num_Malote
    qryRemoveCapaRecepcionada.rdoParameters(0).Direction = rdParamReturnValue
    
    With qryRemoveCapaRecepcionada
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = Geral.Capa.Capa
        .rdoParameters(3) = Geral.Capa.AgOrig
        .rdoParameters(4) = Geral.Capa.Num_Malote
        .Execute
        If .rdoParameters(0).Value <> 0 Then
            Beep
            'Geral.Banco.RollbackTrans
            If MsgBox("Não foi possível atualizar documento referente a Capa. Continua", vbCritical + vbYesNo, App.Title) = vbYes Then
                RemoveCapaRecepcionada
            Else
                Exit Function
            End If
        End If
    End With

Exit Function

TrataErro:
    Select Case TratamentoErro("Não foi possível validar o documento atual.(" & sPosicaoErro & ")", Err, rdoErrors)
           Case vbRetry
                MsgBox "Não é possível repetir a operação!", vbInformation + vbOKOnly, "Atenção"
    End Select
    Me.Hide

End Function
Function AtualizaOcorrenciaCapa(ByVal nIdCapa As Long, ByVal nOcorrencia As Integer) As Boolean

  On Error GoTo ERRO_ATUALIZAOCORRENCIACAPA

  AtualizaOcorrenciaCapa = False

  Set qryAtualizaOcorrenciaCapa = Geral.Banco.CreateQuery("", "{? = call AtualizaOcorrenciaCapa (?,?,?)}")
  With qryAtualizaOcorrenciaCapa
      .rdoParameters(0).Direction = rdParamReturnValue
      .rdoParameters(1) = Geral.DataProcessamento     'Data Proc.
      .rdoParameters(2) = nIdCapa                     'IdCapa
      .rdoParameters(3) = nOcorrencia                 'Ocorrencia
      .Execute
  End With

  If qryAtualizaOcorrenciaCapa(0).Value = 1 Then
    MsgBox "Ocorreu um erro ao atualizar a ocorrência da Capa.", vbInformation + vbOKOnly, App.Title
    Exit Function
  End If

  'Gravar Log
  Call GravaLog(nIdCapa, 0, 3)

  AtualizaOcorrenciaCapa = True

  Exit Function

ERRO_ATUALIZAOCORRENCIACAPA:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Atualizar o Status do Documento.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function
Sub AtualizaStatusCapa(ByVal sIdCapa As Long, sStatus As String)

  On Error GoTo ERRO_ATUALIZASTATUS

  'Se IdCapa passado = 0 , sair da funcao
  If Val(sIdCapa) = 0 Then Exit Sub

  Set qryAtualizaStatusCapa = Geral.Banco.CreateQuery("", "{? = call AtualizaStatusCapa (?,?,?)}")
  With qryAtualizaStatusCapa
      .rdoParameters(0).Direction = rdParamReturnValue
      .rdoParameters(1) = Geral.DataProcessamento 'Data Proc.
      .rdoParameters(2) = sIdCapa                 'IdCapa
      .rdoParameters(3) = sStatus                 'Status
      .Execute
  End With

  If qryAtualizaStatusCapa(0).Value = 1 Then
    MsgBox "Ocorreu um erro ao atualizar o status da capa.", vbInformation + vbOKOnly, App.Title
    Exit Sub
  End If

  'Gravação de Log
  Select Case sStatus
  Case "1"
    'Reenvia para Complementação
    Call GravaLog(sIdCapa, 0, 1)
  Case "D"
    'Devolver Capa
    Call GravaLog(sIdCapa, 0, 3)
  Case "8"
    'Envia para Vinculo Automatico
    Call GravaLog(sIdCapa, 0, 4)
  End Select

  Exit Sub

ERRO_ATUALIZASTATUS:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Atualizar o Status do Documento.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Sub
Function CapaDuplicada() As Boolean

  On Error GoTo ERRO_CAPADUPLICADA

  Dim sSql As String
  Dim rsCapa As rdoResultset

  CapaDuplicada = False

  If aCapa(lstCapa.ListIndex + 1).IdEnv_Mal = "E" And aCapa(lstCapa.ListIndex + 1).AgOrig = 0 Then Exit Function

  If aCapa(lstCapa.ListIndex + 1).Duplicidade = 1 Then
    CapaDuplicada = True
  Else
    CapaDuplicada = False
  End If

  Exit Function

ERRO_CAPADUPLICADA:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Verificar Duplicidade de Capas.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function
Private Sub ChamaTelaComplementacao(ByRef sForm As Form)

   On Error GoTo ERRO_CHAMATELA

   'Elimina brancos no campo Leitura
   Geral.Documento.Leitura = Trim(Geral.Documento.Leitura)

   Load sForm

   sForm.SetParent Me
   sForm.SetPosition (Me.Left + (Me.Width - sForm.Width) / 2), Me.Top
   sForm.Show vbModal, Me

   If sForm.Alterou = True Then
     
      'Verificar se a tela chamada é Envelope ou Malote
      If sForm.Name = "Envelope" Or sForm.Name = "Malote" Then
         'Atualizar Status do documento
         Call AtualizaStatusDocumento(aDoc(grdDocto.Row + 1).IdDocto, "1")

         If sForm.Name = "Envelope" Then
            aDoc(grdDocto.Row + 1).DscTipoDocto = "Envelope"
            aCapa(lstCapa.ListIndex + 1).IdEnv_Mal = "E"
            aCapa(lstCapa.ListIndex + 1).NumMalote = ""
            LblEnv_Mal.Caption = "Envelope"
            Call HDMalote(False)
         Else
            aDoc(grdDocto.Row + 1).DscTipoDocto = "Malote"
            aCapa(lstCapa.ListIndex + 1).IdEnv_Mal = "M"
            aCapa(lstCapa.ListIndex + 1).NumMalote = Geral.Capa.Num_Malote

            'Informar o Numero do Malote na tela
            lblNumMalote.Caption = aCapa(lstCapa.ListIndex + 1).NumMalote

            LblEnv_Mal.Caption = "Malote"
            Call HDMalote(True)
         End If

         'Verificar se a capa está duplicada
         If Geral.Capa.Duplicidade = "1" Then
            'Capa Duplicada -> Não permitir a edição dos documentos
            LblCapaDup.Visible = True
            aCapa(lstCapa.ListIndex + 1).Duplicidade = 1
            bCapaDuplicada = True
         Else
            LblCapaDup.Visible = False
            bCapaDuplicada = False
            aCapa(lstCapa.ListIndex + 1).Duplicidade = 0
         End If

         'Informar o Lote na tela
         lblLote.Caption = Format(Trim(aCapa(lstCapa.ListIndex + 1).IdLote), "0000")

         aDoc(grdDocto.Row + 1).Status = "1"
         aCapa(lstCapa.ListIndex + 1).AgOrig = Geral.Capa.AgOrig
         aCapa(lstCapa.ListIndex + 1).Capa = Geral.Capa.Capa

         lstCapa.List(lstCapa.ListIndex) = Geral.Capa.Capa

'      Else
'         'Nao eh uma capa
'         If Geral.Documento.Status <> "D" And Geral.Documento.Status <> "L" Then
'            'Atualizar Status do documento
'            Call AtualizaStatusDocumento(aDoc(GrdDocto.Row + 1).IdDocto, "1")
'         End If
      End If

      'Refazer a Lista de Documentos
      Call PreencheListDocto(grdDocto.Row)

      'Verificar se documento foi devolvido para gravar log
      If Geral.Documento.Status = "D" Then
         'Documento Devolvido -> Gravar Log
         Call GravaLog(aCapa(lstCapa.ListIndex + 1).IdCapa, Geral.Documento.IdDocto, 2)
      Else
         If Geral.Documento.Status <> "L" Then
            'Documento Recomplementado -> Atualizar o Status para 1 (Complementado)
            Call AtualizaStatusDocumento(aDoc(grdDocto.Row + 1).IdDocto, "1")
         End If
         'Documento Recomplementado -> Gravar Log
         Call GravaLog(aCapa(lstCapa.ListIndex + 1).IdCapa, aDoc(grdDocto.Row + 1).IdDocto, 5)
      End If
   End If

   Unload sForm

   grdDocto.SetFocus
   Exit Sub

ERRO_CHAMATELA:
   Screen.MousePointer = vbDefault
   Select Case TratamentoErro("Erro ao executar tela para Complementação de Documentos.", Err, rdoErrors)
      Case vbCancel
      Case vbRetry
   End Select
End Sub
Sub ChamaTelaComplementacaoCapa(ByVal TipoCapa As Form)

   'Preencher Type de Documentos
   Geral.Documento.IdDocto = aDoc(grdDocto.Row + 1).IdDocto
   Geral.Documento.Leitura = ""
   Geral.Documento.TipoDocto = 1
   Geral.Documento.Status = aDoc(grdDocto.Row + 1).Status
   Geral.Documento.Agencia = aCapa(lstCapa.ListIndex + 1).AgOrig
   Geral.Documento.IdCapa = aCapa(lstCapa.ListIndex + 1).IdCapa

   'Preencher Type de Capa
   Geral.Capa.AgOrig = 0
   Geral.Capa.IdEnv_Mal = aCapa(lstCapa.ListIndex + 1).IdEnv_Mal
   Geral.Capa.IdCapa = aCapa(lstCapa.ListIndex + 1).IdCapa
   Geral.Capa.Status = aCapa(lstCapa.ListIndex + 1).Status
   Geral.Capa.Capa = 0
   Geral.Capa.Num_Malote = 0

   TabTipoDoc.Visible = False
   PicTiposDoc.Visible = False
   Call HDObjetos(True)

   'Zerar os campos capa , num_malote e agencia da capa
   Set qryAtualizaCapa = Geral.Banco.CreateQuery("", "{? = call AtualizaCapa (?,?,?,?,?,?,?)}")
   With qryAtualizaCapa
      .rdoParameters(0).Direction = rdParamReturnValue
      .rdoParameters(1) = Geral.DataProcessamento                'Data Proc.
      .rdoParameters(2) = Geral.Capa.IdCapa                      'IdCapa
      .rdoParameters(3) = 0                                      'Capa
      .rdoParameters(4) = 0                                      'Ag. Origem
      .rdoParameters(5) = Geral.Documento.IdDocto                'IdDocto
      .rdoParameters(6) = Val(aCapa(lstCapa.ListIndex + 1).Capa)  'Num. Malote
      .rdoParameters(7) = ""                                     'CMC7

      '.Execute
   End With

   If qryAtualizaCapa(0).Value = 1 Then
      MsgBox "Ocorreu um erro ao atualizar dados da capa.", vbInformation + vbOKOnly, App.Title
      Exit Sub
   End If

   Call ChamaTelaComplementacao(TipoCapa)
End Sub
Sub ExcluiCapaDuplicada()

  Dim sSql As String
  Dim RsMotExc As rdoResultset

  On Error GoTo ERRO_EXCLUICAPADUPLICADA

  'Atualizar Status da Capa para 'D'
  Call AtualizaStatusCapa(aCapa(lstCapa.ListIndex + 1).IdCapa, "D")
  aCapa(lstCapa.ListIndex + 1).Status = "D"

  'Atualizar Ocorrencia da Capa para '998'
  If Not AtualizaOcorrenciaCapa(aCapa(lstCapa.ListIndex + 1).IdCapa, 998) Then Exit Sub

  'Verificar se esta Capa já foi excluída
  sSql = Geral.DataProcessamento & " , " & aCapa(lstCapa.ListIndex + 1).IdCapa

  Set qryGetMotivoExclusao = Geral.Banco.CreateQuery("", "{call GetMotivoExclusao (" & sSql & ")}")

  Set RsMotExc = qryGetMotivoExclusao.OpenResultset(rdOpenStatic, rdConcurReadOnly)

  If Not RsMotExc.EOF Then
    'Já Existe - Excluir Motivo Antigo
    Set qryRemoveMotivoExclusao = Geral.Banco.CreateQuery("", "{? = call RemoveMotivoExclusao (?,?)}")
    With qryRemoveMotivoExclusao
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1) = Geral.DataProcessamento             'Data Proc.
        .rdoParameters(2) = aCapa(lstCapa.ListIndex + 1).IdCapa 'IdCapa
        .Execute
    End With

    If qryRemoveMotivoExclusao(0).Value = 1 Then
      MsgBox "Ocorreu um erro ao atualizar o motivo de exclusão da capa.", vbInformation + vbOKOnly, App.Title
      Exit Sub
    End If
  End If

  'Gravar Motivo de Exclusão da Capa
  Set qryInsereMotivoExclusao = Geral.Banco.CreateQuery("", "{? = call InsereMotivoExclusao (?,?,?)}")
  With qryInsereMotivoExclusao
      .rdoParameters(0).Direction = rdParamReturnValue
      .rdoParameters(1) = Geral.DataProcessamento               'Data Proc.
      .rdoParameters(2) = aCapa(lstCapa.ListIndex + 1).IdCapa   'IdCapa
      .rdoParameters(3) = "Capa Duplicada."                     'MotivoExclusao
      .Execute
  End With

  If qryInsereMotivoExclusao(0).Value = 1 Then
    MsgBox "Ocorreu um erro ao gravar motivo de exclusão da capa.", vbInformation + vbOKOnly, App.Title
    Exit Sub
  End If

  'Atualizar status e a ocorrencia de todos os documentos da capa
  Set qryAtualizaStatusDocumentosCapa = Geral.Banco.CreateQuery("", "{? = call AtualizaStatusDocumentosCapa (?,?,?,?)}")
  With qryAtualizaStatusDocumentosCapa
      .rdoParameters(0).Direction = rdParamReturnValue              'Parametro de Output
      .rdoParameters(1) = Geral.DataProcessamento                   'Data Proc.
      .rdoParameters(2) = aCapa(lstCapa.ListIndex + 1).IdCapa       'IdCapa
      .rdoParameters(3) = "D"                                       'Status
      .rdoParameters(4) = 999                                       'Ocorrencia
      .Execute
  End With

  If qryAtualizaStatusDocumentosCapa(0).Value = 1 Then
    MsgBox "Ocorreu um erro ao atualizar status dos documentos.", vbInformation + vbOKOnly, App.Title
    Exit Sub
  End If

  Exit Sub

ERRO_EXCLUICAPADUPLICADA:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Excluir Capa Duplicada.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Sub
Private Sub FinalizaCapa(ByVal sCod As String)

    Screen.MousePointer = vbHourglass

    'Atualizar o STATUS da capa para '8' -> Vínculo Automático
    Call AtualizaStatusCapa(aCapa(lstCapa.ListIndex + 1).IdCapa, sCod)
    aCapa(lstCapa.ListIndex + 1).Status = sCod

    'Limpando a variável que armazena a capa Atual
    IdSelecionado = 0
    
    If UCase(sCod) = "A" Then
        GravaLog aCapa(lstCapa.ListIndex + 1).IdCapa, 0, 8
    End If

    Screen.MousePointer = vbDefault

    'Posicionar na próxima Capa da Lista
    Call LimpaListaDocto

    If lstCapa.ListIndex + 1 < lstCapa.ListCount Then
      'Existem mais Capas -> Posicionar
      lstCapa.ListIndex = lstCapa.ListIndex + 1
    Else
      Call CmdAtualizar_Click
    End If
End Sub
Sub HabilitaTimerPesquisa()

  'Esta Função irá verificar a existência de documentos Ilegíveis a cada x segundos
  'de acordo com o campo PARAMETRO.TmAtualizacao
  FrmPesquisa.Visible = True
  tmrPesquisa.Enabled = True
  Progress.Value = 0
  
  ''''''''''''''''''''''''''''''''''''''''''
  'Grava log MDI - Inicio Aguarda documento'
  ''''''''''''''''''''''''''''''''''''''''''
  Call GravaLog(0, 0, 252)
  
End Sub
Sub HDMalote(ByVal bValor As Boolean)

  PctMalote.Visible = bValor
  lblNumMalote.Visible = bValor
  If bValor = False Then
    LblCapaDup.Visible = False
    lblLote.Caption = ""
  End If
End Sub
Sub HDObjetos(bValor As Boolean)

  On Error GoTo ERRO_HDOBJETOS

  'Habilita / Desabilita Objetos e frames para que o TAB 'TipoDocto' fique Modal
  FrmCmd.Enabled = bValor
  FrmCmdImagem.Enabled = bValor
  frmImagem.Enabled = bValor

  lstCapa.Enabled = bValor
  grdDocto.Enabled = bValor
  
  chkFiltro.Enabled = bValor

  Exit Sub

ERRO_HDOBJETOS:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Desabilitar Objetos.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Sub
Sub HDObjetosImagem(bValor As Boolean)

  On Error GoTo ERRO_HDOBJETOS

  cmdAuditoria.Enabled = bValor
  cmdZoomMais.Enabled = bValor
  cmdZoomMenos.Enabled = bValor
  cmdRotacao.Enabled = bValor
  cmdInverteCor.Enabled = bValor
  cmdFrenteVerso.Enabled = bValor
  frmImagem.Visible = bValor
  Lead1.ForceRepaint

  Exit Sub

ERRO_HDOBJETOS:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao preparar botões de manipulação de Imagens.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Sub

Sub LimpaListaDocto()

    grdDocto.Clear
    grdDocto.Rows = 0
    
    lblOcorrencia.Caption = ""

End Sub

Sub LimpaListas()

    lstCapa.Clear
    grdDocto.Clear
    grdDocto.Rows = 0
    
    frmImagem.Visible = False

    Erase aCapa
    
End Sub
Private Sub MarcaDoctoRecaptura()

    Dim X           As Integer
    Dim iLnInicial  As Integer
    Dim iLnFinal    As Integer
    
    If grdDocto.RowSel > grdDocto.Row Then
        iLnInicial = grdDocto.Row
        iLnFinal = grdDocto.RowSel
    Else
        iLnInicial = grdDocto.RowSel
        iLnFinal = grdDocto.Row
    End If

    For X = iLnInicial To iLnFinal
        'Atualizar o Status do Documento para 'A' (Recaptura)
        Call AtualizaStatusDocumento(aDoc(X + 1).IdDocto, "A")
        aDoc(X + 1).Status = "A"
        DoEvents
    Next X
    
    'Refazer a lista de documentos
    Call PreencheListDocto(grdDocto.Row)
    
End Sub
Private Function PossuiDoctoRecaptura() As Boolean

  Dim X As Integer

  PossuiDoctoRecaptura = False

  'Verificar se existe algum documento para recaptura
  If grdDocto.Rows > 0 Then
    For X = 0 To grdDocto.Rows - 1
      If aDoc(X + 1).Status = "A" Then
        'Documento para Recaptura
        PossuiDoctoRecaptura = True
        Exit Function
      End If
      DoEvents
    Next X
  End If
  
End Function
Function PreencheListCapas(Optional pIdModulo As Integer, Optional pIdTitulo As Integer) As Boolean

  On Error GoTo ERRO_PREENCHELISTCAPAS

  Dim rsCapa        As rdoResultset
  Dim sSql          As String
  Dim sPosicaoErro  As String
  Dim X             As Integer
  Dim sLinha        As String
  Dim sDescricao    As String

  Call LimpaListas

  'Passando parâmetros para a Stored Procedure 'GetCapaIlegiveis'
  sSql = Geral.DataProcessamento & " , " & Geral.Intervalo

  'Set qryGetCapa = Geral.Banco.CreateQuery("", "{call GetCapaIlegiveis (" & sSql & ")}")
  Set qryGetCapa = Geral.Banco.CreateQuery("", "{call GetCapaIlegiveis (?,?,?,?)}")
  
  qryGetCapa.rdoParameters(0).Value = Geral.DataProcessamento
  qryGetCapa.rdoParameters(1).Value = Geral.Intervalo
  If pIdModulo <> 0 Then
      qryGetCapa.rdoParameters(2).Value = pIdModulo
  End If
  If pIdTitulo <> 0 Then
      qryGetCapa.rdoParameters(3).Value = pIdTitulo
  End If
  

  Set rsCapa = qryGetCapa.OpenResultset(rdOpenStatic, rdConcurReadOnly)

  If rsCapa.RowCount > 0 Then

    'Desabilitar o Timer de Pesquisa
    tmrPesquisa.Enabled = False
    FrmPesquisa.Visible = False

    'ReDim aCapa(0)
    ReDim Preserve aCapa(rsCapa.RowCount)

    X = 1
    While Not rsCapa.EOF
        'Carregando o Array com as Capas
        'ReDim Preserve aCapa(UBound(aCapa) + 1)
        aCapa(X).IdCapa = rsCapa!IdCapa
        aCapa(X).IdLote = rsCapa!IdLote
        aCapa(X).IdEnv_Mal = rsCapa!IdEnv_Mal
        aCapa(X).Capa = rsCapa!Capa
        aCapa(X).NumMalote = rsCapa!Num_Malote
        aCapa(X).AgOrig = rsCapa!AgOrig
        aCapa(X).Status = rsCapa!Status
        aCapa(X).Duplicidade = rsCapa!Duplicidade
        aCapa(X).Comentario = IIf(IsNull(rsCapa!Comentario), "", rsCapa!Comentario)
        aCapa(X).IdModuloAnterior = IIf(IsNull(rsCapa!IdModuloAnterior), 0, rsCapa!IdModuloAnterior)
        
        
        sDescricao = IIf(IsNull(rsCapa!Descricao), "", rsCapa!Descricao)
        sLinha = rsCapa!Capa & String(28 - (Len(sDescricao) + Len(rsCapa!Capa)), " ") & sDescricao

        lstCapa.AddItem sLinha
        lstCapa.ItemData(lstCapa.NewIndex) = rsCapa!IdCapa

        rsCapa.MoveNext
        X = X + 1
        DoEvents
    Wend
  End If

  'Selecionar a Primeira Capa , caso exista
  If lstCapa.ListCount > 0 Then
    lstCapa.Selected(0) = True
    PreencheListCapas = True
  Else
    PreencheListCapas = False
    Call HDObjetosImagem(False)
    Call HDMalote(False)
  End If

  Exit Function

ERRO_PREENCHELISTCAPAS:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Ler Capas.", Err, rdoErrors)
    Case vbCancel, vbRetry
      Unload Me
  End Select
End Function
Function CapaSelecionadaDisponivel() As Integer

  On Error GoTo ERRO_CAPASELECIONADADISP

  Dim sSql As String

  sSql = Geral.DataProcessamento & " , " & aCapa(lstCapa.ListIndex + 1).IdCapa & _
         ",'5','H'," & Geral.Intervalo

  Set qryGetCapa = Geral.Banco.CreateQuery("", "{? = call VerificaCapaDisponivel (?,?,?,?,?)}")

  With qryGetCapa
    .rdoParameters(0).Direction = rdParamReturnValue
    .rdoParameters(1) = Geral.DataProcessamento             'Data de Processamento
    .rdoParameters(2) = aCapa(lstCapa.ListIndex + 1).IdCapa 'IdCapa
    .rdoParameters(3) = "5"                                 'Status 1
    .rdoParameters(4) = "H"                                 'Status 2 (Pendentes)
    .rdoParameters(5) = Geral.Intervalo                     'Intervalo de Atualização

    .Execute
  End With

  CapaSelecionadaDisponivel = qryGetCapa(0)

  If qryGetCapa(0) = 1 Then
    lstCapa.ListIndex = -1
    grdDocto.Clear
    grdDocto.Rows = 0
    
    frmImagem.Visible = False
    Screen.MousePointer = vbDefault
    MsgBox "Este Envelope / Malote não está disponível. Pode estar sendo tratado por outra estação ou já foi tratado.", vbInformation, App.Title
  End If

  Exit Function

ERRO_CAPASELECIONADADISP:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Verificar se a Capa está Disponível.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function
Sub PreencheListDocto(ByVal Indice As Integer)

  On Error GoTo ERRO_PREENCHELISTDOCTO

  Dim sSql As String
  Dim sLinha As String
  Dim rsDocumentos As rdoResultset
  Dim X As Integer

  grdDocto.Visible = False

  'Selecionar todos os documentos pertencentes à capa selecionada
  sSql = Geral.DataProcessamento & " , " & Val(lstCapa.ItemData(lstCapa.ListIndex))

  Set qryGetDocumentos = Geral.Banco.CreateQuery("", "{call GetDocumentoIlegiveis (" & sSql & ")}")

  Set rsDocumentos = qryGetDocumentos.OpenResultset(rdOpenStatic, rdConcurReadOnly)

  X = 1
  Call LimpaListaDocto
  ReDim aDoc(rsDocumentos.RowCount)

  If Not rsDocumentos.EOF Then
    While Not rsDocumentos.EOF

        'Numero Sequencial
        aDoc(X).NrSeq = X
        sLinha = Format(aDoc(X).NrSeq, "0000") & Space(2)

        'Vinculo
        aDoc(X).Vinculo = Val(rsDocumentos!Vinculo & "")
        'sLinha = sLinha & Format(aDoc(X).Vinculo, String(5, "0")) & Space(9 - Len(Format(aDoc(X).Vinculo, String(5, "0"))))
        sLinha = sLinha & Format(aDoc(X).Vinculo, String(5, "0")) & Space(7 - Len(Format(aDoc(X).Vinculo, String(5, "0"))))

        'Ocorrencia
        aDoc(X).Ocorrencia = Val(rsDocumentos!Ocorrencia & "")
        If Val(rsDocumentos!Ocorrencia & "") <> 0 Then
          'sLinha = sLinha & "S" & Space(6)
          sLinha = sLinha & "S" & Space(5)
        Else
          'sLinha = sLinha & Space(7)
          sLinha = sLinha & Space(6)
        End If
        
        'Retorno Transacao
        aDoc(X).RetornoTransacao = Val(rsDocumentos!RetornoTransacao)

        'Indicador de Duplicidade
        aDoc(X).Duplicidade = Val(rsDocumentos!Duplicidade & "")
        If Val(rsDocumentos!Duplicidade) = 1 Then
          sLinha = sLinha & "S" & Space(5)
        Else
          sLinha = sLinha & " " & Space(5)
        End If

        'Status do Documento
        aDoc(X).Status = rsDocumentos!Status & ""

        'Tipo de Documento
        aDoc(X).TipoDocto = rsDocumentos!TipoDocto & ""

        If aDoc(X).Status = "A" Then
            sLinha = sLinha & "S" & Space(4)
        Else
            sLinha = sLinha & Space(5)
        End If

        Select Case aDoc(X).TipoDocto
          Case 0          'Indefinido
            aDoc(X).DscTipoDocto = "INDEFINIDO   "

          Case 1          'CAPA DE ENVELOPE / MALOTE
            If aCapa(lstCapa.ListIndex + 1).IdEnv_Mal = "E" Then
              If Val(aDoc(X).Status) = 0 Then
                aDoc(X).DscTipoDocto = "ENVELOPE ILEGÍVEL "
              Else
                aDoc(X).DscTipoDocto = "ENVELOPE          "
              End If
            Else
              If Val(aDoc(X).Status) = 0 Then
                aDoc(X).DscTipoDocto = "MALOTE ILEGÍVEL   "
              Else
                aDoc(X).DscTipoDocto = "MALOTE            "
              End If
            End If

          Case 2, 3       'Depósito
            aDoc(X).DscTipoDocto = "DEPOSITO          "

          Case 4          'Aut. Déb.
            aDoc(X).DscTipoDocto = "DEBITO CC         "
          Case 5, 6, 7    'Cheque
            aDoc(X).DscTipoDocto = "CHEQUE            "
          Case 32, 34     'Ajuste de Crédito
            aDoc(X).DscTipoDocto = "AJ. CREDITO       "
          Case 33, 38     'Ajuste de Débito
            aDoc(X).DscTipoDocto = "AJ. DÉBITO        "
          Case 36         'Cartão Avulso
            aDoc(X).DscTipoDocto = "CARTÃO AVULSO     "
          Case 37         'OCT
            aDoc(X).DscTipoDocto = "OCT               "
          Case 39         'Capa OCT
            aDoc(X).DscTipoDocto = "CAPA OCT          "
          Case 41         'LANÇAMENTO INTERNO
            aDoc(X).DscTipoDocto = "LANCTO INTERNO    "
          Case Else       'Pagamento
            aDoc(X).DscTipoDocto = "PAGAMENTO         "
        End Select

        sLinha = sLinha & aDoc(X).DscTipoDocto '& Space(3)

        'Valor do Documento
        aDoc(X).Valor = FormataValor(rsDocumentos!Valor & "", 15)
        sLinha = sLinha & FormataValor(rsDocumentos!Valor, 15)

        'Frente e Verso
        aDoc(X).Frente = Trim(rsDocumentos!Frente & "")
        aDoc(X).Verso = Trim(rsDocumentos!Verso & "")

        'IdDocto
        aDoc(X).IdDocto = rsDocumentos!IdDocto & ""

        'Leitura
        aDoc(X).Leitura = rsDocumentos!Leitura & ""
        
        'Ordem
        aDoc(X).Ordem = rsDocumentos!Ordem & ""

        'Código Motivo de Envio para Ilegíveis
        aDoc(X).CodMotivo = "0" & rsDocumentos!CodMotivo
        
        grdDocto.Rows = grdDocto.Rows + 1
        grdDocto.Row = grdDocto.Rows - 1
        grdDocto.TextMatrix(grdDocto.Row, 0) = sLinha
        grdDocto.TextMatrix(grdDocto.Row, 1) = rsDocumentos!IdDocto
        'Mudar cor da linha para documento indefinido ou Capa Ilegível
        If (aDoc(X).TipoDocto = 0) Or (aDoc(X).TipoDocto = 1 And Val(aDoc(X).Status) = 0) Then
            'Apresenta em vermelho somente documento sem ocorrência
            If aDoc(X).Status <> "D" Then grdDocto.CellForeColor = vbRed
        End If

        rsDocumentos.MoveNext
        X = X + 1
        DoEvents
    Wend
  Else
    Call HDObjetosImagem(False)
  End If

  If lstCapa.ListCount > 0 And grdDocto.Rows > 0 Then
    grdDocto.Row = Indice
    GrdDocto_Click
    IdSelecionado = aCapa(lstCapa.ListIndex + 1).IdCapa
  End If
  
  grdDocto.Visible = True
  grdDocto.SetFocus
  
  'Posiciona na linha atual
  If Indice > 0 Then SendKeys "{LEFT}", True
    
  DoEvents

  Exit Sub

ERRO_PREENCHELISTDOCTO:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Preencher Lista de Documentos.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Sub
Function VerificaDoctosIndefinidos() As Boolean

  Dim X As Integer

  VerificaDoctosIndefinidos = False

  'Verificar se existe algum documento indefinido
  If grdDocto.Rows > 0 Then
    For X = 0 To grdDocto.Rows - 1
      If aDoc(X + 1).TipoDocto = 0 And Val(aDoc(X + 1).Ocorrencia) = 0 Then
        'Documento Indefinido
        VerificaDoctosIndefinidos = True
        Exit Function
      End If
      DoEvents
    Next X
  End If
End Function

Function VerificaDocumentosTransmitidos() As Boolean

   On Error GoTo VerificaDocumentosTransmitidos_Err

   Dim RsDoctosTrans As rdoResultset

   VerificaDocumentosTransmitidos = False

'   Set qryVA_GetDocumentosTransmitidos = Geral.Banco.CreateQuery("", "{ ? = Call VA_GetDocumentosTransmitidos (?,?)}")
   Set qryVA_GetDocumentosTransmitidos = Geral.Banco.CreateQuery("", "{ ? = Call GetDocumentosParaVerificacao (?,?)}")
   

   With qryVA_GetDocumentosTransmitidos
      .rdoParameters(1) = Geral.DataProcessamento
      .rdoParameters(2) = aCapa(lstCapa.ListIndex + 1).IdCapa
   End With

   Set RsDoctosTrans = qryVA_GetDocumentosTransmitidos.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
   If Not RsDoctosTrans.EOF Then
      If RsDoctosTrans!Qtde > 0 Then
         VerificaDocumentosTransmitidos = True
         'Atualizar o Status da Capa para 'V' - Em Analise
         Call AtualizaStatusCapa(aCapa(lstCapa.ListIndex + 1).IdCapa, "V")
         aCapa(lstCapa.ListIndex + 1).Status = "V"

         'Gravar Log
         Call GravaLog(aCapa(lstCapa.ListIndex + 1).IdCapa, 0, 190)
         Call LimpaListaDocto
         MsgBox "Este Envelope/Malote não está mais disponível por já ter sido tratado ou porque esta sendo tratado por outra estação.", vbInformation + vbOKOnly, App.Title
      End If
   End If

   Exit Function

VerificaDocumentosTransmitidos_Err:
   Screen.MousePointer = vbDefault
   Select Case TratamentoErro("Erro ao Verificar Documentos já Transmitidos.", Err, rdoErrors)
      Case vbCancel
      Case vbRetry
   End Select
End Function

Private Sub chkFiltro_Click()

    Dim IdItem      As Integer
    Dim Opcao       As Integer

    If chkFiltro.Value = vbChecked Then
        Opcao = Filtro.ShowModal(IdItem)
        If Opcao Then
            chkFiltro.Value = vbChecked

            ''''''''''''''''''''''''''''''''''''''''''
            'Retira capa selecionada de 'Em Ilegiveis'
            ''''''''''''''''''''''''''''''''''''''''''
            If Opcao = 1 Or Opcao = 2 Then
                If AlterouDocto = True Then
                    'A Capa anterior sofreu alteração
                    If Not VerificaDoctosIndefinidos Then
                        Call AtualizaStatusCapa(IdSelecionado, "8")
                    Else
                        Call AtualizaStatusCapa(IdSelecionado, "5")
                    End If
                Else
                    'A Capa anterior não sofreu alteração , Voltar o Status para '5'
                    Call AtualizaStatusCapa(IdSelecionado, "5")
                End If
                lblOcorrencia.Caption = ""
                lblMotivoDoctoIlegivel.Caption = ""
            End If
            
            '''''''''''''''''''''''''''''
            'Selecionar capas para ???  '
            '''''''''''''''''''''''''''''
            Select Case Opcao
            Case 1
                ''''''''''''''''''''''
                'Selecionou um Titulo'
                ''''''''''''''''''''''
                Call PreencheListCapas(0, IdItem)
            Case 2
                ''''''''''''''''''''''
                'Selecionou um módulo'
                ''''''''''''''''''''''
                Call PreencheListCapas(IdItem, 0)
            End Select
        Else
            chkFiltro.Value = vbUnchecked
        End If
    Else
        ''''''''''''''''''''
        'Seleção sem filtro'
        ''''''''''''''''''''
        Call CmdAtualizar_Click
    End If
End Sub

Private Sub CmdAtualizar_Click()

  If Screen.MousePointer = vbDefault Then
    Screen.MousePointer = vbHourglass

    If AlterouDocto = True Then
      'A Capa anterior sofreu alteração
      If IdSelecionado <> 0 Then
'        If Not VerificaDoctosIndefinidos Then
'          Call AtualizaStatusCapa(IdSelecionado, "8")
'        Else
          Call AtualizaStatusCapa(IdSelecionado, "5")
'        End If
      End If
    ElseIf IdSelecionado <> 0 Then
      'A Capa anterior não sofreu alteração , Voltar o Status para '5'
      Call AtualizaStatusCapa(IdSelecionado, "5")
      IdSelecionado = 0
    End If
    chkFiltro.Value = vbUnchecked
    Screen.MousePointer = vbDefault

    If Not PreencheListCapas Then
      MsgBox "Não Existem Envelopes / Malotes Ilegíveis.", vbInformation, App.Title
      Call HabilitaTimerPesquisa
    End If
  Else
    Call HDMalote(False)
  End If
End Sub

Private Sub cmdAuditoria_Click()

    Geral.Capa.IdCapa = aCapa(lstCapa.ListIndex + 1).IdCapa
    Geral.Capa.Capa = aCapa(lstCapa.ListIndex + 1).Capa
    Geral.Capa.Num_Malote = aCapa(lstCapa.ListIndex + 1).NumMalote
    Geral.Capa.AgOrig = aCapa(lstCapa.ListIndex + 1).AgOrig
    Geral.Capa.IdEnv_Mal = aCapa(lstCapa.ListIndex + 1).IdEnv_Mal
    
    Call Auditoria
    
    Geral.Capa.IdCapa = 0
    Geral.Capa.Capa = 0
    Geral.Capa.Num_Malote = 0
    Geral.Capa.AgOrig = 0
    Geral.Capa.IdEnv_Mal = ""
    
End Sub

Private Sub cmdCalculadora_Click()

  Dim strCommand As String

  strCommand = Space(254)
  GetWindowsDirectory strCommand, 254
  strCommand = Trim(strCommand)
  strCommand = Left(strCommand, Len(strCommand) - 1) & "\calc.exe"

  WinExec strCommand, 9
End Sub

Private Sub cmdCancelar_Click()

   frmLocalizar.Visible = False
   grdDocto.SetFocus
   
End Sub

Private Sub cmdComentario_Click()

    Dim sStr        As String
    
    
    If FrmPesquisa.Visible = False Then

        sStr = aCapa(lstCapa.ListIndex + 1).Comentario

        If Comentario.ShowModal(sStr) Then
        
            If Not InsereControleCapa(Geral.DataProcessamento, aCapa(lstCapa.ListIndex + 1).IdCapa, sStr, aCapa(lstCapa.ListIndex + 1).IdModuloAnterior) Then
                MsgBox "Nâo foi possível inserir o comentário.", vbExclamation
            End If
            
            Call CmdAtualizar_Click
        
        End If
        
        
    End If
End Sub

Private Sub cmdEncerrar_Click()

  Dim X As Integer
  Dim PossuiDocIndef As Boolean

  'Verificar se existe algum processo em andamento
  If Screen.MousePointer = vbDefault And FrmPesquisa.Visible = False Then

    'Verificar se existe alguma selecionada para Encerramento
    If lstCapa.ListIndex = -1 Then
      MsgBox "Nenhum Envelope / Malote selecionado.", vbInformation, App.Title
      Exit Sub
    End If

    PossuiDocIndef = VerificaDoctosIndefinidos

    If bCapaDuplicada = True Then
      'Capa Duplicada
      'Verificar se encontrou documentos indefinidos
      If PossuiDocIndef Then
        If MsgBox("Este " & LblEnv_Mal.Caption & " possui documentos indefinidos que podem ser Envelopes ou Malotes. Confirma Encerramento ?", vbYesNo + vbInformation) = vbYes Then
          Call ExcluiCapaDuplicada
        Else
          grdDocto.SetFocus
          Exit Sub
        End If
      Else
        Call ExcluiCapaDuplicada
      End If

      'Limpando a variável que armazena a capa Atual
      IdSelecionado = 0

      Screen.MousePointer = vbDefault

      'Posicionar na próxima Capa da Lista
      Call LimpaListaDocto

      If lstCapa.ListIndex + 1 < lstCapa.ListCount Then
        'Existem mais Capas -> Posicionar
        lstCapa.ListIndex = lstCapa.ListIndex + 1
      Else
        Call CmdAtualizar_Click
      End If
    Else
      'Capa não duplicada
      'Verificar se a capa possui algum documento para recaptura
      If PossuiDoctoRecaptura Then
        If MsgBox("Confirma o envio da capa para RECAPTURA ?", vbYesNo) = vbYes Then
          Call FinalizaCapa("A")
        End If
      Else
        'Verificar se a capa possui documentos indefinidos
        If PossuiDocIndef Then
          MsgBox "Não é permitido encerrar uma capa com documentos indefinidos.", vbInformation + vbOKOnly, App.Title
          grdDocto.SetFocus
          Exit Sub
        Else
          If MsgBox("Confirma o Encerramento da Capa ?", vbYesNo) = vbYes Then
            Call FinalizaCapa("8")
          End If
        End If
      End If
    End If
  End If

  If (grdDocto.Visible = True) And (grdDocto.Enabled = True) Then grdDocto.SetFocus
  
End Sub

Private Sub cmdEnviarCSP_Click()
    Dim CapaVazia As Boolean

    CapaVazia = False

    If FrmPesquisa.Visible = True Then Exit Sub
    If frmLocalizar.Visible Then Exit Sub

    'Verificar se há alguma capa selecionada
    If lstCapa.ListIndex = -1 Then
        MsgBox "Nenhum Envelope / Malote selecionado.", vbInformation, App.Title
        grdDocto.SetFocus
        Exit Sub
    End If

    'Verificar se a capa possui documentos
    If grdDocto.Rows = 1 Then
        'Capa possui apenas 1 documento -> Verificar se é uma capa
        If aDoc(grdDocto.Row + 1).TipoDocto = 1 Then
            'Capa não possui documentos , somente capa
            CapaVazia = True
            MsgBox "Capa não possui documentos a serem enviados para CSP ", vbInformation + vbOKOnly, App.Title
            grdDocto.SetFocus
            Exit Sub
        End If
    End If

    'Verificar se a capa possui algum documento indefinido
    If Not CapaVazia Then
        If VerificaDoctosIndefinidos Then
            MsgBox "Este " & LblEnv_Mal.Caption & " não pode ser enviado para CSP " & _
            " porque possui documento(s) indefinido(s).", vbInformation + vbOKOnly, App.Title
            Exit Sub
        End If
    End If

    If MsgBox("Confirma o Envio da Capa para CSP ?", vbYesNo) = vbYes Then
        Screen.MousePointer = vbHourglass

        'Atualizar o Status para '1' - Capa Digitalizada
        Call AtualizaStatusCapa(aCapa(lstCapa.ListIndex + 1).IdCapa, "N")
        aCapa(lstCapa.ListIndex + 1).Status = "N"

        'Gravar Log
        Call GravaLog(aCapa(lstCapa.ListIndex + 1).IdCapa, 0, 272)

        IdSelecionado = 0
        Screen.MousePointer = vbDefault

        Call LimpaListaDocto

        'Posicionar na próxima Capa da Lista
        If lstCapa.ListIndex + 1 < lstCapa.ListCount Then
            'Existem mais Capas -> Posicionar
            lstCapa.ListIndex = lstCapa.ListIndex + 1
        Else
            Call CmdAtualizar_Click
        End If
    End If

    grdDocto.SetFocus

End Sub

Private Sub cmdEnviarSupervisor_Click()

  Dim X As Integer
  Dim PossuiDocIndef As Boolean
  Dim CapaVazia As Boolean

  CapaVazia = False

  If FrmPesquisa.Visible = True Then Exit Sub

  'Verificar se há alguma capa selecionada
  If lstCapa.ListIndex = -1 Then
    MsgBox "Nenhum Envelope / Malote selecionado.", vbInformation, App.Title
    grdDocto.SetFocus
    Exit Sub
  End If

  'Verificar se a capa possui documentos
  If grdDocto.Rows = 1 Then
    'Capa possui apenas 1 documento -> Verificar se é uma capa
    If aDoc(grdDocto.Row + 1).TipoDocto = 1 Then
      'Capa não possui documentos , somente capa
      CapaVazia = True
    End If
  End If

  'Verificar se a capa possui algum documento indefinido
  If Not CapaVazia Then
    If Not VerificaDoctosIndefinidos Then
      MsgBox "Este " & LblEnv_Mal.Caption & " não pode ser enviado para Complementação " & _
             " porque não possui documentos indefinidos.", vbInformation + vbOKOnly, App.Title
      Exit Sub
    End If
  End If

  'Verifica se existe documento para recaptura
  For X = 0 To grdDocto.Rows - 1
      If aDoc(X + 1).Status = "A" Then
          MsgBox "Não é permitido reenviar capa para complementação contendo recaptura.", vbInformation, App.Title
          Exit Sub
      End If
  Next

  If MsgBox("Confirma o Envio da Capa para Complementação ?", vbYesNo) = vbYes Then
    Screen.MousePointer = vbHourglass

    'Atualizar o Status para '1' - Capa Digitalizada
    Call AtualizaStatusCapa(aCapa(lstCapa.ListIndex + 1).IdCapa, "1")
    aCapa(lstCapa.ListIndex + 1).Status = "1"
    
    IdSelecionado = 0
    Screen.MousePointer = vbDefault

    Call LimpaListaDocto

    'Posicionar na próxima Capa da Lista
    If lstCapa.ListIndex + 1 < lstCapa.ListCount Then
      'Existem mais Capas -> Posicionar
      lstCapa.ListIndex = lstCapa.ListIndex + 1
    Else
      Call CmdAtualizar_Click
    End If
  End If

  grdDocto.SetFocus
  
End Sub
Private Sub cmdExcluirCapa_Click()

  On Error GoTo ERRO_EXCLUIRCAPA
    
  If Screen.MousePointer = vbDefault And FrmPesquisa.Visible = False Then

    'Verificar se existe alguma capa selecionada
    If lstCapa.ListIndex = -1 Then
      MsgBox "Nenhum Envelope / Malote selecionado.", vbInformation, App.Title
      Exit Sub
    End If

    'Verificar se a capa está duplicada
    If bCapaDuplicada Then
      MsgBox "Não é permitido excluir Envelopes / Malotes Duplicados.", vbInformation + vbOKOnly, App.Title
      Exit Sub
    End If

    'Verificar se a capa pode ser excluida
    If Not VerificaDoctosExcluidosCapa(aCapa(lstCapa.ListIndex + 1).IdCapa) Then
      MsgBox "Não é permitido excluir Envelopes / Malotes em que todos os documentos possuam ocorrência.", vbInformation + vbOKOnly, App.Title
      Exit Sub
    End If

    'Chamar Tela para informar 'Motivo de Exclusão'
    Load MotivoExclusao

    If aCapa(lstCapa.ListIndex + 1).IdEnv_Mal = "E" Then
      MotivoExclusao.LblNroEnv_Mal.Caption = "Nro. Envelope"
    Else
      MotivoExclusao.LblNroEnv_Mal.Caption = "Nro. Malote"
    End If

    MotivoExclusao.LblValorEnv_Mal.Caption = aCapa(lstCapa.ListIndex + 1).Capa

    'Guardar IDCapa para ser usado na atualização da tabela 'MotivoExclusao'
    MotivoExclusao.LblValorEnv_Mal.Tag = aCapa(lstCapa.ListIndex + 1).IdCapa

    MotivoExclusao.Show vbModal, Me

    'Verificar se a Exclusão da Capa foi efetivada
    If MotivoExclusao.Result = True Then
      'Atualizar o Status da Capa para 'D'
      Call AtualizaStatusCapa(aCapa(lstCapa.ListIndex + 1).IdCapa, "D")
      aCapa(lstCapa.ListIndex + 1).Status = "D"

      Unload MotivoExclusao

      'Limpando a variável que armazena a capa Atual
      IdSelecionado = 0

      'Posicionar na próxima Capa da Lista
      Call LimpaListaDocto

      If lstCapa.ListIndex + 1 < lstCapa.ListCount Then
        'Existem mais Capas -> Posicionar
        lstCapa.ListIndex = lstCapa.ListIndex + 1
      Else
        'Não existem mais capas -> Emitir Mensagem
        Call CmdAtualizar_Click
      End If
    End If
    Unload MotivoExclusao
  End If

  If grdDocto.Rows <> 0 Then
     grdDocto.SetFocus
  End If
  
  Exit Sub

ERRO_EXCLUIRCAPA:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Excluir a Capa do Documento", Err, rdoErrors)
    Case vbCancel, vbRetry
      Unload Me
  End Select
End Sub
Private Sub CmdFechar_Click()

  If TabTipoDoc.Visible = True Then
    Call CmdFecharTiposDocto_Click(0)
  Else
    Unload Me
  End If
End Sub
Private Sub CmdFecharPesquisa_Click()

  ''''''''''''''''''''''''''''''''''''''''''
  'Grava log MDI - Inicio Aguarda documento'
  ''''''''''''''''''''''''''''''''''''''''''
  Call GravaLog(0, 0, 253)


  Call CmdFechar_Click
  
End Sub
Private Sub CmdFecharTiposDocto_Click(Index As Integer)

  Call HDObjetos(True)
  TabTipoDoc.Visible = False
  PicTiposDoc.Visible = False
  grdDocto.SetFocus
  
End Sub
Public Sub cmdFrenteVerso_Click()

  On Error GoTo ERRO_FRENTEVERSO

  If teclou Then Exit Sub

  If frmImagem.Visible = False Then Exit Sub

  teclou = True
  'só muda para frente/verso qdo docto vem da Ls500 e da Vips,
  'poi, o canon não gera verso.
  If (aDoc(grdDocto.Row + 1).Ordem = "0") Or (aDoc(grdDocto.Row + 1).Ordem = "2") Then
    If Lead1.Tag = "V" Then
        Lead1.Tag = "F"     'se verso, mostrar frente
        With Lead1
            .AutoRepaint = False
            If Geral.VIPSDLL = eDllProservi Then
              .Load Geral.DiretorioImagens & aDoc(grdDocto.Row + 1).Frente, 0, 0, 1
            Else
              .Load Geral.DiretorioImagens & Format(aCapa(lstCapa.ListIndex + 1).IdLote, "000000000") & "\" & aDoc(grdDocto.Row + 1).Frente, 0, 0, 1
            End If
  
            'se ls500 mostrar mais escuro
            If (aDoc(grdDocto.Row + 1).Ordem = "2") Then
               .Intensity 140
            Else
               .Intensity 220
            End If
           .PaintZoomFactor = 100
           .AutoRepaint = True
        End With
    Else
        Lead1.Tag = "V"     'se frente, mostrar verso
        With Lead1
            .AutoRepaint = False
  
            If Geral.VIPSDLL = eDllProservi Then
              .Load Geral.DiretorioImagens & Trim(aDoc(grdDocto.Row + 1).Verso), 0, 0, 1
            Else
              .Load Geral.DiretorioImagens & Format(aCapa(lstCapa.ListIndex + 1).IdLote, "000000000") & "\" & aDoc(grdDocto.Row + 1).Verso, 0, 0, 1
            End If
  
            'se ls500 mostrar mais escuro
            If (aDoc(grdDocto.Row + 1).Ordem = "2") Then
               .Intensity 140
            Else
               .Intensity 220
            End If
            .PaintZoomFactor = 100
            .AutoRepaint = True
        End With
    End If
  End If
  
  DoEvents
  teclou = False
  Exit Sub

ERRO_FRENTEVERSO:
  frmImagem.Visible = False
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Manipular Imagem do Documento.", Err, rdoErrors)
    Case vbCancel, vbRetry
      Unload Me
  End Select
End Sub
Public Sub cmdInverteCor_Click()

  On Error GoTo ERRO_INVERTECOR

  If teclou Then Exit Sub

  If frmImagem.Visible = False Then Exit Sub

  teclou = True
  Lead1.Invert
  DoEvents
  teclou = False
  Exit Sub

ERRO_INVERTECOR:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao manipular Imagem do Documento.", Err, rdoErrors)
    Case vbCancel, vbRetry
      Unload Me
  End Select
End Sub

Private Sub CmdLocalizar_Click()

   If FrmPesquisa.Visible = False Then
      frmLocalizar.Visible = True
      txtNumEnvMal.SetFocus
   End If
End Sub
Private Sub cmdOcorrencia_Click()

    On Error GoTo ERRO_OCORRENCIA

    Dim X               As Integer
    Dim Valor           As Currency
    Dim DoctosZerados   As Boolean
    Dim StatusDocto     As String
    Dim iLnInicial      As Integer
    Dim iLnFinal        As Integer
    Dim strDescricao    As String

    If FrmPesquisa.Visible = True Then Exit Sub

    'Verificar se a lista de documentos está preenchida
    If grdDocto.Rows = 0 Then
        MsgBox "Nenhum Documento selecionado.", vbInformation, App.Title
        Exit Sub
    End If

    'Verificar se a capa está duplicada
    If bCapaDuplicada Then
        MsgBox "Não é permitido gerar ocorrência para documentos de capas duplicadas.", vbInformation, App.Title
        Exit Sub
    End If

    'Verificar se há mais de 1 documento selecionado
    If grdDocto.Row <> grdDocto.RowSel Then
        
        If grdDocto.RowSel > grdDocto.Row Then
            iLnInicial = grdDocto.Row
            iLnFinal = grdDocto.RowSel
        Else
            iLnInicial = grdDocto.RowSel
            iLnFinal = grdDocto.Row
        End If
        
        For X = iLnInicial To iLnFinal
            'Verificar se o documento é uma Ajuste
            If aDoc(X + 1).TipoDocto = "32" Or aDoc(X + 1).TipoDocto = "33" Or aDoc(X + 1).TipoDocto = "34" Or aDoc(X + 1).TipoDocto = "38" Then
                MsgBox "Não é permitido gerar Ocorrência para Ajustes.", vbInformation, App.Title
                Exit Sub
            End If

            'Verificar se o documento é uma Capa
            If aDoc(X + 1).TipoDocto = "1" Then
                MsgBox "Não é permitido gerar Ocorrência de Capas de " & LblEnv_Mal.Caption & ".", vbInformation, App.Title
                Exit Sub
            End If

            'Verificar se o documento está duplicado
            If aDoc(X + 1).Duplicidade = True Then
                MsgBox "Não é permitido gerar Ocorrência para Documentos Duplicados.", vbInformation, App.Title
                Exit Sub
            End If

            'Verificar se a agencia de origem está zerada
            If aCapa(lstCapa.ListIndex + 1).AgOrig = 0 And aDoc(X + 1).TipoDocto <> 1 Then
                MsgBox "Não é permitido alterar documentos de Envelopes / Malotes Ilegíveis.", vbInformation + vbOKOnly, App.Title
                Exit Sub
            End If

            'Verificar se o documento não possui valor
            If aDoc(X + 1).Valor = 0 Then
                DoctosZerados = True
            End If

            DoEvents
        Next X

    ElseIf grdDocto.Row = grdDocto.RowSel Then
        '--- Apenas 1 documento selecionado ---

        'Verificar se o documento é um Ajuste
        If aDoc(grdDocto.Row + 1).TipoDocto = "32" Or aDoc(grdDocto.Row + 1).TipoDocto = "33" Or aDoc(grdDocto.Row + 1).TipoDocto = "34" Or aDoc(grdDocto.Row + 1).TipoDocto = "38" Then
            MsgBox "Não é permitido gerar Ocorrência para Ajustes.", vbInformation, App.Title
            Exit Sub
        End If

        'Verificar se o documento é uma Capa
        If aDoc(grdDocto.Row + 1).TipoDocto = "1" Then
            MsgBox "Não é permitido gerar Ocorrências de Capas de " & LblEnv_Mal.Caption & ".", vbInformation, App.Title
            grdDocto.SetFocus
            Exit Sub
        End If

        'Verificar se o documento está duplicado
        If aDoc(grdDocto.Row + 1).Duplicidade = True Then
            MsgBox "Não é permitido gerar Ocorrência para Documentos Duplicados.", vbInformation, App.Title
            Exit Sub
        End If

        'Verificar se a agencia de origem está zerada
        If aCapa(lstCapa.ListIndex + 1).AgOrig = 0 And aDoc(grdDocto.Row + 1).TipoDocto <> 1 Then
            MsgBox "Não é permitido alterar documentos de Envelopes / Malotes Ilegíveis.", vbInformation + vbOKOnly, App.Title
            Exit Sub
        End If

        'Verificar se o documento possui valor
        If aDoc(grdDocto.Row + 1).Valor = 0 And aDoc(grdDocto.Row + 1).TipoDocto <> 39 Then
            DoctosZerados = True
        End If
    End If

    'Verificar se o documento não possui valor
    If DoctosZerados Then
        Load DocumentoNaoDefinido
        DocumentoNaoDefinido.Top = Me.Top + 132
        DocumentoNaoDefinido.Show vbModal, Me

        'Verificar se foi informado um valor
        If DocumentoNaoDefinido.InformouValor = True Then
            Valor = DocumentoNaoDefinido.Valor
        Else
            Unload DocumentoNaoDefinido
            grdDocto.SetFocus
            Exit Sub
        End If

        Unload DocumentoNaoDefinido

    End If

    'Busca descrição do complemento de ocorrência, caso exista
    strDescricao = ""
'''    Call GravaComplementoOcorrencia(aDoc(GrdDocto.Row + 1).IdDocto, "C", strDescricao)
    
    Ocorrencia.m_Descricao = Trim(strDescricao)

    Ocorrencia.Show vbModal, Me

    '''''''''''''''''''''''''''''''''''''
    '1 = Colocou ou alterou a ocorrência'
    '2 = Removeu a ocorrência           '
    '''''''''''''''''''''''''''''''''''''
    If Ocorrencia.Result = 1 Or Ocorrencia.Result = 2 Then
    
        aCapa(lstCapa.ListIndex + 1).AlterouDocto = True
        AlterouDocto = True

        'Foi escolhida uma Ocorrência -> Atualizar Documentos selecionados
        If grdDocto.RowSel > grdDocto.Row Then
            iLnInicial = grdDocto.Row
            iLnFinal = grdDocto.RowSel
        Else
            iLnInicial = grdDocto.RowSel
            iLnFinal = grdDocto.Row
        End If
        
        For X = iLnInicial To iLnFinal
                
          'Atualizar o Campo 'OCORRENCIA'
          Set qryAtualizaOcorrencia = Geral.Banco.CreateQuery("", "{? = call AtualizaOcorrenciaDocumento (?,?,?)}")
          With qryAtualizaOcorrencia
              .rdoParameters(0).Direction = rdParamReturnValue
              .rdoParameters(1) = Geral.DataProcessamento  'Data Proc.
              .rdoParameters(2) = aDoc(X + 1).IdDocto      'IdDocto
              .rdoParameters(3) = Ocorrencia.CodOcorr      'Código da Ocorrencia
              .Execute
          End With

          If qryAtualizaOcorrencia(0).Value = 1 Then
              MsgBox "Ocorreu um erro ao atualizar a ocorrência do documento.", vbInformation + vbOKOnly, App.Title
              Exit Sub
          End If
          
'''          If Ocorrencia.Result = 2 Then
'''              'Exclui Complemento da Ocorrência
'''              If Not GravaComplementoOcorrencia(aDoc(X + 1).IdDocto, "E", "") Then Exit Sub
'''          Else
'''              'Grava/Altera Complemento da Ocorrência
'''              If Not GravaComplementoOcorrencia(aDoc(X + 1).IdDocto, IIf(Ocorrencia.m_Descricao = "", "E", "G"), Ocorrencia.m_Descricao) Then Exit Sub
'''          End If
          StatusDocto = "D"
          ''''''''''''''''''''''''''''''''''
          'Selecionou remoção da ocorrência'
          ''''''''''''''''''''''''''''''''''
          If Ocorrencia.Result = 2 Then
              StatusDocto = "0"
              If aDoc(X + 1).TipoDocto <> 0 Then
                  StatusDocto = "1"
              End If
          End If

          If Not AtualizaStatusDocumento(aDoc(X + 1).IdDocto, StatusDocto) Then
              MsgBox "Ocorreu um erro ao atualizar o status do documento.", vbInformation + vbOKOnly, App.Title
              Exit Sub
          End If

          If aDoc(X + 1).TipoDocto <> 39 Then
              If CCur(Valor) <> 0 And CCur(aDoc(X + 1).Valor = 0) Then
                  'Atualizar o Campo 'VALOR' da tabela Documento
                  Set qryAtualizaValorDocumento = Geral.Banco.CreateQuery("", "{? = call AtualizaValorDocIlegiveis (?,?,?)}")
                  With qryAtualizaValorDocumento
                      .rdoParameters(0).Direction = rdParamReturnValue
                      .rdoParameters(1) = Geral.DataProcessamento    'Data Proc.
                      .rdoParameters(2) = aDoc(X + 1).IdDocto        'IdDocto
                      .rdoParameters(3) = Valor                      'Valor do Documento
                      .Execute
                  End With
    
                  If qryAtualizaValorDocumento(0).Value = 1 Then
                      MsgBox "Ocorreu um erro ao atualizar o valor do documento.", vbInformation + vbOKOnly, App.Title
                      Exit Sub
                  End If
              End If
          End If

          'Gravar Log
          Call GravaLog(aCapa(lstCapa.ListIndex + 1).IdCapa, aDoc(X + 1).IdDocto, 2)

        Next X

        Unload Ocorrencia
        Call PreencheListDocto(grdDocto.Row)
    End If

    grdDocto.SetFocus
    Exit Sub

ERRO_OCORRENCIA:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro ao preparar Documento para Ocorrência.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
End Sub
Private Sub cmdProcurar_Click()

    Dim iIndex                   As Integer
    Dim Encontrou                As Boolean
    Dim qryGetDescStatusCapa     As rdoQuery 'Pega descricao do status da capa
    Dim sCapa                    As String
    
    If (Trim(txtNumEnvMal.Text) = "") Or (Not IsNumeric(txtNumEnvMal.Text)) Then
        MsgBox "Entre com um número de capa válido.", vbExclamation
        txtNumEnvMal.SetFocus
        Exit Sub
    End If
    
    Set qryGetDescStatusCapa = Geral.Banco.CreateQuery("", "{Call GetDescStatusCapa(?,?,?)}")

    If Trim(txtNumEnvMal.Text) <> "" Then
        If IsNumeric(txtNumEnvMal.Text) Then
            'Atualizar a lista de capas antes da pesquisa
            Call CmdAtualizar_Click
            
            'Verificar se a capa informada está na lista de capas
            For iIndex = 0 To lstCapa.ListCount - 1
                If CDbl(Left(lstCapa.List(iIndex), 13)) = CDbl(txtNumEnvMal.Text) Then
                    lstCapa.Selected(iIndex) = True
                    Encontrou = True
                    Exit For
                End If
                DoEvents
            Next iIndex
        End If
    End If


    sCapa = txtNumEnvMal.Text
    txtNumEnvMal.Text = ""
    frmLocalizar.Visible = False
    
    'Verificar se encontrou a capa
    If Not Encontrou Then
    
        With qryGetDescStatusCapa
            .rdoParameters(0) = Geral.DataProcessamento
            .rdoParameters(1) = CDbl(sCapa)
            .rdoParameters(2).Direction = rdParamOutput
            .Execute
            
            If Trim(.rdoParameters(2).Value) <> "" Then
                MsgBox .rdoParameters(2).Value, vbInformation
            Else
                MsgBox "Capa não Encontrada.", vbInformation + vbOKOnly, App.Title
                
                'Limpa linha contendo mensagem de ocorrência
                lblOcorrencia.Caption = ""
            End If
            
        End With
    
        
        LblEnv_Mal.Caption = ""
        
        Call HDMalote(False)
        lblNumMalote.Caption = ""
        
        If IdSelecionado <> 0 Then
            If AlterouDocto = True Then
                'A Capa anterior sofreu alteração
                If Not VerificaDoctosIndefinidos Then
                   Call AtualizaStatusCapa(aCapa(lstCapa.ListIndex + 1).IdCapa, "8")
                Else
                   Call AtualizaStatusCapa(aCapa(lstCapa.ListIndex + 1).IdCapa, "5")
                End If
            Else
                'A Capa anterior não sofreu alteração , Voltar o Status para '5'
                Call AtualizaStatusCapa(aCapa(lstCapa.ListIndex + 1).IdCapa, "5")
            End If
        End If
        
        lstCapa.ListIndex = -1
        grdDocto.Clear
        grdDocto.Rows = 0
        
        frmImagem.Visible = False
        Screen.MousePointer = vbDefault
    End If
End Sub
Private Sub CmdRecaptura_Click()

    If lstCapa.ListIndex = -1 Then Exit Sub

    If MsgBox("Os documentos selecionados serão marcados para recaptura. Confirma ?", vbYesNo + vbQuestion) = vbYes Then
        Screen.MousePointer = vbHourglass
        Call MarcaDoctoRecaptura
        Screen.MousePointer = vbDefault
    End If
End Sub
Public Sub cmdRotacao_Click()

  On Error GoTo ERRO_ROTACAO

  If teclou Then Exit Sub

  If frmImagem.Visible = False Then Exit Sub

  teclou = True
  Lead1.FastRotate 90
  DoEvents
  teclou = False
  Exit Sub

ERRO_ROTACAO:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao manipular Imagem do Documento.", Err, rdoErrors)
    Case vbCancel, vbRetry
      Unload Me
  End Select
End Sub

Private Sub CmdTipoDocto_Click()

   If FrmPesquisa.Visible = True Then Exit Sub

   'Verificar se a lista de documentos está preenchida
   If grdDocto.Rows = 0 Then
      MsgBox "Nenhum Documento selecionado.", vbInformation, App.Title
      Exit Sub
   End If

   'Verificar se o usuário selecionou apenas um Documento
   If grdDocto.Row <> grdDocto.RowSel Then
      MsgBox "Não é permitido trocar o tipo de mais de um Documento.", vbInformation, App.Title
      Exit Sub
   End If

   'Verificar se a capa está duplicada
   If bCapaDuplicada Then
      MsgBox "Não é permitido trocar o tipo de um documento com a capa duplicada.", vbInformation, App.Title
      Exit Sub
   End If

   'Verificar se o documento é um Ajuste
   If aDoc(grdDocto.Row + 1).TipoDocto = "32" Or aDoc(grdDocto.Row + 1).TipoDocto = "33" Or aDoc(grdDocto.Row + 1).TipoDocto = "34" Or aDoc(grdDocto.Row + 1).TipoDocto = "38" Then
      MsgBox "Não é permitido alterar Ajustes", vbInformation, App.Title
      Exit Sub
   End If

   'Habilitar todos os tipos de documentos (exceto Envelope e Malote)
   Call HabilitaTiposDocto

   'Verificar se o documento é uma Capa
   If aDoc(grdDocto.Row + 1).TipoDocto = "1" Then
      'Capa de Envelope / Malote
      sCapaOuDocumento = "C"

      'Exibir lista com Envelope e Malote
      Call DesmarcaTiposDocto
      Call DesabilitaTiposDocto
      Call HDObjetos(False)
      TabTipoDoc.Visible = True
      PicTiposDoc.Visible = True
      TabTipoDoc.SetFocus

      Exit Sub
   End If

   'Verificar se o Documento está duplicado
   If aDoc(grdDocto.Row + 1).Duplicidade = True Then
      MsgBox "Não é permitido Alterar Documentos Duplicados.", vbInformation, App.Title
      Exit Sub
   End If

   'Verificar se o documento está com ocorrencia
   If Val(aDoc(grdDocto.Row + 1).Ocorrencia) <> 0 Then
      MsgBox "Não é permitido Alterar Documentos com Ocorrências.", vbInformation, App.Title
      Exit Sub
   End If

   'Verificar se a capa esta ilegivel
   If (aCapa(lstCapa.ListIndex + 1).AgOrig = 0 And aDoc(grdDocto.Row + 1).TipoDocto <> 1) Or _
      (aDoc(1).Status <> "1" And aDoc(grdDocto.Row + 1).TipoDocto <> 1) Then
      MsgBox "Não é permitido alterar documentos de Envelopes / Malotes Ilegíveis.", vbInformation + vbOKOnly, App.Title
      Exit Sub
   End If

   'Setar Flag assumindo que o Documento sofreu uma Alteração
   aCapa(lstCapa.ListIndex + 1).AlterouDocto = True
   AlterouDocto = True

   'Chamar TAB com os tipos de documentos disponíveis
   Call DesmarcaTiposDocto
   Call HDObjetos(False)
   TabTipoDoc.Tab = 0
   TabTipoDoc.Visible = True
   PicTiposDoc.Visible = True
   TabTipoDoc.SetFocus
End Sub
Private Sub CmdTrocaOrdem_Click()

   If lstCapa.ListIndex <> -1 Then
      If aCapa(lstCapa.ListIndex + 1).IdCapa <> 0 Then
         'Gravar Log
         Call GravaLog(aCapa(lstCapa.ListIndex + 1).IdCapa, 0, 6)

         Load frmTrocarOrdemDocumento
         frmTrocarOrdemDocumento.setIdCapaDefault aCapa(lstCapa.ListIndex + 1).IdCapa
         frmTrocarOrdemDocumento.Show vbModal
         Call PreencheListDocto(0)
      End If
   End If
End Sub
Public Sub cmdZoomMais_Click()

  On Error GoTo ERRO_ZOOMMAIS

  If teclou Then Exit Sub
  
  If frmImagem.Visible = False Then Exit Sub

  teclou = True
  If Lead1.PaintZoomFactor <= 400 Then
      Lead1.PaintSizeMode = PAINTSIZEMODE_ZOOM
      Lead1.PaintZoomFactor = Lead1.PaintZoomFactor + 10
  End If
  DoEvents
  teclou = False
  Exit Sub

ERRO_ZOOMMAIS:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao manipular Imagem do Documento.", Err, rdoErrors)
    Case vbCancel, vbRetry
      Unload Me
  End Select
End Sub
Public Sub cmdZoomMenos_Click()

  On Error GoTo ERRO_ZOOMMENOS

  If teclou Then Exit Sub

  If frmImagem.Visible = False Then Exit Sub

  teclou = True
  If Lead1.PaintZoomFactor >= 20 Then
      Lead1.PaintSizeMode = PAINTSIZEMODE_ZOOM
      Lead1.PaintZoomFactor = Lead1.PaintZoomFactor - 10
  End If
  DoEvents
  teclou = False
  Exit Sub

ERRO_ZOOMMENOS:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao manipular Imagem do Documento.", Err, rdoErrors)
    Case vbCancel, vbRetry
      Unload Me
  End Select
End Sub
Private Sub MostraImagemLST()

  On Error GoTo ERRO_MOSTRAIMAGEM

  Dim Ret As Long

  hCtl = Lead1.hwnd

  'Coloca imagem na tela
  With Lead1
    .Tag = "F"
    .AutoRepaint = False
    If Geral.VIPSDLL = eDllProservi Then
      .Load Geral.DiretorioImagens & aDoc(grdDocto.Row + 1).Frente, 0, 0, 1
    Else
      .Load Geral.DiretorioImagens & Format(aCapa(lstCapa.ListIndex + 1).IdLote, "000000000") & "\" & aDoc(grdDocto.Row + 1).Frente, 0, 0, 1
    End If
    
    'Se imagem for da ls500, deixar mais escura
    If aDoc(grdDocto.Row + 1).Ordem <> "2" Then
      .Intensity 220
    Else
      .Intensity 140
    End If
    'Se imagem for do canon, diminui em 50% o tamanho
    If aDoc(grdDocto.Row + 1).Ordem <> "1" Then
      .PaintZoomFactor = 100
    Else
      .PaintZoomFactor = 50
    End If
    .AutoRepaint = True
  End With
  
  frmImagem.Visible = True

  'Posiciona imagem sempre no começo
  Ret = SendMessage(hCtl, WM_VSCROLL, SB_TOP, 0)
  Ret = SendMessage(hCtl, WM_HSCROLL, SB_TOP, 0)

  cmdOcorrencia.Enabled = True

  'Habilita Objetos de Manipulação de Imagens
  Call HDObjetosImagem(True)

  DoEvents

  Exit Sub

ERRO_MOSTRAIMAGEM:
  Screen.MousePointer = vbDefault
  MsgBox "Não foi possível exibir a Imagem do Documento, imagem não encontrada.", vbInformation, App.Title
  Call HDObjetosImagem(False)

End Sub

Private Sub MostraImagem()

  On Error GoTo ERRO_MOSTRAIMAGEM

  Dim Ret As Long

  hCtl = Lead1.hwnd

  'Coloca imagem na tela
  With Lead1
    .Tag = "F"
    .AutoRepaint = False
    If Geral.VIPSDLL = eDllProservi Then
      .Load Geral.DiretorioImagens & aDoc(grdDocto.Row + 1).Frente, 0, 0, 1
    Else
      .Load Geral.DiretorioImagens & Format(aCapa(lstCapa.ListIndex + 1).IdLote, "000000000") & "\" & aDoc(grdDocto.Row + 1).Frente, 0, 0, 1
    End If
    
    'Se imagem for da ls500, deixar mais escura
    If aDoc(grdDocto.Row + 1).Ordem <> "2" Then
      .Intensity 220
    Else
      .Intensity 140
    End If
    'Se imagem for do canon, diminui em 50% o tamanho
    If aDoc(grdDocto.Row + 1).Ordem <> "1" Then
      .PaintZoomFactor = 100
    Else
      .PaintZoomFactor = 50
    End If
    .AutoRepaint = True
  End With
  
  frmImagem.Visible = True

  'Posiciona imagem sempre no começo
  Ret = SendMessage(hCtl, WM_VSCROLL, SB_TOP, 0)
  Ret = SendMessage(hCtl, WM_HSCROLL, SB_TOP, 0)

  cmdOcorrencia.Enabled = True

  'Habilita Objetos de Manipulação de Imagens
  Call HDObjetosImagem(True)

'  DoEvents

  Exit Sub

ERRO_MOSTRAIMAGEM:
  Screen.MousePointer = vbDefault
  MsgBox "Não foi possível exibir a Imagem do Documento, imagem não encontrada.", vbInformation, App.Title
  Call HDObjetosImagem(False)

End Sub

Private Sub Form_Activate()

    On Error GoTo ERRO_ACTIVATE

    'Inclusão de chamada a rotina AtualizaAtividade
    Call AtualizaAtividade(9)

    With Lead1
        .UnlockSupport L_SUPPORT_EXPRESS, "YXPQ3XPPVT"
        .UnlockSupport L_SUPPORT_GIFLZW, "0K3RV9UY3EY"
        .UnlockSupport L_SUPPORT_TIFLZW, "9LE75L0FDXHK"
    End With

    'Preencher List com as Capas de Documentos Ilegíveis
    If PrimeiraVez Then
        PrimeiraVez = False

        AlterouDocto = False
        If Not PreencheListCapas Then
            MsgBox "Não Existem Envelopes / Malotes Ilegíveis.", vbInformation, App.Title

            Call HabilitaTimerPesquisa

            Exit Sub
        End If

        sTempo = 0

        'Habilitar o Timer de Atualização
        tmrAtualiza.Enabled = True
    Else
        Call RemoveCapaRecepcionada
    End If

    Exit Sub

ERRO_ACTIVATE:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro ao Ativar Tela.", Err, rdoErrors)
        Case vbCancel, vbRetry
            Unload Me
    End Select
End Sub
Public Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim Ret As Long

    hCtl = Ilegiveis.Lead1.hwnd

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
        Case vbKeyA
            'Letra A
            If TabTipoDoc.Visible = True And TabTipoDoc.Tab = 0 Then
                Call optGenericos_Click(9)
            End If
        Case vbKeyB
            'Letra B
            If TabTipoDoc.Visible = True And TabTipoDoc.Tab = 0 Then
                Call optGenericos_Click(10)
            End If
        Case vbKeyC
            'Letra C
            If TabTipoDoc.Visible = True And TabTipoDoc.Tab = 0 Then
                Call optGenericos_Click(11)
            End If
        Case vbKey1 To vbKey9
            'Teclado Alfa
            If TabTipoDoc.Visible = True Then
                Select Case TabTipoDoc.Tab
                Case 0
                    Call optGenericos_Click(Chr(KeyCode - 1))
                Case 1
                    Call optTributos_Click(UCase(Chr(KeyCode)) - 1)
                Case 2
                    Call optDiversos_Click(UCase(Chr(KeyCode)) - 1)
                End Select
            End If
        Case vbKeyNumpad1 To vbKeyNumpad9
            'Teclado Numerico
            If TabTipoDoc.Visible = True Then
                Select Case TabTipoDoc.Tab
                Case 0
                    Call optGenericos_Click(Chr(KeyCode - 49))
                Case 1
                    Call optTributos_Click(Chr(KeyCode - 49))
                Case 2
                    Call optDiversos_Click(Chr(KeyCode - 49))
                End Select
            End If
        Case vbKeyF6
            TabTipoDoc.Tab = 0
        Case vbKeyF7
            TabTipoDoc.Tab = 1
        Case vbKeyF8
            TabTipoDoc.Tab = 2
        Case vbKeyF11
            Call cmdFrenteVerso_Click
        Case vbKeyEscape
            If TabTipoDoc.Visible = True Then
                Call CmdFecharTiposDocto_Click(0)
                Exit Sub
            End If
        'Manipulação da Imagem
        Case vbKeyDown
            Ret = SendMessage(hCtl, WM_VSCROLL, SB_LINEDOWN, 0)
        Case vbKeyUp
            Ret = SendMessage(hCtl, WM_VSCROLL, SB_LINEUP, 0)
        Case vbKeyLeft
            Ret = SendMessage(hCtl, WM_HSCROLL, SB_LINEUP, 0)
        Case vbKeyRight
            Ret = SendMessage(hCtl, WM_HSCROLL, SB_LINEDOWN, 0)
    End Select
End Sub
Private Sub Form_Load()

    Dim Control As Control
    
    PrimeiraVez = True
  
    For Each Control In Principal.Controls
        If TypeName(Control) = "Menu" Then
            'Verifica se menu de Troca de ordem está habilitado
            If (Control.Index) = 20 Then
                If Not Control.Enabled Then
                    CmdTrocaOrdem.Enabled = False
                End If
                Exit For
            End If
        End If
    Next
    
    '''''''''''''''''''''''''''
    'Loga a acao Entrar Modulo'
    '''''''''''''''''''''''''''
    Call GravaLog(0, 0, 162)
  
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    'Limpa grid  de documentos
    grdDocto.Rows = 0
    grdDocto.Cols = 2
    grdDocto.ColWidth(0) = grdDocto.Width
    grdDocto.ColAlignment(0) = vbLeftButton
'    GrdDocto.AllowBigSelection = True
    
End Sub



Private Sub Form_Unload(Cancel As Integer)

    'Verificar se foi selecionado uma Capa Anteriormente
    If lstCapa.ListIndex + 1 > 0 Then
        If aCapa(lstCapa.ListIndex + 1).Status <> "V" Then
            Call AtualizaStatusCapa(aCapa(lstCapa.ListIndex + 1).IdCapa, "5")
            aCapa(lstCapa.ListIndex + 1).Status = "5"
        End If
    End If

    IdSelecionado = 0
    
    'Desabilitar os timers
    tmrAtualiza.Enabled = False
    tmrPesquisa.Enabled = False
    
    'Finalizar Conexões
    Set qryGetCapa = Nothing
    Set qryGetDocumentos = Nothing
    Set qryAtualizaStatusCapa = Nothing
    Set qryGetOcorr = Nothing
    Set qryGetMotivoIlegiveis = Nothing
    Set qryAtualizaOcorrencia = Nothing
    Set qryGetMotivoExclusao = Nothing
    Set qryInsereMotivoExclusao = Nothing
    Set qryRemoveMotivoExclusao = Nothing
    Set qryAtualizaValorDocumento = Nothing
    Set qryRemoveCapaRecepcionada = Nothing
  
    '''''''''''''''''''''''''
    'Loga a acao Sair Modulo'
    '''''''''''''''''''''''''
    Call GravaLog(0, 0, 163)

  
End Sub

Private Sub GrdDocto_Click()
   
    On Error GoTo ERRO_DOCTOCLICK

    Dim RsOcorr As rdoResultset
    Dim sSql As String
    Dim X As Integer
    Dim sOcorrencia As String


   'Exibir a Figura do Documento Selecionado
   Call MostraImagem

    lblOcorrencia.Caption = ""
    lblMotivoDoctoIlegivel.Caption = ""
    
   'Verifica se o Documento possui Ocorrência
      'Verificar se a ocorrencia começa com 999
        If Left(aDoc(grdDocto.Row + 1).Ocorrencia, 3) = "999" Then
            If aDoc(grdDocto.Row + 1).RetornoTransacao > 0 Then
                Call ObtemRetornoTransacao(aDoc(grdDocto.Row + 1).RetornoTransacao, sOcorrencia)
            Else
                sOcorrencia = "Erro operacional."
            End If
            lblOcorrencia.Caption = sOcorrencia
        
        Else
            'Verificar se o código da ocorrencia possui 3 ou 5 caracteres
            If Val(aDoc(grdDocto.Row + 1).Ocorrencia) > 999 Then
               '5 Posicoes
               sSql = Left(Trim(aDoc(grdDocto.Row + 1).Ocorrencia), 3)
            Else
               '3 Posicoes
               If Right(Trim(aDoc(grdDocto.Row + 1).Ocorrencia), 2) = "00" Then
                  'Ocorrencia atualizada pelo robo
                  sSql = Val(Trim(aDoc(grdDocto.Row + 1).Ocorrencia)) / 100
               Else
                  'Ocorrencia gerada pelo sistema
                  sSql = Val(Trim(aDoc(grdDocto.Row + 1).Ocorrencia))
               End If
            End If
            
            Set qryGetOcorr = Geral.Banco.CreateQuery("", "{call GetOcorrencia (" & sSql & ")}")
            
            Set RsOcorr = qryGetOcorr.OpenResultset(rdOpenStatic, rdConcurReadOnly)
            
            lblOcorrencia.Caption = ""
            If Not RsOcorr.EOF Then
               lblOcorrencia.ForeColor = &HC0&
               lblOcorrencia.Caption = "Ocorrência : " & RsOcorr!Descricao
            End If

        End If

        'Obtem a descrição do motivo de documento enviado para ilegíveis
        If aDoc(grdDocto.Row + 1).CodMotivo <> 0 And (aDoc(grdDocto.Row + 1).TipoDocto = 0 Or aDoc(grdDocto.Row + 1).Status = "0") Then
            Set RsOcorr = Nothing
        
            Set qryGetMotivoIlegiveis = Geral.Banco.CreateQuery("", "{call GetMotivoIlegiveis(" & aDoc(grdDocto.Row + 1).CodMotivo & ")}")
            Set RsOcorr = qryGetMotivoIlegiveis.OpenResultset(rdOpenStatic, rdConcurReadOnly)
            If Not RsOcorr.EOF() Then
                lblMotivoDoctoIlegivel.Caption = " " & RsOcorr!Descricao & ""
            End If
        End If

   Exit Sub

ERRO_DOCTOCLICK:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Selecionar Documento.", Err, rdoErrors)
    Case vbCancel, vbRetry
      Unload Me
  End Select

End Sub

Private Sub GrdDocto_DblClick()

   If Screen.MousePointer = vbDefault Then

      'Verificar se existe algum documento selecionado
'"      If LstDocto.SelCount = 0 Then
'"         Screen.MousePointer = vbDefault
'"         Exit Sub
'"      End If

    'Verifica se selecionada mais de uma linha
    If grdDocto.Row <> grdDocto.RowSel Then
       MsgBox "Não é permitido trocar o tipo de mais de um Documento.", vbInformation, App.Title
       Exit Sub
    End If
      
      'Verificar se a Capa está duplicada
      If bCapaDuplicada And aDoc(grdDocto.Row + 1).TipoDocto <> 1 Then
         MsgBox "Não é permitido alterar documentos de capas duplicadas.", vbInformation, App.Title
         Exit Sub
      End If

      'Verificar se o documento está duplicado
      If aDoc(grdDocto.Row + 1).Duplicidade = True Then
         MsgBox "Não é Permitido Alterar Documentos Duplicados.", vbInformation, App.Title
         Exit Sub
      End If

      'Verificar se o documento possui ocorrência
      If aDoc(grdDocto.Row + 1).Ocorrencia <> 0 Then
         MsgBox "Não é Permitido Alterar Documentos com Ocorrência.", vbInformation, App.Title
         Exit Sub
      End If

      'Verificar se a capa esta ilegivel
      If (aCapa(lstCapa.ListIndex + 1).AgOrig = 0 And aDoc(grdDocto.Row + 1).TipoDocto <> 1) Or _
         (aDoc(1).Status <> "1" And aDoc(grdDocto.Row + 1).TipoDocto <> 1) Then
         MsgBox "Não é permitido alterar documentos de Envelopes / Malotes Ilegíveis.", vbInformation + vbOKOnly, App.Title
         Exit Sub
      End If

      'Setando o primeiro TAB com default
      TabTipoDoc.Tab = 0

      'Setar Flag assumindo que o Documento sofreu uma Alteração
      aCapa(lstCapa.ListIndex + 1).AlterouDocto = True
      AlterouDocto = True

      'Habilitar todos os tipos de documentos (exceto Envelope e Malote)
      Call HabilitaTiposDocto

      'Verificar se o documento possui tipo
      If aDoc(grdDocto.Row + 1).TipoDocto <> 0 Then
         'Já possui um tipo -> Chamar tela de Complementação
         Select Case aDoc(grdDocto.Row + 1).TipoDocto
         Case 1
            'Capa de Envelope / Malote
            sCapaOuDocumento = "C"

            'Exibir lista com Envelope e Malote
            Call DesmarcaTiposDocto
            Call DesabilitaTiposDocto
            Call HDObjetos(False)
            TabTipoDoc.Visible = True
            PicTiposDoc.Visible = True
            TabTipoDoc.SetFocus

         Case 2, 3
            'Deposito
            Call optGenericos_Click(1)

         Case 4
            'ADCC
            Call optGenericos_Click(6)

         Case 5, 6, 7
            'Cheque
            Call optGenericos_Click(0)

         Case 12
            'Titulos
            Call optDiversos_Click(0)

         Case 13
            'Cobrança Registrada
            Call optDiversos_Click(1)

         Case 14
            'Cobrança Especial
            Call optDiversos_Click(2)

         Case 15
            'DARM
            Call optTributos_Click(4)

         Case 16
            'DARF Preto
            Call optTributos_Click(1)

         Case 17
            'DARF Simples
            Call optTributos_Click(0)

         Case 18
            'GARE
            Call optTributos_Click(3)

         Case 20, 21, 22, 23
            'Arrecadação Eletronica
            Call optGenericos_Click(2)

         Case 24, 25, 26
            'Arrecadação com Valor Indexado
            Call optGenericos_Click(4)

         Case 27
            'Arrecadação Convencional
            Call optDiversos_Click(3)

         Case 28, 29, 30, 31
            'Ficha de Compensação
            Call optGenericos_Click(3)

         Case 32, 33, 34, 38
            MsgBox "Não é permitido alterar Ajustes.", vbInformation, App.Title

         Case 35
            'GPS
            Call optTributos_Click(5)

         Case 36
            'CARTAO AVULSO
            Call optGenericos_Click(7)
        
         Case 37
            'OCT
            Call optGenericos_Click(9)

         Case 39
            'CAPA OCT
            Call optGenericos_Click(10)
         Case 40
           'FGTS
           Call optTributos_Click(6)
           
         Case 41
            'LANÇAMENTO INTERNO
            Call optGenericos_Click(11)

         End Select
      Else
         'Ainda não possui um tipo -> Chamar a tela de Tipos de Docto
         Call DesmarcaTiposDocto
         Call HDObjetos(False)
         TabTipoDoc.Visible = True
         PicTiposDoc.Visible = True
         TabTipoDoc.SetFocus
      End If
   End If

End Sub


Private Sub GrdDocto_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = vbKeyReturn Then
    Call GrdDocto_DblClick
  End If

End Sub

Private Sub GrdDocto_SelChange()

    GrdDocto_Click

End Sub

Private Sub Lead1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

  If Button = 1 Then
    Lead1.AutoRubberBand = True
    Lead1.MousePointer = vbCrosshair
  Else
    Call MostraImagem
  End If
End Sub
Private Sub Lead1_RubberBand()

  On Error GoTo ERRO_RUBBERBAND

  Dim zoomleft As Integer
  Dim zoomtop As Integer
  Dim zoomwidth As Integer
  Dim zoomheight As Integer

  Lead1.MousePointer = 0
  'Zoom in on the selection.
  zoomleft = Lead1.RubberBandLeft
  zoomtop = Lead1.RubberBandTop
  zoomwidth = Lead1.RubberBandWidth
  zoomheight = Lead1.RubberBandHeight
  If (zoomwidth = 0) Or (zoomheight = 0) Then
      Exit Sub
  End If

  'Zoom in on the rectangle defined by the rubberband
  Lead1.ZoomToRect zoomleft, zoomtop, zoomwidth, zoomheight
  Lead1.ForceRepaint

  Exit Sub

ERRO_RUBBERBAND:
  MsgBox "Não é possível redimensionar a Imagem.", vbInformation, App.Title
End Sub

Private Sub lstCapa_Click()

  On Error GoTo ERRO_CAPACLICK

  Dim rsDocumentos As rdoResultset
  Dim sSql As String
  Dim X As Integer
  Dim sLinha As String
  Dim Ret As Integer
  Dim Status As String

  If Screen.MousePointer = vbDefault And lstCapa.ListIndex <> -1 Then
    Screen.MousePointer = vbHourglass

    sTempo = 0

    lblMotivoDoctoIlegivel.Caption = ""
    
    If IdSelecionado <> 0 And (IdSelecionado <> aCapa(lstCapa.ListIndex + 1).IdCapa) Then
      If AlterouDocto = True Then
        'A Capa anterior sofreu alteração
        If Not VerificaDoctosIndefinidos Then
          Call AtualizaStatusCapa(IdSelecionado, "8")
        Else
          Call AtualizaStatusCapa(IdSelecionado, "5")
        End If
      Else
        'A Capa anterior não sofreu alteração , Voltar o Status para '5'
        Call AtualizaStatusCapa(IdSelecionado, "5")
      End If
    End If

    'Verificar se a Capa está Duplicada
    If CapaDuplicada Then
      bCapaDuplicada = True
      LblCapaDup.Visible = True
    Else
      bCapaDuplicada = False
      LblCapaDup.Visible = False
    End If

    'Verificar se a capa mudou
    If IdSelecionado = aCapa(lstCapa.ListIndex + 1).IdCapa Then
      Call PreencheListDocto(0)
    Else
      'Verificar se a capa selecionada continua disponivel
      Ret = CapaSelecionadaDisponivel
      If Ret = 0 Then

        'Verificar se existem documentos transmitidos/expedidos ou com NSU na capa
        If VerificaDocumentosTransmitidos Then
          Screen.MousePointer = vbDefault
          Exit Sub
        End If

        'Excluir Ajustes , se houver
        If Not ExcluiAjuste Then Exit Sub

        Call AtualizaStatusCapa(aCapa(lstCapa.ListIndex + 1).IdCapa, "H")
        aCapa(lstCapa.ListIndex + 1).Status = "H"
        
        ''''''''''''''''''''''''''''''''''
        'Grava log "191 - Selecionar Capa'
        ''''''''''''''''''''''''''''''''''
        GravaLog aCapa(lstCapa.ListIndex + 1).IdCapa, 0, 191

        'Guardar as informações da agência de origem da capa
        Call ValidaAgencia(aCapa(lstCapa.ListIndex + 1).AgOrig, 0, False, True)

        'Preencher os dados dos documentos
        Call PreencheListDocto(0)
        
        If Trim(aCapa(lstCapa.ListIndex + 1).Comentario) <> "" Then
            lblOcorrencia.ForeColor = &H800000 'é o que estava na propriedade do label
            lblOcorrencia.Caption = "Comentário : " & Trim(aCapa(lstCapa.ListIndex + 1).Comentario)
        Else
            lblOcorrencia.ForeColor = &HC0& 'é o que estava na propriedade do label
            lblOcorrencia.Caption = ""
        End If
      Else
        Call HDObjetosImagem(False)
        IdSelecionado = 0
        Screen.MousePointer = vbDefault
        Exit Sub
      End If
    End If

    'Verificar se a Capa é 'ENVELOPE' OU 'MALOTE'
    If UCase(Trim(aCapa(lstCapa.ListIndex + 1).IdEnv_Mal)) = "E" Then
      'Envelope
      LblEnv_Mal.Caption = "Envelope"

      Call HDMalote(False)
    Else
      'Malote
      LblEnv_Mal.Caption = "Malote"

      Call HDMalote(True)
      lblNumMalote.Caption = aCapa(lstCapa.ListIndex + 1).NumMalote
    End If

    'Informar o Lote na tela
    lblLote.Caption = Format(Trim(aCapa(lstCapa.ListIndex + 1).IdLote), "0000-00000")

    'Limpar Objetos
    'lblOcorrencia.Caption = ""

    Screen.MousePointer = vbDefault
  End If

  AlterouDocto = False
  IdSelecionado = aCapa(lstCapa.ListIndex + 1).IdCapa

  Exit Sub

ERRO_CAPACLICK:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Selecionar Capa do Documento.", Err, rdoErrors)
    Case vbCancel, vbRetry
      Unload Me
  End Select
End Sub

Function ExcluiAjuste() As Boolean

  On Error GoTo Erro_ExcluiAjuste

  ExcluiAjuste = False

  Set qryRemoveAjusteCapa = Geral.Banco.CreateQuery("", "{? = call RemoveAjusteCapa (?,?)}")
  With qryRemoveAjusteCapa
    .rdoParameters(0).Direction = rdParamReturnValue
    .rdoParameters(1) = Geral.DataProcessamento                   'Data Proc.
    .rdoParameters(2) = aCapa(lstCapa.ListIndex + 1).IdCapa       'IdCapa
    .Execute
  End With

  If qryRemoveAjusteCapa(0).Value = 1 Then
    'Ocorreu um erro
    MsgBox "Ocorreu um erro ao excluir Ajustes.", vbInformation + vbOKOnly, App.Title
    Exit Function
  End If

  ExcluiAjuste = True

  Exit Function

Erro_ExcluiAjuste:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Verificar Ajustes na Capa.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function

Sub DesabilitaTiposDocto()

   On Error GoTo ERRO_DESABILITATIPOSDOCTO

   Dim X As Integer

   'Genéricos
   For X = 0 To OptGenericos.Count - 1
      If X = 5 Or X = 8 Then
         OptGenericos(X).Enabled = True
      Else
         OptGenericos(X).Enabled = False
      End If
      DoEvents
   Next X

   'Tributos
   For X = 0 To OptTributos.Count - 1
      OptTributos(X).Enabled = False
      DoEvents
   Next X

   'Diversos
   For X = 0 To OptDiversos.Count - 1
      OptDiversos(X).Enabled = False
      DoEvents
   Next X

   Exit Sub

ERRO_DESABILITATIPOSDOCTO:
   Screen.MousePointer = vbDefault
   Select Case TratamentoErro("Erro ao preparar tela para seleção de Tipo de Documento.", Err, rdoErrors)
      Case vbCancel
      Case vbRetry
   End Select
End Sub

Sub HabilitaTiposDocto()

    On Error GoTo ERRO_HABILITATIPOSDOCTO

    Dim X As Integer

    'Genéricos
    For X = 0 To OptGenericos.Count - 1
        If X = 5 Or X = 8 Then
            OptGenericos(X).Enabled = False
        ElseIf X = 11 Then
            If aCapa(lstCapa.ListIndex + 1).IdEnv_Mal = "E" Then
                OptGenericos(X).Enabled = False
            Else
                OptGenericos(X).Enabled = True
            End If
        Else
            OptGenericos(X).Enabled = True
        End If
        DoEvents
    Next X

    'Tributos
    For X = 0 To OptTributos.Count - 1
        'optTributos(x).Enabled = Not CBool(x = 6) 'era igual a True
        OptTributos(X).Enabled = True
        DoEvents
    Next X
    
    'Diversos
    For X = 0 To OptDiversos.Count - 1
        OptDiversos(X).Enabled = True
        DoEvents
    Next X

    Exit Sub

ERRO_HABILITATIPOSDOCTO:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro ao preparar tela para seleção de Tipo de Documento.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
End Sub
Sub DesmarcaTiposDocto()

   On Error GoTo ERRO_DESMARCATIPOSDOCTO

   Dim X As Integer

   'Genéricos
   For X = 0 To OptGenericos.Count - 1
      OptGenericos(X).Value = False
      DoEvents
   Next X

   'Tributos
   For X = 0 To OptTributos.Count - 1
      OptTributos(X).Value = False
      DoEvents
   Next X

   'Diversos
   For X = 0 To OptDiversos.Count - 1
      OptDiversos(X).Value = False
      DoEvents
   Next X

   Exit Sub

ERRO_DESMARCATIPOSDOCTO:
   Screen.MousePointer = vbDefault
   Select Case TratamentoErro("Erro ao preparar tela para seleção de Tipo de Documento.", Err, rdoErrors)
      Case vbCancel
      Case vbRetry
   End Select
End Sub

Private Sub optDiversos_Click(Index As Integer)

   On Error GoTo ERRO_OPTDIVERSOS

   'Preencher Type Global com informações do Documento a ser tratado
   Geral.Documento.IdDocto = aDoc(grdDocto.Row + 1).IdDocto
   Geral.Documento.Leitura = aDoc(grdDocto.Row + 1).Leitura
   Geral.Documento.ValorTotal = aDoc(grdDocto.Row + 1).Valor * 100
   Geral.Documento.TipoDocto = aDoc(grdDocto.Row + 1).TipoDocto
   Geral.Documento.Status = aDoc(grdDocto.Row + 1).Status
   
   Geral.Documento.Agencia = aCapa(lstCapa.ListIndex + 1).AgOrig
   Geral.Capa.AgOrig = aCapa(lstCapa.ListIndex + 1).AgOrig
   Geral.Capa.IdEnv_Mal = aCapa(lstCapa.ListIndex + 1).IdEnv_Mal
   Geral.Capa.Num_Malote = aCapa(lstCapa.ListIndex + 1).NumMalote

   TabTipoDoc.Visible = False
   PicTiposDoc.Visible = False
   Call HDObjetos(True)

   If Val(OptDiversos.Count) > Index Then
      If OptDiversos(Index).Enabled = True Then
         Select Case Index
         Case 0
            'Titulo Convencional
            Call ChamaTelaComplementacao(Titulo)
         Case 1
            'Unicobrança Registrada
            Call ChamaTelaComplementacao(CobrancaRegistrada)
         Case 2
            'Unicobrança Especial
            Call ChamaTelaComplementacao(CobrancaEspecial)
         Case 3
            'Arrecadação Convencional
            Call ChamaTelaComplementacao(ArrecConvencional)
         Case Else
            TabTipoDoc.Visible = True
            PicTiposDoc.Visible = True
            Call HDObjetos(False)
         End Select
      End If
   End If

   If grdDocto.Enabled Then
      grdDocto.SetFocus
   End If

   Exit Sub

ERRO_OPTDIVERSOS:
   Screen.MousePointer = vbDefault
   Select Case TratamentoErro("Erro ao Exibir Tela para Complementação de Documentos.", Err, rdoErrors)
      Case vbCancel, vbRetry
         Unload Me
   End Select
End Sub
Private Sub optGenericos_Click(Index As Integer)

   On Error GoTo ERRO_OPTGENERICOS

   'Preencher Type Global com informações do Documento a ser tratado
   Geral.Documento.IdDocto = aDoc(grdDocto.Row + 1).IdDocto
   Geral.Documento.Leitura = aDoc(grdDocto.Row + 1).Leitura
   Geral.Documento.ValorTotal = aDoc(grdDocto.Row + 1).Valor * 100
   Geral.Documento.TipoDocto = aDoc(grdDocto.Row + 1).TipoDocto
   Geral.Documento.Status = aDoc(grdDocto.Row + 1).Status
   
   Geral.Documento.Agencia = Val(aCapa(lstCapa.ListIndex + 1).AgOrig)
   Geral.Capa.AgOrig = Val(aCapa(lstCapa.ListIndex + 1).AgOrig)
   Geral.Capa.IdEnv_Mal = aCapa(lstCapa.ListIndex + 1).IdEnv_Mal
   Geral.Capa.Num_Malote = Val(aCapa(lstCapa.ListIndex + 1).NumMalote)
   
   TabTipoDoc.Visible = False
   PicTiposDoc.Visible = False
   Call HDObjetos(True)

   If Val(OptGenericos.Count) > Index Then
      If OptGenericos(Index).Enabled = True Then
         Select Case Index
         Case 0
            'CHEQUE
            Call ChamaTelaComplementacao(Cheque)
         Case 1
            'DEPOSITO
            Call ChamaTelaComplementacao(Deposito)
         Case 2
            'CONCESSIONARIA
            Call ChamaTelaComplementacao(ArrecEletronica)
         Case 3
            'FICHA DE COMPENSACAO
            Call ChamaTelaComplementacao(FichaCompensacao)
         Case 4
            'COD. BARRAS COM VALOR INDEXADO
            Call ChamaTelaComplementacao(ArrecValorIndexado)
         Case 5
            'CAPA DE ENVELOPE
            Call ChamaTelaComplementacaoCapa(Envelope)
         Case 6
            'AUTORIZACAO DE DÉBITO
            Call ChamaTelaComplementacao(ADCC)
         Case 7
            'CARTÃO CRÉDITO AVULSO
            Call ChamaTelaComplementacao(CartaoAvulso)
         Case 8
            'CAPA DE MALOTE
            Call ChamaTelaComplementacaoCapa(Malote)
         Case 9
            'OCT
            Call ChamaTelaComplementacao(OCT)
         Case 10
            'CAPA OCT
            Call ChamaTelaComplementacao(CapaOCT)
         Case 11
            'Lançamento Interno
            Call ChamaTelaComplementacao(LancamentoInterno)
         Case Else
            TabTipoDoc.Visible = True
            PicTiposDoc.Visible = True
            Call HDObjetos(False)
         End Select
      End If
   End If

   If grdDocto.Enabled Then
      grdDocto.SetFocus
   End If

   Exit Sub

ERRO_OPTGENERICOS:
   Screen.MousePointer = vbDefault
   Select Case TratamentoErro("Erro ao Exibir Tela para Complementação de Documentos.", Err, rdoErrors)
      Case vbCancel, vbRetry
         Unload Me
   End Select
End Sub

Private Sub optTributos_Click(Index As Integer)

   On Error GoTo ERRO_OPTTRIBUTOS

   'Preencher Type Global com informações do Documento a ser tratado
   Geral.Documento.IdDocto = aDoc(grdDocto.Row + 1).IdDocto
   Geral.Documento.Leitura = aDoc(grdDocto.Row + 1).Leitura
   Geral.Documento.ValorTotal = aDoc(grdDocto.Row + 1).Valor * 100
   Geral.Documento.TipoDocto = aDoc(grdDocto.Row + 1).TipoDocto
   Geral.Documento.Status = aDoc(grdDocto.Row + 1).Status
   
   Geral.Documento.Agencia = aCapa(lstCapa.ListIndex + 1).AgOrig
   Geral.Capa.AgOrig = aCapa(lstCapa.ListIndex + 1).AgOrig
   Geral.Capa.IdEnv_Mal = aCapa(lstCapa.ListIndex + 1).IdEnv_Mal
   Geral.Capa.Num_Malote = Val(aCapa(lstCapa.ListIndex + 1).NumMalote)

   TabTipoDoc.Visible = False
   PicTiposDoc.Visible = False
   Call HDObjetos(True)

   If Val(OptTributos.Count) > Index Then
      If OptTributos(Index).Enabled = True Then
         Select Case Index
         Case 0
            'DARF Simples
            Call ChamaTelaComplementacao(DARFSimples)
         Case 1
            'DARF Preto
            Call ChamaTelaComplementacao(DARFPreto)
         Case 2
            'FGTS
            Call ChamaTelaComplementacao(ArrecConvencional)
         Case 3
            'GARE
            Call ChamaTelaComplementacao(GareICMS)
         Case 4
            'DARM
            Call ChamaTelaComplementacao(DARM)
         Case 5
            'GPS
            Call ChamaTelaComplementacao(GPS)
         Case 6
            'FGTS
            Call ChamaTelaComplementacao(frmFGTS)
         Case Else
            TabTipoDoc.Visible = True
            PicTiposDoc.Visible = True
            Call HDObjetos(False)
         End Select
      End If
   End If

   If grdDocto.Enabled Then
      grdDocto.SetFocus
   End If

   Exit Sub

ERRO_OPTTRIBUTOS:
   Screen.MousePointer = vbDefault
   Select Case TratamentoErro("Erro ao Exibir Tela para Complementação de Documentos.", Err, rdoErrors)
      Case vbCancel, vbRetry
         Unload Me
   End Select
End Sub

Private Sub Text1_Change()

End Sub

Private Sub tmrAtualiza_Timer()

    tmrAtualiza.Enabled = False
    
    If lstCapa.ListIndex <> -1 Then
        If aCapa(lstCapa.ListIndex + 1).IdCapa <> 0 Then
            sTempo = sTempo + Int(tmrAtualiza.Interval / 1000)
            If sTempo + Int(tmrAtualiza.Interval / 1000) >= Geral.Intervalo Then
                'Atualizar o Status da Capa
                Call AtualizaStatusCapa(aCapa(lstCapa.ListIndex + 1).IdCapa, "H")
                sTempo = 0
            End If
            '''''''''''''''''''''''''''''''''''''''
            'Grava log MDI - Fim Aguarda documento'
            '''''''''''''''''''''''''''''''''''''''
            Call GravaLog(0, 0, 253)
        End If
    End If
    
    tmrAtualiza.Enabled = True
End Sub

Private Sub tmrPesquisa_Timer()

  tmrPesquisa.Enabled = False

  sTempo = sTempo + Int(tmrPesquisa.Interval / 1000)

  If sTempo + Int(tmrPesquisa.Interval / 1000) >= Geral.Atualizacao Then
    'Pesquisar por Documentos Ilegíveis
    sTempo = 0
    ''''''''''''''''''''''''''''''''''''''''''
    'Grava log MDI - Inicio Aguarda documento'
    ''''''''''''''''''''''''''''''''''''''''''
    Call GravaLog(0, 0, 252)

    If PreencheListCapas Then Exit Sub

    tmrPesquisa.Enabled = True
  End If

  'Atualizar a Barra de Progresso
  If Progress.Value + 4 > 100 Then
    Progress.Value = 0
  Else
    Progress.Value = Progress.Value + 4
  End If

  DoEvents
  tmrPesquisa.Enabled = True
End Sub

Private Sub TxtNumEnvMal_KeyPress(KeyAscii As Integer)

   If KeyAscii = vbKeyReturn Then
      Call cmdProcurar_Click
   ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
      KeyAscii = 0
   End If
End Sub
