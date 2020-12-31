VERSION 5.00
Object = "{00080003-B1BA-11CE-ABC6-F5B2E79D9E3F}#8.0#0"; "LTOCX80N.OCX"
Object = "{9CBA5D64-E3C8-11D3-9FFC-00104BC8688C}#1.0#0"; "CurrencyEdit.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Consulta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta"
   ClientHeight    =   8100
   ClientLeft      =   636
   ClientTop       =   708
   ClientWidth     =   11160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   11160
   Begin TabDlg.SSTab SSTab1 
      Height          =   7932
      Left            =   168
      TabIndex        =   0
      Top             =   84
      Width           =   10812
      _ExtentX        =   19071
      _ExtentY        =   13991
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   420
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Argumentos Pesquisa"
      TabPicture(0)   =   "Consulta.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame_Indices"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame_Campos(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Picture1(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdSair(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdConfirma"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Resultado"
      TabPicture(1)   =   "Consulta.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblExclusao"
      Tab(1).Control(1)=   "Panel_Situacao"
      Tab(1).Control(2)=   "Panel_SitEnv"
      Tab(1).Control(3)=   "lblLote"
      Tab(1).Control(4)=   "Label_Envelope"
      Tab(1).Control(5)=   "lbl_NumeroMalote"
      Tab(1).Control(6)=   "lbl_status"
      Tab(1).Control(7)=   "Grade"
      Tab(1).Control(8)=   "List_detalhe"
      Tab(1).Control(9)=   "Frame5"
      Tab(1).Control(10)=   "Picture4"
      Tab(1).Control(11)=   "picNumMalote"
      Tab(1).Control(12)=   "Picture1(0)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Frame4"
      Tab(1).ControlCount=   14
      Begin VB.Frame Frame2 
         Caption         =   "Consulta padrão"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   972
         Left            =   816
         TabIndex        =   57
         Top             =   828
         Width           =   9168
         Begin VB.TextBox txtcapa 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   360
            Left            =   1956
            MaxLength       =   14
            TabIndex        =   61
            Top             =   408
            Width           =   2028
         End
         Begin VB.OptionButton Opn_Capa 
            Caption         =   "Capa"
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
            Left            =   300
            TabIndex        =   60
            Top             =   432
            Value           =   -1  'True
            Width           =   2880
         End
         Begin VB.TextBox TxtNumMalote 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   360
            Left            =   6924
            MaxLength       =   12
            TabIndex        =   59
            Top             =   408
            Width           =   2028
         End
         Begin VB.OptionButton Opn_NumMalote 
            Caption         =   "Numero Malote"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   4680
            TabIndex        =   58
            Top             =   432
            Width           =   2544
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Imagem"
         Height          =   4596
         Left            =   -74832
         TabIndex        =   50
         Top             =   3216
         Width           =   8736
         Begin VB.PictureBox Picture2 
            BackColor       =   &H80000009&
            Height          =   252
            Left            =   420
            ScaleHeight     =   204
            ScaleWidth      =   180
            TabIndex        =   52
            Top             =   384
            Visible         =   0   'False
            Width           =   228
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H80000009&
            Height          =   252
            Left            =   168
            ScaleHeight     =   204
            ScaleWidth      =   180
            TabIndex        =   51
            Top             =   384
            Visible         =   0   'False
            Width           =   228
         End
         Begin LeadLib.Lead Lead2 
            Height          =   204
            Left            =   168
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   204
            Visible         =   0   'False
            Width           =   456
            _Version        =   524288
            _ExtentX        =   804
            _ExtentY        =   360
            _StockProps     =   229
            BackColor       =   -2147483639
            BorderStyle     =   1
            ScaleHeight     =   15
            ScaleWidth      =   36
            DataField       =   ""
            BitmapDataPath  =   ""
            AnnDataPath     =   ""
         End
         Begin LeadLib.Lead Lead1 
            Height          =   4332
            Left            =   156
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   192
            Width           =   8484
            _Version        =   524288
            _ExtentX        =   14965
            _ExtentY        =   7641
            _StockProps     =   229
            BackColor       =   -2147483643
            BorderStyle     =   1
            ScaleHeight     =   359
            ScaleWidth      =   705
            DataField       =   ""
            BitmapDataPath  =   ""
            AnnDataPath     =   ""
         End
      End
      Begin VB.CommandButton cmdConfirma 
         Caption         =   "&Confirmar"
         Height          =   696
         Left            =   8244
         Picture         =   "Consulta.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   6912
         Width           =   816
      End
      Begin VB.PictureBox Picture1 
         Height          =   288
         Index           =   0
         Left            =   -74832
         ScaleHeight     =   240
         ScaleWidth      =   1620
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   528
         Width           =   1668
         Begin VB.Label Lbl_tpCapa 
            Caption         =   "Capa"
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
            Height          =   264
            Left            =   24
            TabIndex        =   43
            Top             =   -12
            Width           =   1500
         End
      End
      Begin VB.PictureBox picNumMalote 
         Height          =   288
         Left            =   -69545
         ScaleHeight     =   240
         ScaleWidth      =   1056
         TabIndex        =   40
         Top             =   528
         Visible         =   0   'False
         Width           =   1104
         Begin VB.Label lbl_NumMalote 
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
            Height          =   288
            Left            =   48
            TabIndex        =   41
            Top             =   0
            Width           =   1092
         End
      End
      Begin VB.PictureBox Picture4 
         Height          =   288
         Left            =   -71350
         ScaleHeight     =   240
         ScaleWidth      =   456
         TabIndex        =   38
         Top             =   528
         Width           =   504
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
            TabIndex        =   39
            Top             =   -12
            Width           =   480
         End
      End
      Begin VB.Frame Frame5 
         Height          =   4596
         Left            =   -66024
         TabIndex        =   33
         Top             =   3216
         Width           =   1596
         Begin VB.CommandButton cmdSair 
            BackColor       =   &H80000004&
            Caption         =   "Sair"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   708
            Index           =   1
            Left            =   468
            Picture         =   "Consulta.frx":0342
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   3780
            Width           =   840
         End
         Begin VB.CommandButton Imprime_Detalhe 
            Caption         =   "&Imprimir"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   624
            Left            =   468
            Picture         =   "Consulta.frx":064C
            Style           =   1  'Graphical
            TabIndex        =   49
            Top             =   3120
            Width           =   820
         End
         Begin VB.CommandButton cmdZoomMais 
            Caption         =   "Zoom +"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   696
            Left            =   468
            Picture         =   "Consulta.frx":0956
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   240
            Width           =   820
         End
         Begin VB.CommandButton cmdZoomMenos 
            Caption         =   "Zoom -"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   696
            Left            =   468
            Picture         =   "Consulta.frx":0C60
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   960
            Width           =   820
         End
         Begin VB.CommandButton cmdRotacao 
            Caption         =   "Rotação"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   696
            Left            =   468
            Picture         =   "Consulta.frx":0F6A
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   1680
            Width           =   820
         End
         Begin VB.CommandButton cmdFrenteVerso 
            Caption         =   "Frente/Verso"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   696
            Left            =   468
            Picture         =   "Consulta.frx":1274
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   2400
            Width           =   820
         End
      End
      Begin VB.CommandButton cmdSair 
         BackColor       =   &H80000004&
         Cancel          =   -1  'True
         Caption         =   "Sair"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   708
         Index           =   0
         Left            =   9168
         Picture         =   "Consulta.frx":157E
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   6900
         Width           =   840
      End
      Begin VB.PictureBox Picture1 
         Height          =   552
         Index           =   1
         Left            =   5496
         ScaleHeight     =   504
         ScaleWidth      =   4440
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2352
         Width           =   4488
         Begin VB.Label lblTipo 
            Alignment       =   2  'Center
            Caption         =   "Envelope ou Malote"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   312
            Left            =   96
            TabIndex        =   8
            Top             =   60
            Width           =   4248
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   672
         Index           =   1
         Left            =   792
         TabIndex        =   6
         Top             =   2232
         Width           =   3972
         Begin VB.OptionButton optTipo 
            Caption         =   "&Malote Empresa"
            Height          =   192
            Index           =   1
            Left            =   1944
            TabIndex        =   10
            Top             =   300
            Width           =   1764
         End
         Begin VB.OptionButton optTipo 
            Caption         =   "&Envelope"
            Height          =   192
            Index           =   0
            Left            =   216
            TabIndex        =   9
            Top             =   300
            Width           =   1092
         End
      End
      Begin VB.Frame Frame_Campos 
         Enabled         =   0   'False
         Height          =   3696
         Index           =   1
         Left            =   5508
         TabIndex        =   2
         Top             =   3024
         Width           =   4500
         Begin CURRENCYEDITLib.CurrencyEdit TxtVal 
            Height          =   348
            Left            =   2220
            TabIndex        =   22
            Top             =   408
            Width           =   2028
            _Version        =   65537
            _ExtentX        =   3577
            _ExtentY        =   614
            _StockProps     =   93
            ForeColor       =   -2147483635
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
            Enabled         =   0   'False
            MaxLength       =   10
            BackColor       =   -2147483643
         End
         Begin VB.TextBox TxtBco 
            Enabled         =   0   'False
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
            Left            =   2652
            MaxLength       =   3
            TabIndex        =   14
            Top             =   1776
            Width           =   720
         End
         Begin VB.TextBox TxtAg 
            Enabled         =   0   'False
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
            Left            =   2628
            MaxLength       =   4
            TabIndex        =   15
            Top             =   2244
            Width           =   732
         End
         Begin VB.TextBox TxtCta 
            Enabled         =   0   'False
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
            Left            =   2232
            MaxLength       =   7
            TabIndex        =   17
            Top             =   2724
            Width           =   1164
         End
         Begin VB.TextBox TxtChq 
            Enabled         =   0   'False
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
            Left            =   2232
            MaxLength       =   6
            TabIndex        =   19
            Top             =   3168
            Width           =   1152
         End
         Begin VB.TextBox TxtNSU 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   360
            Left            =   2220
            MaxLength       =   6
            TabIndex        =   3
            Top             =   888
            Width           =   2028
         End
         Begin VB.Label LblBanco 
            Caption         =   "Banco"
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   192
            Left            =   180
            TabIndex        =   21
            Top             =   1788
            Width           =   672
         End
         Begin VB.Label LblAgencia 
            Caption         =   "Agência"
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   192
            Left            =   180
            TabIndex        =   20
            Top             =   2244
            Width           =   792
         End
         Begin VB.Label LblCta 
            Caption         =   "Conta"
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   192
            Left            =   180
            TabIndex        =   18
            Top             =   2772
            Width           =   672
         End
         Begin VB.Label LblCheque 
            AutoSize        =   -1  'True
            Caption         =   "Nro. Cheque"
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   192
            Left            =   180
            TabIndex        =   16
            Top             =   3228
            Width           =   1092
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000003&
            BorderStyle     =   6  'Inside Solid
            X1              =   36
            X2              =   4776
            Y1              =   1512
            Y2              =   1512
         End
         Begin VB.Label LblNSU 
            Caption         =   "NSU"
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   228
            Left            =   180
            TabIndex        =   5
            Top             =   1008
            Width           =   552
         End
         Begin VB.Label LblValor 
            Caption         =   "Valor"
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   228
            Left            =   192
            TabIndex        =   4
            Top             =   468
            Width           =   552
         End
      End
      Begin VB.Frame Frame_Indices 
         Caption         =   " Índices "
         Enabled         =   0   'False
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
         Height          =   3720
         Left            =   780
         TabIndex        =   1
         Top             =   3024
         Width           =   3984
         Begin VB.ComboBox cmbTipoDocto 
            Height          =   288
            ItemData        =   "Consulta.frx":1888
            Left            =   1128
            List            =   "Consulta.frx":188A
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   492
            Width           =   2712
         End
         Begin VB.Frame OptOcorrencia 
            Caption         =   "Escolha uma opção"
            Height          =   900
            Left            =   1560
            TabIndex        =   44
            Top             =   1416
            Width           =   2124
            Begin VB.OptionButton OptOcDocto 
               Caption         =   "por Documento"
               Height          =   264
               Left            =   192
               TabIndex        =   46
               Top             =   552
               Width           =   1572
            End
            Begin VB.OptionButton OptOcCapa 
               Caption         =   "por Capa"
               Height          =   264
               Left            =   192
               TabIndex        =   45
               Top             =   288
               Width           =   1572
            End
         End
         Begin VB.OptionButton opn_Ocorrencia 
            Caption         =   "Ocorrência :"
            Height          =   252
            Left            =   204
            TabIndex        =   32
            Top             =   1440
            Width           =   1452
         End
         Begin VB.OptionButton Opn_Banco 
            Caption         =   "Banco + Ag. + Conta + Cheque"
            Height          =   300
            Left            =   180
            TabIndex        =   31
            Top             =   2532
            Width           =   3072
         End
         Begin VB.OptionButton Opn_Vl 
            Caption         =   "Valor"
            Height          =   264
            Left            =   192
            TabIndex        =   30
            Top             =   456
            Width           =   840
         End
         Begin VB.OptionButton Opn_NSU 
            Caption         =   "NSU"
            Height          =   252
            Left            =   192
            TabIndex        =   29
            Top             =   996
            Width           =   840
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo de Documento"
            ForeColor       =   &H00000000&
            Height          =   264
            Left            =   1704
            TabIndex        =   48
            Top             =   276
            Width           =   1512
         End
      End
      Begin VB.ListBox List_detalhe 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1944
         ItemData        =   "Consulta.frx":188C
         Left            =   -74880
         List            =   "Consulta.frx":188E
         TabIndex        =   11
         Top             =   960
         Visible         =   0   'False
         Width           =   10404
      End
      Begin MSFlexGridLib.MSFlexGrid Grade 
         Height          =   1932
         Left            =   -74880
         TabIndex        =   62
         Top             =   960
         Width           =   10452
         _ExtentX        =   18436
         _ExtentY        =   3408
         _Version        =   393216
         FixedCols       =   0
         FocusRect       =   0
      End
      Begin VB.Label lbl_status 
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   216
         Left            =   -64980
         TabIndex        =   27
         Top             =   120
         Width           =   612
      End
      Begin VB.Label lbl_NumeroMalote 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
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
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   -68352
         TabIndex        =   26
         Top             =   516
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label Label_Envelope 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
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
         ForeColor       =   &H00000000&
         Height          =   288
         Left            =   -73130
         TabIndex        =   25
         Top             =   528
         Width           =   1716
      End
      Begin VB.Label lblLote 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
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
         ForeColor       =   &H00000000&
         Height          =   288
         Left            =   -70800
         TabIndex        =   24
         Top             =   528
         Width           =   1176
      End
      Begin VB.Label Panel_SitEnv 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   600
         Left            =   -66744
         TabIndex        =   23
         Top             =   336
         Width           =   2316
      End
      Begin VB.Label Panel_Situacao 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   348
         Left            =   -66372
         TabIndex        =   28
         Top             =   612
         Visible         =   0   'False
         Width           =   1920
      End
      Begin VB.Label lblExclusao 
         BorderStyle     =   1  'Fixed Single
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
         Height          =   312
         Left            =   -74880
         TabIndex        =   55
         Top             =   2892
         Visible         =   0   'False
         Width           =   10452
      End
   End
End
Attribute VB_Name = "Consulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type tpModulo
    qryGetDocumentosNumCapa As rdoQuery ' Leitura da capa com os respectivos doctos
    qryGetAgCcDoc           As rdoQuery ' Leitura de dados por Agencia / Conta Corrente / Documento
    qryGetStatusCapa        As rdoQuery ' Leitura do Status Capa
    qryGetConsStatusCapa    As rdoQuery ' Leitura do Status Capa
    qryGetDocumentosCapaNSU As rdoQuery ' Leitura do docto que atenda ao Nro. NSU desejado
    qryPesquisaBcoAgCtaChq  As rdoQuery ' Leitura do docto que atenda a estes dados
    qryPesquisaValor        As rdoQuery ' Leitura do docto que atenda ao valor desejado
    qryBuscaEnvelope        As rdoQuery ' Leitura do envelope do cheque escolhido
    qryDocumentoEscolhido   As rdoQuery ' Leitura dos dados do docto escolhido
    qryGetADCC              As rdoQuery ' Leitura do adcc
    qryGetDeposito          As rdoQuery ' Leitura do deposito
    qryGetArrecConvenc      As rdoQuery ' Leitura da arrec.convenc.
    qryGetDARFPreto         As rdoQuery ' Leitura do darf preto
    qryGetDARFSimples       As rdoQuery ' Leitura do darf simples
    qryGetGare              As rdoQuery ' Leitura do gare
    qryGetGps               As rdoQuery ' Leitura do gps
    qryGetDarm              As rdoQuery ' Leitura do darm
    qryGetTitulo            As rdoQuery ' Leitura do titulo
    qryGetCobEspTec         As rdoQuery ' Leitura da cob.esp.teclado
    qryGetCobRegTec         As rdoQuery ' Leitura da cob.reg.teclado
    qryGetCobCodBar         As rdoQuery ' Leitura das cobranças com cod.barras
    qryCartaoAvulso         As rdoQuery ' Leitura do cartao credito avulso
    qryGetDoctosNumMalote   As rdoQuery ' Leitura de Documentos Malote
    qryGetTipoOcorr         As rdoQuery ' Leitura de Tipo de Ocorrencia por Capa ou Documentos
    qryGetIdcDocto          As rdoQuery ' Leitura de IdDocto
    qryGetIdDocto           As rdoQuery ' Leitura de IdDocto
    qryGetIdCapa            As rdoQuery ' Leitura de IdCapa
    qryGetOCT               As rdoQuery ' Leitura de OCT
    qryGettipoDocto         As rdoQuery ' Traz a Descrição do Tipo de Documento
    qryGetMotivoExclusao    As rdoQuery ' Leitura do descritivo do Motivo de exclusão
    qryGetFGTS              As rdoQuery ' Leitura dos detalhes do FGTS
    qryGetLctoInterno       As rdoQuery ' Leitura dos detalhes do Lcto Interno
    End Type
Private Modulo As tpModulo

Private qryGetocorrencia    As rdoQuery ' Leitura de descrivo de ocorrências

Dim tbenv                   As rdoResultset ' Recordset Auxiliar
Dim tbenv1                  As rdoResultset ' Recordset Auxiliar
Dim tb1                     As rdoResultset ' Recordset Auxiliar`
Dim tbdoctos                As rdoResultset ' Recordset Auxiliar
Dim RsAux                   As rdoResultset ' Recordset Auxiliar
Dim RsDeposito              As rdoResultset ' Recordset Depositos
Dim RsBHVC                  As rdoResultset ' Recordset Pesquisa BHVC
Dim RsArrConv               As rdoResultset ' Recordset Arrecadação
Dim RsDARFS                 As rdoResultset ' Recordset Darf Simples
Dim RsDarfP                 As rdoResultset ' Recordset Darf Preto
Dim RsCob3                  As rdoResultset ' Recordset Cobranca de 3º
Dim RsADCC                  As rdoResultset ' Recordset ADCC
Dim RsGps                   As rdoResultset ' Recordset GPS
Dim rsGare                  As rdoResultset ' Recordset Gare
Dim RsDARM                  As rdoResultset ' Recordset Darm
Dim RsCartaoAv              As rdoResultset ' Recordset Cartao Avulso
Dim RsOcorrencia            As rdoResultset ' Recordset Ocorrencias
Dim RsOCT                   As rdoResultset ' Recordset OCT
Dim RsStatusCapa            As rdoResultset ' Recordset StatusCapa
Dim RsConsStatusCapa        As rdoResultset ' Recordset Consulta StatusCapa
Dim RsTipoDocto             As rdoResultset ' Recordset Pesquisa TipoDocto
Dim RsFGTS                  As rdoResultset ' Recordset FGTS
Dim RsLctoInterno           As rdoResultset ' Recordset Lançamento Interno
Dim RsMotivoExclusao        As rdoResultset ' Recordset Motivo de Exlusão
Dim RsCobCodBar             As rdoResultset ' Recordset Cobrança com Código de Barras
Dim RsCobEspTec             As rdoResultset ' Recordset Cobrança Especial
Dim RsCobRegTec             As rdoResultset ' Recordset Cobrança Registrada
Dim Lote                    As Double       ' Guarda Numero do Lote

Private m_Busy, bAlterar As Boolean

Private idenv_antigo, iArq As Long

Private detalhe_consulta As String
Private str_formatada As String * 30

Private indice_corrente, achou As Byte

Private Pegiddocto, PegiddoctoSit2 As Long

Private cont_dt, cont_lote, cont_banco As Integer
Private cont_agencia, cont_conta, cont_cheque As Integer
Private rep, Indice, cont_scroll, TipoDocto    As Integer
Private FlagImp As Integer

Private Aux, buf_scroll, texto_scroll, PegHistorico As String
Private valor_caption, visual, Tip_doc, Status, Status_Capa As String
Private reg_imagem, reg_frente, reg_data As String
Private reg_env, TipoDoctoTab As String

'* Guarda o Nº do Identificação Documento
Dim nDocto_Sel As Long

'* Guarda o Nº da Capa (Malote Ou Envelope)
Dim NCapa_Sel As Long

'* Guarda o Valor Selecionado (Molote ou Envelope)
Dim NEnvMal As String

'* Guarda o Valor Selecionado Tipo de Ocorrencia ( por Capa ou Malote)
Dim NCapaDocto As String

Private Function ObtemOcorrencia(ByVal Ocorrencia As String) As String

    Dim sDescricaoTransacao     As String
    Dim iRetornoTransacao       As Integer
    Dim iOcorrencia             As Integer
    
    On Error GoTo ErroOcorrencia
    rdoErrors.Clear
    
    iRetornoTransacao = Val(Right(Ocorrencia, 2))
    iOcorrencia = Val(Left(Ocorrencia, 3))
    
    Screen.MousePointer = vbHourglass
    
    If iRetornoTransacao > 0 Then
        If ObtemRetornoTransacao(iRetornoTransacao, sDescricaoTransacao) Then
            ObtemOcorrencia = sDescricaoTransacao
            Screen.MousePointer = vbDefault
            Exit Function
        End If
    End If
    
    qryGetocorrencia.rdoParameters(0) = iOcorrencia

    Set RsOcorrencia = qryGetocorrencia.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    If RsOcorrencia.EOF Then
        ObtemOcorrencia = "Codigo da Ocorrencia nao existe: " & Trim(str(iOcorrencia))
    Else
        ObtemOcorrencia = RsOcorrencia!Descricao
    End If
    RsOcorrencia.Close

    Screen.MousePointer = vbDefault
    On Error GoTo 0
    Exit Function

ErroOcorrencia:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro na obtenção da Ocorrência do documento.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    m_Busy = False
    Unload Me

End Function


Sub FormataValor(ByVal vl_doc As String)
   
   'se encontrar o ponto (.), não formata
   If InStr(vl_doc, ".") = 0 Then
        valor_caption = Format(vl_doc, "0.00")
   Else
        valor_caption = vl_doc
   End If
   
   str_formatada = ""
   rep = 1                'contador de caracteres a serem formatados
   Do
      If (Mid$(valor_caption, rep, 1) = "") Then   'verifica término da string
         Exit Do
      End If
      rep = rep + 1
   Loop While (rep < 14)   'tamanho máximo da string a ser formatada
   rep = rep - 1
   
   '---- formata à direita ----
   Mid$(str_formatada, 13 - rep + 1, rep) = Mid$(valor_caption, 1, rep)
   valor_caption = str_formatada    'atualiza valor_caption com dado formatado a direita

End Sub
Sub LeituraDoctosEnvelope()
                   
Dim Count_Lista As Long
Dim Count_Reg   As Long
Dim Count_linha As Long
Dim Count_Grade As Long

Dim Duplicidade      As String
Dim RetornoTransacao As String
        
    On Error GoTo ERRO_LEITURADOCTOS
        
    If indice_corrente = 6 Then
       lblExclusao.Visible = True
       lblExclusao.Caption = IIf(IsNull(tbenv!Descricao), "", tbenv!Descricao)
    Else
       lblExclusao.Visible = False
    End If
        
    'não imprime este dado na tela quando for devolvido pelo robo
    Count_Grade = tbenv.RowCount + 1
        
       'Limita retorno de linhas em dez mil
       If Count_Grade >= 10000 Then
          Grade.Rows = 10000
       Else
          Grade.Rows = Count_Grade
       End If
        
        For Count_Reg = 0 To Count_Grade - 2
            
            Count_linha = Count_Reg + 1
            Grade.Row = Count_linha
            
            FormataValor Trim(Format(tbenv!Valor, "##,##0.00"))
            Grade.TextMatrix(Grade.Row, 0) = IIf(IsNull(tbenv!Vinculo), "0000000000", Format(tbenv!Vinculo, "0000000000"))
            If tbenv!TipoDocto = "1" Then
                Grade.TextMatrix(Grade.Row, 1) = IIf(IsNull(tbenv!StatusCapa), "0", (tbenv!StatusCapa))
                Call GravaLog(tbenv!IdCapa, 0, 140)
            Else
                Grade.TextMatrix(Grade.Row, 1) = IIf(IsNull(tbenv!StatusDocto), "0", (tbenv!StatusDocto))
            End If
            Grade.TextMatrix(Grade.Row, 2) = IIf(IsNull(tbenv!Autenticado), " ", tbenv!Autenticado)
            
            '--Trata Retorno Transacao
             If IsNull(tbenv!RetornoTransacao) = True Then
                RetornoTransacao = "00"
             ElseIf tbenv!RetornoTransacao = 0 Then
                RetornoTransacao = "00"
             Else
                RetornoTransacao = tbenv!RetornoTransacao
             End If
             
            Grade.TextMatrix(Grade.Row, 3) = IIf(IsNull(tbenv!Ocorrencia), "000", Mid(tbenv!Ocorrencia, 1, 3)) & RetornoTransacao
            Grade.TextMatrix(Grade.Row, 7) = IIf(IsNull(tbenv!AgOrig), "0", tbenv!AgOrig)
            Grade.TextMatrix(Grade.Row, 8) = IIf(IsNull(tbenv!Frente), "", tbenv!Frente)
            Grade.TextMatrix(Grade.Row, 9) = IIf(IsNull(tbenv!Verso), "", tbenv!Verso)
            Grade.TextMatrix(Grade.Row, 10) = IIf(IsNull(tbenv!IdDocto), "0000000000", Format(tbenv!IdDocto, "0000000000"))
            Grade.TextMatrix(Grade.Row, 12) = IIf(IsNull(tbenv!Ordem), "0", (tbenv!Ordem))
            Grade.TextMatrix(Grade.Row, 11) = IIf(IsNull(tbenv!TipoDocto), "", tbenv!TipoDocto)
            Grade.TextMatrix(Grade.Row, 14) = IIf(IsNull(tbenv!Ocorrencia), "000", Format(Mid(tbenv!Ocorrencia, 1, 3), "000")) & Format(RetornoTransacao, "00")
            Grade.TextMatrix(Grade.Row, 13) = IIf(IsNull(tbenv!IdCapa), "000000", tbenv!IdCapa)
            Grade.TextMatrix(Grade.Row, 15) = IIf(IsNull(tbenv!NSU), "000000", Format(tbenv!NSU, "######"))
            Grade.TextMatrix(Grade.Row, 16) = IIf(IsNull(tbenv!Terminal), "000", Format(tbenv!Terminal, "000"))
            Grade.TextMatrix(Grade.Row, 17) = IIf(IsNull(tbenv!Leitura), "000", tbenv!Leitura)
            Grade.TextMatrix(Grade.Row, 18) = IIf(IsNull(tbenv!Cortado), "", "S")
            Grade.TextMatrix(Grade.Row, 19) = IIf(IsNull(tbenv!IdLote), "", tbenv!IdLote)
            Grade.TextMatrix(Grade.Row, 20) = IIf(IsNull(tbenv!StatusCapa), "0", tbenv!StatusCapa)
            Grade.TextMatrix(Grade.Row, 21) = Trim(valor_caption)
            Grade.TextMatrix(Grade.Row, 22) = IIf(IsNull(tbenv!RecepcionadoIK), "0", tbenv!RecepcionadoIK)
            
            'Verifica Status de Alçada
            If IsNull(tbenv!Alcada) Or (tbenv!Alcada) = "N" Then
                Grade.TextMatrix(Grade.Row, 4) = ""
            Else
                Grade.TextMatrix(Grade.Row, 4) = tbenv!Alcada
            End If
            
            'Verifica Documentos em duplicidade
            If (IsNull(tbenv!Duplicidade)) Or (tbenv!Duplicidade) = 0 Then
                Grade.TextMatrix(Grade.Row, 5) = " "
            ElseIf (tbenv!Duplicidade) = 1 Then
                Grade.TextMatrix(Grade.Row, 5) = "S"
            End If
            
            'Verifica tipos  de documentos
            If (Grade.TextMatrix(Grade.Row, 11)) = "" Then
                TipoDocto = 0
            Else
                TipoDocto = Grade.TextMatrix(Grade.Row, 11)
            End If
            
            If TipoDocto = 0 Then
               TipoDocto = 0
               Call DescTipoProd
            Else
               TipoDocto = (Grade.TextMatrix(Grade.Row, 11))
               Call DescTipoProd
            End If
            
            Grade.TextMatrix(Grade.Row, 6) = DescTipoProd
            
        tbenv.MoveNext
        Next
    
    
   'Grade.RowSel = 1
   If CLng(nDocto_Sel) <> 0 Then
        
        'Posiciona no documento escolhido na pesquisa
        For Count_Lista = 1 To Grade.Rows - 1
            If Val(Grade.TextMatrix(Count_Lista, 10)) = CLng(nDocto_Sel) Then
                If Count_Lista = 1 Then
                    Grade.RowSel = 0
                    Grade.RowSel = 1
                Else
                    Grade.RowSel = Count_Lista
                    Grade.TopRow = Grade.Row
                    Grade.Col = 0
                    Grade.ColSel = 21
                    Grade.SetFocus
                End If
                Exit Sub
            End If
        Next Count_Lista
   Else
    Grade.Row = 1
    Grade.RowSel = 1
    Grade.TopRow = 1
    Grade.Col = 0
    Grade.ColSel = 21
    Call Grade_SelChange
    Grade.SetFocus
   End If
    
   SSTab1.Tab = 1
   SSTab1.TabEnabled(1) = True
   
Exit Sub
ERRO_LEITURADOCTOS:

    Select Case TratamentoErro("Não foi possível consultar os dados.", Err, rdoErrors)
           Case vbCancel
             Unload Me
           Case vbRetry
           Resume
    End Select

End Sub
Sub LimpaCampos()
   
    'limpa campos DÉBITO / DEPÓSITO
    TxtAg = ""
    TxtCta = ""
    TxtNSU = ""
    TxtChq = ""
    TxtBco = ""
    TxtNumMalote = ""
    txtcapa = ""
       
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If SSTab1.Tab = 0 Then Exit Sub
End Sub
Private Sub Imprime_Detalhe_Click()
       
On Error GoTo ERRO_IMPRESSAO

    Dim Arq, img_verso, NomEnvMal As String
    Dim len_img As Integer, rt_f As Long, rt_v As Long, Count_detalhe As Long
    Dim detalhe As String * 63, Tam As Integer, p As Integer
    
    With Modulo.qryGetIdDocto
         Pegiddocto = CLng(Grade.TextMatrix(Grade.Row, 10))
         .rdoParameters(0).Value = Geral.DataProcessamento
         .rdoParameters(1).Value = Pegiddocto
         Set tbenv1 = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
       
    ' Query que traz os Dados de Bco + Ag + Cc do Documento
    p = 131
    len_img = 0
    img_verso = Grade.TextMatrix(Grade.Row, 9)
    
    If Grade.TextMatrix(Grade.Row, 8) = "DEBITO" Then
        MsgBox "Esta imagem não pode ser impressa.", vbInformation, App.Title
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    
    len_img = Len(reg_imagem)
    If (len_img = 0) Then
        Screen.MousePointer = 0
        MsgBox "É preciso selecionar um documento para mostrar a imagem.", vbInformation, App.Title
        Exit Sub
    End If
   
    'Verifica se Imagem da CANON
    If Mid(Grade.TextMatrix(Grade.Row, 12), 1, 1) = 1 Then
        ' Imprime Frente e Verso do Cheque
        Arq = Dir$("IMGIF.BMP")
        If Arq = "IMGIF.BMP" Then Kill "IMGIF.BMP"
    
        Arq = Dir$("IMGIV.BMP")
        If Arq = "IMGIV.BMP" Then Kill "IMGIV.BMP"
        
        Arq = Dir$("IMGIF.TIF")
        If Arq = "IMGIF.TIF" Then Kill "IMGIF.TIF"
    
        Arq = Dir$("IMGIV.TIF")
        If Arq = "IMGIV.TIF" Then Kill "IMGIV.TIF"
    End If
    
    'Verifica a origem da imagem
    '0 - vips (.jpg)
    '1 - canon (.tif)
    '2 - ls500 (.bmp)
    If Mid(Grade.TextMatrix(Grade.Row, 12), 1, 1) = 0 Or (Mid(Grade.TextMatrix(Grade.Row, 12), 1, 1) = 2) Then
        'frente
'        If Geral.VIPSDLL = eDllProservi Then
'          Lead2.Load  Geral.DiretorioImagens & reg_imagem, 0, 0, 1
'        Else
'          Lead2.Load Geral.DiretorioImagens & Format(Val(Lote), "000000000") & "\" & reg_imagem, 0, 0, 1
'        End If

'        Lead2.PaintSizeMode = PAINTSIZEMODE_ZOOM
'        Lead2.PaintZoomFactor = Lead1.PaintZoomFactor - 70
'        Lead2.Save Geral.DiretorioImagens & "IMGIF.BMP", FILE_BMP, 1, 1, False
        
        'verso
'        If Geral.VIPSDLL = eDllProservi Then
'          Lead2.Load Geral.DiretorioImagens & Grade.TextMatrix(Grade.Row, 9), 0, 0, 1
'        Else
'          Lead2.Load Geral.DiretorioImagens & Format(Val(Lote), "000000000") & "\" & Grade.TextMatrix(Grade.Row, 9), 0, 0, 1
'        End If
'        Lead2.PaintSizeMode = PAINTSIZEMODE_ZOOM
'        Lead2.PaintZoomFactor = Lead1.PaintZoomFactor - 50
'        Lead2.Save Geral.DiretorioImagens & "IMGIV.BMP", FILE_BMP, 1, 1, False
        
'        Picture2.Picture = LoadPicture(Geral.DiretorioImagens & "IMGIF.BMP")
'        Picture3.Picture = LoadPicture(Geral.DiretorioImagens & "IMGIV.BMP")

        'frente
        If Geral.VIPSDLL = eDllProservi Then
          Picture2.Picture = LoadPicture(Geral.DiretorioImagens & reg_imagem)
        Else
          Picture2.Picture = LoadPicture(Geral.DiretorioImagens & Format(Val(Lote), "000000000") & "\" & reg_imagem)
        End If

        'verso
        If Geral.VIPSDLL = eDllProservi Then
          Picture3.Picture = LoadPicture(Geral.DiretorioImagens & Grade.TextMatrix(Grade.Row, 9))
        Else
          Picture3.Picture = LoadPicture(Geral.DiretorioImagens & Format(Val(Lote), "000000000") & "\" & Grade.TextMatrix(Grade.Row, 9))
        End If

'        Picture2.Picture = LoadPicture(Geral.DiretorioImagens & "IMGIF.BMP")
'        Picture3.Picture = LoadPicture(Geral.DiretorioImagens & "IMGIV.BMP")
    
    Else 'se canon - tif
         'frente
        If Geral.VIPSDLL = eDllProservi Then
          Lead2.Load Geral.DiretorioImagens & reg_imagem, 0, 0, 1
        Else
          Lead2.Load Geral.DiretorioImagens & Format(Val(Lote), "000000000") & "\" & reg_imagem, 0, 0, 1
        End If
        
        Lead2.PaintSizeMode = PAINTSIZEMODE_ZOOM
        Lead2.PaintZoomFactor = Lead1.PaintZoomFactor - 50
        Lead2.Save Geral.DiretorioImagens & "IMGIF.BMP", FILE_BMP, 1, 1, False
        Picture2.Picture = LoadPicture(Geral.DiretorioImagens & "IMGIF.BMP")
    End If
    
    ScaleHeight = 1000   ' Set height units.
    ScaleWidth = 300
    Printer.ScaleMode = 6
    Printer.FontSize = 8
    Printer.Font = "Courier New"

    Dim Dia As String, Hora As String, Capa As String
    
    Dia = Format$(Now, " DD/MM/YYYY ")
    Hora = Format$(Now, " HH:MM:SS ")
    
    Printer.Print "Multi-Agência  - " & Geral.AgenciaCentral
    Printer.Print

    Printer.Print App.Title; "  "; ; ; Dia; "-"; Hora
    Printer.Print
    
    If NEnvMal = "M" Then
        NomEnvMal = "Malote"
        Capa = Format(tbenv1!Capa, "00000000000000")
    Else
        NomEnvMal = "Envelope"
        Capa = Format(tbenv1!Capa, "00000000")
    End If
    
    Printer.Print "Documento Processado no movimento de "; Mid(Geral.DataProcessamento, 7, 2); "/"; Mid(Geral.DataProcessamento, 5, 2); "/"; Mid(Geral.DataProcessamento, 1, 4); " - no "; NomEnvMal; ":"; "  "; Capa
    Printer.Print
    Printer.Print
   
    If tbenv1.EOF Then
        Screen.MousePointer = 0   'default
        MsgBox "Não foi possível localizar os detalhes do documento para impressão. Tente novamente.", vbInformation, App.Title
        Exit Sub
    Else
        If tbenv1!TipoDocto <> 1 Then
            With Modulo.qryGetAgCcDoc
                .rdoParameters(0).Value = Geral.DataProcessamento
                .rdoParameters(1).Value = Pegiddocto
                .rdoParameters(2).Value = tbenv1!TipoDocto
                Set RsAux = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
            End With
            
            MontaDetalheConsulta (CInt(Grade.TextMatrix(Grade.Row, 11)))
            
            For Count_detalhe = 0 To List_detalhe.ListCount - 1
                FlagImp = 1
                List_detalhe.Selected(Count_detalhe) = True
                Printer.Print List_detalhe.Text
            Next
                List_detalhe.Selected(Count_detalhe - 1) = False
                FlagImp = 0
        End If
    End If
        
     Printer.PaintPicture Picture2, 5, 60
     
    'so imprime o verso quando for da Vips ou LS500
    If Mid(Grade.TextMatrix(Grade.Row, 12), 1, 1) <> 1 Then
       Printer.PaintPicture Picture3, 5, 165
    '140 / 190
    End If
    
    Printer.EndDoc
    Screen.MousePointer = 0   'default
    Grade.SetFocus
    
    Exit Sub
    
ERRO_IMPRESSAO:
    Screen.MousePointer = vbDefault
    Printer.KillDoc
    
    Select Case TratamentoErro("Não foi possível fazer a impressão.", Err, rdoErrors)
        Case vbCancel
            Unload Me
        Case vbRetry
            Resume
    End Select
End Sub
Private Sub cmbTipoDocto_KeyPress(KeyAscii As Integer)

    If (KeyAscii = 13) Then
        TxtVal.Enabled = True
        TxtVal.SetFocus
    ElseIf (KeyAscii = 27) Then
        KeyAscii = 0
        CmdSair_Click (0)
    End If
    
End Sub
Private Sub cmdConfirma_Click()
 
' 1 --- Tratamento para pesquisa de Capa Envelope e Malote
' 2 --- Tratamento da pesquisa por Banco + Agencia + Conta + Cheque
' 3 --- Tratamento da pesquisa por Valor
' 4 --- Tratamento da pesquisa por NSU
' 5 --- Tratamento da Pesquisa por Número de Malote
' 6 --- Tratamento da Pesquisa Tipo de Documento
' 7 --- Faz a pesquisa de cada opção
                 
   If (indice_corrente = 1) Then
        If (Len(txtcapa) = 0) Then
            Screen.MousePointer = 0
            MsgBox "A digitação do campo Envelope é obrigatória !", vbInformation, App.Title
            txtcapa.SetFocus
            Exit Sub
        End If
        
    ElseIf (indice_corrente = 2) Then
        If Len(Trim(TxtBco)) <> 0 Or Len(Trim(TxtAg)) <> 0 Or Len(Trim(TxtCta)) <> 0 Or Len(Trim(TxtChq)) <> 0 Then
        Else
            Screen.MousePointer = 0
            MsgBox "É obrigatório o preenchimento de pelo menos um campo.", vbInformation, App.Title
            TxtBco.SetFocus
            Exit Sub
        End If
        
    ElseIf (indice_corrente = 3) Then
        If (Val(TxtVal.Text) = 0) Then
            Screen.MousePointer = 0
            MsgBox "É obrigatório o preenchimento deste campo.", vbInformation, App.Title
            TxtVal.SetFocus
            Exit Sub
        End If
    
    ElseIf (indice_corrente = 4) Then
        If (Len(TxtNSU) = 0) Then
            Screen.MousePointer = 0
            MsgBox "É obrigatório o preenchimento deste campo.", vbInformation, App.Title
            TxtNSU.SetFocus
            Exit Sub
        End If
    
    ElseIf (indice_corrente = 5) Then
        If (Len(TxtNumMalote) = 0) Then
            Screen.MousePointer = 0
            MsgBox "É obrigatório o preenchimento deste campo.", vbInformation, App.Title
            TxtNumMalote.SetFocus
            Exit Sub
        End If
    
    ElseIf (indice_corrente = 6) Then
        If Not OptOcCapa.Value = True Then
            If Not OptOcDocto.Value = True Then
                Screen.MousePointer = 0
                MsgBox "É obrigatório o preenchimento deste campo.", vbInformation, App.Title
                cmbTipoDocto.SetFocus
                Exit Sub
            Else
                NCapaDocto = "Doc"
            End If
        Else
            NCapaDocto = "Cap"
        End If
    
    End If
    
    Call FazPesquisa
    
End Sub
Private Sub FazPesquisa()

    On Error GoTo ERRO_FAZPESQUISA
    
    If (indice_corrente = 1) Then

        Call PesqCapaEnvMal

    ElseIf (indice_corrente = 2) Then

        Call PesqCheques
        
    ElseIf (indice_corrente = 3) Then
        
        Call PesqProduto
        Call PesqValores

    ElseIf (indice_corrente = 4) Then

        Call PesqDoctoNSU

    ElseIf (indice_corrente = 5) Then

        Call PesqNumMalote

    ElseIf (indice_corrente = 6) Then

        Call PesqDoctoOcorrencias

    End If
    
Exit Sub
ERRO_FAZPESQUISA:
    
    Select Case TratamentoErro("Não foi possível consultar os dados.", Err, rdoErrors)
        Case vbCancel
            Unload Me
        Case vbRetry
            Resume
    End Select

End Sub
Private Sub cmdFrenteVerso_Click()
    
'verifica se as imagens são da Vips ou da ls500, pois as imagens
'do canon não tem verso

    If (Mid(Grade.TextMatrix(Grade.Row, 12), 1, 1) = "0") Or (Mid(Grade.TextMatrix(Grade.Row, 12), 1, 1) = "2") Then
        If Lead1.Tag = "V" Then
           'se verso, mostrar frente
            Lead1.Tag = "F"
            With Lead1
                .AutoRepaint = False
                If Geral.VIPSDLL = eDllProservi Then
                    .Load Geral.DiretorioImagens & reg_imagem, 0, 0, 1
                Else
                    .Load Geral.DiretorioImagens & Format(Val(Lote), "000000000") & "\" & reg_imagem, 0, 0, 1
                End If
                
                'se ls500 mostrar mais escuro
                If (Mid(Grade.TextMatrix(Grade.Row, 12), 1, 1) = "2") Then
                   .Intensity 140
                Else
                   .Intensity 220
                End If
                .PaintZoomFactor = 100
                .AutoRepaint = True
            End With
        Else
            'se frente, mostrar verso
            Lead1.Tag = "V"
            With Lead1
                .AutoRepaint = False
                If Geral.VIPSDLL = eDllProservi Then
                    .Load Geral.DiretorioImagens & Grade.TextMatrix(Grade.Row, 9), 0, 0, 1
                Else
                    .Load Geral.DiretorioImagens & Format(Val(Lote), "000000000") & "\" & Grade.TextMatrix(Grade.Row, 9), 0, 0, 1
                End If
                
                'se ls500 mostrar mais escuro
                If (Mid(Grade.TextMatrix(Grade.Row, 12), 1, 1) = "2") Then
                   .Intensity 140
                Else
                   .Intensity 220
                End If
                .PaintZoomFactor = 100
                .AutoRepaint = True
            End With
        End If
    End If

End Sub
Private Sub cmdInverteCor_Click()
    Lead1.Invert
End Sub
Private Sub cmdRotacao_Click()
    Lead1.FastRotate 90
End Sub
Private Sub CmdSair_Click(Index As Integer)
    Unload Me
End Sub
Private Sub cmdZoomMais_Click()
    If Lead1.PaintZoomFactor <= 400 Then
        Lead1.PaintSizeMode = PAINTSIZEMODE_ZOOM
        Lead1.PaintZoomFactor = Lead1.PaintZoomFactor + 10
    End If
End Sub
Private Sub cmdZoomMenos_Click()
    If Lead1.PaintZoomFactor >= 20 Then
       Lead1.PaintSizeMode = PAINTSIZEMODE_ZOOM
       Lead1.PaintZoomFactor = Lead1.PaintZoomFactor - 10
    End If
End Sub
Private Sub Command1_Click()
    Unload Me
End Sub
Private Sub Form_Activate()

   'Inclusão de chamada a rotina AtualizaAtividade
   Call AtualizaAtividade(24)

With Grade
    
    .Rows = 1
    .Cols = 23
    .Row = 0
      
    .Col = 0
    .ColWidth(0) = Grade.Width * 0.12
    .ColAlignment(0) = 3
    .Text = "Vinculo"
    
    .Col = 1
    .ColWidth(1) = Grade.Width * 0.06
    .ColAlignment(1) = 3
    .Text = "Status"
    
    .Col = 2
    .ColWidth(2) = Grade.Width * 0.09
    .ColAlignment(2) = 3
    .Text = "Autenticado"
    
    .Col = 3
    .ColWidth(3) = Grade.Width * 0.09
    .ColAlignment(3) = 3
    .Text = "Ocorrência"
    
    .Col = 4
    .ColWidth(4) = Grade.Width * 0.08
    .ColAlignment(4) = 3
    .Text = "Alçada"
    
    .Col = 5
    .ColWidth(5) = Grade.Width * 0.08
    .ColAlignment(5) = 3
    .Text = "Duplicado"
    
    .Col = 6
    .ColWidth(6) = Grade.Width * 0.2
    .ColAlignment(6) = 1
    .Text = "             Tipo de Documento"
    
    .Col = 7
    .ColWidth(7) = Grade.Width * 0.1
    .ColAlignment(7) = 3
    .Text = "AG. Origem"
    
    .Col = 21
    .ColWidth(21) = Grade.Width * 0.14
    .Text = "                Valor"
    
    .ColWidth(8) = Grade.Width * 0  ' Frente da Imagem
    .ColWidth(9) = Grade.Width * 0  ' Verso  da Imagem
    .ColWidth(10) = Grade.Width * 0 ' Iddocto   - Identificação do documento
    .ColWidth(11) = Grade.Width * 0 ' TipoDocto - Tipo de documento
    .ColWidth(12) = Grade.Width * 0 ' Ordem
    .ColWidth(13) = Grade.Width * 0 ' IdCapa
    .ColWidth(14) = Grade.Width * 0 ' Código da ocorrência
    .ColWidth(15) = Grade.Width * 0 ' NSU
    .ColWidth(16) = Grade.Width * 0 ' Terminal
    .ColWidth(17) = Grade.Width * 0 ' Leitura
    .ColWidth(18) = Grade.Width * 0 ' Cortado
    .ColWidth(19) = Grade.Width * 0 ' IdLote
    .ColWidth(20) = Grade.Width * 0 ' StatusCapa
    .ColWidth(22) = Grade.Width * 0 ' Recepcionado IK

End With

txtcapa.SetFocus
    
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
 
  Select Case KeyCode
    Case vbKeyAdd
        If cmdZoomMais.Enabled = False Then Exit Sub
        cmdZoomMais_Click
    Case vbKeySubtract
        If cmdZoomMenos.Enabled = False Then Exit Sub
        cmdZoomMenos_Click
    Case vbKeyF10
        If cmdZoomMenos.Enabled = False Then Exit Sub
        cmdInverteCor_Click
    Case vbKeyDivide
        If cmdRotacao.Enabled = False Then Exit Sub
        cmdRotacao_Click
    Case vbKeyF11
        If cmdFrenteVerso.Enabled = False Then Exit Sub
        cmdFrenteVerso_Click
    Case vbKeyUp, vbKeyDown
    
    Case 27
        Call CmdSair_Click(1)
        
  End Select
  
  KeyCode = 0
  
End Sub
Private Sub Form_Load()

    With Lead1
        .UnlockSupport L_SUPPORT_EXPRESS, "YXPQ3XPPVT"
        .UnlockSupport L_SUPPORT_GIFLZW, "0K3RV9UY3EY"
        .UnlockSupport L_SUPPORT_TIFLZW, "9LE75L0FDXHK"
    End With

    ' Cria query para a leitura de Status da Capa
    Set Modulo.qryGetStatusCapa = Geral.Banco.CreateQuery("", "{call getAgenciasCapa(?,?,?,?)}")
    
    ' Cria query para leitura de Oct
    Set Modulo.qryGetOCT = Geral.Banco.CreateQuery("", "{call GetOct(?,?)}")

    ' Cria query para a leitura de todos os Dados do Documento (Capa)
    Set Modulo.qryGetIdcDocto = Geral.Banco.CreateQuery("", "{call GetIdCapa(?,?)}")
    
    ' Cria query para a leitura de todos os Dados do Documento (Malote)
    Set Modulo.qryGetIdDocto = Geral.Banco.CreateQuery("", "{call GetIdDocto(?,?)}")

    ' Cria query para a leitura de todos os Dados do Documento (IdCapa)
    Set Modulo.qryGetIdCapa = Geral.Banco.CreateQuery("", "{call GetIdDocCapa (?,?)}")

    ' Cria query para a leitura de envelope que contém o cheque consultado '
    Set Modulo.qryBuscaEnvelope = Geral.Banco.CreateQuery("", "{call GetDocumentosNumenv(?,?,?)}")

    ' Cria query para a leitura do docto com bco+ag+cta+chq '
    Set Modulo.qryPesquisaBcoAgCtaChq = Geral.Banco.CreateQuery("", "{? = call getbcoagcccheque(?,?,?,?,?,?)}")

    ' Cria query para a leitura de Tipo de Ocorrência
    Set Modulo.qryGetTipoOcorr = Geral.Banco.CreateQuery("", "{ Call GetTipoOcorr(?,?,?)}")

    ' Cria query para a leitura do docto com valor desejado '
    Set Modulo.qryPesquisaValor = Geral.Banco.CreateQuery("", "{? = call Getdocumentosvalor  (?,?,?,?)}")

    ' Cria query para a leitura do doctos através do Numero de Malote
    Set Modulo.qryGetDoctosNumMalote = Geral.Banco.CreateQuery("", "{call GetDocumentosNumMalote  (?,?)}")

    ' Cria query para a leitura do docto com NSU desejado '
    Set Modulo.qryGetDocumentosCapaNSU = Geral.Banco.CreateQuery("", "{call GetDocumentosCapaNSU(?,?,?)}")

    ' Cria query para a leitura do docto com Numero da Capa '
    Set Modulo.qryGetDocumentosNumCapa = Geral.Banco.CreateQuery("", "{? = call GetDocumentosNumCapa(?,?,?)}")

    ' Cria query para a leitura do documento escolhido '
    Set Modulo.qryDocumentoEscolhido = Geral.Banco.CreateQuery("", "{call GetDocto(?,?,?)}")

    ' Cria query para a leitura da Ag. do Docto Escolhido'
    Set Modulo.qryGetAgCcDoc = Geral.Banco.CreateQuery("", "{call GetAgContaDocumento (?,?,?)}")

    ' Cria query para a leitura de Ocorrencias'
    Set qryGetocorrencia = Geral.Banco.CreateQuery("", "{Call GetOcorrencia (?)}")

    ' Cria query para a Leitura dos detalhes do Depósito'
    Set Modulo.qryGetDeposito = Geral.Banco.CreateQuery("", "{call GetDeposito(?,?)}")
    
    ' Cria query para a Leitura dos detalhes da Cobranca'
    Set Modulo.qryGetCobCodBar = Geral.Banco.CreateQuery("", "{call GetCobrancaCodBar(?,?)}")
    
    ' Cria query para a Leitura dos detalhes do Darf-Simples'
    Set Modulo.qryGetDARFSimples = Geral.Banco.CreateQuery("", "{call GetDarfSimples(?,?)}")
    
    ' Cria query para a Leitura dos detalhes da Arrecadação Convencional'
    Set Modulo.qryGetArrecConvenc = Geral.Banco.CreateQuery("", "{call GetArrecConv(?,?)}")
    
    ' Cria query para a Leitura dos detalhes do Título'
    Set Modulo.qryGetTitulo = Geral.Banco.CreateQuery("", "{call GetTitulo(?,?)}")
    
    ' Cria query para a Leitura dos detalhes do ADCC'
    Set Modulo.qryGetADCC = Geral.Banco.CreateQuery("", "{call GetADCC(?,?)}")
    
    ' Cria query para a Leitura dos detalhes do Darf-Preto'
    Set Modulo.qryGetDARFPreto = Geral.Banco.CreateQuery("", "{call GetDarfPreto(?,?)}")
    
    ' Cria query para a Leitura dos detalhes do Gare'
    Set Modulo.qryGetGare = Geral.Banco.CreateQuery("", "{call GetGare(?,?)}")
    
    ' Cria query para a Leitura dos detalhes da Cobrança Especial'
    Set Modulo.qryGetCobEspTec = Geral.Banco.CreateQuery("", "{call GetCobrancaEspecial(?,?)}")
    
    ' Cria query para a Leitura dos detalhes do GPS'
    Set Modulo.qryGetGps = Geral.Banco.CreateQuery("", "{call GetGps(?,?)}")
    
    ' Cria query para a Leitura dos detalhes do DARM'
    Set Modulo.qryGetDarm = Geral.Banco.CreateQuery("", "{call GetDarm(?,?)}")
    
    ' Cria query para a Leitura dos detalhes do Cartão Avulso'
    Set Modulo.qryCartaoAvulso = Geral.Banco.CreateQuery("", "{call GetCartaoAvulso(?,?)}")
    
    ' Cria query para a Leitura dos detalhes da Cobrança Registrada'
    Set Modulo.qryGetCobRegTec = Geral.Banco.CreateQuery("", "{call GetCobrancaRegistrada(?,?)}")
    
    ' Cria query para a Leitura de descrição dos tipo de documentos'
    Set Modulo.qryGettipoDocto = Geral.Banco.CreateQuery("", "{call Gettipodocto (?)}")
    
    ' Cria query para a Leitura da descrição do Status da Capa
    Set Modulo.qryGetConsStatusCapa = Geral.Banco.CreateQuery("", "{call GetConsStatusCapa (?)}")
   
   ' Cria query para a Leitura de Dados do FGTS
    Set Modulo.qryGetFGTS = Geral.Banco.CreateQuery("", "{call GetFGTS (?,?)}")
    
    ' Cria query para a Leitura de Dados do Lançamento Interno
    Set Modulo.qryGetLctoInterno = Geral.Banco.CreateQuery("", "{call GetLancamentoInterno (?,?)}")
    
    ' Cria query para a Leitura do Motivo de Exclusão para capa Selecionada
    Set Modulo.qryGetMotivoExclusao = Geral.Banco.CreateQuery("", "{call GetMotivoExclusao(?,?)}")
    

    reg_imagem = ""             'nome da imagem frente
    reg_frente = ""             'nome da imagem frente
    indice_corrente = 1         'default - opção envelope
    txtcapa = ""                'default - txtcapa = ""
    visual = "F"                'default - opção envelope
    NEnvMal = "E"               'default - Identificação de Capa
    
    ' Valor default da Tab - Opçes
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = False
    SSTab1.Tab = 0
    
    OptOcorrencia.Enabled = False
    Call TipoDoctos
    cmbTipoDocto.Enabled = False

End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    With Modulo
    
        .qryGetDocumentosNumCapa.Close
        .qryGetAgCcDoc.Close
        .qryGetStatusCapa.Close
        .qryGetDocumentosCapaNSU.Close
        .qryPesquisaBcoAgCtaChq.Close
        .qryPesquisaValor.Close
        .qryBuscaEnvelope.Close
        .qryDocumentoEscolhido.Close
        .qryGetADCC.Close
        .qryGetDeposito.Close
        .qryGetArrecConvenc.Close
        .qryGetDARFPreto.Close
        .qryGetDARFSimples.Close
        .qryGetGare.Close
        .qryGetGps.Close
        .qryGetDarm.Close
        .qryGetTitulo.Close
        .qryGetCobEspTec.Close
        .qryGetCobRegTec.Close
        .qryGetCobCodBar.Close
        .qryCartaoAvulso.Close
        .qryGetDoctosNumMalote.Close
        .qryGetTipoOcorr.Close
        .qryGetIdcDocto.Close
        .qryGetIdDocto.Close
        .qryGetIdCapa.Close
        .qryGetOCT.Close
        .qryGetConsStatusCapa.Close
        .qryGetMotivoExclusao.Close
        .qryGetFGTS.Close
        .qryGetLctoInterno.Close
        
    End With
    
End Sub
Private Sub Grade_SelChange()
'* Seleciona Lina Inteira *'
    
    Static m_SelChange As Boolean
    
    'Controle de Acesso
    If m_SelChange = True Then Exit Sub
    
    m_SelChange = True
        
        Grade.Row = Grade.RowSel
        Grade.Col = 0
        Grade.ColSel = 21
        Grade.SetFocus
        Call Mostra_Imagem
    
    m_SelChange = False
    
    
End Sub
Private Sub Label_Envelope_Change()

    'Formata o Valor Capa de Malote / Envelope
     If NEnvMal = "E" Then
        If Len(Label_Envelope) <> 0 Then
            Label_Envelope = Format(Label_Envelope, "000000000")
        End If
     Else
        If Len(Label_Envelope) <> 0 Then
            Label_Envelope = Format(Label_Envelope, "00000000000000")
        End If
     End If
 
End Sub
Private Sub Lead1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error GoTo ERRO_LEADMOUSEDOWN

    'Lead1.MouseIcon = LoadPicture("mao2.cur")
    Lead1.MousePointer = 99
    Xold = X
    Yold = Y
    IsMove = True
    Lead1.AutoRepaint = False
    Atualiza = 2
    Exit Sub
    
ERRO_LEADMOUSEDOWN:
MsgBox "(OEGT) Não foi possível mostrar o detalhe do documento - arq:mao.cur.", vbInformation, App.Title

End Sub
Private Sub Lead1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo Saida
    
    Dim PontosX, PontosY, Zoom As Integer, MovimentoX, MovimentoY As Integer
    Dim Largura, Altura As Integer, Ret As Integer
    
    PontosX = Screen.TwipsPerPixelX
    PontosY = Screen.TwipsPerPixelY
    
    ' Se não está Clicado, cai fora
    If IsMove = False Then
        Exit Sub
    End If
    
    ' Na primeira vez está Empty
    If (Xold = Empty) Or (Yold = Empty) Then
        Xold = X
        Yold = Y
        Exit Sub
    End If
    
    ' Se não houve movimento, cai fora
    If (X = Xold) And (Y = Yold) Then
        Exit Sub
    End If
    
    Zoom = Lead1.PaintZoomFactor / 100
    
    MovimentoX = (Xold - X) / PontosX
    MovimentoY = (Yold - Y) / PontosY
    
    If Atualiza = 1 Then
        If MovimentoX <> 0 Then
            If MovimentoX > 0 Then
                Ret = SendMessage(hCtl, WM_HSCROLL, SB_LINEDOWN, 0)
            Else
                Ret = SendMessage(hCtl, WM_HSCROLL, SB_LINEUP, 0)
            End If
        End If
        If MovimentoY <> 0 Then
            If MovimentoY > 0 Then
                Ret = SendMessage(hCtl, WM_VSCROLL, SB_LINEDOWN, 0)
            Else
                Ret = SendMessage(hCtl, WM_VSCROLL, SB_LINEUP, 0)
            End If
        End If
    End If
    Atualiza = Atualiza + 1
    If Atualiza >= 5 Then
       Atualiza = 1
    End If
    
    Yold = Y
    Xold = X
Saida:

End Sub
Private Sub Lead1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Lead1.MousePointer = 99
    If IsMove = True Then
        Atualiza = 1
    End If
    IsMove = False
    Lead1.AutoRepaint = True
End Sub
Private Sub Opn_Banco_Click()
 ' Desabilita Objetos que não pertencem a Opção 2
 ' opção de pesquisa por Banco+Agencia+Conta+Cheque

   indice_corrente = 2
   LimpaTela Me
   nDocto_Sel = 0
   txtcapa.Enabled = False
   LblValor.Enabled = False
   TxtVal.Enabled = False
   LblNSU.Enabled = False
   TxtNSU.Enabled = False
   
   OptOcorrencia.Enabled = False
   OptOcCapa.Value = False
   OptOcDocto.Value = False
   
   TxtNumMalote.Enabled = False
   
   ' habilita chave Banco+Agencia+Conta+Cheque
   LblBanco.Enabled = True
   LblAgencia.Enabled = True
   LblCta.Enabled = True
   LblCheque.Enabled = True
   TxtBco.Enabled = True
   TxtAg.Enabled = True
   TxtCta.Enabled = True
   TxtChq.Enabled = True
   TxtBco.SetFocus
   
End Sub
Private Sub Opn_Capa_Click()
' Desabilita Campos que não pertencem a Opção 1
' Default - opção envelope
   
   indice_corrente = 1
   nDocto_Sel = 0
   
   SSTab1.Tab = 0
   LimpaTela Me
   
   Opn_Banco.Value = False
   Opn_NSU.Value = False
   Opn_NumMalote.Value = False
   opn_Ocorrencia.Value = False
   Opn_Vl.Value = False
   OptOcCapa.Value = False
   OptOcDocto.Value = False
   optTipo(0).Value = False
   optTipo(1).Value = False

   LblBanco.Enabled = False
   LblAgencia.Enabled = False
   LblCta.Enabled = False
   LblCheque.Enabled = False
   LblValor.Enabled = False
   TxtBco.Enabled = False
   TxtAg.Enabled = False
   TxtCta.Enabled = False
   TxtChq.Enabled = False
   TxtVal.Enabled = False
   OptOcorrencia.Enabled = False
   LblNSU.Enabled = False
   TxtNSU.Enabled = False
   cmbTipoDocto.Enabled = False
   
    Frame_Campos(1).Enabled = False
    Frame_Indices.Enabled = False
    
   ' Habilita opção Envelope
   txtcapa.Enabled = True
   'txtcapa.SetFocus
   txtcapa.MaxLength = 14
   txtcapa.SetFocus
   
End Sub
Private Sub Opn_NSU_Click()
      
 'Desabilita Objetos que não pertencem a Opção 4
 'opção NSU
   
   indice_corrente = 4
   
   LimpaTela Me
   nDocto_Sel = 0
   LblBanco.Enabled = False
   LblAgencia.Enabled = False
   LblCta.Enabled = False
   LblCheque.Enabled = False
   LblValor.Enabled = False
   TxtBco.Enabled = False
   TxtAg.Enabled = False
   TxtCta.Enabled = False
   TxtChq.Enabled = False
   txtcapa.Enabled = False
   TxtVal.Enabled = False
   
   OptOcorrencia.Enabled = False
   Opn_Capa.Value = False
   Opn_NumMalote.Value = False
   OptOcCapa.Value = False
   OptOcDocto.Value = False
   
   TxtNumMalote.Enabled = False
   
   LblNSU.Enabled = True
   TxtNSU.Enabled = True
       
   'foco para NSU
   TxtNSU.SetFocus
   
End Sub
Private Sub opn_Ocorrencia_Click()

' Desabilita Objetos que não pertencem a Opção 6
' opção de pesquisa por Ocorrência
 
   indice_corrente = 6

   LimpaTela Me
   nDocto_Sel = 0
   LblBanco.Enabled = False
   LblAgencia.Enabled = False
   LblCta.Enabled = False
   LblCheque.Enabled = False
   LblValor.Enabled = False
   TxtBco.Enabled = False
   TxtAg.Enabled = False
   TxtCta.Enabled = False
   TxtChq.Enabled = False
   txtcapa.Enabled = False
   TxtVal.Enabled = False
   LblNSU.Enabled = False
   TxtNSU.Enabled = False
   TxtNumMalote.Enabled = False
   cmbTipoDocto.Enabled = False
   LblValor.Enabled = False
   TxtVal.Enabled = False
   
   'Habilita opção Ocorrência
   OptOcorrencia.Enabled = True
   OptOcCapa.Value = False
   OptOcDocto.Value = False
   OptOcCapa.SetFocus
   Opn_Capa.Value = False
   Opn_NumMalote.Value = False
 
End Sub
Private Sub Opn_Vl_Click()
   
 ' Desabilita Objetos que não pertencem a Opção 3
 ' Opção valor = 3
   
   indice_corrente = 3
   nDocto_Sel = 0
   LimpaTela Me
   
   Opn_Capa.Value = False
   Opn_NumMalote.Value = False
   OptOcCapa.Value = False
   OptOcDocto.Value = False
   
   LblBanco.Enabled = False
   LblAgencia.Enabled = False
   LblCta.Enabled = False
   LblCheque.Enabled = False
   TxtBco.Enabled = False
   TxtAg.Enabled = False
   TxtCta.Enabled = False
   TxtChq.Enabled = False
   OptOcorrencia.Enabled = False
   LblNSU.Enabled = False
   TxtNSU.Enabled = False
   txtcapa.Enabled = False
   TxtNumMalote.Enabled = False
   
   'Habita Campos opção 3
   cmbTipoDocto.Enabled = True
   cmbTipoDocto.Text = "Todos"
   TxtVal.Enabled = True
   LblValor.Enabled = True
   TxtVal.SetFocus
   
End Sub
Private Sub Opn_NumMalote_Click()
 ' Desabilita Objetos que não pertencem a Opção 5
 ' opção Pesquisa por Numero de Malote
 
   indice_corrente = 5
 
   LimpaTela Me
   nDocto_Sel = 0
   LblBanco.Enabled = False
   LblAgencia.Enabled = False
   LblCta.Enabled = False
   LblCheque.Enabled = False
   LblValor.Enabled = False
   TxtBco.Enabled = False
   TxtAg.Enabled = False
   TxtCta.Enabled = False
   TxtChq.Enabled = False
   txtcapa.Enabled = False
   TxtVal.Enabled = False
   LblNSU.Enabled = False
   TxtNSU.Enabled = False
   LblValor.Enabled = False
   TxtVal.Enabled = False
   cmbTipoDocto.Enabled = False
   
   OptOcorrencia.Enabled = False
   Opn_Banco.Value = False
   Opn_NSU.Value = False
   Opn_Capa.Value = False
   opn_Ocorrencia.Value = False
   Opn_Vl.Value = False
   OptOcCapa.Value = False
   OptOcDocto.Value = False
   optTipo(0).Value = False
   optTipo(1).Value = False
   
   Frame_Campos(1).Enabled = False
   Frame_Indices.Enabled = False
    
   'Habilita opção Numero de Malote
   TxtNumMalote.Enabled = True
   
   'foco para Numero de Malote
   TxtNumMalote.SetFocus
   
End Sub
Private Sub optTipo_Click(Index As Integer)
           
nDocto_Sel = 0
Frame_Campos(1).Enabled = True
Frame_Indices.Enabled = True

    If optTipo(0).Value Then
        NEnvMal = "E"
        lblTipo.Caption = "Capa de Envelope"
        txtcapa.Text = ""
        TxtNumMalote.Text = ""
        Opn_Capa.Value = False
        Opn_NumMalote.Value = False
     Else
        NEnvMal = "M"
        lblTipo.Caption = "Capa de Malote"
        txtcapa.Text = ""
        TxtNumMalote.Text = ""
        Opn_Capa.Value = False
        Opn_NumMalote.Value = False
    End If
    
    cmbTipoDocto.Enabled = False
    
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
      
      SSTab1.TabEnabled(1) = False
      SSTab1.TabEnabled(0) = True
      
      If NEnvMal = "E" Then
          Lbl_tpCapa.Caption = "Capa Envelope"
      Else
          Lbl_tpCapa.Caption = "Capa Malote"
      End If
      
      If SSTab1.Tab = 0 Then
        LimpaTela Me
        'Opn_Capa_Click
        'Opn_Capa.SetFocus
        List_detalhe.Visible = False
        achou = 0
            
            If indice_corrente = 1 Then
              Call Opn_Capa_Click
            ElseIf indice_corrente = 2 Then
              Call Opn_Banco_Click
            ElseIf indice_corrente = 3 Then
              Call Opn_Vl_Click
            ElseIf indice_corrente = 4 Then
              Call Opn_NSU_Click
            ElseIf indice_corrente = 5 Then
              Call Opn_NumMalote_Click
            ElseIf indice_corrente = 6 Then
              Call opn_Ocorrencia_Click
            ElseIf indice_corrente = 7 Then
              Call Opn_Vl_Click
            End If

      End If
            
End Sub
Private Sub TxtAg_Change()

If Len(Trim(TxtAg.Text)) = 0 Then Exit Sub
If IsNumeric(TxtAg.Text) = False Then
    MsgBox "Valor inválido para este campo.", vbExclamation + vbOKOnly, App.Title
    TxtAg.Text = ""
    TxtAg.SetFocus
End If

End Sub

Private Sub TxtAg_KeyPress(KeyAscii As Integer)
    InibirTeclaAlfa KeyAscii
    
    If (KeyAscii = 13) Then
        TxtAg = Format(TxtAg, "0000")
        TxtCta.SetFocus
    ElseIf (KeyAscii = 27) Then
        KeyAscii = 0
        CmdSair_Click (0)
    End If
    
End Sub
Private Sub TxtBco_Change()

If Len(Trim(TxtBco.Text)) = 0 Then Exit Sub
If IsNumeric(TxtBco.Text) = False Then
    MsgBox "Valor inválido para este campo.", vbExclamation + vbOKOnly, App.Title
    TxtBco.Text = ""
    TxtBco.SetFocus
End If

End Sub

Private Sub TxtBco_KeyPress(KeyAscii As Integer)
    
    InibirTeclaAlfa KeyAscii
    
    If (KeyAscii = 13) Then
        If (Len(TxtBco) > 0) Then
           TxtBco = Format(TxtBco, "000")
        End If
        TxtAg.SetFocus
    ElseIf (KeyAscii = 27) Then
        KeyAscii = 0
        CmdSair_Click (0)
    End If
    
End Sub
Private Sub txtcapa_Change()

If Len(Trim(txtcapa.Text)) = 0 Then Exit Sub
If IsNumeric(txtcapa.Text) = False Then
    MsgBox "Valor inválido para este campo.", vbExclamation + vbOKOnly, App.Title
    txtcapa.Text = ""
    txtcapa.SetFocus
End If

End Sub
Private Sub txtCapa_GotFocus()

  txtcapa.SelStart = 0
  txtcapa.SelLength = Len(txtcapa.Text)

End Sub
Private Sub TxtChq_Change()

If Len(Trim(TxtChq.Text)) = 0 Then Exit Sub
If IsNumeric(TxtChq.Text) = False Then
    MsgBox "Valor inválido para este campo.", vbExclamation + vbOKOnly, App.Title
    TxtChq.Text = ""
    TxtChq.SetFocus
End If

End Sub

Private Sub TxtChq_KeyPress(KeyAscii As Integer)
    
    InibirTeclaAlfa KeyAscii
    
    If (KeyAscii = 13) Then
        TxtChq = Format$(TxtChq, "000000")
        cmdConfirma_Click
    ElseIf (KeyAscii = 27) Then
        KeyAscii = 0
        CmdSair_Click (0)
    End If
    
End Sub
Private Sub TxtCta_Change()

If Len(Trim(TxtCta.Text)) = 0 Then Exit Sub
If IsNumeric(TxtCta.Text) = False Then
    MsgBox "Valor inválido para este campo.", vbExclamation + vbOKOnly, App.Title
    TxtCta.Text = ""
    TxtCta.SetFocus
End If

End Sub

Private Sub TxtCta_KeyPress(KeyAscii As Integer)
    InibirTeclaAlfa KeyAscii
    If (KeyAscii = 13) Then
        TxtCta = Format(TxtCta, "0000000")
        TxtChq.SetFocus
    ElseIf (KeyAscii = 27) Then
        KeyAscii = 0
        CmdSair_Click (0)
    End If
End Sub
Private Sub txtcapa_KeyPress(KeyAscii As Integer)
    InibirTeclaAlfa KeyAscii
    If (KeyAscii = 13) Then
        If Len(txtcapa) > 0 Then
            txtcapa = Format(txtcapa, "00000000")
            cmdConfirma_Click
       End If
    ElseIf (KeyAscii = 27) Then
        KeyAscii = 0
        CmdSair_Click (0)
    End If
End Sub
Private Sub TxtNSU_Change()

If Len(Trim(TxtNSU.Text)) = 0 Then Exit Sub
If IsNumeric(TxtNSU.Text) = False Then
    MsgBox "Valor inválido para este campo.", vbExclamation + vbOKOnly, App.Title
    TxtNSU.Text = ""
    TxtNSU.SetFocus
End If

End Sub

Private Sub TxtNSU_KeyPress(KeyAscii As Integer)
    InibirTeclaAlfa KeyAscii
       If (KeyAscii = 13) Then
        If Len(TxtNSU) > 0 Then
            TxtNSU = Format(TxtNSU, "000000")
        End If
        cmdConfirma_Click
    ElseIf (KeyAscii = 27) Then
        KeyAscii = 0
        CmdSair_Click (0)
    End If
End Sub
Private Sub TxtVal_Change()
  
    Dim nparou As Integer
    Dim ntamanho As String
    
    If Not bAlterar Then
        Exit Sub
    End If
    
    bAlterar = False
    With TxtVal
        'Guarda posição do cursor
        nparou = .SelStart
        'Guarda tamanho texto
        ntamanho = Len(Trim(.Text))
        'Chama função formata texto
        .Text = Formata_Valor(.Text)
        .SelLength = 0
        'Calcula nova posição do cursor
        nparou = nparou + (Len(Trim(.Text)) - ntamanho)
        If nparou < 0 Then nparou = 0
        .SelStart = nparou
    End With
    bAlterar = True
    
End Sub
Private Sub list_detalhe_Click()
   
    If FlagImp = 1 Then Exit Sub
        List_detalhe.Clear
        List_detalhe.Visible = False
        Grade.SetFocus
        
End Sub
Private Sub Mostra_Imagem()

On Error GoTo ERRO_LOAD_IMG

    Dim Tam, rt, nIdCapa            As Long
    Dim valor_caption, StrMotivo    As String
    Dim Ret, PegIdCapa              As Long
    Dim Ocorrencia                  As Long
    Dim l_ordem                     As Long

    Status = Grade.TextMatrix(Grade.Row, 1)               ' Status do Documento
    Status_Capa = Grade.TextMatrix(Grade.Row, 20)         ' Status da Capa
    reg_imagem = Grade.TextMatrix(Grade.Row, 8)           ' Nome da Imagem - Frente
    Pegiddocto = CLng(Grade.TextMatrix(Grade.Row, 10))    ' Recupera IdDocto
    PegIdCapa = CLng(Grade.TextMatrix(Grade.Row, 13))     ' Recupera IdCapa
    Tip_doc = Grade.TextMatrix(Grade.Row, 11)             ' Só permite mudar para verso se for cheque
    lblLote = Grade.TextMatrix(Grade.Row, 19)             ' Envia Código do Lote para Label
    Lote = lblLote                                        ' Envia Código do Lote para variável
    lblLote = Format(lblLote, "0000-00000")               ' Formata Código do Lote
    Ocorrencia = Grade.TextMatrix(Grade.Row, 3)           ' Código de Ocorrencia
    l_ordem = Grade.TextMatrix(Grade.Row, 12)             ' Código de Ordem

    lblExclusao = ""
    '* Recupera descritivo do Motivo de Exclusão se  Status capa for = 'D'
    '  ,Ocorrencia '999' e Tipo de Documento = 1 (Capa de Envelope/Malote) *'
    If Status_Capa = "D" And Ocorrencia = 99900 And Tip_doc = "1" Then
        lblExclusao = ObtemMotivoExclusao(PegIdCapa)
    Else
    '* Se não recupera descritivo da ocorrência *'
        If Status = "D" Or Status = "F" Or Status = "C" Then
            StrMotivo = ObtemOcorrencia(Grade.TextMatrix(Grade.Row, 14))
            If Not Trim(StrMotivo) = "" Then
                lblExclusao = "Ocorrência:" & " " & StrMotivo
            Else
                lblExclusao = "Retorno de Transação Não Cadastrado"
            End If
        End If
    End If
    
    Call StatusCapa

    If reg_imagem = "" Then
        Frame4.Visible = False
        cmdFrenteVerso.Enabled = False
        cmdRotacao.Enabled = False
        cmdZoomMais.Enabled = False
        cmdZoomMenos.Enabled = False
        Imprime_Detalhe.Enabled = False
        Exit Sub
    End If
    
    If indice_corrente = 2 Or indice_corrente = 3 Or indice_corrente = 4 Or indice_corrente = 5 Or indice_corrente = 6 Then
        Set tbenv = Nothing
    
        With Modulo.qryGetIdcDocto
            .rdoParameters(0).Value = Geral.DataProcessamento
            .rdoParameters(1).Value = Pegiddocto
            Set tb1 = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
        End With

        If Not tb1.EOF Then
            nIdCapa = tb1!IdCapa
            Set tb1 = Nothing
            With Modulo.qryGetIdCapa
                .rdoParameters(0).Value = Geral.DataProcessamento
                .rdoParameters(1).Value = nIdCapa
                Set tb1 = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
            End With

            If Not tb1.EOF Then
                lblLote = tb1!IdLote
                lblLote = Format(lblLote, "0000-00000")
                Label_Envelope = tb1!Capa
                If NEnvMal = "M" Then
                    picNumMalote.Visible = True
                    lbl_NumeroMalote.Visible = True
                    lbl_NumeroMalote = FormataMalote(tb1!Num_Malote)
                Else
                    picNumMalote.Visible = False
                    lbl_NumeroMalote.Visible = False
                End If
            End If
        End If
    End If

    If AjustaBotoes = False Then
        Exit Sub
    End If
    
    Tam = Len(reg_imagem)
    If (Tam = 0) Then
        Screen.MousePointer = 0
        MsgBox "É preciso selecionar um documento para mostra a imagem.", vbInformation, App.Title
        Exit Sub
    End If

    ' Se for credito automatico, preenche campos de agencia + conta + valor + usuario
    If (Mid$(reg_imagem, 1, 6) = "DEBITO") Then
        With Lead1
            .Enabled = False
            .AutoRepaint = True
            If Geral.VIPSDLL = eDllUnibanco Then
                .Load Geral.DiretorioTrabalho & "debito.bmp", 0, 0, 1
            Else
                .Load Geral.DiretorioImagens & Format(Val(Lote), "000000000") & "\" & reg_imagem, 0, 0, 1
            End If
            .Tag = "F"
            .Intensity 220
            .PaintZoomFactor = 100
            .Enabled = True
            .AutoRepaint = True
        End With

        'posiciona imagem sempre no começo
        Ret = SendMessage(hCtl, WM_VSCROLL, SB_TOP, 0)
        Ret = SendMessage(hCtl, WM_HSCROLL, SB_TOP, 0)
        Exit Sub
    Else
    ' Se NUMERÁRIO - TESOURARIA, preenche campos ----
        If (Mid$(reg_imagem, 1, 5) = "MONEY") Then
        Else
        ' Mostra doctos na lead tools
            With Lead1
                .AutoRepaint = False
                If Geral.VIPSDLL = eDllProservi Then
                    .Load Geral.DiretorioImagens & reg_imagem, 0, 0, 1
                Else
                    .Load Geral.DiretorioImagens & Format(Val(Lote), "000000000") & "\" & reg_imagem, 0, 0, 1
                End If
                .Tag = "F"
                 hCtl = .hwnd
                .Enabled = True
                '* Verifica valor do campo ordem que pode ser: *'
                '* 0 - Vips
                '* 1 - Canon
                '* 2 - Ls 500
                '* Se for tipo 1 diminui imagem em 50% de seu valor
        
                ' se imagem for da ls500, deixar mais escura
                If l_ordem <> "2" Then
                   .Intensity 220
                Else
                   .Intensity 140
                End If
                ' se imagem for do canon, diminui em 50% o tamanho
                If l_ordem <> "1" Then
                   .PaintZoomFactor = 100
                Else
                   .PaintZoomFactor = 50
                End If
                .AutoRepaint = True
            End With

        'posiciona imagem sempre no começo
        Ret = SendMessage(hCtl, WM_VSCROLL, SB_TOP, 0)
        Ret = SendMessage(hCtl, WM_HSCROLL, SB_TOP, 0)
        Exit Sub
    End If
End If

DoEvents
Screen.MousePointer = 0
Exit Sub

ERRO_LOAD_IMG:
    Screen.MousePointer = 0

    Select Case Err
        Case 20010
            MsgBox "(OEGT) Não foi possível exibir imagem!", vbInformation, App.Title
            cmdFrenteVerso.Enabled = False
            cmdRotacao.Enabled = False
            cmdZoomMais.Enabled = False
            cmdZoomMenos.Enabled = False
            Imprime_Detalhe.Enabled = False
        Case 3044
            MsgBox "(OEGT) Diretório de Imagem Inválido ! Especifique outro diretório e execute o programa novamente. Erro: " + Error, vbInformation, App.Title
        Case 3050
            MsgBox "(OEGT) SHARE Não instalado! Finalize o WINDOWS e carregue o SHARE.EXE. Erro: " + Error, vbInformation, App.Title
        Case Else
            MsgBox "(OEGT) Erro: " + Error, vbInformation, App.Title
    End Select

End Sub
Private Sub lead1_DblClick()
      
On Error GoTo ERRO_MOSTRADETALHE

    Dim detalhe As String
    Dim Tam As Integer, p As Integer, Selecao As Integer
    
    If List_detalhe.Visible = True Then Exit Sub
    
        ' Busca primeiro os dados básicos na tabela Documento '
        With Modulo.qryDocumentoEscolhido
            .rdoParameters(0).Value = Geral.DataProcessamento
            .rdoParameters(1).Value = Grade.TextMatrix(Grade.Row, 13)
            .rdoParameters(2).Value = CLng(Grade.TextMatrix(Grade.Row, 10))
            Set tb1 = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
        End With
    
    If tb1.EOF Then
        MsgBox "Não foi possível localizar o documento escolhido. Tente novamente.", vbInformation, App.Title
        Exit Sub
    Else
        If tb1!TipoDocto <> 0 Then
            Call MontaDetalheConsulta(CInt(Grade.TextMatrix(Grade.Row, 11)))
            List_detalhe.Visible = True
        End If
    End If
    
Exit Sub

ERRO_MOSTRADETALHE:
    
    Select Case TratamentoErro("Não foi possível consultar os dados.", Err, rdoErrors)
        Case vbCancel
            Unload Me
        Case vbRetry
            Resume
    End Select
    
End Sub
Function MontaDetalheConsulta(CaseTipoProduto As Integer)
'* Esta função monta o detalhe da consulta para cada tipo de produto cadastrado *'

    Call ParteFixaDetalheConsulta

    Select Case CaseTipoProduto
        Case 0
        Case 1
            Call ProdCapa
        Case 2, 3
            Call ProdDeposito
        Case 4
            Call ProdADCC
        Case 5, 6, 7
            Call ProdCheque
        Case 8, 9
        Case 10
            Call ProdFichaCompensacao
        Case 11
        Case 12
            Call ProdTitulo
        Case 13
            Call ProdCobRegistrada
        Case 14
            Call ProdCobEspecial
        Case 15
            Call ProdDARM
        Case 16
            Call ProdDARFPRETO
        Case 17
            Call ProdDARFSIMPLES
        Case 18
            Call ProdGARE
        Case 19
        Case 20 To 26
            Call ProdArrecEletronica
        Case 27
            Call ProdArrecConvencional
        Case 28 To 31
            Call ProdFichadeCompensacao
        Case 32 To 34
            '* 32,33 Ajuste de Crédito *'
            '* 34    Ajuste de Débito  *'
        Case 35
            Call ProdGPS
        Case 36
            Call ProdCartaoAvulso
        Case 37
            Call ProdOCT
        Case 38
            '* Ajuste de Débito * '
        Case 39
            Call ProdCapaOCT
        Case 40
            Call ProdFGTS
        Case 41
            Call ProdLctoInterno
        Case 42 To 43
            '* 42 Ajuste Contábil Receita *'
            '* 43 Ajuste Contábil Despesa  *'
        Case Else
            Call CaseElse
    End Select

End Function
Sub ParteFixaDetalheConsulta()
        
Dim desc_trans, Desc_Ocorrencia, Cod_Ocorrencia As String
Dim strDescricao As String, IdDocto As Long, i As Integer
    
    List_detalhe.Clear
    List_detalhe.AddItem "       Detalhamento da Transação"
    List_detalhe.AddItem "       ========================="
    
'* Se documento possuir Status D ou F traz seus descritivos *'
   If Grade.TextMatrix(Grade.Row, 1) = "D" Or Grade.TextMatrix(Grade.Row, 1) = "F" Or Grade.TextMatrix(Grade.Row, 1) = "C" Then
      PegHistorico = Grade.TextMatrix(Grade.Row, 3)
      PegHistorico = CLng(PegHistorico)
      Desc_Ocorrencia = ObtemOcorrencia(PegHistorico)
      Cod_Ocorrencia = Grade.TextMatrix(Grade.Row, 3)
      List_detalhe.AddItem "Ocorrência Nº : " & Mid(Cod_Ocorrencia, 1, 5)
      List_detalhe.AddItem Desc_Ocorrencia
      'Busca descrição do complemento de ocorrência, caso exista
      strDescricao = ""
      IdDocto = CLng(Grade.TextMatrix(Grade.Row, 10))
'''      Call GravaComplementoOcorrencia(IdDocto, "C", strDescricao)
      strDescricao = Trim(strDescricao)
      If strDescricao <> "" Then
        If Len(strDescricao) > 55 Then
            'Quebra da descrição separando palavra
            For i = 60 To 1 Step -1
                If Mid(strDescricao, i, 1) = " " Then Exit For
            Next
            If i = 0 Then
                List_detalhe.AddItem "Complemento de Ocorrência: " & Left(strDescricao, 55)
                List_detalhe.AddItem "                           " & Mid(strDescricao, 56)
            Else
                List_detalhe.AddItem "Complemento de Ocorrência: " & Left(strDescricao, i)
                List_detalhe.AddItem "                           " & Mid(strDescricao, i + 1)
            End If
        Else
              List_detalhe.AddItem "Complemento de Ocorrência: " & strDescricao
        End If
      End If
    End If

'* Traz descrição de tipos de documentos * '
    With Modulo.qryGettipoDocto
        .rdoParameters(0).Value = CInt(Grade.TextMatrix(Grade.Row, 11))
        Set RsTipoDocto = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
               
    If RsTipoDocto.RowCount <> 0 Then
        If CInt(Grade.TextMatrix(Grade.Row, 11)) = 6 And Mid$(Grade.TextMatrix(Grade.Row, 17), 1, 3) <> "409" Then
            If CInt(Grade.TextMatrix(Grade.Row, 11)) = 6 And Mid$(Grade.TextMatrix(Grade.Row, 17), 1, 3) <> "230" Then
                desc_trans = "CHEQUE TERCEIRO (PAGTO)"
            Else
                desc_trans = RsTipoDocto!Nome
            End If
        Else
            desc_trans = RsTipoDocto!Nome
        End If
    Else
        desc_trans = "Documento - " & Format((Grade.TextMatrix(Grade.Row, 11)), "00")
    End If
    
    Set RsTipoDocto = Nothing
            
    List_detalhe.AddItem "Transação     : " & desc_trans
    List_detalhe.AddItem "Autenticação  : " & Grade.TextMatrix(Grade.Row, 15) & " - Caixa: " & Grade.TextMatrix(Grade.Row, 16) & " - Agência: " & "9" & Format(Grade.TextMatrix(Grade.Row, 7), "0000")
    
End Sub
Private Sub txtNumMalote_Change()

If Len(Trim(TxtNumMalote)) = 0 Then Exit Sub
If IsNumeric(TxtNumMalote.Text) = False Then
    MsgBox "Valor inválido para este campo.", vbExclamation + vbOKOnly, App.Title
    TxtNumMalote.Text = ""
    TxtNumMalote.SetFocus
End If

End Sub
Private Sub TxtNumMalote_GotFocus()
    If Opn_NumMalote.Value = False Then
       TxtNumMalote.Text = ""
       TxtNumMalote.Enabled = False
    Else
       TxtNumMalote.Enabled = True
    End If
End Sub
Private Sub txtNumMalote_KeyPress(KeyAscii As Integer)
    InibirTeclaAlfa KeyAscii
    If (KeyAscii = 13) Then
        If Len(TxtNumMalote) > 0 Then
            If VerificaMalote(TxtNumMalote) = False Then
                MsgBox "Número de Malote inválido.", vbInformation, App.Title
                TxtNumMalote.SetFocus
                Exit Sub
            End If
            cmdConfirma_Click
        End If
    ElseIf (KeyAscii = 27) Then
        KeyAscii = 0
        CmdSair_Click (0)
    End If
End Sub
Private Sub TxtVal_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        If Len(TxtVal.Text) > 0 Then
            cmdConfirma_Click
        End If
    ElseIf (KeyAscii = 27) Then
        KeyAscii = 0
        CmdSair_Click (0)
    End If
End Sub
Function DescTipoProd() As String

    Select Case TipoDocto
    
        Case 0
            DescTipoProd = "INDEFINIDO"
        Case 1
            If NEnvMal = "E" Then
                DescTipoProd = "ENVELOPE"
            Else
                DescTipoProd = "MALOTE"
            End If
        Case 2, 3
            DescTipoProd = "DEPÓSITO"
        Case 4
            DescTipoProd = "DEBITO CC"
        Case 5, 6, 7
            DescTipoProd = "CHEQUE"
        Case 32, 34
            DescTipoProd = "AJUSTE DE CRÉDITO"
        Case 33, 38
            DescTipoProd = "AJUSTE DE DÉBITO"
        Case 36
            DescTipoProd = "CARTÃO AVULSO"
        Case 37
            DescTipoProd = "OCT"
        Case 39
            DescTipoProd = "CAPA OCT"
        Case 40
            DescTipoProd = "FGTS"
        Case 41
            DescTipoProd = "LANÇAMENTO INTERNO"
        Case 42
            DescTipoProd = "AJ. CONTÁBIL RECEITA"
        Case 43
            DescTipoProd = "AJ. CONTÁBIL DESPESA"
        Case Else
            DescTipoProd = "PAGAMENTO"
    
    End Select

End Function
Public Sub PesqCapaEnvMal()
'* Traz a Capa escolhida (Malote/Envelope) *'

Set tbenv = Nothing

    With Modulo.qryGetDocumentosNumCapa
        .rdoParameters(1).Value = Geral.DataProcessamento
        .rdoParameters(2).Value = CDbl(txtcapa.Text)
        .rdoParameters(3).Value = Null
        Set tbenv = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With

    If tbenv.EOF Then
        Set tbenv = Nothing
        With Modulo.qryGetStatusCapa
            .rdoParameters(0).Value = CDbl(txtcapa.Text)
            .rdoParameters(1).Value = Geral.DataProcessamento
            .rdoParameters(2).Value = Null
            .rdoParameters(3).Value = Null
           Set tbenv = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
        End With
        
        If Not tbenv.EOF Then
           If tbenv!Status = "0" Then
               MsgBox "Capa recepcionada!", vbInformation, App.Title
               Call txtCapa_GotFocus
               Exit Sub
            Else
               MsgBox "Capa excluida na preparação!", vbInformation, App.Title
               Call txtCapa_GotFocus
               Exit Sub
            End If
        End If
        
        If NEnvMal = "E" Then
            MsgBox " Envelope não encontrado !", vbInformation, App.Title
            Call txtCapa_GotFocus
        Else
            MsgBox " Malote não encontrado !", vbInformation, App.Title
            Call txtCapa_GotFocus
        End If
        
        txtcapa = ""
        txtcapa.SetFocus
        SSTab1.Tab = 0
        Exit Sub
    Else
        NEnvMal = (tbenv!IdEnv_Mal)
        ' verifica o status do envelope para impressão na tela
        If NEnvMal = "M" Then
            picNumMalote.Visible = True
            lbl_NumeroMalote.Visible = True
            lbl_NumeroMalote = FormataMalote(tbenv!Num_Malote)
        Else
            picNumMalote.Visible = False
            lbl_NumeroMalote.Visible = False
        End If
        ' Leitura dos doctos deste envelope
        Label_Envelope = txtcapa
        LeituraDoctosEnvelope
        SSTab1.Tab = 1
        SSTab1.TabEnabled(1) = True
    End If

End Sub
Public Sub PesqCheques()
'* Pequisa por Banco, Agência, Conta Corrente e Valor
'  Retorna os Cheques de acordo com parâmetros '*

Set tbenv = Nothing

    With Modulo.qryPesquisaBcoAgCtaChq
        .rdoParameters(1).Value = Geral.DataProcessamento
        .rdoParameters(2).Value = IIf(TxtBco.Text <> "", TxtBco, Null)
        .rdoParameters(3).Value = IIf(TxtAg.Text <> "", Format(TxtAg, "0000"), Null)
        .rdoParameters(4).Value = IIf(TxtCta.Text <> "", TxtCta, Null)
        .rdoParameters(5).Value = IIf(TxtChq.Text <> "", TxtChq, Null)
        .rdoParameters(6).Value = NEnvMal
        Set tbenv = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With

    If tbenv.EOF Then
        MsgBox "Não foi localizado este cheque. Tente novamente.", vbInformation, App.Title
        TxtBco.Text = ""
        TxtAg.Text = ""
        TxtCta.Text = ""
        TxtChq.Text = ""
        TxtBco.SetFocus
        SSTab1.Tab = 0
        Exit Sub
    Else
        If Len(TxtBco.Text) <> 0 And _
            Len(TxtAg.Text) <> 0 And _
            Len(TxtCta.Text) <> 0 And _
            Len(TxtChq.Text) <> 0 Then
            Opn_Capa.Value = True
            txtcapa = tbenv!Capa
            PegiddoctoSit2 = tbenv!IdDocto
            Call cmdConfirma_Click
            Exit Sub
        End If
        LeituraDoctosEnvelope  ' Leitura dos doctos deste envelope
        SSTab1.Tab = 1
        SSTab1.TabEnabled(1) = True
    End If

End Sub
Public Sub PesqValores()

' Pesquisa valores de Cheques e Documentos em Geral
Set tbenv = Nothing
idenv_antigo = 0
        
    With Modulo.qryPesquisaValor
        .rdoParameters(1) = Geral.DataProcessamento     ' data processamento
        .rdoParameters(2) = Val(TxtVal.Text) / 100      ' valor
        .rdoParameters(3) = NEnvMal                     ' envelope ou malote
        .rdoParameters(4) = IIf(TipoDoctoTab = "Null", Null, TipoDoctoTab)
        Set tbenv = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With

    If tbenv.EOF Then
        MsgBox "Valor não encontrado !", vbInformation, App.Title
        TxtVal.Text = ""
        TxtVal.SetFocus
        Exit Sub
    Else
        LeituraDoctosEnvelope   ' Leitura dos doctos deste envelope/malote
    End If

End Sub
Public Sub PesqDoctoNSU()
'* Pesquisa de documentos por NSU *'

Set tbdoctos = Nothing
idenv_antigo = 0

    With Modulo.qryGetDocumentosCapaNSU
        .rdoParameters(0) = Geral.DataProcessamento                'Data de processamento
        .rdoParameters(1) = Format(TxtNSU, "######")               'Numero NSU
        .rdoParameters(2) = NEnvMal                                'Tipo de capa
        Set tbdoctos = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With

    Do
        If tbdoctos.EOF Then
            If (achou = 0) Then
                MsgBox "Não foi localizado nenhum documento com este número de NSU. Tente novamente.", vbInformation, App.Title
                TxtNSU = ""
                TxtNSU.SetFocus
                Exit Sub
            End If
            Exit Do
        Else
            nDocto_Sel = tbdoctos!IdDocto
            NCapa_Sel = tbdoctos!IdCapa
            Set tbenv = Nothing
            
            With Modulo.qryBuscaEnvelope
                .rdoParameters(0) = Geral.DataProcessamento
                .rdoParameters(1) = NCapa_Sel
                .rdoParameters(2) = NEnvMal
                Set tbenv = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
            End With

            If tbenv.EOF Then
                MsgBox "Não foi localizado o envelope que contém este NSU. Tente novamente.", vbInformation, App.Title
            Else
                achou = 1
                LeituraDoctosEnvelope
                SSTab1.Tab = 1
                SSTab1.TabEnabled(1) = True
            End If
        End If
        tbdoctos.MoveNext
    Loop
    
End Sub
Public Sub PesqNumMalote()
'* Pesquisa de Documentos por Número de Malote *'

Set tbenv = Nothing

    With Modulo.qryGetDoctosNumMalote
        .rdoParameters(0).Value = Geral.DataProcessamento
        .rdoParameters(1).Value = Val(FormataMalote(TxtNumMalote))
        Set tbenv = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With

    If tbenv.EOF Then
        MsgBox "Não existe Malote com este Número !", vbInformation, App.Title
        TxtNumMalote = ""
        TxtNumMalote.SetFocus
        SSTab1.Tab = 0
        Exit Sub
    Else
        NEnvMal = tbenv!IdEnv_Mal
        LeituraDoctosEnvelope
        SSTab1.Tab = 1
        SSTab1.TabEnabled(1) = True
    End If
        
End Sub
Public Sub PesqDoctoOcorrencias()
'* Pesquisa de Documentos por Ocorrência de Capa ou Documento *'

Set tbenv = Nothing

    With Modulo.qryGetTipoOcorr
        .rdoParameters(0).Value = NCapaDocto
        .rdoParameters(1).Value = Geral.DataProcessamento
        .rdoParameters(2).Value = NEnvMal
        Set tbenv = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With

    If tbenv.EOF Then
        MsgBox "Não existem Documentos com Ocorrência neste período!", vbInformation, App.Title
        txtcapa = ""
        SSTab1.Tab = 0
        Exit Sub
    Else
        LeituraDoctosEnvelope
    End If

End Sub
Public Sub TipoDoctos()
'* Insere tipos fixos de documentos na Lista *'
            
    cmbTipoDocto.AddItem "ADCC"
    cmbTipoDocto.AddItem "Ajuste Contábil Receita"
    cmbTipoDocto.AddItem "Ajuste Contátil Despesa"
    cmbTipoDocto.AddItem "Ajuste de Crédito"
    cmbTipoDocto.AddItem "Ajuste de Dédito"
    cmbTipoDocto.AddItem "Arrecadação Convencional"
    cmbTipoDocto.AddItem "Arrecadação Eletrônica"
    cmbTipoDocto.AddItem "Cartão Avulso"
    cmbTipoDocto.AddItem "CB Indexado"
    cmbTipoDocto.AddItem "Cheque"
    cmbTipoDocto.AddItem "Cobrança Especial"
    cmbTipoDocto.AddItem "Cobrança Registrada"
    cmbTipoDocto.AddItem "DARF-Preto"
    cmbTipoDocto.AddItem "DARF-Simples"
    cmbTipoDocto.AddItem "DARM"
    cmbTipoDocto.AddItem "Depósito"
    cmbTipoDocto.AddItem "Ficha de Compensação"
    cmbTipoDocto.AddItem "FGTS"
    cmbTipoDocto.AddItem "GARE"
    cmbTipoDocto.AddItem "GPS"
    cmbTipoDocto.AddItem "Lançamento Interno"
    cmbTipoDocto.AddItem "OCT"
    cmbTipoDocto.AddItem "Título Convencional"
    cmbTipoDocto.AddItem "Todos"

End Sub
Public Sub PesqProduto()

    Select Case (cmbTipoDocto.ListIndex + 1):

      Case 1
         TipoDoctoTab = "04,00,00,00,00"
      Case 2
         TipoDoctoTab = "42,00,00,00,00"
      Case 3
         TipoDoctoTab = "43,00,00,00,00"
      Case 4
         TipoDoctoTab = "32,34,00,00,00"
      Case 5
         TipoDoctoTab = "33,38,00,00,00"
      Case 6
         TipoDoctoTab = "27,00,00,00,00"
      Case 7
         TipoDoctoTab = "20,21,22,23,00"
      Case 8
         TipoDoctoTab = "36,00,00,00,00"
      Case 9
         TipoDoctoTab = "08,09,24,25,26"
      Case 10
         TipoDoctoTab = "05,06,07,00,00"
      Case 11
         TipoDoctoTab = "14,00,00,00,00"
      Case 12
         TipoDoctoTab = "13,00,00,00,00"
      Case 13
         TipoDoctoTab = "16,00,00,00,00"
      Case 14
         TipoDoctoTab = "17,00,00,00,00"
      Case 15
         TipoDoctoTab = "15,00,00,00,00"
      Case 16
         TipoDoctoTab = "02,03,00,00,00"
      Case 17
         TipoDoctoTab = "28,29,30,31,00"
      Case 18
         TipoDoctoTab = "40,00,00,00,00"
      Case 19
         TipoDoctoTab = "18,00,00,00,00"
      Case 20
         TipoDoctoTab = "35,00,00,00,00"
      Case 21
         TipoDoctoTab = "41,00,00,00,00"
      Case 22
         TipoDoctoTab = "37,00,00,00,00"
      Case 23
         TipoDoctoTab = "12,00,00,00,00"
      Case 24
         TipoDoctoTab = "Null"
    End Select
        
End Sub
Public Sub StatusCapa()
   
    If Status_Capa = "D" Then
        If Grade.TextMatrix(Grade.Row, 14) <> 998 Or Grade.TextMatrix(Grade.Row, 14) <> 99800 Then
            Panel_SitEnv.Caption = "Capa excluida pelo sistema"
            Exit Sub
        End If
    End If
    
    With Modulo.qryGetConsStatusCapa
        .rdoParameters(0) = Status_Capa
         Set RsConsStatusCapa = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
       
    If Not RsConsStatusCapa.EOF Then
        Panel_SitEnv.Caption = RsConsStatusCapa!Descricao
    End If

End Sub
Function AjustaBotoes() As Boolean
    
    If (CInt(Tip_doc) = 32) Then
        Frame4.Visible = False
        cmdRotacao.Enabled = False
        cmdZoomMais.Enabled = False
        lblExclusao.Visible = False
        cmdZoomMenos.Enabled = False
        cmdFrenteVerso.Enabled = False
        Imprime_Detalhe.Enabled = False
        MsgBox "Este Documento foi um ajuste gerado pelo sistema de caixa." & vbCr & "Não há imagem a ser mostrada.", vbInformation, App.Title
        AjustaBotoes = False
    ElseIf (CInt(Tip_doc) = 33) Then
        Frame4.Visible = False
        cmdRotacao.Enabled = False
        cmdZoomMais.Enabled = False
        lblExclusao.Visible = False
        cmdZoomMenos.Enabled = False
        cmdFrenteVerso.Enabled = False
        Imprime_Detalhe.Enabled = False
        MsgBox "Este Documento foi um ajuste gerado pelo sistema de caixa." & vbCr & "Não há imagem a ser mostrada.", vbInformation, App.Title
        AjustaBotoes = False
    ElseIf (CInt(Tip_doc) = 34) Then
        Frame4.Visible = False
        cmdRotacao.Enabled = False
        cmdZoomMais.Enabled = False
        lblExclusao.Visible = False
        cmdZoomMenos.Enabled = False
        cmdFrenteVerso.Enabled = False
        Imprime_Detalhe.Enabled = False
        MsgBox "Este Documento foi um ajuste gerado pelo sistema de caixa." & vbCr & "Não há imagem a ser mostrada.", vbInformation, App.Title
        AjustaBotoes = False
    ElseIf (CInt(Tip_doc) = 38) Then
        Frame4.Visible = False
        cmdRotacao.Enabled = False
        cmdZoomMais.Enabled = False
        lblExclusao.Visible = False
        cmdZoomMenos.Enabled = False
        cmdFrenteVerso.Enabled = False
        Imprime_Detalhe.Enabled = False
        MsgBox "Este Documento foi um ajuste gerado pelo sistema de caixa." & vbCr & "Não há imagem a ser mostrada.", vbInformation, App.Title
        AjustaBotoes = False
    ElseIf (CInt(Tip_doc) = 42) Then
        Frame4.Visible = False
        cmdRotacao.Enabled = False
        cmdZoomMais.Enabled = False
        lblExclusao.Visible = False
        cmdZoomMenos.Enabled = False
        cmdFrenteVerso.Enabled = False
        Imprime_Detalhe.Enabled = False
        MsgBox "Este Documento foi um ajuste gerado pelo sistema de caixa." & vbCr & "Não há imagem a ser mostrada.", vbInformation, App.Title
        AjustaBotoes = False
    ElseIf (CInt(Tip_doc) = 43) Then
        Frame4.Visible = False
        cmdRotacao.Enabled = False
        cmdZoomMais.Enabled = False
        lblExclusao.Visible = False
        cmdZoomMenos.Enabled = False
        cmdFrenteVerso.Enabled = False
        Imprime_Detalhe.Enabled = False
        MsgBox "Este Documento foi um ajuste gerado pelo sistema de caixa." & vbCr & "Não há imagem a ser mostrada.", vbInformation, App.Title
        AjustaBotoes = False
    Else
        Frame4.Visible = True
        cmdRotacao.Enabled = True
        cmdZoomMais.Enabled = True
        lblExclusao.Visible = True
        cmdZoomMenos.Enabled = True
        cmdFrenteVerso.Enabled = True
        Imprime_Detalhe.Enabled = True
        AjustaBotoes = True
    End If
    
End Function
Function ObtemMotivoExclusao(IdcapaMot As Long) As String
'* Esta função tem a finalidade de retornar o descritivo do motivo de exclusão, '
'  para uma capa selecionada e que possua Status = 'D' - Deletado *'

    With Modulo.qryGetMotivoExclusao
        .rdoParameters(0) = Geral.DataProcessamento
        .rdoParameters(1) = IdcapaMot
        Set RsMotivoExclusao = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If Not RsMotivoExclusao.EOF Then
        ObtemMotivoExclusao = RsMotivoExclusao!MotivoExclusao
    End If

End Function
Function VerificaMalote(ValNumMalote As String) As Boolean
'* Verifica se número de malote é Válido *'

    If Len(ValNumMalote) = 12 And CStr(Mid(ValNumMalote, 1, 2)) <> "09" Then
       VerificaMalote = False
    Else
       VerificaMalote = True
    End If
        
End Function
Sub ProdDeposito()
'* Traz os detalhes do documento DEPÓSITO *'

    With Modulo.qryGetDeposito
        .rdoParameters(0) = Geral.DataProcessamento
        .rdoParameters(1) = Pegiddocto
        Set RsDeposito = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If RsDeposito.EOF Then
        MsgBox "Não foi possível localizar os detalhes do Depósito. Tente novamente.", vbInformation, App.Title
        Exit Sub
    Else
        List_detalhe.AddItem "Agência         : " & Format(RsDeposito!Agencia, "0000")
        List_detalhe.AddItem "Conta           : " & Format(RsDeposito!Conta, "0000000")
        If ((Grade.TextMatrix(Grade.Row, 11)) <> 1) Then
            FormataValor ((Grade.TextMatrix(Grade.Row, 21)))
            List_detalhe.AddItem "Valor  Total    : " & Trim(Format(valor_caption, "##,##0.00"))
        End If
        If RsDeposito!TipoConta = "1" Then
            List_detalhe.AddItem "Identificação   : " & (RsDeposito!TipoConta) & "  Tipo de Conta : Corrente"
        Else
            List_detalhe.AddItem "Identificação   : " & (RsDeposito!TipoConta) & " Tipo de Conta : Poupança"
        End If
    End If

    Set RsDeposito = Nothing
    
End Sub
Sub ProdADCC()
'* Traz os detalhes dos documentos do tipo ADCC - Autorização de Débito em Conta Corrente *'

    With Modulo.qryGetADCC
    .rdoParameters(0) = Geral.DataProcessamento
    .rdoParameters(1) = Pegiddocto
    Set RsADCC = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If RsADCC.EOF Then
        MsgBox "Não foi possível localizar os detalhes da Autorização de Débito. Tente novamente.", vbInformation, App.Title
        Exit Sub
    Else
        If ((Grade.TextMatrix(Grade.Row, 11)) <> 1) Then
        FormataValor ((Grade.TextMatrix(Grade.Row, 21)))
        List_detalhe.AddItem "Valor         : " & Trim(Format(valor_caption, "##,##0.00"))
        End If
        List_detalhe.AddItem "Agência       : " & Format(RsADCC!Agencia, "0000")
        List_detalhe.AddItem "Conta         : " & Format(RsADCC!Conta, "0000000")
        List_detalhe.AddItem "Linha do CMC7 : " & Mid(RsADCC!CMC7, 1, 30)
    End If

    Set RsADCC = Nothing

End Sub
Sub ProdCheque()
'* Traz os detalhes dos documentos do tipo Cheque *'

    If (Grade.TextMatrix(Grade.Row, 11) <> 1) Then
        FormataValor (Grade.TextMatrix(Grade.Row, 21))
        List_detalhe.AddItem "Valor         : " & Trim(Format(valor_caption, "##,##0.00"))
    End If
    
    List_detalhe.AddItem "Linha do CMC7 : " & Mid(Grade.TextMatrix(Grade.Row, 17), 1, 30)

End Sub
Sub ProdFichaCompensacao()
'* Traz os detalhes dos documentos do tipo Ficha de Compensação *'

    List_detalhe.AddItem "Linha do CMC7 : " & Mid(tbenv!Leitura, 1, 30)

End Sub
Sub ProdCapa()
'* Traz os detalhes do documento CAPA *'
Dim TpIk As String

    Select Case Trim(Grade.TextMatrix(Grade.Row, 22))

        Case "S":
            TpIk = " Processado OK"
            
        Case "N":
            TpIk = " Não Processado"
        
        Case "P":
            TpIk = " Processado com problema"
        
        Case Else
            Exit Sub
    
    End Select
    
    List_detalhe.AddItem "Situação IK   :" & TpIk
    
End Sub
Sub ProdTitulo()
'* Traz os detalhes dos documentos do tipo Título Outros Bcos Sem Código de Barras *'

    With Modulo.qryGetTitulo
        .rdoParameters(0) = Geral.DataProcessamento
        .rdoParameters(1) = Pegiddocto
        Set RsCob3 = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If RsCob3.EOF Then
        MsgBox "Não foi possível localizar os detalhes da Cobrança de Terceiro. Tente novamente.", vbInformation, App.Title
        Exit Sub
    Else
        List_detalhe.AddItem "Banco         : " & RsCob3!Banco
        List_detalhe.AddItem "Valor Total   : " & Format(RsCob3!Valor, "##,#00.00")
    End If

    Set RsCob3 = Nothing

End Sub
Sub ProdCobRegistrada()
'* Traz os detalhes dos documentos do tipo Cobrança Registrada *'

    With Modulo.qryGetCobRegTec
        .rdoParameters(0) = Geral.DataProcessamento
        .rdoParameters(1) = CLng(Grade.TextMatrix(Grade.Row, 10))
        Set RsCobRegTec = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If RsCobRegTec.EOF Then
        MsgBox "Não foi possível localizar os detalhes da Cobrança registrada. Tente novamente.", vbInformation, App.Title
        Exit Sub
    Else
        If Not IsNull(RsCobRegTec!BHVC_Descricao) = True Then
            List_detalhe.AddItem "Retorno VC    : " & RsCobRegTec!BHVC_Descricao
        End If
        If Not IsNull(RsCobRegTec!CVT) = True Then
            List_detalhe.AddItem "CVT           : " & RsCobRegTec!CVT
        End If
        If Not IsNull(RsCobRegTec!Agencia) = True Then
            List_detalhe.AddItem "Agencia       : " & RsCobRegTec!Agencia
        End If
        If Not IsNull(RsCobRegTec!NossoNumero) = True Then
            List_detalhe.AddItem "Nosso Numero  : " & RsCobRegTec!NossoNumero
        End If
        If Not IsNull(RsCobRegTec!vecto) = True Then
            List_detalhe.AddItem "Vencimento    : " & Mid(RsCobRegTec!vecto, 7, 2) & "/" & Mid(RsCobRegTec!vecto, 5, 2) & "/" & Mid(RsCobRegTec!vecto, 1, 4)
        End If
        If Not IsNull(RsCobRegTec!ValorBase) = True Then
            List_detalhe.AddItem "Valor Base    : " & Format(RsCobRegTec!ValorBase, "##,#00.00")
        End If
        If Not IsNull(RsCobRegTec!Valor) = True Then
            List_detalhe.AddItem "Valor Cobrado : " & Format(RsCobRegTec!Valor, "##,#00.00")
        End If
    End If

    Set RsCobRegTec = Nothing
    
End Sub
Sub ProdCobEspecial()
'* Traz os detalhes dos documentos do tipo Cobrança Especial *'

    With Modulo.qryGetCobEspTec
        .rdoParameters(0) = Geral.DataProcessamento
        .rdoParameters(1) = CLng(Grade.TextMatrix(Grade.Row, 10))
        Set RsCobEspTec = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If RsCobEspTec.EOF Then
        MsgBox "Não foi possível localizar os detalhes da Cobrança Especial UBB. Tente novamente.", vbInformation, App.Title
        Exit Sub
    Else
        If IsNull(RsCobEspTec!BHVC_Descricao) = False Then
            List_detalhe.AddItem "Retorno VC    : " & RsCobEspTec!BHVC_Descricao
        End If
        List_detalhe.AddItem "CVT           : " & RsCobEspTec!CVT
        List_detalhe.AddItem "Agência       : " & RsCobEspTec!Agencia
        List_detalhe.AddItem "Conta         : " & RsCobEspTec!Cedente
        List_detalhe.AddItem "Nosso Número  : " & RsCobEspTec!NossoNumero
        List_detalhe.AddItem "Vencimento    : " & Mid(RsCobEspTec!vecto, 7, 2) & "/" & Mid(RsCobEspTec!vecto, 5, 2) & "/" & Mid(RsCobEspTec!vecto, 1, 4)
        List_detalhe.AddItem "Valor Cob.    : " & Format(RsCobEspTec!ValorBase, "##,#00.00")
        List_detalhe.AddItem "Juros         : " & Format(RsCobEspTec!Juros, "##,#00.00")
        List_detalhe.AddItem "Desconto      : " & Format(RsCobEspTec!Desconto, "##,#00.00")
        List_detalhe.AddItem "Abatimento    : " & Format(RsCobEspTec!Abatimento, "##,#00.00")
    End If

    Set RsCobEspTec = Nothing
    
End Sub
Sub ProdDARM()
'* Traz os detalhes dos documentos do tipo DARM - Docto de Arrecadação de Tributo Mobiliarios *'

    With Modulo.qryGetDarm
        .rdoParameters(0) = Geral.DataProcessamento
        .rdoParameters(1) = Pegiddocto
        Set RsDARM = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If RsDARM.EOF Then
        MsgBox "Não foi possível localizar os detalhes do Darm. Tente novamente.", vbInformation, App.Title
        Exit Sub
    Else
        List_detalhe.AddItem "Incidência    : " & Mid(RsDARM!Incidencia, 3, 2) & "/" & Mid(RsDARM!Incidencia, 5, 4)
        List_detalhe.AddItem "Tributo       : " & RsDARM!Tributo
        List_detalhe.AddItem "C.C.M.        : " & RsDARM!ccm
        FormataValor (RsDARM!Valor)
        List_detalhe.AddItem "Valor         : " + Format(valor_caption, "##,#00.00")
    End If

    Set RsDARM = Nothing
    
End Sub
Sub ProdDARFPRETO()
'* Traz os detalhes dos documentos do tipo DARF - Docto de Arrecadação de Receitas Federais *'

    With Modulo.qryGetDARFPreto
        .rdoParameters(0) = Geral.DataProcessamento
        .rdoParameters(1) = Pegiddocto
        Set RsDarfP = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If RsDarfP.EOF Then
        MsgBox "Não foi possível localizar os detalhes do Darf Preto. Tente novamente.", vbInformation, App.Title
        Exit Sub
    Else
        List_detalhe.AddItem "Vencimento     : " & Mid(RsDarfP!vecto, 7, 2) & "/" & Mid(RsDarfP!vecto, 5, 2) & "/" & Mid(RsDarfP!vecto, 1, 4)
        List_detalhe.AddItem "Período  Apur. : " & Mid(RsDarfP!PeriodoApuracao, 7, 2) & "/" & Mid(RsDarfP!PeriodoApuracao, 5, 2) & "/" & Mid(RsDarfP!PeriodoApuracao, 1, 4)
        List_detalhe.AddItem "CPF/CGC        : " & Trim(RsDarfP!CPFCGC)
        List_detalhe.AddItem "Receita        : " & Trim(RsDarfP!CodigoReceita)
        List_detalhe.AddItem "Referência     : " & Trim(RsDarfP!Referencia)
        
        valor_caption = ""
        Formata_Valor (RsDarfP!Valor)
        
        List_detalhe.AddItem "Valor Principal: " & Trim(Format(RsDarfP!Valor, "#,##0.00"))
        
        If (RsDarfP!ValorMulta) <> 0 Then
            List_detalhe.AddItem "Valor Multa    : " & Trim(Format(RsDarfP!ValorMulta, "##,#00.00"))
        Else
            List_detalhe.AddItem "Valor Multa    : " & "0,00"
        End If
        
        If RsDarfP!Juros <> 0 Then
            List_detalhe.AddItem "Valor Juros    : " & Trim(Format(RsDarfP!Juros, "##,#00.00"))
        Else
            List_detalhe.AddItem "Valor Juros    : " & "0,00"
        End If
        
        If ((Grade.TextMatrix(Grade.Row, 11)) <> 1) Then
            FormataValor ((Grade.TextMatrix(Grade.Row, 21)))
            List_detalhe.AddItem "Valor Total    : " & Trim(Format(valor_caption, "##,##0.00"))
        End If
        
    End If

    Set RsDarfP = Nothing

End Sub
Sub ProdDARFSIMPLES()
'* Traz os detalhes dos documentos do tipo DARF - Docto de Arrecadação de Receitas Federais *'

    With Modulo.qryGetDARFSimples
        .rdoParameters(0) = Geral.DataProcessamento
        .rdoParameters(1) = Pegiddocto
        Set RsDARFS = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If RsDARFS.EOF Then
        MsgBox "Não foi possível localizar os detalhes do Darf Simples. Tente novamente.", vbInformation, App.Title
        Exit Sub
    Else
        List_detalhe.AddItem "Vencimento        : " & Mid(RsDARFS!PeriodoApuracao, 1, 2) & "/" & Mid(RsDARFS!PeriodoApuracao, 3, 2) & "/" & Mid(RsDARFS!PeriodoApuracao, 5, 4)
        List_detalhe.AddItem "CPF/CGC           : " & Trim(RsDARFS!CGC)
        If RsDARFS!ReceitaBruta <> 0 Then
            valor_caption = ""
            FormataValor (RsDARFS!ReceitaBruta)
            valor_caption = Format(valor_caption, "##,##0.00")
            List_detalhe.AddItem "Receita           : " & Trim(valor_caption)
        Else
            List_detalhe.AddItem "Receita             : " & "0,00"
        End If
        valor_caption = ""
        FormataValor (RsDARFS!ValorPrincipal)
        List_detalhe.AddItem "Valor Principal   : " & Trim(Format(valor_caption, "##,##0.00"))
        If RsDARFS!ValorMulta <> 0 Then
            valor_caption = ""
            Formata_Valor (RsDARFS!ValorMulta)
            List_detalhe.AddItem "Valor Multa   : " & Trim(Format(valor_caption, "##,#00.00"))
        Else
            List_detalhe.AddItem "Valor Multa       : " & "0,00"
        End If
        If RsDARFS!Juros <> 0 Then
            valor_caption = ""
            Formata_Valor (RsDARFS!Juros)
            List_detalhe.AddItem "Valor Juros   : " & Trim(Format(valor_caption, "##,#00.00"))
        Else
            List_detalhe.AddItem "Valor Juros       : " & "0,00"
        End If
        If RsDARFS!Percentual <> 0 Then
            List_detalhe.AddItem "Percentual        : " & Format(RsDARFS!Percentual, "#.00")
        Else
            List_detalhe.AddItem "Percentual        : " & "0,00"
        End If

    End If
    
    Set RsDARFS = Nothing
    
End Sub
Sub ProdGARE()
'* Traz os detalhes dos documentos do tipo GARE - Guia de Arrecadação Estadual  *'

    With Modulo.qryGetGare
        .rdoParameters(0) = Geral.DataProcessamento
        .rdoParameters(1) = Pegiddocto
        Set rsGare = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If rsGare.EOF Then
        MsgBox "Não foi possível localizar os detalhes do Gare. Tente novamente.", vbInformation, App.Title
        Exit Sub
    Else
        If rsGare!vecto <> 0 Then
            List_detalhe.AddItem "Vencimento    : " & Mid(rsGare!vecto, 1, 2) & "/" & Mid(rsGare!vecto, 3, 2) & "/" & Mid(rsGare!vecto, 5, 4)
        Else
            List_detalhe.AddItem "Vencimento    : "
        End If
        List_detalhe.AddItem "Receita         : " & Trim(rsGare!Receita)
        List_detalhe.AddItem "Insc. Estadual  : " & Trim(rsGare!InscricaoEstadual)
        List_detalhe.AddItem "CGC/CPF         : " & Trim(rsGare!CPFCGC)
        List_detalhe.AddItem "Divida Ativa    : " & Trim(rsGare!DividaAtiva)
        If rsGare!Referencia <> 0 Then
            List_detalhe.AddItem "Referência      : " & Mid(rsGare!Referencia, 5, 2) & "/" & Mid(rsGare!Referencia, 1, 4)
        Else
            List_detalhe.AddItem "Referência      :"
        End If
        List_detalhe.AddItem "Autent. Digital : " & Trim(rsGare!AutenticacaoDigital)
        List_detalhe.AddItem "AIIM            : " & Trim(rsGare!AIIM)
        List_detalhe.AddItem "Valor da Receita: " & Trim(Format(rsGare!ValorReceita, "##,#00.00"))
        List_detalhe.AddItem "Mora            : " & Trim(Format(rsGare!Juros, "##,#00.00"))
        List_detalhe.AddItem "Multa           : " & Trim(Format(rsGare!Multa, "##,#00.00"))
        List_detalhe.AddItem "Acrescimos      : " & Trim(Format(rsGare!Acrescimo, "##,#00.00"))
        List_detalhe.AddItem "Honorarios      : " & Trim(Format(rsGare!Honorarios, "##,#00.00"))
        List_detalhe.AddItem "Valor Total     : " & Trim(Format(rsGare!Valor, "##,#00.00"))
    End If

    Set rsGare = Nothing
    
End Sub
Public Sub ProdArrecEletronica()
'* Traz os detalhes dos documentos do tipo Arrecadação Eletrônica *'

    If ((Grade.TextMatrix(Grade.Row, 11)) <> 1) Then
        FormataValor ((Grade.TextMatrix(Grade.Row, 21)))
        List_detalhe.AddItem "Valor            : " & Trim(Format(valor_caption, "##,##0.00"))
    End If
    
    List_detalhe.AddItem "Codigo de Barras : " & (Grade.TextMatrix(Grade.Row, 17))

End Sub
Public Sub ProdArrecConvencional()
'* Traz os detalhes dos documentos do tipo Arrecadação Convencional *'

    With Modulo.qryGetArrecConvenc
        .rdoParameters(0) = Geral.DataProcessamento
        .rdoParameters(1) = Pegiddocto
        Set RsArrConv = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If RsArrConv.EOF Then
        MsgBox "Não foi possível localizar os detalhes da Arrecadação. Tente novamente.", vbInformation, App.Title
        Exit Sub
    Else
        List_detalhe.AddItem "Código        : " & RsArrConv!Produto
        List_detalhe.AddItem "Requisição    : " & RsArrConv!Requisicao
        List_detalhe.AddItem "Valor         : " & Format(RsArrConv!Valor, "##,#00.00")
    End If

    Set RsArrConv = Nothing

End Sub
Sub ProdFichadeCompensacao()
'* Traz os detalhes dos documentos do tipo Ficha de Compensação *'

    With Modulo.qryGetCobCodBar
        .rdoParameters(0) = Geral.DataProcessamento
        .rdoParameters(1) = Pegiddocto
        Set RsCobCodBar = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If RsCobCodBar.EOF Then
        MsgBox "Não foi possível localizar os detalhes da Unicobrança. Tente novamente.", vbInformation, App.Title
        Exit Sub
    Else
        List_detalhe.AddItem "Codigo de Barras : " & (Grade.TextMatrix(Grade.Row, 17))
        If Not IsNull(RsCobCodBar!BHVC_Descricao) = True Then
            List_detalhe.AddItem "Retorno VC       : " & RsCobCodBar!BHVC_Descricao
        End If
        List_detalhe.AddItem "Vencimento       : " & Mid(RsCobCodBar!vecto, 7, 2) & "/" & Mid(RsCobCodBar!vecto, 5, 2) & "/" & Mid(RsCobCodBar!vecto, 1, 4)
        List_detalhe.AddItem "Valor Base       : " & Format(RsCobCodBar!ValorBase, "##,#00.00")
        List_detalhe.AddItem "Juros            : " & Format(RsCobCodBar!Juros, "##,#00.00")
        List_detalhe.AddItem "Mora             : " & Format(RsCobCodBar!Mora, "##,#00.00")
        List_detalhe.AddItem "Descontos        : " & Format(RsCobCodBar!Desconto, "##,#00.00")
        List_detalhe.AddItem "Abatimentos      : " & Format(RsCobCodBar!Abatimento, "##,#00.00")
        If (Grade.TextMatrix(Grade.Row, 11) <> 1) Then
            FormataValor (Grade.TextMatrix(Grade.Row, 21))
            List_detalhe.AddItem "Valor Cobrado    : " & Trim(Format(valor_caption, "##,##0.00"))
        End If
    End If

    Set RsCobCodBar = Nothing

End Sub
Sub ProdAjusteCredDeb()
'* Traz os detalhes dos documentos do tipo Ajuste de Crédito / Débito *'

    List_detalhe.AddItem "Ajuste Depósito C/C"
    List_detalhe.AddItem "Valor : " & Format(tb1!ValorTotal, "##,#00.00")
    
End Sub
Sub ProdGPS()
'* Traz os detalhes dos documentos do tipo GPS - Guia da Previdência Social *'

    With Modulo.qryGetGps
        .rdoParameters(0) = Geral.DataProcessamento
        .rdoParameters(1) = Pegiddocto
        Set RsGps = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If RsGps.EOF Then
        MsgBox "Não foi possível localizar os detalhes do Gps. Tente novamente.", vbInformation, App.Title
        Exit Sub
    Else
        List_detalhe.AddItem "Cod. Pagamento    : " & Trim(RsGps!CodigoPagamento)
        If RsGps!Competencia <> 0 Then
            List_detalhe.AddItem "Competência       : " & Mid(RsGps!Competencia, 5, 2) & "/" & Mid(RsGps!Competencia, 1, 4)
        Else
            List_detalhe.AddItem "Competência       : "
        End If
        List_detalhe.AddItem "Identificador     : " & Trim(RsGps!Identificador)
        List_detalhe.AddItem "Valor INSS        : " & Trim(Format(RsGps!ValorINSS, "##,#00.00"))
        List_detalhe.AddItem "Vl. Entidades     : " & Trim(Format(RsGps!ValorEntidades, "##,#00.00"))
        List_detalhe.AddItem "Juros             : " & Trim(Format(RsGps!Juros, "##,#00.00"))
        List_detalhe.AddItem "Total             : " & Trim(Format(RsGps!Valor, "##,#00.00"))
    End If
    
    Set RsGps = Nothing
    
End Sub
Sub ProdCartaoAvulso()
'* Traz os detalhes dos documentos do tipo Cartão Avulso *'

    With Modulo.qryCartaoAvulso
        .rdoParameters(0) = Geral.DataProcessamento
        .rdoParameters(1) = Pegiddocto
        Set RsCartaoAv = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If RsCartaoAv.EOF Then
        MsgBox "Não foi possível localizar os detalhes do Cartao de Credito Avulso. Tente novamente.", vbInformation, App.Title
        Exit Sub
    Else
        List_detalhe.AddItem "Número do Cartão : " & Trim(RsCartaoAv!Cartao)
        List_detalhe.AddItem "Valor            : " & Trim(Format(RsCartaoAv!Valor, "##,#00.00"))
    End If

    Set RsCartaoAv = Nothing
    
End Sub
Sub ProdOCT()
'* Traz os detalhes dos documentos do tipo OCT - Ordem de Crédito Terceiro *'

    With Modulo.qryGetOCT
        .rdoParameters(0) = Geral.DataProcessamento
        .rdoParameters(1) = Pegiddocto
        Set RsOCT = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If RsOCT.EOF Then
        MsgBox "Não foi possível localizar os detalhes do Cartao de Credito Avulso. Tente novamente.", vbInformation, App.Title
        Exit Sub
    Else
        List_detalhe.AddItem "Agência          : " & Trim(RsOCT!AgenciaCredito)
        List_detalhe.AddItem "Conta Corrente   : " & Trim(RsOCT!ContaCredito)
        List_detalhe.AddItem "Ref. Cliente     : " & Trim(RsOCT!Referencia)
        List_detalhe.AddItem "Número OCT       : " & Trim(RsOCT!OrdemCredito)
        List_detalhe.AddItem "Agência Cliente  : " & Trim(RsOCT!AgCliente)
        List_detalhe.AddItem "Conta Cliente    : " & Trim(RsOCT!CtaCliente)
        List_detalhe.AddItem "Valor Total      : " & Format(RsOCT!Valor, "##,#00.00")
    End If
    
    Set RsOCT = Nothing
    
End Sub
Sub ProdCapaOCT()
'* Traz os detalhes dos documentos do tipo Capa de OCT - Ordem de Crédito Terceiro *'

    List_detalhe.AddItem "CMC7          : " & Mid(Grade.TextMatrix(Grade.Row, 17), 1, 44)

End Sub
Sub ProdFGTS()
'* Traz os detalhes dos documentos do tipo FGTS - Fundo de Garantia por Tempo de Serviço *'

    With Modulo.qryGetFGTS
        .rdoParameters(0) = Geral.DataProcessamento
        .rdoParameters(1) = Pegiddocto
        Set RsFGTS = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If RsFGTS.EOF Then
        MsgBox "Não foi possível localizar os detalhes do FGTS. Tente novamente.", vbInformation, App.Title
        Exit Sub
    Else
        List_detalhe.AddItem "Cód.Recolhimento  : " & Trim(RsFGTS!CodRecolhimento)
        List_detalhe.AddItem "CNPJ Empresa      : " & Trim(RsFGTS!CNPJCEI_Empresa)
        List_detalhe.AddItem "Competência       : " & Mid$(RsFGTS!Competencia, 5, 2) & "/" & Mid$(RsFGTS!Competencia, 1, 4)
        List_detalhe.AddItem "Validade          : " & (Mid$(RsFGTS!Validade, 7, 2)) & "/" & Mid$(RsFGTS!Validade, 5, 2) & "/" & Mid$(RsFGTS!Validade, 1, 4)
        List_detalhe.AddItem "CNPJ Tomador      : " & Trim(RsFGTS!CNPJCEI_Tomador)
        List_detalhe.AddItem "Depósito          : " & Format(RsFGTS!Deposito, "##,#00.00")
        List_detalhe.AddItem "Juros/Acres./Mora : " & Format(RsFGTS!JAM, "##,#00.00")
        List_detalhe.AddItem "Multa             : " & Format(RsFGTS!Multa, "##,#00.00")
        List_detalhe.AddItem "Valor             : " & Format(RsFGTS!Valor, "##,#00.00")
    End If

    Set RsFGTS = Nothing

End Sub
Sub ProdLctoInterno()
'* Traz os detalhes dos documentos do tipo Lançamento Interno *'
    
    With Modulo.qryGetLctoInterno
        .rdoParameters(0) = Geral.DataProcessamento
        .rdoParameters(1) = Pegiddocto
        Set RsLctoInterno = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If RsLctoInterno.EOF Then
        MsgBox "Não foi possível localizar os detalhes do Lançamento Interno. Tente novamente.", vbInformation, App.Title
        Exit Sub
    Else
        List_detalhe.AddItem "Cód. Envento      : " & Trim(RsLctoInterno!Evento)
        List_detalhe.AddItem "Controle Banco    : " & Trim(RsLctoInterno!ControleBanco)
        List_detalhe.AddItem "Valor             : " & Format(RsLctoInterno!Valor, "##,#00.00")
    End If

    Set RsLctoInterno = Nothing
    
End Sub
Sub CaseElse()
'* Mensagem padrão para documentos não tratados pela Consulta *'

    MsgBox "Codigo de transação não identificado para o detalhe de Consulta.", vbInformation, App.Title
    List_detalhe.AddItem "Dados Leitura : " & Mid(Grade.TextMatrix(Grade.Row, 17), 1, 44)

End Sub

