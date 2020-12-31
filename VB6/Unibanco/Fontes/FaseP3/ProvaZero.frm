VERSION 5.00
Object = "{00080003-B1BA-11CE-ABC6-F5B2E79D9E3F}#8.0#0"; "LTOCX80N.OCX"
Object = "{9CBA5D64-E3C8-11D3-9FFC-00104BC8688C}#1.0#0"; "CurrencyEdit.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ProvaZero 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Prova Zero"
   ClientHeight    =   8724
   ClientLeft      =   12
   ClientTop       =   168
   ClientWidth     =   11748
   ForeColor       =   &H8000000F&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8724
   ScaleWidth      =   11748
   Begin VB.PictureBox frmLocalizar 
      Height          =   1308
      Left            =   4548
      ScaleHeight     =   1260
      ScaleWidth      =   2604
      TabIndex        =   51
      Top             =   2016
      Visible         =   0   'False
      Width           =   2652
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   300
         Left            =   1392
         TabIndex        =   55
         Top             =   864
         Width           =   972
      End
      Begin VB.CommandButton cmdProcurar 
         Caption         =   "&Localizar"
         Height          =   300
         Left            =   144
         TabIndex        =   53
         Top             =   864
         Width           =   1068
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
         Left            =   144
         MaxLength       =   18
         TabIndex        =   52
         Top             =   384
         Width           =   2316
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
         Left            =   144
         TabIndex        =   54
         Top             =   96
         Width           =   2232
      End
   End
   Begin VB.Timer tmrPesquisa 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   10020
      Top             =   3876
   End
   Begin VB.PictureBox FrmPesquisa 
      Height          =   1932
      Left            =   2718
      ScaleHeight     =   1884
      ScaleWidth      =   6264
      TabIndex        =   47
      Top             =   1788
      Visible         =   0   'False
      Width           =   6312
      Begin VB.CommandButton CmdFecharPesquisa 
         Caption         =   "&Fechar"
         Height          =   312
         Left            =   2604
         TabIndex        =   48
         Top             =   1464
         Width           =   1068
      End
      Begin ComctlLib.ProgressBar Progress 
         Height          =   300
         Left            =   348
         TabIndex        =   49
         Top             =   912
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   529
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Pesquisando por Documentos para Prova Zero. Aguarde ..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   348
         TabIndex        =   50
         Top             =   576
         Width           =   5664
      End
   End
   Begin VB.PictureBox Picture4 
      Height          =   264
      Left            =   2808
      ScaleHeight     =   216
      ScaleWidth      =   552
      TabIndex        =   44
      Top             =   24
      Width           =   600
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
         TabIndex        =   45
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.PictureBox picNumMalote 
      Height          =   264
      Left            =   4740
      ScaleHeight     =   216
      ScaleWidth      =   1176
      TabIndex        =   41
      Top             =   24
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
         TabIndex        =   42
         Top             =   0
         Width           =   1092
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   264
      Left            =   2784
      ScaleHeight     =   216
      ScaleWidth      =   6864
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   360
      Width           =   6912
      Begin VB.Label Label14 
         Caption         =   "Recap."
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
         Left            =   2280
         TabIndex        =   58
         Top             =   0
         Width           =   624
      End
      Begin VB.Label Label6 
         Caption         =   "Valor"
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
         Height          =   216
         Left            =   5568
         TabIndex        =   25
         Top             =   0
         Width           =   540
      End
      Begin VB.Label Label5 
         Caption         =   "Documento"
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
         Left            =   3000
         TabIndex        =   24
         Top             =   0
         Width           =   1104
      End
      Begin VB.Label Label4 
         Caption         =   "Ocorr."
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
         Left            =   1680
         TabIndex        =   23
         Top             =   0
         Width           =   588
      End
      Begin VB.Label Label3 
         Caption         =   "Vínculo"
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
         Left            =   828
         TabIndex        =   22
         Top             =   0
         Width           =   732
      End
      Begin VB.Label Label2 
         Caption         =   "Nro."
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
         Height          =   192
         Left            =   108
         TabIndex        =   21
         Top             =   0
         Width           =   408
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   264
      Left            =   72
      ScaleHeight     =   216
      ScaleWidth      =   2604
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   24
      Width           =   2652
      Begin VB.Label lblCapa 
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
         Height          =   228
         Left            =   12
         TabIndex        =   20
         Top             =   -24
         Width           =   1272
      End
   End
   Begin VB.ListBox lstDocto 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1728
      Left            =   2784
      TabIndex        =   1
      Top             =   732
      Width           =   6912
   End
   Begin VB.ListBox lstCapa 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2088
      IntegralHeight  =   0   'False
      Left            =   72
      TabIndex        =   0
      Top             =   372
      Width           =   2652
   End
   Begin VB.Frame Frame2 
      Height          =   2940
      Left            =   9780
      TabIndex        =   33
      Top             =   -48
      Width           =   1752
      Begin VB.CommandButton CmdRecaptura 
         Caption         =   "Reca&ptura"
         Height          =   288
         Left            =   144
         MaskColor       =   &H00404080&
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   1596
         Width           =   1460
      End
      Begin VB.CommandButton CmdTrocaOrderm 
         Caption         =   "&Troca de Ordem"
         Height          =   276
         Left            =   144
         TabIndex        =   56
         Top             =   1320
         Width           =   1464
      End
      Begin VB.CommandButton cmdLocalizar 
         Caption         =   "L&ocalizar"
         Height          =   276
         Left            =   144
         TabIndex        =   15
         Top             =   2196
         Width           =   1464
      End
      Begin VB.CommandButton cmdAtualizar 
         Caption         =   "&Atualizar"
         Height          =   276
         Left            =   144
         TabIndex        =   11
         Top             =   192
         Width           =   1464
      End
      Begin VB.CommandButton cmdFechar 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   276
         Left            =   144
         TabIndex        =   17
         Top             =   2496
         Width           =   1464
      End
      Begin VB.CommandButton cmdEncerrar 
         Caption         =   "&Encerrar"
         Height          =   276
         Left            =   144
         TabIndex        =   16
         Top             =   1896
         Width           =   1464
      End
      Begin VB.CommandButton cmdIlegiveis 
         Caption         =   "Enviar &Ilegíveis"
         Height          =   276
         Left            =   144
         TabIndex        =   14
         Top             =   1044
         Width           =   1464
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "E&xcluir Capa"
         Height          =   276
         Left            =   144
         TabIndex        =   13
         Top             =   768
         Width           =   1464
      End
      Begin VB.CommandButton cmdCalculadora 
         Caption         =   "&Calculadora"
         Height          =   276
         Left            =   144
         TabIndex        =   12
         Top             =   480
         Width           =   1464
      End
   End
   Begin VB.Frame Frame3 
      Height          =   672
      Left            =   5172
      TabIndex        =   34
      Top             =   2880
      Width           =   6384
      Begin VB.CommandButton cmdAlterar 
         Caption         =   "A&lterar"
         Height          =   324
         Left            =   4740
         TabIndex        =   3
         Top             =   216
         Width           =   1464
      End
      Begin CURRENCYEDITLib.CurrencyEdit txtValor 
         Height          =   348
         Left            =   816
         TabIndex        =   2
         Top             =   204
         Width           =   3720
         _Version        =   65537
         _ExtentX        =   6562
         _ExtentY        =   614
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
      Begin VB.Label Label10 
         Caption         =   "Valor:"
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
         Height          =   276
         Left            =   168
         TabIndex        =   35
         Top             =   240
         Width           =   756
      End
   End
   Begin VB.Frame Frame6 
      Height          =   996
      Left            =   5172
      TabIndex        =   39
      Top             =   3528
      Width           =   6384
      Begin VB.CheckBox chkFiltro 
         Caption         =   "Filtrar capas (Títulos de outros bancos)"
         Height          =   192
         Left            =   108
         TabIndex        =   59
         Top             =   708
         Width           =   3180
      End
      Begin VB.Timer tmrAtualiza 
         Enabled         =   0   'False
         Interval        =   50000
         Left            =   5220
         Top             =   348
      End
      Begin VB.CheckBox chkOcorrencia 
         Caption         =   "Mostrar Documentos com Ocorrência"
         Height          =   204
         Left            =   108
         TabIndex        =   5
         Top             =   432
         Width           =   3012
      End
      Begin VB.CheckBox chkNaoVinculados 
         Caption         =   "Mostrar Somente Documentos Não Vinculados"
         Height          =   204
         Left            =   108
         TabIndex        =   4
         Top             =   168
         Width           =   3684
      End
   End
   Begin VB.Frame FrmImagem 
      Caption         =   "Imagem"
      Height          =   4092
      Left            =   60
      TabIndex        =   36
      Top             =   4560
      Width           =   9816
      Begin LeadLib.Lead Lead1 
         Height          =   3756
         Left            =   144
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   216
         Width           =   9552
         _Version        =   524288
         _ExtentX        =   16849
         _ExtentY        =   6625
         _StockProps     =   229
         BackColor       =   -2147483639
         BorderStyle     =   1
         ScaleHeight     =   311
         ScaleWidth      =   794
         DataField       =   ""
         BitmapDataPath  =   ""
         AnnDataPath     =   ""
      End
   End
   Begin VB.Frame Frame5 
      Height          =   4092
      Left            =   9948
      TabIndex        =   37
      Top             =   4560
      Width           =   1608
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
         Height          =   612
         Left            =   360
         Picture         =   "ProvaZero.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   228
         Width           =   900
      End
      Begin VB.CommandButton cmdFrenteVerso 
         Caption         =   "Frente/Verso"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   360
         Picture         =   "ProvaZero.frx":018A
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3336
         Width           =   900
      End
      Begin VB.CommandButton cmdInverteCor 
         Caption         =   "Inverte cor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   360
         Picture         =   "ProvaZero.frx":0494
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2712
         Width           =   900
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
         Height          =   600
         Left            =   360
         Picture         =   "ProvaZero.frx":079E
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2088
         Width           =   900
      End
      Begin VB.CommandButton cmdZoomMenos 
         Caption         =   "Zoom -"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   360
         Picture         =   "ProvaZero.frx":0AA8
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1476
         Width           =   900
      End
      Begin VB.CommandButton cmdZoomMais 
         Caption         =   "Zoom +"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   360
         Picture         =   "ProvaZero.frx":0DB2
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   852
         Width           =   900
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1644
      Left            =   60
      TabIndex        =   26
      Top             =   2880
      Width           =   4752
      Begin VB.Label Label9 
         Caption         =   "Diferença:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   276
         Left            =   120
         TabIndex        =   32
         Top             =   1044
         Width           =   1548
      End
      Begin VB.Label lblValorDiferenca 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   2196
         TabIndex        =   31
         Top             =   1116
         Width           =   2460
      End
      Begin VB.Label lblValorCheques 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   2196
         TabIndex        =   30
         Top             =   648
         Width           =   2460
      End
      Begin VB.Label Label8 
         Caption         =   "Cheques / Lancto :"
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
         Left            =   120
         TabIndex        =   29
         Top             =   624
         Width           =   1740
      End
      Begin VB.Label lblValorContas 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   2196
         TabIndex        =   28
         Top             =   180
         Width           =   2460
      End
      Begin VB.Label Label7 
         Caption         =   "Dep. / OCT / Pagtos :"
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
         Height          =   240
         Left            =   108
         TabIndex        =   27
         Top             =   228
         Width           =   1872
      End
   End
   Begin VB.Label lblLote 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0001-00001"
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
      Left            =   3456
      TabIndex        =   46
      Top             =   24
      Width           =   1176
   End
   Begin VB.Label lblNumMalote 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "06001100741"
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
      Left            =   6024
      TabIndex        =   43
      Top             =   24
      Width           =   1500
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
      Height          =   288
      Left            =   60
      TabIndex        =   40
      Top             =   2580
      Width           =   9648
   End
End
Attribute VB_Name = "ProvaZero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type tpMyCapa
    IdCapa      As Long
    IdLote      As Long
    IdEnv_Mal   As String * 1
    AgOrig      As Integer
    Capa        As String
    NumMalote   As String
    Diferenca    As Currency
End Type

Private Type tpMyDoc
    IdDocto             As Long
    TipoDocto           As Integer
    Ocorrencia          As Long
    RetornoTransacao    As Long
    Leitura             As String
    Frente              As String
    Verso               As String
    Status              As String * 1
    Vinculo             As Long
    Valor               As Currency
    Ordem               As String * 1
    Conferido           As Boolean
End Type

Private m_IsEvent       As Boolean
Private m_Busy          As Boolean
Private m_Alteracao     As Boolean
Private m_ValContas     As Currency
Private m_ValCheques    As Currency
Private m_ValDiferenca  As Currency
Private m_IdCapa        As Long
Private m_Capa          As tpMyCapa
Private m_Doc           As tpMyDoc
Private aCapa()         As tpMyCapa
Private aDoc()          As tpMyDoc
Private m_CountCapa     As Integer
Private m_CountDocto    As Integer
Private sTempo          As Integer
Private m_FirstActivate As Boolean

Private qryGetCapaProvaZero             As rdoQuery
Private qryGetDocumentoProvaZero        As rdoQuery
Private qryGetocorrencia                As rdoQuery
Private qryAtualizaStatusCapa           As rdoQuery
Private qryAtualizaValorDocumento       As rdoQuery
Private qryVerificaCapaDisponivel       As rdoQuery
Private qryGetDocumentosParaVerificacao As rdoQuery

Private rsCapa                          As rdoResultset
Private rsDoc                           As rdoResultset
Private RsOcorrencia                    As rdoResultset
Private Sub LimparValores()

    lblValorContas.Caption = ""
    lblValorCheques.Caption = ""
    lblValorDiferenca.Caption = ""
    txtValor.Text = ""
End Sub
Private Sub LimparHeader()
    lblCapa.Caption = ""
    lblNumMalote.Caption = ""
    lblLote.Caption = ""
    lblOcorrencia.Caption = ""
End Sub
Private Function PossuiDoctoRecaptura() As Boolean

  Dim X As Integer

  PossuiDoctoRecaptura = False

  'Verificar se existe algum documento para recaptura
  If lstDocto.ListCount > 0 Then
    For X = 0 To lstDocto.ListCount - 1
      If aDoc(X + 1).Status = "A" Then
        'Documento para Recaptura
        PossuiDoctoRecaptura = True
        Exit Function
      End If
      DoEvents
    Next X
  End If
End Function

Private Sub LimparListas()
    lstCapa.Clear
    lstDocto.Clear
End Sub

Private Sub Preenche_lstCapa()

    Dim Count As Integer

    lstCapa.Clear
    For Count = 1 To m_CountCapa
        If Len(Format(aCapa(Count).Diferenca, ".00")) > 9 Then
            lstCapa.AddItem aCapa(Count).Capa & Space(11) & "xxxxx"
        Else
            lstCapa.AddItem aCapa(Count).Capa & Space(15 - Len(aCapa(Count).Capa)) & Space(9 - Len(Format(aCapa(Count).Diferenca, ".00"))) & Format(aCapa(Count).Diferenca, ".00")
        End If
    Next
End Sub
Private Sub Preenche_lstDocto(ByVal Indice As Integer)
    Dim Linha As String
    Dim Count As Integer

    LimparImagem

    txtValor.Text = ""
    lstDocto.Clear
    For Count = 1 To m_CountDocto
        If chkNaoVinculados.Value = 0 Or aDoc(Count).Vinculo = 0 Then
            If chkOcorrencia.Value = 1 Or (aDoc(Count).Status <> "D" And aDoc(Count).Status <> "F") Then

                Linha = Format(Count, "0000") & Space(2)
                Linha = Linha & Format(aDoc(Count).Vinculo, "0000000") & Space(2)
                Linha = Linha & IIf(aDoc(Count).Ocorrencia = 0, " ", "S") & Space(4)
                'Inclusão de linha - Referente a status 'Para/Em Recaptura'
                Linha = Linha & IIf(aDoc(Count).Status = "A", "S", " ") & Space(4)
                Select Case aDoc(Count).TipoDocto
                    Case 1
                        Linha = Linha & "CAPA     " & Space(1)
                    Case 2, 3       ' Depósito
                        Linha = Linha & "DEPOSITO " & Space(1)
                    Case 4          ' ADCC
                        Linha = Linha & "DEBITO CC" & Space(1)
                    Case 5, 6, 7    ' Cheque
                        Linha = Linha & "CHEQUE   " & Space(1)
                    Case 37         ' oct
                        Linha = Linha & "OCT      " & Space(1)
                    Case 32, 34
                        Linha = Linha & "AJ. CRED." & Space(1)
                    Case 33, 38
                        Linha = Linha & "AJ. DEB. " & Space(1)
                    Case 41
                        Linha = Linha & "LANCTO   " & Space(1)
                    Case 42
                        Linha = Linha & "AJ. REC. " & Space(1)
                    Case 43
                        Linha = Linha & "AJ. DESP." & Space(1)
                    Case Else       ' Pagamento
                        Linha = Linha & "PAGAMENTO" & Space(1)
                End Select
                If aDoc(Count).TipoDocto = 1 Then
                    Linha = Linha & Space(22)
                Else
                    Linha = Linha & FormataValor(aDoc(Count).Valor, 20) & Space(2)
                End If
                Linha = Linha & Format(aDoc(Count).IdDocto, "0000000000")
                lstDocto.AddItem Linha
            End If
        End If
    Next
    If lstDocto.ListCount > 0 Then
        If Val(Indice) <> 0 Then
            lstDocto.ListIndex = Indice
        Else
            lstDocto.ListIndex = 0
        End If
    Else
        txtValor.Enabled = False
        cmdAlterar.Enabled = False
    End If
End Sub
Private Function ObtemCapas(Optional pIdTitulo As Integer) As Boolean
    On Error GoTo ErroGetCapa
    rdoErrors.Clear

    Screen.MousePointer = vbHourglass

    Erase aCapa
    m_CountCapa = 0

    Set qryGetCapaProvaZero = Geral.Banco.CreateQuery("", "{Call GetCapaProvaZero (?,?,?)}")

    qryGetCapaProvaZero.rdoParameters(0) = Geral.DataProcessamento
    qryGetCapaProvaZero.rdoParameters(1) = Geral.Intervalo
    
    If pIdTitulo <> 0 Then
        qryGetCapaProvaZero.rdoParameters(2) = pIdTitulo
    End If
    Set rsCapa = qryGetCapaProvaZero.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    If rsCapa.EOF Then
        rsCapa.Close
        ObtemCapas = False
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    ReDim aCapa(rsCapa.RowCount)
    While Not rsCapa.EOF
        m_CountCapa = m_CountCapa + 1
        m_Capa.IdCapa = rsCapa!IdCapa
        m_Capa.IdLote = rsCapa!IdLote
        m_Capa.IdEnv_Mal = rsCapa!IdEnv_Mal
        m_Capa.AgOrig = rsCapa!AgOrig
        m_Capa.Capa = rsCapa!Capa
        m_Capa.NumMalote = rsCapa!Num_Malote
        m_Capa.Diferenca = rsCapa!Diferenca
        aCapa(m_CountCapa) = m_Capa
        rsCapa.MoveNext
    Wend
    rsCapa.Close
    ObtemCapas = True
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    Exit Function
    
ErroGetCapa:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro na obtenção de envelope/malote para Prova Zero.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    m_Busy = False
    Unload Me

End Function

Private Sub ObtemDocumentos(ByVal IdCapa As Long)
    On Error GoTo ErroGetDocto
    rdoErrors.Clear
    
    Screen.MousePointer = vbHourglass
    
    Erase aDoc
    m_CountDocto = 0
    qryGetDocumentoProvaZero.rdoParameters(0) = Geral.DataProcessamento
    qryGetDocumentoProvaZero.rdoParameters(1) = IdCapa
    Set rsDoc = qryGetDocumentoProvaZero.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    ReDim aDoc(rsDoc.RowCount)
    While Not rsDoc.EOF
        m_CountDocto = m_CountDocto + 1
        m_Doc.IdDocto = rsDoc!IdDocto
        m_Doc.TipoDocto = rsDoc!TipoDocto
        m_Doc.Ocorrencia = rsDoc!Ocorrencia
        m_Doc.RetornoTransacao = rsDoc!RetornoTransacao
        m_Doc.Leitura = rsDoc!Leitura
        m_Doc.Frente = rsDoc!Frente
        m_Doc.Verso = rsDoc!Verso
        m_Doc.Status = rsDoc!Status
        m_Doc.Ordem = rsDoc!Ordem
        m_Doc.Vinculo = rsDoc!Vinculo
        m_Doc.Valor = rsDoc!Valor
        m_Doc.Conferido = False
        aDoc(m_CountDocto) = m_Doc
        rsDoc.MoveNext
    Wend
    rsDoc.Close
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErroGetDocto:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro na obtenção de documentos para Prova Zero.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    m_Busy = False
    Unload Me
    
End Sub

Private Function VerificaCapaDisponivel(ByVal IdCapa As Long) As Boolean
    On Error GoTo ErroVerificaCapa
    rdoErrors.Clear
    
    Screen.MousePointer = vbHourglass
    
    With qryVerificaCapaDisponivel
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = IdCapa
        .rdoParameters(3) = "4"
        .rdoParameters(4) = "G"
        .rdoParameters(5) = Geral.Intervalo
        .Execute
        Screen.MousePointer = vbDefault
        If .rdoParameters(0) = 0 Then
            VerificaCapaDisponivel = True
        ElseIf .rdoParameters(0) = 1 Then
            VerificaCapaDisponivel = False
            MsgBox "Este Envelope/Malote não está mais disponível por já ter sido tratado ou porque esta sendo tratado por outra estação.", vbInformation + vbOKOnly, App.Title
        Else
            VerificaCapaDisponivel = False
            MsgBox "Erro. Não foi possível obter o Status do Envelope/Malote.", vbInformation + vbOKOnly, App.Title
        End If
    End With
    On Error GoTo 0
    Exit Function
    
ErroVerificaCapa:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro. Não foi possível obter o Status do Envelope/Malote.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    m_Busy = False
    Unload Me
    
End Function

Private Function CalculaValores() As Boolean
    Dim Count As Integer
   
    m_ValContas = 0
    m_ValCheques = 0
    m_ValDiferenca = 0

    For Count = 1 To m_CountDocto
        If chkNaoVinculados.Value = 1 Then
            If (aDoc(Count).Status <> "D" And aDoc(Count).Status <> "F" And _
                aDoc(Count).Vinculo = 0) Then
                If (aDoc(Count).TipoDocto >= 4 And aDoc(Count).TipoDocto <= 7) Or (aDoc(Count).TipoDocto = 41) Then
                    m_ValCheques = m_ValCheques + aDoc(Count).Valor
                Else
                    m_ValContas = m_ValContas + aDoc(Count).Valor
                End If
            End If
        Else
            If (aDoc(Count).Status <> "D" And aDoc(Count).Status <> "F") Then
                If (aDoc(Count).TipoDocto >= 4 And aDoc(Count).TipoDocto <= 7) Or _
                    (aDoc(Count).TipoDocto = 41) Or _
                    (aDoc(Count).TipoDocto = 43 Or aDoc(Count).TipoDocto = 45) Or _
                    (aDoc(Count).TipoDocto = 33 Or aDoc(Count).TipoDocto = 38) Then
                    m_ValCheques = m_ValCheques + aDoc(Count).Valor
                Else
                    m_ValContas = m_ValContas + aDoc(Count).Valor
                End If
            End If
        End If
    Next

    m_ValDiferenca = m_ValCheques - m_ValContas

    If m_ValDiferenca <> 0 Then
        lblValorContas.Caption = FormataValor(m_ValContas, 20)
        lblValorCheques.Caption = FormataValor(m_ValCheques, 20)
        lblValorDiferenca.Caption = FormataValor(m_ValDiferenca, 21)
        'lstCapa.List(lstCapa.ListIndex + 1) = aCapa(lstCapa.ListIndex + 1).capa & Space(14 - Len(aCapa(lstCapa.ListIndex + 1).capa)) & Space(9 - Len(Format(lblValorDiferenca.Caption, ".00"))) & Format(lblValorDiferenca.Caption, ".00")
        CalculaValores = False
    Else
        'Verificar se a capa possui algum documento para recaptura
        If PossuiDoctoRecaptura Then
            AtualizaStatusCapa m_IdCapa, "A"
            m_IdCapa = 0
            If lstCapa.ListIndex < lstCapa.ListCount - 1 Then
                lstCapa.ListIndex = lstCapa.ListIndex + 1
            Else
                CmdAtualizar_Click
            End If
            lstDocto.SetFocus
            CalculaValores = True
        Else
            If Not m_Alteracao Then
                AtualizaStatusCapa m_IdCapa, "9"
                m_IdCapa = 0
            End If
            If lstCapa.ListIndex < lstCapa.ListCount - 1 Then
                lstCapa.ListIndex = lstCapa.ListIndex + 1
            Else
                CmdAtualizar_Click
            End If
            CalculaValores = True
        End If
    End If
End Function
Private Function Indice(ByVal IdDocto As Long) As Integer
    Dim Count As Integer
    For Count = 1 To m_CountDocto
        If aDoc(Count).IdDocto = IdDocto Then
            Indice = Count
            Exit Function
        End If
    Next
    Indice = 0
End Function

Private Sub MostraImagem()
    Dim i As Integer
    i = Indice(Val(Right(lstDocto.List(lstDocto.ListIndex), 10)))
    
    aDoc(i).Conferido = True
    
    hCtl = Lead1.hwnd
    '''''''''''''''''''''''''''
    ' mostra imagem escolhida '
    '''''''''''''''''''''''''''
    On Error GoTo ErroImagem
    With Lead1
       .Tag = "F"
       .AutoRepaint = False
       If Geral.VIPSDLL = eDllProservi Then
         .Load Geral.DiretorioImagens & aDoc(i).Frente, 0, 0, 1
       Else
         .Load Geral.DiretorioImagens & Format(aCapa(lstCapa.ListIndex + 1).IdLote, "000000000") & "\" & aDoc(i).Frente, 0, 0, 1
       End If

       ' se imagem for da ls500, deixar mais escura
       If aDoc(i).Ordem <> "2" Then
          .Intensity 220
       Else
          .Intensity 140
       End If
       ' se imagem for do canon, diminui em 50% o tamanho
       If aDoc(i).Ordem <> "1" Then
          .PaintZoomFactor = 100
       Else
          .PaintZoomFactor = 50
       End If
       .AutoRepaint = True
    End With
    FrmImagem.Visible = True
    
    'posiciona imagem sempre no começo
    Ret = SendMessage(hCtl, WM_VSCROLL, SB_TOP, 0)
    Ret = SendMessage(hCtl, WM_HSCROLL, SB_BOTTOM, 0)
    
    cmdAuditoria.Enabled = True
    cmdZoomMais.Enabled = True
    cmdZoomMenos.Enabled = True
    cmdRotacao.Enabled = True
    cmdInverteCor.Enabled = True
    cmdFrenteVerso.Enabled = True
    On Error GoTo 0
    DoEvents
    Exit Sub
    
ErroImagem:
    cmdAuditoria.Enabled = False
    cmdZoomMais.Enabled = False
    cmdZoomMenos.Enabled = False
    cmdRotacao.Enabled = False
    cmdInverteCor.Enabled = False
    cmdFrenteVerso.Enabled = False
    FrmImagem.Visible = False
    MsgBox "Não foi possível exibir imagem do documento, imagem não encontrada", vbExclamation + vbOKOnly, App.Title

End Sub

Private Sub LimparImagem()

    FrmImagem.Visible = False
    cmdAuditoria.Enabled = False
    cmdZoomMais.Enabled = False
    cmdZoomMenos.Enabled = False
    cmdRotacao.Enabled = False
    cmdInverteCor.Enabled = False
    cmdFrenteVerso.Enabled = False
End Sub
Private Sub MostraValor()
    Dim i As Integer
    i = Indice(Val(Right(lstDocto.List(lstDocto.ListIndex), 10)))
    txtValor.Text = RetiraPonto(Trim(FormataValor(aDoc(i).Valor, 21)))
    If aDoc(i).Status = "D" Or aDoc(i).Status = "F" Or aDoc(i).TipoDocto = 1 Then
        txtValor.Enabled = False
        cmdAlterar.Enabled = False
    Else
        txtValor.Enabled = True
        cmdAlterar.Enabled = True
    End If
End Sub
Private Function AtualizaStatusCapa(ByVal IdCapa As Long, ByVal Status As String) As Boolean
    On Error GoTo ErroAtualizaStatus
    rdoErrors.Clear

    AtualizaStatusCapa = True
    Screen.MousePointer = vbHourglass

    With qryAtualizaStatusCapa
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = IdCapa
        .rdoParameters(3) = Status
        .Execute
        If .rdoParameters(0) <> 0 Then
            AtualizaStatusCapa = False
            Screen.MousePointer = vbDefault
            MsgBox "Erro na atualização do status do envelope/malote.", vbCritical + vbOKOnly, App.Title
        Else
            'Verificar o Status
            If Status = "5" Then
                'Enviar para Ilegiveis
                Call GravaLog(IdCapa, 0, 63)
            ElseIf Status = "8" Then
                'Enviar para Vinculo Automatico
                Call GravaLog(IdCapa, 0, 64)
            ElseIf Status = "9" Then
                'Enviar para Vinculo Automatico
                Call GravaLog(IdCapa, 0, 65)
            ElseIf Status = "A" Then
                'Enviar para Recaptura
                Call GravaLog(IdCapa, 0, 67)
            ' ElseIf Status = "G" Then
            '    Call GravaLog(IdCapa, 0, 193)
            ' ElseIf Status = "4" Then
            '    Call GravaLog(IdCapa, 0, 194)
            End If
        End If
    End With

    Screen.MousePointer = vbDefault
    On Error GoTo 0
    Exit Function
    
ErroAtualizaStatus:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro na atualização do status do envelope/malote.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    m_Busy = False
    Unload Me
    
End Function

Private Function AtualizaValorDocumento(ByVal IdDocto As Long, _
                ByVal TipoDocto As Integer, ByVal Valor As Currency) As Boolean
    On Error GoTo ErroAtualizaValor
    rdoErrors.Clear

    AtualizaValorDocumento = True
    Screen.MousePointer = vbHourglass
    
    With qryAtualizaValorDocumento
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = IdDocto
        .rdoParameters(3) = TipoDocto
        .rdoParameters(4) = Valor
        .Execute
        If .rdoParameters(0) <> 0 Then
            AtualizaValorDocumento = False
            Screen.MousePointer = vbDefault
            MsgBox "Erro na atualização do valor do documento.", vbCritical + vbOKOnly, App.Title

            lstDocto.SetFocus
        End If
    End With
    Screen.MousePointer = vbDefault
    On Error GoTo 0
    Exit Function

ErroAtualizaValor:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro na atualização do valor do documento.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    m_Busy = False
    Unload Me
    
End Function

Private Sub ObtemOcorrencia()

    Dim i               As Integer
    Dim Ocorrencia      As Long
    Dim sOcorrencia     As String
    
    i = Indice(Val(Right(lstDocto.List(lstDocto.ListIndex), 10)))
    If aDoc(i).Ocorrencia = 0 Then
        lblOcorrencia.Caption = ""
        Exit Sub
    End If
    
    On Error GoTo ErroOcorrencia
    rdoErrors.Clear
    
    Screen.MousePointer = vbHourglass
    
    'Verificar se a ocorrencia começa com 999
    If Left(aDoc(i).Ocorrencia, 3) = "999" Then
        If aDoc(i).RetornoTransacao > 0 Then
            Call ObtemRetornoTransacao(aDoc(i).RetornoTransacao, sOcorrencia)
        Else
            sOcorrencia = "Erro operacional."
        End If
        lblOcorrencia.Caption = sOcorrencia
        
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        '3 Posicoes
        If Right(Trim(aDoc(i).Ocorrencia), 2) = "00" Then
            'Ocorrencia atualizada pelo robo
            Ocorrencia = Val(Trim(aDoc(i).Ocorrencia)) / 100
        Else
            'Ocorrencia gerada pelo sistema
            Ocorrencia = Val(Trim(aDoc(i).Ocorrencia))
        End If
    End If
    
    'qryGetocorrencia.rdoParameters(0) = aDoc(i).Ocorrencia
    qryGetocorrencia.rdoParameters(0) = Ocorrencia
    
    Set RsOcorrencia = qryGetocorrencia.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    If RsOcorrencia.EOF Then
        Screen.MousePointer = vbDefault
        MsgBox "Código da Ocorrência não existe: " & str(aDoc(i).Ocorrencia) & ".", vbExclamation + vbOKOnly, App.Title
    Else
        lblOcorrencia.Caption = "Ocorrência: " & RsOcorrencia!Descricao
    End If
    
    RsOcorrencia.Close
    Screen.MousePointer = vbDefault
    On Error GoTo 0
    Exit Sub
    
ErroOcorrencia:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro na obtenção da Ocorrência do documento.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    m_Busy = False
    Unload Me
    
End Sub

Private Sub HabilitaTimerPesquisa()

  'Esta Função irá verificar a existência de documentos Ilegíveis a cada x segundos
  'de acordo com o campo PARAMETRO.TmAtualizacao
  FrmPesquisa.Visible = True
  tmrPesquisa.Enabled = True
  Progress.Value = 0
  
  
  ''''''''''''''''''''''''''''''''''''''''''
  'Grava Log MDI - Inicio Aguarda documento'
  ''''''''''''''''''''''''''''''''''''''''''
  Call GravaLog(0, 0, 254)
  
  
End Sub

Private Sub chkFiltro_Click()
    
    If m_IsEvent Then Exit Sub

    If chkFiltro.Value = vbChecked Then
        
        '''''''''''''''''''''''''''''''''''''''''''
        'Retira capa selecionada de 'Em Prova-Zero'
        '''''''''''''''''''''''''''''''''''''''''''
        If m_IdCapa > 0 Then
            If m_Alteracao Then
                If Not AtualizaStatusCapa(m_IdCapa, "8") Then
                    m_Busy = False
                    m_IdCapa = 0
                    chkFiltro.Value = vbUnchecked
                    Exit Sub
                End If
            Else
                If Not AtualizaStatusCapa(m_IdCapa, "4") Then
                    m_Busy = False
                    m_IdCapa = 0
                    chkFiltro.Value = vbUnchecked
                    Exit Sub
                End If
            End If
        End If
        
        LimparValores
        LimparHeader
        LimparListas
        
        If Not ObtemCapas(31) Then
            MsgBox "Não existem Envelopes/Malotes com pendência de Prova Zero.", vbExclamation + vbOKOnly, App.Title
            Call CmdAtualizar_Click
            Exit Sub
        End If
        
        Call Preenche_lstCapa
        
        lstCapa.ListIndex = 0
        
    Else
        ''''''''''''''''''''
        'Seleção sem filtro'
        ''''''''''''''''''''
        Call CmdAtualizar_Click
    End If

End Sub

Private Sub chkNaoVinculados_Click()
    If Not m_FirstActivate Then
        Preenche_lstDocto (0)
        CalculaValores
        lstDocto.SetFocus
    End If
End Sub
Private Sub chkOcorrencia_Click()
    If Not m_FirstActivate Then
        Preenche_lstDocto (0)
        lstDocto.SetFocus
    End If
End Sub
Private Sub cmdAlterar_Click()
   If FrmPesquisa.Visible = True Or m_IdCapa = 0 Then Exit Sub
   AlterarValor (True)
End Sub
Private Sub MarcaDoctoRecaptura()

    Dim X As Integer

    For X = 0 To lstDocto.ListCount - 1
        If lstDocto.Selected(X) = True Then
            'Atualizar o Status do Documento para 'A' (Recaptura)
            Call AtualizaStatusDocumento(aDoc(X + 1).IdDocto, "A")
            aDoc(X + 1).Status = "A"
        End If
        DoEvents
    Next X

    Call Preenche_lstDocto(lstDocto.ListIndex)

    lstDocto.SetFocus
End Sub
Private Function AtualizaStatusDocumento(ByVal IdDocto As Long, ByVal Status As String) As Boolean

    On Error GoTo AtualizaStatusDocumento_Err

    AtualizaStatusDocumento = False

    Set qryAtualizaStatusDocumento = Geral.Banco.CreateQuery("", "{? = call AtualizaStatusDocumento (?,?,?)}")
    With qryAtualizaStatusDocumento
       .rdoParameters(0).Direction = rdParamReturnValue     'Parametro de Retorno
       .rdoParameters(1) = Geral.DataProcessamento          'Data Processamento
       .rdoParameters(2) = IdDocto                          'IdDocto
       .rdoParameters(3) = Status                           'Status do Documento
       .Execute
    End With

    If qryAtualizaStatusDocumento(0).Value = "1" Then
       MsgBox "Ocorreu um erro ao Atualizar o Status do Documento.", vbInformation + vbOKOnly, App.Title
       Exit Function
    End If

    AtualizaStatusDocumento = True

    Exit Function

AtualizaStatusDocumento_Err:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Atualizar o Status do documento.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function

Private Sub AlterarValor(ByVal Gravar As Boolean)

    Dim i As Integer
    Dim j As Integer

    j = lstDocto.ListIndex
    i = Indice(Val(Right(lstDocto.List(j), 10)))

    If aDoc(i).TipoDocto = 4 Then
        'Verificar se o valor da ADCC informada é maior que o limite permitido (Parâmetro)
        If Val(InserePonto(txtValor.Text)) > Val(InserePonto(Geral.ValorMaxADCC * 100)) Then
            MsgBox "Valor informado na Autorização maior que o limite permitido.", vbInformation, App.Title
            txtValor.SetFocus
            Exit Sub
        End If
    End If

    Select Case aDoc(i).TipoDocto
        Case 1
            'Verificar se foi selecionado Capa de Malote/Envelope
            MsgBox "Não é permitido alterar Valor de CAPA.", vbInformation + vbOKOnly, App.Title
            Exit Sub
        
        Case 15
            'Verificar se foi selecionado um DARM
            MsgBox "Não é permitido alterar Valores de DARMs.", vbInformation + vbOKOnly, App.Title
            Exit Sub

        Case 20, 21, 22, 23
            'Verificar se foi selecionado uma concessionaria
            MsgBox "Não é permitido alterar Valores de Concessionárias.", vbInformation + vbOKOnly, App.Title
            Exit Sub

        Case 32, 33, 34, 38, 42, 43
           'Verificar se foi selecionado um ajuste
            MsgBox "Não é permitido alterar valores de Ajustes.", vbInformation + vbOKOnly, App.Title
            Exit Sub
        
        Case 41
           'Verificar se foi selecionado um L.I.
            MsgBox "Não é permitido alterar valor de Lancamento Interno.", vbInformation + vbOKOnly, App.Title
            Exit Sub

        Case 13, 14, 16, 17, 18, 28, 29, 30, 31, 35, 36
            'Verificar se foi selecionado algum produto que possui tela de complementação no Prova Zero
            'MsgBox "", vbInformation + vbOKOnly, App.Title
            If Gravar = True Then Exit Sub

    End Select

    If (aDoc(i).Valor <> Val(txtValor.Text) / 100) Then
        If Gravar Then
            If Not AtualizaValorDocumento(aDoc(i).IdDocto, aDoc(i).TipoDocto, Val(txtValor.Text) / 100) Then
                Exit Sub
            End If
        End If
        m_Alteracao = True
        aDoc(i).Valor = Val(txtValor.Text) / 100
        If CalculaValores Then
            Exit Sub
        End If

        'Grava Log
        If lstCapa.ListIndex <> -1 Then
          Call GravaLog(aCapa(lstCapa.ListIndex + 1).IdCapa, aDoc(lstDocto.ListIndex).IdDocto, 60)
        End If

        Preenche_lstDocto (lstDocto.ListIndex)

    End If
    If lstDocto.ListCount > j + 1 Then
        lstDocto.Selected(j + 1) = True
        Preenche_lstDocto (lstDocto.ListIndex)
    Else
       If lstDocto.ListIndex <> -1 Then
          lstDocto.Selected(j) = True
          Preenche_lstDocto (lstDocto.ListIndex)
       End If
    End If

    lstDocto.SetFocus
    DoEvents
End Sub

Private Sub CmdAtualizar_Click()

    m_IsEvent = True
    
    If m_IdCapa > 0 Then
        If m_Alteracao Then
            If Not AtualizaStatusCapa(m_IdCapa, "8") Then
                m_Busy = False
                m_IsEvent = False
                m_IdCapa = 0
                Exit Sub
            End If
        Else
            If Not AtualizaStatusCapa(m_IdCapa, "4") Then
                m_Busy = False
                m_IdCapa = 0
                m_IsEvent = False
                Exit Sub
            End If
        End If
    End If
    LimparValores
    LimparHeader
    LimparListas

    If Not ObtemCapas Then
        MsgBox "Não existem Envelopes/Malotes com pendência de Prova Zero.", vbExclamation + vbOKOnly, App.Title
        m_IdCapa = 0
        m_IsEvent = False
        LimparImagem
        HabilitaTimerPesquisa
        Exit Sub
    Else
        tmrPesquisa.Enabled = False
        FrmPesquisa.Visible = False
    End If
    
    chkFiltro.Value = vbUnchecked
    
    Preenche_lstCapa
    lstCapa.ListIndex = 0
    m_IsEvent = False
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

    If FrmPesquisa.Visible = True Then Exit Sub

    strCommand = Space(254)
    GetWindowsDirectory strCommand, 254
    strCommand = Trim(strCommand)
    strCommand = Left(strCommand, Len(strCommand) - 1) & "\calc.exe"

    WinExec strCommand, 9
End Sub
Private Sub cmdCancelar_Click()

    txtNumEnvMal.Text = ""
    frmLocalizar.Visible = False
    lstDocto.SetFocus

End Sub
Private Sub cmdEncerrar_Click()

    Dim strDoctos As String
    Dim Count As Integer

    If FrmPesquisa.Visible = True Or m_IdCapa = 0 Then Exit Sub

    'Verificar se a capa possui algum documento para recaptura
    If PossuiDoctoRecaptura Then
        AtualizaStatusCapa m_IdCapa, "A"
        m_IdCapa = 0
        If lstCapa.ListIndex < lstCapa.ListCount - 1 Then
            lstCapa.ListIndex = lstCapa.ListIndex + 1
        Else
            CmdAtualizar_Click
        End If
        lstDocto.SetFocus
        Exit Sub
    End If

    strDoctos = ""
    For Count = 1 To m_CountDocto
        If Not aDoc(Count).Conferido And aDoc(Count).Status <> "D" And aDoc(Count).Status <> "F" And aDoc(Count).Vinculo = 0 Then
            strDoctos = strDoctos & Format(Count, "0000") & ", "
        End If
    Next
    If Len(strDoctos) > 0 Then
        MsgBox "Os seguintes documentos não foram conferidos: " & _
            Left(strDoctos, Len(strDoctos) - 2) & "." & vbLf & _
            "Para encerrar um Envelope/Malote é necessário conferir todos os documentos.", _
            vbInformation + vbOKOnly, App.Title
        lstDocto.SetFocus
        Exit Sub
    End If

    AtualizaStatusCapa m_IdCapa, "9"
    m_IdCapa = 0
    If lstCapa.ListIndex < lstCapa.ListCount - 1 Then
        lstCapa.ListIndex = lstCapa.ListIndex + 1
    Else
        CmdAtualizar_Click
    End If
End Sub
Private Sub CmdExcluir_Click()
    
    If FrmPesquisa.Visible = True Or m_IdCapa = 0 Then Exit Sub

    'Verificar se a capa pode ser excluida
    If Not VerificaDoctosExcluidosCapa(aCapa(lstCapa.ListIndex + 1).IdCapa) Then
      MsgBox "Não é permitido excluir Envelopes / Malotes em que todos os documentos possuam ocorrência.", vbInformation + vbOKOnly, App.Title
      Exit Sub
    End If
    
    Load MotivoExclusao
    
    MotivoExclusao.LblValorEnv_Mal.Caption = aCapa(lstCapa.ListIndex + 1).Capa
    MotivoExclusao.LblValorEnv_Mal.Tag = aCapa(lstCapa.ListIndex + 1).IdCapa
    
    MotivoExclusao.Show vbModal, Me
    
    If MotivoExclusao.Result Then
        'Gravar Log
        Call GravaLog(aCapa(lstCapa.ListIndex + 1).IdCapa, 0, 62)

        m_IdCapa = 0
        If lstCapa.ListIndex < lstCapa.ListCount - 1 Then
            lstCapa.ListIndex = lstCapa.ListIndex + 1
        Else
            lstCapa_Click
        End If
    Else
        CmdAtualizar_Click
    End If
    
    Unload MotivoExclusao
    
End Sub

Private Sub CmdFechar_Click()
    Unload Me
End Sub

Private Sub CmdFecharPesquisa_Click()
    '''''''''''''''''''''''''''''''''''''''
    'Grava Log MDI - Fim Aguarda documento'
    '''''''''''''''''''''''''''''''''''''''
    Call GravaLog(0, 0, 255)
    
    CmdFechar_Click
End Sub

Public Sub cmdFrenteVerso_Click()
    Dim i As Integer
    If m_Busy Then
        Exit Sub
    End If
    If Not FrmImagem.Visible Then
        Exit Sub
    End If
    m_Busy = True
    
    On Error GoTo ErroImagem
    
    i = Indice(Val(Right(lstDocto.List(lstDocto.ListIndex), 10)))
    'só muda para frente/verso qdo docto vem da Ls500 e da Vips,
    'poi, o canon não gera verso.
    If (aDoc(i).Ordem = "0") Or (aDoc(i).Ordem = "2") Then
        If Lead1.Tag = "V" Then
            Lead1.Tag = "F"     'se verso, mostrar frente
            With Lead1
               .AutoRepaint = False
               If Geral.VIPSDLL = eDllProservi Then
                 .Load Geral.DiretorioImagens & aDoc(i).Frente, 0, 0, 1
               Else
                 .Load Geral.DiretorioImagens & Format(aCapa(lstCapa.ListIndex + 1).IdLote, "000000000") & "\" & aDoc(i).Frente, 0, 0, 1
               End If

               'se ls500 mostrar mais escuro
               If (aDoc(i).Ordem = "2") Then
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
                 .Load Geral.DiretorioImagens & aDoc(i).Verso, 0, 0, 1
               Else
                 .Load Geral.DiretorioImagens & Format(aCapa(lstCapa.ListIndex + 1).IdLote, "000000000") & "\" & aDoc(i).Verso, 0, 0, 1
               End If

               If (aDoc(i).Ordem = "2") Then
                  .Intensity 140
               Else
                  .Intensity 220
               End If
               .PaintZoomFactor = 100
               .AutoRepaint = True
            End With
        End If
    End If
    m_Busy = False
    Exit Sub
    
ErroImagem:
    m_Busy = False
    cmdAuditoria.Enabled = False
    cmdZoomMais.Enabled = False
    cmdZoomMenos.Enabled = False
    cmdRotacao.Enabled = False
    cmdInverteCor.Enabled = False
    cmdFrenteVerso.Enabled = False
    FrmImagem.Enabled = False
    MsgBox "Não foi possível exibir imagem do documento, imagem não encontrada", vbExclamation + vbOKOnly, App.Title
    
End Sub

Private Sub cmdIlegiveis_Click()

    Dim rst         As RDO.rdoResultset
    Dim sStr        As String

    If FrmPesquisa.Visible = True Or m_IdCapa = 0 Then Exit Sub
    
    AtualizaStatusCapa m_IdCapa, "5"
    
    '''''''''''''''''''''''''''''''
    'Verifica se existe comentario'
    '''''''''''''''''''''''''''''''
    Set rst = GetControleCapa(Geral.DataProcessamento, m_IdCapa)
    
    sStr = ""
    If Not rst.EOF() Then
        sStr = rst!Comentario
    End If
    '''''''''''''''''''''''''''''''''
    'Insere registro no ControleCapa'
    '''''''''''''''''''''''''''''''''
    If Not InsereControleCapa(Geral.DataProcessamento, m_IdCapa, sStr, 10) Then
        MsgBox "Não foi possível inserir o Controle de Capa.", vbExclamation
    End If
        
    m_IdCapa = 0
    If lstCapa.ListIndex < lstCapa.ListCount - 1 Then
        lstCapa.ListIndex = lstCapa.ListIndex + 1
    Else
        CmdAtualizar_Click
    End If

    lstDocto.SetFocus
End Sub

Public Sub cmdInverteCor_Click()
    If m_Busy Then
        Exit Sub
    End If
    If Not FrmImagem.Visible Then
        Exit Sub
    End If
    m_Busy = True
    Lead1.Invert
    m_Busy = False
End Sub

Private Sub CmdLocalizar_Click()

    If FrmPesquisa.Visible = False Then
        frmLocalizar.Visible = True
        txtNumEnvMal.SetFocus
    End If
End Sub

Private Sub cmdProcurar_Click()

    Dim iIndex                   As Integer
    Dim Encontrou                As Boolean
    Dim qryGetDescStatusCapa     As rdoQuery 'Pega descricao do status da capa
    Dim sCapa                    As String
    
    If Not IsNumeric(txtNumEnvMal.Text) Then
        MsgBox "Capa inválida.", vbInformation
        txtNumEnvMal.SetFocus
        Exit Sub
    End If

    Set qryGetDescStatusCapa = Geral.Banco.CreateQuery("", "{Call GetDescStatusCapa(?,?,?)}")
    
    Encontrou = False
    If Trim(txtNumEnvMal.Text) <> "" Then
    
        If IsNumeric(txtNumEnvMal.Text) Then
            'Atualizar a lista de capas antes da pesquisa
            Call CmdAtualizar_Click
    
            'Verificar se a capa informada está na lista de capas
            For iIndex = 0 To lstCapa.ListCount - 1
                If CDbl(Left(lstCapa.List(iIndex), 14)) = CDbl(txtNumEnvMal.Text) Then
                    lstCapa.Selected(iIndex) = True
                    Encontrou = True
                    Exit For
                End If
                DoEvents
            Next iIndex
        End If
    End If

    sCapa = txtNumEnvMal.Text
    'Verificar se encontrou capa
    If m_IdCapa > 0 And Not Encontrou Then
    
        With qryGetDescStatusCapa
            .rdoParameters(0) = Geral.DataProcessamento
            .rdoParameters(1) = CDbl(Left(sCapa, 14))
            .rdoParameters(2).Direction = rdParamOutput
            .Execute
            
            If Trim(.rdoParameters(2).Value) <> "" Then
                MsgBox .rdoParameters(2).Value, vbInformation
            Else
                MsgBox "Capa não Encontrada.", vbInformation
            End If
            
        End With

        If m_Alteracao Then
            If Not AtualizaStatusCapa(m_IdCapa, "8") Then
                m_Busy = False
                m_IdCapa = 0
                Exit Sub
            End If
        Else
            If Not AtualizaStatusCapa(m_IdCapa, "4") Then
                m_Busy = False
                m_IdCapa = 0
                Exit Sub
            End If
        End If

        lstCapa.ListIndex = -1
        m_IdCapa = 0
        m_Busy = False
        m_CountDocto = 0
        Preenche_lstDocto (0)
        LimparValores

        'Desabilitando as informações de Malote
        lblCapa.Caption = ""
        picNumMalote.Visible = False
        lblNumMalote.Visible = False
        lblNumMalote.Caption = ""
    End If

    txtNumEnvMal.Text = ""
    frmLocalizar.Visible = False

End Sub
Private Sub CmdRecaptura_Click()

    Dim j As Integer, i As Integer
    
    If FrmPesquisa.Visible = True Or m_IdCapa = 0 Then Exit Sub
    
    j = lstDocto.ListIndex
    i = Indice(Val(Right(lstDocto.List(j), 10)))
    
    If aDoc(i).TipoDocto = 1 Then
        MsgBox "O documento CAPA não pode ser marcado para recaptura.", vbOKOnly, App.Title
        Exit Sub
    End If
    
    If MsgBox("O documento selecionado será marcado para recaptura. Confirma ?", vbYesNo + vbQuestion) = vbYes Then
        Screen.MousePointer = vbHourglass
        Call MarcaDoctoRecaptura
        Screen.MousePointer = vbDefault
    End If
End Sub
Public Sub cmdRotacao_Click()
    If m_Busy Then
        Exit Sub
    End If
    If Not FrmImagem.Visible Then
        Exit Sub
    End If
    m_Busy = True
    Lead1.FastRotate 90
    m_Busy = False
End Sub
Private Sub CmdTrocaOrderm_Click()
'* Envia capa e documentos para Modulo Troca de Orderm *'

   If lstCapa.ListIndex <> -1 Then
      If aCapa(lstCapa.ListIndex + 1).IdCapa <> 0 Then
         'Gravar Log
         Call GravaLog(aCapa(lstCapa.ListIndex + 1).IdCapa, 0, 6)

         Load frmTrocarOrdemDocumento
         frmTrocarOrdemDocumento.setIdCapaDefault aCapa(lstCapa.ListIndex + 1).IdCapa
         frmTrocarOrdemDocumento.Show vbModal
         lstCapa_Click
      End If
   End If

End Sub

Public Sub cmdZoomMais_Click()
    If m_Busy Then
        Exit Sub
    End If
    If Not FrmImagem.Visible Then
        Exit Sub
    End If
    m_Busy = True
    If Lead1.PaintZoomFactor <= 400 Then
        Lead1.PaintSizeMode = PAINTSIZEMODE_ZOOM
        Lead1.PaintZoomFactor = Lead1.PaintZoomFactor + 10
    End If
    m_Busy = False
End Sub

Public Sub cmdZoomMenos_Click()
    If m_Busy Then
        Exit Sub
    End If
    If Not FrmImagem.Visible Then
        Exit Sub
    End If
    m_Busy = True
    If Lead1.PaintZoomFactor >= 20 Then
        Lead1.PaintSizeMode = PAINTSIZEMODE_ZOOM
        Lead1.PaintZoomFactor = Lead1.PaintZoomFactor - 10
    End If
    m_Busy = False
End Sub

Private Sub Form_Activate()

   'Inclusão de chamada a rotina AtualizaAtividade
    Call AtualizaAtividade(10)
    
    If m_FirstActivate Then
        LimparValores
        LimparHeader
        LimparListas

        chkNaoVinculados.Value = 1
        chkOcorrencia.Value = 0
        tmrAtualiza.Enabled = True
        sTempo = 0
        m_IdCapa = 0

        If Not ObtemCapas Then
            MsgBox "Não existem Envelopes/Malotes com pendência de Prova Zero.", vbExclamation + vbOKOnly, App.Title
            m_IdCapa = 0
            LimparImagem
            HabilitaTimerPesquisa
            Exit Sub
        End If
        Preenche_lstCapa
        m_FirstActivate = False
        lstCapa.ListIndex = 0
    End If
    DoEvents
    If lstDocto.Enabled Then
      lstDocto.SetFocus
    End If
    
    'Centraliza o form
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
End Sub
Private Sub ChamaTelaComplementacao(ByRef sForm As Form)

  On Error GoTo ERRO_CHAMATELA

  Dim i As Integer

  i = Indice(Val(Right(lstDocto.List(lstDocto.ListIndex), 10)))

  'Preencher os campos do type Geral.Documento
  Geral.Documento.IdDocto = aDoc(i).IdDocto
  Geral.Documento.Leitura = aDoc(i).Leitura
  Geral.Documento.ValorTotal = aDoc(i).Valor * 100
  Geral.Documento.TipoDocto = aDoc(i).TipoDocto
  Geral.Documento.Status = aDoc(i).Status
  Geral.Documento.Agencia = aCapa(lstCapa.ListIndex + 1).AgOrig
  Geral.Capa.AgOrig = aCapa(lstCapa.ListIndex + 1).AgOrig
  Geral.Capa.IdEnv_Mal = aCapa(lstCapa.ListIndex + 1).IdEnv_Mal
  Geral.Capa.Num_Malote = aCapa(lstCapa.ListIndex + 1).NumMalote

  Load sForm

  sForm.SetParent Me
  sForm.SetPosition (Me.Left + (Me.Width - sForm.Width) / 2), Me.Top

  sForm.AlteraValor = True

  sForm.Show vbModal, Me

  sForm.AlteraValor = False

  If sForm.Alterou = True Then

    'Gravar Log
    Call GravaLog(aCapa(lstCapa.ListIndex + 1).IdCapa, aDoc(i).IdDocto, 61)

    'Atualizar a Lista de Documentos
    txtValor.Text = Trim(RetiraPonto(FormataValor(Geral.Documento.ValorTotal, 22)))
    AlterarValor (False)
  End If

  lstDocto.SetFocus

  Unload sForm

  Exit Sub

ERRO_CHAMATELA:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao executar tela para Complementação de Documentos.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Sub
Public Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim Ret As Long

    hCtl = ProvaZero.Lead1.hwnd

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

      Case vbKeyDown
          Ret = SendMessage(hCtl, WM_VSCROLL, SB_LINEDOWN, 0)
      Case vbKeyUp
          Ret = SendMessage(hCtl, WM_VSCROLL, SB_LINEUP, 0)
      Case vbKeyLeft
          Ret = SendMessage(hCtl, WM_HSCROLL, SB_LINEUP, 0)
      Case vbKeyRight
          Ret = SendMessage(hCtl, WM_HSCROLL, SB_LINEDOWN, 0)
      Case vbKeyPageUp
          Ret = SendMessage(hCtl, WM_VSCROLL, SB_PAGEUP, 0)
      Case vbKeyPageDown
          Ret = SendMessage(hCtl, WM_VSCROLL, SB_PAGEDOWN, 0)
      Case vbKeyHome
          Ret = SendMessage(hCtl, WM_VSCROLL, SB_TOP, 0)
          Ret = SendMessage(hCtl, WM_HSCROLL, SB_TOP, 0)
      Case vbKeyEnd
          Ret = SendMessage(hCtl, WM_VSCROLL, SB_BOTTOM, 0)
          Ret = SendMessage(hCtl, WM_HSCROLL, SB_BOTTOM, 0)
    End Select

End Sub

Private Sub Form_Load()
    
    Dim Control As Control
    
    For Each Control In Principal.Controls
        If TypeName(Control) = "Menu" Then
            'Verifica se menu de Troca de ordem está habilitado
            If (Control.Index) = 20 Then
                If Not Control.Enabled Then
                    CmdTrocaOrderm.Enabled = False
                End If
                Exit For
            End If
        End If
    Next
    
    With Lead1
        .UnlockSupport L_SUPPORT_EXPRESS, "YXPQ3XPPVT"
        .UnlockSupport L_SUPPORT_GIFLZW, "0K3RV9UY3EY"
        .UnlockSupport L_SUPPORT_TIFLZW, "9LE75L0FDXHK"
    End With

    Set qryGetDocumentoProvaZero = Geral.Banco.CreateQuery("", "{Call GetDocumentoProvaZero (?,?)}")
    Set qryAtualizaStatusCapa = Geral.Banco.CreateQuery("", "{? = Call AtualizaStatusCapa (?,?,?)}")
    Set qryAtualizaValorDocumento = Geral.Banco.CreateQuery("", "{? = Call AtualizaValorDocumento (?,?,?,?)}")
    Set qryGetocorrencia = Geral.Banco.CreateQuery("", "{ Call GetOcorrencia (?)}")
    Set qryObtemStatusCapa = Geral.Banco.CreateQuery("", "{ Call GetCapa (?,?)}")
    Set qryVerificaCapaDisponivel = Geral.Banco.CreateQuery("", "{ ? = Call VerificaCapaDisponivel (?,?,?,?,?)}")
    
    m_FirstActivate = True
    
    
    '''''''''''''''''''''''''''
    'Loga a acao Entrar Modulo'
    '''''''''''''''''''''''''''
    Call GravaLog(0, 0, 164)
    
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If m_Busy Then
        Cancel = 1
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    tmrAtualiza.Enabled = False
    tmrPesquisa.Enabled = False
    
    If m_IdCapa > 0 Then
        If m_Alteracao Then
            AtualizaStatusCapa m_IdCapa, "8"
        Else
            AtualizaStatusCapa m_IdCapa, "4"
        End If
    End If
    qryGetCapaProvaZero.Close
    qryGetDocumentoProvaZero.Close
    qryAtualizaStatusCapa.Close
    qryAtualizaValorDocumento.Close
    qryGetocorrencia.Close
    qryVerificaCapaDisponivel.Close
    
    '''''''''''''''''''''''''
    'Loga a acao Sair Modulo'
    '''''''''''''''''''''''''
    Call GravaLog(0, 0, 165)

End Sub

Private Sub Lead1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Lead1.AutoRubberBand = True
        Lead1.MousePointer = 2
    Else
        MostraImagem
    End If
End Sub

Private Sub Lead1_RubberBand()
    Dim zoomleft As Integer
    Dim zoomtop As Integer
    Dim zoomwidth As Integer
    Dim zoomheight As Integer
    
    On Error GoTo ERRO_RUBBERBAND
    
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
    Dim Count As Integer
    Dim AindaExiste As Boolean
    
    If m_Busy Then
        Exit Sub
    End If
    m_Busy = True

    If lstCapa.ListIndex = -1 Then
        m_Busy = False
        Exit Sub
    End If


    If aCapa(lstCapa.ListIndex + 1).IdEnv_Mal = "E" Then
        lblCapa.Caption = "Envelope"
        picNumMalote.Visible = False
        lblNumMalote.Visible = False
    Else
        lblCapa.Caption = "Malote"
        picNumMalote.Visible = True
        lblNumMalote.Visible = True
        lblNumMalote.Caption = aCapa(lstCapa.ListIndex + 1).NumMalote
    End If
    
    If m_IdCapa > 0 Then
        If m_Alteracao Then
            If Not AtualizaStatusCapa(m_IdCapa, "8") Then
                m_Busy = False
                m_IdCapa = 0
                Exit Sub
            End If
        Else
            If Not AtualizaStatusCapa(m_IdCapa, "4") Then
                m_Busy = False
                m_IdCapa = 0
                Exit Sub
            End If
        End If
    End If

    If m_CountCapa > 0 Then
        m_IdCapa = aCapa(lstCapa.ListIndex + 1).IdCapa
    End If
    
    If Not VerificaCapaDisponivel(m_IdCapa) Then
        m_IdCapa = 0
        m_Busy = False
        m_CountDocto = 0
        Preenche_lstDocto (0)
        LimparValores
        Exit Sub
    End If
    
    'Verificar se existem documentos transmitidos/expedidos ou com NSU
    If VerificaDocumentosTransmitidos Then
        m_IdCapa = 0
        m_Busy = False
        CmdAtualizar_Click
        Exit Sub
    End If
    
    If Not AtualizaStatusCapa(m_IdCapa, "G") Then
        m_IdCapa = 0
        m_Busy = False
        Exit Sub
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''
    'Grava log "193 - Prova Zero - Selecionar Capa'
    '''''''''''''''''''''''''''''''''''''''''''''''
    GravaLog m_IdCapa, 0, 193

    m_Alteracao = False
    lblLote.Caption = Format(aCapa(lstCapa.ListIndex + 1).IdLote, "0000-00000")
    ObtemDocumentos m_IdCapa
    If CalculaValores Then
        m_Busy = False
        Exit Sub
    End If
    sTempo = 0
    Preenche_lstDocto (0)
    lstDocto.SetFocus

    m_Busy = False
End Sub
Private Sub lstCapa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys ("{TAB}")
    End If
End Sub

Private Sub lstDocto_Click()
    Dim i As Integer

    If lstCapa.ListIndex <> -1 And lstDocto.ListIndex <> -1 Then
        ObtemOcorrencia
        MostraImagem
        MostraValor
    End If
End Sub

Private Sub lstDocto_DblClick()

  Dim i As Integer
  
  If lstDocto.ListIndex = -1 Then Exit Sub

  i = Indice(Val(Right(lstDocto.List(lstDocto.ListIndex), 10)))

  Select Case aDoc(i).TipoDocto
    Case 13
      'Cobranca Registrada
      Call ChamaTelaComplementacao(CobrancaRegistrada)

    Case 14
      'Cobranca Especial
      Call ChamaTelaComplementacao(CobrancaEspecial)

    Case 16
      'DARF Preto
      Call ChamaTelaComplementacao(DARFPreto)

    Case 17
      'DARF Simples
       Call ChamaTelaComplementacao(DARFSimples)

    Case 18
      'GARE
      Call ChamaTelaComplementacao(GareICMS)

    Case 20, 21, 22, 23
      MsgBox "Não é permitido alterar valores de concessionárias.", vbInformation + vbOKOnly, App.Title

    Case 28, 29, 30, 31
      'Ficha de Compensacao
      Call ChamaTelaComplementacao(FichaCompensacao)

    Case 32, 33, 34, 38, 42, 43
      'Ajustes
      MsgBox "Não é permitido alterar valores de Ajustes.", vbInformation + vbOKOnly, App.Title

    Case 41
      'Lancamento Interno
      MsgBox "Não é permitido alterar valor de Lancamento Interno.", vbInformation + vbOKOnly, App.Title

    Case 35
      'GPS
      Call ChamaTelaComplementacao(GPS)

    Case 36
      'CARTAO CREDITO AVULSO
      Call ChamaTelaComplementacao(CartaoAvulso)

    Case 40
      'FGTS
      Call ChamaTelaComplementacao(frmFGTS)

    Case Else
      SendKeys ("{TAB}")

  End Select
End Sub
Private Sub lstDocto_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    Call lstDocto_DblClick
  End If
End Sub

Private Sub tmrAtualiza_Timer()
    tmrAtualiza.Enabled = False
    If m_IdCapa > 0 Then
        sTempo = sTempo + Int(tmrAtualiza.Interval / 1000)
        If sTempo + Int(tmrAtualiza.Interval / 1000) >= Geral.Intervalo Then
            AtualizaStatusCapa m_IdCapa, "G"
            sTempo = 0
            '''''''''''''''''''''''''''''''''''''''
            'Grava Log MDI - Fim Aguarda documento'
            '''''''''''''''''''''''''''''''''''''''
            Call GravaLog(0, 0, 255)
        End If

    End If
    tmrAtualiza.Enabled = True
End Sub

Private Sub tmrPesquisa_Timer()
  tmrPesquisa.Enabled = False

  sTempo = sTempo + Int(tmrPesquisa.Interval / 1000)

  If sTempo + Int(tmrPesquisa.Interval / 1000) >= Geral.Atualizacao Then
    sTempo = 0
    If ObtemCapas Then
        FrmPesquisa.Visible = False
        '''''''''''''''''''''''''''''''''''''''
        'Grava Log MDI - Fim Aguarda documento'
        '''''''''''''''''''''''''''''''''''''''
        Call GravaLog(0, 0, 255)
        
        Preenche_lstCapa
        lstCapa.ListIndex = 0
        tmrPesquisa.Enabled = False
        Exit Sub
    End If

    ''''''''''''''''''''''''''''''''''''''''''
    'Grava Log MDI - Inicio Aguarda documento'
    ''''''''''''''''''''''''''''''''''''''''''
    Call GravaLog(0, 0, 254)
    
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

Private Sub txtNumEnvMal_GotFocus()
    SelecionarTexto txtNumEnvMal
End Sub

Private Sub TxtNumEnvMal_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdProcurar_Click
    Else
        SoNumero KeyAscii
    End If
End Sub

Private Sub txtValor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdAlterar_Click
    End If
End Sub
Function VerificaDocumentosTransmitidos() As Boolean

   On Error GoTo VerificaDocumentosTransmitidos_Err

   Dim RsDoctosTrans As rdoResultset

   VerificaDocumentosTransmitidos = False

   Set qryGetDocumentosParaVerificacao = Geral.Banco.CreateQuery("", "{ ? = Call GetDocumentosParaVerificacao (?,?)}")
   

   With qryGetDocumentosParaVerificacao
      .rdoParameters(1) = Geral.DataProcessamento
      .rdoParameters(2) = m_IdCapa
   End With

   Set RsDoctosTrans = qryGetDocumentosParaVerificacao.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
   If Not RsDoctosTrans.EOF Then
      If RsDoctosTrans!Qtde > 0 Then
         VerificaDocumentosTransmitidos = True
         'Atualizar o Status da Capa para 'V' - Em Analise
         Call AtualizaStatusCapa(m_IdCapa, "V")

         'Gravar Log
         Call GravaLog(m_IdCapa, 0, 68)
         lstDocto.Clear
         MsgBox "Este Envelope/Malote não está mais disponível, capa enviada para análise.", vbInformation + vbOKOnly, App.Title
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

