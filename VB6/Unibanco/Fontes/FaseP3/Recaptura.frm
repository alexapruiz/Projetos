VERSION 5.00
Object = "{00080003-B1BA-11CE-ABC6-F5B2E79D9E3F}#8.0#0"; "LTOCX80N.OCX"
Object = "{A1FF59A3-0995-11CF-B3E8-00608C82AA8D}#1.0#0"; "PIXEZOCX.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Recaptura 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recaptura"
   ClientHeight    =   8592
   ClientLeft      =   384
   ClientTop       =   660
   ClientWidth     =   11544
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8592
   ScaleWidth      =   11544
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox frmLocalizar 
      Height          =   1272
      Left            =   4224
      ScaleHeight     =   1224
      ScaleWidth      =   2604
      TabIndex        =   38
      Top             =   1944
      Visible         =   0   'False
      Width           =   2652
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
         TabIndex        =   41
         Top             =   384
         Width           =   2304
      End
      Begin VB.CommandButton cmdProcurar 
         Caption         =   "&Localizar"
         Height          =   300
         Left            =   144
         TabIndex        =   40
         Top             =   816
         Width           =   972
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   300
         Left            =   1464
         TabIndex        =   39
         Top             =   816
         Width           =   972
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
         TabIndex        =   42
         Top             =   96
         Width           =   2232
      End
   End
   Begin VB.PictureBox FrmPesquisa 
      Height          =   1932
      Left            =   2856
      ScaleHeight     =   1884
      ScaleWidth      =   5724
      TabIndex        =   21
      Top             =   2004
      Visible         =   0   'False
      Width           =   5772
      Begin VB.CommandButton CmdFecharPesquisa 
         Caption         =   "&Fechar"
         Height          =   312
         Left            =   2184
         TabIndex        =   28
         Top             =   1464
         Width           =   1068
      End
      Begin ComctlLib.ProgressBar Progress 
         Height          =   300
         Left            =   336
         TabIndex        =   29
         Top             =   912
         Width           =   4932
         _ExtentX        =   8700
         _ExtentY        =   529
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Pesquisando por Documentos para Recaptura. Aguarde ..."
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
         TabIndex        =   30
         Top             =   576
         Width           =   4968
      End
   End
   Begin VB.Timer TmrAtualiza 
      Enabled         =   0   'False
      Interval        =   50000
      Left            =   9420
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      Height          =   264
      Left            =   120
      ScaleHeight     =   216
      ScaleWidth      =   1752
      TabIndex        =   19
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
         TabIndex        =   22
         Top             =   0
         Width           =   1296
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   252
      Left            =   1992
      ScaleHeight     =   204
      ScaleWidth      =   7776
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   384
      Width           =   7824
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Recaptura"
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
         Left            =   1104
         TabIndex        =   34
         Top             =   0
         Width           =   876
      End
      Begin VB.Label Label9 
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
         Left            =   108
         TabIndex        =   18
         Top             =   0
         Width           =   408
      End
      Begin VB.Label Label5 
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
         Height          =   228
         Left            =   2772
         TabIndex        =   17
         Top             =   0
         Width           =   1104
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
         Left            =   6312
         TabIndex        =   16
         Top             =   0
         Width           =   540
      End
   End
   Begin VB.PictureBox PctMalote 
      Height          =   264
      Left            =   3924
      ScaleHeight     =   216
      ScaleWidth      =   1176
      TabIndex        =   13
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
         TabIndex        =   14
         Top             =   0
         Width           =   1092
      End
   End
   Begin VB.PictureBox Picture4 
      Height          =   264
      Left            =   1980
      ScaleHeight     =   216
      ScaleWidth      =   528
      TabIndex        =   11
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
         TabIndex        =   12
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.Frame FrmCmdImagem 
      Height          =   5004
      Left            =   9840
      TabIndex        =   9
      Top             =   3528
      Width           =   1620
      Begin VB.CommandButton cmdZoomMais 
         Caption         =   "Zoom +"
         Height          =   696
         Left            =   384
         Picture         =   "Recaptura.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   612
         Width           =   888
      End
      Begin VB.CommandButton cmdZoomMenos 
         Caption         =   "Zoom -"
         Height          =   696
         Left            =   384
         Picture         =   "Recaptura.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1428
         Width           =   888
      End
      Begin VB.CommandButton cmdRotacao 
         Caption         =   "Rotação"
         Height          =   696
         Left            =   384
         Picture         =   "Recaptura.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2244
         Width           =   888
      End
      Begin VB.CommandButton cmdInverteCor 
         Caption         =   "Inverte cor"
         Height          =   696
         Left            =   384
         Picture         =   "Recaptura.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3060
         Width           =   888
      End
      Begin VB.CommandButton cmdFrenteVerso 
         Caption         =   "Fren/Ver"
         Height          =   696
         Left            =   384
         Picture         =   "Recaptura.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3876
         Width           =   888
      End
   End
   Begin VB.ListBox LstDocto 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2496
      ItemData        =   "Recaptura.frx":0F32
      Left            =   1980
      List            =   "Recaptura.frx":0F34
      MultiSelect     =   2  'Extended
      TabIndex        =   20
      Top             =   672
      Width           =   7836
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
      ItemData        =   "Recaptura.frx":0F36
      Left            =   120
      List            =   "Recaptura.frx":0F38
      TabIndex        =   23
      Top             =   456
      Width           =   1800
   End
   Begin VB.Frame FrmImagem 
      Caption         =   "Imagem"
      Height          =   5004
      Left            =   144
      TabIndex        =   8
      Top             =   3528
      Width           =   9696
      Begin LeadLib.Lead Lead1 
         Height          =   4644
         Left            =   96
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   240
         Width           =   9516
         _Version        =   524288
         _ExtentX        =   16785
         _ExtentY        =   8191
         _StockProps     =   229
         BackColor       =   16777215
         BorderStyle     =   1
         ScaleHeight     =   385
         ScaleWidth      =   791
         DataField       =   ""
         BitmapDataPath  =   ""
         AnnDataPath     =   ""
      End
   End
   Begin VB.Timer TmrPesquisa 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9024
      Top             =   0
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      Height          =   2484
      Left            =   1968
      ScaleHeight     =   2436
      ScaleWidth      =   7788
      TabIndex        =   31
      Top             =   672
      Width           =   7836
   End
   Begin PixezocxLib.PixEzImage EzCanon 
      Height          =   24
      Left            =   11400
      TabIndex        =   36
      Top             =   0
      Visible         =   0   'False
      Width           =   24
      _Version        =   65536
      _ExtentX        =   -42
      _ExtentY        =   42
      _StockProps     =   96
      TAG_OPEN_DIR    =   ""
      TAG_OPEN_SCHEMA =   ""
      TAG_OPEN_EXT    =   ""
      TAG_OPEN_ROOT   =   ""
      TAG_OPEN_DETECTSCHEMA=   1
      TAG_OPEN_FILENAMES=   ""
      TAG_WINDOW_CURPAGE=   0
      PIXEZ_SELECT    =   ""
      TAG_BORDER_COLOR_BG=   8421504
      TAG_BORDER_COLOR_ENVFOCUS=   254
      TAG_BORDER_COLOR_WINFOCUS=   16711422
      TAG_BRIGHTNESS  =   128
      TAG_CONTRAST    =   128
      TAG_BLUEBRIGHTNESS=   128
      TAG_BLUECONTRAST=   128
      TAG_GREENBRIGHTNESS=   128
      TAG_GREENCONTRAST=   128
      TAG_REDBRIGHTNESS=   128
      TAG_REDCONTRAST =   128
      TAG_DOC_OPENATTRIBUTE=   66
      TAG_FILLORDER   =   1
      TAG_HFLIP       =   0
      TAG_PAN_HEIGHT  =   0
      TAG_PAN_WIDTH   =   0
      TAG_PAN_XPOS    =   0
      TAG_PAN_YPOS    =   0
      TAG_PAN_SCALING =   4
      TAG_PAN_TITLE   =   "Pan Window"
      TAG_PAN_SHOW    =   0
      TAG_ONE_ACCELMODE=   0
      TAG_ONE_ACTION_CLOCKWISE=   35
      TAG_ONE_ACTION_CTRCLOCKWISE=   3
      TAG_ONE_ACTION_DEFINEREG=   0
      TAG_ONE_ACTION_DEFINEREGASPECT=   64
      TAG_ONE_ACTION_PAN=   1
      TAG_ONE_ACTION_SWITCHTOTREE=   64
      TAG_ONE_ACTION_ZOOMINREG=   64
      TAG_ONE_ACTION_ZOOMINREGASPECT=   32
      TAG_ONE_ACTION_ZOOMOUTCORNER=   44
      TAG_ONE_SCROLLBARS=   2
      TAG_ONE_SETTINGS_RANGE=   1
      TAG_ORIENTATION =   1
      TAG_OVERSCAN    =   0
      TAG_PHOTOMETRICINTERPRETATION=   0
      TAG_PRINT_COLLATE=   1
      TAG_PRINT_COPIES=   1
      TAG_PRINT_DEVICENO=   0
      TAG_PRINT_DEVNAME1=   ""
      TAG_PRINT_DEVNAME2=   ""
      TAG_PRINT_RANGEMODE=   0
      TAG_PRINT_REGION=   0
      TAG_PRINT_SCALE =   0
      TAG_PRINT_SHOWDLG=   0
      TAG_REGION_COUNT=   0
      TAG_REGION_MODE =   0
      TAG_ROTATION    =   1
      TAG_SCALING     =   1
      TAG_SCALE_X     =   1
      TAG_SCALE_Y     =   1
      TAG_SCAN_ALLOW_TURNOVER=   0
      TAG_SCAN_COLORFORMAT=   0
      TAG_SCAN_COMPRESSION=   4
      TAG_SCAN_CURPAGE=   0
      TAG_SCAN_DISPLAYPAGE=   0
      TAG_SCAN_DIR    =   ""
      TAG_SCAN_DUPLEX =   0
      TAG_SCAN_EXT    =   "."
      TAG_SCAN_FILENAME=   "\SCAN."
      TAG_SCAN_INSERTMODE=   1
      TAG_SCAN_SCHEMA =   ""
      TAG_SCAN_WARNOVERWRITE=   0
      TAG_SCAN_MULTIPAGE=   1
      TAG_SCAN_USESCHEMA=   0
      TAG_SCAN_MAXPAGES=   -1
      TAG_SCAN_ORIENTATION=   1
      TAG_SCAN_PACK   =   196608
      TAG_SCAN_PRECEDENCE=   1
      TAG_SCAN_SAVEFLAG=   0
      TAG_SCAN_ROOT   =   "SCAN"
      TAG_SCAN_USELONGNAMES=   0
      TAG_SAVE_COLORFORMAT=   0
      TAG_SAVE_COMPRESSION=   4
      TAG_SAVE_DIR    =   ""
      TAG_SAVE_EXT    =   "."
      TAG_SAVE_FILENAME=   "\SAVE."
      TAG_SAVE_ORIENTATION=   1
      TAG_SAVE_PACK   =   196608
      TAG_SAVE_PRECEDENCE=   1
      TAG_SAVE_RANGESTR=   ""
      TAG_SAVE_ROOT   =   "SAVE"
      TAG_SAVE_SAVEFLAG=   0
      TAG_SAVE_WARNOVERWRITE=   0
      TAG_SAVE_MULTIPAGE=   1
      TAG_SAVE_USESCHEMA=   0
      TAG_SAVE_USELONGNAMES=   0
      TAG_THRESH_X    =   0
      TAG_THRESH_Y    =   0
      TAG_TREE_COLOR_BG=   8421504
      TAG_TREE_COLOR_NODETEXT=   0
      TAG_TREE_COLOR_NODESELTEXT=   16777215
      TAG_TREE_COLOR_THUMBTEXT=   0
      TAG_TREE_COLOR_THUMBSELTEXT=   16777215
      TAG_TREE_COLOR_LINE=   0
      TAG_TREE_THUMBSTYLE=   528
      TAG_WINDOW_STYLE=   0
      TAG_XPOSITION   =   0
      TAG_YPOSITION   =   0
      TAG_INVERT      =   0
   End
   Begin VB.Frame FrmCmd 
      Height          =   3612
      Left            =   9840
      TabIndex        =   10
      Top             =   -72
      Width           =   1632
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
         Height          =   312
         Left            =   120
         MaskColor       =   &H00404080&
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   2568
         Width           =   1392
      End
      Begin VB.CommandButton CmdDesmarcar 
         Caption         =   "Desmarcar Docto"
         Height          =   312
         Left            =   120
         TabIndex        =   35
         Top             =   1704
         Width           =   1392
      End
      Begin VB.CommandButton CmdCapturar 
         Caption         =   "&Recapturar"
         Height          =   312
         Left            =   120
         TabIndex        =   33
         Top             =   1272
         Width           =   1392
      End
      Begin VB.CommandButton CmdIlegiveis 
         Caption         =   "Enviar &Ilegíveis"
         Height          =   312
         Left            =   120
         TabIndex        =   32
         Top             =   840
         Width           =   1392
      End
      Begin VB.CommandButton CmdAtualizar 
         Caption         =   "&Atualizar"
         Height          =   312
         Left            =   120
         TabIndex        =   0
         Top             =   408
         Width           =   1392
      End
      Begin VB.CommandButton cmdEncerrar 
         Caption         =   "&Encerrar Capa"
         Height          =   312
         Left            =   120
         TabIndex        =   1
         Top             =   2136
         Width           =   1392
      End
      Begin VB.CommandButton cmdFechar 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   312
         Left            =   120
         TabIndex        =   2
         Top             =   3000
         Width           =   1392
      End
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
      TabIndex        =   26
      Top             =   3240
      Width           =   9696
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
      TabIndex        =   25
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
      TabIndex        =   24
      Top             =   48
      Width           =   1224
   End
End
Attribute VB_Name = "Recaptura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Delaração dos Objetos RDO
Private qryGetCapa                          As rdoQuery
Private qryGetCapaDuplicada                 As rdoQuery
Private qryGetDocumentos                    As rdoQuery
Private qryAtualizaStatusCapa               As rdoQuery
'Private qryAtualizaStatusDocumento          As rdoQuery
'Private qryAtualizaOcorrenciaDocumento      As rdoQuery
'Private qryAtualizaDuplicidadeDocumento     As rdoQuery
Private qryGetOcorr                         As rdoQuery
Private qryRemoveAjusteCapa                 As rdoQuery
Private qryAtualizaCapa                     As rdoQuery
Private qryGetDocumentosCapa                As rdoQuery
Private qryRemoveDocumento                  As rdoQuery
Private qryGetCapaLoteEmRecaptura           As rdoQuery
Private qryGetDocumentosParaVerificacao     As rdoQuery

'Declaração dos objetos RDO da Captura
Private qryGetImagem                        As rdoQuery
Private qryProducaoScanner                  As rdoQuery
Private qryInsereCapa                       As rdoQuery
Private qryInsereDocto                      As rdoQuery
Private qryGetUltimaImagemLote              As rdoQuery
Private qryAtualizaOrdemCapturaDocumento    As rdoQuery
Private qryAtualizaTipoDoctoCapa            As rdoQuery

'Declaração de Variáveis
Private AlterouDocto                        As Boolean
Private PrimeiraVez                         As Boolean
Private teclou                              As Boolean
Private IdSelecionado                       As Long
Private sTempo                              As Integer
Private FileLog                             As Integer
Private OrdemCapturaInicial                 As Integer
Private PathIni                             As String

'Declaração dos Arrays
Private aDoc()  As TDoc
Private aCapa() As TCapa

'Type de Capas
Private Type TCapa
  IdCapa        As Long
  IdLote        As Long
  IdEnv_Mal     As String * 1
  Capa          As String * 18
  NumMalote     As String * 11
  AgOrig        As Integer
  Status        As String * 1
  AlterouDocto  As Boolean
  Duplicidade   As Integer
End Type

'Type para Documentos
Private Type TDoc
  NrSeq             As Integer
  IdDocto           As Long
  IdCapa            As Long
  TipoDocto         As Integer
  DscTipoDocto      As String * 18
  Duplicidade       As Boolean
  Ocorrencia        As String * 5
  RetornoTransacao  As Long
  OrdemCaptura      As Integer
  Leitura           As String * 48
  Frente            As String * 20
  Verso             As String * 20
  Status            As String * 1
  Ordem             As Integer
  Vinculo           As Long
  Valor             As Double
End Type
Private Sub HabilitaRecaptura()

    Dim iRet As Long
    Dim ScannerOk As Boolean
    Dim NumBoxes As Long
    Dim MaxDocBox As Long
    Dim BoxDefault As Long
    Dim Threshold As Long
    Dim Compress As Long
    Dim Resolution As Long

    '''''''''''''''''''''''
    ' Inicializar Scanner '
    '''''''''''''''''''''''
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
            
        ElseIf Geral.VIPSDLL = eDllNovaUBB Then 'VipsDll (Nova Versão)
            tSC_ParamDLL.BoxDefault = BoxDefault
            tSC_ParamDLL.MaxDocBox = MaxDocBox
            tSC_ParamDLL.NumBoxes = NumBoxes
            
            If InicializarVips Then
                ScannerOk = True
                bInicializou = True
            End If
            
        Else ' VipsDll do Unibanco
            VIPS_SetBoxes (NumBoxes)
            VIPS_SetMaxDocBox (MaxDocBox)
            VIPS_SetBoxDefault (BoxDefault)
            VIPS_SetCompress (Compress)
            VIPS_SetCutBords (Threshold)
            VIPS_SetCameraFile ("Doc100.cpf")
            VIPS_SetImageDirectory (Geral.DiretorioImagens)
            VIPS_SetResolution (Resolution)
            iRet = VIPS_Init()
            If iRet = 0 Then
                ScannerOk = True
            End If
        End If

    ElseIf Geral.Scanner = escnCanonLS500 Then
      ' Inicializção da LS500 e Canon
      iRet = LS_ProcuraLS500(string1, string2, string3)
      If iRet <> 0 Then
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
    End If
    
    If ScannerOk Then
      ' habilitar botão de recaptura
      Recaptura.CmdCapturar.Enabled = True
    Else
      ' desabilitar botão de recaptura
      Recaptura.CmdCapturar.Enabled = False
    End If
    
End Sub
Private Sub AtualizaOrdemCaptura()

    Dim DoctoSelec  As Boolean
    Dim X           As Integer

    For X = 0 To lstDocto.ListCount - 1
        If lstDocto.Selected(X) = True Then
            DoctoSelec = True
        ElseIf DoctoSelec = True Then
            'Atualizar o campo OrdemCaptura
            Call AtualizaOrdemCapturaDocumento(aDoc(X + 1).IdDocto, OrdemCapturaInicial)
            OrdemCapturaInicial = OrdemCapturaInicial + 1
        End If
        DoEvents
    Next X
End Sub
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
    If sStatus = "1" Then
        'Envia para Complementação
        Call GravaLog(sIdCapa, 0, 202)
    ElseIf sStatus = "5" Then
        'Envia para Ilegiveis
        Call GravaLog(sIdCapa, 0, 203)
    End If

    Exit Sub

ERRO_ATUALIZASTATUS:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro ao Atualizar o Status do Documento.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
End Sub
Private Function AtualizaTipoDoctoCapa(ByVal IdCapa As Long) As Boolean

    On Error GoTo AtualizaTipoDoctoCapa_Err

    AtualizaTipoDoctoCapa = False

    Set qryAtualizaTipoDoctoCapa = Geral.Banco.CreateQuery("", "{? = call AtualizaTipoDoctoCapa (?,?)}")
    With qryAtualizaTipoDoctoCapa
        .rdoParameters(0).Direction = rdParamReturnValue     'Parametro de Retorno
        .rdoParameters(1) = Geral.DataProcessamento          'Data Processamento
        .rdoParameters(2) = IdCapa                           'IdDocto
        .Execute
    End With

    If qryAtualizaTipoDoctoCapa(0).Value <> "0" Then
        MsgBox "Ocorreu um erro ao Atualizar o tipo de documento da capa.", vbInformation + vbOKOnly, App.Title
        Exit Function
    End If

    AtualizaTipoDoctoCapa = True

    Exit Function

AtualizaTipoDoctoCapa_Err:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro ao Atualizar o tipo de documento da capa.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
End Function
Sub HabilitaTimerPesquisa()

    'Esta Função irá verificar a existência de documentos para recaptura a cada x segundos
    'de acordo com o campo PARAMETRO.TmAtualizacao
    FrmPesquisa.Visible = True
    tmrPesquisa.Enabled = True
    Progress.Value = 0
End Sub
Sub HDObjetosImagem(bValor As Boolean)

    On Error GoTo ERRO_HDOBJETOS

    cmdZoomMais.Enabled = bValor
    cmdZoomMenos.Enabled = bValor
    cmdRotacao.Enabled = bValor
    cmdInverteCor.Enabled = bValor
    cmdFrenteVerso.Enabled = bValor
    FrmImagem.Visible = bValor
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

    lstDocto.Clear
    lblOcorrencia.Caption = ""
End Sub
Sub LimpaListas()

    lstCapa.Clear
    lstDocto.Clear
    FrmImagem.Visible = False
    Erase aCapa
End Sub
Function PreencheListCapas() As Boolean

    On Error GoTo ERRO_PREENCHELISTCAPAS

    Dim rsCapa        As rdoResultset
    Dim sSql          As String
    Dim sPosicaoErro  As String
    Dim X             As Integer

    Call LimpaListas

    'Passando parâmetros para a Stored Procedure 'GetCapaRecaptura'
    sSql = Geral.DataProcessamento & " , " & Geral.Intervalo

    Set qryGetCapa = Geral.Banco.CreateQuery("", "{call GetCapaRecaptura (" & sSql & ")}")

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

            lstCapa.AddItem (rsCapa!Capa)
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

    Set qryGetCapa = Geral.Banco.CreateQuery("", "{? = call VerificaCapaDisponivel (?,?,?,?,?)}")

    With qryGetCapa
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1) = Geral.DataProcessamento             'Data de Processamento
        .rdoParameters(2) = aCapa(lstCapa.ListIndex + 1).IdCapa 'IdCapa
        .rdoParameters(3) = "A"                                 'Status 1
        .rdoParameters(4) = "B"                                 'Status 2 (Pendentes)
        .rdoParameters(5) = Geral.Intervalo                     'Intervalo de Atualização

        .Execute
    End With

    CapaSelecionadaDisponivel = qryGetCapa(0)

    If qryGetCapa(0) = 1 Then
        lstCapa.ListIndex = -1
        lstDocto.Clear
        FrmImagem.Visible = False
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

    Dim sSql          As String
    Dim sLinha        As String
    Dim rsDocumentos  As rdoResultset
    Dim X             As Integer

    lstDocto.Visible = False

    'Selecionar todos os documentos pertencentes à capa selecionada
    sSql = Geral.DataProcessamento & " , " & Val(lstCapa.ItemData(lstCapa.ListIndex))

    Set qryGetDocumentos = Geral.Banco.CreateQuery("", "{call GetDocumentoRecaptura (" & sSql & ")}")

    Set rsDocumentos = qryGetDocumentos.OpenResultset(rdOpenStatic, rdConcurReadOnly)

    X = 1
    Call LimpaListaDocto
    ReDim aDoc(rsDocumentos.RowCount)

    If Not rsDocumentos.EOF Then
        While Not rsDocumentos.EOF

            'Numero Sequencial
            aDoc(X).NrSeq = X
            sLinha = Format(aDoc(X).NrSeq, "0000") & Space(9)

            'Status do Documento
            aDoc(X).Status = rsDocumentos!Status & ""
            
            'Ocorrencia do documento
            aDoc(X).Ocorrencia = rsDocumentos!Ocorrencia
            
            'Retorno de Transação
            aDoc(X).RetornoTransacao = rsDocumentos!RetornoTransacao

            'Tipo de Documento
            aDoc(X).TipoDocto = rsDocumentos!TipoDocto & ""

            'Ordem de Captura
            aDoc(X).OrdemCaptura = rsDocumentos!OrdemCaptura & ""

            'Preenche o campo Ordem (0 - VIPS , 1 - Canon , 2 - LS 500)
            aDoc(X).Ordem = rsDocumentos!Ordem & ""

            'Flag que indica se o documento está marcado para recaptura
            If aDoc(X).Status = "A" Then
                sLinha = sLinha & "S" & Space(12)
            Else
                sLinha = sLinha & Space(13)
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

            sLinha = sLinha & aDoc(X).DscTipoDocto & Space(3)

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

            lstDocto.AddItem sLinha
            lstDocto.ItemData(lstDocto.NewIndex) = rsDocumentos!IdDocto
            rsDocumentos.MoveNext
            X = X + 1
            DoEvents
        Wend
    Else
        Call HDObjetosImagem(False)
    End If

    If lstCapa.ListCount > 0 And lstDocto.ListCount > 0 Then
        lstDocto.Selected(Val(Indice)) = True
        IdSelecionado = aCapa(lstCapa.ListIndex + 1).IdCapa
    End If

    lstDocto.Visible = True

    If lstDocto.Visible = True Then
        lstDocto.SetFocus
    End If
    DoEvents

    Exit Sub

ERRO_PREENCHELISTDOCTO:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro ao Preencher Lista de Documentos.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
End Sub
Private Function RemoveDoctosSelecionados(ByVal Desmarca As Boolean, Optional ByVal pRecaptura As Boolean = False) As Boolean

    On Error GoTo REMOVEDOCTOSSELECIONADOS_ERR

    Dim X As Integer

    RemoveDoctosSelecionados = False

    For X = 0 To lstDocto.ListCount - 1
        If lstDocto.Selected(X) = True Then
            'Excluir cada documento da capa
            If (Desmarca = False) Or _
               (Desmarca And aDoc(X + 1).TipoDocto > 1) Then
                With qryRemoveDocumento
                    .rdoParameters(0).Direction = rdParamReturnValue
                    .rdoParameters(1) = Geral.DataProcessamento
                    .rdoParameters(2) = aDoc(X + 1).IdDocto
                    .rdoParameters(3) = IIf(aDoc(X + 1).TipoDocto = 1, 0, Abs(pRecaptura))
                    .Execute
                End With
    
                If qryRemoveDocumento(0).Value = 1 Then
                    MsgBox "Ocorreu um erro ao remover documentos recapturados.", vbInformation + vbOKOnly, App.Title
                    Exit Function
                End If

                'Gravar Log
                Call GravaLog(aCapa(lstCapa.ListIndex + 1).IdCapa, aDoc(X + 1).IdDocto, 204)
            End If
        End If
        DoEvents
    Next X

    RemoveDoctosSelecionados = True

    Exit Function

REMOVEDOCTOSSELECIONADOS_ERR:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro ao remover documentos recapturados.", Err, rdoErrors)
        Case vbCancel, vbRetry
            Unload Me
    End Select
End Function
Function PossuiDoctosRecaptura() As Boolean

    Dim X As Integer

    PossuiDoctosRecaptura = False

    'Verificar se existe algum documento indefinido
    If lstDocto.ListCount > 0 Then
        PossuiDoctosRecaptura = False
        For X = 0 To lstDocto.ListCount - 1
            If aDoc(X + 1).Status = "A" Then
                'Documento Indefinido
                PossuiDoctosRecaptura = True
                Exit Function
            End If
            DoEvents
        Next X
    End If
End Function
Private Function SelecionaPrimeiroDoctoparaRecaptura() As Integer

    Dim X As Integer

    SelecionaPrimeiroDoctoparaRecaptura = 0

    For X = 0 To lstDocto.ListCount - 1
        If lstDocto.Selected(X) = True Then
            'Primeiro Docto selecionado
            SelecionaPrimeiroDoctoparaRecaptura = aDoc(X + 1).OrdemCaptura
            Exit Function
        End If
        DoEvents
    Next X
End Function
Private Function SelectDoctoNaoSubSeq() As Boolean

    Dim X           As Integer
    Dim DoctoSelec  As Boolean
    Dim Y           As Integer

    SelectDoctoNaoSubSeq = False

    Y = 0
    For X = 0 To lstDocto.ListCount - 1
        If lstDocto.Selected(X) = True Then
            DoctoSelec = True
            Y = Y + 1
        Else
            If DoctoSelec = True Then
                'Encontrou item deselecionado apos itens selecionados
                If Y < lstDocto.SelCount Then
                    MsgBox "Não é permitido selecionar documentos não subsequentes.", vbInformation + vbOKOnly, App.Title
                    Exit Function
                End If
            End If
        End If

        DoEvents
    Next X

    SelectDoctoNaoSubSeq = True
End Function
Private Sub CmdAtualizar_Click()

    If Screen.MousePointer = vbDefault Then
        Screen.MousePointer = vbHourglass

        If IdSelecionado <> 0 Then
            Call AtualizaStatusCapa(IdSelecionado, "A")
            IdSelecionado = 0
        End If

        Screen.MousePointer = vbDefault

        If Not PreencheListCapas Then
            MsgBox "Não Existem Envelopes / Malotes Recaptura.", vbInformation, App.Title
            Call HabilitaTimerPesquisa
        End If
    Else
        Call HDMalote(False)
    End If
End Sub

Private Sub cmdCancelar_Click()
   frmLocalizar.Visible = False
   lstDocto.SetFocus
End Sub

Private Sub CmdCapturar_Click()

    Dim QtdDoctos           As Integer
    Dim PrimeiroDoctoRecap  As Integer

    'Verificar se há alguma capa selecionada
    If lstCapa.ListIndex = -1 Then
        MsgBox "Nenhum Envelope / Malote selecionado.", vbInformation, App.Title
        lstDocto.SetFocus
        Exit Sub
    End If

    'Verifica se foi determinado o scanner para captura
    If Geral.Scanner = escnSemScanner Then
        MsgBox "Nenhum scanner foi selecionado, favor verificar.", vbInformation, App.Title
        lstDocto.SetFocus
        Exit Sub
    End If
    
    'Verificar se foi selecionado mais de um docto nao-subsequente
    If Not SelectDoctoNaoSubSeq Then Exit Sub

    Screen.MousePointer = vbHourglass

    'Nome do Arquivo de Retorno
    Geral.RetornoFinal = Format(Geral.DataProcessamento, "00000000") & Format(aCapa(lstCapa.ListIndex + 1).IdLote, "000000000") & "r.txt"

    'Verificar qual a ordemCaptura do primeiro documento a ser recapturado
    OrdemCapturaInicial = SelecionaPrimeiroDoctoparaRecaptura

    'Guardar a posição do primeiro documento recapturado
    PrimeiroDoctoRecap = OrdemCapturaInicial

    If OrdemCapturaInicial <> 0 Then

        If Digitalizar(False, QtdDoctos) = True Then

            'Atualizar a OrdemCaptura dos documentos
            Call AtualizaOrdemCaptura

            'Remover os documentos selecionados
            Call RemoveDoctosSelecionados(False, False)

            'Atualizar o tipodocto da capa para '1'
            Call AtualizaTipoDoctoCapa(aCapa(lstCapa.ListIndex + 1).IdCapa)

            'Atualizar a lista de documentos
            Call PreencheListDocto(PrimeiroDoctoRecap - 1)
        End If

    Else

        MsgBox "Nenhum Documento selecionado para Recaptura.", vbInformation + vbOKOnly, App.Title

    End If

    Screen.MousePointer = vbDefault
End Sub
Private Sub CmdDesmarcar_Click()

    Dim X As Integer
    Dim EncontrouDocsRecap As Boolean

    EncontrouDocsRecap = False
    For X = 0 To lstDocto.ListCount - 1
        If lstDocto.Selected(X) = True Then
            'Verificar se o documento está marcado para recaptura
            If aDoc(X + 1).Status = "A" Then
                If aDoc(X + 1).TipoDocto <> 0 And aDoc(X + 1).Ocorrencia <> 0 Then
                    Call AtualizaStatusDocumento(aDoc(X + 1).IdDocto, "D")
                    aDoc(X + 1).Status = "D"
                ElseIf aDoc(X + 1).TipoDocto = 0 And aDoc(X + 1).Ocorrencia <> 0 Then
                    Call AtualizaStatusDocumento(aDoc(X + 1).IdDocto, "D")
                    aDoc(X + 1).Status = "D"
                ElseIf aDoc(X + 1).TipoDocto <> 0 And aDoc(X + 1).Ocorrencia = 0 Then
                    Call AtualizaStatusDocumento(aDoc(X + 1).IdDocto, "1")
                    aDoc(X + 1).Status = "1"
                Else
                    Call AtualizaStatusDocumento(aDoc(X + 1).IdDocto, "0")
                    aDoc(X + 1).Status = "0"
                End If
'                Call AtualizaStatusDocumento(aDoc(X + 1).IdDocto, "0")
'                Call AtualizaOcorrenciaDocumento(aDoc(X + 1).IdDocto)
'                Call AtualizaDuplicidadeDocumento(aDoc(X + 1).IdDocto)
'                aDoc(X + 1).Status = "0"
'                aDoc(X + 1).Ocorrencia = ""
'                aDoc(X + 1).Duplicidade = False
                EncontrouDocsRecap = True
                
                'Gravar Log
                Call GravaLog(aCapa(lstCapa.ListIndex + 1).IdCapa, aDoc(X + 1).IdDocto, 206)
                
            End If
        End If
        DoEvents
    Next X

    If EncontrouDocsRecap = True Then

'        ''''''''''''''''''''''''''''''''''''''''''''
'        'Remove os documentos que foram desmarcados'
'        ''''''''''''''''''''''''''''''''''''''''''''
'        Call RemoveDoctosSelecionados(True, True)

        'Refazer a lista de documentos
        Call PreencheListDocto(lstDocto.ListIndex)
    Else
        MsgBox "Nenhum documento marcado para recaptura foi selecionado.", vbInformation + vbOKOnly, App.Title
        lstDocto.SetFocus
        Exit Sub
    End If
End Sub
'Private Function AtualizaStatusDocumento(ByVal IdDocto As Long, ByVal Status As String) As Boolean
'
'    On Error GoTo AtualizaStatusDocumento_Err
'
'    AtualizaStatusDocumento = False
'
'    Set qryAtualizaStatusDocumento = Geral.Banco.CreateQuery("", "{? = call AtualizaStatusDocumento (?,?,?)}")
'    With qryAtualizaStatusDocumento
'       .rdoParameters(0).Direction = rdParamReturnValue     'Parametro de Retorno
'       .rdoParameters(1) = Geral.DataProcessamento          'Data Processamento
'       .rdoParameters(2) = IdDocto                          'IdDocto
'       .rdoParameters(3) = Status                           'Status do Documento
'       .Execute
'    End With
'
'    If qryAtualizaStatusDocumento(0).Value = "1" Then
'       MsgBox "Ocorreu um erro ao Atualizar o Status do Documento.", vbInformation + vbOKOnly, App.Title
'       Exit Function
'    End If
'
'    AtualizaStatusDocumento = True
'
'    Exit Function
'
'AtualizaStatusDocumento_Err:
'  Screen.MousePointer = vbDefault
'  Select Case TratamentoErro("Erro ao Atualizar o Status do documento.", Err, rdoErrors)
'    Case vbCancel
'    Case vbRetry
'  End Select
'End Function
'
'Esta rotina atualiza a ocorrencia do documento para Null
'
'Private Function AtualizaOcorrenciaDocumento(ByVal IdDocto As Long) As Boolean
'
'    On Error GoTo Erro_AtualizaOcorrenciaDocumento
'
'    AtualizaOcorrenciaDocumento = False
'
'    Set qryAtualizaOcorrenciaDocumento = Geral.Banco.CreateQuery("", "{? = call AtualizaOcorrenciaDocumento(?,?,?)}")
'    With qryAtualizaOcorrenciaDocumento
'       .rdoParameters(0).Direction = rdParamReturnValue     'Parametro de Retorno
'       .rdoParameters(1) = Geral.DataProcessamento          'Data Processamento
'       .rdoParameters(2) = IdDocto                          'IdDocto
'       .rdoParameters(3) = Null                             'Ocorrência do Documento
'       .Execute
'    End With
'
'    If qryAtualizaOcorrenciaDocumento(0).Value = "1" Then
'       MsgBox "Ocorreu um erro ao Atualizar a ocorrência do Documento.", vbInformation + vbOKOnly, App.Title
'       Exit Function
'    End If
'
'    AtualizaOcorrenciaDocumento = True
'
'    Exit Function
'
'Erro_AtualizaOcorrenciaDocumento:
'  Screen.MousePointer = vbDefault
'  Select Case TratamentoErro("Erro ao Atualizar a ocorrência do documento.", Err, rdoErrors)
'    Case vbCancel
'    Case vbRetry
'  End Select
'End Function

'
'Esta rotina atualiza a duplicidade do documento para 0
'
'Private Function AtualizaDuplicidadeDocumento(ByVal IdDocto As Long) As Boolean
'
'    On Error GoTo Erro_AtualizaDuplicidadeDocumento
'
'    AtualizaDuplicidadeDocumento = False
'
'    Set qryAtualizaDuplicidadeDocumento = Geral.Banco.CreateQuery("", "{? = call AtualizaDuplicidadeDocumento(?,?,?)}")
'    With qryAtualizaDuplicidadeDocumento
'       .rdoParameters(0).Direction = rdParamReturnValue     'Parametro de Retorno
'       .rdoParameters(1) = Geral.DataProcessamento          'Data Processamento
'       .rdoParameters(2) = IdDocto                          'IdDocto
'       .rdoParameters(3) = 0                                'Duplicidade do Documento
'       .Execute
'    End With
'
'    If qryAtualizaDuplicidadeDocumento(0).Value = "1" Then
'       MsgBox "Ocorreu um erro ao Atualizar a duplicidade do Documento.", vbInformation + vbOKOnly, App.Title
'       Exit Function
'    End If
'
'    AtualizaDuplicidadeDocumento = True
'
'    Exit Function
'
'Erro_AtualizaDuplicidadeDocumento:
'  Screen.MousePointer = vbDefault
'  Select Case TratamentoErro("Erro ao Atualizar a duplicidade do documento.", Err, rdoErrors)
'    Case vbCancel
'    Case vbRetry
'  End Select
'End Function


Private Function AtualizaOrdemCapturaDocumento(ByVal IdDocto As Long, ByVal Ordem As Integer) As Boolean

    On Error GoTo AtualizaOrdemCapturaDocumento_Err

    AtualizaOrdemCapturaDocumento = False

    Set qryAtualizaOrdemCapturaDocumento = Geral.Banco.CreateQuery("", "{? = call AtualizaOrdemCapturaDocto (?,?,?)}")
    With qryAtualizaOrdemCapturaDocumento
       .rdoParameters(0).Direction = rdParamReturnValue     'Parametro de Retorno
       .rdoParameters(1) = Geral.DataProcessamento          'Data Processamento
       .rdoParameters(2) = IdDocto                          'IdDocto
       .rdoParameters(3) = Ordem                            'Ordem de Captura do documento
       .Execute
    End With

    If qryAtualizaOrdemCapturaDocumento(0).Value = "1" Then
       MsgBox "Ocorreu um erro ao Atualizar a Ordem de Captura dos documentos.", vbInformation + vbOKOnly, App.Title
       Exit Function
    End If

    AtualizaOrdemCapturaDocumento = True

    Exit Function

AtualizaOrdemCapturaDocumento_Err:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Atualizar a Ordem de Captura dos Documentos.", Err, rdoErrors)
    Case vbCancel
    Case vbRetry
  End Select
End Function
Private Sub cmdEncerrar_Click()

    Dim X As Integer

    'Verificar se existe algum processo em andamento
    If Screen.MousePointer = vbDefault And FrmPesquisa.Visible = False Then

        'Verificar se existe alguma selecionada para Encerramento
        If lstCapa.ListIndex = -1 Then
            MsgBox "Nenhum Envelope / Malote selecionado.", vbInformation, App.Title
            Exit Sub
        End If

        'Verificar se a capa possui algum documento marcado para recaptura
        If PossuiDoctosRecaptura Then
            MsgBox "Ainda existem documentos marcados para recaptura.", vbInformation + vbOKOnly, App.Title
            Exit Sub
        End If

        If MsgBox("Confirma o Encerramento da Capa ?", vbYesNo) = vbYes Then
            Screen.MousePointer = vbHourglass

            'Atualizar o STATUS da capa para '1' -> Complementação
            Call AtualizaStatusCapa(aCapa(lstCapa.ListIndex + 1).IdCapa, "1")
            aCapa(lstCapa.ListIndex + 1).Status = "1"

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
        End If
    End If

    lstDocto.SetFocus
End Sub
Private Sub CmdFechar_Click()

    Unload Me
End Sub
Private Sub CmdFecharPesquisa_Click()

  Call CmdFechar_Click
End Sub
Public Sub cmdFrenteVerso_Click()

  On Error GoTo ERRO_FRENTEVERSO

  If teclou Then Exit Sub

  If FrmImagem.Visible = False Then Exit Sub

  teclou = True
  'só muda para frente/verso qdo docto vem da Ls500 e da Vips,
  'poi, o canon não gera verso.
  If (aDoc(lstDocto.ListIndex + 1).Ordem = "0") Or (aDoc(lstDocto.ListIndex + 1).Ordem = "2") Then
    If Lead1.Tag = "V" Then
        Lead1.Tag = "F"     'se verso, mostrar frente
        With Lead1
            .AutoRepaint = False
            If Geral.VIPSDLL = eDllProservi Then
              .Load Geral.DiretorioImagens & aDoc(lstDocto.ListIndex + 1).Frente, 0, 0, 1
            Else
              .Load Geral.DiretorioImagens & Format(aCapa(lstCapa.ListIndex + 1).IdLote, "000000000") & "\" & aDoc(lstDocto.ListIndex + 1).Frente, 0, 0, 1
            End If
  
            'se ls500 mostrar mais escuro
            If (aDoc(lstDocto.ListIndex + 1).Ordem = "2") Then
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
              .Load Geral.DiretorioImagens & Trim(aDoc(lstDocto.ListIndex + 1).Verso), 0, 0, 1
            Else
              .Load Geral.DiretorioImagens & Format(aCapa(lstCapa.ListIndex + 1).IdLote, "000000000") & "\" & aDoc(lstDocto.ListIndex + 1).Verso, 0, 0, 1
            End If
  
            'se ls500 mostrar mais escuro
            If (aDoc(lstDocto.ListIndex + 1).Ordem = "2") Then
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
  FrmImagem.Visible = False
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Manipular Imagem do Documento.", Err, rdoErrors)
    Case vbCancel, vbRetry
      Unload Me
  End Select
End Sub

Private Sub cmdIlegiveis_Click()

    Dim rst         As RDO.rdoResultset
    Dim sStr        As String

    'Verificar se há alguma capa selecionada
    If lstCapa.ListIndex = -1 Then
        MsgBox "Nenhum Envelope / Malote selecionado.", vbInformation, App.Title
        lstDocto.SetFocus
        Exit Sub
    End If

    Call AtualizaStatusCapa(aCapa(lstCapa.ListIndex + 1).IdCapa, "5")
    
    '''''''''''''''''''''''''''''''
    'Verifica se existe comentario'
    '''''''''''''''''''''''''''''''
    Set rst = GetControleCapa(Geral.DataProcessamento, aCapa(lstCapa.ListIndex + 1).IdCapa)
    
    sStr = ""
    If Not rst.EOF() Then
        sStr = rst!Comentario
    End If
    '''''''''''''''''''''''''''''''''
    'Insere registro no ControleCapa'
    '''''''''''''''''''''''''''''''''
    If Not InsereControleCapa(Geral.DataProcessamento, aCapa(lstCapa.ListIndex + 1).IdCapa, sStr, 7) Then
        MsgBox "Não foi possível inserir o Controle de Capa.", vbExclamation
    End If



    IdSelecionado = 0

    lblOcorrencia.Caption = ""
    If lstCapa.ListIndex < lstCapa.ListCount - 1 Then
        lstCapa.ListIndex = lstCapa.ListIndex + 1
    Else
        Call CmdAtualizar_Click
    End If
    lstDocto.SetFocus
End Sub
Public Sub cmdInverteCor_Click()

  On Error GoTo ERRO_INVERTECOR

  If teclou Then Exit Sub

  If FrmImagem.Visible = False Then Exit Sub

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
                If CDbl(lstCapa.List(iIndex)) = CDbl(txtNumEnvMal.Text) Then
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
                'A Capa anterior não sofreu alteração , Voltar o Status para 'A'
                Call AtualizaStatusCapa(aCapa(lstCapa.ListIndex + 1).IdCapa, "A")
            End If
        End If
        
        lstCapa.ListIndex = -1
        lstDocto.Clear
        FrmImagem.Visible = False
        Screen.MousePointer = vbDefault
    End If
End Sub

Public Sub cmdRotacao_Click()

  On Error GoTo ERRO_ROTACAO

  If teclou Then Exit Sub

  If FrmImagem.Visible = False Then Exit Sub

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

Public Sub cmdZoomMais_Click()

  On Error GoTo ERRO_ZOOMMAIS

  If teclou Then Exit Sub
  
  If FrmImagem.Visible = False Then Exit Sub

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

  If FrmImagem.Visible = False Then Exit Sub

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
Private Sub MostraImagem()

    On Error GoTo ERRO_MOSTRAIMAGEM

    Dim i     As Integer
    Dim Ret   As Long

    hCtl = Lead1.hwnd

    'Coloca imagem na tela
    With Lead1
        .Tag = "F"
        .AutoRepaint = False
        If Geral.VIPSDLL = eDllProservi Then
            .Load Geral.DiretorioImagens & aDoc(lstDocto.ListIndex + 1).Frente, 0, 0, 1
        Else
            .Load Geral.DiretorioImagens & Format(aCapa(lstCapa.ListIndex + 1).IdLote, "000000000") & "\" & aDoc(lstDocto.ListIndex + 1).Frente, 0, 0, 1
        End If

        'Se imagem for da ls500, deixar mais escura
        If aDoc(lstDocto.ListIndex + 1).Ordem <> "2" Then
            .Intensity 220
        Else
            .Intensity 150
        End If

        'Se imagem for do canon, diminui em 50% o tamanho
        If aDoc(lstDocto.ListIndex + 1).Ordem <> "1" Then
            .PaintZoomFactor = 100
        Else
            .PaintZoomFactor = 50
        End If
        .AutoRepaint = True
    End With

    FrmImagem.Visible = True

    'Posiciona imagem sempre no começo
    Ret = SendMessage(hCtl, WM_VSCROLL, SB_TOP, 0)
    Ret = SendMessage(hCtl, WM_HSCROLL, SB_TOP, 0)

    'Habilita Objetos de Manipulação de Imagens
    Call HDObjetosImagem(True)

    DoEvents

    Exit Sub

ERRO_MOSTRAIMAGEM:
    Screen.MousePointer = vbDefault
    MsgBox "Não foi possível exibir a Imagem do Documento, imagem não encontrada.", vbInformation, App.Title
    Call HDObjetosImagem(False)
End Sub
Private Sub Form_Activate()

  On Error GoTo ERRO_ACTIVATE

  'Inclusão de chamada a rotina AtualizaAtividade
  Call AtualizaAtividade(7)
  'Call HabilitaRecaptura
  Recaptura.CmdCapturar.Enabled = True

  With Lead1
      .UnlockSupport L_SUPPORT_EXPRESS, "YXPQ3XPPVT"
      .UnlockSupport L_SUPPORT_GIFLZW, "0K3RV9UY3EY"
      .UnlockSupport L_SUPPORT_TIFLZW, "9LE75L0FDXHK"
  End With

  'Preencher List com as Capas para Recaptura
  If PrimeiraVez Then
    PrimeiraVez = False

    FileLog = FreeFile
    Open Geral.DiretorioTrabalho & "DIG" & Format(Now, "ddmm") & ".TXT" For Append As #FileLog

    'obtem o path do windows para leitura/gravação do canon.ini
    PathIni = String(256, " ")
    GetWindowsDirectory PathIni, 255       'obtem o diretorio do windows
    PathIni = Trim(PathIni)
    PathIni = Left(PathIni, Len(PathIni) - 1) & "\"

    AlterouDocto = False
    If Not PreencheListCapas Then
      MsgBox "Não Existem Envelopes / Malotes para Recaptura.", vbInformation, App.Title

      Call HabilitaTimerPesquisa

      Exit Sub
    End If
    sTempo = 0

    'Habilitar o Timer de Atualização
    tmrAtualiza.Enabled = True
  End If

  Exit Sub

ERRO_ACTIVATE:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Ativar Tela.", Err, rdoErrors)
    Case vbCancel, vbRetry
      Unload Me
  End Select
End Sub
Private Function Digitalizar(ByVal pvbAppend As Boolean, ByRef QtdDoctos As Integer) As Boolean

    Dim dtInicio            As Date
    Dim dtFim               As Date
    Dim iRet                As Long
    Dim NumInicial          As Long
    Dim Estacao             As Long
    Dim GaugeTop, GaugeLeft As Long

    Dim Count               As Integer
    Dim sSql                As String
    
    Dim sFrente             As String
    Dim sVerso              As String
    
    Dim RsUltimaImagem      As rdoResultset
    Dim SeqInic             As Long
    Dim Opcao               As Long

    On Error GoTo ErroGetImagem
    rdoErrors.Clear

    Digitalizar = True
    
    sFrente = ""
    sVerso = ""

    'Selecionar o campo ProxImagem da tabela PARAMETRO
    With qryGetImagem
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2).Direction = rdParamOutput
        .Execute
        If .rdoParameters(0) <> 0 Then
            Screen.MousePointer = vbDefault
            MsgBox "Erro na obtenção do número da imagem.", vbCritical + vbOKOnly, App.Title
            Digitalizar = False
            Exit Function
        End If
        NumInicial = .rdoParameters(2)
    End With
    Print #FileLog, "Usuario: " & Geral.Usuario & " Horario: " & Format(Now, "hh:mm:ss") & " - Obtencao do numero da proxima imagem"

    On Error GoTo ErroDigitalizar

    'Verificar qual o numero da ultima imagem gravada
    sSql = Geral.DataProcessamento & " , " & aCapa(lstCapa.ListIndex + 1).IdLote

    Set qryGetUltimaImagemLote = Geral.Banco.CreateQuery("", "{call GetUltimaImagemLote (" & sSql & ")}")

    Set RsUltimaImagem = qryGetUltimaImagemLote.OpenResultset(rdOpenStatic, rdConcurReadOnly)

    If Not RsUltimaImagem.EOF Then
        SeqInic = Val(RsUltimaImagem!SeqInic) + 1
    Else
        MsgBox "Não foi possível selecionar a última imagem gravada.", vbInformation + vbOKOnly, App.Title
        Digitalizar = False
        Exit Function
    End If

    'Iniciar o contador de tempo de captura
    dtInicio = Now

    Me.Refresh

    pvbAppend = False
    Opcao = 2
    Do While Opcao = 2
        
        If Geral.Scanner = escnVIPS Then
            If Geral.VIPSDLL = eDllNovaUBB Then
                GaugeTop = TwipsYToPixel(Me.Top + 400)
                GaugeLeft = TwipsXToPixel(Me.Left + 9600)

                iRet = DigitalizarVIPS(Val(Geral.AgenciaCentral), Val(Right(CStr(aCapa(lstCapa.ListIndex + 1).IdLote), 5)), SeqInic, Geral.Estacao, GaugeTop, GaugeLeft, pvbAppend)
                
                'tratamento do erro
                If iRet = SC_Erro Then
                    Digitalizar = False
                End If
            Else
                iRet = VIPS_Recaptura(Val(Geral.AgenciaCentral), aCapa(lstCapa.ListIndex + 1).IdLote, SeqInic, Geral.DiretorioDados & Geral.RetornoFinal, pvbAppend)
                Print #FileLog, "Usuario: " & Geral.Usuario & " Horario: " & Format(Now, "hh:mm:ss") & "VIPS_Captura: " & Trim(str(iRet))
                Select Case iRet
                    Case 0
                        Digitalizar = True
                    Case -50
                        MsgBox "Foi detectado documento duplo na captura. Retorno: " & iRet & _
                                vbCrLf & "Repasse o documento.", _
                                vbExclamation + vbOKOnly, App.Title
                        Digitalizar = False
                    Case -1
                        MsgBox "Foi detectado um deslizamento na captura. Retorno: " & iRet, vbExclamation + vbOKOnly, App.Title
                        Digitalizar = False
                    Case -51, -55, -59, -61
                        MsgBox "Foi detectado um atolamento na captura. Retorno: " & iRet & _
                               vbCrLf & "Repasse o documento.", _
                               vbExclamation + vbOKOnly, App.Title
                        Digitalizar = False
                    Case -105
                        MsgBox "Não foi possível digitalizar todo o lote por uma falha de comunicação do scanner. Desligue a VIPS e ligue-a novamente.", vbExclamation + vbOKOnly, App.Title
                        VIPS_Reset
                        Digitalizar = False
                    Case Else
                        MsgBox "Não foi possível digitalizar todo o lote. Codigo de erro: " & iRet, vbExclamation + vbOKOnly, App.Title
                        Digitalizar = False
                End Select
            End If
            
        ElseIf Geral.Scanner = escnCanonLS500 Then
            GaugeTop = TwipsYToPixel(Me.Top + 400)
            GaugeLeft = TwipsXToPixel(Me.Left + 9600)
            Screen.MousePointer = vbHourglass

            iRet = DigitalizarCanonLS500(NumInicial, Geral.Estacao, GaugeTop, GaugeLeft, pvbAppend)

            Screen.MousePointer = vbDefault
            If iRet <> 1 Then
                If iRet = 0 Then
                    MsgBox "Verifique se o scanner contém documentos.", vbExclamation + vbOKOnly, App.Title
                Else
                    MsgBox "Ocorreu o seguinte erro na digitalização com o scanner LS500/Canon: " & iRet, vbExclamation + vbOKOnly, App.Title
                End If
                Digitalizar = False
            End If
        Else
            Opcao = 0
        End If

        dtFim = Now

        If Digitalizar = True Then
            'Varre o arquivo de retorno e obtem o nome da ultima imagem
            
            If Not ObtemUltimaImagem(Geral.DiretorioDados & Geral.RetornoFinal, sFrente, sVerso) Then

                Screen.MousePointer = vbDefault
                MsgBox "Não foi possível obter a última imagem capturada. Repasse os documentos.", vbExclamation + vbOKOnly, App.Title
                Digitalizar = False
                Print #FileLog, "Usuario: " & Geral.Usuario & " Horario: " & Format(Now, "hh:mm:ss") & " - Nao foi localizada a imagem do ultimo documento"
                Opcao = 0
            Else
            
                Geral.Documento.Frente = sFrente
                Geral.Documento.Verso = sVerso

                If Geral.Scanner = escnVIPS Then
                    Load ConfFimLote
                    On Error Resume Next
                    'Verso
                    
                    With ConfFimLote.Lead2
                        .AutoRepaint = False
                        If Geral.VIPSDLL = eDllProservi Then
                            .Load Geral.DiretorioImagens & Geral.Documento.Verso, 0, 0, 1
                        Else
                            .Load Geral.DiretorioImagens & Mid(Geral.RetornoFinal, 9, 9) & "\" & Geral.Documento.Verso, 0, 0, 1
                        End If
                        .Intensity 220
                        .PaintZoomFactor = 100
                        .AutoRepaint = True
                    End With

                    'Frente
                    With ConfFimLote.Lead1
                        .AutoRepaint = False
                        If Geral.VIPSDLL = eDllProservi Then
                            .Load Geral.DiretorioImagens & Geral.Documento.Frente, 0, 0, 1
                        Else
                            .Load Geral.DiretorioImagens & Mid(Geral.RetornoFinal, 9, 9) & "\" & Geral.Documento.Frente, 0, 0, 1
                        End If
                        .Intensity 220
                        .PaintZoomFactor = 100
                        .AutoRepaint = True
                    End With

                    Screen.MousePointer = vbDefault

                    'Modificar caption dos botões
                    ConfFimLote.Command1.Caption = "Confirmar"
                    ConfFimLote.Command2.Caption = "Continuar Recaptura"
                    ConfFimLote.Command3.Caption = "Cancelar"

                    'Exibe tela com a ultima imagem a ser gravada
                    ConfFimLote.Show vbModal, Me
                    
                    ConfFimLote.Lead1.Load 0, 0, 0, 0
                    ConfFimLote.Lead2.Load 0, 0, 0, 0
                    
                    Opcao = ConfFimLote.Resposta
                    Unload ConfFimLote
                    Principal.Refresh

                    Select Case Opcao
                        Case 0 'Cancelou o Lote
                            Print #FileLog, "Usuario: " & Geral.Usuario & " Horario: " & Format(Now, "hh:mm:ss") & " - Cancelamento da captura do lote"
                            MsgBox "Redigitalize este lote, pois o mesmo não foi gravado.", vbExclamation + vbOKOnly, App.Title
                            Digitalizar = False
                        Case 1 'Confirmou o Lote
                            Print #FileLog, "Usuario: " & Geral.Usuario & " Horario: " & Format(Now, "hh:mm:ss") & " - Confirmacao da captura do lote"
                        Case 2 'Continuar capturando
                            pvbAppend = True
                            Print #FileLog, "Usuario: " & Geral.Usuario & " Horario: " & Format(Now, "hh:mm:ss") & " - Continuacao da captura no mesmo lote"
                    End Select
                End If
            End If
        Else
            Print #FileLog, "Usuario: " & Geral.Usuario & " Horario: " & Format(Now, "hh:mm:ss") & " - Erro na captura das imagens dos documentos"
        End If

        SeqInic = SeqInic + Count
        DoEvents
        Me.Refresh
    Loop

    If Digitalizar = True Then
        If Not ProcessaArquivoRetorno(Geral.DiretorioDados & Geral.RetornoFinal, Count) Then
            MsgBox "Não foi possível Capturar os documentos. Tente novamente.", vbInformation + vbOKOnly, App.Title
            Digitalizar = False
            Exit Function
        End If
    End If

    Me.Refresh

    Print #FileLog, "Usuario: " & Geral.Usuario & " Horario: " & Format(Now, "hh:mm:ss") & " - Termino do processamento do arquivo de retorno"

    QtdDoctos = Count

    Exit Function

ErroGetImagem:
    Select Case TratamentoErro("Erro na obtenção do número da imagem.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    Exit Function

ErroDigitalizar:
    TratamentoErro "Erro na digitalização da imagens.", Err, rdoErrors, False
End Function
Private Function ObtemUltimaImagem(ByVal NomeArq As String, _
                                    ByRef Frente As String, ByRef Verso As String) As Boolean

    Dim Arq             As Integer
    Dim RetornoProservi As tpRetornoFinal
    Dim RetornoUnibanco As tpRetornoVips
    Dim RetornoNovaDll  As tpRetornoVipsNovaDLL

    Arq = FreeFile
    Open NomeArq For Binary As #Arq

    RetornoUnibanco.Frente = String(19, "  ")
    RetornoUnibanco.Verso = String(19, "  ")
    
    RetornoNovaDll.Frente = String(19, "  ")
    RetornoNovaDll.Verso = String(19, "  ")

    If Geral.VIPSDLL = eDllNovaUBB Then
        Get #Arq, , RetornoNovaDll
        While Not EOF(Arq)
            Frente = RetornoNovaDll.Frente
            Verso = IIf(Trim(RetornoNovaDll.Verso) = "", RetornoNovaDll.Frente, RetornoNovaDll.Verso)
            Get #Arq, , RetornoNovaDll
        Wend
    Else
        Get #Arq, , RetornoUnibanco
        While Not EOF(Arq)
            Frente = RetornoUnibanco.Frente
            Verso = IIf(Trim(RetornoUnibanco.Verso) = "", RetornoUnibanco.Frente, RetornoUnibanco.Verso)
            Get #Arq, , RetornoUnibanco
        Wend
    End If

    Close #Arq

    If Trim(Frente) <> "" Then
        ObtemUltimaImagem = True
    Else
        ObtemUltimaImagem = False
    End If
End Function
Private Function ProcessaArquivoRetorno(ByVal NomeArq As String, ByRef Count As Integer) As Boolean

    If Geral.Scanner = escnCanonLS500 Then

        ProcessaArquivoRetorno = ProcessaArquivoRetornoCanonLS500(NomeArq, Count)
    Else
        ProcessaArquivoRetorno = ProcessaArquivoRetornoUnibanco(NomeArq, Count)
    End If
End Function
Private Function ProcessaArquivoRetornoCanonLS500(ByVal NomeArq As String, ByRef Count As Integer) As Boolean

    Dim Arq             As Integer
    Dim IdCapa          As Integer
    Dim TipoDoc         As Integer
    Dim CountCapas      As Integer
    Dim Campo1          As String
    Dim Campo2          As String
    Dim Campo3          As String
    Dim Valor           As String
    Dim bVirtual        As Boolean
    Dim IdEnv_Mal       As String
    Dim Linha           As tpRetornoFinal
    Dim DirDestino      As String
    Dim RecCapa         As Long

    Count = 0
    CountCapas = 0
    IdCapa = 0

    On Error Resume Next
    DirDestino = Geral.DiretorioImagens & Format(aCapa(lstCapa.ListIndex + 1).IdLote, "000000000")

    MkDir (DirDestino)

    On Error GoTo ErroCaptura

    Geral.Banco.BeginTrans

    'Grava primeiro registros com codigo de barras
    Arq = FreeFile
    Open NomeArq For Binary As #Arq

    'Verifica se o arquivo de retorno esta vazio
    If LOF(Arq) = 0 Then
        Geral.Banco.RollbackTrans
        Exit Function
    End If

    Get #Arq, , Linha
    Do While Not EOF(Arq)
        If Linha.Tipo <> "A" Then

            'Gravar Documento
            Valor = "000"
            qryInsereDocto.rdoParameters(0).Direction = rdParamReturnValue
            qryInsereDocto.rdoParameters(10).Direction = rdParamOutput
            qryInsereDocto.rdoParameters(1) = Geral.DataProcessamento
            qryInsereDocto.rdoParameters(2) = aCapa(lstCapa.ListIndex + 1).IdCapa
            qryInsereDocto.rdoParameters(3) = 0
            qryInsereDocto.rdoParameters(4) = Trim(Linha.Leitura)
            qryInsereDocto.rdoParameters(5) = Val(Valor) / 100
            qryInsereDocto.rdoParameters(6) = Linha.Frente
            qryInsereDocto.rdoParameters(7) = IIf(Trim(Linha.Verso) = "", Linha.Frente, Linha.Verso)
            qryInsereDocto.rdoParameters(8) = Linha.Origem
            qryInsereDocto.rdoParameters(9) = OrdemCapturaInicial
            qryInsereDocto.Execute

            If qryInsereDocto.rdoParameters(0) <> 0 Then
                GoTo ErroCaptura
            End If

            Call GravaLog(aCapa(lstCapa.ListIndex + 1).IdCapa, qryInsereDocto.rdoParameters(10).Value, 201)

            Count = Count + 1
            OrdemCapturaInicial = OrdemCapturaInicial + 1

            On Error GoTo ErroCopiaImagem
            If Len(Trim(Linha.Frente)) > 0 Then
                FileCopy Geral.DiretorioImagens & Trim(Linha.Frente), DirDestino & "\" & Trim(Linha.Frente)
                On Error GoTo 0
                Kill Geral.DiretorioImagens & Trim(Linha.Frente)
            End If
            On Error GoTo ErroCopiaImagem
            If Len(Trim(Linha.Verso)) > 0 Then
                FileCopy Geral.DiretorioImagens & Trim(Linha.Verso), DirDestino & "\" & Trim(Linha.Verso)
                On Error GoTo 0
                Kill Geral.DiretorioImagens & Trim(Linha.Verso)
            End If
            On Error GoTo ErroCaptura
            
        End If
        Get #Arq, , Linha
    Loop
    Close #Arq

    'Ler novamente o arquivo gravando os documentos com CMC7
    Arq = FreeFile
    Open NomeArq For Binary As #Arq
    Get #Arq, , Linha

    While Not EOF(Arq)
        If Linha.Tipo = "A" Then

            Linha.Leitura = TrataLeitura(Linha.Leitura)
            TipoDoc = 0
            Valor = "000"

            'Gravar Documento
            qryInsereDocto.rdoParameters(0).Direction = rdParamReturnValue
            qryInsereDocto.rdoParameters(10).Direction = rdParamOutput
            qryInsereDocto.rdoParameters(1) = Geral.DataProcessamento
            qryInsereDocto.rdoParameters(2) = aCapa(lstCapa.ListIndex + 1).IdCapa
            qryInsereDocto.rdoParameters(3) = 0
            If Linha.Tipo = "A" Then ' Docto com CMC7
                Valor = ""
                TratarCamposCMC7 Linha.Leitura, Campo1, Campo2, Campo3, Valor
                qryInsereDocto.rdoParameters(4) = Campo1 & Campo2 & Campo3
            ElseIf Linha.Tipo = "B" Then ' Docto com Cod Barras
                qryInsereDocto.rdoParameters(4) = RPad(Trim(Linha.Leitura), 44)
            Else
                qryInsereDocto.rdoParameters(4) = Trim(Linha.Leitura)
            End If
            qryInsereDocto.rdoParameters(5) = Val(Valor) / 100
            qryInsereDocto.rdoParameters(6) = Linha.Frente
            qryInsereDocto.rdoParameters(7) = IIf(Trim(Linha.Verso) = "", Linha.Frente, Linha.Verso)
            qryInsereDocto.rdoParameters(8) = Linha.Origem
            qryInsereDocto.rdoParameters(9) = OrdemCapturaInicial
            qryInsereDocto.Execute

            If qryInsereDocto.rdoParameters(0) <> 0 Then
                GoTo ErroCaptura
            End If

            Call GravaLog(aCapa(lstCapa.ListIndex + 1).IdCapa, qryInsereDocto.rdoParameters(10).Value, 201)

            Count = Count + 1
            OrdemCapturaInicial = OrdemCapturaInicial + 1

            On Error GoTo ErroCopiaImagem
            If Len(Trim(Linha.Frente)) > 0 Then
                FileCopy Geral.DiretorioImagens & Trim(Linha.Frente), DirDestino & "\" & Trim(Linha.Frente)
                On Error GoTo 0
                Kill Geral.DiretorioImagens & Trim(Linha.Frente)
            End If
            On Error GoTo ErroCopiaImagem
            If Len(Trim(Linha.Verso)) > 0 Then
                FileCopy Geral.DiretorioImagens & Trim(Linha.Verso), DirDestino & "\" & Trim(Linha.Verso)
                On Error GoTo 0
                Kill Geral.DiretorioImagens & Trim(Linha.Verso)
            End If
            On Error GoTo ErroCaptura
        
        End If
        Get #Arq, , Linha
    Wend
    Close #Arq

    Geral.Banco.CommitTrans

    ProcessaArquivoRetornoCanonLS500 = True

    Exit Function

ErroCopiaImagem:
    If MsgBox("Erro ao gravar a imagem do documento. Deseja tentar novamente?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        Resume
        Exit Function
    End If

ErroCaptura:
    Geral.Banco.RollbackTrans

    TratamentoErro "Erro no processamento do arquivo de retorno.", Err, rdoErrors, False
    ProcessaArquivoRetornoCanonLS500 = False
End Function
Private Function ProcessaArquivoRetornoUnibanco(ByVal NomeArq As String, ByRef Count As Integer) As Boolean

    Dim Arq             As Integer
    Dim IdCapa          As Integer
    Dim TipoDoc         As Integer
    Dim CountCapas      As Integer
    Dim Campo1          As String
    Dim Campo2          As String
    Dim Campo3          As String
    Dim Valor           As String
    Dim bVirtual        As Boolean
    Dim IdEnv_Mal       As String
    Dim Linha           As tpRetornoVips
    Dim LinhaNovaDLL    As tpRetornoVipsNovaDLL

    Count = 0
    CountCapas = 0
    IdCapa = 0

    On Error GoTo ErroCaptura

    Arq = FreeFile
    Open NomeArq For Binary As #Arq

    If Geral.VIPSDLL = eDllNovaUBB Then
        Get #Arq, , LinhaNovaDLL
    Else
        Get #Arq, , Linha
    End If

    'Se o Arquivo de Retorno estiver vazio , nenhum documento capturado
    If EOF(Arq) Then
        ProcessaArquivoRetornoUnibanco = False
        Close #Arq
        Exit Function
    End If

    Geral.Banco.BeginTrans

    While Not EOF(Arq)
        If Geral.VIPSDLL = eDllNovaUBB Then
            LinhaNovaDLL.Leitura = TrataLeitura(LinhaNovaDLL.Leitura)
        Else
            Linha.Leitura = TrataLeitura(Linha.Leitura)
        End If
        
        TipoDoc = 0

        'Inserir Documento
        Valor = "000"
        qryInsereDocto.rdoParameters(0).Direction = rdParamReturnValue
        qryInsereDocto.rdoParameters(10).Direction = rdParamOutput
        qryInsereDocto.rdoParameters(1) = Geral.DataProcessamento
        qryInsereDocto.rdoParameters(2) = aCapa(lstCapa.ListIndex + 1).IdCapa
        qryInsereDocto.rdoParameters(3) = TipoDoc

        If Geral.VIPSDLL = eDllNovaUBB Then
            If LinhaNovaDLL.Tipo = "A" Then 'Docto com CMC7
                Valor = ""
                TratarCamposCMC7 LinhaNovaDLL.Leitura, Campo1, Campo2, Campo3, Valor
                qryInsereDocto.rdoParameters(4) = Campo1 & Campo2 & Campo3
            ElseIf LinhaNovaDLL.Tipo = "B" Then ' Docto com Cod Barras
                qryInsereDocto.rdoParameters(4) = RPad(Trim(LinhaNovaDLL.Leitura), 44)
            Else
                qryInsereDocto.rdoParameters(4) = Trim(LinhaNovaDLL.Leitura)
            End If
        Else
            If Linha.Tipo = "A" Then 'Docto com CMC7
                Valor = ""
                TratarCamposCMC7 Linha.Leitura, Campo1, Campo2, Campo3, Valor
                qryInsereDocto.rdoParameters(4) = Campo1 & Campo2 & Campo3
            ElseIf Linha.Tipo = "B" Then ' Docto com Cod Barras
                qryInsereDocto.rdoParameters(4) = RPad(Trim(Linha.Leitura), 44)
            Else
                qryInsereDocto.rdoParameters(4) = Trim(Linha.Leitura)
            End If
        End If

        qryInsereDocto.rdoParameters(5) = Val(Valor) / 100
        If Geral.VIPSDLL = eDllNovaUBB Then
            qryInsereDocto.rdoParameters(6) = LinhaNovaDLL.Frente
            qryInsereDocto.rdoParameters(7) = IIf(Trim(LinhaNovaDLL.Verso) = "", LinhaNovaDLL.Frente, LinhaNovaDLL.Verso)
            qryInsereDocto.rdoParameters(8) = "0"
        Else
            qryInsereDocto.rdoParameters(6) = Linha.Frente
            qryInsereDocto.rdoParameters(7) = IIf(Trim(Linha.Verso) = "", Linha.Frente, Linha.Verso)
            qryInsereDocto.rdoParameters(8) = Linha.Origem
        End If
        qryInsereDocto.rdoParameters(9) = OrdemCapturaInicial
        qryInsereDocto.Execute

        If qryInsereDocto.rdoParameters(0) <> 0 Then
            GoTo ErroCaptura
        End If

        Call GravaLog(aCapa(lstCapa.ListIndex + 1).IdCapa, qryInsereDocto.rdoParameters(10), 201)
        Count = Count + 1
        OrdemCapturaInicial = OrdemCapturaInicial + 1
        
        If Geral.VIPSDLL = eDllNovaUBB Then
            Get #Arq, , LinhaNovaDLL
        Else
            Get #Arq, , Linha
        End If
    Wend

    Close #Arq

    Geral.Banco.CommitTrans
    ProcessaArquivoRetornoUnibanco = True

    Exit Function

ErroCaptura:
    If aCapa(lstCapa.ListIndex + 1).IdCapa >= 0 Then
        Geral.Banco.RollbackTrans
    End If

    ProcessaArquivoRetornoUnibanco = False
End Function
Private Function TrataLeitura(ByVal Leitura As String) As String

    Dim bInvalido   As Boolean
    Dim Count       As Integer
    Dim Result      As String
    Dim Char        As String * 1

    Result = ""
    Leitura = Trim(Leitura)
    bInvalido = False
    For Count = 1 To Len(Leitura)
        Char = Mid(Leitura, Count, 1)
        If Char <> "<" And Char <> ">" And Char <> ":" And Char <> ";" Then
            If (Not bInvalido) And (Not IsNumeric(Char)) Then
                bInvalido = True
            End If
            If bInvalido Then
                Result = Result & "0"
            Else
                Result = Result & Char
            End If
        End If
    Next
    If Val(Result) > 0 Then
        TrataLeitura = Result
    Else
        TrataLeitura = ""
    End If
End Function
Private Function DigitalizarCanonLS500(ByVal NumInicial As Long, ByVal Estacao As Long, ByVal GaugeTop As Long, ByVal GaugeLeft As Long, Optional ByVal pvbAppend As Boolean = False) As Long

    Dim iRet, Ret_LS, Ret_Canon As Long
    Dim ImgInicialCanon         As Long
    Dim ArqCanon                As String
    Dim NumLote                 As Long
    Dim Retc                    As Integer
    Dim Arq                     As String

    '---- Função para fazer ou não o append no arquivo de retorno da ls500.
    If pvbAppend Then
        iRet = LS_SetAppend(1)      'faz append no arquivo
    Else
        iRet = LS_SetAppend(0)      'não faz append no arquivo
    End If

    iRet = LS_SetGaugePos(GaugeTop, GaugeLeft)

    iRet = LS_SetFileName(Geral.DiretorioDados & Geral.RetornoFinal)

    iRet = LS_SetCanon(1)   ' Habilita para LS + Canon

    Ret_LS = LS_Digitaliza1(Geral.DiretorioImagens, NumInicial, Estacao)

    iRet = LS_SetCanon(0)   ' Desabilita para LS + Canon

Redigitaliza_Canon:

    ImgInicialCanon = LS_GetLastImage()
    NumLote = LS_GetNumLote()

On Error GoTo ERRO_CANON

    'Digitaliza Doctos via PIXEL TRANSLATIONS

    ArqCanon = "Lote" & Format$(ImgInicialCanon, "0000") & ".TIF"

    'COORDENADAS BÁSICAS PARA PIXEL ----
    EzCanon.Close
    EzCanon.ScanResolution = 200                '200 DPI
    EzCanon.ScanPackaging = &H30000             'TIF
    EzCanon.ScanMultipage = 1                   'múltiplas páginas
    EzCanon.Visible = True
    Arq = Dir(PathIni & "canon.ini")

    If Arq = "" Then
        EzCanon.ScanMoreDialog                  'page size detection
        Retc = EzCanon.ScanStateWrite(PathIni & "canon.ini", "configuracao")
    Else
        Retc = EzCanon.ScanStateRead(PathIni & "canon.ini", "configuracao")
    End If

    DoEvents
    EzCanon.ScanFileName = Geral.DiretorioImagens & ArqCanon

    iRet = UT_GaugeCanonInit(GaugeTop, GaugeLeft)
    iRet = UT_GaugeCanon()

    '--- aciona digitalização no CANON ---
    Ret_Canon = EzCanon.ScanBatch

    iRet = UT_GaugeCanon()
    UT_DestroyGaugeCanon

    EzCanon.Close    'fecha pixel para não dar acesso denied
    EzCanon.Visible = False

    iRet = UT_DesmembraTiff(Geral.DiretorioImagens & ArqCanon, Geral.DiretorioDados & Geral.RetornoFinal, NumLote, ImgInicialCanon, Estacao)

    DigitalizarCanonLS500 = 1

    Exit Function

ERRO_CANON:
    UT_DestroyGaugeCanon
    Screen.MousePointer = vbDefault
    Beep

    Select Case Err
        Case 3044
            MsgBox "Diretório de Imagem Inválido! Verifique e redigitalize somente os documentos deste equipamento. Erro: " + Error, vbCritical + vbOKOnly, App.Title
        Case 3050
            MsgBox "SHARE Não instalado! Finalize o WINDOWS e carregue o SHARE.EXE. Erro: " + Error, vbCritical + vbOKOnly, App.Title
        Case Else
            MsgBox "Desligue o scanner CANON e ligue-o novamente para redigitalizar somente os documentos deste equipamento. Erro: " + Error, vbCritical + vbOKOnly, App.Title
    End Select
    Screen.MousePointer = vbHourglass

    EzCanon.Close
    If Dir(Geral.DiretorioImagens & ArqCanon) <> "" Then
        Kill Geral.DiretorioImagens & ArqCanon
    End If
    GoTo Redigitaliza_Canon

    EzCanon.Close
    EzCanon.Visible = False

    DigitalizarCanonLS500 = -5  ' Erro Canon
    
End Function
Private Function TwipsXToPixel(ByVal Twips As Long) As Long

    TwipsXToPixel = Int(Twips / Screen.TwipsPerPixelX)
End Function
Private Function TwipsYToPixel(ByVal Twips As Long) As Long

    TwipsYToPixel = Int(Twips / Screen.TwipsPerPixelY)
End Function
Public Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

  Dim Ret As Long

  hCtl = Recaptura.Lead1.hwnd

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
    Case vbKeyF11
      Call cmdFrenteVerso_Click
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

    PrimeiraVez = True
  
    'Definindo as querys da Captura
    Set qryProducaoScanner = Geral.Banco.CreateQuery("", "{ Call InsereProducaoScanner (?,?,?,?)}")
    Set qryGetImagem = Geral.Banco.CreateQuery("", "{? = Call GetImagem (?,?)}")
    Set qryInsereCapa = Geral.Banco.CreateQuery("", "{? = Call CapturaCapa (?,?,?,?,?)}")
    Set qryInsereDocto = Geral.Banco.CreateQuery("", "{? = Call CapturaDocumento (?,?,?,?,?,?,?,?,?,?)}")
    Set qryRemoveDocumento = Geral.Banco.CreateQuery("", "{ ? = Call RemoveDocumento (?,?,?)}")
    
End Sub
Private Sub Form_Unload(Cancel As Integer)

    'Verificar se foi selecionado uma Capa Anteriormente
    If lstCapa.ListIndex + 1 > 0 Then
        If aCapa(lstCapa.ListIndex + 1).Status <> "V" Then
            Call AtualizaStatusCapa(aCapa(lstCapa.ListIndex + 1).IdCapa, "A")
            aCapa(lstCapa.ListIndex + 1).Status = "A"
        End If
    End If

    'Limpando variável de controle
    IdSelecionado = 0

    'Limpando variável de arquivo de log da captura
    Close #FileLog

    'Desabilitar os timers
    tmrAtualiza.Enabled = False
    tmrPesquisa.Enabled = False

    'Finalizar Conexões
    Set qryGetCapa = Nothing
    Set qryGetDocumentos = Nothing
    Set qryAtualizaStatusCapa = Nothing
'    Set qryAtualizaStatusDocumento = Nothing
    Set qryGetOcorr = Nothing

    'Finalizar Conexões da Captura
    Set qryGetImagem = Nothing
    Set qryProducaoScanner = Nothing
    Set qryInsereCapa = Nothing
    Set qryInsereDocto = Nothing
    Set qryGetUltimaImagemLote = Nothing
End Sub

Private Sub Lead1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

  If Button = 1 Then
    Lead1.AutoRubberBand = True
    Lead1.MousePointer = vbCrosshair
  Else
    Call MostraImagem
  End If
End Sub
Sub HDMalote(ByVal bValor As Boolean)

    PctMalote.Visible = bValor
    lblNumMalote.Visible = bValor
    If bValor = False Then
        lblLote.Caption = ""
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

        If IdSelecionado <> 0 And (IdSelecionado <> aCapa(lstCapa.ListIndex + 1).IdCapa) Then

            Call AtualizaStatusCapa(IdSelecionado, "A")

        End If
        
       'Verifica se Capa Selecionado pertence a Lote com Capa em Recaptura
        Set qryGetCapaLoteEmRecaptura = Geral.Banco.CreateQuery("", "{? = call GetCapaLoteEmRecaptura (?,?,?)}")

        With qryGetCapaLoteEmRecaptura
            .rdoParameters(0).Direction = rdParamReturnValue
            .rdoParameters(1) = Geral.DataProcessamento             'Data de Processamento
            .rdoParameters(2) = aCapa(lstCapa.ListIndex + 1).IdLote 'IdLote
            .rdoParameters(3) = Geral.Intervalo                     'Intervalo de Atualização
    
            .Execute
        End With

        If qryGetCapaLoteEmRecaptura(0) = 0 Then
            Screen.MousePointer = vbDefault
            MsgBox "    Este Envelope / Malote não está disponível." & vbCrLf & "Pertence a um Lote com Documento em Recaptura.", vbInformation, App.Title
            Exit Sub
        End If
        
        'Verificar se a capa mudou
        If IdSelecionado = aCapa(lstCapa.ListIndex + 1).IdCapa Then
            Call PreencheListDocto(0)
        Else
            'Verificar se a capa selecionada continua disponivel
            Ret = CapaSelecionadaDisponivel
            If Ret = 0 Then
                'Verificar se existem documentos transmitidos/expedidos ou com NSU
                If VerificaDocumentosTransmitidos Then
                    Screen.MousePointer = vbDefault
                    Call CmdAtualizar_Click
                    Exit Sub
                End If

                'Excluir Ajustes , se houver
                If Not ExcluiAjuste Then Exit Sub

                Call AtualizaStatusCapa(aCapa(lstCapa.ListIndex + 1).IdCapa, "B")
                aCapa(lstCapa.ListIndex + 1).Status = "B"

                Call PreencheListDocto(0)

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
        lblOcorrencia.Caption = ""

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
Private Sub lstDocto_Click()

   On Error GoTo ERRO_DOCTOCLICK

   Dim RsOcorr As rdoResultset
   Dim sSql As String
   Dim X As Integer
   Dim sOcorrencia          As String

   'Exibir a Figura do Documento Selecionado
   Call MostraImagem

   lblOcorrencia.Caption = ""

   'Verifica se o Documento possui Ocorrência
'   If aDoc(LstDocto.ListIndex + 1).Status = "D" Or aDoc(LstDocto.ListIndex + 1).Status = "F" Then
      'Verificar se a ocorrencia começa com 999
        If Left(aDoc(lstDocto.ListIndex + 1).Ocorrencia, 3) = "999" Then
            If aDoc(lstDocto.ListIndex + 1).RetornoTransacao > 0 Then
                Call ObtemRetornoTransacao(aDoc(lstDocto.ListIndex + 1).RetornoTransacao, sOcorrencia)
                lblOcorrencia.Caption = sOcorrencia
            Else
                lblOcorrencia.Caption = "Erro operacional."
            End If
        Else
            'Verificar se o código da ocorrencia possui 3 ou 5 caracteres
            If Val(aDoc(lstDocto.ListIndex + 1).Ocorrencia) > 999 Then
               '5 Posicoes
               sSql = Left(Trim(aDoc(lstDocto.ListIndex + 1).Ocorrencia), 3)
            Else
               '3 Posicoes
               If Right(Trim(aDoc(lstDocto.ListIndex + 1).Ocorrencia), 2) = "00" Then
                  'Ocorrencia atualizada pelo robo
                  sSql = Val(Trim(aDoc(lstDocto.ListIndex + 1).Ocorrencia)) / 100
               Else
                  'Ocorrencia gerada pelo sistema
                  sSql = Val(Trim(aDoc(lstDocto.ListIndex + 1).Ocorrencia))
               End If
            End If
            
            Set qryGetOcorr = Geral.Banco.CreateQuery("", "{call GetOcorrencia (" & sSql & ")}")
            
            Set RsOcorr = qryGetOcorr.OpenResultset(rdOpenStatic, rdConcurReadOnly)
            
            lblOcorrencia.Caption = ""
            If Not RsOcorr.EOF Then
               lblOcorrencia.Caption = "Ocorrência : " & RsOcorr!Descricao
            End If
        End If
'   End If

    Exit Sub

ERRO_DOCTOCLICK:
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Selecionar Documento.", Err, rdoErrors)
    Case vbCancel, vbRetry
      Unload Me
  End Select
End Sub

Private Sub tmrAtualiza_Timer()

  tmrAtualiza.Enabled = False

  If lstCapa.ListIndex <> -1 Then
    If aCapa(lstCapa.ListIndex + 1).IdCapa <> 0 Then
      sTempo = sTempo + Int(tmrAtualiza.Interval / 1000)

      If sTempo + Int(tmrAtualiza.Interval / 1000) >= Geral.Intervalo Then
        'Atualizar o Status da Capa
        Call AtualizaStatusCapa(aCapa(lstCapa.ListIndex + 1).IdCapa, "B")

        sTempo = 0
      End If
    End If
  End If

  tmrAtualiza.Enabled = True
End Sub

Private Sub tmrPesquisa_Timer()

  tmrPesquisa.Enabled = False

  sTempo = sTempo + Int(tmrPesquisa.Interval / 1000)

  If sTempo + Int(tmrPesquisa.Interval / 1000) >= Geral.Atualizacao Then
    'Pesquisar por Documentos para Recaptura
    sTempo = 0
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

Function VerificaDoctosIndefinidos() As Boolean

  Dim X As Integer

  VerificaDoctosIndefinidos = False

  'Verificar se existe algum documento indefinido
  If lstDocto.ListCount > 0 Then
    VerificaDoctosIndefinidos = False
    For X = 0 To lstDocto.ListCount - 1
      If aDoc(X + 1).TipoDocto = 0 And Val(aDoc(X + 1).Ocorrencia) = 0 Then
        'Documento Indefinido
        VerificaDoctosIndefinidos = True
        Exit Function
      End If
      DoEvents
    Next X
  End If
End Function

Private Sub TxtNumEnvMal_KeyPress(KeyAscii As Integer)


   If KeyAscii = vbKeyReturn Then
      Call cmdProcurar_Click
   ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
      KeyAscii = 0
   End If

End Sub
Private Function VerificaDocumentosTransmitidos() As Boolean

   On Error GoTo VerificaDocumentosTransmitidos_Err

   Dim RsDoctosTrans As rdoResultset

   VerificaDocumentosTransmitidos = False

   Set qryGetDocumentosParaVerificacao = Geral.Banco.CreateQuery("", "{ ? = Call GetDocumentosParaVerificacao (?,?)}")
   

   With qryGetDocumentosParaVerificacao
      .rdoParameters(1) = Geral.DataProcessamento
      .rdoParameters(2) = aCapa(lstCapa.ListIndex + 1).IdCapa
   End With

   Set RsDoctosTrans = qryGetDocumentosParaVerificacao.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
   If Not RsDoctosTrans.EOF Then
      If RsDoctosTrans!Qtde > 0 Then
         VerificaDocumentosTransmitidos = True
         'Atualizar o Status da Capa para 'V' - Em Analise
         Call AtualizaStatusCapa(aCapa(lstCapa.ListIndex + 1).IdCapa, "V")

         'Gravar Log
         Call GravaLog(aCapa(lstCapa.ListIndex + 1).IdCapa, 0, 205)
         Call LimpaListaDocto
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
Private Function DigitalizarVIPS(ByVal lngAgencia As Long, ByVal lngLote As Long, ByVal pvnInicio As Long, ByVal pvnEstacao As Long, ByVal nTop As Long, ByVal nLeft As Long, Optional ByVal pvbAppend As Boolean = False) As Long
    
    Dim iRet    As Long
    Dim iRet1   As Long
    Dim iRet2   As Long
    
    DigitalizarVIPS = SC_Erro
    
    iRet = SC_SetGaugePos(nTop, nLeft)
    If iRet <> 1 Then
        Call ScanMessageErr(iRet)
        Exit Function
    End If
    
    If pvbAppend Then
        iRet = SC_SetAppend(1)
        If iRet <> 1 Then
            Call ScanMessageErr(iRet)
            Exit Function
        End If
    Else
        iRet = SC_SetAppend(0)
        If iRet <> 1 Then
            Call ScanMessageErr(iRet)
            Exit Function
        End If
    End If

    iRet = SC_AcquireBatch(lngAgencia, lngLote, pvnInicio, Geral.DiretorioDados & Geral.RetornoFinal, pvnEstacao)
    
    If iRet <> 1 Then
        Call ScanMessageErr(iRet)
        Exit Function
    End If
    
    DigitalizarVIPS = SC_OK
        
    
End Function

