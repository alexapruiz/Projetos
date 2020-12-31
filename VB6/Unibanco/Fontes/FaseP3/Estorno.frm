VERSION 5.00
Object = "{00080003-B1BA-11CE-ABC6-F5B2E79D9E3F}#8.0#0"; "LTOCX80N.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Estorno 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estorno de Capas/Documentos"
   ClientHeight    =   7428
   ClientLeft      =   636
   ClientTop       =   1260
   ClientWidth     =   11100
   ControlBox      =   0   'False
   ForeColor       =   &H00404040&
   Icon            =   "Estorno.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7428
   ScaleWidth      =   11100
   Begin VB.Timer tmrAtualiza 
      Interval        =   50000
      Left            =   30
      Top             =   5310
   End
   Begin VB.Frame FramePesquisaEstorno 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2505
      Left            =   0
      MouseIcon       =   "Estorno.frx":000C
      TabIndex        =   20
      Top             =   0
      Width           =   11085
      Begin VB.TextBox TxtNumTerminal 
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
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   5010
         MaxLength       =   3
         TabIndex        =   24
         Top             =   1770
         Width           =   1125
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   408
         Left            =   120
         Picture         =   "Estorno.frx":0316
         ScaleHeight     =   384
         ScaleWidth      =   384
         TabIndex        =   29
         Top             =   210
         Width           =   408
      End
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
         Height          =   800
         Left            =   8256
         Picture         =   "Estorno.frx":0620
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   900
         Width           =   852
      End
      Begin VB.CommandButton CmdPesquisa 
         Caption         =   "&Pesquisar"
         Enabled         =   0   'False
         Height          =   800
         Left            =   7320
         Picture         =   "Estorno.frx":092A
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   900
         Width           =   850
      End
      Begin VB.OptionButton OpPesqCapa 
         Caption         =   "&Capa  Env/Mal:"
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
         Left            =   210
         TabIndex        =   17
         Top             =   930
         Value           =   -1  'True
         Width           =   2040
      End
      Begin VB.TextBox TxtNumCapa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   2730
         MaxLength       =   14
         TabIndex        =   21
         Top             =   810
         Width           =   3405
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
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   2730
         MaxLength       =   11
         TabIndex        =   22
         Top             =   1290
         Width           =   3405
      End
      Begin VB.OptionButton OpPesqMalote 
         Caption         =   "&Número &Malote:"
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
         Left            =   240
         TabIndex        =   18
         Top             =   1410
         Width           =   2010
      End
      Begin VB.TextBox TxtNumNSU 
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
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   2730
         MaxLength       =   6
         TabIndex        =   23
         Top             =   1770
         Width           =   1425
      End
      Begin VB.OptionButton OpPesqNsu 
         Caption         =   "N&SU:"
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
         Left            =   210
         TabIndex        =   19
         Top             =   1860
         Width           =   1980
      End
      Begin VB.Label LblCaixa 
         AutoSize        =   -1  'True
         Caption         =   "CAIXA:"
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
         Left            =   4260
         TabIndex        =   30
         Top             =   1830
         Width           =   615
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Pesquisar Capa/Malote ou NSU de Documento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   840
         TabIndex        =   28
         Top             =   330
         Width           =   4875
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Estornar:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   4980
      TabIndex        =   16
      Top             =   6180
      Width           =   2955
      Begin VB.CommandButton cmdConfirma 
         Caption         =   "Confirmar Estorno"
         Height          =   800
         Left            =   1980
         Picture         =   "Estorno.frx":0C34
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   850
      End
      Begin VB.OptionButton OpEstornaVinculo 
         Caption         =   "&Vinculo"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   570
         Width           =   1005
      End
      Begin VB.OptionButton OpEstornaDocumento 
         Caption         =   "&Documento"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   870
         Value           =   -1  'True
         Width           =   1245
      End
      Begin VB.OptionButton OpEstornaCapa 
         Caption         =   "&Capa"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   270
         Width           =   1665
      End
   End
   Begin VB.Frame fraBotoesSuperiores 
      Caption         =   "Navegação:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   7950
      TabIndex        =   11
      Top             =   6180
      Width           =   2880
      Begin VB.CommandButton cmdDoctoPosterior 
         Caption         =   "Docto Posterior"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   1000
         Picture         =   "Estorno.frx":0F3E
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   250
         Width           =   850
      End
      Begin VB.CommandButton cmdDoctoAnterior 
         Caption         =   "Docto Anterior"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   135
         Picture         =   "Estorno.frx":1380
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   250
         Width           =   850
      End
      Begin VB.CommandButton cmdFinalizar 
         Cancel          =   -1  'True
         Caption         =   "Finalizar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   1850
         Picture         =   "Estorno.frx":17C2
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   250
         Width           =   850
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Imagem:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   300
      TabIndex        =   10
      Top             =   6180
      Width           =   4665
      Begin VB.CommandButton cmdRotacao 
         Caption         =   "Rotação"
         Height          =   800
         Left            =   1905
         Picture         =   "Estorno.frx":1ACC
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   250
         Width           =   825
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
         Height          =   800
         Left            =   2745
         Picture         =   "Estorno.frx":1DD6
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   250
         Width           =   850
      End
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
         Height          =   800
         Left            =   3600
         Picture         =   "Estorno.frx":20E0
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   250
         Width           =   850
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
         Height          =   800
         Left            =   200
         Picture         =   "Estorno.frx":23EA
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   250
         Width           =   850
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
         Height          =   800
         Left            =   1050
         Picture         =   "Estorno.frx":26F4
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   250
         Width           =   850
      End
   End
   Begin VB.Frame FrmImagem 
      Caption         =   "Imagem"
      Height          =   4740
      Left            =   510
      TabIndex        =   0
      Top             =   450
      Width           =   10050
      Begin LeadLib.Lead Lead1 
         Height          =   4455
         Left            =   90
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   270
         Width           =   9870
         _Version        =   524288
         _ExtentX        =   17420
         _ExtentY        =   7853
         _StockProps     =   229
         BackColor       =   16777215
         BorderStyle     =   1
         ScaleHeight     =   369
         ScaleWidth      =   820
         DataField       =   ""
         BitmapDataPath  =   ""
         AnnDataPath     =   ""
      End
   End
   Begin ComctlLib.StatusBar StatusBarEstorno 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   11100
      _ExtentX        =   19579
      _ExtentY        =   614
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   5792
            MinWidth        =   1834
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   8616
            MinWidth        =   4658
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   5088
            MinWidth        =   1130
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblAjuste 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label14"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   996
      Left            =   192
      TabIndex        =   38
      Top             =   3552
      Visible         =   0   'False
      Width           =   6300
   End
   Begin VB.Label LblVrNSU 
      AutoSize        =   -1  'True
      Caption         =   "NSU"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   5190
      TabIndex        =   37
      Top             =   5200
      Width           =   435
   End
   Begin VB.Label LblNSU 
      AutoSize        =   -1  'True
      Caption         =   "NSU:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   4590
      TabIndex        =   36
      Top             =   5200
      Width           =   555
   End
   Begin VB.Label LblValorDocto 
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
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   3255
      TabIndex        =   35
      Top             =   5200
      Width           =   990
   End
   Begin VB.Label LblVrTerminal 
      Caption         =   "Caixa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   1650
      TabIndex        =   34
      Top             =   5200
      Width           =   510
   End
   Begin VB.Label LblTerminal 
      AutoSize        =   -1  'True
      Caption         =   "Terminal:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   600
      TabIndex        =   33
      Top             =   5200
      Width           =   990
   End
   Begin VB.Label LblValor 
      AutoSize        =   -1  'True
      Caption         =   "Valor:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   2600
      TabIndex        =   32
      Top             =   5200
      Width           =   630
   End
   Begin VB.Label LblEstornado 
      AutoSize        =   -1  'True
      Caption         =   "Documento Estornado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   600
      TabIndex        =   31
      Top             =   5505
      Visible         =   0   'False
      Width           =   2310
   End
End
Attribute VB_Name = "Estorno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private teclou                                  As Boolean
Private sTempo                                  As Integer
Private sTela                                   As Boolean

'Variaveis com os dados da Capa
Private rIdLote                                 As Long
Private rIdCapa                                 As Long

'Variaveis com dados dos Doctos
Private aDoc()                                  As TpDocumento
Private IndexDoc                                As Integer

'Constantes
Private Const STATUS_TXT_DOCUMENTO = "Documento :"
Private Const STATUS_TXT_VINCULO = "Vínculo :"
Private Const STATUS_TXT_LOTECAPA = "Lote/Capa :"
Private Const STATUS_TXT_DATA = "Data :"
Private Const HABILITADO = &H80000005           'Window Background
Private Const DESABILITADO = &H8000000F         'Button Face

'Parâmetros vindos da tela de CSP
Public m_lngIdCapaCSP As Long                   'IdCapa para carga dos documentos para Estorno
Public m_strNumCapa  As String                  'Número da Capa de Malote/Envelope
Public m_bHouveEstornoCSP As Boolean            'Variável de retorno informando CSP se houve Estorno
Public m_IndexDoc As Integer                    'Posição atual do documento selecionado na tela de CSP

Private Sub CmdPesquisa_Click()
    Dim RsDoctos    As rdoResultset
    Dim Posicao     As Integer
    Dim qryGetCapaDocumentoEstorno              As RDO.rdoQuery
    
    On Error GoTo ErroPesquisa
    
   'Procura capa, Marca capa para estorno e retorna documentos.
    Set qryGetCapaDocumentoEstorno = Geral.Banco.CreateQuery("", "{? = call GetCapaDocumentoEstorno (?,?,?,?,?)}")
    
    With qryGetCapaDocumentoEstorno
            .rdoParameters(0).Direction = rdParamReturnValue
            .rdoParameters(1) = Geral.DataProcessamento                     'Data de Processamento
            .rdoParameters(2) = IIf(OpPesqCapa, TxtNumCapa.Text, Null)      'Número da Capa Malote/envelope
            .rdoParameters(3) = IIf(OpPesqMalote, TxtNumMalote, Null)       'Número do Malote
            .rdoParameters(4) = IIf(OpPesqNsu, Val(TxtNumNSU), Null)        'NSU
            .rdoParameters(5) = IIf(OpPesqNsu, TxtNumTerminal, Null)        'Terminal
            
            Set RsDoctos = qryGetCapaDocumentoEstorno.OpenResultset(rdOpenStatic, rdConcurReadOnly)
                
            If .rdoParameters(0) = 1 Then
                'Não encontrou nenhum docto com o argumento digitado
                MsgBox "Número de Documento Solicitado não Existe!.", vbCritical
                SetaText False
                Exit Sub
            ElseIf .rdoParameters(0) = 2 Then
               'Capa disponivel para estorno, porem ocorreu erro na Proc. de atualizacao do status
                MsgBox "Não foi possível atualizar Status da Capa", vbCritical
                Exit Sub
            ElseIf .rdoParameters(0) = 3 Then
               'Capa com status <> de 'E' ou 'T'
                MsgBox "Capa Encontrada não Esta disponível para estorno", vbCritical
                Exit Sub
            ElseIf .rdoParameters(0) = 5 Then
               'Mais de um malote com número solicitado.
                MsgBox "Foi encontrado mais de um Malote com este Número, Refaça a Pesquisa pela Capa.", vbCritical
                Exit Sub
            End If
    End With
       
    Tela True
        
    IndexDoc = 0
    If Not RsDoctos.EOF Then
        'Redimensiona o vetor de documentos
         Erase aDoc
         ReDim aDoc(RsDoctos.RowCount - 1)
         rIdCapa = RsDoctos!IdCapa
         rIdLote = RsDoctos!IdLote
         Posicao = RsDoctos!Posicao - 1
         
        'Carrega vetor de documentos marcados para correção de Agencia e Conta'
         While Not RsDoctos.EOF
            'Carrega dados dos documentos
             With RsDoctos
                aDoc(IndexDoc).IdCapa = !IdCapa
                aDoc(IndexDoc).IdDocto = !IdDocto
                aDoc(IndexDoc).Frente = !Frente
                aDoc(IndexDoc).Verso = !Verso
                aDoc(IndexDoc).TipoDocto = !TipoDocto
                aDoc(IndexDoc).Status = !Status
                aDoc(IndexDoc).Vinculo = !Vinculo
                aDoc(IndexDoc).Ordem = !Ordem
                aDoc(IndexDoc).ValorTotal = !Valor
                aDoc(IndexDoc).NSU = IIf(IsNull(!NSU), 0, !NSU)
                aDoc(IndexDoc).Terminal = IIf(IsNull(!Terminal), 0, !Terminal)
                aDoc(IndexDoc).Estornado = IIf(IsNull(!Estornado), "0", !Estornado)
             End With
             
             IndexDoc = IndexDoc + 1
             RsDoctos.MoveNext
         Wend
    End If
    
   'Acao: Estorno seleciona Capa
    Call GravaLog(rIdCapa, 0, 231)
    
    IndexDoc = IIf(OpPesqNsu, Posicao, 1)
    Call MostraImagem
    Exit Sub
    
ErroPesquisa:
    Call TratamentoErro("Erro ao pequisar Documento para Estorno", Err, rdoErrors)
    SetaText True

End Sub
Private Sub cmdConfirma_Click()
    Dim qryInsereDoctoEstorno   As RDO.rdoQuery
    Dim i                       As Integer
    Dim EnvMsg, Marcou          As Boolean
    Dim DcExcluir(3)            As Long       'Posicao(0) = Opção do estorno , Posição(1)= Documento/Vinculo/Capa, Posição(2)= iDocto, Posição(3)= TipoDocto
    
    On Error GoTo ErroConfirma
    
   'Evitar de enviar mais de 1 msg em bloco de doctos, durante o Loop
    EnvMsg = False
    
    If OpEstornaCapa Then
        If (MsgBox("Confirma o Estorno de Todos os Documentos pertencentes a Capa ?", vbExclamation + vbYesNo + vbDefaultButton2, App.Title)) = vbNo Then
            Exit Sub
        Else
            DcExcluir(0) = 1
            DcExcluir(1) = aDoc(IndexDoc).IdCapa
            DcExcluir(3) = aDoc(IndexDoc).TipoDocto
        End If
    ElseIf OpEstornaDocumento Then
        DcExcluir(0) = 2
        DcExcluir(2) = aDoc(IndexDoc).IdDocto
        DcExcluir(1) = aDoc(IndexDoc).Vinculo
        DcExcluir(3) = aDoc(IndexDoc).TipoDocto
    
        If (aDoc(IndexDoc).TipoDocto = 2 Or aDoc(IndexDoc).TipoDocto = 3) Then
            If (MsgBox("Capa Depósito!, Serão Estornados Todos os Doctos desse Vínculo ?", vbQuestion + vbYesNo, App.Title)) = vbNo Then
                Exit Sub
            Else
                DcExcluir(0) = 3
            End If
        ElseIf (aDoc(IndexDoc).TipoDocto = 37 Or aDoc(IndexDoc).TipoDocto = 39) Then
            If (MsgBox("OCT!, Serão Estornados Todos os Doctos desse Vínculo ?", vbQuestion + vbYesNo, App.Title)) = vbNo Then
                Exit Sub
            Else
                DcExcluir(0) = 3
            End If
        Else
            If (MsgBox("Confirma o Estorno do Documento Corrente ?", vbQuestion + vbYesNo, App.Title)) = vbNo Then
                Exit Sub
            End If
        End If
    ElseIf OpEstornaVinculo Then
        If (MsgBox("Confirma o Estorno de Todos os Documentos deste Vinculo ?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title)) = vbNo Then
            Exit Sub
        Else
            DcExcluir(0) = 3
            DcExcluir(1) = aDoc(IndexDoc).Vinculo
            DcExcluir(3) = aDoc(IndexDoc).TipoDocto
        End If
    End If
    
    Geral.Banco.BeginTrans
    
   'Insere documentos na tabela de estorno
    Set qryInsereDoctoEstorno = Geral.Banco.CreateQuery("", "{? = call InsereDocumentoEstorno (?,?,?,?,?,?)}")

    i = 0
    Do
       If (IIf(DcExcluir(0) = 1, aDoc(i).IdCapa, aDoc(i).Vinculo) = DcExcluir(1) And aDoc(i).TipoDocto <> 1) Then
            If ((DcExcluir(0) = 2 And aDoc(i).TipoDocto = 41 And aDoc(i).IdDocto = DcExcluir(2)) Or (DcExcluir(0) <> 2 And aDoc(i).TipoDocto = 41)) Then
                Geral.Banco.RollbackTrans
                MsgBox "Não é possível efetuar o Estorno!, há {Lançamento Interno} entre o(s) documento(s) Selecionados.", vbCritical
                
                'Seja havia marcado doctos volta a tela de pesquisa
                If Marcou Then
                    'Volta o status anterior da capa, se estornada a capa inteira
                    If Not RetornaStatus Then
                        Exit Sub
                    End If

                   'volta para pesquisa
                    SetaText True
                End If
                Exit Sub
            End If
            
            If ((DcExcluir(0) = 2 And (InStr("T*E", aDoc(i).Status) = 0) And aDoc(i).IdDocto = DcExcluir(2)) Or (DcExcluir(0) <> 2 And (InStr("T*E", aDoc(i).Status) = 0))) Then

                Geral.Banco.RollbackTrans
                If OpEstornaDocumento.Value Then
                    MsgBox "Não é possível efetuar o Estorno, documento não foi transmitido.", vbCritical
                Else
                    MsgBox "Não é possível efetuar o Estorno, existe documento não transmitido.", vbCritical
                End If
                
                'Seja havia marcado doctos volta a tela de pesquisa
                If Marcou Then
                    'Volta o status anterior da capa, se estornada a capa inteira
                    If Not RetornaStatus Then
                        Exit Sub
                    End If

                   'volta para pesquisa
                    SetaText True
                End If
                Exit Sub
            End If
            
           'Verifica se no documento(s) selecionado(s) há número(s) de Nsu zerado
            If (Val(aDoc(i).NSU) = 0 And aDoc(i).TipoDocto <> 39 And DcExcluir(0) <> 2) Or (Val(aDoc(i).NSU) = 0 And DcExcluir(0) = 2 And (DcExcluir(2) = aDoc(i).IdDocto And aDoc(i).TipoDocto <> 39)) Then
                If (aDoc(i).Estornado = "0") Then
                    Geral.Banco.RollbackTrans
                    MsgBox "Não é possível efetuar o Estorno" & vbCrLf & "Número de NSU do(s) documento(s) selecionado(s) - Inválido(s). ", vbCritical
                   'Seja havia marcado doctos volta a tela de pesquisa
                    If Marcou Then
                        'Volta o status anterior da capa, se estornada a capa inteira
                         If Not RetornaStatus Then
                            Exit Sub
                         End If
                          
                        'volta para pesquisa
                         SetaText True
                    End If
                    
                    Exit Sub
                End If
            End If
            
           'Verifica se no documento(s) selecionado(s) existe documento não transmitido/expedido
            If ((InStr("T*E", aDoc(i).Status) = 0) And aDoc(i).TipoDocto <> 39 And DcExcluir(0) <> 2) Or ((InStr("T*E", aDoc(i).Status) = 0) And DcExcluir(0) = 2 And (DcExcluir(2) = aDoc(i).IdDocto And aDoc(i).TipoDocto <> 39)) Then
                If (aDoc(i).Estornado = "0") Then
                    Geral.Banco.RollbackTrans
                    If OpEstornaDocumento.Value Then
                        MsgBox "Não é possível efetuar o Estorno, documento não foi transmitido.", vbCritical
                    Else
                        MsgBox "Não é possível efetuar o Estorno, existe documento não transmitido.", vbCritical
                    End If
                   'Seja havia marcado doctos volta a tela de pesquisa
                    If Marcou Then
                        'Volta o status anterior da capa, se estornada a capa inteira
                         If Not RetornaStatus Then
                            Exit Sub
                         End If
                          
                        'volta para pesquisa
                         SetaText True
                    End If
                    
                    Exit Sub
                End If
            End If
            
            If (DcExcluir(0) = 2 And (aDoc(i).TipoDocto = 2 Or aDoc(i).TipoDocto = 3 Or aDoc(i).TipoDocto = 37)) Then
                Geral.Banco.RollbackTrans
                MsgBox "Não é possível efetuar o Estorno, documento(s) c/ Capa de Depósito ou OCT, deve-se excluir o Vínculo. ", vbCritical
                               
               'Seja havia marcado doctos volta a tela de pesquisa
                If Marcou Then
                    'Volta o status anterior da capa, se estornada a capa inteira
                    If Not RetornaStatus Then
                        Exit Sub
                    End If
                  
                   'volta para pesquisa
                    SetaText True
                End If
                
                Exit Sub
            End If
                        
            If (DcExcluir(0) = 1 Or DcExcluir(0) = 3 Or (DcExcluir(0) = 2 And aDoc(i).IdDocto = DcExcluir(2))) Then
                
                With qryInsereDoctoEstorno
                     .rdoParameters(0).Direction = rdParamReturnValue
                     .rdoParameters(1) = Geral.DataProcessamento                'Data de Processamento
                     .rdoParameters(2) = aDoc(i).IdCapa                         'IdCapa
                     .rdoParameters(3) = aDoc(i).IdDocto                        'IdDocto
                     .rdoParameters(4) = aDoc(i).NSU                            'NSU
                     .rdoParameters(5) = IIf(aDoc(i).TipoDocto = 7 Or _
                                             aDoc(i).TipoDocto = 32 Or _
                                             aDoc(i).TipoDocto = 33, "C", "N")  'Status
                     .rdoParameters(6) = Geral.Usuario                          'Usuario
                     .Execute
                     
                    'Retorno 1 ocorreu algum erro na Procedure de Estorno de documentos
                     If .rdoParameters(0) = 1 Then
                         Geral.Banco.RollbackTrans
                         GoTo ErroConfirma:
                         
                    'Retorno 2 docto marcado p/ estorno, Retorno 3, docto já estornado pelo robo
                     ElseIf .rdoParameters(0) = 2 Or .rdoParameters(0) = 3 And Not EnvMsg Then
                         If (MsgBox("Um ou mais documento(s) já estão marcados para Estorno Continua ?", vbQuestion + vbYesNo, App.Title)) = vbNo Then
                             Geral.Banco.RollbackTrans
                             Exit Sub
                         Else
                            'já enviou uma msg no Loop
                             EnvMsg = True
                         End If
                     Else
                        'Atualiza estorno de documento no vetor
                         aDoc(i).Estornado = IIf(aDoc(i).TipoDocto = 7 Or aDoc(i).TipoDocto = 32 Or aDoc(i).TipoDocto = 33, "C", "N")
                         Marcou = True
                     End If
                End With
            End If
        End If
        i = i + 1
    Loop Until i > UBound(aDoc)
   
    Geral.Banco.CommitTrans
    
    'Controle de Estorno do módulo de CSP
    m_bHouveEstornoCSP = True
    
    Call GravaLog(aDoc(IndexDoc).IdCapa, aDoc(IndexDoc).IdDocto, 232)
    MsgBox "Documento(s) Selecionado(s) Marcados para Estorno.", vbInformation + vbOKOnly, App.Title
    
    If DcExcluir(0) = 2 Or DcExcluir(0) = 3 Then
       MostraImagem
    Else
       'Volta o status anterior da capa, se estornada a capa inteira
        If Not RetornaStatus Then
           Exit Sub
        End If
        
       'volta para pesquisa
        SetaText True
    End If
    
    Exit Sub
    
ErroConfirma:
    If Marcou Then Geral.Banco.RollbackTrans
    Call TratamentoErro("Houve Falha ao enviar documento(s) para estorno", Err, rdoErrors)
    Call RetornaStatus
    Unload Me
End Sub
Private Sub cmdDoctoAnterior_Click()

    If IndexDoc = 1 Then
       Beep
    Else
        IndexDoc = IndexDoc - 1
    End If
    
    Call MostraImagem
    
End Sub
Private Sub cmdDoctoPosterior_Click()

    If IndexDoc = UBound(aDoc) Then
        Beep
    Else
        IndexDoc = IndexDoc + 1
    End If
    
    Call MostraImagem
    
End Sub
Private Sub cmdFinalizar_Click()
   
    'Volta o status anterior da capa
    
     If Not RetornaStatus Then
         Exit Sub
     End If
     
     Unload Me
     Exit Sub
     
End Sub
Private Sub CmdSair_Click()
    Unload Me
End Sub
Private Sub cmdZoomMais_Click()

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
Private Sub cmdZoomMenos_Click()

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
Private Sub cmdFrenteVerso_Click()

  On Error GoTo ERRO_FRENTEVERSO

  If teclou Then Exit Sub

  If FrmImagem.Visible = False Then Exit Sub

  teclou = True
  'só muda para frente/verso qdo docto vem da Ls500 e da Vips,'pois, da canon não gera verso.
  
  If (aDoc(IndexDoc).Ordem = "0") Or (aDoc(IndexDoc).Ordem = "2") Then
    If Lead1.Tag = "V" Then
        Lead1.Tag = "F"     'se verso, mostrar frente
        With Lead1
            .AutoRepaint = False
            If Geral.VIPSDLL = eDllProservi Then
              .Load Geral.DiretorioImagens & aDoc(IndexDoc).Frente, 0, 0, 1
            Else
              .Load Geral.DiretorioImagens & Format(rIdLote, "000000000") & "\" & aDoc(IndexDoc).Frente, 0, 0, 1
            End If
  
            'se ls500 mostrar mais escuro
            If (aDoc(m_IndexDoc).Ordem = "2") Then
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
              .Load Geral.DiretorioImagens & Trim(aDoc(IndexDoc).Verso), 0, 0, 1
            Else
              .Load Geral.DiretorioImagens & Format(rIdLote, "000000000") & "\" & aDoc(IndexDoc).Verso, 0, 0, 1
            End If
  
            'se ls500 mostrar mais escuro
            If (aDoc(IndexDoc).Ordem = "2") Then
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
Private Sub cmdInverteCor_Click()
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
Private Sub cmdRotacao_Click()

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
Private Sub Form_Activate()
    rIdCapa = 0
    Tela False
    sTela = False
    LimpaPesquisa 1
    
    With Lead1
        .UnlockSupport L_SUPPORT_EXPRESS, "YXPQ3XPPVT"
        .UnlockSupport L_SUPPORT_GIFLZW, "0K3RV9UY3EY"
        .UnlockSupport L_SUPPORT_TIFLZW, "9LE75L0FDXHK"
    End With
    
    m_bHouveEstornoCSP = False
    'Verifica se Acionada Tela de Estorno através de CSP
    If m_lngIdCapaCSP <> 0 Then
        TxtNumCapa = m_strNumCapa
        If Not PesquisaCapaEstornoCSP Then
            'Limpa parametros do módulo de CSP
            Call LimpaDadosCSP
            Unload Me
        End If
    End If
    
End Sub
Public Sub LimpaPesquisa(iObj As Integer)
Select Case iObj
       Case 1
            TxtNumCapa.BackColor = HABILITADO
            TxtNumCapa.Enabled = True
            TxtNumCapa.Text = ""
            
            TxtNumMalote.BackColor = DESABILITADO
            TxtNumMalote.Text = ""
            TxtNumMalote.Enabled = False
            
            TxtNumNSU.Text = ""
            TxtNumNSU.BackColor = DESABILITADO
            TxtNumNSU.Enabled = False
            TxtNumTerminal.Text = ""
            TxtNumTerminal.BackColor = DESABILITADO
            TxtNumTerminal.Enabled = False
            
            TxtNumCapa.SetFocus
            
       Case 2
            TxtNumCapa.BackColor = DESABILITADO
            TxtNumCapa.Text = ""
            TxtNumCapa.Enabled = False
            
            TxtNumMalote.BackColor = HABILITADO
            TxtNumMalote.Enabled = True
            TxtNumMalote.Text = ""
            
            TxtNumNSU.BackColor = DESABILITADO
            TxtNumNSU.Text = ""
            TxtNumNSU.Enabled = False
            TxtNumTerminal.BackColor = DESABILITADO
            TxtNumTerminal.Text = ""
            TxtNumTerminal.Enabled = False
            
            TxtNumMalote.SetFocus
       Case 3
            TxtNumCapa.BackColor = DESABILITADO
            TxtNumCapa.Text = ""
            TxtNumCapa.Enabled = False
            
            TxtNumMalote.BackColor = DESABILITADO
            TxtNumMalote.Text = ""
            TxtNumMalote.Enabled = False
            
            TxtNumNSU.BackColor = HABILITADO
            TxtNumNSU.Enabled = True
            TxtNumNSU.Text = ""
            TxtNumTerminal.BackColor = HABILITADO
            TxtNumTerminal.Enabled = True
            TxtNumTerminal.Text = ""
            
            TxtNumNSU.SetFocus
       End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
 Dim Ret As Long

  hCtl = Lead1.hwnd

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
Private Sub Form_Unload(Cancel As Integer)

    'Limpa parametros do módulo de CSP
    Call LimpaDadosCSP
    
End Sub
Private Sub OpPesqCapa_Click()
    LimpaPesquisa 1
End Sub
Private Sub OpPesqMalote_Click()
    LimpaPesquisa 2
End Sub
Private Sub OpPesqNsu_Click()
    LimpaPesquisa 3
End Sub
Private Sub TxtNumTerminal_Change()
    If ValidaPesquisa(TxtNumTerminal.Text, TxtNumTerminal) Then
        If ValidaPesquisa(TxtNumNSU.Text, TxtNumNSU) Then
            CmdPesquisa.Enabled = True
        Else
            CmdPesquisa.Enabled = False
        End If
    Else
        CmdPesquisa.Enabled = False
    End If
End Sub
Private Sub TxtNumCapa_Change()
    
    If ValidaPesquisa(TxtNumCapa.Text, TxtNumCapa) Then
        CmdPesquisa.Enabled = True
    Else
        CmdPesquisa.Enabled = False
    End If
   
End Sub
Private Sub TxtNumCapa_GotFocus()
    TxtNumCapa.SelStart = 0
    TxtNumCapa.SelLength = Len(TxtNumCapa.Text)
End Sub
Private Sub TxtNumCapa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And CmdPesquisa.Enabled Then
        CmdPesquisa.SetFocus
        CmdPesquisa_Click
    End If
End Sub
Private Sub txtNumMalote_Change()
    If ValidaPesquisa(TxtNumMalote.Text, TxtNumMalote) Then
        CmdPesquisa.Enabled = True
    Else
        CmdPesquisa.Enabled = False
    End If
End Sub
Private Sub TxtNumMalote_GotFocus()
    SelecionarTexto TxtNumMalote
End Sub
Private Sub TxtNumMalote_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And CmdPesquisa.Enabled Then
        CmdPesquisa.SetFocus
        CmdPesquisa_Click
    End If
End Sub
Private Sub TxtNumNSU_Change()

    If ValidaPesquisa(TxtNumNSU.Text, TxtNumNSU) Then
        If ValidaPesquisa(TxtNumTerminal.Text, TxtNumTerminal) Then
            CmdPesquisa.Enabled = True
        Else
            CmdPesquisa.Enabled = False
        End If
    Else
        CmdPesquisa.Enabled = False
    End If

   If Len(TxtNumNSU.Text) = TxtNumNSU.MaxLength Then TxtNumTerminal.SetFocus

End Sub
Private Sub MostraImagem()

    On Error GoTo ERRO_MOSTRAIMAGEM
    
    Dim Ret As Long
    Cabecalho True
    hCtl = Lead1.hwnd
    
   'em teste nao da erro se nao existe imagem
   'If UCase(Geral.Usuario) = "DESENV" And Dir(Trim(Geral.DiretorioImagens) & aDoc(IndexDoc).Frente, vbArchive) = "" Then GoTo NoImage
   
    If Not (aDoc(IndexDoc).TipoDocto = 32 Or aDoc(IndexDoc).TipoDocto = 33 Or _
            aDoc(IndexDoc).TipoDocto = 34 Or aDoc(IndexDoc).TipoDocto = 38 Or _
            aDoc(IndexDoc).TipoDocto = 42 Or aDoc(IndexDoc).TipoDocto = 43 Or _
            aDoc(IndexDoc).TipoDocto = 44 Or aDoc(IndexDoc).TipoDocto = 45) Then
        
        If lblAjuste.Visible Then lblAjuste.Visible = False
        
        'Coloca imagem na tela
        With Lead1
            .Tag = "F"
            .AutoRepaint = False
            If Geral.VIPSDLL = eDllProservi Then
                .Load Geral.DiretorioImagens & aDoc(IndexDoc).Frente, 0, 0, 1
            Else
                .Load Geral.DiretorioImagens & Format(rIdLote, "000000000") & "\" & aDoc(IndexDoc).Frente, 0, 0, 1
            End If
            
           'Se imagem for da ls500, deixar mais escura
            If aDoc(IndexDoc).Ordem <> "2" Then
                .Intensity 220
            Else
                .Intensity 140
            End If
            
           'Se imagem for do canon, diminui em 50% o tamanho
            If aDoc(IndexDoc).Ordem <> "1" Then
                .PaintZoomFactor = 100
            Else
                .PaintZoomFactor = 50
            End If
            .AutoRepaint = True
        End With
        FrmImagem.Visible = True
    Else
        If Not lblAjuste.Visible Then lblAjuste.Visible = True
        FrmImagem.Visible = False
        ApresentaTelaAjuste (aDoc(IndexDoc).TipoDocto)
    End If

   'Posiciona imagem sempre no começo
    Ret = SendMessage(hCtl, WM_VSCROLL, SB_TOP, 0)
    Ret = SendMessage(hCtl, WM_HSCROLL, SB_TOP, 0)
    
   'Habilita Objetos de Manipulação de Imagens
    If Not (aDoc(IndexDoc).TipoDocto = 32 Or aDoc(IndexDoc).TipoDocto = 33 Or _
            aDoc(IndexDoc).TipoDocto = 34 Or aDoc(IndexDoc).TipoDocto = 38 Or _
            aDoc(IndexDoc).TipoDocto = 42 Or aDoc(IndexDoc).TipoDocto = 43 Or _
            aDoc(IndexDoc).TipoDocto = 44 Or aDoc(IndexDoc).TipoDocto = 45) Then
        Call HDObjetosImagem(True)
    Else
        Call HDObjetosImagem(False)
    End If
NoImage:
   'Carrega Dados do Documento em tela
    LblValorDocto = Format(aDoc(IndexDoc).ValorTotal, "###,###,###,##0.00")
    LblVrTerminal = aDoc(IndexDoc).Terminal
    LblVrNSU = aDoc(IndexDoc).NSU
    
    If aDoc(IndexDoc).Estornado = "S" Then
        LblEstornado.ForeColor = &HC0&
        LblEstornado.Caption = "Documento Estornado"
        LblEstornado.Visible = True
        cmdConfirma.Enabled = False
    ElseIf aDoc(IndexDoc).Estornado = "N" Or aDoc(IndexDoc).Estornado = "P" Then
        LblEstornado.ForeColor = &HC00000
        LblEstornado.Caption = "Documento para Estorno"
        LblEstornado.Visible = True
        cmdConfirma.Enabled = False
     ElseIf aDoc(IndexDoc).Estornado = "F" Then
        LblEstornado.ForeColor = &H808080
        LblEstornado.Caption = "Falha no Processo de Estorno"
        LblEstornado.Visible = True
        cmdConfirma.Enabled = False
     ElseIf aDoc(IndexDoc).Estornado = "C" Then
        LblEstornado.ForeColor = &H808080
        LblEstornado.Caption = "'Documento Vinculo a Capa Deposito/OCT"
        LblEstornado.Visible = True
        cmdConfirma.Enabled = False
    Else
        LblEstornado.Visible = False
        cmdConfirma.Enabled = True
    End If
    
    DoEvents
    
    Exit Sub

ERRO_MOSTRAIMAGEM:
    Screen.MousePointer = vbDefault
    MsgBox "Não foi possível exibir a Imagem do Documento, imagem não encontrada.", vbInformation, App.Title
    Call HDObjetosImagem(False)
End Sub
Sub HDObjetosImagem(bValor As Boolean)

    On Error GoTo ERRO_HDOBJETOS
    
    cmdZoomMais.Enabled = bValor
    cmdZoomMenos.Enabled = bValor
    cmdRotacao.Enabled = bValor
    cmdInverteCor.Enabled = bValor
    cmdFrenteVerso.Enabled = bValor
    cmdConfirma.Enabled = bValor
    FrmImagem.Visible = bValor
    Lead1.ForceRepaint
    
    Exit Sub

ERRO_HDOBJETOS:
    Screen.MousePointer = vbDefault
    Call TratamentoErro("Erro ao preparar botões de manipulação de Imagens.", Err, rdoErrors)
End Sub
Private Sub Cabecalho(ByVal pShow As Boolean)

    Dim sDocumento          As String
    Dim sLoteCapa           As String
    Dim sVinculo            As String

    
    sDocumento = STATUS_TXT_DOCUMENTO & " "
    sLoteCapa = STATUS_TXT_LOTECAPA & " "
    sVinculo = STATUS_TXT_VINCULO & " "
    
    If pShow Then

        sDocumento = sDocumento & IndexDoc + 1 & "/" & UBound(aDoc) + 1
        sVinculo = sVinculo & aDoc(IndexDoc).Vinculo
        sLoteCapa = sLoteCapa & Format(rIdLote, "0000-00000") & " / " & Trim(rIdCapa)

    End If
    
    StatusBarEstorno.Panels(1).Text = sDocumento & " - " & sVinculo
    StatusBarEstorno.Panels(2).Text = sLoteCapa
    StatusBarEstorno.Panels(3).Text = STATUS_TXT_DATA & Format(Format(Geral.DataProcessamento, "0000/00/00"), "dd/mm/yyyy")

End Sub
Sub Tela(Exibe As Boolean)
    If Exibe Then
        cmdFinalizar.Cancel = True
        tmrAtualiza.Enabled = True
        FramePesquisaEstorno.Visible = False
        Estorno.Height = 7800
        sTela = True
    Else
        cmdSair.Cancel = True
        tmrAtualiza.Enabled = False
        FramePesquisaEstorno.Visible = True
        Estorno.Height = 2750
        sTela = True
    End If
End Sub
Private Sub TxtNumNSU_GotFocus()
    SelecionarTexto TxtNumNSU
End Sub
Private Sub TxtNumNSU_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then TxtNumTerminal.SetFocus
End Sub
Private Sub tmrAtualiza_Timer()
    tmrAtualiza.Enabled = False
    
    If rIdCapa <> 0 Then
        sTempo = sTempo + Int(tmrAtualiza.Interval / 1000)
        If sTempo + Int(tmrAtualiza.Interval / 1000) >= Geral.Intervalo Then
        
            'Atualiza Tempo da capa corrente
            If Not Globais.AtualizaStatusCapa(rIdCapa, "W") Then
                Exit Sub
            End If
        
            sTempo = 0
        End If
    End If
    
    tmrAtualiza.Enabled = True

End Sub
Sub SetaText(Limpa As Boolean)

    'Verifica se Acionada Tela de Estorno através de CSP
    If m_lngIdCapaCSP <> 0 Then Exit Sub

    If OpPesqCapa Then
        If Limpa Then
           Tela False
           LimpaPesquisa 1
        End If
        TxtNumCapa.SetFocus

    ElseIf OpPesqMalote Then
        If Limpa Then
            Tela False
            LimpaPesquisa 2
        End If
        TxtNumMalote.SetFocus

    Else
        If Limpa Then
            Tela False
            LimpaPesquisa 3
        End If
        TxtNumNSU.SetFocus

    End If
End Sub
Private Sub TxtNumTerminal_GotFocus()
    SelecionarTexto TxtNumTerminal
End Sub
Private Sub TxtNumTerminal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And CmdPesquisa.Enabled Then
        CmdPesquisa.SetFocus
        CmdPesquisa_Click
    End If
End Sub
Function RetornaStatus() As Boolean
   'Atualiza Status da Capa e limpa linha criada na Tab estorno p/ guardar Status
    Dim qryAtualizaStatusCapaEstorno As RDO.rdoQuery
    Dim AnteriorStatus As String * 1
    On Error GoTo errostatus:
    
    Set qryAtualizaStatusCapaEstorno = Geral.Banco.CreateQuery("", "{? = call AtualizaStatusCapaEstorno (?,?)}")
        
    With qryAtualizaStatusCapaEstorno
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1) = Geral.DataProcessamento   'Data de Processamento
        .rdoParameters(2) = aDoc(i).IdCapa            'IdCapa
        .Execute
                     
        If .rdoParameters(0).Value = 0 Then
            RetornaStatus = True
        Else
            GoTo errostatus:
        End If
    End With
    
    Exit Function
    
errostatus:
    Call TratamentoErro("Falha ao Atualizar Status da Capa", Err, rdoErrors)
    RetornaStatus = False
    
End Function
Function ValidaPesquisa(pValor As String, Texto As TextBox)
    
    If Len(Trim(pValor)) = 0 Then
        Exit Function
    ElseIf IsNumeric(pValor) = False Or InStr(1, pValor, "-", vbTextCompare) > 0 Or InStr(1, pValor, "+", vbTextCompare) > 0 Then
        MsgBox "Valor Inválido para este Campo.", vbExclamation + vbOKOnly, App.Title
        Texto.Text = ""
        Texto.SetFocus
    Else
        ValidaPesquisa = True
    End If
    
End Function
Private Function PesquisaCapaEstornoCSP() As Boolean
    
Dim RsDoctos                    As rdoResultset
Dim Posicao                     As Integer
Dim qryGetDocumentoEstorno  As New rdoQuery

On Error GoTo Err_PesquisaCapaEstornoCSP
    
    PesquisaCapaEstornoCSP = False
    
   'Procura capa, Marca capa para estorno e retorna documentos.
    Set qryGetDocumentoEstorno = Geral.Banco.CreateQuery("", "{call GetDocumentosEstornoCSP (?,?)}")
    
    With qryGetDocumentoEstorno
        .rdoParameters(0) = Geral.DataProcessamento     'Data de Processamento
        .rdoParameters(1) = m_lngIdCapaCSP              'IdCapa
        
        Set RsDoctos = .OpenResultset(rdOpenStatic, rdConcurReadOnly)
                
    End With
       
    If RsDoctos.EOF Then
        MsgBox "Não foi possível obter documentos para estorno. Tente novamente!", vbInformation + vbOKOnly, App.Title
        GoTo Exit_PesquisaCapaEstornoCSP
    End If
    
    Tela True
        
    IndexDoc = 0
    If Not RsDoctos.EOF Then
        'Redimensiona o vetor de documentos
         Erase aDoc
         ReDim aDoc(RsDoctos.RowCount - 1)
         rIdCapa = RsDoctos!IdCapa
         rIdLote = RsDoctos!IdLote
         Posicao = RsDoctos!Posicao - 1
         
        'Carrega vetor de documentos marcados para correção de Agencia e Conta'
         While Not RsDoctos.EOF
            'Não carregar documentos do Tipo Ajuste
             With RsDoctos
                aDoc(IndexDoc).IdCapa = !IdCapa
                aDoc(IndexDoc).IdDocto = !IdDocto
                aDoc(IndexDoc).Frente = !Frente
                aDoc(IndexDoc).Verso = !Verso
                aDoc(IndexDoc).TipoDocto = !TipoDocto
                aDoc(IndexDoc).Status = !Status
                aDoc(IndexDoc).Vinculo = !Vinculo
                aDoc(IndexDoc).Ordem = !Ordem
                aDoc(IndexDoc).ValorTotal = !Valor
                aDoc(IndexDoc).NSU = IIf(IsNull(!NSU), 0, !NSU)
                aDoc(IndexDoc).Terminal = IIf(IsNull(!Terminal), 0, !Terminal)
                aDoc(IndexDoc).Estornado = IIf(IsNull(!Estornado), "0", !Estornado)
            End With
             IndexDoc = IndexDoc + 1
             RsDoctos.MoveNext
         Wend
    End If
    
   'Acao: Estorno seleciona Capa
    Call GravaLog(rIdCapa, 0, 231)
    
    IndexDoc = IIf(OpPesqNsu, Posicao, m_IndexDoc)
    Call MostraImagem
    
    PesquisaCapaEstornoCSP = True
    
Exit_PesquisaCapaEstornoCSP:
    Set qryGetDocumentoEstorno = Nothing
    Exit Function
    
Err_PesquisaCapaEstornoCSP:
    Call TratamentoErro("Erro ao pequisar Documento para Estorno", Err, rdoErrors)
    SetaText True
    GoTo Exit_PesquisaCapaEstornoCSP

End Function
Private Sub LimpaDadosCSP()
            
    m_lngIdCapaCSP = 0
    m_strNumCapa = ""

End Sub
Private Sub ApresentaTelaAjuste(ByVal intTipoDocto As Integer)

    'Centraliza Label de Ajuste
    lblAjuste.Top = FrmImagem.Top + (FrmImagem.Height - lblAjuste.Height) / 2
    lblAjuste.Left = FrmImagem.Left + (FrmImagem.Width - lblAjuste.Width) / 2
    
     'Habilita Objetos de Manipulação de Imagens
    If intTipoDocto = 32 Or intTipoDocto = 34 Or intTipoDocto = 42 Or intTipoDocto = 44 Then
        lblAjuste.Caption = "Ajuste Crédito"
    ElseIf intTipoDocto = 33 Or intTipoDocto = 38 Or intTipoDocto = 43 Or intTipoDocto = 45 Then
        lblAjuste.Caption = "Ajuste Débito"
    End If
    
End Sub


