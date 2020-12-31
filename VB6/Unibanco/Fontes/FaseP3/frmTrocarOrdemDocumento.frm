VERSION 5.00
Object = "{00080003-B1BA-11CE-ABC6-F5B2E79D9E3F}#8.0#0"; "LTOCX80N.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmTrocarOrdemDocumento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Troca de ordem de documentos"
   ClientHeight    =   8304
   ClientLeft      =   36
   ClientTop       =   288
   ClientWidth     =   11232
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8304
   ScaleWidth      =   11232
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   120
      Left            =   3216
      Top             =   0
   End
   Begin VB.Frame Frame3 
      Height          =   4092
      Left            =   9456
      OLEDropMode     =   1  'Manual
      TabIndex        =   16
      Top             =   -48
      Width           =   1740
      Begin VB.CommandButton cmdLimpar 
         Caption         =   "&Limpar"
         Enabled         =   0   'False
         Height          =   348
         Left            =   144
         TabIndex        =   25
         Top             =   576
         Width           =   1452
      End
      Begin VB.Timer tmrAtualiza 
         Enabled         =   0   'False
         Interval        =   50000
         Left            =   1392
         Top             =   3216
      End
      Begin VB.CommandButton cmdFechar 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   348
         Left            =   144
         TabIndex        =   6
         Top             =   960
         Width           =   1452
      End
      Begin VB.CommandButton cmdConfirmar 
         Caption         =   "Confirmar"
         Enabled         =   0   'False
         Height          =   348
         Left            =   144
         TabIndex        =   5
         Top             =   192
         Width           =   1452
      End
   End
   Begin VB.Frame Frame5 
      Height          =   4176
      Left            =   9456
      TabIndex        =   15
      Top             =   4080
      Width           =   1740
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
         Height          =   696
         Left            =   528
         Picture         =   "frmTrocarOrdemDocumento.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   192
         Width           =   820
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
         Height          =   696
         Left            =   528
         Picture         =   "frmTrocarOrdemDocumento.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   960
         Width           =   820
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
         Height          =   696
         Left            =   528
         Picture         =   "frmTrocarOrdemDocumento.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1764
         Width           =   820
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
         Height          =   696
         Left            =   528
         Picture         =   "frmTrocarOrdemDocumento.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2580
         Width           =   820
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
         Height          =   696
         Left            =   528
         Picture         =   "frmTrocarOrdemDocumento.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3384
         Width           =   820
      End
   End
   Begin VB.Frame FrmImagem 
      Caption         =   "Imagem"
      Height          =   4176
      Left            =   48
      TabIndex        =   13
      Top             =   4080
      Width           =   9336
      Begin LeadLib.Lead Lead1 
         Height          =   3900
         Left            =   144
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   216
         Width           =   9120
         _Version        =   524288
         _ExtentX        =   16087
         _ExtentY        =   6879
         _StockProps     =   229
         BackColor       =   -2147483639
         BorderStyle     =   1
         ScaleHeight     =   323
         ScaleWidth      =   758
         DataField       =   ""
         BitmapDataPath  =   ""
         AnnDataPath     =   ""
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4092
      Left            =   48
      TabIndex        =   12
      Top             =   -48
      Width           =   9336
      Begin ComctlLib.ListView ListView1 
         Height          =   2892
         Left            =   4656
         TabIndex        =   4
         Top             =   1104
         Width           =   4620
         _ExtentX        =   8149
         _ExtentY        =   5101
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDragMode     =   1
         OLEDropMode     =   1
         _Version        =   327682
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         OLEDragMode     =   1
         OLEDropMode     =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Documento"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Valor do documento"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   2796
         Left            =   4704
         ScaleHeight     =   2796
         ScaleWidth      =   4524
         TabIndex        =   26
         Top             =   1152
         Width           =   4524
      End
      Begin VB.PictureBox Picture3 
         Height          =   396
         Left            =   96
         ScaleHeight     =   348
         ScaleWidth      =   1860
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   192
         Width           =   1908
         Begin VB.Label lblCapa 
            AutoSize        =   -1  'True
            Caption         =   "Capa"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   312
            Left            =   0
            TabIndex        =   24
            Top             =   0
            Width           =   1992
         End
      End
      Begin VB.PictureBox picNumMalote 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   396
         Left            =   4704
         ScaleHeight     =   348
         ScaleWidth      =   2052
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   192
         Width           =   2100
         Begin VB.Label lblMalote 
            Caption         =   "Número do Malote"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   252
            Left            =   36
            TabIndex        =   22
            Top             =   36
            Width           =   1956
         End
      End
      Begin VB.TextBox txtNumMalote 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   420
         Left            =   6864
         MaxLength       =   12
         TabIndex        =   1
         Top             =   192
         Width           =   2388
      End
      Begin VB.ComboBox cboAgencia 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   396
         Left            =   2028
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   648
         Width           =   2604
      End
      Begin VB.PictureBox Picture5 
         Height          =   396
         Left            =   96
         ScaleHeight     =   348
         ScaleWidth      =   1860
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   648
         Width           =   1908
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Agência"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   -48
            TabIndex        =   18
            Top             =   12
            Width           =   984
         End
      End
      Begin VB.ComboBox cboCapa 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   396
         Left            =   2028
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   204
         Width           =   2604
      End
      Begin ComctlLib.TreeView TreeView1 
         Height          =   2892
         Left            =   96
         TabIndex        =   3
         Top             =   1104
         Width           =   4524
         _ExtentX        =   7980
         _ExtentY        =   5101
         _Version        =   327682
         HideSelection   =   0   'False
         Indentation     =   265
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "ImageList1"
         Appearance      =   1
         OLEDropMode     =   1
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Lote"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   396
         Left            =   4692
         TabIndex        =   20
         Top             =   648
         Width           =   2100
      End
      Begin VB.Label lblLote 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   396
         Left            =   6864
         TabIndex        =   19
         Top             =   648
         Width           =   2388
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   8448
      Top             =   3408
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   9
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTrocarOrdemDocumento.frx":0F32
            Key             =   "PASTA_FECHADA"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTrocarOrdemDocumento.frx":1044
            Key             =   "PASTA_ABERTA"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTrocarOrdemDocumento.frx":1156
            Key             =   "DOCTO"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTrocarOrdemDocumento.frx":1268
            Key             =   "DRAG_DOCTO"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTrocarOrdemDocumento.frx":1582
            Key             =   "NO_DRAG_DOCTO"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTrocarOrdemDocumento.frx":189C
            Key             =   "DEBITO"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTrocarOrdemDocumento.frx":1996
            Key             =   "CREDITO"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTrocarOrdemDocumento.frx":1A90
            Key             =   "ENVELOPE"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTrocarOrdemDocumento.frx":1DAA
            Key             =   "CHEQUE"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmTrocarOrdemDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Utilizado para identificação da capa que veio de Ilegíveis
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Type tpCapaDefault
    tpCapa                              As String
    tpIdCapa                            As String
    tpAgOrig                            As String
    tpCapaCarregada                     As Boolean 'Se true, é capa que veio de ilegiveis
    tpStatusAnterior                    As String
    tpStatusAtual                       As String
End Type
''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Determina a posicao de cada propriedade de cada node
''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Enum enumDRAG
    eImagem_Frente = 1
    eImagem_Verso
    eIdLote_Origem
    eIdLote
    eIdCapa
    eStatusCapa
    eStatusAnterior
    eIdCapa_Origem
    eIdDocto
    eIdTipoDocto
    eValorDocto
    eIdNrSequencia
    eIdStatusDocto
    eStatusDocto
    eNivel
    eAcao
    eOrdem
End Enum
''''''''''''''''''''''''''''
'Campos do objeto em questão
''''''''''''''''''''''''''''
Private Type tpDoctoDRAG
    tpImagem_Frente                     As String
    tpImagem_Verso                      As String
    tpIdLote_Origem                     As String
    tpIdLote                            As String
    tpIdCapa                            As String
    tpStatusCapa                        As String
    tpStatusAnterior                    As String
    tpIddocto                           As String
    tpIdTipoDocto                       As String
    tpValorDocto                        As String
    tpNrSequencia                       As Integer
    tpId_View                           As Integer  '- Indica de qual treeview ele se referencia
    tpIdStatusDocto                     As String
    tpStatusDocto                       As String
    tpNivel                             As String
    tpAcao                              As String
    tpOrdem                             As String
End Type
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'Area em que o mouse pode percorrer sem causar scroll
'''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Type tpAreaDRAG
    Top                                 As Long
    Bottom                              As Long
End Type


Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetScrollRange Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long, lpMinPos As Long, lpMaxPos As Long) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''
'Query que carrega os documentos da capa digitada
''''''''''''''''''''''''''''''''''''''''''''''''''
Dim qryGetDocumentosCapaTrocaOrdem      As rdoQuery
Dim qryAtualizaOrdemCapturaDocumento    As rdoQuery
Dim qryGetStatusCapa                    As rdoQuery
Dim qryAtualizaStatusCapa               As rdoQuery
Dim qryGetCapa                          As rdoQuery
Dim qryGetCapaTrocaOrdem                As rdoQuery
Dim qryGetMaloteTrocaOrdem              As rdoQuery
Dim rsTrocaOrdem                        As rdoResultset

''''''''''''''''''''''''''
'se True - Está arrastando
''''''''''''''''''''''''''
Dim m_bDragging                         As Boolean
Dim m_sUltimaImagem                     As String

Private Const EM_TROCA_ORDEM = "O"
Private Const NUMERO_TOTAL_CAPAS = 2
Private Const SM_CYHSCROLL = 3
Private Const SB_CTL = 2
Private Const SB_HORZ = 0
Private Const SB_VERT = 1
Private Const SALVA_NIVEL = "0"
'''''''''''''''''''''''
'Ações válidas para log
'''''''''''''''''''''''
'Documentos
Private Const ACAO_INSERIR = "130"
Private Const ACAO_EXCLUIR = "131"
Private Const ACAO_REORDENAR = "134"
'Capas
Private Const ACAO_VINCULO = "132"
Private Const ACAO_ILEGIVEIS = "133"
Private Const IDENTIFICACAO = "-> "

''''''''''''''''''''''''''''''''''''''''''''''''''''
'Variavel que define se deve abrir uma capa no load
'mais utilizada par abrir uma capa quando vem de ilegiveis
''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim m_CapaDefault                       As tpCapaDefault
Dim m_Capa(NUMERO_TOTAL_CAPAS - 1)      As tpCapaDefault

Dim m_bEstaFazendo                      As Boolean

'''''''''''''''''''''''''''''''''''''''''''''''''''
'Objeto Node, Quando é feito um drag, esta variavel
'conterá o item (Nó) seleciondado do TreeView
'''''''''''''''''''''''''''''''''''''''''''''''''''
Dim m_ndSelectedNode                    As Object
'''''''''''''''''''''''''''''''''''''''''''
'Objeto membro, para manipulações diversas
'''''''''''''''''''''''''''''''''''''''''''
Dim m_Docto                             As tpDoctoDRAG
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'caminho do Scroll = 0 sem scroll, 1 scroll up, 2 scroll down
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim m_Scroll_UpDown                     As Integer
'''''''''''''''''''''''''''''''''''''
'Guarda a altura do Scroll Horizontal
'''''''''''''''''''''''''''''''''''''
Dim m_lScrollHeight                     As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Guarda o tempo de atualização do status da capa, (para que não fique travada)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim sTempo                              As Integer
Dim teclou                              As Boolean

Private Function AtualizaCapa(ByVal psCapa_Origem As String, _
                              ByVal psCapa_Destino As String, _
                              ByVal psIdDocto As String, _
                              ByVal psTipoDocto As String, _
                              ByVal psOrdemCaptura) As Boolean

    On Error GoTo ERRO_ATUALIZACAPA
    
    AtualizaCapa = False
    
    
    '**********************************************'
    'Se caso o documento foi arrastado de uma capa '
    'para outra, é a procedure quem altera o idCapa.
    '**********************************************'

    With qryAtualizaOrdemCapturaDocumento
        .rdoParameters(0).Direction = rdParamReturnValue    'Retorno de dados
        .rdoParameters(1) = Geral.DataProcessamento         'Data de processamento
        .rdoParameters(2) = psCapa_Origem                   'Id da Capa origem
        .rdoParameters(3) = psCapa_Destino                  'Id da Capa destino
        .rdoParameters(4) = psIdDocto                       'Id do Documento
        .rdoParameters(5) = psTipoDocto                     'Tipo do documento
        .rdoParameters(6) = psOrdemCaptura                  'Ordem nova de captura
        .Execute
        
        If .rdoParameters(0).Value <> 0 Then
            MsgBox "Não foi possível atualizar os documentos.", vbCritical
            Exit Function
        End If
    End With
    
    AtualizaCapa = True
    Exit Function
    
ERRO_ATUALIZACAPA:

    Select Case TratamentoErro("Erro ao atualizar os documentos.", Err, rdoErrors)
    Case vbRetry
        Resume
    Case vbCancel
        
    End Select
End Function


Private Function AtualizaStatusCapa(ByVal pIdCapa As String, ByVal pStatusCapa As String) As Boolean

    On Error GoTo ERRO_ATUALIZA_STATUS_CAPA
    AtualizaStatusCapa = False
    
    
    With qryAtualizaStatusCapa
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = CDbl(pIdCapa)
        .rdoParameters(3) = CStr(pStatusCapa)
        .Execute
        
        If .rdoParameters(0).Value <> 0 Then
            MsgBox "Não foi possível atualizar a capa.", vbCritical
            Exit Function
        End If
        
    End With
    
    
    AtualizaStatusCapa = True
    Exit Function
ERRO_ATUALIZA_STATUS_CAPA:

    Select Case TratamentoErro("Erro ao atulizar o status da capa.", Err, rdoErrors)
        Case vbRetry
            Resume
    End Select
    

End Function


Private Sub endDrag()

    ''''''''''''''''''''''''''''''''''''''''''''
    'Variável membro, Se está ou não arrastando
    ''''''''''''''''''''''''''''''''''''''''''''
    m_bDragging = False
    
    '''''''''''''''''''''''''''''''''''''''''''''
    'Limpa o ultimo item, o qual o usuário soltou
    '''''''''''''''''''''''''''''''''''''''''''''
    Set TreeView1.DropHighlight = Nothing
    Set ListView1.DropHighlight = Nothing
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Finaliza o icone de drag, (Mais utilizado quando .Drag vbBeginDrag)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    TreeView1.Drag vbEndDrag
    ListView1.Drag vbEndDrag
    
    ''''''''''''''''''''''''''''''''''
    'Limpa a seleção de origem do drag
    ''''''''''''''''''''''''''''''''''
    'If Not ListView1.SelectedItem Is Nothing Then ListView1.SelectedItem.Selected = False
    'If Not TreeView1.SelectedItem Is Nothing Then TreeView1.SelectedItem.Selected = False
    

End Sub

Private Function GetStatusCapa(ByVal pCapa As Double, _
                               ByVal pDuplicidade As Long) As String
                               
    On Error GoTo ERRO_GET_STATUS_CAPA
                               
    With qryGetStatusCapa
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(4).Direction = rdParamOutput
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = CDbl(pCapa)
        .rdoParameters(3) = CDbl(pDuplicidade)
        .Execute
        
        
        If .rdoParameters(0).Value <> 0 Then GoTo ERRO_GET_STATUS_CAPA:
        
        If .rdoParameters(4).Value = "" Then GoTo ERRO_GET_STATUS_CAPA:
        
        GetStatusCapa = .rdoParameters(4)
        
    End With
    
    
    Exit Function
ERRO_GET_STATUS_CAPA:

    Call TratamentoErro("Erro ao obter o status da capa.", Err, rdoErrors)

End Function


Private Sub irIdentificacao(ByRef pNode As Node, ByVal pInserir As Boolean)

    If pInserir = True Then
        pNode.Text = IDENTIFICACAO & pNode.Text
    Else
        pNode.Text = Mid(pNode, Len(IDENTIFICACAO) + 1)
    End If

End Sub

Private Sub Limpar()

    lblMalote.Enabled = True
    txtNumMalote.Text = ""
    txtNumMalote.Enabled = True
    cmdConfirmar.Enabled = False
    lblLote.Caption = ""
    lblCapa.Caption = "Capa"
    
    
    cboCapa.Clear
    cboAgencia.Clear
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'E necessário zerar a contagem das capas ja carregadas
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    TreeView1.Tag = ""
    
    On Error Resume Next
    Lead1.Load 0, 0, 0, 0
    On Error GoTo 0


End Sub

Private Function ListView_To_TreeView(ByVal pSrcControl As ListView, _
                                      ByVal pSrcItem As ListItem, _
                                      ByVal pDestControl As TreeView, _
                                      ByVal pDestItem As Node) As Boolean



    Dim iItemCount      As Integer
    Dim sStr            As String
    
    Dim nd_Aux          As Node
    Dim nd_From         As Node
    Dim nd_Novo         As Node
    Dim iTipoFilho      As TreeRelationshipConstants
    
    ListView_To_TreeView = False
    
    If pSrcItem Is Nothing Or pDestItem Is Nothing Then Exit Function
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'se estiver querendo arrastar para debaixo dele mesmo então sai
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If pSrcItem.Key = pDestItem.Key Then Exit Function
    
    ''''''''''''''''''''''''''''''''''''''''''''
    'define nd_From como item origem do TreeView
    ''''''''''''''''''''''''''''''''''''''''''''
    Set nd_From = pDestControl.Nodes(pSrcItem.Key)
    
    '''''''''''''''''''''''''''''''''
    'Está jogando no root do TreeView
    '''''''''''''''''''''''''''''''''
    If pDestItem.Children > 0 Or (getItem(pDestItem.Tag, eIdTipoDocto)) = "1" Then
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'pega o primeiro item após o root a quem devo me referenciar
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Set nd_Aux = pDestItem.Child
        '''''''''''''''''''''''''''''''''
        'define iTipoFilho como anterior
        '''''''''''''''''''''''''''''''''
        iTipoFilho = tvwPrevious
        
        If nd_Aux Is Nothing Then
            Set nd_Aux = pDestItem
            iTipoFilho = tvwChild
        End If
        
        '''''''''''''''''''''''''''''''''''''''''''
        'Verifica se não está jogando na mesma capa
        '''''''''''''''''''''''''''''''''''''''''''
        If CLng(getItem(pSrcItem.Tag, eIdDocto)) = CLng(getItem(nd_Aux.Tag, eIdDocto)) Then
            '''''''''''''''''''''''''
            'ja é o primeiro da lista
            '''''''''''''''''''''''''
            Exit Function
        End If
    Else
        ''''''''''''''''''''''''''''''
        'define nd_Aux como referencia
        ''''''''''''''''''''''''''''''
        Set nd_Aux = pDestItem
        '''''''''''''''''''''''''''''''''
        'define iTipoFilho como proximo
        '''''''''''''''''''''''''''''''''
        iTipoFilho = tvwNext
    End If

    ''''''''''''''''''''''''''''''''''''''''''''''
    'primeiro precisa remover para depois inserir
    'porque se não remover da chave duplicada
    ''''''''''''''''''''''''''''''''''''''''''''''
    pDestControl.Nodes.Remove nd_From.Key
    
    Set nd_Novo = pDestControl.Nodes.Add(nd_Aux.Key, _
                                         iTipoFilho, _
                                         pSrcItem.Key, _
                                         pSrcItem.Text, _
                                         nd_From.Image)

    pSrcControl.ListItems.Remove nd_From.Key

    If Not nd_Novo Is Nothing Then
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Acerta o IdCapa e IdLote do node origem para o idCapa de destino
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        sStr = nd_From.Tag
        insertItem sStr, getItem(pDestItem.Tag, eIdCapa), eIdCapa
        insertItem sStr, getItem(pDestItem.Tag, eIdLote), eIdLote
        ''''''''''''''''''''''''''''''''''''''''''''''''
        'se a capa atual for diferente da capa origem
        ''''''''''''''''''''''''''''''''''''''''''''''''
        If getItem(pDestItem.Tag, eIdCapa) = getItem(pDestItem.Tag, eIdCapa_Origem) Then
            insertItem sStr, ACAO_REORDENAR, eAcao
        Else
            insertItem sStr, ACAO_INSERIR, eAcao
        End If
        nd_Novo.Tag = sStr
        
    End If

    ''''''''''''''''''''''''''''''''''''''''''
    'Reordena o numero de sequencia dos nodes
    ''''''''''''''''''''''''''''''''''''''''''
    If pDestItem.Parent Is Nothing Then
        Set nd_Aux = pDestItem
    Else
        Set nd_Aux = pDestItem.Parent
    End If
    
    CorrigirSequencia nd_Aux, 1
    ListView_To_TreeView = True

End Function



Private Sub MostraDoctos(ByVal pNode As Node)

    Dim i                   As Integer
    Dim nd_Aux              As Node
    Dim lv_Item             As ListItem
    Dim l_ColMaxWidth       As ColumnHeaders
    
    ListView1.ListItems.Clear
    
    
'    If pNode.Children > 0 Then
'        Set nd_Aux = pNode.Child
'    Else
'        Set pNode = pNode.Parent
'    End If

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Se o item selecionado não tem filhos, então define pNode como o pai do item selecionado
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If pNode.Children = 0 Then
        Set pNode = pNode.Parent
    End If

    ''''''''''''''''''''''''''''''''''''''
    'Se for inválido, não posso fazer nada
    ''''''''''''''''''''''''''''''''''''''
    If pNode Is Nothing Then Exit Sub
    
    
    '''''''''''''''''''''''''''''''''''''''''''''
    'Define nd_Aux sempre como o primeiro filho
    'porque devo mostrar sempre os filhos da capa
    '''''''''''''''''''''''''''''''''''''''''''''
    Set nd_Aux = pNode.Child
    
    Set l_ColMaxWidth = ListView1.ColumnHeaders
    
    l_ColMaxWidth(1).Width = TextWidth(l_ColMaxWidth(1).Text)
    l_ColMaxWidth(2).Width = TextWidth(l_ColMaxWidth(2).Text) - 100
    
    l_ColMaxWidth(2).Alignment = lvwColumnRight
    
    For i = 1 To pNode.Children
        Set lv_Item = ListView1.ListItems.Add(, nd_Aux.Key, nd_Aux.Text, nd_Aux.Image, nd_Aux.Image)
        
        
        If TextWidth(nd_Aux.Text) > l_ColMaxWidth(1).Width Then _
            l_ColMaxWidth(1).Width = TextWidth(nd_Aux.Text) + 1
        
        lv_Item.SubItems(1) = Format(getItem(nd_Aux.Tag, eValorDocto), "###,###,###,##0.00")
        
        
        If TextWidth(lv_Item.SubItems(1)) > l_ColMaxWidth(2).Width Then _
            l_ColMaxWidth(2).Width = TextWidth(lv_Item.SubItems(1)) + 1
        
        
        lv_Item.Tag = nd_Aux.Tag
        Set nd_Aux = nd_Aux.Next
    Next

End Sub

Private Function salvaNode(ByVal pNode As Node) As Boolean

    Dim sCapa_Origem        As String
    Dim sCapa_Destino       As String
    Dim sDocto              As String
    Dim sSequencia          As String
    Dim sImagem_Origem      As String
    Dim sImagem_Destino     As String
    Dim iIndex              As Integer
    Dim nd_Aux              As Node
    Dim sAcao               As String
    Dim sTipoDocto          As String
    
    On Error GoTo ERRO_SALVA_NODE
    
    salvaNode = False

    If pNode.Children = 0 Then
    
    
        ''''''''''''''''''''''''''''''''''''''''''''''''''
        'Na operação de Drag, Capa_Origem nunca pode mudar
        ''''''''''''''''''''''''''''''''''''''''''''''''''
        sCapa_Origem = getItem(pNode.Tag, eIdCapa_Origem) 'pega o id da capa
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Determina-se Capa_Destino na função ListView_To_TreeView
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        sCapa_Destino = getItem(pNode.Tag, eIdCapa)       'pega o id da capa
        sDocto = getItem(pNode.Tag, eIdDocto)             'pega o id Docto
        sTipoDocto = getItem(pNode.Tag, eIdTipoDocto)     'pega o tipo do docto
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Quem determina a sequencia é a função CorrigirSequencia
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        sSequencia = getItem(pNode.Tag, eIdNrSequencia)   'pega o nr sequencia
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Move a imagem de um lote para o outro se houver necessidade
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If Val(getItem(pNode.Tag, eIdLote)) <> _
           Val(getItem(pNode, eIdLote_Origem)) Then
           
           
           
            If Geral.VIPSDLL = eDllUnibanco Then
            
                ''''''''''''''
                'Imagem frente
                ''''''''''''''
                sImagem_Origem = Geral.DiretorioImagens & _
                                 Format(getItem(pNode.Tag, eIdLote_Origem), "000000000") & "\" & _
                                 getItem(pNode.Tag, eImagem_Frente)
                
                sImagem_Destino = Geral.DiretorioImagens & _
                                 Format(getItem(pNode.Tag, eIdLote), "000000000") & "\" & _
                                 getItem(pNode.Tag, eImagem_Frente)
                
                If FileExist(sImagem_Origem) Then
                    Name sImagem_Origem As sImagem_Destino
                End If
                                 
                '''''''''''''
                'Imagem verso
                '''''''''''''
                sImagem_Origem = Geral.DiretorioImagens & _
                                 Format(getItem(pNode.Tag, eIdLote_Origem), "000000000") & "\" & _
                                 getItem(pNode.Tag, eImagem_Verso)
                
                sImagem_Destino = Geral.DiretorioImagens & _
                                 Format(getItem(pNode.Tag, eIdLote), "000000000") & "\" & _
                                 getItem(pNode.Tag, eImagem_Verso)
                
                If FileExist(sImagem_Origem) Then
                    Name sImagem_Origem As sImagem_Destino
                End If
            End If
           
        End If
    
        '''''''''''''''''''''''''''
        'Atualiza capa do documento
        '''''''''''''''''''''''''''
        If Not AtualizaCapa(sCapa_Origem, _
                            sCapa_Destino, _
                            sDocto, _
                            sTipoDocto, _
                            sSequencia) Then
            Exit Function
        End If
        
        '''''''''''''''''''''''''''''''''''''''''
        'se mudou o documento de capa entao loga
        '''''''''''''''''''''''''''''''''''''''''
        If sCapa_Origem <> sCapa_Destino Then
            GravaLog sCapa_Origem, sDocto, ACAO_EXCLUIR
            GravaLog sCapa_Destino, sDocto, ACAO_INSERIR
        Else
            If Val(getItem(pNode.Tag, eAcao)) <> 0 Then
                GravaLog sCapa_Origem, sDocto, ACAO_REORDENAR
            End If
        End If
    
    Else
        For iIndex = 1 To pNode.Children
            If Not pNode.Child Is Nothing Then
                If iIndex > 1 Then
                    Set nd_Aux = nd_Aux.Next
                Else
                    Set nd_Aux = pNode.Child
                End If
                If Not salvaNode(nd_Aux) Then Exit Function
            End If
        Next iIndex
    End If
    
    
    salvaNode = True
    Exit Function
    
ERRO_SALVA_NODE:


End Function

Private Function voltarStatusCapa() As Boolean

    On Error GoTo ERRO_RETORNAR_STATUS

    Dim iIndex      As Integer
    
    voltarStatusCapa = False
    ''''''''''''''''''''''''''''''''''''''''''''''''
    'Volta o status da capa para a situacao anterior
    ''''''''''''''''''''''''''''''''''''''''''''''''
    For iIndex = 1 To TreeView1.Nodes.Count
        If (TreeView1.Nodes(iIndex).Children > 0) Or (getItem(TreeView1.Nodes(iIndex).Tag, eNivel) = 0) Then
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Se a capa não veio de ilegiveis e sim por digitacao, voltar ao status anterior
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'If (getItem(TreeView1.Nodes(iIndex).Tag, eIdCapa) <> m_CapaDefault.tpIdCapa) Then
            
                If AtualizaStatusCapa(getItem(TreeView1.Nodes(iIndex).Tag, eIdCapa), getItem(TreeView1.Nodes(iIndex).Tag, eStatusAnterior)) Then
                    If UCase(getItem(TreeView1.Nodes(iIndex).Tag, eStatusAnterior)) = "H" Then
                        GravaLog getItem(TreeView1.Nodes(iIndex).Tag, eIdCapa), 0, ACAO_ILEGIVEIS
                    End If
                End If
            
            'End If
        End If
    Next iIndex
    
    voltarStatusCapa = True
    Exit Function
    
ERRO_RETORNAR_STATUS:

    
End Function

Private Function SalvarTreeView(ByVal pTreeView As TreeView) As Boolean

    On Error GoTo ERRO_SALVATREEVIEW

    Dim iIndex          As Integer
    Dim bAtualizou      As Boolean
    Dim nd_Aux          As Node
    
    
    SalvarTreeView = False
    bAtualizou = False
    
    For iIndex = 1 To pTreeView.Nodes.Count
    
        If (pTreeView.Nodes(iIndex).Children > 0) Or _
           (getItem(pTreeView.Nodes(iIndex).Tag, eNivel) = SALVA_NIVEL) Then
           
            bAtualizou = False
            
            If Not salvaNode(pTreeView.Nodes(iIndex)) Then
                SalvarTreeView = False
                Exit Function
            End If
        
            '''''''''''''''''''''''''''''''''''''''''''''''''''
            'Se a capa que estiver sendo salva não for a capa
            'de ilegiveis, então envia para vínculo automatico
            '''''''''''''''''''''''''''''''''''''''''''''''''''
            If getItem(pTreeView.Nodes(iIndex).Tag, eIdCapa) <> m_CapaDefault.tpIdCapa And bAtualizou = False Then
            
                If AtualizaStatusCapa(getItem(pTreeView.Nodes(iIndex).Tag, eIdCapa), "8") Then
                    GravaLog getItem(pTreeView.Nodes(iIndex).Tag, eIdCapa_Origem), 0, ACAO_VINCULO
                End If
                
                bAtualizou = True
            End If
        End If
    Next iIndex

    SalvarTreeView = True
    Exit Function
    
ERRO_SALVATREEVIEW:

    MsgBox "Não foi possível salvar a sequênçia.", vbCritical

End Function

Private Function FileExist(prmPathName As String) As Boolean

    Dim lclFileNum As Integer

    On Error Resume Next
    
    lclFileNum = FreeFile
    
    Open prmPathName For Input As lclFileNum

    FileExist = IIf(Err = 0, True, False)

    Close lclFileNum

    Err = 0

End Function

'Na chamada inicial
'pNode              - é o item 1 do TreeView
'iInicioSequencia   - 1, porque ele é o primeiro item da sequencia
'
'É uma rotina RECURSIVA
'
Private Sub CorrigirSequencia(ByRef pNode As Node, ByRef iInicioSequencia As Integer)

    On Error GoTo ERRO_Sequencia
    '''''''''''''''''''''''''''''''''''''''''''''''
    Dim iIndex      As Integer      'Auxiliar do For
    '''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim nd_Aux      As Node         'Auxiliar para os nodes
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim sStr        As String 'Para guardar as propriedades do node e atualizar o numero de sequencia
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Se o node não tem filhos, então acerta o numero de sequencia
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If pNode.Children = 0 Then
        ''''''''''''''''''''''''''''''''''
        'Incremento do número de sequencia
        ''''''''''''''''''''''''''''''''''
        iInicioSequencia = iInicioSequencia + 1
        ''''''''''''''''''''''''''''''
        'Obtem as propriedades do node
        ''''''''''''''''''''''''''''''
        sStr = pNode.Tag
        ''''''''''''''''''''''''''''''''''''''''''''
        'Insere o numero sequencial em sStr -> ByRef
        ''''''''''''''''''''''''''''''''''''''''''''
        insertItem sStr, CStr(iInicioSequencia), eIdNrSequencia
        '''''''''''''''''''''''''''''''''
        'Atribui as propriedades no node
        '''''''''''''''''''''''''''''''''
        pNode.Tag = sStr
    Else
        '''''''''''''''''''''''''''''
        'Loop da quantidade de filhos
        '''''''''''''''''''''''''''''
        For iIndex = 1 To pNode.Children
            If Not pNode.Child Is Nothing Then
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'Se > 1, é porque já passou pelo else onde ja tenho nd_Aux
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If iIndex > 1 Then
                    Set nd_Aux = nd_Aux.Next
                Else
                    Set nd_Aux = pNode.Child
                End If
                ''''''''''''''''''
                'Chamada recursiva
                ''''''''''''''''''
                CorrigirSequencia nd_Aux, iInicioSequencia
            End If
        Next iIndex
    End If

    Exit Sub
    
ERRO_Sequencia:

    MsgBox "Erro ao corrigir a sequência.", vbCritical
    

End Sub

'pIdView    -   De qual treeView
'pItem      -   Um objeto item/node selecionado
Private Sub DefinirDocto(ByVal pIdView As Integer, ByVal pItem As Object)
    '''''''''''''''''''''''''''''''''''''''''''''
    'Seta as propriedades do (nó) no objeto Docto
    '''''''''''''''''''''''''''''''''''''''''''''
    m_Docto.tpId_View = pIdView
    m_Docto.tpIdCapa = getItem(pItem.Tag, eIdCapa)
    m_Docto.tpIddocto = getItem(pItem.Tag, eIdDocto)
    m_Docto.tpStatusCapa = getItem(pItem.Tag, eStatusCapa)
    m_Docto.tpIdTipoDocto = getItem(pItem.Tag, eIdTipoDocto)
    
    m_Docto.tpImagem_Frente = getItem(pItem.Tag, eImagem_Frente)
    m_Docto.tpImagem_Verso = getItem(pItem.Tag, eImagem_Verso)
    m_Docto.tpIdLote_Origem = getItem(pItem.Tag, eIdLote_Origem)
    m_Docto.tpIdLote = getItem(pItem.Tag, eIdLote)
    m_Docto.tpStatusAnterior = getItem(pItem.Tag, eStatusAnterior)
    m_Docto.tpValorDocto = getItem(pItem.Tag, eValorDocto)
    m_Docto.tpIdStatusDocto = getItem(pItem.Tag, eIdStatusDocto)
    m_Docto.tpOrdem = getItem(pItem.Tag, eOrdem)
    
End Sub

Private Function getItem(ByVal prmSource As String, ByVal prmPos As enumDRAG) As Variant

    Dim lclI As Integer, lclPos1 As Integer, lclPos2 As Integer
    Dim lclStr As String
    Dim lclLeft As String
    Dim lclRight As String
    Dim lclCopyPos As Double
    
    On Error Resume Next
    
    getItem = ""
    If prmPos < 1 Then Exit Function
    If InStr(prmSource, "|") = False Then Exit Function
    If InCount(prmSource, "|") + 1 < prmPos Then Exit Function
    
    
    lclCopyPos = prmPos
    
    getPositionOfString prmSource, lclLeft, lclRight, "|", lclCopyPos
    
    lclStr = Mid(prmSource, lclCopyPos, InStr(lclCopyPos, prmSource, "|") - lclCopyPos)


    getItem = IIf(IsNull(lclStr), "", lclStr)
    
    Err = 0

End Function

Private Function InCount(ByVal prmSource_ As String, ByVal prmSearch_ As String) As Long

    Dim lclI As Long, Count As Long
    
    InCount = 0

    If Len(prmSearch_) = 0 Then Exit Function
    If Len(prmSource_) = 0 Then Exit Function


    For lclI = 1 To Len(prmSource_)
        If InStr(lclI, prmSource_, prmSearch_) Then
            Count = Count + 1
        Else
            Exit For
        End If
        lclI = InStr(lclI, prmSource_, prmSearch_)
    Next lclI
    InCount = Count
End Function

'prmSource => em qual string se quer gravar
'prmVal    => o item a ser colocado
'prmPos    => em qual posicao sera colocado o item
Private Sub insertItem(ByRef prmSource As String, prmVal As String, ByVal prmPos As enumDRAG)

    Dim lclCopyPos As Double
    Dim lclVirgCount As Long
    
    Dim lclLeft As String
    Dim lclRight As String
    
    If prmPos < 1 Then Exit Sub
    
    'se caso nao tem ou tem virgula e o numero de virgulas é menor do que precisa
    lclVirgCount = InCount(prmSource, "|")
    If lclVirgCount < prmPos Then
        prmSource = prmSource & String(prmPos - lclVirgCount, "|")
    End If
    
    lclCopyPos = prmPos
    
    getPositionOfString prmSource, lclLeft, lclRight, "|", lclCopyPos

    prmSource = lclLeft & prmVal & lclRight

End Sub

Public Sub getPositionOfString(ByVal prmSource_ As String, ByRef prmLeft_ As String, ByRef prmRight_, ByVal prmSeparator_ As String, ByRef prmPosition_ As Double)
    Dim lclI As Long, Count As Long
    
    If InStr(prmSource_, prmSeparator_) = 0 Then Exit Sub
    
    For lclI = 1 To Len(prmSource_)
        If InStr(lclI, prmSource_, prmSeparator_) Then
            Count = Count + 1
            'da para fazer ela ficar mais rápida, bem mais rápida
            'alterando o For ex:
            'if Count < prmPosition_ then
                'count = prmPosition_
                'i = mposition
            'endif
            'mas é preciso estudar mais
            If Count = prmPosition_ Then
                prmLeft_ = Left(prmSource_, lclI - 1)
                prmRight_ = Mid(prmSource_, InStr(lclI, prmSource_, prmSeparator_))
                prmPosition_ = lclI
                Exit Sub
            End If
        Else
            Exit For
        End If
        lclI = InStr(lclI, prmSource_, prmSeparator_)
    Next lclI
    
    prmPosition_ = lclI
    
End Sub

Private Sub CriarQueries()
    ''''''''''''''''''''''''''''''''
    'Cria query de leitura de capas
    ''''''''''''''''''''''''''''''''
    Set qryGetDocumentosCapaTrocaOrdem = Geral.Banco.CreateQuery("", "{? = Call GetDocumentosCapaTrocaOrdem (?,?,?)}")
    ''''''''''''''''''''''''''''''''''''
    'Cria query de atualização de doctos
    ''''''''''''''''''''''''''''''''''''
    Set qryAtualizaOrdemCapturaDocumento = Geral.Banco.CreateQuery("", "{? = Call AtualizaOrdemCapturaDocumento (?,?,?,?,?,?)}")
    '''''''''''''''''''''''''''''''''''''''''
    'Cria query de obtencão do status da capa
    '''''''''''''''''''''''''''''''''''''''''
    Set qryGetStatusCapa = Geral.Banco.CreateQuery("", "{? = Call GetStatusCapa (?,?,?,?)}")
    '''''''''''''''''''''''''''''''''''''''''''''
    'Cria query de atualização do status da capa
    '''''''''''''''''''''''''''''''''''''''''''''
    Set qryAtualizaStatusCapa = Geral.Banco.CreateQuery("", "{ ? = Call AtualizaStatusCapa (?,?,?)}")
    ''''''''''''''''''''
    'Cria query de capa
    ''''''''''''''''''''
    Set qryGetCapa = Geral.Banco.CreateQuery("", "{ Call GetCapa(?,?)}")
    
    Set qryGetCapaTrocaOrdem = Geral.Banco.CreateQuery("", "{ Call GetCapaTrocaOrdem(?,?)}")
    
    Set qryGetMaloteTrocaOrdem = Geral.Banco.CreateQuery("", "{ Call GetMaloteTrocaOrdem(?,?)}")
    
    
End Sub

Private Function PreencheTreeView(ByVal pvsCapa As String) As Boolean

    Dim nd_Capa         As Node
    Dim nd_Docto        As Node
    Dim nd_Ref          As Node '- So para referencia
    Dim sKey            As String
    Dim sStr            As String
    Dim sImagem         As String
    Dim iNrSeq          As Integer
    Dim sStatusCapa     As String
    Dim sTipoDocto      As String
    Dim tb              As rdoResultset
    
    On Error GoTo LBL_ERRO
    
    PreencheTreeView = False
    
    If (Trim(pvsCapa) = "") And (m_CapaDefault.tpCapaCarregada = True) Then
        m_CapaDefault.tpCapaCarregada = False
        cboCapa.Text = m_CapaDefault.tpCapa
        m_CapaDefault.tpCapaCarregada = LocalizarCapa(1)
    ElseIf (Trim(pvsCapa) = "") And (m_CapaDefault.tpCapaCarregada = False) Then
        TreeView1.Nodes.Clear
        ListView1.ListItems.Clear
    End If
    
    If Not IsNumeric(pvsCapa) Then tmrAtualiza.Enabled = False: Exit Function
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Guardo no proprio TreeView o total de capas ja carregas
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Val(TreeView1.Tag) = NUMERO_TOTAL_CAPAS Then
        MsgBox "Já foram carregadas as " & NUMERO_TOTAL_CAPAS & " capas.", vbExclamation
        Exit Function
    End If
    
    ''''''''''''''''''''''''''''
    'Carrega recordset de capas
    ''''''''''''''''''''''''''''
    With qryGetDocumentosCapaTrocaOrdem
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = CDbl(pvsCapa)
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''
        'Se capa carregada = true e pvsCapa = Default não
        'faz nada, estaria querendo carregar a mesma capa
        ''''''''''''''''''''''''''''''''''''''''''''''''''
        If m_CapaDefault.tpCapaCarregada And (pvsCapa = m_CapaDefault.tpCapa) And (rsTrocaOrdem!IdCapa = m_CapaDefault.tpIdCapa) Then
            Exit Function
        ElseIf m_CapaDefault.tpCapaCarregada = False And pvsCapa = m_CapaDefault.tpCapa Then
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Se é capa de ilegiveis, então envio o idCapa de ilegiveis
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            TreeView1.Nodes.Clear
            .rdoParameters(3) = CDbl(m_CapaDefault.tpIdCapa)
        Else
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'caso contrario, envio o que está no itemData do cboAgencia
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            .rdoParameters(3) = CDbl(cboAgencia.ItemData(cboAgencia.ListIndex))
        End If

        Set tb = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
        
        If Not tb.EOF Then
            tb.MoveLast
            tb.MoveFirst
        End If

        If Val(.rdoParameters(0).Value) = 1 Then GoTo LBL_ERRO
    End With
    
    If tb.EOF Then
        MsgBox "Não foi possível encontrar a Capa.", vbExclamation
        Set tb = Nothing
        cboCapa.SetFocus
        SendKeys "{F4}"
        Exit Function
    End If
    
    ''''''''''''''''''''''''''''''''
    'Loop dos documentos desta capa
    ''''''''''''''''''''''''''''''''
    If Not tb.EOF Then
        
        Do While Not tb.EOF
            sKey = "KEY_" & tb!IdDocto
            sStr = ""
            
            If tb!TipoDocto = 1 Then
                Set nd_Capa = TreeView1.Nodes.Add(, , sKey, _
                    " Agencia (" & Format(tb!AgOrig, "0000") & ") - (" & tb!Capa & ")", _
                    "ENVELOPE", _
                    "ENVELOPE")

                Set nd_Ref = nd_Capa
                insertItem sStr, "0", eNivel 'Indica que é uma capa
            Else
                sImagem = "DOCTO"
                If tb!TipoDocto = 32 Or tb!TipoDocto = 34 Then sImagem = "CREDITO"
                If tb!TipoDocto = 33 Or tb!TipoDocto = 38 Then sImagem = "DEBITO"
                If tb!TipoDocto = 5 Or tb!TipoDocto = 6 Or tb!TipoDocto = 7 Then sImagem = "CHEQUE"
                Set nd_Ref = TreeView1.Nodes.Add(nd_Capa.Key, tvwChild, sKey, StrConv(tb!Nome, vbProperCase), sImagem)
                insertItem sStr, "1", eNivel 'Indica que é um documento
            End If
            ''''''''''''''''''''''''''''''''''
            'Obtem dados e insere em sStr
            ''''''''''''''''''''''''''''''''''
            iNrSeq = iNrSeq + 1
            insertItem sStr, tb!Frente, eImagem_Frente
            insertItem sStr, tb!Verso, eImagem_Verso
            insertItem sStr, tb!IdDocto, eIdDocto
            insertItem sStr, EM_TROCA_ORDEM, eStatusCapa
            If Val(rsTrocaOrdem!Capa) = Val(m_CapaDefault.tpCapa) Then
                insertItem sStr, m_CapaDefault.tpStatusAnterior, eStatusAnterior
            Else
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '17/10/2000
                'sempre que for voltar o status anterior desta capa, enviar para vinculo
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'insertItem sStr, IIf(tb!Status_Capa = EM_TROCA_ORDEM, "8", tb!Status_Capa), eStatusAnterior
                insertItem sStr, "8", eStatusAnterior
            End If
            insertItem sStr, tb!Valor, eValorDocto
            insertItem sStr, tb!IdCapa, eIdCapa
            insertItem sStr, tb!IdCapa, eIdCapa_Origem
            insertItem sStr, tb!IdLote, eIdLote_Origem
            insertItem sStr, tb!IdLote, eIdLote
            
            sTipoDocto = tb!TipoDocto
            If tb!TipoDocto = 7 And Val(Left(tb!Leitura, 3)) = 409 Then
                sTipoDocto = "5"
            ElseIf tb!TipoDocto = 7 And Val(Left(tb!Leitura, 3)) <> 409 Then
                sTipoDocto = "6"
            End If
            insertItem sStr, sTipoDocto, eIdTipoDocto
            
            insertItem sStr, tb!Status_Docto, eIdStatusDocto
            insertItem sStr, tb!Status_Descricao, eStatusDocto
            insertItem sStr, CStr(iNrSeq), eIdNrSequencia
            insertItem sStr, "0", eAcao 'nenhuma acao tomada
            insertItem sStr, tb!Ordem, eOrdem
            '''''''''''''''''''''''''''''''''''''''''
            'Atribui propriedades do documento ao nó
            '''''''''''''''''''''''''''''''''''''''''
            nd_Ref.Tag = sStr
            
            tb.MoveNext
        Loop
        
    End If
    
    DoEvents
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Se retorna diferente de zero é porque tem barra de scroll
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'If GetScrollRange(TreeView1.hwnd, SB_HORZ, pMinScrollPos, pMaxScrollPos) = 0 Then
    If GetScrollRange(TreeView1.hwnd, SB_HORZ, 0, 0) = 0 Then
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'se não tem barra de scroll, então desconsidera a altura da barra de scroll
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        m_lScrollHeight = 0
    Else
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Pega a altura da barra de scroll horizontal para que seja considerada no drag-scroll
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        m_lScrollHeight = GetSystemMetrics(SM_CYHSCROLL)
    End If
    
    ''''''''''''''''''''''
    'expande o node aberto
    ''''''''''''''''''''''
    If Not nd_Ref Is Nothing Then nd_Ref.EnsureVisible
    
    ''''''''''''''''''''''''''''''
    'seleciona a capa no treeview
    ''''''''''''''''''''''''''''''
    nd_Capa.Selected = True
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Envia um click na capa para abrir a imagem e definir o objeto m_Docto
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    TreeView1_NodeClick nd_Capa
    Set tb = Nothing
    
    ''''''''''''''''''''''''''''''''''''
    'Determina o numero de capas abertas
    ''''''''''''''''''''''''''''''''''''
    TreeView1.Tag = Val(TreeView1.Tag) + 1
    
    
    
    tmrAtualiza.Enabled = True
    
    PreencheTreeView = True
    cmdConfirmar.Enabled = False
    Exit Function


LBL_ERRO:
    If Err <> 35602 Then
        Call TratamentoErro("Não foi possível carregar a capa.", Err, rdoErrors, True)
    End If

    PreencheTreeView = False
    cmdConfirmar.Enabled = False

End Function

Public Sub setIdCapaDefault(ByVal pIdCapa As String)

    m_CapaDefault.tpIdCapa = pIdCapa

End Sub

Private Function TreeView_To_TreeView(ByVal pControleOrig As TreeView, _
                                      ByVal pControleDest As TreeView) As Boolean

    Dim nd_From     As Node
    Dim nd_To       As Node
    Dim nd_Aux      As Node
    Dim nd_Novo     As Node
    Dim bRemover    As Boolean
    Dim iTipoFilho  As TreeRelationshipConstants
    Dim sStr        As String
    
    TreeView_To_TreeView = False
    
    Set nd_From = pControleOrig.SelectedItem
    Set nd_To = pControleDest.DropHighlight
    
    bRemover = True
    '''''''''''''''''''''''''''''''''''''''''''''
    'Consiste no objeto nó. Não pode estar vazio
    '''''''''''''''''''''''''''''''''''''''''''''
    If (nd_From Is Nothing) Or (nd_To Is Nothing) Then Exit Function
    '''''''''''''''''''''''''''''''''''''''''''''''
    'Consiste no objeto nó. Não pode ser ele mesmo
    '''''''''''''''''''''''''''''''''''''''''''''''
    If (nd_From.Key = nd_To.Key) Then
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Mas pode acontecer de estar querendo mudar a ordem
        'do documento no mesmo TreeView ai é necessário consistir
        'novamente só que agora no objeto de origem do Drag
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If (nd_From.Key = m_ndSelectedNode.Key) Or _
           (nd_To.Key = m_ndSelectedNode.Key) Or _
           (m_ndSelectedNode Is Nothing) Then Exit Function

        '''''''''''''''''''''''''''
        'Beleza, está OK
        '
        'Mudar a ordem do documento
        '''''''''''''''''''''''''''
        
        '''''''''''''''''
        'Cria um auxiliar
        '''''''''''''''''
        Set nd_Aux = nd_To
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Define objeto local com referencia ao objeto de origem do Drag
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Set nd_From = m_ndSelectedNode
        iTipoFilho = tvwNext
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Consiste no node se está vazio, ou seja, não existe nenhum filho
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If Not nd_Aux.Child Is Nothing Then
            '''''''''''''''''''''''''''''''''''''''''''''
            'Caso estaja mesmo vazio, então pega o Parent
            '''''''''''''''''''''''''''''''''''''''''''''
            Set nd_Aux = nd_To.Child
            If nd_Aux Is Nothing Then
                MsgBox "Não é possível completar a operação.", vbCritical
                Exit Function
            End If
            iTipoFilho = tvwPrevious
        End If

        '''''''''''''''''''''''''''
        'É necessário remover antes
        '''''''''''''''''''''''''''
        pControleDest.Nodes.Remove m_ndSelectedNode.Key



        Set nd_Novo = pControleDest.Nodes.Add(nd_Aux.Key, _
                                              iTipoFilho, _
                                              nd_From.Key, _
                                              nd_From.Text, _
                                              nd_From.Image, _
                                              nd_From.SelectedImage)

        bRemover = False
    ''''''''''''''''''''''''''''''''''''''''''''
    'Verifica se está jogando no root (destino)
    ''''''''''''''''''''''''''''''''''''''''''''
    ElseIf CLng(getItem(nd_To.Root.Tag, eIdDocto)) = CLng(getItem(nd_To.Tag, eIdDocto)) Then
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Caso estiver jogando no root, então seleciona o primeiro filho após o root
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Set nd_Aux = nd_To.Child
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Consiste no node se está vazio, ou seja, não existe nenhum filho
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        iTipoFilho = tvwFirst
        If nd_Aux Is Nothing Then
            ''''''''''''''''''''''''''''''''''''''''''''''''''
            'Caso estaja mesmo vazio, então define a variavel
            'iTipoFilho como filho e nd_Aux como root
            ''''''''''''''''''''''''''''''''''''''''''''''''''
            iTipoFilho = tvwChild
            Set nd_Aux = nd_To
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Depois do primeiro item selecionado, crio como o primeiro dos filhos
        '
        'ATENCAO: iTipoFilho varia de acordo com a existencia do filho
        '         Caso exista filho iTipoFilho = tvwFirst
        '         Caso contrato     iTipoFilho = tvwChild
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Set nd_Novo = pControleDest.Nodes.Add(nd_Aux.Key, _
                                              iTipoFilho, _
                                              nd_From.Key, _
                                              nd_From.Text, _
                                              nd_From.Image, _
                                              nd_From.SelectedImage)
    Else
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Caso contrario, joga o node para o treeview selecionado como proximo item
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Set nd_Novo = pControleDest.Nodes.Add(nd_To.Key, _
                                              tvwNext, _
                                              nd_From.Key, _
                                              nd_From.Text, _
                                              nd_From.Image, _
                                              nd_From.SelectedImage)
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Acerta o IdCapa do node origem para o idCapa de destino
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sStr = nd_From.Tag
    insertItem sStr, getItem(nd_To.Tag, eIdCapa), eIdCapa
    nd_Novo.Tag = sStr
    ''''''''''''''''''''''
    'Remove o node origem
    ''''''''''''''''''''''
    If bRemover Then pControleOrig.Nodes.Remove nd_From.Key
    
    ''''''''''''''''''''''''''''''''''''''''''
    'Reordena o numero de sequencia dos nodes
    ''''''''''''''''''''''''''''''''''''''''''
    CorrigirSequencia pControleDest.Nodes(1), 1
    
    ''''''''''''''''''''''''''''''
    'Seleciona o novo item criado
    ''''''''''''''''''''''''''''''
    pControleDest.Nodes(nd_Novo.Index).Selected = True
    pControleDest.SetFocus
    
    TreeView_To_TreeView = True
    
End Function

Private Sub cboAgencia_Click()
    Dim Msg As String
    Dim bPosicionou As Boolean
    Dim iIndex      As Integer
    Dim sStr        As String
    
    If Len(Trim(cboAgencia.Text)) = 0 Then
        Exit Sub
    End If
    
    If rsTrocaOrdem.RowCount > 0 Then
        rsTrocaOrdem.MoveFirst
        Do While Not rsTrocaOrdem.EOF
            If rsTrocaOrdem!IdCapa = cboAgencia.ItemData(cboAgencia.ListIndex) Then
                Exit Do
            End If
            rsTrocaOrdem.MoveNext
        Loop
    End If
    lblLote.Caption = Format(rsTrocaOrdem!IdLote, "0000-00000")
    
    If rsTrocaOrdem!IdEnv_Mal = "E" Then
        lblCapa.Caption = "Envelope"
        lblMalote.Enabled = False
        txtNumMalote.Text = ""
        txtNumMalote.Enabled = False
    Else
        lblCapa.Caption = "Malote"
        lblMalote.Enabled = True
        txtNumMalote.Enabled = True
    End If
    
    
    Msg = "Esta " & _
        " capa não está disponível para Troca de Ordem, porque encontra-se " & vbCrLf
    Select Case rsTrocaOrdem!Status
        Case "0"      'Capa cadastrada
            Msg = Msg & "para Captura de Imagens. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "1"      'Capa digitalizada
            Msg = Msg & "para Complementação. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "2"      'Capa em complementação
            Msg = Msg & "em Complementação. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "3"      'Capa complementada, mas com pendência
            ' Nao existe mais este status
            Msg = Msg & "com Status inválido (3). "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "6"      'Capa para Alcada
            Msg = Msg & "para Alçada. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "7"      'Capa Vinculo Manual
            Msg = Msg & "para Vinculo Manual. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "8"      'Capa para Vínculo Automatico
            If Val(m_CapaDefault.tpCapa) <> Val(rsTrocaOrdem!Capa) Then
               Msg = Msg & "para Vínculo Automático. "
               MsgBox Msg, vbInformation + vbOKOnly, App.Title
               If m_CapaDefault.tpCapaCarregada Then cboCapa.Text = m_CapaDefault.tpCapa
               Exit Sub
            End If
        Case "9"      'Capa p/ Vinc. Automatico, enviada pelo Prova Zero
            Msg = Msg & "para Vínculo Automático. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "A"      'Capa para Recaptura
            Msg = Msg & "para Recaptura. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "B"      'Capa em Recaptura
            Msg = Msg & "em Recaptura. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "D"      'Capa devolvida pelo Sistema
            Msg = Msg & "Devolvida pelo Sistema. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "E"      'Capa Expedida
            Msg = Msg & "Expedida. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "F"      'Capa Devolvida pelo Robo
            Msg = Msg & "Devolvida pelo Robô. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "G"      'Capa em Prova Zero
            If m_CapaDefault.tpIdCapa = "" Then
                Msg = Msg & "em Prova Zero. "
                MsgBox Msg, vbInformation + vbOKOnly, App.Title
                cmdLimpar_Click
                Exit Sub
            End If
        Case "H"      'Capa em Ilegiveis
            If m_CapaDefault.tpIdCapa = "" Then
                Msg = Msg & "em Ilegíveis. "
                MsgBox Msg, vbInformation + vbOKOnly, App.Title
                cmdLimpar_Click
                Exit Sub
            End If
        Case "I"      'Capa em Alcada
            Msg = Msg & "em Alçada. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "J"      'Capa em Vinculo Manual
            Msg = Msg & "em Vínculo Manual. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
            
        Case "O", "G"     'Capa em Troca de ordem
            If rsTrocaOrdem!Intervalo <= Geral.Intervalo Then
                For iIndex = 0 To UBound(m_Capa)
                    If m_Capa(iIndex).tpCapaCarregada Then
                        If Val(m_Capa(iIndex).tpCapa) <> Val(cboCapa.Text) Then
                            Msg = Msg & "em Troca de Ordem por outra estação. "
                            MsgBox Msg, vbInformation + vbOKOnly, App.Title
                            Limpar
                            Exit Sub
                        End If
                    End If
                Next iIndex
            End If
        Case "K"      'Capa em Expedicao
            Msg = Msg & "em Expedição. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "L"      'Capa para Confirmação de Agencia e Conta
            Msg = Msg & "para Confirmação de Agência e Conta. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "M"      'Capa em Confirmação de Agencia e Conta
            Msg = Msg & "em Confirmação de Agência e Conta. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "N"      'Capa para CSP
            Msg = Msg & "para C.S.P. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "P"      'Capa Devolvida pela Preparação
            ' Depois que o Robot transmitir a ocorrencia
            ' desta capa, ele mudarah o status para "X"
            Msg = Msg & "para Transmissão. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "Q"      'Capa em CSP
            Msg = Msg & "em C.S.P. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "R"      'Capa para Transmissão
            Msg = Msg & "para Transmissão. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "S"      'Capa em Transmissão
            Msg = Msg & "em Transmissão. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "T"      'Capa Transmitida
            Msg = Msg & "Transmitida. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "U"      'Capa em Confirmação de Lote
            Msg = Msg & "em Confirmação de Lote. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "V"      'Capa em Verificacao
            Msg = Msg & "em Verificação. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "W"      'Capa em Estorno
            Msg = Msg & "em Estorno. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "X"      'Capa finalizada com ocorrência
            Msg = Msg & "finalizada com ocorrência. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        
        Case "Y"      'Capa para Correção de Agencia e Conta
            Msg = Msg & "para Correção de Agência e Conta. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "Z"      'Capa em Correção de Agencia e Conta
            Msg = Msg & "em Correção de Agência e Conta. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
    End Select

    If PreencheTreeView(rsTrocaOrdem!Capa) Then
    
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Não atualizo o status da Capa, porque se a capa que estiver sendo
        'carregada veio de Ilegiveis (m_CapaDefault.tpCapa), o status deve
        'continuar como H, caso contrario atualiza para "O" em troca de ordem
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'If Val(m_CapaDefault.tpIdCapa) <> Val(rsTrocaOrdem!IdCapa) Then
            m_Capa(TreeView1.Tag - 1).tpCapa = rsTrocaOrdem!Capa
            m_Capa(TreeView1.Tag - 1).tpCapaCarregada = False
            m_Capa(TreeView1.Tag - 1).tpAgOrig = Format(rsTrocaOrdem!AgOrig, "0000")
            m_Capa(TreeView1.Tag - 1).tpIdCapa = rsTrocaOrdem!IdCapa
            m_Capa(TreeView1.Tag - 1).tpStatusAnterior = IIf(rsTrocaOrdem!Status = EM_TROCA_ORDEM, "8", rsTrocaOrdem!Status)
            m_Capa(TreeView1.Tag - 1).tpStatusAtual = EM_TROCA_ORDEM
            m_Capa(TreeView1.Tag - 1).tpCapaCarregada = True
            
            AtualizaStatusCapa rsTrocaOrdem!IdCapa, EM_TROCA_ORDEM
            
            cmdLimpar.Enabled = True
        'End If
        
    End If
    
    
    
End Sub


Private Sub cboCapa_Click()

    If m_bEstaFazendo Then Exit Sub
    
    m_bEstaFazendo = True


    cboAgencia.Clear
    If rsTrocaOrdem.RowCount > 0 Then
        rsTrocaOrdem.MoveFirst
        Do While Not rsTrocaOrdem.EOF
            If rsTrocaOrdem!Capa = Val(cboCapa.Text) Then
                cboAgencia.AddItem Format(rsTrocaOrdem!AgOrig, "0000")
                cboAgencia.ItemData(cboAgencia.NewIndex) = rsTrocaOrdem!IdCapa
            End If
            rsTrocaOrdem.MoveNext
        Loop
    End If
    If cboAgencia.ListCount = 1 Then
        cboAgencia.Text = cboAgencia.List(0)
    Else
        cboAgencia.SetFocus
        SendKeys "{F4}"
    End If
    
    m_bEstaFazendo = False
    
End Sub

Private Sub cboCapa_GotFocus()
    SelecionarTexto cboCapa
End Sub


Private Sub cboCapa_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Val(TreeView1.Tag) = NUMERO_TOTAL_CAPAS Then
            MsgBox "Já foram carregadas as " & NUMERO_TOTAL_CAPAS & " capas.", vbExclamation
            Exit Sub
        End If
        If LocalizarCapa(1) Then
            txtNumMalote.Text = FormataMalote(txtNumMalote.Text)
        End If
    Else
        SoNumero KeyAscii
        If Len(cboCapa.Text) > 13 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub cmdConfirmar_Click()

    Dim iIndex      As Integer
    Dim nd_Aux      As Node

    Screen.MousePointer = vbHourglass
    
    Geral.Banco.BeginTrans
    
    cmdFechar.Enabled = False
    cmdLimpar.Enabled = False
    
    If SalvarTreeView(TreeView1) = False Then
        Geral.Banco.RollbackTrans
        MsgBox "Não foi possível atualizar a troca de ordem dos documentos.", vbExclamation
    Else
        Geral.Banco.CommitTrans
    End If
    
    TreeView1.Tag = ""
    
    Erase m_Capa
    PreencheTreeView ""
    
    If m_CapaDefault.tpCapaCarregada = False Then
        Limpar
    End If
    
    cmdConfirmar.Enabled = False
    cmdLimpar.Enabled = False
    cmdFechar.Enabled = True
    
    cboCapa.SetFocus
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub CmdFechar_Click()

    '''''''''''''''''''''''''''
    'Unload do form ja faz isto
    '''''''''''''''''''''''''''
'    Limpar
'    voltarStatusCapa

    Unload Me

End Sub

Private Sub cmdLimpar_Click()

    Limpar
    
    voltarStatusCapa
    
    ''''''''''''''''''''''''''''''''''
    'Limpa o TreeView automaticamente
    ''''''''''''''''''''''''''''''''''
    If m_CapaDefault.tpCapaCarregada Then
        cboCapa.Text = m_CapaDefault.tpCapa
        m_CapaDefault.tpCapaCarregada = False
        cboCapa_KeyPress vbKeyReturn
        m_CapaDefault.tpCapaCarregada = True
    Else
        PreencheTreeView ""
    End If
    cmdLimpar.Enabled = False
    cboCapa.SetFocus
    
End Sub

Private Sub Form_Activate()

    On Error GoTo ERRO_ACTIVATE:

    Dim tb      As rdoResultset
    Dim i       As Integer
    
    'Inclusão de chamada a rotina AtualizaAtividade
    Call AtualizaAtividade(20)
    
    With Lead1
        .UnlockSupport L_SUPPORT_EXPRESS, "YXPQ3XPPVT"
        .UnlockSupport L_SUPPORT_GIFLZW, "0K3RV9UY3EY"
        .UnlockSupport L_SUPPORT_TIFLZW, "9LE75L0FDXHK"
    End With
    
    CriarQueries
    
    m_bDragging = False
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'se alguma tela definiu a capa a ser aberta, então abre
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If m_CapaDefault.tpIdCapa <> "" Then
    
    
        With qryGetCapa
            .rdoParameters(0) = Geral.DataProcessamento
            .rdoParameters(1) = m_CapaDefault.tpIdCapa
            
            Set tb = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
        End With
    
        If Not tb.EOF Then
            '''''''''''''''''''''''''''''''''''''''
            'Obtem a capa e o status atual da capa.
            '''''''''''''''''''''''''''''''''''''''
            m_CapaDefault.tpCapa = tb!Capa
            m_CapaDefault.tpStatusAnterior = tb!Status
            m_CapaDefault.tpStatusAtual = EM_TROCA_ORDEM
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Sempre que estiver vindo de ilegiveis, sempre teremos o idCapa
            'portanto não me preocupo, porque o KeyPress do combo Capa
            'vai abrir a capa
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            cboCapa.Text = m_CapaDefault.tpCapa
            DoEvents
            
            
            qryGetCapaTrocaOrdem.rdoParameters(0) = Geral.DataProcessamento
            qryGetCapaTrocaOrdem.rdoParameters(1) = m_CapaDefault.tpCapa
            Set rsTrocaOrdem = qryGetCapaTrocaOrdem.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
            
            
            m_CapaDefault.tpCapaCarregada = PreencheTreeView(m_CapaDefault.tpCapa)
            
            rsTrocaOrdem.MoveLast
            rsTrocaOrdem.MoveFirst

            Do While Not rsTrocaOrdem.EOF
                cboAgencia.AddItem Format(rsTrocaOrdem!AgOrig, "0000")
                cboAgencia.ItemData(cboAgencia.NewIndex) = rsTrocaOrdem!IdCapa
                
                rsTrocaOrdem.MoveNext
            Loop
            rsTrocaOrdem.MoveFirst
            
            For i = 0 To cboAgencia.ListCount - 1
                If cboAgencia.ItemData(i) = m_CapaDefault.tpIdCapa Then
                    cboAgencia.ListIndex = i
                End If
            Next i
            'cboAgencia.Text = Format(tb!AgOrig, "0000")
            m_CapaDefault.tpAgOrig = Format(tb!AgOrig, "0000")
            
            
            'm_CapaDefault.tpCapaCarregada = LocalizarCapa(1)
            
            If m_CapaDefault.tpCapaCarregada Then SelecionarTexto cboCapa: cboCapa.SetFocus
            
            cmdLimpar.Enabled = False
        End If
        
        Set tb = Nothing
    End If
    
    Exit Sub
ERRO_ACTIVATE:

    Select Case TratamentoErro("Erro ao iniciar a Troca de Ordem.", Error, rdoErrors)
    Case vbRetry
        Resume
    End Select
    
End Sub


' Chave = 1 - Capa; 2 - Num_Malote
Private Function LocalizarCapa(ByVal Chave As Integer) As Boolean

    On Error GoTo Erro_LocalizarCapa

    LocalizarCapa = False

    If Chave = 1 And cboCapa.ListCount > 0 Then
        Exit Function
    End If
    
    If Chave = 1 Then
        If Not IsNumeric(cboCapa.Text) Then
            SelecionarTexto cboCapa
            MsgBox "Capa informada não é válida.", vbExclamation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Function
        End If
        qryGetCapaTrocaOrdem.rdoParameters(0) = Geral.DataProcessamento
        qryGetCapaTrocaOrdem.rdoParameters(1) = Val(cboCapa.Text)
        Set rsTrocaOrdem = qryGetCapaTrocaOrdem.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    Else
        If Not IsNumeric(txtNumMalote.Text) Then
            SelecionarTexto txtNumMalote
            MsgBox "Número do Malote informado não é válido.", vbExclamation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Function
        End If
        qryGetMaloteTrocaOrdem.rdoParameters(0) = Geral.DataProcessamento
        qryGetMaloteTrocaOrdem.rdoParameters(1) = Val(txtNumMalote.Text)
        Set rsTrocaOrdem = qryGetMaloteTrocaOrdem.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End If
    If rsTrocaOrdem.EOF Then
        If Chave = 1 Then
            SelecionarTexto cboCapa
            MsgBox "Capa informada não foi encontrada.", vbExclamation + vbOKOnly, App.Title
            SelecionarTexto cboCapa
        Else
            SelecionarTexto txtNumMalote
            MsgBox "Número do Malote informado não foi encontrado.", vbExclamation + vbOKOnly, App.Title
            SelecionarTexto txtNumMalote
        End If
        Exit Function
    End If
    cboCapa.Clear
    cboAgencia.Clear
    While Not rsTrocaOrdem.EOF
        If Chave = 1 Then
            cboCapa.Text = Format(rsTrocaOrdem!Capa, IIf(rsTrocaOrdem!IdEnv_Mal = "E", "00000000", "00000000000000"))
            If rsTrocaOrdem!IdEnv_Mal = "E" Then
                txtNumMalote.Text = ""
            Else
                txtNumMalote.Text = Format(rsTrocaOrdem!Num_Malote, "00000000000")
            End If
            cboAgencia.AddItem Format(rsTrocaOrdem!AgOrig, "0000")
            cboAgencia.ItemData(cboAgencia.NewIndex) = rsTrocaOrdem!IdCapa
        Else
            cboCapa.AddItem Format(rsTrocaOrdem!Capa, IIf(rsTrocaOrdem!IdEnv_Mal = "E", "00000000", "00000000000000"))
        End If
        rsTrocaOrdem.MoveNext
    Wend
    If Chave = 1 Then
        If cboAgencia.ListCount = 1 Then
            cboAgencia.ListIndex = 0
        Else
            If (m_CapaDefault.tpIdCapa <> "") And (Trim(Val(m_CapaDefault.tpCapa)) = Trim(CDbl(cboCapa.Text))) Then
                cboAgencia.Text = m_CapaDefault.tpAgOrig
                m_CapaDefault.tpCapaCarregada = True
            Else
                DoEvents
                cboAgencia.SetFocus
                SendKeys "{F4}"
            End If
        End If
    Else
        If cboCapa.ListCount = 1 Then
            cboCapa.ListIndex = 0
        Else
            DoEvents
            cboCapa.SetFocus
            SendKeys "{F4}"
        End If
    End If
    
    Screen.MousePointer = vbDefault
    LocalizarCapa = IIf((rsTrocaOrdem.RowCount = 1) Or (m_CapaDefault.tpCapaCarregada = True), True, False)
    
    Exit Function
Erro_LocalizarCapa:

    Select Case TratamentoErro("Erro ao localizar a capa.", Err, rdoErrors)
    Case vbRetry
        Resume
    End Select

End Function


Public Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim Ret As Long
    
    hCtl = Me.Lead1.hwnd
    
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



Public Sub cmdFrenteVerso_Click()

  On Error GoTo ERRO_FRENTEVERSO

  If teclou Then Exit Sub

  If FrmImagem.Visible = False Then Exit Sub

  teclou = True
  'só muda para frente/verso qdo docto vem da Ls500 e da Vips,
  'poi, o canon não gera verso.
  If (m_Docto.tpOrdem = "0") Or (m_Docto.tpOrdem = "2") Then
    If Lead1.Tag = "V" Then
        Lead1.Tag = "F"     'se verso, mostrar frente
        With Lead1
            .AutoRepaint = False
            If Geral.VIPSDLL = eDllProservi Then
              .Load Geral.DiretorioImagens & m_Docto.tpImagem_Frente, 0, 0, 1
            Else
              .Load Geral.DiretorioImagens & Format(m_Docto.tpIdLote, "000000000") & "\" & m_Docto.tpImagem_Frente, 0, 0, 1
            End If
  
            'se ls500 mostrar mais escuro
            If (m_Docto.tpOrdem = "2") Then
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
              .Load Geral.DiretorioImagens & Trim(m_Docto.tpImagem_Verso), 0, 0, 1
            Else
              .Load Geral.DiretorioImagens & Format(m_Docto.tpIdLote, "000000000") & "\" & m_Docto.tpImagem_Verso, 0, 0, 1
            End If
  
            'se ls500 mostrar mais escuro
            If (m_Docto.tpOrdem = "2") Then
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
  'frmImagem.Visible = False
  Screen.MousePointer = vbDefault
  Select Case TratamentoErro("Erro ao Manipular Imagem do Documento.", Err, rdoErrors)
    Case vbCancel, vbRetry
      Unload Me
  End Select
End Sub

Private Sub Form_Load()
    m_CapaDefault.tpIdCapa = ""
    m_CapaDefault.tpCapa = ""
    m_CapaDefault.tpCapaCarregada = False
    
    Erase m_Capa
    
    m_lScrollHeight = 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Limpar
    voltarStatusCapa



    m_CapaDefault.tpCapa = ""
    m_CapaDefault.tpCapaCarregada = False
    m_CapaDefault.tpIdCapa = 0
    m_CapaDefault.tpStatusAnterior = ""
    m_CapaDefault.tpStatusAtual = ""
    
    Erase m_Capa

    Set qryGetDocumentosCapaTrocaOrdem = Nothing
    Set qryAtualizaOrdemCapturaDocumento = Nothing
    Set qryGetStatusCapa = Nothing
    Set qryAtualizaStatusCapa = Nothing
    Set qryGetCapa = Nothing
    Set qryGetCapaTrocaOrdem = Nothing
    Set qryGetMaloteTrocaOrdem = Nothing
    Set rsTrocaOrdem = Nothing
    
End Sub

Private Sub ListView1_Click()

    If ListView1.SelectedItem Is Nothing Then Exit Sub
    If ListView1.ListItems(ListView1.SelectedItem.Key) Is Nothing Then Exit Sub
    If ListView1.SelectedItem.Selected = False Then Exit Sub
    
    TreeView1.Nodes(ListView1.SelectedItem.Key).Selected = True
    
    DefinirDocto 1, ListView1.SelectedItem

    MostraImagem

End Sub

Private Sub ListView1_DragDrop(Source As Control, X As Single, Y As Single)
    endDrag
End Sub

Private Sub ListView1_ItemClick(ByVal Item As ComctlLib.ListItem)


    If Item Is Nothing Then Exit Sub
    
    If TreeView1.Nodes(Item.Key) Is Nothing Then Exit Sub

    Set m_ndSelectedNode = TreeView1.Nodes(Item.Key)
    
    
    If m_bDragging Then
        DefinirDocto 1, m_ndSelectedNode
        MostraImagem
    End If
    
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)

    m_bDragging = Not (Shift = vbShiftMask)
    
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Se foi MouseDown em um item, então define um objeto
    'como o item selecionado
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Not ListView1.HitTest(X, Y) Is Nothing Then
        'ListView1.HitTest(x, y).Selected = True
        'Set m_ndSelectedNode = ListView1.SelectedItem
        Set m_ndSelectedNode = TreeView1.Nodes(ListView1.SelectedItem.Key)
        '''''''''''''''''''''''''''''''''''''''''''''
        'Definir m_Docto como o item a ser arrastado
        '''''''''''''''''''''''''''''''''''''''''''''
        If m_ndSelectedNode Is Nothing Then Exit Sub
        DefinirDocto 1, m_ndSelectedNode
        ''''''''''''''''''''''''''''''''''''''''''''
        'Variavel membro, se está ou não arrastando
        ''''''''''''''''''''''''''''''''''''''''''''
        m_bDragging = True
    Else
        ''''''''''''''''''''''''''''''''
        'Caso contrario, limpa o objeto
        ''''''''''''''''''''''''''''''''
        Set m_ndSelectedNode = Nothing
        m_bDragging = False
    End If
    
End Sub

Private Sub ListView1_OLECompleteDrag(Effect As Long)
    m_bDragging = False
End Sub

Private Sub ListView1_OLEStartDrag(Data As ComctlLib.DataObject, AllowedEffects As Long)
    m_bDragging = True
End Sub

Private Sub Timer1_Timer()
    
    Dim iWM_UPDOWN  As Long
    
    If m_bDragging Then
    
        If m_Scroll_UpDown = 0 Then Timer1.Enabled = False: Exit Sub
        
        iWM_UPDOWN = IIf(m_Scroll_UpDown = 1, SB_LINEUP, SB_LINEDOWN)
        
        Call SendMessage(TreeView1.hwnd, WM_VSCROLL, iWM_UPDOWN, 0)
        
    End If
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
    
    Dim i As Integer
    Dim Ret As Long
    
    If m_sUltimaImagem = m_Docto.tpIddocto Then Exit Sub
    
    hCtl = Lead1.hwnd
    
    ' Coloca imagem na tela
    With Lead1
        .Tag = "F"
        .AutoRepaint = False
        If Geral.VIPSDLL = eDllProservi Then
            .Load Geral.DiretorioImagens & m_Docto.tpImagem_Frente, 0, 0, 1
        Else
            .Load Geral.DiretorioImagens & Format(m_Docto.tpIdLote_Origem, "000000000") & "\" & m_Docto.tpImagem_Frente, 0, 0, 1
        End If
       
       ' se imagem for da ls500, mostra mais escura
       If m_Docto.tpOrdem <> "2" Then
          .Intensity 220
       Else
          .Intensity 140
       End If
       ' se imagem for da canon, mostra a 50%
       If m_Docto.tpOrdem <> "1" Then
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
    
    m_sUltimaImagem = m_Docto.tpIddocto
    
    DoEvents
    
    Exit Sub

ERRO_MOSTRAIMAGEM:
    Screen.MousePointer = vbDefault
    MsgBox "Não foi possível exibir a Imagem do Documento, imagem não encontrada.", vbInformation, App.Title
    Call HDObjetosImagem(False)
End Sub

Private Sub HDObjetosImagem(bValor As Boolean)

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
    Call TratamentoErro("Erro ao preparar botões de manipulação de Imagens.", Err, rdoErrors)
End Sub

Private Sub tmrAtualiza_Timer()
    tmrAtualiza.Enabled = False

    If m_Docto.tpIdCapa <> 0 Then
        sTempo = sTempo + Int(tmrAtualiza.Interval / 1000)
    
        If sTempo + Int(tmrAtualiza.Interval / 1000) >= Geral.Intervalo Then
            'Atualizar o Status da Capa
            Call AtualizaStatusCapa(m_Docto.tpIdCapa, m_Docto.tpStatusCapa)
            sTempo = 0
        End If
    End If

    tmrAtualiza.Enabled = True

End Sub

Private Sub TreeView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim i As Integer

    If Button = vbRightButton Then
        If Not TreeView1.HitTest(X, Y) Is Nothing Then
        
            For i = 1 To TreeView1.Nodes.Count
                If TreeView1.Nodes(i).Key <> TreeView1.HitTest(X, Y).Key Then
                    If InStr(TreeView1.Nodes(i).Text, IDENTIFICACAO) Then
                        irIdentificacao TreeView1.Nodes(i), False
                    End If
                End If
            Next i
            
            irIdentificacao TreeView1.HitTest(X, Y), Not CBool(InStr(TreeView1.HitTest(X, Y).Text, IDENTIFICACAO))
            
        End If
    End If

End Sub

Private Sub TreeView1_NodeClick(ByVal Node As ComctlLib.Node)

    ''''''''''''''''''''''''''''
    'pode acontecer de vir vazio
    ''''''''''''''''''''''''''''
    If Node Is Nothing Then Exit Sub

    '''''''''''''''''''''''''''''''''''''''
    'Cria o objeto "m_Docto" com idView = 1
    '''''''''''''''''''''''''''''''''''''''
    DoEvents
    
    DefinirDocto 1, Node
    
    MostraImagem
    
    MostraDoctos Node
    
End Sub

Private Sub TreeView1_OLECompleteDrag(Effect As Long)

    m_bDragging = False

End Sub

Private Sub TreeView1_OLEDragDrop(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim nd_Aux      As Node
    Dim iIndex      As Integer

    m_Scroll_UpDown = 0
    '''''''''''''''''''''''''''''''''''''''''''''''''
    'Define o item selecionado como o ultimo do drag
    '''''''''''''''''''''''''''''''''''''''''''''''''
    'Set TreeView1.SelectedItem = TreeView1.DropHighlight

    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Rotina para tratamento do drag entre os documentos
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    For iIndex = ListView1.ListItems.Count To 1 Step -1
        If ListView1.ListItems(iIndex).Selected Then
            If ListView_To_TreeView(ListView1, ListView1.ListItems(iIndex), TreeView1, TreeView1.DropHighlight) Then
                'TreeView1_NodeClick TreeView1.SelectedItem
                cmdConfirmar.Enabled = True
                cmdLimpar.Enabled = True
            End If
        End If
    Next iIndex
    
'    'MostraDoctos TreeView1.DropHighlight
'    If Not TreeView1.DropHighlight Is Nothing Then
'        DefinirDocto 1, TreeView1.DropHighlight
'        MostraImagem
'        MostraDoctos TreeView1.DropHighlight
'
'        ListView1.ListItems(TreeView1.DropHighlight.Key).Selected = True
'        TreeView1.Nodes(TreeView1.DropHighlight.Key).Selected = True
'
'    End If

    If TreeView1.DropHighlight Is Nothing Then GoTo LBL_LIMPA

    Set nd_Aux = TreeView1.DropHighlight.Next
    
    If nd_Aux Is Nothing Then Set nd_Aux = TreeView1.DropHighlight.Child
    
    
    
    If Not nd_Aux Is Nothing Then
    
        If nd_Aux.Children > 0 Then Set nd_Aux = TreeView1.DropHighlight.Child
        
        DefinirDocto 1, nd_Aux
        MostraImagem
        MostraDoctos nd_Aux
        
        TreeView1_NodeClick nd_Aux
        TreeView1.Nodes(nd_Aux.Key).Selected = True
        On Error Resume Next
        ListView1.ListItems(nd_Aux.Key).Selected = True
        On Error GoTo 0
    End If


LBL_LIMPA:

    ''''''''''''''
    'Limpa o drag
    ''''''''''''''
    endDrag

End Sub

Private Sub TreeView1_OLEDragOver(Data As ComctlLib.DataObject, _
                                  Effect As Long, _
                                  Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single, _
                                  State As Integer)

    Dim tpArea      As tpAreaDRAG

    '''''''''''''''''''''''''''''''''''''''''''
    'Só seleciona um item se estiver arrastando
    '''''''''''''''''''''''''''''''''''''''''''
    If m_bDragging And State = 2 Then
        If m_ndSelectedNode Is Nothing Then Exit Sub
        '''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Se o tipo de documento = 1, não criar DropHighlight
        '''''''''''''''''''''''''''''''''''''''''''''''''''''
        If m_Docto.tpIdTipoDocto = 1 Then Exit Sub

'        If TreeView1.HitTest(X, Y) Is Nothing Then Exit Sub

        tpArea.Top = TextHeight(m_ndSelectedNode) * 2 'Duas linhas
        tpArea.Bottom = TreeView1.Height - tpArea.Top - (m_lScrollHeight * Screen.TwipsPerPixelY)
        
        
        If Y < tpArea.Top Then 'Scroll para cima
            m_Scroll_UpDown = 1
            Timer1.Enabled = True
        ElseIf Y > tpArea.Bottom Then
            m_Scroll_UpDown = 2
            Timer1.Enabled = True
        Else
            m_Scroll_UpDown = 0
            Timer1.Enabled = False
        End If

        Set TreeView1.DropHighlight = TreeView1.HitTest(X, Y)
        
        If Not TreeView1.DropHighlight Is Nothing Then
            If getItem(m_ndSelectedNode.Tag, eIdDocto) = getItem(TreeView1.DropHighlight.Tag, eIdDocto) Then
                Effect = 0
            End If
        End If
    ElseIf State = 1 Then
        Set TreeView1.DropHighlight = Nothing
        m_Scroll_UpDown = 0
        Timer1.Enabled = False
    End If
    

End Sub

Private Sub txtNumMalote_GotFocus()
    SelecionarTexto txtNumMalote
End Sub

Private Sub txtNumMalote_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(txtNumMalote) Then
            If Len(txtNumMalote) = 12 Then
                If Left(txtNumMalote.Text, 2) <> "09" Then
                    MsgBox "Número do Malote inválido.", vbExclamation + vbOKOnly, App.Title
                    cmdLimpar_Click
                    Exit Sub
                End If
            End If
            
            If LocalizarCapa(2) Then
                txtNumMalote.Text = FormataMalote(txtNumMalote)
            End If
            Screen.MousePointer = vbDefault
        Else
            MsgBox "Número do Malote inválido.", _
                vbExclamation + vbOKOnly, App.Title
            cmdLimpar_Click
        End If
        KeyAscii = 0
    Else
        SoNumero KeyAscii
    End If
End Sub


