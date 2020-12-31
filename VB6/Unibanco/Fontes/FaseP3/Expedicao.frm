VERSION 5.00
Object = "{00080003-B1BA-11CE-ABC6-F5B2E79D9E3F}#8.0#0"; "LTOCX80N.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Expedicao 
   Caption         =   "Expedição de Documentos"
   ClientHeight    =   8148
   ClientLeft      =   372
   ClientTop       =   720
   ClientWidth     =   11628
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8148
   ScaleWidth      =   11628
   Begin VB.PictureBox picLilas2 
      Height          =   204
      Left            =   6960
      Picture         =   "Expedicao.frx":0000
      ScaleHeight     =   156
      ScaleWidth      =   240
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   2196
      Visible         =   0   'False
      Width           =   288
   End
   Begin VB.PictureBox picLilas 
      Height          =   228
      Left            =   6588
      Picture         =   "Expedicao.frx":0152
      ScaleHeight     =   180
      ScaleWidth      =   252
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2160
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picBranco2 
      Height          =   204
      Left            =   6252
      Picture         =   "Expedicao.frx":02A4
      ScaleHeight     =   156
      ScaleWidth      =   252
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2184
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picBranco 
      Height          =   216
      Left            =   5904
      Picture         =   "Expedicao.frx":03F6
      ScaleHeight     =   168
      ScaleWidth      =   240
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2184
      Visible         =   0   'False
      Width           =   288
   End
   Begin VB.PictureBox picAzul2 
      Height          =   192
      Left            =   8400
      Picture         =   "Expedicao.frx":0548
      ScaleHeight     =   144
      ScaleWidth      =   252
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1872
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picAzul 
      Height          =   216
      Left            =   8040
      Picture         =   "Expedicao.frx":069A
      ScaleHeight     =   168
      ScaleWidth      =   264
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   1860
      Visible         =   0   'False
      Width           =   312
   End
   Begin VB.PictureBox picAmarela2 
      Height          =   228
      Left            =   7656
      Picture         =   "Expedicao.frx":07EC
      ScaleHeight     =   180
      ScaleWidth      =   264
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1860
      Visible         =   0   'False
      Width           =   312
   End
   Begin VB.PictureBox picAmarela 
      Height          =   228
      Left            =   7320
      Picture         =   "Expedicao.frx":093E
      ScaleHeight     =   180
      ScaleWidth      =   240
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1848
      Visible         =   0   'False
      Width           =   288
   End
   Begin VB.PictureBox picVerde2 
      Height          =   228
      Left            =   6924
      Picture         =   "Expedicao.frx":0A90
      ScaleHeight     =   180
      ScaleWidth      =   240
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1872
      Visible         =   0   'False
      Width           =   288
   End
   Begin VB.PictureBox picVerde 
      Height          =   204
      Left            =   6588
      Picture         =   "Expedicao.frx":0BE2
      ScaleHeight     =   156
      ScaleWidth      =   240
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1872
      Visible         =   0   'False
      Width           =   288
   End
   Begin VB.PictureBox picVermelha2 
      Height          =   228
      Left            =   6264
      Picture         =   "Expedicao.frx":0D34
      ScaleHeight     =   180
      ScaleWidth      =   240
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1848
      Visible         =   0   'False
      Width           =   288
   End
   Begin VB.PictureBox picVermelha 
      Height          =   228
      Left            =   5940
      Picture         =   "Expedicao.frx":0E86
      ScaleHeight     =   180
      ScaleWidth      =   240
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1836
      Visible         =   0   'False
      Width           =   288
   End
   Begin MSFlexGridLib.MSFlexGrid grdDocto 
      Height          =   2136
      Left            =   48
      TabIndex        =   3
      Top             =   1320
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   3768
      _Version        =   393216
      Rows            =   9
      Cols            =   6
      FixedCols       =   0
      BackColorSel    =   8388608
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin VB.Frame Frame5 
      Height          =   4164
      Left            =   9804
      TabIndex        =   16
      Top             =   3936
      Width           =   1752
      Begin VB.CommandButton cmdAuditoria 
         Caption         =   "A&uditoria"
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
         Height          =   660
         Left            =   486
         Picture         =   "Expedicao.frx":0FD8
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   170
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
         Height          =   648
         Left            =   486
         Picture         =   "Expedicao.frx":1162
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   840
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
         Height          =   648
         Left            =   486
         Picture         =   "Expedicao.frx":146C
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1500
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
         Height          =   648
         Left            =   486
         Picture         =   "Expedicao.frx":1776
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2160
         Width           =   820
      End
      Begin VB.CommandButton cmdInverteCor 
         Caption         =   "Inverte cor"
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
         Height          =   648
         Left            =   486
         Picture         =   "Expedicao.frx":1A80
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2810
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
         Height          =   648
         Left            =   486
         Picture         =   "Expedicao.frx":1D8A
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3468
         Width           =   820
      End
   End
   Begin VB.Frame frmImagem 
      Caption         =   "Imagem"
      Height          =   4164
      Left            =   60
      TabIndex        =   14
      Top             =   3936
      Width           =   9672
      Begin LeadLib.Lead Lead1 
         Height          =   3816
         Left            =   120
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   240
         Width           =   9444
         _Version        =   524288
         _ExtentX        =   16658
         _ExtentY        =   6731
         _StockProps     =   229
         BackColor       =   -2147483639
         BorderStyle     =   1
         ScaleHeight     =   316
         ScaleWidth      =   785
         DataField       =   ""
         BitmapDataPath  =   ""
         AnnDataPath     =   ""
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1260
      Left            =   9804
      TabIndex        =   13
      Top             =   0
      Width           =   1752
      Begin VB.CommandButton CmdReimpressao 
         Caption         =   "&Reimprimir"
         Enabled         =   0   'False
         Height          =   324
         Left            =   132
         TabIndex        =   5
         Top             =   516
         Width           =   1464
      End
      Begin VB.CommandButton cmdLimpar 
         Caption         =   "&Limpar"
         Height          =   324
         Left            =   132
         TabIndex        =   4
         Top             =   168
         Width           =   1464
      End
      Begin VB.CommandButton cmdFechar 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   324
         Left            =   132
         TabIndex        =   6
         Top             =   864
         Width           =   1464
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1260
      Left            =   36
      TabIndex        =   30
      Top             =   0
      Width           =   9696
      Begin VB.PictureBox picNumMalote 
         Height          =   396
         Left            =   4944
         ScaleHeight     =   348
         ScaleWidth      =   2052
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   240
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
            TabIndex        =   38
            Top             =   36
            Width           =   1956
         End
      End
      Begin VB.PictureBox Picture3 
         Height          =   396
         Left            =   108
         ScaleHeight     =   348
         ScaleWidth      =   2052
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   240
         Width           =   2100
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
            Left            =   12
            TabIndex        =   36
            Top             =   12
            Width           =   1992
         End
      End
      Begin VB.ComboBox cmbCapa 
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
         Left            =   2280
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   2604
      End
      Begin VB.PictureBox Picture5 
         Height          =   396
         Left            =   108
         ScaleHeight     =   348
         ScaleWidth      =   2052
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   732
         Width           =   2100
         Begin VB.Label Label1 
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
            TabIndex        =   34
            Top             =   12
            Width           =   984
         End
      End
      Begin VB.ComboBox cmbAgencia 
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
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   732
         Width           =   2604
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
         Left            =   7116
         MaxLength       =   12
         TabIndex        =   1
         Text            =   "00446000028"
         Top             =   228
         Width           =   2196
      End
      Begin VB.Label lblLote 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0001"
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
         Left            =   7116
         TabIndex        =   32
         Top             =   732
         Width           =   2196
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
         Left            =   4944
         TabIndex        =   31
         Top             =   732
         Width           =   2100
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
      Height          =   288
      Left            =   60
      TabIndex        =   17
      Top             =   3624
      Width           =   11484
   End
End
Attribute VB_Name = "Expedicao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '       Variáveis de Cálculo do novo Super DV
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Ret         As Integer
    Dim Data        As String
    Dim nroCaixa    As String
    Dim Agencia     As String
    Dim TipoAgencia As String
    Dim Valor       As String
    Dim Operador    As String
    Dim SDV         As String * 1


Private Type tpMyDoc
    IdDocto                             As Long
    TipoDocto                           As Integer
    Ocorrencia                          As Long
    RetornoTransacao                    As Long
    Leitura                             As String
    Frente                              As String
    Verso                               As String
    Ordem                               As String * 1
    Status                              As String * 1
    Autenticado                         As String
    AutenticacaoGare                    As String * 64  'Autenticacao Digital (GARE)
    Nr_Autenticacoes_Efetuadas          As Integer 'Quantidade de autenticacoes efetuadas no documento
    Nr_Autenticacoes_Permitidas         As Integer
    NSU                                 As String
    Valor                               As Currency
    Terminal                            As Integer
    Vinculo                             As Long
    ParaAutenticar                      As Boolean
    Impresso                            As Boolean
    LI                                  As Boolean
End Type

Dim m_Gare As TpGare
Private Type TpGare
    DataVencto                          As String
    CodReceita                          As Integer
    InscrEstadual                       As String
    CPFCNPJ                             As String
    InscrDivida                         As String
    NumAIIM                             As String
    NomeAgenciaColeta                   As String
    AutenticacaoGare                    As String * 64
    NumeroViasGARE                      As Integer
    NSU                                 As String
    Terminal                            As Long
    VlrReceita                          As Double
    VlrJuros                            As Double
    VlrMulta                            As Double
    VlrAcrescimo                        As Double
    VlrHonorario                        As Double
    VlrTotal                            As Double
End Type


'* Guarda Dados de Cheque, pesquisa de Saldo Insuficiente *'
Private Type SaldoInsuficiente
    IdDocto                             As Long
    DataSaldo                           As String
    HoraSaldo                           As String
    Agencia                             As String
    Conta                               As String
    SaldoDisponivel                     As String
    LimiteChequeEspecial                As String
    ValorBloqueado                      As String
    ValorSaldoAtual                     As String
End Type

Private Const COL_DOCUMENTO = 0
Private Const COL_DESCRICAO = 1
Private Const COL_VALOR = 2
Private Const COL_TRATAMENTO = 3
Private Const COL_NR_AUTENTICA = 4
Private Const COL_IMAGEM = 5

Private UserReautentica                 As Boolean
Private m_FromEvent                     As Boolean
Private bRunActivate                    As Boolean
Private m_FirstActivate                 As Boolean
Private m_Autenticando                  As Boolean
Private m_IdCapa                        As Long
Private ControleSaldo                   As Long
Private m_Capa                          As String
Private m_IdEnvMal                      As String * 1
Private m_NumMalote                     As String
Private m_Agencia                       As String * 4
Private m_Status                        As String * 1
Private m_Ocorrencia                    As Long
Private m_Busy                          As Boolean
Private m_Doc                           As tpMyDoc
Private aDoc()                          As tpMyDoc
Private SaldoInsuficiente()             As SaldoInsuficiente
Private aIndice()                       As Integer
Private m_CountDocto                    As Integer

Private qryGetCapaExpedicao             As rdoQuery
Private qryGetMaloteExpedicao           As rdoQuery
Private qryGetDocumentoExpedicao        As rdoQuery
Private qryGetDocumentoOcorrencia       As rdoQuery
Private qryGetocorrencia                As rdoQuery
Private qryGetMotivoExclusao            As rdoQuery
Private qryGetDescricaoDocumento        As rdoQuery
Private qryGetAgContaDocumento          As rdoQuery
Private qryGetUsuario                   As rdoQuery
Private qryGetCartaoAvulso              As rdoQuery
Private qryGetBHVCDescricao             As rdoQuery
Private qryAtualizaAutenticacao         As rdoQuery
Private qryAtualizaStatusCapaDoctoExpedido  As rdoQuery
Private qryVerificaBinCartao            As rdoQuery
Private qryLerParametro                 As rdoQuery
Private qryGetSaldoConta                As rdoQuery

Private rsExpedicao                     As rdoResultset
Private rsDoc                           As rdoResultset
Private RsOcorrencia                    As rdoResultset
Private rsMotivo                        As rdoResultset
Private rsDescricao                     As rdoResultset
Private rsAgConta                       As rdoResultset
Private rsDocOcor                       As rdoResultset
Private rsUsuario                       As rdoResultset
Private RsBHVC                          As rdoResultset
Private rsParametro                     As rdoResultset
Private rsSaldoConta                    As rdoResultset

'Declaração da DLL calculo do NSU e Calculo do SDV
Private Declare Function QXCalNsu Lib "qxnsusdv32.dll" (ByVal PCC As String, ByVal Caixa As String, ByVal DataAutentic As String, ByVal Tipo As String, ByVal Valor As String, ByVal Ret As String) As Integer
Private Declare Function QXGetSDV Lib "qxnsusdv32.dll" (ByVal PCC As String, ByVal Caixa As String, ByVal DataAutentic As String, ByVal Tipo As String, ByVal Valor As String, ByVal CIF As String, ByVal Ret As String) As Integer

Private Function FormataString(ByVal pOque As Variant, _
                               ByVal pCompletarCom As Variant, _
                               ByVal pFieldLen As Integer, _
                               ByVal pAEsquerda As Boolean) As Variant
    
    
    If pFieldLen <= 0 Then FormataString = pOque: Exit Function
    If pCompletarCom = "" Then FormataString = pOque: Exit Function
    If pFieldLen < Len(pOque) Then FormataString = pOque: Exit Function

    If pAEsquerda Then
        FormataString = Right(String(pFieldLen - Len(pOque), pCompletarCom) & pOque, pFieldLen)
    Else
        FormataString = Left(pOque & String(pFieldLen - Len(pOque), pCompletarCom), pFieldLen)
    End If
    
End Function
Private Sub LimpaHeader()
    
    lblCapa.Caption = "Capa"
    lblLote.Caption = ""
    cmbCapa.Clear
    TxtNumMalote.Text = ""
    cmbAgencia.Clear
    lblOcorrencia.Caption = ""
    
    'Limpa variaveis com informações da capa
    m_IdCapa = 0
    m_Capa = ""
    m_NumMalote = ""
    m_Agencia = ""
    m_IdEnvMal = ""
    
    'Desabilita botões
    cmdAuditoria.Enabled = False
    cmdZoomMais.Enabled = False
    cmdZoomMenos.Enabled = False
    cmdRotacao.Enabled = False
    cmdInverteCor.Enabled = False
    cmdFrenteVerso.Enabled = False
    FrmImagem.Visible = False
    
End Sub
Private Function FinalizarExpedicao() As Boolean
    Dim bExiste     As Boolean
    Dim Count       As Integer
    Dim resp        As Integer
    
    If m_Status = "D" Or m_Status = "X" Then
        AtualizaStatusCapa m_IdCapa, m_Status
        m_IdCapa = 0
        Exit Function
    End If
    
    bExiste = False
    For Count = 1 To m_CountDocto
        If aDoc(Count).Nr_Autenticacoes_Permitidas > 0 And aDoc(Count).ParaAutenticar And (aDoc(Count).Autenticado = "0" Or (aDoc(Count).Autenticado < aDoc(Count).Nr_Autenticacoes_Permitidas)) Then
            bExiste = True
            Exit For
        End If
    Next
    
    'Testa a existência de expedições pendentes. Se houverem então solicita a confirmação de finalização
    If bExiste Then
        resp = MsgBox("O " & IIf(m_IdEnvMal = "E", "Envelope ", "Malote ") & _
            m_Capa & " da agência " & m_Agencia & _
            " continua pendente de Expedição porque " & _
            " ainda possui documentos para autenticar. Certeza da Saída ?", vbYesNo + vbOKOnly + vbDefaultButton2, App.Title)
            
        If resp = vbNo Then
           FinalizarExpedicao = False
           Exit Function
        Else

            AtualizaStatusCapa m_IdCapa, m_Status
            m_IdCapa = 0
        End If
    Else
        AtualizaStatusCapa m_IdCapa, "E"
        m_IdCapa = 0
    End If
    
    FinalizarExpedicao = True
    
End Function

' Chave = 1 - Capa; 2 - Num_Malote
Private Function LocalizarCapa(ByVal Chave As Integer) As Boolean
    
    On Error GoTo Erro_LocalizarCapa

    LocalizarCapa = False
    
    If Chave = 1 And cmbCapa.ListCount > 0 Then
        Exit Function
    End If
    
    If Chave = 1 Then
        If Not IsNumeric(cmbCapa.Text) Then
            cmbCapa.SelStart = 0
            cmbCapa.SelLength = Len(cmbCapa.Text)
            MsgBox "Capa informada não é válida.", vbExclamation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Function
        End If
        qryGetCapaExpedicao.rdoParameters(0) = Geral.DataProcessamento
        qryGetCapaExpedicao.rdoParameters(1) = Val(cmbCapa.Text)
        Set rsExpedicao = qryGetCapaExpedicao.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    Else
        If Not IsNumeric(TxtNumMalote.Text) Then
            TxtNumMalote.SelStart = 0
            TxtNumMalote.SelLength = Len(TxtNumMalote.Text)
            MsgBox "Número do Malote informado não é válido.", vbExclamation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Function
        End If
        qryGetMaloteExpedicao.rdoParameters(0) = Geral.DataProcessamento
        qryGetMaloteExpedicao.rdoParameters(1) = Val(TxtNumMalote.Text)
        Set rsExpedicao = qryGetMaloteExpedicao.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End If
    If rsExpedicao.EOF Then
        If Chave = 1 Then
            cmbCapa.SelStart = 0
            cmbCapa.SelLength = Len(cmbCapa.Text)
            MsgBox "Capa informada não foi encontrada.", vbExclamation + vbOKOnly, App.Title
            Limpar_grdDocto
            cmdLimpar_Click
        Else
            TxtNumMalote.SelStart = 0
            TxtNumMalote.SelLength = Len(TxtNumMalote.Text)
            MsgBox "Número do Malote informado não foi encontrado.", vbExclamation + vbOKOnly, App.Title
            Limpar_grdDocto
            cmdLimpar_Click
        End If
        Exit Function
    End If
    cmbCapa.Clear
    cmbAgencia.Clear
    While Not rsExpedicao.EOF
        
        If Chave = 1 Then
            cmbCapa.Text = Format(rsExpedicao!Capa, IIf(rsExpedicao!IdEnv_Mal = "E", "00000000", "00000000000000"))
            If rsExpedicao!IdEnv_Mal = "E" Then
                TxtNumMalote.Text = ""
            Else
                TxtNumMalote.Text = Format(rsExpedicao!Num_Malote, "00000000000")
            End If
            cmbAgencia.AddItem Format(rsExpedicao!AgOrig, "0000")
            cmbAgencia.ItemData(cmbAgencia.NewIndex) = rsExpedicao!IdCapa
        Else
            cmbCapa.AddItem Format(rsExpedicao!Capa, IIf(rsExpedicao!IdEnv_Mal = "E", "00000000", "00000000000000"))
        End If
        rsExpedicao.MoveNext
    Wend
    If Chave = 1 Then
        If cmbAgencia.ListCount = 1 Then
            cmbAgencia.ListIndex = 0
        Else
            cmbAgencia.SetFocus
            SendKeys "{F4}"
        End If
    Else
        If cmbCapa.ListCount = 1 Then
            cmbCapa.ListIndex = 0
        Else
            cmbCapa.SetFocus
            SendKeys "{F4}"
        End If
    End If
    
    LocalizarCapa = True
    Exit Function
Erro_LocalizarCapa:

    Select Case TratamentoErro("Não foi possível localizar a capa.", Err, rdoErrors)
        Case vbRetry
            Resume
    End Select

End Function

Private Sub ObtemDocumentos(ByVal IdCapa As Long)

    On Error GoTo ErroGetDocto
    rdoErrors.Clear

    Screen.MousePointer = vbHourglass

    Erase aDoc
    m_CountDocto = 0

    qryGetDocumentoExpedicao.rdoParameters(0) = Geral.DataProcessamento
    qryGetDocumentoExpedicao.rdoParameters(1) = IdCapa
    Set rsDoc = qryGetDocumentoExpedicao.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

    If Not rsDoc.EOF Then
        rsDoc.MoveLast
        rsDoc.MoveFirst
    End If

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Leitura da tabela de parametros sempre na leitura de uma nova capa'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    qryLerParametro.rdoParameters(0) = Geral.DataProcessamento
    Set rsParametro = qryLerParametro.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

    Set rsParametro = Nothing

    ReDim aDoc(rsDoc.RowCount)

    While Not rsDoc.EOF
        m_Doc.IdDocto = rsDoc!IdDocto
        m_Doc.TipoDocto = rsDoc!TipoDocto
        m_Doc.Ocorrencia = rsDoc!Ocorrencia
        m_Doc.RetornoTransacao = rsDoc!RetornoTransacao
        m_Doc.Leitura = Trim(rsDoc!Leitura)
        m_Doc.Frente = rsDoc!Frente
        m_Doc.Verso = rsDoc!Verso
        m_Doc.Ordem = rsDoc!Ordem
        m_Doc.Status = rsDoc!Status
        m_Doc.Autenticado = Trim(rsDoc!Autenticado)
        m_Doc.AutenticacaoGare = rsDoc!AutenticacaoGare & ""
        m_Doc.NSU = Trim(rsDoc!NSU)
        m_Doc.Valor = rsDoc!Valor
        m_Doc.Terminal = rsDoc!Terminal
        m_Doc.Vinculo = rsDoc!Vinculo
        m_Doc.Nr_Autenticacoes_Efetuadas = Val(rsDoc!Autenticado) 'Tabela Documento
        If m_Doc.TipoDocto = 18 Then
            'Para GARE com autenticacao Digital dispensa Autenticacao Manual
            If Len(Trim(m_Doc.AutenticacaoGare)) = 0 Then
                m_Doc.Nr_Autenticacoes_Permitidas = Val(rsDoc!Nr_Autenticacoes) 'Tabela Grupo Documento
            Else
                m_Doc.Nr_Autenticacoes_Permitidas = 0
            End If
        Else
            m_Doc.Nr_Autenticacoes_Permitidas = Val(rsDoc!Nr_Autenticacoes) 'Tabela Grupo Documento
        End If
        If m_Doc.Autenticado = "" Then
            m_Doc.Autenticado = "0"
        End If
        
        If m_Doc.Status <> "D" And _
           m_Doc.Status <> "F" And _
           m_Doc.Status <> "C" And _
           m_Doc.TipoDocto <> 2 And _
           m_Doc.TipoDocto <> 3 And _
           m_Doc.TipoDocto <> 4 And _
           m_Doc.TipoDocto <> 6 And _
           m_Doc.TipoDocto <> 7 And _
           m_Doc.TipoDocto <> 32 And _
           m_Doc.TipoDocto <> 33 And _
           m_Doc.TipoDocto <> 34 And _
           m_Doc.TipoDocto <> 36 And _
           m_Doc.TipoDocto <> 38 And _
           m_Doc.TipoDocto <> 41 And _
           m_Doc.TipoDocto <> 42 And _
           m_Doc.TipoDocto <> 43 And _
           m_Doc.TipoDocto <> 39 And _
           (m_Doc.Autenticado = "0" Or _
           (m_Doc.Autenticado < m_Doc.Nr_Autenticacoes_Permitidas)) Then
            m_Doc.ParaAutenticar = True
        Else
            m_Doc.ParaAutenticar = False
        End If
        m_Doc.Impresso = False

        If (m_Doc.TipoDocto <> 32 And _
            m_Doc.TipoDocto <> 33 And _
            m_Doc.TipoDocto <> 34 And _
            m_Doc.TipoDocto <> 38 And _
            m_Doc.TipoDocto <> 42 And _
            m_Doc.TipoDocto <> 43) Or _
           (m_Doc.Status <> "D" And _
            m_Doc.Status <> "F" And _
            m_Doc.Status <> "C") Then
            m_CountDocto = m_CountDocto + 1
            aDoc(m_CountDocto) = m_Doc
        End If
        
        rsDoc.MoveNext
    Wend
    rsDoc.Close
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    Exit Sub

ErroGetDocto:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro na obtenção de documentos para Expedição.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    m_Busy = False

End Sub

Private Function ObtemDescricaoDocto(ByVal TipoDocto As Integer) As String
    On Error GoTo ErroDescricao
    rdoErrors.Clear
    
    Screen.MousePointer = vbHourglass
    
    qryGetDescricaoDocumento.rdoParameters(0) = TipoDocto
    Set rsDescricao = qryGetDescricaoDocumento.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    
    If rsDescricao.EOF Then
        ObtemDescricaoDocto = "DOCUMENTO TIPO INVÁLIDO"
    Else
        ObtemDescricaoDocto = rsDescricao!Nome
    End If
    rsDescricao.Close
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    Exit Function
    
ErroDescricao:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro na obtenção da descricao do tipo do documento.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    m_Busy = False
    Unload Me
End Function

Private Function ObtemDescricaoBHVC(ByVal IdDocto As Long, ByVal TipoDocto As Integer) As String
    On Error GoTo ErroDescricao
    rdoErrors.Clear
    
    Screen.MousePointer = vbHourglass
    
    qryGetBHVCDescricao.rdoParameters(0) = Geral.DataProcessamento
    qryGetBHVCDescricao.rdoParameters(1) = IdDocto
    qryGetBHVCDescricao.rdoParameters(2) = TipoDocto
    Set RsBHVC = qryGetBHVCDescricao.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    
    If RsBHVC.EOF Then
        ObtemDescricaoBHVC = ""
    Else
        ObtemDescricaoBHVC = RsBHVC!BHVC_Descricao
    End If
    RsBHVC.Close
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    Exit Function
    
ErroDescricao:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro na obtenção da mensagem da Consulta do Titulo.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    m_Busy = False
    Unload Me
End Function

Private Function ObtemOcorrencia(ByVal Ocorrencia As Long, ByVal pRetornoTransacao As Long) As String

    Dim sDescricaoTransacao     As String

    On Error GoTo ErroOcorrencia
    rdoErrors.Clear

    If pRetornoTransacao > 0 Then
        If ObtemRetornoTransacao(pRetornoTransacao, sDescricaoTransacao) Then
            ObtemOcorrencia = sDescricaoTransacao
            Exit Function
        End If
    End If
    
    Screen.MousePointer = vbHourglass

    If Ocorrencia > 999 Then
        '5 posicoes
        qryGetocorrencia.rdoParameters(0) = Val(Left(Ocorrencia, 3))
    Else
        '3 posicoes
        qryGetocorrencia.rdoParameters(0) = Left(Val(Format(Ocorrencia, "00000")), 3)
    End If

    Ocorrencia = Val(Left(Ocorrencia, 3))

    Set RsOcorrencia = qryGetocorrencia.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    If RsOcorrencia.EOF Then
        ObtemOcorrencia = "Codigo da Ocorrencia nao existe: " & Trim(str(Ocorrencia))
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
Private Function ObtemMotivoExclusao(ByVal IdCapa As Long) As String
    On Error GoTo ErroMotivo
    rdoErrors.Clear
    
    Screen.MousePointer = vbHourglass
    
    qryGetMotivoExclusao.rdoParameters(0) = Geral.DataProcessamento
    qryGetMotivoExclusao.rdoParameters(1) = IdCapa
    Set rsMotivo = qryGetMotivoExclusao.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    If rsMotivo.EOF Then
        ObtemMotivoExclusao = "Erro. Motivo de exclusao nao encontrado."
    Else
        ObtemMotivoExclusao = rsMotivo!MotivoExclusao
    End If
    rsMotivo.Close
    Screen.MousePointer = vbDefault
    On Error GoTo 0
    Exit Function
    
ErroMotivo:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro na obtenção do Motivo de Exclusão do Envelope/Malote.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    m_Busy = False
    Unload Me

End Function

Private Function PossuiMotivoExclusao(ByVal IdCapa As Long) As Boolean
    On Error GoTo ErroMotivo
    rdoErrors.Clear
    
    Screen.MousePointer = vbHourglass
    
    qryGetMotivoExclusao.rdoParameters(0) = Geral.DataProcessamento
    qryGetMotivoExclusao.rdoParameters(1) = IdCapa
    Set rsMotivo = qryGetMotivoExclusao.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    If rsMotivo.EOF Then
        PossuiMotivoExclusao = False
    Else
        PossuiMotivoExclusao = True
    End If
    rsMotivo.Close
    Screen.MousePointer = vbDefault
    On Error GoTo 0
    Exit Function
    
ErroMotivo:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro na obtenção do Motivo de Exclusão do Envelope/Malote.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    m_Busy = False
    Unload Me

End Function

Private Function AtualizaStatusCapa(ByVal IdCapa As Long, ByVal Status As String) As Boolean
    
Dim Count As Long

On Error GoTo ErroAtualizaStatus

    rdoErrors.Clear
    
    AtualizaStatusCapa = False
    Screen.MousePointer = vbHourglass
    
    '--- Atualiza status da tabela CAPA e Docto para (E)xpedida ---
    With qryAtualizaStatusCapaDoctoExpedido
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = IdCapa
        .rdoParameters(3) = Status
        .Execute
        
        If .rdoParameters(0) <> 0 Then
            GoTo CancelaAtualizaStatus
        Else
            If Status = "E" Then
              'Gravar Log -> Atualizar Envelope / Malote Expedido
              Call GravaLog(IdCapa, 0, 86)
              'Grava log -> Documento autenticado (Cheques UBB para compensacao)
              For Count = 0 To m_CountDocto
                  If aDoc(Count).TipoDocto = 6 And InStr("409,230", Left(aDoc(Count).Leitura, 3)) Then
                       Call GravaLog(IdCapa, aDoc(Count).IdDocto, 80)
                  End If
             Next
            End If
        End If
    End With
    
    AtualizaStatusCapa = True
    Screen.MousePointer = vbDefault
    On Error GoTo 0
    Exit Function
    
CancelaAtualizaStatus:
    Screen.MousePointer = vbDefault
    MsgBox "Erro na atualização do status do envelope/malote.", vbCritical + vbOKOnly, App.Title
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

Private Function AtualizaAutenticacao(ByVal IdDocto As Long, ByVal Autenticacao As String) As Boolean

    On Error GoTo ErroAtualizaAut
    rdoErrors.Clear

    AtualizaAutenticacao = True
    Screen.MousePointer = vbHourglass

    With qryAtualizaAutenticacao
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = IdDocto
        .rdoParameters(3) = Autenticacao
        .Execute
        If .rdoParameters(0) <> 0 Then
            AtualizaAutenticacao = False
            Screen.MousePointer = vbDefault
            MsgBox "Erro na atualização da autenticação do documento.", vbCritical + vbOKOnly, App.Title
        Else
            If Val(Autenticacao) > 2 Then
                'Gravar Log -> Reautenticar Documento
                Call GravaLog(m_IdCapa, IdDocto, 81)
            Else
                'Gravar Log -> Autenticar Documento
                Call GravaLog(m_IdCapa, IdDocto, 80)
            End If
        End If
    End With
    Screen.MousePointer = vbDefault
    On Error GoTo 0
    Exit Function

ErroAtualizaAut:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro na atualização da autenticação do documento.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    m_Busy = False
    Unload Me
End Function

Private Function ObtemAgConta(ByVal IdDocto As Long, ByVal TipoDocto As Integer, _
                            Agencia As Integer, Conta As Long) As Boolean
    On Error GoTo ErroAgConta
    rdoErrors.Clear
    
    Screen.MousePointer = vbHourglass
    
    qryGetAgContaDocumento.rdoParameters(0) = Geral.DataProcessamento
    qryGetAgContaDocumento.rdoParameters(1) = IdDocto
    qryGetAgContaDocumento.rdoParameters(2) = TipoDocto
    Set rsAgConta = qryGetAgContaDocumento.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    If rsAgConta.EOF Then
        Screen.MousePointer = vbDefault
        MsgBox "Erro na obtenção da Agencia e Conta do documento.", vbExclamation + vbOKOnly, App.Title
        ObtemAgConta = False
    Else
        Agencia = rsAgConta!Agencia
        Conta = rsAgConta!Conta
        ObtemAgConta = True
    End If
    rsAgConta.Close
    Screen.MousePointer = vbDefault
    On Error GoTo 0
    Exit Function
    
ErroAgConta:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro na obtenção da Agencia e Conta do documento.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    m_Busy = False
    Unload Me
                            
End Function

Private Function ObtemBancoCheque(ByVal IdDocto As Long, _
                                        Banco As String, _
                                        Cheque As String) As Boolean
    Dim Count       As Integer
    Dim bExiste     As Boolean
    
    bExiste = False
    Count = 1
    Do While Count <= m_CountDocto
        If aDoc(Count).IdDocto = IdDocto Then
            bExiste = True
            Exit Do
        End If
        Count = Count + 1
    Loop
    If Not bExiste Or Len(aDoc(Count).Leitura) = 0 Then
        ObtemBancoCheque = False
    Else
        ObtemBancoCheque = True
        Banco = Left(aDoc(Count).Leitura, 3)
        Cheque = Mid(aDoc(Count).Leitura, 12, 6)
    End If
End Function
Private Sub selecionaLinha(ByVal pLinha As Long)
    GrdDocto.Row = pLinha
    GrdDocto.Col = 0
    GrdDocto.ColSel = GrdDocto.Cols - 1

End Sub
Private Function UsuarioSupervisor(ByVal User As String) As Boolean
                            
    UsuarioSupervisor = GrupoUsuario(User, eG_SUPERVISOR)
                            
End Function
Private Function PossuiADCC(ByVal Vinculo As Long) As Boolean
    Dim Count As Integer
    
    For Count = 1 To m_CountDocto
        If aDoc(Count).Status <> "D" And aDoc(Count).Status <> "F" And aDoc(Count).Status <> "C" And _
           aDoc(Count).Vinculo = Vinculo And aDoc(Count).TipoDocto = 4 Then
            PossuiADCC = True
            Exit Function
        End If
    Next
    PossuiADCC = False
End Function

Private Function MostraOcorrencia(ByVal Linha As Integer) As String
    Dim Count As Integer
    
    Count = aIndice(Linha)
    
    If (aDoc(Count).Status = "D" Or aDoc(Count).Status = "F" Or aDoc(Count).Status = "C") And _
        aDoc(Count).TipoDocto <> 32 And aDoc(Count).TipoDocto <> 33 And _
        aDoc(Count).TipoDocto <> 34 And aDoc(Count).TipoDocto <> 38 And _
        aDoc(Count).TipoDocto <> 42 And aDoc(Count).TipoDocto <> 43 Then
        MostraOcorrencia = "Ocorrencia: " & ObtemOcorrencia(aDoc(Count).Ocorrencia, aDoc(Count).RetornoTransacao)
    Else
        MostraOcorrencia = ""
    End If
End Function

Private Function AjusteDeposito(ByVal Vinculo As Long) As Boolean
    Dim Count As Integer
    
    AjusteDeposito = False
    For Count = 1 To m_CountDocto
        If aDoc(Count).Status <> "D" And aDoc(Count).Status <> "F" And aDoc(Count).Status <> "C" And _
           aDoc(Count).Vinculo = Vinculo And (aDoc(Count).TipoDocto = 2 Or _
           aDoc(Count).TipoDocto = 3) Then
            AjusteDeposito = True
            Exit Function
        End If
    Next
End Function

Private Sub MostraImagem(ByVal Linha As Integer)
    
    Dim i       As Integer
    Dim Ret     As Integer
    
    i = aIndice(Linha)
    
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
         .Load Geral.DiretorioImagens & Format(rsExpedicao!IdLote, "000000000") & "\" & aDoc(i).Frente, 0, 0, 1
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
    DoEvents
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
    GrdDocto.SetFocus
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
    GrdDocto.SetFocus

End Sub

Private Sub Preenche_grdDocto()
    Dim Count               As Integer
    Dim Linha               As Integer
    Dim strOcor             As String
    Dim strBco              As String
    Dim strCh               As String
    Dim Ag                  As Integer
    Dim Cta                 As Long
    Dim Descr               As String
    Dim lvinculo            As Long
    Dim iIndex              As Integer
    Dim Vinculo             As Integer
    Dim Aviso               As Integer
    Dim bComoFinalizado     As Boolean
    Dim Exibiu              As Boolean
    
    Dim Vinculos()          As Long
    
    Dim NumLancInt          As Integer
    Dim i                   As Integer
    
    GrdDocto.Clear
    Erase aIndice
    ReDim aIndice(m_CountDocto)
    
    bRunActivate = True
    Linha = 1
    bComoFinalizado = False
    
    NumLancInt = 0
    
    Exibiu = False
    
    For Count = 0 To m_CountDocto
        
        If (aDoc(Count).TipoDocto = 2 Or aDoc(Count).TipoDocto = 3 Or aDoc(Count).TipoDocto = 37) And (aDoc(Count).Status <> "C" Or aDoc(Count).Status <> "D" Or aDoc(Count).Status <> "F") And (Len(Trim(aDoc(Count).NSU)) = 0 Or IsNull(aDoc(Count).NSU)) Then
            If Not Exibiu Then
               Aviso = MsgBox("Esta capa possui documentos não processados, enviar o Físico para Supervisor")
               Exibiu = True
            End If
        End If
        
        If aDoc(Count).TipoDocto = 2 Or _
           aDoc(Count).TipoDocto = 3 Or _
           aDoc(Count).TipoDocto = 4 Or _
           aDoc(Count).TipoDocto = 5 Or _
           aDoc(Count).TipoDocto = 6 Or _
           aDoc(Count).TipoDocto = 7 Or _
           aDoc(Count).TipoDocto = 32 Or _
           aDoc(Count).TipoDocto = 33 Or _
           aDoc(Count).TipoDocto = 34 Or _
           aDoc(Count).TipoDocto = 37 Or _
           aDoc(Count).TipoDocto = 38 Then
           
            ObtemAgConta aDoc(Count).IdDocto, aDoc(Count).TipoDocto, Ag, Cta
            
            If aDoc(Count).TipoDocto = 5 Or _
               aDoc(Count).TipoDocto = 6 Or _
               aDoc(Count).TipoDocto = 7 Then
               
                ObtemBancoCheque aDoc(Count).IdDocto, strBco, strCh
                Descr = "Bco: " & strBco & " Ag: " & Format(Ag, "0000") & _
                    " Cta: " & FormataConta(Cta) & " Ch: " & strCh
            Else
                Descr = Space(18) & "Ag: " & Format(Ag, "0000") & _
                    " Cta: " & FormataConta(Cta)
            End If
            
            ''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Percorre os documentos do mesmo vinculo, se houver'
            'lancamento interno ao invez de "Para compensacao" '
            'ser "Finalizado"
            ''''''''''''''''''''''''''''''''''''''''''''''''''''
            If aDoc(Count).TipoDocto = 2 Or aDoc(Count).TipoDocto = 3 Then
                lvinculo = aDoc(Count).Vinculo

                iIndex = 1
                Do While iIndex <= UBound(aDoc)
                
                    If (aDoc(iIndex).Vinculo = lvinculo) Or (aDoc(iIndex).TipoDocto = 0) Then
                        If (aDoc(iIndex).TipoDocto = 41 Or aDoc(iIndex).TipoDocto = 5) Then
                            bComoFinalizado = True
                            Exit Do
                        End If
                    End If
                    
                    iIndex = iIndex + 1
                Loop

            End If
            
        Else
            Descr = " "
        End If
    
        GrdDocto.Rows = Linha + 1
        If Count = 0 Then
            GrdDocto.TextMatrix(0, COL_DOCUMENTO) = "Documento"
            GrdDocto.TextMatrix(0, COL_DESCRICAO) = "Descrição"
            GrdDocto.TextMatrix(0, COL_VALOR) = "Valor"
            GrdDocto.TextMatrix(0, COL_TRATAMENTO) = "Tratamento"
        ' se ocorrencias
        ElseIf (aDoc(Count).Status = "C" Or aDoc(Count).Status = "D" Or aDoc(Count).Status = "F") And _
           aDoc(Count).TipoDocto <> 32 And aDoc(Count).TipoDocto <> 33 And _
           aDoc(Count).TipoDocto <> 34 And aDoc(Count).TipoDocto <> 38 And _
           aDoc(Count).TipoDocto <> 42 And aDoc(Count).TipoDocto <> 43 Then
            GrdDocto.Row = Linha
            GrdDocto.Col = COL_IMAGEM
            Set GrdDocto.CellPicture = picLilas.Picture
            
            GrdDocto.TextMatrix(Linha, COL_DOCUMENTO) = ObtemDescricaoDocto(aDoc(Count).TipoDocto)
            GrdDocto.TextMatrix(Linha, COL_DESCRICAO) = Descr
            GrdDocto.Col = COL_VALOR
            GrdDocto.CellAlignment = flexAlignRightCenter
            GrdDocto.TextMatrix(Linha, COL_VALOR) = Trim(FormataValor(aDoc(Count).Valor, 22))
            strOcor = Format(aDoc(Count).Ocorrencia, "000")
            GrdDocto.TextMatrix(Linha, COL_TRATAMENTO) = "Ocorrência: " & strOcor
            aIndice(Linha) = Count
            Linha = Linha + 1
        
        ' se ajustes
        ElseIf (aDoc(Count).TipoDocto = 32 Or aDoc(Count).TipoDocto = 33 Or _
            aDoc(Count).TipoDocto = 34 Or aDoc(Count).TipoDocto = 38 Or _
            aDoc(Count).TipoDocto = 42 Or aDoc(Count).TipoDocto = 43) And _
           (aDoc(Count).Status <> "C" And aDoc(Count).Status <> "D" And aDoc(Count).Status <> "F") Then
            GrdDocto.Row = Linha
            GrdDocto.Col = COL_IMAGEM
            Set GrdDocto.CellPicture = picVerde.Picture
            GrdDocto.TextMatrix(Linha, COL_DOCUMENTO) = ObtemDescricaoDocto(aDoc(Count).TipoDocto)
            GrdDocto.TextMatrix(Linha, COL_DESCRICAO) = Descr
            GrdDocto.Col = COL_VALOR
            GrdDocto.CellAlignment = flexAlignRightCenter
            GrdDocto.TextMatrix(Linha, COL_VALOR) = Trim(FormataValor(aDoc(Count).Valor, 22))
            aIndice(Linha) = Count
            Linha = Linha + 1
        
        ' se para compensacao
        ElseIf (aDoc(Count).TipoDocto = 2 Or aDoc(Count).TipoDocto = 3 Or _
           aDoc(Count).TipoDocto = 7 Or aDoc(Count).TipoDocto = 39) And aDoc(Count).Status <> "F" Then
            GrdDocto.Row = Linha
            GrdDocto.Col = COL_IMAGEM
            
            If bComoFinalizado Then
                Set GrdDocto.CellPicture = picBranco.Picture
            Else
                Set GrdDocto.CellPicture = picVermelha.Picture
            End If
            
            GrdDocto.TextMatrix(Linha, COL_DOCUMENTO) = ObtemDescricaoDocto(aDoc(Count).TipoDocto)
            GrdDocto.TextMatrix(Linha, COL_DESCRICAO) = Descr
            GrdDocto.Col = COL_VALOR
            GrdDocto.CellAlignment = flexAlignRightCenter
            GrdDocto.TextMatrix(Linha, COL_VALOR) = Trim(FormataValor(aDoc(Count).Valor, 22))
            
            If bComoFinalizado Then
                GrdDocto.TextMatrix(Linha, COL_TRATAMENTO) = " Finalizado "
                bComoFinalizado = False
            Else
                GrdDocto.TextMatrix(Linha, COL_TRATAMENTO) = " Para Compensação "
            End If
            
            aIndice(Linha) = Count
            Linha = Linha + 1
           
        ' Cheque Pagamento para compensacao
        ElseIf aDoc(Count).TipoDocto = 6 Then
            GrdDocto.Row = Linha
            GrdDocto.Col = COL_IMAGEM
            Set GrdDocto.CellPicture = picAmarela.Picture
            
            GrdDocto.TextMatrix(Linha, COL_DOCUMENTO) = ObtemDescricaoDocto(aDoc(Count).TipoDocto)
            GrdDocto.TextMatrix(Linha, COL_DESCRICAO) = Descr
            GrdDocto.Col = COL_VALOR
            GrdDocto.CellAlignment = flexAlignRightCenter
            GrdDocto.TextMatrix(Linha, COL_VALOR) = Trim(FormataValor(aDoc(Count).Valor, 22))
            GrdDocto.Col = COL_TRATAMENTO
            GrdDocto.TextMatrix(Linha, COL_TRATAMENTO) = " Para Compensação "
            aIndice(Linha) = Count
            Linha = Linha + 1
           
        ' se finalizado
        ElseIf (aDoc(Count).TipoDocto = 4 Or aDoc(Count).TipoDocto = 36 Or aDoc(Count).TipoDocto = 41 Or _
            (aDoc(Count).TipoDocto = 5 And aDoc(Count).Nr_Autenticacoes_Permitidas = 0)) Then
               
            'Realizou Lançamento Interno Qualquer Documento Que Esteja
            'Nesta Capa Deve Ser Pintado de Branco
            If aDoc(Count).TipoDocto = 41 Then
               NumLancInt = NumLancInt + 1
               ReDim Preserve Vinculos(NumLancInt) As Long
               Vinculos(NumLancInt) = aDoc(Count).Vinculo
            End If
            
            GrdDocto.Row = Linha
            GrdDocto.Col = COL_IMAGEM
            Set GrdDocto.CellPicture = picBranco.Picture
            GrdDocto.TextMatrix(Linha, COL_DOCUMENTO) = ObtemDescricaoDocto(aDoc(Count).TipoDocto)
            GrdDocto.TextMatrix(Linha, COL_DESCRICAO) = Descr
            GrdDocto.Col = COL_VALOR
            GrdDocto.CellAlignment = flexAlignRightCenter
            GrdDocto.TextMatrix(Linha, COL_VALOR) = Trim(FormataValor(aDoc(Count).Valor, 22))
            GrdDocto.Col = COL_TRATAMENTO
            GrdDocto.TextMatrix(Linha, COL_TRATAMENTO) = " Finalizado "
            aIndice(Linha) = Count
            Linha = Linha + 1
        
        ' para autenticacao
        ElseIf aDoc(Count).Status <> "C" And aDoc(Count).Status <> "D" And aDoc(Count).Status <> "F" Then
            GrdDocto.Row = Linha
            GrdDocto.Col = COL_IMAGEM
            If aDoc(Count).TipoDocto = 27 Or aDoc(Count).TipoDocto = 12 Then
                Set GrdDocto.CellPicture = picAzul.Picture
            Else
                Set GrdDocto.CellPicture = picBranco.Picture
            End If
            GrdDocto.TextMatrix(Linha, COL_DOCUMENTO) = ObtemDescricaoDocto(aDoc(Count).TipoDocto)
            GrdDocto.TextMatrix(Linha, COL_DESCRICAO) = Descr
            GrdDocto.Col = COL_VALOR
            GrdDocto.CellAlignment = flexAlignRightCenter
            GrdDocto.TextMatrix(Linha, COL_VALOR) = Trim(FormataValor(aDoc(Count).Valor, 22))
            
            If Val(aDoc(Count).Autenticado) < aDoc(Count).Nr_Autenticacoes_Permitidas Then
                GrdDocto.TextMatrix(Linha, COL_TRATAMENTO) = " Para Autenticação"
            ElseIf Val(aDoc(Count).Autenticado) = aDoc(Count).Nr_Autenticacoes_Permitidas Then
                GrdDocto.TextMatrix(Linha, COL_TRATAMENTO) = " Autenticado: " & aDoc(Count).NSU
            Else
                GrdDocto.TextMatrix(Linha, COL_TRATAMENTO) = " Re-autenticado: " & aDoc(Count).NSU
            End If
            
            GrdDocto.TextMatrix(Linha, COL_NR_AUTENTICA) = IIf(Val(aDoc(Count).Nr_Autenticacoes_Permitidas) - Val(aDoc(Count).Nr_Autenticacoes_Efetuadas) < 0, 0, Val(aDoc(Count).Nr_Autenticacoes_Permitidas) - Val(aDoc(Count).Nr_Autenticacoes_Efetuadas))
            aIndice(Linha) = Count
            Linha = Linha + 1
        End If
    Next
    
    For i = 1 To NumLancInt
    
        Preenche_grdDocto_LancamentosInternos Vinculos(i)
        
    Next
    
    selecionaLinha 1
    bRunActivate = False
        
End Sub
Private Sub Preenche_grdDocto_LancamentosInternos(ByVal Incidencia As Long)

Dim Contador As Integer
Dim Line     As Integer

Dim ContDoc  As Integer

ContDoc = 0

Line = 0

For Contador = 0 To m_CountDocto
        
    If Line < m_CountDocto And aDoc(Contador).TipoDocto <> 7 And aDoc(Contador).TipoDocto <> 6 And aDoc(Contador).TipoDocto <> 43 And aDoc(Contador).TipoDocto <> 12 And aDoc(Contador).TipoDocto <> 27 And aDoc(Contador).Vinculo = Incidencia Then
       GrdDocto.Row = Line
       GrdDocto.Col = COL_IMAGEM
       Set GrdDocto.CellPicture = picBranco.Picture
       aDoc(Contador).LI = True
    Else
       aDoc(Contador).LI = False
    End If
        
    Line = Line + 1
        
Next

End Sub
Sub Preenche_grdDocto_LancamentosInternosSelecao(ByValIncidencia As Integer)

End Sub

Private Sub SelecionaPasta(ByVal Linha As Integer, _
                           ByVal PastaAberta As Boolean)
    Dim Count   As Integer
    
    Dim Vinculo As Long
    
    Count = aIndice(Linha)
    GrdDocto.Col = COL_IMAGEM
    
    ' se ocorrencias
    If ((aDoc(Count).Status = "C" Or aDoc(Count).Status = "D" Or aDoc(Count).Status = "F") And _
       aDoc(Count).TipoDocto <> 32 And aDoc(Count).TipoDocto <> 33 And _
       aDoc(Count).TipoDocto <> 34 And aDoc(Count).TipoDocto <> 38 And _
       aDoc(Count).TipoDocto <> 42 And aDoc(Count).TipoDocto <> 43) Then
        If PastaAberta Then
            Set GrdDocto.CellPicture = picLilas2.Picture
        Else
            Set GrdDocto.CellPicture = picLilas.Picture
        End If
    
    ' se ajustes
    ElseIf (aDoc(Count).TipoDocto = 32 Or aDoc(Count).TipoDocto = 33 Or _
       aDoc(Count).TipoDocto = 34 Or aDoc(Count).TipoDocto = 38 Or _
       aDoc(Count).TipoDocto = 42 Or aDoc(Count).TipoDocto = 43 And _
       (aDoc(Count).Status <> "C" And aDoc(Count).Status <> "D" And aDoc(Count).Status <> "F")) Then
        If PastaAberta Then
            Set GrdDocto.CellPicture = picVerde2.Picture
        Else
            Set GrdDocto.CellPicture = picVerde.Picture
        End If
    
    ' se para compensacao
    ElseIf (aDoc(Count).TipoDocto = 2 Or aDoc(Count).TipoDocto = 3 Or _
       aDoc(Count).TipoDocto = 7 Or aDoc(Count).TipoDocto = 39) Then
            If PastaAberta Then
                If Not aDoc(Count).LI Then
                    If Trim(GrdDocto.TextMatrix(Count, COL_TRATAMENTO)) = "Finalizado" Then
                        Set GrdDocto.CellPicture = picBranco2.Picture
                    Else
                        Set GrdDocto.CellPicture = picVermelha2.Picture
                    End If
                Else
                   Set GrdDocto.CellPicture = picBranco2.Picture
                End If
            Else
                If Not aDoc(Count).LI Then
                    If Trim(GrdDocto.TextMatrix(Count, COL_TRATAMENTO)) = "Finalizado" Then
                        Set GrdDocto.CellPicture = picBranco.Picture
                    Else
                        Set GrdDocto.CellPicture = picVermelha.Picture
                    End If
                Else
                   Set GrdDocto.CellPicture = picBranco.Picture
                End If
            End If
       
    ' ch. pagamento para compensacao
    ElseIf aDoc(Count).TipoDocto = 6 Then
        If PastaAberta Then
           Set GrdDocto.CellPicture = picAmarela2.Picture
        Else
           Set GrdDocto.CellPicture = picAmarela.Picture
        End If
       

    ' se finalizado
    ElseIf aDoc(Count).TipoDocto = 4 Or aDoc(Count).TipoDocto = 5 Or aDoc(Count).TipoDocto = 36 Or _
           aDoc(Count).TipoDocto = 41 Then
           Vinculo = aDoc(Count).Vinculo
        If PastaAberta Then
            If aDoc(Count).LI Then
               Set GrdDocto.CellPicture = picBranco2.Picture
            End If
        Else
            If aDoc(Count).LI Then
               Set GrdDocto.CellPicture = picBranco.Picture
            End If
        End If
    
    ' para autenticacao
    ElseIf aDoc(Count).Status <> "C" And aDoc(Count).Status <> "D" And aDoc(Count).Status <> "F" Then
        If (aDoc(Count).TipoDocto = 27 Or aDoc(Count).TipoDocto = 12) Then
            If PastaAberta Then
                Set GrdDocto.CellPicture = picAzul2.Picture
            Else
                Set GrdDocto.CellPicture = picAzul.Picture
            End If
        Else
            If PastaAberta Then
                Set GrdDocto.CellPicture = picBranco2.Picture
            Else
                Set GrdDocto.CellPicture = picBranco.Picture
            End If
        End If
    End If
    
    GrdDocto.Col = 0

End Sub

Private Function PosicionaAutenticar() As Boolean
    Dim Count As Integer
    
    PosicionaAutenticar = False
    For Count = 1 To GrdDocto.Rows - 1
        If Trim(GrdDocto.TextMatrix(Count, COL_TRATAMENTO)) = "Para Autenticação" Then
            If Count = GrdDocto.Row Then
                GrdDocto.Row = 2
            End If
            GrdDocto.TopRow = Count
            GrdDocto.Row = Count
            GrdDocto.Col = 0
            GrdDocto.ColSel = 4
            PosicionaAutenticar = True
            Exit Function
        End If
    Next
End Function

Private Function ImprimeHeaderDeposito() As Boolean
    Dim ret_imp, ret_aut    As Integer
    Dim Buff_st             As String * 3
    Dim buff_aut            As String * 45
    Dim buff_linha          As String * 2
    Dim r1, r2, r3          As Integer
     
    ImprimeHeaderDeposito = False
    
    buff_linha = Chr(13)
    
    If rsExpedicao!IdEnv_Mal = "E" Then
       buff_aut = Space(12) & "Envelope - " & Format(rsExpedicao!Capa, "00000000")
    Else
       buff_aut = Space(8) & "Numero Malote - " & Format(rsExpedicao!Num_Malote, "00000000000")
    End If
    
    If Geral.autenticadora = 2 Then
        ret_imp = Autentica.Status(Buff_st)
        r1 = Hex(Asc(Mid(Buff_st, 1, 1)))
        r2 = Hex(Asc(Mid(Buff_st, 2, 1)))
        r3 = Hex(Asc(Mid(Buff_st, 3, 1)))
        If (r1) <> 0 Then
            ret_aut = 1
        Else
            ret_aut = Autentica.Imprimir(buff_aut, False)
        End If
    Else
        ret_aut = Autentica.Imprimir(buff_aut, False)
    End If
    
    If (ret_aut <> 0) Then
        ret_imp = Autentica.Status(Buff_st)
    
        If (ret_imp = 0) Then
           
           r1 = Hex(Asc(Mid(Buff_st, 1, 1)))
           r2 = Hex(Asc(Mid(Buff_st, 2, 1)))
           r3 = Hex(Asc(Mid(Buff_st, 3, 1)))
           
           If (r1 + r2 + r3) <> 0 Then
              'teste do 1 byte
              Select Case r1
                 Case 1
                    MsgBox "Autenticadora está off-line.", vbInformation + vbOKOnly, App.Title
                 Case 2
                    MsgBox "Autenticadora está desligada.", vbInformation + vbOKOnly, App.Title
                 Case 3
                    MsgBox "Autenticadora está com buffer cheio.", vbInformation + vbOKOnly, App.Title
                 Case 4
                    MsgBox "Autenticadora está inoperante.", vbInformation + vbOKOnly, App.Title
              End Select
           
              'teste do 3 byte
              If (r3 <> 0) Then
                 MsgBox "Verifique a bobina da Autentica.", vbInformation + vbOKOnly, App.Title
              End If
              Exit Function
           End If
        Else
           MsgBox "Verifique a Autentica.", vbInformation + vbOKOnly, App.Title
           Exit Function
        End If
    Else
        buff_aut = String(45, "=")
        ret_aut = Autentica.Imprimir(buff_aut, False)
        
        If rsExpedicao!IdEnv_Mal = "M" Then
            buff_aut = Space(4) & "Capa Malote Empresa - " & Format(rsExpedicao!Capa, "00000000000000")
            ret_aut = Autentica.Imprimir(buff_aut, False)
            ret_aut = Autentica.Imprimir(buff_linha, False)
        
            buff_aut = String(45, "-")
            ret_aut = Autentica.Imprimir(buff_aut, False)
        End If
        
        buff_aut = Space(13) & "UBB - Unibanco SA"
        ret_aut = Autentica.Imprimir(buff_aut, False)
    
    End If
    
    ImprimeHeaderDeposito = True
    
End Function

Private Function ImprimeHeaderOcorrencia() As Boolean
    Dim ret_imp, ret_aut        As Integer
    Dim Buff_st                 As String * 3
    Dim buff_aut                As String * 45
    Dim buff_linha              As String * 2
    Dim r1, r2, r3              As Integer
     
    ImprimeHeaderOcorrencia = False
    
    buff_linha = Chr(13)
    
    buff_aut = Space(16) & "OCORRENCIAS"
    
    If Geral.autenticadora = 2 Then
        ret_imp = Autentica.Status(Buff_st)
        r1 = Hex(Asc(Mid(Buff_st, 1, 1)))
        r2 = Hex(Asc(Mid(Buff_st, 2, 1)))
        r3 = Hex(Asc(Mid(Buff_st, 3, 1)))
        If (r1) <> 0 Then
            ret_aut = 1
        Else
            ret_aut = Autentica.Imprimir(buff_aut, False)
        End If
    Else
        ret_aut = Autentica.Imprimir(buff_aut, False)
    End If
    
    If (ret_aut <> 0) Then
        ret_imp = Autentica.Status(Buff_st)
    
        If (ret_imp = 0) Then
           
           r1 = Hex(Asc(Mid(Buff_st, 1, 1)))
           r2 = Hex(Asc(Mid(Buff_st, 2, 1)))
           r3 = Hex(Asc(Mid(Buff_st, 3, 1)))
           
           If (r1 + r2 + r3) <> 0 Then
              'teste do 1 byte
              Select Case r1
                 Case 1
                    MsgBox "Autenticadora está off-line.", vbInformation + vbOKOnly, App.Title
                 Case 2
                    MsgBox "Autenticadora está desligada.", vbInformation + vbOKOnly, App.Title
                 Case 3
                    MsgBox "Autenticadora está com buffer cheio.", vbInformation + vbOKOnly, App.Title
                 Case 4
                    MsgBox "Autenticadora está inoperante.", vbInformation + vbOKOnly, App.Title
              End Select
           
              'teste do 3 byte
              If (r3 <> 0) Then
                 MsgBox "Verifique a bobina da Autentica.", vbInformation + vbOKOnly, App.Title
              End If
              Exit Function
           End If
        Else
           MsgBox "Verifique a Autentica.", vbInformation + vbOKOnly, App.Title
           Exit Function
        End If
    Else
        ret_aut = Autentica.Imprimir(buff_linha, False)
        
        If rsExpedicao!IdEnv_Mal = "E" Then
           buff_aut = Space(12) & "Envelope - " & Format(rsExpedicao!Capa, "00000000")
        Else
           buff_aut = Space(8) & "Numero Malote - " & Format(rsExpedicao!Num_Malote, "00000000000")
        End If
        ret_aut = Autentica.Imprimir(buff_aut, False)
        ret_aut = Autentica.Imprimir(buff_linha, False)
        
        buff_aut = String(45, "=")
        ret_aut = Autentica.Imprimir(buff_aut, False)
        ret_aut = Autentica.Imprimir(buff_linha, False)
        
        If rsExpedicao!IdEnv_Mal = "M" Then
            buff_aut = Space(4) & "Capa Malote Empresa - " & Format(rsExpedicao!Capa, "00000000000000")
            ret_aut = Autentica.Imprimir(buff_aut, False)
            ret_aut = Autentica.Imprimir(buff_linha, False)
        End If
        
        buff_aut = Space(7) & "Data do Movimento: " + Mid(Geral.DataProcessamento, 7, 2) + "/" + Mid(Geral.DataProcessamento, 5, 2) + "/" + Mid(Geral.DataProcessamento, 1, 4)
        ret_aut = Autentica.Imprimir(buff_aut, False)
        
        buff_aut = Space(4) & "Data de Emissão: " & Format(Now, "dd/mm/yyyy hh:mm:ss")
        ret_aut = Autentica.Imprimir(buff_aut, False)
        
        buff_aut = "  Ag. Coleta:" + Format(rsExpedicao!AgOrig, "0000") & " - Ag. Processadora: " + Geral.AgenciaCentral
        ret_aut = Autentica.Imprimir(buff_aut, False)
        ret_aut = Autentica.Imprimir(buff_linha, False)
        
        buff_aut = String(45, "-")
        ret_aut = Autentica.Imprimir(buff_aut, False)
        ret_aut = Autentica.Imprimir(buff_linha, False)
        
    End If
    
    ImprimeHeaderOcorrencia = True
    
End Function
Private Function ImprimeHeaderAutenticaGARE() As Boolean

    Dim ret_imp, ret_aut        As Integer
    Dim Buff_st                 As String * 3
    Dim buff_aut                As String * 45
    Dim buff_linha              As String * 2
    Dim r1, r2, r3              As Integer
     
    ImprimeHeaderAutenticaGARE = False
    
    buff_linha = Chr(13)
    
    buff_aut = String(45, "-")

ContinuaImpressao:

    If Geral.autenticadora = 2 Then
        ret_imp = Autentica.Status(Buff_st)
        r1 = Hex(Asc(Mid(Buff_st, 1, 1)))
        r2 = Hex(Asc(Mid(Buff_st, 2, 1)))
        r3 = Hex(Asc(Mid(Buff_st, 3, 1)))
        If (r1) <> 0 Then
            ret_aut = 1
        Else

            ret_aut = Autentica.Imprimir(buff_aut, False)
        End If
    Else
        ret_aut = Autentica.Imprimir(buff_aut, False)
    End If
    
    If (ret_aut <> 0) Then

        ret_imp = Autentica.Status(Buff_st)
    
        If (ret_imp = 0) Then
           
           r1 = Hex(Asc(Mid(Buff_st, 1, 1)))
           r2 = Hex(Asc(Mid(Buff_st, 2, 1)))
           r3 = Hex(Asc(Mid(Buff_st, 3, 1)))
           
           If (r1 + r2 + r3) <> 0 Then
              'teste do 1 byte
              Select Case r1
                 Case 1
                    MsgBox "Autenticadora está off-line.", vbInformation + vbOKOnly, App.Title
                 Case 2
                    MsgBox "Autenticadora está desligada.", vbInformation + vbOKOnly, App.Title
                 Case 3
                    MsgBox "Autenticadora está com buffer cheio.", vbInformation + vbOKOnly, App.Title
                    ret_imp = Autentica.Inicia()    'Continua impressao
                    Call Espera(10)                 'Espera desvaziar o buffer
                    GoTo ContinuaImpressao
                 Case 4
                    MsgBox "Autenticadora está inoperante.", vbInformation + vbOKOnly, App.Title
              End Select
           
              'teste do 3 byte
              If (r3 <> 0) Then
                If MsgBox("Verifique a bobina da Autentica." + vbCrLf + vbCrLf + "Continuar impressão ?", vbInformation + vbYesNo, App.Title) = vbNo Then
                    Exit Function
                Else
                    ret_imp = Autentica.Inicia()    'Continua impressao
                    Call Espera(10)                 'Espera desvaziar o buffer
                    GoTo ContinuaImpressao
                End If
              End If
           End If
        Else
           MsgBox "Verifique a Autentica.", vbInformation + vbOKOnly, App.Title
           Exit Function
        End If
    Else
        ret_aut = Autentica.Imprimir(buff_linha, False)
        
        buff_aut = Space(13) & "Unibanco - Banco 409"
        ret_aut = Autentica.Imprimir(buff_aut, False)
        
        buff_aut = Space(9) & "Demonstrativo de Pagamento"
        ret_aut = Autentica.Imprimir(buff_aut, False)
        
        buff_aut = "Guia de Arrecadacao Estadual - Demais Receita"
        ret_aut = Autentica.Imprimir(buff_aut, False)
        
        buff_aut = Space(18) & "GARE - DR"
        ret_aut = Autentica.Imprimir(buff_aut, False)
        
        ret_aut = Autentica.Imprimir(buff_linha, False)
        
        buff_aut = String(45, "-")
        ret_aut = Autentica.Imprimir(buff_aut, False)
        
    End If
    
    ImprimeHeaderAutenticaGARE = True

End Function
Private Function ImprimeAutenticacao(ByVal Linha As Integer) As Boolean
    Dim ret_imp, ret_aut    As Integer
    Dim Buff_st             As String * 3
    Dim buff_aut            As String * 60
    Dim r1, r2, r3          As Integer
    Dim Count               As Integer
    Dim strValor            As String
    Dim Ulogo, Blogo, LogoUBB, Space, GraphOn, GraphOff As String
    Dim Autenticado         As String * 1
    Dim bPosicionou         As Boolean
    Dim sOldAutenticado     As String
    
     
    If m_Autenticando Then Exit Function

    'Verifica se grid está vazio
    If GrdDocto.TextMatrix(1, 1) = "" Then Exit Function

    ' Se Campo NSU for 0, então não pode autenticar
    If Val(aDoc(Linha).NSU) = 0 Then

       MsgBox "NSU zerado ou está faltando. Impossível Autenticar", vbExclamation
       
       GrdDocto.SetFocus
       Exit Function
       
    End If
    
    ImprimeAutenticacao = False
    
    
    If Trim(GrdDocto.TextMatrix(Linha, COL_TRATAMENTO)) <> "Para Autenticação" And _
       Left(Trim(GrdDocto.TextMatrix(Linha, COL_TRATAMENTO)), 12) <> "Autenticado:" And _
       Left(Trim(GrdDocto.TextMatrix(Linha, COL_TRATAMENTO)), 15) <> "Re-autenticado:" Then
        Exit Function
    End If
    
    If Geral.autenticadora = 0 Then
        Exit Function
    End If
    
    If InStr(1, UCase(Trim(GrdDocto.TextMatrix(Linha, COL_TRATAMENTO))), "AUTENTICADO", vbTextCompare) <> 0 And Not UserReautentica Then
       
        MsgBox "Documento já está autenticado." & vbCrLf & _
               "Somente o Usuário Autorizado poderá re-autenticar este documento", _
               vbInformation + vbOKOnly, App.Title
        GrdDocto.SetFocus
        Exit Function
    
    End If
    
    Count = aIndice(Linha)
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Verificação da quantidade de autenticações conforme parametros'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If (aDoc(Count).Nr_Autenticacoes_Efetuadas >= aDoc(Count).Nr_Autenticacoes_Permitidas) And _
       (aDoc(Count).Nr_Autenticacoes_Permitidas <> 0) And _
       (Not UserReautentica) And _
       (Geral.Backup = False) Then

        If aDoc(Count).Nr_Autenticacoes_Efetuadas = 1 Then
            MsgBox "Já foi efetuada uma autenticação para este documento.", vbExclamation
        Else
            MsgBox "Já foram efetuadas duas autenticações para este documento.", vbExclamation
        End If
        GrdDocto.SetFocus
        Exit Function
    End If

    If aDoc(Count).TipoDocto = 0 Then
        MsgBox "Não é possível autenticar documento indefinido.", vbInformation + vbOKOnly, App.Title
        Exit Function
    ElseIf aDoc(Count).Valor <= 0 Then
        MsgBox "Não é possível autenticar documento sem valor.", vbInformation + vbOKOnly, App.Title
        Exit Function
    ElseIf aDoc(Count).NSU = "" Or Val(aDoc(Count).NSU) = 0 Then
        MsgBox "Não é possível autenticar documento sem número de NSU.", vbInformation + vbOKOnly, App.Title
        Exit Function
    End If

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '           Cálculo do novo Super DV
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Agencia = m_Agencia
    nroCaixa = Format(aDoc(Count).Terminal, "000")
    Data = Right(Geral.DataProcessamento, 2) & Mid(Geral.DataProcessamento, 5, 2) & Mid(Geral.DataProcessamento, 3, 2)
    TipoAgencia = "9"
    Valor = (aDoc(Count).Valor * 100)
    Operador = "999999"
    
    'Chama a DLL para o cálculo do Super DV
    Ret = QXGetSDV(Agencia, nroCaixa, Data, TipoAgencia, Valor, Operador, SDV)

    'Verifica se não houve erro e se os DVs estão OK
    If Ret <> 0 Then   'Or UCase(SDV) <> UCase(Mid(ControleBanco, 1, 1)) Then Exit Function
        MsgBox "Não é possível autenticar documento, cálculo do Super DV errado.", vbInformation + vbOKOnly, App.Title
        Exit Function
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''

    m_Autenticando = True

    strValor = Trim(FormataValor(aDoc(Count).Valor, 14))
    Select Case aDoc(Count).TipoDocto
        Case 27     'arrec.convenc.
            'A2_OK-227
            buff_aut = Mid(Geral.DataProcessamento, 7, 2) & _
                       Mid(Geral.DataProcessamento, 5, 2) & _
                       Mid(Geral.DataProcessamento, 3, 2) & _
                       SDV & _
                       String(14 - Len(strValor), "*") & strValor & _
                       "R999999#" & Format(aDoc(Count).NSU, "000000") & _
                       "9" & m_Agencia & _
                       Format(aDoc(Count).Terminal, "000") & _
                       "A" & Chr(10)

        Case 12     'tit.terc.sem cb.
            'A2_OK-228
            buff_aut = Mid(Geral.DataProcessamento, 7, 2) & _
                       Mid(Geral.DataProcessamento, 5, 2) & _
                       Mid(Geral.DataProcessamento, 3, 2) & _
                       SDV & _
                       String(14 - Len(strValor), "*") & strValor & _
                       "R999999#" & Format(aDoc(Count).NSU, "000000") & _
                       "9" & m_Agencia & _
                       Format(aDoc(Count).Terminal, "000") & _
                       "T" & Chr(10)

        Case Else
            'A2_OK-229
            buff_aut = Mid(Geral.DataProcessamento, 7, 2) & _
                       Mid(Geral.DataProcessamento, 5, 2) & _
                       Mid(Geral.DataProcessamento, 3, 2) & _
                       SDV & _
                       String(14 - Len(strValor), "*") & strValor & _
                       "R999999#" & Format(aDoc(Count).NSU, "000000") & _
                       "9" & m_Agencia & _
                       Format(aDoc(Count).Terminal, "000") & _
                       Chr(10)

    End Select
    
    If Geral.autenticadora = 1 Then

        'Montando graficamente a letra "U" de UBB
        Ulogo = Chr(&HFC) & Chr(&H0) & Chr(&HFE) & Chr(&H0) & Chr(&H2) & Chr(&H0) & Chr(&H2) & Chr(&H0) & Chr(&HFE) & Chr(&H0)
        'Montando graficamente a letra "B" de UBB
        Blogo = Chr(&HFE) & Chr(&H0) & Chr(&HFE) & Chr(&H0) & Chr(&H92) & Chr(&H0) & Chr(&H72) & Chr(&H0) & Chr(&HC) & Chr(&H0)
        'Montando graficamente um micro espaco
        Space = Chr(&H0)
        'Montando graficamente a string que poe a impressora em modo grafico
        GraphOn = Chr(&H1B) & Chr(&H4B) & Chr(&H22) & Chr(&H0)
        'Montando graficamente a string que tira a impressora do modo grafico
        GraphOff = Chr(&H0)
        'Concatenando todas as variaveis
        LogoUBB = GraphOn & Ulogo & Space & Blogo & Space & Blogo & Space & GraphOff & buff_aut

        
        'Funçao que imprime no Documento
        ret_aut = Autentica.Imprimir(LogoUBB, True)
    ElseIf Geral.autenticadora = 2 Then
        ret_imp = Autentica.Status(Buff_st)
        r1 = Hex(Asc(Mid(Buff_st, 1, 1)))
        r2 = Hex(Asc(Mid(Buff_st, 2, 1)))
        r3 = Hex(Asc(Mid(Buff_st, 3, 1)))
        If (r1) <> 0 Then
            ret_aut = 1
        Else
            ret_aut = Autentica.Imprimir(buff_aut, True)
        End If
    End If
    
    '''''''''''''''''''''''''''''''''''''''''
    'Se ret_aut <> 0, não conseguiu imprimir'
    '''''''''''''''''''''''''''''''''''''''''
    If (ret_aut <> 0) Then
        
        ret_imp = Autentica.Status(Buff_st)
        If (ret_imp = 0) Then
           ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
           'Mesmo com erro na autenticadora, pode acontecer de não'
           'retornar nada nos bits                                '
           ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
           r1 = Hex(Asc(Mid(Buff_st, 1, 1)))
           r2 = Hex(Asc(Mid(Buff_st, 2, 1)))
           r3 = Hex(Asc(Mid(Buff_st, 3, 1)))
           
           If (r1 + r2 + r3) <> 0 Then
              'teste do 1 byte
              Select Case r1
                 Case 1
                    MsgBox "Autenticadora está off-line.", vbInformation + vbOKOnly, App.Title
                 Case 2
                    MsgBox "Autenticadora está desligada.", vbInformation + vbOKOnly, App.Title
                 Case 3
                    MsgBox "Autenticadora está com buffer cheio.", vbInformation + vbOKOnly, App.Title
                 Case 4
                    MsgBox "Autenticadora está inoperante.", vbInformation + vbOKOnly, App.Title
              End Select
              
              If Geral.autenticadora = 1 And r2 <> 0 Then
                 MsgBox "Verifique a Autentica.", vbInformation + vbOKOnly, App.Title
              End If
           
              'teste do 3 byte
              If (r3 <> 0) Then
                 MsgBox "Verifique a bobina da Autentica.", vbInformation + vbOKOnly, App.Title
              End If
           Else
              MsgBox "Falha na Autentica.", vbCritical
           End If
           m_Autenticando = False
           Exit Function
        Else
           MsgBox "Verifique a Autentica.", vbInformation + vbOKOnly, App.Title
           m_Autenticando = False
           Exit Function
        End If
    End If
    
    sOldAutenticado = Autenticado
    
    If aDoc(Count).Autenticado = "0" Then
        Autenticado = "1"
    ElseIf aDoc(Count).Autenticado = "1" Then
        Autenticado = "2"
    Else
        Autenticado = "3"
    End If
    
    If Not AtualizaAutenticacao(aDoc(Count).IdDocto, Autenticado) Then
        m_Autenticando = False
        Autenticado = sOldAutenticado
        Exit Function
    End If
        
    aDoc(Count).Autenticado = Autenticado
    
    If Val(aDoc(Count).Autenticado) < aDoc(Count).Nr_Autenticacoes_Permitidas Then
        GrdDocto.TextMatrix(Linha, COL_TRATAMENTO) = " Para Autenticação"
    ElseIf Val(aDoc(Count).Autenticado) = aDoc(Count).Nr_Autenticacoes_Permitidas Then
        GrdDocto.TextMatrix(Linha, COL_TRATAMENTO) = " Autenticado: " & aDoc(Count).NSU
    Else
        GrdDocto.TextMatrix(Linha, COL_TRATAMENTO) = " Re-autenticado: " & aDoc(Count).NSU
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''
    'Incrementa o Nr de Autenticacoes efetuadas'
    ''''''''''''''''''''''''''''''''''''''''''''
    aDoc(Count).Nr_Autenticacoes_Efetuadas = aDoc(Count).Nr_Autenticacoes_Efetuadas + 1
    
    '''''''''''''''''''''''''''''''
    'Atualiza o Grid de Documentos'
    '''''''''''''''''''''''''''''''
    GrdDocto.TextMatrix(Linha, COL_NR_AUTENTICA) = _
        IIf(aDoc(Count).Nr_Autenticacoes_Permitidas - aDoc(Count).Nr_Autenticacoes_Efetuadas < 0, _
                        0, _
                        aDoc(Count).Nr_Autenticacoes_Permitidas - aDoc(Count).Nr_Autenticacoes_Efetuadas)
    
    
    bPosicionou = PosicionaAutenticar
    
    m_Autenticando = False
    
End Function

Private Function ImprimeTrailler(ByVal ShowMsg As Boolean) As Boolean
    Dim ret_imp, ret_aut        As Integer
    Dim Buff_st                 As String * 3
    Dim buff_aut                As String * 45
    Dim buff_linha              As String * 2
    Dim r1, r2, r3, i           As Integer
     
    ImprimeTrailler = False
    
    buff_linha = Chr(13)
    
    buff_aut = String(45, "-")
    
    ret_aut = Autentica.Imprimir(buff_aut, False)
    
    If (ret_aut <> 0) Then
        ret_imp = Autentica.Status(Buff_st)
    
        If (ret_imp = 0) Then
           
           r1 = Hex(Asc(Mid(Buff_st, 1, 1)))
           r2 = Hex(Asc(Mid(Buff_st, 2, 1)))
           r3 = Hex(Asc(Mid(Buff_st, 3, 1)))
           
           If (r1 + r2 + r3) <> 0 Then
              'teste do 1 byte
              Select Case r1
                 Case 1
                    MsgBox "Autenticadora está off-line.", vbInformation + vbOKOnly, App.Title
                 Case 2
                    MsgBox "Autenticadora está desligada.", vbInformation + vbOKOnly, App.Title
                 Case 3
                    MsgBox "Autenticadora está com buffer cheio.", vbInformation + vbOKOnly, App.Title
                 Case 4
                    MsgBox "Autenticadora está inoperante.", vbInformation + vbOKOnly, App.Title
              End Select
           
              'teste do 3 byte
              If (r3 <> 0) Then
                 MsgBox "Verifique a bobina da Autentica.", vbInformation + vbOKOnly, App.Title
              End If
              Exit Function
           End If
        Else
           MsgBox "Verifique a Autentica.", vbInformation + vbOKOnly, App.Title
           Exit Function
        End If
    Else
        If ShowMsg Then
            buff_aut = Space(9) & "Ticket de Caixa Unibanco."
            ret_aut = Autentica.Imprimir(buff_aut, False)
            
            buff_aut = Space(3) & "Feito para facilitar o seu dia-a-dia."
            ret_aut = Autentica.Imprimir(buff_aut, False)
        End If
        
        'Imprime 10 linhas no final da impressão do ticket
        For i = 1 To 10
            ret_aut = Autentica.Imprimir(buff_linha, False)
        Next i
        
    End If
    
    ImprimeTrailler = True
    
End Function

Private Sub ImprimeOcorrenciaCapa()
    Dim ret_aut         As Integer
    Dim StrMotivo       As String
    Dim buff_aut        As String * 45
    Dim buff_linha      As String * 2
    Dim Pos             As Integer

    If Geral.autenticadora = 0 Then
        Exit Sub
    End If

    buff_linha = Chr(13)

    If rsExpedicao!Status = "D" Then
        StrMotivo = Trim(ObtemMotivoExclusao(rsExpedicao!IdCapa))
    ElseIf rsExpedicao!Status = "X" Then
        StrMotivo = Trim(ObtemOcorrencia(rsExpedicao!Ocorrencia, 0)) 'rsExpedicao!RetornoTransacao))
    End If

    If Not ImprimeHeaderOcorrencia Then
        Exit Sub
    End If

    If rsExpedicao!IdEnv_Mal = "E" Then
        buff_aut = "Envelope Devolvido"
    Else
        buff_aut = "Malote Devolvido"
    End If
    ret_aut = Autentica.Imprimir(buff_aut, False)
    ret_aut = Autentica.Imprimir(buff_linha, False)
    
    buff_aut = "Motivo devolucao: "
    ret_aut = Autentica.Imprimir(buff_aut, False)
    
    If Len(StrMotivo) < 45 Then
        buff_aut = StrMotivo
        ret_aut = Autentica.Imprimir(buff_aut, False)
    Else
        Pos = 45
        While Pos < Len(StrMotivo)
            buff_aut = QuebraBuffer(StrMotivo, Pos)
            ret_aut = Autentica.Imprimir(buff_aut, False)
            StrMotivo = Right(StrMotivo, Len(StrMotivo) - Pos)
        Wend
        If Len(StrMotivo) > 0 Then
            buff_aut = StrMotivo
            ret_aut = Autentica.Imprimir(buff_aut, False)
        End If
    End If
    ret_aut = Autentica.Imprimir(buff_linha, False)

    'Gravar Log -> Imprimir Ocorrencia Capa
    Call GravaLog(m_IdCapa, 0, 82)

    ImprimeTrailler (False)
    
End Sub

Private Sub ImprimeOcorrenciaDoctos()

    Dim ret_aut             As Integer
    Dim StrMotivo           As String
    Dim buff_aut            As String * 45
    Dim buff_linha          As String * 2
    Dim Pos                 As Integer
    Dim Ocorrencia          As Long
    Dim strBco              As String
    Dim strCh               As String
    Dim Ag                  As Integer
    Dim Cta                 As Long
    Dim lRetornoTransacao   As Long
    Dim PrimeiroDocto       As Boolean
    
    Dim bImprimirBHVC       As Boolean
    Dim sDescricaoBHVC      As String

    If Geral.autenticadora = 0 Then
        Exit Sub
    End If

    buff_linha = Chr(13)
    Ocorrencia = 0

    On Error GoTo ErroDocOcor

    qryGetDocumentoOcorrencia.rdoParameters(0) = Geral.DataProcessamento
    qryGetDocumentoOcorrencia.rdoParameters(1) = rsExpedicao!IdCapa
    Set rsDocOcor = qryGetDocumentoOcorrencia.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    If Not rsDocOcor.EOF Then
        If Not ImprimeHeaderOcorrencia Then
            Exit Sub
        End If
    
        If Not rsDocOcor.EOF Then
            ReDim SaldoInsuficiente(rsDocOcor.RowCount)
        End If
        
        bImprimirBHVC = False
        
        '''''''''''''''''''''''''
        'Gambiarra é assim mesmo'
        '''''''''''''''''''''''''
        If Not rsDocOcor.EOF Then
            ''''''''''''''''''''''''''''
            'pega do primeiro documento'
            ''''''''''''''''''''''''''''
            'lRetornoTransacao = IIf(IsNull(rsDocOcor!RetornoTransacao), 0, rsDocOcor!RetornoTransacao)
'            Ocorrencia = rsDocOcor!Ocorrencia
 '           PrimeiroDocto = False
        End If
        
        While Not rsDocOcor.EOF
            '''''''''''''''''''''''''
            'acertar ObtemOcorrencia'
            '''''''''''''''''''''''''
            lRetornoTransacao = IIf(IsNull(rsDocOcor!RetornoTransacao), 0, rsDocOcor!RetornoTransacao)
            
            If Ocorrencia = 0 Then
                Ocorrencia = rsDocOcor!Ocorrencia
            ElseIf (Ocorrencia <> rsDocOcor!Ocorrencia) Or (lRetornoTransacao <> rsDocOcor!RetornoTransacao) Then
            
                ''''''''''''''''''''''''''''''''''''''''''''''
                'Infelizmente, está querendo imprimir dados  '
                'do primeiro registro sendo que se está no   '
                'segundo registro, portanto volto um registro'
                'para pegar o dado correto                   '
                ''''''''''''''''''''''''''''''''''''''''''''''
                rsDocOcor.MovePrevious
                If Not rsDocOcor.BOF Then
                    lRetornoTransacao = IIf(IsNull(rsDocOcor!RetornoTransacao), 0, rsDocOcor!RetornoTransacao)
                    StrMotivo = Trim(ObtemOcorrencia(Ocorrencia, lRetornoTransacao))
                    rsDocOcor.MoveNext
                Else
                    rsDocOcor.MoveNext
                    StrMotivo = Trim(ObtemOcorrencia(Ocorrencia, lRetornoTransacao))
                End If
                
                lRetornoTransacao = IIf(IsNull(rsDocOcor!RetornoTransacao), 0, rsDocOcor!RetornoTransacao)
                
                buff_aut = "Descricao da Ocorrencia: "
                ret_aut = Autentica.Imprimir(buff_aut, False)
                
                If Len(StrMotivo) < 45 Then
                    buff_aut = StrMotivo
                    ret_aut = Autentica.Imprimir(buff_aut, False)
                Else
                    Pos = 45
                    While Pos < Len(StrMotivo)
                        buff_aut = QuebraBuffer(StrMotivo, Pos)
                        ret_aut = Autentica.Imprimir(buff_aut, False)
                        StrMotivo = Right(StrMotivo, Len(StrMotivo) - Pos)
                    Wend
                    If Len(StrMotivo) > 0 Then
                        buff_aut = StrMotivo
                        ret_aut = Autentica.Imprimir(buff_aut, False)
                    End If
                End If
                ret_aut = Autentica.Imprimir(buff_linha, False)
                '''''''''''''''''''''''''''
                'Imprime descricao do BHVC'
                '''''''''''''''''''''''''''
                If bImprimirBHVC Then
                    bImprimirBHVC = False
                    If (sDescricaoBHVC <> "") Then
                        If Len(sDescricaoBHVC) < 45 Then
                            buff_aut = sDescricaoBHVC
                            ret_aut = Autentica.Imprimir(buff_aut, False)
                        Else
                            Pos = 45
                            While Pos < Len(sDescricaoBHVC)
                                buff_aut = QuebraBuffer(sDescricaoBHVC, Pos)
                                ret_aut = Autentica.Imprimir(buff_aut, False)
                                sDescricaoBHVC = Right(sDescricaoBHVC, Len(sDescricaoBHVC) - Pos)
                            Wend
                            If Len(sDescricaoBHVC) > 0 Then
                                buff_aut = sDescricaoBHVC
                                ret_aut = Autentica.Imprimir(buff_aut, False)
                            End If
                        End If
                    End If
                End If
                '''''''''''''''''''''''
                'Fim impressão do BHVC'
                '''''''''''''''''''''''
                buff_aut = String(45, "-")
                ret_aut = Autentica.Imprimir(buff_aut, False)
                ret_aut = Autentica.Imprimir(buff_linha, False)

                Ocorrencia = rsDocOcor!Ocorrencia
            End If

            buff_aut = Trim(ObtemDescricaoDocto(rsDocOcor!TipoDocto))
            ret_aut = Autentica.Imprimir(buff_aut, False)

            If rsDocOcor!TipoDocto = 2 Or _
               rsDocOcor!TipoDocto = 3 Or _
               rsDocOcor!TipoDocto = 4 Or _
               rsDocOcor!TipoDocto = 5 Or _
               rsDocOcor!TipoDocto = 6 Or _
               rsDocOcor!TipoDocto = 7 Or _
               rsDocOcor!TipoDocto = 32 Or _
               rsDocOcor!TipoDocto = 33 Or _
               rsDocOcor!TipoDocto = 34 Or _
               rsDocOcor!TipoDocto = 37 Or _
               rsDocOcor!TipoDocto = 38 Then

                ObtemAgConta rsDocOcor!IdDocto, rsDocOcor!TipoDocto, Ag, Cta

                If (rsDocOcor!TipoDocto = 4 And rsDocOcor!Ocorrencia = 429) Or _
                   (rsDocOcor!TipoDocto = 5 And rsDocOcor!Ocorrencia = 213) Then
                
                    '* Verifica se Conta Possui Saldo Insuficiente *'
                    VerificaSaldoCheque rsDocOcor!IdDocto, rsDocOcor!TipoDocto
                    
                    ObtemBancoCheque rsDocOcor!IdDocto, strBco, strCh
                    
                    buff_aut = "Bco: " & strBco & " Ag: " & Format(Ag, "0000") & _
                        " Cta: " & FormataConta(Cta) & " Ch: " & strCh
                Else
                    buff_aut = "Ag: " & Format(Ag, "0000") & _
                        " Cta: " & FormataConta(Cta)
                End If
            Else
                buff_aut = ""
            End If

            If Len(Trim(buff_aut)) > 0 Then
                ret_aut = Autentica.Imprimir(buff_aut, False)
            End If

            buff_aut = "Valor: " & FormataValor(rsDocOcor!Valor, 22)
            ret_aut = Autentica.Imprimir(buff_aut, False)

            If Ocorrencia = 208 And (rsDocOcor!TipoDocto = 13 Or _
               rsDocOcor!TipoDocto = 14 Or rsDocOcor!TipoDocto = 28 Or _
               rsDocOcor!TipoDocto = 29 Or rsDocOcor!TipoDocto = 30 Or _
               rsDocOcor!TipoDocto = 31) Then
               
                sDescricaoBHVC = ObtemDescricaoBHVC(rsDocOcor!IdDocto, rsDocOcor!TipoDocto)
                bImprimirBHVC = True
            End If
            
            ret_aut = Autentica.Imprimir(buff_linha, False)

            'Gravar Log -> Imprimir Ocorrencia Docto
            Call GravaLog(rsExpedicao!IdCapa, rsDocOcor!IdDocto, 83)

            rsDocOcor.MoveNext
        Wend

        buff_aut = "Descricao da Ocorrencia: "
        ret_aut = Autentica.Imprimir(buff_aut, False)
        
        StrMotivo = Trim(ObtemOcorrencia(Ocorrencia, lRetornoTransacao))
        If Len(StrMotivo) < 45 Then
            buff_aut = StrMotivo
            ret_aut = Autentica.Imprimir(buff_aut, False)
        Else
            Pos = 45
            While Pos < Len(StrMotivo)
                buff_aut = QuebraBuffer(StrMotivo, Pos)
                ret_aut = Autentica.Imprimir(buff_aut, False)
                StrMotivo = Right(StrMotivo, Len(StrMotivo) - Pos)
            Wend
            If Len(StrMotivo) > 0 Then
                buff_aut = StrMotivo
                ret_aut = Autentica.Imprimir(buff_aut, False)
            End If
        End If
        ret_aut = Autentica.Imprimir(buff_linha, False)
        
        If bImprimirBHVC Then
            If (sDescricaoBHVC <> "") Then
                If Len(sDescricaoBHVC) < 45 Then
                    buff_aut = sDescricaoBHVC
                    ret_aut = Autentica.Imprimir(buff_aut, False)
                Else
                    Pos = 45
                    While Pos < Len(sDescricaoBHVC)
                        buff_aut = QuebraBuffer(sDescricaoBHVC, Pos)
                        ret_aut = Autentica.Imprimir(buff_aut, False)
                        sDescricaoBHVC = Right(sDescricaoBHVC, Len(sDescricaoBHVC) - Pos)
                    Wend
                    If Len(sDescricaoBHVC) > 0 Then
                        buff_aut = sDescricaoBHVC
                        ret_aut = Autentica.Imprimir(buff_aut, False)
                    End If
                End If
            End If
        End If

        If Not CBool(ControleSaldo) Then ImprimeTrailler (False)
    End If
    rsDocOcor.Close
    
    If CBool(ControleSaldo) Then ImprimeSaldoConta
    
    ReDim SaldoInsuficiente(0)
    ControleSaldo = 0
    
    Exit Sub
ErroDocOcor:

    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro na obtenção de documentos com ocorrência para Expedição.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    m_Busy = False
End Sub

Private Sub ImprimeAjustes()
    Dim ret_aut         As Integer
    Dim buff_aut        As String * 45
    Dim buff_linha      As String * 2
    Dim Count           As Integer
    Dim Ag              As Integer
    Dim Cta             As Long
    Dim strValor        As String
    Dim Ocorrencia      As Long
    Dim StrMotivo       As String
    Dim Pos             As Integer

    If Geral.autenticadora = 0 Then
        Exit Sub
    End If

    buff_linha = Chr(13)

    For Count = 1 To m_CountDocto
        If (aDoc(Count).TipoDocto = 32 Or aDoc(Count).TipoDocto = 33 Or _
            aDoc(Count).TipoDocto = 34 Or aDoc(Count).TipoDocto = 38) And _
           (aDoc(Count).Status <> "C" And aDoc(Count).Status <> "D" And aDoc(Count).Status <> "F") Then

            If Not ImprimeHeaderDeposito Then
                Exit Sub
            End If

            If Not ObtemAgConta(aDoc(Count).IdDocto, aDoc(Count).TipoDocto, Ag, Cta) Then
                cmdLimpar_Click
                Exit Sub
            End If

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '           Cálculo do novo Super DV
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Agencia = m_Agencia
            nroCaixa = Format(aDoc(Count).Terminal, "000")
            Data = Right(Geral.DataProcessamento, 2) & Mid(Geral.DataProcessamento, 5, 2) & Mid(Geral.DataProcessamento, 3, 2)
            TipoAgencia = "9"
            Valor = (aDoc(Count).Valor * 100)
            Operador = "999999"
            
            'Chama a DLL para o cálculo do Super DV
            Ret = QXGetSDV(Agencia, nroCaixa, Data, TipoAgencia, Valor, Operador, SDV)
        
            'Verifica se não houve erro e se os DVs estão OK
            If Ret <> 0 Then   'Or UCase(SDV) <> UCase(Mid(ControleBanco, 1, 1)) Then Exit Function
                MsgBox "Não é possível autenticar documento, cálculo do Super DV errado.", vbInformation + vbOKOnly, App.Title
                Exit Sub
            End If
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''

            If aDoc(Count).TipoDocto = 32 Or aDoc(Count).TipoDocto = 34 Then
                If aDoc(Count).TipoDocto = 32 Then
                    If AjusteDeposito(aDoc(Count).Vinculo) Then
                        buff_aut = Space(13) & "Aviso de Credito"
                    Else
                        buff_aut = Space(10) & "Aviso de Credito - OCT"
                    End If
                Else
                    buff_aut = Space(13) & "Aviso de Credito"
                End If
                If aDoc(Count).TipoDocto = 32 Then
                    Ocorrencia = 12200
                Else
                    If PossuiADCC(aDoc(Count).Vinculo) Then
                        Ocorrencia = 40300
                    Else
                        Ocorrencia = 20400
                    End If
                End If
            Else
                If aDoc(Count).TipoDocto = 33 Then
                    If AjusteDeposito(aDoc(Count).Vinculo) Then
                        buff_aut = Space(14) & "Aviso de Debito"
                    Else
                        buff_aut = Space(10) & "Aviso de Debito - OCT"
                    End If
                Else
                    buff_aut = Space(14) & "Aviso de Debito"
                End If
                If aDoc(Count).TipoDocto = 33 Then
                    Ocorrencia = 12100
                Else
                    If PossuiADCC(aDoc(Count).Vinculo) Then
                        Ocorrencia = 40200
                    Else
                        Ocorrencia = 20300
                    End If
                End If
            End If
            ret_aut = Autentica.Imprimir(buff_aut, False)

            strValor = Trim(FormataValor(aDoc(Count).Valor, 18))
            buff_aut = "Valor da operacao .......:" & String(18 - Len(strValor), "*") & strValor
            ret_aut = Autentica.Imprimir(buff_aut, False)

            buff_aut = "Agencia emitente: " & Format(Geral.AgenciaCentral, "0000") & " Agencia cliente: " & Format(Ag, "0000")
            ret_aut = Autentica.Imprimir(buff_aut, False)

            buff_aut = "Numero da conta: " & Format(Cta, "000000-0") & "  Num.Doc.: " & IIf(Len(Trim(aDoc(Count).NSU)) = 0, "0000000", Format(aDoc(Count).NSU, "0000000"))
            ret_aut = Autentica.Imprimir(buff_aut, False)

            buff_aut = "Data: " & Mid(Geral.DataProcessamento, 7, 2) + "/" + Mid(Geral.DataProcessamento, 5, 2) + "/" + Mid(Geral.DataProcessamento, 3, 2) & "                Hora: " & Format(Now, "hh:mm:ss")
            ret_aut = Autentica.Imprimir(buff_aut, False)

            buff_aut = "Controle do banco: " & SDV & "999999" & "#" & IIf(Len(Trim(aDoc(Count).NSU)) = 0, "000000", Format(aDoc(Count).NSU, "000000")) & "9" & Format(m_Agencia, "0000") & Format(aDoc(Count).Terminal, "000") & "#"
            ret_aut = Autentica.Imprimir(buff_aut, False)

            ret_aut = Autentica.Imprimir(buff_linha, False)

            buff_aut = "Descricao da Ocorrencia: "
            ret_aut = Autentica.Imprimir(buff_aut, False)
            
            StrMotivo = Trim(ObtemOcorrencia(Ocorrencia, 0))
            If Len(StrMotivo) < 45 Then
                buff_aut = StrMotivo
                ret_aut = Autentica.Imprimir(buff_aut, False)
            Else
                Pos = 45
                While Pos < Len(StrMotivo)
                    buff_aut = QuebraBuffer(StrMotivo, Pos)
                    ret_aut = Autentica.Imprimir(buff_aut, False)
                    StrMotivo = Right(StrMotivo, Len(StrMotivo) - Pos)
                Wend
                If Len(StrMotivo) > 0 Then
                    buff_aut = StrMotivo
                    ret_aut = Autentica.Imprimir(buff_aut, False)
                End If
            End If

            'Gravar Log -> Imprime Comprovante Ajuste DEB / CRED
            Call GravaLog(m_IdCapa, aDoc(Count).IdDocto, 84)

            ImprimeTrailler (False)
        End If
    Next
End Sub

Private Sub ImprimeDepositos()
    Dim ret_aut         As Integer
    Dim buff_aut        As String * 45
    Dim buff_linha      As String * 2
    Dim Count           As Integer
    Dim Ag              As Integer
    Dim Cta             As Long
    Dim strValor        As String
    
    Dim iVinculo        As Integer
    Dim lvinculo        As Long
    Dim sStr            As String
    
    If Geral.autenticadora = 0 Then
        Exit Sub
    End If
    
    buff_linha = Chr(13)
    
    For Count = 1 To m_CountDocto
        If (aDoc(Count).TipoDocto = 2 Or aDoc(Count).TipoDocto = 3) And _
           (aDoc(Count).Status <> "C" And aDoc(Count).Status <> "D" And aDoc(Count).Status <> "F") Then
           
            If Not ImprimeHeaderDeposito Then
                Exit Sub
            End If
        
            If Not ObtemAgConta(aDoc(Count).IdDocto, aDoc(Count).TipoDocto, Ag, Cta) Then
                cmdLimpar_Click
                Exit Sub
            End If
            
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '           Cálculo do novo Super DV
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Agencia = m_Agencia
            nroCaixa = Format(aDoc(Count).Terminal, "000")
            Data = Right(Geral.DataProcessamento, 2) & Mid(Geral.DataProcessamento, 5, 2) & Mid(Geral.DataProcessamento, 3, 2)
            TipoAgencia = "9"
            Valor = (aDoc(Count).Valor * 100)
            Operador = "999999"
            
            'Chama a DLL para o cálculo do Super DV
            Ret = QXGetSDV(Agencia, nroCaixa, Data, TipoAgencia, Valor, Operador, SDV)
        
            'Verifica se não houve erro e se os DVs estão OK
            If Ret <> 0 Then   'Or UCase(SDV) <> UCase(Mid(ControleBanco, 1, 1)) Then Exit Function
                MsgBox "Não é possível autenticar documento, cálculo do Super DV errado.", vbInformation + vbOKOnly, App.Title
                Exit Sub
            End If
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            If aDoc(Count).TipoDocto = 2 Then
            
                lvinculo = aDoc(Count).Vinculo
                If Trim(GrdDocto.TextMatrix(Count, COL_TRATAMENTO)) = "Finalizado" Then
                    sStr = "dinheiro"
                Else
                    sStr = "cheque"
                End If
                
                For iVinculo = 1 To m_CountDocto
                    If lvinculo = aDoc(iVinculo).Vinculo Then
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        'Se tipoDocto = 41 - Lancamento Interno, imprime Deposito em conta - dinheiro'
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        If aDoc(iVinculo).TipoDocto = 41 Then
                            sStr = "dinheiro"
                            Exit For
                        End If
                    End If
                Next iVinculo
            
                buff_aut = Space(4) & "Deposito em conta corrente - " & sStr
            Else
                If Trim(GrdDocto.TextMatrix(Count, COL_TRATAMENTO)) = "Finalizado" Then
                    buff_aut = Space(4) & "Deposito em conta poupanca - dinheiro"
                Else
                    buff_aut = Space(4) & "Deposito em conta poupanca - cheque"
                End If
                
            End If
            ret_aut = Autentica.Imprimir(buff_aut, False)
        
            strValor = Trim(FormataValor(aDoc(Count).Valor, 18))
            buff_aut = "Valor da operacao .......:" & String(18 - Len(strValor), "*") & strValor
            ret_aut = Autentica.Imprimir(buff_aut, False)
            
            buff_aut = "Agencia emitente: " & Format(Geral.AgenciaCentral, "0000") & " Agencia cliente: " & Format(Ag, "0000")
            ret_aut = Autentica.Imprimir(buff_aut, False)
                   
            buff_aut = "Numero da conta: " & Format(Cta, "000000-0") & "  Num.Doc.: " & IIf(Len(Trim(aDoc(Count).NSU)) = 0, "0000000", Format(aDoc(Count).NSU, "0000000"))
            ret_aut = Autentica.Imprimir(buff_aut, False)
            
            buff_aut = "Data: " & Mid(Geral.DataProcessamento, 7, 2) + "/" + Mid(Geral.DataProcessamento, 5, 2) + "/" + Mid(Geral.DataProcessamento, 3, 2) & "                Hora: " & Format(Now, "hh:mm:ss")
            ret_aut = Autentica.Imprimir(buff_aut, False)
            
            buff_aut = "Controle do banco: " & SDV & "999999" & "#" & IIf(Len(Trim(aDoc(Count).NSU)) = 0, "000000", Format(aDoc(Count).NSU, "000000")) & "9" & Format(m_Agencia, "0000") & Format(aDoc(Count).Terminal, "000") & "#"
            ret_aut = Autentica.Imprimir(buff_aut, False)

            'Gravar Log -> Imprimir Comprovante de Deposito
            Call GravaLog(m_IdCapa, aDoc(Count).IdDocto, 85)

            ImprimeTrailler (True)
        End If
    Next
End Sub
Private Function QuebraBuffer(ByVal Buf As String, ByRef Pos As Integer) As String
    Dim Tam As Integer
    
    Tam = Pos
    Do While Tam > 0
        If Mid(Buf, Tam, 1) = " " Then
            Exit Do
        End If
        Tam = Tam - 1
    Loop
    If Tam > 0 Then
        Pos = Tam
    End If
    QuebraBuffer = Mid(Buf, 1, Pos)
End Function

Private Sub cmbAgencia_Click()
    Dim Msg             As String
    Dim bPosicionou     As Boolean

    If Len(Trim(cmbAgencia.Text)) = 0 Then
        Exit Sub
    End If

    If rsExpedicao.RowCount > 0 Then
        rsExpedicao.MoveFirst
        Do While Not rsExpedicao.EOF
            If rsExpedicao!IdCapa = cmbAgencia.ItemData(cmbAgencia.ListIndex) Then
                Exit Do
            End If
            rsExpedicao.MoveNext
        Loop
    End If
    lblLote.Caption = Format(rsExpedicao!IdLote, "0000-00000")

    If rsExpedicao!IdEnv_Mal = "E" Then
        lblCapa.Caption = "Envelope"
        lblMalote.Enabled = False
        TxtNumMalote.Text = ""
        TxtNumMalote.Enabled = False
    Else
        lblCapa.Caption = "Malote"
        lblMalote.Enabled = True
        TxtNumMalote.Enabled = True
    End If

    'Entrar Capa
    GravaLog rsExpedicao!IdCapa, 0, 88

    Msg = "Esta " & _
        " capa não está disponível para Expedição, porque se encontra " & vbCrLf
    Select Case rsExpedicao!Status
        Case "0"      'Capa cadastrada
            Limpar_grdDocto
            Msg = Msg & "para Captura de Imagens. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "1"      'Capa digitalizada
            Limpar_grdDocto
            Msg = Msg & "para Complementação. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "2"      'Capa em complementação
            Limpar_grdDocto
            Msg = Msg & "em Complementação. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "3"      'Capa complementada, mas com pendência
            ' Nao existe mais este status
            Limpar_grdDocto
            Msg = Msg & "com Status inválido (3). "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "4"      'Capa para Prova Zero
            Limpar_grdDocto
            Msg = Msg & "para Prova Zero. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "5"      'Capa para Ilegíveis
            Limpar_grdDocto
            Msg = Msg & "para Ilegíveis. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "6"      'Capa para Alçada
            Limpar_grdDocto
            Msg = Msg & "para Alçada. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "7"      'Capa para Vínculo Manual
            Limpar_grdDocto
            Msg = Msg & "para Vínculo Manual. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "8"      'Capa para Vínculo Automatico
            Limpar_grdDocto
            Msg = Msg & "para Vínculo Automático. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "9"      'Capa p/ Vinc. Automatico, enviada pelo Prova Zero
            Limpar_grdDocto
            Msg = Msg & "para Vínculo Automático. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "A"      'Para Recaptura
            Limpar_grdDocto
            Msg = Msg & "para Recaptura. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "B"      'Em Recaptura
            Limpar_grdDocto
            Msg = Msg & "em Recaptura. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "F"      'Capa Devolvida pelo Robo
            Limpar_grdDocto
            Msg = Msg & "Devolvida pelo Robô. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "G"      'Capa em Prova Zero
            Limpar_grdDocto
            Msg = Msg & "em Prova Zero. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "H"      'Capa em Ilegiveis
            Limpar_grdDocto
            Msg = Msg & "em Ilegíveis. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "I"      'Capa em Alcada
            Limpar_grdDocto
            Msg = Msg & "em Alçada. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "J"      'Capa em Vinculo Manual
            Limpar_grdDocto
            Msg = Msg & "em Vínculo Manual. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "K"      'Capa em Expedicao
            If rsExpedicao!Intervalo <= Geral.Intervalo Then
                Limpar_grdDocto
                Msg = Msg & "em Expedição por outra estação. "
                MsgBox Msg, vbInformation + vbOKOnly, App.Title
                cmdLimpar_Click
                Exit Sub
            End If
        Case "L"       'Capa para Confirmação de Ag/Cc
            Limpar_grdDocto
            Msg = Msg & "para Confirmação de Ag/Cc. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "M"       'Capa em Confirmação de Ag/Cc
            Limpar_grdDocto
            Msg = Msg & "em Confirmação de Ag/Cc. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "N"       'Capa para CSP
            Limpar_grdDocto
            Msg = Msg & "para CSP. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "Q"       'Capa em CSP
            Limpar_grdDocto
            Msg = Msg & "em CSP. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "P"      'Capa Devolvida pela Preparação
            ' Depois que o Robot transmitir a ocorrencia
            ' desta capa, ele mudarah o status para "X"
            Limpar_grdDocto
            Msg = Msg & "para Transmissão. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "R"      'Capa para Transmissão
            Limpar_grdDocto
            Msg = Msg & "para Transmissão. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "S"      'Capa em Transmissão
            Limpar_grdDocto
            Msg = Msg & "em Transmissão. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "O"      'Capa em Troca de Ordem
            Limpar_grdDocto
            Msg = Msg & "em Troca de Ordem. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "W"      'Capa em Estorno
            Limpar_grdDocto
            Msg = Msg & "em Estorno. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "Y"      'Capa para Correção de Agencia e Conta
            Limpar_grdDocto
            Msg = Msg & "para Correção de Agência e Conta. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
        Case "Z"      'Capa em Correção de Agencia e Conta
            Limpar_grdDocto
            Msg = Msg & "em Correção de Agência e Conta. "
            MsgBox Msg, vbInformation + vbOKOnly, App.Title
            cmdLimpar_Click
            Exit Sub
    End Select

    If m_IdCapa > 0 And _
       m_IdCapa <> rsExpedicao!IdCapa Then
        FinalizarExpedicao
    End If

    m_IdCapa = rsExpedicao!IdCapa
    m_IdEnvMal = rsExpedicao!IdEnv_Mal
    m_Capa = Format(rsExpedicao!Capa, IIf(m_IdEnvMal = "E", "00000000", "00000000000000"))
    If m_IdEnvMal = "M" Then
        m_NumMalote = CStr(rsExpedicao!Num_Malote)
    Else
        m_NumMalote = ""
    End If
    m_Agencia = Format(rsExpedicao!AgOrig, "0000")
    m_Status = rsExpedicao!Status
    m_Ocorrencia = rsExpedicao!Ocorrencia
    m_CountDocto = 0

    If m_Status = "K" Then
        If m_Ocorrencia <> 0 Then
            If m_Ocorrencia = 998 Or m_Ocorrencia = 99800 Then
                m_Status = "D"
            Else
                m_Status = "X"
            End If
        ElseIf PossuiMotivoExclusao(m_IdCapa) Then
            m_Status = "D"
        Else
            m_Status = "T"
        End If
    End If

    AtualizaStatusCapa m_IdCapa, "K"

    If m_Status <> "D" And m_Status <> "X" Then
        ObtemDocumentos m_IdCapa
        Preenche_grdDocto
        GrdDocto.SetFocus
    Else
        If m_Ocorrencia <> 998 And m_Ocorrencia <> 99800 Then
            ImprimeOcorrenciaCapa
            MsgBox "Capa excluída. ", vbInformation + vbOKOnly, App.Title
        Else
            MsgBox "Capa excluída automaticamente por duplicidade. ", vbInformation + vbOKOnly, App.Title
        End If
        FinalizarExpedicao
        cmdLimpar_Click
        Exit Sub
    End If
    
    If m_CountDocto > 0 Then
        If rsExpedicao!IdEnv_Mal = "E" Then
            bPosicionou = PosicionaAutenticar
        Else
            bPosicionou = True
        End If
        
        If Not bPosicionou Or rsExpedicao!IdEnv_Mal = "M" Then
            GrdDocto.Row = 1
            selecionaLinha 1
            lblOcorrencia.Caption = MostraOcorrencia(1)
            MostraImagem 1
        End If
    End If
    
    If m_Status = "E" Then
        CmdReimpressao.Enabled = True
    Else
        CmdReimpressao.Enabled = False
        ImprimeOcorrenciaDoctos
        Call AutenticacaoGare
        ImprimeAjustes
        ImprimeCartaoAvulso
        If m_IdEnvMal = "M" Then
            ImprimeDepositos
        End If
    End If
    
    GrdDocto.SetFocus
    
End Sub

Private Sub cmbCapa_Click()
    cmbAgencia.Clear
    If rsExpedicao.RowCount > 0 Then
        rsExpedicao.MoveFirst
        Do While Not rsExpedicao.EOF
            If rsExpedicao!Capa = Val(cmbCapa.Text) Then
                cmbAgencia.AddItem Format(rsExpedicao!AgOrig, "0000")
                cmbAgencia.ItemData(cmbAgencia.NewIndex) = rsExpedicao!IdCapa
            End If
            rsExpedicao.MoveNext
        Loop
    End If
    If cmbAgencia.ListCount = 1 Then
        cmbAgencia.Text = cmbAgencia.List(0)
    Else
        cmbAgencia.SetFocus
        SendKeys "{F4}"
    End If
End Sub

Private Sub cmbCapa_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmbCapa.Text = Left(cmbCapa.Text, 14)
        If IsNumeric(cmbCapa.Text) Then
            If LocalizarCapa(1) Then
                TxtNumMalote.Text = FormataMalote(TxtNumMalote.Text)
            End If
        Else
            MsgBox "Número da Capa inválido.", vbExclamation + vbOKOnly, App.Title
            cmdLimpar_Click
        End If
        KeyAscii = 0
    End If
End Sub

Private Sub cmdAuditoria_Click()

    If m_IdCapa = 0 Then Exit Sub
        
    Geral.Capa.IdCapa = m_IdCapa
    Geral.Capa.Capa = Val(m_Capa)
    Geral.Capa.Num_Malote = Val(m_NumMalote)
    Geral.Capa.AgOrig = Val(m_Agencia)
    Geral.Capa.IdEnv_Mal = m_IdEnvMal
    
    Call Auditoria
    
    Geral.Capa.IdCapa = 0
    Geral.Capa.Capa = 0
    Geral.Capa.Num_Malote = 0
    Geral.Capa.AgOrig = 0
    Geral.Capa.IdEnv_Mal = ""

End Sub

Private Sub CmdFechar_Click()
         
    Unload Me
    
End Sub

Private Sub cmdFrenteVerso_Click()

    Dim i       As Integer
    
    If m_Busy Then
        Exit Sub
    End If
    If Not cmdFrenteVerso.Enabled Then
        Exit Sub
    End If
    m_Busy = True
    
    On Error GoTo ErroImagem
    
    i = aIndice(GrdDocto.Row)
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
                 .Load Geral.DiretorioImagens & Format(rsExpedicao!IdLote, "000000000") & "\" & aDoc(i).Frente, 0, 0, 1
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
                 .Load Geral.DiretorioImagens & Format(rsExpedicao!IdLote, "000000000") & "\" & aDoc(i).Verso, 0, 0, 1
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
    FrmImagem.Visible = False
    MsgBox "Não foi possível exibir imagem do documento, imagem não encontrada", vbExclamation + vbOKOnly, App.Title
    
End Sub

Private Sub cmdInverteCor_Click()
    If m_Busy Then
        Exit Sub
    End If

    If Not cmdInverteCor.Enabled Then Exit Sub

    m_Busy = True
    Lead1.Invert
    m_Busy = False
End Sub

Private Sub cmdLimpar_Click()

    If m_IdCapa > 0 Then

        'Verifica se Expedição deve ser encerrada ou não
        If FinalizarExpedicao Then

            LimpaHeader
            lblMalote.Enabled = True
            TxtNumMalote.Enabled = True
            FrmImagem.Visible = False
            m_CountDocto = 0
        
            bRunActivate = True
            GrdDocto.Clear
            GrdDocto.Rows = 9
            
            GrdDocto.TextMatrix(0, COL_DOCUMENTO) = "Documento"
            GrdDocto.TextMatrix(0, COL_DESCRICAO) = "Descrição"
            GrdDocto.TextMatrix(0, COL_VALOR) = "Valor"
            GrdDocto.TextMatrix(0, COL_TRATAMENTO) = "Tratamento"
            selecionaLinha 1

            bRunActivate = False
            cmbCapa.SetFocus
        End If
    Else
        LimpaHeader
        lblMalote.Enabled = True
        TxtNumMalote.Enabled = True
        FrmImagem.Visible = False
        m_CountDocto = 0
    
        bRunActivate = True
        GrdDocto.Clear
        GrdDocto.Rows = 9
    
        GrdDocto.TextMatrix(0, COL_DOCUMENTO) = "Documento"
        GrdDocto.TextMatrix(0, COL_DESCRICAO) = "Descrição"
        GrdDocto.TextMatrix(0, COL_VALOR) = "Valor"
        GrdDocto.TextMatrix(0, COL_TRATAMENTO) = "Tratamento"
        selecionaLinha 1

        bRunActivate = False
        cmbCapa.SetFocus
    End If
End Sub

Private Sub CmdReimpressao_Click()
    ImprimeOcorrenciaDoctos
    Call AutenticacaoGare
    ImprimeAjustes
    ImprimeCartaoAvulso
    
     If m_IdEnvMal = "M" Then
        ImprimeDepositos
     End If
        
    CmdReimpressao.Enabled = False
End Sub

Private Sub cmdRotacao_Click()
    If m_Busy Then
        Exit Sub
    End If
    If Not cmdRotacao.Enabled Then
        Exit Sub
    End If
    m_Busy = True
    Lead1.FastRotate 90
    m_Busy = False
End Sub

Private Sub cmdZoomMais_Click()
    If m_Busy Then
        Exit Sub
    End If
    If Not cmdZoomMais.Enabled Then
        Exit Sub
    End If
    m_Busy = True
    If Lead1.PaintZoomFactor <= 400 Then
        Lead1.PaintSizeMode = PAINTSIZEMODE_ZOOM
        Lead1.PaintZoomFactor = Lead1.PaintZoomFactor + 10
    End If
    m_Busy = False
End Sub

Private Sub cmdZoomMenos_Click()
    If m_Busy Then
        Exit Sub
    End If
    If Not cmdZoomMenos.Enabled Then
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
    Call AtualizaAtividade(11)
    
    If m_FirstActivate Then
        bRunActivate = True
        
        LimpaHeader
        m_IdCapa = 0
        
        GrdDocto.ColWidth(COL_IMAGEM) = Int(GrdDocto.Width * 0.08)          '   8   %
        GrdDocto.ColWidth(COL_DOCUMENTO) = Int(GrdDocto.Width * 0.32)       '  32   %
        GrdDocto.ColWidth(COL_DESCRICAO) = Int(GrdDocto.Width * 0.335)      '  33,5 %
        GrdDocto.ColWidth(COL_VALOR) = Int(GrdDocto.Width * 0.1)            '  10   %
        GrdDocto.ColWidth(COL_TRATAMENTO) = Int(GrdDocto.Width * 0.15)      '  15   %
        GrdDocto.ColWidth(COL_NR_AUTENTICA) = Int(GrdDocto.Width * 0.015)   '   1,5 %
        
        GrdDocto.TextMatrix(0, COL_DOCUMENTO) = "Documento"
        GrdDocto.TextMatrix(0, COL_DESCRICAO) = "Descrição"
        GrdDocto.TextMatrix(0, COL_VALOR) = "Valor"
        GrdDocto.TextMatrix(0, COL_TRATAMENTO) = "Tratamento"
        GrdDocto.TextMatrix(0, COL_NR_AUTENTICA) = ""
        
        selecionaLinha 1
        
        bRunActivate = False
        m_FirstActivate = False
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyAdd
            cmdZoomMais_Click
        Case vbKeySubtract
            cmdZoomMenos_Click
        Case vbKeyMultiply
            GrdDocto_DblClick
        Case vbKeyDivide
            cmdRotacao_Click
        Case vbKeyF10
            cmdInverteCor_Click
            KeyCode = 0
        Case vbKeyF11
            cmdFrenteVerso_Click
    End Select

End Sub

Private Sub Form_Load()
    With Lead1
        .UnlockSupport L_SUPPORT_EXPRESS, "YXPQ3XPPVT"
        .UnlockSupport L_SUPPORT_GIFLZW, "0K3RV9UY3EY"
        .UnlockSupport L_SUPPORT_TIFLZW, "9LE75L0FDXHK"
    End With
    
    Set qryGetCapaExpedicao = Geral.Banco.CreateQuery("", "{Call GetCapaExpedicao (?,?)}")
    Set qryGetMaloteExpedicao = Geral.Banco.CreateQuery("", "{Call GetMaloteExpedicao (?,?)}")
    Set qryGetocorrencia = Geral.Banco.CreateQuery("", "{Call GetOcorrencia (?)}")
    Set qryGetMotivoExclusao = Geral.Banco.CreateQuery("", "{Call GetMotivoExclusao (?,?)}")
    Set qryGetDescricaoDocumento = Geral.Banco.CreateQuery("", "{Call GetDescricaoDocumento (?)}")
    Set qryGetDocumentoExpedicao = Geral.Banco.CreateQuery("", "{Call GetDocumentoExpedicao (?,?)}")
    Set qryGetDocumentoOcorrencia = Geral.Banco.CreateQuery("", "{Call GetDocumentoOcorrencia (?,?)}")
    Set qryGetAgContaDocumento = Geral.Banco.CreateQuery("", "{Call GetAgContaDocumento (?,?,?)}")
    Set qryGetUsuario = Geral.Banco.CreateQuery("", "{call GetUsuario (?)}")
    Set qryGetCartaoAvulso = Geral.Banco.CreateQuery("", "{call GetCartaoAvulso (?,?)}")
    Set qryGetBHVCDescricao = Geral.Banco.CreateQuery("", "{call GetBHVCDescricao (?,?,?)}")
    Set qryAtualizaAutenticacao = Geral.Banco.CreateQuery("", "{? = Call AtualizaAutenticacaoDocumento (?,?,?)}")
    Set qryAtualizaStatusCapaDoctoExpedido = Geral.Banco.CreateQuery("", "{? = Call AtualizaStatusCapaDoctoExpedido (?,?,?)}")
    Set qryVerificaBinCartao = Geral.Banco.CreateQuery("", "{? = call VerificaBinCartao (?)}")
    Set qryLerParametro = Geral.Banco.CreateQuery("", "{call LerParametro(?)}")
    Set qryGetSaldoConta = Geral.Banco.CreateQuery("", "{call GetSaldoConta(?,?,?)}")
    
    m_FirstActivate = True
    
   'verifica permissao de usuario para reautenticacao
    UserReautentica = IIf(UCase(Geral.Usuario) = "DESENV", True, VerificaAcessoUsuario(Geral.idUsuario, 35))
    
    '''''''''''''''''''''''''''
    'Loga a acao Entrar Módulo'
    '''''''''''''''''''''''''''
    Call GravaLog(0, 0, 166)
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call CmdFechar_Click

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If m_IdCapa > 0 Then
        If Not FinalizarExpedicao Then
           Cancel = -1
           Exit Sub
        End If
    End If
    
    qryGetCapaExpedicao.Close
    qryGetMaloteExpedicao.Close
    qryGetocorrencia.Close
    qryGetMotivoExclusao.Close
    qryGetDescricaoDocumento.Close
    qryGetDocumentoExpedicao.Close
    qryGetDocumentoOcorrencia.Close
    qryGetUsuario.Close
    qryGetCartaoAvulso.Close
    qryAtualizaAutenticacao.Close
    qryAtualizaStatusCapaDoctoExpedido.Close
    qryGetAgContaDocumento.Close
    qryVerificaBinCartao.Close
    qryGetBHVCDescricao.Close
    qryLerParametro.Close
    
    '''''''''''''''''''''''''
    'Loga a acao Sair Modulo'
    '''''''''''''''''''''''''
    Call GravaLog(0, 0, 167)


End Sub

Private Sub GrdDocto_DblClick()
    ImprimeAutenticacao (GrdDocto.Row)
End Sub

Private Sub grdDocto_EnterCell()


    If m_FromEvent Then Exit Sub
    m_FromEvent = True

    If Not bRunActivate And m_CountDocto > 0 Then
        SelecionaPasta GrdDocto.Row, True
        lblOcorrencia.Caption = MostraOcorrencia(GrdDocto.Row)
        MostraImagem GrdDocto.Row
    End If
    m_FromEvent = False
End Sub

Private Sub GrdDocto_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdLimpar_Click
    End If
End Sub

Private Sub grdDocto_LeaveCell()


    If m_FromEvent Then Exit Sub
    m_FromEvent = True
    If Not bRunActivate And m_CountDocto > 0 Then
        SelecionaPasta GrdDocto.Row, False
    End If
    m_FromEvent = False
    
End Sub

Private Sub GrdDocto_SelChange()

    If GrdDocto.Row <> GrdDocto.RowSel Then
        selecionaLinha GrdDocto.RowSel
    End If
End Sub

Private Sub Lead1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Lead1.AutoRubberBand = True
        Lead1.MousePointer = 2
    Else
        MostraImagem GrdDocto.Row
    End If
End Sub

Private Sub Lead1_RubberBand()
    Dim zoomleft        As Integer
    Dim zoomtop         As Integer
    Dim zoomwidth       As Integer
    Dim zoomheight      As Integer
    
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

Private Sub txtNumMalote_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(TxtNumMalote) Then
            If Len(TxtNumMalote) = 12 Then
                If Left(TxtNumMalote.Text, 2) <> "09" Then
                    MsgBox "Número do Malote inválido.", vbExclamation + vbOKOnly, App.Title
                    cmdLimpar_Click
                    Exit Sub
                End If
            End If
            
            If LocalizarCapa(2) Then
                TxtNumMalote.Text = FormataMalote(TxtNumMalote.Text)
            End If
        Else
            MsgBox "Número do Malote inválido.", vbExclamation + vbOKOnly, App.Title
            cmdLimpar_Click
        End If
        KeyAscii = 0
    End If
End Sub

Private Sub Limpar_grdDocto()
    If Not bRunActivate Then
        bRunActivate = True
        
        m_CountDocto = 0
        
        GrdDocto.Clear
        GrdDocto.Rows = 9
        
        GrdDocto.TextMatrix(0, COL_DOCUMENTO) = "Documento"
        GrdDocto.TextMatrix(0, COL_DESCRICAO) = "Descrição"
        GrdDocto.TextMatrix(0, COL_VALOR) = "Valor"
        GrdDocto.TextMatrix(0, COL_TRATAMENTO) = "Tratamento"
        selecionaLinha 1
        
        FrmImagem.Visible = False
        
        bRunActivate = False
    End If
End Sub

Private Sub ImprimeCartaoAvulso()

    Dim ret_aut         As Integer
    Dim buff_aut        As String * 45
    Dim buff_linha      As String * 2
    Dim Count           As Integer
    
    Dim strValor        As String
    Dim NumCartao       As String
    Dim ValorCartao     As Currency
    Dim DespReais       As Currency
    Dim DespDolar       As Currency
    Dim AntSaque        As Currency
    Dim CodBandeira     As String

    If Geral.autenticadora = 0 Then
        Exit Sub
    End If

    buff_linha = Chr(13)

    For Count = 1 To m_CountDocto
        If aDoc(Count).TipoDocto = 36 And aDoc(Count).Status <> "D" And aDoc(Count).Status <> "F" And aDoc(Count).Status <> "C" Then

            If Not ObtemCartao(aDoc(Count).IdDocto, NumCartao, DespReais, DespDolar, AntSaque, ValorCartao, CodBandeira) Then
                cmdLimpar_Click
                Exit Sub
            End If
            
            If Not ImprimeHeaderCartao(NumCartao, CodBandeira) Then
                Exit Sub
            End If

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '           Cálculo do novo Super DV
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Agencia = m_Agencia
            nroCaixa = Format(aDoc(Count).Terminal, "000")
            Data = Right(Geral.DataProcessamento, 2) & Mid(Geral.DataProcessamento, 5, 2) & Mid(Geral.DataProcessamento, 3, 2)
            TipoAgencia = "9"
            Valor = (ValorCartao * 100)
            Operador = "999999"
            
            'Chama a DLL para o cálculo do Super DV
            Ret = QXGetSDV(Agencia, nroCaixa, Data, TipoAgencia, Valor, Operador, SDV)
        
            'Verifica se não houve erro e se os DVs estão OK
            If Ret <> 0 Then   'Or UCase(SDV) <> UCase(Mid(ControleBanco, 1, 1)) Then Exit Function
                MsgBox "Não é possível autenticar documento, cálculo do Super DV errado.", vbInformation + vbOKOnly, App.Title
                Exit Sub
            End If
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''

            strValor = Trim(FormataValor(ValorCartao, 20))
            buff_aut = "Valor Total da operacao:" & String(20 - Len(strValor), "*") & strValor
            ret_aut = Autentica.Imprimir(buff_aut, False)

            buff_aut = "Agencia emitente: " & Format(Geral.AgenciaCentral, "0000")
            ret_aut = Autentica.Imprimir(buff_aut, False)

            buff_aut = "Numero do Cartao:........." & NumCartao
            ret_aut = Autentica.Imprimir(buff_aut, False)
            
            buff_aut = Space(5) & "Recebimento de Despesas em Reais"
            ret_aut = Autentica.Imprimir(buff_aut, False)
            
            strValor = Trim(FormataValor(DespReais, 18))
            buff_aut = "Valor (R$):..............." & String(18 - Len(strValor), "*") & strValor
            ret_aut = Autentica.Imprimir(buff_aut, False)
            
            buff_aut = Space(5) & "Recebimento de Despesas em Dolar"
            ret_aut = Autentica.Imprimir(buff_aut, False)
            
            strValor = Trim(FormataValor(DespDolar, 18))
            buff_aut = "Valor (R$):..............." & String(18 - Len(strValor), "*") & strValor
            ret_aut = Autentica.Imprimir(buff_aut, False)
            
            buff_aut = Space(5) & "Recebimento Antecipado de Saque"
            ret_aut = Autentica.Imprimir(buff_aut, False)
            
            strValor = Trim(FormataValor(AntSaque, 18))
            buff_aut = "Valor (R$):..............." & String(18 - Len(strValor), "*") & strValor
            ret_aut = Autentica.Imprimir(buff_aut, False)
            
            buff_aut = "Data: " & Mid(Geral.DataProcessamento, 7, 2) + "/" + Mid(Geral.DataProcessamento, 5, 2) + "/" + Mid(Geral.DataProcessamento, 3, 2) & "                Hora: " & Format(Now, "hh:mm:ss")
            ret_aut = Autentica.Imprimir(buff_aut, False)

            buff_aut = "Controle do banco: " & SDV & "999999" & "#" & IIf(Len(Trim(aDoc(Count).NSU)) = 0, "000000", Format(aDoc(Count).NSU, "000000")) & "0" & Format(m_Agencia, "0000") & Format(aDoc(Count).Terminal, "000") & "#"
            ret_aut = Autentica.Imprimir(buff_aut, False)
            
            buff_aut = String(44, "-")
            ret_aut = Autentica.Imprimir(buff_aut, False)

            buff_aut = "A exatidao dos dados informados neste  docu-"
            ret_aut = Autentica.Imprimir(buff_aut, False)
            
            buff_aut = "mento e de total responsabilidade  do  asso-"
            ret_aut = Autentica.Imprimir(buff_aut, False)
            
            buff_aut = "ciado, eximindo-se o Unibanco  de  quaisquer"
            ret_aut = Autentica.Imprimir(buff_aut, False)
            
            buff_aut = "responsabilidades pela demora ou nao  quita-"
            ret_aut = Autentica.Imprimir(buff_aut, False)
            
            buff_aut = "cao da fatura acima em virtude de informacao"
            ret_aut = Autentica.Imprimir(buff_aut, False)
            
            buff_aut = "incorreta por parte do associado.           "
            ret_aut = Autentica.Imprimir(buff_aut, False)
            
            'Gravar Log -> Imprime Comprovante cartao avulso
            Call GravaLog(m_IdCapa, aDoc(Count).IdDocto, 87)

            ImprimeTrailler (True)
        End If
    Next
End Sub

Private Function ObtemCartao(ByVal IdDocto As Long, _
                                   NumCartao As String, _
                                   DespReais As Currency, _
                                   DespDolar As Currency, _
                                   AntSaque As Currency, _
                                   Valor As Currency, _
                                   CodBandeira As String) As Boolean
                                   
    Dim rsCartao As rdoResultset
    
    On Error GoTo ErroCartao
    rdoErrors.Clear
    
    Screen.MousePointer = vbHourglass
    
    qryGetCartaoAvulso.rdoParameters(0) = Geral.DataProcessamento
    qryGetCartaoAvulso.rdoParameters(1) = IdDocto
    Set rsCartao = qryGetCartaoAvulso.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    
    If rsCartao.EOF Then
        ObtemCartao = False
    Else
        '''''''''''''''''''''''''''''''''''''''''''''''''''
        'Comentado por existir problemas de arredondamento'
        '''''''''''''''''''''''''''''''''''''''''''''''''''
        NumCartao = FormataString(rsCartao!Cartao, "0", 16, True)
        DespReais = rsCartao!DespReais
        DespDolar = rsCartao!DespDolar
        AntSaque = rsCartao!AntSaque
        Valor = rsCartao!Valor
        
        If ObtemBandeiraCartao(CStr(rsCartao!Cartao), CodBandeira) Then
            ObtemCartao = True
        End If
      
        ObtemCartao = True
    End If
    
    rsCartao.Close
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    Exit Function
    
ErroCartao:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro na obtenção do Cartão Avulso.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    m_Busy = False
    Unload Me
End Function

Private Function ImprimeHeaderCartao(ByVal NumCartao As String, ByVal CodBandeira As String) As Boolean
    Dim ret_imp, ret_aut    As Integer
    Dim Buff_st             As String * 3
    Dim buff_aut            As String * 45
    Dim r1, r2, r3          As Integer
     
    ImprimeHeaderCartao = False
    
    buff_aut = String(44, "-")
    
    If Geral.autenticadora = 2 Then
        ret_imp = Autentica.Status(Buff_st)
        r1 = Hex(Asc(Mid(Buff_st, 1, 1)))
        r2 = Hex(Asc(Mid(Buff_st, 2, 1)))
        r3 = Hex(Asc(Mid(Buff_st, 3, 1)))
        If (r1) <> 0 Then
            ret_aut = 1
        Else
            ret_aut = Autentica.Imprimir(buff_aut, False)
        End If
    Else
        ret_aut = Autentica.Imprimir(buff_aut, False)
    End If
    
    If (ret_aut <> 0) Then
        ret_imp = Autentica.Status(Buff_st)
    
        If (ret_imp = 0) Then
           
           r1 = Hex(Asc(Mid(Buff_st, 1, 1)))
           r2 = Hex(Asc(Mid(Buff_st, 2, 1)))
           r3 = Hex(Asc(Mid(Buff_st, 3, 1)))
           
           If (r1 + r2 + r3) <> 0 Then
              'teste do 1 byte
              Select Case r1
                 Case 1
                    MsgBox "Autenticadora está off-line.", vbInformation + vbOKOnly, App.Title
                 Case 2
                    MsgBox "Autenticadora está desligada.", vbInformation + vbOKOnly, App.Title
                 Case 3
                    MsgBox "Autenticadora está com buffer cheio.", vbInformation + vbOKOnly, App.Title
                 Case 4
                    MsgBox "Autenticadora está inoperante.", vbInformation + vbOKOnly, App.Title
              End Select
           
              'teste do 3 byte
              If (r3 <> 0) Then
                 MsgBox "Verifique a bobina da Autentica.", vbInformation + vbOKOnly, App.Title
              End If
              Exit Function
           End If
        Else
           MsgBox "Verifique a Autentica.", vbInformation + vbOKOnly, App.Title
           Exit Function
        End If
    Else
        buff_aut = Space(13) & "UBB - Unibanco SA"
        ret_aut = Autentica.Imprimir(buff_aut, False)
        
        Select Case CodBandeira
            Case "00"       'Unibanco Visa
                buff_aut = Space(2) & "Cartao Unibanco Visa - Recebimento Avulso"
            Case "01"       'Unibanco MasterCard
                buff_aut = "Cartao Unibanco Mastercard-Recebimento Avulso"
            Case "02"       'Diners
                buff_aut = Space(4) & "Cartao Diners - Recebimento Avulso"
            Case "03"       'Credicard/Mastercard
                buff_aut = Space(4) & "Cartao Credicard - Recebimento Avulso"
        End Select
        
        ret_aut = Autentica.Imprimir(buff_aut, False)
    End If
    
    ImprimeHeaderCartao = True
    
End Function

Private Function ObtemBandeiraCartao(ByVal sCdCartao As String, ByRef sCodigoBandeira As String) As Boolean
'Parâmetro: (sCdCartao) - Seis caracteres iniciais do Número do Cartão
'
'Retorno:   (0)- Sucesso
'           (1)- Erro no SQL
Dim rstBandeira As rdoResultset

On Error GoTo Err_ObtemBandeiraCartao

    ObtemBandeiraCartao = False
    
    With qryVerificaBinCartao
        .rdoParameters(0).Direction = rdParamReturnValue
        
        .rdoParameters(1) = Left(sCdCartao, 6)
        Set rstBandeira = .OpenResultset(rdOpenStatic)
        
        'Verifica se ocorreu erro no SQL
        If .rdoParameters(0).Value = 1 Then GoTo Err_ObtemBandeiraCartao
        
        'Verifica se existe Código BIN do Cartão
        If rstBandeira.EOF Then Exit Function
        
        'Retorna código de Bandeira
        sCodigoBandeira = rstBandeira!crefsbandei
    End With
    
    ObtemBandeiraCartao = True
    Exit Function

Err_ObtemBandeiraCartao:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Não foi possível Verificar Número do Cartão.( VerifBinCartao )", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    m_Busy = False
    Unload Me

End Function
Private Function VerificaSaldoCheque(viddocto As Long, ByVal pTipoDocto As Integer)

    Dim ub      As Integer
    Dim i       As Integer

    With qryGetSaldoConta
        .rdoParameters(0) = Geral.DataProcessamento
        .rdoParameters(1) = viddocto
        .rdoParameters(2) = pTipoDocto
        Set rsSaldoConta = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    ub = UBound(SaldoInsuficiente)
    
    If Not rsSaldoConta.EOF Then
        For i = 0 To ub
            If SaldoInsuficiente(i).Agencia = rsSaldoConta!Agencia And _
               SaldoInsuficiente(i).Conta = Mid(rsSaldoConta!Conta, 1, Len(rsSaldoConta!Conta) - 1) & "-" & Right(rsSaldoConta!Conta, 1) Then
               
                Exit Function
            End If
        Next i
               
        SaldoInsuficiente(ControleSaldo).IdDocto = viddocto
        SaldoInsuficiente(ControleSaldo).Agencia = rsSaldoConta!Agencia
        SaldoInsuficiente(ControleSaldo).Conta = Mid(rsSaldoConta!Conta, 1, Len(rsSaldoConta!Conta) - 1) & "-" & Right(rsSaldoConta!Conta, 1)
        SaldoInsuficiente(ControleSaldo).DataSaldo = Format(Mid(rsSaldoConta!DataHoraSaldo, 1, 8), "DD/MM/YY")
        SaldoInsuficiente(ControleSaldo).HoraSaldo = Mid(IIf(IsNull(rsSaldoConta!DataHoraSaldo), "", rsSaldoConta!DataHoraSaldo), 10, 5)
        SaldoInsuficiente(ControleSaldo).SaldoDisponivel = Format(rsSaldoConta!SaldoDisponivel, "###,###,#00.00")
        SaldoInsuficiente(ControleSaldo).ValorBloqueado = Format(rsSaldoConta!ValorBloqueado, "###,###,#00.00")
        SaldoInsuficiente(ControleSaldo).LimiteChequeEspecial = Format(rsSaldoConta!LimiteChequeEspecial, "###,###,#00.00")
        SaldoInsuficiente(ControleSaldo).ValorSaldoAtual = Format(rsSaldoConta!SaldoTotal, "###,###,#00.00")
        ControleSaldo = ControleSaldo + 1
            
    End If

End Function
Private Sub ImprimeSaldoConta()
    Dim ret_imp, ret_aut    As Integer
    Dim Buff_st             As String * 3
    Dim buff_aut            As String * 45
    Dim buff_linha          As String * 2
    Dim r1, r2, r3          As Integer
    Dim Contador            As Double
    
    On Error GoTo Erro_ImprimeSaldo
    
    For Contador = 0 To ControleSaldo
    
        If SaldoInsuficiente(Contador).IdDocto = 0 Then Exit Sub
        
        buff_aut = String(45, "=")
        ret_aut = Autentica.Imprimir(buff_aut, False)
        
        buff_linha = Chr(13)
        ret_aut = Autentica.Imprimir(buff_linha, False)
        
        buff_aut = Space(7) & "SALDO PARA SIMPLES CONFERENCIA"
        ret_aut = Autentica.Imprimir(buff_aut, False)
        
        buff_aut = Space(15) & "CONTA CORRENTE"
        
        If Geral.autenticadora = 2 Then
            ret_imp = Autentica.Status(Buff_st)
            r1 = Hex(Asc(Mid(Buff_st, 1, 1)))
            r2 = Hex(Asc(Mid(Buff_st, 2, 1)))
            r3 = Hex(Asc(Mid(Buff_st, 3, 1)))
            If (r1) <> 0 Then
                ret_aut = 1
            Else
                ret_aut = Autentica.Imprimir(buff_aut, False)
            End If
        Else
            ret_aut = Autentica.Imprimir(buff_aut, False)
        End If
        
        If (ret_aut <> 0) Then
            ret_imp = Autentica.Status(Buff_st)
        
            If (ret_imp = 0) Then
               
               r1 = Hex(Asc(Mid(Buff_st, 1, 1)))
               r2 = Hex(Asc(Mid(Buff_st, 2, 1)))
               r3 = Hex(Asc(Mid(Buff_st, 3, 1)))
               
               If (r1 + r2 + r3) <> 0 Then
                  'teste do 1 byte
                  Select Case r1
                     Case 1
                        MsgBox "Autenticadora está off-line.", vbInformation + vbOKOnly, App.Title
                     Case 2
                        MsgBox "Autenticadora está desligada.", vbInformation + vbOKOnly, App.Title
                     Case 3
                        MsgBox "Autenticadora está com buffer cheio.", vbInformation + vbOKOnly, App.Title
                     Case 4
                        MsgBox "Autenticadora está inoperante.", vbInformation + vbOKOnly, App.Title
                  End Select
               
                  'teste do 3 byte
                  If (r3 <> 0) Then
                     MsgBox "Verifique a bobina da Autentica.", vbInformation + vbOKOnly, App.Title
                  End If
                  Exit Sub
               End If
            Else
               MsgBox "Verifique a Autentica.", vbInformation + vbOKOnly, App.Title
               Exit Sub
            End If
        Else
            ret_aut = Autentica.Imprimir(buff_linha, False)
        
            buff_aut = "Agencia/Conta" & Space(2) & "Lim Ch Especial" & Space(3) & "Data" & Space(3) & "Hora"
            ret_aut = Autentica.Imprimir(buff_aut, False)
            
            buff_aut = SaldoInsuficiente(Contador).Agencia & "/" & SaldoInsuficiente(Contador).Conta & Space(2) & String(15 - Len(SaldoInsuficiente(Contador).LimiteChequeEspecial), Space(1)) & SaldoInsuficiente(Contador).LimiteChequeEspecial & Space(1) & SaldoInsuficiente(Contador).DataSaldo & Space(1) & SaldoInsuficiente(Contador).HoraSaldo
            ret_aut = Autentica.Imprimir(buff_aut, False)
            
            buff_linha = Chr(13)
            ret_aut = Autentica.Imprimir(buff_linha, False)
            
            buff_aut = "Saldo Disponivel.............." & String(15 - Len(SaldoInsuficiente(Contador).SaldoDisponivel), ".") & SaldoInsuficiente(Contador).SaldoDisponivel
            ret_aut = Autentica.Imprimir(buff_aut, False)
            
            buff_aut = "Valor Bloqueado..............." & String(15 - Len(SaldoInsuficiente(Contador).ValorBloqueado), ".") & SaldoInsuficiente(Contador).ValorBloqueado
            ret_aut = Autentica.Imprimir(buff_aut, False)
            
            buff_aut = "Saldo Total Atual............." & String(15 - Len(SaldoInsuficiente(Contador).ValorSaldoAtual), ".") & SaldoInsuficiente(Contador).ValorSaldoAtual
            ret_aut = Autentica.Imprimir(buff_aut, False)
            
            buff_linha = Chr(13)
            ret_aut = Autentica.Imprimir(buff_linha, False)
            
            buff_aut = String(45, "-")
            ret_aut = Autentica.Imprimir(buff_aut, False)
                        
            buff_linha = Chr(13)
            ret_aut = Autentica.Imprimir(buff_linha, False)
            
            buff_linha = Chr(13)
            ret_aut = Autentica.Imprimir(buff_linha, False)
            
            buff_linha = Chr(13)
            ret_aut = Autentica.Imprimir(buff_linha, False)
            
            buff_linha = Chr(13)
            ret_aut = Autentica.Imprimir(buff_linha, False)
            
            buff_linha = Chr(13)
            ret_aut = Autentica.Imprimir(buff_linha, False)
            
            buff_linha = Chr(13)
            ret_aut = Autentica.Imprimir(buff_linha, False)
            
            buff_linha = Chr(13)
            ret_aut = Autentica.Imprimir(buff_linha, False)
            
        End If
        
    Next
    
Erro_ImprimeSaldo:

End Sub

Private Function ImprimeAutenticacaoGare() As Boolean

    Dim ret_aut             As Integer
    Dim buff_aut            As String * 45
    Dim buff_linha          As String * 2
    Dim nVias               As Integer

    buff_linha = Chr(13)

    On Error GoTo ErroDocOcor
    ImprimeAutenticacaoGare = False
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '           Cálculo do novo Super DV
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Agencia = m_Agencia
    nroCaixa = Format(m_Gare.Terminal, "000")
    Data = Right(Geral.DataProcessamento, 2) & Mid(Geral.DataProcessamento, 5, 2) & Mid(Geral.DataProcessamento, 3, 2)
    TipoAgencia = "9"
    Valor = (m_Gare.VlrTotal * 100)
    Operador = "999999"
    
    'Chama a DLL para o cálculo do Super DV
    Ret = QXGetSDV(Agencia, nroCaixa, Data, TipoAgencia, Valor, Operador, SDV)

    'Verifica se não houve erro e se os DVs estão OK
    If Ret <> 0 Then   'Or UCase(SDV) <> UCase(Mid(ControleBanco, 1, 1)) Then Exit Function
        MsgBox "Não é possível autenticar documento, cálculo do Super DV errado.", vbInformation + vbOKOnly, App.Title
        Exit Function
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    For nVias = 1 To m_Gare.NumeroViasGARE
        'Imprime Header do ticket
        If Not ImprimeHeaderAutenticaGARE Then
        
            Exit Function
        End If
    
        If m_Gare.DataVencto <> 0 Then
            buff_aut = "Data de Vencimento.........: " & Right(m_Gare.DataVencto, 2) & "/" & Mid(m_Gare.DataVencto, 5, 2) & "/" & Left(m_Gare.DataVencto, 4)
        Else
            buff_aut = "Data de Vencimento.........: "
        End If
        ret_aut = Autentica.Imprimir(buff_aut, False)
        
        buff_aut = "Codigo da Receita..........: " & Format(m_Gare.CodReceita, "000-0")
        ret_aut = Autentica.Imprimir(buff_aut, False)
        
        buff_aut = "Inscr.Estadual/Cd.Municipio: " & Format(m_Gare.InscrEstadual, "000000000000")
        ret_aut = Autentica.Imprimir(buff_aut, False)
        
        buff_aut = "CPF/CNPJ...................: " & Format(m_Gare.CPFCNPJ, "0000000000000-00")
        ret_aut = Autentica.Imprimir(buff_aut, False)
        
        buff_aut = "Inscr.Divida/Num.Etiqueta..: " & Format(m_Gare.InscrDivida, "0000000000000")
        ret_aut = Autentica.Imprimir(buff_aut, False)
        
        buff_aut = "Num. AIIM..................: " & Format(m_Gare.NumAIIM, "000000000000")
        ret_aut = Autentica.Imprimir(buff_aut, False)
        
        buff_aut = "Valor da Receita...........: " & Right(String(12, "*") & CStr(Format(m_Gare.VlrReceita, "#########0.00")), 16)
        ret_aut = Autentica.Imprimir(buff_aut, False)
        
        buff_aut = "Juros de Mora..............: " & Right(String(12, "*") & CStr(Format(m_Gare.VlrJuros, "#########0.00")), 16)
        ret_aut = Autentica.Imprimir(buff_aut, False)
        
        buff_aut = "Multa de Mora/Infracao.....: " & Right(String(12, "*") & CStr(Format(m_Gare.VlrMulta, "#########0.00")), 16)
        ret_aut = Autentica.Imprimir(buff_aut, False)
        
        buff_aut = "Acrescimos Financeiros.....: " & Right(String(12, "*") & CStr(Format(m_Gare.VlrAcrescimo, "#########0.00")), 16)
        ret_aut = Autentica.Imprimir(buff_aut, False)
        
        buff_aut = "Honorarios Advocaticios....: " & Right(String(12, "*") & CStr(Format(m_Gare.VlrHonorario, "#########0.00")), 16)
        ret_aut = Autentica.Imprimir(buff_aut, False)
        
        ret_aut = Autentica.Imprimir(buff_linha, False)
        buff_aut = "Valor Total................: " & Right(String(12, "*") & CStr(Format(m_Gare.VlrTotal, "#########0.00")), 16)
        ret_aut = Autentica.Imprimir(buff_aut, False)
    
        ret_aut = Autentica.Imprimir(buff_linha, False)
        buff_aut = "Ag. Emitente: " & Format(m_Agencia, "0000") & " - " & Left(m_Gare.NomeAgenciaColeta, 24)
        ret_aut = Autentica.Imprimir(buff_aut, False)
    
        ret_aut = Autentica.Imprimir(buff_linha, False)
        buff_aut = "Data: " & Format(Now, "dd/mm/yy") & Space(17) & "Hora: " & Format(Now, "hh:mm:ss")
        ret_aut = Autentica.Imprimir(buff_aut, False)
    
        buff_aut = "Controle do banco: " & SDV & "999999" & "#" & Format(m_Gare.NSU, "000000") & "9" & Format(m_Agencia, "0000") & Format(m_Gare.Terminal, "000") & "#"
        ret_aut = Autentica.Imprimir(buff_aut, False)
    
        buff_aut = String(45, "-")
        ret_aut = Autentica.Imprimir(buff_aut, False)
        
        buff_aut = Space(13) & "Autenticacao Digital"
        ret_aut = Autentica.Imprimir(buff_aut, False)
        
        buff_aut = Space(5) & Left(m_Gare.AutenticacaoGare, 8) & " " & Mid(m_Gare.AutenticacaoGare, 9, 8) & " " & Mid(m_Gare.AutenticacaoGare, 17, 8) & " " & Mid(m_Gare.AutenticacaoGare, 25, 8)
        ret_aut = Autentica.Imprimir(buff_aut, False)
        
        buff_aut = Space(5) & Mid(m_Gare.AutenticacaoGare, 33, 8) & " " & Mid(m_Gare.AutenticacaoGare, 41, 8) & " " & Mid(m_Gare.AutenticacaoGare, 49, 8) & " " & Mid(m_Gare.AutenticacaoGare, 57, 8)
        ret_aut = Autentica.Imprimir(buff_aut, False)
    
        buff_aut = String(45, "-")
        ret_aut = Autentica.Imprimir(buff_aut, False)
    
        buff_aut = Space(7) & "Recolhimento conforme Portarias:"
        ret_aut = Autentica.Imprimir(buff_aut, False)
    
        buff_aut = Space(12) & "CAT - 98 de 04/12/1997"
        ret_aut = Autentica.Imprimir(buff_aut, False)
    
        buff_aut = Space(12) & "CAT - 60 de 08/08/2002"
        ret_aut = Autentica.Imprimir(buff_aut, False)
    
        buff_aut = Space(4) & "Autorizado pelo Processo DAN 1816/98"
        ret_aut = Autentica.Imprimir(buff_aut, False)
    
        ret_aut = Autentica.Imprimir(buff_linha, False)
        
        'Verifica se última via
        If nVias = m_Gare.NumeroViasGARE Then
            buff_aut = CStr(nVias) & "a " & "Via-Contribuinte"
        Else
            buff_aut = CStr(nVias) & "a " & "Via"
        End If
        ret_aut = Autentica.Imprimir(buff_aut, False)
    
        buff_aut = String(45, "-")
        ret_aut = Autentica.Imprimir(buff_aut, False)
    
        buff_aut = Space(14) & "Ticket de Caixa."
        ret_aut = Autentica.Imprimir(buff_aut, False)
    
        buff_aut = Space(4) & "Feito para facilitar o seu dia-a-dia."
        ret_aut = Autentica.Imprimir(buff_aut, False)
    
        ret_aut = Autentica.Imprimir(buff_linha, False)
        ret_aut = Autentica.Imprimir(buff_linha, False)
        ret_aut = Autentica.Imprimir(buff_linha, False)
    
        'Espera 10 segundos para impressao da nova via do GARE devido ao Buffer de impressao para Procomp
        If Geral.autenticadora = estAutProcomp Then Espera (10)
    Next
                
    ret_aut = Autentica.Imprimir(buff_linha, False)
    ret_aut = Autentica.Imprimir(buff_linha, False)
    ret_aut = Autentica.Imprimir(buff_linha, False)
    ret_aut = Autentica.Imprimir(buff_linha, False)
    ret_aut = Autentica.Imprimir(buff_linha, False)
    ret_aut = Autentica.Imprimir(buff_linha, False)
    
    ImprimeAutenticacaoGare = True
    Exit Function

ErroDocOcor:

    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro na impressão de GARE com Autenticacão Digital.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    m_Busy = False

End Function
Private Function AutenticacaoGare() As Boolean

Dim nTotDocto As Integer
Dim qryGare As New rdoQuery
Dim rsGare As rdoResultset

AutenticacaoGare = False

On Error GoTo Err_AutenticacaoGare

    If Geral.autenticadora = 0 Then
        AutenticacaoGare = True
        GoTo Exit_AutenticacaoGare
    End If
    
    'Obtem Nome da Agencia de coleta
    m_Gare.NomeAgenciaColeta = ObtemAGENF(m_Agencia)
    
    Set qryGare = Geral.Banco.CreateQuery("", "{ ? = call GetGare(?,?)}")
    With qryGare
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1).Value = Geral.DataProcessamento
    End With
    
    'Desabilita Botões
    Frame2.Enabled = False
    
    For nTotDocto = 1 To UBound(aDoc)
    
        If Len(Trim(aDoc(nTotDocto).AutenticacaoGare)) <> 0 Then
            
            With qryGare
                .rdoParameters(2).Value = aDoc(nTotDocto).IdDocto
                
                Set rsGare = Nothing
                Set rsGare = .OpenResultset(rdOpenStatic, rdConcurReadOnly)
            
                If .rdoParameters(0).Value <> 0 Then GoTo Err_AutenticacaoGare
            End With
            
            If Not rsGare.EOF Then
                m_Gare.NSU = CStr(aDoc(nTotDocto).NSU)
                m_Gare.Terminal = CStr(aDoc(nTotDocto).Terminal)
                
                m_Gare.DataVencto = CStr(rsGare!vecto)
                m_Gare.CodReceita = 0 & rsGare!Receita
                m_Gare.InscrEstadual = CStr(rsGare!InscricaoEstadual)
                m_Gare.CPFCNPJ = CStr(rsGare!CPFCGC)
                m_Gare.InscrDivida = CStr(rsGare!DividaAtiva)
                m_Gare.NumAIIM = CStr(rsGare!AIIM)
                m_Gare.AutenticacaoGare = aDoc(nTotDocto).AutenticacaoGare
                m_Gare.NumeroViasGARE = Val(rsGare!Numero_Vias_Comprovante)
                m_Gare.VlrReceita = CStr(rsGare!ValorReceita)
                m_Gare.VlrJuros = CStr(rsGare!Juros)
                m_Gare.VlrMulta = CStr(rsGare!Multa)
                m_Gare.VlrAcrescimo = CStr(rsGare!Acrescimo)
                m_Gare.VlrHonorario = CStr(rsGare!Honorarios)
                m_Gare.VlrTotal = CStr(rsGare!Valor)
                
                If Not ImprimeAutenticacaoGare Then GoTo Exit_AutenticacaoGare
                
            End If
        
        End If
    Next

    'Habilita Botões
    Frame2.Enabled = True

    AutenticacaoGare = True

Exit_AutenticacaoGare:
    'Habilita Botões
    Frame2.Enabled = True

    If Not (rsGare Is Nothing) Then Set rsGare = Nothing
    qryGare.Close
    Exit Function

Err_AutenticacaoGare:
    'Habilita Botões
    Frame2.Enabled = True
    
    MsgBox "Erro ao obter informações do GARE, tente novamente !", vbCritical + vbOKOnly, App.Title
    GoTo Exit_AutenticacaoGare

End Function
Private Function ObtemAGENF(ByVal intAgOrig As Integer) As String

Dim qryAGENF As New rdoQuery
Dim RsAgenf As rdoResultset

On Error GoTo Err_ObtemAGENF

ObtemAGENF = ""

    Set qryAGENF = Geral.Banco.CreateQuery("", "{ call ObtemAgencia(?)}")
    
    qryAGENF.rdoParameters(0).Value = intAgOrig
        
    Set RsAgenf = qryAGENF.OpenResultset(rdOpenStatic, rdConcurReadOnly)
            
    If Not RsAgenf.EOF Then ObtemAGENF = Trim(RsAgenf!agefsnoagen)
    
    
Exit_ObtemAGENF:
    If Not (RsAgenf Is Nothing) Then Set RsAgenf = Nothing
    qryAGENF.Close
    Exit Function

Err_ObtemAGENF:
    MsgBox "Erro ao obter informações do GARE, tente novamente !", vbCritical + vbOKOnly, App.Title
    GoTo Exit_ObtemAGENF
    
End Function
Private Sub Espera(ByVal iPauseTime As Integer)

Dim start

   start = Timer
   Do While Timer < start + iPauseTime: Loop

End Sub
