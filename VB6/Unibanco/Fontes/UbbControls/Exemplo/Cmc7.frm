VERSION 5.00
Object = "{39F894DF-E245-11D4-B08D-00600899AB13}#1.5#0"; "UbbLVImg.ocx"
Object = "{ED123F48-E23F-11D4-B08D-00600899AB13}#1.0#0"; "UbbEdit.ocx"
Object = "{AD56946E-E23A-11D4-B08D-00600899AB13}#1.0#0"; "BarStat.ocx"
Object = "{798C3070-E23D-11D4-B08D-00600899AB13}#1.0#0"; "UbbFtp.ocx"
Begin VB.Form frmCmc7 
   Caption         =   "CMC7"
   ClientHeight    =   6570
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11880
   Icon            =   "Cmc7.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrEstacaoAtiva 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   2340
      Top             =   60
   End
   Begin UbbLVImg.UBBImage imgControlador 
      Height          =   465
      Left            =   1320
      TabIndex        =   14
      Top             =   60
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   820
      Controller      =   -1  'True
      ImageFile       =   "d:\tmp\get2.tif"
   End
   Begin UbbLVImg.UbbImageRegion imgRegiao 
      Height          =   1455
      Left            =   0
      TabIndex        =   13
      Top             =   780
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   2566
      Object.Top             =   79
      Right           =   80
      Controller      =   "imgControlador"
   End
   Begin UBBStatBar.UBBStatusBar stbBarra 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      Top             =   6315
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   450
   End
   Begin VB.Timer tmrObtencao 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   1860
      Top             =   60
   End
   Begin VB.Frame frmLinhaSup 
      BorderStyle     =   0  'None
      Height          =   1155
      Left            =   60
      TabIndex        =   12
      Top             =   3480
      Visible         =   0   'False
      Width           =   11655
      Begin UbbEdt.UbbEdit edtBanco 
         Height          =   840
         Left            =   1440
         TabIndex        =   4
         Top             =   120
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1482
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
            Name            =   "Courier New"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FieldType       =   3
         TextMaxNumChars =   3
         TitleAlignment  =   2
         Title           =   "Banco"
      End
      Begin UbbEdt.UbbEdit edtCompe 
         Height          =   840
         Left            =   360
         TabIndex        =   3
         Top             =   120
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1482
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
            Name            =   "Courier New"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FieldType       =   14
         TextMaxNumChars =   3
         TitleAlignment  =   2
         Title           =   "Comp."
      End
      Begin UbbEdt.UbbEdit edtAgencia 
         Height          =   840
         Left            =   2520
         TabIndex        =   5
         Top             =   120
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   1482
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
            Name            =   "Courier New"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FieldType       =   4
         TextMaxNumChars =   4
         TitleAlignment  =   2
         Title           =   "Agência"
      End
      Begin UbbEdt.UbbEdit edtC1 
         Height          =   840
         Left            =   3840
         TabIndex        =   6
         Top             =   120
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1482
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
            Name            =   "Courier New"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FieldType       =   8
         TextMaxNumChars =   1
         TitleAlignment  =   2
         Title           =   "C1"
      End
      Begin UbbEdt.UbbEdit edtConta 
         Height          =   840
         Left            =   4800
         TabIndex        =   7
         Top             =   120
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   1482
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
            Name            =   "Courier New"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FieldType       =   15
         TextMaxNumChars =   10
         TitleAlignment  =   2
         Title           =   "Conta"
      End
      Begin UbbEdt.UbbEdit edtC2 
         Height          =   840
         Left            =   7380
         TabIndex        =   8
         Top             =   120
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1482
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
            Name            =   "Courier New"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FieldType       =   9
         TextMaxNumChars =   1
         TitleAlignment  =   2
         Title           =   "C2"
      End
      Begin UbbEdt.UbbEdit edtCheque 
         Height          =   840
         Left            =   8340
         TabIndex        =   9
         Top             =   120
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   1482
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
            Name            =   "Courier New"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FieldType       =   17
         TextMaxNumChars =   6
         TitleAlignment  =   2
         Title           =   "Cheque"
      End
      Begin UbbEdt.UbbEdit edtC3 
         Height          =   840
         Left            =   10080
         TabIndex        =   10
         Top             =   120
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1482
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
            Name            =   "Courier New"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FieldType       =   10
         TextMaxNumChars =   1
         TitleAlignment  =   2
         Title           =   "C3"
         AutoNextControl =   0   'False
      End
   End
   Begin UBBFtp.UbbFtpRexec ftpClient 
      Left            =   120
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   979
   End
   Begin VB.Frame frmLinhaInf 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   0
      TabIndex        =   11
      Top             =   2280
      Visible         =   0   'False
      Width           =   11655
      Begin UbbEdt.UbbEdit edtInf1 
         Height          =   840
         Left            =   1020
         TabIndex        =   0
         Top             =   120
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   1482
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
            Name            =   "Courier New"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FieldType       =   11
         TitleAlignment  =   2
         Title           =   "Campo 1"
      End
      Begin UbbEdt.UbbEdit edtInf2 
         Height          =   840
         Left            =   3780
         TabIndex        =   1
         Top             =   120
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   1482
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
            Name            =   "Courier New"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FieldType       =   12
         TextMaxNumChars =   10
         TitleAlignment  =   2
         Title           =   "Campo 2"
      End
      Begin UbbEdt.UbbEdit edtInf3 
         Height          =   840
         Left            =   7020
         TabIndex        =   2
         Top             =   120
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   1482
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
            Name            =   "Courier New"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FieldType       =   13
         TextMaxNumChars =   12
         TitleAlignment  =   2
         Title           =   "Campo 3"
         AutoNextControl =   0   'False
      End
   End
   Begin UbbEdt.UBBValid valCmc7 
      Left            =   780
      Top             =   60
      _ExtentX        =   794
      _ExtentY        =   820
      ColorOK         =   12582912
      ColorInvalid    =   192
      Campo1          =   "edtInf1"
      Campo2          =   "edtInf2"
      Campo3          =   "edtInf3"
      Campo4          =   "edtCompe"
      Campo5          =   "edtBanco"
      Campo6          =   "edtAgencia"
      Campo7          =   "edtC1"
      Campo8          =   "edtConta"
      Campo9          =   "edtC2"
      Campo10         =   "edtCheque"
      Campo11         =   "edtC3"
   End
   Begin VB.Menu mnuCompl 
      Caption         =   "&Complementação"
      Begin VB.Menu mnuLinha 
         Caption         =   "&Linha Superior"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuVoltaDoc 
         Caption         =   "&Volta ao documento anterior"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSair 
         Caption         =   "&Sair"
      End
   End
   Begin VB.Menu mnuAjuda 
      Caption         =   "&Ajuda"
      Begin VB.Menu mnuAjudaSist 
         Caption         =   "Ajuda do Sistema"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuSepAjuda 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSobre 
         Caption         =   "So&bre"
         Shortcut        =   {F11}
      End
   End
End
Attribute VB_Name = "frmCmc7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' COMPCMC7     : Complementação de CMC-7                                    '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Autor        : Wladimir Leite                                             '
' Data         : 31/Outubro/2000                                            '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Descrição    : Aplicativo para complementação de CMC-7 de documentos.     '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

'Objeto CConfig
Private m_cfgAmb As bcci.CConfig

'Objeto CDatabase
Private m_cdbConn As bcci.CDatabase

'Objeto CApp
Private m_appImg As bcci.CApp

'Nome da Aplicação, Versão e Serviço
Private Const m_strAplicacao As String = "Complementação de CMC-7"
Private Const m_strDirLocal As String = "CompCMC7"
Private Const m_lngServico As Long = 3

'Digitação pela linha inferior (ou superior)
Private m_blnUsaLinhaInf As Boolean

'Buffer de documentos
Private Type Doc
    lngSeq As Long
    lngTipo As Long
    lngTipif As Long
    lngBanco As Long
    lngTipifOriginal As Long
    lngSeqMacroProc As Long
    strCampo(1 To 3) As String
    strC(1 To 3) As String
    vntRowid As Variant
    lngOperacao As Long
    lngEstado As Long
End Type
    
'Buffer de documentos
Private m_docBuffer() As Doc

'Documento Atual
Private m_lngDocAtual As Long

'Total de Documentos carregados
Private m_lngDocTotal As Long

'Conjunto obtido
Private m_lngSeqConjunto As Long

'Momento em que o usuario "pegou" o conjunto de documentos
Private m_vntDataInicio As Variant

'Novo Estado do documento quando é digitado mas não identificado
Private m_lngEstadoNaoIdentificado As Long

'Novo Estado do documento quando é digitado e identificado
Private m_lngEstadoIdentificado As Long

'Estados tratados
Private Const ESTADO_PARA_CMC7 = 35
Private Const ESTADO_EM_CMC7 = 36

'Transições
Private Const TRANSICAO_TRATADO = 2 'Tipo identificado
Private Const TRANSICAO_NAO_TRATADO = 3 'Tipo não identificado

'Operações executadas
Private Const OPERACAO_COMP_LINHA_SUP = 23
Private Const OPERACAO_COMP_LINHA_INF = 24
Private Const OPERACAO_EXCLUI_DOC = 25

'Timeout
Private m_lngTimeout As Long

'Função da API que posiciona o cursor
Private Declare Function SetCursorPos Lib "user32" _
        (ByVal x As Long, ByVal y As Long) As Long


'Carga do formulário
Private Sub Form_Load()
    Dim lgnLogin As bcci.CLogin
    Dim strVersao As String
    
    strVersao = Format$(App.Major, "0") & "." & _
                Format$(App.Minor, "00") & "." & _
                Format$(App.Revision, "00")
    
    If App.PrevInstance Then
        Unload Me
        Exit Sub
    End If
    
    Set m_cfgAmb = bcci.New_CConfig(m_strAplicacao, strVersao, m_lngServico, m_strDirLocal, Me.Icon)
    Set lgnLogin = bcci.New_CLogin(m_cfgAmb)
    
    If Not lgnLogin.Logou Then
        Unload Me
        Exit Sub
    End If
    
    Set m_appImg = bcci.New_CApp(Me, m_cfgAmb, stbBarra, ftpClient, imgControlador)
    Set m_cdbConn = bcci.New_CDatabase(lgnLogin.ADOConn, Me)
    
    'Posiciona o cursor do mouse (no canto superior esquerdo),
    'só para não atrapalhar a visualização dos controles
    SetCursorPos 50, 50
    
    'Acerta configuração da tela
    Show
    MudaLinha True
    
    'Carrega parametro iniciais
    Inicializa
    
    'Tenta obter documentos
    tmrObtencao_Timer

End Sub


'Usuário fechou o aplicativo
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If m_lngSeqConjunto > 0 Then
        DevolveDocumentos True
    End If
End Sub


'Ajusta Controles
Private Sub Form_Resize()
    Dim lngLeft As Long
    Dim lngY As Long
    
    lngY = (Me.ScaleHeight - stbBarra.Height - Screen.TwipsPerPixelY) \ 2
    If lngY > imgRegiao.Height Then
        lngLeft = (Me.ScaleWidth - frmLinhaInf.Width) \ 2
        If lngLeft < 0 Then lngLeft = 0
        frmLinhaInf.Move lngLeft, lngY
        frmLinhaSup.Move lngLeft, lngY
        imgRegiao.Move lngLeft, lngY - imgRegiao.Height
    End If
End Sub


'Menu de mudança de linha "sup/inf"
Private Sub mnuLinha_Click()
    MudaLinha (Not m_blnUsaLinhaInf)
End Sub


'Menu sair
Private Sub mnuSair_Click()
    tmrObtencao.Enabled = False
    Unload Me
End Sub


'Menu sobre
Private Sub mnuSobre_Click()
    bcci.New_CAbout m_cfgAmb
End Sub


'Muda linha inferior/superior
Private Sub MudaLinha(ByVal UsaLinhaInf As Boolean)
    m_blnUsaLinhaInf = UsaLinhaInf
    
    mnuLinha.Caption = IIf(m_blnUsaLinhaInf, _
                           "Linha Superior", _
                           "Linha Inferior")
                           
    If m_blnUsaLinhaInf Then
        frmLinhaSup.Visible = False
        frmLinhaInf.Visible = True
        imgRegiao.SetRect 0, 79, 80, 99
    Else
        frmLinhaSup.Visible = True
        frmLinhaInf.Visible = False
        imgRegiao.SetRect 0, 1, 80, 21
    
        MontaLinhaSuperior False
    End If
    
    AjustaFoco

End Sub


'Menu para voltar para o documento anterior
Private Sub mnuVoltaDoc_Click()
    VoltaDocumento
End Sub


'Evento Timer que indica periodicamente que a estação está ativa
Private Sub tmrEstacaoAtiva_Timer()
    Static lngTDecorrido As Long
    
    tmrEstacaoAtiva.Enabled = False
    
    'Acumula o tempo decorrido
    lngTDecorrido = lngTDecorrido + CLng(tmrEstacaoAtiva.Interval) \ 1000
    
    If lngTDecorrido >= m_lngTimeout Then
        If m_lngSeqConjunto > 0 Then
            If Not m_cdbConn.ExecSQL("update CONJUNTO set DATA_INICIO=sysdate where " & _
                                     "SEQ_CONJUNTO = ?", Array(m_lngSeqConjunto)) Then
                m_cdbConn.Rollback
            Else
                m_cdbConn.Commit
            End If
        End If
        
        'Zera o tempo decorrido
        lngTDecorrido = 0
    End If
    
    tmrEstacaoAtiva.Enabled = True
End Sub


'Evento Timer de Obtencao de documentos
Private Sub tmrObtencao_Timer()
    Static intN As Integer
    
    'Desabilita timer
    tmrObtencao.Enabled = False
    
    stbBarra.SetStatus "Procurando documentos...", 0, 0
    If ObtemDocumentos Then
        m_lngDocAtual = 1
        HabilitaEdicao True
        CarregaDocumento False
        
        'Obtem o momento de início do tratamento do conjunto
        m_vntDataInicio = Now()
        
        'Habilita o Timer que indicará que estação está ativa
        tmrEstacaoAtiva.Enabled = True
    Else
        stbBarra.SetStatus Space$(intN * 2) & "Aguardando documentos...", 0, 0
        intN = intN + 1
        If intN = 10 Then intN = 0
        
        HabilitaEdicao False
        tmrObtencao.Enabled = True
    End If
End Sub


'Carrega documentos que serão tratados
Private Function ObtemDocumentos() As Boolean
    Dim lngNDocs As Long
    Dim lngParamArr(0 To 7) As Long
    Dim strSql As String
    Dim vntArr As Variant
    
    'Inicializa total de documentos carregados
    m_lngDocTotal = 0
    
    'Preparar parâmetros para obtenção de Conjuntos
    lngParamArr(1) = ESTADO_PARA_CMC7  'Estado origem
    lngParamArr(2) = ESTADO_EM_CMC7    'Estado_Destino
    lngParamArr(3) = m_cfgAmb.SeqUsu   'Seq_Usu
    lngParamArr(4) = m_cfgAmb.Estacao  'Seq_Estação
    lngParamArr(5) = 0                 'Banco Apresentante
    lngParamArr(6) = 0                 'Processo
    lngParamArr(7) = 0                 'Tipo_Exceção
            
    'Chama função de obtenção de conjuntos
    m_cdbConn.ExecSQL "{? = call F_OBTEM_CONJUNTO (?, ?, ?, ?, ?, ?, ?)}", _
                      lngParamArr
            
    'Achou ?
    If lngParamArr(0) <= 0 Then
        ObtemDocumentos = False
        Exit Function
    End If
    
    'Conjunto obtido
    m_lngSeqConjunto = lngParamArr(0)
    
    'Le documentos pertencentes ao conjunto da base de dados
    strSql = "select SEQ_DOC, SEQ_MACRO_PROC, " & _
             "CAMPO1_CMC7, CAMPO2_CMC7, CAMPO3_CMC7, ROWID from CAPTURA " & _
             "where SEQ_CONJUNTO = ?"
             
    If Not m_cdbConn.ExecSQL(strSql, Array(m_lngSeqConjunto), vntArr, m_lngDocTotal) Then
        ObtemDocumentos = False
        Exit Function
    End If

    If m_lngDocTotal <= 0 Then
        'Apenas "libera" o conjunto vazio
        If LiberaConjunto Then
            m_cdbConn.Commit
        Else
            m_cdbConn.Rollback
        End If
        ObtemDocumentos = False
        Exit Function
    End If
    
    'Armazena documentos lidos
    ReDim m_docBuffer(1 To m_lngDocTotal)

    For lngNDocs = 0 To (m_lngDocTotal - 1)
        With m_docBuffer(lngNDocs + 1)
            .lngSeq = vntArr(0, lngNDocs)
        
            .lngTipo = 0
            
            .lngSeqMacroProc = vntArr(1, lngNDocs)
            
            If CDec(vntArr(2, lngNDocs)) = CDec(0) Then
                .strCampo(1) = ""
            Else
                .strCampo(1) = Format$(vntArr(3, lngNDocs), "00000000")
            End If
            
            If CDec(vntArr(3, lngNDocs)) = CDec(0) Then
                .strCampo(2) = ""
            Else
                .strCampo(2) = Format$(vntArr(3, lngNDocs), "0000000000")
            End If
            
            If CDec(vntArr(4, lngNDocs)) = CDec(0) Then
                .strCampo(3) = ""
            Else
                .strCampo(3) = Format$(vntArr(4, lngNDocs), "000000000000")
            End If
        
            .strC(1) = ""
            .strC(2) = ""
            .strC(3) = ""
            
            .vntRowid = vntArr(5, lngNDocs)
            
            'Guarda o tipif, se já houver
            If Len(.strCampo(2)) = 10 Then
                .lngTipif = CLng(Mid$(.strCampo(2), 10, 1))
            Else
                .lngTipif = 0
            End If
            .lngTipifOriginal = .lngTipif
        End With
    Next

    ObtemDocumentos = True
End Function


'Carrega o documento atual
Private Sub CarregaDocumento(Optional ByVal blnModoNormal As Boolean = True)
    'Atualiza a barra
    stbBarra.SetStatus "Documento: " & CStr(m_docBuffer(m_lngDocAtual).lngSeq) & _
                       "  (" & CStr(m_lngDocAtual) & "/" & CStr(m_lngDocTotal) & ")", _
                       m_lngDocAtual, _
                       m_lngDocTotal
    
    If blnModoNormal Then
        'Normalmente já estará no buffer
        imgControlador.UseBackImage
    Else
        'Senão carrega
        imgControlador.LoadDoc m_docBuffer(m_lngDocAtual).lngSeq, _
                               m_docBuffer(m_lngDocAtual).lngTipo
    End If
        
    'Preenche o CMC7 linha Inferior
    edtInf1.Text = m_docBuffer(m_lngDocAtual).strCampo(1)
    edtInf2.Text = m_docBuffer(m_lngDocAtual).strCampo(2)
    edtInf3.Text = m_docBuffer(m_lngDocAtual).strCampo(3)
    MontaLinhaSuperior True
    AjustaFoco
    
    'Tem mais documentos a seguir
    If m_lngDocAtual < m_lngDocTotal Then
        'Traz próxima imagem "em background"
        imgControlador.LoadDoc m_docBuffer(m_lngDocAtual + 1).lngSeq, _
                               m_docBuffer(m_lngDocAtual + 1).lngTipo, _
                               True
    End If
End Sub


'Habilita/desabilita todos edits
Private Sub HabilitaEdicao(blnEnabled As Boolean)
    edtInf1.Enabled = blnEnabled
    edtInf2.Enabled = blnEnabled
    edtInf3.Enabled = blnEnabled
    
    edtCompe.Enabled = blnEnabled
    edtBanco.Enabled = blnEnabled
    edtAgencia.Enabled = blnEnabled
    edtC1.Enabled = blnEnabled
    
    edtConta.Enabled = blnEnabled
    edtC2.Enabled = blnEnabled

    edtCheque.Enabled = blnEnabled
    edtC3.Enabled = blnEnabled
    
    'Também limpa os campos
    If Not blnEnabled Then
        edtInf1.Text = ""
        edtInf2.Text = ""
        edtInf3.Text = ""
    
        edtCompe.Text = ""
        edtBanco.Text = ""
        edtAgencia.Text = ""
        edtC1.Text = ""
        
        edtConta.Text = ""
        edtC2.Text = ""
    
        edtCheque.Text = ""
        edtC3.Text = ""
        
        imgControlador.LoadDoc 0
    End If
End Sub


'Trata PAGEUP
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    'Volta ao documento anterior
    If KeyCode = vbKeyPageUp Then
        VoltaDocumento
    End If
End Sub


'Trata tecla enter nos campos "finais"
Private Sub Form_KeyPress(KeyAscii As Integer)
    Static strSupAnterior(1 To 3) As String
    Static lngDocAnterior As Long
    Dim strSupAtual(1 To 3) As String
    Dim blnOk As Boolean
    
    If (KeyAscii = vbKeyReturn) And _
        Not (ActiveControl Is Nothing) Then
        'Não permite passar para o próximo campo
        '(com enter) se não está preenchido
        On Error Resume Next
        If Len(ActiveControl.Text) = 0 Then
            KeyAscii = 0
            Exit Sub
        End If
        On Error GoTo 0
        
        'Digitação pela linha inferior
        If (ActiveControl.Name = edtInf3.Name) And _
           (m_blnUsaLinhaInf) Then
            KeyAscii = 0
            If valCmc7.ValidaLinhaInf = CustomValidStatus.Invalido Then
                MsgBox "Linha inferior inválida !", vbExclamation
                AjustaFoco
            Else
                m_docBuffer(m_lngDocAtual).lngOperacao = OPERACAO_COMP_LINHA_INF
                m_docBuffer(m_lngDocAtual).lngEstado = m_lngEstadoNaoIdentificado
                
                'Linha inferior zerada ?
                If valCmc7.ValidaLinhaInf = CustomValidStatus.Zerado Then
                    'Identificar tipo de documento
                    IdentTipo True, True
                    
                    m_docBuffer(m_lngDocAtual).lngEstado = m_lngEstadoIdentificado
                    
                    If m_docBuffer(m_lngDocAtual).lngTipo = 0 Then
                        m_docBuffer(m_lngDocAtual).lngOperacao = OPERACAO_EXCLUI_DOC
                    End If
                End If
                
                ProximoDocumento
            End If
        ElseIf (ActiveControl.Name = edtC3.Name) And _
           (Not m_blnUsaLinhaInf) Then
            'Linha inferior
            KeyAscii = 0
            
            'O documento anterior e mesmo que o atual
            If m_lngDocAtual <> lngDocAnterior Then
                strSupAnterior(1) = ""
                strSupAnterior(2) = ""
                strSupAnterior(3) = ""
            End If
            'Guarda para a proxima tentativa
            lngDocAnterior = m_lngDocAtual
            
            blnOk = True
            
            'Obtem linha superior digitada
            strSupAtual(1) = edtCompe(FilledNotNull) & edtBanco(FilledNotNull) & edtAgencia(FilledNotNull) & edtC1(FilledNotNull)
            strSupAtual(2) = edtConta(FilledNotNull) & edtC2(FilledNotNull)
            strSupAtual(3) = edtCheque(FilledNotNull) & edtC3(FilledNotNull)
            
            'Campo 1
            If blnOk And (valCmc7.ValidaLinhaSup(1) = CustomValidStatus.Invalido) Then
                'Redigitação bateu ?
                If strSupAtual(1) = strSupAnterior(1) Then
                    If MsgBox("Campo 1 inválido, confirma redigitação ?", vbQuestion Or vbYesNo, m_strAplicacao) <> vbYes Then
                        AjustaFoco 1
                        blnOk = False
                    End If
                Else
                    MsgBox "Campo 1 inválido !", vbExclamation
                    AjustaFoco 1
                    blnOk = False
                End If
            End If
            
            'Campo 2
            If blnOk And (valCmc7.ValidaLinhaSup(2) = CustomValidStatus.Invalido) Then
                'Redigitação bateu ?
                If strSupAtual(2) = strSupAnterior(2) Then
                    If MsgBox("Campo 2 inválido, confirma redigitação ?", vbQuestion Or vbYesNo, m_strAplicacao) <> vbYes Then
                        AjustaFoco 2
                        blnOk = False
                    End If
                Else
                    MsgBox "Campo 2 inválido !", vbExclamation
                    AjustaFoco 2
                    blnOk = False
                End If
            End If
            
            'Campo 3
            If blnOk And (valCmc7.ValidaLinhaSup(3) = CustomValidStatus.Invalido) Then
                'Redigitação bateu ?
                If strSupAtual(3) = strSupAnterior(3) Then
                    If MsgBox("Campo 3 inválido, confirma redigitação ?", vbQuestion Or vbYesNo, m_strAplicacao) <> vbYes Then
                        AjustaFoco 3
                        blnOk = False
                    End If
                Else
                    MsgBox "Campo 3 inválido !", vbExclamation
                    AjustaFoco 3
                    blnOk = False
                End If
            End If
            
            
            If blnOk Then
                m_docBuffer(m_lngDocAtual).lngOperacao = OPERACAO_COMP_LINHA_SUP
                m_docBuffer(m_lngDocAtual).lngEstado = m_lngEstadoNaoIdentificado
                
                'Linha superior parcialmente zerada ?
                If (valCmc7.ValidaLinhaSup(1) = CustomValidStatus.Zerado) Or _
                   (valCmc7.ValidaLinhaSup(2) = CustomValidStatus.Zerado) Or _
                   (valCmc7.ValidaLinhaSup(3) = CustomValidStatus.Zerado) Then
                    'Identificar tipo de documento e tipif
                    IdentTipo True, True
                    
                    m_docBuffer(m_lngDocAtual).lngEstado = m_lngEstadoIdentificado
                    
                    If m_docBuffer(m_lngDocAtual).lngTipo = 0 Then
                        m_docBuffer(m_lngDocAtual).lngOperacao = OPERACAO_EXCLUI_DOC
                        
                        'Zera todos campos
                        edtCompe = "0"
                        edtBanco = "0"
                        edtAgencia = "0"
                        edtC1 = "0"
                        edtConta = "0"
                        edtC2 = "0"
                        edtCheque = "0"
                        edtC3 = "0"
                    End If
                Else
                    'Identificar tipif
                    If (m_docBuffer(m_lngDocAtual).lngTipif = 0) Then
                        IdentTipo False, True
                    End If
                End If
                
                ProximoDocumento
            Else
                'Guarda a linha superior digitada
                strSupAnterior(1) = strSupAtual(1)
                strSupAnterior(2) = strSupAtual(2)
                strSupAnterior(3) = strSupAtual(3)
            End If
        End If
    End If
End Sub


'Acerta o foco
Private Sub AjustaFoco(Optional ByVal intCampo As Integer = 0)
    If m_blnUsaLinhaInf Then
        'Foco no primeiro campo vazio
        If Not Foco(edtInf1) Then
            If Not Foco(edtInf2) Then
                If (edtInf3.Visible) And (edtInf3.Enabled) Then edtInf3.SetFocus
            End If
        End If
    Else
        Select Case intCampo
            Case 1
                If (edtCompe.Visible) And _
                   (edtCompe.Enabled) Then edtCompe.SetFocus
                edtCompe.Text = ""
                edtBanco.Text = ""
                edtAgencia.Text = ""
                edtC1.Text = ""
            Case 2
                If (edtConta.Visible) And _
                   (edtConta.Enabled) Then edtConta.SetFocus
                edtConta.Text = ""
                edtC2.Text = ""
            Case 3
                If (edtCheque.Visible) And _
                   (edtCheque.Enabled) Then edtCheque.SetFocus
                edtCheque.Text = ""
                edtC3.Text = ""
            Case Else
                'Foco no primeiro campo vazio
                If Not Foco(edtCompe) Then
                    If Not Foco(edtBanco) Then
                        If Not Foco(edtC1) Then
                            If Not Foco(edtConta) Then
                                If Not Foco(edtC2) Then
                                    If Not Foco(edtCheque) Then
                                        If (edtC3.Visible) And (edtC3.Enabled) Then edtC3.SetFocus
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
        End Select
    End If
End Sub


'Passa para próximo documento
Private Sub ProximoDocumento()
    Dim lngTipif As Long
    Dim lngDv1 As Long
    Dim lngDv2 As Long
    Dim lngDv3 As Long
    
    If m_blnUsaLinhaInf Then
        'Salva o CMC7 da linha Inferior
        m_docBuffer(m_lngDocAtual).strCampo(1) = edtInf1.Text(FilledNotNull)
        m_docBuffer(m_lngDocAtual).strCampo(2) = edtInf2.Text(FilledNotNull)
        m_docBuffer(m_lngDocAtual).strCampo(3) = edtInf3.Text(FilledNotNull)
        
    Else
        'Salva C1,C2,C3 para o caso de "voltar" o documento
        m_docBuffer(m_lngDocAtual).strC(1) = edtC1.Text(Filled)
        m_docBuffer(m_lngDocAtual).strC(2) = edtC2.Text(Filled)
        m_docBuffer(m_lngDocAtual).strC(3) = edtC3.Text(Filled)
        
        'Monta os campos a partir da linha superior
        'Calcular DV's
        valCmc7.CalcDVsSupInf m_docBuffer(m_lngDocAtual).lngTipif, _
                              lngDv1, lngDv2, lngDv3

        'Campo1
        m_docBuffer(m_lngDocAtual).strCampo(1) = edtBanco.Text(FilledNotNull) & _
                                                 edtAgencia.Text(FilledNotNull) & _
                                                 Format$(lngDv1, "0")
                
        'Campo2
        m_docBuffer(m_lngDocAtual).strCampo(2) = edtCompe.Text(FilledNotNull) & _
                                                 edtCheque.Text(FilledNotNull) & _
                                                 Format$(m_docBuffer(m_lngDocAtual).lngTipif, "0")
                
        'Campo3
        m_docBuffer(m_lngDocAtual).strCampo(3) = Format$(lngDv2, "0") & _
                                                 edtConta.Text(FilledNotNull) & _
                                                 Format$(lngDv3, "0")
    
    End If
    
    'incrementa documento atual
    m_lngDocAtual = m_lngDocAtual + 1
    
    'era o último ?
    If m_lngDocAtual > m_lngDocTotal Then
        DevolveDocumentos
    Else
        CarregaDocumento
    End If
End Sub


'Volta ao documento anterior
Private Sub VoltaDocumento()
    If m_lngDocAtual > 1 Then
        m_lngDocAtual = m_lngDocAtual - 1
        
        'Restaura tipif e tipo originais
        m_docBuffer(m_lngDocAtual).lngTipif = m_docBuffer(m_lngDocAtual).lngTipifOriginal
        m_docBuffer(m_lngDocAtual).lngTipo = 0
        
        CarregaDocumento False
    Else
        Beep
    End If
End Sub


'Ao término da digitação dos documentos do buffer,
'ou por "finalização do aplicativo"
Private Sub DevolveDocumentos(Optional ByVal blnFinalizaApp As Boolean = False)
    Dim blnSucesso As Boolean
    
    blnSucesso = True
    
    'Desabilita o Timer que indica que estação está ativa
    tmrEstacaoAtiva.Enabled = False
    
    'Existem documentos carregados
    If m_lngDocTotal > 0 Then
        'Grava atualizações no banco de dados
        blnSucesso = GravaAlteracoes
        
        If blnSucesso Then
            'Grava informações de produtividade
            blnSucesso = GravaProdutividade
        End If
        
        'Limpa diretório local
        On Error Resume Next
        Kill (m_cfgAmb.DirLocal & "\*.*")
        On Error GoTo 0
    End If
    
    'Libera o conjunto
    If blnSucesso Then
        blnSucesso = LiberaConjunto
    End If
    
    'Atualizações realizadas com sucesso ?
    If blnSucesso Then
        m_cdbConn.Commit
    Else
        m_cdbConn.Rollback
    End If
    
    m_lngSeqConjunto = 0
    
    If Not blnFinalizaApp Then
        'Tenta pegar mais documentos
        tmrObtencao_Timer
    End If
End Sub


'Preenche campos da linha superior
Private Sub MontaLinhaSuperior(ByVal blnLimpa As Boolean)
    Dim strAux As String
        
    'Campo 1
    strAux = edtInf1.Text(FilledNotNull)
    If strAux <> String(Len(strAux), "0") Then
        edtBanco.Text = Mid$(strAux, 1, 3)
        edtAgencia.Text = Mid$(strAux, 4, 4)
    ElseIf blnLimpa Then
        edtBanco.Text = ""
        edtAgencia.Text = ""
        edtC1.Text = ""
    End If

    'Campo 2
    strAux = edtInf2.Text(FilledNotNull)
    If strAux <> String(Len(strAux), "0") Then
        edtCompe.Text = Mid$(strAux, 1, 3)
        edtCheque.Text = Mid$(strAux, 4, 6)
    ElseIf blnLimpa Then
        edtCompe.Text = ""
        edtCheque.Text = ""
    End If

    'Campo 3
    strAux = edtInf3.Text(FilledNotNull)
    If strAux <> String(Len(strAux), "0") Then
        edtConta.Text = Mid$(strAux, 5, 7)
    ElseIf blnLimpa Then
        edtConta.Text = ""
    End If

    'Digitos da linha superior
    If blnLimpa Then
        edtC1.Text = m_docBuffer(m_lngDocAtual).strC(1)
        edtC2.Text = m_docBuffer(m_lngDocAtual).strC(2)
        edtC3.Text = m_docBuffer(m_lngDocAtual).strC(3)
    End If

End Sub


'Identifica tipo de documento, banco (apresentante) e tipif
Private Sub IdentTipo(ByVal blnTipoDoc As Boolean, _
                      ByVal blnTipif As Boolean)
    'Algo a fazer ?
    If (Not blnTipoDoc) And (Not blnTipif) Then Exit Sub
    
    frmIdentTipo.SetParam ftpClient, _
                          m_cfgAmb, _
                          m_cdbConn, _
                          m_docBuffer(m_lngDocAtual).lngSeq, _
                          m_docBuffer(m_lngDocAtual).lngSeqMacroProc, _
                          blnTipoDoc, _
                          blnTipif
    If blnTipoDoc Then
        m_docBuffer(m_lngDocAtual).lngTipo = frmIdentTipo.GetTipoDoc
        m_docBuffer(m_lngDocAtual).lngBanco = frmIdentTipo.GetBanco
    End If
    
    If blnTipif Then
        m_docBuffer(m_lngDocAtual).lngTipif = frmIdentTipo.GetTipif
    End If
End Sub


'Carrega parâmetros iniciais do aplicativo
Private Sub Inicializa()
    Dim strSleep As String
    Dim vntArr As Variant
    Dim lngLin As Long
    
    'Timeout, em segundos
    m_lngTimeout = 100
    On Error Resume Next
    m_lngTimeout = CLng(m_cdbConn.GetParameter("SISTEMA", "TIMEOUT")) \ 2
    On Error GoTo 0
    
    'Intervalo entre execuções
    strSleep = m_cdbConn.GetParameter("COMPL_CMC7", "SLEEP")
    
    On Error Resume Next
    tmrObtencao.Interval = CLng(strSleep) * 1000
    On Error GoTo 0

    'Checar qual o estado futuro (TRATADO)
    m_cdbConn.ExecSQL "select SEQ_ESTADO_DESTINO from TRANSICAO_DOCUMENTO where SEQ_ESTADO_ORIGEM = ? and TIPO_TRANSICAO = ?", _
                      Array(ESTADO_EM_CMC7, TRANSICAO_TRATADO), _
                      vntArr, lngLin
    If lngLin <> 1 Then
        MsgBox "ERRO: Não foi possível carregar Transição de Estados.", vbCritical, m_strAplicacao
        End
    End If
    
    m_lngEstadoIdentificado = vntArr(0, 0)

    'Checar qual o estado futuro (NÃO TRATADO)
    m_cdbConn.ExecSQL "select SEQ_ESTADO_DESTINO from TRANSICAO_DOCUMENTO where SEQ_ESTADO_ORIGEM = ? and TIPO_TRANSICAO = ?", _
                      Array(ESTADO_EM_CMC7, TRANSICAO_NAO_TRATADO), _
                      vntArr, lngLin
    If lngLin <> 1 Then
        MsgBox "ERRO: Não foi possível carregar Transição de Estados.", vbCritical, m_strAplicacao
        End
    End If
    
    m_lngEstadoNaoIdentificado = vntArr(0, 0)
End Sub


'Grava atualizações no banco de dados
Private Function GravaAlteracoes() As Boolean
    Dim lngN As Long
    
    'Percorre os documentos já tratados
    For lngN = 1 To (m_lngDocAtual - 1)
        With m_docBuffer(lngN)
            'Trata tipif e banco (identificados manualmente)
            If .lngBanco > 0 Then
                'Força o banco no campo 1
                .strCampo(1) = Format$(.lngBanco, "000") & Mid$(.strCampo(1), 4)
            End If
            If (.lngTipif > 0) And (Right$(.strCampo(2), 1) = "0") Then
                'Força o tipif no campo 2
                .strCampo(2) = Mid$(.strCampo(2), 1, 9) & Format$(.lngTipif, "0")
            End If
            
            'Atualiza tabela da captura
            If Not (m_cdbConn.ExecSQL("update CAPTURA set CAMPO1_CMC7 = ?, " & _
                                      "CAMPO2_CMC7 = ?, CAMPO3_CMC7 = ?, " & _
                                      "TIPO_DOC = DECODE (?, 0, TIPO_DOC, ?), EST_CMC7 = ?, SEQ_CONJUNTO = 0 where " & _
                                      "ROWID = ? and SEQ_CONJUNTO = ?", _
                                      Array(.strCampo(1), .strCampo(2), .strCampo(3), _
                                            .lngTipo, .lngTipo, .lngEstado, .vntRowid, m_lngSeqConjunto))) Then
                GravaAlteracoes = False
                Exit Function
            End If
        End With
    Next
        
    For lngN = 1 To (m_lngDocAtual - 1)
        With m_docBuffer(lngN)
            'Audita
            If Not (m_cdbConn.Audit(.lngSeq, m_cfgAmb.SeqUsu, .lngOperacao, m_cfgAmb.Estacao)) Then
                GravaAlteracoes = False
                Exit Function
            End If
        End With
    Next

    'Percorre os documentos não tratados
    For lngN = m_lngDocAtual To m_lngDocTotal
        With m_docBuffer(lngN)
            'Apenas desfaz o conjunto
            If Not (m_cdbConn.ExecSQL("update CAPTURA set SEQ_CONJUNTO = 0 where " & _
                                      "ROWID = ? and SEQ_CONJUNTO = ?", _
                                      Array(.vntRowid, m_lngSeqConjunto))) Then
                GravaAlteracoes = False
                Exit Function
            End If
        End With
    Next

    GravaAlteracoes = True
End Function


'Apaga o conjunto já tratado
Private Function LiberaConjunto() As Boolean
    If m_lngSeqConjunto > 0 Then
        LiberaConjunto = m_cdbConn.ExecSQL("delete CONJUNTO where SEQ_CONJUNTO = ?", _
                                           Array(m_lngSeqConjunto))
    Else
        LiberaConjunto = True
    End If
End Function


'Coloca o foco num Edit, se ele está vazio
Private Function Foco(ByRef edtCampo As UbbEdit) As Boolean
    Foco = False
    If Len(edtCampo.Text(Normal)) = 0 Then
        If (edtCampo.Visible) And (edtCampo.Enabled) Then
            edtCampo.SetFocus
            Foco = True
        End If
    End If
End Function


'Grava informações de produtividade no banco de dados
Private Function GravaProdutividade() As Boolean
    Dim lngN As Long
    Dim strSql As String
    Dim lngCount As Long
    Dim lngTempoTotal As Long
    Dim lngOperacao As Long
    Dim lngAux As Long
        
    'Retorno default
    GravaProdutividade = True
    
    'Tempo transcorrido
    lngTempoTotal = DateDiff("s", m_vntDataInicio, Now())
    
    'Tempo não deve ser negativo !
    If lngTempoTotal < 0 Then
        MsgBox "Problema na obtenção do tempo decorrido !", vbCritical, m_strAplicacao
        
        'Mesmo assim considera OK para o resto das atualizações
        Exit Function
    End If
    
    'Inserção na tabela PRODUTIVIDADE
    strSql = "insert into PRODUTIVIDADE (SEQ_USU, SEQ_SERV, SEQ_SUBBLOCO, " & _
             "SEQ_PROC, SEQ_BANCO, SEQ_OPERACAO, TIPO_EXCECAO, QTDE_ALT, " & _
             "QTDE_VIS, QTDE_TOT, TEMPO) values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"

    'Percorre as operações tratadas
    For lngAux = 1 To 3
        Select Case lngAux
            Case 1
                lngOperacao = OPERACAO_COMP_LINHA_SUP
            Case 2
                lngOperacao = OPERACAO_COMP_LINHA_INF
            Case 3
                lngOperacao = OPERACAO_EXCLUI_DOC
        End Select
        
        lngCount = 0
        For lngN = 1 To (m_lngDocAtual - 1)
            If m_docBuffer(lngN).lngOperacao = lngOperacao Then
                lngCount = lngCount + 1
            End If
        Next
        
        'Pelo menos um documento ?
        If (lngCount > 0) Then
            'Insere efetivamente
            GravaProdutividade = m_cdbConn.ExecSQL(strSql, _
                                 Array(m_cfgAmb.SeqUsu, _
                                       m_lngServico, _
                                       0, _
                                       0, _
                                       0, _
                                       lngOperacao, _
                                       0, _
                                       lngCount, _
                                       lngCount, _
                                       lngCount, _
                                       Round(CDbl(lngTempoTotal) * CDbl(lngCount) / CDbl(m_lngDocAtual - 1))))
            
            'Erro ?
            If Not GravaProdutividade Then Exit For
        End If
    Next
End Function


