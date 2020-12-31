VERSION 5.00
Object = "{00080003-B1BA-11CE-ABC6-F5B2E79D9E3F}#8.0#0"; "LTOCX80N.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ConfirmaAgenciaConta 
   Caption         =   "Confirmação de Agência e Conta"
   ClientHeight    =   7176
   ClientLeft      =   48
   ClientTop       =   300
   ClientWidth     =   10248
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7176
   ScaleWidth      =   10248
   StartUpPosition =   1  'CenterOwner
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   10245
      _ExtentX        =   18076
      _ExtentY        =   614
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   5284
            MinWidth        =   1834
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   8108
            MinWidth        =   4658
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   4580
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
   Begin VB.Timer tmrAtualiza 
      Enabled         =   0   'False
      Interval        =   50000
      Left            =   7776
      Top             =   1344
   End
   Begin VB.Timer tmrPesquisa 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7440
      Top             =   1344
   End
   Begin VB.Frame frmAgConta 
      Height          =   876
      Left            =   4032
      TabIndex        =   15
      Top             =   5136
      Width           =   2172
      Begin VB.TextBox txtAgencia 
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
         Left            =   240
         MaxLength       =   4
         TabIndex        =   0
         Top             =   384
         Width           =   612
      End
      Begin VB.TextBox txtConta 
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
         Left            =   996
         MaxLength       =   7
         TabIndex        =   1
         Top             =   384
         Width           =   960
      End
      Begin VB.Label lblConta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Conta"
         Height          =   192
         Left            =   1008
         TabIndex        =   22
         Top             =   192
         Width           =   420
      End
      Begin VB.Label lblAgencia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Agência"
         Height          =   192
         Left            =   240
         TabIndex        =   21
         Top             =   192
         Width           =   600
      End
   End
   Begin VB.Frame fraBotoesSuperiores 
      Height          =   1068
      Left            =   6480
      TabIndex        =   11
      Top             =   6000
      Width           =   3660
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
         Left            =   960
         Picture         =   "ConfirmaAgenciaConta.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   192
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
         Left            =   96
         Picture         =   "ConfirmaAgenciaConta.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   192
         Width           =   850
      End
      Begin VB.CommandButton cmdSupervisor 
         Caption         =   "Docto Ilegível"
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
         Left            =   1824
         Picture         =   "ConfirmaAgenciaConta.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   192
         Width           =   850
      End
      Begin VB.CommandButton cmdFinalizar 
         Cancel          =   -1  'True
         Caption         =   "Finalizar Digitação"
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
         Left            =   2688
         Picture         =   "ConfirmaAgenciaConta.frx":0B8E
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   192
         Width           =   850
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1068
      Left            =   96
      TabIndex        =   4
      Top             =   6000
      Width           =   5388
      Begin VB.CommandButton cmdRotacao 
         Caption         =   "Rotação"
         Height          =   800
         Left            =   1824
         Picture         =   "ConfirmaAgenciaConta.frx":0E98
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   192
         Width           =   850
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
         Left            =   2688
         Picture         =   "ConfirmaAgenciaConta.frx":11A2
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   192
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
         Left            =   3552
         Picture         =   "ConfirmaAgenciaConta.frx":14AC
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   192
         Width           =   852
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
         Left            =   96
         Picture         =   "ConfirmaAgenciaConta.frx":17B6
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   192
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
         Left            =   960
         Picture         =   "ConfirmaAgenciaConta.frx":1AC0
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   192
         Width           =   850
      End
      Begin VB.CommandButton cmdConfirmar 
         Caption         =   "Confirmar"
         Height          =   800
         Left            =   4416
         Picture         =   "ConfirmaAgenciaConta.frx":1DCA
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   192
         Width           =   850
      End
   End
   Begin VB.Frame FrmImagem 
      Caption         =   "Imagem"
      Height          =   4740
      Left            =   78
      TabIndex        =   2
      Top             =   384
      Width           =   10080
      Begin LeadLib.Lead Lead1 
         Height          =   4452
         Left            =   96
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   216
         Width           =   9876
         _Version        =   524288
         _ExtentX        =   17420
         _ExtentY        =   7853
         _StockProps     =   229
         BackColor       =   16777215
         BorderStyle     =   1
         ScaleHeight     =   369
         ScaleWidth      =   821
         DataField       =   ""
         BitmapDataPath  =   ""
         AnnDataPath     =   ""
      End
   End
   Begin VB.PictureBox FrmPesquisa 
      Height          =   1932
      Left            =   2232
      ScaleHeight     =   1884
      ScaleWidth      =   5724
      TabIndex        =   17
      Top             =   2568
      Visible         =   0   'False
      Width           =   5772
      Begin VB.CommandButton CmdFecharPesquisa 
         Caption         =   "&Fechar"
         Height          =   312
         Left            =   2376
         TabIndex        =   20
         Top             =   1464
         Width           =   1068
      End
      Begin ComctlLib.ProgressBar Progress 
         Height          =   300
         Left            =   336
         TabIndex        =   18
         Top             =   912
         Width           =   5028
         _ExtentX        =   8869
         _ExtentY        =   529
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Pesquisando por Documentos para confirmação. Aguarde ..."
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
         TabIndex        =   19
         Top             =   576
         Width           =   5112
      End
   End
End
Attribute VB_Name = "ConfirmaAgenciaConta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private qryGetCapaConfAgConta                   As RDO.rdoQuery
Private qryAtualizaDocumentoConfAgConta         As RDO.rdoQuery
Private qryGetDocumentoConfAgConta              As RDO.rdoQuery

Private aCapa()                                 As tpCapa
Private aDoc()                                  As TpDocumento
Private PrimeiraVez                             As Boolean
Private AlterouDocto                            As Boolean
Private sTempo                                  As Integer
Private m_IndexCapa                             As Integer
Private m_IndexDoc                              As Integer
Private m_EnviarParaIlegiveis                   As Boolean
Private teclou                                  As Boolean

Private Enum ePosicao
    eAnterior
    ePosterior
End Enum

Private Enum eEnviarCapaPara
    eIlegiveis = 5
    eVinculoAutomatico = 8
End Enum

Private Const HABILITADO = &H80000005               'Window Background
Private Const DESABILITADO = &H8000000F             'Button Face
Private Const STATUS_TXT_DOCUMENTO = "Documento :"
Private Const STATUS_TXT_LOTECAPA = "Lote/Capa :"
Private Const STATUS_TXT_DATA = "Data :"

Private Sub Form_Load()

    Set qryGetDocumentoConfAgConta = Geral.Banco.CreateQuery("", "{Call GetDocumentoConfAgConta(?,?)}")
    Set qryAtualizaDocumentoConfAgConta = Geral.Banco.CreateQuery("", "{? = call AtualizaDocumentoConfAgConta (?,?,?,?,?,?)}")
    Set qryGetCapaConfAgConta = Geral.Banco.CreateQuery("", "{call GetCapaConfAgConta(?,?)}")
        
    Cabecalho False

    PrimeiraVez = True
End Sub

Private Sub Form_Unload(Cancel As Integer)

    qryGetCapaConfAgConta.Close
    qryGetDocumentoConfAgConta.Close
    qryAtualizaDocumentoConfAgConta.Close

End Sub

Private Sub Cabecalho(ByVal pShow As Boolean)

    Dim sDocumento          As String
    Dim sLoteCapa           As String

    
    sDocumento = STATUS_TXT_DOCUMENTO & " "
    sLoteCapa = STATUS_TXT_LOTECAPA & " "
    
    If pShow Then

        sDocumento = sDocumento & m_IndexDoc + 1 & "/" & UBound(aDoc) + 1
        sLoteCapa = sLoteCapa & Format(aCapa(m_IndexCapa).IdLote, "0000-00000") & " / " & aCapa(m_IndexCapa).Capa

    End If
    
    StatusBar1.Panels(1).Text = sDocumento
    StatusBar1.Panels(2).Text = sLoteCapa
    
    
    StatusBar1.Panels(3).Text = STATUS_TXT_DATA & Format(Format(Geral.DataProcessamento, "0000/00/00"), "dd/mm/yyyy")

End Sub

'
'Carrega o vetor de Capas
'
Private Function CarregaVetorCapas() As Boolean

    On Error GoTo Erro_CarregaVetorCapas
    
    Dim rsCapa              As rdoResultset
    Dim X                   As Integer
    
    ''''''''''''''''''''''
    'Limpa vetor de capas'
    ''''''''''''''''''''''
    Erase aCapa
    Erase aDoc
    
    CarregaVetorCapas = False
    '''''''''''''''''''''''''''''''''''''''''
    'Pega capas disponíveis para confirmação'
    '''''''''''''''''''''''''''''''''''''''''
    With qryGetCapaConfAgConta
        .rdoParameters(0) = Geral.DataProcessamento
        .rdoParameters(1) = Geral.Intervalo
    End With
    
    Set rsCapa = qryGetCapaConfAgConta.OpenResultset(rdOpenStatic, rdConcurReadOnly)
    
    If Not rsCapa.EOF Then
        '''''''''''''''''''''''''''''''''
        'Desabilitar o Timer de Pesquisa'
        '''''''''''''''''''''''''''''''''
        TmrPesquisa.Enabled = False
        ShowFrmPesquisa False

        ReDim Preserve aCapa(rsCapa.RowCount - 1)

        X = 0
        While Not rsCapa.EOF
            '''''''''''''''''''''''''''''''''
            'Carregando o Array com as Capas'
            '''''''''''''''''''''''''''''''''
            aCapa(X).IdCapa = rsCapa!IdCapa
            aCapa(X).IdLote = rsCapa!IdLote
            aCapa(X).IdEnv_Mal = rsCapa!IdEnv_Mal
            aCapa(X).Capa = rsCapa!Capa
            aCapa(X).Num_Malote = rsCapa!Num_Malote
            aCapa(X).AgOrig = rsCapa!AgOrig
            aCapa(X).Status = rsCapa!Status
            aCapa(X).Duplicidade = rsCapa!Duplicidade
        
            rsCapa.MoveNext
            X = X + 1
        Wend
        CarregaVetorCapas = True
    End If
        
    Exit Function

Erro_CarregaVetorCapas:
    Screen.MousePointer = vbDefault
    Call TratamentoErro("Erro ao carregar as Capas.", Err, rdoErrors)
    Unload Me
End Function

'
'Retorna True se documento é para confirmação
'
Private Function DocumentoConfirmacao(ByRef pDoc As TpDocumento) As Boolean
    ''''''''''''''''''''''''''''''''''''''
    'Procura o documento à ser confirmado'
    ''''''''''''''''''''''''''''''''''''''
    DocumentoConfirmacao = False
    If pDoc.Status = "L" Then
        DocumentoConfirmacao = True
    End If

End Function

Private Function EnviarCapaPara(ByVal pPara As eEnviarCapaPara) As Boolean

    On Error GoTo Erro_CapaIlegivel
    
    EnviarCapaPara = False

    If Globais.AtualizaStatusCapa(aCapa(m_IndexCapa).IdCapa, Trim(pPara)) Then
        ''''''''''''''''''''''''''''''''''''''''''''
        'Determina que esta capa foi para Ilegíveis'
        ''''''''''''''''''''''''''''''''''''''''''''
        aCapa(m_IndexCapa).Status = Trim(pPara)
        
        EnviarCapaPara = True
        
    End If
    
    Exit Function
    
Erro_CapaIlegivel:

End Function

Sub HabilitaTimerPesquisa()

    ''''''''''''''''''''''''''''''''''
    'Desabilitar o Timer de Atualização'
    ''''''''''''''''''''''''''''''''''
    TmrAtualiza.Enabled = False

    'Esta Função irá verificar a existência de documentos Ilegíveis a cada x segundos
    'de acordo com o campo PARAMETRO.TmAtualizacao
    ShowFrmPesquisa True
    TmrPesquisa.Enabled = True
    
    HDObjetosAgenciaConta False
    
    Progress.Value = 0
    
End Sub

Sub HDObjetosAgenciaConta(ByVal pValor As Boolean)

    frmAgConta.Enabled = pValor
            
    txtAgencia.Enabled = pValor
    txtAgencia.BackColor = IIf(pValor = True, HABILITADO, DESABILITADO)
    
    txtConta.Enabled = pValor
    txtConta.BackColor = IIf(pValor = True, HABILITADO, DESABILITADO)
    
    lblAgencia.Enabled = pValor
    lblConta.Enabled = pValor
    
    '''''''''''''''''''''''''''''''''''''''''''
    'Limpa tambem os campos de agencia e conta'
    '''''''''''''''''''''''''''''''''''''''''''
    txtAgencia.Text = ""
    txtConta.Text = ""

End Sub

Private Sub HDObjetosNavegacao(ByVal pValue As Boolean)

    cmdDoctoAnterior.Enabled = pValue
    cmdDoctoPosterior.Enabled = pValue
    cmdSupervisor.Enabled = pValue
    
End Sub

Private Function MostraDocumento(ByVal pPosicao As ePosicao) As Boolean

    Dim bRetorno            As Boolean

    On Error GoTo Erro_MostraDocumento:
    
    
    MostraDocumento = False
    bRetorno = False
    
    If pPosicao = eAnterior Then
        If m_IndexDoc > LBound(aDoc) Then
            m_IndexDoc = m_IndexDoc - 1
            bRetorno = True
        End If
    ElseIf pPosicao = ePosterior Then
        If UBound(aDoc) > m_IndexDoc Then
            m_IndexDoc = m_IndexDoc + 1
            bRetorno = True
        End If
    End If
    
    If bRetorno Then
        
        MostraImagem
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Se o documento corrente é para confirmação, então habilita os campos Agencia e conta'
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        HDObjetosAgenciaConta DocumentoConfirmacao(aDoc(m_IndexDoc))
        
        txtAgencia.Text = IIf(Val(aDoc(m_IndexDoc).Agencia) = 0, "", Format(aDoc(m_IndexDoc).Agencia, "0000"))
        txtConta.Text = IIf(Val(aDoc(m_IndexDoc).Conta) = 0, "", aDoc(m_IndexDoc).Conta)
    End If
    
    MostraDocumento = bRetorno
    
    Exit Function
    
Erro_MostraDocumento:
    

End Function

'
'Rotina desenvolvida especialmente para complementação de documentos para confirmação
'
'Só mostra documentos com status "L"
'
Private Function MostraProximoDocumento() As Boolean

    Dim i           As Integer
    MostraProximoDocumento = False
    
    
    i = 0
    For i = 0 To UBound(aDoc)
        If DocumentoConfirmacao(aDoc(i)) Then
            m_IndexDoc = i
            MostraImagem
            MostraProximoDocumento = True
            
            Cabecalho True
            
            ''''''''''''''''''''''''''''''''''''
            'Limpa os campos de agencia e conta'
            ''''''''''''''''''''''''''''''''''''
            txtAgencia.Text = ""
            txtConta.Text = ""
            
            '''''''''''''''''''''''''''''''
            'Habilita campo para digitação'
            '''''''''''''''''''''''''''''''
            HDObjetosAgenciaConta True
            If Me.Visible Then txtAgencia.SetFocus
            
            Exit Function
        End If
    Next i
    
    Cabecalho False
    
End Function

Private Function ObtemDocumento(ByRef pDoc As TpDocumento) As Boolean

    Dim qryGetDeposito      As RDO.rdoQuery
    Dim qryGetADCC          As RDO.rdoQuery
    Dim qryGetOCT           As RDO.rdoQuery
    Dim rsDocumento         As RDO.rdoResultset

    On Error GoTo Erro_ObtemDocumento:
    
    ObtemDocumento = False
    
    ''''''''''''''''''''''''''''''
    'Deposito em conta e poupanca'
    ''''''''''''''''''''''''''''''
    If pDoc.TipoDocto = 2 Or pDoc.TipoDocto = 3 Then
        Set qryGetDeposito = Geral.Banco.CreateQuery("", "{Call GetDeposito(?,?)}")
    
        With qryGetDeposito
            .rdoParameters(0) = Geral.DataProcessamento
            .rdoParameters(1) = pDoc.IdDocto
            Set rsDocumento = qryGetDeposito.OpenResultset(rdOpenStatic, rdConcurReadOnly)
            If Not rsDocumento.EOF Then
                With rsDocumento
                    pDoc.Agencia = !Agencia
                    pDoc.Conta = !Conta
                End With
            End If
            rsDocumento.Close
        End With
        qryGetDeposito.Close
        ObtemDocumento = True
    '''''''''''''''''''''''''''''''''''''''''
    'Autorização de Débito em Conta Corrente'
    '''''''''''''''''''''''''''''''''''''''''
    ElseIf pDoc.TipoDocto = 4 Then
        Set qryGetADCC = Geral.Banco.CreateQuery("", "{Call GetAdcc(?,?)}")
    
        With qryGetADCC
            .rdoParameters(0) = Geral.DataProcessamento
            .rdoParameters(1) = pDoc.IdDocto
            Set rsDocumento = qryGetADCC.OpenResultset(rdOpenStatic, rdConcurReadOnly)
            If Not rsDocumento.EOF Then
                With rsDocumento
                    pDoc.Agencia = !Agencia
                    pDoc.Conta = !Conta
                End With
            End If
            rsDocumento.Close
        End With
        qryGetADCC.Close
        ObtemDocumento = True
    ''''''''''''''''''''''
    'Ordem de Crédito OCT'
    ''''''''''''''''''''''
    ElseIf pDoc.TipoDocto = 37 Then
        Set qryGetOCT = Geral.Banco.CreateQuery("", "{Call GetOct(?,?)}")
    
        With qryGetOCT
            .rdoParameters(0) = Geral.DataProcessamento
            .rdoParameters(1) = pDoc.IdDocto
            Set rsDocumento = qryGetOCT.OpenResultset(rdOpenStatic, rdConcurReadOnly)
            If Not rsDocumento.EOF Then
                With rsDocumento
                    pDoc.Agencia = !AgenciaCredito
                    pDoc.Conta = !ContaCredito
                End With
            End If
            rsDocumento.Close
        End With
        qryGetOCT.Close
        ObtemDocumento = True
    End If


    Exit Function
Erro_ObtemDocumento:

    Call TratamentoErro("Não foi possível obter o documento.", Err, rdoErrors)

End Function

Private Sub ShowFrmPesquisa(ByVal pShow As Boolean)


    FrmPesquisa.Visible = pShow
    
    If pShow Then FrmPesquisa.ZOrder
    

End Sub

Private Sub cmdConfirmar_Click()

    Dim sTamanho        As String
    Dim sFile           As String
    Dim qryGetAgenf     As RDO.rdoQuery
    Dim rstAgenf        As RDO.rdoResultset
    Dim strEncripta     As String
    Dim iFile           As Integer
    
    On Error GoTo Erro_Confirmacao:

    '''''''''''''''''''''''''
    'Validar Agencia e Conta'
    '''''''''''''''''''''''''

    If (Trim(txtAgencia.Text) = "" Or Val(txtAgencia.Text) = 0) Or (Trim(txtConta.Text) = "" Or Val(txtConta.Text) = 0) Then
        MsgBox "Agência e/ou Conta Inválidos!", vbInformation, App.Title
        If txtAgencia.Enabled Then txtAgencia.SetFocus
        Exit Sub
    End If

    
    sTamanho = Len(Format(txtAgencia.Text, "0000") & Format(txtConta.Text, "0000000"))
    If Not Modulo10(Format(txtAgencia.Text, "0000") & Format(txtConta.Text, "0000000"), sTamanho) Then
        MsgBox "Agência e/ou Conta Inválidos!", vbInformation, App.Title
        If txtAgencia.Enabled Then txtAgencia.SetFocus
        Exit Sub
    End If

    Set qryGetAgenf = Geral.Banco.CreateQuery("", "{call GetAgenf (?)}")

    qryGetAgenf.rdoParameters(0) = Val(txtAgencia.Text)
    
    Set rstAgenf = qryGetAgenf.OpenResultset(rdOpenStatic, rdConcurReadOnly)
    
    If rstAgenf.EOF() Then
        MsgBox "Código da agência inválida.", vbExclamation, App.Title
        If txtAgencia.Enabled Then txtAgencia.SetFocus
        Exit Sub
    End If

    rstAgenf.Close
    qryGetAgenf.Close
    ''''''''''''''''''''''''''''''''''''''''''
    'Acerta o documento com a agencia e conta'
    ''''''''''''''''''''''''''''''''''''''''''
    If Trim(CDbl(aDoc(m_IndexDoc).Agencia)) <> Trim(CDbl(txtAgencia.Text)) Or _
       Trim(CDbl(aDoc(m_IndexDoc).Conta)) <> Trim(CDbl(txtConta.Text)) Then
        Beep
        If (MsgBox("O Número da agência/conta não confere com o que foi digitado anteriormente." & _
            vbCrLf & "Confirma ALTERAÇÃO de agência/conta?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title)) = vbNo Then
            txtAgencia.SetFocus
            Exit Sub
        End If
        
        aDoc(m_IndexDoc).Agencia = txtAgencia.Text
        aDoc(m_IndexDoc).Conta = txtConta.Text
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Atualiza os dados da agencia e conta de qualquer um destes documentos'
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Atualiza campo Autenticação Digital
        If aDoc(m_IndexDoc).TipoDocto = 5 Or aDoc(m_IndexDoc).TipoDocto = 6 Or aDoc(m_IndexDoc).TipoDocto = 7 Then
            'Atualiza campo Autenticação Digital
            strEncripta = G_EncriptaBO(aDoc(m_IndexDoc).TipoDocto, aDoc(m_IndexDoc).Leitura)
            If strEncripta = "" Then
                MsgBox "Não foi possível atualizar agência/conta do documento.", vbCritical
                Exit Sub
            End If
        Else
            'Atualiza campo Autenticação Digital
            strEncripta = G_EncriptaBO(aDoc(m_IndexDoc).TipoDocto, CStr(Val(txtConta.Text)))
            If strEncripta = "" Then
                MsgBox "Não foi possível atualizar agência/conta do documento.", vbCritical
                Exit Sub
            End If
        End If

        With qryAtualizaDocumentoConfAgConta
            .rdoParameters(0).Direction = rdParamReturnValue
            .rdoParameters(1) = Geral.DataProcessamento         'Data de Processamento
            .rdoParameters(2) = aDoc(m_IndexDoc).TipoDocto      'TipoDocto
            .rdoParameters(3) = aDoc(m_IndexDoc).IdDocto        'IdDocto
            .rdoParameters(4) = aDoc(m_IndexDoc).Agencia        'Agencia
            .rdoParameters(5) = aDoc(m_IndexDoc).Conta          'Conta Corrente
            .rdoParameters(6) = strEncripta                     'Autenticacao digital
            .Execute
            If .rdoParameters(0) <> 0 Then
                MsgBox "Não foi possível atualizar agência/conta do documento.", vbCritical
                Exit Sub
            End If
        End With
        
            
    End If

    ''''''''''''''''''''''''''
    ' Grava Log do Documento '
    ''''''''''''''''''''''''''
    Call GravaLog(aCapa(m_IndexCapa).IdCapa, aDoc(m_IndexDoc).IdDocto, 211)

    sFile = IIf(Right(Geral.DiretorioImagens, 1) = "\", Geral.DiretorioImagens, Geral.DiretorioImagens & "\") & "ConfAgConta_" & Geral.Usuario & ".log"
        
    ''''''''''''''''''''''''''''''''''
    'Grava no arquivo texto o De Para'
    ''''''''''''''''''''''''''''''''''
    iFile = FreeFile
    Open sFile For Append As #iFile
        
    Print #iFile, "======== DE - PARA ========="
    Print #iFile, "DataProcessamento - " & Geral.DataProcessamento
    Print #iFile, "IdCapa            - " & aCapa(m_IndexCapa).IdCapa
    Print #iFile, "Capa              - " & aCapa(m_IndexCapa).Capa
    Print #iFile, "IdDocto           - " & aDoc(m_IndexDoc).IdDocto
    Print #iFile, "Login             - " & Geral.Usuario
    Print #iFile, "Hora              - " & Time
    Print #iFile, "----------------------------"
    Print #iFile, "DE   =>          Ag:" & txtAgencia.Text
    Print #iFile, "DE   =>       Conta:" & txtConta.Text
    Print #iFile, "PARA =>          Ag:" & aDoc(m_IndexDoc).Agencia
    Print #iFile, "PARA =>       Conta:" & aDoc(m_IndexDoc).Conta
    Print #iFile, "============================"
        
    Close #iFile

    aDoc(m_IndexDoc).Status = "1"

    If Not Globais.AtualizaStatusDocumento(aDoc(m_IndexDoc).IdDocto, "1") Then
        txtAgencia.SetFocus
        Exit Sub
    End If

    ''''''''''''''''''''''''''''''''''''''''''
    'Mostra o proximo documento da mesma capa'
    ''''''''''''''''''''''''''''''''''''''''''
           
    If Not MostraProximoDocumento Then
        '''''''''''''''''''''''''''
        'Envia capa para Ilegiveis'
        '''''''''''''''''''''''''''
        If m_EnviarParaIlegiveis Then
            m_EnviarParaIlegiveis = False
            If Not EnviarCapaPara(eIlegiveis) Then
                MsgBox "Não foi possível enviar a capa para Ilegíveis.", vbCritical
                Exit Sub
            End If
        Else
            ''''''''''''''''''''''''''''''''
            'enviar para vinculo automatico'
            ''''''''''''''''''''''''''''''''
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Caso não exista mais documentos desta capa, então enviar para Vinculo'
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                  
            Call EnviarCapaPara(eVinculoAutomatico)
        End If
        
        If Not CarregaVetorCapas Then
            FrmImagem.Visible = False
            HDObjetosImagem False
            HDObjetosNavegacao False
            HDObjetosAgenciaConta False
            MsgBox "Não Existem Envelopes / Malotes para confirmação de Agência e Conta.", vbInformation, App.Title
            Call HabilitaTimerPesquisa
            Exit Sub
        End If
        
        Call CarregarDocumentos
    End If
    
    Exit Sub
Erro_Confirmacao:

    Call TratamentoErro("Erro ao atualizar o documento.", Err, rdoErrors)
    
End Sub

Private Sub cmdDoctoAnterior_Click()

    If MostraDocumento(eAnterior) Then
        Cabecalho True
    Else
        Beep
    End If
    
    If DocumentoConfirmacao(aDoc(m_IndexDoc)) Then
        HDObjetosAgenciaConta True
        DoEvents
        txtAgencia.SetFocus
    End If
    
End Sub

Private Sub cmdDoctoPosterior_Click()

    If MostraDocumento(ePosterior) Then
        Cabecalho True
    Else
        Beep
    End If
    
    If DocumentoConfirmacao(aDoc(m_IndexDoc)) Then
        HDObjetosAgenciaConta True
        txtAgencia.SetFocus
    End If

End Sub

Private Sub CmdFecharPesquisa_Click()

    Call cmdFinalizar_Click
    
End Sub

Private Sub cmdFinalizar_Click()

    On Error GoTo Erro_Finalizar

    '''''''''''''''''''''''''''''''''
    'Volta o status anterior da capa'
    '''''''''''''''''''''''''''''''''
    If UBound(aCapa) >= 0 Then
    
        If m_EnviarParaIlegiveis Then
            m_EnviarParaIlegiveis = False
            If Not EnviarCapaPara(eIlegiveis) Then
                MsgBox "Não foi possível enviar a capa para Ilegíveis.", vbCritical
                Exit Sub
            End If
        Else
            '''''''''''''''''''''''''''''''''''''
            'Caso contrario, procedimento normal'
            '''''''''''''''''''''''''''''''''''''
            If aCapa(m_IndexCapa).Status = "M" Then
                ''''''''''''''''''''''''''''''''''''''''''''
                'Chama AtualizaStatusCapa do Modulo GLOBAIS'
                ''''''''''''''''''''''''''''''''''''''''''''
                If Not Globais.AtualizaStatusCapa(aCapa(m_IndexCapa).IdCapa, "L") Then
                    MsgBox "Não foi possível atualizar o status da capa.", vbCritical
                End If
    
            End If
        End If
    End If

Erro_Finalizar:
    Unload Me
End Sub

Private Sub cmdFrenteVerso_Click()

  On Error GoTo ERRO_FRENTEVERSO

  If teclou Then Exit Sub

  If FrmImagem.Visible = False Then Exit Sub

  teclou = True
  
  'só muda para frente/verso qdo docto vem da Ls500 e da Vips, pois a canon não gera verso.
  If (aDoc(m_IndexDoc).Ordem = "0") Or (aDoc(m_IndexDoc).Ordem = "2") Then
    If Lead1.Tag = "V" Then
        Lead1.Tag = "F"     'se verso, mostrar frente
        With Lead1
            .AutoRepaint = False
            If Geral.VIPSDLL = eDllProservi Then
              .Load Geral.DiretorioImagens & aDoc(m_IndexDoc).Frente, 0, 0, 1
            Else
              .Load Geral.DiretorioImagens & Format(aCapa(m_IndexCapa).IdLote, "000000000") & "\" & aDoc(m_IndexDoc).Frente, 0, 0, 1
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
              .Load Geral.DiretorioImagens & Trim(aDoc(m_IndexDoc).Verso), 0, 0, 1
            Else
              .Load Geral.DiretorioImagens & Format(aCapa(m_IndexCapa).IdLote, "000000000") & "\" & aDoc(m_IndexDoc).Verso, 0, 0, 1
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

Private Sub cmdSupervisor_Click()

    m_EnviarParaIlegiveis = True
    
    ''''''''''''''''''''''''''
    ' Grava Log do Documento '
    ''''''''''''''''''''''''''
    Call GravaLog(aCapa(m_IndexCapa).IdCapa, aDoc(m_IndexDoc).IdDocto, 210)

    If Globais.AtualizaStatusDocumento(aDoc(m_IndexDoc).IdDocto, "1") Then
        aDoc(m_IndexDoc).Status = "1"
    End If
    
    If Not MostraProximoDocumento() Then
        ''''''''''''''''''''
        'Busca próxima capa'
        ''''''''''''''''''''
        If Not CarregarDocumentos() Then
        
            '''''''''''''''''''''''''''''''''''''''''''''
            'Não existe mais documentos desta capa e não'
            'tem mais capa para confirmação. Enviar esta'
            'capa para Ilegiveis                        '
            '''''''''''''''''''''''''''''''''''''''''''''
            If Not EnviarCapaPara(eIlegiveis) Then
                MsgBox "Não foi possível enviar a capa para Ilegíveis.", vbCritical
            End If
        
        
            FrmImagem.Visible = False
            HDObjetosImagem False
            HDObjetosNavegacao False
            HDObjetosAgenciaConta False
            Cabecalho False
            Erase aCapa
            MsgBox "Não Existem Envelopes / Malotes para confirmação de Agência e Conta.", vbInformation, App.Title
            Call HabilitaTimerPesquisa
        End If
    End If
    
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

Private Sub Form_Activate()

    On Error GoTo ERRO_ACTIVATE

    'Inclusão de chamada a rotina AtualizaAtividade
    Call AtualizaAtividade(19)
    
    With Lead1
        .UnlockSupport L_SUPPORT_EXPRESS, "YXPQ3XPPVT"
        .UnlockSupport L_SUPPORT_GIFLZW, "0K3RV9UY3EY"
        .UnlockSupport L_SUPPORT_TIFLZW, "9LE75L0FDXHK"
    End With

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Preenche vetor com as capas de documentos à serem confirmados'
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If PrimeiraVez Then
        PrimeiraVez = False

        AlterouDocto = False
        If Not CarregaVetorCapas Then
            FrmImagem.Visible = False
            HDObjetosImagem False
            HDObjetosNavegacao False
            HDObjetosAgenciaConta False
            MsgBox "Não Existem Envelopes / Malotes para confirmação de Agência e Conta.", vbInformation, App.Title
            Call HabilitaTimerPesquisa
            Exit Sub
        End If
        
        sTempo = 0
        
        CarregarDocumentos
        
    End If
    
    Exit Sub

ERRO_ACTIVATE:
    Screen.MousePointer = vbDefault
    
    Call TratamentoErro("Erro ao Ativar Tela.", Err, rdoErrors)
    
    Unload Me
End Sub
'
'Rotina que confirma os documentos
'
Private Function CarregarDocumentos() As Boolean

    Dim i                   As Integer
    Dim rsDocumentos        As RDO.rdoResultset

    On Error GoTo Erro_CarregarDocumentos
    
    CarregarDocumentos = True
    
    ''''''''''''''''
    'Loop das capas'
    ''''''''''''''''
    For i = 0 To UBound(aCapa)
        m_IndexCapa = i
        
        If Not Globais.AtualizaStatusCapa(aCapa(m_IndexCapa).IdCapa, "M") Then
            GoTo Proxima_Capa
        End If
        '''''''''''''''''''''''''''''''''''''''''''''
        'Determina que esta capa está em Confirmação'
        '''''''''''''''''''''''''''''''''''''''''''''
        aCapa(m_IndexCapa).Status = "M"
        
        ''''''''''''''''''''''''''''''''''
        'Habilitar o Timer de Atualização'
        ''''''''''''''''''''''''''''''''''
        TmrAtualiza.Enabled = True
        
        With qryGetDocumentoConfAgConta
            .rdoParameters(0) = Geral.DataProcessamento
            .rdoParameters(1) = aCapa(m_IndexCapa).IdCapa
            Set rsDocumentos = qryGetDocumentoConfAgConta.OpenResultset(rdOpenStatic, rdConcurReadOnly)
        End With
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Tenta atualizar esta capa "Em Confirmação", caso não consiga, tenta outra capa '          '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If Not rsDocumentos.EOF Then
            ''''''''''''''''''''''''''''''''''''
            'Redimensiona o vetor de documentos'
            ''''''''''''''''''''''''''''''''''''
            Erase aDoc
            ReDim aDoc(rsDocumentos.RowCount - 1)

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Carrega vetor de documentos marcados para confirmação de Agencia e Conta'
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            m_IndexDoc = 0
            Do While Not rsDocumentos.EOF
                '''''''''''''''''''''''''''''''''
                'Carregando dados dos documentos'
                '''''''''''''''''''''''''''''''''
                With rsDocumentos
                    aDoc(m_IndexDoc).IdDocto = !IdDocto
                    aDoc(m_IndexDoc).Leitura = !Leitura & ""
                    aDoc(m_IndexDoc).Frente = !Frente
                    aDoc(m_IndexDoc).Verso = !Verso
                    aDoc(m_IndexDoc).TipoDocto = !TipoDocto
                    aDoc(m_IndexDoc).Status = !Status
                    aDoc(m_IndexDoc).Ordem = !Ordem
                    '''''''''''''''''''''''''''''
                    'Obtem o documento pelo tipo'
                    '''''''''''''''''''''''''''''
                    ObtemDocumento aDoc(m_IndexDoc)
                End With
                m_IndexDoc = m_IndexDoc + 1
                rsDocumentos.MoveNext
            Loop

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Carrega imagem do próximo documento à ser confirmado'
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If Not MostraProximoDocumento Then
                '''''''''''''''''''''''''''''''''
                'Volta o status anterior da capa'
                '''''''''''''''''''''''''''''''''
                If aCapa(m_IndexCapa).Status = "M" Then
                    EnviarCapaPara eVinculoAutomatico
                    GoTo Proxima_Capa
                End If
            End If
            
            ''''''''''''''''''''''
            'Preenche o cabeçalho'
            ''''''''''''''''''''''
            Cabecalho True

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Sai do Loop e espera o usuario confirmar a agencia e conta'
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Exit For
        End If
Proxima_Capa:
    Next i
    
    If i > UBound(aCapa) Then CarregarDocumentos = False
    
    Exit Function
    
Erro_CarregarDocumentos:
'    Resume
    Call TratamentoErro("Erro ao obter o documento para confirmação.", Err, rdoErrors)
    
End Function

Private Sub MostraImagem()

    On Error GoTo ERRO_MOSTRAIMAGEM
    
    Dim Ret As Long
    
    hCtl = Lead1.hwnd
    
    'Coloca imagem na tela
    With Lead1
        .Tag = "F"
        .AutoRepaint = False
        If Geral.VIPSDLL = eDllProservi Then
            .Load Geral.DiretorioImagens & aDoc(m_IndexDoc).Frente, 0, 0, 1
        Else
            .Load Geral.DiretorioImagens & Format(aCapa(m_IndexCapa).IdLote, "000000000") & "\" & aDoc(m_IndexDoc).Frente, 0, 0, 1
        End If
        
        'Se imagem for da ls500, deixar mais escura
        If aDoc(m_IndexDoc).Ordem <> "2" Then
            .Intensity 220
        Else
            .Intensity 140
        End If
        'Se imagem for do canon, diminui em 50% o tamanho
        If aDoc(m_IndexDoc).Ordem <> "1" Then
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

Sub HDObjetosImagem(bValor As Boolean)

    On Error GoTo ERRO_HDOBJETOS
    
    cmdZoomMais.Enabled = bValor
    cmdZoomMenos.Enabled = bValor
    cmdRotacao.Enabled = bValor
    cmdInverteCor.Enabled = bValor
    cmdFrenteVerso.Enabled = bValor
    cmdConfirmar.Enabled = bValor
    FrmImagem.Visible = bValor
    Lead1.ForceRepaint
    
    Exit Sub

ERRO_HDOBJETOS:
    Screen.MousePointer = vbDefault
    Call TratamentoErro("Erro ao preparar botões de manipulação de Imagens.", Err, rdoErrors)
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
'        Case vbKeyEscape
'            Call cmdFinalizar_Click
'            Exit Sub
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

Private Sub tmrAtualiza_Timer()

    TmrAtualiza.Enabled = False
    
    If aCapa(m_IndexCapa).IdCapa <> 0 Then
        sTempo = sTempo + Int(TmrAtualiza.Interval / 1000)
        If sTempo + Int(TmrAtualiza.Interval / 1000) >= Geral.Intervalo Then
            'Atualizar o Status da Capa
            Call Globais.AtualizaStatusCapa(aCapa(m_IndexCapa).IdCapa, "M")
            sTempo = 0
        End If
    End If
    
    TmrAtualiza.Enabled = True

End Sub

Private Sub tmrPesquisa_Timer()

    TmrPesquisa.Enabled = False
    
    sTempo = sTempo + Int(TmrPesquisa.Interval / 1000)
    
    If sTempo + Int(TmrPesquisa.Interval / 1000) >= Geral.Atualizacao Then
        'Pesquisar por Documentos Ilegíveis
        sTempo = 0
        
        If CarregaVetorCapas Then
            CarregarDocumentos
            Exit Sub
        End If
        
        TmrPesquisa.Enabled = True
    End If
    
    'Atualizar a Barra de Progresso
    If Progress.Value + 4 > 100 Then
        Progress.Value = 0
    Else
        Progress.Value = Progress.Value + 4
    End If
    
    DoEvents
    TmrPesquisa.Enabled = True

End Sub

Private Sub txtAgencia_Change()

    If Len(txtAgencia.Text) = txtAgencia.MaxLength And txtConta.Enabled Then txtConta.SetFocus

End Sub

Private Sub txtAgencia_GotFocus()

    SelecionarTexto txtAgencia

End Sub

Private Sub txtAgencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtConta.SetFocus
        KeyAscii = 0
    Else
        SoNumero KeyAscii
    End If
End Sub

Private Sub txtConta_Change()

    If Len(txtConta.Text) = txtConta.MaxLength Then cmdConfirmar.SetFocus

End Sub

Private Sub txtConta_GotFocus()

    DoEvents
    SelecionarTexto txtConta

End Sub

Private Sub txtConta_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        cmdConfirmar_Click
        KeyAscii = 0
    Else
        SoNumero KeyAscii
    End If
End Sub
