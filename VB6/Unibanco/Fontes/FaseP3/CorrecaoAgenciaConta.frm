VERSION 5.00
Object = "{00080003-B1BA-11CE-ABC6-F5B2E79D9E3F}#8.0#0"; "LTOCX80N.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form CorrecaoAgenciaConta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Correção de Agência e Conta"
   ClientHeight    =   7428
   ClientLeft      =   36
   ClientTop       =   252
   ClientWidth     =   10284
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7428
   ScaleWidth      =   10284
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrPesquisa 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7920
      Top             =   5280
   End
   Begin VB.Timer tmrAtualiza 
      Enabled         =   0   'False
      Interval        =   50000
      Left            =   8400
      Top             =   5280
   End
   Begin VB.PictureBox FrmPesquisa 
      Height          =   1932
      Left            =   2232
      ScaleHeight     =   1884
      ScaleWidth      =   5724
      TabIndex        =   19
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
         TabIndex        =   21
         Top             =   912
         Width           =   5028
         _ExtentX        =   8869
         _ExtentY        =   529
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Pesquisando por Documentos para correção. Aguarde ..."
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
         Left            =   345
         TabIndex        =   22
         Top             =   570
         Width           =   5100
      End
   End
   Begin VB.Frame FrmImagem 
      Caption         =   "Imagem"
      Height          =   4740
      Left            =   78
      TabIndex        =   17
      Top             =   384
      Width           =   10080
      Begin LeadLib.Lead Lead1 
         Height          =   4455
         Left            =   180
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   240
         Width           =   9870
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
   Begin VB.Frame Frame1 
      Height          =   1068
      Left            =   120
      TabIndex        =   10
      Top             =   6000
      Width           =   6504
      Begin VB.CommandButton CmdDevolDocto 
         Caption         =   "&Devolver"
         Enabled         =   0   'False
         Height          =   800
         Left            =   4600
         Picture         =   "CorrecaoAgenciaConta.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   180
         Width           =   900
      End
      Begin VB.CommandButton cmdConfirmar 
         Caption         =   "Confirmar"
         Height          =   800
         Left            =   5500
         Picture         =   "CorrecaoAgenciaConta.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   180
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
         Height          =   800
         Left            =   1000
         Picture         =   "CorrecaoAgenciaConta.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   180
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
         Height          =   800
         Left            =   100
         Picture         =   "CorrecaoAgenciaConta.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   180
         Width           =   900
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
         Left            =   3700
         Picture         =   "CorrecaoAgenciaConta.frx":0D60
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   180
         Width           =   900
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
         Left            =   2800
         Picture         =   "CorrecaoAgenciaConta.frx":106A
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   180
         Width           =   900
      End
      Begin VB.CommandButton cmdRotacao 
         Caption         =   "Rotação"
         Height          =   800
         Left            =   1900
         Picture         =   "CorrecaoAgenciaConta.frx":1374
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   180
         Width           =   900
      End
   End
   Begin VB.Frame fraBotoesSuperiores 
      Height          =   1068
      Left            =   7290
      TabIndex        =   6
      Top             =   6000
      Width           =   2892
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
         Left            =   1900
         Picture         =   "CorrecaoAgenciaConta.frx":167E
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   192
         Width           =   900
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
         Left            =   100
         Picture         =   "CorrecaoAgenciaConta.frx":1988
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   192
         Width           =   900
      End
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
         Picture         =   "CorrecaoAgenciaConta.frx":1DCA
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   192
         Width           =   900
      End
   End
   Begin VB.Frame frameAgConta 
      Height          =   876
      Left            =   3510
      TabIndex        =   1
      Top             =   5136
      Width           =   4200
      Begin VB.ComboBox CboTipoConta 
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
         Height          =   360
         ItemData        =   "CorrecaoAgenciaConta.frx":220C
         Left            =   2220
         List            =   "CorrecaoAgenciaConta.frx":220E
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   405
         Width           =   1848
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
         Left            =   990
         MaxLength       =   7
         TabIndex        =   3
         Top             =   405
         Width           =   960
      End
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
         TabIndex        =   2
         Top             =   390
         Width           =   612
      End
      Begin VB.Label LblTipoConta 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Conta"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2220
         TabIndex        =   25
         Top             =   180
         Width           =   780
      End
      Begin VB.Label lblAgencia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Agência"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   195
         Width           =   600
      End
      Begin VB.Label lblConta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Conta"
         Height          =   195
         Left            =   1035
         TabIndex        =   4
         Top             =   195
         Width           =   420
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   1  'Align Top
      Height          =   348
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10284
      _ExtentX        =   18140
      _ExtentY        =   614
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   5305
            MinWidth        =   1834
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   8129
            MinWidth        =   4658
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   4601
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
End
Attribute VB_Name = "CorrecaoAgenciaConta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private qryGetCapaCorrAgConta                   As RDO.rdoQuery
Private qryGetDocumentoCorrAgConta              As RDO.rdoQuery
Private qryAtualizaDocumentoCorrAgConta         As RDO.rdoQuery
Private qryAtualizaOcorrDoctoCorrAgConta        As RDO.rdoQuery
        
Private aCapa()                                 As tpCapa
Private aDoc()                                  As TpDocumento
Private m_IndexCapa                             As Integer
Private m_IndexDoc                              As Integer
Private sTempo                                  As Integer
Private teclou                                  As Boolean
Private PrimeiraVez                             As Boolean
Private AlterouDocto                            As Boolean

Private Enum ePosicao
    eAnterior
    ePosterior
End Enum

Private Const HABILITADO = &H80000005           'Window Background
Private Const DESABILITADO = &H8000000F         'Button Face
Private Const STATUS_TXT_DOCUMENTO = "Documento :"
Private Const STATUS_TXT_LOTECAPA = "Lote/Capa :"
Private Const STATUS_TXT_DATA = "Data :"

Private Sub CboTipoConta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdConfirmar_Click
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

Private Sub Form_Load()

    Set qryGetCapaCorrAgConta = Geral.Banco.CreateQuery("", "{call GetCapaCorrAgConta(?,?)}")
    Set qryGetDocumentoCorrAgConta = Geral.Banco.CreateQuery("", "{Call GetDocumentoCorrAgConta(?,?)}")
    Set qryAtualizaDocumentoCorrAgConta = Geral.Banco.CreateQuery("", "{? = call AtualizaDocumentoCorrAgConta (?,?,?,?,?,?,?)}")
    Set qryAtualizaOcorrDoctoCorrAgConta = Geral.Banco.CreateQuery("", "{? = call AtualizaOcorrDoctoCorrAgConta (?,?,?,?)}")
                                                                               
    Cabecalho False
    PrimeiraVez = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
    qryGetCapaCorrAgConta.Close
    qryGetDocumentoCorrAgConta.Close
    qryAtualizaDocumentoCorrAgConta.Close
    qryAtualizaOcorrDoctoCorrAgConta.Close

End Sub

Private Sub CmdDevolDocto_Click()
    On Error GoTo ERRO_OCORRENCIA
    
    Beep
    If (MsgBox("Confirma a Exclusão/Devolução do Documento ", vbQuestion + vbYesNo, App.Title)) = vbNo Then
        Exit Sub
    End If
       
   'Atualizar o Campo 'OCORRENCIA'
    With qryAtualizaOcorrDoctoCorrAgConta
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1) = Geral.DataProcessamento          'Data Proc.
        .rdoParameters(2) = aDoc(m_IndexDoc).TipoDocto       'TIpo Docto
        .rdoParameters(3) = aDoc(m_IndexDoc).Vinculo         'Vinculo dos Doctos
        .rdoParameters(4) = aCapa(m_IndexCapa).IdCapa        'IdCapa
        .Execute
    End With

    If qryAtualizaOcorrDoctoCorrAgConta(0).Value = 1 Then
        MsgBox "Ocorreu um erro ao atualizar a ocorrência do documento.", vbInformation + vbOKOnly, App.Title
        Exit Sub
    End If
    
    aDoc(m_IndexDoc).Status = "D"
        
    'Gravar Log
    Call GravaLog(aCapa(m_IndexCapa).IdCapa, aDoc(m_IndexDoc).IdDocto, 220)
          
    ''''''''''''''''''''''''''''''''''''''''''
    'Mostra o proximo documento da mesma capa'
    ''''''''''''''''''''''''''''''''''''''''''
    If Not MostraProximoDocumento Then
         '''''''''''''''''''''''''''
         ' Enviar para Transmissão '
         '''''''''''''''''''''''''''
        Call EnviarCapaPara
                
        If Not CarregaVetorCapas Then
            FrmImagem.Visible = False
            HDObjetosImagem False
            HDObjetosNavegacao False
            HDObjetosAgenciaConta False
            MsgBox "Não Existem Envelopes / Malotes para correção de Agência e Conta.", vbInformation, App.Title
            Call HabilitaTimerPesquisa
            Exit Sub
        End If
        
        Call CarregarDocumentos
    End If
    
    Exit Sub

ERRO_OCORRENCIA:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro ao preparar Documento para Ocorrência.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select

End Sub

Private Sub cmdDoctoAnterior_Click()

    If MostraDocumento(eAnterior) Then
       Cabecalho True
    Else
        Beep
    End If
    
    If DocumentoCorrecao(aDoc(m_IndexDoc)) Then
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
    
    If DocumentoCorrecao(aDoc(m_IndexDoc)) Then
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
        If aCapa(m_IndexCapa).Status = "Z" Then
                ''''''''''''''''''''''''''''''''''''''''''''
                'Chama AtualizaStatusCapa do Modulo GLOBAIS'
                ''''''''''''''''''''''''''''''''''''''''''''
            If Not Globais.AtualizaStatusCapa(aCapa(m_IndexCapa).IdCapa, "Y") Then
                MsgBox "Não foi possível atualizar o status da capa.", vbCritical
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
  'só muda para frente/verso qdo docto vem da Ls500 e da Vips,'pois, da canon não gera verso.
  
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
    
    ''''''''''''''''''''''''''''''''''''''
    'Pega capas disponíveis para correção'
    ''''''''''''''''''''''''''''''''''''''
    With qryGetCapaCorrAgConta
        .rdoParameters(0) = Geral.DataProcessamento
        .rdoParameters(1) = Geral.Intervalo
    End With
    
    Set rsCapa = qryGetCapaCorrAgConta.OpenResultset(rdOpenStatic, rdConcurReadOnly)
    
    If Not rsCapa.EOF Then
        '''''''''''''''''''''''''''''''''
        'Desabilitar o Timer de Pesquisa'
        '''''''''''''''''''''''''''''''''
        tmrPesquisa.Enabled = False
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
    CmdDevolDocto.Enabled = bValor
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
Private Sub Form_Activate()

    On Error GoTo ERRO_ACTIVATE
    
   'Preencher o Combo de Tipos de Conta
    CboTipoConta.AddItem "0 - Corrente"
    CboTipoConta.ItemData(CboTipoConta.NewIndex) = 2
    
    CboTipoConta.AddItem "9 - Poupança"
    CboTipoConta.ItemData(CboTipoConta.NewIndex) = 3

    'Inclusão de chamada a rotina  de controle de AtualizaAtividade
    Call AtualizaAtividade(25)
    
    With Lead1
        .UnlockSupport L_SUPPORT_EXPRESS, "YXPQ3XPPVT"
        .UnlockSupport L_SUPPORT_GIFLZW, "0K3RV9UY3EY"
        .UnlockSupport L_SUPPORT_TIFLZW, "9LE75L0FDXHK"
    End With

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Preenche vetor com as capas com doctos para corrição '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If PrimeiraVez Then
        PrimeiraVez = False
        AlterouDocto = False
        
        If Not CarregaVetorCapas Then
            FrmImagem.Visible = False
            HDObjetosImagem False
            HDObjetosNavegacao False
            HDObjetosAgenciaConta False
            MsgBox "Não Existem Envelopes / Malotes para correção de Agência e Conta.", vbInformation, App.Title
            Call HabilitaTimerPesquisa
            Exit Sub
        End If
        
        sTempo = 0
        
       'Carrega documentos das capas
        Call CarregarDocumentos
        
    End If
    
    Exit Sub

ERRO_ACTIVATE:
    Screen.MousePointer = vbDefault
    Call TratamentoErro("Erro ao Ativar Tela.", Err, rdoErrors)
    Unload Me
End Sub
Sub HabilitaTimerPesquisa()
    ''''''''''''''''''''''''''''''''''
    'Desabilitar o Timer de Atualização'
    ''''''''''''''''''''''''''''''''''
    tmrAtualiza.Enabled = False
    ShowFrmPesquisa True
    tmrPesquisa.Enabled = True
    HDObjetosAgenciaConta False
    Progress.Value = 0
    
End Sub
Private Sub tmrAtualiza_Timer()
    tmrAtualiza.Enabled = False
    
    If aCapa(m_IndexCapa).IdCapa <> 0 Then
        sTempo = sTempo + Int(tmrAtualiza.Interval / 1000)
        If sTempo + Int(tmrAtualiza.Interval / 1000) >= Geral.Intervalo Then
        
            'Atualiza Tempo da capa corrente
            If Not Globais.AtualizaStatusCapa(aCapa(m_IndexCapa).IdCapa, "Z") Then
                MsgBox "Não foi possível atualizar o status da capa.", vbCritical
            End If
        
            sTempo = 0
        End If
    End If
    
    tmrAtualiza.Enabled = True

End Sub
Private Sub tmrPesquisa_Timer()
    tmrPesquisa.Enabled = False
    
    sTempo = sTempo + Int(tmrPesquisa.Interval / 1000)
    
    If sTempo + Int(tmrPesquisa.Interval / 1000) >= Geral.Atualizacao Then
        'Pesquisar por Documentos para Correção
        sTempo = 0
        
        If CarregaVetorCapas Then
            CarregarDocumentos
            Exit Sub
        End If
        
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
Private Sub cmdConfirmar_Click()

    Dim sTamanho        As String
    Dim iFile           As Integer
    Dim sFile           As String
    Dim sAgencia_Old    As String
    Dim sConta_Old      As String
    Dim strEncripta     As String
    
    On Error GoTo Erro_Confirmacao:

    '''''''''''''''''''''''''
    'Validar Agencia e Conta'
    '''''''''''''''''''''''''
    If Trim(txtAgencia.Text) = "" Or Trim(txtConta.Text) = "" Then
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
    
    ''''''''''''''''''''''''''''''''''''''''''
    'Acerta o documento com a agencia e conta'
    ''''''''''''''''''''''''''''''''''''''''''
     Beep
    
    If (MsgBox("Confirma a Correção da agência/conta ?", vbQuestion + vbYesNo, App.Title)) = vbNo Then
        txtAgencia.SetFocus
        Exit Sub
    End If
        
    sAgencia_Old = aDoc(m_IndexDoc).Agencia
    sConta_Old = aDoc(m_IndexDoc).Conta

    aDoc(m_IndexDoc).Agencia = txtAgencia.Text
    aDoc(m_IndexDoc).Conta = txtConta.Text
    
    If aDoc(m_IndexDoc).TipoDocto <> 4 And aDoc(m_IndexDoc).TipoDocto <> 37 Then
        aDoc(m_IndexDoc).TipoDocto = IIf(CboTipoConta.ListIndex = 0, 2, 3)
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
    
    With qryAtualizaDocumentoCorrAgConta
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1) = Geral.DataProcessamento                         'Data de Processamento
        .rdoParameters(2) = aDoc(m_IndexDoc).TipoDocto                      'TipoDocto
        .rdoParameters(3) = aDoc(m_IndexDoc).IdDocto                        'IdDocto
        .rdoParameters(4) = aDoc(m_IndexDoc).Agencia                        'Agencia
        .rdoParameters(5) = aDoc(m_IndexDoc).Conta                          'Conta Corrente
        If aDoc(m_IndexDoc).TipoDocto = 37 Then
            .rdoParameters(6) = 37                                          'TipoConta
        Else
            .rdoParameters(6) = IIf(CboTipoConta.ListIndex = 0, 1, 2)       'TipoConta
        End If
        .rdoParameters(7) = strEncripta                                      'Autenticacao digital
        .Execute
        
        If .rdoParameters(0) <> 0 Then
            aDoc(m_IndexDoc).Agencia = sAgencia_Old
            aDoc(m_IndexDoc).Conta = sConta_Old
            MsgBox "Não foi possível atualizar agência/conta do documento.", vbCritical
            Exit Sub
        End If
        
        
        ''''''''''''''''''''''''''
        ' Grava Log do Documento '
        ''''''''''''''''''''''''''
        Call GravaLog(aCapa(m_IndexCapa).IdCapa, aDoc(m_IndexDoc).IdDocto, 221)
        
        sFile = IIf(Right(Geral.DiretorioImagens, 1) = "\", Geral.DiretorioImagens, Geral.DiretorioImagens & "\") & "CorrAgConta_" & Geral.Usuario & ".log"
        
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
            Print #iFile, "DE   =>          Ag:" & sAgencia_Old
            Print #iFile, "DE   =>       Conta:" & sConta_Old
            Print #iFile, "PARA =>          Ag:" & aDoc(m_IndexDoc).Agencia
            Print #iFile, "PARA =>       Conta:" & aDoc(m_IndexDoc).Conta
            Print #iFile, "============================"
        
        Close #iFile

    End With
    
    aDoc(m_IndexDoc).Status = "1"
    
    If Not AtualizaStatusDocumento(aDoc(m_IndexDoc).IdDocto, "1") Then
        txtAgencia.SetFocus
        Exit Sub
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''
    'Mostra o proximo documento da mesma capa'
    ''''''''''''''''''''''''''''''''''''''''''
    If Not MostraProximoDocumento Then
    
        'Enviar para Transmissão
        Call EnviarCapaPara
                
        If Not CarregaVetorCapas Then
            FrmImagem.Visible = False
            HDObjetosImagem False
            HDObjetosNavegacao False
            HDObjetosAgenciaConta False
            MsgBox "Não Existem Envelopes / Malotes para correção de Agência e Conta.", vbInformation, App.Title
            Call HabilitaTimerPesquisa
            Exit Sub
        End If
        
        Call CarregarDocumentos
    End If
    
    Exit Sub
    
Erro_Confirmacao:
    Call TratamentoErro("Erro ao atualizar o documento.", Err, rdoErrors)
    
End Sub
Private Function ObtemDocumento(ByRef pDoc As TpDocumento) As Boolean

    Dim qryGetDeposito      As RDO.rdoQuery
    Dim qryGetADCC          As RDO.rdoQuery
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
    End If


    Exit Function
    
Erro_ObtemDocumento:
    Call TratamentoErro("Não foi possível obter o documento.", Err, rdoErrors)

End Function
Sub HDObjetosAgenciaConta(ByVal pValor As Boolean)

    'Habilita e desabilita objetos durante navegação no documentos
        
    txtAgencia.Enabled = pValor
    txtAgencia.BackColor = IIf(pValor = True, HABILITADO, DESABILITADO)
    
    txtConta.Enabled = pValor
    txtConta.BackColor = IIf(pValor = True, HABILITADO, DESABILITADO)
    
    CboTipoConta.Enabled = pValor
    CboTipoConta.BackColor = IIf(pValor = True, HABILITADO, DESABILITADO)
    
    frameAgConta.Enabled = pValor
    If frameAgConta.Enabled Then
        If aDoc(m_IndexDoc).TipoDocto = 2 Or aDoc(m_IndexDoc).TipoDocto = 3 Then
            frameAgConta.Width = 4200
        Else
            frameAgConta.Width = 2200
        End If
    Else
        frameAgConta.Width = 2200
    End If
    
    lblAgencia.Enabled = pValor
    lblConta.Enabled = pValor
    
    CmdDevolDocto.Enabled = pValor
    cmdConfirmar.Enabled = pValor
            
    ''''''''''''''''''''''''''''''''''''
    'Limpa os campos de agencia e conta'
    ''''''''''''''''''''''''''''''''''''
    If Not pValor Then
        txtAgencia.Text = ""
        txtConta.Text = ""
    End If
End Sub
Private Sub HDObjetosNavegacao(ByVal pValue As Boolean)

    cmdDoctoAnterior.Enabled = pValue
    cmdDoctoPosterior.Enabled = pValue
        
End Sub
Private Function MostraProximoDocumento() As Boolean
 
    Dim i As Integer
    MostraProximoDocumento = False
          
    For i = 0 To UBound(aDoc)
        
        If DocumentoCorrecao(aDoc(i)) Then
            m_IndexDoc = i
            MostraImagem
            MostraProximoDocumento = True
            
            Cabecalho True
            
            ''''''''''''''''''''''''''''''''''''
            'Limpa os campos de agencia e conta'
            ''''''''''''''''''''''''''''''''''''
            txtAgencia.Text = IIf(Val(aDoc(m_IndexDoc).Agencia) = 0, "", Format(aDoc(m_IndexDoc).Agencia, "0000"))
            txtConta.Text = IIf(Val(aDoc(m_IndexDoc).Conta) = 0, "", aDoc(m_IndexDoc).Conta)
            
            If aDoc(m_IndexDoc).TipoDocto = 2 Then
                CboTipoConta.Enabled = True
                CboTipoConta.ListIndex = 0
            ElseIf aDoc(m_IndexDoc).TipoDocto = 3 Then
                CboTipoConta.Enabled = True
                CboTipoConta.ListIndex = 1
            Else
                CboTipoConta.Enabled = False
            End If
       
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
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Se o documento corrente é para correção, então habilita os campos Agencia e conta'
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        HDObjetosAgenciaConta DocumentoCorrecao(aDoc(m_IndexDoc))
        
        txtAgencia.Text = IIf(Val(aDoc(m_IndexDoc).Agencia) = 0, "", Format(aDoc(m_IndexDoc).Agencia, "0000"))
        txtConta.Text = IIf(Val(aDoc(m_IndexDoc).Conta) = 0, "", aDoc(m_IndexDoc).Conta)
        
        If aDoc(m_IndexDoc).TipoDocto = 2 Then
            frameAgConta.Width = 4200
            CboTipoConta.Enabled = True
            CboTipoConta.ListIndex = 0
        ElseIf aDoc(m_IndexDoc).TipoDocto = 3 Then
            frameAgConta.Width = 4200
            CboTipoConta.Enabled = True
            CboTipoConta.ListIndex = 0
        Else
            CboTipoConta.Enabled = False
            frameAgConta.Width = 2200
        End If
    End If
    
    MostraDocumento = bRetorno
    
    Exit Function
    
Erro_MostraDocumento:
    Call TratamentoErro("Não foi possível Mostrar o documento.", Err, rdoErrors)
End Function

Private Sub ShowFrmPesquisa(ByVal pShow As Boolean)

    FrmPesquisa.Visible = pShow
    
    If pShow Then FrmPesquisa.ZOrder
    
End Sub

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
        
        ''''''''''''''''''''''''''
        'Atualiza Status da Capa '
        ''''''''''''''''''''''''''
        If Globais.AtualizaStatusCapa(aCapa(m_IndexCapa).IdCapa, "Z") Then
                    
            '''''''''''''''''''''''''''''''''''''''''''
            'Determina que esta capa está em Correção '
            '''''''''''''''''''''''''''''''''''''''''''
            aCapa(m_IndexCapa).Status = "Z"
            
            ''''''''''''''''''''''''''''''''''
            'Habilitar o Timer de Atualização'
            ''''''''''''''''''''''''''''''''''
            tmrAtualiza.Enabled = True
            
            With qryGetDocumentoCorrAgConta
                .rdoParameters(0) = Geral.DataProcessamento
                .rdoParameters(1) = aCapa(m_IndexCapa).IdCapa
                Set rsDocumentos = qryGetDocumentoCorrAgConta.OpenResultset(rdOpenStatic, rdConcurReadOnly)
            End With
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Tenta atualizar esta capa "Em Correção", caso não consiga, tenta outra capa '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If Not rsDocumentos.EOF Then
                ''''''''''''''''''''''''''''''''''''
                'Redimensiona o vetor de documentos'
                ''''''''''''''''''''''''''''''''''''
                Erase aDoc
                ReDim aDoc(rsDocumentos.RowCount - 1)
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'Carrega vetor de documentos marcados para correção de Agencia e Conta'
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
                            aDoc(m_IndexDoc).Vinculo = !Vinculo
                            '''''''''''''''''''''''''''''
                            'Obtem o documento pelo tipo'
                            '''''''''''''''''''''''''''''
                            ObtemDocumento aDoc(m_IndexDoc)
                        End With
                     
                     m_IndexDoc = m_IndexDoc + 1
                     rsDocumentos.MoveNext
                Loop
    
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'Carrega imagem do próximo documento à ser corrigido '
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If MostraProximoDocumento Then
                    Exit For
                End If
            End If
        Else
            MsgBox "Não foi possível atualizar o status da capa.", vbCritical
        End If
    Next i
    
    If i > UBound(aCapa) Then CarregarDocumentos = False
        
    Exit Function
    
Erro_CarregarDocumentos:
'    Resume
    Call TratamentoErro("Erro ao obter o documento para correção.", Err, rdoErrors)

End Function

Function EnviarCapaPara() As Boolean

    On Error GoTo Erro_CapaNaoAtualizada
    
    EnviarCapaPara = False

    If Globais.AtualizaStatusCapa(aCapa(m_IndexCapa).IdCapa, "R") Then
        aCapa(m_IndexCapa).Status = "R"
        EnviarCapaPara = True
    End If
    
    Exit Function
    
Erro_CapaNaoAtualizada:
    Call TratamentoErro("Erro ao Atualizar Status da Capa p/ Transmissão.", Err, rdoErrors)

End Function
Private Function DocumentoCorrecao(ByRef pDoc As TpDocumento) As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Procura o documento à ser corrigido: retorna true/false '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    DocumentoCorrecao = False
    
    If UCase(pDoc.Status) = "Y" And pDoc.Vinculo <> 0 Then
        DocumentoCorrecao = True
    End If

End Function
Private Sub txtAgencia_Change()

    If Len(txtAgencia.Text) = txtAgencia.MaxLength Then SendKeys "{TAB}"
    
End Sub
Private Sub txtAgencia_GotFocus()

    SelecionarTexto txtAgencia
    
End Sub
Private Sub txtAgencia_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then txtConta.SetFocus
    
End Sub
Private Sub txtConta_Change()
    If CboTipoConta.Enabled Then
        If Len(txtConta.Text) = txtConta.MaxLength Then CboTipoConta.SetFocus
    Else
        If Len(txtConta.Text) = txtConta.MaxLength And cmdConfirmar.Enabled Then cmdConfirmar.SetFocus
    End If
End Sub
Private Sub txtConta_GotFocus()

    SelecionarTexto txtConta
    
End Sub
Private Sub txtConta_KeyDown(KeyCode As Integer, Shift As Integer)
    If CboTipoConta.Enabled Then
        If KeyCode = 13 Then CboTipoConta.SetFocus
    Else
        If KeyCode = 13 Then cmdConfirmar_Click
    End If
End Sub
