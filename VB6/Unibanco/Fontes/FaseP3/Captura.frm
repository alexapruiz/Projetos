VERSION 5.00
Object = "{A1FF59A3-0995-11CF-B3E8-00608C82AA8D}#1.0#0"; "PIXEZOCX.OCX"
Begin VB.Form Captura 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Captura de Imagens"
   ClientHeight    =   3252
   ClientLeft      =   36
   ClientTop       =   336
   ClientWidth     =   6192
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3252
   ScaleWidth      =   6192
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin PixezocxLib.PixEzImage EzCanon 
      Height          =   456
      Left            =   84
      TabIndex        =   17
      Top             =   2700
      Visible         =   0   'False
      Width           =   552
      _Version        =   65536
      _ExtentX        =   974
      _ExtentY        =   804
      _StockProps     =   96
      TAG_OPEN_DIR    =   ""
      TAG_OPEN_SCHEMA =   ""
      TAG_OPEN_EXT    =   ""
      TAG_OPEN_ROOT   =   ""
      TAG_OPEN_DETECTSCHEMA=   1
      TAG_OPEN_FILENAMES=   ""
      TAG_WINDOW_CURPAGE=   0
      PIXEZ_SELECT    =   ""
      TAG_BORDER_COLOR_BG=   6579300
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
   Begin VB.CommandButton cmdFechar 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   372
      Left            =   3924
      TabIndex        =   2
      Top             =   2736
      Width           =   1536
   End
   Begin VB.CommandButton cmdSemPrioridade 
      Caption         =   "&Sem Prioridade"
      Default         =   -1  'True
      Height          =   372
      Left            =   732
      TabIndex        =   0
      Top             =   2736
      Width           =   1536
   End
   Begin VB.CommandButton cmdComPrioridade 
      Caption         =   "&Com Prioridade"
      Height          =   372
      Left            =   2328
      TabIndex        =   1
      Top             =   2736
      Width           =   1536
   End
   Begin VB.PictureBox Picture5 
      Height          =   312
      Left            =   72
      ScaleHeight     =   264
      ScaleWidth      =   6012
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2316
      Width           =   6060
      Begin VB.Label lblMsg 
         Alignment       =   2  'Center
         Caption         =   "Scanner Inativo"
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
         Height          =   252
         Left            =   36
         TabIndex        =   16
         Top             =   12
         Width           =   5952
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1512
      Left            =   60
      TabIndex        =   5
      Top             =   732
      Width           =   6072
      Begin VB.PictureBox Picture4 
         Height          =   372
         Left            =   3396
         ScaleHeight     =   324
         ScaleWidth      =   1104
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1032
         Width           =   1152
         Begin VB.Label lblQtdeDocto 
            Caption         =   "00000000"
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
            Left            =   48
            TabIndex        =   14
            Top             =   48
            Width           =   1056
         End
      End
      Begin VB.PictureBox Picture3 
         Height          =   372
         Left            =   3396
         ScaleHeight     =   324
         ScaleWidth      =   1104
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   612
         Width           =   1152
         Begin VB.Label lblQtdeCapa 
            Caption         =   "00000000"
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
            Left            =   48
            TabIndex        =   13
            Top             =   48
            Width           =   1056
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   372
         Left            =   3396
         ScaleHeight     =   324
         ScaleWidth      =   1104
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   192
         Width           =   1152
         Begin VB.Label lblNumLote 
            Caption         =   "0000-00000"
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
            Left            =   48
            TabIndex        =   12
            Top             =   48
            Width           =   1032
         End
      End
      Begin VB.Label Label3 
         Caption         =   "Quantidade de Documentos no Lote:"
         Height          =   192
         Left            =   120
         TabIndex        =   8
         Top             =   1128
         Width           =   2688
      End
      Begin VB.Label lblTitQtdeCapa 
         Caption         =   "Quantidade de Envelopes / Malote no Lote:"
         Height          =   192
         Left            =   120
         TabIndex        =   7
         Top             =   708
         Width           =   3192
      End
      Begin VB.Label Label1 
         Caption         =   "Número do Lote Capturado:"
         Height          =   192
         Left            =   120
         TabIndex        =   6
         Top             =   282
         Width           =   2112
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Scanner"
      Height          =   672
      Left            =   60
      TabIndex        =   3
      Top             =   36
      Width           =   6072
      Begin VB.Label lblScanner 
         Caption         =   "Nome do Scanner"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   19.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   372
         Left            =   120
         TabIndex        =   4
         Top             =   180
         Width           =   5832
      End
   End
End
Attribute VB_Name = "Captura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private IdLote                      As Long
Private LoteCapturado               As Boolean
Private FileLog                     As Integer
Private PathIni                     As String
Private Prioridade                  As Integer
Private qryInsereLote               As rdoQuery
Private qryRemoveLote               As rdoQuery
Private qryInsereCapa               As rdoQuery
Private qryInsereDocto              As rdoQuery
Private qryGetControleQualidade     As rdoQuery
Private qryAtualizaStatusLote       As rdoQuery
Private qryGetImagem                As rdoQuery
Private qryProducaoScanner          As rdoQuery
Private rsContQualidade             As rdoResultset

Private Function Digitalizar(Optional ByVal pvbAppend As Boolean = False) As Boolean
    Dim iRet As Long
    Dim NumInicial As Long
    Dim Estacao As Long
    Dim GaugeTop, GaugeLeft As Long
    
    On Error GoTo ErroGetImagem
    rdoErrors.Clear
    
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
    
    Digitalizar = True
    
    On Error GoTo ErroDigitalizar
    
    lblMsg.Caption = "Scanner Capturando"
    Me.Refresh
    If Geral.Scanner = escnVIPS Then

        If Geral.VIPSDLL = eDllProservi Then
            GaugeTop = TwipsYToPixel(Me.Top + 3235)
            GaugeLeft = TwipsXToPixel(Me.Left + 5860)
'            iRet = DigitalizarVIPS(NumInicial, Geral.Estacao, GaugeTop, GaugeLeft, pvbAppend)
            iRet = DigitalizarVIPS(Val(Geral.AgenciaCentral), IdLote, NumInicial, Geral.Estacao, GaugeTop, GaugeLeft, pvbAppend)
            
            'tratamento do erro
            If iRet = -105 Then
                MsgBox "Não foi possível digitalizar todo o lote por uma falha de comunicação do scanner. Desligue a VIPS e ligue-a novamente.", vbExclamation + vbOKOnly, App.Title
            ElseIf iRet <> 1 Then
                MsgBox "Não foi possível digitalizar todo o lote. Codigo de erro: " & iRet, vbExclamation + vbOKOnly, App.Title
            End If
            If iRet <> 1 Then
                Digitalizar = False
            End If
        
        ElseIf Geral.VIPSDLL = eDllNovaUBB Then
            GaugeTop = TwipsYToPixel(Me.Top + 3235)
            GaugeLeft = TwipsXToPixel(Me.Left + 5860)
            iRet = DigitalizarVIPS(Val(Geral.AgenciaCentral), Val(Right(CStr(IdLote), 5)), 1, Geral.Estacao, GaugeTop, GaugeLeft, pvbAppend)
            
            'tratamento do erro
            If iRet = SC_Erro Then
                Digitalizar = False
            End If
        
        Else
            iRet = VIPS_Captura(Val(Geral.AgenciaCentral), IdLote, Geral.DiretorioDados & Geral.RetornoFinal, pvbAppend)
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

        GaugeTop = TwipsYToPixel(Me.Top + 3235)
        GaugeLeft = TwipsXToPixel(Me.Left + 5860)
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
    End If
    If Digitalizar Then
        lblMsg.Caption = "Scanner Inativo"
    Else
        lblMsg.Caption = "Falha da captura das imagens"
        Print #FileLog, "Usuario: " & Geral.Usuario & " Horario: " & Format(Now, "hh:mm:ss") & " - Erro na captura das imagens dos documentos"
    End If
    Me.Refresh
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

Private Function DigitalizarVIPS(ByVal lngAgencia As Long, ByVal lngLote As Long, ByVal pvnInicio As Long, ByVal pvnEstacao As Long, ByVal nTop As Long, ByVal nLeft As Long, Optional ByVal pvbAppend As Boolean = False) As Long
    
    Dim iRet    As Long
    Dim iRet1   As Long
    Dim iRet2   As Long
    
    If Geral.VIPSDLL = eDllNovaUBB Then
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
        
    Else
        iRet = MC93_SetGaugePos(nTop, nLeft)
        If pvbAppend Then
            MC93_SetAppend (1)
        Else
            MC93_SetAppend (0)
        End If

        iRet = MC93_Digitaliza(pvnInicio, Geral.DiretorioDados & Geral.RetornoFinal, pvnEstacao)
        If iRet = -105 Then
            MC93_DeInit
            MC93_Init
        End If
        
        'se ocorrer erro, reinicializa o scanner
        If iRet <> 1 Then
            iRet1 = MC93_Reset
        End If
    
        'retorna o cod. de retorno da vips
        DigitalizarVIPS = iRet
    End If
    
End Function

Private Function DigitalizarCanonLS500(ByVal NumInicial As Long, ByVal Estacao As Long, ByVal GaugeTop As Long, ByVal GaugeLeft As Long, Optional ByVal pvbAppend As Boolean = False) As Long
    
    Dim iRet, Ret_LS, Ret_Canon As Long
    Dim ImgInicialCanon As Long
    Dim ArqCanon As String
    Dim NumLote As Long
    Dim Retc As Integer
    Dim Arq As String
    
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
   
    If Ret_LS <> 1 Then
        DigitalizarCanonLS500 = Ret_LS ' Erro LS
        Exit Function
    End If
    
Redigitaliza_Canon:
    
    ImgInicialCanon = LS_GetLastImage()
    NumLote = LS_GetNumLote()
   
On Error GoTo ERRO_CANON
       
    '////////////////////////////////////////////////\\
    '//   Digitaliza Doctos via PIXEL TRANSLATIONS   \\
    '////////////////////////////////////////////////\\
    
    ArqCanon = "Lote" & Format$(ImgInicialCanon, "0000") & ".TIF"
    
    '---- COORDENADAS BÁSICAS PARA PIXEL ----
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

    If (Ret_Canon <> 0) Then
        Beep
        Screen.MousePointer = vbDefault
        MsgBox "Desligue o scanner CANON e ligue-o novamente para redigitalizar somente os documentos deste equipamento. Codigo de erro: " & Ret_Canon, vbCritical + vbOKOnly, App.Title
        Screen.MousePointer = vbHourglass
        EzCanon.Close    'fecha pixel para não dar acesso denied
        If Dir(Geral.DiretorioImagens & ArqCanon) <> "" Then
            Kill Geral.DiretorioImagens & ArqCanon
        End If
        GoTo Redigitaliza_Canon
    End If
    
    'verifica se foi gravada alguma imagem
    If (EzCanon.PageCount = 0) Then
        Beep
        Screen.MousePointer = vbDefault
        MsgBox "Não existem documentos para digitalização no scanner Canon. Coloque os documentos no scanner.", vbExclamation + vbOKOnly, App.Title
        Screen.MousePointer = vbHourglass
        EzCanon.Close    'fecha pixel para não dar acesso denied
        If Dir(Geral.DiretorioImagens & ArqCanon) <> "" Then
            Kill Geral.DiretorioImagens & ArqCanon
        End If
        GoTo Redigitaliza_Canon
    End If
    
    EzCanon.Close    'fecha pixel para não dar acesso denied
    EzCanon.Visible = False
    
    iRet = UT_DesmembraTiff(Geral.DiretorioImagens & ArqCanon, Geral.DiretorioDados & Geral.RetornoFinal, NumLote, ImgInicialCanon, Estacao)
    If iRet = 0 Then
        DigitalizarCanonLS500 = -4  ' Erro ao Desmembrar
    Else
        DigitalizarCanonLS500 = 1   ' OK
    End If
    
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

Private Sub cmdComPrioridade_Click()
    Prioridade = 1
    
    If Geral.VIPSDLL = eDllNovaUBB Then
        If bInicializou = False Then
            MsgBox "VIPS não foi inicializada, favor verificar!", vbCritical, App.Title
            Exit Sub
        End If
    End If
    
    IniciarDigitalizacao
    
End Sub

Private Sub CmdFechar_Click()
    Unload Me
End Sub

Private Sub cmdSemPrioridade_Click()
    Prioridade = 0
    
    If Geral.VIPSDLL = eDllNovaUBB Then
        If bInicializou = False Then
            MsgBox "VIPS não foi inicializada, favor verificar!", vbCritical, App.Title
            Exit Sub
        End If
    End If
    
    IniciarDigitalizacao
    
End Sub

Private Sub Form_Activate()
   'Inclusão de chamada a rotina AtualizaAtividade
   Call AtualizaAtividade(5)
End Sub

Private Sub Form_Load()
    
    FileLog = FreeFile
    Open Geral.DiretorioTrabalho & "DIG" & Format(Now, "ddmm") & ".TXT" For Append As #FileLog
    
    '''''''''''
    'Grava log'
    '''''''''''
    GravaLog 0, 0, 42
    
    
    If Geral.Scanner = escnCanonLS500 Then
        lblScanner.Caption = "Scanner Canon e Scanner LS-500"
        Call LS_SetAppend(0)  'default - não faz append na LS500
    ElseIf Geral.Scanner = escnVIPS Then
        lblScanner.Caption = "Scanner VIPS MC93"
    ElseIf Geral.Scanner = escnDummy Then
        lblScanner.Caption = "Simulação de Scanner"
    End If
    
    Set qryInsereLote = Geral.Banco.CreateQuery("", "{? = call InsereLote (?,?,?,?)}")
    Set qryRemoveLote = Geral.Banco.CreateQuery("", "{? = call RemoveLote (?,?)}")
    Set qryInsereCapa = Geral.Banco.CreateQuery("", "{? = call CapturaCapa (?,?,?,?,?)}")
    Set qryInsereDocto = Geral.Banco.CreateQuery("", "{? = call CapturaDocumento (?,?,?,?,?,?,?,?,?,?)}")
    Set qryGetImagem = Geral.Banco.CreateQuery("", "{? = call GetImagem (?,?)}")
    Set qryProducaoScanner = Geral.Banco.CreateQuery("", "{ call InsereProducaoScanner (?,?,?,?)}")
    Set qryGetControleQualidade = Geral.Banco.CreateQuery("", "{ call GetControleQualidade(?)}")
    Set qryAtualizaStatusLote = Geral.Banco.CreateQuery("", "{ ? = call AtualizaStatusLote (?,?,?)}")

    PathIni = String(256, " ")
    'obtem o path do windows para leitura/gravação do canon.ini
    GetWindowsDirectory PathIni, 255       'obtem o diretorio do windows
    PathIni = Trim(PathIni)
    PathIni = Left(PathIni, Len(PathIni) - 1) & "\"
    
    lblNumLote.Caption = ""
    lblQtdeCapa.Caption = ""
    lblQtdeDocto.Caption = ""
    lblMsg.Caption = "Scanner Inativo"
End Sub

Private Sub Form_Unload(Cancel As Integer)

    GravaLog 0, 0, 43
    
    Close #FileLog
    qryInsereLote.Close
    qryRemoveLote.Close
    qryInsereCapa.Close
    qryInsereDocto.Close
    qryGetImagem.Close
    qryProducaoScanner.Close
    qryGetControleQualidade.Close
    qryAtualizaStatusLote.Close
End Sub

Private Sub IniciarDigitalizacao()
    Dim dtInicio                As Date
    Dim dtFim                   As Date
    Dim bDigitalizou            As Boolean
    Dim bAppend                 As Boolean
    Dim sTimer                  As String
    Dim iVirgula                As Long
    Dim Frente                  As String
    Dim Verso                   As String
    Dim Opcao                   As Integer
    Dim Count                   As Integer
    Dim Proc_ConfLote           As New ConfirmacaoLote
    Dim sFormatoData            As String
    Dim sNomeComputador         As String
    Dim sDiretorioDestino       As String
    Dim sDiretorioBULK          As String
    Dim bExcluirArquivoCriacao  As Boolean

    'lblNumLote.Caption = ""
    'lblQtdeCapa.Caption = ""
    'lblQtdeDocto.Caption = ""

    cmdSemPrioridade.Enabled = False
    cmdComPrioridade.Enabled = False
    cmdFechar.Enabled = False
    
    Screen.MousePointer = vbHourglass
    
    On Error GoTo ErroInsereLote
    
    IdLote = 0
    LoteCapturado = False
    
    If Not InsereLote Then
        GoTo ErroInsereLote
    End If
    
    On Error GoTo ErroDigitalizacao
    '''''''''''''''''''''''''
    ' Digitalizar Documento '
    '''''''''''''''''''''''''
    bAppend = False
    sTimer = CStr(Timer)
    iVirgula = InStr(sTimer, ",")
    sTimer = Left(sTimer, iVirgula - 1) & Mid(sTimer, iVirgula + 1)
    If Geral.VIPSDLL = eDllProservi Then
        Geral.RetornoFinal = Format(Now, "yyyymmddhhnnss") & sTimer & ".TXT"
    Else
        Geral.RetornoFinal = Format(Geral.DataProcessamento, "00000000") & Format(IdLote, "000000000") & ".txt"
    End If

    dtInicio = Now  ' Inicio da digitalizacao
    Print #FileLog, "Usuario: " & Geral.Usuario & " Horario: " & Format(Now, "hh:mm:ss") & " - Inicio da captura"

Redigitalizar:
    
    'Captura - Incicio Captura'
    GravaLog 0, 0, 44
    
    bDigitalizou = Digitalizar(bAppend)
    
    dtFim = Now     ' Fim da digitalizacao
    Print #FileLog, "Usuario: " & Geral.Usuario & " Horario: " & Format(Now, "hh:mm:ss") & " - Fim da captura"
    
    If Not bDigitalizou Then
        If Geral.Scanner = escnCanonLS500 Then
            MsgBox "Não foi possível concluir a digitalização.", vbQuestion + vbOKOnly, App.Title
            GoTo FinalDigitalizacao
        Else
            Screen.MousePointer = vbDefault
            If MsgBox("Não foi possível concluir a digitalização. Deseja continuar no mesmo lote?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                bAppend = True
                GoTo Redigitalizar
            Else
                bAppend = False
            End If
        End If
    End If
    
    'Captura - Fim Captura'
    GravaLog 0, 0, 45
    
    ''''''''''''''''''''''''''''''''''''''
    ' Processar retorno da digitalizacao '
    ''''''''''''''''''''''''''''''''''''''
ReprocessarArquivo:
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Varre o arquivo de retorno e obtem o nome da ultima imagem'
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Not ObtemUltimaImagem(Geral.DiretorioDados & Geral.RetornoFinal, Frente, Verso) Then
        Screen.MousePointer = vbDefault
        MsgBox "Não foi possível obter a última imagem capturada. Repasse os documentos.", vbExclamation + vbOKOnly, App.Title
        Print #FileLog, "Usuario: " & Geral.Usuario & " Horario: " & Format(Now, "hh:mm:ss") & " - Nao foi localizada a imagem do ultimo documento"
        GoTo FinalDigitalizacao
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Apresenta img do ult. docto e espera confirmação fechamento do lote'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Geral.Scanner = escnVIPS Then
        Load ConfFimLote
        On Error Resume Next
        'verso
        With ConfFimLote.Lead2
           .AutoRepaint = False
           If Geral.VIPSDLL = eDllProservi Then
             .Load Geral.DiretorioImagens & Verso, 0, 0, 1
           Else
             .Load Geral.DiretorioImagens & Format(IdLote, "000000000") & "\" & Verso, 0, 0, 1
           End If
           .Intensity 220
           .PaintZoomFactor = 100
           .AutoRepaint = True
        End With
        'frente
        With ConfFimLote.Lead1
           .AutoRepaint = False
           If Geral.VIPSDLL = eDllProservi Then
               .Load Geral.DiretorioImagens & Frente, 0, 0, 1
           Else
               .Load Geral.DiretorioImagens & Format(IdLote, "000000000") & "\" & Frente, 0, 0, 1
           End If
           .Intensity 220
           .PaintZoomFactor = 100
           .AutoRepaint = True
        End With
        On Error GoTo ErroDigitalizacao

        Screen.MousePointer = vbDefault
        'confirma se lote deve ser gravado
        ConfFimLote.Show vbModal, Me
        Opcao = ConfFimLote.Resposta
        Unload ConfFimLote
        Principal.Refresh

        Select Case Opcao
            Case 0 ' Cancelou o Lote
                Print #FileLog, "Usuario: " & Geral.Usuario & " Horario: " & Format(Now, "hh:mm:ss") & " - Cancelamento da captura do lote"
                MsgBox "Redigitalize este lote, pois o mesmo não foi gravado.", vbExclamation + vbOKOnly, App.Title
                LoteCapturado = False
                GoTo FinalDigitalizacao
            Case 1 ' Confirmou o Lote
                ' continua o processamento
                Print #FileLog, "Usuario: " & Geral.Usuario & " Horario: " & Format(Now, "hh:mm:ss") & " - Confirmacao da captura do lote"
            Case 2 ' Continuar capturando no mesmo lote
                Print #FileLog, "Usuario: " & Geral.Usuario & " Horario: " & Format(Now, "hh:mm:ss") & " - Continuacao da captura no mesmo lote"
                bAppend = True
                GoTo Redigitalizar
        End Select
    End If
    
    Screen.MousePointer = vbHourglass
    
    '''''''''''''''''''''''''''''''''''''''''''''''''
    ' Varre o arquivo e grava os documentos
    '''''''''''''''''''''''''''''''''''''''''''''''''
    Print #FileLog, "Usuario: " & Geral.Usuario & " Horario: " & Format(Now, "hh:mm:ss") & " - Inicio do processamento do arquivo: " & Geral.RetornoFinal
    
    '''''''''''''''''''''''''''
    'Pega o nome do computador'
    '''''''''''''''''''''''''''
    sNomeComputador = ComputerName()
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Pega o diretorio de destino de gravacao dos arquivos'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sDiretorioDestino = PegarOpcaoINI("Diversos", "GravacaoLote", "\\MDI_NT2\DRIVE_D\MDI_UBB\DADOS\")
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Pega o diretorio de gravacao dos arquivos para o BULK INSERT'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sDiretorioBULK = PegarOpcaoINI("Diversos", "DiretorioBULK", "D:\MDI_UBB\DADOS\")
    ''''''''''''''''''''''''''''''''''''''''''''''
    'Pega o formato da data para gravacao do lote'
    ''''''''''''''''''''''''''''''''''''''''''''''
    sFormatoData = PegarOpcaoINI("Diversos", "FormatoData", "dd-mm-yyyy hh:mm:00")
    '''''''''''''''''''''''''''''''''''''''''''''
    'Pega flag de exclusão do arquivo de criação'
    '''''''''''''''''''''''''''''''''''''''''''''
    bExcluirArquivoCriacao = CBool(PegarOpcaoINI("Diversos", "ExcluirArquivoCriacao", "0"))
    
    Proc_ConfLote.SetConnection Geral.Banco
    Proc_ConfLote.DataProcessamento = Geral.DataProcessamento
    Proc_ConfLote.IdLote = IdLote
    Proc_ConfLote.ArquivoCapaCriacao = "C_" & sNomeComputador & "_" & Geral.RetornoFinal
    Proc_ConfLote.ArquivoDoctoCriacao = "D_" & sNomeComputador & "_" & Geral.RetornoFinal
    Proc_ConfLote.ArquivoOrigem = Geral.RetornoFinal
    Proc_ConfLote.DiretorioOrigem = Geral.DiretorioDados
    Proc_ConfLote.DiretorioCriacao = Geral.DiretorioDados & Geral.DataProcessamento
    Proc_ConfLote.DiretorioDestino = sDiretorioDestino
    Proc_ConfLote.DiretorioBULK = sDiretorioBULK
    Proc_ConfLote.FormatoData = sFormatoData
    Proc_ConfLote.ExcluirArquivoCriacao = bExcluirArquivoCriacao
    
    'Captura - Inicio Confirmar lote
    GravaLog 0, 0, 46
    
    Do While Proc_ConfLote.Processar() = False
        Screen.MousePointer = vbDefault
        If MsgBox("Houve um erro na gravação do lote." & vbCrLf & _
               "Deseja tentar gravar novamente?", vbQuestion + vbYesNo, App.Title) = vbNo Then

            MsgBox "Redigitalize este lote, pois o mesmo não foi gravado.", vbExclamation + vbOKOnly, App.Title
            Print #FileLog, "Usuario: " & Geral.Usuario & " Horario: " & Format(Now, "hh:mm:ss") & " - Erro no processamento do arquivo de retorno, documentos nao gravados"
            LoteCapturado = False
            GravaLog 0, 0, 48
            GoTo FinalDigitalizacao

        End If
    Loop
    
    'Captura - Fim Confirmar lote
    GravaLog 0, 0, 47
    ''''''''''''''''''''
    'Se chegou ate aqui'
    ''''''''''''''''''''
    LoteCapturado = True
    
    lblNumLote.Caption = Format(IdLote, "0000-00000")
    lblQtdeCapa.Caption = Format(Proc_ConfLote.TotalCapas, "0000")
    lblQtdeDocto.Caption = Format(Proc_ConfLote.TotalDocumentos - Proc_ConfLote.TotalCapas, "0000")
    
    AtualizaStatusLote
    
'    While Not ProcessaArquivoRetorno(Geral.DiretorioDados & Geral.RetornoFinal, Count)
'        Screen.MousePointer = vbDefault
'        ' erro na gravacao dos doctos do lote
'        If MsgBox("Houve um erro na gravação do lote." & vbCrLf & _
'               "Deseja tentar gravar novamente?", vbQuestion + vbYesNo, App.Title) = vbNo Then
'
'            MsgBox "Redigitalize este lote, pois o mesmo não foi gravado.", vbExclamation + vbOKOnly, App.Title
'            Print #FileLog, "Usuario: " & Geral.Usuario & " Horario: " & Format(Now, "hh:mm:ss") & " - Erro no processamento do arquivo de retorno, documentos nao gravados"
'            LoteCapturado = False
'            GoTo FinalDigitalizacao
'
'        End If
'
'    Wend
    
    Print #FileLog, "Usuario: " & Geral.Usuario & " Horario: " & Format(Now, "hh:mm:ss") & " - Termino do processamento do arquivo de retorno"
    
    If Geral.Scanner <> escnDummy And Geral.Scanner <> escnSemScanner Then
       With qryProducaoScanner
            .rdoParameters(0) = Now
            If Geral.Scanner = escnVIPS Then
                .rdoParameters(1) = "VIPS"
            ElseIf Geral.Scanner = escnCanonLS500 Then
                .rdoParameters(1) = "LS500 e Canon"
            ElseIf Geral.Scanner = escnLS500 Then
                .rdoParameters(1) = "LS500"
            ElseIf Geral.Scanner = escnCanon Then
                .rdoParameters(1) = "Canon"
            End If
            If DateDiff("s", dtInicio, dtFim) > 0 Then
                .rdoParameters(2) = DateDiff("s", dtInicio, dtFim)
            Else
                .rdoParameters(2) = 1
            End If
            .rdoParameters(3) = Count
            .Execute
        End With
    End If
    
    '''''''''''''''''''''''''''
    ' Finalizar digitalização '
    '''''''''''''''''''''''''''
    GoTo FinalDigitalizacao

ErroInsereLote:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Não foi possível inserir novo Lote.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    GoTo FinalDigitalizacao
    Exit Sub
    
ErroDigitalizacao:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro inesperado na Captura de imagens.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    GoTo FinalDigitalizacao
    Exit Sub

FinalDigitalizacao:
    Screen.MousePointer = vbDefault
    cmdSemPrioridade.Enabled = True
    cmdComPrioridade.Enabled = True
    cmdFechar.Enabled = True
    
    cmdSemPrioridade.SetFocus

    If Not LoteCapturado Then
        If IdLote > 0 Then
            With qryRemoveLote
                .rdoParameters(1) = Geral.DataProcessamento
                .rdoParameters(2) = IdLote
                .Execute
            End With
        End If
    End If
End Sub
Private Function ObtemUltimaImagem(ByVal NomeArq As String, _
                                    ByRef Frente As String, ByRef Verso As String) As Boolean
    Dim Arq As Integer
    Dim RetornoProservi As tpRetornoFinal
    Dim RetornoUnibanco As tpRetornoVips
    Dim RetornoNovaDll  As tpRetornoVipsNovaDLL
    
    Arq = FreeFile
    Open NomeArq For Binary As #Arq
    
    If Geral.VIPSDLL = eDllProservi Then
        RetornoProservi.Frente = String(12, "  ")
        RetornoProservi.Verso = String(12, "  ")
        
        Get #Arq, , RetornoProservi
        While Not EOF(Arq)
            Frente = RetornoProservi.Frente
            Verso = IIf(Trim(RetornoProservi.Verso) = "", RetornoProservi.Frente, RetornoProservi.Verso)
            Get #Arq, , RetornoProservi
        Wend
    ElseIf Geral.VIPSDLL = eDllNovaUBB Then
        RetornoNovaDll.Frente = String(19, "  ")
        RetornoNovaDll.Verso = String(19, "  ")
        
        Get #Arq, , RetornoNovaDll
        While Not EOF(Arq)
            Frente = RetornoNovaDll.Frente
            Verso = IIf(Trim(RetornoNovaDll.Verso) = "", RetornoNovaDll.Frente, RetornoNovaDll.Verso)
            Get #Arq, , RetornoNovaDll
        Wend
    
    Else
        RetornoUnibanco.Frente = String(19, "  ")
        RetornoUnibanco.Verso = String(19, "  ")
        
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
        If Geral.VIPSDLL = eDllProservi Then
            ProcessaArquivoRetorno = ProcessaArquivoRetornoProservi(NomeArq, Count)
        Else
            ProcessaArquivoRetorno = ProcessaArquivoRetornoUnibanco(NomeArq, Count)
        End If
    End If
End Function

Private Function ProcessaArquivoRetornoProservi(ByVal NomeArq As String, ByRef Count As Integer) As Boolean
    Dim Arq             As Integer
    Dim IdCapa          As Long
    Dim TipoDoc         As Integer
    Dim CountCapas      As Integer
    Dim OrdemCaptura    As Integer
    Dim Campo1          As String
    Dim Campo2          As String
    Dim Campo3          As String
    Dim Valor           As String
    Dim bVirtual        As Boolean
    Dim IdEnv_Mal       As String
    Dim Linha           As tpRetornoFinal
    
    Count = 0
    CountCapas = 0
    IdCapa = 0
    OrdemCaptura = 1
    
    On Error GoTo ErroCaptura
    
    Geral.Banco.BeginTrans
    
    Arq = FreeFile
    Open NomeArq For Binary As #Arq
    
    Get #Arq, , Linha
    While Not EOF(Arq)
        Linha.Leitura = TrataLeitura(Linha.Leitura)
        TipoDoc = 0
        If VerificaSeCapa(Linha.Leitura, IdEnv_Mal) Then
            If IdEnv_Mal = "M" Then
                TratarCamposCMC7 Linha.Leitura, Campo1, Campo2, Campo3, Valor
            End If
            TipoDoc = 1 ' Capa
            bVirtual = False
            ' gravar capa
            qryInsereCapa.rdoParameters(0).Direction = rdParamReturnValue
            qryInsereCapa.rdoParameters(1) = Geral.DataProcessamento
            qryInsereCapa.rdoParameters(2) = IdLote
            qryInsereCapa.rdoParameters(3) = IdEnv_Mal
            qryInsereCapa.rdoParameters(4) = IIf(IdEnv_Mal = "E", Val(Trim(Linha.Leitura)), "0" & Mid(Campo2, 1, 9) & Mid(Campo1, 4, 4))
            qryInsereCapa.rdoParameters(5).Direction = rdParamOutput
            qryInsereCapa.Execute
            If qryInsereCapa.rdoParameters(0) <> 0 Then
                GoTo ErroCaptura
            End If
            IdCapa = qryInsereCapa.rdoParameters(5)
            GravaLog IdCapa, 0, 40
            CountCapas = CountCapas + 1
            'Inicia a Ordem de Captura
            OrdemCaptura = 1
        ElseIf IdCapa = 0 Then
            TipoDoc = 1 ' Capa
            bVirtual = True
            ' gravar capa
            qryInsereCapa.rdoParameters(0).Direction = rdParamReturnValue
            qryInsereCapa.rdoParameters(1) = Geral.DataProcessamento
            qryInsereCapa.rdoParameters(2) = IdLote
            qryInsereCapa.rdoParameters(3) = "E" ' Capa virtual sempre sera envelope
            qryInsereCapa.rdoParameters(4) = 9 ' Capa Virtual
            qryInsereCapa.rdoParameters(5).Direction = rdParamOutput
            qryInsereCapa.Execute
            If qryInsereCapa.rdoParameters(0) <> 0 Then
                GoTo ErroCaptura
            End If
            IdCapa = qryInsereCapa.rdoParameters(5)
            GravaLog IdCapa, 0, 40
            CountCapas = CountCapas + 1
            'Inicia a Ordem de Captura
            OrdemCaptura = 1
        End If
        ' gravar documento
        Valor = "000"
        qryInsereDocto.rdoParameters(0).Direction = rdParamReturnValue
        qryInsereDocto.rdoParameters(10).Direction = rdParamOutput
        qryInsereDocto.rdoParameters(1) = Geral.DataProcessamento
        qryInsereDocto.rdoParameters(2) = IdCapa
        qryInsereDocto.rdoParameters(3) = TipoDoc
        If IdEnv_Mal = "M" And TipoDoc = 1 Then ' Capa de Malote
            If Not bVirtual Then
                qryInsereDocto.rdoParameters(4) = "0" & Mid(Campo2, 1, 9) & Mid(Campo1, 4, 4)
            Else
                qryInsereDocto.rdoParameters(4) = "9" ' Capa Virtual
            End If
        ElseIf IdEnv_Mal = "E" And TipoDoc = 1 Then ' Capa de Envelope
            If Not bVirtual Then
                qryInsereDocto.rdoParameters(4) = Trim(Linha.Leitura)
            Else
                qryInsereDocto.rdoParameters(4) = "9" ' Capa Virtual
            End If
        ElseIf Linha.Tipo = "A" Then ' Docto com CMC7
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
        qryInsereDocto.rdoParameters(9) = OrdemCaptura
        qryInsereDocto.Execute
        If qryInsereDocto.rdoParameters(0) <> 0 Then
            GoTo ErroCaptura
        End If
        GravaLog IdCapa, qryInsereDocto.rdoParameters(10), 41
        Count = Count + 1
        OrdemCaptura = OrdemCaptura + 1
        Get #Arq, , Linha
    Wend
    Close #Arq
    Geral.Banco.CommitTrans
    AtualizaStatusLote
    LoteCapturado = True
    ProcessaArquivoRetornoProservi = True
        
    lblNumLote.Caption = Format(IdLote, "0000-00000")
    lblQtdeCapa.Caption = Format(CountCapas, "0000")
    lblQtdeDocto.Caption = Format(Count - CountCapas, "0000")
    Exit Function
    
ErroCaptura:
    If IdCapa >= 0 Then
        Geral.Banco.RollbackTrans
    End If
    
    LoteCapturado = False
    
    TratamentoErro "Erro no processamento do arquivo de retorno.", Err, rdoErrors, False
    MsgBox "Erro no processamento do arquivo de retorno.", vbCritical + vbOKOnly, App.Title
    ProcessaArquivoRetornoProservi = False
End Function

Private Function ProcessaArquivoRetornoUnibanco(ByVal NomeArq As String, ByRef Count As Integer) As Boolean
    Dim Arq             As Integer
    Dim IdCapa          As Long
    Dim TipoDoc         As Integer
    Dim CountCapas      As Integer
    Dim OrdemCaptura    As Integer
    Dim Campo1          As String
    Dim Campo2          As String
    Dim Campo3          As String
    Dim Valor           As String
    Dim bVirtual        As Boolean
    Dim IdEnv_Mal       As String
    Dim Linha           As tpRetornoVips
    
    Count = 0
    CountCapas = 0
    IdCapa = 0
    OrdemCaptura = 1
    
    On Error GoTo ErroCaptura
    
    Geral.Banco.BeginTrans
    
    Arq = FreeFile
    Open NomeArq For Binary As #Arq
    
    Get #Arq, , Linha
    While Not EOF(Arq)
        Linha.Leitura = TrataLeitura(Linha.Leitura)
        TipoDoc = 0
        If VerificaSeCapa(Linha.Leitura, IdEnv_Mal) Then
            If IdEnv_Mal = "M" Then
                TratarCamposCMC7 Linha.Leitura, Campo1, Campo2, Campo3, Valor
            End If
            TipoDoc = 1 ' Capa
            bVirtual = False
            ' gravar capa
            qryInsereCapa.rdoParameters(0).Direction = rdParamReturnValue
            qryInsereCapa.rdoParameters(1) = Geral.DataProcessamento
            qryInsereCapa.rdoParameters(2) = IdLote
            qryInsereCapa.rdoParameters(3) = IdEnv_Mal
            qryInsereCapa.rdoParameters(4) = IIf(IdEnv_Mal = "E", Val(Trim(Linha.Leitura)), "0" & Mid(Campo2, 1, 9) & Mid(Campo1, 4, 4))
            qryInsereCapa.rdoParameters(5).Direction = rdParamOutput
            qryInsereCapa.Execute
            If qryInsereCapa.rdoParameters(0) <> 0 Then
                GoTo ErroCaptura
            End If
            IdCapa = qryInsereCapa.rdoParameters(5)
            GravaLog IdCapa, 0, 40
            CountCapas = CountCapas + 1
            OrdemCaptura = 1
        ElseIf IdCapa = 0 Then
            TipoDoc = 1 ' Capa
            bVirtual = True
            ' gravar capa
            qryInsereCapa.rdoParameters(0).Direction = rdParamReturnValue
            qryInsereCapa.rdoParameters(1) = Geral.DataProcessamento
            qryInsereCapa.rdoParameters(2) = IdLote
            qryInsereCapa.rdoParameters(3) = "E" ' Capa virtual sempre sera envelope
            qryInsereCapa.rdoParameters(4) = 9 ' Capa Virtual
            qryInsereCapa.rdoParameters(5).Direction = rdParamOutput
            qryInsereCapa.Execute
            If qryInsereCapa.rdoParameters(0) <> 0 Then
                GoTo ErroCaptura
            End If
            IdCapa = qryInsereCapa.rdoParameters(5)
            GravaLog IdCapa, 0, 40
            CountCapas = CountCapas + 1
            OrdemCaptura = 1
        End If
        ' gravar documento
        Valor = "000"
        qryInsereDocto.rdoParameters(0).Direction = rdParamReturnValue
        qryInsereDocto.rdoParameters(10).Direction = rdParamOutput
        qryInsereDocto.rdoParameters(1) = Geral.DataProcessamento
        qryInsereDocto.rdoParameters(2) = IdCapa
        qryInsereDocto.rdoParameters(3) = TipoDoc
        If IdEnv_Mal = "M" And TipoDoc = 1 Then ' Capa de Malote
            If Not bVirtual Then
                qryInsereDocto.rdoParameters(4) = "0" & Mid(Campo2, 1, 9) & Mid(Campo1, 4, 4)
            Else
                qryInsereDocto.rdoParameters(4) = "9" ' Capa Virtual
            End If
        ElseIf IdEnv_Mal = "E" And TipoDoc = 1 Then ' Capa de Envelope
            If Not bVirtual Then
                qryInsereDocto.rdoParameters(4) = Trim(Linha.Leitura)
            Else
                qryInsereDocto.rdoParameters(4) = "9" ' Capa Virtual
            End If
        ElseIf Linha.Tipo = "A" Then ' Docto com CMC7
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
        qryInsereDocto.rdoParameters(9) = OrdemCaptura
        qryInsereDocto.Execute
        If qryInsereDocto.rdoParameters(0) <> 0 Then
            GoTo ErroCaptura
        End If
        GravaLog IdCapa, qryInsereDocto.rdoParameters(10), 41
        Count = Count + 1
        OrdemCaptura = OrdemCaptura + 1
        Get #Arq, , Linha
    Wend
    Close #Arq
    Geral.Banco.CommitTrans
    AtualizaStatusLote
    LoteCapturado = True
    ProcessaArquivoRetornoUnibanco = True

    lblNumLote.Caption = Format(IdLote, "0000-00000")
    lblQtdeCapa.Caption = Format(CountCapas, "0000")
    lblQtdeDocto.Caption = Format(Count - CountCapas, "0000")
    Exit Function

ErroCaptura:
    If IdCapa >= 0 Then
        Geral.Banco.RollbackTrans
    End If

    LoteCapturado = False

    TratamentoErro "Erro no processamento do arquivo de retorno.", Err, rdoErrors, False
    MsgBox "Erro no processamento do arquivo de retorno.", vbCritical + vbOKOnly, App.Title
    ProcessaArquivoRetornoUnibanco = False

End Function
Private Function ProcessaArquivoRetornoCanonLS500(ByVal NomeArq As String, ByRef Count As Integer) As Boolean
    Dim Arq             As Integer
    Dim IdCapa          As Long
    Dim TipoDoc         As Integer
    Dim CountCapas      As Integer
    Dim OrdemCaptura    As Integer
    Dim Campo1          As String
    Dim Campo2          As String
    Dim Campo3          As String
    Dim Valor           As String
    Dim bVirtual        As Boolean
    Dim IdEnv_Mal       As String
    Dim Linha           As tpRetornoFinal
    Dim DirDestino      As String
    
    Count = 0
    CountCapas = 0
    IdCapa = 0
    OrdemCaptura = 1
    
    On Error Resume Next
    DirDestino = Geral.DiretorioImagens & Format(IdLote, "000000000")
    MkDir (DirDestino)
    
    On Error GoTo ErroCaptura
    
    Geral.Banco.BeginTrans
    
    ' grava primeiro registros com codigo de barras
    Arq = FreeFile
    Open NomeArq For Binary As #Arq
    Get #Arq, , Linha
    Do While Not EOF(Arq)
        If Linha.Tipo <> "A" Then
        
            IdEnv_Mal = "E"
            If IdCapa = 0 Then
                TipoDoc = 1 ' Capa
                ' gravar capa
                qryInsereCapa.rdoParameters(0).Direction = rdParamReturnValue
                qryInsereCapa.rdoParameters(1) = Geral.DataProcessamento
                qryInsereCapa.rdoParameters(2) = IdLote
                qryInsereCapa.rdoParameters(3) = IdEnv_Mal
                qryInsereCapa.rdoParameters(4) = 9
                qryInsereCapa.rdoParameters(5).Direction = rdParamOutput
                qryInsereCapa.Execute
                If qryInsereCapa.rdoParameters(0) <> 0 Then
                    GoTo ErroCaptura
                End If
                IdCapa = qryInsereCapa.rdoParameters(5)
                GravaLog IdCapa, 0, 40
                CountCapas = CountCapas + 1
                OrdemCaptura = 1
            Else
                TipoDoc = 0
            End If
            
            ' gravar documento
            Valor = "000"
            qryInsereDocto.rdoParameters(0).Direction = rdParamReturnValue
            qryInsereDocto.rdoParameters(10).Direction = rdParamOutput
            qryInsereDocto.rdoParameters(1) = Geral.DataProcessamento
            qryInsereDocto.rdoParameters(2) = IdCapa
            qryInsereDocto.rdoParameters(3) = TipoDoc
            qryInsereDocto.rdoParameters(4) = IIf(TipoDoc = 1, "9", Trim(Linha.Leitura))
            qryInsereDocto.rdoParameters(5) = Val(Valor) / 100
            qryInsereDocto.rdoParameters(6) = Linha.Frente
            qryInsereDocto.rdoParameters(7) = IIf(Trim(Linha.Verso) = "", Linha.Frente, Linha.Verso)
            qryInsereDocto.rdoParameters(8) = Linha.Origem
            qryInsereDocto.rdoParameters(9) = OrdemCaptura
            qryInsereDocto.Execute
            If qryInsereDocto.rdoParameters(0) <> 0 Then
                GoTo ErroCaptura
            End If
            GravaLog IdCapa, qryInsereDocto.rdoParameters(10), 41
            Count = Count + 1
            OrdemCaptura = OrdemCaptura + 1
        
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
    
    ' se nao conseguiu localizar capa
    If IdCapa = 0 Then
        MsgBox "Não foi possivel localizar Capa de Envelope / Malote.", vbExclamation + vbOKOnly, App.Title
        GoTo ErroCaptura
    End If
    
    Arq = FreeFile
    Open NomeArq For Binary As #Arq
    Get #Arq, , Linha
    'OrdemCaptura = 1
    
    While Not EOF(Arq)
        If Linha.Tipo = "A" Then
            Linha.Leitura = TrataLeitura(Linha.Leitura)
            TipoDoc = 0
            Valor = "000"
            ' gravar documento
            qryInsereDocto.rdoParameters(0).Direction = rdParamReturnValue
            qryInsereDocto.rdoParameters(10).Direction = rdParamOutput
            qryInsereDocto.rdoParameters(1) = Geral.DataProcessamento
            qryInsereDocto.rdoParameters(2) = IdCapa
            qryInsereDocto.rdoParameters(3) = TipoDoc
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
            qryInsereDocto.rdoParameters(9) = OrdemCaptura
            qryInsereDocto.Execute
            If qryInsereDocto.rdoParameters(0) <> 0 Then
                GoTo ErroCaptura
            End If
            GravaLog IdCapa, qryInsereDocto.rdoParameters(10), 41
            Count = Count + 1
            OrdemCaptura = OrdemCaptura + 1
        
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
    AtualizaStatusLote
    LoteCapturado = True
    ProcessaArquivoRetornoCanonLS500 = True
        
    lblNumLote.Caption = Format(IdLote, "0000-00000")
    lblQtdeCapa.Caption = Format(CountCapas, "0000")
    lblQtdeDocto.Caption = Format(Count - CountCapas, "0000")
    Exit Function
    
ErroCopiaImagem:
    If MsgBox("Erro ao gravar a imagem do documento. Deseja tentar novamente?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        Resume
        Exit Function
    End If
    
ErroCaptura:
    Geral.Banco.RollbackTrans
    
    LoteCapturado = False
    
    TratamentoErro "Erro no processamento do arquivo de retorno.", Err, rdoErrors, False
    MsgBox "Erro no processamento do arquivo de retorno.", vbCritical + vbOKOnly, App.Title
    ProcessaArquivoRetornoCanonLS500 = False
End Function

Private Function VerificaSeCapa(ByVal Leitura As String, ByRef IdEnv_Mal As String) As Boolean
    Dim Campo1 As String
    Dim Campo2 As String
    Dim Campo3 As String
    Dim Valor As String
    
    VerificaSeCapa = False
    IdEnv_Mal = "E"
    Leitura = Trim(Leitura)
    
    If Leitura = "" Then
        Exit Function
    End If
    If Len(Leitura) = 8 Then ' Envelope
        If IsNumeric(Leitura) Then
            If Right(Leitura, 1) = Modulo11UBB(Val(Left(Leitura, Len(Leitura) - 1))) Or _
               Right(Leitura, 1) = Modulo11Simplificado(Val(Left(Leitura, Len(Leitura) - 1))) Or _
               Right(Leitura, 1) = Modulo11U(Val(Left(Leitura, Len(Leitura) - 1))) Then
                VerificaSeCapa = True
                IdEnv_Mal = "E"
            End If
        End If
    Else ' Malote
        If Len(Leitura) >= 30 Then
            If TratarCamposCMC7(Leitura, Campo1, Campo2, Campo3, Valor) Then
                'If Mid(Campo3, 1, 4) = "0600" And Mid(Campo1, 1, 3) = "409" And Mid(Campo2, 10, 1) = "4" Then
                If Mid(Campo2, 1, 3) = "600" And Mid(Campo1, 1, 3) = "409" And Mid(Campo2, 10, 1) = "4" Then
                    VerificaSeCapa = True
                    IdEnv_Mal = "M"
                End If
            End If
        End If
    End If
End Function

Private Function TwipsXToPixel(ByVal Twips As Long) As Long
    TwipsXToPixel = Int(Twips / Screen.TwipsPerPixelX)
End Function

Private Function TwipsYToPixel(ByVal Twips As Long) As Long
    TwipsYToPixel = Int(Twips / Screen.TwipsPerPixelY)
End Function

Private Function InsereLote() As Boolean
    On Error GoTo ErroLote
    rdoErrors.Clear
    
    Screen.MousePointer = vbDefault
    
    With qryInsereLote
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = Prioridade
        .rdoParameters(3) = CLng(Geral.AgenciaCentral)
        .rdoParameters(4).Direction = rdParamOutput
        .Execute
        If .rdoParameters(0) <> 0 Then
            MsgBox "Erro na gravação do novo lote.", vbCritical + vbOKOnly, App.Title
            InsereLote = False
            Exit Function
        End If
        IdLote = .rdoParameters(4)
    End With
    InsereLote = True
    On Error GoTo 0
    Exit Function

ErroLote:
    Select Case TratamentoErro("Erro na gravação do novo lote.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    Unload Me

End Function

Private Function AtualizaStatusLote() As Boolean
    Dim Status As String
    
    AtualizaStatusLote = True
    Screen.MousePointer = vbHourglass
    
    On Error GoTo ErroContQualidade
    rdoErrors.Clear

    qryGetControleQualidade.rdoParameters(0) = Geral.DataProcessamento

    Set rsContQualidade = qryGetControleQualidade.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    If rsContQualidade.EOF Then
        Status = "2"
    Else
        If rsContQualidade!ControleQualidade = 0 Then
            Status = "2"
        Else
            Status = "0"
        End If
    End If
    rsContQualidade.Close
    
    On Error GoTo ErroAtualizaStatus
    rdoErrors.Clear
    
    
    With qryAtualizaStatusLote
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = IdLote
        .rdoParameters(3) = Status
        .Execute
        If .rdoParameters(0) <> 0 Then
            GoTo ErroAtualizaStatus
        End If
    End With
    Screen.MousePointer = vbDefault
    On Error GoTo 0
    Exit Function
    
ErroContQualidade:
    AtualizaStatusLote = False
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro na obtenção do parametro Controle de Qualidade.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    IdLote = 0
    Exit Function
    
ErroAtualizaStatus:
    AtualizaStatusLote = False
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro na atualização do status do Lote.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    IdLote = 0
    
End Function

Private Function TrataLeitura(ByVal Leitura As String) As String
    Dim bInvalido As Boolean
    Dim Count As Integer
    Dim Result As String
    Dim Char As String * 1
    
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
