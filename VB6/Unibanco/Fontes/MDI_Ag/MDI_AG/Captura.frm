VERSION 5.00
Begin VB.Form frmCaptura 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Captura de Imagens"
   ClientHeight    =   2520
   ClientLeft      =   36
   ClientTop       =   336
   ClientWidth     =   6192
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   6192
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFechar 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   372
      Left            =   3924
      TabIndex        =   2
      Top             =   2010
      Width           =   1536
   End
   Begin VB.CommandButton cmdSemPrioridade 
      Caption         =   "&Sem Prioridade"
      Default         =   -1  'True
      Height          =   372
      Left            =   732
      TabIndex        =   0
      Top             =   2010
      Width           =   1536
   End
   Begin VB.CommandButton cmdComPrioridade 
      Caption         =   "&Com Prioridade"
      Height          =   372
      Left            =   2328
      TabIndex        =   1
      Top             =   2010
      Width           =   1536
   End
   Begin VB.PictureBox Picture5 
      Height          =   312
      Left            =   72
      ScaleHeight     =   264
      ScaleWidth      =   6012
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1590
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
         TabIndex        =   14
         Top             =   12
         Width           =   5952
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1512
      Left            =   60
      TabIndex        =   3
      Top             =   15
      Width           =   6072
      Begin VB.PictureBox Picture4 
         Height          =   372
         Left            =   3396
         ScaleHeight     =   324
         ScaleWidth      =   1104
         TabIndex        =   9
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
            TabIndex        =   12
            Top             =   48
            Width           =   1056
         End
      End
      Begin VB.PictureBox Picture3 
         Height          =   372
         Left            =   3396
         ScaleHeight     =   324
         ScaleWidth      =   1104
         TabIndex        =   8
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
            TabIndex        =   11
            Top             =   48
            Width           =   1056
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   372
         Left            =   3396
         ScaleHeight     =   324
         ScaleWidth      =   1104
         TabIndex        =   7
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
            TabIndex        =   10
            Top             =   48
            Width           =   1032
         End
      End
      Begin VB.Label Label3 
         Caption         =   "Quantidade de Documentos no Lote:"
         Height          =   192
         Left            =   120
         TabIndex        =   6
         Top             =   1128
         Width           =   2688
      End
      Begin VB.Label lblTitQtdeCapa 
         Caption         =   "Quantidade de Envelopes / Malote no Lote:"
         Height          =   192
         Left            =   120
         TabIndex        =   5
         Top             =   708
         Width           =   3192
      End
      Begin VB.Label Label1 
         Caption         =   "Número do Lote Capturado:"
         Height          =   192
         Left            =   120
         TabIndex        =   4
         Top             =   282
         Width           =   2112
      End
   End
End
Attribute VB_Name = "frmCaptura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public m_Connection          As rdo.rdoConnection
Public m_Scanner             As Scanner
Public m_TipoScanner         As enumScanner
Public m_AgenciaApresentante As String
Public m_DataProcessamento   As Long
Public m_DiretorioDados      As String
Public m_DiretorioImagens    As String
Public m_DiretorioTrabalho   As String
Public m_Usuario             As String

Private RetornoFinal As String
Private IdLote As Long
Private LoteCapturado As Boolean
Private FileLog As Integer
Private Prioridade As Integer
Private qryInsereLote As rdoQuery
Private qryRemoveLote As rdoQuery
Private qryInsereCapa As rdoQuery
Private qryInsereDocto As rdoQuery
Private qryAtualizaStatusLote As rdoQuery

Private Function Digitalizar(Optional ByVal pvbAppend As Boolean = False) As Boolean
    Dim iRet As Long
    
    On Error GoTo ErroGetImagem
    rdoErrors.Clear
    
    Digitalizar = True
    
    On Error GoTo ErroDigitalizar
    
    lblMsg.Caption = "Scanner Capturando"
    Me.Refresh
    If m_TipoScanner = escnVIPS Then
        iRet = m_Scanner.Captura(Val(m_AgenciaApresentante), IdLote, m_DiretorioDados & RetornoFinal, pvbAppend)
        Print #FileLog, "Usuario: " & m_Usuario & " Horario: " & Format(Now, "hh:mm:ss") & " - VIPS_Captura: " & Trim(str(iRet))
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
                m_Scanner.Reset
                Digitalizar = False
            Case Else
                MsgBox "Não foi possível digitalizar todo o lote. Codigo de erro: " & iRet, vbExclamation + vbOKOnly, App.Title
                Digitalizar = False
        End Select
        
    End If
    If Digitalizar Then
        lblMsg.Caption = "Scanner Inativo"
    Else
        lblMsg.Caption = "Falha da captura das imagens"
        Print #FileLog, "Usuario: " & m_Usuario & " Horario: " & Format(Now, "hh:mm:ss") & " - Erro na captura das imagens dos documentos"
    End If
    Me.Refresh
    Exit Function
    
ErroGetImagem:
    Select Case TratamentoErro(m_Connection, "Erro na obtenção do número da imagem.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    Exit Function

ErroDigitalizar:
    TratamentoErro m_Connection, "Erro na digitalização da imagens.", Err, rdoErrors, False

End Function

Private Sub cmdComPrioridade_Click()
    Prioridade = 1
    IniciarDigitalizacao
End Sub

Private Sub cmdFechar_Click()
    Unload Me
End Sub

Private Sub cmdSemPrioridade_Click()
    Prioridade = 0
    IniciarDigitalizacao
End Sub

Private Sub Form_Load()
    
    FileLog = FreeFile
    Open m_DiretorioTrabalho & "DIG" & Format(m_DataProcessamento, "00000000") & ".TXT" For Append As #FileLog
    
    Set qryInsereLote = m_Connection.CreateQuery("", "{? = call MDIAG_InsereLote (?,?,?,?)}")
    Set qryRemoveLote = m_Connection.CreateQuery("", "{? = call MDIAG_RemoveLote (?,?)}")
    Set qryInsereCapa = m_Connection.CreateQuery("", "{? = call MDIAG_CapturaCapa (?,?,?,?,?,?)}")
    Set qryInsereDocto = m_Connection.CreateQuery("", "{? = call MDIAG_CapturaDocumento (?,?,?,?,?,?,?,?,?)}")
    Set qryAtualizaStatusLote = m_Connection.CreateQuery("", "{ ? = call MDIAG_AtualizaStatusLote (?,?,?)}")

    lblNumLote.Caption = ""
    lblQtdeCapa.Caption = ""
    lblQtdeDocto.Caption = ""
    lblMsg.Caption = "Scanner Inativo"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Close #FileLog
    qryInsereLote.Close
    qryRemoveLote.Close
    qryInsereCapa.Close
    qryInsereDocto.Close
    qryAtualizaStatusLote.Close
End Sub

Private Sub IniciarDigitalizacao()
    Dim dtInicio As Date
    Dim dtFim As Date
    Dim bDigitalizou As Boolean
    Dim bAppend As Boolean
    Dim Frente As String, Verso As String
    Dim Opcao As Integer
    Dim Count As Integer

    cmdSemPrioridade.Enabled = False
    cmdComPrioridade.Enabled = False
    cmdFechar.Enabled = False
    
    Screen.MousePointer = vbHourglass
    
    On Error GoTo ErroInsereLote
    
    IdLote = 0
    LoteCapturado = False
    
    m_Connection.BeginTrans
    
    If Not InsereLote Then
        GoTo ErroInsereLote
    End If
    
    On Error GoTo ErroDigitalizacao
    '''''''''''''''''''''''''
    ' Digitalizar Documento '
    '''''''''''''''''''''''''
    bAppend = False
    RetornoFinal = Format(m_DataProcessamento, "00000000") & Format(IdLote, "000000000") & ".txt"

    dtInicio = Now  ' Inicio da digitalizacao
    Print #FileLog, "Usuario: " & m_Usuario & " Horario: " & Format(Now, "hh:mm:ss") & " - Inicio da captura"

Redigitalizar:
    
    bDigitalizou = Digitalizar(bAppend)
    
    dtFim = Now     ' Fim da digitalizacao
    Print #FileLog, "Usuario: " & m_Usuario & " Horario: " & Format(Now, "hh:mm:ss") & " - Fim da captura"
    
    If Not bDigitalizou Then
        Screen.MousePointer = vbDefault
        If MsgBox("Não foi possível concluir a digitalização. Deseja continuar no mesmo lote?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            bAppend = True
            GoTo Redigitalizar
        Else
            bAppend = False
        End If
    End If
    ''''''''''''''''''''''''''''''''''''''
    ' Processar retorno da digitalizacao '
    ''''''''''''''''''''''''''''''''''''''
ReprocessarArquivo:
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Varre o arquivo de retorno e obtem o nome da ultima imagem'
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Not ObtemUltimaImagem(m_DiretorioDados & RetornoFinal, Frente, Verso) Then
        Screen.MousePointer = vbDefault
        MsgBox "Não foi possível obter a última imagem capturada. Repasse os documentos.", vbExclamation + vbOKOnly, App.Title
        Print #FileLog, "Usuario: " & m_Usuario & " Horario: " & Format(Now, "hh:mm:ss") & " - Nao foi localizada a imagem do ultimo documento"
        GoTo FinalDigitalizacao
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Apresenta img do ult. docto e espera confirmação fechamento do lote'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If m_TipoScanner = escnVIPS Then
        Load ConfFimLote
        On Error Resume Next
        'verso
        With ConfFimLote.Lead2
           .AutoRepaint = False
           .Load m_DiretorioImagens & Format(IdLote, "000000000") & "\" & Verso, 0, 0, 1
           .Intensity 220
           .PaintZoomFactor = 100
           .AutoRepaint = True
        End With
        'frente
        With ConfFimLote.Lead1
           .AutoRepaint = False
           .Load m_DiretorioImagens & Format(IdLote, "000000000") & "\" & Frente, 0, 0, 1
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

        Select Case Opcao
            Case 0 ' Cancelou o Lote
                Print #FileLog, "Usuario: " & m_Usuario & " Horario: " & Format(Now, "hh:mm:ss") & " - Cancelamento da captura do lote"
                MsgBox "Redigitalize este lote, pois o mesmo não foi gravado.", vbExclamation + vbOKOnly, App.Title
                LoteCapturado = False
                GoTo FinalDigitalizacao
            Case 1 ' Confirmou o Lote
                ' continua o processamento
                Print #FileLog, "Usuario: " & m_Usuario & " Horario: " & Format(Now, "hh:mm:ss") & " - Confirmacao da captura do lote"
            Case 2 ' Continuar capturando no mesmo lote
                Print #FileLog, "Usuario: " & m_Usuario & " Horario: " & Format(Now, "hh:mm:ss") & " - Continuacao da captura no mesmo lote"
                bAppend = True
                GoTo Redigitalizar
        End Select
    End If
    
    Screen.MousePointer = vbHourglass
    
    '''''''''''''''''''''''''''''''''''''''''''''''''
    ' Varre o arquivo e grava os documentos
    '''''''''''''''''''''''''''''''''''''''''''''''''
    Print #FileLog, "Usuario: " & m_Usuario & " Horario: " & Format(Now, "hh:mm:ss") & " - Inicio do processamento do arquivo: " & RetornoFinal
    
    While Not ProcessaArquivoRetorno(m_DiretorioDados & RetornoFinal, Count)
        Screen.MousePointer = vbDefault
        ' erro na gravacao dos doctos do lote
        If MsgBox("Houve um erro na gravação do lote." & vbCrLf & _
               "Deseja tentar gravar novamente?", vbQuestion + vbYesNo, App.Title) = vbNo Then
        
            MsgBox "Redigitalize este lote, pois o mesmo não foi gravado.", vbExclamation + vbOKOnly, App.Title
            Print #FileLog, "Usuario: " & m_Usuario & " Horario: " & Format(Now, "hh:mm:ss") & " - Erro no processamento do arquivo de retorno, documentos nao gravados"
            LoteCapturado = False
            GoTo FinalDigitalizacao
            
        End If
        
    Wend
    Print #FileLog, "Usuario: " & m_Usuario & " Horario: " & Format(Now, "hh:mm:ss") & " - Termino do processamento do arquivo de retorno"
    
    '''''''''''''''''''''''''''
    ' Finalizar digitalização '
    '''''''''''''''''''''''''''
    GoTo FinalDigitalizacao

ErroInsereLote:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro(m_Connection, "Não foi possível inserir novo Lote.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    GoTo FinalDigitalizacao
    Exit Sub

ErroDigitalizacao:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro(m_Connection, "Erro inesperado na Captura de imagens.", Err, rdoErrors)
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
        m_Connection.RollbackTrans
'        If IdLote > 0 Then
'            With qryRemoveLote
'                .rdoParameters(1) = m_DataProcessamento
'                .rdoParameters(2) = IdLote
'                .Execute
'            End With
'        End If
    Else
        m_Connection.CommitTrans
    End If
End Sub

Private Function ObtemUltimaImagem(ByVal NomeArq As String, _
                                    ByRef Frente As String, ByRef Verso As String) As Boolean
    Dim Arq As Integer
    Dim RetornoUnibanco As tpRetornoVips
    
    Arq = FreeFile
    Open NomeArq For Binary As #Arq
    
    RetornoUnibanco.Frente = String(19, "  ")
    RetornoUnibanco.Verso = String(19, "  ")
    
    Get #Arq, , RetornoUnibanco
    While Not EOF(Arq)
        Frente = RetornoUnibanco.Frente
        Verso = IIf(Trim(RetornoUnibanco.Verso) = "", RetornoUnibanco.Frente, RetornoUnibanco.Verso)
        Get #Arq, , RetornoUnibanco
    Wend
    
    Close #Arq
    
    If Trim(Frente) <> "" Then
        ObtemUltimaImagem = True
    Else
        ObtemUltimaImagem = False
    End If
    
End Function

Private Function ProcessaArquivoRetorno(ByVal NomeArq As String, ByRef Count As Integer) As Boolean
    Dim Arq             As Integer
    Dim IdCapa          As Integer
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
    
'    m_Connection.BeginTrans
    
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
            qryInsereCapa.rdoParameters(1) = m_DataProcessamento
            qryInsereCapa.rdoParameters(2) = IdLote
            qryInsereCapa.rdoParameters(3) = IdEnv_Mal
            qryInsereCapa.rdoParameters(4) = IIf(IdEnv_Mal = "E", Val(Trim(Linha.Leitura)), Val(Mid(Campo3, 1, 4) & Mid(Campo2, 4, 6) & Mid(Campo1, 4, 4)))
            qryInsereCapa.rdoParameters(5) = m_AgenciaApresentante
            qryInsereCapa.rdoParameters(6).Direction = rdParamOutput
            qryInsereCapa.Execute
            If qryInsereCapa.rdoParameters(0) <> 0 Then
                GoTo ErroCaptura
            End If
            IdCapa = qryInsereCapa.rdoParameters(6)
            GravaLog m_Connection, m_DataProcessamento, IdCapa, 0, m_Usuario, 40
            CountCapas = CountCapas + 1
            OrdemCaptura = 1
        ElseIf IdCapa = 0 Then
            TipoDoc = 1 ' Capa
            bVirtual = True
            ' gravar capa
            qryInsereCapa.rdoParameters(0).Direction = rdParamReturnValue
            qryInsereCapa.rdoParameters(1) = m_DataProcessamento
            qryInsereCapa.rdoParameters(2) = IdLote
            qryInsereCapa.rdoParameters(3) = "E" ' Capa virtual sempre sera envelope
            qryInsereCapa.rdoParameters(4) = 9 ' Capa Virtual
            qryInsereCapa.rdoParameters(5) = m_AgenciaApresentante
            qryInsereCapa.rdoParameters(6).Direction = rdParamOutput
            qryInsereCapa.Execute
            If qryInsereCapa.rdoParameters(0) <> 0 Then
                GoTo ErroCaptura
            End If
            IdCapa = qryInsereCapa.rdoParameters(6)
            GravaLog m_Connection, m_DataProcessamento, IdCapa, 0, m_Usuario, 40
            CountCapas = CountCapas + 1
            OrdemCaptura = 1
        End If
        ' gravar documento
        Valor = "000"
        qryInsereDocto.rdoParameters(0).Direction = rdParamReturnValue
        qryInsereDocto.rdoParameters(9).Direction = rdParamOutput
        qryInsereDocto.rdoParameters(1) = m_DataProcessamento
        qryInsereDocto.rdoParameters(2) = IdCapa
        qryInsereDocto.rdoParameters(3) = TipoDoc
        If IdEnv_Mal = "M" And TipoDoc = 1 Then ' Capa de Malote
            If Not bVirtual Then
                qryInsereDocto.rdoParameters(4) = Mid(Campo3, 1, 4) & Mid(Campo2, 4, 6) & Mid(Campo1, 4, 4)
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
        qryInsereDocto.rdoParameters(5) = Linha.Frente
        qryInsereDocto.rdoParameters(6) = IIf(Trim(Linha.Verso) = "", Linha.Frente, Linha.Verso)
        qryInsereDocto.rdoParameters(7) = Linha.origem
        qryInsereDocto.rdoParameters(8) = OrdemCaptura
        qryInsereDocto.Execute
        If qryInsereDocto.rdoParameters(0) <> 0 Then
            GoTo ErroCaptura
        End If
        GravaLog m_Connection, m_DataProcessamento, IdCapa, qryInsereDocto.rdoParameters(9), m_Usuario, 41
        Count = Count + 1
        OrdemCaptura = OrdemCaptura + 1
        Get #Arq, , Linha
    Wend
    Close #Arq
'    m_Connection.CommitTrans
    AtualizaStatusLote
    LoteCapturado = True
    ProcessaArquivoRetorno = True

    lblNumLote.Caption = Format(IdLote, "0000-00000")
    lblQtdeCapa.Caption = Format(CountCapas, "0000")
    lblQtdeDocto.Caption = Format(Count - CountCapas, "0000")
    Exit Function

ErroCaptura:
'    If IdCapa >= 0 Then
'        m_Connection.RollbackTrans
'    End If

    LoteCapturado = False

    TratamentoErro m_Connection, "Erro no processamento do arquivo de retorno.", Err, rdoErrors, False
    MsgBox "Erro no processamento do arquivo de retorno.", vbCritical + vbOKOnly, App.Title
    ProcessaArquivoRetorno = False

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
                If Mid(Campo3, 1, 4) = "0600" And Mid(Campo1, 1, 3) = "409" And Mid(Campo2, 10, 1) = "4" Then
                    VerificaSeCapa = True
                    IdEnv_Mal = "M"
                End If
            End If
        End If
    End If
End Function

Private Function InsereLote() As Boolean
    On Error GoTo ErroLote
    rdoErrors.Clear
    
    Screen.MousePointer = vbDefault
    
    With qryInsereLote
        .rdoParameters(1) = m_DataProcessamento
        .rdoParameters(2) = Prioridade
        .rdoParameters(3) = CLng(m_AgenciaApresentante)
        .rdoParameters(4).Direction = rdParamOutput
        .Execute
        If .rdoParameters(0) <> 0 Then
            MsgBox "Erro a gravação do novo lote.", vbCritical + vbOKOnly, App.Title
            InsereLote = False
            Exit Function
        End If
        IdLote = .rdoParameters(4)
    End With
    InsereLote = True
    On Error GoTo 0
    Exit Function

ErroLote:
    Select Case TratamentoErro(m_Connection, "Erro na gravação do novo lote.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    Unload Me

End Function

Private Function AtualizaStatusLote() As Boolean
    
    AtualizaStatusLote = True
    Screen.MousePointer = vbHourglass
    
    On Error GoTo ErroAtualizaStatus
    rdoErrors.Clear
    
    With qryAtualizaStatusLote
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1) = m_DataProcessamento
        .rdoParameters(2) = IdLote
        .rdoParameters(3) = "0" 'Digitalizado
        .Execute
        If .rdoParameters(0) <> 0 Then
            GoTo ErroAtualizaStatus
        End If
    End With
    Screen.MousePointer = vbDefault
    On Error GoTo 0
    Exit Function
    
ErroAtualizaStatus:
    AtualizaStatusLote = False
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro(m_Connection, "Erro na atualização do status do Lote.", Err, rdoErrors)
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
