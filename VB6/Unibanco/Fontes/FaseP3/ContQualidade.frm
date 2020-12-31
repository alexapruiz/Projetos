VERSION 5.00
Object = "{00080003-B1BA-11CE-ABC6-F5B2E79D9E3F}#8.0#0"; "LTOCX80N.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form ControleQualidade 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controle de Qualidade"
   ClientHeight    =   8784
   ClientLeft      =   36
   ClientTop       =   336
   ClientWidth     =   11940
   ForeColor       =   &H8000000F&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8784
   ScaleWidth      =   11940
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1740
      Left            =   48
      TabIndex        =   14
      Top             =   6984
      Width           =   1752
      Begin VB.CommandButton cmdFechar 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   324
         Left            =   132
         TabIndex        =   18
         Top             =   1272
         Width           =   1464
      End
      Begin VB.CommandButton cmdAtualizar 
         Caption         =   "&Atualizar"
         Height          =   324
         Left            =   132
         TabIndex        =   17
         Top             =   204
         Width           =   1464
      End
      Begin VB.CommandButton cmdLiberar 
         Caption         =   "&Liberar Lote"
         Height          =   324
         Left            =   132
         TabIndex        =   16
         Top             =   564
         Width           =   1464
      End
      Begin VB.CommandButton cmdRemover 
         Caption         =   "&Remover Docto."
         Height          =   324
         Left            =   144
         TabIndex        =   15
         Top             =   912
         Width           =   1464
      End
   End
   Begin VB.ListBox lstLote 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1008
      Left            =   48
      TabIndex        =   13
      Top             =   432
      Width           =   1752
   End
   Begin VB.ListBox lstDocto 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4368
      ItemData        =   "ContQualidade.frx":0000
      Left            =   48
      List            =   "ContQualidade.frx":0002
      MultiSelect     =   2  'Extended
      TabIndex        =   12
      Top             =   2184
      Width           =   1752
   End
   Begin VB.PictureBox Picture1 
      Height          =   264
      Left            =   48
      ScaleHeight     =   216
      ScaleWidth      =   1704
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   96
      Width           =   1752
      Begin VB.Label lblLote 
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
         Left            =   12
         TabIndex        =   11
         Top             =   -24
         Width           =   1272
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   264
      Left            =   48
      ScaleHeight     =   216
      ScaleWidth      =   1704
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1848
      Width           =   1752
      Begin VB.Label Label5 
         Caption         =   "Descrição"
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
         TabIndex        =   9
         Top             =   -12
         Width           =   1104
      End
   End
   Begin VB.PictureBox FrmPesquisa 
      Height          =   1932
      Left            =   2724
      ScaleHeight     =   1884
      ScaleWidth      =   6420
      TabIndex        =   4
      Top             =   3444
      Visible         =   0   'False
      Width           =   6468
      Begin VB.CommandButton CmdFecharPesquisa 
         Caption         =   "&Fechar"
         Height          =   312
         Left            =   2724
         TabIndex        =   5
         Top             =   1428
         Width           =   1068
      End
      Begin ComctlLib.ProgressBar Progress 
         Height          =   300
         Left            =   348
         TabIndex        =   6
         Top             =   912
         Width           =   5808
         _ExtentX        =   10245
         _ExtentY        =   529
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Pesquisando por Lotes em Controle de Qualidade. Aguarde ..."
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
         Left            =   300
         TabIndex        =   7
         Top             =   576
         Width           =   5940
      End
   End
   Begin VB.Timer tmrPesquisa 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   972
      Top             =   5388
   End
   Begin VB.Frame frmImagem2 
      Caption         =   "Imagem Verso"
      Height          =   4392
      Left            =   1908
      TabIndex        =   2
      Top             =   4356
      Width           =   9972
      Begin LeadLib.Lead Lead2 
         Height          =   4128
         Left            =   72
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   192
         Width           =   9828
         _Version        =   524288
         _ExtentX        =   17336
         _ExtentY        =   7281
         _StockProps     =   229
         BackColor       =   -2147483639
         BorderStyle     =   1
         ScaleHeight     =   342
         ScaleWidth      =   817
         DataField       =   ""
         BitmapDataPath  =   ""
         AnnDataPath     =   ""
      End
   End
   Begin VB.Timer tmrAtualiza 
      Enabled         =   0   'False
      Interval        =   50000
      Left            =   576
      Top             =   5376
   End
   Begin VB.Frame frmImagem1 
      Caption         =   "Imagem Frente"
      Height          =   4392
      Left            =   1896
      TabIndex        =   0
      Top             =   -24
      Width           =   9972
      Begin LeadLib.Lead Lead1 
         Height          =   4116
         Left            =   48
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   192
         Width           =   9840
         _Version        =   524288
         _ExtentX        =   17357
         _ExtentY        =   7260
         _StockProps     =   229
         BackColor       =   -2147483639
         BorderStyle     =   1
         ScaleHeight     =   341
         ScaleWidth      =   818
         DataField       =   ""
         BitmapDataPath  =   ""
         AnnDataPath     =   ""
      End
   End
End
Attribute VB_Name = "ControleQualidade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type tpMyDoc
    IdDocto     As Long
    IdCapa      As Long
    TipoDocto   As Integer
    Frente      As String
    Verso       As String
    Ordem       As String * 1
    Leitura     As String
    Deleted     As Boolean
End Type

Private m_Busy                          As Boolean
Private m_IdLote                        As Long
Private m_Doc                           As tpMyDoc
Private aDoc()                          As tpMyDoc
Private m_CountDocto                    As Integer
Private m_CountDeleted                  As Integer
Private sTempo                          As Integer
Private m_FirstActivate                 As Boolean

Private qryAtualizaStatusLote           As rdoQuery
Private qryVerificaLoteDisponivel       As rdoQuery
Private qryRemoveDocumento              As rdoQuery
Private qryOrdenaCaptura                As rdoQuery
Private qryRemoveCapa                   As rdoQuery
Private qryRemoveLote                   As rdoQuery
Private qryGetLoteContQualidade         As rdoQuery
Private qryGetDocumentoContQualidade    As rdoQuery
Private qryGetCapasLote                  As rdoQuery
Private qryGravaLog                     As rdoQuery
Private rsLote                          As rdoResultset
Private rsDoc                           As rdoResultset
Private rsCapa                          As rdoResultset



Private Sub LimparListas()
    lstLote.Clear
    lstDocto.Clear
End Sub

Private Sub Preenche_lstDocto()
    Dim Linha As String
    Dim Count As Integer
    
    frmImagem1.Visible = False
    frmImagem2.Visible = False
    lstDocto.Clear
    For Count = 1 To m_CountDocto
        If Not aDoc(Count).Deleted Then
            If aDoc(Count).TipoDocto = "1" Then
                If Left(aDoc(Count).Leitura, 4) = "0600" Then
                    Linha = "MALOTE   " & Space(18)
                Else
                    Linha = "ENVELOPE " & Space(18)
                End If
            Else
                Linha = "DOCUMENTO" & Space(18)
            End If
            Linha = Linha & Format(aDoc(Count).IdDocto, "0000000000")
            lstDocto.AddItem Linha
        End If
    Next
End Sub

Private Function ObtemLote() As Boolean
    On Error GoTo ErroGetLote
    rdoErrors.Clear
    
    Screen.MousePointer = vbHourglass
    lstLote.Clear
    
    qryGetLoteContQualidade.rdoParameters(0) = Geral.DataProcessamento
    qryGetLoteContQualidade.rdoParameters(1) = Geral.Intervalo
    Set rsLote = qryGetLoteContQualidade.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    If rsLote.EOF Then
        rsLote.Close
        ObtemLote = False
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    While Not rsLote.EOF
        lstLote.AddItem Format(rsLote!IdLote, "000000000")
        rsLote.MoveNext
    Wend
    rsLote.Close
    ObtemLote = True
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    Exit Function
    
ErroGetLote:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro na obtenção de envelope/malote para Controle de Qualidade.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    m_IdLote = 0
    m_Busy = False
    Unload Me

End Function

Private Sub ObtemDocumentos(ByVal IdLote As Long)
    On Error GoTo ErroGetDocto
    rdoErrors.Clear
    
    Screen.MousePointer = vbHourglass
    
    Erase aDoc
    m_CountDocto = 0
    qryGetDocumentoContQualidade.rdoParameters(0) = Geral.DataProcessamento
    qryGetDocumentoContQualidade.rdoParameters(1) = IdLote
    Set rsDoc = qryGetDocumentoContQualidade.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    ReDim aDoc(rsDoc.RowCount)
    While Not rsDoc.EOF
        m_CountDocto = m_CountDocto + 1
        m_Doc.IdDocto = rsDoc!IdDocto
        m_Doc.IdCapa = rsDoc!IdCapa
        m_Doc.TipoDocto = rsDoc!TipoDocto
        m_Doc.Frente = rsDoc!Frente
        m_Doc.Verso = rsDoc!Verso
        m_Doc.Leitura = rsDoc!Leitura
        m_Doc.Ordem = rsDoc!Ordem
        m_Doc.Deleted = False
        aDoc(m_CountDocto) = m_Doc
        rsDoc.MoveNext
    Wend
    rsDoc.Close
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErroGetDocto:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro na obtenção de documentos para Controle de Qualidade.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    m_IdLote = 0
    m_Busy = False
    Unload Me
    
End Sub

Private Function VerificaLoteDisponivel(ByVal IdLote As Long) As Boolean
    On Error GoTo ErroVerificaLote
    rdoErrors.Clear
    
    Screen.MousePointer = vbHourglass
    
    With qryVerificaLoteDisponivel
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = IdLote
        .rdoParameters(3) = Geral.Intervalo
        .Execute
        Screen.MousePointer = vbDefault
        If .rdoParameters(0) = 0 Then
            VerificaLoteDisponivel = True
        ElseIf .rdoParameters(0) = 1 Then
            VerificaLoteDisponivel = False
            MsgBox "Este Lote não está mais disponível por já ter sido tratado ou porque esta sendo tratado por outra estação.", vbInformation + vbOKOnly, App.Title
        Else
            VerificaLoteDisponivel = False
            MsgBox "Erro. Não foi possível obter o Status do Lote.", vbInformation + vbOKOnly, App.Title
        End If
    End With
    On Error GoTo 0
    Exit Function
    
ErroVerificaLote:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro. Não foi possível obter o Status do Lote.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    m_IdLote = 0
    m_Busy = False
    Unload Me
    
End Function

Private Function RemoveDocumento(ByVal IdCapa As Long, ByVal IdDocto As Long) As Boolean
    On Error GoTo ErroRemoveDoc
    rdoErrors.Clear

    Screen.MousePointer = vbHourglass

    With qryRemoveDocumento
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = IdDocto
        .Execute
        Screen.MousePointer = vbDefault
        If .rdoParameters(0) = 0 Then
            m_CountDeleted = m_CountDeleted + 1
            RemoveDocumento = True
        Else
            RemoveDocumento = False
        End If
    End With

    'Gravar Log
    Call GravaLog(IdCapa, IdDocto, 31)

    On Error GoTo 0
    Exit Function
    
ErroRemoveDoc:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro. Não foi possível remover o documento.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    m_IdLote = 0
    m_Busy = False
    Unload Me
    
End Function
Private Function OrdenaCaptura(ByVal IdCapa As Long) As Boolean
    
    On Error GoTo Err_OrdenaCapa
    rdoErrors.Clear
    OrdenaCaptura = False
    
    Screen.MousePointer = vbHourglass

    With qryOrdenaCaptura
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = IdCapa
        .Execute
        
        Screen.MousePointer = vbDefault
        If .rdoParameters(0) = 0 Then
            OrdenaCaptura = True
        End If
    End With

    On Error GoTo 0
    Exit Function
    
Err_OrdenaCapa:
    
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro. Não foi possível remover o documento.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    m_IdLote = 0
    m_Busy = False
    Unload Me

End Function
Private Function RemoveCapa(ByVal IdCapa As Long) As Boolean
    Dim Count As Integer
    
    For Count = 1 To m_CountDocto
        If aDoc(Count).Deleted = False And aDoc(Count).IdCapa = IdCapa Then
            If Not RemoveDocumento(aDoc(Count).IdCapa, aDoc(Count).IdDocto) Then
                RemoveCapa = False
                Exit Function
            End If
            aDoc(Count).Deleted = True
        End If
    Next
    
    On Error GoTo ErroRemoveCapa
    rdoErrors.Clear
    
    Screen.MousePointer = vbHourglass
    
    With qryRemoveCapa
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = IdCapa
        .Execute
        Screen.MousePointer = vbDefault
        If .rdoParameters(0) = 0 Then
            RemoveCapa = True
        Else
            RemoveCapa = False
            MsgBox "Erro. Não foi possível remover o Envelope/Malote.", vbCritical + vbOKOnly, App.Title
        End If
    End With

    'Gravar Log
    Call GravaLog(IdCapa, 0, 30)

    On Error GoTo 0
    Exit Function
    
ErroRemoveCapa:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro. Não foi possível remover o Envelope/Malote.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    m_IdLote = 0
    m_Busy = False
    Unload Me
End Function

Private Function RemoveLote(ByVal IdLote As Long) As Boolean
    Dim Count As Integer
    
    For Count = 1 To m_CountDocto
        If aDoc(Count).Deleted = False And aDoc(Count).TipoDocto = 1 Then
            If Not RemoveCapa(aDoc(Count).IdCapa) Then
                RemoveLote = False
                Exit Function
            End If
        End If
    Next
    
    On Error GoTo ErroRemoveLote
    rdoErrors.Clear
    
    Screen.MousePointer = vbHourglass
    
    With qryRemoveLote
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = IdLote
        .Execute
        Screen.MousePointer = vbDefault
        If .rdoParameters(0) = 0 Then
            m_IdLote = 0
            RemoveLote = True
        Else
            RemoveLote = False
            MsgBox "Erro. Não foi possível remover o Lote.", vbCritical + vbOKOnly, App.Title
        End If
    End With
    On Error GoTo 0
    Exit Function
    
ErroRemoveLote:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro. Não foi possível remover o Lote.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    m_IdLote = 0
    m_Busy = False
    Unload Me
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
    Dim Ret As Long
    
    i = Indice(Val(Right(lstDocto.List(lstDocto.ListIndex), 10)))
    
    hCtl = Lead1.hwnd
    ''''''''''''''''''''''''''''''''''
    ' mostra imagem frente escolhida '
    ''''''''''''''''''''''''''''''''''
    On Error GoTo ErroImagemFrente
    With Lead1
       .Tag = "F"
       .AutoRepaint = False

       If Geral.VIPSDLL = eDllProservi Then
         .Load Geral.DiretorioImagens & aDoc(i).Frente, 0, 0, 1
       Else
         .Load Geral.DiretorioImagens & lstLote.List(lstLote.ListIndex) & "\" & aDoc(i).Frente, 0, 0, 1
       End If

       ' se imagem for da ls500, deixar mais escura
       If aDoc(i).Ordem <> "2" Then
          .Intensity 220
       Else
          .Intensity 140
       End If
       ' se imagem for do canon, diminui em 50% o tamanho
       If aDoc(i).Ordem <> "1" Then
          .PaintZoomFactor = 80
       Else
          .PaintZoomFactor = 50
       End If
       .AutoRepaint = True
    End With
    
    'posiciona imagem sempre no começo
    Ret = SendMessage(hCtl, WM_VSCROLL, SB_TOP, 0)
    Ret = SendMessage(hCtl, WM_HSCROLL, SB_BOTTOM, 0)
    
    frmImagem1.Visible = True
    
    hCtl = Lead2.hwnd
    ''''''''''''''''''''''''''''''''''
    ' mostra imagem verso escolhida  '
    ''''''''''''''''''''''''''''''''''
    On Error GoTo ErroImagemVerso
    With Lead2
       .Tag = "F"
       .AutoRepaint = False

       If Geral.VIPSDLL = eDllProservi Then
         .Load Geral.DiretorioImagens & aDoc(i).Verso, 0, 0, 1
       Else
         .Load Geral.DiretorioImagens & lstLote.List(lstLote.ListIndex) & "\" & aDoc(i).Verso, 0, 0, 1
       End If

       ' se imagem for da ls500, deixar mais escura
       If aDoc(i).Ordem <> "2" Then
          .Intensity 220
       Else
          .Intensity 140
       End If
       ' se imagem for do canon, diminui em 60% o tamanho
       If aDoc(i).Ordem <> "1" Then
          .PaintZoomFactor = 80
       Else
          .PaintZoomFactor = 50
       End If
       .AutoRepaint = True
    End With
    'posiciona imagem sempre no começo
    Ret = SendMessage(hCtl, WM_VSCROLL, SB_TOP, 0)
    Ret = SendMessage(hCtl, WM_HSCROLL, SB_BOTTOM, 0)
    
    frmImagem2.Visible = True
    
    On Error GoTo 0
    Exit Sub
    
ErroImagemFrente:
    frmImagem1.Visible = False
    MsgBox "Não foi possível exibir imagem do documento, imagem não encontrada", vbExclamation + vbOKOnly, App.Title
ErroImagemVerso:
    frmImagem1.Visible = False

End Sub

Private Function AtualizaStatusLote(ByVal IdLote As Long, ByVal Status As String) As Boolean
    On Error GoTo ErroAtualizaStatus
    rdoErrors.Clear
    
    AtualizaStatusLote = True
    Screen.MousePointer = vbHourglass
    
    With qryAtualizaStatusLote
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = IdLote
        .rdoParameters(3) = Status
        .Execute
        If .rdoParameters(0) <> 0 Then
            AtualizaStatusLote = False
            Screen.MousePointer = vbDefault
            MsgBox "Erro na atualização do status do Lote.", vbCritical + vbOKOnly, App.Title
        End If
    End With
    Screen.MousePointer = vbDefault
    On Error GoTo 0
    Exit Function
    
ErroAtualizaStatus:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro na atualização do status do Lote.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    m_IdLote = 0
    m_Busy = False
    Unload Me
    
End Function

Private Sub HabilitaTimerPesquisa()

  'Esta Função irá verificar a existência de documentos Ilegíveis a cada x segundos
  'de acordo com o campo PARAMETRO.TmAtualizacao
  FrmPesquisa.Visible = True
  tmrPesquisa.Enabled = True
  Progress.Value = 0
End Sub

Private Sub CmdAtualizar_Click()
    If m_IdLote > 0 Then
        If Not AtualizaStatusLote(m_IdLote, "0") Then
            m_IdLote = 0
            m_Busy = False
            Exit Sub
        End If
    End If
    
    LimparListas
    frmImagem1.Visible = False
    frmImagem2.Visible = False
    
    If Not ObtemLote Then
        MsgBox "Não existem Lotes com pendência de Controle de Qualidade.", vbExclamation + vbOKOnly, App.Title
        m_IdLote = 0
        HabilitaTimerPesquisa
        Exit Sub
    Else
        FrmPesquisa.Visible = False
        tmrPesquisa.Enabled = False
    End If

    lstLote.Selected(0) = True
End Sub
Private Sub CmdFechar_Click()
    Unload Me
End Sub

Private Sub CmdFecharPesquisa_Click()
    CmdFechar_Click
End Sub

Private Sub cmdLiberar_Click()
    
    If FrmPesquisa.Visible = True Then
      lstDocto.SetFocus
      Exit Sub
    End If

    'Gravar Log
    Call LiberaLote(m_IdLote)

    'Atualizar status do lote
    AtualizaStatusLote m_IdLote, "2"

    m_IdLote = 0
    If lstLote.ListIndex < lstLote.ListCount - 1 Then
        lstLote.Selected(lstLote.ListIndex + 1) = True
    Else
        If lstLote.ListIndex = lstLote.ListCount - 1 Then
            CmdAtualizar_Click
        Else
            lstLote_Click
        End If
    End If

    lstDocto.SetFocus

End Sub

Private Sub cmdRemover_Click()
    Dim Count As Integer
    Dim i As Integer
    Dim QtdEnv As Integer
    Dim QtdMal As Integer
    Dim Msg As String
    Dim bExiste As Boolean
    Dim ListIndex As Integer
    Dim RemoverLote As Boolean
    
    If FrmPesquisa.Visible = True Then Exit Sub
    
    RemoverLote = False
    If lstDocto.ListCount = lstDocto.SelCount Then
        Msg = "O Lote inteiro será removido. Confirma remover documentos?"
        If MsgBox(Msg, vbQuestion + vbYesNo, App.Title) = vbNo Then
            lstDocto.SetFocus
            Exit Sub
        End If
        If Not RemoveLote(Val(lstLote.List(lstLote.ListIndex))) Then
            lstDocto.SetFocus
            Exit Sub
        End If
        RemoverLote = True
    End If

    For Count = 0 To lstDocto.ListCount - 1
        If lstDocto.Selected(Count) Then
            i = Indice(Val(Right(lstDocto.List(Count), 10)))
            If Not aDoc(i).Deleted And aDoc(i).TipoDocto = 1 Then
                If Left(aDoc(i).Leitura, 4) = "0600" Then
                    QtdMal = QtdMal + 1
                Else
                    QtdEnv = QtdEnv + 1
                End If
            End If
        End If
    Next
    
    If QtdMal > 0 Then
        Msg = "Entre os documentos selecionados para remover existe(m) " & _
            Format(QtdMal, "00") & " Capa(s) de Malote. Ao remover este(s) Malote(s), " & _
            "serão removidos automaticamente todos os documentos a ele relacionados." & _
            vbCrLf & "Confirma remover documentos?"
        If MsgBox(Msg, vbQuestion + vbYesNo, App.Title) = vbNo Then
            lstDocto.SetFocus
            Exit Sub
        End If
    ElseIf QtdEnv > 0 Then
        Msg = "Entre os documentos selecionados para remover existe(m) " & _
            Format(QtdEnv, "00") & " Envelope(s). Ao remover este(s) Envelope(s), " & _
            "serão removidos automaticamente todos os documentos a ele relacionados." & _
            vbCrLf & "Confirma remover documentos?"
        If MsgBox(Msg, vbQuestion + vbYesNo, App.Title) = vbNo Then
            lstDocto.SetFocus
            Exit Sub
        End If
    ElseIf Not RemoverLote Then
        Msg = "Confirma remover documentos?"
        If MsgBox(Msg, vbQuestion + vbYesNo, App.Title) = vbNo Then
            lstDocto.SetFocus
            Exit Sub
        End If
    End If
    
    m_CountDeleted = 0
    For Count = 0 To lstDocto.ListCount - 1
        If lstDocto.Selected(Count) Then
            i = Indice(Val(Right(lstDocto.List(Count), 10)))
            If Not aDoc(i).Deleted Then
                If aDoc(i).TipoDocto = 1 Then
                    If Not RemoveCapa(aDoc(i).IdCapa) Then
                        lstDocto.SetFocus
                        Exit Sub
                    End If
                Else
                    Geral.Banco.BeginTrans
                    If Not RemoveDocumento(aDoc(i).IdCapa, aDoc(i).IdDocto) Then
                        GoTo Exit_DelDocto
                    End If
                    
                    If Not OrdenaCaptura(aDoc(i).IdCapa) Then
                        GoTo Exit_DelDocto
                    End If
                    
                    Geral.Banco.CommitTrans
                    aDoc(i).Deleted = True
                End If
            End If
            ListIndex = lstDocto.ListIndex
        End If
    Next
    
    For Count = 1 To m_CountDocto - 1
        If Not aDoc(Count).Deleted And aDoc(Count).TipoDocto = 1 Then
            bExiste = False
            i = Count + 1
            Do While aDoc(Count).IdCapa = aDoc(i).IdCapa
                If Not aDoc(i).Deleted Then
                    bExiste = True
                End If
                i = i + 1
                If i > m_CountDocto Then
                    Exit Do
                End If
            Loop
            If Not bExiste Then
                If Not RemoveCapa(aDoc(Count).IdCapa) Then
                    lstDocto.SetFocus
                    Exit Sub
                End If
            End If
        End If
    Next
    
    bExiste = False
    For Count = 1 To m_CountDocto
        If Not aDoc(Count).Deleted Then
            bExiste = True
            Exit For
        End If
    Next
    
    If Not bExiste Then
        If Not RemoveLote(Val(lstLote.List(lstLote.ListIndex))) Then
            lstDocto.SetFocus
            Exit Sub
        End If
    End If
    
    Preenche_lstDocto
    If lstDocto.ListCount > 0 Then
        If ListIndex - m_CountDeleted + 1 < lstDocto.ListCount And _
           ListIndex - m_CountDeleted > 0 Then
            lstDocto.Selected(ListIndex - m_CountDeleted + 1) = True
        ElseIf ListIndex - m_CountDeleted > 0 Then
            lstDocto.Selected(ListIndex - m_CountDeleted) = True
        ElseIf lstDocto.ListCount > 0 Then
            lstDocto.Selected(0) = True
        End If
        lstDocto.SetFocus
    Else
        CmdAtualizar_Click
    End If
    
    Exit Sub
    
Exit_DelDocto:
    Geral.Banco.RollbackTrans
    MsgBox "Erro. Não foi possível remover o documento.", vbCritical + vbOKOnly, App.Title
    lstDocto.SetFocus

End Sub
Private Sub Form_Activate()

    'Inclusão de chamada a rotina AtualizaAtividade
    Call AtualizaAtividade(6)

    If m_FirstActivate Then
        LimparListas

        tmrAtualiza.Enabled = True
        sTempo = 0
        m_IdLote = 0

        If Not ObtemLote Then
            MsgBox "Não existem Lotes com pendência de Controle de Qualidade.", vbExclamation + vbOKOnly, App.Title
            m_IdLote = 0
            HabilitaTimerPesquisa
            Exit Sub
        End If
        m_FirstActivate = False
        lstLote.Selected(0) = True
    End If
End Sub

Private Sub Form_Load()
    
    With Lead1
        .UnlockSupport L_SUPPORT_EXPRESS, "YXPQ3XPPVT"
        .UnlockSupport L_SUPPORT_GIFLZW, "0K3RV9UY3EY"
        .UnlockSupport L_SUPPORT_TIFLZW, "9LE75L0FDXHK"
    End With
    With Lead2
        .UnlockSupport L_SUPPORT_EXPRESS, "YXPQ3XPPVT"
        .UnlockSupport L_SUPPORT_GIFLZW, "0K3RV9UY3EY"
        .UnlockSupport L_SUPPORT_TIFLZW, "9LE75L0FDXHK"
    End With

    Set qryGetLoteContQualidade = Geral.Banco.CreateQuery("", "{Call GetLoteContQualidade (?,?)}")
    Set qryGetDocumentoContQualidade = Geral.Banco.CreateQuery("", "{Call GetDocumentoContQualidade (?,?)}")
    Set qryGetCapasLote = Geral.Banco.CreateQuery("", "{Call GetCapasLote (?,?)}")
    Set qryAtualizaStatusLote = Geral.Banco.CreateQuery("", "{? = Call AtualizaStatusLote (?,?,?)}")
    Set qryVerificaLoteDisponivel = Geral.Banco.CreateQuery("", "{? = Call VerificaLoteDisponivel (?,?,?)}")
    Set qryRemoveDocumento = Geral.Banco.CreateQuery("", "{? = Call RemoveDocumento (?,?)}")
    Set qryOrdenaCaptura = Geral.Banco.CreateQuery("", "{? = call OrdenaCapturaSplitCapa(?,?)}")
    Set qryRemoveCapa = Geral.Banco.CreateQuery("", "{? = Call RemoveCapa (?,?)}")
    Set qryRemoveLote = Geral.Banco.CreateQuery("", "{? = Call RemoveLote (?,?)}")
    
    m_FirstActivate = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If m_Busy Then
        Cancel = 1
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    tmrAtualiza.Enabled = False
    tmrPesquisa.Enabled = False
    
    If m_IdLote > 0 Then
        AtualizaStatusLote m_IdLote, "0"
    End If
    
    qryGetLoteContQualidade.Close
    qryGetDocumentoContQualidade.Close
    qryAtualizaStatusLote.Close
    qryVerificaLoteDisponivel.Close
    qryRemoveDocumento.Close
    qryOrdenaCaptura.Close
    qryRemoveCapa.Close
    qryRemoveLote.Close
End Sub

Private Sub lstLote_Click()
    Dim Count As Integer
    Dim AindaExiste As Boolean
    
    If m_Busy Then
        Exit Sub
    End If
    m_Busy = True
    
    If m_IdLote > 0 Then
        If Not AtualizaStatusLote(m_IdLote, "0") Then
            m_Busy = False
            m_IdLote = 0
            Exit Sub
        End If
    End If
    
    If lstLote.ListCount > 0 Then
        m_IdLote = Val(lstLote.List(lstLote.ListIndex))
    End If
    
    If Not VerificaLoteDisponivel(m_IdLote) Then
        m_IdLote = 0
        m_Busy = False
        m_CountDocto = 0
        Preenche_lstDocto
        Exit Sub
    End If
    
    If Not AtualizaStatusLote(m_IdLote, "1") Then
        m_IdLote = 0
        m_Busy = False
        Exit Sub
    End If
    ObtemDocumentos m_IdLote
    sTempo = 0
    Preenche_lstDocto
    If lstDocto.ListCount > 0 Then
        lstDocto.Selected(0) = True
    End If
    
    m_Busy = False
End Sub

Private Sub lstLote_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys ("{TAB}")
    End If
End Sub

Private Sub lstDocto_Click()
    MostraImagem
    lstDocto.SetFocus
End Sub

Private Sub lstDocto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmdRemover_Click
    End If
End Sub

Private Sub tmrAtualiza_Timer()
    tmrAtualiza.Enabled = False
    If m_IdLote > 0 Then
        sTempo = sTempo + Int(tmrAtualiza.Interval / 1000)
        If sTempo + Int(tmrAtualiza.Interval / 1000) >= Geral.Intervalo Then
            AtualizaStatusLote m_IdLote, "1"
            sTempo = 0
        End If
    End If
    tmrAtualiza.Enabled = True
End Sub

Private Sub tmrPesquisa_Timer()
  tmrPesquisa.Enabled = False

  sTempo = sTempo + Int(tmrPesquisa.Interval / 1000)

  If sTempo + Int(tmrPesquisa.Interval / 1000) >= Geral.Atualizacao Then
    sTempo = 0
    If ObtemLote Then
        FrmPesquisa.Visible = False
        lstLote.Selected(0) = True
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

Private Sub LiberaLote(ByVal IdLote As Long)

    On Error GoTo ErroLiberaLote

    Screen.MousePointer = vbHourglass

    qryGetCapasLote.rdoParameters(0) = Geral.DataProcessamento
    qryGetCapasLote.rdoParameters(1) = IdLote
    Set rsCapa = qryGetCapasLote.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

    If Not rsCapa.EOF Then
        Do Until rsCapa.EOF
            'Grava Log da liberação da capa
            Call GravaLog(rsCapa!IdCapa, 0, 32)

            rsCapa.MoveNext
            DoEvents
        Loop
    End If

    Exit Sub

ErroLiberaLote:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro. Não foi possível gravar o log da liberação do lote.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
End Sub
