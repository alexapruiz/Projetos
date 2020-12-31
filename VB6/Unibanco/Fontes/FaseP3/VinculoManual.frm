VERSION 5.00
Object = "{00080003-B1BA-11CE-ABC6-F5B2E79D9E3F}#8.0#0"; "LTOCX80N.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form VinculoManual 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vínculo Manual"
   ClientHeight    =   8088
   ClientLeft      =   864
   ClientTop       =   432
   ClientWidth     =   10560
   ForeColor       =   &H8000000F&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8088
   ScaleWidth      =   10560
   Begin VB.Timer tmrPesquisa 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   10224
      Top             =   3972
   End
   Begin VB.PictureBox FrmPesquisa 
      Height          =   1932
      Left            =   2160
      ScaleHeight     =   1884
      ScaleWidth      =   5712
      TabIndex        =   45
      Top             =   2568
      Visible         =   0   'False
      Width           =   5760
      Begin VB.CommandButton CmdFecharPesquisa 
         Caption         =   "&Fechar"
         Height          =   312
         Left            =   2316
         TabIndex        =   46
         Top             =   1464
         Width           =   1068
      End
      Begin ComctlLib.ProgressBar Progress 
         Height          =   300
         Left            =   348
         TabIndex        =   47
         Top             =   912
         Width           =   5088
         _ExtentX        =   8975
         _ExtentY        =   529
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Pesquisando por Documentos para Vínculo. Aguarde ..."
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
         Left            =   276
         TabIndex        =   48
         Top             =   576
         Width           =   5304
      End
   End
   Begin VB.Timer tmrAtualiza 
      Enabled         =   0   'False
      Interval        =   50000
      Left            =   9840
      Top             =   3984
   End
   Begin TabDlg.SSTab tabVinculo 
      Height          =   3204
      Left            =   60
      TabIndex        =   22
      Top             =   444
      Width           =   8724
      _ExtentX        =   15388
      _ExtentY        =   5652
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   420
      TabCaption(0)   =   "&A Vincular"
      TabPicture(0)   =   "VinculoManual.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Vinc. &Manualmente"
      TabPicture(1)   =   "VinculoManual.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstManual"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Vinc. Au&tomaticamente"
      TabPicture(2)   =   "VinculoManual.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lstAutomatico"
      Tab(2).Control(1)=   "cmdDesfazerTudo"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "&Ocorrências"
      TabPicture(3)   =   "VinculoManual.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lstOcorrencia"
      Tab(3).ControlCount=   1
      Begin VB.CommandButton cmdDesfazerTudo 
         Caption         =   "Desfazer Vinc. Automático"
         Height          =   372
         Left            =   -71904
         TabIndex        =   8
         Top             =   2664
         Width           =   2460
      End
      Begin VB.ListBox lstOcorrencia 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2688
         Left            =   -74916
         TabIndex        =   9
         Top             =   324
         Width           =   8448
      End
      Begin VB.ListBox lstAutomatico 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2208
         Left            =   -74904
         TabIndex        =   7
         Top             =   324
         Width           =   8448
      End
      Begin VB.ListBox lstManual 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2688
         Left            =   -74916
         TabIndex        =   6
         Top             =   324
         Width           =   8448
      End
      Begin VB.Frame Frame6 
         Caption         =   "Depósitos / OCTs / Pagamentos"
         Height          =   2844
         Left            =   5364
         TabIndex        =   43
         Top             =   276
         Width           =   3276
         Begin VB.ListBox lstPagtos 
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
            Left            =   60
            MultiSelect     =   2  'Extended
            TabIndex        =   5
            Top             =   192
            Width           =   3120
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Cheques / Lançamentos"
         Height          =   2844
         Left            =   84
         TabIndex        =   42
         Top             =   276
         Width           =   3276
         Begin VB.ListBox lstCheques 
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
            Left            =   72
            MultiSelect     =   2  'Extended
            TabIndex        =   1
            Top             =   192
            Width           =   3120
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2856
         Left            =   3396
         TabIndex        =   35
         Top             =   276
         Width           =   1920
         Begin VB.CommandButton cmdOcorrencia 
            Caption         =   "Oco&rrência"
            Height          =   312
            Left            =   96
            TabIndex        =   4
            Top             =   2484
            Width           =   1764
         End
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "&Cancelar"
            Height          =   312
            Left            =   96
            TabIndex        =   3
            Top             =   2148
            Width           =   1764
         End
         Begin VB.CommandButton cmdVincular 
            Caption         =   "&Vincular"
            Height          =   312
            Left            =   96
            TabIndex        =   2
            Top             =   1812
            Width           =   1764
         End
         Begin VB.Label lblValorContas 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0,00"
            Height          =   276
            Left            =   84
            TabIndex        =   41
            Top             =   936
            Width           =   1764
         End
         Begin VB.Label lblValorCheques 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0,00"
            Height          =   276
            Left            =   84
            TabIndex        =   40
            Top             =   384
            Width           =   1764
         End
         Begin VB.Label lblValorDiferenca 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0,00"
            Height          =   276
            Left            =   84
            TabIndex        =   39
            Top             =   1476
            Width           =   1764
         End
         Begin VB.Label Label7 
            Caption         =   "Dep. / OCT / Pagtos.:"
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   84
            TabIndex        =   38
            Top             =   696
            Width           =   1620
         End
         Begin VB.Label Label8 
            Caption         =   "Cheques:"
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   84
            TabIndex        =   37
            Top             =   168
            Width           =   876
         End
         Begin VB.Label Label9 
            Caption         =   "Diferença:"
            ForeColor       =   &H000000C0&
            Height          =   240
            Left            =   84
            TabIndex        =   36
            Top             =   1248
            Width           =   876
         End
      End
   End
   Begin VB.ComboBox cmbCapa 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   1536
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   24
      Width           =   2412
   End
   Begin VB.PictureBox Picture4 
      Height          =   336
      Left            =   4020
      ScaleHeight     =   288
      ScaleWidth      =   528
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   24
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
         TabIndex        =   33
         Top             =   12
         Width           =   480
      End
   End
   Begin VB.PictureBox picNumMalote 
      Height          =   336
      Left            =   5952
      ScaleHeight     =   288
      ScaleWidth      =   1128
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   24
      Width           =   1176
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
         TabIndex        =   30
         Top             =   24
         Width           =   1092
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   336
      Left            =   72
      ScaleHeight     =   288
      ScaleWidth      =   1380
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   24
      Width           =   1428
      Begin VB.Label lblCapa 
         Caption         =   "Capa"
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
         TabIndex        =   24
         Top             =   12
         Width           =   1272
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3684
      Left            =   8880
      TabIndex        =   25
      Top             =   -36
      Width           =   1632
      Begin VB.CommandButton cmdEnviarCSP 
         Caption         =   "Enviar C&SP"
         Height          =   324
         Left            =   96
         TabIndex        =   15
         Top             =   2688
         Width           =   1464
      End
      Begin VB.CommandButton CmdExibirCapa 
         Caption         =   "&Exibir Capa"
         Height          =   324
         Left            =   96
         TabIndex        =   13
         Top             =   1704
         Width           =   1464
      End
      Begin VB.CommandButton cmdAtualizar 
         Caption         =   "&Atualizar"
         Height          =   324
         Left            =   84
         TabIndex        =   10
         Top             =   252
         Width           =   1464
      End
      Begin VB.CommandButton cmdFechar 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   324
         Left            =   84
         TabIndex        =   16
         Top             =   3192
         Width           =   1464
      End
      Begin VB.CommandButton cmdConfirmar 
         Caption         =   "Con&firmar"
         Height          =   324
         Left            =   84
         TabIndex        =   11
         Top             =   732
         Width           =   1464
      End
      Begin VB.CommandButton cmdIlegiveis 
         Caption         =   "Enviar &Ilegíveis"
         Height          =   324
         Left            =   84
         TabIndex        =   14
         Top             =   2184
         Width           =   1464
      End
      Begin VB.CommandButton cmdDesfazer 
         Caption         =   "&Desfazer"
         Height          =   324
         Left            =   84
         TabIndex        =   12
         Top             =   1212
         Width           =   1464
      End
   End
   Begin VB.Frame frmImagem 
      Caption         =   "Imagem"
      Height          =   4020
      Left            =   60
      TabIndex        =   26
      Top             =   4008
      Width           =   8748
      Begin LeadLib.Lead Lead1 
         Height          =   3648
         Left            =   108
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   252
         Width           =   8400
         _Version        =   524288
         _ExtentX        =   14817
         _ExtentY        =   6435
         _StockProps     =   229
         BackColor       =   -2147483639
         BorderStyle     =   1
         ScaleHeight     =   302
         ScaleWidth      =   698
         DataField       =   ""
         BitmapDataPath  =   ""
         AnnDataPath     =   ""
      End
   End
   Begin VB.Frame Frame5 
      Height          =   4020
      Left            =   8880
      TabIndex        =   27
      Top             =   4008
      Width           =   1632
      Begin VB.CommandButton cmdAuditoria 
         Caption         =   "A&uditoria"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   564
         Left            =   384
         Picture         =   "VinculoManual.frx":0070
         Style           =   1  'Graphical
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   288
         Width           =   900
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
         Height          =   588
         Left            =   384
         Picture         =   "VinculoManual.frx":01FA
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   3264
         Width           =   900
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
         Height          =   588
         Left            =   384
         Picture         =   "VinculoManual.frx":0504
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2664
         Width           =   900
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
         Height          =   588
         Left            =   384
         Picture         =   "VinculoManual.frx":080E
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2064
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
         Height          =   588
         Left            =   384
         Picture         =   "VinculoManual.frx":0B18
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1464
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
         Height          =   588
         Left            =   384
         Picture         =   "VinculoManual.frx":0E22
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   864
         Width           =   900
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
      TabIndex        =   44
      Top             =   3696
      Width           =   8724
   End
   Begin VB.Label lblLote 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1270 - 00001"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   4644
      TabIndex        =   34
      Top             =   24
      Width           =   1200
   End
   Begin VB.Label lblNumMalote 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "06001100741"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   7188
      TabIndex        =   31
      Top             =   24
      Width           =   1500
   End
End
Attribute VB_Name = "VinculoManual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type tpMyCapa
    IdCapa                                  As Long
    IdLote                                  As Long
    IdEnv_Mal                               As String * 1
    Capa                                    As String
    NumMalote                               As String
    RegraNova                               As Boolean
    AgOrig                                  As Integer

End Type

Private Type tpMyDoc
    IdDocto                                 As Long
    TipoDocto                               As Integer
    TipoDoctoAnt                            As Integer
    Ocorrencia                              As Integer
    Frente                                  As String
    Verso                                   As String
    Status                                  As String * 1
    StatusAnt                               As String * 1
    Vinculo                                 As Long
    Valor                                   As Currency
    Ordem                                   As String * 1
    Leitura                                 As String
    Alcada                                  As String * 1
    AlcadaAnt                               As String * 1
    ComplementoOcorrencia                   As String
End Type

Private Type tpMyAjuste
    TipoDocto                               As Integer
    Vinculo                                 As Long
    Agencia                                 As Integer
    Conta                                   As Long
    Valor                                   As Currency
End Type

Private m_Busy                              As Boolean
Private m_ValContas                         As Currency
Private m_ValCheques                        As Currency
Private m_ValDiferenca                      As Currency
Private m_IdCapa                            As Long
Private m_Capa                              As tpMyCapa
Private m_Doc                               As tpMyDoc
Private m_Ajuste                            As tpMyAjuste
Private aCapa()                             As tpMyCapa
Private aDoc()                              As tpMyDoc
Private aCheque()                           As tpMyDoc
Private aPagto()                            As tpMyDoc
Private aAjuste(1 To 50)                    As tpMyAjuste
Private m_CountCapa                         As Integer
Private m_CountDocto                        As Integer
Private m_CountCheque                       As Integer
Private m_CountPagto                        As Integer
Private m_CountAjuste                       As Integer
Private m_Frente                            As String
Private m_Verso                             As String
Private m_Ordem                             As String * 1
Private m_ValorAjusteVincManual_Mal         As Currency
Private m_ValorAjusteVincManual_Env         As Currency
Private m_ValorAlcadaSaque_Env              As Currency
Private m_ValorAlcadaSaque_Mal              As Currency
Private m_ValorAjusteContabil               As Currency
Private m_ValorAlcadaOutros_Env             As Currency
Private m_ValorAlcadaOutros_Mal             As Currency
Private m_ValorLimiteMaxDifLancto           As Currency
Private m_FirstActivate                     As Boolean
Private sTempo                              As Integer

Private qryGetCapaVinculoManual             As rdoQuery
Private qryGetDocumentoVinculoManual        As rdoQuery
Private qryGetocorrencia                    As rdoQuery
Private qryAtualizaStatusCapa               As rdoQuery
Private qryAtualizaVinculoDocumento         As rdoQuery
Private qryAtualizaDocumentoExcluido        As rdoQuery
Private qryInsereAjuste                     As rdoQuery
Private qryGetDescricaoDocumento            As rdoQuery
Private qryLerParametros                    As rdoQuery
Private qryGetAgContaDocumento              As rdoQuery
Private qryVerificaCapaDisponivel           As rdoQuery
Private qryGetUltimaOrdemCaptura            As rdoQuery
Private qryRemoveAjustesVinculoManual       As rdoQuery
Private qryVA_GetDocumentosTransmitidos     As rdoQuery
Private qryGetImagemCapa                    As rdoQuery

Private rsCapa                              As rdoResultset
Private rsDoc                               As rdoResultset
Private rsOrdemCaptura                      As rdoResultset
Private RsOcorrencia                        As rdoResultset
Private rsDescricao                         As rdoResultset
Private rsParametro                         As rdoResultset
Private rsAgConta                           As rdoResultset
Private RsDoctosTrans                       As rdoResultset
Private i                                   As Long


Private Function ExisteDeptoPagto() As Boolean

    Dim Count       As Integer
    Dim CountAux    As Integer
    
    ExisteDeptoPagto = True

    For Count = 0 To lstPagtos.ListCount - 1
        If lstPagtos.Selected(Count) Then
            If aPagto(Count + 1).TipoDocto = 2 Or aPagto(Count + 1).TipoDocto = 3 Then
                For CountAux = 0 To lstPagtos.ListCount - 1
                    If lstPagtos.Selected(CountAux) Then
                        If (aPagto(CountAux + 1).TipoDocto <> 2 And aPagto(CountAux + 1).TipoDocto <> 3) And (Count <> CountAux) Then
                            Exit Function
                        End If
                    End If
                Next CountAux
            End If
        End If
    Next Count
    
    ExisteDeptoPagto = False

End Function

Private Sub LimparValores()
    lblValorContas.Caption = ""
    lblValorCheques.Caption = ""
    lblValorDiferenca.Caption = ""
End Sub

Private Sub LimparHeader()
    lblCapa.Caption = "Capa"
    lblNumMalote.Caption = ""
    lblLote.Caption = ""
    lblOcorrencia.Caption = ""
End Sub

Private Sub LimparListas()
    tabVinculo.Tab = 0
    cmbCapa.Clear
    lstCheques.Clear
    lstPagtos.Clear
    lstManual.Clear
    lstAutomatico.Clear
    lstOcorrencia.Clear
End Sub

Private Sub ObtemParametros()
    On Error GoTo ErroParametro
    rdoErrors.Clear

    Screen.MousePointer = vbHourglass

    qryLerParametros.rdoParameters(0) = Geral.DataProcessamento
    Set rsParametro = qryLerParametros.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    Screen.MousePointer = vbDefault
    If rsParametro.EOF Then
        rsParametro.Close
        GoTo ErroParametro
        Exit Sub
    End If
    m_ValorAjusteVincManual_Env = rsParametro!ValorAjusteVincManual_Env
    m_ValorAjusteVincManual_Mal = rsParametro!ValorAjusteVincManual_Mal
    m_ValorAlcadaSaque_Env = rsParametro!ValorAlcada_Env

    m_ValorAlcadaOutros_Mal = rsParametro!ValorAlcadaOutros_Mal
    m_ValorAlcadaOutros_Env = rsParametro!ValorAlcadaOutros_Env

    m_ValorAlcadaSaque_Mal = rsParametro!ValorAlcada_Mal
    m_ValorAjusteContabil = rsParametro!ValorAjusteContabil

    m_ValorLimiteMaxDifLancto = rsParametro!LimiteMaxDifLancto_Mal
    rsParametro.Close
    On Error GoTo 0
    Exit Sub

ErroParametro:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro na obtenção dos parâmetros de ajuste para Vínculo Manual.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    m_Busy = False
    Unload Me
End Sub
Private Function ObtemCapas() As Boolean
    ObtemParametros

    On Error GoTo ErroGetCapa
    rdoErrors.Clear

    Screen.MousePointer = vbHourglass

    Erase aCapa
    m_CountCapa = 0

    qryGetCapaVinculoManual.rdoParameters(0) = Geral.DataProcessamento
    qryGetCapaVinculoManual.rdoParameters(1) = Geral.Intervalo
    Set rsCapa = qryGetCapaVinculoManual.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    If rsCapa.EOF Then
        rsCapa.Close
        ObtemCapas = False
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    ReDim aCapa(rsCapa.RowCount)
    While Not rsCapa.EOF
        m_CountCapa = m_CountCapa + 1
        m_Capa.IdCapa = rsCapa!IdCapa
        m_Capa.IdLote = Format(rsCapa!IdLote, "000000000")
        m_Capa.IdEnv_Mal = rsCapa!IdEnv_Mal
        m_Capa.Capa = rsCapa!Capa
        m_Capa.AgOrig = rsCapa!AgOrig
        If Left(Format(rsCapa!Num_Malote, "00000000000"), 1) = "9" Then
            m_Capa.NumMalote = Format(rsCapa!Num_Malote, "000000000000")
            m_Capa.RegraNova = True
        Else
            m_Capa.NumMalote = Format(rsCapa!Num_Malote, "00000000000")
            m_Capa.RegraNova = False
        End If
        aCapa(m_CountCapa) = m_Capa
        rsCapa.MoveNext
    Wend
    rsCapa.Close
    ObtemCapas = True
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    Exit Function

ErroGetCapa:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro na obtenção de envelope/malote para Vínculo Manual.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    m_Busy = False
    Unload Me
End Function
Private Sub ObtemDocumentos(ByVal IdCapa As Long)

    On Error GoTo ErroGetDocto
    rdoErrors.Clear

    Screen.MousePointer = vbHourglass

    Erase aDoc
    Erase aCheque
    Erase aPagto
    Erase aAjuste
    m_CountDocto = 0
    m_CountCheque = 0
    m_CountPagto = 0
    m_CountAjuste = 0

    qryGetDocumentoVinculoManual.rdoParameters(0) = Geral.DataProcessamento
    qryGetDocumentoVinculoManual.rdoParameters(1) = IdCapa
    Set rsDoc = qryGetDocumentoVinculoManual.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

    ReDim aDoc(rsDoc.RowCount)
    ReDim aCheque(rsDoc.RowCount)
    ReDim aPagto(rsDoc.RowCount)

    While Not rsDoc.EOF
        m_Doc.IdDocto = rsDoc!IdDocto
        m_Doc.TipoDocto = rsDoc!TipoDocto
        m_Doc.TipoDoctoAnt = m_Doc.TipoDocto
        m_Doc.Ocorrencia = rsDoc!Ocorrencia
        m_Doc.Frente = rsDoc!Frente
        m_Doc.Verso = rsDoc!Verso
        m_Doc.Status = rsDoc!Status
        m_Doc.StatusAnt = m_Doc.Status
        m_Doc.Ordem = rsDoc!Ordem
        m_Doc.Vinculo = rsDoc!Vinculo
        m_Doc.Valor = rsDoc!Valor
        m_Doc.Leitura = Trim(rsDoc!Leitura)
        m_Doc.Alcada = rsDoc!Alcada
        m_Doc.AlcadaAnt = m_Doc.Alcada

        m_CountDocto = m_CountDocto + 1
        aDoc(m_CountDocto) = m_Doc

        If rsDoc!Vinculo = 0 And rsDoc!Status <> "D" And rsDoc!Status <> "F" Then
            If rsDoc!TipoDocto >= 4 And rsDoc!TipoDocto <= 7 Or _
               rsDoc!TipoDocto = 33 Or rsDoc!TipoDocto = 38 Or _
               rsDoc!TipoDocto = 41 Then
                m_CountCheque = m_CountCheque + 1
                aCheque(m_CountCheque) = m_Doc
            Else
                If rsDoc!TipoDocto <> 39 Then
                    m_CountPagto = m_CountPagto + 1
                    aPagto(m_CountPagto) = m_Doc
                End If
            End If
        End If
    
        rsDoc.MoveNext
    Wend
    rsDoc.Close
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErroGetDocto:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro na obtenção de documentos para Vínculo Manual.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    m_Busy = False
    Unload Me

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

Private Sub ObtemOcorrencia()
    Dim Ocorrencia As Long
    
    On Error GoTo ErroOcorrencia
    rdoErrors.Clear
    
    Screen.MousePointer = vbHourglass
    
    '3 Posicoes
    If Left(Right(Left(lstOcorrencia.List(lstOcorrencia.ListIndex), Len(lstOcorrencia.List(lstOcorrencia.ListIndex)) - 15), 5), 2) = "00" Then
       'Ocorrencia gerada pelo sistema
       Ocorrencia = Val(Right(Left(lstOcorrencia.List(lstOcorrencia.ListIndex), Len(lstOcorrencia.List(lstOcorrencia.ListIndex)) - 15), 5))
    Else
       'Ocorrencia atualizada pelo robo
       Ocorrencia = Val(Right(Left(lstOcorrencia.List(lstOcorrencia.ListIndex), Len(lstOcorrencia.List(lstOcorrencia.ListIndex)) - 15), 5)) / 100
    End If
    
    qryGetocorrencia.rdoParameters(0) = Ocorrencia
    Set RsOcorrencia = qryGetocorrencia.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    If RsOcorrencia.EOF Then
        Screen.MousePointer = vbDefault
        MsgBox "Código da Ocorrência não existe: " & Mid(lstOcorrencia.List(lstOcorrencia.ListIndex), 74, 4) & ".", vbExclamation + vbOKOnly, App.Title
    Else
        lblOcorrencia.Caption = "Ocorrência: " & RsOcorrencia!Descricao
    End If
    RsOcorrencia.Close
    Screen.MousePointer = vbDefault
    On Error GoTo 0
    Exit Sub
    
ErroOcorrencia:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro na obtenção da Ocorrência do documento.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    m_Busy = False
    Unload Me

End Sub

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
Private Function VerificaCapaDisponivel(ByVal IdCapa As Long) As Boolean
    On Error GoTo ErroVerificaCapa
    rdoErrors.Clear

    Screen.MousePointer = vbHourglass

    With qryVerificaCapaDisponivel
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = IdCapa
        .rdoParameters(3) = "7"
        .rdoParameters(4) = "J"
        .rdoParameters(5) = Geral.Intervalo
        .Execute
        Screen.MousePointer = vbDefault
        If .rdoParameters(0) = 0 Then
            VerificaCapaDisponivel = True
        ElseIf .rdoParameters(0) = 1 Then
            VerificaCapaDisponivel = False
            MsgBox "Este Envelope/Malote não está mais disponível por já ter sido tratado ou porque esta sendo tratado por outra estação.", vbInformation + vbOKOnly, App.Title
        Else
            VerificaCapaDisponivel = False
            MsgBox "Erro. Não foi possível obter o Status do Envelope/Malote.", vbInformation + vbOKOnly, App.Title
        End If
    End With
    On Error GoTo 0
    Exit Function
    
ErroVerificaCapa:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro. Não foi possível obter o Status do Envelope/Malote.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    m_Busy = False
    Unload Me
    
End Function

Private Function AtualizaStatusCapa(ByVal IdCapa As Long, ByVal Status As String) As Boolean
    On Error GoTo ErroAtualizaStatus
    rdoErrors.Clear
    
    AtualizaStatusCapa = True
    Screen.MousePointer = vbHourglass
    
    With qryAtualizaStatusCapa
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = IdCapa
        .rdoParameters(3) = Status
        .Execute
        If .rdoParameters(0) <> 0 Then
            AtualizaStatusCapa = False
            Screen.MousePointer = vbDefault
            MsgBox "Erro na atualização do status do envelope/malote.", vbCritical + vbOKOnly, App.Title
            cmdDesfazer_Click
        End If
    End With
    Screen.MousePointer = vbDefault
    On Error GoTo 0
    Exit Function
    
ErroAtualizaStatus:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro na atualização do status do envelope/malote.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    m_Busy = False
    cmdDesfazer_Click
    Unload Me

End Function


Private Sub Preenche_cmbCapa()
    Dim Count As Integer
    cmbCapa.Clear
    For Count = 1 To m_CountCapa
        cmbCapa.AddItem aCapa(Count).Capa
    Next
End Sub

Private Sub Preenche_lstCheques()
    Dim Linha As String
    Dim Count As Integer
    
    lstCheques.Clear
    For Count = 1 To m_CountCheque
        If aCheque(Count).Vinculo = 0 And aCheque(Count).Status <> "D" And aCheque(Count).Status <> "F" Then
            Select Case aCheque(Count).TipoDocto
                Case 4          'ADCC
                    Linha = "DEBITO CC" & Space(1)
                Case 5          'Cheque UBB Sacado
                    Linha = "CH UBB   " & Space(1)
                Case 6          'Cheque Terceiro Pagto
                    If InStr("409*230", Left(aCheque(Count).Leitura, 3)) > 0 Then
                        Linha = "CH UBB   " & Space(1)
                    Else
                        Linha = "CH TERC  " & Space(1)
                    End If
                Case 7          'Cheque de Deposito
                    Linha = "CH DEP   " & Space(1)
                Case 33, 38     'Ajuste Debito
                    Linha = "AJ DEBITO" & Space(1)
                Case 41         'Lancamento Interno
                    Linha = "LANCTO   " & Space(1)
            End Select
            Linha = Linha & FormataValor(aCheque(Count).Valor, 16) & Space(3)
            Linha = Linha & Format(aCheque(Count).IdDocto, "0000000000")
            lstCheques.AddItem Linha
        End If
    Next
End Sub

Private Sub Preenche_lstPagtos()
    Dim Linha As String
    Dim Count As Integer
    
    lstPagtos.Clear
    For Count = 1 To m_CountPagto
        If aPagto(Count).Vinculo = 0 And aPagto(Count).Status <> "D" And aPagto(Count).Status <> "F" Then
            Select Case aPagto(Count).TipoDocto
                Case 2, 3       ' Deposito
                    Linha = "DEPOSITO " & Space(1)
                Case 12 Or 31
                    Linha = "PAGTO TER" & Space(1)
                Case 37         ' Oct
                    Linha = "OCT      " & Space(1)
                Case 32, 34, 42 ' Ajuste Credito
                    Linha = "AJ CRED  " & Space(1)
                Case Else       ' Pagamentos
                    Linha = "PAGAMENTO" & Space(1)
            End Select
            Linha = Linha & FormataValor(aPagto(Count).Valor, 16) & Space(3)
            Linha = Linha & Format(aPagto(Count).IdDocto, "0000000000")
            lstPagtos.AddItem Linha
        End If
    Next
End Sub

Private Sub Preenche_lstAutomatico()

    Dim Count As Integer
    Dim iVinculo As Long

    lstAutomatico.Clear
    iVinculo = 0
    For Count = 1 To m_CountDocto
        If aDoc(Count).Vinculo > 0 And aDoc(Count).Status <> "D" And aDoc(Count).Status <> "F" Then
            If iVinculo <> aDoc(Count).Vinculo Then
                If iVinculo <> 0 Then
                    lstAutomatico.AddItem " "
                End If
                iVinculo = aDoc(Count).Vinculo
                lstAutomatico.AddItem "Vínculo Nº " & Trim(str(aDoc(Count).Vinculo))
            End If
            
            lstAutomatico.AddItem Space(5) & RPad(ObtemDescricaoDocto(aDoc(Count).TipoDocto), 30) & _
                Space(2) & FormataValor(aDoc(Count).Valor, 20) & Space(13) & Trim(str(aDoc(Count).IdDocto))
        End If
    Next

    If lstAutomatico.ListCount > 0 Then
        cmdDesfazerTudo.Enabled = True
    Else
        cmdDesfazerTudo.Enabled = False
    End If
End Sub
Private Sub Adiciona_lstManual()

    Dim Count As Integer
    Dim i As Integer
    Dim iVinculo As Long
   
    If lstManual.ListCount > 0 Then
        lstManual.AddItem " "
    End If
    iVinculo = 0
    
    For Count = 0 To lstPagtos.ListCount - 1
        If lstPagtos.Selected(Count) Then
            i = IndicePagto(Val(Right(lstPagtos.List(Count), 10)))
            If iVinculo = 0 Then
                iVinculo = aPagto(i).Vinculo
                lstManual.AddItem "Vínculo Nº " & iVinculo
            End If
            lstManual.AddItem Space(5) & RPad(ObtemDescricaoDocto(aPagto(i).TipoDocto), 30) & _
                Space(2) & FormataValor(aPagto(i).Valor, 20) & Space(13) & aPagto(i).IdDocto
        End If
    Next
    For Count = 0 To lstCheques.ListCount - 1
        If lstCheques.Selected(Count) Then
            i = IndiceCheque(Val(Right(lstCheques.List(Count), 10)))
            lstManual.AddItem Space(5) & RPad(ObtemDescricaoDocto(aCheque(i).TipoDocto), 30) & _
                Space(2) & FormataValor(aCheque(i).Valor, 20) & Space(13) & aCheque(i).IdDocto
        End If
    Next
    For Count = 1 To m_CountAjuste
        If aAjuste(Count).Vinculo = iVinculo Then
            lstManual.AddItem Space(5) & RPad(ObtemDescricaoDocto(aAjuste(Count).TipoDocto), 30) & _
                Space(2) & FormataValor(aAjuste(Count).Valor, 20)
        End If
    Next
End Sub
Private Sub Preenche_lstOcorrencia()

    Dim Count As Integer

    lstOcorrencia.Clear
    For Count = 1 To m_CountDocto
        If aDoc(Count).Status = "D" Or aDoc(Count).Status = "F" Then
            lstOcorrencia.AddItem Space(1) & RPad(ObtemDescricaoDocto(aDoc(Count).TipoDocto), 30) & _
                Space(1) & FormataValor(aDoc(Count).Valor, 16) & " Ocorrência Nº " & _
                Format(aDoc(Count).Ocorrencia, "00000") & LPad(aDoc(Count).IdDocto, 15)
        End If
    Next
    For Count = 1 To m_CountCheque
        If aCheque(Count).Status = "D" Or aCheque(Count).Status = "F" Then
            lstOcorrencia.AddItem Space(1) & RPad(ObtemDescricaoDocto(aCheque(Count).TipoDocto), 30) & _
                Space(1) & FormataValor(aCheque(Count).Valor, 16) & " Ocorrência Nº " & _
                Format(aCheque(Count).Ocorrencia, "00000") & LPad(aCheque(Count).IdDocto, 15)
        End If
    Next
    For Count = 1 To m_CountPagto
        If aPagto(Count).Status = "D" Or aPagto(Count).Status = "F" Then
            lstOcorrencia.AddItem Space(1) & RPad(ObtemDescricaoDocto(aPagto(Count).TipoDocto), 30) & _
                Space(1) & FormataValor(aPagto(Count).Valor, 16) & " Ocorrência Nº " & _
                Format(aPagto(Count).Ocorrencia, "00000") & LPad(aPagto(Count).IdDocto, 15)
        End If
    Next
End Sub
Private Function CalculaValores() As Boolean
    Dim Count As Integer
    Dim SelCheque As Boolean
    Dim SelPagto As Boolean
   
    m_ValContas = 0
    m_ValCheques = 0
    m_ValDiferenca = 0
    SelCheque = False
    SelPagto = False
    
    For Count = 0 To lstCheques.ListCount - 1
        If lstCheques.Selected(Count) Then
            SelCheque = True
            m_ValCheques = m_ValCheques + aCheque(IndiceCheque(Val(Right(lstCheques.List(Count), 10)))).Valor
        End If
    Next
    For Count = 0 To lstPagtos.ListCount - 1
        If lstPagtos.Selected(Count) Then
            SelPagto = True
            m_ValContas = m_ValContas + aPagto(IndicePagto(Val(Right(lstPagtos.List(Count), 10)))).Valor
        End If
    Next
    
    m_ValDiferenca = m_ValCheques - m_ValContas
    lblValorContas.Caption = FormataValor(m_ValContas, 20)
    lblValorCheques.Caption = FormataValor(m_ValCheques, 20)
    lblValorDiferenca.Caption = FormataValor(m_ValDiferenca, 21)
    
    If SelCheque Or SelPagto Then
        cmdOcorrencia.Enabled = True
        cmdCancelar.Enabled = True
    End If
    If SelCheque And SelPagto Then
        If m_ValDiferenca < 0 And m_Capa.IdEnv_Mal = "M" And Left(Val(m_Capa.NumMalote), 1) <> "9" Then
             cmdVincular.Enabled = True
        End If
    ElseIf Not SelCheque And Not SelPagto Then
        cmdOcorrencia.Enabled = False
        cmdCancelar.Enabled = False
        cmdVincular.Enabled = False
    End If
    
    If m_ValDiferenca <> 0 Then
        CalculaValores = False
    Else
        CalculaValores = True
    End If
End Function
Private Sub MostraImagem()
    
    If Len(m_Frente) = 0 Then
        Exit Sub
    End If
    
    hCtl = Lead1.hwnd
    '''''''''''''''''''''''''''
    ' mostra imagem escolhida '
    '''''''''''''''''''''''''''
    On Error GoTo ErroImagem
    With Lead1
       .Tag = "F"
       .AutoRepaint = False
       If Geral.VIPSDLL = eDllProservi Then
         .Load Geral.DiretorioImagens & m_Frente, 0, 0, 1
       Else
         .Load Geral.DiretorioImagens & Format(aCapa(cmbCapa.ListIndex + 1).IdLote, "000000000") & "\" & m_Frente, 0, 0, 1
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
    frmImagem.Visible = True
    
    'posiciona imagem sempre no começo
    Ret = SendMessage(hCtl, WM_VSCROLL, SB_BOTTOM, 0)
    Ret = SendMessage(hCtl, WM_HSCROLL, SB_TOP, 0)
    
    cmdAuditoria.Enabled = True
    cmdZoomMais.Enabled = True
    cmdZoomMenos.Enabled = True
    cmdRotacao.Enabled = True
    cmdInverteCor.Enabled = True
    cmdFrenteVerso.Enabled = True
    On Error GoTo 0
    Exit Sub
    
ErroImagem:
    cmdAuditoria.Enabled = False
    cmdZoomMais.Enabled = False
    cmdZoomMenos.Enabled = False
    cmdRotacao.Enabled = False
    cmdInverteCor.Enabled = False
    cmdFrenteVerso.Enabled = False
    frmImagem.Visible = False
    MsgBox "Não foi possível exibir imagem do documento, imagem não encontrada", vbExclamation + vbOKOnly, App.Title

End Sub

Private Sub LimparImagem()
    frmImagem.Visible = False
    cmdAuditoria.Enabled = False
    cmdZoomMais.Enabled = False
    cmdZoomMenos.Enabled = False
    cmdRotacao.Enabled = False
    cmdInverteCor.Enabled = False
    cmdFrenteVerso.Enabled = False
End Sub

Private Function IndiceCheque(ByVal IdDocto As Long) As Integer
    Dim Count As Integer
    
    For Count = 1 To m_CountCheque
        If aCheque(Count).IdDocto = IdDocto Then
            IndiceCheque = Count
            Exit Function
        End If
    Next
    IndiceCheque = 0
End Function

Private Function IndicePagto(ByVal IdDocto As Long) As Integer
    Dim Count As Integer
    
    For Count = 1 To m_CountPagto
        If aPagto(Count).IdDocto = IdDocto Then
            IndicePagto = Count
            Exit Function
        End If
    Next
    IndicePagto = 0
End Function

Private Function IndiceDocto(ByVal IdDocto As Long) As Integer
    Dim Count As Integer
    
    For Count = 1 To m_CountDocto
        If aDoc(Count).IdDocto = IdDocto Then
            IndiceDocto = Count
            Exit Function
        End If
    Next
    IndiceDocto = 0
End Function

Function VerificaDocumentosTransmitidos() As Boolean

    On Error GoTo VerificaDocumentosTransmitidos_Err

    With qryVA_GetDocumentosTransmitidos
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = aCapa(Me.cmbCapa.ListIndex + 1).IdCapa
    End With

    Set RsDoctosTrans = qryVA_GetDocumentosTransmitidos.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    If Not RsDoctosTrans.EOF Then
        If RsDoctosTrans!Qtde > 0 Then
            VerificaDocumentosTransmitidos = True
            
            'Atualizar o Status da Capa para 'V' - Em Analise
            If AtualizaStatusCapa(aCapa(Me.cmbCapa.ListIndex + 1).IdCapa, "V") Then
                'Gravar Log
                Call GravaLog(aCapa(Me.cmbCapa.ListIndex + 1).IdCapa, 0, 76)
                MsgBox "Este Envelope/Malote não está mais disponível por já ter sido tratado ou porque esta sendo tratado por outra estação.", vbInformation + vbOKOnly, App.Title
            End If
        Else
            VerificaDocumentosTransmitidos = False
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
Private Sub VerificaVinculoEnvelope()

    Dim IndCh As Integer
    Dim IndPg As Integer

    cmdVincular.Enabled = False

    ' Se Diferenca Negativa envelope nao processa
    If m_ValDiferenca < 0 Then
        Exit Sub
    End If

    ' Nao pode haver mais de um cheque selecionado
    ' e deve haver pelo menos um pagto selecionado
    If lstCheques.SelCount <> 1 Or lstPagtos.SelCount < 1 Then
        Exit Sub
    End If

    ''''''''''''''''''''''''''''''''''''''''''
    'não é permitido depositos com pagamentos'
    ''''''''''''''''''''''''''''''''''''''''''
    If ExisteDeptoPagto() Then Exit Sub

    IndCh = IndiceCheque(Val(Right(lstCheques.List(lstCheques.ListIndex), 10)))

    If aCheque(IndCh).TipoDocto <> 4 And ((Left(aCheque(IndCh).Leitura, 3) <> "409") And _
                                          (Left(aCheque(IndCh).Leitura, 3) <> "230")) Then
        'Cheque Terceiros
        If lstPagtos.SelCount = 1 Then
            ' 1 x 1
            IndPg = IndicePagto(Val(Right(lstPagtos.List(lstPagtos.ListIndex), 10)))

            'Verificar se o titulo é do UNIBANCO ou do Bandeirantes
            If (Left(aPagto(IndPg).Leitura, 3) = "409" Or Left(aPagto(IndPg).Leitura, 3) = "230") _
            And m_ValDiferenca = 0 Then
                cmdVincular.Enabled = True
                Exit Sub
            End If
            '''''''''''''''''''''''''''''''''''''''
            'Não pode vincular com Concessionarias'
            '''''''''''''''''''''''''''''''''''''''
            If aPagto(IndPg).TipoDocto = 20 Or _
               aPagto(IndPg).TipoDocto = 21 Or _
               aPagto(IndPg).TipoDocto = 22 Or _
               aPagto(IndPg).TipoDocto = 23 Then
                cmdVincular.Enabled = False
                Exit Sub
            End If

            'Se for cartão avulso pode vincular com envelope (fase 6.3 - requisitado pelo usuario)
            If (aPagto(IndPg).TipoDocto = 36) And m_ValDiferenca = 0 Then
                cmdVincular.Enabled = True
                Exit Sub
            End If
            
            'Se docto = cobranca de terceiros ou título de outros bancos -> Não processa em envelopes
            If (aPagto(IndPg).TipoDocto = 12 Or aPagto(IndPg).TipoDocto = 31) Then
                Exit Sub
            End If
        Else
            ' 1 x N
            ' Nao processa
        End If
    Else
        ' Cheque Ubb
        cmdVincular.Enabled = True
    End If
End Sub
Private Sub VerificaVinculoMaloteRegraAntiga()

    Dim Count           As Integer
    Dim CountAux        As Integer
    Dim IndCh           As Integer
    Dim PossuiLancto    As Boolean
    Dim PossuiADCC      As Boolean

    cmdVincular.Enabled = False

    'Deve haver pelo nenos um cheque e um pagto selecionado
    If lstCheques.SelCount < 1 Or lstPagtos.SelCount < 1 Then
        Exit Sub
    End If

    'A unica restricao do malote velho é um ADCC pagando uma conta junto com outro cheque
    PossuiLancto = False
    PossuiADCC = False
    For Count = 0 To lstCheques.ListCount - 1
        If lstCheques.Selected(Count) Then
            IndCh = IndiceCheque(Val(Right(lstCheques.List(Count), 10)))

            If aCheque(IndCh).TipoDocto = 41 Then
                PossuiLancto = True
            End If

            If aCheque(IndCh).TipoDocto = 4 Then
                PossuiADCC = True
            End If
        End If
        DoEvents
    Next Count

    'Mais de um credito com ADCC sem lancamento -> Nao Processa
    If PossuiLancto = False And lstCheques.SelCount > 1 And PossuiADCC = True Then Exit Sub

    'Verificar se existe diferenca maior que ajuste contabil parametrizado
    If Abs(m_ValDiferenca) > m_ValorAjusteVincManual_Mal And m_ValDiferenca < 0 And PossuiLancto = True Then Exit Sub
    ''''''''''''''''''''''''''''''''''''''''''
    'não é permitido depositos com pagamentos'
    ''''''''''''''''''''''''''''''''''''''''''
    If PossuiLancto = False Then
        If ExisteDeptoPagto() Then Exit Sub
    End If

    cmdVincular.Enabled = True
End Sub
Private Sub VerificaVinculoMaloteRegraNova()

    Dim Count                   As Integer
    Dim CountAux                As Integer
    Dim IndCh                   As Integer
    Dim PossuiLancto            As Boolean
    Dim PossuiChqTerceiro       As Boolean
    Dim PossuiADCC              As Boolean
    
    cmdVincular.Enabled = False

    'Deve haver pelo nenos um cheque selecionado
    'e deve haver pelo menos um pagto selecionado
    If lstCheques.SelCount < 1 Or lstPagtos.SelCount < 1 Then
        Exit Sub
    End If

    'Nao pode haver uma ADCC selecionada junto com outros cheques e /ou lancto
    'Nao pode haver um cheque de outros bancos selecionado junto com cheques do unibanco
    'Verificar se foi selecionado algum lancamento interno ou cheque de terceiro
    PossuiLancto = False
    PossuiChqTerceiro = False
    PossuiADCC = False
    
    'Verifica lista de Créditos
    For Count = 0 To lstCheques.ListCount - 1
        If lstCheques.Selected(Count) Then
            IndCh = IndiceCheque(Val(Right(lstCheques.List(Count), 10)))

            If aCheque(IndCh).TipoDocto = 41 Then
                PossuiLancto = True
            End If

            If aCheque(IndCh).TipoDocto <> 4 And aCheque(IndCh).TipoDocto <> 41 And _
               ((Left(aCheque(IndCh).Leitura, 3) <> "409") And _
                (Left(aCheque(IndCh).Leitura, 3) <> "230")) Then
                
                PossuiChqTerceiro = True
            End If

            If aCheque(IndCh).TipoDocto = 4 Then
                PossuiADCC = True
            End If

        End If
        DoEvents
    Next Count

    'Mais de um credito com cheque de outros bancos sem lancamento -> Nao processa
    If PossuiLancto = False And lstCheques.SelCount > 1 And PossuiChqTerceiro = True Then
        'Verifica se Chq Terceiro e usuário = Coordenador ou Supervisor
        If (InStr(Geral.GrupoUsuario, "COO") = 0 And InStr(Geral.GrupoUsuario, "SUP") = 0) Then
            Exit Sub
        End If
    End If

    'Mais de um credito com ADCC sem lancamento -> Nao Processa
    If PossuiLancto = False And lstCheques.SelCount > 1 And PossuiADCC = True Then Exit Sub

    'Verificar se foi selecionado algum Lancamento
    If PossuiLancto Then
        If lstCheques.SelCount > 1 Then
            If lstPagtos.SelCount = 1 Then
                'N x 1
                'Se possui lancamento e cheque terceiro sem deposito -> Nao Processa
                If PossuiChqTerceiro = True And PossuiDeposito = False Then
                    'Verifica se Chq Terceiro e usuário = Coordenador ou Supervisor
                    If (InStr(Geral.GrupoUsuario, "COO") = 0 And InStr(Geral.GrupoUsuario, "SUP") = 0) Then
                        Exit Sub
                    End If
                End If
            Else
                'N x N
                'Se possui mais de um pagamento com Cheque de Terceiro -> Nao Processa
                If PossuiChqTerceiro Then
                    'Verifica se Chq Terceiro e usuário = Coordenador ou Supervisor
                    If (InStr(Geral.GrupoUsuario, "COO") = 0 And InStr(Geral.GrupoUsuario, "SUP") = 0) Then
                        Exit Sub
                    End If
                End If
            End If
        End If
    Else
        'Verifica se Crédito é cheque de outro banco
        If aCheque(IndCh).TipoDocto <> 4 And ((Left(aCheque(IndCh).Leitura, 3) <> "409") And _
                                              (Left(aCheque(IndCh).Leitura, 3) <> "230")) Then
            'Cheque Terceiros
            If lstPagtos.SelCount = 1 Then
                '1 x 1
                IndPg = IndicePagto(Val(Right(lstPagtos.List(lstPagtos.ListIndex), 10)))

                'Verificar se o titulo é do Bandeirantes
                If Left(aPagto(IndPg).Leitura, 3) = "230" Then
                    cmdVincular.Enabled = True
                    Exit Sub
                Else
                    'Se diferente de cobranca de terceiros
                    If Abs(m_ValDiferenca) > m_ValorAjusteContabil Then Exit Sub

                    'Verifica se Chq Terceiro e usuário = Coordenador ou Supervisor
                    If aPagto(IndPg).TipoDocto = 12 Or aPagto(IndPg).TipoDocto = 31 Then
                        If (InStr(Geral.GrupoUsuario, "COO") = 0 And InStr(Geral.GrupoUsuario, "SUP") = 0) Then
                            Exit Sub
                        End If
                    End If
                End If
            Else
                '1 x N
                'Nao processa
                'Verifica se Chq Terceiro e usuário = Coordenador ou Supervisor
                If (InStr(Geral.GrupoUsuario, "COO") = 0 And InStr(Geral.GrupoUsuario, "SUP") = 0) Then
                    Exit Sub
                End If
            End If
        Else
            'Cheque Ubb
            If lstCheques.SelCount > 1 And m_ValDiferenca <> 0 Then
                Exit Sub
            End If
        End If
    End If

    If PossuiLancto = True Then
        'Verificar se existe diferenca maior que ajuste contabil parametrizado
        If Abs(m_ValDiferenca) > (m_ValorAjusteContabil + m_ValorLimiteMaxDifLancto) Then Exit Sub
    Else
        'Verificar se existe diferenca maior que ajuste contabil parametrizado
        If Abs(m_ValDiferenca) > m_ValorLimiteMaxDifLancto Then Exit Sub
    End If
    ''''''''''''''''''''''''''''''''''''''''''
    'não é permitido depositos com pagamentos'
    ''''''''''''''''''''''''''''''''''''''''''
    If PossuiLancto = False Then
        If ExisteDeptoPagto() Then Exit Sub
    End If

    cmdVincular.Enabled = True
End Sub
Private Sub VerificaVinculoMalote()

    If aCapa(cmbCapa.ListIndex + 1).RegraNova Then
        VerificaVinculoMaloteRegraNova
    Else
        VerificaVinculoMaloteRegraAntiga
    End If
End Sub
Private Sub VinculaEnvelope()

    Dim Count As Integer
    Dim iVinculo As Long
    Dim i As Integer
    Dim j As Integer
    Dim Msg As String
    Dim Agencia As Integer
    Dim Conta As Long
    Dim TipoAjuste As Integer

    If m_ValDiferenca <> 0 And Abs(m_ValDiferenca) > m_ValorAjusteContabil Then
        'Obter conta para Credito ou Debito da diferenca
        i = IndiceCheque(Val(Right(lstCheques.List(lstCheques.ListIndex), 10)))
        If Not ObtemAgConta(aCheque(i).IdDocto, aCheque(i).TipoDocto, Agencia, Conta) Then
            Exit Sub
        End If

        Msg = "Este vínculo irá gerar um Ajuste de Crédito no valor de " & _
            Trim(FormataValor(Abs(m_ValDiferenca), 20)) & vbCrLf & "na Conta " & FormataConta(Conta) & _
            " da Agência " & Format(Agencia, "0000") & vbCrLf & "Confirma vínculo? "
        If MsgBox(Msg, vbQuestion + vbYesNo, App.Title) = vbNo Then
            Exit Sub
        End If

        If m_ValDiferenca > 0 Then
            TipoAjuste = 34
        Else
            TipoAjuste = 38
        End If

    ElseIf m_ValDiferenca <> 0 Then
        Msg = "Este vínculo irá gerar um Ajuste Contábil de " & IIf(m_ValDiferenca > 0, "Crédito", "Débito") & _
            " no valor de " & _
            Trim(FormataValor(Abs(m_ValDiferenca), 20)) & vbCrLf & "Confirma vínculo? "
        If MsgBox(Msg, vbQuestion + vbYesNo, App.Title) = vbNo Then
            Exit Sub
        End If

        If m_ValDiferenca > 0 Then
            TipoAjuste = 42
        Else
            TipoAjuste = 43
        End If
    End If

    If Abs(m_ValDiferenca) <> 0 Then
        m_Ajuste.Agencia = Agencia
        m_Ajuste.Conta = Conta
        m_Ajuste.Valor = Abs(m_ValDiferenca)
        m_Ajuste.TipoDocto = TipoAjuste
        m_Ajuste.Vinculo = 0
        m_CountAjuste = m_CountAjuste + 1
        If m_CountAjuste > 50 Then
            MsgBox "Não é possível gerar mais que 50 ajustes num mesmo Envelope.", vbInformation + vbOKOnly, App.Title
            Exit Sub
        End If
        aAjuste(m_CountAjuste) = m_Ajuste
    End If
' Se diferenca negativa envelope nao processa
        
'    ElseIf m_ValDiferenca < 0 Then
'        If Abs(m_ValDiferenca) < m_ValorAjusteVincManual_Env Then
'            Msg = "Este vínculo irá gerar um Ajuste de Débito no valor de " & _
'                Trim(FormataValor(Abs(m_ValDiferenca), 20)) & vbCrLf & "na Conta " & FormataConta(Conta) & _
'                " da Agência " & Format(Agencia, "0000") & vbCrLf & "Confirma vínculo? "
'            If MsgBox(Msg, vbQuestion + vbYesNo, App.Title) = vbNo Then
'                Exit Sub
'            End If
'
'            m_Ajuste.Agencia = Agencia
'            m_Ajuste.Conta = Conta
'            m_Ajuste.Valor = Abs(m_ValDiferenca)
'            m_Ajuste.TipoDocto = 38
'            m_Ajuste.Vinculo = 0
'            m_CountAjuste = m_CountAjuste + 1
'            If m_CountAjuste > 50 Then
'                MsgBox "Não é possível gerar mais que 50 ajustes num mesmo Envelope.", vbInformation + vbOKOnly, App.Title
'                Exit Sub
'            End If
'            aAjuste(m_CountAjuste) = m_Ajuste
'
'        Else
'            MsgBox "Não é possível efetuar o vínculo, pois existe uma diferença de " & _
'                Trim(FormataValor(Abs(m_ValDiferenca), 20)) & _
'                " que é maior ou igual ao valor máximo permitido de " & _
'                Trim(FormataValor(m_ValorAjusteVincManual_Env, 20)) & ".", _
'                vbExclamation + vbOKOnly, App.Title
'            Exit Sub
'        End If

    iVinculo = 0
    For Count = 0 To lstCheques.ListCount - 1
        If lstCheques.Selected(Count) Then
            i = IndiceCheque(Val(Right(lstCheques.List(Count), 10)))
            If iVinculo = 0 Then
                iVinculo = aCheque(i).IdDocto
            End If
            aCheque(i).Vinculo = iVinculo
            If aCheque(i).TipoDocto = 6 Or aCheque(i).TipoDocto = 7 Then
                If (Left(aCheque(i).Leitura, 3) = "409" Or Left(aCheque(i).Leitura, 3) = "230") Then
                    aCheque(i).TipoDocto = 5
                    j = IndiceDocto(Val(Right(lstCheques.List(Count), 10)))
                    aDoc(j).TipoDocto = 5
                    If aCheque(i).Valor >= IIf(m_Capa.IdEnv_Mal = "E", m_ValorAlcadaSaque_Env, m_ValorAlcadaSaque_Mal) Then
                        aCheque(i).Alcada = "S"
                        aDoc(j).Alcada = "S"
                    Else
                        aCheque(i).Alcada = "N"
                        aDoc(j).Alcada = "N"
                    End If
                    
                Else
                    ''''''''''''''''''''''''''''''''''''''''''''
                    'Cheques de terceiros agora vão para alçada'
                    '08-01-2001
                    ''''''''''''''''''''''''''''''''''''''''''''
                    aCheque(i).TipoDocto = 6
                    j = IndiceDocto(Val(Right(lstCheques.List(Count), 10)))
                    aDoc(j).TipoDocto = 6
                    If aCheque(i).Valor >= IIf(m_Capa.IdEnv_Mal = "E", m_ValorAlcadaOutros_Env, m_ValorAlcadaOutros_Mal) Then
                        aCheque(i).Alcada = "S"
                        aDoc(j).Alcada = "S"
                    Else
                        aCheque(i).Alcada = "N"
                        aDoc(j).Alcada = "N"
                    End If
'                    aCheque(i).TipoDocto = 6
'                    aCheque(i).Alcada = "N"
'                    j = IndiceDocto(Val(Right(lstCheques.List(Count), 10)))
'                    aDoc(j).TipoDocto = 6
'                    aDoc(j).Alcada = "N"
                End If
            End If
            If aCheque(i).TipoDocto = 5 And lstCheques.SelCount > 1 Then
                aCheque(i).TipoDocto = 6
                aCheque(i).Alcada = "N"
                j = IndiceDocto(Val(Right(lstCheques.List(Count), 10)))
                aDoc(j).TipoDocto = 6
                aDoc(j).Alcada = "N"
            End If
        End If
    Next

    For Count = 0 To lstPagtos.ListCount - 1
        If lstPagtos.Selected(Count) Then
            i = IndicePagto(Val(Right(lstPagtos.List(Count), 10)))
            aPagto(i).Vinculo = iVinculo
        End If
    Next
    For Count = 1 To m_CountAjuste
        If aAjuste(Count).Vinculo = 0 Then
            aAjuste(Count).Vinculo = iVinculo
        End If
    Next

    Adiciona_lstManual
    Preenche_lstCheques
    Preenche_lstPagtos
    cmdVincular.Enabled = False

    VinculaRestanteEnvelope

End Sub
Private Sub VinculaMalote()

    If aCapa(cmbCapa.ListIndex + 1).RegraNova Then
        VinculaMaloteRegraNova
    Else
        VinculaMaloteRegraAntiga
    End If
End Sub
Private Sub VinculaMaloteRegraNova()

    Dim Count       As Integer
    Dim iVinculo    As Long
    Dim i           As Integer
    Dim j           As Integer
    Dim Msg         As String
    Dim Agencia     As Integer
    Dim Conta       As Long
    Dim TipoAjuste  As Integer
    Dim DigitouAgCc As Boolean

    'Verificar se o valor da diferença é maior que o valor do ajuste contábil e menor que
    'o valor do Ajuste para Vinculo Manual
    If (Abs(m_ValDiferenca) > m_ValorAjusteContabil) And (Abs(m_ValDiferenca) <= (m_ValorLimiteMaxDifLancto + m_ValorAjusteContabil)) Then
        'Obter Conta para efetuar Ajuste de Credito / Debito
        i = IndiceCheque(Val(Right(lstCheques.List(lstCheques.ListIndex), 10)))
        If aCheque(i).TipoDocto = 41 Then
            'Lançamento Interno - Solicitar a digitação de Agencia e Conta para Crédito/Débito
            If Not DigitaAgenciaConta(Agencia, Conta) Then Exit Sub
            DigitouAgCc = True
        Else
            If aCheque(i).TipoDocto = 5 Then
                If Not ObtemAgConta(aCheque(i).IdDocto, aCheque(i).TipoDocto, Agencia, Conta) Then
                    Exit Sub
                End If
            Else
                'O usuário deve informar a agência e conta para ajuste
                If Not DigitaAgenciaConta(Agencia, Conta) Then Exit Sub
            End If
            DigitouAgCc = False
        End If

        'Verificar se é necessário gerar dois ajustes para o vínculo selecionado
        If Abs(m_ValDiferenca) > m_ValorLimiteMaxDifLancto Then
            Msg = "Este vínculo irá gerar um Ajuste no valor de " & _
                Trim(FormataValor(m_ValorLimiteMaxDifLancto, 20)) & vbCrLf & "na Conta " & FormataConta(Conta) & _
                " da Agência " & Format(Agencia, "0000") & " e um ajuste contábil no valor de " & Trim(FormataValor(Abs(m_ValDiferenca) - m_ValorLimiteMaxDifLancto, 4)) & "." & vbCrLf & "Confirma vínculo? "
            If MsgBox(Msg, vbQuestion + vbYesNo, App.Title) = vbNo Then
                Exit Sub
            Else
                'Gerar Ajuste contábil com valor que exceder o parametro
                If Not GeraAjuste(Agencia, Conta, Abs(m_ValDiferenca) - m_ValorLimiteMaxDifLancto, IIf(m_ValDiferenca > 0, 42, 43)) Then
                    MsgBox "Não foi possível gerar o(s) ajuste(s).", vbInformation + vbOKOnly, App.Title
                    Exit Sub
                End If
            End If
        Else
            Msg = "Este vínculo irá gerar um Ajuste no valor de " & _
                Trim(FormataValor(Abs(m_ValDiferenca), 20)) & vbCrLf & "na Conta " & FormataConta(Conta) & _
                " da Agência " & Format(Agencia, "0000") & vbCrLf & "Confirma vínculo? "
            If MsgBox(Msg, vbQuestion + vbYesNo, App.Title) = vbNo Then
                Exit Sub
            End If
        End If

        If m_ValDiferenca > 0 Then
            TipoAjuste = 34
        Else
            TipoAjuste = 38
        End If

    ElseIf m_ValDiferenca <> 0 And Abs(m_ValDiferenca) <= m_ValorAjusteContabil Then
        Msg = "Este vínculo irá gerar um Ajuste Contábil de " & IIf(m_ValDiferenca > 0, "Crédito", "Débito") & _
            " no valor de " & _
            Trim(FormataValor(Abs(m_ValDiferenca), 20)) & vbCrLf & "Confirma vínculo? "
        If MsgBox(Msg, vbQuestion + vbYesNo, App.Title) = vbNo Then
            Exit Sub
        End If

        If m_ValDiferenca > 0 Then
            TipoAjuste = 42
        Else
            TipoAjuste = 43
        End If
    End If

    'Verificar se já foi gerado um ajuste contábil com o excedente do parametro
    If Not GeraAjuste(Agencia, Conta, IIf(Abs(m_ValDiferenca) > m_ValorLimiteMaxDifLancto, m_ValorLimiteMaxDifLancto, Abs(m_ValDiferenca)), TipoAjuste) Then
        MsgBox "Não foi possível gerar o(s) ajuste(s).", vbInformation + vbOKOnly, App.Title
        Exit Sub
    End If

    iVinculo = 0
    For Count = 0 To lstCheques.ListCount - 1
        If lstCheques.Selected(Count) Then
            i = IndiceCheque(Val(Right(lstCheques.List(Count), 10)))
            If iVinculo = 0 Then
                iVinculo = aCheque(i).IdDocto
            End If
            aCheque(i).Vinculo = iVinculo
            If (aCheque(i).TipoDocto = 6 Or aCheque(i).TipoDocto = 7) And lstCheques.SelCount = 1 Then
                If Left(aCheque(i).Leitura, 3) = "409" Or Left(aCheque(i).Leitura, 3) = "230" Then
                    aCheque(i).TipoDocto = 5
                    j = IndiceDocto(Val(Right(lstCheques.List(Count), 10)))
                    aDoc(j).TipoDocto = 5
                    If aCheque(i).Valor >= IIf(m_Capa.IdEnv_Mal = "E", m_ValorAlcadaSaque_Env, m_ValorAlcadaSaque_Mal) Then
                        aCheque(i).Alcada = "S"
                        aDoc(j).Alcada = "S"
                    Else
                        aCheque(i).Alcada = "N"
                        aDoc(j).Alcada = "N"
                    End If
                Else
                    ''''''''''''''''''''''''''''''''''''''''''''
                    'Cheques de terceiros agora vão para alçada'
                    '08-01-2001
                    ''''''''''''''''''''''''''''''''''''''''''''
                    aCheque(i).TipoDocto = 6
                    j = IndiceDocto(Val(Right(lstCheques.List(Count), 10)))
                    aDoc(j).TipoDocto = 6
                    If aCheque(i).Valor >= IIf(m_Capa.IdEnv_Mal = "E", m_ValorAlcadaOutros_Env, m_ValorAlcadaOutros_Mal) Then
                        aCheque(i).Alcada = "S"
                        aDoc(j).Alcada = "S"
                    Else
                        aCheque(i).Alcada = "N"
                        aDoc(j).Alcada = "N"
                    End If
                    
'                    aCheque(i).TipoDocto = 6
'                    aCheque(i).Alcada = "N"
'                    j = IndiceDocto(Val(Right(lstCheques.List(Count), 10)))
'                    aDoc(j).TipoDocto = 6
'                    aDoc(j).Alcada = "N"
                End If

            ElseIf (aCheque(i).TipoDocto = 5) Then
                If lstCheques.SelCount > 1 Then
                    aCheque(i).TipoDocto = 6
                    aDoc(j).TipoDocto = 6
                End If
                
                If aCheque(i).Valor >= IIf(m_Capa.IdEnv_Mal = "E", m_ValorAlcadaSaque_Env, m_ValorAlcadaSaque_Mal) Then
                    aCheque(i).Alcada = "S"
                    aDoc(j).Alcada = "S"
                Else
                    aCheque(i).Alcada = "N"
                    aDoc(j).Alcada = "N"
                End If
            ElseIf (aCheque(i).TipoDocto = 7) Then
                If Not ExisteDeposito() Then
                    aCheque(i).TipoDocto = 6
                    aDoc(j).TipoDocto = 6
                End If
            End If
        End If
    Next

    For Count = 0 To lstPagtos.ListCount - 1
        If lstPagtos.Selected(Count) Then
            i = IndicePagto(Val(Right(lstPagtos.List(Count), 10)))
            aPagto(i).Vinculo = iVinculo
        End If
    Next

    For Count = 1 To m_CountAjuste
        If aAjuste(Count).Vinculo = 0 Then
            aAjuste(Count).Vinculo = iVinculo
        End If
    Next
    
    Adiciona_lstManual
    Preenche_lstCheques
    Preenche_lstPagtos
    cmdVincular.Enabled = False

    VinculaRestanteMalote

End Sub
Private Sub VinculaMaloteRegraAntiga()

    Dim Count As Integer
    Dim iVinculo As Long
    Dim i As Integer
    Dim j As Integer
    Dim Msg As String
    Dim Agencia As Integer
    Dim Conta As Long
    Dim strNumMalote

    strNumMalote = Format(aCapa(cmbCapa.ListIndex + 1).NumMalote, "00000000000")
    Agencia = Val(Left(strNumMalote, 4))
    Conta = Val(Right(strNumMalote, 7))

    If m_ValDiferenca > 0 Then
        Msg = "Este vínculo irá gerar um Ajuste de Crédito no valor de " & _
            Trim(FormataValor(Abs(m_ValDiferenca), 20)) & vbCrLf & "na Conta " & FormataConta(Conta) & _
            " da Agência " & Format(Agencia, "0000") & vbCrLf & "Confirma vínculo? "
        If MsgBox(Msg, vbQuestion + vbYesNo, App.Title) = vbNo Then
            Exit Sub
        End If

        m_Ajuste.Agencia = Agencia
        m_Ajuste.Conta = Conta
        m_Ajuste.Valor = Abs(m_ValDiferenca)
        m_Ajuste.TipoDocto = 34
        m_Ajuste.Vinculo = 0
        m_CountAjuste = m_CountAjuste + 1
        If m_CountAjuste > 50 Then
            MsgBox "Não é possível gerar mais que 50 ajustes num mesmo Malote.", vbInformation + vbOKOnly, App.Title
            Exit Sub
        End If
        aAjuste(m_CountAjuste) = m_Ajuste

    ElseIf m_ValDiferenca < 0 Then

        If Abs(m_ValDiferenca) < m_ValorAjusteVincManual_Mal Then
            Msg = "Este vínculo irá gerar um Ajuste de Débito no valor de " & _
                Trim(FormataValor(Abs(m_ValDiferenca), 20)) & vbCrLf & "na Conta " & FormataConta(Conta) & _
                " da Agência " & Format(Agencia, "0000") & vbCrLf & "Confirma vínculo? "
            If MsgBox(Msg, vbQuestion + vbYesNo, App.Title) = vbNo Then
                Exit Sub
            End If

            m_Ajuste.Agencia = Agencia
            m_Ajuste.Conta = Conta
            m_Ajuste.Valor = Abs(m_ValDiferenca)
            m_Ajuste.TipoDocto = 38
            m_Ajuste.Vinculo = 0
            m_CountAjuste = m_CountAjuste + 1

            If m_CountAjuste > 50 Then
                MsgBox "Não é possível gerar mais que 50 ajustes num mesmo Malote.", vbInformation + vbOKOnly, App.Title
                Exit Sub
            End If
            aAjuste(m_CountAjuste) = m_Ajuste
        Else
            MsgBox "Não é possível efetuar o vínculo, pois existe uma diferença de " & _
                Trim(FormataValor(Abs(m_ValDiferenca), 20)) & _
                " que é maior ou igual ao valor máximo permitido de " & _
                Trim(FormataValor(m_ValorAjusteVincManual_Mal, 20)) & ".", _
                vbExclamation + vbOKOnly, App.Title
            Exit Sub
        End If
    End If

    iVinculo = 0
    For Count = 0 To lstCheques.ListCount - 1
        If lstCheques.Selected(Count) Then
            i = IndiceCheque(Val(Right(lstCheques.List(Count), 10)))
            If iVinculo = 0 Then
                iVinculo = aCheque(i).IdDocto
            End If
            aCheque(i).Vinculo = iVinculo
            If aCheque(i).TipoDocto = 6 Or aCheque(i).TipoDocto = 7 Then
                If Left(aCheque(i).Leitura, 3) = "409" Or Left(aCheque(i).Leitura, 3) = "230" Then
                    aCheque(i).TipoDocto = 5
                    j = IndiceDocto(Val(Right(lstCheques.List(Count), 10)))
                    aDoc(j).TipoDocto = 5
                    If aCheque(i).Valor >= IIf(m_Capa.IdEnv_Mal = "E", m_ValorAlcadaSaque_Env, m_ValorAlcadaSaque_Mal) Then
                        aCheque(i).Alcada = "S"
                        aDoc(j).Alcada = "S"
                    Else
                        aCheque(i).Alcada = "N"
                        aDoc(j).Alcada = "N"
                    End If
                Else
                    ''''''''''''''''''''''''''''''''''''''''''''
                    'Cheques de terceiros agora vão para alçada'
                    '08-01-2001
                    ''''''''''''''''''''''''''''''''''''''''''''
                    aCheque(i).TipoDocto = 6
                    j = IndiceDocto(Val(Right(lstCheques.List(Count), 10)))
                    aDoc(j).TipoDocto = 6
                    If aCheque(i).Valor >= IIf(m_Capa.IdEnv_Mal = "E", m_ValorAlcadaOutros_Env, m_ValorAlcadaOutros_Mal) Then
                        aCheque(i).Alcada = "S"
                        aDoc(j).Alcada = "S"
                    Else
                        aCheque(i).Alcada = "N"
                        aDoc(j).Alcada = "N"
                    End If
                    
                End If
            End If
            
'            If aCheque(i).TipoDocto = 5 And lstCheques.SelCount > 1 Then
'                aCheque(i).TipoDocto = 6
'                aCheque(i).Alcada = "N"
'                j = IndiceDocto(Val(Right(lstCheques.List(Count), 10)))
'                aDoc(j).TipoDocto = 6
'                aDoc(j).Alcada = "N"
            
            If aCheque(i).TipoDocto = 5 Then
                j = IndiceDocto(Val(Right(lstCheques.List(Count), 10)))
                
                If lstCheques.SelCount > 1 Then
                    aCheque(i).TipoDocto = 6
                    aDoc(j).TipoDocto = 6
                End If
                
                If aCheque(i).Valor >= IIf(m_Capa.IdEnv_Mal = "E", m_ValorAlcadaSaque_Env, m_ValorAlcadaSaque_Mal) Then
                    aCheque(i).Alcada = "S"
                    aDoc(j).Alcada = "S"
                Else
                    aCheque(i).Alcada = "N"
                    aDoc(j).Alcada = "N"
                End If

            
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Enviar pagamentos com lançamento interno para alçada'
            '09-01-2001                                          '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ElseIf aCheque(i).TipoDocto = 41 And lstCheques.SelCount = 1 Then
                j = IndiceDocto(Val(Right(lstCheques.List(Count), 10)))
                If aCheque(i).Valor >= IIf(m_Capa.IdEnv_Mal = "E", m_ValorAlcadaOutros_Env, m_ValorAlcadaOutros_Mal) Then
                    aCheque(i).Alcada = "S"
                    aDoc(j).Alcada = "S"
                Else
                    aCheque(i).Alcada = "N"
                    aDoc(j).Alcada = "N"
                End If
            End If
        End If
    Next
    For Count = 0 To lstPagtos.ListCount - 1
        If lstPagtos.Selected(Count) Then
            i = IndicePagto(Val(Right(lstPagtos.List(Count), 10)))
            aPagto(i).Vinculo = iVinculo
        End If
    Next
    For Count = 1 To m_CountAjuste
        If aAjuste(Count).Vinculo = 0 Then
            aAjuste(Count).Vinculo = iVinculo
        End If
    Next

    Adiciona_lstManual
    Preenche_lstCheques
    Preenche_lstPagtos
    cmdVincular.Enabled = False

    VinculaRestanteMalote

End Sub
Private Sub VinculaRestanteEnvelope()

    Dim Count As Integer
    Dim i As Integer
    Dim j As Integer
    Dim iVinculo As Long

    If lstCheques.ListCount <> 1 Or lstPagtos.ListCount = 0 Then
        Exit Sub
    End If

    i = IndiceCheque(Val(Right(lstCheques.List(0), 10)))
    If ((Left(aCheque(i).Leitura, 3) <> "409") And (Left(aCheque(i).Leitura, 3) <> "230")) And _
        lstPagtos.ListCount > 1 Then
        
        Exit Sub
    End If

    lstCheques.Selected(0) = True
    For Count = 0 To lstPagtos.ListCount - 1
        lstPagtos.Selected(Count) = True
    Next

    j = IndicePagto(Val(Right(lstPagtos.List(lstPagtos.ListIndex), 10)))
    If aCheque(i).TipoDocto <> 4 And ((Left(aCheque(i).Leitura, 3) <> "409") And _
                                      (Left(aCheque(i).Leitura, 3) <> "230")) And _
       (aPagto(j).TipoDocto = 12 Or aPagto(j).TipoDocto = 31) Then
       
       Exit Sub
    End If

    If cmdVincular.Enabled = False Then Exit Sub

    CalculaValores
    
    If Abs(m_ValDiferenca) = 0 Then
        i = IndiceCheque(Val(Right(lstCheques.List(0), 10)))
        iVinculo = aCheque(i).IdDocto
        aCheque(i).Vinculo = iVinculo
        If aCheque(i).TipoDocto = 6 Or aCheque(i).TipoDocto = 7 Then
            If (Left(aCheque(i).Leitura, 3) = "409") Or (Left(aCheque(i).Leitura, 3) = "230") Then
                aCheque(i).TipoDocto = 5
                j = IndiceDocto(Val(Right(lstCheques.List(Count), 10)))
                aDoc(j).TipoDocto = 5
                If aCheque(i).Valor >= IIf(m_Capa.IdEnv_Mal = "E", m_ValorAlcadaSaque_Env, m_ValorAlcadaSaque_Mal) Then
                    aCheque(i).Alcada = "S"
                    aDoc(j).Alcada = "S"
                Else
                    aCheque(i).Alcada = "N"
                    aDoc(j).Alcada = "N"
                End If
            Else
                If aCheque(i).Valor >= IIf(m_Capa.IdEnv_Mal = "E", m_ValorAlcadaOutros_Env, m_ValorAlcadaOutros_Mal) Then
                    aCheque(i).Alcada = "S"
                    aDoc(j).Alcada = "S"
                Else
                    aCheque(i).Alcada = "N"
                    aDoc(j).Alcada = "N"
                End If
            End If
        End If
        
        For Count = 0 To lstPagtos.ListCount - 1
            If lstPagtos.Selected(Count) Then
                i = IndicePagto(Val(Right(lstPagtos.List(Count), 10)))
                aPagto(i).Vinculo = iVinculo
            End If
        Next
        Adiciona_lstManual
        Preenche_lstCheques
        Preenche_lstPagtos
        cmdVincular.Enabled = False
    End If
    
End Sub

Private Sub VinculaRestanteMalote()

    If aCapa(cmbCapa.ListIndex + 1).RegraNova Then
        VinculaRestanteMaloteRegraNova
    Else
        VinculaRestanteMaloteRegraAntiga
    End If
End Sub
Private Sub VinculaRestanteMaloteRegraNova()

    Dim Count As Integer
    Dim i As Integer
    Dim j As Integer

    If (lstCheques.ListCount = 1 And lstPagtos.ListCount = 1) Then
        lstCheques.Selected(0) = True
        lstPagtos.Selected(0) = True

        CalculaValores

        i = IndiceCheque(Val(Right(lstCheques.List(0), 10)))
        j = IndicePagto(Val(Right(lstPagtos.List(0), 10)))

        If cmdVincular.Enabled = False Then Exit Sub

        If aCheque(i).TipoDocto <> 4 And ((Left(aCheque(i).Leitura, 3) <> "409") And _
                                          (Left(aCheque(i).Leitura, 3) <> "230")) And _
           (aPagto(j).TipoDocto = 12 Or aPagto(j).TipoDocto = 31) Then
           Exit Sub
        End If

        If Abs(m_ValDiferenca) <> 0 Then
            Exit Sub
        End If

        iVinculo = aCheque(i).IdDocto
        aCheque(i).Vinculo = iVinculo
        If aCheque(i).TipoDocto = 6 Or aCheque(i).TipoDocto = 7 Then
            If (Left(aCheque(i).Leitura, 3) = "409") Or (Left(aCheque(i).Leitura, 3) = "230") Then
                aCheque(i).TipoDocto = 5
                j = IndiceDocto(Val(Right(lstCheques.List(Count), 10)))
                aDoc(j).TipoDocto = 5
                If aCheque(i).Valor >= IIf(m_Capa.IdEnv_Mal = "E", m_ValorAlcadaSaque_Env, m_ValorAlcadaSaque_Mal) Then
                    aCheque(i).Alcada = "S"
                    aDoc(j).Alcada = "S"
                Else
                    aCheque(i).Alcada = "N"
                    aDoc(j).Alcada = "N"
                End If
            Else
                aCheque(i).TipoDocto = 6
                aCheque(i).Alcada = "N"
                j = IndiceDocto(Val(Right(lstCheques.List(Count), 10)))
                aDoc(j).TipoDocto = 6
                aDoc(j).Alcada = "N"
            End If
        End If

        i = IndicePagto(Val(Right(lstPagtos.List(0), 10)))
        aPagto(i).Vinculo = iVinculo

        Adiciona_lstManual
        Preenche_lstCheques
        Preenche_lstPagtos
        cmdVincular.Enabled = False

    ElseIf (lstCheques.ListCount = 1 And lstPagtos.ListCount > 1) Then
        lstCheques.Selected(0) = True
        For Count = 0 To lstPagtos.ListCount - 1
            lstPagtos.Selected(Count) = True
        Next
        CalculaValores

        i = IndiceCheque(Val(Right(lstCheques.List(0), 10)))

        If cmdVincular.Enabled = False Then Exit Sub

        If aCheque(i).TipoDocto <> 4 And ((Left(aCheque(i).Leitura, 3) <> "409") And _
                                          (Left(aCheque(i).Leitura, 3) <> "230")) Then
            Exit Sub
        End If

        If Abs(m_ValDiferenca) = 0 Then
            iVinculo = aCheque(i).IdDocto
            aCheque(i).Vinculo = iVinculo
            If aCheque(i).TipoDocto = 6 Or aCheque(i).TipoDocto = 7 Then
                If (Left(aCheque(i).Leitura, 3) = "409") Or (Left(aCheque(i).Leitura, 3) = "230") Then
                    aCheque(i).TipoDocto = 5
                    j = IndiceDocto(Val(Right(lstCheques.List(Count), 10)))
                    aDoc(j).TipoDocto = 5
                    If aCheque(i).Valor >= IIf(m_Capa.IdEnv_Mal = "E", m_ValorAlcadaSaque_Env, m_ValorAlcadaSaque_Mal) Then
                        aCheque(i).Alcada = "S"
                        aDoc(j).Alcada = "S"
                    Else
                        aCheque(i).Alcada = "N"
                        aDoc(j).Alcada = "N"
                    End If
                Else
                    aCheque(i).TipoDocto = 6
                    aCheque(i).Alcada = "N"
                    j = IndiceDocto(Val(Right(lstCheques.List(Count), 10)))
                    aDoc(j).TipoDocto = 6
                    aDoc(j).Alcada = "N"
                End If
            End If
            
            For Count = 0 To lstPagtos.ListCount - 1
                If lstPagtos.Selected(Count) Then
                    i = IndicePagto(Val(Right(lstPagtos.List(Count), 10)))
                    aPagto(i).Vinculo = iVinculo
                End If
            Next
            Adiciona_lstManual
            Preenche_lstCheques
            Preenche_lstPagtos
            cmdVincular.Enabled = False
        End If
    ElseIf (lstPagtos.ListCount = 1 And lstCheques.ListCount > 1) Then
        lstPagtos.Selected(0) = True
        For Count = 0 To lstCheques.ListCount - 1
            i = IndiceCheque(Val(Right(lstCheques.List(Count), 10)))
            If (aCheque(i).TipoDocto <> 4 And ((Left(aCheque(i).Leitura, 3) <> "409") And _
                                               (Left(aCheque(i).Leitura, 3) <> "230"))) Or aCheque(i).TipoDocto = 4 Then
                Exit Sub
            End If
            lstCheques.Selected(Count) = True
        Next
        CalculaValores

        If cmdVincular.Enabled = False Then Exit Sub

        If Abs(m_ValDiferenca) = 0 Then
            i = IndicePagto(Val(Right(lstPagtos.List(0), 10)))
            iVinculo = aPagto(i).IdDocto
            aPagto(i).Vinculo = iVinculo
            
            For Count = 0 To lstCheques.ListCount - 1
                If lstCheques.Selected(Count) Then
                    i = IndiceCheque(Val(Right(lstCheques.List(Count), 10)))
                    aCheque(i).Vinculo = iVinculo
                    If aCheque(i).TipoDocto = 6 Or aCheque(i).TipoDocto = 7 Then
                        If (Left(aCheque(i).Leitura, 3) = "409") Or (Left(aCheque(i).Leitura, 3) = "230") Then
                            aCheque(i).TipoDocto = 5
                            j = IndiceDocto(Val(Right(lstCheques.List(Count), 10)))
                            aDoc(j).TipoDocto = 5
                            If aCheque(i).Valor >= IIf(m_Capa.IdEnv_Mal = "E", m_ValorAlcadaSaque_Env, m_ValorAlcadaSaque_Mal) Then
                                aCheque(i).Alcada = "S"
                                aDoc(j).Alcada = "S"
                            Else
                                aCheque(i).Alcada = "N"
                                aDoc(j).Alcada = "N"
                            End If
                        Else
                            aCheque(i).TipoDocto = 6
                            aCheque(i).Alcada = "N"
                            j = IndiceDocto(Val(Right(lstCheques.List(Count), 10)))
                            aDoc(j).TipoDocto = 6
                            aDoc(j).Alcada = "N"
                        End If
                    End If
                End If
            Next
            Adiciona_lstManual
            Preenche_lstCheques
            Preenche_lstPagtos
            cmdVincular.Enabled = False
        End If
    End If
End Sub
Private Sub VinculaRestanteMaloteRegraAntiga()

    Dim Count As Integer
    Dim i As Integer
    Dim j As Integer
    Dim iVinculo As Long

    If (lstCheques.ListCount = 1 And lstPagtos.ListCount > 0) Then
        lstCheques.Selected(0) = True
        For Count = 0 To lstPagtos.ListCount - 1
            lstPagtos.Selected(Count) = True
        Next

        CalculaValores

        If cmdVincular.Enabled = False Then Exit Sub

        If Abs(m_ValDiferenca) = 0 Then
            i = IndiceCheque(Val(Right(lstCheques.List(0), 10)))
            iVinculo = aCheque(i).IdDocto
            aCheque(i).Vinculo = iVinculo
            If aCheque(i).TipoDocto = 6 Or aCheque(i).TipoDocto = 7 Then
                If (Left(aCheque(i).Leitura, 3) = "409") Or (Left(aCheque(i).Leitura, 3) = "230") Then
                    aCheque(i).TipoDocto = 5
                    j = IndiceDocto(Val(Right(lstCheques.List(Count), 10)))
                    aDoc(j).TipoDocto = 5
                    If aCheque(i).Valor >= IIf(m_Capa.IdEnv_Mal = "E", m_ValorAlcadaSaque_Env, m_ValorAlcadaSaque_Mal) Then
                        aCheque(i).Alcada = "S"
                        aDoc(j).Alcada = "S"
                    Else
                        aCheque(i).Alcada = "N"
                        aDoc(j).Alcada = "N"
                    End If
                Else
                    aCheque(i).TipoDocto = 6
                    aCheque(i).Alcada = "N"
                    j = IndiceDocto(Val(Right(lstCheques.List(Count), 10)))
                    aDoc(j).TipoDocto = 6
                    aDoc(j).Alcada = "N"
                End If
            End If

            For Count = 0 To lstPagtos.ListCount - 1
                If lstPagtos.Selected(Count) Then
                    i = IndicePagto(Val(Right(lstPagtos.List(Count), 10)))
                    aPagto(i).Vinculo = iVinculo
                End If
            Next
            Adiciona_lstManual
            Preenche_lstCheques
            Preenche_lstPagtos
            cmdVincular.Enabled = False
        End If
    ElseIf (lstPagtos.ListCount = 1 And lstCheques.ListCount > 0) Then
        lstPagtos.Selected(0) = True
        For Count = 0 To lstCheques.ListCount - 1
            lstCheques.Selected(Count) = True
        Next
        CalculaValores

        If cmdVincular.Enabled = False Then Exit Sub

        If Abs(m_ValDiferenca) = 0 Then
            i = IndicePagto(Val(Right(lstPagtos.List(0), 10)))
            iVinculo = aPagto(i).IdDocto
            aPagto(i).Vinculo = iVinculo
            
            For Count = 0 To lstCheques.ListCount - 1
                If lstCheques.Selected(Count) Then
                    i = IndiceCheque(Val(Right(lstCheques.List(Count), 10)))
                    aCheque(i).Vinculo = iVinculo
                    If aCheque(i).TipoDocto = 6 Or aCheque(i).TipoDocto = 7 Then
                        If (Left(aCheque(i).Leitura, 3) = "409") Or (Left(aCheque(i).Leitura, 3) = "230") Then
                            aCheque(i).TipoDocto = 5
                            j = IndiceDocto(Val(Right(lstCheques.List(Count), 10)))
                            aDoc(j).TipoDocto = 5
                            If aCheque(i).Valor >= IIf(m_Capa.IdEnv_Mal = "E", m_ValorAlcadaSaque_Env, m_ValorAlcadaSaque_Mal) Then
                                aCheque(i).Alcada = "S"
                                aDoc(j).Alcada = "S"
                            Else
                                aCheque(i).Alcada = "N"
                                aDoc(j).Alcada = "N"
                            End If
                        Else
                            aCheque(i).TipoDocto = 6
                            aCheque(i).Alcada = "N"
                            j = IndiceDocto(Val(Right(lstCheques.List(Count), 10)))
                            aDoc(j).TipoDocto = 6
                            aDoc(j).Alcada = "N"
                        End If
                    End If
                    If aCheque(i).TipoDocto = 5 And lstCheques.SelCount > 1 Then
                        aCheque(i).TipoDocto = 6
                        aCheque(i).Alcada = "N"
                        j = IndiceDocto(Val(Right(lstCheques.List(Count), 10)))
                        aDoc(j).TipoDocto = 6
                        aDoc(j).Alcada = "N"
                    End If
                End If
            Next
            Adiciona_lstManual
            Preenche_lstCheques
            Preenche_lstPagtos
            cmdVincular.Enabled = False
        End If
    End If
End Sub
Private Function PodeEncerrar() As Boolean
    If lstCheques.ListCount = 0 And lstPagtos.ListCount = 0 Then
        PodeEncerrar = True
    Else
        PodeEncerrar = False
    End If
End Function

Private Sub HabilitaTimerPesquisa()

  'Esta Função irá verificar a existência de documentos Ilegíveis a cada x segundos
  'de acordo com o campo PARAMETRO.TmAtualizacao
  FrmPesquisa.Visible = True
  tmrPesquisa.Enabled = True
  Progress.Value = 0
  
  ''''''''''''''''''''''''''''''''''''''''''
  'Grava log MDI - Inicio Aguarda documento'
  ''''''''''''''''''''''''''''''''''''''''''
  Call GravaLog(0, 0, 256)

  
End Sub

Private Sub cmbCapa_Click()

    If m_Busy Then
        Exit Sub
    End If
    m_Busy = True

    If m_IdCapa > 0 Then
        If PodeEncerrar Then
            If MsgBox("O Envelope/Malote atual não foi encerrado. " & vbCrLf & _
                    "Deseja efetivar os Vínculos Manuais?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                m_Busy = False
                Confirmar True
                Exit Sub
            Else
                If Not AtualizaStatusCapa(m_IdCapa, "7") Then
                    m_Busy = False
                    m_IdCapa = 0
                    Exit Sub
                End If
                ' GravaLog m_IdCapa, 0, 196
            End If
        Else
            If Not AtualizaStatusCapa(m_IdCapa, "7") Then
                m_Busy = False
                m_IdCapa = 0
                Exit Sub
            End If
            ' GravaLog m_IdCapa, 0, 196
        End If
    End If

    'Verificar se a capa possui documentos já transmitidos
    If VerificaDocumentosTransmitidos Then
        m_IdCapa = 0
        m_Busy = False
        tabVinculo.Tab = 0
        lstCheques.Clear
        lstPagtos.Clear
        lstManual.Clear
        lstAutomatico.Clear
        lstOcorrencia.Clear
        LimparValores
        cmbCapa.SetFocus
        Exit Sub
    End If

    If aCapa(cmbCapa.ListIndex + 1).IdEnv_Mal = "E" Then
        lblCapa.Caption = "Envelope"
        picNumMalote.Visible = False
        lblNumMalote.Visible = False
    Else
        lblCapa.Caption = "Malote"
        picNumMalote.Visible = True
        lblNumMalote.Visible = True
        lblNumMalote.Caption = aCapa(cmbCapa.ListIndex + 1).NumMalote
    End If

    If m_CountCapa > 0 Then
        m_IdCapa = aCapa(cmbCapa.ListIndex + 1).IdCapa
    End If

    If Not VerificaCapaDisponivel(m_IdCapa) Then
        m_IdCapa = 0
        m_Busy = False
        tabVinculo.Tab = 0
        lstCheques.Clear
        lstPagtos.Clear
        lstManual.Clear
        lstAutomatico.Clear
        lstOcorrencia.Clear
        LimparValores
        Exit Sub
    End If
    
    If Not AtualizaStatusCapa(m_IdCapa, "J") Then
        m_IdCapa = 0
        m_Busy = False
        Exit Sub
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''
    'Grava acao 195 - Vinculo Manual Selecionar Capa'
    '''''''''''''''''''''''''''''''''''''''''''''''''
    GravaLog m_IdCapa, 0, 195
    
    lblLote.Caption = Format(aCapa(cmbCapa.ListIndex + 1).IdLote, "0000-00000")
    
    ObtemDocumentos m_IdCapa
    
    sTempo = 0
    
    tabVinculo.Tab = 0
    Preenche_lstCheques
    Preenche_lstPagtos
    Preenche_lstAutomatico
    Preenche_lstOcorrencia
    lstManual.Clear
    CalculaValores
    
    cmdOcorrencia.Enabled = False
    cmdCancelar.Enabled = False
    cmdVincular.Enabled = False
    LimparImagem
    
    m_Busy = False
End Sub

Private Sub CmdAtualizar_Click()
    If m_IdCapa > 0 Then
        If PodeEncerrar Then
            If MsgBox("O Envelope/Malote atual não foi encerrado. " & vbCrLf & _
                    "Deseja efetivar os Vínculos Manuais?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                Confirmar False
            Else
                If Not AtualizaStatusCapa(m_IdCapa, "7") Then
                    Exit Sub
                End If
                ' GravaLog m_IdCapa, 0, 196
            End If
        Else
            If Not AtualizaStatusCapa(m_IdCapa, "7") Then
                Exit Sub
            End If
            ' GravaLog m_IdCapa, 0, 196
        End If
        m_IdCapa = 0
    End If

    LimparValores
    LimparHeader
    LimparListas
    
    If Not ObtemCapas Then
        MsgBox "Não existem Envelopes/Malotes com pendência de Vínculo Manual.", vbExclamation + vbOKOnly, App.Title
        m_IdCapa = 0
        LimparImagem
        HabilitaTimerPesquisa
        Exit Sub
    Else
      FrmPesquisa.Visible = False
      tmrPesquisa.Enabled = False
    End If

    Preenche_cmbCapa
    cmbCapa.ListIndex = 0
End Sub

Private Sub cmdAuditoria_Click()

    Geral.Capa.IdCapa = aCapa(cmbCapa.ListIndex + 1).IdCapa
    Geral.Capa.Capa = Val(aCapa(cmbCapa.ListIndex + 1).Capa)
    Geral.Capa.Num_Malote = Val(aCapa(cmbCapa.ListIndex + 1).NumMalote)
    Geral.Capa.AgOrig = aCapa(cmbCapa.ListIndex + 1).AgOrig
    Geral.Capa.IdEnv_Mal = aCapa(cmbCapa.ListIndex + 1).IdEnv_Mal
    
    Call Auditoria
    
    Geral.Capa.IdCapa = 0
    Geral.Capa.Capa = 0
    Geral.Capa.Num_Malote = 0
    Geral.Capa.AgOrig = 0
    Geral.Capa.IdEnv_Mal = ""

End Sub

Private Sub cmdCancelar_Click()
    Dim Count As Integer
    
    For Count = 0 To lstCheques.ListCount - 1
        lstCheques.Selected(Count) = False
    Next
    For Count = 0 To lstPagtos.ListCount - 1
        lstPagtos.Selected(Count) = False
    Next
    CalculaValores
    LimparImagem
End Sub

Private Sub Confirmar(ByVal Atualizar As Boolean)

    Dim Count As Integer
    Dim bAlcada As Boolean
    Dim strAutenticacaoDigital  As String
    
    If m_Busy Then
        Exit Sub
    End If

    m_Busy = True

    If Not PodeEncerrar Then
        MsgBox "Não é possível encerrar Envelope/Malote sem que todos documentos " & _
            "tenham sido vinculados ou devolvidos.", vbExclamation + vbOKOnly, App.Title
        tabVinculo.Tab = 0
        lstCheques.SetFocus
        m_Busy = False
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass

    bAlcada = False

    For Count = 1 To m_CountDocto
    
        If Count <= m_CountCheque Then
            If aCheque(Count).Status = "D" Then
                On Error GoTo ErroExclusao
                With qryAtualizaDocumentoExcluido
                    .rdoParameters(0).Direction = rdParamReturnValue
                    .rdoParameters(1) = Geral.DataProcessamento
                    .rdoParameters(2) = aCheque(Count).IdDocto
                    .rdoParameters(3) = aCheque(Count).Status
                    .rdoParameters(4) = 0 ' duplicidade
                    .rdoParameters(5) = aCheque(Count).Ocorrencia
                    .Execute
                    If .rdoParameters(0) <> 0 Then
                        GoTo ErroExclusao
                    End If
                End With
                'Grava/Altera ou Exclui Complemento da Ocorrência
'''                Call GravaComplementoOcorrencia(aCheque(Count).IdDocto, IIf(aCheque(Count).ComplementoOcorrencia = "", "E", "G"), aCheque(Count).ComplementoOcorrencia)
                
                On Error GoTo 0
                GravaLog m_IdCapa, aCheque(Count).IdDocto, 71
                aDoc(IndiceDocto(aCheque(Count).IdDocto)).Status = "D"
            Else
                On Error GoTo ErroVinculo
                With qryAtualizaVinculoDocumento
                    .rdoParameters(0).Direction = rdParamReturnValue
                    .rdoParameters(1) = Geral.DataProcessamento
                    .rdoParameters(2) = aCheque(Count).IdDocto
                    .rdoParameters(3) = aCheque(Count).Vinculo
                    .rdoParameters(4) = aCheque(Count).TipoDocto
                    .rdoParameters(5) = aCheque(Count).Alcada
                    .Execute
                    If .rdoParameters(0) <> 0 Then
                        GoTo ErroVinculo
                    End If
                End With
                On Error GoTo 0
                GravaLog m_IdCapa, aCheque(Count).IdDocto, 72
            End If
        End If
        
        If Count <= m_CountPagto Then
            If aPagto(Count).Status = "D" Then
                On Error GoTo ErroExclusao
                With qryAtualizaDocumentoExcluido
                    .rdoParameters(0).Direction = rdParamReturnValue
                    .rdoParameters(1) = Geral.DataProcessamento
                    .rdoParameters(2) = aPagto(Count).IdDocto
                    .rdoParameters(3) = aPagto(Count).Status
                    .rdoParameters(4) = 0 ' duplicidade
                    .rdoParameters(5) = aPagto(Count).Ocorrencia
                    .Execute
                    If .rdoParameters(0) <> 0 Then
                        GoTo ErroExclusao
                    End If
                End With
                'Grava/Altera ou Exclui Complemento da Ocorrência
'''                Call GravaComplementoOcorrencia(aPagto(Count).IdDocto, IIf(aPagto(Count).ComplementoOcorrencia = "", "E", "G"), aPagto(Count).ComplementoOcorrencia)
                
                On Error GoTo 0
                GravaLog m_IdCapa, aPagto(Count).IdDocto, 71
                aDoc(IndicePagto(aPagto(Count).IdDocto)).Status = "D"
            Else
                On Error GoTo ErroVinculo
                With qryAtualizaVinculoDocumento
                    .rdoParameters(0).Direction = rdParamReturnValue
                    .rdoParameters(1) = Geral.DataProcessamento
                    .rdoParameters(2) = aPagto(Count).IdDocto
                    .rdoParameters(3) = aPagto(Count).Vinculo
                    .rdoParameters(4) = aPagto(Count).TipoDocto
                    .rdoParameters(5) = aPagto(Count).Alcada
                    .Execute
                    If .rdoParameters(0) <> 0 Then
                        GoTo ErroVinculo
                    End If
                End With
                On Error GoTo 0
                GravaLog m_IdCapa, aPagto(Count).IdDocto, 72
            End If
        End If
        
        If Count <= m_CountAjuste Then
            On Error GoTo ErroAjuste

            'Gera Autenticação Digital
            If aAjuste(Count).TipoDocto = 32 Or aAjuste(Count).TipoDocto = 33 Or _
                aAjuste(Count).TipoDocto = 34 Or aAjuste(Count).TipoDocto = 38 Then
                
                strAutenticacaoDigital = G_EncriptaBO(aAjuste(Count).TipoDocto, CStr(aAjuste(Count).Conta))
                
                If strAutenticacaoDigital = "" Then GoTo ErroAjuste
            Else
                strAutenticacaoDigital = ""
            End If
            
            'Verificar qual o ultimo numero de ordem de captura e incrementar 1
            qryGetUltimaOrdemCaptura.rdoParameters(0).Value = m_IdCapa
            qryGetUltimaOrdemCaptura.rdoParameters(1).Value = Geral.DataProcessamento

            Set rsOrdemCaptura = qryGetUltimaOrdemCaptura.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

            With qryInsereAjuste
                .rdoParameters(0).Direction = rdParamReturnValue
                .rdoParameters(1) = Geral.DataProcessamento
                .rdoParameters(2) = m_IdCapa
                .rdoParameters(3) = aAjuste(Count).TipoDocto
                .rdoParameters(4) = aAjuste(Count).Agencia
                .rdoParameters(5) = aAjuste(Count).Conta
                .rdoParameters(6) = aAjuste(Count).Valor
                .rdoParameters(7) = aAjuste(Count).Vinculo
                .rdoParameters(8) = Val(rsOrdemCaptura!MaiorOrdem) + 1
                .rdoParameters(9) = strAutenticacaoDigital
                .Execute

                If .rdoParameters(0) <> 0 Then
                    GoTo ErroAjuste
                End If
            End With

            On Error GoTo 0
            GravaLog m_IdCapa, 0, 73
        End If
        
        If aDoc(Count).Alcada = "S" And aDoc(Count).Status <> "D" Then
            bAlcada = True
        End If
        
    Next
    
    If bAlcada Then
        If Not AtualizaStatusCapa(m_IdCapa, "6") Then
            m_Busy = False
            Unload Me
            Exit Sub
        End If
        GravaLog m_IdCapa, 0, 74
    Else
        If Not AtualizaStatusCapa(m_IdCapa, "R") Then
            m_Busy = False
            Unload Me
            Exit Sub
        End If
        GravaLog m_IdCapa, 0, 75
    End If
    
    Screen.MousePointer = vbDefault
    m_Busy = False
    m_IdCapa = 0
    
    If Atualizar Then
        CmdAtualizar_Click
    End If
    
    Exit Sub
    
ErroExclusao:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro na atualização do status do documento.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    m_Busy = False
    cmdDesfazer_Click
    Unload Me
    Exit Sub
    
ErroVinculo:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro na atualização do vínculo do documento.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    m_Busy = False
    cmdDesfazer_Click
    Unload Me
    Exit Sub

ErroAjuste:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro na inserção de ajuste de credito/debito.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    m_Busy = False
    cmdDesfazer_Click
    Unload Me
    Exit Sub
End Sub

Private Sub cmdConfirmar_Click()
    If FrmPesquisa.Visible = True Then Exit Sub
    Confirmar True
End Sub

Private Sub cmdDesfazer_Click()

    Dim Count As Integer

    If FrmPesquisa.Visible = True Then Exit Sub

    For Count = 1 To m_CountCheque
        aCheque(Count).Vinculo = 0
        aCheque(Count).Status = aCheque(Count).StatusAnt
        aCheque(Count).TipoDocto = aCheque(Count).TipoDoctoAnt
        aCheque(Count).Alcada = aCheque(Count).AlcadaAnt
        aCheque(Count).Ocorrencia = 0
    Next
    For Count = 1 To m_CountPagto
        aPagto(Count).Vinculo = 0
        aPagto(Count).Status = aPagto(Count).StatusAnt
        aPagto(Count).TipoDocto = aPagto(Count).TipoDoctoAnt
        aPagto(Count).Alcada = aPagto(Count).AlcadaAnt
        aPagto(Count).Ocorrencia = 0
    Next
    Erase aAjuste
    m_CountAjuste = 0

    tabVinculo.Tab = 0
    Preenche_lstCheques
    Preenche_lstPagtos
    Preenche_lstOcorrencia
    lstManual.Clear
    CalculaValores
End Sub
Private Sub cmdDesfazerTudo_Click()

Dim Count As Integer
Dim avinculo()  As Long
Dim Vinculo As Long
Dim bVerificado As Boolean
Dim iVinculo As Integer
Dim x As Integer, iCountLI As Integer, iCountDP As Integer


    If m_Busy Then
        Exit Sub
    End If

    m_Busy = True

    Screen.MousePointer = vbHourglass

    On Error GoTo ErroVinculo

    Vinculo = 0
    ReDim avinculo(0)

    For x = 1 To m_CountDocto
        If aDoc(x).TipoDocto <> 1 And aDoc(x).Vinculo <> 0 Then

            If Vinculo <> aDoc(x).Vinculo Then
                Vinculo = aDoc(x).Vinculo
                
                'Verifica se vinculo já sofreu verificação
                bVerificado = False
                For iVinculo = 1 To UBound(avinculo)
                    If avinculo(iVinculo) = Vinculo Then
                        bVerificado = True
                        Exit For
                    End If
                Next
                
                If Not bVerificado Then
                    ReDim Preserve avinculo(UBound(avinculo) + 1)
                    avinculo(UBound(avinculo)) = Vinculo
                End If
                
                If Not bVerificado Then
                    iCountLI = 0
                    iCountDP = 0
                    
                    'Verifica quantos (Depósitos e LI) existem por vínculo e (OCT)
                    For Count = 1 To m_CountDocto
                        If aDoc(Count).Vinculo = Vinculo Then
                            If aDoc(Count).TipoDocto = 2 Or aDoc(Count).TipoDocto = 3 Then
                                iCountDP = iCountDP + 1
                            End If
                            If aDoc(Count).TipoDocto = 41 Then
                                iCountLI = iCountLI + 1
                            End If
                            If aDoc(Count).TipoDocto = 37 Or aDoc(Count).TipoDocto = 39 Then
                                iCountDP = 1
                                iCountLI = 0
                                Exit For
                            End If
                        End If
                    Next
                    
                    'Não desfazer o vínculo para (Depósito sem LI) e (OCT)
                    If Not (iCountDP >= 1 And iCountLI = 0) Then
                        For Count = 1 To m_CountDocto
                            If aDoc(Count).Vinculo = Vinculo Then
'                                If aDoc(Count).Status <> "D" And aDoc(Count).Status <> "F" And _
'                                   aDoc(Count).TipoDocto <> 2 And aDoc(Count).TipoDocto <> 3 And _
'                                   aDoc(Count).TipoDocto <> 7 And aDoc(Count).TipoDocto <> 32 And _
'                                   aDoc(Count).TipoDocto <> 33 And aDoc(Count).TipoDocto <> 37 And _
'                                   aDoc(Count).TipoDocto <> 39 Then
                                If aDoc(Count).TipoDocto <> 32 And aDoc(Count).TipoDocto <> 33 And _
                                   aDoc(Count).TipoDocto <> 34 And aDoc(Count).TipoDocto <> 38 And _
                                   aDoc(Count).TipoDocto <> 42 And aDoc(Count).TipoDocto <> 43 Then
                                    
                                    With qryAtualizaVinculoDocumento
                                        .rdoParameters(0).Direction = rdParamReturnValue
                                        .rdoParameters(1) = Geral.DataProcessamento
                                        .rdoParameters(2) = aDoc(Count).IdDocto
                                        .rdoParameters(3) = 0
                                        .rdoParameters(4) = aDoc(Count).TipoDocto
                                        .rdoParameters(5) = aDoc(Count).Alcada
                                        .Execute
                                        If .rdoParameters(0) <> 0 Then
                                            GoTo ErroVinculo
                                        End If
                                    End With
                                End If
                            End If
                        Next
                        'Remover Ajustes (34) - CREDITO AUTOMATICO
                        '                (38) - DEBITO AUTOMATICO
                        '                (42) - AJUSTE CONTABIL RECEITA
                        '                (43) - AJUSTE CONTABIL DESPESA
                        '                (32) - AJUSTE DE CRÉDITO
                        '                (33) - AJUSTE DE DÉBITO
                    
                        With qryRemoveAjustesVinculoManual
                            .rdoParameters(0).Direction = rdParamReturnValue
                            .rdoParameters(1) = Geral.DataProcessamento
                            .rdoParameters(2) = m_Capa.IdCapa
                            .rdoParameters(3) = Vinculo
                            .Execute
                    
                            If .rdoParameters(0) <> 0 Then
                                GoTo ErroVinculo
                            End If
                        End With
                    End If
                End If
            End If
        End If
    Next
    
    On Error GoTo 0

    Screen.MousePointer = vbDefault

    ObtemDocumentos m_IdCapa

    m_Busy = False

    Erase aAjuste
    m_CountAjuste = 0

    tabVinculo.Tab = 0

    Preenche_lstCheques
    Preenche_lstPagtos
    Preenche_lstAutomatico
    Preenche_lstOcorrencia
    lstManual.Clear
    CalculaValores

    cmdOcorrencia.Enabled = False
    cmdCancelar.Enabled = False
    cmdVincular.Enabled = False
    LimparImagem

    Exit Sub
    
ErroVinculo:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro na atualização do vínculo do documento.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    m_Busy = False
    Unload Me
End Sub

Private Sub cmdEnviarCSP_Click()
    
    'Verificar se há alguma capa selecionada
    If FrmPesquisa.Visible = True Or m_IdCapa = 0 Then Exit Sub
    
    If lstCheques.ListCount < 1 And lstPagtos.ListCount < 1 Then
        MsgBox "Capa não possui documentos a serem enviados para CSP ", vbInformation + vbOKOnly, App.Title
        Exit Sub
    End If
    
    If MsgBox("Confirma o Envio da Capa para CSP ?", vbYesNo) = vbYes Then
        Screen.MousePointer = vbHourglass

        AtualizaStatusCapa m_IdCapa, "N"

        'Gravar Log
        Call GravaLog(m_IdCapa, 0, 273)

        m_IdCapa = 0
        'Posicionar na próxima Capa da Lista
        If cmbCapa.ListIndex < cmbCapa.ListCount - 1 Then
            cmbCapa.ListIndex = cmbCapa.ListIndex + 1
        Else
            CmdAtualizar_Click
        End If
    End If

End Sub

Private Sub CmdExibirCapa_Click()

    Dim rsDoc As rdoResultset

    qryGetImagemCapa.rdoParameters(0) = Geral.DataProcessamento
    qryGetImagemCapa.rdoParameters(1) = m_IdCapa
    Set rsDoc = qryGetImagemCapa.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    
    If rsDoc.EOF() Then Exit Sub

    m_Frente = rsDoc!Frente
    m_Verso = rsDoc!Verso
    m_Ordem = 0
    If Trim(m_Frente) <> "" Or Trim(m_Verso) <> "" Then
        MostraImagem
    Else
        LimparImagem
    End If
End Sub

Private Sub CmdFechar_Click()
    Unload Me
End Sub

Private Sub CmdFecharPesquisa_Click()
    '''''''''''''''''''''''''''''''''''''''
    'Grava Log MDI - Fim Aguarda documento'
    '''''''''''''''''''''''''''''''''''''''
    Call GravaLog(0, 0, 257)
    
    CmdFechar_Click
End Sub

Private Sub cmdFrenteVerso_Click()
    Dim i As Integer
    If m_Busy Then
        Exit Sub
    End If
    If Not cmdFrenteVerso.Enabled Or Len(m_Frente) = 0 Or Len(m_Verso) = 0 Then
        Exit Sub
    End If
    m_Busy = True
    
    On Error GoTo ErroImagem
    
    'só muda para frente/verso qdo docto vem da Ls500 e da Vips,
    'poi, o canon não gera verso.
    If (m_Ordem = "0") Or (m_Ordem = "2") Then
        If Lead1.Tag = "V" Then
            Lead1.Tag = "F"     'se verso, mostrar frente
            With Lead1
               .AutoRepaint = False
               If Geral.VIPSDLL = eDllProservi Then
                 .Load Geral.DiretorioImagens & m_Frente, 0, 0, 1
               Else
                 .Load Geral.DiretorioImagens & Format(aCapa(cmbCapa.ListIndex + 1).IdLote, "000000000") & "\" & m_Frente, 0, 0, 1
               End If

               'se ls500 mostrar mais escuro
               If (m_Ordem = "2") Then
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
                 .Load Geral.DiretorioImagens & m_Verso, 0, 0, 1
               Else
                 .Load Geral.DiretorioImagens & Format(aCapa(cmbCapa.ListIndex + 1).IdLote, "000000000") & "\" & m_Verso, 0, 0, 1
               End If

               If (m_Ordem = "2") Then
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
    frmImagem.Visible = False
    MsgBox "Não foi possível exibir imagem do documento, imagem não encontrada", vbExclamation + vbOKOnly, App.Title
    
End Sub

Private Sub cmdIlegiveis_Click()

    Dim rst         As RDO.rdoResultset
    Dim sStr        As String

    If FrmPesquisa.Visible = True Or m_IdCapa = 0 Then Exit Sub

    AtualizaStatusCapa m_IdCapa, "5"
    
    '''''''''''''''''''''''''''''''
    'Verifica se existe comentario'
    '''''''''''''''''''''''''''''''
    Set rst = GetControleCapa(Geral.DataProcessamento, m_IdCapa)
    sStr = ""
    If Not rst.EOF() Then
        sStr = rst!Comentario
    End If
    
    '''''''''''''''''''''''''''''''''
    'Insere registro no ControleCapa'
    '''''''''''''''''''''''''''''''''
    If Not InsereControleCapa(Geral.DataProcessamento, m_IdCapa, sStr, 17) Then
        MsgBox "Não foi possível inserir o Controle de Capa.", vbExclamation
    End If
   
    GravaLog m_IdCapa, 0, 70
    m_IdCapa = 0
    If cmbCapa.ListIndex < cmbCapa.ListCount - 1 Then
        cmbCapa.ListIndex = cmbCapa.ListIndex + 1
    Else
        CmdAtualizar_Click
    End If
End Sub

Private Sub cmdInverteCor_Click()

    If m_Busy Then
        Exit Sub
    End If

    If Not cmdInverteCor.Enabled Or Len(m_Frente) = 0 Or Len(m_Verso) = 0 Then
        Exit Sub
    End If

    m_Busy = True
    Lead1.Invert
    m_Busy = False
End Sub

Private Sub cmdOcorrencia_Click()

    Dim iOcorrencia As Integer
    Dim Count As Integer
    Dim i As Integer
    Dim strDescricao As String, IdDocto As Long
    
    If FrmPesquisa.Visible = True Then Exit Sub
    
    'Verificar se existe algum documento selecionado para ocorrencia
    If lstCheques.SelCount < 1 And lstPagtos.SelCount < 1 Then
        MsgBox "Nenhum Documento foi selecionado para ocorrência.", vbInformation + vbOKOnly, App.Title
        Exit Sub
    End If

    Ocorrencia.cmdRemoverOcorrencia.Tag = "CSP"
    
    'Obtem um IdDocto para pesquisa de complemento da ocorrência, esteja em cheque ou contas
    IdDocto = 0
    For Count = 0 To lstCheques.ListCount - 1
        If lstCheques.Selected(Count) Then
            IdDocto = aCheque(Count + 1).IdDocto
            Exit For
        End If
    Next
    If IdDocto = 0 Then
        For Count = 0 To lstPagtos.ListCount - 1
            If lstPagtos.Selected(Count) Then
                IdDocto = aPagto(Count + 1).IdDocto
                Exit For
            End If
        Next
    End If

    'Busca descrição do complemento de ocorrência, caso exista
    strDescricao = ""
'''    Call GravaComplementoOcorrencia(IdDocto, "C", strDescricao)
    
    Ocorrencia.m_Descricao = Trim(strDescricao)

    Ocorrencia.Show vbModal, Me
    
    If Ocorrencia.Result Then
        iOcorrencia = Ocorrencia.CodOcorr
        
        For Count = 0 To lstCheques.ListCount - 1
            If lstCheques.Selected(Count) Then
                i = IndiceCheque(Val(Right(lstCheques.List(Count), 10)))
                aCheque(i).Status = "D"
                aCheque(i).Ocorrencia = iOcorrencia
                aCheque(i).ComplementoOcorrencia = Ocorrencia.m_Descricao
            End If
        Next
        For Count = 0 To lstPagtos.ListCount - 1
            If lstPagtos.Selected(Count) Then
                i = IndicePagto(Val(Right(lstPagtos.List(Count), 10)))
                aPagto(i).Status = "D"
                aPagto(i).Ocorrencia = iOcorrencia
                aPagto(i).ComplementoOcorrencia = Ocorrencia.m_Descricao
            End If
        Next
        Preenche_lstCheques
        Preenche_lstPagtos
        Preenche_lstOcorrencia

    End If

    Unload Ocorrencia

    cmdVincular.Enabled = False

    Preenche_lstCheques
    Preenche_lstPagtos
    Preenche_lstOcorrencia
End Sub

Private Sub cmdRotacao_Click()
    If m_Busy Then
        Exit Sub
    End If
    If Not cmdRotacao.Enabled Or Len(m_Frente) = 0 Or Len(m_Verso) = 0 Then
        Exit Sub
    End If
    m_Busy = True
    Lead1.FastRotate 90
    m_Busy = False
End Sub

Private Sub cmdVincular_Click()

    If FrmPesquisa.Visible = True Then Exit Sub

    If aCapa(cmbCapa.ListIndex + 1).IdEnv_Mal = "E" Then
        VinculaEnvelope
    Else
        VinculaMalote
    End If

    LimparImagem
    CalculaValores
    
    If lstCheques.ListCount = 0 And lstPagtos.ListCount = 0 Then
        cmdConfirmar.SetFocus
    End If

End Sub
Private Sub cmdZoomMais_Click()
    If m_Busy Then
        Exit Sub
    End If
    If Not cmdZoomMais.Enabled Or Len(m_Frente) = 0 Or Len(m_Verso) = 0 Then
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
    If Not cmdZoomMenos.Enabled Or Len(m_Frente) = 0 Or Len(m_Verso) = 0 Then
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
    Call AtualizaAtividade(17)
    
    If m_FirstActivate Then
        m_FirstActivate = False
        LimparValores
        LimparHeader
        LimparListas

        tmrAtualiza.Enabled = True
        sTempo = 0
        m_IdCapa = 0

        If Not ObtemCapas Then
            MsgBox "Não existem Envelopes/Malotes com pendência de Vínculo Manual.", vbExclamation + vbOKOnly, App.Title
            m_IdCapa = 0
            LimparImagem
            HabilitaTimerPesquisa
            Exit Sub
        Else
            'Encontrou Capa -> Desabilitar timer de pesquisa
            FrmPesquisa.Visible = False
            tmrPesquisa.Enabled = False
        End If

        m_IdCapa = 0
        Preenche_cmbCapa
        cmbCapa.ListIndex = 0
    End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyAdd
            cmdZoomMais_Click
        Case vbKeySubtract
            cmdZoomMenos_Click
        Case vbKeyF10
            cmdInverteCor_Click
            KeyCode = 0
        Case vbKeyDivide
            cmdRotacao_Click
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
        
    Set qryGetCapaVinculoManual = Geral.Banco.CreateQuery("", "{Call GetCapaVinculoManual (?,?)}")
    Set qryGetDocumentoVinculoManual = Geral.Banco.CreateQuery("", "{Call GetDocumentoVinculoManual (?,?)}")
    Set qryGetocorrencia = Geral.Banco.CreateQuery("", "{Call GetOcorrencia (?)}")
    Set qryAtualizaStatusCapa = Geral.Banco.CreateQuery("", "{? = Call AtualizaStatusCapa (?,?,?)}")
    Set qryAtualizaVinculoDocumento = Geral.Banco.CreateQuery("", "{? = Call AtualizaVinculoDocumento (?,?,?,?,?)}")
    Set qryAtualizaDocumentoExcluido = Geral.Banco.CreateQuery("", "{ ? = Call AtualizaDocumentoExcluido (?,?,?,?,?)}")
    Set qryInsereAjuste = Geral.Banco.CreateQuery("", "{ ? = Call InsereAjuste (?,?,?,?,?,?,?,?,?)}")
    Set qryGetDescricaoDocumento = Geral.Banco.CreateQuery("", "{Call GetDescricaoDocumento (?)}")
    Set qryLerParametros = Geral.Banco.CreateQuery("", "{Call LerParametro (?)}")
    Set qryGetAgContaDocumento = Geral.Banco.CreateQuery("", "{Call GetAgContaDocumento (?,?,?)}")
    Set qryVerificaCapaDisponivel = Geral.Banco.CreateQuery("", "{ ? = Call VerificaCapaDisponivel (?,?,?,?,?)}")
    Set qryRemoveAjustesVinculoManual = Geral.Banco.CreateQuery("", "{ ? = Call RemoveAjustesVinculoManual (?,?,?)}")
    Set qryVA_GetDocumentosTransmitidos = Geral.Banco.CreateQuery("", "{ ? = Call VA_GetDocumentosTransmitidos (?,?)}")
    Set qryGetUltimaOrdemCaptura = Geral.Banco.CreateQuery("", "{Call GetUltimaOrdemCaptura (?,?)}")
    Set qryGetImagemCapa = Geral.Banco.CreateQuery("", "{Call GetImagemCapa (?,?)}")

    m_FirstActivate = True
    
    '''''''''''''''''''''''''''
    'Loga a acao Entrar Modulo'
    '''''''''''''''''''''''''''
    'Call GravaLog(0, 0, 170)

   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If m_Busy Then
        Cancel = 1
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    tmrAtualiza.Enabled = False
    tmrPesquisa.Enabled = False
    
    If m_IdCapa > 0 Then
        If PodeEncerrar Then
            If MsgBox("O Envelope/Malote atual não foi encerrado. " & vbCrLf & _
                    "Deseja efetivar os Vínculos Manuais?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                m_Busy = False
                Confirmar False
            Else
                If Not AtualizaStatusCapa(m_IdCapa, "7") Then
                    m_Busy = False
                    Exit Sub
                End If
                ' GravaLog m_IdCapa, 0, 196
            End If
        Else
            If Not AtualizaStatusCapa(m_IdCapa, "7") Then
                m_Busy = False
                Exit Sub
            End If
            ' GravaLog m_IdCapa, 0, 196
        End If
    End If
    
    qryGetCapaVinculoManual.Close
    qryGetDocumentoVinculoManual.Close
    qryGetocorrencia.Close
    qryAtualizaStatusCapa.Close
    qryAtualizaVinculoDocumento.Close
    qryAtualizaDocumentoExcluido.Close
    qryInsereAjuste.Close
    qryGetDescricaoDocumento.Close
    qryLerParametros.Close
    qryGetAgContaDocumento.Close
    qryVerificaCapaDisponivel.Close
    qryRemoveAjustesVinculoManual.Close
    qryGetUltimaOrdemCaptura.Close
    
    '''''''''''''''''''''''''
    'Loga a acao Sair Modulo'
    '''''''''''''''''''''''''
    'Call GravaLog(0, 0, 171)

End Sub

Private Sub Lead1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Lead1.AutoRubberBand = True
        Lead1.MousePointer = 2
    Else
        MostraImagem
    End If
End Sub

Private Sub Lead1_RubberBand()
    Dim zoomleft As Integer
    Dim zoomtop As Integer
    Dim zoomwidth As Integer
    Dim zoomheight As Integer
    
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

Private Sub lstAutomatico_Click()

    i = IndiceDocto(Val(Right(lstAutomatico.List(lstAutomatico.ListIndex), 10)))
    
    m_Frente = aDoc(i).Frente
    m_Verso = aDoc(i).Verso
    m_Ordem = aDoc(i).Ordem
    If Trim(m_Frente) <> "" Or Trim(m_Verso) <> "" Then
        MostraImagem
    Else
        LimparImagem
    End If

End Sub

Private Sub lstCheques_Click()

    CalculaValores

    If aCapa(cmbCapa.ListIndex + 1).IdEnv_Mal = "E" Then
        VerificaVinculoEnvelope
    Else
        VerificaVinculoMalote
    End If
End Sub
Private Sub lstCheques_DblClick()
    
    i = IndiceDocto(Val(Right(lstCheques.List(lstCheques.ListIndex), 10)))
    m_Frente = aDoc(i).Frente
    m_Verso = aDoc(i).Verso
    m_Ordem = aDoc(i).Ordem
    
    MostraImagem
End Sub

Private Sub lstManual_Click()

    i = IndiceDocto(Val(Right(lstManual.List(lstManual.ListIndex), 10)))
    
    m_Frente = aDoc(i).Frente
    m_Verso = aDoc(i).Verso
    m_Ordem = aDoc(i).Ordem
    If Trim(m_Frente) <> "" Or Trim(m_Verso) <> "" Then
        MostraImagem
    Else
        LimparImagem
    End If
End Sub

Private Sub lstOcorrencia_Click()
    ObtemOcorrencia
    
    
    ''''''''''''''''''''''''''''''''''''
    'Para mostrar a imagem do documento'
    ''''''''''''''''''''''''''''''''''''
    i = IndiceDocto(Val(Right(lstOcorrencia.List(lstOcorrencia.ListIndex), 10)))
    
    m_Frente = aDoc(i).Frente
    m_Verso = aDoc(i).Verso
    m_Ordem = aDoc(i).Ordem
    If Trim(m_Frente) <> "" Or Trim(m_Verso) <> "" Then
        MostraImagem
    Else
        LimparImagem
    End If
    
End Sub

Private Sub lstPagtos_Click()
    CalculaValores
    
    If aCapa(cmbCapa.ListIndex + 1).IdEnv_Mal = "E" Then
        VerificaVinculoEnvelope
    Else
        VerificaVinculoMalote
    End If
End Sub

Private Sub lstPagtos_DblClick()
    
    i = IndiceDocto(Val(Right(lstPagtos.List(lstPagtos.ListIndex), 10)))
    m_Frente = aDoc(i).Frente
    m_Verso = aDoc(i).Verso
    m_Ordem = aDoc(i).Ordem

    MostraImagem
End Sub

Private Sub tabVinculo_Click(PreviousTab As Integer)
    lblOcorrencia.Caption = ""
    LimparImagem
    m_Frente = ""
    m_Verso = ""
    m_Ordem = ""
End Sub

Private Sub tmrAtualiza_Timer()
    tmrAtualiza.Enabled = False
    If m_IdCapa > 0 Then
        sTempo = sTempo + Int(tmrAtualiza.Interval / 1000)
        If sTempo + Int(tmrAtualiza.Interval / 1000) >= Geral.Intervalo Then
            AtualizaStatusCapa m_IdCapa, "J"
            ' GravaLog m_IdCapa, 0, 195
            sTempo = 0
            '''''''''''''''''''''''''''''''''''''''
            'Grava Log MDI - Fim Aguarda documento'
            '''''''''''''''''''''''''''''''''''''''
            Call GravaLog(0, 0, 257)
        End If
    End If
    tmrAtualiza.Enabled = True
End Sub

Private Sub tmrPesquisa_Timer()
  tmrPesquisa.Enabled = False

  sTempo = sTempo + Int(tmrPesquisa.Interval / 1000)

  If sTempo + Int(tmrPesquisa.Interval / 1000) >= Geral.Atualizacao Then
    sTempo = 0
    If ObtemCapas Then
        FrmPesquisa.Visible = False
        '''''''''''''''''''''''''''''''''''''''
        'Grava log MDI - Fim Aguarda documento'
        '''''''''''''''''''''''''''''''''''''''
        Call GravaLog(0, 0, 257)
        
        Preenche_cmbCapa
        
        cmbCapa.ListIndex = 0
        Exit Sub
    End If

    ''''''''''''''''''''''''''''''''''''''''''
    'Grava log MDI - Inicio Aguarda documento'
    ''''''''''''''''''''''''''''''''''''''''''
    Call GravaLog(0, 0, 256)
    
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
Private Function DigitaAgenciaConta(ByRef Agencia As Integer, ByRef Conta As Long) As Boolean

    On Error GoTo Erro_DigitaAgenciaConta
    
    Dim rsGetDate As rdoResultset
        
    Set rsGetDate = Geral.Banco.OpenResultset("select getdate() Hora")
    DigitaAgenciaConta = False

    Call AgenciaContaAjuste.ShowModal(Agencia, Conta)

    If (Agencia <> 0 And Conta <> 0) Then
    
        DigitaAgenciaConta = True
        
        sFile = IIf(Right(Geral.DiretorioImagens, 1) = "\", Geral.DiretorioImagens, Geral.DiretorioImagens & "\") & "VinculoManual_" & Geral.Usuario & ".log"
        
        ''''''''''''''''''''''''''''''''''
        'Grava no arquivo texto o Ajuste '
        ''''''''''''''''''''''''''''''''''
        iFile = FreeFile
        Open sFile For Append As #iFile
            
        Print #iFile, "====== Ajuste Agencia/Conta Vinculo Manual======="
        Print #iFile, "DataProcessamento - " & Geral.DataProcessamento
        Print #iFile, "Capa              - " & aCapa(cmbCapa.ListIndex + 1).Capa
        Print #iFile, "Login             - " & Geral.Usuario
        Print #iFile, "Hora              - " & str(rsGetDate!Hora)
        Print #iFile, "Agência           - " & Agencia
        Print #iFile, "Conta             - " & Conta
        Print #iFile, "================================================="
        Print #iFile, Chr(13)
        Close #iFile
    End If

    Exit Function

Erro_DigitaAgenciaConta:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro na atualização do vínculo do documento.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    Resume
End Function

Private Function GeraAjuste(ByVal Agencia As Integer, ByVal Conta As Long, ByVal Valor As Currency, ByVal TipoAjuste As Integer) As Boolean

    GeraAjuste = False

    If Abs(m_ValDiferenca) <> 0 Then
        m_Ajuste.Agencia = Agencia
        m_Ajuste.Conta = Conta
        m_Ajuste.Valor = Valor
        m_Ajuste.TipoDocto = TipoAjuste
        m_Ajuste.Vinculo = 0
        m_CountAjuste = m_CountAjuste + 1
        If m_CountAjuste > 50 Then
            MsgBox "Não é possível gerar mais que 50 ajustes numa mesma Capa.", vbInformation + vbOKOnly, App.Title
            Exit Function
        End If

        aAjuste(m_CountAjuste) = m_Ajuste
    End If

    GeraAjuste = True

End Function


Private Function ExisteDeposito() As Boolean

    Dim x As Integer

    ExisteDeposito = False

    For x = 0 To lstPagtos.ListCount - 1
        If lstPagtos.Selected(x) = True Then
            If aPagto(x + 1).TipoDocto = 2 Or aPagto(x + 1).TipoDocto = 3 Then
                ExisteDeposito = True
                Exit Function
            End If
        End If
    Next x
End Function
