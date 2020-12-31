VERSION 5.00
Object = "{00080003-B1BA-11CE-ABC6-F5B2E79D9E3F}#8.0#0"; "LTOCX80N.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Alcada 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alçada"
   ClientHeight    =   7992
   ClientLeft      =   792
   ClientTop       =   768
   ClientWidth     =   10560
   ForeColor       =   &H8000000F&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7992
   ScaleWidth      =   10560
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox FrmPesquisa 
      Height          =   1932
      Left            =   2208
      ScaleHeight     =   1884
      ScaleWidth      =   5712
      TabIndex        =   35
      Top             =   3036
      Visible         =   0   'False
      Width           =   5760
      Begin VB.CommandButton CmdFecharPesquisa 
         Caption         =   "&Fechar"
         Height          =   312
         Left            =   2340
         TabIndex        =   36
         Top             =   1464
         Width           =   1068
      End
      Begin ComctlLib.ProgressBar Progress 
         Height          =   300
         Left            =   348
         TabIndex        =   37
         Top             =   912
         Width           =   5088
         _ExtentX        =   8975
         _ExtentY        =   529
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Pesquisando por Documentos para Alçada. Aguarde ..."
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
         TabIndex        =   38
         Top             =   576
         Width           =   4812
      End
   End
   Begin VB.Frame Frame1 
      Height          =   732
      Left            =   5412
      TabIndex        =   39
      Top             =   2796
      Width           =   3228
      Begin VB.TextBox txtConta 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
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
         Left            =   2016
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   240
         Width           =   1140
      End
      Begin VB.TextBox txtAgencia 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
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
         Left            =   816
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   240
         Width           =   588
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Conta"
         Height          =   192
         Left            =   1536
         TabIndex        =   41
         Top             =   288
         Width           =   420
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Agência"
         Height          =   192
         Left            =   144
         TabIndex        =   40
         Top             =   288
         Width           =   600
      End
   End
   Begin VB.Timer tmrPesquisa 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9000
      Top             =   3084
   End
   Begin VB.Timer tmrAtualiza 
      Enabled         =   0   'False
      Interval        =   50000
      Left            =   8988
      Top             =   2700
   End
   Begin VB.PictureBox Picture4 
      Height          =   264
      Left            =   1728
      ScaleHeight     =   216
      ScaleWidth      =   528
      TabIndex        =   31
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
         TabIndex        =   32
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.PictureBox picNumMalote 
      Height          =   264
      Left            =   3660
      ScaleHeight     =   216
      ScaleWidth      =   1176
      TabIndex        =   28
      Top             =   24
      Width           =   1224
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
         TabIndex        =   29
         Top             =   0
         Width           =   1092
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   264
      Left            =   1728
      ScaleHeight     =   216
      ScaleWidth      =   6864
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   360
      Width           =   6912
      Begin VB.Label Label1 
         Caption         =   "Vínculo"
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
         Height          =   192
         Left            =   708
         TabIndex        =   34
         Top             =   -12
         Width           =   828
      End
      Begin VB.Label Label6 
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
         ForeColor       =   &H00800000&
         Height          =   216
         Left            =   5448
         TabIndex        =   21
         Top             =   0
         Width           =   540
      End
      Begin VB.Label Label5 
         Caption         =   "Documento"
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
         Left            =   3000
         TabIndex        =   20
         Top             =   0
         Width           =   1104
      End
      Begin VB.Label Label4 
         Caption         =   "Ocorr."
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
         Left            =   2328
         TabIndex        =   19
         Top             =   0
         Width           =   636
      End
      Begin VB.Label Label3 
         Caption         =   "Alçada"
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
         Left            =   1620
         TabIndex        =   18
         Top             =   -12
         Width           =   732
      End
      Begin VB.Label Label2 
         Caption         =   "Nro."
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
         Height          =   192
         Left            =   108
         TabIndex        =   17
         Top             =   0
         Width           =   408
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   264
      Left            =   72
      ScaleHeight     =   216
      ScaleWidth      =   1524
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   24
      Width           =   1572
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
         TabIndex        =   16
         Top             =   -24
         Width           =   1272
      End
   End
   Begin VB.ListBox lstDocto 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1728
      Left            =   1728
      TabIndex        =   1
      Top             =   648
      Width           =   6912
   End
   Begin VB.ListBox lstCapa 
      Height          =   2016
      IntegralHeight  =   0   'False
      Left            =   60
      TabIndex        =   0
      Top             =   372
      Width           =   1572
   End
   Begin VB.Frame Frame2 
      Height          =   3588
      Left            =   8724
      TabIndex        =   22
      Top             =   -48
      Width           =   1752
      Begin VB.CommandButton cmdIlegiveis 
         Caption         =   "&Enviar Ilegíveis"
         Height          =   324
         Left            =   132
         TabIndex        =   7
         Top             =   1248
         Width           =   1464
      End
      Begin VB.CommandButton cmdOcorrencia 
         Caption         =   "&Ocorrência"
         Height          =   324
         Left            =   132
         TabIndex        =   6
         Top             =   900
         Width           =   1464
      End
      Begin VB.CommandButton cmdAtualizar 
         Caption         =   "&Atualizar"
         Height          =   324
         Left            =   132
         TabIndex        =   4
         Top             =   204
         Width           =   1464
      End
      Begin VB.CommandButton cmdFechar 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   324
         Left            =   132
         TabIndex        =   8
         Top             =   1596
         Width           =   1464
      End
      Begin VB.CommandButton cmdAprovar 
         Caption         =   "A&provar"
         Height          =   324
         Left            =   132
         TabIndex        =   5
         Top             =   552
         Width           =   1464
      End
   End
   Begin VB.Frame Frame6 
      Height          =   732
      Left            =   60
      TabIndex        =   26
      Top             =   2796
      Width           =   5292
      Begin VB.CheckBox chkOcorrencia 
         Caption         =   "Mostrar Documentos com Ocorrência"
         Height          =   204
         Left            =   108
         TabIndex        =   3
         Top             =   432
         Width           =   3048
      End
      Begin VB.CheckBox chkAlcada 
         Caption         =   "Mostrar Somente Documentos Pendentes de Aprovação"
         Height          =   204
         Left            =   108
         TabIndex        =   2
         Top             =   168
         Width           =   4416
      End
   End
   Begin VB.Frame FrmImagem 
      Caption         =   "Imagem"
      Height          =   4272
      Left            =   60
      TabIndex        =   23
      Top             =   3624
      Width           =   8616
      Begin LeadLib.Lead Lead1 
         Height          =   3900
         Left            =   96
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   216
         Width           =   8400
         _Version        =   524288
         _ExtentX        =   14817
         _ExtentY        =   6879
         _StockProps     =   229
         BackColor       =   16777215
         BorderStyle     =   1
         ScaleHeight     =   323
         ScaleWidth      =   698
         DataField       =   ""
         BitmapDataPath  =   ""
         AnnDataPath     =   ""
      End
   End
   Begin VB.Frame Frame5 
      Height          =   4272
      Left            =   8724
      TabIndex        =   24
      Top             =   3624
      Width           =   1752
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
         Picture         =   "Alcada.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   3480
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
         Picture         =   "Alcada.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2670
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
         Picture         =   "Alcada.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1860
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
         Picture         =   "Alcada.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1050
         Width           =   820
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
         Height          =   696
         Left            =   528
         Picture         =   "Alcada.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   820
      End
   End
   Begin VB.Label lblLote 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0001"
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
      Left            =   2352
      TabIndex        =   33
      Top             =   24
      Width           =   1200
   End
   Begin VB.Label lblNumMalote 
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
      Height          =   264
      Left            =   4968
      TabIndex        =   30
      Top             =   24
      Width           =   1500
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
      Height          =   300
      Left            =   84
      TabIndex        =   27
      Top             =   2496
      Width           =   8556
   End
End
Attribute VB_Name = "Alcada"
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
End Type

Private Type tpMyDoc
    Agencia                                 As String
    Conta                                   As String
    IdDocto                                 As Long
    TipoDocto                               As Integer
    Ocorrencia                              As Integer
    Vinculo                                 As Long
    Frente                                  As String
    Verso                                   As String
    Status                                  As String * 1
    Valor                                   As Currency
    Ordem                                   As String * 1
    Alcada                                  As String * 1
End Type

Private m_Busy                              As Boolean
Private m_IdCapa                            As Long
Private m_Capa                              As tpMyCapa
Private m_Doc                               As tpMyDoc
Private aCapa()                             As tpMyCapa
Private aDoc()                              As tpMyDoc
Private m_CountCapa                         As Integer
Private m_CountDocto                        As Integer
Private sTempo                              As Integer
Private m_FirstActivate                     As Boolean

Private qryGetCapaAlcada                    As rdoQuery
Private qryGetDocumentoAlcada               As rdoQuery
Private qryGetocorrencia                    As rdoQuery
Private qryAtualizaStatusCapa               As rdoQuery
Private qryAtualizaAlcadaDocumento          As rdoQuery
Private qryAtualizaDocumentoExcluido        As rdoQuery
Private qryRemoveDocumento                  As rdoQuery
Private qryVerificaCapaDisponivel           As rdoQuery
Private qryGetDocumentoAlcadaAgConta        As rdoQuery
Private qryGrupoUsuario                     As rdoQuery
Private rsCapa                              As rdoResultset
Private rsDoc                               As rdoResultset
Private RsOcorrencia                        As rdoResultset
Private rsGrupoUsuario                      As rdoResultset

Dim aGrupoUsuario()                          As String

Private Sub LimparHeader()
    lblCapa.Caption = ""
    lblNumMalote.Caption = ""
    lblLote.Caption = ""
    lblOcorrencia.Caption = ""
    chkAlcada.Value = 0
    chkOcorrencia.Value = 0
End Sub

Private Sub LimparListas()
    lstCapa.Clear
    LstDocto.Clear
End Sub

Private Sub Preenche_lstCapa()
    Dim Count                   As Integer
    lstCapa.Clear
    For Count = 1 To m_CountCapa
        lstCapa.AddItem aCapa(Count).Capa
    Next
End Sub

Private Sub Preenche_lstDocto()
    Dim Linha                   As String
    Dim Count                   As Integer
    
    LimparImagem
    LstDocto.Clear
    For Count = 1 To m_CountDocto
        If chkAlcada.Value = 0 Or aDoc(Count).Alcada = "S" Then
            If chkOcorrencia.Value = 1 Or (aDoc(Count).Status <> "D" And aDoc(Count).Status <> "F") Then
        
                Linha = Format(Count, "0000") & Space(1)
                Linha = Linha & Format(aDoc(Count).Vinculo, "0000000") & Space(3)
                Linha = Linha & IIf(aDoc(Count).Alcada = "S" And aDoc(Count).Status <> "D" And aDoc(Count).Status <> "F", "S", " ") & Space(5)
                Linha = Linha & IIf(aDoc(Count).Status <> "D" And aDoc(Count).Status <> "F", " ", "S") & Space(3)
                Select Case aDoc(Count).TipoDocto
                    Case 2, 3       ' Depósito
                        Linha = Linha & "DEPOSITO      " & Space(1)
                    Case 4          ' ADCC
                        Linha = Linha & "DEBITO CC     " & Space(1)
                    Case 5, 6, 7    ' Cheque
                        Linha = Linha & "CHEQUE        " & Space(1)
                    Case 37         ' oct
                        Linha = Linha & "OCT           " & Space(1)
                    Case 32, 34
                        Linha = Linha & "AJ. CRED.     " & Space(1)
                    Case 33, 38
                        Linha = Linha & "AJ. DEB.      " & Space(1)
                    Case 41
                        Linha = Linha & "LANCTO INTERNO" & Space(1)
                    Case 42
                        Linha = Linha & "AJ. REC.      " & Space(1)
                    Case 43
                        Linha = Linha & "AJ. DESP.     " & Space(1)
                    Case Else       ' Pagamento
                        Linha = Linha & "PAGAMENTO     " & Space(1)
                End Select
                Linha = Linha & FormataValor(aDoc(Count).Valor, 15) & Space(2)
                Linha = Linha & Format(aDoc(Count).IdDocto, "0000000000")
                LstDocto.AddItem Linha
            End If
        End If
    Next
End Sub

Private Function ObtemCapas() As Boolean
    On Error GoTo ErroGetCapa
    rdoErrors.Clear
    
    Screen.MousePointer = vbHourglass
    
    Erase aCapa
    m_CountCapa = 0
    
    qryGetCapaAlcada.rdoParameters(0) = Geral.DataProcessamento
    qryGetCapaAlcada.rdoParameters(1) = Geral.Intervalo
    Set rsCapa = qryGetCapaAlcada.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
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
        m_Capa.IdLote = rsCapa!IdLote
        m_Capa.IdEnv_Mal = rsCapa!IdEnv_Mal
        m_Capa.Capa = rsCapa!Capa
        m_Capa.NumMalote = rsCapa!Num_Malote
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
    Select Case TratamentoErro("Erro na obtenção de envelope/malote para Prova Zero.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    m_Busy = False
    Unload Me

End Function

Private Sub ObtemDocumentos(ByVal IdCapa As Long)

    Dim rsAgConta               As RDO.rdoResultset

    On Error GoTo ErroGetDocto
    rdoErrors.Clear
    
    Screen.MousePointer = vbHourglass
    
    Erase aDoc
    m_CountDocto = 0
    qryGetDocumentoAlcada.rdoParameters(0) = Geral.DataProcessamento
    qryGetDocumentoAlcada.rdoParameters(1) = IdCapa
    Set rsDoc = qryGetDocumentoAlcada.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    ReDim aDoc(rsDoc.RowCount)
    While Not rsDoc.EOF
        m_CountDocto = m_CountDocto + 1
        m_Doc.IdDocto = rsDoc!IdDocto
        m_Doc.TipoDocto = rsDoc!TipoDocto
        m_Doc.Ocorrencia = rsDoc!Ocorrencia
        m_Doc.Vinculo = rsDoc!Vinculo
        m_Doc.Frente = rsDoc!Frente
        m_Doc.Verso = rsDoc!Verso
        m_Doc.Status = rsDoc!Status
        m_Doc.Ordem = rsDoc!Ordem
        m_Doc.Valor = rsDoc!Valor
        If rsDoc!Status <> "D" And rsDoc!Status <> "F" Then
            m_Doc.Alcada = rsDoc!Alcada
        Else
            m_Doc.Alcada = ""
        End If
        
        With qryGetDocumentoAlcadaAgConta
            .rdoParameters(0) = Geral.DataProcessamento
            .rdoParameters(1) = m_Doc.TipoDocto
            .rdoParameters(2) = m_Doc.IdDocto
            Set rsAgConta = qryGetDocumentoAlcadaAgConta.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
        End With
        
        If Not rsAgConta.EOF Then
            With rsAgConta
                m_Doc.Agencia = !Agencia
                m_Doc.Conta = !Conta
            End With
        End If
        rsAgConta.Close
        
        aDoc(m_CountDocto) = m_Doc
        rsDoc.MoveNext
    Wend
    rsDoc.Close
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErroGetDocto:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro na obtenção de documentos para Alçada.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    m_Busy = False
    Unload Me
    
End Sub

Private Function VerificaCapaDisponivel(ByVal IdCapa As Long) As Boolean
    On Error GoTo ErroVerificaCapa
    rdoErrors.Clear
    
    Screen.MousePointer = vbHourglass
    
    With qryVerificaCapaDisponivel
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = IdCapa
        .rdoParameters(3) = "6"
        .rdoParameters(4) = "I"
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

Private Function Indice(ByVal IdDocto As Long) As Integer
    Dim Count                   As Integer
    For Count = 1 To m_CountDocto
        If aDoc(Count).IdDocto = IdDocto Then
            Indice = Count
            Exit Function
        End If
    Next
    Indice = 0
End Function

Private Sub MostraImagem()
    Dim i                       As Integer
    i = Indice(Val(Right(LstDocto.List(LstDocto.ListIndex), 10)))
    
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
         .Load Geral.DiretorioImagens & Format(aCapa(lstCapa.ListIndex + 1).IdLote, "000000000") & "\" & aDoc(i).Frente, 0, 0, 1
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
    
    'posiciona imagem sempre no começo
    Ret = SendMessage(hCtl, WM_VSCROLL, SB_BOTTOM, 0)
    Ret = SendMessage(hCtl, WM_HSCROLL, SB_TOP, 0)
    
    cmdZoomMais.Enabled = True
    cmdZoomMenos.Enabled = True
    cmdRotacao.Enabled = True
    cmdInverteCor.Enabled = True
    cmdFrenteVerso.Enabled = True
    On Error GoTo 0
    Exit Sub
    
ErroImagem:
    cmdZoomMais.Enabled = False
    cmdZoomMenos.Enabled = False
    cmdRotacao.Enabled = False
    cmdInverteCor.Enabled = False
    cmdFrenteVerso.Enabled = False
    FrmImagem.Visible = False
    MsgBox "Não foi possível exibir imagem do documento, imagem não encontrada", vbExclamation + vbOKOnly, App.Title

End Sub

Private Sub LimparImagem()
    FrmImagem.Visible = False
    cmdZoomMais.Enabled = False
    cmdZoomMenos.Enabled = False
    cmdRotacao.Enabled = False
    cmdInverteCor.Enabled = False
    cmdFrenteVerso.Enabled = False
End Sub

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
        Else
            'Gravar Log
            If Status = "R" Then
              Call GravaLog(IdCapa, 0, 91)
            ElseIf Status = "5" Then
              Call GravaLog(IdCapa, 0, 92)
            ' ElseIf Status = "I" Then
            '  Call GravaLog(IdCapa, 0, 197)
            ' ElseIf Status = "6" Then
            '  Call GravaLog(IdCapa, 0, 198)
            End If
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
    Unload Me
    
End Function

Private Function AtualizaAlcadaDocumento(ByVal IdDocto As Long, _
                                         ByVal Alcada As String) As Boolean
    On Error GoTo ErroAtualizaAlcada
    rdoErrors.Clear
    
    AtualizaAlcadaDocumento = True
    Screen.MousePointer = vbHourglass
    
    With qryAtualizaAlcadaDocumento
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = IdDocto
        .rdoParameters(3) = Alcada
        .Execute
        If .rdoParameters(0) <> 0 Then
            AtualizaValorDocumento = False
            Screen.MousePointer = vbDefault
            MsgBox "Erro na atualização da alçada do documento.", vbCritical + vbOKOnly, App.Title
        Else
            'Gravar Log
            Call GravaLog(aCapa(lstCapa.ListIndex + 1).IdCapa, IdDocto, 90)
        End If
    End With
    Screen.MousePointer = vbDefault
    On Error GoTo 0
    Exit Function

ErroAtualizaAlcada:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro na atualização da alçada do documento.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    m_Busy = False
    Unload Me
    
End Function

Private Function AtualizaDocumentoExcluido(ByVal IdDocto As Long, ByVal Ocorrencia As Long) As Boolean
    On Error GoTo ErroExclusao
    rdoErrors.Clear
    
    AtualizaDocumentoExcluido = True
    Screen.MousePointer = vbHourglass
    
    With qryAtualizaDocumentoExcluido
        .rdoParameters(0).Direction = rdParamReturnValue
        .rdoParameters(1) = Geral.DataProcessamento
        .rdoParameters(2) = IdDocto
        .rdoParameters(3) = "D" ' status
        .rdoParameters(4) = 0 ' duplicidade
        .rdoParameters(5) = Ocorrencia
        .Execute
        If .rdoParameters(0) <> 0 Then
            GoTo ErroExclusao
        End If
    End With
    Screen.MousePointer = vbDefault
    On Error GoTo 0
    Exit Function
    
ErroExclusao:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro na atualização do status do documento.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    m_Busy = False
    Unload Me

End Function

Private Sub ObtemOcorrencia()
    Dim i                       As Integer
    i = Indice(Val(Right(LstDocto.List(LstDocto.ListIndex), 10)))
    
    If aDoc(i).Status <> "D" And aDoc(i).Status <> "F" Then
        lblOcorrencia.Caption = ""
        Exit Sub
    End If
    
    On Error GoTo ErroOcorrencia
    rdoErrors.Clear
    
    Screen.MousePointer = vbHourglass
    
    qryGetocorrencia.rdoParameters(0) = aDoc(i).Ocorrencia
    Set RsOcorrencia = qryGetocorrencia.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    If RsOcorrencia.EOF Then
        Screen.MousePointer = vbDefault
        MsgBox "Código da Ocorrência não existe: " & str(aDoc(i).Ocorrencia) & ".", vbExclamation + vbOKOnly, App.Title
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

Private Sub PosicionaDocto()
    Dim Count                   As Integer
    Dim bExiste                 As Boolean
    
    bExiste = False
    For Count = 0 To LstDocto.ListCount - 1
        If Mid(LstDocto.List(Count), 16, 1) = "S" Then
            LstDocto.Selected(Count) = True
            bExiste = True
            Exit For
        End If
    Next
    If Not bExiste And LstDocto.ListCount > 0 Then
        LstDocto.Selected(0) = True
    End If
End Sub

Private Sub HabilitaTimerPesquisa()

  'Esta Função irá verificar a existência de documentos Ilegíveis a cada x segundos
  'de acordo com o campo PARAMETRO.TmAtualizacao
  FrmPesquisa.Visible = True
  TmrPesquisa.Enabled = True
  Progress.Value = 0
  
  ''''''''''''''''''''''''''''''''''''''''''
  'Grava log MDI - Inicio Aguarda documento'
  ''''''''''''''''''''''''''''''''''''''''''
  Call GravaLog(0, 0, 258)
  
  
End Sub

Private Sub chkAlcada_Click()
    Preenche_lstDocto
    PosicionaDocto
End Sub

Private Sub chkOcorrencia_Click()
    Preenche_lstDocto
    PosicionaDocto
End Sub

Private Sub cmdAprovar_Click()
    Dim Count                   As Integer
    Dim i                       As Integer
    Dim sFile                   As String
    Dim iFile                   As Integer
    Dim icount                    As Integer
    
    'Diretório e nome do arquivo texto para relatar liberação de alçada
    sFile = IIf(Right(Geral.DiretorioImagens, 1) = "\", Geral.DiretorioImagens, Geral.DiretorioImagens & "\") & "Alcada_" & Geral.Usuario & ".log"
    
    If FrmPesquisa.Visible = True Then Exit Sub
    
    If LstDocto.SelCount > 1 Then
        MsgBox "Não é possivel aprovar alçada com mais de um documento selecionado.", _
            vbInformation + vbOKOnly, App.Title
        Exit Sub
    End If
    
    i = Indice(Val(Right(LstDocto.List(LstDocto.ListIndex), 10)))
    If aDoc(i).Alcada <> "S" Or aDoc(i).Status = "D" Or aDoc(i).Status = "F" Then
        If LstDocto.ListIndex < LstDocto.ListCount - 1 Then
            LstDocto.ListIndex = LstDocto.ListIndex + 1
        End If
        Exit Sub
    End If
    
    'Verifica se alçada para malote
    If m_Capa.IdEnv_Mal = "M" Then
        'Verifica se docto para alçada e valor maior que o limite por usuário <> de Coordenador
        If aDoc(i).Alcada = "S" And aDoc(i).Valor > Geral.ValorAlcadaCoord_Mal Then
            'Verifica se usuário é COOrdenador
            If InStr(Geral.GrupoUsuario, "COO") = 0 Then
                MsgBox "Não é possivel aprovar alçada com valor superior à " & _
                        Format(Geral.ValorAlcadaCoord_Mal, " R$ ###,###,##0.00") & vbCrLf & vbCrLf & _
                        "Favor entrar em contato com o coordenador. ", vbInformation & vbOKOnly, App.Title
                Exit Sub
            End If
            
            ''''''''''''''''''''''''''''''''''''''''''''''
            '   Grava arquivo texto liberação da alçada
            ''''''''''''''''''''''''''''''''''''''''''''''
            iFile = FreeFile
            Open sFile For Append As #iFile
                
            Print #iFile, "================== LIBERACAO DE ALCADA ==================="
            Print #iFile, "DataProcessamento - " & Geral.DataProcessamento
            Print #iFile, "IdCapa            - " & aCapa(lstCapa.ListIndex + 1).IdCapa
            Print #iFile, "Capa              - " & aCapa(lstCapa.ListIndex + 1).Capa
            Print #iFile, "IdDocto           - " & aDoc(i).IdDocto
            Print #iFile, "Login             - " & Geral.Usuario
            Print #iFile, "Usuario           - " & Geral.NomeUsuario
            Print #iFile, "Hora              - " & Time
            
            For icount = 1 To UBound(aGrupoUsuario)
                If InStr(Geral.GrupoUsuario, aGrupoUsuario(icount, 1)) <> 0 Then
                    Print #iFile, "Grupo de usuario  - " & aGrupoUsuario(icount, 2)
                End If
            Next
            Print #iFile, "=========================================================="
            Close #iFile
        
        End If
    End If
    
    aDoc(i).Alcada = "N"
    If Not AtualizaAlcadaDocumento(aDoc(i).IdDocto, aDoc(i).Alcada) Then
        m_Busy = False
        m_IdCapa = 0
        Unload Me
        Exit Sub
    End If
    
    ListIndex = LstDocto.ListIndex
    
    Preenche_lstDocto
    
    FinalizaCapa
    
End Sub

Private Sub FinalizaCapa()
    Dim Count                   As Integer
    Dim bFinal                  As Boolean

    bFinal = True
    
    For Count = 1 To m_CountDocto
        If aDoc(Count).Alcada = "S" And aDoc(Count).Status <> "D" And aDoc(Count).Status <> "F" Then
            bFinal = False
            Exit For
        End If
    Next
    
    If bFinal Then
        If Not AtualizaStatusCapa(m_IdCapa, "R") Then
            m_Busy = False
            m_IdCapa = 0
            Unload Me
            Exit Sub
        End If
        
        m_IdCapa = 0
        
        If lstCapa.ListIndex < lstCapa.ListCount - 1 Then
            lstCapa.ListIndex = lstCapa.ListIndex + 1
        Else
            CmdAtualizar_Click
        End If
    Else
        PosicionaDocto
    End If

    LstDocto.SetFocus
End Sub

Private Sub CmdAtualizar_Click()
    If m_IdCapa > 0 Then
        If Not AtualizaStatusCapa(m_IdCapa, "6") Then
            m_Busy = False
            m_IdCapa = 0
            Exit Sub
        End If
    End If
    
    LimparHeader
    LimparListas
    
    If Not ObtemCapas Then
        MsgBox "Não existem Envelopes/Malotes com pendência de Alçada.", vbExclamation + vbOKOnly, App.Title
        m_IdCapa = 0
        LimparImagem
        HabilitaTimerPesquisa
        Exit Sub
    Else
        TmrPesquisa.Enabled = False
        FrmPesquisa.Visible = False
    End If
    Preenche_lstCapa
    lstCapa.Selected(0) = True
End Sub

Private Sub CmdFechar_Click()
    Unload Me
End Sub

Private Sub CmdFecharPesquisa_Click()
    ''''''''''''''''''''''''''''''''''''''
    'Loga a acao de Fim Aguarda documento'
    ''''''''''''''''''''''''''''''''''''''
    GravaLog 0, 0, 259
    
    CmdFechar_Click
End Sub

Private Sub cmdFrenteVerso_Click()
    Dim i                       As Integer
    If m_Busy Then
        Exit Sub
    End If
    If Not cmdFrenteVerso.Enabled Then
        Exit Sub
    End If
    m_Busy = True
    
    On Error GoTo ErroImagem
    
    i = Indice(Val(Right(LstDocto.List(LstDocto.ListIndex), 10)))
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
                 .Load Geral.DiretorioImagens & Format(aCapa(lstCapa.ListIndex + 1).IdLote, "000000000") & "\" & aDoc(i).Frente, 0, 0, 1
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
                 .Load Geral.DiretorioImagens & Format(aCapa(lstCapa.ListIndex + 1).IdLote, "000000000") & "\" & aDoc(i).Verso, 0, 0, 1
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
    cmdZoomMais.Enabled = False
    cmdZoomMenos.Enabled = False
    cmdRotacao.Enabled = False
    cmdInverteCor.Enabled = False
    cmdFrenteVerso.Enabled = False
    MsgBox "Não foi possível exibir imagem do documento, imagem não encontrada", vbExclamation + vbOKOnly, App.Title
    
End Sub

Private Sub cmdIlegiveis_Click()

    Dim rst         As RDO.rdoResultset
    Dim sStr        As String


    If FrmPesquisa.Visible = True Or m_IdCapa = 0 Then Exit Sub
    
    AtualizaStatusCapa m_IdCapa, "5"
    
    
    '''''''''''''''''''''''''''''''
    'Verifica se existe Comentario'
    '''''''''''''''''''''''''''''''
    Set rst = GetControleCapa(Geral.DataProcessamento, m_IdCapa)
    
    sStr = ""
    If Not rst.EOF() Then
        sStr = rst!Comentario
    End If
    '''''''''''''''''''''''''''''''''
    'Insere registro no ControleCapa'
    '''''''''''''''''''''''''''''''''
    If Not InsereControleCapa(Geral.DataProcessamento, m_IdCapa, sStr, 16) Then
        MsgBox "Não foi possível inserir o Controle de Capa.", vbExclamation
    End If
    
    m_IdCapa = 0
    If lstCapa.ListIndex < lstCapa.ListCount - 1 Then
        lstCapa.ListIndex = lstCapa.ListIndex + 1
    Else
        CmdAtualizar_Click
    End If
End Sub

Private Sub cmdInverteCor_Click()
    If m_Busy Then
        Exit Sub
    End If
    If Not cmdInverteCor.Enabled Then
        Exit Sub
    End If
    m_Busy = True
    Lead1.Invert
    m_Busy = False
End Sub

Private Sub cmdOcorrencia_Click()
    Dim iOcorrencia             As Long
    Dim iVinculo                As Long
    Dim Count                   As Integer
    Dim i                       As Integer
    Dim strDescricao            As String
    Dim IdDocto                 As Long
    
    If FrmPesquisa.Visible = True Then Exit Sub
    
    MsgBox "Atenção! Todos os documentos vinculados ao documento selecionado " & _
           "também serão devolvidos com a mesma ocorrência.", _
           vbInformation + vbOKOnly, App.Title
           
    'Busca descrição do complemento de ocorrência, caso exista
    strDescricao = ""
    IdDocto = Val(Right(LstDocto.List(LstDocto.ListIndex), 10))
'''    Call GravaComplementoOcorrencia(IdDocto, "C", strDescricao)
    
    Ocorrencia.m_Descricao = Trim(strDescricao)
    
    Ocorrencia.Show vbModal, Me
    
    If Ocorrencia.Result Then
        iOcorrencia = Ocorrencia.CodOcorr
        i = Indice(Val(Right(LstDocto.List(LstDocto.ListIndex), 10)))
        iVinculo = aDoc(i).Vinculo
        For Count = 1 To m_CountDocto
            If aDoc(Count).Vinculo = iVinculo And aDoc(Count).Status <> "D" And aDoc(Count).Status <> "F" Then
                If aDoc(Count).TipoDocto = 32 Or aDoc(Count).TipoDocto = 33 Or _
                   aDoc(Count).TipoDocto = 34 Or aDoc(Count).TipoDocto = 38 Then
                    If RemoveDocumento(aDoc(Count).IdDocto) Then
                        ' so para o array achar que ajuste foi devolvido
                        aDoc(Count).Status = "D"
                        aDoc(Count).Ocorrencia = iOcorrencia
                    End If
                Else
                    'Grava/Altera ou Exclui Complemento da Ocorrência
'''                    Call GravaComplementoOcorrencia(aDoc(Count).IdDocto, IIf(Ocorrencia.m_Descricao = "", "E", "G"), Ocorrencia.m_Descricao)
                
                    If AtualizaDocumentoExcluido(aDoc(Count).IdDocto, iOcorrencia) Then
                        aDoc(Count).Status = "D"
                        aDoc(Count).Ocorrencia = iOcorrencia
                    End If
                End If
            End If
        Next
    End If
    
    Preenche_lstDocto
    
    Unload Ocorrencia
    
    FinalizaCapa
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
    Call AtualizaAtividade(16)

    If m_FirstActivate Then
        LimparHeader
        LimparListas
        
        TmrAtualiza.Enabled = True
        sTempo = 0
        m_IdCapa = 0
        
        If Not ObtemCapas Then
            MsgBox "Não existem Envelopes/Malotes com pendência de Alçada.", vbExclamation + vbOKOnly, App.Title
            m_IdCapa = 0
            LimparImagem
            HabilitaTimerPesquisa
            Exit Sub
        End If
        Preenche_lstCapa
        m_FirstActivate = False
        lstCapa.Selected(0) = True
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
        
    Set qryGetCapaAlcada = Geral.Banco.CreateQuery("", "{Call GetCapaAlcada (?,?)}")
    Set qryGetDocumentoAlcada = Geral.Banco.CreateQuery("", "{Call GetDocumentoAlcada (?,?)}")
    Set qryAtualizaStatusCapa = Geral.Banco.CreateQuery("", "{? = Call AtualizaStatusCapa (?,?,?)}")
    Set qryAtualizaAlcadaDocumento = Geral.Banco.CreateQuery("", "{? = Call AtualizaAlcadaDocumento (?,?,?)}")
    Set qryAtualizaDocumentoExcluido = Geral.Banco.CreateQuery("", "{ ? = Call AtualizaDocumentoExcluido (?,?,?,?,?)}")
    Set qryGetocorrencia = Geral.Banco.CreateQuery("", "{ Call GetOcorrencia (?)}")
    Set qryObtemStatusCapa = Geral.Banco.CreateQuery("", "{ Call GetCapa (?,?)}")
    Set qryRemoveDocumento = Geral.Banco.CreateQuery("", "{? = Call RemoveDocumento (?,?)}")
    Set qryVerificaCapaDisponivel = Geral.Banco.CreateQuery("", "{ ? = Call VerificaCapaDisponivel (?,?,?,?,?)}")
    Set qryGetDocumentoAlcadaAgConta = Geral.Banco.CreateQuery("", "{Call GetDocumentoAlcadaAgConta(?,?,?)}")
    Set qryGrupoUsuario = Geral.Banco.CreateQuery("", "{call GetAllGrupos}")
    
    m_FirstActivate = True
    
    '''''''''''''''''''''''''''
    'Loga a acao Entrar Modulo'
    '''''''''''''''''''''''''''
    Call GravaLog(0, 0, 168)
    
    Set rsGrupoUsuario = qryGrupoUsuario.OpenResultset(rdOpenStatic, rdConcurReadOnly)
    With rsGrupoUsuario
        If Not .EOF Then ReDim aGrupoUsuario(.RowCount, 2)
        Do Until .EOF
            aGrupoUsuario(.AbsolutePosition, 1) = UCase(rsGrupoUsuario!IdGrupo)
            aGrupoUsuario(.AbsolutePosition, 2) = rsGrupoUsuario!Descricao
            .MoveNext
        Loop
    End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If m_Busy Then
        Cancel = 1
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Desabilitar os timers
    TmrAtualiza.Enabled = False
    TmrPesquisa.Enabled = False
    
    If m_IdCapa > 0 Then
        AtualizaStatusCapa m_IdCapa, "6"
    End If
    
    qryGetCapaAlcada.Close
    qryGetDocumentoAlcada.Close
    qryAtualizaStatusCapa.Close
    qryAtualizaAlcadaDocumento.Close
    qryAtualizaDocumentoExcluido.Close
    qryGetocorrencia.Close
    qryRemoveDocumento.Close
    qryVerificaCapaDisponivel.Close
    qryGetDocumentoAlcadaAgConta.Close
    qryGrupoUsuario.Close
    
    '''''''''''''''''''''''''
    'Loga a acao Sair Modulo'
    '''''''''''''''''''''''''
    Call GravaLog(0, 0, 169)

End Sub

Private Sub Lead1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Lead1.AutoRubberBand = True
        Lead1.MousePointer = 2
    Else
        MostraImagem
    End If
End Sub

Private Sub Lead1_RubberBand()
    Dim zoomleft                As Integer
    Dim zoomtop                 As Integer
    Dim zoomwidth               As Integer
    Dim zoomheight              As Integer
    
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

Private Sub lstCapa_Click()
    Dim Count                   As Integer
    Dim AindaExiste             As Boolean
    
    If m_Busy Then
        Exit Sub
    End If
    m_Busy = True
    
    If aCapa(lstCapa.ListIndex + 1).IdEnv_Mal = "E" Then
        lblCapa.Caption = "Envelope"
        picNumMalote.Visible = False
        lblNumMalote.Visible = False
    Else
        lblCapa.Caption = "Malote"
        picNumMalote.Visible = True
        lblNumMalote.Visible = True
        lblNumMalote.Caption = aCapa(lstCapa.ListIndex + 1).NumMalote
    End If
    
    If m_IdCapa > 0 Then
        If Not AtualizaStatusCapa(m_IdCapa, "6") Then
            m_Busy = False
            m_IdCapa = 0
            Exit Sub
        End If
    End If
    
    If m_CountCapa > 0 Then
        m_IdCapa = aCapa(lstCapa.ListIndex + 1).IdCapa
    End If
    
    If Not VerificaCapaDisponivel(m_IdCapa) Then
        m_IdCapa = 0
        m_Busy = False
        m_CountDocto = 0
        Preenche_lstDocto
        Exit Sub
    End If
    
    If Not AtualizaStatusCapa(m_IdCapa, "I") Then
        m_IdCapa = 0
        m_Busy = False
        Exit Sub
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''
    'Loga a acao 197 - Alcada - Selecionar Capa'
    ''''''''''''''''''''''''''''''''''''''''''''
    GravaLog m_IdCapa, 0, 197
    
    lblLote.Caption = Format(aCapa(lstCapa.ListIndex + 1).IdLote, "0000-00000")
    ObtemDocumentos m_IdCapa
    sTempo = 0
    Preenche_lstDocto
    PosicionaDocto
    
    m_Busy = False
End Sub

Private Sub lstCapa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys ("{TAB}")
    End If
End Sub

Private Sub lstDocto_Click()
    Dim i                       As Integer
    i = Indice(Val(Right(LstDocto.List(LstDocto.ListIndex), 10)))
    
    ObtemOcorrencia
    MostraImagem
    
    txtAgencia.Text = ""
    txtConta.Text = ""
    
    If aDoc(i).Alcada = "S" Then
        cmdAprovar.Enabled = True
        CmdOcorrencia.Enabled = True
        txtAgencia.Text = Format(aDoc(i).Agencia, "0000")
        txtConta.Text = Format(aDoc(i).Conta, "000000-0")
    Else
        cmdAprovar.Enabled = False
        CmdOcorrencia.Enabled = False
    End If
    LstDocto.SetFocus
    
End Sub

Private Sub lstDocto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call cmdAprovar_Click
        LstDocto.SetFocus
    End If
End Sub

Private Sub tmrAtualiza_Timer()
    TmrAtualiza.Enabled = False
    If m_IdCapa > 0 Then
        sTempo = sTempo + Int(TmrAtualiza.Interval / 1000)
        If sTempo + Int(TmrAtualiza.Interval / 1000) >= Geral.Intervalo Then
            AtualizaStatusCapa m_IdCapa, "I"
            sTempo = 0
            '''''''''''''''''''''''''''''''''''''''
            'Grava Log MDI - Fim Aguarda documento'
            '''''''''''''''''''''''''''''''''''''''
            Call GravaLog(0, 0, 259)
        End If
    End If
    TmrAtualiza.Enabled = True
End Sub

Private Sub tmrPesquisa_Timer()
  TmrPesquisa.Enabled = False

  sTempo = sTempo + Int(TmrPesquisa.Interval / 1000)

  If sTempo + Int(TmrPesquisa.Interval / 1000) >= Geral.Atualizacao Then
    sTempo = 0
    If ObtemCapas Then
        FrmPesquisa.Visible = False
        '''''''''''''''''''''''''''''''''''''''
        'Grava log MDI - Fim Aguarda documento'
        '''''''''''''''''''''''''''''''''''''''
        Call GravaLog(0, 0, 259)
        
        Preenche_lstCapa
        lstCapa.Selected(0) = True
        Exit Sub
    End If
    ''''''''''''''''''''''''''''''''''''''''''
    'Grava Log MDI - Inicio Aguarda documento'
    ''''''''''''''''''''''''''''''''''''''''''
    Call GravaLog(0, 0, 258)

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

Private Function RemoveDocumento(ByVal IdDocto As Long) As Boolean
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
            RemoveDocumento = True
        Else
            RemoveDocumento = False
            MsgBox "Erro. Não foi possível remover o documento.", vbCritical + vbOKOnly, App.Title
        End If
    End With
    On Error GoTo 0
    Exit Function
    
ErroRemoveDoc:
    Screen.MousePointer = vbDefault
    Select Case TratamentoErro("Erro. Não foi possível remover o documento.", Err, rdoErrors)
        Case vbCancel
        Case vbRetry
    End Select
    m_Busy = False
    Unload Me
End Function
