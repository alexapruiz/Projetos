VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmLogUsuario 
   Caption         =   "Log do Usuário"
   ClientHeight    =   7968
   ClientLeft      =   840
   ClientTop       =   756
   ClientWidth     =   11328
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7968
   ScaleWidth      =   11328
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkLogUsuario 
      Caption         =   "Log de usuário"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   6000
      TabIndex        =   21
      Top             =   3636
      Width           =   1692
   End
   Begin VB.Frame Frame3 
      Height          =   1260
      Left            =   12
      TabIndex        =   11
      Top             =   0
      Width           =   9456
      Begin VB.PictureBox Picture2 
         Height          =   396
         Left            =   4944
         ScaleHeight     =   348
         ScaleWidth      =   2052
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   240
         Width           =   2100
         Begin VB.Label lblNumMalote 
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
            Height          =   852
            Left            =   36
            TabIndex        =   17
            Top             =   36
            Width           =   1956
         End
      End
      Begin VB.PictureBox Picture4 
         Height          =   396
         Left            =   108
         ScaleHeight     =   348
         ScaleWidth      =   2052
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   240
         Width           =   2100
         Begin VB.Label Label4 
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
            TabIndex        =   15
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
      Begin VB.PictureBox Picture6 
         Height          =   396
         Left            =   108
         ScaleHeight     =   348
         ScaleWidth      =   2052
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   732
         Width           =   2100
         Begin VB.Label Label5 
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
            TabIndex        =   13
            Top             =   12
            Width           =   984
         End
      End
      Begin VB.ComboBox CmbAgencia 
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
         ItemData        =   "FrmLogUser.frx":0000
         Left            =   2280
         List            =   "FrmLogUser.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   732
         Width           =   2604
      End
      Begin VB.TextBox TxtNumMalote 
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
         TabIndex        =   2
         Top             =   228
         Width           =   2196
      End
      Begin VB.Label lblLote 
         Alignment       =   2  'Center
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
         Left            =   7116
         TabIndex        =   19
         Top             =   732
         Width           =   2196
      End
      Begin VB.Label Label8 
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
         TabIndex        =   18
         Top             =   732
         Width           =   2100
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MsfLogUser 
      Height          =   3816
      Left            =   312
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3960
      Width           =   10704
      _ExtentX        =   18881
      _ExtentY        =   6731
      _Version        =   393216
      Cols            =   5
   End
   Begin VB.Frame Frame2 
      Height          =   1260
      Left            =   9516
      TabIndex        =   1
      Top             =   0
      Width           =   1752
      Begin VB.CommandButton cmdFechar 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   324
         Left            =   156
         TabIndex        =   5
         Top             =   864
         Width           =   1464
      End
      Begin VB.CommandButton cmdLimpar 
         Caption         =   "&Limpar"
         Height          =   324
         Left            =   156
         TabIndex        =   3
         Top             =   180
         Width           =   1464
      End
      Begin VB.CommandButton CmdExec 
         Caption         =   "&Confirma"
         Height          =   324
         Left            =   156
         TabIndex        =   4
         Top             =   516
         Width           =   1464
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grade 
      Height          =   1872
      Left            =   372
      TabIndex        =   9
      Top             =   1380
      Width           =   10704
      _ExtentX        =   18881
      _ExtentY        =   3302
      _Version        =   393216
      AllowBigSelection=   0   'False
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin VB.CheckBox ChkInicio 
      Caption         =   "Inicialização"
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
      Left            =   8760
      TabIndex        =   20
      Top             =   3636
      Width           =   1668
   End
   Begin VB.Label lblExclusao 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   312
      Left            =   372
      TabIndex        =   10
      Top             =   3264
      Width           =   10692
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Lista de  Eventos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   516
      Left            =   -24
      TabIndex        =   7
      Top             =   3672
      Width           =   2388
   End
End
Attribute VB_Name = "FrmLogUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsAux                     As rdoResultset
Dim RsMotivoExclusao          As rdoResultset
Private RsOcorrencia          As rdoResultset
Private qryGetocorrencia      As rdoQuery
Private qryGetRetornoTransacao As rdoQuery
Private lngWidhtColCapa       As Long

Dim StrMotivo                 As String
Dim TEnvMal                   As String
Dim TipoDocto                 As Integer

'Identificador de consulta única referente à capa em
'Complementação, Ilegíveis, Prova Zero, Expedição ou Vínculo Manual
Dim bConsultaUnica            As Boolean

Private Type MdregOcorrencia
    qryGetLstOcorrencias      As rdoQuery
    qryGetMaloteLog           As rdoQuery
    qryGetMudaStatus          As rdoQuery
    qryGetDocumentosNumCapa   As rdoQuery
    qryGetMotivoExclusao      As rdoQuery
End Type
Private Type MdSupExclusao
    qryGetCapaLog             As rdoQuery
    qryGetAgCapaLog           As rdoQuery
    qryGetCapaLog1            As rdoQuery
    qryGetTodosUsuarios       As rdoQuery
End Type

Private MdSupExclusao         As MdSupExclusao
Private MdregOcorrencia       As MdregOcorrencia

'Query que traz o Log dos Usuários
Private QryGetLogUser As rdoQuery
Private RsLogUser As rdoResultset

'* Guarda o Nº do Identificação Documento
Dim nDocto_Sel As Integer

Dim rep As Integer, Indice As Integer, cont_scroll As Integer
Dim buf_scroll As String, texto_scroll As String
Dim valor_caption As String, visual As String, Tip_doc As String
Dim str_formatada As String * 13
Private Function ObtemOcorrencia(ByVal Ocorrencia As Long) As String
    
    Dim RetornoTransacao        As Integer
    
    On Error GoTo ErroOcorrencia
    rdoErrors.Clear
    Screen.MousePointer = vbHourglass
    
    If Ocorrencia > 99900 Then
    
        Set qryGetRetornoTransacao = Geral.Banco.CreateQuery("", "{Call GetRetornoTransacao(?,?)}")
        
        RetornoTransacao = Mid(Trim(str(Ocorrencia)), 4, 2)
    
         With qryGetRetornoTransacao
            .rdoParameters(0).Direction = rdParamInput
            .rdoParameters(1).Direction = rdParamOutput
            .rdoParameters(0) = RetornoTransacao
            .Execute

             ObtemOcorrencia = ""
            If IsNull(.rdoParameters(1).Value) Then
                ObtemOcorrencia = "Retorno de Mensagem não Tratado"
                Exit Function
            End If

            '''''''''''''''''''''
            'Retorna a descricao'
            '''''''''''''''''''''
            ObtemOcorrencia = Trim(.rdoParameters(1).Value)
        End With
        
        qryGetRetornoTransacao.Close
        
        
'        Select Case Mid(Trim(str(Ocorrencia)), 4, 2)
'            Case "41"
'                ObtemOcorrencia = "Erro Operacional - Arrecadaçao nao Conveniada"
'            Case "42"
'                ObtemOcorrencia = "Erro Operacional - Envelope recebido para Processamento"
'            Case "43"
'                ObtemOcorrencia = "Erro Operacional - Pagamento com cheque Roxo"
'            Case "44"
'                ObtemOcorrencia = "Erro Operacional - Agencia nao Cadastrada"
'            Case "45"
'                ObtemOcorrencia = "Erro Operacional - Ficha de Deposito ja utilizada"
'            Case "46"
'                ObtemOcorrencia = "Erro Operacional - Conta Poupança nao Encontrada"
'            Case "47"
'                ObtemOcorrencia = "Erro Operacional - Agencia nao Cadastrada"
'            Case "48"
'                ObtemOcorrencia = "Erro Operacional - Conta Corrente nao Encontrada"
'            Case "49"
'                ObtemOcorrencia = "Erro Operacional - Codigo de Barras Zerado"
'            Case "50"
'                ObtemOcorrencia = "Erro Operacional - Erro no envio da BHS1"
'            Case "51"
'                ObtemOcorrencia = "Erro Operacional - Retorno de Mensagem nao Tratado"
'            Case "52"
'                ObtemOcorrencia = "Erro Operacional - Erro no Vinculo (Cheque x Titulo)"
'            Case "53"
'                ObtemOcorrencia = "Erro Operacional - Valor dos cheques diferente do Informado"
'            Case "54"
'                ObtemOcorrencia = "Erro Operacional - Excluido pelo Supervisor"
'            Case "55"
'                ObtemOcorrencia = "Erro Operacional - Conta nao encontrada"
'            Case "56"
'                ObtemOcorrencia = "Erro Operacional - Conta Unibanco nao Existe"
'            Case Else
'                ObtemOcorrencia = "Erro Operacional - Retorno de Mensagem nao Tratado"
'        End Select
    Else

        If CLng(Len(Grade.TextMatrix(Grade.Row, 14))) > 4 Then
            Ocorrencia = Int(Ocorrencia / 100)
        End If
        qryGetocorrencia.rdoParameters(0) = Ocorrencia
        Set RsOcorrencia = qryGetocorrencia.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
           If RsOcorrencia.EOF Then
                If Trim(str(Ocorrencia)) = 0 Then
                    Else
                ObtemOcorrencia = "Ocorrência não Cadastrada " & " - " & Grade.TextMatrix(Grade.Row, 4)
                End If
            Else
                ObtemOcorrencia = RsOcorrencia!Descricao
            End If
        RsOcorrencia.Close
    End If
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
Function DescTipoProd() As String

    Select Case TipoDocto
    
        Case 0
            DescTipoProd = "INDEFINIDO"
        Case 1
            If TEnvMal = "E" Then
                DescTipoProd = "ENVELOPE"
            Else
                DescTipoProd = "MALOTE"
            End If
        Case 2, 3
            DescTipoProd = "DEPÓSITO"
        Case 4
            DescTipoProd = "DEBITO CC"
        Case 5, 6, 7
            DescTipoProd = "CHEQUE"
        Case 32, 34
            DescTipoProd = "AJUSTE DE CRÉDITO"
        Case 33, 38
            DescTipoProd = "AJUSTE DE DÉBITO"
        Case 36
            DescTipoProd = "CARTÃO AVULSO"
        Case 37
            DescTipoProd = "OCT"
        Case 39
            DescTipoProd = "CAPA OCT"
        Case 40
            DescTipoProd = "FGTS"
        Case 41
            DescTipoProd = "LANÇAMENTO INTERNO"
        Case 42
            DescTipoProd = "AJ. CONTÁBIL RECEITA"
        Case 43
            DescTipoProd = "AJ. CONTÁBIL DESPESA"
        Case Else
            DescTipoProd = "PAGAMENTO"
    
    End Select

End Function

Sub LeituraDoctosEnvelope()
                   
'Preenche List de Acordo com a Pesquisa

On Error GoTo ERRO_LEITURADOCTOS

    Dim Count_Set           As Integer
    Dim Count_Log           As Integer
    Dim RetornoTransacao    As Integer
                
    Grade.Rows = RsAux.RowCount + 1
                
    For Count_Log = 0 To RsAux.RowCount - 1
      'não imprime este dado na tela quando for devolvido pelo robo
        If IsNull(RsAux!Valor) = False Then
            FormataValor Trim(Format(RsAux!Valor, "##,##0.00"))
        Else
            valor_caption = "0"
        End If

        Grade.Row = Count_Log + 1
        Grade.Col = 1
        Grade.Text = IIf(IsNull(RsAux!Vinculo), "0000000000", Format(RsAux!Vinculo, "0000000000"))
        Grade.Col = 2
        If RsAux!TipoDocto = "1" Then
            Grade.Text = IIf(IsNull(RsAux!StatusCapa), "0", (RsAux!StatusCapa))
        Else
            Grade.Text = IIf(IsNull(RsAux!StatusDocto), "0", (RsAux!StatusDocto))
        End If
        Grade.Col = 3
        Grade.Text = IIf(IsNull(RsAux!Autenticado), " ", RsAux!Autenticado)
        Grade.Col = 4
        
        
        '''''''''''''''''''''''''
        'Trata Retorno Transacao'
        '''''''''''''''''''''''''
        If IsNull(RsAux!RetornoTransacao) = True Then
           RetornoTransacao = "00"
        ElseIf RsAux!RetornoTransacao = 0 Then
           RetornoTransacao = "00"
        Else
           RetornoTransacao = RsAux!RetornoTransacao
        End If
        
        'Grade.Text = IIf(IsNull(RsAux!Ocorrencia), "00000", Format(RsAux!Ocorrencia, "00000"))
        
        Grade.Text = IIf(IsNull(RsAux!Ocorrencia), "000", Mid(RsAux!Ocorrencia, 1, 3)) & RetornoTransacao
        
        Grade.Col = 5
        If IsNull(RsAux!Alcada) Or (RsAux!Alcada) = "N" Then
            Grade.Text = ""
        Else
            Grade.Text = RsAux!Alcada
        End If
        Grade.Col = 6
        If (IsNull(RsAux!Duplicidade)) Or (RsAux!Duplicidade) = 0 Then
            Grade.Text = " "
        ElseIf (RsAux!Duplicidade) = 1 Then
            Grade.Text = "S"
        End If
        Grade.Col = 7
        If IsNull(RsAux!TipoDocto) Or RsAux!TipoDocto = 0 Then
            TipoDocto = 0
            Call DescTipoProd
        Else
            TipoDocto = RsAux!TipoDocto
            Call DescTipoProd
        End If
        Grade.ColAlignment(7) = 1
        Grade.Text = DescTipoProd
        Grade.Col = 8
        Grade.Text = Trim(valor_caption)
        Grade.Col = 9
        Grade.Text = IIf(IsNull(RsAux!Frente), "", RsAux!Frente)
        Grade.Col = 10
        Grade.Text = IIf(IsNull(RsAux!Verso), "", RsAux!Frente)
        Grade.Col = 11
        Grade.Text = IIf(IsNull(RsAux!IdDocto), "0000000000", Format(RsAux!IdDocto, "0000000000"))
        Grade.Col = 12
        Grade.Text = IIf(IsNull(RsAux!TipoDocto), "", RsAux!TipoDocto)
        Grade.Col = 13
        Grade.Text = IIf(IsNull(RsAux!IdCapa), CmbAgencia.ItemData(CmbAgencia.ListIndex), RsAux!IdCapa)
        Grade.Col = 14
        Grade.Text = IIf(IsNull(RsAux!Ocorrencia), "00000", RsAux!Ocorrencia)
        
        RsAux.MoveNext

    Next
    
   '--- Posiciona o cursor no docto que foi sugerido na consulta '
    If nDocto_Sel <> 0 Then
        For Count_Log = 1 To Grade.Rows - 1
            Grade.Row = Count_Log
            Grade.Col = 11
            If Val(Grade.Text) = Val(nDocto_Sel) Then
                Grade.RowSel = Grade.Row
                Exit For
            End If
        Next
'Fase620
'    Else
'        Grade.RowSel = 1
'        Grade_SelChange
'        Grade_Click
    End If
   
   Exit Sub
                
ERRO_LEITURADOCTOS:

    Select Case TratamentoErro("Não foi possível consultar os dados.", Err, rdoErrors)
        Case vbCancel
            Unload Me
        Case vbRetry
        Resume
    End Select

End Sub
Private Sub FazPesquisa()
         
    If Trim(CmbAgencia.Text) = "" Then Exit Sub
        
'* Pesquisa Envelope / Malote *'
'* Traz a capa escolhida      *'
    With MdregOcorrencia.qryGetDocumentosNumCapa
        .rdoParameters(1).Value = Geral.DataProcessamento
        .rdoParameters(2).Value = Val(cmbCapa.Text)
        .rdoParameters(3).Value = CmbAgencia.ItemData(CmbAgencia.ListIndex)
        Set RsAux = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With

    If RsAux.EOF Then
        MsgBox "Capa não encontrada !", vbInformation, App.Title
        Call cmdLimpar_Click
        Exit Sub
    Else
        TEnvMal = RsAux!IdEnv_Mal
        'Resultado da pesquisa for <> 0
        'Popula List de Documentos
        LeituraDoctosEnvelope
    End If

End Sub
Sub FormataValor(ByVal vl_doc As String)
   
   'se encontrar o ponto (.), não formata
   If InStr(vl_doc, ".") = 0 Then
        valor_caption = Format(vl_doc, "0.00")
   Else
        valor_caption = vl_doc
   End If
   
   str_formatada = ""
   rep = 1                'contador de caracteres a serem formatados
   Do
      If (Mid$(valor_caption, rep, 1) = "") Then   'verifica término da string
         Exit Do
      End If
      rep = rep + 1
   Loop While (rep < 14)   'tamanho máximo da string a ser formatada
   rep = rep - 1
   
   '---- formata à direita ----
   Mid$(str_formatada, 13 - rep + 1, rep) = Mid$(valor_caption, 1, rep)
   valor_caption = str_formatada    'atualiza valor_caption com dado formatado a direita

End Sub
Private Sub limpa_Header()
    lblLote.Caption = ""
    cmbCapa.Clear
    TxtNumMalote.Text = ""
    CmbAgencia.Clear
End Sub
Private Sub cmdLimpaCampos_Click()
    LimpaTela Me
    txtNumEnvMal.SetFocus
End Sub
Private Sub cmdSair_Click()
    Unload Me
End Sub
Private Sub InicializaSis()

'* Preenche o Grid com os logs do usuário para capa documento *'

Dim ObtemIddocto As Long        'Obtem o Identificador do documento
Dim ObtemIdCapa  As Long        'Obtem o Identificador da capa
Dim ObtemOcorr   As Long        'Obtem o código da ocorrência

Dim StatusCapa   As String      'Recupera o Status da Capa
Dim Tip_doc      As String      'Recupera o Tipo de documento
Dim Buffer_Log   As String      'Recupera o buffer de log

Dim LogsCount    As Integer     'Countador de Loop
Dim CountRow     As Integer     'Countador de Loop - linha
Dim CountLine    As Integer     'Countador de Loop - linha
    
    If ChkInicio.Value = Unchecked Then Exit Sub
    
        CountRow = 0
        CountLine = 0
        MsfLogUser.Rows = 1
        StatusCapa = "0"
        ObtemOcorr = 0
        Tip_doc = 0
        ObtemIdCapa = 0
        ObtenIddocto = 0
       
    With QryGetLogUser
        .rdoParameters(0).Value = Geral.DataProcessamento
        .rdoParameters(1).Value = ObtemIddocto
        .rdoParameters(2).Value = ObtemIdCapa
        .rdoParameters(3).Value = "E"
        .rdoParameters(4).Value = "I"
        Set RsLogUser = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If Not RsLogUser.EOF Then
    
        MsfLogUser.Rows = RsLogUser.RowCount + 1
    
        For LogsCount = 0 To RsLogUser.RowCount - 1
            
            CountRow = CountRow + 1

            MsfLogUser.Row = CountRow
                
            MsfLogUser.Col = 1
            MsfLogUser.Text = Format(RsLogUser!Data, "DD/MM/YYYY") & " - " & Format(RsLogUser!Data, "HH:MM:SS")
            
            MsfLogUser.Col = 2
            MsfLogUser.Text = RsLogUser!Capa
            
            MsfLogUser.Col = 3
            MsfLogUser.CellAlignment = 3
            MsfLogUser.Text = IIf(IsNull(Left(Trim(RsLogUser!Nome), 15)), Trim(RsLogUser!Login), Left(Trim(RsLogUser!Nome), 15))
            
            MsfLogUser.Col = 4
            MsfLogUser.CellAlignment = 1
            MsfLogUser.Text = RsLogUser!Descricao
        
            RsLogUser.MoveNext
        Next
        
    End If
    
End Sub

Private Sub ChkInicio_Click()
    'Verifica se ChkInicio esta Checked
    If ChkInicio.Value = Checked Then
        Call InicializaSis
        Exit Sub
    End If
End Sub

Private Sub chkLogUsuario_Click()

    Static strCaptionLblNumMalote   As String
    Static intLengthLblMalote       As Integer
    Dim lWidth                      As Long
    
    If Not chkLogUsuario.Enabled Then Exit Sub
    
    If chkLogUsuario.Value = vbChecked Then
        chkLogUsuario.Enabled = False
        Call cmdLimpar_Click
        chkLogUsuario.Value = vbChecked
        chkLogUsuario.Enabled = True
        cmbCapa.Clear
        CmbAgencia.Clear
        lblLote.Caption = ""
        

        'Desabilita opções
        Label4.Enabled = False
        Label5.Enabled = False
        Label8.Enabled = False
        cmbCapa.Enabled = False
        CmbAgencia.Enabled = False
        ChkInicio.Enabled = False
        ChkInicio.Value = vbUnchecked
        strCaptionLblNumMalote = lblNumMalote.Caption: lblNumMalote.Caption = "Nome Login"
        intLengthLblMalote = TxtNumMalote.MaxLength
        TxtNumMalote.MaxLength = 10
        TxtNumMalote.SetFocus
        

        
        MsfLogUser.Cols = 6
        lWidth = MsfLogUser.Width / 100
        
        MsfLogUser.ColWidth(0) = lWidth * 20: MsfLogUser.TextMatrix(0, 0) = "Usuário alterado"
        MsfLogUser.ColWidth(1) = lWidth * 15: MsfLogUser.TextMatrix(0, 1) = "Ação efetuada"
        MsfLogUser.ColWidth(2) = lWidth * 15: MsfLogUser.TextMatrix(0, 2) = "Grupo antes"
        MsfLogUser.ColWidth(3) = lWidth * 15: MsfLogUser.TextMatrix(0, 3) = "Grupo depois"
        MsfLogUser.ColWidth(4) = lWidth * 15: MsfLogUser.TextMatrix(0, 4) = "Campo alterado"
        MsfLogUser.ColWidth(5) = lWidth * 20: MsfLogUser.TextMatrix(0, 5) = "Responsável"
        MsfLogUser.FixedCols = 0
        
        
        'lngWidhtColCapa = MsfLogUser.ColWidth(2)    'Esconde coluna com Nr da Capa
        
        'MsfLogUser.ColWidth(2) = 0
        'MsfLogUser.ColWidth(4) = MsfLogUser.ColWidth(4) + lngWidhtColCapa
        
        CmdExec.Enabled = False
    
    Else
        MsfLogUser.Cols = 5
        'Habilia opções
        Label4.Enabled = True
        Label5.Enabled = True
        Label8.Enabled = True
        cmbCapa.Enabled = True
        CmbAgencia.Enabled = True
        ChkInicio.Enabled = True
        chkLogUsuario.Value = vbUnchecked
        lblNumMalote.Caption = strCaptionLblNumMalote
        TxtNumMalote.MaxLength = intLengthLblMalote
        MsfLogUser.ColWidth(2) = lngWidhtColCapa    'Retorna coluna com Nr da Capa
        MsfLogUser.ColWidth(4) = MsfLogUser.ColWidth(4) - lngWidhtColCapa
        CmdExec.Enabled = True
        
        '''''''''''''''''''
        'Acerta as colunas'
        '''''''''''''''''''
        
        MsfLogUser.Row = 0
         
        MsfLogUser.ColWidth(0) = 0
        MsfLogUser.ColAlignment(0) = 3
        
        MsfLogUser.Col = 1
        MsfLogUser.ColWidth(1) = 1700
        MsfLogUser.Text = "Data - Hora"
        MsfLogUser.ColAlignment(1) = 3
        
        MsfLogUser.Col = 2
        MsfLogUser.ColWidth(2) = 1400
        MsfLogUser.ColAlignment(2) = 3
        MsfLogUser.Text = "Capa nº"
        
        MsfLogUser.Col = 3
        MsfLogUser.ColWidth(3) = 1900
        MsfLogUser.ColAlignment(3) = 3
        MsfLogUser.Text = "Usuário"
        
        MsfLogUser.Col = 4
        MsfLogUser.ColWidth(4) = 5620
        MsfLogUser.ColAlignment(4) = 3
        MsfLogUser.Text = "Ação"

        Call cmdLimpar_Click
    End If
    
End Sub

Private Sub cmbAgencia_Change()
    If IsNumeric(CmbAgencia.Text) = True Then
        If CmbAgencia.ListCount = 1 Then
            CmbAgencia.Text = CmbAgencia.List(0)
        End If
    Else
        MsgBox "Informação inválida para este campo.", vbInformation + vbOKOnly, App.Title
        CmbAgencia.SetFocus
    End If
End Sub
Private Sub cmbAgencia_Click()
'Muda o Número de Lote - Quando for Selecionado outra Agência

    If CmbAgencia.Text = "" Then Exit Sub
'Fase620
    If CmbAgencia.Tag = "N" Then Exit Sub

    With MdSupExclusao.qryGetAgCapaLog
        .rdoParameters(0).Value = CDbl(cmbCapa)
        .rdoParameters(1).Value = Geral.DataProcessamento
        .rdoParameters(2).Value = CInt(CmbAgencia)
        .rdoParameters(3).Value = CmbAgencia.ItemData(CmbAgencia.ListIndex)
        Set RsAux = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If Not RsAux.EOF Then
        lblLote = Format(RsAux!IdLote, "0000-00000")
        Call FazPesquisa
'Fase620
        If Not bConsultaUnica Then
            Grade.RowSel = 1
            Grade_SelChange
            Grade_Click
        End If
    End If
    
    
End Sub
Private Sub cmbAgencia_KeyPress(KeyAscii As Integer)

    If (KeyAscii = 13) Then
       If Len(cmbCapa) > 0 Then
            Call FazPesquisa
       End If

    ElseIf (KeyAscii = 27) Then
        KeyAscii = 0
        cmdSair_Click
    End If
    
End Sub
Private Sub cmbCapa_Change()
    If Len(Trim(cmbCapa.Text)) = 0 Then Exit Sub
    If IsNumeric(cmbCapa.Text) = False Then
        MsgBox "Informação inválida para este campo.", vbInformation + vbOKOnly, App.Title
        cmbCapa.Text = ""
        cmbCapa.SetFocus
    End If
    
    If Len(Trim(TxtNumMalote)) <> 0 Then
        If Len(Trim(cmbCapa)) <> 0 And cmbCapa.ListCount > 1 Then
            Call Pesquisa_Dados
        End If
    End If
End Sub
Private Sub cmbCapa_KeyPress(KeyAscii As Integer)
   If (KeyAscii = vbKeyReturn) Then
      If Len(Trim(cmbCapa)) > 0 Then
         Call Pesquisa_Dados
         If Trim(CmbAgencia.Text) <> "" Then
            Call FazPesquisa
'Fase620
            Grade.RowSel = 1
            Grade_SelChange
            Grade_Click
         End If
      End If
   ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
      KeyAscii = 0
   Else
      'Não permitir a digitação de mais de 18 caracteres
      If Len(cmbCapa.Text) >= 18 And cmbCapa.SelLength = 0 And (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Then
         KeyAscii = 0
      End If
   End If
End Sub
Private Sub CmdExec_Click()
        
    'Verifica Preenchimento do Numero da Capa de Malote
    If Len(Trim(cmbCapa)) > 0 Then
        Call Pesquisa_Dados
    End If
     
    'Verifica Preenchimento do Numero do Malote
    If Len(Trim(TxtNumMalote)) > 0 Then
        If VerificaMalote(TxtNumMalote) = False Then
            MsgBox "Número de Malote inválido.", vbInformation, App.Title
            Call cmdLimpar_Click
            TxtNumMalote.SelStart = 0
            TxtNumMalote.SelLength = Len(TxtNumMalote.Text)
            TxtNumMalote.SetFocus
            Exit Sub
        End If
        Call ProcCapa
    End If
    
    'Verifica Preencimento da Agencia da Capa
    If Len(CmbAgencia.Text) <> 0 Then
        Grade.Rows = 1
        MsfLogUser.Rows = 1
        Call FazPesquisa
        Grade.SetFocus
    End If

End Sub
Private Sub CmdFechar_Click()
    Unload Me
End Sub
Private Sub cmdLimpar_Click()
    
    If chkLogUsuario.Value = vbUnchecked Then
        CmbAgencia.Clear
        cmbCapa.Clear
        LimpaTela Me
        lblLote = ""
        lblExclusao.Caption = ""
        ChkInicio.Value = Unchecked
        cmbCapa.SetFocus
    Else
        TxtNumMalote.Text = ""
        TxtNumMalote.SetFocus
    End If
    
    Grade.Rows = 1
    Grade.Rows = 2
    Grade.Tag = ""
    MsfLogUser.Rows = 1
    
End Sub
Private Sub Form_Activate()

Dim nform As Form
   
    'Verifica se form ativado à partir do botão de auditoria vindo de
    'Complementação, Ilegíveis, Prova Zero, Expedição ou Vínculo Manual
    bConsultaUnica = False
    For Each nform In Forms
        Select Case nform.Name
            Case "Complementacao"
                If nform.Visible Then bConsultaUnica = True: Exit For
            Case "Ilegiveis"
                If nform.Visible Then bConsultaUnica = True: Exit For
            Case "ProvaZero"
                If nform.Visible Then bConsultaUnica = True: Exit For
            Case "VinculoManual"
                If nform.Visible Then bConsultaUnica = True: Exit For
            Case "Expedicao"
                If nform.Visible Then bConsultaUnica = True: Exit For
        End Select
    Next
   
   'Inclusão de chamada a rotina AtualizaAtividade
    Call AtualizaAtividade(18)
    
    LimpaTela Me
    
    MsfLogUser.Row = 0
     
    MsfLogUser.ColWidth(0) = 0
    MsfLogUser.ColAlignment(0) = 3
    
    MsfLogUser.Col = 1
    MsfLogUser.ColWidth(1) = 1700
    MsfLogUser.Text = "Data - Hora"
    MsfLogUser.ColAlignment(1) = 3
    
    MsfLogUser.Col = 2
    MsfLogUser.ColWidth(2) = 1400
    MsfLogUser.ColAlignment(2) = 3
    MsfLogUser.Text = "Capa nº"
    
    MsfLogUser.Col = 3
    MsfLogUser.ColWidth(3) = 1900
    MsfLogUser.ColAlignment(3) = 3
    MsfLogUser.Text = "Usuário"
    
    MsfLogUser.Col = 4
    MsfLogUser.ColWidth(4) = 5620
    MsfLogUser.ColAlignment(4) = 3
    MsfLogUser.Text = "Ação"
    
    Grade.Rows = 1
    Grade.Cols = 16
    
    Grade.Row = 0
    Grade.ColWidth(0) = 0
    
    Grade.Col = 1
    Grade.ColAlignment(1) = 3
    Grade.ColWidth(1) = 1250
    Grade.Text = "Vinculo"
    
    Grade.Col = 2
    Grade.ColWidth(2) = 1000
    Grade.ColAlignment(2) = 3
    Grade.Text = "Status"
    
    Grade.Col = 3
    Grade.ColWidth(3) = 1000
    Grade.ColAlignment(3) = 3
    Grade.Text = "Autenticado"
    
    Grade.Col = 4
    Grade.ColWidth(4) = 1100
    Grade.ColAlignment(4) = 3
    Grade.Text = "Ocorrência"
    
    Grade.Col = 5
    Grade.ColWidth(5) = 1000
    Grade.ColAlignment(5) = 3
    Grade.Text = "Alçada"
    
    
    Grade.Col = 6
    Grade.ColWidth(6) = 1000
    Grade.ColAlignment(6) = 3
    Grade.Text = "Duplicado"
    
    Grade.Col = 7
    Grade.ColWidth(7) = 2600
    Grade.ColAlignment(7) = 3
    Grade.Text = "Tipo de Documento"
    
    Grade.Col = 8
    Grade.ColWidth(8) = 1690
    Grade.ColAlignment(8) = 3
    Grade.Text = "Valor"
    
    Grade.Col = 9
    Grade.ColAlignment(9) = 3
    Grade.ColWidth(9) = 0
    Grade.Text = "Frente"
    
    Grade.Col = 10
    Grade.ColWidth(10) = 0
    Grade.ColAlignment(10) = 3
    Grade.Text = "Verso"
    
    Grade.Col = 11
    Grade.ColWidth(11) = 0
    Grade.ColAlignment(11) = 3
    Grade.Text = "Tp"
    
    Grade.Col = 12
    Grade.ColWidth(12) = 0
    Grade.ColAlignment(12) = 3
    Grade.Text = "Tipo Documento"
    
    Grade.Col = 13
    Grade.ColWidth(13) = 0
    Grade.ColAlignment(13) = 3
    Grade.Text = "IdCapa"

    Grade.ColWidth(14) = 0
    Grade.ColWidth(15) = 0
    
    If bConsultaUnica Then
        'Desabilita controles desnecessários
        cmdLimpar.Enabled = False
        CmdExec.Enabled = False
        chkLogUsuario.Enabled = False
        ChkInicio.Enabled = False
        
        'Carrega controles chaves para pesquisa
        cmbCapa.Text = Geral.Capa.Capa
        If Geral.Capa.IdEnv_Mal = "M" Then TxtNumMalote.Text = Geral.Capa.Num_Malote
        CmbAgencia.AddItem Geral.Capa.AgOrig
        CmbAgencia.ItemData(0) = Geral.Capa.IdCapa
        CmbAgencia.ListIndex = 0
    
        Grade.RowSel = 1
        Grade_SelChange
        Grade_Click

        TxtNumMalote.Enabled = False
        CmbAgencia.Enabled = False
        cmbCapa.Enabled = False
    End If
    
End Sub
Private Sub Form_Load()

Dim CountUsers As Integer

'Traz Todas as Capas (Malote ou Envolpe) do Período Corrente
Set MdSupExclusao.qryGetCapaLog = Geral.Banco.CreateQuery("", "{Call GetTodasCapas_OC(?,?,?,?)}")

Set MdSupExclusao.qryGetCapaLog1 = Geral.Banco.CreateQuery("", "{Call GetTodasCapas(?,?,?,?)}")

'Traz Todas as Agencias da Capa (Malote ou Envolpe) do Período Corrente
Set MdregOcorrencia.qryGetMaloteLog = Geral.Banco.CreateQuery("", "{Call GetMaloteExpedicao(?,?)}")

'Traz Todas as Agencias da Capa (Malote ou Envolpe) do Período Corrente
Set MdSupExclusao.qryGetAgCapaLog = Geral.Banco.CreateQuery("", "{Call GetAgenciasCapa(?,?,?,?)}")

'Traz Todas as Capas (Malote ou Envolpe) do Período Corrente
Set MdregOcorrencia.qryGetMudaStatus = Geral.Banco.CreateQuery("", "{Call GetMudaStatus(?,?,?,?)}")

'Traz Todos Documentos para  a Capa informada
Set MdregOcorrencia.qryGetDocumentosNumCapa = Geral.Banco.CreateQuery("", "{? = call GetDocumentosNumCapa(?,?,?)}")

'Traz Todos os Registro de Log para um determinado Documento
Set QryGetLogUser = Geral.Banco.CreateQuery("", "{Call GetLogUser(?,?,?,?,?)}")
    
'Traz Ocorrência dos Documentos
Set qryGetocorrencia = Geral.Banco.CreateQuery("", "{Call GetOcorrencia (?)}")
    
' Cria query para a Leitura do Motivo de Exclusão para capa Selecionada
Set MdregOcorrencia.qryGetMotivoExclusao = Geral.Banco.CreateQuery("", "{call GetMotivoExclusao(?,?)}")
    
    MsfLogUser.Rows = 1
    
End Sub
Private Sub OptCapaEnvelope_Click()
'* Posiciona o foco no Text - txtnumEnvmal *'
     txtNumEnvMal.SetFocus
End Sub
Private Sub OptCapaMalote_Click()
'* Posiciona o foco no Text - txtnumEnvmal *'
     txtNumEnvMal.SetFocus
End Sub
Private Sub TxtNumEnvMal_LostFocus()
'* Formata Numero de capa de envelope e malote *'
    If OptCapaEnvelope.Value Then
        txtNumEnvMal = Format(txtNumEnvMal, "00000000")
    Else
        txtNumEnvMal = Format(txtNumEnvMal, "00000000000000")
    End If

End Sub
Private Sub LstOcorrencias_DblClick()
    Call CmdExec_Click
End Sub
Private Sub Grade_Click()
'* Preenche o Grid com os logs do usuário para capa documento *'

Dim ObtemIddocto As Long        'Obtem o Identificador do documento
Dim ObtemIdCapa  As Long        'Obtem o Identificador da capa
Dim ObtemOcorr   As Long        'Obtem o código da ocorrência

Dim StatusCapa   As String      'Recupera o Status da Capa
Dim Tip_doc      As String      'Recupera o Tipo de documento
Dim Buffer_Log   As String      'Recupera o buffer de log

Dim LogsCount    As Integer     'Countador de Loop
Dim CountRow     As Integer     'Countador de Loop - linha
Dim CountLine    As Integer     'Countador de Loop - linha
    
    If Grade.Row = 0 Then Exit Sub
    
    'Controle para evitar carregar novamente o grid devido ao evento
    If Grade.Tag = Grade.TextMatrix(Grade.Row, 11) Then Exit Sub
    Grade.Tag = Grade.TextMatrix(Grade.Row, 11)
    
    If Grade.Rows <= 2 And Trim(Grade.TextMatrix(Grade.Row, 1)) = "" Then Exit Sub

        CountRow = 0
        CountLine = 0
        MsfLogUser.Rows = 1
        StatusCapa = (Grade.TextMatrix(Grade.Row, 2))
        ObtemOcorr = (Grade.TextMatrix(Grade.Row, 4))
        Tip_doc = (Grade.TextMatrix(Grade.Row, 12))
        ObtemIdCapa = (Grade.TextMatrix(Grade.Row, 13))

    If Grade.TextMatrix(Grade.Row, 12) = "1" Then
       ObtenIddocto = 0
    Else
       ObtemIddocto = Grade.TextMatrix(Grade.Row, 11)
    End If
       
    With QryGetLogUser
        .rdoParameters(0).Value = Geral.DataProcessamento
        .rdoParameters(1).Value = ObtemIddocto
        .rdoParameters(2).Value = CLng(Grade.TextMatrix(Grade.Row, 13))
        .rdoParameters(3).Value = TEnvMal
        .rdoParameters(4).Value = "N"
        Set RsLogUser = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If Not RsLogUser.EOF Then
    
        MsfLogUser.Rows = RsLogUser.RowCount + 1
    
        For LogsCount = 0 To RsLogUser.RowCount - 1
            
            CountRow = CountRow + 1

            MsfLogUser.Row = CountRow
                
            MsfLogUser.Col = 1
            MsfLogUser.Text = Format(RsLogUser!Data, "DD/MM/YYYY") & " - " & Format(RsLogUser!Data, "HH:MM:SS")
            
            MsfLogUser.Col = 2
            MsfLogUser.Text = RsLogUser!Capa
            
            MsfLogUser.Col = 3
            MsfLogUser.CellAlignment = 3
            If IsNull(RsLogUser!Nome) Then
                MsfLogUser.Text = Left(Trim(RsLogUser!Login), 15)
            Else
                MsfLogUser.Text = Left(Trim(RsLogUser!Nome), 15)
            End If
            
            MsfLogUser.Col = 4
            MsfLogUser.CellAlignment = 1
            MsfLogUser.Text = RsLogUser!Descricao
        
            RsLogUser.MoveNext
        Next
        
    End If
    
    '* Recupera descritivo do Motivo de Exclusão se  Status capa for = 'D'
    '  ,Ocorrencia '999' e Tipo de Documento = 1 (Capa de Envelope/Malote) *'
    If StatusCapa = "D" And ObtemOcorr = 999 And Tip_doc = "1" Then
        lblExclusao = ObtemMotivoExclusao(ObtemIdCapa)
    Else
        '* Se não recupera descritivo da ocorrência *'
        StrMotivo = ObtemOcorrencia(ObtemOcorr)
        If Not Trim(StrMotivo) = "" Then
            lblExclusao = "Ocorrência:" & " " & StrMotivo
            Grade.SetFocus
        Else
            lblExclusao = ""
            Grade.SetFocus
        End If
    End If

End Sub
Private Sub Grade_SelChange()
'* Seleciona linha da Grade *'
    
    Static m_VerificaAcesso As Boolean
    
    'Controle de Chamada / Acesso
    If m_VerificaAcesso = True Then Exit Sub
    
       m_VerificaAcesso = True
        
        If Grade.Rows <= 1 Then Exit Sub
            Grade.Row = Grade.RowSel
            Grade.Col = 0
            Grade.ColSel = 12
            Grade.SetFocus

            Grade_Click
    
        m_VerificaAcesso = False
    
End Sub
Private Sub MsfLogUser_Click()
    Grade.SetFocus
End Sub
Public Sub ProcCapa()

Dim CountCapaMalote As Integer
Dim CountRegExcluido As Integer


    If Len(TxtNumMalote.Text) = 9 Or Len(TxtNumMalote.Text) = 11 Or Len(TxtNumMalote.Text) = 12 Then
        TxtNumMalote.Text = FormataMalote(TxtNumMalote)
        CmbAgencia.Clear
        cmbCapa.Clear
    
        If Not (RsAux Is Nothing) Then RsAux.Close
    
            With MdregOcorrencia.qryGetMaloteLog
                .rdoParameters(0).Value = Geral.DataProcessamento
                .rdoParameters(1).Value = Val(TxtNumMalote)
                If Err = 13 Then
                    MsgBox "Dados inválidos, reentre!", vbInformation, App.Title
                    Call cmdLimpar_Click
                    TxtNumMalote.SelStart = 0
                    TxtNumMalote.SelLength = Len(TxtNumMalote.Text)
                    TxtNumMalote.SetFocus
                    Exit Sub
                End If
                Set RsAux = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
            End With
    
            If Not RsAux.EOF Then
                For CountCapaMalote = 0 To RsAux.RowCount - 1
                    cmbCapa.AddItem RsAux!Capa
                    Call RetiraDuplicidade(RsAux!Capa)
                    RsAux.MoveNext
                Next
    
            If cmbCapa.ListCount = 1 Then
                cmbCapa.Text = cmbCapa.List(0)
            ElseIf cmbCapa.ListCount > 1 Then
                cmbCapa.SetFocus
                SendKeys "{F4}"
            End If
    
            Call Pesquisa_Dados
    
        Else
            MsgBox "Registro não encontrado!", vbInformation, App.Title
            Call cmdLimpar_Click
            TxtNumMalote.SelStart = 0
            TxtNumMalote.SelLength = Len(TxtNumMalote.Text)
            TxtNumMalote.SetFocus
            Exit Sub
        End If
        
    Else
        MsgBox "Dados inválidos, reentre !", vbInformation, App.Title
        Call cmdLimpar_Click
        TxtNumMalote.SelStart = 0
        TxtNumMalote.SelLength = Len(TxtNumMalote)
        TxtNumMalote.SetFocus
        Exit Sub
    End If

End Sub
Private Sub txtNumMalote_Change()

    If chkLogUsuario.Value = vbChecked Then
        If MsfLogUser.Rows > 1 Then chkLogUsuario_Click
        Exit Sub
    End If
    
    If Len(Trim(TxtNumMalote)) = 0 Then Exit Sub
    
    If IsNumeric(TxtNumMalote.Text) = False Then
        MsgBox "Informação inválida para este campo.", vbInformation + vbOKOnly, App.Title
        With TxtNumMalote
            .Text = ""
            .SetFocus
        End With
    End If
End Sub
Private Sub TxtNumMalote_GotFocus()
  TxtNumMalote.SelStart = 0
  TxtNumMalote.SelLength = Len(TxtNumMalote.Text)
End Sub
Private Sub txtNumMalote_KeyPress(KeyAscii As Integer)

    Dim idUsuario As Long
    
  If (KeyAscii = vbKeyReturn) Then
     If Len(Trim(TxtNumMalote)) > 0 Then
        'Verifica se consulta somente log de usuário
        If chkLogUsuario.Value = vbUnchecked Then
            If VerificaMalote(TxtNumMalote) = False Then
                MsgBox "Número de Malote inválido.", vbInformation, App.Title
                Call cmdLimpar_Click
                TxtNumMalote.SelStart = 0
                TxtNumMalote.SelLength = Len(TxtNumMalote.Text)
                TxtNumMalote.SetFocus
                Exit Sub
            End If
            Call ProcCapa
            If Trim(CmbAgencia.Text) <> "" Then
               Call FazPesquisa
            End If
        Else
            'Verifica se usuário existe
            If Not ExisteUsuario(idUsuario) Then
                MsgBox "Usuário não localizado, tente novamente", vbInformation + vbOKOnly, App.Title
                TxtNumMalote.SelStart = 0
                TxtNumMalote.SelLength = TxtNumMalote.MaxLength
                TxtNumMalote.SetFocus
                Exit Sub
            End If
            
            'Preenche grid com todos log´s
            Call ListaLogPorUsuario(idUsuario)
            
        End If
     End If
  Else
    If chkLogUsuario.Value = vbUnchecked Then
        If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then KeyAscii = 0
    End If
  End If
  
End Sub
Public Sub Pesquisa_Dados()
      
Dim CountAg As Integer
Dim CountRegExcluido As Integer
       
    CmbAgencia.Clear

    If cmbCapa.Text = "" Then Exit Sub

    With MdSupExclusao.qryGetAgCapaLog
        .rdoParameters(0).Value = CDbl(cmbCapa)
        .rdoParameters(1).Value = Geral.DataProcessamento
        .rdoParameters(2).Value = Null
        .rdoParameters(3).Value = Null
        Set RsAux = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With

    If Not RsAux.EOF Then
        TEnvMal = RsAux!IdEnv_Mal
        For CountAg = 0 To RsAux.RowCount - 1
            CmbAgencia.AddItem RsAux!AgOrig
            CmbAgencia.ItemData(CmbAgencia.NewIndex) = RsAux!IdCapa
            If RsAux!Num_Malote <> 0 Then
                Me.TxtNumMalote.Text = FormataMalote(RsAux!Num_Malote)
            End If
            RsAux.MoveNext
        Next
    Else
       MsgBox "Registro não encontrado!", vbInformation, App.Title
       Call cmdLimpar_Click
       Exit Sub
    End If

    If CmbAgencia.ListCount = 1 Then
'Fase620
        CmbAgencia.Tag = "N"
        CmbAgencia.Text = CmbAgencia.List(0)
        CmbAgencia.Tag = ""
    ElseIf CmbAgencia.ListCount > 1 Then
           CmbAgencia.SetFocus
           SendKeys "{F4}"
    End If

End Sub
Function VerificaMalote(ValNumMalote As String) As Boolean
'* Verifica se número de malote é Válido *'

    If Len(ValNumMalote) = 12 And CStr(Mid(ValNumMalote, 1, 2)) <> "09" Then
       VerificaMalote = False
    Else
       VerificaMalote = True
    End If
        
End Function
Function RetiraDuplicidade(NumCapaMalote As Double)
'* Elimina da Lista de Capas sua duplicidades *'

Dim CountLoop   As Integer  'Conta o Loop de acordo com a quantidade de Capas no Combo
Dim CountCapa   As Integer  'Traz  a quantidade de registros duplidados
Dim GuardaItem  As Integer

    For CountLoop = 0 To cmbCapa.ListCount - 1
        If NumCapaMalote = cmbCapa.List(CountLoop) Then
            CountCapa = CountCapa + 1
            GuardaItem = CountLoop
        End If
        
        If CountCapa >= 2 Then
            cmbCapa.RemoveItem (CountLoop)
        End If
        
    Next
      
End Function
Function ObtemMotivoExclusao(IdcapaMot As Long) As String
'* Esta função tem a finalidade de retornar o descritivo do motivo de exclusão, '
'  para uma capa selecionada e que possua Status = 'D' - Deletado *'

    With MdregOcorrencia.qryGetMotivoExclusao
        .rdoParameters(0) = Geral.DataProcessamento
        .rdoParameters(1) = IdcapaMot
        Set RsMotivoExclusao = .OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    End With
    
    If Not RsMotivoExclusao.EOF Then
        ObtemMotivoExclusao = RsMotivoExclusao!MotivoExclusao
    End If

End Function

Private Function ExisteUsuario(ByRef lngIdUsuario As Long) As Boolean

    Dim rsUsuario           As rdoResultset
    Dim qryLoginUsuario     As rdoQuery
    
    ExisteUsuario = False
    lngIdUsuario = 0
    
    On Error GoTo Err_ExisteUsuario

    Set qryLoginUsuario = Geral.Banco.CreateQuery("", "{? = call GetUsuarioPorLogin(?)}")
    qryLoginUsuario.rdoParameters(0).Direction = rdParamReturnValue
    
    qryLoginUsuario.rdoParameters(1).Value = Trim(TxtNumMalote.Text)
    Set rsUsuario = qryLoginUsuario.OpenResultset(rdOpenStatic, rdConcurReadOnly)
    
    'Verifica se não ocorreu erro na procedure SQL
    If qryLoginUsuario.rdoParameters(0) <> 0 Then
        MsgBox "Erro na leitura de informações do usuário.", vbCritical + vbOKOnly, App.Title
        GoTo Exit_ExisteUsuario
    End If
    
    If rsUsuario.EOF Then GoTo Exit_ExisteUsuario
    
    lngIdUsuario = rsUsuario("IdUsuario")
        
    ExisteUsuario = True
    
    
Exit_ExisteUsuario:
    If Not (rsUsuario Is Nothing) Then Set rsUsuario = Nothing
    qryLoginUsuario.Close
    Exit Function
    
Err_ExisteUsuario:
    Beep
    MsgBox "Não foi possível ler informações do usuário.", vbCritical + vbOKOnly, App.Title
    GoTo Exit_ExisteUsuario

End Function
Private Sub ListaLogPorUsuario(idUsuario As Long)

    Dim rsLogUsuario            As rdoResultset
    Dim qryLogUsuario           As rdoQuery
    Dim icount                  As Integer
    Dim colMaxWidth()           As Double
    Dim i                       As Integer
    
    On Error GoTo Err_ListaLogPorUsuario

    Set qryLogUsuario = Geral.Banco.CreateQuery("", "{call GetLogPorUsuario(?)}")
    qryLogUsuario.rdoParameters(0).Value = idUsuario
    
    Set rsLogUsuario = qryLogUsuario.OpenResultset(rdOpenStatic, rdConcurReadOnly)
    
    If rsLogUsuario.EOF Then
        MsgBox "Não existe Log para este usuário", vbInformation + vbOKOnly, App.Title
        TxtNumMalote.SelStart = 0
        TxtNumMalote.SelLength = TxtNumMalote.MaxLength
        GoTo Exit_ListaLogPorUsuario
    End If
    
    
    MsfLogUser.Rows = rsLogUsuario.RowCount + 1
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Redimensiona o vetor para o numero de colunas do grid'
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ReDim colMaxWidth(MsfLogUser.Cols - 1) As Double
    
    For i = 0 To MsfLogUser.Cols - 1
        colMaxWidth(i) = MsfLogUser.ColWidth(i)
    Next i

    For icount = 0 To rsLogUsuario.RowCount - 1

        MsfLogUser.TextMatrix(icount + 1, 0) = rsLogUsuario("Usuario Alterado")
        MsfLogUser.TextMatrix(icount + 1, 1) = rsLogUsuario("Acao efetuada")
        MsfLogUser.TextMatrix(icount + 1, 2) = ResolveMapa(idUsuario, eRetornaDescricao, rsLogUsuario("GrupoAntes"))
        MsfLogUser.TextMatrix(icount + 1, 3) = ResolveMapa(idUsuario, eRetornaDescricao, rsLogUsuario("GrupoDepois"))
        MsfLogUser.TextMatrix(icount + 1, 4) = rsLogUsuario("Campo alterado")
        MsfLogUser.TextMatrix(icount + 1, 5) = rsLogUsuario("Usuario responsavel")
        
        For i = 0 To MsfLogUser.Cols - 1
            If TextWidth(Trim(MsfLogUser.TextMatrix(icount + 1, i))) > colMaxWidth(i) Then
                MsfLogUser.ColWidth(i) = TextWidth(MsfLogUser.TextMatrix(icount + 1, i)) + 100
                colMaxWidth(i) = TextWidth(Trim(MsfLogUser.TextMatrix(icount + 1, i)))
            End If
        Next i
        
        rsLogUsuario.MoveNext
    Next
        
        
Exit_ListaLogPorUsuario:
    If Not (rsLogUsuario Is Nothing) Then Set rsLogUsuario = Nothing
    qryLogUsuario.Close
    Exit Sub
    
Err_ListaLogPorUsuario:
    Beep
    MsgBox "Não foi possível ler informações de Log.", vbCritical + vbOKOnly, App.Title
    GoTo Exit_ListaLogPorUsuario

End Sub
