VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Supervisor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sistema de Captura - Supervisor"
   ClientHeight    =   7740
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   9930
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   9930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   204
      Left            =   7560
      TabIndex        =   14
      Top             =   7500
      Width           =   2328
      _ExtentX        =   4101
      _ExtentY        =   370
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Frame Frame6 
      Height          =   2412
      Left            =   7464
      TabIndex        =   15
      Top             =   0
      Width           =   2340
      Begin VB.CommandButton cmdAtualizar 
         Caption         =   "&Atualizar"
         Height          =   324
         Left            =   192
         TabIndex        =   4
         Top             =   168
         Width           =   1956
      End
      Begin VB.CommandButton cmdEncerrarBordero 
         Caption         =   "&Encerrar Borderô"
         Height          =   324
         Left            =   192
         TabIndex        =   7
         Top             =   1284
         Width           =   1956
      End
      Begin VB.CommandButton cmdProvaZero 
         Caption         =   "Enviar &Prova Zero"
         Height          =   324
         Left            =   192
         TabIndex        =   6
         Top             =   912
         Width           =   1956
      End
      Begin VB.CommandButton cmdExcluirBordero 
         Caption         =   "E&xcluir Borderô"
         Height          =   324
         Left            =   192
         TabIndex        =   5
         Top             =   540
         Width           =   1956
      End
      Begin VB.CommandButton cmdSair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   324
         Left            =   192
         TabIndex        =   8
         Top             =   1656
         Width           =   1956
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4956
      Left            =   120
      TabIndex        =   11
      Top             =   2424
      Width           =   9684
      Begin VB.Frame Frame5 
         Caption         =   "Motivo de Rejeição"
         Height          =   1824
         Left            =   120
         TabIndex        =   13
         Top             =   3048
         Width           =   9444
         Begin MSFlexGridLib.MSFlexGrid GrdMotivoRejeicao 
            Height          =   1584
            Left            =   48
            TabIndex        =   3
            Top             =   192
            Width           =   9348
            _ExtentX        =   16484
            _ExtentY        =   2778
            _Version        =   393216
            Rows            =   1
            FixedCols       =   0
            BackColorBkg    =   -2147483633
            FocusRect       =   0
            SelectionMode   =   1
            BorderStyle     =   0
            FormatString    =   "Código       |Descrição                             "
         End
      End
      Begin VB.Frame Frame4 
         Height          =   2700
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   9444
         Begin MSFlexGridLib.MSFlexGrid GrdCheques 
            Height          =   2532
            Left            =   48
            TabIndex        =   2
            Top             =   144
            Width           =   9348
            _ExtentX        =   16484
            _ExtentY        =   4445
            _Version        =   393216
            Rows            =   1
            Cols            =   6
            FixedCols       =   0
            BackColorBkg    =   -2147483633
            FocusRect       =   0
            SelectionMode   =   1
            BorderStyle     =   0
            FormatString    =   $"Supervisor.frx":0000
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2412
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   7236
      Begin VB.Frame Frame2 
         Height          =   2172
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   2220
         Begin VB.ListBox lstBordero 
            Height          =   1620
            Left            =   48
            TabIndex        =   0
            Top             =   144
            Width           =   2124
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GrdDatas 
         Height          =   2052
         Left            =   2376
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   216
         Width           =   4740
         _ExtentX        =   8361
         _ExtentY        =   3625
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         FixedCols       =   0
         BackColorBkg    =   -2147483633
         FocusRect       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         BorderStyle     =   0
         FormatString    =   "Data de Depósito    |Quantidade   |Valor do Depósito  "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   16
      Top             =   7440
      Width           =   9924
      _ExtentX        =   17515
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   13229
            MinWidth        =   13229
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4207
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Supervisor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'''''''''''''''''''''''''''''
'Constantes do grid de Datas'
'''''''''''''''''''''''''''''
Private Const COL_DT_DATADEPOSITO = 0
Private Const COL_DT_QUANTIDADE = 1
Private Const COL_DT_VALOR = 2
'''''''''''''''''''''''''''''''
'Constantes do grid de Cheques'
'''''''''''''''''''''''''''''''
Private Const COL_CH_SEQUENCIAL = 0
Private Const COL_CH_AGENCIA = 1
Private Const COL_CH_CONTA = 2
Private Const COL_CH_NUMEROCHEQUE = 3
Private Const COL_CH_CNPJCPF = 4
Private Const COL_CH_VALOR = 5
'''''''''''''''''''''''''''''''''''''''''''''
'Constatntes do grid de Motivos de Rejeições'
'''''''''''''''''''''''''''''''''''''''''''''
Private Const COL_MR_CODIGO = 0
Private Const COL_MR_DESCRICAO = 1
''''''''''''''''''''''''''''''''''''''''
'Constantes da barra de scroll vertical'
''''''''''''''''''''''''''''''''''''''''
Private Const SM_CXVSCROLL = 2
'''''''''''''''''''''''''''''''
'Definição de variaveis membro'
'''''''''''''''''''''''''''''''
Dim m_IdBordero                 As Long
Dim m_IdBordero_Atual           As Long
Dim m_StatusAnterior            As String
Dim m_IdCheque                  As Double
Dim m_DataProcessamento         As Long
Dim m_bColor                    As Double   'Variavel para segurar as cores
Dim m_fColor                    As Double   'Variavel para segurar as cores
Dim m_IsActive                  As Boolean
Dim m_IsEvent                   As Boolean
Dim m_Bordero_Tratado()         As Boolean
Dim objAguardaDocumento         As AguardaDocumento

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Tipo de dados do arquivo Cheques Indevidos para geração do Aviso de Diferenca'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Type tpArquivoAD
    DataOcorrencia              As String * 9
    CodigoOcorrencia            As String * 6
    CodigoMotivo                As String * 6
    DataDeposito                As String * 9
    Num_Bordero                 As String * 20
    CodigoCarteira              As String * 4
    Agencia                     As String * 6
    Conta                       As String * 6
    CodigoDevolucao             As String * 6
    CodigoCompensacao           As String * 3 ' CodigoCompensacao           As String * 6
    BancoEmitente               As String * 6
    AgenciaEmitente             As String * 6
    CcEmitente                  As String * 11
    NrChequeEmitente            As String * 11
    TipoCheque                  As String * 1  ' TipoCheque                  As String * 4
    TipoInscricao               As String * 4
    InscricaoEmitente           As String * 15
    Valor                       As String * 16
    Gerado                      As String * 2
    CrLf                        As String * 2
End Type
Private Sub AguardaDocumento()

    Dim Proc_Selecionar         As New Custodia.Selecionar
    
    Screen.MousePointer = vbHourglass
    
    Set objAguardaDocumento = New AguardaDocumento
    
    objAguardaDocumento.SetConnection g_cMainConnection
    objAguardaDocumento.Tempo = 30
    objAguardaDocumento.SetStatusBar Me.StatusBar
    objAguardaDocumento.SetProgressBar Me.ProgressBar1
    objAguardaDocumento.SQL = Proc_Selecionar.GetSupervisor(m_DataProcessamento, g_Parametros.TMP_Pendente)
    
    objAguardaDocumento.SetStatus "Aguardando novo Borderô..."
    Do While Not objAguardaDocumento.ExisteDocumento() And objAguardaDocumento.Finalizado = False
        DoEvents
        objAguardaDocumento.SQL = Proc_Selecionar.GetSupervisor(m_DataProcessamento, g_Parametros.TMP_Pendente)
    Loop
    
    If Not objAguardaDocumento.Finalizado Then
        If objAguardaDocumento.Recordset.RecordCount > 0 Then
            If PreencheLstBordero Then
                SelecionaBordero lstBordero.ItemData(0)
            End If
        End If
    End If
    
    Set objAguardaDocumento = Nothing
    
    Screen.MousePointer = vbDefault
End Sub

Public Sub LimpaGridCheques()

    GrdCheques.Rows = 1
    '''''''''''''''''''''''''''''''''
    'Remove a cor da seleção do Grid'
    '''''''''''''''''''''''''''''''''
    GrdCheques.BackColorSel = GrdCheques.BackColorFixed
    GrdCheques.ForeColorSel = GrdCheques.ForeColorFixed
    

End Sub

Private Sub LimpaGridDatas()
    
    GrdDatas.Rows = 1
    
    '''''''''''''''''''''''''''''''''
    'Remove a cor de Seleção do Grid'
    '''''''''''''''''''''''''''''''''
    GrdDatas.BackColorSel = GrdDatas.BackColorFixed
    GrdDatas.ForeColorSel = GrdDatas.ForeColorFixed
    
    
    

    
End Sub


Public Sub LimpaGridRejeicoes()

    GrdMotivoRejeicao.Rows = 1
    '''''''''''''''''''''''''''''''''
    'Remove a cor da seleção do Grid'
    '''''''''''''''''''''''''''''''''
    GrdMotivoRejeicao.BackColorSel = GrdMotivoRejeicao.BackColorFixed
    GrdMotivoRejeicao.ForeColorSel = GrdMotivoRejeicao.ForeColorFixed

End Sub

Private Sub PreencheGridCheques(ByVal pDataProcessamento As Long, _
                                ByVal pIdBordero As Long, _
                                ByVal pDataDeposito As Long)

    Dim rst                 As New ADODB.Recordset
    Dim Proc_Selecionar     As New Custodia.Selecionar
    'Dim cCheque             As New clsCheque
    Dim cCheque             As New CalculoCheque
    
    On Error GoTo Erro_PreencheGridCheques
    
    
    Screen.MousePointer = vbHourglass
    
    GrdCheques.Rows = 1
    
    If Not IsDate(Format(pDataDeposito, "0000/00/00")) Then Exit Sub
    '''''''''''''''''''''''''''''''''''''''''
    'Busca os cheques desta Data de Deposito'
    '''''''''''''''''''''''''''''''''''''''''
    Set rst = g_cMainConnection.Execute(Proc_Selecionar.GetChequesBordero( _
                                        pDataProcessamento, _
                                        pIdBordero, _
                                        pDataDeposito))
    ''''''''''''''''''''''''''''
    'Preenche o Grid de Cheques'
    ''''''''''''''''''''''''''''
    Do While Not rst.EOF
        GrdCheques.Rows = GrdCheques.Rows + 1
        
        cCheque.CMC7 = rst!CMC7
        
        cCheque.Calcula

        GrdCheques.TextMatrix(rst.AbsolutePosition, COL_CH_SEQUENCIAL) = FormataString(rst.AbsolutePosition, "0", 5, True)
        GrdCheques.TextMatrix(rst.AbsolutePosition, COL_CH_AGENCIA) = cCheque.Agencia
        GrdCheques.TextMatrix(rst.AbsolutePosition, COL_CH_CONTA) = cCheque.Conta
        GrdCheques.TextMatrix(rst.AbsolutePosition, COL_CH_NUMEROCHEQUE) = cCheque.NumeroCheque
        GrdCheques.TextMatrix(rst.AbsolutePosition, COL_CH_CNPJCPF) = rst!CNPJCPF
        GrdCheques.TextMatrix(rst.AbsolutePosition, COL_CH_VALOR) = Format(rst!Valor, MASK_VALOR)
        
        GrdCheques.RowData(rst.AbsolutePosition) = rst!IdCheque

        rst.MoveNext
    Loop
    
    ''''''''''''''''''''''''''''
    If GrdCheques.Rows > 1 Then
        GrdCheques.BackColorSel = m_bColor
        GrdCheques.ForeColorSel = m_fColor
    End If
    
    rst.Close
    Set Proc_Selecionar = Nothing
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
Erro_PreencheGridCheques:

    Screen.MousePointer = vbDefault

    TratamentoErro "Erro ao preencher o grid de cheques.", Err

End Sub

Private Sub PreencheGridDatas(ByVal pDataProcessamento As Long, ByVal pIdBordero As Long)

    Dim rst                 As New ADODB.Recordset
    Dim Proc_Selecionar     As New Custodia.Selecionar
    
    On Error GoTo Erro_PreencheGridDatas
    
    ''''''''''''''''''''''''''''''
    'Busca as datas deste Borderô'
    ''''''''''''''''''''''''''''''
    Set rst = g_cMainConnection.Execute(Proc_Selecionar.GetDatasBordero( _
                                        pDataProcessamento, _
                                        pIdBordero))
    
    
    ''''''''''''''''''''''''''
    'Preenche o Grid de Datas'
    ''''''''''''''''''''''''''
    GrdDatas.Rows = 1
    Do While Not rst.EOF
        GrdDatas.Rows = GrdDatas.Rows + 1
        GrdDatas.TextMatrix(rst.AbsolutePosition, COL_DT_DATADEPOSITO) = Format(Format(rst!DataDeposito, "0000/00/00"), "dd/mm/yyyy")
        GrdDatas.TextMatrix(rst.AbsolutePosition, COL_DT_QUANTIDADE) = rst!QuantidadeCheques
        GrdDatas.TextMatrix(rst.AbsolutePosition, COL_DT_VALOR) = Format(rst!ValorDeposito, MASK_VALOR)
        rst.MoveNext
    Loop
    '''''''''''''''''''''''''''''''''
    'Coloca a cor de seleção do Grid'
    '''''''''''''''''''''''''''''''''
    If GrdDatas.Rows > 1 Then
        GrdDatas.BackColorSel = m_bColor
        GrdDatas.ForeColorSel = m_fColor
    End If
    
    rst.Close
    Set Proc_Selecionar = Nothing
    
    Exit Sub
    
Erro_PreencheGridDatas:
    
    TratamentoErro "Não foi possível preencher o grid de datas.", Err

End Sub

Private Sub PreencheGridRejeicoes(ByVal pDataProcessamento As Long, _
                                  ByVal pIdBordero As Long, _
                                  ByVal pDataDeposito As Long, _
                                  ByVal pIdCheque As Double)

    Dim rst                 As New ADODB.Recordset
    Dim Proc_Selecionar     As New Custodia.Selecionar
    
    On Error GoTo Erro_PreencheGridRejeicoes
    
    GrdMotivoRejeicao.Rows = 1
    

    If Not IsDate(Format(pDataDeposito, "0000/00/00")) Then Exit Sub
    
    '''''''''''''''''''''''''''''''''
    'Busca as Rejeições deste Cheque'
    '''''''''''''''''''''''''''''''''
    Set rst = g_cMainConnection.Execute(Proc_Selecionar.GetRejeicoesCheque( _
                                        pDataProcessamento, _
                                        pIdBordero, _
                                        pDataDeposito, _
                                        pIdCheque))
    ''''''''''''''''''''''''''''''
    'Preenche o Grid de Rejeições'
    ''''''''''''''''''''''''''''''
    Do While Not rst.EOF
        GrdMotivoRejeicao.Rows = GrdMotivoRejeicao.Rows + 1
        GrdMotivoRejeicao.TextMatrix(rst.AbsolutePosition, COL_MR_CODIGO) = rst!CodigoErro
        GrdMotivoRejeicao.TextMatrix(rst.AbsolutePosition, COL_MR_DESCRICAO) = rst!Descricao
        rst.MoveNext
    Loop
    
    '''''''''''''''''''''''''''''''''
    'Coloca a cor de seleção no Grid'
    '''''''''''''''''''''''''''''''''
    If GrdMotivoRejeicao.Rows > 1 Then
        GrdMotivoRejeicao.BackColorSel = m_bColor
        GrdMotivoRejeicao.ForeColorSel = m_fColor
    End If
    
    
    rst.Close
    Set Proc_Selecionar = Nothing
    
    Exit Sub
    
Erro_PreencheGridRejeicoes:
    TratamentoErro "Não foi possível preencher a lista de Rejeições.", Err
End Sub

Private Function PreencheLstBordero() As Boolean

    Dim rst                 As New ADODB.Recordset
    Dim Proc_Selecionar     As New Custodia.Selecionar
    
    On Error GoTo Erro_PreencheLstBordero:
    
    PreencheLstBordero = False
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Verifica se existe bordero Para Supervisor ou Em Supervisor + Pendente'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set rst = g_cMainConnection.Execute(Proc_Selecionar.GetSupervisor(m_DataProcessamento, g_Parametros.TMP_Pendente))
    
    If Not rst.EOF Then
        lstBordero.Clear
        ''''''''''''''''''''''''''''''''''''
        'Limpa a lista de borderos tratados'
        ''''''''''''''''''''''''''''''''''''
        Erase m_Bordero_Tratado
        
        ''''''''''''''''''''''''''''''''''
        'Redimensiona a lista de borderos'
        ''''''''''''''''''''''''''''''''''
        ReDim m_Bordero_Tratado(rst.RecordCount - 1) As Boolean
        
        Do While Not rst.EOF
            lstBordero.AddItem FormataString(rst!Num_Bordero, "0", rst.Fields("Num_Bordero").DefinedSize, True)
            lstBordero.ItemData(lstBordero.NewIndex) = rst!IdBordero
            rst.MoveNext
        Loop
        PreencheLstBordero = True
    End If
    m_IdBordero_Atual = 0
    
    rst.Close
    Set Proc_Selecionar = Nothing
    
    Exit Function
    
Erro_PreencheLstBordero:
    
    TratamentoErro "Erro ao preencher a lista de Borderôs.", Err

End Function

Private Sub SelecionaBordero(ByVal pIdBordero As Long)
    
    Dim i       As Integer
    
    ''''''''''''''''''''''''''''''
    'Localiza o IdBordero no List'
    ''''''''''''''''''''''''''''''
    For i = 0 To lstBordero.ListCount - 1
        If lstBordero.ItemData(i) = pIdBordero Then
            lstBordero.ListIndex = i
            Exit For
        End If
    Next i
    
    
End Sub

Private Function SelecionaProximoBordero() As Boolean

    Dim i           As Integer

    SelecionaProximoBordero = False
    
    
    If lstBordero.ListIndex = -1 Then
        lstBordero.ListIndex = 0
        SelecionaProximoBordero = True
        Exit Function
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Começa a verificar apartir do index do lstBordero.   '
    'Se não houver mais bordero para ser tratado, reinicia'
    'até o index atual do lstbordero                      '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    ''''''''''''''''''''''''
    'de onde está até o fim'
    ''''''''''''''''''''''''
    For i = lstBordero.ListIndex To lstBordero.ListCount - 1
        If m_Bordero_Tratado(i) = False Then
            lstBordero.ListIndex = i
            SelecionaProximoBordero = True
            Exit Function
        End If
    Next i
    ''''''''''''''''''''''''''
    'do inicio ate aonde está'
    ''''''''''''''''''''''''''
    For i = 0 To lstBordero.ListIndex
        If m_Bordero_Tratado(i) = False Then
            lstBordero.ListIndex = i
            SelecionaProximoBordero = True
            Exit Function
        End If
    Next i


End Function

Private Sub cmdAtualizar_Click()

    Dim Proc_atualizar      As New Custodia.atualizar
    Dim lRetorno            As Long
    
    
    If lstBordero.ListIndex = -1 Then Exit Sub
    
    
    Call g_cMainConnection.Execute(Proc_atualizar.AtualizaStatusBordero( _
                                   m_DataProcessamento, _
                                   m_IdBordero, _
                                   "5"), _
                            lRetorno, _
                            adCmdText)

    If lRetorno = 0 Then
        MsgBox "Não foi possível efetuar o processo de atualização.", vbExclamation, Me.Caption
        Exit Sub
    End If


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Atualiza o status do bordero ativo para o status "Para Supervisor"'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    LimpaGridDatas
    LimpaGridCheques
    LimpaGridRejeicoes

    lstBordero.Clear
    
    '''''''''''''''''''''''''''''
    'Refazer a lista de borderos'
    '''''''''''''''''''''''''''''
    If Not PreencheLstBordero() Then
        ''''''''''''''''''''''''''''''''''''''
        'Aguarda novo borderô para Supervisor'
        ''''''''''''''''''''''''''''''''''''''
        If objAguardaDocumento.Finalizado = True Then
            AguardaDocumento
        End If
    End If

End Sub

Private Sub cmdEncerrarBordero_Click()

    Dim lRetorno                    As Long
    Dim i, j                        As Integer
    Dim sMsg                        As String
    Dim sNumBordero                 As String
    Dim clCB                        As New CalculoBordero
    Dim clSD                        As New SomatoriaDatas
    Dim Proc_Alterar                As New Custodia.atualizar
    Dim Proc_Selecionar             As New Custodia.Selecionar
    Dim Proc_Inserir                As New Custodia.Inserir
    Dim Proc_Excluir                As New Custodia.Excluir
    Dim lQtdChequesIndevidos        As Long
    Dim lQtdDatasIndevidos          As Long
    Dim lSeqOcorrencia              As Long
    Dim rst                         As New ADODB.Recordset
    Dim rstDescMotAD                As New ADODB.Recordset
    Dim sMotivoAD                   As String 'Codigo do motivo = 4
    Dim sMotivoAD2                  As String 'Codigo do motivo = 5
    Dim DB                          As DAO.Database
    Dim ArquivoAD                   As tpArquivoAD
    Dim iFile                       As Integer
    Dim sStr                        As String
    Dim iOldDestination             As Integer
    
    On Error GoTo Erro_EncerrarBordero
    
    
    If lstBordero.ListIndex = -1 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    '''''''''''''''''''''''''''''''''''''''''''''
    'Define as propriedades da classe de calculo'
    '''''''''''''''''''''''''''''''''''''''''''''
    clCB.SetConnection g_cMainConnection
    clCB.DataProcessamento = m_DataProcessamento
    clCB.IdBordero = m_IdBordero
    clCB.QuantidadeMaximaCheques = g_Parametros.QuantidadeCheques
    clCB.QuantidadeMaximaDatas = g_Parametros.QuantidadeDatas
    
    ''''''''''''''''''''''''''''''''''''''
    'Verifica se existe cheques indevidos'
    ''''''''''''''''''''''''''''''''''''''
    
    Call clCB.VoltaStatusChequesIndevidos
    
    If Not clCB.CalculaChequesIndevidosQTDE(lQtdChequesIndevidos, True) Then
        Set clCB = Nothing
        GoTo Erro_GerarAD
    End If

    Set rst = g_cMainConnection.Execute(Proc_Selecionar.GetIdBordero( _
                                        m_DataProcessamento, _
                                        m_IdBordero))
    If rst.EOF() Then
        Set clCB = Nothing
        GoTo Erro_AtualizaBordero
    End If

    '''''''''''''''''''''''''''''''''''''''''''''''''''
    'Existe cheques indevidos, Gerar AD para os mesmos'
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    If lQtdChequesIndevidos > 0 Then
    
        ''''''''''''''''''''
        'Inicia a Transacao'
        ''''''''''''''''''''
        On Error GoTo Erro_AtualizaBordero
        g_cMainConnection.BeginTrans
        '''''''''''''''''''''''''''''''''
        'Obtem a Sequencia da Ocorrencia'
        '''''''''''''''''''''''''''''''''
        
        
        
        lSeqOcorrencia = Format(CLng(Right(g_Parametros.CNPJ_Terceira, 9)) + 10000000, "000000000")
        
        If g_Parametros.Seq_Ocorrencia = 0 Or (g_Parametros.Seq_Ocorrencia > lSeqOcorrencia) Then
            lSeqOcorrencia = Format(CLng(Right(g_Parametros.CNPJ_Terceira, 9)), "000000000")
        Else
            lSeqOcorrencia = Format(g_Parametros.Seq_Ocorrencia, "000000000")
        End If
        
        iFile = FreeFile
        Open App.path & "\AvisoDiferenca.txt" For Binary As iFile
        
        ArquivoAD.CrLf = vbCrLf
        
        ''''''''''''''''''''''''''''
        'Loop das Datas de Deposito'
        ''''''''''''''''''''''''''''
        For i = 1 To clCB.DataDeposito.Count
            '''''''''''''''''''''''''''''''''''''''
            'Loop dos Cheques por Data de Deposito'
            '''''''''''''''''''''''''''''''''''''''
            For j = 1 To clCB.DataDeposito(i).Cheque.Count
            
                If clCB.DataDeposito(i).Cheque.Item(j).Status = "I" Then
                
                    Call clCB.DataDeposito(i).Cheque.Item(j).Calcula

                    'lSeqOcorrencia = lSeqOcorrencia + 1
                    lSeqOcorrencia = Format(lSeqOcorrencia + 1, "000000000")

                    With ArquivoAD
                        .DataOcorrencia = m_DataProcessamento & "*"
                        .CodigoOcorrencia = FormataString(lSeqOcorrencia, "0", Len(.CodigoOcorrencia) - 1, True) & "*"
                        .CodigoMotivo = FormataString(5, "0", Len(.CodigoMotivo) - 1, True) & "*"
                        .DataDeposito = clCB.DataDeposito(i).DataDeposito & "*"
                        .Num_Bordero = FormataString(rst!Num_Bordero, "0", Len(.Num_Bordero) - 1, True) & "*"
                        .CodigoCarteira = FormataString(rst!CodigoCarteira, "0", Len(.CodigoCarteira) - 1, True) & "*"
                        .Agencia = FormataString(rst!Agencia, "0", Len(.Agencia) - 1, True) & "*"
                        .Conta = FormataString(rst!Conta, "0", Len(.Conta) - 1, True) & "*"
                        .CodigoDevolucao = FormataString(0, "0", Len(.CodigoDevolucao) - 1, True) & "*"
                        .CodigoCompensacao = FormataString(clCB.DataDeposito(i).Cheque.Item(j).Comp, "0", Len(.CodigoCompensacao), True) & "*"
                        ' Teste - Ilco em 30.11.01
                        '.CodigoCompensacao = FormataString(clCB.DataDeposito(i).Cheque.Item(j).Comp, "0", Len(.CodigoCompensacao) - 1, True) & "*"
                        .BancoEmitente = FormataString(clCB.DataDeposito(i).Cheque.Item(j).Banco, "0", Len(.BancoEmitente) - 1, True) & "*"
                        .AgenciaEmitente = FormataString(clCB.DataDeposito(i).Cheque.Item(j).Agencia, "0", Len(.AgenciaEmitente) - 1, True) & "*"
                        .CcEmitente = FormataString(clCB.DataDeposito(i).Cheque.Item(j).Conta, "0", Len(.CcEmitente) - 1, True) & "*"
                        .NrChequeEmitente = FormataString(clCB.DataDeposito(i).Cheque.Item(j).NumeroCheque, "0", Len(.NrChequeEmitente) - 1, True) & "*"
                        .TipoCheque = FormataString(clCB.DataDeposito(i).Cheque.Item(j).Tipificacao, "0", Len(.TipoCheque), True) & "*"
                        ' Teste - Ilco em 30.11.01
                        '.TipoCheque = FormataString(clCB.DataDeposito(i).Cheque.Item(j).Tipificacao, "0", Len(.TipoCheque) - 1, True) & "*"
                        .TipoInscricao = FormataString(0, "0", Len(.TipoInscricao) - 1, True) & "*"
                        .InscricaoEmitente = FormataString(0, "0", Len(.InscricaoEmitente) - 1, True) & "*"
                        .Valor = FormataString(clCB.DataDeposito(i).Cheque.Item(j).Valor, "0", Len(.Valor) - 1, True) & "*"
                        .Gerado = 0
                    End With
                    ''''''''''''''''''''''''''''''''''''''''''''''''''
                    'Imprime no arquivo Aviso de Diferenca por cheque'
                    ''''''''''''''''''''''''''''''''''''''''''''''''''
                    Put #iFile, , ArquivoAD
                End If
            Next j
        Next i
        
        Close #iFile
        
        g_Parametros.Seq_Ocorrencia = Format(lSeqOcorrencia, "000000000")
        
        Call g_cMainConnection.Execute(Proc_Alterar.AtualizaParametros( _
                                       m_DataProcessamento, _
                                       g_Parametros.QuantidadeCheques, _
                                       g_Parametros.QuantidadeDatas, _
                                       g_Parametros.DiretorioTransmissao, _
                                       g_Parametros.DiretorioRecepcao, _
                                       g_Parametros.Codigo_USB, _
                                       g_Parametros.CodigoAgAcolhed, _
                                       g_Parametros.CPD_Origem, _
                                       g_Parametros.CPD_Destino, _
                                       g_Parametros.Codigo_Terceira, _
                                       g_Parametros.CNPJ_Terceira, _
                                       g_Parametros.Seq_Ocorrencia, _
                                       g_Parametros.UF_Terceira, _
                                       g_Parametros.CodigoAplicacao, _
                                       g_Parametros.ValorChequeLimite, _
                                       g_Parametros.HeaderAV, g_Parametros.chkSoma, _
                                       g_Parametros.Gerar_Arquivo_CEL, g_Parametros.Comp_Origem_CEL, _
                                       g_Parametros.Numero_Versao_Inicial_CEL, _
                                       g_Parametros.Numero_Versao_Final_CEL, _
                                       g_Parametros.QuantidadeMinimaDias, _
                                       g_Parametros.Cidade_Terceira, _
                                       g_Parametros.Nome_Terceira), _
                                lRetorno, adCmdText)

        If lRetorno = 0 Then
            Set clCB = Nothing
            GoTo Erro_AtualizaBordero
        End If
        
        sStr = PegarOpcaoINI("CONEXAO", "DATABASE", "")

        DBEngine.IniPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Jet\3.5\ISAM Formats\Text"
        Set DB = DBEngine.OpenDatabase(sStr, False)

        DB.Execute Proc_Inserir.InsereArquivoAD(App.path)
        DB.Close

        '''''''''''''''''''''
        'Atualiza alterações'
        '''''''''''''''''''''
        g_cMainConnection.CommitTrans

        ''''''''''''''''''''''''''''''
        'Imprime o Aviso de Diferença'
        ''''''''''''''''''''''''''''''
        Principal.CrystalReport.ReportFileName = App.path & "\Reports\RelAvisoDiferenca.rpt"
        Principal.CrystalReport.SelectionFormula = "{AvisoDiferenca.Num_Bordero} = """ & lstBordero.List(lstBordero.ListIndex) & """"
        Principal.CrystalReport.WindowState = crptMaximized
        Principal.CrystalReport.WindowTitle = "Aviso de Diferença"
        Principal.CrystalReport.PrintReport

    End If
    On Error GoTo Erro_GerarAD

    ''''''''''''''''''''''''''''''''''''''''''''''''''
    'Calcula as Datas de Deposito que estão indevidas'
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    If Not clCB.CalculaChequesIndevidosDATA(lQtdDatasIndevidos) Then
        Set clCB = Nothing
        GoTo Erro_GerarAD
    End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    If Not clCB.Calcula Then

        If MsgBox("Este Borderô ainda apresenta irregularidades." & Chr(10) & _
                  "Deseja gerar o Aviso de Diferença?", vbExclamation + vbYesNo, Me.Caption) = vbNo Then
            '''''''''''''''''''''''''''''''''
            'Mantém o Status do Borderô em 5'
            '''''''''''''''''''''''''''''''''
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        Screen.MousePointer = vbHourglass

        '''''''''''''''''''''''''''''''''
        'Obtem a Sequencia da Ocorrencia'
        '''''''''''''''''''''''''''''''''
        
        lSeqOcorrencia = Format(CLng(Right(g_Parametros.CNPJ_Terceira, 9)) + 10000000, "000000000")
        
        If g_Parametros.Seq_Ocorrencia = 0 Or (g_Parametros.Seq_Ocorrencia > lSeqOcorrencia) Then
            lSeqOcorrencia = Format(CLng(Right(g_Parametros.CNPJ_Terceira, 9)), "000000000")
        Else
            lSeqOcorrencia = Format(g_Parametros.Seq_Ocorrencia, "000000000")
        End If


        '''''''''''''''''''''''''''''''''''
        'Localiza as datas com divergencia'
        '''''''''''''''''''''''''''''''''''
        On Error GoTo Erro_AtualizaBordero
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''
        'Obtem a descrição do Motivo do Aviso de Diferenca'
        '''''''''''''''''''''''''''''''''''''''''''''''''''
        Set rstDescMotAD = g_cMainConnection.Execute(Proc_Selecionar.GetMotivoAD(4))
        sMotivoAD = ""
        If Not rstDescMotAD.EOF Then
            sMotivoAD = rstDescMotAD!Descricao
        End If
        
        Set rstDescMotAD = g_cMainConnection.Execute(Proc_Selecionar.GetMotivoAD(5))
        sMotivoAD2 = ""
        If Not rstDescMotAD.EOF Then
            sMotivoAD2 = rstDescMotAD!Descricao
        End If
        rstDescMotAD.Close
        
        g_cMainConnection.BeginTrans

        For i = 1 To clCB.DataDeposito.Count
            If clCB.DataDeposito(i).DataDivergente Then

                lSeqOcorrencia = Format(lSeqOcorrencia + 1, "000000000")
                '''''''''''''''''''''''''''''
                'Insere o Aviso de Diferenca'
                '''''''''''''''''''''''''''''

                Call g_cMainConnection.Execute(Proc_Inserir.InsereAvisoDiferenca( _
                                               m_DataProcessamento, _
                                               lSeqOcorrencia, _
                                               sMotivoAD, _
                                               clCB.DataDeposito(i).DataDeposito, _
                                               lstBordero.List(lstBordero.ListIndex), _
                                               rst!CodigoCarteira, _
                                               0, _
                                               0, _
                                               2, _
                                               0, _
                                               0, _
                                               0, _
                                               0, _
                                               0, _
                                               0, _
                                               0, _
                                               0, _
                                               Abs(clCB.DataDeposito(i).DiferencaValor), _
                                               -1, "R"), _
                                        lRetorno, _
                                        adCmdText)
                If lRetorno = 0 Then
                    Set clCB = Nothing
                    GoTo Erro_AtualizaBordero
                End If

                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'Não gerar AD de cheques quando diferenca for positiva, ou seja'
                'bordero indicando R$100 e so existe um cheque de R$50
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If Val(clCB.DataDeposito(i).DiferencaValor) < 0 Then
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''
                    'Calcula os cheques que estao com Aviso de Diferenca'
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''
                    clCB.DataDeposito(i).ValorAD = Abs(clCB.DataDeposito(i).DiferencaValor)
                    If Not clCB.DataDeposito(i).CalculaChequesAD(g_cMainConnection) Then
                        Set clCB = Nothing
                        GoTo Erro_AtualizaBordero
                        End If
    
                   ' lSeqOcorrencia = lSeqOcorrencia + 1
                    lSeqOcorrencia = Format(lSeqOcorrencia + 1, "000000000")
                    For j = 1 To clCB.DataDeposito(i).Cheque.Count
                        
                        '''''''''''''''''''''''''''''''''''''''''''''
                        'Calcula o cheque para obter Agencia e Conta'
                        '''''''''''''''''''''''''''''''''''''''''''''
                        clCB.DataDeposito(i).Cheque.Item(j).Calcula
                        
                        If clCB.DataDeposito(i).Cheque.Item(j).Status = "AD" Then
                            '''''''''''''''''''''''''''''''''''''''''''''''''
                            'Insere o Aviso de Diferenca informando o cheque'
                            '''''''''''''''''''''''''''''''''''''''''''''''''
                            Call g_cMainConnection.Execute(Proc_Inserir.InsereAvisoDiferenca( _
                                                           m_DataProcessamento, _
                                                          lSeqOcorrencia, _
                                                           sMotivoAD2, _
                                                           clCB.DataDeposito(i).DataDeposito, _
                                                           lstBordero.List(lstBordero.ListIndex), _
                                                           rst!CodigoCarteira, _
                                                           0, _
                                                           0, _
                                                           2, _
                                                           clCB.DataDeposito(i).Cheque.Item(j).Comp, _
                                                           clCB.DataDeposito(i).Cheque.Item(j).Banco, _
                                                           clCB.DataDeposito(i).Cheque.Item(j).Agencia, _
                                                           clCB.DataDeposito(i).Cheque.Item(j).Conta, _
                                                           clCB.DataDeposito(i).Cheque.Item(j).NumeroCheque, _
                                                           clCB.DataDeposito(i).Cheque.Item(j).Tipificacao, _
                                                           0, _
                                                           0, _
                                                           Abs(clCB.DataDeposito(i).Cheque.Item(j).Valor), _
                                                           -1, "R"), _
                                                    lRetorno, _
                                                    adCmdText)
    
                            If lRetorno = 0 Then
                                Set clCB = Nothing
                                GoTo Erro_AtualizaBordero
                            End If
                        End If
                    
                    Next j
                End If

                '''''''''''''''''''''''''''''''
                'Corrige a tabela DataDeposito'
                '''''''''''''''''''''''''''''''
                Call g_cMainConnection.Execute(Proc_Alterar.AtualizaDataDeposito( _
                                               m_DataProcessamento, _
                                               m_IdBordero, _
                                               clCB.DataDeposito(i).DataDeposito, _
                                               clCB.DataDeposito(i).DataDeposito, _
                                               clCB.DataDeposito(i).SomatoriaQuantidadeCheques, _
                                               InserePonto(RetiraPonto(clCB.DataDeposito(i).SomatoriaValoresCheques))), _
                                        lRetorno, _
                                        adCmdText)
                If lRetorno = 0 Then
                    Set clCB = Nothing
                    GoTo Erro_AtualizaBordero
                End If
            End If
        Next i

        '''''''''''''''''''''''''''''
        'Atualiza a tabela parametro'
        '''''''''''''''''''''''''''''
        g_Parametros.Seq_Ocorrencia = Format(lSeqOcorrencia, "000000000")

        Call g_cMainConnection.Execute(Proc_Alterar.AtualizaParametros( _
                                       m_DataProcessamento, _
                                       g_Parametros.QuantidadeCheques, _
                                       g_Parametros.QuantidadeDatas, _
                                       g_Parametros.DiretorioTransmissao, _
                                       g_Parametros.DiretorioRecepcao, _
                                       g_Parametros.Codigo_USB, _
                                       g_Parametros.CodigoAgAcolhed, _
                                       g_Parametros.CPD_Origem, _
                                       g_Parametros.CPD_Destino, _
                                       g_Parametros.Codigo_Terceira, _
                                       g_Parametros.CNPJ_Terceira, _
                                       g_Parametros.Seq_Ocorrencia, _
                                       g_Parametros.UF_Terceira, _
                                       g_Parametros.CodigoAplicacao, _
                                       g_Parametros.ValorChequeLimite, _
                                       g_Parametros.HeaderAV, g_Parametros.chkSoma, _
                                       g_Parametros.Gerar_Arquivo_CEL, _
                                       g_Parametros.Comp_Origem_CEL, _
                                       g_Parametros.Numero_Versao_Inicial_CEL, _
                                       g_Parametros.Numero_Versao_Final_CEL, _
                                       g_Parametros.QuantidadeMinimaDias, _
                                       g_Parametros.Cidade_Terceira, _
                                       g_Parametros.Nome_Terceira), _
                                lRetorno, adCmdText)
        If lRetorno = 0 Then
            Set clCB = Nothing
            GoTo Erro_AtualizaBordero
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Corrige a tabela Bordero                             '
        'Para não ter que criar um novo type ou clará-lo      '
        'no Globais, estou fazendo um novo select para        '
        'atualizar a tabela Bordero com os campos selecionados'
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Set rst = g_cMainConnection.Execute(Proc_Selecionar.GetIdBordero( _
                                            m_DataProcessamento, _
                                            m_IdBordero))
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Calcula as datas corretas de acordo com a tabela de cheques'
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        clSD.SetConnection g_cMainConnection
        clSD.DataProcessamento = m_DataProcessamento
        clSD.IdBordero = m_IdBordero

        clSD.Calcula

        Call g_cMainConnection.Execute(Proc_Alterar.AtualizaBordero( _
                                       m_DataProcessamento, _
                                       m_IdBordero, _
                                       lstBordero.List(lstBordero.ListIndex), _
                                       rst!Agencia, _
                                       rst!Conta, _
                                       rst!CodigoCarteira, _
                                       rst!CodigoLoja, _
                                       rst!DataEntrada, _
                                       rst!NomeCliente, _
                                       clSD.SomatoriaDatas, _
                                       clSD.SomatoriaQuantidades, _
                                       clSD.SomatoriaValores, _
                                       clSD.SomatoriaControle), _
                                lRetorno, _
                                adCmdText)
        If lRetorno = 0 Then
            Set clSD = Nothing
            GoTo Erro_AtualizaBordero
        End If
        rst.Close

        '''''''''''''''''''''''''''''''''''''''''''''
        'Antes de gerar o Aviso de Diferenca, EXCLUI'
        '''''''''''''''''''''''''''''''''''''''''''''
'        Call g_cMainConnection.Execute(Proc_Excluir.RemoveAvisoDiferenca( _
'                                       m_DataProcessamento, _
'                                       m_IdBordero), _
'                                lRetorno, _
'                                adCmdText)

        '''''''''''''''''''''''''''
        'Gera o Aviso de Diferenca'
        '''''''''''''''''''''''''''
'        Call g_cMainConnection.Execute(Proc_Inserir.InsereAvisoDiferenca( _
'                                       m_DataProcessamento, _
'                                       m_IdBordero, _
'                                       0, _
'                                       0, _
'                                       0, _
'                                       0, _
'                                       0), _
'                                lRetorno, _
'                                adCmdText)
        If lRetorno = 0 Then
            GoTo Erro_AtualizaBordero
        End If

        g_cMainConnection.CommitTrans

        ''''''''''''''''''''''''''''
        'Atualiza Status do Borderô'
        ''''''''''''''''''''''''''''
        Call g_cMainConnection.Execute(Proc_Alterar.AtualizaStatusBordero( _
                                       m_DataProcessamento, _
                                       m_IdBordero, _
                                       "R"), _
                                lRetorno, _
                                adCmdText)
        If lRetorno = 0 Then
            Screen.MousePointer = vbDefault
            MsgBox "Não foi possível atualizar o Status do Borderô.", vbExclamation, Me.Caption
            Exit Sub
        End If

        On Error GoTo Erro_GerarAD
        ''''''''''''''''''''''''
        'Imprime o novo bordero'
        ''''''''''''''''''''''''
        iOldDestination = Principal.CrystalReport.Destination

        Principal.CrystalReport.ReportFileName = App.path & "\Reports\RelGerenciamentoCheques.rpt"
        Principal.CrystalReport.SelectionFormula = "{Bordero.IdBordero} = " & m_IdBordero
        Principal.CrystalReport.WindowState = crptMaximized
        Principal.CrystalReport.WindowTitle = "Borderô de Gerenciamento de Cheques"
        Principal.CrystalReport.CopiesToPrinter = 3
        Principal.CrystalReport.Destination = crptToPrinter
        Principal.CrystalReport.Action = 0
        ''''''''''''''''''''''''''''''
        'Imprime o Aviso de Diferença'
        ''''''''''''''''''''''''''''''
        Principal.CrystalReport.ReportFileName = App.path & "\Reports\RelAvisoDiferenca.rpt"
        Principal.CrystalReport.SelectionFormula = "{AvisoDiferenca.Num_Bordero} = '" & lstBordero.List(lstBordero.ListIndex) & "' AND " & _
                                                   "{AvisoDiferenca.DataOcorrencia} = " & m_DataProcessamento
        Principal.CrystalReport.WindowState = crptMaximized
        Principal.CrystalReport.WindowTitle = "Aviso de Diferença"
        Principal.CrystalReport.CopiesToPrinter = 3
        Principal.CrystalReport.Destination = crptToPrinter
        Principal.CrystalReport.PrintReport

        Principal.CrystalReport.Destination = iOldDestination

    End If

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'ESTÁ TUDO BATIDINHO, BORDERO NÃO APRESENTA IRREGULARIDADES'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'    Call g_cMainConnection.Execute(Proc_Alterar.AtualizaStatusBordero( _
'                                   m_DataProcessamento, _
'                                   m_IdBordero, _
'                                   IIf(m_StatusAnterior = "X", "C", "R")), _
'                           lRetorno, _
'                           adCmdText)

    ' Atualiza Status do Borderô para 'R'
    Call g_cMainConnection.Execute(Proc_Alterar.AtualizaStatusBordero( _
                                   m_DataProcessamento, _
                                   m_IdBordero, _
                                   "R"), _
                           lRetorno, _
                           adCmdText)
                           
    
    
    ' Atualiza Status das dadtas do Borderô
    Call g_cMainConnection.Execute(Proc_Alterar.AtualizaStatusDataDeposito( _
                                   m_DataProcessamento, _
                                   m_IdBordero), _
                           lRetorno, _
                           adCmdText)


    ' Atualiza Status dos cheques do Borderô
    Call g_cMainConnection.Execute(Proc_Alterar.AtualizaStatusCheques( _
                                   m_DataProcessamento, _
                                   m_IdBordero, "C"), _
                           lRetorno, _
                           adCmdText)

    If lRetorno = 0 Then
        MsgBox "Não foi possível encerrar este Borderô.", vbExclamation, Me.Caption
    Else
        m_Bordero_Tratado(lstBordero.ListIndex) = True
        MsgBox "Borderô encerrado.", vbExclamation, Me.Caption
        '''''''''''''''''''''''''''''
        'Refazer a lista de borderos'
        '''''''''''''''''''''''''''''
'            If Not PreencheLstBordero() Then
'                LimpaGridDatas
'                LimpaGridCheques
'                LimpaGridRejeicoes
'
'                lstBordero.Clear
'                ''''''''''''''''''''''''''''''''''''''
'                'Aguarda novo borderô para Supervisor'
'                ''''''''''''''''''''''''''''''''''''''
'                AguardaDocumento
'            End If
        '''''''''''''''''''''''''''''
        'Seleciona o próximo Borderô'
        '''''''''''''''''''''''''''''
        m_IdBordero_Atual = 0
        If Not SelecionaProximoBordero() Then
            LimpaGridDatas
            LimpaGridCheques
            LimpaGridRejeicoes

            lstBordero.Clear
            '''''''''''''''''''''''''''''
            'Refazer a lista de borderos'
            '''''''''''''''''''''''''''''
            If Not PreencheLstBordero() Then
                ''''''''''''''''''''''''''''''''''''''
                'Aguarda novo borderô para Supervisor'
                ''''''''''''''''''''''''''''''''''''''
                AguardaDocumento
            End If

        End If
    End If

    Screen.MousePointer = vbDefault
    
    Exit Sub

Erro_GerarAD:
    Screen.MousePointer = vbDefault
    Call TratamentoErro("Não foi possível finalizar o procedimento de geração de Aviso de Diferença.", Err)
    Exit Sub

    
Erro_AtualizaBordero:

    Screen.MousePointer = vbDefault
    g_cMainConnection.RollbackTrans
    Call TratamentoErro("Não foi possível finalizar o procedimento de atualização do borderô.", Err)
    Exit Sub

Erro_EncerrarBordero:

End Sub

Private Sub cmdExcluirBordero_Click()

    Dim Proc_Alterar        As New Custodia.atualizar
    Dim Proc_Selecionar     As New Custodia.Selecionar
    Dim lRetorno            As Long
    Dim rst                 As New ADODB.Recordset
    
    On Error GoTo Erro_ExcluirBordero

    '''''''''''''''''''''''''''''''''''''''''''
    'Verifica se foi selecionado algum borderô'
    '''''''''''''''''''''''''''''''''''''''''''
    If lstBordero.ListIndex >= 0 Then
        If MsgBox("Confirma a exclusão deste Borderô?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
            Exit Sub
        End If
        
        '''''''''''''''''''''''''''''
        'Respondeu sim, então exclui'
        '''''''''''''''''''''''''''''
        
        ''''''''''''''''''''''''''''''''''''''''''
        'É preciso verificar se existe os cheques'
        'se caso existir cheque, é necessário que'
        'retorne um numero diferente de zero na  '
        'exclusão dos cheques                    '
        ''''''''''''''''''''''''''''''''''''''''''
        
        Set rst = g_cMainConnection.Execute(Proc_Selecionar.GetCheques( _
                                            m_DataProcessamento, _
                                            m_IdBordero))

        g_cMainConnection.BeginTrans
        
        If Not rst.EOF() Then
            
            
            '''''''''''''''''''''''''''''''''''''''
            'Exclui todos os cheques deste borderô'
            '''''''''''''''''''''''''''''''''''''''
            Call g_cMainConnection.Execute(Proc_Alterar.AtualizaStatusCheques( _
                                           m_DataProcessamento, _
                                           m_IdBordero, _
                                           "D"), _
                                    lRetorno, _
                                    adCmdText)
            
            If lRetorno = 0 Then
                g_cMainConnection.RollbackTrans
                MsgBox "Não foi possível excluir os cheques deste borderô.", vbExclamation, Me.Caption
                Exit Sub
            End If
        End If
        
        rst.Close
        
        ''''''''''''''''''
        'Exclui o borderô'
        ''''''''''''''''''
        Call g_cMainConnection.Execute(Proc_Alterar.AtualizaStatusBordero( _
                                       m_DataProcessamento, _
                                       m_IdBordero, _
                                       "D"), _
                                lRetorno, _
                                adCmdText)
        If lRetorno = 0 Then
            g_cMainConnection.RollbackTrans
            MsgBox "Não foi possível excluir este Borderô.", vbExclamation, Me.Caption
            Exit Sub
        End If
        
        g_cMainConnection.CommitTrans
        
        lstBordero.Clear
        LimpaGridDatas
        LimpaGridCheques
        LimpaGridRejeicoes
        '''''''''''''''''''''''''''''
        'Refazer a lista de borderos'
        '''''''''''''''''''''''''''''
        If Not PreencheLstBordero() Then
            ''''''''''''''''''''''''''''''''''''''
            'Aguarda novo borderô para Supervisor'
            ''''''''''''''''''''''''''''''''''''''
            AguardaDocumento
        Else
            '''''''''''''''''''''''''''''
            'Seleciona o próximo Borderô'
            '''''''''''''''''''''''''''''
            m_IdBordero_Atual = 0
            SelecionaProximoBordero
        End If
        
        
    End If
    
    Exit Sub
Erro_ExcluirBordero:

    TratamentoErro "Não foi possível excluir este Borderô.", Err

End Sub

Private Sub cmdGerarAD_Click()

    Dim i                       As Integer
    Dim clB                     As New CalculoBordero
    Dim clSD                    As New SomatoriaDatas
    Dim Proc_Alterar            As New Custodia.atualizar
    Dim Proc_Selecionar         As New Custodia.Selecionar
    Dim Proc_Inserir            As New Custodia.Inserir
    Dim rst                     As New ADODB.Recordset
    Dim lQtdDatasAlteradas      As Long
    Dim lRetorno                As Long
    
    On Error GoTo Erro_GerarAD
    
    Screen.MousePointer = vbHourglass
    
    If lstBordero.ListIndex = -1 Then Exit Sub
    
    If MsgBox("Deseja realmente gerar o aviso de difereça.", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    clB.SetConnection g_cMainConnection
    clB.DataProcessamento = m_DataProcessamento
    clB.IdBordero = m_IdBordero
    
    clB.QuantidadeMaximaCheques = g_Parametros.QuantidadeCheques
    clB.QuantidadeMaximaDatas = g_Parametros.QuantidadeDatas
    
    ''''''''''''''''''''''''''''''''''''''
    'Verifica se existe cheques indevidos'
    ''''''''''''''''''''''''''''''''''''''
    If Not clB.CalculaChequesIndevidosQTDE() Then
        Set clB = Nothing
        GoTo Erro_GerarAD
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''
    'Verifica se existe Datas de Depositos indevidos'
    '''''''''''''''''''''''''''''''''''''''''''''''''
    If Not clB.CalculaChequesIndevidosDATA() Then
        Set clB = Nothing
        GoTo Erro_GerarAD
    End If

    If Not clB.Calcula() Then
        '''''''''''''''''''''''''''''''''''
        'Localiza as datas com divergencia'
        '''''''''''''''''''''''''''''''''''
        On Error GoTo Erro_AtualizaBordero

        g_cMainConnection.BeginTrans
        
        For i = 1 To clB.DataDeposito.Count
            If clB.DataDeposito(i).DataDivergente Then
                '''''''''''''''''''''''''''''''
                'Corrige a tabela DataDeposito'
                '''''''''''''''''''''''''''''''
                Call g_cMainConnection.Execute(Proc_Alterar.AtualizaDataDeposito( _
                                               m_DataProcessamento, _
                                               m_IdBordero, _
                                               clB.DataDeposito(i).DataDeposito, _
                                               clB.DataDeposito(i).DataDeposito, _
                                               clB.DataDeposito(i).SomatoriaQuantidadeCheques, _
                                               InserePonto(RetiraPonto(clB.DataDeposito(i).SomatoriaValoresCheques))), _
                                       lRetorno, _
                                       adCmdText)
                If lRetorno = 0 Then
                    Set clB = Nothing
                    GoTo Erro_AtualizaBordero
                End If
            End If
        Next i
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Corrige a Tabela Bordero                             '
        'Para não ter que criar um novo type ou declará-lo    '
        'no Globais, estou fazendo um novo select para        '
        'atualizar a tabela Bordero com os campos selecionados'
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Set rst = g_cMainConnection.Execute(Proc_Selecionar.GetIdBordero( _
                                            m_DataProcessamento, _
                                            m_IdBordero))
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Calcula as datas corretas de acordo com a tabela de cheques'
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        clSD.SetConnection g_cMainConnection
        clSD.DataProcessamento = m_DataProcessamento
        clSD.IdBordero = m_IdBordero
        
        clSD.Calcula

        Call g_cMainConnection.Execute(Proc_Alterar.AtualizaBordero( _
                                       m_DataProcessamento, _
                                       m_IdBordero, _
                                       lstBordero.List(lstBordero.ListIndex), _
                                       rst!Agencia, _
                                       rst!Conta, _
                                       rst!CodigoCarteira, _
                                       rst!CodigoLoja, _
                                       rst!DataEntrada, _
                                       rst!NomeCliente, _
                                       clSD.SomatoriaDatas, _
                                       clSD.SomatoriaQuantidades, _
                                       clSD.SomatoriaValores, _
                                       clSD.SomatoriaControle), _
                                lRetorno, _
                                adCmdText)

        If lRetorno = 0 Then
            GoTo Erro_AtualizaBordero
        End If

        rst.Close
        '''''''''''''''''''''''''''
        'Gera o Aviso de Diferença'
        '''''''''''''''''''''''''''
        
'        Call g_cMainConnection.Execute(Proc_Inserir.InsereAvisoDiferenca( _
'                                       m_DataProcessamento, _
'                                       m_IdBordero, _
'                                       0, _
'                                       0, _
'                                       0, _
'                                       0, _
'                                       0), _
'                                adCmdText, _
'                                lRetorno)
        
        If lRetorno = 0 Then
            GoTo Erro_AtualizaBordero
        End If
        
        
    
        g_cMainConnection.CommitTrans
    
    Else
        Screen.MousePointer = vbDefault
        MsgBox "Os valores deste Borderô estão batidos, não é permitido gerar Aviso de Diferença.", vbExclamation, Me.Caption
        Exit Sub
    End If
    
    On Error GoTo Erro_GerarAD
    
    ''''''''''''''''''''''''''
    'Atualiza o Grid de Datas'
    ''''''''''''''''''''''''''
    PreencheGridDatas m_DataProcessamento, m_IdBordero
    
    Screen.MousePointer = vbDefault
    MsgBox "Geração de Aviso de Diferença concluída com sucesso.", vbInformation, Me.Caption
    
    '''''''''''''''''''''''''''''
    'Gera a impressao do Borderô'
    '''''''''''''''''''''''''''''
'    Impressao.SetImpressao eBorderoGerenciamentoCheques
'    Impressao.SetSelectionFormula "{Bordero.IdBordero} = " & m_IdBordero
'    Impressao.PrintReport True
    
    ''''''''''''''''''''''''
    'Gera a impressao do AD'
    ''''''''''''''''''''''''
'    Impressao.SetImpressao eAvisoDiferenca
'    'Impressao.SetSelectionFormula "{Bordero.IdBordero} = " & m_IdBordero
'    Impressao.PrintReport True
    
    ''''''''''''''''''''''''''''
    'Atualiza Status do Borderô'
    ''''''''''''''''''''''''''''
    Call g_cMainConnection.Execute(Proc_Alterar.AtualizaStatusBordero( _
                                   m_DataProcessamento, _
                                   m_IdBordero, _
                                   "R"), _
                            lRetorno, _
                            adCmdText)
    If lRetorno = 0 Then
        MsgBox "Não foi possível atualizar o Status do Borderô.", vbExclamation, Me.Caption
        Exit Sub
    End If

    Exit Sub
    
Erro_GerarAD:

    Screen.MousePointer = vbDefault
    Call TratamentoErro("Não foi possível finalizar o procedimento.", Err)
    Exit Sub
    
Erro_AtualizaBordero:
    Screen.MousePointer = vbDefault
    g_cMainConnection.RollbackTrans
    Call TratamentoErro("Não foi possível finalizar o procedimento.", Err)
    
End Sub

Private Sub cmdProvaZero_Click()

    Dim Proc_Alterar        As New Custodia.atualizar
    Dim lRetorno            As Long
    
    On Error GoTo Erro_EnvioProvaZero

    ''''''''''''''''''''''''''''''''''''''''''''
    'Se modo em alteração ja existe m_IdBordero'
    'caso contrario não existe nenhum IdBordero'
    ''''''''''''''''''''''''''''''''''''''''''''
    If lstBordero.ListIndex >= 0 Then
        Call g_cMainConnection.Execute(Proc_Alterar.AtualizaStatusBordero( _
                                       m_DataProcessamento, _
                                       m_IdBordero, _
                                       "4"), _
                               lRetorno, _
                               adCmdText)
                               
                               
         ' Atualiza Status das dadtas do Borderô
        Call g_cMainConnection.Execute(Proc_Alterar.AtualizaStatusDataDeposito( _
                                   m_DataProcessamento, _
                                   m_IdBordero), _
                           lRetorno, _
                           adCmdText)


        ' Atualiza Status dos cheques do Borderô
        Call g_cMainConnection.Execute(Proc_Alterar.AtualizaStatusCheques( _
                                   m_DataProcessamento, _
                                   m_IdBordero, "C"), _
                           lRetorno, _
                           adCmdText)
                               
                               
                               
        If lRetorno = 0 Then
            MsgBox "Não foi possível enviar o Borderô para Prova Zero.", vbExclamation, Me.Caption
            Exit Sub
        End If
        
        m_Bordero_Tratado(lstBordero.ListIndex) = True

        '''''''''''''''''''''''''''''
        'Seleciona o próximo Borderô'
        '''''''''''''''''''''''''''''
        m_IdBordero_Atual = 0
        If Not SelecionaProximoBordero() Then
            LimpaGridDatas
            LimpaGridCheques
            LimpaGridRejeicoes

            lstBordero.Clear
            
            '''''''''''''''''''''''''''''
            'Refazer a lista de borderos'
            '''''''''''''''''''''''''''''
            If Not PreencheLstBordero() Then
                ''''''''''''''''''''''''''''''''''''''
                'Aguarda novo borderô para Supervisor'
                ''''''''''''''''''''''''''''''''''''''
                AguardaDocumento
            End If
        End If
    End If
    
    Exit Sub
Erro_EnvioProvaZero:

    TratamentoErro "Não foi possível enviar o Borderô para Prova Zero.", Err

End Sub


Private Sub cmdSair_Click()
   
    Unload Me
End Sub

Private Sub Form_Activate()

    If m_IsActive Then Exit Sub

    '''''''''''''''''''''''''''''
    'Preenche a lista de Bordero'
    '''''''''''''''''''''''''''''
    If Not PreencheLstBordero() Then
        ''''''''''''''''''''''''''''''''''''''
        'Aguarda novo borderô para Supervisor'
        ''''''''''''''''''''''''''''''''''''''
        m_IsActive = True
        'UpdateWindow Me.hwnd
        Me.Refresh
        DoEvents
        Me.Refresh
        AguardaDocumento
    Else
        m_IsActive = True
        SelecionaBordero lstBordero.ItemData(0)
    End If
    

End Sub

Private Sub Form_Load()

    Dim lWidth              As Long
    
    On Error GoTo Erro_FormLoad
    
    Me.Refresh
    
    ''''''''''''''''''''''''''''''''''''''''
    'Variavel para segurar as cores do Grid'
    ''''''''''''''''''''''''''''''''''''''''
    m_bColor = GrdDatas.BackColorSel
    m_fColor = GrdDatas.ForeColorSel
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''
    'Formata m_DataProcessamento para "yyyymmdd"'
    '''''''''''''''''''''''''''''''''''''''''''''''
    m_DataProcessamento = Geral.DataProcessamento
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Pega largura do scroll vertical para que seja considerado'
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    lWidth = GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX + 10
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                          FORMATA CABEÇALHO DO GRID DE DATAS                                              '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    GrdDatas.Cols = 3
    
    GrdDatas.ColWidth(COL_DT_DATADEPOSITO) = Int((GrdDatas.Width - lWidth) * 0.33)   '  33   %
    GrdDatas.ColWidth(COL_DT_QUANTIDADE) = Int((GrdDatas.Width - lWidth) * 0.33)     '  33   %
    GrdDatas.ColWidth(COL_DT_VALOR) = Int((GrdDatas.Width - lWidth) * 0.34)          '  34   %
    
    GrdDatas.ColAlignment(COL_DT_DATADEPOSITO) = flexAlignLeftCenter
    GrdDatas.ColAlignment(COL_DT_QUANTIDADE) = flexAlignLeftCenter
    GrdDatas.ColAlignment(COL_DT_VALOR) = flexAlignLeftCenter
    
    GrdDatas.TextMatrix(0, COL_DT_DATADEPOSITO) = "Data de Depósito"
    GrdDatas.TextMatrix(0, COL_DT_QUANTIDADE) = "Quantidade"
    GrdDatas.TextMatrix(0, COL_DT_VALOR) = "Valor do Depósito"
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                          FORMATA CABEÇALHO DO GRID DE CHEQUES                                            '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    GrdCheques.Cols = 6
    
    GrdCheques.ColWidth(COL_CH_SEQUENCIAL) = Int((GrdCheques.Width - lWidth) * 0.15)   '  15   %
    GrdCheques.ColWidth(COL_CH_AGENCIA) = Int((GrdCheques.Width - lWidth) * 0.1)       '  10   %
    GrdCheques.ColWidth(COL_CH_CONTA) = Int((GrdCheques.Width - lWidth) * 0.2)         '  20   %
    GrdCheques.ColWidth(COL_CH_NUMEROCHEQUE) = Int((GrdCheques.Width - lWidth) * 0.15) '  15   %
    GrdCheques.ColWidth(COL_CH_CNPJCPF) = Int((GrdCheques.Width - lWidth) * 0.2)       '  20   %
    GrdCheques.ColWidth(COL_CH_VALOR) = Int((GrdCheques.Width - lWidth) * 0.2)         '  20   %
    
    GrdCheques.ColAlignment(COL_CH_SEQUENCIAL) = flexAlignLeftCenter
    GrdCheques.ColAlignment(COL_CH_AGENCIA) = flexAlignLeftCenter
    GrdCheques.ColAlignment(COL_CH_CONTA) = flexAlignLeftCenter
    GrdCheques.ColAlignment(COL_CH_NUMEROCHEQUE) = flexAlignLeftCenter
    GrdCheques.ColAlignment(COL_CH_CNPJCPF) = flexAlignLeftCenter
    GrdCheques.ColAlignment(COL_CH_VALOR) = flexAlignRightCenter
    
    GrdCheques.TextMatrix(0, COL_CH_SEQUENCIAL) = "Nr.Sequencial"
    GrdCheques.TextMatrix(0, COL_CH_AGENCIA) = "Agência"
    GrdCheques.TextMatrix(0, COL_CH_CONTA) = "Conta Corrente"
    GrdCheques.TextMatrix(0, COL_CH_NUMEROCHEQUE) = "Número Cheque"
    GrdCheques.TextMatrix(0, COL_CH_CNPJCPF) = "CNPJ - CPF"
    GrdCheques.TextMatrix(0, COL_CH_VALOR) = "Valor"
    
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                          FORMATA CABEÇALHO DO GRID DE REJEIÇÕES                                          '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    GrdMotivoRejeicao.Cols = 2
    
    GrdMotivoRejeicao.ColWidth(COL_MR_CODIGO) = Int((GrdMotivoRejeicao.Width - lWidth) * 0.15)    '  05   %
    GrdMotivoRejeicao.ColWidth(COL_MR_DESCRICAO) = Int((GrdMotivoRejeicao.Width - lWidth) * 0.85) '  85   %
    
    GrdMotivoRejeicao.ColAlignment(COL_MR_CODIGO) = flexAlignLeftCenter
    GrdMotivoRejeicao.ColAlignment(COL_MR_DESCRICAO) = flexAlignLeftCenter
    
    GrdMotivoRejeicao.TextMatrix(0, COL_MR_CODIGO) = "Código"
    GrdMotivoRejeicao.TextMatrix(0, COL_MR_DESCRICAO) = "Descrição"

   'Inicializa scanner
    Call Principal.SetScanner
    
    Exit Sub
    
Erro_FormLoad:

    TratamentoErro "Não foi possivel iniciar a tela de Supervisor", Err

End Sub


Private Sub Form_Unload(Cancel As Integer)

    Dim Proc_Alterar        As New Custodia.atualizar
    Dim lRetorno            As Long
    Dim sStatus             As String
    
    On Error GoTo Erro_FormUnload

    m_IsActive = False
    ''''''''''''''''''''''''''''''''''''''''''''
    'Volta o Status do antigo borderô se houver'
    ''''''''''''''''''''''''''''''''''''''''''''
    If m_IdBordero_Atual <> 0 Then
    
    
        sStatus = "X"
        
        If m_StatusAnterior = "5" Or _
           m_StatusAnterior = "2" Or _
           m_StatusAnterior = "H" Then sStatus = "5"
           
    
    
        Call g_cMainConnection.Execute(Proc_Alterar.AtualizaStatusBordero( _
                                       m_DataProcessamento, _
                                       m_IdBordero_Atual, _
                                       sStatus), _
                                lRetorno, _
                                adCmdText)

        If lRetorno = 0 Then
            MsgBox "Não foi possível voltar o status do Borderô.", vbExclamation, Me.Caption
            Exit Sub
        End If
    End If
    
    If Not objAguardaDocumento Is Nothing Then
        objAguardaDocumento.Finalizar
    End If
    
    
   'Finaliza Scanner
    Call Principal.DelScanner
    
    Exit Sub
    
Erro_FormUnload:
    TratamentoErro "Erro ao atualizar o Status do Borderô.", Err

End Sub


Private Sub grdCheques_DblClick()

    Dim lRetornoModal       As enumRetornoModal
    Dim sDataDeposito       As String
    
    On Error GoTo Erro_AbreCheque
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    'Não pode adicionar cheque pela tela de Supervisor'
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    If GrdDatas.Rows <= 1 Or GrdCheques.Rows <= 1 Then Exit Sub
    
    ''''''''''''''''''''''''''''''''''''''''''''''''
    'Pega Id Cheque para abertura da tela de Cheque'
    ''''''''''''''''''''''''''''''''''''''''''''''''
    m_IdCheque = GrdCheques.RowData(GrdCheques.Row)
    
    '''''''''''''''''''''''''
    'Pega a Data de Deposito'
    '''''''''''''''''''''''''
    sDataDeposito = Format(GrdDatas.TextMatrix(GrdDatas.Row, COL_DT_DATADEPOSITO), "yyyymmdd")
    If Not IsNumeric(sDataDeposito) Then Exit Sub
    
    Cheque.SetIdCheque m_IdCheque
    
    lRetornoModal = Cheque.ShowModal(m_IdBordero, CLng(sDataDeposito), m_IdCheque)

    If lRetornoModal = eRetornoCancelar Then Exit Sub
    
    PreencheGridCheques m_DataProcessamento, m_IdBordero, sDataDeposito
    
    'If GrdCheques.Rows > 1 Then
        PreencheGridRejeicoes m_DataProcessamento, m_IdBordero, sDataDeposito, m_IdCheque
    'End If
    
    Exit Sub
    
Erro_AbreCheque:
    TratamentoErro "Não foi possível abrir a tela de cheques.", Err

End Sub

Private Sub GrdCheques_EnterCell()

    Dim sDataDeposito       As String

    If GrdCheques.Rows > 1 And GrdCheques.Row > 0 Then
    
        m_IdCheque = GrdCheques.RowData(GrdCheques.Row)
        sDataDeposito = Format(GrdDatas.TextMatrix(GrdDatas.Row, COL_DT_DATADEPOSITO), "yyyymmdd")
        
        ''''''''''''''''''''''''
        'Limpa grid de rejeicao'
        ''''''''''''''''''''''''
        
        LimpaGridRejeicoes
        
        ''''''''''''''''''''''''''''
        'Preenche Grid de Rejeições'
        ''''''''''''''''''''''''''''
        If IsNumeric(sDataDeposito) Then
            PreencheGridRejeicoes m_DataProcessamento, m_IdBordero, sDataDeposito, m_IdCheque
        End If
    End If

End Sub


Private Sub GrdCheques_KeyPress(KeyAscii As Integer)


    '''''''''''''''''''''''''''
    'Responde ao enter no Grid'
    '''''''''''''''''''''''''''
    
    grdCheques_DblClick

End Sub

Private Sub GrdDatas_EnterCell()

    Dim sDataDeposito       As Long
    
    On Error GoTo Erro_EnterCell
    
    If GrdDatas.Rows > 1 And GrdDatas.Row > 0 Then
    
        sDataDeposito = Format(GrdDatas.TextMatrix(GrdDatas.Row, COL_DT_DATADEPOSITO), "yyyymmdd")
        
        '''''''''''''''''''''''''''''''''''
        'Limpa grid de cheques e Rejeicoes'
        '''''''''''''''''''''''''''''''''''
        LimpaGridCheques
        LimpaGridRejeicoes
        
        GrdDatas.Refresh
        DoEvents
        ''''''''''''''''''''''''''
        'Preenche Grid de Cheques'
        ''''''''''''''''''''''''''
        PreencheGridCheques m_DataProcessamento, m_IdBordero, sDataDeposito
        If GrdCheques.Rows > 1 Then
            m_IdCheque = GrdCheques.RowData(GrdCheques.Row)
            ''''''''''''''''''''''''''''
            'Preenche Grid de Rejeições'
            ''''''''''''''''''''''''''''
            PreencheGridRejeicoes m_DataProcessamento, m_IdBordero, sDataDeposito, GrdCheques.RowData(GrdCheques.Row)
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Joga o foco pro grid de cheques para poder teclar enter'
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        GrdCheques.SetFocus
    End If
    
    Exit Sub
    
Erro_EnterCell:
    
    TratamentoErro "Erro ao preencher grid de cheques.", Err

End Sub




Private Sub lstBordero_Click()

    Dim Proc_Alterar        As New Custodia.atualizar
    Dim Proc_Selecionar     As New Custodia.Selecionar
    Dim sDataDeposito       As String
    Dim sStatus             As String
    Dim sMsg                As String
    Dim lRetorno            As Long
    Dim rst                 As New ADODB.Recordset
    
    On Error GoTo Erro_AbreBordero
    
    If m_IsEvent Then Exit Sub
    m_IsEvent = True
    
    m_IdBordero = lstBordero.ItemData(lstBordero.ListIndex)
    
    ''''''''''''''''''''''''''''''''''''''''''''
    'Volta o Status do antigo borderô se houver'
    ''''''''''''''''''''''''''''''''''''''''''''
    If (m_IdBordero_Atual <> 0) Then
    
        sStatus = "X"
        
        If m_StatusAnterior = "2" Or _
           m_StatusAnterior = "5" Or _
           m_StatusAnterior = "H" Then sStatus = "5"
        
        Call g_cMainConnection.Execute(Proc_Alterar.AtualizaStatusBordero(m_DataProcessamento, m_IdBordero_Atual, sStatus), lRetorno, adCmdText)
        If lRetorno = 0 Then
            MsgBox "Não foi possível voltar o status do Borderô.", vbExclamation, Me.Caption
            GoTo Fim_Sub
        End If
    End If
    
    ''''''''''''''''
    'Limpa os grids'
    ''''''''''''''''
    
    LimpaGridDatas
    LimpaGridCheques
    LimpaGridRejeicoes
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'pega o bordero, verifica se ainda está "Para Supervisor"'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set rst = g_cMainConnection.Execute(Proc_Selecionar.GetIdBordero( _
                                        m_DataProcessamento, _
                                        m_IdBordero))

    If Not rst.EOF() Then
        If EstaPara(rst!Status, sMsg, "5", "X", "2") Then
            If rst!Status = "H" Or rst!Status = "J" Then
                '''''''''''''''''''''''''''''''''''
                'Verifica se está em outra estacao'
                '''''''''''''''''''''''''''''''''''
                If DateDiff("s", rst!HoraAtual, Time) <= g_Parametros.TMP_Pendente Then
                    ''''''''''''''''''
                    'Continua travada'
                    ''''''''''''''''''
                    MsgBox "Não permitido." & Chr(10) & "Este borderô se encontra " & Trim(sMsg), vbExclamation, Me.Caption
                    m_IdBordero = 0
                    m_IdBordero_Atual = 0
                    lstBordero.ListIndex = -1
                    GoTo Fim_Sub
                End If
            Else
                ''''''''''''''''''''''''''''''''''
                'Está com outra tela da aplicacao'
                ''''''''''''''''''''''''''''''''''
                MsgBox "Não permitido." & Chr(10) & "Este borderô se encontra " & Trim(sMsg), vbExclamation, Me.Caption
                m_IdBordero = 0
                m_IdBordero_Atual = 0
                
                ''''''''''''''''''''''''''
                'Remove este item do list'
                ''''''''''''''''''''''''''
                'lstBordero.RemoveItem lstBordero.ListIndex
                
                lstBordero.ListIndex = -1
                GoTo Fim_Sub
            End If
        End If
    End If


    '''''''''''''''''''''''''''''''''''''''''''''''''''
    'Atualiza o Status do Borderô para "Em Supervisor"'
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    sStatus = "J"
    
    If rst!Status = "2" Or _
       rst!Status = "5" Or _
       rst!Status = "H" Then sStatus = "H"

    Call g_cMainConnection.Execute(Proc_Alterar.AtualizaStatusBordero(m_DataProcessamento, m_IdBordero, sStatus), lRetorno, adCmdText)
    If lRetorno = 0 Then
        MsgBox "Não foi possível atualizar o status do Borderô.", vbExclamation, Me.Caption
        GoTo Fim_Sub
    End If
    ''''''''''''''''''''''''''
    'Define o IdBordero atual'
    ''''''''''''''''''''''''''
    m_IdBordero_Atual = m_IdBordero
    
    ''''''''''''''''''''''''''
    'Define o Status anterior'
    ''''''''''''''''''''''''''
    m_StatusAnterior = rst!Status
    
    '''''''''''''''''''''''''''''''''''''''
    'Se status 'X', desabilitar Prova Zero'
    '''''''''''''''''''''''''''''''''''''''
    cmdProvaZero.Enabled = Not CBool(m_StatusAnterior = "X")
    
    ''''''''''''''''''''''''
    'Preenche grid de datas'
    ''''''''''''''''''''''''
    PreencheGridDatas m_DataProcessamento, m_IdBordero
    
    If GrdDatas.Rows = 1 Then
        GoTo Fim_Sub
    End If
    sDataDeposito = Format(GrdDatas.TextMatrix(GrdDatas.Row, COL_DT_DATADEPOSITO), "yyyymmdd")
    ''''''''''''''''''''''''''
    'Preenche Grid de Cheques'
    ''''''''''''''''''''''''''
    If IsNumeric(sDataDeposito) Then
        PreencheGridCheques m_DataProcessamento, m_IdBordero, sDataDeposito
        ''''''''''''''''''''''''''''
        'Preenche Grid de Rejeições'
        ''''''''''''''''''''''''''''
        If GrdCheques.Rows > 1 Then
            PreencheGridRejeicoes m_DataProcessamento, m_IdBordero, sDataDeposito, GrdCheques.RowData(GrdCheques.Row)
        End If
    End If
    
Fim_Sub:
    m_IsEvent = False
    Exit Sub
    
Erro_AbreBordero:
    m_IsEvent = False
    TratamentoErro "Erro ao abrir o Borderô.", Err
    
End Sub



Private Sub lstBordero_DblClick()

    Dim lRetornoModal       As enumRetornoModal
    Dim sDataDeposito       As String
    Dim Proc_Alterar        As New Custodia.atualizar
    Dim lRetorno            As Long
    
    On Error GoTo Erro_AbreBordero

    '''''''''''''''''''''''''''''''''''''''''''''''''''
    'Pega o IdBordero para abertura da tela de Borderô'
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    m_IdBordero = lstBordero.ItemData(lstBordero.ListIndex)


    Bordero.SetIdbordero m_IdBordero
    
    lRetornoModal = Bordero.ShowModal()
    
    If lRetornoModal = eRetornoCancelar Then Exit Sub
    
    '''''''''''''''''''''''''''''''''
    'Retira o status "Em Supervisor"'
    '''''''''''''''''''''''''''''''''
    If m_IdBordero_Atual <> 0 Then
        Call g_cMainConnection.Execute(Proc_Alterar.AtualizaStatusBordero(m_DataProcessamento, m_IdBordero_Atual, "5"), lRetorno, adCmdText)
        If lRetorno = 0 Then
            MsgBox "Não foi possível voltar o status do Borderô.", vbExclamation, Me.Caption
            Exit Sub
        End If
        m_IdBordero_Atual = 0
    End If
    
    
    '''''''''''''''''''''''''''''
    'Refazer a lista de borderos'
    '''''''''''''''''''''''''''''
    If Not PreencheLstBordero() Then
        ''''''''''''''''''''''''''''''''''''''
        'Aguarda novo borderô para Supervisor'
        ''''''''''''''''''''''''''''''''''''''
        AguardaDocumento
        Exit Sub
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'SelecionaBordero ja chama as rotinas de preenchimento dos grids'
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    SelecionaBordero m_IdBordero
    
    Exit Sub
    
Erro_AbreBordero:
    
    TratamentoErro "Erro ao carregar o Borderô.", Err

End Sub


