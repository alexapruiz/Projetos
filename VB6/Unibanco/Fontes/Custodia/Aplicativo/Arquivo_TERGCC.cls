VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Arquivo_TERGCC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum enumTipoExportacaoTER
    TerGCC
End Enum


'---------------------------------------------------------------------------------------------
'                                    Criação do TER.GCC
'---------------------------------------------------------------------------------------------
Private Type tpHeaderArquivo_TerGCC
    IdentHeader                 As String * 2       ' "HD"
    CodigoTerceiro              As String * 4       ' "TPS1"
    DataMovto                   As String * 8       ' AAAAMMDD
    NumRemessa                  As String * 4       ' Número da Remessa
    DataGeracao                 As String * 8       ' AAAAMMDD
    HoraGeracao                 As String * 6       ' HHMMSS
    Filler                      As String * 111     ' Brancos
    SequencialArquivo           As String * 7       ' Fixo "0000001"
    CrLf                        As String * 2       ' OK
End Type


Private Type tpTraillerArquivo_TerGCC
    IdentTrailler               As String * 2   ' "TR"
    CodigoTerceiro              As String * 4   ' "TPS1"
    DataMovto                   As String * 8   ' AAAAMMDD
    NumRemessa                  As String * 4   ' Número da Remessa ( Igual Header)
    TotalLimiteQtd              As String * 6   ' Quantidade total de cheques limites
    TotalLimiteValor            As String * 16  ' Valor total de cheques limite
    TotalsuperiorQtd            As String * 6   ' Quantidade total de cheques superior
    TotalsuperiorValor          As String * 16  ' Valor total de cheques superior
    Filler_1                    As String * 12
    Filler_2                    As String * 69
    SequencialArquivo           As String * 7   'Sequencia desde o inicio
End Type


Private Type tpDetalheArquivo_TerGCC
    Tipo                        As String * 2       '"DE"
    AgenciaAutenticacao         As String * 5       ' Agência da autemnticação
    Filler                      As String * 3       ' Brancos
    Filler_2                    As String * 6       ' Brancos
    CMC7                        As String * 42       ' CMC7 do cheque capturado
    ValorProc                   As String * 16      ' Valor processado
    AgeCredito                  As String * 4       ' Agencia de crédito
    CtaCredito                  As String * 7       ' Agencia de crédito
    Filler_3                    As String * 12      ' Brancos
    NumVerCEL                   As String * 4       ' Número da Versão CEL
    NumLoteCEL                  As String * 7       ' Número do Lote CEL
    Filler_4                    As String * 35      ' Brancos
    SequencialArquivo           As String * 7       'Sequencia anterior + 1
    CrLf                        As String * 2       '
End Type



Private Type tpArquivo_TerGCC
    Header                      As tpHeaderArquivo_TerGCC
    Detalhes                    As tpDetalheArquivo_TerGCC
    Trailler                    As tpTraillerArquivo_TerGCC
End Type


Private m_ArquivoTER_GCC        As tpArquivo_TerGCC
Private m_DataProcessamento     As Long
Private m_ProgressBar           As Object
Private m_DiretorioSaida        As String
Private m_ArquivoSaida          As String
Private m_Connection            As New ADODB.Connection
Private m_DataTroca             As Long
Private m_NumRemessa            As Integer
Private m_NovaRemessa           As Boolean
Private m_Inidk                 As String
Private m_Inik7                 As String
Private m_hdx                   As String


Public Property Let ArquivoSaida(ByVal pArquivoSaida As String)
    m_ArquivoSaida = pArquivoSaida
End Property

Public Property Let DataTroca(ByVal pDataTroca As Long)
    m_DataTroca = pDataTroca
End Property

Public Property Get DataTroca() As Long
    DataTroca = m_DataTroca
End Property

Public Property Let NumRemessa(ByVal pNumRemessa As Long)
    m_NumRemessa = pNumRemessa
End Property

Public Property Get NumRemessa() As Long
    NumRemessa = m_NumRemessa
End Property

Public Property Let NovaRemessa(ByVal pNovaRemessa As Boolean)
    m_NovaRemessa = pNovaRemessa
End Property

Public Property Get NovaRemessa() As Boolean
    NovaRemessa = m_NovaRemessa
End Property


Public Property Let DataProcessamento(ByVal pDataProcessamento As Long)
    m_DataProcessamento = pDataProcessamento
End Property

Public Property Get DataProcessamento() As Long
    DataProcessamento = m_DataProcessamento
End Property

Public Property Get ArquivoSaida() As String
    ArquivoSaida = m_ArquivoSaida
End Property

Public Property Let DiretorioSaida(ByVal pDiretorioSaida As String)
    m_DiretorioSaida = pDiretorioSaida & IIf(Right(pDiretorioSaida, 1) <> "\", "\", "")
End Property

Public Property Get DiretorioSaida() As String
    DiretorioSaida = m_DiretorioSaida
End Property


Public Sub SetConnection(ByVal pConnection As ADODB.Connection)
    Set m_Connection = pConnection
End Sub

Public Sub SetProgressBar(ByRef pProgressBar As Object)
    Set m_ProgressBar = pProgressBar
End Sub

Public Property Let TipoArquivo(ByVal pTipoArquivo As enumTipoArquivoCEL)
    m_TipoArquivoCEL = pTipoArquivo
End Property

Public Property Get TipoArquivo() As enumTipoArquivoCEL
    TipoArquivo = m_TipoArquivoCEL
End Property


Public Property Let TipoExportacao(ByVal pTipoExportacao As enumTipoExportacaoCEL605)
    m_TipoExportacao = pTipoExportacao
End Property

Public Property Get TipoExportacao() As enumTipoExportacaoCEL605
    TipoExportacao = m_TipoExportacao
End Property


' Geração do TER.GCC
Public Function Gera_TERGCC() As Boolean

    Dim sstr                As String
    Dim rstDetalhes         As New ADODB.Recordset
    Dim Proc_Selecionar     As New Custodia.Selecionar
    Dim Proc_Atualizar      As New Custodia.Atualizar
    'Dim cCheque             As New clsCheque
    Dim cCheque             As New CalculoCheque
    Dim iFile               As Integer
    Dim dSequencia          As Double
    Dim lRetorno            As Long
    Dim dTotalLimiteValor   As Double
    Dim dTotalLimiteQtd     As Double
    
    Dim dTotalsuperiorValor As Double
    Dim dTotalsuperiorQtd   As Double
    Dim iNumRemessa         As Integer
    
    Dim bTransacaoAberta    As Boolean
    Dim bGeracao            As Boolean
    
    On Error GoTo Erro_Gerar
    
    bTransacaoAberta = False
    Gera_TERGCC = False
    Screen.MousePointer = vbHourglass
    
    Set rstDetalhes = m_Connection.Execute(Proc_Selecionar.GetDetalhesArquivo_TerGCC(m_DataTroca, m_NumRemessa))
    
    If rstDetalhes.EOF Then
        Screen.MousePointer = vbDefault
        MsgBox "Não Existem Cheques para Gerar o Arquivo TERGCC!", vbInformation + vbOKOnly
        Exit Function
    End If
         
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                         CRIAÇÃO E ABERTURA DO ARQUIVO TER_GCC
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If NumRemessa <> 0 Then
        If NovaRemessa Then
            iNumRemessa = NumRemTerGCC()
            m_ArquivoSaida = Format(g_Parametros.Codigo_Terceira, "0000") + Format(iNumRemessa, "0000") & ".GCC"
            bGeracao = True
        Else
            iNumRemessa = m_NumRemessa
            bGeracao = False
        End If
    Else
        iNumRemessa = NumRemTerGCC()
        bGeracao = True
    End If
    
    
    If FileExist(m_DiretorioSaida & m_ArquivoSaida) Then
       Kill m_DiretorioSaida & m_ArquivoSaida
    End If

    iFile = FreeFile
    Open m_DiretorioSaida & m_ArquivoSaida For Binary As #iFile


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                               HEADER DO ARQUIVO                                      '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    If g_Parametros.HeaderAV Then
      
      m_Inidk = F_Inidk(iNumRemessa)
      m_Inik7 = F_Inik7(iNumRemessa)
      m_hdx = F_HDX(iNumRemessa)
    
      Put #iFile, , m_Inidk
      Put #iFile, , m_Inik7
      Put #iFile, , m_hdx
      
    End If
    
    With m_ArquivoTER_GCC.Header
    
        .IdentHeader = "HD"
        .CodigoTerceiro = Format(g_Parametros.Codigo_Terceira, "0000")
        .DataMovto = m_DataTroca
        .NumRemessa = Format(iNumRemessa, "0000")
        .DataGeracao = Format(Date, "YYYYMMDD")
        .HoraGeracao = Format(Now, "HHMMSS")
        .Filler = Space(111)
        .SequencialArquivo = FormataString(Sequencia(dSequencia), "0", Len(.SequencialArquivo), True)
        .CrLf = vbCrLf
        
    End With
    
    Put #iFile, , m_ArquivoTER_GCC.Header

    
    dTotalsuperiorValor = 0
    dTotalsuperiorQtd = 0
    dTotalLimiteValor = 0
    dTotalLimiteQtd = 0

    g_cMainConnection.BeginTrans
    bTransacaoAberta = True
    
    Do While Not rstDetalhes.EOF
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                               DETALHES DO ARQUIVO
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
           
            '''''''''''''''''''''''''''''''''''''''
            'Incrementa o ProgressBar do Principal'
            '''''''''''''''''''''''''''''''''''''''
            m_ProgressBar.Value = rstDetalhes.AbsolutePosition / rstDetalhes.RecordCount * 10
                             
            cCheque.CMC7 = rstDetalhes!CMC7
            cCheque.Calcula
            
            
            With m_ArquivoTER_GCC.Detalhes
                .Tipo = "DE"
                .AgenciaAutenticacao = Format(g_Parametros.CodigoAgAcolhed, "00000")
                .Filler = Space(3)
                .Filler_2 = Space(6)
                .CMC7 = rstDetalhes!CMC7 & Space(12)
                .ValorProc = FormataString(rstDetalhes!Valor * 100, "0", Len(.ValorProc), True)
                .AgeCredito = Format(IIf(IsNull(rstDetalhes!Agencia), 0, rstDetalhes!Agencia), "0000")
                .CtaCredito = Format(IIf(IsNull(rstDetalhes!Conta), 0, rstDetalhes!Conta), "0000000")
                .Filler_3 = Space(12)
                .NumVerCEL = Format(IIf(IsNull(rstDetalhes!Num_Versao_CEL), 0, rstDetalhes!Num_Versao_CEL), "0000")
                .NumLoteCEL = Format(IIf(IsNull(rstDetalhes!Num_Lote_CEL), 0, rstDetalhes!Num_Versao_CEL), "0000000")
                .Filler_4 = Space(35)
                .SequencialArquivo = FormataString(Sequencia(dSequencia), "0", Len(.SequencialArquivo), True)
                .CrLf = vbCrLf
            End With
            
            Put #iFile, , m_ArquivoTER_GCC.Detalhes
            
            '''''''''''''''''''''''''''''''''''
            'Somatoria dos valores dos cheques'
            '''''''''''''''''''''''''''''''''''
            If rstDetalhes!Valor > g_Parametros.ValorChequeLimite Then
              dTotalsuperiorValor = dTotalsuperiorValor + rstDetalhes!Valor
              dTotalsuperiorQtd = dTotalsuperiorQtd + 1
            Else
              dTotalLimiteValor = dTotalLimiteValor + rstDetalhes!Valor
              dTotalLimiteQtd = dTotalLimiteQtd + 1
            End If
            
            
            ' Gravar .T. em Geracao_TER
            If bGeracao Then
            
                Call g_cMainConnection.Execute(Proc_Atualizar.AtualizaGeracao_Ter(rstDetalhes!CMC7, iNumRemessa, m_DataTroca), _
                       lRetorno, _
                      adCmdText)
            End If
            
            
            rstDetalhes.MoveNext
            If rstDetalhes.EOF Then Exit Do
            
        Loop

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                               TRAILLER DO ARQUIVO
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    With m_ArquivoTER_GCC.Trailler
        .IdentTrailler = "TR"
        .CodigoTerceiro = Format(g_Parametros.Codigo_Terceira, "0000")
        .DataMovto = m_DataTroca
        .NumRemessa = Format(iNumRemessa, "0000")
        .TotalLimiteQtd = FormataString(dTotalLimiteQtd, "0", 6, True)  'TotalsuperiorQtd
        .TotalLimiteValor = FormataString(dTotalLimiteValor * 100, "0", 16, True) 'dTotalLimiteValor
        .TotalsuperiorQtd = FormataString(dTotalsuperiorQtd, "0", 6, True) 'TotalsuperiorQtd
        .TotalsuperiorValor = FormataString(dTotalsuperiorValor * 100, "0", 16, True) 'dTotalsuperiorValor
        .Filler_1 = FormataString(Filler_1, "0", Len(.Filler_1), True)  'Filler_1
        .Filler_2 = Space(69)
        .SequencialArquivo = FormataString(Sequencia(dSequencia), "0", Len(.SequencialArquivo), True)
    End With
    
    Put #iFile, , m_ArquivoTER_GCC.Trailler
     
    'Fecha Arquivo txt
    Close
    'Fecha Transação
    If bTransacaoAberta Then g_cMainConnection.CommitTrans

    Gera_TERGCC = True
    Screen.MousePointer = vbDefault
    Exit Function
    
Erro_Gerar:
    
    'Fecha arquivo txt
    Close
    'Cancela Transação
    If bTransacaoAberta Then g_cMainConnection.RollbackTrans
    
    Screen.MousePointer = vbDefault
    MsgBox Err.Description
    
    'Apaga arquivo gerado
    If FileExist(m_DiretorioSaida & m_ArquivoSaida) Then
        Kill m_DiretorioSaida & m_ArquivoSaida
    End If
    
End Function


Private Function Sequencia(ByRef pSequencia As Variant) As Variant
    
    pSequencia = Val(pSequencia) + 1
    
    '''''''''''''''''''
    'Retorno da função'
    '''''''''''''''''''
    Sequencia = pSequencia
    
End Function



Private Function NumRemTerGCC() As Integer
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                '* Retorna o Numero de Remessa para a Remessa Atual *'                      '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro
Dim Proc_Selecionar     As New Custodia.Selecionar
Dim Proc_Atualizar      As New Custodia.Atualizar
Dim rsNumRemessa        As New ADODB.Recordset
Dim iretorno            As Integer
Dim iNumRemessa         As Integer

'''''''''''''''''''''''''''''''''''''''
'* Retorna o Numero de Remessas Atual*'
'''''''''''''''''''''''''''''''''''''''
    
Set rsNumRemessa = m_Connection.Execute(Proc_Selecionar.GetNumRemessaTER(Geral.DataProcessamento))
    
    
If Not rsNumRemessa.EOF Then
    iNumRemessa = rsNumRemessa!Num_Remessa_TER
End If

'''''''''''''''''''''''''''''''''''''''
'    * Atualiza Valor de Remessa *'   '
'''''''''''''''''''''''''''''''''''''''
Call m_Connection.Execute(Proc_Atualizar.AtualizaNumRemessaTER_Parametro(Geral.DataProcessamento, iNumRemessa), iretorno, adCmdText)
        
'''''''''''''''''''''''''''''''''
' Retorno de Linhas  afetadas *'
'''''''''''''''''''''''''''''''''
        
NumRemTerGCC = iNumRemessa + 1
        
Exit Function
TrataErro:
    Call RemoveArquivo
    Call TratamentoErro("Erro ao atualizar Número da Remessa do Parametro.", Err)
    
End Function
