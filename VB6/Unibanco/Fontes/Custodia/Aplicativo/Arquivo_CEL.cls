VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Arquivo_CEL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum enumTipoExportacaoCEL605
    Limite
    Superior
    Todos
End Enum
''''''''''''''''
'Types privados'
''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                           HEADER DO ARQUIVO CEL - 605
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Type tpHeaderArquivo_CEL605
    ControleHeader              As String * 47      '0
    NomeArquivo                 As String * 6       'CEL605
    Compe                       As String * 3       'Tabela parametro Comp_Origem_CEL
    VersaoArquivo               As String * 4       'Número da remessa
    BancoRemetente              As String * 3       '409
    DVBancoRemetente            As String * 1       'X
    IndicadorRemessa            As String * 1       '
    DataMovimento               As String * 8       'Pedir ao usuario
    Filler                      As String * 77      'Brancos
    SequencialArquivo           As String * 10      '1
    CrLf                        As String * 2       'OK
End Type
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                          DETALHES DO ARQUIVO CEL - 605
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Type tpDetalheArquivo_CEL605
    Comp                        As String * 3       '01 CMC7
    Banco                       As String * 3       '02 CMC7
    Agencia                     As String * 4       '03 CMC7
    DV2                         As String * 1       '04 CMC7
    Conta                       As String * 12      '05 CMC7
    DV1                         As String * 1       '06 CMC7
    NumeroCheque                As String * 6       '07 CMC7
    DV3                         As String * 1       '08 CMC7
    Filler_1                    As String * 2       '09 Espaco em branco
    Valor                       As String * 17      '10 o que esta na tabela (Valor)
    Tipificacao                 As String * 1       '11 CMC7
    TipoDocumento               As String * 2       '12 10-Superior ou 14-Limite
    Filler_2                    As String * 2       '13 Zeros
    BancoRemetente              As String * 3       '14 409
    AgenciaAcolhedora           As String * 4       '15 0 - ?
    AgenciaDepositante          As String * 4       '16 0 - ?
    ContaDepositante            As String * 12      '17 0 - ?
    CompAcolhido                As String * 3       '18 Pegar da tabela parametro (Comp_Origem_CEL)
    DataApresentacao            As String * 8       '19 Pedir ao usuario
    LoteCheque                  As String * 7       '20 Iniciar com o que esta no parametro (Num_Lote_CEL) e depois sequencia e depois atualiza a tabela parametro
    SequenciaRegistro           As String * 3       '21 sequencial do lote
    CentroProcessador           As String * 6       '22 0 - ?
    Filler_3                    As String * 45      '23
    SequencialArquivo           As String * 10      '24 sequencia de linha
    CrLf                        As String * 2
End Type
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                          FECHAMENTO DO LOTE CEL - 605
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Type tpFechamentoArquivo_CEL605
    Comp_Destinataria           As String * 3   'Parametro.Comp_Origem_CEL
    Banco_Destinatario          As String * 3   'Banco do Cheque
    Controle_Fechamento         As String * 27  'colocar 9
    Somatorio_Detalhes          As String * 17  'Somatoria dos cheques
    Filler_1                    As String * 1   '
    TipoDocumento               As String * 2   '10 superior ou 14 limite
    Filler_2                    As String * 2   '
    Banco_Remetente             As String * 3   '409
    Versao_Arquivo              As String * 7   'Igual ao Header
    Unidade_Federativa          As String * 2   'Colocar UF na tabela parametro
    Filler_3                    As String * 14  '
    Data_Movimento              As String * 8   'Pedir ao usuario
    Numero_Lote                 As String * 7   'ultimo n do lote gerado em detalhes
    Sequencial_Lote             As String * 3   '999
    Centro_Processador          As String * 6   ' CPD Origem
    Filler_4                    As String * 45  '
    SequencialArquivo           As String * 10  'Sequencia desde o inicio
    CrLf                        As String * 2
End Type
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                         TRAILLER DO ARQUIVO CEL - 605
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Type tpTraillerArquivo_CEL605
    Controle_Trailler           As String * 47  'Igual a 9
    Nome_Arquivo                As String * 6   'CEL605
    Codigo_Comp_Origem          As String * 3   'Parametro.Comp_Origem_CEL
    Versao_Arquivo              As String * 4   'Igual ao Header
    Banco_Remetente             As String * 3   '409
    DV_Banco                    As String * 1   'X
    Indicador_Remessa           As String * 1   '1 Limite ou 2 Superior
    Data_Processamento          As String * 8   'Pedir ao usuario
    Valor_Arquivo               As String * 17  'Somatoria total ( Cheques )
    Filler                      As String * 60
    SequencialArquivo           As String * 10  'Sequencia desde o inicio
End Type
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                               ARQUIVO CEL - 605
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Type tpArquivo_CEL605
    Header                      As tpHeaderArquivo_CEL605
    Detalhes                    As tpDetalheArquivo_CEL605
    Fechamento                  As tpFechamentoArquivo_CEL605
    Trailler                    As tpTraillerArquivo_CEL605
End Type


''''''''''''''''''''''
'Tipo arquivo CEL 605'
''''''''''''''''''''''
Private m_ArquivoCEL            As tpArquivo_CEL605


Private m_TipoArquivoCEL        As enumTipoArquivoCEL
Private m_DataProcessamento     As Long
Private m_ProgressBar           As Object
Private m_DiretorioSaida        As String
Private m_ArquivoSaida          As String
Private m_Connection            As New ADODB.Connection
Private m_DataTroca             As Long
Private m_TipoCheque            As Integer
Private m_TipoExportacao        As enumTipoExportacaoCEL605
Private m_Inidk                 As String
Private m_Inik7                 As String
Private m_hdx                   As String



Public Property Let ArquivoSaida(ByVal pArquivoSaida As String)
    m_ArquivoSaida = pArquivoSaida
End Property

Public Property Let TipoCheque(ByVal pTipoCheque As Integer)
    m_TipoCheque = pTipoCheque
End Property

Public Property Get TipoCheque() As Integer
    TipoCheque = m_TipoCheque
End Property

Public Property Let DataTroca(ByVal pDataTroca As Long)
    m_DataTroca = pDataTroca
End Property

Public Property Get DataTroca() As Long
    DataTroca = m_DataTroca
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

Public Function Exportar() As Boolean

    Dim sStr                As String
    Dim rstDetalhes         As New ADODB.Recordset
    'Dim rstBordero          As New ADODB.Recordset
    Dim Proc_Selecionar     As New Custodia.Selecionar
    Dim Proc_atualizar      As New Custodia.Atualizar
    'Dim cCheque             As New clsCheque
    Dim cCheque             As New CalculoCheque
    Dim iFile               As Integer
    Dim sBanco              As String
    Dim dSequencia          As Double
    Dim dValorLote          As Double
    Dim dValorArquivo       As Double
    Dim lRetorno            As Long
    Dim iVersaoArquivo      As Integer
    
    On Error GoTo Erro_Exportar
    
    Exportar = False
    Screen.MousePointer = vbHourglass
    
    Set rstDetalhes = m_Connection.Execute(Proc_Selecionar.GetDetalhesArquivoCEL_605( _
                                            m_DataTroca, _
                                            m_TipoCheque, _
                                            m_TipoExportacao))
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                            CRIAÇÃO E ABERTURA DO ARQUIVO CEL
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    If FileExist(m_DiretorioSaida & m_ArquivoSaida) Then
        Kill m_DiretorioSaida & m_ArquivoSaida
    End If

    iFile = FreeFile
    Open m_DiretorioSaida & m_ArquivoSaida For Binary As iFile


    iVersaoArquivo = NumRemCEL()
    
    If g_Parametros.HeaderAV Then
      
      m_Inidk = F_Inidk(iVersaoArquivo)
      m_Inik7 = F_Inik7(iVersaoArquivo)
      m_hdx = F_HDX(iVersaoArquivo)
    
      Put #iFile, , m_Inidk
      Put #iFile, , m_Inik7
      Put #iFile, , m_hdx
      
    End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                     HEADER DO ARQUIVO
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    With m_ArquivoCEL.Header
        .ControleHeader = FormataString("", "0", Len(.ControleHeader), True)
        .NomeArquivo = "CEL605"
        .Compe = FormataString(g_Parametros.Comp_Origem_CEL, "0", Len(.Compe), True)
        .VersaoArquivo = FormataString(iVersaoArquivo, "0", Len(.VersaoArquivo), True) 'Parametro.VersaoArquivo_CEL
        .BancoRemetente = "409"
        .DVBancoRemetente = "X"
        .IndicadorRemessa = IIf((m_TipoArquivoCEL = eCheque_Superior) Or (m_TipoArquivoCEL = eCheque_Unibanco), "2", "1")
        .DataMovimento = m_DataTroca
        .Filler = FormataString("", " ", Len(.Filler), True)
        .SequencialArquivo = FormataString(Sequencia(dSequencia), "0", Len(.SequencialArquivo), True)
        .CrLf = vbCrLf
    End With
    
    Put #iFile, , m_ArquivoCEL.Header


    '''''''''''''''''''''''''''''''''''''''''''
    'Leitura do Recordset e escrita no arquivo'
    '''''''''''''''''''''''''''''''''''''''''''
    
    dValorArquivo = 0

    Do While Not rstDetalhes.EOF
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                  DETALHES DO ARQUIVO
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        sBanco = Left(rstDetalhes!CMC7, 3)
        dValorLote = 0
        
        Do While (sBanco = Left(rstDetalhes!CMC7, 3)) And _
                 (Val(m_ArquivoCEL.Detalhes.SequenciaRegistro) < 200)
                 
            '''''''''''''''''''''''''''''''''''''''
            'Incrementa o ProgressBar do Principal'
            '''''''''''''''''''''''''''''''''''''''
            m_ProgressBar.Value = rstDetalhes.AbsolutePosition / rstDetalhes.RecordCount * 10
                           
            cCheque.CMC7 = rstDetalhes!CMC7
            cCheque.Calcula
            
            With m_ArquivoCEL.Detalhes
                .Comp = FormataString(cCheque.Comp, "0", Len(.Comp), True)
                .Banco = FormataString(cCheque.Banco, "0", Len(.Banco), True)
                .Agencia = FormataString(cCheque.Agencia, "0", Len(.Agencia), True)
                .DV2 = cCheque.DV2
                .Conta = FormataString(cCheque.Conta, "0", Len(.Conta), True)
                .DV1 = cCheque.DV1
                .NumeroCheque = FormataString(cCheque.NumeroCheque, "0", Len(.NumeroCheque), True)
                .DV3 = cCheque.DV3
                .Filler_1 = FormataString("", " ", Len(.Filler_1), True)
                .Valor = FormataString(rstDetalhes!Valor * 100, "0", Len(.Valor), True)
                .Tipificacao = cCheque.Tipificacao
                .TipoDocumento = IIf((m_TipoArquivoCEL = eCheque_Superior), "10", "14")
                .Filler_2 = FormataString("", "0", Len(.Filler_2), True)
                .BancoRemetente = "409"
                .AgenciaAcolhedora = FormataString(g_Parametros.CodigoAgAcolhed, "0", Len(.AgenciaAcolhedora), True)  'Parametro.CodigoAgAcolhed
                .AgenciaDepositante = FormataString("0", "0", Len(.AgenciaDepositante), True) 'Bordero.Agencia
                .ContaDepositante = FormataString("0", "0", Len(.ContaDepositante), True) 'Bordero.Conta
                .CompAcolhido = FormataString(g_Parametros.Comp_Origem_CEL, "0", Len(.CompAcolhido), True)   '(Comp_Origem_CEL)
                .DataApresentacao = m_DataTroca
                .LoteCheque = FormataString(g_Parametros.Numero_Lote_CEL, "0", Len(.LoteCheque), True)
                .SequenciaRegistro = FormataString(Sequencia(.SequenciaRegistro), "0", Len(.SequenciaRegistro), True)
                .CentroProcessador = FormataString(g_Parametros.Codigo_USB, "0", Len(.CentroProcessador), True)  'Parametro.Codigo_USB
                .Filler_3 = FormataString("", " ", Len(.Filler_3), True)
                .SequencialArquivo = FormataString(Sequencia(dSequencia), "0", Len(.SequencialArquivo), True)
                .CrLf = vbCrLf
            End With
            
            Put #iFile, , m_ArquivoCEL.Detalhes
            
            '''''''''''''''''''''''''''''''''''
            'Somatoria dos valores dos cheques'
            '''''''''''''''''''''''''''''''''''
            dValorLote = dValorLote + rstDetalhes!Valor * 100
            
            '''''''''''''''''''''''''''''''''
            'Somatoria dos valores dos lotes'
            '''''''''''''''''''''''''''''''''
            dValorArquivo = dValorArquivo + rstDetalhes!Valor * 100
            
            
            ' Gravar .T. em Geracao_CEL
            Call g_cMainConnection.Execute(Proc_atualizar.AtualizaGeracao_CEL(rstDetalhes!CMC7, iVersaoArquivo, m_DataTroca), _
                            lRetorno, _
                            adCmdText)
            
            
            rstDetalhes.MoveNext
            If rstDetalhes.EOF Then Exit Do
            
        Loop
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                 FECHAMENTO DO ARQUIVO
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        With m_ArquivoCEL.Fechamento
            .Comp_Destinataria = FormataString(g_Parametros.Comp_Origem_CEL, "0", Len(.Comp_Destinataria), True)
            .Banco_Destinatario = cCheque.Banco
            .Controle_Fechamento = FormataString("", "9", Len(.Controle_Fechamento), True)
            .Somatorio_Detalhes = FormataString(dValorLote, "0", Len(.Somatorio_Detalhes), True)
            .Filler_1 = FormataString("", " ", Len(.Filler_1), True)
            .TipoDocumento = FormataString(IIf((m_TipoArquivoCEL = eCheque_Superior), "10", "14"), "0", Len(.TipoDocumento), True)
            .Filler_2 = FormataString("", " ", Len(.Filler_2), True)
            .Banco_Remetente = "409"
            .Versao_Arquivo = FormataString(.Versao_Arquivo, "?", Len(.Versao_Arquivo), True) 'Parametro.VersaoArquivo_CEL
            .Unidade_Federativa = g_Parametros.UF_Terceira 'Parametro UF
            .Filler_3 = FormataString("", " ", Len(.Filler_3), True)
            .Data_Movimento = m_DataTroca
            .Numero_Lote = FormataString(g_Parametros.Numero_Lote_CEL, "0", Len(.Numero_Lote), True) 'ultimo n do lote gerado em detalhes
            .Sequencial_Lote = FormataString("", "9", Len(.Sequencial_Lote), True)
            .Centro_Processador = FormataString(g_Parametros.Codigo_USB, "0", Len(.Centro_Processador), True)  'Parametro.Codigo_USB
            .Filler_4 = FormataString(.Filler_4, " ", Len(.Filler_4), True)
            .SequencialArquivo = FormataString(Sequencia(dSequencia), "0", Len(.SequencialArquivo), True)
            .CrLf = vbCrLf
        End With
        
        g_Parametros.Numero_Lote_CEL = Sequencia(g_Parametros.Numero_Lote_CEL)

        
        Put #iFile, , m_ArquivoCEL.Fechamento
        '''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Reseta a sequencia do registro do bloco de Detalhes'
        '''''''''''''''''''''''''''''''''''''''''''''''''''''
        m_ArquivoCEL.Detalhes.SequenciaRegistro = 0
        
    Loop
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                  TRAILLER DO ARQUIVO
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    With m_ArquivoCEL.Trailler
        .Controle_Trailler = FormataString("", "9", Len(.Controle_Trailler), True)
        .Nome_Arquivo = "CEL605"
        .Codigo_Comp_Origem = FormataString(g_Parametros.Comp_Origem_CEL, "0", Len(.Codigo_Comp_Origem), True)
        .Versao_Arquivo = FormataString(iVersaoArquivo, "0", Len(.Versao_Arquivo), True) 'Parametro.VersaoArquivo_CEL
        .Banco_Remetente = "409"
        .DV_Banco = "X"
        .Indicador_Remessa = IIf(m_TipoArquivoCEL = eCheque_Limite, "1", 2)
        .Data_Processamento = m_DataProcessamento
        .Valor_Arquivo = FormataString(dValorArquivo * 100, "0", Len(.Valor_Arquivo), True)
        .SequencialArquivo = FormataString(Sequencia(dSequencia), "0", Len(.SequencialArquivo), True)
        .Filler = FormataString("", Space(1), Len(.Filler), True)
    End With
    
    Put #iFile, , m_ArquivoCEL.Trailler
    
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                FECHAMENTO DO ARQUIVO CEL
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Close iFile

    Exportar = True
    Screen.MousePointer = vbDefault
    Exit Function
    
Erro_Exportar:
    Close
    Screen.MousePointer = vbDefault
    
End Function

Private Function Sequencia(ByRef pSequencia As Variant) As Variant
    
    pSequencia = Val(pSequencia) + 1
    
    '''''''''''''''''''
    'Retorno da função'
    '''''''''''''''''''
    Sequencia = pSequencia
    
End Function


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



Private Function NumRemCEL() As Integer
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                '* Retorna o Numero de Remessa para a Remessa Atual *'                      '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro
Dim Proc_Selecionar     As New Custodia.Selecionar
Dim Proc_atualizar      As New Custodia.Atualizar
Dim rsNumVersaoCEL      As New ADODB.Recordset
Dim iretorno            As Integer
Dim iNumVersaoCEL       As Integer

'''''''''''''''''''''''''''''''''''''''
'* Retorna o Numero de Remessas Atual*'
'''''''''''''''''''''''''''''''''''''''
    
Set rsNumVersaoCEL = m_Connection.Execute(Proc_Selecionar.GetNumVersaoCEL(Geral.DataProcessamento))
    
    
If Not rsNumVersaoCEL.EOF Then
    iNumVersaoCEL = rsNumVersaoCEL!Num_Versao_CEL
End If

'''''''''''''''''''''''''''''''''''''''
'    * Atualiza Valor de Versao *'   '
'''''''''''''''''''''''''''''''''''''''
Call m_Connection.Execute(Proc_atualizar.AtualizaNumVersaoCEL_Parametro(Geral.DataProcessamento, iNumVersaoCEL), iretorno, adCmdText)
        
       
NumRemCEL = iNumVersaoCEL + 1
    
        
Exit Function
TrataErro:
    Call RemoveArquivo
    Call TratamentoErro("Erro ao atualizar Número da Versão CEL do Parametro.", Err)
    
End Function
