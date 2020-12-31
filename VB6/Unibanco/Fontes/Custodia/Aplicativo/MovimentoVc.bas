Attribute VB_Name = "MovimentoVc"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                '* Type de Utilização de Banco *'                           '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Type Procedures
        Inclusao        As New Custodia.Inserir    'Querys de Insert
        Alteracao       As New Custodia.Atualizar  'Querys de Update
        Deletacao       As New Custodia.Excluir    'Querys de Delete
        Selecao         As New Custodia.Selecionar 'Querys de Select
    End Type
    Private Procedures  As Procedures

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                       Header Campos Comuns Geração de Arquivos                             '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Type cg_Header
        Rotulo          As String * 6
        CNPJ            As String * 14
        NumBordero      As String * 18
        Carteira        As String * 2
    End Type
            
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                       Trailler Campos Comuns Geração de Arquivos                           '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Type cg_Trailler
        CodOFIADV       As String * 25
        CrLf            As String * 2
    End Type
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                Detalhe arquivo CHINBO - Registro de Inclusão de Borderô                    '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Type cg_CHINBO
        DtaEntrega      As String * 8
        CodLoja         As String * 10
        AgCliente       As String * 4
        CcCliente       As String * 7
        OrigBordero     As String * 1
        StatusBordero   As String * 1
        CodAgAcolhed    As String * 4
        SomaData        As String * 10
        SomaQtde        As String * 3
        SomaVlr         As String * 15
        SomaTodos       As String * 15
    End Type
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'           Detalhe arquivo CHINDT - Registro de Datas de Depósito do Borderô                '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Type cg_CHINDT
        DtaDeposito     As String * 8
        QtdCheques      As String * 3
        VlDeposito      As String * 13
    End Type

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                Detalhe arquivo CHINCH - Registro de Cheques de Borderô                     '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Type cg_CHINCH
        DtaDeposito     As String * 8
        VlCheque        As String * 13
        CMC7            As String * 34
        CodComp         As String * 3
        NumBcoEmit      As String * 4
        AgEmit          As String * 4
        CcEmit          As String * 11
        NumChEmit       As String * 10
        tpCheque        As String * 1
        TpInscricao     As String * 2
        InscrEmit       As String * 14
    End Type

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                    * Todos os Dados *                                      '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Type Dados_CHINBO
        Header          As cg_Header
        CHINBO          As cg_CHINBO
        Trailler        As cg_Trailler
    End Type
    
    Private Type Dados_CHINDT
        Header          As cg_Header
        CHINDT          As cg_CHINDT
        Trailler        As cg_Trailler
    End Type

    Private Type Dados_CHINCH
        Header          As cg_Header
        CHINCH          As cg_CHINCH
        Trailler        As cg_Trailler
    End Type
        
Private Dados_CHINBO    As Dados_CHINBO
Private Dados_CHINDT    As Dados_CHINDT
Private Dados_CHINCH    As Dados_CHINCH

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                 * Type com IdBordero Processados e em Processamento * '                    '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Type Borderos
        IdBordero       As Long
    End Type
    Private Borderos() As Borderos

'**********************************************************************************************'
'               Função de Leitura / Gravação de dados em arquivos txt                          '
'**********************************************************************************************'
Dim DirTxt          As String  '* Diretório de Arquivos txt *'
Dim LocTxt          As String  '* Local de Arquivos txt *'
Dim ArqTxt          As String  '* Nome  de Arquivos txt *'

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                              * Variáveis Auxiliares *                                      '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim d_DtaProcessamento As Long          'Data de Processamento Formatada 'MM/DD/AAAA'
Dim sDataDeposito      As Long          'Data de Deposito Atual
Dim i_IdBordero        As Long          'Identificação do Borderô Atual
Dim b_Status           As String * 1    'Status do Borderô Atual

Dim NumRemessa         As Integer       'Numero de Remessa Atual
Dim NumBorderos        As Integer       'Quantidade de Borderos
Dim iFile              As Integer

Private Function Define_Arquivo() As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                 '* Define Arquivo Texto *'                                 '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

Dim sExtensao As String

    Define_Arquivo = False

    ''''''''''''''''''''''''''''''''''''''''
    ' '* Define extensão do arquivo texto *'
    ''''''''''''''''''''''''''''''''''''''''
    If b_Status = "R" Then
        sExtensao = ".mVc" 'Movimento para Vc
    Else
        If b_Status = "C" Then
           sExtensao = ".cVc" 'Corrigidos
        Else
           sExtensao = ".rVc" 'Rejeitados
        End If
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''
    ' '* Nome do arquivo DD/MM/Remessa.extensao * '
    '''''''''''''''''''''''''''''''''''''''''''''''
    ArqTxt = Mid(Geral.DataProcessamento, 7, 2) & Mid(Geral.DataProcessamento, 5, 2) & Format(NumRemessa, String(2, "0")) & sExtensao
    
    '''''''''''''''''''''''''''''''''''''''''''
    ' '* Diretório aonde se encontra Arquivo *'
    '''''''''''''''''''''''''''''''''''''''''''
    DirTxt = Trim(g_Parametros.DiretorioTransmissao) & "\" & Geral.DataProcessamento
    
    '''''''''''''''''''''''''
    ' '* Local do Arquivo  *'
    '''''''''''''''''''''''''
    LocTxt = DirTxt & "\" & ArqTxt
    
    '''''''''''''''''''''''''
    ' '* Cria Diretório  *' '
    '''''''''''''''''''''''''
    If Dir(DirTxt, vbDirectory) = "" Then
        MkDir DirTxt
    End If
        
    Define_Arquivo = True
        
Exit Function
TrataErro:
    Call TratamentoErro("Erro durante definição de Arquivo.", Err)
    
End Function
Private Sub ShowGeracao()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                      '* Formatação da Data de Processamento *'                             '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    On Error GoTo TrataErro

    d_DtaProcessamento = Mid(Geral.DataProcessamento, 1, 4) & _
                         Mid(Geral.DataProcessamento, 5, 2) & _
                         Mid(Geral.DataProcessamento, 7, 2) _

    '''''''''''''''''''''''''''''''
    ' * Define Número de Remessa *'
    '''''''''''''''''''''''''''''''
    Call NumeroRemessa
       
    ''''''''''''''''''''''''''''''
    '  * Define Arquivo Texto *' '
    ''''''''''''''''''''''''''''''
    If Define_Arquivo Then
        '''''''''''''''''''''''''''''''
        ' *  1º Processo de Geração  *'
        '''''''''''''''''''''''''''''''
        Call BuscaDadosGeracaoCHINBO
    End If
        
Exit Sub
TrataErro:
    Call RemoveArquivo
    Call TratamentoErro("Erro durante a abertura do Módulo.", Err)
    
End Sub
Private Sub BuscaDadosGeracaoCHINBO()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       Busca Dados para a Geração do arquivo CHINBO - Registro de Inclusão de Borderô       '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim lngRegs                 As Long
    Dim Progress                As New clsProgressBar
    Dim rsGeracaoCHINBO         As New ADODB.Recordset
    Dim bTransacaoAberta        As Boolean
    Dim clSD                    As New SomatoriaDatas
    Dim Proc_Alterar            As New Custodia.Atualizar
    Dim Proc_Excluir            As New Custodia.Excluir
    Dim Proc_Selecionar         As New Custodia.Selecionar
    Dim rstDatasBordero         As New ADODB.Recordset
    Dim rstBordero              As New ADODB.Recordset
    Dim lQtdChequesIndevidos    As Long
    Dim lQtdDatasIndevidos      As Long
    Dim i                       As Integer
    Dim j                       As Integer
    Dim lRetorno                As Long
    Dim lBorderoAtual           As Long

    On Error GoTo TrataErro
    
    bTransacaoAberta = False
    
    Set rsGeracaoCHINBO = g_cMainConnection.Execute _
                          (Procedures.Selecao.GetDadosCHINBO(d_DtaProcessamento _
                                                           , b_Status))
                                                  
    'Abre transação
    g_cMainConnection.BeginTrans
    bTransacaoAberta = True
    
    If Not rsGeracaoCHINBO.EOF Then
        
        iFile = FreeFile
        Open LocTxt For Binary As iFile
        
        'Inicia progress bar
        Progress.ValorMinimo = 1
        Progress.ValorMaximo = rsGeracaoCHINBO.RecordCount
        Progress.DescricaoProcesso = "Gerando Movimento para CH ..."
        Progress.InicializaProgressBar
        lngRegs = 0
            
        Do While Not rsGeracaoCHINBO.EOF
        
            'Acumulador para ProgressBar
            lngRegs = lngRegs + 1
            
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '           * Identificação do Borderô Atual *          '
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            i_IdBordero = rsGeracaoCHINBO!IdBordero
            
            
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '             *Recalcula as datas do Borderô*           '
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            ''''''''''''''''''''''''''''''''''''''''''''''''''
            'Selecionar todas as datas de deposito do borderô'
            ''''''''''''''''''''''''''''''''''''''''''''''''''
            Set rstDatasBordero = g_cMainConnection.Execute(Proc_Selecionar.GetDatasBordero( _
                                                            d_DtaProcessamento, _
                                                            i_IdBordero))
            Do While Not rstDatasBordero.EOF()
                With rstDatasBordero
                    ''''''''''''''''''''''''''''''''''''''''''''
                    'Esta data foi corrigida pela geracao do AD'
                    ''''''''''''''''''''''''''''''''''''''''''''
                    If !QuantidadeCheques = 0 And !ValorDeposito = 0 Then
                        ''''''''''''''''''''''''''
                        'Remove cheques indevidos'
                        ''''''''''''''''''''''''''
                        Call g_cMainConnection.Execute(Proc_Excluir.RemoveChequesIndevidos( _
                                                       d_DtaProcessamento, _
                                                       i_IdBordero, _
                                                       !DataDeposito))
                        ''''''''''''''''''''''''''
                        'Remove cheques Deletados'
                        ''''''''''''''''''''''''''
                        Call g_cMainConnection.Execute(Proc_Excluir.RemoveChequesDeletados( _
                                                       d_DtaProcessamento, _
                                                       i_IdBordero, _
                                                       !DataDeposito))

                        '''''''''''''''''''''''''
                        'Remove data de deposito'
                        '''''''''''''''''''''''''
                        Call g_cMainConnection.Execute(Proc_Excluir.RemoveDataDeposito(d_DtaProcessamento, _
                                                       i_IdBordero, _
                                                       !DataDeposito), _
                                                lRetorno, _
                                                adCmdText)
                        If lRetorno = 0 Then
                            GoTo CancelaGeracao
                        End If
                        Set rsGeracaoCHINBO = g_cMainConnection.Execute _
                                              (Procedures.Selecao.GetDadosCHINBO(d_DtaProcessamento _
                                                                               , b_Status))
                    End If
                    .MoveNext
                End With
            Loop
            '''''''''''''''''
            'Corrige Borderô'
            '''''''''''''''''
            Set rstBordero = g_cMainConnection.Execute(Proc_Selecionar.GetIdBordero( _
                                                d_DtaProcessamento, _
                                                i_IdBordero))
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Calcula as datas corretas de acordo com a tabela de cheques'
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            clSD.SetConnection g_cMainConnection
            clSD.DataProcessamento = d_DtaProcessamento
            clSD.IdBordero = i_IdBordero
    
            clSD.Calcula
    
            Call g_cMainConnection.Execute(Proc_Alterar.AtualizaBordero( _
                                           d_DtaProcessamento, _
                                           i_IdBordero, _
                                           rstBordero!Num_Bordero, _
                                           rstBordero!Agencia, _
                                           rstBordero!Conta, _
                                           rstBordero!CodigoCarteira, _
                                           rstBordero!CodigoLoja, _
                                           rstBordero!DataEntrada, _
                                           rstBordero!NomeCliente, _
                                           clSD.SomatoriaDatas, _
                                           clSD.SomatoriaQuantidades, _
                                           clSD.SomatoriaValores, _
                                           clSD.SomatoriaControle), _
                                    lRetorno, _
                                    adCmdText)
            If lRetorno = 0 Then
                Set clSD = Nothing
                GoTo CancelaGeracao
            End If
            rstBordero.Close
            
            '''''''''''''''''''''''
            'Localiza o Id Bordero'
            '''''''''''''''''''''''
            'lBorderoAtual = rsGeracaoCHINBO.AbsolutePosition
            
            rsGeracaoCHINBO.Requery
            
            rsGeracaoCHINBO.Find "idBordero = " & i_IdBordero
            
            'Set rsGeracaoCHINBO = g_cMainConnection.Execute _
                                  (Procedures.Selecao.GetDadosCHINBO(d_DtaProcessamento _
                                                                   , b_Status))
            'rsGeracaoCHINBO.AbsolutePosition = lBorderoAtual
            
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '    *Fim do Processo de cálculo das datas do Borderô*   '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '   Guarda IdBordero em Array para Tratamento de Erros  '
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ReDim Preserve Borderos(0 To NumBorderos)
            Borderos(NumBorderos).IdBordero = i_IdBordero
            NumBorderos = NumBorderos + 1
            
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '        Atualiza Status para 'S' - em Transmissão      '
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If Not AtualizaStatusBorderoTransmissao("S") Then
                GoTo CancelaGeracao
'                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                ' Se o Borderô atual estiver com Status <> 'R' move para o próximo registro '
'                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                If Not rsGeracaoCHINBO.EOF Then
'                    Call BuscaDadosGeracaoCHINBO
'                End If
'                Exit Sub
            End If
    
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '  Dados de Header - Campos Comuns Geração de Arquivos  '
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            With Dados_CHINBO.Header
                .Rotulo = "CHINBO"
                .CNPJ = Format(rsGeracaoCHINBO!CNPJ_Terceira, String(14, "0"))
                .NumBordero = Format(CStr(Mid(rsGeracaoCHINBO!Num_Bordero, 1, Len(rsGeracaoCHINBO!Num_Bordero) - 1)), String(18, "0"))
                .Carteira = Format(rsGeracaoCHINBO!CodigoCarteira, String(2, "0"))
            End With
            
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '  Dados do Layout CHINBO - Registro de Inclusão de Borderô  '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            With Dados_CHINBO.CHINBO
                .DtaEntrega = rsGeracaoCHINBO!DataProcessamento
                .CodLoja = Format(rsGeracaoCHINBO!CodigoLoja, String(10, "0"))
                .AgCliente = Format(rsGeracaoCHINBO!Agencia, String(4, "0"))
                .CcCliente = Format(rsGeracaoCHINBO!Conta, String(7, "0"))
                .OrigBordero = "5"
                              
                'Verifica se status do borderô é corrigido ou não
                .StatusBordero = IIf(b_Status = "C", 3, 0)
                .CodAgAcolhed = Format(rsGeracaoCHINBO!CodigoAgAcolhed, String(4, "0"))
                .SomaData = Format(rsGeracaoCHINBO!SomaData, String(10, "0"))
                .SomaQtde = Format(rsGeracaoCHINBO!SomaQuantidade, String(3, "0"))
                .SomaVlr = Format(rsGeracaoCHINBO!SomaValor, String(15, "0"))
                .SomaTodos = Format(rsGeracaoCHINBO!SomaTodos, String(15, "0"))
            End With
            
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Dados do Layout CHINDT - Registro de Datas de Depósito do Borderô '
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            With Dados_CHINBO.Trailler
                .CodOFIADV = "CHEQUE" + Space(19)
                .CrLf = vbCrLf
            End With
            
            '''''''''''''''''''''''''''''''''''
            '  Grava Linha CHINBO no Arquivo  '
            '''''''''''''''''''''''''''''''''''
            Put #iFile, , Dados_CHINBO.Header
            Put #iFile, , Dados_CHINBO.CHINBO
            Put #iFile, , Dados_CHINBO.Trailler
           
            
            If Not BuscaDadosGeracaoCHINDT Then GoTo CancelaGeracao
            
        
            '''''''''''''''''''''''''''''''''''
            '   Move para o próximo registro  '
            '''''''''''''''''''''''''''''''''''
            rsGeracaoCHINBO.MoveNext
        
            'Atualiza Progress Bar
            Progress.AtualValue = lngRegs
            Progress.AtualizaBarra
        
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '         Atualiza Status para 'T' - Transmitido        '
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If AtualizaStatusBordero("T") = False Then
                MsgBox "Não foi posível atualizar Status do Borderô.", vbExclamation + vbOKOnly, App.Title
                GoTo CancelaGeracao
            End If
            
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '       Atualiza Número de Remessa do Bordero Atual     '
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If AtualizaNumRemessaBordero = False Then
                MsgBox "Não foi possível atualizar o Número de Remessa do Borderô.", vbExclamation + vbOKOnly, App.Title
                GoTo CancelaGeracao
            End If
            
        Loop

        'Fecha arquivo txt
        Close
        'Encerra transação
        If bTransacaoAberta Then g_cMainConnection.CommitTrans
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '               '* Geração concluida / Geração OK *'                 '
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        MsgBox "Geração concluida com sucesso.", vbInformation + vbOKOnly, App.Title
        'Encerra progress bar
        Set Progress = Nothing
        If Not (rsGeracaoCHINBO Is Nothing) Then Set rsGeracaoCHINBO = Nothing
    Else
        
        'Fecha arquivo txt
        Close
        'Cancela transação
        If bTransacaoAberta Then g_cMainConnection.RollbackTrans
        If Not (rsGeracaoCHINBO Is Nothing) Then Set rsGeracaoCHINBO = Nothing
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '                 '* Sem Movimento para Geração *'                   '
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        MsgBox "Não existem dados para transmissão.", vbExclamation + vbOKOnly, App.Title
           
    End If
    
Exit Sub

CancelaGeracao:
    'Fecha arquivo txt
    Close
    Call RemoveArquivo
    'Cancela transação
    If bTransacaoAberta Then g_cMainConnection.RollbackTrans
    If Not (rsGeracaoCHINBO Is Nothing) Then Set rsGeracaoCHINBO = Nothing
    'Encerra progress bar
    Set Progress = Nothing
    
    MsgBox "Falha durante a geração do Layout CHINBO.", vbCritical, App.Title
    Exit Sub

TrataErro:
    'Fecha arquivo txt
    Close
    Call RemoveArquivo
    Call TratamentoErro("Erro durante geração do Layout CHINBO.", Err)
    If Not (rsGeracaoCHINBO Is Nothing) Then Set rsGeracaoCHINBO = Nothing
    'Encerra progress bar
    Set Progress = Nothing
    
End Sub
Private Function BuscaDadosGeracaoCHINDT() As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  Busca Dados para a Geração do arquivo CHINDT - Registro de Datas de Depósito do Borderô   '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

    Dim rsGeracaoCHINDT As New ADODB.Recordset

    BuscaDadosGeracaoCHINDT = False
    
    Set rsGeracaoCHINDT = g_cMainConnection.Execute _
                          (Procedures.Selecao.GetDadosCHINDT(d_DtaProcessamento _
                                                           , i_IdBordero))

    If Not rsGeracaoCHINDT.EOF Then
            
        Do While Not rsGeracaoCHINDT.EOF
        
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '  Dados de Header - Campos Comuns Geração de Arquivos  '
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            With Dados_CHINBO.Header
            .Rotulo = "CHINDT"
            End With
        
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Dados do Layout CHINDT - Registro de Datas de Depósito do Borderô '
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            With Dados_CHINDT.CHINDT
                .DtaDeposito = rsGeracaoCHINDT!DataDeposito
                sDataDeposito = rsGeracaoCHINDT!DataDeposito
                .QtdCheques = Format(rsGeracaoCHINDT!QuantidadeCheques, String(3, "0"))
                .VlDeposito = Format(rsGeracaoCHINDT!ValorDeposito * 100, String(13, "0"))
            End With
            
            '''''''''''''''''''''''''''''''''''
            '  Grava Linha CHINDT no Arquivo  '
            '''''''''''''''''''''''''''''''''''
            Put #iFile, , Dados_CHINBO.Header
            Put #iFile, , Dados_CHINDT.CHINDT
            Put #iFile, , Dados_CHINBO.Trailler
            
            '''''''''''''''''''''''''''''
            ' * Gera Dados do CHINCH *  '
            '''''''''''''''''''''''''''''
            If Not BuscaDadosGeracaoCHINCH Then Exit Function
            
            rsGeracaoCHINDT.MoveNext
        Loop
        
    End If

    BuscaDadosGeracaoCHINDT = True
    
    Exit Function

TrataErro:
    Call RemoveArquivo
    Call TratamentoErro("Erro durante geração do Layout CHINDT.", Err)

End Function
Private Function BuscaDadosGeracaoCHINCH() As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       Busca Dados para a Geração do arquivo CHINCH - Registro de Cheques de Borderô        '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

    Dim rsGeracaoCHINCH As New ADODB.Recordset
    Dim pCMC7           As String

    BuscaDadosGeracaoCHINCH = False
    
    Set rsGeracaoCHINCH = g_cMainConnection.Execute _
                          (Procedures.Selecao.GetDadosCHINCH(d_DtaProcessamento _
                                                           , i_IdBordero _
                                                           , sDataDeposito))
                                                           
    If Not rsGeracaoCHINCH.EOF Then
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '  Dados de Header - Campos Comuns Geração de Arquivos  '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        With Dados_CHINBO.Header
            .Rotulo = "CHINCH"
        End With
        
        Do While Not rsGeracaoCHINCH.EOF
        
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Dados do Layout CHINCH - Registro de Cheques de Borderô           '
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
           
            With Dados_CHINCH.CHINCH
                .DtaDeposito = rsGeracaoCHINCH!DataDeposito
                .VlCheque = Format(rsGeracaoCHINCH!Valor * 100, String(13, "0"))
                .CMC7 = Space(1) & Mid(rsGeracaoCHINCH!CMC7, 1, 8) & Space(1) & Mid(rsGeracaoCHINCH!CMC7, 9, 10) & Space(1) & Mid(rsGeracaoCHINCH!CMC7, 19, 12) & Space(1)
                .CodComp = Format(Mid(rsGeracaoCHINCH!CMC7, 9, 3), String(3, "0"))
                .NumBcoEmit = Format(Mid(rsGeracaoCHINCH!CMC7, 1, 3), String(4, "0"))
                .AgEmit = Format(Mid(rsGeracaoCHINCH!CMC7, 4, 4), String(4, "0"))
                .CcEmit = Format(Mid(rsGeracaoCHINCH!CMC7, 20, 10), String(11, "0"))
                .NumChEmit = Format(Mid(rsGeracaoCHINCH!CMC7, 12, 6), String(10, "0"))
                .tpCheque = Format(Mid(rsGeracaoCHINCH!CMC7, 18, 1), String(1, "0"))
                .TpInscricao = Format(rsGeracaoCHINCH!TipoInscricao, String(2, "0"))
                .InscrEmit = Format(rsGeracaoCHINCH!CNPJCPF, String(14, "0"))
            End With
            
            '''''''''''''''''''''''''''''''''''
            '  Grava Linha CHINCH no Arquivo  '
            '''''''''''''''''''''''''''''''''''
            Put #iFile, , Dados_CHINBO.Header
            Put #iFile, , Dados_CHINCH.CHINCH
            Put #iFile, , Dados_CHINBO.Trailler
        
        
        
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '  888       Atualiza Status para 'T' - Transmitido        '
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            pCMC7 = rsGeracaoCHINCH!CMC7
            If AtualizaStatusCheque(i_IdBordero, pCMC7, "T") = False Then
                MsgBox "Não foi posível atualizar Status do Cheque.", vbExclamation + vbOKOnly, App.Title
                Exit Function
            End If
            
        
            '''''''''''''''''''''''''''''''''''
            '   Move para o próximo registro  '
            '''''''''''''''''''''''''''''''''''
            rsGeracaoCHINCH.MoveNext
        
        Loop
        
    End If
    
    BuscaDadosGeracaoCHINCH = True
    
    Exit Function

TrataErro:
    Call RemoveArquivo
    Call TratamentoErro("Erro durante geração do Layout CHINCH.", Err)
    
End Function
Public Sub SetStatus(pStatus As String)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                 '* Receberá o Status para a Geração dos Arquivos *'                        '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

    b_Status = pStatus

    ''''''''''''''''''''''''''''
    '*   Inicializa Geração   *'
    ''''''''''''''''''''''''''''
    Call ShowGeracao

Exit Sub
TrataErro:
    Call RemoveArquivo
    Call TratamentoErro("Erro na inicialização.", Err)
    
End Sub
Public Sub RemoveArquivo()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'     '* Caso ocorra qualquer erro durante a Geração do Arquivo o mesmo será Deletado *'     '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro
    
    '''''''''''''''''''
    '* Fecha Arquivo *'
    '''''''''''''''''''
    Close #iFile
    
    ''''''''''''''''''''
    '* Exclui Arquivo *'
    ''''''''''''''''''''
    If FileExist(LocTxt) Then
        Kill LocTxt
    End If

'    If VoltaStatusBordero = False Then
'        MsgBox "Não foi possível atualizar Status do Borderô.", vbExclamation + vbOKOnly, App.Title
'    End If

Exit Sub
TrataErro:
    Call TratamentoErro("Erro ao remover arquivo.", Err)
    
End Sub
Private Function AtualizaStatusBorderoTransmissao(pStatus As String) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'               '* Atualiza Status do Borderô Atual conforme Tratamento *'                   '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

Dim iretorno As Integer
    
    Call g_cMainConnection.Execute(Procedures.Alteracao.AtualizaStatusBorderoTransmissao _
                                  (d_DtaProcessamento, i_IdBordero, pStatus) _
                                  , iretorno, adCmdText)
                                   
    '''''''''''''''''''''''''''''''''
    '* Retorno de Linhas  afetadas *'
    '''''''''''''''''''''''''''''''''
    If iretorno = 0 Then
        AtualizaStatusBorderoTransmissao = False
        MsgBox "Não foi possível atualizar status do arquivo de borderô em transmissão", vbCritical, App.Title
    Else
        AtualizaStatusBorderoTransmissao = True
    End If
    
Exit Function
TrataErro:
    Call RemoveArquivo
    Call TratamentoErro("Erro ao atualizar Status do Borderô em Transmissão.", Err)
    
End Function
Public Function AtualizaStatusBordero(pStatus As String) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'               '* Atualiza Status do Borderô Atual conforme Tratamento *'                   '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

Dim iretorno As Integer
    
    Call g_cMainConnection.Execute(Procedures.Alteracao.AtualizaStatusBordero _
                                  (d_DtaProcessamento, i_IdBordero, pStatus) _
                                  , iretorno, adCmdText)
                                   
    '''''''''''''''''''''''''''''''''
    '* Retorno de Linhas  afetadas *'
    '''''''''''''''''''''''''''''''''
    If iretorno = 0 Then
        AtualizaStatusBordero = False
    Else
        AtualizaStatusBordero = True
    End If

Exit Function
TrataErro:
    Call RemoveArquivo
    Call TratamentoErro("Erro ao atualizar Status do Borderô.", Err)
    
End Function
Private Function NumeroRemessa() As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                '* Retorna o Numero de Remessa para a Remessa Atual *'                      '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

Dim rsNumRemessa As New ADODB.Recordset
Dim iretorno     As Integer

    '''''''''''''''''''''''''''''''''''''''
    '* Retorna o Numero de Remessas Atual*'
    '''''''''''''''''''''''''''''''''''''''
    Set rsNumRemessa = g_cMainConnection.Execute(Procedures.Selecao.GetNumRemessaMov(d_DtaProcessamento))
                                                
    If Not rsNumRemessa.EOF Then
        NumRemessa = IIf(IsNull(rsNumRemessa!Num_Remessa_MOV), 0, rsNumRemessa!Num_Remessa_MOV)
    End If

    '''''''''''''''''''''''''''''''''''''''
    '    * Atualiza Valor de Remessa *'   '
    '''''''''''''''''''''''''''''''''''''''
    Call g_cMainConnection.Execute(Procedures.Alteracao.AtualizaNumRemessaMOV_Parametro(d_DtaProcessamento, NumRemessa), iretorno, adCmdText)
        
    '''''''''''''''''''''''''''''''''
    '* Retorno de Linhas  afetadas *'
    '''''''''''''''''''''''''''''''''
    If iretorno = 0 Then
        NumeroRemessa = False
    Else
        NumeroRemessa = True
        NumRemessa = NumRemessa + 1
    End If
        
Exit Function
TrataErro:
    Call RemoveArquivo
    Call TratamentoErro("Erro ao atualizar Número de Remessa do Parametro.", Err)
    
End Function
Private Function AtualizaNumRemessaBordero() As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                      '* Atualiza Número de Remessas do Borderô Atual *'                    '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

Dim iretorno As Integer

    Call g_cMainConnection.Execute(Procedures.Alteracao.AtualizaNumRemessaMOV_Bordero _
                                  (d_DtaProcessamento _
                                 , i_IdBordero _
                                 , NumRemessa), iretorno, adCmdText)
                                                                 
    '''''''''''''''''''''''''''''''''
    '* Retorno de Linhas  afetadas *'
    '''''''''''''''''''''''''''''''''
    If iretorno = 0 Then
        AtualizaNumRemessaBordero = False
    Else
        AtualizaNumRemessaBordero = True
    End If
    
    Exit Function

TrataErro:
    Call RemoveArquivo
    Call TratamentoErro("Erro ao atualizar Número de Remessa do Borderô.", Err)
    
End Function
Private Function VoltaStatusBordero() As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' * Caso ocorra qualquer erro durante  a Geração do Arquivo volta Status de todos Borderos * '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

Dim contador    As Integer
Dim iretorno    As Integer
    
    For contador = 0 To NumBorderos - 1

        Call g_cMainConnection.Execute(Procedures.Alteracao.AtualizaStatusBordero( _
                                       d_DtaProcessamento _
                                     , Borderos(contador).IdBordero _
                                     , b_Status), iretorno, adCmdText)
                   
        '''''''''''''''''''''''''''''''''
        '* Retorno de Linhas  afetadas *'
        '''''''''''''''''''''''''''''''''
        If iretorno = 0 Then
            VoltaStatusBordero = False
        Else
            VoltaStatusBordero = True
        End If
        
    Next

Exit Function
TrataErro:
    Call RemoveArquivo
    Call TratamentoErro("Erro ao voltar Status do Borderô.", Err)
    
End Function



Public Function AtualizaStatusCheque(i_IdBordero, pCMC7, pStatus As String) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'               '* Atualiza Status do Cheque Atual para trabsmitido                          '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

Dim iretorno As Integer

    
    Call g_cMainConnection.Execute(Procedures.Alteracao.AlteraStatusCheque( _
                                  i_IdBordero, pCMC7, pStatus), _
                                  iretorno, adCmdText)
                                   
    '''''''''''''''''''''''''''''''''
    '* Retorno de Linhas  afetadas *'
    '''''''''''''''''''''''''''''''''
    If iretorno = 0 Then
        AtualizaStatusCheque = False
    Else
        AtualizaStatusCheque = True
    End If

Exit Function
TrataErro:
    Call RemoveArquivo
    Call TratamentoErro("Erro ao atualizar Status do Cheque.", Err)
    
End Function

