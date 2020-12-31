Attribute VB_Name = "Regra"
Option Explicit
' ****************************************************************
' * Definição de Variáveis Utilizadas nas Regras de Encerramento *
' * de Malote Empresa                                            *
' ****************************************************************
Enum enRetornoMalote
    enMaloteDigitar = 0             ' Malote Enviado para Digitação
    enMaloteSupervisor = 1          ' Malote Enviado para o Supervisor
    enMaloteProvaZero = 2           ' Malote Enviado para a Prova Zero
    enMaloteAlcada = 3              ' Malote Enviado para a Alçada de Valores
    enMaloteVinculo = 4             ' Malote Enviado para o Vínculo Manual
    enMaloteRobo = 5                ' Malote Enviado para o Robo
    enMaloteErro = 9                ' Malote Encerrado com Erro
End Enum

Type tpRegraMalote
    sSql As String                  ' String do SQL
    rdoTB As rdoResultset           ' Cursor de Leitura
    qryInserirAjuste As rdoQuery    ' Chamada de Stored Procedure
    
    bDeposito As Boolean            ' Indica se Existe Deposito no Malote
    nIdLote As Double               ' Numero de Identificação do Lote no BD
    nIdMalote As Double             ' Numero de Identificação do Malote no BD
    nAgenciaMalote As Double        ' Numero da Agência do Malote
    nContaMalote As Double          ' Numero da Cotna do Malote
    nQtdeDocumentos As Integer      ' Qtde de Documentos no Malote
    nId As Integer                  ' Variável Auxiliar de Acesso ao Array de Documentos
    sStatus As String               ' Status do Malote Empresa
    sPendenciaValor As String       ' Indicador de Pendência de Valor
    sAlcada As String               ' Indicador de Pendência de Alçada
    sSupervisor As String           ' Indicador de Pendência de Supervisão
    sVinculoManual As String        ' Indicador de Pendência de Vinculo
    
    nValorDebito As Currency        ' Valor de Débitos  do Malote Empresa
    nValorCredito As Currency       ' Valor de Créditos do Malote Empresa
    nValorContas As Currency        ' Valor de Contas / Depósitos
    nValorCheques As Currency       ' Valor de Cheques / ADCC
    nValorAjusteAuto As Currency    ' Valor de Parametro para Ajuste Automático de Diferenças
    
    nIdDocto() As Double            ' Array com o Numero de Identificação do Documento no BD
    nValorTotal() As Currency       ' Array com o Valor do Documento
    nTipoDocto() As Integer         ' Array com o Tipo do Documento
    nVinculo() As Integer           ' Array com o Número do Vinculo de Documentos no Malote
    sAlcadaDocumento() As String    ' Array contendo a Alçada dos Documentos
    sQualDocumento() As String * 2  ' Array com o Tipo de Documento na Transação
    bDesprezarVinculo() As Boolean  ' Array Marcando para desprezar o Vínculo deste Documento
    bVinculou() As Boolean          ' Array Marcando a Alteração do Vínculo
End Type

Public RegraMalote As tpRegraMalote
' **********************************
' * Executa Encerramento do Malote *
' **********************************
Public Function EncerrarMalote(ByVal pvnDataProcessamento As Double, _
                               ByVal pvnIdMalote As Double, _
                               ByVal pvbIgnorarProvaZero As Boolean) As Integer
    Dim nInd As Integer             ' Auxiliar de Acesso ao Array
    Dim sDocumentos As String       ' Documentos a Alterar
    Dim nDiferenca As Currency      ' Balanceamento do Vínculo
    Dim nQtdeContas As Integer      ' Qtde de Contas sem Vínculo
    Dim nQtdeCheques As Integer     ' Qtde de Cheques sem Vínculo
    
    On Error GoTo Erro
    
    ' *****************************************
    ' * Iniciando Transação do Banco de Dados *
    ' *****************************************
    Geral.Banco.BeginTrans
    
    With RegraMalote
        ' ************************************
        ' * Inicializando Situação do Malote *
        ' ************************************
        .bDeposito = False
        .sStatus = ""
        .sPendenciaValor = "N"
        .sAlcada = "N"
        .sSupervisor = "N"
        .sVinculoManual = "N"
        .nValorCheques = 0
        .nValorContas = 0
        .nValorCredito = 0
        .nValorDebito = 0
        
        ' ****************************************************
        ' * Selecionando Dados Necessários do Malote Empresa *
        ' ****************************************************
        SelecionandoMalote pvnDataProcessamento, pvnIdMalote
        
        ' *****************************************************
        ' * Verificar se existem Documentos a serem Digitados *
        ' *****************************************************
        .sSql = ""
        .sSql = .sSql & "Select Count(*) "
        .sSql = .sSql & "From Documento ( NOLOCK Index=indDocumento01 ) "
        .sSql = .sSql & "Where DataProcessamento = " & pvnDataProcessamento & " And "
        .sSql = .sSql & "      IdLote            = " & .nIdLote & " And "
        .sSql = .sSql & "      IdEnvelope        = " & .nIdMalote & " And "
        .sSql = .sSql & "      IdDocto           > 0 And "
        .sSql = .sSql & "      Status            = '0'"
        
        Set .rdoTB = Geral.Banco.OpenResultset(.sSql, rdOpenKeyset, rdConcurRowVer)
        
        .nQtdeDocumentos = IIf(IsNull(.rdoTB(0)), 0, .rdoTB(0))
        
        .rdoTB.Close
        
        If .nQtdeDocumentos > 0 Then
            ' ****************************************************
            ' * Malote Empresa ainda não foi Totalmente Digitado *
            ' * Retornando o mesmo para o Status de Digitação    *
            ' ****************************************************
            .sStatus = "1"
            AlterandoSituacaoMalote pvnDataProcessamento
            GoTo FimOK
        End If
        
        If pvbIgnorarProvaZero Then
            ' ************************************************
            ' * Carregando na Tabela Parametro               *
            ' * o Valor de Ajuste Automático para Diferenças *
            ' ************************************************
            .sSql = ""
            .sSql = .sSql & "Select ValorAjusteAuto "
            .sSql = .sSql & "From Parametro (NOLOCK) "
            .sSql = .sSql & "Where DataProcessamento = " & pvnDataProcessamento
            
            Set .rdoTB = Geral.Banco.OpenResultset(.sSql, rdOpenKeyset, rdConcurRowVer)
                
            RegraMalote.nValorAjusteAuto = .rdoTB!ValorAjusteAuto
                
            .rdoTB.Close
        Else
            RegraMalote.nValorAjusteAuto = 0
        End If
    End With
    
    ' ************************************************
    ' * Selecionando os Documentos do Malote Empresa *
    ' ************************************************
    SelecionandoDocumentos pvnDataProcessamento
    
    If RegraMalote.nQtdeDocumentos = 0 Then
        GoTo EnviarRobo
    End If
    
    ' ******************************************
    ' * Verificando Situação do Malote Empresa *
    ' ******************************************
    With RegraMalote
        For .nId = 1 To .nQtdeDocumentos
            If .nTipoDocto(.nId) = 0 Or _
                .nValorTotal(.nId) = 0 Then
                ' *************************************
                ' * Documento não foi digitado        *
                ' * Enviando Malote para o Supervisor *
                ' *************************************
                .sSupervisor = "S"
                Exit For
            ElseIf .sQualDocumento(.nId) = "DE" Then
                ' **************************
                ' * Malote possui Depósito *
                ' **************************
                .bDeposito = True
                If .nId = .nQtdeDocumentos Then
                    ' ******************************************************
                    ' * Depósito não pode ser o último Documento do Malote *
                    ' * Enviando Malote para o Supervisor                  *
                    ' ******************************************************
                    .sSupervisor = "S"
                    Exit For
                ElseIf .sQualDocumento(.nId + 1) = "DE" Then
                    ' ***************************************
                    ' * Não pode haver 2 Depósitos seguidos *
                    ' * Enviando Malote para o Supervisor   *
                    ' ***************************************
                    .sSupervisor = "S"
                    Exit For
                End If
            ElseIf .bDeposito Then
                ' ***********************************
                ' * Documento depois de um Depósito *
                ' ***********************************
                If .sQualDocumento(.nId) = "CO" Or _
                    .nTipoDocto(.nId) = 4 Then
                    ' *************************************
                    ' * Documento fora de Ordem           *
                    ' * Enviando Malote para o Supervisor *
                    ' *************************************
                    .sSupervisor = "S"
                    Exit For
                ElseIf .sQualDocumento(.nId) <> "AD" And _
                        .sQualDocumento(.nId) <> "AC" Then
                    ' *************************************************
                    ' * Transformando Documento em Cheque de Deposito *
                    ' *************************************************
                    .nTipoDocto(.nId) = 7
                    .sAlcadaDocumento(.nId) = "N"
                    .sQualDocumento(.nId) = "CD"
                End If
            End If
            
            ' *****************************************
            ' * Totalizando Valores do Malote Empresa *
            ' *****************************************
            Select Case .sQualDocumento(.nId)
                Case "DE"
                    ' ************
                    ' * Depósito *
                    ' ************
                    .nValorContas = .nValorContas + .nValorTotal(.nId)
                Case "CO"
                    ' **********
                    ' * Contas *
                    ' **********
                    .nValorContas = .nValorContas + .nValorTotal(.nId)
                Case "CP"
                    ' *******************************
                    ' * ADCC ou Cheque de Pagamento *
                    ' *******************************
                    .nValorCheques = .nValorCheques + .nValorTotal(.nId)
                Case "CD"
                    ' ***********************
                    ' * Cheque de Depósitos *
                    ' ***********************
                    .nValorCheques = .nValorCheques + .nValorTotal(.nId)
                Case "AC"
                    ' *********************
                    ' * Ajuste de Crédito *
                    ' *********************
                    .nValorContas = .nValorContas + .nValorTotal(.nId)
                Case "AD"
                    ' ********************
                    ' * Ajuste de Débito *
                    ' ********************
                    .nValorCheques = .nValorCheques + .nValorTotal(.nId)
            End Select
        Next .nId
        
        If .sSupervisor = "S" Then
            ' *************************************
            ' * Enviar o Malote para o Supervisor *
            ' *************************************
            .sStatus = "3"
            .sSupervisor = "S"
            .nValorCheques = 0
            .nValorContas = 0
            
            AlterandoSituacaoMalote pvnDataProcessamento
            GoTo FimOK
        End If
    End With
    
    DoEvents
    
    ' *********************************************
    ' * Vinculando Depósitos Existentes no Malote *
    ' *********************************************
    If RegraMalote.bDeposito Then
        VinculandoDeposito pvbIgnorarProvaZero
        
        ' *********************************
        ' * Alterando o Tipo de Documento *
        ' * Para os Cheques de Depósito   *
        ' *********************************
        With RegraMalote
            sDocumentos = ""
            .sSql = "Update Documento Set TipoDocto = 7, Alcada = 'N' Where DataProcessamento = " & pvnDataProcessamento & " And IdLote = " & .nIdLote & " And IdDocto In ("
            
            For .nId = 1 To .nQtdeDocumentos
                If .sQualDocumento(.nId) = "CD" Then
                    If sDocumentos = "" Then
                        sDocumentos = .nIdDocto(.nId)
                    Else
                        sDocumentos = sDocumentos & "," & .nIdDocto(.nId)
                    End If
                End If
            Next .nId
            
            .sSql = .sSql & sDocumentos & ")"
        
            ' ***********************************************
            ' * Executando a Alteração do Tipo de Documento *
            ' * Para os Cheques de Depósito                 *
            ' ***********************************************
            Geral.Banco.Execute .sSql, rdExecDirect
        End With
    End If
    
    ' **********************************************
    ' * Vinculando Pagamentos Existentes no Malote *
    ' **********************************************
    VinculandoPagamento
    
    ' **************************************
    ' * Alterando o Vinculo dos Documentos *
    ' **************************************
    With RegraMalote
        .sSql = ""
        For .nId = 1 To .nQtdeDocumentos
            If Not .bVinculou(.nId) Then
                ' **********************************
                ' * Iniciando Vinculo da Transação *
                ' **********************************
                sDocumentos = ""
                .sSql = .sSql & "Update Documento Set Vinculo = " & .nVinculo(.nId) & " "
                .sSql = .sSql & "Where DataProcessamento = " & pvnDataProcessamento & " And "
                .sSql = .sSql & "IdLote = " & .nIdLote & " And "
                .sSql = .sSql & "IdDocto In ("
                For nInd = 1 To .nQtdeDocumentos
                    If Not .bVinculou(nInd) And _
                        .nVinculo(.nId) = .nVinculo(nInd) Then
                        ' ************************************
                        ' * Carregando Documentos a Vincular *
                        ' ************************************
                        .bVinculou(nInd) = True
                        If sDocumentos = "" Then
                            sDocumentos = .nIdDocto(nInd)
                        Else
                            sDocumentos = sDocumentos & "," & .nIdDocto(nInd)
                        End If
                    End If
                Next nInd
                .sSql = .sSql & sDocumentos & ")" & Chr(13)
            End If
        Next .nId
    
        ' ***********************************************************
        ' * Executa a Alteração do Vinculo para todos os Documentos *
        ' ***********************************************************
        Geral.Banco.Execute .sSql, rdExecDirect
        
        ' *****************************************************
        ' * Alterando o Tipo de Documento                     *
        ' * Para os Cheques a serem enviados para Compensação *
        ' *****************************************************
        sDocumentos = ""
        .sSql = ""
        
        For .nId = 1 To .nQtdeDocumentos
            If .nTipoDocto(.nId) = 6 Then
                If sDocumentos = "" Then
                    .sSql = "Update Documento Set TipoDocto = 6 Where DataProcessamento = " & pvnDataProcessamento & " And IdLote = " & .nIdLote & " And IdDocto In ("
                    sDocumentos = .nIdDocto(.nId)
                Else
                    sDocumentos = sDocumentos & "," & .nIdDocto(.nId)
                End If
            End If
        Next .nId
        
        If sDocumentos <> "" Then
            .sSql = .sSql & sDocumentos & ")"
            Geral.Banco.Execute .sSql, rdExecDirect
        End If
    End With
    
    ' ******************************************
    ' * Verificar Prova Zero no Malote Empresa *
    ' ******************************************
    With RegraMalote
        If Not pvbIgnorarProvaZero And _
            .nValorCheques <> .nValorContas Then
            For .nId = 1 To .nQtdeDocumentos
                If .nVinculo(.nId) = 0 Then
                    ' *************************************
                    ' * Enviar o Malote para a Prova Zero *
                    ' *************************************
                    .sStatus = "3"
                    .sPendenciaValor = "S"
                    AlterandoSituacaoMalote pvnDataProcessamento
                    GoTo FimOK
                End If
            Next .nId
        End If
    End With
    
    ' **************************************
    ' * Verificar Alçada no Malote Empresa *
    ' **************************************
    With RegraMalote
        For .nId = 1 To .nQtdeDocumentos
            If .sAlcadaDocumento(.nId) = "S" Then
                ' ***********************************************
                ' * Enviar o Malote para a Alçada de Documentos *
                ' ***********************************************
                .sStatus = "3"
                .sAlcada = "S"
                AlterandoSituacaoMalote pvnDataProcessamento
                GoTo FimOK
            End If
        Next .nId
    End With
        
    ' **********************************************
    ' * Verificar Vínculo Manual no Malote Empresa *
    ' **********************************************
    nQtdeContas = 0
    nQtdeCheques = 0
    
    With RegraMalote
        For .nId = 1 To .nQtdeDocumentos
            If .nVinculo(.nId) = 0 Then
                If .sQualDocumento(.nId) = "CO" Then
                    nQtdeContas = nQtdeContas + 1
                Else
                    nQtdeCheques = nQtdeCheques + 1
                End If
            End If
        Next .nId

        If nQtdeContas > 0 And _
            nQtdeCheques > 0 Then
            ' *****************************************
            ' * Enviar o Malote para o Vínculo Manual *
            ' *****************************************
            .sStatus = "3"
            .sVinculoManual = "S"
            AlterandoSituacaoMalote pvnDataProcessamento
            GoTo FimOK
        ElseIf nQtdeContas > 0 Or nQtdeCheques > 0 Then
            ' *****************************************************
            ' * Enviar o Malote para o Supervisor                 *
            ' * Pois existe somente Contas a Vincular Manualmente *
            ' * ou existe somente Cheques a Vincular Manualmente  *
            ' *****************************************************
            .sStatus = "3"
            .sSupervisor = "S"
            AlterandoSituacaoMalote pvnDataProcessamento
            GoTo FimOK
        End If
    End With
    
EnviarRobo:
    ' *************************************
    ' * Malote Empresa está Finalizado    *
    ' * Enviar para o Robô de Transmissão *
    ' *************************************
    RegraMalote.sStatus = "R"
    AlterandoSituacaoMalote pvnDataProcessamento
    
FimOK:
    ' *******************************************
    ' * Finalizando Transação do Banco de Dados *
    ' *******************************************
    Geral.Banco.CommitTrans
    
    ' ***********************************************************
    ' * Retornando Informação do Encerramento do Malote Empresa *
    ' ***********************************************************
    With RegraMalote
        If .sStatus = "1" Then
            ' *********************************
            ' * Malote Enviado para Digitação *
            ' *********************************
            EncerrarMalote = enRetornoMalote.enMaloteDigitar
        ElseIf .sStatus = "3" Then
            If .sSupervisor = "S" Then
                ' **********************************
                ' * Malote Enviado para Supervisor *
                ' **********************************
                EncerrarMalote = enRetornoMalote.enMaloteSupervisor
            ElseIf .sPendenciaValor = "S" Then
                ' **********************************
                ' * Malote Enviado para Prova Zero *
                ' **********************************
                EncerrarMalote = enRetornoMalote.enMaloteProvaZero
            ElseIf .sAlcada = "S" Then
                ' ******************************
                ' * Malote Enviado para Alçada *
                ' ******************************
                EncerrarMalote = enRetornoMalote.enMaloteAlcada
            ElseIf .sVinculoManual = "S" Then
                ' **************************************
                ' * Malote Enviado para Vinculo Manual *
                ' **************************************
                EncerrarMalote = enRetornoMalote.enMaloteVinculo
            End If
        ElseIf .sStatus = "R" Then
            ' ***********************************
            ' * Malote Enviado para Transmissão *
            ' ***********************************
            EncerrarMalote = enRetornoMalote.enMaloteRobo
        End If
    End With

    Exit Function

Erro:
    ' ******************************************
    ' * Cancelando Transação do Banco de Dados *
    ' ******************************************
    Geral.Banco.RollbackTrans
    
    EncerrarMalote = enRetornoMalote.enMaloteErro
End Function
' **********************************************
' * Executa o Vinculo Automático de Pagamentos *
' **********************************************
Public Sub VinculandoPagamento()
    Dim bNaoVinculo As Boolean          ' Marca se o Vinculo deve ou não ser efetuado
    Dim nInd As Integer                 ' Variável Auxiliar de Acesso ao Array de Documentos
    Dim nQtdeContas As Integer          ' Qtde de Contas a Serem Vinculadas
    Dim nQtdeCheques As Integer         ' Qtde de Cheques a Serem Vinculados
    Dim nQtdeSemVinculo As Integer      ' Qtde de Documentos sem Vinculo
    Dim nValorVinculo As Currency       ' Valor Apurado para Verificar o Vinculo
    Dim nPonteiroInicio As Integer      ' Ponteiro de Acesso ao Array
    Dim nPonteiroDesprezar As Integer   ' Ponteiro de Desprezao de Acesso ao Array
    Dim nConta As Integer               ' Variável Auxiliar para indicar a posição da Conta no Calculo
    Dim nCheque As Integer              ' Variável Auxiliar para indicar a posição da Conta no Calculo
    Dim wIndConta() As Integer          ' Array de Vínculo de Contas
    Dim wIndCheque() As Integer         ' Array de Vínculo de Cheques
    Dim nVinculo As Integer             ' Conteúdo do Vinculo
    Dim nDiferenca As Currency          ' Valor da Diferença dos Pagamentos
    
    With RegraMalote
        ' ****************************************
        ' * Marcar para Desprezar para o Vínculo *
        ' * Os Cheques com o mesmo Valor         *
        ' ****************************************
        For .nId = 1 To .nQtdeDocumentos
            nQtdeCheques = 0
            nQtdeContas = 0
            
            If .nVinculo(.nId) = 0 And _
                .nTipoDocto(.nId) > 3 And _
                .nTipoDocto(.nId) < 6 Then
                ' ***************************
                ' * Cheque/ADCC sem Vínculo *
                ' ***************************
                For nInd = 1 To .nQtdeDocumentos
                    If .nVinculo(nInd) = 0 And _
                        .nValorTotal(nInd) = .nValorTotal(.nId) Then
                        If .nTipoDocto(nInd) = .nTipoDocto(.nId) Then
                            ' ******************************
                            ' * Cheque/ADCC no Mesmo Valor *
                            ' ******************************
                            nQtdeCheques = nQtdeCheques + 1
                        ElseIf .sQualDocumento(nInd) = "CO" Then
                            ' ************************
                            ' * Conta no mesmo Valor *
                            ' ************************
                            nQtdeContas = nQtdeContas + 1
                        End If
                    End If
                Next nInd
            End If
            
            If nQtdeCheques = 1 And _
                nQtdeContas > 1 Then
                ' ***************************************
                ' * Desprezar para o Vínculo Automático *
                ' * 1 Cheque para mais de uma conta     *
                ' ***************************************
                .bDesprezarVinculo(.nId) = True
            ElseIf nQtdeCheques > 1 And _
                nQtdeContas > 0 Then
                ' *********************************************
                ' * Desprezar para o Vínculo Automático       *
                ' * Mais de um Cheque para uma ou mais contas *
                ' *********************************************
                .bDesprezarVinculo(.nId) = True
            End If
        Next .nId
        
        ' ***************************************
        ' * Primeira Fase do Vinculo            *
        ' * Vinculando Um Cheque para Uma Conta *
        ' ***************************************
        For .nId = 1 To .nQtdeDocumentos
            If .sQualDocumento(.nId) = "CP" And _
                .nVinculo(.nId) = 0 And _
                Not .bDesprezarVinculo(.nId) Then
                ' ****************************************
                ' * Cheque ou ADCC a Verificar o Vinculo *
                ' ****************************************
                For nInd = 1 To .nQtdeDocumentos
                    If .sQualDocumento(nInd) = "CO" And _
                        .nVinculo(nInd) = 0 And _
                        .nValorTotal(.nId) = .nValorTotal(nInd) Then
                        ' *********************************
                        ' * Vinculando Conta com o Cheque *
                        ' *********************************
                        .nVinculo(.nId) = .nIdDocto(.nId)
                        .nVinculo(nInd) = .nIdDocto(.nId)
                        Exit For
                    End If
                Next nInd
            End If
        Next .nId
        
        ' ******************************************
        ' * Verificando a Qtde de Contas no Malote *
        ' * A Serem Vinculadas                     *
        ' ******************************************
        nQtdeContas = 0
        
        For nInd = 1 To .nQtdeDocumentos
            If .sQualDocumento(nInd) = "CO" And _
               .nVinculo(nInd) = 0 Then
                nQtdeContas = nQtdeContas + 1
            End If
        Next nInd
        
        If nQtdeContas = 0 Then
            ' *****************************************
            ' * Não Existe Mais Documentos a Vincular *
            ' *****************************************
            Exit Sub
        End If
        
        ReDim wIndConta(nQtdeContas)
        
        ' ****************************************
        ' * Segunda Fase do Vinculo              *
        ' * Vinculando Um Cheque a várias Contas *
        ' ****************************************
        For .nId = 1 To .nQtdeDocumentos
            If .sQualDocumento(.nId) = "CP" And _
                .nVinculo(.nId) = 0 And _
                Not .bDesprezarVinculo(.nId) Then
                ' ****************************************
                ' * Cheque ou ADCC a Verificar o Vinculo *
                ' ****************************************
                nPonteiroInicio = 1
                nPonteiroDesprezar = 1
                
                Do While nPonteiroInicio <= nQtdeContas
                    nValorVinculo = 0
                    nConta = 0
                    
                    For nInd = 1 To nQtdeContas
                        wIndConta(nInd) = 0
                    Next nInd
                    
                    For nInd = 1 To .nQtdeDocumentos
                        If .sQualDocumento(nInd) = "CO" And _
                            .nVinculo(nInd) = 0 Then
                            nConta = nConta + 1
                            If nConta = nPonteiroInicio Or _
                                nConta > nPonteiroDesprezar Then
                                nValorVinculo = nValorVinculo + .nValorTotal(nInd)
                                wIndConta(nConta) = nInd
                                If nValorVinculo = .nValorTotal(.nId) Then
                                    Exit For
                                End If
                            End If
                        End If
                    Next nInd
                    
                    If nValorVinculo = .nValorTotal(.nId) Then
                        For nInd = 1 To nQtdeContas
                            If wIndConta(nInd) > 0 Then
                                .nVinculo(.nId) = .nIdDocto(.nId)
                                .nVinculo(wIndConta(nInd)) = .nIdDocto(.nId)
                            End If
                        Next nInd
                    End If
                    
                    nPonteiroDesprezar = nPonteiroDesprezar + 1
                    
                    If nPonteiroDesprezar > nQtdeContas Then
                        nPonteiroInicio = nPonteiroInicio + 1
                        nPonteiroDesprezar = nPonteiroInicio
                    End If
                Loop
            End If
        Next .nId
        
        ' *******************************************
        ' * Verificando a Qtde de Cheques no Malote *
        ' * A Serem Vinculados                      *
        ' *******************************************
        nQtdeCheques = 0
        
        For nInd = 1 To .nQtdeDocumentos
            If .nTipoDocto(nInd) > 4 And _
                .nTipoDocto(nInd) < 7 And _
                .nVinculo(nInd) = 0 And _
                Not .bDesprezarVinculo(nInd) Then
                nQtdeCheques = nQtdeCheques + 1
            End If
        Next nInd
        
        If nQtdeCheques = 0 Then
            ' *****************************************
            ' * Não Existe Mais Documentos a Vincular *
            ' *****************************************
            Exit Sub
        End If
        
        ReDim wIndCheque(nQtdeCheques)
        
        ' *****************************************
        ' * Terceira Fase do Vinculo              *
        ' * Vinculando Uma Conta a vários Cheques *
        ' *****************************************
        For .nId = 1 To .nQtdeDocumentos
            If .sQualDocumento(.nId) = "CO" And _
                .nVinculo(.nId) = 0 Then
                ' *******************************
                ' * Conta a Verificar o Vinculo *
                ' *******************************
                nPonteiroInicio = 1
                nPonteiroDesprezar = 1
                
                Do While nPonteiroInicio <= nQtdeCheques
                    nValorVinculo = 0
                    nCheque = 0
                    
                    For nInd = 1 To nQtdeCheques
                        wIndCheque(nInd) = 0
                    Next nInd
                    
                    For nInd = 1 To .nQtdeDocumentos
                        If .nTipoDocto(nInd) > 4 And _
                            .nTipoDocto(nInd) < 7 And _
                            .nVinculo(nInd) = 0 And _
                            Not .bDesprezarVinculo(nInd) Then
                            nCheque = nCheque + 1
                            If nCheque = nPonteiroInicio Or _
                                nCheque > nPonteiroDesprezar Then
                                nValorVinculo = nValorVinculo + .nValorTotal(nInd)
                                wIndCheque(nCheque) = nInd
                            End If
                            If nValorVinculo = .nValorTotal(.nId) Then
                                Exit For
                            End If
                        End If
                    Next nInd
                    
                    If nValorVinculo = .nValorTotal(.nId) Then
                        .nVinculo(.nId) = .nIdDocto(.nId)
                        For nInd = 1 To nQtdeCheques
                            If wIndCheque(nInd) > 0 Then
                                .nVinculo(wIndCheque(nInd)) = .nIdDocto(.nId)
                                If .nTipoDocto(wIndCheque(nInd)) = 5 Then
                                    .nTipoDocto(wIndCheque(nInd)) = 6
                                End If
                            End If
                        Next nInd
                    End If
                    
                    nPonteiroDesprezar = nPonteiroDesprezar + 1
                    
                    If nPonteiroDesprezar > nQtdeCheques Then
                        nPonteiroInicio = nPonteiroInicio + 1
                        nPonteiroDesprezar = nPonteiroInicio
                    End If
                Loop
            End If
        Next .nId
        
        ' *********************************
        ' * Verificar se ainda existe     *
        ' * Documentos a serem vinculados *
        ' *********************************
        nQtdeSemVinculo = 0
        
        For .nId = 1 To .nQtdeDocumentos
            If .sQualDocumento(.nId) = "CO" Or _
                .sQualDocumento(.nId) = "CP" Then
                If .nVinculo(.nId) = 0 And _
                    Not .bDesprezarVinculo(.nId) Then
                    nQtdeSemVinculo = nQtdeSemVinculo + 1
                End If
            End If
        Next .nId
        
        If nQtdeSemVinculo = 0 Then
            ' *****************************************
            ' * Não existe mais Documentos a Vincular *
            ' *****************************************
            Exit Sub
        End If
        
        ' *****************************************
        ' * Verificar se pode Vincular o Restante *
        ' *****************************************
        For .nId = 1 To .nQtdeDocumentos
            If .nTipoDocto(.nId) = 4 And _
                .nVinculo(.nId) = 0 Then
                ' **********************************
                ' * Contém ADCC Sem Vinculo        *
                ' * Não deverá vincular o restante *
                ' **********************************
                Exit Sub
            ElseIf .bDesprezarVinculo(.nId) And _
                .nVinculo(.nId) = 0 Then
                ' *************************************************
                ' * Contém mais de um Pagto UBB com o mesmo Valor *
                ' * Não deverá vincular o restante                *
                ' *************************************************
                Exit Sub
            End If
        Next .nId

        ' *********************************
        ' * Verificar se Existe Diferença *
        ' *********************************
        nDiferenca = 0
        
        For .nId = 1 To .nQtdeDocumentos
            If .nVinculo(.nId) = 0 Then
                If .sQualDocumento(.nId) = "CP" Then
                    nDiferenca = nDiferenca + .nValorTotal(.nId)
                ElseIf .sQualDocumento(.nId) = "CO" Then
                    nDiferenca = nDiferenca - .nValorTotal(.nId)
                End If
            End If
        Next .nId
        
        If Abs(nDiferenca) > 0 And _
            Abs(nDiferenca) > .nValorAjusteAuto Then
            ' ****************************************************
            ' * Diferença é Maior que o permitido pelo Parametro *
            ' * Não deverá vincular o restante                   *
            ' ****************************************************
            Exit Sub
        End If
        
        ' *****************************************
        ' * Quarta Fase do Vinculo                *
        ' * Vinculando n Contas com n Cheques     *
        ' *****************************************
        nVinculo = 0
        
        For .nId = 1 To .nQtdeDocumentos
            If .nVinculo(.nId) = 0 Then
                ' *******************************
                ' * Vinculando Contas e Cheques *
                ' *******************************
                If nVinculo = 0 Then
                    nVinculo = .nIdDocto(.nId)
                End If
                .nVinculo(.nId) = nVinculo
                If .nTipoDocto(.nId) = 5 Then
                    .nTipoDocto(.nId) = 6
                End If
            End If
        Next .nId
        
        If nDiferenca <> 0 Then
            ' ****************************************
            ' * Gravando Ajuste de Débito ou Crédito *
            ' ****************************************
            Set .qryInserirAjuste = Geral.Banco.CreateQuery("", "{Call InserirAjuste (?,?,?,?,?,?,?,?,?)}")
            With .qryInserirAjuste
                .rdoParameters(0).Value = Geral.DataProcessamento
                .rdoParameters(1).Value = RegraMalote.nIdLote
                .rdoParameters(2).Value = RegraMalote.nIdMalote
                .rdoParameters(3).Value = RegraMalote.nAgenciaMalote
                .rdoParameters(4).Value = RegraMalote.nContaMalote
                .rdoParameters(5).Value = nVinculo
                .rdoParameters(6).Value = Abs(nDiferenca)
                .rdoParameters(7).Value = IIf(nDiferenca > 0, 34, 38)
                .rdoParameters(8).Value = Format(RegraMalote.nAgenciaMalote, "0000") & Format(RegraMalote.nContaMalote, "0000000") & Trim(Geral.Usuario)
                .Execute
                .Close
            End With
            ' *******************************
            ' * Ajustando Valores do Malote *
            ' *******************************
            If nDiferenca > 0 Then
                .nValorContas = .nValorContas + nDiferenca
            Else
                .nValorCheques = .nValorCheques + Abs(nDiferenca)
            End If
        End If
    End With
End Sub
' *******************************
' * Recuperando Dados do Malote *
' *******************************
Public Sub SelecionandoMalote(ByVal pvnDataProcessamento As Double, _
                              ByVal pvnIdMalote As Double)
    With RegraMalote
        .sSql = ""
        .sSql = .sSql & "Select IdLote, "
        .sSql = .sSql & "       AgenciaMalote, "
        .sSql = .sSql & "       ContaMalote "
        .sSql = .sSql & "From Envelope "
        .sSql = .sSql & "Where DataProcessamento = " & pvnDataProcessamento & " And "
        .sSql = .sSql & "      IdEnvelope        = " & pvnIdMalote
        
        Set RegraMalote.rdoTB = Geral.Banco.OpenResultset(.sSql, rdOpenKeyset, rdConcurRowVer)
        
        If RegraMalote.rdoTB.EOF Then
            RegraMalote.rdoTB.Close
            Exit Sub
        End If
        
        .nIdMalote = pvnIdMalote
        .nIdLote = .rdoTB!IdLote
        .nAgenciaMalote = .rdoTB!AgenciaMalote
        .nContaMalote = .rdoTB!ContaMalote
        
        RegraMalote.rdoTB.Close
    End With
End Sub
' *************************************
' * Selecionando Documentos do Malote *
' *************************************
Public Sub SelecionandoDocumentos(ByVal pvnDataProcessamento As Double)
    With RegraMalote
        .sSql = ""
        .sSql = .sSql & "Select Ordem, "
        .sSql = .sSql & "       IdDocto, "
        .sSql = .sSql & "       ValorTotal, "
        .sSql = .sSql & "       TipoDocto, "
        .sSql = .sSql & "       Alcada, "
        .sSql = .sSql & "       Vinculo "
        .sSql = .sSql & "From Documento "
        .sSql = .sSql & "Where DataProcessamento = " & pvnDataProcessamento & " And "
        .sSql = .sSql & "      IdLote            = " & .nIdLote & " And "
        .sSql = .sSql & "      IdEnvelope        = " & .nIdMalote & " And "
        .sSql = .sSql & "      IdDocto           > 0" & " And "
        .sSql = .sSql & "      Status            = '1' And "
        .sSql = .sSql & "      TipoDocto        <> 1 "
        .sSql = .sSql & "Order by ordem,IdDocto"
        
        Set .rdoTB = Geral.Banco.OpenResultset(.sSql, rdOpenKeyset, rdConcurRowVer)
        
        .nQtdeDocumentos = RegraMalote.rdoTB.RowCount
        
        If .nQtdeDocumentos = 0 Then
            .rdoTB.Close
            Exit Sub
        End If
        
        ' **********************************************************
        ' * Criando Array Contendo os Documentos do Malote Empresa *
        ' **********************************************************
        ReDim .nIdDocto(1 To RegraMalote.rdoTB.RowCount)
        ReDim .nValorTotal(1 To RegraMalote.rdoTB.RowCount)
        ReDim .nTipoDocto(1 To RegraMalote.rdoTB.RowCount)
        ReDim .nVinculo(1 To RegraMalote.rdoTB.RowCount)
        ReDim .sAlcadaDocumento(1 To RegraMalote.rdoTB.RowCount)
        ReDim .sQualDocumento(1 To RegraMalote.rdoTB.RowCount)
        ReDim .bDesprezarVinculo(1 To RegraMalote.rdoTB.RowCount)
        ReDim .bVinculou(1 To RegraMalote.rdoTB.RowCount)
        
        ' ***************************************
        ' * Inicializando o Array de Documentos *
        ' ***************************************
        For .nId = 1 To RegraMalote.rdoTB.RowCount
            .nIdDocto(.nId) = 0
            .nValorTotal(.nId) = 0
            .nTipoDocto(.nId) = 0
            .nVinculo(.nId) = 0
            .sAlcadaDocumento(.nId) = ""
            .sQualDocumento(.nId) = ""
            .bDesprezarVinculo(.nId) = False
            .bVinculou(.nId) = False
        Next .nId
        
        Do While Not .rdoTB.EOF
            ' **********************************************
            ' * Carregando o Array de Documentos do Malote *
            ' **********************************************
            .nIdDocto(.rdoTB.AbsolutePosition) = .rdoTB!IdDocto
            .nValorTotal(.rdoTB.AbsolutePosition) = .rdoTB!ValorTotal
            .nTipoDocto(.rdoTB.AbsolutePosition) = .rdoTB!TipoDocto
            .nVinculo(.rdoTB.AbsolutePosition) = .rdoTB!Vinculo
            .sAlcadaDocumento(.rdoTB.AbsolutePosition) = .rdoTB!Alcada
            
            ' ********************************************
            ' * Setando o Tipo de Documento na Transação *
            ' ********************************************
            Select Case .rdoTB!TipoDocto
                Case 2, 3, 37
                    ' ************
                    ' * Depósito *
                    ' ************
                    .sQualDocumento(.rdoTB.AbsolutePosition) = "DE"
                 Case 4, 5, 6
                    ' *******************************
                    ' * ADCC ou Cheque de Pagamento *
                    ' *******************************
                    .sQualDocumento(.rdoTB.AbsolutePosition) = "CP"
                Case 7
                    ' **********************
                    ' * Cheque de Depósito *
                    ' **********************
                    .sQualDocumento(.rdoTB.AbsolutePosition) = "CD"
                Case 34
                    ' *********************
                    ' * Acerto de Crédito *
                    ' *********************
                    .sQualDocumento(.rdoTB.AbsolutePosition) = "AC"
                Case 38
                    ' ********************
                    ' * Acerto de Débito *
                    ' ********************
                    .sQualDocumento(.rdoTB.AbsolutePosition) = "AD"
                Case Else
                    ' **********
                    ' * Contas *
                    ' **********
                    .sQualDocumento(.rdoTB.AbsolutePosition) = "CO"
            End Select
            .rdoTB.MoveNext
        Loop
        .rdoTB.Close
    End With
End Sub
' ******************************************
' * Executa Vinculo Automático de Depósito *
' ******************************************
Public Sub VinculandoDeposito(ByVal pvbIgnorarProvaZero As Boolean)
    Dim nInd As Integer             ' Variável Auxiliar de Acesso ao Array de Documentos
    Dim wVinculo As Integer         ' Contem o Vinculo do Deposito
    Dim nSoma As Currency           ' Valor da soma dos cheques do Depósito

    With RegraMalote
        For .nId = 1 To .nQtdeDocumentos
            If .sQualDocumento(.nId) = "DE" Then
                ' ************************************
                ' * Foi Identificado um Depósito/OCT *
                ' ************************************
                If .nVinculo(.nId) <> 0 Then
                    ' *************************
                    ' * Depósito já Vinculado *
                    ' *************************
                    wVinculo = .nVinculo(.nId)
                Else
                    wVinculo = 0
                    nSoma = 0
                    For nInd = .nId + 1 To .nQtdeDocumentos
                        If .sQualDocumento(nInd) <> "CD" Then
                            Exit For
                        Else
                            ' ***********************************
                            ' * Somando Cheques do Depósito/OCT *
                            ' ***********************************
                            nSoma = nSoma + .nValorTotal(nInd)
                        End If
                    Next nInd
                    
                    If nSoma = .nValorTotal(.nId) Or pvbIgnorarProvaZero Then
                        ' *********************************
                        ' * Depósito deverá ser Vinculado *
                        ' *********************************
                        wVinculo = .nIdDocto(.nId)
                    End If
                End If
            End If
            
            If .sQualDocumento(.nId) = "DE" Or _
                .sQualDocumento(.nId) = "CD" Then
                ' ****************************
                ' * Vinculando os Documentos *
                ' ****************************
                .nVinculo(.nId) = wVinculo
            End If
        Next .nId
    End With
End Sub
' ************************************
' * Altera Altera Situação do Malote *
' ************************************
Private Sub AlterandoSituacaoMalote(ByVal pvnDataProcessamento As Double)
    With RegraMalote
        .sSql = ""
        .sSql = .sSql & "Update Envelope "
        .sSql = .sSql & "       Set Status = '" & .sStatus & "',"
        .sSql = .sSql & "           PendenciaValor = '" & .sPendenciaValor & "',"
        .sSql = .sSql & "           Alcada         = '" & .sAlcada & "',"
        .sSql = .sSql & "           Supervisor     = '" & .sSupervisor & "',"
        .sSql = .sSql & "           VinculoManual  = '" & .sVinculoManual & "',"
        .sSql = .sSql & "           Conta          = " & FormataValor(.nValorContas) & ","
        .sSql = .sSql & "           Dinheiro       = " & FormataValor(.nValorCheques) & ","
        .sSql = .sSql & "           Diferenca      = " & FormataValor(.nValorContas - .nValorCheques) & " "
        .sSql = .sSql & "Where DataProcessamento = " & pvnDataProcessamento & " And "
        .sSql = .sSql & "      IdEnvelope        = " & .nIdMalote
        
        Geral.Banco.Execute .sSql, rdExecDirect
    End With
End Sub
' *********************
' * Formantando Valor *
' *********************
Private Function FormataValor(ByVal pvnValor As Currency) As String
    Dim nInd As Integer
    Dim svalor As String
    
    svalor = pvnValor
    
    For nInd = 1 To Len(Trim(svalor))
        If Mid(svalor, nInd, 1) = "," Then
            Mid(svalor, nInd, 1) = "."
        End If
    Next nInd
    
    FormataValor = svalor
End Function
