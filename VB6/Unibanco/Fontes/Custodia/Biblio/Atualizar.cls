VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Atualizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function AtualizaBordero(ByVal pDataProcessamento As Long, _
                                ByVal pIdBordero As Long, _
                                ByVal pNum_Bordero As String, _
                                ByVal pAgencia As String, _
                                ByVal pContaCorrente As String, _
                                ByVal pCodigoCarteira As String, _
                                ByVal pCodigoLoja As String, _
                                ByVal pDataEntrada As Long, _
                                ByVal pNomeCliente As String, _
                                ByVal pSomaData As String, _
                                ByVal pSomaQuantidade As String, _
                                ByVal pSomaValor As String, _
                                ByVal pSomaTodos As String) As String

    Dim sStr        As String

           sStr = "UPDATE Bordero SET "
    sStr = sStr & "Num_Bordero = " & getParams(pNum_Bordero) & ", "
    sStr = sStr & "Agencia = " & Val(pAgencia) & ", "
    sStr = sStr & "Conta = " & Val(pContaCorrente) & ", "
    sStr = sStr & "CodigoCarteira = " & Val(pCodigoCarteira) & ", "
    sStr = sStr & "CodigoLoja = " & Val(pCodigoLoja) & ", "
    sStr = sStr & "DataEntrada = " & getParams(pDataEntrada) & ", "
    sStr = sStr & "NomeCliente = " & getParams(pNomeCliente) & ", "
    sStr = sStr & "SomaData = " & Val(pSomaData) & ", "
    sStr = sStr & "SomaQuantidade = " & Val(pSomaQuantidade) & ", "
    sStr = sStr & "SomaValor = " & Val(pSomaValor) & ", "
    sStr = sStr & "SomaTodos = " & Val(pSomaTodos) & ", "
    sStr = sStr & "HoraAtual = '" & Format(Now, "hh:mm:ss") & "'"
    sStr = sStr & " WHERE DataProcessamento = " & pDataProcessamento
    sStr = sStr & "   AND IdBordero = " & pIdBordero

    AtualizaBordero = sStr

End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Atualiza a Data de Deposito dos Cheques de Determinado IdBordero'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AtualizaDataDepositoCheques(ByVal pDataProcessamento As Long, _
                                            ByVal pIdBordero As Long, _
                                            ByVal pDataDepositoAtual As Long, _
                                            ByVal pDataDeposito As Long) As String

    Dim sStr        As String

           sStr = "UPDATE Cheque SET "
    sStr = sStr & "       DataDeposito = " & pDataDeposito
    sStr = sStr & " WHERE DataProcessamento = " & pDataProcessamento
    sStr = sStr & "   AND IdBordero = " & pIdBordero
    sStr = sStr & "   AND DataDeposito = " & pDataDepositoAtual
    
    AtualizaDataDepositoCheques = sStr
    

End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''
'Atualiza a Data de Deposito da Tabela DataDeposito'
''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'Não consegui determinar um nome legal para esta função,
'pode ser que haja conflito com outros processos que
'precisam desta função más que não faça exatamente isto
Public Function AtualizaDataDeposito(ByVal pDataProcessamento As Long, _
                                     ByVal pIdBordero As Long, _
                                     ByVal pDataDepositoAtual As Long, _
                                     ByVal pDataDeposito As Long, _
                                     ByVal pQuantidadeCheques As Integer, _
                                     ByVal pValorDeposito As String) As String

    Dim sStr        As String

           sStr = "UPDATE DataDeposito SET "
    sStr = sStr & "       DataDeposito = " & pDataDeposito & ","
    sStr = sStr & "       QuantidadeCheques = " & pQuantidadeCheques & ","
    sStr = sStr & "       ValorDeposito = " & pValorDeposito
    sStr = sStr & " WHERE DataProcessamento = " & pDataProcessamento
    sStr = sStr & "   AND IdBordero = " & pIdBordero
    sStr = sStr & "   AND DataDeposito = " & pDataDepositoAtual
    
    AtualizaDataDeposito = sStr
    

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Atualiza status dos cheques por DataProcessamento e IdBordero'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AtualizaStatusCheques(ByVal pDataProcessamento As Long, _
                                      ByVal pIdBordero As Long, _
                                      ByVal pStatus As String) As String

    Dim sStr        As String
    
           sStr = "UPDATE CHEQUE SET"
    sStr = sStr & "       CHEQUE.Status = '" & pStatus & "'"
    sStr = sStr & " WHERE CHEQUE.DataProcessamento = " & pDataProcessamento
    sStr = sStr & "   AND CHEQUE.IdBordero = " & pIdBordero
    sStr = sStr & "   AND CHEQUE.Status <> 'D' "
    
    AtualizaStatusCheques = sStr
    
    
End Function


Public Function AtualizaUsuario(Login As String, Nome As String, Cif As String, Senha As String) As String

Dim sStr  As String

     sStr = "update USUARIO set " & _
               "Nome = '" & Nome & "', " & _
               "Cif = '" & Cif & "', " & _
               "Senha = '" & Senha & "' " & _
               "where Login = '" & Login & "'"

     AtualizaUsuario = sStr

End Function
Public Function AtualizaStatusBordero(ByVal pDataProcessamento As Long, _
                                      ByVal pIdBordero As Long, _
                                      ByVal pStatus As String)

    Dim sStr As String
    
    sStr = "Update Bordero Set "
    sStr = sStr & "Status = '" & Trim(pStatus) & "',"
    sStr = sStr & "HoraAtual = '" & Time & "'"
    sStr = sStr & " WHERE DataProcessamento = " & pDataProcessamento
    sStr = sStr & "   AND IdBordero = " & pIdBordero
    
    AtualizaStatusBordero = sStr
    
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'               * Atualiza Dados do Cheque para o Borderô e IdCheque Atual *                '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AtualizaCheque(ByVal pDataProcessamento As Long, _
                               ByVal pIdBordero As Long, _
                               ByVal pIdcheque As Long, _
                               ByVal pDataDeposito As Long, _
                               ByVal pCMC7 As String, _
                               ByVal pcnpjcpf As String, _
                               ByVal pTipoInscricao As Integer, _
                               ByVal pvalor As String) As String

    Dim sStr As String

    sStr = "Update  Cheque Set "
    sStr = sStr & "        Cmc7= '" & pCMC7 & "'"
    sStr = sStr & "       ,Valor= '" & pvalor & "'"
    sStr = sStr & "       ,DataDeposito= " & pDataDeposito
    sStr = sStr & "       ,CNPJCPF= '" & pcnpjcpf & "'"
    sStr = sStr & "       ,PossuiErro=  No"
    sStr = sStr & "       ,Status=  1"
    sStr = sStr & "       ,TipoInscricao= '" & pTipoInscricao & "'"
    sStr = sStr & " WHERE DataProcessamento = " & pDataProcessamento
    sStr = sStr & "   AND IdBordero = " & pIdBordero
    sStr = sStr & "   AND IdCheque  = " & pIdcheque

    AtualizaCheque = sStr

End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'               * Se CPF estiver inválido grava erro na tabela de cheque *                  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AtualizaErroCheque(ByVal pDataProcessamento As Long, _
                                   ByVal pIdBordero As Long, _
                                   ByVal pIdcheque As Long) As String
                                                               
    Dim sStr As String

    sStr = "Update  Cheque Set "
    sStr = sStr & " PossuiErro = Yes"
    sStr = sStr & " WHERE DataProcessamento = " & pDataProcessamento
    sStr = sStr & "   AND IdBordero = " & pIdBordero
    sStr = sStr & "   AND IdCheque  = " & pIdcheque

    AtualizaErroCheque = sStr

End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                     * Atualiza Status do Borderô para Transmissão  *                      '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AtualizaStatusBorderoTransmissao(ByVal pDataProcessamento As Long _
                                               , ByVal pIdBordero As Long _
                                               , ByVal pStatus As String)

    Dim sStr As String
    
    sStr = "Update Bordero Set "
    sStr = sStr & "Status = '" & Trim(pStatus) & "'"
    sStr = sStr & " WHERE DataProcessamento = " & pDataProcessamento
    sStr = sStr & "   AND IdBordero = " & pIdBordero
    sStr = sStr & "   AND Status    = 'R'"
    
    AtualizaStatusBorderoTransmissao = sStr
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'          * Atualiza Numero de Remessa do Arquivo de Movimento*                            '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AtualizaNumRemessaMOV_Parametro(pDataProcessamento As Long _
                                              , pNumRemessa_MOV As Integer)

    Dim sStr As String

    sStr = "Update Parametro Set"
    sStr = sStr & " Num_Remessa_MOV = " & pNumRemessa_MOV + 1
    sStr = sStr & " WHERE DataProcessamento = " & pDataProcessamento

    AtualizaNumRemessaMOV_Parametro = sStr

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'          * Atualiza Numero de Remessa conforme a quantidade de Remessas do dia*           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AtualizaNumRemessaTER_Parametro(pDataProcessamento As Long _
                                              , pNumRemessa_TER As Integer)
    
    Dim sStr As String
    
    sStr = "Update Parametro Set"
    sStr = sStr & " Num_Remessa_TER = " & pNumRemessa_TER + 1
    sStr = sStr & " WHERE DataProcessamento = " & pDataProcessamento
    sStr = sStr & "   AND Num_Remessa_TER = " & pNumRemessa_TER
    
    AtualizaNumRemessaTER_Parametro = sStr
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'          * Atualiza Numero de Remessa conforme a quantidade de Remessas do dia*           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AtualizaNumVersaoCEL_Parametro(pDataProcessamento As Long _
                                              , pNumVersaoCEL As Integer)
    
    Dim sStr As String
    
    sStr = "Update Parametro Set"
    sStr = sStr & " Num_Versao_CEL = " & pNumVersaoCEL + 1
    sStr = sStr & " WHERE DataProcessamento = " & pDataProcessamento
    sStr = sStr & "   AND Num_Versao_CEL = " & pNumVersaoCEL
    
    AtualizaNumVersaoCEL_Parametro = sStr
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                     * Atualiza Numero de Remessa do Borderô Atual *                       '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AtualizaNumRemessaTER_Bordero(ByVal pDataProcessamento As Long _
                                           , ByVal pIdBordero As Long _
                                           , ByVal pNumRemessa_TER As Integer)

    Dim sStr As String
    
    sStr = "Update Bordero Set "
    sStr = sStr & "Num_Remessa_TER =" & pNumRemessa_TER
    sStr = sStr & " WHERE DataProcessamento = " & pDataProcessamento
    sStr = sStr & "   AND IdBordero = " & pIdBordero
    
    AtualizaNumRemessaTER_Bordero = sStr
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                     * Atualiza Numero de Remessa do Borderô Atual *                       '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AtualizaNumRemessaMOV_Bordero(ByVal pDataProcessamento As Long _
                                           , ByVal pIdBordero As Long _
                                           , ByVal pNumRemessa_MOV As Integer)

    Dim sStr As String

    sStr = "Update Bordero Set "
    sStr = sStr & "Num_Remessa =" & pNumRemessa_MOV
    sStr = sStr & " WHERE DataProcessamento = " & pDataProcessamento
    sStr = sStr & "   AND IdBordero = " & pIdBordero

    AtualizaNumRemessaMOV_Bordero = sStr

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Atualiza os parametros na tabela
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AtualizaParametros(ByVal pDataProcessamento As Long, _
                                   ByVal pQuantidadeCheques As Integer, _
                                   ByVal pQuantidadeDatas As Integer, _
                                   ByVal pDiretorioTransmissao As String, _
                                   ByVal pDiretorioRecepcao As String, _
                                   ByVal pCodigo_USB As Integer, _
                                   ByVal pCodigoAgAcolhed As Integer, _
                                   ByVal pCPD_Origem As String, _
                                   ByVal pCPD_Destino As String, _
                                   ByVal pCodigoTerceira As String, _
                                   ByVal pCNPJ_Terceira As String, _
                                   ByVal pSeq_Ocorrencia As Variant, _
                                   ByVal pUF_Terceira As String, _
                                   ByVal pCodigoAplicacao As String, _
                                   ByVal pValorChequeLimite As Currency, _
                                   ByVal pHeaderAV As String, _
                                   ByVal PchkSoma As String, _
                                   ByVal pGerarArquivo_CEL As String, _
                                   ByVal pComp_Origem_CEL As Integer, _
                                   ByVal pNum_Versao_Inicial_CEL As Integer, _
                                   ByVal pNum_Versao_Final_CEL As Integer, _
                                   ByVal pQuantidadeMinimaDias As String, _
                                   ByVal pCidade_Terceira As String, _
                                   ByVal pNome_Terceira As String) As String

    Dim sStr As String

    sStr = " Update Parametro set "
    sStr = sStr & " QuantidadeCheques = " & pQuantidadeCheques
    sStr = sStr & " ,QuantidadeDatas = " & pQuantidadeDatas
    sStr = sStr & " ,DiretorioTransmissao = '" & pDiretorioTransmissao & "'"
    sStr = sStr & " ,DiretorioRecepcao = '" & pDiretorioRecepcao & "'"
    sStr = sStr & " ,Codigo_USB = " & pCodigo_USB
    sStr = sStr & " ,CodigoAgAcolhed = " & pCodigoAgAcolhed
    sStr = sStr & " ,CPD_Origem = '" & pCPD_Origem & "'"
    sStr = sStr & " ,CPD_Destino = '" & pCPD_Destino & "'"
    sStr = sStr & " ,Codigo_Terceira = '" & pCodigoTerceira & "'"
    sStr = sStr & " ,CNPJ_Terceira = '" & pCNPJ_Terceira & "'"
    If pSeq_Ocorrencia <> vbNull Then
        sStr = sStr & " ,Seq_Ocorrencia = '" & pSeq_Ocorrencia & "'"
    End If
    sStr = sStr & " ,UF_Terceira = '" & pUF_Terceira & "'"
    sStr = sStr & " ,Cidade_Terceira = '" & pCidade_Terceira & "'"
    sStr = sStr & " ,Nome_Terceira = '" & pNome_Terceira & "'"
    sStr = sStr & " ,CodigoAplicacao = '" & pCodigoAplicacao & "'"
    sStr = sStr & " ,ValorChequeLimite = '" & pValorChequeLimite & "'"
    sStr = sStr & " ,HeaderAV =  " & pHeaderAV
    sStr = sStr & " ,CriticaSoma =  " & PchkSoma
    sStr = sStr & " ,QuantidadeMinimaDias = " & pQuantidadeMinimaDias
    sStr = sStr & " ,GerarArquivo_CEL = " & pGerarArquivo_CEL
    sStr = sStr & " ,Comp_Origem_CEL = " & pComp_Origem_CEL
    sStr = sStr & " ,Num_Versao_Inicial_CEL = " & pNum_Versao_Inicial_CEL
    sStr = sStr & " ,Num_Versao_Final_CEL = " & pNum_Versao_Final_CEL
    sStr = sStr & " Where DataProcessamento = " & pDataProcessamento

    AtualizaParametros = sStr

End Function

Public Function AtualizaStatusBorderoDePara(ByVal DataProcessamento As Long, ByVal IdBordero As Long, ByVal StatusPara As String, Optional ByVal StatusDe As String) As String
'
'    Atualiza status (DE) bordero para status (PARA)
'

Dim sStr  As String

     sStr = "Update BORDERO set Status = '" & StatusPara & "',"
     sStr = sStr & " HoraAtual = time() "
     sStr = sStr & " where DataProcessamento = " & DataProcessamento
     sStr = sStr & " and IdBordero = " & IdBordero
     
     If StatusDe <> "" Then
          sStr = sStr & " and Status = '" & StatusDe & "'"
     End If

AtualizaStatusBorderoDePara = sStr

End Function


Public Function AtualizaStatusCheque(ByVal DataProcessamento As Long, _
                                     ByVal IdBordero As Long, _
                                     ByVal IdCheque As Long, _
                                     ByVal Status As String) As String
'
'    Atualiza status do cheque
'

Dim sStr  As String

     sStr = "Update CHEQUE set Status = '" & Status & "'" & _
               " where DataProcessamento = " & DataProcessamento & _
               " and IdBordero = " & IdBordero & _
               " and IdCheque = " & IdCheque
     
AtualizaStatusCheque = sStr

End Function
Public Function AlteraStatusCheque(ByVal IdBordero As Long, _
                                   ByVal pCMC7 As String, _
                                   ByVal Status As String) As String
'
'    Alterar o Status do cheque
'

Dim sStr  As String

     sStr = "Update CHEQUE set Status = '" & Status & "'" & _
               " where IdBordero = " & IdBordero & _
               " and CMC7 = '" & pCMC7 & "'" & _
               " and Status <> 'D'"
     
AlteraStatusCheque = sStr

End Function


Public Function AtualizaHoraAtualBordero(ByVal DataProcessamento As Long, _
                                         ByVal IdBordero As Long, _
                                         ByVal sHoraAtual As String) As String
'
'    Atualiza HoraAtual do borderô para controle de registro pendente
'

Dim sStr  As String

     sStr = "Update BORDERO set HoraAtual = #" & sHoraAtual & "#" & _
               " where DataProcessamento = " & DataProcessamento & _
               " and IdBordero = " & IdBordero
     
AtualizaHoraAtualBordero = sStr

End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Atualiza status de cheque para (E)Confirmação de remessa '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AtualizaChequesRemessa(ByVal lngDataProcessamento As Long _
                                        , ByVal lngIdBordero As Long) As String

Dim sStr As String

               sStr = " update CHEQUE "
        sStr = sStr & " set Status = 'E' "
        sStr = sStr & " where DataProcessamento = " & lngDataProcessamento
        sStr = sStr & " and IdBordero = " & lngIdBordero
        sStr = sStr & " and Status    = 'T' "

AtualizaChequesRemessa = sStr

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Atualiza status da Data de Depósito para (1)Confirmação de remessa '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AtualizaStatusDataDeposito(ByVal lngDataProcessamento As Long _
                                        , ByVal lngIdBordero As Long) As String

Dim sStr As String

               sStr = " update DataDeposito "
        sStr = sStr & " set Status = '1' "
        sStr = sStr & " where DataProcessamento = " & lngDataProcessamento
        sStr = sStr & " and IdBordero = " & lngIdBordero
        sStr = sStr & " and Status <> 'I' "

AtualizaStatusDataDeposito = sStr

End Function


Public Function AtualizaFusao(pCMC7 As String, pDataProcessamento As Long) As String
    Dim sStr  As String

    sStr = " Update ChequeDataBoa"
    sStr = sStr & " Set Fusao = TRUE"
    sStr = sStr & " Where Dataprocessamento = " & pDataProcessamento
    sStr = sStr & " And CMC7 = '" & pCMC7 & "'"
    
AtualizaFusao = sStr

End Function

Public Function AtualizaGeracao_Ter(ByVal pCMC7 As String, ByVal pNum_Remessa_TER As Integer, ByVal pDataProcessamento As Long) As String
    Dim sStr  As String

    sStr = " Update ChequeDataBoa"
    sStr = sStr & " Set Geracao_Ter = TRUE"
    sStr = sStr & " ,Num_Remessa_TER = " & pNum_Remessa_TER
    sStr = sStr & " Where Dataprocessamento = " & pDataProcessamento
    sStr = sStr & " And CMC7 = '" & pCMC7 & "'"
    
AtualizaGeracao_Ter = sStr

End Function


Public Function AtualizaGeracao_CEL(ByVal pCMC7 As String, ByVal pNum_Remessa_TER As Integer, ByVal pDataProcessamento As Long) As String
    Dim sStr  As String

    sStr = " Update ChequeDataBoa"
    sStr = sStr & " Set Geracao_CEL = TRUE"
    sStr = sStr & " ,Num_Versao_CEL = " & pNum_Remessa_TER
    sStr = sStr & " Where Dataprocessamento = " & pDataProcessamento
    sStr = sStr & " And CMC7 = '" & pCMC7 & "'"
    
AtualizaGeracao_CEL = sStr

End Function

Public Function AtualizaDiasCheques(pDataProcessamento As Long _
                                       , pQuantidadeMinimaDias As Integer)
    
    Dim sStr As String
    
    sStr = "Update Parametro Set"
    sStr = sStr & " QuantidadeMinimaDias = " & pQuantidadeMinimaDias
    sStr = sStr & " WHERE DataProcessamento = " & pDataProcessamento
    
    AtualizaDiasCheques = sStr
    
End Function

Public Function AtualizaStatusAviso(ByVal pDataOcorrencia As Long, _
                                      ByVal pCodigoOcorrencia As Long, _
                                      ByVal pStatus As String)

    Dim sStr As String
    
    sStr = "Update AvisoDiferenca Set "
    sStr = sStr & "Status = '" & Trim(pStatus) & "'"
    sStr = sStr & " WHERE DataOcorrencia = " & pDataOcorrencia
    sStr = sStr & "   AND CodigoOcorrencia = " & pCodigoOcorrencia
    
    AtualizaStatusAviso = sStr
    
End Function
