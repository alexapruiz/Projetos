VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Inserir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Insere o arquivo gerado dos cheques indevidos na tabela Aviso Diferença'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function InsereArquivoAD(ByVal pPathArquivoAD As String)

    Dim sStr        As String
    
           sStr = "INSERT INTO AvisoDiferenca "
    sStr = sStr & "SELECT *"
    sStr = sStr & "  FROM [AvisoDiferenca.TXT] "
    sStr = sStr & "    IN '" & pPathArquivoAD & "'"
    sStr = sStr & "[text;];"

    InsereArquivoAD = sStr

End Function


Public Function InsereAvisoDiferenca(ByVal pDataOcorrencia As Long, _
                                     ByVal pCodigoOcorrencia As Long, _
                                     ByVal pDescricaoMotivoAD As String, _
                                     ByVal pDataDeposito As Long, _
                                     ByVal pBordero As String, _
                                     ByVal pCodigoCarteira As String, _
                                     ByVal pAgencia As Integer, _
                                     ByVal pConta As Long, _
                                     ByVal pCodigoDevolucao As Integer, _
                                     ByVal pCodigoCompensacao As Integer, _
                                     ByVal pBancoEmitente As Integer, _
                                     ByVal pAgenciaEmitente As Integer, _
                                     ByVal pCcEmitente As Double, _
                                     ByVal pNrChequeEmitente As Long, _
                                     ByVal pTipoCheque As String, _
                                     ByVal pTipoInscricao As String, _
                                     ByVal pInscricaoEmitente As String, _
                                     ByVal pvalor As String, ByVal pGerado As String, _
                                     ByVal pStatus As String) As String
                                      
                                       

    Dim sStr        As String
    
           sStr = "INSERT INTO AvisoDiferenca"
    
    sStr = sStr & "("
    sStr = sStr & "             DataOcorrencia,"
    sStr = sStr & "             CodigoOcorrencia,"
    sStr = sStr & "             CodigoMotivo,"
    sStr = sStr & "             DataDeposito,"
    sStr = sStr & "             Num_Bordero,"
    sStr = sStr & "             CodigoCarteira,"
    sStr = sStr & "             Agencia,"
    sStr = sStr & "             Conta,"
    sStr = sStr & "             CodigoDevolucao,"
    sStr = sStr & "             CodigoCompensacao,"
    sStr = sStr & "             BancoEmitente,"
    sStr = sStr & "             AgenciaEmitente,"
    sStr = sStr & "             CcEmitente,"
    sStr = sStr & "             NrChequeEmitente,"
    sStr = sStr & "             TipoCheque,"
    sStr = sStr & "             TipoInscricao,"
    sStr = sStr & "             InscricaoEmitente,"
    sStr = sStr & "             Valor,"
    sStr = sStr & "             Gerado,"
    sStr = sStr & "             Status"
    
    sStr = sStr & ")"
    
    sStr = sStr & " VALUES "
    sStr = sStr & "("
    sStr = sStr & pDataOcorrencia & ","
    sStr = sStr & pCodigoOcorrencia & ","
    sStr = sStr & "'" & pDescricaoMotivoAD & "',"
    sStr = sStr & pDataDeposito & ","
    sStr = sStr & "'" & pBordero & "',"
    sStr = sStr & pCodigoCarteira & ","
    sStr = sStr & pAgencia & ","
    sStr = sStr & pConta & ","
    sStr = sStr & pCodigoDevolucao & ","
    sStr = sStr & pCodigoCompensacao & ","
    sStr = sStr & pBancoEmitente & ","
    sStr = sStr & pAgenciaEmitente & ","
    sStr = sStr & pCcEmitente & ","
    sStr = sStr & pNrChequeEmitente & ","
    sStr = sStr & pTipoCheque & ","
    sStr = sStr & pTipoInscricao & ","
    sStr = sStr & pInscricaoEmitente & ","
    sStr = sStr & "'" & pvalor & "',"
    sStr = sStr & pGerado & ","
    sStr = sStr & "'" & pStatus & "'"
    sStr = sStr & ")"
    
    InsereAvisoDiferenca = sStr

End Function


Public Function InsereBordero(ByVal pDataProcessamento As Long, _
                              ByVal pBordero As String, _
                              ByVal pAgencia As String, _
                              ByVal pContaCorrente As String, _
                              ByVal pCodigoCarteira As String, _
                              ByVal pCodigoLoja As String, _
                              ByVal pDataEntradaBordero As Long, _
                              ByVal pNomeCliente As String, _
                              ByVal pStatus As String, _
                              ByVal pSomaData As String, _
                              ByVal pSomaQuantidade As String, _
                              ByVal pSomaValor As String, _
                              ByVal pSomaTodos As String)

    Dim sStr        As String
    
           sStr = "INSERT INTO BORDERO "
    sStr = sStr & "("
    sStr = sStr & "       DataProcessamento, "
    sStr = sStr & "       Num_Bordero, "
    sStr = sStr & "       Agencia, "
    sStr = sStr & "       Conta, "
    sStr = sStr & "       CodigoCarteira, "
    sStr = sStr & "       CodigoLoja, "
    sStr = sStr & "       DataEntrada, "
    sStr = sStr & "       NomeCliente, "
    sStr = sStr & "       Status, "
    sStr = sStr & "       SomaData, "
    sStr = sStr & "       SomaQuantidade, "
    sStr = sStr & "       SomaValor, "
    sStr = sStr & "       SomaTodos,"
    sStr = sStr & "       HoraAtual"
    sStr = sStr & ")"
    sStr = sStr & " VALUES "
    sStr = sStr & "("

    sStr = sStr & pDataProcessamento & ", "
    sStr = sStr & getParams(pBordero) & ", "
    sStr = sStr & pAgencia & ", "
    sStr = sStr & pContaCorrente & ", "
    sStr = sStr & pCodigoCarteira & ", "
    sStr = sStr & pCodigoLoja & ", "
    sStr = sStr & pDataEntradaBordero & ", "
    sStr = sStr & getParams(pNomeCliente) & ", "
    sStr = sStr & getParams(pStatus) & ", "
    sStr = sStr & pSomaData & ", "
    sStr = sStr & pSomaQuantidade & ", "
    sStr = sStr & pSomaValor & ", "
    sStr = sStr & pSomaTodos & ", '"
    sStr = sStr & Format(Now, "hh:mm:ss") & "'"
    sStr = sStr & ")"

    InsereBordero = sStr

End Function
Public Function InsereDataDeposito(ByVal pDataProcessamento As Long, _
                                   ByVal pIdBordero As Long, _
                                   ByVal pDataDeposito As Long, _
                                   ByVal pQuantidadeCheques As String, _
                                   ByVal pValorDeposito As String) As String

    Dim sStr        As String
        
           sStr = "INSERT INTO DataDeposito "
    sStr = sStr & "("
    sStr = sStr & tiraAOA(getParams("DataProcessamento", _
                                    "IdBordero", _
                                    "DataDeposito", _
                                    "QuantidadeCheques", _
                                    "ValorDeposito"))
    sStr = sStr & ")"
    sStr = sStr & " SELECT "
    sStr = sStr & AcertaTD(tiraAOA(getParams(pDataProcessamento, _
                                             pIdBordero, _
                                             pDataDeposito, _
                                             pQuantidadeCheques, _
                                             pValorDeposito)))
    
    InsereDataDeposito = sStr



End Function
Public Function InsereGrupoUsuario(IdGrupo As Integer, IdUsuario As String) As String

Dim sStr  As String

     sStr = "insert into GRUPOUSUARIO (IdUsuario, IdGrupo) values (" & _
               IdUsuario & "," & _
               IdGrupo & ")"

     InsereGrupoUsuario = sStr

End Function
Public Function InsereUsuario(Login As String, Nome As String, Cif As String, Senha As String) As String

Dim sStr  As String

     sStr = "insert into USUARIO (Login, Nome, CIF, Senha) values (" & _
               "'" & Login & "'," & _
               "'" & Nome & "'," & _
               "'" & Cif & "'," & _
               "'" & Senha & "')"

     InsereUsuario = sStr

End Function
Public Function InsereCheque(ByVal pDataProcessamento As Long, _
                             ByVal pIdBordero As Long, _
                             ByVal pDataDeposito As Long, _
                             ByVal pCMC7 As String, _
                             ByVal pcnpjcpf As String, _
                             ByVal pTipoInscricao As Integer, _
                             ByVal pvalor As String _
                             ) As String

    Dim sStr As String

          sStr = "Insert Into Cheque(DataProcessamento"
                     sStr = sStr & ",idBordero"
                     sStr = sStr & ",DataDeposito"
                     sStr = sStr & ",Cmc7"
                     sStr = sStr & ",CNPJCPF"
                     sStr = sStr & ",PossuiErro"
                     sStr = sStr & ",TipoInscricao"
                     sStr = sStr & ",Status"
                     sStr = sStr & ",Valor)"

          sStr = sStr & "Values (" & _
                            pDataProcessamento & "," & _
                            "'" & pIdBordero & "'," & _
                            pDataDeposito & "," & _
                            "'" & pCMC7 & "'," & _
                            "'" & pcnpjcpf & "', No , " & _
                            "'" & pTipoInscricao & "'," & _
                            "" & 1 & "," & _
                            "'" & pvalor & "')"
   InsereCheque = sStr
    
End Function
Public Function InsereRejeicaoRemessa(ByVal pDataProcessamento As Long, _
                                        ByVal pIdBordero As Long, _
                                        ByVal pDataDeposito As Long, _
                                        ByVal pIdcheque As Long, _
                                        ByVal pCodErro As Integer, _
                                        ByVal pRotuloOriginal As String) As String

Dim sStr As String
    
           sStr = "insert into RejeicaoRemessa ("
    sStr = sStr & " DataProcessamento,"
    sStr = sStr & " IdBordero,"
    sStr = sStr & " DataDeposito,"
    sStr = sStr & " IdCheque,"
    sStr = sStr & " CodigoErro,"
    sStr = sStr & " RotuloOriginal)"
    sStr = sStr & " Values ("
    sStr = sStr & pDataProcessamento & ","
    sStr = sStr & pIdBordero & ","
    sStr = sStr & pDataDeposito & ","
    sStr = sStr & pIdcheque & ","
    sStr = sStr & pCodErro & ","
    sStr = sStr & "'" & pRotuloOriginal & "')"

InsereRejeicaoRemessa = sStr

End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                * Insere Nova Data de Processamento na Tabela de Parâmetros *                 '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function InsereParametro(ByVal pDataProcessamento As Long, _
                                ByVal pQuantidadeCheques As Integer, _
                                ByVal pQuantidadeDatas As Integer, _
                                ByVal pQuantidadeMinimaDias As String, _
                                ByVal pDiretorioTransmissao As String, _
                                ByVal pDiretorioRecepcao As String, _
                                ByVal pSeq_Bordero As Integer, _
                                ByVal pGerarArquivo_CEL As String, _
                                ByVal pNumLoteCEL As Integer, _
                                ByVal pComp_Origem_CEL As Integer, _
                                ByVal pNum_Versao_Inicial_CEL As Integer, _
                                ByVal pNum_Versao_Final_CEL As Integer, _
                                ByVal pHeaderAV As String, ByVal pchkSoma As String, _
                                ByVal pCodigo_USB As Integer, _
                                ByVal pCPD_Origem As String, _
                                ByVal pCPD_Destino As String, _
                                ByVal pCodigoTerceira As String, _
                                ByVal pCNPJ_Terceira As String, _
                                ByVal pSeq_Ocorrencia As Long, _
                                ByVal pCodigoAgAcolhed As Integer, _
                                ByVal pNum_Remessa_TER As Integer, _
                                ByVal pValorChequeLimite As Currency, _
                                ByVal pTmp_Pendente As Integer, _
                                ByVal pUF_Terceira As String, _
                                ByVal pCodigoAplicacao As String, ByVal pDiasLimpeza As Integer, pCidadeTerceira As String, pNomeTerceira As String) As String

    Dim sStr As String

           sStr = " Insert Into Parametro"
    sStr = sStr & "     (DataProcessamento,"
    sStr = sStr & "     QuantidadeCheques,"
    sStr = sStr & "     QuantidadeDatas,"
    sStr = sStr & "     QuantidadeMinimaDias,"
    sStr = sStr & "     DiretorioTransmissao,"
    sStr = sStr & "     DiretorioRecepcao,"
    sStr = sStr & "     Seq_Bordero,"
    sStr = sStr & "     GerarArquivo_CEL,"
    sStr = sStr & "     Num_Lote_CEL,"
    sStr = sStr & "     Comp_Origem_CEL,"
    sStr = sStr & "     Num_Versao_Inicial_CEL,"
    sStr = sStr & "     Num_Versao_Final_CEL,"
    sStr = sStr & "     HeaderAV,"
    sStr = sStr & "     CriticaSoma,"
    sStr = sStr & "     Codigo_USB,"
    sStr = sStr & "     CPD_Origem,"
    sStr = sStr & "     CPD_Destino,"
    sStr = sStr & "     Codigo_Terceira,"
    sStr = sStr & "     CNPJ_Terceira,"
    sStr = sStr & "     Seq_Ocorrencia,"
    sStr = sStr & "     CodigoAgAcolhed,"
    sStr = sStr & "     Num_Remessa_TER,"
    sStr = sStr & "     ValorChequeLimite,"
    sStr = sStr & "     TMP_Pendente,"
    sStr = sStr & "     Num_Versao_CEL,"
    sStr = sStr & "     Num_Remessa_Mov,"
    sStr = sStr & "     UF_Terceira,"
    sStr = sStr & "     CodigoAplicacao,"
    sStr = sStr & "     DiasLimpeza,"
    sStr = sStr & "     Cidade_Terceira,"
    sStr = sStr & "     Nome_Terceira)"
    sStr = sStr & " Values ("
    sStr = sStr & pDataProcessamento & " ,"
    sStr = sStr & pQuantidadeCheques & ","
    sStr = sStr & pQuantidadeDatas & ","
    sStr = sStr & pQuantidadeMinimaDias & ","
    sStr = sStr & "'" & pDiretorioTransmissao & "',"
    sStr = sStr & "'" & pDiretorioRecepcao & "',"
    sStr = sStr & pSeq_Bordero & ","
    sStr = sStr & pGerarArquivo_CEL & ","
    sStr = sStr & pNumLoteCEL & ","
    sStr = sStr & pComp_Origem_CEL & ","
    sStr = sStr & pNum_Versao_Inicial_CEL & ","
    sStr = sStr & pNum_Versao_Final_CEL & ","
    sStr = sStr & pHeaderAV & ","
    sStr = sStr & pchkSoma & ","
    sStr = sStr & pCodigo_USB & ","
    sStr = sStr & "'" & pCPD_Origem & "',"
    sStr = sStr & "'" & pCPD_Destino & "',"
    sStr = sStr & "'" & pCodigoTerceira & "',"
    sStr = sStr & pCNPJ_Terceira & ","
    sStr = sStr & pSeq_Ocorrencia & ","
    sStr = sStr & pCodigoAgAcolhed & ","
    sStr = sStr & pNum_Remessa_TER & ","
    sStr = sStr & "'" & pValorChequeLimite & "',"
    sStr = sStr & pTmp_Pendente & ","
    sStr = sStr & "0,"
    sStr = sStr & "0,"
    sStr = sStr & "'" & pUF_Terceira & "',"
    sStr = sStr & "'" & pCodigoAplicacao & "',"
    sStr = sStr & "'" & pDiasLimpeza & "','"
    sStr = sStr & pCidadeTerceira & "','"
    sStr = sStr & pNomeTerceira & "')"

    InsereParametro = sStr
    
End Function
Public Function InsereInconsistencia(ByVal strDataChave As String, ByVal strHoraChave As String, ByVal strDescricao As String) As String

Dim sStr As String

     sStr = "Insert into Inconsistencia (DataChave, HoraChave, Descricao) values ( "
     sStr = sStr & "'" & strDataChave & "','" & strHoraChave & "','" & strDescricao & "')"

InsereInconsistencia = sStr

End Function
Public Function InsereAlteracaoData(ByVal lngDataProcessamento As Long, ByVal strNum_Bordero As String, _
                                   ByVal bytCodigoCarteira As Byte, ByVal lngDataAnterior As Long, _
                                   ByVal intBancoEmitente As Integer, ByVal intAgencia As Integer, _
                                   ByVal DblContaCorrente As Double, ByVal lngNumCheque As Long, _
                                   ByVal lngDataNova As Long, ByVal strNossoNumero As String, _
                                   ByVal intCodCompens As Integer, ByVal pvalor As String) As String

Dim sStr As String

     sStr = "Insert into AlteracaoData "
     sStr = sStr & "(DataProcessamento, Num_Bordero, CodigoCarteira, DataAnterior, BancoEmitente, "
     sStr = sStr & "AgenciaEmitente, CcEmitente, NrChequeEmitente, DataNova, NossoNumero, CodigoCompensacao, Valor) "
     sStr = sStr & "values ( "
     sStr = sStr & lngDataProcessamento & ", '"
     sStr = sStr & strNum_Bordero & "', "
     sStr = sStr & bytCodigoCarteira & ", "
     sStr = sStr & lngDataAnterior & ", "
     sStr = sStr & intBancoEmitente & ", "
     sStr = sStr & intAgencia & ", "
     sStr = sStr & DblContaCorrente & ", "
     sStr = sStr & lngNumCheque & ", "
     sStr = sStr & lngDataNova & ", '"
     sStr = sStr & strNossoNumero & "', "
     sStr = sStr & intCodCompens & ", '"
     sStr = sStr & pvalor & "')"
     
     
InsereAlteracaoData = sStr

End Function









Public Function InsereChequesBaixados(ByVal strNum_Bordero As String, ByVal intBancoEmitente As Integer, _
                                      ByVal intAgencia As Integer, ByVal lngNumCheque As Long, _
                                      ByVal DblContaCorrente As Double, ByVal strNossoNumero As String, _
                                      ByVal intCodCompens As Integer, ByVal lngDataDeposito As Long, _
                                      ByVal pvalor As Currency, ByVal lngDataBaixa As Long, ByVal bytCodigoCarteira As Byte)

                                      

Dim sStr As String

     sStr = "Insert into ChequesBaixados "
     sStr = sStr & "(Num_Bordero, BancoEmitente, AgenciaEmitente, NrChequeEmitente, "
     sStr = sStr & "CcEmitente,  NossoNumero, CodigoCompensacao, DataDeposito,ValorCheque,DataBaixa, CodigoCarteira )"
     sStr = sStr & "values ('"
     sStr = sStr & Trim(strNum_Bordero) & "', "
     sStr = sStr & intBancoEmitente & ", "
     sStr = sStr & intAgencia & ", "
     sStr = sStr & lngNumCheque & ", "
     sStr = sStr & DblContaCorrente & ", '"
     sStr = sStr & Trim(strNossoNumero) & "', "
     sStr = sStr & intCodCompens & ", "
     sStr = sStr & lngDataDeposito & ", '"
     sStr = sStr & pvalor & "', "
     sStr = sStr & lngDataBaixa & ", "
     sStr = sStr & bytCodigoCarteira & ")"
     
InsereChequesBaixados = sStr

End Function





Public Function InsChDataBoa(ByVal pDataProcessamento As Long, _
                            ByVal pCMC7 As String, _
                            ByVal pccg_end As String, _
                            ByVal pDataDeposito As Long, _
                            ByVal pBordero As String, _
                            ByVal pCodigoCarteira As Byte, _
                            ByVal pcnpjcpf As String, _
                            ByVal pAgencia As Integer, _
                            ByVal pConta As Double, _
                            ByVal pvalor As String, _
                            ByVal pTipoCheque As Integer, _
                            ByVal pChequeSuperior As String, _
                            ByVal pFusao As String, _
                            ByVal pGeracao_CEL As String, _
                            ByVal pGeracao_TER As String, _
                            ByVal pNum_Versao_CEL As Integer, _
                            ByVal pNum_Lote_CEL As Integer, _
                            ByVal pNum_Remessa_TER As Integer _
                            ) As String

    Dim sStr As String

            sStr = "Insert Into ChequeDataBoa(DataProcessamento"
            sStr = sStr & ",CMC7"
            sStr = sStr & ",CGC_Enderecamento"
            sStr = sStr & ",DataDeposito"
            sStr = sStr & ",Num_Bordero"
            sStr = sStr & ",CodigoCarteira"
            sStr = sStr & ",CNPJCPF"
            sStr = sStr & ",Agencia"
            sStr = sStr & ",Conta"
            sStr = sStr & ",Valor"
            sStr = sStr & ",TipoCheque"
            sStr = sStr & ",ChequeSuperior"
            sStr = sStr & ",Fusao"
            sStr = sStr & ",Geracao_CEL"
            sStr = sStr & ",Geracao_TER"
            sStr = sStr & ",Num_Versao_CEL"
            sStr = sStr & ",Num_Lote_CEL"
            sStr = sStr & ",Num_Remessa_TER )"
            
          sStr = sStr & " Values (" & _
                            pDataProcessamento & "," & _
                            "'" & pCMC7 & "'," & _
                            "'" & pccg_end & "'," & _
                            pDataDeposito & "," & _
                            "'" & pBordero & "'," & _
                            pCodigoCarteira & "," & _
                            "'" & pcnpjcpf & "'," & _
                            pAgencia & "," & _
                            pConta & "," & _
                            "'" & pvalor & "'," & _
                            pTipoCheque & "," & _
                            pChequeSuperior & "," & _
                            pFusao & "," & _
                            pGeracao_CEL & "," & _
                            pGeracao_TER & "," & _
                            pNum_Versao_CEL & "," & _
                            pNum_Lote_CEL & "," & _
                            pNum_Remessa_TER & ")"
                            
   InsChDataBoa = sStr
    
End Function

Public Function InsereComplementoExportacao(ByVal sReferencia As String, ByVal lDataLayOut As Long, ByVal sRotulo As String, ByVal iDelimitador As Integer, ByVal iFimDeLinha As Integer, ByVal sUsuario As String, ByVal sNomeArquivo As String, ByVal sDiretorio As String) As String

    Dim sStr        As String
    
           sStr = "INSERT INTO ComplementoExportacao ("
    sStr = sStr & "Referencia, "
    sStr = sStr & "DataLayOut,"
    sStr = sStr & "Rotulo,"
    sStr = sStr & "Delimitador,"
    sStr = sStr & "FimDeLinha,"
    sStr = sStr & "NomeArquivo,"
    sStr = sStr & "Diretorio,"
    sStr = sStr & "Usuario) "
    sStr = sStr & "Values ("
    sStr = sStr & "'" & sReferencia & "', "
    sStr = sStr & lDataLayOut & ", "
    sStr = sStr & "'" & sRotulo & "', "
    sStr = sStr & iDelimitador & ", "
    sStr = sStr & iFimDeLinha & ", "
    sStr = sStr & "'" & sNomeArquivo & "', "
    sStr = sStr & "'" & sDiretorio & "', "
    sStr = sStr & "'" & Trim(sUsuario) & "'" & ")"

    InsereComplementoExportacao = sStr

End Function
Public Function InsereCamposExportacao(ByVal sReferencia As String, ByVal iOrdenacao As Integer, ByVal sNome As String, ByVal iTamanho As Integer, ByVal iAlinhamento As Integer, ByVal iZeros As Integer, ByVal sSiglaTabela As String) As String

    Dim sStr        As String
    
           sStr = "INSERT INTO CamposExportacao ("
    sStr = sStr & "Referencia, "
    sStr = sStr & "Ordenacao, "
    sStr = sStr & "Nome,"
    sStr = sStr & "Tamanho,"
    sStr = sStr & "Alinhamento,"
    sStr = sStr & "Zeros,"
    sStr = sStr & "SiglaTabela) "
    sStr = sStr & "Values ("
    sStr = sStr & "'" & sReferencia & "', "
    sStr = sStr & iOrdenacao & ", "
    sStr = sStr & "'" & sNome & "', "
    sStr = sStr & iTamanho & ", "
    sStr = sStr & iAlinhamento & ", "
    sStr = sStr & iZeros & ", "
    sStr = sStr & "'" & sSiglaTabela & "')"
    InsereCamposExportacao = sStr

End Function









Public Function InsereRegraGP(ByVal pDataProcessamento As Long, _
                             ByVal pCodidoProduto As Long, _
                             ByVal pCodigoRegraGP As Long, _
                             ByVal pQtdDias As Long _
                             ) As String

    Dim sStr As String

          sStr = "Insert Into RegraGP(DataProcessamento"
                     sStr = sStr & ",CodigoProduto"
                     sStr = sStr & ",CodigoRegra"
                     sStr = sStr & ",QtdDias)"

          sStr = sStr & "Values (" & _
                            pDataProcessamento & ", " & _
                            pCodidoProduto & ", " & _
                            pCodigoRegraGP & ", " & _
                            pQtdDias & ")"
   InsereRegraGP = sStr
    
End Function
