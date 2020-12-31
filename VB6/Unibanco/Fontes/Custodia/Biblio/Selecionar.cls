VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Selecionar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function GetCapaLimpeza(ByVal pDataProcessamento As Long) As String

Dim sStr As String

sStr = "SELECT * FROM Bordero"
sStr = sStr & " WHERE DataProcessamento < " & Trim(pDataProcessamento)
sStr = sStr & " AND Status <> 'E' AND Status <> 'D'"

GetCapaLimpeza = sStr

End Function

Public Function GetChequesIndevidos(ByVal pDataProcessamento As Long, _
                                    ByVal pIdBordero As Long, _
                                    ByVal pDataDeposito As Long) As String

    Dim sStr        As String
    
           sStr = "SELECT IdCheque, "
    sStr = sStr & "       CMC7, "
    sStr = sStr & "       CNPJCPF, "
    sStr = sStr & "       TipoInscricao, "
    sStr = sStr & "       PossuiErro, "
    sStr = sStr & "       Status, "
    sStr = sStr & "       Valor"
    sStr = sStr & "  FROM Cheque"
    sStr = sStr & " WHERE DataProcessamento = " & pDataProcessamento
    sStr = sStr & "   AND IdBordero = " & pIdBordero
    sStr = sStr & "   AND DataDeposito = " & pDataDeposito

    GetChequesIndevidos = sStr

End Function

Public Function GetNaoTransmitidos(ByVal pDataProcessamento As Long) As String

Dim sStr As String

sStr = "SELECT * FROM Bordero"
'sStr = sStr & " WHERE DataProcessamento < " & Trim(pDataProcessamento)
sStr = sStr & " Where Status <> 'E' AND Status <> 'D' AND Status <> 'T'"

GetNaoTransmitidos = sStr

End Function

Public Function GetChequeDataBoa(ByVal pDataProcessamento As Long, _
                                 Optional ByVal pCMC7 As String) As String
Dim sStr As String

sStr = "SELECT * FROM ChequeDataBoa"
sStr = sStr & " WHERE ChequeDataBoa.DataDeposito = " & Trim(Str(pDataProcessamento))

If Len(Trim(pCMC7)) <> 0 Then
   sStr = sStr & " And ChequeDataBoa.CMC7 = '" & pCMC7 & "'"
End If

GetChequeDataBoa = sStr

End Function

Public Function GetDataProcessamento(ByVal pDiasLimpeza As Long) As String

Dim sStr As String

sStr = " SELECT TOP " & Str(pDiasLimpeza) & " Parametro.DataProcessamento"
sStr = sStr & " FROM Parametro"
sStr = sStr & " ORDER BY DataProcessamento DESC"

GetDataProcessamento = sStr

End Function
Public Function GetChequesBaixados(Optional ByVal pNumBordero As String _
                            , Optional ByVal pDteAnterior As Long _
                            , Optional ByVal pBanco As String _
                            , Optional ByVal pAgencia As String _
                            , Optional ByVal pContac As String _
                            , Optional ByVal pNrCheque As String _
                            ) As String
    
    Dim sStr   As String
    Dim sWhere As String
        
    sStr = " SELECT ChequesBaixados.Num_Bordero"
    sStr = sStr & "       ,ChequesBaixados.DataDeposito "
    sStr = sStr & "       ,ChequesBaixados.AgenciaEmitente "
    sStr = sStr & "       ,ChequesBaixados.BancoEmitente "
    sStr = sStr & "       ,ChequesBaixados.CcEmitente "
    sStr = sStr & "       ,ChequesBaixados.NrChequeEmitente "
    sStr = sStr & "       ,ChequesBaixados.ValorCheque "
    sStr = sStr & "  FROM  ChequesBaixados "
    
    sWhere = ""
    
    '''''''''''''''''''''''''''''''''''''''''''
    '* Pesquisa por Identificação do Bordero *'
    '''''''''''''''''''''''''''''''''''''''''''
    If Len(Trim(pNumBordero)) <> 0 Then
        If sWhere = "" Then
           sWhere = "   Where  ChequesBaixados.Num_Bordero = '" & Trim(pNumBordero) & "'"
        Else
           sWhere = sWhere & "   And   ChequesBaixados.Num_Bordero = '" & pNumBordero & "'"
        End If
    End If
   
    '''''''''''''''''''''''''''''''''''''''''''
    '     * Pesquisa por Data Anterior *      '
    '''''''''''''''''''''''''''''''''''''''''''
    If pDteAnterior <> 0 Then
       If sWhere = "" Then
          sWhere = "   Where   ChequesBaixados.DataDeposito = " & pDteAnterior
       Else
          sWhere = sWhere & "   And   ChequesBaixados.DataDeposito = " & pDteAnterior
       End If
    End If

    
    '''''''''''''''''''''''''''''''''''''''''''
    '          * Pesquisa por Banco *         '
    '''''''''''''''''''''''''''''''''''''''''''
    If Len(Trim(pBanco)) <> 0 Then
       If sWhere = "" Then
          sWhere = "   Where   ChequesBaixados.BancoEmitente = " & pBanco
       Else
          sWhere = sWhere & "   And   ChequesBaixados.BancoEmitente = " & pBanco
       End If
    End If

    '''''''''''''''''''''''''''''''''''''''''''
    '        * Pesquisa por Agencia *         '
    '''''''''''''''''''''''''''''''''''''''''''
    If Len(Trim(pAgencia)) <> 0 Then
       If sWhere = "" Then
          sWhere = "   Where   ChequesBaixados.AgenciaEmitente = " & pAgencia
       Else
          sWhere = sWhere & "   And   ChequesBaixados.AgenciaEmitente = " & pAgencia
       End If
    End If

    '''''''''''''''''''''''''''''''''''''''''''
    '          * Pesquisa por Conta *         '
    '''''''''''''''''''''''''''''''''''''''''''
    If Len(Trim(pContac)) <> 0 Then
       If sWhere = "" Then
          sWhere = "   Where   ChequesBaixados.CcEmitente = " & pContac
       Else
          sWhere = sWhere & "   And   ChequesBaixados.CcEmitente = " & pContac
       End If
    End If

    '''''''''''''''''''''''''''''''''''''''''''
    '    * Pesquisa por Número do Cheque *    '
    '''''''''''''''''''''''''''''''''''''''''''
    If Len(Trim(pNrCheque)) <> 0 Then
       If sWhere = "" Then
          sWhere = sWhere & "   Where   ChequesBaixados.NrChequeEmitente = " & pNrCheque
       Else
          sWhere = sWhere & "   And   ChequesBaixados.NrChequeEmitente = " & pNrCheque
       End If
    End If
    
    GetChequesBaixados = sStr & sWhere

End Function

Public Function GetDataBaixa(Optional ByVal pDataBaixa As Long) As String

    Dim sStr   As String
    Dim sWhere As String
            
    sWhere = ""
    
    sStr = " SELECT ChequesBaixados.Num_Bordero"
    sStr = sStr & "  FROM  ChequesBaixados "
    
    If pDataBaixa <> 0 Then
       sWhere = "   Where   ChequesBaixados.DataBaixa = " & pDataBaixa
    End If

    GetDataBaixa = sStr & sWhere

End Function



Public Function GetChequesBordero(ByVal pDataProcessamento As Long, _
                                  ByVal pIdBordero As Long, _
                                  ByVal pDataDeposito As Long) As String

    Dim sStr        As String
    
           sStr = "SELECT "
    sStr = sStr & "       Cheque.IdCheque, "
    sStr = sStr & "       Cheque.CMC7, "
    sStr = sStr & "       Cheque.CNPJCPF, "
    sStr = sStr & "       Cheque.TipoInscricao, "
    sStr = sStr & "       Cheque.Status, "
    sStr = sStr & "       Cheque.Valor "
    sStr = sStr & "  FROM Cheque"
    sStr = sStr & " WHERE DataProcessamento = " & pDataProcessamento
    sStr = sStr & "   AND IdBordero = " & pIdBordero
    sStr = sStr & "   AND DataDeposito = " & pDataDeposito
    sStr = sStr & "   AND Status NOT IN ('D','I')"
    
    GetChequesBordero = sStr

End Function
Public Function GetChequesIndevidosQTDE(ByVal pDataProcessamento As Long, _
                                        ByVal pIdBordero As Long) As String
    Dim sStr As String
   
           sStr = "SELECT IdCheque,"
    sStr = sStr & "       DataDeposito,"
    sStr = sStr & "       CMC7, "
    sStr = sStr & "       Valor"
    sStr = sStr & "  FROM CHEQUE"
    sStr = sStr & " WHERE CHEQUE.DataProcessamento = " & pDataProcessamento
    sStr = sStr & "   AND CHEQUE.IdBordero = " & pIdBordero
    sStr = sStr & "   AND CHEQUE.Status <> 'D'"
    sStr = sStr & " ORDER BY CHEQUE.DataDeposito ASC, IdCheque DESC"
    
    GetChequesIndevidosQTDE = sStr

End Function
Public Function GetChequesIndevidosDATA(ByVal pDataProcessamento As Long, _
                                        ByVal pIdBordero As Long) As String
    Dim sStr        As String
    
    sStr = "SELECT DATADEPOSITO.DataDeposito AS Quantidade "
    sStr = sStr & "  FROM DATADEPOSITO"
    sStr = sStr & " WHERE DATADEPOSITO.DataProcessamento = " & pDataProcessamento
    sStr = sStr & " AND DATADEPOSITO.IdBordero = " & pIdBordero
    sStr = sStr & " AND DATADEPOSITO.Status <> 'D'"
    sStr = sStr & " GROUP BY DATADEPOSITO.DataDeposito"
    sStr = sStr & " ORDER BY DATADEPOSITO.DataDeposito"

    GetChequesIndevidosDATA = sStr

End Function
Public Function GetDataDepositoBordero(ByVal pDataProcessamento As Long, _
                                       ByVal pIdBordero As Long, _
                                       ByVal pDataDeposito As Long) As String

    Dim sStr As String
    
           sStr = "SELECT "
    sStr = sStr & "       DataDeposito.DataProcessamento, "
    sStr = sStr & "       DataDeposito.IdBordero, "
    sStr = sStr & "       DataDeposito.DataDeposito, "
    sStr = sStr & "       DataDeposito.QuantidadeCheques, "
    sStr = sStr & "       DataDeposito.ValorDeposito"
    sStr = sStr & "  FROM DataDeposito"
    sStr = sStr & " WHERE DataProcessamento = " & pDataProcessamento
    sStr = sStr & "   AND IdBordero = " & pIdBordero
    sStr = sStr & "   AND DataDeposito = " & pDataDeposito
    
    GetDataDepositoBordero = sStr

End Function
Public Function GetDetalhesArquivoCEL_605(ByVal pDataProcessamento As Long, _
                                          ByVal pTipoCheque As Integer, _
                                          ByVal pTipoExportacao As Integer)
    
    Dim sStr        As String

           sStr = "SELECT ChequeDataBoa.DataProcessamento, "
    sStr = sStr & "       ChequeDataBoa.CGC_Enderecamento, "
    sStr = sStr & "       ChequeDataBoa.DataDeposito, "
    sStr = sStr & "       ChequeDataBoa.Num_Bordero, "
    sStr = sStr & "       ChequeDataBoa.CodigoCarteira, "
    sStr = sStr & "       ChequeDataBoa.CMC7, "
    sStr = sStr & "       ChequeDataBoa.Valor, "
    sStr = sStr & "       ChequeDataBoa.Fusao, "
    sStr = sStr & "       ChequeDataBoa.Geracao_CEL, "
    sStr = sStr & "       ChequeDataBoa.Geracao_TER"
    sStr = sStr & "  FROM ChequeDataBoa"
    sStr = sStr & " WHERE DataProcessamento = " & pDataProcessamento
    sStr = sStr & "   AND TipoCheque = " & pTipoCheque
    
    If pTipoExportacao = 0 Or pTipoExportacao = 1 Then
        sStr = sStr & "  AND ChequeSuperior = " & CInt(CBool(pTipoExportacao))
    End If
    sStr = sStr & " ORDER BY CMC7"
    
    GetDetalhesArquivoCEL_605 = sStr

End Function

Public Function GetEstatisticaBordero(ByVal pDataProcessamento As Long) As String

    Dim Str As String

    Str = "Select Count(Bordero.IdBordero) as QtdBordero , Bordero.Status "
    Str = Str & " From Bordero "
    Str = Str & " Where Bordero.DataProcessamento = " & pDataProcessamento
    Str = Str & " Group By Bordero.Status "

    GetEstatisticaBordero = Str
End Function
Public Function GetEstatisticaCheque(ByVal pDataProcessamento As Long) As String

    Dim Str As String

    Str = "Select Count(Cheque.IdCheque) as QtdCheque , Bordero.Status "
    Str = Str & " FROM Bordero LEFT JOIN Cheque ON "
    Str = Str & " (Bordero.IdBordero = Cheque.IdBordero) AND (Bordero.DataProcessamento = Cheque.DataProcessamento)"
    Str = Str & " Where Bordero.DataProcessamento = " & pDataProcessamento
    Str = Str & " Group By Bordero.Status"

    GetEstatisticaCheque = Str
    
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Seleciona bordero pela DataProcessamento e IdBordero'
''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetIdBordero(ByVal pDataProcessamento As Long, _
                             ByVal pIdBordero As String, _
                             Optional ByVal pTmp_Pendente As Long = 0) As String

    Dim sStr        As String

           sStr = "SELECT Bordero.DataProcessamento, "
    sStr = sStr & "       Bordero.IdBordero, "
    sStr = sStr & "       Bordero.Num_Bordero, "
    sStr = sStr & "       Bordero.Agencia, "
    sStr = sStr & "       Bordero.Conta, "
    sStr = sStr & "       Bordero.CodigoCarteira, "
    sStr = sStr & "       Bordero.CodigoLoja, "
    sStr = sStr & "       Bordero.DataEntrada, "
    sStr = sStr & "       Bordero.NomeCliente, "
    sStr = sStr & "       Bordero.Status, "
    sStr = sStr & "       Bordero.SomaData, "
    sStr = sStr & "       Bordero.SomaQuantidade, "
    sStr = sStr & "       Bordero.SomaValor, "
    sStr = sStr & "       Bordero.SomaTodos, "
    sStr = sStr & "       Bordero.HoraAtual"
    sStr = sStr & "  FROM Bordero"
    sStr = sStr & " WHERE IdBordero = " & Val(pIdBordero)
    sStr = sStr & "   AND DataProcessamento = " & pDataProcessamento
    
    If pTmp_Pendente > 0 Then
        sStr = sStr & "   AND (Bordero.Status = '5' OR "
        sStr = sStr & "        Bordero.Status = 'X' OR "
        sStr = sStr & "       (Bordero.Status = 'H' AND "
        sStr = sStr & "        DateDiff('s',Bordero.HoraAtual,'" & Time & "') <= " & pTmp_Pendente & ")"
        sStr = sStr & "       )"
    End If
    
    GetIdBordero = sStr

End Function
Public Function GetMotivoAD(ByVal pCodigoMotivo As Long) As String

    Dim sStr    As String
    
           sStr = "SELECT Descricao"
    sStr = sStr & "  FROM MotivosAD"
    sStr = sStr & " WHERE CodigoMotivo = " & pCodigoMotivo

    GetMotivoAD = sStr

End Function

Public Function GetCodMotivoAD(ByVal sDescricao As String) As String

    Dim sStr    As String
    
           sStr = "SELECT CodigoMotivo"
    sStr = sStr & "  FROM MotivosAD"
    sStr = sStr & " WHERE Descricao = '" & sDescricao & "'"

    GetCodMotivoAD = sStr

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Seleciona bordero pela DataProcessamento e NumeroBordero'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetNumeroBordero(ByVal pDataProcessamento As Long _
                               , ByVal pNumeroBordero As String) As String

    Dim sStr        As String

           sStr = "SELECT Bordero.DataProcessamento, "
    sStr = sStr & "       Bordero.IdBordero, "
    sStr = sStr & "       Bordero.Num_Bordero, "
    sStr = sStr & "       Bordero.Agencia, "
    sStr = sStr & "       Bordero.Conta, "
    sStr = sStr & "       Bordero.CodigoCarteira, "
    sStr = sStr & "       Bordero.CodigoLoja, "
    sStr = sStr & "       Bordero.DataEntrada, "
    sStr = sStr & "       Bordero.NomeCliente, "
    sStr = sStr & "       Bordero.Status, "
    sStr = sStr & "       Bordero.SomaData, "
    sStr = sStr & "       Bordero.SomaQuantidade, "
    sStr = sStr & "       Bordero.SomaValor, "
    sStr = sStr & "       Bordero.SomaTodos "
    sStr = sStr & "  FROM Bordero"
    sStr = sStr & " WHERE Num_Bordero = " & getParams(pNumeroBordero)
    sStr = sStr & "   AND DataProcessamento = " & pDataProcessamento
    sStr = sStr & "   AND Status <> 'D'"

    GetNumeroBordero = sStr

End Function
Public Function GetAvisoDiferenca(ByVal pDataProcessamento As Long, _
                                  ByVal pNumeroBordero As String _
                                  ) As String

    Dim sStr        As String

    sStr = "SELECT * FROM AvisoDiferenca"
    sStr = sStr & " WHERE Num_Bordero = " & getParams(pNumeroBordero)
    sStr = sStr & "   AND DataOcorrencia = " & pDataProcessamento
    
    
    GetAvisoDiferenca = sStr

End Function




Public Function GetRecAvisoDiferenca(ByVal pDataProcessamento As Long _
              , ByVal pGerado As Boolean) As String

    Dim sStr        As String

    sStr = "SELECT * FROM AvisoDiferenca"
    sStr = sStr & " WHERE DataOcorrencia = " & pDataProcessamento
    sStr = sStr & "   AND Gerado = " & pGerado
   
    
    GetRecAvisoDiferenca = sStr
    

End Function
Public Function GetCarteira(Optional ByVal pCodigoCarteira As Byte = 0) As String

    Dim sStr        As String
    
    
           sStr = "SELECT Carteira.CodigoCarteira, "
    sStr = sStr & "       Carteira.Descricao"
    sStr = sStr & "  FROM Carteira"
    
    If pCodigoCarteira <> 0 Then
        sStr = sStr & " WHERE CodigoCarteira = " & pCodigoCarteira
    End If
    
    
    GetCarteira = sStr

End Function
''''''''''''''''''''''''''''''''''''
'Pega o tipo de feriado e descricao'
''''''''''''''''''''''''''''''''''''
Public Function GetFeriado(ByVal pDataFeriado As String, _
                           ByVal pAgencia As String) As String

    Dim sStr    As String
    
           sStr = "SELECT Feriado.DataFeriado, "
    sStr = sStr & "       Feriado.PCC, "
    sStr = sStr & "       Feriado.Descricao, "
    sStr = sStr & "       Feriado.TipoFeriado"
    sStr = sStr & "  FROM Feriado"
    sStr = sStr & " WHERE DataFeriado = " & pDataFeriado
    sStr = sStr & "   AND PCC = " & pAgencia
    
    GetFeriado = sStr

End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Busca o Codigo e Descricao da RejeicaoRemessa fazendo JOIN com MotivoRejeicao'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetRejeicoesCheque(ByVal pDataProcessamento As Long, _
                                   ByVal pIdBordero As Long, _
                                   ByVal pDataDeposito As Long, _
                                   ByVal pIdcheque As Long) As String

    Dim sStr        As String
    
           sStr = "SELECT RejeicaoRemessa.CodigoErro, "
    sStr = sStr & "       RejeicaoRemessa.RotuloOriginal, "
    sStr = sStr & "       MotivoRejeicao.Descricao "
    sStr = sStr & "  FROM RejeicaoRemessa, MotivoRejeicao "
    sStr = sStr & " WHERE DataProcessamento = " & pDataProcessamento
    sStr = sStr & "   AND IdBordero = " & pIdBordero
    sStr = sStr & "   AND DataDeposito = " & pDataDeposito
    sStr = sStr & "   AND IdCheque = " & pIdcheque
    sStr = sStr & "   AND MotivoRejeicao.CodigoErro = RejeicaoRemessa.CodigoErro"
    
    GetRejeicoesCheque = sStr

End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Seleciona um registro com a somatoria dos cheques por '
'DataProcessamento, IdBordero e DataDeposito é Opcional'
'onde o envio da DataDeposito Retornar somente um reg  '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetSomatoriaChequesBordero(ByVal pDataProcessamento As Long, _
                                           ByVal pIdBordero As Long, _
                                  Optional ByVal pDataDeposito As Long = 0) As String

    Dim sStr        As String

           sStr = "SELECT "
    sStr = sStr & "        A.Dataprocessamento, "
    sStr = sStr & "        A.Idbordero, "
    sStr = sStr & "        A.DataDeposito, "
    sStr = sStr & "        COUNT(A.valor) AS QtdeCheques, "
    sStr = sStr & "        SUM(A.valor) AS TotCheques"
    sStr = sStr & "  FROM Cheque AS A"
    sStr = sStr & " WHERE DataProcessamento = " & pDataProcessamento
    sStr = sStr & "   AND IdBordero = " & pIdBordero
    If Trim(pDataDeposito) <> 0 Then
        sStr = sStr & "   AND DataDeposito = " & pDataDeposito
    End If
    sStr = sStr & " GROUP BY A.Dataprocessamento, "
    sStr = sStr & "          A.Idbordero, "
    sStr = sStr & "          A.DataDeposito"
    sStr = sStr & " ORDER BY A.Dataprocessamento, "
    sStr = sStr & "          A.DataDeposito;"
    
    GetSomatoriaChequesBordero = sStr

End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Seleciona todos os registros da Tabela StatusBordero'
''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetStatusBordero() As String

    Dim sStr        As String
    
           sStr = "SELECT"
    sStr = sStr & "       Status, "
    sStr = sStr & "       Descricao "
    sStr = sStr & "  FROM StatusBordero"

    GetStatusBordero = sStr

End Function

'
'Seleciona os borderos para Supervisor
'
Public Function GetSupervisor(ByVal pDataProcessamento As Long, _
                              ByVal pTmp_Pendente As Long) As String

    Dim sStr        As String
    
           sStr = "SELECT Bordero.DataProcessamento, "
    sStr = sStr & "       Bordero.IdBordero, "
    sStr = sStr & "       Bordero.Num_Bordero, "
    sStr = sStr & "       Bordero.Agencia, "
    sStr = sStr & "       Bordero.Conta, "
    sStr = sStr & "       Bordero.CodigoCarteira, "
    sStr = sStr & "       Bordero.CodigoLoja, "
    sStr = sStr & "       Bordero.DataEntrada, "
    sStr = sStr & "       Bordero.NomeCliente, "
    sStr = sStr & "       Bordero.Status, "
    sStr = sStr & "       Bordero.SomaData, "
    sStr = sStr & "       Bordero.SomaQuantidade, "
    sStr = sStr & "       Bordero.SomaValor, "
    sStr = sStr & "       Bordero.SomaTodos, "
    sStr = sStr & "       Bordero.HoraAtual "
    sStr = sStr & "  FROM Bordero "
    sStr = sStr & " WHERE Bordero.DataProcessamento = " & pDataProcessamento
    sStr = sStr & "   AND (Bordero.Status = '5' OR "
    sStr = sStr & "        Bordero.Status = 'X' OR "
    sStr = sStr & "      ((Bordero.Status = 'H' AND "
    sStr = sStr & "        DateDiff('s',Bordero.HoraAtual,'" & Time & "') >= " & pTmp_Pendente & ") OR"
    sStr = sStr & "       (Bordero.Status = '2' AND "
    sStr = sStr & "        DateDiff('s',Bordero.HoraAtual,'" & Time & "') >= " & pTmp_Pendente & "))"
    sStr = sStr & "       )"
    
    GetSupervisor = sStr

End Function
Public Function GetTipoCheque(Optional ByVal pTipoCheque As Integer = -1) As String

    
    Dim sStr            As String
    
    
           sStr = "SELECT TipoCheque.TipoCheque,"
    sStr = sStr & "       TipoCheque.Descricao"
    sStr = sStr & "  FROM TipoCheque"
    If pTipoCheque >= 0 Then
        sStr = sStr & " WHERE TipoCheque = " & pTipoCheque
    End If
    
    GetTipoCheque = sStr

End Function
Public Function GetTodosGrupos() As String

    Dim sStr  As String

    sStr = "select IdGrupo, Descricao From GRUPO order by IdGrupo"
     
    GetTodosGrupos = sStr
     
End Function
Public Function GetTotalCapa(ByVal pDataProcessamento As Long) As String

    Dim Str As String

    Str = "select Count(0) as TotalCapa from Bordero "
    Str = Str & " Where DataProcessamento = " & pDataProcessamento

    GetTotalCapa = Str
End Function
Public Function GetTotalDocto(ByVal pDataProcessamento As Long) As String

    Dim Str As String

    Str = "select Count(0) as TotalDocto from Cheque "
    Str = Str & " Where DataProcessamento = " & pDataProcessamento

    GetTotalDocto = Str
End Function
Public Function GetUsuario(Optional Login) As String

Dim sStr  As String

     sStr = "select IdUsuario, Login, Nome, Cif, Senha From USUARIO "
     
     If Not IsMissing(Login) Then
          sStr = sStr & "where Login = '" & Login & "' "
     End If
     sStr = sStr & "order by Login"
     
     GetUsuario = sStr

End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'*  Seleciona dados de Parâmetro para a Data de Processamento Atual ou se a variável Flag  *'
'*  estiver preenchida retorna a maior  Data de Processamento gravada na Base de Dados     *'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetParametros(ByVal pDataProcessamento As Long, _
                     Optional ByVal pFlag As Integer) As String

    Dim sStr        As String
    
           sStr = "SELECT Parametro.DataProcessamento, "
    sStr = sStr & "       Parametro.QuantidadeCheques, "
    sStr = sStr & "       Parametro.QuantidadeDatas, "
    sStr = sStr & "       Parametro.QuantidadeMinimaDias,"
    sStr = sStr & "       Parametro.DiretorioTransmissao, "
    sStr = sStr & "       Parametro.DiretorioRecepcao, "
    sStr = sStr & "       Parametro.Seq_Bordero, "
    sStr = sStr & "       Parametro.GerarArquivo_CEL, "
    sStr = sStr & "       Parametro.Num_Lote_CEL, "
    sStr = sStr & "       Parametro.Comp_Origem_CEL, "
    sStr = sStr & "       Parametro.Num_Versao_Inicial_CEL, "
    sStr = sStr & "       Parametro.Num_Versao_Final_CEL, "
    sStr = sStr & "       Parametro.HeaderAV, "
    sStr = sStr & "       Parametro.CriticaSoma, "
    sStr = sStr & "       Parametro.Codigo_USB, "
    sStr = sStr & "       Parametro.CPD_Origem, "
    sStr = sStr & "       Parametro.CPD_Destino, "
    sStr = sStr & "       Parametro.Codigo_Terceira, "
    sStr = sStr & "       Parametro.Nome_Terceira, "
    sStr = sStr & "       Parametro.CNPJ_Terceira, "
    sStr = sStr & "       Parametro.Cidade_Terceira, "
    sStr = sStr & "       Parametro.UF_Terceira, "
    sStr = sStr & "       Parametro.CodigoAplicacao, "
    sStr = sStr & "       Parametro.DiasLimpeza, "
    sStr = sStr & "       Parametro.Seq_Ocorrencia, "
    sStr = sStr & "       Parametro.CodigoAgAcolhed, "
    sStr = sStr & "       Parametro.Num_Remessa_TER, "
    sStr = sStr & "       Parametro.ValorChequeLimite, "
    sStr = sStr & "       Parametro.CodigoAplicacao, "
    sStr = sStr & "       Parametro.TMP_Pendente"
    sStr = sStr & "  FROM Parametro"

    If pFlag = 0 Then
        sStr = sStr & " WHERE DataProcessamento = " & pDataProcessamento
    Else
        sStr = sStr & " WHERE DataProcessamento = (Select Max(DataProcessamento) From Parametro)"
    End If
    
    GetParametros = sStr

End Function
Public Function GetDiasLimpeza() As String

Dim sStr As String

sStr = "SELECT TOP 1 Parametro.DiasLimpeza"
sStr = sStr & " FROM Parametro"
sStr = sStr & " ORDER BY Parametro.DataProcessamento DESC"

GetDiasLimpeza = sStr

End Function
Public Function GetGrupoUsuario(IdUsuario As Integer) As String

    Dim sStr  As String

            sStr = " Select GrpUser.IdGrupo,Grupo.Descricao "
     sStr = sStr & " From   GRUPOUSUARIO GrpUser, Grupo Grupo"
     sStr = sStr & " Where  IdUsuario = " & IdUsuario
     sStr = sStr & " And    GrpUser.IdGrupo = Grupo.IdGrupo"
     
     GetGrupoUsuario = sStr

End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                    Seleciona Datas de Depósito para o Bordero Atual                        '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetDatasBordero(ByVal pDataProcessamento As Long, _
                                ByVal pIdBordero As Long) As String

    Dim sStr As String

           sStr = "Select DataDeposito.DataProcessamento, "
    sStr = sStr & "       DataDeposito.IdBordero, "
    sStr = sStr & "       DataDeposito.DataDeposito, "
    sStr = sStr & "       DataDeposito.QuantidadeCheques, "
    sStr = sStr & "       DataDeposito.ValorDeposito"
    sStr = sStr & "  FROM DataDeposito "
    sStr = sStr & " WHERE DataProcessamento = " & pDataProcessamento
    sStr = sStr & "   AND IdBordero         = " & pIdBordero

    GetDatasBordero = sStr

End Function
Public Function GetDatas(ByVal pDataProcessamento As Long, _
                         ByVal pIdBordero As Long) As String

    Dim sStr As String

           sStr = "Select DataDeposito.DataDeposito"
    sStr = sStr & " FROM DataDeposito "
    sStr = sStr & " WHERE DataProcessamento = " & pDataProcessamento
    sStr = sStr & " AND IdBordero = " & pIdBordero
    'sStr = sStr & " Order by DataDeposito"
    
    GetDatas = sStr

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'              Seleciona dados do Cheque para o Cheque e Bordero Atual                      '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetDadosCheque(ByVal pDataProcessamento As Long, _
                               ByVal pIdBordero As Long, _
                               ByVal pIdcheque As Long) As String

    Dim sStr        As String

           sStr = "SELECT Cheque.DataDeposito, "
    sStr = sStr & "       Cheque.CMC7, "
    sStr = sStr & "       Cheque.CNPJCPF, "
    sStr = sStr & "       Cheque.Valor "
    sStr = sStr & "  FROM Cheque"
    sStr = sStr & " WHERE DataProcessamento = " & pDataProcessamento
    sStr = sStr & "   AND Idbordero = " & pIdBordero
    sStr = sStr & "   AND IdCheque = " & pIdcheque

    GetDadosCheque = sStr

End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                   Seleciona o maior Idcheque para Bordero Atual                             '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetMaxidCheque(ByVal pDataProcessamento As Long, _
                               ByVal pIdBordero As Long) As String

    Dim sStr        As String

           sStr = "SELECT Max(idCheque)As MaxIdCheque "
    sStr = sStr & "  FROM Cheque"
    sStr = sStr & " WHERE DataProcessamento = " & pDataProcessamento
    sStr = sStr & "   AND Idbordero = " & pIdBordero
  
    GetMaxidCheque = sStr

End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       Busca Dados para a Geração do arquivo CHINBO - Registro de Inclusão de Borderô       '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetDadosCHINBO(ByVal pDataProcessamento As Long, _
                               ByVal pStatus As String) As String

    Dim sStr        As String

           sStr = " SELECT A.CNPJ_Terceira"
    sStr = sStr & "      , B.idBordero"
    sStr = sStr & "      , B.Num_Bordero"
    sStr = sStr & "      , B.CodigoCarteira"
    sStr = sStr & "      , A.DataProcessamento"
    sStr = sStr & "      , B.CodigoLoja"
    sStr = sStr & "      , B.Agencia"
    sStr = sStr & "      , B.Conta"
    sStr = sStr & "      , 3 AS OrigemBordero"
    sStr = sStr & "      , 0 AS StatusBordero"
    sStr = sStr & "      , A.CodigoAgAcolhed"
    sStr = sStr & "      , B.SomaQuantidade"
    sStr = sStr & "      , B.SomaData"
    sStr = sStr & "      , B.SomaValor"
    sStr = sStr & "      , B.SomaTodos"
    sStr = sStr & "  FROM Parametro AS A, Bordero AS B"
    sStr = sStr & " WHERE A.DataProcessamento = " & pDataProcessamento
    sStr = sStr & "   AND A.DataProcessamento = B.DataProcessamento"
    sStr = sStr & "   AND B.Status = '" & pStatus & "'"

    GetDadosCHINBO = sStr

End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  Busca Dados para a Geração do arquivo CHINDT - Registro de Datas de Depósito do Borderô   '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetDadosCHINDT(ByVal pDataProcessamento As Long, _
                               ByVal pIdBordero As Long) As String

    Dim sStr        As String
   
'           sStr = " SELECT  DataDeposito"
'    sStr = sStr & "       , QuantidadeCheques"
'    sStr = sStr & "       , ValorDeposito"
'    sStr = sStr & "  FROM DataDeposito"
'    sStr = sStr & " WHERE DataProcessamento = " & pDataProcessamento
'    sStr = sStr & "   AND IdBordero = " & pIdBordero
    
    sStr = " SELECT A.DataDeposito"
    sStr = sStr & " , A.QuantidadeCheques"
    sStr = sStr & " , A.ValorDeposito"
    sStr = sStr & " , Count(B.IdCheque)"
    sStr = sStr & " FROM DataDeposito A Left Join Cheque B on "
    sStr = sStr & "(A.Dataprocessamento = B.DataProcessamento And"
    sStr = sStr & " A.IdBordero = B.IdBordero And"
    sStr = sStr & " A.DataDeposito = B.DataDeposito)"
    sStr = sStr & " WHERE A.DataProcessamento = " & pDataProcessamento
    sStr = sStr & " AND A.IdBordero = " & pIdBordero
    sStr = sStr & " AND B.Status IN ('1','C')"
    sStr = sStr & " Group By A.DataDeposito, A.QuantidadeCheques , A.ValorDeposito"
    
    GetDadosCHINDT = sStr

End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       Busca Dados para a Geração do arquivo CHINCH - Registro de Cheques de Borderô        '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetDadosCHINCH(ByVal pDataProcessamento As Long, _
                               ByVal pIdBordero As Long, _
                               ByVal pDataDeposito As Long) As String

    Dim sStr        As String
   
           sStr = " SELECT DataDeposito"
    sStr = sStr & "      , Valor"
    sStr = sStr & "      , CMC7"
    sStr = sStr & "      , TipoInscricao"
    sStr = sStr & "      , CNPJCPF"
    sStr = sStr & "  FROM Cheque"
    sStr = sStr & " WHERE DataProcessamento = " & pDataProcessamento
    sStr = sStr & "   AND IdBordero = " & pIdBordero
    sStr = sStr & "   AND DataDeposito = " & pDataDeposito
    sStr = sStr & "   AND Status <> 'D'"
    
    
    
    
    GetDadosCHINCH = sStr

End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                    Seleciona o Numero de Remessa                                '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetNumRemessaTER(ByVal pDataProcessamento As Long) As String
                              
    Dim sStr        As String

           sStr = "SELECT Num_Remessa_TER"
    sStr = sStr & "  FROM Parametro"
    sStr = sStr & " WHERE DataProcessamento = " & pDataProcessamento
    
    GetNumRemessaTER = sStr

End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                    Seleciona o Numero de Remessa                                '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetNumRemessaMov(ByVal pDataProcessamento As Long) As String
                              
    Dim sStr        As String

           sStr = "SELECT Num_Remessa_MOV"
    sStr = sStr & "  FROM Parametro"
    sStr = sStr & " WHERE DataProcessamento = " & pDataProcessamento
    
    GetNumRemessaMov = sStr

End Function


Public Function GetCheques(ByVal DataProcessamento As Long _
                , Optional ByVal IdBordero As Long _
                , Optional DataDeposito As Long) As String
'
'    Seleciona algumas informações de cheque
'

Dim sStr  As String
     
     sStr = "Select "
     sStr = sStr + IIf(IdBordero = 0, "IdBordero, ", "")
     sStr = sStr + IIf(DataDeposito = 0, "DataDeposito, ", "")
     
     sStr = sStr + "IdCheque, Cmc7, CnpjCpf, TipoInscricao, Status, Valor" & _
                    " from CHEQUE where DataProcessamento = " & DataProcessamento
     sStr = sStr + " and Status not in ('B','D')"
     
     If IdBordero <> 0 Then
          sStr = sStr & " and IdBordero = " & IdBordero
     End If
 
     If DataDeposito <> 0 Then
          sStr = sStr & " and DataDeposito = " & DataDeposito
     End If

GetCheques = sStr

End Function
Public Function GetChequesRejeitados(ByVal DataProcessamento As Long) As String

Dim sStr  As String
     
     sStr = "Select * from RejeicaoRemessa "
     sStr = sStr & "Where DataProcessamento = " & DataProcessamento
     'sStr = sStr + " And RotuloOriginal = 'CHINCH'"
     sStr = sStr + " And (RotuloOriginal = 'CHINBO' OR RotuloOriginal = 'CHINCH')"
          
     GetChequesRejeitados = sStr

End Function
Public Function GetDatasBorderosRejeitados(ByVal DataProcessamento As Long) As String

Dim sStr  As String
     
     sStr = "Select * from RejeicaoRemessa "
     sStr = sStr & "Where DataProcessamento = " & DataProcessamento
     'sStr = sStr + " And (RotuloOriginal = 'CHINCH' OR RotuloOriginal = 'CHINDT')"
     sStr = sStr + " And (RotuloOriginal = 'CHINBO' OR RotuloOriginal = 'CHINDT')"
     
     GetDatasBorderosRejeitados = sStr

End Function

Public Function GetBorderoConfirmacao(ByVal DataProcessamento As Long) As String

Dim sStr  As String

sStr = "Select * from BORDERO" & _
       " where DataProcessamento = " & DataProcessamento & _
       " and  status = 'T' "
     
GetBorderoConfirmacao = sStr

End Function
Public Function GetBorderoConfirmados(ByVal DataProcessamento As Long) As String

Dim sStr  As String

sStr = "Select * from BORDERO" & _
       " where DataProcessamento = " & DataProcessamento & _
       " and  status = 'E' "
     
GetBorderoConfirmados = sStr

End Function
Public Function GetSomatoriaCheques(ByVal DataProcessamento As Long _
                                  , ByVal IdBordero As Long _
                                  , ByVal DataDeposito As Long) As String
'
'    Retorno da quantidade e soma de valores dos cheque com status <> (B)Baixado / (D)Excluido
'

Dim sStr  As String
     
     sStr = "Select count(*), SUM(Valor) from CHEQUE where DataProcessamento = " & DataProcessamento & _
               " and IdBordero = " & IdBordero & _
               " and DataDeposito = " & DataDeposito & _
               " and Status not in ('B','D')"
               
GetSomatoriaCheques = sStr

End Function
Public Function GetNumBordero(ByVal DataProcessamento As Long _
                         , Optional IdBordero As Long _
                   , Optional ByVal Status As String) As String
'
'    Obtem informações específicas do borderô
'

Dim sStr  As String

     sStr = "Select IdBordero, Num_Bordero, SomaQuantidade, SomaValor from BORDERO " & _
               "where DataProcessamento = " & DataProcessamento

     If IdBordero <> 0 Then
          sStr = sStr & " and IdBordero = " & IdBordero
     End If
 
     If Status <> "" Then
               sStr = sStr & " and Status = '" & Status & "'"
     End If


GetNumBordero = sStr

End Function

Public Function GetDataDeposito(ByVal DataProcessamento As Long _
                     , Optional ByVal IdBordero As Long _
                     , Optional DataDeposito As Long) As String
'    Seleciona informações específicas de Data Depósito
'
'

Dim sStr  As String
     
     sStr = "Select "
     sStr = sStr + IIf(IdBordero = 0, "IdBordero, ", "")
     sStr = sStr + IIf(DataDeposito = 0, "DataDeposito, ", "")
     
     sStr = sStr + "QuantidadeCheques, ValorDeposito " & _
                    "from DATADEPOSITO where DataProcessamento = " & DataProcessamento

     If IdBordero <> 0 Then
          sStr = sStr & " and IdBordero = " & IdBordero
     End If
 
     If DataDeposito <> 0 Then
          sStr = sStr & " and DataDeposito = " & DataDeposito
     End If
     
' Retirado por Fernando (28/8) à pedidos do usuário
'If DataDeposito = 0 Then sStr = sStr & " order by DataDeposito"
     
GetDataDeposito = sStr

End Function
Public Function GetBorderoProvaZeroBloqueado(ByVal DataProcessamento As Long _
                                           , ByVal HoraLimite As String, Optional bSomenteContagem As Boolean = False) As String

'    Obtem informações específicas do borderô bloqueado em Prova Zero Status (G) e
'    borderô para Prova Zero (4), se somente verificação (optional), seleciona 'UmRegistro'
'    somente se existe borderô para prova zero
'
'    Obs.: Utiliza a tabela StatusBordero apenas para obter a horaatual do servidor
'
Dim sStr  As String

     If bSomenteContagem Then
          sStr = "Select 'UmRegistro' from BORDERO" & _
                    " where DataProcessamento = " & DataProcessamento & _
                    " and  (status = '4' " & _
                    " or (Status = 'G' and horaatual < " & _
                    " (select distinct time() - #" & HoraLimite & "# as hora from STATUSBORDERO)))"
     
     Else
          sStr = "Select IdBordero, Num_Bordero, SomaQuantidade, SomaValor, Status from BORDERO" & _
                    " where DataProcessamento = " & DataProcessamento & _
                    " and  (status = '4' " & _
                    " or (Status = 'G' and horaatual < " & _
                    " (select distinct time() - #" & HoraLimite & "# as hora from STATUSBORDERO)))"
     End If
     
GetBorderoProvaZeroBloqueado = sStr

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  Busca Dados para a Consulta de Cheques: A) Por Data De Depósito,                          '
'                                          B) Por Status                                     '
'                                          C) Por Banco / Ag / Conta                         '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetConsultaCheque(ByVal pDataProcessamento As Long _
                       , Optional ByVal pDataDeposito As Long _
                       , Optional ByVal pStatus As String _
                       , Optional ByVal pBanco As String _
                       , Optional ByVal pAgencia As String _
                       , Optional ByVal pContaCorrente As String _
                       , Optional ByVal pvalor As String) As String

    Dim sStr As String

               sStr = " SELECT Mid(CMC7,1,3)  AS Banco"
        sStr = sStr & "       ,Mid(CMC7,4,4)  AS Agencia"
        sStr = sStr & "       ,Mid(CMC7,20,10) AS Contac"
        sStr = sStr & "       ,Mid(CMC7,12,6) AS NrCheque"
        sStr = sStr & "       ,Cheque.DataDeposito"
        sStr = sStr & "       ,Cheque.CNPJCPF"
        sStr = sStr & "       ,Cheque.Valor"
        sStr = sStr & "       ,Cheque.IdBordero"
        sStr = sStr & "       ,Status.Status"
        sStr = sStr & "       ,Status.Descricao"
        sStr = sStr & "  FROM  Cheque Cheque, StatusCheque Status"
        sStr = sStr & "  WHERE Cheque.DataProcessamento = " & pDataProcessamento
        sStr = sStr & "    AND Status.Status = Cheque.Status"
        
        '''''''''''''''''''''''''''''''''
        ' Pesquisa por Data de Depósito '
        '''''''''''''''''''''''''''''''''
        If pDataDeposito <> 0 Then
            sStr = sStr & "    And DataDeposito      = " & pDataDeposito
        End If
        
        '''''''''''''''''''''''''''''''''
        ' Pesquisa por Status do Cheque '
        '''''''''''''''''''''''''''''''''
        If Len(pStatus) <> 0 Then
            sStr = sStr & "    And Descricao = '" & pStatus & "'"
        End If
    
        '''''''''''''''''''''''''''''''''
        ' Pesquisa por Banco do Cheque  '
        '''''''''''''''''''''''''''''''''
        If Len(pBanco) <> 0 Then
            sStr = sStr & "    And Mid(CMC7,1,3)  = '" & Trim(pBanco) & "'"
        End If
        
        '''''''''''''''''''''''''''''''''
        ' Pesquisa p/ Agência do Cheque '
        '''''''''''''''''''''''''''''''''
        If Len(pAgencia) <> 0 Then
            sStr = sStr & "    And Mid(CMC7,4,4)  ='" & (pAgencia) & "'"
        End If
        
        '''''''''''''''''''''''''''''''''
        ' Pesquisa por Conta do Cheque  '
        '''''''''''''''''''''''''''''''''
        If Len(pContaCorrente) <> 0 Then
            sStr = sStr & "    And Mid(CMC7,20,10) = '" & pContaCorrente & "'"
        End If
        
        '''''''''''''''''''''''
        ' Pesquisa por Valor  '
        '''''''''''''''''''''''
        If pvalor <> 0 Then
            sStr = sStr & "    And Valor = " & pvalor
        End If
                
        GetConsultaCheque = sStr

End Function
Public Function GetBorderoTransmissao(ByVal DataProcessamento As Long) As String

Dim sStr  As String

sStr = "Select * from BORDERO" & _
       " where DataProcessamento = " & DataProcessamento & _
       " and  status = 'R' "
     
GetBorderoTransmissao = sStr

End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                * Consulta de Borderôs: Agência e Conta do Favorecido *                      '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetConsultaBordero(ByVal pDataProcessamento As Long _
                        , Optional ByVal pAgencia As Integer _
                        , Optional ByVal pConta As Long _
                        , Optional ByVal pIdBordero As Long _
                        , Optional ByVal pStatus As String) As String
                                 
       Dim sStr As String

           sStr = " Select  Bordero.IdBordero"
    sStr = sStr & "       , Bordero.Num_Bordero"
    sStr = sStr & "       , Bordero.Agencia"
    sStr = sStr & "       , Bordero.Conta"
    sStr = sStr & "       , Bordero.NomeCliente"
    sStr = sStr & "       , Bordero.CodigoCarteira"
    sStr = sStr & "       , Carteira.Descricao"
    sStr = sStr & "       , Bordero.CodigoLoja"
    sStr = sStr & "       , Bordero.DataEntrada"
    sStr = sStr & "       , Status.Descricao As Status"
    sStr = sStr & "  From   Bordero Bordero, Carteira Carteira, StatusBordero Status"
    sStr = sStr & " Where   Bordero.CodigoCarteira    = Carteira.CodigoCarteira"
    sStr = sStr & "   And   Bordero.DataProcessamento = " & pDataProcessamento
    sStr = sStr & "   And   Bordero.Status            = Status.Status"
        
    '''''''''''''''''''''''''''''''''''''
    ' Pesquisa p/ Agência do Favorecido '
    '''''''''''''''''''''''''''''''''''''
    If pAgencia <> 0 Then
        sStr = sStr & "   And   Bordero.Agencia       = " & pAgencia
    End If
    
    '''''''''''''''''''''''''''''''''''''
    '  Pesquisa p/ Conta do Favorecido  '
    '''''''''''''''''''''''''''''''''''''
    If pConta <> 0 Then
        sStr = sStr & "   And   Bordero.Conta         = " & pConta
    End If

    '''''''''''''''''''''''''''''''''''''''''''
    '  Pesquisa por Identificação do Bordero  '
    '''''''''''''''''''''''''''''''''''''''''''
    If pIdBordero <> 0 Then
        sStr = sStr & "   And   Bordero.IdBordero     = " & pIdBordero
    End If

    '''''''''''''''''''''''''''''''''
    ' Pesquisa por Status do Cheque '
    '''''''''''''''''''''''''''''''''
    If Len(pStatus) <> 0 Then
        sStr = sStr & "    And Status.Descricao = '" & pStatus & "'"
    End If

    GetConsultaBordero = sStr

End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                 * Consulta de Cheques do Borderôs: Numero Borderô  *                        '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetConsultaChequesBordero(ByVal pDataProcessamento As Long _
                               , Optional ByVal pNumBordero As String) As String
                        

       Dim sStr As String
        
           sStr = " SELECT Mid(CMC7,1,3)  AS Banco"
    sStr = sStr & "       ,Mid(CMC7,4,4)  AS Agencia"
    sStr = sStr & "       ,Mid(CMC7,23,7) AS Contac"
    sStr = sStr & "       ,Mid(CMC7,12,6) AS NrCheque"
    sStr = sStr & "       ,Cheque.DataDeposito"
    sStr = sStr & "       ,Cheque.CNPJCPF"
    sStr = sStr & "       ,Cheque.Valor"
    sStr = sStr & "       ,Bordero.IdBordero"
    sStr = sStr & "       ,Status.Status"
    sStr = sStr & "       ,Status.Descricao"
    sStr = sStr & "  FROM  Cheque Cheque, Bordero Bordero, StatusCheque Status"
    sStr = sStr & " Where  Cheque.DataProcessamento = " & pDataProcessamento
    sStr = sStr & "   AND  Cheque.DataProcessamento = Bordero.DataProcessamento"
    sStr = sStr & "   AND  Cheque.IdBordero         = Bordero.IdBordero"
    sStr = sStr & "   AND  Bordero.Num_Bordero      = '" & pNumBordero & "'"
    sStr = sStr & "   AND  Status.Status            = Cheque.Status"
    
    GetConsultaChequesBordero = sStr
  
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                     * Verifica se CMC7 já esta cadastrado na Base *                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetPesquisaCMC7(ByVal pDataProcessamento As Long _
                              , ByVal pCMC7 As String _
                     , Optional ByVal pIdcheque As Long) As String

    Dim sStr As String

           sStr = " Select CMC7 From Cheque "
    sStr = sStr & " Where DataProcessamento = " & pDataProcessamento
    sStr = sStr & "   AND CMC7              = '" & pCMC7 & "'"
    sStr = sStr & "   AND STATUS NOT IN ('D','B')"
    
    If pIdcheque <> 0 Then
        sStr = sStr & "   AND idCheque     <> " & pIdcheque
    End If
    
    GetPesquisaCMC7 = sStr

End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    Pesquisa cheque por Data Deposito e CMC7 retornando dados de Bordero  '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetChequeBorderoPorCMC7(ByVal lngDataDeposito As Long, ByVal strCMC7 As String) As String

    Dim sStr As String

           sStr = " Select Ch.DataProcessamento, Ch.IdBordero, Ch.Status as StatusCheque,"
    sStr = sStr + " Ch.IdCheque, Bo.Status as StatusBordero From CHEQUE Ch"
    sStr = sStr & " left join  BORDERO Bo on Bo.DataProcessamento = Ch.DataProcessamento "
    sStr = sStr & " and   Bo.IdBordero    = Ch.IdBordero"
    sStr = sStr & " where Ch.DataDeposito = " & lngDataDeposito
    sStr = sStr & " and   Ch.CMC7         = '" & strCMC7 & "'"
    
    GetChequeBorderoPorCMC7 = sStr
    
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Pesquisa a existencia do cheque com código do erro   '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetExisteRejeicaoRemessa(ByVal pDataProcessamento As Long, _
                                        ByVal pIdBordero As Long, _
                                        ByVal pDataDeposito As Long, _
                                        ByVal pIdcheque As Long, _
                                        ByVal pCodErro As Integer) As String

    Dim sStr        As String
    
           sStr = "SELECT count(0)"
    sStr = sStr & " FROM  RejeicaoRemessa "
    sStr = sStr & " WHERE DataProcessamento = " & pDataProcessamento
    sStr = sStr & " AND   IdBordero    = " & pIdBordero
    sStr = sStr & " AND   DataDeposito = " & pDataDeposito
    sStr = sStr & " AND   IdCheque     = " & pIdcheque
    sStr = sStr & " AND   CodigoErro   = " & pCodErro
    
    GetExisteRejeicaoRemessa = sStr
End Function
Public Function GetChequeBorderoPorLinha1(ByVal lngDataDeposito As Long _
                                        , ByVal strCompensacao As String _
                                        , ByVal strBanco As String _
                                        , ByVal strAgencia As String _
                                        , ByVal strContaCorrente As String _
                                        , ByVal strNumCheque As String) As String
Dim sStr As String

            sStr = " Select Ch.DataProcessamento, Ch.IdBordero, Ch.Status as StatusCheque,"
     sStr = sStr & " Ch.IdCheque, Bo.Status as StatusBordero From CHEQUE Ch"
     sStr = sStr & " left join  BORDERO Bo "
     sStr = sStr & " on    Bo.DataProcessamento = Ch.DataProcessamento "
     sStr = sStr & " and   Bo.IdBordero         = Ch.IdBordero"
     sStr = sStr & " where Ch.DataDeposito      = " & lngDataDeposito
     sStr = sStr & " and   Mid(Ch.CMC7,9,3)     = '" & strCompensacao & "'"
     sStr = sStr & " and   Mid(Ch.CMC7,1,3)     = '" & Right(strBanco, 3) & "'"
     sStr = sStr & " and   Mid(Ch.CMC7,4,4)     = '" & strAgencia & "'"
     sStr = sStr & " and   Mid(Ch.CMC7,23,7)    = '" & Right(strContaCorrente, 7) & "'"
     sStr = sStr & " and   Mid(Ch.CMC7,12,6)    = '" & Right(strNumCheque, 6) & "'"
     
GetChequeBorderoPorLinha1 = sStr
    
End Function
Public Function GetChequeFusao(pCMC7 As String, pDataProcessamento As Long) As String

    Dim sStr As String

           sStr = ""
           sStr = "SELECT Num_Bordero, CodigoCarteira, Valor, Fusao"
           sStr = sStr & " From ChequeDataBoa"
           sStr = sStr & " Where Dataprocessamento = " & pDataProcessamento
           sStr = sStr & " And CMC7 = '" & pCMC7 & "'"
    
    GetChequeFusao = sStr

End Function
Public Function GetChequePendenteFusao(pDataProcessamento As Long) As String

    Dim sStr As String

           sStr = ""
           sStr = "SELECT Num_Bordero, CodigoCarteira, Valor, Fusao"
           sStr = sStr & " From ChequeDataBoa"
           sStr = sStr & " Where Dataprocessamento = " & pDataProcessamento
           sStr = sStr & " And Fusao = False"
    
    GetChequePendenteFusao = sStr

End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                      * Consulta de Instruções enviadas pelo VC  *                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetConsultaInstrucaoVC(ByVal pDataProcessamento As Long _
                            , Optional ByVal pNumBordero As String _
                            , Optional ByVal pDteAnterior As Long _
                            , Optional ByVal pDteNova As Long _
                            , Optional ByVal pBanco As String _
                            , Optional ByVal pAgencia As String _
                            , Optional ByVal pContac As String _
                            , Optional ByVal pNrCheque As String _
                            , Optional ByVal pvalor As String _
                            ) As String
                        
    Dim sStr As String
        
           sStr = " SELECT AlteracaoData.Num_Bordero"
    sStr = sStr & "       ,AlteracaoData.CodigoCarteira "
    sStr = sStr & "       ,AlteracaoData.DataAnterior "
    sStr = sStr & "       ,AlteracaoData.DataNova "
    sStr = sStr & "       ,AlteracaoData.BancoEmitente "
    sStr = sStr & "       ,AlteracaoData.AgenciaEmitente"
    sStr = sStr & "       ,AlteracaoData.CcEmitente "
    sStr = sStr & "       ,AlteracaoData.NrChequeEmitente "
    sStr = sStr & "       ,AlteracaoData.NossoNumero"
    sStr = sStr & "       ,AlteracaoData.CodigoCompensacao"
    sStr = sStr & "       ,AlteracaoData.Valor"
    sStr = sStr & "  FROM  AlteracaoData AlteracaoData"
    sStr = sStr & " Where  AlteracaoData.DataProcessamento = " & pDataProcessamento
    
    '''''''''''''''''''''''''''''''''''''''''''
    '* Pesquisa por Identificação do Bordero *'
    '''''''''''''''''''''''''''''''''''''''''''
    If Len(Trim(pNumBordero)) <> 0 Then
        sStr = sStr & "   And   AlteracaoData.Num_Bordero = '" & pNumBordero & "'"
    End If
   
    '''''''''''''''''''''''''''''''''''''''''''
    '     * Pesquisa por Data Anterior *      '
    '''''''''''''''''''''''''''''''''''''''''''
    If pDteAnterior <> 0 Then
        sStr = sStr & "   And   AlteracaoData.DataAnterior = " & pDteAnterior
    End If

    '''''''''''''''''''''''''''''''''''''''''''
    '        * Pesquisa por Data Nova *       '
    '''''''''''''''''''''''''''''''''''''''''''
    If pDteNova <> 0 Then
        sStr = sStr & "   And   AlteracaoData.DataNova = " & pDteNova
    End If

    '''''''''''''''''''''''''''''''''''''''''''
    '          * Pesquisa por Banco *         '
    '''''''''''''''''''''''''''''''''''''''''''
    If Len(Trim(pBanco)) <> 0 Then
        sStr = sStr & "   And   AlteracaoData.BancoEmitente = " & pBanco
    End If

    '''''''''''''''''''''''''''''''''''''''''''
    '        * Pesquisa por Agencia *         '
    '''''''''''''''''''''''''''''''''''''''''''
    If Len(Trim(pAgencia)) <> 0 Then
        sStr = sStr & "   And   AlteracaoData.AgenciaEmitente = " & pAgencia
    End If

    '''''''''''''''''''''''''''''''''''''''''''
    '          * Pesquisa por Conta *         '
    '''''''''''''''''''''''''''''''''''''''''''
    If Len(Trim(pContac)) <> 0 Then
        sStr = sStr & "   And   AlteracaoData.CcEmitente = " & pContac
    End If

    '''''''''''''''''''''''''''''''''''''''''''
    '    * Pesquisa por Número do Cheque *    '
    '''''''''''''''''''''''''''''''''''''''''''
    If Len(Trim(pNrCheque)) <> 0 Then
        sStr = sStr & "   And   AlteracaoData.NrChequeEmitente = " & pNrCheque
    End If

    '''''''''''''''''''''''''''''''''''''''''''
    '         * Pesquisa por Valor *          '
    '''''''''''''''''''''''''''''''''''''''''''
    If pvalor <> 0 Then
        sStr = sStr & "   And   AlteracaoData.Valor =" & pvalor
    End If

    GetConsultaInstrucaoVC = sStr
    
End Function
Public Function GetChequesComplementados(ByVal pDataProcessamento As Long, _
                                         ByVal pIdBordero As Long, _
                                Optional ByVal pDataDeposito As Long, _
                                Optional ByVal pIdcheque As Long) As String

    Dim sStr        As String
    
           sStr = " Select" & vbCrLf
    sStr = sStr & " Mid(CMC7, 1, 3) as Banco," & vbCrLf
    sStr = sStr & " Mid(CMC7, 4, 4) as Agencia," & vbCrLf
    sStr = sStr & " Mid(CMC7, 22, 7) as [Conta Corrente]," & vbCrLf
    sStr = sStr & " Mid(CMC7, 29, 1) as [Digito]," & vbCrLf
    sStr = sStr & " Mid(CMC7, 12, 6) as Cheque," & vbCrLf
    sStr = sStr & " CCur(Valor) as [Vr Depósito]," & vbCrLf
    sStr = sStr & " IdCheque " & vbCrLf
    sStr = sStr & " FROM Cheque"
    sStr = sStr & " WHERE DataProcessamento = " & pDataProcessamento
    sStr = sStr & " AND IdBordero = " & pIdBordero
    
    If pDataDeposito <> 0 Then
        sStr = sStr & " AND DataDeposito = " & pDataDeposito
    End If
    
    If pIdcheque <> 0 Then
        sStr = sStr & " AND IdCheque = " & pIdcheque
    End If
    
    GetChequesComplementados = sStr

End Function
Public Function GetDetalhesArquivo_TerGCC(ByVal DataTroca As Long, Optional ByVal iNumremessa As Integer)
    
    Dim sStr        As String

           sStr = "SELECT ChequeDataBoa.Agencia "
    sStr = sStr & "  ,ChequeDataBoa.CMC7"
    sStr = sStr & "  ,ChequeDataBoa.Valor"
    sStr = sStr & "  ,ChequeDataBoa.Agencia"
    sStr = sStr & "  ,ChequeDataBoa.Conta"
    sStr = sStr & "  ,ChequeDataBoa.Num_Versao_CEL"
    sStr = sStr & "  ,ChequeDataBoa.Num_Lote_CEL"
    sStr = sStr & "  ,ChequeDataBoa.Geracao_CEL"
    sStr = sStr & "  ,ChequeDataBoa.Fusao"
    sStr = sStr & "  FROM ChequeDataBoa"
    sStr = sStr & "  WHERE Fusao = True"
    ' sStr = sStr & "   AND Geracao_CEL = True"
    sStr = sStr & "  AND ChequeDataBoa.DataProcessamento = " & DataTroca
    If iNumremessa <> 0 Then
        sStr = sStr & "  AND Num_Remessa_TER = " & iNumremessa
    Else
        sStr = sStr & "  AND Geracao_TER = False"
    End If
    sStr = sStr & " ORDER BY CMC7"
    
    GetDetalhesArquivo_TerGCC = sStr

End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                   Obtem Número do IdBordero                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetNumeroIdBordero(ByVal DataProcessamento As Long, strNumBordero As String) As String

    Dim sStr  As String

    sStr = "Select IdBordero from BORDERO "
    sStr = sStr & "where DataProcessamento = " & DataProcessamento
    sStr = sStr & " and Num_Bordero = '" & strNumBordero & "'"
    sStr = sStr & " and status = 'T'"

    GetNumeroIdBordero = sStr

End Function
Public Function GetDataHoraServidor() As String
'''''''''''''''''''''''''''''''''''''''''''''''''
'    Obtem Data e hora do servidor      '
'''''''''''''''''''''''''''''''''''''''''''''''''
    Dim sStr  As String

    sStr = "select distinct Date() as data, Time() as hora from STATUSBORDERO"
     
    GetDataHoraServidor = sStr

End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                    Seleciona o Numero de Remessa                                '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetNumVersaoCEL(ByVal pDataProcessamento As Long) As String
                              
    Dim sStr        As String

           sStr = "SELECT Num_Versao_CEL"
    sStr = sStr & "  FROM Parametro"
    sStr = sStr & " WHERE DataProcessamento = " & pDataProcessamento
    
    GetNumVersaoCEL = sStr

End Function
Public Function GetAvisoDif(ByVal pDataOcorrencia As Long, _
                            ByVal pCodigoOcorrencia As Long, ByVal sCodigoOcorrencia As Integer) As String

    Dim sStr        As String
    
           sStr = "SELECT "
    sStr = sStr & "       AvisoDiferenca.DataOcorrencia, "
    sStr = sStr & "       AvisoDiferenca.CodigoOcorrencia "
    sStr = sStr & "  FROM AvisoDiferenca"
    sStr = sStr & " WHERE DataOcorrencia = " & pDataOcorrencia
    sStr = sStr & "   AND CodigoOcorrencia = " & pCodigoOcorrencia
    sStr = sStr & "   AND CodigoMotivo = '" & sCodigoOcorrencia & "'"
    
    GetAvisoDif = sStr

End Function

Public Function GetRegraGP(ByVal pDataProcessamento As Long, _
                            ByVal pCodigoProduto As String, ByVal pCodigoRegra As String) As String

    Dim sStr        As String
    
           sStr = "SELECT "
    sStr = sStr & "       RegraGP.DataProcessamento, "
    sStr = sStr & "       RegraGP.CodigoProduto, "
    sStr = sStr & "       RegraGP.CodigoRegra "
    sStr = sStr & "  FROM RegraGP"
    sStr = sStr & " WHERE DataProcessamento = " & pDataProcessamento
    sStr = sStr & "   AND CodigoProduto = " & pCodigoProduto
    sStr = sStr & "   AND CodigoRegra = " & pCodigoRegra
    
    GetRegraGP = sStr

End Function


Public Function GetCamposExportacao(ByVal sReferencia As String) As String
     '***** Referencia (B)orderô (I)nstrução   *****
     Dim sStr As String
     
     sStr = "SELECT * FROM CamposExportacao where Referencia = '" & sReferencia & "' order by ordenacao"
     
     GetCamposExportacao = sStr

End Function

Public Function GetComplementoExportacao(ByVal sReferencia As String) As String
     '***** Referencia (B)orderô (I)nstrução   *****
     Dim sStr As String

     sStr = "SELECT * FROM ComplementoExportacao where Referencia = '" & sReferencia & "'"

     GetComplementoExportacao = sStr

End Function

Public Function GetStructChequesBaixados() As String
        
     Dim sStr As String
     
     sStr = " SELECT * FROM  ChequesBaixados "
     sStr = sStr & "Where Num_Bordero = '-1'"          'Obtem somente a estrutura da tabela
    
     GetStructChequesBaixados = sStr

End Function
Public Function GetStructAlteracaoData() As String

     Dim sStr As String
     
     sStr = " SELECT * FROM  AlteracaoData "
     sStr = sStr & "Where DataProcessamento = 19000101"          'Obtem somente a estrutura da tabela
    
     GetStructAlteracaoData = sStr

End Function
Public Function GetEstruturaCheque(ByVal DataProcessamento As Long) As String

Dim sStr  As String
     
     sStr = "Select * from Cheque "
     sStr = sStr & "Where DataProcessamento = " & DataProcessamento
          
     GetEstruturaCheque = sStr
     
End Function
