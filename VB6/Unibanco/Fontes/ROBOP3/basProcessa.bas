Attribute VB_Name = "basProcessa"
 Option Explicit
Sub LogFGTS()

    Dim TipoIdentificacao As String * 1
    
   'variaveis do header
    Geral.CodTransacao = "0320"
    Geral.Evento = 970
    Geral.TipoTransacao = 2
    Geral.Capa = GetCapa(Geral.idEnvMal)
    Geral.IndTransac = " "
    Geral.ValorTrans = formataValor(Geral.rstDoctos!Valor)
    Geral.Deposito = formataValor(Geral.rstDocto!Deposito)
    Geral.JAM = formataValor(Geral.rstDocto!JAM)
    Geral.Multa = formataValor(Geral.rstDocto!Multa)
    
    Geral.hsSQLa = "Exec recfgts "
      
   'monta header
    MontaHeader
     
    MontaComplemento
     
   'monta parte variavel
    Geral.hsSQLa = Geral.hsSQLa & ", '" & Geral.rstDocto!CodigoBarras & "'"
    
    If Len(Trim(Geral.rstDocto!CNPJCEI_Tomador)) = 0 Or IsNull(Geral.rstDocto!CNPJCEI_Tomador) Then
        TipoIdentificacao = "0"
    ElseIf Len(Geral.rstDocto!CNPJCEI_Tomador) = 12 Then
        TipoIdentificacao = "2"
    Else
        TipoIdentificacao = "1"
    End If

    Geral.hsSQLa = Geral.hsSQLa & ", " & TipoIdentificacao
    Geral.hsSQLa = Geral.hsSQLa & ", '" & IIf(Trim(Geral.rstDocto!CNPJCEI_Tomador) = Empty, String(14, "0"), Format(Geral.rstDocto!CNPJCEI_Tomador, String(14, "0"))) & "'"
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.Deposito)
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.JAM)
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.Multa)
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorTrans)
    Geral.hsSQLa = Geral.hsSQLa & ", " & "1"
    
    MontaComplementoVariavel
    
    LocalLog "FGTS - SP " & Geral.hsSQLa

End Sub
Sub LogFGTS1()

    Dim TipoIdentificacao As String * 1
    
   'variaveis do header
    Geral.CodTransacao = "0420"
    Geral.Evento = 970
    Geral.TipoTransacao = 2
    Geral.Capa = GetCapa(Geral.idEnvMal)
    Geral.IndTransac = " "
    Geral.ValorTrans = formataValor(Geral.rstDoctos!Valor)
    Geral.Deposito = formataValor(Geral.rstDocto!Deposito)
    Geral.JAM = formataValor(Geral.rstDocto!JAM)
    Geral.Multa = formataValor(Geral.rstDocto!Multa)
    
    Geral.hsSQLa = "Exec recfgts1 "
      
   'monta header
    MontaHeader
     
    MontaComplemento
     
   'monta parte variavel
    Geral.hsSQLa = Geral.hsSQLa & ", '" & Geral.rstDocto!CodigoBarras & "'"
    Geral.hsSQLa = Geral.hsSQLa & ", '" & IIf(Trim(Geral.rstDocto!CNPJCEI_Tomador) = Empty, String(16, "0"), Format(Geral.rstDocto!CNPJCEI_Tomador, String(16, "0"))) & "'"
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorTrans)
    Geral.hsSQLa = Geral.hsSQLa & ", " & "1"
    
    MontaComplementoVariavel
    
    LocalLog "FGTS - SP " & Geral.hsSQLa

End Sub
Sub LogAjusteDebitoADCC()
'================================================'
' TRANSAÇÃO 38(nosso número) - debito automatico '
'================================================'

On Error GoTo TrataErro
   
    Dim RstUBB      As Recordset
    Dim spRetorno   As Integer
       
    Geral.Capa = GetCapa(Geral.idEnvMal)
   
   'variaveis do header
    Geral.CodTransacao = "0015"
    Geral.Evento = 580
    Geral.TipoTransacao = 1
    Geral.IndTransac = "O"
    Geral.Capa = GetCapa(Geral.idEnvMal)
    Geral.TipoConta = "C"
    Geral.AgenciaVinculo = Geral.rstDocto!Agencia
    Geral.ContaVinculo = Geral.rstDocto!Conta
    
   'consulta depara
    DePara
    
    If Geral.PreparouLog = 1 Then
        Exit Sub
    End If

    Geral.ValorTrans = formataValor(Geral.rstDoctos!Valor)

    Geral.hsSQLa = "exec avintpar "
   
    MontaHeader
      
   'monta parte variavel
    Geral.hsSQLa = Geral.hsSQLa & ", 2"
    Geral.hsSQLa = Geral.hsSQLa & ", 1565"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.AgenCob
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.ContaCob
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorTrans)
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Val(Parametros.DataServer)
    Geral.hsSQLa = Geral.hsSQLa & ", ' '"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 1"
    Geral.hsSQLa = Geral.hsSQLa & ", ' '"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    
    If GetDocumentoTransmitido(EnumOutros) Then
        Exit Sub
    End If

   'gravar o nsu desta transação antes de envia-la para o UBB-NT
    spRetorno = MDIQuery.updNsuDocto(Geral.DataProcessamento, _
                                     Geral.rstCapa!idcapa, _
                                     Geral.rstDoctos!iddocto, _
                                     Caixa.NSU1, _
                                     Caixa.Caixa)
                    
    If spRetorno <> 0 Then
        MsgBox "ATENÇÃO! Falha na SP [ updNsuDocto ]. ", vbOKOnly + vbCritical, "Atenção"
        Exit Sub
    End If
       
   'enviar a stored procedure
    Set Geral.rst = UBBQuery.ExecuteSQL(Geral.hsSQLa)
    LocalLog "Retorno Ajuste aut. debito: " & Format(RstUBB(0), "000")
   
    If (Val(Geral.rst(0)) <> 0) Then
        Geral.CodOcorrencia = 999
        DevolveDocumentos
    Else
        
       'atualizar o adcc como enviado
        spRetorno = MDIQuery.updDoctoTransmitido(Geral.DataProcessamento, _
                                                 Geral.rstCapa!idcapa, _
                                                 Geral.rstDoctos!iddocto, _
                                                 Caixa.NSU1, _
                                                 Caixa.Caixa, "N")
                                                    
        If spRetorno <> 0 Then
            MsgBox "ATENÇÃO! Falha na SP [ updDoctoTransmitido ]. ", vbOKOnly + vbCritical, "Atenção"
        End If
        
    End If
    
    Exit Sub

TrataErro:
    
    Select Case TratamentoErro("Falha no AjusteDebitoADCC.", Err, eDoctoSubidaLog)
        Case eSair
            End
        Case eRepetir
            Resume
        Case eContinuar
            Resume Next
        Case eFinalizar
            Exit Sub
    End Select

End Sub
Sub LogLanctoInterno()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'verifica se cheque está pagando alguma cobrança Ubb vencida e se esta pode ser paga '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo TrataErro

'    Dim RstMDI  As Recordset
'
'   'variaveis do header
'    Set RstMDI = MDIQuery.getContraPartida(Geral.rstDocto!Evento)
'
'    If RstMDI.EOF Then
'        RstMDI.Close
'        MsgBox "ATENÇÃO !!! (31)Não foi possível localizar a contra-partida para este evento. Tecle <enter> para continuar.", vbOKOnly + vbCritical, "Atenção"
'        End
'    End If

    Geral.Vinculo = Geral.rstDoctos!Vinculo
    Geral.CodTransacao = "BHN3"
'   Geral.Evento = RstMDI!ContraPartida
    Geral.Evento = 561
    Geral.TipoTransacao = 1
    Geral.Capa = GetCapa(Geral.idEnvMal)
    Geral.IndTransac = "O"
    Geral.ValorTrans = formataValor(Geral.rstDoctos!Valor)
        
    Geral.hsSQLa = "exec lannfoff "
    
    MontaHeader
    
   'monta parte variavel
    Geral.hsSQLa = Geral.hsSQLa & ", " & "''"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Right(Trim(Geral.rstDocto!ControleBanco), 14)
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.ValorTrans
    Geral.hsSQLa = Geral.hsSQLa & ", " & "0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Mid(Geral.DataProcessamento, 7, 2) & Mid(Geral.DataProcessamento, 5, 2) & Mid(Geral.DataProcessamento, 1, 4))
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Mid(Geral.rstDocto!DataGeracao, 7, 2) & Mid(Geral.rstDocto!DataGeracao, 5, 2) & Mid(Geral.rstDocto!DataGeracao, 1, 4))
    
    LocalLog "Lancamento Interno - SP " & Geral.hsSQLa
    
    Exit Sub
    
TrataErro:
    
    Screen.MousePointer = 0
    Select Case TratamentoErro("Falha no módulo de Lançto.", Err)
        Case eSair
            End
        Case eRepetir
            Resume
        Case eContinuar
            Resume Next
    End Select
   
End Sub
Sub LogContabilDebito()
      
Dim sSqla As String
 
   'variaveis do header
    Geral.CodTransacao = "3033"
    Geral.Evento = 244
    Geral.TipoTransacao = 1
    Geral.Capa = GetCapa(Geral.idEnvMal)
    Geral.IndTransac = " "
    Geral.ValorTrans = formataValor(Geral.rstDoctos!Valor)
        
    Geral.hsSQLa = "exec lancinte "
      
   'monta header
    MontaHeader
     
   'monta parte variavel
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & "''"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Format(Caixa.NSU1, "000000")
    Geral.hsSQLa = Geral.hsSQLa & Format(Parametros.TipoAgencia, "0")
    Geral.hsSQLa = Geral.hsSQLa & Format(Geral.rstCapa!agorig, "0000")
    Geral.hsSQLa = Geral.hsSQLa & Format(Caixa.Caixa, "000")
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.ValorTrans
    Geral.hsSQLa = Geral.hsSQLa & ", " & "0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & "1"
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Mid(Geral.DataProcessamento, 7, 2) & Mid(Geral.DataProcessamento, 5, 2) & Mid(Geral.DataProcessamento, 1, 4))
    Geral.hsSQLa = Geral.hsSQLa & ", " & "'01232'"
    Geral.hsSQLa = Geral.hsSQLa & ", " & "08199900098"
    Geral.hsSQLa = Geral.hsSQLa & ", " & "66680"
    
    LocalLog "Ajuste Contabil de Débito - SP " & Geral.hsSQLa

End Sub
Sub LogContabilCredito()
      
    Dim sSqla As String
 
   'variaveis do header
    Geral.CodTransacao = "3033"
    Geral.Evento = 263
    Geral.TipoTransacao = 2
    Geral.Capa = GetCapa(Geral.idEnvMal)
    Geral.IndTransac = " "
    Geral.ValorTrans = formataValor(Geral.rstDoctos!Valor)
        
    Geral.hsSQLa = "exec lancinte "
      
   'monta header
    MontaHeader
     
   'monta parte variavel
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & "''"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Format(Caixa.NSU1, "000000")
    Geral.hsSQLa = Geral.hsSQLa & Format(Parametros.TipoAgencia, "0")
    Geral.hsSQLa = Geral.hsSQLa & Format(Geral.rstCapa!agorig, "0000")
    Geral.hsSQLa = Geral.hsSQLa & Format(Caixa.Caixa, "000")
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.ValorTrans
    Geral.hsSQLa = Geral.hsSQLa & ", " & "0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & "1"
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Mid(Geral.DataProcessamento, 7, 2) & Mid(Geral.DataProcessamento, 5, 2) & Mid(Geral.DataProcessamento, 1, 4))
    Geral.hsSQLa = Geral.hsSQLa & ", " & "'00907'"
    Geral.hsSQLa = Geral.hsSQLa & ", " & "7399900010"
    Geral.hsSQLa = Geral.hsSQLa & ", " & "66320"
       
    LocalLog "Ajuste Contabil de Crédito - SP " & Geral.hsSQLa

End Sub
Sub AgenciaColeta()
On Error GoTo TrataErro
 
    Dim sSqla As String
        
   'leitura do número do terminal
    CalculaNSU (True)
    Caixa.NSU1 = Caixa.NSU
    
    sSqla = "exec cxalagcx "
    sSqla = sSqla & "  '" & "XXXX" & "'"
    sSqla = sSqla & ", 0"
    sSqla = sSqla & ", " & Caixa.VersaoAtual
    sSqla = sSqla & ", " & Val(Parametros.AgenciaCentral)
    sSqla = sSqla & ", " & Parametros.AgenciaSatelite
    sSqla = sSqla & ", 0"
    sSqla = sSqla & ", " & 0
    sSqla = sSqla & ", 3"
    sSqla = sSqla & ", " & Caixa.Caixa
    sSqla = sSqla & ", 1"
    sSqla = sSqla & ", " & Caixa.NSU1
    sSqla = sSqla & ", 0"
    sSqla = sSqla & ", " & 0
    sSqla = sSqla & ", " & Geral.Hora
    sSqla = sSqla & ", '" & " " & "'"
    
    If Geral.idEnvMal = "E" Then
       sSqla = sSqla & ", 6"
    Else
       sSqla = sSqla & ", 7"
    End If
    
    sSqla = sSqla & ", " & Geral.TpRep
    sSqla = sSqla & ", " & 0
    sSqla = sSqla & ", '" & GetCapa(Geral.idEnvMal) & "'"
    sSqla = sSqla & ", " & Geral.rstCapa!agorig
      
    Set Geral.rst = UBBQuery.ExecuteSQL(sSqla)
            
    LocalLog "ret troca agencia coleta - " & Format(Geral.rst(0), "0000")
       
    Exit Sub

TrataErro:
    Screen.MousePointer = 0
    Select Case TratamentoErro("Falha no módulo: [Agencia de Coleta] ", Err)
        Case eSair
           End
        Case eRepetir
           Resume
        Case eContinuar
           Resume Next
    End Select
 
End Sub
Function EstornoDeposito(ByVal pcodTran As Integer) As Boolean

On Error GoTo TrataErro
    
    Dim spRetorno                   As Integer
    Dim RespostaConestor            As Integer
    Dim QtdeCheques                 As Integer
    Dim ValorCheque                 As String
    Dim ValorDinheiro               As String
    Dim ValorSomado                 As String
    Dim SDV                         As String * 1

    ValorCheque = 0
    ValorSomado = 0
    ValorDinheiro = 0
    
    RespostaConestor = Conestor()
    
    If RespostaConestor = 1 Then
      'Apenas para verificar(logar) se retorno indevido da conestor - Estorno ja realizado
       LocalLog "Retorno Proc.Conestor - Estorno ja realizado. - CAPA :" & Format(Geral.rstCapa!Capa, "00000000000")
       EstornoDeposito = True
       Exit Function
    ElseIf RespostaConestor = -1 Then
      'Apenas para verificar(logar) se retorno indevido da conestor - Estorno não pode ser efetuado
       LocalLog "Retorno Proc.Conestor - Estorno não pode ser efetuado. - CAPA :" & Format(Geral.rstCapa!Capa, "00000000000")
       Exit Function
    End If
    
    LocalLog "Inicio do Estorno de Deposito/OCT" & " - Retorno Conestor: " & Trim(RespostaConestor)
    
    Call GetChequeCashDeposito(Geral.rstCapa!idcapa, Geral.rstCapa!Vinculo, Geral.rstCapa!Valor, QtdeCheques, ValorDinheiro, ValorCheque, ValorSomado, False)

   'variaveis do header
    Geral.idEnvMal = Geral.rstCapa!idEnv_Mal
    Geral.CodTransacao = pcodTran
    Geral.Evento = 0
    Geral.TipoTransacao = 0
    Geral.Capa = GetCapa(Geral.idEnvMal)
    Geral.IndTransac = " "
    Geral.CapaBack = 0
    Geral.TpRep = 0

    Geral.ValorTrans = formataValor(Geral.rstCapa!Valor)
    
   'SAQUE
    Geral.hsSQLa = "exec estodepo "
   
    MontaHeader
      
   'monta parte variavel
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.rstCapa!NSU
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    
    If Abs(ValorCheque - ValorSomado) = 0 Then
        Geral.ValorTrans = ValorSomado
    Else
        Geral.ValorTrans = formataValor(Abs(ValorCheque - ValorSomado), True)
    End If

    'Geral.ValorTrans = formataValor(Geral.rstCapa!Valor)
    CalculaNSU
    Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.NSU
        
    CalculaNSU
    Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.NSU
    SDV = Caixa.SDV
    
    Geral.ValorTrans = ValorSomado
    CalculaNSU
    Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.NSU
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    
    Geral.hsSQLa = Geral.hsSQLa & ", '" & Caixa.SDV & "'"
    Geral.hsSQLa = Geral.hsSQLa & ", '" & SDV & "'"
    
    LocalLog "Procedure Estorno(Deposito): " & Trim(Geral.rstCapa!Nome) & ": " & Geral.hsSQLa
    Set Geral.rst = UBBQuery.ExecuteSQL(Geral.hsSQLa)
    LocalLog "ret do estorno CPO " & Trim(Geral.rstCapa!Nome) & ": " & Format(Geral.rst(0).Value, "0000")
        
    If Geral.rst(0) = 0 Then
    
        EstornoDeposito = True
        Geral.RetTransacao = 86
        Geral.CodOcorrencia = "999"
            
        Call DbRejeitaDocto(Geral.rstCapa!idcapa, Geral.rstCapa!iddocto)
       
    Else
             
        LocalLog "Não foi possivel realizar estorno deposito ret: " & Geral.rst(0)
        MsgBox "Não foi possível realizar o estorno para o docto [" & Trim(Geral.rstCapa!Nome) & "] no valor de [R$ " & Format(Geral.rstCapa!Valor, ".00") & "]. Verifique. Tecle <Enter> para continuar.", vbOKOnly + vbCritical, "Atenção"
 
       
    End If
            
    Exit Function

TrataErro:
    Screen.MousePointer = 0
    If Conestor() = 1 Then
       LocalLog "Falha no Estorno, Erro : " & Err & ", Conestor Retorno = Estorno efetuado - Continuando..."
       Resume Next
    Else
       Select Case TratamentoErro("Falha ao efetivar estorno.", Err, True)
          Case eSair
              LocalLog "Falha no Estorno: " & Err & ", Conestor Retorno = Estorno não efetuado - Finalizado pelo usuario."
              End
          Case eContinuar
              LocalLog "Falha no Estorno: " & Err & ", Conestor Retorno = Estorno não efetuado - Repetindo Operação."
              Resume
       End Select
    End If

End Function
Function EstornoGeral(ByVal pcodTran As Integer) As Boolean

On Error GoTo TrataErro

    Dim spRetorno           As Integer
    Dim RespostaConestor    As Integer
    
    RespostaConestor = Conestor()
    
    If RespostaConestor = 1 Then
      'Apenas para verificar(logar) se retorno indevido da conestor - Estorno ja realizado
       LocalLog "Retorno Proc.Conestor - Estorno ja realizado. - CAPA: " & Format(Geral.rstCapa!Capa, "00000000000")
       EstornoGeral = True
       Exit Function
    ElseIf RespostaConestor = -1 Then
      'Apenas para verificar(logar) se retorno indevido da conestor - Estorno não pode ser efetuado
       LocalLog "Retorno Proc.Conestor - Estorno não pode ser efetuado. - CAPA: " & Format(Geral.rstCapa!Capa, "00000000000")
       Exit Function
    End If
    
    LocalLog "Inicio da Rotina de Estorno Generica" & " - Retorno conestor: " & Trim(RespostaConestor)

   'variaveis do header
    Geral.idEnvMal = Geral.rstCapa!idEnv_Mal
    Geral.CodTransacao = pcodTran
    Geral.Evento = 0
    Geral.TipoTransacao = 0
    Geral.Capa = GetCapa(Geral.idEnvMal)
    Geral.IndTransac = " "
    Geral.CapaBack = 1
    Geral.TpRep = 0
    
    Geral.ValorTrans = formataValor(Geral.rstCapa!Valor)
   
    Geral.hsSQLa = "exec proestor "

    MontaHeader
      
   'monta parte variavel
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.rstCapa!NSU    'nsu do documento
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"                       'nsu laranja - adap.para SQL
    Geral.hsSQLa = Geral.hsSQLa & ", 0"                       'nsu laranja 2
    Geral.hsSQLa = Geral.hsSQLa & ", 0"                       'nsu laranja 3
    Geral.hsSQLa = Geral.hsSQLa & ", 0"                       'nsu laranja 4
    Geral.hsSQLa = Geral.hsSQLa & ", 0"                       'SDV2
    
    LocalLog "Procedure Estorno(Geral): " & Trim(Geral.rstCapa!Nome) & " : " & Geral.hsSQLa
    Set Geral.rst = UBBQuery.ExecuteSQL(Geral.hsSQLa)
    LocalLog "ret do estorno TOR para " & Trim(Geral.rstCapa!Nome) & " : " & Format(Geral.rst(0), "0000")
    
    If Geral.rst(0) = 0 Then
        EstornoGeral = True
        Geral.RetTransacao = 86
        Geral.CodOcorrencia = "999"
            
        Call DbRejeitaDocto(Geral.rstCapa!idcapa, Geral.rstCapa!iddocto)
       
    Else
             
        MsgBox "Não foi possível realizar o estorno para o docto [" & Trim(Geral.rstCapa!Nome) & "] no valor de [R$ " & Format(Geral.rstCapa!Valor, ".00") & "]. Verifique. Tecle <Enter> para continuar.", vbOKOnly + vbCritical, "Atenção"
       
    End If
    
    Exit Function
 
TrataErro:
    Screen.MousePointer = 0
    If Conestor() = 1 Then
       LocalLog "Falha no Estorno, Erro: " & Err & ", Conestor Retorno = Estorno efetuado - Continuando..."
       Resume Next
    Else
       Select Case TratamentoErro("Falha ao efetivar estorno.", Err, True)
          Case eSair
              LocalLog "Falha no Estorno: " & Err & ", Conestor Retorno = Estorno não efetuado - Finalizado pelo usuario."
              End
          Case eContinuar
              LocalLog "Falha no Estorno: " & Err & ", Conestor Retorno = Estorno não efetuado - Repetindo Operação."
              Resume
       End Select
    End If
End Function
Function EstornoArrecad() As Boolean

On Error GoTo TrataErro

    Dim spRetorno           As Integer
    Dim RespostaConestor    As Integer
    Dim Vez                 As Integer
   
    RespostaConestor = Conestor()
    
    If RespostaConestor = 1 Then
      'Apenas para verificar(logar) se retorno indevido da conestor - Estorno ja realizado
       LocalLog "Retorno Proc.Conestor - Estorno já realizado. - CAPA: " & Format(Geral.rstCapa!Capa, "00000000000")
       EstornoArrecad = True
       Exit Function
    ElseIf RespostaConestor = -1 Then
      'Apenas para verificar(logar) se retorno indevido da conestor - Estorno não pode ser efetuado
       LocalLog "Retorno Proc.Conestor - Estorno não pode ser efetuado. - CAPA: " & Format(Geral.rstCapa!Capa, "00000000000")
       Exit Function
    End If
    
    LocalLog "Inicio do Estorno de Arrecadações Convencionais" & " - Resposta Conestor: " & Trim(RespostaConestor)
    
   'variaveis para o header
    Geral.idEnvMal = Geral.rstCapa!idEnv_Mal
    Geral.CodTransacao = "0032"
    Geral.Evento = 39
    Geral.TipoTransacao = 0
    
    Geral.ValorTrans = formataValor(Geral.rstCapa!Valor)
    
    Geral.Capa = GetCapa(Geral.idEnvMal)
    Geral.IndTransac = " "
    Geral.CapaBack = 1
    Geral.TpRep = 0
   
   'stored procedure da fracassada
    Geral.hsSQLa = "exec proestrc "
    
    MontaHeader
    
   'monta parte variavel
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.rstCapa!NSU
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", ''"    'SDV2
    Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.NSU1
    
    LocalLog "Procedure Estorno(Arrecadacao): " & Trim(Geral.rstCapa!Nome) & " /SP:" & Geral.hsSQLa
    Set Geral.rst = UBBQuery.ExecuteSQL(Geral.hsSQLa)
    LocalLog "Ret do proestrc (0015) para " & Trim(Geral.rstCapa!Nome) & " : " & Format(Geral.rst(0), "0000")
        
    If Geral.rst(0) = 0 Then
        EstornoArrecad = True
        Geral.RetTransacao = 86
        Geral.CodOcorrencia = "999"
            
        Call DbRejeitaDocto(Geral.rstCapa!idcapa, Geral.rstCapa!iddocto)
       
    Else
             
        MsgBox "Não foi possível realizar o estorno para o docto [" & Trim(Geral.rstCapa!Nome) & "] no valor de [R$ " & Format(Geral.rstCapa!Valor, ".00") & "]. Verifique. Tecle <Enter> para continuar.", vbOKOnly + vbCritical, "Atenção"
       
    End If
    
    Exit Function
 
TrataErro:
    Screen.MousePointer = 0
    If Conestor() = 1 Then
       LocalLog "Falha no Estorno, Erro: " & Err & ", Conestor Retorno = Estorno efetuado - Continuando..."
       Resume Next
    Else
       Select Case TratamentoErro("Falha ao efetivar estorno.", Err, True)
          Case eSair
              LocalLog "Falha no Estorno: " & Err & ", Conestor Retorno = Estorno não efetuado - Finalizado pelo usuario."
              End
          Case eContinuar
              LocalLog "Falha no Estorno: " & Err & ", Conestor Retorno = Estorno não efetuado - Repetindo Operação."
              Resume
       End Select
    End If
End Function

Function EstornoTitulo() As Boolean

On Error GoTo TrataErro

    Dim spRetorno           As Integer
    Dim RetQX               As Integer
    Dim Funcao              As String * 14
    Dim MsgIda              As String
    Dim MsgRetorno          As String
    Dim HeaderTx            As String
    Dim TamIda              As String
    Dim NsuSyb              As String
    Dim TrnEstorno          As String
    Dim RespostaConestor    As Integer
    Dim Vez                 As Integer
   
    RespostaConestor = Conestor(NsuSyb, TrnEstorno)
    
    If RespostaConestor = 1 Then
      'Apenas para verificar(logar) se retorno indevido da conestor - Estorno ja realizado
       LocalLog "Retorno Proc.Conestor - Estorno já realizado. - CAPA: " & Format(Geral.rstCapa!Capa, "00000000000")
       EstornoTitulo = True
       Exit Function
    ElseIf RespostaConestor = -1 Then
      'Apenas para verificar(logar) se retorno indevido da conestor - Estorno não pode ser efetuado
       LocalLog "Retorno Proc.Conestor - Estorno não pode ser efetuado. - CAPA: " & Format(Geral.rstCapa!Capa, "00000000000")
       Exit Function
    End If
    
    LocalLog "Inicio do Estorno de Titulos de Outros Bancos" & " - Resposta Conestor: " & Trim(RespostaConestor)
    
   'variaveis para o header
    Geral.idEnvMal = Geral.rstCapa!idEnv_Mal
    Geral.CodTransacao = "0F32"
    Geral.Evento = 291
    Geral.TipoTransacao = 0
    Geral.Capa = GetCapa(Geral.idEnvMal)
    Geral.IndTransac = "1"
    Geral.CapaBack = 0
    Geral.TpRep = 0
    Geral.ValorTrans = formataValor(Geral.rstCapa!Valor)
   
   'stored procedure da fracassada
    Geral.hsSQLa = "exec estootbc "
    
    MontaHeader
    
   'monta parte variavel
    Geral.hsSQLa = Geral.hsSQLa & ", " & NsuSyb
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.NSU1
    Geral.hsSQLa = Geral.hsSQLa & ", ''"     'SDV2
    Geral.hsSQLa = Geral.hsSQLa & ", 0"      'NSU_ADV
    Geral.hsSQLa = Geral.hsSQLa & ", ''"     'SDV2
    
    LocalLog "Procedure Estorno(Titulo): " & Trim(Geral.rstCapa!Nome) & " /SP:" & Geral.hsSQLa
    Set Geral.rst = UBBQuery.ExecuteSQL(Geral.hsSQLa)
    LocalLog "Ret do estorno TBC (0F32) para " & Trim(Geral.rstCapa!Nome) & " : " & Format(Geral.rst(0), "0000")
    
    If (Val(Geral.rst(0)) = 0) Then
        
        CalculaNSU
   
        HeaderTx = "BHNF" & "000000" & Caixa.VersaoAtual & _
                    Format(Parametros.AgenciaCentral, "0000") & _
                    Format(Parametros.AgenciaSatelite, "0000") & "000" & _
                    Format(Caixa.Caixa, "000") & "1" & "000000" & _
                    "000000" & "0" & Format(Now, "HHMM") & "1" & "3" & _
                    "0000000000"
               
        MsgIda = HeaderTx & Format(NsuSyb, "000000") & Format(TrnEstorno, "0000")
             
       'Envia 1ª mensagem ao Host
        TamIda = Format(Len(Trim(MsgIda)), "0000")
        MsgRetorno = String(1921, " ")
        Funcao = "1" & TamIda & "1921****"
        
        LocalLog MsgIda
                
       'Envia BHS1
        Call Abrelinha("BHNF")
        RetQX = qxhostnt(Funcao, MsgIda, MsgRetorno)
        Call FechaLinha("BHNF")
                   
       '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       'se retorno = 20,41,71 -> o micro deverá ser reinicializado '
       '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       
        If (RetQX = 20) Or (RetQX = 41) Or (RetQX = 71) Then
             LocalLog "Retorno BHNF: " & Mid(MsgRetorno, 58, 2)
             MsgBox "Atenção. Ocorreu um erro de comunicação com o Servidor da Agência. Reinicialize este equipamento. Retorno DLL = " & Format(RetQX, "00"), vbOKOnly + vbCritical, "Atenção"
             Close #20
             End
        End If
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Retorno = 21,30,43,47,52,62,80 -> tentar novamente.'
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        If (RetQX = 21) Or (RetQX = 30) Or (RetQX = 43) Or (RetQX = 47) Or _
           (RetQX = 52) Or (RetQX = 62) Or (RetQX = 80) Then
        
            Vez = 1
            Do
               
                Espera (5 * Vez)
                     
               'tentar novamente
                Call Abrelinha("RE-envio BHNF")
                RetQX = qxhostnt(Funcao, MsgIda, MsgRetorno)
                Call FechaLinha("RE-envio BHNF")
                
                Vez = Vez + 1
            
            Loop Until Not ((RetQX = 21) Or (RetQX = 30) Or (RetQX = 43) Or _
                            (RetQX = 47) Or (RetQX = 52) Or (RetQX = 62) Or (RetQX = 80)) And Vez <> 4
        
        End If
        
        LocalLog "Retorno BHNF: " & Format(RetQX, "00") & "MsgRetorno: " & MsgRetorno

       'Recebe retorno da Consulta do Host BHNF
        If (RetQX = 0) Then
            LocalLog MsgRetorno
      
           'Recebe resposta no BHFR/ Grava concretizada 0032
            If Mid(MsgRetorno, 58, 2) = "00" Or Mid(MsgRetorno, 58, 2) = "03" Then 'retorno OK
                          
                Geral.CodTransacao = "0032"
                Geral.IndTransac = " "
                Geral.TipoTransacao = 1
              
               'stored procedure da concretizada
                Geral.hsSQLa = "Exec estootbc "
            
               'monta header
                MontaHeader
            
               'monta parte variavel
                Geral.hsSQLa = Geral.hsSQLa & ", " & NsuSyb
                Geral.hsSQLa = Geral.hsSQLa & ", 0"
                Geral.hsSQLa = Geral.hsSQLa & ", " & Mid(MsgRetorno, 61, 6)
                Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.NSU1
                Geral.hsSQLa = Geral.hsSQLa & ", ''"     'SDV2
                Geral.hsSQLa = Geral.hsSQLa & ", 0"      'NSU_ADV
                Geral.hsSQLa = Geral.hsSQLa & ", ''"     'SDV2
                                                 
               ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
               ' Concretizada - 2ª parte (gravando no server da agencia)  '
               ' confirma a perna fracassada 0F32 que subiu anteriormente '
               ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                
                Set Geral.rst = UBBQuery.ExecuteSQL(Geral.hsSQLa)
                LocalLog "RetQx do estorno TBC para " & Trim(Geral.rstCapa!Nome) & ": " & Format(Geral.rst(0), "0000")
                            
                If (Val(Geral.rst(0)) = 0) Then
                                
                    MsgIda = ""
                    CalculaNSU
                    
                    HeaderTx = "BHS3" & Format(Geral.rst(13), "000000") & _
                    Caixa.VersaoAtual & _
                    Format(Parametros.AgenciaCentral, "0000") & _
                    Format(Parametros.AgenciaSatelite, "0000") & "011" & _
                    Format(Caixa.Caixa, "000") & "1" & _
                    Format(Caixa.NSU, "000000") & "000000" & "0" _
                    & Format(Now, "HHMM") & "110000000002"
                                   
                    MsgIda = HeaderTx & "3" & Mid(MsgRetorno, 61, 6) & "000000" & "1" & Format(Parametros.AgenciaSatelite, "0000")
                    TamIda = Format(Len(Trim(MsgIda)), "0000")
                    MsgRetorno = String(1921, " ")
               
                    Funcao = "1" & TamIda & "1921****"
                                          
                    LocalLog MsgIda
                    
                   'envia BHS3 (confirmação para o Host)
                    Call Abrelinha("BHS3")
                    RetQX = qxhostnt(Funcao, MsgIda, MsgRetorno)
                    Call FechaLinha("BHS3")
                    
                    If RetQX <> 0 Then
                        Err.Raise 963, App.Title, "Falha na Confirmação de Pagamento (BHS3)"
                    End If
                    
                    EstornoTitulo = True
                    Geral.CodOcorrencia = "999"
                    Geral.RetTransacao = 86
               
                    Call DbRejeitaDocto(Geral.rstCapa!idcapa, Geral.rstCapa!iddocto)
                       
                   'libera linha
                    Geral.GereiLog = 1
                          
               End If
                              
            ElseIf Mid(MsgRetorno, 58, 2) = "02" Then
            
                Screen.MousePointer = 0
                MsgBox "Título já extraído sem possibilidade de estorno. Contate a USB de sua região para regularização. Tecle <enter> para continuar.", vbOKOnly + vbCritical, "Atenção"
                
            Else
            
                LocalLog "Retorno " & Mid(MsgRetorno, 58, 2)
                MsgBox "Estorno não Autorizado pelo sistema de Caixa, docto não pode ser estornado.", vbOKOnly + vbCritical, "Atenção"
                
            End If
         
        Else
        
            MsgBox "Retorno Indevido do sistema de Caixa docto não pode ser estornado.", vbOKOnly + vbCritical, "Atenção"
            LocalLog "Retorno da função da DLL para o envio da BHNF-Retorno: " + Str(RetQX)
            
        End If
      
    Else
       
        If (Val(Geral.rst(0)) = 7) Then
            MsgBox "Atenção! Não foi possível finalizar o estorno, pois o caixa está FECHADO. Saia do sistema e entre novamente para executar a rotina de Abertura de Caixa.", vbOKOnly + vbCritical, "Atenção"
        End If
     
    End If

    Exit Function
    
TrataErro:

    If Err.Number = 964 Or Err.Number = 965 Then
       'Erro na Abertura/Fechamento de Linha
        Geral.CodOcorrencia = 999
        Geral.RetTransacao = 98
       'Call DevolveDocumentos
       'Exit Function
    ElseIf Err.Number = 963 Then
       'Erro Subida da confirmaçao (BHS3)
        Geral.CodOcorrencia = 999
        Geral.RetTransacao = 97
       'Call DevolveDocumentos
       'Exit Sub
    End If

   Screen.MousePointer = 0
   If Conestor() = 1 Then
      LocalLog "Falha no Estorno, Erro : " & Err & ", Conestor Retorno = Estorno efetuado... Continuando"
      Resume Next
   Else
      Select Case TratamentoErro("Falha ao efetivar estorno.", Err, True)
         Case eSair
             LocalLog "Falha no Estorno: " & Err & ", Conestor Retorno = Estorno não efetuado... Finalizado pelo usuario."
             End
         Case eContinuar
             LocalLog "Falha no Estorno: " & Err & ", Estorno não efetuado... Repetindo Operação."
             Resume
      End Select
   End If

End Function
Function Conestor(Optional NsuSyb As String, Optional TrnEstorno As String) As Integer
On Error GoTo TrataErro

    Dim RstMDI      As Recordset
    Dim RstUBB      As Recordset
    Dim spRetorno   As Integer

   'variaveis do header
    Geral.idEnvMal = Geral.rstCapa!idEnv_Mal
    Geral.CodTransacao = "0032"
    Geral.Evento = 0
    Geral.TipoTransacao = 0
    Geral.Capa = GetCapa(Geral.idEnvMal)
    Geral.IndTransac = " "
    Geral.CapaBack = 0
    Geral.TpRep = 0
    
    LocalLog "Inicio da Consulta para Estorno"
      
   'stored procedure do saque
    Geral.hsSQLa = "exec conestor "
            
    MontaHeader
      
   'monta parte variavel
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.rstCapa!NSU
    
    Set RstMDI = MDIQuery.getTipoDocto(Geral.rstCapa!TipoDocto)
    LocalLog "Consulta de estorno p/ " & Trim(RstMDI!Nome) & ": " & Geral.hsSQLa
    
    Set RstUBB = UBBQuery.ExecuteSQL(Geral.hsSQLa)
    
    LocalLog "ret da consulta de estorno p/ " & Trim(RstMDI!Nome) & ": " & Format(RstUBB(0), "0000")
    
    If RstUBB(0) = 0 Then               'Pode ser estornado
    
       Conestor = 0
       
       TrnEstorno = RstUBB(15)
       NsuSyb = RstUBB(16)
       
    ElseIf RstUBB(0) = 83 Then          'Ja foi efetivado estorno
       Conestor = 1
    Else                                'Docto não pode ser estornado
    
       Conestor = -1
       
       spRetorno = MDIQuery.insMensagem(Geral.DataProcessamento, _
                                         MSG_EstornoNaoAutorizadoBH, _
                                         Geral.rstCapa!idcapa, _
                                         Geral.rstCapa!iddocto, _
                                         Caixa.Caixa)
                                         
       If spRetorno <> 0 Then MsgBox "Falha Procedure [ InsMensagem ]", vbCritical + vbOKOnly
    
    End If
  
    Call GaugePos(Estorno, Geral.rstCapa!Nome)
    
    Exit Function
 
TrataErro:
    Screen.MousePointer = 0
    
    Select Case TratamentoErro("Falha na atualização do estorno.", Err)
      Case eSair
          End
      Case eRepetir
          Resume
      Case eContinuar
          Resume Next
    End Select

End Function
Function Depara230() As Boolean
    
On Error GoTo TrataErro

    Dim Funcao          As String * 14
    Dim MsgIda          As String
    Dim MsgRetorno      As String
    Dim RetQX           As Integer
    Dim HeaderTx        As String
    Dim TamIda          As String
    Dim Vez             As Integer
    
    Depara230 = True
    
   '''''''''''''''''
   ' Enviar a BHNC '
   '''''''''''''''''
    
    Geral.AgenCob = Geral.AgenciaVinculo
    Geral.ContaCob = Geral.ContaVinculo
 
    HeaderTx = "BHNC" & "000000" & Caixa.VersaoAtual & _
                Format(Parametros.AgenciaCentral, "0000") & _
                Format(Parametros.AgenciaSatelite, "0000") & "000" & _
                Format(Caixa.Caixa, "000") & "1" & "000000" & _
                "000000" & "0" & Format(Now, "HHMM") & "1" & "3" & _
                "0000000000"
    
    MsgIda = HeaderTx & "00" & "1" & _
             Format(Geral.AgenCob, "0000") & _
             Format(Geral.ContaCob, "0000000000") & _
             "0" & "0000" & _
             Format(Geral.IdentDep, "000000000") & _
             "0230" & String(37, " ")

         
   'Envia 1ª mensagem ao Host
    TamIda = Format(Len(Trim(MsgIda)), "0000")
    MsgRetorno = String(1921, " ")
    Funcao = "1" & TamIda & "1921****"
      
    LocalLog MsgIda
             
   'Envia BHNC
    Call Abrelinha("BHNC")
    RetQX = qxhostnt(Funcao, MsgIda, MsgRetorno)
    Call FechaLinha("BHNC")
               
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   'se retorno = 20,41,71 -> o micro deverá ser reinicializado '
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    If (RetQX = 20) Or (RetQX = 41) Or (RetQX = 71) Then
         LocalLog "Retorno BHNC: " & Mid(MsgRetorno, 58, 2)
         MsgBox "Atenção. Ocorreu um erro de comunicação com o Servidor da Agência. Reinicialize este equipamento. Retorno DLL = " & Format(RetQX, "00"), vbOKOnly + vbCritical, "Atenção"
         Close #20
         End
    End If
    
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' Retorno = 21,30,43,47,52,62,80 -> tentar novamente.'
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    If (RetQX = 21) Or (RetQX = 30) Or (RetQX = 43) Or (RetQX = 47) Or _
       (RetQX = 52) Or (RetQX = 62) Or (RetQX = 80) Then
    
        Vez = 1
        Do
           
            Espera (5 * Vez)
                 
           'tentar novamente
            Call Abrelinha("RE-envio BHNC")
            RetQX = qxhostnt(Funcao, MsgIda, MsgRetorno)
            Call FechaLinha("RE-envio BHNC")
            
            Vez = Vez + 1
        
        Loop Until Not ((RetQX = 21) Or (RetQX = 30) Or (RetQX = 43) Or _
                        (RetQX = 47) Or (RetQX = 52) Or (RetQX = 62) Or (RetQX = 80)) And Vez <> 4
    
    End If
    
    LocalLog "Retorno BHNC: " & Format(RetQX, "00") & "MsgRetorno: " & MsgRetorno
         
   'Retorno da BHNC
    If (RetQX = 0) Then
             
       LocalLog MsgRetorno
         
       If Mid(MsgRetorno, 138, 1) = "5" Or Mid(MsgRetorno, 138, 1) = "6" Then
          Geral.AgenCob = Val(Mid(MsgRetorno, 94, 4))
          Geral.ContaCob = Val(Mid(MsgRetorno, 98, 10))
       Else
          LogChequeCompensacao   'altera para compensacao e sai do modulo
          AlteraParaCompensacao     'para tratar novamente
          Depara230 = False
       End If
        
    Else
    
      'para estes codigos de retorno, transformar o saque em compensação
       Select Case RetQX
          Case 21, 30, 43, 47, 52, 62, 80, 33, 36, 42, 48
               LogChequeCompensacao
               AlteraParaCompensacao
          Case Else
               Geral.CodOcorrencia = 999
               Geral.RetTransacao = 51
               
               Call DevolveDocumentos

       End Select
       
       Depara230 = False
    
    End If
   
    Exit Function
    
TrataErro:

    Screen.MousePointer = 0
    
    Select Case TratamentoErro("Falha no módulo: [Consulta Tabela De-Para 230", Err)
        Case eSair
            End
        Case eRepetir
            Resume
        Case eContinuar
            Resume Next

    End Select

End Function
Function AgenciaCadastradaBH() As Boolean

On Error GoTo TrataErro
   
    Dim spRetorno   As Integer
    Static Vezes
    
    LocalLog "Capa : " & Format(Geral.rstCapa!Capa, String(11, "0"))
    
Reinicio:
    Parametros.AgenciaSatelite = Geral.rstCapa!agorig
    Geral.Hora = Format(Now, "HHMM")
  
   'caso não exista o arquivo, então abre caixa
    If GetSetting("Robo", "Caixa", "Aberto", 0) = 0 Then
    
       'Leitura do número do terminal
        Call CalculaNSU(True)
        Caixa.NSU1 = Caixa.NSU
        
       '================================'
       ' ABERTURA / REABERTURA DE CAIXA '
       '================================'
        
        Geral.hsSQLb = "exec abrecx "
        
       'monta header
        Geral.hsSQLb = Geral.hsSQLb & " '0030'"
        Geral.hsSQLb = Geral.hsSQLb & ", 0"
        Geral.hsSQLb = Geral.hsSQLb & ", " & Caixa.VersaoAtual
        Geral.hsSQLb = Geral.hsSQLb & ", " & Parametros.AgenciaCentral
        Geral.hsSQLb = Geral.hsSQLb & ", " & Parametros.AgenciaSatelite
        Geral.hsSQLb = Geral.hsSQLb & ", 2"
        Geral.hsSQLb = Geral.hsSQLb & ", 4"
        Geral.hsSQLb = Geral.hsSQLb & ", 0"
        Geral.hsSQLb = Geral.hsSQLb & ", " & Caixa.Caixa
        Geral.hsSQLb = Geral.hsSQLb & ", 1"
        Geral.hsSQLb = Geral.hsSQLb & ", " & Caixa.NSU1
        Geral.hsSQLb = Geral.hsSQLb & ", 0"
        Geral.hsSQLb = Geral.hsSQLb & ", 1"
        Geral.hsSQLb = Geral.hsSQLb & ", " & Geral.Hora
        Geral.hsSQLb = Geral.hsSQLb & ", ' '"
        
        If Geral.idEnvMal = "E" Then
            Geral.hsSQLb = Geral.hsSQLb & ", 6"
        Else
            Geral.hsSQLb = Geral.hsSQLb & ", 7"
        End If
        
        Geral.hsSQLb = Geral.hsSQLb & ", 0"
        Geral.hsSQLb = Geral.hsSQLb & ", 126"
        Geral.hsSQLb = Geral.hsSQLb & ", ' ', '"
        
        Geral.hsSQLb = Geral.hsSQLb & Caixa.CIF
        Geral.hsSQLb = Geral.hsSQLb & "', '" & Caixa.SDV
        
       'monta parte variavel
        Geral.hsSQLb = Geral.hsSQLb & "', 0"
        Geral.hsSQLb = Geral.hsSQLb & ", ' '"
        Geral.hsSQLb = Geral.hsSQLb & ", 0"
        Geral.hsSQLb = Geral.hsSQLb & ", 0"
        Geral.hsSQLb = Geral.hsSQLb & ", 0"
         
        LocalLog "Abertura de Caixa - para verificação 143 - SP " & Geral.hsSQLb
        Set Geral.rst = UBBQuery.ExecuteSQL(Geral.hsSQLb)
        LocalLog "Retorno sp_Abertura de Caixa para verificação 143 - " & Format(Geral.rst(0), "00")
       
        If (Val(Geral.rst(0)) = 0) Then
           'Caixa foi aberto OK
            SaveSetting appname:="Robo", section:="Caixa", Key:="Aberto", setting:=1
            LogFechamentoCaixa ("A")
            AgenciaCadastradaBH = True
        Else
            SaveSetting appname:="Robo", section:="Caixa", Key:="Aberto", setting:=0
            
            MDIQuery.insLog Geral.DataProcessamento, Geral.rstCapa!idcapa, "0", Caixa.UsuarioAtual, "124"
            AgenciaCadastradaBH = False

            If (Val(Geral.rst(0)) = 143) Then
            
                LocalLog "Ocorreu o erro " & Str(Geral.rst(0)) & " na abertura do Caixa Agência de coleta: [" & Parametros.AgenciaSatelite & "]. não cadastrada."
                
                Call GaugeTitulo(2)
                
                spRetorno = MDIQuery.insMensagem(Geral.DataProcessamento, _
                                                 MSG_AgenCapanaoCadastrada, _
                                                 Geral.rstCapa!idcapa, 0, _
                                                 Caixa.Caixa)
            
               'Seta Capa para Ilegíveis
                spRetorno = MDIQuery.UpdCapaStatusCaixaControle(Geral.DataProcessamento, Geral.rstCapa!idcapa, ST_CapaParaIlegivel, Caixa.Caixa)
                If spRetorno <> 0 Then MsgBox "Falha Procedure [UpdCapaStatusCaixaControle]", vbCritical + vbOKOnly
          
               'Insere log
                MDIQuery.insLog Geral.DataProcessamento, Geral.rstCapa!idcapa, "0", Caixa.UsuarioAtual, "124"
                
            ElseIf Val(Geral.rst(0)) = 102 Then
                LocalLog "Ocorreu o erro " & Str(Geral.rst(0)) & " na abertura do Caixa Agência de coleta: [" & Parametros.AgenciaSatelite & "]. movimento não liberado."
                MsgBox "ATENÇÃO ! Erro " & Str(Geral.rst(0)) & " na abertura da agência de coleta [" & Parametros.AgenciaSatelite & "]. movimento não liberado. Verifique ! ", vbOKOnly + vbCritical, "Atenção"
                Close #20
                End
            ElseIf Val(Geral.rst(0)) = 148 Then
                LocalLog "Ocorreu o erro " & Str(Geral.rst(0)) & " na abertura do Caixa Agência de coleta: [" & Parametros.AgenciaSatelite & "]. Versão Incorreta do Caixa."
                MsgBox "ATENÇÃO ! Erro " & Str(Geral.rst(0)) & " na abertura da agência de coleta [" & Parametros.AgenciaSatelite & "]. Versão Incorreta do Caixa. Verifique ! ", vbOKOnly + vbCritical, "Atenção"
                Close #20
                End
            ElseIf Val(Geral.rst(0)) = 144 Then
                LocalLog "Ocorreu o erro " & Str(Geral.rst(0)) & " na abertura do Caixa Agência de coleta: [" & Parametros.AgenciaSatelite & "]. não permitido Agência Centralizadora."
                
                Call GaugeTitulo(2)
                
                spRetorno = MDIQuery.insMensagem(Geral.DataProcessamento, _
                                                 MSG_NaoPermitidaAgCentral, _
                                                 Geral.rstCapa!idcapa, 0, _
                                                 Caixa.Caixa)
            
               'Seta Capa para Ilegíveis
                spRetorno = MDIQuery.UpdCapaStatusCaixaControle(Geral.DataProcessamento, Geral.rstCapa!idcapa, ST_CapaParaIlegivel, Caixa.Caixa)
                If spRetorno <> 0 Then MsgBox "Falha Procedure [UpdCapaStatusCaixaControle]", vbCritical + vbOKOnly
          
               'Insere log
                MDIQuery.insLog Geral.DataProcessamento, Geral.rstCapa!idcapa, "0", Caixa.UsuarioAtual, "124"
                        
            Else
                LocalLog "Ocorreu o erro " & Str(Geral.rst(0)) & " na abertura do Caixa Agência de coleta: [" & Parametros.AgenciaSatelite & "]. Desconhecido."
                MsgBox "ATENÇÃO ! Erro " & Str(Geral.rst(0)) & " na abertura da agência de coleta  [" & Parametros.AgenciaSatelite & "]. Verifique! ", vbOKOnly + vbCritical, "Atenção"
                Close #20
                End
            End If
        End If
    Else
        AgenciaCadastradaBH = True
    End If

   'Calcula o NSU para a proxima transação
   ' CalculaNSU
   ' Caixa.NSU1 = Caixa.NSU
    
    Vezes = 0
    
    Exit Function

TrataErro:
    Vezes = Vezes + 1
    If Vezes = 5 Then
        Select Case TratamentoErro("Falha no módulo: [Consulta de Agência BH] .", Err)
            Case eSair
                End
            Case eRepetir
                Resume
            Case eContinuar
                Resume Next
        End Select
    Else
        Espera (0.5)
        Resume Reinicio
    End If
End Function
Function TituloExcedePrazoVC(ByVal pTabela As String, _
                             ByVal Vencimento As Long, _
                             ByVal PrazoVencimento_Mal As Integer, _
                             ByVal PrazoVencimento_Env As Integer) As Boolean
On Error GoTo TrataErro

   'Verifica Nr.Malote Novo ou Antigo
    If Geral.idEnvMal = "M" Then
        If PrazoVencimento_Mal > 0 Then
            If Vencimento + PrazoVencimento_Mal <= Geral.DataMovAnt Then
               TituloExcedePrazoVC = True
            End If
        Else
            TituloExcedePrazoVC = True
        End If
    Else
        If PrazoVencimento_Env > 0 Then
            If Vencimento + PrazoVencimento_Env <= Geral.DataMovAnt Then
                TituloExcedePrazoVC = True
            End If
        Else
            TituloExcedePrazoVC = True
        End If

    End If
    
    Exit Function

TrataErro:

Screen.MousePointer = 0
    Select Case TratamentoErro("Falha no módulo: [TituloExcedePrazoVC] .", Err)
        Case eSair
            End
        Case eRepetir
            Resume
        Case eContinuar
            Resume Next
    End Select
 
End Function
Sub LogCartaoAvulso()
   '====================================================='
   ' TRANSAÇÃO 36(nosso número) - Cartao Credito Avulso  '
   '====================================================='

    On Error GoTo TrataErro

    Dim valorReais      As String
    Dim valorDolar      As String
    Dim valorAntSaque   As String
    Dim ConvCartao      As String
    Dim crtTamanho      As Integer
    Dim RstMDI          As Recordset

   'variaveis do header
    Geral.CodTransacao = "0083"
    Geral.Evento = 711
    Geral.TipoTransacao = 2
    Geral.Capa = GetCapa(Geral.idEnvMal)
    Geral.IndTransac = " "
    Geral.CapaBack = 0
    Geral.TpRep = 0
    Geral.ValorTrans = formataValor(Geral.rstDoctos!Valor)
    
    valorReais = formataValor(Geral.rstDocto!DespReais)
    valorDolar = formataValor(Geral.rstDocto!DespDolar)
    valorAntSaque = formataValor(Geral.rstDocto!AntSaque)
        
    Geral.hsSQLa = "exec cxrcvisa "
   
   'monta header
    MontaHeader
    
   'monta parte variavel
    Set RstMDI = MDIQuery.getVerificaBinCartao(Trim(Geral.rstDocto!Cartao))
    
    crtTamanho = Len(Trim(Geral.rstDocto!Cartao))
    ConvCartao = String(16 - crtTamanho, "0") & Trim(Geral.rstDocto!Cartao)
    
    Geral.hsSQLa = Geral.hsSQLa & ", 1"
    Geral.hsSQLa = Geral.hsSQLa & ", ' '"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Format(RstMDI!crefscdagen, "0000")
    Geral.hsSQLa = Geral.hsSQLa & ", " & Format(RstMDI!crefsnuccor, "0000000")
    Geral.hsSQLa = Geral.hsSQLa & ", ' '"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 1"
    Geral.hsSQLa = Geral.hsSQLa & ", " & ConvCartao
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", ' '"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"

    Geral.hsSQLa = Geral.hsSQLa & ", " & ConvCartao
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(valorReais)
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(valorDolar)
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(valorAntSaque)
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorTrans)

    MontaComplementoVariavel

    LocalLog "Cartao Avulso - SP " & Geral.hsSQLa

    Exit Sub

TrataErro:
    Screen.MousePointer = 0
    Select Case TratamentoErro("Falha no módulo: [Cartão Avulso] .", Err)
        Case eSair
            End
        Case eRepetir
            Resume
        Case eContinuar
            Resume Next
    End Select
End Sub
Sub LogGPS()

  '=================================='
  ' TRANSAÇÃO 35(nosso número) - GPS '
  '=================================='
   
On Error GoTo TrataErro
    
    Dim ValINSS As String, ValEntidades As String, ValJuros As String, DataCompetencia As String
           
   'variaveis do header
    Geral.CodTransacao = "0099"
    Geral.Evento = 866
    Geral.TipoTransacao = 2
    Geral.Capa = GetCapa(Geral.idEnvMal)
    Geral.IndTransac = " "
    Geral.TpRep = 0
    Geral.ValorTrans = formataValor(Geral.rstDoctos!Valor)
     
    ValINSS = formataValor(Geral.rstDocto!ValorInss)
    ValEntidades = formataValor(Geral.rstDocto!ValorEntidades)
    ValJuros = formataValor(Geral.rstDocto!Juros)

    DataCompetencia = Mid(Geral.rstDocto!Competencia, 5, 2) & Mid(Geral.rstDocto!Competencia, 1, 4)   'mmaaaa
    
    Geral.hsSQLa = "exec recgps "
   
   'monta header
    MontaHeader
   
   'monta parte variavel
    MontaComplemento
   
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.rstDocto!CodigoPagamento)
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(DataCompetencia)
    Geral.hsSQLa = Geral.hsSQLa & ", '" & Format(Geral.rstDocto!Identificador, "00000000000000") & "'"
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(ValINSS)
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(ValEntidades)
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(ValJuros)
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorTrans)
    Geral.hsSQLa = Geral.hsSQLa & ", '" & Geral.rstDocto!Confirmacao & "'"
   
    MontaComplementoVariavel

    LocalLog "Gps - SP " & Geral.hsSQLa
    
    Exit Sub

TrataErro:

Screen.MousePointer = 0
    Select Case TratamentoErro("Falha no módulo: [GPS] .", Err)
        Case eSair
            End
        Case eRepetir
            Resume
        Case eContinuar
            Resume Next
    End Select
End Sub
Sub MontaHeaderDepInter()

    Parametros.AgenciaSatelite = Geral.rstCapa!agorig
    Geral.Hora = Format(Now, "HHMM")
    
   'O caixa só será aberto qdo estacao local com caixa fechado
    LogAberturaCaixa
        
   'parametros para header da procedure
    Geral.hsSQLa = Geral.hsSQLa & "  '" & Geral.CodTransacao & "'"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.VersaoAtual
    Geral.hsSQLa = Geral.hsSQLa & ", " & Parametros.AgenciaCentral
    Geral.hsSQLa = Geral.hsSQLa & ", " & Parametros.AgenciaSatelite
    
   'este parametro não verifica a existencia de contas e duplicidades
    Geral.hsSQLa = Geral.hsSQLa & ", 1"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.TipoTransacao
    Geral.hsSQLa = Geral.hsSQLa & ", 3"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.Caixa
    Geral.hsSQLa = Geral.hsSQLa & ", 1"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.NSU1
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.CapaBack
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.Hora
    Geral.hsSQLa = Geral.hsSQLa & ", '" & Geral.IndTransac & "'"
    
    If Geral.idEnvMal = "E" Then
       Geral.hsSQLa = Geral.hsSQLa & ", 6"
    Else
       Geral.hsSQLa = Geral.hsSQLa & ", 7"
    End If
    
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.TpRep
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.Evento
    Geral.hsSQLa = Geral.hsSQLa & ", '" & Geral.Capa & "'"

End Sub
Sub ProcessaCapa()

On Error GoTo TrataErro
            
    Dim RstUBB          As Recordset
    Dim RstMDI          As Recordset
    Dim spRetorno       As Integer
    Dim Ano             As String
    Dim ErroDescricao   As String
        
    ErroDescricao = "Obtenção de Parametros"
   
   'leitura da praça da Agência de Coleta
    Set RstUBB = UBBQuery.getPracaCompensacao(Geral.rstCapa!agorig)
    
   'Leitura da praça de compensação
    Parametros.PracaCompensacao = Format$(RstUBB("agefsnuprco"), "000")
           
   'leitura da data de movto anterior
    Ano = Format(Mid(Format$(RstUBB("agefsdtmvan"), "000000"), 5, 2), "00")
    If (Ano >= "00") And (Ano <= "51") Then
       Geral.DataMovAnt = "20" & Mid(Format$(RstUBB("agefsdtmvan"), "000000"), 5, 2) & Mid(Format$(RstUBB("agefsdtmvan"), "000000"), 3, 2) & Mid(Format$(RstUBB("agefsdtmvan"), "000000"), 1, 2) 'data no formato DDMMAAAA
    ElseIf (Ano > "51") And (Ano <= "99") Then
       Geral.DataMovAnt = "19" & Mid(Format$(RstUBB("agefsdtmvan"), "000000"), 5, 2) & Mid(Format$(RstUBB("agefsdtmvan"), "000000"), 3, 2) & Mid(Format$(RstUBB("agefsdtmvan"), "000000"), 1, 2) 'data no formato DDMMAAAA
    End If
             
   'Verifica se trata-se de envelope ou malote
    Geral.idEnvMal = Geral.rstCapa!idEnv_Mal
   
    If Not Geral.PrimeiraVez Then
       'Informa agencia de coleta p/ o UBB-NT
        ErroDescricao = "Agencia de coleta"
        AgenciaColeta
    End If
                                        
    If Geral.rstCapa!Status = "R" Then
    
        ErroDescricao = "Checa Capa"
        Set RstMDI = MDIQuery.getChecaCapa(Geral.DataProcessamento, Geral.rstCapa!idcapa, spRetorno)
                               
        If spRetorno <> 0 Then
        
            Call GaugeTitulo(2)
                        
            'Seta Capa para Ilegíveis ou CSP conforme Origem
            If Geral.rstCapa!csp Then
                LocalLog "Capa com problema de batimento setada para CSP: " & RstMDI!Msg
                
               'Insere Mensagem p/ Usuario
                spRetorno = MDIQuery.insMensagem(Geral.DataProcessamento, _
                                                 MSG_CapaProblemaBatimentoCSP, _
                                                 Geral.rstCapa!idcapa, 0, _
                                                 Caixa.Caixa)
            
                spRetorno = MDIQuery.UpdCapaStatusCaixaControle(Geral.DataProcessamento, Geral.rstCapa!idcapa, ST_CapaParaCSP, Caixa.Caixa)
                If spRetorno <> 0 Then MsgBox "Falha Procedure [UpdCapaStatusCaixaControle]", vbCritical + vbOKOnly
                
               'Insere log
                MDIQuery.insLog Geral.DataProcessamento, Geral.rstCapa!idcapa, "0", Caixa.UsuarioAtual, "126"
                                
            Else
                LocalLog "Capa com problema de batimento setada para ilegiveis: " & RstMDI!Msg
                
               'Insere Mensagem p/ Usuario
                spRetorno = MDIQuery.insMensagem(Geral.DataProcessamento, _
                             MSG_CapaProblemaBatimento, _
                             Geral.rstCapa!idcapa, 0, _
                             Caixa.Caixa)
            
                spRetorno = MDIQuery.UpdCapaStatusCaixaControle(Geral.DataProcessamento, Geral.rstCapa!idcapa, ST_CapaParaIlegivel, Caixa.Caixa)
                If spRetorno <> 0 Then MsgBox "Falha Procedure [UpdCapaStatusCaixaControle]", vbCritical + vbOKOnly
                
               'Insere log
                MDIQuery.insLog Geral.DataProcessamento, Geral.rstCapa!idcapa, "0", Caixa.UsuarioAtual, "124"

            End If
                        
            Exit Sub
         
        End If
        
    End If
                  
    If Geral.rstCapa!Ocorrencia <> 0 Then
       'encontrou capa para enviar ocorrencia
        ErroDescricao = "Envia Ocorrencia da Capa"
        EnviaOcorrenciaCapa
    Else
                
        Geral.GereiLog = 0
      
       'Inicia Transmissão
        ErroDescricao = "Processa Documentos"
        Call ProcessaDocumentos
        
Finaliza:

        ErroDescricao = "Finaliza"
        
       'Capa para CSP
        If Geral.PreparouLog = 4 Then
        
            LocalLog "Envia Capa para CSP " & Geral.rstCapa!Capa
            MDIQuery.insLog Geral.DataProcessamento, Geral.rstCapa!idcapa, "0", Caixa.UsuarioAtual, "126"

            spRetorno = MDIQuery.UpdCapaStatusCaixaControle(Geral.DataProcessamento, Geral.rstCapa!idcapa, ST_CapaParaCSP, Caixa.Caixa)
            If spRetorno <> 0 Then
                Err.Raise 999, App.Title, "Atualizar Status da Capa (CSP)"
            End If
            
           'frmShow.LabelTitulo.Caption = "Enviando Capa para CSP."
            Call GaugeTitulo(4)
            Espera (0.3)
            
            '''''''''''''''''''''''''''''''''''
            'Grava a procedure de fim de ciclo'
            '''''''''''''''''''''''''''''''''''
            
            LogFimCiclo
            
        ElseIf Geral.PreparouLog = 5 Then
        
            LocalLog "Envia Capa para Correcao Agencia/Conta " & Geral.rstCapa!Capa
            MDIQuery.insLog Geral.DataProcessamento, Geral.rstCapa!idcapa, "0", Caixa.UsuarioAtual, "127"
        
            spRetorno = MDIQuery.UpdCapaStatusCaixaControle(Geral.DataProcessamento, Geral.rstCapa!idcapa, ST_CapaCorrecaoAgConta, Caixa.Caixa)
            If spRetorno <> 0 Then
                Err.Raise 999, App.Title, "Atualizar Status da Capa (Correção AG/Conta)"
            End If
        
            Call GaugeTitulo(1)
            Espera (0.3)
        
           '''''''''''''''''''''''''''''''''''
           'Grava a procedure de fim de ciclo'
           '''''''''''''''''''''''''''''''''''
        
            LogFimCiclo
            
        ElseIf Geral.PreparouLog = 6 Then
            LocalLog "Envia Capa para Ilegiveis " & Geral.rstCapa!Capa
            MDIQuery.insLog Geral.DataProcessamento, Geral.rstCapa!idcapa, "0", Caixa.UsuarioAtual, "127"
        
            spRetorno = MDIQuery.UpdCapaStatusCaixaControle(Geral.DataProcessamento, Geral.rstCapa!idcapa, ST_CapaParaIlegivel, Caixa.Caixa)
            If spRetorno <> 0 Then
                Err.Raise 999, App.Title, "Atualizar Status da Capa (Ilegiveis)"
            End If
        
            Call GaugeTitulo(1)
            Espera (0.3)
        
           '''''''''''''''''''''''''''''''''''
           'Grava a procedure de fim de ciclo'
           '''''''''''''''''''''''''''''''''''
        
            LogFimCiclo

        ElseIf (Geral.PreparouLog <> 3) Then
            LocalLog "Atualiza Capa Transmitida " & Geral.rstCapa!Capa
            
            spRetorno = MDIQuery.UpdCapaStatusCaixaControle(Geral.DataProcessamento, Geral.rstCapa!idcapa, ST_CapaTransmitida, Caixa.Caixa)
            If spRetorno <> 0 Then
                Err.Raise 999, App.Title, "Atualizar Status da Capa (Transmitida)"
            End If
            
            '''''''''''''''''''''''''''''''''''
            'Grava a procedure de fim de ciclo'
            '''''''''''''''''''''''''''''''''''
            
            LogFimCiclo
                    
        End If
                          
    End If
    
    Exit Sub

TrataErro:

    LocalLog "Falha no processamento da Capa - Nota " & ErroDescricao
    
    Select Case TratamentoErro("Falha no Tratamento da Capa - Nota: " & ErroDescricao, Err, IIf(ErroDescricao <> "Finaliza", eCapa, eDefault))
            Case eSair
                End
            Case eRepetir
                Resume
            Case eContinuar
                Resume Next
            Case eFinalizar
                Resume Finaliza
        End Select
    
End Sub
Public Function LogRecepcao() As Boolean
    
On Error GoTo TrataErro

    Dim localStatus     As String * 1
    Dim spRetorno       As Integer

   'variaveis do header
    Geral.idEnvMal = Geral.rstCapa!idEnv_Mal
    Geral.CodTransacao = "IKRM"
    Geral.Evento = 706
    Geral.TipoTransacao = 0
    Geral.Capa = GetCapa(Geral.idEnvMal)
    Geral.IndTransac = " "
    Geral.CapaBack = 0
    Geral.TpRep = 0
    
    LocalLog "Inicio de Recepcao da Capa"
        
    Geral.hsSQLa = "exec tarccxma "

    MontaHeader
    
    Call GaugePos(Recepcao)
    Espera (0.1)
    
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.rstCapa!agorig
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Mid(Geral.DataProcessamento, 7, 2) & Mid(Geral.DataProcessamento, 5, 2) & Mid(Geral.DataProcessamento, 1, 4))
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"

    Set Geral.rst = UBBQuery.ExecuteSQL(Geral.hsSQLa)

    Call GaugePos(Recepcao)
    Espera (0.1)
  
    LocalLog "Recepcao da Capa - SP " & Geral.hsSQLa

    If Geral.rst(0) <> 0 Then
        MsgBox "ATENÇÃO ! - não foi possível atualizar 'Capa (" & Geral.rstCapa!idcapa & ") no MDI para recepção", vbOKOnly + vbCritical, "Atenção"
    End If

    LocalLog "Retorno da Recepcao: " & Geral.rst(0)
   
    If (Val(Geral.rst(0)) = 0) Then
       localStatus = "S"
    Else
       localStatus = "P"
    End If

    spRetorno = MDIQuery.updCapaRecepcionada(Geral.DataProcessamento, Geral.rstCapa!idcapa, localStatus)

    If spRetorno <> 0 Then
        MsgBox "ATENÇÃO !!! (83) Erro na Atualizacao da capa selecionada. Tecle <enter> para continuar.", vbOKOnly + vbCritical, "Atenção"
    End If

    Call GaugePos(Recepcao)
    Espera (0.1)
   
Exit Function

TrataErro:
   
    Screen.MousePointer = 0
    Select Case TratamentoErro("Atenção! - Falha na Rotina de Recepção de Capa", Err)
      Case eSair
          End
      Case eRepetir
          Resume
      Case eContinuar
          Resume Next
    End Select

End Function
Function Comunica() As Boolean

On Error GoTo TrataErro

    Dim RstMDI          As Recordset
    Dim spRetorno       As Integer
    Dim Processou       As Boolean
    Dim MudouCx         As Boolean
    Static Vezes
        
    DoEvents
    
    Geral.ehVinculoManual = False
    Geral.PrimeiraVez = True
    Geral.RecebendoCapa = False
    
    If GetSetting("Robo", "Capa", "Transmissao", 0) = 1 Then
    
        Set Geral.rstCapa = MDIQuery.getCapaPendente(Geral.DataProcessamento, Caixa.Caixa)
               
        If Not Geral.rstCapa.EOF Then
            
            LocalLog "Encontrada Capa Pendente Capa: " & Geral.rstCapa!Capa
        
           '************************'
           'Encontrou capa pendente '
           '************************'
          
            If AgenciaCadastradaBH() Then
            
                frmShow.LabelStatus = "Status: Em Transmissão..."
                DoEvents
                 
                Call ProcessaCapa
                
               'Antes de enviar a proxima capa, verifica se caixa deve ser fechado
                If Geral.CaixaAberto Then
                    LogFechamentoCaixa ("A")
                    Geral.CaixaAberto = False
                End If
                  
                If Geral.FecharCaixa Then
                    Call GetCaixa
                    Geral.FecharCaixa = False
                End If
            End If
            Processou = True
        End If
    End If
                
   '*****************************'
   'Verifica se existem estornos '
   '*****************************'
   
    Set Geral.rstCapa = MDIQuery.getDocumentoEstorno(Geral.DataProcessamento, Caixa.Estacao)
            
    If Not Geral.rstCapa.EOF Then
        LocalLog "Capa: " & Format(Geral.rstCapa!Capa, String(11, "0"))
        
       'Atualiza Tabela Estorno Docto
        spRetorno = MDIQuery.updStatusEstorno(Geral.DataProcessamento, _
                                              Geral.rstCapa!idcapa, _
                                              Geral.rstCapa!iddocto, _
                                              CodigoDoctoEmEstorno)
                                                        
       'Verifica se Docto foi transmitido por outro caixa
        If Caixa.Caixa <> Geral.rstCapa!Terminal Then
            LocalLog "Mudando Caixa para Efetivar Estorno"
            MudouCx = True
                        
            Set Geral.rst = UBBQuery.ExecuteSQL("Select * From Tfstcxag Where tcxfsnucaix = " & Caixa.Caixa)
                        
           'Fecha Caixa Anterior
            If Not Geral.rst.EOF Then
                LogFechamentoCaixa ("A")
            End If
            
            spRetorno = MDIQuery.AtualizarCaixaMDI(Geral.DataProcessamento, Caixa.Estacao, Caixa.Caixa, "N")
                        
           'Seleciona Caixa que transmitiu Docto e ultimo NSU usado no mesmo
            Caixa.Caixa = Geral.rstCapa!Terminal
            Caixa.BaseNSU = MDIQuery.getUltimoNSU(Geral.DataProcessamento, Geral.rstCapa!Terminal)

            frmShow.LabelTerminal.Caption = Caixa.Caixa
            
           'frmShow.LabelTerminal.Caption = Caixa.Caixa
            spRetorno = MDIQuery.AtualizarCaixaMDI(Geral.DataProcessamento, Caixa.Estacao, Caixa.Caixa, "S")
             
        End If

       'Existe docto a ser estornado
        frmShow.LabelStatus = "Status: Enviando Estorno..."
        DoEvents
        
       'Atualização do Gauge
        Call GaugeInit("Capa " & Format(Geral.rstCapa!Capa, "00000000000"), 4, Estorno)
        Call GaugePos(Estorno, Geral.rstCapa!Nome)
        Espera (0.5)
                                        
       'Verifica o Tipo de Docto, seleciona o Estorno e executa
        If ProcessaEstorno() Then
            Call GaugePos(Estorno, Geral.rstCapa!Nome)
            Espera (0.5)
            MDIQuery.insLog Geral.DataProcessamento, Geral.rstCapa!idcapa, Geral.rstCapa!iddocto, Caixa.UsuarioAtual, "128"
            Call AtualizaEstorno(True)
        Else
            Call GaugePos(Estorno, "F A L H O U    E S T O R N O")
            Espera (2)
            MDIQuery.insLog Geral.DataProcessamento, Geral.rstCapa!idcapa, Geral.rstCapa!iddocto, Caixa.UsuarioAtual, "129"
            Call AtualizaEstorno(False)
        End If
                                                   
       'Encerra estorno, fecha o caixa
        If Geral.CaixaAberto Then
            LogFechamentoCaixa ("A")
            Geral.CaixaAberto = False
        End If
        
        If MudouCx Then
            MudouCx = False
            spRetorno = MDIQuery.AtualizarCaixaMDI(Geral.DataProcessamento, Caixa.Estacao, Caixa.Caixa, "N")
            Call GetCaixa
        End If
                                    
        If Geral.FecharCaixa Then
            Geral.FecharCaixa = False
        End If
         
        frmShow.LabelTerminal.Caption = Caixa.Caixa
        Processou = True
      
    Else
        If ProcessaRecepcao Then
            Processou = True
            LocalLog "Capa Recepcionada: " & Format(Geral.rstCapa!Capa, String(11, "0"))
            
        Else
        
            '********************************'
            ' Procura Capas para Transmissão '
            '********************************'
            
            Set Geral.rstCapa = MDIQuery.getCapaTransmitir(Geral.DataProcessamento, Caixa.Caixa, spRetorno)
            
            If spRetorno = 0 Then
                LocalLog "Capa Transmissao: " & Geral.rstCapa!Capa
                SaveSetting appname:="Robo", section:="Capa", Key:="Transmissao", setting:=1
                
                frmShow.LabelStatus = "Status: Iniciando Transmissão..."
                DoEvents
                
                If AgenciaCadastradaBH() Then
                
                    Call ProcessaCapa
                    
                   'Antes de enviar proxima capa, verifica se caixa deve ser fechado
                    If Geral.CaixaAberto Then
                       LogFechamentoCaixa ("A")
                       Geral.CaixaAberto = False
                    End If
                    
                    If Geral.FecharCaixa Then
                       Call GetCaixa
                       Geral.FecharCaixa = False
                    End If
                     
                End If
                
                SaveSetting appname:="Robo", section:="Capa", Key:="Transmissao", setting:=0
                
                Processou = True
                
            End If
            
        End If
        
    End If
    
    frmShow.LabelStatus = ""
    frmShow.LabelTerminal.Caption = Caixa.Caixa
    
   'Se processou Capa e Usuario não fechou o Caixa Pesquisa Novamente
    If Processou And frmShow.ProgressDelayPesquisa.Tag <> "FIM" Then
        Comunica = True
        frmShow.ProgressDelayPesquisa.Tag = "OK"
    End If
    
    Call DestroyGauge
    
    Vezes = 0
    
    Exit Function
    
TrataErro:
    Vezes = Vezes + 1
    LocalLog "Falha modulo Comunica Vezes:" & Str(Vezes)

    If Vezes = 5 Then
            
        Select Case TratamentoErro("Não foi possível localizar a próxima capa.", Err)
          Case eSair
              End
          Case eRepetir
              Resume
          Case eContinuar
              Resume Next
        End Select
    End If
End Function
Sub ProcessaDocumentos()
   
On Error GoTo TrataErro

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Esta rotina envia transação a transação de cada capa     '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim QtdeRegistros       As Integer
    Dim spRetorno           As Integer
    Dim RstMDI              As Recordset
    Dim RstMdiVinculo       As Recordset
    Dim ContCheques         As Long
    Dim ErroDescricao       As String
    Static ConsultouVinculo
    
    ConsultouVinculo = Empty
    
    ContCheques = 0
    Geral.PreparouLog = 0
    Geral.VincProcAnt = 0
  
   'Traz todos os doctos da capa corrente
    Set Geral.rstDoctos = MDIQuery.getDocumentos(Geral.DataProcessamento, _
                                                 Geral.rstCapa!idcapa, _
                                                 spRetorno)
                                                 
    If spRetorno <> 0 Then Exit Sub
                                                 
   'Se houver depositos e/ou OCT valida agencia/conta dos mesmos
    ErroDescricao = "*Verifica Ag/Conta Depositos"
    If VerificaAgenciaConta Then Exit Sub
    
   'Se houver ajustes validar agencia / conta
    ErroDescricao = "*Verifica Ag/Conta Ajustes"
    If VerificaAgenciaContaAjustes() Then Exit Sub
                                                         
   'Seta qtde de docto para o Gauge
    If Geral.rstDoctos.RecordCount <= 0 Then QtdeRegistros = 1 Else QtdeRegistros = Geral.rstDoctos.RecordCount
    
   'Inicializa Gauge - Status da tranmissão e Painel
    Call GaugeInit("Capa " & Format(Geral.rstCapa!Capa, "00000000000"), QtdeRegistros, Transmissao)
    Call GaugeInit("", QtdeRegistros, Estorno)
    frmShow.LabelStatus.Caption = "Status: Em Transmissão..."
    frmShow.LabelQtde.Caption = QtdeRegistros
        
   'Log de inicio da transmissão dos documentos
    MDIQuery.insLog Geral.DataProcessamento, Geral.rstCapa!idcapa, "0", Caixa.UsuarioAtual, "121"
                                             
    Do While Not Geral.rstDoctos.EOF
    
        ErroDescricao = "*Verifica Documento com Ocorrencia"
        
        While Geral.rstDoctos!Status = "D"
           'verifica se ocorrencia deve ser informada ao BH
            If VerificaCapaOcorrencia(ContCheques) Then
                ContCheques = ContCheques + 1
                If ContCheques <= 0 Then
                    ContCheques = 1
                End If
                        
                Call GaugePos(Transmissao, "Enviando Ocorrência")
                
            End If
            
            Geral.rstDoctos.MoveNext
            
            If Geral.rstDoctos.EOF Then
                Exit Do
            End If
            
        Wend
                           
        ErroDescricao = "*Verifica mudanca de Vinculo"
        If ConsultouVinculo <> Geral.rstDoctos!Vinculo Then
            ConsultouVinculo = Geral.rstDoctos!Vinculo
            
           'Consulta Vencimento e Irregularidades na cobranças deste vinculo
            ErroDescricao = "*Consulta Cobranca"
            Call ConsultaCobrancaUBB(Geral.rstDoctos!Vinculo)
                               
            If Geral.BHAceitaCobranca = 0 Then
                Exit Sub
            End If
                    
        End If
        
        Set Geral.rstDocto = MDIQuery.getDocumento(Geral.DataProcessamento, _
                                           Geral.rstDoctos!iddocto, _
                                           Geral.rstDoctos!TipoDocto, _
                                           Caixa.Caixa, _
                                           spRetorno)
                                           
        ErroDescricao = "#Subindo Log"
                
        Geral.PreparouLog = 0
        Geral.CapaBack = 0
        Geral.TpRep = 2
        Geral.RetTransacao = 0
        Geral.CodOcorrencia = 0
                          
        Call GaugePos(Transmissao, Geral.rstDoctos!Nome)
        
       'Valida Chave de Encriptação, se invalida devolve documentos e capa para CSP
        ErroDescricao = "#Valida Criptografia"
        If Geral.Criptografia Then
           
            If Not ValidaEncriptaBO Then
                Geral.CodOcorrencia = 999
                Geral.RetTransacao = 91
                
                Call DevolveDocumentos
                Exit Sub
            
            End If
        End If
                
       'Processamento das transações
        ErroDescricao = "#Subindo Log"
        LocalLog "Inicia Transmissao Docto: " & Geral.rstDoctos!Nome
        
       ' 5_6_8_9_12_20_21_22_23_24_25_26_42_43
        
        Select Case Geral.rstDoctos!TipoDocto
            
            Case 2                  ' Deposito c.corrente
                LogDepositoCC
            Case 3                  ' Deposito c.poupanca
                LogDepositoCP
            Case 4                  ' ADCC - Aviso de Débito
                LogADCC
            Case 5                  ' Saque cheque UBB interagencia
                If Mid(Geral.rstDoctos!leitura, 23, 6) = "688611" Or _
                   Mid(Geral.rstDoctos!leitura, 23, 6) = "688612" Or _
                   Mid(Geral.rstDoctos!leitura, 4, 4) = "0927" Then 'Cheque ADM
                                          
                    LogChequeCompensacao
                    AlteraParaCompensacao
                Else
                    LogChequeInteragencia ' Saque cheque UBB interagencia
                End If
            Case 6                  ' Cheque terceiro pagto
                LogChequeCompensacao
            Case 7                  ' Cheque deposito - não tem log
                'Nao ha transacao para este docto ele sobe junto com a capa de deposito
            Case 11                 ' INSS
                LogINSS
            Case 12                 ' Cobrança terceiros - SEM codigo de barras
                LogTitulo
            Case 13                 ' Cobrança registrada via teclado - SEM codigo de barras
                LogCobRegistradaSemCB
            Case 14                 ' Cobrança especial via teclado - SEM codigo de barras
                LogCobEspecialSemCB
            Case 15                 ' DARM
                LogDarm
            Case 16                 ' DARF - Preto
                LogDarfPreto
            Case 17                 ' DARF - Simples
                LogDarfSimples
            Case 18                 ' GARE
                LogGARE
            Case 19                 ' GRPS
                LogGRPS
            Case 20 To 26           ' Trib.municipais,estaduais,federais e arrecadacoes
                LogConcessionarias
            Case 27                 ' Arrecadacao convencional
                LogArrecConvenc
            Case 28                 ' Unicobrança - COM codigo de barras
                LogUnicobrancaUBB
            Case 29                 ' Cobrança Imediata UNIBANCO - COM codigo de barras
                LogCobImediataUBB
            Case 30                 ' Cobrança especial - COM codigo de barras
                LogCobEspecialUBB
            Case 31                 ' Cobrança terceiros - COM codigo de barras
                LogCobTerceiroComCB
            Case 32, 33
                LogAjusteDeposito
            Case 34                 ' Aviso de credito - cc
                LogAjusteCredito
            Case 35                 ' gps
                LogGPS
            Case 36                 ' cartao credito avulso
                LogCartaoAvulso
            Case 37                 ' OCT
                LogOCT
            Case 38                 ' Aviso de debito - cc
                LogAjusteDebito
            Case 40                 ' FGTS
            
                If InStr(1, "0178_0179_0180_0181_0182", Mid(Geral.rstDoctos!leitura, 16, 4), vbTextCompare) <> 0 Then
                    Call LogFGTS1
                Else
                    Call LogFGTS
                End If
                
            Case 41
                ' Lancamento Interno
                'If UCase(Command) <> "DEBUG" Then
                    LogLanctoInterno
                'Else
                '    LogLiTemp
                'End If
            Case 42                 ' Ajuste Contabil Receita
                LogContabilCredito
            Case 43                 ' Ajuste Contabil Despesa
                LogContabilDebito
            Case Else
                Geral.PreparouLog = 1
                Screen.MousePointer = 0
                MsgBox "LOG ainda não especificado. - " & Geral.rstDoctos!TipoDocto & " - Capa: " & Format(Geral.rstCapa!Capa, "00000000000"), vbOKOnly + vbCritical, "Atenção"
                Beep
        End Select
        
        ErroDescricao = "Fim Transmissao"
       
       'Rotinas genericas sobe o Docto ou rejeita
        If (Geral.PreparouLog = 0) Then
           
            If Geral.rstDoctos!TipoDocto <> 32 And Geral.rstDoctos!TipoDocto <> 33 And Geral.rstDoctos!TipoDocto <> 37 And Geral.rstDoctos!TipoDocto > 7 Then
                Call LogGeral
            End If
            
        ElseIf Geral.PreparouLog = 1 Then
            Call DevolveDocumentos
        End If
        
       'Call GaugePos(Transmissao, Geral.rstDoctos!Nome)
        LocalLog "Fim Transmissão Docto: " & Geral.rstDoctos!Nome
        
       'Final da transmissao por problemas
        If Geral.PreparouLog = 3 Or Geral.PreparouLog = 4 Or Geral.PreparouLog = 5 Then
            Exit Do
        End If
        
        Geral.rstDoctos.MoveNext
            
    Loop
    
    frmShow.LabelStatus = "Status: Finalizando ..."
    
    MDIQuery.insLog Geral.DataProcessamento, Geral.rstCapa!idcapa, "0", Caixa.UsuarioAtual, "122"
   
    ContCheques = 0
    Espera (0.2)
   
    Exit Sub

TrataErro:
    LocalLog "Falha no processamento do Documento - Nota " & ErroDescricao
    
   'se ja aberto rst especifico
    Select Case TratamentoErro("Falha na Processa Doctos. Nota: " & ErroDescricao, Err, IIf(Left(ErroDescricao, 1) = "#", eDoctoSubidaLog, eDoctoProcesso))
        Case eSair
            Call DestroyGauge
            End
        Case eRepetir
            Resume
        Case eContinuar
            Resume Next
        Case Else
            Exit Sub
    End Select
    
End Sub
Sub LogAjusteDebito()

  '=================================================
  ' TRANSAÇÃO 38(nosso número) - debito automatico '
  '================================================='

On Error GoTo TrataErro:

    Dim RstMDI  As Recordset

    Geral.Capa = GetCapa(Geral.idEnvMal)
    LocalLog "Gravacao Ik20 (ocorrencia 203) - Cheque menor do que as contas."

    'envia log de ocorrencia para o ajuste de debito
    Set RstMDI = MDIQuery.getTipoPagto(Geral.DataProcessamento, Geral.rstCapa!idcapa, Geral.rstDoctos!Vinculo)
    
    If RstMDI!TipoDocto = 4 Then
       Geral.CodOcorrencia = 402  'ADCC
    Else
       Geral.CodOcorrencia = 203  'Cheque
    End If

    Geral.Transacao = ""
    LogOcorrencia
   
   'variaveis do header
    Geral.CodTransacao = "0015"
    Geral.Evento = 580
    
    Geral.TipoTransacao = 1
    Geral.IndTransac = "O"
    
    Geral.Capa = GetCapa(Geral.idEnvMal)
   
   'aviso debito em conta corrente
    Geral.TipoConta = "C"
      
    Geral.IdentDep = 0
    Geral.AgenciaVinculo = Geral.rstDocto!Agencia
    Geral.ContaVinculo = Geral.rstDocto!Conta
    
    Set RstMDI = MDIQuery.GetBancoOrigemDoAjuste(Geral.DataProcessamento, _
                                                 Geral.rstCapa!idcapa, _
                                                 Geral.rstDoctos!Vinculo)

    If RstMDI!Banco = "230" And Val(Format(Geral.rstDocto!Agencia, "0000") & _
        Format(Geral.rstDocto!Conta, "0000000")) <> Geral.rstCapa!Num_malote Then

       'pesquisa tabela depara 230x409
        If Not Depara230() Then
            Exit Sub
        End If
    Else
       'pesquisa tabela depara 409x409 (antiga)
        DePara
    End If

    If (Geral.PreparouLog = 1) Then
        Exit Sub
    End If
   
    Geral.ValorTrans = formataValor(Geral.rstDoctos!Valor)

   'stored procedure do aviso debito
    Geral.hsSQLa = "exec avintpar "

   'monta header
    MontaHeader

   'monta parte variavel
    Geral.hsSQLa = Geral.hsSQLa & ", 2"
    Geral.hsSQLa = Geral.hsSQLa & ", 1565"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.AgenCob
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.ContaCob
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorTrans)
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Val(Parametros.DataServer)
    Geral.hsSQLa = Geral.hsSQLa & ", ' '"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 1"
    Geral.hsSQLa = Geral.hsSQLa & ", ' '"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"

    LocalLog "Ajuste Debito (cheque < contas) - SP " & Geral.hsSQLa

Exit Sub

TrataErro:
Screen.MousePointer = 0

    Select Case TratamentoErro("Falha no módulo: [LogAjusteDebito] .", Err)
        Case eSair
            End
        Case eRepetir
            Resume
        Case eContinuar
            Resume Next
    End Select
End Sub
Sub MontaHeaderDebito()
    
    Parametros.AgenciaSatelite = Geral.rstCapa!agorig
    Geral.Hora = Format(Now, "HHMM")
    
   'O caixa só será aberto qdo estacao local com caixa fechado
    LogAberturaCaixa
        
    Geral.hsSQLa = Geral.hsSQLa & "  '" & Geral.CodTransacao & "'"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.VersaoAtual
    Geral.hsSQLa = Geral.hsSQLa & ", " & Parametros.AgenciaCentral
    Geral.hsSQLa = Geral.hsSQLa & ", " & Parametros.AgenciaSatelite
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.TipoTransacao
    Geral.hsSQLa = Geral.hsSQLa & ", 1"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.Caixa
    Geral.hsSQLa = Geral.hsSQLa & ", 1"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.NSU1
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.Hora
    Geral.hsSQLa = Geral.hsSQLa & ", '" & Geral.IndTransac & "'"
    Geral.hsSQLa = Geral.hsSQLa & ", 7"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.Evento
    Geral.hsSQLa = Geral.hsSQLa & ", '" & Geral.Capa & "'"

End Sub
Sub LogOCT()

On Error GoTo TrataErro:

   '===================================
   ' TRANSAÇÃO 37(nosso número) - OCT '
   '=================================='
    Dim RstMDI                      As Recordset
    Dim RstUBB_BHQC                 As Recordset     'Capa de cheques
    Dim RstUBB_BHQQ                 As Recordset     'Cheques
    Dim RstMdiChequesDeposito       As Recordset
    Dim spRetorno                   As Integer
    Dim contLaco                    As Integer
    Dim totLaco                     As Integer
    Dim i                           As Integer
    Dim ContCheques                 As Integer
    Dim QtdeCheques                 As Integer
    Dim ValorCheque                 As String
    Dim ValorDinheiro               As String
    Dim ValorSomado                 As String
    Dim DigitoData                  As Integer
    Dim NumeroOCT                   As String
    
    ValorCheque = 0
    ValorSomado = 0
    ValorDinheiro = 0
       
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' Obtem os cheque Deposito (Cheque diversos) ou dinheiro (Cheque UBB / LI) '
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Not GetChequeCashDeposito(Geral.rstCapa!idcapa, Geral.rstDoctos!Vinculo, Geral.rstDoctos!Valor, QtdeCheques, ValorDinheiro, ValorCheque, ValorSomado, True) Then Exit Sub
    
    Geral.ValorTrans = formataValor(Geral.rstDoctos!Valor)
    
   'calculo do digito da data
    DigitoData = Modulo10(Format(Parametros.AgenciaSatelite, "0000") & Parametros.DataServer, 10)
    
   'pesquisa tabela depara
    Geral.AgenciaVinculo = Geral.rstDocto!agenciaCredito
    Geral.ContaVinculo = Geral.rstDocto!ContaCredito
    
    DePara
   
    If Geral.PreparouLog = 1 Then
        Exit Sub
    End If
    
    Set RstMDI = MDIQuery.GetCMC7CapaOCT(Geral.DataProcessamento, Geral.rstCapa!idcapa, Geral.rstDoctos!Vinculo)
    
    If Not RstMDI.EOF Then
    
       'verifica se existe o CMC-7 da capa de OCT, caso não exista, devolver a OCT.
        If Val(RstMDI!leitura) = 0 Then
            Geral.PreparouLog = 1
            Exit Sub
        End If
    
        NumeroOCT = Mid(RstMDI!leitura, 9, 3) & Mid(RstMDI!leitura, 12, 6) & Mid(RstMDI!leitura, 4, 4)
    Else
        Geral.CodOcorrencia = 999
        Geral.RetTransacao = 92
        Call DevolveDocumentos
        Exit Sub
    End If
    
   'NumeroOCT = Mid(Geral.rstDoctos!Leitura, 9, 3) & Mid(Geral.rstDoctos!Leitura, 12, 6) & Mid(Geral.rstDoctos!Leitura, 4, 4)
    
   'variaveis para o header - ENVIO LOCAL
    Geral.CodTransacao = "0180"
    Geral.Evento = 828
    Geral.TipoTransacao = 2
    
    Geral.Capa = Format(Geral.rstCapa!Num_malote, "00000000000") 'malote
    
    Geral.IndTransac = " "
    Geral.CapaBack = 1
    Geral.TpRep = 0
      
   '=====================================
   '  Executa stored procedure CAIXAOCT '
   '=====================================
    
   'stored procedure da OCT
    Geral.hsSQLa = "Exec CaixaOct "
       
   'monta header
    MontaHeader
      
   'monta parte variavel
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Parametros.AgenciaSatelite
    Geral.hsSQLa = Geral.hsSQLa & ", " & Val(Parametros.DataServer)
    Geral.hsSQLa = Geral.hsSQLa & ", " & DigitoData
    Geral.hsSQLa = Geral.hsSQLa & ", 77500"
    Geral.hsSQLa = Geral.hsSQLa & ", '" & Format(Geral.rstDocto!OrdemCredito, "000000000000000") & "'"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.AgenCob
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.ContaCob
    Geral.hsSQLa = Geral.hsSQLa & ", '" & Mid(Geral.rstDocto!Referencia, 1, 20) & "'"
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(ValorDinheiro)
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(ValorCheque)
    Geral.hsSQLa = Geral.hsSQLa & ", '" & NumeroOCT & "'"
    Geral.hsSQLa = Geral.hsSQLa & ", " & QtdeCheques
    Geral.hsSQLa = Geral.hsSQLa & ", '" & Parametros.TipoAgencia & "'"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"     'cif_capt
    Geral.hsSQLa = Geral.hsSQLa & ", 0"     'age_hst
    Geral.hsSQLa = Geral.hsSQLa & ", 0"     'cta_hst
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    
   'CalculaNSU
    Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.NSU1
    
    Geral.ValorTrans = ValorCheque
    CalculaNSU
    Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.NSU
    Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.SDV
    
    CalculaNSU
    Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.NSU
    Geral.SeqPagto = Caixa.NSU2
    Geral.SeqRecto = Caixa.NSU1
    
    If GetDocumentoTransmitido(EnumDeposito) Then
        Exit Sub
    End If
        
   'gravar o novo nsu desta transação antes de envia-la para o UBB-NT
    spRetorno = MDIQuery.updNsuDocto(Geral.DataProcessamento, _
                                     Geral.rstCapa!idcapa, _
                                     Geral.rstDoctos!iddocto, _
                                     Caixa.NSU1, _
                                     Caixa.Caixa)
    
    If spRetorno <> 0 Then
        MsgBox "5490. ATENÇÃO! OCT a ser enviado Log, não atualizado o NSU. ", vbOKOnly + vbCritical, "Atenção"
        Exit Sub
    End If
       
    LocalLog "OCT - SP " & Geral.hsSQLa
    
   'Executa sp da OCT
    Set Geral.rst = UBBQuery.ExecuteSQL(Geral.hsSQLa)
    LocalLog "ret sp-caixaOCT - " & Format(Geral.rst(0), "00")
    
    If (Val(Geral.rst(0)) = 0) Then
        
        Geral.GereiLog = 1
        
        spRetorno = MDIQuery.updDoctoTransmitido(Geral.DataProcessamento, _
                                                 Geral.rstCapa!idcapa, _
                                                 Geral.rstDoctos!iddocto, _
                                                 Geral.SeqRecto, _
                                                 Caixa.Caixa, "N")
        
        If spRetorno <> 0 Then
            MsgBox "339. ATENÇÃO! Documento OCT - já enviado Log não atualizado no SQL. ", vbOKOnly + vbCritical, "Atenção"
            Exit Sub
        End If
                
        If ValorCheque <> 0 Then
        
            '=================================================
            '      Montando a BHQC   (Capa de Cheques)   SP3 '
            '=================================================
         
           'variaveis para o Header da BHQC
            Geral.Hora = Format(Now, "HHMM")
       
           'O caixa só será aberto qdo estacao local com caixa fechado
            Geral.ValorTrans = ValorCheque
            LogAberturaCaixa
         
            Geral.SeqBHQC = Caixa.NSU1
       
           'monta parte fixa
            Geral.hsSQLa = "exec compbhqc "
            Geral.hsSQLa = Geral.hsSQLa & "  'BHQC'"
            Geral.hsSQLa = Geral.hsSQLa & ", 0"
            Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.VersaoAtual
            Geral.hsSQLa = Geral.hsSQLa & ", " & Parametros.AgenciaCentral
            Geral.hsSQLa = Geral.hsSQLa & ", " & Parametros.AgenciaSatelite
            Geral.hsSQLa = Geral.hsSQLa & ", 0"
            Geral.hsSQLa = Geral.hsSQLa & ", 1"
            Geral.hsSQLa = Geral.hsSQLa & ", 1"
            Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.Caixa
            Geral.hsSQLa = Geral.hsSQLa & ", 1"
            Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.NSU1
            Geral.hsSQLa = Geral.hsSQLa & ", 0"
            Geral.hsSQLa = Geral.hsSQLa & ", 1"
            Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.Hora
            Geral.hsSQLa = Geral.hsSQLa & ", ' '"
            Geral.hsSQLa = Geral.hsSQLa & ", 7"
            Geral.hsSQLa = Geral.hsSQLa & ", 1"
            Geral.hsSQLa = Geral.hsSQLa & ", 827"
            Geral.hsSQLa = Geral.hsSQLa & ", '" & Geral.Capa & "'"
            Geral.hsSQLa = Geral.hsSQLa & ", '" & Caixa.CIF
            Geral.hsSQLa = Geral.hsSQLa & "', '" & Caixa.SDV
                     
           'monta parte variavel
            Geral.hsSQLa = Geral.hsSQLa & "', 0"
            Geral.hsSQLa = Geral.hsSQLa & ", '" & NumeroOCT & "'"
            Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(ValorCheque)
            Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(ValorSomado)
            Geral.hsSQLa = Geral.hsSQLa & ", " & QtdeCheques
            Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.AgenciaVinculo
            Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.ContaCob
            Geral.hsSQLa = Geral.hsSQLa & ", 1"
            Geral.hsSQLa = Geral.hsSQLa & ", ' '"
                     
            '***************************************************************'
            ' Verifica se valor informado é diferente do valor processado.  '
            ' Caso seja diferente, devemos informar para a stored procedure '
            ' o número do NSU para o acerto a ser criado pela mesma.        '
            '***************************************************************'
          
            If ValorCheque <> ValorSomado Then
            
                Geral.ValorTrans = Abs(ValorCheque - ValorSomado)
                
                CalculaNSU
                Caixa.NSU2 = Caixa.NSU
                Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.NSU2
                CalculaNSU
                Caixa.NSU3 = Caixa.NSU
                Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.NSU3
                
            Else
                
                Geral.hsSQLa = Geral.hsSQLa & ", 0"
                Geral.hsSQLa = Geral.hsSQLa & ", 0"
                
            End If
            
           'SDV2
            Geral.hsSQLa = Geral.hsSQLa & ", 0"
         
            '=========================================
            '  Executa stored procedure da capa BHQC '
            '=========================================
          
            LocalLog Geral.hsSQLa
            Set RstUBB_BHQC = UBBQuery.ExecuteSQL(Geral.hsSQLa)
                              
            LocalLog "ret sp-BHQC OCT - " & Format(RstUBB_BHQC(0), "00")
           
            If (Val(RstUBB_BHQC(0)) = 0) Then
                   
                '===========================================================================
                ' A Stored BHQQ deve ser enviada a cada lote de 5 cheques. Se nâo houverem '
                ' 5 cheques, ela é enviada assim mesmo, com zeros a direita.               '
                ' Para isso foi calculado quantas vezes a mesma será enviada.              '
                '===========================================================================
               
               'totLaço = qtos laços de 5 cheques existem neste OCT.
                totLaco = QtdeCheques \ 5
                If QtdeCheques Mod 5 <> 0 Then
                    totLaco = totLaco + 1
                End If
         
                '================================================'
                ' Procedure para ler novamente os cheques do OCT '
                '================================================'
               
                Set RstMdiChequesDeposito = MDIQuery.GetChequesDeposito(Geral.DataProcessamento, _
                                                                        Geral.rstCapa!idcapa, _
                                                                        Geral.rstDoctos!Vinculo)
                ContCheques = 0
                     
               'envio de cada laço de 5 cheques
                For contLaco = 1 To totLaco
                                                                                        
                    Geral.TpCtaBHQQ = 1
                    Geral.TipoOperacaoDeposito = "06"
                
                   'Monta o Header e parte variável da BHQQ para cada 05 cheques
                    MontaHeaderBHQQ
                    
                   'continua com montagem da parte variável
                    Geral.hsSQLa = Geral.hsSQLa & ", " & Val(RstUBB_BHQC(30))
    
                   'envio de cada 05 cheques
                    Do
                 
                        ContCheques = ContCheques + 1
                    
                        CalculaNSU
                                
                        Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.NSU
                        Geral.hsSQLa = Geral.hsSQLa & ", '" & RstMdiChequesDeposito!leitura & "'"
                    
                        If Not ValidaCMC7Cheque(Trim(RstMdiChequesDeposito!leitura)) Then
                            MsgBox "Encontrada inconsistencia no CMC7 cheque-deposito", vbCritical + vbOKOnly, "ATENÇÃO: COMUNIQUE SUPORTE"
                            MDIQuery.insLogErro Geral.DataProcessamento, Caixa.Estacao, 123456, "Encontrada inconsistencia no CMC7 cheque-deposito, Capa: " & Trim(Geral.Capa)
                            End
                        End If
                   
                       'verifica o tipo do cheque (1ºposição do nro do cheque)
                        If Mid(RstMdiChequesDeposito!leitura, 12, 1) = "8" Then
                            Geral.hsSQLa = Geral.hsSQLa & ", 9"     'tipo do cheque - cheque roxo
                        Else
                            Geral.hsSQLa = Geral.hsSQLa & ", 5"     'tipo do cheque - cheque comum
                        End If
                   
                       'valor deste cheque
                        Geral.hsSQLa = Geral.hsSQLa & ", " & formataValor(RstMdiChequesDeposito!Valor)
                  
                       'tipo de compensação do cheque, se cheque terceiro, e valor >= valor inferior, então este é SUPERIOR
                        If (Mid(RstMdiChequesDeposito!leitura, 1, 3) <> "409") And (RstMdiChequesDeposito!Valor >= Val(Parametros.ValorLimiteInferior)) Then
                            Geral.hsSQLa = Geral.hsSQLa & ", 1"
                        Else
                           'se cheque terceiro, e valor < valor inferior, então este é INFERIOR
                            If (Mid(RstMdiChequesDeposito!leitura, 1, 3) <> "409") And (RstMdiChequesDeposito!Valor < Val(Parametros.ValorLimiteInferior)) Then
                                Geral.hsSQLa = Geral.hsSQLa & ", 2"
                            Else
                               'se cheque UBB, e conta = 688111 ou 688112, então este é ADM
                                If ((Mid(RstMdiChequesDeposito!leitura, 1, 3) = "409") Or (Mid(RstMdiChequesDeposito!leitura, 1, 3) = "415")) And _
                                    ((Mid(RstMdiChequesDeposito!leitura, 23, 6) = "688111") Or (Mid(RstMdiChequesDeposito!leitura, 23, 6) = "688112")) Then
                                    Geral.hsSQLa = Geral.hsSQLa & ", 4"
                                Else
                                   'se cheque UBB, e conta <> 688111 E 688112, então este é INTERNA
                                    If ((Mid(RstMdiChequesDeposito!leitura, 1, 3) = "409") Or (Mid(RstMdiChequesDeposito!leitura, 1, 3) = "415")) And _
                                        ((Mid(RstMdiChequesDeposito!leitura, 23, 6) <> "688111") And (Mid(RstMdiChequesDeposito!leitura, 23, 6) <> "688112")) Then
                                        Geral.hsSQLa = Geral.hsSQLa & ", 3"
                                    End If
                                End If
                            End If
                        End If
                  
                       '''''''''''''''''''''''''''''''''''
                       ' Atualiza cheque como já enviado '
                       '''''''''''''''''''''''''''''''''''
                        Geral.GereiLog = 1
                       
                        spRetorno = MDIQuery.updChequeDepositoTransmitido(Geral.DataProcessamento, _
                                                                          Geral.rstCapa!idcapa, _
                                                                          RstMdiChequesDeposito!iddocto, _
                                                                          Caixa.NSU, _
                                                                          Caixa.Caixa)
                        If spRetorno <> 0 Then
                            MsgBox "160. ATENÇÃO! Documento - cheque de deposito - já enviado Log não atualizado no SQL. ", vbOKOnly + vbCritical, "Atenção"
                            Exit Sub
                        End If
                    
                        Call GaugePos(Transmissao, "Cheque Deposito")
                        Espera (0.2)
                        RstMdiChequesDeposito.MoveNext
                        If ContCheques = (5 * contLaco) Then Exit Do
                   
                    Loop Until RstMdiChequesDeposito.EOF
                 
                   'se for ultimo laço de cheques, seta flag de Fim
                    If ContCheques < QtdeCheques Then
                        Geral.hsSQLa = Geral.hsSQLa & ", 'C'"   'indicativo que continua
                    Else
                 
                       'preenche os cheques com branco sobrando no laço
                        If QtdeCheques Mod 5 <> 0 Then
                            For i = 1 To 5 - (QtdeCheques Mod 5)
                                Geral.hsSQLa = Geral.hsSQLa & ", 0"
                                Geral.hsSQLa = Geral.hsSQLa & ", ' '"
                                Geral.hsSQLa = Geral.hsSQLa & ", 0"
                                Geral.hsSQLa = Geral.hsSQLa & ", 0"
                                Geral.hsSQLa = Geral.hsSQLa & ", 0"
                            Next
                        End If
                        Geral.hsSQLa = Geral.hsSQLa & ", 'F'"   'indicativo que NAO continua
                    End If
                 
                   'novo campo numero da oct
                    Geral.hsSQLa = Geral.hsSQLa & ", '0" & String(14, " ") & "'"   'campo ordem credito em branco - char(15)
                 
                   '***********************************************************'
                   '  Executa stored procedure de cada BHQQ com até 05 cheques '
                   '***********************************************************'
                
                    LocalLog "Cheques do Deposito - SP " & Geral.hsSQLa
                    Set RstUBB_BHQQ = UBBQuery.ExecuteSQL(Geral.hsSQLa)
                    LocalLog "ret sp-BHQQ C/C - " & Format(RstUBB_BHQQ(0), "00")
                   
                    If Val(RstUBB_BHQQ(0)) <> 0 Then
                        Call DevolveDocumentos(RstUBB_BHQQ)
                        Exit Sub
                    End If
               
                   'Call GaugePos(Transmissao, Geral.rstDoctos!Nome)
                             
                Next
           
                Geral.GereiLog = 1
               
            Else
                   
                LocalLog "Ocorreu o seguinte retorno na sp-BHQC " & Format(RstUBB_BHQC(0), "00")
                Call DevolveDocumentos(RstUBB_BHQC)
                
            End If
       
        End If
           
    Else
          
        Call DevolveDocumentos
        
    End If

    Exit Sub
    
TrataErro:

    Select Case TratamentoErro("Falha na OCT.", Err, eDoctoSubidaLog)
        Case eSair
            End
        Case eRepetir
            Resume
        Case eContinuar
            Resume Next
        Case eFinalizar
            Exit Sub
    End Select

End Sub
Sub MontaHeaderBHQQ_OCT(Optional ByVal TpCtaBHQQ As String, _
                        Optional ByVal TipoOpDep As String, _
                        Optional ByVal SeqBHQC As String)
    
   '======================================'
   ' Monta a BHQQ   (Detalhe dos Cheques) '
   '======================================'
    
   'variaveis para o header
    Geral.Capa = Format(Geral.rstCapa!Num_malote, "00000000000") 'malote
    Geral.Hora = Format$(Now, "HHMM")
      
   'O caixa só será aberto qdo estacao local com caixa fechado
    LogAberturaCaixa (True)
    
    Geral.hsSQLa = "exec cheqdepo "
   
    Geral.hsSQLa = Geral.hsSQLa & "  'BHQQ'"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.VersaoAtual
    Geral.hsSQLa = Geral.hsSQLa & ", " & Parametros.AgenciaCentral
    Geral.hsSQLa = Geral.hsSQLa & ", " & Parametros.AgenciaSatelite
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 2"
    Geral.hsSQLa = Geral.hsSQLa & ", 1"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.Caixa
    Geral.hsSQLa = Geral.hsSQLa & ", 1"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.NSU1
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 1"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.Hora
    Geral.hsSQLa = Geral.hsSQLa & ", ' '"
    Geral.hsSQLa = Geral.hsSQLa & ", 7"
    Geral.hsSQLa = Geral.hsSQLa & ", 1"
    Geral.hsSQLa = Geral.hsSQLa & ", 826"
    Geral.hsSQLa = Geral.hsSQLa & ", '" & Geral.Capa & "'"
   
   'parte variável
    Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.Caixa
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.SeqPagto
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.SeqRecto
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.TpCtaBHQQ
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.AgenCob
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.ContaCob
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.SeqBHQC
    Geral.hsSQLa = Geral.hsSQLa & ", '" & Geral.TipoOperacaoDeposito & "'"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Val(Parametros.PracaCompensacao)

End Sub
Sub EnviaOcorrenciaCapa()
   
On Error GoTo TrataErro
    Dim spRetorno   As Integer

    Call DestroyGauge
    Call GaugeInit("Capa " & Format(Geral.rstCapa!Capa, "00000000000"), 2, Transmissao)
    
    Geral.CodOcorrencia = Geral.rstCapa!Ocorrencia
    Geral.Capa = GetCapa(Geral.idEnvMal)
    
   'gera ocorrencia
    Geral.Transacao = ""
    LogOcorrencia
    
    Call GaugePos
   
   'atualiza o status desta capa para X (já enviou ocorrencia)
    spRetorno = MDIQuery.updStatusCapa(Geral.DataProcessamento, Geral.rstCapa!idcapa, "X")

    If spRetorno <> 0 Then
        MsgBox "ATENÇÃO !!! (25) Capa com ocorrencia enviada não atualizado. Tecle <enter> para continuar.", vbOKOnly + vbCritical, "Atenção"
    End If
   
    Call GaugePos
    Call DestroyGauge
    
    Exit Sub
    
TrataErro:

    Screen.MousePointer = 0
    Select Case TratamentoErro("Não foi possível atualizar o status da capa com a informação de já enviado ocorrência.", Err)
        Case eSair
            End
        Case eRepetir
            Resume
        Case eContinuar
            Resume Next
    End Select

End Sub
Sub LogINSS()
   
   '========================================='
   ' TRANSAÇÃO 11(nosso número) - GR6 (INSS) '
   '========================================='
   
   'variaveis do header
    Geral.CodTransacao = "0039"
    Geral.Evento = 577
    Geral.TipoTransacao = 2
    Geral.Capa = GetCapa(Geral.idEnvMal)
    Geral.IndTransac = " "
    Geral.CapaBack = 0
    Geral.TpRep = 0
    Geral.ValorTrans = formataValor(Geral.rstDoctos!Valor)
      
    Geral.hsSQLa = "exec recinss "
    
   'monta header
    MontaHeader
   
   'monta parte variavel
    MontaComplemento
   
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.rstDoctos!identificacao)
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Mid(Geral.rstDoctos!Competencia, 5, 2) & Mid(Geral.rstDoctos!Competencia, 3, 2))  'competencia MMAA
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorTrans)
    
    MontaComplementoVariavel
    
    LocalLog "Inss - SP " & Geral.hsSQLa

End Sub
Sub DePara()
   On Error GoTo TrataErro
   
   '--------------------------'
   ' CONSULTA A TABELA DEPARA '
   '--------------------------'
     
    LocalLog "Inicio De Para "
    
    Geral.AgenCob = CLng(Geral.AgenciaVinculo)
    Geral.ContaCob = CDbl(Geral.ContaVinculo)
    
    Geral.hsSQLa = "exec depara "

    Geral.hsSQLa = Geral.hsSQLa & "  '" & Geral.TipoConta & "'"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.AgenCob
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.ContaCob
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", ' '"
      
    LocalLog "DEPARA - SP " & Geral.hsSQLa
    Set Geral.rst = UBBQuery.ExecuteSQL(Geral.hsSQLa)
    LocalLog "Retorno sp-Depara: " & Format(Geral.rst(0), "00")
   
    If (Val(Geral.rst(0)) = 0) Then
      'Se ag/cta for migrada, pega nova ag/cta para pesquisa
       If (Val(Geral.rst(1)) = 1) Then
            Geral.AgenCob = Geral.rst(2)
            Geral.ContaCob = Geral.rst(3)
            LocalLog "Nova ag/conta depara: " & Format(Geral.AgenCob, "0000") & "-" & Format(Geral.ContaCob, "00000000")
       End If
    Else
        Geral.PreparouLog = 1
        Geral.CodOcorrencia = 105
        Geral.Transacao = ""
        LogOcorrencia
    End If
        
    Exit Sub

TrataErro:

Screen.MousePointer = 0
    Select Case TratamentoErro("Falha no módulo: [Consulta Tabela De-Para] .", Err)
        Case eSair
            End
        Case eRepetir
            Resume
        Case eContinuar
            Resume Next
    End Select
End Sub
Sub LogOcorrencia(Optional ByVal pCobrancaIddocto As Double)

   '==========================================================='
   ' TRANSAÇÃO IKRO - ocorrencia malote empresa/caixa expresso '
   '==========================================================='

   On Error GoTo TrataErro
   
    Dim Data        As Long
    Dim Ano         As String
    Dim AposPosi    As Integer
    Dim RstUBB      As Recordset
    Dim RstMDI      As Recordset
   
    Geral.Capa = GetCapa(Geral.idEnvMal)
    Geral.CodTransacao = "IKRO"
    Geral.Evento = 705
    
    Ano = Format(Mid(Parametros.DataServer, 5, 2), "00")
    If (Ano >= "00") And (Ano <= "51") Then
        Data = Mid(Parametros.DataServer, 1, 4) & "20" & Mid(Parametros.DataServer, 5, 2)   'data no formato DDMMAAAA
    ElseIf (Ano > "51") And (Ano <= "99") Then
        Data = Mid(Parametros.DataServer, 1, 4) & "19" & Mid(Parametros.DataServer, 5, 2)   'data no formato DDMMAAAA
    End If
    
    Parametros.AgenciaSatelite = Geral.rstCapa!agorig
    Geral.Hora = Format(Now, "HHMM")
   
    Geral.hsSQLa = "exec tareocor "
   
   'Geral.ValorTrans = formataValor(Geral.rstDoctos!Valor)
    Geral.ValorTrans = 0

    LogAberturaCaixa
       
   'monta header
    Geral.hsSQLa = Geral.hsSQLa & "'" & Geral.CodTransacao & "'"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.VersaoAtual
    Geral.hsSQLa = Geral.hsSQLa & ", " & Parametros.AgenciaCentral
    Geral.hsSQLa = Geral.hsSQLa & ", " & Parametros.AgenciaSatelite
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 1"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.Caixa
    Geral.hsSQLa = Geral.hsSQLa & ", 1"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.NSU1
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.Hora
    Geral.hsSQLa = Geral.hsSQLa & ", ' '"
    
    If Geral.idEnvMal = "E" Then
       Geral.hsSQLa = Geral.hsSQLa & ", 6"
    Else
       Geral.hsSQLa = Geral.hsSQLa & ", 7"
    End If
    
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.Evento
    Geral.hsSQLa = Geral.hsSQLa & ", '" & Geral.Capa & "'"
    Geral.hsSQLa = Geral.hsSQLa & ", '" & Caixa.CIF
    Geral.hsSQLa = Geral.hsSQLa & "', '" & Caixa.SDV

   'monta parte variavel
    Geral.hsSQLa = Geral.hsSQLa & "', " & Geral.rstCapa!agorig
    Geral.hsSQLa = Geral.hsSQLa & ", " & Data
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.CodOcorrencia

    If pCobrancaIddocto <> Empty Then
        Set RstMDI = MDIQuery.getComplOcorrencia(Geral.DataProcessamento, pCobrancaIddocto)
    Else
        Set RstMDI = MDIQuery.getComplOcorrencia(Geral.DataProcessamento, Geral.rstDoctos!iddocto)
    End If
    
    If RstMDI.EOF Then
        Geral.Transacao = Replace(Geral.Transacao, "'", " ", 1, Len(Geral.Transacao), vbTextCompare)
    Else
        Geral.Transacao = RstMDI!Descricao
    End If

    Geral.hsSQLa = Geral.hsSQLa & ", '" & Geral.Transacao & "'"
   
    LocalLog "Ocorrencia IK20 - SP " & Geral.hsSQLa
    Set RstUBB = UBBQuery.ExecuteSQL(Geral.hsSQLa)
    LocalLog "Retorno sp-Ocorrencia: " & Format(RstUBB(0), "00")
   
    If (Val(RstUBB(0)) <> 0) Then
       
        If (Val(RstUBB(0)) = 7) Then
            MsgBox "Atenção! Não foi possível enviar as transações deste malote, pois o caixa está FECHADO. Saia do sistema e entre novamente para executar a rotina de Abertura de Caixa.", vbOKOnly + vbCritical, "Atenção"
        Else
            MsgBox "ATENÇÃO ! Ocorreu o erro " & Str(RstUBB(0)) & " procedure de ocorrência. Tecle <enter> para continuar ", vbOKOnly + vbCritical, "Atenção"
        End If
        
    End If
    
    Exit Sub

TrataErro:

    Select Case TratamentoErro("Falha no módulo: [Gravação de Ocorrência] .", Err)
        Case eSair
            End
        Case eRepetir
            Resume
        Case eContinuar
            Resume Next
    End Select

End Sub
Sub LogAjusteCredito()

   On Error GoTo TrataErro
       
    Dim RstMDI  As Recordset
    
    Geral.Capa = GetCapa(Geral.idEnvMal)

   'envia log de ocorrencia para o ajuste de credito
    Set RstMDI = MDIQuery.getTipoPagto(Geral.DataProcessamento, _
                                       Geral.rstCapa!idcapa, _
                                       Geral.rstDoctos!Vinculo)

    If RstMDI!TipoDocto = 4 Then
       Geral.CodOcorrencia = 403  'adcc
    Else
       Geral.CodOcorrencia = 204  'cheque
    End If
         
    Geral.Transacao = ""
    LogOcorrencia
   
   '================================================='
   ' TRANSAÇÃO 34(nosso número) - credito automatico '
   '================================================='
        
    LocalLog "Inicio geracao log de ajuste de credito"
     
   'variaveis do header
    Geral.CodTransacao = "0015"
    Geral.Evento = 581
    
    Geral.TipoTransacao = 2
    Geral.Capa = GetCapa(Geral.idEnvMal)
    Geral.IndTransac = "O"

   'aviso credito em conta corrente
    Geral.TipoConta = "C"

    Geral.IdentDep = 0
    Geral.AgenciaVinculo = Geral.rstDocto!Agencia
    Geral.ContaVinculo = Geral.rstDocto!Conta

    Set RstMDI = Nothing
    Set RstMDI = MDIQuery.GetBancoOrigemDoAjuste(Geral.DataProcessamento, _
                                                 Geral.rstCapa!idcapa, _
                                                 Geral.rstDoctos!Vinculo)
    If RstMDI!Banco = "230" And _
       Val(Format(Geral.rstDocto!Agencia, "0000") & _
       Format(Geral.rstDocto!Conta, "0000000")) <> Geral.rstCapa!Num_malote Then

      'pesquisa tabela depara 230x409
       If Not Depara230() Then
          Exit Sub
       End If
    Else
      'pesquisa tabela depara 409x409 (antiga)
       DePara
    End If
    
    If (Geral.PreparouLog = 1) Then
        Exit Sub
    End If
    
   'formataValor valor com . (ex 12.00). O valor tem que ter 12 posicoes
    Geral.ValorTrans = formataValor(Geral.rstDoctos!Valor)

   'stored procedure do aviso credito
    
    Geral.hsSQLa = "exec avintpar "
    
   'monta header
    MontaHeader
      
   'monta parte variavel
    Geral.hsSQLa = Geral.hsSQLa & ", 2"
    Geral.hsSQLa = Geral.hsSQLa & ", 1566"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.AgenCob
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.ContaCob
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorTrans)
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Val(Parametros.DataServer)
    Geral.hsSQLa = Geral.hsSQLa & ", ' '"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 1"
    Geral.hsSQLa = Geral.hsSQLa & ", ' '"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
                
    LocalLog "Ajuste de Credito - SP " & Geral.hsSQLa
    
    Exit Sub

TrataErro:

Screen.MousePointer = 0
    Select Case TratamentoErro("Falha no módulo: [Ajuste de Crédito] .", Err)
        Case eSair
            End
        Case eRepetir
            Resume
        Case eContinuar
            Resume Next
    End Select
End Sub
Sub LogAjusteDeposito()
On Error GoTo TrataErro

    Dim RstMDI      As Recordset
    Dim spRetorno   As Integer
    
    Set RstMDI = MDIQuery.getOcorrenciasDeposito(Geral.DataProcessamento, _
                                                 Geral.rstCapa!idcapa, _
                                                 Geral.rstDoctos!Vinculo)
           
    If Not RstMDI.EOF Then
      'envia log de ocorrencia para o ajuste do deposito
       If (Geral.rstDoctos!TipoDocto = 33) Then
           Geral.CodOcorrencia = 121        'deposito a maior
          LocalLog "Gravacao Ik20 (ocorrencia 121) - Capa de deposito maior do que cheques."
       Else
          Geral.CodOcorrencia = 122        'deposito a menor
          LocalLog "Gravacao Ik20 (ocorrencia 122) - Capa de deposito menor do que cheques."
       End If
       
       MDIQuery.updOcorrenciaDocumento Geral.DataProcessamento, _
                                       Geral.rstDoctos!iddocto, _
                                       Geral.CodOcorrencia
                       
       Geral.Transacao = ""
       LogOcorrencia
       
       spRetorno = MDIQuery.updDoctoTransmitido(Geral.DataProcessamento, _
                                                    Geral.rstCapa!idcapa, _
                                                    Geral.rstDoctos!iddocto, _
                                                    Caixa.NSU1, _
                                                    Caixa.Caixa, "N")

       If spRetorno <> 0 Then
          MsgBox "211. ATENÇÃO! Documento - cheque UBB saque interagência - já enviado Log, não foi atualizado no SQL. Tecle <enter> para continuar.", vbOKOnly + vbCritical, "Atenção"
       End If
       
    Else
    
       MDIQuery.updStatusDocto Geral.DataProcessamento, Geral.rstDoctos!iddocto, "D"
    
    End If
    
    Exit Sub
    
TrataErro:

    Screen.MousePointer = 0
    
    Select Case TratamentoErro("Não foi possível gerar a informação do ajuste de depósito na base de dados do MDI-Ubb.", Err)
        Case eSair
            End
        Case eRepetir
            Resume
        Case eContinuar
            Resume Next
    End Select

End Sub
Sub DevolveDocumentos(Optional rst As Recordset = Nothing)
On Error GoTo TrataErro

    Dim spRetorno       As Integer
    
    If rst Is Nothing Then
        Set rst = Geral.rst
    End If
    
   'O erro 115 é de gravaçao do registro. Neste caso deve-se regerar o mesmo
    If Val(rst(0).Value) = 115 Then
        Exit Sub
    End If
    
   'Mis generico
    If Val(rst(0).Value) = 480 Then
        Geral.CodOcorrencia = 999
        Geral.RetTransacao = 94
    End If
    
    LocalLog "Devolução de Doctos: " & Geral.rstDoctos!Nome & " - Retorno Procedure: " & Geral.rst(0).Value
    
   'exceto para documento devolvido por (mis = 94 ou triggererror = 90)
    If Not (Geral.CodOcorrencia = 999 And (Geral.RetTransacao = 90 Or Geral.RetTransacao = 94 Or Geral.RetTransacao = 91)) Then
        
        Select Case Geral.rstDoctos!TipoDocto
            Case 2
            
                '--------------------------'
                ' DEPOSITO CONTA CORRENTE  '
                '--------------------------'
    
                Select Case Val(rst(0).Value)
                    Case 0
                        Geral.CodOcorrencia = 999
                        Geral.RetTransacao = 51
                    Case 7
                        Geral.CodOcorrencia = 999
                        Geral.RetTransacao = rst(0).Value
                    Case 143
                        Geral.CodOcorrencia = 999
                        Geral.RetTransacao = 44
                    Case 214
                        If Mid(rst(15), 1, 1) = "1" Then
                            Geral.CodOcorrencia = 110
                            Geral.Transacao = ""
                            LogOcorrencia
                        ElseIf Mid(rst(15), 2, 1) = "1" Then
                            Geral.CodOcorrencia = 115
                            Geral.Transacao = ""
                            LogOcorrencia
                        ElseIf Mid(rst(15), 3, 1) = "1" Then
                            Geral.CodOcorrencia = 112
                            Geral.Transacao = ""
                            LogOcorrencia
                        ElseIf Mid(rst(15), 4, 1) = "1" Then
                            Geral.CodOcorrencia = 111
                            Geral.Transacao = ""
                            LogOcorrencia
                        End If
                    Case 234
                        'Ficha de deposito já utilizada, a instrucao e p/ trocar a ficha, mais ??  aqui eu estamos deletando
                        Geral.CodOcorrencia = 999
                        Geral.RetTransacao = 45
                    Case 243
                        ' Ficha de deposito com restritivo  2025, tratamento efetuado em 14/08/2000
                        Geral.CodOcorrencia = 111
                        Geral.Transacao = ""
                        LogOcorrencia
                    Case 273
                        Geral.CodOcorrencia = 999
                        Geral.RetTransacao = 48
                    Case Else
                        Geral.CodOcorrencia = 999
                        If Geral.PreparouLog = 2 Then
                           Geral.RetTransacao = 53
                        Else
                           Geral.RetTransacao = rst(0).Value
                        End If
                End Select
            
            Case 3
                '-------------------'
                ' DEPOSITO POUPANCA '
                '-------------------'
             
                Select Case Val(rst(0).Value)
                    Case 7
                        Geral.CodOcorrencia = 999
                        Geral.RetTransacao = rst(0).Value
                    Case 143
                        Geral.CodOcorrencia = 999
                        Geral.RetTransacao = 44
                    Case 234
                       'Ficha de deposito já utilizada, a instrucao e p/ trocar a ficha, mais ??  aqui eu estamos deletando
                        Geral.CodOcorrencia = 999
                        Geral.RetTransacao = 45
                     Case 274
                        Geral.CodOcorrencia = 999
                        Geral.RetTransacao = 46
                     Case Else
                        Geral.CodOcorrencia = 999
                        If Geral.PreparouLog = 2 Then
                           Geral.RetTransacao = 53
                        Else
                           Geral.RetTransacao = rst(0).Value
                        End If
                End Select
              
    
            Case 4
            
                 '-------------------'
                 '       ADCC        '
                 '-------------------'
                
                'tratamento das ocorrencias
                 If Val(rst(0).Value) = 7 Then
                    Geral.CodOcorrencia = 999
                    Geral.RetTransacao = rst(0).Value
                 ElseIf Val(rst(0).Value) = 143 Then
                    Geral.CodOcorrencia = 999
                    Geral.RetTransacao = 44
                 ElseIf Val(rst(0).Value) = 273 Then
                    Geral.CodOcorrencia = 999
                    Geral.RetTransacao = 48
                
                 Else
                 
                    If Geral.CodOcorrencia = 0 Then
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 79
                    End If
       
                 End If
                 
            Case 5
                '-------------------'
                ' SAQUE CHEQUE UBB  '
                '-------------------'
           
                'se cheque UBB está sendo devolvido devido a arrecadação não conveniada,
                'não gera ocorrência aqui, pois já foi gerada na função Log_102/001/025.
                
                 If Val(rst(0).Value) = 278 Then
                    Geral.CodOcorrencia = 999
                    Geral.RetTransacao = 43
                 Else
                    If Val(rst(0).Value) = 88 Then
                       LogChequeCompensacao
                       AlteraParaCompensacao
                       Exit Sub
                    ElseIf Val(rst(0).Value) = 7 Then
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = rst(0).Value
                    Else
                       'tratamento das ocorrencias
                        If Geral.CodOcorrencia = 0 Then
                           Geral.CodOcorrencia = 999
                           Geral.RetTransacao = 79
                        End If
                    End If
                 End If
            Case 6
            
                If Val(rst(0).Value) = 7 Then
                    Geral.CodOcorrencia = 999
                    Geral.RetTransacao = rst(0).Value
                ElseIf Geral.CodOcorrencia = 0 Then
                    Geral.CodOcorrencia = 999
                    Geral.RetTransacao = 99
                End If
            
    
            Case 11
                '----------'
                ' INSS(11) '
                '----------'
                If Val(rst(0).Value) = 7 Then
                    Geral.CodOcorrencia = 999
                    Geral.RetTransacao = rst(0).Value
                Else
                    Geral.CodOcorrencia = 206        'nao conveniado
                    Geral.Transacao = ""
                    LogOcorrencia
                End If
        
            Case 15 To 19
                If Val(rst(0).Value) = 7 Then
                    Geral.CodOcorrencia = 999
                    Geral.RetTransacao = rst(0).Value
                    
                ElseIf (Geral.CodOcorrencia <> 999 And Geral.RetTransacao <> 93) Then
                
                   '------------------------------------------------------'
                   ' DARM(15), DARF_P(16), DARF_S(17), GARE(18), GRPS(19) '
                   '------------------------------------------------------'
                    Geral.CodOcorrencia = 206        'nao conveniado
                    Geral.Transacao = ""
                    LogOcorrencia
                End If
    
    
            Case 20 To 26
                '----------------'
                ' CONCESSIONARIA '
                '----------------'
                If Val(rst(0).Value) = 132 Or Val(rst(0).Value) = 277 Then
                    Geral.CodOcorrencia = 999
                    Geral.RetTransacao = 41
                ElseIf Val(rst(0).Value) = 7 Then
                    Geral.CodOcorrencia = 999
                    Geral.RetTransacao = rst(0).Value
                Else
                    Geral.CodOcorrencia = 999
                    Geral.RetTransacao = 99
                End If
            Case 27
            
                '------------------------'
                ' ARREC.CONVENCIONAL(27) '
                '------------------------'
                If Val(rst(0).Value) = 7 Then
                    Geral.CodOcorrencia = 999
                    Geral.RetTransacao = rst(0).Value
                Else
                    Geral.CodOcorrencia = 206        'nao conveniado
                    Geral.Transacao = ""
                    LogOcorrencia
                End If
          
            Case 31
                '---------------------------------'
                ' FICHA COMPENSACAO OUTROS BANCOS '
                '---------------------------------'
                If Val(rst(0).Value) = 314 Then
                    Geral.CodOcorrencia = 999
                    Geral.RetTransacao = 49
                ElseIf Val(rst(0).Value) = 7 Then
                    Geral.CodOcorrencia = 999
                    Geral.RetTransacao = rst(0).Value
                Else
                    Geral.CodOcorrencia = 999
                    Geral.RetTransacao = 99
                End If
                
            Case 32, 33, 34, 36, 38
                If Val(rst(0).Value) = 143 Then
                    Geral.CodOcorrencia = 999
                    Geral.RetTransacao = 44
                ElseIf Val(rst(0).Value) = 7 Then
                    Geral.CodOcorrencia = 999
                    Geral.RetTransacao = rst(0).Value
                Else
                    Geral.CodOcorrencia = 999
                    Geral.RetTransacao = 99
                End If
        
            Case 37
            
                If Val(rst(0).Value) = 143 Then
                    Geral.CodOcorrencia = 999
                    Geral.RetTransacao = 44
                ElseIf Val(rst(0).Value) = 7 Then
                    Geral.CodOcorrencia = 999
                    Geral.RetTransacao = rst(0).Value
                Else
                    If Trim(Geral.CodOcorrencia) = "" Then
                        Geral.CodOcorrencia = 999
                        Geral.RetTransacao = 99
                    End If
                End If
               
               spRetorno = MDIQuery.updStatusChqOCT(Geral.DataProcessamento, Geral.rstCapa!idcapa, Geral.rstDoctos!Vinculo)
                   
        End Select
    End If
    
   'Mandar capa para CSP
    Geral.PreparouLog = 4
     
   'Exclui docto que gerou erro
    LocalLog "ID-Documento Rejeitado: " & Trim(Geral.rstDoctos!iddocto) & " - Ocorr/RetTransacao: " & Trim(Geral.CodOcorrencia) & " / " & Trim(Geral.RetTransacao)
     
    If Not DbRejeitaDocto(Geral.rstCapa!idcapa, Geral.rstDoctos!iddocto) Then
        Screen.MousePointer = 0
        MsgBox "Falha na exclusão de Documento. Tecle <enter> para continuar.", vbOKOnly + vbCritical, "Atenção"
    End If
    
    MDIQuery.insLog Geral.DataProcessamento, Geral.rstCapa!idcapa, Geral.rstDoctos!iddocto, Caixa.UsuarioAtual, "123"
    
    Exit Sub
    
TrataErro:

    Screen.MousePointer = 0
    
    Select Case TratamentoErro("Não foi possível fazer a exclusão dos documentos rejeitados pelo Servidor da Agência.", Err)
        Case eSair
            End
        Case eRepetir
            Resume
        Case eContinuar
            Resume Next
    End Select

End Sub
Sub LogChequeCompensacao()
   
On Error GoTo TrataErro

    Dim spRetorno       As Integer
    
    Geral.Vinculo = Geral.rstDoctos!Vinculo
    
   '==================================================================================
   ' TRANSAÇÃO 05(nosso numero) - Cheques para Compensação (Cheque Unibanco Cruzado) '
   '==================================================================================
   
   'variaveis do header
    Geral.CodTransacao = "0025"
    Geral.Evento = 642
    Geral.TipoTransacao = 1
    Geral.Capa = GetCapa(Geral.idEnvMal)
    Geral.IndTransac = " "
    Geral.ValorTrans = formataValor(Geral.rstDoctos!Valor)
    Parametros.AgenciaSatelite = Geral.rstCapa!agorig
    Geral.Hora = Format(Now, "HHMM")
            
    Geral.hsSQLa = "exec pagcheq "
    
   'O caixa só será aberto qdo estacao local com caixa fechado
    LogAberturaCaixa
    
    Geral.hsSQLa = Geral.hsSQLa & "  '" & Geral.CodTransacao & "'"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.VersaoAtual
    Geral.hsSQLa = Geral.hsSQLa & ", " & Parametros.AgenciaCentral
    Geral.hsSQLa = Geral.hsSQLa & ", " & Parametros.AgenciaSatelite
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.TipoTransacao
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.Caixa
    Geral.hsSQLa = Geral.hsSQLa & ", 1"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.NSU1
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 1"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.Hora
    Geral.hsSQLa = Geral.hsSQLa & ", '" & Geral.IndTransac & "'"
    
    If Geral.idEnvMal = "E" Then
       Geral.hsSQLa = Geral.hsSQLa & ", 6"
    Else
       Geral.hsSQLa = Geral.hsSQLa & ", 7"
    End If
    
    Geral.hsSQLa = Geral.hsSQLa & ", 2"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.Evento
    Geral.hsSQLa = Geral.hsSQLa & ", '" & Geral.Capa
    
    Geral.hsSQLa = Geral.hsSQLa & "', '" & Caixa.CIF
    Geral.hsSQLa = Geral.hsSQLa & "', '" & Caixa.SDV

   'monta parte variavel
    Geral.hsSQLa = Geral.hsSQLa & "', " & CVar(Geral.ValorTrans)
    Geral.hsSQLa = Geral.hsSQLa & ", '" & Geral.rstDoctos!leitura & "'"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", ' '"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    
    If GetDocumentoTransmitido(EnumOutros) Then
        Exit Sub
    End If
    
   'gravar o nsu desta transação antes de envia-la para o UBB-NT
    spRetorno = MDIQuery.updNsuDocto(Geral.DataProcessamento, _
                                     Geral.rstCapa!idcapa, _
                                     Geral.rstDoctos!iddocto, _
                                     Caixa.NSU1, _
                                     Caixa.Caixa)
    If spRetorno <> 0 Then
        MsgBox "5595. ATENÇÃO! Saque Local a ser enviado Log, não atualizado o NSU. ", vbOKOnly + vbCritical, "Atenção"
        Exit Sub
    End If
   
    LocalLog "Cheque Compensacao - SP " & Geral.hsSQLa
    Set Geral.rst = UBBQuery.ExecuteSQL(Geral.hsSQLa)
    LocalLog "Retorno sp_ChequeCompensacao: " & Format(Geral.rst(0), "00")
   
    If (Val(Geral.rst(0)) = 0) Then
        Geral.GereiLog = 1
        
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''
       ' Procedure para atualizar o cheque como já enviado e'
       ' setar flag de cheque compensado.                   '
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''
       
        spRetorno = MDIQuery.updDoctoTransmitido(Geral.DataProcessamento, _
                                                 Geral.rstCapa!idcapa, _
                                                 Geral.rstDoctos!iddocto, _
                                                 Caixa.NSU1, _
                                                 Caixa.Caixa, "S")
                                    
        If spRetorno <> 0 Then
            MsgBox "785. ATENÇÃO! Documento - cheque UBB compensação - já enviado Log, não foi atualizado no SQL. Tecle <enter> para continuar.", vbOKOnly + vbCritical, "Atenção"
        End If
        
        Geral.PreparouLog = 2
            
    Else
              
        Call DevolveDocumentos
        
    End If
    
    Exit Sub
    
TrataErro:

    Select Case TratamentoErro("Não foi possível finalizar a transação de Cheque Compensação.", Err)
        Case eSair
            End
        Case eRepetir
            Resume
        Case eContinuar
            Resume Next
    End Select

End Sub
Sub LogFechamentoCaixa(ByVal pUnidade As String)
   
On Error GoTo TrataErro

    Dim RstUBB          As Recordset
    Dim spRetorno       As Integer
    Dim ValorDiferenca  As String
    
    If pUnidade <> "T" Then
       
       If Not Geral.RecebendoCapa Then
          spRetorno = MDIQuery.updHoraUltimaTransmissao(Geral.DataProcessamento, Caixa.Caixa)
             
          If spRetorno <> 0 Then
             MsgBox "Ocorreu algum erro na gravação do horário da última transmissão. Tecle <enter> para continuar.", vbOKOnly + vbCritical, "Atenção"
          End If
          
       End If
              
       Parametros.AgenciaSatelite = Geral.rstCapa!agorig
      
       Geral.Hora = Format$(Now, "HHMM")
       Geral.ValorTrans = 0
   
      'verifica qual caixa está aberto para fazer fechamento
       Geral.hsSQLb = "Exec Fechamdi "
       
      'monta header
       MontaFechamento (0) 'Fecha sem diferenca
  
       LocalLog "Fechamento Normal Caixa MDI - SP " & Geral.hsSQLb
       
       If GetSetting("Robo", "Caixa", "Aberto", 0) = 1 Then
           Set Geral.rst = UBBQuery.ExecuteSQL(Geral.hsSQLb)
           LocalLog "Retorno Fechamento Normal de Caixa MDI - " & Format(Geral.rst(0), "00")
                       
            If CDbl(Geral.rst(0)) = 10 Then
            
                If CDbl(Geral.rst(30)) <> 0 Then 'caixa com diferenca
                
                  'leitura da diferenca do caixa
                   Set RstUBB = UBBQuery.getNsuTabelacaixa(Caixa.Caixa)
                   
                   If RstUBB!tcxfsvlacrc > RstUBB!tcxfsvlacpg Then
                     'Insere Mensagem p/ Usuario
                      spRetorno = MDIQuery.insMensagem(Geral.DataProcessamento, _
                                            MSG_DiferencaFechamentoCx, _
                                            Geral.rstCapa!idcapa, 0, _
                                            Caixa.Caixa)
                                   
                     'MsgBox "Foi constatado diferenca de R$ " & Format(Geral.Rst(30) / 100, "#.00") & "  no fechamento do caixa durante processamento da capa  " & Geral.rstCapa!Capa & " . Tecle <enter> para continuar.", vbOKOnly + vbCritical, "Atenção"
                      ValorDiferenca = formataValor(Geral.rst(30))
                      
                      LocalLog "Foi constatado diferenca de R$ " & Format(Geral.rst(30), "#.00") & "] no fechamento do caixa durante processamento da capa [" & Geral.rstCapa!Capa & "]."
                      
                   Else
                     'Insere Mensagem p/ Usuario
                      spRetorno = MDIQuery.insMensagem(Geral.DataProcessamento, _
                                            MSG_DiferencaFechamentoCx, _
                                            Geral.rstCapa!idcapa, 0, _
                                            Caixa.Caixa)
                                   
                     'MsgBox "Foi constatado diferenca de  R$ -" & Format(Geral.Rst(30) / 100, "#.00") & "  no fechamento do caixa durante processamento da capa  " & Geral.rstCapa!Capa & " . Tecle <enter> para continuar.", vbOKOnly + vbCritical, "Atenção"
                      ValorDiferenca = formataValor(0 - Geral.rst(30))
                      
                      LocalLog "Foi constatado diferenca de [R$ -" & Format(Geral.rst(30) / 100, "000.00") & "] no fechamento do caixa durante processamento da capa [" & Geral.rstCapa!Capa & "]."
                   
                   End If
                   
                  'gravação dos dados da tabela DIFERENCACAIXA
                   MDIQuery.insDiferencaCaixa Geral.DataProcessamento, _
                                              Geral.rstCapa!idcapa, _
                                              Caixa.Caixa, _
                                              Caixa.Estacao, ValorDiferenca
                   
                  'verifica qual caixa está aberto para fazer fechamento
                  
                   Geral.ValorTrans = ValorDiferenca
                   Geral.hsSQLb = "exec fechamdi "
                
                   MontaFechamento (1)   'Força fechamento com diferença
                   LocalLog "Fechamento de Erro Caixa MDI - SP " & Geral.hsSQLb
                   
                   Set Geral.rst = UBBQuery.ExecuteSQL(Geral.hsSQLb)
                   LocalLog "Retorno Fechamento de Erro de Caixa MDI - " & Format(Geral.rst(0), "00")
                   
                End If
                
            ElseIf CDbl(Geral.rst(0)) = 246 Then
                LocalLog "Fechamento Caixa Retorno: " & Format(Geral.rst(0).Value, "000")
                MsgBox "ATENÇÃO ! Caixa Fechado, Retorno " & Str$(Geral.rst(0)) & " no fechamento do caixa. Contate o suporte técnico.", vbOKOnly + vbCritical, "Atenção, Capa: " & Geral.rstCapa!Capa
            Else
                LocalLog "Fechamento Caixa Retorno: " & Format(Geral.rst(0).Value, "000") & " Diferença: " & Str(Geral.rst(30).Value)
                If Val(Geral.rst(0)) <> 0 And Val(Geral.rst(0)) <> 10 And Val(Geral.rst(0)) <> 363 Then
                    MsgBox "ATENÇÃO ! Ocorreu o retorno " & Str$(Geral.rst(0)) & " no fechamento do caixa. Contate o suporte técnico.", vbOKOnly + vbCritical, "Atenção"
                End If
              
           End If
           
           SaveSetting appname:="Robo", section:="Caixa", Key:="Aberto", setting:=0
        
       End If
       
    End If

    If pUnidade = "M" Or pUnidade = "T" Then  'Se for MultiAgência
        '--------------------------------------------------'
        ' Fechamento do caixa no MDI (flag N - não aberto) '
        '--------------------------------------------------'
        spRetorno = MDIQuery.AtualizarCaixaMDI(Geral.DataProcessamento, Caixa.Estacao, Caixa.Caixa, "N")
           
        If spRetorno <> 0 Then
            MsgBox "Ocorreu algum erro com o fechamento do caixa de controle no MDI-Ubb. Tecle <enter> para continuar.", vbOKOnly + vbCritical, "Atenção"
        End If
       
    End If

    Exit Sub

TrataErro:

    Select Case TratamentoErro("Não foi possível finalizar a transação de Fechamento de Caixa.", Err)
        Case eSair
            End
        Case eRepetir
            Resume
        Case eContinuar
    End Select

End Sub
Sub LogFimCiclo()
   On Error GoTo TrataErro
   
   '==============================='
   ' TRANSAÇÃO BHFC - fim de ciclo '
   '==============================='
   
    Parametros.AgenciaSatelite = Geral.rstCapa!agorig
    Geral.Capa = GetCapa(Geral.idEnvMal)
    Geral.Hora = Format(Now, "HHMM")
   
   ''O caixa só será aberto qdo estacao local com caixa fechado
    LogAberturaCaixa (True)
       
   'monta header
    Geral.hsSQLa = "exec fimciclo "
   
    Geral.hsSQLa = Geral.hsSQLa & "  'BHFC'"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.VersaoAtual
    Geral.hsSQLa = Geral.hsSQLa & ", " & Parametros.AgenciaCentral
    Geral.hsSQLa = Geral.hsSQLa & ", " & Parametros.AgenciaSatelite
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 1"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.Caixa
    Geral.hsSQLa = Geral.hsSQLa & ", 1"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.NSU1
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.Hora
    Geral.hsSQLa = Geral.hsSQLa & ", ' '"
        
    If Geral.idEnvMal = "E" Then
       Geral.hsSQLa = Geral.hsSQLa & ", 6"
    Else
       Geral.hsSQLa = Geral.hsSQLa & ", 7"
    End If
    
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 134"
    Geral.hsSQLa = Geral.hsSQLa & ", '" & Geral.Capa & "'"
    Geral.hsSQLa = Geral.hsSQLa & ", '" & Caixa.CIF
    Geral.hsSQLa = Geral.hsSQLa & "', '" & Caixa.SDV & "'"

    LocalLog "Fim de Ciclo - SP " & Geral.hsSQLa
   
   'stored procedure
    Set Geral.rst = UBBQuery.ExecuteSQL(Geral.hsSQLa)
    LocalLog "Retorno sp-Fim de Ciclo - " & Format(Geral.rst(0), "00")
   
    If (Val(Geral.rst(0)) <> 0) Then
        MsgBox "ATENÇÃO ! Ocorreu o retorno " & Str(Geral.rst(0)) & " na procedure de Fim de Ciclo. Tecle <enter> para continuar.", vbOKOnly + vbCritical, "Atenção"
    End If
    
    Exit Sub

TrataErro:

Screen.MousePointer = 0
    Select Case TratamentoErro("Falha no módulo: [Finaliza Ciclo] .", Err)
        Case eSair
            End
        Case eRepetir
            Resume
        Case eContinuar
            Resume Next
    End Select
       
End Sub
Sub LogADCC()
   
On Error GoTo TrataErro

   '====================================='
   ' TRANSAÇÃO 04 (nosso número) - ADCC  '
   '====================================='
   
    Dim RstUBB          As Recordset
    Dim RstMDI          As Recordset
    Dim spRetorno       As Integer
    Dim MsgIda          As String
    Dim MsgRetorno      As String
    Dim HeaderTx        As String
    Dim TamIda          As String
    Dim Funcao          As String * 14
    Dim RetQX           As Integer
    Dim Vez             As Integer
            
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   'verifica se ADCC está pagando alguma cobrança Ubb vencida e se esta pode ser paga '
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
    Geral.Vinculo = Geral.rstDoctos!Vinculo
    
   'variaveis do header
    Geral.CodTransacao = "0F15"
    Geral.Evento = 580
    Geral.TipoTransacao = 1
    Geral.Capa = GetCapa(Geral.idEnvMal)
    Geral.IndTransac = "1"
    Geral.ValorTrans = formataValor(Geral.rstDoctos!Valor)
    Geral.TipoConta = "C"
    Geral.AgenciaVinculo = Geral.rstDocto!Agencia
    Geral.ContaVinculo = Geral.rstDocto!Conta
    
   'pesquisa tabela depara
    DePara
   
    If (Geral.PreparouLog = 1) Then
        Exit Sub
    End If
      
   'stored procedure do SAQUE
    Geral.hsSQLa = "exec avintpar "
      
   'monta header
    MontaHeader
   
   'monta parte variavel da 1ª perna do log
    Geral.hsSQLa = Geral.hsSQLa & ", 2"
    Geral.hsSQLa = Geral.hsSQLa & ", 1841"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.AgenCob
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.ContaCob
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorTrans)
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.rstDoctos!CodCenape
    Geral.hsSQLa = Geral.hsSQLa & ", " & Val(Parametros.DataServer)
    Geral.hsSQLa = Geral.hsSQLa & ", ' '"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 1"
    Geral.hsSQLa = Geral.hsSQLa & ", ' '"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    
   'executa Fracassada (0F15)
    LocalLog "Aut.Debito (0F15) - SP " & Geral.hsSQLa
    Set Geral.rst = UBBQuery.ExecuteSQL(Geral.hsSQLa)
    LocalLog "Retorno saque aut.debito(0F15): " & Format(Geral.rst(0), "00")
    
    If (Val(Geral.rst(0)) = 0) Then
                 
        CalculaNSU
        HeaderTx = "BHS1" & Format(Geral.rst(13), "000000") & _
                    Caixa.VersaoAtual & _
                    Format(Parametros.AgenciaCentral, "0000") & _
                    Format(Parametros.AgenciaSatelite, "0000") & "011" & _
                    Format(Caixa.Caixa, "000") & "1" & _
                    Format(Caixa.NSU, "000000") & "0000000" & _
                    Format(Now, "HHMM") & "110000000002"
        
        MsgIda = HeaderTx & "05" & Format(Geral.AgenCob, "0000") & _
                "0" & Format(Geral.ContaCob, "0000000000000000") & _
                "0000" & Format(Geral.IdentDep, "000000000") & _
                 String(13, "0") & String(5, "0") & _
                 Parametros.DataServer & "00" & _
                 Format(Geral.rstDoctos!Valor * 100, "0000000000000000") & _
                 String(16, "0") & Format(Geral.rstDoctos!Valor * 100, "0000000000000000") & _
                 Format(Now, "HHMM") & "0" & "0000" & "0000000000" & String(4, "0") & _
                 String(37, "0") & String(6, "0") & String(6, "0") & "02"
           
       'Envia 1ª mensagem ao Host
        TamIda = Format(Len(Trim(MsgIda)), "0000")
        MsgRetorno = String(1921, " ")
        Funcao = "1" & TamIda & "1921****"
        
        LocalLog MsgIda
        
       'Envia BHS1
        Call Abrelinha("BHS1")
        RetQX = qxhostnt(Funcao, MsgIda, MsgRetorno)
        Call FechaLinha("BHS1")
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'se retorno = 20,41,71 -> o micro deverá ser reinicializado '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If (RetQX = 20) Or (RetQX = 41) Or (RetQX = 71) Then
             LocalLog "Retorno BHS1: " & Mid(MsgRetorno, 58, 2)
             MsgBox "Atenção. Ocorreu um erro de comunicação com o Servidor da Agência. Reinicialize este equipamento. Retorno DLL = " & Format(RetQX, "00"), vbOKOnly + vbCritical, "Atenção"
             Close #20
             End
        End If
        
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''
       ' Retorno = 21,30,43,47,52,62,80 -> tentar novamente.'
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        If (RetQX = 21) Or (RetQX = 30) Or (RetQX = 43) Or (RetQX = 47) Or _
           (RetQX = 52) Or (RetQX = 62) Or (RetQX = 80) Then
        
            Vez = 1
            Do
               
                Espera (5 * Vez)
                     
               'tentar novamente
                Call Abrelinha("RE-envio BHS1")
                RetQX = qxhostnt(Funcao, MsgIda, MsgRetorno)
                Call FechaLinha("RE-envio BHS1")
                
                Vez = Vez + 1
            
            Loop Until Not ((RetQX = 21) Or (RetQX = 30) Or (RetQX = 43) Or _
                            (RetQX = 47) Or (RetQX = 52) Or (RetQX = 62) Or (RetQX = 80)) And Vez <> 4
        
        End If
        
        LocalLog "Retorno BHS1: " & Format(RetQX, "00") & "MsgRetorno: " & MsgRetorno
        
       'Recebe retorno da Consulta do Host BHS1
        If (RetQX = 0) Then
        
           'Recebe resposta no BHS2, Grava concretizada 6001
            If Mid(MsgRetorno, 58, 2) = "00" Then
                                  
               'variaveis do header
                Geral.CodTransacao = "0015"
                Geral.Evento = 580
                Geral.TipoTransacao = 1
                Geral.Capa = GetCapa(Geral.idEnvMal)
                Geral.IndTransac = " "
                Geral.ValorTrans = formataValor(Geral.rstDoctos!Valor)
                
               'monta header
                Parametros.AgenciaSatelite = Geral.rstCapa!agorig
                
                Geral.Hora = Format(Now, "HHMM")
                     
               'stored procedure do aviso debito
                Geral.hsSQLa = "exec avintpar "
                              
               ''O caixa só será aberto qdo estacao local com caixa fechado
                LogAberturaCaixa
                
                Geral.hsSQLa = Geral.hsSQLa & "  '" & Geral.CodTransacao & "'"
                Geral.hsSQLa = Geral.hsSQLa & ", " & Format(Geral.rst(13), "000000")
                Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.VersaoAtual
                Geral.hsSQLa = Geral.hsSQLa & ", " & Parametros.AgenciaCentral
                Geral.hsSQLa = Geral.hsSQLa & ", " & Parametros.AgenciaSatelite
                Geral.hsSQLa = Geral.hsSQLa & ", 0"
                Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.TipoTransacao
                Geral.hsSQLa = Geral.hsSQLa & ", 3"
                Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.Caixa
                Geral.hsSQLa = Geral.hsSQLa & ", 1"
                Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.NSU1
                Geral.hsSQLa = Geral.hsSQLa & ", 0"
                Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.CapaBack
                Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.Hora
                Geral.hsSQLa = Geral.hsSQLa & ", '" & Geral.IndTransac & "'"
                
                If Geral.idEnvMal = "E" Then
                    Geral.hsSQLa = Geral.hsSQLa & ", 6"
                Else
                    Geral.hsSQLa = Geral.hsSQLa & ", 7"
                End If
                
                Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.TpRep
                Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.Evento
                Geral.hsSQLa = Geral.hsSQLa & ", '" & Geral.Capa & "'"
                Geral.hsSQLa = Geral.hsSQLa & ", '" & Caixa.CIF
                Geral.hsSQLa = Geral.hsSQLa & "', '" & Caixa.SDV

               'monta parte variavel
                Geral.hsSQLa = Geral.hsSQLa & "', 2"
                Geral.hsSQLa = Geral.hsSQLa & ", 1841"
                Geral.hsSQLa = Geral.hsSQLa & ", 0"
                Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.AgenCob
                Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.ContaCob
                Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorTrans)
                Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.rstDoctos!CodCenape
                Geral.hsSQLa = Geral.hsSQLa & ", " & Val(Parametros.DataServer)
                Geral.hsSQLa = Geral.hsSQLa & ", ' '"
                Geral.hsSQLa = Geral.hsSQLa & ", " & Mid(MsgRetorno, 5, 6)
                Geral.hsSQLa = Geral.hsSQLa & ", 1"
                Geral.hsSQLa = Geral.hsSQLa & ", ' '"
                Geral.hsSQLa = Geral.hsSQLa & ", 0"
                
                If GetDocumentoTransmitido(EnumADCC) Then
                    Exit Sub
                End If
                
               'gravar o nsu desta transação antes de envia-la para o UBB-NT
                spRetorno = MDIQuery.updNsuDocto(Geral.DataProcessamento, _
                                                 Geral.rstCapa!idcapa, _
                                                 Geral.rstDoctos!iddocto, _
                                                 Caixa.NSU1, _
                                                 Caixa.Caixa)
                                       
                If spRetorno <> 0 Then
                    MsgBox "5435. ATENÇÃO! Aut.Debito a ser enviado Log, não atualizado o NSU. ", vbOKOnly + vbCritical, "Atenção"
                    Exit Sub
                End If
                
               'Executa concretizada
                LocalLog "Aut. Debito (0015) - SP " & Geral.hsSQLa
                Set Geral.rst = UBBQuery.ExecuteSQL(Geral.hsSQLa)
                LocalLog "Retorno saque aut.debito(0015): " & Format(Geral.rst(0), "00")
                
                If (Val(Geral.rst(0)) = 0) Then
                    
                    CalculaNSU
              
                    HeaderTx = "BHS3" & Format(Geral.rst(13), "000000") & Caixa.VersaoAtual & Format(Parametros.AgenciaCentral, "0000") & Format(Parametros.AgenciaSatelite, "0000") & "011" & Format(Caixa.Caixa, "000") & "1" & Format(Caixa.NSU, "000000") & "000000" & "0" & Format(Now, "HHMM") & "110000000002"
                    MsgIda = HeaderTx & Mid(MsgRetorno, 60, 1) & Mid(MsgRetorno, 61, 6) & Format(Geral.rst(13), "000000") & "1" & Format(Geral.AgenCob, "0000")
                    
                    TamIda = Format(Len(Trim(MsgIda)), "0000")
                    MsgRetorno = String(1921, " ")
                    Funcao = "1" & TamIda & "1921****"
                                  
                    LocalLog MsgIda
              
                   'envia BHS3 (confirmação para o Host)
                    Call Abrelinha("BHS3")
                    RetQX = qxhostnt(Funcao, MsgIda, MsgRetorno)
                    Call FechaLinha("BHS3")
                    
                    If RetQX <> 0 Then
                        Err.Raise 963, App.Title, "Falha na Confirmação de Pagamento (BHS3)"
                    End If
                                         
                    If Not (IsNull(Geral.rstDoctos!RetornoTransacao) And _
                            Geral.rstDoctos!RetornoTransacao = 75) Then
                        
                       'Conta Unibanco reinformada Corretamente
                        MDIQuery.updCancelarRetornoTransacao Geral.DataProcessamento, _
                                                             Geral.rstCapa!idcapa, _
                                                             Geral.rstDoctos!iddocto, _
                                                             Caixa.Caixa
                    End If
                                            
                   ''''''''''''''''''''''''''''''''''''''''''''''''
                   ' Procedure para atualizar o adcc como enviado '
                   ''''''''''''''''''''''''''''''''''''''''''''''''
                   
                    spRetorno = MDIQuery.updDoctoTransmitido(Geral.DataProcessamento, _
                                                             Geral.rstCapa!idcapa, _
                                                             Geral.rstDoctos!iddocto, _
                                                             Caixa.NSU1, _
                                                             Caixa.Caixa, "N")
                                        
                    If spRetorno <> 0 Then
                        MsgBox "590. ATENÇÃO! Documento - débito automático - já enviado Log, não atualizado no SQL. ", vbOKOnly + vbCritical, "Atenção"
                    End If
                                            
                   'libera linha
                    Geral.GereiLog = 1
              
                Else
                    Call DevolveDocumentos
                    Exit Sub
                End If
            
            ElseIf Mid(MsgRetorno, 58, 2) = "03" Then
               'Pagto de Conta/Titulo com cheque Ubb com insuficiencia de saldo
                Geral.CodOcorrencia = 429
                Geral.Transacao = Mid$(MsgRetorno, 84, 31)
                LogOcorrencia
                Call DevolveDocumentos
                
               'Atualiza Tabela ADCC Saldo/Limite/Vinculado
                spRetorno = MDIQuery.updADCCSaldo(Geral.DataProcessamento, _
                                                  Geral.rstDoctos!iddocto, _
                                                  MsgRetorno)
        
                If spRetorno <> 0 Then
                   MsgBox "Ocorreu algum erro na gravação do saldo para ocorrencia de ADCC. Tecle <enter> para continuar.", vbOKOnly + vbCritical, "Atenção"
                End If
                
                Exit Sub
            
            Else
                                
                Select Case UCase(Trim(Mid(MsgRetorno, 84, 31)))
                
                    Case "CONTA UNIBANCO NAO EXISTE"    ' (06) **Conta Unibanco nao existe
                       Geral.CodOcorrencia = 414        ' alteração solicitada por Selma em 30/11/99
                       
                    Case "TIPO CONTA UNIBANCO INVALIDA" ' (06) **Conta Unibanco nao existe
                       Geral.CodOcorrencia = 414
                       
                    Case "BLOQUEIO RESOLUCAO 2025"      ' (10) Bloqueio Resolucao 2025
                       Geral.CodOcorrencia = 420
                       
                    Case "CONTA UNIBANCO INVALIDA"      ' (12) **Conta Unibanco Invalida
                       Geral.CodOcorrencia = 414        ' alteração solicitada por Selma em 30/11/99
                       
                    Case "TRANSACAO CANCELADA"          ' (36) **Transação Cancelada
                       Geral.CodOcorrencia = 0

                    Case "CONTA SEM SALDO"              ' (83) Conta sem Saldo
                       
                      'Atualiza Tabela ADCC Saldo/Limite/Vinculado
                       spRetorno = MDIQuery.updADCCSaldo(Geral.DataProcessamento, _
                                                         Geral.rstDoctos!iddocto, _
                                                         MsgRetorno)
                    
                       If spRetorno <> 0 Then
                          MsgBox "ATENÇÃO !!! (96)Erro na Atualizacao de saldo insuficiente para ADCC. Tecle <enter> para continuar.", vbOKOnly + vbCritical, "Atenção"
                       End If
                    
                       Geral.CodOcorrencia = 429
                       
                    Case "SALDO BLOQUEADO"              ' (84) Saldo Bloqueado
                       Geral.CodOcorrencia = 426
                       
                    Case "CONTA ENCERRADA"              ' (94) Conta Encerrada
                       Geral.CodOcorrencia = 419
                       
                    Case "CONTA BLOQUEADA"              ' (95) Conta Bloqueada
                       Geral.CodOcorrencia = 425
                       
                    Case "RESIDE EXTERIOR"              ' (96) Reside Exterior
                       Geral.CodOcorrencia = 421
                       
                    Case "CONTA PARALISADA"             ' (97) Conta Paralisada
                       Geral.CodOcorrencia = 422
                       
                    Case "ADIANTAMENTO DEPOSITANTE"     ' (99) Adiantamento depositante
                       Geral.CodOcorrencia = 424
                       
                    Case "CONTA INATIVA-BLOQUEADA"       '(01) Conta Bloqueada por Inatividade
                       Geral.CodOcorrencia = 425
                    
                    Case "HORARIO EXPIRADO"       '(01)
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 1
                       
                    Case "PROBLEMAS NA CONSIST. DO CP"       '(01)
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 2
                       
                    Case "CONTA CONTABIL"       '(01)
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 3
                       
                    Case "CONTA UNIBANCO NAO ABERTA"       '(01)
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 5
                       
                    Case "AGENCIA NACIONAL INVALIDA"       '(01)
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 59
                    
                    Case "C/C NACIONAL INVALIDA"       '(01)
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 60
                     
                    Case "EXCEDEU LIMITE"       '(01)
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 62
                     
                     Case "DEP. OUTRA PRACA"       '(01)
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 64
                     
                    Case "CONTA COM RESTRITIVO(S)"       '(01)
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 66
                     
                    Case "CONTATE AGEN TITULAR DA CONTA"       '(01)
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 67
                     
                    Case "VALOR SAQUE RAPIDO EXCEDIDO"       '(01)
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 69
                     
                    Case "DEP INT DISP.AAG"       '(01)
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 71
                     
                    Case "INTERAGENCIA NAO PERMITIDO"       '(01)
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 73
                     
                    Case "SALDO BLOQUEADO"       '(01)
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 74
                       
                    Case Else                           ' nao tratado
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 51
                End Select
                               
                If Geral.CodOcorrencia = 414 And _
                   (IsNull(Geral.rstDoctos!RetornoTransacao) Or _
                    Geral.rstDoctos!RetornoTransacao <> 75) Then
                    
                    Call GaugeTitulo(1)
                                        
                    Geral.RetTransacao = 75
                    Geral.CodOcorrencia = "0"
            
                    LocalLog "Agencia/Conta Invalida sendo enviada para correcao de agencia/conta"
                    Call DbRejeitaDocto(Geral.rstCapa!idcapa, Geral.rstDoctos!iddocto, ST_DoctoCorrecaoAgConta)
                    
                   'Capa para correcao de Agencia/Conta
                    Geral.PreparouLog = 5
                    Espera (0.5)
                Else
                                  
                   'grava ocorrencia
                    If Geral.CodOcorrencia > 0 And Geral.CodOcorrencia < 998 Then
                        Geral.Transacao = Mid$(MsgRetorno, 84, 31)
                        LogOcorrencia
                    End If
                    
                    LocalLog "Agencia/Conta Reinformada Invalida sendo enviada para CSP"
                    
                    Call DevolveDocumentos
                
                End If
                
                Exit Sub
        
            End If
        Else
            
           'Falhou BHS1
            LocalLog "Retorno da função da DLL para o envio da BHS1-Retorno: " & Str(RetQX) & " - Msgretorno: " & MsgRetorno

            Select Case RetQX
                Case 21, 30, 43, 47, 52, 62, 80, 33, 36, 42, 48
                    LogAjusteDebitoADCC
                Case Else
                    Geral.CodOcorrencia = 999
                    Geral.RetTransacao = 51
                    Call DevolveDocumentos
            End Select
            
            Exit Sub
            
        End If
        
    Else
    
       'Falhou Fracassada devolve com retorno padrao da mesma
        Call DevolveDocumentos
        
    End If
   
    Exit Sub
    
TrataErro:
    LocalLog "Falha no ADCC " & Err.Description

    If Err.Number = 964 Or Err.Number = 965 Then
       'Erro na Abertura/Fechamento de Linha
        Geral.CodOcorrencia = 999
        Geral.RetTransacao = 98
        Call DevolveDocumentos
        Exit Sub
    ElseIf Err.Number = 963 Then
       'Erro Subida da confirmaçao do Saque (BHS3)
        Geral.CodOcorrencia = 999
        Geral.RetTransacao = 97
        Call DevolveDocumentos
        Exit Sub
    End If
    
    Select Case TratamentoErro("Falha no ADCC.", Err, eDoctoSubidaLog)
        Case eSair
            End
        Case eRepetir
            Resume
        Case eContinuar
            Resume Next
        Case eFinalizar
            Exit Sub
    End Select

End Sub
Sub MontaHeader()
    
    Parametros.AgenciaSatelite = Geral.rstCapa!agorig
    Geral.Hora = Format(Now, "HHMM")
    
   'O caixa só será aberto qdo estacao local com caixa fechado
    LogAberturaCaixa
        
    Geral.hsSQLa = Geral.hsSQLa & "  '" & Geral.CodTransacao & "'"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.VersaoAtual
    Geral.hsSQLa = Geral.hsSQLa & ", " & Parametros.AgenciaCentral
    Geral.hsSQLa = Geral.hsSQLa & ", " & Parametros.AgenciaSatelite
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.TipoTransacao
    Geral.hsSQLa = Geral.hsSQLa & ", 3"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.Caixa
    Geral.hsSQLa = Geral.hsSQLa & ", 1"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.NSU1
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.CapaBack
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.Hora
    Geral.hsSQLa = Geral.hsSQLa & ", '" & Geral.IndTransac & "'"
    
    If InStr(1, "EF", Geral.idEnvMal, 1) <> 0 Then
       Geral.hsSQLa = Geral.hsSQLa & ", 6"
    ElseIf Geral.idEnvMal = "M" Then
       Geral.hsSQLa = Geral.hsSQLa & ", 7"
    End If
    
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.TpRep
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.Evento
    Geral.hsSQLa = Geral.hsSQLa & ", '" & Geral.Capa & "'"
    Geral.hsSQLa = Geral.hsSQLa & ", '" & Caixa.CIF
    Geral.hsSQLa = Geral.hsSQLa & "', '" & Caixa.SDV & "'"

End Sub
Sub MontaFechamento(ByVal pTipo As Integer)
    
   'leitura do número do terminal
    CalculaNSU
    Caixa.NSU1 = Caixa.NSU
    
    CalculaNSU
    Caixa.NSU2 = Caixa.NSU
    
   'monta header
    Geral.hsSQLb = Geral.hsSQLb & " '0031'"
    Geral.hsSQLb = Geral.hsSQLb & ", 0"
    Geral.hsSQLb = Geral.hsSQLb & ", " & Caixa.VersaoAtual
    Geral.hsSQLb = Geral.hsSQLb & ", " & Parametros.AgenciaCentral
    Geral.hsSQLb = Geral.hsSQLb & ", " & Parametros.AgenciaSatelite
    Geral.hsSQLb = Geral.hsSQLb & ", " & pTipo
    Geral.hsSQLb = Geral.hsSQLb & ", 4"
    Geral.hsSQLb = Geral.hsSQLb & ", 0"
    Geral.hsSQLb = Geral.hsSQLb & ", " & Caixa.Caixa
    Geral.hsSQLb = Geral.hsSQLb & ", 1"
    Geral.hsSQLb = Geral.hsSQLb & ", " & Caixa.NSU1
    Geral.hsSQLb = Geral.hsSQLb & ", 0"
    Geral.hsSQLb = Geral.hsSQLb & ", 1"
    Geral.hsSQLb = Geral.hsSQLb & ", " & Geral.Hora
    Geral.hsSQLb = Geral.hsSQLb & ", ' '"
    
    If Geral.idEnvMal = "E" Then
       Geral.hsSQLb = Geral.hsSQLb & ", 6"
    Else
       Geral.hsSQLb = Geral.hsSQLb & ", 7"
    End If
    
    Geral.hsSQLb = Geral.hsSQLb & ", 0"
    Geral.hsSQLb = Geral.hsSQLb & ", 129"
    Geral.hsSQLb = Geral.hsSQLb & ", ' '"
    
    Geral.hsSQLb = Geral.hsSQLb & ", '" & Caixa.CIF
    Geral.hsSQLb = Geral.hsSQLb & "', '" & Caixa.SDV
         
   'monta parte variavel
    Geral.hsSQLb = Geral.hsSQLb & "', 0"
    Geral.hsSQLb = Geral.hsSQLb & ", " & Caixa.Caixa
    Geral.hsSQLb = Geral.hsSQLb & ", " & Caixa.NSU2

End Sub
Sub LogConcessionarias()

    Dim Valor As Long
   
    ' .......................................................
    ' ARRECADAÇÃO:
    ' .......................................................
    ' TRANSAÇÃO 20 (nosso número) - Recebimento de AGUA
    ' TRANSAÇÃO 21 (nosso número) - Recebimento de GAS
    ' TRANSAÇÃO 22 (nosso número) - Recebimento de LUZ
    ' TRANSAÇÃO 23 (nosso número) - Recebimento de TELEFONE
    ' TRANSAÇÃO 24 (nosso número) - tributo municipal
    ' TRANSAÇÃO 25 (nosso número) - tributo estadual
    ' TRANSAÇÃO 26 (nosso número) - tributo federal
    ' .......................................................
   
   'variaveis do header
    Geral.CodTransacao = "0020"
    Geral.Evento = 116
    Geral.TipoTransacao = 2
    Geral.Capa = GetCapa(Geral.idEnvMal)
    Geral.IndTransac = " "

    Geral.ValorTrans = formataValor(Geral.rstDoctos!Valor)

    Geral.hsSQLa = "exec arrecad "
    
   'monta header
    MontaHeader
   
    MontaComplemento
   
   'monta parte variavel
    Geral.hsSQLa = Geral.hsSQLa & ", '" & Geral.rstDoctos!leitura & "'"
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorTrans)
    
   'verifica se codigo de barras foi corrigido
    If Geral.rstDoctos!CodBarComplem = "N" Then
        Geral.hsSQLa = Geral.hsSQLa & ", 1"
    Else
        Geral.hsSQLa = Geral.hsSQLa & ", 2"
    End If
    
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    
    MontaComplementoVariavel

    LocalLog "Concessionaria - SP " & Geral.hsSQLa
End Sub
Sub LogGRPS()

    ' .................................
    ' TRANSAÇÃO 19(nosso número) - GRPS
    ' .................................
   
    Dim ValorLiquido As Double
   
   'variaveis do header
    Geral.CodTransacao = "0089"
    Geral.Evento = 22
    
    Geral.TipoTransacao = 2
    Geral.Capa = GetCapa(Geral.idEnvMal)
    Geral.IndTransac = " "
    Geral.TpRep = 0

    Geral.ValorTrans = formataValor(Geral.rstDoctos!Valor)
    ValorLiquido = Geral.rstDoctos!segurados + Geral.rstDoctos!empresa + Geral.rstDoctos!valorterceiro - Geral.rstDoctos!DeducaoFpas
    
    Geral.hsSQLa = "exec grpselet "
   
   'monta header
    MontaHeader
   
   'monta parte variavel
    MontaComplemento
   
    Geral.hsSQLa = Geral.hsSQLa & ", 24309"
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.rstDoctos!TipoIdentificacao)
    Geral.hsSQLa = Geral.hsSQLa & ", '" & Geral.rstDoctos!identificacao & "'"
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.rstDoctos!Fpas)
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Mid(Geral.rstDoctos!Competencia, 5, 2) & Mid(Geral.rstDoctos!Competencia, 3, 2))
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(formataValor(Geral.rstDoctos!segurados))
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(formataValor(Geral.rstDoctos!empresa))
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.rstDoctos!CodigoTerceiro)
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(formataValor(Geral.rstDoctos!valorterceiro))
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(formataValor(Geral.rstDoctos!DeducaoFpas))
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(formataValor(ValorLiquido))
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(formataValor(Geral.rstDoctos!AtualizacaoMonetaria))
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(formataValor(Geral.rstDoctos!Jurosval_jur))
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorTrans)
   
    MontaComplementoVariavel

    LocalLog "Grps - SP " & Geral.hsSQLa
End Sub
Sub LogGARE()

On Error GoTo TrataErro
      
    '==================================='
    ' TRANSAÇÃO 18(nosso número) - GARE '
    '==================================='

    Dim MsgIda          As String
    Dim MsgRetorno      As String '
    Dim Funcao          As String * 14
    Dim RetQX           As Integer
    Dim HeaderTx        As String
    Dim TamIda          As String
    Dim Vez             As Integer
    Dim TipoDocto       As Byte
    Dim spRetorno       As Integer
        
   'se campo do cgc/cpf esteja com zeros devido ao cod.receita,o campo de tipo de pessoa devera ser 0
    If Val(Format(Geral.rstDocto!CPFCGC, "00000000000000")) = 0 Then
        TipoDocto = 0            'nem pessoa juridica nem fisica
    Else
        If Len(Trim(Val(Format(Geral.rstDocto!CPFCGC, "00000000000000")))) > 11 Then
            TipoDocto = 2         'pessoa juridica 2
        Else
            TipoDocto = 1         'pessoa fisica 1
        End If
    End If
    
    If Geral.rstDocto!Indicador_Autenticacao = "S" Then
                   
        CalculaNSU
    
         HeaderTx = "BHGG" & "000000" & Caixa.VersaoAtual & _
                     Format(Parametros.AgenciaCentral, "0000") & _
                     Format(Parametros.AgenciaSatelite, "0000") & "000" & _
                     Format(Caixa.Caixa, "000") & "1" & "000000" & _
                     "000000" & "0" & Format(Now, "HHMM") & "1" & "3" & _
                     "0000000000"
                
        MsgIda = HeaderTx & Trim(Geral.rstDocto!Indicador_Servico_Autenticacao) & _
                 Format(Geral.rstDocto!CPFCGC, String(14, "0")) & _
                 TipoDocto & _
                 Format(Parametros.AgenciaSatelite, "0000") & _
                 Geral.DataProcessamento & Format(Geral.rstDocto!valorreceita * 100, String(13, "0")) & _
                 CVar(Geral.rstDocto!Tipo_Servico)
                 
       'Envia 1ª mensagem ao Host
        TamIda = Format(Len(Trim(MsgIda)), "0000")
        MsgRetorno = String(1921, " ")
        Funcao = "1" & TamIda & "1921****"
         
        LocalLog MsgIda
        
       'Envia BHS1
        Call Abrelinha("BHGG")
        RetQX = qxhostnt(Funcao, MsgIda, MsgRetorno)
        Call FechaLinha("BHGG")
                   
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'se retorno = 20,41,71 -> o micro deverá ser reinicializado '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If (RetQX = 20) Or (RetQX = 41) Or (RetQX = 71) Then
             LocalLog "Retorno BHGG: " & Mid(MsgRetorno, 58, 2)
             MsgBox "Atenção. Ocorreu um erro de comunicação com o Servidor da Agência. Reinicialize este equipamento. Retorno DLL = " & Format(RetQX, "00"), vbOKOnly + vbCritical, "Atenção"
             Close #20
             End
        End If
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Retorno = 21,30,43,47,52,62,80 -> tentar novamente.'
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        If (RetQX = 21) Or (RetQX = 30) Or (RetQX = 43) Or (RetQX = 47) Or _
           (RetQX = 52) Or (RetQX = 62) Or (RetQX = 80) Then
        
            Vez = 1
            Do
               
                Espera (5 * Vez)
                     
               'tentar novamente
                Call Abrelinha("RE-envio BHGG")
                RetQX = qxhostnt(Funcao, MsgIda, MsgRetorno)
                Call FechaLinha("RE-envio BHGG")
                
                Vez = Vez + 1
            
            Loop Until Not ((RetQX = 21) Or (RetQX = 30) Or (RetQX = 43) Or _
                            (RetQX = 47) Or (RetQX = 52) Or (RetQX = 62) Or (RetQX = 80)) And Vez <> 4
        
        End If
        
        LocalLog "Retorno BHGG: " & Format(RetQX, "00") & "MsgRetorno: " & MsgRetorno
                                
    End If
    
    If (RetQX = 0 And (Mid(MsgRetorno, 58, 2) = "00" Or Mid(MsgRetorno, 58, 2) = "03")) Or _
        Geral.rstDocto!Indicador_Autenticacao = "N" Then
              
       'variaveis do header
        Geral.CodTransacao = "0238"
        Geral.Evento = 576
        Geral.TipoTransacao = 2
        Geral.Capa = GetCapa(Geral.idEnvMal)
        Geral.IndTransac = " "
        Geral.CapaBack = 0
        Geral.TpRep = 2
        Geral.ValorTrans = formataValor(Geral.rstDoctos!Valor)
        Geral.ValorMora = formataValor(Geral.rstDocto!Multa)
        
        Geral.hsSQLa = "exec garesp "
        
       'monta header
        MontaHeader
        
        MontaComplemento
        
        Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Mid(Geral.rstDocto!Vecto, 7, 2) & Mid(Geral.rstDocto!Vecto, 5, 2) & Mid(Geral.rstDocto!Vecto, 1, 4))
        Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.rstDocto!receita)
        Geral.hsSQLa = Geral.hsSQLa & ", '" & Format(Geral.rstDocto!InscricaoEstadual, String(12, "0")) & "'"
        Geral.hsSQLa = Geral.hsSQLa & ", " & TipoDocto
        Geral.hsSQLa = Geral.hsSQLa & ", '" & Format(Geral.rstDocto!CPFCGC, "00000000000000") & "'"
        Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.rstDocto!dividaAtiva
        Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Mid(Geral.rstDocto!Referencia, 5, 2) & Mid(Geral.rstDocto!Referencia, 1, 4))
        Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.rstDocto!AIIM)
        Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(formataValor(Geral.rstDocto!valorreceita))
        Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(formataValor(Geral.rstDocto!Juros))
        Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorMora)
        Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(formataValor(Geral.rstDocto!acrescimo))
        Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(formataValor(Geral.rstDocto!honorarios))
        Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorTrans)
        Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.rstDocto!Indicador_Servico_Autenticacao)
        Geral.hsSQLa = Geral.hsSQLa & ", '" & CVar(Geral.rstDocto!Indicador_Autenticacao)
        
        If Geral.rstDocto!Indicador_Autenticacao = "S" Then
        
            Geral.hsSQLa = Geral.hsSQLa & "', '" & CVar(Mid(MsgRetorno, 134, 64))
            LocalLog "MsgRetorno: " & CVar(Mid(MsgRetorno, 134, 64))
            
            spRetorno = MDIQuery.updGareAD(Geral.DataProcessamento, Geral.rstDoctos!iddocto, CVar(Mid(MsgRetorno, 134, 64)))
            
            If spRetorno <> 0 Then
                Err.Raise 999, App.Title, "Falha ao Atualizar AutenticaçãoDigital"
            End If
            
            Geral.hsSQLa = Geral.hsSQLa & "', " & CVar(Geral.rstDocto!Tipo_Servico)
            
           'Complemento Variavel
            Geral.hsSQLa = Geral.hsSQLa & ", '" & CVar(Mid(MsgRetorno, 198, 5))
            Geral.hsSQLa = Geral.hsSQLa & "', '" & CVar(Mid(MsgRetorno, 203, 16))
            Geral.hsSQLa = Geral.hsSQLa & "', '" & CVar(Mid(MsgRetorno, 219, 40)) & "'"
            
        Else
        
            Geral.hsSQLa = Geral.hsSQLa & "', '"
            Geral.hsSQLa = Geral.hsSQLa & "', " & CVar(Geral.rstDocto!Tipo_Servico)
            Geral.hsSQLa = Geral.hsSQLa & ", 0"
            Geral.hsSQLa = Geral.hsSQLa & ", ''"
            Geral.hsSQLa = Geral.hsSQLa & ", ''"
            
        End If
        
        MontaComplementoVariavel
        
        LocalLog "Gare - SP " & Geral.hsSQLa

    Else
    
       'Verificar retornos da BHC9
        Geral.CodOcorrencia = 999
        Geral.RetTransacao = 93
        Geral.PreparouLog = 1
        
    End If
    
    Exit Sub
 
TrataErro:

    If Err.Number = 964 Or Err.Number = 965 Then
       'Erro na Abertura/Fechamento de Linha
        Geral.CodOcorrencia = 999
        Geral.RetTransacao = 98
        Call DevolveDocumentos
        Exit Sub
    ElseIf Err.Number = 963 Then
       'Erro Subida da confirmaçao (BHS3)
        Geral.CodOcorrencia = 999
        Geral.RetTransacao = 97
        Call DevolveDocumentos
        Exit Sub
    End If
   
    Select Case TratamentoErro("Falha no GARE.", Err)
      Case eSair
          End
      Case eRepetir
          Resume
      Case eContinuar
          Resume Next
    End Select

End Sub
Sub LogArrecConvenc()
   
    '======================================================='
    ' TRANSAÇÃO 27(nosso número) - Arrecadação Convencional '
    '======================================================='
   
   'variaveis do header
    Geral.CodTransacao = "0120"
    Geral.Evento = 850
    
    Geral.TipoTransacao = 2
    Geral.Capa = GetCapa(Geral.idEnvMal)
    Geral.IndTransac = " "
    Geral.CapaBack = 1
    Geral.TpRep = 1
    Geral.ValorTrans = formataValor(Geral.rstDoctos!Valor)
       
    Geral.hsSQLa = "exec trn0120"
    
   'monta header
    MontaHeader
   
   'monta parte variavel
    MontaComplemento
    
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.rstDocto!Produto)
    
    If Geral.rstDocto!Produto = 3160 Or Geral.rstDocto!Produto = 3170 Then
       Geral.hsSQLa = Geral.hsSQLa & ", "
       Geral.hsSQLa = Geral.hsSQLa & CVar(Geral.rstDocto!Requisicao)
    Else
       Geral.hsSQLa = Geral.hsSQLa & ", 0"
    End If
    
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorTrans)
    Geral.hsSQLa = Geral.hsSQLa & ", 9"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    
    MontaComplementoVariavel2

    LocalLog "Arrecadacao Convencional - SP " & Geral.hsSQLa

End Sub
Sub LogDarfSimples()
   
   '============================================='
   ' TRANSAÇÃO 17(nosso número) - DARF - Simples '
   '============================================='
   
   'variaveis do header
    Geral.CodTransacao = "0188"
    Geral.Evento = 21
    Geral.TipoTransacao = 2
    Geral.Capa = GetCapa(Geral.idEnvMal)
    Geral.IndTransac = " "
    Geral.CapaBack = 0
    Geral.TpRep = 2
    Geral.ValorTrans = formataValor(Geral.rstDoctos!Valor)
    Geral.ValorCob = formataValor(Geral.rstDocto!ValorPrincipal) 'valor principal
    Geral.ValorMora = formataValor(Geral.rstDocto!valorMulta)
    Geral.ValorDesc = formataValor(Geral.rstDocto!Juros)
        
    Geral.hsSQLa = "exec darfsimp "
    
   'monta header
    MontaHeader
   
   'monta parte variavel
    MontaComplemento
   
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Mid(Geral.rstDocto!PeriodoApuracao, 7, 2) & Mid(Geral.rstDocto!PeriodoApuracao, 5, 2) & Mid(Geral.rstDocto!PeriodoApuracao, 3, 2)) 'periodo de apuracao  ddmmaa
    Geral.hsSQLa = Geral.hsSQLa & ", '" & Format(Geral.rstDocto!CGC, "00000000000000") & "'"
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.rstDocto!receita)
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorCob)
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.rstDocto!percentual * 100
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorCob)
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorMora)
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorDesc)
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorTrans)
    
    MontaComplementoVariavel

    LocalLog "Darf Simples - SP " & Geral.hsSQLa

End Sub

Sub LogDarfPreto()
        
    '==========================================='
    ' TRANSAÇÃO 16(nosso número) - DARF - Preto '
    '==========================================='
   
   'variaveis do header
    Geral.CodTransacao = "0388"
    Geral.Evento = 21
    Geral.TipoTransacao = 2
    Geral.Capa = GetCapa(Geral.idEnvMal)
    Geral.IndTransac = " "
    Geral.CapaBack = 0
    Geral.TpRep = 2
    Geral.ValorTrans = formataValor(Geral.rstDoctos!Valor)
    Geral.ValorCob = formataValor(Geral.rstDocto!ValorPrincipal)    'valor principal
    Geral.ValorMora = formataValor(Geral.rstDocto!valorMulta)
    Geral.ValorDesc = formataValor(Geral.rstDocto!Juros)            'juros
       
    Geral.hsSQLa = "exec darfpret "
      
   'monta header
    MontaHeader
   
   'monta parte variavel
    MontaComplemento
    
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Mid(Geral.rstDocto!PeriodoApuracao, 7, 2) & Mid(Geral.rstDocto!PeriodoApuracao, 5, 2) & Mid(Geral.rstDocto!PeriodoApuracao, 3, 2)) 'periodo de apuracao  ddmmaa
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Mid(Geral.rstDocto!Vecto, 7, 2) & Mid(Geral.rstDocto!Vecto, 5, 2) & Mid(Geral.rstDocto!Vecto, 3, 2)) 'vencimento  ddmmaa
    Geral.hsSQLa = Geral.hsSQLa & ", '" & Format(Geral.rstDocto!CPFCGC, "00000000000000") & "'"   'cgc/cpf - char(14)
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.rstDocto!CodigoReceita)                'codigo da receita
    Geral.hsSQLa = Geral.hsSQLa & ", '" & Format(Geral.rstDocto!Referencia, "00000000000000000") & "'" 'referencia cliente - char
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorCob)
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorMora)
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorDesc)
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorTrans)
    
    MontaComplementoVariavel

    LocalLog "Darf Preto - SP " & Geral.hsSQLa

End Sub
Sub LogCobEspecialSemCB()
   '==================================================================='
   ' TRANSAÇÃO 14 (nosso número) - Cobrança Especial UBB - Via Teclado '
   '==================================================================='
   
    Dim DtVencimento             As Long
    Dim localStatus              As Integer
    Dim localCodOcorrencia       As Integer
    Dim RstMDI                   As Recordset
    Dim spRetorno                As Integer
   
   'variaveis do header
    Geral.CodTransacao = "0086"
    Geral.Evento = 26
    Geral.TipoTransacao = 2
    Geral.Capa = GetCapa(Geral.idEnvMal)
    Geral.IndTransac = " "
    Geral.TpRep = 1
   
    DtVencimento = Mid(Geral.rstDocto!Vecto, 7, 2) & Mid(Geral.rstDocto!Vecto, 5, 2) & Mid(Geral.rstDocto!Vecto, 3, 2)  'ddmmaa

    Geral.AgenCob = Geral.rstDocto!Agencia
    Geral.ContaCob = Geral.rstDocto!Cedente
    Geral.NossoNumCob = Geral.rstDocto!NossoNumero
    Geral.CodCVTCob = Geral.rstDocto!CVT
    Geral.ValorCob = formataValor(Geral.rstDocto!ValorBase)
    Geral.ValorMora = formataValor(Geral.rstDocto!Juros)
    Geral.ValorDesc = formataValor(Geral.rstDocto!Desconto + Geral.rstDocto!Abatimento)
    Geral.ValorTrans = formataValor(Geral.rstDoctos!Valor)
            
    Geral.hsSQLa = "exec cobresp "
    
   'monta header
    MontaHeader

   'monta parte variavel
    MontaComplemento
        
    Geral.hsSQLa = Geral.hsSQLa & ", '" & String(44, " ") & "'"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.CodCVTCob
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.AgenCob
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & DtVencimento
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.ContaCob
    Geral.hsSQLa = Geral.hsSQLa & ", '" & Format(Geral.NossoNumCob, String(15, "0")) & "'"
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorCob)
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorDesc)
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorMora)
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorTrans)
    
    Geral.hsSQLa = Geral.hsSQLa & ", 2"
         
    If IsNull(Geral.rstDocto!BHVC_Descricao) Or (Mid(Geral.rstDocto!BHVC_Descricao, 1, 1) = " ") Then 'neste caso não foi feita pesquisa
       localStatus = 0
       localCodOcorrencia = 0
    Else
       localStatus = Val(Mid(Geral.rstDocto!BHVC_Descricao, 1, 2))
       localCodOcorrencia = Val(Mid(Geral.rstDocto!BHVC_Descricao, 3, 2))
    
       If Not Geral.ehVinculoManual Then 'Grava informação na capa correspondente q trata-se de vínculo manual
       
          spRetorno = MDIQuery.updStatusCapa(Geral.DataProcessamento, Geral.rstCapa!idcapa, ST_CapaVinculoManual)
       
          If spRetorno <> 0 Then
             MsgBox "ATENÇÃO !!! (25)Capa para vinculo manual não atualizada. Tecle <enter> para continuar.", vbOKOnly + vbCritical, "Atenção"
          End If
                             
          Geral.ehVinculoManual = True
       
       End If
       
       MDIQuery.updOcorrenciaDocumento Geral.DataProcessamento, Geral.rstDoctos!iddocto, localCodOcorrencia
        
    End If
       
    Geral.hsSQLa = Geral.hsSQLa & ", " & localStatus          'status se pode receber titulo
    Geral.hsSQLa = Geral.hsSQLa & ", " & localCodOcorrencia   'codigo de mensagem de erro
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    
    MontaComplementoVariavel
    
    LocalLog "Cobranca Esp. sem Cod.Barras - SP " & Geral.hsSQLa

End Sub
Sub LogCobRegistradaSemCB()
   
    '==================================================================='
    ' TRANSAÇÃO 13 (nosso número) -  UNICOBRANCA REGISTRADA Via Teclado '
    '==================================================================='
    
    Dim localStatus                     As Integer
    Dim localCodOcorrencia              As Integer
    Dim localVrCobrancaRegistrada       As String
    Dim RstMDI                          As Recordset
    Dim spRetorno                       As Integer
   
   'variaveis do header
    Geral.CodTransacao = "0084"
    Geral.Evento = 25
    Geral.TipoTransacao = 2
    Geral.Capa = GetCapa(Geral.idEnvMal)
    Geral.IndTransac = " "
    Geral.TpRep = 1
         
   'variaveis da parte variavel da stored procedure
    Geral.VencCob = CLng(Mid(Geral.rstDocto!Vecto, 7, 2) & Mid(Geral.rstDocto!Vecto, 5, 2) & Mid(Geral.rstDocto!Vecto, 3, 2))  'ddmmaa
    Geral.AgenCob = CLng(Geral.rstDocto!Agencia)
    Geral.NossoNumCob = Format(Geral.rstDocto!NossoNumero, String(15, "0"))     'nosso nro
    Geral.CodCVTCob = CLng(Geral.rstDocto!CVT)          'codigo cvt
   
   'valor do documento
    localVrCobrancaRegistrada = formataValor(Geral.rstDocto!ValorBase)  'valor documento
   
    Geral.ValorTrans = formataValor(Geral.rstDoctos!Valor)
        
    Geral.ValorCob = "0.00"
    Geral.ValorMora = "0.00"
    Geral.ValorDesc = "0.00"
    Geral.ValorAbat = "0.00"
       
    Geral.hsSQLa = "exec unicobcx "
    
   'monta header
    MontaHeader
   
   'monta parte variavel
    MontaComplemento
   
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.CodCVTCob
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.VencCob
    Geral.hsSQLa = Geral.hsSQLa & ", ' '"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.AgenCob
    Geral.hsSQLa = Geral.hsSQLa & ", '" & Geral.NossoNumCob & "'"
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(localVrCobrancaRegistrada)
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorMora)
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorDesc)
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorAbat)
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorTrans)
    Geral.hsSQLa = Geral.hsSQLa & ", ' '"
   
    Geral.hsSQLa = Geral.hsSQLa & ", 2"
    
    If IsNull(Geral.rstDocto!BHVC_Descricao) Or (Mid(Geral.rstDocto!BHVC_Descricao, 1, 1) = " ") Then
       localStatus = 0
       localCodOcorrencia = 0
    Else
       localStatus = Val(Mid(Geral.rstDocto!BHVC_Descricao, 1, 2))
       localCodOcorrencia = Val(Mid(Geral.rstDocto!BHVC_Descricao, 3, 2))
        
       If Not Geral.ehVinculoManual Then 'Grava informação na capa correspondente q trata-se de vínculo manual
       
          spRetorno = MDIQuery.updStatusCapa(Geral.DataProcessamento, Geral.rstCapa!idcapa, "7")
                          
          If spRetorno <> 0 Then
             MsgBox "ATENÇÃO !!! (25)Capa para vinculo manual não atualizada. Tecle <enter> para continuar.", vbOKOnly + vbCritical, "Atenção"
          End If
                             
          Geral.ehVinculoManual = True
       
       End If
        
      'grava código de ocorrência p/ documento
       MDIQuery.updOcorrenciaDocumento Geral.DataProcessamento, Geral.rstDoctos!iddocto, localCodOcorrencia
       
    End If
     
    Geral.hsSQLa = Geral.hsSQLa & ", " & localStatus          'status se pode receber titulo
    Geral.hsSQLa = Geral.hsSQLa & ", " & localCodOcorrencia   'codigo de mensagem de erro
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"

    MontaComplementoVariavel
    
    LocalLog "Cobranca Reg. sem Cod.Barras - SP " & Geral.hsSQLa
    
    Exit Sub

TrataErro:

Screen.MousePointer = 0
    Select Case TratamentoErro("Falha no módulo: [Cobrança Registrada sem CB] .", Err)
        Case eSair
            End
        Case eRepetir
            Resume
        Case eContinuar
            Resume Next
    End Select

End Sub
Sub LogDepositoCC()
    
On Error GoTo TrataErro

   '=================================================='
   ' TRANSAÇÃO 02(nosso número) - Deposito C.Corrente '
   '=================================================='
       
    Dim RstUBB_BHQC                 As Recordset     'Capa de cheques
    Dim RstUBB_BHQQ                 As Recordset     'Cheques
    Dim RstMdiChequesDeposito       As Recordset
    Dim spRetorno                   As Integer
    Dim contLaco                    As Integer
    Dim totLaco                     As Integer
    Dim i                           As Integer
    Dim ContCheques                 As Integer
    Dim QtdeCheques                 As Integer
    Dim ValorCheque                 As String
    Dim ValorDinheiro               As String
    Dim ValorSomado                 As String
    
    ValorCheque = 0
    ValorSomado = 0
    ValorDinheiro = 0
      
   'Verifica CMC7 da capa de deposito
    If Not ValidaCMC7Deposito(Geral.rstDoctos!leitura) Then
        MDIQuery.insLogErro Geral.DataProcessamento, Caixa.Estacao, 123456, "Encontrada inconsistencia na Capa de Deposito, Capa: " & Trim(Geral.Capa)
    End If
       
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' Obtem os cheque Deposito (Cheque diversos) ou dinheiro (Cheque UBB / LI) '
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Not GetChequeCashDeposito(Geral.rstCapa!idcapa, Geral.rstDoctos!Vinculo, Geral.rstDoctos!Valor, QtdeCheques, ValorDinheiro, ValorCheque, ValorSomado, True) Then Exit Sub
    
    Geral.ValorTrans = formataValor(Geral.rstDoctos!Valor)
    Geral.NroDeposito = Mid(Geral.rstDoctos!leitura, 9, 3) & Mid(Geral.rstDoctos!leitura, 12, 6) & Mid(Geral.rstDoctos!leitura, 4, 4)
    
   'verifica se o campo de identificação está em branco, se estiver,preenche com zeros
    If Val(Geral.rstDocto!identificado) = 0 Then
        Geral.IdentDep = 0
    Else
        Geral.IdentDep = Geral.rstDocto!identificado
    End If
   
    Geral.TipoConta = "C"
    Geral.AgenciaVinculo = Geral.rstDocto!Agencia
    Geral.ContaVinculo = Geral.rstDocto!Conta
   
    DePara
   
    If Geral.PreparouLog = 1 Then
        Exit Sub
    End If
   
   '---------------------------'
   'processamento INTERAGENCIA '
   '---------------------------'
  
   'variaveis para o header
    Geral.CodTransacao = "2003"
    Geral.Evento = 4
    Geral.TipoTransacao = 2
    Geral.Capa = GetCapa(Geral.idEnvMal)
    Geral.IndTransac = "O"
    Geral.CapaBack = 1
    Geral.TpRep = 0
  
    Geral.hsSQLa = "exec depccoff "
    
   'monta header
    MontaHeader
  
   'monta parte variavel
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.ContaCob
    Geral.hsSQLa = Geral.hsSQLa & ", ' '"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.IdentDep
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(ValorDinheiro)
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(ValorCheque)
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.AgenCob
    Geral.hsSQLa = Geral.hsSQLa & ", '" & Geral.NroDeposito & "'"
    Geral.hsSQLa = Geral.hsSQLa & ", " & QtdeCheques
    Geral.hsSQLa = Geral.hsSQLa & ", '" & Parametros.TipoAgencia & "'"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.NSU1
  
    CalculaNSU
    Caixa.NSU2 = Caixa.NSU
    Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.NSU2
    Geral.SeqPagto = Caixa.NSU2
    Geral.SeqRecto = Caixa.NSU1
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", '000000000000000000000000000000000000'"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"     'SDV2
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
 
    LocalLog "Deposito Interagencia - SP " & Geral.hsSQLa
    
    If GetDocumentoTransmitido(EnumDeposito) Then
        Exit Sub
    End If
    
   'gravar o novo nsu desta transação antes de envia-la para o UBB-NT
    spRetorno = MDIQuery.updNsuDocto(Geral.DataProcessamento, _
                                     Geral.rstCapa!idcapa, _
                                     Geral.rstDoctos!iddocto, _
                                     Caixa.NSU1, _
                                     Caixa.Caixa)
                           
    If spRetorno <> 0 Then
        MsgBox "5490. ATENÇÃO! Deposito a ser enviado Log, não atualizado o NSU. ", vbOKOnly + vbCritical, "Atenção"
        Exit Sub
    End If
    
    LocalLog "Chamada da sp_depcc_inter: " & Format(Now, "hh:mm:ss")
    Set Geral.rst = UBBQuery.ExecuteSQL(Geral.hsSQLa)
    LocalLog "Retorno da sp_depcc_inter: " & Format(Now, "hh:mm:ss")
    
   'se neste caso (interag) o 15º parametro for > 1, reenviar pois encontrou mais de uma conta.
    If (Val(Geral.rst(0)) = 0) And Geral.rst.Fields.Count > 14 Then
            
        If (Val(Geral.rst(14)) > 1) Then
            
           'variaveis para o header
            Geral.CodTransacao = "2003"
            Geral.Evento = 4
            Geral.TipoTransacao = 2
            Geral.Capa = GetCapa(Geral.idEnvMal)
            Geral.IndTransac = "O"
            Geral.CapaBack = 1
            Geral.TpRep = 0
     
           'stored procedure do deposito INTERAGENCIA
            Geral.hsSQLa = "exec depccoff "
      
           'monta header
            MontaHeaderDepInter
      
           'monta parte variavel
            Geral.hsSQLa = Geral.hsSQLa & ", 0"
            Geral.hsSQLa = Geral.hsSQLa & ", 0"
            Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.ContaCob
            Geral.hsSQLa = Geral.hsSQLa & ", ' '"
            Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.IdentDep
            Geral.hsSQLa = Geral.hsSQLa & ", 0"
            Geral.hsSQLa = Geral.hsSQLa & ", 0"
            Geral.hsSQLa = Geral.hsSQLa & ", 0"
            Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(ValorDinheiro)
            Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(ValorCheque)
            Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.AgenCob
            Geral.hsSQLa = Geral.hsSQLa & ", '" & Geral.NroDeposito & "'"
            Geral.hsSQLa = Geral.hsSQLa & ", " & QtdeCheques
            Geral.hsSQLa = Geral.hsSQLa & ", '" & Parametros.TipoAgencia & "'"
            Geral.hsSQLa = Geral.hsSQLa & ", 0"
            Geral.hsSQLa = Geral.hsSQLa & ", 0"
            Geral.hsSQLa = Geral.hsSQLa & ", 0"
            Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.NSU1
            CalculaNSU
           
            Caixa.NSU2 = Caixa.NSU
            
            Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.NSU2
            Geral.SeqPagto = Caixa.NSU2
            Geral.SeqRecto = Caixa.NSU1
            Geral.hsSQLa = Geral.hsSQLa & ", 0"
            Geral.hsSQLa = Geral.hsSQLa & ", 0"
            Geral.hsSQLa = Geral.hsSQLa & ", 0"
            Geral.hsSQLa = Geral.hsSQLa & ", 0"
            Geral.hsSQLa = Geral.hsSQLa & ", 0"
            Geral.hsSQLa = Geral.hsSQLa & ", '000000000000000000000000000000000000'"
            Geral.hsSQLa = Geral.hsSQLa & ", 0"
            Geral.hsSQLa = Geral.hsSQLa & ", 0"
            Geral.hsSQLa = Geral.hsSQLa & ", 0"
    
            LocalLog "reenvio da 2003 - SP " & Geral.hsSQLa
        
            If GetDocumentoTransmitido(EnumDeposito) Then
                Exit Sub
            End If
                       
           'gravar o novo nsu desta transação antes de envia-la para o UBB-NT
            spRetorno = MDIQuery.updNsuDocto(Geral.DataProcessamento, _
                                             Geral.rstCapa!idcapa, _
                                             Geral.rstDoctos!iddocto, _
                                             Caixa.NSU1, _
                                             Caixa.Caixa)

            If spRetorno <> 0 Then
                MsgBox "5590. ATENÇÃO! Deposito a ser enviado Log, não atualizado o NSU. ", vbOKOnly + vbCritical, "Atenção"
                Exit Sub
            End If
           
            LocalLog "Chamada da sp_depcc_inter2: " & Format(Now, "hh:mm:ss")
            Set Geral.rst = UBBQuery.ExecuteSQL(Geral.hsSQLa)
            LocalLog "Retorno da sp_depcc_inter2: " & Format(Now, "hh:mm:ss")
        
        End If
        
    End If
        
    LocalLog "Retorno da sp-deposito c/c: " & Format(Geral.rst(0), "000")
    
    If (Val(Geral.rst(0)) = 0) Then
        
        Geral.GereiLog = 1
        spRetorno = MDIQuery.updDoctoTransmitido(Geral.DataProcessamento, _
                                                 Geral.rstCapa!idcapa, _
                                                 Geral.rstDoctos!iddocto, _
                                                 Geral.SeqRecto, _
                                                 Caixa.Caixa, "N")
 
        If spRetorno <> 0 Then
            MsgBox "333. ATENÇÃO! Documento - deposito c/c - já enviado Log não atualizado no SQL. ", vbOKOnly + vbCritical, "Atenção"
            Exit Sub
        End If
       
       
        If ValorCheque <> 0 Then
            '------------------------------------------------'
            '      Montando a BHQC   (Capa de Cheques)   SP3 '
            '------------------------------------------------'
         
           'variaveis para o Header da BHQC
            Geral.Hora = Format(Now, "HHMM")    'hora
       
           'O caixa só será aberto qdo estacao local com caixa fechado
            Geral.ValorTrans = ValorSomado
            LogAberturaCaixa
        
            Geral.SeqBHQC = Caixa.NSU1
      
           'monta parte fixa
            Geral.hsSQLa = "exec compbhqc "
        
            Geral.hsSQLa = Geral.hsSQLa & "  'BHQC'"
            Geral.hsSQLa = Geral.hsSQLa & ", 0"
            Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.VersaoAtual
            Geral.hsSQLa = Geral.hsSQLa & ", " & Parametros.AgenciaCentral
            Geral.hsSQLa = Geral.hsSQLa & ", " & Parametros.AgenciaSatelite
            Geral.hsSQLa = Geral.hsSQLa & ", 0"
            Geral.hsSQLa = Geral.hsSQLa & ", 1"
            Geral.hsSQLa = Geral.hsSQLa & ", 1"
            Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.Caixa
            Geral.hsSQLa = Geral.hsSQLa & ", 1"
            Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.NSU1
            Geral.hsSQLa = Geral.hsSQLa & ", 0"
            Geral.hsSQLa = Geral.hsSQLa & ", 1"
            Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.Hora
            Geral.hsSQLa = Geral.hsSQLa & ", ' '"
                     
            If Geral.idEnvMal = "E" Then
                Geral.hsSQLa = Geral.hsSQLa & ", 6"
            Else
                Geral.hsSQLa = Geral.hsSQLa & ", 7"
            End If
            
            Geral.hsSQLa = Geral.hsSQLa & ", 1"
            Geral.hsSQLa = Geral.hsSQLa & ", 827"
            Geral.hsSQLa = Geral.hsSQLa & ", '" & Geral.Capa & "'"
            Geral.hsSQLa = Geral.hsSQLa & ", '" & Caixa.CIF
            Geral.hsSQLa = Geral.hsSQLa & "', '" & Caixa.SDV
            
           'monta parte variavel
            Geral.hsSQLa = Geral.hsSQLa & "', 0"
            Geral.hsSQLa = Geral.hsSQLa & ", '" & Geral.NroDeposito & "'"
            Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(ValorCheque)
            Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(ValorSomado)
            Geral.hsSQLa = Geral.hsSQLa & ", " & QtdeCheques
            Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.AgenCob
            Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.ContaCob
            Geral.hsSQLa = Geral.hsSQLa & ", 1"
            Geral.hsSQLa = Geral.hsSQLa & ", ' '"
        
            '***************************************************************'
            'Verifica se valor informado é diferente do valor processado.   '
            'Caso seja diferente, devemos informar para a stored procedure  '
            'o número do NSU para o acerto a ser criado pela mesma.         '
            '***************************************************************'
         
            If ValorCheque <> ValorSomado Then
            
                Geral.ValorTrans = Abs(ValorCheque - ValorSomado)
                
                CalculaNSU
                Caixa.NSU2 = Caixa.NSU
                Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.NSU2
                CalculaNSU
                Caixa.NSU3 = Caixa.NSU
                Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.NSU3
                
            Else
            
                Geral.hsSQLa = Geral.hsSQLa & ", 0"
                Geral.hsSQLa = Geral.hsSQLa & ", 0"
                
            End If
            
           'SDV2
            Geral.hsSQLa = Geral.hsSQLa & ", '" & Caixa.SDV & "'"
         
            '***************************************'
            ' Executa stored procedure da capa BHQC '
            '***************************************'
           
            LocalLog "Capa Deposito - SP " & Geral.hsSQLa
            Set RstUBB_BHQC = UBBQuery.ExecuteSQL(Geral.hsSQLa)
            LocalLog "ret sp-BHQC C/C - " & RstUBB_BHQC(0)
           
            If (Val(RstUBB_BHQC(0)) = 0) Then
               'Enviar a stored procedure dos cheques - BHQQ (Captura 100%)
             
               '**************************************************************************'
               ' A Stored BHQQ deve ser enviada a cada lote de 5 cheques. Se nâo houverem '
               ' 5 cheques, ela é enviada assim mesmo, com zeros a direita.               '
               ' Para isso foi calculado quantas vezes a mesma será enviada.              '
               '**************************************************************************'
             
               'TotLaço = qtos laços de 5 cheques existem neste deposito
                totLaco = QtdeCheques \ 5
                If QtdeCheques Mod 5 <> 0 Then
                    totLaco = totLaco + 1
                End If
                
                Set RstMdiChequesDeposito = MDIQuery.GetChequesDeposito(Geral.DataProcessamento, _
                                                                         Geral.rstCapa!idcapa, _
                                                                         Geral.rstDoctos!Vinculo)
                
               'RstMdiChequesDeposito.Requery
                ContCheques = 0
                     
               'envio de cada laço de 5 cheques
                For contLaco = 1 To totLaco
                    Geral.TpCtaBHQQ = 1
                    Geral.TipoOperacaoDeposito = "06"
                  
                   'Monta o Header e parte variável da BHQQ para cada 05 cheques
                    MontaHeaderBHQQ
                    
                   'continua com montagem da parte variável
                    Geral.hsSQLa = Geral.hsSQLa & ", " & Val(RstUBB_BHQC(30))
    
                   'envio de cada 05 cheques
                    Do
                 
                        ContCheques = ContCheques + 1
                    
                        CalculaNSU
                                
                        Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.NSU
                        Geral.hsSQLa = Geral.hsSQLa & ", '" & RstMdiChequesDeposito!leitura & "'"
                    
                        If Not ValidaCMC7Cheque(Trim(RstMdiChequesDeposito!leitura)) Then
                            MsgBox "Encontrada inconsistencia no CMC7 cheque-deposito", vbCritical + vbOKOnly, "ATENÇÃO: COMUNIQUE SUPORTE"
                            MDIQuery.insLogErro Geral.DataProcessamento, Caixa.Estacao, 123456, "Encontrada inconsistencia no CMC7 cheque-deposito, Capa: " & Trim(Geral.Capa)
                            End
                        End If
                   
                        'verifica o tipo do cheque (1ºposição do nro do cheque)
                        If Mid(RstMdiChequesDeposito!leitura, 12, 1) = "8" Then
                            Geral.hsSQLa = Geral.hsSQLa & ", 9"     'tipo do cheque - cheque roxo
                        Else
                            Geral.hsSQLa = Geral.hsSQLa & ", 5"     'tipo do cheque - cheque comum
                        End If
                   
                       'valor deste cheque
                        Geral.hsSQLa = Geral.hsSQLa & ", " & formataValor(RstMdiChequesDeposito!Valor)
                  
                       'tipo de compensação do cheque, se cheque terceiro, e valor >= valor inferior, então este é SUPERIOR
                        If (Mid(RstMdiChequesDeposito!leitura, 1, 3) <> "409") And (RstMdiChequesDeposito!Valor >= Val(Parametros.ValorLimiteInferior)) Then
                            Geral.hsSQLa = Geral.hsSQLa & ", 1"
                        Else
                           'se cheque terceiro, e valor < valor inferior, então este é INFERIOR
                            If (Mid(RstMdiChequesDeposito!leitura, 1, 3) <> "409") And (RstMdiChequesDeposito!Valor < Val(Parametros.ValorLimiteInferior)) Then
                                Geral.hsSQLa = Geral.hsSQLa & ", 2"
                            Else
                               'se cheque UBB, e conta = 688111 ou 688112, então este é ADM
                                If ((Mid(RstMdiChequesDeposito!leitura, 1, 3) = "409") Or (Mid(RstMdiChequesDeposito!leitura, 1, 3) = "415")) And _
                                    ((Mid(RstMdiChequesDeposito!leitura, 23, 6) = "688111") Or (Mid(RstMdiChequesDeposito!leitura, 23, 6) = "688112")) Then
                                    Geral.hsSQLa = Geral.hsSQLa & ", 4"
                                Else
                                   'se cheque UBB, e conta <> 688111 E 688112, então este é INTERNA
                                    If ((Mid(RstMdiChequesDeposito!leitura, 1, 3) = "409") Or (Mid(RstMdiChequesDeposito!leitura, 1, 3) = "415")) And _
                                        ((Mid(RstMdiChequesDeposito!leitura, 23, 6) <> "688111") And (Mid(RstMdiChequesDeposito!leitura, 23, 6) <> "688112")) Then
                                        Geral.hsSQLa = Geral.hsSQLa & ", 3"
                                    End If
                                End If
                            End If
                        End If
                  
                        '''''''''''''''''''''''''''''''''''
                        ' Atualiza cheque como já enviado '
                        '''''''''''''''''''''''''''''''''''
                    
                        Geral.GereiLog = 1
                       
                        spRetorno = MDIQuery.updChequeDepositoTransmitido(Geral.DataProcessamento, _
                                                                          Geral.rstCapa!idcapa, _
                                                                          RstMdiChequesDeposito!iddocto, _
                                                                          Caixa.NSU, _
                                                                          Caixa.Caixa)
                        If spRetorno <> 0 Then
                            MsgBox "160. ATENÇÃO! Documento - cheque de deposito - já enviado Log não atualizado no SQL. ", vbOKOnly + vbCritical, "Atenção"
                            Exit Sub
                        End If
                    
                        Call GaugePos(Transmissao, "Cheque Deposito")
                        Espera (0.2)
                        
                        RstMdiChequesDeposito.MoveNext
                        
                        If ContCheques = (5 * contLaco) Then Exit Do
                   
                    Loop Until RstMdiChequesDeposito.EOF
                 
                   'se for ultimo laço de cheques, seta flag de Fim
                    If ContCheques < QtdeCheques Then
                        Geral.hsSQLa = Geral.hsSQLa & ", 'C'"   'indicativo que continua
                    Else
                 
                       'preenche os cheques com branco sobrando no laço
                        If QtdeCheques Mod 5 <> 0 Then
                            For i = 1 To 5 - (QtdeCheques Mod 5)
                                Geral.hsSQLa = Geral.hsSQLa & ", 0"
                                Geral.hsSQLa = Geral.hsSQLa & ", ' '"
                                Geral.hsSQLa = Geral.hsSQLa & ", 0"
                                Geral.hsSQLa = Geral.hsSQLa & ", 0"
                                Geral.hsSQLa = Geral.hsSQLa & ", 0"
                            Next
                        End If
                        Geral.hsSQLa = Geral.hsSQLa & ", 'F'"   'indicativo que NAO continua
                    End If
                 
                   'novo campo numero da oct
                    Geral.hsSQLa = Geral.hsSQLa & ", '0" & String(14, " ") & "'"   'campo ordem credito em branco - char(15)
                 
                    '**********************************************************'
                    ' Executa stored procedure de cada BHQQ com até 05 cheques '
                    '**********************************************************'
                
                    LocalLog "Cheques do Deposito - SP " & Geral.hsSQLa
                    Set RstUBB_BHQQ = UBBQuery.ExecuteSQL(Geral.hsSQLa)
                    LocalLog "ret sp-BHQQ C/C - " & Format(RstUBB_BHQQ(0), "00")
                   
                    If Val(RstUBB_BHQQ(0)) <> 0 Then
                        Call DevolveDocumentos(RstUBB_BHQQ)
                        Exit Sub
                    End If
                                            
                Next
           
                Geral.GereiLog = 1
          
            Else
          
                LocalLog "Retorno de Erro na BHQC: " & Str$(RstUBB_BHQC(0))
                Call DevolveDocumentos(RstUBB_BHQC)
           
            End If
       
        End If
    
    Else
    
        Call DevolveDocumentos
        
    End If

    Exit Sub
    
TrataErro:
    
    Select Case TratamentoErro("Falha no Deposito C/C.", Err, eDoctoSubidaLog)
        Case eSair
            End
        Case eRepetir
            Resume
        Case eContinuar
            Resume Next
        Case eFinalizar
            Exit Sub
    End Select

End Sub
Sub LogDepositoCP()
   
On Error GoTo TrataErro
  '=================================================='
  'TRANSAÇÃO 03(nosso número) - Deposito C.Poupanca  '
  '=================================================='
          
    Dim RstUBB_BHQC                 As Recordset     'Capa de cheques
    Dim RstUBB_BHQQ                 As Recordset     'Cheques
    Dim RstMdiChequesDeposito       As Recordset
    Dim spRetorno                   As Integer
    Dim contLaco                    As Integer
    Dim totLaco                     As Integer
    Dim i                           As Integer
    Dim ContCheques                 As Integer
    Dim QtdeCheques                 As Integer
    Dim ValorCheque                 As String
    Dim ValorDinheiro               As String
    Dim ValorSomado                 As String
    
    ValorCheque = 0
    ValorSomado = 0
    ValorDinheiro = 0
      
   'Verifica CMC7 da capa de deposito
    If Not ValidaCMC7Deposito(Geral.rstDoctos!leitura) Then
        MDIQuery.insLogErro Geral.DataProcessamento, Caixa.Estacao, 123456, "Encontrada inconsistencia na Capa de Deposito, Capa: " & Trim(Geral.Capa)
    End If
       
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' Obtem os cheque Deposito (Cheque diversos) ou dinheiro (Cheque UBB / LI) '
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Not GetChequeCashDeposito(Geral.rstCapa!idcapa, Geral.rstDoctos!Vinculo, Geral.rstDoctos!Valor, QtdeCheques, ValorDinheiro, ValorCheque, ValorSomado, True) Then Exit Sub
    
    Geral.ValorTrans = formataValor(Geral.rstDoctos!Valor)
    Geral.NroDeposito = Mid(Geral.rstDoctos!leitura, 9, 3) & Mid(Geral.rstDoctos!leitura, 12, 6) & Mid(Geral.rstDoctos!leitura, 4, 4)
   
   'verifica se o campo de identificação está em branco, se estiver,preenche com zeros
    If Val(Geral.rstDocto!identificado) = 0 Then
        Geral.IdentDep = 0
    Else
        Geral.IdentDep = Geral.rstDocto!identificado
    End If
   
   'tipo de conta no deposito é poupanca
    Geral.TipoConta = "P"
   
   'pesquisa tabela depara
    Geral.AgenciaVinculo = Geral.rstDocto!Agencia
    Geral.ContaVinculo = Geral.rstDocto!Conta
        
    DePara
   
    If Geral.PreparouLog = 1 Then
        Exit Sub
    End If
   
   'variaveis para o header
    Geral.CodTransacao = "2004"
    Geral.Evento = 822
    Geral.TipoTransacao = 2
    Geral.Capa = GetCapa(Geral.idEnvMal)
    Geral.IndTransac = " "
    Geral.CapaBack = 1
    Geral.TpRep = 0
        
   'stored procedure do deposito
    Geral.hsSQLa = "exec depppcta "
    
   'monta header
    MontaHeader
    
   'monta parte variavel
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.ContaCob
    Geral.hsSQLa = Geral.hsSQLa & ", ' '"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.IdentDep
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(ValorDinheiro)
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(ValorCheque)
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.AgenCob
    Geral.hsSQLa = Geral.hsSQLa & ", '" & Geral.NroDeposito & "'"
    Geral.hsSQLa = Geral.hsSQLa & ", " & QtdeCheques
    Geral.hsSQLa = Geral.hsSQLa & ", '" & Parametros.TipoAgencia & "'"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.NSU1
    
    CalculaNSU
    Caixa.NSU2 = Caixa.NSU
    Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.NSU2
    Geral.SeqPagto = Caixa.NSU2
    Geral.SeqRecto = Caixa.NSU1
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"      'campo não utilizado Ed_Mat_Cap
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"  'SDV2
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
  
    LocalLog "Deposito C/P - SP " & Geral.hsSQLa
    If GetDocumentoTransmitido(EnumDeposito) Then
        Exit Sub
    End If
    
   'gravar o novo nsu desta transação antes de envia-la para o UBB-NT
    spRetorno = MDIQuery.updNsuDocto(Geral.DataProcessamento, _
                                     Geral.rstCapa!idcapa, _
                                     Geral.rstDoctos!iddocto, _
                                     Caixa.NSU1, _
                                     Caixa.Caixa)
                           
    If spRetorno <> 0 Then
        MsgBox "5590. ATENÇÃO! Deposito a ser enviado Log, não atualizado o NSU. ", vbOKOnly + vbCritical, "Atenção"
        Exit Sub
    End If

   'executa stored procedure do DEPOSITO POUPANCA
    Set Geral.rst = UBBQuery.ExecuteSQL(Geral.hsSQLa)
    LocalLog "Retorno da sp_depcp: " & Format(Geral.rst(0), "00")
    
   'verifica se encontrou mais de uma conta
    If (Val(Geral.rst(0)) = 0) And (Val(Geral.rst(23)) > 1) Then
            
        Geral.CodTransacao = "2004"
        Geral.Evento = 822
        Geral.TipoTransacao = 2
        Geral.Capa = GetCapa(Geral.idEnvMal)
        Geral.IndTransac = " "
        Geral.CapaBack = 1
        Geral.TpRep = 0
        
       'stored procedure do deposito
        Geral.hsSQLa = "exec depppcta "
    
       'monta header
        MontaHeaderDepInter
    
       'monta parte variavel
        Geral.hsSQLa = Geral.hsSQLa & ", 0"
        Geral.hsSQLa = Geral.hsSQLa & ", 0"
        Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.ContaCob
        Geral.hsSQLa = Geral.hsSQLa & ", ' '"
        Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.IdentDep
        Geral.hsSQLa = Geral.hsSQLa & ", 0"
        Geral.hsSQLa = Geral.hsSQLa & ", 0"
        Geral.hsSQLa = Geral.hsSQLa & ", 0"
        Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(ValorDinheiro)
        Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(ValorCheque)
        Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.AgenCob
        Geral.hsSQLa = Geral.hsSQLa & ", '" & Geral.NroDeposito & "'"
        Geral.hsSQLa = Geral.hsSQLa & ", " & QtdeCheques
        Geral.hsSQLa = Geral.hsSQLa & ", '" & Parametros.TipoAgencia & "'"
        Geral.hsSQLa = Geral.hsSQLa & ", 0"
        Geral.hsSQLa = Geral.hsSQLa & ", 0"
        Geral.hsSQLa = Geral.hsSQLa & ", 0"
        Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.NSU1
    
        CalculaNSU
        
        Caixa.NSU2 = Caixa.NSU
        Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.NSU2
        Geral.SeqPagto = Caixa.NSU2
        Geral.SeqRecto = Caixa.NSU1
        
        Geral.hsSQLa = Geral.hsSQLa & ", 0"
        Geral.hsSQLa = Geral.hsSQLa & ", 0"
        Geral.hsSQLa = Geral.hsSQLa & ", 0"
        Geral.hsSQLa = Geral.hsSQLa & ", 0"
        Geral.hsSQLa = Geral.hsSQLa & ", 0"
        Geral.hsSQLa = Geral.hsSQLa & ", 0"      'campo não utilizado Ed_Mat_Cap
        Geral.hsSQLa = Geral.hsSQLa & ", 0"
        Geral.hsSQLa = Geral.hsSQLa & ", 0"
  
        LocalLog "reenvio da 2004 - SP " & Geral.hsSQLa
        
        If GetDocumentoTransmitido(EnumDeposito) Then
            Exit Sub
        End If

       'gravar o novo nsu desta transação antes de envia-la para o UBB-NT
        spRetorno = MDIQuery.updNsuDocto(Geral.DataProcessamento, _
                                         Geral.rstCapa!idcapa, _
                                         Geral.rstDoctos!iddocto, _
                                         Caixa.NSU1, _
                                         Caixa.Caixa)

        If spRetorno <> 0 Then
            MsgBox "5590. ATENÇÃO! Deposito a ser enviado Log, não atualizado o NSU. ", vbOKOnly + vbCritical, "Atenção"
            Exit Sub
        End If
        
       'executa stored procedure do DEPOSITO POUPANCA
        LocalLog "Chamada da sp_depcp2: " & Format(Now, "hh:mm:ss")
        Set Geral.rst = UBBQuery.ExecuteSQL(Geral.hsSQLa)
        LocalLog "Retorno da sp_depcp2: " & Format(Now, "hh:mm:ss")
    End If
    
    LocalLog "Retorno da sp-deposito c/p: " & Format(Geral.rst(0), "000")
    
    If (Val(Geral.rst(0)) = 0) Then

        Geral.GereiLog = 1
                
       ''''''''''''''''''''''''''''''''''''''''''''''
       ' Atualiza o status do deposito para enviado '
       ''''''''''''''''''''''''''''''''''''''''''''''
        spRetorno = MDIQuery.updDoctoTransmitido(Geral.DataProcessamento, _
                                                 Geral.rstCapa!idcapa, _
                                                 Geral.rstDoctos!iddocto, _
                                                 Geral.SeqRecto, _
                                                 Caixa.Caixa, "N")
                               
        If spRetorno <> 0 Then
            MsgBox "233. ATENÇÃO! Documento - deposito c/p - já enviado Log não atualizado no SQL. ", vbOKOnly + vbCritical, "Atenção"
            Exit Sub
        End If
                
        If ValorCheque <> 0 Then
           '---------------------------------------'
           'Montando a BHQC (Capa de Cheques)  SP3 '
           '---------------------------------------'
            Geral.Hora = Format$(Now, "HHMM")
            Geral.ValorTrans = ValorSomado
            
            LogAberturaCaixa
            
            Geral.SeqBHQC = Caixa.NSU1
            Geral.hsSQLa = "exec compbhqc "
            
           'monta parte fixa
            Geral.hsSQLa = Geral.hsSQLa & "  'BHQC'"
            Geral.hsSQLa = Geral.hsSQLa & ", 0"
            Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.VersaoAtual
            Geral.hsSQLa = Geral.hsSQLa & ", " & Parametros.AgenciaCentral
            Geral.hsSQLa = Geral.hsSQLa & ", " & Parametros.AgenciaSatelite
            Geral.hsSQLa = Geral.hsSQLa & ", 0"
            Geral.hsSQLa = Geral.hsSQLa & ", 1"
            Geral.hsSQLa = Geral.hsSQLa & ", 1"
            Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.Caixa
            Geral.hsSQLa = Geral.hsSQLa & ", 1"
            Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.NSU1
            Geral.hsSQLa = Geral.hsSQLa & ", 0"
            Geral.hsSQLa = Geral.hsSQLa & ", 1"
            Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.Hora
            Geral.hsSQLa = Geral.hsSQLa & ", ' '"
                    
            If Geral.idEnvMal = "E" Then
                Geral.hsSQLa = Geral.hsSQLa & ", 6"
            Else
                Geral.hsSQLa = Geral.hsSQLa & ", 7"
            End If
            
            Geral.hsSQLa = Geral.hsSQLa & ", 1"
            Geral.hsSQLa = Geral.hsSQLa & ", 827"
            Geral.hsSQLa = Geral.hsSQLa & ", '" & Geral.Capa & "'"
            Geral.hsSQLa = Geral.hsSQLa & ", '" & Caixa.CIF
            Geral.hsSQLa = Geral.hsSQLa & "', '" & Caixa.SDV
            
           'monta parte variavel
            Geral.hsSQLa = Geral.hsSQLa & "', 0"
            Geral.hsSQLa = Geral.hsSQLa & ", '" & Geral.NroDeposito & "'"
            Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(ValorCheque)
            Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(ValorSomado)     'valor dos cheques
            Geral.hsSQLa = Geral.hsSQLa & ", " & QtdeCheques
            Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.AgenCob
            Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.ContaCob
            Geral.hsSQLa = Geral.hsSQLa & ", 2"
            Geral.hsSQLa = Geral.hsSQLa & ", ' '"
              
            If ValorCheque <> ValorSomado Then
            
                Geral.ValorTrans = Abs(ValorCheque - ValorSomado)
                
                CalculaNSU
                Caixa.NSU2 = Caixa.NSU
                Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.NSU2
                CalculaNSU
                Caixa.NSU3 = Caixa.NSU
                Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.NSU3
                
            Else
            
                Geral.hsSQLa = Geral.hsSQLa & ", 0"
                Geral.hsSQLa = Geral.hsSQLa & ", 0"
                
            End If
            
           'SDV2
            Geral.hsSQLa = Geral.hsSQLa & ", '" & Caixa.SDV & "'"
           
           'executa stored procedure da BHQC
            LocalLog "Capa deposito C/P - SP " & Geral.hsSQLa
            Set RstUBB_BHQC = UBBQuery.ExecuteSQL(Geral.hsSQLa)
            LocalLog "Retorno sp-Capa Deposito C/P: " & Format(RstUBB_BHQC(0), "00")
          
            If (Val(RstUBB_BHQC(0)) = 0) Then
                
                '--------------------------------------------------------------------------'
                ' A Stored BHQQ deve ser enviada a cada lote de 5 cheques. Se nâo houverem '
                ' 5 cheques, ela é enviada assim mesmo, com zeros a direita.               '
                ' Para isso foi calculado quantas vezes a mesma será enviada.              '
                '--------------------------------------------------------------------------'
                
               'totLaço = qtos laços de 5 cheques existem neste OCT.
                totLaco = QtdeCheques \ 5
                If QtdeCheques Mod 5 <> 0 Then
                    totLaco = totLaco + 1
                End If
                
                Set RstMdiChequesDeposito = MDIQuery.GetChequesDeposito(Geral.DataProcessamento, _
                                                                         Geral.rstCapa!idcapa, _
                                                                         Geral.rstDoctos!Vinculo)
                
               'RstMdiChequesDeposito.Requery
                ContCheques = 0
            
               'envio de cada laço de 5 cheques
                For contLaco = 1 To totLaco
                               
                    Geral.TpCtaBHQQ = 2
                    Geral.TipoOperacaoDeposito = "08"
                    
                   'Monta o Header e parte variável da BHQQ para cada 05 cheques
                    MontaHeaderBHQQ
                    
                   'continua com montagem da parte variável
                    Geral.hsSQLa = Geral.hsSQLa & ", " & Val(RstUBB_BHQC(30))
                    
                   'envio de cada 05 cheques
                    Do
                                         
                        ContCheques = ContCheques + 1
                        
                        CalculaNSU
                                
                        Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.NSU
                        Geral.hsSQLa = Geral.hsSQLa & ", '" & RstMdiChequesDeposito!leitura & "'"
                    
                        If Not ValidaCMC7Cheque(Trim(RstMdiChequesDeposito!leitura)) Then
                            MsgBox "Encontrada inconsistencia no CMC7 cheque-deposito", vbCritical + vbOKOnly, "ATENÇÃO: COMUNIQUE SUPORTE"
                            MDIQuery.insLogErro Geral.DataProcessamento, Caixa.Estacao, 123456, "Encontrada inconsistencia no CMC7 cheque-deposito, Capa: " & Trim(Geral.Capa)
                            End
                        End If
                   
                       'verifica o tipo do cheque (1ºposição do nro do cheque)
                        If Mid(RstMdiChequesDeposito!leitura, 12, 1) = "8" Then
                            Geral.hsSQLa = Geral.hsSQLa & ", 9"     'tipo do cheque - cheque roxo
                        Else
                            Geral.hsSQLa = Geral.hsSQLa & ", 5"     'tipo do cheque - cheque comum
                        End If
                   
                       'valor deste cheque
                        Geral.hsSQLa = Geral.hsSQLa & ", " & formataValor(RstMdiChequesDeposito!Valor)
                      
                       'tipo de compensação do cheque, se cheque terceiro, e valor >= valor inferior, então este é SUPERIOR
                        If (Mid(RstMdiChequesDeposito!leitura, 1, 3) <> "409") And (RstMdiChequesDeposito!Valor >= Val(Parametros.ValorLimiteInferior)) Then
                            Geral.hsSQLa = Geral.hsSQLa & ", 1"
                        Else
                           'se cheque terceiro, e valor < valor inferior, então este é INFERIOR
                            If (Mid(RstMdiChequesDeposito!leitura, 1, 3) <> "409") And (RstMdiChequesDeposito!Valor < Val(Parametros.ValorLimiteInferior)) Then
                                Geral.hsSQLa = Geral.hsSQLa & ", 2"
                            Else
                               'se cheque UBB, e conta = 688111 ou 688112, então este é ADM
                                If ((Mid(RstMdiChequesDeposito!leitura, 1, 3) = "409") Or (Mid(RstMdiChequesDeposito!leitura, 1, 3) = "415")) And _
                                    ((Mid(RstMdiChequesDeposito!leitura, 23, 6) = "688111") Or (Mid(RstMdiChequesDeposito!leitura, 23, 6) = "688112")) Then
                                    Geral.hsSQLa = Geral.hsSQLa & ", 4"
                                Else
                                   'se cheque UBB, e conta <> 688111 E 688112, então este é INTERNA
                                    If ((Mid(RstMdiChequesDeposito!leitura, 1, 3) = "409") Or (Mid(RstMdiChequesDeposito!leitura, 1, 3) = "415")) And _
                                        ((Mid(RstMdiChequesDeposito!leitura, 23, 6) <> "688111") And (Mid(RstMdiChequesDeposito!leitura, 23, 6) <> "688112")) Then
                                        Geral.hsSQLa = Geral.hsSQLa & ", 3"
                                    End If
                                End If
                            End If
                        End If
                  
                        '''''''''''''''''''''''''''''''''''
                        ' Atualiza cheque como já enviado '
                        '''''''''''''''''''''''''''''''''''
                    
                        Geral.GereiLog = 1
                       
                        spRetorno = MDIQuery.updChequeDepositoTransmitido(Geral.DataProcessamento, _
                                                                      Geral.rstCapa!idcapa, _
                                                                      RstMdiChequesDeposito!iddocto, _
                                                                      Caixa.NSU, _
                                                                      Caixa.Caixa)
                        If spRetorno <> 0 Then
                            MsgBox "164. ATENÇÃO! Documento - cheque de deposito - já enviado Log não atualizado no SQL. ", vbOKOnly + vbCritical, "Atenção"
                            Exit Sub
                        End If
                    
                        Call GaugePos(Transmissao, "Cheque Deposito")
                        Espera (0.2)
                        RstMdiChequesDeposito.MoveNext
                        If ContCheques = (5 * contLaco) Then Exit Do
                   
                    Loop Until RstMdiChequesDeposito.EOF
                 
                   'se for ultimo laço de cheques, seta flag de Fim
                    If ContCheques < QtdeCheques Then
                        Geral.hsSQLa = Geral.hsSQLa & ", 'C'"   'indicativo que continua
                    Else
                 
                       'preenche os cheques com branco se necessario
                        If QtdeCheques Mod 5 <> 0 Then
                            For i = 1 To 5 - (QtdeCheques Mod 5)
                                Geral.hsSQLa = Geral.hsSQLa & ", 0"
                                Geral.hsSQLa = Geral.hsSQLa & ", ' '"
                                Geral.hsSQLa = Geral.hsSQLa & ", 0"
                                Geral.hsSQLa = Geral.hsSQLa & ", 0"
                                Geral.hsSQLa = Geral.hsSQLa & ", 0"
                            Next
                        End If
                        Geral.hsSQLa = Geral.hsSQLa & ", 'F'"   'indicativo que NAO continua
                    End If
                 
                   'novo campo numero da oct
                    Geral.hsSQLa = Geral.hsSQLa & ", '0" & String(14, " ") & "'"   'campo ordem credito em branco - char(15)
                 
                    '¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨'
                    '  Executa stored procedure de cada BHQQ com até 05 cheques '
                    '¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨¨'
                
                    LocalLog "Cheques do Deposito - SP " & Geral.hsSQLa
                    Set RstUBB_BHQQ = UBBQuery.ExecuteSQL(Geral.hsSQLa)
                    LocalLog "ret sp-BHQQ C/C - " & Format(RstUBB_BHQQ(0), "00")
                   
                    If Val(RstUBB_BHQQ(0)) <> 0 Then
                        Call DevolveDocumentos(RstUBB_BHQQ)
                        Exit Sub
                    End If
               
                   'Call GaugePos(Transmissao, Geral.rstDoctos!Nome)
                             
                Next
          
               Geral.GereiLog = 1
                      
            Else
            
                LocalLog "Retorno de Erro na BHQC: " & Str$(RstUBB_BHQC(0))
                Call DevolveDocumentos(RstUBB_BHQC)
                
            End If
            
       End If
        
    Else
    
        Call DevolveDocumentos
        
    End If
    
    Exit Sub
    
TrataErro:

    Select Case TratamentoErro("Falha no Deposito C/P.", Err, eDoctoSubidaLog)
        Case eSair
            End
        Case eRepetir
            Resume
        Case eContinuar
            Resume Next
        Case eFinalizar
            Exit Sub
    End Select
    
End Sub
Private Sub MontaHeaderBHQQ()
           
   '''''''''''''''''''''''''''''''''''''''
   ' Monta a BHQQ   (Detalhe dos Cheques)'
   '''''''''''''''''''''''''''''''''''''''
   
   'variaveis para o header
    Geral.Capa = GetCapa(Geral.idEnvMal)
    Geral.Hora = Format$(Now, "HHMM")
    
   'O caixa só será aberto qdo estacao local com caixa fechado
    LogAberturaCaixa (True)
  
    Geral.hsSQLa = "exec cheqdepo "
   
    Geral.hsSQLa = Geral.hsSQLa & "  'BHQQ'"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.VersaoAtual
    Geral.hsSQLa = Geral.hsSQLa & ", " & Parametros.AgenciaCentral
    Geral.hsSQLa = Geral.hsSQLa & ", " & Parametros.AgenciaSatelite
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 2"
    Geral.hsSQLa = Geral.hsSQLa & ", 1"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.Caixa
    Geral.hsSQLa = Geral.hsSQLa & ", 1"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.NSU1
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 1"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.Hora
    Geral.hsSQLa = Geral.hsSQLa & ", ' '"
    
    If Geral.idEnvMal = "E" Then
       Geral.hsSQLa = Geral.hsSQLa & ", 6"
    Else
       Geral.hsSQLa = Geral.hsSQLa & ", 7"
    End If
    
    Geral.hsSQLa = Geral.hsSQLa & ", 1"
    Geral.hsSQLa = Geral.hsSQLa & ", 826"
    Geral.hsSQLa = Geral.hsSQLa & ", '" & Geral.Capa & "'"
    Geral.hsSQLa = Geral.hsSQLa & ", '" & Caixa.CIF
    Geral.hsSQLa = Geral.hsSQLa & "', '" & Caixa.SDV
    
   'parte variável
    Geral.hsSQLa = Geral.hsSQLa & "', " & Caixa.Caixa
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.SeqPagto
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.SeqRecto
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.TpCtaBHQQ
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.AgenCob
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.ContaCob
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.SeqBHQC
    Geral.hsSQLa = Geral.hsSQLa & ", '" & Geral.TipoOperacaoDeposito & "'"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Val(Parametros.PracaCompensacao)
    
End Sub
Sub LogCobTerceiroComCB()
   
   '================================================================'
   ' TRANSAÇÃO 31(nosso número) - Ficha Compensação - Outros bancos '
   '================================================================'

   On Error GoTo TrataErro
   
   Dim AuxLeitura As String
   
   'variaveis do header
    Geral.CodTransacao = "0041"
    Geral.Evento = 639
    Geral.TipoTransacao = 2
    Geral.Capa = GetCapa(Geral.idEnvMal)
    Geral.IndTransac = " "
    Geral.ValorTrans = formataValor(Geral.rstDoctos!Valor)
       
    Geral.hsSQLa = "exec titouban "
      
   'monta header
    MontaHeader
   
   'monta parte variavel
    MontaComplemento
   
    If Val(Mid(Geral.rstDoctos!leitura, 6, 1)) <> 0 Then
        AuxLeitura = Geral.rstDoctos!leitura
    Else
        AuxLeitura = Left(Geral.rstDoctos!leitura, 4) & "000000000000000" & Right(Geral.rstDoctos!leitura, 25)
    End If
   
    Geral.hsSQLa = Geral.hsSQLa & ", '" & AuxLeitura & "'"                   'dados do cod barras - char(44)
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorTrans)
   
   'verifica se codigo de barras foi corrigido
    If Geral.rstDoctos!CodBarComplem = "N" Then
        Geral.hsSQLa = Geral.hsSQLa & ", 1"
    Else
        Geral.hsSQLa = Geral.hsSQLa & ", 2"
    End If
    
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0.00"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", '0'"   'SDV2
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", ' '"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
   
   LocalLog "Cobranca Terceiro com Cod.Barras - SP " & Geral.hsSQLa
   
   Exit Sub

TrataErro:

Screen.MousePointer = 0
    Select Case TratamentoErro("Falha no módulo: [Cobrança de Terceiro sem CB] .", Err)
        Case eSair
            End
        Case eRepetir
            Resume
        Case eContinuar
            Resume Next
    End Select

End Sub
Sub LogTitulo()
   
   '===================================================================='
   ' TRANSAÇÃO 12(nosso número) - Titulos Outros Bancos SEM Cod. Barras '
   '===================================================================='
   
   'variaveis do header
    Geral.CodTransacao = "0141"
    Geral.Evento = 817
    Geral.TipoTransacao = 2
    Geral.Capa = GetCapa(Geral.idEnvMal)
    Geral.IndTransac = " "
    Geral.CapaBack = 1
    Geral.TpRep = 0
    Geral.ValorTrans = formataValor(Geral.rstDoctos!Valor)
       
    Geral.hsSQLa = "exec tit0141 "
      
   'monta header
    MontaHeader
   
   'monta parte variavel
    MontaComplemento
   
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.rstDocto!Banco            'codigo do banco
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorTrans)
    Geral.hsSQLa = Geral.hsSQLa & ", 0"                                  'cobra tarifa ou não
    Geral.hsSQLa = Geral.hsSQLa & ", 0.00"                               'valor da tarifa - numeric(16,2)
    Geral.hsSQLa = Geral.hsSQLa & ", 0"                                  'nsu da tarifa
    Geral.hsSQLa = Geral.hsSQLa & ", 0"                                  'SDV2
    Geral.hsSQLa = Geral.hsSQLa & ", 0"                                  'flag de contrapartida
    Geral.hsSQLa = Geral.hsSQLa & ", 0"                                  'codigo de mensagem de erro
    
    MontaComplementoVariavel

    LocalLog "Titulo sem Cod.Barras - SP " & Geral.hsSQLa

End Sub
Sub LogUnicobrancaUBB()
   '==========================================='
   ' TRANSAÇÃO 28(nosso número) -  UNICOBRANCA '
   '==========================================='

   On Error GoTo TrataErro
   
    Dim localStatus             As Integer
    Dim localCodOcorrencia      As Integer
    Dim AuxLeitura              As String
    Dim spRetorno               As Integer
   
   'variaveis do header
    Geral.CodTransacao = "0084"
    Geral.Evento = 25
    Geral.TipoTransacao = 2
    Geral.Capa = GetCapa(Geral.idEnvMal)
    Geral.IndTransac = " "
    Geral.TpRep = 1
    
   'variaveis da parte variavel da stored procedure
    Geral.VencCob = Mid(Geral.rstDoctos!leitura, 26, 2) & Mid(Geral.rstDoctos!leitura, 24, 2) & Mid(Geral.rstDoctos!leitura, 22, 2)
    Geral.AgenCob = CLng(Mid(Geral.rstDoctos!leitura, 28, 5))
    Geral.NossoNumCob = "001" & Mid(Geral.rstDoctos!leitura, 33, 12)
    Geral.CodCVTCob = 55395
    
    Geral.ValorMora = formataValor(0)
    Geral.ValorDesc = formataValor(0)
    Geral.ValorAbat = formataValor(0)
    
    Geral.ValorTrans = formataValor(Geral.rstDoctos!Valor)
       
    Geral.hsSQLa = "exec unicobcx "
    
   'monta header
    MontaHeader
   
   'monta parte variavel
    MontaComplemento
    
    If Val(Mid(Geral.rstDoctos!leitura, 6, 1)) <> 0 Then
        AuxLeitura = Geral.rstDoctos!leitura
    Else
        AuxLeitura = Left(Geral.rstDoctos!leitura, 4) & "000000000000000" & Right(Geral.rstDoctos!leitura, 25)
    End If
    
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.CodCVTCob
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.VencCob
    Geral.hsSQLa = Geral.hsSQLa & ", '" & AuxLeitura & "'"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.AgenCob
    Geral.hsSQLa = Geral.hsSQLa & ", '" & Geral.NossoNumCob & "'"
        
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorTrans)
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorMora)
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorDesc)
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorAbat)
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorTrans)
    Geral.hsSQLa = Geral.hsSQLa & ", ' '"
   
   'verifica se codigo de barras foi corrigido
    If Geral.rstDoctos!CodBarComplem = "N" Then
        Geral.hsSQLa = Geral.hsSQLa & ", 1"
    Else
        Geral.hsSQLa = Geral.hsSQLa & ", 2"
    End If
      
    If IsNull(Geral.rstDocto!BHVC_Descricao) Or (Mid(Geral.rstDocto!BHVC_Descricao, 1, 1) = " ") Then 'neste caso não foi feita pesquisa
       localStatus = 0
       localCodOcorrencia = 0
    Else
       localStatus = Val(Mid(Geral.rstDocto!BHVC_Descricao, 1, 2))
       localCodOcorrencia = Val(Mid(Geral.rstDocto!BHVC_Descricao, 3, 2))
    
       If Not Geral.ehVinculoManual Then 'Grava informação na capa correspondente q trata-se de vínculo manual
          spRetorno = MDIQuery.updStatusCapa(Geral.DataProcessamento, Geral.rstCapa!idcapa, "7")
                       
          If spRetorno <> 0 Then
             MsgBox "ATENÇÃO !!!, Capa para vinculo manual não atualizada. Tecle <enter> para continuar.", vbOKOnly + vbCritical, "Atenção"
          End If
                             
          Geral.ehVinculoManual = True
       
       End If
       
       MDIQuery.updOcorrenciaDocumento Geral.DataProcessamento, Geral.rstDoctos!iddocto, localCodOcorrencia
        
    End If
    
    Geral.hsSQLa = Geral.hsSQLa & ", " & localStatus          'localStatus se pode receber titulo
    Geral.hsSQLa = Geral.hsSQLa & ", " & localCodOcorrencia   'codigo de mensagem de erro
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    
    MontaComplementoVariavel
    
    LocalLog "UniCobranca Ubb - SP " & Geral.hsSQLa
    
    Exit Sub

TrataErro:

Screen.MousePointer = 0
    Select Case TratamentoErro("Falha no módulo: [Unicobrança UBB ] .", Err)
        Case eSair
            End
        Case eRepetir
            Resume
        Case eContinuar
            Resume Next
    End Select
    
End Sub
Sub LogCobImediataUBB()

  '=========================================================='
  ' TRANSAÇÃO 29 (nosso número) - Cobrança Imediata UNIBANCO '
  '=========================================================='
  
   On Error GoTo TrataErro
   
    Dim localStatus             As Integer
    Dim localCodOcorrencia      As Integer
    Dim AuxLeitura              As String
    Dim DataVencimento          As Long
    Dim spRetorno               As Integer
    
   'variaveis do header
    Geral.CodTransacao = "0980"
    Geral.Evento = 18
    Geral.TipoTransacao = 2
    Geral.Capa = GetCapa(Geral.idEnvMal)
    Geral.IndTransac = "O"                   'ind transacional (O-offline)
        
    DataVencimento = Mid(Geral.rstDoctos!Vecto, 7, 2) & Mid(Geral.rstDoctos!Vecto, 5, 2) & Mid(Geral.rstDoctos!Vecto, 3, 2)   'ddmmaa
        
   'codigo cvt
    Geral.CodCVTCob = 80020
   
    Geral.AgenCob = CLng(Mid(Geral.rstDoctos!leitura, 28, 4))
    Geral.ContaCob = CLng(Mid(Geral.rstDoctos!leitura, 21, 7))
   
   'nosso numero - calcular modulo 11
    If (Mid(Geral.rstDoctos!leitura, 33, 12) = "000000000000") Then
        Geral.NossoNumCob = "0000000000" & Mid(Geral.rstDoctos!leitura, 26, 4)
        Modulo11 Geral.NossoNumCob
        Geral.NossoNumCob = Geral.NossoNumCob & Geral.RetDigMod11
    Else
        Geral.NossoNumCob = "000" & Mid(Geral.rstDoctos!leitura, 33, 12)
    End If
    
    Geral.ValorTrans = formataValor(Geral.rstDoctos!Valor)
    Geral.ValorMora = formataValor(Geral.rstDoctos!Juros)
    Geral.ValorDesc = formataValor(Geral.rstDoctos!Desconto + Geral.rstDoctos!Abatimento)
       
    If Val(Mid(Geral.rstDoctos!leitura, 6, 1)) <> 0 Then
        Geral.ValorCob = formataValor(Geral.rstDoctos!ValorBase)
    Else
        Geral.ValorCob = "0"
    End If
    
    Geral.hsSQLa = "exec cobrimed "
    
   'monta header
    MontaHeader
   
   'monta parte variavel
    MontaComplemento
   
    If Val(Mid(Geral.rstDoctos!leitura, 6, 1)) <> 0 Then
        AuxLeitura = Geral.rstDoctos!leitura
    Else
        AuxLeitura = Left(Geral.rstDoctos!leitura, 4) & "000000000000000" & Right(Geral.rstDoctos!leitura, 25)
    End If
 
    Geral.hsSQLa = Geral.hsSQLa & ", '" & AuxLeitura & "'"             'dados cod barras - char(44)
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.CodCVTCob
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & DataVencimento
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.AgenCob
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.ContaCob
    Geral.hsSQLa = Geral.hsSQLa & ", '" & Geral.NossoNumCob & "'"
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorCob)
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorDesc)
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorMora)
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorTrans)
    Geral.hsSQLa = Geral.hsSQLa & ", 0"                                'digito localStatus online/offline
      
    If IsNull(Geral.rstDoctos!BHVC_Descricao) Or (Mid(Geral.rstDoctos!BHVC_Descricao, 1, 1) = " ") Then 'neste caso não foi feita pesquisa
       localStatus = 0
       localCodOcorrencia = 0
    Else
       localStatus = Val(Mid(Geral.rstDoctos!BHVC_Descricao, 1, 2))
       localCodOcorrencia = Val(Mid(Geral.rstDoctos!BHVC_Descricao, 3, 2))
    
       If Not Geral.ehVinculoManual Then 'Grava informação na capa correspondente q trata-se de vínculo manual
          spRetorno = MDIQuery.updStatusCapa(Geral.DataProcessamento, Geral.rstCapa!idcapa, "7")
                       
          If spRetorno <> 0 Then
             MsgBox "ATENÇÃO !!! Capa para vinculo manual não atualizada. Tecle <enter> para continuar.", vbOKOnly + vbCritical, "Atenção"
          End If
                            
          Geral.ehVinculoManual = True
       
       End If
           
       MDIQuery.updOcorrenciaDocumento Geral.DataProcessamento, Geral.rstDoctos!iddocto, localCodOcorrencia
        
    End If
    
    MontaComplementoVariavel
   
    LocalLog "Cobranca Imediata Ubb - SP " & Geral.hsSQLa
    
    Exit Sub

TrataErro:

Screen.MousePointer = 0
    Select Case TratamentoErro("Falha no módulo: [Cobrança Imediata UBB] .", Err)
        Case eSair
            End
        Case eRepetir
            Resume
        Case eContinuar
            Resume Next
    End Select

End Sub
Sub LogCobEspecialUBB()

  '=========================================================='
  ' TRANSAÇÃO 30 (nosso número) - Cobrança Especial UNIBANCO '
  '=========================================================='

   On Error GoTo TrataErro
   
    Dim localStatus             As Integer
    Dim localCodOcorrencia      As Integer
    Dim AuxLeitura              As String
    Dim DataVencimento          As Long
    Dim spRetorno               As Integer
    Dim Diferenca               As Double
   
   '=========================================================='
   ' TRANSAÇÃO 30 (nosso número) - Cobrança Especial UNIBANCO '
   '=========================================================='
   
   'variaveis do header
    Geral.CodTransacao = "0086"
    Geral.Evento = 26
    Geral.TipoTransacao = 2
    Geral.Capa = GetCapa(Geral.idEnvMal)
    Geral.IndTransac = " "
    Geral.TpRep = 1
   
    DataVencimento = Mid(Geral.rstDocto!Vecto, 7, 2) & Mid(Geral.rstDocto!Vecto, 5, 2) & Mid(Geral.rstDocto!Vecto, 3, 2)   'ddmmaa
        
    Geral.AgenCob = 9000
    Geral.ContaCob = CLng(Mid(Geral.rstDoctos!leitura, 21, 7))
   
   'nosso numero - calcular modulo 11
    If (Mid(Geral.rstDoctos!leitura, 33, 12) = "000000000000") Then
        Geral.NossoNumCob = "0000000000" & Mid(Geral.rstDoctos!leitura, 26, 4)
        Modulo11 Geral.NossoNumCob
        Geral.NossoNumCob = Geral.NossoNumCob & Geral.RetDigMod11
    Else
        Geral.NossoNumCob = Mid(Geral.rstDoctos!leitura, 30, 15)
    End If
   
   'codigo cvt
    Select Case Mid(Geral.rstDoctos!leitura, 20, 1)
        Case "1"
            Geral.CodCVTCob = 77330
        Case "2"
            Geral.CodCVTCob = 77437
        Case "3"
            Geral.CodCVTCob = 77445
        Case "5"
            Geral.CodCVTCob = 77445
    End Select
   
    Geral.ValorCob = formataValor(Geral.rstDocto!ValorBase)
    Geral.ValorMora = formataValor(Geral.rstDocto!Juros)
    Geral.ValorDesc = formataValor(Geral.rstDocto!Desconto + Geral.rstDocto!Abatimento)
    Geral.ValorTrans = formataValor(Geral.rstDoctos!Valor)

    Geral.hsSQLa = "exec cobresp "
   
   'monta header
    MontaHeader

   'monta parte variavel
    MontaComplemento
    
    If Val(Mid(Geral.rstDoctos!leitura, 6, 1)) <> 0 Then
        AuxLeitura = Geral.rstDoctos!leitura
    Else
        AuxLeitura = Left(Geral.rstDoctos!leitura, 4) & "000000000000000" & Right(Geral.rstDoctos!leitura, 25)
    End If
    
    Geral.hsSQLa = Geral.hsSQLa & ", '" & AuxLeitura & "'"            'dados cod barras - char(44)
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.CodCVTCob
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.AgenCob
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & DataVencimento
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.ContaCob
    Geral.hsSQLa = Geral.hsSQLa & ", '" & Format(Geral.NossoNumCob, String(15, "0")) & "'"
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorCob)
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorDesc)
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorMora)
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorTrans)
   
   'verifica se codigo de barras foi corrigido
    If Geral.rstDoctos!CodBarComplem = "N" Then
        Geral.hsSQLa = Geral.hsSQLa & ", 1"
    Else
        Geral.hsSQLa = Geral.hsSQLa & ", 2"
    End If
    
    If IsNull(Geral.rstDocto!BHVC_Descricao) Or (Mid(Geral.rstDocto!BHVC_Descricao, 1, 1) = " ") Then 'neste caso não foi feita pesquisa
       localStatus = 0
       localCodOcorrencia = 0
    Else
       localStatus = Val(Mid(Geral.rstDocto!BHVC_Descricao, 1, 2))
       localCodOcorrencia = Val(Mid(Geral.rstDocto!BHVC_Descricao, 3, 2))
    
       If Not Geral.ehVinculoManual Then 'Grava informação na capa correspondente q trata-se de vínculo manual
       
          spRetorno = MDIQuery.updStatusCapa(Geral.DataProcessamento, Geral.rstCapa!idcapa, "7")
                          
          If spRetorno <> 0 Then
             MsgBox "ATENÇÃO !!! Capa para vinculo manual não atualizada. Tecle <enter> para continuar.", vbOKOnly + vbCritical, "Atenção"
          End If
                             
          Geral.ehVinculoManual = True
       
       End If
       
       MDIQuery.updOcorrenciaDocumento Geral.DataProcessamento, Geral.rstDoctos!iddocto, localCodOcorrencia
        
    End If
 
    Geral.hsSQLa = Geral.hsSQLa & ", " & localStatus            'localStatus se pode receber titulo
    Geral.hsSQLa = Geral.hsSQLa & ", " & localCodOcorrencia     'codigo de mensagem de erro
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    
    MontaComplementoVariavel

    LocalLog "Cobranca Especial Ubb - SP " & Geral.hsSQLa
    
    Exit Sub

TrataErro:

    Select Case TratamentoErro("Falha no módulo: [Cobrança Especial UBB] .", Err)
        Case eSair
            End
        Case eRepetir
            Resume
        Case eContinuar
            Resume Next
    End Select

End Sub
Sub MontaComplemento()
       
    Geral.hsSQLa = Geral.hsSQLa & ", 1"
    Geral.hsSQLa = Geral.hsSQLa & ", ' '"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", ' '"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 1"
    Geral.hsSQLa = Geral.hsSQLa & ", ' '"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", ' '"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"

End Sub
Sub LogAberturaCaixa(Optional ByVal outProcess As Boolean)

   On Error GoTo TrataErro
   
   Dim spRetorno        As Integer
    
   '============================================================='
   ' ABERTURA / REABERTURA DE CAIXA / CALCULO NSU DOCTO CORRENTE '
   '============================================================='
    
   'verifica se este caixa já está aberto
    If GetSetting("Robo", "Caixa", "Aberto", 0) = 0 Then
    
      'leitura do número do terminal
        CalculaNSU (True)
        Caixa.NSU1 = Caixa.NSU
        
        Geral.hsSQLb = "exec abrecx "
        
       'monta header
        Geral.hsSQLb = Geral.hsSQLb & " '0030'"
        Geral.hsSQLb = Geral.hsSQLb & ", 0"
        Geral.hsSQLb = Geral.hsSQLb & ", " & Caixa.VersaoAtual
        Geral.hsSQLb = Geral.hsSQLb & ", " & Parametros.AgenciaCentral
        Geral.hsSQLb = Geral.hsSQLb & ", " & Parametros.AgenciaSatelite
        Geral.hsSQLb = Geral.hsSQLb & ", 2"
        Geral.hsSQLb = Geral.hsSQLb & ", 4"
        Geral.hsSQLb = Geral.hsSQLb & ", 0"
        Geral.hsSQLb = Geral.hsSQLb & ", " & Caixa.Caixa
        Geral.hsSQLb = Geral.hsSQLb & ", 1"
        Geral.hsSQLb = Geral.hsSQLb & ", " & Caixa.NSU1
        Geral.hsSQLb = Geral.hsSQLb & ", 0"
        Geral.hsSQLb = Geral.hsSQLb & ", 1"
        Geral.hsSQLb = Geral.hsSQLb & ", " & Geral.Hora
        Geral.hsSQLb = Geral.hsSQLb & ", ' '"
             
        If Geral.idEnvMal = "E" Then
           Geral.hsSQLb = Geral.hsSQLb & ", 6"
        Else
           Geral.hsSQLb = Geral.hsSQLb & ", 7"
        End If
        
        Geral.hsSQLb = Geral.hsSQLb & ", 0"
        Geral.hsSQLb = Geral.hsSQLb & ", 126"
        Geral.hsSQLb = Geral.hsSQLb & ", ' ', '"
        Geral.hsSQLb = Geral.hsSQLb & Caixa.CIF
        Geral.hsSQLb = Geral.hsSQLb & "', '" & Caixa.SDV
        
       'monta parte variavel
        Geral.hsSQLb = Geral.hsSQLb & "', 0"
        Geral.hsSQLb = Geral.hsSQLb & ", ' '"
        Geral.hsSQLb = Geral.hsSQLb & ", 0"
        Geral.hsSQLb = Geral.hsSQLb & ", 0"
        Geral.hsSQLb = Geral.hsSQLb & ", 0"
                    
        LocalLog "Abertura de Caixa - SP " & Geral.hsSQLb
        Set Geral.rst = UBBQuery.ExecuteSQL(Geral.hsSQLb)
        LocalLog "Retorno sp_Abertura de Caixa - " & Format(Geral.rst(0), "00")
       
        If (Val(Geral.rst(0)) = 0) Then
            SaveSetting appname:="Robo", section:="Caixa", Key:="Aberto", setting:=1
               
            If Not Geral.RecebendoCapa And Geral.PrimeiraVez Then
               Geral.PrimeiraVez = False
               AgenciaColeta
            End If
            
        Else
        
            MsgBox "ATENÇÃO! Ocorreu o erro " & Str(Geral.rst(0)) & " na abertura da agência de coleta  [" & Parametros.AgenciaSatelite & "]. & A Capa está sendo enviada para ilegíveis. Verifique para que a comunicação prossiga.", vbOKOnly + vbInformation, "Atenção"
            
            Call GaugeTitulo(2)
            
            spRetorno = MDIQuery.UpdCapaStatusCaixaControle(Geral.DataProcessamento, Geral.rstCapa!idcapa, ST_CapaParaIlegivel, Caixa.Caixa)
            If spRetorno <> 0 Then MsgBox "Falha Procedure [ UpdCapaStatusCaixaControle ]", vbCritical + vbOKOnly
            
            MDIQuery.insLog Geral.DataProcessamento, Geral.rstCapa!idcapa, "0", Caixa.UsuarioAtual, "124"
            End
                        
        End If
         
    Else
       
        If Not Geral.RecebendoCapa And Geral.PrimeiraVez Then
           Geral.PrimeiraVez = False
           AgenciaColeta
        End If
        
    End If
   
   'calcula o NSU para a proxima transação
    CalculaNSU (outProcess)
    Caixa.NSU1 = Caixa.NSU
   
    Geral.CaixaAberto = True
    
    Exit Sub

TrataErro:

    Select Case TratamentoErro("Falha no módulo: [Abertura de Caixa] .", Err)
        Case eSair
            End
        Case eRepetir
            Resume
        Case eContinuar
            Resume Next
    End Select
           
End Sub
Sub LogDarm()
   
   '================================='
   ' TRANSAÇÃO DARM (tipodocto = 15) '
   '================================='
   
   'variaveis do header
    Geral.CodTransacao = "0020"
    Geral.Evento = 116
    Geral.TipoTransacao = 2
    Geral.Capa = GetCapa(Geral.idEnvMal)
    Geral.IndTransac = " "
    Geral.CapaBack = 0
    Geral.TpRep = 1
   
    Geral.ValorTrans = formataValor(Geral.rstDoctos!Valor)
       
    Geral.hsSQLa = "exec trndarm "
        
   'monta header
    MontaHeader
   
    MontaComplemento
    
    Geral.hsSQLa = Geral.hsSQLa & ", '" & Geral.rstDoctos!leitura & "'"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Mid(Geral.rstDocto!incidencia, 5, 2) & Mid(Geral.rstDocto!incidencia, 3, 2) 'incidencia  (mmaa)
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.rstDocto!tributo
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorTrans)
    Geral.hsSQLa = Geral.hsSQLa & ", 2"
    
    MontaComplementoVariavel

    LocalLog "Darm - SP " & Geral.hsSQLa

End Sub
Sub MontaComplementoVariavel()
   
   Geral.hsSQLa = Geral.hsSQLa & ", 0"   'saldo base para autorização - numeric(16,2)
   Geral.hsSQLa = Geral.hsSQLa & ", 0"
   Geral.hsSQLa = Geral.hsSQLa & ", 0"
   Geral.hsSQLa = Geral.hsSQLa & ", ' '"
   
  'calcula NSU de pagto dos cheques
   CalculaNSU
   Caixa.NSU2 = Caixa.NSU
   
   Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.NSU2  'nsu pagto cheque

End Sub
Sub MontaComplementoVariavel2()
   
   Geral.hsSQLa = Geral.hsSQLa & ", 0"   'saldo base para autorização - numeric(16,2)
   Geral.hsSQLa = Geral.hsSQLa & ", 0"
   Geral.hsSQLa = Geral.hsSQLa & ", 0"
   Geral.hsSQLa = Geral.hsSQLa & ", ' '"
   Geral.hsSQLa = Geral.hsSQLa & ", 0"

End Sub

Public Function ProcessaEstorno() As Boolean

On Error GoTo TrataErro

    Dim spRetorno As Integer
       
    ProcessaEstorno = True
    
    LocalLog "Processando Estorno "
    
    If Geral.rstCapa!TipoDocto = 4 Or Geral.rstCapa!TipoDocto = 6 Or _
       Geral.rstCapa!TipoDocto = 11 Or _
      (Geral.rstCapa!TipoDocto >= 12 And Geral.rstCapa!TipoDocto < 27) Or _
       Geral.rstCapa!TipoDocto = 28 Or Geral.rstCapa!TipoDocto = 30 Or _
      (Geral.rstCapa!TipoDocto >= 32 And Geral.rstCapa!TipoDocto <= 36) Or _
       Geral.rstCapa!TipoDocto = 38 Or Geral.rstCapa!TipoDocto = 40 Or _
      (Geral.rstCapa!TipoDocto >= 42 And Geral.rstCapa!TipoDocto <= 43) Then
                      
        ProcessaEstorno = EstornoGeral("0032")
       
    ElseIf Geral.rstCapa!TipoDocto = 27 Then
        ProcessaEstorno = EstornoArrecad
    ElseIf Geral.rstCapa!TipoDocto = 5 Then
        ProcessaEstorno = EstornoGeral("5032")
    ElseIf Geral.rstCapa!TipoDocto = 2 Then
        ProcessaEstorno = EstornoDeposito("5032")
    ElseIf Geral.rstCapa!TipoDocto = 3 Then
        ProcessaEstorno = EstornoDeposito("2033")
    ElseIf Geral.rstCapa!TipoDocto = 37 Then
        ProcessaEstorno = EstornoDeposito("0032")
    ElseIf Geral.rstCapa!TipoDocto = 31 Then
        ProcessaEstorno = EstornoTitulo()
    Else
        ProcessaEstorno = False
                
        spRetorno = MDIQuery.insMensagem(Geral.DataProcessamento, _
                                         MSG_TipoDoctoNaoPodeEstor, _
                                         Geral.rstCapa!idcapa, _
                                         Geral.rstCapa!iddocto, _
                                         Caixa.Caixa)
                                        
        If spRetorno <> 0 Then MsgBox "Falha Procedure [ InsMensagem ]", vbCritical + vbOKOnly
       
    End If
    
    Call GaugePos(Estorno, Geral.rstCapa!Nome)
    
    Exit Function
     
TrataErro:

Screen.MousePointer = 0
    Select Case TratamentoErro("Falha no módulo: [Processa Estorno] .", Err)
        Case eSair
            End
        Case eRepetir
            Resume
        Case eContinuar
            Resume Next
    End Select
             
End Function
Public Sub LogGeral()

On Error GoTo TrataErro
    
    Dim spRetorno   As Integer

    If GetDocumentoTransmitido(EnumOutros) Then
        Exit Sub
    End If
           
    LocalLog "ProcessaDocumento: " & Trim(Geral.rstDoctos!Nome) & " - SP: " & Geral.hsSQLa
    
   'gravar o nsu antes de enviar transação para o UBB-NT
    spRetorno = MDIQuery.updNsuDocto(Geral.DataProcessamento, _
                                     Geral.rstCapa!idcapa, _
                                     Geral.rstDoctos!iddocto, _
                                     Caixa.NSU1, _
                                     Caixa.Caixa)
       
    If spRetorno <> 0 Then
       MsgBox "ATENÇÃO! Falha na SP [ updNsuDocto ]. ", vbOKOnly + vbCritical, "Atenção"
       Exit Sub
    End If
    
    Set Geral.rst = UBBQuery.ExecuteSQL(Geral.hsSQLa)
    LocalLog "Retorno sp-generica: " & Format(Geral.rst(0), "000")
    
    If (Val(Geral.rst(0)) = 0) Then
        Geral.GereiLog = 1
    
       'atualizar o docto com já transmitido
        spRetorno = MDIQuery.updDoctoTransmitido(Geral.DataProcessamento, _
                                                 Geral.rstCapa!idcapa, _
                                                 Geral.rstDoctos!iddocto, _
                                                 Caixa.NSU1, _
                                                 Caixa.Caixa, "N")
    
        If spRetorno <> 0 Then
            MsgBox "ATENÇÃO! Falha na SP [ updDoctoTransmitido ]. ", vbOKOnly + vbCritical, "Atenção"
            Exit Sub
        End If
    Else
                        
        Call DevolveDocumentos
    
    End If
      
    Exit Sub
    
TrataErro:
    
    Select Case TratamentoErro("Falha no " & Geral.rstDoctos!Nome, Err, eDoctoSubidaLog)
        Case eSair
            End
        Case eRepetir
            Resume
        Case eContinuar
            Resume Next
        Case eFinalizar
            Exit Sub
    End Select

End Sub
Sub LogChequeInteragencia()

'================================================='
' TRANSAÇÃO 05(nosso número) - Saque Interagencia '
'================================================='
 
On Error GoTo TrataErro
   
    Dim RstUBB          As Recordset
    Dim spRetorno       As Integer
    Dim MsgIda          As String
    Dim MsgRetorno      As String
    Dim Funcao          As String * 14
    Dim RetQX           As Integer
    Dim HeaderTx        As String
    Dim TamIda          As String
    Dim Vez             As Integer
        
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'verifica se cheque está pagando alguma cobrança Ubb vencida e se esta pode ser paga '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Geral.Vinculo = Geral.rstDoctos!Vinculo
        
   'variaveis para o header
    Geral.CodTransacao = "6F01"
    Geral.Evento = 2
    Geral.TipoTransacao = 2
    Geral.Capa = GetCapa(Geral.idEnvMal)
    Geral.IndTransac = "1"
    Geral.ValorTrans = formataValor(Geral.rstDoctos!Valor)
    Geral.TipoConta = "C"
    Geral.IdentDep = Mid(Geral.rstDoctos!leitura, 12, 6)
    Geral.AgenciaVinculo = Mid(Geral.rstDoctos!leitura, 4, 4)
    Geral.ContaVinculo = Mid(Geral.rstDoctos!leitura, 23, 7)
    
    If Mid(Geral.rstDoctos!leitura, 1, 3) = "230" Then
      'pesquisa tabela depara 230x409
       If Not Depara230() Then
          Exit Sub
       End If
    Else
      'pesquisa tabela depara 409x409 (antiga)
       DePara
    End If
         
    If (Geral.PreparouLog = 1) Then
        Exit Sub
    End If
   
   'stored procedure do SAQUE
    Geral.hsSQLa = "exec saqccchi "
      
   'monta header
    MontaHeader
   
   'monta parte variavel da 1ºperna do log
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.rstDoctos!CodCenape
    Geral.hsSQLa = Geral.hsSQLa & ", 1"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.ContaCob
    Geral.hsSQLa = Geral.hsSQLa & ", ' '"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.IdentDep
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorTrans)
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.AgenCob
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Mid(Geral.rstDoctos!leitura, 1, 3)
    Geral.hsSQLa = Geral.hsSQLa & ", 2"                                   '
    
   'executa fracassada (6F01)
    LocalLog "Saque Cheque (6F01)- SP " & Geral.hsSQLa
    Set Geral.rst = UBBQuery.ExecuteSQL(Geral.hsSQLa)
    LocalLog "Retorno sp_saque cheque (6F01): " & Format(Geral.rst(0), "00")
    
    If (Val(Geral.rst(0)) = 0) Then
         
        CalculaNSU
        HeaderTx = "BHS1" & Format(Geral.rst(13), "000000") & _
                    Caixa.VersaoAtual & _
                    Format(Parametros.AgenciaCentral, "0000") & _
                    Format(Parametros.AgenciaSatelite, "0000") & "011" & _
                    Format(Caixa.Caixa, "000") & "1" & _
                    Format(Caixa.NSU, "000000") & "0000000" & _
                    Format(Now, "HHMM") & "110000000002"
                   
        MsgIda = HeaderTx & "01" & Format(Geral.AgenCob, "0000") & "0" & _
                 Format(Geral.ContaCob, "0000000000000000") & "0000" & _
                 Format(Geral.IdentDep, "000000000") & String(13, "0") & _
                 String(5, "0") & Parametros.DataServer & "00" & _
                 Format(Geral.rstDoctos!Valor * 100, "0000000000000000") & _
                 String(16, "0") & Format(Geral.rstDoctos!Valor * 100, "0000000000000000") & _
                 Format(Now, "HHMM") & "0" & "0000" & "0000000000" & String(4, "0") & _
                 String(37, "0") & String(6, "0") & String(6, "0") & "02"
        
       'Envia 1ª mensagem ao Host
        TamIda = Format(Len(Trim(MsgIda)), "0000")
        MsgRetorno = String(1921, " ")
        Funcao = "1" & TamIda & "1921****"
        
        LocalLog MsgIda
        
       'Envia BHS1
        Call Abrelinha("BHS1")
        RetQX = qxhostnt(Funcao, MsgIda, MsgRetorno)
        Call FechaLinha("BHS1")
                   
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'se retorno = 20,41,71 -> o micro deverá ser reinicializado '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If (RetQX = 20) Or (RetQX = 41) Or (RetQX = 71) Then
             LocalLog "Retorno BHS1: " & Mid(MsgRetorno, 58, 2)
             MsgBox "Atenção. Ocorreu um erro de comunicação com o Servidor da Agência. Reinicialize este equipamento. Retorno DLL = " & Format(RetQX, "00"), vbOKOnly + vbCritical, "Atenção"
             Close #20
             End
        End If
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Retorno = 21,30,43,47,52,62,80 -> tentar novamente.'
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        If (RetQX = 21) Or (RetQX = 30) Or (RetQX = 43) Or (RetQX = 47) Or _
           (RetQX = 52) Or (RetQX = 62) Or (RetQX = 80) Then
        
            Vez = 1
            Do
               
                Espera (5 * Vez)
                     
               'tentar novamente
                Call Abrelinha("RE-envio BHS1")
                RetQX = qxhostnt(Funcao, MsgIda, MsgRetorno)
                Call FechaLinha("RE-envio BHS1")
                
                Vez = Vez + 1
            
            Loop Until Not ((RetQX = 21) Or (RetQX = 30) Or (RetQX = 43) Or _
                            (RetQX = 47) Or (RetQX = 52) Or (RetQX = 62) Or (RetQX = 80)) And Vez <> 4
        
        End If
        
        LocalLog "Retorno BHS1: " & Format(RetQX, "00") & "MsgRetorno: " & MsgRetorno
        
       'Recebe retorno da Consulta do Host BHS1
        If (RetQX = 0) Then
            
           'Recebe resposta no BHS2 / Grava concretizada 6001
            If Mid(MsgRetorno, 58, 2) = "00" Then
            
                Geral.CodTransacao = "6001"
                Geral.IndTransac = " "
                Geral.TipoTransacao = 1
                              
               'stored procedure do SAQUE
                Geral.hsSQLa = "exec saqccchi "
           
                Parametros.AgenciaSatelite = Geral.rstCapa!agorig
           
                Geral.Hora = Format(Now, "HHMM")
                                                  
               'O caixa só será aberto qdo estacao local com caixa fechado
                LogAberturaCaixa
                
                Geral.hsSQLa = Geral.hsSQLa & "  '" & Geral.CodTransacao & "'"
                Geral.hsSQLa = Geral.hsSQLa & ", " & Format(Geral.rst(13), "000000")
                Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.VersaoAtual
                Geral.hsSQLa = Geral.hsSQLa & ", " & Parametros.AgenciaCentral
                Geral.hsSQLa = Geral.hsSQLa & ", " & Parametros.AgenciaSatelite
                Geral.hsSQLa = Geral.hsSQLa & ", 0"
                Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.TipoTransacao
                Geral.hsSQLa = Geral.hsSQLa & ", 3"
                Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.Caixa
                Geral.hsSQLa = Geral.hsSQLa & ", 1"
                Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.NSU1
                Geral.hsSQLa = Geral.hsSQLa & ", 0"
                Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.CapaBack
                Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.Hora
                Geral.hsSQLa = Geral.hsSQLa & ", '" & Geral.IndTransac & "'"
                
                If Geral.idEnvMal = "E" Then
                   Geral.hsSQLa = Geral.hsSQLa & ", 6"
                Else
                   Geral.hsSQLa = Geral.hsSQLa & ", 7"
                End If
                
                Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.TpRep
                Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.Evento
                Geral.hsSQLa = Geral.hsSQLa & ", '" & Geral.Capa
                Geral.hsSQLa = Geral.hsSQLa & "', '" & Caixa.CIF
                Geral.hsSQLa = Geral.hsSQLa & "', '" & Caixa.SDV
                
               'monta parte variavel
                Geral.hsSQLa = Geral.hsSQLa & "', " & Geral.rstDoctos!CodCenape
                Geral.hsSQLa = Geral.hsSQLa & ", " & Val(Geral.TipoConta)
                Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.ContaCob
                Geral.hsSQLa = Geral.hsSQLa & ", ' '"
                Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.IdentDep
                Geral.hsSQLa = Geral.hsSQLa & ", 0"
                Geral.hsSQLa = Geral.hsSQLa & ", 0"
                Geral.hsSQLa = Geral.hsSQLa & ", 0"
                Geral.hsSQLa = Geral.hsSQLa & ", 0"
                Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorTrans)
                Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.AgenCob
                Geral.hsSQLa = Geral.hsSQLa & ", " & Mid(MsgRetorno, 5, 6)
                Geral.hsSQLa = Geral.hsSQLa & ", 0"
                Geral.hsSQLa = Geral.hsSQLa & ", 0"
                Geral.hsSQLa = Geral.hsSQLa & ", 0"
                Geral.hsSQLa = Geral.hsSQLa & ", 0"
                Geral.hsSQLa = Geral.hsSQLa & ", " & Mid(Geral.rstDoctos!leitura, 1, 3)
                Geral.hsSQLa = Geral.hsSQLa & ", 2"                           '
                
                If GetDocumentoTransmitido(EnumCheque) Then
                    Exit Sub
                End If
                                
               'gravar o nsu desta transação antes de envia-la para o UBB-NT
                spRetorno = MDIQuery.updNsuDocto(Geral.DataProcessamento, _
                                                 Geral.rstCapa!idcapa, _
                                                 Geral.rstDoctos!iddocto, _
                                                 Caixa.NSU1, _
                                                 Caixa.Caixa)
                If spRetorno <> 0 Then
                    MsgBox "5595. ATENÇÃO! Saque Local a ser enviado Log, não atualizado o NSU. ", vbOKOnly + vbCritical, "Atenção"
                    Exit Sub
                End If
                
               'Executa concretizada
                LocalLog "Saque Cheque (6001)- SP " & Geral.hsSQLa
                Set Geral.rst = UBBQuery.ExecuteSQL(Geral.hsSQLa)
                LocalLog "Retorno sp_saque cheque (6001): " & Format(Geral.rst(0), "00")
        
                If (Val(Geral.rst(0)) = 0) Then
                    
                    CalculaNSU
              
                    HeaderTx = "BHS3" & Format(Geral.rst(13), "000000") & _
                    Caixa.VersaoAtual & Format(Parametros.AgenciaCentral, "0000") & _
                    Format(Parametros.AgenciaSatelite, "0000") & "011" & Format(Caixa.Caixa, "000") & _
                    "1" & Format(Caixa.NSU, "000000") & "000000" & "0" & Format(Now, "HHMM") & "110000000002"
                    
                    MsgIda = HeaderTx & Mid(MsgRetorno, 60, 1) & Mid(MsgRetorno, 61, 6) & _
                    Format(Geral.rst(13), "000000") & "1" & Format(Geral.AgenCob, "0000")
              
                    TamIda = Format(Len(Trim(MsgIda)), "0000")
                    MsgRetorno = String(1921, " ")
                    Funcao = "1" & TamIda & "1921****"
            
                    LocalLog MsgIda
                    
                   'envia BHS3 (confirmação para o Host)
                    Call Abrelinha("BHS3")
                    RetQX = qxhostnt(Funcao, MsgIda, MsgRetorno)
                    Call FechaLinha("BHS3")
                    
                    If RetQX <> 0 Then
                        Err.Raise 963, App.Title, "Falha na Confirmação de Pagamento (BHS3)"
                    End If
                    
                   '''''''''''''''''''''''''''''''''''''''''''''''''''
                   ' Procedure para atualizar cheque como já enviado '
                   '''''''''''''''''''''''''''''''''''''''''''''''''''
                    spRetorno = MDIQuery.updDoctoTransmitido(Geral.DataProcessamento, _
                                                             Geral.rstCapa!idcapa, _
                                                             Geral.rstDoctos!iddocto, _
                                                             Caixa.NSU1, _
                                                             Caixa.Caixa, "N")
                    If spRetorno <> 0 Then
                        MsgBox "211. ATENÇÃO! Documento - cheque UBB saque interagência - já enviado Log, não foi atualizado no SQL. Tecle <enter> para continuar.", vbOKOnly + vbCritical, "Atenção"
                    End If
                    
                   'libera linha
                    Geral.GereiLog = 1
                           
                Else
                                                           
                    Call DevolveDocumentos
        
                End If
            
            ElseIf Mid(MsgRetorno, 58, 2) = "03" Then
                                                    
               'Pagto de Conta/Titulo com cheque Ubb com insuficiencia de saldo
                Geral.CodOcorrencia = 213
                Geral.Transacao = Mid(MsgRetorno, 84, 31)
                                      
                LogOcorrencia
                
                Call DevolveDocumentos
                
               'Atualiza saldo do cheque qdo s/ saldo
                spRetorno = MDIQuery.updChequeSaldo(Geral.DataProcessamento, _
                                                    Geral.rstDoctos!iddocto, _
                                                    MsgRetorno)
                
                If spRetorno <> 0 Then
                   MsgBox "Ocorreu algum erro na gravação do saldo para ocorrencia de cheque. Tecle <enter> para continuar.", vbOKOnly + vbCritical, "Atenção"
                End If
                
                Exit Sub
            
            Else
                
                If UCase(Trim(Mid(MsgRetorno, 84, 31))) = "CONTA SEM SALDO" Then
                        
                    Geral.CodOcorrencia = 223
                    Geral.Transacao = Mid(MsgRetorno, 84, 31)
                                            
                    LogOcorrencia
                    Call DevolveDocumentos
                    
                   'Atualiza saldo do cheque qdo s/ saldo
                    spRetorno = MDIQuery.updChequeSaldo(Geral.DataProcessamento, _
                                                        Geral.rstDoctos!iddocto, _
                                                        MsgRetorno)
                    
                    If spRetorno <> 0 Then
                        MsgBox "ATENÇÃO !!! (94) Erro na Atualizacao de saldo insuficiente para SAQUE INTERAGENCIA. Tecle <enter> para continuar.", vbOKOnly + vbCritical, "Atenção"
                    End If
                    
                    Exit Sub
                
                Else
                                        
                    Select Case UCase(Trim(Mid(MsgRetorno, 84, 31)))
                    
                    Case "CONTA UNIBANCO NAO EXISTE"    ' (06) **Conta Unibanco nao existe
                       Geral.CodOcorrencia = 216
                       
                    Case "TIPO CONTA UNIBANCO INVALIDA" ' (06) **Conta Unibanco nao existe
                       Geral.CodOcorrencia = 216
                       
                    Case "BLOQUEIO RESOLUCAO 2025"      ' (10) Bloqueio Resolucao 2025
                       Geral.CodOcorrencia = 217
                       
                    Case "CONTA UNIBANCO INVALIDA"      ' (12) **Conta Unibanco Invalida
                        Geral.CodOcorrencia = 216
                        
                    Case "TRANSACAO CANCELADA"          ' (36) **Transação Cancelada
                        Geral.CodOcorrencia = 999
                        Geral.RetTransacao = 79
                        
                    Case "CHEQUE BLOQUEADO"             ' (72) Cheque bloqueado
                       Geral.CodOcorrencia = 228
                       
                    Case "CONTA SEM SALDO"              ' (83) Conta sem Saldo
                       Geral.CodOcorrencia = 213
                       
                    Case "SALDO BLOQUEADO"              ' (84) Saldo Bloqueado
                       Geral.CodOcorrencia = 223
                       
                    Case "CHEQUE SUSTADO"               ' (92) Cheque sustado
                       Geral.CodOcorrencia = 229
                       
                    Case "CONTA ENCERRADA"              ' (94) Conta Encerrada
                       Geral.CodOcorrencia = 216
                       
                    Case "CONTA BLOQUEADA"              ' (95) Conta Bloqueada
                       Geral.CodOcorrencia = 222
                       
                    Case "RESIDE EXTERIOR"              ' (96) Reside Exterior
                       Geral.CodOcorrencia = 218
                       
                    Case "CONTA PARALISADA"             ' (97) Conta Paralisada
                       Geral.CodOcorrencia = 219
                       
                    Case "ADIANTAMENTO DEPOSITANTE"     ' (99) Adiantamento depositante
                       Geral.CodOcorrencia = 221
                       
                    Case "CONTA INATIVA-BLOQUEADA"       '(01) Conta Bloqueada por Inatividade
                       Geral.CodOcorrencia = 222
                       
                    Case "HORARIO EXPIRADO"                 '(01)
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 1
                       
                    Case "PROBLEMAS NA CONSIST. DO CP"      '(01)
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 2
                       
                    Case "CONTA CONTABIL"                   '(01)
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 3
                       
                    Case "CONTA UNIBANCO NAO ABERTA"        '(01)
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 5
                       
                    Case "NRO DE CHEQUE INIC. INVALIDO"     '(01)
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 57
                     
                    Case "NRO DE CHEQUE FINAL INVALIDO"     '(01)
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 58
                    
                    Case "AGENCIA NACIONAL INVALIDA"        '(01)
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 59
                    
                    Case "C/C NACIONAL INVALIDA"            '(01)
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 60
                     
                    Case "EXCEDEU LIMITE"                   '(01)
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 62
                     
                    Case "DEP. OUTRA PRACA"                 '(01)
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 64
                     
                    Case "CONTA COM RESTRITIVO(S)"          '(01)
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 66
                     
                    Case "CONTATE AGEN TITULAR DA CONTA"    '(01)
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 67
                     
                    Case "VALOR SAQUE RAPIDO EXCEDIDO"      '(01)
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 69
                     
                    Case "DEP INT DISP.AAG"                 '(01)
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 71
                     
                    Case "INTERAGENCIA NAO PERMITIDO"       '(01)
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 73
                     
                    Case "SALDO BLOQUEADO"                  '(01)
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 74
                       
                    Case Else                               'nao tratado
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 51
                       
                    End Select
                  
                    If Geral.CodOcorrencia < 998 Then
                       Geral.Transacao = Mid(MsgRetorno, 84, 31)
                       LogOcorrencia
                    End If
                                            
                    Call DevolveDocumentos
                    Exit Sub
           
                End If
                
            End If
        
        Else
                    
           'Falhou BHS1
            LocalLog "Retorno da função da DLL para o envio da BHS1-Retorno: " & Str(RetQX) & " - Msgretorno: " & MsgRetorno
            
            Geral.CodOcorrencia = 999
            Geral.RetTransacao = 51
            
            Call DevolveDocumentos
                
        End If
   
    Else
        
       'Falhou Fracassada devolve com retorno padrao da mesma
        Call DevolveDocumentos
        
    End If

    Exit Sub
    
TrataErro:
    
    If Err.Number = 964 Or Err.Number = 965 Then
       'Erro na Abertura/Fechamento de Linha
        Geral.CodOcorrencia = 999
        Geral.RetTransacao = 98
        Call DevolveDocumentos
        Exit Sub
    ElseIf Err.Number = 963 Then
       'Erro Subida da confirmaçao do Saque (BHS3)
        Geral.CodOcorrencia = 999
        Geral.RetTransacao = 97
        Call DevolveDocumentos
        Exit Sub
    End If
    
    Select Case TratamentoErro("Falha no Cheque UBB Intergencia.", Err, eDoctoSubidaLog)
        Case eSair
            End
        Case eRepetir
            Resume
        Case eContinuar
            Resume Next
        Case eFinalizar
            Exit Sub
    End Select

End Sub
Sub BHVC(RstMDI_Cobrancas As Recordset, RstMDI_Cobranca As Recordset)
    
On Error GoTo TrataErro

    Dim Funcao          As String * 14
    Dim MsgIda          As String
    Dim MsgRetorno      As String
    Dim HeaderTx        As String
    Dim TamIda          As String
    Dim dataAAAA        As String
    Dim RetQX           As Integer
    Dim Diferenca       As Double
    Dim RstMDI          As Recordset
    Dim Vez             As Integer

    Select Case RstMDI_Cobrancas!TipoDocto
        Case 13     'cobrança registrada sem CB
        
            Geral.CodCVTCob = CLng(RstMDI_Cobranca!CVT)
            Geral.NossoNumCob = Format(RstMDI_Cobranca!NossoNumero, String(15, "0"))
            Geral.VencCob = CLng(Mid(RstMDI_Cobranca!Vecto, 7, 2) & Mid(RstMDI_Cobranca!Vecto, 5, 2) & Mid(RstMDI_Cobranca!Vecto, 3, 2))
            
            Geral.ValorCob = Format(RstMDI_Cobranca!ValorBase * 100, "0000000000000000")
            Geral.ValorMora = "0000000000000000"
            Geral.ValorDesc = "0000000000000000"
            Geral.ValorAbat = "0000000000000000"
            Geral.ValorTrans = Format(RstMDI_Cobrancas!Valor * 100, "0000000000000000")

        Case 14     'cobrança especial sem CB
            Geral.CodCVTCob = CLng(RstMDI_Cobranca!CVT)
            Geral.NossoNumCob = Format(RstMDI_Cobranca!NossoNumero, String(15, "0"))
            Geral.VencCob = CLng(Mid(RstMDI_Cobranca!Vecto, 7, 2) & Mid(RstMDI_Cobranca!Vecto, 5, 2) & Mid(RstMDI_Cobranca!Vecto, 3, 2))
            
            Geral.ValorCob = Format(RstMDI_Cobranca!ValorBase * 100, "0000000000000000")
            Geral.ValorMora = Format(RstMDI_Cobranca!Juros * 100, "0000000000000000")
            Geral.ValorDesc = Format(RstMDI_Cobranca!Desconto * 100, "0000000000000000")
            Geral.ValorAbat = Format(RstMDI_Cobranca!Abatimento * 100, "0000000000000000")
            Geral.ValorTrans = Format(RstMDI_Cobrancas!Valor * 100, "0000000000000000")
        
        Case 28     'cobrança registrada com CB
            Geral.CodCVTCob = 55395
            Geral.NossoNumCob = "001" & Mid(RstMDI_Cobrancas!leitura, 33, 12)
            Geral.VencCob = CLng(Mid(RstMDI_Cobranca!Vecto, 7, 2) & Mid(RstMDI_Cobranca!Vecto, 5, 2) & Mid(RstMDI_Cobranca!Vecto, 3, 2))
            
            Geral.ValorCob = Format(RstMDI_Cobranca!ValorBase * 100, "0000000000000000")
            Geral.ValorTrans = Format(RstMDI_Cobrancas!Valor * 100, "0000000000000000")
            Geral.ValorMora = Format(RstMDI_Cobranca!Juros * 100, "0000000000000000")
            Geral.ValorDesc = Format(RstMDI_Cobranca!Desconto * 100, "0000000000000000")
            Geral.ValorAbat = Format(RstMDI_Cobranca!Abatimento * 100, "0000000000000000")

        Case 30     'cobrança especial com CB
            Select Case Mid(RstMDI_Cobrancas!leitura, 20, 1)
                Case "1"
                    Geral.CodCVTCob = 77330
                    Geral.BHAceitaCobranca = 1  'neste caso, não precisa fazer consulta
                    Exit Sub
                Case "2"
                    Geral.CodCVTCob = 77437
                    Geral.BHAceitaCobranca = 1   'neste caso, não precisa fazer consulta
                    Exit Sub
                Case "3"
                    Geral.CodCVTCob = 77445
                Case "5"
                    Geral.CodCVTCob = 77445
            End Select
         
           'NOSSO NUMERO
            If (Mid(RstMDI_Cobrancas!leitura, 33, 12) = "000000000000") Then
                Geral.NossoNumCob = "0000000000" & Mid(RstMDI_Cobrancas!leitura, 26, 4)
                Modulo11 Geral.NossoNumCob
                Geral.NossoNumCob = Geral.NossoNumCob & Geral.RetDigMod11
            Else
                Geral.NossoNumCob = Mid(RstMDI_Cobrancas!leitura, 30, 15)
            End If
        
             Geral.VencCob = CLng(Mid(RstMDI_Cobranca!Vecto, 7, 2) & Mid(RstMDI_Cobranca!Vecto, 5, 2) & Mid(RstMDI_Cobranca!Vecto, 3, 2))
             
             Geral.ValorCob = Format(RstMDI_Cobranca!ValorBase * 100, "0000000000000000")
             Geral.ValorTrans = Format(RstMDI_Cobrancas!Valor * 100, "0000000000000000")
             Geral.ValorMora = Format(RstMDI_Cobranca!Juros * 100, "0000000000000000")
             Geral.ValorDesc = Format(RstMDI_Cobranca!Desconto * 100, "0000000000000000")
             Geral.ValorAbat = Format(RstMDI_Cobranca!Abatimento * 100, "0000000000000000")

    End Select
    
   '''''''''''''''''
   ' Enviar a BHVC '
   '''''''''''''''''
         
    If Val(Mid(Geral.VencCob, 5, 2)) > 50 Then
       dataAAAA = Mid(Format(Geral.VencCob, "000000"), 1, 4) + "19" + Mid(Format(Geral.VencCob, "000000"), 5, 2)
    Else
       dataAAAA = Mid(Format(Geral.VencCob, "000000"), 1, 4) + "20" + Mid(Format(Geral.VencCob, "000000"), 5, 2)
    End If
     
    HeaderTx = "BHVC" & "000000" & Caixa.VersaoAtual & _
               Format(Parametros.AgenciaCentral, "0000") & _
               Format(Parametros.AgenciaSatelite, "0000") & "000" & _
               Format(Caixa.Caixa, "000") & "1" & "000000" & _
               "000000" & "0" & Format(Now, "HHMM") & "1" & "3" & "0000000000"
               
    MsgIda = HeaderTx & Geral.CodCVTCob & Geral.NossoNumCob & _
             Geral.ValorTrans & Geral.ValorMora & Geral.ValorDesc & _
             Geral.ValorAbat & Geral.ValorCob & dataAAAA & "U"
     
   'Envia 1ª mensagem ao Host
    TamIda = Format(Len(Trim(MsgIda)), "0000")
    MsgRetorno = String(1921, " ")
    Funcao = "1" & TamIda & "1921****"
     
    LocalLog MsgIda
        
   'Envia BHVC
    Call Abrelinha("BHVC")
    RetQX = qxhostnt(Funcao, MsgIda, MsgRetorno)
    Call FechaLinha("BHVC")
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'se retorno = 20,41,71 -> o micro deverá ser reinicializado '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    If (RetQX = 20) Or (RetQX = 41) Or (RetQX = 71) Then
        LocalLog "Retorno BHS2: " & Mid(MsgRetorno, 58, 2)
        MsgBox "Atenção. Ocorreu um erro de comunicação com o Servidor da Agência. Reinicialize este equipamento. Retorno DLL = " & Format(RetQX, "00"), vbOKOnly + vbCritical, "Atenção"
        Close #20
        End
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    'retorno = 21,30,43,47,52,62,80 -> tentar novamente.'
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    If (RetQX = 21) Or (RetQX = 30) Or (RetQX = 43) Or (RetQX = 47) Or _
        (RetQX = 52) Or (RetQX = 62) Or (RetQX = 80) Then
    
        Vez = 1
        Do
     
            '1º RE-ENVIO - 20 SEGUNDOS
             Espera (5 * Vez)
                  
             
            'tentar novamente
             Call Abrelinha("Re-Enviao BHVC")
             RetQX = qxhostnt(Funcao, MsgIda, MsgRetorno)
             Call FechaLinha("Re-Envio BHVC")
             
             Vez = Vez + 1
        
        Loop Until Not ((RetQX = 21) Or (RetQX = 30) Or (RetQX = 43) Or (RetQX = 47) Or _
                        (RetQX = 52) Or (RetQX = 62) Or (RetQX = 80)) And Vez <> 4
    End If
        
    LocalLog "Retorno BHS1: " & Format(RetQX, "00") & "MsgRetorno: " & MsgRetorno
            
   'Retorno da BHVC
    If (RetQX = 0) Then
                
       'Gravação dos dados retornados na tabela documento (Ok ou não OK)
        Call MDIQuery.updAtualizaOcorrenciaBHVR(Geral.DataProcessamento, RstMDI_Cobrancas!iddocto, Mid(MsgRetorno, 62, 120))
                                   
       'Tratamento da resposta no BHVR
        If Mid(MsgRetorno, 58, 2) = "00" Then
            Geral.BHAceitaCobranca = 1
            
        ElseIf Mid(MsgRetorno, 58, 2) = "01" And Mid(MsgRetorno, 62, 19) = "BASES INDISPONIVEIS" Then
            Geral.BHAceitaCobranca = 0
            
            Geral.CodOcorrencia = 999
            Geral.RetTransacao = "56"
            
        ElseIf Mid(MsgRetorno, 58, 2) = "01" Then
            Geral.BHAceitaCobranca = 0
            
            Geral.CodOcorrencia = 208
            Geral.RetTransacao = 0
            Geral.Transacao = "TITULO VENCIDO"
            Call MDIQuery.insLog(Geral.DataProcessamento, Geral.rstCapa!idcapa, RstMDI_Cobranca!iddocto, Caixa.UsuarioAtual, "123")
            Call LogOcorrencia(RstMDI_Cobranca!iddocto)

        End If
        
    Else
    
        Geral.BHAceitaCobranca = 0
        Geral.CodOcorrencia = 999
        Geral.RetTransacao = "51"

    End If
    
    Exit Sub
    
TrataErro:

    Screen.MousePointer = 0
    
    Select Case TratamentoErro("Não foi possível finalizar a transação de consulta de Cobranças Unibanco.", Err)
        Case eSair
            End
        Case eRepetir
            Resume
        Case eContinuar
            Resume Next
    End Select

End Sub
Sub ConsultaCobrancaUBB(pVinculo As Double)

On Error GoTo TrataErro

    Dim RstMDI_Cobrancas    As Recordset
    Dim RstMDI_Cobranca     As Recordset
    Dim RstMDI_Parametro    As Recordset
    Dim RstMDI              As Recordset
    Dim spRetorno           As Integer
    Dim nTabela             As String
    Dim tbAtuCaixa          As Integer
    Dim qtdADCC             As Integer
    Dim qtdChqOutros        As Integer
    Dim qtdChqUBB           As Integer
    Dim qtdLancInt          As Integer

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' BHAceitaCobranca = 0 -> devolver cobrança + cheque                                    '
    ' BHAceitaCobranca = 1 -> consulta OK, processar cobrança + cheque normalmente          '
    ' BHAceitaCobranca = 2 -> processar saque normalmente, pois não há consulta a ser feita '
    ' BHAceitaCobranca = 3 -> devolver cobrança + cheque (Rejeitado pelo Robo)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

   'verifica se existem cobranças Ubb que são pagas com este cheque/adcc
    Set RstMDI_Cobrancas = MDIQuery.getConsultaCobrancasUBB(Geral.DataProcessamento, _
                                                            Geral.rstCapa!idcapa, _
                                                            pVinculo)
    If RstMDI_Cobrancas.EOF = True Then
       'não há cobrança UBB para este pagamento
        Geral.BHAceitaCobranca = 2
        Exit Sub
    End If
    
    Do While Not RstMDI_Cobrancas.EOF
    
        Geral.BHAceitaCobranca = 0
    
        Select Case RstMDI_Cobrancas!TipoDocto
              
           Case 13
                'cobrança registrada sem CB
                Set RstMDI_Cobranca = MDIQuery.getConsultaCobrancaUBB("GetCobrancaRegistrada", _
                                                                       Geral.DataProcessamento, _
                                                                       RstMDI_Cobrancas!iddocto)
                                                        
                nTabela = "CobrancaRegistrada"
               
            Case 14
               'cobrança especial sem CB
                Set RstMDI_Cobranca = MDIQuery.getConsultaCobrancaUBB("GetCobrancaEspecial", _
                                                                      Geral.DataProcessamento, _
                                                                      RstMDI_Cobrancas!iddocto)

                nTabela = "CobrancaEspecial"
            Case 28, 30
                'cobrança registrada com CB,cobrança especial com CB
                Set RstMDI_Cobranca = MDIQuery.getConsultaCobrancaUBB("GetCobrancaCodBar", _
                                                                       Geral.DataProcessamento, _
                                                                       RstMDI_Cobrancas!iddocto)

                nTabela = "FichaCompensacao"
        End Select
        
        If RstMDI_Cobranca.EOF Then
        
            Geral.BHAceitaCobranca = 0
            Geral.CodOcorrencia = 999
            Geral.RetTransacao = 52
            
           'Mandar capa para CSP
            Geral.PreparouLog = 4
     
            If Not DbRejeitaDocto(Geral.rstCapa!idcapa, RstMDI_Cobranca!iddocto) Then
                Screen.MousePointer = 0
                MsgBox "Falha na exclusão de Documento. Tecle <enter> para continuar.", vbOKOnly + vbCritical, "Atenção"
            End If
      
            LocalLog "BHVC - Dados especificos da cobranca UBB nao localizados. Devolvendo Cobranca. " & Trim(Geral.CodOcorrencia) & " / " & Trim(Geral.RetTransacao)
            
           'spRetorno = MDIQuery.updDoctoPendente(Geral.DataProcessamento, Geral.rstCapa!idcapa)
            
            Exit Do
        Else
        
            Set RstMDI_Parametro = MDIQuery.getParametro(Geral.DataProcessamento)
                 
            If TituloExcedePrazoVC(nTabela, _
                                   RstMDI_Cobranca!Vecto, _
                                   RstMDI_Parametro!PrazoVencimento_Mal, _
                                   RstMDI_Parametro!PrazoVencimento_Env) Then

               'montar dados de cada cobrança para envio da BHVC
                Call BHVC(RstMDI_Cobrancas, RstMDI_Cobranca)
                
                If Geral.BHAceitaCobranca <> 1 Then
                                    
                   'Mandar capa para CSP
                    Geral.PreparouLog = 4
                    
                    If Not DbRejeitaDocto(Geral.rstCapa!idcapa, RstMDI_Cobranca!iddocto) Then
                        Screen.MousePointer = 0
                        MsgBox "Falha na exclusão de Documento. Tecle <enter> para continuar.", vbOKOnly + vbCritical, "Atenção"
                    End If
                    
                    LocalLog "BHVC - Devolvendo Cobranca devido retorno BHVC. " & Trim(Geral.CodOcorrencia) & " / " & Trim(Geral.RetTransacao)
                   'spRetorno = MDIQuery.updDoctoPendente(Geral.DataProcessamento, Geral.rstCapa!idcapa)
                    
                    Exit Do
  
                End If
                                
            Else
            
                Geral.BHAceitaCobranca = 2
                
            End If
            
        End If
        
        RstMDI_Cobrancas.MoveNext
        
    Loop

    Exit Sub

TrataErro:

    Screen.MousePointer = 0
    Select Case TratamentoErro("Falha no módulo: [Consulta de cobrança] .", Err)
        Case eSair
            End
        Case eRepetir
            Resume
        Case eContinuar
            Resume Next
    End Select
End Sub
Sub LogLiTemp()
   
'Função para converter LI para ADCC durante testes
'p/ nao comprometer rotinas com LI's sem contrapartida.

On Error GoTo TrataErro

   '====================================='
   ' TRANSAÇÃO 04 (nosso número) - ADCC  '
   '====================================='
   
    Dim RstUBB          As Recordset
    Dim RstMDI          As Recordset
    Dim spRetorno       As Integer
    Dim MsgIda          As String
    Dim MsgRetorno      As String
    Dim HeaderTx        As String
    Dim TamIda          As String
    Dim Funcao          As String * 14
    Dim RetQX           As Integer
    Dim Vez             As Integer
            
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   'verifica se ADCC está pagando alguma cobrança Ubb vencida e se esta pode ser paga '
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
    Geral.Vinculo = Geral.rstDoctos!Vinculo
   
    If Geral.BHAceitaCobranca = 0 Then
      'Geral.PreparouLog = 2    'Vai para novo select
       Exit Sub
    End If
   
   'variaveis do header
    Geral.CodTransacao = "0F15"
    Geral.Evento = 580
    Geral.TipoTransacao = 1
    Geral.Capa = GetCapa(Geral.idEnvMal)
    Geral.IndTransac = "1"
    Geral.ValorTrans = formataValor(Geral.rstDoctos!Valor)
    Geral.TipoConta = "C"
    Geral.AgenciaVinculo = 44
    Geral.ContaVinculo = 1000007
    
   'pesquisa tabela depara
    DePara
   
    If (Geral.PreparouLog = 1) Then
        Exit Sub
    End If
      
   'stored procedure do SAQUE
    Geral.hsSQLa = "exec avintpar "
      
   'monta header
    MontaHeader
   
   'monta parte variavel da 1ª perna do log
    Geral.hsSQLa = Geral.hsSQLa & ", 2"
    Geral.hsSQLa = Geral.hsSQLa & ", 1841"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.AgenCob
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.ContaCob
    Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorTrans)
    Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.rstDoctos!CodCenape
    Geral.hsSQLa = Geral.hsSQLa & ", " & Val(Parametros.DataServer)
    Geral.hsSQLa = Geral.hsSQLa & ", ' '"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    Geral.hsSQLa = Geral.hsSQLa & ", 1"
    Geral.hsSQLa = Geral.hsSQLa & ", ' '"
    Geral.hsSQLa = Geral.hsSQLa & ", 0"
    
   'executa Fracassada (0F15)
    LocalLog "Aut.Debito (0F15) - SP " & Geral.hsSQLa
    Set Geral.rst = UBBQuery.ExecuteSQL(Geral.hsSQLa)
    LocalLog "Retorno saque aut.debito(0F15): " & Format(Geral.rst(0), "00")
    
    If (Val(Geral.rst(0)) = 0) Then
                 
        CalculaNSU
        HeaderTx = "BHS1" & Format(Geral.rst(13), "000000") & _
                    Caixa.VersaoAtual & _
                    Format(Parametros.AgenciaCentral, "0000") & _
                    Format(Parametros.AgenciaSatelite, "0000") & "011" & _
                    Format(Caixa.Caixa, "000") & "1" & _
                    Format(Caixa.NSU, "000000") & "0000000" & _
                    Format(Now, "HHMM") & "110000000002"
        
        MsgIda = HeaderTx & "05" & Format(Geral.AgenCob, "0000") & _
                "0" & Format(Geral.ContaCob, "0000000000000000") & _
                "0000" & Format(Geral.IdentDep, "000000000") & _
                 String(13, "0") & String(5, "0") & _
                 Parametros.DataServer & "00" & _
                 Format(Geral.rstDoctos!Valor * 100, "0000000000000000") & _
                 String(16, "0") & Format(Geral.rstDoctos!Valor * 100, "0000000000000000") & _
                 Format(Now, "HHMM") & "0" & "0000" & "0000000000" & String(4, "0") & _
                 String(37, "0") & String(6, "0") & String(6, "0") & "02"
           
       'Envia 1ª mensagem ao Host
        TamIda = Format(Len(Trim(MsgIda)), "0000")
        MsgRetorno = String(1921, " ")
        Funcao = "1" & TamIda & "1921****"
        
        LocalLog MsgIda
        
       'Envia BHS1
        Call Abrelinha("BHS1")
        RetQX = qxhostnt(Funcao, MsgIda, MsgRetorno)
        Call FechaLinha("BHS1")
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'se retorno = 20,41,71 -> o micro deverá ser reinicializado '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If (RetQX = 20) Or (RetQX = 41) Or (RetQX = 71) Then
             LocalLog "Retorno BHS1: " & Mid(MsgRetorno, 58, 2)
             MsgBox "Atenção. Ocorreu um erro de comunicação com o Servidor da Agência. Reinicialize este equipamento. Retorno DLL = " & Format(RetQX, "00"), vbOKOnly + vbCritical, "Atenção"
             Close #20
             End
        End If
        
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''
       ' Retorno = 21,30,43,47,52,62,80 -> tentar novamente.'
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        If (RetQX = 21) Or (RetQX = 30) Or (RetQX = 43) Or (RetQX = 47) Or _
           (RetQX = 52) Or (RetQX = 62) Or (RetQX = 80) Then
        
            Vez = 1
            Do
               
                Espera (5 * Vez)
                     
               'tentar novamente
                Call Abrelinha("RE-envio BHS1")
                RetQX = qxhostnt(Funcao, MsgIda, MsgRetorno)
                Call FechaLinha("RE-envio BHS1")
                
                Vez = Vez + 1
            
            Loop Until Not ((RetQX = 21) Or (RetQX = 30) Or (RetQX = 43) Or _
                            (RetQX = 47) Or (RetQX = 52) Or (RetQX = 62) Or (RetQX = 80)) And Vez <> 4
        
        End If
        
        LocalLog "Retorno BHS1: " & Format(RetQX, "00") & "MsgRetorno: " & MsgRetorno
        
       'Recebe retorno da Consulta do Host BHS1
        If (RetQX = 0) Then
        
           'Recebe resposta no BHS2, Grava concretizada 6001
            If Mid(MsgRetorno, 58, 2) = "00" Then
                                  
               'variaveis do header
                Geral.CodTransacao = "0015"
                Geral.Evento = 580
                Geral.TipoTransacao = 1
                Geral.Capa = GetCapa(Geral.idEnvMal)
                Geral.IndTransac = " "
                Geral.ValorTrans = formataValor(Geral.rstDoctos!Valor)
                
               'monta header
                Parametros.AgenciaSatelite = Geral.rstCapa!agorig
                
                Geral.Hora = Format(Now, "HHMM")
                     
               'stored procedure do aviso debito
                Geral.hsSQLa = "exec avintpar "
                              
               ''O caixa só será aberto qdo estacao local com caixa fechado
                LogAberturaCaixa
                
                Geral.hsSQLa = Geral.hsSQLa & "  '" & Geral.CodTransacao & "'"
                Geral.hsSQLa = Geral.hsSQLa & ", " & Format(Geral.rst(13), "000000")
                Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.VersaoAtual
                Geral.hsSQLa = Geral.hsSQLa & ", " & Parametros.AgenciaCentral
                Geral.hsSQLa = Geral.hsSQLa & ", " & Parametros.AgenciaSatelite
                Geral.hsSQLa = Geral.hsSQLa & ", 0"
                Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.TipoTransacao
                Geral.hsSQLa = Geral.hsSQLa & ", 3"
                Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.Caixa
                Geral.hsSQLa = Geral.hsSQLa & ", 1"
                Geral.hsSQLa = Geral.hsSQLa & ", " & Caixa.NSU1
                Geral.hsSQLa = Geral.hsSQLa & ", 0"
                Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.CapaBack
                Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.Hora
                Geral.hsSQLa = Geral.hsSQLa & ", '" & Geral.IndTransac & "'"
                
                If Geral.idEnvMal = "E" Then
                    Geral.hsSQLa = Geral.hsSQLa & ", 6"
                Else
                    Geral.hsSQLa = Geral.hsSQLa & ", 7"
                End If
                
                Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.TpRep
                Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.Evento
                Geral.hsSQLa = Geral.hsSQLa & ", '" & Geral.Capa & "'"
                Geral.hsSQLa = Geral.hsSQLa & ", '" & Caixa.CIF
                Geral.hsSQLa = Geral.hsSQLa & "', '" & Caixa.SDV

               'monta parte variavel
                Geral.hsSQLa = Geral.hsSQLa & "', 2"
                Geral.hsSQLa = Geral.hsSQLa & ", 1841"
                Geral.hsSQLa = Geral.hsSQLa & ", 0"
                Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.AgenCob
                Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.ContaCob
                Geral.hsSQLa = Geral.hsSQLa & ", " & CVar(Geral.ValorTrans)
                Geral.hsSQLa = Geral.hsSQLa & ", " & Geral.rstDoctos!CodCenape
                Geral.hsSQLa = Geral.hsSQLa & ", " & Val(Parametros.DataServer)
                Geral.hsSQLa = Geral.hsSQLa & ", ' '"
                Geral.hsSQLa = Geral.hsSQLa & ", " & Mid(MsgRetorno, 5, 6)
                Geral.hsSQLa = Geral.hsSQLa & ", 1"
                Geral.hsSQLa = Geral.hsSQLa & ", ' '"
                Geral.hsSQLa = Geral.hsSQLa & ", 0"
                
                If GetDocumentoTransmitido(EnumADCC) Then
                    Exit Sub
                End If
                
               'gravar o nsu desta transação antes de envia-la para o UBB-NT
                spRetorno = MDIQuery.updNsuDocto(Geral.DataProcessamento, _
                                                 Geral.rstCapa!idcapa, _
                                                 Geral.rstDoctos!iddocto, _
                                                 Caixa.NSU1, _
                                                 Caixa.Caixa)
                                       
                If spRetorno <> 0 Then
                    MsgBox "5435. ATENÇÃO! Aut.Debito a ser enviado Log, não atualizado o NSU. ", vbOKOnly + vbCritical, "Atenção"
                    Exit Sub
                End If
                
               'Executa concretizada
                LocalLog "Aut. Debito (0015) - SP " & Geral.hsSQLa
                Set Geral.rst = UBBQuery.ExecuteSQL(Geral.hsSQLa)
                LocalLog "Retorno saque aut.debito(0015): " & Format(Geral.rst(0), "00")
                
                If (Val(Geral.rst(0)) = 0) Then
                    
                    CalculaNSU
              
                    HeaderTx = "BHS3" & Format(Geral.rst(13), "000000") & Caixa.VersaoAtual & Format(Parametros.AgenciaCentral, "0000") & Format(Parametros.AgenciaSatelite, "0000") & "011" & Format(Caixa.Caixa, "000") & "1" & Format(Caixa.NSU, "000000") & "000000" & "0" & Format(Now, "HHMM") & "110000000002"
                    MsgIda = HeaderTx & Mid(MsgRetorno, 60, 1) & Mid(MsgRetorno, 61, 6) & Format(Geral.rst(13), "000000") & "1" & Format(Geral.AgenCob, "0000")
                    
                    TamIda = Format(Len(Trim(MsgIda)), "0000")
                    MsgRetorno = String(1921, " ")
                    Funcao = "1" & TamIda & "1921****"
                                  
                    LocalLog MsgIda
              
                   'envia BHS3 (confirmação para o Host)
                    Call Abrelinha("BHS3")
                    RetQX = qxhostnt(Funcao, MsgIda, MsgRetorno)
                    Call FechaLinha("BHS3")
                    
                    If RetQX <> 0 Then
                        Err.Raise 963, App.Title, "Falha na Confirmação de Pagamento (BHS3)"
                    End If
                                         
                    If Not (IsNull(Geral.rstDoctos!RetornoTransacao) And _
                            Geral.rstDoctos!RetornoTransacao = 75) Then
                        
                       'Conta Unibanco reinformada Corretamente
                        MDIQuery.updCancelarRetornoTransacao Geral.DataProcessamento, _
                                                             Geral.rstCapa!idcapa, _
                                                             Geral.rstDoctos!iddocto, _
                                                             Caixa.Caixa
                    End If
                                            
                   ''''''''''''''''''''''''''''''''''''''''''''''''
                   ' Procedure para atualizar o adcc como enviado '
                   ''''''''''''''''''''''''''''''''''''''''''''''''
                   
                    spRetorno = MDIQuery.updDoctoTransmitido(Geral.DataProcessamento, _
                                                             Geral.rstCapa!idcapa, _
                                                             Geral.rstDoctos!iddocto, _
                                                             Caixa.NSU1, _
                                                             Caixa.Caixa, "N")
                                        
                    If spRetorno <> 0 Then
                        MsgBox "590. ATENÇÃO! Documento - débito automático - já enviado Log, não atualizado no SQL. ", vbOKOnly + vbCritical, "Atenção"
                    End If
                                            
                   'libera linha
                    Geral.GereiLog = 1
              
                Else
                    Call DevolveDocumentos
                    Exit Sub
                End If
            
            ElseIf Mid(MsgRetorno, 58, 2) = "03" Then
               'Pagto de Conta/Titulo com cheque Ubb com insuficiencia de saldo
                Geral.CodOcorrencia = 429
                Geral.Transacao = Mid$(MsgRetorno, 84, 31)
                LogOcorrencia
                Call DevolveDocumentos
                
               'Atualiza Tabela ADCC Saldo/Limite/Vinculado
                spRetorno = MDIQuery.updADCCSaldo(Geral.DataProcessamento, _
                                                  Geral.rstDoctos!iddocto, _
                                                  MsgRetorno)
        
                If spRetorno <> 0 Then
                   MsgBox "Ocorreu algum erro na gravação do saldo para ocorrencia de ADCC. Tecle <enter> para continuar.", vbOKOnly + vbCritical, "Atenção"
                End If
                
                Exit Sub
            
            Else
                                
                Select Case UCase(Trim(Mid(MsgRetorno, 84, 31)))
                
                    Case "CONTA UNIBANCO NAO EXISTE"    ' (06) **Conta Unibanco nao existe
                       Geral.CodOcorrencia = 414        ' alteração solicitada por Selma em 30/11/99
                       
                    Case "TIPO CONTA UNIBANCO INVALIDA" ' (06) **Conta Unibanco nao existe
                       Geral.CodOcorrencia = 414
                       
                    Case "BLOQUEIO RESOLUCAO 2025"      ' (10) Bloqueio Resolucao 2025
                       Geral.CodOcorrencia = 420
                       
                    Case "CONTA UNIBANCO INVALIDA"      ' (12) **Conta Unibanco Invalida
                       Geral.CodOcorrencia = 414        ' alteração solicitada por Selma em 30/11/99
                       
                    Case "TRANSACAO CANCELADA"          ' (36) **Transação Cancelada
                       Geral.CodOcorrencia = 0

                    Case "CONTA SEM SALDO"              ' (83) Conta sem Saldo
                       
                      'Atualiza Tabela ADCC Saldo/Limite/Vinculado
                       spRetorno = MDIQuery.updADCCSaldo(Geral.DataProcessamento, _
                                                         Geral.rstDoctos!iddocto, _
                                                         MsgRetorno)
                    
                       If spRetorno <> 0 Then
                          MsgBox "ATENÇÃO !!! (96)Erro na Atualizacao de saldo insuficiente para ADCC. Tecle <enter> para continuar.", vbOKOnly + vbCritical, "Atenção"
                       End If
                    
                       Geral.CodOcorrencia = 429
                       
                    Case "SALDO BLOQUEADO"              ' (84) Saldo Bloqueado
                       Geral.CodOcorrencia = 426
                       
                    Case "CONTA ENCERRADA"              ' (94) Conta Encerrada
                       Geral.CodOcorrencia = 419
                       
                    Case "CONTA BLOQUEADA"              ' (95) Conta Bloqueada
                       Geral.CodOcorrencia = 425
                       
                    Case "RESIDE EXTERIOR"              ' (96) Reside Exterior
                       Geral.CodOcorrencia = 421
                       
                    Case "CONTA PARALISADA"             ' (97) Conta Paralisada
                       Geral.CodOcorrencia = 422
                       
                    Case "ADIANTAMENTO DEPOSITANTE"     ' (99) Adiantamento depositante
                       Geral.CodOcorrencia = 424
                       
                    Case "CONTA INATIVA-BLOQUEADA"       '(01) Conta Bloqueada por Inatividade
                       Geral.CodOcorrencia = 425
                    
                    Case "HORARIO EXPIRADO"       '(01)
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 1
                       
                    Case "PROBLEMAS NA CONSIST. DO CP"       '(01)
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 2
                       
                    Case "CONTA CONTABIL"       '(01)
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 3
                       
                    Case "CONTA UNIBANCO NAO ABERTA"       '(01)
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 5
                       
                    Case "AGENCIA NACIONAL INVALIDA"       '(01)
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 59
                    
                    Case "C/C NACIONAL INVALIDA"       '(01)
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 60
                     
                    Case "EXCEDEU LIMITE"       '(01)
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 62
                     
                     Case "DEP. OUTRA PRACA"       '(01)
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 64
                     
                    Case "CONTA COM RESTRITIVO(S)"       '(01)
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 66
                     
                    Case "CONTATE AGEN TITULAR DA CONTA"       '(01)
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 67
                     
                    Case "VALOR SAQUE RAPIDO EXCEDIDO"       '(01)
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 69
                     
                    Case "DEP INT DISP.AAG"       '(01)
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 71
                     
                    Case "INTERAGENCIA NAO PERMITIDO"       '(01)
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 73
                     
                    Case "SALDO BLOQUEADO"       '(01)
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 74
                       
                    Case Else                           ' nao tratado
                       Geral.CodOcorrencia = 999
                       Geral.RetTransacao = 51
                End Select
                               
                If Geral.CodOcorrencia = 414 And _
                   (IsNull(Geral.rstDoctos!RetornoTransacao) Or _
                    Geral.rstDoctos!RetornoTransacao <> 75) Then
                    
                    Call GaugeTitulo(1)
                                        
                    Geral.RetTransacao = 75
                    Geral.CodOcorrencia = "0"
            
                    LocalLog "Agencia/Conta Invalida sendo enviada para correcao de agencia/conta"
                    Call DbRejeitaDocto(Geral.rstCapa!idcapa, Geral.rstDoctos!iddocto, ST_DoctoCorrecaoAgConta)
                    
                   'Capa para correcao de Agencia/Conta
                    Geral.PreparouLog = 5
                    Espera (0.5)
                Else
                                  
                   'grava ocorrencia
                    If Geral.CodOcorrencia > 0 And Geral.CodOcorrencia < 998 Then
                        Geral.Transacao = Mid$(MsgRetorno, 84, 31)
                        LogOcorrencia
                    End If
                    
                    LocalLog "Agencia/Conta Reinformada Invalida sendo enviada para CSP"
                    
                    Call DevolveDocumentos
                
                End If
                
                Exit Sub
        
            End If
        Else
            
           'Falhou BHS1
            LocalLog "Retorno da função da DLL para o envio da BHS1-Retorno: " & Str(RetQX) & " - Msgretorno: " & MsgRetorno

            Select Case RetQX
                Case 21, 30, 43, 47, 52, 62, 80, 33, 36, 42, 48
                    LogAjusteDebitoADCC
                Case Else
                    Geral.CodOcorrencia = 999
                    Geral.RetTransacao = 51
                    Call DevolveDocumentos
            End Select
            
            Exit Sub
            
        End If
        
    Else
    
       'Falhou Fracassada devolve com retorno padrao da mesma
        Call DevolveDocumentos
        
    End If
   
    Exit Sub
    
TrataErro:
    LocalLog "Falha no ADCC " & Err.Description

    If Err.Number = 964 Or Err.Number = 965 Then
       'Erro na Abertura/Fechamento de Linha
        Geral.CodOcorrencia = 999
        Geral.RetTransacao = 98
        Call DevolveDocumentos
        Exit Sub
    ElseIf Err.Number = 963 Then
       'Erro Subida da confirmaçao do Saque (BHS3)
        Geral.CodOcorrencia = 999
        Geral.RetTransacao = 97
        Call DevolveDocumentos
        Exit Sub
    End If
    
    Select Case TratamentoErro("Falha no ADCC.", Err, eDoctoSubidaLog)
        Case eSair
            End
        Case eRepetir
            Resume
        Case eContinuar
            Resume Next
        Case eFinalizar
            Exit Sub
    End Select

End Sub
Function ProcessaRecepcao() As Boolean

    '********************************************'
    ' Recepciona Capa no Ik, se Caixa Autorizado '
    '********************************************'
    
Dim spRetorno   As Integer

    If frmShow.CheckRecepcao = 0 Then Exit Function

    Set Geral.rstCapa = MDIQuery.getCapaRecepcao(Geral.DataProcessamento, spRetorno)
    
    If spRetorno = 0 Then
    
        Geral.RecebendoCapa = True
        
        If AgenciaCadastradaBH() Then
        
            If Geral.rstCapa!idEnv_Mal = "F" Then
                frmShow.LabelStatus.Caption = "Status: Recepção Fininvest..."
                Espera (0.25)
                Call GaugeInit("Capa " & Format(Geral.rstCapa!Capa, "00000000000"), 4, Fininvest)
            Else
                frmShow.LabelStatus.Caption = "Status: Recepção Malote/Envelope..."
                Espera (0.25)
                Call GaugeInit("Capa " & Format(Geral.rstCapa!Capa, "00000000000"), 4, Recepcao)
            End If
      
            Call GaugePos(Recepcao)
            
            Espera (0.1)
            
            Call LogRecepcao
       
            If Geral.CaixaAberto Then
                LogFechamentoCaixa ("A")
                Geral.CaixaAberto = False
            End If
            
            If Geral.FecharCaixa Then
                Call GetCaixa
                Geral.FecharCaixa = False
            End If
        Else
        
            Call MDIQuery.updCapaRecepcionada(Geral.DataProcessamento, Geral.rstCapa!idcapa, "P")
            
        End If
        
        ProcessaRecepcao = True
        Geral.RecebendoCapa = False
        
    End If


End Function

