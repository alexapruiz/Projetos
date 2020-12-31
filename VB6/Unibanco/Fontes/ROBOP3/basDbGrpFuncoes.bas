Attribute VB_Name = "basDbGrpFuncoes"
'=================================='
' Funções de Banco de Dados Gerais '
'=================================='
Option Explicit
Public Function DbRejeitaDocto(ByVal pIdcapa As Long, ByVal pIdDocto As Long, Optional ByVal pStatus As String = ST_DoctoDeletadaRobo) As Boolean
On Error GoTo TrataErro

    Dim spRetorno As Integer
                               
    'Atualizar Dados do Documento
     spRetorno = MDIQuery.updDoctoRejeitado(Geral.DataProcessamento, _
                                            pIdcapa, _
                                            pIdDocto, _
                                            Geral.CodOcorrencia, _
                                            Geral.RetTransacao, _
                                            pStatus)
     If spRetorno = 0 Then
         DbRejeitaDocto = True
     Else
         Screen.MousePointer = 0
         MsgBox "Houve alguma Falha na atualização do docto. Tecle <enter> para continuar.", vbOKOnly + vbCritical, "Atenção"
     End If
     
     LocalLog "Rejeita Docto(Ocorrencia:" & Trim(Geral.CodOcorrencia) & "/" & Trim(Geral.RetTransacao) & ") - AtualizaCaixa "
     
    'Atualiza referencia na Tabela Caixa
    
'     spRetorno = MDIQuery.updCaixadocto(Geral.DataProcessamento, Caixa.Caixa)
'     If spRetorno = 0 Then
'         DbRejeitaDocto = True
'     Else
'         DbRejeitaDocto = False
'         Screen.MousePointer = 0
'         MsgBox "Falha na atualização do docto, na Tabela Caixa. Tecle <enter> para continuar.", vbOKOnly + vbCritical, "Atenção"
'     End If
     
     Exit Function

TrataErro:

    Screen.MousePointer = 0
    Select Case TratamentoErro("Falha no módulo: [DbRejeitaDocto].", Err)
        Case eSair
            End
        Case eRepetir
            Resume
        Case eContinuar
            Resume Next
    End Select
            
End Function
Sub IniOpcoes()
On Error GoTo TrataErro

    'Fecha caixa
     If AntIniOpcoes.FechaCx Then
         SaveSetting appname:="Robo", section:="Caixa", Key:="Aberto", setting:=0
     End If
     
    'Limpa tabela Caixa
     If AntIniOpcoes.ClearCapaCx Then
         Call MDIQuery.updCaixaCapa(Geral.DataProcessamento, Caixa.Caixa)
     End If
     
'     If AntIniOpcoes.ClearDoctoCX Then
'         Call MDIQuery.updCaixadocto(Geral.DataProcessamento, Caixa.Caixa)
'     End If
     
     Exit Sub
     
TrataErro:

    Select Case TratamentoErro("Falha no módulo: [IniOpções].", Err)
        Case eSair
            End
        Case eRepetir
            Resume
        Case eContinuar
            Resume Next
    End Select

End Sub
Sub AlteraParaCompensacao()

On Error GoTo TrataErro

    Dim spRetorno As Integer
    
    spRetorno = MDIQuery.updAlteraTipoDocto(Geral.DataProcessamento, Geral.rstDoctos!iddocto, "6")
    
    If spRetorno <> 0 Then
       MsgBox "ATENÇÃO !!! Falha na SP [ updAlteraTipoDocto ]. Tecle <enter> para continuar.", vbOKOnly + vbCritical, "Atenção"
    End If
    LocalLog "altera para compesanção: " & Trim(Geral.rstDoctos!iddocto)
    Exit Sub
    
TrataErro:

Screen.MousePointer = 0
    Select Case TratamentoErro("Falha no módulo: [AlteraParaCompensacao].", Err)
        Case eSair
            End
        Case eRepetir
            Resume
        Case eContinuar
            Resume Next
    End Select

End Sub
Function GetCapa(ByVal pTipoCapa As String) As String

    If InStr(1, "EF", pTipoCapa, 1) <> 0 Then
        GetCapa = Format(Geral.rstCapa!Capa, "00000000000")
    Else
        GetCapa = Format(Geral.rstCapa!Num_malote, "00000000000")
    End If
 
End Function
Function AtualizaEstorno(ByVal EstornoAutorizado As Boolean)

On Error GoTo TrataErro:

    Dim spRetorno As Integer
    
    spRetorno = MDIQuery.updStatusEstorno(Geral.DataProcessamento, _
                                          Geral.rstCapa!idcapa, _
                                          Geral.rstCapa!iddocto, _
                                          IIf(EstornoAutorizado, CodigoDoctoEstornado, CodigoDoctoNoEstorno))
                                             
    If spRetorno <> 0 Then
       MsgBox "Ocorreu algum erro com a atualizacao do estorno realizado para esta capa. Tecle <enter> para continuar.", vbOKOnly + vbCritical, "Atenção"
    End If
    
    Exit Function
    
    LocalLog "Atualiza Estorno - SUCESSO: " & IIf(EstornoAutorizado, "Sim", "Não")
    
TrataErro:

    Select Case TratamentoErro("Falha no módulo: [AtualizaEstorno].", Err)
        Case eSair
            End
        Case eRepetir
            Resume
        Case eContinuar
            Resume Next
    End Select
    
End Function
Function VerificaAgenciaConta() As Boolean

On Error GoTo TrataErro

    Dim RstUbbAgConta       As Recordset
    Dim RstMdiDepositos     As Recordset
    Dim spRetorno           As Integer
    Dim TipoConta           As String * 1

    Dim Rst1, Rst2, Rst3    As Recordset
    
    LocalLog "Valida Ag/Conta de: Depositos e OCT"
       
    Set Rst1 = Geral.rstDoctos.Clone(adLockReadOnly)
    Set Rst2 = Geral.rstDoctos.Clone(adLockReadOnly)
    Set Rst3 = Geral.rstDoctos.Clone(adLockReadOnly)
    
    Rst1.Filter = "Tipodocto like 2"
    Rst2.Filter = "Tipodocto like 3"
    Rst3.Filter = "Tipodocto like 39"
    
    If Rst1.EOF And Rst2.EOF And Rst3.EOF Then
        Exit Function
    End If
    
    Set RstMdiDepositos = MDIQuery.getAgenciaContaDosDepositos(Geral.DataProcessamento, Geral.rstCapa!idcapa, spRetorno)
    If spRetorno <> 0 Then Exit Function
    
    Do While Not RstMdiDepositos.EOF
        
        If RstMdiDepositos!TipoConta = 5 Then         'OCT
            TipoConta = "O"
        ElseIf RstMdiDepositos!TipoConta = 2 Then     'Deposito CP
            TipoConta = "P"
        Else                                           'Deposito CC
            TipoConta = "C"
        End If
                    
        Set RstUbbAgConta = UBBQuery.getAgenciaConta(RstMdiDepositos!Agencia, _
                                                     RstMdiDepositos!Conta, _
                                                     TipoConta)
                    
       'se nao encontrar a agencia/conta do deposito
        If RstUbbAgConta.EOF() And _
           (IsNull(RstMdiDepositos!RetornoTransacao) Or _
           RstMdiDepositos!RetornoTransacao <> 75) Then
                       
            Call GaugeTitulo(1)
            
            Geral.RetTransacao = 75
            Geral.CodOcorrencia = "0"
            
            Call DbRejeitaDocto(Geral.rstCapa!idcapa, RstMdiDepositos!iddocto, ST_DoctoCorrecaoAgConta)
            
           'Capa para correcao de Agencia/Conta
            Geral.PreparouLog = 5
            VerificaAgenciaConta = True
            Espera (0.5)
            
            LocalLog "Agencia/Conta Invalida (" & Format(RstMdiDepositos!Agencia, "0000") & _
                     "/" & Format(RstMdiDepositos!Conta, "0000000") & _
                     ") sendo enviada para correcao de agencia/conta"
                        
        ElseIf RstUbbAgConta.EOF() And _
               Not (IsNull(RstMdiDepositos!RetornoTransacao) Or _
               RstMdiDepositos!RetornoTransacao <> 75) Then
               
            Call GaugeTitulo(4)
        
            Geral.CodOcorrencia = 999
           
            If RstMdiDepositos!TipoDocto = 3 Then
                Geral.RetTransacao = 46
            Else
                Geral.RetTransacao = 48
            End If
           
            If Not DbRejeitaDocto(Geral.rstCapa!idcapa, RstMdiDepositos!iddocto) Then
                Screen.MousePointer = 0
                MsgBox "Falha na exclusão de Documento. Tecle <enter> para continuar.", vbOKOnly + vbCritical, "Atenção"
            End If
          
           'capa para CSP
            Geral.PreparouLog = 4
            VerificaAgenciaConta = True
           
            LocalLog "Agencia/Conta Invalida reinformada (" & _
                     Format(RstMdiDepositos!Agencia, "0000") & "/" & _
                     Format(RstMdiDepositos!Conta, "0000000") & _
                    ") pelo modulo de correcao"
                    
        ElseIf Not RstUbbAgConta.EOF() And IIf(IsNull(RstMdiDepositos!RetornoTransacao), 0, RstMdiDepositos!RetornoTransacao) = 75 Then
        
            Call MDIQuery.updCancelarRetornoTransacao(Geral.DataProcessamento, Geral.rstCapa!idcapa, RstMdiDepositos!iddocto, Caixa.Caixa)
        
        End If
    
        RstMdiDepositos.MoveNext
    
    Loop
    
Exit Function

TrataErro:

    Select Case TratamentoErro("Falha no módulo: [VerificaAgenciaConta].", Err)
        Case eSair
            End
        Case eRepetir
            Resume
        Case eContinuar
            Resume Next
    End Select

End Function
Function VerificaCapaOcorrencia(ByVal pNovaCapa As Integer) As Boolean

On Error GoTo TrataErro

    Dim spRetorno           As Integer
    Dim RstMDI              As Recordset
    Dim Ocorrencia          As Integer
    Dim ErroDescricao       As String
    Static OcorrenciaAntiga
    
    LocalLog "Verifica se ha Docto com Ocorrencia"
    
    If Geral.rstDoctos!Status = "D" Then
     
        If Mid(Geral.rstDoctos!Ocorrencia, 1, 3) <> 999 And _
           Mid(Geral.rstDoctos!Ocorrencia, 1, 3) <> 998 And _
           OcorrenciaAntiga <> Geral.rstDoctos!Ocorrencia Then
                
            OcorrenciaAntiga = Geral.rstDoctos!Ocorrencia
            Geral.CodOcorrencia = Mid(Geral.rstDoctos!Ocorrencia, 1, 3)
        
            Geral.Transacao = ""    'descricao em branco já que o codigo está sendo informado
            LogOcorrencia
            LocalLog "Gravacao IKRO Cod.Ocorrencia: " & Format(Mid(Geral.rstDoctos!Ocorrencia, 1, 3), "000")
        End If
                   
       'atualizar o docto com já enviado ocorrencia, mesmo se for deletado sem ocorrencia
        ErroDescricao = "Atualizada Documento Transmitido"
        spRetorno = MDIQuery.updDoctoTransmitido(Geral.DataProcessamento, _
                                                 Geral.rstCapa!idcapa, _
                                                 Geral.rstDoctos!iddocto, _
                                                 "0", "0", "N")
        If spRetorno <> 0 Then
            MsgBox "211. ATENÇÃO! Documento com ocorrencia enviada não atualizado. Tecle <enter> para continuar.", vbOKOnly + vbCritical, "Atenção"
        End If
        
        VerificaCapaOcorrencia = True
        
    End If
    
    Exit Function

TrataErro:

    Select Case TratamentoErro("Falha em VerificaCapaOcorrencia. Nota: " & ErroDescricao, Err, eCapa)
        Case eSair
            End
        Case eRepetir
            Resume
        Case eContinuar
            Resume Next
        Case eFinalizar
            Exit Function
    End Select

End Function
Function GetChequeCashDeposito(ByVal pIdcapa As Long, _
                               ByVal pVinculo As Long, _
                               ByVal pValor As Double, _
                               ByRef QtdeCheques As Integer, _
                               ByRef ValorDinheiro As String, _
                               ByRef ValorCheque As String, _
                               ByRef ValorSomado As String, _
                               Optional ByVal InProcess As Boolean = True) As Boolean
                              
    On Error GoTo TrataErro

    Dim RstMDI, RstTMP              As Recordset
    Dim ContCheques                 As Integer
    Dim QtdeDinheiro                As Integer
    Dim ValorCheques                As Double
    Dim spRetorno                   As Integer
    Dim LocalVinculo                As Long
    Dim LocalIndex                  As Integer
    GetChequeCashDeposito = True
    
   'InProcess somente leitura dos cheques, para rotina de estorno

   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' Procedure os cheque Deposito (Cheque diversos) ou dinheiro (Cheque UBB / LI) '
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
    Set RstMDI = MDIQuery.GetChequesDeposito(Geral.DataProcessamento, pIdcapa, pVinculo)
                                                             
    If RstMDI.EOF And InProcess Then
    
        Geral.CodOcorrencia = 999
        Geral.Transacao = 94
        GetChequeCashDeposito = False
        
        Call DevolveDocumentos
        Exit Function
    End If
                                                                        
    Do
    
        If RstMDI!TipoDocto = 41 Then
           'para LI continua o processamento normal o log do LI ira depois
            If InProcess Then LocalLog "Deposito em dinheiro LI"
            QtdeDinheiro = QtdeDinheiro + 1
        
       'Sacar cheque para confirmar deposito
        ElseIf RstMDI!TipoDocto = 5 Then
        
            QtdeDinheiro = QtdeDinheiro + 1
            If InProcess Then LocalLog "Deposito em dinheiro Cheque-UBB"
            
        ElseIf RstMDI!TipoDocto = 7 Then
            
            If InProcess Then LocalLog "Deposito em Cheque Documento: " & Str(RstMDI!Valor)
            QtdeCheques = QtdeCheques + 1
            ValorCheques = ValorCheques + RstMDI!Valor
            
        End If
        
        RstMDI.MoveNext
        
    Loop Until RstMDI.EOF
    
    ValorSomado = formataValor(ValorCheques)
    
   'caso a qtd cheques do deposito esteja com 0, excluir este deposito
    If (QtdeDinheiro > 0 And QtdeCheques = 0) Then
        ValorDinheiro = formataValor(pValor)
        ValorCheque = 0
    ElseIf QtdeCheques > 0 And QtdeDinheiro = 0 Then
        ValorDinheiro = 0
        ValorCheque = formataValor(pValor)
    Else
        Geral.PreparouLog = 1
        Exit Function
    End If
    
    LocalLog "Dinheiro: " & QtdeDinheiro & " / " & ValorDinheiro & " - Cheque: " & QtdeCheques & " / " & ValorCheque & ""
    
    Exit Function
    
TrataErro:
    GetChequeCashDeposito = False
    Select Case TratamentoErro("Falha no Deposito C/C.", Err, eDoctoSubidaLog)
        Case eSair
            End
        Case eRepetir
            Resume
        Case eContinuar
            Resume Next
        Case eFinalizar
            Exit Function
    End Select

End Function
Function VerificaAgenciaContaAjustes() As Boolean

On Error GoTo TrataErro

    Dim RstUbbAgConta       As Recordset
    Dim RstMdiAjustes       As Recordset
    Dim spRetorno           As Integer
    Dim TipoConta           As String * 1

    Dim Rst1, Rst2, Rst3, Rst4    As Recordset
    
    LocalLog "Valida Ag/Conta de: Ajustes de Credito/Debito"
       
    Set Rst1 = Geral.rstDoctos.Clone(adLockReadOnly)
    Set Rst2 = Geral.rstDoctos.Clone(adLockReadOnly)
    Set Rst3 = Geral.rstDoctos.Clone(adLockReadOnly)
    Set Rst4 = Geral.rstDoctos.Clone(adLockReadOnly)
    
    Rst1.Filter = "Tipodocto = 32"
    Rst2.Filter = "Tipodocto = 33"
    Rst3.Filter = "Tipodocto = 34"
    Rst4.Filter = "Tipodocto = 38"
    
    If Rst1.EOF And Rst2.EOF And Rst3.EOF And Rst4.EOF Then
        Exit Function
    End If

    Set RstMdiAjustes = MDIQuery.getAgenciaContaAjustes(Geral.DataProcessamento, Geral.rstCapa!idcapa, spRetorno)
    
    If spRetorno = 0 Then
    
        While Not RstMdiAjustes.EOF
                                                                         
            Set RstUbbAgConta = UBBQuery.getAgenciaConta(RstMdiAjustes!Agencia, RstMdiAjustes!Conta)
           
           'se nao encontrar a agencia/conta do ajuste
            If RstUbbAgConta.EOF() Then
                                   
                Call GaugeTitulo(2)
                LocalLog "Agencia/conta de ajuste não cadastradas: " & RstMdiAjustes!Agencia & " / " & RstMdiAjustes!Conta
        
               'Seta Capa para Ilegíveis ou CSP conforme Origem
                If Geral.rstCapa!csp Then
                   'Insere Mensagem p/ Usuario
                    spRetorno = MDIQuery.insMensagem(Geral.DataProcessamento, _
                                                     MSG_AjusteAgenContaInexisCSP, _
                                                     Geral.rstCapa!idcapa, 0, _
                                                     Caixa.Caixa)
                   'CSP
                    Geral.PreparouLog = 4
                    
                Else
                   'Insere Mensagem p/ Usuario
                    spRetorno = MDIQuery.insMensagem(Geral.DataProcessamento, _
                                 MSG_AjusteAgenContaInexis, _
                                 Geral.rstCapa!idcapa, 0, _
                                 Caixa.Caixa)
                   'Ilegiveis
                    Geral.PreparouLog = 6

                End If
                             
                VerificaAgenciaContaAjustes = True
                Exit Function
              
            End If
           
            RstMdiAjustes.MoveNext
                       
        Wend
        
    End If
    
    Exit Function
    
TrataErro:

    Select Case TratamentoErro("Falha no Deposito C/C.", Err, eDoctoSubidaLog)
        Case eSair
            End
        Case eRepetir
            Resume
        Case eContinuar
            Resume Next
        Case eFinalizar
            Exit Function
    End Select



End Function
