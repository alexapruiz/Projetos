Attribute VB_Name = "Recepcao"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''
''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private myrecord   As TpRecordRejeicao

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'         Define type de registro do arquivo de Rejeição           '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Type TpRecordRejeicao
     Rot_Mensag          As String * 6       'Rótulo da mensagem (REJEIC)
     CGC_Ender           As String * 14      'CGC de Endereçamento do contrato de serviço
     Num_Bordero         As String * 18      'Número do borderô
     Dat_Deposito        As String * 8       'Data de Entrega/Depósito (AAAAMMDD)
     Num_BcoCliente      As String * 4       'Número do Banco Cliente
     Num_AgenCliente     As String * 4       'Número da Agência Cliente
     Num_ContaCliente    As String * 7       'Número da Conta Corrente Cliente
     Num_Cheque          As String * 7       'Número do Cheque emitente
     Rot_Original        As String * 6       'Rótulo Original
     Qtd_Erros           As String * 2       'Quantidade de Erros
     Cod_Erros           As String * 114     'Código de Erros (38 ocorrências de 3 bytes cada)
     Tip_Identificacao   As String * 2       'Identificador de Campo (CM ou "  " e demais dados do cheque)
     Cod_CMC7            As String * 34      'Código do CMC7
     Cod_OfiAdv          As String * 25      'Código do OFI/ADV
     CrLf                As String * 2       'OK
End Type

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'         Define type de inconsistência na recepção de arquivos (Módulo Principal)   '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Type tpRecepcao
     Num_Bordero    As String * 19
     DataDeposito   As String * 8       'Data formato (aaaammdd)
     Banco          As String * 4
     Agencia        As String * 4
     Conta          As String * 11
     Cheque         As String * 6
     NossoNumero    As String * 11
     Inconsistencia As String * 50
End Type

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'         Type para inserção de dados na tabela rejeicaoremessa                      '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Type tpRejeicao
    DataProcessamento As Long
    IdBordero         As Long
    DataDeposito      As Long
    IdCheque          As Long
    CodErro           As String * 114
    Rotulo            As String * 6
End Type


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'         Define type de inconsistência na recepção de confirmação (Módulo Principal)     '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Type tpRecepcaoConfir
     Num_Bordero    As String * 19
     Banco          As String * 4
     Agencia        As String * 4
     Conta          As String * 7
     Inconsistencia As String * 50
End Type

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'         Define type de registro do arquivo de Confirmação           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Type RecordConfirmacao
     Rot_Mensag          As String * 6       'Rótulo da mensagem (CONFIR)
     CGC_Ender           As String * 14      'CGC de Endereçamento do Contrato de serviço
     Num_Bordero         As String * 18      'Número do borderô
     Dat_Processamento   As String * 8       'Data de Processamento (AAAAMMDD)
     Num_Banco           As String * 4       'Número do Banco (0409)
     Num_Agencia         As String * 4       'Número da Agência do cliente
     Num_ContaCorrente   As String * 7       'Número da Conta Corrente do cliente
     Reg_Vago            As String * 7       'Registro Vago (Zeros)
     Rot_Original        As String * 6       'Rótulo Original (CHINBO)
     Qtd_Erros           As String * 2       'Quantidade de Erros
     Cod_Erro            As String * 3       'Código do Erro
     Cod_OfiAdv          As String * 25      'Código do OFI/ADV
     CrLf                As String * 2       'OK
     'Controle            As String * 1       'Byte de controle para fim de arquivo (LineFeed)
End Type

'Registro de Cheque Data Boa -  Rótulo CHDBOA
Private Type ChDataBoa_Reg
    CGC_Enderecamento               As String * 14
    Num_Bordero                     As String * 18
    CodigoCarteira                  As String * 2
    DataDeposito                    As String * 8
    Valor                           As String * 13
    CMC7                            As String * 30
    AgenciaCliente                  As String * 4
    ContaCliente                    As String * 7
    TipoCPFCGC                      As String * 2
    CNPJCPF                         As String * 14
    CrLf                            As String * 2
End Type

' Registro de Confirmacao de Baixa- Rotulo CHRBAI
Private Type ChrBai_Reg
    Num_Bordero                    As String * 18
    CodigoCarteira                 As String * 2
    DataDeposito                   As String * 8
    BancoEmitente                  As String * 4
    AgenciaEmitente                As String * 4
    NrChequeEmitente               As String * 6
    NossoNumero                    As String * 11
    CodigoCompensacao              As String * 3
    CcEmitente                     As String * 11
    ValorCheque                    As Currency
    CrLf                           As String * 2
End Type

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'         Define type de registro do arquivo de Instruções (CHRDTV)   '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Type RecordInstrucoes
     Rot_Mensag          As String * 6       'Rótulo da mensagem (CHRDTV)
     CGC_Ender           As String * 14      'CGC de Endereçamento do Contrato de serviço
     Cod_Carteira        As String * 2       'Código da carteira (16-Custódia) / (17-Caução)
     Dat_DepAnterior     As String * 8       'Data de Depósito Anterior (AAAAMMDD)
     Num_Banco           As String * 4       'Número de Banco do emitente
     Num_Agencia         As String * 4       'Número de Agência do emitente
     Num_Cheque          As String * 6       'Número de Cheque do emitente
     Dat_DepNova         As String * 8       'Data de Depósito Nova (AAAAMMDD)
     Cod_NossoNumero     As String * 11      'Nosso Número (Controle do banco)
     Cod_compensacao     As String * 3       'Código de compensação
     Num_ContaCorrente   As String * 11      'Número da Conta Corrente do emitente
     Num_Bordero         As String * 18      'Número do borderô
     ValorCheque         As String * 15      'Valor de Cheque
     CrLf                As String * 2
End Type

' Registro de Aviso de Diferença -  Rótulo CHADIF
Private Type AvisoDif_Reg
    Rot_Mensag                      As String * 6   'Rótulo da mensagem (CHADIF)
    CGC_Ender                       As String * 14  'CGC de Endereçamento do Contrato de serviço
    Num_Bordero                     As String * 18  'Número do borderô
    CodigoCarteira                  As String * 2   'Código da carteira (16-Custódia) / (17-Caução)
    DataOcorrencia                  As String * 8
    CodigoOcorrencia                As String * 9
    Agencia                         As String * 4
    Conta                           As String * 7
    CodigoDevolucao                 As String * 2
    CodigoCompensacao               As String * 3
    BancoEmitente                   As String * 4
    AgenciaEmitente                 As String * 4
    CcEmitente                      As String * 11
    NrChequeEmitente                As String * 10
    TipoCheque                      As String * 1
    TipoInscricao                   As String * 2
    InscricaoEmitente               As String * 14
    DataDeposito                    As String * 8
    Valor                           As String * 13
    CMC7                            As String * 34
    MotivoDevolucao                 As String * 2
    OcorrCHADIF                     As String * 1
    CrLf                            As String * 2
End Type


' Registro da Regra do GP -  Rótulo CHREGP
Private Type RegraGP_Reg
    Rot_Mensag                      As String * 6   'Rótulo da mensagem (CHREPG)
    CGC_Ender                       As String * 14  'CGC de Endereçamento do Contrato de serviço
    DataProcessamento               As String * 8
    CodigoProduto                   As String * 5
    CodigoRegra                     As String * 4
    QtdDias                         As String * 5
    CrLf                            As String * 2
End Type


Private Enum LabelAcao
     Apresenta = 1
     Finaliza
     Atualiza
End Enum

Private Proc_Selecionar As New Custodia.Selecionar
Private Proc_Inserir As New Custodia.Inserir
Private Proc_Atualizar As New Custodia.Atualizar
Private Proc_Excluir As New Custodia.Excluir
'

Private Function LerArquivosDiretorio(ByVal sDiretorio As String, ByVal sRotulo As String, ByRef sArquivos As String) As Boolean
     
     Dim iCount As Integer
     Dim sArquivo As String
     
     Dim myrec As RecordConfirmacao
     
     LerArquivosDiretorio = False
     sRotulo = UCase(sRotulo)
     
     On Error GoTo Err_TrataDiretorio
     Principal.filArquivosRecepcao.path = sDiretorio
     
     On Error GoTo Err_LerArquivosDiretorio
     
     Principal.filArquivosRecepcao.Pattern = "*.*"
     Principal.filArquivosRecepcao.Normal = True
     
     If Principal.filArquivosRecepcao.ListCount > 0 Then

          For iCount = 0 To Principal.filArquivosRecepcao.ListCount
               sArquivo = sDiretorio & Principal.filArquivosRecepcao.List(iCount)
'               If Not (Right(Principal.filArquivosRecepcao.List(iCount), 5) = "_READ" or
               If Not (Right(Principal.filArquivosRecepcao.List(iCount), 4) = "_ERRO" Or _
                    Right(Principal.filArquivosRecepcao.List(iCount), 3) = "_OK") Then
                    
                    myrec.Rot_Mensag = ""
                    
                    Open sArquivo For Random As #1 Len = Len(myrec)
                    Get #1, 1, myrec
     
                    If UCase(myrec.Rot_Mensag) = sRotulo Then
                         'Se arquivo foi aberto para leitura (READ) então retirar extensão
                         If Right(Principal.filArquivosRecepcao.List(iCount), 5) = "_READ" Then
                              sArquivos = sArquivos & Left(Principal.filArquivosRecepcao.List(iCount), _
                                                       Len(Principal.filArquivosRecepcao.List(iCount)) - 5) & ","
                         Else
                              sArquivos = sArquivos & Principal.filArquivosRecepcao.List(iCount) & ","
                         End If
                    End If
                    Close #1
               End If
          Next
     End If
     
     Principal.filArquivosRecepcao.Normal = False
     LerArquivosDiretorio = True
     Exit Function

Err_TrataDiretorio:
     Beep
     If Err.Number = 76 Then
          Close
          Principal.filArquivosRecepcao.Normal = False
          MsgBox "Diretório para recepção não localizado!" & vbCrLf & vbCrLf & "Verifique diretório de recepção em parametros do sistema", vbCritical, App.Title
          Exit Function
     End If

Err_LerArquivosDiretorio:
     'Verifica se arquivo não tem permissão de abertura
     If Err.Number = 70 Or Err.Number = 52 Or Err.Number = 75 Or Err.Number = 76 Then
          Resume Next
     End If
     
     Beep
     Close
     Principal.filArquivosRecepcao.Normal = False
     MsgBox "Erro na leitura do diretório ( " & sDiretorio & " )", vbCritical, App.Title
     
End Function

Public Sub RecConfRemessa()

Dim myrecord             As RecordConfirmacao
Dim lngTotalRegistros    As Long, lngRegs As Long, bArquivoReading As Boolean
Dim strNomeArquivo       As String, strNomeArquivoREA As String, strNomeArquivoOK As String, strNomeArquivoErr As String
Dim rsRetorno            As New ADODB.Recordset
Dim lngRetorno           As Long
Dim bTransacaoAberta     As Boolean
Dim recepcao()           As tpRecepcaoConfir
Dim lngCount             As Long
Dim lngIdBordero         As Long
Dim strDataChave         As String
Dim strHoraChave         As String
Dim strDigBordero        As String      'Dígito verificador do borderô
Dim Progress             As New clsProgressBar
Dim sTituloTela          As String
Dim sArquivoRecep        As String
Dim aArquivoRecep()
Dim iCount               As Integer, aIndex As Integer

On Error GoTo Erro_RecConfRemessa
     
     sTituloTela = "Recepção da Confirmação de remessa"
     
     bTransacaoAberta = False
     
     If MsgBox("Inicializa processo de recepção", vbQuestion + vbYesNo, sTituloTela) = vbNo Then
          Set rsRetorno = Nothing
          Exit Sub
     End If

     If Not LerArquivosDiretorio(g_Parametros.DiretorioRecepcao & "\", "CONFIR", sArquivoRecep) Then
          Set rsRetorno = Nothing
          Exit Sub
     End If
     If Len(sArquivoRecep) = 0 Then
          Beep
          MsgBox "Não existe arquivo(s) para recepção referente à Confirmação de Remessa", vbInformation, sTituloTela
          Set rsRetorno = Nothing
          Exit Sub
     End If
     
     ReDim recepcao(0)
     ReDim aArquivoRecep(0)
     Screen.MousePointer = vbHourglass
     
     Do While True
          iCount = InStr(sArquivoRecep, ",")
          If iCount = 0 Then Exit Do
          aIndex = UBound(aArquivoRecep) + 1
          ReDim Preserve aArquivoRecep(aIndex)
          aArquivoRecep(aIndex) = Mid(sArquivoRecep, 1, (iCount - 1))
          sArquivoRecep = Mid(sArquivoRecep, (iCount + 1))
     Loop

     Call LabelRecepcao(LabelAcao.Apresenta)

     For iCount = 1 To aIndex
     
          Call LabelRecepcao(LabelAcao.Atualiza, iCount, aIndex, aArquivoRecep(iCount))
          
          strNomeArquivo = Trim(g_Parametros.DiretorioRecepcao) & "\" & aArquivoRecep(iCount)
          'Muda extensão do arquivo para (REA) qdo estiver sendo acessado para leitura
          strNomeArquivoREA = strNomeArquivo & "_READ"
          'Muda extensão do arquivo para (OK) qdo finalizado o processo
          strNomeArquivoOK = strNomeArquivo & "_OK"
          'Muda extensão do arquivo para (Erro) qdo ocorrer erro
          strNomeArquivoErr = strNomeArquivo & "_ERRO"
          
          bArquivoReading = False
          
          'Verifica se arq. confirmação existe no diretório e está sendo acessado por outro usuário
          If Dir(strNomeArquivo, vbDirectory) = "" Then
               
               'Verifica se arq. confirmação está sendo lido por outro usuário (.REA)
               If Dir(strNomeArquivoREA, vbDirectory) = "" Then

                    'Abre arquivo para verificar erro de inexistencia do diretório/Arquivo
                    Open strNomeArquivo For Input As #1
                    Close #1
                    Exit Sub
               Else
                    'Força abertura do arquivo para executar tratamento do erro
                    'Caso consiga abrir o arquivo, é sinal de que houve queda de execução
                    'e o arquivo ficou em modo Reading...
                    Open strNomeArquivoREA For Input Access Read As #1
                    Close #1
                    bArquivoReading = True
               End If
          End If
          
          'Renomeia o arquivo para extensão (.REA)
          If Not bArquivoReading Then
               Name strNomeArquivo As strNomeArquivoREA
          End If
          
          'Abre o arquivo CONFIRMAÇÃO DE REMESSA com extensão (.REA)
          Open strNomeArquivoREA For Random As #1 Len = Len(myrecord)
          
          'Obtem o total de bytes do arquivo de confirmação de remessa
          lngTotalRegistros = FileLen(strNomeArquivoREA) / Len(myrecord)
          
         'Inicializa Progress Bar
          Progress.ValorMinimo = 1
          Progress.ValorMaximo = lngTotalRegistros
          Progress.DescricaoProcesso = "Recepcionando confirmação de remessa ..."
          Progress.InicializaProgressBar
          Progress.AtualizaBarra
          
          'Ler cada registro e atualizar tabela de cheque
          For lngRegs = 1 To lngTotalRegistros
               'Ler linha arquivo confirmação de remessa
               Get #1, lngRegs, myrecord
               'Atualiza tabela cheque somente se registro não contém erro
               If CInt(myrecord.Qtd_Erros) = 0 Then
     
                    ' Calcula dígito verificador do borderô
                    ' Para Calcular o digito considerar: Mid(NumeroBordero, 13, 6)
                    
                    
                    strDigBordero = RetornaDigitoModulo11Simplificado(Mid(myrecord.Num_Bordero, 13, 6))
     
                    Set rsRetorno = Nothing
                    'Obtem número do IdBordero
                    Set rsRetorno = g_cMainConnection.Execute(Proc_Selecionar.GetNumeroIdBordero( _
                                                       CLng(myrecord.Dat_Processamento), _
                                                       myrecord.Num_Bordero & strDigBordero), lngRetorno, adCmdText)
                    
                    If rsRetorno.EOF Then
                         Call GeraInconsConfRemessa("Não localizado borderô", myrecord, recepcao, strDigBordero)
                    Else
                         lngIdBordero = rsRetorno(0).Value
                         
                         Set rsRetorno = Nothing
                         
                         'Inicializa transação
                         g_cMainConnection.BeginTrans
                         bTransacaoAberta = True
                         
                         'Atualiza status Borderô correspondente para (E)Confirmado
                         Set rsRetorno = g_cMainConnection.Execute(Proc_Atualizar.AtualizaStatusBordero( _
                                                            CLng(myrecord.Dat_Processamento), _
                                                            lngIdBordero, "E"), lngRetorno, adCmdText)
                         
                         If lngRetorno <> 1 Then
                              Call GeraInconsConfRemessa("Não foi possível atualizar borderô para confirmado", myrecord, recepcao, strDigBordero)
                              
                              'Cancela transação
                              g_cMainConnection.RollbackTrans
                              bTransacaoAberta = False
                         Else
                    
                              Set rsRetorno = Nothing
                              
                              'Atualiza status cheque correspondente para (E)Confirmado
                              Set rsRetorno = g_cMainConnection.Execute(Proc_Atualizar.AtualizaChequesRemessa( _
                                                                 CLng(myrecord.Dat_Processamento), _
                                                                 lngIdBordero))
                              If lngRetorno < 1 Then
                                   Call GeraInconsConfRemessa("Não foi possível atualizar cheque para confirmado", myrecord, recepcao, strDigBordero)
                              
                                   'Cancela transação
                                   g_cMainConnection.RollbackTrans
                                   bTransacaoAberta = False
                              
                              End If
                              
                              
                              Set rsRetorno = Nothing
                              
                              'Atualiza status Data de Depósito correspondente para (1)Confirmado
                              Set rsRetorno = g_cMainConnection.Execute(Proc_Atualizar.AtualizaStatusDataDeposito( _
                                                                 CLng(myrecord.Dat_Processamento), _
                                                                 lngIdBordero))
                              If lngRetorno < 1 Then
                                   Call GeraInconsConfRemessa("Não foi possível atualizar Data de Depósito para confirmada", myrecord, recepcao, strDigBordero)
                              
                                   'Cancela transação
                                   g_cMainConnection.RollbackTrans
                                   bTransacaoAberta = False
                              Else
                                   'Finaliza transação
                                   g_cMainConnection.CommitTrans
                                   bTransacaoAberta = False
                              End If
                              
                              
                         End If
                    End If
               End If
               
               'Atualiza Progress Bar
               Progress.AtualValue = lngRegs
               Progress.AtualizaBarra
          Next
          
          'Fecha arquivo de confirmação
          Close #1

          'Renomeia o arquivo para extensão (.OK)
          Name strNomeArquivoREA As strNomeArquivoOK
     
     Next
     
     ReDim aArquivoRecep(0)
     
     'Verifica se houve ocorrencia na recepção
     If UBound(recepcao) > 0 Then
          
          Set rsRetorno = Nothing
          'Obtem data e hora do servidor (Chave para a tabela INCONSISTENCIA)
          Set rsRetorno = g_cMainConnection.Execute(Proc_Selecionar.GetDataHoraServidor(), lngRetorno, adCmdText)
          
          If lngRetorno <> 0 Then
               Beep
               MsgBox "Problema na geração do relatório de inconsistências!", vbCritical, sTituloTela
               GoTo Sair
          End If
          
          strDataChave = Format(rsRetorno!Data, "yyyymmdd")
          strHoraChave = rsRetorno!Hora
          
          For lngCount = 1 To UBound(recepcao)
               'Insere registro com dados de inconsistência
               Call g_cMainConnection.Execute(Proc_Inserir.InsereInconsistencia( _
                                                  strDataChave, _
                                                  strHoraChave, _
                                                  recepcao(lngCount).Num_Bordero & _
                                                  recepcao(lngCount).Banco & _
                                                  recepcao(lngCount).Agencia & _
                                                  recepcao(lngCount).Conta & _
                                                  recepcao(lngCount).Inconsistencia))

          Next
     
'          'Imprime relação de ocorrências do processo de recepção
'          Call ImprimirInconsistencia("RelConfirmacaoRemessa", "Ocorrências da confirmação de remessa", myrecord.Dat_Processamento, strDataChave, strHoraChave)
          
          'Fecha tela apresentação de arquivos
          Call LabelRecepcao(LabelAcao.Finaliza)
          MsgBox Space(32) & "A T E N Ç Ã O" & vbCrLf & vbCrLf & _
                 "Finalizada a recepção com ocorrência(s).  Verifique !", vbCritical, sTituloTela
                 
          'Remove registro com dados de inconsistência
          Call g_cMainConnection.Execute(Proc_Excluir.RemoveInconsistencia(strDataChave, _
                                                                           strHoraChave), _
                                                                           lngRetorno, adCmdText)
     Else
          
          Call LabelRecepcao(LabelAcao.Finaliza)
          
          MsgBox "Finalizada a confirmação de remessa de cheques.", vbInformation, sTituloTela
     End If
     
     
Sair:
     Call LabelRecepcao(LabelAcao.Finaliza)
     
     Set Progress = Nothing
     
     Screen.MousePointer = vbDefault
     Set rsRetorno = Nothing
     
     Exit Sub
     
     
Erro_RecConfRemessa:

     Beep
     
     'Cancela transação
     If bTransacaoAberta Then g_cMainConnection.RollbackTrans: bTransacaoAberta = False

     'Tratamento para diretório não localizado
     If Err.Number = 76 Then
          MsgBox "Não localizado o diretório com o arquivo de confirmação." & vbCrLf & vbCrLf & _
                    "Favor verificar diretório em parâmetros.", vbCritical, sTituloTela
          GoTo Sair
     End If
     If Err.Number = 53 Then
          MsgBox "Não localizado o arquivo de confirmação." & vbCrLf & vbCrLf & _
                    "Favor verificar e tentar novamente.", vbCritical, sTituloTela
          GoTo Sair
     End If
     If Err.Number = 55 Then
          MsgBox "Arquivo de confirmação em utilização por outro usuário. Favor verificar !", vbCritical, sTituloTela
          GoTo Sair
     End If

     'Fecha arquivo de confirmação
     Close #1

     If UBound(aArquivoRecep) >= 1 Then
          'Renomeia o arquivo para finalização com erro
          Name strNomeArquivoREA As strNomeArquivoErr
     End If
     
     MsgBox Err.Description, vbCritical, sTituloTela
     GoTo Sair

End Sub

Private Sub GeraInconsConfRemessa(strDescrMsg As String, myrecord As RecordConfirmacao, recepcao() As tpRecepcaoConfir, ByVal strDigBordero As String)

Dim lngCount As Long
     
     ReDim Preserve recepcao(UBound(recepcao) + 1)
     lngCount = UBound(recepcao)

     'Acumula em type os borderôs com ocorrência
     recepcao(lngCount).Num_Bordero = myrecord.Num_Bordero & strDigBordero
     recepcao(lngCount).Banco = myrecord.Num_Banco
     recepcao(lngCount).Agencia = myrecord.Num_Agencia
     recepcao(lngCount).Conta = myrecord.Num_ContaCorrente
     recepcao(lngCount).Inconsistencia = strDescrMsg

End Sub
Sub GeraInconsRejeitados(strDescrMsg As String, recepcao() As tpRecepcao, ByVal sBordero As String)
     
    Dim lngCount As Long
     
    ReDim Preserve recepcao(UBound(recepcao) + 1)
    lngCount = UBound(recepcao)
     
   'Acumula em type os borderôs com ocorrência
    recepcao(lngCount).Num_Bordero = Format(sBordero, String(19, "0"))
    recepcao(lngCount).DataDeposito = Geral.DataProcessamento
    recepcao(lngCount).Banco = IIf(myrecord.Tip_Identificacao = "CM", "0" & Mid(myrecord.Cod_CMC7, 2, 3), myrecord.Num_BcoCliente)
    recepcao(lngCount).Agencia = IIf(myrecord.Tip_Identificacao = "CM", Mid(myrecord.Cod_CMC7, 5, 4), myrecord.Num_AgenCliente)
    recepcao(lngCount).Conta = IIf(myrecord.Tip_Identificacao = "CM", Mid(myrecord.Cod_CMC7, 26, 7), myrecord.Num_ContaCliente)
    recepcao(lngCount).Inconsistencia = strDescrMsg

End Sub
Private Sub ImprimirInconsistencia(ByVal strNomeRel As String, strTitulo As String, ByVal strDataProcessamento As String, ByVal strDataChave As String, ByVal strHoraChave As String)

     With Principal.CrystalReport
          .ReportFileName = App.path & "\Reports\" & strNomeRel & ".rpt"
          .Destination = crptToWindow
          .WindowState = crptMaximized
          .WindowTitle = strTitulo
          .Formulas(0) = "DataChave = '" & strDataChave & "'"
          .Formulas(1) = "HoraChave = '" & strHoraChave & "'"
          .Formulas(2) = "DataProcessamento = '" & FormataData(CLng(strDataProcessamento), DD_MM_AAAA) & "'"
          
          .WindowState = crptNormal
          .WindowLeft = 30
          .WindowTop = 30
          .WindowHeight = 700
          .WindowWidth = 950
        
          .Action = 1

          .ReportFileName = Empty
          .Formulas(0) = Empty
          .Formulas(1) = Empty
          .Formulas(2) = Empty
     End With
     
End Sub
Public Sub RecChqDataBoa()
 
    Dim rstChDataBoa        As New ADODB.Recordset
    Dim cCheque             As New CalculoCheque
    Dim DatFile             As Integer
    Dim lRetorno            As Long
    Dim Reg                 As String * 124
    Dim OffSet              As Long
    Dim nCheques            As Long
    Dim CH                  As ChDataBoa_Reg
    Dim bArquivoReading     As Boolean
    Dim sstr                As String
    Dim sDigito             As String
    Dim sNumBordero         As String
    Dim sWhere              As String
    Dim strNomeArquivo      As String
    Dim strNomeArquivoREA   As String
    Dim strNomeArquivoOK    As String
    Dim strNomeArquivoErr   As String
    Dim Progress            As New clsProgressBar
    Dim lngRegs             As Long
    Dim sTituloTela         As String
    Dim sPathName           As String
    Dim sArquivoRecep       As String
    Dim aArquivoRecep()
    Dim iCount              As Integer, aIndex As Integer
    
    On Error GoTo ErroLeitura

     sTituloTela = "Recepção Movimento Data Boa"
     nCheques = 0
        
     If MsgBox("Inicializa processo de recepção", vbQuestion + vbYesNo, sTituloTela) = vbNo Then
          Exit Sub
     End If
    
     If Not LerArquivosDiretorio(g_Parametros.DiretorioRecepcao & "\", "CHDBOA", sArquivoRecep) Then
          Exit Sub
     End If
     If Len(sArquivoRecep) = 0 Then
          Beep
          MsgBox "Não existe arquivo(s) para recepção do Movimento de Data Boa.", vbInformation, sTituloTela
          Exit Sub
     End If
     
     ReDim aArquivoRecep(0)
     Screen.MousePointer = vbHourglass
     
     Do While True
          iCount = InStr(sArquivoRecep, ",")
          If iCount = 0 Then Exit Do
          aIndex = UBound(aArquivoRecep) + 1
          ReDim Preserve aArquivoRecep(aIndex)
          aArquivoRecep(aIndex) = Mid(sArquivoRecep, 1, (iCount - 1))
          sArquivoRecep = Mid(sArquivoRecep, (iCount + 1))
     Loop

     Call LabelRecepcao(LabelAcao.Apresenta)
    
     DatFile = FreeFile

     'Recepciona quantos arquivos existirem
     For iCount = 1 To aIndex
     
          Call LabelRecepcao(LabelAcao.Atualiza, iCount, aIndex, aArquivoRecep(iCount))
          
          strNomeArquivo = Trim(g_Parametros.DiretorioRecepcao) & "\" & aArquivoRecep(iCount)
          'Muda extensão do arquivo para (REA) qdo estiver sendo acessado para leitura
          strNomeArquivoREA = strNomeArquivo & "_READ"
          'Muda extensão do arquivo para (OK) qdo finalizado o processo
          strNomeArquivoOK = strNomeArquivo & "_OK"
          'Muda extensão do arquivo para (Erro) qdo ocorrer erro
          strNomeArquivoErr = strNomeArquivo & "_ERRO"
          
          bArquivoReading = False
          
          'Verifica se arq. confirmação existe no diretório e está sendo acessado por outro usuário
          If Dir(strNomeArquivo, vbDirectory) = "" Then
               
               'Verifica se arq. confirmação está sendo lido por outro usuário (.REA)
               If Dir(strNomeArquivoREA, vbDirectory) = "" Then

                    'Abre arquivo para verificar erro de inexistencia do diretório/Arquivo
                    Open strNomeArquivo For Input As #1
                    Close #1
                    Exit Sub
               Else
                    'Força abertura do arquivo para executar tratamento do erro
                    'Caso consiga abrir o arquivo, é sinal de que houve queda de execução
                    'e o arquivo ficou em modo Reading...
                    Open strNomeArquivoREA For Input Access Read As #1
                    Close #1
                    bArquivoReading = True
               End If
          End If
          
          'Renomeia o arquivo para extensão (.REA)
          If Not bArquivoReading Then
               Name strNomeArquivo As strNomeArquivoREA
          End If
          
          Open strNomeArquivoREA For Binary Access Read Lock Read Write As #DatFile
              
          OffSet = 1
      
          Get #DatFile, OffSet, Reg
          
          'Obtem o total de registros do arquivo de leitura
          If Not EOF(DatFile) Then
                'Inicia progress bar
                Progress.ValorMinimo = 1
                Progress.ValorMaximo = Fix(FileLen(strNomeArquivoREA) / Len(Reg))
                Progress.DescricaoProcesso = "Recepcionando Movimento Data Boa ..."
                Progress.InicializaProgressBar
                Progress.AtualizaBarra
                lngRegs = 0
          End If
          
          While Not EOF(DatFile)
          
              'Acumulador de registros lidos
              lngRegs = lngRegs + 1
          
              ' Se arquivo foi lido ok
              If Len(Reg) < 111 Then
                   MsgBox "Erro de Leitura", vbOKOnly + vbCritical, sTituloTela
                   GoTo FimLeituraComErro
              End If
              
              ' Ver rótulo do arquivo
              If Mid(Reg, 1, 6) <> "CHDBOA" Then
                   MsgBox "Rótulo do Arquivo de Cheques da Data Boa Inválido.", vbOKOnly + vbCritical, sTituloTela
                   GoTo FimLeituraComErro
              End If
              
              ' Ver se CGC de terceira é válido
              If CStr(Mid(Reg, 7, 14)) <> Format(g_Parametros.CNPJ_Terceira, String(14, "0")) Then
                 MsgBox "Endereçamento da Terceira Inválido", vbOKOnly + vbCritical, sTituloTela
                 GoTo FimLeituraComErro
              End If
                      
              ' Atribuir registos do arquivo
              CH.CGC_Enderecamento = Mid(Reg, 7, 14)
              CH.Num_Bordero = Mid(Reg, 21, 18)
              CH.CodigoCarteira = Mid(Reg, 39, 2)
              CH.DataDeposito = Mid(Reg, 41, 8)
              CH.Valor = Mid(Reg, 49, 13)
              CH.CMC7 = Mid(Reg, 63, 8) & Mid(Reg, 72, 10) & Mid(Reg, 83, 12)
              CH.AgenciaCliente = Mid(Reg, 96, 4)
              CH.ContaCliente = Mid(Reg, 100, 7)
              CH.TipoCPFCGC = Mid(Reg, 107, 2)
              CH.CNPJCPF = Mid(Reg, 109, 14)
              
              'Calcula dígito verificador do borderô
              ' Para Calcular o digito considerar: Mid(NumeroBordero, 13, 6)
              
              sDigito = RetornaDigitoModulo11Simplificado(Mid(CH.Num_Bordero, 13, 6))
              sNumBordero = CH.Num_Bordero & Trim(sDigito)
              cCheque.CMC7 = CH.CMC7
              cCheque.Calcula
              
              Set rstChDataBoa = g_cMainConnection.Execute(Proc_Selecionar.GetChequeDataBoa(Geral.DataProcessamento, CH.CMC7))
              
              If rstChDataBoa.EOF Then

                ' Gravar o Cheque da Data Boa
                Call g_cMainConnection.Execute(Proc_Inserir.InsChDataBoa(Geral.DataProcessamento, _
                                               CH.CMC7, CH.CGC_Enderecamento, CLng(CH.DataDeposito), _
                                               sNumBordero, CByte(CH.CodigoCarteira), CH.CNPJCPF, _
                                               CInt(cCheque.Agencia), CDbl(cCheque.Conta), _
                                               Format(Val(InserePonto(CH.Valor)), MASK_VALOR), _
                                               CInt(cCheque.Tipificacao), IIf(Val(CH.Valor) < g_Parametros.ValorChequeLimite, 0, 1), 0, 0, _
                                               0, 0, 0, 0), lRetorno, adCmdText)

                  nCheques = nCheques + lRetorno
                  
              End If
              
              OffSet = OffSet + Len(Reg)
              
              Get #DatFile, OffSet, Reg
              
              'Atualiza Progress Bar
              Progress.AtualValue = lngRegs
              Progress.AtualizaBarra
              
          Wend
      
          Close #DatFile
          
          'Renomeia o arquivo para extensão (.OK)
          Name strNomeArquivoREA As strNomeArquivoOK
          
     Next
     
     'Encerra progress bar
     Set Progress = Nothing
     
     'Fecha tela apresentação de arquivos
     Call LabelRecepcao(LabelAcao.Finaliza)
     
     Screen.MousePointer = vbDefault
     MsgBox "Foram Processados " & CStr(nCheques) & " Cheque(s) da Data Boa.", vbOKOnly + vbExclamation, sTituloTela
     Exit Sub
     
FimLeitura:
     'Fecha tela apresentação de arquivos
     Call LabelRecepcao(LabelAcao.Finaliza)

     Close #DatFile
     'Encerra progress bar
     Set Progress = Nothing
     
     Screen.MousePointer = vbDefault
     Exit Sub
    
FimLeituraComErro:

     Close #DatFile

     If UBound(aArquivoRecep) >= 1 Then
          'Renomeia o arquivo para finalização com erro
          Name strNomeArquivoREA As strNomeArquivoErr
     End If
     GoTo FimLeitura
     
ErroLeitura:
     Beep
     Screen.MousePointer = vbDefault
     
     'Tratamento para diretório não localizado
     If Err.Number = 76 Then
          MsgBox "Não localizado o diretório com o arquivo de movimento Data Boa." & vbCrLf & vbCrLf & _
                    "Favor verificar diretório em parâmetros.", vbCritical, sTituloTela
          GoTo FimLeitura
     End If
     If Err.Number = 53 Then
          MsgBox "Não localizado o arquivo de movimento de data boa." & vbCrLf & vbCrLf & _
                    "Favor verificar e tentar novamente.", vbCritical, sTituloTela
          GoTo FimLeitura
     End If
     If Err.Number = 55 Then
          MsgBox "Arquivo com movimento de data boa em utilização por outro usuário. Favor verificar !", vbCritical, sTituloTela
          GoTo FimLeitura
     End If

    MsgBox "Erro na Leitura do Arquivo de Cheques da Data Boa.", vbOKOnly + vbCritical, sTituloTela

    GoTo FimLeitura

End Sub


Public Sub RecConfBaixas()
 
    Dim rstCheques            As New ADODB.Recordset
    Dim InsCheques            As New Custodia.Inserir
    Dim SelCheques            As New Custodia.Selecionar
    Dim DatFile               As Integer
    Dim lRetorno              As Long
    Dim Reg                   As String * 104
    Dim bArquivoReading       As Boolean
    Dim strNomeArquivo      As String
    Dim strNomeArquivoREA   As String
    Dim strNomeArquivoOK    As String
    Dim strNomeArquivoErr   As String
    Dim OffSet                As Long
    Dim nBaixas               As Long
    Dim nJaBaixado            As Long
    Dim ChBaixado             As ChrBai_Reg
    Dim sstr                  As String
    Dim sDigito               As String
    Dim sNumBordero           As String
    Dim sWhere                As String
    Dim Progress              As New clsProgressBar
    Dim lngRegs               As Long
    Dim sTituloTela           As String
    Dim sPathName             As String
    Dim sArquivoRecep         As String
    Dim aArquivoRecep()
    Dim iCount                As Integer, aIndex As Integer
    
     On Error GoTo ErroLeitura
        
     nBaixas = 0
     sTituloTela = "Recepção da Baixa de Cheques"
    
     If MsgBox("Inicializa processo de recepção", vbQuestion + vbYesNo, sTituloTela) = vbNo Then
          Exit Sub
     End If
    
     If Not LerArquivosDiretorio(g_Parametros.DiretorioRecepcao & "\", "CHRBAI", sArquivoRecep) Then
          Exit Sub
     End If
     If Len(sArquivoRecep) = 0 Then
          Beep
          MsgBox "Não existe arquivo(s) para recepção referente à Baixa de Cheques.", vbInformation, sTituloTela
          Exit Sub
     End If
     
     ReDim aArquivoRecep(0)
     Screen.MousePointer = vbHourglass
     
     Do While True
          iCount = InStr(sArquivoRecep, ",")
          If iCount = 0 Then Exit Do
          aIndex = UBound(aArquivoRecep) + 1
          ReDim Preserve aArquivoRecep(aIndex)
          aArquivoRecep(aIndex) = Mid(sArquivoRecep, 1, (iCount - 1))
          sArquivoRecep = Mid(sArquivoRecep, (iCount + 1))
     Loop

     Call LabelRecepcao(LabelAcao.Apresenta)
    
     DatFile = FreeFile
     
     'Recepciona quantos arquivos existirem
     For iCount = 1 To aIndex
     
          Call LabelRecepcao(LabelAcao.Atualiza, iCount, aIndex, aArquivoRecep(iCount))
          
          strNomeArquivo = Trim(g_Parametros.DiretorioRecepcao) & "\" & aArquivoRecep(iCount)
          'Muda extensão do arquivo para (REA) qdo estiver sendo acessado para leitura
          strNomeArquivoREA = strNomeArquivo & "_READ"
          'Muda extensão do arquivo para (OK) qdo finalizado o processo
          strNomeArquivoOK = strNomeArquivo & "_OK"
          'Muda extensão do arquivo para (Erro) qdo ocorrer erro
          strNomeArquivoErr = strNomeArquivo & "_ERRO"
          
          bArquivoReading = False
          
          'Verifica se arq. confirmação existe no diretório e está sendo acessado por outro usuário
          If Dir(strNomeArquivo, vbDirectory) = "" Then
               
               'Verifica se arq. confirmação está sendo lido por outro usuário (.REA)
               If Dir(strNomeArquivoREA, vbDirectory) = "" Then

                    'Abre arquivo para verificar erro de inexistencia do diretório/Arquivo
                    Open strNomeArquivo For Input As #1
                    Close #1
                    Exit Sub
               Else
                    'Força abertura do arquivo para executar tratamento do erro
                    'Caso consiga abrir o arquivo, é sinal de que houve queda de execução
                    'e o arquivo ficou em modo Reading...
                    Open strNomeArquivoREA For Input Access Read As #1
                    Close #1
                    bArquivoReading = True
               End If
          End If
          
          'Renomeia o arquivo para extensão (.REA)
          If Not bArquivoReading Then
               Name strNomeArquivo As strNomeArquivoREA
          End If
          
          Open strNomeArquivoREA For Binary Access Read Lock Read Write As #DatFile
        
          OffSet = 1
          
          Get #DatFile, OffSet, Reg
    
          If Not EOF(DatFile) Then
                'Inicia progress bar
                Progress.ValorMinimo = 1
                Progress.ValorMaximo = Fix(FileLen(strNomeArquivoREA) / Len(Reg))
                Progress.DescricaoProcesso = "Recepcionando Baixa de Cheques ..."
                Progress.InicializaProgressBar
                Progress.AtualizaBarra
                lngRegs = 0
          End If
    
          While Not EOF(DatFile)
              'Acumulador de registros lidos
              lngRegs = lngRegs + 1
          
              ' Se arquivo foi lido ok
              If Len(Reg) < 104 Then
                   MsgBox "Erro de Leitura", vbOKOnly + vbCritical, sTituloTela
                   GoTo FimLeituraComErro
              End If
                  
              ' Ver se CGC de terceira é válido
              If CStr(Mid(Reg, 7, 14)) <> Format(g_Parametros.CNPJ_Terceira, String(14, "0")) Then
                   MsgBox "CNPJ da Terceira Inválido", vbOKOnly + vbCritical, sTituloTela
                   GoTo FimLeituraComErro
              End If
                      
              ' Atribuir registos do arquivo
              ChBaixado.Num_Bordero = Mid(Reg, 21, 18)
              ChBaixado.CodigoCarteira = Mid(Reg, 39, 2)
              ChBaixado.DataDeposito = Mid(Reg, 41, 8)
              ChBaixado.BancoEmitente = Mid(Reg, 49, 4)
              ChBaixado.AgenciaEmitente = Mid(Reg, 53, 4)
              ChBaixado.NrChequeEmitente = Mid(Reg, 57, 6)
              ChBaixado.NossoNumero = Mid(Reg, 63, 11)
              ChBaixado.CodigoCompensacao = Mid(Reg, 74, 3)
              ChBaixado.CcEmitente = Mid(Reg, 77, 11)
              ChBaixado.ValorCheque = CDbl(Mid(Reg, 88, 15)) / 100
              
             
              'Calcula dígito verificador do borderô
              ' Para Calcular o digito considerar: Mid(NumeroBordero, 13, 6)
              
              sDigito = RetornaDigitoModulo11Simplificado(Mid(ChBaixado.Num_Bordero, 13, 6))
              sNumBordero = Mid(Reg, 21, 18) & Trim(sDigito)
              
              
              Set rstCheques = g_cMainConnection.Execute(SelCheques.GetChequesBaixados(sNumBordero, CLng(ChBaixado.DataDeposito), _
                                                  ChBaixado.BancoEmitente, ChBaixado.AgenciaEmitente, ChBaixado.CcEmitente, _
                                                  ChBaixado.NrChequeEmitente))
              
              
              
              If rstCheques.EOF Then
              
                  ' Gravar Registro de Cheques Baixados
                  Call g_cMainConnection.Execute(InsCheques.InsereChequesBaixados(sNumBordero, CInt(ChBaixado.BancoEmitente), _
                                                                              CInt(ChBaixado.AgenciaEmitente), CLng(ChBaixado.NrChequeEmitente), _
                                                                              CDbl(ChBaixado.CcEmitente), ChBaixado.NossoNumero, _
                                                                              CInt(ChBaixado.CodigoCompensacao), CLng(ChBaixado.DataDeposito), _
                                                                              ChBaixado.ValorCheque, Geral.DataProcessamento, CByte(ChBaixado.CodigoCarteira)), _
                                                                              lRetorno, adCmdText)
              
              
                  If lRetorno > 0 Then
                      nBaixas = nBaixas + lRetorno
                  End If
              
              Else
                  nJaBaixado = nJaBaixado + 1
              
              End If
              
              OffSet = OffSet + Len(Reg)
              
              Get #DatFile, OffSet, Reg
              
              'Atualiza Progress Bar
              Progress.AtualValue = lngRegs
              Progress.AtualizaBarra
              
          Wend
    
          Close #DatFile
          
          'Renomeia o arquivo para extensão (.OK)
          Name strNomeArquivoREA As strNomeArquivoOK
          
     Next
    
     'Encerra progress bar
     Set Progress = Nothing
    
     'Fecha tela apresentação de arquivos
     Call LabelRecepcao(LabelAcao.Finaliza)
    
     Screen.MousePointer = vbDefault
     MsgBox "Foram Processados " & CStr(nBaixas + nJaBaixado) & " Cheques, Foram Baixados " & CStr(nBaixas), vbOKOnly + vbExclamation, sTituloTela
     Exit Sub
    
FimLeitura:
     'Fecha tela apresentação de arquivos
     Call LabelRecepcao(LabelAcao.Finaliza)

     Close #DatFile
    
     'Encerra progress bar
     Set Progress = Nothing
     
     Screen.MousePointer = vbDefault
     Exit Sub

FimLeituraComErro:

     Close #DatFile

     If UBound(aArquivoRecep) >= 1 Then
          'Renomeia o arquivo para finalização com erro
          Name strNomeArquivoREA As strNomeArquivoErr
     End If
     GoTo FimLeitura
    
ErroLeitura:
     Beep
     Screen.MousePointer = vbDefault
     
     'Tratamento para diretório não localizado
     If Err.Number = 76 Then
          MsgBox "Não localizado o diretório com o arquivo baixa de cheque." & vbCrLf & vbCrLf & _
                    "Favor verificar diretório em parâmetros.", vbCritical, sTituloTela
          GoTo FimLeitura
     End If
     If Err.Number = 53 Then
          MsgBox "Não localizado o arquivo de baixa de cheque." & vbCrLf & vbCrLf & _
                    "Favor verificar e tentar novamente.", vbCritical, sTituloTela
          GoTo FimLeitura
     End If
     If Err.Number = 55 Then
          MsgBox "Arquivo baixa de cheque em utilização por outro usuário. Favor verificar !", vbCritical, sTituloTela
          GoTo FimLeitura
     End If

     MsgBox "Erro na Leitura do Arquivo da Baixas", vbOKOnly + vbCritical, sTituloTela
    
     GoTo FimLeitura
     
End Sub


Public Sub RecInstrucoes()

Dim myrecord             As RecordInstrucoes
Dim lngTotalRegistros    As Long
Dim strNomeArquivo       As String, strNomeArquivoREA As String, strNomeArquivoOK As String, strNomeArquivoErr As String
Dim lngRegs              As Long, lngRetorno As Long
Dim recepcao()           As tpRecepcao
Dim Progress             As New clsProgressBar
Dim bArquivoReading      As Boolean
Dim strDataChave         As String
Dim strHoraChave         As String
Dim rsRetorno            As New ADODB.Recordset
Dim strDataProcessamento As String
Dim lngCount             As Long
Dim strDigBordero        As String      'Dígito verificador do borderô
Dim sTituloTela          As String
Dim sArquivoRecep        As String
Dim aArquivoRecep()
Dim iCount               As Integer, aIndex As Integer

On Error GoTo Erro_RecInstrucoes

     sTituloTela = "Recepção da Tabela de instrução VC"

     If MsgBox("Inicializa processo de recepção da tabela de instruções do VC", vbQuestion + vbYesNo, sTituloTela) = vbNo Then
          Set rsRetorno = Nothing
          Exit Sub
     End If
     
     If Not LerArquivosDiretorio(g_Parametros.DiretorioRecepcao & "\", "CHRDTV", sArquivoRecep) Then
          Set rsRetorno = Nothing
          Exit Sub
     End If
     If Len(sArquivoRecep) = 0 Then
          Beep
          MsgBox "Não existe arquivo(s) para recepção referente à Tabela de Instrução", vbInformation, sTituloTela
          Set rsRetorno = Nothing
          Exit Sub
     End If
     
     ReDim recepcao(0)
     ReDim aArquivoRecep(0)
     Screen.MousePointer = vbHourglass
     
     Do While True
          iCount = InStr(sArquivoRecep, ",")
          If iCount = 0 Then Exit Do
          aIndex = UBound(aArquivoRecep) + 1
          ReDim Preserve aArquivoRecep(aIndex)
          aArquivoRecep(aIndex) = Mid(sArquivoRecep, 1, (iCount - 1))
          sArquivoRecep = Mid(sArquivoRecep, (iCount + 1))
     Loop

     Call LabelRecepcao(LabelAcao.Apresenta)

     For iCount = 1 To aIndex
     
          Call LabelRecepcao(LabelAcao.Atualiza, iCount, aIndex, aArquivoRecep(iCount))
          
          strNomeArquivo = Trim(g_Parametros.DiretorioRecepcao) & "\" & aArquivoRecep(iCount)
          'Muda extensão do arquivo para (REA) qdo estiver sendo acessado para leitura
          strNomeArquivoREA = strNomeArquivo & "_READ"
          'Muda extensão do arquivo para (OK) qdo finalizado o processo
          strNomeArquivoOK = strNomeArquivo & "_OK"
          'Muda extensão do arquivo para (Erro) qdo ocorrer erro
          strNomeArquivoErr = strNomeArquivo & "_ERRO"
     
          bArquivoReading = False
     
          'Verifica se arquivo existe no diretório e está sendo acessado por outro usuário
          If Dir(strNomeArquivo, vbDirectory) = "" Then
     
               'Verifica se arquivo está sendo lido por outro usuário (.REA)
               If Dir(strNomeArquivoREA, vbDirectory) = "" Then
                    
                    'Abre arquivo para verificar erro de inexistencia do diretório/Arquivo
                    Open strNomeArquivoREA For Input As #1
                    Close #1
                    Exit Sub
               Else
                    'Força abertura do arquivo para executar tratamento do erro
                    'Caso consiga abrir o arquivo, é sinal de que houve queda de execução
                    'e o arquivo ficou em modo Reading...
                    Open strNomeArquivoREA For Input Access Read As #1
                    Close #1
                    bArquivoReading = True
               End If
          End If
     
          'Renomeia o arquivo para extensão (.REA)
          If Not bArquivoReading Then
               Name strNomeArquivo As strNomeArquivoREA
          End If
          
          'Abre o arquivo de rejeições com extensão (.REA)
          Open strNomeArquivoREA For Random As #1 Len = Len(myrecord)
          
          'Obtem o total de bytes do arquivo de rejeitados
          lngTotalRegistros = FileLen(strNomeArquivoREA) / Len(myrecord)
          
         'Inicializa Progress Bar
          Progress.ValorMinimo = 1
          Progress.ValorMaximo = lngTotalRegistros
          Progress.DescricaoProcesso = "Recepcionando tabela de instruções ..."
          Progress.InicializaProgressBar
          Progress.AtualizaBarra
          
          'Ler cada registro e atualizar tabelas
          For lngRegs = 1 To lngTotalRegistros
               'Ler linha à linha
               Get #1, lngRegs, myrecord
               
               'Calcula dígito verificador do borderô
               
               strDigBordero = RetornaDigitoModulo11Simplificado(Mid(myrecord.Num_Bordero, 13, 6))
               
               'Verifica se existe código da carteira
               If Val(myrecord.Cod_Carteira) = 0 Then
                    Call GeraInconsInstrucao("Não localizado código de carteira", myrecord, recepcao, strDigBordero)
               Else
                    'Localiza código da carteira na tabela de Carteira
                    Set rsRetorno = g_cMainConnection.Execute(Proc_Selecionar.GetCarteira(myrecord.Cod_Carteira), _
                                                            lngRetorno, adCmdText)
               
                    If rsRetorno.EOF Then
                         Call GeraInconsInstrucao("Não localizado código de carteira", myrecord, recepcao, strDigBordero)
                    Else
                    
                         'Insere registro na tabela de Alteração de Data Depósito
                         Call g_cMainConnection.Execute(Proc_Inserir.InsereAlteracaoData(Geral.DataProcessamento, _
                                                       myrecord.Num_Bordero & strDigBordero, _
                                                       myrecord.Cod_Carteira, _
                                                       myrecord.Dat_DepAnterior, _
                                                       myrecord.Num_Banco, _
                                                       myrecord.Num_Agencia, _
                                                       myrecord.Num_ContaCorrente, _
                                                       myrecord.Num_Cheque, _
                                                       myrecord.Dat_DepNova, _
                                                       myrecord.Cod_NossoNumero, _
                                                       myrecord.Cod_compensacao, _
                                                       myrecord.ValorCheque / 100 _
                                                       ), lngRetorno, adCmdText)
                         If lngRetorno <= 0 Then
                              Call GeraInconsInstrucao("Não foi possível atualizar a alteração de data", myrecord, recepcao, strDigBordero)
                         End If
                    End If
               End If
               
               'Atualiza Progress Bar
               Progress.AtualValue = lngRegs
               Progress.AtualizaBarra
               
          Next
     
          'Fecha arquivo de rejeições
          Close #1

          'Renomeia o arquivo para extensão (.OK)
          Name strNomeArquivoREA As strNomeArquivoOK
     
     Next

     'Verifica se houve ocorrencia na recepção
     If UBound(recepcao) > 0 Then
          
          Set rsRetorno = Nothing
          'Obtem data e hora do servidor (Chave para a tabela INCONSISTENCIA)
          Set rsRetorno = g_cMainConnection.Execute(Proc_Selecionar.GetDataHoraServidor(), lngRetorno, adCmdText)
          
          If lngRetorno <> 0 Then
               Beep
               MsgBox "Problema na geração do relatório de inconsistências!", vbCritical, sTituloTela
               GoTo Sair
          End If
          
          strDataChave = Format(rsRetorno!Data, "yyyymmdd")
          strHoraChave = rsRetorno!Hora
          
          For lngCount = 1 To UBound(recepcao)
               'Insere registro com dados de inconsistência
               Call g_cMainConnection.Execute(Proc_Inserir.InsereInconsistencia( _
                                                  strDataChave, _
                                                  strHoraChave, _
                                                  recepcao(lngCount).Num_Bordero & _
                                                  recepcao(lngCount).Banco & _
                                                  recepcao(lngCount).Agencia & _
                                                  recepcao(lngCount).Conta & _
                                                  recepcao(lngCount).Cheque & _
                                                  recepcao(lngCount).NossoNumero & _
                                                  recepcao(lngCount).Inconsistencia))

          Next
     
'          'Imprime relação de ocorrências do processo de recepção
'          Call ImprimirInconsistencia("RelRecepcaoInstrucoes", "Ocorrências na recepção da tabela de instrução do VC", Geral.DataProcessamento, strDataChave, strHoraChave)
          
          'Fecha tela apresentação de arquivos
          Call LabelRecepcao(LabelAcao.Finaliza)
          MsgBox Space(32) & "A T E N Ç Ã O" & vbCrLf & vbCrLf & _
                 "Finalizada a recepção com ocorrência(s).  Verifique !", vbCritical, sTituloTela
          
          
          'Remove registro com dados de inconsistência
          Call g_cMainConnection.Execute(Proc_Excluir.RemoveInconsistencia(strDataChave, _
                                                                           strHoraChave), _
                                                                           lngRetorno, adCmdText)
     Else
          'Fecha tela apresentação de arquivos
          Call LabelRecepcao(LabelAcao.Finaliza)
          
          MsgBox "Finalizada a recepção da tabela de instruções", vbInformation, sTituloTela
     End If
     
     
Sair:
     
     'Fecha tela apresentação de arquivos
     Call LabelRecepcao(LabelAcao.Finaliza)

     Set Progress = Nothing
     
     Set rsRetorno = Nothing
     
     Screen.MousePointer = vbDefault
     Exit Sub
     
SairSemFinalizar:
     'Fecha arquivo de rejeições
     Close #1
     
     'Renomeia o arquivo para extensão inicial ao processo
     Name strNomeArquivoREA As strNomeArquivo
     GoTo Sair
     
Erro_RecInstrucoes:
     
     Beep
     Screen.MousePointer = vbDefault
     
     'Tratamento para diretório não localizado
     If Err.Number = 76 Then
          MsgBox "Não localizado o diretório com o arquivo de rejeição, " & _
                    "verifique diretório no módulo parâmetros.", vbCritical, sTituloTela
          GoTo Sair
     End If
     If Err.Number = 53 Then
          MsgBox "Não localizado o arquivo de rejeição, " & _
                 "verifique e tentar novamente.", vbCritical, sTituloTela
          GoTo Sair
     End If
     If Err.Number = 55 Then
          MsgBox "Arquivo de rejeição em utilização por outro usuário. Favor verificar!", vbCritical, sTituloTela
          GoTo Sair
     End If
     
    'Fecha arquivo de rejeições
     Close #1

     If UBound(aArquivoRecep) >= 1 Then
         'Renomeia o arquivo para finalização com erro
          Name strNomeArquivoREA As strNomeArquivoErr
     End If
     
     MsgBox Err.Description, vbCritical, sTituloTela
     GoTo Sair

End Sub
Private Sub GeraInconsInstrucao(strDescrMsg As String, myrecord As RecordInstrucoes, recepcao() As recepcao.tpRecepcao, ByVal strDigBordero As String)

Dim lngCount As Long
     
     ReDim Preserve recepcao(UBound(recepcao) + 1)
     lngCount = UBound(recepcao)

     'Acumula em type os borderôs com ocorrência
     recepcao(lngCount).Num_Bordero = myrecord.Num_Bordero & strDigBordero
     recepcao(lngCount).Banco = myrecord.Num_Banco
     recepcao(lngCount).Agencia = myrecord.Num_Agencia
     recepcao(lngCount).Conta = myrecord.Num_ContaCorrente
     recepcao(lngCount).Cheque = myrecord.Num_Cheque
     recepcao(lngCount).NossoNumero = myrecord.Cod_NossoNumero
     recepcao(lngCount).Inconsistencia = strDescrMsg

End Sub
Sub RecRejeitados()
     On Error GoTo Erro:
     
     Dim lngTotalRegistros    As Long
     Dim lngAtualRegistro     As Long
     Dim lngRegs              As Long
     Dim lngRetorno           As Long
     Dim recepcao()           As tpRecepcao
     Dim bArquivoReading      As Boolean
     Dim Progress             As New clsProgressBar
     Dim strDataChave         As String
     Dim strHoraChave         As String
     Dim strCMC7              As String
     Dim lngCount             As Long
     Dim strDigBordero        As String
     Dim sstr                 As String
     Dim Bordero              As String
     Dim TabRejeicao          As tpRejeicao
     Dim strNomeArquivo       As String
     Dim strNomeArquivoREA    As String
     Dim strNomeArquivoOK     As String
     Dim strNomeArquivoErr    As String
     Dim rs                   As New ADODB.Recordset
     Dim sTituloTela          As String
     Dim sArquivoRecep        As String
     Dim aArquivoRecep()
     Dim iCount               As Integer, aIndex As Integer
     
     sTituloTela = "Recepção de Cheques Rejeitados"
     
     If MsgBox("Inicializa processo da recepção de Cheques Rejeitados", vbQuestion + vbYesNo, sTituloTela) = vbNo Then
          Set rs = Nothing
          Exit Sub
     End If
     
     If Not LerArquivosDiretorio(g_Parametros.DiretorioRecepcao & "\", "REJEIC", sArquivoRecep) Then
          Set rs = Nothing
          Exit Sub
     End If
     
     If Len(sArquivoRecep) = 0 Then
          Beep
          MsgBox "Não existe arquivo(s) para recepção referente à Cheques Rejeitados", vbInformation, sTituloTela
          Set rs = Nothing
          Exit Sub
     End If
     
     ReDim recepcao(0)
     ReDim aArquivoRecep(0)
     Screen.MousePointer = vbHourglass
     
     Do While True
          iCount = InStr(sArquivoRecep, ",")
          If iCount = 0 Then Exit Do
          aIndex = UBound(aArquivoRecep) + 1
          ReDim Preserve aArquivoRecep(aIndex)
          aArquivoRecep(aIndex) = Mid(sArquivoRecep, 1, (iCount - 1))
          sArquivoRecep = Mid(sArquivoRecep, (iCount + 1))
     Loop

     Call LabelRecepcao(LabelAcao.Apresenta)

     For iCount = 1 To aIndex
     
          Call LabelRecepcao(LabelAcao.Atualiza, iCount, aIndex, aArquivoRecep(iCount))
          strNomeArquivo = Trim(g_Parametros.DiretorioRecepcao) & "\" & aArquivoRecep(iCount)
          
         'Muda extensão do arquivo para (REA) qdo estiver sendo acessado para leitura
          strNomeArquivoREA = strNomeArquivo & "_READ"
          
         'Muda extensão do arquivo para (OK) qdo finalizado o processo
          strNomeArquivoOK = strNomeArquivo & "_OK"
          
         'Muda extensão do arquivo para (Erro) qdo ocorrer erro
          strNomeArquivoErr = strNomeArquivo & "_ERRO"
     
          bArquivoReading = False

         'arq. rejeição existe no diretório e está sendo acessado por outro usuário ?
          If Dir(strNomeArquivo, vbDirectory) = "" Then
              'arq. rejeição está sendo lido por outro usuário ? (.REA)
               If Dir(strNomeArquivoREA, vbDirectory) = "" Then
                   'Abre arquivo para verificar erro de inexistencia do diretório/Arquivo
                    Open strNomeArquivo For Input As #1
                    Close #1
                    Exit Sub
               Else
                    Open strNomeArquivoREA For Input Access Read As #1
                    Close #1
                    bArquivoReading = True
               End If
          End If
          
         'Renomeia o arquivo para extensão (.REA)
          If Not bArquivoReading Then
               Name strNomeArquivo As strNomeArquivoREA
          End If
          
         'Abre o arquivo de rejeições com extensão (.REA)
          Open strNomeArquivoREA For Random As #1 Len = Len(myrecord)
          
         'Obtem o total de bytes do arquivo de rejeitados
          lngTotalRegistros = FileLen(strNomeArquivoREA) / Len(myrecord)
          
         'Inicializa Progress Bar
          Progress.ValorMinimo = 0
          Progress.ValorMaximo = lngTotalRegistros
          Progress.DescricaoProcesso = "Recepcionando rejeição de cheque ..."
          Progress.InicializaProgressBar
          Progress.AtualizaBarra
          
         'Ler cada registro e atualizar tabelas
          For lngRegs = 1 To lngTotalRegistros
             'Ler linha à linha
              Get #1, lngRegs, myrecord
              
             'Verifica se Arquivo é válido
              If Not IsNumeric(myrecord.Num_Bordero) Or myrecord.Rot_Mensag <> "REJEIC" Then
                 Beep
                 Err.Raise 910, App.Title, "Falha no módulo de Recepção do Arquivo de Cheques Rejeitados, Arquivo com formato desconhecido"
              End If
              
             'Calcula Digito do bordero
              Bordero = myrecord.Num_Bordero & RetornaDigitoModulo11Simplificado(Mid(myrecord.Num_Bordero, 13, 6))
              
             'Procura Bordero na Base
              sstr = "Select DataProcessamento, Idbordero, Status From Bordero Where Num_bordero  = '" & Format(Bordero, String(19, "0")) & "'" & " And Status IN('T','C')"
              Set rs = g_cMainConnection.Execute(sstr, lngRetorno, adCmdText)
                       
             'Se Não Encontrou Bordero
              If rs.EOF Then
                 Call GeraInconsRejeitados("Número de Borderô não Encontrado", recepcao, Bordero)
              ElseIf rs.RecordCount > 1 Then
                'Se Encontrado mais de um Bordero
                 Call GeraInconsRejeitados("Encontrado mais de Um Borderô com mesmo número", recepcao, myrecord.Num_Bordero & strDigBordero)
              ElseIf rs("Status").Value = "E" Then
                Call GeraInconsRejeitados("Borderô confirmado", recepcao, Bordero)
              Else
                'Bordero Encontrado, joga dados no type p/ posterior inserçao na tabela de rejeicao
                 TabRejeicao.DataProcessamento = rs("DataProcessamento").Value
                 TabRejeicao.IdBordero = rs("IdBordero").Value
                 TabRejeicao.DataDeposito = 0
                 TabRejeicao.IdCheque = 0
                 TabRejeicao.Rotulo = myrecord.Rot_Original
                 TabRejeicao.CodErro = 0
              
                 If rs("Status").Value <> "X" Then
                     sstr = "Update Bordero Set Status = 'X' where IdBordero = " & TabRejeicao.IdBordero  '  & Format(Bordero, String(19, "0")) & "'"
                     Call g_cMainConnection.Execute(sstr, lngRetorno, adCmdText)
     
                     If lngRetorno <> 1 Then
                         Beep
                         Err.Raise 910, App.Title, "Falha no módulo de Recepção do Arquivo de Cheques Rejeitados, Atualização de Status do Borderô."
                     End If
                 End If
                 
                '*********************************************
                'Inserir na Tabela Rejeição remessa os erros *
                '*********************************************
                 If myrecord.Rot_Original = "CHINBO" Then
                 
                    Call CodigoErro(TabRejeicao, Bordero)
                 
                'Verifica se o rotulo é de Cheque e trata
                 ElseIf myrecord.Rot_Original = "CHINCH" Then
                     
                     'Altera status do cheque para (X)
                      strCMC7 = Replace(myrecord.Cod_CMC7, " ", "", 1, Len(myrecord.Cod_CMC7), vbTextCompare)
                      
                      sstr = "Select DataDeposito, IdCheque From Cheque Where " & _
                             " IdBordero = " & TabRejeicao.IdBordero & _
                             " And CMC7 = '" & strCMC7 & _
                             " ' And Status = 'T'" & _
                             " And DataProcessamento = " & TabRejeicao.DataProcessamento
                              
                      Set rs = g_cMainConnection.Execute(sstr, lngRetorno, adCmdText)
                      
                      If rs.RecordCount <> 0 Then
                        
                        'Joga dados no type p/ posterior inserçao na tabela de rejeicao
                         TabRejeicao.DataDeposito = rs("DataDeposito").Value
                         TabRejeicao.IdCheque = rs("IdCheque").Value
                         
                         sstr = "Update Cheque Set Status = 'X' " & _
                                " Where Dataprocessamento = " & TabRejeicao.DataProcessamento & _
                                " And IdBordero = " & TabRejeicao.IdBordero & _
                                " And DataDeposito = " & TabRejeicao.DataDeposito & _
                                " And Status = 'T'" & _
                                " And IdCheque = " & TabRejeicao.IdCheque
                                
                         Call g_cMainConnection.Execute(sstr, lngRetorno, adCmdText)
                             
                         If lngRetorno <= 0 Then
                           'se Houve falha na atualizaçao de status do cheque
                            Call GeraInconsRejeitados("Não foi possível atualizar o Cheque", recepcao, Bordero)
                         Else
                            'Se encontrado Ok
                            Call CodigoErro(TabRejeicao, Bordero)
                         End If
                      Else
                        'Se não encontrado cheque na Base
                         Call GeraInconsRejeitados("Cheque não Localizado na Base", recepcao, Bordero)
                      End If
                'Verifica se o rotulo é de Data e trata
                ElseIf myrecord.Rot_Original = "CHINDT" Then
                 
                     sstr = "Select DataDeposito From DataDeposito Where " & _
                             " IdBordero = " & TabRejeicao.IdBordero & _
                             " And Status = '1' " & _
                             " And DataProcessamento = " & TabRejeicao.DataProcessamento
                              
                      Set rs = g_cMainConnection.Execute(sstr, lngRetorno, adCmdText)
                      
                      If rs.RecordCount <> 0 Then
                         TabRejeicao.DataDeposito = rs("DataDeposito").Value
                                          
                         sstr = "Update DataDeposito Set Status = 'X' " & _
                                " Where Dataprocessamento = " & TabRejeicao.DataProcessamento & _
                                " And IdBordero = " & TabRejeicao.IdBordero
                                
                         Call g_cMainConnection.Execute(sstr, lngRetorno, adCmdText)
                             
                         If lngRetorno <= 0 Then
                           'Não localizada data na base
                            Call GeraInconsRejeitados("Não foi possível atualizar a Data Depósito", recepcao, Bordero)
                         Else
                           'Se encontrado Ok
                            Call CodigoErro(TabRejeicao, Bordero)
                         End If
                      Else
                         Call GeraInconsRejeitados("Data Depósito não Localizada na Base", recepcao, Bordero)
                      End If
                      
                 End If
                 
              End If
                   
             'Atualiza Progress Bar
             Progress.AtualValue = lngRegs
             Progress.AtualizaBarra
                  
          Next
          
         'Fecha arquivo de rejeições
          Close #1
     
          'Renomeia o arquivo para extensão (.OK)
          Name strNomeArquivoREA As strNomeArquivoOK
     Next
     
    'Verifica se houve ocorrencia na recepção
     If UBound(recepcao) > 0 Then
          
          Set rs = Nothing
          
         'Obtem data e hora do servidor (Chave para a tabela INCONSISTENCIA)
          Set rs = g_cMainConnection.Execute(Proc_Selecionar.GetDataHoraServidor(), lngRetorno, adCmdText)
          
          If lngRetorno <> 0 Then
               Beep
               Err.Raise 910, App.Title, "Falha no módulo de Recepção do Arquivo de Cheques Rejeitados,(Obtenção de Hora do Servidor)"
          End If
          
          strDataChave = Format(rs!Data, "yyyymmdd")
          strHoraChave = rs!Hora
          
          For lngCount = 1 To UBound(recepcao)
              'Insere registro com dados de inconsistência
               Call g_cMainConnection.Execute(Proc_Inserir.InsereInconsistencia( _
                                                  strDataChave, _
                                                  strHoraChave, _
                                                  recepcao(lngCount).Num_Bordero & _
                                                  recepcao(lngCount).DataDeposito & _
                                                  recepcao(lngCount).Banco & _
                                                  recepcao(lngCount).Agencia & _
                                                  recepcao(lngCount).Conta & _
                                                  recepcao(lngCount).Inconsistencia))
          Next
     
         'Imprime relação de ocorrências do processo de recepção
          Call ImprimirInconsistencia("RelInconsistenciaRecepcao", "Ocorrências da recepção de rejeições", Geral.DataProcessamento, strDataChave, strHoraChave)
          
         'Remove registro com dados de inconsistência
          Call g_cMainConnection.Execute(Proc_Excluir.RemoveInconsistencia(strDataChave, strHoraChave), lngRetorno, adCmdText)
     
     End If
     
Fim:
     'Fecha tela apresentação de arquivos
     Call LabelRecepcao(LabelAcao.Finaliza)

     Set Progress = Nothing
     
     Set rs = Nothing
       
     Screen.MousePointer = vbDefault
     Exit Sub
     
Erro:
     Beep
     Screen.MousePointer = vbDefault
     
     'Tratamento para diretório não localizado
     If Err.Number = 76 Then
        Call TratamentoErro("Não localizado o diretório com o arquivo de rejeição, " & _
                            "verifique o diretório no módulo parâmetros.", Err, False, False)
        Resume Fim
     End If
     If Err.Number = 53 Then
        Call TratamentoErro("Não localizado o arquivo de rejeição, " & _
                            "verifique e tente novamente.", Err, False, False)
        Resume Fim
     End If
     If Err.Number = 55 Then
        Call TratamentoErro("Arquivo de rejeição em utilização por outro usuário. Favor verificar!", Err, False, False)
        Resume Fim
     End If
     
    'Fecha arquivo de rejeições
     Close #1
     
     If UBound(aArquivoRecep) > 1 Then
         'Renomeia o arquivo para finalização com erro
          Name strNomeArquivoREA As strNomeArquivoErr
     End If
     
     Call TratamentoErro("Falha no módulo de Recepção de Cheques", Err, False, False)
     Resume Fim

End Sub
Function CodigoErro(Registro As tpRejeicao, pBordero As String) As Boolean
Dim iCodErro, Retorno     As Integer
Dim intCodErro            As String
Dim rs                    As New ADODB.Recordset

For iCodErro = 1 To CInt(myrecord.Qtd_Erros)
    'Obtem separadamente o cód. de erro existente no registro do arq. Rejeições
     intCodErro = Mid(myrecord.Cod_Erros, (iCodErro * 3) - 2, 3)
          
     Set rs = g_cMainConnection.Execute(Proc_Selecionar.GetExisteRejeicaoRemessa(Registro.DataProcessamento, Registro.IdBordero, Registro.DataDeposito, Registro.IdCheque, intCodErro), Retorno, adCmdText)

     If Retorno <> 0 Then
        Beep
        MsgBox "Processo cancelado." & vbCrLf & vbCrLf & "Não foi possível ler informação de Rejeição do borderô (" & pBordero & ")", vbCritical, App.Title
        Exit Function
     End If
    
     If rs(0).Value = 0 Then

         Call g_cMainConnection.Execute(Proc_Inserir.InsereRejeicaoRemessa(Registro.DataProcessamento, Registro.IdBordero, Registro.DataDeposito, Registro.IdCheque, intCodErro, Registro.Rotulo), Retorno, adCmdText)

         If Retorno <= 0 Then
            Beep
            MsgBox "Processo cancelado." & vbCrLf & vbCrLf & "Não foi possível inserir informações de rejeição no arquivo.", vbCritical, App.Title
            Exit Function
         End If
     End If
Next

CodigoErro = True

End Function

Private Sub LabelRecepcao(ByVal iAcao As Integer, _
                         Optional iArq_Atual As Integer, _
                         Optional iArq_Total As Integer, _
                         Optional ByVal sArquivo As String)

     With Principal
          
          If iAcao = LabelAcao.Apresenta Then      'Apresenta Label de arquivos em recepção
               .lblRecepcao.Caption = ""
               .lblRecepcao.Top = (Screen.Height - .lblRecepcao.Height) / 2
               .lblRecepcao.Left = (Screen.Width - .lblRecepcao.Width) / 2
               .lblRecepcao.Visible = True
     
               .lblRecepcaoArquivo.Caption = ""
               .lblRecepcaoArquivo.Top = (Screen.Height - .lblRecepcaoArquivo.Height + 350) / 2
               .lblRecepcaoArquivo.Left = (Screen.Width - .lblRecepcaoArquivo.Width) / 2
               .lblRecepcaoArquivo.Visible = True
          
          ElseIf iAcao = LabelAcao.Finaliza Then  'Fecha Label com apresentação de arquivos em recepção
               .lblRecepcao.Visible = False
               .lblRecepcaoArquivo.Visible = False
          
          ElseIf iAcao = LabelAcao.Atualiza Then  'Preenche Label de arquivos em recepção com as informações
               .lblRecepcao.Caption = "Recepcionando arquivo " & "( " & CStr(iArq_Atual) & " de " & CStr(iArq_Total) & " )"
               .lblRecepcao.Refresh
               .lblRecepcaoArquivo.Caption = "( " & sArquivo & " )"
               .lblRecepcaoArquivo.Refresh
          End If
     
     End With
     
End Sub

Public Sub RecAvisoDiferenca()
 
    Dim rstAvisoDif         As New ADODB.Recordset
    Dim rstMotivo           As New ADODB.Recordset
    Dim InsAviso            As New Custodia.Inserir
    Dim SelAviso            As New Custodia.Selecionar
    Dim DatFile             As Integer
    Dim vMotivo             As Integer
    Dim vPos                As Integer
    Dim sCodOcorrencia      As String
    Dim bArquivoReading     As Boolean
    Dim strNomeArquivo      As String
    Dim strNomeArquivoREA   As String
    Dim strNomeArquivoOK    As String
    Dim strNomeArquivoErr   As String
    Dim lRetorno            As Long
    Dim Reg                 As String * 179
    Dim OffSet              As Long
    Dim nAvisos             As Long
    Dim AD                  As AvisoDif_Reg
    Dim sstr                As String
    Dim sWhere              As String
    Dim Progress            As New clsProgressBar
    Dim sArquivoRecep       As String
    Dim aArquivoRecep()
    Dim sTituloTela         As String
    Dim NumBordero          As String
    Dim iCount              As Integer, aIndex As Integer
    Dim lngRegs             As Long
    
     On Error GoTo ErroLeitura
        
     sTituloTela = "Recepção Aviso de Diferença"
     nAvisos = 0
     vPos = 1
    
     If MsgBox("Inicializa processo de recepção", vbQuestion + vbYesNo, sTituloTela) = vbNo Then
          Exit Sub
     End If
    
     If Not LerArquivosDiretorio(g_Parametros.DiretorioRecepcao & "\", "CHADIF", sArquivoRecep) Then
          Exit Sub
     End If
     If Len(sArquivoRecep) = 0 Then
          Beep
          MsgBox "Não existe arquivo(s) para recepção do Aviso de Diferença.", vbInformation, sTituloTela
          Exit Sub
     End If
    
     ReDim aArquivoRecep(0)
     Screen.MousePointer = vbHourglass
    
     Do While True
          iCount = InStr(sArquivoRecep, ",")
          If iCount = 0 Then Exit Do
          aIndex = UBound(aArquivoRecep) + 1
          ReDim Preserve aArquivoRecep(aIndex)
          aArquivoRecep(aIndex) = Mid(sArquivoRecep, 1, (iCount - 1))
          sArquivoRecep = Mid(sArquivoRecep, (iCount + 1))
     Loop

     Call LabelRecepcao(LabelAcao.Apresenta)
    
     DatFile = FreeFile
    
     'Recepciona quantos arquivos existirem
     For iCount = 1 To aIndex
     
          Call LabelRecepcao(LabelAcao.Atualiza, iCount, aIndex, aArquivoRecep(iCount))
          
          strNomeArquivo = Trim(g_Parametros.DiretorioRecepcao) & "\" & aArquivoRecep(iCount)
          'Muda extensão do arquivo para (REA) qdo estiver sendo acessado para leitura
          strNomeArquivoREA = strNomeArquivo & "_READ"
          'Muda extensão do arquivo para (OK) qdo finalizado o processo
          strNomeArquivoOK = strNomeArquivo & "_OK"
          'Muda extensão do arquivo para (Erro) qdo ocorrer erro
          strNomeArquivoErr = strNomeArquivo & "_ERRO"
          
          bArquivoReading = False
          
          'Verifica se arq. existe no diretório e está sendo acessado por outro usuário
          If Dir(strNomeArquivo, vbDirectory) = "" Then
               
               'Verifica se arq. confirmação está sendo lido por outro usuário (.REA)
               If Dir(strNomeArquivoREA, vbDirectory) = "" Then

                    'Abre arquivo para verificar erro de inexistencia do diretório/Arquivo
                    Open strNomeArquivo For Input As #1
                    Close #1
                    Exit Sub
               Else
                    'Força abertura do arquivo para executar tratamento do erro
                    'Caso consiga abrir o arquivo, é sinal de que houve queda de execução
                    'e o arquivo ficou em modo Reading...
                    Open strNomeArquivoREA For Input Access Read As #1
                    Close #1
                    bArquivoReading = True
               End If
          End If
          
          'Renomeia o arquivo para extensão (.REA)
          If Not bArquivoReading Then
               Name strNomeArquivo As strNomeArquivoREA
          End If
    
          Open strNomeArquivoREA For Binary Access Read Lock Read Write As #DatFile
    
        
          OffSet = 1
    
          Get #DatFile, OffSet, Reg
    
          'Obtem o total de registros do arquivo de leitura
          If Not EOF(DatFile) Then
                'Inicia progress bar
                Progress.ValorMinimo = 0
                Progress.ValorMaximo = Fix(FileLen(strNomeArquivoREA) / Len(Reg))
                Progress.DescricaoProcesso = "Recepcionando Aviso de Diferença ..."
                Progress.InicializaProgressBar
                Progress.AtualizaBarra
                lngRegs = 0
          End If
          
          While Not EOF(DatFile)
              'Acumulador de registros lidos
              lngRegs = lngRegs + 1
              
              ' Se arquivo foi lido ok
              If Len(Reg) < 178 Then
                   MsgBox "Erro de Leitura", vbOKOnly + vbCritical, sTituloTela
                   GoTo FimLeituraComErro
              End If
              
              ' Ver rótulo do arquivo
              If Mid(Reg, 1, 6) <> "CHADIF" Then
                   MsgBox "Rótulo do Arquivo de Diferença Inválido.", vbOKOnly + vbCritical, sTituloTela
                   GoTo FimLeituraComErro
              End If
              
              ' Ver se CGC de terceira é válido
              If CStr(Mid(Reg, 7, 14)) <> FormataString(g_Parametros.CNPJ_Terceira, "0", 14, True) Then
                   MsgBox "CNPJ da Terceira Inválido", vbOKOnly + vbCritical, sTituloTela
                   GoTo FimLeituraComErro
              End If
                      
              ' Atribuir registos do arquivo
              AD.Num_Bordero = Mid(Reg, 21, 18)
              AD.CodigoCarteira = Mid(Reg, 39, 2)
              AD.DataOcorrencia = Mid(Reg, 41, 8)
              AD.CodigoOcorrencia = Mid(Reg, 49, 9)
              AD.Agencia = Mid(Reg, 58, 4)
              AD.Conta = Mid(Reg, 62, 7)
              AD.CodigoDevolucao = Mid(Reg, 69, 2)
              AD.CodigoCompensacao = Mid(Reg, 71, 3)
              AD.BancoEmitente = Mid(Reg, 74, 4)
              AD.AgenciaEmitente = Mid(Reg, 78, 4)
              AD.CcEmitente = Mid(Reg, 82, 11)
              AD.NrChequeEmitente = Mid(Reg, 93, 10)
              AD.TipoCheque = Mid(Reg, 103, 1)
              AD.TipoInscricao = Mid(Reg, 104, 2)
              AD.InscricaoEmitente = Mid(Reg, 106, 14)
              AD.DataDeposito = Mid(Reg, 120, 8)
              AD.Valor = Mid(Reg, 128, 13)
              AD.MotivoDevolucao = Mid(Reg, 175, 2)
              
              
              'Calcula Digito do bordero
              NumBordero = AD.Num_Bordero & RetornaDigitoModulo11Simplificado(Mid(AD.Num_Bordero, 13, 6))
             
             'AD.Gerado = False
              
              For vMotivo = 1 To 10
              
                    sCodOcorrencia = Mid(AD.MotivoDevolucao, vPos, 2)
                    
                   If sCodOcorrencia = "" Then
                      Exit For
                    End If
                    
                    
                    Set rstMotivo = g_cMainConnection.Execute(SelAviso.GetMotivoAD(sCodOcorrencia))
                    
                    Set rstAvisoDif = g_cMainConnection.Execute(SelAviso.GetAvisoDif(CLng(AD.DataOcorrencia), CLng(AD.CodigoOcorrencia), CInt(sCodOcorrencia)))
                    
                    If rstAvisoDif.EOF Then
                    
                      ' Gravar Registro do Aviso de Diferença
                      
                      Call g_cMainConnection.Execute(InsAviso.InsereAvisoDiferenca(CLng(AD.DataOcorrencia), _
                                                     CLng(AD.CodigoOcorrencia), rstMotivo!Descricao, CLng(AD.DataDeposito), _
                                                     NumBordero, CByte(AD.CodigoCarteira), _
                                                     CInt(AD.Agencia), CLng(AD.Conta), CInt(AD.CodigoDevolucao), _
                                                     CInt(AD.CodigoCompensacao), CInt(AD.BancoEmitente), _
                                                     CInt(AD.AgenciaEmitente), CDbl(AD.CcEmitente), _
                                                     CLng(AD.NrChequeEmitente), CByte(AD.TipoCheque), _
                                                     CByte(AD.TipoInscricao), AD.InscricaoEmitente, _
                                                     Format(Val(InserePonto(AD.Valor)), MASK_VALOR), 0, "T"), _
                                                     lRetorno, adCmdText)
                               
                        nAvisos = nAvisos + lRetorno
                        vPos = vPos + 2
                        
                    End If
              Next
              
              OffSet = OffSet + Len(Reg)
              
              Get #DatFile, OffSet, Reg
              vPos = 1
              
              'Atualiza Progress Bar
              Progress.AtualValue = lngRegs
              Progress.AtualizaBarra
        
          Wend
    
          Close #DatFile

          'Renomeia o arquivo para extensão (.OK)
          Name strNomeArquivoREA As strNomeArquivoOK

     Next
     
     'Encerra progress bar
     Set Progress = Nothing
    
     'Fecha tela apresentação de arquivos
     Call LabelRecepcao(LabelAcao.Finaliza)
    
     Screen.MousePointer = vbDefault
     MsgBox "Foram Processados " & CStr(nAvisos) & " Avisos de Diferença.", vbOKOnly + vbExclamation, sTituloTela
     Exit Sub
    
FimLeitura:

     'Fecha tela apresentação de arquivos
     Call LabelRecepcao(LabelAcao.Finaliza)
    
     Close #DatFile
     'Encerra progress bar
     Set Progress = Nothing
     
     Screen.MousePointer = vbDefault
     Exit Sub
    
FimLeituraComErro:

     Close #DatFile

     If UBound(aArquivoRecep) >= 1 Then
          'Renomeia o arquivo para finalização com erro
          Name strNomeArquivoREA As strNomeArquivoErr
     End If
     GoTo FimLeitura

ErroLeitura:
     Beep
     Screen.MousePointer = vbDefault
     
     'Tratamento para diretório não localizado
     If Err.Number = 76 Then
          MsgBox "Não localizado o diretório com o arquivo de Aviso de Diferença." & vbCrLf & vbCrLf & _
                    "Favor verificar diretório em parâmetros.", vbCritical, sTituloTela
          GoTo FimLeitura
     End If
     If Err.Number = 53 Then
          MsgBox "Não localizado o arquivo de Aviso de Diferença." & vbCrLf & vbCrLf & _
                    "Favor verificar e tentar novamente.", vbCritical, sTituloTela
          GoTo FimLeitura
     End If
     If Err.Number = 55 Then
          MsgBox "Arquivo com Aviso de Diferença em utilização por outro usuário. Favor verificar !", vbCritical, sTituloTela
          GoTo FimLeitura
     End If

    MsgBox "Erro na Leitura do Arquivo de Aviso de Diferença.", vbOKOnly + vbCritical, sTituloTela

    GoTo FimLeitura

End Sub

Public Sub FusaoAutomatica()
 
    Dim rstChFusao          As New ADODB.Recordset
    Dim DatFile             As Integer
    Dim Sel_Fusao           As New Custodia.Selecionar
    Dim Atu_Fusao           As New Custodia.Atualizar
    Dim lRetorno            As Long
    Dim Reg                 As String * 152
    Dim OffSet              As Long
    Dim nCheques            As Long
    Dim sstr                As String
    Dim sWhere              As String
    Dim strNomeArquivo      As String
    Dim sCMC7               As String
    Dim Progress            As New clsProgressBar
    Dim lngRegs             As Long
    Dim sTituloTela         As String
 
    
    
    On Error GoTo ErroLeitura

     sTituloTela = "Fusão Automática"
     nCheques = 0
     
     strNomeArquivo = Trim(FusaoDialog.TxtFusao.Text)
     
     If Not FileExist(strNomeArquivo) Then
        Beep
        MsgBox "Arquivo de Fusão Não Encontrado", vbExclamation, sTituloTela
        Exit Sub
     End If
     
     Screen.MousePointer = vbHourglass
     
     DatFile = FreeFile
          
     Open strNomeArquivo For Binary Access Read Lock Read Write As #DatFile
              
     OffSet = 1
      
     Get #DatFile, OffSet, Reg
          
          
          'Obtem o total de registros do arquivo de leitura
            If Not EOF(DatFile) Then
                'Inicia progress bar
                Progress.ValorMinimo = 1
                
                Progress.ValorMaximo = Fix(FileLen(strNomeArquivo) / Len(Reg)) - 1
                
                Progress.DescricaoProcesso = "Processando a Fusão Automática"
                Progress.InicializaProgressBar
                Progress.AtualizaBarra
                lngRegs = 0
            End If
          
           ' Se arquivo foi lido ok
            If Len(Reg) < 150 Then
                MsgBox "Erro de Leitura", vbOKOnly + vbCritical, sTituloTela
                GoTo FimLeituraComErro
            End If
              
            ' Ver rótulo do arquivo
            If Mid(Reg, 1, 2) <> "HD" Then
                MsgBox "Rótulo do Arquivo Inválido.", vbOKOnly + vbCritical, sTituloTela
                GoTo FimLeituraComErro
            End If
              
            ' Verifica Código da Terceira
            If CStr(Mid(Reg, 3, 4)) <> g_Parametros.Codigo_Terceira Then
               MsgBox "Código da Terceira Inválido", vbOKOnly + vbCritical, sTituloTela
               GoTo FimLeituraComErro
            End If
          
            OffSet = OffSet + Len(Reg)
              
            Get #DatFile, OffSet, Reg
          
          While Not EOF(DatFile)
          
              'Acumulador de registros lidos
              lngRegs = lngRegs + 1
              
              
              ' Atribui cmc7 do arquivo
              sCMC7 = Trim(Mid(Reg, 17, 30))
                            
              Set rstChFusao = g_cMainConnection.Execute(Sel_Fusao.GetChequeFusao(sCMC7, Geral.DataProcessamento))
              
                If Not rstChFusao.EOF Then
                    If rstChFusao!fusao Then
                        MsgBox "Documento já Processado" & vbCrLf & vbCrLf & "Cheque nº " & Mid(sCMC7, 12, 6) & "  Banco " & Mid(sCMC7, 1, 3), vbExclamation + vbOKOnly
                    Else
                        Call g_cMainConnection.Execute(Atu_Fusao.AtualizaFusao(sCMC7, Geral.DataProcessamento), lRetorno, adCmdText)
                        If lRetorno = 0 Then
                            Err.Raise 998, App.Title, "Erro ao Atualizar - Fusão"
                            Exit Sub
                        Else
                            nCheques = nCheques + lRetorno
                        End If
                        
                    End If
            
                Else
                        
                    MsgBox "Documento não encontrado", vbExclamation + vbOKOnly, sTituloTela
        
                End If


              
                OffSet = OffSet + Len(Reg)
              
                Get #DatFile, OffSet, Reg
              
                'Atualiza Progress Bar
                Progress.AtualValue = lngRegs
                Progress.AtualizaBarra
              
          Wend
      
          Close #DatFile
          
               
     'Encerra progress bar
     Set Progress = Nothing
     
     
     Screen.MousePointer = vbDefault
     MsgBox vbCrLf & CStr(lngRegs) & " Cheque(s) Foram Processados " & vbCrLf & CStr(nCheques) & " Cheque(s) Foram Fundidos", vbOKOnly + vbInformation, sTituloTela
     
     
     Set FusaoDialog = Nothing
     Exit Sub
     
FimLeitura:
     
     Close #DatFile
     'Encerra progress bar
     Set Progress = Nothing
     
     Screen.MousePointer = vbDefault
     Exit Sub
    
FimLeituraComErro:

     Close #DatFile
     GoTo FimLeitura
     
ErroLeitura:
     Beep
     Screen.MousePointer = vbDefault
     
     'Tratamento para diretório não localizado
     If Err.Number = 76 Then
          MsgBox "Diretório do Arquivo Não Localizado." & vbCrLf & vbCrLf & _
                    "Favor verificar diretório em parâmetros.", vbCritical, sTituloTela
          GoTo FimLeitura
     End If
     If Err.Number = 53 Then
          MsgBox "Arquivo Não Localizado." & vbCrLf & vbCrLf & _
                    "Favor verificar e tentar novamente.", vbCritical, sTituloTela
          GoTo FimLeitura
     End If
     If Err.Number = 55 Then
          MsgBox "Arquivo em utilização por outro usuário. Favor verificar!", vbCritical, sTituloTela
          GoTo FimLeitura
     End If

    MsgBox "Erro na Leitura do Arquivo.", vbOKOnly + vbCritical, sTituloTela

    GoTo FimLeitura

End Sub

Public Sub RecRegraGP()
 
    Dim rstRegraGP          As New ADODB.Recordset
    Dim InsRegraGP          As New Custodia.Inserir
    Dim SelRegraGP          As New Custodia.Selecionar
    Dim AtuParametros       As New Custodia.Atualizar
    Dim DatFile             As Integer
    Dim bArquivoReading     As Boolean
    Dim strNomeArquivo      As String
    Dim strNomeArquivoREA   As String
    Dim strNomeArquivoOK    As String
    Dim strNomeArquivoErr   As String
    Dim lRetorno            As Long
    Dim Reg                 As String * 44
    Dim OffSet              As Long
    Dim nRegraGP            As Long
    Dim GP                  As RegraGP_Reg
    Dim sstr                As String
    Dim sWhere              As String
    Dim Progress            As New clsProgressBar
    Dim sArquivoRecep       As String
    Dim aArquivoRecep()
    Dim sTituloTela         As String
    Dim iCount              As Integer, aIndex As Integer
    Dim lngRegs             As Long
    
     On Error GoTo ErroLeitura
        
     sTituloTela = "Recepção Regra do GP"
     nRegraGP = 0
     
    
     If MsgBox("Inicializa processo de recepção", vbQuestion + vbYesNo, sTituloTela) = vbNo Then
          Exit Sub
     End If
    
     If Not LerArquivosDiretorio(g_Parametros.DiretorioRecepcao & "\", "CHREGP", sArquivoRecep) Then
          Exit Sub
     End If
     If Len(sArquivoRecep) = 0 Then
          Beep
          MsgBox "Não existe arquivo(s) para recepção da Regra do GP.", vbInformation, sTituloTela
          Exit Sub
     End If
    
     ReDim aArquivoRecep(0)
     Screen.MousePointer = vbHourglass
    
     Do While True
          iCount = InStr(sArquivoRecep, ",")
          If iCount = 0 Then Exit Do
          aIndex = UBound(aArquivoRecep) + 1
          ReDim Preserve aArquivoRecep(aIndex)
          aArquivoRecep(aIndex) = Mid(sArquivoRecep, 1, (iCount - 1))
          sArquivoRecep = Mid(sArquivoRecep, (iCount + 1))
     Loop

     Call LabelRecepcao(LabelAcao.Apresenta)
    
     DatFile = FreeFile
    
     'Recepciona quantos arquivos existirem
     For iCount = 1 To aIndex
     
          Call LabelRecepcao(LabelAcao.Atualiza, iCount, aIndex, aArquivoRecep(iCount))
          
          strNomeArquivo = Trim(g_Parametros.DiretorioRecepcao) & "\" & aArquivoRecep(iCount)
          'Muda extensão do arquivo para (REA) qdo estiver sendo acessado para leitura
          strNomeArquivoREA = strNomeArquivo & "_READ"
          'Muda extensão do arquivo para (OK) qdo finalizado o processo
          strNomeArquivoOK = strNomeArquivo & "_OK"
          'Muda extensão do arquivo para (Erro) qdo ocorrer erro
          strNomeArquivoErr = strNomeArquivo & "_ERRO"
          
          bArquivoReading = False
          
          'Verifica se arq. existe no diretório e está sendo acessado por outro usuário
          If Dir(strNomeArquivo, vbDirectory) = "" Then
               
               'Verifica se arq. confirmação está sendo lido por outro usuário (.REA)
               If Dir(strNomeArquivoREA, vbDirectory) = "" Then

                    'Abre arquivo para verificar erro de inexistencia do diretório/Arquivo
                    Open strNomeArquivo For Input As #1
                    Close #1
                    Exit Sub
               Else
                    'Força abertura do arquivo para executar tratamento do erro
                    'Caso consiga abrir o arquivo, é sinal de que houve queda de execução
                    'e o arquivo ficou em modo Reading...
                    Open strNomeArquivoREA For Input Access Read As #1
                    Close #1
                    bArquivoReading = True
               End If
          End If
          
          'Renomeia o arquivo para extensão (.REA)
          If Not bArquivoReading Then
               Name strNomeArquivo As strNomeArquivoREA
          End If
    
          Open strNomeArquivoREA For Binary Access Read Lock Read Write As #DatFile
    
        
          OffSet = 1
    
          Get #DatFile, OffSet, Reg
    
          'Obtem o total de registros do arquivo de leitura
          If Not EOF(DatFile) Then
                'Inicia progress bar
                Progress.ValorMinimo = 0
                Progress.ValorMaximo = Fix(FileLen(strNomeArquivoREA) / Len(Reg))
                Progress.DescricaoProcesso = "Recepcionando Regra do GP..."
                Progress.InicializaProgressBar
                Progress.AtualizaBarra
                lngRegs = 0
          End If
          
          While Not EOF(DatFile)
              'Acumulador de registros lidos
              lngRegs = lngRegs + 1
              
              ' Se arquivo foi lido ok
              If Len(Reg) < 42 Then
                   MsgBox "Erro de Leitura", vbOKOnly + vbCritical, sTituloTela
                   GoTo FimLeituraComErro
              End If
              
              ' Ver rótulo do arquivo
              If Mid(Reg, 1, 6) <> "CHREGP" Then
                   MsgBox "Rótulo da Regra do GP Inválido.", vbOKOnly + vbCritical, sTituloTela
                   GoTo FimLeituraComErro
              End If
              
              ' Ver se CGC de terceira é válido
'              If CStr(Mid(Reg, 7, 14)) <> g_Parametros.CNPJ_Terceira Then
'                   MsgBox "CNPJ da Terceira Inválido", vbOKOnly + vbCritical, sTituloTela
'                   GoTo FimLeituraComErro
'              End If
              
              ' Ver se Códido do Produto = 43265
              If CStr(Mid(Reg, 29, 5)) <> "43265" Then
                   MsgBox "Códido do Produto Inválido", vbOKOnly + vbCritical, sTituloTela
                   GoTo FimLeituraComErro
              End If
                      
              ' Atribuir registos do arquivo
              GP.DataProcessamento = Mid(Reg, 21, 8)
              GP.CodigoProduto = Mid(Reg, 29, 5)
              GP.CodigoRegra = Mid(Reg, 34, 4)
              GP.QtdDias = Mid(Reg, 38, 5)
              
              Set rstRegraGP = g_cMainConnection.Execute(SelRegraGP.GetRegraGP(CLng(GP.DataProcessamento), _
              GP.CodigoProduto, GP.CodigoRegra))
                    
                    If rstRegraGP.EOF Then
                    
                      ' Gravar Registro da Regra do GP
                      Call g_cMainConnection.Execute(InsRegraGP.InsereRegraGP(CLng(GP.DataProcessamento), _
                                                     CLng(GP.CodigoProduto), _
                                                     CLng(GP.CodigoRegra), _
                                                     CLng(GP.QtdDias)), _
                                                     lRetorno, adCmdText)
                               
                        
                        
                        ' Se Regra = (0664) atualiza Parametro no Sistema
                        If CStr(GP.CodigoRegra) = "0664" Then
                            Call g_cMainConnection.Execute(AtuParametros.AtualizaDiasCheques(CLng(GP.DataProcessamento), _
                                                      CInt(GP.QtdDias)), _
                                                      lRetorno, adCmdText)
                                                                                  
                            nRegraGP = nRegraGP + lRetorno
                            
                        End If
                        
                    End If
              
              OffSet = OffSet + Len(Reg)
              
              Get #DatFile, OffSet, Reg
              
              'Atualiza Progress Bar
              Progress.AtualValue = lngRegs
              Progress.AtualizaBarra
        
          Wend
    
          Close #DatFile

          'Renomeia o arquivo para extensão (.OK)
          Name strNomeArquivoREA As strNomeArquivoOK

     Next
     
     'Encerra progress bar
     Set Progress = Nothing
    
     'Fecha tela apresentação de arquivos
     Call LabelRecepcao(LabelAcao.Finaliza)
    
     Screen.MousePointer = vbDefault
     MsgBox "Foram Processados " & CStr(nRegraGP) & " Regra do GP.", vbOKOnly + vbExclamation, sTituloTela
     Exit Sub
    
FimLeitura:

     'Fecha tela apresentação de arquivos
     Call LabelRecepcao(LabelAcao.Finaliza)
    
     Close #DatFile
     'Encerra progress bar
     Set Progress = Nothing
     
     Screen.MousePointer = vbDefault
     Exit Sub
    
FimLeituraComErro:

     Close #DatFile

     If UBound(aArquivoRecep) >= 1 Then
          'Renomeia o arquivo para finalização com erro
          Name strNomeArquivoREA As strNomeArquivoErr
     End If
     GoTo FimLeitura

ErroLeitura:
     Beep
     Screen.MousePointer = vbDefault
     
     'Tratamento para diretório não localizado
     If Err.Number = 76 Then
          MsgBox "Não localizado o diretório com o arquivo Regra do GP." & vbCrLf & vbCrLf & _
                    "Favor verificar diretório em parâmetros.", vbCritical, sTituloTela
          GoTo FimLeitura
     End If
     If Err.Number = 53 Then
          MsgBox "Não localizado o arquivo de Aviso de Diferença." & vbCrLf & vbCrLf & _
                    "Favor verificar e tentar novamente.", vbCritical, sTituloTela
          GoTo FimLeitura
     End If
     If Err.Number = 55 Then
          MsgBox "Arquivo com Regra do GP em utilização por outro usuário. Favor verificar !", vbCritical, sTituloTela
          GoTo FimLeitura
     End If

    MsgBox "Erro na Leitura do Arquivo de Regra do GP.", vbOKOnly + vbCritical, sTituloTela

    GoTo FimLeitura

End Sub
