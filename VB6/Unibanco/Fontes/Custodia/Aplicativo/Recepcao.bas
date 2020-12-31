Attribute VB_Name = "Recepcao"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''
''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private myrecord   As TpRecordRejeicao

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'         Define type de registro do arquivo de Rejei��o           '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Type TpRecordRejeicao
     Rot_Mensag          As String * 6       'R�tulo da mensagem (REJEIC)
     CGC_Ender           As String * 14      'CGC de Endere�amento do contrato de servi�o
     Num_Bordero         As String * 18      'N�mero do border�
     Dat_Deposito        As String * 8       'Data de Entrega/Dep�sito (AAAAMMDD)
     Num_BcoCliente      As String * 4       'N�mero do Banco Cliente
     Num_AgenCliente     As String * 4       'N�mero da Ag�ncia Cliente
     Num_ContaCliente    As String * 7       'N�mero da Conta Corrente Cliente
     Num_Cheque          As String * 7       'N�mero do Cheque emitente
     Rot_Original        As String * 6       'R�tulo Original
     Qtd_Erros           As String * 2       'Quantidade de Erros
     Cod_Erros           As String * 114     'C�digo de Erros (38 ocorr�ncias de 3 bytes cada)
     Tip_Identificacao   As String * 2       'Identificador de Campo (CM ou "  " e demais dados do cheque)
     Cod_CMC7            As String * 34      'C�digo do CMC7
     Cod_OfiAdv          As String * 25      'C�digo do OFI/ADV
     CrLf                As String * 2       'OK
End Type

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'         Define type de inconsist�ncia na recep��o de arquivos (M�dulo Principal)   '
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
'         Type para inser��o de dados na tabela rejeicaoremessa                      '
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
'         Define type de inconsist�ncia na recep��o de confirma��o (M�dulo Principal)     '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Type tpRecepcaoConfir
     Num_Bordero    As String * 19
     Banco          As String * 4
     Agencia        As String * 4
     Conta          As String * 7
     Inconsistencia As String * 50
End Type

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'         Define type de registro do arquivo de Confirma��o           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Type RecordConfirmacao
     Rot_Mensag          As String * 6       'R�tulo da mensagem (CONFIR)
     CGC_Ender           As String * 14      'CGC de Endere�amento do Contrato de servi�o
     Num_Bordero         As String * 18      'N�mero do border�
     Dat_Processamento   As String * 8       'Data de Processamento (AAAAMMDD)
     Num_Banco           As String * 4       'N�mero do Banco (0409)
     Num_Agencia         As String * 4       'N�mero da Ag�ncia do cliente
     Num_ContaCorrente   As String * 7       'N�mero da Conta Corrente do cliente
     Reg_Vago            As String * 7       'Registro Vago (Zeros)
     Rot_Original        As String * 6       'R�tulo Original (CHINBO)
     Qtd_Erros           As String * 2       'Quantidade de Erros
     Cod_Erro            As String * 3       'C�digo do Erro
     Cod_OfiAdv          As String * 25      'C�digo do OFI/ADV
     CrLf                As String * 2       'OK
     'Controle            As String * 1       'Byte de controle para fim de arquivo (LineFeed)
End Type

'Registro de Cheque Data Boa -  R�tulo CHDBOA
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
'         Define type de registro do arquivo de Instru��es (CHRDTV)   '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Type RecordInstrucoes
     Rot_Mensag          As String * 6       'R�tulo da mensagem (CHRDTV)
     CGC_Ender           As String * 14      'CGC de Endere�amento do Contrato de servi�o
     Cod_Carteira        As String * 2       'C�digo da carteira (16-Cust�dia) / (17-Cau��o)
     Dat_DepAnterior     As String * 8       'Data de Dep�sito Anterior (AAAAMMDD)
     Num_Banco           As String * 4       'N�mero de Banco do emitente
     Num_Agencia         As String * 4       'N�mero de Ag�ncia do emitente
     Num_Cheque          As String * 6       'N�mero de Cheque do emitente
     Dat_DepNova         As String * 8       'Data de Dep�sito Nova (AAAAMMDD)
     Cod_NossoNumero     As String * 11      'Nosso N�mero (Controle do banco)
     Cod_compensacao     As String * 3       'C�digo de compensa��o
     Num_ContaCorrente   As String * 11      'N�mero da Conta Corrente do emitente
     Num_Bordero         As String * 18      'N�mero do border�
     ValorCheque         As String * 15      'Valor de Cheque
     CrLf                As String * 2
End Type

' Registro de Aviso de Diferen�a -  R�tulo CHADIF
Private Type AvisoDif_Reg
    Rot_Mensag                      As String * 6   'R�tulo da mensagem (CHADIF)
    CGC_Ender                       As String * 14  'CGC de Endere�amento do Contrato de servi�o
    Num_Bordero                     As String * 18  'N�mero do border�
    CodigoCarteira                  As String * 2   'C�digo da carteira (16-Cust�dia) / (17-Cau��o)
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


' Registro da Regra do GP -  R�tulo CHREGP
Private Type RegraGP_Reg
    Rot_Mensag                      As String * 6   'R�tulo da mensagem (CHREPG)
    CGC_Ender                       As String * 14  'CGC de Endere�amento do Contrato de servi�o
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
                         'Se arquivo foi aberto para leitura (READ) ent�o retirar extens�o
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
          MsgBox "Diret�rio para recep��o n�o localizado!" & vbCrLf & vbCrLf & "Verifique diret�rio de recep��o em parametros do sistema", vbCritical, App.Title
          Exit Function
     End If

Err_LerArquivosDiretorio:
     'Verifica se arquivo n�o tem permiss�o de abertura
     If Err.Number = 70 Or Err.Number = 52 Or Err.Number = 75 Or Err.Number = 76 Then
          Resume Next
     End If
     
     Beep
     Close
     Principal.filArquivosRecepcao.Normal = False
     MsgBox "Erro na leitura do diret�rio ( " & sDiretorio & " )", vbCritical, App.Title
     
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
Dim strDigBordero        As String      'D�gito verificador do border�
Dim Progress             As New clsProgressBar
Dim sTituloTela          As String
Dim sArquivoRecep        As String
Dim aArquivoRecep()
Dim iCount               As Integer, aIndex As Integer

On Error GoTo Erro_RecConfRemessa
     
     sTituloTela = "Recep��o da Confirma��o de remessa"
     
     bTransacaoAberta = False
     
     If MsgBox("Inicializa processo de recep��o", vbQuestion + vbYesNo, sTituloTela) = vbNo Then
          Set rsRetorno = Nothing
          Exit Sub
     End If

     If Not LerArquivosDiretorio(g_Parametros.DiretorioRecepcao & "\", "CONFIR", sArquivoRecep) Then
          Set rsRetorno = Nothing
          Exit Sub
     End If
     If Len(sArquivoRecep) = 0 Then
          Beep
          MsgBox "N�o existe arquivo(s) para recep��o referente � Confirma��o de Remessa", vbInformation, sTituloTela
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
          'Muda extens�o do arquivo para (REA) qdo estiver sendo acessado para leitura
          strNomeArquivoREA = strNomeArquivo & "_READ"
          'Muda extens�o do arquivo para (OK) qdo finalizado o processo
          strNomeArquivoOK = strNomeArquivo & "_OK"
          'Muda extens�o do arquivo para (Erro) qdo ocorrer erro
          strNomeArquivoErr = strNomeArquivo & "_ERRO"
          
          bArquivoReading = False
          
          'Verifica se arq. confirma��o existe no diret�rio e est� sendo acessado por outro usu�rio
          If Dir(strNomeArquivo, vbDirectory) = "" Then
               
               'Verifica se arq. confirma��o est� sendo lido por outro usu�rio (.REA)
               If Dir(strNomeArquivoREA, vbDirectory) = "" Then

                    'Abre arquivo para verificar erro de inexistencia do diret�rio/Arquivo
                    Open strNomeArquivo For Input As #1
                    Close #1
                    Exit Sub
               Else
                    'For�a abertura do arquivo para executar tratamento do erro
                    'Caso consiga abrir o arquivo, � sinal de que houve queda de execu��o
                    'e o arquivo ficou em modo Reading...
                    Open strNomeArquivoREA For Input Access Read As #1
                    Close #1
                    bArquivoReading = True
               End If
          End If
          
          'Renomeia o arquivo para extens�o (.REA)
          If Not bArquivoReading Then
               Name strNomeArquivo As strNomeArquivoREA
          End If
          
          'Abre o arquivo CONFIRMA��O DE REMESSA com extens�o (.REA)
          Open strNomeArquivoREA For Random As #1 Len = Len(myrecord)
          
          'Obtem o total de bytes do arquivo de confirma��o de remessa
          lngTotalRegistros = FileLen(strNomeArquivoREA) / Len(myrecord)
          
         'Inicializa Progress Bar
          Progress.ValorMinimo = 1
          Progress.ValorMaximo = lngTotalRegistros
          Progress.DescricaoProcesso = "Recepcionando confirma��o de remessa ..."
          Progress.InicializaProgressBar
          Progress.AtualizaBarra
          
          'Ler cada registro e atualizar tabela de cheque
          For lngRegs = 1 To lngTotalRegistros
               'Ler linha arquivo confirma��o de remessa
               Get #1, lngRegs, myrecord
               'Atualiza tabela cheque somente se registro n�o cont�m erro
               If CInt(myrecord.Qtd_Erros) = 0 Then
     
                    ' Calcula d�gito verificador do border�
                    ' Para Calcular o digito considerar: Mid(NumeroBordero, 13, 6)
                    
                    
                    strDigBordero = RetornaDigitoModulo11Simplificado(Mid(myrecord.Num_Bordero, 13, 6))
     
                    Set rsRetorno = Nothing
                    'Obtem n�mero do IdBordero
                    Set rsRetorno = g_cMainConnection.Execute(Proc_Selecionar.GetNumeroIdBordero( _
                                                       CLng(myrecord.Dat_Processamento), _
                                                       myrecord.Num_Bordero & strDigBordero), lngRetorno, adCmdText)
                    
                    If rsRetorno.EOF Then
                         Call GeraInconsConfRemessa("N�o localizado border�", myrecord, recepcao, strDigBordero)
                    Else
                         lngIdBordero = rsRetorno(0).Value
                         
                         Set rsRetorno = Nothing
                         
                         'Inicializa transa��o
                         g_cMainConnection.BeginTrans
                         bTransacaoAberta = True
                         
                         'Atualiza status Border� correspondente para (E)Confirmado
                         Set rsRetorno = g_cMainConnection.Execute(Proc_Atualizar.AtualizaStatusBordero( _
                                                            CLng(myrecord.Dat_Processamento), _
                                                            lngIdBordero, "E"), lngRetorno, adCmdText)
                         
                         If lngRetorno <> 1 Then
                              Call GeraInconsConfRemessa("N�o foi poss�vel atualizar border� para confirmado", myrecord, recepcao, strDigBordero)
                              
                              'Cancela transa��o
                              g_cMainConnection.RollbackTrans
                              bTransacaoAberta = False
                         Else
                    
                              Set rsRetorno = Nothing
                              
                              'Atualiza status cheque correspondente para (E)Confirmado
                              Set rsRetorno = g_cMainConnection.Execute(Proc_Atualizar.AtualizaChequesRemessa( _
                                                                 CLng(myrecord.Dat_Processamento), _
                                                                 lngIdBordero))
                              If lngRetorno < 1 Then
                                   Call GeraInconsConfRemessa("N�o foi poss�vel atualizar cheque para confirmado", myrecord, recepcao, strDigBordero)
                              
                                   'Cancela transa��o
                                   g_cMainConnection.RollbackTrans
                                   bTransacaoAberta = False
                              
                              End If
                              
                              
                              Set rsRetorno = Nothing
                              
                              'Atualiza status Data de Dep�sito correspondente para (1)Confirmado
                              Set rsRetorno = g_cMainConnection.Execute(Proc_Atualizar.AtualizaStatusDataDeposito( _
                                                                 CLng(myrecord.Dat_Processamento), _
                                                                 lngIdBordero))
                              If lngRetorno < 1 Then
                                   Call GeraInconsConfRemessa("N�o foi poss�vel atualizar Data de Dep�sito para confirmada", myrecord, recepcao, strDigBordero)
                              
                                   'Cancela transa��o
                                   g_cMainConnection.RollbackTrans
                                   bTransacaoAberta = False
                              Else
                                   'Finaliza transa��o
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
          
          'Fecha arquivo de confirma��o
          Close #1

          'Renomeia o arquivo para extens�o (.OK)
          Name strNomeArquivoREA As strNomeArquivoOK
     
     Next
     
     ReDim aArquivoRecep(0)
     
     'Verifica se houve ocorrencia na recep��o
     If UBound(recepcao) > 0 Then
          
          Set rsRetorno = Nothing
          'Obtem data e hora do servidor (Chave para a tabela INCONSISTENCIA)
          Set rsRetorno = g_cMainConnection.Execute(Proc_Selecionar.GetDataHoraServidor(), lngRetorno, adCmdText)
          
          If lngRetorno <> 0 Then
               Beep
               MsgBox "Problema na gera��o do relat�rio de inconsist�ncias!", vbCritical, sTituloTela
               GoTo Sair
          End If
          
          strDataChave = Format(rsRetorno!Data, "yyyymmdd")
          strHoraChave = rsRetorno!Hora
          
          For lngCount = 1 To UBound(recepcao)
               'Insere registro com dados de inconsist�ncia
               Call g_cMainConnection.Execute(Proc_Inserir.InsereInconsistencia( _
                                                  strDataChave, _
                                                  strHoraChave, _
                                                  recepcao(lngCount).Num_Bordero & _
                                                  recepcao(lngCount).Banco & _
                                                  recepcao(lngCount).Agencia & _
                                                  recepcao(lngCount).Conta & _
                                                  recepcao(lngCount).Inconsistencia))

          Next
     
'          'Imprime rela��o de ocorr�ncias do processo de recep��o
'          Call ImprimirInconsistencia("RelConfirmacaoRemessa", "Ocorr�ncias da confirma��o de remessa", myrecord.Dat_Processamento, strDataChave, strHoraChave)
          
          'Fecha tela apresenta��o de arquivos
          Call LabelRecepcao(LabelAcao.Finaliza)
          MsgBox Space(32) & "A T E N � � O" & vbCrLf & vbCrLf & _
                 "Finalizada a recep��o com ocorr�ncia(s).  Verifique !", vbCritical, sTituloTela
                 
          'Remove registro com dados de inconsist�ncia
          Call g_cMainConnection.Execute(Proc_Excluir.RemoveInconsistencia(strDataChave, _
                                                                           strHoraChave), _
                                                                           lngRetorno, adCmdText)
     Else
          
          Call LabelRecepcao(LabelAcao.Finaliza)
          
          MsgBox "Finalizada a confirma��o de remessa de cheques.", vbInformation, sTituloTela
     End If
     
     
Sair:
     Call LabelRecepcao(LabelAcao.Finaliza)
     
     Set Progress = Nothing
     
     Screen.MousePointer = vbDefault
     Set rsRetorno = Nothing
     
     Exit Sub
     
     
Erro_RecConfRemessa:

     Beep
     
     'Cancela transa��o
     If bTransacaoAberta Then g_cMainConnection.RollbackTrans: bTransacaoAberta = False

     'Tratamento para diret�rio n�o localizado
     If Err.Number = 76 Then
          MsgBox "N�o localizado o diret�rio com o arquivo de confirma��o." & vbCrLf & vbCrLf & _
                    "Favor verificar diret�rio em par�metros.", vbCritical, sTituloTela
          GoTo Sair
     End If
     If Err.Number = 53 Then
          MsgBox "N�o localizado o arquivo de confirma��o." & vbCrLf & vbCrLf & _
                    "Favor verificar e tentar novamente.", vbCritical, sTituloTela
          GoTo Sair
     End If
     If Err.Number = 55 Then
          MsgBox "Arquivo de confirma��o em utiliza��o por outro usu�rio. Favor verificar !", vbCritical, sTituloTela
          GoTo Sair
     End If

     'Fecha arquivo de confirma��o
     Close #1

     If UBound(aArquivoRecep) >= 1 Then
          'Renomeia o arquivo para finaliza��o com erro
          Name strNomeArquivoREA As strNomeArquivoErr
     End If
     
     MsgBox Err.Description, vbCritical, sTituloTela
     GoTo Sair

End Sub

Private Sub GeraInconsConfRemessa(strDescrMsg As String, myrecord As RecordConfirmacao, recepcao() As tpRecepcaoConfir, ByVal strDigBordero As String)

Dim lngCount As Long
     
     ReDim Preserve recepcao(UBound(recepcao) + 1)
     lngCount = UBound(recepcao)

     'Acumula em type os border�s com ocorr�ncia
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
     
   'Acumula em type os border�s com ocorr�ncia
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

     sTituloTela = "Recep��o Movimento Data Boa"
     nCheques = 0
        
     If MsgBox("Inicializa processo de recep��o", vbQuestion + vbYesNo, sTituloTela) = vbNo Then
          Exit Sub
     End If
    
     If Not LerArquivosDiretorio(g_Parametros.DiretorioRecepcao & "\", "CHDBOA", sArquivoRecep) Then
          Exit Sub
     End If
     If Len(sArquivoRecep) = 0 Then
          Beep
          MsgBox "N�o existe arquivo(s) para recep��o do Movimento de Data Boa.", vbInformation, sTituloTela
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
          'Muda extens�o do arquivo para (REA) qdo estiver sendo acessado para leitura
          strNomeArquivoREA = strNomeArquivo & "_READ"
          'Muda extens�o do arquivo para (OK) qdo finalizado o processo
          strNomeArquivoOK = strNomeArquivo & "_OK"
          'Muda extens�o do arquivo para (Erro) qdo ocorrer erro
          strNomeArquivoErr = strNomeArquivo & "_ERRO"
          
          bArquivoReading = False
          
          'Verifica se arq. confirma��o existe no diret�rio e est� sendo acessado por outro usu�rio
          If Dir(strNomeArquivo, vbDirectory) = "" Then
               
               'Verifica se arq. confirma��o est� sendo lido por outro usu�rio (.REA)
               If Dir(strNomeArquivoREA, vbDirectory) = "" Then

                    'Abre arquivo para verificar erro de inexistencia do diret�rio/Arquivo
                    Open strNomeArquivo For Input As #1
                    Close #1
                    Exit Sub
               Else
                    'For�a abertura do arquivo para executar tratamento do erro
                    'Caso consiga abrir o arquivo, � sinal de que houve queda de execu��o
                    'e o arquivo ficou em modo Reading...
                    Open strNomeArquivoREA For Input Access Read As #1
                    Close #1
                    bArquivoReading = True
               End If
          End If
          
          'Renomeia o arquivo para extens�o (.REA)
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
              
              ' Ver r�tulo do arquivo
              If Mid(Reg, 1, 6) <> "CHDBOA" Then
                   MsgBox "R�tulo do Arquivo de Cheques da Data Boa Inv�lido.", vbOKOnly + vbCritical, sTituloTela
                   GoTo FimLeituraComErro
              End If
              
              ' Ver se CGC de terceira � v�lido
              If CStr(Mid(Reg, 7, 14)) <> Format(g_Parametros.CNPJ_Terceira, String(14, "0")) Then
                 MsgBox "Endere�amento da Terceira Inv�lido", vbOKOnly + vbCritical, sTituloTela
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
              
              'Calcula d�gito verificador do border�
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
          
          'Renomeia o arquivo para extens�o (.OK)
          Name strNomeArquivoREA As strNomeArquivoOK
          
     Next
     
     'Encerra progress bar
     Set Progress = Nothing
     
     'Fecha tela apresenta��o de arquivos
     Call LabelRecepcao(LabelAcao.Finaliza)
     
     Screen.MousePointer = vbDefault
     MsgBox "Foram Processados " & CStr(nCheques) & " Cheque(s) da Data Boa.", vbOKOnly + vbExclamation, sTituloTela
     Exit Sub
     
FimLeitura:
     'Fecha tela apresenta��o de arquivos
     Call LabelRecepcao(LabelAcao.Finaliza)

     Close #DatFile
     'Encerra progress bar
     Set Progress = Nothing
     
     Screen.MousePointer = vbDefault
     Exit Sub
    
FimLeituraComErro:

     Close #DatFile

     If UBound(aArquivoRecep) >= 1 Then
          'Renomeia o arquivo para finaliza��o com erro
          Name strNomeArquivoREA As strNomeArquivoErr
     End If
     GoTo FimLeitura
     
ErroLeitura:
     Beep
     Screen.MousePointer = vbDefault
     
     'Tratamento para diret�rio n�o localizado
     If Err.Number = 76 Then
          MsgBox "N�o localizado o diret�rio com o arquivo de movimento Data Boa." & vbCrLf & vbCrLf & _
                    "Favor verificar diret�rio em par�metros.", vbCritical, sTituloTela
          GoTo FimLeitura
     End If
     If Err.Number = 53 Then
          MsgBox "N�o localizado o arquivo de movimento de data boa." & vbCrLf & vbCrLf & _
                    "Favor verificar e tentar novamente.", vbCritical, sTituloTela
          GoTo FimLeitura
     End If
     If Err.Number = 55 Then
          MsgBox "Arquivo com movimento de data boa em utiliza��o por outro usu�rio. Favor verificar !", vbCritical, sTituloTela
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
     sTituloTela = "Recep��o da Baixa de Cheques"
    
     If MsgBox("Inicializa processo de recep��o", vbQuestion + vbYesNo, sTituloTela) = vbNo Then
          Exit Sub
     End If
    
     If Not LerArquivosDiretorio(g_Parametros.DiretorioRecepcao & "\", "CHRBAI", sArquivoRecep) Then
          Exit Sub
     End If
     If Len(sArquivoRecep) = 0 Then
          Beep
          MsgBox "N�o existe arquivo(s) para recep��o referente � Baixa de Cheques.", vbInformation, sTituloTela
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
          'Muda extens�o do arquivo para (REA) qdo estiver sendo acessado para leitura
          strNomeArquivoREA = strNomeArquivo & "_READ"
          'Muda extens�o do arquivo para (OK) qdo finalizado o processo
          strNomeArquivoOK = strNomeArquivo & "_OK"
          'Muda extens�o do arquivo para (Erro) qdo ocorrer erro
          strNomeArquivoErr = strNomeArquivo & "_ERRO"
          
          bArquivoReading = False
          
          'Verifica se arq. confirma��o existe no diret�rio e est� sendo acessado por outro usu�rio
          If Dir(strNomeArquivo, vbDirectory) = "" Then
               
               'Verifica se arq. confirma��o est� sendo lido por outro usu�rio (.REA)
               If Dir(strNomeArquivoREA, vbDirectory) = "" Then

                    'Abre arquivo para verificar erro de inexistencia do diret�rio/Arquivo
                    Open strNomeArquivo For Input As #1
                    Close #1
                    Exit Sub
               Else
                    'For�a abertura do arquivo para executar tratamento do erro
                    'Caso consiga abrir o arquivo, � sinal de que houve queda de execu��o
                    'e o arquivo ficou em modo Reading...
                    Open strNomeArquivoREA For Input Access Read As #1
                    Close #1
                    bArquivoReading = True
               End If
          End If
          
          'Renomeia o arquivo para extens�o (.REA)
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
                  
              ' Ver se CGC de terceira � v�lido
              If CStr(Mid(Reg, 7, 14)) <> Format(g_Parametros.CNPJ_Terceira, String(14, "0")) Then
                   MsgBox "CNPJ da Terceira Inv�lido", vbOKOnly + vbCritical, sTituloTela
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
              
             
              'Calcula d�gito verificador do border�
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
          
          'Renomeia o arquivo para extens�o (.OK)
          Name strNomeArquivoREA As strNomeArquivoOK
          
     Next
    
     'Encerra progress bar
     Set Progress = Nothing
    
     'Fecha tela apresenta��o de arquivos
     Call LabelRecepcao(LabelAcao.Finaliza)
    
     Screen.MousePointer = vbDefault
     MsgBox "Foram Processados " & CStr(nBaixas + nJaBaixado) & " Cheques, Foram Baixados " & CStr(nBaixas), vbOKOnly + vbExclamation, sTituloTela
     Exit Sub
    
FimLeitura:
     'Fecha tela apresenta��o de arquivos
     Call LabelRecepcao(LabelAcao.Finaliza)

     Close #DatFile
    
     'Encerra progress bar
     Set Progress = Nothing
     
     Screen.MousePointer = vbDefault
     Exit Sub

FimLeituraComErro:

     Close #DatFile

     If UBound(aArquivoRecep) >= 1 Then
          'Renomeia o arquivo para finaliza��o com erro
          Name strNomeArquivoREA As strNomeArquivoErr
     End If
     GoTo FimLeitura
    
ErroLeitura:
     Beep
     Screen.MousePointer = vbDefault
     
     'Tratamento para diret�rio n�o localizado
     If Err.Number = 76 Then
          MsgBox "N�o localizado o diret�rio com o arquivo baixa de cheque." & vbCrLf & vbCrLf & _
                    "Favor verificar diret�rio em par�metros.", vbCritical, sTituloTela
          GoTo FimLeitura
     End If
     If Err.Number = 53 Then
          MsgBox "N�o localizado o arquivo de baixa de cheque." & vbCrLf & vbCrLf & _
                    "Favor verificar e tentar novamente.", vbCritical, sTituloTela
          GoTo FimLeitura
     End If
     If Err.Number = 55 Then
          MsgBox "Arquivo baixa de cheque em utiliza��o por outro usu�rio. Favor verificar !", vbCritical, sTituloTela
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
Dim strDigBordero        As String      'D�gito verificador do border�
Dim sTituloTela          As String
Dim sArquivoRecep        As String
Dim aArquivoRecep()
Dim iCount               As Integer, aIndex As Integer

On Error GoTo Erro_RecInstrucoes

     sTituloTela = "Recep��o da Tabela de instru��o VC"

     If MsgBox("Inicializa processo de recep��o da tabela de instru��es do VC", vbQuestion + vbYesNo, sTituloTela) = vbNo Then
          Set rsRetorno = Nothing
          Exit Sub
     End If
     
     If Not LerArquivosDiretorio(g_Parametros.DiretorioRecepcao & "\", "CHRDTV", sArquivoRecep) Then
          Set rsRetorno = Nothing
          Exit Sub
     End If
     If Len(sArquivoRecep) = 0 Then
          Beep
          MsgBox "N�o existe arquivo(s) para recep��o referente � Tabela de Instru��o", vbInformation, sTituloTela
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
          'Muda extens�o do arquivo para (REA) qdo estiver sendo acessado para leitura
          strNomeArquivoREA = strNomeArquivo & "_READ"
          'Muda extens�o do arquivo para (OK) qdo finalizado o processo
          strNomeArquivoOK = strNomeArquivo & "_OK"
          'Muda extens�o do arquivo para (Erro) qdo ocorrer erro
          strNomeArquivoErr = strNomeArquivo & "_ERRO"
     
          bArquivoReading = False
     
          'Verifica se arquivo existe no diret�rio e est� sendo acessado por outro usu�rio
          If Dir(strNomeArquivo, vbDirectory) = "" Then
     
               'Verifica se arquivo est� sendo lido por outro usu�rio (.REA)
               If Dir(strNomeArquivoREA, vbDirectory) = "" Then
                    
                    'Abre arquivo para verificar erro de inexistencia do diret�rio/Arquivo
                    Open strNomeArquivoREA For Input As #1
                    Close #1
                    Exit Sub
               Else
                    'For�a abertura do arquivo para executar tratamento do erro
                    'Caso consiga abrir o arquivo, � sinal de que houve queda de execu��o
                    'e o arquivo ficou em modo Reading...
                    Open strNomeArquivoREA For Input Access Read As #1
                    Close #1
                    bArquivoReading = True
               End If
          End If
     
          'Renomeia o arquivo para extens�o (.REA)
          If Not bArquivoReading Then
               Name strNomeArquivo As strNomeArquivoREA
          End If
          
          'Abre o arquivo de rejei��es com extens�o (.REA)
          Open strNomeArquivoREA For Random As #1 Len = Len(myrecord)
          
          'Obtem o total de bytes do arquivo de rejeitados
          lngTotalRegistros = FileLen(strNomeArquivoREA) / Len(myrecord)
          
         'Inicializa Progress Bar
          Progress.ValorMinimo = 1
          Progress.ValorMaximo = lngTotalRegistros
          Progress.DescricaoProcesso = "Recepcionando tabela de instru��es ..."
          Progress.InicializaProgressBar
          Progress.AtualizaBarra
          
          'Ler cada registro e atualizar tabelas
          For lngRegs = 1 To lngTotalRegistros
               'Ler linha � linha
               Get #1, lngRegs, myrecord
               
               'Calcula d�gito verificador do border�
               
               strDigBordero = RetornaDigitoModulo11Simplificado(Mid(myrecord.Num_Bordero, 13, 6))
               
               'Verifica se existe c�digo da carteira
               If Val(myrecord.Cod_Carteira) = 0 Then
                    Call GeraInconsInstrucao("N�o localizado c�digo de carteira", myrecord, recepcao, strDigBordero)
               Else
                    'Localiza c�digo da carteira na tabela de Carteira
                    Set rsRetorno = g_cMainConnection.Execute(Proc_Selecionar.GetCarteira(myrecord.Cod_Carteira), _
                                                            lngRetorno, adCmdText)
               
                    If rsRetorno.EOF Then
                         Call GeraInconsInstrucao("N�o localizado c�digo de carteira", myrecord, recepcao, strDigBordero)
                    Else
                    
                         'Insere registro na tabela de Altera��o de Data Dep�sito
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
                              Call GeraInconsInstrucao("N�o foi poss�vel atualizar a altera��o de data", myrecord, recepcao, strDigBordero)
                         End If
                    End If
               End If
               
               'Atualiza Progress Bar
               Progress.AtualValue = lngRegs
               Progress.AtualizaBarra
               
          Next
     
          'Fecha arquivo de rejei��es
          Close #1

          'Renomeia o arquivo para extens�o (.OK)
          Name strNomeArquivoREA As strNomeArquivoOK
     
     Next

     'Verifica se houve ocorrencia na recep��o
     If UBound(recepcao) > 0 Then
          
          Set rsRetorno = Nothing
          'Obtem data e hora do servidor (Chave para a tabela INCONSISTENCIA)
          Set rsRetorno = g_cMainConnection.Execute(Proc_Selecionar.GetDataHoraServidor(), lngRetorno, adCmdText)
          
          If lngRetorno <> 0 Then
               Beep
               MsgBox "Problema na gera��o do relat�rio de inconsist�ncias!", vbCritical, sTituloTela
               GoTo Sair
          End If
          
          strDataChave = Format(rsRetorno!Data, "yyyymmdd")
          strHoraChave = rsRetorno!Hora
          
          For lngCount = 1 To UBound(recepcao)
               'Insere registro com dados de inconsist�ncia
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
     
'          'Imprime rela��o de ocorr�ncias do processo de recep��o
'          Call ImprimirInconsistencia("RelRecepcaoInstrucoes", "Ocorr�ncias na recep��o da tabela de instru��o do VC", Geral.DataProcessamento, strDataChave, strHoraChave)
          
          'Fecha tela apresenta��o de arquivos
          Call LabelRecepcao(LabelAcao.Finaliza)
          MsgBox Space(32) & "A T E N � � O" & vbCrLf & vbCrLf & _
                 "Finalizada a recep��o com ocorr�ncia(s).  Verifique !", vbCritical, sTituloTela
          
          
          'Remove registro com dados de inconsist�ncia
          Call g_cMainConnection.Execute(Proc_Excluir.RemoveInconsistencia(strDataChave, _
                                                                           strHoraChave), _
                                                                           lngRetorno, adCmdText)
     Else
          'Fecha tela apresenta��o de arquivos
          Call LabelRecepcao(LabelAcao.Finaliza)
          
          MsgBox "Finalizada a recep��o da tabela de instru��es", vbInformation, sTituloTela
     End If
     
     
Sair:
     
     'Fecha tela apresenta��o de arquivos
     Call LabelRecepcao(LabelAcao.Finaliza)

     Set Progress = Nothing
     
     Set rsRetorno = Nothing
     
     Screen.MousePointer = vbDefault
     Exit Sub
     
SairSemFinalizar:
     'Fecha arquivo de rejei��es
     Close #1
     
     'Renomeia o arquivo para extens�o inicial ao processo
     Name strNomeArquivoREA As strNomeArquivo
     GoTo Sair
     
Erro_RecInstrucoes:
     
     Beep
     Screen.MousePointer = vbDefault
     
     'Tratamento para diret�rio n�o localizado
     If Err.Number = 76 Then
          MsgBox "N�o localizado o diret�rio com o arquivo de rejei��o, " & _
                    "verifique diret�rio no m�dulo par�metros.", vbCritical, sTituloTela
          GoTo Sair
     End If
     If Err.Number = 53 Then
          MsgBox "N�o localizado o arquivo de rejei��o, " & _
                 "verifique e tentar novamente.", vbCritical, sTituloTela
          GoTo Sair
     End If
     If Err.Number = 55 Then
          MsgBox "Arquivo de rejei��o em utiliza��o por outro usu�rio. Favor verificar!", vbCritical, sTituloTela
          GoTo Sair
     End If
     
    'Fecha arquivo de rejei��es
     Close #1

     If UBound(aArquivoRecep) >= 1 Then
         'Renomeia o arquivo para finaliza��o com erro
          Name strNomeArquivoREA As strNomeArquivoErr
     End If
     
     MsgBox Err.Description, vbCritical, sTituloTela
     GoTo Sair

End Sub
Private Sub GeraInconsInstrucao(strDescrMsg As String, myrecord As RecordInstrucoes, recepcao() As recepcao.tpRecepcao, ByVal strDigBordero As String)

Dim lngCount As Long
     
     ReDim Preserve recepcao(UBound(recepcao) + 1)
     lngCount = UBound(recepcao)

     'Acumula em type os border�s com ocorr�ncia
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
     
     sTituloTela = "Recep��o de Cheques Rejeitados"
     
     If MsgBox("Inicializa processo da recep��o de Cheques Rejeitados", vbQuestion + vbYesNo, sTituloTela) = vbNo Then
          Set rs = Nothing
          Exit Sub
     End If
     
     If Not LerArquivosDiretorio(g_Parametros.DiretorioRecepcao & "\", "REJEIC", sArquivoRecep) Then
          Set rs = Nothing
          Exit Sub
     End If
     
     If Len(sArquivoRecep) = 0 Then
          Beep
          MsgBox "N�o existe arquivo(s) para recep��o referente � Cheques Rejeitados", vbInformation, sTituloTela
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
          
         'Muda extens�o do arquivo para (REA) qdo estiver sendo acessado para leitura
          strNomeArquivoREA = strNomeArquivo & "_READ"
          
         'Muda extens�o do arquivo para (OK) qdo finalizado o processo
          strNomeArquivoOK = strNomeArquivo & "_OK"
          
         'Muda extens�o do arquivo para (Erro) qdo ocorrer erro
          strNomeArquivoErr = strNomeArquivo & "_ERRO"
     
          bArquivoReading = False

         'arq. rejei��o existe no diret�rio e est� sendo acessado por outro usu�rio ?
          If Dir(strNomeArquivo, vbDirectory) = "" Then
              'arq. rejei��o est� sendo lido por outro usu�rio ? (.REA)
               If Dir(strNomeArquivoREA, vbDirectory) = "" Then
                   'Abre arquivo para verificar erro de inexistencia do diret�rio/Arquivo
                    Open strNomeArquivo For Input As #1
                    Close #1
                    Exit Sub
               Else
                    Open strNomeArquivoREA For Input Access Read As #1
                    Close #1
                    bArquivoReading = True
               End If
          End If
          
         'Renomeia o arquivo para extens�o (.REA)
          If Not bArquivoReading Then
               Name strNomeArquivo As strNomeArquivoREA
          End If
          
         'Abre o arquivo de rejei��es com extens�o (.REA)
          Open strNomeArquivoREA For Random As #1 Len = Len(myrecord)
          
         'Obtem o total de bytes do arquivo de rejeitados
          lngTotalRegistros = FileLen(strNomeArquivoREA) / Len(myrecord)
          
         'Inicializa Progress Bar
          Progress.ValorMinimo = 0
          Progress.ValorMaximo = lngTotalRegistros
          Progress.DescricaoProcesso = "Recepcionando rejei��o de cheque ..."
          Progress.InicializaProgressBar
          Progress.AtualizaBarra
          
         'Ler cada registro e atualizar tabelas
          For lngRegs = 1 To lngTotalRegistros
             'Ler linha � linha
              Get #1, lngRegs, myrecord
              
             'Verifica se Arquivo � v�lido
              If Not IsNumeric(myrecord.Num_Bordero) Or myrecord.Rot_Mensag <> "REJEIC" Then
                 Beep
                 Err.Raise 910, App.Title, "Falha no m�dulo de Recep��o do Arquivo de Cheques Rejeitados, Arquivo com formato desconhecido"
              End If
              
             'Calcula Digito do bordero
              Bordero = myrecord.Num_Bordero & RetornaDigitoModulo11Simplificado(Mid(myrecord.Num_Bordero, 13, 6))
              
             'Procura Bordero na Base
              sstr = "Select DataProcessamento, Idbordero, Status From Bordero Where Num_bordero  = '" & Format(Bordero, String(19, "0")) & "'" & " And Status IN('T','C')"
              Set rs = g_cMainConnection.Execute(sstr, lngRetorno, adCmdText)
                       
             'Se N�o Encontrou Bordero
              If rs.EOF Then
                 Call GeraInconsRejeitados("N�mero de Border� n�o Encontrado", recepcao, Bordero)
              ElseIf rs.RecordCount > 1 Then
                'Se Encontrado mais de um Bordero
                 Call GeraInconsRejeitados("Encontrado mais de Um Border� com mesmo n�mero", recepcao, myrecord.Num_Bordero & strDigBordero)
              ElseIf rs("Status").Value = "E" Then
                Call GeraInconsRejeitados("Border� confirmado", recepcao, Bordero)
              Else
                'Bordero Encontrado, joga dados no type p/ posterior inser�ao na tabela de rejeicao
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
                         Err.Raise 910, App.Title, "Falha no m�dulo de Recep��o do Arquivo de Cheques Rejeitados, Atualiza��o de Status do Border�."
                     End If
                 End If
                 
                '*********************************************
                'Inserir na Tabela Rejei��o remessa os erros *
                '*********************************************
                 If myrecord.Rot_Original = "CHINBO" Then
                 
                    Call CodigoErro(TabRejeicao, Bordero)
                 
                'Verifica se o rotulo � de Cheque e trata
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
                        
                        'Joga dados no type p/ posterior inser�ao na tabela de rejeicao
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
                           'se Houve falha na atualiza�ao de status do cheque
                            Call GeraInconsRejeitados("N�o foi poss�vel atualizar o Cheque", recepcao, Bordero)
                         Else
                            'Se encontrado Ok
                            Call CodigoErro(TabRejeicao, Bordero)
                         End If
                      Else
                        'Se n�o encontrado cheque na Base
                         Call GeraInconsRejeitados("Cheque n�o Localizado na Base", recepcao, Bordero)
                      End If
                'Verifica se o rotulo � de Data e trata
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
                           'N�o localizada data na base
                            Call GeraInconsRejeitados("N�o foi poss�vel atualizar a Data Dep�sito", recepcao, Bordero)
                         Else
                           'Se encontrado Ok
                            Call CodigoErro(TabRejeicao, Bordero)
                         End If
                      Else
                         Call GeraInconsRejeitados("Data Dep�sito n�o Localizada na Base", recepcao, Bordero)
                      End If
                      
                 End If
                 
              End If
                   
             'Atualiza Progress Bar
             Progress.AtualValue = lngRegs
             Progress.AtualizaBarra
                  
          Next
          
         'Fecha arquivo de rejei��es
          Close #1
     
          'Renomeia o arquivo para extens�o (.OK)
          Name strNomeArquivoREA As strNomeArquivoOK
     Next
     
    'Verifica se houve ocorrencia na recep��o
     If UBound(recepcao) > 0 Then
          
          Set rs = Nothing
          
         'Obtem data e hora do servidor (Chave para a tabela INCONSISTENCIA)
          Set rs = g_cMainConnection.Execute(Proc_Selecionar.GetDataHoraServidor(), lngRetorno, adCmdText)
          
          If lngRetorno <> 0 Then
               Beep
               Err.Raise 910, App.Title, "Falha no m�dulo de Recep��o do Arquivo de Cheques Rejeitados,(Obten��o de Hora do Servidor)"
          End If
          
          strDataChave = Format(rs!Data, "yyyymmdd")
          strHoraChave = rs!Hora
          
          For lngCount = 1 To UBound(recepcao)
              'Insere registro com dados de inconsist�ncia
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
     
         'Imprime rela��o de ocorr�ncias do processo de recep��o
          Call ImprimirInconsistencia("RelInconsistenciaRecepcao", "Ocorr�ncias da recep��o de rejei��es", Geral.DataProcessamento, strDataChave, strHoraChave)
          
         'Remove registro com dados de inconsist�ncia
          Call g_cMainConnection.Execute(Proc_Excluir.RemoveInconsistencia(strDataChave, strHoraChave), lngRetorno, adCmdText)
     
     End If
     
Fim:
     'Fecha tela apresenta��o de arquivos
     Call LabelRecepcao(LabelAcao.Finaliza)

     Set Progress = Nothing
     
     Set rs = Nothing
       
     Screen.MousePointer = vbDefault
     Exit Sub
     
Erro:
     Beep
     Screen.MousePointer = vbDefault
     
     'Tratamento para diret�rio n�o localizado
     If Err.Number = 76 Then
        Call TratamentoErro("N�o localizado o diret�rio com o arquivo de rejei��o, " & _
                            "verifique o diret�rio no m�dulo par�metros.", Err, False, False)
        Resume Fim
     End If
     If Err.Number = 53 Then
        Call TratamentoErro("N�o localizado o arquivo de rejei��o, " & _
                            "verifique e tente novamente.", Err, False, False)
        Resume Fim
     End If
     If Err.Number = 55 Then
        Call TratamentoErro("Arquivo de rejei��o em utiliza��o por outro usu�rio. Favor verificar!", Err, False, False)
        Resume Fim
     End If
     
    'Fecha arquivo de rejei��es
     Close #1
     
     If UBound(aArquivoRecep) > 1 Then
         'Renomeia o arquivo para finaliza��o com erro
          Name strNomeArquivoREA As strNomeArquivoErr
     End If
     
     Call TratamentoErro("Falha no m�dulo de Recep��o de Cheques", Err, False, False)
     Resume Fim

End Sub
Function CodigoErro(Registro As tpRejeicao, pBordero As String) As Boolean
Dim iCodErro, Retorno     As Integer
Dim intCodErro            As String
Dim rs                    As New ADODB.Recordset

For iCodErro = 1 To CInt(myrecord.Qtd_Erros)
    'Obtem separadamente o c�d. de erro existente no registro do arq. Rejei��es
     intCodErro = Mid(myrecord.Cod_Erros, (iCodErro * 3) - 2, 3)
          
     Set rs = g_cMainConnection.Execute(Proc_Selecionar.GetExisteRejeicaoRemessa(Registro.DataProcessamento, Registro.IdBordero, Registro.DataDeposito, Registro.IdCheque, intCodErro), Retorno, adCmdText)

     If Retorno <> 0 Then
        Beep
        MsgBox "Processo cancelado." & vbCrLf & vbCrLf & "N�o foi poss�vel ler informa��o de Rejei��o do border� (" & pBordero & ")", vbCritical, App.Title
        Exit Function
     End If
    
     If rs(0).Value = 0 Then

         Call g_cMainConnection.Execute(Proc_Inserir.InsereRejeicaoRemessa(Registro.DataProcessamento, Registro.IdBordero, Registro.DataDeposito, Registro.IdCheque, intCodErro, Registro.Rotulo), Retorno, adCmdText)

         If Retorno <= 0 Then
            Beep
            MsgBox "Processo cancelado." & vbCrLf & vbCrLf & "N�o foi poss�vel inserir informa��es de rejei��o no arquivo.", vbCritical, App.Title
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
          
          If iAcao = LabelAcao.Apresenta Then      'Apresenta Label de arquivos em recep��o
               .lblRecepcao.Caption = ""
               .lblRecepcao.Top = (Screen.Height - .lblRecepcao.Height) / 2
               .lblRecepcao.Left = (Screen.Width - .lblRecepcao.Width) / 2
               .lblRecepcao.Visible = True
     
               .lblRecepcaoArquivo.Caption = ""
               .lblRecepcaoArquivo.Top = (Screen.Height - .lblRecepcaoArquivo.Height + 350) / 2
               .lblRecepcaoArquivo.Left = (Screen.Width - .lblRecepcaoArquivo.Width) / 2
               .lblRecepcaoArquivo.Visible = True
          
          ElseIf iAcao = LabelAcao.Finaliza Then  'Fecha Label com apresenta��o de arquivos em recep��o
               .lblRecepcao.Visible = False
               .lblRecepcaoArquivo.Visible = False
          
          ElseIf iAcao = LabelAcao.Atualiza Then  'Preenche Label de arquivos em recep��o com as informa��es
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
        
     sTituloTela = "Recep��o Aviso de Diferen�a"
     nAvisos = 0
     vPos = 1
    
     If MsgBox("Inicializa processo de recep��o", vbQuestion + vbYesNo, sTituloTela) = vbNo Then
          Exit Sub
     End If
    
     If Not LerArquivosDiretorio(g_Parametros.DiretorioRecepcao & "\", "CHADIF", sArquivoRecep) Then
          Exit Sub
     End If
     If Len(sArquivoRecep) = 0 Then
          Beep
          MsgBox "N�o existe arquivo(s) para recep��o do Aviso de Diferen�a.", vbInformation, sTituloTela
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
          'Muda extens�o do arquivo para (REA) qdo estiver sendo acessado para leitura
          strNomeArquivoREA = strNomeArquivo & "_READ"
          'Muda extens�o do arquivo para (OK) qdo finalizado o processo
          strNomeArquivoOK = strNomeArquivo & "_OK"
          'Muda extens�o do arquivo para (Erro) qdo ocorrer erro
          strNomeArquivoErr = strNomeArquivo & "_ERRO"
          
          bArquivoReading = False
          
          'Verifica se arq. existe no diret�rio e est� sendo acessado por outro usu�rio
          If Dir(strNomeArquivo, vbDirectory) = "" Then
               
               'Verifica se arq. confirma��o est� sendo lido por outro usu�rio (.REA)
               If Dir(strNomeArquivoREA, vbDirectory) = "" Then

                    'Abre arquivo para verificar erro de inexistencia do diret�rio/Arquivo
                    Open strNomeArquivo For Input As #1
                    Close #1
                    Exit Sub
               Else
                    'For�a abertura do arquivo para executar tratamento do erro
                    'Caso consiga abrir o arquivo, � sinal de que houve queda de execu��o
                    'e o arquivo ficou em modo Reading...
                    Open strNomeArquivoREA For Input Access Read As #1
                    Close #1
                    bArquivoReading = True
               End If
          End If
          
          'Renomeia o arquivo para extens�o (.REA)
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
                Progress.DescricaoProcesso = "Recepcionando Aviso de Diferen�a ..."
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
              
              ' Ver r�tulo do arquivo
              If Mid(Reg, 1, 6) <> "CHADIF" Then
                   MsgBox "R�tulo do Arquivo de Diferen�a Inv�lido.", vbOKOnly + vbCritical, sTituloTela
                   GoTo FimLeituraComErro
              End If
              
              ' Ver se CGC de terceira � v�lido
              If CStr(Mid(Reg, 7, 14)) <> FormataString(g_Parametros.CNPJ_Terceira, "0", 14, True) Then
                   MsgBox "CNPJ da Terceira Inv�lido", vbOKOnly + vbCritical, sTituloTela
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
                    
                      ' Gravar Registro do Aviso de Diferen�a
                      
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

          'Renomeia o arquivo para extens�o (.OK)
          Name strNomeArquivoREA As strNomeArquivoOK

     Next
     
     'Encerra progress bar
     Set Progress = Nothing
    
     'Fecha tela apresenta��o de arquivos
     Call LabelRecepcao(LabelAcao.Finaliza)
    
     Screen.MousePointer = vbDefault
     MsgBox "Foram Processados " & CStr(nAvisos) & " Avisos de Diferen�a.", vbOKOnly + vbExclamation, sTituloTela
     Exit Sub
    
FimLeitura:

     'Fecha tela apresenta��o de arquivos
     Call LabelRecepcao(LabelAcao.Finaliza)
    
     Close #DatFile
     'Encerra progress bar
     Set Progress = Nothing
     
     Screen.MousePointer = vbDefault
     Exit Sub
    
FimLeituraComErro:

     Close #DatFile

     If UBound(aArquivoRecep) >= 1 Then
          'Renomeia o arquivo para finaliza��o com erro
          Name strNomeArquivoREA As strNomeArquivoErr
     End If
     GoTo FimLeitura

ErroLeitura:
     Beep
     Screen.MousePointer = vbDefault
     
     'Tratamento para diret�rio n�o localizado
     If Err.Number = 76 Then
          MsgBox "N�o localizado o diret�rio com o arquivo de Aviso de Diferen�a." & vbCrLf & vbCrLf & _
                    "Favor verificar diret�rio em par�metros.", vbCritical, sTituloTela
          GoTo FimLeitura
     End If
     If Err.Number = 53 Then
          MsgBox "N�o localizado o arquivo de Aviso de Diferen�a." & vbCrLf & vbCrLf & _
                    "Favor verificar e tentar novamente.", vbCritical, sTituloTela
          GoTo FimLeitura
     End If
     If Err.Number = 55 Then
          MsgBox "Arquivo com Aviso de Diferen�a em utiliza��o por outro usu�rio. Favor verificar !", vbCritical, sTituloTela
          GoTo FimLeitura
     End If

    MsgBox "Erro na Leitura do Arquivo de Aviso de Diferen�a.", vbOKOnly + vbCritical, sTituloTela

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

     sTituloTela = "Fus�o Autom�tica"
     nCheques = 0
     
     strNomeArquivo = Trim(FusaoDialog.TxtFusao.Text)
     
     If Not FileExist(strNomeArquivo) Then
        Beep
        MsgBox "Arquivo de Fus�o N�o Encontrado", vbExclamation, sTituloTela
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
                
                Progress.DescricaoProcesso = "Processando a Fus�o Autom�tica"
                Progress.InicializaProgressBar
                Progress.AtualizaBarra
                lngRegs = 0
            End If
          
           ' Se arquivo foi lido ok
            If Len(Reg) < 150 Then
                MsgBox "Erro de Leitura", vbOKOnly + vbCritical, sTituloTela
                GoTo FimLeituraComErro
            End If
              
            ' Ver r�tulo do arquivo
            If Mid(Reg, 1, 2) <> "HD" Then
                MsgBox "R�tulo do Arquivo Inv�lido.", vbOKOnly + vbCritical, sTituloTela
                GoTo FimLeituraComErro
            End If
              
            ' Verifica C�digo da Terceira
            If CStr(Mid(Reg, 3, 4)) <> g_Parametros.Codigo_Terceira Then
               MsgBox "C�digo da Terceira Inv�lido", vbOKOnly + vbCritical, sTituloTela
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
                        MsgBox "Documento j� Processado" & vbCrLf & vbCrLf & "Cheque n� " & Mid(sCMC7, 12, 6) & "  Banco " & Mid(sCMC7, 1, 3), vbExclamation + vbOKOnly
                    Else
                        Call g_cMainConnection.Execute(Atu_Fusao.AtualizaFusao(sCMC7, Geral.DataProcessamento), lRetorno, adCmdText)
                        If lRetorno = 0 Then
                            Err.Raise 998, App.Title, "Erro ao Atualizar - Fus�o"
                            Exit Sub
                        Else
                            nCheques = nCheques + lRetorno
                        End If
                        
                    End If
            
                Else
                        
                    MsgBox "Documento n�o encontrado", vbExclamation + vbOKOnly, sTituloTela
        
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
     
     'Tratamento para diret�rio n�o localizado
     If Err.Number = 76 Then
          MsgBox "Diret�rio do Arquivo N�o Localizado." & vbCrLf & vbCrLf & _
                    "Favor verificar diret�rio em par�metros.", vbCritical, sTituloTela
          GoTo FimLeitura
     End If
     If Err.Number = 53 Then
          MsgBox "Arquivo N�o Localizado." & vbCrLf & vbCrLf & _
                    "Favor verificar e tentar novamente.", vbCritical, sTituloTela
          GoTo FimLeitura
     End If
     If Err.Number = 55 Then
          MsgBox "Arquivo em utiliza��o por outro usu�rio. Favor verificar!", vbCritical, sTituloTela
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
        
     sTituloTela = "Recep��o Regra do GP"
     nRegraGP = 0
     
    
     If MsgBox("Inicializa processo de recep��o", vbQuestion + vbYesNo, sTituloTela) = vbNo Then
          Exit Sub
     End If
    
     If Not LerArquivosDiretorio(g_Parametros.DiretorioRecepcao & "\", "CHREGP", sArquivoRecep) Then
          Exit Sub
     End If
     If Len(sArquivoRecep) = 0 Then
          Beep
          MsgBox "N�o existe arquivo(s) para recep��o da Regra do GP.", vbInformation, sTituloTela
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
          'Muda extens�o do arquivo para (REA) qdo estiver sendo acessado para leitura
          strNomeArquivoREA = strNomeArquivo & "_READ"
          'Muda extens�o do arquivo para (OK) qdo finalizado o processo
          strNomeArquivoOK = strNomeArquivo & "_OK"
          'Muda extens�o do arquivo para (Erro) qdo ocorrer erro
          strNomeArquivoErr = strNomeArquivo & "_ERRO"
          
          bArquivoReading = False
          
          'Verifica se arq. existe no diret�rio e est� sendo acessado por outro usu�rio
          If Dir(strNomeArquivo, vbDirectory) = "" Then
               
               'Verifica se arq. confirma��o est� sendo lido por outro usu�rio (.REA)
               If Dir(strNomeArquivoREA, vbDirectory) = "" Then

                    'Abre arquivo para verificar erro de inexistencia do diret�rio/Arquivo
                    Open strNomeArquivo For Input As #1
                    Close #1
                    Exit Sub
               Else
                    'For�a abertura do arquivo para executar tratamento do erro
                    'Caso consiga abrir o arquivo, � sinal de que houve queda de execu��o
                    'e o arquivo ficou em modo Reading...
                    Open strNomeArquivoREA For Input Access Read As #1
                    Close #1
                    bArquivoReading = True
               End If
          End If
          
          'Renomeia o arquivo para extens�o (.REA)
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
              
              ' Ver r�tulo do arquivo
              If Mid(Reg, 1, 6) <> "CHREGP" Then
                   MsgBox "R�tulo da Regra do GP Inv�lido.", vbOKOnly + vbCritical, sTituloTela
                   GoTo FimLeituraComErro
              End If
              
              ' Ver se CGC de terceira � v�lido
'              If CStr(Mid(Reg, 7, 14)) <> g_Parametros.CNPJ_Terceira Then
'                   MsgBox "CNPJ da Terceira Inv�lido", vbOKOnly + vbCritical, sTituloTela
'                   GoTo FimLeituraComErro
'              End If
              
              ' Ver se C�dido do Produto = 43265
              If CStr(Mid(Reg, 29, 5)) <> "43265" Then
                   MsgBox "C�dido do Produto Inv�lido", vbOKOnly + vbCritical, sTituloTela
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

          'Renomeia o arquivo para extens�o (.OK)
          Name strNomeArquivoREA As strNomeArquivoOK

     Next
     
     'Encerra progress bar
     Set Progress = Nothing
    
     'Fecha tela apresenta��o de arquivos
     Call LabelRecepcao(LabelAcao.Finaliza)
    
     Screen.MousePointer = vbDefault
     MsgBox "Foram Processados " & CStr(nRegraGP) & " Regra do GP.", vbOKOnly + vbExclamation, sTituloTela
     Exit Sub
    
FimLeitura:

     'Fecha tela apresenta��o de arquivos
     Call LabelRecepcao(LabelAcao.Finaliza)
    
     Close #DatFile
     'Encerra progress bar
     Set Progress = Nothing
     
     Screen.MousePointer = vbDefault
     Exit Sub
    
FimLeituraComErro:

     Close #DatFile

     If UBound(aArquivoRecep) >= 1 Then
          'Renomeia o arquivo para finaliza��o com erro
          Name strNomeArquivoREA As strNomeArquivoErr
     End If
     GoTo FimLeitura

ErroLeitura:
     Beep
     Screen.MousePointer = vbDefault
     
     'Tratamento para diret�rio n�o localizado
     If Err.Number = 76 Then
          MsgBox "N�o localizado o diret�rio com o arquivo Regra do GP." & vbCrLf & vbCrLf & _
                    "Favor verificar diret�rio em par�metros.", vbCritical, sTituloTela
          GoTo FimLeitura
     End If
     If Err.Number = 53 Then
          MsgBox "N�o localizado o arquivo de Aviso de Diferen�a." & vbCrLf & vbCrLf & _
                    "Favor verificar e tentar novamente.", vbCritical, sTituloTela
          GoTo FimLeitura
     End If
     If Err.Number = 55 Then
          MsgBox "Arquivo com Regra do GP em utiliza��o por outro usu�rio. Favor verificar !", vbCritical, sTituloTela
          GoTo FimLeitura
     End If

    MsgBox "Erro na Leitura do Arquivo de Regra do GP.", vbOKOnly + vbCritical, sTituloTela

    GoTo FimLeitura

End Sub
