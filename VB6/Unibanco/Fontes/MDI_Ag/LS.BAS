Attribute VB_Name = "LS_M"

Option Explicit


'---- VARI�VEIS USADAS PARA CTREE ----
Global CtreeVars As Long
Global lixo, lixao, llock As Integer
Global Current_Indice As Integer

'---- N�MEROS DOS ARQUIVOS/INDICES - CTREE ----
Global Const ArqLok = 0         'numero do arq de LOK

Global Const IndDat = 1         'numero do arq de Dados - Ind
Global Const IndDtLote = 4      'numero do Indice Data+Lote
Global Const IndBcAg = 5        'numero do Indice Banco+Agencia
Global Const IndBcAgCta = 6     'numero do Indice Banco+Agencia+Conta
Global Const IndBcAgCh = 7      'numero do Indice Banco+Agencia+Cheque
Global Const IndDtFLtImg = 8    'numero do Indice Data+Flag_CMC7+Lote+Img_frente

Global Const RecupDat = 2       'numero do arq de Dados - Recupera��o
Global Const RecupDtLote = 9    'numero do Indice Data+Lote

Global Const CbaDat = 3         'numero do arq de Dados - CADBAN
Global Const CbaAge = 10        'numero do Indice N�mero Agencia

'---------------------------------------------------------------
Type Indices                     'VB      CTREE
   Data_Movimento As String * 6  '1       0        data (AAMMDD)        chave
   F_CMC7 As String * 5          '7       6        flag CMC7            chave
   lote As String * 5            '12      11       lote                 chave
   Banco As String * 3           '17      16       banco                chave
   agencia As String * 4         '20      19       ag�ncia              chave
   C1 As String * 1              '24      23       c1
   Conta As String * 6           '25      24       conta                chave
   C2 As String * 1              '31      30       c2
   Comp As String * 3            '32      31       compensa��o
   cheque As String * 6          '35      34       cheque               chave
   C3 As String * 1              '41      40       c3
   Adicional As String * 5       '42      41       tipifica��o do cheque + raz�o C/C + DAC
   Agencia_Vinc As String * 5    '47      46       ag�ncia vinculada
   Conta_Vinc As String * 8      '52      51       conta vinculada
   Valor_Vinc As String * 12     '60      59       valor vinculado
   I_Frente As String * 12       '72      71       imagem frente        chave
   I_Verso As String * 12        '84      83       imagem verso
   Dias As String * 2            '96      95       n�mero de dias de vinculo
   usuario As String * 10        '98      97       nome do usu�rio que colocou este cheque para pend�ncia
End Type
    
Global reg_ind As Indices        'estrutura Ind
'tamanho desta estrutura = 107

'----------------------------------------------------------------
Type Recuperacao                    'VB      CTREE
    Data_Movimento As String * 6    '1       0        data (AAMMDD)        chave
    lote As String * 5              '7       6        lote em quest�o      chave
    Prod_Comp As String * 1         '12      11       status da produ��o e compress�o
    Corr_Rev As String * 2          '13      12       status da corre��o e Entrada de Dados
    Pendencia As String * 3         '15      14       status pend�ncia de : valor , corre��o , Entrada de Dados
    usuario As String * 10          '18      17       nome do usu�rio que efetuou a opera��o
    Data_Inicio As String * 8       '28      27       data inicial da opera��o
    Hora_Inicio As String * 8       '36      35       hora inicial da opera��o
    Data_Final As String * 8        '44      43       data final da opera��o
    Hora_Final As String * 8        '52      51       hora final da opera��o
    Imagem As String * 12           '60      59       n�mero da �ltima imagem
    Diretorio As String * 40        '72      71       diret�rio de trabalho
    Documentos As String * 3        '112    111       qtde de documentos desse lote        *** p/ff ***
    Nosso_Lote As String * 5        '115    114       numero do lote gerado pelo sistema   *** p/ff ***
    Docs_Rev As String * 3          '120    119       quantidade de documentos que ja foram revinculados   *** p/ff ***
    Diferenca_Lote As String * 13   '123    122       diferen�a do lote (tam = 13 em fun��o do sinal)
    Limite As String * 1            '136    135       flag do lote : superior / inferior
    Total_Lote As String * 12       '137    136       valor do lote
End Type

Global reg_recup As Recuperacao  'estrutura recupera��o
'tamanho desta estrutura = 148


'-----------------------------------------------------------------
Type Agencia_CADBAN
    Age_CBA As String * 5
    Age_aux As String * 5
End Type

Global reg_cba As Agencia_CADBAN
'tamanho desta estrutura = 10



'****************************************************************
'----       VARI�VEIS USADAS NO DECORRER DO PROGRAMA         ----
'****************************************************************

'*** string's fixas ***
Global bakkey As String * 100       'tamanho de uma linha do arquivo de retorno = 90 + 2 (crlf)
Global sebuf As String * 5600       '92 colunas * 60 linhas (bloco lido) = 5520
Global lote_arq As String * 5
Global Const branco = "  "          ' usado para montar arquivo de log


'*** string's vari�veis ***
Global servidor As String
Global linha_log, operacao As String      'usado para montar arquivo de log
Global user_acesso, base_imp As String
Global lote_anterior, buffer As String
Global data_mov, data_inverte As String
Global proxima_imagem, revinc_imagem As String
Global valor_capa_lote, novo_num_lote As String
Global reg_imagem, reg_frente, reg_verso As String
Global agencia_RV_aux, conta_RV_aux As String
Global agencia_RV, conta_RV, agencia_print As String
Global diretorio_imagem, diretorio_producao As String
Global frente_o, frente_d, verso_o, verso_d As String
Global nro_C1, nro_C2, nro_C3, nro_adicional As String
Global nome_frente, nome_verso, producao_valor_limite As String
Global nro_banco, nro_agencia, nro_conta, nro_cheque, nro_comp, nro_valor_vinc As String

'*** inteiros ***
Global ret, situacao As Integer
Global cancela_data, ret_msg As Integer
Global ja_fiz_capa_deste_lote As Integer
Global revinc_click, corrige_click As Integer
Global stop_revinculo, stop_correcao As Integer
Global fecha_janelas, cheques_feitos As Integer
Global origem_pendencia, ret_lock, erro As Integer
Global conf_cadban, par_setup, env_supervisor As Integer
Global pendente_cheque_rev, pendente_lote_rev As Integer
Global pendente_cheque_cor, pendente_lote_cor As Integer
Global problema_imagem, frente_verso, flag_pos As Integer
Global cheque_excluido_CMC7, cheque_excluido_rev As Integer
Global input_posicao, input_rotacao, input_scroll As Integer
Global input_col, input_lin, input_larg, input_alt As Integer

'*** double ***
Global total_capa_lote As Double

'*** single ***
Global input_zoom As Single



'********************************************
'*** Estruturas para pesquisa ***

'indice DATA+LOTE do arquivo IND.DAT (DUPLICIDADE)
Type a
    DtLote As String * 11
    Deslo As Long
End Type
Global t_lote As a


'indice BANCO+AGENCIA do arquivo IND.DAT (DUPLICIDADE)
Type b
   Banco As String * 3
   agencia As String * 4
   Deslo As Long
End Type
Global t_ba As b


'indice BANCO+AGENCIA+CONTA do arquivo IND.DAT (DUPLICIDADE)
Type C
   Banco As String * 3
   agencia As String * 4
   Conta As String * 6
   Deslo As Long
End Type
Global t_bac As C


'indice BANCO+AGENCIA+CHEQUE do arquivo IND.DAT (DUPLICIDADE)
Type d
   Banco As String * 3
   agencia As String * 4
   cheque As String * 6
   Deslo As Long
End Type
Global t_bach As d


'indice DATA+LOTE do arquivo RECUP.DAT (N�O DUPLICIDADE)
Type E
    DtLote As String * 11
End Type
Global recup_lote As E


'indice DATA+FLAG_CMC7+LOTE+IMAGEM_FRENTE do arquivo IND.DAT (N�O DUPLICIDADE)
Type f
   DtFgLoteImg As String * 28
End Type
Global t_flaglote As f


'indice AGENCIA do arquivo CBA.DAT (N�O DUPLICIDADE)
Type M
   Agenc_a As String * 5
End Type
Global t_age As M


'-----------------------------------------------------------------------------------------------------------------
'              ARQUIVOS TEXTO (.TXT) usados no sistema
'
' M�DULO PRODU��O:      - (05) RETORNOS.TXT (cont�m CMC-7 dos cheques captados)
'                       - (50) FLAGLOTE.TXT (diz se o arquivo Lote.Txt est� sendo usado)
'                       - (51) LOTE.TXT     (cont�m o n�mero do pr�ximo lote a ser digitalizado)
'                       - (73) BAKKEY.TXT   (controle interno do sistema para checar a rotina de prepara��o do CMC-7)
'                       - (06) TPYYMMDD.TXT (cont�m status do lotes e cheques - situa��o)
'***
' M�DULO RELAT�RIO:     - (25) I500I.TXT    (cont�m listagem dos cheques inferiores)
'                       - (30) I500S.TXT    (cont�m listagem dos cheques superiores)
'                       - (27) I500E.TXT    (cont�m listagem dos cheques excluidos)
'                       - (52) LIMITE.TXT   (cont�m valor do limite)
'***
' M�DULO CORRE��O:      - (15) P00?????.TXT (informa o lote com pend�ncia de CMC-7)
'***
' M�DULO ENT. DADOS:    - (16) P00?????.TXT (informa o lote com pend�ncia de Ent.Dados)
'***
' M�DULO COMPRESS�O:    - (60) COMPRESS.TXT (informa que compress�o est� sendo feita por algum servidor)
'***
' M�DULO LOG:           - (75) LOG???? (arquiv�o de log)
' M�DULO LOG:           - (76) LG?????? (arquivo de log de cada esta��o)
'***
' M�DULO PRINCIPAL:     - (09) CADBAN.TXT   (cont�m n�mero das ag�ncias do Bradesco)
'                       - (97) DIRPROD      (cont�m o PATH para o diret�rio de produ��o)
'                       - (96) DIRIMA       (cont�m o PATH para o diret�rio de imagens)
'***
' M�DULO CONSULTA:      - (99) LE_LOTE.TXT  (listagem que cont�m todos os lotes digitalizados e seus status)
'                       - (88) LE_CHEQUE.TXT(listagem que cont�m todos os cheques digitalizados e seus status)
'                       - (96) DIRIMA       (cont�m o PATH para o diret�rio de imagens)
'***
' M�DULO SENHA.BAS:     - (01) LS.INI       (arquivo .INI)
'***
' M�DULO SENHA_PENDENTE:- (70) PENDENTE.TXT (informa que pend�ncia de valor j� est� sendo feita por alguma esta��o)
'                         (71) PEND_COR.TXT (informa que pend�ncia de corre��o j� est� sendo feita por alguma esta��o)
'                         (72) CONSISTE.TXT (informa que consist�ncia de corre��o j� est� sendo feita por alguma esta��o)
' M�DULO SENHA_DEL��O:  - (74) DELECAO.TXT  (informa que arquivos do CTREE n�o foram apagados)
'
'
'-----------------------------------------------------------------------------------------------------------------

Function Get_Lock()
   
   ' TRUE - Obteve Lock
   ' FALSE- NAO Lock
   
   Dim tempo As Integer

   tempo = 30                 ' numero de tentativas para obter lock
   
   Do
      llock = LKISAM(2)       ' enabled CTREE
      If (llock = 0) Then
         Exit Do
      End If
      
      tempo = tempo - 1

   Loop While (tempo > 0)

   Get_Lock = False           ' verifica se obteve lock
   If tempo > 0 Then
      Get_Lock = True
   End If

End Function


Sub Grava_Log()

   Close #75     'fecha caso esteja aberto
   
   Dim dir_log, arquivo, log_operador, valor_caption, nro_maquina As String
   Dim tamanho, espacos, rep As Integer
   Dim str_formatada As String * 13

   'arquivo = "LG" + data_inverte
   
   ' verifica se existe arquivo com n�mero da m�quina.
   nro_maquina = Dir$("c:\windows\system\*.log")
   If (nro_maquina = "") Then
      nro_maquina = "99MICROA.LOG"
   End If
   
   arquivo = "LG" + Mid$(data_inverte, 3, 2) + Mid$(data_inverte, 5, 2) + Mid$(nro_maquina, 1, 2)

   dir_log = Dir(diretorio_imagem & arquivo)

   '*******************************************************************************
   '*** linha_log cont�m as seguintes informa��es: lote + banco + ag�ncia + comp +
   '*** cheque + conta + agencia_vinc + conta_vinc + valor_vinc + operador + data da
   '*** operacao + opera��o
   '*******************************************************************************
   log_operador = user_acesso       'usu�rio de acesso

   If (user_acesso = "") Then
      log_operador = String(10, " ")
   Else
      'completa campo USU�RIO com brancos no final.
      tamanho = Len(user_acesso)
      espacos = 10 - tamanho
      log_operador = user_acesso & String(espacos, " ")
   End If
   
   str_formatada = ""     'limpa string para guardar valor formatado
   rep = 1                'contador de caracteres a serem formatados
   valor_caption = Format$(Val(reg_ind.Valor_Vinc) / 100, "##########.00")

   Do
     If (Mid$(valor_caption, rep, 1) = "") Then   'verifica t�rmino da string
        Exit Do
     End If
     rep = rep + 1
   Loop While (rep < 14)   'tamanho m�ximo da string a ser formatada
   rep = rep - 1
            
   'formata � direita
   Mid$(str_formatada, 13 - rep + 1, rep) = Mid$(valor_caption, 1, rep)
   valor_caption = str_formatada    'atualiza valor_caption com dado formatado a direita

   
   linha_log = reg_recup.lote + branco + reg_ind.Banco + branco + reg_ind.agencia + branco + reg_ind.Comp + branco + reg_ind.cheque + branco + reg_ind.Conta + branco + reg_ind.Agencia_Vinc + branco + reg_ind.Conta_Vinc + branco + valor_caption + branco + log_operador + branco + Format$(Now, "DD/MM/YY") + branco + operacao
   
   'caso arquivo de LOG n�o exista, cria novo.
   If dir_log = "" Then
      Open (diretorio_imagem & arquivo) For Output As #75
      Print #75, linha_log
   Else
      Open (diretorio_imagem & arquivo) For Append As #75
      Print #75, linha_log
   End If

   Close #75

End Sub

Sub Main()
   
   '*** Inicializa vari�veis do CTREE ***
   CtreeVars = InitCtreeGV()
   
   Principal.Show

End Sub

Function Trata_Msg_Erro(nro_msg As Integer, ident_msg As Integer)

   Select Case nro_msg
      Case 1
         MsgBox "N�o foi poss�vel carregar a imagem , Verifique se as imagens est�o gravadas no diret�rio de origem! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 2
         MsgBox "Base de Dados n�o Encontrada! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 3
         MsgBox "Digite todos os campos! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 4
         MsgBox "Usu�rio n�o cadastrado! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 5
         MsgBox "Senha de Acesso incorreta! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 6
         MsgBox "N�o foi poss�vel ler o primeiro Lote do Banco de Dados! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 7
         MsgBox "Atualiza��o no arquivo de lotes negada , Reinicie Processo! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 8
         MsgBox "Data Inv�lida! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 9
         MsgBox "Registro n�o encontrado! " + "(" + Str$(ident_msg) + ")", 48 + 0, "Aten��o"
      Case 10
         MsgBox "Digite o dia e m�s (DD/MM) da Data do Movimento! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 11
         MsgBox "Lote n�o encontrado no Banco de Dados Imagens! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 12
         MsgBox "O campo est� incompleto. Digite todos os caracteres! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 13
         MsgBox "Arquivo de identifica��o n�o encontrado. Contate o Supervisor! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 14
         MsgBox "N�o existe nenhum Lote a ser tratado! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 15
         MsgBox "Contador de Cheques n�o atualizado. Continue Processo! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 16
         MsgBox "Atualiza��o do Arquivo de Cheques negada , Reinicie Processo! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 17
         MsgBox "Reinicie o M�dulo para confirma��o de t�rmino! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 18
         MsgBox "Digite o campo " + Str$(ident_msg) + " do CMC-7! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 19
         MsgBox "Reinicie a Pend�ncia de Entrada de Dados! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 20
         MsgBox "Anote este n�mero e informe ao Supervisor e/ou Suporte Datacheck: " + reg_ind.I_Frente + " -Comp - (" + Str$(ident_msg) + ")", MB_OK, "Aten��o"
      Case 21
         MsgBox "Impress�o de cheques negada , Reinicie Processo! " + "(" + Str$(ident_msg) + ")", MB_OK, "Aten��o"
      Case 22
         MsgBox "Campo " + Str$(ident_msg) + " n�o confere! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 23
         MsgBox "Digite o Campo Valor! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 24
         MsgBox "Digite o Campo Ag�ncia! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 25
         MsgBox "N�o existe nenhum Lote pertencente a esta Data! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 26
         MsgBox "T�rmino da Gera��o do Arquivo Texto! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 27
         MsgBox "Atualiza��o do novo valor do cheque negada , Reinicie Processo! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 28
         MsgBox "N�o existe nenhum lote gravado! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 29
         MsgBox "N�o existem mais cheques pertencentes a este Lote! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 30
         MsgBox "Reinicie CORRE��O! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 31
         MsgBox "Problema interno , Reinicie o procedimento! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 32
         MsgBox "Digite o N�mero do Lote! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 33
         MsgBox "Cheque n�o encontrado para Atualiza��o do Valor. Contate o Suporte da Datacheck! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 34
         MsgBox "Lote n�o encontrado para Atualiza��o do Valor. Contate o Suporte da Datacheck! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 35
         MsgBox "N�o existe Lote Gravado nesta Data! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 36
         MsgBox "Leitura do valor do Lote para impress�o negada , Reinicie Processo! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 37
         MsgBox "Todos os registros atuais ser�o deletados! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 38
         MsgBox "Exclus�o da Tabela negada, Reinicie Processo! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 39
         MsgBox "Arquivo CADBAN.TXT n�o encontrado! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 40
         MsgBox "Atualiza��o da Diferen�a do valor do cheque negada, Reinicie Processo! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 41
         MsgBox "Gera��o da Nova Tabela negada, Reinicie Processo! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 42
         MsgBox "� necess�rio fazer a Manuten��o do CADBAN antes da produ��o! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 43
         MsgBox "Usu�rio j� cadastrado! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 44
         MsgBox "Disco inv�lido! " + "(" + Str$(ident_msg) + ")", MB_ICONEXCLAMATION
      Case 45
         MsgBox "Selecione a Imagem a ser Impressa! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 46
         MsgBox "Selecione item do scroll para visualizar Imagem! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 47
         MsgBox "T�rmino do processo! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 48
         MsgBox "Lote n�o encontrado! " + "(" + Str$(ident_msg) + ")", 48 + 0, "Aten��o"
      Case 49
         MsgBox "Problema ao tentar ler cheques deste Lote! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 50
         MsgBox "Cheque n�o encontrado! " + "(" + Str$(ident_msg) + ")", 48 + 0, "Aten��o"
      Case 51
         MsgBox "Problema ao carregar os registros na tela! " + "(" + Str$(ident_msg) + ")", MB_OK, "Aten��o"
      Case 52
         MsgBox "Reinicie Entrada de Dados! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 53
         MsgBox "Base de Dados n�o Encontrada para fazer batimento das ag�ncias! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 54
         MsgBox "Arquivo de Lote nao encontrado , Contate o Supervisor! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 55
         MsgBox "Digite o n�mero de Cheques pertencentes ao Lote! " + "(" + Str$(ident_msg) + ")", 48, " Aten��o"
      Case 56
         MsgBox "Tratamento de abertura do arquivo (TXT) de lote negada, Reinicie Processo! " + "(" + Str$(ident_msg) + ")", MB_OK, "Aten��o"
      Case 57
         MsgBox "Verificar Timer-Tic da LS-500! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 58
         MsgBox "Verifique se todas as LS-500 est�o ligadas! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 59
         MsgBox "Grava��o de CMC-7 negada , Reinicie Processo! " + "(" + Str$(ident_msg) + ")", MB_OK, "Aten��o"
      Case 60
         MsgBox "Verifique se as LS-500 cont�m cheques! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 61
         MsgBox "N�mero de cheques n�o confere! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 62
         MsgBox "Problema na grava��o de cheques, Refa�a LOTE! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 63
         MsgBox "Problema na Grava��o do N�mero do Lote! ...Refa�a LOTE! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 64
         MsgBox "Digite apenas caracteres num�ricos! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 65
         MsgBox "Este Documento n�o � uma capa de lote! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 66
         MsgBox "Ja foi informada a Capa deste lote! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 67
         MsgBox "CMC-7 n�o confere! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 68
         MsgBox "Digite os Campos Anteriores! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 69
         MsgBox "Ag�ncia inexistente! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 70
         MsgBox "Digite o campo Conta! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 71
         MsgBox "Conta inexistente! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 72
         MsgBox "T�rmino do Processo de Corre��o do CMC-7! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 73
         MsgBox "Pr�ximo Lote n�o lido, Reinicie o Processo! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 74
         MsgBox "T�rmino do Lote, pois a interrup��o do Processo de Corre��o do CMC-7 foi solicitada pelo operador! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 75
         MsgBox "C�pia j� solicitada por outra esta��o! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 76
         MsgBox "Digite a Data do Movimento! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 77
         MsgBox "Digite a Data do Movimento para cheques inferiores! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 78
         MsgBox "N�o encontrado valor limite. Para grav�-lo, v� para o menu Supervisor, entre na Op��o <Novo Valor Limite>! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 79
         MsgBox "Aten��o !!! N�mero de Lote j� existente !!! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 80
         MsgBox "N�o existe nenhum Lote com Pend�ncia de Valor! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 81
         MsgBox "N�o existem mais Lotes com Pend�ncia de Valor! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 82
         MsgBox "Reinicie Processo! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 83
         MsgBox "Lote Batido! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 84
         MsgBox "Digite a senha de acesso! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 85
         MsgBox "Pendencia de Valor j� solicitada por outra esta��o! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 86
         MsgBox "CMC-7 n�o confere! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 87
         MsgBox "T�rmino do Processo de Pend�ncia do CMC-7! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 88
         MsgBox "T�rmino do Processo de Pend�ncia de Entrada de Dados! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 89
         MsgBox "Agencia n�o existente no CADBAN! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 90
         MsgBox "Ag�ncia inexistente! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 91
         MsgBox "Digite o campo Conta! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 92
         MsgBox "Problema ao ler primeiro registro! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 93
         MsgBox "N�o encontrado arquivo de LOG. Confirme a Data! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 94
         MsgBox "Arquivo de Estat�stica n�o encontrado. Contate o Supervisor! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 95
         MsgBox "Digite o Valor M�nimo para Impress�o dos Cheques! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 96
         MsgBox "Pr�ximo Lote n�o lido, Reinicie o Processo! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 97
         MsgBox "Pend�ncia de Corre��o de CMC-7 j� solicitada por outra esta��o! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 98
         MsgBox "Consist�ncia de Corre��o de CMC-7 j� solicitada por outra esta��o! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 99
         MsgBox "N�o foi poss�vel atualizar o registro(CMC76) desta imagem! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 100
         MsgBox "N�o foi poss�vel carregar esta imagem na tela! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 101
         MsgBox "N�o foi poss�vel imprimir a imagem " + reg_ind.I_Frente + " ! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 102
         MsgBox "N�o existe arquivo de Log pertencente a esta data! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 103
         MsgBox "Esta Dele��o s� poder� ser feita do Servidor PRINCIPAL! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"
      Case 104
         MsgBox "Esta Dele��o s� poder� ser feita do Servidor SECUND�RIO! " + "(" + Str$(ident_msg) + ")", 48, "Aten��o"

   
   End Select

End Function


