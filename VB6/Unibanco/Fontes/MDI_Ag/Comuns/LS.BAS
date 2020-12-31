Attribute VB_Name = "LS_M"

Option Explicit


'---- VARIÁVEIS USADAS PARA CTREE ----
Global CtreeVars As Long
Global lixo, lixao, llock As Integer
Global Current_Indice As Integer

'---- NÚMEROS DOS ARQUIVOS/INDICES - CTREE ----
Global Const ArqLok = 0         'numero do arq de LOK

Global Const IndDat = 1         'numero do arq de Dados - Ind
Global Const IndDtLote = 4      'numero do Indice Data+Lote
Global Const IndBcAg = 5        'numero do Indice Banco+Agencia
Global Const IndBcAgCta = 6     'numero do Indice Banco+Agencia+Conta
Global Const IndBcAgCh = 7      'numero do Indice Banco+Agencia+Cheque
Global Const IndDtFLtImg = 8    'numero do Indice Data+Flag_CMC7+Lote+Img_frente

Global Const RecupDat = 2       'numero do arq de Dados - Recuperação
Global Const RecupDtLote = 9    'numero do Indice Data+Lote

Global Const CbaDat = 3         'numero do arq de Dados - CADBAN
Global Const CbaAge = 10        'numero do Indice Número Agencia

'---------------------------------------------------------------
Type Indices                     'VB      CTREE
   Data_Movimento As String * 6  '1       0        data (AAMMDD)        chave
   F_CMC7 As String * 5          '7       6        flag CMC7            chave
   lote As String * 5            '12      11       lote                 chave
   Banco As String * 3           '17      16       banco                chave
   agencia As String * 4         '20      19       agência              chave
   C1 As String * 1              '24      23       c1
   Conta As String * 6           '25      24       conta                chave
   C2 As String * 1              '31      30       c2
   Comp As String * 3            '32      31       compensação
   cheque As String * 6          '35      34       cheque               chave
   C3 As String * 1              '41      40       c3
   Adicional As String * 5       '42      41       tipificação do cheque + razão C/C + DAC
   Agencia_Vinc As String * 5    '47      46       agência vinculada
   Conta_Vinc As String * 8      '52      51       conta vinculada
   Valor_Vinc As String * 12     '60      59       valor vinculado
   I_Frente As String * 12       '72      71       imagem frente        chave
   I_Verso As String * 12        '84      83       imagem verso
   Dias As String * 2            '96      95       número de dias de vinculo
   usuario As String * 10        '98      97       nome do usuário que colocou este cheque para pendência
End Type
    
Global reg_ind As Indices        'estrutura Ind
'tamanho desta estrutura = 107

'----------------------------------------------------------------
Type Recuperacao                    'VB      CTREE
    Data_Movimento As String * 6    '1       0        data (AAMMDD)        chave
    lote As String * 5              '7       6        lote em questão      chave
    Prod_Comp As String * 1         '12      11       status da produção e compressão
    Corr_Rev As String * 2          '13      12       status da correção e Entrada de Dados
    Pendencia As String * 3         '15      14       status pendência de : valor , correção , Entrada de Dados
    usuario As String * 10          '18      17       nome do usuário que efetuou a operação
    Data_Inicio As String * 8       '28      27       data inicial da operação
    Hora_Inicio As String * 8       '36      35       hora inicial da operação
    Data_Final As String * 8        '44      43       data final da operação
    Hora_Final As String * 8        '52      51       hora final da operação
    Imagem As String * 12           '60      59       número da última imagem
    Diretorio As String * 40        '72      71       diretório de trabalho
    Documentos As String * 3        '112    111       qtde de documentos desse lote        *** p/ff ***
    Nosso_Lote As String * 5        '115    114       numero do lote gerado pelo sistema   *** p/ff ***
    Docs_Rev As String * 3          '120    119       quantidade de documentos que ja foram revinculados   *** p/ff ***
    Diferenca_Lote As String * 13   '123    122       diferença do lote (tam = 13 em função do sinal)
    Limite As String * 1            '136    135       flag do lote : superior / inferior
    Total_Lote As String * 12       '137    136       valor do lote
End Type

Global reg_recup As Recuperacao  'estrutura recuperação
'tamanho desta estrutura = 148


'-----------------------------------------------------------------
Type Agencia_CADBAN
    Age_CBA As String * 5
    Age_aux As String * 5
End Type

Global reg_cba As Agencia_CADBAN
'tamanho desta estrutura = 10



'****************************************************************
'----       VARIÁVEIS USADAS NO DECORRER DO PROGRAMA         ----
'****************************************************************

'*** string's fixas ***
Global bakkey As String * 100       'tamanho de uma linha do arquivo de retorno = 90 + 2 (crlf)
Global sebuf As String * 5600       '92 colunas * 60 linhas (bloco lido) = 5520
Global lote_arq As String * 5
Global Const branco = "  "          ' usado para montar arquivo de log


'*** string's variáveis ***
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


'indice DATA+LOTE do arquivo RECUP.DAT (NÃO DUPLICIDADE)
Type E
    DtLote As String * 11
End Type
Global recup_lote As E


'indice DATA+FLAG_CMC7+LOTE+IMAGEM_FRENTE do arquivo IND.DAT (NÃO DUPLICIDADE)
Type f
   DtFgLoteImg As String * 28
End Type
Global t_flaglote As f


'indice AGENCIA do arquivo CBA.DAT (NÃO DUPLICIDADE)
Type M
   Agenc_a As String * 5
End Type
Global t_age As M


'-----------------------------------------------------------------------------------------------------------------
'              ARQUIVOS TEXTO (.TXT) usados no sistema
'
' MÓDULO PRODUÇÃO:      - (05) RETORNOS.TXT (contém CMC-7 dos cheques captados)
'                       - (50) FLAGLOTE.TXT (diz se o arquivo Lote.Txt está sendo usado)
'                       - (51) LOTE.TXT     (contém o número do próximo lote a ser digitalizado)
'                       - (73) BAKKEY.TXT   (controle interno do sistema para checar a rotina de preparação do CMC-7)
'                       - (06) TPYYMMDD.TXT (contém status do lotes e cheques - situação)
'***
' MÓDULO RELATÓRIO:     - (25) I500I.TXT    (contém listagem dos cheques inferiores)
'                       - (30) I500S.TXT    (contém listagem dos cheques superiores)
'                       - (27) I500E.TXT    (contém listagem dos cheques excluidos)
'                       - (52) LIMITE.TXT   (contém valor do limite)
'***
' MÓDULO CORREÇÃO:      - (15) P00?????.TXT (informa o lote com pendência de CMC-7)
'***
' MÓDULO ENT. DADOS:    - (16) P00?????.TXT (informa o lote com pendência de Ent.Dados)
'***
' MÓDULO COMPRESSÃO:    - (60) COMPRESS.TXT (informa que compressão está sendo feita por algum servidor)
'***
' MÓDULO LOG:           - (75) LOG???? (arquivão de log)
' MÓDULO LOG:           - (76) LG?????? (arquivo de log de cada estação)
'***
' MÓDULO PRINCIPAL:     - (09) CADBAN.TXT   (contém número das agências do Bradesco)
'                       - (97) DIRPROD      (contém o PATH para o diretório de produção)
'                       - (96) DIRIMA       (contém o PATH para o diretório de imagens)
'***
' MÓDULO CONSULTA:      - (99) LE_LOTE.TXT  (listagem que contém todos os lotes digitalizados e seus status)
'                       - (88) LE_CHEQUE.TXT(listagem que contém todos os cheques digitalizados e seus status)
'                       - (96) DIRIMA       (contém o PATH para o diretório de imagens)
'***
' MÓDULO SENHA.BAS:     - (01) LS.INI       (arquivo .INI)
'***
' MÓDULO SENHA_PENDENTE:- (70) PENDENTE.TXT (informa que pendência de valor já está sendo feita por alguma estação)
'                         (71) PEND_COR.TXT (informa que pendência de correçào já está sendo feita por alguma estação)
'                         (72) CONSISTE.TXT (informa que consistência de correçào já está sendo feita por alguma estação)
' MÓDULO SENHA_DELÇÃO:  - (74) DELECAO.TXT  (informa que arquivos do CTREE não foram apagados)
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
   
   ' verifica se existe arquivo com número da máquina.
   nro_maquina = Dir$("c:\windows\system\*.log")
   If (nro_maquina = "") Then
      nro_maquina = "99MICROA.LOG"
   End If
   
   arquivo = "LG" + Mid$(data_inverte, 3, 2) + Mid$(data_inverte, 5, 2) + Mid$(nro_maquina, 1, 2)

   dir_log = Dir(diretorio_imagem & arquivo)

   '*******************************************************************************
   '*** linha_log contém as seguintes informações: lote + banco + agência + comp +
   '*** cheque + conta + agencia_vinc + conta_vinc + valor_vinc + operador + data da
   '*** operacao + operação
   '*******************************************************************************
   log_operador = user_acesso       'usuário de acesso

   If (user_acesso = "") Then
      log_operador = String(10, " ")
   Else
      'completa campo USUÁRIO com brancos no final.
      tamanho = Len(user_acesso)
      espacos = 10 - tamanho
      log_operador = user_acesso & String(espacos, " ")
   End If
   
   str_formatada = ""     'limpa string para guardar valor formatado
   rep = 1                'contador de caracteres a serem formatados
   valor_caption = Format$(Val(reg_ind.Valor_Vinc) / 100, "##########.00")

   Do
     If (Mid$(valor_caption, rep, 1) = "") Then   'verifica término da string
        Exit Do
     End If
     rep = rep + 1
   Loop While (rep < 14)   'tamanho máximo da string a ser formatada
   rep = rep - 1
            
   'formata à direita
   Mid$(str_formatada, 13 - rep + 1, rep) = Mid$(valor_caption, 1, rep)
   valor_caption = str_formatada    'atualiza valor_caption com dado formatado a direita

   
   linha_log = reg_recup.lote + branco + reg_ind.Banco + branco + reg_ind.agencia + branco + reg_ind.Comp + branco + reg_ind.cheque + branco + reg_ind.Conta + branco + reg_ind.Agencia_Vinc + branco + reg_ind.Conta_Vinc + branco + valor_caption + branco + log_operador + branco + Format$(Now, "DD/MM/YY") + branco + operacao
   
   'caso arquivo de LOG não exista, cria novo.
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
   
   '*** Inicializa variáveis do CTREE ***
   CtreeVars = InitCtreeGV()
   
   Principal.Show

End Sub

Function Trata_Msg_Erro(nro_msg As Integer, ident_msg As Integer)

   Select Case nro_msg
      Case 1
         MsgBox "Não foi possível carregar a imagem , Verifique se as imagens estão gravadas no diretório de origem! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 2
         MsgBox "Base de Dados não Encontrada! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 3
         MsgBox "Digite todos os campos! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 4
         MsgBox "Usuário não cadastrado! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 5
         MsgBox "Senha de Acesso incorreta! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 6
         MsgBox "Não foi possível ler o primeiro Lote do Banco de Dados! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 7
         MsgBox "Atualização no arquivo de lotes negada , Reinicie Processo! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 8
         MsgBox "Data Inválida! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 9
         MsgBox "Registro não encontrado! " + "(" + Str$(ident_msg) + ")", 48 + 0, "Atenção"
      Case 10
         MsgBox "Digite o dia e mês (DD/MM) da Data do Movimento! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 11
         MsgBox "Lote não encontrado no Banco de Dados Imagens! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 12
         MsgBox "O campo está incompleto. Digite todos os caracteres! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 13
         MsgBox "Arquivo de identificação não encontrado. Contate o Supervisor! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 14
         MsgBox "Não existe nenhum Lote a ser tratado! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 15
         MsgBox "Contador de Cheques não atualizado. Continue Processo! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 16
         MsgBox "Atualização do Arquivo de Cheques negada , Reinicie Processo! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 17
         MsgBox "Reinicie o Módulo para confirmação de término! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 18
         MsgBox "Digite o campo " + Str$(ident_msg) + " do CMC-7! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 19
         MsgBox "Reinicie a Pendência de Entrada de Dados! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 20
         MsgBox "Anote este número e informe ao Supervisor e/ou Suporte Datacheck: " + reg_ind.I_Frente + " -Comp - (" + Str$(ident_msg) + ")", MB_OK, "Atenção"
      Case 21
         MsgBox "Impressão de cheques negada , Reinicie Processo! " + "(" + Str$(ident_msg) + ")", MB_OK, "Atenção"
      Case 22
         MsgBox "Campo " + Str$(ident_msg) + " não confere! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 23
         MsgBox "Digite o Campo Valor! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 24
         MsgBox "Digite o Campo Agência! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 25
         MsgBox "Não existe nenhum Lote pertencente a esta Data! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 26
         MsgBox "Término da Geração do Arquivo Texto! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 27
         MsgBox "Atualização do novo valor do cheque negada , Reinicie Processo! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 28
         MsgBox "Não existe nenhum lote gravado! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 29
         MsgBox "Não existem mais cheques pertencentes a este Lote! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 30
         MsgBox "Reinicie CORREÇÃO! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 31
         MsgBox "Problema interno , Reinicie o procedimento! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 32
         MsgBox "Digite o Número do Lote! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 33
         MsgBox "Cheque não encontrado para Atualização do Valor. Contate o Suporte da Datacheck! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 34
         MsgBox "Lote não encontrado para Atualização do Valor. Contate o Suporte da Datacheck! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 35
         MsgBox "Não existe Lote Gravado nesta Data! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 36
         MsgBox "Leitura do valor do Lote para impressão negada , Reinicie Processo! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 37
         MsgBox "Todos os registros atuais serão deletados! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 38
         MsgBox "Exclusão da Tabela negada, Reinicie Processo! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 39
         MsgBox "Arquivo CADBAN.TXT não encontrado! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 40
         MsgBox "Atualização da Diferença do valor do cheque negada, Reinicie Processo! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 41
         MsgBox "Geração da Nova Tabela negada, Reinicie Processo! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 42
         MsgBox "É necessário fazer a Manutenção do CADBAN antes da produção! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 43
         MsgBox "Usuário já cadastrado! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 44
         MsgBox "Disco inválido! " + "(" + Str$(ident_msg) + ")", MB_ICONEXCLAMATION
      Case 45
         MsgBox "Selecione a Imagem a ser Impressa! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 46
         MsgBox "Selecione item do scroll para visualizar Imagem! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 47
         MsgBox "Término do processo! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 48
         MsgBox "Lote não encontrado! " + "(" + Str$(ident_msg) + ")", 48 + 0, "Atenção"
      Case 49
         MsgBox "Problema ao tentar ler cheques deste Lote! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 50
         MsgBox "Cheque não encontrado! " + "(" + Str$(ident_msg) + ")", 48 + 0, "Atenção"
      Case 51
         MsgBox "Problema ao carregar os registros na tela! " + "(" + Str$(ident_msg) + ")", MB_OK, "Atenção"
      Case 52
         MsgBox "Reinicie Entrada de Dados! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 53
         MsgBox "Base de Dados não Encontrada para fazer batimento das agências! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 54
         MsgBox "Arquivo de Lote nao encontrado , Contate o Supervisor! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 55
         MsgBox "Digite o número de Cheques pertencentes ao Lote! " + "(" + Str$(ident_msg) + ")", 48, " Atenção"
      Case 56
         MsgBox "Tratamento de abertura do arquivo (TXT) de lote negada, Reinicie Processo! " + "(" + Str$(ident_msg) + ")", MB_OK, "Atenção"
      Case 57
         MsgBox "Verificar Timer-Tic da LS-500! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 58
         MsgBox "Verifique se todas as LS-500 estão ligadas! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 59
         MsgBox "Gravação de CMC-7 negada , Reinicie Processo! " + "(" + Str$(ident_msg) + ")", MB_OK, "Atenção"
      Case 60
         MsgBox "Verifique se as LS-500 contém cheques! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 61
         MsgBox "Número de cheques não confere! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 62
         MsgBox "Problema na gravação de cheques, Refaça LOTE! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 63
         MsgBox "Problema na Gravação do Número do Lote! ...Refaça LOTE! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 64
         MsgBox "Digite apenas caracteres numéricos! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 65
         MsgBox "Este Documento não é uma capa de lote! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 66
         MsgBox "Ja foi informada a Capa deste lote! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 67
         MsgBox "CMC-7 não confere! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 68
         MsgBox "Digite os Campos Anteriores! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 69
         MsgBox "Agência inexistente! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 70
         MsgBox "Digite o campo Conta! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 71
         MsgBox "Conta inexistente! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 72
         MsgBox "Término do Processo de Correção do CMC-7! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 73
         MsgBox "Próximo Lote não lido, Reinicie o Processo! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 74
         MsgBox "Término do Lote, pois a interrupção do Processo de Correção do CMC-7 foi solicitada pelo operador! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 75
         MsgBox "Cópia já solicitada por outra estação! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 76
         MsgBox "Digite a Data do Movimento! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 77
         MsgBox "Digite a Data do Movimento para cheques inferiores! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 78
         MsgBox "Não encontrado valor limite. Para gravá-lo, vá para o menu Supervisor, entre na Opção <Novo Valor Limite>! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 79
         MsgBox "Atenção !!! Número de Lote já existente !!! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 80
         MsgBox "Não existe nenhum Lote com Pendência de Valor! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 81
         MsgBox "Não existem mais Lotes com Pendência de Valor! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 82
         MsgBox "Reinicie Processo! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 83
         MsgBox "Lote Batido! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 84
         MsgBox "Digite a senha de acesso! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 85
         MsgBox "Pendencia de Valor já solicitada por outra estação! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 86
         MsgBox "CMC-7 não confere! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 87
         MsgBox "Término do Processo de Pendência do CMC-7! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 88
         MsgBox "Término do Processo de Pendência de Entrada de Dados! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 89
         MsgBox "Agencia não existente no CADBAN! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 90
         MsgBox "Agência inexistente! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 91
         MsgBox "Digite o campo Conta! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 92
         MsgBox "Problema ao ler primeiro registro! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 93
         MsgBox "Não encontrado arquivo de LOG. Confirme a Data! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 94
         MsgBox "Arquivo de Estatística não encontrado. Contate o Supervisor! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 95
         MsgBox "Digite o Valor Mínimo para Impressão dos Cheques! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 96
         MsgBox "Próximo Lote não lido, Reinicie o Processo! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 97
         MsgBox "Pendência de Correção de CMC-7 já solicitada por outra estação! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 98
         MsgBox "Consistência de Correção de CMC-7 já solicitada por outra estação! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 99
         MsgBox "Não foi possível atualizar o registro(CMC76) desta imagem! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 100
         MsgBox "Não foi possível carregar esta imagem na tela! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 101
         MsgBox "Não foi possível imprimir a imagem " + reg_ind.I_Frente + " ! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 102
         MsgBox "Não existe arquivo de Log pertencente a esta data! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 103
         MsgBox "Esta Deleção só poderá ser feita do Servidor PRINCIPAL! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"
      Case 104
         MsgBox "Esta Deleção só poderá ser feita do Servidor SECUNDÁRIO! " + "(" + Str$(ident_msg) + ")", 48, "Atenção"

   
   End Select

End Function


