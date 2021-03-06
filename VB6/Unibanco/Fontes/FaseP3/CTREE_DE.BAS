Attribute VB_Name = "CTREE_DEF"

'////////////////////////////////////////////////////
'////    DEFINI��ES DAS ESTRUTURAS - CTREE    ///////
'////////////////////////////////////////////////////


'---- VARI�VEIS USADAS PARA CTREE ----
Global CtreeVars As Long
Global lixo As Integer
Global lixao As Integer
Global llock As Integer
Global Current_Indice As Integer


'---- N�MEROS DOS ARQUIVOS/INDICES NO PARAMETER FILE ----

Global Const ArqLok = 0             'numero do arq de LOK

Global Const IndDat = 1             'numero do arq de Dados - Ind
Global Const IndDtEnvImg = 5        'numero do Indice Data+Envelope+Imagem
Global Const IndEnvStPeSq = 6       'numero do Indice Envelope+Status+Pendencia+SeqImg
Global Const IndBcAgCtCh = 7        'numero do Indice Bco+Ag+Cta+Cheque
Global Const IndEnvDoctoVlId = 8    'numero do Indice Envelope+TipoDocto+Valor+IdChPg
Global Const IndDuplicidade = 9     'numero do Indice Duplicidade

Global Const RecupDat = 2           'numero do arq de Dados - Recup
Global Const RecupDtEnv = 10        'numero do Indice Data+Envelope
Global Const RecupStPeHo = 11       'numero do Indice Status+Pend�ncia+Hora_consulta
Global Const RecupDtAgEnv = 12      'numero do indice Data + Agencia + Envelope

Global Const LogDat = 3             'numero do arq de dados - Log
Global Const LogEnv = 13            'numero do indice Envelope
Global Const LogUsu = 14            'numero do indice Usu�rio

Global Const ContDat = 4            'numero do arq de Dados - Controles
Global Const ContDt = 15            'numero do Indice DataMovimento

'-------------------------------------------------------------------------------------------
Type Est_Indices                 'VB      CTREE
   Data_Movimento As String * 8  '1       0        data (AAMMDD)        chave
   Nro_Envelope As String * 11   '9       8        envelope             chave
   I_Frente As String * 15       '20      19       imagem frente        chave
   Status_Docto As String * 1    '35      34       percurso docto       chave
   Pendencia As String * 2       '36      35       pendencia            chave
   Tipo_docto As String * 4      '38      37       tipo docto
   Valor_Doc As String * 12      '42      41       valor
   dados As String * 145         '54      53       dados dos documentos digitados
   Agencia_Vinc As String * 5    '199     198      ag�ncia vinculada
   Conta_Vinc As String * 8      '204     203      conta vinculada
   Codigo_Barras As String * 44  '212     211      conteudo do Codigo de Barras (lido ou digitado)
   Hora_Consulta As String * 8   '256     255      guarda nsu, cod ocorrencia
   Cod_trans As String * 4       '264     263      transa��o
   Seq_img As String * 3         '268     267      sequencia da imagem no envelope
   nro_log_trans As String * 7   '271     270
   Num_Deposito As String * 6    '278     277      n�mero do dep�sito associado
   cmc7_assoc As String * 30     '284     283      cmc-7 do cheque que est� associado a conta
   NSU As String * 6             '314     313      NSU para autentica��o
   Ident_ChPg As String * 1      '320     319      identifica se ja fez associa��o de cheque/pagamento
   Duplicidade As String * 44    '321     320      guarda cmc-7(30) ou cod.barras(44) ou nro_imagem(15)
End Type
    
Global reg_ind As Est_Indices
'tamanho desta estrutura = 364


'--------------- IND.DAT -------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
'|  Envelope          |  Status  |  Pendencia  | Coment�rio
'
'|   nro_envelope     |     3    |      00     | pronto para LOG
'|        "           |     3    |      10     | pendencia sup. IO
'|        "           |     3    |      30     | devolve docto
'|        "           |     3    |      01     | vai para consulta,incrementa contador Recup
'|        "           |     3    |      02     | aguardando consulta
'|        "           |     4    |      00     | gerou LOG,decrementar contador Recup
'-------------------------------------------------------------------------------------------

'---- Campo STATUS_DOCTO ----
'
' 0 - digitalizado -- leitura CMC7 / BAR CODE  OK
' 1 - digitalizado -- erro leitura CMC7
' 2 - digitalizado -- erro leitura BAR CODE
' 3 - digitado/pronto para enviar consulta/aguardando consulta
' 4 - gerou LOG
' D - devolvido (qdo n�o puder vincular ou qdo arrec.convenc n�o encontrada ou qdo devolvido pelo supervisor)
' 9 - excluido  (pelo usuario ou pelo restri��o do LOG)

'---- Campo PENDENCIA ----
'
' 00 - sem pendencia supervisor
' 30 - com devolu��o docto

'---- Campo TIPO_DOCTO ap�s digitaliza��o ----
'
' 1111 - Cheque+dep�sito
' 2222 - doctos a pagar
' 3333 - envelope

'---- Campo TIPO_DOCTO ap�s digita��o ----
'
' 1111 (dinheiro) - Cheque   digitalizado pela LS500
' 1112 (conta)    - Dep�sito digitalizado pela LS500
' 1113 (dinheiro) - Tesouraria
' 1114 (dinheiro) - ADCC
' 2222 (conta)    - doctos a pagar
' 3333 (envelope) - envelope
' 4444 (debito)   - despesa gerada para diferen�a <=0,99   - evento 836
' 5555 (credito)  - lan�amento interno para ajuste <= 0,99 - evento 835
' 6666 (credito)  - lan�amento interno para ajuste deposito - evento 834
' 7777 (debito)   - lan�amento interno para ajuste deposito - evento 833
' 9994 (conta)    - Credito c/c criado
' 9995 (dinheiro) - Debito  c/c criado
' 8884 (conta)    - Credito c/p criado
' 8885 (dinheiro) - Debito  c/p criado
' 9992 (credito)  - Credito c/c criado


        
'---- Campo COD_TRANS ap�s digita��o ----
' 0102         - saque
' 0112         - adcc
' 0307         - deposito interagencia c/c
' 0308         - deposito interagencia c/p
' 0367         - arrecadacao convencional
' 0368         - darf preto
' 0369         - darf simples
' 0370         - tributo federdal
' 0371         - tributo estadual
' 0372         - tributo municipal
' 0380 To 0386 - Concessionarias
' 0410         - Cobran�a terceiros
' 0412         - Cobran�a Imediata UNIBANCO
' 0413         - cobran�a especial
' 0431         - UNIcobran�a
' 0333         - envelope
' 0444         - despesa gerada para diferen�a <= 0,99  - evento 836
' 0555         - lan�amento interno para ajuste <= 0,99 - evento 835
' 0666         - lan�amento interno credito - evento 834
' 0777         - lan�amento interno debito  - evento 833
' 0960         - cheque
' 8884         - credito de poupan�a - ajuste
' 8885         - debito de poupan�a - ajuste
' 9994         - credito de conta corrente - ajuste
' 9995         - debito conta corrente - ajuste


'---- Conte�do do campo Dados por tipo de documento
'
' 1111 (dinheiro) - Cheque   digitalizado pela LS500
'
'   Banco       1   3
'   Agencia     4   4
'   C1          8   1
'   Conta       9   6
'   C2          15  1
'   Comp        16  3
'   Cheque      19  6
'   C3          25  1
'   Adicional   26  5



'-------------------------------------------------------------------------------------------
Type Est_Recuperacao                'VB      CTREE
   Data_Movimento As String * 8     '1       0        data (AAMMDD)        chave
   Nro_Envelope As String * 11      '9       8        envelope             chave
   Status_Envelope As String * 1    '20      19       percurso envelope    chave
   Pendencia As String * 3          '21      20       pend�ncia            chave
   Hora_Consulta As String * 8      '24      23       hora inicio consulta chave
   Nosso_Numero As String * 5       '32      31       nosso numero de envelope
   Doctos_Cons As String * 3        '37      36       qtd doctos a serem consultados
   Qtd_docs As String * 3           '40      39       qtd doctos do envelope
   Docs_dig As String * 3           '43      42       qtd doctos digitados
   Usuario As String * 10           '46      45       nome usuario realizou transa��o
   Vl_Conta As String * 12          '56      55       valor das contas
   Vl_Dinheiro As String * 12       '68      67       valor em dinheiro/cheque/d�bito
   Vl_Diferenca As String * 12      '80      79       valor diferen�a entre vl_contas e vl_dinheiro
   Ag_orig As String * 4            '92      91       Agencia de origem do envelope
   Cta_orig As String * 7           '96      95       Conta de origem do envelope
   Valor_Especie As String * 12     '103     102      valor em dinheiro
   Existe_Duplicidade As String * 1 '115     114      indica se ha duplicidade no envelope
End Type

Global reg_recup As Est_Recuperacao
'tamanho desta estrutura = 115


'--------------- RECUP.DAT -------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
'| Status |  Pendencia  | Coment�rio
'
'|   6    |     000     | pronto para LOG
'|   6    |     100     | pendencia de valor
'|   6    |     010     | pendencia de sup
'|   6    |     020     | pendencia de alcada (n�o tem )
'|   6    |     030     | devolu��o de envelope
'|   7    |     000     | gerou LOG
'-------------------------------------------------------------------------------------------

'---- Campo STATUS_ENVELOPE ----
'
' 0 - cadastrado
' 1 - enviado a programa��o
' 2 - enviado a tesouraria
' 3 - enviado a prepara��o
' 4 - digitalizado
' 5 - compactado
' X - est� sendo feita digita��o
' 6 - digitado/pronto para enviar consulta/aguardando consulta
' Y - est� sendo feita pend�ncia de valor
' S - est� com pendencia de supervisor
' V - est� com pendencia de supervisor para fazer o vinculo de ch/pagto
' Z - est� sendo feito o vinculo de doctos
' 7 - gerou LOG
' 8 - expedido
' 9 - devolve envelope inteiro

'---- Campo PENDENCIA ----
'
' 000 - sem pendencia valor
' 100 - com pendencia valor
' 999 - envelope tem numer�rio;ap�s digitalizar este campo passa para '000'



'----------------------------------------------------------------
Type Est_Log                        'VB      CTREE
   Envelope As String * 11          '1       0        envelope          chave
   Usuario As String * 10           '12      11       nome usuario      chave
   Data_Log As String * 8           '22      21       data (AAMMDD)
   Hora_Log As String * 8           '30      29       hora (HH:MM:SS)
   Acao As String * 80              '38      37       a��o realizada
End Type

Global reg_log As Est_Log
'tamanho desta estrutura = 117



'----------------------------------------------------------------
Type Est_Controle                      'VB      CTREE
   Data_Movimento As String * 8        '1       0
   Base_NSU As String * 7              '9       8
End Type

Global reg_cont As Est_Controle
'tamanho desta estrutura = 15


'******************************************************************************
' +++  BUFFERS DE PESQUISA  +++
'******************************************************************************

'////////////////////////////  IND.DAT  ///////////////////////////////////////

' indice DATA+ENVELOPE+IMAGEM do arquivo IND.DAT (N�O DUPLICIDADE)
Type Ind_1
    DtEnvImg As String * 34      '8+11+15
End Type
Global id_dei As Ind_1

' indice ENVELOPE+STATUS+PENDENCIA+SEQIMG do arquivo IND.DAT (N�O DUPLICIDADE)
Type Ind_2
   EnvStPeSq As String * 17      '11+1+2+3
End Type
Global id_espc As Ind_2


' indice BANCO+AGENCIA+CONTA+CHEQUE do arquivo IND.DAT (DUPLICIDADE)
Type Ind_3
   BcAgCtCh As String * 20       '3+4+7+6
   Deslo As Long
End Type
Global id_bacc As Ind_3


' indice ENVELOPE+TIPO_DOCTO+VALOR_DOCTO+IDENT_CHPG do arquivo IND.DAT (DUPLICIDADE)
Type Ind_4
   EnvDoctoVlId As String * 28   '11+4+12+1
   Deslo As Long
End Type
Global id_docvl As Ind_4


' indice DUPLICIDADE do arquivo IND.DAT (N�O DUPLICIDADE)
Type Ind_5
   Duplicidade As String * 44   '44
End Type
Global id_duplic As Ind_5


'/////////////////////////////  RECUP.DAT  //////////////////////////////////////

' indice DATA+ENVELOPE do arquivo RECUP.DAT (N�O DUPLICIDADE)
Type Ind_6
   DtEnv As String * 19          '8+11
End Type
Global rp_de As Ind_6

' indice STATUS+PENDENCIA+HORA do arquivo RECUP.DAT (DUPLICIDADE)
Type Ind_7
   StPeHo As String * 12         '1+3+8
   Deslo As Long
End Type
Global rp_sph As Ind_7

' indice DATA+AGENCIA+ENVELOPE do arquivo RECUP.DAT (N�O DUPLICIDADE)
Type Ind_8
   DtAgEnv As String * 23        '8+4+11
End Type
Global rp_dae As Ind_8


'/////////////////////////////  LOG.DAT  //////////////////////////////////////

' indice ENVELOPE do arquivo LOG.DAT (DUPLICIDADE)
Type Ind_9
   Env As String * 11            '11
   Deslo As Long
End Type
Global lg_e As Ind_9

' indice USUARIO do arquivo LOG.DAT (DUPLICIDADE)
Type Ind_10
   N_Usuario As String * 10      '10
   Deslo As Long
End Type
Global lg_u As Ind_10



'/////////////////////////////  CONT.DAT  //////////////////////////////////////

' indice DATA do arquivo CONT.DAT (N�O DUPLICIDADE)
Type Ind_11
   data As String * 8            '8
End Type
Global Ct_dat As Ind_11

