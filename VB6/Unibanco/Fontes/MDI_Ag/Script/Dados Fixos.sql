
set nocount on

		/*-----------------------------------
		   INSERT DOS REGISTROS NAS TABELAS 
		-----------------------------------*/
--STATUSCAPA


INSERT INTO StatusCapa VALUES ('0' ,'Capa cadastrada')
INSERT INTO StatusCapa VALUES ('1' ,'Capa digitalizada')
INSERT INTO StatusCapa VALUES ('2' ,'Capa em complementacao')
INSERT INTO StatusCapa VALUES ('3' ,'Capa complementada, mas com pendencia')
INSERT INTO StatusCapa VALUES ('4' ,'Capa para Prova Zero')
INSERT INTO StatusCapa VALUES ('5' ,'Capa para Ilegiveis')
INSERT INTO StatusCapa VALUES ('6' ,'Capa para Alcada')
INSERT INTO StatusCapa VALUES ('7' ,'Capa para Vinculo Manual')
INSERT INTO StatusCapa VALUES ('8' ,'Capa para Vinculo Automatico')
INSERT INTO StatusCapa VALUES ('9' ,'Capa p/ Vinc. Automatico, enviada pelo Prova Zero')
INSERT INTO StatusCapa VALUES ('D' ,'Capa Devolvida pelo Sistema')
INSERT INTO StatusCapa VALUES ('E' ,'Capa Expedida')
INSERT INTO StatusCapa VALUES ('F' ,'Capa Devolvida pelo Caixa Robo')
INSERT INTO StatusCapa VALUES ('G' ,'Capa em Prova Zero')
INSERT INTO StatusCapa VALUES ('H' ,'Capa em Ilegiveis')
INSERT INTO StatusCapa VALUES ('I' ,'Capa em Alcada')
INSERT INTO StatusCapa VALUES ('J' ,'Capa em Vinculo Manual')
INSERT INTO StatusCapa VALUES ('K' ,'Capa em Expedicao')
INSERT INTO StatusCapa VALUES ('O' ,'Capa em Troca de Ordem')
INSERT INTO StatusCapa VALUES ('P' ,'Capa Devolvida pela Preparacao')
INSERT INTO StatusCapa VALUES ('R' ,'Capa para Transmissao')
INSERT INTO StatusCapa VALUES ('S' ,'Capa em Transmissao')
INSERT INTO StatusCapa VALUES ('T' ,'Capa Transmitida')
INSERT INTO StatusCapa VALUES ('V' ,'Capa em Verificacao')
INSERT INTO StatusCapa VALUES ('X' ,'Capa ja enviada a ocorrencia para Ubb')

-- ACAO
INSERT INTO Acao VALUES ('1'   ,'Ilegiveis - Reenviar para Complementacao')
INSERT INTO Acao VALUES ('2'   ,'Ilegiveis - Documento registrado ocorrencia')
INSERT INTO Acao VALUES ('3'   ,'Ilegiveis - Devolver Envelope / Malote')
INSERT INTO Acao VALUES ('4'   ,'Ilegiveis - Enviar para Vinculo Automatico')
INSERT INTO Acao VALUES ('5'   ,'Ilegiveis - Documento Corrigido')
INSERT INTO Acao VALUES ('6'   ,'Ilegiveis - Enviar Capa para Troca de Ordem')
INSERT INTO Acao VALUES ('7'   ,'Ilegiveis - Remover Documento para Recaptura')
INSERT INTO Acao VALUES ('8'   ,'Ilegiveis - Enviar Capa para Recaptura')
INSERT INTO Acao VALUES ('10'  ,'Complementacao - Alterar Tipo de Documento')
INSERT INTO Acao VALUES ('11'  ,'Complementacao - Documento digitado')
INSERT INTO Acao VALUES ('12'  ,'Complementacao - Complementacao Automatica')
INSERT INTO Acao VALUES ('13'  ,'Complementacao - Devolver por Duplicidade (Auto)')
INSERT INTO Acao VALUES ('14'  ,'Complementacao - Cadastrar Envelope / Malote')
INSERT INTO Acao VALUES ('15'  ,'Complementacao - Enviar para Ilegiveis')
INSERT INTO Acao VALUES ('16'  ,'Complementacao - Enviar para Vinculo Auto. (Auto)')
INSERT INTO Acao VALUES ('17'  ,'Complementacao - Enviar para Ilegiveis (Auto)')
INSERT INTO Acao VALUES ('20'  ,'Recepcao - Envelope/Malote recepcionado')
INSERT INTO Acao VALUES ('21'  ,'Recepcao - Envelope/Malote registrado ocorrencia')
INSERT INTO Acao VALUES ('30'  ,'Controle de Qualidade - Remover Envelope / Malote')
INSERT INTO Acao VALUES ('31'  ,'Controle de Qualidade - Remover Documento')
INSERT INTO Acao VALUES ('40'  ,'Captura - Capturar Envelope / Malote')
INSERT INTO Acao VALUES ('41'  ,'Captura - Documento capturado')
INSERT INTO Acao VALUES ('50'  ,'Inicializacao - Criar Parametro')
INSERT INTO Acao VALUES ('51'  ,'Inicializacao - Inicializar Link')
INSERT INTO Acao VALUES ('60'  ,'Prova Zero - Documento corrigido valor')
INSERT INTO Acao VALUES ('61'  ,'Prova Zero - Documento corrigido')
INSERT INTO Acao VALUES ('62'  ,'Prova Zero - Devolver Envelope / Malote')
INSERT INTO Acao VALUES ('63'  ,'Prova Zero - Enviar para Ilegiveis')
INSERT INTO Acao VALUES ('64'  ,'Prova Zero - Enviar para Vinc. Auto. C/ Alt. Valor')
INSERT INTO Acao VALUES ('65'  ,'Prova Zero - Enviar para Vinc. Auto. Apos Conferir')
INSERT INTO Acao VALUES ('66'  ,'Prova Zero - Enviar para Troca de Ordem.')
INSERT INTO Acao VALUES ('70'  ,'Vinculo Manual - Enviar para Ilegiveis')
INSERT INTO Acao VALUES ('71'  ,'Vinculo Manual - Documento registrado ocorrencia')
INSERT INTO Acao VALUES ('72'  ,'Vinculo Manual - Documento vinculado manualmente')
INSERT INTO Acao VALUES ('73'  ,'Vinculo Manual - Inserir Ajuste Debito / Credito')
INSERT INTO Acao VALUES ('74'  ,'Vinculo Manual - Enviar para Alcada(Automatico)')
INSERT INTO Acao VALUES ('75'  ,'Vinculo Manual - Enviar para Transmissao(Auto)')
INSERT INTO Acao VALUES ('76'  ,'Vinculo Manual - Enviar para Analise')
INSERT INTO Acao VALUES ('80'  ,'Expedicao - Documento autenticado')
INSERT INTO Acao VALUES ('81'  ,'Expedicao - Reautenticar Documento')
INSERT INTO Acao VALUES ('82'  ,'Expedicao - Imprimir Ocorr. do Envelope / Malote')
INSERT INTO Acao VALUES ('83'  ,'Expedicao - Imprimir Ocorrencia do Documento')
INSERT INTO Acao VALUES ('84'  ,'Expedicao - Imprimir Comp. Ajuste Debito / Credito')
INSERT INTO Acao VALUES ('85'  ,'Expedicao - Imprimir Comprovante de Deposito')
INSERT INTO Acao VALUES ('86'  ,'Expedicao - Atualizar Env. / Mal. Expedido(Auto)')
INSERT INTO Acao VALUES ('87'  ,'Expedicao - Imprimir Comp. Cartao Avulso')
INSERT INTO Acao VALUES ('88'  ,'Expedicao - Entrar Capa')
INSERT INTO Acao VALUES ('89'  ,'Expedicao - Imprimir Comp. Pagamento')
INSERT INTO Acao VALUES ('90'  ,'Alcada - Documento liberado com alcada')
INSERT INTO Acao VALUES ('91'  ,'Alcada - Enviar Envelope / Malote p/ Trans.(Auto)')
INSERT INTO Acao VALUES ('92'  ,'Alcada - Enviar para Ilegiveis')
INSERT INTO Acao VALUES ('100' ,'Supervisao - Excluir Envelope / Malote')
INSERT INTO Acao VALUES ('110' ,'Vinc. Automatico - Enviar para Prova Zero')
INSERT INTO Acao VALUES ('111' ,'Vinc. Automatico - Enviar para Ilegiveis')
INSERT INTO Acao VALUES ('112' ,'Vinc. Automatico - Enviar para Alcada')
INSERT INTO Acao VALUES ('113' ,'Vinc. Automatico - Enviar para Vinculo Manual')
INSERT INTO Acao VALUES ('114' ,'Vinc. Automatico - Enviar para Transmissao')
INSERT INTO Acao VALUES ('115' ,'Vinc. Automatico - Enviar para Analise')
INSERT INTO Acao VALUES ('120' ,'Robo - Inicializacao')
INSERT INTO Acao VALUES ('121' ,'Robo - Transmite Capa')
INSERT INTO Acao VALUES ('122' ,'Robo - Transmite Documento')
INSERT INTO Acao VALUES ('123' ,'Robo - Grava Ocorrencia')
INSERT INTO Acao VALUES ('124' ,'Robo - Capa com Diferenca enviada para Ilegiveis')
INSERT INTO Acao VALUES ('130' ,'Troca de Ordem - Documento inserido')
INSERT INTO Acao VALUES ('131' ,'Troca de Ordem - Documento excluido')
INSERT INTO Acao VALUES ('132' ,'Troca de Ordem - Envio Env/Mal para Vinculo Aut.')
INSERT INTO Acao VALUES ('133' ,'Troca de Ordem - Reenvio para Ilegiveis')
INSERT INTO Acao VALUES ('134' ,'Troca de Ordem - Documento reordenado')
INSERT INTO Acao VALUES ('140' ,'Consulta - Consultar Capa')
INSERT INTO Acao VALUES ('150' ,'Complementacao - Split Capa (Inicial)')
INSERT INTO Acao VALUES ('151' ,'Complementacao - Split Capa (Final)')
INSERT INTO Acao VALUES ('152' ,'Complementacao - Split Capa Anterior (Inicial)')
INSERT INTO Acao VALUES ('153' ,'Complementacao - Split Capa Anterior (Final)')
INSERT INTO Acao VALUES ('190' ,'Ilegiveis - Enviar para Analise')
INSERT INTO Acao VALUES ('191' ,'Ilegiveis - Selecionar Capa')
INSERT INTO Acao VALUES ('192' ,'Ilegiveis - Deselecionar Capa')
INSERT INTO Acao VALUES ('193' ,'Prova Zero - Selecionar Capa')
INSERT INTO Acao VALUES ('194' ,'Prova Zero - Deselecionar Capa')
INSERT INTO Acao VALUES ('195' ,'Vinc. Manual - Selecionar Capa')
INSERT INTO Acao VALUES ('196' ,'Vinc. Manual - Deselecionar Capa')
INSERT INTO Acao VALUES ('197' ,'Alcada - Selecionar Capa')
INSERT INTO Acao VALUES ('198' ,'Alcada - Deselecionar Capa')

-- OCORRENCIA

INSERT INTO Ocorrencia VALUES ('1'   ,'Quantidade de Envelope maior que informada no Protocolo Remessa')
INSERT INTO Ocorrencia VALUES ('2'   ,'Quantidade de Envelope menor que informada Protocolo Remessa')
INSERT INTO Ocorrencia VALUES ('3'   ,'Envelope Vazio')
INSERT INTO Ocorrencia VALUES ('4'   ,'Envelope so com Dinheiro')
INSERT INTO Ocorrencia VALUES ('5'   ,'Envelope so com Cheque')
INSERT INTO Ocorrencia VALUES ('6'   ,'Envelope so com o Documento a ser Pago')
INSERT INTO Ocorrencia VALUES ('7'   ,'Envelope so com  Ficha de Deposito')
INSERT INTO Ocorrencia VALUES ('101' ,'Deposito em Dinheiro')
INSERT INTO Ocorrencia VALUES ('102' ,'Cheque para Deposito em Dinheiro')
INSERT INTO Ocorrencia VALUES ('104' ,'Deposito Dinheiro Agencia/Conta Invalida')
INSERT INTO Ocorrencia VALUES ('105' ,'Deposito em Cheque para Agencia/Conta Invalida')
INSERT INTO Ocorrencia VALUES ('106' ,'Deposito Misto em Cheque e em Dinheiro')
INSERT INTO Ocorrencia VALUES ('107' ,'Deposito em Poupanca, apos horario corte')
INSERT INTO Ocorrencia VALUES ('108' ,'Deposito em Cheque c/ desdobramento varias fichas de deposito')
INSERT INTO Ocorrencia VALUES ('109' ,'Deposito Cheque c/ Irregularidade ( sem Assinatura, Erro  Preenchimento ,Validade')
INSERT INTO Ocorrencia VALUES ('110' ,'Deposito em C/C Inexistente / Encerrada')
INSERT INTO Ocorrencia VALUES ('111' ,'Deposito em Conta Corrente na situacao  2025')
INSERT INTO Ocorrencia VALUES ('112' ,'Deposito em C/C na condicao CC5 - Residentes no Exterior')
INSERT INTO Ocorrencia VALUES ('113' ,'Deposito em Conta Corrente Paralisada')
INSERT INTO Ocorrencia VALUES ('114' ,'Deposito em C/C na condicao MR')
INSERT INTO Ocorrencia VALUES ('115' ,'Deposito em C/C  bloqueada')
INSERT INTO Ocorrencia VALUES ('120' ,'Deposito com Moedas')
INSERT INTO Ocorrencia VALUES ('121' ,'Deposito a maior - Valor do cheque menor que valor informado na ficha de deposito')
INSERT INTO Ocorrencia VALUES ('122' ,'Deposito a menor - Valor do cheque maior que valor informado na ficha de deposito')
INSERT INTO Ocorrencia VALUES ('123' ,'Deposito com Cheque UBB')
INSERT INTO Ocorrencia VALUES ('124' ,'Deposito com Cheque de Outros Bancos')
INSERT INTO Ocorrencia VALUES ('201' ,'Pagto em Dinheiro')
INSERT INTO Ocorrencia VALUES ('202' ,'Pagto Conta/Titulo com Cheque e Dinheiro')
INSERT INTO Ocorrencia VALUES ('203' ,'Pagto Conta/Titulo com Cheque valor menor ao valor documento a ser liquidado')
INSERT INTO Ocorrencia VALUES ('204' ,'Pagto Conta/Titulo com Cheque valor maior  ao valor do documento a ser liquidado')
INSERT INTO Ocorrencia VALUES ('205' ,'Pagto Conta/Titulo Quitada')
INSERT INTO Ocorrencia VALUES ('206' ,'Pagto Conta/Titulo nao aceito pelo Unibanco')
INSERT INTO Ocorrencia VALUES ('207' ,'Pagto Conta/Titulo nao aceito pelo UBB c/ Conta/Titulo aceito UBB mesmo envelope')
INSERT INTO Ocorrencia VALUES ('208' ,'Pagto Titulo Vencido')
INSERT INTO Ocorrencia VALUES ('209' ,'Pagto Contas Cheque UBB ou Outros Bancos, sem descricao Finalidade no verso')
INSERT INTO Ocorrencia VALUES ('210' ,'Pagto varios Documentos UBB com apenas um cheque de outro Banco')
INSERT INTO Ocorrencia VALUES ('211' ,'Pagto apenas uma Conta/Titulo com diversos cheques de Outros Bancos')
INSERT INTO Ocorrencia VALUES ('212' ,'Pagto Conta/Titulo emitido por terceiros a ser Pago Cheque Outro Banco')
INSERT INTO Ocorrencia VALUES ('213' ,'Pagto Conta/Titulo Cheque UBB c/Insuficiencia Saldo')
INSERT INTO Ocorrencia VALUES ('214' ,'Irregularidade no numero DARF')
INSERT INTO Ocorrencia VALUES ('215' ,'Pagto Conta/Titulo c/ irreguralaridade Cheque (sem assinatura, erro preenchimento')
INSERT INTO Ocorrencia VALUES ('216' ,'Pagto Conta/Titulo com Cheque de conta corrente Encerrada/Inexistente')
INSERT INTO Ocorrencia VALUES ('217' ,'Pagto Conta/Titulo com cheque de conta corrente na situacao 2025')
INSERT INTO Ocorrencia VALUES ('218' ,'Pagto Conta/Titulo com Cheque Conta Corrente condicao CC5 - Residente Exterior')
INSERT INTO Ocorrencia VALUES ('219' ,'Pagto Conta/Titulo com Cheque Conta Corrente Paralisada')
INSERT INTO Ocorrencia VALUES ('220' ,'Pagto Conta/Titulo com Cheque Conta Corrente na Condicao MR')
INSERT INTO Ocorrencia VALUES ('221' ,'Pagto Conta/Titulo com Cheque Conta Corrente na Condicao CL/AD')
INSERT INTO Ocorrencia VALUES ('222' ,'Pagto Conta/Titulo com Cheque c/ Conta com Saldo Bloqueado')
INSERT INTO Ocorrencia VALUES ('223' ,'Pagto Conta/Titulo c/ Cheque,  Conta Saldo Bloqueado e fora Limite Contratual')
INSERT INTO Ocorrencia VALUES ('224' ,'Pagto Conta faltando via ou Titulo com apenas uma via')
INSERT INTO Ocorrencia VALUES ('226' ,'Pagto Conta/Titulo sem Cheque')
INSERT INTO Ocorrencia VALUES ('227' ,'Pagto Conta/Titulo documentos cadastrados em Debito Automatico')
INSERT INTO Ocorrencia VALUES ('228' ,'Pagto Conta/Titulo com Cheque Bloqueado motivo 29 - sustado')
INSERT INTO Ocorrencia VALUES ('229' ,'Pagto Conta/Titulo c/ cheque sustado.')
INSERT INTO Ocorrencia VALUES ('234' ,'Irregularidade no num./codigo do documento')
INSERT INTO Ocorrencia VALUES ('301' ,'Retirada da Poupanca')
INSERT INTO Ocorrencia VALUES ('302' ,'DOC')
INSERT INTO Ocorrencia VALUES ('303' ,'Requisicao de Talao de Cheque')
INSERT INTO Ocorrencia VALUES ('304' ,'Nao e envelope do caixa expresso')
INSERT INTO Ocorrencia VALUES ('306' ,'Resumo de Vendas de Cartao de Credito')
INSERT INTO Ocorrencia VALUES ('401' ,'Pagto Conta/Titulo, por Autor. de Debito, assinatura nao confere.')
INSERT INTO Ocorrencia VALUES ('402' ,'Pagto Conta/Titulo, por Autor. de Debito, com valor a menor')
INSERT INTO Ocorrencia VALUES ('403' ,'Pagto Conta/Titulo, por Autor. de Debito, com valor a maior')
INSERT INTO Ocorrencia VALUES ('404' ,'Pagto Conta/Titulo Nao Aceito Unibanco,   atraves de Autorizacao de Debito')
INSERT INTO Ocorrencia VALUES ('405' ,'Pagto Conta/Titulo, sem assinatura na Autorizacao de Debito')
INSERT INTO Ocorrencia VALUES ('406' ,'Pagto Conta/Titulo, sem descrever o valor a ser debitado na Autorizacao de Debito')
INSERT INTO Ocorrencia VALUES ('407' ,'Pagto Conta/Titulo, com Cheque Terceiro e Autor. de Debito')
INSERT INTO Ocorrencia VALUES ('408' ,'Pagto Conta/Titulo, com Cheque UBB e Autor de Debito')
INSERT INTO Ocorrencia VALUES ('409' ,'Pagto Conta/Titulo, com Dinheiro e Autor. de Debito')
INSERT INTO Ocorrencia VALUES ('411' ,'Pagto Conta/Titulo Vencidos, atraves de Autorizacao de Debito')
INSERT INTO Ocorrencia VALUES ('412' ,'Pagto Conta/Titulo Quitado, atraves Autorizacao de Debito')
INSERT INTO Ocorrencia VALUES ('413' ,'Pagto Conta/Titulo nao aceito junto com outro aceito, por Autor. Debito')
INSERT INTO Ocorrencia VALUES ('414' ,'Pagto Conta/Titulo, atraves de Autorizacao de Debito, com Conta Invalida')
INSERT INTO Ocorrencia VALUES ('415' ,'Pagto Conta/Titulo, atraves de Autorizacao de Debito, sem assinatura capturada')
INSERT INTO Ocorrencia VALUES ('416' ,'Pagto Conta/Titulo c/ Autor. de Debito e  Cheque, ambos valor igual total')
INSERT INTO Ocorrencia VALUES ('417' ,'Autorizacao de Debito p/ Deposito')
INSERT INTO Ocorrencia VALUES ('419' ,'Pagto Conta/Titulo, por Autor. de Debito, conta corrente Encerrada/Inexistente')
INSERT INTO Ocorrencia VALUES ('420' ,'Pagto Conta/Titulo, por Autor. de Debito, conta corrente na situacao 2025.')
INSERT INTO Ocorrencia VALUES ('421' ,'Pagto Conta/Titulo, por Autor. de Debito, C/C condicao CC5 - Residente Exterior')
INSERT INTO Ocorrencia VALUES ('422' ,'Pagto Conta/Titulo, por Autor. de Debito, Conta Corrente Paralisada')
INSERT INTO Ocorrencia VALUES ('424' ,'Pagto Conta/Titulo, por Autor. de Debito, Conta Corrente na Condicao CL/AD')
INSERT INTO Ocorrencia VALUES ('425' ,'Pagto Conta/Titulo, por Autor. de Debito, C/C  bloqueada')
INSERT INTO Ocorrencia VALUES ('426' ,'Pagto Conta/Titulo, por Autor. de Debito, c/ Conta com Saldo Bloqueado.')
INSERT INTO Ocorrencia VALUES ('429' ,'Pagto Conta/Titulo, por Autor. de Debito, c/Insuficiencia Saldo.')
INSERT INTO Ocorrencia VALUES ('430' ,'Autorizacao de Debito para assinatura inexistente.')
INSERT INTO Ocorrencia VALUES ('431' ,'Autorizacao de Debito Acima de R$2.000,00.')
INSERT INTO Ocorrencia VALUES ('432' ,'Autorizacao de Debito para Pessoa Juridica.')
INSERT INTO Ocorrencia VALUES ('998' ,'Documento em Duplicidade')
INSERT INTO Ocorrencia VALUES ('999' ,'Erro Operacional')


-- PARAMETRO
INSERT INTO Parametro
       (DataProcessamento,
	Hm_Abertura,
	Hm_Fechamento,
	AgenciaCentral,
	AgenciaApresentante,
	Tm_Pendente,
	Tm_Atualizacao,
	Dir_Dados,
	Dir_Imagens,
	Dir_Trabalho
	)
VALUES
       (20001127, 		--DataProcessamento
	GETDATE(),		--Hm_Abertura
	GETDATE(),		--Hm_Fechamento
	9999,			--AgenciaCentral
	9999,			--AgenciaApresentante
	300,			--Tm_Pendente
	30,			--Tm_Atualizacao
	'c:\mdi_ag\dados',	--Dir_Dados
	'c:\mdi_ag\imagens',	--Dir_Imagens
	'c:\mdi_ag\trabalho\'	--Dir_Trabalho
	)


-- STATUSDOCUMENTO
INSERT INTO StatusDocumento VALUES ('0'   ,'Documento nao complementado')
INSERT INTO StatusDocumento VALUES ('1'   ,'Documento complementado')
INSERT INTO StatusDocumento VALUES ('2'   ,'Documento em transmissao')
INSERT INTO StatusDocumento VALUES ('3'   ,'Documento com ocorrencia em transmissao')
INSERT INTO StatusDocumento VALUES ('A'   ,'Documento para Alcada')
INSERT INTO StatusDocumento VALUES ('D'   ,'Documento deletado pelo sistema')
INSERT INTO StatusDocumento VALUES ('E'   ,'Documento Expedido')
INSERT INTO StatusDocumento VALUES ('F'   ,'Documento deletado pelo caixa robo')
INSERT INTO StatusDocumento VALUES ('G'   ,'Acerto de Debito/Credito gerado por diferenca')
INSERT INTO StatusDocumento VALUES ('T'   ,'Documento Transmitido')

--STATUSLOTE
INSERT INTO StatusLote VALUES ('0'   ,'Lote digitalizado')
INSERT INTO StatusLote VALUES ('1'   ,'Lote em liberacao')
INSERT INTO StatusLote VALUES ('2'   ,'Lote liberado')
INSERT INTO StatusLote VALUES ('3'   ,'Lote em Captura')


--TIPODOCTO
INSERT INTO TipoDocto VALUES ('0'   ,'DOCUMENTO INDEFINIDO')
INSERT INTO TipoDocto VALUES ('1'   ,'CAPA ENVELOPE / MALOTE EMPRESA')
INSERT INTO TipoDocto VALUES ('2'   ,'DEPOSITO CONTA CORRENTE')
INSERT INTO TipoDocto VALUES ('3'   ,'DEPOSITO CONTA POUPANCA')
INSERT INTO TipoDocto VALUES ('4'   ,'AUTORIZACAO DE DEBITO EM C/C')
INSERT INTO TipoDocto VALUES ('5'   ,'CHEQUE UBB SACADO')
INSERT INTO TipoDocto VALUES ('6'   ,'CHEQUE COMPENSADO (CP)')
INSERT INTO TipoDocto VALUES ('7'   ,'CHEQUE DEPOSITO')
INSERT INTO TipoDocto VALUES ('8'   ,'CONCESSIONARIA VALOR REAL')
INSERT INTO TipoDocto VALUES ('9'   ,'CONCESSIONARIA VALOR INDEXADO')
INSERT INTO TipoDocto VALUES ('10'  ,'FICHA COMPENSACAO')
INSERT INTO TipoDocto VALUES ('11'  ,'INSS')
INSERT INTO TipoDocto VALUES ('12'  ,'TIT. OUTROS BCOS CONVENCIONAL')
INSERT INTO TipoDocto VALUES ('13'  ,'COBRANCA REGISTRADA (SEM CB)')
INSERT INTO TipoDocto VALUES ('14'  ,'COBRANCA ESPECIAL (SEM CB)')
INSERT INTO TipoDocto VALUES ('15'  ,'DARM')
INSERT INTO TipoDocto VALUES ('16'  ,'DARF PRETO')
INSERT INTO TipoDocto VALUES ('17'  ,'DARF SIMPLES')
INSERT INTO TipoDocto VALUES ('18'  ,'GARE')
INSERT INTO TipoDocto VALUES ('19'  ,'GRPS')
INSERT INTO TipoDocto VALUES ('20'  ,'AGUA')
INSERT INTO TipoDocto VALUES ('21'  ,'GAS')
INSERT INTO TipoDocto VALUES ('22'  ,'LUZ')
INSERT INTO TipoDocto VALUES ('23'  ,'TELEFONE')
INSERT INTO TipoDocto VALUES ('24'  ,'TRIBUTOS MUNICIPAIS')
INSERT INTO TipoDocto VALUES ('25'  ,'TRIBUTOS ESTADUAIS')
INSERT INTO TipoDocto VALUES ('26'  ,'TRIBUTOS FEDERAIS')
INSERT INTO TipoDocto VALUES ('27'  ,'ARRECADACAO CONVENCIONAL')
INSERT INTO TipoDocto VALUES ('28'  ,'UNICOBRANCA UBB')
INSERT INTO TipoDocto VALUES ('29'  ,'COBRANCA IMEDIATA UBB')
INSERT INTO TipoDocto VALUES ('30'  ,'COBRANCA ESPECIAL UBB')
INSERT INTO TipoDocto VALUES ('31'  ,'TITULO OUTROS BCOS ELETRONICO')
INSERT INTO TipoDocto VALUES ('32'  ,'AJUSTE CREDITO')
INSERT INTO TipoDocto VALUES ('33'  ,'AJUSTE DEBITO')
INSERT INTO TipoDocto VALUES ('34'  ,'CREDITO AUTOMATICO')
INSERT INTO TipoDocto VALUES ('35'  ,'GPS')
INSERT INTO TipoDocto VALUES ('36'  ,'CARTAO AVULSO')
INSERT INTO TipoDocto VALUES ('37'  ,'OCT')
INSERT INTO TipoDocto VALUES ('38'  ,'DEBITO AUTOMATICO')
INSERT INTO TipoDocto VALUES ('39'  ,'CAPA OCT')
INSERT INTO TipoDocto VALUES ('40'  ,'FGTS')
INSERT INTO TipoDocto VALUES ('41'  ,'LANCAMENTO INTERNO')
INSERT INTO TipoDocto VALUES ('42'  ,'AJUSTE CONTABIL RECEITA')
INSERT INTO TipoDocto VALUES ('43'  ,'AJUSTE CONTABIL DESPESA')


set nocount off