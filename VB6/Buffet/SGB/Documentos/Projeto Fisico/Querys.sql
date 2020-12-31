-- Valor Total de cheques recebidos
select sum(valor_par) as TOTAL_CHEQUES 		from parcela_contrato

-- Valor Total de cheques recebidos
-- Cheques a serem depositados em cada mes
select sum(valor_par) as CHEQUES_FEVEREIRO 	from parcela_contrato where convert(char,data_par,112) between '20050201' and '20050228' and COMPENSADO = 'N'
select sum(valor_par) as CHEQUES_MARCO 		from parcela_contrato where convert(char,data_par,112) between '20050301' and '20050331' and COMPENSADO = 'N'
select sum(valor_par) as CHEQUES_ABRIL	 	from parcela_contrato where convert(char,data_par,112) between '20050401' and '20050430' and COMPENSADO = 'N'

-- Qtde de festas de cada mes
select count(id_festa) as FESTAS_JANEIRO 	from festas where convert(char,data_festa,112) between '20050101' and '20050131'
select count(id_festa) as FESTAS_FEVEREIRO 	from festas where convert(char,data_festa,112) between '20050201' and '20050228'
select count(id_festa) as FESTAS_MARCO 		from festas where convert(char,data_festa,112) between '20050301' and '20050331'
select count(id_festa) as FESTAS_ABRIL 		from festas where convert(char,data_festa,112) between '20050401' and '20050430'

-- Qtde de festas fechadas de cada mes
select count(id_cnt) as CONTRATOS_JANEIRO	from CONTRATOS 	where convert(char,data_cnt,112) between '20050101' and '20050131'
select count(id_cnt) as CONTRATOS_FEVEREIRO 	from CONTRATOS 	where convert(char,data_cnt,112) between '20050201' and '20050228'
select count(id_cnt) as CONTRATOS_MARCO 	from CONTRATOS	where convert(char,data_cnt,112) between '20050301' and '20050331'
select count(id_cnt) as CONTRATOS_ABRIL 	from CONTRATOS 	where convert(char,data_cnt,112) between '20050401' and '20050430'

-- Qtde de festas realizadas
select count(0) as FESTAS_REALIZADAS from festas where convert(char,data_festa,112) <= getdate()

-- Qtde de festas a serem realizadas
select count(0) as FESTAS_A_SEREM_REALIZADAS from festas where convert(char,data_festa,112) > getdate()
