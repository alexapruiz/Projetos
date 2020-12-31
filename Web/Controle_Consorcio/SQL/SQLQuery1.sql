/* create table RESULTADO (SEGMENTO char(200) , PERFIL smallint , PERIODO_PREV char(7) , TOTAL_UST int) */

select	SEGMENTO , substring(periodo_prev,6,2) + '/' + substring(periodo_prev,1,4) as PERIODO , SUM(TOTAL_UST) AS TOTAL_UST
from	DEMANDAS
where	Status <> 'Cancelado'
and		prazo_final between '21/05/2016 00:00:00' and '20/06/2016 23:59:59'
and		Unidade = 'CEDESSP'
group	by SEGMENTO , PERIODO_PREV
order	by PERIODO_PREV , SEGMENTO

select * from funcionarios order by nome

select nome,
case 
when Situacao = 1 Then
	'Ativo'
else
	'Inativo'
end 
as Situacao
from funcionarios


UPDATE	Funcionarios 
SET		Nome = 'ALEX APARECIDO RUIZ', funcao = 8 ,  Codigo_Secao = '2.0007.28.2.10.03.03.1.12' ,  Descricao_Secao = 'SP_PR_CEF_3122/2013' ,  Localizacao = 'Caixa - SP' ,  Horario_Escala_Trabalho = '09:00 às 13:00 / 14:00 às 18:00' ,  Data_Admissao = '01/10/2010' ,  Centro_Custo = '2.10.03.03.1.12' ,  Situacao = 0 
WHERE	Matricula = 0046020

select * from valores

insert into valores (CONTRATO,VALOR_PERFIL_1, VALOR_PERFIL_2 , VALOR_PERFIL_3 , VIGENCIA_INICIO , VIGENCIA_FIM) values ('CTMARG' , 123.43,93.62,0,'01/04/2015','31/12/2018')
insert into valores (CONTRATO,VALOR_PERFIL_1, VALOR_PERFIL_2 , VALOR_PERFIL_3 , VIGENCIA_INICIO , VIGENCIA_FIM) values ('CTMONSI' , 127.67,93.05,80.94,'01/04/2015','31/12/2018')


select * from funcionarios order by nome

SELECT	* FROM SERVICOS WHERE CONTRATO = 'CTMONSI'

select Nome from funcionarios where lider = '1'