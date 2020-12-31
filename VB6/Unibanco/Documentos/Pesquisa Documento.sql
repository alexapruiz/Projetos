Set noCount on
select * from capa where capa = 6005647151011

declare @DataProcessamento 	int
declare @Idcapa 		int

select @DataProcessamento = 20030818
select @IdCapa = 8

select 	IdDocto , TipoDocto , Ocorrencia , Leitura , Status , Frente, Vinculo , NSU , Terminal , Valor , RetornoTransacao
from 	documento 
where 	dataprocessamento = @DataProcessamento and idcapa = @Idcapa order by vinculo

Select 	Log.IdCapa , Log.IdDocto , Log.Login , Log.Data , Acao.Descricao
From 	Log , Acao
Where	Log.Acao = Acao.Acao
And	Log.DataProcessamento = @DataProcessamento
And	Log.IdCapa = @Idcapa
Order	by Log.Data

select * from logerro where DataProcessamento = @DataProcessamento order by Data

select * from caixa where DataProcessamento = @DataProcessamento

select * from MDI_UBB..LogExtra where DataProcessamento = @DataProcessamento

Set noCount off