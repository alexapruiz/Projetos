-- Quantidade de registros
select 	count(0) as QuantidadeProcessos 
from 	trace_mdi_p3
where 	SPID = 17
and	(textdata is not null and reads is not null and CPU is not null)
and	(textdata not like '%sp%')

-- Transações com maior tempo de duração
select 	SPID , textdata as Comando , duration as Tempo, reads as leitura , writes as escrita, CPU
from 	trace_mdi_p3
where 	SPID = 17
and	(textdata is not null and reads is not null and CPU is not null)
and	(textdata not like '%sp%')
order	by duration desc

-- Quantidade de cada comando
select 	convert(char,textData) as Comando , count(0) as Total
from 	trace_mdi_p3
where 	SPID = 17
and	(textdata is not null and reads is not null and CPU is not null)
and	(textdata not like '%sp%')
Group	By convert(char,textData)
order	By total desc
