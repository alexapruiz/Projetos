use ESTUDOS
select * from Cardio

select	count(0) as total , cardiopatia
from	cardio
group by cardiopatia

select	*
from	cardio
where	pressao_max > 150
and		cardiopatia = 1