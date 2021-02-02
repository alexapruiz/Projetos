select * from jogadoresXpartidas

select * from partidas


select 	j.nome , avg(jp.qtdepontos / p.qtdjogos) as MediaPontos
from	jogadoresXPartidas jp , partidas p , jogadores j
where	jp.idpartida = p.idpartida
and	jp.idjogador = j.idjogador
group	by j.nome