select count(0) as Fumantes from Doenca where fumante = 1
select count(0) as Noa_Fumantes from Doenca where fumante = 0

select	* 
from	Doenca 
where	diabetes = 1
and		cardiopatia = 0


select	id, idade, peso, altura, (peso / ((altura / 100.00) * (altura / 100.00))) as IMC
from	Doenca
where	((peso / ((altura / 100.00) * (altura / 100.00))) < 16) and (peso < 40)
order by IMC

select * from Doenca where altura > 130 and peso < 40 order by peso