create table Comunidade_Usuarios
(COMUNIDADE varchar(100),
SQUAD varchar(50),
MATRICULA char(7),
PAPEL varchar(50))

use CAIXA
drop table Comunidade_Usuarios
delete from Comunidade_Usuarios
select * from Comunidade_Usuarios

select COMUNIDADE, SQUAD, PAPEL, count(0) AS TOTAL from Comunidade_Usuarios group by COMUNIDADE, SQUAD, PAPEL order by COMUNIDADE, SQUAD, PAPEL

select COMUNIDADE, SQUAD, PAPEL, MATRICULA from Comunidade_Usuarios order by COMUNIDADE, SQUAD, PAPEL, MATRICULA

select	MATRICULA 
from	Comunidade_Usuarios 
where	COMUNIDADE = 'ARRECADA��O, CONV�NIOS E COBRAN�A' 
AND		SQUAD = 'AC10 - Cobran�a 1'
AND		PAPEL = 'Time (Desenvolvedores)'

SELECT DISTINCT(COMUNIDADE) FROM Comunidade_Usuarios ORDER BY COMUNIDADE
SELECT DISTINCT(SQUAD) as SQUAD FROM Comunidade_Usuarios WHERE COMUNIDADE = 'ARRECADA��O, CONV�NIOS E COBRAN�A' ORDER BY SQUAD
SELECT DISTINCT(PAPEL) as PAPEL FROM Comunidade_Usuarios WHERE COMUNIDADE = 'ARRECADA��O, CONV�NIOS E COBRAN�A' AND SQUAD = 'AC10 - Cobran�a 1' ORDER BY PAPEL
SELECT MATRICULA FROM Comunidade_Usuarios WHERE COMUNIDADE = 'ARRECADA��O, CONV�NIOS E COBRAN�A' AND SQUAD = 'AC10 - Cobran�a 1' AND PAPEL = 'Time (Desenvolvedores)' ORDER BY MATRICULA


SELECT DISTINCT(PAPEL) as PAPEL FROM Comunidade_Usuarios
update Comunidade_Usuarios set PAPEL = 'agente_qualidade' where PAPEL = 'Agente de Qualidade'


SELECT distinct(comunidade) FROM Comunidade_Usuarios where squad = 'CI03 - A��es on-line - Sustenta��o Frontoffice'
SELECT * FROM Comunidade_Usuarios where comunidade = 'Comunidade'
