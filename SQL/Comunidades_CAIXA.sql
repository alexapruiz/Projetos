use CAIXA

create table Comunidade_Usuarios
(COMUNIDADE varchar(100),
SQUAD varchar(50),
MATRICULA char(7),
PAPEL varchar(50))

drop table Comunidade_Usuarios
delete from Comunidade_Usuarios
select * from Comunidade_Usuarios

select COMUNIDADE, SQUAD, PAPEL, count(0) AS TOTAL from Comunidade_Usuarios group by COMUNIDADE, SQUAD, PAPEL order by COMUNIDADE, SQUAD, PAPEL

select COMUNIDADE, SQUAD, PAPEL, MATRICULA from Comunidade_Usuarios order by COMUNIDADE, SQUAD, PAPEL, MATRICULA

select	MATRICULA 
from	Comunidade_Usuarios 
where	COMUNIDADE = 'ARRECADAÇÃO, CONVÊNIOS E COBRANÇA' 
AND		SQUAD = 'AC10 - Cobrança 1'
AND		PAPEL = 'Time (Desenvolvedores)'

SELECT DISTINCT(COMUNIDADE) FROM Comunidade_Usuarios ORDER BY COMUNIDADE
SELECT DISTINCT(SQUAD) as SQUAD FROM Comunidade_Usuarios WHERE COMUNIDADE = 'ARRECADAÇÃO, CONVÊNIOS E COBRANÇA' ORDER BY SQUAD
SELECT DISTINCT(PAPEL) as PAPEL FROM Comunidade_Usuarios WHERE COMUNIDADE = 'ARRECADAÇÃO, CONVÊNIOS E COBRANÇA' AND SQUAD = 'AC10 - Cobrança 1' ORDER BY PAPEL
SELECT MATRICULA FROM Comunidade_Usuarios WHERE COMUNIDADE = 'ARRECADAÇÃO, CONVÊNIOS E COBRANÇA' AND SQUAD = 'AC10 - Cobrança 1' AND PAPEL = 'Time (Desenvolvedores)' ORDER BY MATRICULA


SELECT DISTINCT(PAPEL) as PAPEL FROM Comunidade_Usuarios
update Comunidade_Usuarios set PAPEL = 'agente_qualidade' where PAPEL = 'Agente de Qualidade'


SELECT DISTINCT(COMUNIDADE) FROM Comunidade_Usuarios order by Comunidade
SELECT * FROM Comunidade_Usuarios where matricula = 'C083132'


--Stored Procedures
exec COMUNIDADE_SEL_DISTINCT 'Comunidade Habitação' , 'ABH1 - ADMINISTRAÇÃO - SQUAD A' , 'dono_produto'


