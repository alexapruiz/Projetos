use CAIXA
delete from import13
select * from import13
delete from arquivo3
select * from Demandas_BRQ
select * from Demandas_BRQ order by ID
select * from Demandas_BRQ for xml PATH

select	D.ID, D.QTDE, D.COMPLEXIDADE,
		S.COMPLEXIDADE_BAIXA, S.COMPLEXIDADE_MEDIA,S.COMPLEXIDADE_ALTA
from	Demandas_BRQ D , Servicos S
where	D.SERVICO = S.SERVICO

select	PERIODO, sum(UST) as USTs from	Demandas_BRQ where	PERIODO is not null group by PERIODO ORDER BY PERIODO

select	sum(UST) as USTs, PERIODO , FERRAMENTA
from	Demandas_BRQ
where	GRUPO = 'Grupo 2'
group	by PERIODO , FERRAMENTA
order	by PERIODO , FERRAMENTA , USTs

select PRAZO_FINAL, PERIODO from Demandas_BRQ order by PRAZO_FINAL

select * from Demandas_BRQ
select * from Servicos


DROP TABLE Demandas_BRQ
go
CREATE TABLE [dbo].[Demandas_BRQ](
	[ID] [int] NOT NULL,
	[RESUMO] [nchar](250) NULL,
	[STATUS] [nchar](30) NULL,
	[QTDE] [smallint] NULL,
	[COMPLEXIDADE] [nchar](5) NULL,
	[DATA_CRIACAO] [datetime] NULL,
	[PRAZO_FINAL] [datetime] NULL,
	[SOLICITANTE] [nchar](30) NULL,
	[PREPOSTO] [nchar](50) NULL,
	[SERVICO] [nchar](100) NULL,
	[UST] [smallint] NULL,
	[GRUPO] [nchar](20) NULL,
	[PERIODO] [nchar](10) NULL,
	[FERRAMENTA] [nchar](10) NULL,
 CONSTRAINT [PK_Demandas_BRQ] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO

select * from Demandas_BRQ where FERRAMENTA = 'None'

truncate table Demandas_BRQ
