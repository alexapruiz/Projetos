use CAIXA
delete from import13
select * from import13
delete from arquivo3
select * from Demandas_BRQ
select * from Demandas_BRQ order by ID
select * from Demandas_BRQ for xml PATH

CREATE TABLE [dbo].[Demandas_BRQ](
	[ID] [int] NOT NULL,
	[RESUMO] [nchar](150) NULL,
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
 CONSTRAINT [PK_Demandas_BRQ] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO

insert into Servicos (SERVICO,COMPLEXIDADE_BAIXA,COMPLEXIDADE_MEDIA,COMPLEXIDADE_ALTA) values ('Apoio � Solu��o de Problemas Relacionados �s Ferramentas',4,8,12)
insert into Servicos (SERVICO,COMPLEXIDADE_BAIXA,COMPLEXIDADE_MEDIA,COMPLEXIDADE_ALTA) values ('Cria��o/Manuten��o de �rea de projeto',2,4,8)
insert into Servicos (SERVICO,COMPLEXIDADE_BAIXA,COMPLEXIDADE_MEDIA,COMPLEXIDADE_ALTA) values ('Mentoring',2,4,8)
insert into Servicos (SERVICO,COMPLEXIDADE_BAIXA,COMPLEXIDADE_MEDIA,COMPLEXIDADE_ALTA) values ('Manuten��o em permiss�es/perfis de usu�rio',1,2,4)
insert into Servicos (SERVICO,COMPLEXIDADE_BAIXA,COMPLEXIDADE_MEDIA,COMPLEXIDADE_ALTA) values ('Cria��o de �rea de projeto integrada/associa��o de �reas de projeto',1,2,4)
insert into Servicos (SERVICO,COMPLEXIDADE_BAIXA,COMPLEXIDADE_MEDIA,COMPLEXIDADE_ALTA) values ('Manuten��o em Indicador / Relat�rio',8,16,32)
insert into Servicos (SERVICO,COMPLEXIDADE_BAIXA,COMPLEXIDADE_MEDIA,COMPLEXIDADE_ALTA) values ('Manuten��o em Itens de Configura��o (ex:Retirada de check-out)',2,2,2)
insert into Servicos (SERVICO,COMPLEXIDADE_BAIXA,COMPLEXIDADE_MEDIA,COMPLEXIDADE_ALTA) values ('Manuten��o em pain�is',2,4,8)
insert into Servicos (SERVICO,COMPLEXIDADE_BAIXA,COMPLEXIDADE_MEDIA,COMPLEXIDADE_ALTA) values ('Cria��o e Configura��o de Projeto',4,4,4)
insert into Servicos (SERVICO,COMPLEXIDADE_BAIXA,COMPLEXIDADE_MEDIA,COMPLEXIDADE_ALTA) values ('Cria��o / Manuten��o de Atributos',1,2,4)
insert into Servicos (SERVICO,COMPLEXIDADE_BAIXA,COMPLEXIDADE_MEDIA,COMPLEXIDADE_ALTA) values ('G1 - Capacita��o',0,0,0)
insert into Servicos (SERVICO,COMPLEXIDADE_BAIXA,COMPLEXIDADE_MEDIA,COMPLEXIDADE_ALTA) values ('Cria��o / Manuten��o de View',2,2,2)
insert into Servicos (SERVICO,COMPLEXIDADE_BAIXA,COMPLEXIDADE_MEDIA,COMPLEXIDADE_ALTA) values ('Cria��o / Manuten��o de VOB',8,8,8)
insert into Servicos (SERVICO,COMPLEXIDADE_BAIXA,COMPLEXIDADE_MEDIA,COMPLEXIDADE_ALTA) values ('Cria��o / Manuten��o de Artefatos',2,4,8)

select	D.ID, D.QTDE, D.COMPLEXIDADE,
		S.COMPLEXIDADE_BAIXA, S.COMPLEXIDADE_MEDIA,S.COMPLEXIDADE_ALTA
from	Demandas_BRQ D , Servicos S
where	D.SERVICO = S.SERVICO

select * from Demandas_BRQ

select	PERIODO, sum(UST) as USTs from	Demandas_BRQ where	PERIODO is not null group by PERIODO ORDER BY PERIODO

select * from Demandas_BRQ WHERE id = 12524545
select * from servicos




USE [CAIXA]
GO

/****** Object:  Table [dbo].[Demandas_BRQ]    Script Date: 07/12/2020 19:54:40 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE TABLE [dbo].[Demandas_BRQ](
	[ID] [int] NOT NULL,
	[RESUMO] [nchar](150) NULL,
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
	[PERIODO] [nchar](10) NULL
)