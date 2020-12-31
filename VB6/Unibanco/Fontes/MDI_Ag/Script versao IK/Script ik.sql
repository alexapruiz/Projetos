drop table Parametro

GO

CREATE TABLE [dbo].[Parametro] (
	[DataProcessamento] 	[int] NOT NULL ,
	[Hm_Abertura] 		[smalldatetime] NOT NULL ,
	[Hm_Fechamento] 	[smalldatetime] NULL ,
	[AgenciaCentral] 	[numeric](4, 0) NOT NULL ,
	[AgenciaApresentante] 	[numeric](4, 0) NOT NULL ,
	[Tm_Pendente] 		[int] NOT NULL ,
	[Tm_Atualizacao] 	[int] NOT NULL ,
	[Dir_Dados] 		[varchar] (255) NOT NULL ,
	[Dir_Imagens] 		[varchar] (255) NOT NULL ,
	[Dir_Trabalho] 		[varchar] (255) NOT NULL ,
	[Versao] 		[smallint] NOT NULL
) ON [PRIMARY]

GO

INSERT INTO Parametro
       (DataProcessamento,
	Hm_Abertura,
	Hm_Fechamento,
	AgenciaCentral,
	AgenciaApresentante,
	Tm_Pendente,
	Tm_Atualizacao,
	Dir_Dados,
	Dir_Imagens,
	Dir_Trabalho,
	Versao
	)
VALUES
       (20001127, 		--DataProcessamento
	GETDATE(),		--Hm_Abertura
	GETDATE(),		--Hm_Fechamento
	9999,			--AgenciaCentral
	9999,			--AgenciaApresentante
	300,			--Tm_Pendente
	30,			--Tm_Atualizacao
	'c:\mdi_ag\dados',	--Dir_Dados
	'c:\mdi_ag\imagens',	--Dir_Imagens
	'c:\mdi_ag\trabalho\',	--Dir_Trabalho
	509			--Versao
	)
