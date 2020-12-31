ALTER TABLE [dbo].[Log] DROP CONSTRAINT FK_LOG_REF_15195_ACAO
GO

ALTER TABLE [dbo].[Capa] DROP CONSTRAINT FK_CAPA_REF_21052_OCORRENC
GO

ALTER TABLE [dbo].[Capa] DROP CONSTRAINT FK_CAPA_REF_8846_STATUSCA
GO

ALTER TABLE [dbo].[Documento] DROP CONSTRAINT FK_DOCUMENT_REF_8858_STATUSDO
GO

ALTER TABLE [dbo].[Lote] DROP CONSTRAINT FK_LOTE_REF_8852_STATUSLO
GO

ALTER TABLE [dbo].[Documento] DROP CONSTRAINT FK_DOCUMENT_REF_4541_TIPODOCT
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[tI_Capa]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[tI_Capa]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[tD_Capa]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[tD_Capa]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[tI_Documento]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[tI_Documento]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[tD_Lote]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
drop trigger [dbo].[tD_Lote]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_AlteraTipoDocto]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_AlteraTipoDocto]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_AtualizaCapa]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_AtualizaCapa]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_AtualizaDuplicidadeCapa]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_AtualizaDuplicidadeCapa]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_AtualizaStatusLote]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_AtualizaStatusLote]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_CapturaCapa]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_CapturaCapa]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_CapturaDocumento]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_CapturaDocumento]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_ChecarCapaEnvelope]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_ChecarCapaEnvelope]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_GetAgenciasCapa]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_GetAgenciasCapa]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_GetDocumentoContQualidade]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_GetDocumentoContQualidade]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_GetEstatistica]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_GetEstatistica]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_GetLoteContQualidade]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_GetLoteContQualidade]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_GetMaloteExpedicao_oc]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_GetMaloteExpedicao_oc]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_GetMudaStatus]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_GetMudaStatus]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_GetRecuperaStatus]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_GetRecuperaStatus]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_GetTodasCapas]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_GetTodasCapas]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_GetTodasCapas_Oc]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_GetTodasCapas_Oc]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_GetTotalCapa]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_GetTotalCapa]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_GetTotalDocumento]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_GetTotalDocumento]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_GetUltimaImagemLote]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_GetUltimaImagemLote]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_InsereCapa]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_InsereCapa]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_InsereLog]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_InsereLog]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_InsereLote]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_InsereLote]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_LimpaMovimento]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_LimpaMovimento]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_ObtemCapasLote]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_ObtemCapasLote]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_ObtemDocumentosCapa]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_ObtemDocumentosCapa]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_ObtemLog]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_ObtemLog]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_ObtemLotesExportacao]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_ObtemLotesExportacao]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_RemoveCapa]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_RemoveCapa]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_RemoveDocumento]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_RemoveDocumento]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_RemoveLote]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_RemoveLote]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_SplitCapa]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_SplitCapa]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_SplitCapaAnterior]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_SplitCapaAnterior]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_TotalizaAgencia]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_TotalizaAgencia]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_VerificaCapaExportar]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_VerificaCapaExportar]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_VerificaCapaRecepcionada]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_VerificaCapaRecepcionada]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_VerificaLoteDisponivel]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_VerificaLoteDisponivel]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_VerificaLotePendente]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_VerificaLotePendente]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_AtualizaAgencia]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_AtualizaAgencia]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_EncerraMovimento]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_EncerraMovimento]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_GetAllTipoDocumento]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_GetAllTipoDocumento]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_GetTodasOcorrencia]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_GetTodasOcorrencia]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_InsereAgencia]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_InsereAgencia]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_InsereLogErro]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_InsereLogErro]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_LerParametro]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_LerParametro]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_ObtemAgencias]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_ObtemAgencias]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[Capa]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Capa]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[Documento]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Documento]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[Log]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Log]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[Lote]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Lote]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[Acao]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Acao]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[Agencia]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Agencia]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[LogErro]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[LogErro]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[Ocorrencia]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Ocorrencia]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[Parametro]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Parametro]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[StatusCapa]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[StatusCapa]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[StatusDocumento]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[StatusDocumento]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[StatusLote]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[StatusLote]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[TipoDocto]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[TipoDocto]
GO

if not exists (select * from sysusers where name = N'mdi' and uid < 16382)
	EXEC sp_grantdbaccess N'mdi', N'mdi'
GO

exec sp_addrolemember N'db_owner', N'mdi'
GO

CREATE TABLE [dbo].[Acao] (
	[Acao] [tinyint] NOT NULL ,
	[Descricao] [varchar] (50) NOT NULL 
)
GO

CREATE TABLE [dbo].[Agencia] (
	[Agencia] [smallint] NOT NULL ,
	[Lacre] [decimal](8, 0) NOT NULL ,
	[QtdInformada] [int] NOT NULL ,
	[QtdGravada] [int] NULL ,
	[Identificador] [char] (10) NULL ,
	[HoraChegada] [char] (5) NULL ,
	[HoraCadastrada] [char] (5) NOT NULL ,
	[idEnv_Mal] [char] (1) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[LogErro] (
	[Data] [datetime] NOT NULL ,
	[Erro] [int] NULL ,
	[Descricao] [varchar] (255) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Ocorrencia] (
	[Ocorrencia] [decimal](5, 0) NOT NULL ,
	[Descricao] [varchar] (82) NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Parametro] (
	[DataProcessamento] [int] NOT NULL ,
	[Hm_Abertura] [smalldatetime] NOT NULL ,
	[Hm_Fechamento] [smalldatetime] NULL ,
	[AgenciaCentral] [numeric](4, 0) NOT NULL ,
	[AgenciaApresentante] [numeric](4, 0) NOT NULL ,
	[Tm_Pendente] [int] NOT NULL ,
	[Tm_Atualizacao] [int] NOT NULL ,
	[Dir_Dados] [varchar] (255) NOT NULL ,
	[Dir_Imagens] [varchar] (255) NOT NULL ,
	[Dir_Trabalho] [varchar] (255) NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[StatusCapa] (
	[Status] [char] (1) NOT NULL ,
	[Descricao] [varchar] (50) NOT NULL 
)
GO

CREATE TABLE [dbo].[StatusDocumento] (
	[Status] [char] (1) NOT NULL ,
	[Descricao] [varchar] (50) NOT NULL 
)
GO

CREATE TABLE [dbo].[StatusLote] (
	[Status] [char] (1) NOT NULL ,
	[Descricao] [varchar] (50) NOT NULL 
)
GO

CREATE TABLE [dbo].[TipoDocto] (
	[TipoDocto] [smallint] NOT NULL ,
	[Nome] [varchar] (30) NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Capa] (
	[DataProcessamento] [int] NOT NULL ,
	[IdCapa] [int] IDENTITY (1, 1) NOT NULL ,
	[IdLote] [int] NOT NULL ,
	[idEnv_Mal] [char] (1) NOT NULL ,
	[Capa] [numeric](18, 0) NOT NULL ,
	[Num_Malote] [numeric](11, 0) NULL ,
	[AgOrig] [smallint] NOT NULL ,
	[Status] [char] (1) NOT NULL ,
	[DataCriacao] [smalldatetime] NOT NULL ,
	[Ocorrencia] [decimal](5, 0) NULL ,
	[Duplicidade] [numeric](1, 0) NULL ,
	[HoraAtual] [datetime] NULL 
)
GO

CREATE TABLE [dbo].[Documento] (
	[DataProcessamento] [int] NOT NULL ,
	[IdDocto] [int] IDENTITY (1, 1) NOT NULL ,
	[IdCapa] [int] NOT NULL ,
	[OrdemCaptura] [smallint] NOT NULL ,
	[TipoDocto] [smallint] NOT NULL ,
	[Leitura] [varchar] (48) NULL ,
	[Frente] [varchar] (20) NOT NULL ,
	[Verso] [varchar] (20) NOT NULL ,
	[Status] [char] (1) NOT NULL ,
	[Ordem] [char] (1) NULL 
)
GO

CREATE TABLE [dbo].[Log] (
	[DataProcessamento] [int] NOT NULL ,
	[IdCapa] [int] NULL ,
	[IdDocto] [int] NULL ,
	[Data] [datetime] NOT NULL ,
	[Login] [varchar] (10) NOT NULL ,
	[Acao] [tinyint] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Lote] (
	[DataProcessamento] [int] NOT NULL ,
	[IdLote] [int] NOT NULL ,
	[Status] [char] (1) NOT NULL ,
	[Prioridade] [smallint] NOT NULL ,
	[HoraAtual] [datetime] NULL 
)
GO

ALTER TABLE [dbo].[Acao] WITH NOCHECK ADD 
	CONSTRAINT [PK_Acao] PRIMARY KEY  CLUSTERED 
	(
		[Acao]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[StatusCapa] WITH NOCHECK ADD 
	CONSTRAINT [PK_STATUSCAPA] PRIMARY KEY  CLUSTERED 
	(
		[Status]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[StatusDocumento] WITH NOCHECK ADD 
	CONSTRAINT [PK_STATUSDOCUMENTO] PRIMARY KEY  CLUSTERED 
	(
		[Status]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[StatusLote] WITH NOCHECK ADD 
	CONSTRAINT [PK_STATUSLOTE] PRIMARY KEY  CLUSTERED 
	(
		[Status]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Capa] WITH NOCHECK ADD 
	CONSTRAINT [PK_Capa] PRIMARY KEY  CLUSTERED 
	(
		[DataProcessamento],
		[IdCapa]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Documento] WITH NOCHECK ADD 
	CONSTRAINT [PK_Documento] PRIMARY KEY  CLUSTERED 
	(
		[DataProcessamento],
		[IdDocto]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Lote] WITH NOCHECK ADD 
	CONSTRAINT [PK_Lote] PRIMARY KEY  CLUSTERED 
	(
		[DataProcessamento],
		[IdLote]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Ocorrencia] WITH NOCHECK ADD 
	CONSTRAINT [PK_Ocorrencia] PRIMARY KEY  NONCLUSTERED 
	(
		[Ocorrencia]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[TipoDocto] WITH NOCHECK ADD 
	CONSTRAINT [PK_TipoDocto] PRIMARY KEY  NONCLUSTERED 
	(
		[TipoDocto]
	)  ON [PRIMARY] 
GO

 CREATE  INDEX [IndDataIdCapa] ON [dbo].[Documento]([DataProcessamento], [IdCapa]) ON [PRIMARY]
GO

GRANT  SELECT ,  INSERT ,  DELETE ,  UPDATE  ON [dbo].[Acao]  TO [mdi]
GO

GRANT  SELECT ,  INSERT ,  DELETE ,  UPDATE  ON [dbo].[Agencia]  TO [mdi]
GO

GRANT  SELECT ,  INSERT ,  DELETE ,  UPDATE  ON [dbo].[LogErro]  TO [mdi]
GO

GRANT  SELECT ,  INSERT ,  DELETE ,  UPDATE  ON [dbo].[Ocorrencia]  TO [mdi]
GO

GRANT  SELECT ,  INSERT ,  DELETE ,  UPDATE  ON [dbo].[Parametro]  TO [mdi]
GO

GRANT  SELECT ,  INSERT ,  DELETE ,  UPDATE  ON [dbo].[StatusCapa]  TO [mdi]
GO

GRANT  SELECT ,  INSERT ,  DELETE ,  UPDATE  ON [dbo].[StatusDocumento]  TO [mdi]
GO

GRANT  SELECT ,  INSERT ,  DELETE ,  UPDATE  ON [dbo].[StatusLote]  TO [mdi]
GO

GRANT  SELECT ,  INSERT ,  DELETE ,  UPDATE  ON [dbo].[TipoDocto]  TO [mdi]
GO

GRANT  SELECT ,  INSERT ,  DELETE ,  UPDATE  ON [dbo].[Capa]  TO [mdi]
GO

GRANT  SELECT ,  INSERT ,  DELETE ,  UPDATE  ON [dbo].[Documento]  TO [mdi]
GO

GRANT  SELECT ,  INSERT ,  DELETE ,  UPDATE  ON [dbo].[Log]  TO [mdi]
GO

GRANT  SELECT ,  INSERT ,  DELETE ,  UPDATE  ON [dbo].[Lote]  TO [mdi]
GO

ALTER TABLE [dbo].[Capa] ADD 
	CONSTRAINT [FK_CAPA_REF_21052_OCORRENC] FOREIGN KEY 
	(
		[Ocorrencia]
	) REFERENCES [dbo].[Ocorrencia] (
		[Ocorrencia]
	),
	CONSTRAINT [FK_CAPA_REF_8846_STATUSCA] FOREIGN KEY 
	(
		[Status]
	) REFERENCES [dbo].[StatusCapa] (
		[Status]
	)
GO

ALTER TABLE [dbo].[Documento] ADD 
	CONSTRAINT [FK_DOCUMENT_REF_4541_TIPODOCT] FOREIGN KEY 
	(
		[TipoDocto]
	) REFERENCES [dbo].[TipoDocto] (
		[TipoDocto]
	),
	CONSTRAINT [FK_DOCUMENT_REF_8858_STATUSDO] FOREIGN KEY 
	(
		[Status]
	) REFERENCES [dbo].[StatusDocumento] (
		[Status]
	)
GO

ALTER TABLE [dbo].[Log] ADD 
	CONSTRAINT [FK_LOG_REF_15195_ACAO] FOREIGN KEY 
	(
		[Acao]
	) REFERENCES [dbo].[Acao] (
		[Acao]
	)
GO

ALTER TABLE [dbo].[Lote] ADD 
	CONSTRAINT [FK_LOTE_REF_8852_STATUSLO] FOREIGN KEY 
	(
		[Status]
	) REFERENCES [dbo].[StatusLote] (
		[Status]
	)
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO



/****** Object:  Stored Procedure dbo.MDIAG_AtualizaAgencia    Script Date: 17/11/00 11:22:09 ******/
CREATE PROCEDURE MDIAG_AtualizaAgencia

	@pAgenciaApresentante		Numeric(5)

AS

	UPDATE Parametro
		Set AgenciaApresentante = @pAgenciaApresentante

	Return @@Error



GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO


CREATE PROCEDURE MDIAG_EncerraMovimento
	@pDataProcessamento	Int
AS

Declare @Erro	Int

	Select @Erro = 0

	/*----------------------------------------------------------------
	   Atualiza hora de fechamento no parametro
	----------------------------------------------------------------*/
	UPDATE Parametro
		SET Hm_Fechamento	= GETDATE()
	WHERE DataProcessamento 	= @pDataProcessamento


	Select @Erro = @@Error

	Return @Erro


GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO


CREATE Procedure MDIAG_GetAllTipoDocumento

As

Select TipoDocto, Nome
From TipoDocto
Order By TipoDocto




GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO



/****** Object:  Stored Procedure dbo.MDIAG_GetTodasOcorrencia    Script Date: 17/11/00 11:22:09 ******/
CREATE PROCEDURE MDIAG_GetTodasOcorrencia

As

	Select	Ocorrencia,
		Descricao
	  From 	Ocorrencia 
	 Order 	by  Ocorrencia asc




GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO



CREATE PROCEDURE MDIAG_InsereAgencia
	@Agencia		SmallInt,
	@Lacre			Decimal(8),
	@QtdInformada		Int,
	@QtdGravada		Int,
	@Identificador		Char(10),
	@HoraChegada		Char(5),
	@HoraCadastrada		Char(5),
	@IdEnvMal		Char(1)

As

	SET NOCOUNT ON
	Begin Transaction
	
	/****************************
	 * Insere na tabela Agencia *
	 ****************************/
	INSERT INTO Agencia 
		(Agencia, 
		 Lacre, 
		 QtdInformada,
		 QtdGravada,
		 Identificador,
		 HoraChegada,
		 HoraCadastrada,
		 IdEnv_Mal)
	VALUES
		(@Agencia,
		 @Lacre,
		 @QtdInformada,
		 @QtdGravada,
		 @Identificador,
		 @HoraChegada,
		 @HoraCadastrada,
		 @IdEnvMal)

	If @@Error = 0 Begin
		Commit Transaction
		Return(0)
	End
	Else Begin
		Rollback Transaction
		Return(@@Error)
	End





GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO



/****** Object:  Stored Procedure dbo.MDIAG_InsereLogErro    Script Date: 17/11/00 11:22:09 ******/
CREATE Procedure MDIAG_InsereLogErro
  @Erro            		int,
  @Descricao       	varchar(255)
As
Begin Transaction
  INSERT INTO LogErro(
                       Data,
                       Erro,
                       Descricao )
              values(
                       GetDate(),
                      @Erro,
                      @Descricao )
  If @@Error = 0 Begin
    Commit Transaction
  End
  Else Begin
    Rollback Transaction
  End





GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO



/****** Object:  Stored Procedure dbo.MDIAG_LerParametro    Script Date: 17/11/00 11:22:09 ******/
CREATE PROCEDURE MDIAG_LerParametro
AS
	Select	DataProcessamento,
		Hm_Abertura,
		Hm_Fechamento,
		AgenciaCentral,
		AgenciaApresentante,
		TM_Pendente,
		TM_Atualizacao,
		Dir_Dados,
		Dir_Imagens,
		Dir_Trabalho
	  From	Parametro




GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  ON    SET ANSI_NULLS  ON 
GO


CREATE PROCEDURE MDIAG_ObtemAgencias

AS


	SELECT 	Agencia,
		Lacre,
		QtdInformada,
		HoraCadastrada,
		IdEnv_Mal
	  FROM	Agencia



GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO


CREATE Proc MDIAG_AlteraTipoDocto

	@pDataProcessamento	Int,
	@pIdDocto		Int,
	@pTipoDocto		Int

As

Declare @Erro           int,
	@LinhasAfetadas int


	/*-----------------
	  Starta Transação
	-----------------*/

	Begin Transaction AlteraTipoDocto


	-- Altera Documento

	UPDATE Documento SET
		TipoDocto		= @pTipoDocto
	 WHERE DataProcessamento	= @pDataProcessamento
	   AND IdDocto	        	= @pIdDocto



	SELECT 	@Erro           = @@Error,
		@LinhasAfetadas = @@Rowcount

	-- Verifica Codigo de Erro
	If @LinhasAfetadas <> 1 or @Erro <> 0 Begin
		Rollback Transaction AlteraTipoDocto
		Return(1)
	End


	Commit Transaction

	Return(0)









GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO



/*
     	Esta procedure atualiza Agencia, Capa de Envelope/Malote e N£mero de Malote para tabela capa 
	j  cadastrada, vincula a Capa de Envelope e C¢digo	 do CMC7 na tabela Documento, apenas se 
	fornecido o n£mero do documento (IdDocto).
    	Se informado o N£mero do Malote Empresa ser  atualizado o CMC7 no campo Leitura senƒo
	ser  atualizado o n£mero da Capa de Envelope na Tabela Documento.


     Retorno:  (0)-Sucesso
               (1)-Erro
*/

CREATE Procedure MDIAG_AtualizaCapa
	@Data		Int,
	@IdCapa		Int,
	@Capa 		Numeric(18),
	@AgOrig		SmallInt,
	@IdDocto	Int 		= 0,
	@Num_Malote	Numeric(11) 	= 0,
	@CMC7		Char(30) 	= NULL
As

	Begin Transaction

	UPDATE 	Capa SET
		AgOrig 			= @AgOrig,
		Capa			= @Capa,
		Num_Malote		= @Num_Malote,
		Ocorrencia		= null,
		Duplicidade		= 0,
		idEnv_Mal		= CASE @Num_Malote WHEN 0 THEN "E" ELSE "M" END
	 WHERE 	DataProcessamento 	= @Data
	   AND	IdCapa			= @IdCapa

	If @IdDocto > 0 Begin
		UPDATE	Documento SET
		 	Leitura 		= CASE @Num_Malote WHEN 0 THEN (Right("00000000" + Convert(Varchar(8), @Capa),8)) ELSE @CMC7 END ,
			TipoDocto		= 1	---* Envelope/Malote *---
		 WHERE	DataProcessamento 	= @Data
		   AND	IdCapa			= @IdCapa		
		   AND	IdDocto			= @IdDocto

	End

	If @@Error = 0 Begin
		Commit Transaction
		Return(0)
	End
	Else Begin
		Rollback Transaction
		Return(1)
	End













GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO



CREATE PROCEDURE MDIAG_AtualizaDuplicidadeCapa
	@DataProcessamento              int,
	@IdCapa                         int,
	@Status 			char(1),
	@Duplicidade			numeric(1)
as

	Begin Transaction
	--------------------------------------------------------------------------------------------------
	-- Atualiza Dados de Capa
	UPDATE Capa SET
		Status		= @Status,
		Duplicidade	= @Duplicidade,
		Ocorrencia	= CASE WHEN @Duplicidade = 0 THEN Null Else 998 END
	 WHERE  DataProcessamento	= @DataProcessamento
	   AND	IdCapa			= @IdCapa
	--------------------------------------------------------------------------------------------------
	-- Verifica Codigo de Erro
	If @@Error = 0 Begin
		Commit Transaction
		Return(0)
	End
	Else Begin
		Rollback Transaction
		Return(1)
	End



GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO



/****** Object:  Stored Procedure dbo.MDIAG_AtualizaStatusLote    Script Date: 17/11/00 11:22:09 ******/
CREATE Procedure MDIAG_AtualizaStatusLote
	 @DataProcessamento		int
	,@IdLote             		int
	,@Status             		char(1)

As

Declare @Erro           int
Declare @LinhasAfetadas int

Begin Transaction

	Update Lote Set
		Status 			= @Status,
		HoraAtual 		= GetDate()
	 Where 	DataProcessamento	= @DataProcessamento
		And IdLote              = @IdLote

	Select 	@Erro 		= @@Error,
		@LinhasAfetadas = @@RowCount

	If @LinhasAfetadas <> 1 or @Erro <> 0 
		Begin 	
			RollBack Transaction
   			Return(1)	
		End

Commit Transaction
Return(0)


GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO



/****** Object:  Stored Procedure dbo.MDIAG_CapturaCapa    Script Date: 17/11/00 11:22:10 ******/
CREATE Proc MDIAG_CapturaCapa
	@DataProcessamento		Int,
	@IdLote 			Int,
	@idEnv_Mal			Char(1),
	@Capa				Numeric(18),
	@AgOrig				Smallint,
	@IdCapa 			Int Output

as

Declare @Status    char(1)

	-- Verifica se capa virtual (9)
        if @Capa = 9 Begin
		------------------------------------------------------------------------------------------
		Insert Capa (DataProcessamento,	IdLote, idEnv_Mal, Capa, Num_Malote, AgOrig, 
				Status, DataCriacao, Duplicidade )
	              Values(  @DataProcessamento, @IdLote, @idEnv_Mal, @Capa, 0, @AgOrig, 
				'1', GetDate(), 0 )
		if @@Error = 0 Begin
			Select @IdCapa = @@Identity
		End
		if @@Error = 0 Begin
			Return(0)
		End
		Else Begin
			Return(1)
		End
        End

	-- Verifica se nao existe a Capa 
	if not exists ( Select 1
			  From Capa 
			 Where DataProcessamento 	= @DataProcessamento
			   And Capa 			= @Capa
			   And Status not in 		  ('F','D','P')) Begin
		------------------------------------------------------------------------------------------
		Insert Capa ( 	DataProcessamento, IdLote, idEnv_Mal, Capa, Num_Malote, AgOrig, 
				Status, DataCriacao, Duplicidade )
	              Values(  @DataProcessamento, @IdLote, @idEnv_Mal, @Capa, 0, @AgOrig, 
				'1', GetDate(), 0 )
		if @@Error = 0 Begin
			Select @IdCapa = @@Identity
		End
		if @@Error = 0 Begin
			Return(0)
		End
		Else Begin
			Return(1)
		End
	End
	Else Begin

		-- Ja existe capa com mesmo DataProc + Capa
		------------------------------------------------------------------------------------------
		Select	@IdCapa			= IdCapa,
			@Status			= Status
		  From	Capa 
		 Where	DataProcessamento 	= @DataProcessamento
		   And	Capa 			= @Capa
		   And	Status not in 		('F','D','P')

		-- Se o Status for Capa Cadastrada entao atualiza Capa
		if @Status = '0' Begin
			Update Capa Set 
				IdLote 			= @IdLote,
				Status 			= '1' 
		   	 Where DataProcessamento	= @DataProcessamento
			   And IdCapa 			= @IdCapa
		End
		Else Begin
			Insert Capa ( DataProcessamento, IdLote, idEnv_Mal, Capa, Num_Malote, AgOrig, 
					Status, DataCriacao, Duplicidade )
	                Values      ( @DataProcessamento, @idLote, @idEnv_Mal, @Capa, 0, @AgOrig, 
				'1', GetDate(), 1 )

			if @@Error = 0 Begin
				Select @IdCapa = @@Identity
			End
		End
		if @@Error = 0 Begin
			Return(0)
		End
		Else Begin
			Return(1)
		End
	End



GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO



/****** Object:  Stored Procedure dbo.MDIAG_CapturaDocumento    Script Date: 17/11/00 11:22:10 ******/
CREATE Procedure MDIAG_CapturaDocumento
	@DataProcessamento		Int,
	@IdCapa 			Int,
	@TipoDocto			Smallint,
	@Leitura			Varchar(48),
	@Frente				Varchar(20),
	@Verso				Varchar(20),
	@Ordem				Char(1),
	@OrdemCaptura			Int,
	@IdDocto			Int          Output

AS

	Declare @Error                  int

	/*--------------------------------------------------------------
	  Simplesmente insere na tabela documento com a ordem de captura
	--------------------------------------------------------------*/
	Insert into Documento (	DataProcessamento, 	IdCapa, 	TipoDocto, 	Leitura,
				Frente,			Verso,		Status,		Ordem,
				OrdemCaptura)

                     values (	@DataProcessamento, 	@IdCapa, 	@TipoDocto, 	@Leitura, 
				@Frente, 		@Verso, 	'0',		@Ordem,
			 	@OrdemCaptura)


	Select @Error = @@Error, @IdDocto = @@Identity

	If @Error = 0 Begin
		Return(0)
	End
	Else Begin
		Return(1)
	End






GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO





---------------------------------------------------------------------------
---     Verifica se existe N£mero de Capa j  cadastrada (Duplicidade)	---
---									---
---	Pesquisa por Agˆncia ou apenas por N£mero de Capa passando	---
---	o parƒmetro AgOrig						---
---									---
---	Retorno: (0) - Sucesso						---
---		 (1) - Erro						---
---------------------------------------------------------------------------
CREATE PROCEDURE MDIAG_ChecarCapaEnvelope
	@DataProcessamento	Int,
	@AgOrig			SmallInt = Null,
	@Capa			Numeric(18),
	@Registros		Int Output,
	@IdCapa			Int
As	

	SELECT	@Registros = Count(*)
	  FROM	Capa (NOLOCK)
	 WHERE	DataProcessamento	= @DataProcessamento 
	   AND	Capa			= @Capa
	   AND	IdCapa			<> @IdCapa
	   AND	Convert(Char(10),AgOrig) like Case @AgOrig WHEN Null THEN "%" ELSE Convert(Char(10),@AgOrig) END
	   AND	Status			not in ("0","1","D","F","P")	--- Verifica duplicidade somente para capa complementada ou adiante ---


	If @@Error <> 0
		Return(1)

	Return(0)







GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO



/****** Object:  Stored Procedure dbo.MDIAG_GetAgenciasCapa    Script Date: 17/11/00 11:22:10 ******/

/****** Object:  Stored Procedure dbo.GetAgenciasCapa    Script Date: 03/11/00 18:02:43 ******/
/****** Object:  Stored Procedure dbo.GetAgenciasCapa    Script Date: 14/09/00 13:02:06 ******/
/****** Object:  Stored Procedure dbo.GetAgenciasCapa    Script Date: 15/08/00 13:54:53 ******/
CREATE PROCEDURE MDIAG_GetAgenciasCapa

	@Capa		Numeric(18),
	@DataProc	Int,
	@AgOrig		Smallint,
	@idCapa		Int

AS

	if @AgOrig Is Null
		SELECT	Cap.IdCapa,
			Cap.AgOrig,
			Cap.IdLote,
			Cap.Status,
			Stc.Descricao,
			Cap.Num_Malote,
			Ocorrencia,
			Cap.Idenv_Mal
		  FROM	Capa Cap, StatusCapa Stc 
		 WHERE	Cap.Capa		= @Capa
		   And  Cap.dataprocessamento	= @dataProc
		   And	Cap.Status             	= Stc.Status
		ORDER BY Cap.AgOrig

	Else
		SELECT	Cap.IdCapa,
			Cap.AgOrig,
			Cap.IdLote,
			Cap.Status,
			Stc.Descricao,
			Cap.Num_Malote,
			Ocorrencia,
			Cap.IdEnv_Mal
		  FROM	Capa Cap, StatusCapa Stc 
		 WHERE 	Cap.IdCapa		= @idCapa
		   And	Cap.dataprocessamento	= @dataProc		 
		   And	Cap.Status		= Stc.Status
		ORDER BY Cap.AgOrig





GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO



/****** Object:  Stored Procedure dbo.MDIAG_GetDocumentoContQualidade    Script Date: 17/11/00 11:22:10 ******/
CREATE PROCEDURE MDIAG_GetDocumentoContQualidade
	@DataProc	Int,
	@IdLote		Int
As


	SELECT	D.IdDocto, D.IdCapa, D.TipoDocto, D.Frente, D.Verso, IsNull(D.Leitura, '') AS Leitura, 
		D.Ordem, C.IdEnv_Mal, D.Status
	  FROM	Documento D, Capa C (NOLOCK)
	 WHERE	D.DataProcessamento 	= @DataProc
	   And	D.IdCapa		= C.IdCapa
	   And	C.IdLote		= @IdLote
	 ORDER BY IdDocto









GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO


CREATE Procedure MDIAG_GetEstatistica
@DataProc		int,
@IdEnv_Mal		char(1)

As

if @IdEnv_Mal = 'T' Begin   -- Todos

	SELECT 	Count(Distinct(C.IdCapa)) as QtdCapa, Count(D.IdDocto) as QtdDoc, C.Status
	  FROM 	Capa C (NOLOCK) Left Outer Join Documento D (NOLOCK INDEX = IndDataIdCapa) on (D.IdCapa = C.IdCapa And
		D.DataProcessamento = C.DataProcessamento) 
	 WHERE	C.DataProcessamento = @DataProc
	 GROUP BY C.Status

End
Else Begin

	SELECT 	Count(Distinct(C.IdCapa)) as QtdCapa, Count(D.IdDocto) as QtdDoc, C.Status
	  FROM 	Capa C (NOLOCK) Left Outer Join Documento D (NOLOCK INDEX = IndDataIdCapa) on (D.IdCapa = C.IdCapa And
		D.DataProcessamento = C.DataProcessamento)
	 WHERE	C.DataProcessamento 	= @DataProc
	   And	C.IdEnv_Mal 		= @IdEnv_Mal
	 GROUP BY C.Status

End






GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO



/****** Object:  Stored Procedure dbo.MDIAG_GetLoteContQualidade    Script Date: 17/11/00 11:22:09 ******/
CREATE PROCEDURE MDIAG_GetLoteContQualidade
	@DataProc	Int,
	@Intervalo	Int
As

	(
	SELECT IdLote, Status
	  FROM Lote 
	 WHERE DataProcessamento = @DataProc
	   And IdLote            > 0
	   And Status            = '0'
	)
	UNION ALL
	(
	SELECT IdLote, Status
	  FROM Lote 
	 WHERE DataProcessamento = @DataProc
	   And IdLote            > 0
	   And Status            = '1'
	   And DateDiff(Second, HoraAtual, GetDate()) > @Intervalo
	)
	Order By IdLote







GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO



/****** Object:  Stored Procedure dbo.MDIAG_GetMaloteExpedicao_oc    Script Date: 17/11/00 11:22:10 ******/
CREATE PROCEDURE MDIAG_GetMaloteExpedicao_oc
	@DataProc	Int,
	@NumMalote	Numeric(11)

As

	SELECT	Capa,IdCapa, IdLote, IdEnv_Mal,Idcapa, Num_Malote, 
		AgOrig,Status, Status as StatusAnt, 
		DateDiff(second, HoraAtual, GetDate()) as Intervalo,
		IsNull(Ocorrencia, 0) as Ocorrencia
	  FROM	Capa
	 WHERE 	DataProcessamento 	= @DataProc
	   And	Num_Malote		= @NumMalote
	   And	Status in 		('0','P')
	 ORDER	By Capa,AgOrig





GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO



/****** Object:  Stored Procedure dbo.MDIAG_GetMudaStatus    Script Date: 17/11/00 11:22:10 ******/
CREATE PROCEDURE MDIAG_GetMudaStatus
	@Status			VarChar (1),
	@NOcorrencia		Decimal(5),
	@DataProcessamento	Int,
	@IdCapa		Int

AS

Update Capa
	Set Status 	= @Status,
	Ocorrencia 	= @NOcorrencia
Where 	IdCapa 		= @IdCapa
And   	DataProcessamento = @DataProcessamento






GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO



/****** Object:  Stored Procedure dbo.MDIAG_GetRecuperaStatus    Script Date: 17/11/00 11:22:10 ******/
CREATE PROCEDURE MDIAG_GetRecuperaStatus

	@DataProc	Int,
	@NumMalote	Numeric(11),
	@Idcapa	Int,
	@Capa		Numeric(18)

As

	If @Capa is null 
		Select 	Capa.IdCapa, Capa.IdLote, Capa.IdEnv_Mal, Capa.Capa,Idcapa, Capa.Num_Malote, 
			Capa.AgOrig, Capa.Status, Capa.Status as StatusAnt,status.Descricao ,
			DateDiff(second, HoraAtual, GetDate()) as Intervalo,
			IsNull(Ocorrencia, 0) as Ocorrencia
		From	Capa Capa, StatusCapa Status
		Where 	Capa.DataProcessamento 	= @DataProc  	And
			Capa.Num_Malote        	= @NumMalote 	And
			Capa.Status            		= Status.Status 	And
			Capa.Idcapa		= @Idcapa   
		Order 	By Capa.Capa, Capa.AgOrig
	Else
		Select 	Capa.IdCapa, Capa.IdLote, Capa.IdEnv_Mal, Capa.Capa,Idcapa, Capa.Num_Malote, 
			Capa.AgOrig, Capa.Status, Capa.Status as StatusAnt,status.Descricao ,
			DateDiff(second, HoraAtual, GetDate()) as Intervalo,
			IsNull(Ocorrencia, 0) as Ocorrencia
		From	Capa Capa, StatusCapa Status
		Where	Capa.DataProcessamento 	= @DataProc  	And
			Capa.Capa  		= @Capa	And
			Capa.Status		= Status.Status
		Order	By Capa.Capa, Capa.AgOrig




GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO



/****** Object:  Stored Procedure dbo.MDIAG_GetTodasCapas    Script Date: 17/11/00 11:22:10 ******/
CREATE PROCEDURE MDIAG_GetTodasCapas 

	@DataProcessamento		Int,
	@IdCapa				Numeric(18),
	@idEnvMal			VarChar(1)

As

Declare 	@Erro           Int,
	@LinhasAfetadas int

	-------------------------------------------------------------------------------------------------
	-- Busca Todas Capas de Malote ou Envelope Cadastradas
	-------------------------------------------------------------------------------------------------

	if @IdCapa Is Null
		Select	Distinct(Capa),IdCapa,Status,Ocorrencia
		  From	Capa
		 Where	DataProcessamento  = @DataProcessamento
		   And	IdEnv_Mal          = @idEnvMal
		 Order 	By Capa

	Else
		Select	Distinct(Capa),IdCapa,Status,Ocorrencia
		  From	Capa
		 Where  DataProcessamento = @dataProcessamento
		   And	IdEnv_Mal         = @idEnvMal 
		   And	IdCapa            = @IdCapa 
		 Order  By Capa


	Select 	@Erro		= @@Error,
		@LinhasAfetadas	= @@Rowcount


	If @LinhasAfetadas = 0 or @Erro <> 0
		Return(1)

	Return(0)





GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO



/****** Object:  Stored Procedure dbo.MDIAG_GetTodasCapas_Oc    Script Date: 17/11/00 11:22:10 ******/
CREATE PROCEDURE MDIAG_GetTodasCapas_Oc 

	@DataProcessamento	Int,
	@idEnvMal		VarChar(1),		
	@IdCapa			Numeric(18), 
	@AgOrig			Smallint

As

Declare	@Erro		Int,
	@LinhasAfetadas	Int

	-------------------------------------------------------------------------------------------------
	-- Busca Tod as Capas de Malote ou Envelope Cadastradas
	-------------------------------------------------------------------------------------------------

	If @IdCapa Is Null
		Select 	Distinct(Capa),IdCapa
		  From Capa
		 Where	DataProcessamento = @dataProcessamento
		   And	IdEnv_Mal         = @idEnvMal 
		   And  Status            = '0'
		 Order By Capa
	Else
		Select Distinct(Capa),IdCapa
		  From Capa
		 Where  dataProcessamento = @dataProcessamento
		   And  IdEnv_Mal         = @idEnvMal 
		   And  Capa              = @IdCapa 
		   And  AgOrig            = @AgOrig 
		   And  Status            = '0'
		 Order By Capa

	Select 	@Erro           	= @@Error,
		@LinhasAfetadas 	= @@Rowcount

	If @LinhasAfetadas = 0 or @Erro <> 0
		   Return(1)


	Return(0)








GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO


CREATE Procedure MDIAG_GetTotalCapa
	@DataProc   	int,
	@IdEnv_Mal  	char(1)

As

if @IdEnv_Mal = 'T' Begin

	Select Count(*) as Total
	  From Capa (NOLOCK)
	 Where DataProcessamento = @DataProc

End
Else Begin

	Select Count(*) as Total
	  From Capa (NOLOCK)
	 Where DataProcessamento = @DataProc
	   And IdEnv_Mal         = @IdEnv_Mal

End







GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO


CREATE Procedure MDIAG_GetTotalDocumento
	@DataProc   	int,
	@IdEnv_Mal  	char(1)

As

if @IdEnv_Mal = 'T' Begin

	Select Count(*) as Total 
	  From Documento (NOLOCK)
	 Where DataProcessamento = @DataProc

End
Else Begin

	Select Count(*) as Total 
	  From Documento D, Capa C (NOLOCK)
	 Where D.DataProcessamento = @DataProc
	   And C.DataProcessamento = @DataProc
	   And D.IdCapa            = C.IdCapa
	   And C.IdEnv_Mal         = @IdEnv_Mal

End







GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

CREATE PROCEDURE MDIAG_GetUltimaImagemLote

@DataProcessamento 	Int ,
@IdLote			Int

AS

Select 	Max(substring(Frente,10,5)) as SeqInic
From	Documento
Where	DataProcessamento = @DataProcessamento
And       isnumeric(substring(Frente,10,5)) <> 0
And	Idcapa in (Select IdCapa from Capa Where DataProcessamento = @DataProcessamento And IdLote = @IdLote)

GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

GRANT  EXECUTE  ON [dbo].[MDIAG_GetUltimaImagemLote]  TO [mdi]
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO





/****** Object:  Stored Procedure dbo.MDIAG_InsereCapa    Script Date: 17/11/00 11:22:10 ******/
-----------------------------------------------------------------------------------------------------------
CREATE Proc MDIAG_InsereCapa
	@DataProcessamento 	Int,
	@IdLote   		Int,
	@idEnv_Mal  		Char(1),
	@Capa			Numeric(18),
	@Num_Malote		Numeric(11),
	@AgOrig			Smallint,
	@Status			Char(1),
	@IdCapa			Int OUTPUT
As

	Declare @TranCount	Int


	Select @TranCount = @@TranCount

	If @TranCount = 0
		Begin Transaction InsereCapa
	Else
		Save Transaction InsereCapa

--------------------------------------------------------------------------------------------------
	if @AgOrig <> 0 Begin
		-- Verifica se nao existe a Capa 
		if not exists (Select 1
				 From Capa 
				Where DataProcessamento = @DataProcessamento
				  And AgOrig   		= @AgOrig
				  And Capa     		= @Capa
				  And Status not in 	('D','F')) Begin

			------------------------------------------------------------------------------------------
			Insert Capa (  DataProcessamento, IdLote, idEnv_Mal, Capa, Num_Malote, AgOrig, 
				       Status, DataCriacao, Duplicidade )
			                Values(  @DataProcessamento, @idLote, @idEnv_Mal, @Capa, @Num_Malote, @AgOrig, 
				              @Status, GetDate(), 0 )
			if @@Error = 0 Begin
				Select @IdCapa = @@Identity
			End
			if @@Error = 0 Begin
				Commit Transaction InsereCapa
				Return(0)
			End
			Else Begin
				Rollback Transaction InsereCapa
				Return(2)
			End
		End
		Else Begin
		-- Ja existe capa com mesmo DataProc + AgOrig + Capa
			Rollback Transaction InsereCapa
			Return(1)
		End
	End
	Else Begin
		-- AgOrig = 0
		-- Verifica se nao existe a Capa 
		if not exists (Select 1
				 From Capa 
				Where DataProcessamento 	= @DataProcessamento
				  And Capa   			= @Capa
				  And Status not in 		('D','F')) Begin
			------------------------------------------------------------------------------------------
			Insert Capa (  DataProcessamento, IdLote, idEnv_Mal, Capa, Num_Malote, AgOrig, 
				        Status, DataCriacao, Duplicidade )
			          Values(  @DataProcessamento, @idLote, @idEnv_Mal, @Capa, @Num_Malote, @AgOrig, 
				         @Status, GetDate(), 0 )
			if @@Error = 0 Begin
				Select @IdCapa = @@Identity
			End
			if @@Error = 0 Begin
				Commit Transaction InsereCapa
				Return(0)
			End
			Else Begin
				Rollback Transaction InsereCapa
				Return(2)
			End
		End
		Else Begin
			-- Ja existe capa com mesmo DataProc + Capa
			------------------------------------------------------------------------------------------
			Insert Capa (  DataProcessamento, IdLote, idEnv_Mal, Capa, Num_Malote, AgOrig, 
				        Status, DataCriacao, Duplicidade )
			Values(@DataProcessamento, @idLote, @idEnv_Mal, @Capa, @Num_Malote, @AgOrig, 
			               @Status, GetDate(), @@RowCount )
			if @@Error = 0 Begin
				Select @IdCapa = @@Identity
			End
			if @@Error = 0 Begin
				Commit Transaction InsereCapa
				Return(0)
			End
			Else Begin
				Rollback Transaction InsereCapa
				Return(2)
			End
		End
	End



GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO



/****** Object:  Stored Procedure dbo.MDIAG_InsereLog    Script Date: 17/11/00 11:22:09 ******/
CREATE Procedure MDIAG_InsereLog
	@DataProc int,
	@IdCapa   int,
	@IdDocto  int,
	@Usuario  Char(10),
	@Acao     tinyint
As

Begin Transaction
	Insert Into Log
		(DataProcessamento,
		 IdCapa,
		 IdDocto,
		 Data,
		 Login,
		 Acao)
	Values
		(@DataProc,
		 @IdCapa,
		 @IdDocto,
		 GetDate(),
		 @Usuario,
		 @Acao)

if @@Error = 0 begin
	Commit Transaction
end
else begin
	Rollback Transaction
end




GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO



/****** Object:  Stored Procedure dbo.MDIAG_InsereLote    Script Date: 17/11/00 11:22:09 ******/
CREATE Proc MDIAG_InsereLote
	@DataProcessamento		int,
	@Prioridade			smallint,
	@Agencia			int,
	@IdLote 			int Output

as

	--------------------------------------------------------------------------------------------------
	-- Busca ultimo Lote da Agenica
	Select 	@IdLote = IsNull((Max(IdLote)+1),
		@Agencia * 100000 + 1)
	  From	Lote 
	 Where	IdLote between (@Agencia * 100000)
	   And ((@Agencia + 1) * 100000)
	   And  DataProcessamento = @DataProcessamento

	--------------------------------------------------------------------------------------------------
	-- Grava Lote
	Insert Lote (  DataProcessamento, IdLote, Status, Prioridade )
	     Values ( @DataProcessamento, @IdLote, '3', @Prioridade )
	if @@Error = 0
	Begin
		Return(0)
	End
	Else Begin
		select @IdLote = -1
		Return(1)
	End
	--------------------------------------------------------------------------------------------------





GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO



/****** Object:  Stored Procedure dbo.MDIAG_LimpaMovimento    Script Date: 17/11/00 11:22:10 ******/
CREATE PROCEDURE MDIAG_LimpaMovimento
	@pDataProcessamento 	Int
AS

Declare @Erro Int
	 Begin Transaction LimpaMovimento
	 /*-----------------------------------
	    Limpa tabela Documento
	 -----------------------------------*/
	 Truncate Table UbbMDI..Documento
	 Select @Erro = @@Error
	 If @Erro <> 0 Goto lbl_Erro
	 /*-----------------------------------
	    Limpa tabela Capa
	 -----------------------------------*/
	Truncate Table UbbMDI..Capa
	Select @Erro = @@Error
	 If @Erro <> 0 Goto lbl_Erro
	 /*-----------------------------------
	    Limpa tabela Lote
	 -----------------------------------*/
	Truncate Table UbbMDI..Lote
	 Select @Erro = @@Error
	 If @Erro <> 0 Goto lbl_Erro
	 /*-----------------------------------
	    Limpa tabela Agencia
	 -----------------------------------*/
	 Truncate Table UbbMDI..Agencia
	 Select @Erro = @@Error
	 If @Erro <> 0 Goto lbl_Erro
	 /*-----------------------------------
	    Limpa tabela Log
	 -----------------------------------*/
	 Truncate Table UbbMDI..Log
	 Select @Erro = @@Error
	 If @Erro <> 0 Goto lbl_Erro
	 /*-----------------------------------
	    Limpa tabela LogErro
	 -----------------------------------*/
	 Truncate Table UbbMDI..LogErro
	 Select @Erro = @@Error
	 If @Erro <> 0 Goto lbl_Erro
	 /*-------------------------------------------------------------------------------------------
	    Inserção do Lote 0
	 -------------------------------------------------------------------------------------------*/
	 Insert Into Lote
	 Values
	 (@pDataProcessamento, 0, '0','0',GETDATE())
	 Select @Erro = @@Error
	 If @Erro <> 0 Goto lbl_Erro
	 /*------------------------------------------------------------------------------------------
	    Atualização na tabela Parametro
	 -------------------------------------------------------------------------------------------*/
	 Update Parametro Set
	  DataProcessamento  = @pDataProcessamento,
	  Hm_Abertura  = GETDATE(),
	  Hm_Fechamento = NULL
	 Select @Erro = @@Error
	 If @Erro <> 0 Goto lbl_Erro

	 Commit Transaction LimpaMovimento

	 Update Statistics UbbMDI..Agencia
	 Update Statistics UbbMDI..Capa
	 Update Statistics UbbMDI..Documento
	 Update Statistics UbbMDI..Log
	 Update Statistics UbbMDI..LogErro
	 Update Statistics UbbMDI..Lote
lbl_Erro:
	 If @Erro <> 0 RollBack Transaction LimpaMovimento
 
	 Return @Erro











GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO


CREATE PROCEDURE MDIAG_ObtemCapasLote
	@pDataProcessamento	Int,
	@pIdLote		Int

AS


	SELECT	'C' AS TipoRegistro,
		IdCapa,
		IdLote,
		IdEnv_Mal,
		Capa,
		Num_Malote,
		AgOrig,
		Status,
		IsNull(Ocorrencia,0) AS Ocorrencia,
		Duplicidade
	  FROM	CAPA
	 WHERE	DataProcessamento 	= @pDataProcessamento
	   AND 	IdLote			= @pIdLote







GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO


CREATE PROCEDURE MDIAG_ObtemDocumentosCapa
	@pDataProcessamento	Int,
	@pIdCapa		Int
AS


	SELECT	'D' AS TipoRegistro,
		IdDocto,
		OrdemCaptura,
		TipoDocto,
		Leitura,
		Frente,
		Verso,
		Status,
		Ordem
	  FROM	DOCUMENTO
	 WHERE	DataProcessamento 	= @pDataProcessamento And
	   		IdCapa			= @pIdCapa
	Order By OrdemCaptura




GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO


CREATE PROCEDURE MDIAG_ObtemLog
	@pDataProcessamento	Int,
	@pIdCapa		Int,
	@pIdDocto		Int

AS


	If (@pIdCapa Is Not NULL) OR (@pIdCapa <> 0) Begin
		/*--------------------------------------------
		   Seleção dos logs das capas
		--------------------------------------------*/
		SELECT	'G' AS TipoRegistro,
			Data,
			Login,
			Acao	
		  FROM	LOG
		 WHERE	DataProcessamento 	= @pDataProcessamento
		   AND	IdCapa			= @pIdCapa
		   AND	IdDocto			= 0

	End
	Else Begin
		/*-----------------------------------------------------
		   Seleção dos logs dos documentos
		------------------------------------------------------*/
		SELECT	'G' AS TipoRegistro,
			Data,
			Login,
			Acao	
		  FROM	LOG
		 WHERE	DataProcessamento 	= @pDataProcessamento
		   AND	IdDocto			= @pIdDocto
	End







GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO


CREATE PROCEDURE MDIAG_ObtemLotesExportacao
	@pDataProcessamento	Int

AS


	SELECT	'L' AS TipoRegistro,
		IdLote,
		Status,
		Prioridade
	  FROM	LOTE
	 WHERE	DataProcessamento = @pDataProcessamento






GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO



/****** Object:  Stored Procedure dbo.MDIAG_RemoveCapa    Script Date: 17/11/00 11:22:10 ******/
CREATE PROCEDURE MDIAG_RemoveCapa
	@DataProcessamento	Int,
	@IdCapa		Int

As

	Begin Transaction

	--------------------------------------------------------------------------------------------------
	-- Deleta Capa
	DELETE 	FROM Capa
	 WHERE	DataProcessamento 	= @DataProcessamento
	   And  IdCapa			= @IdCapa
	if @@Error = 0 Begin
		Commit Transaction
		Return(0)
	End
	Else Begin
		RollBack Transaction
		Return(1)
	End
	-------------------------------------------------------------------------------------------------- 




GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO



/****** Object:  Stored Procedure dbo.MDIAG_RemoveDocumento    Script Date: 17/11/00 11:22:10 ******/
CREATE PROCEDURE MDIAG_RemoveDocumento
	@DataProcessamento	Int,
	@IdDocto		Int
As

	Begin Transaction

   -------------------------------
   -- Remove da Tabela Documento
   -------------------------------
   Delete Documento 
    Where DataProcessamento = @DataProcessamento
      And IdDocto           = @IdDocto

  ---------------------------------------
   -- Se tudo OK, Commit, senao RollBack
   ---------------------------------------
   If @@Error = 0 Begin
	Commit Transaction
	Return(0)
   End
   Else Begin
	RollBack Transaction
   	Return(1)
   End








GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO



/****** Object:  Stored Procedure dbo.MDIAG_RemoveLote    Script Date: 17/11/00 11:22:10 ******/
CREATE Proc MDIAG_RemoveLote
	@DataProcessamento	int,
	@IdLote 			int 

as

	Begin Transaction
 
	--------------------------------------------------------------------------------------------------
	-- Deleta Lote
	DELETE FROM Lote
	 WHERE DataProcessamento = @DataProcessamento
	   And IdLote            = @IdLote
	if @@Error = 0 Begin
		Commit Transaction
		Return(0)
	End
	Else Begin
		RollBack Transaction
		Return(1)
	End
	-------------------------------------------------------------------------------------------------- 




GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO




CREATE PROCEDURE MDIAG_SplitCapa
	@pDataProcessamento	Int,
	@pIdCapaAtual		Int,
	@pIdDoctoAtual		Int,
	@pCapa			Numeric(9),
	@pNum_Malote		Numeric(9),
	@pIdEnv_Mal		Char(1),
	@pAgOrig		SmallInt
AS


	Declare @vErro			Int
	Declare @vTranCount		Int
	Declare @vIdLote		Int
	Declare @vIdCapa		Int
	Declare @vRet			Int
	Declare @vOrdemCaptura		Int
	Declare @vDataProcessamento	Int
	Declare @vIdDocto		Int


	DECLARE cDocumentos INSENSITIVE CURSOR FOR
		SELECT 	DataProcessamento, IdDocto
		  FROM	Documento
		 WHERE  DataProcessamento 	 = @pDataProcessamento
		   AND  IdCapa			 = @pIdCapaAtual
		   AND  IdDocto			>= @pIdDoctoAtual

	Begin Transaction

	/*------------------------
	  Definição de variaveis
	------------------------*/
	Select 	@vErro 			= 0,
		@vOrdemCaptura 	= 1

	Open cDocumentos

	FETCH cDocumentos INTO @vDataProcessamento, @vIdDocto

	If @@FETCH_STATUS IN(-1,-2) Begin
		Select @vErro = 1
		Goto Fechar_Cursor
	End


	Select @vTranCount = @@TranCount

	/*--------------------------
	  Pega IdLote da capa atual
	--------------------------*/
	SELECT 	@vIdLote	= IdLote
	  FROM 	Capa
	 WHERE 	DataProcessamento = @pDataProcessamento
	   AND 	IdCapa		  = @pIdCapaAtual


	/*-----------------------------------------------------------------------------
	  Não utilizo a SP MDIAG_InsereCapa, porque não funciona de acordo com a regra
	-----------------------------------------------------------------------------*/

	Insert Capa (DataProcessamento, IdLote, idEnv_Mal, Capa, Num_Malote, AgOrig, 
		Status, DataCriacao, Duplicidade )

	Values(@pDataProcessamento, @vIdLote, @pIdEnv_Mal, @pCapa,@pNum_Malote, @pAgOrig, 
		'1', GetDate(), 0 )

	Select @vErro 	= @@Error
	Select @vIdCapa = @@Identity

	If (@vErro <> 0) Or (@vIdCapa Is Null) Goto Fechar_Cursor

	/*------------------------------------------------------
	  LOOP dos documentos para alterar o campo OrdemCaptura
	------------------------------------------------------*/
	Select @vIdDocto = @pIdDoctoAtual

	While @@FETCH_STATUS NOT IN (-1,-2) Begin

		UPDATE Documento SET
			OrdemCaptura 	= @vOrdemCaptura,
			IdCapa		= @vIdCapa
		 WHERE DataProcessamento 	= @pDataProcessamento
		   AND IdCapa			= @pIdCapaAtual
		   AND IdDocto		        = @vIdDocto


		Select @vErro = @@Error

		If @vErro <> 0 Goto Fechar_Cursor

		Select @vOrdemCaptura = @vOrdemCaptura + 1
		FETCH cDocumentos INTO @vDataProcessamento, @vIdDocto
	End


	/*---------------------------------------------
	  Atualiza Tipo Docto para Envelope ou Malote
	---------------------------------------------*/
	UPDATE Documento SET
		TipoDocto = 1
	 WHERE IdCapa 	= @vIdCapa
	   AND IdDocto 	= @pIdDoctoAtual

	Select @vErro = @@Error

Fechar_Cursor:
	/*-------------------------
	  Fecha e desaloca o cursor
	-------------------------*/
	Close cDocumentos
	Deallocate cDocumentos

Fim_Proc:
	If @vErro <> 0 Begin
		RollBack Transaction
		Return (@vErro)
	End
	Else Begin	
		Commit Transaction
		Return @vErro
	End



GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  ON    SET ANSI_NULLS  ON 
GO

CREATE PROCEDURE MDIAG_SplitCapaAnterior
	@pDataProcessamento	Int,
	@pIdCapaAtual		Int,
	@pIdDoctoAtual		Int,
	@pTipoDocto		SmallInt
AS

	Declare @vErro			Int
	Declare @vTranCount		Int
	Declare @vIdCapaAnterior		Int
	Declare @vIdDoctoAnterior	Int
	Declare @vOrdemCapturaAnterior	Int
	Declare @vIdDocto		Int


	/*------------------------
	  Definição de variaveis
	------------------------*/
	Select 	@vErro 		= 0

	/*-------------------------------
	  Pega o IdDocto e Capa Anterior 
	-------------------------------*/
	SELECT	@vIdDoctoAnterior	= Max(IdDocto)
	  FROM	Documento
	 WHERE	DataProcessamento 	= @pDataProcessamento
	   AND	IdDocto			< @pIdDoctoAtual


	SELECT 	@vIdCapaAnterior  	= IdCapa,
			@vOrdemCapturaAnterior	= OrdemCaptura + 1
	  FROM 	Documento
	 WHERE	DataProcessamento	= @pDataProcessamento
	   AND	IdDocto			= @vIdDoctoAnterior

	/*--------------------------------
	  Abertura do cursor de documentos
	--------------------------------*/
	DECLARE cDocumentos INSENSITIVE CURSOR FOR
		SELECT 	IdDocto
		  FROM	Documento
		 WHERE 	DataProcessamento 	 = @pDataProcessamento
		   AND  IdCapa			 = @pIdCapaAtual
		   AND  IdDocto			>= @pIdDoctoAtual

	/*----------------------
	  Controle de Transação
	----------------------*/
	Begin Transaction

	Open cDocumentos

	FETCH cDocumentos INTO @vIdDocto

	If @@FETCH_STATUS IN(-1,-2) Begin
		Select @vErro = 1
		Goto Fechar_Cursor
	End

	Select @vTranCount = @@TranCount

	/*------------------------------------------------------
	  LOOP dos documentos para alterar os campos IdCapa, OrdemCaptura
	------------------------------------------------------*/
	Select @vIdDocto = @pIdDoctoAtual

	While @@FETCH_STATUS NOT IN (-1,-2) Begin

		UPDATE Documento SET
			IdCapa		= @vIdCapaAnterior,
			OrdemCaptura 	= @vOrdemCapturaAnterior
		 WHERE DataProcessamento 	= @pDataProcessamento
		   AND IdCapa			= @pIdCapaAtual
		   AND IdDocto		        = @vIdDocto

		Select @vErro = @@Error

		If @vErro <> 0 Goto Fechar_Cursor

		Select @vOrdemCapturaAnterior = @vOrdemCapturaAnterior + 1
		FETCH cDocumentos INTO @vIdDocto
	End

	/*--------------------------------------------------
	  Atualiza Tipo Docto para o documento selecionado
	--------------------------------------------------*/
	UPDATE Documento SET
		TipoDocto = @pTipoDocto
	 WHERE DataProcessamento = @pDataProcessamento
	   AND IdCapa 	= @vIdCapaAnterior
	   AND IdDocto 	= @pIdDoctoAtual

	Select @vErro = @@Error

	If @vErro = 0 Begin
		DELETE FROM Capa
		 WHERE DataProcessamento = @pDataProcessamento
		   AND IdCapa 	= @pIdCapaAtual

		Select @vErro = @@Error
	End

Fechar_Cursor:
	/*-------------------------
	  Fecha e desaloca o cursor
	-------------------------*/
	Close cDocumentos
	Deallocate cDocumentos

Fim_Proc:
	If @vErro <> 0 Begin
		RollBack Transaction
		Return (@vErro)
	End
	Else Begin
		Commit Transaction
		Return @vErro
	End






GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO


CREATE PROCEDURE MDIAG_TotalizaAgencia
	@pAgencia	SmallInt,
	@pLacre	Decimal(8)
AS


Declare @vTotal_Env		Int
Declare @vTotal_Mal		Int
Declare @vEnv_HoraCadastrada	Char(5)
Declare @vMal_HoraCadastrada	Char(5)
Declare @Erro			Int


	Select @Erro = 0

	/*--------------------------
	   Insere Envelope
	--------------------------*/
	SELECT	@vTotal_Env 		= Count(CP.IdEnv_Mal),
		@vEnv_HoraCadastrada 	= Convert(Char(5),Max(CP.DataCriacao),108)
	  FROM 	CAPA CP
	 WHERE 	CP.IdEnv_Mal = 'E'

	If Exists(Select 1
		    From Agencia
		   Where IdEnv_Mal = 'E'
		     AND QtdInformada = @vTotal_Env) Begin
		UPDATE Agencia	SET
			Agencia		=	@pAgencia,
			Lacre		=	@pLacre,
			QtdInformada	=	@vTotal_Env,
			HoraCadastrada	=	@vEnv_HoraCadastrada
		WHERE	QtdInformada	= 	@vTotal_Env
		  AND	IdEnv_Mal	=	'E'

		Select @Erro = @@Error
		If @Erro <> 0 Goto lbl_Erro
	End
	Else Begin		If @vTotal_Env > 0 Begin
			INSERT INTO Agencia
			(Agencia,	Lacre,		QtdInformada,	HoraCadastrada,		IdEnv_Mal)
			VALUES
			(@pAgencia,	@pLacre,	@vTotal_Env,	@vEnv_HoraCadastrada,	'E')

			Select @Erro = @@Error
		End

		If @Erro <> 0 Goto lbl_Erro
	End

	/*---------------------------
	   Insere Malote
	---------------------------*/

	SELECT @vTotal_Mal 		= Count(CP.IdEnv_Mal),
		@vMal_HoraCadastrada 	= Convert(Char(5),Max(CP.DataCriacao),108)
	  FROM CAPA CP
	 WHERE CP.IdEnv_Mal = 'M'

	If Exists(Select 1 From Agencia Where IdEnv_Mal = 'M' AND QtdInformada = @vTotal_Mal) Begin
		UPDATE	Agencia	SET
			Agencia		=	@pAgencia,
			Lacre		=	@pLacre,
			QtdInformada	=	@vTotal_Mal,
			HoraCadastrada	=	@vMal_HoraCadastrada
		 WHERE	QtdInformada	=	@vTotal_Mal
		   AND	IdEnv_Mal	=	'M'

		Select @Erro = @@Error
	End
	Else Begin
		if @vTotal_Mal > 0 Begin
			INSERT INTO Agencia
			(Agencia,	Lacre,		QtdInformada,	HoraCadastrada,		IdEnv_Mal)
			VALUES
			(@pAgencia,	@pLacre,	@vTotal_Mal,	@vMal_HoraCadastrada,	'M')

			Select @Erro = @@Error
		End
	End

lbl_Erro:


	Return @Erro




GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  ON    SET ANSI_NULLS  ON 
GO

CREATE PROCEDURE MDIAG_VerificaCapaExportar
	@pDataProcessamento	Int

AS

	/*----------------------------------------------------------------------------------------------------------
	   Se existir capa recepcionada, com ocorrencia ou capturada pode continuar
	------------------------------------------------------------------------------------------------------------*/

	If Exists(SELECT 1
		    FROM Capa
		   WHERE DataProcessamento 	= @pDataProcessamento) Begin
		Return (0)
	End
	Else
		/*---------------------------------------------
		   Caso contrario nao pode continuar
		---------------------------------------------*/
		Return (1)






GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

GRANT  EXECUTE  ON [dbo].[MDIAG_VerificaCapaExportar]  TO [mdi]
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO




CREATE PROCEDURE MDIAG_VerificaCapaRecepcionada
	@DataProcessamento              Int,
	@Capa                           Numeric(18),
	@AgOrig				SmallInt,
	@Num_Malote			Numeric(11)

as

	--- Pesquisa Capa recepcionada ---
	if Exists(SELECT 1
		    FROM Capa 
		   WHERE DataProcessamento	= @DataProcessamento  
		     AND Capa 			= @Capa
		     AND Status			= "0"
		     AND Convert(Char(10),AgOrig) like CASE @Num_Malote WHEN 0 THEN Convert(Char(10),@AgOrig) ELSE "%" END)
		Return(1)

	Return(0)








GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO



/****** Object:  Stored Procedure dbo.MDIAG_VerificaLoteDisponivel    Script Date: 17/11/00 11:22:10 ******/
CREATE PROCEDURE MDIAG_VerificaLoteDisponivel
	@DataProc	Int,
	@IdLote		Int,
	@Intervalo	Int

As

Declare	@Count		Int,
	@Error		Int,
	@m_Status	Char(1),
	@m_Dif		Int

	SELECT 	@m_Status = Status,
		@m_Dif = DateDiff(second, HoraAtual, GetDate())
	  FROM 	Lote
	 WHERE	DataProcessamento 	= @DataProc
	   AND	IdLote			= @IdLote

	Select @Error = @@Error, @Count = @@RowCount

	if @Error <> 0 Begin
		Return(2)
	End
	Else If @Count = 0 Begin
		Return(1)
	End
	Else If (@m_Status = "0") Or (@m_Status = "1" And @m_Dif >= @Intervalo) Begin
		Return(0)
	End
	Else Begin
		Return(1)
	End



GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO


CREATE PROCEDURE MDIAG_VerificaLotePendente
	@pDataProcessamento	Int

AS



	/*-------------------------------------------------------------------------------------------
	   Se existir mais de um lote com status <> 2 não pode continuar
	-------------------------------------------------------------------------------------------*/

	If Exists(SELECT 1
		    FROM LOTE
		   WHERE DataProcessamento 	= @pDataProcessamento
		     AND IdLote			> 0
		     AND Status			<> 2) Begin
		Return (1)
	End
	Else
		/*---------------------------------------------
		   Caso contrario pode continuar
		---------------------------------------------*/
		Return (0)





GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO


/****** Object:  Trigger dbo.tI_Capa    Script Date: 17/11/00 11:18:10 ******/
----------------------------------------------------------------------------------------------------------------
create trigger tI_Capa on Capa for 
INSERT, UPDATE as
begin
  declare  @numrows  int,
           @nullcnt  int,
           @validcnt int,
           @errno    int,
           @errmsg   varchar(255)
  select @numrows = @@rowcount
  ------------------------------------------------------------------------------------------------------  
  if @numrows = 0 
     return
  ------------------------------------------------------------------------------------------------------  
  -- Verifica se existe o codigo de Lote
  if not exists ( select * from inserted,Lote  where 
    inserted.DataProcessamento  = Lote.DataProcessamento And
    inserted.idLote   = Lote.idLote )
    begin
      select @errno  = 30001,
             @errmsg = 'Atencao nao existe este codigo de Lote.'
      goto error
  end
  ------------------------------------------------------------------------------------------------------
  return
error:
    raiserror @errno @errmsg
    rollback transaction
end


GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO


/****** Object:  Trigger dbo.tD_Capa    Script Date: 17/11/00 11:18:10 ******/
----------------------------------------------------------------------------------------------------------------
create trigger tD_Capa on Capa for 
DELETE as
begin
  declare  @numrows int,
           @nullcnt int,
           @validcnt int,
           @errno   int,
           @errmsg  varchar(255)
  select @numrows = @@rowcount
  ------------------------------------------------------------------------------------------------------  
  if @numrows = 0 
     return
  ------------------------------------------------------------------------------------------------------  
  -- Verifica se existe o codigo de Capa na Tabela Documento
  if exists ( select * from deleted,Documento where 
      deleted.DataProcessamento = Documento.DataProcessamento And
      deleted.idCapa     = Documento.idCapa )
    begin
      select @errno  = 30001,
             @errmsg = 'Atencao existe Documento desta Capa.'
      goto error
  end
  ------------------------------------------------------------------------------------------------------
  return
error:
    raiserror @errno @errmsg
    rollback transaction
end


GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO


/****** Object:  Trigger dbo.tI_Documento    Script Date: 17/11/00 11:18:10 ******/
----------------------------------------------------------------------------------------------------------------
create trigger tI_Documento on Documento for 
INSERT, UPDATE as
begin
  declare  @numrows  int,
           @nullcnt  int,
           @validcnt int,
           @errno    int,
           @errmsg   varchar(255)
  select @numrows = @@rowcount
  ------------------------------------------------------------------------------------------------------  
  if @numrows = 0 
     return
  ------------------------------------------------------------------------------------------------------  
  -- Verifica se existe o codigo de Capa
  if not exists ( select * from inserted,Capa  where 
    inserted.DataProcessamento  = Capa.DataProcessamento And
    inserted.idCapa   = Capa.idCapa )
    begin
      select @errno  = 30001,
             @errmsg = 'Atencao nao existe este codigo de Capa.'
      goto error
  end
  ------------------------------------------------------------------------------------------------------      
  return
error:
    raiserror @errno @errmsg
    rollback transaction
end


GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO


/****** Object:  Trigger dbo.tD_Lote    Script Date: 17/11/00 11:18:11 ******/
----------------------------------------------------------------------------------------------------------------
create trigger tD_Lote on Lote for 
DELETE as
begin
  declare  @numrows int,
           @nullcnt int,
           @validcnt int,
           @errno   int,
           @errmsg  varchar(255)
  select @numrows = @@rowcount
  ------------------------------------------------------------------------------------------------------  
  if @numrows = 0 
     return
  ------------------------------------------------------------------------------------------------------  
  -- Verifica se existe o codigo de Capa
  if exists ( select * from deleted,Capa where 
      deleted.DataProcessamento = Capa.DataProcessamento And
      deleted.idLote     = Capa.idLote )
    begin
      select @errno  = 30001,
             @errmsg = 'Atencao existe Capa deste Lote.'
      goto error
  end
  ------------------------------------------------------------------------------------------------------      
  return
error:
    raiserror @errno @errmsg
    rollback transaction
end


GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

