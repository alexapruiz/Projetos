ALTER TABLE [dbo].[Log] DROP CONSTRAINT FK_LOG_REF_15195_ACAO
GO

ALTER TABLE [dbo].[Documento] DROP CONSTRAINT FK_DOCUMENT_REF_4572_CAPA
GO

ALTER TABLE [dbo].[Capa] DROP CONSTRAINT FK_CAPA_REF_4584_LOTE
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

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_AtualizaAgencia]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_AtualizaAgencia]
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

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_EncerraMovimento]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_EncerraMovimento]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_GetAgenciasCapa]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_GetAgenciasCapa]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_GetAllAgenf]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_GetAllAgenf]
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

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_GetTodasOcorrencia]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_GetTodasOcorrencia]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_GetTotalCapa]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_GetTotalCapa]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_GetTotalDocumento]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_GetTotalDocumento]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_InsereCapa]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_InsereCapa]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_InsereLog]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_InsereLog]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_InsereLogErro]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_InsereLogErro]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_InsereLote]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_InsereLote]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_LerParametro]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_LerParametro]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_LimpaMovimento]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_LimpaMovimento]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_ObtemAgencias]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_ObtemAgencias]
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

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_TotalizaAgencia]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_TotalizaAgencia]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_VerificaLoteDisponivel]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_VerificaLoteDisponivel]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_VerificaLotePendente]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_VerificaLotePendente]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[Acao]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Acao]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[Agencia]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Agencia]
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

if exists (select * from sysobjects where id = object_id(N'[dbo].[LogErro]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[LogErro]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[Lote]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Lote]
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

if not exists (select * from master..syslogins where name = N'CASH\Administrador')
	exec sp_grantlogin N'CASH\Administrador'
	exec sp_defaultdb N'CASH\Administrador', N'master'
	exec sp_defaultlanguage N'CASH\Administrador', N'us_english'
GO

if not exists (select * from master..syslogins where name = N'i')
BEGIN
	declare @logindb nvarchar(132), @loginlang nvarchar(132) select @logindb = N'MDI_Ubb', @loginlang = N'Português'
	if @logindb is null or not exists (select * from master..sysdatabases where name = @logindb)
		select @logindb = N'master'
	if @loginlang is null or (not exists (select * from master..syslanguages where name = @loginlang) and @loginlang <> N'us_english')
		select @loginlang = @@language
	exec sp_addlogin N'i', null, @logindb, @loginlang
END
GO

if not exists (select * from master..syslogins where name = N'mdi')
BEGIN
	declare @logindb nvarchar(132), @loginlang nvarchar(132) select @logindb = N'MDI_Ubb', @loginlang = N'Português'
	if @logindb is null or not exists (select * from master..sysdatabases where name = @logindb)
		select @logindb = N'master'
	if @loginlang is null or (not exists (select * from master..syslanguages where name = @loginlang) and @loginlang <> N'us_english')
		select @loginlang = @@language
	exec sp_addlogin N'mdi', null, @logindb, @loginlang
END
GO

if not exists (select * from master..syslogins where name = N'ot')
BEGIN
	declare @logindb nvarchar(132), @loginlang nvarchar(132) select @logindb = N'MDI_Ubb', @loginlang = N'Português'
	if @logindb is null or not exists (select * from master..sysdatabases where name = @logindb)
		select @logindb = N'master'
	if @loginlang is null or (not exists (select * from master..syslanguages where name = @loginlang) and @loginlang <> N'us_english')
		select @loginlang = @@language
	exec sp_addlogin N'ot', null, @logindb, @loginlang
END
GO

exec sp_addsrvrolemember N'BUILTIN\Administradores', sysadmin
GO

exec sp_addsrvrolemember N'CASH\Administrador', sysadmin
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

CREATE TABLE [dbo].[LogErro] (
	[Data] [datetime] NOT NULL ,
	[Erro] [int] NULL ,
	[Descricao] [varchar] (255) NULL 
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

ALTER TABLE [dbo].[Acao] WITH NOCHECK ADD 
	CONSTRAINT [PK_Acao] PRIMARY KEY  CLUSTERED 
	(
		[Acao]
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

GRANT  SELECT ,  INSERT ,  DELETE ,  UPDATE  ON [dbo].[Capa]  TO [mdi]
GO

GRANT  SELECT ,  INSERT ,  DELETE ,  UPDATE  ON [dbo].[Documento]  TO [mdi]
GO

GRANT  SELECT ,  INSERT ,  DELETE ,  UPDATE  ON [dbo].[Log]  TO [mdi]
GO

GRANT  SELECT ,  INSERT ,  DELETE ,  UPDATE  ON [dbo].[LogErro]  TO [mdi]
GO

GRANT  SELECT ,  INSERT ,  DELETE ,  UPDATE  ON [dbo].[Lote]  TO [mdi]
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

ALTER TABLE [dbo].[Capa] ADD 
	CONSTRAINT [FK_CAPA_REF_21052_OCORRENC] FOREIGN KEY 
	(
		[Ocorrencia]
	) REFERENCES [dbo].[Ocorrencia] (
		[Ocorrencia]
	),
	CONSTRAINT [FK_CAPA_REF_4584_LOTE] FOREIGN KEY 
	(
		[DataProcessamento],
		[IdLote]
	) REFERENCES [dbo].[Lote] (
		[DataProcessamento],
		[IdLote]
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
	CONSTRAINT [FK_DOCUMENT_REF_4572_CAPA] FOREIGN KEY 
	(
		[DataProcessamento],
		[IdCapa]
	) REFERENCES [dbo].[Capa] (
		[DataProcessamento],
		[IdCapa]
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

GRANT  EXECUTE  ON [dbo].[MDIAG_AtualizaAgencia]  TO [mdi]
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

GRANT  EXECUTE  ON [dbo].[MDIAG_AtualizaStatusLote]  TO [mdi]
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

GRANT  EXECUTE  ON [dbo].[MDIAG_CapturaCapa]  TO [mdi]
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

GRANT  EXECUTE  ON [dbo].[MDIAG_CapturaDocumento]  TO [mdi]
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

GRANT  EXECUTE  ON [dbo].[MDIAG_GetAgenciasCapa]  TO [mdi]
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO


/****** Object:  Stored Procedure dbo.MDIAG_GetAllAgenf    Script Date: 17/11/00 11:22:09 ******/
CREATE PROCEDURE MDIAG_GetAllAgenf
AS
Select  '0035' As agefscdagen,
 'Santos' As agefsnoagen


GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

GRANT  EXECUTE  ON [dbo].[MDIAG_GetAllAgenf]  TO [mdi]
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO


/****** Object:  Stored Procedure dbo.MDIAG_GetDocumentoContQualidade    Script Date: 17/11/00 11:22:10 ******/
CREATE PROCEDURE MDIAG_GetDocumentoContQualidade
	@DataProc	Int,
	@IdLote		Int
As


	SELECT	D.IdDocto, D.IdCapa, D.TipoDocto, D.Frente, D.Verso, IsNull(D.Leitura, '') AS Leitura, D.Ordem
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

GRANT  EXECUTE  ON [dbo].[MDIAG_GetEstatistica]  TO [mdi]
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

GRANT  EXECUTE  ON [dbo].[MDIAG_GetLoteContQualidade]  TO [mdi]
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

GRANT  EXECUTE  ON [dbo].[MDIAG_GetMaloteExpedicao_oc]  TO [mdi]
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

GRANT  EXECUTE  ON [dbo].[MDIAG_GetMudaStatus]  TO [mdi]
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

GRANT  EXECUTE  ON [dbo].[MDIAG_GetRecuperaStatus]  TO [mdi]
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

GRANT  EXECUTE  ON [dbo].[MDIAG_GetTodasCapas]  TO [mdi]
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

GRANT  EXECUTE  ON [dbo].[MDIAG_GetTodasCapas_Oc]  TO [mdi]
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

GRANT  EXECUTE  ON [dbo].[MDIAG_GetTodasOcorrencia]  TO [mdi]
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

GRANT  EXECUTE  ON [dbo].[MDIAG_GetTotalCapa]  TO [mdi]
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

GRANT  EXECUTE  ON [dbo].[MDIAG_GetTotalDocumento]  TO [mdi]
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
	@AgOrig		Smallint,
	@Status			Char(1),
	@IdCapa		Int OUTPUT
As

Begin Transaction
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
				Commit Transaction
				Return(0)
			End
			Else Begin
				Rollback Transaction
				Return(2)
			End
		End
		Else Begin
		-- Ja existe capa com mesmo DataProc + AgOrig + Capa
			Rollback Transaction
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
				Commit Transaction
				Return(0)
			End
			Else Begin
				Rollback Transaction
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
				Commit Transaction
				Return(0)
			End
			Else Begin
				Rollback Transaction
				Return(2)
			End
		End
	End




GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

GRANT  EXECUTE  ON [dbo].[MDIAG_InsereCapa]  TO [mdi]
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

GRANT  EXECUTE  ON [dbo].[MDIAG_InsereLog]  TO [mdi]
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

GRANT  EXECUTE  ON [dbo].[MDIAG_InsereLogErro]  TO [mdi]
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

GRANT  EXECUTE  ON [dbo].[MDIAG_InsereLote]  TO [mdi]
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

GRANT  EXECUTE  ON [dbo].[MDIAG_LerParametro]  TO [mdi]
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
	 Delete From UbbMDI..Capa
	 Select @Erro = @@Error
	 If @Erro <> 0 Goto lbl_Erro
	 /*-----------------------------------
	    Limpa tabela Lote
	 -----------------------------------*/
	 Delete From UbbMDI..Lote
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

GRANT  EXECUTE  ON [dbo].[MDIAG_LimpaMovimento]  TO [mdi]
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
	   AND	IdLote > 0




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

Declare	@TipoDocto		SmallInt,
	@Erro			Int,
	@LinhasAfetadas		Int

---------------------------------------------------
--Seleciona o Tipo do Documento que serah removido
---------------------------------------------------
	SELECT @TipoDocto = TipoDocto
	  FROM Documento
	 WHERE DataProcessamento = @DataProcessamento
	   And IdDocto           = @IdDocto

	Select @Erro           = @@Error
	      ,@LinhasAfetadas = @@Rowcount

	If @LinhasAfetadas <> 1  or @Erro <> 0       -- Linha nao encontrada ou erro 
	   Return(1)


	Begin Transaction

	----------------------------------------------------------------------------	
	-- Remove o Documento da sua tabela especifica, se houver (ver comentarios)
	----------------------------------------------------------------------------


	-- TipoDocto = 0 (DOCUMENTO INDEFINIDO) soh estah na tabela de Documento
	-- TipoDocto = 1 (CAPA MALOTE EMPRESA) soh apaga da tabela de Documento


	If @TipoDocto in (2,3)        -- DEPOSITO CONTA CORRENTE E DEPOSITO CONTA POUPANCA        
           	Delete Deposito 
	     	 Where DataProcessamento = @DataProcessamento
	     	   And IdDocto           = @IdDocto
  
	Else If @TipoDocto = 4        -- ADCC                           
 	     Delete ADCC 
	      Where DataProcessamento = @DataProcessamento
	        And IdDocto           = @IdDocto

	Else If @TipoDocto in (5,6,7) -- CHEQUE UBB SACADO (PAGTO), CHEQUE TERCEIRO (PAGTO), 
                                    -- CHEQUE DEPOSITO      
  	     Delete Cheque 
	      Where DataProcessamento = @DataProcessamento
	        And IdDocto           = @IdDocto

	Else If @TipoDocto IN (8, 9, 24, 25, 26) -- CONCESSIONARIA VALOR REAL,
                                               -- CONCESSIONARIA VALOR INDEXADO
                                               -- TRIBUTOS MUNICIPAIS, TRIBUTOS ESTADUAIS,
                                               -- TRIBUTOS FEDERAIS
	     Delete CBIndex
	      Where DataProcessamento = @DataProcessamento
	        And IdDocto           = @IdDocto

	Else If @TipoDocto IN (10,28,29,30,31)    -- FICHA COMPENSACAO, UNICOBRANCA UBB,
                                                -- COBRANCA IMEDIATA UBB, COBRANCA ESPECIAL UBB
                                                -- COBRANCA TERCEIROS
	     Delete FichaCompensacao 
	      Where DataProcessamento = @DataProcessamento
	        And IdDocto           = @IdDocto

	Else If @TipoDocto IN (11,19) -- INSS e GRPS - nao existem mais
	     Goto TrataErro

	Else If @TipoDocto = 12       -- TITULOS (TERCEIROS SEM CB) 
	     Delete Titulo 
	      Where DataProcessamento = @DataProcessamento
	        And IdDocto           = @IdDocto

	Else If @TipoDocto = 13       -- COBRANCA REGISTRADA (SEM CB)   
   	     Delete CobrancaRegistrada 
     	      Where DataProcessamento = @DataProcessamento
	        And IdDocto           = @IdDocto

	Else If @TipoDocto = 14       -- COBRANCA ESPECIAL (SEM CB)     
	     Delete CobrancaEspecial
	      Where DataProcessamento = @DataProcessamento
	        And IdDocto           = @IdDocto

	Else If @TipoDocto = 15       -- DARM                           
	     Delete Darm 
	      Where DataProcessamento = @DataProcessamento
	        And IdDocto           = @IdDocto

	Else If @TipoDocto = 16       -- DARF PRETO                     
	     Delete DarfPreto 
	      Where DataProcessamento = @DataProcessamento
	        And IdDocto           = @IdDocto

	Else If @TipoDocto = 17       -- DARF SIMPLES                   
	     Delete DarfSimples 
	      Where DataProcessamento = @DataProcessamento
	        And IdDocto           = @IdDocto

	Else If @TipoDocto = 18       -- GARE                           
	     Delete Gare

	      Where DataProcessamento = @DataProcessamento
	        And IdDocto           = @IdDocto


	-- 19 - GRPS nao existe mais, ver tratamento junto com 11 - INSS

	Else If @TipoDocto IN (20, 21, 22, 23) -- AGUA, GAS, LUZ, TELEFONE
	     Delete ArrecadacaoEletronica
	      Where DataProcessamento = @DataProcessamento
	        And IdDocto           = @IdDocto


	Else If @TipoDocto = 27       -- ARRECADACAO CONVENCIONAL
	     Delete ArrecadacaoConvencional
	      Where DataProcessamento = @DataProcessamento
	        And IdDocto           = @IdDocto

	-- 28 - UNICOBRANCA UBB, 29 - COBRANCA IMEDIATA UBB, 30 - COBRANCA ESPECIAL UBB
      -- 31 - COBRANCA TERCEIROS estao tratados acima na Ficha de Compensacao

	Else If @TipoDocto IN (32,34) -- AJUSTE CREDITO DEPOSITO, CREDITO AUTOMATICO                           
	     Delete AjusteCredito
	      Where DataProcessamento = @DataProcessamento
	        And IdDocto           = @IdDocto

	Else If @TipoDocto IN (33,38) -- AJUSTE DEBITO DEPOSITO, DEBITO AUTOMATICO                           
	     Delete AjusteDebito
	      Where DataProcessamento = @DataProcessamento
	        And IdDocto           = @IdDocto

	Else If @TipoDocto = 35      -- GPS                            
	     Delete GPS
	      Where DataProcessamento = @DataProcessamento
	        And IdDocto           = @IdDocto

	Else If @TipoDocto = 36      -- CARTAOAVULSO                   
	     Delete CartaoAvulso
	      Where DataProcessamento = @DataProcessamento
	        And IdDocto           = @IdDocto

	Else If @TipoDocto = 37      -- OCT                            
 	     Delete OCT
	      Where DataProcessamento = @DataProcessamento
	        And IdDocto           = @IdDocto
  
      -- 38 - DEBITO AUTOMATICO estah tratado acima no AjusteDebito                      
	-- 39 - CAPA OCT, soh apaga na Tabela Documento
	Else If @TipoDocto = 40      -- FGTS
 	     Delete FGTS
	      Where DataProcessamento = @DataProcessamento
	        And IdDocto           = @IdDocto

   Select @Erro           = @@Error
         ,@LinhasAfetadas = @@Rowcount

   -----------------------------------------------------------------------------------------
   -- Se não removeu da Tabela específica e o documento não for 0,1 ou 39, que soh estao na 
   -- tabela Documento, o Tipo de Documento não existe ou ocorreu um erro -> TrataErro
   -----------------------------------------------------------------------------------------
   If (@LinhasAfetadas <> 1 and @TipoDocto not in (0,1,39)) or @Erro <> 0 
	Goto TrataErro

   -------------------------------
   -- Remove da Tabela Documento
   -------------------------------
   Delete Documento 
    Where DataProcessamento = @DataProcessamento
      And IdDocto           = @IdDocto

   Select @Erro = @@Error
         ,@LinhasAfetadas = @@Rowcount

   ---------------------------------------
   -- Se tudo OK, Commit, senao RollBack
   ---------------------------------------
   If @LinhasAfetadas = 1 and @Erro = 0 Begin
		Commit Transaction
   		Return(0)
   End
   

TrataErro:
   Begin
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

GRANT  EXECUTE  ON [dbo].[MDIAG_RemoveLote]  TO [mdi]
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



set nocount on

		/*-----------------------------------
		   INSERT DOS REGISTROS NAS TABELAS 
		-----------------------------------*/
--STATUSCAPA


INSERT INTO StatusCapa VALUES ('0' ,'Capa cadastrada')
INSERT INTO StatusCapa VALUES ('1' ,'Capa digitalizada')
INSERT INTO StatusCapa VALUES ('2' ,'Capa em complementacao')
INSERT INTO StatusCapa VALUES ('3' ,'Capa complementada, mas com pendencia')
INSERT INTO StatusCapa VALUES ('4' ,'Capa para Prova Zero')
INSERT INTO StatusCapa VALUES ('5' ,'Capa para Ilegiveis')
INSERT INTO StatusCapa VALUES ('6' ,'Capa para Alcada')
INSERT INTO StatusCapa VALUES ('7' ,'Capa para Vinculo Manual')
INSERT INTO StatusCapa VALUES ('8' ,'Capa para Vinculo Automatico')
INSERT INTO StatusCapa VALUES ('9' ,'Capa p/ Vinc. Automatico, enviada pelo Prova Zero')
INSERT INTO StatusCapa VALUES ('D' ,'Capa Devolvida pelo Sistema')
INSERT INTO StatusCapa VALUES ('E' ,'Capa Expedida')
INSERT INTO StatusCapa VALUES ('F' ,'Capa Devolvida pelo Caixa Robo')
INSERT INTO StatusCapa VALUES ('G' ,'Capa em Prova Zero')
INSERT INTO StatusCapa VALUES ('H' ,'Capa em Ilegiveis')
INSERT INTO StatusCapa VALUES ('I' ,'Capa em Alcada')
INSERT INTO StatusCapa VALUES ('J' ,'Capa em Vinculo Manual')
INSERT INTO StatusCapa VALUES ('K' ,'Capa em Expedicao')
INSERT INTO StatusCapa VALUES ('O' ,'Capa em Troca de Ordem')
INSERT INTO StatusCapa VALUES ('P' ,'Capa Devolvida pela Preparacao')
INSERT INTO StatusCapa VALUES ('R' ,'Capa para Transmissao')
INSERT INTO StatusCapa VALUES ('S' ,'Capa em Transmissao')
INSERT INTO StatusCapa VALUES ('T' ,'Capa Transmitida')
INSERT INTO StatusCapa VALUES ('V' ,'Capa em Verificacao')
INSERT INTO StatusCapa VALUES ('X' ,'Capa ja enviada a ocorrencia para Ubb')

-- ACAO
INSERT INTO Acao VALUES ('1'   ,'Ilegiveis - Reenviar para Complementacao')
INSERT INTO Acao VALUES ('2'   ,'Ilegiveis - Documento registrado ocorrencia')
INSERT INTO Acao VALUES ('3'   ,'Ilegiveis - Devolver Envelope / Malote')
INSERT INTO Acao VALUES ('4'   ,'Ilegiveis - Enviar para Vinculo Automatico')
INSERT INTO Acao VALUES ('5'   ,'Ilegiveis - Documento Corrigido')
INSERT INTO Acao VALUES ('6'   ,'Ilegiveis - Enviar Capa para Troca de Ordem')
INSERT INTO Acao VALUES ('7'   ,'Ilegiveis - Remover Documento para Recaptura')
INSERT INTO Acao VALUES ('8'   ,'Ilegiveis - Enviar Capa para Recaptura')
INSERT INTO Acao VALUES ('10'  ,'Complementacao - Alterar Tipo de Documento')
INSERT INTO Acao VALUES ('11'  ,'Complementacao - Documento digitado')
INSERT INTO Acao VALUES ('12'  ,'Complementacao - Complementacao Automatica')
INSERT INTO Acao VALUES ('13'  ,'Complementacao - Devolver por Duplicidade (Auto)')
INSERT INTO Acao VALUES ('14'  ,'Complementacao - Cadastrar Envelope / Malote')
INSERT INTO Acao VALUES ('15'  ,'Complementacao - Enviar para Ilegiveis')
INSERT INTO Acao VALUES ('16'  ,'Complementacao - Enviar para Vinculo Auto. (Auto)')
INSERT INTO Acao VALUES ('17'  ,'Complementacao - Enviar para Ilegiveis (Auto)')
INSERT INTO Acao VALUES ('20'  ,'Recepcao - Envelope/Malote recepcionado')
INSERT INTO Acao VALUES ('21'  ,'Recepcao - Envelope/Malote registrado ocorrencia')
INSERT INTO Acao VALUES ('30'  ,'Controle de Qualidade - Remover Envelope / Malote')
INSERT INTO Acao VALUES ('31'  ,'Controle de Qualidade - Remover Documento')
INSERT INTO Acao VALUES ('40'  ,'Captura - Capturar Envelope / Malote')
INSERT INTO Acao VALUES ('41'  ,'Captura - Documento capturado')
INSERT INTO Acao VALUES ('50'  ,'Inicializacao - Criar Parametro')
INSERT INTO Acao VALUES ('51'  ,'Inicializacao - Inicializar Link')
INSERT INTO Acao VALUES ('60'  ,'Prova Zero - Documento corrigido valor')
INSERT INTO Acao VALUES ('61'  ,'Prova Zero - Documento corrigido')
INSERT INTO Acao VALUES ('62'  ,'Prova Zero - Devolver Envelope / Malote')
INSERT INTO Acao VALUES ('63'  ,'Prova Zero - Enviar para Ilegiveis')
INSERT INTO Acao VALUES ('64'  ,'Prova Zero - Enviar para Vinc. Auto. C/ Alt. Valor')
INSERT INTO Acao VALUES ('65'  ,'Prova Zero - Enviar para Vinc. Auto. Apos Conferir')
INSERT INTO Acao VALUES ('66'  ,'Prova Zero - Enviar para Troca de Ordem.')
INSERT INTO Acao VALUES ('70'  ,'Vinculo Manual - Enviar para Ilegiveis')
INSERT INTO Acao VALUES ('71'  ,'Vinculo Manual - Documento registrado ocorrencia')
INSERT INTO Acao VALUES ('72'  ,'Vinculo Manual - Documento vinculado manualmente')
INSERT INTO Acao VALUES ('73'  ,'Vinculo Manual - Inserir Ajuste Debito / Credito')
INSERT INTO Acao VALUES ('74'  ,'Vinculo Manual - Enviar para Alcada(Automatico)')
INSERT INTO Acao VALUES ('75'  ,'Vinculo Manual - Enviar para Transmissao(Auto)')
INSERT INTO Acao VALUES ('76'  ,'Vinculo Manual - Enviar para Analise')
INSERT INTO Acao VALUES ('80'  ,'Expedicao - Documento autenticado')
INSERT INTO Acao VALUES ('81'  ,'Expedicao - Reautenticar Documento')
INSERT INTO Acao VALUES ('82'  ,'Expedicao - Imprimir Ocorr. do Envelope / Malote')
INSERT INTO Acao VALUES ('83'  ,'Expedicao - Imprimir Ocorrencia do Documento')
INSERT INTO Acao VALUES ('84'  ,'Expedicao - Imprimir Comp. Ajuste Debito / Credito')
INSERT INTO Acao VALUES ('85'  ,'Expedicao - Imprimir Comprovante de Deposito')
INSERT INTO Acao VALUES ('86'  ,'Expedicao - Atualizar Env. / Mal. Expedido(Auto)')
INSERT INTO Acao VALUES ('87'  ,'Expedicao - Imprimir Comp. Cartao Avulso')
INSERT INTO Acao VALUES ('88'  ,'Expedicao - Entrar Capa')
INSERT INTO Acao VALUES ('89'  ,'Expedicao - Imprimir Comp. Pagamento')
INSERT INTO Acao VALUES ('90'  ,'Alcada - Documento liberado com alcada')
INSERT INTO Acao VALUES ('91'  ,'Alcada - Enviar Envelope / Malote p/ Trans.(Auto)')
INSERT INTO Acao VALUES ('92'  ,'Alcada - Enviar para Ilegiveis')
INSERT INTO Acao VALUES ('100' ,'Supervisao - Excluir Envelope / Malote')
INSERT INTO Acao VALUES ('110' ,'Vinc. Automatico - Enviar para Prova Zero')
INSERT INTO Acao VALUES ('111' ,'Vinc. Automatico - Enviar para Ilegiveis')
INSERT INTO Acao VALUES ('112' ,'Vinc. Automatico - Enviar para Alcada')
INSERT INTO Acao VALUES ('113' ,'Vinc. Automatico - Enviar para Vinculo Manual')
INSERT INTO Acao VALUES ('114' ,'Vinc. Automatico - Enviar para Transmissao')
INSERT INTO Acao VALUES ('115' ,'Vinc. Automatico - Enviar para Analise')
INSERT INTO Acao VALUES ('120' ,'Robo - Inicializacao')
INSERT INTO Acao VALUES ('121' ,'Robo - Transmite Capa')
INSERT INTO Acao VALUES ('122' ,'Robo - Transmite Documento')
INSERT INTO Acao VALUES ('123' ,'Robo - Grava Ocorrencia')
INSERT INTO Acao VALUES ('124' ,'Robo - Capa com Diferenca enviada para Ilegiveis')
INSERT INTO Acao VALUES ('130' ,'Troca de Ordem - Documento inserido')
INSERT INTO Acao VALUES ('131' ,'Troca de Ordem - Documento excluido')
INSERT INTO Acao VALUES ('132' ,'Troca de Ordem - Envio Env/Mal para Vinculo Aut.')
INSERT INTO Acao VALUES ('133' ,'Troca de Ordem - Reenvio para Ilegiveis')
INSERT INTO Acao VALUES ('134' ,'Troca de Ordem - Documento reordenado')
INSERT INTO Acao VALUES ('140' ,'Consulta - Consultar Capa')
INSERT INTO Acao VALUES ('150' ,'Complementacao - Split Capa (Inicial)')
INSERT INTO Acao VALUES ('151' ,'Complementacao - Split Capa (Final)')
INSERT INTO Acao VALUES ('152' ,'Complementacao - Split Capa Anterior (Inicial)')
INSERT INTO Acao VALUES ('153' ,'Complementacao - Split Capa Anterior (Final)')
INSERT INTO Acao VALUES ('190' ,'Ilegiveis - Enviar para Analise')
INSERT INTO Acao VALUES ('191' ,'Ilegiveis - Selecionar Capa')
INSERT INTO Acao VALUES ('192' ,'Ilegiveis - Deselecionar Capa')
INSERT INTO Acao VALUES ('193' ,'Prova Zero - Selecionar Capa')
INSERT INTO Acao VALUES ('194' ,'Prova Zero - Deselecionar Capa')
INSERT INTO Acao VALUES ('195' ,'Vinc. Manual - Selecionar Capa')
INSERT INTO Acao VALUES ('196' ,'Vinc. Manual - Deselecionar Capa')
INSERT INTO Acao VALUES ('197' ,'Alcada - Selecionar Capa')
INSERT INTO Acao VALUES ('198' ,'Alcada - Deselecionar Capa')

-- OCORRENCIA

INSERT INTO Ocorrencia VALUES ('1'   ,'Quantidade de Envelope maior que informada no Protocolo Remessa')
INSERT INTO Ocorrencia VALUES ('2'   ,'Quantidade de Envelope menor que informada Protocolo Remessa')
INSERT INTO Ocorrencia VALUES ('3'   ,'Envelope Vazio')
INSERT INTO Ocorrencia VALUES ('4'   ,'Envelope so com Dinheiro')
INSERT INTO Ocorrencia VALUES ('5'   ,'Envelope so com Cheque')
INSERT INTO Ocorrencia VALUES ('6'   ,'Envelope so com o Documento a ser Pago')
INSERT INTO Ocorrencia VALUES ('7'   ,'Envelope so com  Ficha de Deposito')
INSERT INTO Ocorrencia VALUES ('101' ,'Deposito em Dinheiro')
INSERT INTO Ocorrencia VALUES ('102' ,'Cheque para Deposito em Dinheiro')
INSERT INTO Ocorrencia VALUES ('104' ,'Deposito Dinheiro Agencia/Conta Invalida')
INSERT INTO Ocorrencia VALUES ('105' ,'Deposito em Cheque para Agencia/Conta Invalida')
INSERT INTO Ocorrencia VALUES ('106' ,'Deposito Misto em Cheque e em Dinheiro')
INSERT INTO Ocorrencia VALUES ('107' ,'Deposito em Poupanca, apos horario corte')
INSERT INTO Ocorrencia VALUES ('108' ,'Deposito em Cheque c/ desdobramento varias fichas de deposito')
INSERT INTO Ocorrencia VALUES ('109' ,'Deposito Cheque c/ Irregularidade ( sem Assinatura, Erro  Preenchimento ,Validade')
INSERT INTO Ocorrencia VALUES ('110' ,'Deposito em C/C Inexistente / Encerrada')
INSERT INTO Ocorrencia VALUES ('111' ,'Deposito em Conta Corrente na situacao  2025')
INSERT INTO Ocorrencia VALUES ('112' ,'Deposito em C/C na condicao CC5 - Residentes no Exterior')
INSERT INTO Ocorrencia VALUES ('113' ,'Deposito em Conta Corrente Paralisada')
INSERT INTO Ocorrencia VALUES ('114' ,'Deposito em C/C na condicao MR')
INSERT INTO Ocorrencia VALUES ('115' ,'Deposito em C/C  bloqueada')
INSERT INTO Ocorrencia VALUES ('120' ,'Deposito com Moedas')
INSERT INTO Ocorrencia VALUES ('121' ,'Deposito a maior - Valor do cheque menor que valor informado na ficha de deposito')
INSERT INTO Ocorrencia VALUES ('122' ,'Deposito a menor - Valor do cheque maior que valor informado na ficha de deposito')
INSERT INTO Ocorrencia VALUES ('123' ,'Deposito com Cheque UBB')
INSERT INTO Ocorrencia VALUES ('124' ,'Deposito com Cheque de Outros Bancos')
INSERT INTO Ocorrencia VALUES ('201' ,'Pagto em Dinheiro')
INSERT INTO Ocorrencia VALUES ('202' ,'Pagto Conta/Titulo com Cheque e Dinheiro')
INSERT INTO Ocorrencia VALUES ('203' ,'Pagto Conta/Titulo com Cheque valor menor ao valor documento a ser liquidado')
INSERT INTO Ocorrencia VALUES ('204' ,'Pagto Conta/Titulo com Cheque valor maior  ao valor do documento a ser liquidado')
INSERT INTO Ocorrencia VALUES ('205' ,'Pagto Conta/Titulo Quitada')
INSERT INTO Ocorrencia VALUES ('206' ,'Pagto Conta/Titulo nao aceito pelo Unibanco')
INSERT INTO Ocorrencia VALUES ('207' ,'Pagto Conta/Titulo nao aceito pelo UBB c/ Conta/Titulo aceito UBB mesmo envelope')
INSERT INTO Ocorrencia VALUES ('208' ,'Pagto Titulo Vencido')
INSERT INTO Ocorrencia VALUES ('209' ,'Pagto Contas Cheque UBB ou Outros Bancos, sem descricao Finalidade no verso')
INSERT INTO Ocorrencia VALUES ('210' ,'Pagto varios Documentos UBB com apenas um cheque de outro Banco')
INSERT INTO Ocorrencia VALUES ('211' ,'Pagto apenas uma Conta/Titulo com diversos cheques de Outros Bancos')
INSERT INTO Ocorrencia VALUES ('212' ,'Pagto Conta/Titulo emitido por terceiros a ser Pago Cheque Outro Banco')
INSERT INTO Ocorrencia VALUES ('213' ,'Pagto Conta/Titulo Cheque UBB c/Insuficiencia Saldo')
INSERT INTO Ocorrencia VALUES ('214' ,'Irregularidade no numero DARF')
INSERT INTO Ocorrencia VALUES ('215' ,'Pagto Conta/Titulo c/ irreguralaridade Cheque (sem assinatura, erro preenchimento')
INSERT INTO Ocorrencia VALUES ('216' ,'Pagto Conta/Titulo com Cheque de conta corrente Encerrada/Inexistente')
INSERT INTO Ocorrencia VALUES ('217' ,'Pagto Conta/Titulo com cheque de conta corrente na situacao 2025')
INSERT INTO Ocorrencia VALUES ('218' ,'Pagto Conta/Titulo com Cheque Conta Corrente condicao CC5 - Residente Exterior')
INSERT INTO Ocorrencia VALUES ('219' ,'Pagto Conta/Titulo com Cheque Conta Corrente Paralisada')
INSERT INTO Ocorrencia VALUES ('220' ,'Pagto Conta/Titulo com Cheque Conta Corrente na Condicao MR')
INSERT INTO Ocorrencia VALUES ('221' ,'Pagto Conta/Titulo com Cheque Conta Corrente na Condicao CL/AD')
INSERT INTO Ocorrencia VALUES ('222' ,'Pagto Conta/Titulo com Cheque c/ Conta com Saldo Bloqueado')
INSERT INTO Ocorrencia VALUES ('223' ,'Pagto Conta/Titulo c/ Cheque,  Conta Saldo Bloqueado e fora Limite Contratual')
INSERT INTO Ocorrencia VALUES ('224' ,'Pagto Conta faltando via ou Titulo com apenas uma via')
INSERT INTO Ocorrencia VALUES ('226' ,'Pagto Conta/Titulo sem Cheque')
INSERT INTO Ocorrencia VALUES ('227' ,'Pagto Conta/Titulo documentos cadastrados em Debito Automatico')
INSERT INTO Ocorrencia VALUES ('228' ,'Pagto Conta/Titulo com Cheque Bloqueado motivo 29 - sustado')
INSERT INTO Ocorrencia VALUES ('229' ,'Pagto Conta/Titulo c/ cheque sustado.')
INSERT INTO Ocorrencia VALUES ('234' ,'Irregularidade no num./codigo do documento')
INSERT INTO Ocorrencia VALUES ('301' ,'Retirada da Poupanca')
INSERT INTO Ocorrencia VALUES ('302' ,'DOC')
INSERT INTO Ocorrencia VALUES ('303' ,'Requisicao de Talao de Cheque')
INSERT INTO Ocorrencia VALUES ('304' ,'Nao e envelope do caixa expresso')
INSERT INTO Ocorrencia VALUES ('306' ,'Resumo de Vendas de Cartao de Credito')
INSERT INTO Ocorrencia VALUES ('401' ,'Pagto Conta/Titulo, por Autor. de Debito, assinatura nao confere.')
INSERT INTO Ocorrencia VALUES ('402' ,'Pagto Conta/Titulo, por Autor. de Debito, com valor a menor')
INSERT INTO Ocorrencia VALUES ('403' ,'Pagto Conta/Titulo, por Autor. de Debito, com valor a maior')
INSERT INTO Ocorrencia VALUES ('404' ,'Pagto Conta/Titulo Nao Aceito Unibanco,   atraves de Autorizacao de Debito')
INSERT INTO Ocorrencia VALUES ('405' ,'Pagto Conta/Titulo, sem assinatura na Autorizacao de Debito')
INSERT INTO Ocorrencia VALUES ('406' ,'Pagto Conta/Titulo, sem descrever o valor a ser debitado na Autorizacao de Debito')
INSERT INTO Ocorrencia VALUES ('407' ,'Pagto Conta/Titulo, com Cheque Terceiro e Autor. de Debito')
INSERT INTO Ocorrencia VALUES ('408' ,'Pagto Conta/Titulo, com Cheque UBB e Autor de Debito')
INSERT INTO Ocorrencia VALUES ('409' ,'Pagto Conta/Titulo, com Dinheiro e Autor. de Debito')
INSERT INTO Ocorrencia VALUES ('411' ,'Pagto Conta/Titulo Vencidos, atraves de Autorizacao de Debito')
INSERT INTO Ocorrencia VALUES ('412' ,'Pagto Conta/Titulo Quitado, atraves Autorizacao de Debito')
INSERT INTO Ocorrencia VALUES ('413' ,'Pagto Conta/Titulo nao aceito junto com outro aceito, por Autor. Debito')
INSERT INTO Ocorrencia VALUES ('414' ,'Pagto Conta/Titulo, atraves de Autorizacao de Debito, com Conta Invalida')
INSERT INTO Ocorrencia VALUES ('415' ,'Pagto Conta/Titulo, atraves de Autorizacao de Debito, sem assinatura capturada')
INSERT INTO Ocorrencia VALUES ('416' ,'Pagto Conta/Titulo c/ Autor. de Debito e  Cheque, ambos valor igual total')
INSERT INTO Ocorrencia VALUES ('417' ,'Autorizacao de Debito p/ Deposito')
INSERT INTO Ocorrencia VALUES ('419' ,'Pagto Conta/Titulo, por Autor. de Debito, conta corrente Encerrada/Inexistente')
INSERT INTO Ocorrencia VALUES ('420' ,'Pagto Conta/Titulo, por Autor. de Debito, conta corrente na situacao 2025.')
INSERT INTO Ocorrencia VALUES ('421' ,'Pagto Conta/Titulo, por Autor. de Debito, C/C condicao CC5 - Residente Exterior')
INSERT INTO Ocorrencia VALUES ('422' ,'Pagto Conta/Titulo, por Autor. de Debito, Conta Corrente Paralisada')
INSERT INTO Ocorrencia VALUES ('424' ,'Pagto Conta/Titulo, por Autor. de Debito, Conta Corrente na Condicao CL/AD')
INSERT INTO Ocorrencia VALUES ('425' ,'Pagto Conta/Titulo, por Autor. de Debito, C/C  bloqueada')
INSERT INTO Ocorrencia VALUES ('426' ,'Pagto Conta/Titulo, por Autor. de Debito, c/ Conta com Saldo Bloqueado.')
INSERT INTO Ocorrencia VALUES ('429' ,'Pagto Conta/Titulo, por Autor. de Debito, c/Insuficiencia Saldo.')
INSERT INTO Ocorrencia VALUES ('430' ,'Autorizacao de Debito para assinatura inexistente.')
INSERT INTO Ocorrencia VALUES ('431' ,'Autorizacao de Debito Acima de R$2.000,00.')
INSERT INTO Ocorrencia VALUES ('432' ,'Autorizacao de Debito para Pessoa Juridica.')
INSERT INTO Ocorrencia VALUES ('998' ,'Documento em Duplicidade')
INSERT INTO Ocorrencia VALUES ('999' ,'Erro Operacional')


-- PARAMETRO
INSERT INTO Parametro
       (DataProcessamento,
	Hm_Abertura,
	AgenciaCentral,
	AgenciaApresentante,
	Tm_Pendente,
	Tm_Atualizacao,
	Dir_Dados,
	Dir_Imagens,
	Dir_Trabalho
	)
VALUES
       (20001127, 		--DataProcessamento
	GETDATE(),		--Hm_Abertura
	9999,			--AgenciaCentral
	9999,			--AgenciaApresentante
	300,			--Tm_Pendente
	30,			--Tm_Atualizacao
	'c:\mdi_ag\dados',	--Dir_Dados
	'c:\mdi_ag\imagens',	--Dir_Imagens
	'c:\mdi_ag\trabalho\'	--Dir_Trabalho
	)


-- STATUSDOCUMENTO
INSERT INTO StatusDocumento VALUES ('0'   ,'Documento nao complementado')
INSERT INTO StatusDocumento VALUES ('1'   ,'Documento complementado')
INSERT INTO StatusDocumento VALUES ('2'   ,'Documento em transmissao')
INSERT INTO StatusDocumento VALUES ('3'   ,'Documento com ocorrencia em transmissao')
INSERT INTO StatusDocumento VALUES ('A'   ,'Documento para Alcada')
INSERT INTO StatusDocumento VALUES ('D'   ,'Documento deletado pelo sistema')
INSERT INTO StatusDocumento VALUES ('E'   ,'Documento Expedido')
INSERT INTO StatusDocumento VALUES ('F'   ,'Documento deletado pelo caixa robo')
INSERT INTO StatusDocumento VALUES ('G'   ,'Acerto de Debito/Credito gerado por diferenca')
INSERT INTO StatusDocumento VALUES ('T'   ,'Documento Transmitido')

--STATUSLOTE
INSERT INTO StatusLote VALUES ('0'   ,'Lote digitalizado')
INSERT INTO StatusLote VALUES ('1'   ,'Lote em liberacao')
INSERT INTO StatusLote VALUES ('2'   ,'Lote liberado')
INSERT INTO StatusLote VALUES ('3'   ,'Lote em Captura')


--TIPODOCTO
INSERT INTO TipoDocto VALUES ('0'   ,'DOCUMENTO INDEFINIDO')
INSERT INTO TipoDocto VALUES ('1'   ,'CAPA ENVELOPE / MALOTE EMPRESA')
INSERT INTO TipoDocto VALUES ('2'   ,'DEPOSITO CONTA CORRENTE')
INSERT INTO TipoDocto VALUES ('3'   ,'DEPOSITO CONTA POUPANCA')
INSERT INTO TipoDocto VALUES ('4'   ,'AUTORIZACAO DE DEBITO EM C/C')
INSERT INTO TipoDocto VALUES ('5'   ,'CHEQUE UBB SACADO')
INSERT INTO TipoDocto VALUES ('6'   ,'CHEQUE COMPENSADO (CP)')
INSERT INTO TipoDocto VALUES ('7'   ,'CHEQUE DEPOSITO')
INSERT INTO TipoDocto VALUES ('8'   ,'CONCESSIONARIA VALOR REAL')
INSERT INTO TipoDocto VALUES ('9'   ,'CONCESSIONARIA VALOR INDEXADO')
INSERT INTO TipoDocto VALUES ('10'  ,'FICHA COMPENSACAO')
INSERT INTO TipoDocto VALUES ('11'  ,'INSS')
INSERT INTO TipoDocto VALUES ('12'  ,'TIT. OUTROS BCOS CONVENCIONAL')
INSERT INTO TipoDocto VALUES ('13'  ,'COBRANCA REGISTRADA (SEM CB)')
INSERT INTO TipoDocto VALUES ('14'  ,'COBRANCA ESPECIAL (SEM CB)')
INSERT INTO TipoDocto VALUES ('15'  ,'DARM')
INSERT INTO TipoDocto VALUES ('16'  ,'DARF PRETO')
INSERT INTO TipoDocto VALUES ('17'  ,'DARF SIMPLES')
INSERT INTO TipoDocto VALUES ('18'  ,'GARE')
INSERT INTO TipoDocto VALUES ('19'  ,'GRPS')
INSERT INTO TipoDocto VALUES ('20'  ,'AGUA')
INSERT INTO TipoDocto VALUES ('21'  ,'GAS')
INSERT INTO TipoDocto VALUES ('22'  ,'LUZ')
INSERT INTO TipoDocto VALUES ('23'  ,'TELEFONE')
INSERT INTO TipoDocto VALUES ('24'  ,'TRIBUTOS MUNICIPAIS')
INSERT INTO TipoDocto VALUES ('25'  ,'TRIBUTOS ESTADUAIS')
INSERT INTO TipoDocto VALUES ('26'  ,'TRIBUTOS FEDERAIS')
INSERT INTO TipoDocto VALUES ('27'  ,'ARRECADACAO CONVENCIONAL')
INSERT INTO TipoDocto VALUES ('28'  ,'UNICOBRANCA UBB')
INSERT INTO TipoDocto VALUES ('29'  ,'COBRANCA IMEDIATA UBB')
INSERT INTO TipoDocto VALUES ('30'  ,'COBRANCA ESPECIAL UBB')
INSERT INTO TipoDocto VALUES ('31'  ,'TITULO OUTROS BCOS ELETRONICO')
INSERT INTO TipoDocto VALUES ('32'  ,'AJUSTE CREDITO')
INSERT INTO TipoDocto VALUES ('33'  ,'AJUSTE DEBITO')
INSERT INTO TipoDocto VALUES ('34'  ,'CREDITO AUTOMATICO')
INSERT INTO TipoDocto VALUES ('35'  ,'GPS')
INSERT INTO TipoDocto VALUES ('36'  ,'CARTAO AVULSO')
INSERT INTO TipoDocto VALUES ('37'  ,'OCT')
INSERT INTO TipoDocto VALUES ('38'  ,'DEBITO AUTOMATICO')
INSERT INTO TipoDocto VALUES ('39'  ,'CAPA OCT')
INSERT INTO TipoDocto VALUES ('40'  ,'FGTS')
INSERT INTO TipoDocto VALUES ('41'  ,'LANCAMENTO INTERNO')
INSERT INTO TipoDocto VALUES ('42'  ,'AJUSTE CONTABIL RECEITA')
INSERT INTO TipoDocto VALUES ('43'  ,'AJUSTE CONTABIL DESPESA')


set nocount off