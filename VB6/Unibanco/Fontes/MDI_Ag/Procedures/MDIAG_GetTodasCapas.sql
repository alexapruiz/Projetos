SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_GetTodasCapas]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_GetTodasCapas]
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

