SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_GetTodasCapas_Oc]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_GetTodasCapas_Oc]
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

