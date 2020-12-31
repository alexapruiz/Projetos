SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_AlteraTipoDocto]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_AlteraTipoDocto]
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

