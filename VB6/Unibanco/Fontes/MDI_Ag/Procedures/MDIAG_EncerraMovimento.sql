SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_EncerraMovimento]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_EncerraMovimento]
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

