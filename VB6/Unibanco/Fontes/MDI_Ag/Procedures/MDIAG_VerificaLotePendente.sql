SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_VerificaLotePendente]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_VerificaLotePendente]
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

