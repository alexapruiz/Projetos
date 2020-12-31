SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_ObtemLog]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_ObtemLog]
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

