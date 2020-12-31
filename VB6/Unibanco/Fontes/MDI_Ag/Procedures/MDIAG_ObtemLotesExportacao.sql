SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_ObtemLotesExportacao]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_ObtemLotesExportacao]
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

