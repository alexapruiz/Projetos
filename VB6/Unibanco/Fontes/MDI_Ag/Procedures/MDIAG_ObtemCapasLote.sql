SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_ObtemCapasLote]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_ObtemCapasLote]
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

