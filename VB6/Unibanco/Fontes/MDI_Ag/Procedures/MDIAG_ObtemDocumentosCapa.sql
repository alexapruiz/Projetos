SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_ObtemDocumentosCapa]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_ObtemDocumentosCapa]
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

