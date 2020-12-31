SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_GetDocumentoContQualidade]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_GetDocumentoContQualidade]
GO


/****** Object:  Stored Procedure dbo.MDIAG_GetDocumentoContQualidade    Script Date: 17/11/00 11:22:10 ******/
CREATE PROCEDURE MDIAG_GetDocumentoContQualidade
	@DataProc	Int,
	@IdLote		Int
As


	SELECT	D.IdDocto, D.IdCapa, D.TipoDocto, D.Frente, D.Verso, IsNull(D.Leitura, '') AS Leitura, 
		D.Ordem, C.IdEnv_Mal, D.Status
	  FROM	Documento D, Capa C (NOLOCK)
	 WHERE	D.DataProcessamento 	= @DataProc
	   And	D.IdCapa		= C.IdCapa
	   And	C.IdLote		= @IdLote
	 ORDER BY IdDocto






GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

