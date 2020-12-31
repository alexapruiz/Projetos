SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_GetAgenciasCapa]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_GetAgenciasCapa]
GO


/****** Object:  Stored Procedure dbo.MDIAG_GetAgenciasCapa    Script Date: 17/11/00 11:22:10 ******/

/****** Object:  Stored Procedure dbo.GetAgenciasCapa    Script Date: 03/11/00 18:02:43 ******/
/****** Object:  Stored Procedure dbo.GetAgenciasCapa    Script Date: 14/09/00 13:02:06 ******/
/****** Object:  Stored Procedure dbo.GetAgenciasCapa    Script Date: 15/08/00 13:54:53 ******/
CREATE PROCEDURE MDIAG_GetAgenciasCapa

	@Capa		Numeric(18),
	@DataProc	Int,
	@AgOrig		Smallint,
	@idCapa		Int

AS

	if @AgOrig Is Null
		SELECT	Cap.IdCapa,
			Cap.AgOrig,
			Cap.IdLote,
			Cap.Status,
			Stc.Descricao,
			Cap.Num_Malote,
			Ocorrencia,
			Cap.Idenv_Mal
		  FROM	Capa Cap, StatusCapa Stc 
		 WHERE	Cap.Capa		= @Capa
		   And  Cap.dataprocessamento	= @dataProc
		   And	Cap.Status             	= Stc.Status
		ORDER BY Cap.AgOrig

	Else
		SELECT	Cap.IdCapa,
			Cap.AgOrig,
			Cap.IdLote,
			Cap.Status,
			Stc.Descricao,
			Cap.Num_Malote,
			Ocorrencia,
			Cap.IdEnv_Mal
		  FROM	Capa Cap, StatusCapa Stc 
		 WHERE 	Cap.IdCapa		= @idCapa
		   And	Cap.dataprocessamento	= @dataProc		 
		   And	Cap.Status		= Stc.Status
		ORDER BY Cap.AgOrig




GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

