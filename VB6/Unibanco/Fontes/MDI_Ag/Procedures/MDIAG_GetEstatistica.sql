SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_GetEstatistica]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_GetEstatistica]
GO

CREATE Procedure MDIAG_GetEstatistica
@DataProc		int,
@IdEnv_Mal		char(1)

As

if @IdEnv_Mal = 'T' Begin   -- Todos

	SELECT 	Count(Distinct(C.IdCapa)) as QtdCapa, Count(D.IdDocto) as QtdDoc, C.Status
	  FROM 	Capa C (NOLOCK) Left Outer Join Documento D (NOLOCK INDEX = IndDataIdCapa) on (D.IdCapa = C.IdCapa And
		D.DataProcessamento = C.DataProcessamento) 
	 WHERE	C.DataProcessamento = @DataProc
	 GROUP BY C.Status

End
Else Begin

	SELECT 	Count(Distinct(C.IdCapa)) as QtdCapa, Count(D.IdDocto) as QtdDoc, C.Status
	  FROM 	Capa C (NOLOCK) Left Outer Join Documento D (NOLOCK INDEX = IndDataIdCapa) on (D.IdCapa = C.IdCapa And
		D.DataProcessamento = C.DataProcessamento)
	 WHERE	C.DataProcessamento 	= @DataProc
	   And	C.IdEnv_Mal 		= @IdEnv_Mal
	 GROUP BY C.Status

End





GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

