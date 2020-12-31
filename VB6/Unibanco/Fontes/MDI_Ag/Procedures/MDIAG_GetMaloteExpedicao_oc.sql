SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_GetMaloteExpedicao_oc]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_GetMaloteExpedicao_oc]
GO


/****** Object:  Stored Procedure dbo.MDIAG_GetMaloteExpedicao_oc    Script Date: 17/11/00 11:22:10 ******/
CREATE PROCEDURE MDIAG_GetMaloteExpedicao_oc
	@DataProc	Int,
	@NumMalote	Numeric(11)

As

	SELECT	Capa,IdCapa, IdLote, IdEnv_Mal,Idcapa, Num_Malote, 
		AgOrig,Status, Status as StatusAnt, 
		DateDiff(second, HoraAtual, GetDate()) as Intervalo,
		IsNull(Ocorrencia, 0) as Ocorrencia
	  FROM	Capa
	 WHERE 	DataProcessamento 	= @DataProc
	   And	Num_Malote		= @NumMalote
	   And	Status in 		('0','P')
	 ORDER	By Capa,AgOrig




GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

