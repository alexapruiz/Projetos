SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_GetRecuperaStatus]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_GetRecuperaStatus]
GO


/****** Object:  Stored Procedure dbo.MDIAG_GetRecuperaStatus    Script Date: 17/11/00 11:22:10 ******/
CREATE PROCEDURE MDIAG_GetRecuperaStatus

	@DataProc	Int,
	@NumMalote	Numeric(11),
	@Idcapa	Int,
	@Capa		Numeric(18)

As

	If @Capa is null 
		Select 	Capa.IdCapa, Capa.IdLote, Capa.IdEnv_Mal, Capa.Capa,Idcapa, Capa.Num_Malote, 
			Capa.AgOrig, Capa.Status, Capa.Status as StatusAnt,status.Descricao ,
			DateDiff(second, HoraAtual, GetDate()) as Intervalo,
			IsNull(Ocorrencia, 0) as Ocorrencia
		From	Capa Capa, StatusCapa Status
		Where 	Capa.DataProcessamento 	= @DataProc  	And
			Capa.Num_Malote        	= @NumMalote 	And
			Capa.Status            		= Status.Status 	And
			Capa.Idcapa		= @Idcapa   
		Order 	By Capa.Capa, Capa.AgOrig
	Else
		Select 	Capa.IdCapa, Capa.IdLote, Capa.IdEnv_Mal, Capa.Capa,Idcapa, Capa.Num_Malote, 
			Capa.AgOrig, Capa.Status, Capa.Status as StatusAnt,status.Descricao ,
			DateDiff(second, HoraAtual, GetDate()) as Intervalo,
			IsNull(Ocorrencia, 0) as Ocorrencia
		From	Capa Capa, StatusCapa Status
		Where	Capa.DataProcessamento 	= @DataProc  	And
			Capa.Capa  		= @Capa	And
			Capa.Status		= Status.Status
		Order	By Capa.Capa, Capa.AgOrig



GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

