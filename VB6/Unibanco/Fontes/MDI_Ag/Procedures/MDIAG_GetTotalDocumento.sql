SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_GetTotalDocumento]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_GetTotalDocumento]
GO

CREATE Procedure MDIAG_GetTotalDocumento
	@DataProc   	int,
	@IdEnv_Mal  	char(1)

As

if @IdEnv_Mal = 'T' Begin

	Select Count(*) as Total 
	  From Documento (NOLOCK)
	 Where DataProcessamento = @DataProc

End
Else Begin

	Select Count(*) as Total 
	  From Documento D, Capa C (NOLOCK)
	 Where D.DataProcessamento = @DataProc
	   And C.DataProcessamento = @DataProc
	   And D.IdCapa            = C.IdCapa
	   And C.IdEnv_Mal         = @IdEnv_Mal

End






GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

