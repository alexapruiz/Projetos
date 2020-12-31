SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_GetTotalCapa]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_GetTotalCapa]
GO

CREATE Procedure MDIAG_GetTotalCapa
	@DataProc   	int,
	@IdEnv_Mal  	char(1)

As

if @IdEnv_Mal = 'T' Begin

	Select Count(*) as Total
	  From Capa (NOLOCK)
	 Where DataProcessamento = @DataProc

End
Else Begin

	Select Count(*) as Total
	  From Capa (NOLOCK)
	 Where DataProcessamento = @DataProc
	   And IdEnv_Mal         = @IdEnv_Mal

End






GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

