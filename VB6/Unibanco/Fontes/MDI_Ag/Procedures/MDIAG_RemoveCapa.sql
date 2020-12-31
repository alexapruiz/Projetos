SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_RemoveCapa]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_RemoveCapa]
GO


/****** Object:  Stored Procedure dbo.MDIAG_RemoveCapa    Script Date: 17/11/00 11:22:10 ******/
CREATE PROCEDURE MDIAG_RemoveCapa
	@DataProcessamento	Int,
	@IdCapa		Int

As

	Begin Transaction

	--------------------------------------------------------------------------------------------------
	-- Deleta Capa
	DELETE 	FROM Capa
	 WHERE	DataProcessamento 	= @DataProcessamento
	   And  IdCapa			= @IdCapa
	if @@Error = 0 Begin
		Commit Transaction
		Return(0)
	End
	Else Begin
		RollBack Transaction
		Return(1)
	End
	-------------------------------------------------------------------------------------------------- 



GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

