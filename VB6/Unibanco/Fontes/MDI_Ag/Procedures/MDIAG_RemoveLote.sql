SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_RemoveLote]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_RemoveLote]
GO


/****** Object:  Stored Procedure dbo.MDIAG_RemoveLote    Script Date: 17/11/00 11:22:10 ******/
CREATE Proc MDIAG_RemoveLote
	@DataProcessamento	int,
	@IdLote 			int 

as

	Begin Transaction

	--------------------------------------------------------------------------------------------------
	-- Deleta Lote
	DELETE FROM Lote
	 WHERE DataProcessamento = @DataProcessamento
	   And IdLote            = @IdLote
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

