SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_InsereLog]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_InsereLog]
GO


/****** Object:  Stored Procedure dbo.MDIAG_InsereLog    Script Date: 17/11/00 11:22:09 ******/
CREATE Procedure MDIAG_InsereLog
	@DataProc int,
	@IdCapa   int,
	@IdDocto  int,
	@Usuario  Char(10),
	@Acao     tinyint
As

Begin Transaction
	Insert Into Log
		(DataProcessamento,
		 IdCapa,
		 IdDocto,
		 Data,
		 Login,
		 Acao)
	Values
		(@DataProc,
		 @IdCapa,
		 @IdDocto,
		 GetDate(),
		 @Usuario,
		 @Acao)

if @@Error = 0 begin
	Commit Transaction
end
else begin
	Rollback Transaction
end



GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

