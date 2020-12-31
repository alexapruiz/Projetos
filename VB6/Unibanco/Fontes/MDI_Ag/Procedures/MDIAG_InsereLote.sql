SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_InsereLote]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_InsereLote]
GO


/****** Object:  Stored Procedure dbo.MDIAG_InsereLote    Script Date: 17/11/00 11:22:09 ******/
CREATE Proc MDIAG_InsereLote
	@DataProcessamento		int,
	@Prioridade			smallint,
	@Agencia			int,
	@IdLote 			int Output

as

	--------------------------------------------------------------------------------------------------
	-- Busca ultimo Lote da Agenica
	Select 	@IdLote = IsNull((Max(IdLote)+1),
		@Agencia * 100000 + 1)
	  From	Lote 
	 Where	IdLote between (@Agencia * 100000)
	   And ((@Agencia + 1) * 100000)
	   And  DataProcessamento = @DataProcessamento

	--------------------------------------------------------------------------------------------------
	-- Grava Lote
	Insert Lote (  DataProcessamento, IdLote, Status, Prioridade )
	     Values ( @DataProcessamento, @IdLote, '3', @Prioridade )
	if @@Error = 0
	Begin
		Return(0)
	End
	Else Begin
		select @IdLote = -1
		Return(1)
	End
	--------------------------------------------------------------------------------------------------




GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

