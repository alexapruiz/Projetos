SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_AtualizaStatusLote]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_AtualizaStatusLote]
GO


/****** Object:  Stored Procedure dbo.MDIAG_AtualizaStatusLote    Script Date: 17/11/00 11:22:09 ******/
CREATE Procedure MDIAG_AtualizaStatusLote
	 @DataProcessamento		int
	,@IdLote             		int
	,@Status             		char(1)

As

Declare @Erro           int
Declare @LinhasAfetadas int

Begin Transaction

	Update Lote Set
		Status 			= @Status,
		HoraAtual 		= GetDate()
	 Where 	DataProcessamento	= @DataProcessamento
		And IdLote              = @IdLote

	Select 	@Erro 		= @@Error,
		@LinhasAfetadas = @@RowCount

	If @LinhasAfetadas <> 1 or @Erro <> 0 
		Begin 	
			RollBack Transaction
   			Return(1)	
		End

Commit Transaction
Return(0)

GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

