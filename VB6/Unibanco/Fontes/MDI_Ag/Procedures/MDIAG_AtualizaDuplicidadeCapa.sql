SET QUOTED_IDENTIFIER	OFF
SET ANSI_NULLS		ON 
GO

If Exists (SELECT * FROM SYSOBJECTS WHERE id = object_id(N'MDIAG_AtualizaDuplicidadeCapa') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	DROP PROCEDURE MDIAG_AtualizaDuplicidadeCapa
GO

CREATE PROCEDURE MDIAG_AtualizaDuplicidadeCapa
	@DataProcessamento              int,
	@IdCapa                         int,
	@Status 			char(1),
	@Duplicidade			numeric(1)
as

	Begin Transaction
	--------------------------------------------------------------------------------------------------
	-- Atualiza Dados de Capa
	UPDATE Capa SET
		Status		= @Status,
		Duplicidade	= @Duplicidade,
		Ocorrencia	= CASE WHEN @Duplicidade = 0 THEN Null Else 998 END
	 WHERE  DataProcessamento	= @DataProcessamento
	   AND	IdCapa			= @IdCapa
	--------------------------------------------------------------------------------------------------
	-- Verifica Codigo de Erro
	If @@Error = 0 Begin
		Commit Transaction
		Return(0)
	End
	Else Begin
		Rollback Transaction
		Return(1)
	End
