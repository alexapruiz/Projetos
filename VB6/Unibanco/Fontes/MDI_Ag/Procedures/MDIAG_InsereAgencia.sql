SET QUOTED_IDENTIFIER	OFF
SET ANSI_NULLS		ON 
GO

If Exists (SELECT * FROM SYSOBJECTS WHERE id = object_id(N'MDIAG_InsereAgencia') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	DROP PROCEDURE MDIAG_InsereAgencia
GO
CREATE PROCEDURE MDIAG_InsereAgencia
	@Agencia		SmallInt,
	@Lacre			Decimal(8),
	@QtdInformada		Int,
	@QtdGravada		Int,
	@Identificador		Char(10),
	@HoraChegada		Char(5),
	@HoraCadastrada		Char(5),
	@IdEnvMal		Char(1)

As

	SET NOCOUNT ON
	Begin Transaction
	
	/****************************
	 * Insere na tabela Agencia *
	 ****************************/
	INSERT INTO Agencia 
		(Agencia, 
		 Lacre, 
		 QtdInformada,
		 QtdGravada,
		 Identificador,
		 HoraChegada,
		 HoraCadastrada,
		 IdEnv_Mal)
	VALUES
		(@Agencia,
		 @Lacre,
		 @QtdInformada,
		 @QtdGravada,
		 @Identificador,
		 @HoraChegada,
		 @HoraCadastrada,
		 @IdEnvMal)

	If @@Error = 0 Begin
		Commit Transaction
		Return(0)
	End
	Else Begin
		Rollback Transaction
		Return(@@Error)
	End


