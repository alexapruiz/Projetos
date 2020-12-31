SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_TotalizaAgencia]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_TotalizaAgencia]
GO

CREATE PROCEDURE MDIAG_TotalizaAgencia
	@pAgencia	SmallInt,
	@pLacre	Decimal(8)
AS


Declare @vTotal_Env		Int
Declare @vTotal_Mal		Int
Declare @vEnv_HoraCadastrada	Char(5)
Declare @vMal_HoraCadastrada	Char(5)
Declare @Erro			Int


	Select @Erro = 0

	/*--------------------------
	   Insere Envelope
	--------------------------*/
	SELECT	@vTotal_Env 		= Count(CP.IdEnv_Mal),
		@vEnv_HoraCadastrada 	= Convert(Char(5),Max(CP.DataCriacao),108)
	  FROM 	CAPA CP
	 WHERE 	CP.IdEnv_Mal = 'E'

	If Exists(Select 1
		    From Agencia
		   Where IdEnv_Mal = 'E'
		     AND QtdInformada = @vTotal_Env) Begin
		UPDATE Agencia	SET
			Agencia		=	@pAgencia,
			Lacre		=	@pLacre,
			QtdInformada	=	@vTotal_Env,
			HoraCadastrada	=	@vEnv_HoraCadastrada
		WHERE	QtdInformada	= 	@vTotal_Env
		  AND	IdEnv_Mal	=	'E'

		Select @Erro = @@Error
		If @Erro <> 0 Goto lbl_Erro
	End
	Else Begin		If @vTotal_Env > 0 Begin
			INSERT INTO Agencia
			(Agencia,	Lacre,		QtdInformada,	HoraCadastrada,		IdEnv_Mal)
			VALUES
			(@pAgencia,	@pLacre,	@vTotal_Env,	@vEnv_HoraCadastrada,	'E')

			Select @Erro = @@Error
		End

		If @Erro <> 0 Goto lbl_Erro
	End

	/*---------------------------
	   Insere Malote
	---------------------------*/

	SELECT @vTotal_Mal 		= Count(CP.IdEnv_Mal),
		@vMal_HoraCadastrada 	= Convert(Char(5),Max(CP.DataCriacao),108)
	  FROM CAPA CP
	 WHERE CP.IdEnv_Mal = 'M'

	If Exists(Select 1 From Agencia Where IdEnv_Mal = 'M' AND QtdInformada = @vTotal_Mal) Begin
		UPDATE	Agencia	SET
			Agencia		=	@pAgencia,
			Lacre		=	@pLacre,
			QtdInformada	=	@vTotal_Mal,
			HoraCadastrada	=	@vMal_HoraCadastrada
		 WHERE	QtdInformada	=	@vTotal_Mal
		   AND	IdEnv_Mal	=	'M'

		Select @Erro = @@Error
	End
	Else Begin
		if @vTotal_Mal > 0 Begin
			INSERT INTO Agencia
			(Agencia,	Lacre,		QtdInformada,	HoraCadastrada,		IdEnv_Mal)
			VALUES
			(@pAgencia,	@pLacre,	@vTotal_Mal,	@vMal_HoraCadastrada,	'M')

			Select @Erro = @@Error
		End
	End

lbl_Erro:


	Return @Erro



GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

