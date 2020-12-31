SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_CapturaCapa]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_CapturaCapa]
GO


/****** Object:  Stored Procedure dbo.MDIAG_CapturaCapa    Script Date: 17/11/00 11:22:10 ******/
CREATE Proc MDIAG_CapturaCapa
	@DataProcessamento		Int,
	@IdLote 			Int,
	@idEnv_Mal			Char(1),
	@Capa				Numeric(18),
	@AgOrig				Smallint,
	@IdCapa 			Int Output

as

Declare @Status    char(1)

	-- Verifica se capa virtual (9)
        if @Capa = 9 Begin
		------------------------------------------------------------------------------------------
		Insert Capa (DataProcessamento,	IdLote, idEnv_Mal, Capa, Num_Malote, AgOrig, 
				Status, DataCriacao, Duplicidade )
	              Values(  @DataProcessamento, @IdLote, @idEnv_Mal, @Capa, 0, @AgOrig, 
				'1', GetDate(), 0 )
		if @@Error = 0 Begin
			Select @IdCapa = @@Identity
		End
		if @@Error = 0 Begin
			Return(0)
		End
		Else Begin
			Return(1)
		End
        End

	-- Verifica se nao existe a Capa 
	if not exists ( Select 1
			  From Capa 
			 Where DataProcessamento 	= @DataProcessamento
			   And Capa 			= @Capa
			   And Status not in 		  ('F','D','P')) Begin
		------------------------------------------------------------------------------------------
		Insert Capa ( 	DataProcessamento, IdLote, idEnv_Mal, Capa, Num_Malote, AgOrig, 
				Status, DataCriacao, Duplicidade )
	              Values(  @DataProcessamento, @IdLote, @idEnv_Mal, @Capa, 0, @AgOrig, 
				'1', GetDate(), 0 )
		if @@Error = 0 Begin
			Select @IdCapa = @@Identity
		End
		if @@Error = 0 Begin
			Return(0)
		End
		Else Begin
			Return(1)
		End
	End
	Else Begin

		-- Ja existe capa com mesmo DataProc + Capa
		------------------------------------------------------------------------------------------
		Select	@IdCapa			= IdCapa,
			@Status			= Status
		  From	Capa 
		 Where	DataProcessamento 	= @DataProcessamento
		   And	Capa 			= @Capa
		   And	Status not in 		('F','D','P')

		-- Se o Status for Capa Cadastrada entao atualiza Capa
		if @Status = '0' Begin
			Update Capa Set 
				IdLote 			= @IdLote,
				Status 			= '1' 
		   	 Where DataProcessamento	= @DataProcessamento
			   And IdCapa 			= @IdCapa
		End
		Else Begin
			Insert Capa ( DataProcessamento, IdLote, idEnv_Mal, Capa, Num_Malote, AgOrig, 
					Status, DataCriacao, Duplicidade )
	                Values      ( @DataProcessamento, @idLote, @idEnv_Mal, @Capa, 0, @AgOrig, 
				'1', GetDate(), 1 )

			if @@Error = 0 Begin
				Select @IdCapa = @@Identity
			End
		End
		if @@Error = 0 Begin
			Return(0)
		End
		Else Begin
			Return(1)
		End
	End


GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

