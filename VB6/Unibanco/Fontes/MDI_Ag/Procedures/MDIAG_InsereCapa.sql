SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

if exists (select * from sysobjects where id = object_id(N'MDIAG_InsereCapa') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure MDIAG_InsereCapa
GO


/****** Object:  Stored Procedure dbo.MDIAG_InsereCapa    Script Date: 17/11/00 11:22:10 ******/
-----------------------------------------------------------------------------------------------------------
CREATE Proc MDIAG_InsereCapa
	@DataProcessamento 	Int,
	@IdLote   		Int,
	@idEnv_Mal  		Char(1),
	@Capa			Numeric(18),
	@Num_Malote		Numeric(11),
	@AgOrig			Smallint,
	@Status			Char(1),
	@IdCapa			Int OUTPUT
As

	Declare @TranCount	Int


	Select @TranCount = @@TranCount

	If @TranCount = 0
		Begin Transaction InsereCapa
	Else
		Save Transaction InsereCapa

--------------------------------------------------------------------------------------------------
	if @AgOrig <> 0 Begin
		-- Verifica se nao existe a Capa 
		if NOT EXISTS(Select 1
				 From Capa 
				Where DataProcessamento = @DataProcessamento
				  And AgOrig   		= @AgOrig
				  And Capa     		= @Capa
				  And Status not in 	('D','F')) Begin

			------------------------------------------------------------------------------------------
			Insert Capa (  DataProcessamento, IdLote, idEnv_Mal, Capa, Num_Malote, AgOrig, 
				       Status, DataCriacao, Duplicidade )
			                Values(  @DataProcessamento, @idLote, @idEnv_Mal, @Capa, @Num_Malote, @AgOrig, 
				              @Status, GetDate(), 0 )
			if @@Error = 0 Begin
				Select @IdCapa = @@Identity
			End
			if @@Error = 0 Begin
				Commit Transaction InsereCapa
				Return(0)
			End
			Else Begin
				Rollback Transaction InsereCapa
				Return(2)
			End
		End
		Else Begin
		-- Ja existe capa com mesmo DataProc + AgOrig + Capa
			Rollback Transaction InsereCapa
			Return(1)
		End
	End
	Else Begin
		-- AgOrig = 0
		-- Verifica se nao existe a Capa 
		if not exists (Select 1
				 From Capa 
				Where DataProcessamento 	= @DataProcessamento
				  And Capa   			= @Capa
				  And Status not in 		('D','F')) Begin
			------------------------------------------------------------------------------------------
			Insert Capa (  DataProcessamento, IdLote, idEnv_Mal, Capa, Num_Malote, AgOrig, 
				        Status, DataCriacao, Duplicidade )
			          Values(  @DataProcessamento, @idLote, @idEnv_Mal, @Capa, @Num_Malote, @AgOrig, 
				         @Status, GetDate(), 0 )
			if @@Error = 0 Begin
				Select @IdCapa = @@Identity
			End
			if @@Error = 0 Begin
				Commit Transaction InsereCapa
				Return(0)
			End
			Else Begin
				Rollback Transaction InsereCapa
				Return(2)
			End
		End
		Else Begin
			-- Ja existe capa com mesmo DataProc + Capa
			------------------------------------------------------------------------------------------
			Insert Capa (  DataProcessamento, IdLote, idEnv_Mal, Capa, Num_Malote, AgOrig, 
				        Status, DataCriacao, Duplicidade )
			Values(@DataProcessamento, @idLote, @idEnv_Mal, @Capa, @Num_Malote, @AgOrig, 
			               @Status, GetDate(), @@RowCount )
			if @@Error = 0 Begin
				Select @IdCapa = @@Identity
			End
			if @@Error = 0 Begin
				Commit Transaction InsereCapa
				Return(0)
			End
			Else Begin
				Rollback Transaction InsereCapa
				Return(2)
			End
		End
	End

GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

