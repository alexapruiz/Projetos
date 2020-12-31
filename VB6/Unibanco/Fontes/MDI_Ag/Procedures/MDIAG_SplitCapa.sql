IF Exists(Select * from sysobjects where id = object_id(N'MDIAG_SplitCapa') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	DROP PROCEDURE MDIAG_SplitCapa
GO

CREATE PROCEDURE MDIAG_SplitCapa
	@pDataProcessamento	Int,
	@pIdCapaAtual		Int,
	@pIdDoctoAtual		Int,
	@pCapa			Numeric(9),
	@pNum_Malote		Numeric(9),
	@pIdEnv_Mal		Char(1),
	@pAgOrig		SmallInt
AS


	Declare @vErro			Int
	Declare @vTranCount		Int
	Declare @vIdLote		Int
	Declare @vIdCapa		Int
	Declare @vRet			Int
	Declare @vOrdemCaptura		Int
	Declare @vDataProcessamento	Int
	Declare @vIdDocto		Int


	DECLARE cDocumentos INSENSITIVE CURSOR FOR
		SELECT 	DataProcessamento, IdDocto
		  FROM	Documento
		 WHERE 	DataProcessamento 	 = @pDataProcessamento
		   AND  IdCapa			 = @pIdCapaAtual
		   AND  IdDocto			>= @pIdDoctoAtual

	/*------------------------
	  Definição de variaveis
	------------------------*/
	Select 	@vErro 		= 0,
		@vOrdemCaptura 	= 1

	Open cDocumentos


	FETCH cDocumentos INTO @vDataProcessamento, @vIdDocto

	If @@FETCH_STATUS IN(-1,-2) Begin
		Select @vErro = 1
		Goto Fechar_Cursor
	End


	Select @vTranCount = @@TranCount

	If @vTranCount = 0
		Begin Transaction SplitCapa
	Else
		Save Transaction SplitCapa

	/*--------------------------
	  Pega IdLote da capa atual
	--------------------------*/
	SELECT 	@vIdLote	= IdLote
	  FROM 	Capa
	 WHERE 	DataProcessamento = @pDataProcessamento
	   AND 	IdCapa		  = @pIdCapaAtual



	/*-----------------------------------------------------------------------------
	  Não utilizo a SP MDIAG_InsereCapa, porque não funciona de acordo com a regra
	-----------------------------------------------------------------------------*/

--	Exec @vErro = MDIAG_InsereCapa 
--			@pDataProcessamento,
--			@vIdLote,
--			@pIdEnv_Mal,
--			@pCapa,
--			@pNum_Malote,
--			@pAgOrig,
--			'1', --Status da capa
--			@vIdCapa OUTPUT


	Insert Capa (DataProcessamento, IdLote, idEnv_Mal, Capa, Num_Malote, AgOrig, 
		Status, DataCriacao, Duplicidade )

	Values(@pDataProcessamento, @vIdLote, @pIdEnv_Mal, @pCapa,@pNum_Malote, @pAgOrig, 
		'1', GetDate(), 0 )

	Select @vErro 	= @@Error
	Select @vIdCapa = @@Identity

	If (@vErro <> 0) Or (@vIdCapa Is Null) Goto Fechar_Cursor

	/*------------------------------------------------------
	  LOOP dos documentos para alterar o campo OrdemCaptura
	------------------------------------------------------*/
	Select @vIdDocto = @pIdDoctoAtual

	While @@FETCH_STATUS NOT IN (-1,-2) Begin

		UPDATE Documento SET
			OrdemCaptura 	= @vOrdemCaptura,
			IdCapa		= @vIdCapa
		 WHERE DataProcessamento 	= @pDataProcessamento
		   AND IdCapa			= @pIdCapaAtual
		   AND IdDocto		        = @vIdDocto


		Select @vErro = @@Error

		If @vErro <> 0 Goto Fechar_Cursor

		Select @vOrdemCaptura = @vOrdemCaptura + 1
		FETCH cDocumentos INTO @vDataProcessamento, @vIdDocto
	End


	/*---------------------------------------------
	  Atualiza Tipo Docto para Envelope ou Malote
	---------------------------------------------*/
	UPDATE Documento SET
		TipoDocto = 1
	 WHERE IdCapa 	= @vIdCapa
	   AND IdDocto 	= @pIdDoctoAtual

	Select @vErro = @@Error

Fechar_Cursor:
	/*-------------------------
	  Fecha e desaloca o cursor
	-------------------------*/
	Close cDocumentos
	Deallocate cDocumentos

Fim_Proc:
	If @vErro <> 0 Begin
		RollBack Transaction SplitCapa
		Return (@vErro)
	End
	Else Begin
		If @vTranCount = 0
			Commit Transaction SplitCapa
		Return @vErro
	End

