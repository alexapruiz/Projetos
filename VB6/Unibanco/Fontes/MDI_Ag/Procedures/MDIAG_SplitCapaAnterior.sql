IF Exists(Select * from sysobjects where id = object_id(N'MDIAG_SplitCapaAnterior') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	DROP PROCEDURE MDIAG_SplitCapaAnterior
GO

CREATE PROCEDURE MDIAG_SplitCapaAnterior
	@pDataProcessamento	Int,
	@pIdCapaAtual		Int,
	@pIdDoctoAtual		Int,
	@pTipoDocto		SmallInt
AS


	Declare @vErro			Int
	Declare @vTranCount		Int
	Declare @vIdCapaAnterior	Int
	Declare @vIdDoctoAnterior	Int
	Declare @vOrdemCapturaAnterior	Int
	Declare @vIdDocto		Int


	/*------------------------
	  Definição de variaveis
	------------------------*/
	Select 	@vErro 		= 0

	/*-------------------------------
	  Pega o IdDocto e Capa Anterior 
	-------------------------------*/
	SELECT	@vIdDoctoAnterior	= Max(IdDocto)
	  FROM	Documento
	 WHERE	DataProcessamento 	= @pDataProcessamento
	   AND	IdDocto			< @pIdDoctoAtual


	SELECT 	@vIdCapaAnterior  	= IdCapa,
		@vOrdemCapturaAnterior	= OrdemCaptura + 1
	  FROM 	Documento
	 WHERE	DataProcessamento	= @pDataProcessamento
	   AND	IdDocto			= @vIdDoctoAnterior

	/*--------------------------------
	  Abertura do cursor de documentos
	--------------------------------*/
	DECLARE cDocumentos INSENSITIVE CURSOR FOR
		SELECT 	IdDocto
		  FROM	Documento
		 WHERE 	DataProcessamento 	 = @pDataProcessamento
		   AND  IdCapa			 = @pIdCapaAtual
		   AND  IdDocto			>= @pIdDoctoAtual

	Open cDocumentos


	FETCH cDocumentos INTO @vIdDocto

	If @@FETCH_STATUS IN(-1,-2) Begin
		Select @vErro = 1
		Goto Fechar_Cursor
	End

	/*----------------------
	  Controle de Transação
	----------------------*/
	Select @vTranCount = @@TranCount
	If @vTranCount = 0
		Begin Transaction SplitCapaAnterior
	Else
		Save Transaction SplitCapaAnterior

	/*------------------------------------------------------
	  LOOP dos documentos para alterar o campo OrdemCaptura
	------------------------------------------------------*/
	Select @vIdDocto = @pIdDoctoAtual

	While @@FETCH_STATUS NOT IN (-1,-2) Begin

		UPDATE Documento SET
			IdCapa		= @vIdCapaAnterior,
			OrdemCaptura 	= @vOrdemCapturaAnterior
		 WHERE DataProcessamento 	= @pDataProcessamento
		   AND IdCapa			= @pIdCapaAtual
		   AND IdDocto		        = @vIdDocto

		Select @vErro = @@Error

		If @vErro <> 0 Goto Fechar_Cursor

		Select @vOrdemCapturaAnterior = @vOrdemCapturaAnterior + 1
		FETCH cDocumentos INTO @vIdDocto
	End

	/*--------------------------------------------------
	  Atualiza Tipo Docto para o documento selecionado
	--------------------------------------------------*/
	UPDATE Documento SET
		TipoDocto = @pTipoDocto
	 WHERE IdCapa 	= @vIdCapaAnterior
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
		RollBack Transaction SplitCapaAnterior
		Return (@vErro)
	End
	Else Begin
		If @vTranCount = 0
			Commit Transaction SplitCapaAnterior
		Return @vErro
	End

