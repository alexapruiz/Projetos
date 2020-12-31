SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_CapturaDocumento]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_CapturaDocumento]
GO


/****** Object:  Stored Procedure dbo.MDIAG_CapturaDocumento    Script Date: 17/11/00 11:22:10 ******/
CREATE Procedure MDIAG_CapturaDocumento
	@DataProcessamento		Int,
	@IdCapa 			Int,
	@TipoDocto			Smallint,
	@Leitura			Varchar(48),
	@Frente				Varchar(20),
	@Verso				Varchar(20),
	@Ordem				Char(1),
	@OrdemCaptura			Int,
	@IdDocto			Int          Output

AS

	Declare @Error                  int

	/*--------------------------------------------------------------
	  Simplesmente insere na tabela documento com a ordem de captura
	--------------------------------------------------------------*/
	Insert into Documento (	DataProcessamento, 	IdCapa, 	TipoDocto, 	Leitura,
				Frente,			Verso,		Status,		Ordem,
				OrdemCaptura)

                     values (	@DataProcessamento, 	@IdCapa, 	@TipoDocto, 	@Leitura, 
				@Frente, 		@Verso, 	'0',		@Ordem,
			 	@OrdemCaptura)


	Select @Error = @@Error, @IdDocto = @@Identity

	If @Error = 0 Begin
		Return(0)
	End
	Else Begin
		Return(1)
	End





GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

