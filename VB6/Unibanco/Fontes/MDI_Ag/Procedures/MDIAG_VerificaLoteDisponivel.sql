SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_VerificaLoteDisponivel]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_VerificaLoteDisponivel]
GO


/****** Object:  Stored Procedure dbo.MDIAG_VerificaLoteDisponivel    Script Date: 17/11/00 11:22:10 ******/
CREATE PROCEDURE MDIAG_VerificaLoteDisponivel
	@DataProc	Int,
	@IdLote		Int,
	@Intervalo	Int

As

Declare	@Count		Int,
	@Error		Int,
	@m_Status	Char(1),
	@m_Dif		Int

	SELECT 	@m_Status = Status,
		@m_Dif = DateDiff(second, HoraAtual, GetDate())
	  FROM 	Lote
	 WHERE	DataProcessamento 	= @DataProc
	   AND	IdLote			= @IdLote

	Select @Error = @@Error, @Count = @@RowCount

	if @Error <> 0 Begin
		Return(2)
	End
	Else If @Count = 0 Begin
		Return(1)
	End
	Else If (@m_Status = "0") Or (@m_Status = "1" And @m_Dif >= @Intervalo) Begin
		Return(0)
	End
	Else Begin
		Return(1)
	End


GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

