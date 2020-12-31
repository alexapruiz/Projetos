SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

if exists (select * from sysobjects where id = object_id(N'MDIAG_AtualizaAgencia') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure MDIAG_AtualizaAgencia
GO


/****** Object:  Stored Procedure dbo.MDIAG_AtualizaAgencia    Script Date: 17/11/00 11:22:09 ******/
CREATE PROCEDURE MDIAG_AtualizaAgencia

	@pAgenciaApresentante		Numeric(5)

AS

	UPDATE Parametro
		Set AgenciaApresentante = @pAgenciaApresentante

	Return @@Error


GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

