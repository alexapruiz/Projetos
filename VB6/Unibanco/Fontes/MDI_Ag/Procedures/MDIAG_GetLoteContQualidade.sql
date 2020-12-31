SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_GetLoteContQualidade]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_GetLoteContQualidade]
GO


/****** Object:  Stored Procedure dbo.MDIAG_GetLoteContQualidade    Script Date: 17/11/00 11:22:09 ******/
CREATE PROCEDURE MDIAG_GetLoteContQualidade
	@DataProc	Int,
	@Intervalo	Int
As

	(
	SELECT IdLote, Status
	  FROM Lote 
	 WHERE DataProcessamento = @DataProc
	   And IdLote            > 0
	   And Status            = '0'
	)
	UNION ALL
	(
	SELECT IdLote, Status
	  FROM Lote 
	 WHERE DataProcessamento = @DataProc
	   And IdLote            > 0
	   And Status            = '1'
	   And DateDiff(Second, HoraAtual, GetDate()) > @Intervalo
	)
	Order By IdLote






GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

