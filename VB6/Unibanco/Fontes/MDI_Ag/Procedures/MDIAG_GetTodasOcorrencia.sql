SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_GetTodasOcorrencia]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_GetTodasOcorrencia]
GO


/****** Object:  Stored Procedure dbo.MDIAG_GetTodasOcorrencia    Script Date: 17/11/00 11:22:09 ******/
CREATE PROCEDURE MDIAG_GetTodasOcorrencia

As

	Select	Ocorrencia,
		Descricao
	  From 	Ocorrencia 
	 Order 	by  Ocorrencia asc



GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

