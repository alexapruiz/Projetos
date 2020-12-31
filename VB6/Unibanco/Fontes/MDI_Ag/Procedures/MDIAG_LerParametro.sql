SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_LerParametro]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_LerParametro]
GO


/****** Object:  Stored Procedure dbo.MDIAG_LerParametro    Script Date: 17/11/00 11:22:09 ******/
CREATE PROCEDURE MDIAG_LerParametro
AS
	Select	DataProcessamento,
		Hm_Abertura,
		Hm_Fechamento,
		AgenciaCentral,
		AgenciaApresentante,
		TM_Pendente,
		TM_Atualizacao,
		Dir_Dados,
		Dir_Imagens,
		Dir_Trabalho
	  From	Parametro



GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

