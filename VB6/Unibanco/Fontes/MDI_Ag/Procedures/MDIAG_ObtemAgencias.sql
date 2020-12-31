SET QUOTED_IDENTIFIER  ON    SET ANSI_NULLS  ON 
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_ObtemAgencias]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_ObtemAgencias]
GO

CREATE PROCEDURE MDIAG_ObtemAgencias

AS


	SELECT 	Agencia,
		Lacre,
		QtdInformada,
		HoraCadastrada,
		IdEnv_Mal
	  FROM	Agencia


GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

