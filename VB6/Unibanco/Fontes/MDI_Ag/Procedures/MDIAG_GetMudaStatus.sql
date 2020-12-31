SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_GetMudaStatus]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_GetMudaStatus]
GO


/****** Object:  Stored Procedure dbo.MDIAG_GetMudaStatus    Script Date: 17/11/00 11:22:10 ******/
CREATE PROCEDURE MDIAG_GetMudaStatus
	@Status			VarChar (1),
	@NOcorrencia		Decimal(5),
	@DataProcessamento	Int,
	@IdCapa		Int

AS

Update Capa
	Set Status 	= @Status,
	Ocorrencia 	= @NOcorrencia
Where 	IdCapa 		= @IdCapa
And   	DataProcessamento = @DataProcessamento





GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

