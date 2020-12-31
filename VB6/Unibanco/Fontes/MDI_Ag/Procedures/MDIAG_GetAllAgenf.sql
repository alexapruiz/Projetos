SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_GetAllAgenf]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_GetAllAgenf]
GO


/****** Object:  Stored Procedure dbo.MDIAG_GetAllAgenf    Script Date: 17/11/00 11:22:09 ******/
CREATE PROCEDURE MDIAG_GetAllAgenf
AS
Select  '0035' As agefscdagen,
 'Santos' As agefsnoagen


GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

