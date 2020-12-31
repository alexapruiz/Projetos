SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

if Exists (select * from sysobjects where id = object_id(N'MDIAG_ObtemAgenf') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	DROP PROCEDURE MDIAG_ObtemAgenf
GO

CREATE PROCEDURE MDIAG_ObtemAgenf
	@AgeFsCdAgen	SmallInt
As


	SELECT DISTINCT AgeFsNoAgen
	  FROM AGENF
	 WHERE AgeFsCdAgen = @AgeFsCdAgen

