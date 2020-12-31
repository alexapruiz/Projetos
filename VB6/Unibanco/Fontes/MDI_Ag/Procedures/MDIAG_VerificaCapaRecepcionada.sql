
if Exists (SELECT * FROM SYSOBJECTS WHERE id = object_id(N'MDIAG_VerificaCapaRecepcionada') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	DROP PROCEDURE MDIAG_VerificaCapaRecepcionada
GO

CREATE PROCEDURE MDIAG_VerificaCapaRecepcionada
	@DataProcessamento              Int,
	@Capa                           Numeric(18),
	@AgOrig				SmallInt,
	@Num_Malote			Numeric(11)

as

	--- Pesquisa Capa recepcionada ---
	if Exists(SELECT 1
		    FROM Capa 
		   WHERE DataProcessamento	= @DataProcessamento  
		     AND Capa 			= @Capa
		     AND Status			= "0"
		     AND Convert(Char(10),AgOrig) like CASE @Num_Malote WHEN 0 THEN Convert(Char(10),@AgOrig) ELSE "%" END)
		Return(1)

	Return(0)





