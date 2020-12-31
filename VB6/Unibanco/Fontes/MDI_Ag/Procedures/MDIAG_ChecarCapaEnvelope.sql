
If Exists (SELECT * FROM SYSOBJECTS WHERE id = object_id(N'MDIAG_ChecarCapaEnvelope') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure MDIAG_ChecarCapaEnvelope
GO


---------------------------------------------------------------------------
---     Verifica se existe N£mero de Capa j  cadastrada (Duplicidade)	---
---									---
---	Pesquisa por Agˆncia ou apenas por N£mero de Capa passando	---
---	o parƒmetro AgOrig						---
---									---
---	Retorno: (0) - Sucesso						---
---		 (1) - Erro						---
---------------------------------------------------------------------------
CREATE PROCEDURE MDIAG_ChecarCapaEnvelope
	@DataProcessamento	Int,
	@AgOrig			SmallInt = Null,
	@Capa			Numeric(18),
	@Registros		Int Output,
	@IdCapa			Int
As	

	SELECT	@Registros = Count(*)
	  FROM	Capa (NOLOCK)
	 WHERE	DataProcessamento	= @DataProcessamento 
	   AND	Capa			= @Capa
	   AND	IdCapa			<> @IdCapa
	   AND	Convert(Char(10),AgOrig) like Case @AgOrig WHEN Null THEN "%" ELSE Convert(Char(10),@AgOrig) END
	   AND	Status			not in ("0","1","D","F","P")	--- Verifica duplicidade somente para capa complementada ou adiante ---


	If @@Error <> 0
		Return(1)

	Return(0)




