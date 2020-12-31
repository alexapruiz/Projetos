SET QUOTED_IDENTIFIER	OFF
SET ANSI_NULLS		ON 
GO

If Exists (SELECT * FROM SYSOBJECTS WHERE id = object_id(N'MDIAG_AtualizaCapa') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	DROP PROCEDURE MDIAG_AtualizaCapa
GO

/*
     	Esta procedure atualiza Agencia, Capa de Envelope/Malote e N£mero de Malote para tabela capa 
	j  cadastrada, vincula a Capa de Envelope e C¢digo	 do CMC7 na tabela Documento, apenas se 
	fornecido o n£mero do documento (IdDocto).
    	Se informado o N£mero do Malote Empresa ser  atualizado o CMC7 no campo Leitura senƒo
	ser  atualizado o n£mero da Capa de Envelope na Tabela Documento.


     Retorno:  (0)-Sucesso
               (1)-Erro
*/

CREATE Procedure MDIAG_AtualizaCapa
	@Data		Int,
	@IdCapa		Int,
	@Capa 		Numeric(18),
	@AgOrig		SmallInt,
	@IdDocto	Int 		= 0,
	@Num_Malote	Numeric(11) 	= 0,
	@CMC7		Char(30) 	= NULL
As

	Begin Transaction

	UPDATE 	Capa SET
		AgOrig 			= @AgOrig,
		Capa			= @Capa,
		Num_Malote		= @Num_Malote,
		Ocorrencia		= null,
		Duplicidade		= 0,
		idEnv_Mal		= CASE @Num_Malote WHEN 0 THEN "E" ELSE "M" END
	 WHERE 	DataProcessamento 	= @Data
	   AND	IdCapa			= @IdCapa

	If @IdDocto > 0 Begin
		UPDATE	Documento SET
		 	Leitura 		= CASE @Num_Malote WHEN 0 THEN (Right("00000000" + Convert(Varchar(8), @Capa),8)) ELSE @CMC7 END ,
			TipoDocto		= 1	---* Envelope/Malote *---
		 WHERE	DataProcessamento 	= @Data
		   AND	IdCapa			= @IdCapa		
		   AND	IdDocto			= @IdDocto

	End

	If @@Error = 0 Begin
		Commit Transaction
		Return(0)
	End
	Else Begin
		Rollback Transaction
		Return(1)
	End










