SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_RemoveDocumento]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_RemoveDocumento]
GO


/****** Object:  Stored Procedure dbo.MDIAG_RemoveDocumento    Script Date: 17/11/00 11:22:10 ******/
CREATE PROCEDURE MDIAG_RemoveDocumento
	@DataProcessamento	Int,
	@IdDocto		Int
As

Declare	@TipoDocto		SmallInt,
	@Erro			Int,
	@LinhasAfetadas		Int

---------------------------------------------------
--Seleciona o Tipo do Documento que serah removido
---------------------------------------------------
	SELECT @TipoDocto = TipoDocto
	  FROM Documento
	 WHERE DataProcessamento = @DataProcessamento
	   And IdDocto           = @IdDocto

	Select @Erro           = @@Error
	      ,@LinhasAfetadas = @@Rowcount

	If @LinhasAfetadas <> 1  or @Erro <> 0       -- Linha nao encontrada ou erro 
	   Return(1)


	Begin Transaction

	----------------------------------------------------------------------------	
	-- Remove o Documento da sua tabela especifica, se houver (ver comentarios)
	----------------------------------------------------------------------------


	-- TipoDocto = 0 (DOCUMENTO INDEFINIDO) soh estah na tabela de Documento
	-- TipoDocto = 1 (CAPA MALOTE EMPRESA) soh apaga da tabela de Documento


	If @TipoDocto in (2,3)        -- DEPOSITO CONTA CORRENTE E DEPOSITO CONTA POUPANCA        
           	Delete Deposito 
	     	 Where DataProcessamento = @DataProcessamento
	     	   And IdDocto           = @IdDocto
  
	Else If @TipoDocto = 4        -- ADCC                           
 	     Delete ADCC 
	      Where DataProcessamento = @DataProcessamento
	        And IdDocto           = @IdDocto

	Else If @TipoDocto in (5,6,7) -- CHEQUE UBB SACADO (PAGTO), CHEQUE TERCEIRO (PAGTO), 
                                    -- CHEQUE DEPOSITO      
  	     Delete Cheque 
	      Where DataProcessamento = @DataProcessamento
	        And IdDocto           = @IdDocto

	Else If @TipoDocto IN (8, 9, 24, 25, 26) -- CONCESSIONARIA VALOR REAL,
                                               -- CONCESSIONARIA VALOR INDEXADO
                                               -- TRIBUTOS MUNICIPAIS, TRIBUTOS ESTADUAIS,
                                               -- TRIBUTOS FEDERAIS
	     Delete CBIndex
	      Where DataProcessamento = @DataProcessamento
	        And IdDocto           = @IdDocto

	Else If @TipoDocto IN (10,28,29,30,31)    -- FICHA COMPENSACAO, UNICOBRANCA UBB,
                                                -- COBRANCA IMEDIATA UBB, COBRANCA ESPECIAL UBB
                                                -- COBRANCA TERCEIROS
	     Delete FichaCompensacao 
	      Where DataProcessamento = @DataProcessamento
	        And IdDocto           = @IdDocto

	Else If @TipoDocto IN (11,19) -- INSS e GRPS - nao existem mais
	     Goto TrataErro

	Else If @TipoDocto = 12       -- TITULOS (TERCEIROS SEM CB) 
	     Delete Titulo 
	      Where DataProcessamento = @DataProcessamento
	        And IdDocto           = @IdDocto

	Else If @TipoDocto = 13       -- COBRANCA REGISTRADA (SEM CB)   
   	     Delete CobrancaRegistrada 
     	      Where DataProcessamento = @DataProcessamento
	        And IdDocto           = @IdDocto

	Else If @TipoDocto = 14       -- COBRANCA ESPECIAL (SEM CB)     
	     Delete CobrancaEspecial
	      Where DataProcessamento = @DataProcessamento
	        And IdDocto           = @IdDocto

	Else If @TipoDocto = 15       -- DARM                           
	     Delete Darm 
	      Where DataProcessamento = @DataProcessamento
	        And IdDocto           = @IdDocto

	Else If @TipoDocto = 16       -- DARF PRETO                     
	     Delete DarfPreto 
	      Where DataProcessamento = @DataProcessamento
	        And IdDocto           = @IdDocto

	Else If @TipoDocto = 17       -- DARF SIMPLES                   
	     Delete DarfSimples 
	      Where DataProcessamento = @DataProcessamento
	        And IdDocto           = @IdDocto

	Else If @TipoDocto = 18       -- GARE                           
	     Delete Gare

	      Where DataProcessamento = @DataProcessamento
	        And IdDocto           = @IdDocto


	-- 19 - GRPS nao existe mais, ver tratamento junto com 11 - INSS

	Else If @TipoDocto IN (20, 21, 22, 23) -- AGUA, GAS, LUZ, TELEFONE
	     Delete ArrecadacaoEletronica
	      Where DataProcessamento = @DataProcessamento
	        And IdDocto           = @IdDocto


	Else If @TipoDocto = 27       -- ARRECADACAO CONVENCIONAL
	     Delete ArrecadacaoConvencional
	      Where DataProcessamento = @DataProcessamento
	        And IdDocto           = @IdDocto

	-- 28 - UNICOBRANCA UBB, 29 - COBRANCA IMEDIATA UBB, 30 - COBRANCA ESPECIAL UBB
      -- 31 - COBRANCA TERCEIROS estao tratados acima na Ficha de Compensacao

	Else If @TipoDocto IN (32,34) -- AJUSTE CREDITO DEPOSITO, CREDITO AUTOMATICO                           
	     Delete AjusteCredito
	      Where DataProcessamento = @DataProcessamento
	        And IdDocto           = @IdDocto

	Else If @TipoDocto IN (33,38) -- AJUSTE DEBITO DEPOSITO, DEBITO AUTOMATICO                           
	     Delete AjusteDebito
	      Where DataProcessamento = @DataProcessamento
	        And IdDocto           = @IdDocto

	Else If @TipoDocto = 35      -- GPS                            
	     Delete GPS
	      Where DataProcessamento = @DataProcessamento
	        And IdDocto           = @IdDocto

	Else If @TipoDocto = 36      -- CARTAOAVULSO                   
	     Delete CartaoAvulso
	      Where DataProcessamento = @DataProcessamento
	        And IdDocto           = @IdDocto

	Else If @TipoDocto = 37      -- OCT                            
 	     Delete OCT
	      Where DataProcessamento = @DataProcessamento
	        And IdDocto           = @IdDocto
  
      -- 38 - DEBITO AUTOMATICO estah tratado acima no AjusteDebito                      
	-- 39 - CAPA OCT, soh apaga na Tabela Documento
	Else If @TipoDocto = 40      -- FGTS
 	     Delete FGTS
	      Where DataProcessamento = @DataProcessamento
	        And IdDocto           = @IdDocto

   Select @Erro           = @@Error
         ,@LinhasAfetadas = @@Rowcount

   -----------------------------------------------------------------------------------------
   -- Se não removeu da Tabela específica e o documento não for 0,1 ou 39, que soh estao na 
   -- tabela Documento, o Tipo de Documento não existe ou ocorreu um erro -> TrataErro
   -----------------------------------------------------------------------------------------
   If (@LinhasAfetadas <> 1 and @TipoDocto not in (0,1,39)) or @Erro <> 0 
	Goto TrataErro

   -------------------------------
   -- Remove da Tabela Documento
   -------------------------------
   Delete Documento 
    Where DataProcessamento = @DataProcessamento
      And IdDocto           = @IdDocto

   Select @Erro = @@Error
         ,@LinhasAfetadas = @@Rowcount

   ---------------------------------------
   -- Se tudo OK, Commit, senao RollBack
   ---------------------------------------
   If @LinhasAfetadas = 1 and @Erro = 0 Begin
		Commit Transaction
   		Return(0)
   End
   

TrataErro:
   Begin
	RollBack Transaction
   	Return(1)
   End





GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

