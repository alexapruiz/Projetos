SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_LimpaMovimento]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_LimpaMovimento]
GO


/****** Object:  Stored Procedure dbo.MDIAG_LimpaMovimento    Script Date: 17/11/00 11:22:10 ******/
CREATE PROCEDURE MDIAG_LimpaMovimento
	@pDataProcessamento 	Int
AS

Declare @Erro Int
	 Begin Transaction LimpaMovimento
	 /*-----------------------------------
	    Limpa tabela Documento
	 -----------------------------------*/
	 Truncate Table UbbMDI..Documento
	 Select @Erro = @@Error
	 If @Erro <> 0 Goto lbl_Erro
	 /*-----------------------------------
	    Limpa tabela Capa
	 -----------------------------------*/
	 Delete From UbbMDI..Capa
	 Select @Erro = @@Error
	 If @Erro <> 0 Goto lbl_Erro
	 /*-----------------------------------
	    Limpa tabela Lote
	 -----------------------------------*/
	 Delete From UbbMDI..Lote
	 Select @Erro = @@Error
	 If @Erro <> 0 Goto lbl_Erro
	 /*-----------------------------------
	    Limpa tabela Agencia
	 -----------------------------------*/
	 Truncate Table UbbMDI..Agencia
	 Select @Erro = @@Error
	 If @Erro <> 0 Goto lbl_Erro
	 /*-----------------------------------
	    Limpa tabela Log
	 -----------------------------------*/
	 Truncate Table UbbMDI..Log
	 Select @Erro = @@Error
	 If @Erro <> 0 Goto lbl_Erro
	 /*-----------------------------------
	    Limpa tabela LogErro
	 -----------------------------------*/
	 Truncate Table UbbMDI..LogErro
	 Select @Erro = @@Error
	 If @Erro <> 0 Goto lbl_Erro
	 /*-------------------------------------------------------------------------------------------
	    Inserção do Lote 0
	 -------------------------------------------------------------------------------------------*/
	 Insert Into Lote
	 Values
	 (@pDataProcessamento, 0, '0','0',GETDATE())
	 Select @Erro = @@Error
	 If @Erro <> 0 Goto lbl_Erro
	 /*------------------------------------------------------------------------------------------
	    Atualização na tabela Parametro
	 -------------------------------------------------------------------------------------------*/
	 Update Parametro Set
	  DataProcessamento  = @pDataProcessamento,
	  Hm_Abertura  = GETDATE(),
	  Hm_Fechamento = NULL
	 Select @Erro = @@Error
	 If @Erro <> 0 Goto lbl_Erro

	 Commit Transaction LimpaMovimento

	 Update Statistics UbbMDI..Agencia
	 Update Statistics UbbMDI..Capa
	 Update Statistics UbbMDI..Documento
	 Update Statistics UbbMDI..Log
	 Update Statistics UbbMDI..LogErro
	 Update Statistics UbbMDI..Lote
lbl_Erro:
	 If @Erro <> 0 RollBack Transaction LimpaMovimento
 
	 Return @Erro







GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

