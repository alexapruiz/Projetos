SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_InsereLogErro]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_InsereLogErro]
GO


/****** Object:  Stored Procedure dbo.MDIAG_InsereLogErro    Script Date: 17/11/00 11:22:09 ******/
CREATE Procedure MDIAG_InsereLogErro
  @Erro            		int,
  @Descricao       	varchar(255)
As
Begin Transaction
  INSERT INTO LogErro(
                       Data,
                       Erro,
                       Descricao )
              values(
                       GetDate(),
                      @Erro,
                      @Descricao )
  If @@Error = 0 Begin
    Commit Transaction
  End
  Else Begin
    Rollback Transaction
  End




GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

