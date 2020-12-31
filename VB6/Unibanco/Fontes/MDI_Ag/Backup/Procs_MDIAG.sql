if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_ChecarCapaEnvelope]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_ChecarCapaEnvelope]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_InsereAgencia]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_InsereAgencia]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_InsereCapa]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_InsereCapa]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_InsereControleImportacao]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_InsereControleImportacao]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_InsereDocumento]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_InsereDocumento]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_InsereLog]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_InsereLog]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_InsereLote]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_InsereLote]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_ObtemAgenf]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_ObtemAgenf]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[MDIAG_VerificaImportacao]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[MDIAG_VerificaImportacao]
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO



---------------------------------------------------------------------------
---     Verifica se existe Nœmero de Capa jÿ cadastrada (Duplicidade) ---
---         ---
--- Pesquisa por Ag^ncia ou apenas por Nœmero de Capa passando ---
--- o parŸmetro AgOrig      ---
---         ---
--- Retorno: (0) - Sucesso      ---
---   (1) - Erro      ---
---------------------------------------------------------------------------
CREATE PROCEDURE MDIAG_ChecarCapaEnvelope
 @DataProcessamento Int,
 @AgOrig   SmallInt = Null,
 @Capa   Numeric(18),
 @Registros  Int Output,
 @IdCapa   Int
As 
 SELECT @Registros = Count(*)
   FROM Capa
  WHERE DataProcessamento = @DataProcessamento 
    AND Capa   = @Capa
    AND IdCapa   <> @IdCapa
    AND Convert(Char(10),AgOrig) like Case @AgOrig WHEN Null THEN "%" ELSE Convert(Char(10),@AgOrig) END
    AND Status   not in ("0","1","D","F","P") --- Verifica duplicidade somente para capa complementada ou adiante ---
 If @@Error <> 0
  Return(1)
 Return(0)




GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO



Create Procedure MDIAG_InsereAgencia
 @DataProc  int,
 @Agencia  smallint,
 @Lacre   decimal(8),
 @QtdInformada  int,
 @HoraCadastrada  char(5),
 @IdEnvMal         char(1)
As
 insert into Agencia 
   (DataProcessamento, 
                             Agencia, 
                             Lacre, 
                             QtdInformada,
                             QtdGravada,
                             Identificador,
                             HoraChegada,
                             HoraCadastrada,
                             IdEnv_Mal)
 values (@DataProc,
                @Agencia,
                @Lacre,
                @QtdInformada,
                @QtdInformada,
                '',
                substring(convert(varchar(5), GetDate(), 108), 1, 5),
                @HoraCadastrada,
                @IdEnvMal)
 if @@Error = 0 begin
  return(0)
 end 
 else begin
  return(1)
 end



GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO



CREATE PROCEDURE MDIAG_InsereCapa
 @DataProc  int,
 @IdLote  int,
 @IdEnv_Mal  char(1),
 @Capa   numeric(18),
 @Num_Malote  numeric(9),
 @AgOrig  smallint,
 @Status  char(1),
 @Ocorrencia  decimal(5),
 @Duplicidade  decimal(5),
 @IdCapa  int  OutPut
 AS
  Insert Capa (  DataProcessamento, IdLote, idEnv_Mal, Capa, Num_Malote, AgOrig, 
    Status, DataCriacao, Ocorrencia, Duplicidade )
               Values(  @DataProc, @IdLote, @idEnv_Mal, @Capa, @Num_Malote, @AgOrig, 
    @Status, GetDate(), @Ocorrencia, @Duplicidade )
  if @@Error = 0 Begin
   Select @IdCapa = @@Identity
  End
  if @@Error = 0 Begin
   Return(0)
  End
  Else Begin
   Return(1)
  End



GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO



CREATE PROCEDURE MDIAG_InsereControleImportacao 
 @DataProc  int,
 @DataGravacao varchar(17),
 @Agencia  smallint,
 @Remessa  int,
 @Cd   tinyint,
 @QtdLotes  int,
 @QtdCapas  int,
 @QtdDocumentos int
AS
 insert into ControleImportacao(DataProcessamento,
  Agencia,
  Remessa,
  CD,
  DataGravacao,
  DataImportacao,
  QtdLotes,
  QtdCapas,
  QtdDocumentos)
 values (@DataProc,
  @Agencia,
  @Remessa,
  @CD,
  convert(smalldatetime, @DataGravacao),
  GetDate(),
  @QtdLotes,
  @QtdCapas,
  @QtdDocumentos)
 if @@Error = 0 begin
  return(0)
 end 
 else begin
  return(1)
 end



GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO



CREATE PROCEDURE MDIAG_InsereDocumento 
 @DataProc  int,
 @IdCapa  int,
 @OrdemCaptura  smallint,
 @TipoDocto  smallint,
 @Leitura  varchar(48),
 @Frente  varchar(20),
 @Verso   varchar(20),
 @Status  char(1),
 @Ordem  char(1),
 @IdDocto  int OutPut
AS
 Declare @Error                  int
 Insert into Documento ( DataProcessamento,  IdCapa,  TipoDocto,  Leitura,
    Frente,   Verso,  Status,  Alcada, 
    Autenticado,  OcorrenciaOk, Vinculo,  Duplicidade, 
    CodCenape,  CodBarComplem, Ordem,  Valor,
    OrdemCaptura)
                     values (  @DataProc,   @IdCapa,  @TipoDocto,  @Leitura, 
    @Frente,   @Verso,  @Status, 'N', 
    '0',    '',   0,   0, 
    0,    'N',   @Ordem,  0,
    @OrdemCaptura)
 Select @Error = @@Error, @IdDocto = @@Identity
 If @Error = 0 Begin
  Return(0)
 End
 Else Begin
  Return(1)
 End



GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO



Create Procedure MDIAG_InsereLog
 @DataProc   int,
 @IdCapa     int,
 @IdDocto    int,
 @Data   varchar(17),
 @Usuario    Char(10),
 @Acao       tinyint
As
 Insert Into Log
  (DataProcessamento,
   IdCapa,
   IdDocto,
   Data,
   Login,
   Acao)
 Values
  (@DataProc,
   @IdCapa,
   @IdDocto,
   convert(DateTime, @Data),
   @Usuario,
   @Acao)
if @@Error = 0 begin
 return(0)
end
else begin
 return(1)
end



GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO



CREATE Proc MDIAG_InsereLote
 @DataProcessamento  int,
 @Prioridade   smallint,
 @Agencia   int,
 @IdLote    int
as
 --------------------------------------------------------------------------------------------------
 -- Busca ultimo Lote da Agenica
/* Select @IdLote = IsNull((Max(IdLote)+1), @Agencia * 100000 + 1)
 From Lote 
 Where IdLote between (@Agencia * 100000) And ((@Agencia + 1) * 100000) and
  DataProcessamento = @DataProcessamento
*/
 --------------------------------------------------------------------------------------------------
 -- Grava Lote
 Insert Lote (  DataProcessamento, IdLote, Status, Prioridade )
      Values ( @DataProcessamento, @IdLote, '2', @Prioridade )
 if @@Error = 0
 Begin
  Return(0)
 End
 Else Begin
  select @IdLote = -1
  Return(1)
 End
 --------------------------------------------------------------------------------------------------



GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO



CREATE PROCEDURE MDIAG_ObtemAgenf
 @AgeFsCdAgen SmallInt
As
 SELECT DISTINCT AgeFsNoAgen
   FROM AGENF
  WHERE AgeFsCdAgen = @AgeFsCdAgen



GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO



CREATE PROCEDURE MDIAG_VerificaImportacao 
 @DataProc   int,
 @Agencia  smallint,
 @Remessa  int,
 @CD   tinyint
AS
 if exists( Select 1 
  From ControleImportacao
  Where DataProcessamento = @DataProc  And
   Agencia = @Agencia    And
   Remessa = @Remessa   And
   CD = @CD )  Begin
  return(1)
 End
 Else Begin
  return(0)
 End
 
 



GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

