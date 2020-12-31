/* ============================================================ */
/*   Database name:  MDI_AG                                     */
/*   DBMS name:      Microsoft SQL Server 6.x                   */
/*   Created on:     07/11/00  18:27                            */
/* ============================================================ */

/* ============================================================ */
/*   Database name:  MDI_AG                                     */
/* ============================================================ */
create database MDI_AG
go

/* ============================================================ */
/*   Table: Agencia                                             */
/* ============================================================ */
create table Agencia
(
    Agencia              smallint              not null,
    Lacre                decimal(8)            not null,
    QtdInformada         int                   not null,
    QtdGravada           int                   not null,
    Identificador        char(10)              not null,
    HoraChegada          char(5)               not null,
    HoraCadastrada       char(5)               not null,
    idEnv_Mal            char(1)               null    
)
go

/* ============================================================ */
/*   Table: Parametro                                           */
/* ============================================================ */
create table Parametro
(
    AgenciaCentral       numeric(4)            not null,
    AgenciaApresentante  numeric(4)            null    ,
    Tm_Pendente          int                   not null,
    Tm_Atualizacao       int                   not null,
    Dir_Dados            varchar(255)          not null,
    Dir_Imagens          varchar(255)          not null,
    Dir_Trabalho         varchar(255)          not null
)
go

/* ============================================================ */
/*   Table: LogErro                                             */
/* ============================================================ */
create table LogErro
(
    Data                 datetime              not null,
    Estacao              varchar(12)           not null,
    Login                varchar(10)           not null,
    Rotina               varchar(30)           null    ,
    Erro                 int                   null    ,
    Descricao            varchar(255)          null    
)
go

/* ============================================================ */
/*   Table: Ocorrencia                                          */
/* ============================================================ */
create table Ocorrencia
(
    Ocorrencia           dec(5)                not null,
    Descricao            varchar(82)           not null,
    constraint PK_Ocorrencia primary key nonclustered (Ocorrencia)
)
go

/* ============================================================ */
/*   Table: Usuario                                             */
/* ============================================================ */
create table Usuario
(
    idUsuario            int                   identity,
    Login                varchar(10)           not null,
    Nome                 varchar(50)           not null,
    Senha                varchar(10)           not null,
    CIF                  varchar(9)            null    ,
    constraint PK_USUARIO primary key (idUsuario)
)
go

/* ============================================================ */
/*   Table: TipoDocto                                           */
/* ============================================================ */
create table TipoDocto
(
    TipoDocto            smallint              not null,
    Nome                 varchar(30)           not null,
    constraint PK_TipoDocto primary key nonclustered (TipoDocto)
)
go

/* ============================================================ */
/*   Table: Grupo                                               */
/* ============================================================ */
create table Grupo
(
    idGrupo              char(3)               not null,
    Descricao            varchar(50)           not null,
    constraint PK_GRUPO primary key (idGrupo)
)
go

/* ============================================================ */
/*   Table: StatusCapa                                          */
/* ============================================================ */
create table StatusCapa
(
    Status               char(1)               not null,
    Descricao            varchar(50)           not null,
    constraint PK_STATUSCAPA primary key (Status)
)
go

/* ============================================================ */
/*   Table: StatusLote                                          */
/* ============================================================ */
create table StatusLote
(
    Status               char(1)               not null,
    Descricao            varchar(50)           not null,
    constraint PK_STATUSLOTE primary key (Status)
)
go

/* ============================================================ */
/*   Table: StatusDocumento                                     */
/* ============================================================ */
create table StatusDocumento
(
    Status               char(1)               not null,
    Descricao            varchar(50)           not null,
    constraint PK_STATUSDOCUMENTO primary key (Status)
)
go

/* ============================================================ */
/*   Table: Acao                                                */
/* ============================================================ */
create table Acao
(
    Acao                 tinyint               not null,
    Descricao            varchar(50)           not null,
    constraint PK_Acao primary key (Acao)
)
go

/* ============================================================ */
/*   Table: Lote                                                */
/* ============================================================ */
create table Lote
(
    DataProcessamento    int                   not null,
    IdLote               int                   not null,
    Status               char(1)               not null,
    Prioridade           smallint              not null,
    HoraAtual            datetime              null    ,
    constraint PK_Lote primary key nonclustered (DataProcessamento, IdLote),
    constraint CKT_LOTE check (
            (QtdEnvelope >= 0 and (QtdEnvelope <= 999999999)))
)
go

/* ============================================================ */
/*   Table: Capa                                                */
/* ============================================================ */
create table Capa
(
    DataProcessamento    int                   not null,
    IdCapa               int                   identity,
    IdLote               int                   not null,
    idEnv_Mal            char(1)               not null,
    Capa                 numeric(18)           not null,
    Num_Malote           numeric(11)           null    ,
    AgOrig               smallint              not null,
    Status               char(1)               not null,
    OrdemCaptura         int                   not null,
    DataCriacao          smalldatetime         not null,
    Ocorrencia           decimal(5)            null    ,
    Duplicidade          numeric(1)            null    ,
    HoraAtual            datetime              null    ,
    constraint PK_Capa primary key clustered (DataProcessamento, IdCapa)
)
go

/* ============================================================ */
/*   Table: Documento                                           */
/* ============================================================ */
create table Documento
(
    DataProcessamento    int                   not null,
    IdDocto              int                   identity,
    IdCapa               int                   not null,
    OrdemCaptura         smallint              not null,
    TipoDocto            smallint              not null,
    Leitura              varchar(48)           null    ,
    Frente               varchar(20)           not null,
    Verso                varchar(20)           not null,
    Status               char(1)               not null,
    Ordem                char(1)               null    ,
    constraint PK_Documento primary key nonclustered (DataProcessamento, IdDocto)
)
go

/* ============================================================ */
/*   Table: Log                                                 */
/* ============================================================ */
create table Log
(
    DataProcessamento    int                   not null,
    IdCapa               int                   null    ,
    IdDocto              int                   null    ,
    Data                 datetime              not null,
    Login                varchar(10)           not null,
    Acao                 tinyint               not null
)
go

/* ============================================================ */
/*   Table: GrupoUsuario                                        */
/* ============================================================ */
create table GrupoUsuario
(
    idUsuario            int                   not null,
    idGrupo              char(3)               not null,
    constraint PK_GRUPOUSUARIO primary key (idUsuario, idGrupo)
)
go

alter table Lote
    add constraint FK_LOTE_REF_8852_STATUSLO foreign key  (Status)
       references StatusLote (Status)
go

alter table Capa
    add constraint FK_CAPA_REF_4584_LOTE foreign key  (DataProcessamento, IdLote)
       references Lote (DataProcessamento, IdLote)
go

alter table Capa
    add constraint FK_CAPA_REF_8846_STATUSCA foreign key  (Status)
       references StatusCapa (Status)
go

alter table Capa
    add constraint FK_CAPA_REF_21052_OCORRENC foreign key  (Ocorrencia)
       references Ocorrencia (Ocorrencia)
go

alter table Documento
    add constraint FK_DOCUMENT_REF_4541_TIPODOCT foreign key  (TipoDocto)
       references TipoDocto (TipoDocto)
go

alter table Documento
    add constraint FK_DOCUMENT_REF_4572_CAPA foreign key  (DataProcessamento, IdCapa)
       references Capa (DataProcessamento, IdCapa)
go

alter table Documento
    add constraint FK_DOCUMENT_REF_8858_STATUSDO foreign key  (Status)
       references StatusDocumento (Status)
go

alter table Log
    add constraint FK_LOG_REF_15195_ACAO foreign key  (Acao)
       references Acao (Acao)
go

alter table GrupoUsuario
    add constraint FK_GRUPOUSU_REF_5955_USUARIO foreign key  (idUsuario)
       references Usuario (idUsuario)
go

alter table GrupoUsuario
    add constraint FK_GRUPOUSU_REF_5959_GRUPO foreign key  (idGrupo)
       references Grupo (idGrupo)
go

