USE ADRRBR01

/*

SELECT * FROM SYSOBJECTS WHERE XTYPE = 'U' ORDER BY CRDATE 
SELECT 'drop table ' + name FROM SYSOBJECTS WHERE XTYPE = 'U' and name like 'GRF%' ORDER BY crdate DESC

select * into GRF_Trabalhos_Fases_Acao_Simultanea_BKP from GRF_Trabalhos_Fases_Acao_Simultanea
select * into GRF_Trabalhos_Fases_Status_BKP          from GRF_Trabalhos_Fases_Status
select * into GRF_Trabalhos_Fases_Manut12_BKP         from GRF_Trabalhos_Fases_Manut12
select * into GRF_Trabalhos_Fases_Manut11_BKP         from GRF_Trabalhos_Fases_Manut11
select * into GRF_Trabalhos_Fases_Manut10_BKP         from GRF_Trabalhos_Fases_Manut10
select * into GRF_Trabalhos_Fases_Manut9_BKP          from GRF_Trabalhos_Fases_Manut9
select * into GRF_Trabalhos_Fases_Manut8_BKP          from GRF_Trabalhos_Fases_Manut8
select * into GRF_Trabalhos_Fases_Manut7_BKP          from GRF_Trabalhos_Fases_Manut7
select * into GRF_Trabalhos_Fases_Manut6_BKP          from GRF_Trabalhos_Fases_Manut6
select * into GRF_Trabalhos_Fases_Manut5_BKP          from GRF_Trabalhos_Fases_Manut5
select * into GRF_Trabalhos_Fases_Manut4_BKP          from GRF_Trabalhos_Fases_Manut4
select * into GRF_Trabalhos_Fases_Manut3_BKP          from GRF_Trabalhos_Fases_Manut3
select * into GRF_Trabalhos_Fases_Manut2_BKP          from GRF_Trabalhos_Fases_Manut2
select * into GRF_Trabalhos_Fases_Manut1_BKP          from GRF_Trabalhos_Fases_Manut1
select * into GRF_Fases_BKP                           from GRF_Fases
select * into GRF_Trabalhos_BKP                       from GRF_Trabalhos
select * into GRF_Trabalhos_Fases_Status_Tipos_BKP    from GRF_Trabalhos_Fases_Status_Tipos
select * into GRF_Trabalhos_Tipos_BKP                 from GRF_Trabalhos_Tipos
select * into GRF_Produtos_BKP                        from GRF_Produtos
select * into GRF_Clientes_BKP                        from GRF_Clientes
select * into GRF_Representantes_BKP                  from GRF_Representantes
select * into GRF_Interrupcao_Tipos_BKP               from GRF_Interrupcao_Tipos
select * into GRF_Avaliacao_Tipos_BKP                 from GRF_Avaliacao_Tipos
select * into GRF_Aprovacao_Tipos_BKP                 from GRF_Aprovacao_Tipos
select * into GRF_Filiais_BKP                         from GRF_Filiais

drop table GRF_Trabalhos_Fases_Acao_Simultanea
drop table GRF_Trabalhos_Fases_Status
drop table GRF_Trabalhos_Fases_Manut_Generico
drop table GRF_Trabalhos_Fases_Manut12
drop table GRF_Trabalhos_Fases_Manut11
drop table GRF_Trabalhos_Fases_Manut10
drop table GRF_Trabalhos_Fases_Manut9
drop table GRF_Trabalhos_Fases_Manut8
drop table GRF_Trabalhos_Fases_Manut7
drop table GRF_Trabalhos_Fases_Manut6
drop table GRF_Trabalhos_Fases_Manut5
drop table GRF_Trabalhos_Fases_Manut4
drop table GRF_Trabalhos_Fases_Manut3
drop table GRF_Trabalhos_Fases_Manut2
drop table GRF_Trabalhos_Fases_Manut1
drop table GRF_Fases
drop table GRF_Trabalhos
drop table GRF_Trabalhos_Fases_Status_Tipos
drop table GRF_Trabalhos_Tipos
drop table GRF_Produtos
drop table GRF_Clientes
drop table GRF_Representantes
drop table GRF_Interrupcao_Tipos
drop table GRF_Avaliacao_Tipos
drop table GRF_Aprovacao_Tipos
drop table GRF_Filiais
*/      

CREATE TABLE    dbo.GRF_Filiais(
	pk_filial         int          identity primary key NOT NULL,
	codigo            int          NOT NULL,
	nome              varchar(30)  NOT NULL,
	descricao         varchar(max) NULL
)
create unique index IND_GRF_Filiais_1 ON GRF_Filiais(codigo)

CREATE TABLE    dbo.GRF_Aprovacao_Tipos(
	pk_aprovacao_tipo int          identity primary key NOT NULL,
	codigo            int          NOT NULL,
	nome              varchar(30)  NOT NULL,
	descricao         varchar(max) NULL
)
create unique index IND_GRF_Aprovacao_Tipos_1 ON GRF_Aprovacao_Tipos(codigo)

CREATE TABLE    dbo.GRF_Avaliacao_Tipos(
	pk_avaliacao_tipo int           identity primary key NOT NULL,
	codigo            int           NOT NULL,
	nome              varchar (30)  NOT NULL,
	descricao         varchar (max) NULL
)
create unique index IND_GRF_Avaliacao_Tipos_1 ON GRF_Avaliacao_Tipos(codigo)

CREATE TABLE    dbo.GRF_Interrupcao_Tipos(
	pk_interrupcao_tipo       int           identity primary key NOT NULL,
	codigo                    int           NOT NULL,
	nome                      varchar(30)   NOT NULL,
	descricao                 varchar(max)  NULL
) 
create unique index IND_GRF_Interrupcao_Tipos_1 ON GRF_Interrupcao_Tipos(codigo)

CREATE TABLE    dbo.GRF_Clientes(
	pk_cliente        int           identity primary key NOT NULL,
	Codigo            varchar(60)   NOT NULL,
	nome              varchar(150)  NOT NULL,
	nome_completo     varchar(250)  NULL
) 
create unique index IND_GRF_Clientes_1 ON GRF_Clientes(codigo)
create unique index IND_GRF_Clientes_2 ON GRF_Clientes(nome)

CREATE TABLE    dbo.GRF_Representantes(
	pk_representante   int           identity primary key NOT NULL,
	codigo             varchar(60)   NOT NULL,
	nome               varchar(150)  NOT NULL,
	nome_completo      varchar(250)  NULL
) 
create unique index IND_GRF_Representantes_1 ON GRF_Representantes(codigo)
create unique index IND_GRF_Representantes_2 ON GRF_Representantes(nome)

CREATE TABLE    dbo.GRF_Produtos(
	pk_produto         int           identity primary key NOT NULL,
	codigo             varchar(60)   NOT NULL,
	nome               varchar(500)  NOT NULL,
	descricao          varchar(max)  NULL, 
	caminho_foto       varchar(max)  NULL 
) 
create unique index IND_GRF_Produtos_1 ON GRF_Produtos(codigo)
create unique index IND_GRF_Produtos_2 ON GRF_Produtos(nome)

CREATE TABLE    dbo.GRF_Trabalhos_Fases_Status_Tipos(
	pk_trabalho_fase_status_tipo  int            identity primary key NOT NULL,
	codigo                        int            NOT NULL,
	nome                          varchar(50)    NOT NULL,
	tipo_acao                     varchar(20)    NOT NULL,  
	descricao                     varchar(max)   NULL,
	acao_operador                 tinyint        NOT NULL,
	interrompe_lead_time          tinyint        NOT NULL
)
create unique index IND_GRF_Trabalhos_Fases_Status_Tipos_1 ON GRF_Trabalhos_Fases_Status_Tipos(codigo)

CREATE TABLE    dbo.GRF_Trabalhos_Tipos(
	pk_trabalho_tipo   int           identity primary key NOT NULL,
	codigo             int           NOT NULL,
	nome               varchar(50)   NOT NULL,
	descricao          varchar(max)  NULL 
) 
create unique index IND_GRF_Trabalhos_Tipos_1 ON GRF_Trabalhos_Tipos(codigo)

CREATE TABLE    dbo.GRF_Trabalhos(
	pk_trabalho             int          identity primary key                              NOT NULL,
	fk_trabalho_alt_seq     int          references GRF_Trabalhos(pk_trabalho)             NULL,
	fk_trabalho_tipo        int          references GRF_Trabalhos_Tipos(pk_Trabalho_Tipo)  NOT NULL,
	fk_produto              int          references GRF_Produtos(pk_produto)               NOT NULL,
	fk_cliente              int          references GRF_Clientes(pk_cliente)               NOT NULL,
	fk_representante        int          references GRF_Representantes(pk_Representante)   NOT NULL,
	fk_aprovacao_tipo       int          references GRF_Aprovacao_Tipos(pk_aprovacao_Tipo) NOT NULL,  
	fk_filial               int          references GRF_Filiais(pk_filial)                 NOT NULL,        
	num_pedido              varchar(60)  NOT NULL,
	num_pedido_antigo       varchar(60)  NOT NULL, -- Zarasipa
	num_pedido_novo         varchar(60)  NOT NULL, -- Protheus
	lead_time_programado    varchar(6)   NOT NULL,
	dt_inclusao             datetime     NOT NULL,
	dt_alteracao            datetime     NULL,
	dt_exclusao             datetime     NULL
) 
create unique index IND_GRF_Trabalhos_1 ON GRF_Trabalhos(num_pedido,fk_trabalho_alt_seq)
create unique index IND_GRF_Trabalhos_2 ON GRF_Trabalhos(num_pedido_novo,fk_trabalho_alt_seq)

CREATE TABLE    dbo.GRF_Fases(
	pk_fase                 int           identity primary key NOT NULL,
	codigo                  int                                NOT NULL,
	nome                    varchar(100)                       NOT NULL,
	descricao               varchar(max)                       NULL,
	lead_time_perc_definido float                              NULL
) 
create unique index IND_GRF_Fases_1 ON GRF_Fases(codigo)
create unique index IND_GRF_Fases_2 ON GRF_Fases(nome)

CREATE TABLE    dbo.GRF_Trabalhos_Fases_Manut1(
	pk_trabalho_fase_manut1          int          identity primary key                                           NOT NULL,
	fk_trabalho_fase_manut1_alt_seq  int          references GRF_Trabalhos_Fases_Manut1(pk_trabalho_fase_manut1) NULL,
	fk_trabalho                      int          references GRF_Trabalhos(pk_trabalho)                          NOT NULL,
	fk_fase                          int          references GRF_Fases(pk_fase)                                  NOT NULL,
	fk_operador                      int          references APL_Usuarios(pk_usuario)                            NOT NULL,
	cores_alteradas                  int          NULL,
	circ_cilindro                    int          NULL,
	observacoes                      varchar(max) NULL,
	dt_inclusao                      datetime     NOT NULL,
	dt_alteracao                     datetime     NULL
) 
create unique index IND_GRF_Trabalhos_Fases_Manut1_1 ON GRF_Trabalhos_Fases_Manut1(fk_trabalho,fk_fase,fk_trabalho_fase_manut1_alt_seq)

CREATE TABLE    dbo.GRF_Trabalhos_Fases_Manut2(
	pk_trabalho_fase_manut2          int          identity primary key                                           NOT NULL,
	fk_trabalho_fase_manut2_alt_seq  int          references GRF_Trabalhos_Fases_Manut2(pk_trabalho_fase_manut2) NULL,
	fk_trabalho                      int          references GRF_Trabalhos(pk_trabalho)                          NOT NULL,
	fk_fase                          int          references GRF_Fases(pk_fase)                                  NOT NULL,
	fk_operador                      int          references APL_Usuarios(pk_usuario)                            NOT NULL,
	observacoes                      varchar(max) NULL,
	dt_inclusao                      datetime     NOT NULL,
	dt_alteracao                     datetime     NULL
) 
create unique index IND_GRF_Trabalhos_Fases_Manut2_1 ON GRF_Trabalhos_Fases_Manut2(fk_trabalho,fk_fase,fk_trabalho_fase_manut2_alt_seq)

CREATE TABLE    dbo.GRF_Trabalhos_Fases_Manut3(
	pk_trabalho_fase_manut3          int          identity primary key                                           NOT NULL,
	fk_trabalho_fase_manut3_alt_seq  int          references GRF_Trabalhos_Fases_Manut3(pk_trabalho_fase_manut3) NULL,
	fk_trabalho                      int          references GRF_Trabalhos(pk_trabalho)                          NOT NULL,
	fk_fase                          int          references GRF_Fases(pk_fase)                                  NOT NULL,
	fk_operador                      int          references APL_Usuarios(pk_usuario)                            NOT NULL,
	observacoes                      varchar(max) NULL,
	dt_inclusao                      datetime     NOT NULL,
	dt_alteracao                     datetime     NULL
) 
create unique index IND_GRF_Trabalhos_Fases_Manut3_1 ON GRF_Trabalhos_Fases_Manut3(fk_trabalho,fk_fase,fk_trabalho_fase_manut3_alt_seq)

CREATE TABLE    dbo.GRF_Trabalhos_Fases_Manut4(
	pk_trabalho_fase_manut4          int          identity primary key                                           NOT NULL,
	fk_trabalho_fase_manut4_alt_seq  int          references GRF_Trabalhos_Fases_Manut4(pk_trabalho_fase_manut4) NULL,
	fk_trabalho                      int          references GRF_Trabalhos(pk_trabalho)                          NOT NULL,
	fk_fase                          int          references GRF_Fases(pk_fase)                                  NOT NULL,
	fk_operador                      int          references APL_Usuarios(pk_usuario)                            NOT NULL,
	observacoes                      varchar(max) NULL,
	dt_inclusao                      datetime     NOT NULL,
	dt_alteracao                     datetime     NULL
) 
create unique index IND_GRF_Trabalhos_Fases_Manut4_1 ON GRF_Trabalhos_Fases_Manut4(fk_trabalho,fk_fase,fk_trabalho_fase_manut4_alt_seq)

CREATE TABLE    dbo.GRF_Trabalhos_Fases_Manut5(
	pk_trabalho_fase_manut5          int          identity primary key                                           NOT NULL,
	fk_trabalho_fase_manut5_alt_seq  int          references GRF_Trabalhos_Fases_Manut5(pk_trabalho_fase_manut5) NULL,
	fk_trabalho                      int          references GRF_Trabalhos(pk_trabalho)                          NOT NULL,
	fk_fase                          int          references GRF_Fases(pk_fase)                                  NOT NULL,
	fk_operador                      int          references APL_Usuarios(pk_usuario)                            NOT NULL,
	observacoes                      varchar(max) NULL,
	dt_inclusao                      datetime     NOT NULL,
	dt_alteracao                     datetime     NULL
) 
create unique index IND_GRF_Trabalhos_Fases_Manut5_1 ON GRF_Trabalhos_Fases_Manut5(fk_trabalho,fk_fase,fk_trabalho_fase_manut5_alt_seq)

CREATE TABLE    dbo.GRF_Trabalhos_Fases_Manut6(
	pk_trabalho_fase_manut6          int          identity primary key                                           NOT NULL,
	fk_trabalho_fase_manut6_alt_seq  int          references GRF_Trabalhos_Fases_Manut6(pk_trabalho_fase_manut6) NULL,
	fk_trabalho                      int          references GRF_Trabalhos(pk_trabalho)                          NOT NULL,
	fk_fase                          int          references GRF_Fases(pk_fase)                                  NOT NULL,
	fk_operador                      int          references APL_Usuarios(pk_usuario)                            NOT NULL,
	observacoes                      varchar(max) NULL,
	dt_inclusao                      datetime     NOT NULL,
	dt_alteracao                     datetime     NULL
) 
create unique index IND_GRF_GRF_Trabalhos_Fases_Manut6_1 ON GRF_Trabalhos_Fases_Manut6(fk_trabalho,fk_fase,fk_trabalho_fase_manut6_alt_seq)

CREATE TABLE dbo.GRF_Trabalhos_Fases_Manut7(
	pk_trabalho_fase_manut7          int          identity primary key                                           NOT NULL,
	fk_trabalho_fase_manut7_alt_seq  int          references GRF_Trabalhos_Fases_Manut7(pk_trabalho_fase_manut7) NULL,
	fk_trabalho                      int          references GRF_Trabalhos(pk_trabalho)                          NOT NULL,
	fk_fase                          int          references GRF_Fases(pk_fase)                                  NOT NULL,
	fk_operador                      int          references APL_Usuarios(pk_usuario)                            NOT NULL,
	observacoes                      varchar(max) NULL,
	dt_inclusao                      datetime     NOT NULL,
	dt_alteracao                     datetime     NULL
) 
create unique index IND_GRF_Trabalhos_Fases_Manut7_1 ON GRF_Trabalhos_Fases_Manut7(fk_trabalho,fk_fase,fk_trabalho_fase_manut7_alt_seq)

CREATE TABLE    dbo.GRF_Trabalhos_Fases_Manut8(
	pk_trabalho_fase_manut8          int          identity primary key                                           NOT NULL,
	fk_trabalho_fase_manut8_alt_seq  int          references GRF_Trabalhos_Fases_Manut8(pk_trabalho_fase_manut8) NULL,
	fk_trabalho                      int          references GRF_Trabalhos(pk_trabalho)                          NOT NULL,
	fk_fase                          int          references GRF_Fases(pk_fase)                                  NOT NULL,
	fk_operador                      int          references APL_Usuarios(pk_usuario)                            NOT NULL,
	observacoes                      varchar(max) NULL,
	dt_inclusao                      datetime     NOT NULL,
	dt_alteracao                     datetime      NULL
) 
create unique index IND_GRF_Trabalhos_Fases_Manut8_1 ON GRF_Trabalhos_Fases_Manut8(fk_trabalho,fk_fase,fk_trabalho_fase_manut8_alt_seq)

CREATE TABLE    dbo.GRF_Trabalhos_Fases_Manut9(
	pk_trabalho_fase_manut9          int          identity primary key                                           NOT NULL,
	fk_trabalho_fase_manut9_alt_seq  int          references GRF_Trabalhos_Fases_Manut9(pk_trabalho_fase_manut9) NULL,
	fk_trabalho                      int          references GRF_Trabalhos(pk_trabalho)                          NOT NULL,
	fk_fase                          int          references GRF_Fases(pk_fase)                                  NOT NULL,
	fk_operador                      int          references APL_Usuarios(pk_usuario)                            NOT NULL,
	fk_avaliacao_tipo                int          references GRF_Avaliacao_Tipos(pk_avaliacao_Tipo)              NULL,  
	dt_avaliacao                     datetime     NULL,
	dt_envio_cliente                 datetime     NULL,
	dt_envio_padroes_CQ              datetime     NULL,
	observacoes                      varchar(max) NULL, 
    dt_inclusao                      datetime     NOT NULL,
	dt_alteracao                     datetime     NULL
) 
create unique index IND_GRF_Trabalhos_Fases_Manut9_1 ON GRF_Trabalhos_Fases_Manut9(fk_trabalho,fk_fase,fk_trabalho_fase_manut9_alt_seq)

CREATE TABLE    dbo.GRF_Trabalhos_Fases_Manut10(
	pk_trabalho_fase_manut10          int          identity primary key                                             NOT NULL,
	fk_trabalho_fase_manut10_alt_seq  int          references GRF_Trabalhos_Fases_Manut10(pk_trabalho_fase_manut10) NULL,
	fk_trabalho                       int          references GRF_Trabalhos(pk_trabalho)                            NOT NULL,
	fk_fase                           int          references GRF_Fases(pk_fase)                                    NOT NULL,
	fk_operador                       int          references APL_Usuarios(pk_usuario)                              NOT NULL,
	total_cores                       int          NULL,
	observacoes                       varchar(max) NULL,
	dt_inclusao                       datetime     NOT NULL,
	dt_alteracao                      datetime     NULL
) 
create unique index IND_GRF_GRF_Trabalhos_Fases_Manut10_1 ON GRF_Trabalhos_Fases_Manut10(fk_trabalho,fk_fase,fk_trabalho_fase_manut10_alt_seq)

CREATE TABLE    dbo.GRF_Trabalhos_Fases_Manut11(
	pk_trabalho_fase_manut11          int          identity primary key                                             NOT NULL,
	fk_trabalho_fase_manut11_alt_seq  int          references GRF_Trabalhos_Fases_Manut11(pk_trabalho_fase_manut11) NULL,
	fk_trabalho                       int          references GRF_Trabalhos(pk_trabalho)                            NOT NULL,
	fk_fase                           int          references GRF_Fases(pk_fase)                                    NOT NULL,
	fk_operador                       int          references APL_Usuarios(pk_usuario)                              NOT NULL,
	observacoes                       varchar(max) NULL,
	dt_inclusao                       datetime     NOT NULL,
	dt_alteracao                      datetime     NULL
) 
create unique index IND_GRF_Trabalhos_Fases_Manut11_1 ON GRF_Trabalhos_Fases_Manut11(fk_trabalho,fk_fase,fk_trabalho_fase_manut11_alt_seq)

CREATE TABLE    dbo.GRF_Trabalhos_Fases_Manut12(
	pk_trabalho_fase_manut12          int          identity primary key                                             NOT NULL,
	fk_trabalho_fase_manut12_alt_seq  int          references GRF_Trabalhos_Fases_Manut12(pk_trabalho_fase_manut12) NULL,
	fk_trabalho                       int          references GRF_Trabalhos(pk_trabalho)                            NOT NULL,
	fk_fase                           int          references GRF_Fases(pk_fase)                                    NOT NULL,
	fk_operador                       int          references APL_Usuarios(pk_usuario)                              NOT NULL,
	fk_avaliacao_tipo                 int          references GRF_Avaliacao_Tipos(pk_avaliacao_Tipo)                NULL,  
	dt_avaliacao                      datetime     NULL,
	dt_receb_rolinho                  datetime     NULL,
	dt_envio_laminacao                datetime     NULL,
	dt_receb_laminacao                datetime     NULL,
	dt_envio_padroes_cliente          datetime     NULL,
	dt_envio_padroes_CQ               datetime     NULL,
	observacoes                       varchar(max) NULL,
	dt_inclusao                       datetime     NOT NULL,
	dt_alteracao                      datetime     NULL
) 
create unique index IND_GRF_Trabalhos_Fases_Manut12_1 ON GRF_Trabalhos_Fases_Manut12(fk_trabalho,fk_fase,fk_trabalho_fase_manut12_alt_seq)

CREATE TABLE    dbo.GRF_Trabalhos_Fases_Manut_Generico(
	pk_trabalho_fase_manut_gen         int          identity primary key                                                      NOT NULL,
	fk_trabalho_fase_manut_gen_alt_seq int          references GRF_Trabalhos_Fases_Manut_Generico(pk_trabalho_fase_manut_gen) NULL,
	fk_trabalho                        int          references GRF_Trabalhos(pk_trabalho)                                     NOT NULL,
	fk_fase                            int          references GRF_Fases(pk_fase)                                             NOT NULL,
	fk_operador                        int          references APL_Usuarios(pk_usuario)                                       NOT NULL,
	dt_liberacao_gravacao              datetime     NULL,
	dt_inclusao                        datetime     NOT NULL,
	dt_alteracao                       datetime     NULL
) 
create unique index IND_GRF_Trabalhos_Fases_Manut_Generico_1 ON GRF_Trabalhos_Fases_Manut_Generico(fk_trabalho,fk_fase,fk_trabalho_fase_manut_gen_alt_seq)

--*************
-- Históricos
--*************
CREATE TABLE    dbo.GRF_Trabalhos_Fases_Acao_Simultanea( 
	pk_trabalho_fase_acao_simu        int          identity primary key                                                      NOT NULL,
	fk_trabalho                       int          references GRF_Trabalhos(pk_trabalho)                                     NOT NULL,
	fk_fase                           int          references GRF_Fases(pk_Fase)                                             NOT NULL,
	fk_operador_acao_simu             int          references APL_Usuarios(pk_usuario)                                       NOT NULL,                  
	fk_operador                       int          references APL_Usuarios(pk_usuario)                                       NOT NULL,
	observacoes                       varchar(max) NULL,
	dt_inclusao                       datetime     NOT NULL,
	dt_liberacao_acao_simu            datetime     NULL
)
--create unique index IND_GRF_Trabalhos_Fases_Acao_Simultanea_1 ON GRF_Trabalhos_Fases_Acao_Simultanea(fk_trabalho,fk_fase,fk_operador_acao_simu)

-- Esta tabela posiciona a fase de forma manual ou automática. A verificação da fase atual do trabalho é feita através da < dt_inclusao > mais atual.
CREATE TABLE    dbo.GRF_Trabalhos_Fases_Status( 
	pk_trabalho_fase_status           int          identity primary key                                                      NOT NULL,
	fk_trabalho                       int          references GRF_Trabalhos(pk_trabalho)                                     NOT NULL,
	fk_fase                           int          references GRF_Fases(pk_Fase)                                             NOT NULL,
	fk_trabalho_fase_status_tipo      int          references GRF_Trabalhos_Fases_Status_Tipos(pk_Trabalho_Fase_Status_Tipo) NOT NULL,
	fk_operador                       int          references APL_Usuarios(pk_usuario)                                       NULL,
	obs_status                        varchar(max) NULL,
	dt_inclusao                       datetime     NOT NULL
)
create unique index IND_GRF_Trabalhos_Fases_Status_1 ON GRF_Trabalhos_Fases_Status(fk_trabalho,fk_fase,fk_trabalho_fase_status_tipo,dt_inclusao)
--*************

--*******************************************************************************************************************************************************************************

--***********************
-- Dados Tabelas Básicas
--***********************

--***** GRF_Fases
insert into GRF_Fases(codigo,nome,descricao,lead_time_perc_definido)
values(1,'RECEPÇÃO','',4.41)

insert into GRF_Fases(codigo,nome,descricao,lead_time_perc_definido)
values(2,'ARTE','',2.21)

insert into GRF_Fases(codigo,nome,descricao,lead_time_perc_definido)
values(3,'REVISÃO DA ARTE','',2.21)

insert into GRF_Fases(codigo,nome,descricao,lead_time_perc_definido)
values(4,'APROVAÇÃO DA ARTE','',13.23)

insert into GRF_Fases(codigo,nome,descricao,lead_time_perc_definido)
values(5,'RETOQUE','',21.17)

insert into GRF_Fases(codigo,nome,descricao,lead_time_perc_definido)
values(6,'PREPARAÇÃO','',10.59)

insert into GRF_Fases(codigo,nome,descricao,lead_time_perc_definido)
values(7,'REVISÃO DA PREPARAÇÃO','',3.57)

insert into GRF_Fases(codigo,nome,descricao,lead_time_perc_definido)
values(8,'PROVA','',17.65)

insert into GRF_Fases(codigo,nome,descricao,lead_time_perc_definido)
values(9,'APROVAÇÃO DA PROVA','',17.65)

insert into GRF_Fases(codigo,nome,descricao,lead_time_perc_definido)
values(10,'REVISÃO DE PROCEDIMENTO','',2.9)

insert into GRF_Fases(codigo,nome,descricao,lead_time_perc_definido)
values(11,'REVISÃO DIGITAL FINAL','',4.41)

insert into GRF_Fases(codigo,nome,descricao,lead_time_perc_definido)
values(12,'PADRÃO DE CORES','',0.00)

select * from GRF_Fases order by codigo

--***** GRF_Filiais
insert into GRF_Filiais(codigo,nome,descricao)
values(1,'Matriz','')

insert into GRF_Filiais(codigo,nome,descricao)
values(2,'Laminados','')

insert into GRF_Filiais(codigo,nome,descricao)
values(3,'Cumbica','')

insert into GRF_Filiais(codigo,nome,descricao)
values(4,'Anhanguera','')

insert into GRF_Filiais(codigo,nome,descricao)
values(5,'Cabreúva','')

select * from GRF_Filiais order by codigo

--***** GRF_Aprovacao_Tipos
insert into GRF_Aprovacao_Tipos(codigo,nome,descricao)
values(1,'Roland','')

insert into GRF_Aprovacao_Tipos(codigo,nome,descricao)
values(2,'GMS Externa','')

insert into GRF_Aprovacao_Tipos(codigo,nome,descricao)
values(3,'GMS Interna','')

insert into GRF_Aprovacao_Tipos(codigo,nome,descricao)
values(4,'Conferência','')

insert into GRF_Aprovacao_Tipos(codigo,nome,descricao)
values(5,'Contra-Prova','')

insert into GRF_Aprovacao_Tipos(codigo,nome,descricao)
values(6,'Teste','')

select * from GRF_Aprovacao_Tipos

--***** GRF_Avaliacao_Tipos
insert into GRF_Avaliacao_Tipos(codigo,nome,descricao)
values(1,'Aprovado','')

insert into GRF_Avaliacao_Tipos(codigo,nome,descricao)
values(2,'Reprovado','')

insert into GRF_Avaliacao_Tipos(codigo,nome,descricao)
values(3,'Restrição','')

select * from GRF_Avaliacao_Tipos

--***** GRF_Interrupcao_Tipos
insert into GRF_Interrupcao_Tipos(codigo,nome,descricao)
values(1,'Retrabalho Cliente','')

insert into GRF_Interrupcao_Tipos(codigo,nome,descricao)
values(2,'Retrabalho Zarapast','')

insert into GRF_Interrupcao_Tipos(codigo,nome,descricao)
values(3,'Ajuste Fábrica','')

insert into GRF_Interrupcao_Tipos(codigo,nome,descricao)
values(4,'Arte Complexa','')

insert into GRF_Interrupcao_Tipos(codigo,nome,descricao)
values(5,'Arte com Ajuste','')

insert into GRF_Interrupcao_Tipos(codigo,nome,descricao)
values(6,'Preparação com Ajuste','')

insert into GRF_Interrupcao_Tipos(codigo,nome,descricao)
values(7,'Prova Digital Urgente','')

insert into GRF_Interrupcao_Tipos(codigo,nome,descricao)
values(8,'Reunião','')

insert into GRF_Interrupcao_Tipos(codigo,nome,descricao)
values(9,'Aguardando Virar Sistema','')

insert into GRF_Interrupcao_Tipos(codigo,nome,descricao)
values(10,'Assistência Operacional','')

select * from GRF_Interrupcao_Tipos

--***** GRF_Trabalhos_Tipos
insert into GRF_Trabalhos_Tipos(codigo,nome,descricao)
values(1,'Novo','')

insert into GRF_Trabalhos_Tipos(codigo,nome,descricao)
values(2,'Alteração','')

insert into GRF_Trabalhos_Tipos(codigo,nome,descricao)
values(3,'Correção','')

select * from GRF_Trabalhos_Tipos

--***** GRF_Trabalhos_Fases_Status_Tipos

/*
  Fluxo da mudança de fases através dos status
  --------------------------------------------
  Exemplo:
   FASE X 
      Ações Automáticas
	      (INICIADO)
          (TRABALHANDO) - Ao Inserir registro nas tabelas de manutenção das fases (GRF_Trabalhos_Fases_ManutN),
		                  mantendo este status durante qualquer tipo de alteração nas referidas tabelas (Inclusão/Manutenção).
	  
	  Ações Operador
	      (AGUARDANDO) 
		  (RETROCER)    - FASE Y = Repete Fluxo FASE X
		  (LIBERAR)     - FASE Z = Repete Fluxo FASE X
   FASE 2
*/

insert into GRF_Trabalhos_Fases_Status_Tipos(codigo,nome,tipo_acao,descricao,acao_operador,interrompe_lead_time)
values(1,'INICIADO','AUTOMÁTICO','Ação automática do sistema para indicar que o trabalho chegou em determinada fase.',0,0)

insert into GRF_Trabalhos_Fases_Status_Tipos(codigo,nome,tipo_acao,descricao,acao_operador,interrompe_lead_time)
values(2,'TRABALHANDO','AUTOMÁTICO','Ação automática do sistema para indicar que o trabalho recebeu manutenção em determinada fase.',0,0)

insert into GRF_Trabalhos_Fases_Status_Tipos(codigo,nome,tipo_acao,descricao,acao_operador,interrompe_lead_time)
values(3,'AGUARDANDO','AGUARDAR','Ação indicada pelo operador responsável por determinada fase, onde permite o congelamento do tempo de trabalho na fase.',1,1)

insert into GRF_Trabalhos_Fases_Status_Tipos(codigo,nome,tipo_acao,descricao,acao_operador,interrompe_lead_time)
values(4,'RETROCEDIDO','RETROCEDER','Ação indicada pelo operador responsável por determinada fase, onde permite devolver o trabalho para uma fase anterior a ser escolhida no momento.',1,1)

insert into GRF_Trabalhos_Fases_Status_Tipos(codigo,nome,tipo_acao,descricao,acao_operador,interrompe_lead_time)
values(5,'LIBERADO','LIBERAR','Ação indicada pelo operador responsável por determinada fase, onde permite avançar o trabalho para a fase seguinte, encerrando o trabalho na fase atual.',1,1)

insert into GRF_Trabalhos_Fases_Status_Tipos(codigo,nome,tipo_acao,descricao,acao_operador,interrompe_lead_time)
values(6,'FINALIZADO PREPRESS','AUTOMÁTICO','Ação automática do sistema para indicar que o trabalho recebeu a liberação na última fase em que se apura o LEAD TIME.',0,1)

insert into GRF_Trabalhos_Fases_Status_Tipos(codigo,nome,tipo_acao,descricao,acao_operador,interrompe_lead_time)
values(7,'FINALIZADO TRABALHO','AUTOMÁTICO','Ação automática do sistema para indicar que o trabalho recebeu a liberação na última fase prática, sem interferência do LEAD TIME.',0,1)

insert into GRF_Trabalhos_Fases_Status_Tipos(codigo,nome,tipo_acao,descricao,acao_operador,interrompe_lead_time)
values(8,'INTERROMPIDO','INTERROMPER','Ação indicada pelo operador responsável por determinada fase, onde permite a interrupção do trabalho.',1,1)

insert into GRF_Trabalhos_Fases_Status_Tipos(codigo,nome,tipo_acao,descricao,acao_operador,interrompe_lead_time)
values(9,'GRAVACAO LIBERADA','AUTOMÁTICO','Ação automática do sistema para indicar que o operador informou ao trabalho a data de liberação para gravação.',0,1)

select * from GRF_Trabalhos_Fases_Status_Tipos

--***** GRF_Clientes

insert into GRF_Clientes(codigo,nome,nome_completo)
values(1,'Acroquima','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(2,'AB Brasil','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(3,'Agrocria','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(4,'Alicorp','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(5,'Alispec','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(6,'Alisul','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(7,'Angelo Auricchio','')

insert into GRF_Clientes(codigo,nome,nome_completo)
values(8,'Apti','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(9,'Arruda Alimentos','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(10,'Bagley','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(11,'Barbosa Marques','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(12,'Beba Brasil','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(13,'Bel Alimentos','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(14,'Big Brand','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(15,'Big Sal','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(16,'Bimbo','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(17,'Bioquima','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(18,'BR Brasil Foods','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(19,'Bunge','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(20,'Cacique','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(21,'Café Astro','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(22,'Café Barão','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(23,'Café Bom Dia','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(24,'Café São Braz','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(25,'Caiçara','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(26,'Camil','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(27,'Cargil','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(28,'Casa Suíça','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(29,'Cavnic','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(30,'Cepêra','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(31,'Cereale','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(32,'Cerrilos','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(33,'Cousa','')

insert into GRF_Clientes(codigo,nome,nome_completo)
values(34,'CRM Alimentos','')

insert into GRF_Clientes(codigo,nome,nome_completo)
values(35,'Dauper','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(36,'DSM Tortuga','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(37,'Ducoco','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(38,'Embare Alimentos','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(39,'Fabiani Saúde Animal','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(40,'Fazenda Sertãozinho','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(41,'Ferrero Rocher','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(42,'Fugini','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(43,'Garoto','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(44,'Gois Alimentos','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(45,'Goias Minas','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(46,'Granfino','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(47,'GSA','')

insert into GRF_Clientes(codigo,nome,nome_completo)
values(48,'Hersheys','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(49,'Invivo','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(50,'Itamaraty','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(51,'Itambé','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(52,'J. Macedo','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(53,'Jaguari','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(54,'Junior Alimentos','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(55,'Kelloggs','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(56,'Kimberly','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(57,'Kisabor','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(58,'Kobber','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(59,'Laticínios','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(60,'Liane','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(61,'Liotécnica','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(62,'Live Alimentos','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(63,'M Dias Branco','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(64,'Marilan','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(65,'Mariza','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(66,'Mart Minas','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(67,'Masterfoods','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(68,'Mavi','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(69,'Mecano Pack','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(70,'Mercur','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(71,'Milho de Ouro','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(72,'Mogiana','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(73,'Mondelez','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(74,'Montevergine','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(75,'Morrinhos','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(76,'Motrisa','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(77,'Nautamares','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(78,'Nestlé','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(79,'Netuno','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(80,'Ninfa','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(81,'Nutrimental','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(82,'Nutrisul','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(83,'Odrebrecht','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(84,'Pandurata','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(85,'Paraiso Nutrição Animal','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(86,'Parati','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(87,'Perfetti','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(88,'Pirahy','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(89,'Polenghi','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(90,'Pontevedra Alimentos','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(91,'Predilecta','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(92,'Premier','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(93,'Química Amparo','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(94,'Ração Reis','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(95,'Ruston','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(96,'Safeeds','')

insert into GRF_Clientes(codigo,nome,nome_completo)
values(97,'Sanchez Cano','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(98,'Santo André','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(99,'Sara Lee','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(100,'Selmi','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(101,'Serra da Grama','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(102,'Siol Alimentos','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(103,'Sta Amalia','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(104,'Stella Doro','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(105,'Tirol','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(106,'Tirolez','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(107,'Total Alimentos','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(108,'Três Corações','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(109,'Três de Maio','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(110,'United Mills','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(111,'Upsite Alimentos','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(112,'Val Alimentos','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(113,'Vigor','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(114,'Vitamais','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(115,'Wilson','') 

insert into GRF_Clientes(codigo,nome,nome_completo)
values(116,'Zaraplast','') 

select * from GRF_Clientes

