
/*
=============================================================================================
Tipo de Consulta: Trabalho Fases Manutenção - Consulta 2 (TRAB.EM DETERMINADA FASE E STATUS)
=============================================================================================
*/

IF OBJECT_ID ( 'dbo.PRC_GRF_Consulta_Trabalhos_2') IS NOT NULL DROP PROCEDURE dbo.PRC_GRF_Consulta_Trabalhos_2
GO  

CREATE PROCEDURE dbo.PRC_GRF_Consulta_Trabalhos_2 (
							                       @pk_fase                      int
												  ,@fk_trabalho_fase_status_tipo int	
							                      )

WITH ENCRYPTION
AS

SET NOCOUNT    ON
SET XACT_ABORT ON  

select   TB_pk_trabalho                               TB_pk_trabalho
	    ,TB_num_pedido                                [Número Pedido]
	    ,TB_num_pedido_antigo                         [Número Antigo]
	    ,TB_num_pedido_novo                           [Número Novo]
	    ,TB_R_TT_tipo_trabalho                        [Tipo Trabalho] 
		,TB_R_P_produto                               [Produto]
		,TB_R_C_cliente                               [Cliente]
		,TB_R_R_representante                         [Representante]
		,TB_R_AT_aprovacao_tipo                       [Tipo Aprovação]
		,TB_lead_time_programado                      [LEAD TIME Programado]
		,TB_dt_inclusao                               [Dt.Incl.Recepção]
		,TB_dt_alteracao                              [Dt.Últ.Alt.Recepção]
		---------------------------------------------------------------------------
		,FM1_R_F_fase                                 [Fase 1]
		,FM1_R_O_operador                             [Operador Fase 1]
		,FM1_dt_inclusao                              [Dt.Incl.Fase 1]
		,FM1_dt_alteracao                             [Dt.Últ.Alt.Fase 1]
		---------------------------------------------------------------------------
		,FM2_R_F_fase                                 [Fase 2]
		,FM2_R_O_operador                             [Operador Fase 2]
		,FM2_dt_inclusao                              [Dt.Incl.Fase 2]
		,FM2_dt_alteracao                             [Dt.Últ.Alt.Fase 2]
		---------------------------------------------------------------------------
		,FM3_R_F_fase                                 [Fase 3]
		,FM3_R_O_operador                             [Operador Fase 3]
		,FM3_dt_inclusao                              [Dt.Incl.Fase 3]
		,FM3_dt_alteracao                             [Dt.Últ.Alt.Fase 3]
		---------------------------------------------------------------------------
		,FM4_R_F_fase                                 [Fase 4]
		,FM4_R_O_operador                             [Operador Fase 4]
		,FM4_dt_inclusao                              [Dt.Incl.Fase 4]
		,FM4_dt_alteracao                             [Dt.Últ.Alt.Fase 4]
		---------------------------------------------------------------------------
		,FM5_R_F_fase                                 [Fase 5]
		,FM5_R_O_operador                             [Operador Fase 5]
		,FM5_dt_inclusao                              [Dt.Incl.Fase 5]
		,FM5_dt_alteracao                             [Dt.Últ.Alt.Fase 5]
		---------------------------------------------------------------------------
		,FM6_R_F_fase                                 [Fase 6]
		,FM6_R_O_operador                             [Operador Fase 6]
		,FM6_dt_inclusao                              [Dt.Incl.Fase 6]
		,FM6_dt_alteracao                             [Dt.Últ.Alt.Fase 6]
		---------------------------------------------------------------------------
		,FM7_R_F_fase                                 [Fase 7]
		,FM7_R_O_operador                             [Operador Fase 7]
		,FM7_dt_inclusao                              [Dt.Incl.Fase 7]
		,FM7_dt_alteracao                             [Dt.Últ.Alt.Fase 7]
		---------------------------------------------------------------------------
		,FM8_R_F_fase                                 [Fase 8]
		,FM8_R_O_operador                             [Operador Fase 8]
		,FM8_dt_inclusao                              [Dt.Incl.Fase 8]
		,FM8_dt_alteracao                             [Dt.Últ.Alt.Fase 8]
		---------------------------------------------------------------------------
		,FM9_R_F_fase                                 [Fase 9]
		,FM9_R_O_operador                             [Operador Fase 9]
		,FM9_dt_inclusao                              [Dt.Incl.Fase 9]
		,FM9_dt_alteracao                             [Dt.Últ.Alt.Fase 9]
		---------------------------------------------------------------------------
		,FM10_R_F_fase                                [Fase 10]
		,FM10_R_O_operador                            [Operador Fase 10]
		,FM10_dt_inclusao                             [Dt.Incl.Fase 10]
		,FM10_dt_alteracao                            [Dt.Últ.Alt.Fase 10]
		---------------------------------------------------------------------------
		,FM11_R_F_fase                                [Fase 11]
		,FM11_R_O_operador                            [Operador Fase 11]
		,FM11_dt_inclusao                             [Dt.Incl.Fase 11]
		,FM11_dt_alteracao                            [Dt.Últ.Alt.Fase 11]
		---------------------------------------------------------------------------
		,FM12_R_F_fase                                [Fase 12]
		,FM12_R_O_operador                            [Operador Fase 12]
		,FM12_dt_inclusao                             [Dt.Incl.Fase 12]
		,FM12_dt_alteracao                            [Dt.Últ.Alt.Fase 12]
		---------------------------------------------------------------------------
		,TFSO_nome_fase_operador                      [Fase Inf.Colab.]
		,TFSO_tipo_status_operador                    [Status Inf.Colab.]
		,TFSO_observacao_status                       [Obs.Status.Inf.Colab.] 
		,TFSO_data_inclusao_status                    [Dt.Incl.Status Inf.Colab.] 
		,TFSO_nome_operador_status                    [Colab.Inf.Status] 
		,TFSO_pk_trabalho_fase_status_operador        TFSO_pk_trabalho_fase_status_operador 
		,TFSO_fk_trabalho_fase_status_tipo_operador   TFSO_fk_trabalho_fase_status_tipo_operador
		,TFSO_codigo_fase_operador                    TFSO_codigo_fase_operador
		---------------------------------------------------------------------------
		,TFSA_nome_fase_atual                         [Fase Atual]
		,TFSA_tipo_status_atual                       [Status Fase Atual]
		,GEN_dt_liberacao_gravacao                    [Dt.Lib.Gravação]
		,TFSA_pk_trabalho_fase_status_atual           TFSA_pk_trabalho_fase_status_atual
		,TFSA_fk_trabalho_fase_status_tipo_atual      TFSA_fk_trabalho_fase_status_tipo_atual
		,TFSA_pk_fase_atual                           TFSA_pk_fase_atual
		,TFSA_codigo_fase_atual						  TFSA_codigo_fase_atual
		---------------------------------------------------------------------------
from    VIW_GRF_Trabalhos_Fases_Manut
where   TFSA_pk_fase_atual                      = @pk_fase
and     TFSA_fk_trabalho_fase_status_tipo_atual = @fk_trabalho_fase_status_tipo
