
/*
==============================================
Retorna Todos os Status de Todos os Trabalhos 
==============================================
*/

IF OBJECT_ID ( 'dbo.VIW_GRF_Trabalhos_Status') IS NOT NULL DROP VIEW dbo.VIW_GRF_Trabalhos_Status
GO  

CREATE VIEW dbo.VIW_GRF_Trabalhos_Status 

WITH ENCRYPTION
  
AS

select     st.pk_trabalho_fase_status
		  ,st.fk_trabalho_fase_status_tipo
          ,stt.nome                          tipo_status
          ,t.pk_trabalho 
		  ,t.num_pedido                      numero_pedido
		  ,fs.pk_fase
		  ,fs.codigo                         numero_fase
		  ,fs.nome                           nome_fase
		  ,st.fk_operador  
		  ,op.VCH_NOME                       nome_operador
	  	  ,st.obs_status                     observacao_status
	  	  ,st.dt_inclusao                    data_inclusao_status
		  ,stt.acao_operador
		  ,stt.interrompe_lead_time        
from       GRF_Trabalhos_Fases_Status        st
inner join GRF_Trabalhos_Fases_Status_Tipos  stt on  stt.pk_trabalho_fase_status_tipo = st.fk_trabalho_fase_status_tipo
inner join VIW_GRF_Trabalhos                 t   on  t.pk_trabalho                    = st.fk_trabalho
inner join GRF_Fases                         fs  on  fs.pk_fase                       = st.fk_fase
left  join APL_Usuarios                      op  on  op.PK_USUARIO                    = st.fk_operador


