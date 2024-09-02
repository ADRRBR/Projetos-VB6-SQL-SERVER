
/*
============================================================
Tipo de Consulta: Histórico de Status das Fases do Trabalho 
============================================================
*/

IF OBJECT_ID ( 'dbo.PRC_GRF_Consulta_Status_Fases_Trabalhos_Hist') IS NOT NULL DROP PROCEDURE dbo.PRC_GRF_Consulta_Status_Fases_Trabalhos_Hist
GO  

CREATE PROCEDURE dbo.PRC_GRF_Consulta_Status_Fases_Trabalhos_Hist (
												                   @pk_trabalho int		
							                                      ,@pk_fase     int = null
							                                      )

WITH ENCRYPTION
AS

SET NOCOUNT    ON
SET XACT_ABORT ON  

select   pk_trabalho_fase_status        pk_trabalho_fase_status
		,fk_trabalho_fase_status_tipo   fk_trabalho_fase_status_tipo
		,tipo_status                    [Status]
		,pk_trabalho                    pk_trabalho            
		,numero_pedido                  [Número Pedido]
		,pk_fase                        pk_fase
		,numero_fase                    [Número Fase]
		,nome_fase                      [Fase]
		,fk_operador                    fk_operador                    
		,nome_operador                  Colaborador
		,observacao_status              [Obs.Status]
		,data_inclusao_status           [Dt.Incl.Status]
		,case interrompe_lead_time 
			when 0 then 'Não' 
			when 1 then 'Sim' 
			else ''  
  	     end                            [Interrompe LEAD TIME?]
from     VIW_GRF_Trabalhos_Status
where    pk_trabalho = @pk_trabalho
and      pk_fase     = isnull(@pk_fase,pk_fase)
order by pk_trabalho_fase_status desc


