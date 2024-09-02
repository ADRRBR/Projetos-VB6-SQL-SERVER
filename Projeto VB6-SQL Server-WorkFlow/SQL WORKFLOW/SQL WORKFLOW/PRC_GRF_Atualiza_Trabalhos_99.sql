
/*
===========================================
Tipo de Atualização - 99 (Mudança de Fase)  
===========================================
*/

IF OBJECT_ID ( 'dbo.PRC_GRF_Atualiza_Trabalhos_99') IS NOT NULL DROP PROCEDURE dbo.PRC_GRF_Atualiza_Trabalhos_99
GO  

CREATE PROCEDURE dbo.PRC_GRF_Atualiza_Trabalhos_99 
                            (
							 -- GRF_Trabalhos_Fases_Status
							 @fk_trabalho                    int
							,@fk_fase                        int 
							,@fk_trabalho_fase_status_tipo   int
							,@fk_operador                    int
							,@obs_status                     varchar(max)
							)

WITH ENCRYPTION
AS

declare  @fk_fase_Atual                                    int
	    ,@fk_trabalho_fase_status_tipo_INICIADO            int
		,@fk_Trabalho_Fase_Status_Tipo_LIBERADO            int
		,@fk_Trabalho_Fase_Status_Tipo_INTERROMPIDO        int
		,@fk_Trabalho_Fase_Status_Tipo_FINALIZADO_PREPRESS int
		,@fk_Trabalho_Fase_Status_Tipo_FINALIZADO_TRABALHO int
		,@pk_ultima_fase_lead_time                         int
		,@pk_ultima_fase_trabalho                          int
	    ,@dt_inclusao                                      datetime

SET NOCOUNT    ON
SET XACT_ABORT ON  

-- Verifica Parâmetros Principais
if not exists(select 1 from VIW_GRF_Trabalhos where pk_trabalho = @fk_trabalho) begin
	select '6' Status, 'O registro < @pk_trabalho > = ' + convert(varchar,@fk_trabalho) + ', não foi localizado na view < VIW_GRF_Trabalhos >!' mensagem
	return
end

set    @dt_inclusao = getdate()
select @fk_fase_Atual = isnull(dbo.FNC_GRF_Trabalho_Fase_Atual(@fk_trabalho),0)
select @fk_Trabalho_Fase_Status_Tipo_INICIADO            = pk_Trabalho_Fase_Status_Tipo from GRF_Trabalhos_Fases_Status_Tipos where upper(nome) = 'INICIADO' 
select @fk_Trabalho_Fase_Status_Tipo_LIBERADO            = pk_Trabalho_Fase_Status_Tipo from GRF_Trabalhos_Fases_Status_Tipos where upper(nome) = 'LIBERADO' 
select @fk_Trabalho_Fase_Status_Tipo_INTERROMPIDO        = pk_Trabalho_Fase_Status_Tipo from GRF_Trabalhos_Fases_Status_Tipos where upper(nome) = 'INTERROMPIDO' 
select @fk_Trabalho_Fase_Status_Tipo_FINALIZADO_PREPRESS = pk_Trabalho_Fase_Status_Tipo from GRF_Trabalhos_Fases_Status_Tipos where upper(nome) = 'FINALIZADO PREPRESS' 
select @fk_Trabalho_Fase_Status_Tipo_FINALIZADO_TRABALHO = pk_Trabalho_Fase_Status_Tipo from GRF_Trabalhos_Fases_Status_Tipos where upper(nome) = 'FINALIZADO TRABALHO' 
select @pk_ultima_fase_lead_time = pk_fase from GRF_Fases where codigo = 11
select @pk_ultima_fase_trabalho  = pk_fase from GRF_Fases where codigo = 12

BEGIN TRY 
	BEGIN TRANSACTION
		-- Insere Fase Status Informado
		insert into GRF_Trabalhos_Fases_Status
		(
		    fk_trabalho
		   ,fk_fase
		   ,fk_trabalho_fase_status_tipo
		   ,fk_operador
		   ,obs_status
		   ,dt_inclusao
		)
		values 							 
		(
			@fk_trabalho
		   ,@fk_fase_Atual
		   ,@fk_trabalho_fase_status_tipo
		   ,@fk_operador
		   ,@obs_status
		   ,@dt_inclusao
		)

		-- Demais Status (Exceto Interrupção)
		if @fk_trabalho_fase_status_tipo <> @fk_Trabalho_Fase_Status_Tipo_INTERROMPIDO begin
			if @fk_fase <> 0 begin
				if @fk_trabalho_fase_status_tipo = @fk_Trabalho_Fase_Status_Tipo_LIBERADO begin
					if @fk_fase_Atual = @pk_ultima_fase_lead_time begin
						-- Insere Fase Status - FINALIZADO PREPRESS
						insert into GRF_Trabalhos_Fases_Status
						(
							fk_trabalho
						   ,fk_fase
						   ,fk_trabalho_fase_status_tipo
						   ,fk_operador
						   ,obs_status
						   ,dt_inclusao
						)
						values 							 
						(
							@fk_trabalho
						   ,@fk_fase_Atual
						   ,@fk_Trabalho_Fase_Status_Tipo_FINALIZADO_PREPRESS
						   ,null
						   ,'PRC_GRF_Atualiza_Trabalhos_99'
						   ,@dt_inclusao
						)			    
					end
				end

				-- Insere Fase Status - INICIADO
				insert into GRF_Trabalhos_Fases_Status
				(
					fk_trabalho
				   ,fk_fase
				   ,fk_trabalho_fase_status_tipo
				   ,fk_operador
				   ,obs_status
				   ,dt_inclusao
				)
				values 							 
				(
					@fk_trabalho
				   ,@fk_fase
				   ,@fk_Trabalho_Fase_Status_Tipo_INICIADO
				   ,null
				   ,'PRC_GRF_Atualiza_Trabalhos_99'
				   ,@dt_inclusao
				)
			end
   			else begin
				if @fk_trabalho_fase_status_tipo = @fk_Trabalho_Fase_Status_Tipo_LIBERADO begin
					if @fk_fase_Atual = @pk_ultima_fase_trabalho begin
						-- Insere Fase Status - FINALIZADO TRABALHO
						insert into GRF_Trabalhos_Fases_Status
						(
							fk_trabalho
						   ,fk_fase
						   ,fk_trabalho_fase_status_tipo
						   ,fk_operador
						   ,obs_status
						   ,dt_inclusao
						)
						values 							 
						(
							@fk_trabalho
						   ,@fk_fase_Atual
						   ,@fk_Trabalho_Fase_Status_Tipo_FINALIZADO_TRABALHO
						   ,null
						   ,'PRC_GRF_Atualiza_Trabalhos_99'
						   ,@dt_inclusao
						)	
					end
				end
			end
		end

 		select  '3'                                  Status
	           ,'Atualização realizada com sucesso!' mensagem   

	COMMIT TRANSACTION
	RETURN
END TRY 

-- Tratamento de Erros
BEGIN CATCH  
   DECLARE @ErrorMessage  NVARCHAR(4000)        
   DECLARE @ErrorSeverity INT        
   DECLARE @ErrorState    INT 
        
   SELECT  @ErrorMessage  = ERROR_PROCEDURE() + ' - Linha ' + CONVERT(VARCHAR(15),ERROR_LINE()) + ' - ' + ERROR_MESSAGE()        
          ,@ErrorSeverity = ERROR_SEVERITY()        
          ,@ErrorState    = ERROR_STATE()        
  
   RAISERROR (@ErrorMessage        
             ,@ErrorSeverity        
             ,@ErrorState )  

   ROLLBACK TRANSACTION	
END CATCH

   