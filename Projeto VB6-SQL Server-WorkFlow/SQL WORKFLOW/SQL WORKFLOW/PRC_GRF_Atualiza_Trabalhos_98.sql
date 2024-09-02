
/*
=================================================================
Tipo de Atualização - 98 (Inclusão/Alteração - Campos Genéricos)
=================================================================
*/

IF OBJECT_ID ( 'dbo.PRC_GRF_Atualiza_Trabalhos_98') IS NOT NULL DROP PROCEDURE dbo.PRC_GRF_Atualiza_Trabalhos_98
GO  

CREATE PROCEDURE dbo.PRC_GRF_Atualiza_Trabalhos_98
                            (
							-- GRF_Trabalhos_Fases_Manut_Generico
							 @pk_trabalho_fase_manut_gen       int = NULL        
							,@fk_trabalho                      int = NULL      
							,@fk_fase                          int = NULL      
							,@fk_operador                      int          
							,@dt_liberacao_gravacao            datetime
							)

WITH ENCRYPTION
AS

declare  @fk_Trabalho_Fase_Status_Tipo_GRAVACAO_LIBERADA int 
		,@fk_trabalho_fase_manut_gen_alt_seq             int    
	    ,@codigo_novo                                    int
		,@fk_Trabalho_Fase_Status_Tipo_ATUAL             int
	    ,@dt_inclusao                                    datetime

SET NOCOUNT    ON
SET XACT_ABORT ON  

-- Verifica Parâmetros Principais
if @pk_trabalho_fase_manut_gen is null begin
	if @fk_trabalho is null begin
		select '6' Status,'Se não informar o parâmetro < @pk_trabalho_fase_manut_gen >, informe obrigatoriamente o parâmetro < @fk_trabalho >!' mensagem
		return
	end 
	if @fk_fase is null begin
		select '6' Status,'Se não informar o parâmetro < @pk_trabalho_fase_manut_gen >, informe obrigatoriamente o parâmetro < @fk_fase >!' mensagem
		return
	end
end
else begin
	if not exists(select 1 from VIW_GRF_Trabalhos_Fases_Manut_Generico where pk_trabalho_fase_manut_gen = @pk_trabalho_fase_manut_gen) begin
		select '6' Status, 'O registro < @pk_trabalho_fase_manut_gen > = ' + convert(varchar,@pk_trabalho_fase_manut_gen) + ', não foi localizado na view < VIW_GRF_Trabalhos_Fases_Manut_Generico >!' mensagem
		return
	end
	if @fk_trabalho is not null begin
		select '6' Status,'Se informar o parâmetro < @pk_trabalho_fase_manut_gen >, não informar o parâmetro < @fk_trabalho >!' mensagem
		return
	end 
	if @fk_fase is not null begin
		select '6' Status,'Se informar o parâmetro < @pk_trabalho_fase_manut_gen >, não informar o parâmetro < @fk_fase >!' mensagem
		return
	end 
end

set    @dt_inclusao = getdate()
select @fk_Trabalho_Fase_Status_Tipo_GRAVACAO_LIBERADA = pk_Trabalho_Fase_Status_Tipo from GRF_Trabalhos_Fases_Status_Tipos where upper(nome) = 'GRAVACAO LIBERADA' 

BEGIN TRY 
	BEGIN TRANSACTION
		if @pk_trabalho_fase_manut_gen is null begin
			-- ***** INCLUSÃO

			-- Insere os Campos Genéricos
			insert into GRF_Trabalhos_Fases_Manut_Generico
			(
			 	 fk_trabalho_fase_manut_gen_alt_seq
				,fk_trabalho
				,fk_fase
				,fk_operador
				,dt_liberacao_gravacao
				,dt_inclusao
				,dt_alteracao
			)
			values 							 
			(
			 	 null
				,@fk_trabalho
				,@fk_fase
				,@fk_operador
				,@dt_liberacao_gravacao
				,@dt_inclusao
				,null
			)
			
			set @pk_trabalho_fase_manut_gen = @@IDENTITY

			-- Insere Fase Status - TRABALHANDO
			insert into GRF_Trabalhos_Fases_Status
			(
				 fk_trabalho
				,fk_fase
				,fk_Trabalho_Fase_Status_Tipo
				,fk_operador
				,obs_status
				,dt_inclusao
			)
			values 							 
			(
				 @fk_trabalho
				,@fk_fase
				,@fk_Trabalho_Fase_Status_Tipo_GRAVACAO_LIBERADA
				,null
				,'PRC_GRF_Atualiza_Trabalhos_98'
				,@dt_inclusao
		   	)
		end
		else begin
			-- ***** ALTERAÇÃO

			-- Insere os Campos Genéricos Antes da Alteração (Histórico)
			select @fk_trabalho = fk_trabalho
			      ,@fk_fase     = fk_fase 
			from   GRF_Trabalhos_Fases_Manut_Generico 
			where  pk_trabalho_fase_manut_gen = @pk_trabalho_fase_manut_gen

			select @fk_trabalho_fase_manut_gen_alt_seq = max(pk_trabalho_fase_manut_gen) 
			from   GRF_Trabalhos_Fases_Manut_Generico 
			where  fk_trabalho = @fk_trabalho
			and    fk_fase     = @fk_fase

			select @fk_Trabalho_Fase_Status_Tipo_ATUAL = fk_trabalho_fase_status_tipo 
			from   VIW_GRF_Trabalhos_Status 
			where  pk_trabalho_fase_status = dbo.FNC_GRF_Trabalho_Fase_Status_Atual(@fk_trabalho)

			insert into GRF_Trabalhos_Fases_Manut_Generico
			(
				 fk_trabalho_fase_manut_gen_alt_seq
				,fk_trabalho
				,fk_fase
				,fk_operador
				,dt_liberacao_gravacao
				,dt_inclusao
				,dt_alteracao
			)
			select @fk_trabalho_fase_manut_gen_alt_seq  -- Guarda a Referência do Registro Sequencialmente
				  ,fk_trabalho
				  ,fk_fase
				  ,fk_operador
				  ,dt_liberacao_gravacao
				  ,@dt_inclusao
				  ,null
			from   GRF_Trabalhos_Fases_Manut_Generico
			where  pk_trabalho_fase_manut_gen = @pk_trabalho_fase_manut_gen 

			-- Altera os Campos Genéricos
			update GRF_Trabalhos_Fases_Manut_Generico         
			set    fk_operador                = @fk_operador
			 	  ,dt_liberacao_gravacao      = @dt_liberacao_gravacao
				  ,dt_alteracao               = @dt_inclusao
			where  pk_trabalho_fase_manut_gen = @pk_trabalho_fase_manut_gen

			IF @fk_Trabalho_Fase_Status_Tipo_ATUAL <> @fk_Trabalho_Fase_Status_Tipo_GRAVACAO_LIBERADA begin
				-- Insere Fase Status - TRABALHANDO
				insert into GRF_Trabalhos_Fases_Status
				(
					 fk_trabalho
					,fk_fase
					,fk_Trabalho_Fase_Status_Tipo
					,fk_operador
					,obs_status
					,dt_inclusao
				)
				values 							 
				(
					 @fk_trabalho
					,@fk_fase
					,@fk_Trabalho_Fase_Status_Tipo_GRAVACAO_LIBERADA
					,null
					,'PRC_GRF_Atualiza_Trabalhos_98'
					,@dt_inclusao
		   		)
			end
		end 

	COMMIT TRANSACTION

	select  '3'                                  Status
	       ,'Atualização realizada com sucesso!' mensagem                  
	       ,@pk_trabalho_fase_manut_gen          pk_trabalho_fase_manut_gen
	
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

   