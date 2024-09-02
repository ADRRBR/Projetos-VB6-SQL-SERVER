
/*
======================================================
Tipo de Atualiza��o - 4 (Inclus�o/Altera��o - Fase 3)
======================================================
*/

IF OBJECT_ID ( 'dbo.PRC_GRF_Atualiza_Trabalhos_4') IS NOT NULL DROP PROCEDURE dbo.PRC_GRF_Atualiza_Trabalhos_4
GO  

CREATE PROCEDURE dbo.PRC_GRF_Atualiza_Trabalhos_4
                            (
							-- GRF_Trabalhos_Fases_Manut3
							 @pk_trabalho_fase_manut3          int = NULL        
							,@fk_trabalho                      int = NULL      
							,@fk_fase                          int = NULL      
							,@fk_operador                      int          
							,@observacoes                      varchar(max) 
							)

WITH ENCRYPTION
AS

declare  @fk_Trabalho_Fase_Status_Tipo_TRABALHANDO int 
		,@fk_trabalho_fase_manut3_alt_seq          int    
	    ,@codigo_novo                              int
		,@fk_Trabalho_Fase_Status_Tipo_ATUAL       int
	    ,@dt_inclusao                              datetime

SET NOCOUNT    ON
SET XACT_ABORT ON  

-- Verifica Par�metros Principais
if @pk_trabalho_fase_manut3 is null begin
	if @fk_trabalho is null begin
		select '6' Status,'Se n�o informar o par�metro < @pk_trabalho_fase_manut3 >, informe obrigatoriamente o par�metro < @fk_trabalho >!' mensagem
		return
	end 
	if @fk_fase is null begin
		select '6' Status,'Se n�o informar o par�metro < @pk_trabalho_fase_manut3 >, informe obrigatoriamente o par�metro < @fk_fase >!' mensagem
		return
	end
end
else begin
	if not exists(select 1 from VIW_GRF_Trabalhos_Fases_Manut3 where pk_trabalho_fase_manut3 = @pk_trabalho_fase_manut3) begin
		select '6' Status, 'O registro < @pk_trabalho_fase_manut3 > = ' + convert(varchar,@pk_trabalho_fase_manut3) + ', n�o foi localizado na view < VIW_GRF_Trabalhos_Fases_Manut3 >!' mensagem
		return
	end
	if @fk_trabalho is not null begin
		select '6' Status,'Se informar o par�metro < @pk_trabalho_fase_manut3 >, n�o informar o par�metro < @fk_trabalho >!' mensagem
		return
	end 
	if @fk_fase is not null begin
		select '6' Status,'Se informar o par�metro < @pk_trabalho_fase_manut3 >, n�o informar o par�metro < @fk_fase >!' mensagem
		return
	end 
end

set    @dt_inclusao = getdate()
select @fk_Trabalho_Fase_Status_Tipo_TRABALHANDO = pk_Trabalho_Fase_Status_Tipo from GRF_Trabalhos_Fases_Status_Tipos where upper(nome) = 'TRABALHANDO' 

BEGIN TRY 
	BEGIN TRANSACTION
		if @pk_trabalho_fase_manut3 is null begin
			-- ***** INCLUS�O

			-- Insere Fase de Manuten��o 3
			insert into GRF_Trabalhos_Fases_Manut3
			(
			 	 fk_trabalho_fase_manut3_alt_seq
				,fk_trabalho
				,fk_fase
				,fk_operador
				,observacoes
				,dt_inclusao
				,dt_alteracao
			)
			values 							 
			(
			 	 null
				,@fk_trabalho
				,@fk_fase
				,@fk_operador
				,@observacoes
				,@dt_inclusao
				,null
			)
			
			set @pk_trabalho_fase_manut3 = @@IDENTITY

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
				,@fk_Trabalho_Fase_Status_Tipo_TRABALHANDO
				,null
				,'PRC_GRF_Atualiza_Trabalhos_4'
				,@dt_inclusao
		   	)
		end
		else begin
			-- ***** ALTERA��O

			-- Insere a Fase de Manuten��o 3 Antes da Altera��o (Hist�rico)
			select @fk_trabalho = fk_trabalho
			      ,@fk_fase     = fk_fase 
			from   GRF_Trabalhos_Fases_Manut3 
			where  pk_trabalho_fase_manut3 = @pk_trabalho_fase_manut3

			select @fk_trabalho_fase_manut3_alt_seq = max(pk_trabalho_fase_manut3) 
			from   GRF_Trabalhos_Fases_Manut3 
			where  fk_trabalho = @fk_trabalho
			and    fk_fase     = @fk_fase

			select @fk_Trabalho_Fase_Status_Tipo_ATUAL = fk_trabalho_fase_status_tipo 
			from   VIW_GRF_Trabalhos_Status 
			where  pk_trabalho_fase_status = dbo.FNC_GRF_Trabalho_Fase_Status_Atual(@fk_trabalho)

			insert into GRF_Trabalhos_Fases_Manut3
			(
				 fk_trabalho_fase_manut3_alt_seq
				,fk_trabalho
				,fk_fase
				,fk_operador
				,observacoes
				,dt_inclusao
				,dt_alteracao
			)
			select @fk_trabalho_fase_manut3_alt_seq  -- Guarda a Refer�ncia do Registro Sequencialmente
				  ,fk_trabalho
				  ,fk_fase
				  ,fk_operador
				  ,observacoes
				  ,@dt_inclusao
				  ,null
			from   GRF_Trabalhos_Fases_Manut3
			where  pk_trabalho_fase_manut3 = @pk_trabalho_fase_manut3 

			-- Altera a Fase de Manuten��o 3
			update GRF_Trabalhos_Fases_Manut3         
			set    fk_operador             = @fk_operador
				  ,observacoes             = @observacoes
				  ,dt_alteracao            = @dt_inclusao
			where  pk_trabalho_fase_manut3 = @pk_trabalho_fase_manut3

			IF @fk_Trabalho_Fase_Status_Tipo_ATUAL <> @fk_Trabalho_Fase_Status_Tipo_TRABALHANDO begin
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
					,@fk_Trabalho_Fase_Status_Tipo_TRABALHANDO
					,null
					,'PRC_GRF_Atualiza_Trabalhos_4'
					,@dt_inclusao
		   		)
			end
		end 

	COMMIT TRANSACTION

	select  '3'                                  Status
	       ,'Atualiza��o realizada com sucesso!' mensagem                  
	       ,@pk_trabalho_fase_manut3             pk_trabalho_fase_manut3
	
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

   