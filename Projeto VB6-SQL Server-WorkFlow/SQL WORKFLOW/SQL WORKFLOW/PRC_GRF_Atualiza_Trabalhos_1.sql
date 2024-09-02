
/*
======================================================
Tipo de Atualização - 1 (Inclusão de Trabalhos Novos)  
======================================================
*/

IF OBJECT_ID ( 'dbo.PRC_GRF_Atualiza_Trabalhos_1') IS NOT NULL DROP PROCEDURE dbo.PRC_GRF_Atualiza_Trabalhos_1
GO  

CREATE PROCEDURE dbo.PRC_GRF_Atualiza_Trabalhos_1 
                            (
							 -- GRF_Trabalhos
 							 @fk_trabalho_tipo     int
							,@fk_aprovacao_tipo    int
							,@fk_cliente           int
							,@fk_filial            int
							,@num_pedido           varchar(60)  
							,@num_pedido_antigo    varchar(60)  
							,@num_pedido_novo      varchar(60)  
							,@lead_time_programado varchar(6)          
							
							-- GRF_Produtos
							,@produto_nome         varchar(500)     

							-- GRF_Representantes
							,@representante_nome   varchar(150) 

							-- GRF_Trabalhos_Fases_Manut1
							,@fk_fase              int          
							,@fk_operador          int          
							,@cores_alteradas      int          
							,@circ_cilindro        int          
							,@observacoes          varchar(max) 
							)

WITH ENCRYPTION
AS

declare @pk_trabalho                              int
	   ,@pk_produto                               int
	   ,@pk_representante                         int
	   ,@pk_trabalho_fase_manut1                  int
	   ,@fk_Trabalho_Fase_Status_Tipo_INICIADO    int   
	   ,@fk_Trabalho_Fase_Status_Tipo_TRABALHANDO int      
	   ,@codigo_novo                              int
	   ,@dt_inclusao                              datetime

SET NOCOUNT    ON
SET XACT_ABORT ON  

set    @dt_inclusao = getdate()
select @fk_Trabalho_Fase_Status_Tipo_INICIADO    = pk_Trabalho_Fase_Status_Tipo from GRF_Trabalhos_Fases_Status_Tipos where upper(nome) = 'INICIADO' 
select @fk_Trabalho_Fase_Status_Tipo_TRABALHANDO = pk_Trabalho_Fase_Status_Tipo from GRF_Trabalhos_Fases_Status_Tipos where upper(nome) = 'TRABALHANDO' 

BEGIN TRY 
	BEGIN TRANSACTION
		-- Localiza ou Cadastra Produto Inexistente Através do Nome
		select @pk_produto = pk_produto from GRF_Produtos where nome = @produto_nome

		if @pk_produto is null begin
		   select @codigo_novo = isnull(max(codigo),0) + 1 from GRF_Produtos	
   
		   insert into GRF_Produtos
		   (
			 codigo
			,nome
			,descricao
			,caminho_foto
		   )
		   values
		   (
			 @codigo_novo
			,@produto_nome
			,null
			,null
		   )

		   set @pk_produto = @@IDENTITY
		end 

		-- Localiza ou Cadastra Representante Inexistente Através do Nome
		select @pk_representante = pk_representante from GRF_Representantes where nome = @representante_nome

		if @pk_representante is null begin
		   select @codigo_novo = isnull(max(codigo),0) + 1 from GRF_Representantes
   
		   insert into GRF_Representantes
		   (
			 codigo
			,nome
			,nome_completo
		   )
		   values
		   (
			 @codigo_novo
			,@representante_nome
			,null
		   )

		   set @pk_representante = @@IDENTITY
		end

		-- Insere Trabalho Novo 
		insert into GRF_Trabalhos
		(
			 fk_trabalho_alt_seq
            ,fk_trabalho_tipo
			,fk_produto
			,fk_representante
			,fk_aprovacao_tipo
			,fk_cliente
			,fk_filial
			,num_pedido
			,num_pedido_antigo
			,num_pedido_novo
			,lead_time_programado
			,dt_inclusao
			,dt_alteracao
			,dt_exclusao
		)
		values 							 
		(
		    null
		   ,@fk_trabalho_tipo  
		   ,@pk_produto
		   ,@pk_representante   
		   ,@fk_aprovacao_tipo
		   ,@fk_cliente
		   ,@fk_filial    
		   ,@num_pedido           
		   ,@num_pedido_antigo    
		   ,@num_pedido_novo      
		   ,@lead_time_programado 
		   ,@dt_inclusao
		   ,null
		   ,null        
		)

		set @pk_trabalho = @@IDENTITY

		-- Insere Fase de Manutenção 1
		insert into GRF_Trabalhos_Fases_Manut1
		(
		    fk_trabalho_fase_manut1_alt_seq
		   ,fk_trabalho
		   ,fk_fase
		   ,fk_operador
		   ,cores_alteradas
		   ,circ_cilindro
		   ,observacoes
		   ,dt_inclusao
		   ,dt_alteracao
		)
		values 							 
		(
		    null
		   ,@pk_trabalho
		   ,@fk_fase
		   ,@fk_operador
		   ,@cores_alteradas
		   ,@circ_cilindro
		   ,@observacoes
		   ,@dt_inclusao
		   ,null
		)

		set @pk_trabalho_fase_manut1 = @@IDENTITY

		-- Insere Fase Status - INICIADO
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
			@pk_trabalho
		   ,@fk_fase
		   ,@fk_Trabalho_Fase_Status_Tipo_INICIADO
		   ,null
		   ,'PRC_GRF_Atualiza_Trabalhos_1'
		   ,@dt_inclusao
		)

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
			@pk_trabalho
		   ,@fk_fase
		   ,@fk_Trabalho_Fase_Status_Tipo_TRABALHANDO
		   ,null
		   ,'PRC_GRF_Atualiza_Trabalhos_1'
		   ,@dt_inclusao
		)

	COMMIT TRANSACTION

	select  '3'                                  Status
		   ,'Atualização realizada com sucesso!' mensagem          		               
	       ,@pk_trabalho                         pk_trabalho                       
	       ,@pk_produto                          pk_produto                   
	       ,@pk_representante                    pk_representante                        
	       ,@pk_trabalho_fase_manut1             pk_trabalho_fase_manut1   

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

   