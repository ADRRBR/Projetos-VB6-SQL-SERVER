
/*
===========================================================
Tipo de Atualização - 2 (Alteração de Trabalhos na Fase 1)  
===========================================================
*/

IF OBJECT_ID ( 'dbo.PRC_GRF_Atualiza_Trabalhos_2') IS NOT NULL DROP PROCEDURE dbo.PRC_GRF_Atualiza_Trabalhos_2
GO  

CREATE PROCEDURE dbo.PRC_GRF_Atualiza_Trabalhos_2 
                            (
							 -- GRF_Trabalhos
                             @pk_trabalho             int
 							,@fk_trabalho_tipo        int
                            ,@fk_aprovacao_tipo       int
							,@fk_cliente              int
							,@fk_filial               int
							,@num_pedido_antigo       varchar(60)  
							,@num_pedido_novo         varchar(60)  
							,@lead_time_programado    varchar(6)          
							
							-- GRF_Produtos
							,@produto_nome            varchar(500)     

							-- GRF_Representantes
							,@representante_nome      varchar(150) 

							-- GRF_Trabalhos_Fases_Manut1
							,@pk_trabalho_fase_manut1 int 
							,@fk_operador             int          
							,@cores_alteradas         int          
							,@circ_cilindro           int          
							,@observacoes             varchar(max) 
							)

WITH ENCRYPTION
AS

declare @pk_produto                               int
	   ,@pk_representante                         int
	   ,@num_pedido                               varchar(60)
	   ,@fk_trabalho_alt_seq                      int
	   ,@fk_fase                                  int
	   ,@fk_trabalho_fase_manut1_alt_seq          int
	   ,@codigo_novo                              int
	   ,@fk_Trabalho_Fase_Status_Tipo_TRABALHANDO int
	   ,@fk_Trabalho_Fase_Status_Tipo_ATUAL       int
	   ,@dt_inclusao                              datetime

SET NOCOUNT    ON
SET XACT_ABORT ON  

-- Verifica Parâmetros Principais
if not exists(select 1 from VIW_GRF_Trabalhos where pk_trabalho = @pk_trabalho) begin
	select '6' Status, 'O registro < @pk_trabalho > = ' + convert(varchar,@pk_trabalho) + ', não foi localizado na view < VIW_GRF_Trabalhos >!' mensagem
	return
end
if not exists(select 1 from VIW_GRF_Trabalhos_Fases_Manut1 where pk_trabalho_fase_manut1 = @pk_trabalho_fase_manut1) begin
	select '6' Status, 'O registro < @pk_trabalho_fase_manut1 > = ' + convert(varchar,@pk_trabalho_fase_manut1) + ', não foi localizado na view < VIW_GRF_Trabalhos_Fases_Manut1 >!' mensagem
	return
end

set    @dt_inclusao = getdate()
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

		-- Insere o Trabalho Antes da Alteração (Histórico)
		select @num_pedido = num_pedido 
		from   GRF_Trabalhos 
		where  pk_trabalho = @pk_trabalho

		select @fk_trabalho_alt_seq = max(pk_trabalho) 
		from   GRF_Trabalhos 
		where  num_pedido = @num_pedido

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
		select 
		      @fk_trabalho_alt_seq     -- Guarda a Referência do Registro Sequencialmente
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
			 ,@dt_inclusao
			 ,null
		 	 ,dt_exclusao        
        from  GRF_Trabalhos
		where pk_trabalho = @pk_trabalho 

		-- Insere a Fase de Manutenção 1 Antes da Alteração (Histórico)
		select @fk_fase = fk_fase 
		from   GRF_Trabalhos_Fases_Manut1 
		where  pk_trabalho_fase_manut1 = @pk_trabalho_fase_manut1

		select @fk_trabalho_fase_manut1_alt_seq = max(pk_trabalho_fase_manut1) 
		from   GRF_Trabalhos_Fases_Manut1 
		where  fk_trabalho = @pk_trabalho
		and    fk_fase     = @fk_fase

		select @fk_Trabalho_Fase_Status_Tipo_ATUAL = fk_trabalho_fase_status_tipo 
		from   VIW_GRF_Trabalhos_Status 
		where  pk_trabalho_fase_status = dbo.FNC_GRF_Trabalho_Fase_Status_Atual(@pk_trabalho)

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
		select
		      @fk_trabalho_fase_manut1_alt_seq  -- Guarda a Referência do Registro Sequencialmente
		     ,fk_trabalho
		     ,fk_fase
		     ,fk_operador
		     ,cores_alteradas
		     ,circ_cilindro
		     ,observacoes
		     ,@dt_inclusao
		     ,null
        from  GRF_Trabalhos_Fases_Manut1
		where pk_trabalho_fase_manut1 = @pk_trabalho_fase_manut1 

		-- Altera o Trabalho
		update GRF_Trabalhos         
		set    fk_trabalho_tipo     = @fk_trabalho_tipo
		   	  ,fk_produto           = @pk_produto
			  ,fk_representante     = @pk_representante
              ,fk_aprovacao_tipo    = @fk_aprovacao_tipo
			  ,fk_cliente           = @fk_cliente
			  ,fk_filial            = @fk_filial
              ,num_pedido_antigo    = @num_pedido_antigo
              ,num_pedido_novo      = @num_pedido_novo
              ,lead_time_programado = @lead_time_programado
              ,dt_alteracao         = @dt_inclusao
        where  pk_trabalho          = @pk_trabalho			

		-- Altera a Fase de Manutenção 1
		update GRF_Trabalhos_Fases_Manut1         
		set    fk_operador             = @fk_operador
			  ,cores_alteradas         = @cores_alteradas
			  ,circ_cilindro           = @circ_cilindro
			  ,observacoes             = @observacoes
			  ,dt_alteracao            = @dt_inclusao
        where  pk_trabalho_fase_manut1 = @pk_trabalho_fase_manut1

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
				@pk_trabalho
			   ,@fk_fase
			   ,@fk_Trabalho_Fase_Status_Tipo_TRABALHANDO
			   ,null
			   ,'PRC_GRF_Atualiza_Trabalhos_2'
			   ,@dt_inclusao
			)
		end

	COMMIT TRANSACTION

	select  '3'                                  Status
		   ,'Atualização realizada com sucesso!' mensagem          		               
	       ,@pk_produto                          pk_produto                   
	       ,@pk_representante                    pk_representante                        

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

   