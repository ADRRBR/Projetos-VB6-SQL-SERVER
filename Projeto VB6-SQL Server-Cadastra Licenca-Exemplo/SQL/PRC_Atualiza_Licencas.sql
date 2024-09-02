/*
=====================================
Atualizal��o da Tabela Tab_Licencas
=====================================
*/

IF OBJECT_ID ( 'dbo.PRC_Atualiza_Licencas') IS NOT NULL DROP PROCEDURE dbo.PRC_Atualiza_Licencas
GO  

CREATE PROCEDURE dbo.PRC_Atualiza_Licencas
                            (
							 @ID_software            int                  
							,@nome_software          varchar(100)       
							,@tipo_software          varchar(50)    -- SO / OFFICE / UTILITARIO      
							,@serial   	             varchar(1000)          
							,@data_expiracao 	     date 
							,@nome_usuario_ult_manut varchar(100) 
							,@tipo_manut             varchar(3)  -- 'INC' / 'ALT' / 'EXC'
							)

WITH ENCRYPTION
AS

declare @ID_software_novo int 
       ,@data_ult_manut   datetime  
	   ,@mensagem         varchar(2000)

SET NOCOUNT    ON
SET XACT_ABORT ON  

-- Verifica Par�metros Principais
if upper(isnull(@tipo_manut,' ')) not in ('INC','ALT','EXC') begin
	select 'Informe o par�metro < @tipo_manut >  INC / ALT / EXC' mensagem
    return
end
if upper(isnull(@tipo_manut,' ')) in ('ALT','EXC') and isnull(@ID_software,-1) <= 0 begin
	select 'Informe o par�metro < @ID_software >  para ALT / EXC' mensagem
    return
end
if isnull(@nome_software,' ') is null begin
	select 'Informe o par�metro < @nome_software >' mensagem
    return
end
if upper(isnull(@tipo_software,' ')) not in ('SO','OFFICE','UTILITARIO') begin
	select 'Informe o par�metro < @tipo_software > SO / OFFICE / UTILITARIO' mensagem
	return
end
if isnull(@serial,' ') is null begin
	select 'Informe o par�metro < @serial >' mensagem
	return
end
if isnull(@data_expiracao,' ') is null begin
	select 'Informe o par�metro < @data_expiracao >' mensagem
	return
end
if isnull(@nome_usuario_ult_manut,' ') is null begin
	select 'Informe o par�metro < @@nome_usuario_ult_manut >' mensagem
	return
end

set  @ID_software_novo  = null
set  @data_ult_manut    = getdate()

BEGIN TRY 
	BEGIN TRANSACTION
		if @tipo_manut = 'INC' begin
			-- ***** INCLUS�O

			-- Insere Registro
			insert into TAB_Licencas
			(
			 	 nome_software           
				,tipo_software           
				,serial   	              
				,data_expiracao 	      
				-- Auditoria de Manuten��o no Registro
				,data_inc                
				,data_ult_manut          
				,nome_usuario_ult_manut  
			)
			values 							 
			(
				 @nome_software           
				,@tipo_software           
				,@serial   	              
				,@data_expiracao 	      
				-- Auditoria de Manuten��o no Registro
				,@data_ult_manut          
				,@data_ult_manut
				,@nome_usuario_ult_manut  
			)
			
			set @ID_software_novo = @@IDENTITY

			set @mensagem = 'Regitro alterado com sucesso! Novo ID = ' + convert(varchar,@ID_software_novo)
		end
		else if @tipo_manut = 'ALT' begin
			-- ***** ALTERA��O

			update TAB_Licencas         
			set    nome_software            = @nome_software         
				  ,tipo_software            = @tipo_software
				  ,serial   	            = @serial
				  ,data_expiracao 	        = @data_expiracao
				   -- Auditoria de Manuten��o no Registro
				  ,data_ult_manut           = @data_ult_manut  
				  ,nome_usuario_ult_manut   = @nome_usuario_ult_manut
			where  ID_software              = @ID_software

			set @mensagem = 'Registro alterado com sucesso!'
		end 
		else if @tipo_manut = 'EXC' begin
			-- ***** ALTERA��O

			delete  
			from   TAB_Licencas          
			where  ID_software  = @ID_software

			set @mensagem = 'Registro exclu�do com sucesso!'
		end 
	COMMIT TRANSACTION

	select @mensagem mensagem

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
