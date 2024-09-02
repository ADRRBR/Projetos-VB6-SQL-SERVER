
-- Inclusão
EXEC dbo.PRC_Atualiza_Licencas
       NULL            --@ID_software           
      ,'SOFTWARE 1'	   --@nome_software         
	  ,'SO'			   --@tipo_software   SO / OFFICE / UTILITARIO        
	  ,'111111'		   --@serial   	            
      ,'2024-08-09'    --@data_expiracao 	                       
	  ,'ADRIANO'       --@nome_usuario_ult_manut	
	  ,'INC'           --@tipo_manut     INC / ALT / EXC

-- Alteração
EXEC dbo.PRC_Atualiza_Licencas
       2            --@ID_software           
      ,'SOFTWARE 1'	   --@nome_software         
	  ,'SO'			   --@tipo_software   SO / OFFICE / UTILITARIO        
	  ,'22222'		   --@serial   	            
      ,'2024-08-09'    --@data_expiracao 	                       
	  ,'ADRIANO'       --@nome_usuario_ult_manut	
	  ,'ALT'           --@tipo_manut     INC / ALT / EXC


-- Exclusão
EXEC dbo.PRC_Atualiza_Licencas
       2               --@ID_software           
      ,'SOFTWARE 1'	   --@nome_software         
	  ,'SO'			   --@tipo_software   SO / OFFICE / UTILITARIO        
	  ,'22222'		   --@serial   	            
      ,'2024-08-09'    --@data_expiracao 	                       
	  ,'ADRIANO'       --@nome_usuario_ult_manut	
	  ,'EXC'           --@tipo_manut     INC / ALT / EXC

