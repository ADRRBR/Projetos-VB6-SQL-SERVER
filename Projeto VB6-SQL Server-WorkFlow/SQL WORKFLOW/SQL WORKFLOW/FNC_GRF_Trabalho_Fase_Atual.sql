
/*
==========================================================================
Retorna a Chave Primária da Fase em que se Encontra Atualmente o Trabalho  
==========================================================================
*/

IF OBJECT_ID ( 'dbo.FNC_GRF_Trabalho_Fase_Atual') IS NOT NULL DROP FUNCTION dbo.FNC_GRF_Trabalho_Fase_Atual
GO  

CREATE FUNCTION dbo.FNC_GRF_Trabalho_Fase_Atual (
                                                 @pk_trabalho INT
                                                ) 
RETURNS INT

WITH ENCRYPTION
  
AS

BEGIN
	declare @pk_fase_atual int

	select  @pk_fase_atual = pk_fase 
	from    VIW_GRF_Trabalhos_Status
	where   pk_trabalho_fase_status = dbo.FNC_GRF_Trabalho_Fase_Status_Atual(@pk_trabalho)
		   
    RETURN(isnull(@pk_fase_atual,0))
END


