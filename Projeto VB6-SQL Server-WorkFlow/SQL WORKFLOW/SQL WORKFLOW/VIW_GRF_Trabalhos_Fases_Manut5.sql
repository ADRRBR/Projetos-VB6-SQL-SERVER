
/*
=========================================================================
Retorna Todos os Registros da Fase de Manutenção 5, Exceto os Históricos
=========================================================================
*/

IF OBJECT_ID ( 'dbo.VIW_GRF_Trabalhos_Fases_Manut5') IS NOT NULL DROP VIEW dbo.VIW_GRF_Trabalhos_Fases_Manut5
GO

CREATE VIEW dbo.VIW_GRF_Trabalhos_Fases_Manut5

WITH ENCRYPTION
  
AS

select FM5.*
from   GRF_Trabalhos_Fases_Manut5 FM5 
where  FM5.fk_trabalho_fase_manut5_alt_seq is null

