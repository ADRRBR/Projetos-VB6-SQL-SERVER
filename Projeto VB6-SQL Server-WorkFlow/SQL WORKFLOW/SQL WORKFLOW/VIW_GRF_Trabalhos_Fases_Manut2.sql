
/*
=========================================================================
Retorna Todos os Registros da Fase de Manutenção 2, Exceto os Históricos
=========================================================================
*/

IF OBJECT_ID ( 'dbo.VIW_GRF_Trabalhos_Fases_Manut2') IS NOT NULL DROP VIEW dbo.VIW_GRF_Trabalhos_Fases_Manut2
GO

CREATE VIEW dbo.VIW_GRF_Trabalhos_Fases_Manut2

WITH ENCRYPTION
  
AS

select FM2.*
from   GRF_Trabalhos_Fases_Manut2 FM2 
where  FM2.fk_trabalho_fase_manut2_alt_seq is null

