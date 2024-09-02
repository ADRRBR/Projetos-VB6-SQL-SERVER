
/*
=========================================================================
Retorna Todos os Registros da Fase de Manutenção 4, Exceto os Históricos
=========================================================================
*/

IF OBJECT_ID ( 'dbo.VIW_GRF_Trabalhos_Fases_Manut4') IS NOT NULL DROP VIEW dbo.VIW_GRF_Trabalhos_Fases_Manut4
GO

CREATE VIEW dbo.VIW_GRF_Trabalhos_Fases_Manut4

WITH ENCRYPTION
  
AS

select FM4.*
from   GRF_Trabalhos_Fases_Manut4 FM4 
where  FM4.fk_trabalho_fase_manut4_alt_seq is null

