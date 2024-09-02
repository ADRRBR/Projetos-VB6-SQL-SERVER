
/*
=========================================================================
Retorna Todos os Registros da Fase de Manutenção 3, Exceto os Históricos
=========================================================================
*/

IF OBJECT_ID ( 'dbo.VIW_GRF_Trabalhos_Fases_Manut3') IS NOT NULL DROP VIEW dbo.VIW_GRF_Trabalhos_Fases_Manut3
GO

CREATE VIEW dbo.VIW_GRF_Trabalhos_Fases_Manut3

WITH ENCRYPTION
  
AS

select FM3.*
from   GRF_Trabalhos_Fases_Manut3 FM3 
where  FM3.fk_trabalho_fase_manut3_alt_seq is null

