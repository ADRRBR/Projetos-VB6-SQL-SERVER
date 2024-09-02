
/*
=========================================================================
Retorna Todos os Registros da Fase de Manutenção 7, Exceto os Históricos
=========================================================================
*/

IF OBJECT_ID ( 'dbo.VIW_GRF_Trabalhos_Fases_Manut7') IS NOT NULL DROP VIEW dbo.VIW_GRF_Trabalhos_Fases_Manut7
GO

CREATE VIEW dbo.VIW_GRF_Trabalhos_Fases_Manut7

WITH ENCRYPTION
  
AS

select FM7.*
from   GRF_Trabalhos_Fases_Manut7 FM7 
where  FM7.fk_trabalho_fase_manut7_alt_seq is null

