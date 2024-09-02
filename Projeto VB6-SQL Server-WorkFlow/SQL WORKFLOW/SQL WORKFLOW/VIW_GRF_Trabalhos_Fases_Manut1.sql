
/*
=========================================================================
Retorna Todos os Registros da Fase de Manutenção 1, Exceto os Históricos
=========================================================================
*/

IF OBJECT_ID ( 'dbo.VIW_GRF_Trabalhos_Fases_Manut1') IS NOT NULL DROP VIEW dbo.VIW_GRF_Trabalhos_Fases_Manut1
GO

CREATE VIEW dbo.VIW_GRF_Trabalhos_Fases_Manut1

WITH ENCRYPTION
  
AS

select FM1.*
from   GRF_Trabalhos_Fases_Manut1 FM1 
where  FM1.fk_trabalho_fase_manut1_alt_seq is null

