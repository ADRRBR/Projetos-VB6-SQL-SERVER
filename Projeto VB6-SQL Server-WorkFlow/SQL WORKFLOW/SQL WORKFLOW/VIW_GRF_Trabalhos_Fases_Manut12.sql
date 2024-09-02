
/*
=========================================================================
Retorna Todos os Registros da Fase de Manutenção 12, Exceto os Históricos
=========================================================================
*/

IF OBJECT_ID ( 'dbo.VIW_GRF_Trabalhos_Fases_Manut12') IS NOT NULL DROP VIEW dbo.VIW_GRF_Trabalhos_Fases_Manut12
GO

CREATE VIEW dbo.VIW_GRF_Trabalhos_Fases_Manut12

WITH ENCRYPTION
  
AS

select FM12.*
from   GRF_Trabalhos_Fases_Manut12 FM12 
where  FM12.fk_trabalho_fase_manut12_alt_seq is null

