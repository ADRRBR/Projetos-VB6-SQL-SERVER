
/*
==========================================================================
Retorna Todos os Registros da Fase de Manutenção 10, Exceto os Históricos
==========================================================================
*/

IF OBJECT_ID ( 'dbo.VIW_GRF_Trabalhos_Fases_Manut10') IS NOT NULL DROP VIEW dbo.VIW_GRF_Trabalhos_Fases_Manut10
GO

CREATE VIEW dbo.VIW_GRF_Trabalhos_Fases_Manut10

WITH ENCRYPTION
  
AS

select FM10.*
from   GRF_Trabalhos_Fases_Manut10 FM10 
where  FM10.fk_trabalho_fase_manut10_alt_seq is null

