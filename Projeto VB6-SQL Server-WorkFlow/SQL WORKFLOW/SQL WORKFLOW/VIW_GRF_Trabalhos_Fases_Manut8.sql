
/*
=========================================================================
Retorna Todos os Registros da Fase de Manutenção 8, Exceto os Históricos
=========================================================================
*/

IF OBJECT_ID ( 'dbo.VIW_GRF_Trabalhos_Fases_Manut8') IS NOT NULL DROP VIEW dbo.VIW_GRF_Trabalhos_Fases_Manut8
GO

CREATE VIEW dbo.VIW_GRF_Trabalhos_Fases_Manut8

WITH ENCRYPTION
  
AS

select FM8.*
from   GRF_Trabalhos_Fases_Manut8 FM8 
where  FM8.fk_trabalho_fase_manut8_alt_seq is null

