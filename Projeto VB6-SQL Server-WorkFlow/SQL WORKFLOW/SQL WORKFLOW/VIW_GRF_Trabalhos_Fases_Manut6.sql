
/*
=========================================================================
Retorna Todos os Registros da Fase de Manuten��o 6, Exceto os Hist�ricos
=========================================================================
*/

IF OBJECT_ID ( 'dbo.VIW_GRF_Trabalhos_Fases_Manut6') IS NOT NULL DROP VIEW dbo.VIW_GRF_Trabalhos_Fases_Manut6
GO

CREATE VIEW dbo.VIW_GRF_Trabalhos_Fases_Manut6

WITH ENCRYPTION
  
AS

select FM6.*
from   GRF_Trabalhos_Fases_Manut6 FM6 
where  FM6.fk_trabalho_fase_manut6_alt_seq is null

