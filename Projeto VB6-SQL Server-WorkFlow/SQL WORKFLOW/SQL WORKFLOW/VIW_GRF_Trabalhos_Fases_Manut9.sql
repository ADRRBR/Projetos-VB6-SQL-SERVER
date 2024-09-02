
/*
=========================================================================
Retorna Todos os Registros da Fase de Manuten��o 9, Exceto os Hist�ricos
=========================================================================
*/

IF OBJECT_ID ( 'dbo.VIW_GRF_Trabalhos_Fases_Manut9') IS NOT NULL DROP VIEW dbo.VIW_GRF_Trabalhos_Fases_Manut9
GO

CREATE VIEW dbo.VIW_GRF_Trabalhos_Fases_Manut9

WITH ENCRYPTION
  
AS

select FM9.*
from   GRF_Trabalhos_Fases_Manut9 FM9 
where  FM9.fk_trabalho_fase_manut9_alt_seq is null

