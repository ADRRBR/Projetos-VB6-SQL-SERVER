
/*
==========================================================================
Retorna Todos os Registros da Fase de Manuten��o 11, Exceto os Hist�ricos
==========================================================================
*/

IF OBJECT_ID ( 'dbo.VIW_GRF_Trabalhos_Fases_Manut11') IS NOT NULL DROP VIEW dbo.VIW_GRF_Trabalhos_Fases_Manut11
GO

CREATE VIEW dbo.VIW_GRF_Trabalhos_Fases_Manut11

WITH ENCRYPTION
  
AS

select FM11.*
from   GRF_Trabalhos_Fases_Manut11 FM11 
where  FM11.fk_trabalho_fase_manut11_alt_seq is null

