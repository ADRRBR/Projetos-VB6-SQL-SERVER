
/*
=====================================================================
Retorna Todos os Registros de Campos Genéricos, Exceto os Históricos
=====================================================================
*/

IF OBJECT_ID ( 'dbo.VIW_GRF_Trabalhos_Fases_Manut_Generico') IS NOT NULL DROP VIEW dbo.VIW_GRF_Trabalhos_Fases_Manut_Generico
GO

CREATE VIEW dbo.VIW_GRF_Trabalhos_Fases_Manut_Generico

WITH ENCRYPTION
  
AS

select GEN.*
from   GRF_Trabalhos_Fases_Manut_Generico GEN
where  GEN.fk_trabalho_fase_manut_gen_alt_seq is null

