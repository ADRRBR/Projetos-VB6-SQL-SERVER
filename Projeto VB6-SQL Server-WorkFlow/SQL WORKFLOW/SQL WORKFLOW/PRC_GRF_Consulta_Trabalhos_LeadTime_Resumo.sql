
/*
==================================================================
Tipo de Consulta: Trabalho Fases Manutenção - Resumo de Lead Time
==================================================================
*/

IF OBJECT_ID ( 'dbo.PRC_GRF_Consulta_Trabalhos_LeadTime_Resumo') IS NOT NULL DROP PROCEDURE dbo.PRC_GRF_Consulta_Trabalhos_LeadTime_Resumo
GO  

CREATE PROCEDURE dbo.PRC_GRF_Consulta_Trabalhos_LeadTime_Resumo (
							                                     @pk_trabalho  int
							                                    )

WITH ENCRYPTION
AS

declare  @FM1_lead_time_perc_definido     float
		,@FM2_lead_time_perc_definido     float
		,@FM3_lead_time_perc_definido     float
		,@FM4_lead_time_perc_definido     float
		,@FM5_lead_time_perc_definido     float
		,@FM6_lead_time_perc_definido     float
		,@FM7_lead_time_perc_definido     float
		,@FM8_lead_time_perc_definido     float
		,@FM9_lead_time_perc_definido     float
		,@FM10_lead_time_perc_definido    float
		,@FM11_lead_time_perc_definido    float
		,@FM12_lead_time_perc_definido    float
		
		,@TB_lead_time_programado         varchar(6)
		,@TB_lead_time_consumo            varchar(6)
        ,@TB_lead_time_perc_consumido     float

		,@FM1_lead_time_consumo           varchar(6)
        ,@FM2_lead_time_consumo           varchar(6)
        ,@FM3_lead_time_consumo           varchar(6)
        ,@FM4_lead_time_consumo           varchar(6)
        ,@FM5_lead_time_consumo           varchar(6)
        ,@FM6_lead_time_consumo           varchar(6)
        ,@FM7_lead_time_consumo           varchar(6)
        ,@FM8_lead_time_consumo           varchar(6)
        ,@FM9_lead_time_consumo           varchar(6)
        ,@FM10_lead_time_consumo          varchar(6)
        ,@FM11_lead_time_consumo          varchar(6)
        ,@FM12_lead_time_consumo          varchar(6)

		,@FM1_lead_time_perc_consumido    float
		,@FM2_lead_time_perc_consumido    float
		,@FM3_lead_time_perc_consumido    float
		,@FM4_lead_time_perc_consumido    float
		,@FM5_lead_time_perc_consumido    float
		,@FM6_lead_time_perc_consumido    float
		,@FM7_lead_time_perc_consumido    float
		,@FM8_lead_time_perc_consumido    float
		,@FM9_lead_time_perc_consumido    float
		,@FM10_lead_time_perc_consumido   float
		,@FM11_lead_time_perc_consumido   float
		,@FM12_lead_time_perc_consumido   float

		,@TotalMinutosTrabalho_Programado float
		,@TotalHorasTrabalho_Consumido    float
		,@TotalMinutosTrabalho_Consumido  float
		,@TotalMinutosFase_Consumido      float

SET NOCOUNT    ON
SET XACT_ABORT ON  

-- Verifica Parâmetros Principais
if not exists(select 1 from VIW_GRF_Trabalhos where pk_trabalho = @pk_trabalho) begin
	select '6' Status, 'O registro < @pk_trabalho > = ' + convert(varchar,@pk_trabalho) + ', não foi localizado na view < VIW_GRF_Trabalhos >!' mensagem
	return
end

select @FM1_lead_time_perc_definido  = lead_time_perc_definido from GRF_Fases where codigo = 1 
select @FM2_lead_time_perc_definido  = lead_time_perc_definido from GRF_Fases where codigo = 2
select @FM3_lead_time_perc_definido  = lead_time_perc_definido from GRF_Fases where codigo = 3
select @FM4_lead_time_perc_definido  = lead_time_perc_definido from GRF_Fases where codigo = 4
select @FM5_lead_time_perc_definido  = lead_time_perc_definido from GRF_Fases where codigo = 5
select @FM6_lead_time_perc_definido  = lead_time_perc_definido from GRF_Fases where codigo = 6
select @FM7_lead_time_perc_definido  = lead_time_perc_definido from GRF_Fases where codigo = 7
select @FM8_lead_time_perc_definido  = lead_time_perc_definido from GRF_Fases where codigo = 8
select @FM9_lead_time_perc_definido  = lead_time_perc_definido from GRF_Fases where codigo = 9
select @FM10_lead_time_perc_definido = lead_time_perc_definido from GRF_Fases where codigo = 10
select @FM11_lead_time_perc_definido = lead_time_perc_definido from GRF_Fases where codigo = 11
select @FM12_lead_time_perc_definido = lead_time_perc_definido from GRF_Fases where codigo = 12

select @TB_lead_time_programado = TB_lead_time_programado 
from   VIW_GRF_Trabalhos_Fases_Manut 
where  TB_pk_trabalho = @pk_trabalho 

select @FM1_lead_time_consumo  = dbo.FNC_GRF_Trabalho_Fase_LeadTime_Consumido(@pk_trabalho,1)
      ,@FM2_lead_time_consumo  = dbo.FNC_GRF_Trabalho_Fase_LeadTime_Consumido(@pk_trabalho,2)
      ,@FM3_lead_time_consumo  = dbo.FNC_GRF_Trabalho_Fase_LeadTime_Consumido(@pk_trabalho,3)
      ,@FM4_lead_time_consumo  = dbo.FNC_GRF_Trabalho_Fase_LeadTime_Consumido(@pk_trabalho,4)
      ,@FM5_lead_time_consumo  = dbo.FNC_GRF_Trabalho_Fase_LeadTime_Consumido(@pk_trabalho,5)
      ,@FM6_lead_time_consumo  = dbo.FNC_GRF_Trabalho_Fase_LeadTime_Consumido(@pk_trabalho,6)
      ,@FM7_lead_time_consumo  = dbo.FNC_GRF_Trabalho_Fase_LeadTime_Consumido(@pk_trabalho,7)
      ,@FM8_lead_time_consumo  = dbo.FNC_GRF_Trabalho_Fase_LeadTime_Consumido(@pk_trabalho,8)
      ,@FM9_lead_time_consumo  = dbo.FNC_GRF_Trabalho_Fase_LeadTime_Consumido(@pk_trabalho,9)
      ,@FM10_lead_time_consumo = dbo.FNC_GRF_Trabalho_Fase_LeadTime_Consumido(@pk_trabalho,10)
      ,@FM11_lead_time_consumo = dbo.FNC_GRF_Trabalho_Fase_LeadTime_Consumido(@pk_trabalho,11)
      ,@FM12_lead_time_consumo = dbo.FNC_GRF_Trabalho_Fase_LeadTime_Consumido(@pk_trabalho,12)

set @TotalMinutosTrabalho_Consumido = 0

-- *** Trabalho
set @TotalMinutosTrabalho_Programado = dbo.FNC_GRF_Converte_LeadTime_Minutos(@TB_lead_time_programado)

-- *** Fases de Manutenção
set @TotalMinutosFase_Consumido = dbo.FNC_GRF_Converte_LeadTime_Minutos(@FM1_lead_time_consumo)
set @FM1_lead_time_perc_consumido = round((@TotalMinutosFase_Consumido / @TotalMinutosTrabalho_Programado) * 100,1)
set @TotalMinutosTrabalho_Consumido = @TotalMinutosTrabalho_Consumido + @TotalMinutosFase_Consumido

set @TotalMinutosFase_Consumido = dbo.FNC_GRF_Converte_LeadTime_Minutos(@FM2_lead_time_consumo)
set @FM2_lead_time_perc_consumido = round((@TotalMinutosFase_Consumido / @TotalMinutosTrabalho_Programado) * 100,1)
set @TotalMinutosTrabalho_Consumido = @TotalMinutosTrabalho_Consumido + @TotalMinutosFase_Consumido

set @TotalMinutosFase_Consumido = dbo.FNC_GRF_Converte_LeadTime_Minutos(@FM3_lead_time_consumo)
set @FM3_lead_time_perc_consumido = round((@TotalMinutosFase_Consumido / @TotalMinutosTrabalho_Programado) * 100,1)
set @TotalMinutosTrabalho_Consumido = @TotalMinutosTrabalho_Consumido + @TotalMinutosFase_Consumido

set @TotalMinutosFase_Consumido = dbo.FNC_GRF_Converte_LeadTime_Minutos(@FM4_lead_time_consumo)
set @FM4_lead_time_perc_consumido = round((@TotalMinutosFase_Consumido / @TotalMinutosTrabalho_Programado) * 100,1)
set @TotalMinutosTrabalho_Consumido = @TotalMinutosTrabalho_Consumido + @TotalMinutosFase_Consumido

set @TotalMinutosFase_Consumido = dbo.FNC_GRF_Converte_LeadTime_Minutos(@FM5_lead_time_consumo)
set @FM5_lead_time_perc_consumido = round((@TotalMinutosFase_Consumido / @TotalMinutosTrabalho_Programado) * 100,1)
set @TotalMinutosTrabalho_Consumido = @TotalMinutosTrabalho_Consumido + @TotalMinutosFase_Consumido

set @TotalMinutosFase_Consumido = dbo.FNC_GRF_Converte_LeadTime_Minutos(@FM6_lead_time_consumo)
set @FM6_lead_time_perc_consumido = round((@TotalMinutosFase_Consumido / @TotalMinutosTrabalho_Programado) * 100,1)
set @TotalMinutosTrabalho_Consumido = @TotalMinutosTrabalho_Consumido + @TotalMinutosFase_Consumido

set @TotalMinutosFase_Consumido = dbo.FNC_GRF_Converte_LeadTime_Minutos(@FM7_lead_time_consumo)
set @FM7_lead_time_perc_consumido = round((@TotalMinutosFase_Consumido / @TotalMinutosTrabalho_Programado) * 100,1)
set @TotalMinutosTrabalho_Consumido = @TotalMinutosTrabalho_Consumido + @TotalMinutosFase_Consumido

set @TotalMinutosFase_Consumido = dbo.FNC_GRF_Converte_LeadTime_Minutos(@FM8_lead_time_consumo)
set @FM8_lead_time_perc_consumido = round((@TotalMinutosFase_Consumido / @TotalMinutosTrabalho_Programado) * 100,2)
set @TotalMinutosTrabalho_Consumido = @TotalMinutosTrabalho_Consumido + @TotalMinutosFase_Consumido

set @TotalMinutosFase_Consumido = dbo.FNC_GRF_Converte_LeadTime_Minutos(@FM9_lead_time_consumo)
set @FM9_lead_time_perc_consumido = round((@TotalMinutosFase_Consumido / @TotalMinutosTrabalho_Programado) * 100,1)
set @TotalMinutosTrabalho_Consumido = @TotalMinutosTrabalho_Consumido + @TotalMinutosFase_Consumido

set @TotalMinutosFase_Consumido = dbo.FNC_GRF_Converte_LeadTime_Minutos(@FM10_lead_time_consumo)
set @FM10_lead_time_perc_consumido = round((@TotalMinutosFase_Consumido / @TotalMinutosTrabalho_Programado) * 100,1)
set @TotalMinutosTrabalho_Consumido = @TotalMinutosTrabalho_Consumido + @TotalMinutosFase_Consumido

set @TotalMinutosFase_Consumido = dbo.FNC_GRF_Converte_LeadTime_Minutos(@FM11_lead_time_consumo)
set @FM11_lead_time_perc_consumido = round((@TotalMinutosFase_Consumido / @TotalMinutosTrabalho_Programado) * 100,1)
set @TotalMinutosTrabalho_Consumido = @TotalMinutosTrabalho_Consumido + @TotalMinutosFase_Consumido

set @TotalMinutosFase_Consumido = dbo.FNC_GRF_Converte_LeadTime_Minutos(@FM12_lead_time_consumo)
set @FM12_lead_time_perc_consumido = round((@TotalMinutosFase_Consumido / @TotalMinutosTrabalho_Programado) * 100,1)
set @TotalMinutosTrabalho_Consumido = @TotalMinutosTrabalho_Consumido + @TotalMinutosFase_Consumido

-- *** Consumo do Trabalho
set @TB_lead_time_perc_consumido = round((@TotalMinutosTrabalho_Consumido / @TotalMinutosTrabalho_Programado) * 100,1)

-- Converte o Total de Minutos de Todos os Intervalos em Horas e Minutos
set @TotalHorasTrabalho_Consumido = 0
while @TotalMinutosTrabalho_Consumido >= 60 begin	
	set @TotalHorasTrabalho_Consumido   = @TotalHorasTrabalho_Consumido + 1
	set @TotalMinutosTrabalho_Consumido = @TotalMinutosTrabalho_Consumido - 60
end
set @TB_lead_time_consumo = right('000' + convert(varchar,@TotalHorasTrabalho_Consumido),3) + ':' + right('00' + convert(varchar,@TotalMinutosTrabalho_Consumido),2)

-- Retorno
select  @TB_lead_time_programado       TB_lead_time_programado 
	   ,@TB_lead_time_consumo          TB_lead_time_consumo
	   ,@TB_lead_time_perc_consumido   TB_lead_time_perc_consumido
	    --------------------------------------------------------------  
	   ,@FM1_lead_time_consumo         FM1_lead_time_consumo         
       ,@FM1_lead_time_perc_consumido  FM1_lead_time_perc_consumido
       ,@FM1_lead_time_perc_definido   FM1_lead_time_perc_definido
	    --------------------------------------------------------------  
	   ,@FM2_lead_time_consumo         FM2_lead_time_consumo         
       ,@FM2_lead_time_perc_consumido  FM2_lead_time_perc_consumido
       ,@FM2_lead_time_perc_definido   FM2_lead_time_perc_definido
	    --------------------------------------------------------------  
	   ,@FM3_lead_time_consumo         FM3_lead_time_consumo         
       ,@FM3_lead_time_perc_consumido  FM3_lead_time_perc_consumido
       ,@FM3_lead_time_perc_definido   FM3_lead_time_perc_definido
	    --------------------------------------------------------------  
	   ,@FM4_lead_time_consumo         FM4_lead_time_consumo         
       ,@FM4_lead_time_perc_consumido  FM4_lead_time_perc_consumido
       ,@FM4_lead_time_perc_definido   FM4_lead_time_perc_definido
	    --------------------------------------------------------------  
	   ,@FM5_lead_time_consumo         FM5_lead_time_consumo         
       ,@FM5_lead_time_perc_consumido  FM5_lead_time_perc_consumido
       ,@FM5_lead_time_perc_definido   FM5_lead_time_perc_definido
	    --------------------------------------------------------------  
	   ,@FM6_lead_time_consumo         FM6_lead_time_consumo         
       ,@FM6_lead_time_perc_consumido  FM6_lead_time_perc_consumido
       ,@FM6_lead_time_perc_definido   FM6_lead_time_perc_definido
	    --------------------------------------------------------------  
	   ,@FM7_lead_time_consumo         FM7_lead_time_consumo         
       ,@FM7_lead_time_perc_consumido  FM7_lead_time_perc_consumido
       ,@FM7_lead_time_perc_definido   FM7_lead_time_perc_definido
	    --------------------------------------------------------------  
	   ,@FM8_lead_time_consumo         FM8_lead_time_consumo         
       ,@FM8_lead_time_perc_consumido  FM8_lead_time_perc_consumido
       ,@FM8_lead_time_perc_definido   FM8_lead_time_perc_definido
	    --------------------------------------------------------------  
	   ,@FM9_lead_time_consumo         FM9_lead_time_consumo         
       ,@FM9_lead_time_perc_consumido  FM9_lead_time_perc_consumido
       ,@FM9_lead_time_perc_definido   FM9_lead_time_perc_definido
	    --------------------------------------------------------------  
	   ,@FM10_lead_time_consumo        FM10_lead_time_consumo         
       ,@FM10_lead_time_perc_consumido FM10_lead_time_perc_consumido
       ,@FM10_lead_time_perc_definido  FM10_lead_time_perc_definido
	    --------------------------------------------------------------  
	   ,@FM11_lead_time_consumo        FM11_lead_time_consumo         
       ,@FM11_lead_time_perc_consumido FM11_lead_time_perc_consumido
       ,@FM11_lead_time_perc_definido  FM11_lead_time_perc_definido
	    --------------------------------------------------------------  
	   ,@FM12_lead_time_consumo        FM12_lead_time_consumo         
       ,@FM12_lead_time_perc_consumido FM12_lead_time_perc_consumido
       ,@FM12_lead_time_perc_definido  FM12_lead_time_perc_definido
