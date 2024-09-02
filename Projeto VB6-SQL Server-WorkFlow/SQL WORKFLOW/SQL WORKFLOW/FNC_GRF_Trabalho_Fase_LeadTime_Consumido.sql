
/*
======================================================================
Retorna o Lead Time Consumido por uma Determinada Fase de um Trabalho
======================================================================
*/

IF OBJECT_ID ( 'dbo.FNC_GRF_Trabalho_Fase_LeadTime_Consumido') IS NOT NULL DROP FUNCTION dbo.FNC_GRF_Trabalho_Fase_LeadTime_Consumido
GO  

CREATE FUNCTION dbo.FNC_GRF_Trabalho_Fase_LeadTime_Consumido (
                                                              @pk_trabalho INT
                                                             ,@pk_fase     INT
														     ) 
RETURNS VARCHAR(6)

WITH ENCRYPTION
  
AS

BEGIN
	declare @NovoIntervalo                 int
		   ,@DataIntervaloInicio           datetime
		   ,@DataIntervaloFim              datetime
		   ,@TotalHoras                    int
		   ,@TotalMinutos                  int

		   ,@data_inclusao_status	       datetime	
		   ,@interrompe_lead_time          tinyint	

	-- Cursor Para Cálculo do LEAD TIME de uma Determinada Fase de um Trabalho (Diferentes Intervalos de Trabalho)
	declare c_Trab_Fase_Status cursor for
		select   data_inclusao_status	        	
				,interrompe_lead_time 
		from     VIW_GRF_Trabalhos_Status 
		where    pk_trabalho = @pk_trabalho 
		and      pk_fase     = @pk_fase
		order by data_inclusao_status

	open c_Trab_Fase_Status

	set @NovoIntervalo = 1
	set @TotalHoras    = 0
	set @TotalMinutos  = 0

	while 1 = 1	begin
		fetch next 
		from  c_Trab_Fase_Status 
		into  @data_inclusao_status	        	
			 ,@interrompe_lead_time    
		    		
		if @@fetch_status  <> 0 break

		if @interrompe_lead_time = 0 begin
			if @NovoIntervalo = 1 begin
				set @NovoIntervalo = 0
				set @DataIntervaloInicio = @data_inclusao_status
			end 
		end
		else begin
			set @NovoIntervalo = 1
			set @DataIntervaloFim = @data_inclusao_status

			-- Identifica o Intervalo Entre as Datas em Minutos
			select @TotalMinutos = @TotalMinutos + datediff(minute, @DataIntervaloInicio, @DataIntervaloFim) 
		end  
	end

	close c_Trab_Fase_Status
	deallocate c_Trab_Fase_Status

	-- Caso o Último Status da Fase Referente ao Trabalho Não Seja Para Interrupção de Intervalo, Considerar a Data do Sistema como Fechamento do Último Intervalo, ou Seja, a Fase Está em Andamento.
	if @interrompe_lead_time = 0 begin
		set @DataIntervaloFim = getdate()

		-- Identifica o Intervalo Entre as Datas em Minutos
		select @TotalMinutos = @TotalMinutos + datediff(minute, @DataIntervaloInicio, @DataIntervaloFim) 
	end
	
	-- Converte o Total de Minutos de Todos os Intervalos em Horas e Minutos
	while @TotalMinutos >= 60 begin	
	   set @TotalHoras   = @TotalHoras + 1
	   set @TotalMinutos = @TotalMinutos - 60
	end

    RETURN(right('000' + convert(varchar,@TotalHoras),3) + ':' + right('00' + convert(varchar,@TotalMinutos),2))
END




