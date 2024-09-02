
/*
===============================
Retorna o Lead Time em Minutos
===============================
*/

IF OBJECT_ID ( 'dbo.FNC_GRF_Converte_LeadTime_Minutos') IS NOT NULL DROP FUNCTION dbo.FNC_GRF_Converte_LeadTime_Minutos
GO  

CREATE FUNCTION dbo.FNC_GRF_Converte_LeadTime_Minutos (
                                                       @lead_time varchar(6)
						                              ) 
RETURNS INT

WITH ENCRYPTION
               
AS

BEGIN
	declare @PosSeparador  int
   		   ,@Horas         int
		   ,@Minutos       int
		   ,@TotalMinutos  int

	set @PosSeparador = charindex(':',@lead_time) 
	set @Horas = substring(@lead_time,1,@PosSeparador-1) 
	set @Minutos = substring(@lead_time,@PosSeparador+1,len(ltrim(rtrim(@lead_time)))-@PosSeparador)
	set @TotalMinutos = (@Horas * 60) + @Minutos

    RETURN @TotalMinutos 
END
