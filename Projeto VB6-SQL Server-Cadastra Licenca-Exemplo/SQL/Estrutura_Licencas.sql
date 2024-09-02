
-- ******************** Script para Criação da Tabela
IF OBJECT_ID ('dbo.TAB_Licencas') IS NOT NULL 
    RETURN

CREATE TABLE dbo.TAB_Licencas
(
  ID_software             int          identity(1,1) NOT NULL
 ,nome_software           varchar(100)               NOT NULL
 ,tipo_software           varchar(50)                NOT NULL
 ,serial   	              varchar(1000)              NOT NULL
 ,data_expiracao 	      date                       NOT NULL
 -- Auditoria de Manutenção no Registro
 ,data_inc                date                       NOT NULL
 ,data_ult_manut          date                       NOT NULL
 ,nome_usuario_ult_manut  varchar(100)	             NOT NULL

)

alter table dbo.TAB_Licencas add constraint PK_TAB_Licencas primary key(ID_software);

-- SO / OFFICE / UTILITARIO 
alter table dbo.TAB_Licencas add constraint TAB_Licencas_CHK_1 CHECK (tipo_software IN ('SO','OFFICE','UTILITARIO')); 

create unique index IND_ID_software_1 ON dbo.TAB_Licencas(nome_software)



