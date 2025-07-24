USE SAFC_ECB
GO

CREATE TABLE dbo.Exoneracion_Operacion  
         ( Id_Exoneracion char(10)  NOT NULL , Descripcion varchar(300)  NULL ) 
         ALTER TABLE dbo.Exoneracion_Operacion ADD CONSTRAINT XPKExoneracion_Operacion PRIMARY KEY  CLUSTERED (Id_Exoneracion ASC)  

CREATE TABLE dbo.Modalidad_Servicio  
                  ( Id_Modalidad char(10)  NOT NULL , Descripcion varchar(300)  NULL )  
                  ALTER TABLE dbo.Modalidad_Servicio ADD CONSTRAINT XPKModalidad_Servicio PRIMARY KEY  CLUSTERED (Id_Modalidad ASC) 
                  
CREATE TABLE dbo.Pais  
                  ( Id_Pais char(10)  NOT NULL , Nom_Pais varchar(300)  NULL )  
                  ALTER TABLE dbo.Pais ADD CONSTRAINT XPKPais PRIMARY KEY  CLUSTERED (Id_Pais ASC) 
                  
CREATE TABLE dbo.Tipo_Renta  
                  ( Id_Tipo_Renta char(10)  NOT NULL , Descrip_Tipo_Renta varchar(300)  NULL )  
                  ALTER TABLE dbo.Tipo_Renta ADD CONSTRAINT XPKTipo_Renta PRIMARY KEY  CLUSTERED (Id_Tipo_Renta ASC) 
                  
CREATE TABLE dbo.Vinculo_Economico  
                  ( Id_Vinculo_Economico char(10)  NOT NULL , Descrip_Vinculo_Economico varchar(300)  NULL )  
                  ALTER TABLE dbo.Vinculo_Economico ADD CONSTRAINT XPKVinculo_Economico PRIMARY KEY  CLUSTERED (Id_Vinculo_Economico ASC) 

CREATE TABLE dbo.Aduana  
                  ( Id_Aduana char(10)  NOT NULL , Nom_Aduana varchar(300)  NULL)  
                  ALTER TABLE dbo.Aduana ADD CONSTRAINT XPKAduana PRIMARY KEY (Id_Aduana ASC)  

CREATE TABLE dbo.Clasificacion_Servicio  
                  ( Id_Clasific_Servicio char(10)  NOT NULL ,  Nom_Clasific_Servicio varchar(300)  NULL )  
                  ALTER TABLE dbo.Clasificacion_Servicio  
                  ADD CONSTRAINT XPKClasificacion_Servicio PRIMARY KEY (Id_Clasific_Servicio ASC)  

CREATE TABLE dbo.Convenio  
                  ( Id_Convenio char(10)  NOT NULL , Nom_Convenio varchar(300)  NULL  )  
                  ALTER TABLE dbo.Convenio ADD CONSTRAINT XPKConvenio PRIMARY KEY (Id_Convenio ASC) 
                
                
   
                
ALTER TABLE dbo.CNC_ASIENTO_VOUCHER  
         ADD Id_Exoneracion CHAR(10), FOREIGN KEY (Id_Exoneracion) REFERENCES dbo.Exoneracion_Operacion(Id_Exoneracion) 

ALTER TABLE dbo.CNC_ASIENTO_VOUCHER  
                  ADD Id_Tipo_Renta CHAR(10), FOREIGN KEY (Id_Tipo_Renta) REFERENCES dbo.Tipo_Renta(Id_Tipo_Renta) 

ALTER TABLE dbo.CNC_ASIENTO_VOUCHER  
                  ADD Id_Modalidad CHAR(10), FOREIGN KEY (Id_Modalidad) REFERENCES dbo.Modalidad_Servicio(Id_Modalidad) 
                  
ALTER TABLE dbo.CND_ASIENTO_VOUCHER  
         ADD Id_Exoneracion CHAR(10), FOREIGN KEY (Id_Exoneracion) REFERENCES dbo.Exoneracion_Operacion(Id_Exoneracion) 

ALTER TABLE dbo.CND_ASIENTO_VOUCHER  
                  ADD Id_Tipo_Renta CHAR(10), FOREIGN KEY (Id_Tipo_Renta) REFERENCES dbo.Tipo_Renta(Id_Tipo_Renta) 

ALTER TABLE dbo.CND_ASIENTO_VOUCHER  
                  ADD Id_Modalidad CHAR(10), FOREIGN KEY (Id_Modalidad) REFERENCES dbo.Modalidad_Servicio(Id_Modalidad) 

ALTER TABLE dbo.CNM_ENTIDAD  
                  ADD Id_Pais CHAR(10), FOREIGN KEY (Id_Pais) REFERENCES dbo.Pais(Id_Pais) 

ALTER TABLE dbo.CNM_ENTIDAD  
                  ADD Id_Vinculo_Economico CHAR(10), FOREIGN KEY (Id_Vinculo_Economico) REFERENCES dbo.Vinculo_Economico(Id_Vinculo_Economico) 

ALTER TABLE dbo.CND_ASIENTO_VOUCHER  
                  ADD Id_Aduana CHAR(10), FOREIGN KEY (Id_Aduana) REFERENCES Aduana (Id_Aduana)  

ALTER TABLE dbo.CND_ASIENTO_VOUCHER  
                  ADD Id_Clasific_Servicio CHAR(10), FOREIGN KEY(Id_Clasific_Servicio) REFERENCES Clasificacion_Servicio(Id_Clasific_Servicio)  

ALTER TABLE dbo.CNM_ENTIDAD  
                  ADD Id_Convenio CHAR(10), FOREIGN KEY (Id_Convenio) REFERENCES Convenio (Id_Convenio)  