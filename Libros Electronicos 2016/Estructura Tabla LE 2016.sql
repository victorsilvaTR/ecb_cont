SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO
/*-----------------------------------------------------------------------------------------------------------------            
MODULO DE CONTABILIDAD            
DESCRIPCION  : Crea tabla temporales a utilizarse en diferentes reportes            
------------------------------------------------------------------------------------------------------------------*/            
CREATE PROCEDURE [dbo].[spCn_CrearTablaTemporal1](            
@Tabla char(20)            
)             
--WITH ENCRYPTION            
AS              
SET DATEFORMAT DMY            
            
IF @Tabla = 'COMPRAS'            
BEGIN            
CREATE TABLE [TMPREGISTROCOMPRAS] (            
 [Emp_cCodigo] [char] (3) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Emp_cCodigo] DEFAULT (''),            
 [Ase_cNumMov] [char] (10) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Ase_cNumMov] DEFAULT (''),            
 [Per_cPeriodo] [char] (2) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Per_cPeriodo] DEFAULT (''),            
 [Ase_nVoucher] [char] (10) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Ase_nVoucher] DEFAULT (''),            
 [Asd_nItem] [int] NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Asd_nItem] DEFAULT (0),            
 [Asd_nTipoCambio] [numeric](14, 3)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Asd_nTipoCambio] DEFAULT (0),            
 [Asd_cGlosa] [varchar] (250) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Asd_cGlosa] DEFAULT (''),            
 [Ent_cCodEntidad] [char] (5) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Ent_cCodEntidad] DEFAULT (''),            
 [Ent_nRuc] [varchar] (15) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Ent_nRuc] DEFAULT (''),            
 [Ent_cPersona] [varchar] (120) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Ent_cPersona] DEFAULT (''),            
 [Asd_cTipoDoc] [char] (3) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Asd_cTipoDoc] DEFAULT (''),            
 [Asd_cSerieDoc] [varchar] (20) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Asd_cSerieDoc] DEFAULT (''),            
 [Asd_cNumDoc] [varchar] (25) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Asd_cNumDoc] DEFAULT (''),            
 [Asd_dFecDoc] [datetime] NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Asd_dFecDoc] DEFAULT (''),            
 [Asd_cTipoDocRef] [char] (3) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Asd_cTipoDocRef] DEFAULT (''),            
 [Asd_cSerieDocRef] [char] (5) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Asd_cSerieDocRef] DEFAULT (''),            
 [Asd_cNumDocRef] [char] (10) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Asd_cNumDocRef] DEFAULT (''),            
 [Asd_dFecDocRef] [datetime] NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Asd_dFecDocRef] DEFAULT (''),            
 [Asd_nMontoInafecto] [numeric](14, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Asd_nMontoInafecto] DEFAULT (0),            
 [Asd_cRetencion] [char] (1) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Asd_cRetencion] DEFAULT (''),            
 [Tdo_cNombreLargo] [varchar] (60) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Tdo_cNombreLargo] DEFAULT (''),            
 [Ase_cTipoMoneda] [char] (3) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Ase_cTipoMoneda] DEFAULT (''),            
 [CtaBaseA] [varchar] (12) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_CtaBaseA] DEFAULT (''),            
 [NombreCtaBaseA] [varchar] (120) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_NombreCtaBaseA] DEFAULT (''),            
 [CtaBaseB] [varchar] (12) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_CtaBaseB] DEFAULT (''),            
 [NombreCtaBaseB] [varchar] (120) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_NombreCtaBaseB] DEFAULT (''),            
 [CtaBaseC] [varchar] (12) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_CtaBaseC] DEFAULT (''),            
 [NombreCtaBaseC] [varchar] (120) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_NombreCtaBaseC] DEFAULT (''),            
 [CtaIgvA] [varchar] (12) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_CtaIgvA] DEFAULT (''),            
 [NombreIgvA] [varchar] (120) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_NombreIgvA] DEFAULT (''),            
 [CtaIgvB] [varchar] (12) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_CtaIgvB] DEFAULT (''),            
 [NombreIgvB] [varchar] (120) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_NombreIgvB] DEFAULT (''),            
 [CtaIgvC] [varchar] (12) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_CtaIgvC] DEFAULT (''),            
 [NombreIgvC] [varchar] (120) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_NombreIgvC] DEFAULT (''), [CtaProv] [varchar] (12) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_CtaProv] DEFAULT (''),            
 [NombreProv] [varchar] (120) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_NombreProv] DEFAULT (''),            
 [MontoSBaseA] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_MontoSBaseA] DEFAULT (0),            
 [MontoDBaseA] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_MontoDBaseA] DEFAULT (0),            
 [MontoSBaseB] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_MontoSBaseB] DEFAULT (0),            
 [MontoDBaseB] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_MontoDBaseB] DEFAULT (0),            
 [MontoSBaseC] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_MontoSBaseC] DEFAULT (0),            
 [MontoDBaseC] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_MontoDBaseC] DEFAULT (0),            
 [MontoSIgvA] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_MontoSIgvA] DEFAULT (0),            
 [MontoDIgvA] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_MontoDIgvA] DEFAULT (0),            
 [MontoSIgvB] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_MontoSIgvB] DEFAULT (0),            
 [MontoDIgvB] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_MontoDIgvB] DEFAULT (0),            
 [MontoSIgvC] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_MontoSIgvC] DEFAULT (0),            
 [MontoDIgvC] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_MontoDIgvC] DEFAULT (0),            
 [MontoSOtros] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_MontoSOtros] DEFAULT (0),            
 [MontoDOtros] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_MontoDOtros] DEFAULT (0),            
 [MontoSProv] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_MontoSProv] DEFAULT (0),            
 [MontoDProv] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_MontoDProv] DEFAULT (0),             
            
 [MontoSISC] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_MontoSISC] DEFAULT (0),            
 [MontoDISC] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_MontoDISC] DEFAULT (0),             
            
 [MontoSDIFC] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_MontoSDIFC] DEFAULT (0),            
 [MontoDDIFC] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_MontoDDIFC] DEFAULT (0),             
            
            
 [ASD_DFECHASPOT] [datetime] NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_ASD_DFECHASPOT] DEFAULT (''),            
 [ASD_CNUMSPOT] [char] (25) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_ASD_CNUMSPOT] DEFAULT (''),            
 [imp_nPorcentaje] [numeric](14, 3)  NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_imp_nPorcentaje] DEFAULT (0),             
 [Asd_dFecVen] [datetime] NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Asd_dFecVen] DEFAULT ('')            
) ON [PRIMARY]            
END            
            
IF @Tabla = 'VENTAS'            
BEGIN            
CREATE TABLE [TMPREGISTROVENTAS] (            
 [Emp_cCodigo] [char] (3) null  CONSTRAINT [DF_TMPREGISTRVENTAS_Emp_cCodigo] DEFAULT (''),            
 [Ase_cNumMov] [char] (10) null  CONSTRAINT [DF_TMPREGISTRVENTAS_Ase_cNumMov] DEFAULT (''),            
 [Per_cPeriodo] [char] (2) null  CONSTRAINT [DF_TMPREGISTRVENTAS_Per_cPeriodo] DEFAULT (''),            
 [Ase_nVoucher] [char] (10) null  CONSTRAINT [DF_TMPREGISTRVENTAS_Ase_nVoucher] DEFAULT (''),            
 [Asd_nItem] [int] null  CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_nItem] DEFAULT (0),            
 [Asd_nTipoCambio] [numeric](14, 3) null  CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_nTipoCambio] DEFAULT (0),            
 [Asd_cGlosa] [varchar] (250) null  CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_cGlosa] DEFAULT (''),            
 [Ent_cCodEntidad] [char] (5) null  CONSTRAINT [DF_TMPREGISTRVENTAS_Ent_cCodEntidad] DEFAULT (''),            
 [Ent_nRuc] [varchar] (15) null  CONSTRAINT [DF_TMPREGISTRVENTAS_Ent_nRuc] DEFAULT (''),            
 [Ent_cPersona] [varchar] (120) null  CONSTRAINT [DF_TMPREGISTRVENTAS_Ent_cPersona] DEFAULT (''),            
 [Asd_cTipoDoc] [char] (3) null  CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_cTipoDoc] DEFAULT (''),            
 [Asd_cSerieDoc] [varchar] (20) null  CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_cSerieDoc] DEFAULT (''),            
 [Asd_cNumDoc] [varchar] (25) null  CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_cNumDoc] DEFAULT (''),            
 [Asd_dFecDoc] [datetime] NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_dFecDoc] DEFAULT (''),            
 [Asd_cTipoDocRef] [char] (3) null  CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_cTipoDocRef] DEFAULT (''),            
 [Asd_cSerieDocRef] [char] (5) null  CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_cSerieDocRef] DEFAULT (''),            
 [Asd_cNumDocRef] [char] (10) null  CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_cNumDocRef] DEFAULT (''),            
 [Asd_dFecDocRef] [datetime] NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_dFecDocRef] DEFAULT (''),            
 [Asd_nMontoInafecto] [numeric](14, 3) null  CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_nMontoInafecto] DEFAULT (0),            
 [Asd_cRetencion] [char] (1) null  CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_cRetencion] DEFAULT (''),            
 [Tdo_cNombreLargo] [varchar] (60) null  CONSTRAINT [DF_TMPREGISTRVENTAS_Tdo_cNombreLargo] DEFAULT (''),            
 [Ase_cTipoMoneda] [char] (3) null  CONSTRAINT [DF_TMPREGISTRVENTAS_Ase_cTipoMoneda] DEFAULT (''),            
 [CtaBaseGra] [varchar] (12) null  CONSTRAINT [DF_TMPREGISTRVENTAS_CtaBaseGra] DEFAULT (''),            
 [NombreCtaBaseGra] [varchar] (120) null  CONSTRAINT [DF_TMPREGISTRVENTAS_NombreCtaBaseGra] DEFAULT (''),            
 [CtaBaseExo] [varchar] (12) null  CONSTRAINT [DF_TMPREGISTRVENTAS_CtaBaseExo] DEFAULT (''),            
 [NombreCtaBaseExo] [varchar] (120) null  CONSTRAINT [DF_TMPREGISTRVENTAS_NombreCtaBaseExo] DEFAULT (''),            
 [CtaIgv] [varchar] (12) null  CONSTRAINT [DF_TMPREGISTRVENTAS_CtaIgv] DEFAULT (''),            
 [NombreIgv] [varchar] (120) null  CONSTRAINT [DF_TMPREGISTRVENTAS_NombreIgv] DEFAULT (''),            
 [CtaCli] [varchar] (12) null  CONSTRAINT [DF_TMPREGISTRVENTAS_CtaCli] DEFAULT (''),            
 [NombreCli] [varchar] (120) null  CONSTRAINT [DF_TMPREGISTRVENTAS_NombreCli] DEFAULT (''),            
 [MontoSBaseGra] [numeric](15, 3) null  CONSTRAINT [DF_TMPREGISTRVENTAS_MontoSBaseGra] DEFAULT (0),            
 [MontoDBaseGra] [numeric](15, 3) null  CONSTRAINT [DF_TMPREGISTRVENTAS_MontoDBaseGra] DEFAULT (0),            
 [MontoSBaseExo] [numeric](15, 3) null  CONSTRAINT [DF_TMPREGISTRVENTAS_MontoSBaseExo] DEFAULT (0),            
 [MontoDBaseExo] [numeric](15, 3) null  CONSTRAINT [DF_TMPREGISTRVENTAS_MontoDBaseExo] DEFAULT (0),            
 [MontoSBaseIev] [numeric](15, 3) null  CONSTRAINT [DF_TMPREGISTRVENTAS_MontoSBaseIev] DEFAULT (0),            
 [MontoDBaseIev] [numeric](15, 3) null  CONSTRAINT [DF_TMPREGISTRVENTAS_MontoDBaseIev] DEFAULT (0),            
 [MontoSIsc] [numeric](15, 3) null  CONSTRAINT [DF_TMPREGISTRVENTAS_MontoSIsc] DEFAULT (0),            
 [MontoDIsc] [numeric](15, 3) null  CONSTRAINT [DF_TMPREGISTRVENTAS_MontoDIsc] DEFAULT (0),            
 [MontoSIgv] [numeric](15, 3) null  CONSTRAINT [DF_TMPREGISTRVENTAS_MontoSIgv] DEFAULT (0),            
 [MontoDIgv] [numeric](15, 3) null  CONSTRAINT [DF_TMPREGISTRVENTAS_MontoDIgv] DEFAULT (0),            
 [MontoSIgvIev] [numeric](15, 3) null  CONSTRAINT [DF_TMPREGISTRVENTAS_MontoSIgvIev] DEFAULT (0),            
 [MontoDIgvIev] [numeric](15, 3) null  CONSTRAINT [DF_TMPREGISTRVENTAS_MontoDIgvIev] DEFAULT (0),            
 [MontoSFob] [numeric](15, 3) null  CONSTRAINT [DF_TMPREGISTRVENTAS_MontoSFob] DEFAULT (0),       
 [MontoDFob] [numeric](15, 3) null  CONSTRAINT [DF_TMPREGISTRVENTAS_MontoDFob] DEFAULT (0),            
 [MontoSFlete] [numeric](15, 3) null  CONSTRAINT [DF_TMPREGISTRVENTAS_MontoSFlete] DEFAULT (0),            
 [MontoDFlete] [numeric](15, 3) null  CONSTRAINT [DF_TMPREGISTRVENTAS_MontoDFlete] DEFAULT (0),            
 [MontoSRedondeo] [numeric](15, 3) null  CONSTRAINT [DF_TMPREGISTRVENTAS_MontoSRedondeo] DEFAULT (0),            
 [MontoDRedondeo] [numeric](15, 3) null  CONSTRAINT [DF_TMPREGISTRVENTAS_MontoDRedondeo] DEFAULT (0),            
 [MontoSCli] [numeric](15, 3) null  CONSTRAINT [DF_TMPREGISTRVENTAS_MontoSCli] DEFAULT (0),            
 [MontoDCli] [numeric](15, 3) null  CONSTRAINT [DF_TMPREGISTRVENTAS_MontoDCli] DEFAULT (0)            
) ON [PRIMARY]            
            
END            
            
         
IF @Tabla = 'REGVENTAS'            
BEGIN            
CREATE TABLE [TMPREGISTROVENTAS] (            
 [Emp_cCodigo] [char] (3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Emp_cCodigo] DEFAULT (''),            
 [Ase_cNumMov] [char] (10) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Ase_cNumMov] DEFAULT (''),            
 [Per_cPeriodo] [char] (2) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Per_cPeriodo] DEFAULT (''),            
 [Asd_dFecDoc] [datetime] NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_dFecDoc] DEFAULT (''),            
 [Asd_cTipoDoc] [char] (3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_cTipoDoc] DEFAULT (''),            
 [Asd_cSerieDoc] [varchar] (20) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_cSerieDoc] DEFAULT (''),            
 [Asd_cNumDoc] [varchar] (25) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_cNumDoc] DEFAULT (''),            
 [Asd_cTipoDocRef] [char] (3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_cTipoDocRef] DEFAULT (''),            
 [Asd_cSerieDocRef] [char] (5) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_cSerieDocRef] DEFAULT (''),            
 [Asd_cNumDocRef] [char] (10) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_cNumDocRef] DEFAULT (''),            
 [Ase_nVoucher] [char] (10) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Ase_nVoucher] DEFAULT (''),            
 [Ent_nRuc] [varchar] (15) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Ent_nRuc] DEFAULT (''),            
 [Ent_cPersona] [varchar] (120) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Ent_cPersona] DEFAULT (''),            
 [Asd_nItem] [int] NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_nItem] DEFAULT (0),            
 [MontoBaseGra] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_MontoBaseGra] DEFAULT (0),            
 [MontoInafecto] [numeric](14, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_MontoInafecto] DEFAULT (0),            
 [MontoBaseOtImp] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_MontoBaseOtImp] DEFAULT (0),            
 [MontoISC] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_MontoISC] DEFAULT (0),            
 [MontoIGV] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_MontoIGV] DEFAULT (0),            
 [MontoFOB] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_MontoFOB] DEFAULT (0),            
 [MontoFlete] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_MontoFlete] DEFAULT (0),            
 [MontoOtros] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_MontoOtros] DEFAULT (0),            
 [MontoDifCmb] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_MontoDifCmb] DEFAULT (0),            
 [MontoTotal] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_MontoTotal] DEFAULT (0),            
        [cNombreDoc] [varchar] (20)  NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_cNombreDoc] DEFAULT (''),            
            
 [MontoExp] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_MontoSExp] DEFAULT (0)            
            
            
            
) ON [PRIMARY]            
            
END            
            
IF @Tabla = 'REGHONORARIOS'            
BEGIN            
      CREATE TABLE [TMPREGISTROHONORARIOS] (            
 [Emp_cCodigo] [char] (3) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_Emp_cCodigo] DEFAULT (''),            
 [Ase_cNumMov] [char] (10) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_Ase_cNumMov] DEFAULT (''),            
 [Per_cPeriodo] [char] (2) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_Per_cPeriodo] DEFAULT (''),            
 [Asd_dFecDoc] [datetime] NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_Asd_dFecDoc] DEFAULT (''),            
 [Asd_cTipoDoc] [char] (3) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_Asd_cTipoDoc] DEFAULT (''),            
 [Asd_cSerieDoc] [varchar] (20) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_Asd_cSerieDoc] DEFAULT (''),            
 [Asd_cNumDoc] [varchar] (25) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_Asd_cNumDoc] DEFAULT (''),            
 [Ase_nVoucher] [char] (10) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_Ase_nVoucher] DEFAULT (''),            
 [Ent_nRuc] [varchar] (15) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_Ent_nRuc] DEFAULT (''),            
 [Ent_cPersona] [varchar] (120) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_Ent_cPersona] DEFAULT (''),            
 [MontoSBase] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_MontoSBase] DEFAULT (0),            
 [MontoSCuarta] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_MontoSCuarta] DEFAULT (0),            
 [MontoSProv] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_MontoSProv] DEFAULT (0),            
 [Asd_nTipoCambio] [numeric](14, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_Asd_nTipoCambio] DEFAULT (0),            
 [MontoDProv] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_MontoDProv] DEFAULT (0),            
 [Asd_cGlosa] [varchar] (250) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_Asd_cGlosa] DEFAULT ('')            
      ) ON [PRIMARY]            
END            
            
IF @Tabla = 'HONORARIOS'            
BEGIN            
CREATE TABLE [TMPREGISTROHONORARIOS] (            
 [Emp_cCodigo] [char] (3) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_Emp_cCodigo] DEFAULT (''),            
 [Ase_cNumMov] [char] (10) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_Ase_cNumMov] DEFAULT (''),            
 [Per_cPeriodo] [char] (2) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_Per_cPeriodo] DEFAULT (''),            
 [Ase_nVoucher] [char] (10) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_Ase_nVoucher] DEFAULT (''),            
 [Asd_nItem] [int] NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_Asd_nItem] DEFAULT (0),            
 [Asd_nTipoCambio] [numeric](14, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_Asd_nTipoCambio] DEFAULT (0),            
 [Asd_cGlosa] [varchar] (250) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_Asd_cGlosa] DEFAULT (''),            
 [Ent_cCodEntidad] [char] (5) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_Ent_cCodEntidad] DEFAULT (''),            
 [Ent_nRuc] [varchar] (15) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_Ent_nRuc] DEFAULT (''),            
 [Ent_cPersona] [varchar] (120) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_Ent_cPersona] DEFAULT (''),            
 [Asd_cTipoDoc] [char] (3) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_Asd_cTipoDoc] DEFAULT (''),            
 [Asd_cSerieDoc] [varchar] (20) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_Asd_cSerieDoc] DEFAULT (''),            
 [Asd_cNumDoc] [varchar] (25) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_Asd_cNumDoc] DEFAULT (''),            
 [Asd_dFecDoc] [datetime] NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_Asd_dFecDoc] DEFAULT (''),            
 [Asd_cTipoDocRef] [char] (3) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_Asd_cTipoDocRef] DEFAULT (''),            
 [Asd_cSerieDocRef] [char] (5) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_Asd_cSerieDocRef] DEFAULT (''),            
 [Asd_cNumDocRef] [char] (10) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_Asd_cNumDocRef] DEFAULT (''),            
 [Asd_dFecDocRef] [datetime] NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_Asd_dFecDocRef] DEFAULT (''),            
 [Asd_nMontoInafecto] [numeric](14, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_Asd_nMontoInafecto] DEFAULT (0),            
 [Asd_cRetencion] [char] (1) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_Asd_cRetencion] DEFAULT (''),            
 [Tdo_cNombreLargo] [varchar] (60) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_Tdo_cNombreLargo] DEFAULT (''),            
 [Ase_cTipoMoneda] [char] (3) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_Ase_cTipoMoneda] DEFAULT (''),            
 [CtaBase] [varchar] (12) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_CtaBase] DEFAULT (''),            
 [NombreCtaBase] [varchar] (120) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_NombreCtaBase] DEFAULT (''),            
 [CtaCuarta] [varchar] (12) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_CtaCuarta] DEFAULT (''),            
 [NombreCtaCuarta] [varchar] (120) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_NombreCtaCuarta] DEFAULT (''),            
 [CtaIes] [varchar] (12) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_CtaIes] DEFAULT (''),            
 [NombreIes] [varchar] (120) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_NombreIes] DEFAULT (''),            
 [CtaProv] [varchar] (12) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_CtaProv] DEFAULT (''),            
 [NombreProv] [varchar] (120) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_NombreProv] DEFAULT (''),            
 [MontoSBase] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_MontoSBase] DEFAULT (0),            
 [MontoDBase] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_MontoDBase] DEFAULT (0),            
 [MontoSCuarta] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_MontoSCuarta] DEFAULT (0),            
 [MontoDCuarta] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_MontoDCuarta] DEFAULT (0),            
 [MontoSIes] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_MontoSIes] DEFAULT (0),            
[MontoDIes] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_MontoDIes] DEFAULT (0),            
 [MontoSProv] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_MontoSProv] DEFAULT (0),            
 [MontoDProv] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRHONORARIOS_MontoDProv] DEFAULT (0)            
) ON [PRIMARY]            
            
END            
            
            
IF @Tabla = 'RETENCIONES'--TEMPORAL            
BEGIN            
      CREATE TABLE [TMPRETENCIONES] (            
 [Ase_cNummov] [char] (10)  NULL CONSTRAINT [DF_TMPRETENCIONES_Ase_cNummov] DEFAULT (''),            
 [Ase_nVoucher] [char] (10)  NULL CONSTRAINT [DF_TMPRETENCIONES_Ase_nVoucher] DEFAULT (''),            
 [Per_cPeriodo] [char] (20)  NULL CONSTRAINT [DF_TMPRETENCIONES_Per_cPeriodo] DEFAULT (''),            
 [Asd_nItem] [int] NULL CONSTRAINT [DF_TMPRETENCIONES_Asd_nItem] DEFAULT (0),            
 [Pla_cCuentaContable] [char] (12)  NULL CONSTRAINT [DF_TMPRETENCIONES_Pla_cCuentaContable] DEFAULT (''),            
 [Asd_cGlosa] [char] (250)  NULL CONSTRAINT [DF_TMPRETENCIONES_Asd_cGlosa] DEFAULT (''),            
 [Asd_nDebe] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPRETENCIONES_Asd_nDebeSoles] DEFAULT (0),            
 [Asd_nHaber] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPRETENCIONES_Asd_nHaberSoles] DEFAULT (0),            
 [Asd_nTipoCambio] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPRETENCIONES_Asd_nTipoCambio] DEFAULT (0),            
 [Asd_cTipoDoc] [char] (3)  NULL CONSTRAINT [DF_TMPRETENCIONES_Asd_cTipoDoc] DEFAULT (''),            
 [Tdo_cNombreLargo] [char] (250)  NULL CONSTRAINT [DF_TMPRETENCIONES_Tdo_cNombreLargo] DEFAULT (''),            
 [Asd_cSerieDoc] [varchar] (20)   NULL CONSTRAINT [DF_TMPRETENCIONES_Asd_cSerieDoc] DEFAULT (''),            
 [Asd_cNumDoc] [varchar] (25)  NULL CONSTRAINT [DF_TMPRETENCIONES_Asd_cNumDoc] DEFAULT (''),            
 [Asd_dFecDoc] [datetime]   NULL CONSTRAINT [DF_TMPRETENCIONES_Asd_dFecDoc] DEFAULT (''),            
 [Asd_dFecVen] [datetime]   NULL CONSTRAINT [DF_TMPRETENCIONES_Asd_dFecVen] DEFAULT (''),            
 [Asd_cTipoDocRef] [char] (3)  NULL CONSTRAINT [DF_TMPRETENCIONES_Asd_cTipoDocRef] DEFAULT (''),            
 [Tdo_cNombreLargo_Ref] [char] (250)  NULL CONSTRAINT [DF_TMPRETENCIONES_Tdo_cNombreLargo_Ref] DEFAULT (''),            
 [Asd_cSerieDocRef] [char] (5)   NULL CONSTRAINT [DF_TMPRETENCIONES_Asd_cSerieDocRef] DEFAULT (''),            
 [Asd_cNumDocRef] [char] (15)  NULL CONSTRAINT [DF_TMPRETENCIONES_Asd_cNumDocRef] DEFAULT (''),            
 [Asd_dFecDocRef] [datetime]   NULL CONSTRAINT [DF_TMPRETENCIONES_Asd_dFecDocRef] DEFAULT (''),            
 [Asd_nMontoInafecto] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPRETENCIONES_Asd_nMontoInafecto] DEFAULT (0),            
 [Asd_cRetencion] [char] (1)  NULL CONSTRAINT [DF_TMPRETENCIONES_Asd_cRetencion] DEFAULT (''),            
 [Asd_dFechaSpot] [datetime]   NULL CONSTRAINT [DF_TMPRETENCIONES_Asd_dFechaSpot] DEFAULT (''),            
 [Asd_cNumSpot] [char] (25)  NULL CONSTRAINT [DF_TMPRETENCIONES_Asd_cNumSpot] DEFAULT (''),            
 [Asd_nCorre] [INT]  NULL CONSTRAINT [DF_TMPRETENCIONES_] DEFAULT (0),            
 [Lib_cTrans] [char] (3)   NULL CONSTRAINT [DF_TMPRETENCIONES_Lib_cTrans] DEFAULT (''),            
 [Lib_cDescripcionTrans] [char] (250)  NULL CONSTRAINT [DF_TMPRETENCIONES_Lib_cDescripcionTrans] DEFAULT (''),            
            
 [Ase_cRuc] [char] (15)  NULL CONSTRAINT [DF_TMPRETENCIONES_Ase_cRuc] DEFAULT (''),            
 [Ase_cRazonSocial] [char] (250)  NULL CONSTRAINT [DF_TMPRETENCIONES_RazonSocial] DEFAULT (''),            
 [Ent_cCodEntidad]  [char] (5)  NULL CONSTRAINT [DF_TMPRETENCIONES_CodEntidad] DEFAULT (''),            
            
 [Ent_cVoucher]  [char] (10)  NULL CONSTRAINT [DF_TMPRETENCIONES_Ent_cVoucher] DEFAULT (''),            
 [Lib_cTipoLibro]  [char] (2)  NULL CONSTRAINT [DF_TMPRETENCIONES_Lib_cTipoLibro] DEFAULT (''),            
 [Asd_nAux]  [numeric](15, 3)  NULL CONSTRAINT [DF_TMPRETENCIONES_Asd_nAux] DEFAULT (0),            
 [Ten_cTipoEntidad]  [char] (1)  NULL CONSTRAINT [DF_TMPRETENCIONES_Ten_cTipoEntidad] DEFAULT (''),            
 [Asd_nSaldo]  [numeric](15, 3)  NULL CONSTRAINT [DF_TMPRETENCIONES_Asd_nSaldo] DEFAULT (0) ,            
 [Asd_cGrupo]  [varchar] (4)  NULL CONSTRAINT [DF_TMPRETENCIONES_Asd_cGrupo] DEFAULT ('')            
            
      ) ON [PRIMARY]            
END            
            
IF @Tabla = 'ENTIDADES'            
BEGIN            
            
CREATE TABLE [dbo].[TMP_PENDIENTES_ENTIDADES] (            
 [Ase_cNummov] [char] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,            
 [Per_cPeriodo] [char] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,            
 [Ase_dFecha] [char] (8) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,            
 [Lib_cTipoLibro] [char] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,            
 [Lib_cDescripcion] [varchar] (60) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,            
 [Ase_nVoucher] [char] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,            
 [Ase_cTipoMoneda] [char] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,            
 [Mon_cNombreLargo] [varchar] (60) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,            
 [Ase_nTipoCambio] [decimal](14, 3) NULL ,            
 [Ase_cGlosa] [varchar] (250) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,            
 [Asd_nItem] [int] NULL ,            
 [Pla_cCuentaContable] [varchar] (12) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,            
 [Pla_cNombreCuenta] [varchar] (120) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,            
 [Pla_cProvision] [char] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,            
 [Asd_cGlosa] [varchar] (250) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,            
 [Asd_nDebeSoles] [decimal](14, 3) NULL ,            
 [Asd_nDebeMonExt] [decimal](14, 3) NULL ,            
 [Asd_nHaberSoles] [decimal](14, 3) NULL ,            
 [Asd_nHaberMonExt] [decimal](14, 3) NULL ,            
 [Asd_nTipoCambio] [decimal](14, 3) NULL ,            
 [Cos_cCodigo] [varchar] (12) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,            
 [Cos_cDescripcion] [varchar] (60) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,            
 [Ten_cTipoEntidad] [char] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,            
 [Ent_cCodEntidad] [char] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,            
 [Ent_nRuc] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,            
 [Ent_cPersona] [varchar] (120) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,            
 [Ten_cNombreEntidad] [varchar] (60) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,            
 [Asd_cTipoDoc] [char] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,            
 [Asd_cSerieDoc] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,            
 [Asd_cNumDoc] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,            
 [Asd_dFecDoc] [datetime] NULL ,            
 [Asd_cTipoDocRef] [char] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,            
 [Asd_cSerieDocRef] [char] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,            
 [Asd_cNumDocRef] [char] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,            
 [Asd_dFecDocRef] [datetime] NULL ,            
 [Asd_nMontoInafecto] [decimal](14, 3) NULL ,            
 [Asd_cRetencion] [char] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,            
 [Asd_dFechaSpot] [datetime] NULL ,            
 [Asd_cNumSpot] [char] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,            
 [Asd_cDestino] [char] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,            
 [Asd_nCorre] [int] NULL ,            
 [Asd_cProvCanc] [char] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,            
 [Tdo_cNombreLargo] [varchar] (60) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,            
 [Tdo_cNombreCorto] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,            
 [Tdo_cNombreLargoRef] [varchar] (60) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,            
 [Tdo_cNombreCortoRef] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,            
 [Ase_cUserModifica] [varchar] (40) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,            
 [Emp_cNombreLargo] [varchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,            
 [Emp_cNombreCorto] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,            
 [FechaHoraImp] [datetime] NULL ,            
 [desde] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,            
 [hasta] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,            
 [Mon_cNombreLargoV] [varchar] (12) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,            
 [Mon_cMNac] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,            
 [Asd_cOperaTC] [char] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,            
 [Asd_dFecVen] [char] (8) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,            
 [USUARIO] [char] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,            
 [Asd_cMonAdic] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,            
 [Asd_cImpAdic] [decimal](14, 3) NULL ,            
 [Ase_dFechaISO] [char] (8) COLLATE SQL_Latin1_General_CP1_CI_AS NULL             
) ON [PRIMARY]            
            
            
--CREATE UNIQUE CLUSTERED INDEX IX_1 on #temp_employee_v1 (lname, fname, emp_id)            
            
--CREATE INDEX IX_2 on TMP_PENDIENTES_ENTIDADES (Pla_cCuentaContable,Ten_cTipoEntidad, Ent_cCodEntidad,Asd_cTipoDoc, Asd_cSerieDoc, Asd_cNumDoc)            
            
            
            
END            
            
            
IF @Tabla = 'REP_ESTANDAR'            
BEGIN            
            
CREATE TABLE [dbo].[TMP_REPESTANDAR] (            
     [Campo01] char (400) NULL ,            
     [Campo02] char (400) NULL ,            
     [Campo03] char (400) NULL ,            
     [Campo04] char (400) NULL ,            
     [Campo05] char (400) NULL ,            
     [Campo06] char (400) NULL ,            
     [Campo07] char (400) NULL             
) ON [PRIMARY]            
            
END            
            
IF @Tabla = 'PROC_CIERRE'            
BEGIN            
            
CREATE TABLE [dbo].[TMP_CIERRE] (            
     [Pla_cCuentaContable] [char] (12)  NULL ,            
     [MontoSoles] [decimal](14, 3) NULL ,            
     [MontoDolares] [decimal](14, 3) NULL             
) ON [PRIMARY]            
            
END            
            
            
            
IF @Tabla = 'REGVENTASF1401'            
BEGIN            
CREATE TABLE [TMPREGISTROVENTAS] (            
 [Emp_cCodigo] [char] (3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Emp_cCodigo] DEFAULT (''),            
 [Ase_cNumMov] [char] (10) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Ase_cNumMov] DEFAULT (''),            
 [Per_cPeriodo] [char] (2) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Per_cPeriodo] DEFAULT (''),            
 [Asd_dFecDoc] [datetime] NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_dFecDoc] DEFAULT (''),            
 [Asd_cTipoDoc] [char] (3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_cTipoDoc] DEFAULT (''),            
 [Asd_cSerieDoc] [varchar] (20) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_cSerieDoc] DEFAULT (''),            
 [Asd_cNumDoc] [varchar] (25) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_cNumDoc] DEFAULT (''),            
 [Asd_cTipoDocRef] [char] (3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_cTipoDocRef] DEFAULT (''),            
 [Asd_cSerieDocRef] [char] (5) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_cSerieDocRef] DEFAULT (''),            
 [Asd_cNumDocRef] [char] (12) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_cNumDocRef] DEFAULT (''),            
 [Ase_nVoucher] [char] (10) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Ase_nVoucher] DEFAULT (''),            
 [Ent_nRuc] [varchar] (15) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Ent_nRuc] DEFAULT (''),            
 [Ent_cPersona] [varchar] (120) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Ent_cPersona] DEFAULT (''),            
 [Asd_nItem] [int] NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_nItem] DEFAULT (0),            
 [MontoBaseGra] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_MontoBaseGra] DEFAULT (0),            
 [MontoInafecto] [numeric](14, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_MontoInafecto] DEFAULT (0),            
 [MontoBaseOtImp] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_MontoBaseOtImp] DEFAULT (0),            
 [MontoISC] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_MontoISC] DEFAULT (0),            
 [MontoIGV] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_MontoIGV] DEFAULT (0),            
 [MontoFOB] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_MontoFOB] DEFAULT (0),            
 [MontoFlete] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_MontoFlete] DEFAULT (0),            
 [MontoOtros] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_MontoOtros] DEFAULT (0),            
 [MontoDifCmb] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_MontoDifCmb] DEFAULT (0),            
 [MontoTotal] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_MontoTotal] DEFAULT (0),            
    [cNombreDoc] [varchar] (20)  NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_cNombreDoc] DEFAULT (''),            
 [MontoExp] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_MontoSExp] DEFAULT (0),            
            
 [Asd_dFecVen] [datetime] NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_dFecVen] DEFAULT (''),            
 [Asd_dFecDocRef] [datetime] NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_dFecDocRef] DEFAULT (''),            
 [Asd_cCodSunat] [char] (3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_cCodSunat] DEFAULT (''),            
 [MontoExon] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_MontoExon] DEFAULT (0),            
            
 [Ten_cTipoEntidad] [char] (1)  NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Ten_cTipoEntidad] DEFAULT (''),            
 [Ent_cCodEntidad] [char] (5)  NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Ent_cCodEntidad] DEFAULT (''),            
 [Asd_nTipoCambio] [numeric](14, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_nTipoCambio] DEFAULT (0)            
            
) ON [PRIMARY]            
            
END     
            
IF @Tabla = 'LibroElectVentas'            
BEGIN            
CREATE TABLE [TMPREGISTROLEVENTAS] (            
 [Emp_cCodigo] [char] (3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Emp_cCodigo] DEFAULT (''),            
 [Ase_cNumMov] [char] (10) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Ase_cNumMov] DEFAULT (''),            
 [Per_cPeriodo] [char] (2) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Per_cPeriodo] DEFAULT (''),            
 [Asd_dFecDoc] [datetime] NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_dFecDoc] DEFAULT (''),            
 [Asd_cTipoDoc] [char] (3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_cTipoDoc] DEFAULT (''),            
 [Asd_cSerieDoc] [varchar] (20) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_cSerieDoc] DEFAULT (''),            
 [Asd_cNumDoc] [varchar] (25) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_cNumDoc] DEFAULT (''),          
 [Ase_dFecha] [datetime] NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Ase_dFecha] DEFAULT (''),              
 [Asd_cTipoDocRef] [char] (3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_cTipoDocRef] DEFAULT (''),            
 [Asd_cSerieDocRef] [char] (5) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_cSerieDocRef] DEFAULT (''),            
 [Asd_cNumDocRef] [char] (12) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_cNumDocRef] DEFAULT (''),            
 [Ase_nVoucher] [char] (10) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Ase_nVoucher] DEFAULT (''),            
 [Ent_nRuc] [varchar] (15) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Ent_nRuc] DEFAULT (''),            
 [Ent_cPersona] [varchar] (120) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Ent_cPersona] DEFAULT (''),            
 [Asd_nItem] [int] NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_nItem] DEFAULT (0),            
 [MontoBaseGra] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_MontoBaseGra] DEFAULT (0),            
 [MontoInafecto] [numeric](14, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_MontoInafecto] DEFAULT (0),            
 [MontoBaseOtImp] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_MontoBaseOtImp] DEFAULT (0),            
 [MontoISC] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_MontoISC] DEFAULT (0),            
 [MontoIGV] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_MontoIGV] DEFAULT (0),            
 [MontoFOB] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_MontoFOB] DEFAULT (0),            
 [MontoFlete] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_MontoFlete] DEFAULT (0),            
 [MontoOtros] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_MontoOtros] DEFAULT (0),            
 [MontoDifCmb] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_MontoDifCmb] DEFAULT (0),            
 [MontoTotal] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_MontoTotal] DEFAULT (0),            
 [cNombreDoc] [varchar] (20)  NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_cNombreDoc] DEFAULT (''),            
 [MontoExp] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_MontoSExp] DEFAULT (0),  
 [Asd_dFecVen] [datetime] NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_dFecVen] DEFAULT (''),  
 [Asd_dFecDocRef] [datetime] NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_dFecDocRef] DEFAULT (''),  
 [Asd_cCodSunat] [char] (3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_cCodSunat] DEFAULT (''),            
 [MontoExon] [numeric](15, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_MontoExon] DEFAULT (0),            
 [Ten_cTipoEntidad] [char] (1)  NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Ten_cTipoEntidad] DEFAULT (''),            
 [Ent_cCodEntidad] [char] (5)  NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Ent_cCodEntidad] DEFAULT (''),            
 [Asd_nTipoCambio] [numeric](14, 3) NOT NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_nTipoCambio] DEFAULT (0),            
 [Asd_cEstadoO] [char](1)  NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_cEstadoO] DEFAULT (''),             
 [Asd_cEstadoD] [char](1)  NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_cEstadoD] DEFAULT (''),    
 [Asd_cImpAdic] [numeric](14,2)  NULL CONSTRAINT [DF_TMPREGISTRVENTAS_Asd_cImpAdic] DEFAULT (0)  
) ON [PRIMARY]            
            
END            
            
            
IF @Tabla = 'COMPRASF0801'            
BEGIN            
CREATE TABLE [TMPREGISTROCOMPRAS] (            
 [Emp_cCodigo] [char] (3) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Emp_cCodigo] DEFAULT (''),            
 [Ase_cNumMov] [char] (10) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Ase_cNumMov] DEFAULT (''),            
 [Per_cPeriodo] [char] (2) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Per_cPeriodo] DEFAULT (''),            
 [Ase_nVoucher] [char] (10) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Ase_nVoucher] DEFAULT (''),            
 [ANIO] [char] (4) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_ANIO] DEFAULT (''),            
 [Asd_nItem] [int]  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Asd_nItem] DEFAULT (0),            
 [Asd_nTipoCambio] [numeric](14, 3)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Asd_nTipoCambio] DEFAULT (0),            
 [Asd_cGlosa] [varchar] (250)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Asd_cGlosa] DEFAULT (''),            
 [Ent_cCodEntidad] [char] (5)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Ent_cCodEntidad] DEFAULT (''),            
 [Ent_nRuc] [varchar] (15)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Ent_nRuc] DEFAULT (''),            
 [Ent_cPersona] [varchar] (120)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Ent_cPersona] DEFAULT (''),            
 [Asd_cTipoDoc] [char] (3)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Asd_cTipoDoc] DEFAULT (''),            
 [Asd_cSerieDoc] [varchar] (20)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Asd_cSerieDoc] DEFAULT (''),            
 [Asd_cNumDoc] [varchar] (25)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Asd_cNumDoc] DEFAULT (''),            
 [Asd_dFecDoc] [datetime] NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Asd_dFecDoc] DEFAULT (''),            
 [Asd_cTipoDocRef] [char] (3)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Asd_cTipoDocRef] DEFAULT (''),            
 [Asd_cSerieDocRef] [char] (5)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Asd_cSerieDocRef] DEFAULT (''),            
 [Asd_cNumDocRef] [char] (10)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Asd_cNumDocRef] DEFAULT (''),            
 [Asd_dFecDocRef] [datetime] NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Asd_dFecDocRef] DEFAULT (''),            
 [Asd_nMontoInafecto] [numeric](14, 3)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Asd_nMontoInafecto] DEFAULT (0),            
 [Asd_cRetencion] [char] (1)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Asd_cRetencion] DEFAULT (''),            
 [Tdo_cNombreLargo] [varchar] (60)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Tdo_cNombreLargo] DEFAULT (''),            
 [Ase_cTipoMoneda] [char] (3)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Ase_cTipoMoneda] DEFAULT (''),            
 [CtaBaseA] [varchar] (12)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_CtaBaseA] DEFAULT (''),            
 [NombreCtaBaseA] [varchar] (120)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_NombreCtaBaseA] DEFAULT (''),            
 [CtaBaseB] [varchar] (12)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_CtaBaseB] DEFAULT (''),            
 [NombreCtaBaseB] [varchar] (120)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_NombreCtaBaseB] DEFAULT (''),            
 [CtaBaseC] [varchar] (12)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_CtaBaseC] DEFAULT (''),            
 [NombreCtaBaseC] [varchar] (120)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_NombreCtaBaseC] DEFAULT (''),            
 [CtaIgvA] [varchar] (12)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_CtaIgvA] DEFAULT (''),            
 [NombreIgvA] [varchar] (120)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_NombreIgvA] DEFAULT (''),            
 [CtaIgvB] [varchar] (12)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_CtaIgvB] DEFAULT (''),            
 [NombreIgvB] [varchar] (120)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_NombreIgvB] DEFAULT (''),            
 [CtaIgvC] [varchar] (12)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_CtaIgvC] DEFAULT (''),            
 [NombreIgvC] [varchar] (120)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_NombreIgvC] DEFAULT (''), [CtaProv] [varchar] (12) NOT NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_CtaProv] DEFAULT (''),            
 [NombreProv] [varchar] (120)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_NombreProv] DEFAULT (''),            
 [MontoSBaseA] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_MontoSBaseA] DEFAULT (0),            
 [MontoDBaseA] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_MontoDBaseA] DEFAULT (0),            
 [MontoSBaseB] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_MontoSBaseB] DEFAULT (0),            
 [MontoDBaseB] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_MontoDBaseB] DEFAULT (0),            
 [MontoSBaseC] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_MontoSBaseC] DEFAULT (0),            
 [MontoDBaseC] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_MontoDBaseC] DEFAULT (0),            
 [MontoSIgvA] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_MontoSIgvA] DEFAULT (0),            
 [MontoDIgvA] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_MontoDIgvA] DEFAULT (0),            
 [MontoSIgvB] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_MontoSIgvB] DEFAULT (0),            
 [MontoDIgvB] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_MontoDIgvB] DEFAULT (0),            
 [MontoSIgvC] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_MontoSIgvC] DEFAULT (0),            
 [MontoDIgvC] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_MontoDIgvC] DEFAULT (0),            
 [MontoSOtros] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_MontoSOtros] DEFAULT (0),            
 [MontoDOtros] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_MontoDOtros] DEFAULT (0),            
 [MontoSProv] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_MontoSProv] DEFAULT (0),            
 [MontoDProv] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_MontoDProv] DEFAULT (0),             
            
 [MontoSISC] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_MontoSISC] DEFAULT (0),            
 [MontoDISC] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_MontoDISC] DEFAULT (0),             
            
 [MontoSDIFC] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_MontoSDIFC] DEFAULT (0),        
 [MontoDDIFC] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_MontoDDIFC] DEFAULT (0),             
            
            
 [ASD_DFECHASPOT] [datetime] NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_ASD_DFECHASPOT] DEFAULT (''),            
 [ASD_CNUMSPOT] [char] (25)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_ASD_CNUMSPOT] DEFAULT (''),            
 [imp_nPorcentaje] [numeric](14, 3)   NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_imp_nPorcentaje] DEFAULT (0),             
 [Asd_dFecVen] [datetime] NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Asd_dFecVen] DEFAULT (''),            
            
 [Asd_cCodSunat] [char] (3)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Asd_cCodSunat] DEFAULT (''),            
 [Asd_cComprob] [char] (15)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Asd_cComprob] DEFAULT (''),            
            
 [Ten_cTipoEntidad] [char] (1)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_Ten_cTipoEntidad] DEFAULT (''),            
          
 [MontoSReintegro] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_MontoSReintegro] DEFAULT (0),             
 [MontoDReintegro] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPREGISTRCOMPRAS_MontoDReintegro] DEFAULT (0),      
           
            
) ON [PRIMARY]            
END            
            
IF @Tabla = 'LibroElectCompras'            
BEGIN            
CREATE TABLE [TMPREGISTROLECOMPRAS] (            
 [Emp_cCodigo] [char] (3) NOT NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_Emp_cCodigo] DEFAULT (''),            
 [Ase_cNumMov] [char] (10) NOT NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_Ase_cNumMov] DEFAULT (''),            
 [Per_cPeriodo] [char] (2) NOT NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_Per_cPeriodo] DEFAULT (''),            
 [Ase_nVoucher] [char] (10) NOT NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_Ase_nVoucher] DEFAULT (''),       
 [Ase_dfecha] [datetime] NOT NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_Ase_dfecha] DEFAULT (''),                  
 [ANIO] [char] (4) NOT NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_ANIO] DEFAULT (''),            
 [Asd_nItem] [int]  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_Asd_nItem] DEFAULT (0),            
 [Asd_nTipoCambio] [numeric](14, 3)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_Asd_nTipoCambio] DEFAULT (0),            
 [Asd_cGlosa] [varchar] (250)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_Asd_cGlosa] DEFAULT (''),            
 [Ent_cCodEntidad] [char] (5)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_Ent_cCodEntidad] DEFAULT (''),            
 [Ent_nRuc] [varchar] (15)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_Ent_nRuc] DEFAULT (''),            
 [Ent_cPersona] [varchar] (120)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_Ent_cPersona] DEFAULT (''),            
 [Asd_cTipoDoc] [char] (3)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_Asd_cTipoDoc] DEFAULT (''),            
 [Asd_cSerieDoc] [varchar] (20)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_Asd_cSerieDoc] DEFAULT (''),            
 [Asd_cNumDoc] [varchar] (25)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_Asd_cNumDoc] DEFAULT (''),            
 [Asd_dFecDoc] [datetime] NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_Asd_dFecDoc] DEFAULT (''),            
 [Asd_cTipoDocRef] [char] (3)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_Asd_cTipoDocRef] DEFAULT (''),            
 [Asd_cSerieDocRef] [char] (5)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_Asd_cSerieDocRef] DEFAULT (''),            
 [Asd_cNumDocRef] [char] (10)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_Asd_cNumDocRef] DEFAULT (''),            
 [Asd_dFecDocRef] [datetime] NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_Asd_dFecDocRef] DEFAULT (''),            
 [Asd_nMontoInafecto] [numeric](14, 3)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_Asd_nMontoInafecto] DEFAULT (0),            
 [Asd_cRetencion] [char] (1)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_Asd_cRetencion] DEFAULT (''),            
 [Tdo_cNombreLargo] [varchar] (60)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_Tdo_cNombreLargo] DEFAULT (''),            
 [Ase_cTipoMoneda] [char] (3)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_Ase_cTipoMoneda] DEFAULT (''),            
 [CtaBaseA] [varchar] (12)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_CtaBaseA] DEFAULT (''),            
 [NombreCtaBaseA] [varchar] (120)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_NombreCtaBaseA] DEFAULT (''),            
 [CtaBaseB] [varchar] (12)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_CtaBaseB] DEFAULT (''),            
 [NombreCtaBaseB] [varchar] (120)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_NombreCtaBaseB] DEFAULT (''),            
 [CtaBaseC] [varchar] (12)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_CtaBaseC] DEFAULT (''),            
 [NombreCtaBaseC] [varchar] (120)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_NombreCtaBaseC] DEFAULT (''),            
 [CtaIgvA] [varchar] (12)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_CtaIgvA] DEFAULT (''),            
 [NombreIgvA] [varchar] (120)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_NombreIgvA] DEFAULT (''),            
 [CtaIgvB] [varchar] (12)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_CtaIgvB] DEFAULT (''),            
 [NombreIgvB] [varchar] (120)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_NombreIgvB] DEFAULT (''),            
 [CtaIgvC] [varchar] (12)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_CtaIgvC] DEFAULT (''),            
 [NombreIgvC] [varchar] (120)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_NombreIgvC] DEFAULT (''),       
 [CtaProv] [varchar] (12) NOT NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_CtaProv] DEFAULT (''),            
 [NombreProv] [varchar] (120)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_NombreProv] DEFAULT (''),            
 [MontoSBaseA] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_MontoSBaseA] DEFAULT (0),            
 [MontoDBaseA] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_MontoDBaseA] DEFAULT (0),            
 [MontoSBaseB] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_MontoSBaseB] DEFAULT (0),            
 [MontoDBaseB] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_MontoDBaseB] DEFAULT (0),            
 [MontoSBaseC] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_MontoSBaseC] DEFAULT (0),            
 [MontoDBaseC] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_MontoDBaseC] DEFAULT (0),            
 [MontoSIgvA] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_MontoSIgvA] DEFAULT (0),            
 [MontoDIgvA] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_MontoDIgvA] DEFAULT (0),            
 [MontoSIgvB] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_MontoSIgvB] DEFAULT (0),            
 [MontoDIgvB] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_MontoDIgvB] DEFAULT (0),            
 [MontoSIgvC] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_MontoSIgvC] DEFAULT (0),            
 [MontoDIgvC] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_MontoDIgvC] DEFAULT (0),            
 [MontoSOtros] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_MontoSOtros] DEFAULT (0),            
 [MontoDOtros] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_MontoDOtros] DEFAULT (0),            
 [MontoSProv] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_MontoSProv] DEFAULT (0),            
 [MontoDProv] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_MontoDProv] DEFAULT (0),             
 [MontoSISC] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_MontoSISC] DEFAULT (0),            
 [MontoDISC] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_MontoDISC] DEFAULT (0),             
 [MontoSDIFC] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_MontoSDIFC] DEFAULT (0),            
 [MontoDDIFC] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_MontoDDIFC] DEFAULT (0),             
 [ASD_DFECHASPOT] [datetime] NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_ASD_DFECHASPOT] DEFAULT (''),            
 [ASD_CNUMSPOT] [char] (25)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_ASD_CNUMSPOT] DEFAULT (''),            
 [imp_nPorcentaje] [numeric](14, 3)   NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_imp_nPorcentaje] DEFAULT (0),             
 [Asd_dFecVen] [datetime] NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_Asd_dFecVen] DEFAULT (''),            
 [Asd_cCodSunat] [char] (3)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_Asd_cCodSunat] DEFAULT (''),            
 [Asd_cComprob] [char] (15)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_Asd_cComprob] DEFAULT (''),            
 [Ten_cTipoEntidad] [char] (1)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_Ten_cTipoEntidad] DEFAULT (''),            
 [MontoSReintegro] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_MontoSReintegro] DEFAULT (0),             
 [MontoDReintegro] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_MontoDReintegro] DEFAULT (0),      
 [Asd_cEstadoO] [char](1)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_Asd_cEstadoO] DEFAULT (''),             
 [Asd_cEstadoD] [char](1)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_Asd_cEstadoD] DEFAULT (''),      
 [DifCambio] [numeric](15, 3)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_DifCambio] DEFAULT (0)
 ,
 [Ent_cTipoDoc] [char](1)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_Ent_cTipoDoc] DEFAULT (''),      
 [Mon_cCodSunat] [char](3)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_Mon_cCodSunat] DEFAULT (''),
 [Id_Exoneracion] [char](10)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_IdExoneracion] DEFAULT (''),
 [Id_Tipo_Renta] [char](10)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_Tipo_Renta] DEFAULT (''),
 [Id_Modalidad] [char](10)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_Modalidad] DEFAULT (''),
 [Id_Aduana] [char](10)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_Aduana] DEFAULT (''),
 [Id_Clasific_Servicio] [char](10)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_Clasificacion] DEFAULT (''),
 [Ent_cFlagDomiciliado] [char](1)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_Ent_cFlagDomiciliado] DEFAULT ('1'),
 [Id_Pais] [char](10)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_Pais] DEFAULT (''),
 [Id_Convenio] [char](10)  NULL CONSTRAINT [DF_TMPREGISTROLECOMPRAS_Convenio] DEFAULT ('')
 
            
) ON [PRIMARY]            
END            
            
            
IF @Tabla = 'PATRIMONIO'            
BEGIN            
            
CREATE TABLE [dbo].[PATRIMONIO] (            
     emp_ccodigo char (3)  NULL ,            
     pan_canio   char (4)  NULL ,            
     pat_ccodigo char (3)  NULL ,            
     pat_ccol01  decimal(14, 3) NULL DEFAULT (0),            
     pat_ccol02  decimal(14, 3) NULL DEFAULT (0),            
     pat_ccol03  decimal(14, 3) NULL DEFAULT (0),            
     pat_ccol04  decimal(14, 3) NULL DEFAULT (0),            
     pat_ccol05  decimal(14, 3) NULL DEFAULT (0),            
     pat_ccol06  decimal(14, 3) NULL DEFAULT (0),            
     pat_ccol07  decimal(14, 3) NULL DEFAULT (0)            
) ON [PRIMARY]         
            
END            
            
            
IF @Tabla = 'REPORTEFLUJO'            
BEGIN            
            
CREATE TABLE [dbo].[REPORTEFLUJO] (            
     emp_ccodigo char (3)  NULL ,            
     pan_canio   char (4)  NULL ,            
     rep_ccodigo char (3)  NULL ,            
     pat_nsaldo  decimal(14, 3) NULL DEFAULT (0)            
) ON [PRIMARY]            
            
END            
            
            
IF @Tabla = 'PROCESO1002'            
BEGIN            
            
CREATE TABLE [dbo].[PROCESOS] (            
     pro_cgrupo  char (2)  NULL ,            
     pro_cDesgrupo  char (250)  NULL ,            
     emp_ccodigo  char (3)  NULL ,            
     pan_canio    char (4)  NULL ,            
     pro_ccodigo  char (12)  NULL ,            
     pro_cDescrip  char (250)  NULL ,            
     pro_cTitulo  char (1)  NULL ,            
     pro_cSTitulo  char (1)  NULL ,            
     pro_ncol01   decimal(14, 3) NULL DEFAULT (0),            
     pro_ncol02   decimal(14, 3) NULL DEFAULT (0),            
     pro_ncol03   decimal(14, 3) NULL DEFAULT (0),            
     pro_ncol04   decimal(14, 3) NULL DEFAULT (0),            
     pro_ncol05   decimal(14, 3) NULL DEFAULT (0),            
     pro_ncol06   decimal(14, 3) NULL DEFAULT (0),            
     pro_ncol07   decimal(14, 3) NULL DEFAULT (0),            
     pro_ncol08   decimal(14, 3) NULL DEFAULT (0),            
     pro_ncol09   decimal(14, 3) NULL DEFAULT (0),            
     pro_ncol10   decimal(14, 3) NULL DEFAULT (0),            
     pro_ncol11   decimal(14, 3) NULL DEFAULT (0),            
     pro_ncol12   decimal(14, 3) NULL DEFAULT (0),            
     pro_ncol13   decimal(14, 3) NULL DEFAULT (0)            
) ON [PRIMARY]            
            
END            
            
            
IF @Tabla = 'PROCESO1003'            
BEGIN            
            
CREATE TABLE [dbo].[PROCESOS] (            
     pro_cgrupo  char (2)  NULL ,            
  --   pro_cDesgrupo  char (250)  NULL ,            
     emp_ccodigo  char (3)  NULL ,            
     pan_canio    char (4)  NULL ,            
     pro_ccodigo  char (12)  NULL ,            
     pro_cDescrip  char (250)  NULL ,            
     pro_cTitulo  char (1)  NULL ,            
     pro_cSTitulo  char (1)  NULL ,            
     pro_ncol01   varchar(250) NULL,            
     pro_ncol02   varchar(250) NULL,            
     pro_ncol03   varchar(250) NULL,            
     pro_ncol04   varchar(250) NULL,            
     pro_ncol05   varchar(250) NULL,            
     pro_ncol06   varchar(250) NULL,            
     pro_ncol07   varchar(250) NULL,            
     pro_ncol08   varchar(250) NULL,            
     pro_ncol09   varchar(250) NULL,            
     pro_ncol10   varchar(250) NULL,            
     pro_ncol11   varchar(250) NULL,            
     pro_ncol12   varchar(250) NULL,            
     pro_ncol13   varchar(250) NULL            
) ON [PRIMARY]            
            
END            
            
IF @Tabla = 'COSTO_VALOR'            
BEGIN            
            
CREATE TABLE [dbo].[TMP_VALORACION_MENSUAL] (            
 Pan_cAnio char (4) NULL ,            
 Val_cTipo char (1) NULL ,            
 Cli_cCodEntidad char (5) NULL ,            
 Rdm_cCodArea char (3) NULL ,            
 Cos_cCodigo varchar (12) NULL ,            
 Val_nCosto numeric(14, 2) NULL ,            
 Val_cTurno char (1) NULL ,            
 Val_c01 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c02 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c03 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c04 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c05 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c06 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c07 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c08 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c09 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c10 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c11 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c12 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c13 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c14 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c15 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c16 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c17 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c18 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c19 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c20 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c21 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c22 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c23 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c24 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c25 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c26 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c27 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c28 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c29 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c30 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c31 numeric(14, 2) NULL  DEFAULT (0),            
 Val_nTotal numeric(14, 2) NULL  DEFAULT (0),     Val_nTotCosto numeric(14, 2) NULL  DEFAULT (0)            
) ON [PRIMARY]            
            
            
END            
            
            
IF @Tabla = 'TMP_EMPRESAS'            
BEGIN            
            
CREATE TABLE [dbo].[TMP_EMPRESAS] (            
     emp_ccodigo char (3)  NULL             
) ON [PRIMARY]            
            
END            
            
            
IF @Tabla = 'REP_PRESUPUESTO'            
BEGIN            
CREATE TABLE [dbo].[TMP_PRESUPUESTO] (            
 Periodo char(2) null,         
 Cos_cDescripcion VARCHAR(400)  NULL ,            
 COS_CCODIGO VARCHAR(12)  NULL ,            
 EJECSOLESACUM numeric(14, 2) NULL  DEFAULT (0),            
 EJECDOLARACUM numeric(14, 2) NULL  DEFAULT (0),            
 EJECSOLES numeric(14, 2) NULL  DEFAULT (0),            
 EJECDOLAR numeric(14, 2) NULL  DEFAULT (0),            
 PRESUPSOLESACUM numeric(14, 2) NULL  DEFAULT (0),            
 PRESUPDOLARACUM numeric(14, 2) NULL  DEFAULT (0),            
 PRESUPSOLES numeric(14, 2) NULL  DEFAULT (0),            
 PRESUPDOLAR numeric(14, 2) NULL  DEFAULT (0),            
 NIVEL CHAR(1)  NULL ,            
 DETALLE CHAR(1)  NULL ,            
 MONNAC CHAR(1)  NULL ,            
 PRM_CTIPO CHAR(1)  NULL ,            
 NOMMONEDA VARCHAR(50)  NULL             
) ON [PRIMARY]            
            
            
END            
            
            
IF @Tabla = 'REP_DIARIO'            
BEGIN            
            
CREATE TABLE [dbo].[TMP_REP_DIARIO] (            
 Pan_cAnio char (4) NULL ,            
 Val_cTipo char (1) NULL ,            
 Ten_cTipoEntidad char (1) NULL ,            
 Cod_cCodEntidad char (5) NULL ,            
 Rdm_cCodArea char (3) NULL ,            
 Cos_cCodigo varchar (12) NULL ,            
 Val_cTurno char (1) NULL ,            
 Val_c01 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c02 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c03 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c04 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c05 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c06 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c07 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c08 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c09 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c10 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c11 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c12 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c13 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c14 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c15 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c16 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c17 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c18 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c19 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c20 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c21 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c22 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c23 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c24 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c25 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c26 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c27 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c28 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c29 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c30 numeric(14, 2) NULL  DEFAULT (0),            
 Val_c31 numeric(14, 2) NULL  DEFAULT (0),            
 Val_nTotal numeric(14, 2) NULL  DEFAULT (0)            
            
) ON [PRIMARY]            
            
END
GO
