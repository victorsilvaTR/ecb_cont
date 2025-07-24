USE SAFC_ECB
GO

USE SAFC_ECB
GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO

ALTER PROCEDURE [dbo].[spCn_LibroElectronicoVentas3]          
 @Emp_cCodigo char(3)='',                          
 @Pan_cAnio char(4)='',                    
 @desde varchar(10) = '',                    
 @hasta varchar(10) = '',                    
 @moneda char(3) = '',                    
 @Per_cPeriodo char(2)='',
 @Simplificado CHAR(1) = ''                    
          
--WITH ENCRYPTION                                              
AS      
SET DATEFORMAT DMY      
set nocount on      
      
declare @RUC char(50)          
select @RUC= dbo.RUC(@Emp_cCodigo)          
--------------          
declare @Texto varchar(100)          
set @Texto =''          
          
-- *** Hallando el tipo de Moneda(Si es Nac o No) y su nombre                                              
DECLARE @Mon_cNombreLargo varchar(100)                                              
DECLARE @Mon_cMNac char(1)                                              
DECLARE @NombreEmpresa varchar(250)                                              
DECLARE @Emp_cNumRuc varchar (15)                                              
DECLARE @vItem int                                              
DECLARE @DocNC Char(1)                                              
----------------------------------------                                              
--- Cursor Cabecera                                              
Declare @vAse_cNummov Char(10)                                              
Declare @vEmp_cCodigo Char(3)                                              
Declare @vPan_cAnio Char(4)                                              
Declare @vPer_cPeriodo Char(2)                                              
Declare @vLib_cTipoLibro Char(2)                                              
Declare @vAse_nVoucher Char(10)                                              
Declare @vAse_dFecha DateTime                                              
Declare @vAse_cGlosa VarChar(250)                                              
Declare @vAse_cTipoMoneda Char(3)                                              
Declare @vAse_nTipoCambio Numeric(14,3)                                              
Declare @vAse_cEstado Char(1)                        
-- Campos temporal detalle                    
Declare @TAse_cNummov Char(10)                    
Declare @TEmp_cCodigo Char(3)                    
Declare @TPan_cAnio Char(4)                    
Declare @TPer_cPeriodo Char(2)                    
Declare @TLib_cTipoLibro Char(2)                    
Declare @TAse_nVoucher Char(10)                     
--- Cursor Detalle                                              
Declare @vAsd_nItem Int                                              
Declare @vPla_cCuentaContable VarChar(12)                                              
Declare @vAsd_cGlosa VarChar(250)                                              
Declare @vAsd_nDebeSoles Numeric(14,3)                                              
Declare @vAsd_nHaberSoles Numeric(14,3)                                              
Declare @vAsd_nTipoCambio Numeric(14,3)                                              
Declare @vAsd_nDebeMonExt Numeric(14,3)                                              
Declare @vAsd_nHaberMonExt Numeric(14,3)                                              
Declare @vCos_cCodigo VarChar(12)                                              
Declare @vTen_cTipoEntidad Char(1)                                              
Declare @vEnt_cCodEntidad Char(5)                                              
Declare @vAsd_cTipoDoc Char(3)                                              
Declare @vAsd_dFecDoc DateTime                                              
Declare @vAsd_cSerieDoc VARCHAR(20)                                  
Declare @vAsd_cNumDoc VARCHAR(25)                                              
Declare @vAsd_cTipoDocRef Char(3)                                              
Declare @vAsd_dFecDocRef DateTime                                              
Declare @vAsd_cSerieDocRef Char(5)                   
Declare @vAsd_cNumDocRef VarChar(12)                                              
Declare @vAsd_nMontoInafecto Numeric(14,3)                                  
Declare @vAsd_cTipoMoneda Char(3)                                              
Declare @vEnt_nRucT VarChar(15)                                              
Declare @vEnt_cNombre VarChar(50)                                              
Declare @Tdo_cNombreCortoT VarChar(15)                               
Declare @vAsd_cBaseImp VarChar(3)                    
Declare @vAsd_cCodSunat VarChar(3)                    
Declare @vAsd_dFecVen DateTime                    
declare @Ase_dFecha datetime                    
                                         
--Declare @vAsd_dFecDocRef DateTime                                              
                                              
--- Tabla Temporal de Impresion                                              
Declare @FecEmi DateTime                                              
Declare @CmpTd VarChar(15)                                              
Declare @CmpSr VARCHAR(20)                          
Declare @CmpNm VarChar(25)                                              
Declare @Reffec datetime                                              
Declare @RefTd VarChar(15)                                              
Declare @RefSr Char(5)                                              
Declare @RefNm Char(10)                                              
Declare @NumVou Char(10)                                              
Declare @NumRuc VarChar(15)                                               
Declare @NomClie VarChar(250)                                  
Declare @TipoEntidad VarChar(1)                                              
Declare @CodEntidad VarChar(5)                                              
Declare @BaseImpo Numeric(14,3)                                              
Declare @cTipoIgv char(1)                                              
Declare @BaseInaf Numeric(14,3)                                               
Declare @BaseIpOIm Numeric(14,3)                                               
Declare @ImpISC Numeric(14,3)                                              
Declare @ImpIGV Numeric(14,3)                                              
Declare @ImpFOB Numeric(14,3)                                              
Declare @ImpFlete Numeric(14,3)                                              
Declare @ImpOtros Numeric(14,3)                                              
Declare @ImpDifCmb Numeric(14,3)                                              
Declare @ImpTotal Numeric(14,3)                                              
Declare @cNombreDoc VarChar(15)                                              
Declare @Bonif Numeric(14,3)                                              
Declare @Exon Numeric (14,3)                                              
Declare @MontoExp Numeric(14,3)                                              
Declare @cTipDocNC CHAR(2)                    
Declare @Asd_cEstadoO char(1)                    
Declare @Asd_cEstadoD char(1)    
declare @ValEmb decimal (14,2)  
declare @Asd_cImpAdic decimal(14,2)                  
                                              
Declare @TC_CTA40 Numeric(14,3)                              
----------------------------------------                         
Declare @Numreg as int                    
set @Numreg = 0                    
------------------------                    
Declare @sqlTMP varchar(MAX)            
          
                                         
SELECT @Mon_cNombreLargo = Mon_cNombreLargo, @Mon_cMNac = Mon_cMNac                                 
FROM CNT_TIPO_MONEDA                         
WHERE Emp_cCodigo = @Emp_cCodigo and Mon_cCodigo = @moneda                              
                                              
SET @Mon_cMNac= ISNULL(@Mon_cMNac, '')                                               
----------------------------------------------       
                                              
DECLARE @Lib_cTipoLibro CHAR(2)                                              
select @Lib_cTipoLibro= CFL_CVENTAS from CNT_CONFIG_LIBROS where EMP_CCODIGO=@Emp_cCodigo and Pan_cAnio=@Pan_cAnio                 
SET @Lib_cTipoLibro = ISNULL(@Lib_cTipoLibro , '')                              
-------------------------------------------------------------------------                                              
SELECT @NombreEmpresa = Emp_cNombreLargo, @Emp_cNumRuc = Emp_cNumRuc                                               
FROM EMPRESA WHERE Emp_cCodigo = @Emp_cCodigo                                              
------------------------------------------------------------------------                                               
                                              
Select  Ase_cNummov,Emp_cCodigo,Pan_cAnio,Per_cPeriodo,Lib_cTipoLibro,Ase_nVoucher,          
     Ase_dFecha,Ase_cGlosa,Ase_cTipoMoneda,Ase_nTipoCambio,Ase_cEstado,Asd_cEstadoO, Asd_cEstadoD          
Into #TMPVOUCHER_CAB          
From  CNC_ASIENTO_VOUCHER  WITH (NOLOCK)          
Where  Emp_cCodigo=@Emp_cCodigo And          
       Pan_cAnio=@Pan_cAnio And          
 --Per_cperiodo >'00' and          
       Lib_cTipoLibro IN (SELECT Cfl_cVentas  FROM CNT_CONFIG_LIBROS WHERE Emp_cCodigo=@Emp_cCodigo and Pan_cAnio=@Pan_cAnio) And          
--        Ase_dFecha >= @desde AND Ase_dFecha <= @hasta  And          
  Per_cPeriodo = @Per_cPeriodo  and (CNC_ASIENTO_VOUCHER.Asd_cEstadoD not in ('9','8'))  and          
       (Ase_cEstado='A' Or Ase_cEstado='X')  And          
       Ase_cDeleted<>'*'          
                    
/******  Incluye registros con estado 8 y 9   *********/                    
Select  Ase_cNummov,Emp_cCodigo,Pan_cAnio,Per_cPeriodo,Lib_cTipoLibro,Ase_nVoucher,                                              
     Ase_dFecha,Ase_cGlosa,Ase_cTipoMoneda,Ase_nTipoCambio,Ase_cEstado,Asd_cEstadoO, Asd_cEstadoD                    
Into #TMPVOUCHER_CAB2                    
From  CNC_ASIENTO_VOUCHER  WITH (NOLOCK)                     
Where  Emp_cCodigo=@Emp_cCodigo And                     
      -- Pan_cAnio=@Pan_cAnio And                     
       Lib_cTipoLibro IN (SELECT Cfl_cVentas  FROM CNT_CONFIG_LIBROS WHERE Emp_cCodigo= @Emp_cCodigo and Pan_cAnio=@Pan_cAnio)                    
  --Per_cPeriodo = @Per_cPeriodo and                     
  AND year(CNC_ASIENTO_VOUCHER.Ase_dfechaModifica) = @Pan_cAnio/*YEAR(getdate())*/ and month(CNC_ASIENTO_VOUCHER.Ase_dfechaModifica) = @Per_cPeriodo                  
  and (CNC_ASIENTO_VOUCHER.Asd_cEstadoD = '9' or CNC_ASIENTO_VOUCHER.Asd_cEstadoD = '8') and                    
       (Ase_cEstado='A' Or Ase_cEstado='X') And Ase_cDeleted<>'*'                        
                                              
insert into #TMPVOUCHER_CAB select * from #TMPVOUCHER_CAB2                    
                    
                           
--select * from #TMPVOUCHER_CAB --where Ase_nVoucher= '0503000004'                                              
--return                                              
                                       
-----------------------------------------------------------------------------------------------                                              
SET @cTipDocNC=(SELECT DBO.TRIMSQL(Cod_cValorParam) FROM CND_CONFIG_OPERA                                              
WHERE Emp_cCodigo=@Emp_cCodigo AND Pan_cAnio=@Pan_cAnio AND Cop_cCodigo='012')                                              
                                              
SET @cTipDocNC= ISNULL(@cTipDocNC, '')                                               
-----------------------------------------------------------------------------------------------                                              
EXEC spCn_CrearTablaTemporal1 'LibroElectVentas'                                            
SELECT * INTO #TMPREGISTROLEVENTAS FROM TMPREGISTROLEVENTAS                                              
DROP TABLE TMPREGISTROLEVENTAS                                              
  
-----------------------------------------------------------------------------------------------                                              
Set @vAsd_nItem=0                                              
-----------------------------------------------------------------------------------------------                                              
                                              
DECLARE @CTA_DIFGAN varchar(12)                                              
DECLARE @CTA_DIFPER varchar(12)                                              
DECLARE @CTA_REDGAN varchar(12)                                              
DECLARE @CTA_REDPER varchar(12)                                              
                                              
SELECT @CTA_DIFGAN = Pla_cCuentaContable FROM CNM_PLAN_CTA  WITH (NOLOCK)                                              
WHERE Emp_cCodigo = @Emp_cCodigo AND Pan_cAnio = @Pan_cAnio AND pla_cDifCambio = 'G'                                              
                                              
SELECT @CTA_DIFPER = Pla_cCuentaContable FROM CNM_PLAN_CTA  WITH (NOLOCK)                                              
WHERE Emp_cCodigo = @Emp_cCodigo AND Pan_cAnio = @Pan_cAnio AND pla_cDifCambio = 'P'                                              
                                              
SELECT @CTA_REDGAN = Pla_cCuentaContable FROM CNM_PLAN_CTA  WITH (NOLOCK)                                              
WHERE Emp_cCodigo = @Emp_cCodigo AND Pan_cAnio = @Pan_cAnio AND Pla_cRedondeo = 'G'                                              
                                   
SELECT @CTA_REDPER = Pla_cCuentaContable FROM CNM_PLAN_CTA  WITH (NOLOCK)                                              
WHERE Emp_cCodigo = @Emp_cCodigo AND Pan_cAnio = @Pan_cAnio AND Pla_cRedondeo = 'P'                                              
                                              
-----------------------------------------------------------------------------------------------                              
DECLARE @VariableTabla TABLE                                               
(                                              
 Ase_nVoucher char(10),                                              
 Ase_cNummov char(10),                                              
 Asd_nItem int,                                              
 Pla_cCuentaContable  varchar(12),                                              
 Asd_cGlosa varchar(250),                                              
 Asd_nDebeSoles numeric(14,3),                                              
 Asd_nHaberSoles numeric(14,3),                                              
 Asd_nTipoCambio numeric(14,3),                                              
 Asd_nDebeMonExt numeric(14,3),                                              
 Asd_nHaberMonExt numeric(14,3),                                              
 Cos_cCodigo varchar(12),                                              
 Ten_cTipoEntidad char(1),                                              
 Ent_cCodEntidad char(5),                                              
 Asd_cTipoDoc char(2),                                              
 Asd_dFecDoc datetime,                                              
 Asd_cSerieDoc  varchar(20),                                              
 Asd_cNumDoc varchar(25),                                              
 Asd_cTipoDocRef  char(2),                                              
 Asd_cSerieDocRef  varchar(5),                                              
 Asd_cNumDocRef  varchar(12),                                               
 Asd_nMontoInafecto numeric(14,3),                                             
 Asd_cTipoMoneda char(3),                                              
 Ent_nRucT varchar(50),                                              
 Ent_cNombre varchar(100),                                              
 Tdo_cNombreCortoT varchar(100),                                              
 Asd_cBaseImp char(3),                                              
 Asd_dFecVen datetime,                                              
 Asd_dFecDocRef datetime,  
 Asd_cImpAdic decimal(14,2)  
)                                              
                   
                                             
-----------------------------------------------------------------------------------------------                                              
-- insert into @VariableTabla                                              
 Select                                               
 CND_ASIENTO_VOUCHER.Ase_nVoucher,                                              
 CND_ASIENTO_VOUCHER.Ase_cNummov,                                              
 CND_ASIENTO_VOUCHER.Asd_nItem,                                              
 CND_ASIENTO_VOUCHER.Pla_cCuentaContable,                                              
 CND_ASIENTO_VOUCHER.Asd_cGlosa,                                              
    CND_ASIENTO_VOUCHER.Asd_nDebeSoles,                                              
 CND_ASIENTO_VOUCHER.Asd_nHaberSoles,                               
 CND_ASIENTO_VOUCHER.Asd_nTipoCambio,                                           
    CND_ASIENTO_VOUCHER.Asd_nDebeMonExt,                                              
 CND_ASIENTO_VOUCHER.Asd_nHaberMonExt,                                              
 CND_ASIENTO_VOUCHER.Cos_cCodigo,                                              
    CND_ASIENTO_VOUCHER.Ten_cTipoEntidad,                                              
 CND_ASIENTO_VOUCHER.Ent_cCodEntidad,                                              
 CND_ASIENTO_VOUCHER.Asd_cTipoDoc,                                              
    CND_ASIENTO_VOUCHER.Asd_dFecDoc,                                              
 CND_ASIENTO_VOUCHER.Asd_cSerieDoc,                                              
 CND_ASIENTO_VOUCHER.Asd_cNumDoc,                                              
    CND_ASIENTO_VOUCHER.Asd_cTipoDocRef,                                              
 CND_ASIENTO_VOUCHER.Asd_cSerieDocRef,                                              
    CND_ASIENTO_VOUCHER.Asd_cNumDocRef,                                               
 CND_ASIENTO_VOUCHER.Asd_nMontoInafecto,                                              
 CND_ASIENTO_VOUCHER.Asd_cTipoMoneda,                                              
 @Texto AS Ent_nRucT ,                                              
 @Texto As Ent_cNombre,                                      
 @Texto AS Tdo_cNombreCortoT,                                              
 CND_ASIENTO_VOUCHER.Asd_cBaseImp,                                              
 CND_ASIENTO_VOUCHER.Asd_dFecVen,                                              
 CND_ASIENTO_VOUCHER.Asd_dFecDocRef,  
 CND_ASIENTO_VOUCHER.Asd_cImpAdic                                       
into #TMP_DETALLE                                              
From   CND_ASIENTO_VOUCHER  WITH (NOLOCK)                              
Where                              
 CND_ASIENTO_VOUCHER.Emp_cCodigo=@Emp_cCodigo And                              
    CND_ASIENTO_VOUCHER.Pan_cAnio=@Pan_cAnio And                     
    CND_ASIENTO_VOUCHER.Per_cPeriodo=@Per_cPeriodo And                               
    CND_ASIENTO_VOUCHER.Lib_cTipoLibro=@Lib_cTipoLibro And                              
 CND_ASIENTO_VOUCHER.Asd_cDeleted <>'*'                                          
                        
                                              
Declare Cursor_DetPend  CURSOR LOCAL FORWARD_ONLY STATIC READ_ONLY  For                            
  select Ase_cNummov,Emp_cCodigo,Pan_cAnio,Per_cPeriodo,Lib_cTipoLibro,Ase_nVoucher from  #TMPVOUCHER_CAB2                            
Open Cursor_DetPend                     
Fetch Next From Cursor_DetPend Into @TAse_cNummov,@TEmp_cCodigo,@TPan_cAnio,@TPer_cPeriodo,@TLib_cTipoLibro,@TAse_nVoucher                    
WHILE @@FETCH_STATUS = 0                                      
Begin                       
insert into @VariableTabla                     
 Select                                               
 CND_ASIENTO_VOUCHER.Ase_nVoucher, CND_ASIENTO_VOUCHER.Ase_cNummov,                                              
 CND_ASIENTO_VOUCHER.Asd_nItem,  CND_ASIENTO_VOUCHER.Pla_cCuentaContable,                                              
 CND_ASIENTO_VOUCHER.Asd_cGlosa,    CND_ASIENTO_VOUCHER.Asd_nDebeSoles,                                              
 CND_ASIENTO_VOUCHER.Asd_nHaberSoles, CND_ASIENTO_VOUCHER.Asd_nTipoCambio,                                           
    CND_ASIENTO_VOUCHER.Asd_nDebeMonExt, CND_ASIENTO_VOUCHER.Asd_nHaberMonExt,                                              
 CND_ASIENTO_VOUCHER.Cos_cCodigo,  CND_ASIENTO_VOUCHER.Ten_cTipoEntidad,                                              
 CND_ASIENTO_VOUCHER.Ent_cCodEntidad, CND_ASIENTO_VOUCHER.Asd_cTipoDoc,                                              
    CND_ASIENTO_VOUCHER.Asd_dFecDoc, CND_ASIENTO_VOUCHER.Asd_cSerieDoc,                                              
 CND_ASIENTO_VOUCHER.Asd_cNumDoc, CND_ASIENTO_VOUCHER.Asd_cTipoDocRef,                                              
 CND_ASIENTO_VOUCHER.Asd_cSerieDocRef,    CND_ASIENTO_VOUCHER.Asd_cNumDocRef,                                               
 CND_ASIENTO_VOUCHER.Asd_nMontoInafecto, CND_ASIENTO_VOUCHER.Asd_cTipoMoneda,                                              
 @Texto AS Ent_nRucT , @Texto As Ent_cNombre,                                              
 @Texto AS Tdo_cNombreCortoT, CND_ASIENTO_VOUCHER.Asd_cBaseImp,                                              
 CND_ASIENTO_VOUCHER.Asd_dFecVen, CND_ASIENTO_VOUCHER.Asd_dFecDocRef,CND_ASIENTO_VOUCHER.Asd_cImpAdic                                       
--into #TMP_DETALLE2                    
From   CND_ASIENTO_VOUCHER  WITH (NOLOCK)                              
Where                      
CND_ASIENTO_VOUCHER.Ase_cNummov = @TAse_cNummov and                    
 CND_ASIENTO_VOUCHER.Emp_cCodigo = @TEmp_cCodigo And                              
    CND_ASIENTO_VOUCHER.Pan_cAnio = @TPan_cAnio And                              
    CND_ASIENTO_VOUCHER.Per_cPeriodo = @TPer_cPeriodo And                               
    CND_ASIENTO_VOUCHER.Lib_cTipoLibro = @TLib_cTipoLibro And                              
  CND_ASIENTO_VOUCHER.Ase_nVoucher = @TAse_nVoucher                    
                    
   Fetch Next From Cursor_DetPend                                               
   Into @TAse_cNummov,@TEmp_cCodigo,@TPan_cAnio,@TPer_cPeriodo,@TLib_cTipoLibro,@TAse_nVoucher                                          
  End                        
                    
Close Cursor_DetPend                                              
Deallocate Cursor_DetPend                      
                    
insert into #TMP_DETALLE select * from @VariableTabla  
                    
--drop table @VariableTabla                    
drop table #TMPVOUCHER_CAB2                    
                    
                                              
CREATE INDEX IX_1 on #TMP_DETALLE (Ase_nVoucher,Ase_cNummov)                            
--select * from #TMPVOUCHER_CAB                                            
--select * from #TMP_DETALLE                          
   Print '============== INICIO DEL CURSOR '                            
-----------------------------------------------------------------------------------------------                                  
                                              
Declare Cursor_Cabecera  CURSOR LOCAL FORWARD_ONLY STATIC READ_ONLY  For                                              
Select                                               
 Ase_cNummov,Emp_cCodigo,Pan_cAnio,Per_cPeriodo,Lib_cTipoLibro,Ase_nVoucher,                                              
 Ase_dFecha,Ase_cGlosa,Ase_cTipoMoneda,Ase_nTipoCambio,Ase_cEstado,Asd_cEstadoO, Asd_cEstadoD                                          
From #TMPVOUCHER_CAB  WITH (NOLOCK)                                              
Open Cursor_Cabecera                                               
Fetch Next From Cursor_Cabecera Into @vAse_cNummov,@vEmp_cCodigo,@vPan_cAnio,@vPer_cPeriodo,@vLib_cTipoLibro,                                              
                                     @vAse_nVoucher,@vAse_dFecha,@vAse_cGlosa,@vAse_cTipoMoneda,@vAse_nTipoCambio,                                              
                                     @vAse_cEstado,@Asd_cEstadoO, @Asd_cEstadoD  
                    
WHILE @@FETCH_STATUS = 0                                                
Begin                                 
  -----------------------------------------------------------------------------------------------                                              
  Set @CmpTd=''                                              
  Set @CmpSr=''                                              
  Set @CmpNm=''                                              
  Set @RefTd=''                                               
  Set @RefSr=''                       
  Set @RefNm=''                                              
  Set @NumVou=''                                              
  Set @NumRuc=''                   
  Set @NomClie=''                                              
  Set @DocNC='0'                                              
                                              
  Set @BaseImpo= 0                                              
  Set @BaseInaf= 0                                              
  Set @ImpISC= 0                                              
  Set @ImpIGV= 0                         
  Set @ImpFOB= 0                                              
  Set @ImpFlete= 0                                              
  Set @ImpOtros= 0                                              
  Set @ImpDifCmb= 0                                              
  Set @ImpTotal= 0                                              
  SET @MontoExp =0                                              
  Set @Bonif= 0                                              
  SET @Exon = 0                                              
  SET @BaseIpOIm = 0      
  set @Asd_cImpAdic = 0     
  set @ValEmb=0                                       
                                              
                                              
  Declare Cursor_Detalle  CURSOR LOCAL FORWARD_ONLY STATIC READ_ONLY  For                                               
          Select DET.Asd_nItem,                                              
     DET.Pla_cCuentaContable,                                              
     DET.Asd_cGlosa,                                              
     DET.Asd_nDebeSoles,                                              
     DET.Asd_nHaberSoles,                                              
     DET.Asd_nTipoCambio,                                              
     DET.Asd_nDebeMonExt,                                              
     DET.Asd_nHaberMonExt,                                              
     DET.Cos_cCodigo,                                              
     DET.Ten_cTipoEntidad, DET.Ent_cCodEntidad,                                              
     DET.Asd_cTipoDoc, DET.Asd_cSerieDoc, DET.Asd_cNumDoc, DET.Asd_dFecDoc,                                               
     DET.Asd_cTipoDocRef, DET.Asd_cSerieDocRef, DET.Asd_cNumDocRef, DET.Asd_dFecDocRef,                                     
     DET.Asd_nMontoInafecto,                                              
     DET.Asd_cTipoMoneda,                                              
     @Texto AS Ent_nRucT ,                                              
     @Texto As Ent_cNombre,                                              
     @Texto AS Tdo_cNombreCortoT,                                              
      DET.Asd_cBaseImp,                                              
     DET.Asd_dFecVen,DET.Asd_cImpAdic                                              
          From   #TMP_DETALLE as DET  WITH (NOLOCK)                                              
          Where  DET.Ase_nVoucher=@vAse_nVoucher And                        
   DET.Ase_cNummov=@vAse_cNummov                                       
                                    
  SET @TC_CTA40 = 0       
                                   
  Open Cursor_Detalle                                              
  Fetch Next From Cursor_Detalle                                               
  Into @vAsd_nItem, @vPla_cCuentaContable, @vAsd_cGlosa, @vAsd_nDebeSoles, @vAsd_nHaberSoles,                                              
        @vAsd_nTipoCambio, @vAsd_nDebeMonExt, @vAsd_nHaberMonExt, @vCos_cCodigo, @vTen_cTipoEntidad,                                              
      @vEnt_cCodEntidad, @vAsd_cTipoDoc, @vAsd_cSerieDoc, @vAsd_cNumDoc, @vAsd_dFecDoc,                   
        @vAsd_cTipoDocRef, @vAsd_cSerieDocRef, @vAsd_cNumDocRef, @vAsd_dFecDocRef,                                              
        @vAsd_nMontoInafecto, @vAsd_cTipoMoneda, @vEnt_nRucT, @vEnt_cNombre, @Tdo_cNombreCortoT,                                              
 @vAsd_cBaseImp, @vAsd_dFecVen, @Asd_cImpAdic                                            
  While @@FETCH_STATUS = 0                                              
  Begin                                              
     -----------------------------------------------------------------------------------------------                                                                           
                                           
     If @vAse_cEstado='X'                                              
   Begin                                              
    Set @FecEmi= @vAsd_dFecDoc                                              
    If @vAsd_cTipoDoc<>'' Set @CmpTd= @vAsd_cTipoDoc                                              
    If @vAsd_cSerieDoc<>'' Set @CmpSr= @vAsd_cSerieDoc                                              
    If @vAsd_cNumDoc<>'' Set @CmpNm= @vAsd_cNumDoc
    
    IF LEFT(@vPla_cCuentaContable, 2) = '12'
    BEGIN
    	If @vAsd_cTipoDocRef<>'' Set @RefTd= @vAsd_cTipoDocRef                                              
		If @vAsd_cSerieDocRef<>'' Set @RefSr= @vAsd_cSerieDocRef                                    
		If @vAsd_cNumDocRef<>'' Set @RefNm= @vAsd_cNumDocRef                                              
		If @vAsd_dFecDocRef<>'' Set @Reffec= @vAsd_dFecDocRef 
    END
                                                  
    
                                                 
    If @vAse_nVoucher<>'' Set @NumVou= @vAse_nVoucher                                              
    If @Tdo_cNombreCortoT<>'' Set @cNombreDoc=@Tdo_cNombreCortoT                                              
                                                
    Set @NomClie= 'A   N   U   L   A   D   O'                                              
    Set @BaseImpo= 0                                              
    Set @BaseInaf= 0                                              
    Set @ImpISC= 0                                              
    Set @ImpIGV= 0                                              
    Set @ImpFOB= 0                                              
    Set @ImpFlete= 0                                              
    Set @ImpOtros= 0                                              
    Set @ImpDifCmb= 0                                              
    Set @Bonif= 0                                              
   End                                              
     Else                                              
   Begin             
	IF LEFT(@vPla_cCuentaContable, 2) = '12'
	BEGIN
		If @vAsd_cTipoDocRef<>'' Set @RefTd= @vAsd_cTipoDocRef                                              
		If @vAsd_cSerieDocRef<>'' Set @RefSr= @vAsd_cSerieDocRef                                              
		If @vAsd_cNumDocRef<>'' Set @RefNm= @vAsd_cNumDocRef                                              
		If @vAsd_dFecDocRef<>'' Set @Reffec= @vAsd_dFecDocRef 
	END  
	                               
          
    If @Asd_cImpAdic <>0 Set @ValEmb= @Asd_cImpAdic  
                          
    Set @FecEmi= @vAsd_dFecDoc                                              
                                       
   if @vAsd_cTipoDoc<>''        -- ANTES ESTABA ASI  if @vAsd_cTipoDoc<>'99' and @vAsd_cTipoDoc<>''                                              
 begin                                              
    Set @CmpTd= @vAsd_cTipoDoc                                              
    Set @CmpSr= @vAsd_cSerieDoc                                              
    Set @CmpNm= @vAsd_cNumDoc                                              
                                 
    IF LEFT(@vPla_cCuentaContable, 2) = '12'
    BEGIN
    	Set @RefTd= @vAsd_cTipoDocRef                                              
		Set @RefSr= @vAsd_cSerieDocRef                                              
		Set @RefNm= @vAsd_cNumDocRef                                              
		set @Reffec = @vAsd_dFecDocRef
    END                                    
                                       
    Set @NumVou= @vAse_nVoucher                                              
    Set @cNombreDoc=@Tdo_cNombreCortoT                                              
                                           
    Set @NumRuc= @vEnt_nRucT                                              
    Set @NomClie= @vEnt_cNombre                                              
                                          
    if @vEnt_cCodEntidad <>''                                              
     begin                                              
     Set @TipoEntidad= @vTen_cTipoEntidad                                              
     Set @CodEntidad= @vEnt_cCodEntidad                                              
     end                                              
           
  end                                              
                                   
 --- Tipo de Documento PARA SABER SI ES NOTA DE CREDITO                                              
 If @vAsd_cTipoDoc=@cTipDocNC                                              
     --Set @DocNC='1' --                                           
                                
     Set @DocNC='0'                                              
 Else                                              
     Set @DocNC='0'                                              
                                    
          --- Base Imponible PARA SABER SI PERTENECE A REG VENTAS                                          
          Print 'BASE: '  + str(@vAsd_cBaseImp)                                    
                                              
  If @vAsd_cBaseImp='002' and @vAsd_cBaseImp<>'047' and left(@vPla_cCuentaContable, 2 ) <>'40'                                              
     Begin                                               
       If @Mon_cMNac='1'                                              
          Set @BaseImpo= (Case  When @DocNC='0' Then @BaseImpo+(@vAsd_nHaberSoles - @vAsd_nDebeSoles) Else  @BaseImpo+(@vAsd_nDebeSoles-@vAsd_nHaberSoles )  End)                                              
       Else                                               
          Set @BaseImpo= (Case  When @DocNC='0' Then @BaseImpo+(@vAsd_nHaberMonExt - @vAsd_nDebeMonExt) Else  @BaseImpo+(@vAsd_nDebeMonExt-@vAsd_nHaberMonExt )  End)                                              
     End                                              
                                      
          --- Inafecto                         
    if ( @vAsd_cBaseImp='999' OR @vAsd_cBaseImp='997' ) --AND @vAsd_nMontoInafecto = 1                                               
       BEGIN                                              
               If @Mon_cMNac='1'                                                              
                  Set @BaseInaf= (Case  When @DocNC='0' Then @BaseInaf+(@vAsd_nHaberSoles - @vAsd_nDebeSoles) Else  @BaseInaf+(@vAsd_nDebeSoles-@vAsd_nHaberSoles )  End)                                              
               Else                                              
                  Set @BaseInaf= (Case  When @DocNC='0' Then @BaseInaf+(@vAsd_nHaberMonExt - @vAsd_nDebeMonExt) Else  @BaseInaf+(@vAsd_nDebeMonExt-@vAsd_nHaberMonExt )  End)                                              
       END                                              
                                              
          --- EXONERADA                                              
    if @vAsd_cBaseImp='998' --AND @vAsd_nMontoInafecto = 1                                               
       BEGIN                                     
               If @Mon_cMNac='1'                                                              
                  Set @Exon = (Case  When @DocNC='0' Then @Exon+(@vAsd_nHaberSoles - @vAsd_nDebeSoles) Else  @Exon+(@vAsd_nDebeSoles-@vAsd_nHaberSoles )  End)                                              
               Else                                              
                  Set @Exon = (Case  When @DocNC='0' Then @Exon+(@vAsd_nHaberMonExt - @vAsd_nDebeMonExt) Else  @Exon+(@vAsd_nDebeMonExt-@vAsd_nHaberMonExt )  End)                                              
       END                                         
                                              
          --- ISC                                              
     if @vAsd_cBaseImp='017' --If dbo.fConfigRegVta(@Emp_cCodigo,'017',@vPla_cCuentaContable,@Pan_cAnio)='1'                                              
             Begin                                               
               If @Mon_cMNac='1'                                                              
                  Set @ImpISC= (Case  When @DocNC='0' Then @ImpISC+(@vAsd_nHaberSoles - @vAsd_nDebeSoles) Else  @ImpISC+(@vAsd_nDebeSoles-@vAsd_nHaberSoles )  End)                                              
               Else                                              
                  Set @ImpISC= (Case  When @DocNC='0' Then @ImpISC+(@vAsd_nHaberMonExt - @vAsd_nDebeMonExt) Else  @ImpISC+(@vAsd_nDebeMonExt-@vAsd_nHaberMonExt )  End)                                              
             End                                              
                                              
      --- EXPORTACION                                              
     if @vAsd_cBaseImp='021'                                              
             Begin                                               
               If @Mon_cMNac='1'                                                              
                  Set @MontoExp= (Case  When @DocNC='0' Then @MontoExp+(@vAsd_nHaberSoles - @vAsd_nDebeSoles) Else  @MontoExp+(@vAsd_nDebeSoles-@vAsd_nHaberSoles )  End)                                              
               Else                                              
                  Set @MontoExp= (Case  When @DocNC='0' Then @MontoExp+(@vAsd_nHaberMonExt - @vAsd_nDebeMonExt) Else  @MontoExp+(@vAsd_nDebeMonExt-@vAsd_nHaberMonExt )  End)                                              
             End                                              
                                              
                               
      --- BONIFICACIONES                                              
       If @vAsd_cBaseImp='047' --dbo.fConfigRegVta(@Emp_cCodigo,'047',@vPla_cCuentaContable,@Pan_cAnio)='1'                                               
             Begin                                               
               If @Mon_cMNac='1'                                        
                  Set @Bonif= (Case  When @DocNC='0' Then @Bonif+(@vAsd_nHaberSoles - @vAsd_nDebeSoles) Else  @Bonif+(@vAsd_nDebeSoles-@vAsd_nHaberSoles )  End)                                              
               Else                                              
                  Set @Bonif= (Case  When @DocNC='0' Then @Bonif+(@vAsd_nHaberMonExt - @vAsd_nDebeMonExt) Else  @Bonif+(@vAsd_nDebeMonExt-@vAsd_nHaberMonExt )  End)                                              
             End                                              
                                                 
          --- IGV                                
                  
                  --SET @TC_CTA40 =  @vAsd_nTipoCambio  
		  IF @vAsd_cGlosa <> 'DIFERENCIA POR TIPO DE CAMBIO'
		  BEGIN
		  		IF ABS(@vAsd_nDebeSoles) > 0 AND ABS(@vAsd_nDebeMonExt) > 0
				 BEGIN
					 SET @TC_CTA40 = ABS(@vAsd_nDebeSoles) / ABS(@vAsd_nDebeMonExt) -- @vAsd_nTipoCambio
				 END
				 
			  ELSE IF ABS(@vAsd_nHaberSoles) > 0 AND ABS(@vAsd_nHaberMonExt) > 0
				 BEGIN
					 SET @TC_CTA40 = ABS(@vAsd_nHaberSoles) / ABS(@vAsd_nHaberMonExt)
				 END
			  ELSE
				BEGIN
					SET @TC_CTA40 = @vAsd_nTipoCambio
				END
		  END
		  
				  
				  --SELECT @vAsd_nHaberSoles, @vAsd_nDebeSoles, @vAsd_nHaberMonExt, @vAsd_nDebeMonExt, @vAsd_nTipoCambio, @Mon_cMNac
				  
          If left (@vPla_cCuentaContable,2)='40'   and -- es igv 40                               
 @vAsd_cBaseImp<>'017'-- no es isc 40  --dbo.fConfigRegVta(@Emp_cCodigo,'017',@vPla_cCuentaContable,@Pan_cAnio)='0' -- no es isc 40                                              
             Begin                                               
    SET @TC_CTA40 =  @vAsd_nTipoCambio                                
                           
                              
                      
  select @Numreg = COUNT(*) from CND_CONFIG_OPERA                    
  where Emp_cCodigo = @Emp_cCodigo and Pan_cAnio = @Pan_cAnio and Cod_cValorParam = @vPla_cCuentaContable                    
                       
   If @Mon_cMNac='1' and @Numreg <> 0                           
    begin                    
                  Set @ImpIGV= (Case  When @DocNC='0' Then @ImpIGV+(@vAsd_nHaberSoles - @vAsd_nDebeSoles) Else  @ImpIGV+(@vAsd_nDebeSoles-@vAsd_nHaberSoles )  End)                                              
                end                    
             Else                    
    begin                    
    if @Numreg <> 0                    
                  Set @ImpIGV= (Case  When @DocNC='0' Then @ImpIGV+(@vAsd_nHaberMonExt - @vAsd_nDebeMonExt) Else  @ImpIGV+(@vAsd_nDebeMonExt-@vAsd_nHaberMonExt )  End)                                              
                end                    
             End  
                                              
          --- FOB                                       
          If dbo.fConfigRegVta(@Emp_cCodigo,'016',@vPla_cCuentaContable,@Pan_cAnio)='1'                                              
             Begin                                              
               If @Mon_cMNac='1'                                                              
     Set @ImpFOB=@ImpFOB+(@vAsd_nHaberSoles-@vAsd_nDebeSoles)                                              
               Else                                              
     Set @ImpFOB=@ImpFOB+(@vAsd_nHaberMonExt-@vAsd_nDebeMonExt)                                              
             End                                                
                                              
                                              
          --- Flete                                              
          If dbo.fConfigRegVta(@Emp_cCodigo,'014',@vPla_cCuentaContable,@Pan_cAnio)='1'                                              
             Begin                                              
               If @Mon_cMNac='1'                                              
           Set @ImpFlete=@ImpFlete+(@vAsd_nHaberSoles-@vAsd_nDebeSoles)                                              
               Else                                              
              Set @ImpFlete=@ImpFlete+(@vAsd_nHaberMonExt-@vAsd_nDebeMonExt)                                              
             End                                                    
                                              
       --- Redondeo                                              
          If (@vPla_cCuentaContable = @CTA_REDGAN or @vPla_cCuentaContable = @CTA_REDPER  OR                                              
    LEFT(@vPla_cCuentaContable ,2) IN ('41','42','45','46','47') OR                         
    dbo.fConfigRegVta(@Emp_cCodigo,'003',@vPla_cCuentaContable,@Pan_cAnio)='1' ) -- base de honorario y otros                                             
       AND @vAsd_cBaseImp<>'047'  -- no es bonif y transf gratuita                                                              
             Begin                                              
                If @Mon_cMNac='1'   
                 Set @ImpOtros=@ImpOtros+(@vAsd_nHaberSoles-@vAsd_nDebeSoles)                                                
                Else                                               
  Set @ImpOtros= @ImpOtros+(@vAsd_nHaberMonExt-@vAsd_nDebeMonExt)                                               
   End                                             
                                             
          --- Base Imp Otros Impuestos                                              
          Set @BaseIpOIm=0                                              
                                              
          --- Diferencia de Cambio                                              
          If @vPla_cCuentaContable =@CTA_DIFGAN  or @vPla_cCuentaContable =@CTA_DIFPER                      
             Begin                                              
                                              
             
               If @Mon_cMNac='1'                                              
					Set @ImpDifCmb=@ImpDifCmb+(@vAsd_nHaberSoles-@vAsd_nDebeSoles)                         
             
               Else                                                
					Set @ImpDifCmb=@ImpDifCmb+(@vAsd_nHaberMonExt-@vAsd_nDebeMonExt)                                              
               End                                              
        End      
      
     -----------------------------------------------------------------------------------------------                                              
   Fetch Next From Cursor_Detalle                                               
   Into @vAsd_nItem, @vPla_cCuentaContable, @vAsd_cGlosa, @vAsd_nDebeSoles, @vAsd_nHaberSoles,                                              
         @vAsd_nTipoCambio, @vAsd_nDebeMonExt, @vAsd_nHaberMonExt, @vCos_cCodigo, @vTen_cTipoEntidad,                                              
        @vEnt_cCodEntidad, @vAsd_cTipoDoc, @vAsd_cSerieDoc, @vAsd_cNumDoc, @vAsd_dFecDoc,                                              
         @vAsd_cTipoDocRef, @vAsd_cSerieDocRef, @vAsd_cNumDocRef, @vAsd_dFecDocRef,                                      
         @vAsd_nMontoInafecto, @vAsd_cTipoMoneda, @vEnt_nRucT, @vEnt_cNombre, @Tdo_cNombreCortoT,          
  @vAsd_cBaseImp, @vAsd_dFecVen,@Asd_cImpAdic                                             
  End          
  -----------------------------------------------------------------------------------------------                         
  Set @vItem=@vItem+1                                              
  --- Total                                              
  Set @ImpOtros= isnull(@ImpOtros,0)+ isnull(@Bonif,0) --+ isnull(@ImpDifCmb,0)                                              
  
  IF ISNULL(@Bonif, 0) = 0
  BEGIN
  	Set @ImpOtros = @ImpDifCmb
  END                     
                            
--  set @ImpDifCmb = 0                                               
                                              
  Set @ImpTotal= @BaseImpo+@BaseInaf+@ImpIGV+@ImpFOB+@ImpFlete+@ImpOtros+@MontoExp+@ImpISC- @Bonif +  @Exon                                              
                                           
  set @vAsd_nTipoCambio = isnull(@vAsd_nTipoCambio , 0)                                          

-----------------------------------------------------------------------------------------------  
                    
  Insert into #TMPREGISTROLEVENTAS with(rowlock)                                              
         (Emp_cCodigo,Ase_cNumMov,Per_cPeriodo,Asd_dFecDoc,Asd_cTipoDoc,Asd_cSerieDoc,Asd_cNumDoc, Ase_dFecha,                    
          Asd_cTipoDocRef,Asd_cSerieDocRef,Asd_cNumDocRef,Ase_nVoucher,Ent_nRuc,Ent_cPersona,Asd_nItem,MontoBaseGra,                                              
          MontoInafecto,MontoBaseOtImp,MontoISC,MontoIGV,MontoFOB,MontoFlete,MontoOtros,MontoDifCmb,MontoTotal,cNombreDoc,                                     
      MontoExp, Asd_dFecVen , Asd_dFecDocRef , Asd_cCodSunat , MontoExon, Ent_cCodEntidad, Ten_cTipoEntidad, Asd_nTipoCambio,Asd_cEstadoO, Asd_cEstadoD, Asd_cImpAdic )                                              
  Values  
      (@vEmp_cCodigo,@vAse_cNummov,@vPer_cPeriodo,@FecEmi,isnull(@CmpTd,''),isnull(@CmpSr,''),isnull(@CmpNm,''), @vAse_dFecha,                                             
      isnull(@RefTd,''),isnull(@RefSr,''),isnull(@RefNm,''),isnull(@NumVou,''),isnull(@NumRuc,''),isnull(@NomClie,''),isnull(@vItem,1),isnull(@BaseImpo,''),                                              
      @BaseInaf,@BaseIpOIm,@ImpISC,@ImpIGV,@ImpFOB,@ImpFlete,@ImpOtros, @ImpDifCmb,@ImpTotal,isnull(@cNombreDoc,''),                                               
      @MontoExp, isnull(@vAsd_dFecVen,'') , isnull(@Reffec,''), '' , @Exon , isnull(@CodEntidad,''), isnull(@TipoEntidad ,''), isnull(@TC_CTA40,''),@Asd_cEstadoO, @Asd_cEstadoD, @ValEmb ) --@vAsd_nTipoCambio )                                         
  
-----------------------------------------------------------------------------------------------                              
                    
 -- print  'Ase_cNumMov -->' + @vAse_cNummov + ' , Ase_nVoucher->' + @NumVou             
           
          
  Close Cursor_Detalle                 
  Deallocate Cursor_Detalle                                              
  Fetch Next From Cursor_Cabecera Into @vAse_cNummov,@vEmp_cCodigo,@vPan_cAnio,@vPer_cPeriodo,@vLib_cTipoLibro,                                              
                                       @vAse_nVoucher,@vAse_dFecha,@vAse_cGlosa,@vAse_cTipoMoneda,@vAse_nTipoCambio,                                              
                                       @vAse_cEstado,@Asd_cEstadoO, @Asd_cEstadoD                              
End              
          
                  
-----------------------------------------------------------------------------------------------                                              
Close Cursor_Cabecera                                              
Deallocate Cursor_Cabecera     
    
-----------------------------------------------------------------------------------------------                                              
 --   SELECT * FROM #TMPREGISTROLEVENTAS REP LEFT JOIN dbo.CNM_ENTIDAD CE
	--ON REP.Emp_cCodigo = CE.Emp_cCodigo AND CE.Ten_cTipoEntidad = REP.Ten_cTipoEntidad AND CE.Ent_cCodEntidad = REP.Ent_cCodEntidad
	--INNER JOIN dbo.CNC_ASIENTO_VOUCHER CAV ON REP.Emp_cCodigo = CAV.Emp_cCodigo AND CAV.Ase_nVoucher = REP.Ase_nVoucher
	--INNER JOIN dbo.CNT_TIPO_MONEDA CTM ON cav.Emp_cCodigo = CTM.Emp_cCodigo AND cav.Ase_cTipoMoneda = CTM.Mon_cCodigo
	--WHERE LEFT(REP.Asd_cTipoDoc,2) NOT IN ('02','09','10','20','22','31','33','40','41','46','50','51','52','53','54','89','91','96',
	--																			   '97','98')
	--Order by convert(varchar(4),year(REP.Ase_dFecha)) + REP.Per_cPeriodo + '00',REP.Ase_nVoucher  
	
	--SELECT * FROM #TMPREGISTROLEVENTAS REP INNER JOIN dbo.CNM_ENTIDAD CE
	--ON REP.Emp_cCodigo = CE.Emp_cCodigo AND REP.Ten_cTipoEntidad = CE.Ten_cTipoEntidad AND REP.Ent_cCodEntidad = CE.Ent_cCodEntidad
	--ORDER BY rep.ent_ccodentidad      
    --RETURN                
        
	--	SELECT * FROM #TMPREGISTROLEVENTAS REP
	--	 LEFT JOIN dbo.CNM_ENTIDAD CE
	--ON REP.Emp_cCodigo = CE.Emp_cCodigo AND CE.Ten_cTipoEntidad = REP.Ten_cTipoEntidad AND CE.Ent_cCodEntidad = REP.Ent_cCodEntidad
	--LEFT JOIN dbo.CNC_ASIENTO_VOUCHER CAV ON REP.Emp_cCodigo = CAV.Emp_cCodigo AND CAV.Ase_nVoucher = REP.Ase_nVoucher AND cav.Ase_cNummov = REP.Ase_cNummov
	--LEFT JOIN dbo.CNT_TIPO_MONEDA CTM ON cav.Emp_cCodigo = CTM.Emp_cCodigo AND cav.Ase_cTipoMoneda = CTM.Mon_cCodigo
	--WHERE 
	--LEFT(REP.Asd_cTipoDoc,2) NOT IN ('02','09','10','20','22','31','33','40','41','46','50','51','52','53','54','89','91','96',
	--																			   '97','98') 
	--AND 
	--																			   cav.Pan_cAnio = @Pan_cAnio --AND RTRIM(LTRIM(CE.Ent_cDeleted)) <> '*'
	--Order by convert(varchar(4),year(REP.Ase_dFecha)) + REP.Per_cPeriodo + '00',REP.Ase_nVoucher
	--SELECT * FROM #TMPREGISTROLEVENTAS
	--RETURN
        
DECLARE @Separador varchar(1)                        
                        
SELECT @Separador = '|'  

IF @Simplificado = '0'
BEGIN
	SELECT (CASE WHEN (REP.Asd_cEstadoO = '8' OR REP.Asd_cEstadod = '8') THEN CAST(YEAR(REP.Asd_dFecDoc) AS CHAR(4)) ELSE convert(varchar(4),year(REP.Ase_dFecha)) END + CASE WHEN (REP.Asd_cEstadoO = '8' OR REP.Asd_cEstadod = '8') THEN RIGHT(100 + DATEPART(MM,REP.Asd_dFecDoc), 2) ELSE REP.Per_cPeriodo END + '00' + @Separador + 
	        LTRIM(RTRIM(LEFT(REP.Ase_nVoucher,40))) + @Separador +
		    LEFT(LTRIM(RTRIM('M' + left(REP.Ase_nVoucher,4) + right(REP.Ase_nVoucher,5) )),100) + @Separador +
		    convert(varchar(10),REP.Asd_dFecDoc,103) + @Separador + CASE WHEN YEAR(REP.Asd_dFecVen) = 1900 THEN '' ELSE convert(varchar(10),REP.Asd_dFecVen,103) END + @Separador +
		    LEFT(REP.Asd_cTipoDoc,2) + @Separador + RTRIM(REP.Asd_cSerieDoc) + @Separador + 
		   
		   CASE WHEN (CHARINDEX('-', REP.Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(REP.Asd_cNumDoc, 0, CHARINDEX('-', REP.Asd_cNumDoc)), 7) 
		        WHEN (CHARINDEX('/', REP.Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(REP.Asd_cNumDoc, 0, CHARINDEX('/', REP.Asd_cNumDoc)), 7) 
		        WHEN LEFT(REP.Asd_cTipoDoc,2) IN ('01', '03', '04', '07', '08') THEN RIGHT(RTRIM(LTRIM(REP.Asd_cNumDoc)), 7)
		        WHEN LEFT(REP.Asd_cTipoDoc,2) IN ('00', '13', '14', '15', '16', '17', '18') THEN RIGHT(RTRIM(LTRIM(REP.Asd_cNumDoc)), 20)
		   ELSE RIGHT(RTRIM(REP.Asd_cNumDoc), 7) END + @Separador + 
		   
		   CASE WHEN (CHARINDEX('-', REP.Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(REP.Asd_cNumDoc, CHARINDEX('-', REP.Asd_cNumDoc) + 1, LEN(REP.Asd_cNumDoc) - CHARINDEX('-', REP.Asd_cNumDoc)), 7) 
		        WHEN (CHARINDEX('/', REP.Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(REP.Asd_cNumDoc, CHARINDEX('/', REP.Asd_cNumDoc) + 1, LEN(REP.Asd_cNumDoc) - CHARINDEX('/', REP.Asd_cNumDoc)), 7)
		   ELSE '' END + @Separador 
		   
		   + (SELECT t.Tab_cCodSunat FROM dbo.TABLA T WHERE T.Emp_cCodigo = REP.Emp_cCodigo AND T.Tab_cTabla = '003' AND T.Tab_cCodigo = CE.Ent_cTipoDoc) + @Separador +
		   RTRIM(LTRIM(CE.Ent_nRuc)) + @Separador + RTRIM(LTRIM(CE.Ent_cPersona)) + @Separador + 
		   CAST(CAST(REP.MontoExp AS NUMERIC(18, 2)) AS VARCHAR(50)) + @Separador + CAST(CAST(REP.MontoBaseGra AS NUMERIC(18, 2)) AS VARCHAR(50)) + @Separador + '0.00' + @Separador +
		   CAST(CAST(REP.MontoIGV AS NUMERIC(18, 2)) AS VARCHAR(50)) + @Separador + '0.00' + @Separador +
		   CAST(CAST(MontoExon AS NUMERIC(18, 2)) AS VARCHAR(50)) + @Separador + CAST(CAST(REP.MontoInafecto AS NUMERIC(18, 2)) AS VARCHAR(50)) + @Separador + 
		   CAST(CAST(REP.MontoISC AS NUMERIC(18, 2)) AS VARCHAR(50)) + @Separador + '0.00' + @Separador + '0.00' + @Separador + 
		   CAST(CAST(REP.MontoOtros AS NUMERIC(18, 2)) AS VARCHAR(50)) + @Separador + 
		   CAST(CAST(REP.MontoTotal AS NUMERIC(18, 2)) AS VARCHAR(50)) + @Separador +
		   ctm.Mon_cCodSunat + @Separador + CASE WHEN ctm.Mon_cCodSunat = 'PEN' THEN '1.000' ELSE CAST(CAST(REP.Asd_nTipoCambio AS NUMERIC(4, 3)) AS VARCHAR(50)) END + @Separador + CASE  WHEN LEFT(REP.Asd_cTipoDoc,2) IN ('07', '08', '87', '88') THEN convert(varchar(10),REP.Asd_dFecDoc,103) ELSE '' END + @Separador +
		   CASE  WHEN LEFT(REP.Asd_cTipoDoc,2) IN ('07', '08', '87', '88') THEN RTRIM(LTRIM(LEFT(REP.Asd_cTipoDocRef,2))) ELSE '' END + @Separador +
		   CASE  WHEN LEFT(REP.Asd_cTipoDoc,2) IN ('07', '08', '87', '88') THEN RIGHT(RTRIM(REP.Asd_cSerieDocRef), 4) ELSE '' END + @Separador +
		   CASE  WHEN LEFT(REP.Asd_cTipoDoc,2) IN ('07', '08', '87', '88') THEN RIGHT(RTRIM(REP.Asd_cNumDocRef), 7) ELSE '' END + @Separador + '' + @Separador + '' + @Separador + '' + @Separador +
		   LTRIM(RTRIM(LEFT(case when REP.Asd_cEstadoD = '' then REP.Asd_cEstadoO else  REP.Asd_cEstadoD end,1))) + @Separador) AS Registro
	FROM #TMPREGISTROLEVENTAS REP LEFT JOIN dbo.CNM_ENTIDAD CE
	ON REP.Emp_cCodigo = CE.Emp_cCodigo AND CE.Ten_cTipoEntidad = REP.Ten_cTipoEntidad AND CE.Ent_cCodEntidad = REP.Ent_cCodEntidad
	LEFT JOIN dbo.CNC_ASIENTO_VOUCHER CAV ON REP.Emp_cCodigo = CAV.Emp_cCodigo AND CAV.Ase_nVoucher = REP.Ase_nVoucher AND cav.Ase_cNummov = REP.Ase_cNummov
	LEFT JOIN dbo.CNT_TIPO_MONEDA CTM ON cav.Emp_cCodigo = CTM.Emp_cCodigo AND cav.Ase_cTipoMoneda = CTM.Mon_cCodigo
	WHERE LEFT(REP.Asd_cTipoDoc,2) NOT IN ('02','09','10','20','22','31','33','40','41','46','50','51','52','53','54','89','91','96',
																				   '97','98') AND cav.Pan_cAnio = @Pan_cAnio --AND CE.Ent_cDeleted <> '*'
	Order by convert(varchar(4),year(REP.Ase_dFecha)) + REP.Per_cPeriodo + '00',REP.Ase_nVoucher
END


IF @Simplificado = '1'
BEGIN
	SELECT (CASE WHEN (REP.Asd_cEstadoO = '8' OR REP.Asd_cEstadod = '8') THEN CAST(YEAR(REP.Asd_dFecDoc) AS CHAR(4)) ELSE convert(varchar(4),year(REP.Ase_dFecha)) END + CASE WHEN (REP.Asd_cEstadoO = '8' OR REP.Asd_cEstadod = '8') THEN RIGHT(100 + DATEPART(MM,REP.Asd_dFecDoc), 2) ELSE REP.Per_cPeriodo END + '00' + @Separador + LTRIM(RTRIM(LEFT(REP.Ase_nVoucher,40))) + @Separador + LEFT(LTRIM(RTRIM('M' + left(REP.Ase_nVoucher,4) + right(REP.Ase_nVoucher,5) )),100) + @Separador + convert(varchar(10),REP.Asd_dFecDoc,103) + @Separador + CASE WHEN YEAR(REP.Asd_dFecVen) = 1900 THEN '' ELSE convert(varchar(10),REP.Asd_dFecVen,103) END + @Separador +
	LEFT(REP.Asd_cTipoDoc,2) + @Separador + RTRIM(REP.Asd_cSerieDoc) + @Separador + 
	
	CASE WHEN (CHARINDEX('-', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('-', Asd_cNumDoc)), 7) 
		 WHEN (CHARINDEX('/', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('-', Asd_cNumDoc)), 7) 
		 ELSE RIGHT(RTRIM(Asd_cNumDoc), 7) END + @Separador + 
	
	CASE WHEN (CHARINDEX('-', Asd_cNumDoc) > 0) AND LEFT(REP.Asd_cTipoDoc,2) IN ('00', '03', '12') THEN RIGHT(SUBSTRING(Asd_cNumDoc, CHARINDEX('-', Asd_cNumDoc) + 1, LEN(Asd_cNumDoc) - CHARINDEX('-', Asd_cNumDoc)), 7) 
		 WHEN (CHARINDEX('/', Asd_cNumDoc) > 0) AND LEFT(REP.Asd_cTipoDoc,2) IN ('00', '03', '12') THEN RIGHT(SUBSTRING(Asd_cNumDoc, CHARINDEX('/', Asd_cNumDoc) + 1, LEN(Asd_cNumDoc) - CHARINDEX('/', Asd_cNumDoc)), 7)
		   ELSE '' END + 
	
	@Separador + (SELECT t.Tab_cCodSunat FROM dbo.TABLA T WHERE T.Emp_cCodigo = REP.Emp_cCodigo AND T.Tab_cTabla = '003' AND T.Tab_cCodigo = CE.Ent_cTipoDoc) + @Separador +
	RTRIM(LTRIM(CE.Ent_nRuc)) + @Separador + RTRIM(LTRIM(CE.Ent_cPersona)) + @Separador + CAST(CAST(REP.MontoBaseGra AS NUMERIC(18, 2)) AS VARCHAR(50)) + @Separador + CAST(CAST(REP.MontoIGV AS NUMERIC(18, 2)) AS VARCHAR(50)) + @Separador + CAST(CAST(REP.MontoOtros AS NUMERIC(18, 2)) AS VARCHAR(50)) + @Separador + CAST(CAST(REP.MontoTotal AS NUMERIC(18, 2)) AS VARCHAR(50)) + @Separador + ctm.Mon_cCodSunat + @Separador +
	CASE WHEN ctm.Mon_cCodSunat = 'PEN' THEN '1.000' ELSE CAST(CAST(REP.Asd_nTipoCambio AS NUMERIC(4, 3)) AS VARCHAR(50)) END + @Separador + CASE WHEN LEFT(REP.Asd_cTipoDoc,2) IN ('07', '08') THEN convert(varchar(10),REP.Asd_dFecDoc,103) ELSE '' END + @Separador + CASE WHEN LEFT(REP.Asd_cTipoDoc,2) IN ('07', '08') THEN LEFT(REP.Asd_cTipoDocRef,2) ELSE '' END + @Separador +
	CASE WHEN LEFT(REP.Asd_cTipoDoc,2) IN ('07', '08') THEN RTRIM(REP.Asd_cSerieDocRef) ELSE '' END + @Separador + CASE WHEN LEFT(REP.Asd_cTipoDoc,2) IN ('07', '08') THEN RIGHT(RTRIM(REP.Asd_cNumDocRef), 7) ELSE '' END + @Separador + '' + @Separador + '' + @Separador + LTRIM(RTRIM(LEFT(case when REP.Asd_cEstadoD = '' then REP.Asd_cEstadoO else  REP.Asd_cEstadoD end,1))) + @Separador ) AS Registro FROM #TMPREGISTROLEVENTAS REP LEFT JOIN dbo.CNM_ENTIDAD CE
	ON REP.Emp_cCodigo = CE.Emp_cCodigo AND CE.Ten_cTipoEntidad = REP.Ten_cTipoEntidad AND CE.Ent_cCodEntidad = REP.Ent_cCodEntidad
	LEFT JOIN dbo.CNC_ASIENTO_VOUCHER CAV ON REP.Emp_cCodigo = CAV.Emp_cCodigo AND CAV.Ase_nVoucher = REP.Ase_nVoucher AND cav.Ase_cNummov = REP.Ase_cNummov
	LEFT JOIN dbo.CNT_TIPO_MONEDA CTM ON cav.Emp_cCodigo = CTM.Emp_cCodigo AND cav.Ase_cTipoMoneda = CTM.Mon_cCodigo
	WHERE LEFT(REP.Asd_cTipoDoc,2) NOT IN ('02','04','05','06','09','10','11','13','14','15','16','17','18','19','20','21','22','23','24','25','26','27','28','29','30',
														'31', '32', '33', '34', '35','36','37', '40', '41', '42', '43', '44', '45', '46', '48', '49', '50', '51', '52', '53', '54', '55',
														'56' , '87', '88' , '89', '91', '96', '97', '98') AND cav.Pan_cAnio = @Pan_cAnio
	Order by convert(varchar(4),year(REP.Ase_dFecha)) + REP.Per_cPeriodo + '00',REP.Ase_nVoucher
END

--RETURN 

--Select                      
--(convert(varchar(4),year(REP.Ase_dFecha)) + REP.Per_cPeriodo + '00' + @Separador +                    
--LTRIM(RTRIM(LEFT(REP.Ase_nVoucher,40))) + @Separador +       
-- LEFT(LTRIM(RTRIM('M' + left(REP.Ase_nVoucher,4) + right(REP.Ase_nVoucher,5) )),100) + @Separador +       
--LTRIM(RTRIM(LEFT(case when CNM_ENTIDAD.Ent_cPersona = 'ANULADO' then convert(varchar(10),REP.Asd_dFecDoc,103) else convert(varchar(10),REP.Asd_dFecDoc,103)end,10))) + @Separador +                    
--LTRIM(RTRIM(LEFT(case when CNM_ENTIDAD.Ent_cPersona <> 'ANULADO' then                    
--case when REP.Asd_dFecVen = '' then '' else  convert(varchar(10),REP.Asd_dFecVen,103) end else '' end,10))) + @Separador +                    
--LTRIM(RTRIM(LEFT(REP.Asd_cTipoDoc,2))) + @Separador +  
-- case when RTRIM(REP.Asd_cTipoDoc) in ('05','55') then
--	right(RTRIM(REP.Asd_cSerieDoc),1)  
-- else 
--	LTRIM(RTRIM(LEFT(REP.Asd_cSerieDoc,4)))
-- end + @Separador +
--case when CHARINDEX('/', REP.Asd_cNumDoc) = 0 then                   
--case when CHARINDEX('-', REP.Asd_cNumDoc) <> 0 then
-- left(REP.Asd_cNumDoc,charindex('-',REP.Asd_cNumDoc)-1)                  
--else     
--case when REP.Asd_cTipoDoc in ('01','03','04','06','07','08','23') then    
-- right(RTRIM(REP.Asd_cNumDoc ),7)     
-- else RTRIM(REP.Asd_cNumDoc ) end     
--end                   
--else          
--case when REP.Asd_cTipoDoc in ('01','03','04','06','07','08','23') then    
-- right(RTRIM(left(REP.Asd_cNumDoc,charindex('/',REP.Asd_cNumDoc)-1) ),7)     
-- else RTRIM(left(REP.Asd_cNumDoc,charindex('/',REP.Asd_cNumDoc)-1)) end          
--end + @Separador +                   
--LTRIM(RTRIM(LEFT(case when CHARINDEX('/', REP.Asd_cNumDoc) = 0 then                   
--case when CHARINDEX('-', REP.Asd_cNumDoc) <> 0 then        
-- Ltrim(SUBSTRING(REP.Asd_cNumDoc,charindex('-',REP.Asd_cNumDoc)+1,20))                  
--else '0' end                   
--else                  
-- Ltrim(SUBSTRING(REP.Asd_cNumDoc,charindex('/',REP.Asd_cNumDoc)+1,20))                  
--end,20)))  + @Separador +                   
--LTRIM(RTRIM(LEFT(case when CNM_ENTIDAD.Ent_cPersona in ('ANULADO','ANULADA','CLIENTES VARIOS') then '0' else TABLA.Tab_cCodSunat end,1))) + @Separador +                    
--LTRIM(RTRIM(LEFT(case when CNM_ENTIDAD.Ent_cPersona in ('ANULADO','ANULADA','CLIENTES VARIOS') then '-' else CNM_ENTIDAD.Ent_nRuc end,15))) + @Separador +                    
--LTRIM(RTRIM(LEFT(replace(replace(CNM_ENTIDAD.Ent_cPersona, char(13)+char(10), ' '), char(9), ' '),60))) + @Separador +                    
--LTRIM(RTRIM(LEFT(CONVERT(money,REP.MontoExp, 0),12))) + @Separador +            
--LTRIM(RTRIM(LEFT(CONVERT(money,REP.MontoBaseGra, 0),12))) + @Separador +                    
--LTRIM(RTRIM(LEFT(CONVERT(money,REP.MontoExon, 0),12))) + @Separador +                    
--LTRIM(RTRIM(LEFT(CONVERT(money,REP.MontoInafecto, 0),12))) + @Separador +                    
--LTRIM(RTRIM(LEFT(CONVERT(money,REP.MontoISC, 0),12))) + @Separador +                    
--LTRIM(RTRIM(LEFT(CONVERT(money,REP.MontoIGV, 0),12))) + @Separador +                    
--'0.00'  + @Separador +                    
--'0.00'  + @Separador +                    
--LTRIM(RTRIM(LEFT(CONVERT(money,REP.MontoOtros, 0),12))) + @Separador +                    
--LTRIM(RTRIM(LEFT(CONVERT(money,REP.MontoTotal, 0),12))) + @Separador +                    
--LTRIM(RTRIM(LEFT(convert(varchar(5),abs(REP.Asd_nTipoCambio)),5))) + @Separador +                    
--LTRIM(RTRIM(LEFT(case when REP.Asd_cTipoDoc IN ('07','08','87','88','97','98') then                  
--case when CNM_ENTIDAD.Ent_cPersona in ('ANULADO','ANULADA') then '01/01/0001' else convert(varchar(10),REP.Asd_dFecDocRef,103)end else '01/01/0001' end,10))) + @Separador +                    
--LTRIM(RTRIM(LEFT(case when CNM_ENTIDAD.Ent_cPersona not in ('ANULADO','ANULADA') then                  
--case when REP.Asd_cTipoDoc IN ('07','08','87','88','97','98') then                   
--  REP.Asd_cTipoDocRef                   
--  else '00' end                  
--else '00' end,2))) + @Separador +                  
                  
--LTRIM(RTRIM(LEFT(case when CNM_ENTIDAD.Ent_cPersona not in ('ANULADO','ANULADA') then                  
--case when REP.Asd_cTipoDoc IN ('07','08','87','88','97','98') then                   
--  REP.Asd_cSerieDocRef                  
--  else '-' end                  
--else '-' end,20))) + @Separador +                  
--LTRIM(RTRIM(LEFT(case when CNM_ENTIDAD.Ent_cPersona not in ('ANULADO','ANULADA') then                  
--case when REP.Asd_cTipoDoc IN ('07','08','87','88','97','98') then                   
--  REP.Asd_cNumDocRef                  
--  else '-' end           
--else '-' end,20))) + @Separador +      
-- LTRIM(RTRIM(LEFT(CONVERT(money,REP.Asd_cImpAdic),12)))   + @Separador +  
----LTRIM(RTRIM(LEFT(case when CNM_ENTIDAD.Ent_cPersona <> 'ANULADO' then                    
----case when Asd_cEstadoD = '' then Asd_cEstadoO else  Asd_cEstadoD end else '2' end,1)))+ @Separador) AS 'Registro'                    
--LTRIM(RTRIM(LEFT(case when Asd_cEstadoD = '' then Asd_cEstadoO else  Asd_cEstadoD end,1)))+ @Separador +                    
--convert(varchar(15),CONVERT(money, abs(MontoDifCmb)),0)+ @Separador) AS 'Registro'    
--From   #TMPREGISTROLEVENTAS as REP          
--LEFT JOIN CNT_TIPODOC          
--ON REP.Emp_cCodigo = CNT_TIPODOC.Emp_cCodigo And          
--    REP.Asd_cTipoDoc =CNT_TIPODOC.Tdo_cCodigo          
-- LEFT JOIN CNM_ENTIDAD          
--ON REP.Emp_cCodigo = CNM_ENTIDAD.Emp_cCodigo AND          
--  REP.Ten_cTipoEntidad = CNM_ENTIDAD.Ten_cTipoEntidad AND          
--  REP.Ent_cCodEntidad = CNM_ENTIDAD.Ent_cCodEntidad          
-- LEFT JOIN TABLA          
-- ON  CNM_ENTIDAD.EMP_CCODIGO = TABLA.EMP_CCODIGO AND          
--CNM_ENTIDAD.ENT_CTIPODOC  = TABLA.TAB_CCODIGO              
--WHERE  TABLA.TAB_CTABLA = '003'      
--Order by convert(varchar(4),year(REP.Ase_dFecha)) + REP.Per_cPeriodo + '00',Ase_nVoucher
GO
