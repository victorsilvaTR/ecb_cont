USE SAFC_ECB
GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO
/*-----------------------------------------------------------------------------------------------------------------                                                      
MODULO DE CONTABILIDAD                                                      
DESCRIPCION : Reporte Electrònico de Registro de Compras                            
Fecha Mod: 01/09/2011                            
Usuario: Pool Berrospi                                                      
------------------------------------------------------------------------------------------------------------------*/                                                                                
--Exec  spCn_LibroElectronicoCompras2 '072','2014','01/01/2014','31/01/2014','038','01'                               
                            
CREATE PROCEDURE [dbo].[spCn_LibroElectronicoCompras3]                                          
 @Emp_cCodigo char(3)='',                                                        
 @Pan_cAnio char(4)='',                                                        
 @desde varchar(10) = '',                                                        
 @hasta varchar(10) = '',                                                         
 @moneda char(3) = '',                                                        
 @Per_cPeriodo char(2)='',
 @Ent_cFlagDomiciliado CHAR(1) = ''                                                        
--WITH ENCRYPTION                                                        
AS                                                        
SET DATEFORMAT DMY                                                        
SET NOCOUNT ON                                                        
                                                        
-- *** Hallando el tipo de Moneda(Si es Nac o No) y su nombre                                                        
DECLARE @Mon_cNombreLargo varchar(100)                                                        
DECLARE @Mon_cMNac char(1)                                                        
DECLARE @NombreEmpresa varchar(250)                                                        
DECLARE @Emp_cNumRuc varchar(15)                                                        
DECLARE @Mon_cCodMNac CHAR(3)                                                        
DECLARE @Registros NUMERIC(14,0)                                                        
DECLARE @ISC varchar(5)                                                        
-------------------------------------------------------------------------                                                        
-- DECLARE @Cta_nLeasing varchar(12)                                                        
-------------------------------------------------------------------------                                                        
DECLARE @nIGV numeric(14,3)                                                        
SET @nIGV = isnull ( DBO.fBuscaConfOP (@Emp_cCodigo,@Pan_cAnio,'053') , 0)                                                        
-------------------------------------------------------------------------                                                        
declare @RUC char(50)                                                        
select @RUC= dbo.RUC(@Emp_cCodigo)                                                        
-------------------------------------------------------------------------                                                        
SELECT @Mon_cNombreLargo = Mon_cNombreLargo, @Mon_cMNac = Mon_cMNac                                                         
FROM CNT_TIPO_MONEDA                                                        
WHERE Emp_cCodigo = @Emp_cCodigo and Mon_cCodigo = @moneda                                                        
                                                        
SET @Mon_cNombreLargo= ISNULL(@Mon_cNombreLargo, '')                                                         
SET @Mon_cMNac= ISNULL(@Mon_cMNac, '')                                                         
-------------------------------------------------------------------------                                   
SELECT @NombreEmpresa = Emp_cNombreLargo, @Emp_cNumRuc = Emp_cNumRuc                                                         
FROM EMPRESA WHERE Emp_cCodigo = @Emp_cCodigo                                                        
                                  
SET @NombreEmpresa= ISNULL(@NombreEmpresa, '')                                    
SET @Emp_cNumRuc= ISNULL(@Emp_cNumRuc, '')                          
-------------------------------------------------------------------------                                             
SELECT @Mon_cCodMNac=Mon_cCodigo                   
FROM CNT_TIPO_MONEDA                                                        
WHERE Emp_cCodigo = @Emp_cCodigo AND Mon_cMNac = '1'                                                     
                                                        
SET @Mon_cCodMNac= ISNULL(@Mon_cCodMNac, '')                                       
-------------------------------------------------------------------------                                                        
SELECT Cod_cValorParam AS 'CUENTAS'                                                        
into #TMP_REINTEGRO FROM CND_CONFIG_OPERA                                                         
WHERE Emp_cCodigo=@Emp_cCodigo and pan_cAnio=@Pan_cAnio and cop_ccodigo='050'   AND LEFT(Cod_cValorParam,2)='40'                                                                 
--------------------------------------------------------------------------                                                        
SELECT Cod_cValorParam AS 'CUENTAS'                                                        
into #TMP_ISC FROM CND_CONFIG_OPERA                                                         
WHERE Emp_cCodigo=@Emp_cCodigo and pan_cAnio=@Pan_cAnio and cop_ccodigo='017'   AND LEFT(Cod_cValorParam,2)='40'                                                          
-------------------------------------------------------------------------                                                        
DECLARE @LIBRO CHAR (2)                                                        
DECLARE @LIBROHON CHAR(2)                                                        
-------------------------------------------------------------------------                                                    
SELECT @LIBRO = Cfl_cCompras FROM CNT_CONFIG_LIBROS WHERE Emp_cCodigo=@Emp_cCodigo and Pan_cAnio= @Pan_cAnio                            
                                                        
SET @LIBRO= ISNULL(@LIBRO, '')                                                         
-------------------------------------------------------------------------                                                        
SELECT @LIBROHON =Cfl_cHonorarios FROM CNT_CONFIG_LIBROS WHERE Emp_cCodigo=@Emp_cCodigo  and Pan_cAnio= @Pan_cAnio                                                      
SET @LIBROHON= ISNULL(@LIBROHON, '')                                                         
-------------------------------------------------------------------------                                                        
--SELECT                      
--CNC_ASIENTO_VOUCHER.Ase_dfecha,CNC_ASIENTO_VOUCHER.Emp_cCodigo,                                                         
--CNC_ASIENTO_VOUCHER.Ase_cNumMov,CNC_ASIENTO_VOUCHER.Per_cPeriodo,                                                        
--CNC_ASIENTO_VOUCHER.Pan_cAnio,CNC_ASIENTO_VOUCHER.Lib_cTipoLibro,              
--CNC_ASIENTO_VOUCHER.Ase_nVoucher,CND_ASIENTO_VOUCHER.Asd_nItem,                                                         
--CND_ASIENTO_VOUCHER.Pla_cCuentaContable,CNM_PLAN_CTA.Pla_cNombreCuenta,                                                         
--CNM_PLAN_CTA.Pla_cProvision,CND_ASIENTO_VOUCHER.Asd_cGlosa,                                                         
--CND_ASIENTO_VOUCHER.Asd_nDebeSoles,CND_ASIENTO_VOUCHER.Asd_nHaberSoles,                             
--CND_ASIENTO_VOUCHER.Asd_nDebeMonExt,CND_ASIENTO_VOUCHER.Asd_nHaberMonExt,                
--CND_ASIENTO_VOUCHER.Asd_nTipoCambio,CNM_ENTIDAD.Ent_cCodEntidad,                                                         
--Ent_nRuc,CNM_ENTIDAD.Ent_cPersona,CND_ASIENTO_VOUCHER.Asd_cTipoDoc,                                                         
--CND_ASIENTO_VOUCHER.Asd_cSerieDoc,CND_ASIENTO_VOUCHER.Asd_cNumDoc,                                      
--CND_ASIENTO_VOUCHER.Asd_dFecDoc,CND_ASIENTO_VOUCHER.Asd_cTipoDocRef,                                                         
--CND_ASIENTO_VOUCHER.Asd_cSerieDocRef,CND_ASIENTO_VOUCHER.Asd_cNumDocRef,                                                         
--case when year(CND_ASIENTO_VOUCHER.Asd_dFecDocRef)=1900 then null else CND_ASIENTO_VOUCHER.Asd_dFecDocRef end as 'Asd_dFecDocRef',                                                         
--CND_ASIENTO_VOUCHER.Com_cTipoIgv,CND_ASIENTO_VOUCHER.Asd_nMontoInafecto,                                                         
--CND_ASIENTO_VOUCHER.Asd_cRetencion,CNT_TIPODOC.Tdo_cNombreLargo,                                                         
--CNC_ASIENTO_VOUCHER.Ase_cTipoMoneda,case when year(CND_ASIENTO_VOUCHER.Asd_dFechaSpot)=1900 then null else CND_ASIENTO_VOUCHER.Asd_dFechaSpot end as 'Asd_dFechaSpot',                                                        
--CND_ASIENTO_VOUCHER.Asd_cNumSpot,CND_ASIENTO_VOUCHER.Imp_nPorcentaje,                                                         
--case when year(CND_ASIENTO_VOUCHER.Asd_dFecVen)=1900 then null else CND_ASIENTO_VOUCHER.Asd_dFecVen end as 'Asd_dFecVen',                                                        
--CND_ASIENTO_VOUCHER.Asd_cBaseImp,' '  AS Asd_cCodSunat,                                                        
--CND_ASIENTO_VOUCHER.Asd_cComprobante as Asd_cComprob,                                                        
--CNM_ENTIDAD.Ten_cTipoEntidad,CNC_ASIENTO_VOUCHER.Asd_cEstadoO,                    
--CNC_ASIENTO_VOUCHER.Asd_cEstadoD,                    
--0 as DifCambio                    
--INTO #TMPCOMPRAS                    
--FROM    CNC_ASIENTO_VOUCHER                                                         
--LEFT JOIN CND_ASIENTO_VOUCHER ON                                                        
--CNC_ASIENTO_VOUCHER.Emp_cCodigo = CND_ASIENTO_VOUCHER.Emp_cCodigo AND                                                         
--CNC_ASIENTO_VOUCHER.Pan_cAnio = CND_ASIENTO_VOUCHER.Pan_cAnio AND                                                         
--CNC_ASIENTO_VOUCHER.Per_cPeriodo = CND_ASIENTO_VOUCHER.Per_cPeriodo AND                                                         
--CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = CND_ASIENTO_VOUCHER.Lib_cTipoLibro AND                                                         
--CNC_ASIENTO_VOUCHER.Ase_nVoucher = CND_ASIENTO_VOUCHER.Ase_nVoucher                                                        
--LEFT JOIN CNT_TIPODOC ON                                                          
--CND_ASIENTO_VOUCHER.Asd_cTipoDoc = CNT_TIPODOC.Tdo_cCodigo AND                                                         
--CND_ASIENTO_VOUCHER.Emp_cCodigo = CNT_TIPODOC.Emp_cCodigo                                                         
--LEFT JOIN CNM_ENTIDAD ON                                                         
--CND_ASIENTO_VOUCHER.Emp_cCodigo = CNM_ENTIDAD.Emp_cCodigo AND                                                         
--CND_ASIENTO_VOUCHER.Ten_cTipoEntidad = CNM_ENTIDAD.Ten_cTipoEntidad AND                                                         
--CND_ASIENTO_VOUCHER.Ent_cCodEntidad = CNM_ENTIDAD.Ent_cCodEntidad                                                         
--LEFT JOIN  CNM_PLAN_CTA ON                                                         
--CND_ASIENTO_VOUCHER.Emp_cCodigo = CNM_PLAN_CTA.Emp_cCodigo AND                                                         
--CND_ASIENTO_VOUCHER.Pan_cAnio = CNM_PLAN_CTA.Pan_cAnio AND                                                         
--CND_ASIENTO_VOUCHER.Pla_cCuentaContable = CNM_PLAN_CTA.Pla_cCuentaContable                                                                                     
--WHERE                             
--CNC_ASIENTO_VOUCHER.Ase_cDeleted <> '*' AND CND_ASIENTO_VOUCHER.Asd_cDeleted <> '*' AND                                                         
--CND_ASIENTO_VOUCHER.Asd_cDestino = '0' AND CNC_ASIENTO_VOUCHER.Emp_cCodigo=@Emp_cCodigo AND                                                         
--(CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = @LIBRO OR CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = @LIBROHON) AND                                                        
--CNC_ASIENTO_VOUCHER.Pan_cAnio = @Pan_cAnio AND CNC_ASIENTO_VOUCHER.Per_cPeriodo=@Per_cPeriodo                                                         
---- And CNC_ASIENTO_VOUCHER.Ase_nVoucher='0612000075'                       
--ORDER BY CNC_ASIENTO_VOUCHER.Per_cPeriodo, CNC_ASIENTO_VOUCHER.Ase_nVoucher, CND_ASIENTO_VOUCHER.Asd_nItem                                                        
                    
SELECT cav.Ase_dFecha, cav.Emp_cCodigo, cav.Ase_cNummov, cav.Per_cPeriodo, cav.Pan_cAnio, cav.Lib_cTipoLibro, cav.Ase_nVoucher, CAV2.Asd_nItem,
CAV2.Pla_cCuentaContable, cpc.Pla_cNombreCuenta, cpc.Pla_cProvision, CAV2.Asd_cGlosa, CAV2.Asd_nDebeSoles, CAV2.Asd_nHaberSoles, CAV2.Asd_nDebeMonExt, CAV2.Asd_nHaberMonExt,
CAV2.Asd_nTipoCambio, ce.Ent_cCodEntidad,
 ce.Ent_nRuc, ce.Ent_cPersona, CAV2.Asd_cTipoDoc, CAV2.Asd_cSerieDoc, CAV2.Asd_cNumDoc,
CAV2.Asd_dFecDoc, CAV2.Asd_cTipoDocRef, CAV2.Asd_cSerieDocRef, CAV2.Asd_cNumDocRef, CASE WHEN YEAR(cav2.Asd_dFecDocRef) = 1900 THEN NULL ELSE CAV2.Asd_dFecDocRef END AS 'Asd_dFecDocRef',
CAV2.Com_cTipoIgv, CAV2.Asd_nMontoInafecto, CAV2.Asd_cRetencion, ct.Tdo_cNombreLargo, cav.Ase_cTipoMoneda,
CASE WHEN YEAR(CAV2.Asd_dFechaSpot) = 1900 THEN NULL ELSE CAV2.Asd_dFechaSpot END AS 'Asd_dFechaSpot', CAV2.Asd_cNumSpot, CAV2.Imp_nPorcentaje,
CASE WHEN YEAR(CAV2.Asd_dFecVen) = 1900 THEN NULL ELSE CAV2.Asd_dFecVen END AS 'Asd_dFecVen', CAV2.Asd_cBaseImp, '' AS Asd_cCodSunat, CAV2.Asd_cComprobante AS Asd_cComprob,
ISNULL(ce.Ten_cTipoEntidad, '') AS Ten_cTipoEntidad, cav.Asd_cEstadoO, cav.Asd_cEstadoD, 0 AS DifCambio, 
ISNULL((SELECT t.Tab_cCodSunat FROM dbo.TABLA T WHERE T.Emp_cCodigo = CAV2.Emp_cCodigo AND T.Tab_cCodigo = ce.Ent_cTipoDoc AND t.Tab_cTabla = '003'), '') AS Ent_cTipoDoc, CTM.Mon_cCodSunat,
ISNULL(CAV2.Id_Exoneracion, '') AS Id_Exoneracion, ISNULL(CAV2.Id_Tipo_Renta, '') AS Id_Tipo_Renta, ISNULL(CAV2.Id_Modalidad, '') AS Id_Modalidad, ISNULL(CAV2.Id_Aduana, '') AS Id_Aduana,
ISNULL(CAV2.Id_Clasific_Servicio, '') AS Id_Clasific_Servicio, ISNULL(ce.Ent_cFlagDomiciliado, '1') AS Ent_cFlagDomiciliado, ISNULL(Id_Pais, '') AS Id_Pais, ISNULL(Id_Convenio, '') AS Id_Convenio 
INTO #TMPCOMPRAS   FROM dbo.CNC_ASIENTO_VOUCHER CAV LEFT JOIN dbo.CND_ASIENTO_VOUCHER CAV2
ON CAV.Emp_cCodigo = CAV2.Emp_cCodigo AND CAV.Pan_cAnio = CAV2.Pan_cAnio AND CAV.Per_cPeriodo = CAV2.Per_cPeriodo
AND CAV.Lib_cTipoLibro = CAV2.Lib_cTipoLibro AND CAV.Ase_nVoucher = CAV2.Ase_nVoucher
LEFT JOIN dbo.CNT_TIPODOC CT ON CAV2.Emp_cCodigo = CT.Emp_cCodigo AND CAV2.Asd_cTipoDoc = CT.Tdo_cCodigo
LEFT JOIN dbo.CNM_ENTIDAD CE ON CAV2.Emp_cCodigo = CE.Emp_cCodigo
AND CAV2.Ten_cTipoEntidad = CE.Ten_cTipoEntidad
AND CAV2.Ent_cCodEntidad = CE.Ent_cCodEntidad
LEFT JOIN dbo.CNM_PLAN_CTA CPC ON CAV2.Emp_cCodigo = CPC.Emp_cCodigo
AND CAV2.Pan_cAnio = CPC.Pan_cAnio AND CAV2.Pla_cCuentaContable = CPC.Pla_cCuentaContable
LEFT JOIN dbo.CNT_TIPO_MONEDA CTM ON CTM.Emp_cCodigo = CAV2.Emp_cCodigo AND CTM.Mon_cCodigo = CAV2.Asd_cTipoMoneda
WHERE CAV.Ase_cDeleted <> '*' AND CAV2.Asd_cDeleted <> '*' AND CAV2.Asd_cDestino = '0' AND CAV.Emp_cCodigo = @Emp_cCodigo
AND (CAV.Lib_cTipoLibro = @LIBRO OR CAV.Lib_cTipoLibro = @LIBROHON) AND CAV.Pan_cAnio = @Pan_cAnio AND CAV.Per_cPeriodo = @Per_cPeriodo
ORDER BY cav.Per_cPeriodo, cav.Ase_nVoucher, CAV2.Asd_nItem                    

                    
/* Ahora cargo los registros con estado 9*/             
                    
--select                     
--CNC_ASIENTO_VOUCHER.Ase_dfecha,                    
--CNC_ASIENTO_VOUCHER.Emp_cCodigo,CNC_ASIENTO_VOUCHER.Ase_cNumMov,                                                         
--CNC_ASIENTO_VOUCHER.Per_cPeriodo,CNC_ASIENTO_VOUCHER.Pan_cAnio,              
--CNC_ASIENTO_VOUCHER.Lib_cTipoLibro,CNC_ASIENTO_VOUCHER.Ase_nVoucher,                                                         
--CND_ASIENTO_VOUCHER.Asd_nItem,CND_ASIENTO_VOUCHER.Pla_cCuentaContable,                                                         
--CNM_PLAN_CTA.Pla_cNombreCuenta,CNM_PLAN_CTA.Pla_cProvision,                                                         
--CND_ASIENTO_VOUCHER.Asd_cGlosa,CND_ASIENTO_VOUCHER.Asd_nDebeSoles,                                                    
--CND_ASIENTO_VOUCHER.Asd_nHaberSoles,CND_ASIENTO_VOUCHER.Asd_nDebeMonExt,                                                         
--CND_ASIENTO_VOUCHER.Asd_nHaberMonExt,CND_ASIENTO_VOUCHER.Asd_nTipoCambio,                                                         
--CNM_ENTIDAD.Ent_cCodEntidad,Ent_nRuc,                                                         
--CNM_ENTIDAD.Ent_cPersona,CND_ASIENTO_VOUCHER.Asd_cTipoDoc,                                                         
--CND_ASIENTO_VOUCHER.Asd_cSerieDoc,CND_ASIENTO_VOUCHER.Asd_cNumDoc,                                                         
--CND_ASIENTO_VOUCHER.Asd_dFecDoc,CND_ASIENTO_VOUCHER.Asd_cTipoDocRef,                                                         
--CND_ASIENTO_VOUCHER.Asd_cSerieDocRef,CND_ASIENTO_VOUCHER.Asd_cNumDocRef,                                                         
--case when year(CND_ASIENTO_VOUCHER.Asd_dFecDocRef)=1900 then null else CND_ASIENTO_VOUCHER.Asd_dFecDocRef end as 'Asd_dFecDocRef',                                                         
--CND_ASIENTO_VOUCHER.Com_cTipoIgv,                                                         
--CND_ASIENTO_VOUCHER.Asd_nMontoInafecto,CND_ASIENTO_VOUCHER.Asd_cRetencion,                                                         
--CNT_TIPODOC.Tdo_cNombreLargo,CNC_ASIENTO_VOUCHER.Ase_cTipoMoneda,                                                         
--case when year(CND_ASIENTO_VOUCHER.Asd_dFechaSpot)=1900 then null else CND_ASIENTO_VOUCHER.Asd_dFechaSpot end as 'Asd_dFechaSpot',                                                        
--CND_ASIENTO_VOUCHER.Asd_cNumSpot,CND_ASIENTO_VOUCHER.Imp_nPorcentaje,                                                         
--case when year(CND_ASIENTO_VOUCHER.Asd_dFecVen)=1900 then null else CND_ASIENTO_VOUCHER.Asd_dFecVen end as 'Asd_dFecVen',                                                        
--CND_ASIENTO_VOUCHER.Asd_cBaseImp,' '  AS Asd_cCodSunat,                                                        
--CND_ASIENTO_VOUCHER.Asd_cComprobante as Asd_cComprob,                                                        
--CNM_ENTIDAD.Ten_cTipoEntidad,CNC_ASIENTO_VOUCHER.Asd_cEstadoO,                    
--CNC_ASIENTO_VOUCHER.Asd_cEstadoD,                    
--0 as DifCambio                    
--INTO #TMPCOMPRAS2                    
--FROM    CNC_ASIENTO_VOUCHER                   
--LEFT JOIN CND_ASIENTO_VOUCHER ON                                                        
--CNC_ASIENTO_VOUCHER.Emp_cCodigo = CND_ASIENTO_VOUCHER.Emp_cCodigo AND                                                         
--CNC_ASIENTO_VOUCHER.Pan_cAnio = CND_ASIENTO_VOUCHER.Pan_cAnio AND                               
--CNC_ASIENTO_VOUCHER.Per_cPeriodo = CND_ASIENTO_VOUCHER.Per_cPeriodo AND                                                         
--CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = CND_ASIENTO_VOUCHER.Lib_cTipoLibro AND                                                         
--CNC_ASIENTO_VOUCHER.Ase_nVoucher = CND_ASIENTO_VOUCHER.Ase_nVoucher                                                        
--LEFT JOIN CNT_TIPODOC ON                                                          
--CND_ASIENTO_VOUCHER.Asd_cTipoDoc = CNT_TIPODOC.Tdo_cCodigo AND                                                         
--CND_ASIENTO_VOUCHER.Emp_cCodigo = CNT_TIPODOC.Emp_cCodigo                                                         
--LEFT JOIN CNM_ENTIDAD ON                                                         
--CND_ASIENTO_VOUCHER.Emp_cCodigo = CNM_ENTIDAD.Emp_cCodigo AND                                                         
--CND_ASIENTO_VOUCHER.Ten_cTipoEntidad = CNM_ENTIDAD.Ten_cTipoEntidad AND                                                         
--CND_ASIENTO_VOUCHER.Ent_cCodEntidad = CNM_ENTIDAD.Ent_cCodEntidad                                                         
--LEFT JOIN  CNM_PLAN_CTA ON                                                         
--CND_ASIENTO_VOUCHER.Emp_cCodigo = CNM_PLAN_CTA.Emp_cCodigo AND                                                         
--CND_ASIENTO_VOUCHER.Pan_cAnio = CNM_PLAN_CTA.Pan_cAnio AND                                                         
--CND_ASIENTO_VOUCHER.Pla_cCuentaContable = CNM_PLAN_CTA.Pla_cCuentaContable                                                                            
--WHERE                                               
--CNC_ASIENTO_VOUCHER.Ase_cDeleted <> '*' AND CND_ASIENTO_VOUCHER.Asd_cDeleted <> '*' AND                                                  
--CND_ASIENTO_VOUCHER.Asd_cDestino = '0' AND CNC_ASIENTO_VOUCHER.Emp_cCodigo= @Emp_cCodigo AND                                                         
--(CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = @LIBRO OR CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = @LIBROHON) AND                     
--year(CNC_ASIENTO_VOUCHER.Ase_dfechaModifica) = @Pan_cAnio/*YEAR(getdate())*/ and month(CNC_ASIENTO_VOUCHER.Ase_dfechaModifica) = @Per_cPeriodo                    
--and CNC_ASIENTO_VOUCHER.Asd_cEstadoD = '9'                    
--ORDER BY CNC_ASIENTO_VOUCHER.Per_cPeriodo, CNC_ASIENTO_VOUCHER.Ase_nVoucher, CND_ASIENTO_VOUCHER.Asd_nItem                    
                     
SELECT cav.Ase_dFecha, cav.Emp_cCodigo, cav.Ase_cNummov, cav.Per_cPeriodo, cav.Pan_cAnio, cav.Lib_cTipoLibro, cav.Ase_nVoucher, CAV2.Asd_nItem,
CAV2.Pla_cCuentaContable, cpc.Pla_cNombreCuenta, cpc.Pla_cProvision, CAV2.Asd_cGlosa, CAV2.Asd_nDebeSoles, CAV2.Asd_nHaberSoles, CAV2.Asd_nDebeMonExt, CAV2.Asd_nHaberMonExt,
CAV2.Asd_nTipoCambio, ce.Ent_cCodEntidad,
 ce.Ent_nRuc, ce.Ent_cPersona, CAV2.Asd_cTipoDoc, CAV2.Asd_cSerieDoc, CAV2.Asd_cNumDoc,
CAV2.Asd_dFecDoc, CAV2.Asd_cTipoDocRef, CAV2.Asd_cSerieDocRef, CAV2.Asd_cNumDocRef, CASE WHEN YEAR(cav2.Asd_dFecDocRef) = 1900 THEN NULL ELSE CAV2.Asd_dFecDocRef END AS 'Asd_dFecDocRef',
CAV2.Com_cTipoIgv, CAV2.Asd_nMontoInafecto, CAV2.Asd_cRetencion, ct.Tdo_cNombreLargo, cav.Ase_cTipoMoneda,
CASE WHEN YEAR(CAV2.Asd_dFechaSpot) = 1900 THEN NULL ELSE CAV2.Asd_dFechaSpot END AS 'Asd_dFechaSpot', CAV2.Asd_cNumSpot, CAV2.Imp_nPorcentaje,
CASE WHEN YEAR(CAV2.Asd_dFecVen) = 1900 THEN NULL ELSE CAV2.Asd_dFecVen END AS 'Asd_dFecVen', CAV2.Asd_cBaseImp, '' AS Asd_cCodSunat, CAV2.Asd_cComprobante AS Asd_cComprob,
ISNULL(ce.Ten_cTipoEntidad, '') AS Ten_cTipoEntidad, cav.Asd_cEstadoO, cav.Asd_cEstadoD, 0 AS DifCambio, 
ISNULL((SELECT t.Tab_cCodSunat FROM dbo.TABLA T WHERE T.Emp_cCodigo = CAV2.Emp_cCodigo AND T.Tab_cCodigo = ce.Ent_cTipoDoc AND t.Tab_cTabla = '003'), '') AS Ent_cTipoDoc, CTM.Mon_cCodSunat,
ISNULL(CAV2.Id_Exoneracion, '') AS Id_Exoneracion, ISNULL(CAV2.Id_Tipo_Renta, '') AS Id_Tipo_Renta, ISNULL(CAV2.Id_Modalidad, '') AS Id_Modalidad, ISNULL(CAV2.Id_Aduana, '') AS Id_Aduana,
ISNULL(CAV2.Id_Clasific_Servicio, '') AS Id_Clasific_Servicio, ISNULL(ce.Ent_cFlagDomiciliado, '1') AS Ent_cFlagDomiciliado, ISNULL(Id_Pais, '') AS Id_Pais, ISNULL(Id_Convenio, '') AS Id_Convenio  
INTO #TMPCOMPRAS2  FROM dbo.CNC_ASIENTO_VOUCHER CAV LEFT JOIN dbo.CND_ASIENTO_VOUCHER CAV2
ON CAV.Emp_cCodigo = CAV2.Emp_cCodigo AND CAV.Pan_cAnio = CAV2.Pan_cAnio AND CAV.Per_cPeriodo = CAV2.Per_cPeriodo
AND CAV.Lib_cTipoLibro = CAV2.Lib_cTipoLibro AND CAV.Ase_nVoucher = CAV2.Ase_nVoucher
LEFT JOIN dbo.CNT_TIPODOC CT ON CAV2.Emp_cCodigo = CT.Emp_cCodigo AND CAV2.Asd_cTipoDoc = CT.Tdo_cCodigo
LEFT JOIN dbo.CNM_ENTIDAD CE ON CAV2.Emp_cCodigo = CE.Emp_cCodigo
AND CAV2.Ten_cTipoEntidad = CE.Ten_cTipoEntidad
AND CAV2.Ent_cCodEntidad = CE.Ent_cCodEntidad
LEFT JOIN dbo.CNM_PLAN_CTA CPC ON CAV2.Emp_cCodigo = CPC.Emp_cCodigo
AND CAV2.Pan_cAnio = CPC.Pan_cAnio AND CAV2.Pla_cCuentaContable = CPC.Pla_cCuentaContable
LEFT JOIN dbo.CNT_TIPO_MONEDA CTM ON CTM.Emp_cCodigo = CAV2.Emp_cCodigo AND CTM.Mon_cCodigo = CAV2.Asd_cTipoMoneda
WHERE CAV.Ase_cDeleted <> '*' AND CAV2.Asd_cDeleted <> '*' AND CAV2.Asd_cDestino = '0' AND CAV.Emp_cCodigo = @Emp_cCodigo
AND (CAV.Lib_cTipoLibro = @LIBRO OR CAV.Lib_cTipoLibro = @LIBROHON) AND 
YEAR(cav.Ase_dFechaModifica) = @Pan_cAnio AND MONTH(cav.Ase_dFechaModifica) = @Per_cPeriodo AND cav.Asd_cEstadoD = '9'
ORDER BY cav.Per_cPeriodo, cav.Ase_nVoucher, CAV2.Asd_nItem
                     
insert into #TMPCOMPRAS select * from #TMPCOMPRAS2                    
                    
Delete from #TMPCOMPRAS2                    
                            
                            
--SELECT * FROM CND_ASIENTO_VOUCHER WHERE Ase_nVoucher = '0608000054' And Pan_cAnio = '2010' and Per_cPeriodo = '08' And Asd_cDestino = '0'                            
----------------------------------------------------                                                        
-- ACTUALIZA TIPO DE CAMBIO VEP EN DOCUMENTOS, LOS DE LA CUENTA 40                                                         
              
DECLARE @varPla_cCta varchar(12)              
DECLARE @varAsd_nTipoCambio NUMERIC(14,3)              
DECLARE @varAse_cNumMov char(10)              
DECLARE @varAse_nVoucher char(10)              
declare @vFETCH_STATUS_CAB int              
declare @vFETCH_STATUS_DET int              
-------------------------------                          
DECLARE Compras_Cursor_TC CURSOR FOR              
SELECT Ase_cNumMov, Ase_nVoucher, Pla_cCuentaContable, Asd_nTipoCambio              
FROM #TMPCOMPRAS              
WHERE Emp_cCodigo = @Emp_cCodigo and (left(Pla_cCuentaContable,2)='40' or left(Pla_cCuentaContable,2)='60' or left(Pla_cCuentaContable,2)='70')                      
              
OPEN Compras_Cursor_TC              
FETCH NEXT FROM Compras_Cursor_TC              
INTO @varAse_cNumMov, @varAse_nVoucher, @varPla_cCta, @varAsd_nTipoCambio              
              
 set @vFETCH_STATUS_CAB = @@FETCH_STATUS              
              
 WHILE @vFETCH_STATUS_CAB = 0              
 BEGIN              
              
  update #TMPCOMPRAS set Asd_nTipoCambio = @varAsd_nTipoCambio              
  where Ase_cNumMov = @varAse_cNumMov and Ase_nVoucher = @varAse_nVoucher and Asd_nTipoCambio <> 0              
              
  FETCH NEXT FROM Compras_Cursor_TC              
  INTO  @varAse_cNumMov, @varAse_nVoucher, @varPla_cCta, @varAsd_nTipoCambio              
              
  set @vFETCH_STATUS_CAB = @@FETCH_STATUS              
 END              
CLOSE Compras_Cursor_TC              
DEALLOCATE Compras_Cursor_TC              
---------------------------------------------------                        
DECLARE @auxEnt_cCodEntidad char(5)                                                        
DECLARE @auxAsd_cTipoDoc char(3)                                                        
DECLARE @auxAsd_cSerieDoc VARCHAR(20)                                                       
DECLARE @auxAsd_cNumDoc VARCHAR(25)                                                    
DECLARE @auxAsd_dFecDoc datetime                                   
DECLARE @VarAsd_dFecVen datetime                                                        
DECLARE @auxItem int                                                        
                                                        
DECLARE @varLib_cTipoLibro CHAR(2)                      
                                                        
DECLARE @varAsd_nItem int                                                        
DECLARE @varEnt_cCodEntidad char(5)                                           
DECLARE @varAsd_cTipoDoc char(3)                                                        
DECLARE @varAsd_cSerieDoc VARCHAR(20)                                                        
DECLARE @varAsd_cNumDoc VARCHAR(25)                                                        
                                      
DECLARE @varPla_cNomCta varchar(120)                                                        
DECLARE @varAsd_nDebeSoles decimal(18,3)                      
DECLARE @varAsd_nHaberSoles decimal(18,3)                                                 
DECLARE @varAsd_nDebeMonExt decimal(18,3)                                                        
DECLARE @varAsd_nHaberMonExt decimal(18,3)                                                        
DECLARE @varAsd_nMontoInafecto decimal(18,3)                            
                                                        
DECLARE @varAsd_cBaseImp char(3)                                          
DECLARE @x_varAsd_cBaseImp char(3)                                          
DECLARE @Asd_cBaseImp char(3)                                                        
                                                        
                                                        
DECLARE @ASD_DFECHASPOT DATETIME                                                        
DECLARE @ASD_CNUMSPOT VARCHAR(25)                                                        
DECLARE @Asd_cRetencion CHAR(1)                                                        
DECLARE @Imp_nPorcentaje NUMERIC(14,3)                                                        
                                                        
                                                        
DECLARE @IGV_A VARCHAR(12)                             
DECLARE @IGV_B VARCHAR(12)                                                        
DECLARE @IGV_C VARCHAR(12)                                                        
           
DECLARE @BASE_A VARCHAR(12)                                                        
DECLARE @BASE_B VARCHAR(12)                                                        
DECLARE @BASE_C VARCHAR(12)                                                        
                                                        
DECLARE @CTA_IGV VARCHAR(12)                                                        
DECLARE @CTA_RDEOG VARCHAR(12)                                                
                                                        
DECLARE @CTA_RDEOP VARCHAR(12)                                                        
DECLARE @CTA_DIFCG VARCHAR(12)                                                        
DECLARE @CTA_DIFCP VARCHAR(12)                                                        
DECLARE @cTipDocNC CHAR(2)                                                        
DECLARE @OTROS VARCHAR(12)             
                                                        
DECLARE @cTipoDocReg CHAR(2)                                              
DECLARE @TipMonedaCab CHAR(3)                                                        
DECLARE @MontoTotalS NUMERIC(18,3)                                                        
DECLARE @MontoTotalD NUMERIC(18,3)                                                        
DECLARE @MontoOtros_S NUMERIC(18,3)                                                        
DECLARE @MontoDifCam_S NUMERIC(18,3)                                                        
DECLARE @MontoOtros_D NUMERIC(18,3)                                                        
DECLARE @MontoDifCam_D NUMERIC(18,3)                                                        
DECLARE @MontoBaseTot_S NUMERIC(18,3)                                                        
DECLARE @MontoBaseTot_D NUMERIC(18,3)                                                        
DECLARE @MontoTotalIgv_S NUMERIC(18,3)                                                        
DECLARE @MontoTotalIgv_D NUMERIC(18,3)                                                        
DECLARE @MontoBaseA_S NUMERIC(18,3)                                                        
DECLARE @MontoBaseB_S NUMERIC(18,3)                                                        
DECLARE @MontoBaseC_S NUMERIC(18,3)                                                        
DECLARE @MontoIgvA_S NUMERIC(18,3)                                                        
DECLARE @MontoIgvB_S NUMERIC(18,3)                                                        
DECLARE @MontoIgvC_S NUMERIC(18,3)                                 
DECLARE @MontoBaseA_D NUMERIC(18,3)                                                        
DECLARE @MontoBaseB_D NUMERIC(18,3)                                                        
DECLARE @MontoBaseC_D NUMERIC(18,3)                                                        
DECLARE @MontoIgvA_D NUMERIC(18,3)                                                        
DECLARE @MontoIgvB_D NUMERIC(18,3)                                            
DECLARE @MontoIgvC_D NUMERIC(18,3)                                                        
DECLARE @CTAREG_IGVA VARCHAR(12)                                                
DECLARE @CTAREG_IGVB VARCHAR(12)                                                        
DECLARE @CTAREG_IGVC VARCHAR(12)                                                        
DECLARE @MONTOINAFECTO_S NUMERIC(18,3)                                                        
DECLARE @MONTOINAFECTO_D NUMERIC(18,3)                                                        
DECLARE @nPorcen_igv_A NUMERIC(18,3)                                                        
DECLARE @nPorcen_igv_B NUMERIC(18,3)                                                        
DECLARE @nPorcen_igv_C NUMERIC(18,3)                                                        
DECLARE @nPorcentajeIgv NUMERIC(5,3)                                                        
                                              
DECLARE @MontoSISC_S NUMERIC(18,3)                                                
DECLARE @MontoSISC_D NUMERIC(18,3)                                                        
                                                        
DECLARE @MontoSReintegro NUMERIC(18,3)                                                        
DECLARE @MontoDReintegro NUMERIC(18,3)                                                        
                                                        
DECLARE @CONSTANTE_NC NUMERIC(18,3)            
DECLARE @cCadena varchar(400)                                                        
                                                      
DECLARE @NroRegConfigImp Integer                                    
DECLARE @NroRegConfigOpera Integer                                    
DECLARE @xPla_cCuentaContable VarChar(12)                                             
                                                        
---------------------------------------------------------------------------------------------                                                        
SET @cTipDocNC= ISNULL((SELECT DBO.TRIMSQL(RTRIM(Cod_cValorParam)) FROM CND_CONFIG_OPERA                                                        
  WHERE Emp_cCodigo=@Emp_cCodigo AND Pan_cAnio=@Pan_cAnio AND Cop_cCodigo='012'), '')                                                        
---------------------------------------------------------------------------------------------                                                        
/** BUSCA CTAS DE REDONDEO **/                                                         
SET @CTA_RDEOG = ISNULL((SELECT Pla_cCuentaContable FROM CNM_PLAN_CTA                   
    WHERE Emp_cCodigo=@Emp_cCodigo AND Pan_cAnio=@Pan_cAnio AND Pla_cRedondeo='G'),'')                                                        
SET @CTA_RDEOP = ISNULL((SELECT Pla_cCuentaContable FROM CNM_PLAN_CTA                                                         
    WHERE Emp_cCodigo=@Emp_cCodigo AND Pan_cAnio=@Pan_cAnio AND Pla_cRedondeo='P'), '')                                                        
---------------------------------------------------------------------------------------------                                                        
/** BUSCA CTAS DE DIFERENCIA DE CAMBIO **/                                                         
SET @CTA_DIFCG = ISNULL((SELECT Pla_cCuentaContable FROM CNM_PLAN_CTA                                                         
    WHERE Emp_cCodigo=@Emp_cCodigo AND Pan_cAnio=@Pan_cAnio AND Pla_cDifCambio='G'),'')                              
SET @CTA_DIFCP = ISNULL((SELECT Pla_cCuentaContable FROM CNM_PLAN_CTA                                                         
    WHERE Emp_cCodigo=@Emp_cCodigo AND Pan_cAnio=@Pan_cAnio AND Pla_cDifCambio='P'),'')                                                        
                                                        
---------------------------------------------------------------------------------------------                                                        
-- BUSCANDO CUENTAS DE TIPO ISC                                                        
SELECT @Registros = COUNT(CUENTAS) FROM #TMP_ISC                                                        
IF  isnull(@Registros,0) > 0                    
 set @ISC = 'ISC'                    
ELSE                    
 set @ISC = ''                    
---------------------------------------------------------------------------------------------                                      
/** SE GENERA TABLA TEMPORAL DE COMPRAS **/                                                        
IF EXISTS (SELECT * FROM ..sysobjects WHERE name like 'TMPREGISTROLECOMPRAS%') DROP TABLE TMPREGISTROLECOMPRAS                                                        
IF EXISTS (SELECT * FROM tempdb..sysobjects WHERE name like '#TMP_REGLECOMPRAS') DROP TABLE #TMP_REGLECOMPRAS                                 
                                                        
EXEC spCn_CrearTablaTemporal1 'LibroElectCompras'                                       
SELECT * INTO #TMP_REGLECOMPRAS FROM TMPREGISTROLECOMPRAS                                                        
                                                        
                                                        
IF EXISTS (SELECT * FROM ..sysobjects WHERE name like 'TMPREGISTROLECOMPRAS%') DROP TABLE TMPREGISTROLECOMPRAS                                                        
---------------------------------------------------------------------------------------------                                        
 
/** Agrupando Datos por Sub-Diario, Nro Movto y Nro Voucher CABECERA**/                                             
DECLARE Compras_Cursor_CabeceraLE CURSOR FOR                                                        
SELECT Lib_cTipoLibro, Ase_cNumMov, Ase_nVoucher, Ase_cTipoMoneda                                                        
FROM #TMPCOMPRAS                                                         
WHERE Emp_cCodigo = @Emp_cCodigo                                                     
                                                        
GROUP BY Lib_cTipoLibro, Ase_cNumMov, Ase_nVoucher, Ase_cTipoMoneda                                               
ORDER BY Lib_cTipoLibro, Ase_cNumMov, Ase_nVoucher                                                        
                                                        
OPEN Compras_Cursor_CabeceraLE                                                        
FETCH NEXT FROM Compras_Cursor_CabeceraLE                                                         
INTO @varLib_cTipoLibro, @varAse_cNumMov, @varAse_nVoucher, @TipMonedaCab                                                        
                                                        
 set @vFETCH_STATUS_CAB = @@FETCH_STATUS                                                                                                           
 WHILE @vFETCH_STATUS_CAB = 0                                                        
 BEGIN                                                        
  SET @cTipoDocReg = ''                                                        
  SET @TipMonedaCab = ''                                                        
  SET @MontoTotalS = 0                                                         
  SET @MontoTotalD = 0                                                         
  SET @MontoOtros_S = 0                                                         
  SET @MontoOtros_D = 0                                                         
  SET @MontoDifCam_S = 0                      
  SET @MontoDifCam_D = 0                                                         
  SET @MontoBaseTot_S = 0                                                         
  SET @MontoBaseTot_D = 0                                        
  SET @MontoTotalIgv_S = 0                                                         
  SET @MontoTotalIgv_D = 0                                                         
  SET @MontoBaseA_S = 0                                                         
  SET @MontoBaseB_S = 0                                                         
  SET @MontoBaseC_S = 0                                                         
  SET @MontoIgvA_S = 0                                                         
  SET @MontoIgvB_S = 0                 
  SET @MontoIgvC_S = 0                                                         
  SET @MontoBaseA_D = 0                                                        
  SET @MontoBaseB_D = 0                                                         
  SET @MontoBaseC_D = 0                                 
                                  
  SET @MontoIgvA_D = 0                                                         
  SET @MontoIgvB_D = 0                                                         
  SET @MontoIgvC_D = 0                                                        
  SET @CTAREG_IGVA = ''        
  SET @CTAREG_IGVB = ''                                                        
  SET @CTAREG_IGVC = ''                                                        
  SET @MONTOINAFECTO_S = 0                                                
  SET @MONTOINAFECTO_D = 0                                                         
  SET @nPorcen_igv_A = 0                                                         
  SET @nPorcen_igv_B = 0                                                         
  SET @nPorcen_igv_C = 0                                                         
  SET @nPorcentajeIgv = 0        
  SET @Asd_cBaseImp = ''                                                        
  SET @MontoSISC_S=0                                                   
  SET @MontoSISC_D=0                                                        
  SET @MontoSReintegro = 0                                                        
  SET @MontoDReintegro = 0                                                        
  SET @varAsd_nMontoInafecto = 0                                                         
  ---------------------------------------------------------------------------------------------                                          
  -- Buscando CTA X Pagar y Obtener Montos                                       
   ------------------------------------------------------------------------------------------------------------------------------------------                                    
   SELECT @xPla_cCuentaContable = Pla_cCuentaContable                                     
   FROM #TMPCOMPRAS                                                         
   WHERE Lib_cTipoLibro=@varLib_cTipoLibro AND Ase_cNumMov=@varAse_cNumMov AND Ase_nVoucher=@varAse_nVoucher                                     
         and DBO.TRIMSQL(Asd_cBaseImp) = ''                                      
                                             
  SELECT @NroRegConfigOpera = COUNT(1)                                    
  FROM #TMPCOMPRAS                                                         
  WHERE Lib_cTipoLibro=@varLib_cTipoLibro AND Ase_cNumMov=@varAse_cNumMov AND Ase_nVoucher=@varAse_nVoucher AND                                                         
   LEFT(Pla_cCuentaContable,2) IN (SELECT LEFT(Cod_cValorParam,2) FROM CND_CONFIG_OPERA                                                         
       WHERE Emp_cCodigo=@Emp_cCodigo AND Pan_cAnio=@Pan_cAnio AND Cop_cCodigo='010')                                    
       AND DBO.TRIMSQL(Asd_cBaseImp) = ''                                    
                                            
  IF @NroRegConfigOpera = 0                                    
   BEGIN                                    
     INSERT INTO CND_CONFIG_OPERA (Emp_cCodigo, Pan_cAnio, Cop_cCodigo, Cod_cValorParam, Cod_nIgvPorc, Cod_cEstado,                                     
           Cod_cDeleted, Cod_cUserCrea, Cod_dFechaCrea, Cod_cUserModifica, Cod_dFechaModifica, Cod_cEquipoUser)                                    
  VALUES(@Emp_cCodigo, @Pan_cAnio, '010', LEFT(@xPla_cCuentaContable,2), 0, 'A', '', HOST_NAME(), GETDATE(), HOST_NAME(), GETDATE(), HOST_NAME())                                    
   END                                     
   ------------------------------------------------------------------------------------------------------------------------------------------                                    
                                                          
  SELECT @auxEnt_cCodEntidad = ISNULL(Ent_cCodEntidad, ''),                                                         
   @auxAsd_cTipoDoc = Asd_cTipoDoc,                    
   @cTipoDocReg = Asd_cTipoDoc,                    
   @auxAsd_cSerieDoc = Asd_cSerieDoc,                    
   @auxAsd_cNumDoc = Asd_cNumDoc,                    
   @auxAsd_dFecDoc = Asd_dFecDoc,                    
   @TipMonedaCab=Ase_cTipoMoneda,                    
                           
   @MontoTotalS = Asd_nHaberSoles - Asd_nDebeSoles  ,                                                     
   @MontoTotalD = Asd_nHaberMonExt - Asd_nDebeMonExt  ,                                                   
                                                        
   @Asd_cBaseImp = Asd_cBaseImp                                                        
  FROM #TMPCOMPRAS                                                         
  WHERE Lib_cTipoLibro=@varLib_cTipoLibro AND Ase_cNumMov=@varAse_cNumMov AND Ase_nVoucher=@varAse_nVoucher AND                                                         
   LEFT(Pla_cCuentaContable,2) IN (SELECT LEFT(Cod_cValorParam,2) FROM CND_CONFIG_OPERA          
       WHERE Emp_cCodigo=@Emp_cCodigo AND Cop_cCodigo='010') --AND Pan_cAnio=@Pan_cAnio)                                                        
       and DBO.TRIMSQL(Asd_cBaseImp) = ''                                                        
  ---------------------------------------------------------------------------------------------                                -- Grabando Datos Generales en le Registro de Compras                                                         
  INSERT #TMP_REGLECOMPRAS                                                        
  SELECT DISTINCT Emp_cCodigo, Ase_cNumMov, Per_cPeriodo, Ase_nVoucher,Ase_dfecha,@Pan_cAnio, '1', Asd_nTipoCambio,                                                         
   DBO.TRIMSQL(Asd_cGlosa), ISNULL(Ent_cCodEntidad, ''), ISNULL(Ent_nRuc, ''), ISNULL(Ent_cPersona, ''),                                                  
   DBO.TRIMSQL(Asd_cTipoDoc), DBO.TRIMSQL(Asd_cSerieDoc),DBO.TRIMSQL(Asd_cNumDoc), DBO.TRIMSQL(Asd_dFecDoc), DBO.TRIMSQL(Asd_cTipoDocRef), DBO.TRIMSQL(Asd_cSerieDocRef),                                                         
   DBO.TRIMSQL(Asd_cNumDocRef), ISNULL(Asd_dFecDocRef,''), 0, Asd_cRetencion, ISNULL(Tdo_cNombreLargo, ''),                                                          
   ISNULL(Ase_cTipoMoneda,''), '', '', '', '', '', '', '', '', '', '', '', '', '', '',                                                         
   0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,0,0,0,0, ISNULL(ASD_DFECHASPOT,''), ISNULL(ASD_CNUMSPOT,''),                                                         
Imp_nPorcentaje, ISNULL(Asd_dFecVen,'') ,                                                        
   ISNULL(Asd_cCodSunat,''), ISNULL(Asd_cComprob,'') ,ISNULL(Ten_cTipoEntidad, '') ,0,0,Asd_cEstadoO,Asd_cEstadoD,0, Ent_cTipoDoc, Mon_cCodSunat, Id_Exoneracion, Id_Tipo_Renta,
   Id_Modalidad, Id_Aduana, Id_Clasific_Servicio, Ent_cFlagDomiciliado , Id_Pais, Id_Convenio                                                                        
  FROM #TMPCOMPRAS                                                     
  WHERE DBO.TRIMSQL(Asd_cTipoDoc)<>'' and                                                         
        Lib_cTipoLibro=@varLib_cTipoLibro AND Ase_cNumMov=@varAse_cNumMov AND Ase_nVoucher=@varAse_nVoucher AND                                                         
        LEFT(Pla_cCuentaContable,2) IN (SELECT LEFT(Cod_cValorParam,2) FROM CND_CONFIG_OPERA                                                         
       WHERE Emp_cCodigo=@Emp_cCodigo  AND Cop_cCodigo='010') --AND Pan_cAnio=@Pan_cAnio )                                                        
       and DBO.TRIMSQL(Asd_cBaseImp) = ''                         
                                                          
  ---------------------------------------------------------------------------------------------                                                        
                                            
  -- Filtrando Datos de Movimiento por Voucher (Cuentas) DETALLE                                                        
  DECLARE Compras_Cursor_Detalle CURSOR FOR                                                        
                                                        
 SELECT A.Asd_nItem, A.Ent_cCodEntidad, A.Asd_cTipoDoc, A.Asd_cSerieDoc, A.Asd_cNumDoc,                                                       
   A.Pla_cCuentaContable, A.Pla_cNombreCuenta, A.Asd_nDebeSoles, A.Asd_nHaberSoles, A.Asd_nDebeMonExt,                                       
   A.Asd_nHaberMonExt, A.Asd_nMontoInafecto, A.Imp_nPorcentaje, A.Asd_nTipoCambio,                                                     
   A.Asd_cBaseImp,                                      
   IsNull((SELECT B.Cop_cCodigo                                          
   FROM CND_CONFIG_OPERA B WHERE B.Emp_cCodigo = A.Emp_cCodigo AND B.Pan_cAnio = A.Pan_cAnio AND                                           
   B.Cod_cValorParam = A.Pla_cCuentaContable AND B.Cop_cCodigo = '099'),'XXX') AS x_varAsd_cBaseImp                                         
FROM #TMPCOMPRAS A                                          
  WHERE A.Lib_cTipoLibro=@varLib_cTipoLibro AND                                              
   A.Ase_cNumMov=@varAse_cNumMov AND                                                       
   A.Ase_nVoucher=@varAse_nVoucher                                   
  ORDER BY A.Asd_nItem, A.Ase_cNumMov, A.Ase_nVoucher                                                  
                                            
  OPEN Compras_Cursor_Detalle                
  FETCH NEXT FROM Compras_Cursor_Detalle                                                        
  INTO @varAsd_nItem, @varEnt_cCodEntidad, @varAsd_cTipoDoc, @varAsd_cSerieDoc, @varAsd_cNumDoc,                                                         
   @varPla_cCta, @varPla_cNomCta, @varAsd_nDebeSoles, @varAsd_nHaberSoles, @varAsd_nDebeMonExt,                                                        
   @varAsd_nHaberMonExt, @varAsd_nMontoInafecto, @Imp_nPorcentaje, @varAsd_nTipoCambio,                 
   @varAsd_cBaseImp, @x_varAsd_cBaseImp                                      
                                                        
   SET @vFETCH_STATUS_DET = @@FETCH_STATUS                                                         
                                                        
   WHILE @vFETCH_STATUS_DET = 0                                                        
   BEGIN                                                  
                                                        
    -- 024 = OTROS                                
    -- 026 = HONORARIOS                                                        

                                                        
    SET @varAsd_cBaseImp= DBO.TRIMSQL ( isnull(@varAsd_cBaseImp,''))                                                        
    SELECT @NroRegConfigImp = COUNT(1) FROM CND_CONFIG_OPERA WITH(READUNCOMMITTED) WHERE Emp_cCodigo = @Emp_cCodigo  AND Pan_cAnio = @Pan_cAnio                       
    AND Cod_cValorParam = @varPla_cCta AND Cop_cCodigo = '099'                                                                     
                                                  
    -- INICIANDO VARIABLES                                                        
    SET @CTA_IGV = ''                                                        
    SET @IGV_A = ''                                                     
    SET @IGV_B = ''                                                        
    SET @IGV_C = ''                                                        
    SET @OTROS=''                        
                                                        
    SET @BASE_A = ''                                                        
    SET @BASE_B = ''                                                        
    SET @BASE_C = ''                                                        
                                                        
    SET @cCadena =''                                          
                                                            
  -- BUSCANDO CUENTAS DE LA COLUMNA A                                                        
IF ( @varAsd_cBaseImp= '006' )                                              
 BEGIN                                                    
  IF (LEFT(@varPla_cCta,2)='40' AND @NroRegConfigImp = '1') or (LEFT(@varPla_cCta,2)='64'  AND @NroRegConfigImp = '1' )                              
   BEGIN                                              
    SET @IGV_A = @varPla_cCta                             
                              
    print '*****************************************************************************'                          
    print @IGV_A                          
    print '*****************************************************************************'                              
                                                     
   END                                              
  ELSE                                       BEGIN                                              
    SET @BASE_A = @varPla_cCta                                                
   END                                              
  IF @NroRegConfigImp <> 0 AND (@varAsd_cBaseImp= '006') -- Configurada como Cuenta de Impuesto                                                          
   BEGIN                                                 
    IF ( @varAsd_cBaseImp= '006' )                                              
    BEGIN                                                    
     IF LEFT(@varPla_cCta,1)='6' OR LEFT(@varPla_cCta,1)='9'                                                
      SET @IGV_A = @varPla_cCta                                                
    END                                                          
   END                                                    
 END                      
                                                           
  -- BUSCANDO CUENTAS DE LA COLUMNA B  
                                                  
  IF (@varAsd_cBaseImp= '007')                                              
  BEGIN                         IF (LEFT(@varPla_cCta,2)='40' and @NroRegConfigImp <> 0 ) or (LEFT(@varPla_cCta,2)='64' and @NroRegConfigImp <> 0)                                                   
    BEGIN                                              
     SET @IGV_B= @varPla_cCta                                              
    END                                              
   ELSE                                              
    BEGIN                                              
     SET @BASE_B= @varPla_cCta                                              
    END                                              
   IF @NroRegConfigImp <> 0 AND (@varAsd_cBaseImp= '007') -- Configurada como Cuenta de Impuesto                                                          
    BEGIN                             
  IF ( @varAsd_cBaseImp= '007' )                                              
      BEGIN                                                    
       IF LEFT(@varPla_cCta,1)='6' OR LEFT(@varPla_cCta,1)='9'                                                
       SET @IGV_B = @varPla_cCta                                                
      END                           
    END                                                        
  END            
                             
 Print 'BASE IMPONIBLE '                            
 Print @varAsd_cBaseImp                            
 Print 'CUENTA '                            
 Print @varPla_cCta                            
 Print '>>>>>>>>>>>>>>'                            
 Print '>>>>>>>>>>>>>>'                            
                             
  -- BUSCANDO CUENTAS DE LA COLUMNA C                                                        
 IF ( @varAsd_cBaseImp= '008' )                                              
 BEGIN                                                        
  IF (LEFT(@varPla_cCta,2)='40' OR (LEFT(@varPla_cCta,1)='6') And @NroRegConfigImp = 1)                                
   BEGIN                                             
    SET @IGV_C= @varPla_cCta                                                    
   END                                              
  ELSE                                                        
   BEGIN                                              
    SET @BASE_C= @varPla_cCta                                       
   END                                     
   Print 'CTA. ' + @BASE_C                                
   Print 'IGV. ' + @IGV_C                            
  IF (@NroRegConfigImp <> 0) AND (@varAsd_cBaseImp= '099') OR (@varAsd_cBaseImp= '008') -- Configurada como Cuenta de Impuesto                                                          
  BEGIN                                              
   IF (( @varAsd_cBaseImp = '008' ) OR ( @varAsd_cBaseImp= '099' )) And @NroRegConfigImp <> 0                                
   BEGIN                                                    
    IF LEFT(@varPla_cCta,1)='6' OR LEFT(@varPla_cCta,1)='9'                                                
    SET @IGV_C = @varPla_cCta                                              
   END                                              
  END                       
 END                      
                                               
    -- ACUMULANDO MONTO INAFECTO                                                         
    IF ( @varAsd_cBaseImp= '999')                                                         
    BEGIN                                                        
     SET @MONTOINAFECTO_S =  isnull(@MONTOINAFECTO_S,0) + ( isnull(@varAsd_nDebeSoles,0) - isnull(@varAsd_nHaberSoles,0) )                                                        
     SET @MONTOINAFECTO_D =  isnull(@MONTOINAFECTO_D,0) + ( isnull(@varAsd_nDebeMonExt,0) - isnull(@varAsd_nHaberMonExt,0) )                                                        
    END                                                        
                
    -- ACUMULANDO IMPORTES DE IGV A                                                       
    IF @IGV_A <> '' and  @varAsd_cBaseImp<> '026' AND @varPla_cCta NOT IN (SELECT CUENTAS FROM #TMP_ISC)                                                        
    BEGIN                                                     
                                  
     -- SELECT @NroRegConfigImp = COUNT(1) FROM CND_CONFIG_OPERA WITH(READUNCOMMITTED) WHERE Emp_cCodigo = @Emp_cCodigo  AND Pan_cAnio = @Pan_cAnio  AND Cod_cValorParam = @varPla_cCta AND Cop_cCodigo = '099'                                               
  
         
     IF @NroRegConfigImp <> 0 -- Configurada como Cuenta de Impuesto                                                      
      BEGIN                                                      
       SET @MontoIgvA_S = @MontoIgvA_S + (@varAsd_nDebeSoles - @varAsd_nHaberSoles )                                        
SET @MontoIgvA_D = @MontoIgvA_D + (@varAsd_nDebeMonExt - @varAsd_nHaberMonExt )                                                        
       SET @CTAREG_IGVA = @varPla_cCta                                                        
       SET @MontoTotalIgv_S = @MontoTotalIgv_S + (@varAsd_nDebeSoles - @varAsd_nHaberSoles )                                                        
       SET @MontoTotalIgv_D = @MontoTotalIgv_D + (@varAsd_nDebeMonExt - @varAsd_nHaberMonExt )                                                        
       SET @nPorcentajeIgv =@nIGV                                                      
      END                                                      
    END                                                       
                                                          
    -- ACUMULANDO IMPORTES DE IGV C                                                        
    IF @IGV_C <> '' and  @varAsd_cBaseImp<> '026' AND @varPla_cCta NOT IN (SELECT CUENTAS FROM #TMP_ISC)                                                              
    BEGIN                                                    
     IF @NroRegConfigImp <> 0 -- Configurada como Cuenta de Impuesto                                                          
  BEGIN                                                       
       SET @MontoIgvC_S = @MontoIgvC_S + (@varAsd_nDebeSoles - @varAsd_nHaberSoles )                                                        
  SET @MontoIgvC_D = @MontoIgvC_D + (@varAsd_nDebeMonExt - @varAsd_nHaberMonExt )                                                        
       SET @CTAREG_IGVC = @varPla_cCta                                                        
       SET @MontoTotalIgv_S = @MontoTotalIgv_S + (@varAsd_nDebeSoles - @varAsd_nHaberSoles )                                                        
       SET @MontoTotalIgv_D = @MontoTotalIgv_D + (@varAsd_nDebeMonExt - @varAsd_nHaberMonExt )                                                        
       SET @nPorcentajeIgv = @nIGV                                                        
     END                                      
       IF  @varPla_cCta IN (SELECT CUENTAS FROM #TMP_REINTEGRO)                                                        
       BEGIN                                                    
        SET @MontoSReintegro = @MontoSReintegro + (@varAsd_nDebeSoles - @varAsd_nHaberSoles )                                                        
        SET @MontoDReintegro = @MontoDReintegro + (@varAsd_nDebeMonExt - @varAsd_nHaberMonExt )                                                        
       END                                              
    END                                                      
    Print 'CTA. 1 ' + @BASE_C                            
    -- ACUMULANDO IMPORTES DE IGV B                                                        
    IF @IGV_B <> '' and @varAsd_cBaseImp<> '026' AND @varPla_cCta NOT IN (SELECT CUENTAS FROM #TMP_ISC)                                                              
    BEGIN                                               
     IF @NroRegConfigImp <> 0 -- Configurada como Cuenta de Impuesto                              
      BEGIN                                                          
       SET @MontoIgvB_S = @MontoIgvB_S + (@varAsd_nDebeSoles - @varAsd_nHaberSoles )                                                        
       SET @MontoIgvB_D = @MontoIgvB_D + (@varAsd_nDebeMonExt - @varAsd_nHaberMonExt )                                                        
       SET @CTAREG_IGVB = @varPla_cCta                     
       SET @MontoTotalIgv_S = @MontoTotalIgv_S + (@varAsd_nDebeSoles - @varAsd_nHaberSoles )                                       
       SET @MontoTotalIgv_D = @MontoTotalIgv_D + (@varAsd_nDebeMonExt - @varAsd_nHaberMonExt )                                                        
       SET @nPorcentajeIgv = @nIGV                                                        
     END                                           
    END                                                          
    Print 'CTA. 2 ' + @BASE_C                                                    
    -- BUSCANDO CTAS DE DIFERENCIA CAMBIO                                                       
    IF @varPla_cCta = @CTA_DIFCP                                                         
    BEGIN                                                        
     SET @MontoDifCam_S = @MontoDifCam_S + (@varAsd_nDebeSoles - @varAsd_nHaberSoles )                                                         
     SET @MontoDifCam_D = @MontoDifCam_D + (@varAsd_nDebeMonExt - @varAsd_nHaberMonExt )                      
--     SET @MontoDifCam_S = @MontoDifCam_S + (@varAsd_nHaberSoles  - @varAsd_nDebeSoles)                    
--     SET @MontoDifCam_D = @MontoDifCam_D + (@varAsd_nHaberMonExt - @varAsd_nDebeMonExt)                    
                                                            
    END                    
    IF @varPla_cCta = @CTA_DIFCG                    
    BEGIN                                                        
     SET @MontoDifCam_S = @MontoDifCam_S + (@varAsd_nDebeSoles - @varAsd_nHaberSoles)                                                        
     SET @MontoDifCam_D = @MontoDifCam_D + (@varAsd_nDebeMonExt - @varAsd_nHaberMonExt)                    
--     SET @MontoDifCam_S = @MontoDifCam_S + (@varAsd_nHaberSoles - @varAsd_nDebeSoles)                    
--     SET @MontoDifCam_D = @MontoDifCam_D + (@varAsd_nHaberMonExt - @varAsd_nDebeMonExt )                    
    END                                                        
                    
/*    -- ACUMULANDO IMPORTE DE OTROS                                                        
    IF @varPla_cCta = isnull(@OTROS,'') OR @varAsd_cBaseImp= '024'                    
    BEGIN                    
     SET @MontoOtros_S = @MontoOtros_S + (@varAsd_nDebeSoles - @varAsd_nHaberSoles )                    
     SET @MontoOtros_D = @MontoOtros_D + (@varAsd_nDebeMonExt - @varAsd_nHaberMonExt )                    
     SET @MontoOtros_S = @MontoOtros_S + (@varAsd_nHaberSoles - @varAsd_nDebeSoles )                   
     SET @MontoOtros_D = @MontoOtros_D + (@varAsd_nHaberMonExt - @varAsd_nDebeMonExt)                    
    END                    
                    
    -- BUSCANDO CTAS DE REDONDEO                                                        
    IF @varPla_cCta = @CTA_RDEOP  and @varAsd_cBaseImp ='024'  
    BEGIN                
     SET @MontoOtros_S = @MontoOtros_S + (@varAsd_nDebeSoles - @varAsd_nHaberSoles )                                                        
     SET @MontoOtros_D = @MontoOtros_D + (@varAsd_nDebeMonExt - @varAsd_nHaberMonExt )                                                     
--     SET @MontoOtros_S = @MontoOtros_S + (@varAsd_nHaberSoles - @varAsd_nDebeSoles )                    
--     SET @MontoOtros_D = @MontoOtros_D + (@varAsd_nHaberMonExt - @varAsd_nDebeMonExt )                    
    END                    
                                                        
    IF @varPla_cCta = @CTA_RDEOG  and @varAsd_cBaseImp ='024'  
    BEGIN                                                   
     SET @MontoOtros_S = @MontoOtros_S + (@varAsd_nDebeSoles - @varAsd_nHaberSoles)                    
     SET @MontoOtros_D = @MontoOtros_D + (@varAsd_nDebeMonExt - @varAsd_nHaberMonExt)                    
--     SET @MontoOtros_S = @MontoOtros_S + (@varAsd_nHaberSoles - @varAsd_nDebeSoles)                    
--     SET @MontoOtros_D = @MontoOtros_D + (@varAsd_nHaberMonExt - @varAsd_nDebeMonExt)                    
                                    
    END                                                        
*/      


    IF @varAsd_cBaseImp ='024'  
    BEGIN                                                   
     SET @MontoOtros_S = @MontoOtros_S + (@varAsd_nDebeSoles - @varAsd_nHaberSoles)                    
     SET @MontoOtros_D = @MontoOtros_D + (@varAsd_nDebeMonExt - @varAsd_nHaberMonExt)                    
	END

 
    -- BUSCANDO CTAS DE ISC                                                        
    IF @ISC = 'ISC'                                                        
    BEGIN                                                        
     IF  @varPla_cCta IN (SELECT CUENTAS FROM #TMP_ISC)                                                        
     BEGIN                            
      SET @MontoSISC_S = @MontoSISC_S + (@varAsd_nDebeSoles - @varAsd_nHaberSoles)                                                        
      SET @MontoSISC_D = @MontoSISC_D + (@varAsd_nDebeMonExt - @varAsd_nHaberMonExt)                                                        
     END                            
    END                            
                                                        
    -- BUSCANDO BASES IMPONIBLES                                                        
    SET @cCadena = ISNULL(@BASE_A,'') + ISNULL(@BASE_B,'') + ISNULL(@BASE_C,'')                                                        
                            
 Print 'CTA. 3 ' + @BASE_C                            
 Print @varAsd_nMontoInafecto                            
    IF DBO.TRIMSQL(@cCadena) <> '' --AND @varAsd_nMontoInafecto <> 1                              
  AND                                 
       @varPla_cCta <> @CTA_DIFCP AND @varPla_cCta <> @CTA_DIFCG AND                        
       @varPla_cCta <> @CTA_RDEOP AND @varPla_cCta <> @CTA_RDEOG AND                                                        
        LEFT (@varPla_cCta,2) <>'40'                            
    BEGIN                                                     
-- CUANDO EXISTE SOLO A                                                    
    IF @BASE_A <> ''                                      
    BEGIN                                                    
     SET @MontoBaseA_S = @MontoBaseA_S + (@varAsd_nDebeSoles - @varAsd_nHaberSoles)                                                    
     SET @MontoBaseA_D = @MontoBaseA_D + (@varAsd_nDebeMonExt - @varAsd_nHaberMonExt)                                                    
                                   
    END                                      
    -- CUANDO EXISTE SOLO B                                                        
    IF @BASE_B <> ''                                                              
 BEGIN                                                    
  SET @MontoBaseB_S = @MontoBaseB_S + (@varAsd_nDebeSoles - @varAsd_nHaberSoles)                                                        
  SET @MontoBaseB_D = @MontoBaseB_D + (@varAsd_nDebeMonExt - @varAsd_nHaberMonExt)                                              
 END                                                        
    -- CUANDO EXISTE SOLO C                            
    Print 'CTA. 4 ' + @BASE_C                            
  Print 'PARTE 2 '                            
    Print 'CUENTA: '                            
    Print @BASE_C                            
    Print @varAsd_nDebeSoles                            
 Print @varAsd_nHaberSoles                            
                             
    IF @BASE_C <> ''                                 
  BEGIN                            
   SET @MontoBaseC_S = @MontoBaseC_S + (@varAsd_nDebeSoles - @varAsd_nHaberSoles)                                
   SET @MontoBaseC_D = @MontoBaseC_D + (@varAsd_nDebeMonExt - @varAsd_nHaberMonExt)                                 
  END                            
  Print @MontoBaseC_S                            
  if @IGV_C <> ''  Print (@varAsd_nDebeSoles - @varAsd_nHaberSoles)                                                      
                               
 END                                                    
                                                        
    FETCH NEXT FROM Compras_Cursor_Detalle                                                        
    INTO @varAsd_nItem, @varEnt_cCodEntidad, @varAsd_cTipoDoc, @varAsd_cSerieDoc, @varAsd_cNumDoc,                                                         
     @varPla_cCta, @varPla_cNomCta, @varAsd_nDebeSoles, @varAsd_nHaberSoles, @varAsd_nDebeMonExt,                                                         
     @varAsd_nHaberMonExt, @varAsd_nMontoInafecto, @imp_nPorcentaje, @varAsd_nTipoCambio,                                                        
     @varAsd_cBaseImp, @x_varAsd_cBaseImp                                       
              
    set @vFETCH_STATUS_DET = @@FETCH_STATUS              
              
   END                                                        
                                                        
                                                        
  CLOSE Compras_Cursor_Detalle                                                        
  DEALLOCATE Compras_Cursor_Detalle                          
                      
  ---------------------------------------------------------------------------------------------                                 
  SET @CONSTANTE_NC = 1                                                        
              
  Print 'BASE C: '                             
  Print @varAse_nVoucher                            
  Print @MontoBaseC_S                            
              
  UPDATE #TMP_REGLECOMPRAS              
  SET CtaBaseA = 'CTA', NombreCtaBaseA = 'DESCRIP',              
   CtaBaseB = 'CTA', NombreCtaBaseB = 'DESCRIP',              
   CtaBaseC = 'CTA', NombreCtaBaseC = 'DESCRIP',              
   CtaIgvA = 'CTA', NombreIgvA = 'DESCRIP',              
   CtaIgvB = 'CTA', NombreIgvB = 'DESCRIP',              
   CtaIgvC = 'CTA', NombreIgvC = 'DESCRIP',              
   CtaProv = 'CTA', NombreProv = 'DESCRIP',              
   MontoSBaseA = (CASE WHEN @cTipoDocReg=@cTipDocNC THEN (@MontoBaseA_S * @CONSTANTE_NC) ELSE @MontoBaseA_S END),              
   MontoDBaseA = (CASE WHEN @cTipoDocReg=@cTipDocNC THEN (@MontoBaseA_D * @CONSTANTE_NC) ELSE @MontoBaseA_D END),                                                         
   MontoSBaseB = (CASE WHEN @cTipoDocReg=@cTipDocNC THEN (@MontoBaseB_S * @CONSTANTE_NC) ELSE @MontoBaseB_S END),                                                         
   MontoDBaseB = (CASE WHEN @cTipoDocReg=@cTipDocNC THEN (@MontoBaseB_D * @CONSTANTE_NC) ELSE @MontoBaseB_D END),                                                         
   MontoSBaseC = (CASE WHEN @cTipoDocReg=@cTipDocNC THEN (@MontoBaseC_S * @CONSTANTE_NC) ELSE @MontoBaseC_S END),                                                         
   MontoDBaseC = (CASE WHEN @cTipoDocReg=@cTipDocNC THEN (@MontoBaseC_D * @CONSTANTE_NC) ELSE @MontoBaseC_D END),                                                   
                                       
   MontoSReintegro = (CASE WHEN @cTipoDocReg=@cTipDocNC THEN (@MontoSReintegro * @CONSTANTE_NC) ELSE @MontoSReintegro END),                                                         
   MontoDReintegro = (CASE WHEN @cTipoDocReg=@cTipDocNC THEN (@MontoDReintegro * @CONSTANTE_NC) ELSE @MontoDReintegro END),                                                         
                
   Asd_nMontoInafecto = (CASE WHEN @Mon_cMNac='1' THEN @MONTOINAFECTO_S ELSE @MONTOINAFECTO_D END),                                                        
                
   MontoSIgvA = (CASE WHEN @cTipoDocReg=@cTipDocNC THEN (@MontoIgvA_S * @CONSTANTE_NC) ELSE @MontoIgvA_S END),                                                      
   MontoDIgvA = (CASE WHEN @cTipoDocReg=@cTipDocNC THEN (@MontoIgvA_D * @CONSTANTE_NC) ELSE @MontoIgvA_D END),                                                         
   MontoSIgvB = (CASE WHEN @cTipoDocReg=@cTipDocNC THEN (@MontoIgvB_S * @CONSTANTE_NC) ELSE @MontoIgvB_S END),                                                         
   MontoDIgvB = (CASE WHEN @cTipoDocReg=@cTipDocNC THEN (@MontoIgvB_D * @CONSTANTE_NC) ELSE @MontoIgvB_D END),                                                         
   MontoSIgvC = (CASE WHEN @cTipoDocReg=@cTipDocNC THEN (@MontoIgvC_S * @CONSTANTE_NC) ELSE @MontoIgvC_S END),                                                         
   MontoDIgvC = (CASE WHEN @cTipoDocReg=@cTipDocNC THEN (@MontoIgvC_D * @CONSTANTE_NC) ELSE @MontoIgvC_D END),                                                         
   MontoSOtros = @MontoOtros_S, --abs(@MontoOtros_S) /*- abs(@MontoDifCam_S)*/,                
   MontoDOtros = @MontoOtros_D, --abs(@MontoOtros_D) /*- abs(@MontoDifCam_D)*/,                
   MontoSDIFC  = 0,                
   MontoDDIFC  = 0,                
                
   MontoSISC = (CASE WHEN @cTipoDocReg=@cTipDocNC THEN (@MontoSISC_S * @CONSTANTE_NC) ELSE @MontoSISC_S END),                                                         
   MontoDISC = (CASE WHEN @cTipoDocReg=@cTipDocNC THEN (@MontoSISC_D * @CONSTANTE_NC) ELSE @MontoSISC_D END),                                                         
                                                        
   MontoSProv = @MontoBaseA_S + @MontoBaseB_S + @MontoBaseC_S + @MONTOINAFECTO_S + @MontoIgvA_S + @MontoIgvB_S + @MontoIgvC_S + @MontoOtros_S + @MontoDifCam_S + @MontoSISC_S,                                                        
   MontoDProv = @MontoBaseA_D + @MontoBaseB_D + @MontoBaseC_D + @MONTOINAFECTO_D + @MontoIgvA_D + @MontoIgvB_D + @MontoIgvC_D + @MontoOtros_D + @MontoDifCam_D + @MontoSISC_D,                                                        
   DifCambio = abs(@MontoDifCam_S),                    
   Imp_nPorcentaje =  @nPorcentajeIgv                                                        
  WHERE Ase_cNumMov=@varAse_cNumMov AND Ase_nVoucher=@varAse_nVoucher                                                                             
                                                        
  FETCH NEXT FROM Compras_Cursor_CabeceraLE                                                         
  INTO @varLib_cTipoLibro, @varAse_cNumMov, @varAse_nVoucher, @TipMonedaCab                    
                                
  set @vFETCH_STATUS_CAB = @@FETCH_STATUS                                                         
 END                                                        
CLOSE Compras_Cursor_CabeceraLE                                           
DEALLOCATE Compras_Cursor_CabeceraLE                       
                    
DECLARE @Separador varchar(1)                        
                        
 SELECT @Separador = '|'                      
                                          
---------------------------------------------------------------------------------------------                                                        
UPDATE #TMP_REGLECOMPRAS SET ANIO = '' WHERE NOT Asd_cTipoDoc IN(SELECT TDO_CCODIGO FROM CNT_TIPODOC WHERE EMP_CCODIGO= @Emp_cCodigo AND TDO_CNOMBRELARGO LIKE '%POLIZA%IMP%')                                                        
---------------------------------------------------------------------------------------------                                                        
                    
--SELECT                 
--convert(varchar(4),year(#TMP_REGLECOMPRAS.Ase_dfecha)) + #TMP_REGLECOMPRAS.Per_cPeriodo + '00'  as 'Per_cPeriodo' ,                    
--#TMP_REGLECOMPRAS.Ase_nVoucher ,                    
--case when year(Asd_dFecDoc)=1900 then null else convert(varchar(10),Asd_dFecDoc,103)   end as 'Asd_dFecDoc',                
--#TMP_REGLECOMPRAS.Ase_cNumMov,                  
--case when #TMP_REGLECOMPRAS.Asd_cTipoDoc IN('14') then convert(varchar(10),Asd_dFecVen,103)                    
--else '' end as 'Asd_dFecVen',                    
--#TMP_REGLECOMPRAS.Asd_cTipoDoc  ,                    
--#TMP_REGLECOMPRAS.Asd_cSerieDoc ,                    
--case when #TMP_REGLECOMPRAS.Asd_cTipoDoc IN ('50','52')then year(Asd_dFecDoc) else 0 end as 'AnioDUADSI' ,                    
--#TMP_REGLECOMPRAS.Asd_cNumDoc ,                    
--'0' as 'Campo9' ,                    
--case when TABLA.Tab_cCodSunat = '' then '0' else TABLA.Tab_cCodSunat end as 'Tab_cCodSunat' ,                    
--#TMP_REGLECOMPRAS.Ent_nRuc ,                    
--#TMP_REGLECOMPRAS.Ent_cPersona ,                    
--#TMP_REGLECOMPRAS.MontoSBaseA ,                    
--#TMP_REGLECOMPRAS.MontoSIgvA ,                    
--#TMP_REGLECOMPRAS.MontoSBaseB ,                    
--#TMP_REGLECOMPRAS.MontoSIgvB ,                    
--#TMP_REGLECOMPRAS.MontoSBaseC ,                    
--#TMP_REGLECOMPRAS.MontoSIgvC ,                    
--#TMP_REGLECOMPRAS.Asd_nMontoInafecto ,                    
--MontoSISC,                     
--#TMP_REGLECOMPRAS.MontoSOtros,                    
--#TMP_REGLECOMPRAS.MontoSProv ,                    
--#TMP_REGLECOMPRAS.Asd_nTipoCambio ,                 
--case when #TMP_REGLECOMPRAS.Asd_cTipoDoc IN('07','08','87','88','97','98') then convert(varchar(10),#TMP_REGLECOMPRAS.Asd_dFecDocRef,103) else '01/01/0001'  end as 'Asd_dFecDocRef',                    
--case when #TMP_REGLECOMPRAS.Asd_cTipoDoc IN('07','08','87','88','97','98') then #TMP_REGLECOMPRAS.Asd_cTipoDocRef else '00' end as 'Asd_cTipoDocRef' ,                    
--case when #TMP_REGLECOMPRAS.Asd_cSerieDocRef = '' then '-' else #TMP_REGLECOMPRAS.Asd_cSerieDocRef end as 'Asd_cSerieDocRef' ,                    
--case when Asd_cTipoDocref IN ('50','52') then right(LTRIM(RTRIM(#TMP_REGLECOMPRAS.Asd_cSerieDocRef )),3) else ''  end as 'CodDUADSI',              
--case when #TMP_REGLECOMPRAS.Asd_cTipoDoc IN('07','08','87','88','97','98') then #TMP_REGLECOMPRAS.Asd_cNumDocRef else '-' end as 'Asd_cNumDocRef' ,                    
--case when #TMP_REGLECOMPRAS.Asd_cTipoDoc IN('91','97','98') then Asd_cComprob else '-' end as 'Asd_cComprob' ,                    
--case when #TMP_REGLECOMPRAS.Asd_dFechaSpot ='' then '01/01/0001' else convert(varchar(10),#TMP_REGLECOMPRAS.Asd_dFechaSpot,103) end as 'Asd_dFechaSpot',                    
--case when #TMP_REGLECOMPRAS.ASD_CNUMSPOT = '' then '0' else #TMP_REGLECOMPRAS.ASD_CNUMSPOT end as 'ASD_CNUMSPOT',                    
--case when #TMP_REGLECOMPRAS.asd_cRetencion = 'R' then '1' else '0' end AS 'Retencion' ,                    
--case when #TMP_REGLECOMPRAS.Asd_cEstadoD <> '9' then #TMP_REGLECOMPRAS.Asd_cEstadoO else #TMP_REGLECOMPRAS.Asd_cEstadoD end as 'Estado',                    
--DifCambio                    
--into #TMPREGISTROLECOMPRAS                           
--FROM #TMP_REGLECOMPRAS                           
                                                        
--LEFT JOIN CNM_ENTIDAD  ON                           
--#TMP_REGLECOMPRAS.Emp_cCodigo = CNM_ENTIDAD.Emp_cCodigo AND                            
--#TMP_REGLECOMPRAS.Ten_cTipoEntidad = CNM_ENTIDAD.Ten_cTipoEntidad AND                                                         
--#TMP_REGLECOMPRAS.Ent_cCodEntidad = CNM_ENTIDAD.Ent_cCodEntidad                                                         
                                                        
--LEFT JOIN TABLA  ON                                                        
--CNM_ENTIDAD.EMP_CCODIGO = TABLA.EMP_CCODIGO AND                                                        
--CNM_ENTIDAD.ENT_CTIPODOC  = TABLA.TAB_CCODIGO                                                          
          
--WHERE                            
--TABLA.TAB_CTABLA = '003'                                    
--group by                       
--#TMP_REGLECOMPRAS.Ase_dfecha,                                                      
--#TMP_REGLECOMPRAS.Emp_cCodigo,                                                         
--#TMP_REGLECOMPRAS.Ase_cNumMov,                                                         
--#TMP_REGLECOMPRAS.Per_cPeriodo,                                                       
--#TMP_REGLECOMPRAS.Ase_nVoucher,                                                         
--#TMP_REGLECOMPRAS.ANIO,                                                         
--#TMP_REGLECOMPRAS.Asd_nItem,                                                         
--#TMP_REGLECOMPRAS.Asd_nTipoCambio,                                                         
--#TMP_REGLECOMPRAS.Asd_cGlosa,                                    
--#TMP_REGLECOMPRAS.Ent_cCodEntidad,                                                         
--#TMP_REGLECOMPRAS.Ent_nRuc,                                                         
--#TMP_REGLECOMPRAS.Ent_cPersona,                                                  
--#TMP_REGLECOMPRAS.Asd_cTipoDoc,                                                         
--#TMP_REGLECOMPRAS.Asd_cSerieDoc,                                                         
--#TMP_REGLECOMPRAS.Asd_cNumDoc,                                                         
--#TMP_REGLECOMPRAS.Asd_dFecDoc,                                                         
--#TMP_REGLECOMPRAS.Asd_cTipoDocRef,                                                         
--#TMP_REGLECOMPRAS.Asd_cSerieDocRef,                                                         
--#TMP_REGLECOMPRAS.Asd_cNumDocRef,                                                         
--#TMP_REGLECOMPRAS.Asd_dFecDocRef,                                                         
--#TMP_REGLECOMPRAS.Asd_nMontoInafecto,                                                         
--#TMP_REGLECOMPRAS.Asd_cRetencion,                                                         
--#TMP_REGLECOMPRAS.Tdo_cNombreLargo,                                                        
--#TMP_REGLECOMPRAS.Ase_cTipoMoneda,                                                        
--#TMP_REGLECOMPRAS.CtaBaseA,                                                         
--#TMP_REGLECOMPRAS.NombreCtaBaseA,                                                         
--#TMP_REGLECOMPRAS.CtaBaseB,                                                         
--#TMP_REGLECOMPRAS.NombreCtaBaseB,                                                         
--#TMP_REGLECOMPRAS.CtaBaseC,                                                         
--#TMP_REGLECOMPRAS.NombreCtaBaseC,                                           
--#TMP_REGLECOMPRAS.CtaIgvA,                                                         
--#TMP_REGLECOMPRAS.NombreIgvA,                                                         
--#TMP_REGLECOMPRAS.CtaIgvB,                                                         
--#TMP_REGLECOMPRAS.NombreIgvB,                                                         
--#TMP_REGLECOMPRAS.CtaIgvC,                                                         
--#TMP_REGLECOMPRAS.NombreIgvC,                                                         
--#TMP_REGLECOMPRAS.CtaProv,                                                         
--#TMP_REGLECOMPRAS.NombreProv,                
--#TMP_REGLECOMPRAS.MontoSBaseA,                                                         
--#TMP_REGLECOMPRAS.MontoDBaseA,                                                         
--#TMP_REGLECOMPRAS.MontoSBaseB,                                                         
--#TMP_REGLECOMPRAS.MontoDBaseB,                                                         
--#TMP_REGLECOMPRAS.MontoSBaseC,                                                         
--#TMP_REGLECOMPRAS.MontoDBaseC,                                                        
--#TMP_REGLECOMPRAS.MontoSIgvA,                                                         
--#TMP_REGLECOMPRAS.MontoDIgvA,                                                         
--#TMP_REGLECOMPRAS.MontoSIgvB,                                   
--#TMP_REGLECOMPRAS.MontoDIgvB,                                     
--#TMP_REGLECOMPRAS.MontoSIgvC,                                                         
--#TMP_REGLECOMPRAS.MontoDIgvC,                                                        
--#TMP_REGLECOMPRAS.MontoSOtros,                                                         
--#TMP_REGLECOMPRAS.MontoDOtros,                                                         
--#TMP_REGLECOMPRAS.MontoSProv,                  
--#TMP_REGLECOMPRAS.MontoDProv,                                                         
--#TMP_REGLECOMPRAS.MontoSReintegro,                                                        
--#TMP_REGLECOMPRAS.MontoDReintegro,                                                        
--#TMP_REGLECOMPRAS.imp_nPorcentaje,                                                          
--#TMP_REGLECOMPRAS.ASD_DFECHASPOT,                                                         
--#TMP_REGLECOMPRAS.ASD_CNUMSPOT,                                              
--#TMP_REGLECOMPRAS.Asd_dFecVen ,                                                        
--#TMP_REGLECOMPRAS.MontoSISC,                                                         
--#TMP_REGLECOMPRAS.MontoDISC,                                                         
--#TMP_REGLECOMPRAS.MontoSDIFC,                                                         
--#TMP_REGLECOMPRAS.MontoDDIFC,                                                  
--#TMP_REGLECOMPRAS.Asd_cCodSunat,                                                         
--#TMP_REGLECOMPRAS.Asd_cComprob,                                                        
--TABLA.Tab_cCodSunat,                                                         
--CNM_ENTIDAD.ENT_CTIPODOC,                    
--#TMP_REGLECOMPRAS.Asd_cEstadoO,                                                 
--#TMP_REGLECOMPRAS.Asd_cEstadoD,                    
--#TMP_REGLECOMPRAS.DifCambio                                                           
--ORDER BY Asd_cTipoDoc,ASE_NVOUCHER,Asd_dFecDoc, ASD_NITEM                                                        
---------------------------------------------------------------------------------------------              
             
             --SELECT * FROM #TMP_REGLECOMPRAS
IF @Ent_cFlagDomiciliado = '3'
BEGIN
	--SELECT * FROM #TMP_REGLECOMPRAS RE
	--WHERE ISNULL(Ent_cFlagDomiciliado, '2') = '2'
	--Order by Per_cPeriodo,Ase_nVoucher
	--RETURN
	SELECT (CAST(YEAR(RE.Ase_dFecha) AS CHAR(4)) + RE.Per_cPeriodo + '00' + @Separador + LEFT(LTRIM(RTRIM(Ase_nVoucher)),40) + @Separador +
		   LEFT(LTRIM(RTRIM('M' + left(Ase_nVoucher,4) + right(Ase_nVoucher,5) )),100) + @Separador +
		   CONVERT(NCHAR(10), Asd_dFecDoc, 103) + @Separador + 
		   CASE WHEN YEAR(Asd_dFecVen) = 1900 THEN '' ELSE CONVERT(NCHAR(10), Asd_dFecVen, 103) END + @Separador +
		   LEFT(LTRIM(RTRIM(Asd_cTipoDoc)),2) + @Separador +
		   CASE WHEN LEFT(LTRIM(RTRIM(Asd_cTipoDoc)),2) IN ('50', '52') THEN RIGHT(RTRIM(Asd_cSerieDoc), 3) ELSE RTRIM(Asd_cSerieDoc) END + @Separador + 		   
		   CASE WHEN (CHARINDEX('-', Asd_cNumDoc) > 0) THEN CASE WHEN LEFT(LTRIM(RTRIM(Asd_cTipoDoc)),2) IN ('50', '52') THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('-', Asd_cNumDoc)), 6) ELSE RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('-', Asd_cNumDoc)), 7) END
		        WHEN (CHARINDEX('/', Asd_cNumDoc) > 0) THEN CASE WHEN LEFT(LTRIM(RTRIM(Asd_cTipoDoc)),2) IN ('50', '52') THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('/', Asd_cNumDoc)), 6) ELSE RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('/', Asd_cNumDoc)), 7) END 
		   ELSE CASE WHEN LEFT(LTRIM(RTRIM(Asd_cTipoDoc)),2) IN ('50', '52') THEN RIGHT(RTRIM(Asd_cNumDoc), 6) ELSE RIGHT(RTRIM(Asd_cNumDoc), 7) END END + @Separador +	
		   	    
		   CASE WHEN (CHARINDEX('-', Asd_cNumDoc) > 0) THEN CASE WHEN LEFT(LTRIM(RTRIM(Asd_cTipoDoc)),2) IN ('50', '52') THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('-', Asd_cNumDoc)), 6) ELSE RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('-', Asd_cNumDoc)), 7) END 
		        WHEN (CHARINDEX('/', Asd_cNumDoc) > 0) THEN CASE WHEN LEFT(LTRIM(RTRIM(Asd_cTipoDoc)),2) IN ('50', '52') THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('/', Asd_cNumDoc)), 6) ELSE RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('/', Asd_cNumDoc)), 7) END
		   ELSE '' END + @Separador  +            
		   Ent_cTipoDoc + @Separador + 
		   Ent_nRuc + @Separador + 
		   Ent_cPersona + @Separador +
		   CAST(CAST(MontoSBaseA + MontoSBaseB + MontoSBaseC AS NUMERIC(18, 2)) AS VARCHAR(50)) + @Separador + 
		   CAST(CAST(MontoSIgvB + MontoSIgvC + MontoSIgvA AS NUMERIC(18, 2)) AS VARCHAR(50)) + @Separador + 
		   CAST(CAST(MontoSOtros AS NUMERIC(18, 2)) AS VARCHAR(50)) + @Separador + 
		   CAST(CAST(MontoSBaseA +MontoSIgvA + MontoSBaseB + MontoSIgvB + MontoSBaseC + MontoSIgvC + MontoSOtros AS NUMERIC(18, 2)) AS VARCHAR(50)) + @Separador + 
		   Mon_cCodSunat + @Separador +
		   CAST(Asd_nTipoCambio AS VARCHAR(50)) + @Separador + 
		   CASE WHEN LEFT(LTRIM(RTRIM(Asd_cTipoDoc)),2) IN ('07', '08', '87', '88', '97', '98') THEN CONVERT(NCHAR(10), Asd_dFecDocRef, 103) ELSE '' END + @Separador +
		   CASE WHEN LEFT(LTRIM(RTRIM(Asd_cTipoDoc)),2) IN ('07', '08', '87', '88', '97', '98') THEN LEFT(LTRIM(RTRIM(Asd_cTipoDocRef)),2) ELSE '' END + @Separador +
		   CASE WHEN LEFT(LTRIM(RTRIM(Asd_cTipoDoc)),2) IN ('07', '08', '87', '88', '97', '98') THEN RTRIM(Asd_cSerieDocRef) ELSE '' END + @Separador +
		   CASE WHEN LEFT(LTRIM(RTRIM(Asd_cTipoDoc)),2) IN ('07', '08', '87', '88', '97', '98') THEN RIGHT(RTRIM(Asd_cNumDocRef), 7) ELSE '' END + @Separador +
		   RTRIM(LTRIM(CASE WHEN YEAR(asd_dFechaSpot) = 1900 THEN '' ELSE CONVERT(NCHAR(10), asd_dFechaSpot, 103) END)) + @Separador + 
		   RTRIM(LTRIM(asd_cnumspot)) + @Separador +
		   case when asd_cRetencion = 'R' then '1' else '' END + @Separador + 
		   RIGHT(RTRIM(LTRIM(Id_Clasific_Servicio)), 1) + @Separador + 
		   '' + @Separador + 
		   '' + @Separador + 
		   '' + @Separador + 
		   '' + @Separador + 
		   case when Asd_cEstadoD <> '9' then Asd_cEstadoO else Asd_cEstadoD END + @Separador) AS Registro  FROM #TMP_REGLECOMPRAS RE
	WHERE LEFT(LTRIM(RTRIM(Asd_cTipoDoc)),2) NOT IN ('02', '04', '07', '08', '09', '20', '21', '25', '31', '33', '34', '35', '40', '41', '44', '48', '49', '89', '91', '97', '98')
	--('01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12', '13', '14', '15', '16', '17', '18', '19', '20', '21', '22',
	--																							 '23', '24', '25', '26', '27', '28', '29', '30', '31', '32', '33', '34', '35', '36', '37', '40', '41', '42', '43', '44', '45', '46',
	--																							 '48', '49', '50', '51', '52', '53', '54', '55', '56', '87', '88', '89', '96')
	Order by Per_cPeriodo,Ase_nVoucher
END
ELSE IF @Ent_cFlagDomiciliado = '2'
BEGIN

	SELECT (CAST(YEAR(RE.Ase_dFecha) AS CHAR(4)) + RE.Per_cPeriodo + '00' + @Separador + LEFT(LTRIM(RTRIM(Ase_nVoucher)),40) + @Separador +
		   LEFT(LTRIM(RTRIM('M' + left(Ase_nVoucher,4) + right(Ase_nVoucher,5) )),100) + @Separador + 
		   CONVERT(NCHAR(10), Asd_dFecDoc, 103) + @Separador + 
		   LEFT(LTRIM(RTRIM(Asd_cTipoDoc)),2) + @Separador + 
		   RTRIM(Asd_cSerieDoc) + @Separador + 
		   RIGHT(RTRIM(Asd_cNumDoc), 7) + @Separador + 
		   CAST(CAST((MontoSBaseA +MontoSIgvA + MontoSBaseB + MontoSIgvB + MontoSBaseC + MontoSIgvC + Asd_nMontoInafecto + MontoSISC) AS NUMERIC(20, 2)) AS VARCHAR(50)) + @Separador +
		   CAST(CAST(MontoSOtros AS NUMERIC(20, 2)) AS VARCHAR(50)) + @Separador +
		   CAST(CAST((MontoSBaseA +MontoSIgvA + MontoSBaseB + MontoSIgvB + MontoSBaseC + MontoSIgvC + Asd_nMontoInafecto + MontoSISC + MontoSOtros) AS NUMERIC(20, 2)) AS VARCHAR(50)) + @Separador +
		   CASE WHEN RTRIM(LTRIM(Asd_cTipoDocRef)) IN ('00', '46', '50', '51', '52', '53') THEN RTRIM(LTRIM(Asd_cTipoDocRef)) ELSE '' END + @Separador +
		   CASE WHEN RTRIM(LTRIM(Asd_cTipoDocRef)) IN ('50', '51', '52', '53') THEN RTRIM(LTRIM(Id_Aduana)) ELSE RTRIM(Asd_cSerieDocRef) END + @Separador +
		   CASE WHEN RTRIM(LTRIM(Asd_cTipoDocRef)) IN ('50', '52') THEN CAST(YEAR(Asd_dFecDocRef) AS VARCHAR(4)) ELSE '' END + @Separador + 
		   CASE WHEN RTRIM(LTRIM(Asd_cTipoDocRef)) IN ('50', '51', '52', '53') THEN RIGHT(RTRIM(Asd_cNumDocRef), 6) ELSE RIGHT(RTRIM(Asd_cNumDocRef), 7) END + @Separador + 
		   '0.00' + @Separador + 
		   RTRIM(LTRIM(Mon_cCodSunat)) + @Separador +
		   CAST(Asd_nTipoCambio AS VARCHAR(50)) + @Separador + 
		   RTRIM(LTRIM(Id_Pais)) + @Separador + 
		   Ent_cPersona + @Separador + 
		   '' + @Separador + 
		   RTRIM(LTRIM(Ent_nRuc)) + @Separador + 
		   RTRIM(LTRIM(@Emp_cNumRuc)) + @Separador + 
		   @NombreEmpresa + @Separador +
		   '' + @Separador + 
		   '' + @Separador + 
		   '0.00' + @Separador + 
		   '0.00' + @Separador + 
		   '0.00' + @Separador + 
		   '0.00' + @Separador + 
		   '0.00' + @Separador + 
		   RTRIM(LTRIM(Id_Convenio)) + @Separador + 
		   RTRIM(Id_Exoneracion) + @Separador +
		   RTRIM(LTRIM(Id_Tipo_Renta)) + @Separador + 
		   RTRIM(LTRIM(Id_Modalidad)) + @Separador + 
		   '' + @Separador + 
		   '0'/*case when Asd_cEstadoD <> '9' then Asd_cEstadoO else Asd_cEstadoD END*/ + @Separador) AS Registro  FROM #TMP_REGLECOMPRAS RE 

	WHERE ISNULL(Ent_cFlagDomiciliado, '1') = '2' AND LEFT(LTRIM(RTRIM(Asd_cTipoDoc)),2) NOT IN ('01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12', '13', '14', '15', '16', '17', '18', '19', '20', '21', '22',
																								 '23', '24', '25', '26', '27', '28', '29', '30', '31', '32', '33', '34', '35', '36', '37', '40', '41', '42', '43', '44', '45', '46',
																								 '48', '49', '50', '51', '52', '53', '54', '55', '56', '87', '88', '89', '96')
	Order by Per_cPeriodo,Ase_nVoucher
END
ELSE IF @Ent_cFlagDomiciliado = '1' 
BEGIN
		
	SELECT (CAST(YEAR(RE.Ase_dFecha) AS CHAR(4)) + RE.Per_cPeriodo + '00' + @Separador + 
	        LEFT(LTRIM(RTRIM(Ase_nVoucher)),40) + @Separador +
		   LEFT(LTRIM(RTRIM('M' + left(Ase_nVoucher,4) + right(Ase_nVoucher,5) )),100) + @Separador + 
		   CONVERT(NCHAR(10), Asd_dFecDoc, 103) + @Separador + 
		   CASE WHEN YEAR(Asd_dFecVen) = '1900' THEN '' ELSE CONVERT(NCHAR(10), Asd_dFecVen, 103) END + @Separador +
		   LEFT(LTRIM(RTRIM(Asd_cTipoDoc)),2) + @Separador +
		   CASE WHEN LEFT(LTRIM(RTRIM(Asd_cTipoDoc)),2) IN ('50', '52', '54') THEN RIGHT(RTRIM(Asd_cSerieDoc), 3) 
			    WHEN LEFT(LTRIM(RTRIM(Asd_cTipoDoc)),2) IN ('05') THEN RIGHT(RTRIM(Asd_cSerieDoc), 1)
		   ELSE RTRIM(Asd_cSerieDoc) END + @Separador + 
		   CASE WHEN LEFT(LTRIM(RTRIM(Asd_cTipoDoc)),2) IN ('50', '52') THEN CAST(YEAR(Asd_dFecDoc) AS CHAR(4)) ELSE '' END + @Separador +
		   
		   CASE WHEN (CHARINDEX('-', Asd_cNumDoc) > 0) THEN CASE WHEN LEFT(LTRIM(RTRIM(Asd_cTipoDoc)),2) IN ('50', '52') THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('-', Asd_cNumDoc)), 6) ELSE RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('-', Asd_cNumDoc)), 7) END 
		        WHEN (CHARINDEX('/', Asd_cNumDoc) > 0) THEN CASE WHEN LEFT(LTRIM(RTRIM(Asd_cTipoDoc)),2) IN ('50', '52') THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('/', Asd_cNumDoc)), 6) ELSE RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('/', Asd_cNumDoc)), 7) END 
		   ELSE CASE WHEN LEFT(LTRIM(RTRIM(Asd_cTipoDoc)),2) IN ('50', '52') THEN RIGHT(RTRIM(Asd_cNumDoc), 6) 
		             WHEN LEFT(LTRIM(RTRIM(Asd_cTipoDoc)),2) IN ('12', '13', '14', '15', '16', '17', '18', '19', '20', '21', '22') THEN RIGHT(RTRIM(Asd_cNumDoc), 20)
		             WHEN LEFT(LTRIM(RTRIM(Asd_cTipoDoc)),2) IN ('05') THEN RIGHT(RTRIM(Asd_cNumDoc), 11)
		             ELSE RIGHT(RTRIM(Asd_cNumDoc), 7) END END + @Separador +		   
		   CASE WHEN (CHARINDEX('-', Asd_cNumDoc) > 0) THEN CASE WHEN LEFT(LTRIM(RTRIM(Asd_cTipoDoc)),2) IN ('50', '52') THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('-', Asd_cNumDoc)), 6) ELSE RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('-', Asd_cNumDoc)), 7) END
		        WHEN (CHARINDEX('/', Asd_cNumDoc) > 0) THEN CASE WHEN LEFT(LTRIM(RTRIM(Asd_cTipoDoc)),2) IN ('50', '52') THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('/', Asd_cNumDoc)), 6) ELSE RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('/', Asd_cNumDoc)), 7) END
		   ELSE '' END + @Separador 
		   
		   + RTRIM(LTRIM(ent_cTipoDoc)) + @Separador +
		   Ent_nRuc + @Separador + Ent_cPersona + @Separador +		   
		   CAST(CAST(MontoSBaseA AS NUMERIC(20, 2)) AS VARCHAR(50)) + @Separador + 
		   CAST(CAST(MontoSIgvA AS NUMERIC(20, 2)) AS VARCHAR(50)) + @Separador + 
		   CAST(CAST(MontoSBaseB AS NUMERIC(20, 2)) AS VARCHAR(50)) + @Separador + 
		   CAST(CAST(MontoSIgvB AS NUMERIC(20, 2)) AS VARCHAR(50)) + @Separador + 
		   CAST(CAST(MontoSBaseC AS NUMERIC(20, 2)) AS VARCHAR(50)) + @Separador + 
		   CAST(CAST(MontoSIgvC AS NUMERIC(20, 2)) AS VARCHAR(50)) + @Separador + 
		   CAST(CAST(Asd_nMontoInafecto AS NUMERIC(20, 2)) AS VARCHAR(50)) + @Separador + 
		   CAST(CAST(MontoSISC AS NUMERIC(20, 2)) AS VARCHAR(50)) + @Separador + 
		   CAST(CAST(MontoSOtros AS NUMERIC(20, 2)) AS VARCHAR(50)) + @Separador + 
		   CAST(CAST(MontoSBaseA +MontoSIgvA + MontoSBaseB + MontoSIgvB + MontoSBaseC + MontoSIgvC + Asd_nMontoInafecto + MontoSISC + MontoSOtros AS NUMERIC(18, 2)) AS VARCHAR(50)) + @Separador + 
		   RTRIM(LTRIM(Mon_cCodSunat))+ @Separador + 
		   CAST(Asd_nTipoCambio AS VARCHAR(50)) + @Separador +
		   CASE WHEN LEFT(LTRIM(RTRIM(Asd_cTipoDoc)),2) IN ('07','08','87','88') THEN CONVERT(NCHAR(10), Asd_dFecDocRef, 103) ELSE '' END + @Separador +
		   CASE WHEN LEFT(LTRIM(RTRIM(Asd_cTipoDoc)),2) IN ('07','08','87','88') THEN LEFT(LTRIM(RTRIM(Asd_cTipoDocRef)),2) ELSE '' END + @Separador +
		   CASE WHEN LEFT(LTRIM(RTRIM(Asd_cTipoDoc)),2) IN ('07','08','87','88') THEN RTRIM(Asd_cSerieDocRef) ELSE '' END + @Separador + 
		   CASE WHEN LEFT(LTRIM(RTRIM(Asd_cTipoDocRef)),2) IN ('50','52') THEN RIGHT(RTRIM(Asd_cSerieDoc), 3) ELSE '' END + @Separador + 		   
		   CASE WHEN LEFT(LTRIM(RTRIM(Asd_cTipoDoc)),2) IN ('07','08','87','88') THEN RIGHT(RTRIM(Asd_cNumDocRef), 7) ELSE '' END + @Separador + 
		   CASE WHEN YEAR(Asd_dFechaSpot) = '1900' THEN '' ELSE CONVERT(NCHAR(10), Asd_dFechaSpot, 103) END + @Separador + 
		   RTRIM(LTRIM(Asd_cNumSpot)) + @Separador +
		   case when asd_cRetencion = 'R' then '1' else '' END + @Separador + 
		   RIGHT(RTRIM(LTRIM(Id_Clasific_Servicio)), 1) + @Separador + 
		   '' + @Separador + 
		   '' + @Separador + 
		   '' + @Separador + 
		   '' + @Separador + 
		   '' + @Separador +
		   '' + @Separador + 
		   case when Asd_cEstadoD <> '9' then Asd_cEstadoO else Asd_cEstadoD END + @Separador) AS Registro  FROM #TMP_REGLECOMPRAS RE
		   
	  WHERE 
	   --AND LEFT(LTRIM(RTRIM(Asd_cTipoDoc)),2) NOT IN ('02', '04', '07', '08', '09', '20', '21', '25', '31', '33', '34', '35', '40', '41', '44', '48', '49', '89', '91', '97', '98')
	    LEFT(LTRIM(RTRIM(Asd_cTipoDoc)),2) NOT IN ('09', '20', '31', '33', '40', '41', '91', '97', '98')
	AND ISNULL(Ent_cFlagDomiciliado, '1') = '1'
	Order by Per_cPeriodo,Ase_nVoucher
END


            /*
            
            RTRIM(LTRIM(Asd_cNumDoc)) NOT IN (SELECT RTRIM(LTRIM(Asd_cNumDoc)) FROM #TMP_REGLECOMPRAS
								WHERE LEFT(LTRIM(RTRIM(Asd_cTipoDoc)),2) IN ('00','03','05','06','11','12','13','14','15','16','18','19','23','26','28','30','36','37','55','56')
									  AND (CHARINDEX('-', Asd_cNumDoc) = 0 AND CHARINDEX('/', Asd_cNumDoc) = 0))
            */  
             
-- select              
-- (              
-- LEFT(LTRIM(RTRIM(Per_cPeriodo)),8) + @Separador +                  
-- LEFT(LTRIM(RTRIM(Ase_nVoucher)),40) + @Separador +                  
-- LEFT(LTRIM(RTRIM('M' + left(Ase_nVoucher,4) + right(Ase_nVoucher,5) )),100) + @Separador +                    
-- LEFT(LTRIM(RTRIM(Asd_dFecDoc)),10) + @Separador +                    
-- LEFT(LTRIM(RTRIM(Asd_dFecVen)),10) + @Separador +      
-- LEFT(LTRIM(RTRIM(Asd_cTipoDoc)),2) + @Separador +                    
-- --(case when Asd_cTipoDoc IN ('50','52') then LTRIM(RTRIM(right(Asd_cSerieDoc,4))) else                    
-- --LTRIM(RTRIM(LEFT(Asd_cSerieDoc,4)))  end) + @Separador +              
-- case when RTRIM(Asd_cTipoDoc) in ('05','55') then      
-- right(RTRIM(Asd_cSerieDoc),1)        
-- when RTRIM(Asd_cTipoDoc) in ('50','51','52','53','54') then      
-- right(RTRIM(Asd_cSerieDoc),3)        
-- else       
-- LTRIM(RTRIM(LEFT(Asd_cSerieDoc,4)))      
-- end + @Separador +      
-- convert(varchar(4),LTRIM(RTRIM(LEFT(AnioDUADSI,4)))) + @Separador +           
-- case when RTRIM(Asd_cTipoDoc) in ('01','02','03','04','06','07','08','23','25','34','35','48','89') then        
-- right(RTRIM(Asd_cNumDoc),7)         
-- when RTRIM(Asd_cTipoDoc) in ('50','51','52','53') then        
-- right(RTRIM(Asd_cNumDoc),6)        
-- else          
-- LTRIM(LEFT(RTRIM(Asd_cNumDoc),20))        
-- end + @Separador +         
-- LTRIM(RTRIM(LEFT(Campo9,20))) + @Separador +                    
-- LTRIM(RTRIM(LEFT(Tab_cCodSunat,1))) + @Separador +                    
-- LTRIM(RTRIM(LEFT(Ent_nRuc,15))) + @Separador +                    
-- LEFT(LTRIM(RTRIM(replace(replace(Ent_cPersona, char(13)+char(10), ' '), char(9), ' ') )),60) + @Separador +                    
-- LTRIM(RTRIM(LEFT(convert(varchar(15),CONVERT(money, MontoSBaseA), 0),15))) + @Separador +                    
-- LTRIM(RTRIM(LEFT(convert(varchar(15),CONVERT(money, MontoSIgvA), 0),15))) + @Separador +                    
-- LTRIM(RTRIM(LEFT(convert(varchar(15),CONVERT(money, MontoSBaseB),0),15))) + @Separador +                    
-- LTRIM(RTRIM(LEFT(convert(varchar(15),CONVERT(money, MontoSIgvB),0),15))) + @Separador +                    
-- LTRIM(RTRIM(LEFT(convert(varchar(15),CONVERT(money, MontoSBaseC),0),15))) + @Separador +                    
-- LTRIM(RTRIM(LEFT(convert(varchar(15),CONVERT(money, MontoSIgvC),0),15))) + @Separador +                    
-- LTRIM(RTRIM(LEFT(convert(varchar(15),CONVERT(money, Asd_nMontoInafecto),0),15))) + @Separador +                    
-- LTRIM(RTRIM(LEFT(convert(varchar(15),CONVERT(money, MontoSISC),0),15))) + @Separador +                    
-- LTRIM(RTRIM(LEFT(convert(varchar(15),CONVERT(money, MontoSOtros),0),15))) + @Separador +                    
-- --LTRIM(RTRIM(LEFT(convert(varchar(15),CONVERT(money, MontoSProv),0),15))) + @Separador  +                    
-- LTRIM(RTRIM(LEFT(convert(varchar(15),CONVERT(money, MontoSBaseA +MontoSIgvA + MontoSBaseB + MontoSIgvB + MontoSBaseC + MontoSIgvC + Asd_nMontoInafecto + MontoSISC + MontoSOtros  ),0),15))) + @Separador  +                  
-- LTRIM(RTRIM(LEFT(convert(varchar(5), Asd_nTipoCambio),5))) + @Separador +                    
-- LTRIM(RTRIM(LEFT(Asd_dFecDocRef,10))) + @Separador +                    
-- LTRIM(RTRIM(LEFT(Asd_cTipoDocRef,2))) + @Separador +                    
-- LTRIM(RTRIM(LEFT(Asd_cSerieDocRef,20))) + @Separador +                
-- CodDUADSI + @Separador +              
-- LTRIM(RTRIM(right(Asd_cNumDocRef,20))) + @Separador +                    
-- LTRIM(RTRIM(LEFT(Asd_cComprob,20))) + @Separador +                    
-- LTRIM(RTRIM(LEFT(Asd_dFechaSpot,10))) + @Separador +                    
-- LTRIM(RTRIM(LEFT(ASD_CNUMSPOT,20))) + @Separador +                    
-- LTRIM(RTRIM(LEFT(Retencion,1))) + @Separador +                    
-- LTRIM(RTRIM(LEFT(Estado,1))) + @Separador +                    
-- --convert(varchar(15),CONVERT(money, Difcambio),0) + @Separador) AS Registro                    
--convert(varchar(15),0,0) + @Separador) AS Registro 
--  from #TMPREGISTROLECOMPRAS                      
 --Order by Per_cPeriodo,Ase_nVoucher
GO
