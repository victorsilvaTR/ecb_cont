SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO
/*-----------------------------------------------------------------------------------------------------------------                                                                                    
MODULO DE CONTABILIDAD                                                                                    
Creador: Pool Berrospi                                                    
Fecha Creaciòn: 07/03/2013                                                    
DESCRIPCION  : Reporte de Libro Diario Electronico                                                    
------------------------------------------------------------------------------------------------------------------*/                                                                                    
--spCn_RptDiarioElectronico2 '015','2014','01','01/01/2013','31/01/2013','038','01', '1'                                              
                                                    
CREATE PROCEDURE [dbo].[spCn_RptDiarioElectronico4]                
 @Emp_cCodigo char(3)='',      
 @Pan_cAnio char(4)='',      
 @Per_cPeriodo char(2)='',      
 @desde varchar(10) = '',      
 @hasta varchar(10) = '',      
 @moneda char(3) = '',      
 @Per_cPeriodoFin char(2)='',      
 @TipoLibDi char(1) = '',
 @Simplificado CHAR(1) = '0'      
--WITH ENCRYPTION                                                    
As                                          
                                          
Declare @ctaini  char(12)                                                     
Declare @ctafin  char(12)                                                     
Declare @Separador varchar(1)                                                        
Declare @Nro int                                             
                                                        
 Set @Separador = '|'                                                      
 set @TipoLibDi = '0'                                                    
 set @ctaini = ''            
 set @ctafin = ''                                             
 set @Nro = 1                                                   
                                                    
                                                    
 if @Per_cPeriodo = '01'                                                    
 set @Per_cPeriodo = '00'                                                    
                                                   
 if @Per_cPeriodo = '12'                                                    
 set @Per_cPeriodoFin = '14'                                                    
                                                     
                                                    
SET DATEFORMAT DMY                                                                                    
SET NOCOUNT ON                                                                                    
declare @RUC char(50)                                                                                    
select @RUC= dbo.RUC(@Emp_cCodigo)                                                                                    
                                                                                    
-- *** Hallando el tipo de Moneda(Si es Nac o No) y su nombre                                                                                    
                                                                                    
Declare @Mon_cNombreLargo varchar(100)                                                                                    
Declare @Mon_cMNac char(1)                                                                                    
                                                                                    
SELECT @Mon_cNombreLargo = Mon_cNombreLargo, @Mon_cMNac = Mon_cMNac                                                                                     
FROM CNT_TIPO_MONEDA                    
WHERE Emp_cCodigo = @Emp_cCodigo and Mon_cCodigo = @moneda                                                      
                         
IF  @ctaini  = ''                                                  
 BEGIN           
  SET @ctaini = '10'                                    
  SET @ctafin = '9999999999'           
END                                                                         
                                                    
SET @ctafin = DBO.TRIMSQL(@ctafin)  + '9999999999'                                          
          
   -- *** SELECCIONAR ASIENTOS POR RANGO DE FECHAS          
  SELECT            
  Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then                                              
  SUBSTRING(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,3,2)            
  Else dbo.CNC_ASIENTO_VOUCHER.Per_cPeriodo End As Per_cPeriodo,            
  cnc_asiento_voucher.Ase_nVoucher as CUO,          
 /* Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' OR dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') and CNC_ASIENTO_VOUCHER.Asd_cEstadoD not IN('8','9') Then                                            
 LEFT(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + 'R00001'            
  when @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' OR dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') and CNC_ASIENTO_VOUCHER.Asd_cEstadoD = '8' Then                                            
 LEFT(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + 'R00002'            
  when @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' OR dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') and CNC_ASIENTO_VOUCHER.Asd_cEstadoD = '9' Then                                            
 LEFT(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + 'R00003'            
  Else*/            
  case when left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) = '0100' then            
  LEFT(LTRIM(RTRIM('A' + left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + right(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,5) )),100)            
  when left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) = '0813' or left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) = '0814' then            
  LEFT(LTRIM(RTRIM('C' + left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + right(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,5) )),100)            
  else            
  LEFT(LTRIM(RTRIM('M' + left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + right(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,5) )),100)            
  end As Correlativo            
  ,            
  /*End As Ase_nVoucher,*/            
  '01' as 'PCGE',            
  dbo.CND_ASIENTO_VOUCHER.Pla_cCuentaContable,            
  Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then                                                                    
  DateAdd(day,-1, DateAdd(Month,1 , DateAdd(day,-(Day(dbo.CNC_ASIENTO_VOUCHER.Ase_dFecha)-1),dbo.CNC_ASIENTO_VOUCHER.Ase_dFecha))) Else                                                                    
  dbo.CNC_ASIENTO_VOUCHER.Ase_dFecha End As Ase_dFecha,                                                
  Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then                                                   
  'CENTRALIZACION REGISTRO DE ' +                                                                          
  CASE WHEN dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Then 'VENTAS ' Else 'COMPRAS ' End + @Per_cPeriodo + ' - ' + @Pan_cAnio Else                                                                     
  dbo.CNC_ASIENTO_VOUCHER.Ase_cGlosa End As Ase_cGlosa,                                                
   Case When @TipoLibDi = '1' And dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Then                                                                   
    /*Case When (Sum(dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles - dbo.CND_ASIENTO_VOUCHER.Asd_nHaberSoles)) > 0 Then                      
     (Sum(dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles - dbo.CND_ASIENTO_VOUCHER.Asd_nHaberSoles))                                                           
    Else                                                           
     0                                               
  End */                                            
      --Sum(dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles)
      dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles                                                           
   Else                                                         
    Case When @TipoLibDi = '1' And dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06' Then                  
     /*Case When (Sum(dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles - dbo.CND_ASIENTO_VOUCHER.Asd_nHaberSoles)) > 0 Then                                                                  
      (Sum(dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles - dbo.CND_ASIENTO_VOUCHER.Asd_nHaberSoles))                                                     
     Else 0                                                     
     End */                                            
     --Sum(dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles)
     dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles                                                           
   Else                                                                  
    --Sum(dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles)   
    dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles                               
   End                                                                   
  End As Asd_nDebeSoles,                                              
  Case When @TipoLibDi = '1' And dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Then                                                               
  /*Case When (Sum(dbo.CND_ASIENTO_VOUCHER.Asd_nHaberSoles - dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles)) > 0 Then                                                                   
  Sum(dbo.CND_ASIENTO_VOUCHER.Asd_nHaberSoles - dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles) Else 0 End */                                            
       --Sum(dbo.CND_ASIENTO_VOUCHER.Asd_nHaberSoles)
       dbo.CND_ASIENTO_VOUCHER.Asd_nHaberSoles                                                            
  Else                                                         
  Case When @TipoLibDi = '1' And dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06' Then                                                                   
  /*CASE WHEN (Sum(dbo.CND_ASIENTO_VOUCHER.Asd_nHaberSoles - dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles)) > 0 Then                                                                  
  Sum(dbo.CND_ASIENTO_VOUCHER.Asd_nHaberSoles - dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles) Else 0 End */                                            
  --Sum(dbo.CND_ASIENTO_VOUCHER.Asd_nHaberSoles)     
	dbo.CND_ASIENTO_VOUCHER.Asd_nHaberSoles                                                            
  Else                                                                  
  --Sum(dbo.CND_ASIENTO_VOUCHER.Asd_nHaberSoles) 
  dbo.CND_ASIENTO_VOUCHER.Asd_nHaberSoles                                                                 
  End   End As Asd_nHaberSoles,                
  '' as CUOVentas,            
  '' as CUOCompras,            
  '' as CUOConsignacion,                    
  isnull(CNC_ASIENTO_VOUCHER.Asd_cEstadoO,'') as 'Asd_cEstadoO',                                              
  isnull(CNC_ASIENTO_VOUCHER.Asd_cEstadoD,'') as 'Asd_cEstadoD', dbo.CNT_TIPO_MONEDA.Mon_cCodSunat, dbo.CND_ASIENTO_VOUCHER.Asd_cSerieDoc, dbo.CND_ASIENTO_VOUCHER.Asd_cNumDoc,
  dbo.CND_ASIENTO_VOUCHER.Asd_dFecVen, dbo.CND_ASIENTO_VOUCHER.Asd_dFecDoc, dbo.CND_ASIENTO_VOUCHER.Asd_cTipoDoc, dbo.CND_ASIENTO_VOUCHER.Lib_cTipoLibro, dbo.CND_ASIENTO_VOUCHER.Ase_nVoucher,
  dbo.CND_ASIENTO_VOUCHER.Ten_cTipoEntidad, dbo.CND_ASIENTO_VOUCHER.Ent_cCodEntidad, dbo.CNM_ENTIDAD.Ent_nRuc, (SELECT t.Tab_cCodSunat FROM dbo.TABLA T
                                                                                                                WHERE T.Emp_cCodigo = @Emp_cCodigo AND T.Tab_cTabla = '003' AND t.Tab_cCodigo = dbo.CNM_ENTIDAD.Ent_cTipoDoc) as Ent_cTipoDoc/*,                    
  case when CND_ASIENTO_VOUCHER.Ent_cCodEntidad <> '' then CND_ASIENTO_VOUCHER.Asd_dFecDoc else '' end 'Asd_dFecDoc'*/                    
 into #TMPDiarioPLE1                                              
 FROM dbo.EMPRESA RIGHT OUTER JOIN                                                                              
                      dbo.CNC_ASIENTO_VOUCHER ON dbo.EMPRESA.Emp_cCodigo = dbo.CNC_ASIENTO_VOUCHER.Emp_cCodigo LEFT OUTER JOIN                     
                      dbo.CND_ASIENTO_VOUCHER LEFT OUTER JOIN                                                 
                      dbo.CNT_TIPODOC CNT_TIPODOC_2 ON dbo.CND_ASIENTO_VOUCHER.Asd_cTipoDocRef = CNT_TIPODOC_2.Tdo_cCodigo AND                                                                    
            dbo.CND_ASIENTO_VOUCHER.Emp_cCodigo = CNT_TIPODOC_2.Emp_cCodigo LEFT OUTER JOIN                                                                              
                      dbo.CNT_TIPODOC CNT_TIPODOC_1 ON dbo.CND_ASIENTO_VOUCHER.Asd_cTipoDoc = CNT_TIPODOC_1.Tdo_cCodigo AND                    
                      dbo.CND_ASIENTO_VOUCHER.Emp_cCodigo = CNT_TIPODOC_1.Emp_cCodigo LEFT OUTER JOIN                                                                              
                      dbo.CNM_ENTIDAD ON dbo.CND_ASIENTO_VOUCHER.Emp_cCodigo = dbo.CNM_ENTIDAD.Emp_cCodigo AND                                                                               
       dbo.CND_ASIENTO_VOUCHER.Ten_cTipoEntidad = dbo.CNM_ENTIDAD.Ten_cTipoEntidad AND                                                                               
                      dbo.CND_ASIENTO_VOUCHER.Ent_cCodEntidad = dbo.CNM_ENTIDAD.Ent_cCodEntidad LEFT OUTER JOIN                              
                      dbo.CNM_PLAN_CTA ON dbo.CND_ASIENTO_VOUCHER.Emp_cCodigo = dbo.CNM_PLAN_CTA.Emp_cCodigo AND                    
                      dbo.CND_ASIENTO_VOUCHER.Pan_cAnio = dbo.CNM_PLAN_CTA.Pan_cAnio AND                                                           
                      dbo.CND_ASIENTO_VOUCHER.Pla_cCuentaContable = dbo.CNM_PLAN_CTA.Pla_cCuentaContable LEFT OUTER JOIN                                                                              
                      dbo.CNT_CENTRO_COSTO ON dbo.CND_ASIENTO_VOUCHER.Emp_cCodigo = dbo.CNT_CENTRO_COSTO.Emp_cCodigo AND                           
                      dbo.CND_ASIENTO_VOUCHER.Pan_cAnio = dbo.CNT_CENTRO_COSTO.Pan_cAnio AND                                                                               
                      dbo.CND_ASIENTO_VOUCHER.Cos_cCodigo = dbo.CNT_CENTRO_COSTO.Cos_cCodigo ON                                                        
                      dbo.CNC_ASIENTO_VOUCHER.Emp_cCodigo = dbo.CND_ASIENTO_VOUCHER.Emp_cCodigo AND                                                                               
                      dbo.CNC_ASIENTO_VOUCHER.Pan_cAnio = dbo.CND_ASIENTO_VOUCHER.Pan_cAnio AND                                                                               
                      dbo.CNC_ASIENTO_VOUCHER.Per_cPeriodo = dbo.CND_ASIENTO_VOUCHER.Per_cPeriodo AND                                                                               
                      dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = dbo.CND_ASIENTO_VOUCHER.Lib_cTipoLibro AND                                                                             
                      dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher = dbo.CND_ASIENTO_VOUCHER.Ase_nVoucher                                                                               
  LEFT OUTER JOIN dbo.CNT_LIBRO_OPERA ON              
  dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = dbo.CNT_LIBRO_OPERA.Lib_cTipoLibro AND              
  dbo.CNC_ASIENTO_VOUCHER.PAN_CANIO =dbo. CNT_LIBRO_OPERA.PAN_CANIO AND              
  dbo.CNC_ASIENTO_VOUCHER.Emp_cCodigo = dbo.CNT_LIBRO_OPERA.Emp_cCodigo              
  LEFT OUTER JOIN  dbo.CNT_TIPO_MONEDA ON              
  dbo.CNC_ASIENTO_VOUCHER.Ase_cTipoMoneda = dbo.CNT_TIPO_MONEDA.Mon_cCodigo AND              
  dbo.CNC_ASIENTO_VOUCHER.Emp_cCodigo = dbo.CNT_TIPO_MONEDA.Emp_cCodigo              
  LEFT OUTER JOIN  dbo.CNT_ENTIDAD ON              
  dbo.CNM_ENTIDAD.Ten_cTipoEntidad = dbo.CNT_ENTIDAD.Ten_cTipoEntidad AND              
  dbo.CNM_ENTIDAD.Emp_cCodigo = dbo.CNT_ENTIDAD.Emp_cCodigo              
 WHERE  CNC_ASIENTO_VOUCHER.Emp_cCodigo = @Emp_cCodigo  AND (dbo.CNC_ASIENTO_VOUCHER.Pan_cAnio = @Pan_cAnio )              
 AND dbo.CNC_ASIENTO_VOUCHER.Per_cPeriodo >= @Per_cPeriodo              
 AND dbo.CNC_ASIENTO_VOUCHER.Per_cPeriodo <= @Per_cPeriodoFin              
 AND (dbo.CNC_ASIENTO_VOUCHER.Ase_cDeleted <> '*') AND (dbo.CND_ASIENTO_VOUCHER.Asd_cDeleted <> '*')              
 AND  dbo.CND_ASIENTO_VOUCHER.Pla_cCuentaContable >= @ctaini              
 and   dbo.CND_ASIENTO_VOUCHER.Pla_cCuentaContable <= @ctafin              
 GROUP BY              
   CNC_ASIENTO_VOUCHER.Asd_cEstadoO, dbo.CNM_ENTIDAD.Ent_nRuc, dbo.CNM_ENTIDAD.Ent_cTipoDoc, dbo.CND_ASIENTO_VOUCHER.Asd_nHaberSoles , dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles , CNC_ASIENTO_VOUCHER.Asd_cEstadoD,              
   dbo.CNC_ASIENTO_VOUCHER.Emp_cCodigo,              
  /*Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' OR dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') and CNC_ASIENTO_VOUCHER.Asd_cEstadoD not IN('8','9') Then         
 LEFT(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + 'R00001'                                             
  when @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' OR dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') and CNC_ASIENTO_VOUCHER.Asd_cEstadoD = '8' Then                                            
 LEFT(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + 'R00002'                                             
  when @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' OR dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') and CNC_ASIENTO_VOUCHER.Asd_cEstadoD = '9' Then                                   
 LEFT(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + 'R00003'                                             
  Else                                             
  dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher                                             
  End*/            
  dbo.cnc_asiento_voucher.Ase_nVoucher,          
  case when left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) = '0100' then            
  LEFT(LTRIM(RTRIM('A' + left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + right(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,5) )),100)            
  when left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) = '0813' or left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) = '0814' then            
  LEFT(LTRIM(RTRIM('C' + left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + right(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,5) )),100)            
  else            
  LEFT(LTRIM(RTRIM('M' + left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + right(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,5) )),100)            
  end ,             
  --cnc_asiento_voucher.Ase_cNummov,                            
   Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                     
   CNC_ASIENTO_VOUCHER.Ase_cNummov End,                                                                     
   Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then                              SUBSTRING(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher, 3 ,2) Else                               
  
    
      
        
          
            
   dbo.CNC_ASIENTO_VOUCHER.Per_cPeriodo End,                                              
   Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then                                                                        
   DateAdd(day,-1, DateAdd(Month,1 , DateAdd(day,-(Day(dbo.CNC_ASIENTO_VOUCHER.Ase_dFecha)-1),dbo.CNC_ASIENTO_VOUCHER.Ase_dFecha))) Else                                                                          
   dbo.CNC_ASIENTO_VOUCHER.Ase_dFecha End,                                              
   dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro, dbo.CNT_LIBRO_OPERA.Lib_cDescripcion,                                              
        Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                     
        dbo.CNC_ASIENTO_VOUCHER.Ase_cTipoMoneda End,                                              
        Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                    
        dbo.CNT_TIPO_MONEDA.Mon_cNombreLargo End, dbo.CNC_ASIENTO_VOUCHER.Ase_nTipoCambio,                                              
  Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then                                                         
  'CENTRALIZACION REGISTRO DE ' +                                              
        CASE WHEN dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Then 'VENTAS ' Else 'COMPRAS ' End + @Per_cPeriodo + ' - ' + @Pan_cAnio ELse                                                                          
        dbo.CNC_ASIENTO_VOUCHER.Ase_cGlosa End,         
        Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then                                             
        0 Else                                              
        dbo.CND_ASIENTO_VOUCHER.Asd_nItem End, dbo.CND_ASIENTO_VOUCHER.Pla_cCuentaContable, dbo.CNM_PLAN_CTA.Pla_cNombreCuenta,                                                                    
        dbo.CNM_PLAN_CTA.Pla_cProvision,                                                                           
        Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then 'CENTRALIZACION REGISTRO DE ' +                                                                          
        CASE WHEN dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Then 'VENTAS ' Else 'COMPRAS ' End + @Per_cPeriodo + ' - ' + @Pan_cAnio  Else                                                                          
        dbo.CND_ASIENTO_VOUCHER.Asd_cGlosa End,                                                                          
        dbo.CND_ASIENTO_VOUCHER.Cos_cCodigo,                                                                               
        dbo.CNT_CENTRO_COSTO.Cos_cDescripcion, dbo.CND_ASIENTO_VOUCHER.Ten_cTipoEntidad,                                                                                                 
  Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                              
        dbo.CNM_ENTIDAD.Ent_nRuc End,                                                    
        Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                     
        dbo.CNM_ENTIDAD.Ent_cPersona End ,                                                                           
        Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                          
        dbo.CNT_ENTIDAD.Ten_cNombreEntidad End,                                                                           
        Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                             
        dbo.CND_ASIENTO_VOUCHER.Asd_cTipoDoc End,                                                                               
        Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                          
        dbo.CND_ASIENTO_VOUCHER.Asd_cSerieDoc End,                                                            
        Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                          
        dbo.CND_ASIENTO_VOUCHER.Asd_cNumDoc End,                                                                           
        /*Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                     
        dbo.CND_ASIENTO_VOUCHER.Asd_dFecDoc End,*/                                                                               
        Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                 
        dbo.CND_ASIENTO_VOUCHER.Asd_cTipoDocRef End ,                                                                           
        Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                     
        dbo.CND_ASIENTO_VOUCHER.Asd_cSerieDocRef End,                                                                           
        Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                          
        dbo.CND_ASIENTO_VOUCHER.Asd_cNumDocRef End,                                                                               
        Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                             
        dbo.CND_ASIENTO_VOUCHER.Asd_dFecDocRef End,                                                                    
        Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then 0 Else                                                                          
        dbo.CND_ASIENTO_VOUCHER.Asd_nMontoInafecto End,                                                                           
        Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                         
        dbo.CND_ASIENTO_VOUCHER.Asd_cRetencion End,                                                                               
        Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                          
        dbo.CND_ASIENTO_VOUCHER.Asd_dFechaSpot End,                                                                           
        Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                          
        dbo.CND_ASIENTO_VOUCHER.Asd_cNumSpot End,                                                                           
        Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                          
        dbo.CND_ASIENTO_VOUCHER.Asd_cDestino End,                                                                                                 
        Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                          
        dbo.CND_ASIENTO_VOUCHER.Asd_nCorre End,                                                                          
        Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                       
        CNT_TIPODOC_1.Tdo_cNombreLargo End,                                                                           
        Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                          
  CNT_TIPODOC_1.Tdo_cNombreCorto End,                        
        Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                          
        CNT_TIPODOC_2.Tdo_cNombreLargo End,                                                                               
        Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                          
        CNT_TIPODOC_2.Tdo_cNombreCorto End,                                                                           
        Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                          
       Ase_cUserModifica End,                                                                               
        Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                        
        dbo.EMPRESA.Emp_cNombreLargo End,                                                                           
        Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                       
        dbo.EMPRESA.Emp_cNombreCorto  End,                                                                          
        Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                          
        Ase_cUserModifica END , dbo.CNT_TIPO_MONEDA.Mon_cCodSunat, dbo.CND_ASIENTO_VOUCHER.Asd_cSerieDoc, dbo.CND_ASIENTO_VOUCHER.Asd_cNumDoc,
  dbo.CND_ASIENTO_VOUCHER.Asd_dFecVen, dbo.CND_ASIENTO_VOUCHER.Asd_dFecDoc, dbo.CND_ASIENTO_VOUCHER.Asd_cTipoDoc, dbo.CND_ASIENTO_VOUCHER.Lib_cTipoLibro, dbo.CND_ASIENTO_VOUCHER.Ase_nVoucher, dbo.CND_ASIENTO_VOUCHER.Ten_cTipoEntidad, dbo.CND_ASIENTO_VOUCHER.Ent_cCodEntidad/*,                                    
        , dbo.CND_ASIENTO_VOUCHER.Ten_cTipoEntidad, dbo.CND_ASIENTO_VOUCHER.Ent_cCodEntidad
        CND_ASIENTO_VOUCHER.Ent_cCodEntidad,CND_ASIENTO_VOUCHER.Asd_dFecDoc*/                                   
  ORDER BY dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles ,                                                    
  Case When @TipoLibDi = '1' And                                                    
  (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then                                                   
  SUBSTRING(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher, 3 ,2)                                                   
  Else  dbo.CNC_ASIENTO_VOUCHER.Per_cPeriodo End,             
                                              
  /*Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' OR dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') and CNC_ASIENTO_VOUCHER.Asd_cEstadoD not IN('8','9') Then                                            
 LEFT(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + 'R00001'                                             
  when @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' OR dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') and CNC_ASIENTO_VOUCHER.Asd_cEstadoD = '8' Then                                            
 LEFT(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + 'R00002'                                             
  when @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' OR dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') and CNC_ASIENTO_VOUCHER.Asd_cEstadoD = '9' Then                                            
 LEFT(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + 'R00003'                                
  Else                                             
  dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher                                             
  End*/           
    case when left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) = '0100' then            
  LEFT(LTRIM(RTRIM('A' + left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + right(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,5) )),100)            
  when left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) = '0813' or left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) = '0814' then            
  LEFT(LTRIM(RTRIM('C' + left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + right(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,5) )),100)            
  else            
  LEFT(LTRIM(RTRIM('M' + left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + right(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,5) )),100)            
  end,              
  --cnc_asiento_voucher.Ase_cNummov,            
  Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then               
   DateAdd(day,-1, DateAdd(Month,1 , DateAdd(day,-(Day(dbo.CNC_ASIENTO_VOUCHER.Ase_dFecha)-1),dbo.CNC_ASIENTO_VOUCHER.Ase_dFecha))) Else                                                                          
   dbo.CNC_ASIENTO_VOUCHER.Ase_dFecha End                                                  
                                          
--Agrego los estados 9                                            
                                          
  SELECT Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then                                              
  SUBSTRING(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,3,2)  Else dbo.CNC_ASIENTO_VOUCHER.Per_cPeriodo End As Per_cPeriodo,                                              
              
  /*Case when @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' OR dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') and CNC_ASIENTO_VOUCHER.Asd_cEstadoD = '8' Then                                            
 LEFT(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + 'R00004'                                          
  when @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' OR dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') and CNC_ASIENTO_VOUCHER.Asd_cEstadoD = '9' Then                                            
 LEFT(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + 'R00005'                                          
  Else dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher End As Ase_nVoucher,  */            
  cnc_asiento_voucher.Ase_nVoucher as CUO,            
 case when left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) = '0100' then            
  LEFT(LTRIM(RTRIM('A' + left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + right(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,5) )),100)            
  when left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) = '0813' or left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) = '0814' then            
  LEFT(LTRIM(RTRIM('C' + left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + right(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,5) )),100)            
  else            
  LEFT(LTRIM(RTRIM('M' + left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + right(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,5) )),100)            
  end As Correlativo,              
  '01' as 'PCGE', dbo.CND_ASIENTO_VOUCHER.Pla_cCuentaContable,              
  Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then                                                                    
  DateAdd(day,-1, DateAdd(Month,1 , DateAdd(day,-(Day(dbo.CNC_ASIENTO_VOUCHER.Ase_dFecha)-1),dbo.CNC_ASIENTO_VOUCHER.Ase_dFecha))) Else                                                                    
  dbo.CNC_ASIENTO_VOUCHER.Ase_dFecha End As Ase_dFecha,                                 
  Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then                                                   
  'CENTRALIZACION REGISTRO DE ' + CASE WHEN dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Then 'VENTAS ' Else 'COMPRAS ' End + @Per_cPeriodo + ' - ' + @Pan_cAnio Else                                                                     
  dbo.CNC_ASIENTO_VOUCHER.Ase_cGlosa End As Ase_cGlosa,                                            
   Case When @TipoLibDi = '1' And dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Then                                                                   
    /*Case When (Sum(dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles - dbo.CND_ASIENTO_VOUCHER.Asd_nHaberSoles)) > 0 Then                          
     (Sum(dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles - dbo.CND_ASIENTO_VOUCHER.Asd_nHaberSoles))                                               
    Else  0 End*/ /*Sum(dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles)*/ dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles  Else Case When @TipoLibDi = '1' And dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06' Then                                                                   
     /*Case When (Sum(dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles - dbo.CND_ASIENTO_VOUCHER.Asd_nHaberSoles)) > 0 Then             
      (Sum(dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles - dbo.CND_ASIENTO_VOUCHER.Asd_nHaberSoles))                                                   
     ----Else 0 End*//*Sum(dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles)*/dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles Else dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles--Sum(dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles)                                                                  
   End End As Asd_nDebeSoles, Case When @TipoLibDi = '1' And dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Then                                                                   
  /*Case When (Sum(dbo.CND_ASIENTO_VOUCHER.Asd_nHaberSoles - dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles)) > 0 Then                                    
  Sum(dbo.CND_ASIENTO_VOUCHER.Asd_nHaberSoles - dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles) Else 0 End  */                                            
         --Sum(dbo.CND_ASIENTO_VOUCHER.Asd_nHaberSoles)    
         dbo.CND_ASIENTO_VOUCHER.Asd_nHaberSoles                                                     
  Else Case When @TipoLibDi = '1' And dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06' Then                                                                   
  /*CASE WHEN (Sum(dbo.CND_ASIENTO_VOUCHER.Asd_nHaberSoles - dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles)) > 0 Then                                                                  
  Sum(dbo.CND_ASIENTO_VOUCHER.Asd_nHaberSoles - dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles) Else 0 End */                                            
  --Sum(dbo.CND_ASIENTO_VOUCHER.Asd_nHaberSoles)
  dbo.CND_ASIENTO_VOUCHER.Asd_nHaberSoles                                                                 
  Else /*Sum(dbo.CND_ASIENTO_VOUCHER.Asd_nHaberSoles)*/ dbo.CND_ASIENTO_VOUCHER.Asd_nHaberSoles End                                                                  
  End As Asd_nHaberSoles,             
  '' as CUOVentas,            
  '' as CUOCompras,            
  '' as CUOConsignacion,               
  isnull(CNC_ASIENTO_VOUCHER.Asd_cEstadoO,'') as 'Asd_cEstadoO',                                              
  isnull(CNC_ASIENTO_VOUCHER.Asd_cEstadoD,'') as 'Asd_cEstadoD', dbo.CNT_TIPO_MONEDA.Mon_cCodSunat, dbo.CND_ASIENTO_VOUCHER.Asd_cSerieDoc, dbo.CND_ASIENTO_VOUCHER.Asd_cNumDoc,
  dbo.CND_ASIENTO_VOUCHER.Asd_dFecVen, dbo.CND_ASIENTO_VOUCHER.Asd_dFecDoc, dbo.CND_ASIENTO_VOUCHER.Asd_cTipoDoc, dbo.CND_ASIENTO_VOUCHER.Lib_cTipoLibro, dbo.CND_ASIENTO_VOUCHER.Ase_nVoucher, dbo.CND_ASIENTO_VOUCHER.Ten_cTipoEntidad, dbo.CND_ASIENTO_VOUCHER.Ent_cCodEntidad, dbo.CNM_ENTIDAD.Ent_nRuc, (SELECT t.Tab_cCodSunat FROM dbo.TABLA T
                                                                                                                WHERE T.Emp_cCodigo = @Emp_cCodigo AND T.Tab_cTabla = '003' AND t.Tab_cCodigo = dbo.CNM_ENTIDAD.Ent_cTipoDoc) as Ent_cTipoDoc           
  /*,                                     
  case when CND_ASIENTO_VOUCHER.Ent_cCodEntidad <> '' then CND_ASIENTO_VOUCHER.Asd_dFecDoc else '' end 'Asd_dFecDoc'*/                                            
 into #TMPDiarioPLE9                                             
 FROM dbo.EMPRESA RIGHT OUTER JOIN                                                                              
  dbo.CNC_ASIENTO_VOUCHER ON dbo.EMPRESA.Emp_cCodigo = dbo.CNC_ASIENTO_VOUCHER.Emp_cCodigo LEFT OUTER JOIN                                                                              
  dbo.CND_ASIENTO_VOUCHER LEFT OUTER JOIN                                                                              
  dbo.CNT_TIPODOC CNT_TIPODOC_2 ON dbo.CND_ASIENTO_VOUCHER.Asd_cTipoDocRef = CNT_TIPODOC_2.Tdo_cCodigo AND                                                                               
  dbo.CND_ASIENTO_VOUCHER.Emp_cCodigo = CNT_TIPODOC_2.Emp_cCodigo LEFT OUTER JOIN                                   
  dbo.CNT_TIPODOC CNT_TIPODOC_1 ON dbo.CND_ASIENTO_VOUCHER.Asd_cTipoDoc = CNT_TIPODOC_1.Tdo_cCodigo AND                                                                               
  dbo.CND_ASIENTO_VOUCHER.Emp_cCodigo = CNT_TIPODOC_1.Emp_cCodigo LEFT OUTER JOIN                                                                              
  dbo.CNM_ENTIDAD ON dbo.CND_ASIENTO_VOUCHER.Emp_cCodigo = dbo.CNM_ENTIDAD.Emp_cCodigo AND                                                                               
  dbo.CND_ASIENTO_VOUCHER.Ten_cTipoEntidad = dbo.CNM_ENTIDAD.Ten_cTipoEntidad AND                                              
  dbo.CND_ASIENTO_VOUCHER.Ent_cCodEntidad = dbo.CNM_ENTIDAD.Ent_cCodEntidad LEFT OUTER JOIN                                              
  dbo.CNM_PLAN_CTA ON dbo.CND_ASIENTO_VOUCHER.Emp_cCodigo = dbo.CNM_PLAN_CTA.Emp_cCodigo AND                                                                               
  dbo.CND_ASIENTO_VOUCHER.Pan_cAnio = dbo.CNM_PLAN_CTA.Pan_cAnio AND                                                           
  dbo.CND_ASIENTO_VOUCHER.Pla_cCuentaContable = dbo.CNM_PLAN_CTA.Pla_cCuentaContable LEFT OUTER JOIN                                              
  dbo.CNT_CENTRO_COSTO ON dbo.CND_ASIENTO_VOUCHER.Emp_cCodigo = dbo.CNT_CENTRO_COSTO.Emp_cCodigo AND                                                                               
  dbo.CND_ASIENTO_VOUCHER.Pan_cAnio = dbo.CNT_CENTRO_COSTO.Pan_cAnio AND                                                                               
  dbo.CND_ASIENTO_VOUCHER.Cos_cCodigo = dbo.CNT_CENTRO_COSTO.Cos_cCodigo ON                                                                               
  dbo.CNC_ASIENTO_VOUCHER.Emp_cCodigo = dbo.CND_ASIENTO_VOUCHER.Emp_cCodigo AND                                                                         
  dbo.CNC_ASIENTO_VOUCHER.Pan_cAnio = dbo.CND_ASIENTO_VOUCHER.Pan_cAnio AND                                                                            
  dbo.CNC_ASIENTO_VOUCHER.Per_cPeriodo = dbo.CND_ASIENTO_VOUCHER.Per_cPeriodo AND                                                                               
  dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = dbo.CND_ASIENTO_VOUCHER.Lib_cTipoLibro AND                                                                             
  dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher = dbo.CND_ASIENTO_VOUCHER.Ase_nVoucher                                                                               
  LEFT OUTER JOIN dbo.CNT_LIBRO_OPERA ON                                                                               
  dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = dbo.CNT_LIBRO_OPERA.Lib_cTipoLibro AND                                                                               
  dbo.CNC_ASIENTO_VOUCHER.PAN_CANIO =dbo. CNT_LIBRO_OPERA.PAN_CANIO AND                                                                               
  dbo.CNC_ASIENTO_VOUCHER.Emp_cCodigo = dbo.CNT_LIBRO_OPERA.Emp_cCodigo                                                                               
  LEFT OUTER JOIN  dbo.CNT_TIPO_MONEDA ON                                          
  dbo.CNC_ASIENTO_VOUCHER.Ase_cTipoMoneda = dbo.CNT_TIPO_MONEDA.Mon_cCodigo AND                                   
  dbo.CNC_ASIENTO_VOUCHER.Emp_cCodigo = dbo.CNT_TIPO_MONEDA.Emp_cCodigo                                                                               
  LEFT OUTER JOIN  dbo.CNT_ENTIDAD ON                                          
  dbo.CNM_ENTIDAD.Ten_cTipoEntidad = dbo.CNT_ENTIDAD.Ten_cTipoEntidad AND                                                                               
  dbo.CNM_ENTIDAD.Emp_cCodigo = dbo.CNT_ENTIDAD.Emp_cCodigo              
 WHERE  CNC_ASIENTO_VOUCHER.Emp_cCodigo = @Emp_cCodigo  AND               
 year(CNC_ASIENTO_VOUCHER.Ase_dfechaModifica) = @Pan_cAnio              
 and month(CNC_ASIENTO_VOUCHER.Ase_dfechaModifica) = @Per_cPeriodo                                  
 and (CNC_ASIENTO_VOUCHER.Asd_cEstadoD = '9')                                             
 AND (dbo.CNC_ASIENTO_VOUCHER.Ase_cDeleted <> '*') AND (dbo.CND_ASIENTO_VOUCHER.Asd_cDeleted <> '*')                                                                               
 AND  dbo.CND_ASIENTO_VOUCHER.Pla_cCuentaContable >= @ctaini              
 and   dbo.CND_ASIENTO_VOUCHER.Pla_cCuentaContable <= @ctafin                                                    
 GROUP BY                                          
   CNC_ASIENTO_VOUCHER.Asd_cEstadoO, dbo.CNM_ENTIDAD.Ent_cTipoDoc, dbo.CNM_ENTIDAD.Ent_nRuc, dbo.CND_ASIENTO_VOUCHER.Asd_nHaberSoles , dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles, CNC_ASIENTO_VOUCHER.Asd_cEstadoD,                                                                        
   dbo.CNC_ASIENTO_VOUCHER.Emp_cCodigo,              
 /* Case when @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' OR dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') and CNC_ASIENTO_VOUCHER.Asd_cEstadoD = '8' Then                                            
 LEFT(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + 'R00004'                      
  when @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' OR dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') and CNC_ASIENTO_VOUCHER.Asd_cEstadoD = '9' Then                                            
 LEFT(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + 'R00005'                                             
  Else                  
  dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher              
  End*/            
  dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,          
case when left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) = '0100' then            
  LEFT(LTRIM(RTRIM('A' + left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + right(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,5) )),100)            
  when left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) = '0813' or left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) = '0814' then            
  LEFT(LTRIM(RTRIM('C' + left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + right(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,5) )),100)            
  else            
  LEFT(LTRIM(RTRIM('M' + left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + right(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,5) )),100)            
  end             
  ,               
  --cnc_asiento_voucher.Ase_cNummov,                                         
   Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                     
   CNC_ASIENTO_VOUCHER.Ase_cNummov End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then                                                   
   SUBSTRING(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher, 3 ,2) Else dbo.CNC_ASIENTO_VOUCHER.Per_cPeriodo End,                                              
   Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then                             
   DateAdd(day,-1, DateAdd(Month,1 , DateAdd(day,-(Day(dbo.CNC_ASIENTO_VOUCHER.Ase_dFecha)-1),dbo.CNC_ASIENTO_VOUCHER.Ase_dFecha))) Else                                                                          
   dbo.CNC_ASIENTO_VOUCHER.Ase_dFecha End, dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro, dbo.CNT_LIBRO_OPERA.Lib_cDescripcion,                                              
    Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                     
    dbo.CNC_ASIENTO_VOUCHER.Ase_cTipoMoneda End,                                              
    Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                    
    dbo.CNT_TIPO_MONEDA.Mon_cNombreLargo End, dbo.CNC_ASIENTO_VOUCHER.Ase_nTipoCambio,                      
  Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then                                                         
  'CENTRALIZACION REGISTRO DE ' + CASE WHEN dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Then 'VENTAS ' Else 'COMPRAS ' End + @Per_cPeriodo + ' - ' + @Pan_cAnio ELse                                                                          
    dbo.CNC_ASIENTO_VOUCHER.Ase_cGlosa End,                                              
    Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then                                                     
    0 Else dbo.CND_ASIENTO_VOUCHER.Asd_nItem End, dbo.CND_ASIENTO_VOUCHER.Pla_cCuentaContable, dbo.CNM_PLAN_CTA.Pla_cNombreCuenta,                                                                               
    dbo.CNM_PLAN_CTA.Pla_cProvision, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then 'CENTRALIZACION REGISTRO DE ' +                                                     
  
    
      
        
          
            
              
    CASE WHEN dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Then 'VENTAS ' Else 'COMPRAS ' End + @Per_cPeriodo + ' - ' + @Pan_cAnio  Else                                                                          
    dbo.CND_ASIENTO_VOUCHER.Asd_cGlosa End,             
    dbo.CND_ASIENTO_VOUCHER.Cos_cCodigo, dbo.CNT_CENTRO_COSTO.Cos_cDescripcion, dbo.CND_ASIENTO_VOUCHER.Ten_cTipoEntidad,                                                                                                 
  Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                              
    dbo.CNM_ENTIDAD.Ent_nRuc End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                  
    dbo.CNM_ENTIDAD.Ent_cPersona End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                          
    dbo.CNT_ENTIDAD.Ten_cNombreEntidad End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                       
  
   
    dbo.CND_ASIENTO_VOUCHER.Asd_cTipoDoc End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                
    dbo.CND_ASIENTO_VOUCHER.Asd_cSerieDoc End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                    
  
   
       
    dbo.CND_ASIENTO_VOUCHER.Asd_cNumDoc End, /*Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                    
 
    dbo.CND_ASIENTO_VOUCHER.Asd_dFecDoc End,*/ Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                    
  
    
      
    dbo.CND_ASIENTO_VOUCHER.Asd_cTipoDocRef End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                  
  
    
      
        
    dbo.CND_ASIENTO_VOUCHER.Asd_cSerieDocRef End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                 
  
    
      
        
         
    dbo.CND_ASIENTO_VOUCHER.Asd_cNumDocRef End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                   
    dbo.CND_ASIENTO_VOUCHER.Asd_dFecDocRef End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then 0 Else                                                                    
  
    
      
    dbo.CND_ASIENTO_VOUCHER.Asd_nMontoInafecto End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                               
  
    dbo.CND_ASIENTO_VOUCHER.Asd_cRetencion End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                   
  
    
      
       
    dbo.CND_ASIENTO_VOUCHER.Asd_dFechaSpot End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                   
  
    
      
       
    dbo.CND_ASIENTO_VOUCHER.Asd_cNumSpot End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                        
    dbo.CND_ASIENTO_VOUCHER.Asd_cDestino End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                     
  
    
     
    dbo.CND_ASIENTO_VOUCHER.Asd_nCorre End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                       
  
   
    CNT_TIPODOC_1.Tdo_cNombreLargo End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                          
 CNT_TIPODOC_1.Tdo_cNombreCorto End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                       
    CNT_TIPODOC_2.Tdo_cNombreLargo End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                          
    CNT_TIPODOC_2.Tdo_cNombreCorto End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                     
    Ase_cUserModifica End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                               
    dbo.EMPRESA.Emp_cNombreLargo End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                              
    dbo.EMPRESA.Emp_cNombreCorto  End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                 
    Ase_cUserModifica END , dbo.CNT_TIPO_MONEDA.Mon_cCodSunat, dbo.CND_ASIENTO_VOUCHER.Asd_cSerieDoc, dbo.CND_ASIENTO_VOUCHER.Asd_cNumDoc,
  dbo.CND_ASIENTO_VOUCHER.Asd_dFecVen, dbo.CND_ASIENTO_VOUCHER.Asd_dFecDoc, dbo.CND_ASIENTO_VOUCHER.Asd_cTipoDoc, dbo.CND_ASIENTO_VOUCHER.Lib_cTipoLibro, dbo.CND_ASIENTO_VOUCHER.Ase_nVoucher, dbo.CND_ASIENTO_VOUCHER.Ten_cTipoEntidad, dbo.CND_ASIENTO_VOUCHER.Ent_cCodEntidad/*,                                    
    CND_ASIENTO_VOUCHER.Ent_cCodEntidad,CND_ASIENTO_VOUCHER.Asd_dFecDoc*/                                                 
  ORDER BY                                            
  Case When @TipoLibDi = '1' And                                            
  (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then                                                   
  SUBSTRING(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher, 3 ,2)                                            
  Else  dbo.CNC_ASIENTO_VOUCHER.Per_cPeriodo End,                                             
 /* Case when @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' OR dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') and CNC_ASIENTO_VOUCHER.Asd_cEstadoD = '8' Then                                            
 LEFT(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + 'R00004'                                             
  when @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' OR dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') and CNC_ASIENTO_VOUCHER.Asd_cEstadoD = '9' Then                                            
 LEFT(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + 'R00005'                                             
  Else                                          
  dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher          
  End*/            
  dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher, dbo.CNM_ENTIDAD.Ent_nRuc, dbo.CNM_ENTIDAD.Ent_cTipoDoc,          
case when left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) = '0100' then            
  LEFT(LTRIM(RTRIM('A' + left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + right(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,5) )),100)            
  when left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) = '0813' or left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) = '0814' then            
  LEFT(LTRIM(RTRIM('C' + left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + right(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,5) )),100)            
  else            
  LEFT(LTRIM(RTRIM('M' + left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + right(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,5) )),100)            
  end ,              
            
  --cnc_asiento_voucher.Ase_cNummov,            
  Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then                                                                           
   DateAdd(day,-1, DateAdd(Month,1 , DateAdd(day,-(Day(dbo.CNC_ASIENTO_VOUCHER.Ase_dFecha)-1),dbo.CNC_ASIENTO_VOUCHER.Ase_dFecha))) Else                                                                          
   dbo.CNC_ASIENTO_VOUCHER.Ase_dFecha End                 
                 
--Agrego los estados 8              
              
  SELECT Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then                                              
  SUBSTRING(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,3,2)  Else dbo.CNC_ASIENTO_VOUCHER.Per_cPeriodo End As Per_cPeriodo,                                              
  /*Case when @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' OR dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') and CNC_ASIENTO_VOUCHER.Asd_cEstadoD = '8' Then                                            
 LEFT(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + 'R00004'                                          
  when @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' OR dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') and CNC_ASIENTO_VOUCHER.Asd_cEstadoD = '9' Then                                            
 LEFT(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + 'R00005'                                          
  Else dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher End As Ase_nVoucher*/            
  cnc_asiento_voucher.Ase_nVoucher as CUO,          
case when left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) = '0100' then            
 LEFT(LTRIM(RTRIM('A' + left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + right(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,5) )),100)            
  when left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) = '0813' or left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) = '0814' then            
  LEFT(LTRIM(RTRIM('C' + left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + right(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,5) )),100)            
  else            
  LEFT(LTRIM(RTRIM('M' + left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + right(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,5) )),100)            
  end As Correlativo,            
  --cnc_asiento_voucher.Ase_cNummov,                                                                   
  '01' as 'PCGE', dbo.CND_ASIENTO_VOUCHER.Pla_cCuentaContable,              
  Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then                                                                    
  DateAdd(day,-1, DateAdd(Month,1 , DateAdd(day,-(Day(dbo.CNC_ASIENTO_VOUCHER.Ase_dFecha)-1),dbo.CNC_ASIENTO_VOUCHER.Ase_dFecha))) Else                                                                    
  dbo.CNC_ASIENTO_VOUCHER.Ase_dFecha End As Ase_dFecha,                                 
  Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then                                                   
  'CENTRALIZACION REGISTRO DE ' + CASE WHEN dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Then 'VENTAS ' Else 'COMPRAS ' End + @Per_cPeriodo + ' - ' + @Pan_cAnio Else                                                                     
  dbo.CNC_ASIENTO_VOUCHER.Ase_cGlosa End As Ase_cGlosa,                                            
   Case When @TipoLibDi = '1' And dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Then                                                                   
    /*Case When (Sum(dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles - dbo.CND_ASIENTO_VOUCHER.Asd_nHaberSoles)) > 0 Then                                                           
     (Sum(dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles - dbo.CND_ASIENTO_VOUCHER.Asd_nHaberSoles))                                               
    Else  0 End*/ /*Sum(dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles)*/ dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles  Else Case When @TipoLibDi = '1' And dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06' Then                                                                   
     /*Case When (Sum(dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles - dbo.CND_ASIENTO_VOUCHER.Asd_nHaberSoles)) > 0 Then                                                                   
      (Sum(dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles - dbo.CND_ASIENTO_VOUCHER.Asd_nHaberSoles))                                                   
     Else 0 End*//*Sum(dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles)*/ dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles Else dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles--Sum(dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles)                                                                  
   End End As Asd_nDebeSoles, Case When @TipoLibDi = '1' And dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Then                                                                   
  /*Case When (Sum(dbo.CND_ASIENTO_VOUCHER.Asd_nHaberSoles - dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles)) > 0 Then                                                                   
  Sum(dbo.CND_ASIENTO_VOUCHER.Asd_nHaberSoles - dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles) Else 0 End  */                                            
         --Sum(dbo.CND_ASIENTO_VOUCHER.Asd_nHaberSoles)                                                         
         dbo.CND_ASIENTO_VOUCHER.Asd_nHaberSoles
  Else Case When @TipoLibDi = '1' And dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06' Then                                                                   
  /*CASE WHEN (Sum(dbo.CND_ASIENTO_VOUCHER.Asd_nHaberSoles - dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles)) > 0 Then                                                                  
  Sum(dbo.CND_ASIENTO_VOUCHER.Asd_nHaberSoles - dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles) Else 0 End */                                            
  --Sum(dbo.CND_ASIENTO_VOUCHER.Asd_nHaberSoles) 
  dbo.CND_ASIENTO_VOUCHER.Asd_nHaberSoles                                                                
  Else /*Sum(dbo.CND_ASIENTO_VOUCHER.Asd_nHaberSoles)*/ dbo.CND_ASIENTO_VOUCHER.Asd_nHaberSoles End                                                                  
  End As Asd_nHaberSoles,             
  '' as CUOVentas,            
  '' as CUOCompras,            
  '' as CUOConsignacion,              
  isnull(CNC_ASIENTO_VOUCHER.Asd_cEstadoO,'') as 'Asd_cEstadoO',                                              
  isnull(CNC_ASIENTO_VOUCHER.Asd_cEstadoD,'') as 'Asd_cEstadoD' , dbo.CNT_TIPO_MONEDA.Mon_cCodSunat, dbo.CND_ASIENTO_VOUCHER.Asd_cSerieDoc, dbo.CND_ASIENTO_VOUCHER.Asd_cNumDoc,
  dbo.CND_ASIENTO_VOUCHER.Asd_dFecVen, dbo.CND_ASIENTO_VOUCHER.Asd_dFecDoc, dbo.CND_ASIENTO_VOUCHER.Asd_cTipoDoc, dbo.CND_ASIENTO_VOUCHER.Lib_cTipoLibro, dbo.CND_ASIENTO_VOUCHER.Ase_nVoucher, dbo.CND_ASIENTO_VOUCHER.Ten_cTipoEntidad, dbo.CND_ASIENTO_VOUCHER.Ent_cCodEntidad, dbo.CNM_ENTIDAD.Ent_nRuc, (SELECT t.Tab_cCodSunat FROM dbo.TABLA T
                                                                                                                WHERE T.Emp_cCodigo = @Emp_cCodigo AND T.Tab_cTabla = '003' AND t.Tab_cCodigo = dbo.CNM_ENTIDAD.Ent_cTipoDoc) as Ent_cTipoDoc/*,                                     
  case when CND_ASIENTO_VOUCHER.Ent_cCodEntidad <> '' then CND_ASIENTO_VOUCHER.Asd_dFecDoc else '' end 'Asd_dFecDoc'*/                                            
 into #TMPDiarioPLE8                                             
 FROM dbo.EMPRESA RIGHT OUTER JOIN                                                                              
  dbo.CNC_ASIENTO_VOUCHER ON dbo.EMPRESA.Emp_cCodigo = dbo.CNC_ASIENTO_VOUCHER.Emp_cCodigo LEFT OUTER JOIN                                                     
  dbo.CND_ASIENTO_VOUCHER LEFT OUTER JOIN                                                                              
  dbo.CNT_TIPODOC CNT_TIPODOC_2 ON dbo.CND_ASIENTO_VOUCHER.Asd_cTipoDocRef = CNT_TIPODOC_2.Tdo_cCodigo AND                                                                               
  dbo.CND_ASIENTO_VOUCHER.Emp_cCodigo = CNT_TIPODOC_2.Emp_cCodigo LEFT OUTER JOIN                                   
  dbo.CNT_TIPODOC CNT_TIPODOC_1 ON dbo.CND_ASIENTO_VOUCHER.Asd_cTipoDoc = CNT_TIPODOC_1.Tdo_cCodigo AND                                                                               
  dbo.CND_ASIENTO_VOUCHER.Emp_cCodigo = CNT_TIPODOC_1.Emp_cCodigo LEFT OUTER JOIN                                                                              
  dbo.CNM_ENTIDAD ON dbo.CND_ASIENTO_VOUCHER.Emp_cCodigo = dbo.CNM_ENTIDAD.Emp_cCodigo AND           
  dbo.CND_ASIENTO_VOUCHER.Ten_cTipoEntidad = dbo.CNM_ENTIDAD.Ten_cTipoEntidad AND                                              
  dbo.CND_ASIENTO_VOUCHER.Ent_cCodEntidad = dbo.CNM_ENTIDAD.Ent_cCodEntidad LEFT OUTER JOIN                                              
  dbo.CNM_PLAN_CTA ON dbo.CND_ASIENTO_VOUCHER.Emp_cCodigo = dbo.CNM_PLAN_CTA.Emp_cCodigo AND                                                                               
  dbo.CND_ASIENTO_VOUCHER.Pan_cAnio = dbo.CNM_PLAN_CTA.Pan_cAnio AND                                                           
  dbo.CND_ASIENTO_VOUCHER.Pla_cCuentaContable = dbo.CNM_PLAN_CTA.Pla_cCuentaContable LEFT OUTER JOIN                                                                              
  dbo.CNT_CENTRO_COSTO ON dbo.CND_ASIENTO_VOUCHER.Emp_cCodigo = dbo.CNT_CENTRO_COSTO.Emp_cCodigo AND                                                                               
  dbo.CND_ASIENTO_VOUCHER.Pan_cAnio = dbo.CNT_CENTRO_COSTO.Pan_cAnio AND                                                                               
  dbo.CND_ASIENTO_VOUCHER.Cos_cCodigo = dbo.CNT_CENTRO_COSTO.Cos_cCodigo ON                                                                               
  dbo.CNC_ASIENTO_VOUCHER.Emp_cCodigo = dbo.CND_ASIENTO_VOUCHER.Emp_cCodigo AND                                                                         
  dbo.CNC_ASIENTO_VOUCHER.Pan_cAnio = dbo.CND_ASIENTO_VOUCHER.Pan_cAnio AND                         
  dbo.CNC_ASIENTO_VOUCHER.Per_cPeriodo = dbo.CND_ASIENTO_VOUCHER.Per_cPeriodo AND                                                                               
  dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = dbo.CND_ASIENTO_VOUCHER.Lib_cTipoLibro AND                                                                             
  dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher = dbo.CND_ASIENTO_VOUCHER.Ase_nVoucher                                                                               
  LEFT OUTER JOIN dbo.CNT_LIBRO_OPERA ON           
  dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = dbo.CNT_LIBRO_OPERA.Lib_cTipoLibro AND                                                                               
  dbo.CNC_ASIENTO_VOUCHER.PAN_CANIO =dbo. CNT_LIBRO_OPERA.PAN_CANIO AND                                                                               
  dbo.CNC_ASIENTO_VOUCHER.Emp_cCodigo = dbo.CNT_LIBRO_OPERA.Emp_cCodigo                                                   
  LEFT OUTER JOIN  dbo.CNT_TIPO_MONEDA ON                                          
  dbo.CNC_ASIENTO_VOUCHER.Ase_cTipoMoneda = dbo.CNT_TIPO_MONEDA.Mon_cCodigo AND                                   
  dbo.CNC_ASIENTO_VOUCHER.Emp_cCodigo = dbo.CNT_TIPO_MONEDA.Emp_cCodigo                                                                               
  LEFT OUTER JOIN  dbo.CNT_ENTIDAD ON                                          
  dbo.CNM_ENTIDAD.Ten_cTipoEntidad = dbo.CNT_ENTIDAD.Ten_cTipoEntidad AND                                                                               
  dbo.CNM_ENTIDAD.Emp_cCodigo = dbo.CNT_ENTIDAD.Emp_cCodigo                                                                               
 WHERE  CNC_ASIENTO_VOUCHER.Emp_cCodigo = @Emp_cCodigo  AND               
 CNC_ASIENTO_VOUCHER.Pan_cAnio  = @Pan_cAnio              
 and CNC_ASIENTO_VOUCHER.Per_cPeriodo = @Per_cPeriodo                                  
 and (CNC_ASIENTO_VOUCHER.Asd_cEstadoD = '8')                                             
 AND (dbo.CNC_ASIENTO_VOUCHER.Ase_cDeleted <> '*') AND (dbo.CND_ASIENTO_VOUCHER.Asd_cDeleted <> '*')                                                                               
 AND  dbo.CND_ASIENTO_VOUCHER.Pla_cCuentaContable >= @ctaini              
 and   dbo.CND_ASIENTO_VOUCHER.Pla_cCuentaContable <= @ctafin                                                    
 GROUP BY                                          
   CNC_ASIENTO_VOUCHER.Asd_cEstadoO,dbo.CNM_ENTIDAD.Ent_cTipoDoc, dbo.CNM_ENTIDAD.Ent_nRuc, dbo.CND_ASIENTO_VOUCHER.Ten_cTipoEntidad, dbo.CND_ASIENTO_VOUCHER.Ent_cCodEntidad, dbo.CND_ASIENTO_VOUCHER.Asd_nHaberSoles, dbo.CND_ASIENTO_VOUCHER.Asd_nDebeSoles, CNC_ASIENTO_VOUCHER.Asd_cEstadoD,                                                                        
   dbo.CNC_ASIENTO_VOUCHER.Emp_cCodigo,              
  /*Case when @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' OR dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') and CNC_ASIENTO_VOUCHER.Asd_cEstadoD = '8' Then                                            
 LEFT(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + 'R00004'                      
  when @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' OR dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') and CNC_ASIENTO_VOUCHER.Asd_cEstadoD = '9' Then                                            
 LEFT(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + 'R00005'                                             
  Else                                     
  dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher              
  End*/            
  dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher  ,          
case when left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) = '0100' then            
  LEFT(LTRIM(RTRIM('A' + left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + right(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,5) )),100)            
  when left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) = '0813' or left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) = '0814' then            
  LEFT(LTRIM(RTRIM('C' + left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + right(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,5) )),100)            
  else            
  LEFT(LTRIM(RTRIM('M' + left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + right(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,5) )),100)            
  end,                                            
  --cnc_asiento_voucher.Ase_cNummov,            
   Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                     
   CNC_ASIENTO_VOUCHER.Ase_cNummov End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then                       
   SUBSTRING(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher, 3 ,2) Else dbo.CNC_ASIENTO_VOUCHER.Per_cPeriodo End,                                              
   Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then                             
   DateAdd(day,-1, DateAdd(Month,1 , DateAdd(day,-(Day(dbo.CNC_ASIENTO_VOUCHER.Ase_dFecha)-1),dbo.CNC_ASIENTO_VOUCHER.Ase_dFecha))) Else                                                                          
   dbo.CNC_ASIENTO_VOUCHER.Ase_dFecha End, dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro, dbo.CNT_LIBRO_OPERA.Lib_cDescripcion,                                              
    Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                     
    dbo.CNC_ASIENTO_VOUCHER.Ase_cTipoMoneda End,                                              
    Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                    
    dbo.CNT_TIPO_MONEDA.Mon_cNombreLargo End, dbo.CNC_ASIENTO_VOUCHER.Ase_nTipoCambio,                                              
  Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then                                                         
  'CENTRALIZACION REGISTRO DE ' + CASE WHEN dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Then 'VENTAS ' Else 'COMPRAS ' End + @Per_cPeriodo + ' - ' + @Pan_cAnio ELse                                                                          
    dbo.CNC_ASIENTO_VOUCHER.Ase_cGlosa End,                                              
    Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then                                                     
    0 Else dbo.CND_ASIENTO_VOUCHER.Asd_nItem End, dbo.CND_ASIENTO_VOUCHER.Pla_cCuentaContable, dbo.CNM_PLAN_CTA.Pla_cNombreCuenta,                                                                               
    dbo.CNM_PLAN_CTA.Pla_cProvision, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then 'CENTRALIZACION REGISTRO DE ' +                                                     
  
    
       
       
          
             
             
    CASE WHEN dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Then 'VENTAS ' Else 'COMPRAS ' End + @Per_cPeriodo + ' - ' + @Pan_cAnio  Else                                                                          
    dbo.CND_ASIENTO_VOUCHER.Asd_cGlosa End,                    
    dbo.CND_ASIENTO_VOUCHER.Cos_cCodigo, dbo.CNT_CENTRO_COSTO.Cos_cDescripcion, dbo.CND_ASIENTO_VOUCHER.Ten_cTipoEntidad,                                                                                                 
  Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                              
    dbo.CNM_ENTIDAD.Ent_nRuc End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                  
    dbo.CNM_ENTIDAD.Ent_cPersona End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                          
    dbo.CNT_ENTIDAD.Ten_cNombreEntidad End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                    
   
    dbo.CND_ASIENTO_VOUCHER.Asd_cTipoDoc End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                
    dbo.CND_ASIENTO_VOUCHER.Asd_cSerieDoc End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                   
   
    
      
    dbo.CND_ASIENTO_VOUCHER.Asd_cNumDoc End, /*Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                    
 
    dbo.CND_ASIENTO_VOUCHER.Asd_dFecDoc End,*/ Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                    
  
    
      
    dbo.CND_ASIENTO_VOUCHER.Asd_cTipoDocRef End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                  
  
    
      
        
    dbo.CND_ASIENTO_VOUCHER.Asd_cSerieDocRef End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                 
  
    
      
        
         
    dbo.CND_ASIENTO_VOUCHER.Asd_cNumDocRef End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                   
    dbo.CND_ASIENTO_VOUCHER.Asd_dFecDocRef End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then 0 Else                                                                    
  
    
      
    dbo.CND_ASIENTO_VOUCHER.Asd_nMontoInafecto End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                     
    dbo.CND_ASIENTO_VOUCHER.Asd_cRetencion End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                   
  
   
       
       
    dbo.CND_ASIENTO_VOUCHER.Asd_dFechaSpot End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                   
  
    
      
       
    dbo.CND_ASIENTO_VOUCHER.Asd_cNumSpot End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                        
    dbo.CND_ASIENTO_VOUCHER.Asd_cDestino End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                     
  
    
     
    dbo.CND_ASIENTO_VOUCHER.Asd_nCorre End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                       
  
   
    CNT_TIPODOC_1.Tdo_cNombreLargo End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                          
 CNT_TIPODOC_1.Tdo_cNombreCorto End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                          
    CNT_TIPODOC_2.Tdo_cNombreLargo End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                          
    CNT_TIPODOC_2.Tdo_cNombreCorto End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                     
    Ase_cUserModifica End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                                        
    dbo.EMPRESA.Emp_cNombreLargo End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                                              
    dbo.EMPRESA.Emp_cNombreCorto  End, Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then '' Else                                       
    Ase_cUserModifica END , dbo.CNT_TIPO_MONEDA.Mon_cCodSunat, dbo.CND_ASIENTO_VOUCHER.Asd_cSerieDoc, dbo.CND_ASIENTO_VOUCHER.Asd_cNumDoc,
  dbo.CND_ASIENTO_VOUCHER.Asd_dFecVen, dbo.CND_ASIENTO_VOUCHER.Asd_dFecDoc, dbo.CND_ASIENTO_VOUCHER.Asd_cTipoDoc, dbo.CND_ASIENTO_VOUCHER.Lib_cTipoLibro, dbo.CND_ASIENTO_VOUCHER.Ase_nVoucher/*,                                    
    CND_ASIENTO_VOUCHER.Ent_cCodEntidad,CND_ASIENTO_VOUCHER.Asd_dFecDoc*/                                                 
  ORDER BY                                            
  Case When @TipoLibDi = '1' And                                            
  (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then                                                   
  SUBSTRING(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher, 3 ,2)                                            
  Else  dbo.CNC_ASIENTO_VOUCHER.Per_cPeriodo End,                                             
  /*Case when @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' OR dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') and CNC_ASIENTO_VOUCHER.Asd_cEstadoD = '8' Then                                            
 LEFT(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + 'R00004'                                             
  when @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' OR dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') and CNC_ASIENTO_VOUCHER.Asd_cEstadoD = '9' Then                                            
 LEFT(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + 'R00005'                                             
  Else                                          
  dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher                                             
  End*/            
  dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher, dbo.CNM_ENTIDAD.Ent_nRuc, dbo.CND_ASIENTO_VOUCHER.Ten_cTipoEntidad, dbo.CND_ASIENTO_VOUCHER.Ent_cCodEntidad,          
  case when left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) = '0100' then            
  LEFT(LTRIM(RTRIM('A' + left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + right(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,5) )),100)            
  when left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) = '0813' or left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) = '0814' then            
  LEFT(LTRIM(RTRIM('C' + left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + right(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,5) )),100)            
  else            
  LEFT(LTRIM(RTRIM('M' + left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + right(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,5) )),100)            
  end,             
  --cnc_asiento_voucher.Ase_cNummov,            
  Case When @TipoLibDi = '1' And (dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '05' Or dbo.CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = '06') Then                                                                           
   DateAdd(day,-1, DateAdd(Month,1 , DateAdd(day,-(Day(dbo.CNC_ASIENTO_VOUCHER.Ase_dFecha)-1),dbo.CNC_ASIENTO_VOUCHER.Ase_dFecha))) Else                                                                          
   dbo.CNC_ASIENTO_VOUCHER.Ase_dFecha End                 
          
                             
select *                                           
into #TMPDiarioPLE                                          
from #TMPDiarioPLE1                                          
union all                                          
select * from #TMPDiarioPLE9               
union all               
select * from #TMPDiarioPLE8          
          
                   
delete from #TMPDiarioPLE          
where Asd_nDebeSoles = 0 and Asd_nHaberSoles = 0          
                                   
                                   
                                   
select /*case when Asd_cEstadoO in ('6','7') then LTRIM(RTRIM(LEFT(convert(varchar(4),YEAR(Asd_dFecDoc)) +  convert(varchar(2),replicate ('0',(2 - len(MONTH(Asd_dFecDoc))))) + convert(varchar(2),MONTH(Asd_dFecDoc)) + '00',8)))                             
         
  else*/ LTRIM(RTRIM(LEFT(convert(varchar(4),YEAR(Ase_dFecha)) +  case Per_cPeriodo when '00' then '01' when '13' then '12' when '14' then '12' else Per_cPeriodo end  + '00',8)))                                     
  /*end*/ as 'Periodo',                
      LTRIM(RTRIM(LEFT(CUO,40))) as 'CUO',          
      LTRIM(RTRIM(LEFT(Correlativo,40))) as 'Correlativo',           
      LTRIM(RTRIM(LEFT(PCGE,2))) as 'PCGE',                
      LTRIM(RTRIM(LEFT(Pla_cCuentaContable,24))) as 'Pla_cCuentaContable',                
      LTRIM(RTRIM(LEFT(convert(varchar(10),Ase_dFecha,103),10))) as 'Ase_dFecha',                
      LTRIM(RTRIM(LEFT(replace(replace(replace(Ase_cGlosa, char(13), ' '), char(10), ' '), char(9), ' '),100)))  as 'Ase_cGlosa',                                                      
      CONVERT(money,Asd_nDebeSoles) as 'Asd_nDebeSoles',                
      CONVERT(money,Asd_nHaberSoles) as 'Asd_nHaberSoles',                
   CUOVentas,            
   CUOCompras,            
   CUOConsignacion,             
      case when Asd_cEstadoD = '' then                                                    
      case when Asd_cEstadoO in ('0','1','6','7') then '1' end
      else Asd_cEstadoD end as 'Estado', Mon_cCodSunat, Asd_cSerieDoc, Asd_cNumDoc,
  Asd_dFecVen, Asd_dFecDoc, Asd_cTipoDoc, Lib_cTipoLibro, Ase_nVoucher, Ten_cTipoEntidad, Ent_cCodEntidad, Ent_nRuc, Ent_cTipoDoc
      --when Asd_cEstadoD = '9' and Per_cPeriodo = @Per_cPeriodo and YEAR(Ase_dFecha) = @Pan_cAnio then '1' else Asd_cEstadoD end as 'Estado'
into #TMPDiarioPLE6                                            
from #TMPDiarioPLE                                                 
order by convert(varchar(4),YEAR(Ase_dFecha)) +  Per_cPeriodo  + '00', CUO                            
                                      
delete from #TMPDiarioPLE6                                                  
where Estado = '2'   
              
                            
/*Update  A                       
set A.Ase_dFecha=B.Ase_dFecha                            
from #TMPDiarioPLE6 A inner join #TMPDiarioPLE6 B                                  
on A.Ase_nVoucher=B.Ase_nVoucher and B.Estado ='8' and year(B.Ase_dFecha) <> '1900'*/                             
                            
create table #TMPDiarioPLEDH                                            
( Per_cPeriodo varchar(8),            
 CUO varchar(12),            
 Correlativo varchar(10),            
 PCGE varchar(2),            
 Pla_cCuentaContable varchar(12),            
 Ase_dFecha datetime,            
 Ase_cGlosa varchar(100),                                      
 Asd_nDebeSoles decimal(14,2),                                            
 Asd_nHaberSoles decimal(14,2),            
 CUOVentas varchar(40),            
 CUOCompras varchar(40),            
 CUOConsignacion varchar(40),            
 Estado char(1)    , 
 Mon_cCodSunat CHAR(3), 
 Asd_cSerieDoc VARCHAR(20), 
 Asd_cNumDoc VARCHAR(25),
 Asd_dFecVen DATETIME, 
 Asd_dFecDoc DATETIME, 
 Asd_cTipoDoc CHAR(2), 
 Lib_cTipoLibro CHAR(2), 
 Ase_nVoucher VARCHAR(10) ,
 Corr INT IDENTITY(1,1),
 Ten_cTipoEntidad CHAR(1),
 Ent_cCodEntidad CHAR(5),
 Ent_nRuc VARCHAR(15), 
 Ent_cTipoDoc CHAR(2)                    
 )                            
          
insert into #TMPDiarioPLEDH                             
select Periodo, CUO,Correlativo,PCGE, Pla_cCuentaContable, Ase_dFecha, Ase_cGlosa,/*SUM(Asd_nDebeSoles)*/ Asd_nDebeSoles as 'Asd_nDebeSoles',             
/*SUM(Asd_nHaberSoles)*/ Asd_nHaberSoles as 'Asd_nHaberSoles',CUOVentas,CUOCompras,CUOConsignacion, Estado, Mon_cCodSunat, Asd_cSerieDoc, Asd_cNumDoc,
  Asd_dFecVen, Asd_dFecDoc, Asd_cTipoDoc, Lib_cTipoLibro, Ase_nVoucher, Ten_cTipoEntidad, Ent_cCodEntidad, ISNULL(Ent_nRuc, '') AS Ent_nRuc, ISNULL(Ent_cTipoDoc, '') AS Ent_cTipoDoc       
--case when month(ase_dfecha)= month(cast(@desde as datetime)) and year(ase_dfecha)= year(cast(@desde as datetime)) then '1' else Estado end   
--case when month(ase_dfecha)= month(@desde) then '1' else Estado end     
from #TMPDiarioPLE6                                             
--group by Periodo, CUO ,Correlativo,PCGE, Pla_cCuentaContable, Ase_dFecha, Ase_cGlosa,CUOVentas,CUOCompras,CUOConsignacion, Estado , Mon_cCodSunat, Asd_cSerieDoc, Asd_cNumDoc,
--  Asd_dFecVen, Asd_dFecDoc, Asd_cTipoDoc, Lib_cTipoLibro, Ase_nVoucher , Asd_nDebeSoles , Asd_nHaberSoles                                           
order by  Periodo, CUO                                            
               
               
                                           
Declare @Per_cPeriodoDH varchar(8)                                            
Declare @CUO varchar(12)            
Declare @Correlativo varchar(10)            
Declare @PCGE varchar(2)                                   
Declare @Pla_cCuentaContable varchar(15)                                            
Declare @Ase_dFecha datetime                                            
Declare @Ase_cGlosa varchar(50)                                            
Declare @Asd_nDebeSoles decimal(14,2)                                            
Declare @Asd_nHaberSoles decimal(14,2)               
Declare @CUOVentas varchar(40)                                            
Declare @CUOCompras varchar(40)            
Declare @CUOConsignacion varchar(40)            
Declare @Estado char(1) 
DECLARE @Mon_cCodSunat CHAR(3)
DECLARE @Asd_cSerieDoc VARCHAR(20) 
DECLARE @Asd_cNumDoc VARCHAR(25)
DECLARE @Asd_dFecVen DATETIME 
DECLARE @Asd_dFecDoc DATETIME 
DECLARE @Asd_cTipoDoc CHAR(2) 
DECLARE @Lib_cTipoLibro CHAR(2) 
DECLARE @Ase_nVoucher VARCHAR(10)
DECLARE @Ten_cTipoEntidad CHAR(1)
DECLARE @Ent_cCodEntidad CHAR(5)
DECLARE @Ent_nRuc VARCHAR(15) 
DECLARE @Ent_cTipoDoc CHAR(2)    
     
set @Nro = 1                                
          
DECLARE @TipoDocAuxiliar CHAR(2)
          
DECLARE D_H_Cursor_TC CURSOR FOR          
SELECT Per_cPeriodo, CUO,Correlativo, PCGE ,Pla_cCuentaContable ,Ase_dFecha,Ase_cGlosa,/*sum(Asd_nDebeSoles)*/ Asd_nDebeSoles, /*sum(Asd_nHaberSoles)*/ Asd_nHaberSoles,CUOVentas,CUOCompras,CUOConsignacion,Estado
, Mon_cCodSunat, Asd_cSerieDoc, Asd_cNumDoc,
  Asd_dFecVen, Asd_dFecDoc, Asd_cTipoDoc, Lib_cTipoLibro, Ase_nVoucher, Ten_cTipoEntidad, Ent_cCodEntidad, Ent_nRuc, Ent_cTipoDoc          
FROM #TMPDiarioPLEDH          
group by Per_cPeriodo, CUO,Correlativo, PCGE ,Pla_cCuentaContable ,Ase_dFecha,Ase_cGlosa,CUOVentas,CUOCompras,CUOConsignacion,Estado
, Mon_cCodSunat, Asd_cSerieDoc, Asd_cNumDoc,
  Asd_dFecVen, Asd_dFecDoc, Asd_cTipoDoc, Lib_cTipoLibro, Ase_nVoucher, Asd_nDebeSoles, Asd_nHaberSoles , Ten_cTipoEntidad, Ent_cCodEntidad, Ent_nRuc, Ent_cTipoDoc         
having sum(Asd_nDebeSoles) > 0 and sum(Asd_nHaberSoles) > 0          
                                            
OPEN D_H_Cursor_TC                                              
FETCH NEXT FROM D_H_Cursor_TC                                              
INTO @Per_cPeriodoDH, @CUO,@Correlativo, @PCGE ,@Pla_cCuentaContable ,@Ase_dFecha,@Ase_cGlosa,@Asd_nDebeSoles,@Asd_nHaberSoles,@CUOVentas,@CUOCompras,@CUOConsignacion,@Estado
, @Mon_cCodSunat, @Asd_cSerieDoc, @Asd_cNumDoc,
  @Asd_dFecVen, @Asd_dFecDoc, @Asd_cTipoDoc, @Lib_cTipoLibro, @Ase_nVoucher, @Ten_cTipoEntidad, @Ent_cCodEntidad, @Ent_nRuc, @Ent_cTipoDoc                                            
                              
 WHILE @@FETCH_STATUS = 0                                              
 BEGIN           
    
            
 insert into #TMPDiarioPLEDH          
 select @Per_cPeriodoDH as 'Per_cPeriodo',CUO, LEFT(@Correlativo,5) + 'R' + replicate('0',4 - len(@Nro)) + convert(varchar(10),@Nro)  as 'Correlativo',          
 @PCGE as 'PCGE' ,@Pla_cCuentaContable as 'Pla_cCuentaContable' ,@Ase_dFecha as 'Ase_dFecha',          
 @Ase_cGlosa as 'Ase_cGlosa',0 as 'Asd_nDebeSoles',@Asd_nHaberSoles,@CUOVentas,@CUOCompras,@CUOConsignacion,@Estado
 , @Mon_cCodSunat, @Asd_cSerieDoc, @Asd_cNumDoc,
  @Asd_dFecVen, @Asd_dFecDoc, @Asd_cTipoDoc, @Lib_cTipoLibro, @Ase_nVoucher, @Ten_cTipoEntidad, @Ent_cCodEntidad, @Ent_nRuc, @Ent_cTipoDoc from #TMPDiarioPLEDH          
 where Per_cPeriodo = @Per_cPeriodoDH and CUO  = @CUO and Pla_cCuentaContable = @Pla_cCuentaContable          
           
 --select * from  #TMPDiarioPLEDH          
 --where Per_cPeriodo = @Per_cPeriodoDH and CUO  = @CUO and Pla_cCuentaContable = @Pla_cCuentaContable          
        
 update #TMPDiarioPLEDH          
 set Asd_nHaberSoles = 0          
 where Per_cPeriodo = @Per_cPeriodoDH and CUO = @CUO and Pla_cCuentaContable = @Pla_cCuentaContable and Correlativo = @Correlativo           
          
 set @Nro = @Nro + 1          
          
   FETCH NEXT FROM D_H_Cursor_TC                                                                                 
 INTO  @Per_cPeriodoDH, @CUO,@Correlativo, @PCGE ,@Pla_cCuentaContable ,@Ase_dFecha,@Ase_cGlosa,@Asd_nDebeSoles,@Asd_nHaberSoles,@CUOVentas,@CUOCompras,@CUOConsignacion,@Estado
 , @Mon_cCodSunat, @Asd_cSerieDoc, @Asd_cNumDoc,
  @Asd_dFecVen, @Asd_dFecDoc, @Asd_cTipoDoc, @Lib_cTipoLibro, @Ase_nVoucher, @Ten_cTipoEntidad, @Ent_cCodEntidad, @Ent_nRuc, @Ent_cTipoDoc                                             
          
 END          
CLOSE D_H_Cursor_TC          
DEALLOCATE D_H_Cursor_TC               
          
          
DECLARE @Cu VARCHAR(10)
DECLARE @Corr VARCHAR(10)
DECLARE @Serie VARCHAR(20)
DECLARE @Numero VARCHAR(25)
DECLARE @TipoDoc CHAR(2)
--DECLARE C_Final CURSOR FOR              
--SELECT CUO, Correlativo, RTRIM(LTRIM(Asd_cSerieDoc)), LTRIM(RTRIM(Asd_cNumDoc)), RTRIM(LTRIM(Asd_cTipoDoc)) FROM #TMPDiarioPLEDH
--WHERE RTRIM(LTRIM(Asd_cNumDoc)) <> '' AND RTRIM(LTRIM(Asd_cSerieDoc)) <> ''

--OPEN C_Final
--FETCH NEXT FROM C_Final INTO @Cu, @Corr, @Serie, @Numero, @TipoDoc
--WHILE @@FETCH_STATUS = 0
--BEGIN
--	UPDATE #TMPDiarioPLEDH
--		SET Asd_cSerieDoc = @Serie, Asd_cNumDoc = @Numero, Asd_cTipoDoc = @TipoDoc
--	WHERE CUO = @Cu AND Correlativo = @Corr AND RTRIM(LTRIM(Asd_cSerieDoc)) = '' AND RTRIM(LTRIM(Asd_cNumDoc)) = ''
--	FETCH NEXT FROM C_Final INTO @Cu, @Corr, @Serie, @Numero, @TipoDoc
--END

--SELECT Per_cPeriodo, CUO, Correlativo, PCGE, Pla_cCuentaContable, Ase_dFecha, Ase_cGlosa, Asd_nDebeSoles, Asd_nHaberSoles, CUOVentas, CUOCompras,
--CUOConsignacion, Estado, Mon_cCodSunat, Asd_cSerieDoc, Asd_cNumDoc, Asd_dFecVen, Asd_dFecDoc, Asd_cTipoDoc, Lib_cTipoLibro, Ase_nVoucher FROM #TMPDiarioPLEDH
--RETURN 
	
IF @Simplificado = '1'
BEGIN
	SELECT (D.Per_cPeriodo + @Separador + 
	        D.CUO + @Separador + 
	        --RIGHT(D.Correlativo, 4) + @Separador + 
	        LEFT(D.Correlativo, 5) + RIGHT(CAST(10000 + Corr AS VARCHAR(60)), 5) + @Separador + 
	        D.Pla_cCuentaContable + @Separador + 
	        '' + @Separador + 
	        '' + @Separador + 
	        Mon_cCodSunat + @Separador + 
	        CASE WHEN Lib_cTipoLibro IN ('05', '02') THEN '6'
	             WHEN Lib_cTipoLibro IN ('06', '04') THEN RTRIM(LTRIM(Ent_cTipoDoc)) ELSE '' END + @Separador + 
	        CASE WHEN Lib_cTipoLibro IN ('05', '02') THEN REPLACE(@RUC, 'RUC : ', '')
	             WHEN Lib_cTipoLibro IN ('06', '04') THEN RTRIM(LTRIM(Ent_nRuc)) ELSE '' END + @Separador +
			RTRIM(LTRIM(Asd_cTipoDoc)) + @Separador +
			CASE WHEN LEFT(LTRIM(RTRIM(Asd_cTipoDoc)),2) IN ('50', '52', '54') THEN RIGHT(RTRIM(Asd_cSerieDoc), 3) 
			    WHEN LEFT(LTRIM(RTRIM(Asd_cTipoDoc)),2) IN ('05') THEN RIGHT(RTRIM(Asd_cSerieDoc), 1)
		    ELSE RTRIM(Asd_cSerieDoc) END + @Separador +
			--CASE WHEN /*Lib_cTipoLibro = '06' AND*/ RTRIM(LTRIM(Asd_cTipoDoc)) IN ('50', '52') THEN RIGHT(RTRIM(LTRIM(Asd_cSerieDoc)), 3) ELSE RTRIM(LTRIM(Asd_cSerieDoc)) END + @Separador +  
   --         CASE WHEN /*Lib_cTipoLibro = '05' AND*/ RTRIM(LTRIM(Asd_cTipoDoc)) IN ('01', '04', '07', '08') THEN  RIGHT(RTRIM(LTRIM(Asd_cNumDoc)), 7) 
   --              WHEN /*Lib_cTipoLibro = '05' AND*/ RTRIM(LTRIM(Asd_cTipoDoc)) IN ('00', '10', '12', '13', '14', '15', '16', '17', '18', '19', '21', '22', '24', '26', '27', '28', '29') THEN RIGHT(RTRIM(LTRIM(Asd_cNumDoc)), 20)
   --              WHEN /*Lib_cTipoLibro = '06' AND*/ RTRIM(LTRIM(Asd_cTipoDoc)) IN ('50', '52') THEN RIGHT(RTRIM(LTRIM(Asd_cNumDoc)), 6) 
   --              WHEN RTRIM(LTRIM(Asd_cTipoDoc)) IN ('03', '12') THEN RIGHT(RTRIM(LTRIM(Asd_cNumDoc)), 7)
   --              WHEN RTRIM(LTRIM(Asd_cTipoDoc)) IN ('05') THEN RIGHT(RTRIM(LTRIM(Asd_cNumDoc)), 11)
   --              WHEN RTRIM(LTRIM(Asd_cTipoDoc)) IN ('11') THEN RIGHT(RTRIM(LTRIM(Asd_cNumDoc)), 15)
   --              ELSE RIGHT(RTRIM(LTRIM(Asd_cNumDoc)), 7) END + @Separador +
			CASE WHEN /*Lib_cTipoLibro = '05' AND*/ RTRIM(LTRIM(Asd_cTipoDoc)) IN ('01', '04', '07', '08') THEN  CASE WHEN (CHARINDEX('-', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('-', Asd_cNumDoc)), 7)
                                                                                                                      WHEN (CHARINDEX('/', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('/', Asd_cNumDoc)), 7)
                                                                                                                 ELSE RIGHT(RTRIM(LTRIM(Asd_cNumDoc)), 7)  END
                 WHEN /*Lib_cTipoLibro = '05' AND*/ RTRIM(LTRIM(Asd_cTipoDoc)) IN ('00', '10', '12', '13', '14', '15', '16', '17', '18', '19', '21', '22', '24', '26', '27', '28', '29') THEN CASE WHEN (CHARINDEX('-', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('-', Asd_cNumDoc)), 20)
																																																   WHEN (CHARINDEX('/', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('/', Asd_cNumDoc)), 20)
																																															  ELSE RIGHT(RTRIM(LTRIM(Asd_cNumDoc)), 20) END
                 --WHEN /*Lib_cTipoLibro = '06' AND*/ RTRIM(LTRIM(Asd_cTipoDoc)) IN ('50', '52') THEN RIGHT(RTRIM(LTRIM(Asd_cNumDoc)), 6) 
				 WHEN RTRIM(LTRIM(Asd_cTipoDoc)) IN ('50', '52') THEN CASE WHEN (CHARINDEX('-', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('-', Asd_cNumDoc)), 6)
																		   WHEN (CHARINDEX('/', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('/', Asd_cNumDoc)), 6)
																	  ELSE RIGHT(RTRIM(LTRIM(Asd_cNumDoc)), 6) END
				 WHEN RTRIM(LTRIM(Asd_cTipoDoc)) IN ('03', '12') THEN CASE WHEN (CHARINDEX('-', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('-', Asd_cNumDoc)), 7)
																		   WHEN (CHARINDEX('/', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('/', Asd_cNumDoc)), 7)
																	  ELSE RIGHT(RTRIM(LTRIM(Asd_cNumDoc)), 7) END
				 WHEN RTRIM(LTRIM(Asd_cTipoDoc)) IN ('05') THEN CASE WHEN (CHARINDEX('-', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('-', Asd_cNumDoc)), 11)
																     WHEN (CHARINDEX('/', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('/', Asd_cNumDoc)), 11)
																ELSE RIGHT(RTRIM(LTRIM(Asd_cNumDoc)), 11) END
                 WHEN RTRIM(LTRIM(Asd_cTipoDoc)) IN ('11') THEN CASE WHEN (CHARINDEX('-', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('-', Asd_cNumDoc)), 15)
																	 WHEN (CHARINDEX('/', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('/', Asd_cNumDoc)), 15)
																ELSE RIGHT(RTRIM(LTRIM(Asd_cNumDoc)), 15) END
                 ELSE RIGHT(RTRIM(LTRIM(Asd_cNumDoc)), 7) END + @Separador +
                  
			RTRIM(LTRIM(CONVERT(NCHAR(10), Ase_dFecha, 103))) + @Separador + 
		  CASE WHEN YEAR(CONVERT(NCHAR(10), ISNULL(Asd_dFecVen, ''), 103)) = 1900 THEN '' ELSE CONVERT(NCHAR(10), ISNULL(Asd_dFecVen, ''), 103) END  + @Separador +
		  CONVERT(NCHAR(10), Asd_dFecDoc, 103) + @Separador + Ase_cGlosa + @Separador + '' + @Separador + CAST(CAST(SUM(Asd_nDebeSoles) AS NUMERIC(18, 2)) AS VARCHAR(50)) + @Separador +
		  CAST(CAST(SUM(Asd_nHaberSoles) AS NUMERIC(18, 2)) AS VARCHAR(50)) + @Separador +
		  CASE WHEN Lib_cTipoLibro = '05' THEN '140100&' + convert(varchar(4),year(Ase_dFecha)) + Per_cPeriodo + '&' + LTRIM(RTRIM(LEFT(Ase_nVoucher,40))) + '&' + LEFT(LTRIM(RTRIM('M' + left(Ase_nVoucher,4) + right(Ase_nVoucher,5) )),100) 
			   WHEN Lib_cTipoLibro = '06' THEN '080200&' + CAST(YEAR(Ase_dFecha) AS CHAR(4)) + Per_cPeriodo + '00&' +  LEFT(LTRIM(RTRIM(Ase_nVoucher)),40) + '&' +
			   LEFT(LTRIM(RTRIM('M' + left(Ase_nVoucher,4) + right(Ase_nVoucher,5) )),100) 
			   ELSE '080200&' + CAST(YEAR(Ase_dFecha) AS CHAR(4)) + Per_cPeriodo + '00&' +  LEFT(LTRIM(RTRIM(Ase_nVoucher)),40) + '&' + LEFT(LTRIM(RTRIM('M' + left(Ase_nVoucher,4) + right(Ase_nVoucher,5) )),100)  END + @Separador + Estado + @Separador + 
			   
			   CASE WHEN /*Lib_cTipoLibro = '05' AND*/ RTRIM(LTRIM(Asd_cTipoDoc)) IN ('01', '04', '07', '08') THEN  CASE WHEN (CHARINDEX('-', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, CHARINDEX('-', Asd_cNumDoc) + 1, LEN(Asd_cNumDoc) - CHARINDEX('-', Asd_cNumDoc)), 7) + '|'
                                                                                                                         WHEN (CHARINDEX('/', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, CHARINDEX('/', Asd_cNumDoc) + 1, LEN(Asd_cNumDoc) - CHARINDEX('/', Asd_cNumDoc)), 7) + '|'
                                                                                                                    ELSE '|'  END
                 WHEN /*Lib_cTipoLibro = '05' AND*/ RTRIM(LTRIM(Asd_cTipoDoc)) IN ('00', '10', '12', '13', '14', '15', '16', '17', '18', '19', '21', '22', '24', '26', '27', '28', '29') THEN CASE WHEN (CHARINDEX('-', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, CHARINDEX('-', Asd_cNumDoc) + 1, LEN(Asd_cNumDoc) - CHARINDEX('-', Asd_cNumDoc)), 20) + '|'
																																																   WHEN (CHARINDEX('/', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, CHARINDEX('/', Asd_cNumDoc) + 1, LEN(Asd_cNumDoc) - CHARINDEX('/', Asd_cNumDoc)), 20) + '|'
																																															  ELSE '|' END
                 --WHEN /*Lib_cTipoLibro = '06' AND*/ RTRIM(LTRIM(Asd_cTipoDoc)) IN ('50', '52') THEN RIGHT(RTRIM(LTRIM(Asd_cNumDoc)), 6) 
				 WHEN RTRIM(LTRIM(Asd_cTipoDoc)) IN ('50', '52') THEN CASE WHEN (CHARINDEX('-', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, CHARINDEX('-', Asd_cNumDoc) + 1, LEN(Asd_cNumDoc) - CHARINDEX('-', Asd_cNumDoc)), 6) + '|'
																		   WHEN (CHARINDEX('/', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, CHARINDEX('/', Asd_cNumDoc) + 1, LEN(Asd_cNumDoc) - CHARINDEX('/', Asd_cNumDoc)), 6) + '|'
																	  ELSE '|' END
				 WHEN RTRIM(LTRIM(Asd_cTipoDoc)) IN ('03', '12') THEN CASE WHEN (CHARINDEX('-', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, CHARINDEX('-', Asd_cNumDoc) + 1, LEN(Asd_cNumDoc) - CHARINDEX('-', Asd_cNumDoc)), 7) + '|'
																		   WHEN (CHARINDEX('/', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, CHARINDEX('/', Asd_cNumDoc) + 1, LEN(Asd_cNumDoc) - CHARINDEX('/', Asd_cNumDoc)), 7) + '|'
																	  ELSE '|' END
				 WHEN RTRIM(LTRIM(Asd_cTipoDoc)) IN ('05') THEN CASE WHEN (CHARINDEX('-', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, CHARINDEX('-', Asd_cNumDoc) + 1, LEN(Asd_cNumDoc) - CHARINDEX('-', Asd_cNumDoc)), 11) + '|'
																     WHEN (CHARINDEX('/', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, CHARINDEX('/', Asd_cNumDoc) + 1, LEN(Asd_cNumDoc) - CHARINDEX('/', Asd_cNumDoc)), 11) + '|'
																ELSE '|' END
                 WHEN RTRIM(LTRIM(Asd_cTipoDoc)) IN ('11') THEN CASE WHEN (CHARINDEX('-', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, CHARINDEX('-', Asd_cNumDoc) + 1, LEN(Asd_cNumDoc) - CHARINDEX('-', Asd_cNumDoc)), 15) + '|'
																	 WHEN (CHARINDEX('/', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, CHARINDEX('/', Asd_cNumDoc) + 1, LEN(Asd_cNumDoc) - CHARINDEX('/', Asd_cNumDoc)), 15) + '|'
																ELSE '|' END
                 ELSE '|' END 
			   ) AS Registro   FROM #TMPDiarioPLEDH D	
	group by Per_cPeriodo, CUO ,Correlativo ,PCGE, Pla_cCuentaContable, Ase_dFecha, Ase_cGlosa,CUOVentas,CUOCompras,CUOConsignacion, Estado,
		   Asd_cSerieDoc, Asd_cNumDoc,
			   Asd_dFecVen, Asd_cTipoDoc, Lib_cTipoLibro, Ase_nVoucher, Mon_cCodSunat, Asd_dFecDoc, Corr, Ent_cTipoDoc, Ent_nRuc
			   order by  RIGHT(CUO ,3), Ase_dFecha
	
	--SELECT D.Per_cPeriodo AS Campo1, D.CUO  AS Campo2, D.Correlativo AS Campo3, D.Pla_cCuentaContable AS Campo4 , '' AS Campo5 , '' AS Campo6 , Mon_cCodSunat AS Campo7 , '6' AS Campo8, REPLACE(@RUC, 'RUC : ', '') AS Campo9,
	--		Lib_cTipoLibro, RTRIM(LTRIM(Asd_cTipoDoc)) AS Campo10, RTRIM(LTRIM(Asd_cSerieDoc)) AS Campo11 ,  RIGHT(RTRIM(LTRIM(Asd_cNumDoc)), 7) AS Campo12, RTRIM(LTRIM(CONVERT(NCHAR(10), Ase_dFecha, 103))) AS Campo13 , 
	--	  CASE WHEN YEAR(CONVERT(NCHAR(10), ISNULL(Asd_dFecVen, ''), 103)) = 1900 THEN '' ELSE CONVERT(NCHAR(10), ISNULL(Asd_dFecVen, ''), 103) END  AS Campo14,
	--	  CONVERT(NCHAR(10), Asd_dFecDoc, 103) AS Campo15 , Ase_cGlosa AS Campo16, '' AS Campo17 , CAST(CAST(SUM(Asd_nDebeSoles) AS NUMERIC(18, 2)) AS VARCHAR(50)) AS Campo18 ,
	--	  CAST(CAST(SUM(Asd_nHaberSoles) AS NUMERIC(18, 2)) AS VARCHAR(50)) AS Campo19 ,
	--	  CASE WHEN Lib_cTipoLibro = '05' THEN '140100&' + convert(varchar(4),year(Ase_dFecha)) + Per_cPeriodo + '&' + LTRIM(RTRIM(LEFT(Ase_nVoucher,40))) + '&' + LEFT(LTRIM(RTRIM('M' + left(Ase_nVoucher,4) + right(Ase_nVoucher,5) )),100) 
	--		   WHEN Lib_cTipoLibro = '06' THEN '080200&' + CAST(YEAR(Ase_dFecha) AS CHAR(4)) + Per_cPeriodo + '00&' +  LEFT(LTRIM(RTRIM(Ase_nVoucher)),40) + '&' +
	--		   LEFT(LTRIM(RTRIM('M' + left(Ase_nVoucher,4) + right(Ase_nVoucher,5) )),100) ELSE '' END AS Campo20 , Estado AS Campo21   FROM #TMPDiarioPLEDH D	
	--group by Per_cPeriodo, CUO ,Correlativo ,PCGE, Pla_cCuentaContable, Ase_dFecha, Ase_cGlosa,CUOVentas,CUOCompras,CUOConsignacion, Estado,
	--	   Asd_cSerieDoc, Asd_cNumDoc,
	--		   Asd_dFecVen, Asd_cTipoDoc, Lib_cTipoLibro, Ase_nVoucher, Mon_cCodSunat, Asd_dFecDoc
	--		   order by  RIGHT(CUO ,3), Ase_dFecha
	
END
ELSE IF @Simplificado = '0'
BEGIN
	
	SELECT (D.Per_cPeriodo + @Separador + 
	        D.CUO + @Separador + 
	        LEFT(D.Correlativo, 5) + RIGHT(CAST(10000 + Corr AS VARCHAR(60)), 5) + @Separador + 	
	        D.Pla_cCuentaContable + @Separador + '' + @Separador + '' + @Separador +
	        Mon_cCodSunat + @Separador +
	        CASE WHEN Lib_cTipoLibro IN ('05', '02') THEN '6'
	             WHEN Lib_cTipoLibro IN ('06', '04') THEN RTRIM(LTRIM(Ent_cTipoDoc)) ELSE '' END + @Separador + 
	        CASE WHEN Lib_cTipoLibro IN ('05', '02') THEN REPLACE(@RUC, 'RUC : ', '')
	             WHEN Lib_cTipoLibro IN ('06', '04') THEN RTRIM(LTRIM(Ent_nRuc)) ELSE '' END + @Separador +
            RTRIM(LTRIM(Asd_cTipoDoc)) + @Separador + 
            
            CASE WHEN Lib_cTipoLibro <> '05' AND RTRIM(LTRIM(Asd_cTipoDoc)) IN ('50', '52', '54') THEN RIGHT(RTRIM(LTRIM(Asd_cSerieDoc)), 3) 
                 WHEN RTRIM(LTRIM(Asd_cTipoDoc)) IN ('50', '52') THEN RIGHT(RTRIM(LTRIM(Asd_cSerieDoc)), 3)
                 WHEN RTRIM(LTRIM(Asd_cTipoDoc)) IN ('05') THEN RIGHT(RTRIM(LTRIM(Asd_cSerieDoc)), 1)
            ELSE RTRIM(LTRIM(Asd_cSerieDoc)) END + @Separador +  
            
            CASE WHEN /*Lib_cTipoLibro = '05' AND*/ RTRIM(LTRIM(Asd_cTipoDoc)) IN ('01', '04', '07', '08') THEN  CASE WHEN (CHARINDEX('-', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('-', Asd_cNumDoc)), 7)
                                                                                                                      WHEN (CHARINDEX('/', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('/', Asd_cNumDoc)), 7)
                                                                                                                 ELSE RIGHT(RTRIM(LTRIM(Asd_cNumDoc)), 7)  END
                 WHEN /*Lib_cTipoLibro = '05' AND*/ RTRIM(LTRIM(Asd_cTipoDoc)) IN ('00', '10', '12', '13', '14', '15', '16', '17', '18', '19', '21', '22', '24', '26', '27', '28', '29') THEN CASE WHEN (CHARINDEX('-', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('-', Asd_cNumDoc)), 20)
																																																   WHEN (CHARINDEX('/', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('/', Asd_cNumDoc)), 20)
																																															  ELSE RIGHT(RTRIM(LTRIM(Asd_cNumDoc)), 20) END
                 --WHEN /*Lib_cTipoLibro = '06' AND*/ RTRIM(LTRIM(Asd_cTipoDoc)) IN ('50', '52') THEN RIGHT(RTRIM(LTRIM(Asd_cNumDoc)), 6) 
				 WHEN RTRIM(LTRIM(Asd_cTipoDoc)) IN ('50', '52') THEN CASE WHEN (CHARINDEX('-', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('-', Asd_cNumDoc)), 6)
																		   WHEN (CHARINDEX('/', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('/', Asd_cNumDoc)), 6)
																	  ELSE RIGHT(RTRIM(LTRIM(Asd_cNumDoc)), 6) END
				 WHEN RTRIM(LTRIM(Asd_cTipoDoc)) IN ('03', '12') THEN CASE WHEN (CHARINDEX('-', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('-', Asd_cNumDoc)), 7)
																		   WHEN (CHARINDEX('/', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('/', Asd_cNumDoc)), 7)
																	  ELSE RIGHT(RTRIM(LTRIM(Asd_cNumDoc)), 7) END
				 WHEN RTRIM(LTRIM(Asd_cTipoDoc)) IN ('05') THEN CASE WHEN (CHARINDEX('-', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('-', Asd_cNumDoc)), 11)
																     WHEN (CHARINDEX('/', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('/', Asd_cNumDoc)), 11)
																ELSE RIGHT(RTRIM(LTRIM(Asd_cNumDoc)), 11) END
                 WHEN RTRIM(LTRIM(Asd_cTipoDoc)) IN ('11') THEN CASE WHEN (CHARINDEX('-', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('-', Asd_cNumDoc)), 15)
																	 WHEN (CHARINDEX('/', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('/', Asd_cNumDoc)), 15)
																ELSE RIGHT(RTRIM(LTRIM(Asd_cNumDoc)), 15) END
                 ELSE RIGHT(RTRIM(LTRIM(Asd_cNumDoc)), 7) END + @Separador + 
                 
            RTRIM(LTRIM(CONVERT(NCHAR(10), Ase_dFecha, 103))) + @Separador + 
            CASE WHEN YEAR(CONVERT(NCHAR(10), ISNULL(Asd_dFecVen, ''), 103)) = 1900 THEN '' ELSE CONVERT(NCHAR(10), ISNULL(Asd_dFecVen, ''), 103) END  + @Separador +
            CONVERT(NCHAR(10), Asd_dFecDoc, 103) + @Separador + Ase_cGlosa + @Separador + '' + @Separador + CAST(CAST(SUM(Asd_nDebeSoles) AS NUMERIC(18, 2)) AS VARCHAR(50)) + @Separador +
            CAST(CAST(SUM(Asd_nHaberSoles) AS NUMERIC(18, 2)) AS VARCHAR(50)) + @Separador +
            CASE WHEN Lib_cTipoLibro = '05' THEN '140100&' + convert(varchar(4),year(Ase_dFecha)) + Per_cPeriodo + '&' + LTRIM(RTRIM(LEFT(Ase_nVoucher,40))) + '&' + LEFT(LTRIM(RTRIM('M' + left(Ase_nVoucher,4) + right(Ase_nVoucher,5) )),100) 
                 WHEN Lib_cTipoLibro = '06' THEN '080200&' + CAST(YEAR(Ase_dFecha) AS CHAR(4)) + Per_cPeriodo + '00&' +  LEFT(LTRIM(RTRIM(Ase_nVoucher)),40) + '&' + LEFT(LTRIM(RTRIM('M' + left(Ase_nVoucher,4) + right(Ase_nVoucher,5) )),100) 
		   
		   ELSE '080200&' + CAST(YEAR(Ase_dFecha) AS CHAR(4)) + Per_cPeriodo + '00&' +  LEFT(LTRIM(RTRIM(Ase_nVoucher)),40) + '&' + LEFT(LTRIM(RTRIM('M' + left(Ase_nVoucher,4) + right(Ase_nVoucher,5) )),100)  END + @Separador + Estado + @Separador + 
		   
		   CASE WHEN /*Lib_cTipoLibro = '05' AND*/ RTRIM(LTRIM(Asd_cTipoDoc)) IN ('01', '04', '07', '08') THEN  CASE WHEN (CHARINDEX('-', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, CHARINDEX('-', Asd_cNumDoc) + 1, LEN(Asd_cNumDoc) - CHARINDEX('-', Asd_cNumDoc)), 7) + '|'
                                                                                                                     WHEN (CHARINDEX('/', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, CHARINDEX('/', Asd_cNumDoc) + 1, LEN(Asd_cNumDoc) - CHARINDEX('/', Asd_cNumDoc)), 7) + '|'
                                                                                                                 ELSE '|'  END
                 WHEN /*Lib_cTipoLibro = '05' AND*/ RTRIM(LTRIM(Asd_cTipoDoc)) IN ('00', '10', '12', '13', '14', '15', '16', '17', '18', '19', '21', '22', '24', '26', '27', '28', '29') THEN CASE WHEN (CHARINDEX('-', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, CHARINDEX('-', Asd_cNumDoc) + 1, LEN(Asd_cNumDoc) - CHARINDEX('-', Asd_cNumDoc)), 20) + '|'
																																																   WHEN (CHARINDEX('/', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, CHARINDEX('/', Asd_cNumDoc) + 1, LEN(Asd_cNumDoc) - CHARINDEX('/', Asd_cNumDoc)), 20) + '|'
																																															  ELSE '|' END
                 --WHEN /*Lib_cTipoLibro = '06' AND*/ RTRIM(LTRIM(Asd_cTipoDoc)) IN ('50', '52') THEN RIGHT(RTRIM(LTRIM(Asd_cNumDoc)), 6) 
				 WHEN RTRIM(LTRIM(Asd_cTipoDoc)) IN ('50', '52') THEN CASE WHEN (CHARINDEX('-', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, CHARINDEX('-', Asd_cNumDoc) + 1, LEN(Asd_cNumDoc) - CHARINDEX('-', Asd_cNumDoc)), 6) + '|'
																		   WHEN (CHARINDEX('/', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, CHARINDEX('/', Asd_cNumDoc) + 1, LEN(Asd_cNumDoc) - CHARINDEX('/', Asd_cNumDoc)), 6) + '|'
																	  ELSE '|' END
				 WHEN RTRIM(LTRIM(Asd_cTipoDoc)) IN ('03', '12') THEN CASE WHEN (CHARINDEX('-', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, CHARINDEX('-', Asd_cNumDoc) + 1, LEN(Asd_cNumDoc) - CHARINDEX('-', Asd_cNumDoc)), 7) + '|'
																		   WHEN (CHARINDEX('/', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, CHARINDEX('/', Asd_cNumDoc) + 1, LEN(Asd_cNumDoc) - CHARINDEX('/', Asd_cNumDoc)), 7) + '|'
																	  ELSE '|' END
				 WHEN RTRIM(LTRIM(Asd_cTipoDoc)) IN ('05') THEN CASE WHEN (CHARINDEX('-', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, CHARINDEX('-', Asd_cNumDoc) + 1, LEN(Asd_cNumDoc) - CHARINDEX('-', Asd_cNumDoc)), 11) + '|'
																     WHEN (CHARINDEX('/', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, CHARINDEX('/', Asd_cNumDoc) + 1, LEN(Asd_cNumDoc) - CHARINDEX('/', Asd_cNumDoc)), 11) + '|'
																ELSE '|' END
                 WHEN RTRIM(LTRIM(Asd_cTipoDoc)) IN ('11') THEN CASE WHEN (CHARINDEX('-', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, CHARINDEX('-', Asd_cNumDoc) + 1, LEN(Asd_cNumDoc) - CHARINDEX('-', Asd_cNumDoc)), 15) + '|'
																	 WHEN (CHARINDEX('/', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, CHARINDEX('/', Asd_cNumDoc) + 1, LEN(Asd_cNumDoc) - CHARINDEX('/', Asd_cNumDoc)), 15) + '|'
																ELSE '|' END
                 ELSE '|' END 
		   
		   ) AS Registro  FROM #TMPDiarioPLEDH D
group by Per_cPeriodo, D.CUO ,D.Correlativo, Corr ,PCGE, Pla_cCuentaContable, Ase_dFecha, Ase_cGlosa,CUOVentas,CUOCompras,CUOConsignacion, Estado,
       Asd_cSerieDoc, Asd_cNumDoc,
           Asd_dFecVen, Asd_cTipoDoc, Lib_cTipoLibro, Ase_nVoucher, Mon_cCodSunat, Asd_dFecDoc, Ent_cTipoDoc, Ent_nRuc
           order by  RIGHT(D.CUO ,3), Ase_dFecha
END

                                
--select (Per_cPeriodo + @Separador +                    
--      CUO   + @Separador +           
--      Correlativo  + @Separador +               
--      PCGE + @Separador +                    
--      Pla_cCuentaContable  + @Separador +                    
--      LTRIM(RTRIM(LEFT(convert(varchar(10),Ase_dFecha,103),10)))  + @Separador +                    
--      Ase_cGlosa + @Separador +                    
--      LTRIM(RTRIM(LEFT(convert(varchar(15),CONVERT(money,SUM(Asd_nDebeSoles)),0),15)))  + @Separador +                    
--      LTRIM(RTRIM(LEFT(convert(varchar(15),CONVERT(money,SUM(Asd_nHaberSoles)),0),15))) + @Separador +            
--      CUOVentas  + @Separador +               
--      CUOCompras  + @Separador +               
--      CUOConsignacion + @Separador +                                  
--      Estado + @Separador) AS 'Registro'                    
--from #TMPDiarioPLEDH                    
--group by Per_cPeriodo, CUO ,Correlativo ,PCGE, Pla_cCuentaContable, Ase_dFecha, Ase_cGlosa,CUOVentas,CUOCompras,CUOConsignacion, Estado                                
--order by  RIGHT(CUO ,3), Ase_dFecha
GO
