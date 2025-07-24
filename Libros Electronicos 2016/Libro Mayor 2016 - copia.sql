USE SAFC_ECB
GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO
ALTER PROCEDURE [dbo].[spCn_RptMayorElectronico3]
 @Emp_cCodigo char(3)='',              
 @Pan_cAnio char(4)='',              
 @Per_cPeriodoDesde char(2)='',              
 @Per_cPeriodoHasta char(2)='',              
 @moneda char(3) = '',              
 @CtaDesde varchar(12) = '',              
 @CtaHasta varchar(12) = '',              
 @Tipo_Lib_Mayor Char(1),              
 @Per_Fechadesde varchar(10) = '',              
 @Per_Fechahasta varchar(10) = ''              
--WITH ENCRYPTION                                                              
AS                                                              
SET NOCOUNT ON                                                          
SET DATEFORMAT DMY                                                              
                                                      
--select @Emp_cCodigo ='001',                                                      
-- @Pan_cAnio ='2013',                                                      
-- @Per_cPeriodoDesde ='01',                                                      
-- @Per_cPeriodoHasta ='02',                                                      
-- @moneda = '038',                                                         
-- @CtaDesde = '',                                                      
-- @CtaHasta = ''                                                      
                                                              
Declare @sql varchar(MAX)                                                               
--------------------------------------------------------------------------                                                                          
declare @nValor numeric(14,3)                                                                          
set @nValor = 0                                                                          
--------------------------------------------------------------------------                                                                          
declare @RUC char(50)                                                                          
select @RUC= dbo.RUC(@Emp_cCodigo) 
SET @RUC = REPLACE(@RUC, 'RUC : ', '')                                                                         
--------------------------------------------------------------------------                                                                          
Declare @Mon_cNombreLargo varchar(100)                                                                          
Declare @Mon_cMNac char(1)                                                              
Declare @Separador varchar(1)                    
declare @Per_cPeriodo varchar(2)                    
                                                          
 Set @Separador = '|'                                                                 
--------------------------------------------------------------------------                                                                          
SELECT @Mon_cNombreLargo = Mon_cNombreLargo, @Mon_cMNac = Mon_cMNac                                                                           
FROM CNT_TIPO_MONEDA                                                
WHERE Emp_cCodigo = @Emp_cCodigo and Mon_cCodigo = @moneda                                                                          
--------------------------------------------------------------------------                                  
IF DBO.TRIMSQL(@CtaDesde) = ''                          
BEGIN                                                      
 SET @CtaDesde = '10'                                
 SET @CtaHasta = '999999999999'                                                      
END                              
                    
set @Per_cPeriodo = @Per_cPeriodoDesde                    
                                                      
if @Per_cPeriodoDesde = '01'                                              
 set @Per_cPeriodoDesde = '00'                                                      
                                                   
if @Per_cPeriodoDesde = '12'                                                    
 set @Per_cPeriodoHasta = '14'                                                    
                                                                          
SET @CtaHasta = DBO.TRIMSQL(@CtaHasta)  + '999999999999'                                                                          
                                                              
Declare @tabla as varchar(35)                                                       
Declare @tabla2 as varchar(35)                                                       
Declare @Query as nvarchar(max)                                                      
Declare @randon as int                                                              
Declare @randon2 as int                                                       
SELECT  @randon=abs(checksum(newid()))                                                              
SELECT  @randon2=abs(checksum(newid()))                                                              
                                            
set @tabla='TMPASIENTOS'+ cast(@randon as varchar(100))                                                      
                                                      
set @tabla2='TMPASIENTOS'+ cast(@randon2 as varchar(100))                                                      
                                                    
                                                      
 set @Query = 'CREATE TABLE ' + @tabla + '(                                                                          
Pla_cCuentaContable varchar(30),CUO varchar(30),Correlativo varchar(30), D3 varchar(10), D2 varchar(10), D2_sumas varchar(10),                                                              
 SaldoAntMonNac numeric(14,3), SaldoAntMonExt numeric(14,3), Asd_nDebeSoles numeric(14,3), Asd_nDebeMonExt numeric(14,3),                                                              
 Asd_nHaberSoles numeric(14,3), Asd_nHaberMonExt numeric(14,3), Emp_cCodigo varchar(5), Pan_cAnio varchar(10),                                                              
 Per_cPeriodo varchar(20), Lib_cTipoLibro varchar(20),  Lib_cDescripcion varchar(350), Ase_nVoucher varchar(20),                                                           
 Asd_cTipoMoneda varchar(20), Asd_nItem int, Ase_cGlosa varchar(350), Asd_nTipoCambio numeric(14,3),                                                              
 Cos_cCodigo varchar(20), Ten_cTipoEntidad varchar(20), Ent_cCodEntidad varchar(20), Ent_cPersona varchar(120),                                                              
 Ten_cNombreEntidad varchar(250), Asd_cTipoDoc varchar(20), Asd_cSerieDoc varchar(20), Asd_cNumDoc char(25),                                                              
 Asd_dFecDoc datetime, Asd_cTipoDocRef varchar(20), Asd_cSerieDocRef varchar(20), Asd_cNumDocRef char(50),                                                              
 Asd_dFecDocRef datetime, Asd_nMontoInafecto numeric(14,3), Asd_cRetencion char(1), Asd_dFechaSpot datetime,                                                              
 Asd_cNumSpot varchar(50), Asd_cDestino varchar(20), Asd_nCorre int, Tab_cCodSunat varchar(20), Mon_cMNac varchar(20),                                                      
 Asd_cEstadoO char(1),Asd_cEstadoD char(1), Id_Aduana char(10))'                                                      
 exec(@Query)
                 
 set @Query = 'CREATE TABLE ' + @tabla2 + '(                                                                          
 Pla_cCuentaContable varchar(15),CUO varchar(15),Correlativo varchar(20), D3 char(3), D2 char(2), D2_sumas char(2),           
 SaldoAntMonNac numeric(14,3), SaldoAntMonExt numeric(14,3), Asd_nDebeSoles numeric(14,3), Asd_nDebeMonExt numeric(14,3),                                                              
Asd_nHaberSoles numeric(14,3), Asd_nHaberMonExt numeric(14,3), Emp_cCodigo char(3), Pan_cAnio char(4),                 
 Per_cPeriodo char(2), Lib_cTipoLibro char(2),  Lib_cDescripcion varchar(80), Ase_nVoucher char(10),                                                          
 Asd_cTipoMoneda char(3), Asd_nItem int, Ase_cGlosa varchar(250), Asd_nTipoCambio numeric(14,3),                                                              
 Cos_cCodigo varchar(12), Ten_cTipoEntidad char(1), Ent_cCodEntidad char(15), Ent_cPersona varchar(120),                                                              
 Ten_cNombreEntidad varchar(120), Asd_cTipoDoc char(3), Asd_cSerieDoc varchar(20), Asd_cNumDoc varchar(25),                                                           
Asd_dFecDoc datetime, Asd_cTipoDocRef char(3), Asd_cSerieDocRef char(5), Asd_cNumDocRef char(20),                                                   
 Asd_dFecDocRef datetime, Asd_nMontoInafecto numeric(14,3), Asd_cRetencion char(1), Asd_dFechaSpot datetime,                                                              
 Asd_cNumSpot char(25), Asd_cDestino char(1), Asd_nCorre int, Tab_cCodSunat varchar(10), Mon_cMNac char(1),                                                      
 Pla_cNombreCuenta varchar(120), Pla_cNombreCuentaD3 varchar(120),Pla_cNombreCuentaD2 varchar(100),Asd_cEstadoO char(1),Asd_cEstadoD char(1), Id_Aduana char(10))'                                                      
 exec(@Query)                                                      
                                                     
--------------------------------------------------------------------------                                                                          
--- MOVIMIENTOS DEL PERIODO INICIAL AL FINAL                                                                          
-- SELECT TOP 5 * FROM CND_ASIENTO_VOUCHER where Pla_cCuentaContable ='42101000'                                                                 
set @sql = ' SELECT CND_ASIENTO_VOUCHER.Pla_cCuentaContable,              
cnc_asiento_voucher.Ase_nVoucher as CUO,              
  case when left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) = ''0100'' then              
  LEFT(LTRIM(RTRIM(''A'' + left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + right(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,5) )),100)              
  when left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) = ''0813'' or left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) = ''0814'' then              
  LEFT(LTRIM(RTRIM(''C'' + left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + right(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,5) )),100)              
  else LEFT(LTRIM(RTRIM(''M'' + left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + right(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,5) )),100)              
  end As Correlativo,              
LEFT (CND_ASIENTO_VOUCHER.Pla_cCuentaContable, 3) AS ''D3'',                                                              
LEFT (CND_ASIENTO_VOUCHER.Pla_cCuentaContable, 2) AS ''D2'' ,LEFT (CND_ASIENTO_VOUCHER.Pla_cCuentaContable, 2) AS ''D2_SUMAS'','                                                              
+ convert(varchar(12),@nValor) + ' as ''SaldoAntMonNac'',' + convert(varchar(12),@nValor) + ' as ''SaldoAntMonExt'',                                                 
CND_ASIENTO_VOUCHER.Asd_nDebeSoles,CND_ASIENTO_VOUCHER.Asd_nDebeMonExt, CND_ASIENTO_VOUCHER.Asd_nHaberSoles,CND_ASIENTO_VOUCHER.Asd_nHaberMonExt,              
CND_ASIENTO_VOUCHER.Emp_cCodigo,CND_ASIENTO_VOUCHER.Pan_cAnio, CND_ASIENTO_VOUCHER.Per_cPeriodo,CND_ASIENTO_VOUCHER.Lib_cTipoLibro,                                                              
CNT_LIBRO_OPERA.Lib_cDescripcion,CND_ASIENTO_VOUCHER.Ase_nVoucher,CND_ASIENTO_VOUCHER.Asd_cTipoMoneda,CND_ASIENTO_VOUCHER.Asd_nItem,                                                              
CNc_ASIENTO_VOUCHER.Ase_cGlosa,CND_ASIENTO_VOUCHER.Asd_nTipoCambio,CND_ASIENTO_VOUCHER.Cos_cCodigo,CND_ASIENTO_VOUCHER.Ten_cTipoEntidad,                   
CNM_ENTIDAD.Ent_nRuc as Ent_cCodEntidad,CNM_ENTIDAD.Ent_cPersona,CNT_ENTIDAD.Ten_cNombreEntidad,CND_ASIENTO_VOUCHER.Asd_cTipoDoc,                                                              
CND_ASIENTO_VOUCHER.Asd_cSerieDoc,CND_ASIENTO_VOUCHER.Asd_cNumDoc,convert(varchar(10), (case when year(CNc_ASIENTO_VOUCHER.Ase_dFecha)=1900 then null else CNc_ASIENTO_VOUCHER.Ase_dFecha end) ,103) as Asd_dFecDoc,                        
CND_ASIENTO_VOUCHER.Asd_cTipoDocRef,CND_ASIENTO_VOUCHER.Asd_cSerieDocRef, CND_ASIENTO_VOUCHER.Asd_cNumDocRef,CND_ASIENTO_VOUCHER.Asd_dFecDocRef,                                                              
CND_ASIENTO_VOUCHER.Asd_nMontoInafecto,CND_ASIENTO_VOUCHER.Asd_cRetencion,CND_ASIENTO_VOUCHER.Asd_dFechaSpot,CND_ASIENTO_VOUCHER.Asd_cNumSpot,                                                              
CND_ASIENTO_VOUCHER.Asd_cDestino,CND_ASIENTO_VOUCHER.Asd_nCorre,TABLA.Tab_cCodSunat,''' + @Mon_cMNac + ''' as Mon_cMNac,                                                      
Asd_cEstadoO, Asd_cEstadoD, CND_ASIENTO_VOUCHER.Id_Aduana                                    
FROM CND_ASIENTO_VOUCHER LEFT JOIN CNc_ASIENTO_VOUCHER ON CNC_ASIENTO_VOUCHER.Emp_cCodigo = CND_ASIENTO_VOUCHER.Emp_cCodigo AND                                                    
CNC_ASIENTO_VOUCHER.Pan_cAnio = CND_ASIENTO_VOUCHER.Pan_cAnio AND CNC_ASIENTO_VOUCHER.Per_cPeriodo = CND_ASIENTO_VOUCHER.Per_cPeriodo AND                                      
CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = CND_ASIENTO_VOUCHER.Lib_cTipoLibro AND CNC_ASIENTO_VOUCHER.Ase_nVoucher = CND_ASIENTO_VOUCHER.Ase_nVoucher                                                    
LEFT JOIN CNM_ENTIDAD ON CND_ASIENTO_VOUCHER.Emp_cCodigo = CNM_ENTIDAD.Emp_cCodigo AND                                                    
CND_ASIENTO_VOUCHER.Ten_cTipoEntidad = CNM_ENTIDAD.Ten_cTipoEntidad AND CND_ASIENTO_VOUCHER.Ent_cCodEntidad = CNM_ENTIDAD.Ent_cCodEntidad                                                    
LEFT  JOIN TABLA ON  CNM_ENTIDAD.Emp_cCodigo = TABLA.Emp_cCodigo AND CNM_ENTIDAD.Ent_cTipoDoc = TABLA.Tab_cCodigo AND                                                    
TABLA.tab_ctabla = ''003'' LEFT JOIN CNT_LIBRO_OPERA ON CND_ASIENTO_VOUCHER.Lib_cTipoLibro = CNT_LIBRO_OPERA.Lib_cTipoLibro AND                                                    
CND_ASIENTO_VOUCHER.Pan_cAnio = CNT_LIBRO_OPERA.Pan_cAnio AND CND_ASIENTO_VOUCHER.Emp_cCodigo = CNT_LIBRO_OPERA.Emp_cCodigo                                                    
LEFT JOIN CNT_ENTIDAD ON CNM_ENTIDAD.Ten_cTipoEntidad = CNT_ENTIDAD.Ten_cTipoEntidad AND                                                    
CNM_ENTIDAD.Emp_cCodigo = CNT_ENTIDAD.Emp_cCodigo                                    
WHERE   CND_ASIENTO_VOUCHER.Asd_cDeleted <> ''*'' and CNC_ASIENTO_VOUCHER.AsE_cDeleted <> ''*'' and                                                                 
CND_ASIENTO_VOUCHER.Emp_cCodigo = ''' + @Emp_cCodigo + '''  AND  CND_ASIENTO_VOUCHER.Pan_cAnio = ''' + @Pan_cAnio + ''' and                                                                          
CND_ASIENTO_VOUCHER.Pla_cCuentaContable  >= ''' + @CtaDesde + ''' AND CND_ASIENTO_VOUCHER.Pla_cCuentaContable  <= ''' + @CtaHasta + ''''       
if @Per_cPeriodoDesde = ''                                                           
begin                                                              
 set @sql = @sql + ' AND CNC_ASIENTO_VOUCHER.Ase_dFecha between ''' + @Per_Fechadesde + ''' and ''' + @Per_Fechahasta + ''''                                                              
 set @Per_cPeriodoDesde = month(@Per_Fechadesde)                                                                 
end                                        
else                                  
begin                                        
 set @sql = @sql + ' AND CND_ASIENTO_VOUCHER.Per_cPeriodo >= ''' + @Per_cPeriodoDesde + ''' and CND_ASIENTO_VOUCHER.Per_cPeriodo <= ''' + @Per_cPeriodoHasta + ''''                                  
end                                                         
                                                      
set @sql = 'insert into ' + @tabla + @sql
exec (@sql)                                                      
                                                    
--Agrego los estados 8                                                   
                                                    
set @sql = ' SELECT CND_ASIENTO_VOUCHER.Pla_cCuentaContable,              
cnc_asiento_voucher.Ase_nVoucher as CUO,              
  case when left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) = ''0100'' then              
  LEFT(LTRIM(RTRIM(''A'' + left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + right(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,5) )),100)              
  when left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) = ''0813'' or left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) = ''0814'' then              
  LEFT(LTRIM(RTRIM(''C'' + left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + right(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,5) )),100)              
  else LEFT(LTRIM(RTRIM(''M'' + left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + right(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,5) )),100)              
  end As Correlativo,              
LEFT (CND_ASIENTO_VOUCHER.Pla_cCuentaContable, 3) AS ''D3'',                                                              
LEFT (CND_ASIENTO_VOUCHER.Pla_cCuentaContable, 2) AS ''D2'' ,LEFT (CND_ASIENTO_VOUCHER.Pla_cCuentaContable, 2) AS ''D2_SUMAS'','                                                              
+ convert(varchar(12),@nValor) + ' as ''SaldoAntMonNac'',' + convert(varchar(12),@nValor) + ' as ''SaldoAntMonExt'',                                                              
CND_ASIENTO_VOUCHER.Asd_nDebeSoles,CND_ASIENTO_VOUCHER.Asd_nDebeMonExt, CND_ASIENTO_VOUCHER.Asd_nHaberSoles,CND_ASIENTO_VOUCHER.Asd_nHaberMonExt,                                                              
CND_ASIENTO_VOUCHER.Emp_cCodigo,CND_ASIENTO_VOUCHER.Pan_cAnio, CND_ASIENTO_VOUCHER.Per_cPeriodo,CND_ASIENTO_VOUCHER.Lib_cTipoLibro,                                                              
CNT_LIBRO_OPERA.Lib_cDescripcion,CND_ASIENTO_VOUCHER.Ase_nVoucher,CND_ASIENTO_VOUCHER.Asd_cTipoMoneda,CND_ASIENTO_VOUCHER.Asd_nItem,                              
CNc_ASIENTO_VOUCHER.Ase_cGlosa,CND_ASIENTO_VOUCHER.Asd_nTipoCambio,CND_ASIENTO_VOUCHER.Cos_cCodigo,CND_ASIENTO_VOUCHER.Ten_cTipoEntidad,                                                              
CNM_ENTIDAD.Ent_nRuc as Ent_cCodEntidad,CNM_ENTIDAD.Ent_cPersona,CNT_ENTIDAD.Ten_cNombreEntidad,CND_ASIENTO_VOUCHER.Asd_cTipoDoc,                                                              
CND_ASIENTO_VOUCHER.Asd_cSerieDoc,CND_ASIENTO_VOUCHER.Asd_cNumDoc,convert(varchar(10), (case when year(CNc_ASIENTO_VOUCHER.Ase_dFecha)=1900 then null else CNc_ASIENTO_VOUCHER.Ase_dFecha end) ,103) as Asd_dFecDoc,                        
CND_ASIENTO_VOUCHER.Asd_cTipoDocRef,CND_ASIENTO_VOUCHER.Asd_cSerieDocRef, CND_ASIENTO_VOUCHER.Asd_cNumDocRef,CND_ASIENTO_VOUCHER.Asd_dFecDocRef,                                                        
CND_ASIENTO_VOUCHER.Asd_nMontoInafecto,CND_ASIENTO_VOUCHER.Asd_cRetencion,CND_ASIENTO_VOUCHER.Asd_dFechaSpot,CND_ASIENTO_VOUCHER.Asd_cNumSpot,                                                              
CND_ASIENTO_VOUCHER.Asd_cDestino,CND_ASIENTO_VOUCHER.Asd_nCorre,TABLA.Tab_cCodSunat,''' + @Mon_cMNac + ''' as Mon_cMNac,                                                      
Asd_cEstadoO, Asd_cEstadoD, CND_ASIENTO_VOUCHER.Id_Aduana                                     
FROM CND_ASIENTO_VOUCHER LEFT JOIN CNc_ASIENTO_VOUCHER ON CNC_ASIENTO_VOUCHER.Emp_cCodigo = CND_ASIENTO_VOUCHER.Emp_cCodigo AND                                                    
CNC_ASIENTO_VOUCHER.Pan_cAnio = CND_ASIENTO_VOUCHER.Pan_cAnio AND CNC_ASIENTO_VOUCHER.Per_cPeriodo = CND_ASIENTO_VOUCHER.Per_cPeriodo AND                                                    
CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = CND_ASIENTO_VOUCHER.Lib_cTipoLibro AND CNC_ASIENTO_VOUCHER.Ase_nVoucher = CND_ASIENTO_VOUCHER.Ase_nVoucher                                                    
LEFT JOIN CNM_ENTIDAD ON CND_ASIENTO_VOUCHER.Emp_cCodigo = CNM_ENTIDAD.Emp_cCodigo AND                                                    
CND_ASIENTO_VOUCHER.Ten_cTipoEntidad = CNM_ENTIDAD.Ten_cTipoEntidad AND CND_ASIENTO_VOUCHER.Ent_cCodEntidad = CNM_ENTIDAD.Ent_cCodEntidad                                                    
LEFT  JOIN TABLA ON  CNM_ENTIDAD.Emp_cCodigo = TABLA.Emp_cCodigo AND CNM_ENTIDAD.Ent_cTipoDoc = TABLA.Tab_cCodigo AND                                                    
TABLA.tab_ctabla = ''003'' LEFT JOIN CNT_LIBRO_OPERA ON CND_ASIENTO_VOUCHER.Lib_cTipoLibro = CNT_LIBRO_OPERA.Lib_cTipoLibro AND                                                    
CND_ASIENTO_VOUCHER.Pan_cAnio = CNT_LIBRO_OPERA.Pan_cAnio AND CND_ASIENTO_VOUCHER.Emp_cCodigo = CNT_LIBRO_OPERA.Emp_cCodigo                                                    
LEFT JOIN CNT_ENTIDAD ON CNM_ENTIDAD.Ten_cTipoEntidad = CNT_ENTIDAD.Ten_cTipoEntidad AND                                                    
CNM_ENTIDAD.Emp_cCodigo = CNT_ENTIDAD.Emp_cCodigo                                              
WHERE   CND_ASIENTO_VOUCHER.Asd_cDeleted <> ''*'' and CNC_ASIENTO_VOUCHER.AsE_cDeleted <> ''*'' and                                                       
CND_ASIENTO_VOUCHER.Emp_cCodigo = ''' + @Emp_cCodigo + '''  AND  CNC_ASIENTO_VOUCHER.pan_canio = ''' + @Pan_cAnio + ''' and                                                                          
CND_ASIENTO_VOUCHER.Pla_cCuentaContable  >= ''' + @CtaDesde + ''' AND CND_ASIENTO_VOUCHER.Pla_cCuentaContable  <= ''' + @CtaHasta + '''                                                    
and CNC_ASIENTO_VOUCHER.Per_cPeriodo = ''' + @Per_cPeriodoDesde + ''' and (CNC_ASIENTO_VOUCHER.Asd_cEstadoD = ''8'')'                                                      
                                                    
set @sql = ' insert into ' + @tabla + @sql                                                     
exec (@sql)                    
                  
--Agrego los estados 9                    
                  
set @sql = ' SELECT CND_ASIENTO_VOUCHER.Pla_cCuentaContable,              
cnc_asiento_voucher.Ase_nVoucher as CUO,              
  case when left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) = ''0100'' then              
  LEFT(LTRIM(RTRIM(''A'' + left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + right(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,5) )),100)              
  when left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) = ''0813'' or left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) = ''0814'' then              
  LEFT(LTRIM(RTRIM(''C'' + left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + right(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,5) )),100)              
  else LEFT(LTRIM(RTRIM(''M'' + left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + right(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,5) )),100)              
  end As Correlativo,              
LEFT (CND_ASIENTO_VOUCHER.Pla_cCuentaContable, 3) AS ''D3'',                                                              
LEFT (CND_ASIENTO_VOUCHER.Pla_cCuentaContable, 2) AS ''D2'' ,LEFT (CND_ASIENTO_VOUCHER.Pla_cCuentaContable, 2) AS ''D2_SUMAS'','                                                              
+ convert(varchar(12),@nValor) + ' as ''SaldoAntMonNac'',' + convert(varchar(12),@nValor) + ' as ''SaldoAntMonExt'',                                            
CND_ASIENTO_VOUCHER.Asd_nDebeSoles,CND_ASIENTO_VOUCHER.Asd_nDebeMonExt, CND_ASIENTO_VOUCHER.Asd_nHaberSoles,CND_ASIENTO_VOUCHER.Asd_nHaberMonExt,                                                              
CND_ASIENTO_VOUCHER.Emp_cCodigo,CND_ASIENTO_VOUCHER.Pan_cAnio, CND_ASIENTO_VOUCHER.Per_cPeriodo,CND_ASIENTO_VOUCHER.Lib_cTipoLibro,                                                              
CNT_LIBRO_OPERA.Lib_cDescripcion,CND_ASIENTO_VOUCHER.Ase_nVoucher,CND_ASIENTO_VOUCHER.Asd_cTipoMoneda,CND_ASIENTO_VOUCHER.Asd_nItem,                                                              
CNc_ASIENTO_VOUCHER.Ase_cGlosa,CND_ASIENTO_VOUCHER.Asd_nTipoCambio,CND_ASIENTO_VOUCHER.Cos_cCodigo,CND_ASIENTO_VOUCHER.Ten_cTipoEntidad,                                                              
CNM_ENTIDAD.Ent_nRuc as Ent_cCodEntidad,CNM_ENTIDAD.Ent_cPersona,CNT_ENTIDAD.Ten_cNombreEntidad,CND_ASIENTO_VOUCHER.Asd_cTipoDoc,                                                              
CND_ASIENTO_VOUCHER.Asd_cSerieDoc,CND_ASIENTO_VOUCHER.Asd_cNumDoc,convert(varchar(10), (case when year(CNc_ASIENTO_VOUCHER.Ase_dFecha)=1900 then null else CNc_ASIENTO_VOUCHER.Ase_dFecha end) ,103) as Asd_dFecDoc,                        
CND_ASIENTO_VOUCHER.Asd_cTipoDocRef,CND_ASIENTO_VOUCHER.Asd_cSerieDocRef, CND_ASIENTO_VOUCHER.Asd_cNumDocRef,CND_ASIENTO_VOUCHER.Asd_dFecDocRef,                                                        
CND_ASIENTO_VOUCHER.Asd_nMontoInafecto,CND_ASIENTO_VOUCHER.Asd_cRetencion,CND_ASIENTO_VOUCHER.Asd_dFechaSpot,CND_ASIENTO_VOUCHER.Asd_cNumSpot,                     
CND_ASIENTO_VOUCHER.Asd_cDestino,CND_ASIENTO_VOUCHER.Asd_nCorre,TABLA.Tab_cCodSunat,''' + @Mon_cMNac + ''' as Mon_cMNac,                                                      
Asd_cEstadoO, Asd_cEstadoD, CND_ASIENTO_VOUCHER.Id_Aduana                                     
FROM CND_ASIENTO_VOUCHER LEFT JOIN CNc_ASIENTO_VOUCHER ON CNC_ASIENTO_VOUCHER.Emp_cCodigo = CND_ASIENTO_VOUCHER.Emp_cCodigo AND                                                    
CNC_ASIENTO_VOUCHER.Pan_cAnio = CND_ASIENTO_VOUCHER.Pan_cAnio AND CNC_ASIENTO_VOUCHER.Per_cPeriodo = CND_ASIENTO_VOUCHER.Per_cPeriodo AND                                                    
CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = CND_ASIENTO_VOUCHER.Lib_cTipoLibro AND CNC_ASIENTO_VOUCHER.Ase_nVoucher = CND_ASIENTO_VOUCHER.Ase_nVoucher                                                    
LEFT JOIN CNM_ENTIDAD ON CND_ASIENTO_VOUCHER.Emp_cCodigo = CNM_ENTIDAD.Emp_cCodigo AND                                                    
CND_ASIENTO_VOUCHER.Ten_cTipoEntidad = CNM_ENTIDAD.Ten_cTipoEntidad AND CND_ASIENTO_VOUCHER.Ent_cCodEntidad = CNM_ENTIDAD.Ent_cCodEntidad                                                    
LEFT  JOIN TABLA ON  CNM_ENTIDAD.Emp_cCodigo = TABLA.Emp_cCodigo AND CNM_ENTIDAD.Ent_cTipoDoc = TABLA.Tab_cCodigo AND                                                    
TABLA.tab_ctabla = ''003'' LEFT JOIN CNT_LIBRO_OPERA ON CND_ASIENTO_VOUCHER.Lib_cTipoLibro = CNT_LIBRO_OPERA.Lib_cTipoLibro AND                                                    
CND_ASIENTO_VOUCHER.Pan_cAnio = CNT_LIBRO_OPERA.Pan_cAnio AND CND_ASIENTO_VOUCHER.Emp_cCodigo = CNT_LIBRO_OPERA.Emp_cCodigo                                                    
LEFT JOIN CNT_ENTIDAD ON CNM_ENTIDAD.Ten_cTipoEntidad = CNT_ENTIDAD.Ten_cTipoEntidad AND                                                    
CNM_ENTIDAD.Emp_cCodigo = CNT_ENTIDAD.Emp_cCodigo                                                    
WHERE   CND_ASIENTO_VOUCHER.Asd_cDeleted <> ''*'' and CNC_ASIENTO_VOUCHER.AsE_cDeleted <> ''*'' and                                                       
CND_ASIENTO_VOUCHER.Emp_cCodigo = ''' + @Emp_cCodigo + '''  AND  year(CNC_ASIENTO_VOUCHER.Ase_dfechaModifica) = ''' + @Pan_cAnio + ''' and                                                                          
CND_ASIENTO_VOUCHER.Pla_cCuentaContable  >= ''' + @CtaDesde + ''' AND CND_ASIENTO_VOUCHER.Pla_cCuentaContable  <= ''' + @CtaHasta + '''                                                    
and month(CNC_ASIENTO_VOUCHER.Ase_dfechaModifica) = ''' + @Per_cPeriodoDesde + ''' and (CNC_ASIENTO_VOUCHER.Asd_cEstadoD = ''9'')'                                                      
                                                    
set @sql = 'insert into ' + @tabla + @sql                                                      
exec (@sql)                                                  
                                                      
-- MOVIMIENTOS DESE EL PERIODO DE APERTURA AL PERIODO INICIAL SOLO IMPORTES Y CUENTAS            
IF @Per_cPeriodoDesde > '00'                                                                          
BEGIN                                                
                                                    
set @sql = ' SELECT CND_ASIENTO_VOUCHER.Pla_cCuentaContable,              
cnc_asiento_voucher.Ase_nVoucher as CUO,              
  case when left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) = ''0100'' then              
  LEFT(LTRIM(RTRIM(''A'' + left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + right(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,5) )),100)              
  when left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) = ''0813'' or left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) = ''0814'' then              
  LEFT(LTRIM(RTRIM(''C'' + left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + right(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,5) )),100)              
  else LEFT(LTRIM(RTRIM(''M'' + left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + right(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,5) )),100)              
  end As Correlativo, LEFT (CND_ASIENTO_VOUCHER.Pla_cCuentaContable, 3) AS D3,                                            
    LEFT (CND_ASIENTO_VOUCHER.Pla_cCuentaContable, 2) AS D2,LEFT (CND_ASIENTO_VOUCHER.Pla_cCuentaContable, 2) AS D2_SUMAS,                                            
    --SUM(CND_ASIENTO_VOUCHER.Asd_nDebeSoles - CND_ASIENTO_VOUCHER.Asd_nHaberSoles) as SaldoAntMonNac, SUM(CND_ASIENTO_VOUCHER.Asd_nDebeMonExt - CND_ASIENTO_VOUCHER.Asd_nHaberMonExt) as SaldoAntMonExt,                   
    CND_ASIENTO_VOUCHER.Asd_nDebeSoles - CND_ASIENTO_VOUCHER.Asd_nHaberSoles as SaldoAntMonNac, CND_ASIENTO_VOUCHER.Asd_nDebeMonExt - CND_ASIENTO_VOUCHER.Asd_nHaberMonExt as SaldoAntMonExt,'                                                                    
    + convert(varchar(12),@nValor) + ' as Asd_nDebeSoles,' + convert(varchar(12),@nValor) + ' as Asd_nDebeMonExt,'                                            
    + convert(varchar(12),@nValor) + ' as Asd_nHaberSoles,' + convert(varchar(12),@nValor) + ' as Asd_nHaberMonExt,                                            
CND_ASIENTO_VOUCHER.Emp_cCodigo,    CND_ASIENTO_VOUCHER.Pan_cAnio, '''' AS Per_cPeriodo,   ''XX'' AS Lib_cTipoLibro,                                            
    '''' AS Lib_cDescripcion,    '''' AS Ase_nVoucher, '''' AS Asd_cTipoMoneda,' + convert(varchar(12),@nValor) +  'AS Asd_nItem,                                            
    ''SALDO INICIAL'' AS Ase_cGlosa,' + convert(varchar(12),@nValor) + ' AS Asd_nTipoCambio,                                            
    '''' AS Cos_cCodigo,   '''' AS Ten_cTipoEntidad, '''' AS Ent_cCodEntidad, '''' AS Ent_cPersona, '''' AS Ten_cNombreEntidad, '''' AS Asd_cTipoDoc, '''' AS Asd_cSerieDoc,  '''' AS Asd_cNumDoc,                                    
    '''' AS Asd_dFecDoc,  '''' AS Asd_cTipoDocRef,'''' AS Asd_cSerieDocRef,   '''' AS Asd_cNumDocRef, NULL AS Asd_dFecDocRef,' + convert(varchar(12),@nValor) + ' AS Asd_nMontoInafecto,                                    
    '''' AS Asd_cRetencion, NULL AS Asd_dFechaSpot, '''' AS Asd_cNumSpot,  '''' AS Asd_cDestino,' + convert(varchar(12),@nValor) + ' AS Asd_nCorre, '''' AS Tab_cCodSunat,''' + @Mon_cMNac + ''' as Mon_cMNac,                                    
    Asd_cEstadoO, Asd_cEstadoD, CND_ASIENTO_VOUCHER.Id_Aduana  FROM CND_ASIENTO_VOUCHER LEFT JOIN CNM_ENTIDAD ON CND_ASIENTO_VOUCHER.Emp_cCodigo = CNM_ENTIDAD.Emp_cCodigo AND            
    CND_ASIENTO_VOUCHER.Ten_cTipoEntidad = CNM_ENTIDAD.Ten_cTipoEntidad AND CND_ASIENTO_VOUCHER.Ent_cCodEntidad = CNM_ENTIDAD.Ent_cCodEntidad                                                             
 inner join CNC_ASIENTO_VOUCHER on CND_ASIENTO_VOUCHER.Ase_cNummov = CNC_ASIENTO_VOUCHER.Ase_cNummov and CND_ASIENTO_VOUCHER.Emp_cCodigo = CNC_ASIENTO_VOUCHER.Emp_cCodigo and                                                        
 CND_ASIENTO_VOUCHER.Pan_cAnio = CNC_ASIENTO_VOUCHER.Pan_cAnio and CND_ASIENTO_VOUCHER.Per_cPeriodo = CNC_ASIENTO_VOUCHER.Per_cPeriodo and                                                        
 CND_ASIENTO_VOUCHER.Lib_cTipoLibro = CNC_ASIENTO_VOUCHER.Lib_cTipoLibro and CND_ASIENTO_VOUCHER.Ase_nVoucher = CNC_ASIENTO_VOUCHER.Ase_nVoucher                                                        
 WHERE   CND_ASIENTO_VOUCHER.Asd_cDeleted <> ''*'' and                        
   CND_ASIENTO_VOUCHER.Emp_cCodigo = ''' + @Emp_cCodigo  + ''' and                                                               
   CND_ASIENTO_VOUCHER.Pla_cCuentaContable  >= ''' + @CtaDesde + ''' AND CND_ASIENTO_VOUCHER.Pla_cCuentaContable  <= ''' + @CtaHasta + '''                                                              
   AND  CND_ASIENTO_VOUCHER.Pan_cAnio = ''' + @Pan_cAnio + '''                                                          
   AND CND_ASIENTO_VOUCHER.Per_cPeriodo >= ''00'' and '                                                    
                                    
 if @Per_cPeriodoHasta = ''                                                              
 begin                                        
  set @sql = @sql + ' CNC_ASIENTO_VOUCHER.Ase_dFecha < ''' + @Per_Fechadesde + ''''                                                        
 end                   
 else                                                        
 begin                                                        
  set @sql = @sql + 'CND_ASIENTO_VOUCHER.Per_cPeriodo < ''' + @Per_cPeriodoDesde + ''''                                                        
 end                                                   
                                                         
 set @sql = @sql + ' GROUP BY CND_ASIENTO_VOUCHER.Pla_cCuentaContable,              
    CND_ASIENTO_VOUCHER.Emp_cCodigo,cnc_asiento_voucher.Ase_nVoucher,              
    CND_ASIENTO_VOUCHER.Pan_cAnio,Asd_cEstadoO, Asd_cEstadoD, CND_ASIENTO_VOUCHER.Id_Aduana, CND_ASIENTO_VOUCHER.Asd_nDebeSoles, CND_ASIENTO_VOUCHER.Asd_nHaberSoles, CND_ASIENTO_VOUCHER.Asd_nDebeMonExt, CND_ASIENTO_VOUCHER.Asd_nHaberMonExt '              
                                                                
 set @sql = 'insert into ' + @tabla + @sql                                                               
 exec (@sql)                                                          
                                                    
--Agrego los estados 8                                                    
set @sql = ' SELECT CND_ASIENTO_VOUCHER.Pla_cCuentaContable,              
cnc_asiento_voucher.Ase_nVoucher as CUO,              
  case when left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) = ''0100'' then              
  LEFT(LTRIM(RTRIM(''A'' + left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + right(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,5) )),100)              
  when left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) = ''0813'' or left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) = ''0814'' then              
  LEFT(LTRIM(RTRIM(''C'' + left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + right(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,5) )),100)              
  else LEFT(LTRIM(RTRIM(''M'' + left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + right(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,5) )),100)              
  end As Correlativo, LEFT (CND_ASIENTO_VOUCHER.Pla_cCuentaContable, 3) AS D3,                                            
    LEFT (CND_ASIENTO_VOUCHER.Pla_cCuentaContable, 2) AS D2,LEFT (CND_ASIENTO_VOUCHER.Pla_cCuentaContable, 2) AS D2_SUMAS,                                            
--    SUM(CND_ASIENTO_VOUCHER.Asd_nDebeSoles - CND_ASIENTO_VOUCHER.Asd_nHaberSoles) as SaldoAntMonNac, SUM(CND_ASIENTO_VOUCHER.Asd_nDebeMonExt - CND_ASIENTO_VOUCHER.Asd_nHaberMonExt) as SaldoAntMonExt,
      CND_ASIENTO_VOUCHER.Asd_nDebeSoles - CND_ASIENTO_VOUCHER.Asd_nHaberSoles as SaldoAntMonNac, CND_ASIENTO_VOUCHER.Asd_nDebeMonExt - CND_ASIENTO_VOUCHER.Asd_nHaberMonExt as SaldoAntMonExt,'                                            
    + convert(varchar(12),@nValor) + ' as Asd_nDebeSoles,' + convert(varchar(12),@nValor) + ' as Asd_nDebeMonExt,'                                            
    + convert(varchar(12),@nValor) + ' as Asd_nHaberSoles,' + convert(varchar(12),@nValor) + ' as Asd_nHaberMonExt,                                            
    CND_ASIENTO_VOUCHER.Emp_cCodigo,    CND_ASIENTO_VOUCHER.Pan_cAnio, '''' AS Per_cPeriodo,   ''XX'' AS Lib_cTipoLibro,                               
    '''' AS Lib_cDescripcion,    '''' AS Ase_nVoucher, '''' AS Asd_cTipoMoneda,' + convert(varchar(12),@nValor) +  'AS Asd_nItem,                                            
    ''SALDO INICIAL'' AS Ase_cGlosa,'  + convert(varchar(12),@nValor) + ' AS Asd_nTipoCambio,                                            
    '''' AS Cos_cCodigo,   '''' AS Ten_cTipoEntidad, '''' AS Ent_cCodEntidad, '''' AS Ent_cPersona, '''' AS Ten_cNombreEntidad, '''' AS Asd_cTipoDoc, '''' AS Asd_cSerieDoc,  '''' AS Asd_cNumDoc,                                    
    '''' AS Asd_dFecDoc,  '''' AS Asd_cTipoDocRef,'''' AS Asd_cSerieDocRef,   '''' AS Asd_cNumDocRef, NULL AS Asd_dFecDocRef,' + convert(varchar(12),@nValor) + ' AS Asd_nMontoInafecto,                                    
    '''' AS Asd_cRetencion, NULL AS Asd_dFechaSpot, '''' AS Asd_cNumSpot,  '''' AS Asd_cDestino,'                                    
    + convert(varchar(12),@nValor) + ' AS Asd_nCorre, '''' AS Tab_cCodSunat,''' + @Mon_cMNac + ''' as Mon_cMNac,                        
    Asd_cEstadoO, Asd_cEstadoD, CND_ASIENTO_VOUCHER.Id_Aduana  FROM CND_ASIENTO_VOUCHER LEFT JOIN CNM_ENTIDAD ON CND_ASIENTO_VOUCHER.Emp_cCodigo = CNM_ENTIDAD.Emp_cCodigo AND                                    
    CND_ASIENTO_VOUCHER.Ten_cTipoEntidad = CNM_ENTIDAD.Ten_cTipoEntidad AND CND_ASIENTO_VOUCHER.Ent_cCodEntidad = CNM_ENTIDAD.Ent_cCodEntidad                                                            
 inner join CNC_ASIENTO_VOUCHER on CND_ASIENTO_VOUCHER.Ase_cNummov = CNC_ASIENTO_VOUCHER.Ase_cNummov and CND_ASIENTO_VOUCHER.Emp_cCodigo = CNC_ASIENTO_VOUCHER.Emp_cCodigo and                                                        
 CND_ASIENTO_VOUCHER.Pan_cAnio = CNC_ASIENTO_VOUCHER.Pan_cAnio and CND_ASIENTO_VOUCHER.Per_cPeriodo = CNC_ASIENTO_VOUCHER.Per_cPeriodo and                                                        
 CND_ASIENTO_VOUCHER.Lib_cTipoLibro = CNC_ASIENTO_VOUCHER.Lib_cTipoLibro and CND_ASIENTO_VOUCHER.Ase_nVoucher = CNC_ASIENTO_VOUCHER.Ase_nVoucher                                                        
 WHERE   CND_ASIENTO_VOUCHER.Asd_cDeleted <> ''*'' and                                                               
   CND_ASIENTO_VOUCHER.Emp_cCodigo = ''' + @Emp_cCodigo  + ''' and                                                               
   CND_ASIENTO_VOUCHER.Pla_cCuentaContable  >= ''' + @CtaDesde + ''' AND CND_ASIENTO_VOUCHER.Pla_cCuentaContable  <= ''' + @CtaHasta + '''                                                              
   AND  CNC_ASIENTO_VOUCHER.pan_canio = ''' + @Pan_cAnio + '''                                                          
   and CNC_ASIENTO_VOUCHER.Per_cPeriodo = ''' + @Per_cPeriodoDesde + ''' and (CNC_ASIENTO_VOUCHER.Asd_cEstadoD = ''8'')'                             
                                                    
 set @sql = @sql + ' GROUP BY CND_ASIENTO_VOUCHER.Pla_cCuentaContable,              
    CND_ASIENTO_VOUCHER.Emp_cCodigo,cnc_asiento_voucher.Ase_nVoucher,              
    CND_ASIENTO_VOUCHER.Pan_cAnio,Asd_cEstadoO, Asd_cEstadoD, CND_ASIENTO_VOUCHER.Id_Aduana, CND_ASIENTO_VOUCHER.Asd_nDebeSoles, CND_ASIENTO_VOUCHER.Asd_nHaberSoles, CND_ASIENTO_VOUCHER.Asd_nDebeMonExt, CND_ASIENTO_VOUCHER.Asd_nHaberMonExt '              
                                                        
 set @sql = 'insert into ' + @tabla + @sql                                                               
 exec (@sql)                     
                   
 --Agrego los estados 9                  
                   
 set @sql = ' SELECT CND_ASIENTO_VOUCHER.Pla_cCuentaContable,              
 cnc_asiento_voucher.Ase_nVoucher as CUO,              
  case when left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) = ''0100'' then              
  LEFT(LTRIM(RTRIM(''A'' + left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + right(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,5) )),100)              
  when left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) = ''0813'' or left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) = ''0814'' then              
  LEFT(LTRIM(RTRIM(''C'' + left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + right(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,5) )),100)              
  else LEFT(LTRIM(RTRIM(''M'' + left(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,4) + right(dbo.CNC_ASIENTO_VOUCHER.Ase_nVoucher,5) )),100)              
  end As Correlativo, LEFT (CND_ASIENTO_VOUCHER.Pla_cCuentaContable, 3) AS D3,                                            
    LEFT (CND_ASIENTO_VOUCHER.Pla_cCuentaContable, 2) AS D2,LEFT (CND_ASIENTO_VOUCHER.Pla_cCuentaContable, 2) AS D2_SUMAS,                                            
--    SUM(CND_ASIENTO_VOUCHER.Asd_nDebeSoles - CND_ASIENTO_VOUCHER.Asd_nHaberSoles) as SaldoAntMonNac, SUM(CND_ASIENTO_VOUCHER.Asd_nDebeMonExt - CND_ASIENTO_VOUCHER.Asd_nHaberMonExt) as SaldoAntMonExt,
 CND_ASIENTO_VOUCHER.Asd_nDebeSoles - CND_ASIENTO_VOUCHER.Asd_nHaberSoles as SaldoAntMonNac, CND_ASIENTO_VOUCHER.Asd_nDebeMonExt - CND_ASIENTO_VOUCHER.Asd_nHaberMonExt as SaldoAntMonExt,'                                       
    + convert(varchar(12),@nValor) + ' as Asd_nDebeSoles,' + convert(varchar(12),@nValor) + ' as Asd_nDebeMonExt,'                                            
    + convert(varchar(12),@nValor) + ' as Asd_nHaberSoles,' + convert(varchar(12),@nValor) + ' as Asd_nHaberMonExt,                                            
    CND_ASIENTO_VOUCHER.Emp_cCodigo,    CND_ASIENTO_VOUCHER.Pan_cAnio, '''' AS Per_cPeriodo,   ''XX'' AS Lib_cTipoLibro,                                            
    '''' AS Lib_cDescripcion,    '''' AS Ase_nVoucher, '''' AS Asd_cTipoMoneda,' + convert(varchar(12),@nValor) +  'AS Asd_nItem,                                            
    ''SALDO INICIAL'' AS Ase_cGlosa,'  + convert(varchar(12),@nValor) + ' AS Asd_nTipoCambio,                                            
    '''' AS Cos_cCodigo,   '''' AS Ten_cTipoEntidad, '''' AS Ent_cCodEntidad, '''' AS Ent_cPersona, '''' AS Ten_cNombreEntidad, '''' AS Asd_cTipoDoc, '''' AS Asd_cSerieDoc,  '''' AS Asd_cNumDoc,                                    
    '''' AS Asd_dFecDoc,  '''' AS Asd_cTipoDocRef,'''' AS Asd_cSerieDocRef,   '''' AS Asd_cNumDocRef, NULL AS Asd_dFecDocRef,' + convert(varchar(12),@nValor) + ' AS Asd_nMontoInafecto,                                    
    '''' AS Asd_cRetencion, NULL AS Asd_dFechaSpot, '''' AS Asd_cNumSpot,  '''' AS Asd_cDestino,'                                    
    + convert(varchar(12),@nValor) + ' AS Asd_nCorre, '''' AS Tab_cCodSunat,''' + @Mon_cMNac + ''' as Mon_cMNac,                                    
    Asd_cEstadoO, Asd_cEstadoD, CND_ASIENTO_VOUCHER.Id_Aduana  FROM CND_ASIENTO_VOUCHER LEFT JOIN CNM_ENTIDAD ON CND_ASIENTO_VOUCHER.Emp_cCodigo = CNM_ENTIDAD.Emp_cCodigo AND                                    
    CND_ASIENTO_VOUCHER.Ten_cTipoEntidad = CNM_ENTIDAD.Ten_cTipoEntidad AND CND_ASIENTO_VOUCHER.Ent_cCodEntidad = CNM_ENTIDAD.Ent_cCodEntidad                                                            
 inner join CNC_ASIENTO_VOUCHER on CND_ASIENTO_VOUCHER.Ase_cNummov = CNC_ASIENTO_VOUCHER.Ase_cNummov and CND_ASIENTO_VOUCHER.Emp_cCodigo = CNC_ASIENTO_VOUCHER.Emp_cCodigo and                                                        
 CND_ASIENTO_VOUCHER.Pan_cAnio = CNC_ASIENTO_VOUCHER.Pan_cAnio and CND_ASIENTO_VOUCHER.Per_cPeriodo = CNC_ASIENTO_VOUCHER.Per_cPeriodo and                                                        
 CND_ASIENTO_VOUCHER.Lib_cTipoLibro = CNC_ASIENTO_VOUCHER.Lib_cTipoLibro and CND_ASIENTO_VOUCHER.Ase_nVoucher = CNC_ASIENTO_VOUCHER.Ase_nVoucher                             
WHERE   CND_ASIENTO_VOUCHER.Asd_cDeleted <> ''*'' and                  
   CND_ASIENTO_VOUCHER.Emp_cCodigo = ''' + @Emp_cCodigo  + ''' and                  
   CND_ASIENTO_VOUCHER.Pla_cCuentaContable  >= ''' + @CtaDesde + ''' AND CND_ASIENTO_VOUCHER.Pla_cCuentaContable  <= ''' + @CtaHasta + '''                                                  
   AND  year(CNC_ASIENTO_VOUCHER.Ase_dfechaModifica) = ''' + @Pan_cAnio + '''                                                          
   and month(CNC_ASIENTO_VOUCHER.Ase_dfechaModifica) = ''' + @Per_cPeriodoDesde + ''' and (CNC_ASIENTO_VOUCHER.Asd_cEstadoD = ''9'')'                                                      
                  
 set @sql = @sql + ' GROUP BY CND_ASIENTO_VOUCHER.Pla_cCuentaContable,              
    CND_ASIENTO_VOUCHER.Emp_cCodigo,cnc_asiento_voucher.Ase_nVoucher,              
    CND_ASIENTO_VOUCHER.Pan_cAnio,Asd_cEstadoO, Asd_cEstadoD, CND_ASIENTO_VOUCHER.Id_Aduana, CND_ASIENTO_VOUCHER.Asd_nDebeSoles, CND_ASIENTO_VOUCHER.Asd_nHaberSoles, CND_ASIENTO_VOUCHER.Asd_nDebeMonExt, CND_ASIENTO_VOUCHER.Asd_nHaberMonExt '              
                          
 set @sql = 'insert into ' + @tabla + @sql                                                               
 exec (@sql)                     
                                                     
                                                        
END                                                                        
                                                        
/*                                                              
--------------------------------------------------------------------------                                                                          
-- exec spCn_RptFormato0601 'TODOS', '014', '2006', '06', '07', '038', '10', '99'                                                               
-- exec spCn_RptFormato0601 'TODOS', '014', '2006', '00', '07', '038', '10', '99'                                                                          
*/                                                              
                                                              
set @sql= 'insert into ' + @tabla2 + ' SELECT TMP.Pla_cCuentaContable,TMP.CUO, TMP.Correlativo, TMP.D3 , TMP.D2 , TMP.D2_sumas ,                                                            
 TMP.SaldoAntMonNac , TMP.SaldoAntMonExt , TMP.Asd_nDebeSoles , TMP.Asd_nDebeMonExt ,                                                          
 TMP.Asd_nHaberSoles , TMP.Asd_nHaberMonExt ,TMP.Emp_cCodigo , TMP.Pan_cAnio ,                                                             
 TMP.Per_cPeriodo , TMP.Lib_cTipoLibro ,  TMP.Lib_cDescripcion , TMP.Ase_nVoucher ,                                                              
 TMP.Asd_cTipoMoneda , TMP.Asd_nItem , TMP.Ase_cGlosa , TMP.Asd_nTipoCambio ,                                               
 TMP.Cos_cCodigo , TMP.Ten_cTipoEntidad , TMP.Ent_cCodEntidad , TMP.Ent_cPersona ,                                                              
 TMP.Ten_cNombreEntidad , TMP.Asd_cTipoDoc ,  TMP.Asd_cSerieDoc , TMP.Asd_cNumDoc ,                                                              
 convert(varchar(10), (case when year(TMP.Asd_dFecDoc)=1900 then null else TMP.Asd_dFecDoc end) ,103) as Asd_dFecDoc,                                                            
 TMP.Asd_cTipoDocRef , TMP.Asd_cSerieDocRef , TMP.Asd_cNumDocRef ,                                                              
 convert(varchar(10),TMP.Asd_dFecDocRef,103) , TMP.Asd_nMontoInafecto ,TMP.Asd_cRetencion , TMP.Asd_dFechaSpot ,                                                              
 TMP.Asd_cNumSpot , TMP.Asd_cDestino , TMP.Asd_nCorre , TMP.Tab_cCodSunat ,                                                              
 TMP.Mon_cMNac , CTADET.Pla_cNombreCuenta, CTAD3.Pla_cNombreCuenta AS Pla_cNombreCuentaD3,                                                                 
CTAD2.Pla_cNombreCuenta AS Pla_cNombreCuentaD2 ,TMP.Asd_cEstadoO, TMP.Asd_cEstadoD, TMP.Id_Aduana                                                      
FROM ' + @tabla + ' as TMP                                                      
LEFT JOIN CNM_PLAN_CTA CTADET ON  TMP.Pla_cCuentaContable = CTADET.Pla_cCuentaContable AND                                                              
 TMP.Pan_cAnio = CTADET.Pan_cAnio AND TMP.Emp_cCodigo = CTADET.Emp_cCodigo                                                              
LEFT JOIN CNM_PLAN_CTA CTAD2 ON  TMP.D2 = CTAD2.Pla_cCuentaContable AND                                                              
 TMP.Pan_cAnio = CTAD2.Pan_cAnio AND TMP.Emp_cCodigo = CTAD2.Emp_cCodigo                                                              
LEFT JOIN CNM_PLAN_CTA CTAD3 ON TMP.D3 = CTAD3.Pla_cCuentaContable AND                                                              
 TMP.Pan_cAnio = CTAD3.Pan_cAnio AND TMP.Emp_cCodigo = CTAD3.Emp_cCodigo                                                       
ORDER BY  TMP.Pla_cCuentaContable, TMP.Per_cPeriodo, TMP.Ase_nVoucher'                                                      
exec(@sql)                                                      
                                          
set @sql = 'delete from ' + @tabla2 + ' where SaldoAntMonNac > ''0'''                              
exec(@sql)                                                      
                                                      
set @sql = 'delete from ' + @tabla2 + ' where Asd_cEstadoO = ''2'' or Asd_cEstadoD = ''2'' or Ase_nVoucher = '''''                                                      
exec(@sql)                                                      
                         

--SET @SQL =  'SELECT * FROM ' + @tabla2                           
--EXEC (@SQL)
                        
--                        return 
                            
set @sql = 'create table TMPDiarioPLEDH              
( Per_cPeriodo varchar(8),              
 CUO varchar(15),              
 Correlativo varchar(10),              
Pla_cCuentaContable varchar(12),              
 Ase_dFecha datetime,              
 Ase_cGlosa varchar(100),              
 Asd_nDebeSoles decimal(14,2),              
 Asd_nHaberSoles decimal(14,2),              
 Estado char(1),
 Asd_cTipoMoneda char(3),
 Asd_cTipoDocEmisor char(1),
 Asd_cNumEmisor varchar(20),
 Asd_cTipoDoc char(2),
 Asd_cSerieDoc varchar(20),
 Asd_cNumDoc varchar(25),
 Ase_cGlosaRef varchar(100),
 Ase_cTipoLibro varchar(100),
 Pan_cAnio char(4),
 Per_cPeriodoRef char(2),
 Ase_nVoucher char(10),
 Id_Aduana char(10), 
 Corr INT IDENTITY(1,1)               
 )'              
                         
exec(@sql)                                           
         
             /*Modificar Aqui*/                                       
set @sql = 'insert into TMPDiarioPLEDH select LTRIM(RTRIM(LEFT(Pan_cAnio +  case Per_cPeriodo when ''00'' then ''01'' when ''13'' then ''12'' when ''14'' then ''12'' else Per_cPeriodo end + ''00'',8))) as ''Per_cPeriodo'',                        
LTRIM(RTRIM(LEFT(Cuo,40))) as ''CUO'',              
LTRIM(RTRIM(LEFT(Correlativo,40))) as ''Correlativo'',                          
LTRIM(RTRIM(LEFT(Pla_cCuentaContable,24))) as ''Pla_cCuentaContable'',                                                    
Asd_dFecDoc as ''Ase_dFecha'',                                        
LTRIM(RTRIM(LEFT(replace(replace(replace(Ase_cGlosa, char(13), '' ''), char(10), '' ''), char(9), '' ''),100))) as ''Ase_cGlosa'',                                               
CONVERT(money,Asd_nDebeSoles) as ''Asd_nDebeSoles'',                                    
CONVERT(money,Asd_nHaberSoles) as ''Asd_nHaberSoles'',                                    
case when Asd_cEstadoD = '''' then                                              
case when Asd_cEstadoO in (''0'',''1'',''6'',''7'') then ''1''end                     
when Asd_cEstadoD = ''9'' and Per_cPeriodo = ''' + @Per_cPeriodo + ''' and Pan_cAnio= ''' +  @Pan_cAnio + ''' then ''1'' else Asd_cEstadoD end as ''Estado'',
CASE WHEN Asd_cTipoMoneda = ''038'' THEN ''PEN'' ELSE ''USD'' END AS Asd_cTipoMoneda, ''6'' AS Asd_cTipoDocEmisor, ''' + @Ruc + ''' AS Asd_cNumEmisor, Asd_cTipoDoc,
Asd_cSerieDoc, Asd_cNumDoc, '''' AS Ase_cGlosaRef, CASE WHEN Lib_cTipoLibro = ''05'' THEN ''140100&'' + convert(varchar(4),year(Asd_dFecDoc)) + Per_cPeriodo + ''&'' + LTRIM(RTRIM(LEFT(Ase_nVoucher,40))) + ''&'' + LEFT(LTRIM(RTRIM(''M'' + left(Ase_nVoucher,4) + right(Ase_nVoucher,5) )),100) 
			   WHEN Lib_cTipoLibro = ''06'' THEN ''080200&'' + CAST(YEAR(Asd_dFecDoc) AS CHAR(4)) + Per_cPeriodo + ''00&'' +  LEFT(LTRIM(RTRIM(Ase_nVoucher)),40) + ''&'' +
			   LEFT(LTRIM(RTRIM(''M'' + left(Ase_nVoucher,4) + right(Ase_nVoucher,5) )),100) ELSE '''' END AS Ase_cTipoLibro, Pan_cAnio, Per_cPeriodo AS Per_cPeriodoRef, Ase_nVoucher, Id_Aduana                                        
 from ' + @tabla2 + ''                                                     
exec(@sql)    
   
                                            
set @sql = 'Declare @Per_cPeriodoDH varchar(8)        
Declare @CUO varchar(15)         
Declare @Correlativo varchar(10)        
Declare @Pla_cCuentaContable varchar(12)        
Declare @Ase_dFecha datetime         
Declare @Ase_cGlosa varchar(50)         
Declare @Asd_nDebeSoles decimal(14,2)         
Declare @Asd_nHaberSoles decimal(14,2)         
Declare @Estado char(1)        
Declare @Nro int    
DECLARE @Asd_cTipoMoneda char(3)
DECLARE @Asd_cTipoDocEmisor char(1)
DECLARE @Asd_cNumEmisor varchar(20)
DECLARE @Asd_cTipoDoc char(2)
DECLARE @Asd_cSerieDoc varchar(20)
DECLARE @Asd_cNumDoc varchar(25)
DECLARE @Ase_cGlosaRef varchar(100)
DECLARE @Ase_cTipoLibro varchar(100)
DECLARE @Pan_cAnio char(4)
DECLARE @Per_cPeriodoRef char(2)
DECLARE @Ase_nVoucher char(10)  
DECLARE @Id_Aduana char(10)   
        
set @Nro = ''1''        
                                              
DECLARE D_H_Cursor_TC CURSOR FOR      
SELECT Per_cPeriodo, CUO, Correlativo, Pla_cCuentaContable ,Ase_dFecha,Ase_cGlosa,/*sum(Asd_nDebeSoles)*/ Asd_nDebeSoles, /*sum(Asd_nHaberSoles)*/ Asd_nHaberSoles,Estado,
Asd_cTipoMoneda, Asd_cTipoDocEmisor, Asd_cNumEmisor, Asd_cTipoDoc, Asd_cSerieDoc, Asd_cNumDoc, Ase_cGlosaRef, ISNULL(Ase_cTipoLibro, ''''), Pan_cAnio, Per_cPeriodoRef, Ase_nVoucher, Id_Aduana                                              
FROM TMPDiarioPLEDH                                  
group by Per_cPeriodo, CUO, Correlativo, Pla_cCuentaContable ,Ase_dFecha,Ase_cGlosa,Estado,
Asd_cTipoMoneda, Asd_cTipoDocEmisor, Asd_cNumEmisor, Asd_cTipoDoc, Asd_cSerieDoc, Asd_cNumDoc, Ase_cGlosaRef, Ase_cTipoLibro, Pan_cAnio, Per_cPeriodoRef, Ase_nVoucher, Id_Aduana, Asd_nDebeSoles, Asd_nHaberSoles                             
having sum(Asd_nDebeSoles) > 0 and sum(Asd_nHaberSoles) > 0                                             
                                              
OPEN D_H_Cursor_TC                                                
FETCH NEXT FROM D_H_Cursor_TC                                               
INTO @Per_cPeriodoDH, @CUO,@Correlativo, @Pla_cCuentaContable ,@Ase_dFecha,@Ase_cGlosa,@Asd_nDebeSoles,@Asd_nHaberSoles,@Estado,
@Asd_cTipoMoneda, @Asd_cTipoDocEmisor, @Asd_cNumEmisor, @Asd_cTipoDoc, @Asd_cSerieDoc, @Asd_cNumDoc, @Ase_cGlosaRef, @Ase_cTipoLibro, @Pan_cAnio, @Per_cPeriodoRef, @Ase_nVoucher, @Id_Aduana                                               
                                                                            
 WHILE @@FETCH_STATUS = 0                                                
 BEGIN                                  
                                               
 insert into TMPDiarioPLEDH                                              
 select distinct @Per_cPeriodoDH as ''Per_cPeriodo'', CUO, LEFT(@Correlativo,5) + ''R'' + replicate(''0'',4 - len(@Nro)) + convert(varchar(10),@Nro)  as ''Correlativo'',                                               
 @Pla_cCuentaContable as ''Pla_cCuentaContable'' ,@Ase_dFecha as ''Ase_dFecha'',        
@Ase_cGlosa as ''Ase_cGlosa'',0 as ''Asd_nDebeSoles'',@Asd_nHaberSoles,@Estado,
@Asd_cTipoMoneda, @Asd_cTipoDocEmisor, @Asd_cNumEmisor, @Asd_cTipoDoc, @Asd_cSerieDoc, @Asd_cNumDoc, @Ase_cGlosaRef, @Ase_cTipoLibro, @Pan_cAnio, @Per_cPeriodoRef, @Ase_nVoucher, @Id_Aduana from TMPDiarioPLEDH         
 where Per_cPeriodo = @Per_cPeriodoDH and Cuo = @Cuo and Pla_cCuentaContable = @Pla_cCuentaContable         
        
print @Per_cPeriodoDH         
print @Cuo        
print @Pla_cCuentaContable        
                                                
 update TMPDiarioPLEDH                                              
 set Asd_nHaberSoles = 0                                              
 where Per_cPeriodo = @Per_cPeriodoDH and Cuo = @Cuo and Pla_cCuentaContable = @Pla_cCuentaContable and Correlativo = @Correlativo                                                
 set @Nro = @Nro + 1                                              
                                  
   FETCH NEXT FROM D_H_Cursor_TC                                                                                   
 INTO  @Per_cPeriodoDH, @CUO,@Correlativo, @Pla_cCuentaContable ,@Ase_dFecha,@Ase_cGlosa,@Asd_nDebeSoles,@Asd_nHaberSoles,@Estado,
 @Asd_cTipoMoneda, @Asd_cTipoDocEmisor, @Asd_cNumEmisor, @Asd_cTipoDoc, @Asd_cSerieDoc, @Asd_cNumDoc, @Ase_cGlosaRef, @Ase_cTipoLibro, @Pan_cAnio, @Per_cPeriodoRef, @Ase_nVoucher, @Id_Aduana                                              
                                         
END                                              
CLOSE D_H_Cursor_TC                                              
DEALLOCATE D_H_Cursor_TC'                                              
                                              
exec(@sql)           
        
set @sql = 'delete from TMPDiarioPLEDH where Asd_nDebeSoles = ''0'' and Asd_nHaberSoles = ''0'''                                                  
exec(@sql)                        
                                        
exec(@sql)       

--SET @SQL = 'DECLARE @CUO VARCHAR(15)
--			DECLARE @Asd_cNumDoc VARCHAR(25)
--			DECLARE @Correlativo VARCHAR(10)
--			DECLARE @Asd_cTipoDoc CHAR(2)
--			DECLARE @Asd_cSerieDoc VARCHAR(20)
			
--            DECLARE C_MAYOR CURSOR FOR
--			SELECT RTRIM(LTRIM(CUO)) AS CUO, RTRIM(LTRIM(Correlativo)) AS Correlativo, RTRIM(LTRIM(Asd_cSerieDoc)), RTRIM(LTRIM(Asd_cNumDoc)), RTRIM(LTRIM(Asd_cTipoDoc)) FROM TMPDiarioPLEDH
--			WHERE RTRIM(LTRIM(Asd_cSerieDoc)) <> '''' AND RTRIM(LTRIM(Asd_cNumDoc)) <> ''''
			
--			OPEN C_MAYOR
--			FETCH NEXT FROM C_MAYOR INTO @CUO, @Correlativo, @Asd_cSerieDoc, @Asd_cNumDoc, @Asd_cTipoDoc
--			WHILE @@FETCH_STATUS = 0
--			BEGIN
			
--				UPDATE TMPDiarioPLEDH
--					SET Asd_cNumDoc = @Asd_cNumDoc, Asd_cSerieDoc = @Asd_cSerieDoc, Asd_cTipoDoc = @Asd_cTipoDoc 
--				WHERE RTRIM(LTRIM(CUO)) = RTRIM(LTRIM(@CUO)) AND RTRIM(LTRIM(Correlativo)) = RTRIM(LTRIM(@CORRELATIVO)) AND RTRIM(LTRIM(Asd_cSerieDoc)) = '''' AND RTRIM(LTRIM(Asd_cNumDoc)) = ''''
				
--				FETCH NEXT FROM C_MAYOR INTO @CUO, @Correlativo, @Asd_cSerieDoc, @Asd_cNumDoc, @Asd_cTipoDoc
--			END
--			CLOSE C_MAYOR
--			DEALLOCATE C_MAYOR'                              
--EXEC(@SQL)


                                            
--set @sql = 'select (LTRIM(RTRIM(Per_cPeriodo)) + ''' + @separador + '''+                                                      
--LTRIM(RTRIM(LEFT(CUO,40))) + ''' + @separador + '''+                 
--LTRIM(RTRIM(LEFT(Correlativo,10))) + ''' + @separador + '''+               
--LTRIM(RTRIM(LEFT(Pla_cCuentaContable,24))) + ''' + @separador + '''+          
--'''' + ''' + @separador + '''+                                                      
--Asd_cTipoMoneda + ''' + @separador + '''+                                                      
--''6''+ ''' + @separador + '''+                                              
--RTRIM(LTRIM(Asd_cNumEmisor)  + ''' + @separador + '''+                       
--Asd_cTipoDoc  + ''' + @separador + '''+           
--Asd_cSerieDoc + ''' + @separador + '''+           
--RIGHT(Asd_cNumDoc, 7) + ''' + @separador + '''+           
--CONVERT(NCHAR(10), Ase_dFecha, 103) + ''' + @separador + '''+ 
--'''' + ''' + @separador + '''+ 
--CONVERT(NCHAR(10), Ase_dFecha, 103) + ''' + @separador + '''+
--Ase_cGlosa + ''' + @separador + '''+ 
--'''' + ''' + @separador + '''+ 
--CAST(CAST(Asd_nDebeSoles AS NUMERIC(20, 2)) AS VARCHAR(50)) + ''' + @separador + '''+  
--CAST(CAST(Asd_nHaberSoles AS NUMERIC(20, 2)) AS VARCHAR(50)) + ''' + @separador + '''+ 
--Ase_cTipoLibro + ''' + @separador + '''+                                                
--Estado + ''' + @separador + ''') as Registro                                                      
-- from TMPDiarioPLEDH                                                      
-- group by Per_cPeriodo, Cuo, Correlativo, Pla_cCuentaContable, Ase_dFecha, Ase_cGlosa, Asd_cTipoMoneda, Estado, Ase_cTipoLibro, Asd_cTipoDoc, Asd_cSerieDoc , Asd_cNumDoc                                                     
--order by Cuo, Pla_cCuentaContable, Per_cPeriodo, Asd_cTipoMoneda, , Estado, Ase_cTipoLibro, Asd_cTipoDoc, Asd_cSerieDoc , Asd_cNumDoc  '                                                   
                                       
--exec(@sql)                                                     
             
       
DECLARE @Sepa CHAR(1)
SET @Sepa = '|'

--SELECT * FROM TMPDiarioPLEDH
SELECT (LTRIM(RTRIM(Per_cPeriodo)) + @Sepa + LTRIM(RTRIM(LEFT(CUO,40))) + @Sepa + 
        --LTRIM(RTRIM(LEFT(Correlativo,10))) 
        LEFT(Correlativo, 5) + RIGHT(CAST(10000 + Corr AS VARCHAR(20)) , 5)
        + @Sepa + LEFT(Pla_cCuentaContable, 24) + @Sepa + '' + @Sepa + '' + @Sepa +  Asd_cTipoMoneda + @Sepa + Asd_cTipoDocEmisor + @Sepa + RTRIM(LTRIM(Asd_cNumEmisor)) + @Sepa + RTRIM(LTRIM(Asd_cTipoDoc)) + @Sepa + 
        CASE WHEN RTRIM(LTRIM(Asd_cTipoDoc)) IN ('50', '52') THEN RIGHT(RTRIM(LTRIM(Asd_cSerieDoc)), 3) 
             WHEN RTRIM(LTRIM(Asd_cTipoDoc)) IN ('05') THEN RIGHT(RTRIM(LTRIM(Asd_cSerieDoc)), 1)
        ELSE RTRIM(LTRIM(Asd_cSerieDoc)) END + @Sepa + 
        
    --    CASE WHEN RTRIM(LTRIM(Asd_cTipoDoc)) IN ('50', '52') THEN RIGHT(RTRIM(LTRIM(CAST(Asd_cNumDoc AS VARCHAR(25)))), 6) 
			 --WHEN LEFT(Ase_nVoucher, 2) = '05' AND RTRIM(LTRIM(Asd_cTipoDoc)) IN ('01', '03', '04', '07', '08') THEN  RIGHT(RTRIM(LTRIM(Asd_cNumDoc)), 7)
			 --WHEN LEFT(Ase_nVoucher, 2) = '05' AND RTRIM(LTRIM(Asd_cTipoDoc)) IN ('00', '13', '14', '15', '16', '17', '18') THEN RIGHT(RTRIM(LTRIM(Asd_cNumDoc)), 20)
    --    ELSE RIGHT(RTRIM(LTRIM(CAST(Asd_cNumDoc AS VARCHAR(25)))), 7) END + @Sepa + 
    
        CASE WHEN /*Lib_cTipoLibro = '05' AND*/ RTRIM(LTRIM(Asd_cTipoDoc)) IN ('01', '04', '07', '08') THEN CASE WHEN (CHARINDEX('-', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('-', Asd_cNumDoc)), 7) 
																												 WHEN (CHARINDEX('/', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('/', Asd_cNumDoc)), 7) 
																											ELSE RIGHT(RTRIM(LTRIM(Asd_cNumDoc)), 7) END
                 WHEN /*Lib_cTipoLibro = '05' AND*/ RTRIM(LTRIM(Asd_cTipoDoc)) IN ('00', '10', '12', '13', '14', '15', '16', '17', '18', '19', '21', '22', '24', '26', '27', '28', '29') THEN CASE WHEN (CHARINDEX('-', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('-', Asd_cNumDoc)), 20) 
																																																   WHEN (CHARINDEX('/', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('/', Asd_cNumDoc)), 20) 
																																															  ELSE RIGHT(RTRIM(LTRIM(Asd_cNumDoc)), 20) END
                 WHEN /*Lib_cTipoLibro = '06' AND*/ RTRIM(LTRIM(Asd_cTipoDoc)) IN ('50', '52') THEN CASE WHEN (CHARINDEX('-', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('-', Asd_cNumDoc)), 6) 
																										 WHEN (CHARINDEX('/', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('/', Asd_cNumDoc)), 6) 
																									ELSE RIGHT(RTRIM(LTRIM(Asd_cNumDoc)), 6) END 
				 --WHEN RTRIM(LTRIM(Asd_cTipoDoc)) IN ('50', '52') THEN RIGHT(RTRIM(LTRIM(Asd_cNumDoc)), 6)
				 WHEN RTRIM(LTRIM(Asd_cTipoDoc)) IN ('03', '12') THEN CASE WHEN (CHARINDEX('-', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('-', Asd_cNumDoc)), 7) 
																		   WHEN (CHARINDEX('/', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('/', Asd_cNumDoc)), 7) 
																	  ELSE RIGHT(RTRIM(LTRIM(Asd_cNumDoc)), 7) END
				 WHEN RTRIM(LTRIM(Asd_cTipoDoc)) IN ('05') THEN CASE WHEN (CHARINDEX('-', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('-', Asd_cNumDoc)), 11) 
																	 WHEN (CHARINDEX('/', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('/', Asd_cNumDoc)), 11) 
																ELSE RIGHT(RTRIM(LTRIM(Asd_cNumDoc)), 11) END
                 WHEN RTRIM(LTRIM(Asd_cTipoDoc)) IN ('11') THEN CASE WHEN (CHARINDEX('-', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('-', Asd_cNumDoc)), 15) 
																	 WHEN (CHARINDEX('/', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, 0, CHARINDEX('/', Asd_cNumDoc)), 15) 
																ELSE RIGHT(RTRIM(LTRIM(Asd_cNumDoc)), 15) END
                 ELSE RIGHT(RTRIM(LTRIM(Asd_cNumDoc)), 7) END + @Sepa +
        
        
        CONVERT(NCHAR(10), Ase_dFecha, 103) + @Sepa + '' + @Sepa + CONVERT(NCHAR(10), Ase_dFecha, 103) + @Sepa + Ase_cGlosa + @Sepa + '' + @Sepa + CAST(CAST(Asd_nDebeSoles AS NUMERIC(20, 2)) AS VARCHAR(50)) + @Sepa + CAST(CAST(Asd_nHaberSoles AS NUMERIC(20, 2)) AS VARCHAR(50)) + @Sepa + Ase_cTipoLibro + @Sepa + Estado + @Sepa + 
        
        CASE WHEN /*Lib_cTipoLibro = '05' AND*/ RTRIM(LTRIM(Asd_cTipoDoc)) IN ('01', '04', '07', '08') THEN CASE WHEN (CHARINDEX('-', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, CHARINDEX('-', Asd_cNumDoc) + 1, LEN(Asd_cNumDoc) - CHARINDEX('-', Asd_cNumDoc)), 7) + '|' 
																												 WHEN (CHARINDEX('/', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, CHARINDEX('/', Asd_cNumDoc) + 1, LEN(Asd_cNumDoc) - CHARINDEX('/', Asd_cNumDoc)), 7) + '|' 
																											ELSE '|' END
                 WHEN /*Lib_cTipoLibro = '05' AND*/ RTRIM(LTRIM(Asd_cTipoDoc)) IN ('00', '10', '12', '13', '14', '15', '16', '17', '18', '19', '21', '22', '24', '26', '27', '28', '29') THEN CASE WHEN (CHARINDEX('-', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, CHARINDEX('-', Asd_cNumDoc) + 1, LEN(Asd_cNumDoc) - CHARINDEX('-', Asd_cNumDoc)), 20) + '|' 
																																																   WHEN (CHARINDEX('/', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, CHARINDEX('/', Asd_cNumDoc) + 1, LEN(Asd_cNumDoc) - CHARINDEX('/', Asd_cNumDoc)), 20) + '|'  
																																															  ELSE '|' END
                 WHEN /*Lib_cTipoLibro = '06' AND*/ RTRIM(LTRIM(Asd_cTipoDoc)) IN ('50', '52') THEN CASE WHEN (CHARINDEX('-', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, CHARINDEX('-', Asd_cNumDoc) + 1, LEN(Asd_cNumDoc) - CHARINDEX('-', Asd_cNumDoc)), 6) + '|'  
																										 WHEN (CHARINDEX('/', Asd_cNumDoc) > 0) THEN RIGHT(SUBSTRING(Asd_cNumDoc, CHARINDEX('/', Asd_cNumDoc) + 1, LEN(Asd_cNumDoc) - CHARINDEX('/', Asd_cNumDoc)), 6) + '|' 
																									ELSE '|' END 
				 --WHEN RTRIM(LTRIM(Asd_cTipoDoc)) IN ('50', '52') THEN RIGHT(RTRIM(LTRIM(Asd_cNumDoc)), 6)
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
        
        ) AS Registro FROM TMPDiarioPLEDH
     
     
set @sql = 'drop table ' + @tabla                                        
exec(@sql)                                        
set @sql = 'drop table TMPDiarioPLEDH'                                        
exec(@sql)                                        
set @sql = 'drop table ' + @tabla2                                        
exec(@sql)  
GO
