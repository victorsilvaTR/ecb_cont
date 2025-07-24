USE SAFC_ECB
GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO
--spCn_ConsultaProvisiones 'SEL_PEND', '002', '2010','', '01/01/2010', '31/01/2010', 'MIX','',''                  
/*
Modificado por: Pool Berrospi
Fecha de modificación: 19/12/2014
Proyecto: Pasaje Historico
*/
ALTER PROCEDURE [dbo].[spCn_ReprocesoImportacionXLSv2](                  
@Accion char(20)='',                  
@Emp_cCodigo char(3)='',                  
@Pan_cAnio char(4)='',                  
@Par_cEnt char(1)='0',                  
@Par_cTca char(1)='0',                  
@Par_cApe char(1)='0',                  
@Par_cCom char(1)='0',                  
@Par_cVen char(1)='0',                  
@Par_cCin char(1)='0',                  
@Par_cCeg char(1)='0',                  
@Par_cPlan char(1)='0',                  
@User varchar(20)='',
@Periodo char(2)=''  
                  
)                   
--WITH ENCRYPTION                  
AS                   
SET NOCOUNT ON                  
SET DATEFORMAT DMY                  
                  
DECLARE @cTab_cCodProc CHAR(1)                  
SET @cTab_cCodProc  = case  when @Par_cEnt='1' then '1'            
        when @Par_cTca='1' then '2'            
        when @Par_cApe='1' then '3'            
        when @Par_cCom='1' then '4'            
        when @Par_cVen='1' then '5'            
        when @Par_cCin='1' then '6'            
        when @Par_cCeg='1' then '7'            
        when @Par_cPlan='1' then '8'            
      end            
            
IF @Accion ='REPORTE'            
BEGIN            
 SELECT * FROM CNT_REPORTE_IMPORTACION            
 WHERE  Emp_cCodigo = @Emp_cCodigo            
 ORDER BY 1,2,3            
END            
                  
IF @Accion ='ELIMINACION' OR @Accion ='IMPORTACION'                  
BEGIN                  
  IF NOT exists (select * from .sysobjects where id = object_id(N'CNT_REPORTE_IMPORTACION') and OBJECTPROPERTY(id, N'IsUserTable') = 1)                  
  BEGIN                  
  CREATE TABLE CNT_REPORTE_IMPORTACION (                  
   Emp_cCodigo char (3) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,                  
   Tab_cCodProc char (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,                  
   Tab_cProceso nvarchar (250) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,                  
   Tab_cTabla nvarchar (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,                  
   Tab_cError nvarchar (1500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,                  
   Tab_cMensaje nvarchar (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,                  
   Tab_cUserCrea nvarchar (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,                  
   Tab_dFechaCrea datetime NULL ,                  
   Tab_cEquipoUser nvarchar (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL                   
                    
  )                   
  END                   
                  
END                  
                  
IF @Accion ='ELIMINACION'                  
BEGIN                  
  DELETE FROM CNT_REPORTE_IMPORTACION WHERE Emp_cCodigo = @Emp_cCodigo                   
END                  
                  
IF @Accion ='IMPORTACION'                  
BEGIN                  
  ----------------------------------------------------------------------                  
  DECLARE @cErr varchar(500)                  
  DECLARE @nRegistros int                  
  ----------------------------------------------------------------------                  
  -- IMPORTACION DE ENTIDADES                  
  ----------------------------------------------------------------------                  
  IF @Par_cEnt ='1'                  
  BEGIN                  
   -- 1. CREAMOS LA TABLA TEMPORAL ENTIDADES                  
   SELECT TOP 0 * INTO #TMP_ENTIDADES FROM CNM_ENTIDAD WITH(NOLOCK) WHERE Emp_cCodigo = @Emp_cCodigo                   
            
   -- 2. LLENAMOS LA TABLA TEMPORAL            
   INSERT INTO #TMP_ENTIDADES            
    ( Emp_cCodigo, Ent_cCodEntidad, Ten_cTipoEntidad, Ent_cPersona, Ent_cDireccion, Ent_nRuc, Ent_cRepresentante,            
    Ent_cTipoDoc, Ent_cFlagPersona, Ent_cEstadoEntidad, Ent_cEstado, Ent_cDeleted, Ent_cUserCrea, Ent_dFechaCrea, Ent_cEquipoUser, Ent_cFlagDomiciliado,
    Id_Pais, Id_Vinculo_Economico, Id_Convenio, PorcentajeSunat)            
   SELECT @Emp_cCodigo, Ent_cCodEntidad, Ten_cTipoEntidad, Ent_cPersona, isnull(Ent_cDireccion,''), isnull(Ent_nRuc,''), '',            
    isnull(Ent_cTipoDoc,''), isnull(Ent_cFlagPersona,''), 'A', 'A','', @User, GETDATE(), host_name(), Ent_cFlagDomiciliado, Id_Pais, Id_Vinculo_Economico, Id_Convenio, PorcentajeSunat            
   FROM ZIMP_ENTIDAD WITH(NOLOCK)            
   IF @@ERROR <> 0            
   BEGIN            
    INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK) (Emp_cCodigo, Tab_cCodProc, Tab_cProceso, Tab_cTabla, Tab_cError, Tab_cMensaje, Tab_cUserCrea, Tab_dFechaCrea, Tab_cEquipoUser )                  
    VALUES(@Emp_cCodigo,@cTab_cCodProc, 'ENTIDADES', 'CNM_ENTIDAD', 'Error al insertar los valores en la tabla maestra' ,'verifique los campos obligatorios', @User,GETDATE(), host_name() )                  
    GOTO tagFinEnt                  
   END                  
                    
   -- 3. VALIDAMOS LA TABLA ENTIDADES                  
                   
   --VALIDA LA DUPLICIDAD DE ENTIDADES Y RUC EN EL TEMPORAL                  
   INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)                   
   SELECT @Emp_cCodigo,@cTab_cCodProc, 'ENTIDADES', 'CNM_ENTIDAD', 'El documento ' + Ent_cTipoDoc + '-' + Ent_nRuc + ', esta repetido.','', @User,GETDATE(), host_name()                   
   FROM #TMP_ENTIDADES                   
   GROUP BY Ten_cTipoEntidad, Ent_cTipoDoc, Ent_nRuc                  
   HAVING count(Ten_cTipoEntidad + Ent_cTipoDoc + Ent_nRuc) > 1                  
                    
   --VALIDA LA EXISTENCIA DEL TIPO ENTIDAD EN LA TABLA DE TIPO DE ENTIDADES                  
   INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)                   
   SELECT @Emp_cCodigo,@cTab_cCodProc, 'ENTIDADES', 'CNM_ENTIDAD', 'El Tipo de Entidad ' + Ten_cTipoEntidad + ', no se encuentra el Maestro de Tipos de Entidad','', @User,GETDATE(), host_name()                   
   FROM #TMP_ENTIDADES                  
   WHERE Ten_cTipoEntidad not in (SELECT Ten_cTipoEntidad FROM CNT_ENTIDAD WITH(NOLOCK) WHERE Emp_cCodigo = @Emp_cCodigo )                  
                    
   --VALIDA LA EXISTENCIA DEL CODIGO EN LA TABLA DE ENTIDADES                  
   INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)                   
   SELECT @Emp_cCodigo,@cTab_cCodProc, 'ENTIDADES', 'CNM_ENTIDAD', 'El Codigo de entidad ' + Ten_cTipoEntidad + '-' + Ent_cCodEntidad + ', ya existe en el Maestro de Entidades.','', @User,GETDATE(), host_name()                   
   FROM #TMP_ENTIDADES                  
   WHERE Ten_cTipoEntidad + Ent_cCodEntidad IN (SELECT Ten_cTipoEntidad + Ent_cCodEntidad FROM CNM_ENTIDAD WITH(NOLOCK) WHERE Emp_cCodigo = @Emp_cCodigo )                  
                    
   --VALIDA LA EXISTENCIA DEL DOCUMENTO Y TIPO DE ENTIDAD EN EL MAESTRO                  
   INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)                   
   SELECT @Emp_cCodigo,@cTab_cCodProc, 'ENTIDADES', 'CNM_ENTIDAD', 'El documento '+ Ent_cTipoDoc + '-' + Ent_nRuc + ' de la entidad ' + Ten_cTipoEntidad + '-' + Ent_cCodEntidad + ' ya esta existe en el Maestro de Entidades','', @User,GETDATE(), host_name() 
   FROM #TMP_ENTIDADES                  
   WHERE Ten_cTipoEntidad + Ent_cTipoDoc + Ent_nRuc IN (SELECT Ten_cTipoEntidad + Ent_cTipoDoc + Ent_nRuc FROM CNM_ENTIDAD WITH(NOLOCK) WHERE Emp_cCodigo = @Emp_cCodigo )                  
                    
   -- 4. INSERTA LOS DATOS VALIDADOS EN LA TABLA ENTIDADES                  
   IF NOT exists (SELECT * FROM CNT_REPORTE_IMPORTACION  WHERE Emp_cCodigo= @Emp_cCodigo AND Tab_cCodProc= '1')                  
   BEGIN                  
    INSERT INTO CNM_ENTIDAD WITH(ROWLOCK) SELECT * FROM #TMP_ENTIDADES                   
   END                  
                    
                  
                  
  tagFinEnt:                  
                    
  END                  
                    
                   
  ----------------------------------------------------------------------                  
  -- IMPORTACION DE TIPO DE CAMBIO                  
  ----------------------------------------------------------------------                  
  IF @Par_cTca = '1'                  
  BEGIN                  
   -- 1. CREAMOS LA TABLA TEMPORAL DE TIPO DE CAMBIO       
   SELECT TOP 0 * INTO #TMP_TIPO_CAMBIO FROM CNT_TIPO_CAMBIO WITH(NOLOCK) WHERE Emp_cCodigo = @Emp_cCodigo                   
                     
   -- 2. LLENAMOS LA TABLA TEMPORAL DE TIPO DE CAMBIO                  
   INSERT INTO #TMP_TIPO_CAMBIO                   
    ( Emp_cCodigo, Tca_dFecha, Tca_cCodigoOrigen, Tca_cCodigoDestino, Tca_nCompra, Tca_nVenta,                   
    Tca_nCompraP, Tca_nVentaP, Tca_cEstado, Tca_cDeleted, Tca_cUserCrea, Tca_dFechaCrea, Tca_cEquipoUser )                  
   SELECT @Emp_cCodigo, CONVERT(datetime, Tca_dFecha), isnull(Tca_cCodigoOrigen,''), isnull(Tca_cCodigoDestino,''), isnull(Tca_nCompra,0), isnull(Tca_nVenta,0),                   
    0, isnull(Tca_nVentaP,0), 'A', '', @User, GETDATE(), host_name()                  
   FROM ZIMP_TIPOCAMBIO WITH(NOLOCK)                  
   IF @@ERROR <> 0                   
   BEGIN                  
    INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK) (Emp_cCodigo, Tab_cCodProc, Tab_cProceso, Tab_cTabla, Tab_cError, Tab_cMensaje, Tab_cUserCrea, Tab_dFechaCrea, Tab_cEquipoUser )                  
    VALUES(@Emp_cCodigo,@cTab_cCodProc, 'TIPO DE CAMBIO', 'CNT_TIPO_CAMBIO', 'Error al insertar los valores en la tabla maestra' ,'verifique los campos obligatorios', @User,GETDATE(), host_name() )                  
    GOTO tagFinTC                  
   END                  
                    
   -- 3. VALIDAMOS LA TABLA DE TIPO DE CAMBIO                 
   --VALIDA LA DUPLICIDAD DE FECHAS EN EL TEMPORAL                  
   INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)                   
   SELECT @Emp_cCodigo,@cTab_cCodProc, 'TIPO DE CAMBIO', 'CNT_TIPO_CAMBIO', 'La fecha ' + CONVERT(char(10), Tca_dFecha, 103) + ', esta repetida.','', @User,GETDATE(), host_name()                   
   FROM #TMP_TIPO_CAMBIO                   
   GROUP BY Tca_dFecha                  
   HAVING count(Tca_dFecha) > 1                  
                    
   --VALIDA LA EXISTENCIA LA MONEDA ORIGEN EN LA TABLA DE TIPO DE MONEDAS                  
   INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)                   
   SELECT @Emp_cCodigo,@cTab_cCodProc, 'TIPO DE CAMBIO', 'CNT_TIPO_CAMBIO', 'La Moneda de origen ' + Tca_cCodigoOrigen + ', no existe en el Maestro de Monedas.','', @User,GETDATE(), host_name()                   
   FROM #TMP_TIPO_CAMBIO                  
   WHERE Tca_cCodigoOrigen NOT IN (SELECT Mon_cCodigo FROM CNT_TIPO_MONEDA WITH(NOLOCK) WHERE Emp_cCodigo = @Emp_cCodigo )                  
                    
   --VALIDA LA EXISTENCIA LA MONEDA DESTINO EN LA TABLA DE TIPO DE MONEDAS                  
   INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)                   
   SELECT @Emp_cCodigo,@cTab_cCodProc, 'TIPO DE CAMBIO', 'CNT_TIPO_CAMBIO', 'La Moneda de destino ' + Tca_cCodigoDestino + ', no existe en el Maestro de Monedas.','', @User,GETDATE(), host_name()                   
   FROM #TMP_TIPO_CAMBIO                  
   WHERE Tca_cCodigoDestino NOT IN (SELECT Mon_cCodigo FROM CNT_TIPO_MONEDA WITH(NOLOCK) WHERE Emp_cCodigo = @Emp_cCodigo )                  
                    
   --VALIDA LA EXISTENCIA LA FECHA EN LA TABLA DE TIPO DE CAMBIO                  
   INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)                   
   SELECT @Emp_cCodigo,@cTab_cCodProc, 'TIPO DE CAMBIO', 'CNT_TIPO_CAMBIO', 'La fecha ' + CONVERT(char(10), Tca_dFecha, 103) + ', para convertir de ' + Tca_cCodigoOrigen + ' a ' + Tca_cCodigoDestino + 
   ' ya existe en el Maestro de Tipos de Cambio.','', @User,GETDATE(), host_name()                   
   FROM #TMP_TIPO_CAMBIO                  
   WHERE Tca_cCodigoOrigen + Tca_cCodigoDestino + convert(varchar(8),Tca_dFecha ,112) IN                  
    (SELECT Tca_cCodigoOrigen + Tca_cCodigoDestino + convert(varchar(8),Tca_dFecha ,112)                   
    FROM CNT_TIPO_CAMBIO WITH(NOLOCK) WHERE Emp_cCodigo = @Emp_cCodigo )                  
                    
   -- 4. INSERTA LOS DATOS VALIDADOS EN LA TABLA DE TIPO DE CAMBIO                  
   IF NOT exists (SELECT * FROM CNT_REPORTE_IMPORTACION  WHERE Emp_cCodigo= @Emp_cCodigo AND Tab_cCodProc= '2')                  
   BEGIN                  
    INSERT INTO CNT_TIPO_CAMBIO WITH(ROWLOCK)  SELECT * FROM #TMP_TIPO_CAMBIO                   
   END                  
                    
  tagFinTC:                  
  END                  
        
  ----------------------------------------------------------------------                  
  -- IMPORTACION DE MOVIMIENTOS        
  ----------------------------------------------------------------------                  
                    
  -- SALIR SI NO ESTAN MARCADOS                  
  IF @Par_cApe = '0' AND @Par_cCom = '0' AND @Par_cVen = '0' AND @Par_cCin = '0' AND @Par_cCeg = '0' and @Par_cPlan ='0' GOTO tagExit                  
                    
  -- VALIDAR QUE SOLO SE ENVIE UNO DE ESTOS PARAMENTROS A LA VEZ                  
  IF CONVERT(INT, @Par_cApe) + CONVERT(INT, @Par_cCom) + CONVERT(INT, @Par_cVen) + CONVERT(INT, @Par_cCin) + CONVERT(INT, @Par_cCeg)  + CONVERT(INT, @Par_cPlan) > 1                  
  BEGIN                  
   INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK) (Emp_cCodigo, Tab_cCodProc, Tab_cProceso, Tab_cTabla, Tab_cError, Tab_cMensaje, Tab_cUserCrea, Tab_dFechaCrea, Tab_cEquipoUser )                  
   VALUES(@Emp_cCodigo,@cTab_cCodProc, 'CABECERA VOUCHER', 'CNC_ASIENTO_VOUCHER', 'Solo se puede importar un libro a la vez.' ,'verifique los campos obligatorios', @User,GETDATE(), host_name() )                  
   GOTO tagFinMov                  
  END                  
                    
  IF @Par_cApe = '1' OR @Par_cCom = '1' OR @Par_cVen = '1' OR @Par_cCin = '1' OR @Par_cCeg = '1' OR @Par_cPlan = '1'                  
  BEGIN                  
                  
   -- 1. CREAMOS LA TABLA TEMPORAL DE CABECERA DE VOUCHER                  
   SELECT TOP 0 * INTO #TMP_C_ASIENTO_VOUCHER FROM CNC_ASIENTO_VOUCHER WITH(NOLOCK) WHERE Emp_cCodigo = @Emp_cCodigo                   
                     
   -- 2. LLENAMOS LA TABLA TEMPORAL DE CABECERA DE VOUCHER                  
   IF @Par_cApe = '1'              
   BEGIN              
    INSERT INTO #TMP_C_ASIENTO_VOUCHER    
     (Ase_cNummov, Emp_cCodigo, Pan_cAnio, Per_cPeriodo, Lib_cTipoLibro, Ase_nVoucher, Ase_dFecha, Ase_cGlosa, Ase_cTipoMoneda                  
     , Ase_nTipoCambio, Ase_cOperaTC, Ase_cOperaCaja, Ase_cEstado, Ase_cDeleted, Ase_cUserCrea, Ase_dFechaCrea, Ase_cEquipoUser                  
     , Ase_cCuadreManual, Ase_cCodSoft,Asd_cEstadoO ,Asd_cEstadoD, CreditoFiscal, MaterialConstruccion)                  
    SELECT Ase_cNummov, @Emp_cCodigo, Pan_cAnio, Per_cPeriodo, Lib_cTipoLibro, Ase_nVoucher, Ase_dFecha, Ase_cGlosa, Ase_cTipoMoneda                  
     , 0, '', '', 'A', '', @User, GETDATE(), host_name()                  
     , '0', 'IMP','1','', CreditoFiscal, MaterialConstruccion                 
    FROM ZIMP_APE_CAB WITH(NOLOCK)                  
   END                  
                  
   IF @Par_cPlan = '1'                  
   BEGIN              
    INSERT INTO #TMP_C_ASIENTO_VOUCHER              
     (Ase_cNummov, Emp_cCodigo, Pan_cAnio, Per_cPeriodo, Lib_cTipoLibro, Ase_nVoucher, Ase_dFecha, Ase_cGlosa, Ase_cTipoMoneda                  
     , Ase_nTipoCambio, Ase_cOperaTC, Ase_cOperaCaja, Ase_cEstado, Ase_cDeleted, Ase_cUserCrea, Ase_dFechaCrea, Ase_cEquipoUser                  
     , Ase_cCuadreManual, Ase_cCodSoft,Asd_cEstadoO ,Asd_cEstadoD, CreditoFiscal, MaterialConstruccion)                  
    SELECT Ase_cNummov, @Emp_cCodigo, Pan_cAnio, Per_cPeriodo, Lib_cTipoLibro, Ase_nVoucher, Ase_dFecha, Ase_cGlosa, Ase_cTipoMoneda                  
     , 0, '', '', 'A', '', @User, GETDATE(), host_name()                  
     , '0', 'IMP','1','', CreditoFiscal, MaterialConstruccion                 
    FROM ZIMP_PLAN_CAB WITH(NOLOCK)                  
  END            
                     
   IF @Par_cCom = '1'                  
   BEGIN                  
    INSERT INTO #TMP_C_ASIENTO_VOUCHER                  
     (Ase_cNummov, Emp_cCodigo, Pan_cAnio, Per_cPeriodo, Lib_cTipoLibro, Ase_nVoucher, Ase_dFecha, Ase_cGlosa, Ase_cTipoMoneda                  
     , Ase_nTipoCambio, Ase_cOperaTC, Ase_cOperaCaja, Ase_cEstado, Ase_cDeleted, Ase_cUserCrea, Ase_dFechaCrea, Ase_cEquipoUser                  
     , Ase_cCuadreManual, Ase_cCodSoft,Asd_cEstadoO ,Asd_cEstadoD, CreditoFiscal, MaterialConstruccion)                  
    SELECT Ase_cNummov, @Emp_cCodigo, Pan_cAnio, Per_cPeriodo, Lib_cTipoLibro, Ase_nVoucher, Ase_dFecha, Ase_cGlosa, Ase_cTipoMoneda                  
     , 0, '', '', 'A', '', @User, GETDATE(), host_name()                  
     , '0', 'IMP','1','', CreditoFiscal, MaterialConstruccion                  
    FROM ZIMP_COMPRAS_CAB WITH(NOLOCK)                  
   END         
            
   IF @Par_cVen = '1'                  
   BEGIN               
           
    INSERT INTO #TMP_C_ASIENTO_VOUCHER                  
     (Ase_cNummov, Emp_cCodigo, Pan_cAnio, Per_cPeriodo, Lib_cTipoLibro, Ase_nVoucher, Ase_dFecha, Ase_cGlosa, Ase_cTipoMoneda                  
     , Ase_nTipoCambio, Ase_cOperaTC, Ase_cOperaCaja, Ase_cEstado, Ase_cDeleted, Ase_cUserCrea, Ase_dFechaCrea, Ase_cEquipoUser                  
     , Ase_cCuadreManual, Ase_cCodSoft,Asd_cEstadoO ,Asd_cEstadoD, CreditoFiscal, MaterialConstruccion )                  
    SELECT Ase_cNummov, @Emp_cCodigo, Pan_cAnio, Per_cPeriodo, Lib_cTipoLibro, Ase_nVoucher, Ase_dFecha, Ase_cGlosa, Ase_cTipoMoneda                  
     , 0, '', '', 'A', '', @User, GETDATE(), host_name()                  
     , '0', 'IMP',Ase_cEstadoO,Ase_cEstadoD , CreditoFiscal, MaterialConstruccion               
    FROM ZIMP_VENTAS_CAB WITH(NOLOCK)                  
        
   END                 
            
   IF @Par_cCin = '1'                  
   BEGIN                  
    INSERT INTO #TMP_C_ASIENTO_VOUCHER                  
     (Ase_cNummov, Emp_cCodigo, Pan_cAnio, Per_cPeriodo, Lib_cTipoLibro, Ase_nVoucher, Ase_dFecha, Ase_cGlosa, Ase_cTipoMoneda                  
     , Ase_nTipoCambio, Ase_cOperaTC, Ase_cOperaCaja, Ase_cEstado, Ase_cDeleted, Ase_cUserCrea, Ase_dFechaCrea, Ase_cEquipoUser                  
     , Ase_cCuadreManual, Ase_cCodSoft,Asd_cEstadoO ,Asd_cEstadoD, CreditoFiscal, MaterialConstruccion)                  
    SELECT Ase_cNummov, @Emp_cCodigo, Pan_cAnio, Per_cPeriodo, Lib_cTipoLibro, Ase_nVoucher, Ase_dFecha, Ase_cGlosa, Ase_cTipoMoneda                  
     , 0, '', 'I', 'A', '', @User, GETDATE(), host_name()                  
     , '0', 'IMP','1','', CreditoFiscal, MaterialConstruccion                  
    FROM ZIMP_CAJAING_CAB WITH(NOLOCK)                  
   END                  
   IF @Par_cCeg = '1'                  
   BEGIN                  
    INSERT INTO #TMP_C_ASIENTO_VOUCHER                  
     (Ase_cNummov, Emp_cCodigo, Pan_cAnio, Per_cPeriodo, Lib_cTipoLibro, Ase_nVoucher, Ase_dFecha, Ase_cGlosa, Ase_cTipoMoneda                  
     , Ase_nTipoCambio, Ase_cOperaTC, Ase_cOperaCaja, Ase_cEstado, Ase_cDeleted, Ase_cUserCrea, Ase_dFechaCrea, Ase_cEquipoUser                  
     , Ase_cCuadreManual, Ase_cCodSoft,Asd_cEstadoO ,Asd_cEstadoD, CreditoFiscal, MaterialConstruccion)                  
    SELECT Ase_cNummov, @Emp_cCodigo, Pan_cAnio, Per_cPeriodo, Lib_cTipoLibro, Ase_nVoucher, Ase_dFecha, Ase_cGlosa, Ase_cTipoMoneda                  
     , 0, '', 'E', 'A', '', @User, GETDATE(), host_name()                  
     , '0', 'IMP','1','', CreditoFiscal, MaterialConstruccion                
    FROM ZIMP_CAJAEGR_CAB WITH(NOLOCK)                  
   END                  
           
   IF @@ERROR <> 0                   
   BEGIN                  
    INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK) (Emp_cCodigo, Tab_cCodProc, Tab_cProceso, Tab_cTabla, Tab_cError, Tab_cMensaje, Tab_cUserCrea, Tab_dFechaCrea, Tab_cEquipoUser )                  
    VALUES(@Emp_cCodigo,@cTab_cCodProc, 'CABECERA VOUCHER', 'CNC_ASIENTO_VOUCHER', 'Error al insertar los valores en la tabla CNC_ASIENTO_VOUCHER' ,'verifique los campos obligatorios', @User,GETDATE(), host_name() )                  
    GOTO tagFinMov                  
   END              
                  
   -- 3. VALIDAMOS LA TABLA DE CABECERA DE VOUCHER                  
                    
   -- VALIDAR ANCHO DEL NUMERO DE MOVIMIENTO                  
   INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)
   SELECT @Emp_cCodigo,@cTab_cCodProc, 'CABECERA VOUCHER', 'CNC_ASIENTO_VOUCHER', 'El Numero de Movimiento ' + Ase_cNummov + ' debe tener un ancho de 10 caracteres y deben ser Digitos.','', @User,GETDATE(), host_name()                   
   FROM #TMP_C_ASIENTO_VOUCHER                  
   WHERE CONVERT(INT, Ase_cNummov) = 0 OR LEN(RTRIM(LTRIM(Ase_cNummov))) <> 10                  
        
   -- VALIDAR ANCHO DEL NUMERO DE VOUCHER                  
   INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)                   
   SELECT @Emp_cCodigo,@cTab_cCodProc, 'CABECERA VOUCHER', 'CNC_ASIENTO_VOUCHER', 'El Numero de Voucher ' + Ase_nVoucher + ' debe tener un ancho de 10 caracteres y deben ser Digitos.','', @User,GETDATE(), host_name()                   
   FROM #TMP_C_ASIENTO_VOUCHER                  
   WHERE CONVERT(INT, Ase_nVoucher) = 0 OR LEN(RTRIM(LTRIM(Ase_nVoucher))) <> 10                  
                    
   -- VALIDAR PERIODO                  
   IF @Par_cApe = '1'                   
   BEGIN                  
    INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)                   
    SELECT @Emp_cCodigo,@cTab_cCodProc, 'CABECERA VOUCHER', 'CNC_ASIENTO_VOUCHER', 'El Periodo debe ser 00 para el Movimiento ' + Ase_cNummov + ' con Voucher ' + Ase_nVoucher + ' del Libro ' + Lib_cTipoLibro + '.','', @User,GETDATE(), host_name()        
    FROM #TMP_C_ASIENTO_VOUCHER                  
    WHERE Per_cPeriodo <> '00'                  
   END                  
   ELSE                  
   BEGIN                  
    INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)                   
    SELECT @Emp_cCodigo,@cTab_cCodProc, 'CABECERA VOUCHER', 'CNC_ASIENTO_VOUCHER', 'El Periodo NO debe ser 00 para el Movimiento ' + Ase_cNummov + ' con Voucher ' + Ase_nVoucher + ' del Libro ' + Lib_cTipoLibro + '.','', @User,GETDATE(), host_name()      
    FROM #TMP_C_ASIENTO_VOUCHER                  
    WHERE Per_cPeriodo = '00'                  
   END        
                    
   declare @cLibro   varchar(20)                  
                  
   SELECT @cLibro  = CASE @Par_cApe WHEN '1' THEN Cfl_cApertura ELSE '' END                  
     + CASE @Par_cCom WHEN '1' THEN Cfl_cCompras ELSE '' END                  
     + CASE @Par_cVen WHEN '1' THEN Cfl_cVentas ELSE '' END                  
     + CASE @Par_cCin WHEN '1' THEN Cfl_cCajaIngresos ELSE '' END                  
     + CASE @Par_cCeg WHEN '1' THEN Cfl_cCajaEgresos ELSE '' END                  
     + CASE @Par_cPlan WHEN '1' THEN Cfl_cDiario  ELSE '' END                  
    FROM CNT_CONFIG_LIBROS WITH(NOLOCK) WHERE Emp_cCodigo = @Emp_cCodigo                  
                  
   -- VALIDAR QUE PERTENECA AL LIBRO SELECCIONADO                  
   INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)                   
   SELECT @Emp_cCodigo,@cTab_cCodProc, 'CABECERA VOUCHER', 'CNC_ASIENTO_VOUCHER', 'El Libro ' + Lib_cTipoLibro +  ' del Movimiento ' + Ase_cNummov + ' para el año ' + Pan_cAnio + ', no es el Libro configurado.','', @User,GETDATE(), host_name()            
   FROM #TMP_C_ASIENTO_VOUCHER                  
   WHERE Lib_cTipoLibro <> dbo.trimsql(@cLibro  )                  
                  
                    
   --VALIDA LA DUPLICIDAD DE Movimiento EN EL TEMPORAL                  
   INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)                   
   SELECT @Emp_cCodigo,@cTab_cCodProc, 'CABECERA VOUCHER', 'CNC_ASIENTO_VOUCHER', 'El Movimiento ' + Ase_cNummov + ' para el año ' + Pan_cAnio + ', esta repetido.','', @User,GETDATE(), host_name()                   
   FROM #TMP_C_ASIENTO_VOUCHER    
   GROUP BY Ase_cNummov, Pan_cAnio                  
   HAVING count(Ase_cNummov) > 1               
                    
   --VALIDA LA DUPLICIDAD DEL Voucher EN EL TEMPORAL                  
   INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)                   
  SELECT @Emp_cCodigo,@cTab_cCodProc, 'CABECERA VOUCHER', 'CNC_ASIENTO_VOUCHER', 'El Voucher ' + Ase_nVoucher + ' del Libro ' + Lib_cTipoLibro + ' para el periodo ' + Per_cPeriodo + '-' + Pan_cAnio + ', esta repetido.','', @User,GETDATE(), host_name()   
    FROM #TMP_C_ASIENTO_VOUCHER                  
   GROUP BY Pan_cAnio, Per_cPeriodo, Lib_cTipoLibro, Ase_nVoucher                  
   HAVING count(Pan_cAnio + Per_cPeriodo + Lib_cTipoLibro + Ase_nVoucher) > 1                  
                    
   -- VALIDA LA FECHA DE REGISTRO DEL VOUCHER PERTENEZCA AL PERIODO                  
   INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)                   
   SELECT @Emp_cCodigo,@cTab_cCodProc, 'CABECERA VOUCHER', 'CNC_ASIENTO_VOUCHER', 'La fecha ' + CONVERT(char(10), Ase_dFecha, 103) + ' del Voucher ' + Ase_nVoucher + ' del Libro ' + Lib_cTipoLibro + ', no pertenece al periodo '         
   + Per_cPeriodo + '-' + Pan_cAnio + '.' ,'', @User,GETDATE(), host_name()                   
   FROM #TMP_C_ASIENTO_VOUCHER                  
   WHERE CONVERT(char(6), Ase_dFecha, 112) <> Pan_cAnio + Per_cPeriodo and Per_cPeriodo<>'00'                  
           
   -- VALIDA QUE EL VOUCHER EMPIECE CON EL NUMERO DE LIBRO Y EL PERIODO                  
   INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)                   
   SELECT @Emp_cCodigo,@cTab_cCodProc, 'CABECERA VOUCHER', 'CNC_ASIENTO_VOUCHER', 'El Voucher ' + Ase_nVoucher + ' no coincide con el Libro y periodos ingresados.','', @User,GETDATE(), host_name()                   
   FROM #TMP_C_ASIENTO_VOUCHER                  
   WHERE Lib_cTipoLibro + Per_cPeriodo <> LEFT(Ase_nVoucher, 4)                  
                    
   -- VALIDA QUE EL AÐO Y PERIODO EXISTAN EN EL MAESTRO DE PERIODOS                  
   INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)                   
   SELECT @Emp_cCodigo,@cTab_cCodProc, 'CABECERA VOUCHER', 'CNC_ASIENTO_VOUCHER', 'El Periodo ' + Pan_cAnio + '-' + Per_cPeriodo + ', no existe en el Maestro de Periodos.','', @User,GETDATE(), host_name()                   
   FROM #TMP_C_ASIENTO_VOUCHER                  
   WHERE Pan_cAnio + Per_cPeriodo NOT IN (SELECT Pan_cAnio + Per_cPeriodo FROM CNT_PERIODO WITH(NOLOCK) WHERE Emp_cCodigo = @Emp_cCodigo )                  
          
   -- VALIDA QUE EL LIBRO EXISTA EN EL MAESTRO DE LIBROS                  
   INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)                   
   SELECT @Emp_cCodigo,@cTab_cCodProc, 'CABECERA VOUCHER', 'CNC_ASIENTO_VOUCHER', 'El Libro Contable ' + Lib_cTipoLibro + ', no existe en el Maestro de Libros.','', @User,GETDATE(), host_name()                   
   FROM #TMP_C_ASIENTO_VOUCHER                  
   WHERE Lib_cTipoLibro NOT IN (SELECT Lib_cTipoLibro FROM CNT_LIBRO_OPERA WITH(NOLOCK) WHERE Emp_cCodigo = @Emp_cCodigo )                  
                    
   -- VALIDA QUE LA MONEDA EXISTA EN EL MAESTRO DE MONEDAS                  
   INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)                   
   SELECT @Emp_cCodigo,@cTab_cCodProc, 'CABECERA VOUCHER', 'CNC_ASIENTO_VOUCHER', 'La Moneda ' + Ase_cTipoMoneda + ', no existe en el Maestro de Monedas o no es la Monedas seleccionada como Nacional o Extranjera.','', @User,GETDATE(), host_name()         
   FROM #TMP_C_ASIENTO_VOUCHER                  
   WHERE Ase_cTipoMoneda NOT IN (SELECT Mon_cCodigo FROM CNT_TIPO_MONEDA WITH(NOLOCK) WHERE Emp_cCodigo = @Emp_cCodigo AND Mon_cMNac + Mon_cMExt IN ('10', '01') )                  
        
   -- VALIDA LA DUPLICIDAD DEL Movimiento EN EL REGISTRO DE VOUCHER                  
   INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)                   
   SELECT @Emp_cCodigo,@cTab_cCodProc, 'CABECERA VOUCHER', 'CNC_ASIENTO_VOUCHER', 'El Movimiento ' + Ase_cNummov + ' para el año ' + Pan_cAnio + ' ya existe en el Registro de Asientos Contables.','', @User,GETDATE(), host_name()                   
   FROM #TMP_C_ASIENTO_VOUCHER                  
   WHERE Pan_cAnio + Ase_cNummov IN (SELECT Pan_cAnio + Ase_cNummov FROM CND_ASIENTO_VOUCHER WITH(NOLOCK) WHERE Emp_cCodigo = @Emp_cCodigo  AND Asd_cDeleted <> '*')                  
                    
   --VALIDA LA DUPLICIDAD DEL Voucher EN EL REGISTRO DE VOUCHER                  
   INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)                   
   SELECT @Emp_cCodigo,@cTab_cCodProc, 'CABECERA VOUCHER', 'CNC_ASIENTO_VOUCHER', 'El Voucher ' + Ase_nVoucher + ' del Libro ' + Lib_cTipoLibro + ' para el periodo ' + Per_cPeriodo + '-' + Pan_cAnio + ', esta repetido.','', @User,GETDATE(), host_name()   
   FROM #TMP_C_ASIENTO_VOUCHER         
   WHERE Pan_cAnio + Per_cPeriodo + Lib_cTipoLibro + Ase_nVoucher IN                   
    (SELECT Pan_cAnio + Per_cPeriodo + Lib_cTipoLibro + Ase_nVoucher                   
    FROM CND_ASIENTO_VOUCHER WITH(NOLOCK) WHERE Emp_cCodigo = @Emp_cCodigo AND Asd_cDeleted <> '*')                  
                    
   -- 4. CREAMOS LA TABLA TEMPORAL DE DETALLE DE VOUCHER                  
   SELECT TOP 0 * INTO #TMP_D_ASIENTO_VOUCHER FROM CND_ASIENTO_VOUCHER WITH(NOLOCK) WHERE Emp_cCodigo = @Emp_cCodigo                   
                  
                     
   -- 5. LLENAMOS LA TABLA TEMPORAL DE DETALLE DE VOUCHER                  
   IF @Par_cApe = '1'                   
   BEGIN                  
    INSERT INTO #TMP_D_ASIENTO_VOUCHER                  
     (Ase_cNummov, Emp_cCodigo, Pan_cAnio, Per_cPeriodo, Lib_cTipoLibro, Ase_nVoucher, Pla_cCuentaContable, Asd_nItem, Asd_cGlosa                  
     , Asd_nDebeSoles, Asd_nHaberSoles, Asd_nTipoCambio, Asd_nDebeMonExt, Asd_nHaberMonExt, Cos_cCodigo                  
     , Ten_cTipoEntidad, Ent_cCodEntidad, Asd_cTipoDoc, Asd_dFecDoc, Asd_cSerieDoc, Asd_cNumDoc, Asd_dFecVen                  
     , Asd_cProvCanc, Asd_cOperaTC, Asd_cTipoMoneda                  
     , Asd_cTipoDocRef, Asd_cSerieDocRef, Asd_cNumDocRef, Com_cTipoIgv, Asd_nMontoInafecto, Asd_cBaseImp, Asd_cRetencion                  
     , Asd_cDestino, Asd_nCorre, Imp_nPorcentaje, Asd_cMonedaCalculo, Tra_cCodigo, Asd_cFormaPago, Asd_cMonAdic, Asd_cImpAdic, Asd_cComprobante                  
     , Asd_cEstado, Asd_cDeleted, Asd_cUserCrea, Asd_dFechaCrea, Asd_cEquipoUser)                  
    SELECT Ase_cNummov, @Emp_cCodigo, Pan_cAnio, Per_cPeriodo, Lib_cTipoLibro, Ase_nVoucher, Pla_cCuentaContable, Asd_nItem, Asd_cGlosa                  
     , ISNULL(round(Asd_nDebeSoles,2), 0), ISNULL(round(Asd_nHaberSoles,2), 0), ISNULL(round(Asd_nTipoCambio,2), 0), ISNULL(round(Asd_nDebeMonExt,2), 0), ISNULL(round(Asd_nHaberMonExt,2), 0), ISNULL(Cos_cCodigo, '')                  
     , ISNULL(Ten_cTipoEntidad, ''), ISNULL(Ent_cCodEntidad, ''), ISNULL(Asd_cTipoDoc, ''), Asd_dFecDoc, ISNULL(Asd_cSerieDoc, ''), ISNULL(Asd_cNumDoc, ''), Asd_dFecVen                  
     , ISNULL(Asd_cProvCanc, ''), ISNULL(Asd_cOperaTC, ''), ISNULL(Asd_cTipoMoneda, '')                  
     , '', '', '', '', 0.00, '', ''                  
     , '0', 0, 0.00, '', '', '', '', '0', ''                  
     , 'A', '', @User, GETDATE(), host_name()                  
    FROM ZIMP_APE_DET WITH(NOLOCK)                  
   END                  
                    
   IF @Par_cPlan = '1'                   
   BEGIN                  
    INSERT INTO #TMP_D_ASIENTO_VOUCHER                  
     (Ase_cNummov, Emp_cCodigo, Pan_cAnio, Per_cPeriodo, Lib_cTipoLibro, Ase_nVoucher, Pla_cCuentaContable, Asd_nItem, Asd_cGlosa                  
     , Asd_nDebeSoles, Asd_nHaberSoles, Asd_nTipoCambio, Asd_nDebeMonExt, Asd_nHaberMonExt, Cos_cCodigo                  
     , Ten_cTipoEntidad, Ent_cCodEntidad, Asd_cTipoDoc, Asd_dFecDoc, Asd_cSerieDoc, Asd_cNumDoc, Asd_dFecVen                  
     , Asd_cProvCanc, Asd_cOperaTC, Asd_cTipoMoneda                  
     , Asd_cTipoDocRef, Asd_cSerieDocRef, Asd_cNumDocRef, Com_cTipoIgv, Asd_nMontoInafecto, Asd_cBaseImp, Asd_cRetencion                  
     , Asd_cDestino, Asd_nCorre, Imp_nPorcentaje, Asd_cMonedaCalculo, Tra_cCodigo, Asd_cFormaPago, Asd_cMonAdic, Asd_cImpAdic, Asd_cComprobante                  
     , Asd_cEstado, Asd_cDeleted, Asd_cUserCrea, Asd_dFechaCrea, Asd_cEquipoUser)                  
    SELECT Ase_cNummov, @Emp_cCodigo, Pan_cAnio, Per_cPeriodo, Lib_cTipoLibro, Ase_nVoucher, Pla_cCuentaContable, Asd_nItem, Asd_cGlosa                  
     , ISNULL(Asd_nDebeSoles, 0), ISNULL(Asd_nHaberSoles, 0), ISNULL(Asd_nTipoCambio, 0), ISNULL(Asd_nDebeMonExt, 0), ISNULL(Asd_nHaberMonExt, 0), ISNULL(Cos_cCodigo, '')                  
     , ISNULL(Ten_cTipoEntidad, ''), ISNULL(Ent_cCodEntidad, ''), ISNULL(Asd_cTipoDoc, ''), Asd_dFecDoc, ISNULL(Asd_cSerieDoc, ''), ISNULL(Asd_cNumDoc, ''), Asd_dFecVen                  
     , ISNULL(Asd_cProvCanc, ''), ISNULL(Asd_cOperaTC, ''), ISNULL(Asd_cTipoMoneda, '')                  
     , '', '', '', '', 0.00, '', ''                  
     , '0', 0, 0.00, '', '', '', '', '0', ''                  
     , 'A', '', @User, GETDATE(), host_name()                  
    FROM ZIMP_PLAN_DET WITH(NOLOCK)                  
   END                  
                       
   IF @Par_cCom = '1'                  
   BEGIN                  
    INSERT INTO #TMP_D_ASIENTO_VOUCHER                  
     (Ase_cNummov, Emp_cCodigo, Pan_cAnio, Per_cPeriodo, Lib_cTipoLibro, Ase_nVoucher, Pla_cCuentaContable, Asd_nItem, Asd_cGlosa                  
     , Asd_nDebeSoles, Asd_nHaberSoles, Asd_nTipoCambio, Asd_nDebeMonExt, Asd_nHaberMonExt, Cos_cCodigo                  
     , Ten_cTipoEntidad, Ent_cCodEntidad, Asd_cTipoDoc, Asd_dFecDoc, Asd_cSerieDoc, Asd_cNumDoc, Asd_dFecVen                  
     , Asd_cProvCanc, Asd_cOperaTC, Asd_cTipoMoneda, Com_cTipoIgv                  
     , Asd_cTipoDocRef, Asd_dFecDocRef, Asd_cSerieDocRef, Asd_cNumDocRef, Asd_nMontoInafecto, Asd_cBaseImp, Asd_cRetencion, Asd_dFechaSpot, Asd_cNumSpot                  
     , Asd_cDestino, Asd_nCorre, Imp_nPorcentaje, Asd_cMonedaCalculo, Tra_cCodigo, Asd_cFormaPago, Asd_cMonAdic, Asd_cImpAdic, Asd_cComprobante                  
     , Asd_cEstado, Asd_cDeleted, Asd_cUserCrea, Asd_dFechaCrea, Asd_cEquipoUser, Id_Exoneracion, Id_Tipo_Renta, Id_Modalidad, Id_Aduana, Id_Clasific_Servicio)                  
    SELECT Ase_cNummov, @Emp_cCodigo, Pan_cAnio, Per_cPeriodo, Lib_cTipoLibro, Ase_nVoucher, Pla_cCuentaContable, Asd_nItem, Asd_cGlosa                  
     , ISNULL(Asd_nDebeSoles, 0), ISNULL(Asd_nHaberSoles, 0), ISNULL(Asd_nTipoCambio, 0), ISNULL(Asd_nDebeMonExt, 0), ISNULL(Asd_nHaberMonExt, 0), ISNULL(Cos_cCodigo, '')                  
     , ISNULL(Ten_cTipoEntidad, ''), ISNULL(Ent_cCodEntidad, ''), ISNULL(Asd_cTipoDoc, ''), Asd_dFecDoc, ISNULL(Asd_cSerieDoc, ''), ISNULL(Asd_cNumDoc, ''), Asd_dFecVen                  
     , ISNULL(Asd_cProvCanc, ''), ISNULL(Asd_cOperaTC, ''), ISNULL(Asd_cTipoMoneda, ''), ''                  
     , Asd_cTipoDocRef, Asd_dFecDocRef, Asd_cSerieDocRef, Asd_cNumDocRef, CASE ISNULL(Asd_cBaseImp, '') WHEN '999' THEN 1 ELSE  0 END, ISNULL(Asd_cBaseImp, ''), CASE ISNULL(Asd_cRetencion, '') WHEN '' THEN '' ELSE 'D' END, Asd_dFechaSpot, Asd_cNumSpot    
 
    
     , '0', 0, 0.00, '', '', '', '', '0', Asd_cComprobante                  
     , 'A', '', @User, GETDATE(), host_name(), Id_Exoneracion, Id_Tipo_Renta, Id_Modalidad, Id_Aduana, Id_Clasific_Servicio                  
    FROM ZIMP_COMPRAS_DET WITH(NOLOCK)                  
   END                  
        
   IF @Par_cVen = '1'                  
   BEGIN                  
    INSERT INTO #TMP_D_ASIENTO_VOUCHER                  
     (Ase_cNummov, Emp_cCodigo, Pan_cAnio, Per_cPeriodo, Lib_cTipoLibro, Ase_nVoucher, Pla_cCuentaContable, Asd_nItem, Asd_cGlosa     
                  
     , Asd_nDebeSoles, Asd_nHaberSoles, Asd_nTipoCambio, Asd_nDebeMonExt, Asd_nHaberMonExt, Cos_cCodigo  
                     
     , Ten_cTipoEntidad, Ent_cCodEntidad, Asd_cTipoDoc, Asd_dFecDoc, Asd_cSerieDoc, Asd_cNumDoc, Asd_dFecVen   
                    
     , Asd_cProvCanc, Asd_cOperaTC, Asd_cTipoMoneda 
                      
     , Asd_cTipoDocRef, Asd_dFecDocRef, Asd_cSerieDocRef, Asd_cNumDocRef, Com_cTipoIgv, Asd_nMontoInafecto, Asd_cBaseImp, Asd_cRetencion                  
     , Asd_cDestino, Asd_nCorre, Imp_nPorcentaje, Asd_cMonedaCalculo, Tra_cCodigo, Asd_cFormaPago, Asd_cMonAdic, Asd_cImpAdic, Asd_cComprobante                  
     , Asd_cEstado, Asd_cDeleted, Asd_cUserCrea, Asd_dFechaCrea, Asd_cEquipoUser, Id_Exoneracion, Id_Tipo_Renta, Id_Modalidad, Id_Aduana, Id_Clasific_Servicio)      
                 
    SELECT Ase_cNummov, @Emp_cCodigo, Pan_cAnio, Per_cPeriodo, Lib_cTipoLibro, Ase_nVoucher, Pla_cCuentaContable, Asd_nItem, Asd_cGlosa
                      
     , ISNULL(Asd_nDebeSoles, 0), ISNULL(Asd_nHaberSoles, 0), ISNULL(Asd_nTipoCambio, 0), ISNULL(Asd_nDebeMonExt, 0), ISNULL(Asd_nHaberMonExt, 0), ISNULL(Cos_cCodigo, '')   
                    
     , ISNULL(Ten_cTipoEntidad, ''), ISNULL(Ent_cCodEntidad, ''), ISNULL(Asd_cTipoDoc, ''), Asd_dFecDoc, ISNULL(Asd_cSerieDoc, ''), ISNULL(Asd_cNumDoc, ''), Asd_dFecVen  
                
     , ISNULL(Asd_cProvCanc, ''), ISNULL(Asd_cOperaTC, ''), ISNULL(Asd_cTipoMoneda, '')   
                    
     , Asd_cTipoDocRef, Asd_dFecDocRef, Asd_cSerieDocRef, Asd_cNumDocRef, '', CASE ISNULL(Asd_cBaseImp, '') WHEN '999' THEN 1 WHEN '998' THEN 1 ELSE  0 END, ISNULL(Asd_cBaseImp, ''), ''                  
     , '0', 0, 0.00, '', '', '', '', '0', ''                  
     , 'A', '', @User, GETDATE(), host_name(), Id_Exoneracion, Id_Tipo_Renta, Id_Modalidad, Id_Aduana, Id_Clasific_Servicio                   
    FROM ZIMP_VENTAS_DET WITH(NOLOCK)                  
   END       
   
   
   
   if @Par_cCin = '1' or @Par_cCeg  = '1'
   begin
	  create table #TMPProvisiones (Cnp_nCorre integer, Ten_cTipoEntidad char(1),Ent_cCodEntidad char(5),Ent_cPersona varchar (120), Pla_cCuentaContable varchar(12),
	  Pla_cNombreCuenta varchar(120),Pla_cCentroCosto char(1), Pla_cProvision char(1), Pla_cDocumento char(1), Asd_cGlosa varchar(120),Cnp_nMonSolProv decimal(14,3), 
	  Cnp_nMonSolCancel decimal(14,2), Cnp_nMonExtProv decimal(14,2), Cnp_nMonExtCancel decimal(14,2), Asd_nTipoCambio decimal(14,3), Cnp_cDebeHaber char(1), 
	  SdoSoles decimal(14,2), SdoDolares decimal(14,2), Asd_cTipoDoc char(2), Asd_cSerieDoc VARCHAR(20), Asd_cNumDoc VARCHAR(25),Asd_dFecDoc datetime, Ase_dFecha datetime, 
	  Ase_cTipoMoneda char(3), ent_nRuc varchar(15), Asd_cOperaTC char(3), Mone char(3), NomMoneCorto varchar(20), FlgMonNac char(1), TipoEntiCta char(1), Asd_dFecVen datetime, 
	  Cos_cCodigo varchar(12), cos_cdescripcion varchar(300),TDO_CNOMBRELARGO varchar(60), Lib_cTipoLibro char(2), DifCambio varchar(5), Ase_cCuadreManual char(1), 
	  Ase_cNumMov char(10), Ase_nVoucher char(10), Diferencia numeric (14,4)) 
	  
	  insert into #TMPProvisiones exec ('spCn_ConsultaProvisiones ''SEL_PEND'', ''' + @Emp_cCodigo + ''', ''' + @Pan_cAnio + ''','''', ''01/01/' + @Pan_cAnio + ''', ''31/12/' + @Pan_cAnio + ''', ''MIX'','''',''''')
   end 
              
        
   IF @Par_cCin = '1'                  
   BEGIN                  
    INSERT INTO #TMP_D_ASIENTO_VOUCHER                  
     (Ase_cNummov, Emp_cCodigo, Pan_cAnio, Per_cPeriodo, Lib_cTipoLibro, Ase_nVoucher, Pla_cCuentaContable, Asd_nItem, Asd_cGlosa                  
     , Asd_nDebeSoles, Asd_nHaberSoles, Asd_nTipoCambio, Asd_nDebeMonExt, Asd_nHaberMonExt, Cos_cCodigo                  
     , Ten_cTipoEntidad, Ent_cCodEntidad, Asd_cTipoDoc, Asd_dFecDoc, Asd_cSerieDoc, Asd_cNumDoc, Asd_dFecVen                  
     , Asd_cProvCanc, Asd_cOperaTC, Asd_cTipoMoneda                  
     , Asd_cTipoDocRef, Asd_dFecDocRef, Asd_cSerieDocRef, Asd_cNumDocRef, Com_cTipoIgv, Asd_nMontoInafecto, Asd_cBaseImp, Asd_cRetencion                  
     , Asd_cDestino, Asd_nCorre, Imp_nPorcentaje, Asd_cMonedaCalculo, Tra_cCodigo, Asd_cFormaPago, Asd_cMonAdic, Asd_cImpAdic, Asd_cComprobante                  
     , Asd_cEstado, Asd_cDeleted, Asd_cUserCrea, Asd_dFechaCrea, Asd_cEquipoUser)                  
    SELECT Ase_cNummov, @Emp_cCodigo, Pan_cAnio, Per_cPeriodo, Lib_cTipoLibro, Ase_nVoucher, Pla_cCuentaContable, Asd_nItem, Asd_cGlosa                  
     , ISNULL(Asd_nDebeSoles, 0), ISNULL(Asd_nHaberSoles, 0), ISNULL(Asd_nTipoCambio, 0), ISNULL(Asd_nDebeMonExt, 0), ISNULL(Asd_nHaberMonExt, 0), ISNULL(Cos_cCodigo, '')                  
     , ISNULL(Ten_cTipoEntidad, ''), ISNULL(Ent_cCodEntidad, ''), ISNULL(Asd_cTipoDoc, ''), Asd_dFecDoc, ISNULL(Asd_cSerieDoc, ''), ISNULL(Asd_cNumDoc, ''), Asd_dFecVen                  
     , ISNULL(Asd_cProvCanc, ''), ISNULL(Asd_cOperaTC, ''), ISNULL(Asd_cTipoMoneda, '')                  
     , Asd_cTipoDocRef, Asd_dFecDocRef, Asd_cSerieDocRef, Asd_cNumDocRef, '', 0.00, '', ''         
     , '0', 0, 0.00, '', ISNULL(Tra_cCodigo, ''), ISNULL(Asd_cFormaPago, ''), '', '0', ''            
     , 'A', '', @User, GETDATE(), host_name()                  
    FROM ZIMP_CAJAING_DET WITH(NOLOCK)           
   END        
        
   IF @Par_cCeg = '1'                  
   BEGIN                  
    INSERT INTO #TMP_D_ASIENTO_VOUCHER                  
     (Ase_cNummov, Emp_cCodigo, Pan_cAnio, Per_cPeriodo, Lib_cTipoLibro, Ase_nVoucher, Pla_cCuentaContable, Asd_nItem, Asd_cGlosa                  
     , Asd_nDebeSoles, Asd_nHaberSoles, Asd_nTipoCambio, Asd_nDebeMonExt, Asd_nHaberMonExt, Cos_cCodigo                  
     , Ten_cTipoEntidad, Ent_cCodEntidad, Asd_cTipoDoc, Asd_dFecDoc, Asd_cSerieDoc, Asd_cNumDoc, Asd_dFecVen                  
     , Asd_cProvCanc, Asd_cOperaTC, Asd_cTipoMoneda                  
     , Asd_cTipoDocRef, Asd_dFecDocRef, Asd_cSerieDocRef, Asd_cNumDocRef, Com_cTipoIgv, Asd_nMontoInafecto, Asd_cBaseImp, Asd_cRetencion                  
     , Asd_cDestino, Asd_nCorre, Imp_nPorcentaje, Asd_cMonedaCalculo, Tra_cCodigo, Asd_cFormaPago, Asd_cMonAdic, Asd_cImpAdic, Asd_cComprobante                  
     , Asd_cEstado, Asd_cDeleted, Asd_cUserCrea, Asd_dFechaCrea, Asd_cEquipoUser)                  
    SELECT Ase_cNummov, @Emp_cCodigo, Pan_cAnio, Per_cPeriodo, Lib_cTipoLibro, Ase_nVoucher, Pla_cCuentaContable, Asd_nItem, Asd_cGlosa                  
     , ISNULL(Asd_nDebeSoles, 0), ISNULL(Asd_nHaberSoles, 0), ISNULL(Asd_nTipoCambio, 0), ISNULL(Asd_nDebeMonExt, 0), ISNULL(Asd_nHaberMonExt, 0), ISNULL(Cos_cCodigo, '')                  
     , ISNULL(Ten_cTipoEntidad, ''), ISNULL(Ent_cCodEntidad, ''), ISNULL(Asd_cTipoDoc, ''), Asd_dFecDoc, ISNULL(Asd_cSerieDoc, ''), ISNULL(Asd_cNumDoc, ''), Asd_dFecVen                  
     , ISNULL(Asd_cProvCanc, ''), ISNULL(Asd_cOperaTC, ''), ISNULL(Asd_cTipoMoneda, '')                  
     , Asd_cTipoDocRef, Asd_dFecDocRef, Asd_cSerieDocRef, Asd_cNumDocRef, '', 0.00, '', ''                  
, '0', 0, 0.00, '', ISNULL(Tra_cCodigo, ''), ISNULL(Asd_cFormaPago, ''), '', '0', ''                  
     , 'A', '', @User, GETDATE(), host_name()                  
    FROM ZIMP_CAJAEGR_DET WITH(NOLOCK)                  
   END                  
        
   IF @@ERROR <> 0                   
   BEGIN                  
    INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK) (Emp_cCodigo, Tab_cCodProc, Tab_cProceso, Tab_cTabla, Tab_cError, Tab_cMensaje, Tab_cUserCrea, Tab_dFechaCrea, Tab_cEquipoUser )                  
    VALUES(@Emp_cCodigo,@cTab_cCodProc, 'DETALLE VOUCHER', 'CND_ASIENTO_VOUCHER', 'Error al insertar los valores en la tabla CND_ASIENTO_VOUCHER' ,'verifique los campos obligatorios', @User,GETDATE(), host_name() )                  
    GOTO tagFinMov                  
   END                  
                  
   -- 6. VALIDAMOS LA TABLA DE DETALLE DE VOUCHER                  
        
   -- VERIFICAR QUE EL DETALLE EXISTE EN LA CABECERA                  
   INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)                   
   SELECT @Emp_cCodigo,@cTab_cCodProc, 'DETALLE VOUCHER', 'CND_ASIENTO_VOUCHER', 'El Movimiento ' + Ase_cNummov + ' con Voucher ' + Ase_nVoucher + ' del Libro ' + Lib_cTipoLibro + ' para el periodo ' + Per_cPeriodo + '-' + Pan_cAnio + ', no esta en la Cabecera.','', @User,GETDATE(), host_name()                   
   FROM #TMP_D_ASIENTO_VOUCHER                  
   WHERE Ase_cNummov + Pan_cAnio + Per_cPeriodo + Lib_cTipoLibro + Ase_nVoucher                   
    NOT IN  (SELECT Ase_cNummov + Pan_cAnio + Per_cPeriodo + Lib_cTipoLibro + Ase_nVoucher FROM #TMP_C_ASIENTO_VOUCHER )                  
                    
   -- VERIFICAR QUE LA CABECERA TENGA DETALLE                  
   INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)                   
   SELECT @Emp_cCodigo,@cTab_cCodProc, 'DETALLE VOUCHER', 'CND_ASIENTO_VOUCHER', 'El Movimiento ' + Ase_cNummov + ' con Voucher ' + Ase_nVoucher + ' del Libro ' + Lib_cTipoLibro + ' para el periodo ' + Per_cPeriodo + '-' + Pan_cAnio + ', no tiene detalle.
  
         
','', @User,GETDATE(), host_name()                   
   FROM #TMP_C_ASIENTO_VOUCHER                  
   WHERE Ase_cNummov + Pan_cAnio + Per_cPeriodo + Lib_cTipoLibro + Ase_nVoucher                   
    NOT IN  (SELECT Ase_cNummov + Pan_cAnio + Per_cPeriodo + Lib_cTipoLibro + Ase_nVoucher FROM #TMP_D_ASIENTO_VOUCHER )                  
                    
   -- VERIFICAR QUE NO SE DUPLIQUE LOS ITEMS DEL DETALLE                  
   INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)                   
   SELECT @Emp_cCodigo,@cTab_cCodProc, 'DETALLE VOUCHER', 'CND_ASIENTO_VOUCHER', 'El Item ' + str(Asd_nItem, 3) + ' del Movimiento ' + Ase_cNummov + ' para el año ' + Pan_cAnio + ', esta repetido.','', @User,GETDATE(), host_name()                   
   FROM #TMP_D_ASIENTO_VOUCHER                  
   GROUP BY Ase_cNummov, Pan_cAnio, Per_cPeriodo, Lib_cTipoLibro, Ase_nVoucher, Asd_nItem                  
   HAVING count(Ase_cNummov + Pan_cAnio + Per_cPeriodo + Lib_cTipoLibro + Ase_nVoucher + DBO.TRIMSQL(STR(Asd_nItem))) > 1                  
                    
   -- VALIDAR QUE EL TOTAL DEL DEBE Y DEL HABER EN MONEDA NACIONAL SEAN IGUAL                  
   INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)                   
   SELECT @Emp_cCodigo,@cTab_cCodProc, 'DETALLE VOUCHER', 'CND_ASIENTO_VOUCHER', 'El Movimiento ' + Ase_cNummov + ' con Voucher ' + Ase_nVoucher + ' del Libro ' + Lib_cTipoLibro + ' para el periodo ' + Per_cPeriodo + '-' + Pan_cAnio + ', tiene una diferencia entre el Total del Debe y el Haber en Moneda Nacional.','', @User,GETDATE(), host_name()                   
   FROM #TMP_D_ASIENTO_VOUCHER  
   GROUP BY Ase_cNummov, Pan_cAnio, Per_cPeriodo, Lib_cTipoLibro, Ase_nVoucher                  
   HAVING SUM(Asd_nDebeSoles) - SUM(Asd_nHaberSoles) <> 0                  
                    
   -- VALIDAR QUE EL TOTAL DEL DEBE Y DEL HABER EN MONEDA EXTRANJERA SEAN IGUAL                  
   IF EXISTS(SELECT Emp_cCodigo FROM EMPRESA WHERE Emp_cCodigo = @Emp_cCodigo AND ISNULL(Emp_ByMoneda, '') = '1')                  
   BEGIN                  
    INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)                   
    SELECT @Emp_cCodigo,@cTab_cCodProc, 'DETALLE VOUCHER', 'CND_ASIENTO_VOUCHER', 'El Movimiento ' + Ase_cNummov + ' con Voucher ' + Ase_nVoucher + ' del Libro ' + Lib_cTipoLibro + ' para el periodo ' + Per_cPeriodo + '-' + Pan_cAnio + ', tiene una diferencia entre el Total del Debe y el Haber en Moneda Extranjera.','', @User,GETDATE(), host_name()                   
    FROM #TMP_D_ASIENTO_VOUCHER                  
 GROUP BY Ase_cNummov, Pan_cAnio, Per_cPeriodo, Lib_cTipoLibro, Ase_nVoucher                  
    HAVING SUM(Asd_nDebeMonExt) - SUM(Asd_nHaberMonExt) <> 0                  
   END                  
                    
   -- VALIDAR QUE NO SE REPITA CUENTA CONTABLE, ENTIDAD Y DOCUMENTO PARA LAS PROVISIONES EN EL TEMPORAL                  
   INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)                   
   SELECT @Emp_cCodigo,@cTab_cCodProc, 'DETALLE VOUCHER', 'CND_ASIENTO_VOUCHER', 'El Documento ' + Asd_cTipoDoc + ' - ' + Asd_cSerieDoc + '-' + Asd_cNumDoc + ' de la entidad ' + Ten_cTipoEntidad + '-' + Ent_cCodEntidad + ' y cuenta ' + pla_ccuentacontable
  + ' , esta repetido.','', @User,GETDATE(), host_name()                
   FROM #TMP_D_ASIENTO_VOUCHER                  
   WHERE ISNULL(Asd_cProvCanc, '') = 'P' and ( Asd_cTipoDoc<>'98' or Asd_cTipoDoc<>'99' )                  
   GROUP BY pla_ccuentacontable, Ten_cTipoEntidad, Ent_cCodEntidad, Asd_cTipoDoc, Asd_cSerieDoc, Asd_cNumDoc                  
   HAVING count(pla_ccuentacontable + Ten_cTipoEntidad + Ent_cCodEntidad + Asd_cTipoDoc + Asd_cSerieDoc + Asd_cNumDoc) > 1                  
                  
   -- VALIDAR QUE NO SE REPITA LA CUENTA CONTABLE, ENTIDAD Y DOCUMENTO PARA LAS PROVISIONES Y CANCELACIONES EN EL REGISTRO DE VOUCHERS     
   
   if @Par_cCin = '1' or @Par_cCeg  = '1' 
   BEGIN   
                 
	   INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)                   
	   SELECT @Emp_cCodigo,@cTab_cCodProc, 'DETALLE VOUCHER', 'CND_ASIENTO_VOUCHER', 'El Documento ' + Asd_cTipoDoc + ' - ' + Asd_cSerieDoc + '-' + Asd_cNumDoc + ' de la entidad ' + Ten_cTipoEntidad + '-' + Ent_cCodEntidad + ' y cuenta ' + pla_ccuentacontable
	  +  ', no puede ser mayor el monto al saldo pendiente.','', @User,GETDATE(), host_name()                   
	   FROM #TMP_D_ASIENTO_VOUCHER                  
	   WHERE dbo.TRIMSQL ( pla_ccuentacontable) + dbo.TRIMSQL (Ten_cTipoEntidad )+ dbo.TRIMSQL (Ent_cCodEntidad )+ dbo.TRIMSQL (Asd_cTipoDoc )+ dbo.TRIMSQL (Asd_cSerieDoc )+ dbo.TRIMSQL (Asd_cNumDoc ) + ISNULL(Asd_cProvCanc, '')                  
		 IN ( SELECT dbo.TRIMSQL ( pla_ccuentacontable) + dbo.TRIMSQL (Ten_cTipoEntidad )+ dbo.TRIMSQL (Ent_cCodEntidad )+ dbo.TRIMSQL (Asd_cTipoDoc )+ dbo.TRIMSQL (Asd_cSerieDoc )+ dbo.TRIMSQL (Asd_cNumDoc ) + 'C'                  
		   FROM #TMPProvisiones WITH(NOLOCK) WHERE Emp_cCodigo = @Emp_cCodigo AND pan_canio=@pan_canio and ISNULL(Asd_cProvCanc, '') <> '' AND Asd_cDeleted <> '*'  and  Asd_nHaberSoles > SdoSoles                
		  )   
	end
   else
    begin
	   INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)                   
	   SELECT @Emp_cCodigo,@cTab_cCodProc, 'DETALLE VOUCHER', 'CND_ASIENTO_VOUCHER', 'El Documento ' + Asd_cTipoDoc + ' - ' + Asd_cSerieDoc + '-' + Asd_cNumDoc + ' de la entidad ' + Ten_cTipoEntidad + '-' + Ent_cCodEntidad + ' y cuenta ' + pla_ccuentacontable +  
	   ', ya existe en el Registro de Asientos.','', @User,GETDATE(), host_name()                   
	   FROM #TMP_D_ASIENTO_VOUCHER                  
	   WHERE dbo.TRIMSQL ( pla_ccuentacontable) + dbo.TRIMSQL (Ten_cTipoEntidad )+ dbo.TRIMSQL (Ent_cCodEntidad )+ dbo.TRIMSQL (Asd_cTipoDoc )+ dbo.TRIMSQL (Asd_cSerieDoc )+ dbo.TRIMSQL (Asd_cNumDoc ) + ISNULL(Asd_cProvCanc, '')                  
		 IN ( SELECT dbo.TRIMSQL ( pla_ccuentacontable) + dbo.TRIMSQL (Ten_cTipoEntidad )+ dbo.TRIMSQL (Ent_cCodEntidad )+ dbo.TRIMSQL (Asd_cTipoDoc )+ dbo.TRIMSQL (Asd_cSerieDoc )+ dbo.TRIMSQL (Asd_cNumDoc ) + ISNULL(Asd_cProvCanc, '')                  
		   FROM CND_ASIENTO_VOUCHER WITH(NOLOCK) WHERE Emp_cCodigo = @Emp_cCodigo AND pan_canio=@pan_canio and ISNULL(Asd_cProvCanc, '') <> '' AND Asd_cDeleted <> '*'                  
		  )   
    end
      
   IF @@ROWCOUNT > 0                  
   BEGIN                  
    GOTO tagExit                  
   END                  
                  
-- EXEC spCn_ReprocesoImportacionXLSv2 'IMPORTACION', '037', '2009','0','0','0','0','0','1','0','0','ADMIN'                 
                  
    -- VALIDAR QUE UNA CANCELACION ESTE PROVISIONADA                  
   INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)                   
   SELECT @Emp_cCodigo,@cTab_cCodProc, 'DETALLE VOUCHER', 'CND_ASIENTO_VOUCHER', 'El Movimiento ' + Ase_cNummov + ' con Voucher ' + Ase_nVoucher +  + ' , item ' + DBO.TRIMSQL(str(Asd_nItem)) + ', El Documento ' + Asd_cTipoDoc + ' - ' + Asd_cSerieDoc + '- '  
    + Asd_cNumDoc + ' de la entidad ' + Ten_cTipoEntidad + '-' + Ent_cCodEntidad + ', no ha sido Provisionado.','', @User,GETDATE(), host_name()                   
   FROM #TMP_D_ASIENTO_VOUCHER                  
   WHERE dbo.TRIMSQL ( pla_ccuentacontable) + dbo.TRIMSQL (Ten_cTipoEntidad )+ dbo.TRIMSQL (Ent_cCodEntidad )+ dbo.TRIMSQL (Asd_cTipoDoc )+ dbo.TRIMSQL (Asd_cSerieDoc )+ dbo.TRIMSQL (Asd_cNumDoc )                  
    NOT IN (SELECT dbo.TRIMSQL ( pla_ccuentacontable) + dbo.TRIMSQL (Ten_cTipoEntidad )+ dbo.TRIMSQL (Ent_cCodEntidad )+ dbo.TRIMSQL (Asd_cTipoDoc )+ dbo.TRIMSQL (Asd_cSerieDoc )+ dbo.TRIMSQL (Asd_cNumDoc )                  
     FROM CND_ASIENTO_VOUCHER WITH(NOLOCK) WHERE Emp_cCodigo = @Emp_cCodigo AND pan_canio=@pan_canio  AND ISNULL(Asd_cProvCanc, '') = 'P'  AND Asd_cDeleted <> '*'                   
     UNION                  
     SELECT dbo.TRIMSQL ( pla_ccuentacontable) + dbo.TRIMSQL (Ten_cTipoEntidad )+ dbo.TRIMSQL (Ent_cCodEntidad )+ dbo.TRIMSQL (Asd_cTipoDoc )+ dbo.TRIMSQL (Asd_cSerieDoc )+ dbo.TRIMSQL (Asd_cNumDoc )                  
     FROM #TMP_D_ASIENTO_VOUCHER WITH(NOLOCK) WHERE Emp_cCodigo = @Emp_cCodigo AND pan_canio=@pan_canio  AND ISNULL(Asd_cProvCanc, '') = 'P')                   
    AND ISNULL(Asd_cProvCanc, '') = 'C'                   
            
            
   -- VALIDA QUE SE REGISTREN LOS TIPOS DE PROVISIONES O CANCELACIONES                  
   INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)                   
   SELECT @Emp_cCodigo,@cTab_cCodProc, 'DETALLE VOUCHER', 'CND_ASIENTO_VOUCHER', 'El Movimiento ' + Ase_cNummov + ' con Voucher ' + Ase_nVoucher + ' del Libro ' + Lib_cTipoLibro + ' para el periodo ' + Per_cPeriodo + '-' + Pan_cAnio + ', Debe tener los valores C o P en la columna "Asd_cProvCanc".','', @User,GETDATE(), host_name()                   
   FROM #TMP_D_ASIENTO_VOUCHER                  
   WHERE ISNULL(Asd_cProvCanc, '') NOT IN ('', 'P', 'C')                  
                    
   -- VALIDA LA FECHA DE DETALLE DEL VOUCHER PERTENEZCA AL PERIODO O A UN PERIODO ANTERIOR                  
   INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)                   
   SELECT @Emp_cCodigo,@cTab_cCodProc, 'DETALLE VOUCHER', 'CND_ASIENTO_VOUCHER', 'La fecha ' + CONVERT(char(10), Asd_dFecDoc, 103) + ' del Voucher ' + Ase_nVoucher + ' del Libro ' + Lib_cTipoLibro + ', no pertenece al periodo ' + Per_cPeriodo + '-' +     
  
    
   Pan_cAnio + '.' ,'', @User,GETDATE(), host_name()                   
   FROM #TMP_D_ASIENTO_VOUCHER                  
   WHERE CONVERT(char(6), Asd_dFecDoc, 112) > Pan_cAnio + Per_cPeriodo and Per_cPeriodo<>'00'                   
                    
   -- VALIDAR QUE LA FECHA DE VENCIMIENTO SEA MAYOR O IGUAL A LA FECHA DEL DOCUMENTO                  
   INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)                   
   SELECT @Emp_cCodigo,@cTab_cCodProc, 'DETALLE VOUCHER', 'CND_ASIENTO_VOUCHER', 'El Movimiento ' + Ase_cNummov + ' con Voucher ' + Ase_nVoucher + ' del Libro ' + Lib_cTipoLibro + ' para el periodo ' + Per_cPeriodo + '-' + Pan_cAnio + ' tiene Fecha de Vencimiento menor a la Fecha del Documento.' ,'', @User,GETDATE(), host_name()                   
   FROM #TMP_D_ASIENTO_VOUCHER                  
   WHERE case when ISNULL(Asd_dFecVen, Asd_dFecDoc) = '' Then Asd_dFecDoc Else ISNULL(Asd_dFecVen, Asd_dFecDoc) End < Asd_dFecDoc                  
    --select * from #TMP_D_ASIENTO_VOUCHER                  
                     
                     
   -- VALIDAR QUE LA OPERACION TIPO DE CAMBIO SEA VALIDO                  
   INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)                   
   SELECT @Emp_cCodigo,@cTab_cCodProc, 'DETALLE VOUCHER', 'CND_ASIENTO_VOUCHER', 'La Operacion Tipo de Cambio ' + Asd_cOperaTC + ' del Movimiento ' + Ase_cNummov + ' con Voucher ' + Ase_nVoucher + ', item ' + DBO.TRIMSQL(str(asd_nitem)) + ' del Libro ' + 
Lib_cTipoLibro + ' para el periodo ' + Per_cPeriodo + '-' + Pan_cAnio + ' no es una Operacion valida.' ,'', @User,GETDATE(), host_name()                   
   FROM #TMP_D_ASIENTO_VOUCHER                  
   WHERE Asd_cOperaTC NOT IN ('SCV', 'COM', 'VEN', 'VEP', 'OTR')              
                    
   -- VALIDAR QUE EL TIPO DE CAMBIO PARA LAS OPERACIONES DE TIPO DE CAMBIO                  
   INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)                   
   SELECT @Emp_cCodigo,@cTab_cCodProc, 'DETALLE VOUCHER', 'CND_ASIENTO_VOUCHER',                   
         'La Operacion Tipo de Cambio ' + Asd_cOperaTC + ' del Movimiento ' + Ase_cNummov + ' con Voucher ' + Ase_nVoucher + ', item ' + DBO.TRIMSQL(str(asd_nitem)) + ' del Libro ' + Lib_cTipoLibro + ' para el periodo ' + Per_cPeriodo + '-' + Pan_cAnio   
  
    
         + CASE Asd_cOperaTC WHEN 'SCV' THEN CASE ISNULL(Asd_nTipoCambio, 0) WHEN 0 THEN '' ELSE ' Debe ser CERO.' END                  
         ELSE CASE ISNULL(Asd_nTipoCambio, 0) WHEN 0 THEN ' Debe ser diferente de CERO.' ELSE '' END END                  
         ,'', @User,GETDATE(), host_name()                   
   FROM #TMP_D_ASIENTO_VOUCHER                  
   WHERE ( Asd_cOperaTC = 'SCV' AND ISNULL(Asd_nTipoCambio, 0) <> 0) OR (Asd_cOperaTC <> 'SCV'  AND ISNULL(Asd_nTipoCambio, 0) = 0) AND Asd_cOperaTC <>'OTR'                  
                     
   -- VALIDAR QUE EXISTA CUENTA CONTABLE EN PLAN DE CUENTAS                  
   INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)                   
   SELECT @Emp_cCodigo,@cTab_cCodProc, 'DETALLE VOUCHER', 'CND_ASIENTO_VOUCHER', 'La Cuenta ' + ISNULL(Pla_cCuentaContable, 'NULL') + ' del Movimiento ' + Ase_cNummov + ' con Voucher ' + Ase_nVoucher + ' del Libro ' + Lib_cTipoLibro + ' para el periodo ' 
  
         
+ Per_cPeriodo + '-' + Pan_cAnio + ' no existe en el Plan de Cuentas o no es vßlida.','', @User,GETDATE(), host_name()                   
   FROM #TMP_D_ASIENTO_VOUCHER                  
   WHERE dbo.TRIMSQL( Pan_cAnio) + dbo.TRIMSQL(Pla_cCuentaContable) NOT IN                   
    (SELECT dbo.TRIMSQL(Pan_cAnio) + dbo.TRIMSQL(Pla_cCuentaContable) FROM CNM_PLAN_CTA WITH(NOLOCK) WHERE Emp_cCodigo = @Emp_cCodigo AND Pla_cTitulo <> 'S')                  
                    
   -- VALIDAR QUE EL CENTRO DE COSTO EXISTA                  
   INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)                   
   SELECT @Emp_cCodigo,@cTab_cCodProc, 'DETALLE VOUCHER', 'CND_ASIENTO_VOUCHER', 'El Centro de Costo ' + Cos_cCodigo + ' del Movimiento ' + Ase_cNummov + ' con Voucher ' + Ase_nVoucher + ' del Libro ' + Lib_cTipoLibro + ' para el periodo ' + Per_cPeriodo 
+ '-' + Pan_cAnio + ' no existe en el Maestro de Centros de Costo.' ,'', @User,GETDATE(), host_name()                   
   FROM #TMP_D_ASIENTO_VOUCHER                  
   WHERE Pan_cAnio + dbo.TRIMSQL(Cos_cCodigo) NOT IN (SELECT Pan_cAnio + dbo.TRIMSQL(Cos_cCodigo) FROM CNT_CENTRO_COSTO WITH(NOLOCK) WHERE Emp_cCodigo = @Emp_cCodigo)                  
    AND ISNULL(Cos_cCodigo, '') <> ''                  
                    
   -- VALIDAR QUE LA ENTIDAD EXISTA EN EL MAESTRO DE ENTIDADES                  
   INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)           
   SELECT @Emp_cCodigo,@cTab_cCodProc, 'DETALLE VOUCHER', 'CND_ASIENTO_VOUCHER', 'La Entidad ' + Ten_cTipoEntidad + '-' + Ent_cCodEntidad + ' del Movimiento ' + Ase_cNummov + ' con Voucher ' + Ase_nVoucher + ' del Libro ' + Lib_cTipoLibro + ' para el periodo ' + Per_cPeriodo + '-' + Pan_cAnio + ' no existe en el Maestro de entidades.','', @User,GETDATE(), host_name()                   
   FROM #TMP_D_ASIENTO_VOUCHER                  
   WHERE Ten_cTipoEntidad + Ent_cCodEntidad NOT IN (SELECT Ten_cTipoEntidad + Ent_cCodEntidad FROM CNM_ENTIDAD WITH(NOLOCK) WHERE Emp_cCodigo = @Emp_cCodigo)                  
    AND ISNULL(Ent_cCodEntidad, '') <> ''                  
                    
   -- VALIDAR QUE EL TIPO DE DOCUMENTO EXISTA EN EL MAESTRO DE TIPOS DE DOCUMENTOS                  
   INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)                   
   SELECT @Emp_cCodigo,@cTab_cCodProc, 'DETALLE VOUCHER', 'CND_ASIENTO_VOUCHER', 'Tipo de Documento ' + Asd_cTipoDoc + ' del Movimiento ' + Ase_cNummov + ' con Voucher ' + Ase_nVoucher + ' del Libro ' + Lib_cTipoLibro + ' para el periodo ' + Per_cPeriodo 
+ '-' + Pan_cAnio + ' no existe en el Maestro de Tipo de Documento.','', @User,GETDATE(), host_name()                   
   FROM #TMP_D_ASIENTO_VOUCHER                  
   WHERE Asd_cTipoDoc NOT IN (SELECT Tdo_cCodigo FROM CNT_TIPODOC WITH(NOLOCK) WHERE Emp_cCodigo = @Emp_cCodigo)                  
    AND ISNULL(Asd_cTipoDoc, '') <> ''                  
                    
   -- VALIDAR QUE EL TIPO DE MONEDA EXISTA EN EL MAESTRO DE MONEDAS                  
   INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)                   
   SELECT @Emp_cCodigo,@cTab_cCodProc, 'DETALLE VOUCHER', 'CND_ASIENTO_VOUCHER', 'La Moneda ' + Asd_cTipoMoneda + ' del Movimiento ' + Ase_cNummov + ' con Voucher ' + Ase_nVoucher + ' del Libro ' + Lib_cTipoLibro + ' para el periodo ' + Per_cPeriodo + '-'
 + Pan_cAnio + ' no existe en el Maestro de Monedas o no es la Monedas seleccionada como Nacional o Extranjera.','', @User,GETDATE(), host_name()                   
   FROM #TMP_D_ASIENTO_VOUCHER                  
   WHERE Asd_cTipoMoneda NOT IN ( SELECT Mon_cCodigo FROM CNT_TIPO_MONEDA WITH(NOLOCK)                   
       WHERE Emp_cCodigo = @Emp_cCodigo AND Mon_cMNac + Mon_cMExt IN ('10', '01') )                  
    AND ISNULL(Asd_cTipoMoneda, '') <> ''                  
                    
   -- ESPECIFICAS DE COMPRAS                  
   IF @Par_cCom = '1'                  
   BEGIN                  
    -- VALIDAR EL TIPO DE BASE IMPONIBLE                  
    INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)                  
    SELECT @Emp_cCodigo,@cTab_cCodProc, 'DETALLE VOUCHER', 'CND_ASIENTO_VOUCHER', 'El Tipo de Base Imponible ' + Asd_cBaseImp + ' del Movimiento ' + Ase_cNummov + ' con Voucher ' + Ase_nVoucher + ' del Libro ' +         
    Lib_cTipoLibro + ' para el periodo ' + Per_cPeriodo + '-' + Pan_cAnio + ', NO es valido.','', @User,GETDATE(), host_name()                   
    FROM #TMP_D_ASIENTO_VOUCHER                  
    WHERE Asd_cBaseImp NOT IN ('', '006', '007', '008', '017', '999')                  
                    
    -- VALIDAR QUE LAS DETRACCIONES TENGAN FECHA Y NUMERO                  
    INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)                   
    SELECT @Emp_cCodigo,@cTab_cCodProc, 'DETALLE VOUCHER', 'CND_ASIENTO_VOUCHER', 'El Movimiento ' + Ase_cNummov + ' con Voucher ' + Ase_nVoucher + ' del Libro ' + Lib_cTipoLibro + ' para el periodo ' + Per_cPeriodo + '-' + Pan_cAnio + ' es una detracciónse debe ingresar Fecha y Numero del Comprobante.','', @User,GETDATE(), host_name()                   
    FROM #TMP_D_ASIENTO_VOUCHER                  
    WHERE Asd_cRetencion = 'D' AND (Asd_dFechaSpot IS NULL OR ISNULL(Asd_cNumSpot, '') = '')                  
   END                  
   -- ESPECIFICAS DE VENTAS                  
   IF @Par_cVen = '1'                  
   BEGIN                  
    -- VALIDAR EL TIPO DE BASE IMPONIBLE                  
    INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)                   
    SELECT @Emp_cCodigo,@cTab_cCodProc, 'DETALLE VOUCHER', 'CND_ASIENTO_VOUCHER', 'El Tipo de Base Imponible ' + Asd_cBaseImp + ' del Movimiento ' + Ase_cNummov + ' con Voucher ' + Ase_nVoucher + ' del Libro '         
    + Lib_cTipoLibro + ' para el periodo ' + Per_cPeriodo + '-' + Pan_cAnio + ', NO es valido.','', @User,GETDATE(), host_name()                   
    FROM #TMP_D_ASIENTO_VOUCHER                  
    WHERE Asd_cBaseImp NOT IN ('', '002', '021', '047', '017', '998', '999')                  
   END                  
                    
   -- ESPECIFICAS DE CAJA                  
   IF @Par_cCin = '1' OR @Par_cCeg = '1'                  
   BEGIN                  
    -- VALIDAR QUE EN LA RETENCION SE INGRESEN VALORES DE RETENCION O PERCEPCION                  
    INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)                   
    SELECT @Emp_cCodigo,@cTab_cCodProc, 'DETALLE VOUCHER', 'CND_ASIENTO_VOUCHER', 'El Tipo de Retencion/Percepcion ' + Asd_cRetencion + ' del Movimiento ' + Ase_cNummov + ' con Voucher ' + Ase_nVoucher + ' del Libro ' + Lib_cTipoLibro + ' para el periodo 
  
' + Per_cPeriodo + '-' + Pan_cAnio + ', NO es valido.','', @User,GETDATE(), host_name()                   
    FROM #TMP_D_ASIENTO_VOUCHER                  
    WHERE Asd_cRetencion NOT IN ('', 'R', 'P')                  
                   
    -- VALIDAR SI SE INGRESO RETENCION SE INGRESEN LOS DATOS DEL DOCUMENTO                  
    INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)                   
    SELECT @Emp_cCodigo,@cTab_cCodProc, 'DETALLE VOUCHER', 'CND_ASIENTO_VOUCHER', 'El Movimiento ' + Ase_cNummov + ' con Voucher ' + Ase_nVoucher + ' del Libro ' + Lib_cTipoLibro + ' para el periodo ' + Per_cPeriodo + '-' + Pan_cAnio + ', NO tiene Documento de ' + CASE Asd_cRetencion WHEN 'P' THEN 'Percepcion.' WHEN 'R' THEN 'Retencion.' END,'', @User,GETDATE(), host_name()                   
    FROM #TMP_D_ASIENTO_VOUCHER                  
    WHERE Asd_cRetencion IN ('R', 'P')                  
     AND ( ISNULL(Asd_cTipoDocRef, '') = '' OR ISNULL(Asd_cSerieDocRef, '') = '' OR ISNULL(Asd_cNumDocRef, '') = '' OR Asd_dFecDocRef IS NULL)                  
                    
    -- VALIDAR QUE EL TIPO DE DOCUMENTO DE LA RETENCION SEA VALIDO                  
    INSERT INTO CNT_REPORTE_IMPORTACION  WITH(ROWLOCK)                   
    SELECT @Emp_cCodigo,@cTab_cCodProc, 'DETALLE VOUCHER', 'CND_ASIENTO_VOUCHER', 'El Movimiento ' + Ase_cNummov + ' con Voucher ' + Ase_nVoucher + ' del Libro ' + Lib_cTipoLibro + ' para el periodo ' + Per_cPeriodo + '-' + Pan_cAnio + ' Tiene Tipo de Documento de ' + CASE Asd_cRetencion WHEN 'P' THEN 'Percepcion' WHEN 'R' THEN 'Retencion' END + ' NO valido.','', @User,GETDATE(), host_name()                   
    FROM #TMP_D_ASIENTO_VOUCHER                  
    WHERE Asd_cRetencion IN ('R', 'P')                  
     AND ISNULL(Asd_cTipoDocRef, '') NOT IN (SELECT Tdo_cCodigo FROM CNT_TIPODOC WITH(NOLOCK) WHERE Emp_cCodigo = @Emp_cCodigo)                  
                    
   END                  
                  
   -- 8. INSERTA LOS DATOS VALIDADOS EN LA TABLA CABECERA DE VOUCHER Y DETALLE                  
                  
   SELECT @nRegistros   = count(emp_ccodigo) FROM CNT_REPORTE_IMPORTACION  WHERE Emp_cCodigo= @Emp_cCodigo --and Tab_cCodProc = @cTab_cCodProc                  
                  
   IF isnull(@nRegistros  ,0) < 1                  
   BEGIN                  
    INSERT INTO CNC_ASIENTO_VOUCHER WITH(ROWLOCK) SELECT * FROM #TMP_C_ASIENTO_VOUCHER    
    INSERT INTO CND_ASIENTO_VOUCHER WITH(ROWLOCK) SELECT * FROM #TMP_D_ASIENTO_VOUCHER    
    
    DECLARE @Ase_nVoucherAux VARCHAR(10)
    SELECT @Ase_nVoucherAux = Ase_nVoucher FROM #TMP_D_ASIENTO_VOUCHER
    
    EXEC [dbo].[spCn_RSaldosProvVoucher] @Emp_cCodigo, @Pan_cAnio, @Ase_nVoucherAux
    
    
   END     
      
                    
  tagFinMov:                  
                    
  END                  
                    
  tagExit:                  
                  
                  
END  
GO
