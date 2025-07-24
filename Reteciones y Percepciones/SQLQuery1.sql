USE SAFC_ECB
GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO
/*-----------------------------------------------------------------------------------------------------------------          
MODULO DE CONTABILIDAD         
Modificado por: Berrospi Valladares        
Fecha de Modificación: 02/01/2013          
DESCRIPCION  : Mantenimiento de Cabecera de Asiento          
------------------------------------------------------------------------------------------------------------------*/          
ALTER PROCEDURE [dbo].[spCn_GrabaAsientoCab]          
  @Accion varchar(20)          
, @Ase_cNummov char(10) = ''          
, @Emp_cCodigo char(3) = ''        
, @Pan_cAnio char(4) = ''        
, @Per_cPeriodo char(2) = ''        
, @Lib_cTipoLibro char(2) = ''        
, @Ase_nVoucher char(10) = ''        
, @Ase_dFecha datetime = '01/01/1900 12:00 AM'        
, @Ase_cOperaTC char(3) = ''          
, @Ase_cTipoMoneda char(3) = ''          
, @Ase_nTipoCambio numeric(14,3) = 0        
, @Ase_cGlosa varchar(250) = ''        
, @Ase_cOperaCaja char(1) = ''        
, @Ase_cUserCrea varchar(20) = ''        
, @Ase_cEstado char(1) = ''        
, @Ase_cCuadreManual char(1) = ''        
, @Ase_cCodSoft  char(3)  = ''        
, @Ase_cElimSoft char(1) = '0'        
, @Asd_cEstadoO char(1) = '1'        
, @Asd_cEstadoD char(1) = ''
, @Id_Exoneracion CHAR(10) = NULL
, @Id_Tipo_Renta CHAR(10) = NULL
, @Id_Modalidad CHAR(10) = NULL   
, @CreditoFiscal CHAR(1) = '0'
, @MaterialConstruccion CHAR(1) = '0'     
--WITH ENCRYPTION          
AS        
SET DATEFORMAT DMY    
    
declare @RegPLE as int    
declare @Per_cPeriodo2 char(2)  
    
set @RegPLE = 0    
    
IF @Accion = 'ELIMINAR_DESTINO'    
BEGIN    
 DELETE CND_ASIENTO_VOUCHER WITH(ROWLOCK)          
 WHERE Emp_cCodigo = @Emp_cCodigo AND Pan_cAnio=@Pan_cAnio           
  AND Per_cPeriodo=@Per_cPeriodo AND Lib_cTipoLibro=@Lib_cTipoLibro           
  AND Ase_nVoucher=@Ase_nVoucher AND Asd_cDestino = '1'           
  AND Ase_cNummov = @Ase_cNummov    
END    
    
IF @Accion = 'INSERTAR'          
BEGIN        
  
 if @Per_cPeriodo = '13' or @Per_cPeriodo = '14'  
 begin
  set @Per_cPeriodo2 = '12'  
 end
 else
 begin
  set @Per_cPeriodo2 = @Per_cPeriodo
 end
 
 if @Per_cPeriodo = '00'
 begin
  set @Per_cPeriodo2 = '01'
 end
 else
 begin
  set @Per_cPeriodo2 = @Per_cPeriodo
 end

 select @RegPLE = COUNT(*) from CNT_lIBROSGENERADOS     
 where Emp_cCodigo = @Emp_cCodigo and Pan_cAnio = @Pan_cAnio and Per_cPeriodo = @Per_cPeriodo2 and Lib_cTipoLibro = '03' and Estado ='A'    
    
 if @RegPLE = 1 and @Lib_cTipoLibro in ('07','01','08')  
  set @Asd_cEstadoD = '8'  
    
 INSERT  CNC_ASIENTO_VOUCHER WITH(ROWLOCK)          
  (Emp_cCodigo, Ase_cNummov, Pan_cAnio,Per_cPeriodo,Lib_cTipoLibro,Ase_nVoucher,Ase_dFecha,Ase_cTipoMoneda, Ase_cOperaTC, Ase_nTipoCambio, Ase_cEstado,        
  Ase_cGlosa,Ase_cDeleted,Ase_cUserCrea,Ase_cUserModifica,Ase_dFechaCrea,Ase_dFechaModifica,Ase_cEquipoUser, Ase_cOperaCaja, Ase_cCuadreManual, Ase_cCodSoft, Ase_cElimSoft, Asd_cEstadoO, Asd_cEstadoD, Id_Exoneracion, Id_Tipo_Renta, Id_Modalidad, CreditoFiscal, MaterialConstruccion)        
 VALUES (@Emp_cCodigo, @Ase_cNummov, @Pan_cAnio,@Per_cPeriodo,@Lib_cTipoLibro,@Ase_nVoucher,@Ase_dFecha,@Ase_cTipoMoneda, @Ase_cOperaTC, @Ase_nTipoCambio, @Ase_cEstado,        
  @Ase_cGlosa,'',@Ase_cUserCrea,@Ase_cUserCrea,getdate(),getdate(),host_name(), @Ase_cOperaCaja, @Ase_cCuadreManual , @Ase_cCodSoft, @Ase_cElimSoft, @Asd_cEstadoO, @Asd_cEstadoD, @Id_Exoneracion, @Id_Tipo_Renta, @Id_Modalidad, @CreditoFiscal, @MaterialConstruccion)        
END          
        
IF @Accion = 'EDITAR'          
BEGIN          
  
 if @Per_cPeriodo = '13' or @Per_cPeriodo = '14'  
  set @Per_cPeriodo2 = '12'  

 if @Per_cPeriodo = '00'
  set @Per_cPeriodo2 = '01'  

 if @Per_cPeriodo2 not in ('00','12')
  set @Per_cPeriodo2 = @Per_cPeriodo  
  
 select @RegPLE = COUNT(*) from CNT_lIBROSGENERADOS     
 where Emp_cCodigo = @Emp_cCodigo and Pan_cAnio = @Pan_cAnio and Per_cPeriodo = @Per_cPeriodo2 and Lib_cTipoLibro = '03' and Estado ='A'    
    
 if @RegPLE = 1  and @Lib_cTipoLibro in ('07','01','08')  
  set @Asd_cEstadoD = '9'    
    
 UPDATE CNC_ASIENTO_VOUCHER WITH(ROWLOCK)          
 SET Ase_dFecha = @Ase_dFecha, Ase_cTipoMoneda = @Ase_cTipoMoneda, Ase_nTipoCambio = @Ase_nTipoCambio, Ase_cOperaTC = @Ase_cOperaTC          
  , Ase_cGlosa = @Ase_cGlosa, Ase_cUserModifica = @Ase_cUserCrea, Ase_cEstado = @Ase_cEstado, Ase_cOperaCaja = @Ase_cOperaCaja          
  , Ase_dFechaModifica = GETDATE(), Ase_cEquipoUser = HOST_NAME(), Ase_cCuadreManual = @Ase_cCuadreManual, Asd_cEstadoO  = @Asd_cEstadoO , Asd_cEstadoD = @Asd_cEstadoD,
  Id_Exoneracion = @Id_Exoneracion, Id_Tipo_Renta = @Id_Tipo_Renta, Id_Modalidad = @Id_Modalidad, CreditoFiscal = @CreditoFiscal, MaterialConstruccion = @MaterialConstruccion          
 WHERE Emp_cCodigo = @Emp_cCodigo AND Pan_cAnio = @Pan_cAnio AND Per_cPeriodo = @Per_cPeriodo          
    AND Lib_cTipoLibro = @Lib_cTipoLibro AND Ase_nVoucher = @Ase_nVoucher AND Ase_cDeleted <> '*' AND Ase_cNummov = @Ase_cNummov          
END          
          
IF @Accion = 'ANULAR'          
BEGIN          
 -- ANULAR CABECERA          
 UPDATE CNC_ASIENTO_VOUCHER WITH(ROWLOCK)          
 SET  Ase_cGlosa = 'ANULADO', Ase_cEstado = 'X', Ase_dFechaModifica = GETDATE(), Ase_cEquipoUser = HOST_NAME()          
 WHERE Emp_cCodigo = @Emp_cCodigo AND Pan_cAnio = @Pan_cAnio AND Per_cPeriodo = @Per_cPeriodo          
  AND Lib_cTipoLibro = @Lib_cTipoLibro AND Ase_nVoucher = @Ase_nVoucher AND Ase_cDeleted <>'*' AND Ase_cNummov = @Ase_cNummov          
           
 -- ANULAR EL DETALLE          
 UPDATE CND_ASIENTO_VOUCHER WITH(ROWLOCK)          
 SET Asd_cGlosa = 'ANULADO', Asd_nDebeSoles = 0, Asd_nHaberSoles = 0, Asd_nDebeMonExt = 0, Asd_nHaberMonExt = 0, Asd_nMontoInafecto = 0,  Asd_cEstado = 'X'          
  , Asd_dFechaModifica = GETDATE(), Asd_cEquipoUser = HOST_NAME()          
 WHERE Emp_cCodigo = @Emp_cCodigo AND Pan_cAnio = @Pan_cAnio AND Per_cPeriodo = @Per_cPeriodo          
  AND Lib_cTipoLibro = @Lib_cTipoLibro AND Ase_nVoucher = @Ase_nVoucher AND Asd_cDeleted <> '*' AND Ase_cNummov = @Ase_cNummov          
          
 --- SI HUBIESE CUENTAS DE PROVISION ELIMINARLOS          
 DELETE CND_ASIENTO_PROV          
 WHERE Emp_cCodigo = @Emp_cCodigo AND Pan_cAnio = @Pan_cAnio AND Per_cPeriodo = @Per_cPeriodo          
  AND Lib_cTipoLibro = @Lib_cTipoLibro AND Ase_nVoucher = @Ase_nVoucher AND Ase_cNummov = @Ase_cNummov          
          
 EXEC spCn_EliminaProvisionSobrante @Emp_cCodigo          
END          
          
IF @Accion = 'ELIMINAR'          
BEGIN          
-------------------------------------------------------------------------------------------------------          
 -- MARCANDO COMO ELIMINANDO LA CABECERA          
 UPDATE CNC_ASIENTO_VOUCHER WITH ( ROWLOCK)          
 SET  Ase_cDeleted = '*', Ase_dFechaModifica = GETDATE(), Ase_cEquipoUser = HOST_NAME()          
 WHERE Emp_cCodigo = @Emp_cCodigo AND Pan_cAnio = @Pan_cAnio AND Per_cPeriodo = @Per_cPeriodo          
  AND Lib_cTipoLibro = @Lib_cTipoLibro AND Ase_nVoucher = @Ase_nVoucher AND Ase_cNummov = @Ase_cNummov          
           
 -- MARCANDO COMO ELIMINANDO EL DETALLE          
 UPDATE CND_ASIENTO_VOUCHER WITH ( ROWLOCK)          
 SET Asd_cDeleted = '*', Asd_dFechaModifica = GETDATE(), Asd_cEquipoUser = HOST_NAME()          
 WHERE Emp_cCodigo = @Emp_cCodigo AND Pan_cAnio = @Pan_cAnio AND Per_cPeriodo = @Per_cPeriodo          
  AND Lib_cTipoLibro = @Lib_cTipoLibro AND Ase_nVoucher = @Ase_nVoucher AND Ase_cNummov = @Ase_cNummov          
          
 --- SI HUBIESE CUENTAS DE PROVISION ELIMINARLOS          
 DELETE CND_ASIENTO_PROV           
 WHERE Emp_cCodigo = @Emp_cCodigo AND Pan_cAnio = @Pan_cAnio AND Per_cPeriodo = @Per_cPeriodo          
  AND Lib_cTipoLibro = @Lib_cTipoLibro AND Ase_nVoucher = @Ase_nVoucher AND Ase_cNummov = @Ase_cNummov          
          
 --exec spCn_EliminaProvisionSobrante @Emp_cCodigo          
          
-------------------------------------------------------------------------------------------------------          
 DECLARE @provision char(1)          
 DECLARE @cuenta varchar(12)          
 DECLARE @correl int          
 DECLARE @item int          
 DECLARE @Asd_cProvCanc char(1)          
 DECLARE @Entidad char(5)          
 DECLARE @Tipo char(2)          
 DECLARE @Serie VARCHAR(20)          
 DECLARE @Numero VARCHAR(25)          
          
 DECLARE @PROVNAC NUMERIC (14,3)          
 DECLARE @PROVEXT NUMERIC (14,3)          
 DECLARE @CANCELNAC NUMERIC (14,3)          
 DECLARE @CANCELEXT NUMERIC (14,3)          
           
 DECLARE Saldo_Cursor CURSOR FOR          
  SELECT Ent_cCodEntidad, Pla_cCuentaContable, Asd_cTipoDoc, Asd_cSerieDoc, Asd_cNumDoc, Asd_nCorre, Asd_cProvCanc , Asd_nItem          
  FROM CND_ASIENTO_VOUCHER  WITH(NOLOCK)          
  WHERE Emp_cCodigo = @Emp_cCodigo AND Pan_cAnio = @Pan_cAnio AND Per_cPeriodo = @Per_cPeriodo          
   AND Lib_cTipoLibro = @Lib_cTipoLibro AND Ase_nVoucher = @Ase_nVoucher --AND Asd_cDeleted = 'E'          
   AND Ase_cNummov = @Ase_cNummov AND Asd_cTipoDoc <> '' AND Ent_cCodEntidad <> ''-- AND Asd_nCorre > 0          
  ORDER BY Asd_nItem          
          
 OPEN Saldo_Cursor           
 FETCH NEXT FROM Saldo_Cursor INTO @Entidad, @Cuenta, @Tipo, @Serie, @Numero, @correl, @Asd_cProvCanc, @item          
 WHILE @@FETCH_STATUS = 0          
 BEGIN          
  SET @provision = (SELECT Pla_cProvision FROM CNM_PLAN_CTA  WITH ( NOLOCK)          
  WHERE Emp_cCodigo = @Emp_cCodigo AND Pan_cAnio = @Pan_cAnio AND Pla_cCuentaContable = @cuenta)          
           
  IF @Asd_cProvCanc <> 'P' and @provision = '1' And (@Per_cPeriodo>='00' And @Per_cPeriodo<'13')          
  BEGIN          
           
   SELECT @CANCELNAC = abs(SUM(asd_ndebesoles - asd_nhabersoles)) ,           
    @CANCELEXT = abs(SUM(asd_ndebemonext - asd_nhabermonext))           
   FROM CND_ASIENTO_VOUCHER            
   WHERE Emp_cCodigo = @Emp_cCodigo AND Pan_cAnio = @Pan_cAnio          
    AND Ent_cCodEntidad = @Entidad AND Pla_cCuentaContable = @Cuenta           
    AND Asd_cTipoDoc = @Tipo AND Asd_cSerieDoc = @Serie AND Asd_cNumDoc = @Numero          
    AND Asd_cDeleted <>'*' AND ASD_CPROVCANC <> 'P'          
    AND Ase_nVoucher <> @Ase_nVoucher AND Ase_cNummov <> @Ase_cNummov           
           
      UPDATE CND_ASIENTO_PROV  WITH ( ROWLOCK)          
      SET Cnp_nMonSolCancel = isnull(@CANCELNAC,0)          
       , Cnp_nMonExtCancel =isnull(@CANCELEXT,0)           
      WHERE Emp_cCodigo = @Emp_cCodigo AND Pan_cAnio = @Pan_cAnio AND Cnp_nCorre = @correl          
  END          
  -- SIGUIENTE REGISTRO          
  FETCH NEXT FROM Saldo_Cursor INTO @Entidad, @Cuenta, @Tipo, @Serie, @Numero, @correl, @Asd_cProvCanc, @item          
 END          
 CLOSE Saldo_Cursor          
 DEALLOCATE Saldo_Cursor          
-------------------------------------------------------------------------------------------------------          
END          
          
IF @Accion = 'ELIMINARPROV'          
BEGIN          
 DELETE CND_ASIENTO_PROV WITH(ROWLOCK)          
 WHERE Emp_cCodigo = @Emp_cCodigo AND Pan_cAnio = @Pan_cAnio AND Per_cPeriodo = @Per_cPeriodo          
  AND Lib_cTipoLibro = @Lib_cTipoLibro AND Ase_nVoucher = @Ase_nVoucher AND Ase_cNummov = @Ase_cNummov          
          
 EXEC spCn_EliminaProvisionSobrante @Emp_cCodigo          
END
GO
