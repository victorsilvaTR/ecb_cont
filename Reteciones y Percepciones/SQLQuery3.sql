USE SAFC_ECB
GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO
--spCn_ConsultaAsientos 'SEL_REG', '0000001119','002', '2009', '01', '03', '0301000001', '', '', '', ''      
ALTER PROCEDURE [dbo].[spCn_ConsultaAsientos]               
 @Tipo    varchar(20)= '',               
 @Ase_cNummov  char(10)='',               
 @Emp_cCodigo  char(3)='',              
 @Pan_cAnio   char(4)='',              
 @Per_cPeriodo  char(2)='',              
 @Lib_cTipoLibro char(2)='',              
 @Ase_nVoucher  char(10)='',              
 @desde    varchar(10) = '',              
 @hasta    varchar(10) = '',              
 @Pla_cCuentaContable varchar(12) = '',              
 @Moneda   varchar(3) = ''              
--WITH ENCRYPTION              
AS              
SET DATEFORMAT DMY        
SET NOCOUNT ON      
            
declare @NDecimal int              
set @NDecimal = 3              
              
DECLARE @Mon_cCodMNac CHAR(3)              
DECLARE @vAsd_cTipoDoc char (3)              
DECLARE @vAsd_cSerieDoc char (5)              
DECLARE @vAsd_cNumDoc char (15)              
DECLARE @vAsd_cNumMov char (10)              
DECLARE @NumMov char (10)              
DECLARE @vAse_nVoucher char (10)              
DECLARE @vEnt_cVoucher char (10)              
              
DECLARE @vAsd_nDebe numeric(15, 3)               
DECLARE @vAsd_nHaber numeric(15, 3)               
              
DECLARE @vAse_cRuc char (15)              
DECLARE @vAse_cRazonSocial char (250)              
DECLARE @vAsd_nCorre int              
              
DECLARE @vEnt_cCodEntidad char (5)              
              
DECLARE @vPer_cPeriodo char (20)              
DECLARE @vPla_cCuentaContable char (12)              
DECLARE @vAsd_dFecDoc datetime              
DECLARE @Registros int              
DECLARE @vLib_cTrans char (3)              
DECLARE @vLib_cDescripcionTrans char (250)              
DECLARE @vAsd_cTipoDocref char (3)              
DECLARE @vTdo_cNombreLargo char (250)              
              
DECLARE @Lib_Compras char (2)              
DECLARE @Lib_Ventas char (2)              
DECLARE @Lib_Diario char (2)              
DECLARE @Lib_Caja char (2)              
DECLARE @Lib_CajaIng char (2)              
DECLARE @Lib_CajaEgr char (2)              
--DECLARE @TipoBoletas char (2)              
DECLARE @UIT numeric(15, 3)               
              
              
DECLARE @vvAse_nNumMov char (10)              
DECLARE @vvAse_nVoucher char (10)              
DECLARE @vvAsd_nDebe numeric(15, 3)               
DECLARE @vvAsd_nHaber numeric(15, 3)               
DECLARE @vvAsd_nFecDoc datetime              
              
              
DECLARE @vLetraCob char (2)              
DECLARE @vLetraPag char (2)              
DECLARE @vNC char (2)              
DECLARE @vND char (2)              
              
DECLARE @CTA_GANANCIA VARCHAR(12)              
DECLARE @CTA_PERDIDA VARCHAR(12)              
              
declare @Aux numeric(14,3)              
-------------------------------------------------------------------              
SELECT @CTA_GANANCIA = [Pla_cCuentaContable]              
FROM CNM_PLAN_CTA WITH(READUNCOMMITTED) WHERE Emp_cCodigo = @Emp_cCodigo AND Pan_cAnio = @Pan_cAnio AND pla_cDifCambio = 'G'              
              
SELECT @CTA_PERDIDA = [Pla_cCuentaContable]              
FROM CNM_PLAN_CTA WITH(READUNCOMMITTED) WHERE Emp_cCodigo = @Emp_cCodigo AND Pan_cAnio = @Pan_cAnio AND pla_cDifCambio = 'P'              
-------------------------------------------------------------------              
              
SET @vNC = DBO.fBuscaConfOP (@Emp_cCodigo,@Pan_cAnio,'012')              
SET @vND = DBO.fBuscaConfOP (@Emp_cCodigo,@Pan_cAnio,'013')              
              
SET @vLetraCob = DBO.fBuscaConfOP (@Emp_cCodigo,@Pan_cAnio,'022')              
SET @vLetraPag = DBO.fBuscaConfOP (@Emp_cCodigo,@Pan_cAnio,'023')              
              
--SET @TipoBoletas = DBO.fBuscaConfOP (@Emp_cCodigo,@Pan_cAnio,'028')              
SET @UIT = DBO.fBuscaConfOP (@Emp_cCodigo,@Pan_cAnio,'027')              
              
-------------------------------------------------------------------            
SELECT @Lib_Compras = Cfl_cCompras, @Lib_Ventas = Cfl_cVentas , @Lib_diario = Cfl_cDiario, @Lib_Caja = Cfl_cCaja, @Lib_CajaIng = Cfl_cCajaIngresos, @Lib_CajaEgr = Cfl_cCajaEgresos              
FROM CNT_CONFIG_LIBROS WITH(READUNCOMMITTED) WHERE Emp_cCodigo=@Emp_cCodigo              
-------------------------------------------------------------------              
SELECT @Mon_cCodMNac=Mon_cCodigo FROM CNT_TIPO_MONEDA WITH(READUNCOMMITTED) WHERE Emp_cCodigo = @Emp_cCodigo AND Mon_cMNac = '1'              
              
set @UIT = @UIT /2              
              
-- spCn_ConsultaAsientos 'BUSCAVOUCHER', '','025', '2008', '04', '05', '0504000005'              
IF @Tipo = 'BUSCAVOUCHER'               
BEGIN-- *** SELECCIONAR LOS DATOS DE UN ASIENTO CONTABLE              
 SELECT A.Ase_cNummov, A.Per_cPeriodo, A.Ase_nVoucher               
 FROM   CNC_ASIENTO_VOUCHER A WITH(READUNCOMMITTED)            
 WHERE A.Emp_cCodigo = @Emp_cCodigo AND A.Pan_cAnio = @Pan_cAnio  AND A.Per_cPeriodo = @Per_cPeriodo AND A.Ase_nVoucher = @Ase_nVoucher AND A.Ase_cDeleted <> '*'               
END            
              
-------------------------------------------------------------------              
IF @Tipo = 'ELIMINACIERRE'               
BEGIN              
          
 DELETE FROM CND_ASIENTO_VOUCHER  WHERE               
 Emp_cCodigo = @Emp_cCodigo AND Pan_cAnio = @Pan_cAnio AND Per_cPeriodo = @Per_cPeriodo AND Lib_cTipoLibro = @Lib_cTipoLibro               
               
 DELETE FROM CNC_ASIENTO_VOUCHER WHERE               
 Emp_cCodigo = @Emp_cCodigo AND Pan_cAnio = @Pan_cAnio AND Per_cPeriodo = @Per_cPeriodo AND Lib_cTipoLibro = @Lib_cTipoLibro               
               
END              
              
IF @Tipo = 'SEL_REG'               
     Begin -- *** SELECCIONAR LOS DATOS DE UN ASIENTO CONTABLE              
 SELECT A.Ase_cNummov, A.Per_cPeriodo, A.Ase_dFecha, A.Lib_cTipoLibro,               
  I.Lib_cDescripcion, A.Ase_nVoucher, A.Ase_cTipoMoneda,j.Mon_cNombreCorto, A.Ase_cOperaCaja,               
         J.Mon_cNombreLargo, A.Ase_nTipoCambio, A.Ase_cOperaTC, C.Asd_cTipoMoneda, L.Mon_cNombreCorto As NomMoneCorto, L.Mon_cMNac As FlgMonNac,               
  A.Ase_cGlosa, C.Asd_nItem, C.Pla_cCuentaContable, G.Pla_cNombreCuenta,               
         G.Pla_cProvision, G.Pla_cCentroCosto, G.Pla_cDifCambio, G.Pla_cRedondeo, G.Pla_cDocumento, G.Ten_cTipoEntidad As TipoEntiCta,              
  C.Asd_cGlosa,               
              
--  round(C.Asd_nDebeSoles,@NDecimal) as 'Asd_nDebeSoles',               
--  round(C.Asd_nDebeMonExt,@NDecimal) as 'Asd_nDebeMonExt',               
--  round(C.Asd_nHaberSoles,@NDecimal) as 'Asd_nHaberSoles',               
--  round(C.Asd_nHaberMonExt,@NDecimal) as 'Asd_nHaberMonExt',               
              
  C.Asd_nDebeSoles,              
  C.Asd_nDebeMonExt,               
  C.Asd_nHaberSoles,               
  C.Asd_nHaberMonExt,               
              
  C.Asd_nTipoCambio, C.Asd_cOperaTC, C.Cos_cCodigo,               
         H.Cos_cDescripcion, C.Ten_cTipoEntidad, C.Ent_cCodEntidad,               
   CASE K.Ten_cPlame     
  WHEN '1' THEN F.Ent_cApaterno+' '+F.Ent_cAmaterno+' '+F.Ent_cNombres    
  ELSE  F.Ent_cPersona    
  END as 'Ent_cPersona'  
    , K.Ten_cNombreEntidad, C.Asd_cTipoDoc,               
         C.Asd_cSerieDoc, C.Asd_cNumDoc, C.Asd_dFecDoc,               
              C.Asd_cTipoDocRef, C.Asd_cSerieDocRef, C.Asd_cNumDocRef,               
              C.Asd_dFecDocRef, C.Asd_nMontoInafecto, C.Asd_cRetencion, C.Asd_cFlgSpot,              
              C.Asd_dFechaSpot, C.Asd_cNumSpot, C.Asd_cDestino,               
              C.Asd_nCorre, Asd_cProvCanc,   C.Asd_cMonedaCalculo, C.Imp_nPorcentaje,               
              E.Tdo_cNombreLargo,E.Tdo_cCodigo ,E.Tdo_cNombreCorto, D.Tdo_cNombreLargo AS Tdo_cNombreLargoRef,               
              D.Tdo_cNombreCorto AS Tdo_cNombreCortoRef, Ase_cUserModifica,               
              B.Emp_cNombreLargo, B.Emp_cNombreCorto, GETDATE() AS FechaHoraImp, C.Com_cTipoIgv, C.Asd_dFecVen,               
       G.Pla_cTipoAfect, C.Tra_cCodigo, C.Asd_cFormaPago, C.Asd_cBaseImp,              
       G.pla_cdetraccion, G.pla_cretencion, G.pla_cpercepcion, G.pla_cNCND,              
       C.Asd_cMonAdic, C.Asd_cImpAdic , c.Asd_cComprobante, c.Asd_cProceso,               
       c.ECP_COPERACION, c.Asd_cRegAux, c.Asd_cRegAuxDet, c.Asd_cManual  ,               
              
       C.Asd_cGrupo , C.Asd_cCodConcepto, CPTO.Asl_cDescripcion, A.Asd_cEstadoO, A.Asd_cEstadoD, ISNULL(C.Id_Exoneracion, '') AS Id_Exoneracion, 
       ISNULL(C.Id_Tipo_Renta, '') AS Id_Tipo_Renta, ISNULL(C.Id_Modalidad, '') AS Id_Modalidad, ISNULL(C.Id_Aduana, '') AS Id_Aduana, ISNULL(C.Id_Clasific_Servicio, '') AS Id_Clasific_Servicio,
       ISNULL(CreditoFiscal, 0) AS CreditoFiscal, ISNULL(MaterialConstruccion, 0) AS MaterialConstruccion                
              
 FROM                 
  EMPRESA B WITH(READUNCOMMITTED) INNER JOIN CNC_ASIENTO_VOUCHER A WITH(READUNCOMMITTED) ON                
  B.Emp_cCodigo = A.Emp_cCodigo               
  LEFT JOIN CND_ASIENTO_VOUCHER C WITH(READUNCOMMITTED) ON                
  A.Emp_cCodigo = C.Emp_cCodigo AND A.Pan_cAnio = C.Pan_cAnio AND A.Per_cPeriodo = C.Per_cPeriodo AND A.Lib_cTipoLibro = C.Lib_cTipoLibro AND               
  A.Ase_nVoucher = C.Ase_nVoucher and a.Ase_cNummov = c.Ase_cNummov              
  LEFT JOIN CNT_TIPODOC D WITH(READUNCOMMITTED) ON              
  C.Asd_cTipoDocRef = D.Tdo_cCodigo AND C.Emp_cCodigo = D.Emp_cCodigo               
  LEFT JOIN CNT_TIPODOC E WITH(READUNCOMMITTED) ON               
  C.Asd_cTipoDoc = E.Tdo_cCodigo AND C.Emp_cCodigo = E.Emp_cCodigo               
     LEFT JOIN CNM_ENTIDAD F WITH(READUNCOMMITTED) ON               
     C.Emp_cCodigo = F.Emp_cCodigo AND C.Ten_cTipoEntidad = F.Ten_cTipoEntidad AND               
  C.Ent_cCodEntidad = F.Ent_cCodEntidad               
  LEFT JOIN CNM_PLAN_CTA G WITH(READUNCOMMITTED) ON               
  C.Emp_cCodigo = G.Emp_cCodigo AND C.Pan_cAnio = G.Pan_cAnio AND C.Pla_cCuentaContable = G.Pla_cCuentaContable               
  LEFT JOIN CNT_CENTRO_COSTO H WITH(READUNCOMMITTED) ON                
  C.Emp_cCodigo = H.Emp_cCodigo AND C.Pan_cAnio = H.Pan_cAnio AND C.Cos_cCodigo = H.Cos_cCodigo               
    LEFT JOIN CNT_LIBRO_OPERA I WITH(READUNCOMMITTED) ON               
  A.Lib_cTipoLibro = I.Lib_cTipoLibro AND A.PAN_CANIO = I.PAN_CANIO AND A.Emp_cCodigo = I.Emp_cCodigo               
   LEFT JOIN .CNT_TIPO_MONEDA J WITH(READUNCOMMITTED) ON               
   A.Ase_cTipoMoneda = J.Mon_cCodigo AND A.Emp_cCodigo = J.Emp_cCodigo                
  LEFT JOIN CNT_ENTIDAD K WITH(READUNCOMMITTED) ON               
  F.Ten_cTipoEntidad = K.Ten_cTipoEntidad AND F.Emp_cCodigo = K.Emp_cCodigo                
   LEFT JOIN CNT_TIPO_MONEDA L WITH(READUNCOMMITTED) ON               
   C.Asd_cTipoMoneda = L.Mon_cCodigo AND C.Emp_cCodigo = L.Emp_cCodigo                
  LEFT JOIN CNT_CONCEPTO_LIBRO CPTO WITH(READUNCOMMITTED) ON              
  C.Emp_cCodigo = CPTO.Emp_cCodigo AND C.Pan_cAnio = CPTO.Pan_cAnio AND C.Lib_cTipoLibro  = CPTO.Lib_cTipoLibro and C.Asd_cCodConcepto = CPTO.Asl_cCodigo                
              
 WHERE                   
  (A.Emp_cCodigo = @Emp_cCodigo)               
  AND A.Ase_cNummov = @Ase_cNummov              
  AND (A.Pan_cAnio = @Pan_cAnio)               
  AND (A.Per_cPeriodo = @Per_cPeriodo)               
  AND (A.Lib_cTipoLibro = @Lib_cTipoLibro)               
  AND (A.Ase_nVoucher = @Ase_nVoucher)               
  AND (A.Ase_cDeleted <> '*')               
  AND (C.Asd_cDeleted <> '*')               
  ORDER BY Asd_nItem              
     End              
              
IF @Tipo = 'SEL_REG_EDIT'                
     Begin -- *** SELECCIONAR LOS DATOS DE UN ASIENTO CONTABLE              
-- spCn_ConsultaAsientos 'SEL_REG_EDIT', '','004', '2006', '01', '03', '0301000001', '', '', '', ''              
 SELECT                
  A.Ase_cNummov, A.Per_cPeriodo, A.Ase_dFecha, A.Lib_cTipoLibro,               
   I.Lib_cDescripcion, A.Ase_nVoucher, A.Ase_cTipoMoneda,j.Mon_cNombreCorto, A.Ase_cOperaCaja,               
               J.Mon_cNombreLargo, A.Ase_nTipoCambio, A.Ase_cOperaTC, C.Asd_cTipoMoneda, L.Mon_cNombreCorto As NomMoneCorto, L.Mon_cMNac As FlgMonNac,               
  A.Ase_cGlosa, C.Asd_nItem, C.Pla_cCuentaContable, G.Pla_cNombreCuenta,               
              G.Pla_cProvision, G.Pla_cCentroCosto, G.Pla_cDifCambio, G.Pla_cRedondeo, G.Pla_cDocumento, G.Ten_cTipoEntidad As TipoEntiCta,              
  C.Asd_cGlosa, C.Asd_nDebeSoles, C.Asd_nDebeMonExt, C.Asd_nHaberSoles,               
              C.Asd_nHaberMonExt, C.Asd_nTipoCambio, C.Asd_cOperaTC, C.Cos_cCodigo,               
              H.Cos_cDescripcion, C.Ten_cTipoEntidad, C.Ent_cCodEntidad,  
    CASE K.Ten_cPlame     
    WHEN '1' THEN F.Ent_cApaterno+' '+F.Ent_cAmaterno+' '+F.Ent_cNombres    
    ELSE  F.Ent_cPersona    
    END as 'Ent_cPersona',                
               K.Ten_cNombreEntidad, C.Asd_cTipoDoc,               
              C.Asd_cSerieDoc, C.Asd_cNumDoc, C.Asd_dFecDoc,               
              C.Asd_cTipoDocRef, C.Asd_cSerieDocRef, C.Asd_cNumDocRef,               
              C.Asd_dFecDocRef, C.Asd_nMontoInafecto, C.Asd_cRetencion, C.Asd_cFlgSpot,              
              C.Asd_dFechaSpot, C.Asd_cNumSpot, C.Asd_cDestino,               
              C.Asd_nCorre, Asd_cProvCanc,   C.Asd_cMonedaCalculo, C.Imp_nPorcentaje,               
              E.Tdo_cNombreLargo,E.Tdo_cCodigo ,E.Tdo_cNombreCorto, D.Tdo_cNombreLargo AS Tdo_cNombreLargoRef,               
              D.Tdo_cNombreCorto AS Tdo_cNombreCortoRef, Ase_cUserModifica,               
              B.Emp_cNombreLargo, B.Emp_cNombreCorto, GETDATE() AS FechaHoraImp, C.Com_cTipoIgv, C.Asd_dFecVen,               
       G.Pla_cTipoAfect, C.Tra_cCodigo, C.Asd_cFormaPago, C.Asd_cBaseImp,              
       G.pla_cdetraccion, G.pla_cretencion, G.pla_cpercepcion, G.pla_cNCND,              
       C.Asd_cMonAdic, C.Asd_cImpAdic, c.Asd_cComprobante, c.Asd_cProceso,               
       c.ECP_COPERACION, c.Asd_cRegAux, c.Asd_cRegAuxDet, c.Asd_cManual  ,               
              
       C.Asd_cGrupo ,C.Asd_cCodConcepto, CPTO.Asl_cDescripcion, A.Asd_cEstadoO, A.Asd_cEstadoD , ISNULL(c.Id_Exoneracion, '') AS Id_Exoneracion, 
       ISNULL(c.Id_Tipo_Renta, '') AS Id_Tipo_Renta, ISNULL(c.Id_Modalidad, '') AS Id_Modalidad, ISNULL(c.Id_Aduana, '') AS Id_Aduana,
       ISNULL(c.Id_Clasific_Servicio, '') AS Id_Clasific_Servicio,
       ISNULL(CreditoFiscal, 0) AS CreditoFiscal, ISNULL(MaterialConstruccion, 0) AS MaterialConstruccion              
       FROM   EMPRESA B WITH(READUNCOMMITTED) INNER JOIN CNC_ASIENTO_VOUCHER A WITH(READUNCOMMITTED) ON  B.Emp_cCodigo = A.Emp_cCodigo               
  LEFT  JOIN CND_ASIENTO_VOUCHER C WITH(READUNCOMMITTED) ON                
    A.Emp_cCodigo = C.Emp_cCodigo AND A.Pan_cAnio = C.Pan_cAnio AND A.Per_cPeriodo = C.Per_cPeriodo AND A.Lib_cTipoLibro = C.Lib_cTipoLibro AND               
    A.Ase_nVoucher = C.Ase_nVoucher and a.Ase_cNummov = c.Ase_cNummov              
  LEFT JOIN CNT_TIPODOC  D WITH(READUNCOMMITTED) ON  C.Asd_cTipoDocRef = D.Tdo_cCodigo AND C.Emp_cCodigo = D.Emp_cCodigo               
  LEFT JOIN CNT_TIPODOC E WITH(READUNCOMMITTED) ON C.Asd_cTipoDoc = E.Tdo_cCodigo AND C.Emp_cCodigo = E.Emp_cCodigo               
     LEFT JOIN CNM_ENTIDAD F WITH(READUNCOMMITTED) ON C.Emp_cCodigo = F.Emp_cCodigo AND C.Ten_cTipoEntidad = F.Ten_cTipoEntidad AND               
    C.Ent_cCodEntidad = F.Ent_cCodEntidad               
  LEFT JOIN CNM_PLAN_CTA G WITH(READUNCOMMITTED) ON C.Emp_cCodigo = G.Emp_cCodigo AND C.Pan_cAnio = G.Pan_cAnio AND C.Pla_cCuentaContable = G.Pla_cCuentaContable               
  LEFT JOIN CNT_CENTRO_COSTO H WITH(READUNCOMMITTED) ON  C.Emp_cCodigo = H.Emp_cCodigo AND C.Pan_cAnio = H.Pan_cAnio AND C.Cos_cCodigo = H.Cos_cCodigo               
    LEFT JOIN CNT_LIBRO_OPERA I WITH(READUNCOMMITTED) ON               
   A.Lib_cTipoLibro = I.Lib_cTipoLibro AND A.PAN_CANIO = I.PAN_CANIO AND A.Emp_cCodigo = I.Emp_cCodigo               
   LEFT JOIN .CNT_TIPO_MONEDA J WITH(READUNCOMMITTED) ON A.Ase_cTipoMoneda = J.Mon_cCodigo AND A.Emp_cCodigo = J.Emp_cCodigo                
  LEFT JOIN CNT_ENTIDAD K WITH(READUNCOMMITTED) ON F.Ten_cTipoEntidad = K.Ten_cTipoEntidad AND F.Emp_cCodigo = K.Emp_cCodigo                
   LEFT JOIN CNT_TIPO_MONEDA L WITH(READUNCOMMITTED) ON C.Asd_cTipoMoneda = L.Mon_cCodigo AND C.Emp_cCodigo = L.Emp_cCodigo                
  LEFT JOIN CNT_CONCEPTO_LIBRO CPTO WITH(READUNCOMMITTED) ON              
  C.Emp_cCodigo = CPTO.Emp_cCodigo AND C.Pan_cAnio = CPTO.Pan_cAnio AND C.Lib_cTipoLibro  = CPTO.Lib_cTipoLibro and C.Asd_cCodConcepto = CPTO.Asl_cCodigo                
                 
               
 WHERE                   
  (A.Emp_cCodigo = @Emp_cCodigo)               
  AND A.Ase_cNummov = @Ase_cNummov              
  AND (A.Pan_cAnio = @Pan_cAnio)               
  AND (A.Per_cPeriodo = @Per_cPeriodo)               
  AND (A.Lib_cTipoLibro = @Lib_cTipoLibro)               
  AND (A.Ase_nVoucher = @Ase_nVoucher)               
              
              
  AND ISNULL(a.Ase_cDeleted,'') <> '*'               
  AND ISNULL(c.Asd_cDeleted,'') <> '*'               
  AND ISNULL(C.Asd_cDestino,'') <> '1'               
  ORDER BY Asd_nItem              
     End              
-- spCn_ConsultaAsientos 'SEL_REG_EDIT', '0000000140','001', '2008', '01', '05', '0501000002', '', '', '', ''              
-- spCn_ConsultaAsientos 'SEL_REG_EDIT_NC', '0000000139','001', '2008', '01', '05', '0501000001', '', '', '', ''              
IF @Tipo = 'SEL_REG_EDIT_NC'              
     Begin        
		IF @Lib_cTipoLibro <> '01'
		BEGIN
			SELECT                
			  A.Ase_cNummov, A.Per_cPeriodo, A.Ase_dFecha, A.Lib_cTipoLibro,               
			   I.Lib_cDescripcion, A.Ase_nVoucher, A.Ase_cTipoMoneda,j.Mon_cNombreCorto, A.Ase_cOperaCaja,               
						   J.Mon_cNombreLargo, A.Ase_nTipoCambio, A.Ase_cOperaTC, C.Asd_cTipoMoneda, L.Mon_cNombreCorto As NomMoneCorto, L.Mon_cMNac As FlgMonNac,               
			  A.Ase_cGlosa, C.Asd_nItem, C.Pla_cCuentaContable, G.Pla_cNombreCuenta,               
						  G.Pla_cProvision, G.Pla_cCentroCosto, G.Pla_cDifCambio, G.Pla_cRedondeo, G.Pla_cDocumento, G.Ten_cTipoEntidad As TipoEntiCta,              
			  C.Asd_cGlosa, C.Asd_nDebeSoles, C.Asd_nDebeMonExt, C.Asd_nHaberSoles,               
						  C.Asd_nHaberMonExt, C.Asd_nTipoCambio, C.Asd_cOperaTC, C.Cos_cCodigo,               
						  H.Cos_cDescripcion, C.Ten_cTipoEntidad, C.Ent_cCodEntidad,               
			   CASE K.Ten_cPlame     
			   WHEN '1' THEN F.Ent_cApaterno+' '+F.Ent_cAmaterno+' '+F.Ent_cNombres    
			   ELSE  F.Ent_cPersona    
			   END as 'Ent_cPersona',    
				 K.Ten_cNombreEntidad, C.Asd_cTipoDoc,               
						  C.Asd_cSerieDoc, C.Asd_cNumDoc, C.Asd_dFecDoc,               
						  C.Asd_cTipoDocRef, C.Asd_cSerieDocRef, C.Asd_cNumDocRef,               
						  C.Asd_dFecDocRef, C.Asd_nMontoInafecto, C.Asd_cRetencion, C.Asd_cFlgSpot,              
						  C.Asd_dFechaSpot, C.Asd_cNumSpot, C.Asd_cDestino,               
						  C.Asd_nCorre, Asd_cProvCanc,   C.Asd_cMonedaCalculo, C.Imp_nPorcentaje,               
						  E.Tdo_cNombreLargo,E.Tdo_cCodigo ,E.Tdo_cNombreCorto, D.Tdo_cNombreLargo AS Tdo_cNombreLargoRef,               
						  D.Tdo_cNombreCorto AS Tdo_cNombreCortoRef, Ase_cUserModifica,               
						  B.Emp_cNombreLargo, B.Emp_cNombreCorto, GETDATE() AS FechaHoraImp, C.Com_cTipoIgv, C.Asd_dFecVen,               
				   G.Pla_cTipoAfect, C.Tra_cCodigo, C.Asd_cFormaPago, C.Asd_cBaseImp,              
				   G.pla_cdetraccion, G.pla_cretencion, G.pla_cpercepcion, G.pla_cNCND,              
				   C.Asd_cMonAdic, C.Asd_cImpAdic, c.Asd_cComprobante, c.Asd_cProceso,               
				   c.ECP_COPERACION, c.Asd_cRegAux, c.Asd_cRegAuxDet, c.Asd_cManual  ,               
				   C.Asd_cGrupo , C.Asd_cCodConcepto, CPTO.Asl_cDescripcion, A.Asd_cEstadoO, A.Asd_cEstadoD, ISNULL(c.Id_Exoneracion, '') AS Id_Exoneracion, 
				   ISNULL(c.Id_Tipo_Renta, '') AS Id_Tipo_Renta, ISNULL(c.Id_Modalidad, '') AS Id_Modalidad, ISNULL(c.Id_Aduana, '') AS Id_Aduana,
				   ISNULL(c.Id_Clasific_Servicio, '') AS Id_Clasific_Servicio, 
			       ISNULL(CreditoFiscal, 0) AS CreditoFiscal, ISNULL(MaterialConstruccion, 0) AS MaterialConstruccion
			 FROM   EMPRESA B WITH(READUNCOMMITTED) INNER JOIN CNC_ASIENTO_VOUCHER A WITH(READUNCOMMITTED) ON  B.Emp_cCodigo = A.Emp_cCodigo               
			  LEFT  JOIN CND_ASIENTO_VOUCHER  C WITH(READUNCOMMITTED) ON                
				A.Emp_cCodigo = C.Emp_cCodigo AND A.Pan_cAnio = C.Pan_cAnio AND A.Per_cPeriodo = C.Per_cPeriodo AND A.Lib_cTipoLibro = C.Lib_cTipoLibro AND               
				A.Ase_nVoucher = C.Ase_nVoucher and a.Ase_cNummov = c.Ase_cNummov              
			  LEFT JOIN CNT_TIPODOC  D WITH(READUNCOMMITTED) ON  C.Asd_cTipoDocRef = D.Tdo_cCodigo AND C.Emp_cCodigo = D.Emp_cCodigo               
			  LEFT JOIN CNT_TIPODOC E WITH(READUNCOMMITTED) ON C.Asd_cTipoDoc = E.Tdo_cCodigo AND C.Emp_cCodigo = E.Emp_cCodigo               
				 LEFT JOIN CNM_ENTIDAD F WITH(READUNCOMMITTED) ON C.Emp_cCodigo = F.Emp_cCodigo AND C.Ten_cTipoEntidad = F.Ten_cTipoEntidad AND               
				C.Ent_cCodEntidad = F.Ent_cCodEntidad               
			  LEFT JOIN CNM_PLAN_CTA G WITH(READUNCOMMITTED) ON C.Emp_cCodigo = G.Emp_cCodigo AND C.Pan_cAnio = G.Pan_cAnio AND C.Pla_cCuentaContable = G.Pla_cCuentaContable               
			  LEFT JOIN CNT_CENTRO_COSTO H WITH(READUNCOMMITTED) ON  C.Emp_cCodigo = H.Emp_cCodigo AND C.Pan_cAnio = H.Pan_cAnio AND C.Cos_cCodigo = H.Cos_cCodigo               
				LEFT JOIN CNT_LIBRO_OPERA I WITH(READUNCOMMITTED) ON               
			   A.Lib_cTipoLibro = I.Lib_cTipoLibro AND A.PAN_CANIO = I.PAN_CANIO AND A.Emp_cCodigo = I.Emp_cCodigo               
			   LEFT JOIN .CNT_TIPO_MONEDA J WITH(READUNCOMMITTED) ON A.Ase_cTipoMoneda = J.Mon_cCodigo AND A.Emp_cCodigo = J.Emp_cCodigo                
			  LEFT JOIN CNT_ENTIDAD K WITH(READUNCOMMITTED) ON F.Ten_cTipoEntidad = K.Ten_cTipoEntidad AND F.Emp_cCodigo = K.Emp_cCodigo                
			   LEFT JOIN CNT_TIPO_MONEDA L WITH(READUNCOMMITTED) ON C.Asd_cTipoMoneda = L.Mon_cCodigo AND C.Emp_cCodigo = L.Emp_cCodigo                
			               
			  LEFT JOIN CNT_CONCEPTO_LIBRO CPTO WITH(READUNCOMMITTED) ON              
			  C.Emp_cCodigo = CPTO.Emp_cCodigo AND C.Pan_cAnio = CPTO.Pan_cAnio AND C.Lib_cTipoLibro  = CPTO.Lib_cTipoLibro and C.Asd_cCodConcepto = CPTO.Asl_cCodigo                
			WHERE                   
			  (A.Emp_cCodigo = @Emp_cCodigo)               
			  AND A.Ase_cNummov = @Ase_cNummov              
			  AND (A.Pan_cAnio = @Pan_cAnio)               
			  AND (A.Per_cPeriodo = @Per_cPeriodo)               
			  AND (A.Lib_cTipoLibro = @Lib_cTipoLibro)               
			  AND (A.Ase_nVoucher = @Ase_nVoucher)               
			  AND a.Ase_cDeleted <> '*'               
			  AND c.Asd_cDeleted <> '*'               
			  AND (C.Asd_cDestino <> '1')               
			  AND C.Pla_cCuentaContable <> @CTA_GANANCIA              
			  AND C.Pla_cCuentaContable <> @CTA_PERDIDA              
			  ORDER BY Asd_nItem
			END
		ELSE
			BEGIN
				SELECT                
				  A.Ase_cNummov, A.Per_cPeriodo, A.Ase_dFecha, A.Lib_cTipoLibro,               
				   I.Lib_cDescripcion, A.Ase_nVoucher, A.Ase_cTipoMoneda,j.Mon_cNombreCorto, A.Ase_cOperaCaja,               
							   J.Mon_cNombreLargo, A.Ase_nTipoCambio, A.Ase_cOperaTC, C.Asd_cTipoMoneda, L.Mon_cNombreCorto As NomMoneCorto, L.Mon_cMNac As FlgMonNac,               
				  A.Ase_cGlosa, C.Asd_nItem, C.Pla_cCuentaContable, G.Pla_cNombreCuenta,               
							  G.Pla_cProvision, G.Pla_cCentroCosto, G.Pla_cDifCambio, G.Pla_cRedondeo, G.Pla_cDocumento, G.Ten_cTipoEntidad As TipoEntiCta,              
				  C.Asd_cGlosa, C.Asd_nDebeSoles, C.Asd_nDebeMonExt, C.Asd_nHaberSoles,               
							  C.Asd_nHaberMonExt, C.Asd_nTipoCambio, C.Asd_cOperaTC, C.Cos_cCodigo,               
							  H.Cos_cDescripcion, C.Ten_cTipoEntidad, C.Ent_cCodEntidad,               
				   CASE K.Ten_cPlame     
				   WHEN '1' THEN F.Ent_cApaterno+' '+F.Ent_cAmaterno+' '+F.Ent_cNombres    
				   ELSE  F.Ent_cPersona    
				   END as 'Ent_cPersona',    
					 K.Ten_cNombreEntidad, C.Asd_cTipoDoc,               
							  C.Asd_cSerieDoc, C.Asd_cNumDoc, C.Asd_dFecDoc,               
							  C.Asd_cTipoDocRef, C.Asd_cSerieDocRef, C.Asd_cNumDocRef,               
							  C.Asd_dFecDocRef, C.Asd_nMontoInafecto, C.Asd_cRetencion, C.Asd_cFlgSpot,              
							  C.Asd_dFechaSpot, C.Asd_cNumSpot, C.Asd_cDestino,               
							  C.Asd_nCorre, Asd_cProvCanc,   C.Asd_cMonedaCalculo, C.Imp_nPorcentaje,               
							  E.Tdo_cNombreLargo,E.Tdo_cCodigo ,E.Tdo_cNombreCorto, D.Tdo_cNombreLargo AS Tdo_cNombreLargoRef,               
							  D.Tdo_cNombreCorto AS Tdo_cNombreCortoRef, Ase_cUserModifica,               
							  B.Emp_cNombreLargo, B.Emp_cNombreCorto, GETDATE() AS FechaHoraImp, C.Com_cTipoIgv, C.Asd_dFecVen,               
					   G.Pla_cTipoAfect, C.Tra_cCodigo, C.Asd_cFormaPago, C.Asd_cBaseImp,              
					   G.pla_cdetraccion, G.pla_cretencion, G.pla_cpercepcion, G.pla_cNCND,              
					   C.Asd_cMonAdic, C.Asd_cImpAdic, c.Asd_cComprobante, c.Asd_cProceso,               
					   c.ECP_COPERACION, c.Asd_cRegAux, c.Asd_cRegAuxDet, c.Asd_cManual  ,               
					   C.Asd_cGrupo , C.Asd_cCodConcepto, CPTO.Asl_cDescripcion, A.Asd_cEstadoO, A.Asd_cEstadoD, ISNULL(c.Id_Exoneracion, '') AS Id_Exoneracion, 
					   ISNULL(c.Id_Tipo_Renta, '') AS Id_Tipo_Renta, ISNULL(c.Id_Modalidad, '') AS Id_Modalidad, ISNULL(c.Id_Aduana, '') AS Id_Aduana,
					   ISNULL(c.Id_Clasific_Servicio, '') AS Id_Clasific_Servicio,
					   ISNULL(CreditoFiscal, 0) AS CreditoFiscal, ISNULL(MaterialConstruccion, 0) AS MaterialConstruccion 
				              
				 FROM   EMPRESA B WITH(READUNCOMMITTED) INNER JOIN CNC_ASIENTO_VOUCHER A WITH(READUNCOMMITTED) ON  B.Emp_cCodigo = A.Emp_cCodigo               
				  LEFT  JOIN CND_ASIENTO_VOUCHER  C WITH(READUNCOMMITTED) ON                
					A.Emp_cCodigo = C.Emp_cCodigo AND A.Pan_cAnio = C.Pan_cAnio AND A.Per_cPeriodo = C.Per_cPeriodo AND A.Lib_cTipoLibro = C.Lib_cTipoLibro AND               
					A.Ase_nVoucher = C.Ase_nVoucher and a.Ase_cNummov = c.Ase_cNummov              
				  LEFT JOIN CNT_TIPODOC  D WITH(READUNCOMMITTED) ON  C.Asd_cTipoDocRef = D.Tdo_cCodigo AND C.Emp_cCodigo = D.Emp_cCodigo               
				  LEFT JOIN CNT_TIPODOC E WITH(READUNCOMMITTED) ON C.Asd_cTipoDoc = E.Tdo_cCodigo AND C.Emp_cCodigo = E.Emp_cCodigo               
					 LEFT JOIN CNM_ENTIDAD F WITH(READUNCOMMITTED) ON C.Emp_cCodigo = F.Emp_cCodigo AND C.Ten_cTipoEntidad = F.Ten_cTipoEntidad AND               
					C.Ent_cCodEntidad = F.Ent_cCodEntidad               
				  LEFT JOIN CNM_PLAN_CTA G WITH(READUNCOMMITTED) ON C.Emp_cCodigo = G.Emp_cCodigo AND C.Pan_cAnio = G.Pan_cAnio AND C.Pla_cCuentaContable = G.Pla_cCuentaContable               
				  LEFT JOIN CNT_CENTRO_COSTO H WITH(READUNCOMMITTED) ON  C.Emp_cCodigo = H.Emp_cCodigo AND C.Pan_cAnio = H.Pan_cAnio AND C.Cos_cCodigo = H.Cos_cCodigo               
					LEFT JOIN CNT_LIBRO_OPERA I WITH(READUNCOMMITTED) ON               
				   A.Lib_cTipoLibro = I.Lib_cTipoLibro AND A.PAN_CANIO = I.PAN_CANIO AND A.Emp_cCodigo = I.Emp_cCodigo               
				   LEFT JOIN .CNT_TIPO_MONEDA J WITH(READUNCOMMITTED) ON A.Ase_cTipoMoneda = J.Mon_cCodigo AND A.Emp_cCodigo = J.Emp_cCodigo                
				  LEFT JOIN CNT_ENTIDAD K WITH(READUNCOMMITTED) ON F.Ten_cTipoEntidad = K.Ten_cTipoEntidad AND F.Emp_cCodigo = K.Emp_cCodigo                
				   LEFT JOIN CNT_TIPO_MONEDA L WITH(READUNCOMMITTED) ON C.Asd_cTipoMoneda = L.Mon_cCodigo AND C.Emp_cCodigo = L.Emp_cCodigo                
				               
				  LEFT JOIN CNT_CONCEPTO_LIBRO CPTO WITH(READUNCOMMITTED) ON              
				  C.Emp_cCodigo = CPTO.Emp_cCodigo AND C.Pan_cAnio = CPTO.Pan_cAnio AND C.Lib_cTipoLibro  = CPTO.Lib_cTipoLibro and C.Asd_cCodConcepto = CPTO.Asl_cCodigo                
				WHERE                   
				  (A.Emp_cCodigo = @Emp_cCodigo)               
				  AND A.Ase_cNummov = @Ase_cNummov              
				  AND (A.Pan_cAnio = @Pan_cAnio)               
				  AND (A.Per_cPeriodo = @Per_cPeriodo)               
				  AND (A.Lib_cTipoLibro = @Lib_cTipoLibro)               
				  AND (A.Ase_nVoucher = @Ase_nVoucher)               
				  AND a.Ase_cDeleted <> '*'               
				  AND c.Asd_cDeleted <> '*'               
				  AND (C.Asd_cDestino <> '1')               
				  AND C.Pla_cCuentaContable <> @CTA_GANANCIA              
				  AND C.Pla_cCuentaContable <> @CTA_PERDIDA  
				  AND  C.Asd_cNumDoc = RTRIM(LTRIM(@hasta)) AND C.Asd_cSerieDoc = RTRIM(LTRIM(@desde))               
				  ORDER BY Asd_nItem
			END
               
     End              
IF @Tipo = 'SEL_ALL'               
     Begin -- *** SELECCIONAR ASIENTOS POR RANGO DE FECHAS              
      SELECT  CNC_ASIENTO_VOUCHER.Ase_cNummov, CNC_ASIENTO_VOUCHER.Per_cPeriodo, CNC_ASIENTO_VOUCHER.Ase_dFecha, CNC_ASIENTO_VOUCHER.Lib_cTipoLibro,               
                      CNT_LIBRO_OPERA.Lib_cDescripcion, CNC_ASIENTO_VOUCHER.Ase_nVoucher, CNC_ASIENTO_VOUCHER.Ase_cTipoMoneda,              
                      CNT_TIPO_MONEDA.Mon_cNombreLargo, CNC_ASIENTO_VOUCHER.Ase_nTipoCambio, CNC_ASIENTO_VOUCHER.Ase_cGlosa,               
                      CND_ASIENTO_VOUCHER.Asd_nItem, CND_ASIENTO_VOUCHER.Pla_cCuentaContable, CNM_PLAN_CTA.Pla_cNombreCuenta,               
                      CNM_PLAN_CTA.Pla_cProvision, CND_ASIENTO_VOUCHER.Asd_cGlosa, CND_ASIENTO_VOUCHER.Asd_nDebeSoles,               
                      CND_ASIENTO_VOUCHER.Asd_nDebeMonExt, CND_ASIENTO_VOUCHER.Asd_nHaberSoles,               
   CND_ASIENTO_VOUCHER.Asd_nHaberMonExt, CND_ASIENTO_VOUCHER.Asd_nTipoCambio, CND_ASIENTO_VOUCHER.Cos_cCodigo,               
                      CNT_CENTRO_COSTO.Cos_cDescripcion, CND_ASIENTO_VOUCHER.Ten_cTipoEntidad, CND_ASIENTO_VOUCHER.Ent_cCodEntidad,  
       CASE CNt_ENTIDAD.Ten_cPlame     
      WHEN '1' THEN CNM_ENTIDAD.Ent_cApaterno+' '+CNM_ENTIDAD.Ent_cAmaterno+' '+CNM_ENTIDAD.Ent_cNombres    
      ELSE  CNM_ENTIDAD.Ent_cPersona    
      END as 'Ent_cPersona',   
                      CNT_ENTIDAD.Ten_cNombreEntidad, CND_ASIENTO_VOUCHER.Asd_cTipoDoc,               
                      CND_ASIENTO_VOUCHER.Asd_cSerieDoc, CND_ASIENTO_VOUCHER.Asd_cNumDoc, CND_ASIENTO_VOUCHER.Asd_dFecDoc, CND_ASIENTO_VOUCHER.Asd_dFecVen,               
                      CND_ASIENTO_VOUCHER.Asd_cTipoDocRef, CND_ASIENTO_VOUCHER.Asd_cSerieDocRef, CND_ASIENTO_VOUCHER.Asd_cNumDocRef,               
                      CND_ASIENTO_VOUCHER.Asd_dFecDocRef, CND_ASIENTO_VOUCHER.Asd_nMontoInafecto, CND_ASIENTO_VOUCHER.Asd_cRetencion,               
                      CND_ASIENTO_VOUCHER.Asd_dFechaSpot, CND_ASIENTO_VOUCHER.Asd_cNumSpot, CND_ASIENTO_VOUCHER.Asd_cDestino,                       CND_ASIENTO_VOUCHER.Asd_nCorre, Asd_cProvCanc,               
                      CNT_TIPODOC_1.Tdo_cNombreLargo, CNT_TIPODOC_1.Tdo_cNombreCorto, CNT_TIPODOC_2.Tdo_cNombreLargo AS Tdo_cNombreLargoRef,               
                      CNT_TIPODOC_2.Tdo_cNombreCorto AS Tdo_cNombreCortoRef, Ase_cUserModifica,               
                      EMPRESA.Emp_cNombreLargo, EMPRESA.Emp_cNombreCorto, GETDATE() AS FechaHoraImp ,              
          CND_ASIENTO_VOUCHER.Asd_cComprobante, CND_ASIENTO_VOUCHER.Asd_cProceso,               
          CND_ASIENTO_VOUCHER.ECP_COPERACION,  CND_ASIENTO_VOUCHER.Asd_cRegAux,              
          CND_ASIENTO_VOUCHER.Asd_cRegAuxDet, CND_ASIENTO_VOUCHER.Asd_cManual  ,              
              
          CND_ASIENTO_VOUCHER.Asd_cGrupo , CND_ASIENTO_VOUCHER.Asd_cCodConcepto, CPTO.Asl_cDescripcion               
              
     FROM     EMPRESA WITH(READUNCOMMITTED) INNER JOIN CNC_ASIENTO_VOUCHER WITH(READUNCOMMITTED) ON EMPRESA.Emp_cCodigo = CNC_ASIENTO_VOUCHER.Emp_cCodigo               
       LEFT JOIN  CND_ASIENTO_VOUCHER WITH(READUNCOMMITTED) ON               
                   CNC_ASIENTO_VOUCHER.Emp_cCodigo = CND_ASIENTO_VOUCHER.Emp_cCodigo AND               
                   CNC_ASIENTO_VOUCHER.Pan_cAnio = CND_ASIENTO_VOUCHER.Pan_cAnio AND               
                   CNC_ASIENTO_VOUCHER.Per_cPeriodo = CND_ASIENTO_VOUCHER.Per_cPeriodo AND               
                   CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = CND_ASIENTO_VOUCHER.Lib_cTipoLibro AND               
                   CNC_ASIENTO_VOUCHER.Ase_nVoucher = CND_ASIENTO_VOUCHER.Ase_nVoucher               
          LEFT JOIN CNT_TIPODOC CNT_TIPODOC_2 WITH(READUNCOMMITTED) ON CND_ASIENTO_VOUCHER.Asd_cTipoDocRef = CNT_TIPODOC_2.Tdo_cCodigo AND               
                   CND_ASIENTO_VOUCHER.Emp_cCodigo = CNT_TIPODOC_2.Emp_cCodigo               
         LEFT JOIN CNT_TIPODOC CNT_TIPODOC_1 WITH(READUNCOMMITTED) ON CND_ASIENTO_VOUCHER.Asd_cTipoDoc = CNT_TIPODOC_1.Tdo_cCodigo AND               
                   CND_ASIENTO_VOUCHER.Emp_cCodigo = CNT_TIPODOC_1.Emp_cCodigo               
       LEFT JOIN CNM_ENTIDAD WITH(READUNCOMMITTED) ON CND_ASIENTO_VOUCHER.Emp_cCodigo = CNM_ENTIDAD.Emp_cCodigo AND               
                   CND_ASIENTO_VOUCHER.Ten_cTipoEntidad = CNM_ENTIDAD.Ten_cTipoEntidad AND               
         CND_ASIENTO_VOUCHER.Ent_cCodEntidad = CNM_ENTIDAD.Ent_cCodEntidad               
       LEFT OUTER JOIN CNM_PLAN_CTA WITH(READUNCOMMITTED) ON CND_ASIENTO_VOUCHER.Emp_cCodigo = CNM_PLAN_CTA.Emp_cCodigo AND               
                   CND_ASIENTO_VOUCHER.Pan_cAnio = CNM_PLAN_CTA.Pan_cAnio AND               
                   CND_ASIENTO_VOUCHER.Pla_cCuentaContable = CNM_PLAN_CTA.Pla_cCuentaContable               
       LEFT JOIN CNT_CENTRO_COSTO WITH(READUNCOMMITTED) ON CND_ASIENTO_VOUCHER.Pan_cAnio = CNT_CENTRO_COSTO.Pan_cAnio AND               
       CND_ASIENTO_VOUCHER.Emp_cCodigo = CNT_CENTRO_COSTO.Emp_cCodigo AND CND_ASIENTO_VOUCHER.Cos_cCodigo = CNT_CENTRO_COSTO.Cos_cCodigo               
       LEFT JOIN CNT_LIBRO_OPERA WITH(READUNCOMMITTED) ON  CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = CNT_LIBRO_OPERA.Lib_cTipoLibro AND               
        CNC_ASIENTO_VOUCHER.PAN_CANIO = CNT_LIBRO_OPERA.PAN_CANIO AND               
                   CNC_ASIENTO_VOUCHER.Emp_cCodigo = CNT_LIBRO_OPERA.Emp_cCodigo               
                   LEFT JOIN  CNT_TIPO_MONEDA WITH(READUNCOMMITTED) ON CNC_ASIENTO_VOUCHER.Ase_cTipoMoneda = CNT_TIPO_MONEDA.Mon_cCodigo AND              
                   CNC_ASIENTO_VOUCHER.Emp_cCodigo = CNT_TIPO_MONEDA.Emp_cCodigo               
       LEFT JOIN  CNT_ENTIDAD WITH(READUNCOMMITTED) ON CNM_ENTIDAD.Ten_cTipoEntidad = CNT_ENTIDAD.Ten_cTipoEntidad AND               
       CNM_ENTIDAD.Emp_cCodigo = CNT_ENTIDAD.Emp_cCodigo                
  LEFT JOIN CNT_CONCEPTO_LIBRO CPTO WITH(READUNCOMMITTED) ON              
  CND_ASIENTO_VOUCHER.Emp_cCodigo = CPTO.Emp_cCodigo AND CND_ASIENTO_VOUCHER.Pan_cAnio = CPTO.Pan_cAnio AND CND_ASIENTO_VOUCHER.Lib_cTipoLibro  = CPTO.Lib_cTipoLibro and CND_ASIENTO_VOUCHER.Asd_cCodConcepto = CPTO.Asl_cCodigo                
              
                     
    WHERE  CNC_ASIENTO_VOUCHER.Emp_cCodigo = @Emp_cCodigo              
 AND Ase_dFecha >= @desde AND Ase_dFecha <= @hasta              
 AND (CNC_ASIENTO_VOUCHER.Ase_cDeleted <> '*')               
 AND (CND_ASIENTO_VOUCHER.Asd_cDeleted <> '*')               
    ORDER BY CNC_ASIENTO_VOUCHER.Per_cPeriodo,               
 CNC_ASIENTO_VOUCHER.Lib_cTipoLibro,               
 CNC_ASIENTO_VOUCHER.Ase_nVoucher,               
 Asd_nItem              
     End              
              
IF @Tipo = 'SEL_ALLCTA'               
     Begin -- *** SELECCIONAR ASIENTOS POR RANGO DE FECHAS              
     SELECT   CNC_ASIENTO_VOUCHER.Ase_cNummov, CNC_ASIENTO_VOUCHER.Per_cPeriodo, CNC_ASIENTO_VOUCHER.Ase_dFecha, CNC_ASIENTO_VOUCHER.Lib_cTipoLibro,               
                      CNT_LIBRO_OPERA.Lib_cDescripcion, CNC_ASIENTO_VOUCHER.Ase_nVoucher, CNC_ASIENTO_VOUCHER.Ase_cTipoMoneda,               
                      CNT_TIPO_MONEDA.Mon_cNombreLargo, CNC_ASIENTO_VOUCHER.Ase_nTipoCambio, CNC_ASIENTO_VOUCHER.Ase_cGlosa,               
                      CND_ASIENTO_VOUCHER.Asd_nItem, CND_ASIENTO_VOUCHER.Pla_cCuentaContable, CNM_PLAN_CTA.Pla_cNombreCuenta,               
                      CNM_PLAN_CTA.Pla_cProvision, CND_ASIENTO_VOUCHER.Asd_cGlosa, CND_ASIENTO_VOUCHER.Asd_nDebeSoles,               
                      CND_ASIENTO_VOUCHER.Asd_nDebeMonExt, CND_ASIENTO_VOUCHER.Asd_nHaberSoles,               
                      CND_ASIENTO_VOUCHER.Asd_nHaberMonExt, CND_ASIENTO_VOUCHER.Asd_nTipoCambio, CND_ASIENTO_VOUCHER.Cos_cCodigo,               
                      CNT_CENTRO_COSTO.Cos_cDescripcion, CND_ASIENTO_VOUCHER.Ten_cTipoEntidad, CND_ASIENTO_VOUCHER.Ent_cCodEntidad,               
       CASE CNt_ENTIDAD.Ten_cPlame     
      WHEN '1' THEN CNM_ENTIDAD.Ent_cApaterno+' '+CNM_ENTIDAD.Ent_cAmaterno+' '+CNM_ENTIDAD.Ent_cNombres    
      ELSE  CNM_ENTIDAD.Ent_cPersona    
      END as 'Ent_cPersona',  
      CNT_ENTIDAD.Ten_cNombreEntidad, CND_ASIENTO_VOUCHER.Asd_cTipoDoc,             
                      CND_ASIENTO_VOUCHER.Asd_cSerieDoc, CND_ASIENTO_VOUCHER.Asd_cNumDoc, CND_ASIENTO_VOUCHER.Asd_dFecDoc,               
                      CND_ASIENTO_VOUCHER.Asd_cTipoDocRef, CND_ASIENTO_VOUCHER.Asd_cSerieDocRef, CND_ASIENTO_VOUCHER.Asd_cNumDocRef,               
                      CND_ASIENTO_VOUCHER.Asd_dFecDocRef, CND_ASIENTO_VOUCHER.Asd_nMontoInafecto, CND_ASIENTO_VOUCHER.Asd_cRetencion,               
                      CND_ASIENTO_VOUCHER.Asd_dFechaSpot, CND_ASIENTO_VOUCHER.Asd_cNumSpot, CND_ASIENTO_VOUCHER.Asd_cDestino,         CND_ASIENTO_VOUCHER.Asd_nCorre, Asd_cProvCanc,               
                      CNT_TIPODOC_1.Tdo_cNombreLargo, CNT_TIPODOC_1.Tdo_cNombreCorto, CNT_TIPODOC_2.Tdo_cNombreLargo AS Tdo_cNombreLargoRef,               
                      CNT_TIPODOC_2.Tdo_cNombreCorto AS Tdo_cNombreCortoRef, Ase_cUserModifica,               
                      EMPRESA.Emp_cNombreLargo, EMPRESA.Emp_cNombreCorto, GETDATE() AS FechaHoraImp,              
        cnd_asiento_voucher.Asd_cComprobante, cnd_asiento_voucher.Asd_cProceso,              
        cnd_asiento_voucher.ECP_COPERACION,              
        cnd_asiento_voucher.Asd_cRegAux, cnd_asiento_voucher.Asd_cRegAuxDet,              
        cnd_asiento_voucher.Asd_cManual  ,              
                         
           CND_ASIENTO_VOUCHER.Asd_cGrupo , CND_ASIENTO_VOUCHER.Asd_cCodConcepto, CPTO.Asl_cDescripcion                    
              
     FROM    EMPRESA WITH(READUNCOMMITTED) INNER JOIN CNC_ASIENTO_VOUCHER WITH(READUNCOMMITTED) ON EMPRESA.Emp_cCodigo = CNC_ASIENTO_VOUCHER.Emp_cCodigo               
       LEFT JOIN CND_ASIENTO_VOUCHER WITH(READUNCOMMITTED) ON               
                   CNC_ASIENTO_VOUCHER.Emp_cCodigo = CND_ASIENTO_VOUCHER.Emp_cCodigo AND               
                   CNC_ASIENTO_VOUCHER.Pan_cAnio = CND_ASIENTO_VOUCHER.Pan_cAnio AND               
                   CNC_ASIENTO_VOUCHER.Per_cPeriodo = CND_ASIENTO_VOUCHER.Per_cPeriodo AND               
                   CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = CND_ASIENTO_VOUCHER.Lib_cTipoLibro AND               
                   CNC_ASIENTO_VOUCHER.Ase_nVoucher = CND_ASIENTO_VOUCHER.Ase_nVoucher               
       LEFT JOIN CNT_TIPODOC CNT_TIPODOC_2 WITH(READUNCOMMITTED) ON CND_ASIENTO_VOUCHER.Asd_cTipoDocRef = CNT_TIPODOC_2.Tdo_cCodigo AND               
                   CND_ASIENTO_VOUCHER.Emp_cCodigo = CNT_TIPODOC_2.Emp_cCodigo               
       LEFT JOIN CNT_TIPODOC CNT_TIPODOC_1 WITH(READUNCOMMITTED) ON CND_ASIENTO_VOUCHER.Asd_cTipoDoc = CNT_TIPODOC_1.Tdo_cCodigo AND               
                   CND_ASIENTO_VOUCHER.Emp_cCodigo = CNT_TIPODOC_1.Emp_cCodigo               
       LEFT JOIN CNM_ENTIDAD WITH(READUNCOMMITTED) ON CND_ASIENTO_VOUCHER.Emp_cCodigo = CNM_ENTIDAD.Emp_cCodigo AND               
                   CND_ASIENTO_VOUCHER.Ten_cTipoEntidad = CNM_ENTIDAD.Ten_cTipoEntidad AND               
                   CND_ASIENTO_VOUCHER.Ent_cCodEntidad = CNM_ENTIDAD.Ent_cCodEntidad               
         LEFT JOIN CNM_PLAN_CTA WITH(READUNCOMMITTED) ON CND_ASIENTO_VOUCHER.Emp_cCodigo = CNM_PLAN_CTA.Emp_cCodigo AND               
                   CND_ASIENTO_VOUCHER.Pan_cAnio = CNM_PLAN_CTA.Pan_cAnio AND               
                   CND_ASIENTO_VOUCHER.Pla_cCuentaContable = CNM_PLAN_CTA.Pla_cCuentaContable               
       LEFT JOIN CNT_CENTRO_COSTO WITH(READUNCOMMITTED) ON CND_ASIENTO_VOUCHER.Pan_cAnio = CNT_CENTRO_COSTO.Pan_cAnio AND               
       CND_ASIENTO_VOUCHER.Emp_cCodigo = CNT_CENTRO_COSTO.Emp_cCodigo AND CND_ASIENTO_VOUCHER.Cos_cCodigo = CNT_CENTRO_COSTO.Cos_cCodigo               
       LEFT JOIN CNT_LIBRO_OPERA WITH(READUNCOMMITTED) ON CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = CNT_LIBRO_OPERA.Lib_cTipoLibro AND               
         CNC_ASIENTO_VOUCHER.PAN_CANIO = CNT_LIBRO_OPERA.PAN_CANIO AND CNC_ASIENTO_VOUCHER.Emp_cCodigo = CNT_LIBRO_OPERA.Emp_cCodigo               
       LEFT JOIN  CNT_TIPO_MONEDA WITH(READUNCOMMITTED)ON CNC_ASIENTO_VOUCHER.Ase_cTipoMoneda = CNT_TIPO_MONEDA.Mon_cCodigo AND              
                   CNC_ASIENTO_VOUCHER.Emp_cCodigo = CNT_TIPO_MONEDA.Emp_cCodigo               
       LEFT JOIN  CNT_ENTIDAD WITH(READUNCOMMITTED) ON CNM_ENTIDAD.Ten_cTipoEntidad = CNT_ENTIDAD.Ten_cTipoEntidad AND               
       CNM_ENTIDAD.Emp_cCodigo = CNT_ENTIDAD.Emp_cCodigo                
  LEFT JOIN CNT_CONCEPTO_LIBRO CPTO WITH(READUNCOMMITTED) ON              
  CND_ASIENTO_VOUCHER.Emp_cCodigo = CPTO.Emp_cCodigo AND CND_ASIENTO_VOUCHER.Pan_cAnio = CPTO.Pan_cAnio AND CND_ASIENTO_VOUCHER.Lib_cTipoLibro  = CPTO.Lib_cTipoLibro and CND_ASIENTO_VOUCHER.Asd_cCodConcepto = CPTO.Asl_cCodigo                
              
                     
     WHERE CNC_ASIENTO_VOUCHER.Emp_cCodigo = @Emp_cCodigo              
 AND Ase_dFecha >= @desde AND Ase_dFecha <= @hasta              
 AND (CNC_ASIENTO_VOUCHER.Ase_cDeleted <> '*')               
 AND (CND_ASIENTO_VOUCHER.Asd_cDeleted <> '*')               
 AND CND_ASIENTO_VOUCHER.Pla_cCuentaContable = @Pla_cCuentaContable              
     ORDER BY CNC_ASIENTO_VOUCHER.Per_cPeriodo,               
 CNC_ASIENTO_VOUCHER.Lib_cTipoLibro,               
 CNC_ASIENTO_VOUCHER.Ase_nVoucher,               
 Asd_nItem              
End              
              
IF @Tipo = 'SEL_ALLCAB'               
     Begin -- *** SELECCIONAR CABECERAS DE ASIENTOS POR RANGO DE FECHAS              
 SELECT CNC_ASIENTO_VOUCHER.Ase_cNummov, CNC_ASIENTO_VOUCHER.PAN_CANIO, CNC_ASIENTO_VOUCHER.Per_cPeriodo, Ase_dFecha,               
    CNC_ASIENTO_VOUCHER.Lib_cTipoLibro, Lib_cDescripcion, CNC_ASIENTO_VOUCHER.Ase_nVoucher , Ase_cGlosa,               
   Ase_cTipoMoneda, Mon_cNombreLargo, Ase_nTipoCambio, TABLA.Tab_cDescripCampo As NomEstado,              
   --Sum(CND_ASIENTO_VOUCHER.Asd_nDebeSoles)as debe, Sum(CND_ASIENTO_VOUCHER.Asd_nHaberSoles)as haber              
   0 as debe, 0 as haber, CNC_ASIENTO_VOUCHER.Ase_cCuadreManual, Ase_cCodSoft, Ase_cElimSoft              
              
 FROM CNC_ASIENTO_VOUCHER WITH(READUNCOMMITTED)            
 LEFT OUTER JOIN CNT_LIBRO_OPERA WITH(READUNCOMMITTED) ON CNC_ASIENTO_VOUCHER.Lib_cTipoLibro = CNT_LIBRO_OPERA.Lib_cTipoLibro AND               
        CNC_ASIENTO_VOUCHER.Pan_cAnio = CNT_LIBRO_OPERA.Pan_cAnio AND               
        CNC_ASIENTO_VOUCHER.Emp_cCodigo = CNT_LIBRO_OPERA.Emp_cCodigo               
 LEFT OUTER JOIN CNT_TIPO_MONEDA WITH(READUNCOMMITTED) ON CNC_ASIENTO_VOUCHER.Ase_cTipoMoneda = CNT_TIPO_MONEDA.Mon_cCodigo AND               
        CNC_ASIENTO_VOUCHER.Emp_cCodigo = CNT_TIPO_MONEDA.Emp_cCodigo              
 LEFT JOIN TABLA WITH(READUNCOMMITTED) ON  TABLA.Emp_cCodigo = CNC_ASIENTO_VOUCHER.Emp_cCodigo AND TABLA.Tab_cTabla = '043'  AND TABLA.Tab_cCodigo = CNC_ASIENTO_VOUCHER.Ase_cEstado              
              
 WHERE CNC_ASIENTO_VOUCHER.Emp_cCodigo = @Emp_cCodigo               
   AND CNC_ASIENTO_VOUCHER.PAN_CANIO = @Pan_cAnio               
   AND Ase_cDeleted <> '*' --AND CND_ASIENTO_VOUCHER.Asd_cDestino = '0'               
   AND Ase_dFecha >= @desde              
   AND Ase_dFecha <= @hasta              
   and CNC_ASIENTO_VOUCHER.Lib_cTipoLibro  = @Lib_cTipoLibro               
 GROUP BY CNC_ASIENTO_VOUCHER.Ase_cNummov, CNC_ASIENTO_VOUCHER.PAN_CANIO, CNC_ASIENTO_VOUCHER.Per_cPeriodo, Ase_dFecha,               
        CNC_ASIENTO_VOUCHER.Lib_cTipoLibro, Lib_cDescripcion, CNC_ASIENTO_VOUCHER.Ase_nVoucher , Ase_cGlosa,               
        Ase_cTipoMoneda, Mon_cNombreLargo, Ase_nTipoCambio, TABLA.Tab_cDescripCampo, CNC_ASIENTO_VOUCHER.Ase_cCuadreManual,              
     Ase_cCodSoft, Ase_cElimSoft              
 ORDER BY CNC_ASIENTO_VOUCHER.Lib_cTipoLibro, CNC_ASIENTO_VOUCHER.Ase_nVoucher              
   End              
              
IF @Tipo = 'SEL_ALLCABPER'               
BEGIN  -- *** SELECCIONAR CABECERAS DE ASIENTOS POR RANGO DE FECHAS              
 SELECT DISTINCT A.Ase_cNummov, A.PAN_CANIO, A.Per_cPeriodo, A.Ase_dFecha,               
   A.Lib_cTipoLibro, C.Lib_cDescripcion, A.Ase_nVoucher, A.Ase_cGlosa,               
   A.Ase_cTipoMoneda, D.Mon_cNombreLargo, A.Ase_nTipoCambio, E.Tab_cDescripCampo As NomEstado,              
   0 As debe, 0 As haber , A.Ase_cCuadreManual, A.Ase_cCodSoft, A.Ase_cElimSoft              
   --sum(B.Asd_nDebeSoles) As debe, sum(B.Asd_nHaberSoles) As haber                 
 FROM CNC_ASIENTO_VOUCHER A WITH(READUNCOMMITTED) LEFT JOIN CNT_LIBRO_OPERA C WITH(READUNCOMMITTED)            
 ON A.Lib_cTipoLibro = C.Lib_cTipoLibro AND A.Pan_cAnio = C.Pan_cAnio AND A.Emp_cCodigo = C.Emp_cCodigo               
 LEFT JOIN CNT_TIPO_MONEDA D WITH(READUNCOMMITTED)            
 ON A.Ase_cTipoMoneda = D.Mon_cCodigo AND A.Emp_cCodigo = D.Emp_cCodigo              
 LEFT JOIN TABLA E WITH(READUNCOMMITTED)            
 ON  E.Emp_cCodigo = A.Emp_cCodigo AND E.Tab_cTabla = '043'  AND E.Tab_cCodigo = A.Ase_cEstado              
 WHERE A.Emp_cCodigo = @Emp_cCodigo               
   AND A.PAN_CANIO = @Pan_cAnio               
   AND A.Ase_cDeleted <> '*' --AND B.Asd_cDestino = '0'              
   AND A.Per_cPeriodo = @Per_cPeriodo              
   and a.Lib_cTipoLibro  = @Lib_cTipoLibro          
  ORDER BY A.Lib_cTipoLibro, A.Ase_nVoucher              
                
 -- SELECT * FROM CNC_ASIENTO_VOUCHER 7              
END               
               
-- spCn_ConsultaAsientos 'SEL_ALLCAB_REGAUX', '', '014', '2006', '05', '06', '', '', '', '', ''              
IF @Tipo = 'SEL_ALLCAB_REGAUX'               
     Begin               
 SELECT   A.Ase_cNummov, A.PAN_CANIO, A.Per_cPeriodo, A.Ase_dFecha,               
    B.Asd_cTipoDoc, B.Asd_cSerieDoc, B.Asd_cNumDoc, A.Ase_nVoucher,              
   A.Ase_cTipoMoneda, D.Mon_cNombreLargo, A.Ase_nTipoCambio, E.Tab_cDescripCampo As NomEstado,              
  TD.Tdo_cNombreLArgo,              
   abs(sum(B.Asd_nDebeSoles - B.Asd_nHaberSoles)) As Nac,               
   abs(sum(B.Asd_nDebeMonExt - B.Asd_nHaberMonExt)) As Ext              
              
 FROM CNC_ASIENTO_VOUCHER A WITH(READUNCOMMITTED) INNER JOIN CND_ASIENTO_VOUCHER B WITH(READUNCOMMITTED) ON               
  A.Emp_cCodigo = B.Emp_cCodigo AND A.Pan_cAnio = B.Pan_cAnio AND A.Ase_cNummov = B.Ase_cNummov AND               
  A.Ase_nVoucher = B.Ase_nVoucher AND A.Per_cPeriodo = B.Per_cPeriodo AND A.Lib_cTipoLibro = B.Lib_cTipoLibro               
  INNER JOIN CNT_LIBRO_OPERA C WITH(READUNCOMMITTED) ON A.Lib_cTipoLibro = C.Lib_cTipoLibro AND A.Pan_cAnio = C.Pan_cAnio AND A.Emp_cCodigo = C.Emp_cCodigo               
  INNER JOIN CNT_TIPO_MONEDA D WITH(READUNCOMMITTED) ON A.Ase_cTipoMoneda = D.Mon_cCodigo AND A.Emp_cCodigo = D.Emp_cCodigo              
  LEFT JOIN CNT_TIPODOC TD WITH(READUNCOMMITTED) ON A.Emp_cCodigo=  TD.Emp_cCodigo and  B.Asd_cTipoDoc = TD.Tdo_cCodigo              
  LEFT JOIN TABLA E WITH(READUNCOMMITTED) ON  E.Emp_cCodigo = A.Emp_cCodigo AND E.Tab_cTabla = '043'  AND E.Tab_cCodigo = A.Ase_cEstado              
  WHERE A.Emp_cCodigo = @Emp_cCodigo               
   AND A.PAN_CANIO = @Pan_cAnio               
   AND A.Ase_cDeleted <> '*'               
   AND B.Asd_cDestino = '0'              
   AND A.Per_cPeriodo = @Per_cPeriodo              
   AND A.Lib_cTipoLibro = @Lib_cTipoLibro               
   AND B.Asd_cRegAux = '1'               
   AND B.Asd_cTipodoc IN ( SELECT CONF.cod_cvalorparam FROM CND_CONFIG_OPERA CONF WITH(READUNCOMMITTED)            
      WHERE CONF.EMP_CCODIGO=@Emp_cCodigo AND CONF.PAN_CANIO=@Pan_cAnio AND CONF.COP_CCODIGO='028')              
              
  GROUP BY A.Ase_cNummov, A.PAN_CANIO, A.Per_cPeriodo, A.Ase_dFecha,               
     B.Asd_cTipoDoc, B.Asd_cSerieDoc, B.Asd_cNumDoc,  A.Ase_nVoucher,              
    A.Ase_cTipoMoneda, D.Mon_cNombreLargo, A.Ase_nTipoCambio, E.Tab_cDescripCampo ,              
   TD.Tdo_cNombreLArgo              
--  HAVING sum(B.Asd_nDebeSoles) > @UIT              
              
   End
GO
