USE SAFC_ECB
GO

CREATE PROC USP_Exportar_Percepcion_Ventas
@Emp_cCodigo CHAR(3),
@Pan_cAnio CHAR(4),
@Per_cPeriodo CHAR(2),
@Moneda CHAR(3)
AS
BEGIN
	SET NOCOUNT ON
	
	--DECLARE @Emp_cCodigo CHAR(3)
	--DECLARE @Pan_cAnio CHAR(4)
	--DECLARE @Per_cPeriodo CHAR(2)
	--DECLARE @Moneda CHAR(3)

	--SET @Emp_cCodigo = '017'
	--SET @Pan_cAnio = '2016'
	--SET @Per_cPeriodo = '02'
	--SET @Moneda = '038'

	DECLARE @Cta_Ganancia VARCHAR(12)
	DECLARE @Cta_Perdida VARCHAR(12)
	DECLARE @Letra CHAR(2)
	DECLARE @CodMoneda CHAR(3)
	DECLARE @LibroDiario CHAR(2)
	DECLARE @LibroApertura CHAR(2)
	DECLARE @LibroVenta CHAR(2)

	/*Obteniendo Cuenta Ganancia*/
	SELECT @Cta_Ganancia = cpc.Pla_cCuentaContable FROM dbo.CNM_PLAN_CTA CPC
	WHERE CPC.Emp_cCodigo = @Emp_cCodigo AND CPC.Pan_cAnio = @Pan_cAnio AND CPC.Pla_cDifCambio = 'G'

	/*Obteniendo Cuenta Perdida*/
	SELECT @Cta_Perdida = cpc.Pla_cCuentaContable FROM dbo.CNM_PLAN_CTA CPC
	WHERE CPC.Emp_cCodigo = @Emp_cCodigo AND CPC.Pan_cAnio = @Pan_cAnio AND CPC.Pla_cDifCambio = 'P'

	/*Obteniendo Letra*/
	SET @Letra = dbo.fBuscaConfOP(@Emp_cCodigo, @Pan_cAnio, '023')

	/*Obteniendo Libros*/
	SELECT @LibroApertura = CCL.Cfl_cApertura, @LibroDiario = Cfl_cDiario, @LibroVenta = CCL.Cfl_cVentas FROM dbo.CNT_CONFIG_LIBROS CCL
	WHERE CCL.Emp_cCodigo = @Emp_cCodigo

	/*Obteniendo Moneda*/
	SELECT @CodMoneda = ctm.Mon_cCodigo FROM dbo.CNT_TIPO_MONEDA CTM WHERE CTM.Emp_cCodigo = @Emp_cCodigo AND CTM.Mon_cMNac = '1'

	/*Obteniendo Retenciones*/
	SELECT cav.Emp_cCodigo, cav.Ase_cNummov, cav.Ase_nVoucher, cav.Per_cPeriodo, cav.Asd_nItem, cav.Pla_cCuentaContable, cav.Asd_cGlosa,
	   (CASE WHEN @CodMoneda = @Moneda THEN cav.Asd_nDebeSoles ELSE cav.Asd_nDebeMonExt END) AS Asd_nDebe,
	   (CASE WHEN @CodMoneda = @Moneda THEN cav.Asd_nHaberSoles ELSE cav.Asd_nHaberMonExt END) AS Asd_nHaber,
	   cav.Asd_nTipoCambio, cav.Asd_cTipoDoc, cav.Asd_cSerieDoc, cav.Asd_cNumDoc, cav.Asd_dFecDoc, cav.Asd_dFecVen,
	   cav.Asd_cTipoDocRef, cav.Asd_cSerieDocRef, cav.Asd_cNumDocRef, cav.Asd_dFecDocRef, cav.Asd_nMontoInafecto, cav.Asd_cRetencion, cav.Asd_cNumSpot,
	   cav.Asd_nCorre, cav.Asd_cFormaPago AS Lib_cTrans, cav.Ent_cCodEntidad, cav.Ten_cTipoEntidad, cav.Lib_cTipoLibro, 0 AS Asd_nAux,
	   ISNULL(CAV3.CreditoFiscal, '0') AS CreditoFiscal, ISNULL(CAV3.MaterialConstruccion, '0') AS MaterialConstruccion
	INTO #TMPPERCEPCION FROM dbo.CND_ASIENTO_VOUCHER CAV INNER JOIN dbo.CNC_ASIENTO_VOUCHER CAV3 ON cav.Ase_cNummov = CAV3.Ase_cNummov AND cav.Emp_cCodigo = CAV3.Emp_cCodigo AND cav.Pan_cAnio = CAV3.Pan_cAnio 
		 AND cav.Per_cPeriodo = CAV3.Per_cPeriodo AND cav.Lib_cTipoLibro = CAV3.Lib_cTipoLibro AND cav.Ase_nVoucher = CAV3.Ase_nVoucher
	WHERE CAV.Ase_nVoucher IN (SELECT CAV2.Ase_nVoucher FROM dbo.CND_ASIENTO_VOUCHER CAV2
						   WHERE CAV2.Asd_cRetencion = 'P' AND CAV2.Emp_cCodigo = @Emp_cCodigo AND CAV2.Per_cPeriodo <= @Per_cPeriodo AND CAV2.Asd_cDeleted <> '*'
								 AND CAV2.Pan_cAnio = @Pan_cAnio AND CAV2.Asd_cDestino = '0')
	  AND CAV.Emp_cCodigo = @Emp_cCodigo AND CAV.Per_cPeriodo <= @Per_cPeriodo AND CAV.Pan_cAnio = @Pan_cAnio AND CAV.Asd_cDeleted <> '*'
	  AND CAV.Asd_cRetencion = 'P' AND CAV.Asd_cDestino = '0'
	  
	/*Eliminar Cuentas distintas a Cuentas por Cobrar si el Libro es de Ventas*/	  
	DELETE #TMPPERCEPCION
	WHERE Lib_cTipoLibro = @LibroVenta AND LEFT(Pla_cCuentaContable, 2) NOT IN (SELECT CCO.Cod_cValorParam FROM dbo.CND_CONFIG_OPERA CCO
																			WHERE CCO.Emp_cCodigo = @Emp_cCodigo AND CCO.Pan_cAnio = @Pan_cAnio AND CCO.Cop_cCodigo = '009')

	/*Eliminar las Cuentas por cobrar si el libro es distinto de Ventas*/
	DELETE #TMPPERCEPCION
	WHERE Lib_cTipoLibro <> @LibroVenta AND Lib_cTipoLibro <> @LibroApertura AND LEFT(Pla_cCuentaContable, 2) IN (SELECT Cod_cValorParam FROM dbo.CND_CONFIG_OPERA CCO
																											  WHERE Emp_cCodigo = @Emp_cCodigo AND Pan_cAnio = @Pan_cAnio AND Cop_cCodigo = '009')
	  AND Asd_cTipoDoc <> @Letra
	  
	/*Eliminar Cuentas por Pagar si el libro es distinto de Compras y de Diario*/
	DELETE #TMPPERCEPCION
	WHERE Lib_cTipoLibro <> @LibroVenta AND Lib_cTipoLibro <> @LibroDiario AND Lib_cTipoLibro <> @LibroApertura
	  AND LEFT(Pla_cCuentaContable, 2) IN (SELECT cco.Cod_cValorParam FROM dbo.CND_CONFIG_OPERA CCO
										   WHERE CCO.Emp_cCodigo = @Emp_cCodigo AND Pan_cAnio = @Pan_cAnio AND CCO.Cop_cCodigo = '009')	  																											  
										   
	/*Eliminar que tengan cuentas por Pagar (40 o 10)*/										   
	DELETE #TMPPERCEPCION
	WHERE Ase_nVoucher IN (SELECT Ase_nVoucher from #TMPPERCEPCION    
					   where left(pla_ccuentaContable,2 ) in (select Cod_cValorParam from cnd_config_opera     
															  where emp_ccodigo=@Emp_cCodigo and pan_canio= @Pan_cAnio and cop_ccodigo='010'))
	                                                          
	/*Invertir Cuando son cobros o Letras*/                                                              
	UPDATE #TMPPERCEPCION
	SET Asd_nAux = Asd_nDebe
	WHERE Lib_cTipoLibro <> @LibroVenta AND Lib_cTipoLibro <> @LibroApertura AND Asd_cTipoDoc <> @Letra

	UPDATE #TMPPERCEPCION
	SET Asd_nDebe = Asd_nHaber, Asd_nHaber = Asd_nAux
	WHERE Lib_cTipoLibro <> @LibroVenta AND Lib_cTipoLibro <> @LibroApertura AND Asd_cTipoDoc <> @Letra

	/*Actualizar Entidad*/
	DECLARE @Serie VARCHAR(20)
	DECLARE @Numero VARCHAR(20)
	DECLARE @CodEntidad CHAR(5)
	DECLARE @TipoEntidad CHAR(1)
	DECLARE @MontoDebe DECIMAL(14, 3)
	DECLARE C_Entidad CURSOR FOR
	SELECT Asd_cSerieDoc, Asd_cNumDoc, Ten_cTipoEntidad, Ent_cCodEntidad, Asd_nDebe FROM #TMPPERCEPCION T
	WHERE T.Asd_cSerieDocRef = ''

	OPEN C_Entidad
	FETCH NEXT FROM C_Entidad INTO @Serie, @Numero, @TipoEntidad, @CodEntidad, @MontoDebe
	WHILE @@FETCH_STATUS = 0
	BEGIN

	UPDATE #TMPPERCEPCION
		SET Ten_cTipoEntidad = @TipoEntidad, Ent_cCodEntidad = @CodEntidad, Asd_nDebe = @MontoDebe
	WHERE Asd_cSerieDocRef = @Serie AND Asd_cNumDocRef = @Numero

	FETCH NEXT FROM C_Entidad INTO @Serie, @Numero, @TipoEntidad, @CodEntidad, @MontoDebe	
	END
	CLOSE C_Entidad
	DEALLOCATE C_Entidad


	DECLARE @Separador CHAR(1)
	SET @Separador = '|'

	--SELECT (ce.Ent_nRuc + @Separador + ce.Ent_cPersona + @Separador + ISNULL(ce.Ent_cApaterno, '') + @Separador + ISNULL(ce.Ent_cAmaterno, '') + @Separador + ISNULL(ce.Ent_cNombres, '') + @Separador +
	--   t.Asd_cSerieDoc + @Separador + t.Asd_cNumDoc + @Separador +
	--   CONVERT(NCHAR(10), t.Asd_dFecDocRef, 103) + @Separador + CAST(CAST(t.Asd_nHaber AS NUMERIC(12, 2)) AS VARCHAR(50)) + @Separador + RTRIM(LTRIM(t.Asd_cTipoDocRef)) + @Separador + RTRIM(LTRIM(t.Asd_cSerieDocRef)) + @Separador + 
	--   @Separador + RTRIM(LTRIM(t.Asd_cNumDocRef)) + @Separador + CONVERT(NCHAR(10), t.Asd_dFecDoc, 103) + @Separador +
	--   CAST(CAST(t.Asd_nDebe AS NUMERIC(12, 2)) AS VARCHAR(50)) + @Separador) AS Registro FROM #TMPPERCEPCION T
	--LEFT JOIN dbo.CNM_ENTIDAD CE ON T.Emp_cCodigo = CE.Emp_cCodigo AND T.Ent_cCodEntidad = CE.Ent_cCodEntidad AND T.Ten_cTipoEntidad = CE.Ten_cTipoEntidad
	--WHERE t.Asd_nHaber > 0

	SELECT (T2.Tab_cCodSunat + @Separador + CE.Ent_nRuc + @Separador + CE.Ent_cPersona + @Separador + ISNULL(CE.Ent_cApaterno, '') + @Separador + ISNULL(CE.Ent_cAmaterno, '') + @Separador + 
			ISNULL(CE.Ent_cNombres, '') + @Separador +
			RTRIM(LTRIM(T.Asd_cSerieDocRef)) + @Separador + RTRIM(LTRIM(t.Asd_cNumDocRef)) + @Separador +
			CONVERT(NCHAR(10), t.Asd_dFecDoc, 103) + @Separador + ISNULL(CreditoFiscal, '0') + @Separador + ISNULL(MaterialConstruccion, '0') + @Separador + ISNULL(ce.PorcentajeSunat, '0') + @Separador + CAST(CAST(t.Asd_nHaber AS NUMERIC(12, 2)) AS VARCHAR(50)) + @Separador +
			RTRIM(LTRIM(t.Asd_cTipoDocRef)) + @Separador + t.Asd_cSerieDoc + @Separador + t.Asd_cNumDoc + @Separador + CONVERT(NCHAR(10), t.Asd_dFecDocRef, 103) + @Separador +
			CAST(CAST(t.Asd_nDebe AS NUMERIC(12, 2)) AS VARCHAR(50)) + @Separador) AS Registro FROM #TMPPERCEPCION T
	LEFT JOIN dbo.CNM_ENTIDAD CE ON T.Emp_cCodigo = CE.Emp_cCodigo AND T.Ent_cCodEntidad = CE.Ent_cCodEntidad AND T.Ten_cTipoEntidad = CE.Ten_cTipoEntidad
	LEFT JOIN dbo.TABLA T2 ON CE.Emp_cCodigo = T2.Emp_cCodigo AND CE.Ent_cTipoDoc = T2.Tab_cCodigo AND T2.Tab_cTabla = '003'
	WHERE t.Asd_nHaber > 0 AND t.Per_cPeriodo = @Per_cPeriodo AND (CreditoFiscal <> '0' OR MaterialConstruccion <> '0' OR PorcentajeSunat <> '0')
	
END
GO

