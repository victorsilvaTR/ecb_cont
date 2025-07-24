USE SAFC_ECB
GO

ALTER PROC USP_Exportar_Retencion_PDT_621
@Emp_cCodigo CHAR(3),
@Pan_cAnio CHAR(4),
@Per_cPeriodo CHAR(2),
@Moneda CHAR(3)
AS
BEGIN
	SET NOCOUNT ON
	
	--	DECLARE @Emp_cCodigo CHAR(3)
	--DECLARE @Pan_cAnio CHAR(4)
	--DECLARE @Per_cPeriodo CHAR(2)
	--DECLARE @Moneda CHAR(3)

	--SET @Emp_cCodigo = '017'
	--SET @Pan_cAnio = '2016'
	--SET @Per_cPeriodo = '03'
	--SET @Moneda = '038'


	/*Variables*/
	DECLARE @Cta_Ganancia VARCHAR(12)
	DECLARE @Cta_Perdida VARCHAR(12)
	DECLARE @LibroCompra CHAR(2)
	DECLARE @LibroDiario CHAR(2)
	DECLARE @LibroApertura CHAR(2)
	DECLARE @Letra CHAR(2)
	DECLARE @CodMoneda CHAR(3)

	/*Obteniendo Cuenta Ganancia*/
	SELECT @Cta_Ganancia = cpc.Pla_cCuentaContable FROM dbo.CNM_PLAN_CTA CPC
	WHERE CPC.Emp_cCodigo = @Emp_cCodigo AND CPC.Pan_cAnio = @Pan_cAnio AND CPC.Pla_cDifCambio = 'G'

	/*Obteniendo Cuenta Perdida*/
	SELECT @Cta_Perdida = cpc.Pla_cCuentaContable FROM dbo.CNM_PLAN_CTA CPC
	WHERE CPC.Emp_cCodigo = @Emp_cCodigo AND CPC.Pan_cAnio = @Pan_cAnio AND CPC.Pla_cDifCambio = 'P'

	/*Obteniendo Letra*/
	SET @Letra = dbo.fBuscaConfOP(@Emp_cCodigo, @Pan_cAnio, '022')

	/*Obteniendo Libros*/
	SELECT @LibroApertura = CCL.Cfl_cApertura, @LibroCompra = CCL.Cfl_cCompras, @LibroDiario = CCL.Cfl_cDiario FROM dbo.CNT_CONFIG_LIBROS CCL
	WHERE CCL.Emp_cCodigo = @Emp_cCodigo

	/*Obteniendo Codigo Moneda*/
	SELECT @CodMoneda = ctm.Mon_cCodigo FROM dbo.CNT_TIPO_MONEDA CTM
	WHERE CTM.Emp_cCodigo = @Emp_cCodigo AND CTM.Mon_cMNac = '1'


	/*Obteniendo Asientos de Retencion*/
	SELECT CAV.Ase_cNummov, CAV.Ase_nVoucher, CAV.Per_cPeriodo, CAV.Asd_nItem, CAV.Pla_cCuentaContable, CAV.Asd_cGlosa,
		   (CASE WHEN @Moneda = @CodMoneda THEN CAV.Asd_nDebeSoles ELSE CAV.Asd_nDebeMonExt END) AS Asd_nDebe,
		   (CASE WHEN @Moneda = @CodMoneda THEN CAV.Asd_nHaberSoles ELSE CAV.Asd_nHaberMonExt END) AS Asd_nHaber, CAV.Asd_nTipoCambio, CAV.Asd_cTipoDoc, CAV.Asd_cSerieDoc,
		   CAV.Asd_cNumDoc, CONVERT(NCHAR(10), CAV.Asd_dFecDoc, 103) AS Asd_dFecDoc, CASE WHEN YEAR(ISNULL(CAV.Asd_dFecVen, '')) = 1900 THEN '' ELSE CONVERT(NCHAR(10), ISNULL(CAV.Asd_dFecVen, ''), 103) END AS Asd_dFecVen, 
		   ISNULL(CAV.Asd_cTipoDocRef, '') AS Asd_cTipoDocRef, ISNULL(CAV.Asd_cSerieDocRef, '') AS Asd_cSerieDocRef, ISNULL(CAV.Asd_cNumDocRef,'') AS Asd_cNumDocRef, 
		   CASE WHEN YEAR(ISNULL(CAV.Asd_dFecDocRef, '')) = 1900 THEN '' ELSE CONVERT(NCHAR(10), ISNULL(CAV.Asd_dFecDocRef, ''), 103) END AS Asd_dFecDocRef, CAV.Asd_nMontoInafecto, CAV.Asd_cRetencion,
		   CAV.Asd_dFechaSpot, CAV.Asd_cNumSpot, CAV.Asd_nCorre, CAV.Asd_cFormaPago, CAV.Lib_cTipoLibro, CAV.Ten_cTipoEntidad, CAV.Ent_cCodEntidad, ISNULL(ce.Ent_cPersona, '') AS Ent_cPersona,
		   ISNULL(ce.Ent_cApaterno, '') AS Ent_cApaterno, ISNULL(ce.Ent_cAmaterno, '') AS Ent_cAmaterno, ISNULL(ce.Ent_cNombres, '') AS Ent_cNombres, 
		   ISNULL(t.Tab_cCodSunat, '') AS Tab_cCodSunat, ISNULL(ce.Ent_nRuc, '') AS Ent_nRuc, 0 AS Asd_nAux 
		   INTO #TMPRETENCIONES FROM dbo.CND_ASIENTO_VOUCHER CAV
		   LEFT JOIN dbo.CNM_ENTIDAD CE ON CAV.Emp_cCodigo = CE.Emp_cCodigo AND CAV.Ent_cCodEntidad = CE.Ent_cCodEntidad AND CAV.Ten_cTipoEntidad = CE.Ten_cTipoEntidad
		   LEFT JOIN dbo.TABLA T ON CAV.Emp_cCodigo = T.Emp_cCodigo AND T.Tab_cTabla = '003' AND T.Tab_cCodigo = ce.Ent_cTipoDoc
	WHERE CAV.Ase_nVoucher IN (SELECT CAV2.Ase_nVoucher FROM dbo.CND_ASIENTO_VOUCHER CAV2 WHERE CAV2.Asd_cRetencion = 'R' AND CAV2.Emp_cCodigo = @Emp_cCodigo AND CAV2.Per_cPeriodo <= @Per_cPeriodo
							   AND CAV2.Asd_cDeleted <> '*' AND CAV2.Pan_cAnio = @Pan_cAnio AND CAV2.Asd_cDestino = '0')
		  AND CAV.Emp_cCodigo = @Emp_cCodigo AND CAV.Per_cPeriodo <= @Per_cPeriodo AND CAV.Pan_cAnio = @Pan_cAnio AND CAV.Asd_cRetencion = 'R' AND CAV.Asd_cDestino = '0'
		 	  
	/*Eliminar Cuentas distintas al Libro Compras*/	
	DELETE #TMPRETENCIONES
	WHERE Lib_cTipoLibro = @LibroCompra
		  AND LEFT(Pla_cCuentaContable, 2) NOT IN (SELECT CCO.Cod_cValorParam FROM dbo.CND_CONFIG_OPERA CCO
												   WHERE CCO.Emp_cCodigo = @Emp_cCodigo AND CCO.Pan_cAnio = @Pan_cAnio AND CCO.Cop_cCodigo = '010')
																	   
	/*Asignar Codigo y Tipo de Entidad*/
	DECLARE @Ent_cCodEntidad CHAR(5)
	DECLARE @Ten_cTipoEntidad CHAR(1)
	DECLARE @Serie VARCHAR(20)
	DECLARE @Numero VARCHAR(25)
	DECLARE @Voucher CHAR(10)
	DECLARE @Movimiento CHAR(10)
	DECLARE @Ent_cPersona VARCHAR(100)
	DECLARE @Tab_cCodSunat CHAR(1)
	DECLARE @Ent_nRuc VARCHAR(20)

	DECLARE C_EntidadReferencia CURSOR FOR
	SELECT DISTINCT T.Asd_cSerieDocRef, T.Asd_cNumDocRef, T.Ent_cCodEntidad, T.Ten_cTipoEntidad, T.Ent_cPersona, T.Tab_cCodSunat, T.Ent_nRuc FROM #TMPRETENCIONES T
	WHERE T.Asd_cSerieDocRef <> '' AND T.Asd_cNumDocRef <> '' AND T.Ent_cCodEntidad <> '' AND T.Ten_cTipoEntidad <> ''
	ORDER BY T.Ent_cCodEntidad

	OPEN C_EntidadReferencia
	FETCH NEXT FROM C_EntidadReferencia INTO @Serie, @Numero, @Ent_cCodEntidad, @Ten_cTipoEntidad, @Ent_cPersona, @Tab_cCodSunat, @Ent_nRuc
	WHILE @@FETCH_STATUS = 0
	BEGIN
		
		UPDATE #TMPRETENCIONES
			SET Ent_cCodEntidad = @Ent_cCodEntidad, Ten_cTipoEntidad = @Ten_cTipoEntidad, Ent_cPersona = @Ent_cPersona, Tab_cCodSunat = @Tab_cCodSunat, Ent_nRuc = @Ent_nRuc
		WHERE Asd_cSerieDoc = @Serie AND Asd_cNumDoc = @Numero
		
		FETCH NEXT FROM C_EntidadReferencia INTO @Serie, @Numero, @Ent_cCodEntidad, @Ten_cTipoEntidad, @Ent_cPersona, @Tab_cCodSunat, @Ent_nRuc
	END
	CLOSE C_EntidadReferencia
	DEALLOCATE C_EntidadReferencia
	----------------------------------------------------------------------------------------------------------------
	DECLARE C_Entidad CURSOR FOR
	SELECT T.Asd_cSerieDoc, T.Asd_cNumDoc, T.Ent_cCodEntidad, T.Ten_cTipoEntidad, T.Ent_cPersona, T.Tab_cCodSunat, T.Ent_nRuc FROM #TMPRETENCIONES T
	WHERE T.Asd_cSerieDoc <> '' AND T.Asd_cNumDoc <> '' AND T.Ent_cCodEntidad <> '' AND T.Ten_cTipoEntidad <> ''
	ORDER BY T.Ent_cCodEntidad

	OPEN C_Entidad
	FETCH NEXT FROM C_Entidad INTO @Serie, @Numero, @Ent_cCodEntidad, @Ten_cTipoEntidad, @Ent_cPersona, @Tab_cCodSunat, @Ent_nRuc
	WHILE @@FETCH_STATUS = 0
	BEGIN
		
		UPDATE #TMPRETENCIONES
			SET Ent_cCodEntidad = @Ent_cCodEntidad, Ten_cTipoEntidad = @Ten_cTipoEntidad, Ent_cPersona = @Ent_cPersona, Tab_cCodSunat = @Tab_cCodSunat, Ent_nRuc = @Ent_nRuc
		WHERE Asd_cSerieDocRef = @Serie AND Asd_cNumDocRef = @Numero
		
		FETCH NEXT FROM C_Entidad INTO @Serie, @Numero, @Ent_cCodEntidad, @Ten_cTipoEntidad, @Ent_cPersona, @Tab_cCodSunat, @Ent_nRuc
	END
	CLOSE C_Entidad
	DEALLOCATE C_Entidad											   
							
							
	/*Eliminar las cuentas por Pagar si el Libro es distinto a Compras y no Letra*/										   
	DELETE #TMPRETENCIONES
	WHERE Lib_cTipoLibro <> @LibroCompra AND Lib_cTipoLibro <> @LibroApertura
		  AND LEFT(Pla_cCuentaContable, 2) IN (SELECT CCO.Cod_cValorParam FROM dbo.CND_CONFIG_OPERA CCO
											   WHERE CCO.Emp_cCodigo = @Emp_cCodigo AND CCO.Pan_cAnio = @Pan_cAnio AND CCO.Cop_cCodigo = '010')
											   
	/*Eliminar cuentas por Pagar si el Libro es distinto a Compras y Diario*/										   
	DELETE #TMPRETENCIONES
	WHERE Lib_cTipoLibro <> @LibroCompra AND Lib_cTipoLibro <> @LibroApertura AND Lib_cTipoLibro = @LibroDiario
		  AND LEFT(Pla_cCuentaContable, 2) IN (SELECT CCO.Cod_cValorParam FROM dbo.CND_CONFIG_OPERA CCO
											   WHERE CCO.Emp_cCodigo = @Emp_cCodigo AND CCO.Pan_cAnio = @Pan_cAnio AND CCO.Cop_cCodigo = '010')
											   
	/*Eliminar los Voucher que tengan las cuentas por Cobrar y otras cuentas como la 40 y 10*/
	DELETE #TMPRETENCIONES
	WHERE ase_nVoucher IN (SELECT T.Ase_nVoucher FROM #TMPRETENCIONES T
						   WHERE LEFT(Pla_cCuentaContable, 2) IN (SELECT CCO.Cod_cValorParam FROM dbo.CND_CONFIG_OPERA CCO
																  WHERE CCO.Emp_cCodigo = @Emp_cCodigo AND CCO.Pan_cAnio = @Pan_cAnio AND CCO.Cop_cCodigo = '009'))	
																  
	/*Invertir Montos*/															  									   
	UPDATE #TMPRETENCIONES
		SET Asd_nAux = Asd_nDebe
	WHERE (Lib_cTipoLibro <> @LibroCompra AND Lib_cTipoLibro <> @LibroApertura) AND Asd_cTipoDoc <> @Letra


	DECLARE C_SerieNumero CURSOR
	FOR SELECT Ase_cNummov, Ase_nVoucher, Asd_cSerieDoc, Asd_cNumDoc FROM #TMPRETENCIONES WHERE LEFT(Pla_cCuentaContable, 5) = '40114'
	OPEN C_SerieNumero
	FETCH NEXT FROM C_SerieNumero INTO @Movimiento, @Voucher, @Serie, @Numero
	WHILE @@FETCH_STATUS = 0
	BEGIN
		
		UPDATE #TMPRETENCIONES
			SET Asd_cSerieDoc = @Serie, Asd_cNumDoc = @Numero
		WHERE Ase_cNummov = @Movimiento AND Ase_nVoucher = @Voucher AND LEFT(Pla_cCuentaContable, 5) <> '40114'
		
		FETCH NEXT FROM C_SerieNumero INTO @Movimiento, @Voucher, @Serie, @Numero	
	END
	CLOSE C_SerieNumero
	DEALLOCATE C_SerieNumero

	UPDATE #TMPRETENCIONES
		SET Asd_nDebe = Asd_nHaber, Asd_nHaber = Asd_nAux
	WHERE (Lib_cTipoLibro <> @LibroCompra AND Lib_cTipoLibro <> @LibroApertura) AND Asd_cTipoDoc <> @Letra

	/*Actualizar Serie y Numero igual al Documento de Referencia*/
	DECLARE @SerieRef CHAR(5)
	DECLARE @NumeroRef CHAR(10)
	DECLARE @TipoDocRef CHAR(2)
	DECLARE @FechaRef CHAR(10)

	DECLARE @MontoHaber DECIMAL(14, 3)
	DECLARE C_Monto CURSOR FOR
	SELECT DISTINCT Asd_cSerieDocRef, Asd_cNumDocRef FROM #TMPRETENCIONES
	WHERE Asd_cSerieDocRef <> ''
	OPEN C_Monto
	FETCH NEXT FROM C_Monto INTO @SerieRef, @NumeroRef
	WHILE @@FETCH_STATUS = 0
	BEGIN
		
		SELECT @MontoHaber = Asd_nHaber FROM #TMPRETENCIONES
		WHERE Asd_cSerieDoc = @SerieRef AND Asd_cNumDoc = @NumeroRef
		
		UPDATE #TMPRETENCIONES
			SET Asd_nHaber = @MontoHaber
		WHERE Asd_cSerieDocRef = @SerieRef AND Asd_cNumDocRef = @NumeroRef
		
		FETCH NEXT FROM C_Monto INTO @SerieRef, @NumeroRef	
	END 
	CLOSE C_Monto
	DEALLOCATE C_Monto


	DECLARE C_ActualizaDocumento CURSOR FOR
	SELECT DISTINCT Asd_cSerieDocRef, Asd_cNumDocRef, Asd_cSerieDoc, Asd_cNumDoc, Asd_cTipoDocRef, Asd_dFecDocRef FROM #TMPRETENCIONES
	WHERE Asd_cTipoDocRef <> ''
	OPEN C_ActualizaDocumento
	FETCH NEXT FROM C_ActualizaDocumento INTO @SerieRef, @NumeroRef, @Serie, @Numero, @TipoDocRef, @FechaRef
	WHILE @@FETCH_STATUS = 0
	BEGIN
		
		UPDATE #TMPRETENCIONES
			SET Asd_cSerieDoc = @Serie , Asd_cNumDoc = @Numero, Asd_cSerieDocRef = @SerieRef, Asd_cNumDocRef = @NumeroRef, Asd_cTipoDocRef = @TipoDocRef, Asd_dFecDocRef = @FechaRef
		WHERE Asd_cSerieDoc = @SerieRef AND Asd_cNumDoc = @NumeroRef
		
		FETCH NEXT FROM C_ActualizaDocumento INTO @SerieRef, @NumeroRef, @Serie, @Numero,  @TipoDocRef, @FechaRef	
	END
	CLOSE C_ActualizaDocumento
	DEALLOCATE C_ActualizaDocumento

	--SELECT * FROM #TMPRETENCIONES T
	--WHERE LEFT(T.Pla_cCuentaContable, 5) <> '40114'

	DECLARE @Separador CHAR(1)
	SET @Separador = '|'
	
	SELECT (Ent_nRuc + @Separador + Asd_cSerieDoc + @Separador + Asd_cNumDoc + @Separador + Asd_dFecDoc + @Separador + CAST(CAST(Asd_nDebe AS NUMERIC(14, 2)) AS VARCHAR(50)) + @Separador +
	       RTRIM(LTRIM(Asd_cTipoDocRef)) + @Separador + RTRIM(LTRIM(Asd_cSerieDocRef)) + @Separador + RTRIM(LTRIM(Asd_cNumDocRef)) + @Separador + Asd_dFecDocRef + @Separador + CAST(CAST(Asd_nHaber AS NUMERIC(14, 2)) AS VARCHAR(50)) + @Separador) AS Registro
	        FROM #TMPRETENCIONES
	WHERE LEFT(Pla_cCuentaContable, 5) <> '40114' AND Per_cPeriodo = @Per_cPeriodo AND Asd_nDebe > 0
	
	
	--SELECT (T.Ent_nRuc + @Separador + T.Ent_cPersona + @Separador + T.Ent_cApaterno + @Separador + T.Ent_cAmaterno + @Separador + T.Ent_cNombres + @Separador + T.Asd_cSerieDoc
	--	   + @Separador + T.Asd_cNumDoc + @Separador + T.Asd_dFecDoc + @Separador + CAST(CAST(SUM(T.Asd_nDebe) AS NUMERIC(14, 2)) AS VARCHAR(50)) + @Separador +
	--	   RTRIM(LTRIM(T.Asd_cTipoDocRef)) + @Separador + RTRIM(LTRIM(T.Asd_cSerieDocRef)) + @Separador + RTRIM(LTRIM(T.Asd_cNumDocRef)) + @Separador + T.Asd_dFecDocRef + @Separador + CAST(CAST(T.Asd_nHaber AS NUMERIC(14, 2)) AS VARCHAR(50)) + @Separador) AS Registro FROM #TMPRETENCIONES T
	--WHERE LEFT(Pla_cCuentaContable, 5) <> '40114' AND Per_cPeriodo = @Per_cPeriodo
	--GROUP BY T.Ent_nRuc, T.Ent_cPersona, T.Ent_cApaterno, T.Ent_cAmaterno, T.Ent_cNombres, T.Asd_cSerieDoc, T.Asd_cNumDoc, T.Asd_dFecDoc,
	--		 T.Asd_cTipoDocRef, T.Asd_cSerieDocRef, T.Asd_cNumDocRef, T.Asd_dFecDocRef, T.Asd_nHaber
	
END
GO

