USE SAFC_ECB
GO

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
GO
ALTER PROCEDURE [dbo].[spCn_GrabaEntidad](    
@Accion varchar(20),    
@Emp_cCodigo char(3)='',    
@Ent_cCodEntidad char(5)='',    
@Ten_cTipoEntidad char(1)='',    
@Ent_cPersona varchar(120)='',    
@Ent_cDireccion varchar(120)='',    
@Ent_nRuc varchar(15)='',    
@Ent_cRepresentante varchar(80)='',    
@Ent_cTipoDoc char(3)='',    
@Ent_cFlagPersona char(1)='',    
@Ent_cEstadoEntidad char(1)='',    
@Ent_cEstado char(1)='',    
@Ent_cUserCrea varchar(20)='',
@Ent_cApaterno varchar(40)='', /*************NUEVOS REGISTROS 02/07/2013 - PAUL CUEVA */
@Ent_cAmaterno varchar(40)='',
@Ent_cNombres varchar(40)='',
@Ent_cFlagDomiciliado char(1)='0',
@Ent_cAconvenio char(3)='0',
@Per_cPeriodo CHAR(2) = '',
@Id_Pais CHAR(10) = NULL,
@Id_Vinculo_Economico CHAR(10) = NULL,   /* FIN DE REGISTROS */ 
@Id_Convenio CHAR(10) = NULL,
@PorcentajeSunat CHAR(1) = '0'
) 
--WITH ENCRYPTION    
AS      
SET DATEFORMAT DMY    
-----
    
DECLARE @INTERNO VARCHAR(5)    
IF @Accion = 'INSERTAR'    
BEGIN   

--Asignar Correlativo Entidad (Multiusuario)
 SELECT @Ent_cCodEntidad = ISNULL(max(convert(int,Ent_cCodEntidad))+1, 1) FROM CNM_ENTIDAD WITH(NOLOCK)    
 WHERE Emp_cCodigo = @Emp_cCodigo AND Ten_cTipoEntidad = @Ten_cTipoEntidad  AND Ent_cCodEntidad<>'99999'
 SET @Ent_cCodEntidad = REPLICATE('0', 5-len(@Ent_cCodEntidad)) + @Ent_cCodEntidad 

 INSERT INTO CNM_ENTIDAD WITH(ROWLOCK)    
  (Emp_cCodigo, Ent_cCodEntidad, Ten_cTipoEntidad, Ent_cPersona, Ent_cDireccion, Ent_nRuc, Ent_cRepresentante,     
  Ent_cTipoDoc, Ent_cFlagPersona, Ent_cEstadoEntidad, Ent_cEstado,  Ent_cDeleted, Ent_cUserCrea, Ent_dFechaCrea,
  Ent_cUserModifica, Ent_dFechaModifica, Ent_cEquipoUser, Ent_cApaterno, Ent_cAmaterno, Ent_cNombres, Ent_cFlagDomiciliado,
  Ent_cAconvenio, Id_Pais, Id_Vinculo_Economico, Id_Convenio, PorcentajeSunat)
 VALUES
  (@Emp_cCodigo,@Ent_cCodEntidad,@Ten_cTipoEntidad,@Ent_cPersona,@Ent_cDireccion,@Ent_nRuc,@Ent_cRepresentante,
  @Ent_cTipoDoc,@Ent_cFlagPersona,  '1','A','',@Ent_cUserCrea,getdate(),@Ent_cUserCrea,getdate(),host_name(),
  @Ent_cApaterno,@Ent_cAmaterno,/* NUEVOS REGISTROS 02/07/2013 - PAUL CUEVA */
  @Ent_cNombres,@Ent_cflagDomiciliado, @Ent_cAconvenio, @Id_Pais, @Id_Vinculo_Economico, @Id_Convenio, @PorcentajeSunat )/* NUEVOS REGISTROS 02/07/2013 - PAUL CUEVA */    
END

IF @Accion = 'EDITAR'    
BEGIN    
 UPDATE CNM_ENTIDAD WITH(ROWLOCK)    
 SET Ent_cPersona=@Ent_cPersona, Ent_cDireccion=@Ent_cDireccion, Ent_nRuc=@Ent_nRuc,
  Ent_cRepresentante=@Ent_cRepresentante, Ent_cTipoDoc=@Ent_cTipoDoc, Ent_cEstado='A',     
  Ent_cFlagPersona=@Ent_cFlagPersona, Ent_cEstadoEntidad='1',     
  Ent_cUserModifica=@Ent_cUserCrea, Ent_dFechaModifica=getdate(), Ent_cEquipoUser = host_name(), 
  Ent_cApaterno=@Ent_cApaterno , Ent_cAmaterno=@Ent_cAmaterno, Id_Pais = @Id_Pais, Id_Vinculo_Economico = @Id_Vinculo_Economico, Ent_cNombres=@Ent_cNombres,/* NUEVOS REGISTROS 02/07/2013 - PAUL CUEVA */   
  Ent_cFlagDomiciliado=@Ent_cFlagDomiciliado, Ent_cAconvenio=@Ent_cAconvenio,/* NUEVOS REGISTROS 02/07/2013 - PAUL CUEVA */   
  Ent_cDeleted=CASE WHEN Ent_cDeleted<>'*' THEN '' ELSE '*' END , Id_Convenio = @Id_Convenio, PorcentajeSunat = @PorcentajeSunat   
 WHERE Emp_cCodigo = @Emp_cCodigo AND Ent_cCodEntidad = @Ent_cCodEntidad     
  AND Ten_cTipoEntidad = @Ten_cTipoEntidad    
END    
    
-- exec spCn_GrabaEntidad 'ELIMINAR', '001', '00001', 'C', '', '', '', '', '   ', ' ', ' ', 'I', 'ORTIZ'    
    
IF @Accion = 'ELIMINAR'    
BEGIN    
 DELETE FROM CNM_ENT_DIRECCION WITH(ROWLOCK)    
 WHERE Emp_cCodigo = @Emp_cCodigo AND Ent_cCodEntidad = @Ent_cCodEntidad     
  AND Ten_cTipoEntidad = @Ten_cTipoEntidad    
    
    
 DELETE FROM CNM_ENT_CONTACTO WITH(ROWLOCK)    
 WHERE Emp_cCodigo = @Emp_cCodigo AND Ent_cCodEntidad = @Ent_cCodEntidad     
  AND Ten_cTipoEntidad = @Ten_cTipoEntidad    
    
 DELETE FROM CNM_ENTIDAD WITH(ROWLOCK)    
 WHERE Emp_cCodigo = @Emp_cCodigo AND Ent_cCodEntidad = @Ent_cCodEntidad     
  AND Ten_cTipoEntidad = @Ten_cTipoEntidad    
END    
    
IF @Accion = 'SEL_REG'    
BEGIN    
 SELECT ENT.Emp_cCodigo, ENT.Ent_cCodEntidad, ENT.Ten_cTipoEntidad, ENT.Ent_cPersona,     
  ENT.Ent_cDireccion, ENT.Ent_nRuc, ENT.Ent_cRepresentante, ENT.Ent_cTipoDoc, ENT.Ent_cEstado,     
  TE.Ten_cNombreEntidad, EMP.Emp_cNombreLargo, EMP.Emp_cNombreCorto, EMP.Emp_cNumRuc,     
  ENT.Ent_cFlagPersona, ENT.Ent_cEstadoEntidad, ENT.Ent_cApaterno, ENT.Ent_cAmaterno, ENT.Ent_cNombres,/* NUEVOS REGISTROS 02/07/2013 - PAUL CUEVA */   
  ENT.Ent_cFlagDomiciliado, ENT.Ent_cAconvenio, '' AS EstadoSunat,  --EE.Tab_cDescripCampo AS EstadoSunat,     
  TD.Tab_cDescripCampo AS TipoDocumento, ISNULL(Id_Pais, '') AS Id_Pais, ISNULL(Id_Vinculo_Economico, '') AS Id_Vinculo_Economico, ISNULL(Id_Convenio, '00') AS Id_Convenio,
  ISNULL(PorcentajeSunat, 0) AS PorcentajeSunat
 FROM CNM_ENTIDAD ENT WITH(NOLOCK)    
 INNER JOIN EMPRESA EMP WITH(NOLOCK) ON ENT.Emp_cCodigo = EMP.Emp_cCodigo     
 LEFT OUTER JOIN TABLA TD WITH(NOLOCK) ON ENT.Ent_cTipoDoc = TD.Tab_cCodigo AND ENT.Emp_cCodigo = TD.Emp_cCodigo And TD.Tab_cTabla = '003'    
-- LEFT OUTER JOIN TABLA EE WITH(NOLOCK) ON ENT.Ent_cEstadoEntidad = EE.Tab_cCodigo AND ENT.Emp_cCodigo = EE.Emp_cCodigo  And EE.Tab_cTabla = '018'     
 LEFT OUTER JOIN CNT_ENTIDAD TE WITH(NOLOCK) ON ENT.Emp_cCodigo = TE.Emp_cCodigo AND ENT.Ten_cTipoEntidad = TE.Ten_cTipoEntidad    
 WHERE ENT.Emp_cCodigo = @Emp_cCodigo AND ENT.Ent_cCodEntidad = @Ent_cCodEntidad AND ENT.Ten_cTipoEntidad = @Ten_cTipoEntidad    
END    
    
 -- spCn_GrabaEntidad 'SEL_ALL', '022', '', 'C', '', '', '', '', '', '', '', '', ''     
IF @Accion = 'SEL_ALL'    
BEGIN    
 SELECT ENT.Emp_cCodigo, ENT.Ent_cCodEntidad, ENT.Ten_cTipoEntidad, /* NUEVOS REGISTROS 02/07/2013 - PAUL CUEVA */   
CASE TE.Ten_cPlame 
WHEN '1' THEN ENT.Ent_cApaterno+' '+ENT.Ent_cAmaterno+' '+ENT.Ent_cNombres
ELSE  ENT.Ent_cPersona
END as 'Ent_cPersona' /* FIN DE REGISTROS */
,     
  ENT.Ent_cDireccion, ENT.Ent_nRuc, ENT.Ent_cRepresentante, ENT.Ent_cTipoDoc,     
  ENT.Ent_cEstado,    
  TE.Ten_cNombreEntidad, EMP.Emp_cNombreLargo, EMP.Emp_cNombreCorto, EMP.Emp_cNumRuc,     
  ENT.Ent_cFlagPersona,     
  ENT.Ent_cEstadoEntidad, '' AS EstadoSunat,ENT.Ent_cApaterno, ENT.Ent_cAmaterno, ENT.Ent_cNombres,ENT.Ent_cFlagDomiciliado,/* NUEVOS REGISTROS 02/07/2013 - PAUL CUEVA */
  ENT.Ent_cAconvenio,/* NUEVOS REGISTROS 02/07/2013 - PAUL CUEVA */ 
  --EE.Tab_cDescripCampo AS EstadoSunat,
  TD.Tab_cDescripCampo AS TipoDocumento, GETDATE() AS FechaHoraImp 
 FROM CNM_ENTIDAD ENT WITH(NOLOCK)
 INNER JOIN EMPRESA EMP WITH(NOLOCK) ON ENT.Emp_cCodigo = EMP.Emp_cCodigo     
 LEFT  JOIN TABLA TD WITH(NOLOCK) ON TD.Tab_cTabla = '003' AND ENT.Ent_cTipoDoc = TD.Tab_cCodigo AND ENT.Emp_cCodigo = TD.Emp_cCodigo 
-- LEFT OUTER JOIN TABLA EE WITH(NOLOCK) ON EE.Tab_cTabla = '018' AND ENT.Ent_cEstadoEntidad = EE.Tab_cCodigo AND ENT.Emp_cCodigo = EE.Emp_cCodigo 
 LEFT  JOIN CNT_ENTIDAD TE WITH(NOLOCK) ON ENT.Emp_cCodigo = TE.Emp_cCodigo AND ENT.Ten_cTipoEntidad = TE.Ten_cTipoEntidad
 WHERE ENT.Emp_cCodigo = @Emp_cCodigo    
  AND ENT.Ten_cTipoEntidad = CASE ISNULL(@Ten_cTipoEntidad, '') WHEN '' THEN ENT.Ten_cTipoEntidad ELSE @Ten_cTipoEntidad END    
    
END     
-- SELECT * FROM CNM_ENTIDAD    
--  SELECT * FROM TABLA WHERE Tab_cTabla = '018'      
IF @Accion = 'SEL_PRINT'
BEGIN
 SELECT ENT.Emp_cCodigo, ENT.Ent_cCodEntidad, ENT.Ten_cTipoEntidad, ENT.Ent_cPersona,     
  ENT.Ent_cDireccion, ENT.Ent_nRuc, ENT.Ent_cRepresentante, ENT.Ent_cTipoDoc,    
  ENT.Ent_cEstado,     
  TE.Ten_cNombreEntidad, EMP.Emp_cNombreLargo, EMP.Emp_cNombreCorto, EMP.Emp_cNumRuc,     
  ENT.Ent_cFlagPersona,     
  ENT.Ent_cEstadoEntidad,ENT.Ent_cApaterno, ENT.Ent_cAmaterno, ENT.Ent_cNombres,/* NUEVOS REGISTROS 02/07/2013 - PAUL CUEVA */
  ENT.Ent_cFlagDomiciliado,ENT.Ent_cAconvenio,'' AS EstadoSunat,
  TD.Tab_cDescripCampo AS TipoDocumento, getdate() AS FechaHoraImp    
    
 FROM CNM_ENTIDAD ENT WITH(NOLOCK)    
 LEFT  JOIN EMPRESA EMP WITH(NOLOCK) ON ENT.Emp_cCodigo = EMP.Emp_cCodigo     
 LEFT  JOIN TABLA TD WITH(NOLOCK) ON TD.Tab_cTabla = '003'  AND ENT.Ent_cTipoDoc = TD.Tab_cCodigo AND ENT.Emp_cCodigo = TD.Emp_cCodigo     
-- LEFT  JOIN TABLA EE WITH(NOLOCK) ON EE.Tab_cTabla = '018'  AND  ENT.Ent_cEstadoEntidad = EE.Tab_cCodigo AND ENT.Emp_cCodigo = EE.Emp_cCodigo     
 LEFT  JOIN CNT_ENTIDAD TE WITH(NOLOCK) ON ENT.Emp_cCodigo = TE.Emp_cCodigo AND ENT.Ten_cTipoEntidad = TE.Ten_cTipoEntidad    
 WHERE ENT.Emp_cCodigo = @Emp_cCodigo AND     
  ENT.Ten_cTipoEntidad LIKE CASE WHEN DBO.TRIMSQL(@Ten_cTipoEntidad) <> 'x' THEN @Ten_cTipoEntidad ELSE '%' END     
END     
    
IF @Accion = 'SEL_RUC'    
BEGIN    
 SELECT Ent_cCodEntidad FROM CNM_ENTIDAD WITH(NOLOCK)    
 WHERE Emp_cCodigo = @Emp_cCodigo AND Ten_cTipoEntidad = @Ten_cTipoEntidad AND Ent_nRuc = @Ent_nRuc    
END    
    
IF @Accion = 'CORREL'    
BEGIN    
 SELECT @INTERNO = ISNULL(max(convert(int,Ent_cCodEntidad))+1, 1) FROM CNM_ENTIDAD WITH(NOLOCK)    
 WHERE Emp_cCodigo = @Emp_cCodigo AND Ten_cTipoEntidad = @Ten_cTipoEntidad  AND Ent_cCodEntidad<>'99999'  
     
 SET @INTERNO = REPLICATE('0', 5-len(@INTERNO)) + @INTERNO     
     
 SELECT @INTERNO    
END    
    
IF @Accion = 'SEL_MVTOS'     
BEGIN -- *** SELECCIONAR SI ENTIDAD TIENE MOVIMIENTOS    
 SELECT COUNT(*) FROM CND_ASIENTO_VOUCHER WITH(NOLOCK)    
 WHERE Emp_cCodigo = @Emp_cCodigo AND Ent_cCodEntidad = @Ent_cCodEntidad AND Ten_cTipoEntidad = @Ten_cTipoEntidad    
   and asd_cdeleted<>'*'    
END     
    
IF @Accion = 'BUSCARREGISTRO'    
BEGIN    
 SELECT * FROM CNM_ENTIDAD WITH(NOLOCK)    
 WHERE Emp_cCodigo = @Emp_cCodigo AND Ent_cCodEntidad = @Ent_cCodEntidad AND Ten_cTipoEntidad = @Ten_cTipoEntidad    
END     
    
IF @Accion = 'BUSCARTODOS'    
BEGIN    
 SELECT * FROM CNM_ENTIDAD WITH(NOLOCK)    
 WHERE Emp_cCodigo = @Emp_cCodigo    
END

DECLARE @Separador varchar(1)  
 SELECT @Separador = '|'
  
IF @Accion = 'BUSCARPLAME'  /* NUEVOS REGISTROS 02/07/2013 - PAUL CUEVA */     
BEGIN    
-- SELECT 
--(LTRIM(RTRIM(RIGHT('0' + t.Tab_cCodSunat, 2))) + @Separador +  
--   LTRIM(RTRIM(LEFT(Ent_nRuc,15))) + @Separador +  
--   LTRIM(RTRIM(LEFT(Ent_cApaterno,40))) + @Separador +     
--   LTRIM(RTRIM(LEFT(Ent_cAmaterno,40))) + @Separador +     
--   LTRIM(RTRIM(LEFT(Ent_cNombres,40))) + @Separador +   
--LTRIM(RTRIM(LEFT(Ent_cFlagDomiciliado,1))) + @Separador +   
--LTRIM(RTRIM(LEFT(case when Ent_cAconvenio <> '' then '1' else '0' end,1))) + @Separador )   
--AS Registros
-- FROM CNM_ENTIDAD Ent,CNT_ENTIDAD Ten, dbo.TABLA T
--WITH(NOLOCK)    
-- WHERE Ent.Ten_cTipoEntidad=Ten.Ten_cTipoEntidad AND Ent.Emp_cCodigo = Ten.Emp_cCodigo AND
-- Ten.Ten_cPlame='1' AND T.Emp_cCodigo = Ent.Emp_cCodigo AND T.Tab_cCodigo = Ent.Ent_cTipoDoc AND T.Tab_cTabla = '003' AND
-- Ent.Emp_cCodigo = @Emp_cCodigo   

	SELECT RIGHT('0' + t.Tab_cCodSunat, 2) + @Separador + 
        LEFT(ce.Ent_nRuc, 15) + @Separador + LEFT(ce.Ent_cApaterno, 40) + @Separador +  LEFT(ce.Ent_cAmaterno, 40) + @Separador + LEFT(ce.Ent_cNombres, 40) + @Separador + ISNULL(ce.Ent_cFlagDomiciliado, 0) + 
        @Separador +  CASE WHEN ce.Ent_cAconvenio <> '' THEN '1' ELSE '0' END + @Separador AS Registros FROM dbo.CNM_ENTIDAD CE INNER JOIN dbo.CNT_ENTIDAD CE2 
        ON CE.Emp_cCodigo = CE2.Emp_cCodigo AND CE.Ten_cTipoEntidad = CE2.Ten_cTipoEntidad
 INNER JOIN dbo.CNT_EXPORTA_PDT CEP ON CE.Emp_cCodigo = CEP.Emp_cCodigo AND CE.Ent_nRuc = CEP.Ent_nRuc INNER JOIN dbo.TABLA T ON CEP.Emp_cCodigo = T.Emp_cCodigo AND CE.Ent_cTipoDoc = T.Tab_cCodigo
 AND T.Tab_cTabla = '003'
 WHERE CE.Emp_cCodigo = @Emp_cCodigo AND CE2.Ten_cPlame = '1' AND CEP.Per_cPeriodo = @Per_cPeriodo

END						/* FIN DE REGISTROS */                                     

EXEC USP_Separar_Nombre_Entidad

GO
