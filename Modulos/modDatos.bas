Attribute VB_Name = "modDatos"
Option Explicit

Public Function BuscaConfigAnual(cCampo As String) As String
    BuscaConfigAnual = fRetornaValor("spCFG_PARAMETROS 'BUSCAR_CONTA', '" & gsEmpresa & "', '" & gsAnio & "', '" & cCampo & "'")
End Function

Public Function BuscaNombreCuenta(ByVal cCodigo As String, ByRef bExiste As Boolean)
Dim gtxtSQL As String
Dim rsCta As New ADODB.Recordset
    BuscaNombreCuenta = ""
    cCodigo = CE(cCodigo)
    If cCodigo <> "" Then
        gtxtSQL = "spCn_ConsultaCuentas 'SEL_REG_ALL', '" & gsEmpresa & "', '" & gsAnio & "', '" & cCodigo & "'"
        Set rsCta = fRetornaRS(gtxtSQL)
        If GetRsRecordCount(rsCta) > 0 Then
            BuscaNombreCuenta = CE(rsCta!Pla_cNombreCuenta)
            bExiste = True
        Else
            'Mensajes "Cuenta Contable no existe"
            bExiste = False
        End If
        Call CerrarRecordSet(rsCta)
    End If
End Function

Public Function BuscarCtaContable(ByRef oFormCall As Form, Optional ByVal sCadBusq As String = "") As String
Dim gtxtSQL As String
On Error Resume Next
    BuscarCtaContable = ""
    gsCodigo = ""
    If grstBusqueda Is Nothing Then
        gtxtSQL = "spCn_ConsultaCuentas 'SEL_ALL_NOTITULO', '" & gsEmpresa & "', '" & gsAnio & "', ''"
        Set grstBusqueda = fRetornaRS(gtxtSQL)
        If GetRsRecordCount(grstBusqueda) > 0 Then
            oFormCall.Enabled = False
            
            frmBusqueda.Caption = "Busqueda de Cuentas Contables"
            frmBusqueda.tdbgListado.Columns(0).Caption = "Nro.Cuenta"
            frmBusqueda.tdbgListado.Columns(0).DataField = "Pla_cCuentaContable"
            frmBusqueda.tdbgListado.Columns(0).Width = 1500
            frmBusqueda.tdbgListado.Columns(1).Caption = "Descripción de Cuenta"
            frmBusqueda.tdbgListado.Columns(1).DataField = "Pla_cNombreCuenta"
            frmBusqueda.tdbgListado.Columns(1).Width = 4500
            frmBusqueda.tdbgListado.Columns(2).Caption = "Referencia"
            frmBusqueda.tdbgListado.Columns(2).DataField = ""
            frmBusqueda.tdbgListado.Columns(2).Width = 1500
            frmBusqueda.tdbgListado.Columns(3).Caption = ""
            frmBusqueda.tdbgListado.Columns(3).DataField = ""
            frmBusqueda.tdbgListado.Columns(3).Width = 0
            frmBusqueda.tdbgListado.Columns(3).Visible = False
            frmBusqueda.tdbgListado.Columns(4).Width = 0
            frmBusqueda.tdbgListado.Columns(4).Visible = False
            frmBusqueda.tdbgListado.Columns(5).Width = 0
            frmBusqueda.tdbgListado.Columns(5).Visible = False
            frmBusqueda.tdbgListado.DataSource = grstBusqueda
            frmBusqueda.tdbgListado.Columns(0).FilterText = sCadBusq
            frmBusqueda.Show 1
            oFormCall.Enabled = True
        Else
            Mensajes "No existen Cuentas contables en el Plan de Cuentas."
        End If
        BuscarCtaContable = gsCodigo
        Call CerrarRecordSet(grstBusqueda)
        If gsCodigo <> "" Then Call EnterTab(vbKeyReturn)
    Else
        Mensajes "Esta realizando una busqueda desde otro formulario"
    End If
End Function

Public Function GrabaConfigAnual(cCampo As String, cValor As String) As Boolean
    GrabaConfigAnual = False

    Dim clsMante As clsMantoTablas
    Set clsMante = New clsMantoTablas
    Screen.MousePointer = vbHourglass
    Dim lArrMnt(4) As Variant
    lArrMnt(0) = "EDITAR_CONTA"
    lArrMnt(1) = gsEmpresa
    lArrMnt(2) = gsAnio
    lArrMnt(3) = cCampo
    lArrMnt(4) = cValor
    
    If Not clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCFG_PARAMETROS", lArrMnt(), True) Then
        Mensajes "El proceso no se ha realizado. Verificar...", vbInformation
        Screen.MousePointer = vbDefault
        Set clsMante = Nothing
        Exit Function
    End If
    
    Set clsMante = Nothing
    Screen.MousePointer = vbNormal
    
    GrabaConfigAnual = True
    
End Function


Public Sub GrabaCondigOPDet(cCodigoOP As String, cValor As String)
    Dim clsMante As clsMantoTablas
    Set clsMante = New clsMantoTablas
    ' *** Eliminando la Cuenta
    Screen.MousePointer = vbHourglass
    Dim lArrMnt(9) As Variant
    lArrMnt(0) = "EDITAR_MIX"
    lArrMnt(1) = gsEmpresa
    lArrMnt(2) = gsAnio
    lArrMnt(3) = cCodigoOP '"027" 'UIT
    lArrMnt(4) = cValor 'tdbnUIT.Value
    lArrMnt(5) = 0
    lArrMnt(6) = Null
    lArrMnt(7) = gsUsuario
    lArrMnt(8) = gsUsuario
    lArrMnt(9) = Null
    
    If Not clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCND_CONFIG_OPERA", lArrMnt(), True) Then
        Mensajes "El proceso no se ha realizado. Verificar...", vbInformation
        Screen.MousePointer = vbDefault
    End If
    
    Set clsMante = Nothing
    Screen.MousePointer = vbNormal
    
End Sub

Public Sub pCargaCfgCtas()

    Dim sqlver As String
    Dim rsCta As New ADODB.Recordset
    
    sqlver = "SELECT Pla_cCuentaContable FROM CNM_PLAN_CTA WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' AND pla_cDifCambio = 'G'"
    LlenarRecordSet sqlver, rsCta
    If Not rsCta Is Nothing And rsCta.State = adStateOpen Then gsCuentaDifGan = CE(rsCta!Pla_cCuentaContable)
    
    sqlver = "SELECT Pla_cCuentaContable FROM CNM_PLAN_CTA WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' AND pla_cDifCambio = 'P'"
    LlenarRecordSet sqlver, rsCta
    If Not rsCta Is Nothing And rsCta.State = adStateOpen Then gsCuentaDifPer = CE(rsCta!Pla_cCuentaContable)
    
    sqlver = "SELECT Pla_cCuentaContable FROM CNM_PLAN_CTA WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' AND Pla_cRedondeo = 'G'"
    LlenarRecordSet sqlver, rsCta
    If Not rsCta Is Nothing And rsCta.State = adStateOpen Then gsCuentaRedGan = CE(rsCta!Pla_cCuentaContable)
    
    sqlver = "SELECT Pla_cCuentaContable FROM CNM_PLAN_CTA WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' AND Pla_cRedondeo = 'P'"
    LlenarRecordSet sqlver, rsCta
    If Not rsCta Is Nothing And rsCta.State = adStateOpen Then gsCuentaRedPer = CE(rsCta!Pla_cCuentaContable)

    CerrarRecordSet rsCta
    
    
End Sub


Public Function ConsultarDatosRs(sqlDatos As String) As ADODB.Recordset
    Dim rsDatos As ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
    
    Set clDatos = New clsMantoTablas
    arrDatos = Array(sqlDatos)
    Set rsDatos = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    Set ConsultarDatosRs = rsDatos
    'Call CerrarRecordSet(rsDatos)
End Function

Public Function FechaServidor() As Date
    Dim rsFecha As New ADODB.Recordset
    
    Conectar
    rsFecha.Open " SELECT convert(char(10), GETDATE(), 103) as fecha", gcnSistema
    'FechaServidor = rsFecha("fecha") 'Format(Date, "dd/MM/yyyy") ' -- PGBV Se comenta esta linea para que tome la fecha de la PC y no del Servidor
    FechaServidor = Format(Date, "dd/MM/yyyy")
    Call CerrarRecordSet(rsFecha)
    Desconectar
End Function

Public Function TipoCambio(fecha As Variant, cOperaTC As String, Optional sMoneda As String = "") As Double
    If Not IsDate(fecha) Then Exit Function
    
    Dim rsTc As New ADODB.Recordset
    TipoCambio = 0
    Call Conectar
    
    'Si no esta activado el control chkBimoneda de la empresa
    If sMoneda = "" And gintBiMoneda = 0 Then
        sMoneda = gsMonedaExt
    End If
    
    If cOperaTC = "PRO" Then
        Set rsTc = gcnSistema.Execute("spCn_HallarTcMensual '" & gsEmpresa & "','" & gsAnio & "','" & gsPeriodo & "','" & sMoneda & "','0','PRO'")
        If Not rsTc.EOF And Not rsTc.BOF Then
            TipoCambio = NE(rsTc("TC").Value)
        Else
            TipoCambio = -1
        End If
            
    Else
        Set rsTc = gcnSistema.Execute("spCn_HallarTC '" & gsEmpresa & "', '" & fecha & "', '" & sMoneda & "'")
        If Not rsTc.EOF And Not rsTc.BOF Then
           Select Case Trim(cOperaTC)
                  Case "COM"
                       TipoCambio = NE(rsTc("Tca_nCompra"))
                  Case "COP"
                       TipoCambio = NE(rsTc("Tca_nCompraP"))
                  Case "VEN"
                       TipoCambio = NE(rsTc("Tca_nVenta"))
                  Case "VEP"
                       TipoCambio = NE(rsTc("Tca_nVentaP"))
           End Select
        Else
            TipoCambio = -1
        End If
    End If
    
    Call CerrarRecordSet(rsTc)
    Call Desconectar
   
End Function

Public Function correlativoCodigoEnt(Tipo As String) As String
    Dim sqlSp As String
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
    
    Set clDatos = New clsMantoTablas
    Dim rsCodigo As New ADODB.Recordset
    correlativoCodigoEnt = "00001"
    sqlSp = "spCn_GrabaEntidad 'CORREL', '" & gsEmpresa & "', '', '" & Tipo & "', '', '', '', '', '', '', '', '', '' "
    arrDatos = Array(sqlSp)
    Set rsCodigo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If Not rsCodigo Is Nothing Then
        correlativoCodigoEnt = rsCodigo(0).Value
    End If
    Call CerrarRecordSet(rsCodigo)
    Set clDatos = Nothing
End Function

Public Function correlativoCtaRatio() As String
    Dim sqlSp As String
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
    
    Set clDatos = New clsMantoTablas
    Dim rsCodigo As New ADODB.Recordset
    correlativoCtaRatio = "0001"
    sqlSp = "spCn_GrabaCuentaRatio 'CORREL', '" & gsEmpresa & "', '', '', '', '' "
    arrDatos = Array(sqlSp)
    Set rsCodigo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If Not rsCodigo Is Nothing Then
        correlativoCtaRatio = rsCodigo(0).Value
    End If
    Call CerrarRecordSet(rsCodigo)
    Set clDatos = Nothing
End Function

Public Function numeroVoucher(Tipo As String, año As String, periodo As String, Libro As String) As String
On Error GoTo serror
    Dim sqlSp As String
    Dim arrDatos() As Variant
    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    
    Set clDatos = New clsMantoTablas
    sqlSp = "spCn_HallarVoucher '" & Tipo & "', '" & gsEmpresa & "', '" & año & "', '" & periodo & "', '" & Libro & "' "
    arrDatos = Array(sqlSp)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    numeroVoucher = CE(rsArreglo(0).Value)
    Call CerrarRecordSet(rsArreglo)
    Set clDatos = Nothing
    Exit Function
serror:
    Mensajes Err.Description
End Function

Public Function numeroEmpresa() As String
    Dim sqlSp As String
    Dim arrDatos() As Variant
    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    
    Set clDatos = New clsMantoTablas
    sqlSp = "spCn_GrabaEmpresa 'CORRELATIVO', '', '', '', '', '', '', '' "
    arrDatos = Array(sqlSp)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    numeroEmpresa = rsArreglo(0).Value
    Call CerrarRecordSet(rsArreglo)
    Set clDatos = Nothing
End Function

Public Function EjecutaQuery(sql As String) As Integer
    On Error GoTo ERROR
    Dim Afectados As Integer
    Conectar
    
    gcnSistema.Execute sql, Afectados
    
    Desconectar
    EjecutaQuery = Afectados
    Exit Function
ERROR:
    EjecutaQuery = -1
End Function


Public Function ExtraeDescripcion(sql As String) As String
    Dim rsDatos As ADODB.Recordset
    
    LlenarRecordSet sql, rsDatos
    
    ExtraeDescripcion = ""
    'Conectar
    'rsDatos.Open Sql, gcnSistema
    If Not rsDatos Is Nothing Then
        If Not (rsDatos.EOF And rsDatos.BOF) Then
            ExtraeDescripcion = CE(rsDatos(0).Value)
        End If
    End If
    Call CerrarRecordSet(rsDatos)
    'Desconectar
End Function

Public Function CuentaCCosto(Cuenta As String, año As String) As Boolean
    Dim rsCosto As New ADODB.Recordset
    Dim sqlver As String
    
    sqlver = "SELECT ISNULL(count(Pla_cCuentaContable), 0) From CNM_PLAN_CTA WHERE Emp_cCodigo = '" & gsEmpresa & "' " & _
             "AND Pla_cCuentaContable = '" & Cuenta & "' AND Pla_cTitulo = 'N' " & _
             "AND Pla_cDeleted <> '*' AND Pla_cCentroCosto = '1' " & _
             "AND Pan_cAnio = '" & año & "' "
    Call Conectar
    rsCosto.Open sqlver, gcnSistema
    If rsCosto(0) = 0 Then
        CuentaCCosto = False
    Else
        CuentaCCosto = True
    End If
    Call CerrarRecordSet(rsCosto)
    Call Desconectar
End Function

Public Function CuentaEntidad(Cuenta As String, año As String) As String
    Dim rsCosto As New ADODB.Recordset
    Dim sqlver As String
    
    CuentaEntidad = ""
    sqlver = "SELECT Ten_CTipoEntidad From CNM_PLAN_CTA WHERE Emp_cCodigo = '" & gsEmpresa & "' " & _
             "AND Pla_cCuentaContable = '" & Cuenta & "' AND Pla_cTitulo = 'N' " & _
             "AND Pla_cDeleted <> '*' AND Pan_cAnio = '" & año & "' "
    Call Conectar
    rsCosto.Open sqlver, gcnSistema
    If Not rsCosto.EOF And Not rsCosto.BOF Then CuentaEntidad = Trim(rsCosto("Ten_CTipoEntidad").Value)
    Call CerrarRecordSet(rsCosto)
    Call Desconectar
End Function

Public Function ProvisionCancelada(Numero As Single) As Boolean
    Dim sqlProv As String
    Dim rsProv As New ADODB.Recordset
    
    ProvisionCancelada = False
    sqlProv = "SELECT Ase_cNummov, Ase_nVoucher, Cnp_nMonSolCancel, Cnp_nMonExtCancel FROM CND_ASIENTO_PROV WHERE Cnp_nCorre = '" & Numero & "'"
    Call Conectar
    rsProv.Open sqlProv, gcnSistema
    If Not rsProv.EOF And Not rsProv.BOF Then
        If rsProv("Cnp_nMonSolCancel") = 0 Or rsProv("Cnp_nMonExtCancel") = 0 Then
            ProvisionCancelada = False
        Else
            ProvisionCancelada = True
        End If
    End If
    Call CerrarRecordSet(rsProv)
    Call Desconectar
End Function

Public Function ExisteDocumentoProvi(Tipo As String, ent As String, td As String, _
                                     SerieDoc As String, NumDoc As String) As Boolean
    
    'FUNCION UTILIZADA EN : LLENA TIPO DE ASIENTO (llenaTipoAsiento)
    Dim rsDetAsiento As New ADODB.Recordset
    Dim sqlver As String
    
    ExisteDocumentoProvi = False
    Call Conectar
    
    'BUSCA EN NUMERO INTERNO DEL VOUCHER DEL DETALLE
    sqlver = "SELECT Ase_cNumMov, Ase_nVoucher, Asd_nItem, Asd_cTipoMoneda  " & _
             "From CND_ASIENTO_VOUCHER " & _
             "WHERE Emp_cCodigo = '" & gsEmpresa & "' " & _
             "AND Ten_cTipoEntidad = '" & Tipo & "' " & _
             "AND Ent_cCodEntidad = '" & ent & "' " & _
             "AND Asd_cTipoDoc = '" & td & "' AND Asd_cSerieDoc  = '" & SerieDoc & "' " & _
             "AND Asd_cNumDoc = '" & NumDoc & "' AND Asd_cDeleted <> '*' AND Asd_cProvCanc='P' " & _
             "AND Pan_cAnio = '" & gsAnio & "'"
    
    Set rsDetAsiento = gcnSistema.Execute(sqlver)
    
    If Not rsDetAsiento.EOF And Not rsDetAsiento.BOF Then
        ExisteDocumentoProvi = True
    End If
    
    Call CerrarRecordSet(rsDetAsiento)
    Call Desconectar
    
End Function

Public Function ExisteDocumentoProviEntidad(Nummov As String, Voucher As String, correl As Double) As String
    Dim rsDetAsiento As New ADODB.Recordset
    Dim sqlver As String
    
    'FUNCION UTILIZADA EN : CARGA DATOS DE REGSTRO (CargaDatosRegistro)
    ExisteDocumentoProviEntidad = ""
    
    sqlver = "select Pan_cAnio, Per_cPeriodo, Lib_cTipoLibro, Ase_nVoucher " & _
             "from CND_ASIENTO_PROV WITH(READUNCOMMITTED) where Emp_cCodigo = '" & gsEmpresa & "' AND  " & _
             "Ase_cNummov = '" & Nummov & "' AND Pan_cAnio = '" & gsAnio & "' AND  ase_nVoucher='" & Voucher & "' and " & _
             "( Cnp_nMonSolCancel > 0 or Cnp_nMonExtCancel > 0 ) and cnp_ncorre=" & correl
    
    Call Conectar
    rsDetAsiento.Open sqlver, gcnSistema
    If Not rsDetAsiento.EOF And Not rsDetAsiento.BOF Then
       ExisteDocumentoProviEntidad = rsDetAsiento("Ase_nVoucher") & " DEL LIBRO " & rsDetAsiento("Lib_cTipoLibro") & " - " & rsDetAsiento("Per_cPeriodo") & "/" & rsDetAsiento("Pan_cAnio")
    End If
    
    Call CerrarRecordSet(rsDetAsiento)
    Call Desconectar
End Function

Public Function NroCorre_Documento_Provisionado(Tipo As String, ent As String, td As String, _
                SerieDoc As String, NumDoc As String, fecha As String, ByRef correl As Double, ByRef Cuenta As String) As String
    Dim rsDetAsiento As New ADODB.Recordset, sqlver As String
    
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
    Set clDatos = New clsMantoTablas
   
    NroCorre_Documento_Provisionado = ""
             
    sqlver = "spCn_ConsultaProvisionesDoc 'BUSCA_SI_PROV_DOC', '" & gsEmpresa & "','" & gsAnio & "', '" & ent & "', NULL , '" & td & "', '" & SerieDoc & "', '" & NumDoc & "', '" & Cuenta & "', '" & Tipo & "', '" & gsPeriodo & "'"
    arrDatos = Array(sqlver)
    Set rsDetAsiento = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    
    If Not rsDetAsiento Is Nothing Then
        If rsDetAsiento.State = adStateOpen Then
            If Not rsDetAsiento.EOF And Not rsDetAsiento.BOF Then
               NroCorre_Documento_Provisionado = CE(rsDetAsiento("Ase_nVoucher")) & " DEL LIBRO " & CE(rsDetAsiento("Lib_cTipoLibro")) & " - " & CE(rsDetAsiento("Asd_dFecDoc"))
               correl = NE(rsDetAsiento("Asd_nCorre"))
            End If
        End If
    End If
    
    Call CerrarRecordSet(rsDetAsiento)
    Set clDatos = Nothing
End Function

Public Function BuscarCadRs(cadenas As String, rs As Recordset, Col As Integer) As Integer
    On Error GoTo serror
    BuscarCadRs = -1
    If rs Is Nothing Then Exit Function
    If rs.BOF And rs.EOF Then Exit Function
    
    Dim Fila As Integer
    BuscarCadRs = 0
    Fila = 0
    rs.MoveFirst
    Do While Not rs.EOF
        If Trim(cadenas) = Trim(rs(Col).Value) Then
            BuscarCadRs = Fila
            Exit Function
        End If
        rs.MoveNext
        Fila = Fila + 1
    Loop
    
    Exit Function
serror:
    BuscarCadRs = -1
End Function

Public Function UltimoDiaMes(Mes As String, año As String) As Date
    If NE(Mes) < 1 Then
        UltimoDiaMes = "31/01/" & año
    ElseIf NE(Mes) > 12 Then
        UltimoDiaMes = "31/12/" & año
    Else
        UltimoDiaMes = DateSerial(NE(año), NE(Mes) + 1, 0)
    End If
End Function

Public Function PrimerDiaMes(Mes As String, año As String) As Date
    If NE(Mes) < 1 Then
        PrimerDiaMes = "01/01/" & año
    ElseIf NE(Mes) > 12 Then
        PrimerDiaMes = "01/12/" & año
    Else
        PrimerDiaMes = "01/" & Mes & "/" & año
    End If
End Function

Public Function CuentaCfgAuto(Valor As String) As String

    ' *** Verificar q codigo exista
    
    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
    ' *** Cargando Datos de la Cuenta
    Dim sqlSp As String
    
    CuentaCfgAuto = ""
    
    Set clDatos = New clsMantoTablas
    sqlSp = "spCn_ConsultaCuentas '" & Valor & "', '" & gsEmpresa & "', '" & gsAnio & "', ''"
    arrDatos = Array(sqlSp)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo.State <> 0 Then
        CuentaCfgAuto = rsArreglo("Pla_cCuentaContable").Value
    End If
    Call CerrarRecordSet(rsArreglo)
    ' ***
End Function

Public Function monedaNacional(Codigo As String) As Boolean

    If Codigo = gsMonedaNac Then
       monedaNacional = True
    Else
       monedaNacional = False
    End If

End Function

Public Function rutaReportes() As String
    rutaReportes = App.Path & "\Reportes\" 'BuscaRutaReportes("CONTABILIDAD")
End Function

Public Function Fct_Devolver_Num_Dig_Cta_Detalle() As Long
    Dim VarCnx As New ADODB.Connection
    VarCnx.Open gsCadenaConexion
    
    Dim VarCmd As New ADODB.Command
    VarCmd.CommandText = ""
End Function

