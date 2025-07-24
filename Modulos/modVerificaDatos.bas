Attribute VB_Name = "ModVerificaDatos"
Option Explicit

Public Function BuscaTamanioDoc(Codigo As String) As String
    Dim clDatos As New clsMantoTablas
    Dim rsArreglo As ADODB.Recordset
    Dim arrDatos() As Variant, sqlver As String

    sqlver = "spMAN_TABLA 'SEL_REG', '" & gsEmpresa & "', '003', '" & Codigo & "', '', '', 0 "
    arrDatos = Array(sqlver)
    
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    
    If Not rsArreglo Is Nothing Then
        BuscaTamanioDoc = CE(rsArreglo!Tab_nLongitud)
    Else
        BuscaTamanioDoc = "0"
    End If
    
    Call CerrarRecordSet(rsArreglo)
    Set clDatos = Nothing
End Function

Public Function ExisteDato(Sql As String) As Boolean
    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As New clsMantoTablas
    Dim arrDatos() As Variant
    
    ExisteDato = False
    arrDatos = Array(Sql)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo.State = 0 Then
        On Error Resume Next
        Call CerrarRecordSet(rsArreglo)
        Set clDatos = Nothing
        Exit Function
    End If
    If rsArreglo(0) > 0 Then ExisteDato = True
    Call CerrarRecordSet(rsArreglo)
    Set clDatos = Nothing
End Function

Public Function ExisteRegistro(Valor As String, sp As String, año As Boolean) As Boolean
    ' *** Verificar q codigo exista dependiendo de la tabla y el SP q se le envie
    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
    ' *** Cargando Datos de la Cuenta
    Dim sqlSp As String
    
    ExisteRegistro = False
    Set clDatos = New clsMantoTablas
    
    sqlSp = sp + " 'SEL_REG', '" & gsEmpresa & "', "
    
    If año = True Then sqlSp = sqlSp + " '" & gsAnio & "', "
    sqlSp = sqlSp + " '" & Valor & "'  "
    'sqlSp = "spCn_GrabaLibroOpera 'SEL_REG', '" & gsEmpresa & "', '" & gsAnio & "', '" & valor & "', '', '', '', '', '' "
    'sqlSp = "spCn_GrabaTipoMoneda 'SEL_REG', '" & gsEmpresa & "', '" & valor & "', '', '', '', '', '', '' "
    arrDatos = Array(sqlSp)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo.State <> 0 Then
        ExisteRegistro = True
    End If
    Call CerrarRecordSet(rsArreglo)
    Set clDatos = Nothing
    ' ***
End Function

Public Function ExisteCta(valorCta As String) As String
    If CE(valorCta) = "" Then
        ExisteCta = ""
        Exit Function
    End If

    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
    
    ExisteCta = ""
    Dim sqlSp As String
        
    Set clDatos = New clsMantoTablas
    sqlSp = "spCn_ConsultaCuentas 'SEL_REG_ALL', '" & gsEmpresa & "', '" & gsAnio & "', '" & valorCta & "'"
    arrDatos = Array(sqlSp)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo.State = 0 Then
        ExisteCta = ""
        
    Else
        ExisteCta = CE(rsArreglo!Pla_cNombreCuenta)
    End If
    Call CerrarRecordSet(rsArreglo)
    Set clDatos = Nothing
End Function

Public Function ExisteCtaNoTitulo(valorCta As String, Tipo As String) As String
    If CE(valorCta) = "" Then
        ExisteCtaNoTitulo = ""
        Exit Function
    End If

    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
    
    ' *** Verificando q exista la cuenta contable
    ExisteCtaNoTitulo = ""
    Dim sqlSp As String
        
    Set clDatos = New clsMantoTablas
    sqlSp = "spCn_ConsultaCuentas 'SEL_REG_ALL', '" & gsEmpresa & "', '" & gsAnio & "', '" & valorCta & "'"
    arrDatos = Array(sqlSp)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo.State = 0 Then
        Mensajes "Codigo no existe... ", vbInformation
        
        ExisteCtaNoTitulo = ""
        
    Else
        If rsArreglo!Pla_cTitulo = "S" Then
            If Tipo = "N" Then
                Mensajes "Cuenta es de titulo. Verifique... ", vbInformation
                ExisteCtaNoTitulo = ""
            Else
                ExisteCtaNoTitulo = rsArreglo!Pla_cNombreCuenta
            End If
        Else
            ExisteCtaNoTitulo = rsArreglo!Pla_cNombreCuenta
        End If
    End If
    Call CerrarRecordSet(rsArreglo)
    Set clDatos = Nothing
End Function

Public Function BuscaCtasLibro(valorLibro As String) As Boolean
    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
    
    ' *** Verificando q exista la cuenta contable con el libro
    
    Dim sqlSp As String
        
    Set clDatos = New clsMantoTablas
    sqlSp = "spCn_ConsultaCuentasLibro 'SEL_CTALIBRO' , '" & gsEmpresa & "', '" & gsAnio & "','" & Left(valorLibro, 2) & "',''"
    
    arrDatos = Array(sqlSp)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo.State = adStateOpen Then
       If rsArreglo.RecordCount > 0 Then
            If NE(rsArreglo!registros) > 0 Then
                BuscaCtasLibro = True
            Else
                BuscaCtasLibro = False
            End If
        Else
            BuscaCtasLibro = False
        End If
    Else
       BuscaCtasLibro = False
    End If
    Call CerrarRecordSet(rsArreglo)
    Set clDatos = Nothing
End Function

Public Function ExisteCtaLibroCuenta(valorLibro As String, valorCta As String) As String
    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
    
    ' *** Verificando q exista la cuenta contable con el libro
    ExisteCtaLibroCuenta = ""
    Dim sqlSp As String
        
    Set clDatos = New clsMantoTablas
    sqlSp = "spCn_ConsultaLibroCuenta '" & valorLibro & "', '" & valorCta & "', '" & gsEmpresa & "', '" & gsAnio & "'"
    
    arrDatos = Array(sqlSp)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo.State = 0 Then
       ExisteCtaLibroCuenta = ""
    Else
       ExisteCtaLibroCuenta = CE(rsArreglo!Pla_cCuentaContable)
    End If
    Call CerrarRecordSet(rsArreglo)
    Set clDatos = Nothing
End Function


Public Function ExisteTipoAsi(Libro As String) As Boolean
    Dim sqlSp As String
    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As New clsMantoTablas
    Dim arrDatos() As Variant
    
    ExisteTipoAsi = False
    
    sqlSp = "spCn_GrabaTipoAsiento 'EXISTE_ASILIB', '" & gsEmpresa & "', '" & gsAnio & "', '" & Libro & "', '', 0, '', '', '', 0, '', '' "
    arrDatos = Array(sqlSp)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo.State = 0 Then Exit Function
    If rsArreglo(0) > 0 Then ExisteTipoAsi = True
    Call CerrarRecordSet(rsArreglo)
    Set clDatos = Nothing
End Function

Public Function CodigoEntidadRuc(Ruc As String, tipoEnt As String) As String
    ' *** Si existe te devuelve el codigo, sino te devuelve vacio
    Dim sqlSp As String
    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As New clsMantoTablas
    Dim arrDatos() As Variant
    
    CodigoEntidadRuc = ""
    sqlSp = "spCn_GrabaEntidad 'SEL_RUC', '" & gsEmpresa & "', '', '" & tipoEnt & "', '', '', '" & Ruc & "', '', '', '', '', '', '' "
    arrDatos = Array(sqlSp)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo.State = 0 Then Exit Function
    If rsArreglo(0) > 0 Then CodigoEntidadRuc = rsArreglo(0).Value
    Call CerrarRecordSet(rsArreglo)
    Set clDatos = Nothing
End Function

