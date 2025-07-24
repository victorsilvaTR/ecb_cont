Attribute VB_Name = "ModReprocesoSaldos"
Public Sub ProcesarSaldos(cMes As String)
    Dim i As Integer
    Dim Mes As String
    Dim Ultimo As Integer
    Dim inicio As Integer
    
    Ultimo = Val(cMes)
    inicio = 0
    
    For i = inicio To Ultimo
        Mes = Right("00" & i, 2)
        If Proceso(Mes) = False Then Exit For
    Next i
    
    For i = inicio To Ultimo
        Mes = Right("00" & i, 2)
        Call ActualizaSaldosSp(Mes)
    Next i
End Sub

Public Function Proceso(Mes As String) As Boolean
    Dim clsMante As clsMantoTablas
    Set clsMante = New clsMantoTablas

    Dim lArrMnt() As Variant
    ReDim lArrMnt(4) As Variant
    On Local Error GoTo ErrorEjecucion
    lArrMnt(0) = gsEmpresa          ' Empresa
    lArrMnt(1) = gsAnio             ' Codigo
    lArrMnt(2) = Mes                ' Nombre
    lArrMnt(3) = "A"                ' Nombre Plantilla
    lArrMnt(4) = gsUsuario          ' Usuario
    If CierreMes(Mes) = False Then
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_ReprocesoSaldos", lArrMnt(), True) = False Then
            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
            Screen.MousePointer = vbDefault
            Proceso = False
            Set clsMante = Nothing
            Exit Function
        End If
        
    End If
    Proceso = True
    Set clsMante = Nothing
    Exit Function
ErrorEjecucion:
    Proceso = False
    Mensajes Str(Err.Number) & " " & Err.Description, vbInformation
    Set clsMante = Nothing
End Function

Public Sub ActualizaSaldosSp(Mes As String)
    Dim clsMante2 As clsMantoTablas
    Dim lArrMnt(3) As Variant
    Set clsMante2 = New clsMantoTablas
    
    On Local Error GoTo ErrorEjecucion
    
    lArrMnt(0) = gsEmpresa
    lArrMnt(1) = gsAnio
    lArrMnt(2) = Mes
    lArrMnt(3) = Null
    'cuenta contable, cuando es reproceso total debe ser null,
    'solo se asigna una cuenta cuando se graba el voucher,en spCn_ActualizaSaldos
    
    If clsMante2.MantenimientoDeTablas(gsCadenaConexion, "spCn_AcumulaSaldosTit", lArrMnt(), False) = False Then
        Mensajes "El proceso no ha concluido. Verificar...", vbInformation
        Set clsMante2 = Nothing
        Exit Sub
    End If
    DoEvents
    If clsMante2.MantenimientoDeTablas(gsCadenaConexion, "spCn_AcumulaSaldosAnt", lArrMnt(), True) = False Then
        Mensajes "El proceso no ha concluido. Verificar...", vbInformation
        Set clsMante2 = Nothing
        Exit Sub
    End If
    
    Set clsMante2 = Nothing
    
    Exit Sub
ErrorEjecucion:
    Mensajes Str(Err.Number) & " " & Err.Description, vbInformation
    Set clsMante2 = Nothing
End Sub

Public Function ProcesarSaldosProv(Mes As String) As Boolean
    Dim clsMantenim As clsMantoTablas
    Set clsMantenim = New clsMantoTablas

    Dim lArrMnt() As Variant
    ReDim lArrMnt(2) As Variant
    
    On Local Error GoTo ErrorEjecucion
    lArrMnt(0) = gsEmpresa          ' Empresa
    lArrMnt(1) = gsAnio             ' Codigo
    lArrMnt(2) = Trim(Mes)          ' Periodo
    
        If clsMantenim.MantenimientoDeTablas(gsCadenaConexion, "spCn_ReprocesoSaldosProv", lArrMnt(), True) = False Then
            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
            Screen.MousePointer = vbDefault
            ProcesarSaldosProv = False
            
            Set clsMantenim = Nothing
            Exit Function
        End If
    
    Set clsMantenim = Nothing
    ProcesarSaldosProv = True
    Exit Function
ErrorEjecucion:
    Mensajes Str(Err.Number) & " " & Err.Description, vbInformation
End Function

Public Function ProcesarSaldosProvVoucher(Voucher As String) As Boolean
    Dim clsMantenim As clsMantoTablas
    Set clsMantenim = New clsMantoTablas

    Dim lArrMnt() As Variant
    ReDim lArrMnt(2) As Variant
    On Local Error GoTo ErrorEjecucion
    lArrMnt(0) = gsEmpresa          ' Empresa
    lArrMnt(1) = gsAnio             ' Codigo
    lArrMnt(2) = Voucher
        
        If clsMantenim.MantenimientoDeTablas(gsCadenaConexion, "spCn_RSaldosProvVoucher", lArrMnt(), True) = False Then
            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
            Screen.MousePointer = vbDefault
            ProcesarSaldosProvVoucher = False
            
            Set clsMantenim = Nothing
            Exit Function
        End If
    
    Set clsMantenim = Nothing
    ProcesarSaldosProvVoucher = True
    Exit Function
ErrorEjecucion:
    Mensajes Str(Err.Number) & " " & Err.Description, vbInformation
End Function

Public Function CierreMes(Mes As String, Optional TipoLibro As String) As Boolean
    Dim rsDatos As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Dim sqlDatos As String
    Dim arrDatos() As Variant
    
    Set clDatos = New clsMantoTablas
    sqlDatos = "spCn_GrabaCierre 'EXISTEREG', '" & gsEmpresa & "', '" & gsAnio & "', '" & Mes & "', 'I', '', '" & TipoLibro & "'"
    arrDatos = Array(sqlDatos)
    Set rsDatos = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If Not rsDatos Is Nothing Then
        If Not (rsDatos.EOF And rsDatos.BOF) Then
            If rsDatos(0).Value = 0 Then
                CierreMes = False
            Else
                CierreMes = True
            End If
        End If
    Else
        CierreMes = False
    End If
    
    Call CerrarRecordSet(rsDatos)
    Set clDatos = Nothing
    
End Function
