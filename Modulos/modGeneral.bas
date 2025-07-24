Attribute VB_Name = "modGeneral"
Option Explicit
Global oRegistroLock As Object
Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Const HKEY_LOCAL_MACHINE = &H80000002

Public Function MensajesRet(ByVal sMsg As String, Optional ByVal Tipo As VbMsgBoxStyle = vbOKOnly + vbInformation) As VbMsgBoxResult
    MensajesRet = MsgBox(sMsg, Tipo)
End Function

Public Sub pSendKeys(cCadena As String)
    On Error Resume Next
    SendKeys cCadena
End Sub

Public Function SeteaFondoForm(ByRef oForm As Form)
    If oForm.WindowState <> vbMaximized Then
        oForm.Picture = Nothing
    Else
        oForm.Picture = frmMDIConta.Picture
    End If
End Function
Public Function ValidaFormula(Formula As String, LongVar As Double) As Boolean
    ValidaFormula = False
    
    Dim Pos As Integer
    
    If CE(Formula) = "" Then
        Mensajes "Ingrese una formula para la cuenta"
        Exit Function
    End If
    '-------------------------------------------------------------------------------------------
    Pos = 1
    Do While Pos >= 1
        Pos = InStr(1, Formula, "BGE")
        If Pos > 0 Then
           Formula = Mid(Formula, 1, Pos - 1) & "1" & Mid(Formula, Pos + LongVar, Len(Formula))
        End If
    Loop
    '-------------------------------------------------------------------------------------------
    Pos = 1
    Do While Pos >= 1
        Pos = InStr(1, Formula, "FUN")
        If Pos > 0 Then
           Formula = Mid(Formula, 1, Pos - 1) & "1" & Mid(Formula, Pos + LongVar, Len(Formula))
        End If
    Loop
    '-------------------------------------------------------------------------------------------
    Pos = 1
    Do While Pos >= 1
        Pos = InStr(1, Formula, "NAT")
        If Pos > 0 Then
           Formula = Mid(Formula, 1, Pos - 1) & "1" & Mid(Formula, Pos + LongVar, Len(Formula))
        End If
    Loop
    '-------------------------------------------------------------------------------------------
    Pos = 1
    Do While Pos >= 1
        Pos = InStr(1, Formula, "CTA")
        Formula = Replace(Formula, "CTA", "")
        'If pos > 0 Then
        '   Formula = Mid(Formula, 1, pos - 1) & "1" & Mid(Formula, pos + LongVar, Len(Formula))
        'End If
    Loop
    '-------------------------------------------------------------------------------------------
    Pos = 1
    Do While Pos >= 1
        Pos = InStr(1, Formula, "ANA")
        If Pos > 0 Then
           Formula = Mid(Formula, 1, Pos - 1) & "1" & Mid(Formula, Pos + LongVar, Len(Formula))
        End If
    Loop
    '-------------------------------------------------------------------------------------------
    Dim rs As ADODB.Recordset
    LlenarRecordSet "select " & Formula, rs, False
    
    If rs Is Nothing Then
       Mensajes "Verifique la formula esta mal diseñada"
       ValidaFormula = False
    Else
       ValidaFormula = True
    End If
    
    CerrarRecordSet rs
    
End Function

Public Function GrabaPeriodoActivo() As Boolean
    
    On Error GoTo serror
    Dim cn As ADODB.Connection
    Set cn = New ADODB.Connection
    cn.ConnectionString = gsCadenaConexion
    cn.Open
    cn.Execute "UPDATE SGM_USUARIOS SET LOG_EMPCOD='" & gsEmpresa & "', LOG_PANANIO='" & gsAnio & "' , LOG_PERIODO= '" & gsPeriodo & "' " & _
               "WHERE usu_cCodUsuario='" & gsUsuario & "' "

    If cn.State = adStateOpen Then cn.Close
    Set cn = Nothing
    
    GrabaPeriodoActivo = True
    Exit Function
    
serror:
    GrabaPeriodoActivo = False
End Function
Public Function ValidarCuentaCostoVenta()

    Dim rs As New ADODB.Recordset
    Dim ObjFuncion As New ClsFuncionesExecute
    Dim sql As String

    sql = "SPCN_VALIDARCUENTACOSTOVENTA '" & gsEmpresa & "', '" & gsAnio & "'"
    Set rs = ObjFuncion.fRetornaRS(sql)

    If Not rs.EOF Then
        gstrCuentaCostoVenta = rs!Pla_cCuentaContable
    End If

    Exit Function

End Function

Public Function Titulo(sCaption As String, STitulo As String)
    If STitulo = "" Then
       Titulo = sCaption
    Else
       Titulo = STitulo
    End If
End Function

Public Function ExisteFile(sDirPath As String)
Dim VLman_arch As New filesystemobject
'Set VLman_arch = New FileSystemObject


If Right(sDirPath, 1) <> "\" Then sDirPath = sDirPath & "\"
       
    ' -- Linea de código Opcional para  Comprobar previamente si el path ya existe
    If Len(Dir(sDirPath, vbDirectory)) = 0 Then
        Call MakeSureDirectoryPathExists(sDirPath)
        'ExisteFile = CBool(Len(Dir(sDirPath, vbDirectory)))
    End If
'End If

' lsLibShow = True
'If Not VLman_arch.FolderExists(ruta) Then
' MsgBox "La carpeta " & Chr(34) & ruta & Chr(34) & "no existe", vbInformation, "Contabilidad"
'lsLibShow = False
'ExisteFile = False
'Else
'ExisteFile = True
'End If
End Function

 

Public Function GenMenuUserProfile(ByVal sPrivilegio As String) As String
Dim sProfile As String
    sPrivilegio = Left(Trim(sPrivilegio) & "0000", 4)
    sProfile = Mid(sPrivilegio, 2, 1) _
            & "1" _
            & "1" _
            & Mid(sPrivilegio, 4, 1) _
            & Mid(sPrivilegio, 3, 1) _
            & Mid(sPrivilegio, 1, 1) _
            & "1" _
            & "1"
    ' PROFILE = NUEVO + CONSULTA + GRABAR + ELIMINAR + MODIFICAR + ELIMINAR + IMPRIMIR + CANCELAR + SALIR
    GenMenuUserProfile = sProfile
End Function

Public Function ExistenDatosOp(CodigoBuscaOP As String) As Boolean
    ExistenDatosOp = False
    Dim rsDatos As New ADODB.Recordset
    Set rsDatos = BuscaCodigosOp(CodigoBuscaOP)
    If Not rsDatos Is Nothing Then
        If rsDatos.State = adStateOpen Then
            If Not (rsDatos.EOF And rsDatos.BOF) Then
                ExistenDatosOp = True
            End If
        End If
    End If
End Function

Public Function fgDocCfgOpera(Valor As String) As String
    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant

    Dim sqlSp As String
    
    fgDocCfgOpera = ""
    Set clDatos = New clsMantoTablas
    sqlSp = "SELECT Cod_cValorParam FROM CND_CONFIG_OPERA WHERE Emp_cCodigo = '" & gsEmpresa & "' And Pan_cAnio = '" & gsAnio & "' And Cop_cCodigo = '" & Valor & "'"
    arrDatos = Array(sqlSp)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo.State <> 0 Then
        fgDocCfgOpera = rsArreglo("Cod_cValorParam").Value
    End If
    Call CerrarRecordSet(rsArreglo)
    Set clDatos = Nothing

End Function

Public Function BuscaCodigosOp(CodigoBuscarOP As String) As ADODB.Recordset
    Dim clDatos As New clsMantoTablas
    Dim arrDatos() As Variant
    Dim sqlver As String
    
    sqlver = "SELECT DISTINCT Cod_cValorParam as valor FROM CND_CONFIG_OPERA " & _
             "WHERE Emp_cCodigo='" & gsEmpresa & "' and pan_cAnio='" & gsAnio & "' "
    
    arrDatos = Array(sqlver)
    Set BuscaCodigosOp = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    
    Set clDatos = Nothing
End Function

Public Function BuscaValorEnOp(CodigoBuscarOP As String) As String
    Dim sqlver As String
    sqlver = "SELECT DISTINCT Cod_cValorParam as valor FROM CND_CONFIG_OPERA " & _
             "WHERE Emp_cCodigo='" & gsEmpresa & "' and pan_cAnio='" & gsAnio & "' and " & _
             "cop_ccodigo='" & CE(CodigoBuscarOP) & "'"
    
    BuscaValorEnOp = fRetornaValor(sqlver)
    
End Function


Public Function BuscaCodigoConfOpRango(CodigoBuscar As String, CodigoBuscarOPIni As String, CodigoBuscarOPFin As String) As Boolean
    Dim clDatos As New clsMantoTablas
    Dim rsArreglo As ADODB.Recordset
    Dim arrDatos() As Variant
    Dim sqlver As String
    
    BuscaCodigoConfOpRango = False
    
    sqlver = "SELECT DISTINCT Cod_cValorParam as valor FROM CND_CONFIG_OPERA " & _
             "WHERE Emp_cCodigo='" & gsEmpresa & "' and pan_cAnio='" & gsAnio & "' and  " & _
             "cop_ccodigo>='" & CE(CodigoBuscarOPIni) & "' and cop_ccodigo<='" & CE(CodigoBuscarOPFin) & "'"
    
    arrDatos = Array(sqlver)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If Not rsArreglo Is Nothing Then
        If Not (rsArreglo.EOF And rsArreglo.BOF) Then
            Do While Not rsArreglo.EOF
                If CE(rsArreglo!Valor) = CE(CodigoBuscar) Then
                    BuscaCodigoConfOpRango = True
                    Exit Do
                End If
                rsArreglo.MoveNext
            Loop
        End If
    End If
    
    Call CerrarRecordSet(rsArreglo)
    Set clDatos = Nothing
End Function

Public Function BuscaCodigoConfOp(CodigoBuscar As String, CodigoBuscarOP As String, Optional xCod_cValorParam As String) As Boolean
    Dim clDatos As New clsMantoTablas
    Dim rsArreglo As ADODB.Recordset
    Dim arrDatos() As Variant
    Dim sqlver As String
    
    BuscaCodigoConfOp = False
    
    sqlver = "SELECT DISTINCT Cod_cValorParam as valor FROM CND_CONFIG_OPERA " & _
             "WHERE Emp_cCodigo='" & gsEmpresa & "' and pan_cAnio='" & gsAnio & "' and  " & _
             "cop_ccodigo='" & CE(CodigoBuscarOP) & "' AND Cod_cValorParam = '" & Trim(xCod_cValorParam) & "'"
    
    arrDatos = Array(sqlver)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If Not rsArreglo Is Nothing Then
        Do While Not rsArreglo.EOF
            If CE(rsArreglo!Valor) = CE(CodigoBuscar) Then
                BuscaCodigoConfOp = True
                Exit Do
            End If
            rsArreglo.MoveNext
        Loop
    End If
    
    Call CerrarRecordSet(rsArreglo)
    Set clDatos = Nothing
End Function

Public Function pCargaCfgLibro() As Boolean
    Dim clDatos As New clsMantoTablas
    Dim rsArreglo As ADODB.Recordset
    Dim arrDatos() As Variant, sqlver As String

    sqlver = "SELECT * From CNT_CONFIG_LIBROS WHERE Emp_cCodigo = '" & gsEmpresa & "' And Pan_cAnio = '" & gsAnio & "'"
    arrDatos = Array(sqlver)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If Not rsArreglo Is Nothing Then
       lsLibroCom = CE(rsArreglo("Cfl_cCompras"))
       lsLibroVen = CE(rsArreglo("Cfl_cVentas"))
       lsLibroHon = CE(rsArreglo("Cfl_cHonorarios"))
       lsLibroDiario = CE(rsArreglo("Cfl_cDiario"))
       lsLibroDif = CE(rsArreglo("Cfl_cDifCam"))
       lsLibroCierre = CE(rsArreglo("Cfl_cCierre"))
       lsLibroRet = CE(rsArreglo("Cfl_cRetencion"))
       On Error Resume Next
       lsLibroApe = CE(rsArreglo("Cfl_cApertura"))
       If Len(Trim(CE(rsArreglo("Cfl_cCaja")))) > 0 Then
          lsLibroCajIng = CE(rsArreglo("Cfl_cCaja"))
          lsLibroCajEgr = CE(rsArreglo("Cfl_cCaja"))
       Else
          lsLibroCajIng = CE(rsArreglo("Cfl_cCajaIngresos"))
          lsLibroCajEgr = CE(rsArreglo("Cfl_cCajaEgresos"))
       End If
              
       gsTipoPlan = NE(rsArreglo("Cfl_cTipoPlan"))
       gsNumDigDiarioSimpRep = NE(rsArreglo("Cfl_nDigDiarioRep"))
       gsBaseImpDefCom = CE(rsArreglo("Cfl_cBaseDefCom"))
       gsBaseImpDefVtas = CE(rsArreglo("Cfl_cBaseDefVtas"))
       gsDiarioSimplificado = NE(rsArreglo("Cfl_cDiarioSimplificado"))
       gsPLE = NE(rsArreglo("Cfl_cple"))
       gintLEVentaSimplificado = IIf(NE(rsArreglo("Cfl_cLEVenta")) = Null, 0, NE(rsArreglo("Cfl_cLEVenta")))
       gintLECompraSimplificado = IIf(NE(rsArreglo("Cfl_cLECompra")) = Null, 0, NE(rsArreglo("Cfl_cLECompra")))
       lsLibroTransferenciaCancelacion = rsArreglo("Cfl_cTransferencia")
       lsLibroTransCancAutomatico = rsArreglo("Cfl_cTransAutomatico")
       lsLibroAjusteNIF = IIf(CE(rsArreglo("Cfl_cAjusteNIF")) = vbNullString, "", rsArreglo("Cfl_cAjusteNIF"))
       gstrVersionLE = CE(rsArreglo("Cfl_cVersionLE"))
       gsRVIE = CE(rsArreglo("Cfl_cRVIE")) 'frt_rvie
    End If
    
    Call CerrarRecordSet(rsArreglo)
    
    gsTDNC = fgDocCfgOpera("012")
    gsLetCobrar = fgDocCfgOpera("023")
    gsLetPagar = fgDocCfgOpera("022")
    gsCheque = fgDocCfgOpera("019")
    
    pCargaCfgLibro = True
    
    If lsLibroCom = "" Or lsLibroVen = "" Or lsLibroDiario = "" Or _
    lsLibroDif = "" Or lsLibroCierre = "" Or lsLibroApe = "" Or _
    (lsLibroCajIng = "" Or lsLibroCajEgr = "") Then
       Mensajes "Falta definir los códigos de los libros, revise parametros iniciales"
       pCargaCfgLibro = False
    End If
    
'    If lsLibroCom = "" Then
'       Mensajes "Falta definir el código del libro de Compras, revise parametros iniciales"
'       pCargaCfgLibro = False
'    End If
'
'    If lsLibroVen = "" Then
'       Mensajes "Falta definir el código del libro de Ventas, revise parametros iniciales"
'       pCargaCfgLibro = False
'    End If
'
'    If lsLibroDiario = "" Then
'       Mensajes "Falta definir el código del libro de Diario, revise parametros iniciales"
'       pCargaCfgLibro = False
'    End If
'
'    If lsLibroDif = "" Then
'       Mensajes "Falta definir el código del libro de Diferencia de Cambio, revise parametros iniciales"
'       pCargaCfgLibro = False
'    End If
'
'    If lsLibroCierre = "" Then
'       Mensajes "Falta definir el código del libro de Cierre, revise parametros iniciales"
'       pCargaCfgLibro = False
'    End If
'
'    If lsLibroApe = "" Then
'       Mensajes "Falta definir el código del libro de Apertura, revise parametros iniciales"
'       pCargaCfgLibro = False
'    End If
'
'    If lsLibroCajIng = "" Or lsLibroCajEgr = "" Then
'       Mensajes "Falta definir el código del libro de Caja, revise parametros iniciales"
'       pCargaCfgLibro = False
'    End If
    
    
End Function

Function Cifrar(ByVal s As String, ByVal P As String) As String
    Dim i As Integer, R As String
    Dim c1 As Integer, c2 As Integer
    R = ""
    If Len(P) > 0 Then
    For i = 1 To Len(s)
    c1 = Asc(Mid(s, i, 1))
    If i > Len(P) Then
    c2 = Asc(Mid(P, i Mod Len(P) + 1, 1))
    Else
    c2 = Asc(Mid(P, i, 1))
    End If
    c1 = c1 + c2 + 64
    If c1 > 255 Then c1 = c1 - 256
    R = R + Chr(c1)
    Next i
    Else
    R = s
    End If
    Cifrar = R
End Function

Function DesCifrar(ByVal s As String, ByVal P As String) As String
    Dim i As Integer, R As String
    Dim c1 As Integer, c2 As Integer
    R = ""
    If Len(P) > 0 Then
    For i = 1 To Len(s)
    c1 = Asc(Mid(s, i, 1))
    If i > Len(P) Then
    c2 = Asc(Mid(P, i Mod Len(P) + 1, 1))
    Else
    c2 = Asc(Mid(P, i, 1))
    End If
    c1 = c1 - c2 - 64
    If Sgn(c1) = -1 Then c1 = 256 + c1
    R = R + Chr(c1)
    Next i
    Else
    R = s
    End If
    DesCifrar = R
End Function

Public Sub ControlAbs(ByRef Control As Object)
    On Error GoTo serror
    If IsNumeric(Control) Then
        Control = Abs(NE(Control))
        Control = Format(NE(Control), "###,###,##0.00")
        
        'Control.Refresh
        'DoEvents
    End If
    Exit Sub
serror:
    Control = Abs(NE(Control))
End Sub

Public Function Sincomas(xx As String)
    Dim ss As Variant
    Dim x As Integer
    Dim yy As String
    If Len(xx) >= 9 Then
    ss = Left(xx, Len(xx) - 3)
        If Len(ss) >= 9 Then
            For x = 1 To Len(ss)
                If Mid(ss, x, 1) <> "," Then
                    yy = yy & Mid(ss, x, 1)
                End If
            Next
            Sincomas = yy & Right(xx, 3)
        Else
            Sincomas = xx
        End If
    Else
        If xx = "" Then
            Sincomas = 0
        Else
            Sincomas = xx
        End If
    End If
End Function

Public Sub ActivarControl(ByRef Control As Object, Valor As Boolean, Optional nColor As OLE_COLOR)
    Dim gsColor  As OLE_COLOR
    
    Control.Enabled = Valor
    If nColor <> 0 Then
        gsColor = nColor
    Else
        If Valor = True Then
           gsColor = gsColorActivado
           nColor = gsColor
        Else
           gsColor = gsColorDesactivado
        End If
    End If
    
    If Valor = True Then
        If gsColor <> 0 Then
            gsColor = nColor
        Else
            gsColor = gsColorActivado
        End If
    
    End If
    
    Control.BackColor = gsColor
End Sub

Public Function ValidaDigitos(nDigitos As String) As Boolean

    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As New clsMantoTablas
    Dim cadena As String
    Dim entro As Boolean
    Dim arrDatos() As Variant
    Dim sqlSp As String
    
    ValidaDigitos = False
    entro = False
    
    sqlSp = "spCn_ConsultaCuentas 'SEL_DIG', '" & gsEmpresa & "', '" & gsAnio & "', ''"
    arrDatos = Array(sqlSp)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo.State = 0 Then
        Mensajes "No se encontraron registros.", vbInformation
    Else
        cadena = "Los digitos permitidos son "
        Do While Not rsArreglo.EOF
            cadena = cadena & CE(rsArreglo!Digitos) & " , "
            If CE(rsArreglo!Digitos) = CE(nDigitos) Then
                entro = True
            End If
            rsArreglo.MoveNext
        Loop
        cadena = Mid(cadena, 1, Len(cadena) - 2)
        
        If entro = False Then
            Mensajes cadena, vbOKOnly + vbInformation
            ValidaDigitos = False
        Else
            ValidaDigitos = True
        End If
    End If
    
    Set rsArreglo = Nothing
    Set clDatos = Nothing
    
End Function

Public Sub AlmacenaArray(registros As ADODB.Recordset)
    On Error Resume Next
    Set gsArray = New XArrayDB
    gsArray.ReDim 0, 150, 0, 1
    gsArray.Clear
    
    If Not registros Is Nothing Then
        If registros.State = adStateOpen Then
            gsArray.ReDim 0, registros.RecordCount, 0, 1
            gsArray.Clear
            
            registros.MoveFirst
            
            Do While Not registros.EOF
                gsArray(registros.AbsolutePosition, 0) = CE(registros!opm_cnomobj)
                gsArray(registros.AbsolutePosition, 1) = CE(registros!PFL_CCODPERFIL)
                registros.MoveNext
            Loop
        End If
    End If
End Sub

Public Function BuscaArray(NombreObjeto As String) As String
    On Error Resume Next
    Dim Indice As Long
    Indice = gsArray.Find(0, 0, NombreObjeto)
    BuscaArray = gsArray(Indice, 1)
End Function

Public Sub SeteaBarraHerramientas(BarraHerram As Object, gsGrupo As String)
    With BarraHerram
        If .Buttons(1).Enabled Then .Buttons(1).Enabled = ExtraerCodigo(gsGrupo, G_INGRESAR) 'NUEVO
        If .Buttons(2).Enabled Then .Buttons(2).Enabled = ExtraerCodigo(gsGrupo, G_CONSULTAR)  'CONSULTA
        If .Buttons(3).Enabled Then .Buttons(3).Enabled = ExtraerCodigo(gsGrupo, G_MODIFICAR)  'GRABAR
        If .Buttons(4).Enabled Then .Buttons(4).Enabled = ExtraerCodigo(gsGrupo, G_ELIMINAR)  'ELIMINAR
        If .Buttons(5).Enabled Then .Buttons(5).Enabled = ExtraerCodigo(gsGrupo, G_MODIFICAR)  'MODIFICAR
    End With
End Sub

Public Function ExtraerCodigo(Grupo As String, TipoG As TipoGrupo) As Boolean

    Dim Codigo As String
    
    If Grupo = gsPrivilegioAdmin Then
        Codigo = "1"
    Else
        Select Case TipoG
            Case TipoGrupo.G_CONSULTAR
                 Codigo = Mid(Grupo, 1, 1)
            Case TipoGrupo.G_INGRESAR
                 Codigo = Mid(Grupo, 2, 1)
            Case TipoGrupo.G_MODIFICAR
                 Codigo = Mid(Grupo, 3, 1)
            Case TipoGrupo.G_ELIMINAR
                 Codigo = Mid(Grupo, 4, 1)
        End Select
    End If
    If Codigo = "0" Then
        ExtraerCodigo = False
    Else
        ExtraerCodigo = True
    End If
    
End Function
Public Sub pSetFocus(Control As Object)
On Error Resume Next
    If Control.Enabled = True Then
        Control.SetFocus
    End If
End Sub
Public Function Salto(Lineas As Integer) As String
    Dim i As Integer
    For i = 1 To Lineas
        Salto = Salto & Chr(10) + Chr(13)
    Next i
End Function

Public Function Mensajes(ByVal sMsg As String, Optional ByVal Tipo As VbMsgBoxStyle = vbOKOnly + vbInformation) As VbMsgBoxResult
    Mensajes = MsgBox(sMsg, Tipo, gsNombreModulo)
End Function

Public Function BuscaDescripcionCuenta(empresa As String, Anio As String, Cuenta As String) As String
    
    Dim sqlver As String
    Dim valorDato As String
    
    sqlver = "SELECT Pla_cNombreCuenta FROM CNM_PLAN_CTA " & _
             "WHERE Emp_cCodigo='" & empresa & "' AND Pan_cAnio='" & Anio & "' AND " & _
             "Pla_cCuentaContable='" & Cuenta & "' AND Pla_cTitulo='N'"
    valorDato = ExtraeDescripcion(sqlver)
    
    BuscaDescripcionCuenta = valorDato
    
End Function

Function AbreReporteParam(File_DNS As String, ByRef reporte As Object, ByVal RutaYNombreReporte As String, _
                          Destino As DestinationConstants, _
                          ByVal TituloVentana As String, ByVal FormulaSeleccion As String, _
                          ByRef Parametros(), ByRef formulas(), _
                          Optional ByVal pOrientacion As Orientacion_Pagina = Orientacion_Pagina.defecto, _
                          Optional ByVal pTipoPagina As Tipo_Pagina = Tipo_Pagina.defecto) As Boolean
   On Error GoTo ErrReporte
   
   Call ImprimirReporte(RutaYNombreReporte, Parametros(), formulas(), TituloVentana, False, pOrientacion, pTipoPagina)
   
   Exit Function

ErrReporte:
    MsgBox "Error # " & Err.Number & vbCrLf & "Descripcion  :" & Err.Description, vbCritical, "Error al Cargar el Reporte"
    AbreReporteParam = False
End Function

Public Function fValidarNroRuc(cRuc As String) As Boolean
    Dim objPDT As Object
    
    fValidarNroRuc = False
    If Len(Trim(cRuc)) = 0 Then
       Exit Function
    End If
    
    fValidarNroRuc = True
    
    On Error GoTo eError
    Set objPDT = CreateObject("PDTCommonFuncs.clsCommonFuncs")
    
    If Not objPDT.ValidaRuc(cRuc) Then
       fValidarNroRuc = False
    End If
    Set objPDT = Nothing
    
    Exit Function
eError:
Set objPDT = Nothing
End Function

Public Sub HabilitaControl(ByRef Formulario As Form)
   Dim ctrl As Control
   For Each ctrl In Formulario.Controls
        'ctrl.Enabled = True
       If ctrl.Tag <> "" Then
        ctrl.Enabled = Not ctrl.Enabled
'            MsgBox ctrl.Name
       End If
       'MsgBox ctrl.Name
   Next ctrl
End Sub

Public Sub HabilitaControlBool(ByRef Formulario As Form, Valor As Boolean)
   Dim ctrl As Control
   For Each ctrl In Formulario.Controls
        'ctrl.Enabled = True
       If ctrl.Tag <> "" Then ctrl.Enabled = Valor
       'MsgBox ctrl.Name
   Next ctrl
End Sub

Public Sub AseguraControl(ByRef Formulario As Form, Valor As Boolean)
    Dim ctrl As Control
    For Each ctrl In Formulario.Controls
        If ctrl.Tag <> "" Then
            If TypeOf ctrl Is TDBCombo Then ctrl.Locked = Valor
            If TypeOf ctrl Is TDBText Then ctrl.ReadOnly = Valor
            If TypeOf ctrl Is TDBNumber Then ctrl.ReadOnly = Valor
            If TypeOf ctrl Is TDBDate Then ctrl.ReadOnly = Valor
            If TypeOf ctrl Is CheckBox Then
                ctrl.Enabled = ctrl.Enabled
            Else
                ctrl.Enabled = Valor
            End If
        End If
    Next ctrl
End Sub

Public Function NumeroLleno(ByRef Texto As Control, cadena As String) As Boolean
    ' *** Para controles q tienen la propiedad psetfocus
    NumeroLleno = True
    If Texto = 0 Then
       Mensajes "Ingrese datos en : " & cadena, vbInformation
       pSetFocus Texto
       NumeroLleno = False
    End If
End Function

Public Function TextoLleno(ByRef Texto As Control, cadena As String) As Boolean
    ' *** Para controles q tienen la propiedad psetfocus
    TextoLleno = True
'    If texto = "" Then
'       Mensajes "Insertar datos en : " & cadena, vbInformation
'       pSetFocus texto
'       TextoLleno = False
'    End If
End Function

Public Function TextoSeleccionado(ByRef Texto As Control, cadena As String) As Boolean
    TextoSeleccionado = True
    If Texto = "" Then
       Mensajes "Seleccione un registro de la lista de  " & cadena, vbInformation
       pSetFocus Texto
       TextoSeleccionado = False
    End If

End Function

Public Function TextoLleno2(ByRef Texto As Control, cadena As String) As Boolean
    TextoLleno2 = True
    If Texto = "" Then
       Mensajes "Ingrese los datos de " & cadena, vbInformation
       pSetFocus Texto
       TextoLleno2 = False
    End If
End Function

Public Function datoLleno(Dato As Variant, Nombre As String) As Boolean
    ' *** Para controles q no tienen la propiedad psetfocus
    datoLleno = True
    If Dato = "" Then
        datoLleno = False
        Mensajes "Ingrese datos al campo: " & Nombre, vbInformation
    End If
End Function

Public Sub LimpiaTexto(ByRef Formulario As Form)
   Dim ctrl As Control
   For Each ctrl In Formulario.Controls
      If TypeOf ctrl Is TDBText Then
         If ctrl.Tag <> "" Then ctrl.Text = ""
      End If
      If TypeOf ctrl Is TDBNumber Then
         If ctrl.Tag <> "" Then ctrl.Value = 0
      End If
      If TypeOf ctrl Is CheckBox Then
         If ctrl.Tag <> "" Then ctrl.Value = "0"
      End If
   Next ctrl
End Sub

Public Sub SelTexto(txt As Variant)
    txt.SelStart = 0
    txt.SelLength = Len(txt.Text)
End Sub

Public Function BE(Valor As Variant) As Boolean
    On Error GoTo serror
    If IsNull(Valor) Then
        BE = False
    Else
        If NE(Valor) = 1 Then
            BE = True
        Else
            BE = False
        End If
    End If
    
    Exit Function
    
serror:
    
    BE = False
End Function

Public Function CE(Valor As Variant) As String
    On Error GoTo serror
    If IsNull(Valor) Then
        CE = ""
    Else
        CE = Trim(Valor)
    End If
    
    Exit Function
    
serror:
    
    CE = Valor
End Function

Public Function NE(Valor As Variant) As Double
    On Error GoTo serror
    If Not IsNumeric(Valor) Then
        NE = 0
    Else
        If IsNull(Valor) Or Valor = "" Or Valor = " " Then
            NE = 0
        Else
            NE = Valor
        End If
    End If
    
    Exit Function
serror:
    
    NE = Valor
    
End Function


Public Function NuloText(Valor As Field)
    NuloText = IIf(IsNull(Valor.Value), "", Valor.Value)
End Function

Public Function NuloNum(Valor As Field)
    NuloNum = IIf(IsNull(Valor.Value), 0, Valor.Value)
End Function

Public Sub LlamaBuscar(ByRef formOrigen As Form, nomControl As String, _
                       Control As String, tabla As String, ByRef formul As Form, Optional pPeriodo As String, _
                       Optional auxiliar As String, Optional pCuenta As String = "", _
                       Optional pEntidad As String = "", _
                       Optional pSerie As String = "", Optional pNumero As String = "", _
                       Optional pTipoCuenta As String = "", Optional pAux As String = "", _
                       Optional TipoEntidad As String, Optional pVTipLib As Boolean = False)

On Error GoTo serror

    If formOrigen.enUso = True Then
        If formOrigen.Name = "frmBuscador" Then
            frmBuscador.Cerrar
            DoEvents
        Else
            Mensajes "Actualmente esta realizando una busqueda desde otro Formulario... Cancele la busqueda primero.", vbInformation
            Exit Sub
        End If
    End If
    
    Control = nomControl
    formOrigen.tabla = tabla
    formOrigen.auxiliar = auxiliar
        
    If nomControl = "Provisiones" Then
       formOrigen.Cuenta = pCuenta
       formOrigen.Entidad = pEntidad
       formOrigen.Serie = pSerie
       formOrigen.Numero = pNumero
'       formOrigen.TipoEntidad = TipoEntidad
'        formOrigen.Show vbModal
    End If
    
    If formul.Name = "frmManPDBVentas" Then
        formOrigen.Libro = auxiliar
        formOrigen.auxiliar = "V"
        formOrigen.pPeriodo = gsPeriodoCOA
    ElseIf formul.Name = "frmManPDBCompras" Then
        formOrigen.Libro = auxiliar
        formOrigen.auxiliar = "C"
        formOrigen.pPeriodo = gsPeriodoCOA
    ElseIf formul.Name = "frmManPDBPagos" Then
        formOrigen.Libro = auxiliar
        formOrigen.auxiliar = "C"
        formOrigen.pPeriodo = gsPeriodoCOA
    End If
    
    If formul.Name = "frmManAsientosContables" Then
        formul.Enabled = False
        Control = nomControl
        If formOrigen.Name = "frmBusTipoAsiento" Then
            frmBusTipoAsiento.tabla = tabla
            frmBusTipoAsiento.auxiliar = auxiliar
            frmBusTipoAsiento.NombreOrigen = formul.Name
            frmBusTipoAsiento.NombreBuscador = formOrigen.Name
            frmBusTipoAsiento.VarLibCom = pVTipLib
            frmBusTipoAsiento.FechaReg = pAux
            frmBusTipoAsiento.dtpFecha.Text = pAux
            frmBusTipoAsiento.Show
        ElseIf formOrigen.Name = "frmBusProvisiones" Then
            frmBusProvisiones.tabla = tabla
            frmBusProvisiones.auxiliar = auxiliar
            frmBusProvisiones.NombreOrigen = formul.Name
            frmBusProvisiones.NombreBuscador = formOrigen.Name
            frmBusProvisiones.Show
        Else
            frmBuscador.tabla = tabla
            frmBuscador.auxiliar = auxiliar
            frmBuscador.NombreOrigen = formul.Name
            frmBuscador.NombreBuscador = formOrigen.Name
            frmBuscador.Show vbModal
        End If
'    ElseIf formul.Name = "FrmManRegAcc" Then
'            frmBuscador.tabla = tabla
'            frmBuscador.auxiliar = auxiliar
'            frmBuscador.NombreOrigen = formul.Name
'            frmBuscador.NombreBuscador = formOrigen.Name
'            frmBuscador.Show vbModal
        Else
            If formOrigen.Name = "frmBuscador" Then
               If frmMDIConta.BuscaForm("frmBuscador") Then frmBuscador.Cerrar
            End If
            Set formOrigen.frmOrigen = formul
            formOrigen.nDigitos = NE(formul.Tag)
            formOrigen.NombreOrigen = formul.Name
            formOrigen.NombreBuscador = formOrigen.Name
            formul.Enabled = False
            If formOrigen.Name = "frmBuscador" Then
                formOrigen.Show vbModal
            Else
                formOrigen.Show
            End If
    
        
    End If
    
    Exit Sub
serror:
MsgBox Err.Description
End Sub

Public Function VerificaFecha(cadena As String) As Boolean
    VerificaFecha = True
    Dim Mes As Integer
    Dim dia As Integer
    Dim año As Integer
    Dim lenAño As Integer
    
    On Local Error GoTo ErrorFecha
    dia = Int(Mid(cadena, 1, 2))
    Mes = Int(Mid(cadena, 4, 2))
    año = Int(Mid(cadena, 7, 4)) Mod 4
    ' *** En caso de año a 2 digitos. Verificar
    lenAño = Len(Trim(Mid(cadena, 7, 4)))
    If lenAño < 2 Or lenAño = 3 Then GoTo ErrorFecha
    If dia = 0 Then GoTo CambiaFunction
    If Mes > 12 Then
        Mensajes "Número de mes no existe. Verifique...", vbInformation
        GoTo CambiaFunction
    End If
    Select Case Mes
        Case 1, 3, 5, 7, 8, 10, 12
            If dia > 31 Then
                Mensajes "Número de día no existe para el mes indicado. Verifique...", vbInformation
                GoTo CambiaFunction
            End If
        Case 4, 6, 9, 11
            If dia > 30 Then
                Mensajes "Número de día no existe para el mes indicado. Verifique...", vbInformation
                GoTo CambiaFunction
            End If
        Case 2
            If año <> 0 Then
                If dia > 28 Then
                    Mensajes "Número de día no existe para el mes indicado. Verifique...", vbInformation
                    GoTo CambiaFunction
                End If
            Else
                If dia > 29 Then
                    Mensajes "Número de día no existe para el mes indicado. Verifique...", vbInformation
                    GoTo CambiaFunction
                End If
            End If
    End Select
    Exit Function
CambiaFunction:
    VerificaFecha = False
    Exit Function
ErrorFecha:
    VerificaFecha = False
    Mensajes "Fecha mal ingresada. Verificar", vbInformation
End Function

Public Function VerificaFechaSM(cadena As String) As Boolean
    VerificaFechaSM = True
    Dim Mes As Integer
    Dim dia As Integer
    Dim año As Integer
    Dim lenAño As Integer
    On Local Error GoTo ErrorFecha
    dia = Int(Mid(cadena, 1, 2))
    Mes = Int(Mid(cadena, 4, 2))
    año = Int(Mid(cadena, 7, 4)) Mod 4
    ' *** En caso de año a 2 digitos. Verificar
    lenAño = Len(Trim(Mid(cadena, 7, 4)))
    If lenAño < 2 Or lenAño = 3 Then GoTo ErrorFecha
    If dia = 0 Then GoTo CambiaFunction
    If Mes > 12 Then GoTo CambiaFunction
    Select Case Mes
        Case 1, 3, 5, 7, 8, 10, 12
            If dia > 31 Then GoTo CambiaFunction
        Case 4, 6, 9, 11
            If dia > 30 Then GoTo CambiaFunction
        Case 2
            If año <> 0 Then
                If dia > 28 Then GoTo CambiaFunction
            Else
                If dia > 29 Then GoTo CambiaFunction
            End If
    End Select
    Exit Function
CambiaFunction:
    VerificaFechaSM = False
    Exit Function
ErrorFecha:
    VerificaFechaSM = False
End Function

Public Function FormatoFecha(cadena As String) As String
    Dim dia As String
    Dim Mes As String
    Dim Anio As String
    
    dia = Trim(Mid(cadena, 1, 2))
    Mes = Trim(Mid(cadena, 4, 2))
    Anio = Trim(Mid(cadena, 7, 4))
    cadena = Format(dia, "00") + Mid(cadena, 3)
    cadena = Mid(cadena, 1, 3) + Format(Mes, "00") + Mid(cadena, 6)
    If Len(Anio) = 2 Then
        If Right(Anio, 2) > 70 Then
            cadena = Mid(cadena, 1, 6) + Format(Anio, "1900")
        Else
            cadena = Mid(cadena, 1, 6) + Format(Anio, "2000")
        End If
    End If
    FormatoFecha = cadena
End Function

Public Function FechaRegistro(Valor As Variant) As String
    FechaRegistro = IIf(IsNull(Valor), "", Valor)
    If Valor = "01/01/1900" Then FechaRegistro = ""
End Function

Function RR(xx As Double)
Dim ss As Double

Dim Numero As String
Dim nuevo As String
Dim Pos As Integer

Numero = CStr(xx)
Pos = InStr(1, Numero, ".")

nuevo = "0.00"

If Pos > 0 Then
    nuevo = "0" & Mid(Numero, Pos, Len(Numero) - (Pos - 1))
End If

ss = CDbl(nuevo)

If Len(CStr(ss)) > 4 Then
    If Mid(CStr(ss), 5, 1) = 5 Then
        RR = Round(xx + 0.001, 2)
    Else
        RR = Round(xx, 2)
    End If
Else
    RR = Round(xx, 2)
End If


End Function

Public Function RedondearV2(Valor As Double) As Double
    Dim cadena As String
    Dim Salida As String
    Dim Pos As Integer
    Dim Retorno As Double
    Dim Digito As String
    cadena = CStr(Valor)
    Pos = InStr(1, cadena, ".")
    Digito = Mid(cadena, Pos + 3, 1)
    Salida = Left(cadena, Pos + 2)
    
    Retorno = CDbl(Salida)
    
    If Digito > "5" Then
        Retorno = Retorno + 0.01
    End If
    
    RedondearV2 = Retorno

End Function

Public Function Redondear(dblnToR As Double, Optional intCntDec As Double) As Double

        Redondear = RR(dblnToR)

End Function

Public Sub SoloLetras(KeyAscii As Integer, Optional Tabula As Boolean = False, Optional Mayusc As Boolean = False, Optional Tabula3 As Boolean = False)
    Select Case KeyAscii
          Case 13
               If Tabula Then
                  pSendKeys "{TAB}"
               End If
               If Tabula3 Then
                  pSendKeys "{TAB}" + "{TAB}" + "{TAB}"
               End If
          Case 8, 32, 209, 34
          Case 39
            KeyAscii = 180
          Case Asc("A") To Asc("Z")
          Case Asc("a") To Asc("z")
               If Mayusc Then
                  KeyAscii = Asc(UCase(Chr(KeyAscii)))
               End If
          Case 118, 241
               KeyAscii = 209
          Case Asc("0") To Asc("9")
            KeyAscii = 0
   End Select
End Sub

Public Sub ConsAlfaNumerico(KeyAscii As Integer, Optional Tabula As Boolean = False, Optional Mayusc As Boolean = False, Optional Tabula3 As Boolean = False)
   Select Case KeyAscii
          Case 13
               If Tabula Then
                  pSendKeys "{TAB}"
               End If
               If Tabula3 Then
                  pSendKeys "{TAB}" + "{TAB}" + "{TAB}"
               End If
          Case 8, 32, 209, 34
          Case 39
            KeyAscii = 180
          Case Asc("A") To Asc("Z")
          Case Asc("a") To Asc("z")
               If Mayusc Then
                  KeyAscii = Asc(UCase(Chr(KeyAscii)))
               End If
          Case 118, 241
               KeyAscii = 209
          Case Asc("0") To Asc("9")
          Case Else
               If InStr("Ñ\/!@#$%&*()=+-.,|:<>?{}[]´", Chr(KeyAscii)) > 0 Then
               Else
                  KeyAscii = 0
                  Beep
               End If
   End Select
End Sub

Public Function fIgv() As Double
    fIgv = NE(fRetornaValor("SELECT DBO.fBuscaConfOP ('" & gsEmpresa & "','" & gsAnio & "','053')"))
End Function
Public Function fRetHon() As Double
    fRetHon = NE(fRetornaValor("SELECT dbo.fxParTrib('RET', '000', '31/12/" & gsAnio & "')"))
End Function

Public Sub CrearDirectorio(directorio As String, Optional bMensaje As Boolean = True)
    On Error GoTo serror
    Dim fs
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Not (fs.FolderExists(directorio)) Then
        fs.CreateFolder (directorio)
    End If
    Set fs = Nothing
    Exit Sub
serror:
    If bMensaje Then Mensajes "Crear Directorio: " & directorio & Salto(1) & Err.Description

End Sub

Public Sub EliminaArchivo(archivo As String, Optional bMensaje As Boolean = True)
    On Error GoTo serror
    Dim fs
    Set fs = CreateObject("Scripting.FileSystemObject")
    If (fs.FileExists(archivo)) Then
        Call Kill(archivo)
    End If
    Set fs = Nothing
    Exit Sub
serror:
    If bMensaje Then Mensajes "Eliminar archivo: " & archivo & Salto(1) & Err.Description

End Sub

Public Function ExisteArchivo(archivo As String, Optional bMensaje As Boolean = True) As Boolean
    On Error GoTo serror
    Dim fs
    Set fs = CreateObject("Scripting.FileSystemObject")
    ExisteArchivo = True
    If Not (fs.FileExists(archivo)) Then
        ExisteArchivo = False
        Set fs = Nothing
        Exit Function
    End If
    Exit Function
serror:
    If bMensaje Then Mensajes "Busqueda de archivo: " & archivo & Salto(1) & Err.Description
End Function

Public Function NombreMes(periodo As String)
    Select Case periodo
           Case "00": NombreMes = "APERTURA"
           Case "01": NombreMes = "ENERO"
           Case "02": NombreMes = "FEBRERO"
           Case "03": NombreMes = "MARZO"
           Case "04": NombreMes = "ABRIL"
           Case "05": NombreMes = "MAYO"
           Case "06": NombreMes = "JUNIO"
           Case "07": NombreMes = "JULIO"
           Case "08": NombreMes = "AGOSTO"
           Case "09": NombreMes = "SETIEMBRE"
           Case "10": NombreMes = "OCTUBRE"
           Case "11": NombreMes = "NOVIEMBRE"
           Case "12": NombreMes = "DICIEMBRE"
           Case "13": NombreMes = "AJUSTE"
           Case "14": NombreMes = "CIERRE"
    End Select

End Function


Public Sub ImprimirReporte(strNombreDelReporte As String, ByRef Parametros(), ByRef formulas(), _
                           Optional ByVal TituloVentana As String = "", _
                           Optional ByVal bPrintDirect As Boolean = False, _
                           Optional ByVal pOrientacion As Orientacion_Pagina = Orientacion_Pagina.defecto, _
                           Optional ByVal pTipoPagina As Tipo_Pagina = Tipo_Pagina.defecto)
    If gsGenDsn = "SI" Then
        Call CrearDsn(gsServidor, gsBDUS, gsBDPW, gsBD, gsAutenticacion)
    End If

    Dim strRutaDelReporte As String
    Dim oPreviewReport As frmReportPreview
    Dim i As Integer
    Dim n As Integer
    Dim j As Integer
    'REPORTE
    Dim crApp As New CRAXDRT.Application
    Dim crRept As CRAXDRT.Report
    Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions
    Dim crParamDef As CRAXDRT.ParameterFieldDefinition
    Dim crDBTab As CRAXDRT.DatabaseTable
    'SUB REPORTE
    Dim SubReport As CRAXDRT.Report
    Dim Sections As CRAXDRT.Sections
    Dim Section As CRAXDRT.Section
    Dim RepObjs As CRAXDRT.ReportObjects
    Dim SubReportObj As CRAXDRT.SubreportObject
    
    Dim crpFormula As FormulaFieldDefinition
    Dim crpFormulas As FormulaFieldDefinitions
    
    If UCase(Right(strNombreDelReporte, 4)) = ".RPT" Then strNombreDelReporte = Mid(strNombreDelReporte, 1, Len(strNombreDelReporte) - 4)
    strRutaDelReporte = strNombreDelReporte & ".RPT"
    
    On Error GoTo ErrorImpresion
      
    ' ABRIR ARCHIVO DE REPORTE
    Set crRept = crApp.OpenReport(strRutaDelReporte)
    ' INGRESAR SQL server
        
    
    crRept.Database.LogOnServer "p2sODBC.dll", gsDSN, gsBD, gsBDUS, gsBDPW
    
        
        
    'PARA CRYSTAL 11.0
'    For i = 1 To crRept.Database.Tables.Count
'        crRept.Database.Tables(i).SetLogOnInfo gsDSN, gsBD, gsBDUS, gsBDPW
'        crRept.Database.Tables(i).Location = gsBD & ".dbo." & Mid(crRept.Database.Tables(i).Location, 6, Len(crRept.Database.Tables(i).Location) - 6)
'    Next
    
    For i = 1 To crRept.Database.Tables.Count
        crRept.Database.Tables(i).SetLogOnInfo gsDSN, gsBD, gsBDUS, gsBDPW
        crRept.Database.Tables(i).Location = gsBD & Mid(crRept.Database.Tables(i).Location, _
        InStr(1, crRept.Database.Tables(i).Location, "."), Len(crRept.Database.Tables(i).Location) - 1)
    Next
        
    
    'PARA SUB REPORTES
    Set Sections = crRept.Sections
'    If Sections.Count > 0 Then
'        MsgBox ""
'    End If
    For n = 1 To Sections.Count   'let's verify all sections,
      Set Section = Sections.item(n)
      Set RepObjs = Section.ReportObjects
      For i = 1 To RepObjs.Count
         If RepObjs.item(i).Kind = crSubreportObject Then
              Set SubReportObj = RepObjs.item(i)
              Set SubReport = SubReportObj.OpenSubreport
              For j = 1 To SubReport.Database.Tables.Count
                SubReport.Database.Tables(j).SetLogOnInfo gsDSN, gsBD, gsBDUS, gsBDPW
                SubReport.Database.Tables(j).Location = gsBD & Mid(SubReport.Database.Tables(j).Location, _
                InStr(1, SubReport.Database.Tables(j).Location, "."), _
                Len(SubReport.Database.Tables(j).Location) - 1)
              Next j
         End If
      Next i
     Next n
           
    ' DESAHABILITAR LA SOLICITUD DE PARAMETROS AL USUARIO FINAL
    crRept.EnableParameterPrompting = False
    ' OBTENER LA LISTA DE PARAMETROS DISPONIBLES PARA EL REPORTE
    Set crParamDefs = crRept.ParameterFields
    
    For Each crParamDef In crParamDefs
        crParamDef.SetCurrentValue BuscaValor(crParamDef.ParameterFieldName, crParamDef.ValueType, Parametros())
    Next
    
      ' OBTENER LA LISTA DE FORMULAS DISPONIBLES PARA EL REPORTE
    Set crpFormulas = crRept.FormulaFields

    For Each crpFormula In crpFormulas
        If BuscaFormula(crpFormula.FormulaFieldName, formulas()) Then
            crpFormula.Text = BuscaValorformula(crpFormula.FormulaFieldName, formulas())
        End If
    Next
        
    
'    Debug.Print "Printer Name: " & gsPrnPrinterName
'    Debug.Print "Printer Port: " & gsPrnPortName
'    Debug.Print "Printer Driver: " & gsPrnDriverName

    If Trim(gsPrnDriverName) <> "" And Trim(gsPrnPrinterName) <> "" And Trim(gsPrnPortName) <> "" Then
        crRept.SelectPrinter gsPrnDriverName, gsPrnPrinterName, gsPrnPortName
    End If
           
    
    If bPrintDirect Then
        '*** ENVIAR DIRECTO A LA IMPRESORA
        crRept.PrintOut False
        '***
    Else
        Set oPreviewReport = New frmReportPreview
        oPreviewReport.Visible = False
        oPreviewReport.Caption = TituloVentana
        oPreviewReport.Orientacion = pOrientacion
        oPreviewReport.TipoPagina = pTipoPagina
        oPreviewReport.SetReporte crRept
        Set oPreviewReport = Nothing
    End If
    Exit Sub
'
ErrorImpresion:
    Set crRept = Nothing
    Set frmReportPreview = Nothing

    If Err.Number = -2147189547 Then
        Mensajes "Se excedió la cantidad de reportes abiertos ciérrelos y vuelva a intentar"
        Exit Sub
    Else
        Mensajes Err.Description, vbCritical
    End If
    Resume
End Sub

'''Public Sub ImprimirReporte(strNombreDelReporte As String, ByRef Parametros(), ByRef formulas())
'''
'''    If gsGenDsn = "SI" Then
'''        Call CrearDsn(gsServidor, gsBDUS, gsBDPW, gsBD, gsAutenticacion)
'''    End If
'''
'''  Dim strRutaDelReporte As String
'''  Dim oPreviewReport As frmReportPreview
'''
'''  If UCase(Right(strNombreDelReporte, 4)) = ".RPT" Then strNombreDelReporte = Mid(strNombreDelReporte, 1, Len(strNombreDelReporte) - 4)
'''  strRutaDelReporte = strNombreDelReporte & ".RPT"
'''
'''  On Error GoTo ErrorImpresion
'''
'''  Dim crApp As New craxdrt.Application
'''  Dim crRept As craxdrt.Report
'''  Dim crParamDefs As craxdrt.ParameterFieldDefinitions
'''  Dim crParamDef As craxdrt.ParameterFieldDefinition
'''  Dim crDBTab As craxdrt.DatabaseTable
'''
'''  Dim crpFormula As FormulaFieldDefinition
'''  Dim crpFormulas As FormulaFieldDefinitions
'''
'''  ' Open Report File
'''  Set crRept = crApp.OpenReport(strRutaDelReporte)
'''
'''  ' Logon to SQL server
'''  'crRept.Database.LogOnServer "p2ssql.dll", gsServidor, gsBD, gsBDUS, gsBDPW
'''  crRept.Database.LogOnServerEx "p2sODBC.dll", gsDSN, gsBD, gsBDUS, gsBDPW
'''
'''  Dim i As Integer
'''  For i = 1 To crRept.Database.Tables.Count
'''    crRept.Database.Tables(i).SetLogOnInfo gsDSN, gsBD, gsBDUS, gsBDPW
'''    crRept.Database.Tables(i).Location = gsBD & Mid(crRept.Database.Tables(i).Location, InStr(1, crRept.Database.Tables(i).Location, "."), Len(crRept.Database.Tables(i).Location) - 1)
'''  Next i
'''
'''  ' Disable Parameter Prompting for the end user
'''  crRept.EnableParameterPrompting = False
'''
'''  ' Gather the list of available parameters from the report
'''  Set crParamDefs = crRept.ParameterFields
'''
'''  For Each crParamDef In crParamDefs
'''    crParamDef.SetCurrentValue BuscaValor(crParamDef.ParameterFieldName, crParamDef.ValueType, Parametros())
'''  Next
'''
'''  Set crpFormulas = crRept.FormulaFields
'''
'''  For Each crpFormula In crpFormulas
'''    If BuscaFormula(crpFormula.FormulaFieldName, formulas()) Then
'''       crpFormula.Text = BuscaValorformula(crpFormula.FormulaFieldName, formulas())
'''    End If
'''  Next
'''
'''  Set oPreviewReport = New frmReportPreview
'''  oPreviewReport.Visible = False
'''  oPreviewReport.Caption = "" 'strCaption
'''  oPreviewReport.SetReporte crRept
'''  Set oPreviewReport = Nothing
'''
'''  Exit Sub
'''
'''ErrorImpresion:
'''  Set crRept = Nothing
'''  Set frmReportPreview = Nothing
'''
'''  If Err.Number = -2147189547 Then
'''    Mensajes "Se excedió la cantidad de reportes abiertos ciérrelos y vuelva a intentar", vbOKOnly + vbInformation
'''  Else
'''    Mensajes Err.Description, vbCritical
'''  End If
'''
'''End Sub

Private Function BuscaFormula(cNombre As String, ByRef Parametros()) As Boolean
    Dim i As Integer
    Dim Pos As Integer
    On Error GoTo serror
    BuscaFormula = False
    
    For i = 0 To UBound(Parametros)

        If InStr(1, UCase(Parametros(i)), UCase(cNombre)) > 0 Then
            BuscaFormula = True
            
            Exit For
        End If
    Next
        
    Exit Function
serror:
    MsgBox Err.Description
End Function

Private Function BuscaValorformula(cParam As String, ByRef Parametros()) As String
    Dim i As Integer
    Dim Pos As Integer
    On Error GoTo serror

    For i = 0 To UBound(Parametros)

        If InStr(1, UCase(Parametros(i)), UCase(cParam)) > 0 Then
            Pos = InStr(1, Parametros(i), "=")
            BuscaValorformula = Mid(Parametros(i), Pos + 1, Len(Parametros(i)) - Pos)
            Exit For
        End If
    Next
    Exit Function
serror:
    MsgBox Err.Description
End Function

Private Function BuscaValor(cParam As String, cTipo As Integer, ByRef Parametros()) As Variant
    Dim i As Integer
    Dim Pos As Integer
    Dim PosFin As Integer
    Dim Param As String
    Dim Valor As Variant
    Dim Contenido As String

    On Error GoTo serror

    For i = 0 To UBound(Parametros)

        Contenido = CE(Parametros(i))
        
        If Contenido <> "" Then
            Pos = InStr(1, Contenido, ";")
            PosFin = InStr(Pos + 1, Contenido, ";")
    
            Param = Left(Contenido, Pos - 1)
            If PosFin > 0 Then
               Valor = Mid(Contenido, Pos + 1, PosFin - Pos - 1)
            Else
               Valor = ""
            End If
    
            If cParam = Param Then
                If cTipo >= 1 And cTipo <= 7 Then
                   BuscaValor = NE(Valor)
                Else
                   BuscaValor = CE(Valor)
                End If
    
                Exit For
            End If
        End If
    Next
    Exit Function
serror:
    MsgBox Err.Description
End Function

Public Sub Ayuda(sForm As String)
    'Debug.Print sForm
    Select Case sForm
        Case "frmManPlanCuentas": ShowTopicID 0, 3
        Case "frmManCentroCosto": ShowTopicID 0, 4
        Case "frmManLibros": ShowTopicID 0, 5
        Case "frmManTipoEntidad": ShowTopicID 0, 7
        Case "frmManCfgEntDoc": ShowTopicID 0, 8
        Case "frmManEntidades": ShowTopicID 0, 9
        Case "frmManTipoDocumento": ShowTopicID 0, 10
        Case "frmManTablas":
            If frmManTablas.tdbcTabla.BoundText = "044" Then ShowTopicID 0, 12 'percepciones
            If frmManTablas.tdbcTabla.BoundText = "045" Then ShowTopicID 0, 13 'retenciones
            If frmManTablas.tdbcTabla.BoundText = "047" Then ShowTopicID 0, 14 'perfiles
            If frmManTablas.tdbcTabla.BoundText = "022" Then ShowTopicID 0, 15 'tipos ratios
            If frmManTablas.tdbcTabla.BoundText = "003" Then ShowTopicID 0, 16 'docs identidad
        
        Case "frmManTipoMoneda": ShowTopicID 0, 17
        Case "frmManTipoCambio": ShowTopicID 0, 18
        Case "frmManBancos": ShowTopicID 0, 19
        Case "frmManCuentaCorriente": ShowTopicID 0, 20
        Case "frmManPlantillaTipoAsiento": ShowTopicID 0, 21
        Case "frmManPlantillaBalance": ShowTopicID 0, 22
        Case "frmManSubDiarioTDoc": ShowTopicID 0, 23
        Case "frmConfigOperaciones": ShowTopicID 0, 24
        
        Case "frmManAsientosContables": ShowTopicID 0, 28
        Case "frmManPresupuestos": ShowTopicID 0, 39
        Case "frmPrcRegistroCoaImp": ShowTopicID 0, 40
        Case "frmPrcRegistroCoaExp": ShowTopicID 0, 41
        
        Case "FrmManRegAuxiliarVentas": ShowTopicID 0, 42
        
        Case "frmPrcCambioCierre": ShowTopicID 0, 44
        Case "frmPrcCierreEjercicio": ShowTopicID 0, 45
        Case "frmPrcAsientoApertura": ShowTopicID 0, 46
        Case "frmPrcAsientoCierre": ShowTopicID 0, 47
        Case "frmPrcCierreMes": ShowTopicID 0, 48
        Case "frmPrcConversionMoneda": ShowTopicID 0, 49

        Case "frmManEstractoBancario": ShowTopicID 0, 51
        Case "frmRepMovimientosBancos": ShowTopicID 0, 52
        Case "frmRepSeguimientoCheques": ShowTopicID 0, 53
        Case "frmRepChequesPendientes": ShowTopicID 0, 54
        
        Case "frmRepLibroMayor": ShowTopicID 0, 58
        Case "frmRepSaldosNetos": ShowTopicID 0, 60
        Case "frmRepSaldosCuenta": ShowTopicID 0, 61
        
        Case "frmRepBalanceSumasSaldos": ShowTopicID 0, 64
        Case "frmRepBalanceResultadosConasev": ShowTopicID 0, 65
        Case "frmRepRegistroCompras": ShowTopicID 0, 67
        Case "frmRepRegistroRetencion": ShowTopicID 0, 68
        
        Case "frmRepAnaliticoProveedores": ShowTopicID 0, 70
        Case "frmRepAsientosxLibros": ShowTopicID 0, 71
        Case "frmRepAsientosCCosto": ShowTopicID 0, 73
        Case "frmRepResumenCentrocosto": ShowTopicID 0, 74
        Case "frmRepResumenCentroCostoMes": ShowTopicID 0, 75
        Case "frmRepCoa": ShowTopicID 0, 76
        Case "frmRepDaot": ShowTopicID 0, 77
        Case "frmRepPresupuestoEjecucion": ShowTopicID 0, 78
        
        Case "frmManIndicadores": ShowTopicID 0, 80
        Case "frmPrcRatios": ShowTopicID 0, 81

        Case "frmPrcExportarDatos": ShowTopicID 0, 84
        Case "frmPrcImportarDatosSistema": ShowTopicID 0, 85
        Case "frmPrcEliminaImportaciones": ShowTopicID 0, 86
        Case "frmConAuditoriaAsientos": ShowTopicID 0, 87
        Case "frmPrcExportarCoa": ShowTopicID 0, 88
        Case "frmPrcExportarDaot": ShowTopicID 0, 89
        Case "frmPrcBackup": ShowTopicID 0, 90
        Case "frmPrcRestore": ShowTopicID 0, 91
        Case "frmPrcReindex": ShowTopicID 0, 92
        Case "frmPrcActualizaSaldos": ShowTopicID 0, 93
        Case "frmPrcActualizaDestino": ShowTopicID 0, 94
        Case "frmManPassword": 'ShowTopicID 0, 95
        Case "frmManAnio": ShowTopicID 0, 96
        Case "frmManEmpresas": ShowTopicID 0, 97
        Case "frmManPerfilUsuario": ShowTopicID 0, 98
        Case "frmManUsrEmp": ShowTopicID 0, 99
        Case "frmManMenu": ShowTopicID 0, 100
        Case "frmManPerfiles": ShowTopicID 0, 101

        Case "frmManFlujoProceso": ShowTopicID 0, 104
        Case "frmManFlujoSaldos": ShowTopicID 0, 105
        Case "frmManFlujoReporte": ShowTopicID 0, 106
        Case "frmManPatrimonioNeto": ShowTopicID 0, 107
        Case "frmManCapital": ShowTopicID 0, 108
        
        
        Case "frmPrcFlujos": ShowTopicID 0, 130
        Case "frmRepPatrimonioNeto": ShowTopicID 0, 131
        
        Case "frmRepBalanceComprobacion":
                If Not IsFormato(frmRepBalanceComprobacion.Caption, "Formato") Then ShowTopicID 0, 63
                If IsFormato(frmRepBalanceComprobacion.Caption, "Formato") Then ShowTopicID 0, 129
                
        Case "frmRepAnexoInvBalance":
                'If Not IsFormato(frmRepAnexoInvBalance.Caption, "Formato") Then ShowTopicID 0, 66
                If IsFormato(frmRepAnexoInvBalance.Caption, "Formato 3.1") Then ShowTopicID 0, 116
                If IsFormato(frmRepAnexoInvBalance.Caption, "Formato 3.2") Then ShowTopicID 0, 117
                If IsFormato(frmRepAnexoInvBalance.Caption, "Formato 3.4") Then ShowTopicID 0, 118
                If IsFormato(frmRepAnexoInvBalance.Caption, "Formato 3.5") Then ShowTopicID 0, 119
                If IsFormato(frmRepAnexoInvBalance.Caption, "Formato 3.6") Then ShowTopicID 0, 120
                If IsFormato(frmRepAnexoInvBalance.Caption, "Formato 3.8") Then ShowTopicID 0, 121
                If IsFormato(frmRepAnexoInvBalance.Caption, "Formato 3.10") Then ShowTopicID 0, 122
                If IsFormato(frmRepAnexoInvBalance.Caption, "Formato 3.11") Then ShowTopicID 0, 123
                If IsFormato(frmRepAnexoInvBalance.Caption, "Formato 3.12") Then ShowTopicID 0, 124
                If IsFormato(frmRepAnexoInvBalance.Caption, "Formato 3.13") Then ShowTopicID 0, 125
                If IsFormato(frmRepAnexoInvBalance.Caption, "Formato 3.14") Then ShowTopicID 0, 126
                If IsFormato(frmRepAnexoInvBalance.Caption, "Formato 3.15") Then ShowTopicID 0, 127
                If IsFormato(frmRepAnexoInvBalance.Caption, "Formato 3.16") Then ShowTopicID 0, 128
                If IsFormato(frmRepAnexoInvBalance.Caption, "Formato 3.20") Then ShowTopicID 0, 132
                If IsFormato(frmRepAnexoInvBalance.Caption, "Formato 4.1") Then ShowTopicID 0, 133
                If IsFormato(frmRepAnexoInvBalance.Caption, "Formato 8.1") Then ShowTopicID 0, 136
                If IsFormato(frmRepAnexoInvBalance.Caption, "Formato 14.1") Then ShowTopicID 0, 141
                If IsFormato(frmRepAnexoInvBalance.Caption, "Formato 10.1") Then ShowTopicID 0, 138
                If IsFormato(frmRepAnexoInvBalance.Caption, "Formato 10.2") Then ShowTopicID 0, 139
                If IsFormato(frmRepAnexoInvBalance.Caption, "Formato 10.3") Then ShowTopicID 0, 140
                
                If IsFormato(frmRepAnexoInvBalance.Caption, "Formato 3.3") Then ShowTopicID 0, 146
                If InStr(1, frmRepAnexoInvBalance.Caption, "naturaleza") > 0 Then ShowTopicID 0, 147
               
        Case "frmManCostos": ShowTopicID 0, 143
        Case "frmManInvProc": ShowTopicID 0, 144
        Case "frmManFlujosCuentas": ShowTopicID 0, 145
        
        
        Case "frmRepLibroBancos":
                If Not IsFormato(frmRepLibroBancos.Caption, "Formato") Then ShowTopicID 0, 69
                If IsFormato(frmRepLibroBancos.Caption, "Formato 1.1") Then ShowTopicID 0, 113
                If IsFormato(frmRepLibroBancos.Caption, "Formato 1.2") Then ShowTopicID 0, 114
                    
        Case "frmRepLibroDiario":
                If Not IsFormato(frmRepLibroDiario.Caption, "Formato") Then ShowTopicID 0, 57
                If IsFormato(frmRepLibroDiario.Caption, "Formato") Then ShowTopicID 0, 134
                
        Case "frmRepLibroMayorAnalitico":
                If IsFormato(frmRepLibroMayorAnalitico.Caption, "Formato") Then ShowTopicID 0, 59
                If IsFormato(frmRepLibroMayorAnalitico.Caption, "Formato") Then ShowTopicID 0, 135
        Case "frmManValores":
                ShowTopicID 0, 110
                
        Case "frmPrcImportarDatosDbf":
                ShowTopicID 0, 148
                
        
    End Select
End Sub

Private Function IsFormato(sCaption As String, scadena) As Boolean
    If Left(sCaption, Len(scadena)) = scadena Then
        IsFormato = True
    Else
        IsFormato = False
    End If
End Function

Public Function GrabarTCDAOT(nvalor As Double, Optional bMensaje As Boolean = True) As Boolean
    GrabarTCDAOT = False
    On Local Error GoTo ErrorEjecucion
    Dim lArrMnt(9) As Variant
    Dim clsMante As clsMantoTablas
    Set clsMante = New clsMantoTablas
    
    lArrMnt(0) = "EDITAR_MIX"
    lArrMnt(1) = gsEmpresa
    lArrMnt(2) = gsAnio
    lArrMnt(3) = "053"
    lArrMnt(4) = CE(nvalor)
    lArrMnt(5) = 0
    lArrMnt(6) = "A"
    lArrMnt(7) = gsUsuario
    lArrMnt(8) = gsUsuario
    lArrMnt(9) = "1"
    
    If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCND_CONFIG_OPERA", lArrMnt(), True) = False Then
        If bMensaje = True Then
            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
        End If
        Exit Function
    End If

    GrabarTCDAOT = True
    
    If bMensaje = True Then
        Mensajes "El tipo de cambio DAOT, fue grabado correctamente"
    End If
    
    Set clsMante = Nothing
    
    Exit Function
ErrorEjecucion:
    GrabarTCDAOT = False
    Mensajes Err.Description
End Function

Public Function VerificaTcDAOT(sFechaIni As String, sFechaFin As String, nMonto As Double, sMoneda, sTipo As String) As Boolean
    On Error GoTo serror
    Dim clDatos As New clsMantoTablas
    Dim arrDatos() As Variant
    Dim sqlver As String
    Dim rs As New ADODB.Recordset
    Dim sMensaje As String
    
    sqlver = "exec spCn_RptDaot 'VALIDACION', '" & gsEmpresa & "', '" & gsAnio & "', '" & sFechaIni & "'," & _
             "'" & sFechaFin & "', " & nMonto & ", '" & sMoneda & "', '" & sTipo & "'"
    
    arrDatos = Array(sqlver)
    Set rs = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    sMensaje = ""
    
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then
            If rs.Fields("Reporte") <> "OK" Then
                sMensaje = "Faltan los tipos de cambio de Venta Publicación" & Salto(1) & "de las siguientes Fechas" & Salto(2)
                Do While Not rs.EOF
                    sMensaje = sMensaje & " FECHA: " & rs.Fields("Asd_dFecDoc") & ", VOUCHER: " & rs.Fields("Ase_nVoucher") & Salto(1)
                    rs.MoveNext
                Loop
            End If
        End If
    End If
    
    If sMensaje = "" Then
        VerificaTcDAOT = True
    Else
        Mensajes sMensaje
        VerificaTcDAOT = False
'        Resume
    End If
    
    CerrarRecordSet rs
    Set clDatos = Nothing
    Exit Function
serror:
   Mensajes Err.Description
   Resume
   VerificaTcDAOT = True
   CerrarRecordSet rs
    Set clDatos = Nothing
     
End Function


Public Function fRetornaRS(ByVal sql As String, Optional ByVal nMode As Integer = 0) As ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Set clDatos = New clsMantoTablas
    Dim arrDatos() As Variant
    arrDatos = Array(sql)
    Set fRetornaRS = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
   
    Set clDatos = Nothing
End Function

Public Function SetRsBookMarkOfCol(ByVal cadBusca As String, ByRef rs As ADODB.Recordset, ByVal Col As Integer) As Long
    On Error Resume Next
    SetRsBookMarkOfCol = 0
    rs.MoveFirst
    If Trim(cadBusca) <> "" Then
        rs.Find rs(Col).Name & "=#" & cadBusca & "#"
        If rs.EOF Then If rs.BOF = False Then rs.MoveLast
    End If
    SetRsBookMarkOfCol = rs.Bookmark
End Function


Public Sub EnterTab(nTecla As Integer)
    If nTecla = 13 Then pSendKeys "{tab}"
End Sub

Public Function fRetornaValor(cadena As String) As Variant
    Dim ObjC As clsMantoTablas
    Dim rs As ADODB.Recordset
    Dim arr() As Variant
    arr = Array(cadena)
    Set ObjC = New clsMantoTablas
    Set rs = ObjC.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arr)
    If rs Is Nothing Then fRetornaValor = "": Exit Function
    
    fRetornaValor = IIf(IsNull(rs(0).Value), "NULL", rs(0).Value)
    CerrarRecordSet rs
    Set ObjC = Nothing
End Function

Public Sub Centrar_form(ByRef Formulario As Form)
If Formulario.Name <> "frmIntroEmpresa" Then
    Formulario.Left = (Screen.Width - Formulario.Width) / 2
    Formulario.Top = (frmMDIConta.ScaleHeight - Formulario.Height) / 2
Else
    Formulario.Left = (Screen.Width - Formulario.Width) / 2
    Formulario.Top = (Screen.Height - Formulario.Height) / 2
End If
End Sub

Public Sub Centrar_Objeto(ByRef Objeto As Object, ByRef Contenedor As Object, Optional nLeft As Long = 0, Optional nTop As Long = 0)
    Objeto.Left = (Contenedor.Width - Objeto.Width) / 2 + nLeft
    Objeto.Top = (Contenedor.Height - Objeto.Height) / 2 - 200 + nTop
End Sub

Public Sub CentrarTitulo(ByRef oTitulo As Label, ByRef Contenedor As Object, ByRef Formulario As Object)
    oTitulo.Caption = Formulario.Caption
    DoEvents
    Call Centrar_Objeto(oTitulo, Formulario)
    
    oTitulo.Top = Contenedor.Top - oTitulo.Height - 200
    oTitulo.Alignment = vbCenter
    DoEvents
End Sub

Public Function FE(Valor As Variant) As Variant
    If IsNull(Valor) Then
        FE = Null
    ElseIf Valor = "" Then
        FE = Null
    ElseIf Valor = "01/01/1900" Then
        FE = Null
    ElseIf Valor = "__/__/____" Then
        FE = Null
    Else
        FE = Trim(Valor)
    End If
End Function

Public Function LicEmpresa() As Boolean
    On Error GoTo serror
    Dim sqlSp As String
    Dim arrDatos() As Variant
    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    
    Set clDatos = New clsMantoTablas
    LicEmpresa = False
    sqlSp = "spCn_GrabaTipoCambio 'VALORES'"
    arrDatos = Array(sqlSp)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo.State = 0 Then Exit Function
'    If rsArreglo("movim") < rsArreglo("valor") Then
    LicEmpresa = True
    Call CerrarRecordSet(rsArreglo)
    Set clDatos = Nothing
    Exit Function
serror:
    LicEmpresa = False
    Call CerrarRecordSet(rsArreglo)
    Set clDatos = Nothing
End Function


Public Sub UpdateGrilla(ByRef tdbgGrilla As TDBGrid)
    On Error Resume Next
    tdbgGrilla.Update
    DoEvents
End Sub

Public Sub RefreshGrilla(ByRef tdbgGrilla As TDBGrid)
    On Error Resume Next
    tdbgGrilla.Refresh
    DoEvents
End Sub

Public Sub RebindGrilla(ByRef tdbgGrilla As TDBGrid)
    On Error Resume Next
    tdbgGrilla.ReBind
    DoEvents
End Sub

Public Function ValidaRegistroEcb() As Boolean

Dim myObject As DllEcbLicencia.ISerial
Set myObject = New DllEcbLicencia.Serial

Dim myObject1 As DllEcbLicencia.ISerial
Set myObject1 = New DllEcbLicencia.Serial


Dim clave_registrada As String
Dim clave_registrada0 As String
clave_registrada = Replace(LeerLicencia, " ", "") 'myObject.Decrypt(LeerLicencia)



Dim x As String
Dim Clave  As String

If Trim(clave_registrada) = "" Then
ValidaRegistroEcb = False
Else
 Dim mo As DllEcbLicencia.ISerial
  Set mo = New DllEcbLicencia.Serial
  x = mo.Decrypt(clave_registrada)
  Clave = myObject.GetUniqueID(True, False, False, False)
    
  If Trim(x) <> Trim(Clave) Then
ValidaRegistroEcb = False
ElseIf Trim(x) = Trim(Clave) Then
ValidaRegistroEcb = True
End If


End If


End Function

Public Function LeerLicencia() As String

Dim mo As DllEcbLicencia.ISerial
  Set mo = New DllEcbLicencia.Serial
  LeerLicencia = mo.GetLicencia
'  Dim reg As Object
'Set reg = CreateObject("WScript.Shell")
'Dim prueba As String
'
'prueba = reg.RegRead("HKEY_CURRENT_USER\Software\ecblicencia\licencia")
'
''reg.RegWrite "HKEY_CURRENT_USER\Software\ecblicencia22\licencia", "pruebadatos"


  
  
End Function

Public Function EscribirLicencia(licencia As String) As String

Dim mo As DllEcbLicencia.ISerial
  Set mo = New DllEcbLicencia.Serial
  mo.SetLicencia (licencia)
'
   
'     Dim sh As Object
'
'    Set sh = CreateObject("WScript.Shell")
'    sh.RegWrite "HKEY_CURRENT_USER\Software\ecblicencia\licencia", licencia
'    Set sh = Nothing



End Function

Public Function ValidaRegistro() As Boolean
    Set oRegistroLock = CreateObject("activelock1884.ActiveLock")
    Dim x As Boolean
    ValidaRegistro = False
    
        oRegistroLock.SoftwareName = "ECB-Cont"
        oRegistroLock.SoftwarePassword = "977611"
        oRegistroLock.RegistryHive = "HKLM"
        oRegistroLock.RegistryKey = "ECB"
        oRegistroLock.RegistryName = "ECB"
        oRegistroLock.RegistryPath = "ECBCont"
        oRegistroLock.LiberationKeyLength = 40
        oRegistroLock.SoftwareCodeLength = 16
        oRegistroLock.LockToHardDrive = True
        oRegistroLock.LockToWindowsSerial = True
        oRegistroLock.LockToComputerName = False
        oRegistroLock.LockToRandomNumber = False
        oRegistroLock.LockToMACAddress = False
        oRegistroLock.LockToCustomString = ""
        
        oRegistroLock.HashAlgorithm = htSHA1AA2
        oRegistroLock.UseDataLock = False
        oRegistroLock.RegCounterKey = "ECBCK"
        oRegistroLock.RegLastRunDateKey = "ECBLRDK"
        oRegistroLock.RegRandomKey = "ECBRK"
        oRegistroLock.RegLiberationKey = "ECBLK"
        oRegistroLock.RegInitialRunDateKey = "ECBIRDK"
        
        ValidaRegistro = oRegistroLock.RegisteredUser
        
End Function

Public Sub LlenaComboBaseImponiblePublico(ByRef tdbc As TDBCombo, Libro As String, Optional bSoloBASES As Boolean = False)
    DoEvents
    With tdbc
        .Clear
        
        If Libro = lsLibroCom Then
            .AddItem "" & ";" & "    "
            .AddItem "006" & ";" & " (A) DEST. A OP.GRAV Y/O EXPORTACION"
            .AddItem "007" & ";" & " (B) DEST. A OP.GRAV Y/O EXP. Y NO GRAV."
            .AddItem "008" & ";" & " (C) DEST. A OP. NO GRAVADAS"
            '.AddItem "017" & ";" & " CUENTA DE I.S.C."
            If bSoloBASES = False Then
                .AddItem "999" & ";" & " VALOR DE ADQUISICION NO GRAVADO"
                .AddItem "027" & ";" & " ICBP" 'frt_202011
                .AddItem "024" & ";" & " OTROS"
                '.AddItem "025" & ";" & " CTA. DE REINTEGRO"
                '.AddItem "026" & ";" & " CTA. DE HONORARIOS"
            End If
        End If
        
' ANTES : MANAGMMENT GROUP
'        If Libro = lsLibroVen Then
'            .AddItem "" & ";" & "    "
'            .AddItem "002" & ";" & " GRAVABLE VENTAS"
'
'            If bSoloBASES = False Then
'                .AddItem "021" & ";" & " EXPORTACIONES"
'                .AddItem "047" & ";" & " BONIF. Y TRANSF. GRATUITA"
'                .AddItem "997" & ";" & " BONIF. Y TRANSF. GRATUITA INAFECTA"
'                '.AddItem "017" & ";" & " CUENTA DE I.S.C."
'                .AddItem "998" & ";" & " EXONERADA"
'                .AddItem "999" & ";" & " INAFECTO"
'            End If
'        End If
        
' DESPUES : HTC
        If Libro = lsLibroVen Then
            .AddItem "" & ";" & "    "
            .AddItem "002" & ";" & " GRAVABLE VENTAS"

            If bSoloBASES = False Then
                .AddItem "021" & ";" & " EXPORTACIONES"
                .AddItem "998" & ";" & " EXONERADA"
                .AddItem "999" & ";" & " INAFECTO"
                .AddItem "027" & ";" & " ICBP" 'frt_202011
                .AddItem "024" & ";" & " OTROS" 'rmc20191001
            End If
        End If

        If Libro = lsLibroDiario Then
            .AddItem "" & ";" & "    "
            .AddItem "018" & ";" & " BASE IMPONIBLE DAOT"
        End If
        
        .Columns(0).Visible = False
        .Bookmark = 0
        .ListField = "column1"
        .BoundColumn = "column0"
    
    End With
End Sub

Public Sub LlenaComboIGVPublico(ByRef tdbc As TDBCombo, cadena As String)
    Dim i As Integer
    i = 1

    tdbc.Clear
    tdbc.AddItem "" + ";" + "<Seleccione Sistema / Regimen>"
    If Mid(cadena, 1, 1) = "N" Then tdbc.AddItem "N;DOCUMENTO DE REFERENCIA"

    Do While i <= Len(cadena)
       If Mid(cadena, i, 1) = "D" Then tdbc.AddItem "D;DETRACCIONES"
       If Mid(cadena, i, 1) = "P" Then tdbc.AddItem "P;PERCEPCIONES"
       If Mid(cadena, i, 1) = "R" Then tdbc.AddItem "R;RETENCIONES"
       i = i + 1
    Loop

    tdbc.Bookmark = 0
    tdbc.ListField = "column1"
    tdbc.BoundColumn = "column0"
    tdbc.ReBind

End Sub

Public Function FormularioCargado(NombreFormulario As String) As Boolean
Dim Formulario As Form
FormularioCargado = False
For Each Formulario In Forms
    If (UCase(Formulario.Name) = UCase(NombreFormulario)) Then
    FormularioCargado = True
    Exit For
    End If
Next
End Function

Public Function BuscaRutaLog() As String
    Dim cRuta As String
    cRuta = App.Path
    If Right(cRuta, 1) = "\" Then cRuta = Mid(cRuta, 1, Len(cRuta) - 1)
    cRuta = cRuta & "\LOG"
    Call CrearDirectorio(cRuta)
    
    BuscaRutaLog = cRuta
End Function

Public Sub EscribirLog(cMensaje As String, cNombreFormulario As String)
    If gsGenLog <> "SI" Then Exit Sub
    
    Dim cArchivoLOG As String
    Dim cRuta  As String
    On Error GoTo serror
    
    cArchivoLOG = BuscaRutaLog & "\ECB-Cont.log"
    
    If ExisteArchivo(cArchivoLOG) Then
        Open cArchivoLOG For Append As #1 Len = 220
    Else
        Open cArchivoLOG For Output As #1 Len = 220
    End If
    
    Print #1, "FEC:" & vbTab & Date & vbTab & _
              "HOR:" & vbTab & Time & vbTab & _
              "USR:" & vbTab & gsUsuario & vbTab & _
              "SRV:" & vbTab & gsServidor & vbTab & _
              "EMP:" & vbTab & gsEmpresa & vbTab & _
              "AÑO:" & vbTab & gsAnio & vbTab & _
              "FRM:" & vbTab & cNombreFormulario & vbTab & _
              "MEN:" & vbTab & cMensaje & Chr(10) + Chr(13)
    Close #1
    
    DoEvents
    
    Exit Sub
serror:
    Mensajes Err.Description
End Sub

Public Sub EscribirMovimLog(cNombreFormulario As String, cNombreSP As String, cValores As String, cNummov As String, cVoucher As String)
    If gsGenLogMov <> "SI" Then Exit Sub
    
    Dim cArchivoMovLOG As String
    Dim cRuta As String
    On Error GoTo serror

    cArchivoMovLOG = BuscaRutaLog & "\" & CStr(Year(Date)) & Right("00" & Month(Date), 2) & Right("00" & Day(Date), 2) & "_MOV_" & gsBD & ".log"
    
    If ExisteArchivo(cArchivoMovLOG) Then
        Open cArchivoMovLOG For Append As #1 Len = 220
    Else
        Open cArchivoMovLOG For Output As #1 Len = 220
    End If
    
    cNombreSP = Left(cNombreSP & "                                   ", 35)
    cNummov = Left(cNummov & "          ", 10)
    cVoucher = Left(cVoucher & "          ", 10)
    
    
    Print #1, "FEC:" & vbTab & Date & vbTab & _
              "HOR:" & vbTab & Time & vbTab & _
              "USR:" & vbTab & gsUsuario & vbTab & _
              "SRV:" & vbTab & gsServidor & vbTab & _
              "EMP:" & vbTab & gsEmpresa & vbTab & _
              "AÑO:" & vbTab & gsAnio & vbTab & _
              "SPN:" & vbTab & cNombreSP & vbTab & _
              "MOV:" & vbTab & cNummov & vbTab & _
              "VOU:" & vbTab & cVoucher & vbTab & _
              "PAR:" & vbTab & cValores & Chr(10) + Chr(13)
    Close #1
    
    DoEvents
    
    Exit Sub
serror:
    Mensajes Err.Description
End Sub

Public Sub EscribirConsultaLog(cNombreFormulario As String, cNombreSP As String, cValores As String, cNummov As String, cVoucher As String)
    If gsGenLogConsulta <> "SI" Then Exit Sub
    
    Dim cArchivoMovLOG As String
    Dim cRuta As String
    On Error GoTo serror

    cArchivoMovLOG = BuscaRutaLog & "\" & CStr(Year(Date)) & Right("00" & Month(Date), 2) & Right("00" & Day(Date), 2) & "_LEC_" & gsBD & ".log"
    
    If ExisteArchivo(cArchivoMovLOG) Then
        Open cArchivoMovLOG For Append As #1 Len = 220
    Else
        Open cArchivoMovLOG For Output As #1 Len = 220
    End If
    
    cNombreSP = Left(cNombreSP & "                                   ", 35)
    cNummov = Left(cNummov & "          ", 10)
    cVoucher = Left(cVoucher & "          ", 10)
    
    
    Print #1, "FEC:" & vbTab & Date & vbTab & _
              "HOR:" & vbTab & Time & vbTab & _
              "USR:" & vbTab & gsUsuario & vbTab & _
              "SRV:" & vbTab & gsServidor & vbTab & _
              "EMP:" & vbTab & gsEmpresa & vbTab & _
              "AÑO:" & vbTab & gsAnio & vbTab & _
              "SPN:" & vbTab & cNombreSP & vbTab & _
              "MOV:" & vbTab & cNummov & vbTab & _
              "VOU:" & vbTab & cVoucher & vbTab & _
              "PAR:" & vbTab & cValores & Chr(10) + Chr(13)
    Close #1
    
    DoEvents
    
    Exit Sub
serror:
    Mensajes Err.Description
End Sub

Public Sub ConfigurarBarraEstado()
        With frmMDIConta
            .stbMdi.Panels(1).Text = "  USUARIO: " & CE(gsUsuario) & "  "
            .stbMdi.Panels(2).Text = "  " & NombreMes(gsPeriodo) & " DEL " & gsAnio & "  "
            .stbMdi.Panels(3).Text = " EMPRESA: " & gsEmpresa & "  " '
            .stbMdi.Panels(4).Text = "   SERVIDOR: " & gsServidor & "   "
            .stbMdi.Panels(5).Text = "   BD: " & gsBD & "   "
            
            '.stbMdi.Panels(1).Width = 1000
            '.stbMdi.Panels(2).Width = 1000
            '.stbMdi.Panels(3).Width = 1000
            
            .stbMdi.Panels(1).AutoSize = sbrContents
            .stbMdi.Panels(2).AutoSize = sbrSpring
            .stbMdi.Panels(3).AutoSize = sbrContents
            .stbMdi.Panels(4).AutoSize = sbrContents
            .stbMdi.Panels(5).AutoSize = sbrContents
            
            .stbMdi.Panels(1).Alignment = sbrCenter
            .stbMdi.Panels(2).Alignment = sbrCenter
            .stbMdi.Panels(3).Alignment = sbrCenter
            .stbMdi.Panels(4).Alignment = sbrCenter
            .stbMdi.Panels(5).Alignment = sbrCenter
        End With
End Sub

Public Sub EnviarTecla(cCadena As String)
    On Error Resume Next
    SendKeys cCadena
End Sub

Public Sub SeleccionarChecks(bValor As Boolean, ByRef Formulario As Form)
   Dim ctrl As Control
   For Each ctrl In Formulario.Controls
      If TypeOf ctrl Is CheckBox Then
         If bValor = True Then
            ctrl.Value = vbChecked
         Else
            ctrl.Value = vbUnchecked
         End If
      End If
   Next ctrl
  
   Set ctrl = Nothing
End Sub
'HT : 20091030
Public Sub ConfigForm(frm As Form, ww As Integer, hh As Integer)
   frm.Icon = frmMDIConta.Icon
   frm.Width = ww
   frm.Height = hh
   Call Centrar_form(frm)
End Sub
'HT : 20091111
Public Function ValidaSoloNumeros(ByVal vnCode As Integer) As Boolean

   ' Valida sólo números.
   If vnCode >= 48 And vnCode <= 57 Then
      ValidaSoloNumeros = True
   ElseIf vnCode = vbKeyBack Then
      ValidaSoloNumeros = True
   ElseIf vnCode = vbKeyEscape Then
      ValidaSoloNumeros = True
   ElseIf vnCode = vbKeyReturn Then
      ValidaSoloNumeros = True
   Else
      ValidaSoloNumeros = False
   End If

End Function
'HT : 20091111
Public Function nGetIniValueAscii() As Integer
Beep
nGetIniValueAscii = 0

End Function

'HT : 20091111
Sub CallReports()
Dim nCopias As Integer
ReDim gvMeses(11)
   
gvMeses(0) = "Enero"
gvMeses(1) = "Febrero"
gvMeses(2) = "Marzo"
gvMeses(3) = "Abril"
gvMeses(4) = "Mayo"
gvMeses(5) = "Junio"
gvMeses(6) = "Julio"
gvMeses(7) = "Agosto"
gvMeses(8) = "Setiembre"
gvMeses(9) = "Octubre"
gvMeses(10) = "Noviembre"
gvMeses(11) = "Diciembre"
   
'----------------------------------
' Rutina que Invoca a los Reportes.
'----------------------------------
For nCopias = 1 To giCopias
   gsPagina = 0
'   If GsDestino = "Archivo" Then
'   Else
'   End If

   Select Case gsAccionRep
      Case 1: frmRepAnexoInvBalance.ReporteDiarioSimplificado
      Case 2: frmRepLibroBancos.ReporteLibBancosDetMovEfectivo
      Case 3: frmRepLibroBancos.ReporteLibBancosDetMovCtaCte
      Case 4: frmRepLibroDiario.ReporteLibDiario
      Case 5: frmRepLibroMayorAnalitico.ReporteMatMayGeneral
      Case 7: frmRepLibroDiario.ReporteLibDiario (28) 'frt_rvie
      'Case 6: frmRepLibroMayorAnalitico.ReporteMatMayAnalitico
   End Select
Next

End Sub

'PB: 20110820
'Función que verifica si existe un archivo
'Public Function ExisteArchivo(sNombreArchivo As String) As Boolean
'Dim AttrDev%
'
'On Error Resume Next
'
'   AttrDev = GetAttr(sNombreArchivo)
'   If Err.Number Then
'      Err.Clear
'      ExisteArchivo = False
'   Else
'      ExisteArchivo = True
'   End If
'End Function

'HT : 20091111
'Función que devuelve la Hora del Servidor

Public Function DevuelveHoraServidor() As String

    Dim sSQLGenCod As String
    
    Dim adoRSGenCod As ADODB.Recordset
    Set adoRSGenCod = New ADODB.Recordset
    
    sSQLGenCod = "spCn_HoraServidor"
    
    With adoRSGenCod
        .ActiveConnection = gsCadenaConexion
        .CursorType = adOpenStatic
        .CursorLocation = adUseClient
        .LockType = adLockReadOnly
        .Open sSQLGenCod
        Set .ActiveConnection = Nothing
    End With
    'DevuelveHoraServidor = Left$(adoRSGenCod.Fields(0).Value, 11)
    DevuelveHoraServidor = Time
   Set adoRSGenCod = Nothing
End Function
'HT : 20091111
Public Sub AlinearDosTextos(pColumnas As Integer, pTextIzquierda As String, ptextDerecha)
'-----------------------------------------------------------------------------------------------
' Objetivo  :  Alinear dos textos, uno a la derecha y otro a la izquierda en Impresora matricial
Dim cString As String
    cString = Space(pColumnas)
    RSet cString = ptextDerecha
    Mid(cString, 1) = pTextIzquierda
    printl cString
End Sub
Public Sub AlinearTextoDerecha(pColumnas As Integer, ptextDerecha As String)
'-----------------------------------------------------------------------------------------------
' Objetivo  :  Alinear texto a la derecha del reporte para Impresora matricial
Dim cString As String
    cString = Space(pColumnas)
    RSet cString = ptextDerecha
    printl cString
End Sub
'HT : 20091111
Public Sub CentrarTexto(pCadena As String, pColumnas As Integer)
'---------------------------------------------------
' Objetivo  :  Centrar texto en Impresora matricial

'---------------------------------------------------
Dim nCol As Integer
    nCol = CInt((pColumnas - Len(pCadena)) / 2)
    printl (Space(nCol) & pCadena)
End Sub
'HT : 20091111
Sub printl(gsLinea As String)

'Dim liQuiebre   As Integer
Dim lbIsPrinter As Boolean

On Error GoTo ERRORIMPRESORA
lbIsPrinter = True
If GsDestino = "Archivo" Or GsDestino = "PrintFile" Then lbIsPrinter = False

If Len(Trim$(gsLinea)) = 0 Then
   If lbIsPrinter Then
        If CDbl(xGs_DesdePag) = 0 And CDbl(xGs_HastaPag) = 0 And CDbl(xGs_Principal) = 0 Then
            If gsPagina + 1 >= Gs_DesdePag Then Printer.Print
        ElseIf CDbl(xGs_DesdePag) <> 0 And CDbl(xGs_HastaPag) <> 0 And CDbl(xGs_Principal) <> 0 Then
          If CDbl(xGs_DesdePag) = CDbl(xGs_HastaPag) And CDbl(gsPagina) = CDbl(xGs_DesdePag) And CDbl(xGs_Principal) = CDbl(gsPaginaPrincipal) Then
            If gsPagina + 1 >= Gs_DesdePag Then Printer.Print
          ElseIf CDbl(xGs_DesdePag) <> CDbl(xGs_HastaPag) Then
            If CDbl(gsPagina) >= CDbl(xGs_DesdePag) And CDbl(gsPagina) <= CDbl(xGs_HastaPag) And CDbl(xGs_Principal) = CDbl(gsPaginaPrincipal) Then
             If gsPagina + 1 >= Gs_DesdePag Then Printer.Print
            End If
          Else
            gsLinea = ""
          End If
        End If
   Else
        If CDbl(xGs_DesdePag) = 0 And CDbl(xGs_HastaPag) = 0 And CDbl(xGs_Principal) = 0 Then
            Print #1, " "
        ElseIf CDbl(xGs_DesdePag) <> 0 And CDbl(xGs_HastaPag) <> 0 And CDbl(xGs_Principal) <> 0 Then
          If CDbl(xGs_DesdePag) = CDbl(xGs_HastaPag) And CDbl(gsPagina) = CDbl(xGs_DesdePag) And CDbl(xGs_Principal) = CDbl(gsPaginaPrincipal) Then
            Print #1, " "
          ElseIf CDbl(xGs_DesdePag) <> CDbl(xGs_HastaPag) Then
            If CDbl(gsPagina) >= CDbl(xGs_DesdePag) And CDbl(gsPagina) <= CDbl(xGs_HastaPag) And CDbl(xGs_Principal) = CDbl(gsPaginaPrincipal) Then
             Print #1, " "
             ElseIf CDbl(gsPagina) >= CDbl(xGs_DesdePag) And CDbl(gsPagina) <= CDbl(xGs_HastaPag) Then
             Print #1, " "
            End If
          Else
            gsLinea = ""
          End If
        End If
   End If
Else
   If lbIsPrinter Then
        If CDbl(xGs_DesdePag) = 0 And CDbl(xGs_HastaPag) = 0 And CDbl(xGs_Principal) = 0 Then
            If gsPagina + 1 >= Gs_DesdePag Then Printer.Print gsLinea
        ElseIf CDbl(xGs_DesdePag) <> 0 And CDbl(xGs_HastaPag) <> 0 And CDbl(xGs_Principal) <> 0 Then
          If CDbl(xGs_DesdePag) = CDbl(xGs_HastaPag) And CDbl(gsPagina) = CDbl(xGs_DesdePag) And CDbl(xGs_Principal) = CDbl(gsPaginaPrincipal) Then
            If gsPagina + 1 >= Gs_DesdePag Then Printer.Print gsLinea
          ElseIf CDbl(xGs_DesdePag) <> CDbl(xGs_HastaPag) Then
            If CDbl(gsPagina) >= CDbl(xGs_DesdePag) And CDbl(gsPagina) <= CDbl(xGs_HastaPag) And CDbl(xGs_Principal) = CDbl(gsPaginaPrincipal) Then
             If gsPagina + 1 >= Gs_DesdePag Then Printer.Print gsLinea
            End If
          Else
            gsLinea = ""
          End If
        End If
   Else
      If CDbl(xGs_DesdePag) = 0 And CDbl(xGs_HastaPag) = 0 And CDbl(xGs_Principal) = 0 Then
             Print #1, gsLinea
        ElseIf CDbl(xGs_DesdePag) <> 0 And CDbl(xGs_HastaPag) <> 0 And CDbl(xGs_Principal) <> 0 Then
            'And CDbl(xGs_Principal) = CDbl(gsPaginaPrincipal)
          If CDbl(xGs_DesdePag) = CDbl(xGs_HastaPag) And CDbl(gsPagina) = CDbl(xGs_DesdePag) Then
             Print #1, gsLinea
          ElseIf CDbl(xGs_DesdePag) <> CDbl(xGs_HastaPag) Then
            'And CDbl(xGs_Principal) = CDbl(gsPaginaPrincipal)
            'CDbl(gsPagina) >= CDbl(xGs_DesdePag) And CDbl(gsPagina) <= CDbl(xGs_HastaPag)
            If CDbl(gsPagina) >= CDbl(xGs_DesdePag) And CDbl(gsPagina) <= CDbl(xGs_HastaPag) Then
             Print #1, gsLinea
            End If
          Else
           gsLinea = ""
          End If
        End If
        
   End If
End If
Debug.Print gsLinea

'Crear una Nueva hoja de impresion al encontrar la palabra TOTALES
'If (InStr(gsLinea, "TOTALES") > 0 Or InStr(gsLinea, "VAN...") > 0) Then
'   Print #1, vbCrLf 'Salto de linea para la Impresion
''    Printer.Print ""
''    Printer.NewPage
''    Printer.FontSize = 10
'End If

Select Case gsAccionRep
Case 1
 liQuiebre = 76
 giLineas = giLineas + 1
Case 2, 3, 4, 5, 7
 If giLineas >= 1000 Then
    liQuiebre = 2000
 Else
    liQuiebre = 76
 End If
 
 giLineas = giLineas + 1
End Select

If giLineas >= liQuiebre Then
   giLineas = 0
   If lbIsPrinter Then
      If Gs_DesdePag > 1 Then
         If CDbl(xGs_DesdePag) = 0 And CDbl(xGs_HastaPag) = 0 And CDbl(xGs_Principal) = 0 Then
            If gsPagina + 1 >= Gs_DesdePag Then
            Printer.Print ""
            'Printer.Print "sdfasfsafsfsdf"
            Printer.NewPage
            Printer.FontSize = 10
            Else
            Printer.KillDoc
            End If
         ElseIf CDbl(xGs_DesdePag) <> 0 And CDbl(xGs_HastaPag) <> 0 And CDbl(xGs_Principal) <> 0 Then
            If CDbl(xGs_DesdePag) = CDbl(xGs_HastaPag) And CDbl(gsPagina) = CDbl(xGs_DesdePag) And CDbl(xGs_Principal) = CDbl(gsPaginaPrincipal) Then
              If gsPagina + 1 >= Gs_DesdePag Then
              Printer.Print ""
              Printer.NewPage
              Printer.FontSize = 10
              Else
              Printer.KillDoc
              End If
            ElseIf CDbl(xGs_DesdePag) <> CDbl(xGs_HastaPag) Then
              If CDbl(gsPagina) >= CDbl(xGs_DesdePag) And CDbl(gsPagina) <= CDbl(xGs_HastaPag) And CDbl(xGs_Principal) = CDbl(gsPaginaPrincipal) Then
                If gsPagina + 1 >= Gs_DesdePag Then
                    Printer.Print ""
                    Printer.NewPage
                    Printer.FontSize = 10
                    Else
                    Printer.KillDoc
                End If
              End If
            End If
        End If
      Else
       If CDbl(xGs_DesdePag) = 0 And CDbl(xGs_HastaPag) = 0 And CDbl(xGs_Principal) = 0 Then
        Select Case gsAccionRep
        Case 1
         Printer.Print ""
         Printer.NewPage
         Printer.FontSize = 10
        End Select
       ElseIf CDbl(xGs_DesdePag) <> 0 And CDbl(xGs_HastaPag) <> 0 And CDbl(xGs_Principal) <> 0 Then
         If CDbl(xGs_DesdePag) = CDbl(xGs_HastaPag) And CDbl(gsPagina) = CDbl(xGs_DesdePag) And CDbl(xGs_Principal) = CDbl(gsPaginaPrincipal) Then
                Select Case gsAccionRep
                Case 1, 2, 3
                 Printer.Print ""
                 Printer.NewPage
                 Printer.FontSize = 10
                End Select
         ElseIf CDbl(xGs_DesdePag) <> CDbl(xGs_HastaPag) Then
            If CDbl(gsPagina) >= CDbl(xGs_DesdePag) And CDbl(gsPagina) <= CDbl(xGs_HastaPag) And CDbl(xGs_Principal) = CDbl(gsPaginaPrincipal) Then
                Select Case gsAccionRep
                Case 1, 2, 3
                 Printer.Print ""
                 Printer.NewPage
                 Printer.FontSize = 10
                End Select
            End If
         End If
       End If
       If gsAccionRep = 1 Then
          If (frmRepAnexoInvBalance.UltCol - 1) < (frmRepAnexoInvBalance.rsArreglo.Fields.Count - 1) Then
             If gsTipoImp <> "1" Then
              If Not frmRepAnexoInvBalance.rsArreglo.EOF And (frmRepAnexoInvBalance.UltCol - 1) < (frmRepAnexoInvBalance.rsArreglo.Fields.Count - 1) And gsPaginaPrincipal = 1 Then
               frmRepAnexoInvBalance.rsArreglo.MoveFirst
              Else
               frmRepAnexoInvBalance.rsArreglo.AbsolutePosition = frmRepAnexoInvBalance.PosReg
              End If
             Else
               frmRepAnexoInvBalance.rsArreglo.MoveFirst
                If Not frmRepAnexoInvBalance.rsArreglo.EOF And (frmRepAnexoInvBalance.UltCol - 1) < (frmRepAnexoInvBalance.rsArreglo.Fields.Count - 1) And gsPaginaPrincipal = 1 Then
                 'frmRepAnexoInvBalance.rsArreglo.MoveFirst
                 Debug.Print gsPaginaPrincipal & " " & gsPagina
                Else
                 frmRepAnexoInvBalance.rsArreglo.AbsolutePosition = frmRepAnexoInvBalance.PosReg
                End If
             End If
            'frmRepAnexoInvBalance.rsArreglo.MoveFirst
             Select Case gsAccionRep
             Case 1: Call frmRepAnexoInvBalance.CabeceraDiarioSimplificado: Call frmRepAnexoInvBalance.ImprimeDetalle
             End Select
            Exit Sub
          Else
            frmRepAnexoInvBalance.PosReg = frmRepAnexoInvBalance.rsArreglo.AbsolutePosition + 1
            Exit Sub
          End If
        ElseIf gsAccionRep = 2 Then
         If CDbl(xGs_DesdePag) = 0 And CDbl(xGs_HastaPag) = 0 And CDbl(xGs_Principal) = 0 Then Printer.NewPage
         Call frmRepLibroBancos.CabeceraDetEfectivo
        ElseIf gsAccionRep = 3 Then
         If CDbl(xGs_DesdePag) = 0 And CDbl(xGs_HastaPag) = 0 And CDbl(xGs_Principal) = 0 Then Printer.NewPage
         Call frmRepLibroBancos.CabeceraDetMovCtaCte
        ElseIf gsAccionRep = 4 Or gsAccionRep = 7 Then 'frt_rvie
         If CDbl(xGs_DesdePag) = 0 And CDbl(xGs_HastaPag) = 0 And CDbl(xGs_Principal) = 0 Then Printer.NewPage
         Call frmRepLibroDiario.CabeceraLibroDiario(IIf(gsAccionRep = 7, 28, 0)) 'frt_rvie
        ElseIf gsAccionRep = 5 Then
         If CDbl(xGs_DesdePag) = 0 And CDbl(xGs_HastaPag) = 0 And CDbl(xGs_Principal) = 0 Then Printer.NewPage
         Call frmRepLibroMayorAnalitico.CabeceraLibMayorGeneral
        End If
       End If
   Else
    Select Case gsAccionRep
    Case 1
         If (frmRepAnexoInvBalance.UltCol - 1) < (frmRepAnexoInvBalance.rsArreglo.Fields.Count - 1) Then
          If gsTipoImp <> "1" Then
           If Not frmRepAnexoInvBalance.rsArreglo.EOF And (frmRepAnexoInvBalance.UltCol - 1) < (frmRepAnexoInvBalance.rsArreglo.Fields.Count - 1) And gsPaginaPrincipal = 1 Then
            frmRepAnexoInvBalance.rsArreglo.MoveFirst
           Else
            'frmRepAnexoInvBalance.rsArreglo.AbsolutePosition = frmRepAnexoInvBalance.PosReg
           End If
          Else
           frmRepAnexoInvBalance.rsArreglo.MoveFirst
           If Not frmRepAnexoInvBalance.rsArreglo.EOF And (frmRepAnexoInvBalance.UltCol - 1) < (frmRepAnexoInvBalance.rsArreglo.Fields.Count - 1) And gsPaginaPrincipal = 1 Then
            'frmRepAnexoInvBalance.rsArreglo.MoveFirst
            Debug.Print gsPaginaPrincipal & " " & gsPagina
           Else
            frmRepAnexoInvBalance.rsArreglo.AbsolutePosition = frmRepAnexoInvBalance.PosReg
           End If
          End If
          'frmRepAnexoInvBalance.rsArreglo.MoveFirst
            Select Case gsAccionRep
            Case 1:
                Call frmRepAnexoInvBalance.ImprimeTotales
                Call frmRepAnexoInvBalance.CabeceraDiarioSimplificado
                If VarGsIndDS = True Then
                    Call frmRepAnexoInvBalance.ImprimeDetalle
                Else
                    GoTo SALTAR
                End If
            End Select
        
           Exit Sub
         Else
           frmRepAnexoInvBalance.PosReg = frmRepAnexoInvBalance.rsArreglo.AbsolutePosition + 1
           Exit Sub
         End If
    Case 2 ' Caja detalle efectivo
     If CDbl(xGs_DesdePag) = 0 And CDbl(xGs_HastaPag) = 0 And CDbl(xGs_Principal) = 0 Then
        
       ' Print #1, "": Print #1, "": Print #1, "": Print #1, "": Print #1, ""
        
        'Print #1, "SALTO DE PAGINA"
     ElseIf CDbl(xGs_DesdePag) <> 0 And CDbl(xGs_HastaPag) <> 0 And CDbl(xGs_Principal) <> 0 Then
       If CDbl(xGs_DesdePag) = CDbl(xGs_HastaPag) And CDbl(gsPagina) = CDbl(xGs_DesdePag) And CDbl(xGs_Principal) = CDbl(gsPaginaPrincipal) Then
        
      '  Print #1, "": Print #1, "": Print #1, "": Print #1, "": Print #1, ""
        
        'Print #1, "SALTO DE PAGINA"
       'ElseIf CDbl(xGs_DesdePag) < CDbl(xGs_HastaPag) And CDbl(gsPagina) = CDbl(xGs_DesdePag) And CDbl(xGs_Principal) = CDbl(gsPaginaPrincipal) Then
        ElseIf CDbl(xGs_DesdePag) < CDbl(xGs_HastaPag) And CDbl(xGs_Principal) = CDbl(gsPaginaPrincipal) Then
        
        'Print #1, "": Print #1, "": Print #1, "": Print #1, "": Print #1, ""
        
        'Print #1, "SALTO DE PAGINA"
       End If
     End If
      If SwHoja Then Call frmRepLibroBancos.CabeceraDetEfectivo Else SwHoja = False
    Case 3 'Caja movimiento Efectivo
     If CDbl(xGs_DesdePag) = 0 And CDbl(xGs_HastaPag) = 0 And CDbl(xGs_Principal) = 0 Then
       ' Print #1, Chr(27) & Chr(12)
        'If Not gsNomTipoImp Then
        'Print #1, "": Print #1, "": Print #1, ""
        'Print #1, "SALTO DE PAGINA"
     ElseIf CDbl(xGs_DesdePag) <> 0 And CDbl(xGs_HastaPag) <> 0 And CDbl(xGs_Principal) <> 0 Then
       If CDbl(xGs_DesdePag) < CDbl(xGs_HastaPag) And CDbl(gsPagina) = CDbl(xGs_DesdePag) And CDbl(xGs_Principal) = CDbl(gsPaginaPrincipal) Then
        'Print #1, Chr(27) & Chr(12)
        'If Not gsNomTipoImp Then
        'Print #1, "": Print #1, "": Print #1, ""
        'Print #1, "SALTO DE PAGINA"
       'ElseIf CDbl(xGs_DesdePag) < CDbl(xGs_HastaPag) And CDbl(gsPagina) = CDbl(xGs_DesdePag) And CDbl(xGs_Principal) = CDbl(gsPaginaPrincipal) Then
        'ElseIf CDbl(gsPagina) >= CDbl(xGs_DesdePag) And CDbl(gsPagina) <= CDbl(xGs_HastaPag) Then
        'Print #1, Chr(27) & Chr(12)
        'If Not gsNomTipoImp Then
        'Print #1, "": Print #1, "": Print #1, ""
        'Print #1, "SALTO DE PAGINA"
       End If
     End If
      If SwHoja Then Call frmRepLibroBancos.CabeceraDetMovCtaCte Else SwHoja = False
    Case 4, 7 ' Libro Diario 'frt_rvie
     If CDbl(xGs_DesdePag) = 0 And CDbl(xGs_HastaPag) = 0 And CDbl(xGs_Principal) = 0 Then
        'Print #1, Chr(27) & Chr(12)
        'If Not gsNomTipoImp Then

        'Print #1, "": Print #1, "": Print #1, "": Print #1, ""
        
        'Print #1, "SALTO DE PAGINA"
     ElseIf CDbl(xGs_DesdePag) <> 0 And CDbl(xGs_HastaPag) <> 0 And CDbl(xGs_Principal) <> 0 Then
        If CDbl(xGs_DesdePag) < CDbl(xGs_HastaPag) And CDbl(gsPagina) = CDbl(xGs_DesdePag) And CDbl(xGs_Principal) = CDbl(gsPaginaPrincipal) Then
            'Print #1, Chr(27) & Chr(12)
            'If Not gsNomTipoImp Then

         '   Print #1, "": Print #1, "": Print #1, "": Print #1, ""
        
        'Print #1, "SALTO DE PAGINA"
       'ElseIf CDbl(xGs_DesdePag) < CDbl(xGs_HastaPag) And CDbl(gsPagina) = CDbl(xGs_DesdePag) And CDbl(xGs_Principal) = CDbl(gsPaginaPrincipal) Then
        ElseIf CDbl(gsPagina) >= CDbl(xGs_DesdePag) And CDbl(gsPagina) <= CDbl(xGs_HastaPag) Then 'CDbl(xGs_DesdePag) < CDbl(xGs_HastaPag) And CDbl(xGs_Principal) = CDbl(gsPaginaPrincipal) Then
            'Print #1, Chr(27) & Chr(12)
            'If Not gsNomTipoImp Then

         '   Print #1, "": Print #1, "": Print #1, "": Print #1, ""
'Print #1, "SALTO DE PAGINA"
        End If
     End If
     Call frmRepLibroDiario.CabeceraLibroDiario(IIf(gsAccionRep = 7, 28, 0)) 'frt_rvie
    Case 5 ' Libro mayor
     If CDbl(xGs_DesdePag) = 0 And CDbl(xGs_HastaPag) = 0 And CDbl(xGs_Principal) = 0 Then
        'Print #1, Chr(27) & Chr(12)
        'If Not gsNomTipoImp Then
        
       ' Print #1, "": Print #1, "": Print #1, "": Print #1, "": Print #1, ""
        
        'Print #1, "SALTO DE PAGINA"
     ElseIf CDbl(xGs_DesdePag) <> 0 And CDbl(xGs_HastaPag) <> 0 And CDbl(xGs_Principal) <> 0 Then
       If CDbl(xGs_DesdePag) < CDbl(xGs_HastaPag) And CDbl(gsPagina) = CDbl(xGs_DesdePag) And CDbl(xGs_Principal) = CDbl(gsPaginaPrincipal) Then
        'Print #1, Chr(27) & Chr(12)
        'If Not gsNomTipoImp Then
        
       ' Print #1, "": Print #1, "": Print #1, "": Print #1, "": Print #1, ""
        
        'Print #1, "SALTO DE PAGINA"
       'ElseIf CDbl(xGs_DesdePag) < CDbl(xGs_HastaPag) And CDbl(gsPagina) = CDbl(xGs_DesdePag) And CDbl(xGs_Principal) = CDbl(gsPaginaPrincipal) Then
        ElseIf CDbl(gsPagina) >= CDbl(xGs_DesdePag) And CDbl(gsPagina) <= CDbl(xGs_HastaPag) Then ' CDbl(xGs_DesdePag) < CDbl(xGs_HastaPag) And CDbl(xGs_Principal) <> CDbl(gsPaginaPrincipal) Then
        'Print #1, Chr(27) & Chr(12)
        'If Not gsNomTipoImp Then
        
       ' Print #1, "": Print #1, "": Print #1, "": Print #1, "": Print #1, ""
       
       'Print #1, "SALTO DE PAGINA"
       End If
     End If
      Call frmRepLibroMayorAnalitico.CabeceraLibMayorGeneral
    End Select
   End If
End If
SALTAR:
'giLineas = 0
Exit Sub
ERRORIMPRESORA:
    If Err.Number = 482 Then
       MsgBox Err.Description, vbCritical, App.Title
       Resume
    End If
End Sub


'HT : 20091111
'-------------------------------------------------------------
' Procedure : SelectedText()
' Propósito : Selecciona el Texto de un TextBox
'-------------------------------------------------------------
Sub SelectedText(poTextBox As TextBox, Optional ValorBool As Boolean)
   poTextBox.SelStart = 0
   poTextBox.SelLength = Len(poTextBox.Text)
   If ValorBool Then poTextBox.BackColor = &H80000018
End Sub

Public Function Fct_Obt_Sumas_DH_Sg_Periodo(ByRef Par_EmpCodigo As String, _
ByRef Par_Anio As String, ByRef Par_Periodo As String, ByRef Par_Pla_cCuentaContable As String) As Recordset
    Dim VarCmd As New ADODB.Command
    
    With VarCmd
        .CommandType = adCmdStoredProc
        .CommandText = "Usp_Obt_Sumas_DH_Sg_Periodo"
        .ActiveConnection = gsCadenaConexion
        .Parameters.Append .CreateParameter("@Pe_EmpCodigo", adVarChar, adParamInput, 4, Par_EmpCodigo)
        .Parameters.Append .CreateParameter("@Pe_Anio", adVarChar, adParamInput, 4, Par_Anio)
        .Parameters.Append .CreateParameter("@Pe_Periodo", adVarChar, adParamInput, 4, Par_Periodo)
        .Parameters.Append .CreateParameter("@Pla_cCuentaContable", adVarChar, adParamInput, 12, Par_Pla_cCuentaContable)
    
        Set Fct_Obt_Sumas_DH_Sg_Periodo = .Execute
    End With
    
    Set VarCmd = Nothing
End Function

Public Function Fct_Obtener_Periodos_Sg_Anio_Empresa(ByRef Par_EmpCodigo As String, _
ByRef Par_cAnio As String) As Recordset
    Dim VarCmd As New ADODB.Command
    
    With VarCmd
        .CommandType = adCmdStoredProc
        .CommandText = "Usp_Obtener_Periodos_Sg_Anio_Empresa"
        .ActiveConnection = gsCadenaConexion
        .Parameters.Append .CreateParameter("@Pe_EmpCodigo", adVarChar, adParamInput, 4, Par_EmpCodigo)
        .Parameters.Append .CreateParameter("@Pe_cAnio", adVarChar, adParamInput, 4, Par_cAnio)
    
        Set Fct_Obtener_Periodos_Sg_Anio_Empresa = .Execute
    End With
    
    Set VarCmd = Nothing
End Function
    
Public Function Fct_Obtener_Cuentas_Inversion(ByRef Par_EmpCodigo As String, _
ByRef Par_cAnio As String) As Recordset
    Dim VarCmd As New ADODB.Command
    
    With VarCmd
        .CommandType = adCmdStoredProc
        .CommandText = "Usp_Obtener_Cuentas_Inversion"
        .ActiveConnection = gsCadenaConexion
        .Parameters.Append .CreateParameter("@Pe_EmpCodigo", adVarChar, adParamInput, 4, Par_EmpCodigo)
        .Parameters.Append .CreateParameter("@Pe_cAnio", adVarChar, adParamInput, 4, Par_cAnio)
    
        Set Fct_Obtener_Cuentas_Inversion = .Execute
    End With
    
    Set VarCmd = Nothing
End Function

Public Function Fct_Obtener_Lista_de_Movimiento(ByRef Par_EmpCodigo As String, _
ByRef Par_cAnio As String, ByRef Par_Cuenta As String, ByRef Par_Periodo As String, ByRef Par_NumFila As Integer) As Recordset

    Dim VarCmd As New ADODB.Command
    
    With VarCmd
        .CommandType = adCmdStoredProc
        .CommandText = "Usp_Obtener_Lista_de_Movimiento"
        .ActiveConnection = gsCadenaConexion
        .Parameters.Append .CreateParameter("@Pe_EmpCodigo", adVarChar, adParamInput, 4, Par_EmpCodigo)
        .Parameters.Append .CreateParameter("@Pe_cAnio", adVarChar, adParamInput, 4, Par_cAnio)
        .Parameters.Append .CreateParameter("@Pe_Cuenta", adVarChar, adParamInput, 2, Par_Cuenta)
        .Parameters.Append .CreateParameter("@Pe_Periodo", adVarChar, adParamInput, 2, Par_Periodo)
    
        Set Fct_Obtener_Lista_de_Movimiento = .Execute(Par_NumFila)
    End With
    
    Set VarCmd = Nothing
    
End Function

Public Function Fct_Listar_Denominaciones(ByRef Par_Tipo As String, ByRef Par_EmpCodigo As String) As Recordset

    Dim VarCmd As New ADODB.Command
    
    With VarCmd
        .CommandType = adCmdStoredProc
        .CommandText = "Usp_Listar_Denominaciones"
        .ActiveConnection = gsCadenaConexion
'        .Parameters.Append .CreateParameter("@Tipo", adVarChar, adParamInput, 20, Par_Tipo)
        .Parameters.Append .CreateParameter("@Pe_EmpCodigo", adVarChar, adParamInput, 4, Par_EmpCodigo)
    
        Set Fct_Listar_Denominaciones = .Execute
    End With
    
    Set VarCmd = Nothing
    
End Function

Public Function Fct_Listar_C_Valores_Detalle(ByRef Par_Emp_cCodigo As String, _
ByRef Par_Pan_cAnio As String, ByRef Par_Per_cPeriodo As String, ByRef Par_NumFila As Integer, _
ByRef Par_Tipo As String, ByRef Par_Cuenta As String) As Recordset

    Dim VarCmd As New ADODB.Command
    
    With VarCmd
        .CommandType = adCmdStoredProc
        .CommandText = "Usp_Listar_" & Par_Tipo & "_Valores_Detalle"
        .ActiveConnection = gsCadenaConexion
        .Parameters.Append .CreateParameter("@Pe_EmpCodigo", adVarChar, adParamInput, 3, Par_Emp_cCodigo)
        .Parameters.Append .CreateParameter("@Pan_cAnio", adVarChar, adParamInput, 4, Par_Pan_cAnio)
        .Parameters.Append .CreateParameter("@Per_cPeriodo", adVarChar, adParamInput, 2, Par_Per_cPeriodo)
        .Parameters.Append .CreateParameter("@Cuenta", adVarChar, adParamInput, 15, Left(Par_Cuenta, 2))
        Set Fct_Listar_C_Valores_Detalle = .Execute(Par_NumFila)
    End With
    
    Set VarCmd = Nothing
    
End Function
    
Public Function Fct_Listar_Saldo_Anterior_LibDiario(ByRef Par_Emp_cCodigo As String, _
ByRef Par_Pan_cAnio As String, ByRef Par_Per_cPeriodo As String, ByRef Par_cCtaini As String, _
ByRef Par_cCtaFin As String) As Recordset

    Dim VarCmd As New ADODB.Command
    
    With VarCmd
        .CommandType = adCmdStoredProc
        .CommandText = "Usp_Listar_Saldo_Anterior_LibDiario"
        .ActiveConnection = gsCadenaConexion
        .Parameters.Append .CreateParameter("@Pe_EmpCodigo", adVarChar, adParamInput, 5, Par_Emp_cCodigo)
        .Parameters.Append .CreateParameter("@Pan_cAnio", adVarChar, adParamInput, 5, Par_Pan_cAnio)
        .Parameters.Append .CreateParameter("@Per_cPeriodo", adVarChar, adParamInput, 5, Par_Per_cPeriodo)
        .Parameters.Append .CreateParameter("@Pe_cCtaini", adVarChar, adParamInput, 25, Par_cCtaini)
        .Parameters.Append .CreateParameter("@Pe_cCtaFin", adVarChar, adParamInput, 25, Par_cCtaFin)
        Set Fct_Listar_Saldo_Anterior_LibDiario = .Execute
    End With
    
    Set VarCmd = Nothing
End Function
    
Public Function Sb_Grabar_Valores_Detalle(ByRef Par_Emp_cCodigo As String, _
ByRef Par_Pan_cAnio As String, ByRef Par_Per_cPeriodo As String, _
ByRef Par_Ase_nVoucher As String, ByRef Par_Asd_nItem As Integer, _
ByRef Par_Ten_cTipoEntidad As String, ByRef Par_Ent_cCodentidad As String, _
ByRef Par_Val_cTitulo As String, ByRef Par_Val_cDesTitulo As String, _
ByRef Par_Val_cBaseMedicion As String, ByRef Par_Val_cDesBaseMedicion As String, _
ByRef Par_Val_nValorNom As Double, ByRef Par_Val_nCantidad As Double, _
ByRef Par_Val_nCostoTot As Double, ByRef Par_Val_nProvTot As Double, _
ByRef Par_Val_nTotalNeto As Double, ByRef Par_Ajus_Val_Raz As Double, _
ByRef Par_Otros_Costos As Double, ByRef Par_Pe_Accion As String, ByRef Par_Cuenta As String, Optional ByRef Ase_FecAdquisicion As Date) As Boolean
On Error GoTo MIERROR
    Dim VarCmd As New ADODB.Command
    
    With VarCmd
        .CommandType = adCmdStoredProc
        .CommandText = "Usp_Grabar_Valores_Detalle"
        .ActiveConnection = gsCadenaConexion
        .Parameters.Append .CreateParameter("@Emp_Codigo", adChar, adParamInput, 3, Par_Emp_cCodigo)
        .Parameters.Append .CreateParameter("@Pan_cAnio", adChar, adParamInput, 4, Par_Pan_cAnio)
        .Parameters.Append .CreateParameter("@Per_cPeriodo", adChar, adParamInput, 2, Par_Per_cPeriodo)
        .Parameters.Append .CreateParameter("@Ase_nVoucher", adChar, adParamInput, 10, Par_Ase_nVoucher)
        .Parameters.Append .CreateParameter("@Asd_nItem", adInteger, adParamInput, , Par_Asd_nItem)
        .Parameters.Append .CreateParameter("@Ten_cTipoEntidad", adChar, adParamInput, 1, Par_Ten_cTipoEntidad)
        .Parameters.Append .CreateParameter("@Ent_cCodentidad", adChar, adParamInput, 5, Par_Ent_cCodentidad)
        .Parameters.Append .CreateParameter("@Val_cTitulo", adChar, adParamInput, 2, LTrim(RTrim(Par_Val_cTitulo)))
        .Parameters.Append .CreateParameter("@Val_cDesTitulo", adVarChar, adParamInput, 250, Par_Val_cDesTitulo)
        .Parameters.Append .CreateParameter("@Val_cBaseMedicion", adChar, adParamInput, 2, LTrim(RTrim(Par_Val_cBaseMedicion)))
        .Parameters.Append .CreateParameter("@Val_cDesBaseMedicion", adVarChar, adParamInput, 250, Par_Val_cDesBaseMedicion)
        .Parameters.Append .CreateParameter("@Val_nValorNom", adDouble, adParamInput, , Par_Val_nValorNom)
        .Parameters.Append .CreateParameter("@Val_nCantidad", adDouble, adParamInput, , Par_Val_nCantidad)
        .Parameters.Append .CreateParameter("@Val_nCostoTot", adDouble, adParamInput, , Par_Val_nCostoTot)
        .Parameters.Append .CreateParameter("@Val_nProvTot", adDouble, adParamInput, , Par_Val_nProvTot)
        .Parameters.Append .CreateParameter("@Val_nTotalNeto", adDouble, adParamInput, , Par_Val_nTotalNeto)
        .Parameters.Append .CreateParameter("@Ajus_Val_Raz", adDouble, adParamInput, , Par_Ajus_Val_Raz)
        .Parameters.Append .CreateParameter("@Otros_Costos", adDouble, adParamInput, , Par_Otros_Costos)
        .Parameters.Append .CreateParameter("@Cuenta", adVarChar, adParamInput, 15, Par_Cuenta)
        .Parameters.Append .CreateParameter("@Pe_Accion", adChar, adParamInput, 1, Par_Pe_Accion)
        .Parameters.Append .CreateParameter("@Ase_FecAdquisicion", adDate, adParamInput, 1, Ase_FecAdquisicion)
        
        sSql = "Usp_Grabar_Valores_Detalle '" & Par_Emp_cCodigo & "','" & Par_Pan_cAnio & "','" & _
        Par_Per_cPeriodo & "','" & Par_Ase_nVoucher & "'," & Par_Asd_nItem & ",'" & Par_Ten_cTipoEntidad & "','" & _
        Par_Ent_cCodentidad & "','" & LTrim(RTrim(Par_Val_cTitulo)) & "','" & Par_Val_cDesTitulo & "','" & _
        LTrim(RTrim(Par_Val_cBaseMedicion)) & "','" & Par_Val_cDesBaseMedicion & "'," & _
        Par_Val_nValorNom & "," & Par_Val_nCantidad & "," & Par_Val_nCostoTot & "," & Par_Val_nProvTot & "," & _
        Par_Val_nTotalNeto & "," & Par_Ajus_Val_Raz & "," & Par_Otros_Costos & ",'" & _
        Par_Cuenta & "','" & Par_Pe_Accion & "', '" & Ase_FecAdquisicion & "'"
        
        .Execute
    End With
    
    Set VarCmd = Nothing
    Sb_Grabar_Valores_Detalle = True
    Exit Function
MIERROR:
    MsgBox Err.Description
'    RESUME
Sb_Grabar_Valores_Detalle = False
    Set VarCmd = Nothing
End Function

Public Sub Sb_Borrar_Valores_Detalle(ByRef Par_Accion As String, ByRef Par_Emp_cCodigo As String, _
ByRef Par_Pan_cAnio As String, ByRef Par_Per_cPeriodo As String, _
ByRef Par_Ase_nVoucher As String, ByRef Par_Asd_nItem As Integer, _
ByRef Par_Ten_cTipoEntidad As String, ByRef Par_Ent_cCodentidad As String, _
ByRef Par_Val_cTitulo As String, ByRef Par_Val_cBaseMedicion As String, ByRef Par_Cuenta As String)
On Error GoTo MIERROR
    Dim VarCmd As New ADODB.Command
    
    With VarCmd
    
        .CommandType = adCmdStoredProc
        .CommandText = "Usp_Borrar_Valores_Detalle"
        .ActiveConnection = gsCadenaConexion
        .Parameters.Append .CreateParameter("@Accion", adVarChar, adParamInput, 20, Par_Accion)
        .Parameters.Append .CreateParameter("@Pe_EmpCodigo", adChar, adParamInput, 3, Par_Emp_cCodigo)
        .Parameters.Append .CreateParameter("@Pan_cAnio", adChar, adParamInput, 4, Par_Pan_cAnio)
        .Parameters.Append .CreateParameter("@Per_cPeriodo", adChar, adParamInput, 2, Par_Per_cPeriodo)
        .Parameters.Append .CreateParameter("@Ase_nVoucher", adChar, adParamInput, 10, Par_Ase_nVoucher)
        .Parameters.Append .CreateParameter("@Asd_nItem", adInteger, adParamInput, , Par_Asd_nItem)
        .Parameters.Append .CreateParameter("@Ten_cTipoEntidad", adChar, adParamInput, 1, Par_Ten_cTipoEntidad)
        .Parameters.Append .CreateParameter("@Ent_cCodentidad", adChar, adParamInput, 5, Par_Ent_cCodentidad)
        .Parameters.Append .CreateParameter("@Val_cTitulo", adChar, adParamInput, 2, Left(Par_Val_cTitulo, 2))
        .Parameters.Append .CreateParameter("@Val_cBaseMedicion", adChar, adParamInput, 2, Left(Par_Val_cBaseMedicion, 2))
        .Parameters.Append .CreateParameter("@Val_cCuenta", adChar, adParamInput, 2, Left(Par_Cuenta, 2))
        
        sSql = "Usp_Borrar_Valores_Detalle '" & Par_Accion & "','" & Par_Emp_cCodigo & "','" & Par_Pan_cAnio & "','" & _
        Par_Per_cPeriodo & "','" & Par_Ase_nVoucher & "'," & Par_Asd_nItem & ",'" & Par_Ten_cTipoEntidad & "','" & _
        Par_Ent_cCodentidad & "','" & Left(Par_Val_cTitulo, 2) & "','" & Left(Par_Val_cBaseMedicion, 2) & "', Par_Cuenta"
         
        
        .Execute
    End With
    
    Set VarCmd = Nothing
    Exit Sub
MIERROR:
    MsgBox Err.Description
'    RESUME
    Set VarCmd = Nothing
End Sub

Public Sub Sb_Limpiar_Grilla(ByRef Msh As MSHFlexGrid)
Dim c As Integer
Msh.Rows = 2
On Error GoTo Control
For c = 0 To Msh.COLS - 1
    Msh.TextMatrix(1, c) = ""
Next c
Msh.TextMatrix(1, 14) = Msh.Row
Control:
End Sub
    
    
'Sub printl(gsLinea As String)
'
'Dim liQuiebre   As Integer
'Dim lbIsPrinter As Boolean
'
'On Error GoTo ERRORIMPRESORA
'lbIsPrinter = True
'If GsDestino = "Archivo" Or GsDestino = "PrintFile" Then lbIsPrinter = False
'
'If Len(Trim$(gsLinea)) = 0 Then
'   If lbIsPrinter Then
'     If gsPagina + 1 >= Gs_DesdePag Then Printer.Print
'   Else
'     Print #1, " "
'   End If
'Else
'   If lbIsPrinter Then
'     If gsPagina + 1 >= Gs_DesdePag Then Printer.Print gsLinea
'   Else
'     Print #1, gsLinea
'   End If
'End If
'
'liQuiebre = 68
'giLineas = giLineas + 1
'
'If giLineas >= liQuiebre Then
'   giLineas = 0
'   If lbIsPrinter Then
'      If Gs_DesdePag > 1 Then
'         If gsPagina + 1 >= Gs_DesdePag Then
'         Printer.Print ""
'         Printer.NewPage
'         Printer.FontSize = 10
'         Else
'         Printer.KillDoc
'         End If
'      Else
'         Printer.Print ""
'         Printer.NewPage
'         Printer.FontSize = 10
'
'        If (frmRepAnexoInvBalance.UltCol - 1) < (frmRepAnexoInvBalance.rsArreglo.Fields.Count - 1) Then
'           If gsTipoImp <> "1" Then
'            If Not frmRepAnexoInvBalance.rsArreglo.EOF And (frmRepAnexoInvBalance.UltCol - 1) < (frmRepAnexoInvBalance.rsArreglo.Fields.Count - 1) And gsPaginaPrincipal = 1 Then
'             frmRepAnexoInvBalance.rsArreglo.MoveFirst
'            Else
'             frmRepAnexoInvBalance.rsArreglo.AbsolutePosition = frmRepAnexoInvBalance.PosReg
'            End If
'           Else
'             frmRepAnexoInvBalance.rsArreglo.MoveFirst
'           End If
'          'frmRepAnexoInvBalance.rsArreglo.MoveFirst
'           Select Case gsAccionRep
'           Case 1: Call frmRepAnexoInvBalance.CabeceraDiarioSimplificado: Call frmRepAnexoInvBalance.ImprimeDetalle
'           End Select
'          frmRepAnexoInvBalance.rsArreglo.MoveNext
'          Exit Sub
'        Else
'          frmRepAnexoInvBalance.PosReg = frmRepAnexoInvBalance.rsArreglo.AbsolutePosition + 1
'          Exit Sub
'        End If
'
'      End If
'   Else
'    If (frmRepAnexoInvBalance.UltCol - 1) < (frmRepAnexoInvBalance.rsArreglo.Fields.Count - 1) Then
'     If gsTipoImp <> "1" Then
'      If Not frmRepAnexoInvBalance.rsArreglo.EOF And (frmRepAnexoInvBalance.UltCol - 1) < (frmRepAnexoInvBalance.rsArreglo.Fields.Count - 1) And gsPaginaPrincipal = 1 Then
'       frmRepAnexoInvBalance.rsArreglo.MoveFirst
'      Else
'       frmRepAnexoInvBalance.rsArreglo.AbsolutePosition = frmRepAnexoInvBalance.PosReg
'      End If
'     Else
'       frmRepAnexoInvBalance.rsArreglo.MoveFirst
'     End If
'     'frmRepAnexoInvBalance.rsArreglo.MoveFirst
'       Select Case gsAccionRep
'       Case 1: Call frmRepAnexoInvBalance.CabeceraDiarioSimplificado: Call frmRepAnexoInvBalance.ImprimeDetalle
'       End Select
'      frmRepAnexoInvBalance.rsArreglo.MoveNext
'      Exit Sub
'    Else
'      frmRepAnexoInvBalance.PosReg = frmRepAnexoInvBalance.rsArreglo.AbsolutePosition + 1
'      Exit Sub
'    End If
'   End If
'
'End If
'
'Exit Sub
'ERRORIMPRESORA:
'    If Err.Number = 482 Then
'       MsgBox Err.Description, vbCritical, App.Title
'    End If
'End Sub
'

Public Function Execute_SQL(ADODB_RecordSet As ADODB.Recordset, sSql As String)
    Set ADODB_RecordSet = New ADODB.Recordset
    ADODB_RecordSet.Open sSql, gsCadenaConexion, adOpenStatic, adLockOptimistic
End Function

Public Function FormatDec(dNumero As Double) As Double
On Error GoTo MIERROR
    FormatDec = Format(dNumero, "###,###,##0.00")
    Exit Function
MIERROR:
    MsgBox Err.Description
End Function

Public Function ExtraeCampo(sCampo As String, sTabla As String, Optional sWhere As String) As Variant
    
On Error GoTo MIERROR
    Dim AdoRsExtrae As ADODB.Recordset
    
    sSql = "SELECT " & sCampo & " FROM " & sTabla & " WHERE " & sWhere
    Set AdoRsExtrae = New ADODB.Recordset
    Execute_SQL AdoRsExtrae, sSql
    
    If AdoRsExtrae.RecordCount > 0 Then
        ExtraeCampo = IIf(IsNull(AdoRsExtrae.Fields(0)), "", AdoRsExtrae.Fields(0))
    End If
    
    Set AdoRsExtrae = Nothing
    Exit Function
MIERROR:
    MsgBox Err.Description, 48
    Set AdoRsExtrae = Nothing

End Function

Public Function StrZero(dNumero As Double, iZeros) As String
On Error GoTo MIERROR
    Dim AdoRsZero As ADODB.Recordset
    sSql = "SELECT REPLICATE('0'," & iZeros & "-LEN(LTRIM(CAST(" & dNumero & " AS VARCHAR(" & iZeros & ")))))+LTRIM(CAST(" & dNumero & " AS VARCHAR(" & iZeros & ")))"
    Set AdoRsZero = New ADODB.Recordset
    Execute_SQL AdoRsZero, sSql
    
    If AdoRsZero.RecordCount > 0 Then
        StrZero = AdoRsZero.Fields(0)
    End If
    
    Set AdoRsZero = Nothing
    Exit Function
MIERROR:
    MsgBox Err.Description, 48
    Set AdoRsZero = Nothing
End Function

'Public Function CambiarChrEsp(ByVal Palabra As String) As String
'Dim Char As String
'Dim charesp As String
'  Char = "ÑñÁáÉéÍíÓóÚúº"
'charesp = "¥¤AaE‚I¡O¢U£" & "."
'Dim i As Integer
'For i = 1 To Len(Char)
'Palabra = Replace(Palabra, Mid(Char, i, 1), Mid(charesp, i, 1))
'Next
'CambiarChrEsp = Palabra
'End Function

'Function impresora() As String
'
'    Dim buffer As String
'    Dim ret As Integer
'
'    buffer = Space(255)
'    ret = GetProfileString("Windows", ByVal "device", "", _
'                                 buffer, Len(buffer))
'
'    If ret Then
'        impresora = UCase(Left(buffer, _
'                                   InStr(buffer, ",") - 1))
'    End If
'End Function

'Public Function ObtenerImpRed(Optional ImpDialog As String) As String
'Dim obj_Imp As WshNetwork
'Dim i As Integer
'Dim ImpPre As String
'Dim ImpWsh As String
'Dim sw As Integer
'On Error Resume Next
'
'sw = 0
'Set obj_Imp = New WshNetwork
'If Len(ImpDialog) = 0 Then
'    ImpPre = impresora()
'Else
'    ImpPre = ImpDialog
'End If
'
'For i = 1 To obj_Imp.EnumPrinterConnections.Count - 1 Step 2
'ImpWsh = obj_Imp.EnumPrinterConnections(i)
'    If ImpPre = ImpWsh Then
'        ObtenerImpRed = obj_Imp.EnumPrinterConnections(i)
'        sw = 1
'        Exit Function
'    End If
'Next
'If sw = 0 Then ObtenerImpRed = ImpPre
'
''Elimina la referencia
'    Set obj_Imp = Nothing
'End Function

'Public Function NomPC() As String
'Dim nPC As String
'Dim buffer As String
'Dim estado As Long
'Dim pc As String
'
'buffer = String$(255, " ")
'estado = GetComputerName(buffer, 255)
'
'If estado <> 0 Then
'    nPC = Left(buffer, 255)
'End If
'
'NomPC = LTrim(RTrim(nPC))
'Dim i As Integer
'pc = ""
'For i = 1 To Len(NomPC)
'pc = Replace(NomPC, Mid(pc, i, 1), "")
'Next
'NomPC = pc
'End Function

Sub Guardar_Archivo(ByVal filename As String, Optional Texto As String)
On Error GoTo ERROR
    Close
    Open filename For Output Shared As #1
    If Err > 0 Then Exit Sub
        'Guarda información del TextBox al archivo
        Print #1, Texto
        Close #1
    Exit Sub
ERROR:
    'error
    MsgBox "Número de error: " & Err.Number & vbNewLine & _
           "Descripción del error: " & Err.Description, vbCritical
End Sub

'Public Function Abrir_ArchivoBat(ByVal FileName As String) As String
'    On Error GoTo ERROR
'    Close
'    Open FileName For Input As #1
'    If Err > 0 Then Exit Function
'        ' Devuelve el valor del Bat
'        Abrir_ArchivoBat = Input(LOF(1), 1)
'    Exit Function
'ERROR:
'    'error
'    MsgBox "Número de error: " & Err.Number & vbNewLine & _
'           "Descripción del error: " & Err.Description, vbCritical
'End Function

