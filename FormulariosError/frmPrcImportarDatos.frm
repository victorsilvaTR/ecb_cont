VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmPrcImportarDatos 
   Caption         =   "Importar Datos al Sistema"
   ClientHeight    =   2250
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6210
   Icon            =   "frmPrcImportarDatos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2250
   ScaleWidth      =   6210
   Begin VB.Frame Frame2 
      Height          =   1905
      Left            =   255
      TabIndex        =   0
      Top             =   135
      Width           =   5715
      Begin VB.CommandButton cmdSeleccionar 
         Caption         =   "..."
         Height          =   345
         Left            =   2355
         TabIndex        =   1
         Top             =   300
         Width           =   495
      End
      Begin TDBText6Ctl.TDBText tdbtArchivo 
         Height          =   375
         Left            =   195
         TabIndex        =   2
         Top             =   690
         Width           =   5280
         _Version        =   65536
         _ExtentX        =   9313
         _ExtentY        =   661
         Caption         =   "frmPrcImportarDatos.frx":1982
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPrcImportarDatos.frx":19EE
         Key             =   "frmPrcImportarDatos.frx":1A0C
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   1
         ShowContextMenu =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MarginBottom    =   1
         Enabled         =   -1
         MousePointer    =   0
         Appearance      =   1
         BorderStyle     =   1
         AlignHorizontal =   0
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   ""
         FormatMode      =   1
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   0
         LengthAsByte    =   0
         Text            =   ""
         Furigana        =   0
         HighlightText   =   -1
         IMEMode         =   0
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin MSForms.CommandButton cmd_salir 
         Height          =   435
         Left            =   2880
         TabIndex        =   5
         Top             =   1215
         Width           =   1665
         Caption         =   " Salir"
         PicturePosition =   327683
         Size            =   "2937;767"
         Picture         =   "frmPrcImportarDatos.frx":1A50
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdImportarDatos 
         Height          =   435
         Left            =   1035
         TabIndex        =   4
         Top             =   1215
         Width           =   1665
         Caption         =   " Importar Datos"
         PicturePosition =   327683
         Size            =   "2937;767"
         Picture         =   "frmPrcImportarDatos.frx":1FEA
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "SELECCIONE ARCHIVO:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   195
         TabIndex        =   3
         Top             =   375
         Width           =   2055
      End
   End
   Begin MSComDlg.CommonDialog dlgAbrirArchivo 
      Left            =   5715
      Top             =   1275
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmPrcImportarDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_salir_Click()
Unload Me
End Sub

'Option Explicit
'
'Private Sub cmdImportarDatos_Click()
'    Dim appWorld As Excel.Application
'    Dim wbWorld As Excel.Workbook
'    If Trim(tdbtArchivo) = "" Then
'        Mensajes "Seleccione un archivo", vbInformation
'        Exit Sub
'    End If
'    On Error Resume Next 'Ignorar errores
'
'    ' *** Abrir la Hoja Excel
'    Set appWorld = GetObject(, "Excel.Application") 'buscar una copia de Excel en ejecución
'    If Err.Number <> 0 Then 'Si no se ejecuta Excel
'        Set appWorld = CreateObject("Excel.Application") 'ejcutarlo
'    End If
'    Err.Clear   ' Borrar el objeto Err si se produce un error.
'
'    On Error GoTo 0 'Reaunudar el procesamiento normal de errores
'    Set wbWorld = appWorld.Workbooks.Open(tdbtArchivo.Text)
'    ' *** Fin de Abrir la Hoja Excel
'
'    Dim shtContinent As Excel.Worksheet
'    Dim rngFeatureList As Excel.Range
'    Dim rngFeatureCol As Excel.Range
'    Dim intFirstBlankCell As Integer
'    Dim intFirstBlankCol As Integer
'
'    ' *** Obtiene la hoja cuyo nombre es el del continente seleccionado en el cuadro combinado Continentes.
'    Set shtContinent = wbWorld.Sheets("Asientos")
'    ' *** Asigna la primera fila de esta hoja a un objeto.
'    Set rngFeatureList = shtContinent.Rows(1)
'    Set rngFeatureCol = shtContinent.Columns(1)
'    ' *** Comprueba si es lista vacía. Busca la primera celda en blanco de la fila y columna
'    If (rngFeatureList.Cells(1, 1) = "") Then
'        intFirstBlankCell = 0
'        Mensajes "No hay Datos para importar", vbInformation
'        GoTo cierraDatos
'    Else
'        intFirstBlankCell = rngFeatureList.Find("").Column
'        intFirstBlankCol = rngFeatureCol.Find("").Row
'    End If
'
'    ' *** CREANDO LOS ASIENTOS CONTABLES
'    '
'    '
'    ' *** Declarando variables para los asientos
'    Dim i As Integer
'    Dim j As Integer
'    Dim lInterno As String      ' *** Numero interno
'    Dim lVoucher As String      ' *** Num Voucher
'    Dim numAsi As String        ' *** Asiento
'    Dim numAsiAux As String     ' *** Asiento Aux
'    Dim numitm As Integer       ' *** Item
'    Dim montoSoles As Double    ' *** Monto Soles
'    Dim montoDolar As Double    ' *** Monto Dolar
'    Dim codEntidad As String    ' *** Para Hallar el codigo de la entidad
'    Dim lArrCab() As Variant    ' *** Variable para la cabecera
'    Dim xlibro As String
'    Dim xperiodo As String
'    Dim clsMante As clsMantoTablas
'    Set clsMante = New clsMantoTablas
'
'    ' *** Lee los datos de una fila
'    Screen.MousePointer = 11
'    numAsiAux = ""
'    With rngFeatureList
'    For j = 2 To intFirstBlankCol - 1
'        ' *** Generar los asientos Contables
'        numAsi = .Cells(j, 4)
'        If numAsi <> numAsiAux Then
'            ' *** Crea Destino, Actualiza saldos y cierra conexion en caso no sea 2
'            If j > 2 Then GoSub DestinoDistribucion
'            ' *** Validaciones de cabecera
'            If validarCabecera(j, .Cells(j, 1), .Cells(j, 2), .Cells(j, 6), .Cells(j, 3), .Cells(j, 7), .Cells(j, 8)) = False Then GoTo cierraDatos
'            ' *** Generando la cabecera
'            Set clsMante = New clsMantoTablas
'            lVoucher = numeroVoucher("VOUCHER", .Cells(j, 1), Format(.Cells(j, 2), "00"), Format(.Cells(j, 3), "00"))
'            lInterno = numeroVoucher("INTERNO", .Cells(j, 1), "", "")
'            xlibro = Format(.Cells(j, 3), "00")
'            xperiodo = Format(.Cells(j, 2), "00")
'            GoSub CargaArregloCab
'            If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaAsientoCab", lArrCab(), False) = False Then
'                Mensajes "El proceso no ha concluido. Verificar...", vbInformation
'                Screen.MousePointer = 0
'                Exit Sub
'            End If
'            numAsiAux = .Cells(j, 4)
'            numitm = 0
'        End If
'
'        ' *** Validando el detalle
'        If validarDetalle(j, .Cells(j, 10), .Cells(j, 8), .Cells(j, 9), .Cells(j, 14)) = False Then GoTo cierraDatos
'        ' *** Generando el detalle
'        numitm = numitm + 1
'        GoSub CargaArregloDet
'        If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaAsientoDet", lArrCab(), False) = False Then
'            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
'            Screen.MousePointer = 0
'            Exit Sub
'        End If
'        ' ***
'    Next
'    End With
'
'    ' *** Crea Destino, Actualiza saldos y cierra conexion en caso no sea 2
'    GoSub DestinoDistribucion
'
'    Mensajes "Los asientos se generaron con exito", vbInformation
'    cmdImportarDatos.Enabled = False
'    Screen.MousePointer = 0
'
'    ' *** Actualiza todas las operaciones
'    'clsMante.CommitTrans
'    'clsMante.FinalizaClase
'    'Call cambiarRutaArchivo
'
'    '
'    ' *** FIN CREANDO LOS ASIENTO CONTABLES
'cierraDatos:
'    ' *** Limpiar los datos
'    Set shtContinent = Nothing
'    Set rngFeatureList = Nothing
'    Set rngFeatureCol = Nothing
'    Set appWorld = Nothing
'    Set wbWorld = Nothing
'    Screen.MousePointer = 0
'    ' *** finaliza la clase
'    'clsMante.FinalizaClase
'    Exit Sub
'
'DestinoDistribucion:
'    ' *** Crear asientos con destino
'    GoSub CargaArregloDest
'    If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaAsientoDestino", lArrCab(), False) = False Then
'        Mensajes "El proceso no ha concluido. Verificar...", vbInformation
'        Exit Sub
'    End If
'    ' *** Actualizar los Saldos
'    GoSub CargaArregloSaldos
'    If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_ActualizaSaldos", lArrCab(), True) = False Then
'        Mensajes "El proceso no ha concluido. Verificar...", vbInformation
'        Exit Sub
'    End If
'    ' ***
'    Return
'
'CargaArregloCab:
'    ' *** Cargar los datos a grabar en un arreglo
'    ReDim lArrCab(12) As Variant
'    lArrCab(0) = "INSERTAR"     ' Accion
'    lArrCab(1) = lInterno       ' Numero Interno
'    lArrCab(2) = gsEmpresa      ' Empresa
'    lArrCab(3) = gsAnio         ' Año
'    lArrCab(4) = xperiodo       ' Periodo
'    lArrCab(5) = xlibro         ' Libro
'    lArrCab(6) = lVoucher       ' Voucher
'    lArrCab(7) = rngFeatureList.Cells(j, 6)     ' Fecha
'    lArrCab(8) = rngFeatureList.Cells(j, 7)     ' Moneda
'    lArrCab(9) = rngFeatureList.Cells(j, 8)     ' TipoCambio
'    lArrCab(10) = ""        ' Observaciones
'    lArrCab(11) = gsUsuario
'    lArrCab(12) = "A"           ' @Asd_cEstado
'    Return
'
'CargaArregloDet:
'    ' *** Cargar los datos a grabar en un arreglo
'    ReDim lArrCab(36) As Variant
'    lArrCab(0) = "INSERTAR"     ' Accion
'    lArrCab(1) = lInterno       ' Numero Interno
'    lArrCab(2) = gsEmpresa      ' Empresa
'    lArrCab(3) = gsAnio         ' Año
'    lArrCab(4) = xperiodo       ' Periodo
'    lArrCab(5) = xlibro         ' Libro
'    lArrCab(6) = lVoucher       ' Voucher
'    lArrCab(7) = numitm         ' item
'    With rngFeatureList
'        lArrCab(8) = .Cells(j, 10)      ' Plan Cuenta
'        lArrCab(9) = .Cells(j, 11)      ' Glosa
'
'        ' *** Verificar dependiendo del Tipo de Documento
'        If monedaNacional(rngFeatureList.Cells(j, 7)) = True Then
'            montoSoles = .Cells(j, 12)
'            montoDolar = Redondear(.Cells(j, 12) / .Cells(j, 8), 2)
'        Else
'            montoSoles = Redondear(.Cells(j, 12) * .Cells(j, 8), 2)
'            montoDolar = .Cells(j, 12)
'        End If
'
'        If .Cells(j, 9) = "D" Then
'            lArrCab(10) = montoSoles    ' DebeSoles
'            lArrCab(11) = montoDolar    ' DebeMonExt
'            lArrCab(12) = 0             ' HaberSoles
'            lArrCab(13) = 0             ' HaberMonExt
'        Else
'            lArrCab(10) = 0             ' DebeSoles
'            lArrCab(11) = 0             ' DebeMonExt
'            lArrCab(12) = montoSoles    ' HaberSoles
'            lArrCab(13) = montoDolar    ' HaberMonExt
'        End If
'        ' *** Fin Verificar dependiendo del Tipo de Documento
'
'        lArrCab(14) = .Cells(j, 8)      ' Tipo de Cambio
'        lArrCab(15) = .Cells(j, 13)     ' CCosto
'        lArrCab(16) = .Cells(j, 14)     ' Tipo Entidad
'
'        ' *** En todo caso verificar por el Ruc .Cells(j, 16) y Rz .Cells(j, 17)
'        If Not (Trim(.Cells(j, 14)) = "" And Trim(.Cells(j, 16)) = "") Then
'            ' *** Hallando la entidad
'            codEntidad = CodigoEntidadRuc(.Cells(j, 16), .Cells(j, 14))
'            If codEntidad = "" Then
'                ' *** Crear la entidad con los datos obtenidos
'                codEntidad = correlativoCodigoEnt(.Cells(j, 14))
'                Call crearEntidad(codEntidad, .Cells(j, 14), .Cells(j, 16), .Cells(j, 17))
'            End If
'            ' ***
'        Else
'            codEntidad = ""
'        End If
'
'        lArrCab(17) = codEntidad        ' Codigo Entidad
'        ' ***
'
'        lArrCab(18) = .Cells(j, 18)     ' Tipo Doc
'        lArrCab(19) = .Cells(j, 19)     ' Serie
'        lArrCab(20) = .Cells(j, 20)     ' Numero Doc
'        lArrCab(21) = .Cells(j, 21)     ' Fecha Doc
'
'        lArrCab(22) = .Cells(j, 25)     ' TipoDoc Ref
'        lArrCab(23) = .Cells(j, 26)     ' SerieDoc Ref
'        lArrCab(24) = .Cells(j, 27)     ' NumeroDoc Ref
'        lArrCab(25) = .Cells(j, 28)     ' FechaDoc Ref
'        lArrCab(26) = 0     ' MontoInafecto
'        lArrCab(27) = ""    ' Retencion
'        lArrCab(28) = .Cells(j, 23)     ' Fecha Spot
'        lArrCab(29) = .Cells(j, 22)     ' NumSpot
'        lArrCab(30) = "0"    ' Destino
'
'        lArrCab(31) = 0     ' Correlativo Provision
'        lArrCab(32) = gsUsuario         ' Usuario
'        lArrCab(33) = "A"   ' Estado
'        lArrCab(34) = .Cells(j, 5)      ' Indica Si es Provision/Cancelacion o Ninguno
'        lArrCab(35) = .Cells(j, 24)     ' TipoCompra (A B C)
'        lArrCab(36) = ""                ' TipoCompra (A B C)
'    End With
'    Return
'
'CargaArregloDest:
'    ReDim lArrCab(6) As Variant
'    lArrCab(0) = lInterno       ' Numero Interno
'    lArrCab(1) = gsEmpresa      ' Empresa
'    lArrCab(2) = gsAnio         ' Año
'    lArrCab(3) = xperiodo       ' Periodo
'    lArrCab(4) = xlibro         ' Libro
'    lArrCab(5) = lVoucher       ' Voucher
'    lArrCab(6) = numitm + 1     ' asdItem
'    Return
'
'CargaArregloSaldos:
'    ReDim lArrCab(7) As Variant
'    lArrCab(0) = "ACTUALIZAR"   ' Accion
'    lArrCab(1) = lInterno       ' Numero Interno
'    lArrCab(2) = gsEmpresa      ' Empresa
'    lArrCab(3) = gsAnio         ' Año
'    lArrCab(4) = xperiodo       ' Periodo
'    lArrCab(5) = xlibro         ' Libro
'    lArrCab(6) = lVoucher       ' Voucher
'    lArrCab(7) = gsUsuario
'    Return
'
'End Sub
'
'Private Function validarDetalle(fila As Integer, cuenta As String, tc As Double, tipomov As String, tipoEnt As String) As Boolean
'    Dim valorDato As String
'    validarDetalle = False
'    If tc = 0 Then
'        Mensajes "El Tipo de Cambio no puede ser igual a 0. Fila: " & fila & ".  ", vbInformation
'        Exit Function
'    End If
'    If Not (tipomov = "H" Or tipomov = "D") Then
'        Mensajes "El Tipo de Mov debe ser 'D' o 'H' . Fila: " & fila & ".  ", vbInformation
'        Exit Function
'    End If
'    valorDato = ExisteCtaNoTitulo(cuenta, "N")
'    If valorDato = "" Then
'        Mensajes "Cuenta de la fila: " & fila & " no existe o es cuenta no titulo. ", vbInformation
'        Exit Function
'    End If
'    ' ***
'    If Trim(tipoEnt) <> "" Then
'        If ExisteRegistro(tipoEnt, "spCn_GrabaTipoEntidad", False) = False Then
'            Mensajes "El Tipo de Entidad de la fila: " & fila & " no existe. ", vbInformation
'            Exit Function
'        End If
'    End If
'    validarDetalle = True
'End Function
'
'Private Function validarCabecera(fila As Integer, anio As String, mes As String, fecha As String, libro As String, moneda As String, tc As Double) As Boolean
'    validarCabecera = False
'    If anio <> gsAnio Then
'        Mensajes "Dato de la fila: " & fila & " no pertenece al Año de trabajo Actual. ", vbInformation
'        Exit Function
'    End If
'    If CierreMes(Format(mes, "00")) = True Then
'        Mensajes "La mes de la fila: " & fila & " ha sido cerrado. ", vbInformation
'        Exit Function
'    End If
'    If Format(mes, "00") <> Format(Month(fecha), "00") Or Year(fecha) <> gsAnio Then
'        Mensajes "La fecha de la fila: " & fila & " no coincide con el Año o periodo del mismo registro. ", vbInformation
'        Exit Function
'    End If
'    If ExisteRegistro(libro, "spCn_GrabaLibroOpera", True) = False Then
'        Mensajes "El Codigo de libro de la fila: " & fila & " no existe. ", vbInformation
'        Exit Function
'    End If
'    If ExisteRegistro(moneda, "spCn_GrabaTipoMoneda", False) = False Then
'        Mensajes "El Tipo de Moneda de la fila: " & fila & " no existe. ", vbInformation
'        Exit Function
'    End If
'    If tc = 0 Then
'        Mensajes "El Tipo de Cambio no puede ser igual a 0. Fila: " & fila & ".  ", vbInformation
'        Exit Function
'    End If
'    validarCabecera = True
'End Function
'
'Private Sub crearEntidad(codigo As String, tipo As String, ruc As String, razonSoc As String)
'    Dim clsEntidad As New clsMantoTablas
'    Dim lArrEnt(12) As Variant
'    lArrEnt(0) = "INSERTAR"     ' Accion
'    lArrEnt(1) = gsEmpresa      ' Empresa
'    lArrEnt(2) = codigo         ' Codigo
'    lArrEnt(3) = tipo           ' Codigo
'    lArrEnt(4) = razonSoc       ' Nombre o Razon Social
'    lArrEnt(5) = ""             ' Direccion
'    lArrEnt(6) = ruc            ' Numero de Documento
'    lArrEnt(7) = ""             ' Representante
'    lArrEnt(8) = "04"           ' Tipo de Documento
'    lArrEnt(9) = "0"            ' *** Tipo de Persona
'    lArrEnt(10) = "1"           ' *** Estado Sunat
'    lArrEnt(11) = "A"           ' *** Estado de Entidad
'    lArrEnt(12) = gsUsuario     ' Usuario
'    If clsEntidad.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaEntidad", lArrEnt(), True) = False Then
'        Mensajes "El proceso no ha concluido. Verificar...", vbInformation
'        Exit Sub
'    End If
'End Sub
'
'Private Sub cambiarRutaArchivo()
'    Dim fso As Object
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    fso.MoveFile Me.tdbtArchivo.Text, "C:\"
'    Set fso = Nothing
'    ' ***
'End Sub
'
'Private Sub cmdSeleccionar_Click()
'    Me.tdbtArchivo = ""
'    On Local Error GoTo ErrorEjecucion
'    With Me.dlgAbrirArchivo
'        .DialogTitle = "Archivo de Datos de Asientos"
'        .InitDir = "C:"
'        .Filter = "Archivos de Datos(*.xls;*.dbf)| *.xls;*.dbf"
'        .CancelError = True
'        .ShowOpen
'        If .FileName = "" Then
'            Mensajes "Selecciones un archivo", vbInformation
'        Else
'            tdbtArchivo = .FileName
'        End If
'    End With
'
'    Exit Sub
'ErrorEjecucion:
'    If Err.Number <> 32755 Then Mensajes Str(Err.Number) & Err.Description, vbInformation
'End Sub
'
'Private Sub Form_Load()
'
'End Sub
Private Sub cmdImportarDatos_Click()

End Sub
