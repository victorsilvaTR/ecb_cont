VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFechaVencimiento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Tipo de Cambio"
   ClientHeight    =   5985
   ClientLeft      =   855
   ClientTop       =   2970
   ClientWidth     =   7965
   Icon            =   "frmFechaVencimiento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5985
   ScaleWidth      =   7965
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   80
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFechaVencimiento.frx":0ECA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFechaVencimiento.frx":12A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFechaVencimiento.frx":167E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFechaVencimiento.frx":1A58
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFechaVencimiento.frx":1E32
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFechaVencimiento.frx":220C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFechaVencimiento.frx":25E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFechaVencimiento.frx":29C0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imglstdisabled 
      Left            =   11025
      Top             =   2475
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFechaVencimiento.frx":39DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFechaVencimiento.frx":3B34
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFechaVencimiento.frx":3C8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFechaVencimiento.frx":3DE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFechaVencimiento.frx":3F42
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFechaVencimiento.frx":409C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFechaVencimiento.frx":41F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFechaVencimiento.frx":4350
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFechaVencimiento.frx":44AA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imglstTool 
      Left            =   11025
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFechaVencimiento.frx":4604
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFechaVencimiento.frx":4B9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFechaVencimiento.frx":5138
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFechaVencimiento.frx":56D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFechaVencimiento.frx":5C6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFechaVencimiento.frx":6206
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFechaVencimiento.frx":67A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFechaVencimiento.frx":6D3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFechaVencimiento.frx":72D4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrOpciones 
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "imglstTool"
      DisabledImageList=   "imglstdisabled"
      HotImageList    =   "imglstTool"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Grabar F4"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cancelar o Salir ESC"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   4860
      Left            =   300
      TabIndex        =   0
      Top             =   450
      Width           =   7395
      Begin TDBDate6Ctl.TDBDate dtpFecha 
         Height          =   300
         Left            =   4875
         TabIndex        =   1
         Top             =   1095
         Width           =   1545
         _Version        =   65536
         _ExtentX        =   2725
         _ExtentY        =   529
         Calendar        =   "frmFechaVencimiento.frx":786E
         Caption         =   "frmFechaVencimiento.frx":7970
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFechaVencimiento.frx":79D4
         Keys            =   "frmFechaVencimiento.frx":79F2
         Spin            =   "frmFechaVencimiento.frx":7A5E
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "dd/mm/yyyy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   1
         ForeColor       =   -2147483640
         Format          =   "dd/mm/yyyy"
         HighlightText   =   0
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxDate         =   73415
         MinDate         =   2
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "03/08/2004"
         ValidateMode    =   0
         ValueVT         =   2010185735
         Value           =   38202
         CenturyMode     =   0
      End
   End
End
Attribute VB_Name = "frmFechaVencimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim lArrMnt() As Variant        ' *** Arreglo para los mantenimientos
'Dim lTipoMnt As String          ' *** Tipo de Mantenimiento (insert/update/delete)
'Dim lrsTabla(1) As Recordset
'Dim lRegElim As Boolean         ' *** Verifica si registro ha sido eliminado desde otra sesion
'Dim gsPeriodoAnterior(1) As String
'Dim IndiceTab As Integer
'Dim IndiceTabAnt As Integer
'Public Voucher As Boolean
'Public RegAux  As Boolean
'Public MonAdic As Boolean
'Public ColumnaTC As Integer
'Public Asientos As Boolean
'
'
'Dim gsGrupo As String
'Public Property Let Grupo(ByVal Grupo As String)
'     gsGrupo = Grupo
'End Property
'
'Private Sub chkFecha_Click(Index As Integer)
'    Call FiltrarRecordSet(Index)
'End Sub
'
'Private Sub chkFecha_KeyPress(Index As Integer, KeyAscii As Integer)
'If KeyAscii = 13 Then
'    'pSetFocus dtpFechaBus
'End If
'
'End Sub
''
''
''
''Private Sub dtpFechaBus_Change(Index As Integer)
''    If chkFecha(Index).Value = 1 Then Call FiltrarRecordSet(Index)
''End Sub
'
'Private Sub Form_Resize()
'On Error GoTo errHand
'    If Me.WindowState <> vbMinimized Then
'
'        '*** REDIMENSIONAR SST
'        With SSTCentroCosto
'            .Width = Me.Width - .Left + 15 - 200
'            .Height = Me.Height - .Top + 15 - 500
'            '*** REDIMENSIONAR FRAME PRINCIPAL
'            Frame1.Width = .Width - IIf(.TabOrientation = ssTabOrientationLeft, .TabHeight, 0) - 500
'            Frame1.Height = .Height - IIf(.TabOrientation = ssTabOrientationTop, .TabHeight, 0) - 500
'        End With
'
'        With tdbgMoneda(0)
'            '*** REDIMENSIONAR CUADRICULA DE LISTADO
'            .Width = Frame1.Width - .Left - 500
'            .Height = Frame1.Height - .Top - 200
'        End With
'
'        With tdbgMoneda(1)
'            '*** REDIMENSIONAR CUADRICULA DE LISTADO
'            .Width = Frame1.Width - .Left - 500
'            .Height = Frame1.Height - .Top - 200
'        End With
'
'        '*** REDIMENSIONAR DETALLE
'        Frame4.Height = Frame1.Height
'        Frame4.Width = Frame1.Width
'
'        Frame2(0).Height = Frame1.Height
'        Frame2(0).Width = Frame1.Width
'
'        Frame2(1).Height = Frame1.Height
'        Frame2(1).Width = Frame1.Width
'
'        tbrOpciones.Width = Me.Width
'    End If
'Exit Sub
'errHand:
'End Sub
'
'Private Sub SSTCentroCosto_Click(PreviousTab As Integer)
'    If PreviousTab < 2 Then IndiceTabAnt = PreviousTab
'    IndiceTab = SSTCentroCosto.Tab
'
'    If SSTCentroCosto.Tab = 2 Or SSTCentroCosto.Tab = 4 Then
'        tbrOpciones.Buttons(1).Enabled = False
'        tbrOpciones.Buttons(3).Enabled = False
'        tbrOpciones.Buttons(4).Enabled = False
'        tbrOpciones.Buttons(5).Enabled = False
'
'    ElseIf SSTCentroCosto.Tab = 0 Or SSTCentroCosto.Tab = 1 Then
'        tbrOpciones.Buttons(1).Enabled = True
'        tbrOpciones.Buttons(3).Enabled = True
'        tbrOpciones.Buttons(4).Enabled = True
'        tbrOpciones.Buttons(5).Enabled = True
'
'        SeteaBarraHerramientas Me.tbrOpciones, gsGrupo
'    End If
'
'    If SSTCentroCosto.Tab = 2 Then
'        Call CargaMeses
'    End If
'End Sub
'
'Private Sub tbrOpciones_ButtonClick(ByVal Button As MSComctlLib.Button)
'    If IndiceTab < 2 Then IndiceTabAnt = IndiceTab
'
'    IndiceTab = SSTCentroCosto.Tab
'    Dim respuesta As String
'    Select Case Button.Index
'        Case 1: ManNuevo
'        Case 2: VerDatos
'        Case 3: Grabar
'                SeteaBarraHerramientas Me.tbrOpciones, gsGrupo
'
'        Case 4: Borrar (IndiceTab)
'        Case 5: Editar
'        Case 6: Imprimir
'        Case 7
'            If SSTCentroCosto.TabEnabled(3) = False Then ' *** Grabar
''                respuesta = MsgBox("Esta seguro que desea salir del formulario actual", vbYesNo+ vbQuestion , "Confirmar Salir")
''                If respuesta = vbYes Then Unload Me
'                Unload Me
'            Else
'                respuesta = MsgBox("Desea cancelar la siguiente operación", vbYesNo + vbQuestion, "Confirmar Cancelar")
'                If respuesta = vbYes Then
'                    Call Cancelar
'                    SeteaBarraHerramientas Me.tbrOpciones, gsGrupo
'                End If
'            End If
'    End Select
'End Sub
'
'Private Sub Borrar(Indice As Integer)
'    If SSTCentroCosto.Tab = 2 Then
'        Exit Sub
'    End If
'
'    ' *** Eliminar los datos; segun el q esta seleccionado
'    Dim respuesta As String
'    If Trim(tdbgMoneda(Indice).Columns(0).Value) <> "" Then
'        respuesta = MsgBox("Desea eliminar definitivamente el registro Seleccionado", vbYesNo + vbQuestion + vbDefaultButton2, "Confirmar Eliminar Registro")
'        If respuesta = vbYes Then
'            Dim clsMante As clsMantoTablas
'            Set clsMante = New clsMantoTablas
'            ' *** Eliminando la Cuenta
'            Screen.MousePointer = vbHourglass
'            Call CargaArregloMnt
'            lArrMnt(0) = "ELIMINAR"                     ' Accion
'            lArrMnt(2) = tdbgMoneda(Indice).Columns(0).Value  ' Fecha
'            lArrMnt(3) = gsMonedaNac     ' MonedaOrigen
'            lArrMnt(4) = tdbgMoneda(Indice).Columns(1).Value
'            If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaTipoCambio", lArrMnt(), True) = False Then
'                Mensajes "El proceso no se ha realizado. Verificar...", vbInformation
'                Screen.MousePointer = vbDefault
'                Exit Sub
'            End If
'            Call CargaTablaMoneda(IndiceTabAnt)
'            Screen.MousePointer = vbDefault
'            FiltrarRecordSet (IndiceTabAnt)
'            Mensajes "Registro ha sido eliminado", vbInformation
'        End If
'    Else
'        Mensajes "Debe Seleccionar el registro a eliminar...", vbInformation
'    End If
'    ' ***
'End Sub
'
'Private Sub VerDatos()
'    lblMante = "VER REGISTRO"
'    Call CargaDatosRegistro(IndiceTab)
'    SSTCentroCosto.TabEnabled(3) = True
'    SSTCentroCosto.TabEnabled(0) = False
'    SSTCentroCosto.Tab = 3
'    tbrOpciones.Buttons(4).Enabled = False  ' *** Borrar
'    tbrOpciones.Buttons(7).Image = 8
'    lTipoMnt = "EDITAR"
'    Call AseguraControl(Me, True)
'End Sub
'
'Private Sub Editar()
'    If SSTCentroCosto.Tab = 2 Then
'        Exit Sub
'    End If
'
'    Call CargaDatosRegistro(IndiceTab)
'    If lRegElim = False Then
'        lTipoMnt = "EDITAR"
'        If Me.lblMante = "VER REGISTRO" Then Call AseguraControl(Me, False)
'        Call HabilitaControl(Me)
'        lblMante = "MODIFICANDO REGISTRO"
'        Call TabMantenimiento(True)
'        dtpFecha.ReadOnly = True
'        'tdbcA.Locked = True
'        tdbcDe(0).Locked = True
'        pSetFocus tdbnCompra
'
'    Else
'        lRegElim = False
'    End If
'End Sub
'
'Public Sub ManNuevo()
'    On Error GoTo ERROR
'    lTipoMnt = "INSERTAR"
'    Call LimpiaTexto(Me)
'    Call HabilitaControl(Me)
'    ' ***
'    lblMante = "NUEVO REGISTRO"
'    Call TabMantenimiento(True)
'    dtpFecha.ReadOnly = False
'    'tdbcA.Locked = False
'    tdbcDe(0).Locked = False
'    dtpFecha = FechaServidor
'    pSetFocus dtpFecha
'
'    If IndiceTabAnt = 0 Then
'        tdbcDe(0).BoundText = gsMonedaExt
'        tdbcDe(0).Locked = True
'    Else
'        tdbcDe(0).BoundText = tdbcMonedaAdic.BoundText
'        tdbcDe(0).Locked = False
'    End If
'    If Me.tdbgMoneda(IndiceTabAnt).Row > 0 Then Me.tdbgMoneda(IndiceTabAnt).Row = 0
'    If CE(Me.tdbgMoneda(IndiceTabAnt).Columns(0)) <> "" Then
'        dtpFecha.Value = DateAdd("d", 1, Me.tdbgMoneda(IndiceTabAnt).Columns(0))
'        Me.tdbcMes(IndiceTabAnt).BoundText = Right("00" & Month(dtpFecha), 2)
'    Else
'        dtpFecha.Value = "01/" + tdbcMes(IndiceTabAnt).BoundText + "/" + gsAnio
'
'    End If
'
'    tdbnCompra.Enabled = True
'    tdbnVenta.Enabled = True
'    tdbnVentaP.Enabled = True
'    pSetFocus tdbnCompra
'    Exit Sub
'ERROR:
'
'End Sub
'
'Private Sub BuscarMonedas(Indice As Integer)
'    On Error GoTo ERROR
'    ' *** Busca la moneda nacional y extranjera por defecto
'    Dim i As Integer
'    Dim Cont As Integer
'    Cont = 0
'    For i = 0 To tdbcDe(Indice).ListCount - 1
'        tdbcDe(Indice).Row = i
'        If tdbcDe(Indice).Columns(2) = "1" Then
'            tdbcDe(Indice).Bookmark = i
'            Cont = Cont + 1
'        End If
'        If tdbcDe(Indice).Columns(3) = "1" Then
'            'tdbcA.Bookmark = i
'            Cont = Cont + 1
'        End If
'        If Cont = 2 Then Exit For
'    Next
'    Exit Sub
'ERROR:
'
'End Sub
'
'Private Sub TabMantenimiento(Valor As Boolean)
'    SSTCentroCosto.TabEnabled(3) = Valor
'    SSTCentroCosto.TabEnabled(0) = Not Valor
'    SSTCentroCosto.TabEnabled(1) = Not Valor
'    SSTCentroCosto.TabEnabled(2) = Not Valor
'    'SSTCentroCosto.TabEnabled(4) = Not Valor
'
'    If Valor = True Then SSTCentroCosto.Tab = 3
'    If Valor = False Then SSTCentroCosto.Tab = 0
'    tbrOpciones.Buttons(1).Enabled = Not Valor  ' *** Nuevo
'    tbrOpciones.Buttons(2).Enabled = Not Valor  ' *** Buscar
'    tbrOpciones.Buttons(3).Enabled = Valor      ' *** Grabar
'    tbrOpciones.Buttons(4).Enabled = Not Valor  ' *** Borrar
'    tbrOpciones.Buttons(5).Enabled = Not Valor  ' *** Editar
'    If Valor = True Then
'        tbrOpciones.Buttons(7).Image = 8
'    Else
'        tbrOpciones.Buttons(7).Image = 7
'    End If
'End Sub
'
'Private Sub Cancelar()
'    If Me.lblMante = "VER REGISTRO" Then
'        Call AseguraControl(Me, False)
'    Else
'        Call HabilitaControl(Me)
'    End If
'    Call TabMantenimiento(False)
'    SSTCentroCosto.Tab = IndiceTabAnt
'    pSetFocus tdbgMoneda(IndiceTabAnt)
'End Sub
'
'Private Function BuscaFecha() As Date
'    Select Case SSTCentroCosto.Tab
'           Case 0
'                BuscaFecha = dtpFechaBus(0).Value
'           Case 1
'                BuscaFecha = dtpFechaBus(0).Value
'    End Select
'End Function
'
'Private Function BuscaMoneda() As String
'    Select Case SSTCentroCosto.Tab
'           Case 0
'                BuscaMoneda = gsMonedaExt
'           Case 1
'                BuscaMoneda = tdbcMonedaAdic.BoundText
'    End Select
'End Function
'
'Private Sub Imprimir()
'    Dim matriz(15) As Variant
'    Dim Titulo As String
'
'
'    If SSTCentroCosto.Tab = 0 Then
'        Titulo = "Tipo de Cambio - " & gsNombreMonedaExt
'
'    ElseIf SSTCentroCosto.Tab = 1 Then
'        If tdbcMonedaAdic.BoundText = "" Then
'            Mensajes "Seleccione una moneda Adicional", vbOKOnly + vbInformation
'            Exit Sub
'        End If
'        Titulo = "Tipo de Cambio - " & tdbcMonedaAdic.Text
'
'    ElseIf SSTCentroCosto.Tab = 2 Then
'        Titulo = "Tipo de Cambio Mensual " & gsNombreMonedaExt & " " & gsAnio
'    End If
'
'
'    Titulo = UCase(Titulo)
'    matriz(0) = "@NomEmp;" & gsEmpresaNom & ";True"
'    matriz(1) = "@Titulo00;" & Titulo & ";True"
'    matriz(2) = "@Titulo01;;True"
'    matriz(3) = "@Titulo02;;True"
'    matriz(4) = "@Titulo03;FECHA;True"
'    matriz(5) = "@Titulo04;COMPRA VIG.;True"
'    matriz(6) = "@Titulo05;VENTA VIG.;True"
'    matriz(7) = "@Titulo06;VENTA PUBL.;True"
'    matriz(8) = "@Titulo07;;True"
'
'    matriz(14) = "@EMPRESA;" & gsEmpresaNom & ";True"
'    matriz(15) = "@RUC;" & "RUC : " & gsRUC & ";True"
'
'    If SSTCentroCosto.Tab <= 1 Then
'        matriz(9) = "@Tipo;TIPO_CAMBIO;True"
'    Else
'        matriz(9) = "@Tipo;TIPO_CAMBIO_MENSUAL;True"
'    End If
'
'    matriz(10) = "@Emp_cCodigo;" & gsEmpresa & ";True"
'    matriz(11) = "@Pan_cAnio;" & gsAnio & ";True"
'
'    Dim formulas(0) As Variant
'
'    If SSTCentroCosto.Tab = 0 Then
'        matriz(12) = "@Per_cPeriodo;" & tdbcMes(0).BoundText & ";True"
'        matriz(13) = "@Aux;" & gsMonedaExt & ";True"
'        AbreReporteParam gsDSN, Me, rutaReportes & "RptEstandar.rpt", crptToWindow, Titulo, "", matriz(), formulas()
'
'    ElseIf SSTCentroCosto.Tab = 1 Then
'        matriz(12) = "@Per_cPeriodo;" & tdbcMes(1).BoundText & ";True"
'        matriz(13) = "@Aux;" & tdbcMonedaAdic.BoundText & ";True"
'        AbreReporteParam gsDSN, Me, rutaReportes & "RptEstandar.rpt", crptToWindow, Titulo, "", matriz(), formulas()
'
'    ElseIf SSTCentroCosto.Tab = 2 Then
'        matriz(12) = "@Per_cPeriodo;;True"
'        matriz(13) = "@Aux;" & tdbcTipoMensual.BoundText & ";True"
'        AbreReporteParam gsDSN, Me, rutaReportes & "RptTipoCambioMensual.rpt", crptToWindow, Titulo, "", matriz(), formulas()
'
'    End If
'
'
'
'
'End Sub
'
'Private Sub Grabar()
'
'    If CE(tdbcDe(0).BoundText) = "" Then
'        Mensajes "Seleccione un tipo de moneda la lista", vbOKOnly + vbInformation
'        Exit Sub
'    End If
'
'
'    Dim clsMante As clsMantoTablas
'    Dim i As Integer
'    Dim condicion As Boolean
'    If validarDatos = False Then Exit Sub
'    Set clsMante = New clsMantoTablas
'
'    On Local Error GoTo ErrorEjecucion
'
'    Call CargaArregloMnt
'
'    If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaTipoCambio", lArrMnt(), True) = False Then
'        Mensajes "El proceso no ha concluido. Verificar...", vbInformation
'        Exit Sub
'    End If
'    '-----------------------------------------------------------'
'
'    Dim cMes As String
'    cMes = Right("00" & dtpFecha.Month, 2)
'
'    If UltimoDiaMes(cMes, gsAnio) = dtpFecha.Value Then
'        Call GrabaTcMensual(cMes, tdbnCompra.Value, TCM_COMPRA)
'        Call GrabaTcMensual(cMes, tdbnVenta.Value, TCM_VENTA)
'    End If
'    '-----------------------------------------------------------'
'    Call Cancelar
'    CargaTablaMoneda (IndiceTab)
'    FiltrarRecordSet (IndiceTab)
'    On Error Resume Next
'    lrsTabla(IndiceTab).Find "Tca_dFecha = '" & dtpFecha & "'"
'
'    Mensajes "Los datos se grabaron con exito...", vbInformation + vbOKOnly
'    tdbgMoneda(IndiceTab).HighlightRowStyle = "HighlightRow"
'
'    If Voucher = True Then 'si fue llamado del mant de voucher asignarle el TC VEP a la celda
'        On Error Resume Next
'        frmManAsientosContables.Enabled = True
'        frmManAsientosContables.tdbgDetalle.Columns(16) = Me.tdbnVentaP
'        Unload Me
'        pSetFocus frmManAsientosContables.tdbgDetalle
'        pSendKeys "{Enter}"
'        On Error GoTo 0
'    End If
'
'    If MonAdic = True Then  'si fue llamado del mant de voucher asignarle el TC VEP a la celda
'        On Error Resume Next
'        frmManAsientosContables.Enabled = True
'        frmManAsientosContables.tdbgDetalle.Columns(ColumnaTC) = Me.tdbnVentaP.Value * frmManAsientosContables.tdbtMonedaAdic.Value
'        frmManAsientosContables.ValoresMonedaAdic
'        Unload Me
'        pSetFocus frmManAsientosContables.tdbgDetalle
'
'
'        pSendKeys "{Enter}"
'        On Error GoTo 0
'    End If
'
'    If RegAux = True Then  'si fue llamado de registro de auxiliares
'        On Error Resume Next
'        FrmManRegAuxiliarVentas.Enabled = True
'        FrmManRegAuxiliarVentas.tdbTC.Value = Me.tdbnVentaP.Value
'        Unload Me
'        pSetFocus FrmManRegAuxiliarVentas.tdbTC
'
'        pSendKeys "{Enter}"
'        On Error GoTo 0
'    End If
'    Exit Sub
'
'ErrorEjecucion:
'    Mensajes Str(Err.Number) & Err.Description, vbInformation
'End Sub
'
'Private Function validarDatos() As Boolean
'    validarDatos = False
'    ' *** Validar q los datos necesarios esten ingresados
'    If lTipoMnt = "INSERTAR" Then
'        If ExisteCambio = True Then Exit Function
'    End If
'
'    ' ***
'    validarDatos = True
'End Function
'
'Private Function ExisteCambio() As Boolean
'    Dim rsArreglo As New ADODB.Recordset
'    Dim clDatos As clsMantoTablas
'    Set clDatos = New clsMantoTablas
'    Dim arrDatos() As Variant
'    ' *** Cargando Datos de la Cuenta
'    Dim sqlSp As String
'    ExisteCambio = False
'    sqlSp = "spCn_GrabaTipoCambio 'SEL_REG', '" & gsEmpresa & "', '" & dtpFecha.Value & "', '" & gsMonedaNac & "', '" & tdbcDe(0).BoundText & "', 0, 0,0, 0, '' "
'        arrDatos = Array(sqlSp)
'    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
'    If rsArreglo.State <> 0 Then
'        ExisteCambio = True
'        Mensajes "Cambio con fecha indicada ya existe. Verifique...", vbInformation
'        pSetFocus dtpFecha
'    End If
'    Call CerrarRecordSet(rsArreglo)
'End Function
'
'Private Sub CargaArregloMnt()
'    ' *** Cargar los datos a grabar en un arreglo
'    ReDim lArrMnt(10) As Variant
'    lArrMnt(0) = lTipoMnt           ' Accion
'    lArrMnt(1) = gsEmpresa          ' Empresa
'    lArrMnt(2) = dtpFecha           ' Fecha
'    lArrMnt(3) = gsMonedaNac        ' MonedaOrigen
'    lArrMnt(4) = tdbcDe(0).BoundText ' MonedaDestino
'    lArrMnt(5) = tdbnCompra.Value          ' Compra
'    lArrMnt(6) = tdbnVenta.Value           ' Venta
'    lArrMnt(7) = 0                  ' Compra Publicación
'    lArrMnt(8) = tdbnVentaP         ' Venta Publicación
'    lArrMnt(9) = gsUsuario          ' Usuario
'    lArrMnt(10) = gsPeriodo
'End Sub
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    Dim respuesta As String
'    Select Case KeyCode
'        Case 27:
'            If SSTCentroCosto.TabEnabled(3) = False Then ' *** Grabar
''                respuesta = MsgBox("Esta seguro que desea salir del formulario actual", vbYesNo+ vbQuestion , "Confirmar Salir")
''                If respuesta = vbYes Then Unload Me
'                Unload Me
'            Else
'                respuesta = MsgBox("Desea cancelar la siguiente operación", vbYesNo + vbQuestion, "Confirmar Cancelar")
'                If respuesta = vbYes Then Call Cancelar
'            End If
'        Case 113: If tbrOpciones.Buttons(1).Enabled Then ManNuevo
'        Case 114: If tbrOpciones.Buttons(2).Enabled Then VerDatos
'        Case 115: If tbrOpciones.Buttons(3).Enabled Then Grabar
'        Case 116: If tbrOpciones.Buttons(4).Enabled Then Borrar (IndiceTab)
'        Case 117: If tbrOpciones.Buttons(5).Enabled Then Editar
'        Case 118: If tbrOpciones.Buttons(5).Enabled Then Imprimir
'    End Select
'    ' ***
'End Sub
'
'Private Sub Form_Load()
'   On Error GoTo ERROR
'    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))
'    Dim Mes As String
'    Asientos = False
'    Mes = gsPeriodo
'
'    If Mes = "00" Then Mes = "01"
'    If Mes > "12" Then Mes = "12"
'
'
'    Me.MonAdic = False 'variable que indica si fue llamado del mant de voucher (mon adicional)
'    Me.Voucher = False 'variable que indica si fue llamado del mant de voucher
'
'    Call Centrar_form(Me)
'
'
'    dtpFecha.MinDate = "01/01/1000"
'    dtpFecha.MaxDate = "31/12/2500"
'
'    dtpFechaBus(0).MinDate = "01/01/1000"
'    dtpFechaBus(0).MaxDate = "31/12/2500"
'
'    dtpFechaBus(1).MinDate = "01/01/1000"
'    dtpFechaBus(1).MaxDate = "31/12/2500"
'
'
'    Call LlenaCombos
'    Call LlenaComboMesActivo(tdbcMes(0))
'    Call LlenaComboMesActivo(tdbcMes(1))
'
'    tdbcDe(1).Locked = True
'
'    lTipoMnt = "INSERTAR"
'    SSTCentroCosto.TabEnabled(3) = False
'
'    lRegElim = False
'
'
'    IndiceTab = 0
'    tdbgMoneda(0).HighlightRowStyle = "HighlightRow"
'    tdbgMoneda(1).HighlightRowStyle = "HighlightRow"
'
'    Call CargaMeses
''    Call BuscarTCDAOT
'    On Error Resume Next
'
'    tdbcTipoMensual.BoundText = "1"
'    tdbcDe(1).BoundText = gsMonedaExt
'
'
'    tdbcMes(0).BoundText = Mes
'    tdbcMes(1).BoundText = Mes
'
'    tdbcMes(0).ReBind
'    tdbcMes(1).ReBind
'
'    tdbcDe(0).ReBind
'    tdbcDe(1).ReBind
'
'    tdbcTipoMensual.ReBind
'    tdbcMonedaAdic.ReBind
'
'    DoEvents
'
'    SeteaBarraHerramientas Me.tbrOpciones, gsGrupo
'    SSTCentroCosto.Tab = 0
'
'
'    If Mid(gsGrupo, 3, 1) <> "1" And gsGrupo <> gsPrivilegioAdmin Then
'        cmdGrabar.Enabled = False
'    Else
'        cmdGrabar.Enabled = True
'    End If
'
'
'   Exit Sub
'
'ERROR:
'    Mensajes Err.Description, vbOKOnly + vbCritical
'
'End Sub
'
'Private Sub SeteaBarraHerram()
'
'End Sub
'
'Private Sub BuscarTCDAOT()
'
'End Sub
'
'
'
'Public Sub ConfigurarControlFecha(Indice As Integer)
'   On Error GoTo ERROR
'   Dim FechaIni As Date, FechaFin As Date, NuevaFecha As Date
'   Dim Mes As String
'   Mes = gsPeriodo
'   If Mes < "01" Then Mes = "01"
'   If Mes > "12" Then Mes = "12"
'
'
'   FechaIni = dtpFechaBus(Indice).MinDate
'   FechaFin = dtpFechaBus(Indice).MaxDate
'   NuevaFecha = "01/" + tdbcMes(Indice).BoundText + "/" + gsAnio
'   DoEvents
'   If Indice > 1 Then Exit Sub
'
'    If Val(tdbcMes(Indice).BoundText) > 0 And Val(tdbcMes(Indice).BoundText) < 13 Then
'        dtpFechaBus(Indice).Enabled = True
'
'
'        If Format(NuevaFecha, "yyyyMMdd") <= Format(FechaIni, "yyyyMMdd") Then
'            dtpFechaBus(Indice).MinDate = "01/" + tdbcMes(Indice).BoundText + "/" + gsAnio
'            dtpFechaBus(Indice).MaxDate = UltimoDiaMes(tdbcMes(Indice).BoundText, gsAnio)
'        End If
'
'        If Format(NuevaFecha, "yyyyMMdd") >= Format(FechaFin, "yyyyMMdd") Then
'            dtpFechaBus(Indice).MaxDate = UltimoDiaMes(tdbcMes(Indice).BoundText, gsAnio)
'            dtpFechaBus(Indice).MinDate = "01/" + tdbcMes(Indice).BoundText + "/" + gsAnio
'        End If
'
'        If Val(tdbcMes(Indice).BoundText) = Val(Month(Date)) And gsAnio = Val(Year(Date)) Then
'            dtpFechaBus(Indice).Value = Date
'        Else
'            dtpFechaBus(Indice).Value = "01/" + tdbcMes(Indice).BoundText + "/" + gsAnio
'        End If
'
'        gsPeriodoAnterior(Indice) = tdbcMes(Indice).BoundText
'
'        DoEvents
'    Else
'        dtpFechaBus(Indice).Enabled = False
'        dtpFechaBus(Indice) = "01/" + tdbcMes(Indice).BoundText + "/" + gsAnio
'    End If
'    DoEvents
'    Exit Sub
'ERROR:
'    Mensajes Err.Description & Chr(10) + Chr(13) & '            "Rango: " & dtpFechaBus(Indice).Value & Chr(10) + Chr(13) & '            "Min  : " & dtpFechaBus(Indice).MinDate & Chr(10) + Chr(13) & '            "Max  : " & dtpFechaBus(Indice).MaxDate, vbOKOnly + vbCritical
'End Sub
'Private Sub Form_Unload(Cancel As Integer)
'    Call CerrarRecordSet(lrsTabla(0))
'    Call CerrarRecordSet(lrsTabla(1))
'
'
'    If Me.Voucher = True Or Me.MonAdic = True Then
'        frmManAsientosContables.Enabled = True
'        pSetFocus frmManAsientosContables.tdbgDetalle
'    End If
'
'    If RegAux = True Then
'        FrmManRegAuxiliarVentas.Enabled = True
'    End If
'
'    If Asientos = True Then
'        frmBusTipoAsiento.Enabled = True
'        'frmBusTipoAsiento.Insertar
'        frmBusTipoAsiento.CargaTC
'    End If
'
'    Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))
'
'
'End Sub
'
'Private Sub CargaMeses()
'    Dim sql As String
'    Dim arrDatos() As Variant
'    Dim rsAddItem As ADODB.Recordset
'    Dim clDatos As clsMantoTablas
'
'    Sql = "select * from CNT_TIPO_CAMBIO_MENSUAL " & '          "where emp_ccodigo='" & gsEmpresa & "' and pan_canio='" & gsAnio & "' and " & '          "tca_cmoneda='" & gsMonedaExt & "' and tca_ctipo='" & Me.tdbcTipoMensual.BoundText & "'"
'
'    Set clDatos = New clsMantoTablas
'    arrDatos = Array(sql)
'    Set rsAddItem = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
'
'    If Not rsAddItem Is Nothing Then
'        Do While Not rsAddItem.EOF
'            tdbcTCMes(0) = NE(rsAddItem!Tca_cEne)
'            tdbcTCMes(1) = NE(rsAddItem!Tca_cFeb)
'            tdbcTCMes(2) = NE(rsAddItem!Tca_cMar)
'            tdbcTCMes(3) = NE(rsAddItem!Tca_cAbr)
'            tdbcTCMes(4) = NE(rsAddItem!Tca_cMay)
'            tdbcTCMes(5) = NE(rsAddItem!Tca_cJun)
'            tdbcTCMes(6) = NE(rsAddItem!Tca_cJul)
'            tdbcTCMes(7) = NE(rsAddItem!Tca_cAgo)
'            tdbcTCMes(8) = NE(rsAddItem!Tca_cSet)
'            tdbcTCMes(9) = NE(rsAddItem!Tca_cOct)
'            tdbcTCMes(10) = NE(rsAddItem!Tca_cNov)
'            tdbcTCMes(11) = NE(rsAddItem!Tca_cDic)
'
'            rsAddItem.MoveNext
'        Loop
'    Else
'        tdbcTCMes(0) = 0
'        tdbcTCMes(1) = 0
'        tdbcTCMes(2) = 0
'        tdbcTCMes(3) = 0
'        tdbcTCMes(4) = 0
'        tdbcTCMes(5) = 0
'        tdbcTCMes(6) = 0
'        tdbcTCMes(7) = 0
'        tdbcTCMes(8) = 0
'        tdbcTCMes(9) = 0
'        tdbcTCMes(10) = 0
'        tdbcTCMes(11) = 0
'    End If
'
'    Call CerrarRecordSet(rsAddItem)
'    Set clDatos = Nothing
'End Sub
'
'Private Function Valida() As Boolean
'    If CE(tdbcTipoMensual.BoundText) = "" Then
'        Mensajes "Seleccione un tipo de moneda", vbOKOnly + vbInformation
'        Valida = False
'        Exit Function
'    End If
'
'    If CE(tdbcDe(1).BoundText) = "" Then
'        Mensajes "Seleccione una moneda", vbOKOnly + vbInformation
'        Valida = False
'        Exit Function
'    End If
'
'    Valida = True
'End Function
'
'Private Sub GrabaTcMensual(cMes As String, nvalor As Double, cTipo As Tipo_Cambio)
'
'    Dim sql As String
'    Dim rsAddItem As ADODB.Recordset
'    Dim clDatos As New ClsFuncionesExecute
'    Dim Existe As Boolean
'    On Error GoTo ERROR
'    Dim cCadena As String
'
'    Select Case cMes
'        Case "01": cCadena = "Tca_cEne"
'        Case "02": cCadena = "Tca_cFeb"
'        Case "03": cCadena = "Tca_cMar"
'        Case "04": cCadena = "Tca_cAbr"
'        Case "05": cCadena = "Tca_cMay"
'        Case "06": cCadena = "Tca_cJun"
'        Case "07": cCadena = "Tca_cJul"
'        Case "08": cCadena = "Tca_cAgo"
'        Case "09": cCadena = "Tca_cSet"
'        Case "10": cCadena = "Tca_cOct"
'        Case "11": cCadena = "Tca_cNov"
'        Case "12": cCadena = "Tca_cDic"
'    End Select
'
'    Sql = "select count(emp_ccodigo) as Registro from CNT_TIPO_CAMBIO_MENSUAL " & '          "where emp_ccodigo='" & gsEmpresa & "' and pan_canio='" & gsAnio & "' and " & '          "tca_cmoneda='" & gsMonedaExt & "' and tca_ctipo='" & cTipo & "' "
'
'    Existe = False
'    Set rsAddItem = clDatos.fRetornaRS(sql)
'    If Not rsAddItem Is Nothing Then
'        Do While Not rsAddItem.EOF
'            If NE(rsAddItem!Registro) > 0 Then
'                Existe = True
'            End If
'            rsAddItem.MoveNext
'        Loop
'    End If
'
'    If Existe = True Then
'        Sql = "Update CNT_TIPO_CAMBIO_MENSUAL set " & cCadena & "=" & nvalor & " " & '              "where emp_ccodigo='" & gsEmpresa & "' and pan_canio='" & gsAnio & "' and " & '              "tca_cmoneda='" & gsMonedaExt & "' and tca_ctipo='" & cTipo & "'"
'
'    Else
'        Sql = "Insert into CNT_TIPO_CAMBIO_MENSUAL (" & cCadena & "," & '                                                    "emp_ccodigo,pan_canio,tca_cmoneda,tca_ctipo) values (" & '                                                    nvalor & "," & '                                                    "'" & gsEmpresa & "','" & gsAnio & "','" & gsMonedaExt & "'," & '                                                    "'" & cTipo & "')"
'    End If
'
'    clDatos.pEjecutaSQL (sql)
'    Call CerrarRecordSet(rsAddItem)
'    Set clDatos = Nothing
'
'    Exit Sub
'
'ERROR:
'    Call CerrarRecordSet(rsAddItem)
'    Set clDatos = Nothing
'
'End Sub
'Private Sub cmdGrabar_Click()
'    If Valida = False Then Exit Sub
'
'    Dim sql As String
'    Dim rsAddItem As ADODB.Recordset
'    Dim clDatos As New ClsFuncionesExecute
'    Dim Existe As Boolean
'    Screen.MousePointer = vbHourglass
'    On Error GoTo ERROR
'    Sql = "select count(*) as Registro from CNT_TIPO_CAMBIO_MENSUAL " & '          "where emp_ccodigo='" & gsEmpresa & "' and pan_canio='" & gsAnio & "' and " & '          "tca_cmoneda='" & gsMonedaExt & "' and tca_ctipo='" & Me.tdbcTipoMensual.BoundText & "'"
'
'    Existe = False
'    Set rsAddItem = clDatos.fRetornaRS(sql)
'    If Not rsAddItem Is Nothing Then
'        Do While Not rsAddItem.EOF
'            If NE(rsAddItem!Registro) > 0 Then
'                Existe = True
'            End If
'            rsAddItem.MoveNext
'        Loop
'    End If
'
'    If Existe = True Then
'        Sql = "Update CNT_TIPO_CAMBIO_MENSUAL set Tca_cEne=" & NE(tdbcTCMes(0)) & "," & '                                                 "Tca_cFeb=" & NE(tdbcTCMes(1)) & "," & '                                                 "Tca_cMar=" & NE(tdbcTCMes(2)) & "," & '                                                 "Tca_cAbr=" & NE(tdbcTCMes(3)) & "," & '                                                 "Tca_cMay=" & NE(tdbcTCMes(4)) & "," & '                                                 "Tca_cJun=" & NE(tdbcTCMes(5)) & "," & '                                                 "Tca_cJul=" & NE(tdbcTCMes(6)) & "," & '                                                 "Tca_cAgo=" & NE(tdbcTCMes(7)) & "," & '                                                 "Tca_cSet=" & NE(tdbcTCMes(8)) & "," & '                                                 "Tca_cOct=" & NE(tdbcTCMes(9)) & "," & '                                                 "Tca_cNov=" & NE(tdbcTCMes(10)) & "," & '
'       "Tca_cDic=" & NE(tdbcTCMes(11)) & " " & '              "where emp_ccodigo='" & gsEmpresa & "' and pan_canio='" & gsAnio & "' and " & '              "tca_cmoneda='" & gsMonedaExt & "' and tca_ctipo='" & Me.tdbcTipoMensual.BoundText & "'"
'
'    Else
'        Sql = "Insert into CNT_TIPO_CAMBIO_MENSUAL (Tca_cEne,Tca_cFeb,Tca_cMar,Tca_cAbr," & '                                                    "Tca_cMay,Tca_cJun,Tca_cJul,Tca_cAgo," & '                                                    "Tca_cSet,Tca_cOct,Tca_cNov,Tca_cDic," & '                                                    "emp_ccodigo,pan_canio,tca_cmoneda,tca_ctipo) values (" & '                                                     NE(tdbcTCMes(0)) & "," & NE(tdbcTCMes(1)) & "," & '                                                     NE(tdbcTCMes(2)) & "," & NE(tdbcTCMes(3)) & "," & '                                                     NE(tdbcTCMes(4)) & "," & NE(tdbcTCMes(5)) & "," & '                                                     NE(tdbcTCMes(6)) & "," & NE(tdbcTCMes(7)) & "," & '                                                     NE(tdbcTCMes(8)) & "," & NE(tdbcTCMes(9)) & "," & '                                                     NE(tdbcTCMes(10)) & "," & NE(tdbcTCMes(11)) & "," & '
'                                                "'" & gsEmpresa & "','" & gsAnio & "','" & gsMonedaExt & "'," & '                                                    "'" & Me.tdbcTipoMensual.BoundText & "')"
'    End If
'
'    clDatos.pEjecutaSQL (sql)
'    Call CerrarRecordSet(rsAddItem)
'    Set clDatos = Nothing
'
'    CargaMeses
'    Screen.MousePointer = vbNormal
'    Mensajes "Se grabo correctamente los tipos de cambios mensuales", vbOKOnly + vbInformation
'
'    Exit Sub
'
'ERROR:
'    Call CerrarRecordSet(rsAddItem)
'    Set clDatos = Nothing
'    Screen.MousePointer = vbNormal
'    Mensajes "No se grabo correctamente los tipos de cambios mensuales", vbOKOnly + vbInformation
'
'
'End Sub
'
'Private Sub LlenaCombos()
'    On Error GoTo ERROR
'    Dim sqlcombos As String
'
'    sqlcombos = "SELECT Mon_cCodigo, Mon_cNombreLargo From CNT_TIPO_MONEDA " & '                "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Mon_cMNac<> '1' " & '                "ORDER BY Mon_cNombreLargo"
'    LlenarComboAddItem tdbcDe(0), sqlcombos
'
'    sqlcombos = "SELECT Mon_cCodigo, Mon_cNombreLargo From CNT_TIPO_MONEDA " & '                "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Mon_cMExt = '1' " & '                "ORDER BY Mon_cNombreLargo"
'    LlenarComboAddItem tdbcDe(1), sqlcombos
'
'    sqlcombos = "SELECT Mon_cCodigo, Mon_cNombreLargo, Mon_cMNac, Mon_cMExt, Mon_cNombreCorto From CNT_TIPO_MONEDA " & '                "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Mon_cMNac <> '1' and Mon_cMExt <> '1' " & '                "ORDER BY Mon_cNombreLargo"
'
'    LlenarComboAddItem tdbcMonedaAdic, sqlcombos, True
'
'    sqlcombos = "select Tab_cCodigo, Tab_cDescripCampo from tabla  " & '                "where emp_ccodigo='" & gsEmpresa & "' and tab_ctabla='046' AND Tab_cDescripCampo LIKE 'CIERRE%' " & '                "ORDER BY Tab_cCodigo"
'
'    LlenarComboAddItem tdbcTipoMensual, sqlcombos
'    Exit Sub
'ERROR:
'
'End Sub
'
'
'
'Private Sub CargaTablaMoneda(Indice As Integer)
'    If Indice > 1 Then Exit Sub
'    Dim sqlSp As String
'    Dim clDatos As clsMantoTablas
'    Dim arrDatos() As Variant
'    Set clDatos = New clsMantoTablas
'    Set lrsTabla(Indice) = New ADODB.Recordset
'    Set tdbgMoneda(Indice).DataSource = Nothing
'    Dim fechaAux As String
'    Dim Moneda As String
'
'    If Indice = 0 Then
'        Moneda = gsMonedaExt
'    Else
'        Moneda = tdbcMonedaAdic.BoundText
'    End If
'
'    fechaAux = "01/" & Right("00" & dtpFechaBus(Indice).Month, 2) & "/" & gsAnio
'    sqlSp = "spCn_GrabaTipoCambio 'SEL_ALL', '" & gsEmpresa & "', '" & fechaAux & "', '" & gsMonedaNac & "', '" & Moneda & "', 0, 0,0, 0, ''"
'    arrDatos = Array(sqlSp)
'    Set lrsTabla(Indice) = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
'    If Not lrsTabla(Indice) Is Nothing Then
'        If Not (lrsTabla(Indice).EOF And lrsTabla(Indice).BOF) Then
'
'        lrsTabla(Indice).Sort = "Tca_dFecha desc"
'        tdbgMoneda(Indice).DataSource = lrsTabla(Indice)
'
'        End If
'    End If
'End Sub
'
'
'Private Sub CargaDatosRegistro(Indice As Integer)
'    Dim rsArreglo As New ADODB.Recordset
'    Dim clDatos As clsMantoTablas
'    Set clDatos = New clsMantoTablas
'    Dim arrDatos() As Variant
'    ' *** Cargando Datos de la Cuenta
'    Dim sqlSp As String
'
'    Dim sMoneda As String
'
'
'
'    With tdbgMoneda(Indice)
'        On Error GoTo serror
'
'        sMoneda = .Columns(1).Value
'
'        If SSTCentroCosto.Tab = 0 Then sMoneda = gsMonedaExt
'        If SSTCentroCosto.Tab = 1 Then sMoneda = tdbcMonedaAdic.BoundText
'
'        sqlSp = "spCn_GrabaTipoCambio 'SEL_REG', '" & gsEmpresa & "', '" & .Columns(0).Value & "', '" & gsMonedaNac & "','" & sMoneda & "', 0, 0,0, 0, '' "
'    End With
'
'    arrDatos = Array(sqlSp)
'    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
'    If rsArreglo.State = 0 Then
'        lRegElim = True
'        'Mensajes "Seleccione un tipo de moneda", vbInformation
'        Set rsArreglo = Nothing
'        'Exit Sub
'    End If
'    ' *** Asignando Datos del Tipo de Cambio
'    dtpFecha = CE(rsArreglo!Tca_dFecha)
'
'    tdbcDe(0).BoundText = CE(rsArreglo!Tca_cCodigoDestino)
'
'    'tdbcA.BoundText = rsArreglo!Tca_cCodigoDestino
'    tdbnCompra = NE(rsArreglo!Tca_nCompra)
'    tdbnVenta = NE(rsArreglo!Tca_nVenta)
'    'tdbnCompraP = NuloNum(rsArreglo!Tca_nCompraP)  Operacion
'    tdbnVentaP = NE(rsArreglo!Tca_nVentaP)
'    Call CerrarRecordSet(rsArreglo)
'    ' ***
'    Exit Sub
'serror:
'    Mensajes "Seleccione un tipo de moneda", vbOKOnly + vbInformation
'End Sub
'
'Private Sub FiltrarRecordSet(Indice As Integer)
'    DoEvents
'    ' *** Filtrar segun los textos indicados
'    Dim cadena As String
'    Dim filtros(2) As String
'    Dim i As Integer
'    If lrsTabla(Indice) Is Nothing Then Exit Sub
'    If IsNull(dtpFechaBus) Then Exit Sub
'    cadena = ""
'    If chkFecha(Indice).Value = 1 Then filtros(2) = "Tca_dFecha like '" & Me.dtpFechaBus(Indice) & "'"
'    For i = 0 To 2
'        If filtros(i) <> "" Then
'            If cadena = "" Then
'                cadena = cadena + filtros(i)
'            Else
'                cadena = cadena + " and " + filtros(i)
'            End If
'        End If
'    Next
'    ' *** Filtrando segun campos
'    lrsTabla(Indice).Filter = 0
'    If Trim(cadena) <> "" Then
'        On Error Resume Next
'        lrsTabla(Indice).Filter = cadena
'    Else
'        lrsTabla(Indice).Filter = 0
'    End If
'End Sub
'
'Private Sub tdbcMes_ItemChange(Index As Integer)
'    ConfigurarControlFecha (Index)
'    CargaTablaMoneda (Index)
'End Sub
'
'Private Sub tdbcTipoMensual_ItemChange()
'    Call CargaMeses
'End Sub
'
'Private Sub tdbgMoneda_GotFocus(Index As Integer)
'    On Error Resume Next
'    tdbgMoneda(IndiceTab).HighlightRowStyle = "HighlightRow"
'End Sub
'
'Private Sub tdbgMoneda_HeadClick(Index As Integer, ByVal ColIndex As Integer)
'If Not lrsTabla(Index) Is Nothing Then
'    If lrsTabla(Index).RecordCount > 0 Then
'        lrsTabla(Index).Sort = tdbgMoneda(Index).Columns(ColIndex).DataField
'        tdbgMoneda(Index).DataSource = lrsTabla(Index)
'    End If
'End If
'End Sub
'
'Private Sub tdbgMoneda_KeyPress(Index As Integer, KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        Editar
'    End If
'End Sub
'
'Private Sub tdbgMoneda_LostFocus(Index As Integer)
'    tdbgMoneda(Index).HighlightRowStyle = ""
'End Sub
'
'
'Private Sub tdbcMonedaAdic_ItemChange()
'    If IndiceTab < 2 Then
'    tdbcMes_ItemChange (IndiceTab)
'    End If
'
'End Sub
'
'Private Sub tdbnVentaP_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        Grabar
'    End If
'End Sub
'
'
