VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmBusPeriodo 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Cambio de Empresa y Año de Trabajo"
   ClientHeight    =   3060
   ClientLeft      =   2676
   ClientTop       =   4188
   ClientWidth     =   6864
   Icon            =   "frmBusPeriodo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Picture         =   "frmBusPeriodo.frx":0ECA
   ScaleHeight     =   3060
   ScaleWidth      =   6864
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7875
      Top             =   2250
   End
   Begin MSDataListLib.DataCombo tdbcEmpresa 
      Height          =   264
      Left            =   1032
      TabIndex        =   1
      Top             =   1656
      Width           =   4788
      _ExtentX        =   8446
      _ExtentY        =   466
      _Version        =   393216
      Style           =   2
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo tdbcAnio 
      Height          =   288
      Left            =   1032
      TabIndex        =   2
      Top             =   2340
      Width           =   1212
      _ExtentX        =   2138
      _ExtentY        =   508
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin MSForms.CommandButton cmdAceptar 
      Height          =   390
      Left            =   4155
      TabIndex        =   5
      Top             =   2295
      Width           =   1665
      Caption         =   "  Ingresar"
      PicturePosition =   327683
      Size            =   "2937;688"
      Picture         =   "frmBusPeriodo.frx":728F
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Año:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   1
      Left            =   1035
      TabIndex        =   4
      Top             =   2070
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   0
      Left            =   1035
      TabIndex        =   3
      Top             =   1350
      Width           =   930
   End
   Begin VB.Label lblEmpresa 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "XXX"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   1035
      TabIndex        =   0
      Top             =   1665
      Width           =   4830
   End
End
Attribute VB_Name = "frmBusPeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : frmBusPeriodo
'    Project    : Contabilidad
'
'    Description: Formulario de cambio de ejercicio
'--------------------------------------------------------------------------------

Option Explicit
Dim lArrAnio As New XArrayDB
Dim Mes As String
Dim Segundos As Integer
Dim gsGrupo As String

Dim rstEmpresa As New ADODB.Recordset
Dim rstAnio As New ADODB.Recordset

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Grupo
' Description:       Propiedad de asignacion de grupo
'
' Parameters :       Grupo (String)
'--------------------------------------------------------------------------------
Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       cmdAceptar_Click
' Description:       Evento que se ejecuta al presionar el boton aceptar cambio de ejercicio
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub cmdAceptar_Click()
    cmdAceptar.Enabled = False
    
    DoEvents
    
    If CE(tdbcEmpresa.BoundText) = "" Then
       DoEvents
       Mensajes "Seleccione una empresa de la lista"
       cmdAceptar.Enabled = True
       Exit Sub
    End If

    If CE(tdbcAnio.BoundText) = "" Then
       DoEvents
       Mensajes "Seleccione un periodo contable de la lista"
       cmdAceptar.Enabled = True
       Exit Sub
    End If

    Screen.MousePointer = vbHourglass
    gsByMoneda = "0" 'SIEMPRE ES MONEDA NACIONAL NO DEBE EXIGIR EL INGRESO DE DOLARES NE(rstEmpresa.Fields("Emp_Bymoneda"))
    Mes = gsPeriodo
    gsSucursal = CE(rstEmpresa.Fields("Emp_cCodsuc"))
    
    gsInicio = True
    
    gsRUC = CE(rstEmpresa.Fields("Emp_cNumRuc"))
    gsAnio = tdbcAnio.BoundText
    gsEmpresa = CE(rstEmpresa.Fields("Emp_cCodigo"))
    gsEmpresaNom = CE(rstEmpresa.Fields("Emp_cNombreLargo"))
    'Obtiene el valor del Campo Emp_Bymoneda
    gintBiMoneda = IIf(IsNull(CE(rstEmpresa.Fields("Emp_Bymoneda"))), 0, CE(rstEmpresa.Fields("Emp_Bymoneda")))
    gintPercepcion = CE(rstEmpresa.Fields("Emp_AgentePercepcion"))
    gintRetencion = CE(rstEmpresa.Fields("Emp_AgenteRetencion"))
    
    If CE(gsPeriodo) = "" Then
        gsPeriodo = "00"
        frmMDIConta.stbMdi.Panels(6).Text = "  " & NombreMes(gsPeriodo) & " DEL " & gsAnio & "  "
    End If
        
    Call GrabaPeriodoActivo
    Call pCargaCfgLibro
    Call ValidarCuentaCostoVenta
    Unload Me

    If gsPLE = "0" Then
        frmMDIConta.mnuLibroElec.Enabled = False
    Else
       frmMDIConta.mnuLibroElec.Enabled = True
    End If
    
    Dim sCfl_cTipoPlan As String
    sCfl_cTipoPlan = ExtraeCampo("Cfl_cTipoPlan", "CNT_CONFIG_LIBROS", "Emp_cCodigo='" & gsEmpresa & "' And Pan_cAnio='" & gsAnio & "'")
    gsTipoPlan = IIf(sCfl_cTipoPlan = "", 0, sCfl_cTipoPlan)
       
    '----------------------------------------
    Call frmMDIConta.ActivaMenuSegunTipoPlan
    '----------------------------------------
    
    Screen.MousePointer = vbNormal
    
End Sub


'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       BuscaDatosLogeo
' Description:       Procedimiento de busqueda de ultimo dato de logeo registrado del usuario
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub BuscaDatosLogeo()
    Dim sqlSp As String
    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As New clsMantoTablas
    Dim arrDatos() As Variant
    
    On Error GoTo serror
    
    gsAnio = ""
    gsEmpresa = ""
    gsPeriodo = "01"
    
    sqlSp = "EXEC spSg_GrabaUsuarios 'DATOS_LOGEO', '" & gsUsuario & "'"
    arrDatos = Array(sqlSp)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())

    
    If Not rsArreglo Is Nothing Then
        If rsArreglo.State = adStateOpen Then
            If Not (rsArreglo.EOF And rsArreglo.BOF) Then
                gsAnio = CE(rsArreglo.Fields("LOG_PANANIO"))
                gsEmpresa = CE(rsArreglo.Fields("LOG_EMPCOD"))
                gsPeriodo = CE(rsArreglo.Fields("LOG_PERIODO"))
            End If
        End If
    End If
    
serror:
    
    tdbcEmpresa.BoundText = gsEmpresa

    Call CerrarRecordSet(rsArreglo)
    Set clDatos = Nothing
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       LlenaEmpresas
' Description:       Procedimiento de llenado de empresas segun el usuariologeado
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub LlenaEmpresas()
    Dim sqlCadena As String
    Dim sAux As String
    
    If gsUsuario = "ADMINISTRADOR" Then
        sAux = "BUSCARXADMIN_LOG"
    Else
        sAux = "BUSCARXUSUARIO_LOG"
    End If
    
    sqlCadena = "spCN_GestionEmpresas '" & sAux & "', '','','','','','','','','','','','','','" & gsUsuario & "'"
    
    '----------------------------------------------
    
    Call LlenarRecordSet(sqlCadena, rstEmpresa)
    
    Set tdbcEmpresa.RowSource = rstEmpresa
        tdbcEmpresa.ListField = "Emp_cNombreLargo"
        tdbcEmpresa.BoundColumn = "Emp_cCodigo"

        
    If gsEmpresa <> "" Then
        tdbcEmpresa.BoundText = gsEmpresa
        tdbcAnio.BoundText = gsAnio
    End If
    If GetRsRecordCount(rstEmpresa) > 0 Then
        If tdbcEmpresa.BoundText = "" Then tdbcEmpresa.BoundText = CE(rstEmpresa!Emp_cCodigo)
    End If
    
    'Call BuscaDatosLogeo
    pSetFocus tdbcEmpresa
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Form_Load
' Description:       Evento que se ejecuta al iniciar el formulario
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub Form_Load()
    Segundos = 0
    Label1(0).ForeColor = RGB(4, 62, 74)
    Label1(1).ForeColor = RGB(4, 62, 74)
    
    
    If gsInicio = True Then
        tdbcEmpresa.Locked = True
        cmdAceptar.Enabled = True
        tdbcEmpresa.Visible = False
        lblEmpresa.Caption = gsEmpresaNom
        cmdAceptar.Caption = " Cambiar"
    Else
        tdbcEmpresa.Visible = True
        tdbcEmpresa.Locked = False
        cmdAceptar.Enabled = False
        lblEmpresa.Caption = ""
        
        tmTimer.Enabled = True
        
    End If
    
    lsFecha = FechaServidor
    
    Call LlenaEmpresas
    DoEvents
    
    pSetFocus tdbcEmpresa
End Sub
'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Form_Unload
' Description:       Evento que se ejecuta al cerrar el formulario
'
' Parameters :       Cancel (Integer)
'--------------------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
    
    If gsCambioEmpresa = True Then
        Call CargaMDIPrincipal
    Else
        If gsInicio = False Then
            
            End
        Else
            Call CargaMDIPrincipal
        End If
    End If
    
End Sub

Private Sub CargaMDIPrincipal()

    gsCambioEmpresa = False
    frmMDIConta.Show
    frmMDIConta.Caption = gsNombreModulo & gsVersion & " - Empresa : " & gsEmpresaNom
    Call ConfigurarBarraEstado
    
    Set lArrAnio = Nothing

End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbcAnio_KeyDown
' Description:       Evento que se ejecuta al presionar una tecla en el combo de años
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbcAnio_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then pSetFocus cmdAceptar
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       BuscaAnios
' Description:       Procedimiento de busqueda de anios por la empresa ingresada en el parametro
'
' Parameters :       sEmpresa (String)
'--------------------------------------------------------------------------------
Private Sub BuscaAnios(sEmpresa As String)
    Dim sqlCadena As String
    Dim posicion As Integer
    
    On Local Error GoTo ErrorEjecucion
    
    sqlCadena = "EXEC spCn_GrabaAnio 'BUSCA_ANIO_EMP', '" & sEmpresa & "','','','','" & gsUsuario & "' "
    
    
    Dim gtxtSQL As String
   
    If tdbcEmpresa.BoundText <> "" Then rstEmpresa.Bookmark = tdbcEmpresa.SelectedItem
    
    Set tdbcAnio.RowSource = Nothing
    Call CerrarRecordSet(rstAnio)
    Call LlenarRecordSet(sqlCadena, rstAnio)
    If rstAnio.State = 1 Then
        Set tdbcAnio.RowSource = rstAnio
        tdbcAnio.ListField = "Pan_cAnio"
        tdbcAnio.BoundColumn = "Pan_cAnio"
        rstAnio.MoveLast
        tdbcAnio.BoundText = CE(rstAnio!Pan_cAnio)
    Else
        tdbcAnio.BoundText = ""
    End If
    
    'ComboArreglo lArrAnio, tdbcAnio, sqlCadena
    'tdbcAnio.BoundText = gsAnio
    
    Exit Sub
ErrorEjecucion:

End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbcEmpresa_Change
' Description:       Evento que se ejecuta al cambiar la empresa
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbcEmpresa_Change()
    If gsInicio = False Then
        Call BuscaAnios(tdbcEmpresa.BoundText)
        'Call BuscaAnios(gsEmpresa)
    Else
        Call BuscaAnios(gsEmpresa)
    End If
    
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbcEmpresa_ItemChange
' Description:       Evento que se ejecuta al cambiar la empresa
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbcEmpresa_ItemChange()
            
End Sub
'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbcEmpresa_KeyDown
' Description:       Evento que se ejecuta al presionar una tecla en la empresa
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbcEmpresa_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 Then pSetFocus tdbcAnio
End Sub
'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       NombreUsuario
' Description:       Busca el nombre del usuario
'
' Parameters :       usuario (String)
'--------------------------------------------------------------------------------
Private Function NombreUsuario(usuario As String) As String
    ' *** Validar el ingreso al sistema
    Dim sqlSp As String
    Dim rsArreglo As ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
    
    ' *** Verificando q cuenta exista
    Set clDatos = New clsMantoTablas
    sqlSp = "spSg_GrabaUsuarios 'SEL_REG', '" & usuario & "', '', '', '', '', '', '' ,'" & gsSOFT & "'"
    arrDatos = Array(sqlSp)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo Is Nothing Then
        NombreUsuario = ""
    Else
        NombreUsuario = rsArreglo("usu_cNombres").Value
    End If
    CerrarRecordSet rsArreglo
    ' ***
End Function

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tmTimer_Timer
' Description:       Evento que se ejecuta al ejecutar el timer
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tmTimer_Timer()
    Segundos = Segundos + 1
    'cmdAceptar.Caption = " Espere " & CStr(3 - Segundos) & "seg"
    If Segundos = 1 Then
        cmdAceptar.Caption = " Ingresar"
        tmTimer.Enabled = False
        cmdAceptar.Enabled = True
        'pSetFocus cmdAceptar
    End If
    
End Sub
