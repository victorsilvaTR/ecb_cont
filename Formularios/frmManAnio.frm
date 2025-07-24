VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmManAnio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Apertura de Periodo Contables"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7290
   Icon            =   "frmManAnio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3960
   ScaleWidth      =   7290
   Begin VB.Frame fraTodo 
      Height          =   3840
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   7170
      Begin TDBNumber6Ctl.TDBNumber tdbnAnio 
         Height          =   300
         Left            =   2220
         TabIndex        =   1
         Top             =   915
         Width           =   1125
         _Version        =   65536
         _ExtentX        =   1984
         _ExtentY        =   529
         Calculator      =   "frmManAnio.frx":0ECA
         Caption         =   "frmManAnio.frx":0EEA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmManAnio.frx":0F56
         Keys            =   "frmManAnio.frx":0F74
         Spin            =   "frmManAnio.frx":0FCC
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   3000
         MinValue        =   2000
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   5
         Value           =   2004
         MaxValueVT      =   1447624709
         MinValueVT      =   1885667333
      End
      Begin TrueOleDBList70.TDBCombo tdbcEmpresa 
         Height          =   300
         Left            =   2220
         TabIndex        =   2
         Tag             =   "enabled"
         Top             =   450
         Width           =   3225
         _ExtentX        =   5689
         _ExtentY        =   529
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         _DropdownWidth  =   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).AllowRowSizing=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=609"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=529"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=847"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=767"
         Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=1138"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1058"
         Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(15)=   "Column(2).Visible=0"
         Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
         Splits.Count    =   1
         Appearance      =   1
         BorderStyle     =   1
         ComboStyle      =   2
         AutoCompletion  =   0   'False
         LimitToList     =   0   'False
         ColumnHeaders   =   0   'False
         ColumnFooters   =   0   'False
         DataMode        =   5
         DefColWidth     =   0
         Enabled         =   -1  'True
         HeadLines       =   1
         FootLines       =   1
         RowDividerStyle =   0
         Caption         =   ""
         EditFont        =   "Size=6.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         LayoutName      =   ""
         LayoutFileName  =   ""
         MultipleLines   =   0
         EmptyRows       =   -1  'True
         CellTips        =   0
         EditHeight      =   299.906
         AutoSize        =   -1  'True
         GapHeight       =   30.047
         ListField       =   ""
         BoundColumn     =   ""
         IntegralHeight  =   0   'False
         CellTipsWidth   =   0
         CellTipsDelay   =   1000
         AutoDropdown    =   -1  'True
         RowTracking     =   -1  'True
         RightToLeft     =   0   'False
         RowMember       =   ""
         MouseIcon       =   0
         MouseIcon.vt    =   3
         MousePointer    =   0
         MatchEntryTimeout=   2000
         OLEDragMode     =   0
         OLEDropMode     =   0
         AnimateWindow   =   0
         AnimateWindowDirection=   0
         AnimateWindowTime=   200
         AnimateWindowClose=   0
         DropdownPosition=   0
         Locked          =   -1  'True
         ScrollTrack     =   0   'False
         RowDividerColor =   12632256
         RowSubDividerColor=   12632256
         AddItemSeparator=   ";"
         _PropDict       =   $"frmManAnio.frx":0FF4
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=675,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Arial"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(20)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(26)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
         _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Named:id=33:Normal"
         _StyleDefs(45)  =   ":id=33,.parent=0"
         _StyleDefs(46)  =   "Named:id=34:Heading"
         _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(48)  =   ":id=34,.wraptext=-1"
         _StyleDefs(49)  =   "Named:id=35:Footing"
         _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(51)  =   "Named:id=36:Selected"
         _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(53)  =   "Named:id=37:Caption"
         _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(55)  =   "Named:id=38:HighlightRow"
         _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(57)  =   "Named:id=39:EvenRow"
         _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(59)  =   "Named:id=40:OddRow"
         _StyleDefs(60)  =   ":id=40,.parent=33"
         _StyleDefs(61)  =   "Named:id=41:RecordSelector"
         _StyleDefs(62)  =   ":id=41,.parent=34"
         _StyleDefs(63)  =   "Named:id=42:FilterBar"
         _StyleDefs(64)  =   ":id=42,.parent=33"
      End
      Begin MSForms.CommandButton cmdReplicar 
         Height          =   570
         Left            =   405
         TabIndex        =   7
         Top             =   2160
         Width           =   6300
         Caption         =   " Replicar Datos del AÑO ACTUAL DEL SISTEMA"
         PicturePosition =   327683
         Size            =   "11112;1005"
         Picture         =   "frmManAnio.frx":107B
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdAperturar 
         Height          =   570
         Left            =   405
         TabIndex        =   6
         Top             =   1530
         Width           =   6300
         Caption         =   " Apertura de Año y Replica de Datos del AÑO ACTUAL DEL SISTEMA"
         PicturePosition =   327683
         Size            =   "11112;1005"
         Picture         =   "frmManAnio.frx":2A0D
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdEliminarEjercicio 
         Height          =   615
         Left            =   405
         TabIndex        =   5
         Top             =   2835
         Visible         =   0   'False
         Width           =   6300
         Caption         =   " Elimina el AÑO ACTUAL DEL SISTEMA"
         PicturePosition =   327683
         Size            =   "11112;1085"
         Picture         =   "frmManAnio.frx":439F
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "EMPRESA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   1140
         TabIndex        =   4
         Top             =   450
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "AÑO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1125
         TabIndex        =   3
         Top             =   990
         Width           =   375
      End
   End
   Begin VB.Label lblTitulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BALANCE DE COMPROBACION"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   4365
   End
End
Attribute VB_Name = "frmManAnio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : frmManAnio
'    Project    : Contabilidad
'
'    Description: Formulario de apertura de ejercicios
'--------------------------------------------------------------------------------
Option Explicit
Public gsAnioForm As String

Dim gsGrupo As String
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
' Procedure  :       cmdAperturar_Click
' Description:       Evento que se ejecuta al hacer clic en aperturar ejercicio
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub cmdAperturar_Click()
    Dim respuesta As String
    
    ' *** Verificar q registro no existe.
    If ConsultaAño("EXISTEREG", tdbnAnio) = True Then
        Mensajes "El año seleccionado ya existe.", vbInformation
        Me.cmdReplicar.Enabled = True
        Exit Sub
    End If
    
    ' *** Verificar si se ha registrado movimientos el año anterior.
    If ConsultaAño("EXISTEMOV", tdbnAnio - 1) = False Then
        respuesta = MsgBox("El año anterior no ha sido creado o no tiene movimientos registrados." + Chr(13) + _
                    "Esta seguro de Aperturar este año", vbYesNo + vbQuestion, "Confirmar Apertura y Replica de Año")
        If respuesta = vbNo Then
            Exit Sub
        End If
'        Mensajes "No se puede aperturar año. " + Chr(13) +
'        "El año anterior no ha sido creado o no tiene movimientos registrados. Verifique", vbInformation
    End If
    
    ' *** Aperturar
    Screen.MousePointer = vbHourglass
    cmdAperturar.Enabled = False
    cmdReplicar.Enabled = False
    DoEvents
    
    Call Aperturar
    
    cmdAperturar.Enabled = True
    cmdReplicar.Enabled = True
    
    Screen.MousePointer = vbNormal
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Aperturar
' Description:       Procedimiento que apertura el ejercicio de una empresa
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub Aperturar()
    Dim clsMante As clsMantoTablas
    Dim lArrMnt() As Variant
    
    Set clsMante = New clsMantoTablas
    ReDim lArrMnt(5) As Variant
    ' *** Grabando Centro de Costo
    On Local Error GoTo ErrorEjecucion
    
    Call EscribirLog("Inicando el proceso de generacion de asiento de apertura de la empresa " & gsEmpresa & " " & gsEmpresaNom, Me.Name)
    
    lArrMnt(0) = "INSERTAR"         ' Accion
    lArrMnt(1) = Me.tdbcEmpresa.BoundText   ' Empresa
    lArrMnt(2) = tdbnAnio           ' Codigo(año)
    lArrMnt(3) = "A"                 ' Codigo(año del Sistema)
    lArrMnt(4) = gsUsuario                ' Nombre Plantilla
    lArrMnt(5) = gsUsuario          ' Usuario
        
    If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaAnio", lArrMnt(), False) = False Then
        Mensajes "El proceso no ha concluido. Verificar...", vbInformation
        Set clsMante = Nothing
        Exit Sub
    End If
    
    Dim lArrPlantCO() As Variant        ' *** Arreglo para los mantenimientos
    ReDim lArrPlantCO(3) As Variant
    lArrPlantCO(0) = gsEmpresa           ' Empresa
    lArrPlantCO(1) = gsAnio              ' Anio
    lArrPlantCO(2) = gsBD
    If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_UpdPeriodoPlantillaEEFF", lArrPlantCO(), True) = False Then
     Debug.Print "No se actualizo..."
    End If
        
    Set clsMante = New clsMantoTablas
    
    ReDim lArrMnt(2) As Variant
    lArrMnt(0) = Me.tdbcEmpresa.BoundText   ' Empresa
    lArrMnt(1) = tdbnAnio           ' Codigo(año)
    lArrMnt(2) = gsUsuario          ' Usuario
    
    If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_ReplicarAnio", lArrMnt(), True) = False Then
        Mensajes "El proceso no ha concluido. Verificar...", vbInformation
        Set clsMante = Nothing
        Exit Sub
    End If
        
    Mensajes "El año se aperturo con exito...", vbInformation
    
    Call EscribirLog("Finalizo el proceso de generacion de asiento de apertura de la empresa " & gsEmpresa & " " & gsEmpresaNom, Me.Name)
    
    Set clsMante = Nothing
    Exit Sub
ErrorEjecucion:
    Mensajes Str(Err.Number) & Err.Description, vbInformation
    Resume
    Call EscribirLog("Error al aperturar año contable, [" & Err.Description & "] de la empresa " & gsEmpresa & " " & gsEmpresaNom, Me.Name)
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       ConsultaAño
' Description:       Funcion que retorna si el ejercicio de la empresa existe
'
' Parameters :       Tipo (String)
'                    año (String)
'--------------------------------------------------------------------------------
Private Function ConsultaAño(Tipo As String, año As String) As Boolean
    Dim rsDatos As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Dim sqlDatos As String
    Dim arrDatos() As Variant
    
    Set clDatos = New clsMantoTablas
    sqlDatos = "spCn_GrabaAnio '" & Tipo & "', '" & Me.tdbcEmpresa.BoundText & "', '" & año & "', '', ''"
    arrDatos = Array(sqlDatos)
    Set rsDatos = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsDatos(0).Value = 0 Then
        ConsultaAño = False
    Else
        ConsultaAño = True
    End If
    Call CerrarRecordSet(rsDatos)
End Function

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       cmdEliminarEjercicio_Click
' Description:       Evento que se ejecuta alhacer clic en eliminar ejercicio
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub cmdEliminarEjercicio_Click()
    If MsgBox("Desea eliminar el ejercicio contable " & gsAnio & " de la empresa " & gsEmpresaNom, vbQuestion + vbYesNo) = vbYes Then
        Call DelAnioConta
    End If
End Sub



'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       cmdReplicar_Click
' Description:       Evento que se ejecuta al hacer clic en la replica del ejercicio
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub cmdReplicar_Click()
    ' *** Verificar q registro no existe.
    If ConsultaAño("EXISTEREG", tdbnAnio) = False Then
        Mensajes "El año seleccionado no existe. No se puede Replicar Datos", vbInformation
        Me.cmdReplicar.Enabled = True
        Exit Sub
    End If
    
    ' *** Verificar si se ha registrado movimientos durante.
    If ConsultaAño("EXISTEMOV", tdbnAnio) = True Then
        Mensajes "No se puede Replicar Datos. " & vbCrLf & _
        "El año seleccionado ya tiene asientos registrados.", vbInformation
        Exit Sub
        
    End If
    ' *** Aperturar
    Screen.MousePointer = vbHourglass
    cmdAperturar.Enabled = False
    cmdReplicar.Enabled = False
    
    DoEvents
    Call Replicar
    
    cmdAperturar.Enabled = True
    cmdReplicar.Enabled = True
    
    Screen.MousePointer = vbNormal
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Replicar
' Description:       Procedimiento que solo replica el ejercicio de la empresa
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub Replicar()
    Dim clsMante As clsMantoTablas
    Dim lArrMnt(2) As Variant
    
    Call EscribirLog("Iniciando la replica de datos de la empresa " & gsEmpresa & " " & gsEmpresaNom, Me.Name)
    
    Set clsMante = New clsMantoTablas
    lArrMnt(0) = Me.tdbcEmpresa.BoundText   ' Empresa
    lArrMnt(1) = tdbnAnio                   ' Codigo(año)
    lArrMnt(2) = gsUsuario                  ' Usuario
        
    Dim lArrPlantCO() As Variant        ' *** Arreglo para los mantenimientos
    ReDim lArrPlantCO(3) As Variant
    lArrPlantCO(0) = gsEmpresa           ' Empresa
    lArrPlantCO(1) = gsAnio              ' Anio
    
    If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_UpdPeriodoPlantillaEEFF", lArrPlantCO(), True) = False Then
     Debug.Print "No se actualizo..."
    End If
        
    Set clsMante = New clsMantoTablas
        
    If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_ReplicarAnio", lArrMnt(), True) = False Then
        Mensajes "El proceso no ha concluido. Verificar...", vbInformation
        Exit Sub
    End If
    Mensajes "Los datos del año anterior se replicaron con exito...", vbInformation
    
    Call EscribirLog("Finalizo la replica de datos de la empresa " & gsEmpresa & " " & gsEmpresaNom, Me.Name)
    
    Set clsMante = Nothing
    Exit Sub
ErrorEjecucion:
    Mensajes Str(Err.Number) & Err.Description, vbInformation
    
    Call EscribirLog("Error en la replica de datos del año, [" & Err.Description & "] de la empresa " & gsEmpresa & " " & gsEmpresaNom, Me.Name)
End Sub




'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       DelAnioConta
' Description:       Procedimiento que elimina los asientos temporales eliminados de auditoria
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub DelAnioConta()
    Dim clsMante As clsMantoTablas
    Dim lArrMnt(7) As Variant
    
    Set clsMante = New clsMantoTablas
    lArrMnt(0) = "DEL_ANIO_CONTA"
    lArrMnt(1) = gsEmpresa
    lArrMnt(2) = Null
    lArrMnt(3) = Null
    lArrMnt(4) = Null
    lArrMnt(5) = Null
    lArrMnt(6) = Null
    lArrMnt(7) = gsAnio
    
    If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaEmpresa", lArrMnt(), True) = False Then
        Mensajes "El proceso no ha concluido. Verificar...", vbInformation
        Exit Sub
    End If
    Mensajes "Se borraron los movimientos eliminados con exito (son movimientos que quedan almacenados para auditoria) ...", vbInformation
    
    Set clsMante = Nothing
    Exit Sub
ErrorEjecucion:
    Mensajes Str(Err.Number) & Err.Description, vbInformation
End Sub




'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Form_KeyDown
' Description:       Evento que se ejecuta al presionar una tecla en el formulario
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 And Shift = 1 Then
        If InputBox("Ingrese codigo", "Activar eliminar ejercicio contable") = "977611" Then
            cmdEliminarEjercicio.Visible = True
        End If
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Form_Load
' Description:       Evento que se ejecuta al iniciar el formulario
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub Form_Load()
    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))

    Dim sqlCadena As String
    Centrar_form Me
    cmdEliminarEjercicio.Visible = False
    
    sqlCadena = "SELECT EMP_CCODIGO, EMP_CNOMBRELARGO FROM EMPRESA " & _
                " ORDER BY EMP_CCODIGO"
    LlenarComboAddItem tdbcEmpresa, sqlCadena
    
    If Mid(gsGrupo, 3, 1) <> "1" And gsGrupo <> gsPrivilegioAdmin Then
        Me.cmdAperturar.Enabled = False
        Me.cmdReplicar.Enabled = False
    Else
        Me.cmdAperturar.Enabled = True
        Me.cmdReplicar.Enabled = True
    End If
End Sub

Private Sub Form_Resize()
On Error GoTo errHand
    If Me.WindowState <> vbMinimized Then
        Call SeteaFondoForm(Me)
        Call Centrar_Objeto(fratodo, Me)
        Call CentrarTitulo(lblTitulo, fratodo, Me)
    End If
Exit Sub
errHand:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))

End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbcEmpresa_KeyDown
' Description:       Evento que se ejecuta al presionar una tecla en la lista de empresas
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbcEmpresa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then pSendKeys "{Tab}"
End Sub
