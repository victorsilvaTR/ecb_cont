VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmManRutaReportes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ruta de Reportes Contabilidad"
   ClientHeight    =   2385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6015
   Icon            =   "frmManRutaReportes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2385
   ScaleWidth      =   6015
   Begin VB.Frame Frame2 
      Height          =   1905
      Left            =   135
      TabIndex        =   0
      Top             =   210
      Width           =   5715
      Begin VB.CommandButton cmdSeleccionar 
         Caption         =   "..."
         Height          =   345
         Left            =   2355
         TabIndex        =   1
         Top             =   315
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
         Caption         =   "frmManRutaReportes.frx":1982
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmManRutaReportes.frx":19EE
         Key             =   "frmManRutaReportes.frx":1A0C
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
         Format          =   "@"
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
      Begin MSForms.CommandButton cmdSalir 
         Height          =   390
         Left            =   2970
         TabIndex        =   5
         Top             =   1350
         Width           =   1665
         Caption         =   " Salir"
         PicturePosition =   327683
         Size            =   "2937;688"
         Picture         =   "frmManRutaReportes.frx":1A50
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdGrabar 
         Height          =   390
         Left            =   1125
         TabIndex        =   4
         Top             =   1350
         Width           =   1665
         Caption         =   " Grabar"
         PicturePosition =   327683
         Size            =   "2937;688"
         Picture         =   "frmManRutaReportes.frx":1FEA
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "RUTA DE REPORTES:"
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
         Width           =   1860
      End
   End
   Begin MSComDlg.CommonDialog dlgAbrirArchivo 
      Left            =   5535
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmManRutaReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gsGrupo As String
Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

Private Sub cmdGrabar_Click()
    Dim respuesta As String
    respuesta = MsgBox("Esta seguro de cambiar la ruta de los reportes", vbYesNo + vbQuestion, "Confirmar Cambiar Ruta Reportes")
    If respuesta = vbYes Then
        ' *** Graba la ruta del reporte
        Dim clsMante As clsMantoTablas
        If validarDatos = False Then Exit Sub
        Set clsMante = New clsMantoTablas
        ' *** Grabando Centro de Costo
        'On Local Error GoTo ErrorEjecucion
        ReDim lArrMnt(3) As Variant
        lArrMnt(0) = "INSERTAR"     ' Accion
        lArrMnt(1) = "CONTABILIDAD" ' Sistema
        lArrMnt(2) = tdbtArchivo    ' Año de Trabajo
        lArrMnt(3) = gsUsuario      ' Codigo
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spGrabaRutaReportes", lArrMnt(), True) = False Then
            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
            Exit Sub
        End If
        Mensajes "La ruta de los reporte se han cambiado con exito.", vbInformation
        ' ***
    End If
End Sub

Private Function validarDatos() As Boolean
    validarDatos = False
    ' *** Validar q los datos necesarios esten ingresados
    If TextoLleno(Me.tdbtArchivo, "Ruta de Reporte") = False Then Exit Function
    validarDatos = True
End Function

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdSeleccionar_Click()
    Me.tdbtArchivo = ""
    On Local Error GoTo ErrorEjecucion
    With Me.dlgAbrirArchivo
        .DialogTitle = "Directorio de los Reportes"
        .InitDir = "C:"
        .Filter = "Reportes| *.rpt"
        .CancelError = True
        .ShowOpen
        If .FileName <> "" Then
            tdbtArchivo = ExtraeCarpeta(.FileName)
        End If
    End With
    Exit Sub
ErrorEjecucion:
    If Err.Number <> 32755 Then Mensajes Str(Err.Number) & Err.Description, vbInformation
End Sub

Private Sub Form_Load()
    Me.Top = (frmMDIConta.ScaleHeight - Me.Height) / 2
    Me.Left = (frmMDIConta.ScaleWidth - Me.Width) / 2
    
    ' *** Cargar la ruta del reporte
    tdbtArchivo = BuscaRutaReportes("CONTABILIDAD")
    ' ***
End Sub

Private Function ExtraeCarpeta(cadena As String)
    Dim i As Integer
    ExtraeCarpeta = ""
    For i = Len(Trim(cadena)) To 1 Step -1
        If Mid(cadena, i, 1) = "\" Then
            ExtraeCarpeta = Mid(cadena, 1, i)
            Exit For
        End If
    Next
End Function

