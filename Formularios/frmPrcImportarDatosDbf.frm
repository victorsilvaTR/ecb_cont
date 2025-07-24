VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmPrcImportarDatosXLS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importar datos de Plantilla XLS"
   ClientHeight    =   6732
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   6000
   Icon            =   "frmPrcImportarDatosDbf.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6732
   ScaleWidth      =   6000
   Begin VB.Frame fraTodo 
      Height          =   6675
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   5955
      Begin VB.Frame Frame1 
         Caption         =   "Tipo Importacion"
         Height          =   1215
         Left            =   3300
         TabIndex        =   23
         Top             =   3720
         Width           =   2295
         Begin VB.OptionButton optImportarArchivoTexto 
            Caption         =   "Importacion Archivo Texto"
            Height          =   375
            Left            =   240
            TabIndex        =   25
            Top             =   720
            Width           =   1935
         End
         Begin VB.OptionButton optImportarExcel 
            Caption         =   "Importacion Excel"
            Height          =   375
            Left            =   240
            TabIndex        =   24
            Top             =   360
            Value           =   -1  'True
            Width           =   1815
         End
      End
      Begin VB.CheckBox chkOpcion 
         Caption         =   "DIARIO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   360
         TabIndex        =   18
         Top             =   4680
         Width           =   2760
      End
      Begin VB.CheckBox chkOpcion 
         Caption         =   "CAJA EGRESO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   360
         TabIndex        =   10
         Top             =   4410
         Width           =   2760
      End
      Begin VB.CheckBox chkOpcion 
         Caption         =   "CAJA INGRESO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   360
         TabIndex        =   9
         Top             =   4140
         Width           =   2760
      End
      Begin VB.CheckBox chkOpcion 
         Caption         =   "REGISTRO DE VENTAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   360
         TabIndex        =   8
         Top             =   3870
         Width           =   2760
      End
      Begin VB.CheckBox chkOpcion 
         Caption         =   "REGISTRO DE COMPRAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   360
         TabIndex        =   7
         Top             =   3600
         Width           =   2760
      End
      Begin VB.CheckBox chkOpcion 
         Caption         =   "LIBRO DE APERTURA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   360
         TabIndex        =   6
         Top             =   3330
         Width           =   2760
      End
      Begin VB.CheckBox chkOpcion 
         Caption         =   "TIPO DE CAMBIO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   360
         TabIndex        =   5
         Top             =   3060
         Width           =   2760
      End
      Begin VB.CheckBox chkOpcion 
         Caption         =   "ENTIDADES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   2790
         Width           =   2760
      End
      Begin TDBText6Ctl.TDBText tdbtArchivo 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   4830
         _Version        =   65536
         _ExtentX        =   8520
         _ExtentY        =   661
         Caption         =   "frmPrcImportarDatosDbf.frx":0ECA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPrcImportarDatosDbf.frx":0F36
         Key             =   "frmPrcImportarDatosDbf.frx":0F54
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   0
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
      Begin MSComctlLib.ProgressBar pbAvance 
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   5760
         Width           =   5625
         _ExtentX        =   9927
         _ExtentY        =   339
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin TDBText6Ctl.TDBText lblCorrelativo 
         Height          =   375
         Left            =   180
         TabIndex        =   0
         Top             =   540
         Width           =   4830
         _Version        =   65536
         _ExtentX        =   8520
         _ExtentY        =   661
         Caption         =   "frmPrcImportarDatosDbf.frx":0F98
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPrcImportarDatosDbf.frx":1004
         Key             =   "frmPrcImportarDatosDbf.frx":1022
         BackColor       =   14737632
         EditMode        =   0
         ForeColor       =   16711680
         ReadOnly        =   1
         ShowContextMenu =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MarginBottom    =   1
         Enabled         =   -1
         MousePointer    =   0
         Appearance      =   0
         BorderStyle     =   1
         AlignHorizontal =   2
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
      Begin TDBText6Ctl.TDBText tdbtArchivodet 
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   1800
         Visible         =   0   'False
         Width           =   4830
         _Version        =   65536
         _ExtentX        =   8520
         _ExtentY        =   661
         Caption         =   "frmPrcImportarDatosDbf.frx":1066
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPrcImportarDatosDbf.frx":10D2
         Key             =   "frmPrcImportarDatosDbf.frx":10F0
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   0
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
      Begin MSForms.CommandButton cmdSeleccionard 
         Height          =   390
         Left            =   5040
         TabIndex        =   27
         Top             =   1800
         Visible         =   0   'False
         Width           =   690
         Caption         =   "DET"
         PicturePosition =   262148
         Size            =   "1217;688"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmd_salir 
         Height          =   435
         Left            =   3855
         TabIndex        =   21
         Top             =   6075
         Width           =   1665
         Caption         =   " Salir"
         PicturePosition =   327683
         Size            =   "2937;767"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdImportarDatos 
         Height          =   435
         Left            =   315
         TabIndex        =   20
         Top             =   6060
         Width           =   1665
         Caption         =   " Importar Datos"
         PicturePosition =   327683
         Size            =   "2937;767"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdImprimir 
         Height          =   435
         Left            =   2100
         TabIndex        =   19
         Top             =   6075
         Width           =   1665
         Caption         =   " Imprimir Errores"
         PicturePosition =   327683
         Size            =   "2937;767"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   390
         Left            =   5040
         TabIndex        =   1
         ToolTipText     =   " Vuelve a cargar los datos almacenados "
         Top             =   540
         Width           =   450
         PicturePosition =   262148
         Size            =   "794;688"
         Picture         =   "frmPrcImportarDatosDbf.frx":1134
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CORRELATIVO:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   3
         Left            =   180
         TabIndex        =   17
         Top             =   225
         Width           =   1275
      End
      Begin MSForms.CommandButton cmdSeleccionar 
         Height          =   390
         Left            =   5040
         TabIndex        =   3
         Top             =   1320
         Width           =   690
         Caption         =   "EXCEL"
         PicturePosition =   262148
         Size            =   "1217;688"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PROCESO:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   5160
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "SELECCIONE LOS DATOS A IMPORTAR:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   2400
         Width           =   3300
      End
      Begin VB.Label lblAvance 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         TabIndex        =   14
         Top             =   5460
         Width           =   5325
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "SELECCIONE ARCHIVO:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   0
         Left            =   195
         TabIndex        =   12
         Top             =   945
         Width           =   1950
      End
   End
   Begin MSComDlg.CommonDialog dlgAbrirArchivo 
      Left            =   -75
      Top             =   4455
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblTitulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BALANCE DE COMPROBACION"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   4365
   End
End
Attribute VB_Name = "frmPrcImportarDatosXLS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gsGrupo As String
Dim PeriodoXls As String
Dim rsUpdEstado As Recordset
Dim lblnValor, lblnExistenciaTemporal As Boolean
Dim lstrSqlC, lstrSqlCDetalle As String
    
Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

Private Sub cmd_salir_Click()
Unload Me
End Sub

Private Sub cambiarRutaArchivo()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.MoveFile Me.tdbtArchivo.Text, "C:\"
    Set fso = Nothing
    ' ***
End Sub

Function GetExistenciaTemporal() As Boolean
    
    Dim ObjFunciones As ClsFuncionesExecute
    Dim strSql As String
    Dim rsTemporal As ADODB.Recordset
    
    GetExistenciaTemporal = False
    lblnExistenciaTemporal = False
    
    Set rsTemporal = New ADODB.Recordset
    Set ObjFunciones = New ClsFuncionesExecute
    
    strSql = "SPCN_Validacion_Temporales"
    Set rsTemporal = ObjFunciones.fRetornaRS(strSql)
    
    If rsTemporal("Campo").Value > 0 Then
        If MsgBox("¿Desea continuar modificando la Informacion anterior?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
            Call DeleteTablasTemporales
        Else
            lblnExistenciaTemporal = True
            GetExistenciaTemporal = True
            
            Dim FrmValidacioImportacion As FrmValidacionImportacionOtroSistema
            Set FrmValidacioImportacion = New FrmValidacionImportacionOtroSistema
            
            FrmValidacioImportacion.pstrEntidad = IIf(rsTemporal("Campo").Value = "1", "1", "0")
            FrmValidacioImportacion.pstrTipoCambio = IIf(rsTemporal("Campo").Value = "2", "2", "0")
            FrmValidacioImportacion.pstrApertura = IIf(rsTemporal("Campo").Value = "3", "3", "0")
            FrmValidacioImportacion.pstrCompra = IIf(rsTemporal("Campo").Value = "4", "4", "0")
            FrmValidacioImportacion.pstrVenta = IIf(rsTemporal("Campo").Value = "5", "5", "0")
            FrmValidacioImportacion.pstrIngreso = IIf(rsTemporal("Campo").Value = "6", "6", "0")
            FrmValidacioImportacion.pstrEgreso = IIf(rsTemporal("Campo").Value = "7", "7", "0")
            FrmValidacioImportacion.pstrDiario = IIf(rsTemporal("Campo").Value = "8", "8", "0")
            Unload Me
            FrmValidacioImportacion.Show
        End If
    End If

End Function

Private Sub DeleteTablasTemporales()
    
    Dim cn As New ADODB.Connection
    cn.Open gsCadenaConexion
    
    Call cn.Execute("if exists (select 1 from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_ENTIDAD]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & _
                    "drop table [dbo].[ZIMP_ENTIDAD]")
    Call cn.Execute("if exists (select 1 from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_TIPOCAMBIO]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & _
                    "drop table [dbo].[ZIMP_TIPOCAMBIO]")
    Call cn.Execute("if exists (select 1 from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_APE_CAB]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & _
                    "drop table [dbo].[ZIMP_APE_CAB]")
    Call cn.Execute("if exists (select 1 from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_APE_DET]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & _
                    "drop table [dbo].[ZIMP_APE_DET]")
    Call cn.Execute("if exists (select 1 from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_COMPRAS_CAB]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & _
                    "drop table [dbo].[ZIMP_COMPRAS_CAB]")
    Call cn.Execute("if exists (select 1 from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_COMPRAS_DET]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & _
                    "drop table [dbo].[ZIMP_COMPRAS_DET]")
    Call cn.Execute("if exists (select 1 from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_VENTAS_CAB]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & _
                    "drop table [dbo].[ZIMP_VENTAS_CAB]")
    Call cn.Execute("if exists (select 1 from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_VENTAS_DET]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & _
                    "drop table [dbo].[ZIMP_VENTAS_DET]")
    Call cn.Execute("if exists (select 1 from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_CAJAING_CAB]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & _
                    "drop table [dbo].[ZIMP_CAJAING_CAB]")
    Call cn.Execute("if exists (select 1 from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_CAJAING_DET]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & _
                    "drop table [dbo].[ZIMP_CAJAING_DET]")
    Call cn.Execute("if exists (select 1 from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_CAJAEGR_CAB]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & _
                    "drop table [dbo].[ZIMP_CAJAEGR_CAB]")
    Call cn.Execute("if exists (select 1 from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_CAJAEGR_DET]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & _
                    "drop table [dbo].[ZIMP_CAJAEGR_DET]")
    Call cn.Execute("if exists (select 1 from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_PLAN_CAB]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & _
                    "drop table [dbo].[ZIMP_PLAN_CAB]")
    Call cn.Execute("if exists (select 1 from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_PLAN_DET]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & _
                    "drop table [dbo].[ZIMP_PLAN_DET]")
    
End Sub

Private Sub SaveGeneral()
        
    Dim cn As New ADODB.Connection
    cn.Open gsCadenaConexion
    cn.CommandTimeout = 900000000
    Dim intContador As Integer
    Dim strRuta, strSqlD, strLinea, strCaracter, strDato As String
    Dim intLongitudCadena, intInicio As Integer
    strRuta = Me.tdbtArchivo & "\" & IIf(Me.chkOpcion(0).Value = 1, "Entidad.txt", IIf(Me.chkOpcion(1).Value = 1, "TipoCambio.txt", ""))
    strCaracter = "|"
    intContador = 1
    Open strRuta For Input As #1
    Do Until EOF(1)
        intInicio = 0
        Line Input #1, strLinea
        intLongitudCadena = Len(strLinea)
        strSqlD = vbNullString
        Call GetTabla
        intContador = 1
        Do While intLongitudCadena > 0
            strDato = vbNullString
            intInicio = InStr(strLinea, strCaracter)
            strDato = Replace(Left(strLinea, intInicio), "|", "")
            strLinea = Right(strLinea, Len(strLinea) - intInicio)
            intLongitudCadena = intLongitudCadena - intInicio
            
            If Left(strDato, 1) = "0" Then
                strSqlD = strSqlD & "'" & strDato & "', "
            Else
                If Me.chkOpcion(0).Value = 1 Then
                    If intContador = 9 Or intContador = 10 Or intContador = 11 Then
                        strDato = IIf(strDato = vbNullString, "Null", "'" & strDato & "'") & ","
                        strSqlD = strSqlD & strDato
                    Else
                        strSqlD = strSqlD & IIf(IsNumeric(strDato), strDato, "'" & strDato & "'") & ","
                    End If
                Else
                    strSqlD = strSqlD & IIf(IsNumeric(strDato), strDato, "'" & strDato & "'") & ","
                End If
                
            End If
            intContador = intContador + 1
'            strSqlD = strSqlD & IIf(IsNumeric(strDato), strDato, "'" & strDato & "'") & ","
        Loop
        strSqlD = Left(strSqlD, Len(strSqlD) - 1) & ")"
        lstrSqlC = lstrSqlC & strSqlD
        Call cn.Execute(lstrSqlC)
    Loop
    Close #1
    
End Sub

Private Sub SaveCabecera()
    
    Dim cn As New ADODB.Connection
    cn.Open gsCadenaConexion
    cn.CommandTimeout = 90000000
    Dim strRuta, strSqlD, strLinea, strCaracter, strDato As String
    Dim intLongitudCadena, intInicio As Integer
    
    If Me.chkOpcion(2).Value = 1 Then
        strRuta = Me.tdbtArchivo.Text & "\Apertura.txt"
    ElseIf Me.chkOpcion(3).Value = 1 Then
        strRuta = Me.tdbtArchivo.Text & "\Compra.txt"
    ElseIf Me.chkOpcion(4).Value = 1 Then
        strRuta = Me.tdbtArchivo.Text & "\Venta.txt"
    ElseIf Me.chkOpcion(5).Value = 1 Then
        strRuta = Me.tdbtArchivo.Text & "\Ingreso.txt"
    ElseIf Me.chkOpcion(6).Value = 1 Then
        strRuta = Me.tdbtArchivo.Text & "\Egreso.txt"
    ElseIf Me.chkOpcion(7).Value = 1 Then
        strRuta = Me.tdbtArchivo.Text & "\Diario.txt"
    End If
    
    
    strCaracter = "|"
    Open strRuta For Input As #1
    Do Until EOF(1)
        intInicio = 0
        Line Input #1, strLinea
        intLongitudCadena = Len(strLinea)
        strSqlD = vbNullString
        Call GetTabla
        Do While intLongitudCadena > 0
            strDato = vbNullString
            intInicio = InStr(strLinea, strCaracter)
            strDato = Replace(Left(strLinea, intInicio), "|", "")
            strLinea = Right(strLinea, Len(strLinea) - intInicio)
            intLongitudCadena = intLongitudCadena - intInicio
                        
            If Left(strDato, 1) = "0" Then
                strSqlD = strSqlD & "'" & strDato & "',"
            Else
                strSqlD = strSqlD & IIf(IsNumeric(strDato), strDato, "'" & strDato & "'") & ","
            End If
            
        Loop
        
        If strSqlD <> vbNullString Then
            strSqlD = Left(strSqlD, Len(strSqlD) - 1) & ")"
            lstrSqlC = lstrSqlC & strSqlD
            Call cn.Execute(lstrSqlC)
        End If
    Loop
    Close #1
    
End Sub

Private Sub SaveDetalle()
    
    Dim cn As New ADODB.Connection
    cn.Open gsCadenaConexion
    Dim strRuta, strSqlD, strLinea, strCaracter, strDato As String
    Dim intLongitudCadena, intInicio As Integer
    
    If Me.chkOpcion(2).Value = 1 Then
        strRuta = Me.tdbtArchivo.Text & "\DetalleApertura.txt"
    ElseIf Me.chkOpcion(3).Value = 1 Then
        strRuta = Me.tdbtArchivo.Text & "\DetalleCompra.txt"
    ElseIf Me.chkOpcion(4).Value = 1 Then
        strRuta = Me.tdbtArchivo.Text & "\DetalleVenta.txt"
    ElseIf Me.chkOpcion(5).Value = 1 Then
        strRuta = Me.tdbtArchivo.Text & "\DetalleIngreso.txt"
    ElseIf Me.chkOpcion(6).Value = 1 Then
        strRuta = Me.tdbtArchivo.Text & "\DetalleEgreso.txt"
    ElseIf Me.chkOpcion(7).Value = 1 Then
        strRuta = Me.tdbtArchivo.Text & "\DetalleDiario.txt"
    End If
    
    Dim intContador As Integer
    
    strCaracter = "|"
    Open strRuta For Input As #1
    Do Until EOF(1)
        intInicio = 0
        Line Input #1, strLinea
        intLongitudCadena = Len(strLinea)
        strSqlD = vbNullString
        intContador = 1
        Call GetTabla
        Do While intLongitudCadena > 0
            strDato = vbNullString
            intInicio = InStr(strLinea, strCaracter)
            strDato = Replace(Left(strLinea, intInicio), "|", "")
            strLinea = Right(strLinea, Len(strLinea) - intInicio)
            intLongitudCadena = intLongitudCadena - intInicio
            
            If Left(strDato, 1) = "0" Then
                strSqlD = strSqlD & "'" & strDato & "',"
            Else
                If intContador = 9 Or intContador = 10 Or intContador = 11 Or intContador = 12 Or intContador = 13 Then
                    If strDato = vbNullString Then
                        strSqlD = strSqlD & "0,"
                    Else
                        strSqlD = strSqlD & strDato & ","
                    End If
                Else
                    If Me.chkOpcion(3).Value = 1 Or Me.chkOpcion(4).Value = 1 Then
                        If intContador = 35 Or intContador = 36 Or intContador = 37 Or intContador = 38 Or intContador = 39 Then
                            strSqlD = strSqlD & IIf(strDato = vbNullString, "NULL", "'" & strDato & "',")
                        Else
                            strSqlD = strSqlD & IIf(IsNumeric(strDato), strDato, "'" & strDato & "'") & ","
                        End If
                    Else
                        strSqlD = strSqlD & IIf(IsNumeric(strDato), strDato, "'" & strDato & "'") & ","
                    End If
                    
                End If
                
            End If
            intContador = intContador + 1
'            strSqlD = strSqlD & IIf(IsNumeric(strDato), strDato, "'" & strDato & "'") & ","
        Loop
        
        If strSqlD <> vbNullString Then
            strSqlD = Left(strSqlD, Len(strSqlD) - 1) & ")"
            lstrSqlCDetalle = lstrSqlCDetalle & strSqlD
            Call cn.Execute(lstrSqlCDetalle)
        End If
    Loop
    Close #1
    
End Sub

Private Function GetValidacionArchivo() As Boolean
    
    Dim strExistencia, strRutaAuxiliar As String
    If Me.chkOpcion(0).Value = 1 Then
        strRuta = Me.tdbtArchivo.Text & "\Entidad.txt"
    ElseIf Me.chkOpcion(1).Value = 1 Then
        strRuta = Me.tdbtArchivo.Text & "\TipoCambio.txt"
    ElseIf Me.chkOpcion(2).Value = 1 Then
        strRuta = Me.tdbtArchivo.Text & "\Apertura.txt"
        strRutaAuxiliar = Me.tdbtArchivo.Text & "\DetalleApertura.txt"
    ElseIf Me.chkOpcion(3).Value = 1 Then
        strRuta = Me.tdbtArchivo.Text & "\Compra.txt"
        strRutaAuxiliar = Me.tdbtArchivo.Text & "\DetalleCompra.txt"
    ElseIf Me.chkOpcion(4).Value = 1 Then
        strRuta = Me.tdbtArchivo.Text & "\Venta.txt"
        strRutaAuxiliar = Me.tdbtArchivo.Text & "\DetalleVenta.txt"
    ElseIf Me.chkOpcion(5).Value = 1 Then
        strRuta = Me.tdbtArchivo.Text & "\Ingreso.txt"
        strRutaAuxiliar = Me.tdbtArchivo.Text & "\DetalleIngreso.txt"
    ElseIf Me.chkOpcion(6).Value = 1 Then
        strRuta = Me.tdbtArchivo.Text & "\Egreso.txt"
        strRutaAuxiliar = Me.tdbtArchivo.Text & "\DetalleEgreso.txt"
    ElseIf Me.chkOpcion(7).Value = 1 Then
        strRuta = Me.tdbtArchivo.Text & "\Diario.txt"
        strRutaAuxiliar = Me.tdbtArchivo.Text & "\DetalleDiario.txt"
    End If
    
    If Me.chkOpcion(2).Value = 1 Or Me.chkOpcion(3).Value = 1 Or Me.chkOpcion(4).Value = 1 Or Me.chkOpcion(5).Value = 1 Or Me.chkOpcion(6).Value = 1 Or Me.chkOpcion(7).Value = 1 Then
        strExistencia = Dir$(strRutaAuxiliar)
        If (Len(strExistencia) = 0) Then
            GetValidacionArchivo = False
            MsgBox "No se encontro el Archivo " & strRuta, vbCritical, "Sistema ECB-Cont"
            Exit Function
        Else
            GetValidacionArchivo = True
        End If
    End If
    
    strExistencia = Dir$(strRuta)
    If (Len(strExistencia) = 0) Then
        GetValidacionArchivo = False
        MsgBox "No se encontro el Archivo " & strRuta, vbCritical, "Sistema ECB-Cont"
        Exit Function
    Else
        GetValidacionArchivo = True
    End If
    
End Function

Private Sub GetTabla()

    lstrSqlC = vbNullString
    lstrSqlCDetalle = vbNullString
    
    If Me.chkOpcion(0).Value = 1 Then
        lstrSqlC = "Insert Into ZIMP_ENTIDAD(Ent_cCodEntidad, Ten_cTipoEntidad, Ent_cPersona, Ent_cDireccion, " & _
                   "Ent_nRuc, Ent_cTipoDoc, Ent_cFlagPersona, Ent_cFlagDomiciliado, Id_Pais, Id_Vinculo_Economico, Id_Convenio, PorcentajeSunat) Values ("
    End If
    If Me.chkOpcion(1).Value = 1 Then
        lstrSqlC = "Insert Into ZIMP_TIPOCAMBIO(Tca_dFecha, Tca_cCodigoOrigen, Tca_cCodigoDestino, Tca_nCompra, " & _
                    "Tca_nVenta, Tca_nVentaP) Values ("
    End If
    If Me.chkOpcion(2).Value = 1 Then
        lstrSqlC = "Insert Into ZIMP_APE_CAB(Ase_cNummov, Pan_cAnio, Per_cPeriodo, " & _
                    "Lib_cTipoLibro, Ase_nVoucher, Ase_dFecha, Ase_cGlosa, Ase_cTipoMoneda, CreditoFiscal, MaterialConstruccion) Values ("
                    
        lstrSqlCDetalle = "Insert Into ZIMP_APE_DET(Ase_cNummov, Pan_cAnio, Per_cPeriodo, " & _
                          "Lib_cTipoLibro, Ase_nVoucher, Pla_cCuentaContable, Asd_nItem, Asd_cGlosa, Asd_nDebeSoles, Asd_nHaberSoles, Asd_nTipoCambio, " & _
                          "Asd_nDebeMonExt, Asd_nHaberMonExt, Cos_cCodigo, Ten_cTipoEntidad, Ent_cCodEntidad, Asd_cTipoDoc, Asd_dFecDoc, Asd_cSerieDoc, Asd_cNumDoc, " & _
                          "Asd_dFecVen, Asd_cProvCanc, Asd_cOperaTC, Asd_cTipoMoneda) Values ("
        
    End If
    If Me.chkOpcion(3).Value = 1 Then
        lstrSqlC = "Insert Into ZIMP_COMPRAS_CAB(Ase_cNummov, Pan_cAnio, Per_cPeriodo, " & _
                   "Lib_cTipoLibro, Ase_nVoucher, Ase_dFecha, Ase_cGlosa, Ase_cTipoMoneda, CreditoFiscal, MaterialConstruccion) Values ("
                   
        lstrSqlCDetalle = "Insert Into ZIMP_COMPRAS_DET(Ase_cNummov, Pan_cAnio, Per_cPeriodo, "
        lstrSqlCDetalle = lstrSqlCDetalle & "Lib_cTipoLibro, Ase_nVoucher, Pla_cCuentaContable, Asd_nItem, Asd_cGlosa, Asd_nDebeSoles, Asd_nHaberSoles, Asd_nTipoCambio, "
        lstrSqlCDetalle = lstrSqlCDetalle & "Asd_nDebeMonExt, Asd_nHaberMonExt, Cos_cCodigo, Ten_cTipoEntidad, Ent_cCodEntidad, Asd_cTipoDoc, Asd_dFecDoc, Asd_cSerieDoc, Asd_cNumDoc, "
        lstrSqlCDetalle = lstrSqlCDetalle & "Asd_dFecVen, Asd_cTipoDocRef, Asd_dFecDocRef, Asd_cSerieDocRef, Asd_cNumDocRef, Asd_nMontoInafecto, Asd_cBaseImp, Asd_cRetencion, Asd_dFechaSpot, "
        lstrSqlCDetalle = lstrSqlCDetalle & "Asd_cNumSpot, Asd_cProvCanc, Asd_cOperaTC, Asd_cTipoMoneda, Asd_cComprobante, Id_Exoneracion, Id_Tipo_Renta, Id_Modalidad, Id_Aduana, Id_Clasific_Servicio) Values ("
    End If
    If Me.chkOpcion(4).Value = 1 Then
        lstrSqlC = "Insert Into ZIMP_VENTAS_CAB(Ase_cNummov, Pan_cAnio, Per_cPeriodo, " & _
                    "Lib_cTipoLibro, Ase_nVoucher, Ase_dFecha, Ase_cGlosa, Ase_cTipoMoneda, Id_Exoneracion, Id_Tipo_Renta, Id_Modalidad, CreditoFiscal, MaterialConstruccion) Values ("
                    
        lstrSqlCDetalle = "Insert Into ZIMP_VENTAS_DET(Ase_cNummov, Pan_cAnio, Per_cPeriodo, "
        lstrSqlCDetalle = lstrSqlCDetalle & "Lib_cTipoLibro, Ase_nVoucher, Pla_cCuentaContable, Asd_nItem, Asd_cGlosa, Asd_nDebeSoles, Asd_nHaberSoles, Asd_nTipoCambio, "
        lstrSqlCDetalle = lstrSqlCDetalle & "Asd_nDebeMonExt, Asd_nHaberMonExt, Cos_cCodigo, Ten_cTipoEntidad, Ent_cCodEntidad, Asd_cTipoDoc, Asd_dFecDoc, Asd_cSerieDoc, Asd_cNumDoc, "
        lstrSqlCDetalle = lstrSqlCDetalle & "Asd_dFecVen, Asd_nMontoInafecto, Asd_cBaseImp, Asd_cProvCanc, Asd_cOperaTC, Asd_cTipoMoneda, Asd_cTipoDocRef, Asd_dFecDocRef, Asd_cSerieDocRef, Asd_cNumDocRef, Id_Exoneracion, Id_Tipo_Renta, Id_Modalidad, Id_Aduana, Id_Clasific_Servicio) Values ("
    End If
    If Me.chkOpcion(5).Value = 1 Then
        lstrSqlC = "Insert Into ZIMP_CAJAING_CAB(Ase_cNummov, Pan_cAnio, Per_cPeriodo, " & _
                    "Lib_cTipoLibro, Ase_nVoucher, Ase_dFecha, Ase_cGlosa, Ase_cTipoMoneda, CreditoFiscal, MaterialConstruccion) Values ("
        
        lstrSqlCDetalle = "Insert Into ZIMP_CAJAING_DET(Ase_cNummov, Pan_cAnio, Per_cPeriodo, "
        lstrSqlCDetalle = lstrSqlCDetalle & "Lib_cTipoLibro, Ase_nVoucher, Pla_cCuentaContable, Asd_nItem, Asd_cGlosa, Asd_nDebeSoles, Asd_nHaberSoles, Asd_nTipoCambio, "
        lstrSqlCDetalle = lstrSqlCDetalle & "Asd_nDebeMonExt, Asd_nHaberMonExt, Cos_cCodigo, Ten_cTipoEntidad, Ent_cCodEntidad, Asd_cTipoDoc, Asd_dFecDoc, Asd_cSerieDoc, Asd_cNumDoc, "
        lstrSqlCDetalle = lstrSqlCDetalle & "Asd_dFecVen, Asd_cTipoDocRef, Asd_dFecDocRef, Asd_cSerieDocRef, Asd_cNumDocRef, Asd_cRetencion, Asd_cProvCanc, Asd_cOperaTC, "
        lstrSqlCDetalle = lstrSqlCDetalle & "Asd_cTipoMoneda, Tra_cCodigo, Asd_cFormaPago) Values ("
    End If
    If Me.chkOpcion(6).Value = 1 Then
        lstrSqlC = "Insert Into ZIMP_CAJAEGR_CAB(Ase_cNummov, Pan_cAnio, Per_cPeriodo, " & _
                   "Lib_cTipoLibro, Ase_nVoucher, Ase_dFecha, Ase_cGlosa, Ase_cTipoMoneda, CreditoFiscal, MaterialConstruccion) Values ("
                   
        lstrSqlCDetalle = "Insert Into ZIMP_CAJAEGR_DET(Ase_cNummov, Pan_cAnio, Per_cPeriodo, " & _
                          "Lib_cTipoLibro, Ase_nVoucher, Pla_cCuentaContable, Asd_nItem, Asd_cGlosa, Asd_nDebeSoles, Asd_nHaberSoles, Asd_nTipoCambio, " & _
                          "Asd_nDebeMonExt, Asd_nHaberMonExt, Cos_cCodigo, Ten_cTipoEntidad, Ent_cCodEntidad, Asd_cTipoDoc, Asd_dFecDoc, Asd_cSerieDoc, Asd_cNumDoc, " & _
                          "Asd_dFecVen, Asd_cTipoDocRef, Asd_dFecDocRef, Asd_cSerieDocRef, Asd_cNumDocRef, Asd_cRetencion, Asd_cProvCanc, Asd_cOperaTC, " & _
                          "Asd_cTipoMoneda, Tra_cCodigo, Asd_cFormaPago) Values ("
    End If
    If Me.chkOpcion(7).Value = 1 Then
        lstrSqlC = "Insert Into ZIMP_PLAN_CAB(Ase_cNummov, Pan_cAnio, Per_cPeriodo, " & _
                   "Lib_cTipoLibro, Ase_nVoucher, Ase_dFecha, Ase_cGlosa, Ase_cTipoMoneda, CreditoFiscal, MaterialConstruccion) Values ("
                   
        lstrSqlCDetalle = "Insert Into ZIMP_PLAN_DET(Ase_cNummov, Pan_cAnio, Per_cPeriodo, " & _
                          "Lib_cTipoLibro, Ase_nVoucher, Pla_cCuentaContable, Asd_nItem, Asd_cGlosa, Asd_nDebeSoles, " & _
                          "Asd_nHaberSoles, Asd_nTipoCambio, " & _
                          "Asd_nDebeMonExt, Asd_nHaberMonExt, Cos_cCodigo, Ten_cTipoEntidad, Ent_cCodEntidad, " & _
                          "Asd_cTipoDoc, Asd_dFecDoc, Asd_cSerieDoc, " & _
                          "Asd_cNumDoc, Asd_dFecVen, Asd_cProvCanc, Asd_cOperaTC, " & _
                          "Asd_cTipoMoneda) Values ("
    End If
    
End Sub

Private Sub cmdImportarDatos_Click()
On Error GoTo Control
    Dim i, Cont As Integer
    Dim Resultado As Boolean
    Cont = 0
    If optImportarArchivoTexto.Value Then
        If chkOpcion(3) = 1 And chkOpcion(4) = 1 Then
            Mensajes "Debe seleccionar Compras o Ventas"
            chkOpcion(3) = 0
            chkOpcion(4) = 0
            Exit Sub
        ElseIf chkOpcion(3) = 0 And chkOpcion(4) = 0 Then
            Mensajes "Debe seleccionar una casilla, Compras o Ventas"
            Exit Sub
        End If
        If CE(tdbtArchivo.Text) = "" Then Mensajes "Seleccione el archivo TXT Encabezado": Exit Sub
        If CE(tdbtArchivodet.Text) = "" Then Mensajes "Seleccione el archivo TXT Detalle": Exit Sub
            
        Procesar_Archivos_Texto
        Exit Sub
    Else
        For i = 0 To 7
            If chkOpcion(i) = 1 Then
                Cont = Cont + 1
            End If
        Next
        If Cont > 1 Then
            MsgBox ("Debe seleccionar un solo registro a importar")
            For i = 0 To 7
                chkOpcion(i) = 0
            Next
            Exit Sub
        End If
        If CE(tdbtArchivo.Text) = "" Then Mensajes "Seleccione el archivo a importar": Exit Sub
    End If

    Dim gsTipolibgsPeriodoro As String, gsPeriodo As String, gsPeriodoant As String
    Dim gsNroVoucher As String, gsnummovant As String
    Dim FrmValidacioImportacion As FrmValidacionImportacionOtroSistema
    Set FrmValidacioImportacion = New FrmValidacionImportacionOtroSistema

    If GetExistenciaTemporal() Then Exit Sub

    Dim strIdExoneracion, strIdTipoRenta, strIdModalidad, strIdAduana, strIdClasificacion As String
    Dim strDomiciliado, strIdPais, strIdVinculo, strIdConvenio As String
    cmdImportarDatos.Enabled = False
    cmdSeleccionar.Enabled = False
    cmdImprimir.Enabled = False
    cmd_salir.Enabled = False
    Me.MousePointer = vbHourglass
    DoEvents
    
    Dim Ex As Excel.Application
    Dim Wb As Excel.Workbook
    Dim Sht As Excel.Worksheet

    If Me.optImportarExcel.Value Then
        Set Ex = CreateObject("Excel.Application")
        Set Wb = Ex.Workbooks.Open(Me.tdbtArchivo.Text)
    End If
    sBaseDatos = gsBD

    Dim TipoLibroXls As String
    Dim VarContFila As Long
    Dim cn As New ADODB.Connection
    
    cn.Open gsCadenaConexion
    cn.Execute "set dateformat dmy"
    
    If chkOpcion(0).Value = 1 Then 'Entidad
        
        Call cn.Execute("if exists (select 1 from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_ENTIDAD]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & _
        "drop table [dbo].[ZIMP_ENTIDAD]")
        
        sSql = "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_ENTIDAD] (" & _
            "[Ent_cCodEntidad] char (5) NULL, " & "[Ten_cTipoEntidad] char (1) NULL, " & _
            "[Ent_cPersona] nvarchar (255) NULL, " & "[Ent_cDireccion] nvarchar (255) NULL, " & _
            "[Ent_nRuc] nvarchar (255) NULL, " & "[Ent_cTipoDoc] nvarchar (255) NULL, " & _
            "[Ent_cFlagPersona] nvarchar (255) NULL, [Ent_cFlagDomiciliado] char(1) NULL, [Id_Pais] char(10) NULL, " & _
            "[Id_Vinculo_Economico] char(10) NULL, [Id_Convenio] char(10) NULL, [PorcentajeSunat] char(1) NULL)"
        cn.Execute sSql
        
        If Me.optImportarExcel.Value Then
            Set Sht = Wb.Worksheets("ENTIDAD")
            For VarContFila = 2 To 70000
                If Sht.Range("A" & VarContFila).Value = "" Then Exit For
            Next
            lblAvance.Caption = "IMPORTANDO A TABLA TEMPORAL - ENTIDAD"
            Me.pbAvance.Min = 0
            Me.pbAvance.Value = 0
            Me.pbAvance.Max = VarContFila - 1

            For VarContFila = 2 To 70000
                If Sht.Range("A" & VarContFila).Value = "" Then Exit For
                pbAvance.Value = pbAvance.Value + 1
                DoEvents
                
                'Evaluando los nuevos campos LE
                strDomiciliado = IIf(CE(Sht.Range("H" & VarContFila).Value) = vbNullString, "'1'", CE(Sht.Range("H" & VarContFila).Value))
                strIdPais = IIf(CE(Sht.Range("I" & VarContFila).Value) = vbNullString, "Null", "'" & CE(Sht.Range("I" & VarContFila).Value) & "'")
                strIdVinculo = IIf(CE(Sht.Range("J" & VarContFila).Value) = vbNullString, "Null", "'" & CE(Sht.Range("J" & VarContFila).Value) & "'")
                strIdConvenio = IIf(CE(Sht.Range("K" & VarContFila).Value) = vbNullString, "Null", "'" & CE(Sht.Range("K" & VarContFila).Value) & "'")
                
                '*****SI ENCUENTRA UNA SOLA COMILLA EN LA GLOSA*****
                If InStr(Sht.Range("C" & VarContFila).Value, "'") > 0 Then
                    Sht.Range("C" & VarContFila).Value = Replace(Sht.Range("C" & VarContFila).Value, "'", "")
                End If
                If InStr(Sht.Range("D" & VarContFila).Value, "'") > 0 Then
                    Sht.Range("D" & VarContFila).Value = Replace(Sht.Range("D" & VarContFila).Value, "'", "")
                End If

                'Modifique la columna L no viene en el Excel JCS 21/04/2017 OCTALIA.S.A Chile
                sSql = "Insert Into ZIMP_ENTIDAD (" & _
                    "Ent_cCodEntidad, Ten_cTipoEntidad, Ent_cPersona, Ent_cDireccion, " & _
                    "Ent_nRuc, Ent_cTipoDoc, Ent_cFlagPersona, Ent_cFlagDomiciliado, Id_Pais, " & _
                    "Id_Vinculo_Economico, Id_Convenio, PorcentajeSunat) Values ('" & _
                    Sht.Range("A" & VarContFila).Value & "','" & Sht.Range("B" & VarContFila).Value & "','" & _
                    Sht.Range("C" & VarContFila).Value & "','" & Sht.Range("D" & VarContFila).Value & "','" & _
                    Sht.Range("E" & VarContFila).Value & "','" & Sht.Range("F" & VarContFila).Value & "','" & _
                    Sht.Range("G" & VarContFila).Value & "', " & strDomiciliado & ", " & strIdPais & ", " & _
                    strIdVinculo & ", " & strIdConvenio & ",NULL)"
                                    
                cn.BeginTrans
                cn.Execute sSql
                cn.CommitTrans
                
                strDomiciliado = vbNullString: strIdPais = vbNullString: strIdVinculo = vbNullString: strIdConvenio = vbNullString
            Next
        Else
            If Not GetValidacionArchivo Then Exit Sub
            Call SaveGeneral
        End If
    End If

    If chkOpcion(1).Value = 1 Then 'Tipo Cambio
        Call cn.Execute("if exists (select 1 from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_TIPOCAMBIO]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & _
        "drop table [dbo].[ZIMP_TIPOCAMBIO]")

        sSql = "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_TIPOCAMBIO] (" & _
            "[Tca_dFecha] DateTime NULL, " & "[Tca_cCodigoOrigen] nvarchar (255) NULL, " & _
            "[Tca_cCodigoDestino] nvarchar (255) NULL, " & "[Tca_nCompra] nvarchar (255) NULL, " & _
            "[Tca_nVenta] nvarchar (255) NULL, " & "[Tca_nVentaP] nvarchar (255) NULL)"
        cn.Execute sSql

        If Me.optImportarExcel.Value = True Then
            Set Sht = Wb.Worksheets("TIPOCAMBIO")
            For VarContFila = 2 To 70000
                If Sht.Range("A" & VarContFila).Value = "" Then Exit For
            Next

            lblAvance.Caption = "IMPORTANDO A TABLA TEMPORAL - TIPO DE CAMBIO"
            Me.pbAvance.Min = 0
            Me.pbAvance.Value = 0
            Me.pbAvance.Max = VarContFila - 1

            For VarContFila = 2 To 70000
                If Sht.Range("A" & VarContFila).Value = "" Then Exit For
                pbAvance.Value = pbAvance.Value + 1
                DoEvents
                
                sSql = "Insert Into ZIMP_TIPOCAMBIO (" & _
                    "Tca_dFecha, Tca_cCodigoOrigen, Tca_cCodigoDestino, Tca_nCompra, " & _
                    "Tca_nVenta, Tca_nVentaP) Values ('" & _
                    Format(Sht.Range("A" & VarContFila).Value, "dd/mm/yyyy") & "','" & _
                    Format(Sht.Range("B" & VarContFila).Value, "000") & "','" & _
                    Format(Sht.Range("C" & VarContFila).Value, "000") & "','" & _
                    Sht.Range("D" & VarContFila).Value & "','" & _
                    Sht.Range("E" & VarContFila).Value & "','" & _
                    Sht.Range("F" & VarContFila).Value & "')"

                cn.BeginTrans
                cn.Execute sSql
                cn.CommitTrans
            Next
        Else
            If Not GetValidacionArchivo Then Exit Sub
            Call SaveGeneral
        End If
    End If

    If chkOpcion(2).Value = 1 Then 'Apertura
        Call cn.Execute("if exists (select 1 from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_APE_CAB]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & _
        "drop table [dbo].[ZIMP_APE_CAB]")

        sSql = "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_APE_CAB] (" & _
            "[Ase_cNummov] char (10) NULL, " & "[Pan_cAnio] char (4) NULL, " & _
            "[Per_cPeriodo] char (2) NULL, " & "[Lib_cTipoLibro] char (2) NULL, " & _
            "[Ase_nVoucher] char (10) NULL, " & "[Ase_dFecha] nvarchar (255) NULL, " & _
            "[Ase_cGlosa] nvarchar (255) NULL, " & "[Ase_cTipoMoneda] nvarchar (255) NULL, " & _
            "[CreditoFiscal] char (1) NULL, [MaterialConstruccion] char (1) NULL, [PercepcionIncluida] int NULL)"
        cn.Execute sSql

        If Me.optImportarExcel.Value Then
            Set Sht = Wb.Worksheets("APE_CAB")
            For VarContFila = 2 To 70000
                If Sht.Range("A" & VarContFila).Value = "" Then Exit For
            Next

            lblAvance.Caption = "IMPORTANDO A TABLA TEMPORAL - APERTURACAB"
            Me.pbAvance.Min = 0
            Me.pbAvance.Value = 0
            Me.pbAvance.Max = VarContFila - 1

            For VarContFila = 2 To 70000
                If Sht.Range("A" & VarContFila).Value = "" Then Exit For
                pbAvance.Value = pbAvance.Value + 1
                DoEvents

                'Modifique la columna I,J y K no vienen en el Excel JCS 21/04/2017 OCTALIA.S.A Chile
                sSql = "Insert Into ZIMP_APE_CAB(Ase_cNummov, Pan_cAnio, Per_cPeriodo, " & _
                    "Lib_cTipoLibro, Ase_nVoucher, Ase_dFecha, Ase_cGlosa, Ase_cTipoMoneda, CreditoFiscal, MaterialConstruccion, PercepcionIncluida) Values " & _
                    "('" & Sht.Range("A" & VarContFila).Value & "','" & Sht.Range("B" & VarContFila).Value & "','" & Sht.Range("C" & VarContFila).Value & _
                    "','" & Sht.Range("D" & VarContFila).Value & "','" & Format(Val(Sht.Range("E" & VarContFila).Value), "0000000000") & "','" & Sht.Range("F" & VarContFila).Value & _
                    "','" & Sht.Range("G" & VarContFila).Value & "','" & Format(Sht.Range("H" & VarContFila).Value, "000") & "',NULL,NULL,NULL)"

                cn.BeginTrans
                cn.Execute sSql
                cn.CommitTrans
            Next
        ElseIf Me.optImportarArchivoTexto.Value Then
            If Not GetValidacionArchivo Then Exit Sub
            Call SaveCabecera
        End If

        Call cn.Execute("if exists (select 1 from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_APE_DET]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & _
        "drop table [dbo].[ZIMP_APE_DET]")
''hlp20230710
''''        sSql = "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_APE_DET] (" & _
'''''            "[Ase_cNummov] char (10) NULL, " & "[Pan_cAnio] char (4) NULL, " & _
'''''            "[Per_cPeriodo] char (2) NULL, " & "[Lib_cTipoLibro] char (2) NULL, " & _
'''''            "[Ase_nVoucher] char (10) NULL, " & "[Pla_cCuentaContable] varchar (12) NULL, " & _
'''''            "[Asd_nItem] INT NULL, " & "[Asd_cGlosa] nvarchar (255) NULL, " & _
'''''            "[Asd_nDebeSoles] NUMERIC (14,2) NULL , " & "[Asd_nHaberSoles] NUMERIC (14,2) NULL , " & _
'''''            '''"[Asd_nTipoCambio] NUMERIC (14,2) NULL , " & "[Asd_nDebeMonExt] NUMERIC (14,2) NULL , " & _
'''''            "[Asd_nHaberMonExt] NUMERIC (14,2) NULL , " & "[Cos_cCodigo] varchar (12) NULL, " & _
'''''            "[Ten_cTipoEntidad] char (1) NULL, " & "[Ent_cCodEntidad] char (5) NULL, " & _
'''''            "[Asd_cTipoDoc] char (3) NULL, " & "[Asd_dFecDoc] datetime null , " & _
'''''            "[Asd_cSerieDoc] varchar (20) NULL, " & "[Asd_cNumDoc] varchar (25) NULL, " & _
'''''            "[Asd_dFecVen] datetime NULL, " & "[Asd_cProvCanc] char (1) NULL, " & _
'''''            "[Asd_cOperaTC] char (3) NULL, " & "[Asd_cTipoMoneda] char (3) NULL)"
             sSql = "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_APE_DET] (" & _
            "[Ase_cNummov] char (10) NULL, " & "[Pan_cAnio] char (4) NULL, " & _
            "[Per_cPeriodo] char (2) NULL, " & "[Lib_cTipoLibro] char (2) NULL, " & _
            "[Ase_nVoucher] char (10) NULL, " & "[Pla_cCuentaContable] varchar (12) NULL, " & _
            "[Asd_nItem] INT NULL, " & "[Asd_cGlosa] nvarchar (255) NULL, " & _
            "[Asd_nDebeSoles] NUMERIC (14,3) NULL , " & "[Asd_nHaberSoles] NUMERIC (14,3) NULL , " & _
            "[Asd_nTipoCambio] NUMERIC (14,3) NULL , " & "[Asd_nDebeMonExt] NUMERIC (14,3) NULL , " & _
            "[Asd_nHaberMonExt] NUMERIC (14,3) NULL , " & "[Cos_cCodigo] varchar (12) NULL, " & _
            "[Ten_cTipoEntidad] char (1) NULL, " & "[Ent_cCodEntidad] char (5) NULL, " & _
            "[Asd_cTipoDoc] char (3) NULL, " & "[Asd_dFecDoc] datetime null , " & _
            "[Asd_cSerieDoc] varchar (20) NULL, " & "[Asd_cNumDoc] varchar (25) NULL, " & _
            "[Asd_dFecVen] datetime NULL, " & "[Asd_cProvCanc] char (1) NULL, " & _
            "[Asd_cOperaTC] char (3) NULL, " & "[Asd_cTipoMoneda] char (3) NULL)"




        cn.Execute sSql

        If Me.optImportarExcel.Value Then
            Set Sht = Wb.Worksheets("APE_DET")
            For VarContFila = 2 To 70000
                If Sht.Range("A" & VarContFila).Value = "" Then Exit For
            Next
            lblAvance.Caption = "IMPORTANDO A TABLA TEMPORAL - APERTURADET"
            Me.pbAvance.Value = 0
            Me.pbAvance.Max = VarContFila - 1

            For VarContFila = 2 To 70000
                If Sht.Range("A" & VarContFila).Value = "" Then Exit For
                pbAvance.Value = pbAvance.Value + 1
                DoEvents

                sSql = "Insert Into ZIMP_APE_DET(Ase_cNummov, Pan_cAnio, Per_cPeriodo, " & _
                    "Lib_cTipoLibro, Ase_nVoucher, Pla_cCuentaContable, Asd_nItem, Asd_cGlosa, Asd_nDebeSoles, Asd_nHaberSoles, Asd_nTipoCambio, " & _
                    "Asd_nDebeMonExt, Asd_nHaberMonExt, Cos_cCodigo, Ten_cTipoEntidad, Ent_cCodEntidad, Asd_cTipoDoc, Asd_dFecDoc, Asd_cSerieDoc, Asd_cNumDoc, " & _
                    "Asd_dFecVen, Asd_cProvCanc, Asd_cOperaTC, Asd_cTipoMoneda) Values " & _
                    "('" & Sht.Range("A" & VarContFila).Value & "','" & Sht.Range("B" & VarContFila).Value & "','" & Sht.Range("C" & VarContFila).Value & _
                    "','" & Sht.Range("D" & VarContFila).Value & "','" & Format(Val(Sht.Range("E" & VarContFila).Value), "0000000000") & "','" & Sht.Range("F" & VarContFila).Value & _
                    "','" & Sht.Range("G" & VarContFila).Value & "','" & Sht.Range("H" & VarContFila).Value & "', " & Format(Val(Sht.Range("I" & VarContFila).Value), "0.00") & _
                    ", " & Format(Val(Sht.Range("J" & VarContFila).Value), "0.00") & ", " & Format(Val(Sht.Range("K" & VarContFila).Value), "0.00") & ", " & Format(Val(Sht.Range("L" & VarContFila).Value), "0.00") & _
                    ", " & Format(Val(Sht.Range("M" & VarContFila).Value), "0.00") & ", '" & Sht.Range("N" & VarContFila).Value & "', '" & Sht.Range("O" & VarContFila).Value & _
                    "', '" & Sht.Range("P" & VarContFila).Value & "', '" & Sht.Range("Q" & VarContFila).Value & "', '" & Format(Sht.Range("R" & VarContFila).Value, "dd/MM/yyyy") & _
                    "', '" & Sht.Range("S" & VarContFila).Value & "', '" & Sht.Range("T" & VarContFila).Value & "', '" & IIf(Sht.Range("U" & VarContFila).Value = "", "", Format(Sht.Range("U" & VarContFila).Value, "dd/mm/yyyy")) & _
                    "', '" & Sht.Range("V" & VarContFila).Value & "', '" & Sht.Range("W" & VarContFila).Value & "', '" & Sht.Range("X" & VarContFila).Value & "')"

                cn.BeginTrans
                cn.Execute sSql
                cn.CommitTrans
            Next
        ElseIf Me.optImportarArchivoTexto.Value Then
            If Not GetValidacionArchivo Then Exit Sub
            Call SaveDetalle
        End If
    End If

    If chkOpcion(3).Value = 1 Then ' Importar Compras
        If Me.optImportarExcel.Value Then
            'Validacion del ingreso de montos en las dos monedas
            If gintBiMoneda = 1 Then
                lblnValor = True
                Set Sht = Wb.Worksheets("COMPRAS_DET")
                For VarContFila = 2 To 70000
                    If Sht.Range("A" & VarContFila).Value <> "" Then
                        'Debe Soles - Debe Soles Dolar
                        If (CDbl(Sht.Range("E" & VarContFila).Value) > 0 And CDbl(Sht.Range("H" & VarContFila).Value) = 0) Or CDbl(Sht.Range("F" & VarContFila).Value) > 0 And CDbl(Sht.Range("I" & VarContFila).Value) = 0 Then
                            lblnValor = False
                            Exit For
                        End If
                    Else
                        Exit For
                    End If
                Next

                If lblnValor = False Then
                    MsgBox "La funcionalidad Bimoneda esta activada, ingrese los montos para ambas monedas [COMPRAS_DET]", vbCritical, "Mensaje: Sistema"
                    Exit Sub
                End If
            End If
        End If
        
        gsTipolibro = "06"
        
        Call cn.Execute("if exists (select 1 from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_COMPRAS_CAB]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & _
        "drop table [dbo].[ZIMP_COMPRAS_CAB]")

        sSql = "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_COMPRAS_CAB] (" & _
            "[Ase_cNummov] char (10) NULL, [Pan_cAnio] char (4) NULL, " & _
            "[Per_cPeriodo] char (2) NULL, [Lib_cTipoLibro] char (2) NULL, " & _
            "[Ase_nVoucher] char (10) NULL, [Ase_dFecha] DATETIME NULL, " & _
            "[Ase_cGlosa] nvarchar (255) NULL, [Ase_cTipoMoneda] nvarchar (255) NULL, " & _
            "[Ase_cEstadoO] char (1) NULL, [Ase_cEstadoD] char (1) NULL, " & _
            "[Id_Exoneracion] char(10) NULL, [Id_Tipo_Renta] char(10) NULL, [Id_Modalidad] char(10) NULL, " & _
            "[CreditoFiscal] char(1) NULL, [MaterialConstruccion] char(1) NULL, [PercepcionIncluida] int NULL)"
        cn.Execute sSql

        VarContFila = 0
        gsNroVoucher = ""
        gsPeriodo = ""
        If Me.optImportarExcel.Value Then
            Set Sht = Wb.Worksheets("COMPRAS_CAB")
            lblAvance.Caption = "IMPORTANDO A TABLA TEMPORAL - COMPRASCAB"
            Me.pbAvance.Min = 0
            Me.pbAvance.Value = 0
            Me.pbAvance.Max = Sht.Range("A65536").End(xlUp).Row

            For VarContFila = 2 To 70000
                If Sht.Range("A" & VarContFila).Value = "" Then Exit For
                pbAvance.Value = pbAvance.Value + 1
                PeriodoXls = Format(Val(Sht.Range("B" & VarContFila).Value), "00")
                'pbAvance.Refresh
                DoEvents

                '*****SI ENCUENTRA UNA SOLA COMILLA EN LA GLOSA*****
                If InStr(Sht.Range("E" & VarContFila).Value, "'") > 0 Then
                    Sht.Range("E" & VarContFila).Value = Replace(Sht.Range("E" & VarContFila).Value, "'", "")
                End If

                If gsPeriodo = "" Then
                    gsPeriodo = Format(Val(Sht.Range("B" & VarContFila).Value), "00")
                    gsPeriodoant = Format(Val(gsPeriodo), "00")
                    gsNroVoucher = BuscaNvoucher(gsTipolibro, gsPeriodo)
                Else
                    gsPeriodo = Format(Val(Sht.Range("B" & VarContFila).Value), "00")
                    If gsPeriodo <> gsPeriodoant Then
                        gsPeriodoant = Format(Val(gsPeriodo), "00")
                        gsNroVoucher = BuscaNvoucher(gsTipolibro, gsPeriodo)
                    Else
                        gsPeriodo = Format(Val(Sht.Range("B" & VarContFila).Value), "00")
                        gsPeriodoant = Format(Val(gsPeriodo), "00")
                        gsNroVoucher = gsNroVoucher + 1
                        gsNroVoucher = Format(Val(gsNroVoucher), "0000000000")
                    End If
                End If
                sSql = "Insert Into ZIMP_COMPRAS_CAB (" & _
                    "Ase_cNummov, Pan_cAnio, Per_cPeriodo, Lib_cTipoLibro, Ase_nVoucher, " & _
                    "Ase_dFecha, Ase_cGlosa, Ase_cTipoMoneda, Id_Exoneracion, Id_Tipo_Renta, Id_Modalidad, CreditoFiscal, " & _
                    "MaterialConstruccion, PercepcionIncluida) Values ('" & _
                    Format(Val(Sht.Range("A" & VarContFila).Value), "0000000000") & "','" & gsAnio & "','" & _
                    Format(Val(Sht.Range("B" & VarContFila).Value), "00") & "','" & _
                    gsTipolibro & "','" & gsNroVoucher & "','" & _
                    Sht.Range("C" & VarContFila).Value & "','" & _
                    Sht.Range("D" & VarContFila).Value & "','" & _
                    Format(Val(Sht.Range("E" & VarContFila).Value), "000") & "',null,null,null,null,null,null)"

                cn.BeginTrans
                cn.Execute sSql
                cn.CommitTrans
            Next
        ElseIf Me.optImportarArchivoTexto.Value Then
            If Not GetValidacionArchivo Then Exit Sub
            Call SaveCabecera
        End If

        Call cn.Execute("if exists (select 1 from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_COMPRAS_DET]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & _
        "drop table [dbo].[ZIMP_COMPRAS_DET]")
''hlp20230710
''''        sSql = "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_COMPRAS_DET] (" & _
''''            "[Ase_cNummov] char (10) NULL, " & "[Pan_cAnio] char (4) NULL, " & _
''''            "[Per_cPeriodo] char (2) NULL, " & "[Lib_cTipoLibro] char (2) NULL, " & _
''''            "[Ase_nVoucher] char (10) NULL, " & "[Pla_cCuentaContable] varchar (12) NULL, " & _
''''            "[Asd_nItem] INT NULL, " & "[Asd_cGlosa] nvarchar (255) NULL, " & _
''''            "[Asd_nDebeSoles] NUMERIC (14,2) NULL , " & "[Asd_nHaberSoles] NUMERIC (14,2) NULL , " & _
''''            "[Asd_nTipoCambio] NUMERIC (14,2) NULL , " & "[Asd_nDebeMonExt] NUMERIC (14,2) NULL , " & _
''''            "[Asd_nHaberMonExt] NUMERIC (14,2) NULL , " & "[Cos_cCodigo] varchar (12) NULL, " & _
''''            "[Ten_cTipoEntidad] char (1) NULL, " & "[Ent_cCodEntidad] char (5) NULL, " & _
''''            "[Asd_cTipoDoc] char (2) NULL, " & "[Asd_dFecDoc] datetime NULL, " & _
''''            "[Asd_cSerieDoc] varchar (20) NULL, " & "[Asd_cNumDoc] varchar (25) NULL, " & _
''''            "[Asd_dFecVen] datetime NULL, " & "[Asd_cTipoDocRef] nvarchar (255) NULL, " & _
''''            "[Asd_dFecDocRef] datetime NULL, " & "[Asd_cSerieDocRef] nvarchar (255) NULL, " & _
''''            "[Asd_cNumDocRef] nvarchar (255) NULL, " & "[Asd_nMontoInafecto] nvarchar (255) NULL, " & _
''''            "[Asd_cBaseImp] nvarchar (255) NULL, " & "[Asd_cRetencion] nvarchar (255) NULL, " & _
''''            "[Asd_dFechaSpot] datetime NULL, " & "[Asd_cNumSpot] nvarchar (255) NULL, " & _
''''            "[Asd_cProvCanc] char (1) NULL, " & "[Asd_cOperaTC] char (3) NULL, " & _
''''            "[Asd_cTipoMoneda] char (3) NULL, " & "[Asd_cComprobante] nvarchar (255) NULL, " & _
''''            "[Id_Exoneracion] char(10) NULL, [Id_Tipo_Renta] char(10) NULL, [Id_Modalidad] char(10) NULL, " & _
''''            "[Id_Aduana] char(10) NULL, [Id_Clasific_Servicio] char(10) NULL)"

            sSql = "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_COMPRAS_DET] (" & _
            "[Ase_cNummov] char (10) NULL, " & "[Pan_cAnio] char (4) NULL, " & _
            "[Per_cPeriodo] char (2) NULL, " & "[Lib_cTipoLibro] char (2) NULL, " & _
            "[Ase_nVoucher] char (10) NULL, " & "[Pla_cCuentaContable] varchar (12) NULL, " & _
            "[Asd_nItem] INT NULL, " & "[Asd_cGlosa] nvarchar (255) NULL, " & _
            "[Asd_nDebeSoles] NUMERIC (14,3) NULL , " & "[Asd_nHaberSoles] NUMERIC (14,3) NULL , " & _
            "[Asd_nTipoCambio] NUMERIC (14,3) NULL , " & "[Asd_nDebeMonExt] NUMERIC (14,3) NULL , " & _
            "[Asd_nHaberMonExt] NUMERIC (14,3) NULL , " & "[Cos_cCodigo] varchar (12) NULL, " & _
            "[Ten_cTipoEntidad] char (1) NULL, " & "[Ent_cCodEntidad] char (5) NULL, " & _
            "[Asd_cTipoDoc] char (2) NULL, " & "[Asd_dFecDoc] datetime NULL, " & _
            "[Asd_cSerieDoc] varchar (20) NULL, " & "[Asd_cNumDoc] varchar (25) NULL, " & _
            "[Asd_dFecVen] datetime NULL, " & "[Asd_cTipoDocRef] nvarchar (255) NULL, " & _
            "[Asd_dFecDocRef] datetime NULL, " & "[Asd_cSerieDocRef] nvarchar (255) NULL, " & _
            "[Asd_cNumDocRef] nvarchar (255) NULL, " & "[Asd_nMontoInafecto] nvarchar (255) NULL, " & _
            "[Asd_cBaseImp] nvarchar (255) NULL, " & "[Asd_cRetencion] nvarchar (255) NULL, " & _
            "[Asd_dFechaSpot] datetime NULL, " & "[Asd_cNumSpot] nvarchar (255) NULL, " & _
            "[Asd_cProvCanc] char (1) NULL, " & "[Asd_cOperaTC] char (3) NULL, " & _
            "[Asd_cTipoMoneda] char (3) NULL, " & "[Asd_cComprobante] nvarchar (255) NULL, " & _
            "[Id_Exoneracion] char(10) NULL, [Id_Tipo_Renta] char(10) NULL, [Id_Modalidad] char(10) NULL, " & _
            "[Id_Aduana] char(10) NULL, [Id_Clasific_Servicio] char(10) NULL)"


        cn.Execute sSql

        VarContFila = 0
        gsNroVoucher = ""
        gsPeriodo = ""
        If Me.optImportarExcel.Value Then
            Set Sht = Wb.Worksheets("COMPRAS_DET")
            lblAvance.Caption = "IMPORTANDO A TABLA TEMPORAL - COMPRASDET"
            Me.pbAvance.Min = 0
            Me.pbAvance.Value = 0
            Me.pbAvance.Max = Sht.Range("A65536").End(xlUp).Row

            For VarContFila = 2 To 70000
                pbAvance.Value = pbAvance.Value + 1
                'pbAvance.Refresh
                DoEvents
                If Sht.Range("A" & VarContFila).Value = "" Then Exit For

                '*****SI ENCUENTRA UNA SOLA COMILLA EN LA GLOSA*****
                If InStr(Sht.Range("E" & VarContFila).Value, "'") > 0 Then
                    Sht.Range("E" & VarContFila).Value = Replace(Sht.Range("E" & VarContFila).Value, "'", "")
                End If

                'Evaluando los nuevos campos LE
                strIdExoneracion = IIf(CE(Sht.Range("AD" & VarContFila).Value) = vbNullString, "Null", "'" & CE(Sht.Range("AD" & VarContFila).Value) & "'")
                strIdTipoRenta = IIf(CE(Sht.Range("AE" & VarContFila).Value) = vbNullString, "Null", "'" & CE(Sht.Range("AE" & VarContFila).Value) & "'")
                strIdModalidad = IIf(CE(Sht.Range("AF" & VarContFila).Value) = vbNullString, "Null", "'" & CE(Sht.Range("AF" & VarContFila).Value) & "'")
                strIdAduana = IIf(CE(Sht.Range("AG" & VarContFila).Value) = vbNullString, "Null", "'" & CE(Sht.Range("AG" & VarContFila).Value) & "'")
                strIdClasificacion = IIf(CE(Sht.Range("AH" & VarContFila).Value) = vbNullString, "Null", "'" & CE(Sht.Range("AH" & VarContFila).Value) & "'")

'                gscta = Left(Sht.Range("C" & VarContFila).Value, 2)
'                If gscta = "42" Then
'                    gsTipoEnt = "P"
'                    gsProvCanc = "P"
'                Else
'                    gsTipoEnt = ""
'                    gsProvCanc = ""
'                End If
                gsTipoEnt = Left(Sht.Range("L" & VarContFila).Value, 2)
                gsProvCanc = IIf(gsTipoEnt <> "", "P", "")

                If gsPeriodo = "" Then
                    gsPeriodo = Format(Val(Sht.Range("B" & VarContFila).Value), "00")
                    gsPeriodoant = gsPeriodo
                    gsNroVoucher = BuscaNvoucher(gsTipolibro, gsPeriodo)
                    gsnummovant = Format(Val(Sht.Range("A" & VarContFila).Value), "0000000000")
                Else
                    If gsnummovant <> Format(Val(Sht.Range("A" & VarContFila).Value), "0000000000") Then
                       gsNroVoucher = gsNroVoucher + 1
                       gsPeriodo = Format(Val(Sht.Range("B" & VarContFila).Value), "00")
                       If gsPeriodo = gsPeriodoant Then
                           gsNroVoucher = Format(Val(gsNroVoucher), "0000000000")
                           gsnummovant = Format(Val(Sht.Range("A" & VarContFila).Value), "0000000000")
                       Else
                           gsPeriodoant = gsPeriodo
                           gsNroVoucher = BuscaNvoucher(gsTipolibro, gsPeriodo)
                           gsnummovant = Format(Val(Sht.Range("A" & VarContFila).Value), "0000000000")
                        End If
                    Else
                        gsNroVoucher = Format(Val(gsNroVoucher), "0000000000")
                    End If
                End If

                sSql = "Insert Into ZIMP_COMPRAS_DET(Ase_cNummov, Pan_cAnio, Per_cPeriodo, " & _
                    "Lib_cTipoLibro, Ase_nVoucher, Pla_cCuentaContable, Asd_nItem, Asd_cGlosa, Asd_nDebeSoles, Asd_nHaberSoles, " & _
                    "Asd_nTipoCambio, Asd_nDebeMonExt, Asd_nHaberMonExt, Cos_cCodigo, Ten_cTipoEntidad, Ent_cCodEntidad, " & _
                    "Asd_cTipoDoc, Asd_dFecDoc, Asd_cSerieDoc, Asd_cNumDoc, Asd_dFecVen, Asd_cTipoDocRef, Asd_dFecDocRef, " & _
                    "Asd_cSerieDocRef, Asd_cNumDocRef, Asd_nMontoInafecto, Asd_cBaseImp, Asd_cRetencion, Asd_dFechaSpot, " & _
                    "Asd_cNumSpot, Asd_cProvCanc, Asd_cOperaTC, Asd_cTipoMoneda, Asd_cComprobante, Id_Exoneracion, " & _
                    "Id_Tipo_Renta, Id_Modalidad, Id_Aduana, Id_Clasific_Servicio) Values ('"

                sSql = sSql & Format(Val(Sht.Range("A" & VarContFila).Value), "0000000000") & "','" & gsAnio & "','"
                sSql = sSql & Format(Val(Sht.Range("B" & VarContFila).Value), "00") & "','"
                sSql = sSql & gsTipolibro & "','" & gsNroVoucher & "','"
                sSql = sSql & Sht.Range("C" & VarContFila).Value & "','"
                sSql = sSql & Sht.Range("D" & VarContFila).Value & "','"
                sSql = sSql & Sht.Range("E" & VarContFila).Value & "', "
                sSql = sSql & Sht.Range("F" & VarContFila).Value & ","
                sSql = sSql & NE(Sht.Range("G" & VarContFila).Value) & ","
                sSql = sSql & NE(Sht.Range("H" & VarContFila).Value) & ","
                sSql = sSql & NE(Sht.Range("I" & VarContFila).Value) & ","
                sSql = sSql & NE(Sht.Range("J" & VarContFila).Value) & ",'"
                sSql = sSql & Sht.Range("K" & VarContFila).Value & "','" & gsTipoEnt & "','"
                sSql = sSql & IIf(LTrim(RTrim(Sht.Range("M" & VarContFila).Value)) = "", "", Format(Val(Sht.Range("M" & VarContFila).Value), "00000")) & "','"
                sSql = sSql & IIf(Sht.Range("N" & VarContFila).Value = "", "", Format(Val(Sht.Range("N" & VarContFila).Value), "00")) & "','"
                sSql = sSql & IIf(Sht.Range("O" & VarContFila).Value = "", "", Format(Sht.Range("O" & VarContFila).Value, "dd/MM/yyyy")) & "','"
                sSql = sSql & IIf(IsNumeric(Sht.Range("P" & VarContFila).Value), Format(Val(Sht.Range("P" & VarContFila).Value), "0000"), Sht.Range("P" & VarContFila).Value) & "','"
                sSql = sSql & IIf(IsNumeric(Sht.Range("Q" & VarContFila).Value), Format(Val(Sht.Range("Q" & VarContFila).Value), "00000000"), vbNullString) & "','"
                sSql = sSql & IIf(Sht.Range("R" & VarContFila).Value = "", "", Format(Sht.Range("R" & VarContFila).Value, "dd/MM/yyyy")) & "','"
                sSql = sSql & Sht.Range("S" & VarContFila).Value & "','"
                sSql = sSql & IIf(Sht.Range("T" & VarContFila).Value = "", "", Format(Sht.Range("T" & VarContFila).Value, "dd/MM/yyyy")) & "','"
                sSql = sSql & Format(Val(Sht.Range("U" & VarContFila).Value), "0000") & "','"
                sSql = sSql & Sht.Range("V" & VarContFila).Value & "','" & Null & "','"
                sSql = sSql & Sht.Range("W" & VarContFila).Value & "','"
                sSql = sSql & Sht.Range("X" & VarContFila).Value & "','"
                sSql = sSql & IIf(Sht.Range("Y" & VarContFila).Value = "", "", Format(Sht.Range("Y" & VarContFila).Value, "dd/mm/yyyy")) & "','"
                sSql = sSql & Sht.Range("Z" & VarContFila).Value & "','" & gsProvCanc & "','"
                sSql = sSql & Sht.Range("AA" & VarContFila).Value & "', '"
                sSql = sSql & Format(Val(Sht.Range("AB" & VarContFila).Value), "000") & "','"
                sSql = sSql & Sht.Range("AC" & VarContFila).Value & "',"
                sSql = sSql & strIdExoneracion & ", " & strIdTipoRenta & ", " & strIdModalidad & ", " & strIdAduana & ", " & strIdClasificacion & ")"

                cn.BeginTrans
                cn.Execute sSql
                cn.CommitTrans

                strIdExoneracion = vbNullString: strIdTipoRenta = vbNullString: strIdModalidad = vbNullString: strIdAduana = vbNullString: strIdClasificacion = vbNullString
            Next
        ElseIf Me.optImportarArchivoTexto.Value Then
            If Not GetValidacionArchivo Then Exit Sub
            Call SaveDetalle
        End If
        ConectarAdvance
    End If

    If chkOpcion(4).Value = 1 Then 'Importar Ventas
        If Me.optImportarExcel.Value Then
            'Validacion del ingreso de montos en las dos monedas
            If gintBiMoneda = 1 Then
                lblnValor = True
                Set Sht = Wb.Worksheets("VENTAS_DET")
                For VarContFila = 2 To 70000
                    If Sht.Range("A" & VarContFila).Value <> "" Then
                        'Debe Soles - Debe Soles Dolar
                        If (CDbl(Sht.Range("F" & VarContFila).Value) > 0 And CDbl(Sht.Range("I" & VarContFila).Value) = 0) Or CDbl(Sht.Range("G" & VarContFila).Value) > 0 And CDbl(Sht.Range("J" & VarContFila).Value) = 0 Then
                            lblnValor = False
                            Exit For
                        End If
                    Else
                        Exit For
                    End If
                Next
                If lblnValor = False Then
                    MsgBox "La funcionalidad Bimoneda esta activada, ingrese los montos para ambas monedas [VENTAS_DET]", vbCritical, "Mensaje: Sistema"
                    Exit Sub
                End If
            End If
        End If

        gsTipolibro = "05"

        Call cn.Execute("if exists (select 1 from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_VENTAS_CAB]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & _
            "drop table [dbo].[ZIMP_VENTAS_CAB]")

        sSql = "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_VENTAS_CAB] (" & _
            "[Ase_cNummov] char (10) NULL, [Pan_cAnio] char (4) NULL, " & _
            "[Per_cPeriodo] char (2) NULL, [Lib_cTipoLibro] char (2) NULL, " & _
            "[Ase_nVoucher] char (10) NULL, [Ase_dFecha] nvarchar (255) NULL, " & _
            "[Ase_cGlosa] nvarchar (255) NULL, [Ase_cTipoMoneda] nvarchar (255) NULL, " & _
            "[Ase_cEstadoO] char (1) NULL, [Ase_cEstadoD] char (1) NULL, " & _
            "[Id_Exoneracion] char(10) NULL, [Id_Tipo_Renta] char(10) NULL, " & _
            "[Id_Modalidad] char(10) NULL, [CreditoFiscal] char(1) NULL, " & _
            "[MaterialConstruccion] char(1) NULL, [PercepcionIncluida] int NULL)"
        cn.Execute sSql

        gsPeriodo = ""
        gsNroVoucher = ""

        If Me.optImportarExcel.Value Then
            Set Sht = Wb.Worksheets("VENTAS_CAB")
            lblAvance.Caption = "IMPORTANDO A TABLA TEMPORAL - VENTASCAB"
            Me.pbAvance.Min = 0
            Me.pbAvance.Value = 0
            Me.pbAvance.Max = Sht.Range("A65536").End(xlUp).Row

            For VarContFila = 2 To 70000
                If Sht.Range("A" & VarContFila).Value = "" Then Exit For
                pbAvance.Value = pbAvance.Value + 1
                'pbAvance.Refresh
                DoEvents

                PeriodoXls = Format(Val(Sht.Range("B" & VarContFila).Value), "00")

                '*****SI ENCUENTRA UNA SOLA COMILLA EN LA GLOSA*****
                If InStr(Sht.Range("D" & VarContFila).Value, "'") > 0 Then
                    Sht.Range("D" & VarContFila).Value = Replace(Sht.Range("D" & VarContFila).Value, "'", "")
                End If

                If gsPeriodo = "" Then
                    gsPeriodo = Format(Val(Sht.Range("B" & VarContFila).Value), "00")
                    gsPeriodoant = Format(Val(gsPeriodo), "00")
                    gsNroVoucher = BuscaNvoucher(gsTipolibro, gsPeriodo)
                Else
                    gsPeriodo = Format(Val(Sht.Range("B" & VarContFila).Value), "00")
                    If gsPeriodo <> gsPeriodoant Then
                        gsPeriodoant = Format(Val(gsPeriodo), "00")
                        gsNroVoucher = BuscaNvoucher(gsTipolibro, gsPeriodo)
                    Else
                        gsPeriodo = Format(Val(Sht.Range("B" & VarContFila).Value), "00")
                        gsPeriodoant = Format(Val(gsPeriodo), "00")
                        gsNroVoucher = gsNroVoucher + 1
                        gsNroVoucher = Format(Val(gsNroVoucher), "0000000000")
                    End If
                End If

                sSql = "Insert Into ZIMP_VENTAS_CAB(Ase_cNummov, Pan_cAnio, Per_cPeriodo, " & _
                    "Lib_cTipoLibro, Ase_nVoucher, Ase_dFecha, Ase_cGlosa, Ase_cTipoMoneda, Id_Exoneracion, " & _
                    "Id_Tipo_Renta, Id_Modalidad, CreditoFiscal, MaterialConstruccion) Values ('" & _
                    Format(Val(Sht.Range("A" & VarContFila).Value), "0000000000") & "','" & gsAnio & "','" & _
                    Format(Val(Sht.Range("B" & VarContFila).Value), "00") & "','" & _
                    gsTipolibro & "','" & gsNroVoucher & "','" & _
                    Format(Sht.Range("C" & VarContFila).Value, "dd/mm/yyyy") & "','" & _
                    Sht.Range("D" & VarContFila).Value & "','" & _
                    Format(Sht.Range("E" & VarContFila).Value, "000") & "',null,null,null,'" & _
                    Sht.Range("F" & VarContFila).Value & "','" & Sht.Range("G" & VarContFila).Value & "')"
                
                cn.BeginTrans
                cn.Execute sSql
                cn.CommitTrans
            Next
        ElseIf Me.optImportarArchivoTexto.Value Then
            If Not GetValidacionArchivo Then Exit Sub
            Call SaveCabecera
        End If

        Call cn.Execute("if exists (select 1 from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_VENTAS_DET]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & _
        "drop table [dbo].[ZIMP_VENTAS_DET]")
''hlp20230710
''        sSql = "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_VENTAS_DET] (" & _
''            "[Ase_cNummov] char (10) NULL, [Pan_cAnio] char (4) NULL, " & _
''            "[Per_cPeriodo] char (2) NULL, [Lib_cTipoLibro] char (2) NULL, " & _
''            "[Ase_nVoucher] char (10) NULL, [Pla_cCuentaContable] varchar (12) NULL, " & _
''            "[Asd_nItem] INT NULL, [Asd_cGlosa] nvarchar (255) NULL, " & _
''            "[Asd_nDebeSoles] NUMERIC (14,2) NULL, [Asd_nHaberSoles] NUMERIC (14,2) NULL , " & _
''            "[Asd_nTipoCambio] NUMERIC (14,2) NULL, [Asd_nDebeMonExt] NUMERIC (14,2) NULL , " & _
''            "[Asd_nHaberMonExt] NUMERIC (14,2) NULL, [Cos_cCodigo] varchar (12) NULL, " & _
''            "[Ten_cTipoEntidad] char (1) NULL, [Ent_cCodEntidad] char (5) NULL, " & _
''            "[Asd_cTipoDoc] char (2) NULL, [Asd_dFecDoc] varchar(10) NULL, " & _
''            "[Asd_cSerieDoc] varchar (20) NULL, [Asd_cNumDoc] varchar (25) NULL, " & _
''            "[Asd_dFecVen] varchar(10) NULL, [Asd_nMontoInafecto] nvarchar (255) NULL, " & _
''            "[Asd_cBaseImp] nvarchar (255) NULL, [Asd_cProvCanc] char (1) NULL, " & _
''            "[Asd_cOperaTC] char (3) NULL, [Asd_cTipoMoneda] char (3) NULL, " & _
''            "[Asd_cTipoDocRef] nvarchar (255) NULL, [Asd_dFecDocRef] datetime NULL, " & _
''            "[Asd_cSerieDocRef] nvarchar (255) NULL, [Asd_cNumDocRef] nvarchar (255) NULL, " & _
''            "[Id_Exoneracion] char(10) NULL, [Id_Tipo_Renta] char(10) NULL, " & _
''            "[Id_Modalidad] char(10) NULL, [Id_Aduana] char(10) NULL, [Id_Clasific_Servicio] char(10) NULL)"


        sSql = "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_VENTAS_DET] (" & _
            "[Ase_cNummov] char (10) NULL, [Pan_cAnio] char (4) NULL, " & _
            "[Per_cPeriodo] char (2) NULL, [Lib_cTipoLibro] char (2) NULL, " & _
            "[Ase_nVoucher] char (10) NULL, [Pla_cCuentaContable] varchar (12) NULL, " & _
            "[Asd_nItem] INT NULL, [Asd_cGlosa] nvarchar (255) NULL, " & _
            "[Asd_nDebeSoles] NUMERIC (14,3) NULL, [Asd_nHaberSoles] NUMERIC (14,3) NULL , " & _
            "[Asd_nTipoCambio] NUMERIC (14,3) NULL, [Asd_nDebeMonExt] NUMERIC (14,3) NULL , " & _
            "[Asd_nHaberMonExt] NUMERIC (14,3) NULL, [Cos_cCodigo] varchar (12) NULL, " & _
            "[Ten_cTipoEntidad] char (1) NULL, [Ent_cCodEntidad] char (5) NULL, " & _
            "[Asd_cTipoDoc] char (2) NULL, [Asd_dFecDoc] varchar(10) NULL, " & _
            "[Asd_cSerieDoc] varchar (20) NULL, [Asd_cNumDoc] varchar (25) NULL, " & _
            "[Asd_dFecVen] varchar(10) NULL, [Asd_nMontoInafecto] nvarchar (255) NULL, " & _
            "[Asd_cBaseImp] nvarchar (255) NULL, [Asd_cProvCanc] char (1) NULL, " & _
            "[Asd_cOperaTC] char (3) NULL, [Asd_cTipoMoneda] char (3) NULL, " & _
            "[Asd_cTipoDocRef] nvarchar (255) NULL, [Asd_dFecDocRef] datetime NULL, " & _
            "[Asd_cSerieDocRef] nvarchar (255) NULL, [Asd_cNumDocRef] nvarchar (255) NULL, " & _
            "[Id_Exoneracion] char(10) NULL, [Id_Tipo_Renta] char(10) NULL, " & _
            "[Id_Modalidad] char(10) NULL, [Id_Aduana] char(10) NULL, [Id_Clasific_Servicio] char(10) NULL)"




        cn.Execute sSql

        gsNroVoucher = ""
        lblAvance.Caption = "IMPORTANDO A TABLA TEMPORAL - VENTASDET"
        Me.pbAvance.Min = 0
        Me.pbAvance.Value = 0
        gsPeriodo = ""

        If Me.optImportarExcel.Value Then
            Set Sht = Wb.Worksheets("VENTAS_DET")
            Me.pbAvance.Max = Sht.Range("A65536").End(xlUp).Row

            For VarContFila = 2 To 70000
                If Sht.Range("A" & VarContFila).Value = "" Then Exit For
                pbAvance.Value = pbAvance.Value + 1
                'pbAvance.Refresh
                DoEvents

                strIdExoneracion = IIf(CE(Sht.Range("Z" & VarContFila).Value) = vbNullString, "Null", "'" & CE(Sht.Range("Z" & VarContFila).Value) & "'")
                strIdTipoRenta = IIf(CE(Sht.Range("AA" & VarContFila).Value) = vbNullString, "Null", "'" & CE(Sht.Range("AA" & VarContFila).Value) & "'")
                strIdModalidad = IIf(CE(Sht.Range("AB" & VarContFila).Value) = vbNullString, "Null", "'" & CE(Sht.Range("AB" & VarContFila).Value) & "'")
                strIdAduana = IIf(CE(Sht.Range("AC" & VarContFila).Value) = vbNullString, "Null", "'" & CE(Sht.Range("AC" & VarContFila).Value) & "'")
                strIdClasificacion = IIf(CE(Sht.Range("AD" & VarContFila).Value) = vbNullString, "Null", "'" & CE(Sht.Range("AD" & VarContFila).Value) & "'")

                If InStr(Sht.Range("E" & VarContFila).Value, "'") > 0 Then
                    Sht.Range("E" & VarContFila).Value = Replace(Sht.Range("E" & VarContFila).Value, "'", "")
                End If

'                gscta = Left(Sht.Range("C" & VarContFila).Value, 2)
'                If gscta = "12" Then
'                    gsTipoEnt = "C"
'                    gsProvCanc = "P"
'                Else
'                    gsTipoEnt = ""
'                    gsProvCanc = ""
'                End If
                gsTipoEnt = Left(Sht.Range("L" & VarContFila).Value, 2)
                gsProvCanc = IIf(gsTipoEnt <> "", "P", "")
                
                If gsPeriodo = "" Then
                    gsPeriodo = Format(Val(Sht.Range("B" & VarContFila).Value), "00")
                    gsPeriodoant = gsPeriodo
                    gsNroVoucher = BuscaNvoucher(gsTipolibro, gsPeriodo)
                    gsnummovant = Format(Val(Sht.Range("A" & VarContFila).Value), "0000000000")
                Else
                    If gsnummovant <> Format(Val(Sht.Range("A" & VarContFila).Value), "0000000000") Then
                        gsNroVoucher = gsNroVoucher + 1
                        gsPeriodo = Format(Val(Sht.Range("B" & VarContFila).Value), "00")
                        If gsPeriodo = gsPeriodoant Then
                            gsNroVoucher = Format(Val(gsNroVoucher), "0000000000")
                            gsnummovant = Format(Val(Sht.Range("A" & VarContFila).Value), "0000000000")
                        Else
                            gsPeriodoant = gsPeriodo
                            gsNroVoucher = BuscaNvoucher(gsTipolibro, gsPeriodo)
                            gsnummovant = Format(Val(Sht.Range("A" & VarContFila).Value), "0000000000")
                        End If
                    Else
                        gsNroVoucher = Format(Val(gsNroVoucher), "0000000000")
                    End If
                End If

                sSql = "Insert Into ZIMP_VENTAS_DET(Ase_cNummov, Pan_cAnio, Per_cPeriodo, "
                sSql = sSql & "Lib_cTipoLibro, Ase_nVoucher, Pla_cCuentaContable, Asd_nItem, Asd_cGlosa, Asd_nDebeSoles, Asd_nHaberSoles, Asd_nTipoCambio, "
                sSql = sSql & "Asd_nDebeMonExt, Asd_nHaberMonExt, Cos_cCodigo, Ten_cTipoEntidad, Ent_cCodEntidad, Asd_cTipoDoc, Asd_dFecDoc, Asd_cSerieDoc, Asd_cNumDoc, "
                sSql = sSql & "Asd_dFecVen, Asd_nMontoInafecto, Asd_cBaseImp, Asd_cProvCanc, Asd_cOperaTC, Asd_cTipoMoneda, Asd_cTipoDocRef, Asd_dFecDocRef, Asd_cSerieDocRef, Asd_cNumDocRef, Id_Exoneracion, Id_Tipo_Renta, Id_Modalidad, Id_Aduana, Id_Clasific_Servicio) Values "
                sSql = sSql & "('" & Format(Val(Sht.Range("A" & VarContFila).Value), "0000000000") & "','"
                sSql = sSql & gsAnio & "','"
                sSql = sSql & Format(Val(Sht.Range("B" & VarContFila).Value), "00")
                sSql = sSql & "','" & gsTipolibro & "','"
                sSql = sSql & gsNroVoucher & "','"
                sSql = sSql & Sht.Range("C" & VarContFila).Value & "','"
                sSql = sSql & Sht.Range("D" & VarContFila).Value & "','"
                sSql = sSql & Sht.Range("E" & VarContFila).Value & "',"
                sSql = sSql & NE(Sht.Range("F" & VarContFila).Value) & ","
                sSql = sSql & NE(Sht.Range("G" & VarContFila).Value) & ","
                sSql = sSql & NE(Sht.Range("H" & VarContFila).Value) & ","
                sSql = sSql & NE(Sht.Range("I" & VarContFila).Value) & ","
                sSql = sSql & NE(Sht.Range("J" & VarContFila).Value) & ",'"
                sSql = sSql & Sht.Range("K" & VarContFila).Value & "','"
                sSql = sSql & gsTipoEnt & "','"
                sSql = sSql & IIf(LTrim(RTrim(Sht.Range("M" & VarContFila).Value)) = "", "", Format(Val(Sht.Range("M" & VarContFila).Value), "00000")) & "','"
                sSql = sSql & IIf(LTrim(RTrim(Sht.Range("N" & VarContFila).Value)) = "", "", Format(Sht.Range("N" & VarContFila).Value, "00")) & "','"
                sSql = sSql & IIf(LTrim(RTrim(Sht.Range("O" & VarContFila).Value)) = "", "", Format(CDate(Sht.Range("O" & VarContFila).Value), "dd/mm/yyyy")) & "','"
                If LTrim(RTrim(Sht.Range("P" & VarContFila).Value)) <> "" Then
                    sSql = sSql & LTrim(RTrim(Sht.Range("P" & VarContFila).Value)) & "','"
                Else
                    sSql = sSql & Format(Val(LTrim(RTrim(Sht.Range("P" & VarContFila).Value))), "0000") & "','"
                End If
                sSql = sSql & LTrim(RTrim(Sht.Range("Q" & VarContFila).Value)) & "','"

                If LTrim(RTrim(Sht.Range("R" & VarContFila).Value)) <> "" Then
                    sSql = sSql & Format(CDate(Sht.Range("R" & VarContFila).Value), "dd/mm/yyyy")
                End If

                sSql = sSql & "','" & Null & "','" & Sht.Range("S" & VarContFila).Value & "','" & gsProvCanc & "','" & _
                    Sht.Range("T" & VarContFila).Value & "','" & Format(Val(Sht.Range("U" & VarContFila).Value), "000") & "','" & _
                    Sht.Range("V" & VarContFila).Value & "','" & Sht.Range("W" & VarContFila).Value & "','" & _
                    Sht.Range("X" & VarContFila).Value & "','" & Sht.Range("Y" & VarContFila).Value & "'," & _
                    strIdExoneracion & "," & strIdTipoRenta & "," & strIdModalidad & "," & strIdAduana & "," & strIdClasificacion & ")"

                cn.BeginTrans
                cn.Execute sSql
                cn.CommitTrans

                sSql = "Update ZIMP_VENTAS_CAB Set Ase_cEstadoO=" & IIf(Trim(Sht.Range("E" & VarContFila).Value) = "ANULADO", "2", "1")
                sSql = sSql & ",Ase_cEstadoD='' Where Pan_cAnio='" & gsAnio & "' and "
                sSql = sSql & "Per_cPeriodo='" & Format(Val(Sht.Range("B" & VarContFila).Value), "00") & "' and "
                sSql = sSql & "Ase_nVoucher='" & gsNroVoucher & "'"

                cn.BeginTrans
                cn.Execute sSql
                cn.CommitTrans

                strIdExoneracion = vbNullString: strIdTipoRenta = vbNullString: strIdModalidad = vbNullString: strIdAduana = vbNullString: strIdClasificacion = vbNullString
            Next
            ConectarAdvance
        ElseIf Me.optImportarArchivoTexto.Value Then
            If Not GetValidacionArchivo Then Exit Sub
            Call SaveDetalle
        End If
    End If

    If chkOpcion(5).Value = 1 Then 'caja ingreso
        If Me.optImportarExcel.Value Then
            'Validacion del ingreso de montos en las dos monedas
            If gintBiMoneda = 1 Then
                lblnValor = True
                Set Sht = Wb.Worksheets("CAJAING_DET")
                For VarContFila = 2 To 70000
                    If Sht.Range("A" & VarContFila).Value <> "" Then
                        'Debe Soles - Debe Soles Dolar
                        If (CDbl(Sht.Range("I" & VarContFila).Value) > 0 And CDbl(Sht.Range("L" & VarContFila).Value) = 0) Or CDbl(Sht.Range("J" & VarContFila).Value) > 0 And CDbl(Sht.Range("M" & VarContFila).Value) = 0 Then
                            lblnValor = False
                            Exit For
                        End If
                    Else
                        Exit For
                    End If
                Next
                If lblnValor = False Then
                    MsgBox "La funcionalidad Bimoneda esta activada, ingrese los montos para ambas monedas [CAJAING_DET]", vbCritical, "Mensaje: Sistema"
                    Exit Sub
                End If
            End If
        End If

        Call cn.Execute("if exists (select 1 from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_CAJAING_CAB]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & _
        "drop table [dbo].[ZIMP_CAJAING_CAB]")

        sSql = "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_CAJAING_CAB] (" & _
            "[Ase_cNummov] char (10) NULL, " & _
            "[Pan_cAnio] char (4) NULL, " & _
            "[Per_cPeriodo] char (2) NULL, " & _
            "[Lib_cTipoLibro] char (2) NULL, " & _
            "[Ase_nVoucher] char (10) NULL, " & _
            "[Ase_dFecha] nvarchar (255) NULL, " & _
            "[Ase_cGlosa] nvarchar (255) NULL, " & _
            "[Ase_cTipoMoneda] nvarchar (255) NULL, [CreditoFiscal] char (1) NULL, " & _
            "[MaterialConstruccion] char (1) NULL, [PercepcionIncluida] int NULL)"
        cn.Execute sSql

        If Me.optImportarExcel.Value Then
            Set Sht = Wb.Worksheets("CAJAING_CAB")
            For VarContFila = 2 To 70000
                If Sht.Range("A" & VarContFila).Value = "" Then Exit For
            Next

            lblAvance.Caption = "IMPORTANDO A TABLA TEMPORAL - CAJAINGCAB"
            Me.pbAvance.Min = 0
            Me.pbAvance.Value = 0
            Me.pbAvance.Max = VarContFila - 1

            For VarContFila = 2 To 70000
                If Sht.Range("A" & VarContFila).Value = "" Then Exit For
                pbAvance.Value = pbAvance.Value + 1
                pbAvance.Refresh
                DoEvents

                'Modifique la columna I,J y K no vienen en el Excel JCS 21/04/2017 OCTALIA.S.A Chile
                sSql = "Insert Into ZIMP_CAJAING_CAB(Ase_cNummov, Pan_cAnio, Per_cPeriodo, " & _
                  "Lib_cTipoLibro, Ase_nVoucher, Ase_dFecha, Ase_cGlosa, Ase_cTipoMoneda, CreditoFiscal, MaterialConstruccion, PercepcionIncluida) Values " & _
                  "('" & Format(Val(Sht.Range("A" & VarContFila).Value), "0000000000") & "','" & Sht.Range("B" & VarContFila).Value & "','" & Format(Val(Sht.Range("C" & VarContFila).Value), "00") & _
                  "','" & Format(Val(Sht.Range("D" & VarContFila).Value), "00") & "','" & Format(Val(Sht.Range("E" & VarContFila).Value), "0000000000") & "','" & IIf(Sht.Range("F" & VarContFila).Value = "", "", Format(Sht.Range("F" & VarContFila).Value, "dd/mm/yyyy")) & _
                  "','" & Sht.Range("G" & VarContFila).Value & "','" & Format(Val(Sht.Range("H" & VarContFila).Value), "000") & "',NULL,NULL,NULL)"

                cn.BeginTrans
                cn.Execute sSql
                cn.CommitTrans
            Next
        ElseIf Me.optImportarArchivoTexto.Value Then
            If Not GetValidacionArchivo Then Exit Sub
            Call SaveCabecera
        End If

        Call cn.Execute("if exists (select 1 from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_CAJAING_DET]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & _
        "drop table [dbo].[ZIMP_CAJAING_DET]")

        sSql = "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_CAJAING_DET] (" & _
            "[Ase_cNummov] char (10) NULL, " & "[Pan_cAnio] char (4) NULL, " & _
            "[Per_cPeriodo] char (2) NULL, " & "[Lib_cTipoLibro] char (2) NULL, " & _
            "[Ase_nVoucher] char (10) NULL, " & "[Pla_cCuentaContable] varchar (12) NULL, " & _
            "[Asd_nItem] INT NULL, " & "[Asd_cGlosa] nvarchar (250) NULL, " & _
            "[Asd_nDebeSoles] NUMERIC (14,3) NULL , " & "[Asd_nHaberSoles] NUMERIC (14,3) NULL , " & _
            "[Asd_nTipoCambio] NUMERIC (14,3) NULL , " & "[Asd_nDebeMonExt] NUMERIC (14,3) NULL , " & _
            "[Asd_nHaberMonExt] NUMERIC (14,3) NULL , " & "[Cos_cCodigo] varchar (12) NULL, " & _
            "[Ten_cTipoEntidad] char (1) NULL, " & "[Ent_cCodEntidad] char (5) NULL, " & _
            "[Asd_cTipoDoc] char (2) NULL, " & "[Asd_dFecDoc] varchar(10) NULL, " & _
            "[Asd_cSerieDoc] varchar (20) NULL, " & "[Asd_cNumDoc] varchar (25) NULL, " & _
            "[Asd_dFecVen] varchar(10) NULL, " & "[Asd_cTipoDocRef] nvarchar(255) NULL, " & _
            "[Asd_dFecDocRef] varchar(10) NULL, " & "[Asd_cSerieDocRef] nvarchar (255) NULL, " & _
            "[Asd_cNumDocRef] nvarchar (255) NULL, " & "[Asd_cRetencion] nvarchar (255) NULL, " & _
            "[Asd_cProvCanc] char (1) NULL, " & "[Asd_cOperaTC] char (3) NULL, " & _
            "[Asd_cTipoMoneda] char (3) NULL, " & "[Tra_cCodigo] nvarchar (255) NULL, " & _
            "[Asd_cFormaPago] nvarchar (255) NULL)"
        cn.Execute sSql

        If Me.optImportarExcel.Value Then
            Set Sht = Wb.Worksheets("CAJAING_DET")
            For VarContFila = 2 To 70000
                If Sht.Range("A" & VarContFila).Value = "" Then Exit For
            Next

            lblAvance.Caption = "IMPORTANDO A TABLA TEMPORAL - CAJAINGDET"
            Me.pbAvance.Min = 0
            Me.pbAvance.Value = 0
            Me.pbAvance.Max = VarContFila - 1

            For VarContFila = 2 To 70000
                If Sht.Range("A" & VarContFila).Value = "" Then Exit For
                pbAvance.Value = pbAvance.Value + 1
                pbAvance.Refresh
                DoEvents

                sSql = "Insert Into ZIMP_CAJAING_DET(Ase_cNummov, Pan_cAnio, Per_cPeriodo, "
                sSql = sSql & "Lib_cTipoLibro, Ase_nVoucher, Pla_cCuentaContable, Asd_nItem, Asd_cGlosa, Asd_nDebeSoles, Asd_nHaberSoles, Asd_nTipoCambio, "
                sSql = sSql & "Asd_nDebeMonExt, Asd_nHaberMonExt, Cos_cCodigo, Ten_cTipoEntidad, Ent_cCodEntidad, Asd_cTipoDoc, Asd_dFecDoc, Asd_cSerieDoc, Asd_cNumDoc, "
                sSql = sSql & "Asd_dFecVen, Asd_cTipoDocRef, Asd_dFecDocRef, Asd_cSerieDocRef, Asd_cNumDocRef, Asd_cRetencion, Asd_cProvCanc, Asd_cOperaTC, "
                sSql = sSql & "Asd_cTipoMoneda, Tra_cCodigo, Asd_cFormaPago) Values "
                sSql = sSql & "('" & Format(Val(Sht.Range("A" & VarContFila).Value), "0000000000") & "','" & Sht.Range("B" & VarContFila).Value & "','" & IIf(LTrim(RTrim(Sht.Range("C" & VarContFila).Value)) = "", "", Format(Val(Sht.Range("C" & VarContFila).Value), "00"))
                sSql = sSql & "','" & Format(Val(Sht.Range("D" & VarContFila).Value), "00") & "','" & Format(Val(Sht.Range("E" & VarContFila).Value), "0000000000") & "','" & Sht.Range("F" & VarContFila).Value
                sSql = sSql & "','" & Sht.Range("G" & VarContFila).Value & "','" & Sht.Range("H" & VarContFila).Value & "', " & NE(Sht.Range("I" & VarContFila).Value)
                sSql = sSql & ", " & NE(Sht.Range("J" & VarContFila).Value) & ", " & NE(Sht.Range("K" & VarContFila).Value) & ", " & NE(Sht.Range("L" & VarContFila).Value)
                sSql = sSql & ", " & NE(Sht.Range("M" & VarContFila).Value) & ", '" & Sht.Range("N" & VarContFila).Value & "', '" & Sht.Range("O" & VarContFila).Value
                sSql = sSql & "', '" & IIf(LTrim(RTrim(Sht.Range("P" & VarContFila).Value)) = "", "", Format(Val(Sht.Range("P" & VarContFila).Value), "00000")) & "', '" & IIf(LTrim(RTrim(Sht.Range("Q" & VarContFila).Value)) = "", "", Format(Val(Sht.Range("Q" & VarContFila).Value), "00")) & "', '" & IIf(LTrim(RTrim(Sht.Range("R" & VarContFila).Value)) = "", "", Format(Sht.Range("R" & VarContFila).Value, "dd/mm/yyyy"))
                sSql = sSql & "', '" & IIf(LTrim(RTrim(Sht.Range("S" & VarContFila).Value)) = "", "", IIf(IsNumeric(Sht.Range("S" & VarContFila).Value), Format(Val(Sht.Range("S" & VarContFila).Value), "0000"), Sht.Range("S" & VarContFila).Value)) & "', '" & IIf(LTrim(RTrim(Sht.Range("T" & VarContFila).Value)) = "", "", Sht.Range("T" & VarContFila).Value) & "', '" & IIf(LTrim(RTrim(Sht.Range("U" & VarContFila).Value)) = "", "", Format(Sht.Range("U" & VarContFila).Value, "dd/mm/yyyy"))
                sSql = sSql & "', '" & IIf(LTrim(RTrim(Sht.Range("V" & VarContFila).Value)) = "", "", Format(Val(Sht.Range("V" & VarContFila).Value), "00")) & "', '" & IIf(LTrim(RTrim(Sht.Range("W" & VarContFila).Value)) = "", "", Format(Sht.Range("W" & VarContFila).Value, "dd/mm/yyyy")) & "', '" & Sht.Range("X" & VarContFila).Value & "', '" & Sht.Range("Y" & VarContFila).Value
                sSql = sSql & "', '" & Sht.Range("Z" & VarContFila).Value & "', '" & Sht.Range("AA" & VarContFila).Value & "', '" & Sht.Range("AB" & VarContFila).Value
                sSql = sSql & "', '" & IIf(LTrim(RTrim(Sht.Range("AC" & VarContFila).Value)) = "", "", Format(Val(Sht.Range("AC" & VarContFila).Value), "000")) & "', '" & Sht.Range("AD" & VarContFila).Value & "', '" & Sht.Range("AE" & VarContFila).Value
                sSql = sSql & "')"

                cn.BeginTrans
                cn.Execute sSql
                cn.CommitTrans
            Next
        ElseIf Me.optImportarArchivoTexto.Value Then
            If Not GetValidacionArchivo Then Exit Sub
            Call SaveDetalle
        End If
    End If

    If chkOpcion(6).Value = 1 Then 'caja egreso
        If Me.optImportarExcel.Value Then
            'Validacion del ingreso de montos en las dos monedas
            If gintBiMoneda = 1 Then
                lblnValor = True
                Set Sht = Wb.Worksheets("CAJAEGR_DET")
                For VarContFila = 2 To 70000
                    If Sht.Range("A" & VarContFila).Value <> "" Then
                        'Debe Soles - Debe Soles Dolar
                        If (CDbl(Sht.Range("I" & VarContFila).Value) > 0 And CDbl(Sht.Range("L" & VarContFila).Value) = 0) Or CDbl(Sht.Range("J" & VarContFila).Value) > 0 And CDbl(Sht.Range("M" & VarContFila).Value) = 0 Then
                            lblnValor = False
                            Exit For
                        End If
                    Else
                        Exit For
                    End If
                Next
                If lblnValor = False Then
                    MsgBox "La funcionalidad Bimoneda esta activada, ingrese los montos para ambas monedas [CAJAEGR_DET]", vbCritical, "Mensaje: Sistema"
                    Exit Sub
                End If
            End If
        End If

        Call cn.Execute("if exists (select 1 from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_CAJAEGR_CAB]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & _
        "drop table [dbo].[ZIMP_CAJAEGR_CAB]")

        sSql = "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_CAJAEGR_CAB] (" & _
            "[Ase_cNummov] char (10) NULL, [Pan_cAnio] char (4) NULL, " & _
            "[Per_cPeriodo] char (2) NULL, [Lib_cTipoLibro] char (2) NULL, " & _
            "[Ase_nVoucher] char (10) NULL, [Ase_dFecha] nvarchar (255) NULL, " & _
            "[Ase_cGlosa] nvarchar (255) NULL, [Ase_cTipoMoneda] nvarchar (255) NULL, " & _
            "[CreditoFiscal] char (1) NULL, [MaterialConstruccion] char (1) NULL, " & _
            "[PercepcionIncluida] int NULL)"
        cn.Execute sSql

        If Me.optImportarExcel.Value Then
            Set Sht = Wb.Worksheets("CAJAEGR_CAB")
            For VarContFila = 2 To 70000
                If Sht.Range("A" & VarContFila).Value = "" Then Exit For
            Next

            lblAvance.Caption = "IMPORTANDO A TABLA TEMPORAL - CAJAEGRESOCAB"
            Me.pbAvance.Min = 0
            Me.pbAvance.Value = 0
            Me.pbAvance.Max = VarContFila - 1

            For VarContFila = 2 To 70000
                If Sht.Range("A" & VarContFila).Value = "" Then Exit For
                pbAvance.Value = pbAvance.Value + 1
                pbAvance.Refresh
                DoEvents

                'Modifique la columna I,J y K no vienen en el Excel JCS 21/04/2017 OCTALIA.S.A Chile
                sSql = "Insert Into ZIMP_CAJAEGR_CAB(Ase_cNummov, Pan_cAnio, Per_cPeriodo, " & _
                    "Lib_cTipoLibro, Ase_nVoucher, Ase_dFecha, Ase_cGlosa, Ase_cTipoMoneda, CreditoFiscal, MaterialConstruccion, PercepcionIncluida) Values " & _
                    "('" & Format(Val(Sht.Range("A" & VarContFila).Value), "0000000000") & "','" & Sht.Range("B" & VarContFila).Value & "','" & Format(Val(Sht.Range("C" & VarContFila).Value), "00") & _
                    "','" & Format(Sht.Range("D" & VarContFila).Value, "00") & "','" & Format(Val(Sht.Range("E" & VarContFila).Value), "0000000000") & "','" & Sht.Range("F" & VarContFila).Value & _
                    "','" & Sht.Range("G" & VarContFila).Value & "','" & Format(Val(Sht.Range("H" & VarContFila).Value), "000") & "',NULL,NULL,NULL)"

                cn.BeginTrans
                cn.Execute sSql
                cn.CommitTrans
            Next
        ElseIf Me.optImportarArchivoTexto.Value Then
            If Not GetValidacionArchivo Then Exit Sub
            Call SaveCabecera
        End If

        Call cn.Execute("if exists (select 1 from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_CAJAEGR_DET]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & _
        "drop table [dbo].[ZIMP_CAJAEGR_DET]")

        sSql = "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_CAJAEGR_DET] (" & _
            "[Ase_cNummov] char (10) NULL, [Pan_cAnio] char (4) NULL, " & _
            "[Per_cPeriodo] char (2) NULL, [Lib_cTipoLibro] char (2) NULL, " & _
            "[Ase_nVoucher] char (10) NULL, [Pla_cCuentaContable] varchar (12) NULL, " & _
            "[Asd_nItem] INT NULL, [Asd_cGlosa] nvarchar (255) NULL, " & _
            "[Asd_nDebeSoles] NUMERIC (14,3) NULL, [Asd_nHaberSoles] NUMERIC (14,3) NULL , " & _
            "[Asd_nTipoCambio] NUMERIC (14,3) NULL, [Asd_nDebeMonExt] NUMERIC (14,3) NULL , " & _
            "[Asd_nHaberMonExt] NUMERIC (14,3) NULL, [Cos_cCodigo] varchar (12) NULL, " & _
            "[Ten_cTipoEntidad] char (1) NULL, [Ent_cCodEntidad] char (5) NULL, " & _
            "[Asd_cTipoDoc] char (2) NULL, [Asd_dFecDoc] datetime NULL, " & _
            "[Asd_cSerieDoc] varchar (20) NULL, [Asd_cNumDoc] varchar (25) NULL, " & _
            "[Asd_dFecVen] datetime NULL, [Asd_cTipoDocRef] nvarchar (255) NULL, " & _
            "[Asd_dFecDocRef] datetime NULL, [Asd_cSerieDocRef] nvarchar (255) NULL, " & _
            "[Asd_cNumDocRef] nvarchar (255) NULL, [Asd_cRetencion] nvarchar (255) NULL, " & _
            "[Asd_cProvCanc] char (1) NULL, [Asd_cOperaTC] char (3) NULL, " & _
            "[Asd_cTipoMoneda] char (3) NULL, [Tra_cCodigo] nvarchar (255) NULL, " & _
            "[Asd_cFormaPago] nvarchar (255) NULL)"
        cn.Execute sSql

        If Me.optImportarExcel.Value Then
            Set Sht = Wb.Worksheets("CAJAEGR_DET")
            For VarContFila = 2 To 70000
                If Sht.Range("A" & VarContFila).Value = "" Then Exit For
            Next

            lblAvance.Caption = "IMPORTANDO A TABLA TEMPORAL - CAJAEGRESODET"
            Me.pbAvance.Min = 0
            Me.pbAvance.Value = 0
            Me.pbAvance.Max = VarContFila - 1

            For VarContFila = 2 To 70000
                If Sht.Range("A" & VarContFila).Value = "" Then Exit For
                pbAvance.Value = pbAvance.Value + 1
                'pbAvance.Refresh
                DoEvents
                
                sSql = "Insert Into ZIMP_CAJAEGR_DET(Ase_cNummov, Pan_cAnio, Per_cPeriodo, " & _
                    "Lib_cTipoLibro, Ase_nVoucher, Pla_cCuentaContable, Asd_nItem, Asd_cGlosa, Asd_nDebeSoles, Asd_nHaberSoles, Asd_nTipoCambio, " & _
                    "Asd_nDebeMonExt, Asd_nHaberMonExt, Cos_cCodigo, Ten_cTipoEntidad, Ent_cCodEntidad, Asd_cTipoDoc, Asd_dFecDoc, Asd_cSerieDoc, Asd_cNumDoc, " & _
                    "Asd_dFecVen, Asd_cTipoDocRef, Asd_dFecDocRef, Asd_cSerieDocRef, Asd_cNumDocRef, Asd_cRetencion, Asd_cProvCanc, Asd_cOperaTC, " & _
                    "Asd_cTipoMoneda, Tra_cCodigo, Asd_cFormaPago) Values " & _
                    "('" & Format(Val(Sht.Range("A" & VarContFila).Value), "0000000000") & "','" & Sht.Range("B" & VarContFila).Value & "','" & Format(Val(Sht.Range("C" & VarContFila).Value), "00") & _
                    "','" & Format(Val(Sht.Range("D" & VarContFila).Value), "00") & "','" & Format(Val(Sht.Range("E" & VarContFila).Value), "0000000000") & "','" & Sht.Range("F" & VarContFila).Value & _
                    "','" & Sht.Range("G" & VarContFila).Value & "','" & Sht.Range("H" & VarContFila).Value & "', " & NE(Sht.Range("I" & VarContFila).Value) & _
                    ", " & NE(Sht.Range("J" & VarContFila).Value) & ", " & NE(Sht.Range("K" & VarContFila).Value) & ", " & NE(Sht.Range("L" & VarContFila).Value) & _
                    ", " & NE(Sht.Range("M" & VarContFila).Value) & ", '" & Sht.Range("N" & VarContFila).Value & "', '" & Sht.Range("O" & VarContFila).Value & _
                    "', '" & IIf(LTrim(RTrim(Sht.Range("P" & VarContFila).Value)) = "", "", Format(Val(Sht.Range("P" & VarContFila).Value), "00000")) & "', '" & IIf(LTrim(RTrim(Sht.Range("Q" & VarContFila).Value)) = "", "", Format(Val(Sht.Range("Q" & VarContFila).Value), "00")) & "', '" & IIf(LTrim(RTrim(Sht.Range("R" & VarContFila).Value)) = "", "", Format(Sht.Range("R" & VarContFila).Value, "dd/mm/yyyy")) & _
                    "', '" & IIf(IsNumeric(LTrim(RTrim(Sht.Range("S" & VarContFila).Value))) = False, RTrim(LTrim(Sht.Range("S" & VarContFila).Value)), Sht.Range("S" & VarContFila).Value) & "', '" & IIf(IsNumeric(Sht.Range("T" & VarContFila).Value) = False, RTrim(LTrim(Sht.Range("T" & VarContFila).Value)), Sht.Range("T" & VarContFila).Value) & "', '" & Sht.Range("U" & VarContFila).Value & _
                    "', '" & Sht.Range("V" & VarContFila).Value & "', '" & Sht.Range("W" & VarContFila).Value & "', '" & Sht.Range("X" & VarContFila).Value & _
                    "', '" & Sht.Range("Y" & VarContFila).Value & "', '" & Sht.Range("Z" & VarContFila).Value & "', '" & Sht.Range("AA" & VarContFila).Value & _
                    "', '" & Sht.Range("AB" & VarContFila).Value & "', '" & Format(Val(Sht.Range("AC" & VarContFila).Value), "000") & "', '" & Sht.Range("AD" & VarContFila).Value & _
                    "', '" & Sht.Range("AE" & VarContFila).Value & "')"

                cn.BeginTrans
                cn.Execute sSql
                cn.CommitTrans
            Next
        ElseIf Me.optImportarArchivoTexto.Value Then
            If Not GetValidacionArchivo Then Exit Sub
            Call SaveDetalle
        End If
    End If

    If chkOpcion(7).Value = 1 Then 'diario
        Call cn.Execute("if exists (select 1 from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_PLAN_CAB]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & _
        "drop table [dbo].[ZIMP_PLAN_CAB]")

        sSql = "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_PLAN_CAB] (" & _
            "[Ase_cNummov] char (10) NULL, [Pan_cAnio] char (4) NULL, " & _
            "[Per_cPeriodo] char (2) NULL, [Lib_cTipoLibro] char (2) NULL, " & _
            "[Ase_nVoucher] char (10) NULL, [Ase_dFecha] nvarchar (255) NULL, " & _
            "[Ase_cGlosa] nvarchar (255) NULL, [Ase_cTipoMoneda] nvarchar (255) NULL, " & _
            "[CreditoFiscal] char (1) NULL, [MaterialConstruccion] char (1) NULL, " & _
            "[PercepcionIncluida] int NULL)"
        cn.Execute sSql

        If Me.optImportarExcel.Value Then
            'Validacion del ingreso de montos en las dos monedas
            If gintBiMoneda = 1 Then
                lblnValor = True
                Set Sht = Wb.Worksheets("DIARIO_DET")
                For VarContFila = 2 To 70000
                    If Sht.Range("A" & VarContFila).Value <> "" Then
                        'Debe Soles - Debe Soles Dolar
                        If (CDbl(Sht.Range("I" & VarContFila).Value) > 0 And CDbl(Sht.Range("L" & VarContFila).Value) = 0) Or CDbl(Sht.Range("J" & VarContFila).Value) > 0 And CDbl(Sht.Range("M" & VarContFila).Value) = 0 Then
                            lblnValor = False
                            Exit For
                        End If
                    Else
                        Exit For
                    End If
                Next
                If lblnValor = False Then
                    MsgBox "La funcionalidad Bimoneda esta activada, ingrese los montos para ambas monedas [DIARIO_DET]", vbCritical, "Mensaje: Sistema"
                    Exit Sub
                End If
            End If

            Set Sht = Wb.Worksheets("DIARIO_CAB")
            For VarContFila = 2 To 70000
                If Sht.Range("A" & VarContFila).Value = "" Then Exit For
            Next

            lblAvance.Caption = "IMPORTANDO A TABLA TEMPORAL - DIARIOCAB"
            Me.pbAvance.Min = 0
            Me.pbAvance.Value = 0
            Me.pbAvance.Max = VarContFila - 1

            For VarContFila = 2 To 70000
                If Sht.Range("A" & VarContFila).Value = "" Then Exit For
                pbAvance.Value = pbAvance.Value + 1
                'pbAvance.Refresh
                DoEvents

                'Modifique la columna I,J y K no vienen en el Excel JCS 21/04/2017 OCTALIA.S.A Chile
                sSql = "Insert Into ZIMP_PLAN_CAB(Ase_cNummov, Pan_cAnio, Per_cPeriodo, " & _
                    "Lib_cTipoLibro, Ase_nVoucher, Ase_dFecha, Ase_cGlosa, Ase_cTipoMoneda, CreditoFiscal, MaterialConstruccion, PercepcionIncluida) Values " & _
                    "('" & Sht.Range("A" & VarContFila).Value & "','" & Sht.Range("B" & VarContFila).Value & "','" & Sht.Range("C" & VarContFila).Value & _
                    "','" & Sht.Range("D" & VarContFila).Value & "','" & Sht.Range("E" & VarContFila).Value & "','" & Sht.Range("F" & VarContFila).Value & _
                    "','" & Sht.Range("G" & VarContFila).Value & "','" & Sht.Range("H" & VarContFila).Value & "',NULL,NULL,NULL)"

                cn.BeginTrans
                cn.Execute sSql
                cn.CommitTrans
            Next
        ElseIf Me.optImportarArchivoTexto.Value Then
            If Not GetValidacionArchivo Then Exit Sub
            Call SaveCabecera
        End If

        Call cn.Execute("if exists (select 1 from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_PLAN_DET]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & _
        "drop table [dbo].[ZIMP_PLAN_DET]")

        sSql = "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_PLAN_DET] (" & _
            "[Ase_cNummov] char (10) NULL, [Pan_cAnio] char (4) NULL, " & _
            "[Per_cPeriodo] char (2) NULL, [Lib_cTipoLibro] char (2) NULL, " & _
            "[Ase_nVoucher] char (10) NULL, [Pla_cCuentaContable] varchar (12) NULL, " & _
            "[Asd_nItem] INT NULL, [Asd_cGlosa] nvarchar (255) NULL, " & _
            "[Asd_nDebeSoles] NUMERIC (14,3) NULL, [Asd_nHaberSoles] NUMERIC (14,3) NULL , " & _
            "[Asd_nTipoCambio] NUMERIC (14,3) NULL, [Asd_nDebeMonExt] NUMERIC (14,3) NULL , " & _
            "[Asd_nHaberMonExt] NUMERIC (14,3) NULL, [Cos_cCodigo] varchar (12) NULL, " & _
            "[Ten_cTipoEntidad] char (1) NULL, [Ent_cCodEntidad] char (5) NULL, " & _
            "[Asd_cTipoDoc] char (2) NULL, [Asd_dFecDoc] datetime NULL, " & _
            "[Asd_cSerieDoc] varchar (20) NULL, [Asd_cNumDoc] varchar (25) NULL, " & _
            "[Asd_dFecVen] datetime NULL, [Asd_cProvCanc] char (1) NULL, " & _
            "[Asd_cOperaTC] char (3) NULL, [Asd_cTipoMoneda] char (3) NULL)"
        cn.Execute sSql

        If Me.optImportarExcel.Value Then
            Set Sht = Wb.Worksheets("DIARIO_DET")
            For VarContFila = 2 To 70000
                If Sht.Range("A" & VarContFila).Value = "" Then Exit For
            Next

            lblAvance.Caption = "IMPORTANDO A TABLA TEMPORAL - DIARIODET"
            Me.pbAvance.Min = 0
            Me.pbAvance.Value = 0
            Me.pbAvance.Max = VarContFila - 1

            For VarContFila = 2 To 70000
                If Sht.Range("A" & VarContFila).Value = "" Then Exit For
                pbAvance.Value = pbAvance.Value + 1
                'pbAvance.Refresh
                DoEvents
                sSql = "Insert Into ZIMP_PLAN_DET(Ase_cNummov, Pan_cAnio, Per_cPeriodo, " & _
                    "Lib_cTipoLibro, Ase_nVoucher, Pla_cCuentaContable, Asd_nItem, Asd_cGlosa, Asd_nDebeSoles, " & _
                    "Asd_nHaberSoles, Asd_nTipoCambio, Asd_nDebeMonExt, Asd_nHaberMonExt, Cos_cCodigo, " & _
                    "Ten_cTipoEntidad, Ent_cCodEntidad, Asd_cTipoDoc, Asd_dFecDoc, Asd_cSerieDoc, " & _
                    "Asd_cNumDoc, Asd_dFecVen, Asd_cProvCanc, Asd_cOperaTC, Asd_cTipoMoneda) Values ('" & _
                    Sht.Range("A" & VarContFila).Value & "','" & Sht.Range("B" & VarContFila).Value & "','" & _
                    Sht.Range("C" & VarContFila).Value & "','" & Sht.Range("D" & VarContFila).Value & "','" & _
                    Sht.Range("E" & VarContFila).Value & "','" & Sht.Range("F" & VarContFila).Value & "','" & _
                    Sht.Range("G" & VarContFila).Value & "','" & Sht.Range("H" & VarContFila).Value & "', " & _
                    NE(Sht.Range("I" & VarContFila).Value) & ", " & NE(Sht.Range("J" & VarContFila).Value) & ", " & _
                    NE(Sht.Range("K" & VarContFila).Value) & ", " & NE(Sht.Range("L" & VarContFila).Value) & ", " & _
                    NE(Sht.Range("M" & VarContFila).Value) & ", '" & Sht.Range("N" & VarContFila).Value & "', '" & _
                    Sht.Range("O" & VarContFila).Value & "', '" & Sht.Range("P" & VarContFila).Value & "', '" & _
                    Sht.Range("Q" & VarContFila).Value & "', '" & Sht.Range("R" & VarContFila).Value & "', '" & _
                    IIf(IsNumeric(Sht.Range("S" & VarContFila).Value), Format(Val(Sht.Range("S" & VarContFila).Value), "0000"), Sht.Range("S" & VarContFila).Value) & "', '" & _
                    Sht.Range("T" & VarContFila).Value & "', '" & Sht.Range("U" & VarContFila).Value & "', '" & _
                    Sht.Range("V" & VarContFila).Value & "', '" & Sht.Range("W" & VarContFila).Value & "', '" & _
                    Sht.Range("X" & VarContFila).Value & "')"

                cn.BeginTrans
                cn.Execute sSql
                cn.CommitTrans
            Next
        ElseIf Me.optImportarArchivoTexto.Value Then
            If Not GetValidacionArchivo Then Exit Sub
            Call SaveDetalle
        End If
    End If
    
    'If Me.optImportarExcel.Value Then
    '    Wb.Saved = True
    'End If

    Wb.Close
    Set Sht = Nothing
    Set Wb = Nothing
    
    Ex.Quit
    Set Ex = Nothing
    Set cn = Nothing
    
    If Me.chkOpcion(0).Value = 0 And Me.chkOpcion(1).Value = 0 And Me.chkOpcion(2).Value = 0 And Me.chkOpcion(3).Value = 0 And Me.chkOpcion(4).Value = 0 And Me.chkOpcion(5).Value = 0 And Me.chkOpcion(6).Value = 0 And Me.chkOpcion(7).Value = 0 Then
        cmdImportarDatos.Enabled = True
        cmdSeleccionar.Enabled = True
        cmd_salir.Enabled = True
        cmdImprimir.Enabled = True
        Me.MousePointer = vbNormal
        MsgBox "Debe seleccionar una opcion", vbCritical, "Sistema ECB-Cont"
        Exit Sub
    End If
    
    FrmValidacioImportacion.pstrEntidad = Me.chkOpcion(0).Value
    FrmValidacioImportacion.pstrTipoCambio = Me.chkOpcion(1).Value
    FrmValidacioImportacion.pstrApertura = Me.chkOpcion(2).Value
    FrmValidacioImportacion.pstrCompra = Me.chkOpcion(3).Value
    FrmValidacioImportacion.pstrVenta = Me.chkOpcion(4).Value
    FrmValidacioImportacion.pstrIngreso = Me.chkOpcion(5).Value
    FrmValidacioImportacion.pstrEgreso = Me.chkOpcion(6).Value
    FrmValidacioImportacion.pstrDiario = Me.chkOpcion(7).Value
    Unload Me
    FrmValidacioImportacion.Show

    If Not FrmValidacioImportacion.pblnValidacion Then Exit Sub
    
    If ProcesaTablas Then
        DoEvents
        If Me.chkOpcion(2).Value = 1 Or chkOpcion(3).Value = 1 Or chkOpcion(4).Value = 1 Or _
          chkOpcion(5).Value = 1 Or chkOpcion(6).Value = 1 Or chkOpcion(7).Value = 1 Then
            If Mensajes("Desea actualizar los saldos ahora", vbQuestion + vbYesNo) = vbYes Then
                Call ActualizaSaldos
            End If
        End If
        DoEvents
    End If

    lblAvance.Caption = ""
    pbAvance.Value = 0
    pbAvance.Refresh
    DoEvents
    
    cmdImportarDatos.Enabled = True
    cmdSeleccionar.Enabled = True
    cmd_salir.Enabled = True
    cmdImprimir.Enabled = True
    Me.MousePointer = vbNormal
    Exit Sub

Control:
    Mensajes Err.Description & sSql & vbcrfl & "Fila: " & VarContFila, vbCritical + vbSystemModal
    cmdImportarDatos.Enabled = True
    cmdSeleccionar.Enabled = True
    cmd_salir.Enabled = True
    cmdImprimir.Enabled = True
    Me.MousePointer = vbNormal

    If Me.optImportarExcel.Value Then Ex.Quit
    Set cn = Nothing
    Set Sht = Nothing
    Set Wb = Nothing
    Set Ex = Nothing
End Sub

Private Sub Procesar_Archivos_Texto()

    Dim gsTipolibgsPeriodoro As String, gsPeriodo As String, gsPeriodoant As String
    Dim gsNroVoucher As String, gsnummovant As String
    Dim FrmValidacioImportacion As FrmValidacionImportacionOtroSistema
    Set FrmValidacioImportacion = New FrmValidacionImportacionOtroSistema

    Dim strIdExoneracion, strIdTipoRenta, strIdModalidad, strIdAduana, strIdClasificacion As String
    Dim strDomiciliado, strIdPais, strIdVinculo, strIdConvenio As String
    Dim File, registro, Ase_cNummov_enc, Per_cPeriodo_enc, Ase_cGlosa_enc, Ase_cTipoMoneda_enc, CreditoFiscal_enc, MateriaConstruccion_enc As String
    Dim numFile As Integer
    Dim Ase_dFecha_enc As Date
    Dim FileDet, Ase_cNummov_det, Per_cPeriodo_det, Pla_cCuentaContable_det As String
    Dim numFileDet, Asd_nItem_det As Integer
    Dim Ase_cGlosa_det As String
    Dim Asd_nDebeSoles_det, Asd_nHaberSoles_det, Asd_nTipoCambio_det, Asd_nDebeMonExt_det, Asd_nHaberMonExt_det As Single
    Dim Cos_cCodigo_det, Ent_cCodEntidad_det, Asd_cTipoDoc_det As String
    Dim Asd_dFecDoc_det, Asd_dFecVen_det, Asd_dFecDocRef_det, Asd_dFechaSpot_det As Date
    Dim Asd_cSerieDoc_det, Asd_cNumDoc_det, Asd_cTipoDocRef_det, Asd_cSerieDocRef_det, Asd_cNumDocRef_det As String
    Dim Asd_cBaseImp_det, Asd_cRetencion_det, Asd_cNumSpot_det, Asd_cOperaTC_det, Asd_cTipoMoneda_det As String
    Dim Asd_cComprobante_det, Id_Exoneracion_det, Id_Tipo_Renta_det, Id_Modalidad_det, Id_Aduana_det, Id_Clasific_Servicio_det As String
    Dim registrodet As String
    Dim tArraydet() As String
    Dim tArray() As String

    cmdImportarDatos.Enabled = False
    cmdSeleccionar.Enabled = False
    cmdSeleccionard.Enabled = False
    cmdImprimir.Enabled = False
    cmd_salir.Enabled = False
    Me.MousePointer = vbHourglass
    DoEvents
    
    sBaseDatos = gsBD
    Dim TipoLibroXls As String
    Dim VarSql As String
    Dim VarContFila As Long
    Dim cn As New ADODB.Connection
    
    cn.Open gsCadenaConexion
    cn.Execute ("set dateformat dmy")
    
    If chkOpcion(3).Value = 1 Then ' Importar Compras
        VarContFila = 0
        gsNroVoucher = ""
        gsPeriodo = ""
        gsTipolibro = "06"
    
        lblAvance.Caption = "IMPORTANDO A TABLA TEMPORAL - COMPRASCAB"
        Me.pbAvance.Min = 0
        Me.pbAvance.Value = 0
    
        File = tdbtArchivo.Text
        numFile = FreeFile
        Open File For Input As #numFile
        Do While Not EOF(numFile)
            Line Input #numFile, registro
            tArray = Split(registro, ";")
            If UBound(tArray) <> 6 Then
                Mensajes "Error en archivo plano (ENC) " & registro, vbCritical
                GoTo ERROR
            End If
            VarContFila = VarContFila + 1
        Loop
        Close #numFile
        
        Me.pbAvance.Value = 0
        Me.pbAvance.Max = VarContFila
        
        Call cn.Execute("if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_COMPRAS_CAB]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & _
        "drop table [dbo].[ZIMP_COMPRAS_CAB]")
        
        VarSql = "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_COMPRAS_CAB] (" & _
            "[Ase_cNummov] char (10) NULL, " & _
            "[Pan_cAnio] char (4) NULL, " & _
            "[Per_cPeriodo] char (2) NULL, " & _
            "[Lib_cTipoLibro] char (2) NULL, " & _
            "[Ase_nVoucher] char (10) NULL, " & _
            "[Ase_dFecha] DATETIME NULL, " & _
            "[Ase_cGlosa] nvarchar (255) NULL, " & _
            "[Ase_cTipoMoneda] nvarchar (255) NULL," & _
            "[Ase_cEstadoO] char (1) NULL , " & " [Ase_cEstadoD] char (1) NULL" & ", [Id_Exoneracion] char(10) NULL, [Id_Tipo_Renta] char(10) NULL, [Id_Modalidad] char(10) NULL, [CreditoFiscal] char(1) NULL, [MaterialConstruccion] char(1) NULL, [PercepcionIncluida] int NULL)"
        Call cn.Execute(VarSql)
        
        Open File For Input As #numFile
        Do While Not EOF(numFile)
            Line Input #numFile, registro
            tArray = Split(registro, ";")
            
            Ase_cNummov_enc = tArray(0)
            Per_cPeriodo_enc = tArray(1)
            Ase_dFecha_enc = IIf(tArray(2) <> "", tArray(2), "")
            Ase_cGlosa_enc = tArray(3)
            Ase_cTipoMoneda_enc = tArray(4)
            CreditoFiscal_enc = tArray(5)
            MateriaConstruccion_enc = tArray(6)
            
            pbAvance.Value = pbAvance.Value + 1
            'pbAvance.Refresh
            DoEvents
            If Ase_cNummov_enc <> "" Then
                If gsPeriodo = "" Then
                    gsPeriodo = Per_cPeriodo_enc
                    gsPeriodoant = Format(Val(gsPeriodo), "00")
                    gsNroVoucher = BuscaNvoucher(gsTipolibro, gsPeriodo)
                Else
                    gsPeriodo = Per_cPeriodo_enc
                    If gsPeriodo <> gsPeriodoant Then
                        gsPeriodoant = Format(Val(gsPeriodo), "00")
                        gsNroVoucher = BuscaNvoucher(gsTipolibro, gsPeriodo)
                    Else
                        gsPeriodo = Per_cPeriodo_enc
                        gsPeriodoant = Format(Val(gsPeriodo), "00")
                        gsNroVoucher = gsNroVoucher + 1
                        gsNroVoucher = Format(Val(gsNroVoucher), "0000000000")
                    End If
                End If
                cn.BeginTrans
                cn.Execute ("Insert Into ZIMP_COMPRAS_CAB(Ase_cNummov, Pan_cAnio, Per_cPeriodo, " & _
                "Lib_cTipoLibro, Ase_nVoucher, Ase_dFecha, Ase_cGlosa, Ase_cTipoMoneda, Id_Exoneracion, Id_Tipo_Renta, Id_Modalidad, CreditoFiscal, MaterialConstruccion, PercepcionIncluida) Values " & _
                "('" & Format(Ase_cNummov_enc, "0000000000") & "','" & gsAnio & "','" & _
                Format(Per_cPeriodo_enc, "00") & _
                "','" & gsTipolibro & "','" & gsNroVoucher & "', " & _
                "'" & Format(Ase_dFecha_enc, "dd/mm/yyyy") & "'" & _
                ",'" & Ase_cGlosa_enc & "','" & _
                Format(Ase_cTipoMoneda_enc, "000") & "', NULL, NULL, NULL,NULL,NULL,NULL)")
                cn.CommitTrans
            End If
        Loop
        Close #numFile
        
        'Carga detalle desde archivo plano
        VarContFila = 0
        gsNroVoucher = ""
        gsPeriodo = ""
        lblAvance.Caption = "IMPORTANDO A TABLA TEMPORAL - COMPRASDET"
        Me.pbAvance.Min = 0
        Me.pbAvance.Value = 0
        File = tdbtArchivodet.Text
        numFile = FreeFile
        Open File For Input As #numFile
        Do While Not EOF(numFile)
            registrodet = ""
            Line Input #numFile, registrodet
            tArraydet = Split(registrodet, ";")
            If UBound(tArraydet) <> 33 Then
                Mensajes "Error en archivo plano (DET) " & registrodet, vbCritical
                GoTo ERROR
            End If
            VarContFila = VarContFila + 1
        Loop
        Close #numFile
        Me.pbAvance.Max = VarContFila
        
        Call cn.Execute("if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_COMPRAS_DET]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & _
        "drop table [dbo].[ZIMP_COMPRAS_DET]")
 ''hlp20230710
'''''        VarSql = "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_COMPRAS_DET] (" & _
'''''            "[Ase_cNummov] char (10) NULL, " & "[Pan_cAnio] char (4) NULL, " & _
'''''            "[Per_cPeriodo] char (2) NULL, " & "[Lib_cTipoLibro] char (2) NULL, " & _
'''''            "[Ase_nVoucher] char (10) NULL, " & "[Pla_cCuentaContable] varchar (12) NULL, " & _
'''''            "[Asd_nItem] INT NULL, " & "[Asd_cGlosa] nvarchar (255) NULL, " & _
'''''            "[Asd_nDebeSoles] NUMERIC (14,2) NULL , " & "[Asd_nHaberSoles] NUMERIC (14,2) NULL , " & _
'''''            "[Asd_nTipoCambio] NUMERIC (14,2) NULL , " & "[Asd_nDebeMonExt] NUMERIC (14,2) NULL , " & _
'''''            "[Asd_nHaberMonExt] NUMERIC (14,2) NULL , " & "[Cos_cCodigo] varchar (12) NULL, " & _
'''''            "[Ten_cTipoEntidad] char (1) NULL, " & "[Ent_cCodEntidad] char (5) NULL, " & _
'''''            "[Asd_cTipoDoc] char (2) NULL, " & "[Asd_dFecDoc] datetime NULL, " & _
'''''            "[Asd_cSerieDoc] varchar (20) NULL, " & "[Asd_cNumDoc] varchar (25) NULL, " & _
'''''            "[Asd_dFecVen] datetime NULL, " & "[Asd_cTipoDocRef] nvarchar (255) NULL, " & _
'''''            "[Asd_dFecDocRef] datetime NULL, " & "[Asd_cSerieDocRef] nvarchar (255) NULL, " & _
'''''            "[Asd_cNumDocRef] nvarchar (255) NULL, " & "[Asd_nMontoInafecto] nvarchar (255) NULL, " & _
'''''            "[Asd_cBaseImp] nvarchar (255) NULL, " & "[Asd_cRetencion] nvarchar (255) NULL, " & _
'''''            "[Asd_dFechaSpot] datetime NULL, " & "[Asd_cNumSpot] nvarchar (255) NULL, " & _
'''''            "[Asd_cProvCanc] char (1) NULL, " & "[Asd_cOperaTC] char (3) NULL, " & _
'''''            "[Asd_cTipoMoneda] char (3) NULL, " & "[Asd_cComprobante] nvarchar (255) NULL" & _
'''''            " , [Id_Exoneracion] char(10) NULL, [Id_Tipo_Renta] char(10) NULL, [Id_Modalidad] char(10) NULL, [Id_Aduana] char(10) NULL, [Id_Clasific_Servicio] char(10) NULL)"
'''''
         VarSql = "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_COMPRAS_DET] (" & _
            "[Ase_cNummov] char (10) NULL, " & "[Pan_cAnio] char (4) NULL, " & _
            "[Per_cPeriodo] char (2) NULL, " & "[Lib_cTipoLibro] char (2) NULL, " & _
            "[Ase_nVoucher] char (10) NULL, " & "[Pla_cCuentaContable] varchar (12) NULL, " & _
            "[Asd_nItem] INT NULL, " & "[Asd_cGlosa] nvarchar (255) NULL, " & _
            "[Asd_nDebeSoles] NUMERIC (14,3) NULL , " & "[Asd_nHaberSoles] NUMERIC (14,3) NULL , " & _
            "[Asd_nTipoCambio] NUMERIC (14,3) NULL , " & "[Asd_nDebeMonExt] NUMERIC (14,3) NULL , " & _
            "[Asd_nHaberMonExt] NUMERIC (14,3) NULL , " & "[Cos_cCodigo] varchar (12) NULL, " & _
            "[Ten_cTipoEntidad] char (1) NULL, " & "[Ent_cCodEntidad] char (5) NULL, " & _
            "[Asd_cTipoDoc] char (2) NULL, " & "[Asd_dFecDoc] datetime NULL, " & _
            "[Asd_cSerieDoc] varchar (20) NULL, " & "[Asd_cNumDoc] varchar (25) NULL, " & _
            "[Asd_dFecVen] datetime NULL, " & "[Asd_cTipoDocRef] nvarchar (255) NULL, " & _
            "[Asd_dFecDocRef] datetime NULL, " & "[Asd_cSerieDocRef] nvarchar (255) NULL, " & _
            "[Asd_cNumDocRef] nvarchar (255) NULL, " & "[Asd_nMontoInafecto] nvarchar (255) NULL, " & _
            "[Asd_cBaseImp] nvarchar (255) NULL, " & "[Asd_cRetencion] nvarchar (255) NULL, " & _
            "[Asd_dFechaSpot] datetime NULL, " & "[Asd_cNumSpot] nvarchar (255) NULL, " & _
            "[Asd_cProvCanc] char (1) NULL, " & "[Asd_cOperaTC] char (3) NULL, " & _
            "[Asd_cTipoMoneda] char (3) NULL, " & "[Asd_cComprobante] nvarchar (255) NULL" & _
            " , [Id_Exoneracion] char(10) NULL, [Id_Tipo_Renta] char(10) NULL, [Id_Modalidad] char(10) NULL, [Id_Aduana] char(10) NULL, [Id_Clasific_Servicio] char(10) NULL)"
            
            
            
            
        Call cn.Execute(VarSql)
        
        Open File For Input As #numFile
        Do While Not EOF(numFile)
            Line Input #numFile, registrodet
            tArraydet = Split(registrodet, ";")
            Ase_cNummov_det = tArraydet(0)
            Per_cPeriodo_det = tArraydet(1)
            Pla_cCuentaContable_det = tArraydet(2)
            Asd_nItem_det = tArraydet(3)
            Asd_cGlosa_det = tArraydet(4)
            Asd_nDebeSoles_det = Replace(tArraydet(5), ",", "")
            Asd_nHaberSoles_det = Replace(tArraydet(6), ",", "")
            Asd_nTipoCambio_det = Replace(tArraydet(7), ",", "")
            Asd_nDebeMonExt_det = Replace(tArraydet(8), ",", "")
            Asd_nHaberMonExt_det = Replace(tArraydet(9), ",", "")
            Cos_cCodigo_det = tArraydet(10)
            Ten_cTipoEntidad_det = tArraydet(11)
            Ent_cCodEntidad_det = tArraydet(12)
            Asd_cTipoDoc_det = tArraydet(13)
            Asd_dFecDoc_det = tArraydet(14)
            Asd_cSerieDoc_det = tArraydet(15)
            Asd_cNumDoc_det = tArraydet(16)
            Asd_dFecVen_det = tArraydet(17)
            Asd_cTipoDocRef_det = tArraydet(18)
            Asd_dFecDocRef_det = tArraydet(19)
            Asd_cSerieDocRef_det = tArraydet(20)
            Asd_cNumDocRef_det = tArraydet(21)
            Asd_cBaseImp_det = tArraydet(22)
            Asd_cRetencion_det = tArraydet(23)
            Asd_dFechaSpot_det = IIf(tArraydet(24) = "", "01/01/1900", tArraydet(24))
            Asd_cNumSpot_det = tArraydet(25)
            Asd_cOperaTC_det = tArraydet(26)
            Asd_cTipoMoneda_det = tArraydet(27)
            Asd_cComprobante_det = tArraydet(28)
            Id_Exoneracion_det = tArraydet(29)
            Id_Tipo_Renta_det = tArraydet(30)
            Id_Modalidad_det = tArraydet(31)
            Id_Aduana_det = tArraydet(32)
            Id_Clasific_Servicio_det = tArraydet(33)

            'Validacion del ingreso de montos en las dos monedas
            If gintBiMoneda = 1 Then
                lblnValor = True
                If Ase_cNummov_det <> "" Then
                    'Debe Soles - Debe Soles Dolar
                    If (CDbl(Asd_nDebeSoles_det) > 0 And CDbl(Asd_nDebeMonExt) = 0) Or CDbl(Asd_nHaberSoles) > 0 And CDbl(Asd_nHaberMonExt) = 0 Then
                        lblnValor = False
                    End If
                End If
                If lblnValor = False Then
                    Mensajes "La funcionalidad Bimoneda esta activada, ingrese los montos para ambas monedas [COMPRAS_DET]", vbCritical
                    GoTo ERROR
                End If
            End If
            pbAvance.Value = pbAvance.Value + 1
            'pbAvance.Refresh
            DoEvents
            
            strIdExoneracion = IIf(CE(Id_Exoneracion_det) = vbNullString, "Null", "'" & CE(Id_Exoneracion_det) & "'")
            strIdTipoRenta = IIf(CE(Id_Tipo_Renta_det) = vbNullString, "Null", "'" & CE(Id_Tipo_Renta_det) & "'")
            strIdModalidad = IIf(CE(Id_Modalidad_det) = vbNullString, "Null", "'" & CE(Id_Modalidad_det) & "'")
            strIdAduana = IIf(CE(Id_Aduana_det) = vbNullString, "Null", "'" & CE(Id_Aduana_det) & "'")
            strIdClasificacion = IIf(CE(Id_Clasific_Servicio_det) = vbNullString, "Null", "'" & CE(Id_Clasific_Servicio_det) & "'")
            If InStr(Asd_cGlosa_det, "'") > 0 Then
                Asd_cGlosa_det = Replace(Asd_cGlosa_det, "'", "")
            End If

'            gscta = Left(Pla_cCuentaContable_det, 2)
'            If gscta = "42" Then
'                gsTipoEnt = "P"
'                gsProvCanc = "P"
'            Else
'                gsTipoEnt = ""
'                gsProvCanc = ""
'            End If
            gsTipoEnt = Ten_cTipoEntidad_det
            gsProvCanc = IIf(gsTipoEnt <> "", "P", "")

            If gsPeriodo = "" Then
                gsPeriodo = Per_cPeriodo_det
                gsPeriodoant = gsPeriodo
                gsNroVoucher = BuscaNvoucher(gsTipolibro, gsPeriodo)
                gsnummovant = Ase_cNummov_det
            Else
                If gsnummovant <> Ase_cNummov_det Then
                   gsNroVoucher = gsNroVoucher + 1
                   gsPeriodo = Per_cPeriodo_det
                   If gsPeriodo = gsPeriodoant Then
                       gsNroVoucher = Format(Val(gsNroVoucher), "0000000000")
                       gsnummovant = Ase_cNummov_det
                   Else
                       gsPeriodoant = gsPeriodo
                       gsNroVoucher = BuscaNvoucher(gsTipolibro, gsPeriodo)
                       gsnummovant = Ase_cNummov_det
                    End If
                Else
                    gsNroVoucher = Format(Val(gsNroVoucher), "0000000000")
                End If
            End If
            cn.BeginTrans
            sSql = "Insert Into ZIMP_COMPRAS_DET(Ase_cNummov, Pan_cAnio, Per_cPeriodo, "
            sSql = sSql & "Lib_cTipoLibro, Ase_nVoucher, Pla_cCuentaContable, Asd_nItem, Asd_cGlosa, Asd_nDebeSoles, Asd_nHaberSoles, Asd_nTipoCambio, "
            sSql = sSql & "Asd_nDebeMonExt, Asd_nHaberMonExt, Cos_cCodigo, Ten_cTipoEntidad, Ent_cCodEntidad, Asd_cTipoDoc, Asd_dFecDoc, Asd_cSerieDoc, Asd_cNumDoc, "
            sSql = sSql & "Asd_dFecVen, Asd_cTipoDocRef, Asd_dFecDocRef, Asd_cSerieDocRef, Asd_cNumDocRef, Asd_nMontoInafecto, Asd_cBaseImp, Asd_cRetencion, Asd_dFechaSpot, "
            sSql = sSql & "Asd_cNumSpot, Asd_cProvCanc, Asd_cOperaTC, Asd_cTipoMoneda, Asd_cComprobante, Id_Exoneracion, Id_Tipo_Renta, Id_Modalidad, Id_Aduana, Id_Clasific_Servicio) Values "
            sSql = sSql & "('" & Format(Ase_cNummov_det, "0000000000") & "','" & gsAnio & "','" & Format(Per_cPeriodo_det, "00")
            sSql = sSql & "','" & gsTipolibro & "','" & gsNroVoucher & "','" & Pla_cCuentaContable_det
            sSql = sSql & "','" & Asd_nItem_det & "','" & Asd_cGlosa_det & "', " & Asd_nDebeSoles_det
            sSql = sSql & ", " & NE(Asd_nHaberSoles_det) & ", " & NE(Asd_nTipoCambio_det) & ", " & NE(Asd_nDebeMonExt_det)
            sSql = sSql & ", " & NE(Asd_nHaberMonExt_det) & ", '" & Cos_cCodigo_det & "', '" & gsTipoEnt
            sSql = sSql & "', '" & IIf(LTrim(Ent_cCodEntidad_det) = "", "", Format(Val(Ent_cCodEntidad_det), "00000")) & "', '" & IIf(Asd_cTipoDoc_det = "", "", Format(Val(Asd_cTipoDoc_det), "00")) & "', '" & IIf(Asd_dFecDoc_det = "", "", Format(Asd_dFecDoc_det, "dd/MM/yyyy"))
            sSql = sSql & "', '" & IIf(IsNumeric(Asd_cSerieDoc_det), Format(Val(Asd_cSerieDoc_det), "0000"), Asd_cSerieDoc_det) & "', '" & IIf(IsNumeric(Asd_cNumDoc_det), Format(Val(Asd_cNumDoc_det), "00000000"), vbNullString) & "', '" & IIf(Asd_cNumDoc_det = "", "", Format(Asd_dFecVen_det, "dd/MM/yyyy"))
            sSql = sSql & "', '" & Asd_cTipoDocRef_det & "', '" & IIf(Asd_dFecDocRef_det = "", "", Format(Asd_dFecDocRef_det, "dd/MM/yyyy")) & "', '" & Format(Val(Asd_cSerieDocRef_det), "0000")
            sSql = sSql & "', '" & Asd_cNumDocRef_det & "', '" & Null & "', '" & Asd_cBaseImp_det
            sSql = sSql & "', '" & Asd_cRetencion_det & "', '"
'            sSql = sSql & IIf(Asd_dFechaSpot_det = "", "01/01/1900", Format(Asd_dFechaSpot_det, "dd/mm/yyyy"))
            sSql = sSql & Asd_dFechaSpot_det
            sSql = sSql & "', '" & Asd_cNumSpot_det
            sSql = sSql & "', '" & gsProvCanc & "', '" & Asd_cOperaTC_det & "', '" & Format(Val(Asd_cTipoMoneda_det), "000")
            sSql = sSql & "', '" & Asd_cComprobante_det & "', " & strIdExoneracion & ", " & strIdTipoRenta & ", " & strIdModalidad & ", " & strIdAduana & ", " & strIdClasificacion & ")"
            
            cn.Execute (sSql)
            cn.CommitTrans
        Loop
        Close #numFile
        ConectarAdvance
        cn.Close
    End If
    
    If chkOpcion(4).Value = 1 Then ' Importar Ventas
        VarContFila = 0
        gsNroVoucher = ""
        gsPeriodo = ""
        gsTipolibro = "05"
    
        lblAvance.Caption = "IMPORTANDO A TABLA TEMPORAL - VENTASCAB"
        Me.pbAvance.Min = 0
        Me.pbAvance.Value = 0
    
        File = tdbtArchivo.Text
        numFile = FreeFile
        Open File For Input As #numFile
        Do While Not EOF(numFile)
            Line Input #numFile, registro
            tArray = Split(registro, ";")
            If UBound(tArray) <> 6 Then
                Mensajes "Error en archivo plano (ENC) " & registro, vbCritical
                GoTo ERROR
            End If
            VarContFila = VarContFila + 1
        Loop
        Close #numFile
        Me.pbAvance.Max = VarContFila
        
        Call cn.Execute("if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_VENTAS_CAB]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & _
        "drop table [dbo].[ZIMP_VENTAS_CAB]")
        
        VarSql = "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_VENTAS_CAB] (" & _
            "[Ase_cNummov] char (10) NULL," & _
            "[Pan_cAnio] char (4) NULL," & _
            "[Per_cPeriodo] char (2) NULL," & _
            "[Lib_cTipoLibro] char (2) NULL," & _
            "[Ase_nVoucher] char (10) NULL," & _
            "[Ase_dFecha] DATETIME NULL," & _
            "[Ase_cGlosa] nvarchar (255) NULL," & _
            "[Ase_cTipoMoneda] nvarchar (255) NULL," & _
            "[Ase_cEstadoO] char (1) NULL," & _
            "[Ase_cEstadoD] char (1) NULL," & _
            "[Id_Exoneracion] char(10) NULL," & _
            "[Id_Tipo_Renta] char(10) NULL," & _
            "[Id_Modalidad] char(10) NULL," & _
            "[CreditoFiscal] char(1) NULL," & _
            "[MaterialConstruccion] char(1) NULL," & _
            "[PercepcionIncluida] int NULL)"
        Call cn.Execute(VarSql)
        
        Open File For Input As #numFile
        Do While Not EOF(numFile)
            Line Input #numFile, registro
            tArray = Split(registro, ";")
            
            Ase_cNummov_enc = tArray(0)
            Per_cPeriodo_enc = tArray(1)
            Ase_dFecha_enc = IIf(tArray(2) <> "", tArray(2), "")
            Ase_cGlosa_enc = tArray(3)
            Ase_cTipoMoneda_enc = tArray(4)
            CreditoFiscal_enc = tArray(5)
            MateriaConstruccion_enc = tArray(6)

            pbAvance.Value = pbAvance.Value + 1
            'pbAvance.Refresh
            DoEvents
            If Ase_cNummov_enc <> "" Then
                If gsPeriodo = "" Then
                    gsPeriodo = Per_cPeriodo_enc
                    gsPeriodoant = Format(Val(gsPeriodo), "00")
                    gsNroVoucher = BuscaNvoucher(gsTipolibro, gsPeriodo)
                Else
                    gsPeriodo = Per_cPeriodo_enc
                    If gsPeriodo <> gsPeriodoant Then
                        gsPeriodoant = Format(Val(gsPeriodo), "00")
                        gsNroVoucher = BuscaNvoucher(gsTipolibro, gsPeriodo)
                    Else
                        gsPeriodo = Per_cPeriodo_enc
                        gsPeriodoant = Format(Val(gsPeriodo), "00")
                        gsNroVoucher = gsNroVoucher + 1
                        gsNroVoucher = Format(Val(gsNroVoucher), "0000000000")
                    End If
                End If
                cn.BeginTrans
                cn.Execute "Insert Into ZIMP_VENTAS_CAB(Ase_cNummov,Pan_cAnio,Per_cPeriodo,Lib_cTipoLibro,Ase_nVoucher," & _
                    "Ase_dFecha,Ase_cGlosa,Ase_cTipoMoneda,Id_Exoneracion,Id_Tipo_Renta,Id_Modalidad,CreditoFiscal," & _
                    "MaterialConstruccion,PercepcionIncluida) Values ('" & _
                    Format(Ase_cNummov_enc, "0000000000") & "','" & gsAnio & "','" & _
                    Format(Per_cPeriodo_enc, "00") & "','" & gsTipolibro & "','" & gsNroVoucher & "','" & _
                    Format(Ase_dFecha_enc, "dd/mm/yyyy") & "','" & Ase_cGlosa_enc & "','" & _
                    Format(Ase_cTipoMoneda_enc, "000") & "',null,null,null,null,null,null)"
                cn.CommitTrans
            End If
        Loop
        Close #numFile
        
        'Carga detalle desde archivo plano
        VarContFila = 0
        gsNroVoucher = ""
        gsPeriodo = ""
        lblAvance.Caption = "IMPORTANDO A TABLA TEMPORAL - VENTASDET"
        Me.pbAvance.Min = 0
        Me.pbAvance.Value = 0
        File = tdbtArchivodet.Text
        numFile = FreeFile
        Open File For Input As #numFile
        Do While Not EOF(numFile)
            registrodet = ""
            Line Input #numFile, registrodet
            tArraydet = Split(registrodet, ";")
            If UBound(tArraydet) <> 29 Then
                Mensajes "Error en archivo plano (DET) " & registrodet, vbCritical
                GoTo ERROR
            End If
            VarContFila = VarContFila + 1
        Loop
        Close #numFile
        Me.pbAvance.Max = VarContFila
        
        Call cn.Execute("if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_VENTAS_DET]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & _
        "drop table [dbo].[ZIMP_VENTAS_DET]")
 ''hlp20230710
'''        VarSql = "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_VENTAS_DET] (" & _
'''            "[Ase_cNummov] char (10) NULL, " & "[Pan_cAnio] char (4) NULL, " & _
'''            "[Per_cPeriodo] char (2) NULL, " & "[Lib_cTipoLibro] char (2) NULL, " & _
'''            "[Ase_nVoucher] char (10) NULL, " & "[Pla_cCuentaContable] varchar (12) NULL, " & _
'''            "[Asd_nItem] INT NULL, " & "[Asd_cGlosa] nvarchar (255) NULL, " & _
'''            "[Asd_nDebeSoles] NUMERIC (14,2) NULL , " & "[Asd_nHaberSoles] NUMERIC (14,2) NULL , " & _
'''            "[Asd_nTipoCambio] NUMERIC (14,2) NULL , " & "[Asd_nDebeMonExt] NUMERIC (14,2) NULL , " & _
'''            "[Asd_nHaberMonExt] NUMERIC (14,2) NULL , " & "[Cos_cCodigo] varchar (12) NULL, " & _
'''            "[Ten_cTipoEntidad] char (1) NULL, " & "[Ent_cCodEntidad] char (5) NULL, " & _
'''            "[Asd_cTipoDoc] char (2) NULL, " & "[Asd_dFecDoc] varchar(10) NULL, " & _
'''            "[Asd_cSerieDoc] varchar (20) NULL, " & "[Asd_cNumDoc] varchar (25) NULL, " & _
'''            "[Asd_dFecVen] varchar(10) NULL, " & "[Asd_nMontoInafecto] nvarchar (255) NULL, " & _
'''            "[Asd_cBaseImp] nvarchar (255) NULL, " & "[Asd_cProvCanc] char (1) NULL, " & _
'''            "[Asd_cOperaTC] char (3) NULL, " & "[Asd_cTipoMoneda] char (3) NULL, " & _
'''            "[Asd_cTipoDocRef] nvarchar (255) NULL, " & _
'''            "[Asd_dFecDocRef] datetime NULL, " & "[Asd_cSerieDocRef] nvarchar (255) NULL, " & _
'''            "[Asd_cNumDocRef] nvarchar (255) NULL, [Id_Exoneracion] char(10) NULL, [Id_Tipo_Renta] char(10) NULL, [Id_Modalidad] char(10) NULL, [Id_Aduana] char(10) NULL, [Id_Clasific_Servicio] char(10) NULL )"


        VarSql = "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_VENTAS_DET] (" & _
            "[Ase_cNummov] char (10) NULL, " & "[Pan_cAnio] char (4) NULL, " & _
            "[Per_cPeriodo] char (2) NULL, " & "[Lib_cTipoLibro] char (2) NULL, " & _
            "[Ase_nVoucher] char (10) NULL, " & "[Pla_cCuentaContable] varchar (12) NULL, " & _
            "[Asd_nItem] INT NULL, " & "[Asd_cGlosa] nvarchar (255) NULL, " & _
            "[Asd_nDebeSoles] NUMERIC (14,3) NULL , " & "[Asd_nHaberSoles] NUMERIC (14,3) NULL , " & _
            "[Asd_nTipoCambio] NUMERIC (14,3) NULL , " & "[Asd_nDebeMonExt] NUMERIC (14,3) NULL , " & _
            "[Asd_nHaberMonExt] NUMERIC (14,3) NULL , " & "[Cos_cCodigo] varchar (12) NULL, " & _
            "[Ten_cTipoEntidad] char (1) NULL, " & "[Ent_cCodEntidad] char (5) NULL, " & _
            "[Asd_cTipoDoc] char (2) NULL, " & "[Asd_dFecDoc] varchar(10) NULL, " & _
            "[Asd_cSerieDoc] varchar (20) NULL, " & "[Asd_cNumDoc] varchar (25) NULL, " & _
            "[Asd_dFecVen] varchar(10) NULL, " & "[Asd_nMontoInafecto] nvarchar (255) NULL, " & _
            "[Asd_cBaseImp] nvarchar (255) NULL, " & "[Asd_cProvCanc] char (1) NULL, " & _
            "[Asd_cOperaTC] char (3) NULL, " & "[Asd_cTipoMoneda] char (3) NULL, " & _
            "[Asd_cTipoDocRef] nvarchar (255) NULL, " & _
            "[Asd_dFecDocRef] datetime NULL, " & "[Asd_cSerieDocRef] nvarchar (255) NULL, " & _
            "[Asd_cNumDocRef] nvarchar (255) NULL, [Id_Exoneracion] char(10) NULL, [Id_Tipo_Renta] char(10) NULL, [Id_Modalidad] char(10) NULL, [Id_Aduana] char(10) NULL, [Id_Clasific_Servicio] char(10) NULL )"


        Call cn.Execute(VarSql)

        Open File For Input As #numFile
        Do While Not EOF(numFile)
            Line Input #numFile, registrodet
            tArraydet = Split(registrodet, ";")

            Ase_cNummov_det = tArraydet(0)
            Per_cPeriodo_det = tArraydet(1)
            Pla_cCuentaContable_det = tArraydet(2)
            Asd_nItem_det = tArraydet(3)
            Asd_cGlosa_det = tArraydet(4)
            Asd_nDebeSoles_det = Replace(tArraydet(5), ",", "")
            Asd_nDebeSoles_det = Replace(Asd_nDebeSoles_det, "-", 0)
            Asd_nHaberSoles_det = Replace(tArraydet(6), ",", "")
            Asd_nHaberSoles_det = Replace(Asd_nHaberSoles_det, "-", 0)
            Asd_nTipoCambio_det = Replace(tArraydet(7), ",", "")
            Asd_nTipoCambio_det = Replace(Asd_nTipoCambio_det, "-", 0)
            Asd_nDebeMonExt_det = Replace(tArraydet(8), ",", "")
            Asd_nDebeMonExt_det = Replace(Asd_nDebeMonExt_det, "-", 0)
            Asd_nHaberMonExt_det = Replace(tArraydet(9), ",", "")
            Asd_nHaberMonExt_det = Replace(Asd_nHaberMonExt_det, "-", 0)
            Cos_cCodigo_det = tArraydet(10)
            Ten_cTipoEntidad_det = tArraydet(11)
            Ent_cCodEntidad_det = tArraydet(12)
            Asd_cTipoDoc_det = tArraydet(13)
            Asd_dFecDoc_det = tArraydet(14)
            Asd_cSerieDoc_det = tArraydet(15)
            Asd_cNumDoc_det = tArraydet(16)
            Asd_dFecVen_det = tArraydet(17)
            Asd_cBaseImp_det = tArraydet(18)
            Asd_cOperaTC_det = tArraydet(19)
            Asd_cTipoMoneda_det = tArraydet(20)
            Asd_cTipoDocRef_det = tArraydet(21)
            Asd_dFecDocRef_det = tArraydet(22)
            Asd_cSerieDocRef_det = tArraydet(23)
            Asd_cNumDocRef_det = tArraydet(24)
            Id_Exoneracion_det = tArraydet(25)
            Id_Tipo_Renta_det = tArraydet(26)
            Id_Modalidad_det = tArraydet(27)
            Id_Aduana_det = tArraydet(28)
            Id_Clasific_Servicio_det = tArraydet(29)
            
            'Validacion del ingreso de montos en las dos monedas
            If gintBiMoneda = 1 Then
                lblnValor = True
                If Ase_cNummov_det <> "" Then
                    'Debe Soles - Debe Soles Dolar
                    If (CDbl(Asd_nDebeSoles_det) > 0 And CDbl(Asd_nDebeMonExt) = 0) Or CDbl(Asd_nHaberSoles) > 0 And CDbl(Asd_nHaberMonExt) = 0 Then
                        lblnValor = False
                    End If
                End If
                If lblnValor = False Then
                    Mensajes "La funcionalidad Bimoneda esta activada, ingrese los montos para ambas monedas [VENTAS_DET]", vbCritical
                    GoTo ERROR
                End If
            End If
            pbAvance.Value = pbAvance.Value + 1
            'pbAvance.Refresh
            DoEvents

            strIdExoneracion = IIf(CE(Id_Exoneracion_det) = vbNullString, "Null", "'" & CE(Id_Exoneracion_det) & "'")
            strIdTipoRenta = IIf(CE(Id_Tipo_Renta_det) = vbNullString, "Null", "'" & CE(Id_Tipo_Renta_det) & "'")
            strIdModalidad = IIf(CE(Id_Modalidad_det) = vbNullString, "Null", "'" & CE(Id_Modalidad_det) & "'")
            strIdAduana = IIf(CE(Id_Aduana_det) = vbNullString, "Null", "'" & CE(Id_Aduana_det) & "'")
            strIdClasificacion = IIf(CE(Id_Clasific_Servicio_det) = vbNullString, "Null", "'" & CE(Id_Clasific_Servicio_det) & "'")
            If InStr(Asd_cGlosa, "'") > 0 Then
                Asd_cGlosa_det = Replace(Asd_cGlosa_det, "'", "")
            End If
'            gscta = Left(Pla_cCuentaContable_det, 2)
'            If gscta = "12" Then
'                gsTipoEnt = "C"
'                gsProvCanc = "P"
'            Else
'                gsTipoEnt = ""
'                gsProvCanc = ""
'            End If
            gsTipoEnt = Ten_cTipoEntidad_det
            gsProvCanc = IIf(gsTipoEnt <> "", "P", "")
            
            If gsPeriodo = "" Then
                gsPeriodo = Per_cPeriodo_det
                gsPeriodoant = gsPeriodo
                gsNroVoucher = BuscaNvoucher(gsTipolibro, gsPeriodo)
                gsnummovant = Ase_cNummov_det
            Else
                If gsnummovant <> Ase_cNummov_det Then
                    gsNroVoucher = gsNroVoucher + 1
                    gsPeriodo = Per_cPeriodo_det
                    If gsPeriodo = gsPeriodoant Then
                        gsNroVoucher = Format(Val(gsNroVoucher), "0000000000")
                        gsnummovant = Ase_cNummov_det
                    Else
                        gsPeriodoant = gsPeriodo
                        gsNroVoucher = BuscaNvoucher(gsTipolibro, gsPeriodo)
                        gsnummovant = Ase_cNummov_det
                    End If
                Else
                    gsNroVoucher = Format(Val(gsNroVoucher), "0000000000")
                End If
            End If
            
            sSql = "Insert Into ZIMP_VENTAS_DET(Ase_cNummov, Pan_cAnio, Per_cPeriodo, "
            sSql = sSql & "Lib_cTipoLibro, Ase_nVoucher, Pla_cCuentaContable, Asd_nItem, Asd_cGlosa, Asd_nDebeSoles, Asd_nHaberSoles, Asd_nTipoCambio, "
            sSql = sSql & "Asd_nDebeMonExt, Asd_nHaberMonExt, Cos_cCodigo, Ten_cTipoEntidad, Ent_cCodEntidad, Asd_cTipoDoc, Asd_dFecDoc, Asd_cSerieDoc, Asd_cNumDoc, "
            sSql = sSql & "Asd_dFecVen, Asd_nMontoInafecto, Asd_cBaseImp, Asd_cProvCanc, Asd_cOperaTC, Asd_cTipoMoneda, Asd_cTipoDocRef, Asd_dFecDocRef, Asd_cSerieDocRef, Asd_cNumDocRef, Id_Exoneracion, Id_Tipo_Renta, Id_Modalidad, Id_Aduana, Id_Clasific_Servicio) Values "
            sSql = sSql & "('" & Format(Val(Ase_cNummov_det), "0000000000") & "','"
            sSql = sSql & gsAnio & "','"
            sSql = sSql & Format(Val(Per_cPeriodo_det), "00")
            sSql = sSql & "','" & gsTipolibro & "','"
            sSql = sSql & gsNroVoucher & "','"
            sSql = sSql & Pla_cCuentaContable_det
            sSql = sSql & "','" & Asd_nItem_det & "','"
            sSql = sSql & Asd_cGlosa_det & "', "
            sSql = sSql & NE(Asd_nDebeSoles_det)
            sSql = sSql & ", " & NE(Asd_nHaberSoles_det) & ", "
            sSql = sSql & NE(Asd_nTipoCambio_det) & ", "
            sSql = sSql & NE(Asd_nDebeMonExt_det)
            sSql = sSql & ", " & NE(Asd_nHaberMonExt_det) & ", '"
            sSql = sSql & Cos_cCodigo_det & "', '"
            sSql = sSql & gsTipoEnt
            sSql = sSql & "', '" & IIf(LTrim(RTrim(Ent_cCodEntidad_det)) = "", "", Format(Val(Ent_cCodEntidad_det), "00000")) & "', '"
            sSql = sSql & IIf(LTrim(RTrim(Asd_cTipoDoc_det)) = "", "", Format(Asd_cTipoDoc_det, "00")) & "', '"
            sSql = sSql & IIf(LTrim(RTrim(Asd_dFecDoc_det)) = "", "", Format(CDate(Asd_dFecDoc_det), "dd/mm/yyyy"))
            If LTrim(RTrim(Asd_cSerieDoc_det)) <> "" Then
                sSql = sSql & "', '" & LTrim(RTrim(Asd_cSerieDoc_det)) & "',"
            Else
                sSql = sSql & "', '" & Format(Val(LTrim(Asd_cSerieDoc_det)), "0000") & "',"
            End If
            sSql = sSql & "'" & LTrim(RTrim(Asd_cNumDoc_det)) & "',"
            
            If LTrim(RTrim(Asd_dFecVen_det)) = "" Then
                sSql = sSql & "''"
            Else
                sSql = sSql & "'" & Format(CDate(Asd_dFecVen_det), "dd/mm/yyyy") & "'"
            End If
            sSql = sSql & ", '" & Null & "', '" & Asd_cBaseImp_det & "', '" & gsProvCanc
            sSql = sSql & "', '" & Asd_cOperaTC_det & "', '" & Format(Val(Asd_cTipoMoneda_det), "000") & "','" & _
                Asd_cTipoDocRef_det & "', '" & Asd_dFecDocRef_det & "', '" & Asd_cSerieDocRef_det & "', '" & Asd_cNumDocRef_det & "', " & strIdExoneracion & "," & strIdTipoRenta & ", " & strIdModalidad & ", " & strIdAduana & ", " & strIdClasificacion & ")"
        
            cn.BeginTrans
            Call cn.Execute(sSql)
            cn.CommitTrans
    
            sSql = "Update ZIMP_VENTAS_CAB Set Ase_cEstadoO=" & IIf(Trim(Asd_cGlosa_det) = "ANULADO", "2", "1")
            sSql = sSql & ",Ase_cEstadoD='' Where Pan_cAnio='" & gsAnio & "' and Per_cPeriodo='" & Format(Val(Per_cPeriodo_det), "00") & "'"
            sSql = sSql & " and Ase_nVoucher='" & gsNroVoucher & "'"
            cn.BeginTrans
            Call cn.Execute(sSql)
            cn.CommitTrans
                    
            strIdExoneracion = vbNullString: strIdTipoRenta = vbNullString: strIdModalidad = vbNullString: strIdAduana = vbNullString: strIdClasificacion = vbNullString
        Loop
        Close #numFile
        ConectarAdvance
        cn.Close
    End If

    FrmValidacioImportacion.pstrEntidad = Me.chkOpcion(0).Value
    FrmValidacioImportacion.pstrTipoCambio = Me.chkOpcion(1).Value
    FrmValidacioImportacion.pstrApertura = Me.chkOpcion(2).Value
    FrmValidacioImportacion.pstrCompra = Me.chkOpcion(3).Value
    FrmValidacioImportacion.pstrVenta = Me.chkOpcion(4).Value
    FrmValidacioImportacion.pstrIngreso = Me.chkOpcion(5).Value
    FrmValidacioImportacion.pstrEgreso = Me.chkOpcion(6).Value
    FrmValidacioImportacion.pstrDiario = Me.chkOpcion(7).Value
    Unload Me
    FrmValidacioImportacion.Show

    If Not FrmValidacioImportacion.pblnValidacion Then Exit Sub
    
    If ProcesaTablas Then
        DoEvents
        If chkOpcion(3).Value = 1 Or chkOpcion(4).Value = 1 Then
            If Mensajes("Desea actualizar los saldos ahora", vbQuestion + vbYesNo) = vbYes Then
                Call ActualizaSaldos
            End If
        End If
    End If
    
    DoEvents
    Exit Sub
    
ERROR:
    Close #numFile
    lblAvance.Caption = ""
    pbAvance.Value = 0
    cmdImportarDatos.Enabled = True
    cmdSeleccionar.Enabled = True
    cmdSeleccionard.Enabled = True
    cmd_salir.Enabled = True
    cmdImprimir.Enabled = True
    Me.MousePointer = vbNormal
End Sub
Public Sub Procesar()
    
    If ProcesaTablas Then
        DoEvents
        If Me.chkOpcion(2).Value = 1 Or chkOpcion(3).Value = 1 Or chkOpcion(4).Value = 1 Or _
          chkOpcion(5).Value = 1 Or chkOpcion(6).Value = 1 Or chkOpcion(7).Value = 1 Then
            If Mensajes("Desea actualizar los saldos ahora", vbQuestion + vbYesNo) = vbYes Then
                Call ActualizaSaldos
            End If
        End If
        DoEvents
    End If
    
End Sub

Private Sub cmdImprimir_Click()
    cmdImprimir.Enabled = False
    DoEvents
   
    Dim matriz_fecha(3) As Variant
    Screen.MousePointer = vbHourglass
    
    matriz_fecha(0) = "@Accion;REPORTE;True"
    matriz_fecha(1) = "@Emp_cCodigo;" & gsEmpresa & ";True"
    matriz_fecha(2) = "@EMPRESA;" & gsEmpresaNom & ";True"
    matriz_fecha(3) = "@RUC;" & gsRUC & ";True"
    
    Dim formulas(0) As Variant
    
    AbreReporteParam gsDSN, Me, rutaReportes & "RptImportacionDatos.rpt", crptToWindow, "Reporte de Importación de Datos", "", matriz_fecha(), formulas()
    
    Screen.MousePointer = vbNormal
    cmdImprimir.Enabled = True
End Sub

Private Sub cmdRefresh_Click()
    lblCorrelativo.Text = BuscaCorrelNummov()
End Sub

Private Sub cmdSeleccionar_Click()
    Me.tdbtArchivo = ""
    On Local Error GoTo ErrorEjecucion
    
    If Me.optImportarExcel.Value = True Then
        
        With Me.dlgAbrirArchivo
            .DialogTitle = "Archivo de Datos de Asientos"
            .filename = ""
            .InitDir = "C:"
            .Filter = "Archivos de Datos(*.xls)| *.xls"
            .CancelError = True
            .ShowOpen
            If .filename = "" Then
                Mensajes "Selecciones un archivo Excel", vbInformation
            Else
                tdbtArchivo = .filename
            End If
        End With
        
    Else
        If Me.optImportarArchivoTexto.Value = True Then
        
            With Me.dlgAbrirArchivo
               .DialogTitle = "Archivo TXT de Datos de Asientos Encabezado"
               .InitDir = "C:"
               .Filter = "Archivos de Datos(*.txt)| *.txt"
               .filename = ""
               .CancelError = True
                .ShowOpen
                If .filename = "" Then
                    Mensajes "Seleccione un archivo TXT Encabezado", vbInformation
               Else
                    tdbtArchivo = .filename
                End If
            End With
        End If
    End If
'        Me.tdbtArchivo = modAbrirCarpeta.GetFolder(frmPrcImportarDatosXLS.hwnd, "Select folder")
        
'    End If
    
    Exit Sub
ErrorEjecucion:
    If Err.Number <> 32755 Then Mensajes Str(Err.Number) & Err.Description, vbInformation
End Sub

Private Function BuscaCorrelNummov() As String
    On Error GoTo serror
    Dim sql As String
    sql = "select distinct max(ase_cnummov) as correlativo from cnd_asiento_voucher where emp_ccodigo='" & gsEmpresa & "' and pan_canio='" & gsAnio & "'"
    Dim scadena As String
    
    scadena = Right("0000000000" & NE(fRetornaValor(sql)) + 1, 10)
    BuscaCorrelNummov = scadena
    
    Exit Function
    
serror:
    BuscaCorrelNummov = ""

End Function

' tata-007 Función agregada para encontrar el siguiente nro. de voucher
Private Function BuscaNvoucher(gsTipolibro, gsPeriodo As String) As String
    On Error GoTo serror
    Dim sql As String
    Dim gsLibro As String
    sql = "select distinct MAX(Ase_nVoucher) as nVoucher from cnc_asiento_voucher where Emp_cCodigo='" & gsEmpresa & "' and Pan_cAnio='" & gsAnio & "' and Lib_cTipoLibro='" & gsTipolibro & "' and Per_cPeriodo='" & gsPeriodo & "'"
    Dim scadena As String
    scadena = Right("00" & NE(gsTipolibro), 2) & Right("00" & NE(gsPeriodo), 2) & Right("000000" & NE(fRetornaValor(sql)) + 1, 6)
    BuscaNvoucher = scadena
    
    Exit Function
    
serror:
    BuscaNvoucher = ""

End Function
' fin tata-007
' TATA-008 IMPORTATR ARCHIVO TXT
Private Sub cmdSeleccionard_Click()
'optImportarArchivoTexto
'    Me.tdbtArchivo = ""
'    On Local Error GoTo ErrorEjecucion
    Me.tdbtArchivodet = ""
    On Local Error GoTo ErrorEjecucion
    
    If Me.optImportarArchivoTexto.Value = True Then
        
        With Me.dlgAbrirArchivo
            .DialogTitle = "Archivo TXT de Datos de Asientos detalle"
            .InitDir = "C:"
            .Filter = "Archivos de Datos(*.txt)| *.txt"
            .CancelError = True
            .filename = ""
            .ShowOpen
            If .filename = "" Then
                Mensajes "Selecciones un archivo TXT de Asientos Detalle", vbInformation
            Else
                tdbtArchivodet = .filename
            End If
        End With
        
'        Me.tdbtArchivodet = modAbrirCarpeta.GetFolder(frmPrcImportarDatosXLS.hwnd, "Select folder")
        
    End If
    
    Exit Sub
ErrorEjecucion:
    If Err.Number <> 32755 Then Mensajes Str(Err.Number) & Err.Description, vbInformation
End Sub
' TATA-008 FIN

Private Sub Form_Load()
    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))
' TATA-008 IMPORTACION ARCHIVO TXT
    Label1(0).Caption = "SELECCIONE ARCHIVO EXCEL:"
    cmdSeleccionar.Caption = "EXCEL"
' TATA-008 FIN
    Call Centrar_form(Me)
    
    If Mid(gsGrupo, 3, 1) <> "1" And gsGrupo <> gsPrivilegioAdmin Then
        Me.cmdImportarDatos.Enabled = False
        Me.cmdSeleccionar.Enabled = False
    Else
        Me.cmdImportarDatos.Enabled = True
        Me.cmdSeleccionar.Enabled = True
    End If
    
    pbAvance.Min = 0
    pbAvance.Max = 24
    
    Call cmdRefresh_Click
End Sub

Private Function ProcesaTablas() As Boolean
    Dim lArrMnt(12) As Variant
    Dim i As Integer
    Dim sql As String
    Dim scadena As String
    Dim entro As Boolean
    Dim TipoLib As String
    Dim clsMante2 As clsMantoTablas
    
    Call EscribirLog("Iniciando importacion de plantilla XLS de la empresa " & gsEmpresa & " " & gsEmpresaNom, Me.Name)
    
    On Local Error GoTo ErrorEjecucion
    Set clsMante2 = New clsMantoTablas
    
    ProcesaTablas = True
    entro = False

    
    Screen.MousePointer = vbHourglass
    '-------------------------------------------'
    lArrMnt(0) = "ELIMINACION"
    lArrMnt(1) = gsEmpresa
    lArrMnt(2) = gsAnio
    lArrMnt(11) = gsUsuario
    
    clsMante2.InicializaClase
    clsMante2.BeginTrans
    
    If clsMante2.MantenimientoDeTablas(gsCadenaConexion, "spCn_ReprocesoImportacionXLSv2", lArrMnt(), False) = False Then
        gsImportacion = False
        ProcesaTablas = False
    End If
    
    clsMante2.CommitTrans
    clsMante2.FinalizaClase
    '-------------------------------------------'
    lArrMnt(0) = "IMPORTACION"
    lArrMnt(1) = gsEmpresa
    lArrMnt(2) = gsAnio
    lArrMnt(11) = gsUsuario
    
    For i = 0 To 7
        If chkOpcion(i).Value = vbChecked Then
            entro = True
            
            lArrMnt(3) = IIf(i = 0, "1", "0") 'ENTIDAD
            lArrMnt(4) = IIf(i = 1, "1", "0") 'TIPO DE CAMBIO
            lArrMnt(5) = IIf(i = 2, "1", "0") 'APERTURA
            lArrMnt(6) = IIf(i = 3, "1", "0") 'COMPRAS
            lArrMnt(7) = IIf(i = 4, "1", "0") 'VENTAS
            lArrMnt(8) = IIf(i = 5, "1", "0") 'CAJA INGRESO
            lArrMnt(9) = IIf(i = 6, "1", "0") 'CAJA EGRESO
            lArrMnt(10) = IIf(i = 7, "1", "0") 'PLANILLA
        
            gsImportacion = True
            
            DoEvents
            
            Select Case i
                Case 0: scadena = "ENTIDAD"
                Case 1: scadena = "TIPO DE CAMBIO"
                Case 2: scadena = "APERTURA"
                Case 3: scadena = "COMPRAS"
                Case 4: scadena = "VENTAS"
                Case 5: scadena = "CAJA INGRESO"
                Case 6: scadena = "CAJA EGRESO"
                Case 7: scadena = "PLANILLA"
            End Select
            
            Me.lblAvance.Caption = "PROCESANDO -> " & scadena
            
            lblAvance.Caption = "MIGRANDO A LAS TABLAS DE LA BD DEL PROCESO -> " & scadena
            DoEvents
            Me.pbAvance.Min = 0
            Me.pbAvance.Value = 0
            Me.pbAvance.Max = 30
            
            If pbAvance.Value + 1 > pbAvance.Max Then pbAvance.Max = pbAvance.Max + 1
            
            pbAvance.Value = pbAvance.Value + 1
            pbAvance.Refresh
            lblAvance.Caption = "Importando -> " & archivo
            lblAvance.Refresh
            
            clsMante2.InicializaClase
            clsMante2.BeginTrans
            
            If clsMante2.MantenimientoDeTablas(gsCadenaConexion, "spCn_ReprocesoImportacionXLSv2", lArrMnt(), False) = False Then
                gsImportacion = False
                ProcesaTablas = False
            End If
            
            pbAvance.Value = pbAvance.Max
            
            If scadena = "COMPRAS" Then
            
                lArrMnt(0) = "06"
                lArrMnt(1) = gsEmpresa
                lArrMnt(2) = gsAnio
                lArrMnt(3) = PeriodoXls
                lArrMnt(4) = "06"
                
                If clsMante2.MantenimientoDeTablas(gsCadenaConexion, "sp_Actualiza_Estados", lArrMnt(), False) = False Then
                End If

            End If
            
            lblAvance.Caption = "Proceso terminado ... "
            lblAvance.Caption = ""
            lblAvance.Refresh
            pbAvance.Value = 0
            
            clsMante2.CommitTrans
            clsMante2.FinalizaClase
            
            DoEvents
        End If
    Next i
    
    gsImportacion = False
    Set clsMante2 = Nothing
     
    '-----------------------------------
    sql = "select top 1 emp_ccodigo from CNT_REPORTE_IMPORTACION where emp_ccodigo='" & gsEmpresa & "' "
    scadena = CE(fRetornaValor(sql))
    
    If scadena <> "" Then
        Call Mensajes("Error en los datos de importación, revise el reporte de errores")
        ProcesaTablas = False
    Else
        If entro = False Then
            Call Mensajes("Seleccione una opcion")
        Else
            Call Mensajes("Proceso terminado")
            Call EscribirLog("Finalizo la importacion de plantilla XLS de la empresa " & gsEmpresa & " " & gsEmpresaNom, Me.Name)
        End If
        
    End If
    
    If entro = False Then ProcesaTablas = False
    '-----------------------------------
    Screen.MousePointer = vbNormal
    Exit Function
ErrorEjecucion:
    Call EscribirLog("Error de importacion de plantilla XLS [" & Err.Description & "] de la empresa " & gsEmpresa & " " & gsEmpresaNom, Me.Name)
    Mensajes Str(Err.Number) & " " & Err.Description, vbInformation
    Resume
    Set clsMante2 = Nothing
    gsImportacion = False
    Screen.MousePointer = vbNormal
End Function

Private Sub ActualizaSaldos()
        'Mensajes "SE INICIARA LA ACTUALIZACION DE CUENTAS DE DESTINO", vbOKOnly + vbExclamation
        frmPrcActualizaDestino.Show
        frmPrcActualizaDestino.cmdProcesar.Visible = False
        DoEvents
        frmPrcActualizaDestino.chkMes.Value = vbChecked
        frmPrcActualizaDestino.chkMes.Enabled = False
        frmPrcActualizaDestino.tdbcMes.BoundText = "14"
        DoEvents
        frmPrcActualizaDestino.gsMensaje = False
        frmPrcActualizaDestino.gsSinSaldos = True
        DoEvents
        frmPrcActualizaDestino.Procesar
        DoEvents
        frmPrcActualizaDestino.Cerrar

        DoEvents
        
        'Mensajes "SE INICIARA LA ACTUALIZACION DE SALDOS", vbOKOnly + vbExclamation
        frmPrcActualizaSaldos.Show
        frmPrcActualizaSaldos.cmdProcesar.Visible = False
        DoEvents
        frmPrcActualizaSaldos.chkMes.Value = vbChecked
        frmPrcActualizaSaldos.chkMes.Enabled = False
        DoEvents
        frmPrcActualizaSaldos.tdbcMes.BoundText = "14"
        DoEvents
        frmPrcActualizaSaldos.gsMensaje = False
        frmPrcActualizaSaldos.Procesar
        DoEvents
        frmPrcActualizaSaldos.Cerrar

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

Private Sub optImportarArchivoTexto_Click()
    If Me.optImportarArchivoTexto.Value = True Then
' TATA-008 IMPORTACION ARCHIVO TXT
        chkOpcion(0).Enabled = False
        chkOpcion(1).Enabled = False
        chkOpcion(2).Enabled = False
        chkOpcion(5).Enabled = False
        chkOpcion(6).Enabled = False
        chkOpcion(7).Enabled = False
        cmdSeleccionar.Default = True
        tdbtArchivo.Text = ""
        tdbtArchivodet.Text = ""
        Label1(0).Caption = "SELECCIONE ARCHIVOS DE ENCABEZADO Y DETALLE TXT:"
        cmdSeleccionar.Caption = "ENC"
        cmdSeleccionard.Visible = True
        tdbtArchivodet.Visible = True
        Me.tdbtArchivo.Enabled = False
        Me.tdbtArchivodet.Enabled = False
' TATA-008 FIN
    End If
End Sub

Private Sub optImportarExcel_Click()
    If Me.optImportarExcel.Value Then
' TATA-008 IMPORTACION ARCHIVO TXT
        chkOpcion(0).Enabled = True
        chkOpcion(1).Enabled = True
        chkOpcion(2).Enabled = True
        chkOpcion(5).Enabled = True
        chkOpcion(6).Enabled = True
        chkOpcion(7).Enabled = True
        tdbtArchivo.Text = ""
        tdbtArchivodet.Text = ""
        Label1(0).Caption = "SELECCIONE ARCHIVO EXCEL:"
        cmdSeleccionard.Visible = False
        cmdSeleccionar.Caption = "EXCEL"
        tdbtArchivodet.Visible = False
        Me.tdbtArchivo.Enabled = False
' TATA-008 FIN
    End If
End Sub


