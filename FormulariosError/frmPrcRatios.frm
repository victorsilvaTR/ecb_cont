VERSION 5.00
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmPrcRatios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Ratios Financieros"
   ClientHeight    =   3885
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5940
   Icon            =   "frmPrcRatios.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3885
   ScaleWidth      =   5940
   Begin VB.Frame fraTodo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3795
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   5745
      Begin VB.Frame Frame1 
         Caption         =   " Tipo de Reporte "
         Height          =   1050
         Left            =   225
         TabIndex        =   7
         Top             =   2025
         Width           =   5370
         Begin VB.OptionButton optDetalle 
            Caption         =   "Ver detalle del calculo"
            Height          =   375
            Index           =   0
            Left            =   1035
            TabIndex        =   3
            Top             =   585
            Width           =   3255
         End
         Begin VB.OptionButton optDetalle 
            Caption         =   "Ver solo resultado del calculo"
            Height          =   375
            Index           =   1
            Left            =   1050
            TabIndex        =   2
            Top             =   270
            Value           =   -1  'True
            Width           =   3255
         End
      End
      Begin TrueOleDBList70.TDBCombo tdbcMes 
         Height          =   300
         Left            =   1380
         TabIndex        =   0
         Top             =   1080
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   529
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         _DropdownWidth  =   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "codigo"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "descripcion"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   2
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).AllowRowSizing=   0   'False
         Splits(0).DividerStyle=   2
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=661"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=582"
         Splits(0)._ColumnProps(4)=   "Column(0)._VertColor=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=688"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=609"
         Splits(0)._ColumnProps(11)=   "Column(1)._VertColor=0"
         Splits(0)._ColumnProps(12)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
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
         RowDividerStyle =   1
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
         ListField       =   "descripcion"
         BoundColumn     =   ""
         IntegralHeight  =   0   'False
         CellTipsWidth   =   0
         CellTipsDelay   =   1000
         AutoDropdown    =   0   'False
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
         Locked          =   0   'False
         ScrollTrack     =   0   'False
         RowDividerColor =   12632256
         RowSubDividerColor=   12632256
         AddItemSeparator=   ";"
         _PropDict       =   $"frmPrcRatios.frx":0ECA
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000005&,.bold=0,.fontsize=675"
         _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(40)  =   "Named:id=33:Normal"
         _StyleDefs(41)  =   ":id=33,.parent=0"
         _StyleDefs(42)  =   "Named:id=34:Heading"
         _StyleDefs(43)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(44)  =   ":id=34,.wraptext=-1"
         _StyleDefs(45)  =   "Named:id=35:Footing"
         _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(47)  =   "Named:id=36:Selected"
         _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(49)  =   "Named:id=37:Caption"
         _StyleDefs(50)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(51)  =   "Named:id=38:HighlightRow"
         _StyleDefs(52)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(53)  =   "Named:id=39:EvenRow"
         _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(55)  =   "Named:id=40:OddRow"
         _StyleDefs(56)  =   ":id=40,.parent=33"
         _StyleDefs(57)  =   "Named:id=41:RecordSelector"
         _StyleDefs(58)  =   ":id=41,.parent=34"
         _StyleDefs(59)  =   "Named:id=42:FilterBar"
         _StyleDefs(60)  =   ":id=42,.parent=33"
      End
      Begin MSForms.CommandButton cmdSalir 
         Height          =   435
         Left            =   2910
         TabIndex        =   9
         Top             =   3195
         Width           =   1665
         Caption         =   " Salir"
         PicturePosition =   327683
         Size            =   "2937;767"
         Picture         =   "frmPrcRatios.frx":0F51
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdReporteRatio 
         Height          =   435
         Left            =   1350
         TabIndex        =   8
         Top             =   3210
         Width           =   1485
         Caption         =   " Imprimir"
         PicturePosition =   327683
         Size            =   "2619;767"
         Picture         =   "frmPrcRatios.frx":14EB
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdGenerar 
         Height          =   435
         Left            =   2138
         TabIndex        =   1
         Top             =   1530
         Width           =   1665
         Caption         =   " Calcular Ratios"
         PicturePosition =   327683
         Size            =   "2937;767"
         Picture         =   "frmPrcRatios.frx":1A85
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "RATIOS FINANCIEROS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1260
         TabIndex        =   6
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Seleccione el Mes"
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
         Left            =   1380
         TabIndex        =   5
         Top             =   720
         Width           =   1500
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
      TabIndex        =   10
      Top             =   0
      Width           =   4365
   End
End
Attribute VB_Name = "frmPrcRatios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gsGrupo As String
Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

Private Sub cmd_salir_Click()
    Unload Me
End Sub

Private Sub Procesar()
    Screen.MousePointer = vbHourglass
    Dim clsMante As clsMantoTablas
    Dim lArrMnt(5) As Variant
    Set clsMante = New clsMantoTablas

    lArrMnt(0) = "INSERTAR"
    lArrMnt(1) = gsEmpresa
    lArrMnt(2) = gsAnio
    lArrMnt(3) = tdbcMes.BoundText
    lArrMnt(4) = gsUsuario
    If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaProcIndicadores", lArrMnt(), True) = False Then
        Mensajes "El proceso no ha concluido. Verificar...", vbInformation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Screen.MousePointer = vbDefault
End Sub

Private Function VerificaPlantilla() As Boolean
    Dim sqlSp As String
    Dim clDatos As clsMantoTablas
    Dim lrsTabla As ADODB.Recordset
    Dim arrDatos() As Variant
    Dim registros As Integer
    
    VerificaPlantilla = True
    registros = 0
    
    Set clDatos = New clsMantoTablas
    Set lrsTabla = New ADODB.Recordset
    
    sqlSp = "spCn_GrabaIndicadores 'SEL_ALL', '" & gsEmpresa & "', '', '', '', 0, 0, '', '', '', '' "
    arrDatos = Array(sqlSp)
    Set lrsTabla = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If Not lrsTabla Is Nothing Then
       If lrsTabla.State = adStateOpen Then
            registros = lrsTabla.RecordCount
        End If
    End If
    
    If registros <= 0 Then
        Mensajes "Configure los indicadores , faltan datos para realizar el proceso de calculo de indicadores", vbOKOnly + vbInformation
        VerificaPlantilla = False
    End If
    
    Set lrsTabla = Nothing
    Set clDatos = Nothing
End Function


Private Sub cmdGenerar_Click()
    If VerificaPlantilla = True Then
        cmdGenerar.Enabled = False
        cmdsalir.Enabled = False
        cmdReporteRatio.Enabled = False
        DoEvents
        Screen.MousePointer = vbHourglass
        Call Procesar
        Screen.MousePointer = vbNormal
        Mensajes "Proceso de calculo Terminado"
        
        cmdGenerar.Enabled = True
        cmdsalir.Enabled = True
        cmdReporteRatio.Enabled = True
        
    End If
End Sub

Private Sub cmdReporteRatio_Click()
    cmdReporteRatio.Enabled = False
    
    Screen.MousePointer = vbHourglass
    
    If VerificaPlantilla = True Then
        
        Screen.MousePointer = vbHourglass
        Dim matriz_fecha(10) As Variant
        Dim Tipo As String
        
    
        matriz_fecha(0) = "@Accion;CONSULTA;True"
        matriz_fecha(1) = "@Emp_cCodigo;" & gsEmpresa & ";True"
        matriz_fecha(2) = "@Pan_cAnio;" & gsAnio & ";True"
        matriz_fecha(3) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
        matriz_fecha(4) = "@Ind_cUserCrea;" & gsUsuario & ";True"
        matriz_fecha(5) = "@NombreMes;" & NombreMes(tdbcMes.BoundText) & ";True"
        matriz_fecha(6) = "@Moneda;" & gsMonedaNac & ";True"
        matriz_fecha(7) = "EmpresaNom;" & gsEmpresaNom & ";True"
        
        
        matriz_fecha(8) = "@EMPRESA;" & gsEmpresaNom & ";True"
        matriz_fecha(9) = "@RUC;" & "RUC : " & gsRUC & ";True"
        matriz_fecha(10) = "@NomPeriodo;" & NombreMes(tdbcMes.BoundText) & ";True"
    
        Dim formulas(0) As Variant
        
        If optDetalle(1).Value = True Then
            AbreReporteParam gsDSN, Me, rutaReportes & "RptRatios.rpt", crptToWindow, "Reporte de Analisis de Ratios", "", matriz_fecha(), formulas()
            
        End If
        If optDetalle(0).Value = True Then
            AbreReporteParam gsDSN, Me, rutaReportes & "RptRatiosConValores.rpt", crptToWindow, "Reporte de Analisis de Ratios", "", matriz_fecha(), formulas()
        End If
        
        
    End If
    Screen.MousePointer = vbDefault
    ' ***
    cmdReporteRatio.Enabled = True
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Activate()
If gsPeriodo = "00" Then
        tdbcMes.BoundText = "01"
    ElseIf gsPeriodo = "13" Or gsPeriodo = "14" Then
        tdbcMes.BoundText = "12"
    Else
        tdbcMes.BoundText = gsPeriodo
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub Form_Load()
    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))
    
    
    Call Centrar_form(Me)
    
    Call LlenaComboMesApeAddItem(tdbcMes)
    If gsPeriodo = "00" Then
        tdbcMes.BoundText = "01"
    ElseIf gsPeriodo = "13" Or gsPeriodo = "14" Then
        tdbcMes.BoundText = "12"
    Else
        tdbcMes.BoundText = gsPeriodo
    End If
End Sub

Private Sub Form_Resize()
On Error GoTo errHand
    If Me.WindowState <> vbMinimized Then
        Call SeteaFondoForm(Me)
        Call Centrar_Objeto(fraTodo, Me)
        Call CentrarTitulo(lblTitulo, fraTodo, Me)
    End If
Exit Sub
errHand:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))

End Sub

Private Sub tdbcMes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pSetFocus cmdReporteRatio
End If
End Sub
