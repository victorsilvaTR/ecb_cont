VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importar datos de Plantilla XLS"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11670
   Icon            =   "frmPrcImportarDatosNueva.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8610
   ScaleWidth      =   11670
   Begin VB.Frame fraTodo 
      Height          =   8475
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11595
      Begin VB.Frame Frame1 
         Caption         =   "Lista de Datos"
         Height          =   5415
         Left            =   120
         TabIndex        =   16
         Top             =   1920
         Width           =   10815
         Begin VB.ListBox LstSeleccionados 
            Appearance      =   0  'Flat
            DragIcon        =   "frmPrcImportarDatosNueva.frx":0ECA
            Height          =   4710
            Left            =   3960
            TabIndex        =   19
            Top             =   360
            Width           =   3255
         End
         Begin VB.ListBox LstColumnas 
            Appearance      =   0  'Flat
            Height          =   4710
            Left            =   7320
            TabIndex        =   18
            Top             =   360
            Width           =   3255
         End
         Begin VB.ListBox LstListas 
            Appearance      =   0  'Flat
            DragIcon        =   "frmPrcImportarDatosNueva.frx":1454
            Height          =   4710
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   3255
         End
      End
      Begin VB.ComboBox CboEntidades 
         Height          =   315
         ItemData        =   "frmPrcImportarDatosNueva.frx":19DE
         Left            =   5880
         List            =   "frmPrcImportarDatosNueva.frx":19FA
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   600
         Width           =   5055
      End
      Begin TDBText6Ctl.TDBText tdbtArchivo 
         Height          =   375
         Left            =   180
         TabIndex        =   1
         Top             =   1395
         Width           =   10710
         _Version        =   65536
         _ExtentX        =   18891
         _ExtentY        =   661
         Caption         =   "frmPrcImportarDatosNueva.frx":1A78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPrcImportarDatosNueva.frx":1AE4
         Key             =   "frmPrcImportarDatosNueva.frx":1B02
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
         Appearance      =   0
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
         Left            =   180
         TabIndex        =   2
         Top             =   7680
         Width           =   11325
         _ExtentX        =   19976
         _ExtentY        =   344
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin TDBText6Ctl.TDBText lblCorrelativo 
         Height          =   375
         Left            =   180
         TabIndex        =   3
         Top             =   540
         Width           =   4830
         _Version        =   65536
         _ExtentX        =   8520
         _ExtentY        =   661
         Caption         =   "frmPrcImportarDatosNueva.frx":1B46
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPrcImportarDatosNueva.frx":1BB2
         Key             =   "frmPrcImportarDatosNueva.frx":1BD0
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
      Begin MSComDlg.CommonDialog dlgAbrirArchivo 
         Left            =   -120
         Top             =   1200
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSForms.CommandButton CommandButton1 
         Height          =   435
         Left            =   240
         TabIndex        =   15
         Top             =   7980
         Width           =   1665
         Caption         =   "Cargar Configuracion"
         PicturePosition =   327683
         Size            =   "2937;767"
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
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   345
         Index           =   0
         Left            =   195
         TabIndex        =   13
         Top             =   1080
         Width           =   1950
      End
      Begin VB.Label lblAvance 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -11100
         TabIndex        =   12
         Top             =   6240
         Width           =   5325
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
         Left            =   5880
         TabIndex        =   11
         Top             =   240
         Width           =   3300
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
         Left            =   240
         TabIndex        =   10
         Top             =   7320
         Width           =   900
      End
      Begin MSForms.CommandButton cmdSeleccionar 
         Height          =   390
         Left            =   10920
         TabIndex        =   9
         Top             =   1350
         Width           =   450
         PicturePosition =   262148
         Size            =   "794;688"
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
         TabIndex        =   8
         Top             =   225
         Width           =   1275
      End
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   390
         Left            =   5040
         TabIndex        =   7
         ToolTipText     =   " Vuelve a cargar los datos almacenados "
         Top             =   540
         Width           =   450
         PicturePosition =   262148
         Size            =   "794;688"
         Picture         =   "frmPrcImportarDatosNueva.frx":1C14
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdImprimir 
         Height          =   435
         Left            =   3780
         TabIndex        =   6
         Top             =   7995
         Width           =   1665
         Caption         =   " Imprimir Errores"
         PicturePosition =   327683
         Size            =   "2937;767"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdImportarDatos 
         Height          =   435
         Left            =   1995
         TabIndex        =   5
         Top             =   7980
         Width           =   1665
         Caption         =   " Importar Datos"
         PicturePosition =   327683
         Size            =   "2937;767"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmd_salir 
         Height          =   435
         Left            =   5535
         TabIndex        =   4
         Top             =   7995
         Width           =   1665
         Caption         =   " Salir"
         PicturePosition =   327683
         Size            =   "2937;767"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gsGrupo As String

Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

Private Sub CommandButton1_Click()
 Call Cargar_Columnas
End Sub

Private Sub Form_Load()
'Call Limpiar
Dim i As Integer
  
  
'Agrega algunos datos en LstListas
'For i = 1 To 12
'    LstListas.AddItem MonthName(i)
'Next
Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))
    
    
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

Private Sub LstListas_DragDrop(Source As Control, x As Single, y As Single)
 If Source Is LstSeleccionados Then
       If LstSeleccionados.ListIndex <> -1 Then
            LstListas.AddItem LstSeleccionados.List(LstSeleccionados.ListIndex)
            LstSeleccionados.RemoveItem LstSeleccionados.ListIndex
       End If
    End If
End Sub
Private Sub LstListas_MouseDown(Button As Integer, Shift As Integer, _
                            x As Single, y As Single)
    LstListas.Drag vbBeginDrag
End Sub
Private Sub LstSeleccionados_DragDrop(Source As Control, _
                           x As Single, y As Single)
      
    ' Si el control es el LstListas entonces..
    If Source Is LstListas Then
        If LstListas.ListIndex <> -1 Then
           LstSeleccionados.AddItem LstListas.List(LstListas.ListIndex)
           LstListas.RemoveItem LstListas.ListIndex
        End If
    End If
End Sub
Private Sub LstSeleccionados_MouseDown(Button As Integer, Shift As Integer, _
                                        x As Single, y As Single)
    LstSeleccionados.Drag vbBeginDrag
End Sub

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
Private Sub Cargar_Columnas()
On Error GoTo Control
    If CE(tdbtArchivo.Text) = "" Then Mensajes "Seleccione el archivo a importar": Exit Sub

    cmdImportarDatos.Enabled = False
    cmdSeleccionar.Enabled = False
    cmdImprimir.Enabled = False
    cmd_salir.Enabled = False
    Me.MousePointer = vbHourglass
    DoEvents
    
    Dim primeraletra, letra2 As Integer
    primeraletra = 0
    letra2 = 64
    
    Dim Columna As String
    Dim Valor As String
    Dim Ex As Object
    Dim Wb As Object
    Set Ex = CreateObject("Excel.Application")
'    Set Wb = Ex.Workbooks.Open(Me.dlgAbrirArchivo.FileName)
    Set Wb = Ex.Workbooks.Open(Me.tdbtArchivo.Text)
    sBaseDatos = gsBD
    'If CrearTablas(Me.dlgAbrirArchivo.FileName, pbAvance, lblAvance, cParametros) Then
        Dim VarSql As String
        Dim VarContFila As Long
            VarContFila = 0
            Set Sht = Wb.Worksheets("ENTIDAD")
        
        
            For VarContFila = 1 To 70000
                letra2 = letra2 + 1
                If primeraletra <> 0 Then
                Columna = Trim$(Chr(primeraletra)) & Chr(letra2)
                Else
                Columna = Chr(letra2)
                End If
                
                Valor = Trim$(Columna) & "1"
                If Sht.Range(Valor).Value <> "" Then
                    Dim Dato As String
                    Dato = Columna + "-" + Sht.Range(Valor).Value
                    LstListas.AddItem (Dato)
                    
                Else
                    Exit For
                End If
                
                If letra2 = 90 And primeraletra = 0 Then
                primeraletra = 65
                ElseIf Letra = 90 Then
                
                 primeraletra = primeraletra + 1
                 letra2 = 65
                End If
                
                
                
            Next VarContFila
                
        Ex.Save
        Ex.Quit
        'cn.Close
        'Set cn = Nothing
        Set Sht = Nothing
        Set Wb = Nothing
        Set Ex = Nothing
        
        
    'End If
  
    pbAvance.Value = 0
'    pbAvance.Max = 0
    pbAvance.Refresh
    DoEvents
    
    cmdImportarDatos.Enabled = True
    cmdSeleccionar.Enabled = True
    cmd_salir.Enabled = True
    cmdImprimir.Enabled = True
    Me.MousePointer = vbNormal
     ' Ex.Quit
'        cn.Close
        Set cn = Nothing
        Set Sht = Nothing
        Set Wb = Nothing
        Set Ex = Nothing
    
Exit Sub
Control:

'MsgBox "Error en la importación", vbCritical + vbSystemModal, "Mensaje: Sistema"
MsgBox Err.Description & vbCrLf & sSql & vbcrfl & "Fila: " & VarContFila, vbCritical + vbSystemModal, "Mensaje: Sistema"
'Resume
        Ex.Quit
'        cn.Close
        Set cn = Nothing
        Set Sht = Nothing
        Set Wb = Nothing
        Set Ex = Nothing

End Sub
Private Sub cmdImportarDatos_Click()
'On Error GoTo Control
'    If CE(tdbtArchivo.Text) = "" Then Mensajes "Seleccione el archivo a importar": Exit Sub
'
'    cmdImportarDatos.Enabled = False
'    cmdSeleccionar.Enabled = False
'    cmdImprimir.Enabled = False
'    cmd_salir.Enabled = False
'    Me.MousePointer = vbHourglass
'    DoEvents
'
'    Dim cParametros As String
'
'
'    Dim Ex As Object
'    Dim Wb As Object
'    Set Ex = CreateObject("Excel.Application")
''    Set Wb = Ex.Workbooks.Open(Me.dlgAbrirArchivo.FileName)
'    Set Wb = Ex.Workbooks.Open(Me.tdbtArchivo.Text)
'    sBaseDatos = gsBD
'    'If CrearTablas(Me.dlgAbrirArchivo.FileName, pbAvance, lblAvance, cParametros) Then
'        Dim VarSql As String
'        Dim VarContFila As Long
'        Dim cn As New ADODB.Connection
'        cn.Open gsCadenaConexion
'        cn.Execute ("set dateformat dmy")
'
'        If chkOpcion(0).Value = 1 Then
'
'            Call cn.Execute("if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_ENTIDAD]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf & _
'            "drop table [dbo].[ZIMP_ENTIDAD]")
'
'            VarSql = "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_ENTIDAD] (" & vbCrLf & _
'            "[Ent_cCodEntidad] char (5) NULL, " & vbCrLf & "[Ten_cTipoEntidad] char (1) NULL, " & vbCrLf & _
'            "[Ent_cPersona] nvarchar (255) NULL, " & vbCrLf & "[Ent_cDireccion] nvarchar (255) NULL, " & vbCrLf & _
'            "[Ent_nRuc] nvarchar (255) NULL, " & vbCrLf & "[Ent_cTipoDoc] nvarchar (255) NULL, " & vbCrLf & _
'            "[Ent_cFlagPersona] nvarchar (255) NULL)"
'
'            Call cn.Execute(VarSql)
'
'            VarContFila = 0
'            Set Sht = Wb.Worksheets("ENTIDAD")
'
'            For VarContFila = 2 To 70000
'                If Sht.Range("A" & VarContFila).Value = "" Then
'                    Exit For
'                End If
'            Next VarContFila
'
'            lblAvance.Caption = "IMPORTANDO A TABLA TEMPORAL - ENTIDAD"
'            Me.pbAvance.Min = 0
'            Me.pbAvance.Value = 0
'            Me.pbAvance.Max = VarContFila - 1
'
'            For VarContFila = 2 To 70000
'                If Sht.Range("A" & VarContFila).Value = "" Then Exit For
'                pbAvance.Value = pbAvance.Value + 1
'                pbAvance.Refresh
'                DoEvents
'                cn.BeginTrans
'                '*****SI ENCUENTRA UNA SOLA COMILLA EN LA GLOSA*****
'                If InStr(Sht.Range("C" & VarContFila).Value, "'") > 0 Then
'                    Sht.Range("C" & VarContFila).Value = Replace(Sht.Range("C" & VarContFila).Value, "'", "")
'                End If
'                If InStr(Sht.Range("D" & VarContFila).Value, "'") > 0 Then
'                    Sht.Range("D" & VarContFila).Value = Replace(Sht.Range("D" & VarContFila).Value, "'", "")
'                End If
'                cn.Execute ("Insert Into ZIMP_ENTIDAD(Ent_cCodEntidad, Ten_cTipoEntidad, Ent_cPersona, Ent_cDireccion, " & _
'                "Ent_nRuc, Ent_cTipoDoc, Ent_cFlagPersona) Values " & _
'                "('" & Sht.Range("A" & VarContFila).Value & "','" & Sht.Range("B" & VarContFila).Value & "','" & Sht.Range("C" & VarContFila).Value & _
'                "','" & Sht.Range("D" & VarContFila).Value & "','" & Sht.Range("E" & VarContFila).Value & "','" & Sht.Range("F" & VarContFila).Value & _
'                "','" & Sht.Range("G" & VarContFila).Value & "')")
'                cn.CommitTrans
'            Next VarContFila
'
'        End If
'
'        If chkOpcion(1).Value = 1 Then
'            Call cn.Execute("if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_TIPOCAMBIO]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf & _
'            "drop table [dbo].[ZIMP_TIPOCAMBIO]")
'
'            VarSql = "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_TIPOCAMBIO] (" & vbCrLf & _
'            "[Tca_dFecha] DateTime NULL, " & vbCrLf & "[Tca_cCodigoOrigen] nvarchar (255) NULL, " & vbCrLf & _
'            "[Tca_cCodigoDestino] nvarchar (255) NULL, " & vbCrLf & "[Tca_nCompra] nvarchar (255) NULL, " & vbCrLf & _
'            "[Tca_nVenta] nvarchar (255) NULL, " & vbCrLf & "[Tca_nVentaP] nvarchar (255) NULL" & vbCrLf & ")"
'            Call cn.Execute(VarSql)
'
'            VarContFila = 0
'            Set Sht = Wb.Worksheets("TIPOCAMBIO")
'
'            For VarContFila = 2 To 70000
'                If Sht.Range("A" & VarContFila).Value = "" Then
'                    Exit For
'                End If
'            Next VarContFila
'
'            lblAvance.Caption = "IMPORTANDO A TABLA TEMPORAL - TIPO DE CAMBIO"
'            Me.pbAvance.Min = 0
'            Me.pbAvance.Value = 0
'            Me.pbAvance.Max = VarContFila - 1
'
'            For VarContFila = 2 To 70000
'                If Sht.Range("A" & VarContFila).Value = "" Then Exit For
'                pbAvance.Value = pbAvance.Value + 1
'                pbAvance.Refresh
'                DoEvents
'                cn.BeginTrans
'                    cn.Execute ("Insert Into ZIMP_TIPOCAMBIO(Tca_dFecha, Tca_cCodigoOrigen, Tca_cCodigoDestino, Tca_nCompra, " & _
'                    "Tca_nVenta, Tca_nVentaP) Values " & _
'                    "('" & Format(Sht.Range("A" & VarContFila).Value, "dd/mm/yyyy") & "','" & Format(Sht.Range("B" & VarContFila).Value, "000") & "','" & Format(Sht.Range("C" & VarContFila).Value, "000") & _
'                    "','" & Sht.Range("D" & VarContFila).Value & "','" & Sht.Range("E" & VarContFila).Value & "','" & Sht.Range("F" & VarContFila).Value & "')")
'                cn.CommitTrans
'            Next VarContFila
'
'        End If
'
'        If chkOpcion(2).Value = 1 Then
'            Call cn.Execute("if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_APE_CAB]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf & _
'            "drop table [dbo].[ZIMP_APE_CAB]")
'
'            VarSql = "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_APE_CAB] (" & vbCrLf & _
'            "[Ase_cNummov] char (10) NULL, " & vbCrLf & "[Pan_cAnio] char (4) NULL, " & vbCrLf & _
'            "[Per_cPeriodo] char (2) NULL, " & vbCrLf & "[Lib_cTipoLibro] char (2) NULL, " & vbCrLf & _
'            "[Ase_nVoucher] char (10) NULL, " & vbCrLf & "[Ase_dFecha] nvarchar (255) NULL, " & vbCrLf & _
'            "[Ase_cGlosa] nvarchar (255) NULL, " & vbCrLf & "[Ase_cTipoMoneda] nvarchar (255) NULL" & vbCrLf & ")"
'            Call cn.Execute(VarSql)
'
'            VarContFila = 0
'            Set Sht = Wb.Worksheets("APE_CAB")
'
'            For VarContFila = 2 To 70000
'                If Sht.Range("A" & VarContFila).Value = "" Then
'                    Exit For
'                End If
'            Next VarContFila
'
'            lblAvance.Caption = "IMPORTANDO A TABLA TEMPORAL - APERTURACAB"
'            Me.pbAvance.Min = 0
'            Me.pbAvance.Value = 0
'            Me.pbAvance.Max = VarContFila - 1
'
'            For VarContFila = 2 To 70000
'                If Sht.Range("A" & VarContFila).Value = "" Then Exit For
'                pbAvance.Value = pbAvance.Value + 1
'                pbAvance.Refresh
'                DoEvents
'                cn.BeginTrans
'                    cn.Execute ("Insert Into ZIMP_APE_CAB(Ase_cNummov, Pan_cAnio, Per_cPeriodo, " & _
'                    "Lib_cTipoLibro, Ase_nVoucher, Ase_dFecha, Ase_cGlosa, Ase_cTipoMoneda) Values " & _
'                    "('" & Sht.Range("A" & VarContFila).Value & "','" & Sht.Range("B" & VarContFila).Value & "','" & Sht.Range("C" & VarContFila).Value & _
'                    "','" & Sht.Range("D" & VarContFila).Value & "','" & Format(Val(Sht.Range("E" & VarContFila).Value), "0000000000") & "','" & Sht.Range("F" & VarContFila).Value & _
'                    "','" & Sht.Range("G" & VarContFila).Value & "','" & Format(Sht.Range("H" & VarContFila).Value, "000") & "')")
'                cn.CommitTrans
'            Next VarContFila
'
'            Call cn.Execute("if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_APE_DET]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf & _
'            "drop table [dbo].[ZIMP_APE_DET]")
'
'            VarSql = "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_APE_DET] (" & vbCrLf & _
'            "[Ase_cNummov] char (10) NULL, " & vbCrLf & "[Pan_cAnio] char (4) NULL, " & vbCrLf & _
'            "[Per_cPeriodo] char (2) NULL, " & vbCrLf & "[Lib_cTipoLibro] char (2) NULL, " & vbCrLf & _
'            "[Ase_nVoucher] char (10) NULL, " & vbCrLf & "[Pla_cCuentaContable] varchar (12) NULL, " & vbCrLf & _
'            "[Asd_nItem] INT NULL, " & vbCrLf & "[Asd_cGlosa] nvarchar (255) NULL, " & vbCrLf & _
'            "[Asd_nDebeSoles] NUMERIC (14,3) NULL , " & vbCrLf & "[Asd_nHaberSoles] NUMERIC (14,3) NULL , " & vbCrLf & _
'            "[Asd_nTipoCambio] NUMERIC (14,3) NULL , " & vbCrLf & "[Asd_nDebeMonExt] NUMERIC (14,3) NULL , " & vbCrLf & _
'            "[Asd_nHaberMonExt] NUMERIC (14,3) NULL , " & vbCrLf & "[Cos_cCodigo] varchar (12) NULL, " & vbCrLf & _
'            "[Ten_cTipoEntidad] char (1) NULL, " & vbCrLf & "[Ent_cCodEntidad] char (5) NULL, " & vbCrLf & _
'            "[Asd_cTipoDoc] char (3) NULL, " & vbCrLf & "[Asd_dFecDoc] datetime null , " & vbCrLf & _
'            "[Asd_cSerieDoc] varchar (5) NULL, " & vbCrLf & "[Asd_cNumDoc] varchar (15) NULL, " & vbCrLf & _
'            "[Asd_dFecVen] datetime NULL, " & vbCrLf & "[Asd_cProvCanc] char (1) NULL, " & vbCrLf & _
'            "[Asd_cOperaTC] char (3) NULL, " & vbCrLf & "[Asd_cTipoMoneda] char (3) NULL" & vbCrLf & ")"
'            Call cn.Execute(VarSql)
'
'            VarContFila = 0
'            Set Sht = Wb.Worksheets("APE_DET")
'
'            For VarContFila = 2 To 70000
'                If Sht.Range("A" & VarContFila).Value = "" Then
'                    Exit For
'                End If
'            Next VarContFila
'            lblAvance.Caption = "IMPORTANDO A TABLA TEMPORAL - APERTURADET"
'            Me.pbAvance.Min = 0
'            Me.pbAvance.Value = 0
'            Me.pbAvance.Max = VarContFila - 1
'
'            For VarContFila = 2 To 70000
'                If Sht.Range("A" & VarContFila).Value = "" Then Exit For
'                pbAvance.Value = pbAvance.Value + 1
'                pbAvance.Refresh
'                DoEvents
'                cn.BeginTrans
'                sSql = "Insert Into ZIMP_APE_DET(Ase_cNummov, Pan_cAnio, Per_cPeriodo, " & _
'                    "Lib_cTipoLibro, Ase_nVoucher, Pla_cCuentaContable, Asd_nItem, Asd_cGlosa, Asd_nDebeSoles, Asd_nHaberSoles, Asd_nTipoCambio, " & _
'                    "Asd_nDebeMonExt, Asd_nHaberMonExt, Cos_cCodigo, Ten_cTipoEntidad, Ent_cCodEntidad, Asd_cTipoDoc, Asd_dFecDoc, Asd_cSerieDoc, Asd_cNumDoc, " & _
'                    "Asd_dFecVen, Asd_cProvCanc, Asd_cOperaTC, Asd_cTipoMoneda) Values " & _
'                    "('" & Sht.Range("A" & VarContFila).Value & "','" & Sht.Range("B" & VarContFila).Value & "','" & Sht.Range("C" & VarContFila).Value & _
'                    "','" & Sht.Range("D" & VarContFila).Value & "','" & Format(Val(Sht.Range("E" & VarContFila).Value), "0000000000") & "','" & Sht.Range("F" & VarContFila).Value & _
'                    "','" & Sht.Range("G" & VarContFila).Value & "','" & Sht.Range("H" & VarContFila).Value & "', " & Format(Val(Sht.Range("I" & VarContFila).Value), "0.00") & _
'                    ", " & Format(Val(Sht.Range("J" & VarContFila).Value), "0.00") & ", " & Format(Val(Sht.Range("K" & VarContFila).Value), "0.00") & ", " & Format(Val(Sht.Range("L" & VarContFila).Value), "0.00") & _
'                    ", " & Format(Val(Sht.Range("M" & VarContFila).Value), "0.00") & ", '" & Sht.Range("N" & VarContFila).Value & "', '" & Sht.Range("O" & VarContFila).Value & _
'                    "', '" & Sht.Range("P" & VarContFila).Value & "', '" & Sht.Range("Q" & VarContFila).Value & "', '" & Format(Sht.Range("R" & VarContFila).Value, "dd/MM/yyyy") & _
'                    "', '" & Sht.Range("S" & VarContFila).Value & "', '" & Sht.Range("T" & VarContFila).Value & "', '" & IIf(Sht.Range("U" & VarContFila).Value = "", "", Format(Sht.Range("U" & VarContFila).Value, "dd/mm/yyyy")) & _
'                    "', '" & Sht.Range("V" & VarContFila).Value & "', '" & Sht.Range("W" & VarContFila).Value & "', '" & Sht.Range("X" & VarContFila).Value & "')"
'                    Call cn.Execute(sSql)
'                cn.CommitTrans
'            Next VarContFila
'
'        End If
'
'        If chkOpcion(3).Value = 1 Then
'            Call cn.Execute("if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_COMPRAS_CAB]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf & _
'            "drop table [dbo].[ZIMP_COMPRAS_CAB]")
'
'            VarSql = "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_COMPRAS_CAB] (" & vbCrLf & _
'            "[Ase_cNummov] char (10) NULL, " & vbCrLf & _
'            "[Pan_cAnio] char (4) NULL, " & vbCrLf & _
'            "[Per_cPeriodo] char (2) NULL, " & vbCrLf & _
'            "[Lib_cTipoLibro] char (2) NULL, " & vbCrLf & _
'            "[Ase_nVoucher] char (10) NULL, " & vbCrLf & _
'            "[Ase_dFecha] DATETIME NULL, " & vbCrLf & _
'            "[Ase_cGlosa] nvarchar (255) NULL, " & vbCrLf & _
'            "[Ase_cTipoMoneda] nvarchar (255) NULL" & vbCrLf & _
'            ")"
'            Call cn.Execute(VarSql)
'
'            VarContFila = 0
'            Set Sht = Wb.Worksheets("COMPRAS_CAB")
'
'            For VarContFila = 2 To 70000
'                If Sht.Range("A" & VarContFila).Value = "" Then
'                    Exit For
'                End If
'            Next VarContFila
'            lblAvance.Caption = "IMPORTANDO A TABLA TEMPORAL - COMPRASCAB"
'            Me.pbAvance.Min = 0
'            Me.pbAvance.Value = 0
'            Me.pbAvance.Max = VarContFila - 1
''            cn.Execute ("set dateformat dmy")
'            For VarContFila = 2 To 70000
'                If Sht.Range("A" & VarContFila).Value = "" Then Exit For
'                pbAvance.Value = pbAvance.Value + 1
'                pbAvance.Refresh
'                DoEvents
'                cn.BeginTrans
'                    '*****SI ENCUENTRA UNA SOLA COMILLA EN LA GLOSA*****
'                    If InStr(Sht.Range("G" & VarContFila).Value, "'") > 0 Then
'                        Sht.Range("G" & VarContFila).Value = Replace(Sht.Range("G" & VarContFila).Value, "'", "")
'                    End If
'                    cn.Execute ("Insert Into ZIMP_COMPRAS_CAB(Ase_cNummov, Pan_cAnio, Per_cPeriodo, " & _
'                    "Lib_cTipoLibro, Ase_nVoucher, Ase_dFecha, Ase_cGlosa, Ase_cTipoMoneda) Values " & _
'                    "('" & Format(Val(Sht.Range("A" & VarContFila).Value), "0000000000") & "','" & _
'                    Sht.Range("B" & VarContFila).Value & "','" & Format(Val(Sht.Range("C" & VarContFila).Value), "00") & _
'                    "','" & Format(Val(Sht.Range("D" & VarContFila).Value), "00") & "','" & _
'                    Format(Val(Sht.Range("E" & VarContFila).Value), "0000000000") & "', " & _
'                    "'" & Sht.Range("F" & VarContFila).Value & "'" & _
'                    ",'" & Sht.Range("G" & VarContFila).Value & "','" & _
'                    Format(Val(Sht.Range("H" & VarContFila).Value), "000") & "')")
'                cn.CommitTrans
'            Next VarContFila
'
'            Call cn.Execute("if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_COMPRAS_DET]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf & _
'            "drop table [dbo].[ZIMP_COMPRAS_DET]")
'
'            VarSql = "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_COMPRAS_DET] (" & vbCrLf & _
'                        "[Ase_cNummov] char (10) NULL, " & vbCrLf & "[Pan_cAnio] char (4) NULL, " & vbCrLf & _
'                        "[Per_cPeriodo] char (2) NULL, " & vbCrLf & "[Lib_cTipoLibro] char (2) NULL, " & vbCrLf & _
'                        "[Ase_nVoucher] char (10) NULL, " & vbCrLf & "[Pla_cCuentaContable] varchar (12) NULL, " & vbCrLf & _
'                        "[Asd_nItem] INT NULL, " & vbCrLf & "[Asd_cGlosa] nvarchar (255) NULL, " & vbCrLf & _
'                        "[Asd_nDebeSoles] NUMERIC (14,3) NULL , " & vbCrLf & "[Asd_nHaberSoles] NUMERIC (14,3) NULL , " & vbCrLf & _
'                        "[Asd_nTipoCambio] NUMERIC (14,3) NULL , " & vbCrLf & "[Asd_nDebeMonExt] NUMERIC (14,3) NULL , " & vbCrLf & _
'                        "[Asd_nHaberMonExt] NUMERIC (14,3) NULL , " & vbCrLf & "[Cos_cCodigo] varchar (12) NULL, " & vbCrLf & _
'                        "[Ten_cTipoEntidad] char (1) NULL, " & vbCrLf & "[Ent_cCodEntidad] char (5) NULL, " & vbCrLf & _
'                        "[Asd_cTipoDoc] char (2) NULL, " & vbCrLf & "[Asd_dFecDoc] datetime NULL, " & vbCrLf & _
'                        "[Asd_cSerieDoc] varchar (5) NULL, " & vbCrLf & "[Asd_cNumDoc] varchar (15) NULL, " & vbCrLf & _
'                        "[Asd_dFecVen] datetime NULL, " & vbCrLf & "[Asd_cTipoDocRef] nvarchar (255) NULL, " & vbCrLf & _
'                        "[Asd_dFecDocRef] datetime NULL, " & vbCrLf & "[Asd_cSerieDocRef] nvarchar (255) NULL, " & vbCrLf & _
'                        "[Asd_cNumDocRef] nvarchar (255) NULL, " & vbCrLf & "[Asd_nMontoInafecto] nvarchar (255) NULL, " & vbCrLf & _
'                        "[Asd_cBaseImp] nvarchar (255) NULL, " & vbCrLf & "[Asd_cRetencion] nvarchar (255) NULL, " & vbCrLf & _
'                        "[Asd_dFechaSpot] datetime NULL, " & vbCrLf & "[Asd_cNumSpot] nvarchar (255) NULL, " & vbCrLf & _
'                        "[Asd_cProvCanc] char (1) NULL, " & vbCrLf & "[Asd_cOperaTC] char (3) NULL, " & vbCrLf & _
'                        "[Asd_cTipoMoneda] char (3) NULL, " & vbCrLf & "[Asd_cComprobante] nvarchar (255) NULL" & vbCrLf & _
'                        ")"
'            Call cn.Execute(VarSql)
'
'            VarContFila = 0
'            Set Sht = Wb.Worksheets("COMPRAS_DET")
'
'            For VarContFila = 2 To 70000
'                If Sht.Range("A" & VarContFila).Value = "" Then
'                    Exit For
'                End If
'            Next VarContFila
'
'            lblAvance.Caption = "IMPORTANDO A TABLA TEMPORAL - COMPRASDET"
'            Me.pbAvance.Min = 0
'            Me.pbAvance.Value = 0
'            Me.pbAvance.Max = VarContFila - 1
'
'            For VarContFila = 2 To 70000
'                pbAvance.Value = pbAvance.Value + 1
'                pbAvance.Refresh
'                DoEvents
'                If Sht.Range("A" & VarContFila).Value = "" Then Exit For
''                If VarContFila = 68 Then
''                    MsgBox "XXX"
''                End If
'                    '*****SI ENCUENTRA UNA SOLA COMILLA EN LA GLOSA*****
'                    If InStr(Sht.Range("H" & VarContFila).Value, "'") > 0 Then
'                        Sht.Range("H" & VarContFila).Value = Replace(Sht.Range("H" & VarContFila).Value, "'", "")
'                    End If
'                cn.BeginTrans
'                sSql = "Insert Into ZIMP_COMPRAS_DET(Ase_cNummov, Pan_cAnio, Per_cPeriodo, "
'                sSql = sSql & "Lib_cTipoLibro, Ase_nVoucher, Pla_cCuentaContable, Asd_nItem, Asd_cGlosa, Asd_nDebeSoles, Asd_nHaberSoles, Asd_nTipoCambio, "
'                sSql = sSql & "Asd_nDebeMonExt, Asd_nHaberMonExt, Cos_cCodigo, Ten_cTipoEntidad, Ent_cCodEntidad, Asd_cTipoDoc, Asd_dFecDoc, Asd_cSerieDoc, Asd_cNumDoc, "
'                sSql = sSql & "Asd_dFecVen, Asd_cTipoDocRef, Asd_dFecDocRef, Asd_cSerieDocRef, Asd_cNumDocRef, Asd_nMontoInafecto, Asd_cBaseImp, Asd_cRetencion, Asd_dFechaSpot, "
'                sSql = sSql & "Asd_cNumSpot, Asd_cProvCanc, Asd_cOperaTC, Asd_cTipoMoneda, Asd_cComprobante) Values "
'                sSql = sSql & "('" & Format(Val(Sht.Range("A" & VarContFila).Value), "0000000000") & "','" & Sht.Range("B" & VarContFila).Value & "','" & Format(Val(Sht.Range("C" & VarContFila).Value), "00")
'                sSql = sSql & "','" & Format(Val(Sht.Range("D" & VarContFila).Value), "00") & "','" & Format(Val(Sht.Range("E" & VarContFila).Value), "0000000000") & "','" & Sht.Range("F" & VarContFila).Value
'                sSql = sSql & "','" & Sht.Range("G" & VarContFila).Value & "','" & Sht.Range("H" & VarContFila).Value & "', " & Sht.Range("I" & VarContFila).Value
'                sSql = sSql & ", " & NE(Sht.Range("J" & VarContFila).Value) & ", " & NE(Sht.Range("K" & VarContFila).Value) & ", " & NE(Sht.Range("L" & VarContFila).Value)
'                sSql = sSql & ", " & NE(Sht.Range("M" & VarContFila).Value) & ", '" & Sht.Range("N" & VarContFila).Value & "', '" & Sht.Range("O" & VarContFila).Value
'                sSql = sSql & "', '" & IIf(LTrim(RTrim(Sht.Range("P" & VarContFila).Value)) = "", "", Format(Val(Sht.Range("P" & VarContFila).Value), "00000")) & "', '" & IIf(Sht.Range("Q" & VarContFila).Value = "", "", Format(Val(Sht.Range("Q" & VarContFila).Value), "00")) & "', '" & IIf(Sht.Range("R" & VarContFila).Value = "", "", Format(Sht.Range("R" & VarContFila).Value, "dd/MM/yyyy"))
'                sSql = sSql & "', '" & Sht.Range("S" & VarContFila).Value & "', '" & Sht.Range("T" & VarContFila).Value & "', '" & IIf(Sht.Range("U" & VarContFila).Value = "", "", Format(Sht.Range("U" & VarContFila).Value, "dd/MM/yyyy"))
'                sSql = sSql & "', '" & Sht.Range("V" & VarContFila).Value & "', '" & IIf(Sht.Range("W" & VarContFila).Value = "", "", Format(Sht.Range("W" & VarContFila).Value, "dd/MM/yyyy")) & "', '" & Sht.Range("X" & VarContFila).Value
'                sSql = sSql & "', '" & Sht.Range("Y" & VarContFila).Value & "', '" & Sht.Range("Z" & VarContFila).Value & "', '" & Sht.Range("AA" & VarContFila).Value
'                sSql = sSql & "', '" & Sht.Range("AB" & VarContFila).Value & "', '" & IIf(Sht.Range("AC" & VarContFila).Value = "", "", Format(Sht.Range("AC" & VarContFila).Value, "dd/mm/yyyy")) & "', '" & Sht.Range("AD" & VarContFila).Value
'                sSql = sSql & "', '" & Sht.Range("AE" & VarContFila).Value & "', '" & Sht.Range("AF" & VarContFila).Value & "', '" & Format(Val(Sht.Range("AG" & VarContFila).Value), "000")
'                sSql = sSql & "', '" & Sht.Range("AH" & VarContFila).Value & "')"
'                    cn.Execute (sSql)
'                cn.CommitTrans
'            Next VarContFila
'        End If
'
'        If chkOpcion(4).Value = 1 Then
'            Call cn.Execute("if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_VENTAS_CAB]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf & _
'            "drop table [dbo].[ZIMP_VENTAS_CAB]")
'            VarSql = "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_VENTAS_CAB] (" & vbCrLf & _
'            "[Ase_cNummov] char (10) NULL, " & vbCrLf & "[Pan_cAnio] char (4) NULL, " & vbCrLf & _
'            "[Per_cPeriodo] char (2) NULL, " & vbCrLf & "[Lib_cTipoLibro] char (2) NULL, " & vbCrLf & _
'            "[Ase_nVoucher] char (10) NULL, " & vbCrLf & "[Ase_dFecha] nvarchar (255) NULL, " & vbCrLf & _
'            "[Ase_cGlosa] nvarchar (255) NULL, " & vbCrLf & "[Ase_cTipoMoneda] nvarchar (255) NULL" & vbCrLf & ")"
'            Call cn.Execute(VarSql)
'
'            VarContFila = 0
'            Set Sht = Wb.Worksheets("VENTAS_CAB")
'
'            For VarContFila = 2 To 70000
'                If Sht.Range("A" & VarContFila).Value = "" Then
'                    Exit For
'                End If
'            Next VarContFila
'            lblAvance.Caption = "IMPORTANDO A TABLA TEMPORAL - VENTASCAB"
'            Me.pbAvance.Min = 0
'            Me.pbAvance.Value = 0
'            Me.pbAvance.Max = VarContFila - 1
'
'            For VarContFila = 2 To 70000
'                If Sht.Range("A" & VarContFila).Value = "" Then Exit For
'                pbAvance.Value = pbAvance.Value + 1
'                pbAvance.Refresh
'                DoEvents
'                cn.BeginTrans
'                    '*****SI ENCUENTRA UNA SOLA COMILLA EN LA GLOSA*****
'                    If InStr(Sht.Range("G" & VarContFila).Value, "'") > 0 Then
'                        Sht.Range("G" & VarContFila).Value = Replace(Sht.Range("G" & VarContFila).Value, "'", "")
'                    End If
'                    cn.Execute ("Insert Into ZIMP_VENTAS_CAB(Ase_cNummov, Pan_cAnio, Per_cPeriodo, " & _
'                    "Lib_cTipoLibro, Ase_nVoucher, Ase_dFecha, Ase_cGlosa, Ase_cTipoMoneda) Values " & _
'                    "('" & Format(Val(Sht.Range("A" & VarContFila).Value), "0000000000") & "','" & _
'                    Sht.Range("B" & VarContFila).Value & "','" & _
'                    Format(Val(Sht.Range("C" & VarContFila).Value), "00") & _
'                    "','" & Format(Val(Sht.Range("D" & VarContFila).Value), "00") & "','" & _
'                    Format(Sht.Range("E" & VarContFila).Value, "0000000000") & "','" & _
'                    Format(Sht.Range("F" & VarContFila).Value, "dd/mm/yyyy") & _
'                    "','" & Sht.Range("G" & VarContFila).Value & "','" & _
'                    Format(Sht.Range("H" & VarContFila).Value, "000") & "')")
'                cn.CommitTrans
'            Next VarContFila
'
'            Call cn.Execute("if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_VENTAS_DET]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf & _
'            "drop table [dbo].[ZIMP_VENTAS_DET]")
'
'            VarSql = "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_VENTAS_DET] (" & vbCrLf & _
'                        "[Ase_cNummov] char (10) NULL, " & vbCrLf & "[Pan_cAnio] char (4) NULL, " & vbCrLf & _
'                        "[Per_cPeriodo] char (2) NULL, " & vbCrLf & "[Lib_cTipoLibro] char (2) NULL, " & vbCrLf & _
'                        "[Ase_nVoucher] char (10) NULL, " & vbCrLf & "[Pla_cCuentaContable] varchar (12) NULL, " & vbCrLf & _
'                        "[Asd_nItem] INT NULL, " & vbCrLf & "[Asd_cGlosa] nvarchar (255) NULL, " & vbCrLf & _
'                        "[Asd_nDebeSoles] NUMERIC (14,3) NULL , " & vbCrLf & "[Asd_nHaberSoles] NUMERIC (14,3) NULL , " & vbCrLf & _
'                        "[Asd_nTipoCambio] NUMERIC (14,3) NULL , " & vbCrLf & "[Asd_nDebeMonExt] NUMERIC (14,3) NULL , " & vbCrLf & _
'                        "[Asd_nHaberMonExt] NUMERIC (14,3) NULL , " & vbCrLf & "[Cos_cCodigo] varchar (12) NULL, " & vbCrLf & _
'                        "[Ten_cTipoEntidad] char (1) NULL, " & vbCrLf & "[Ent_cCodEntidad] char (5) NULL, " & vbCrLf & _
'                        "[Asd_cTipoDoc] char (2) NULL, " & vbCrLf & "[Asd_dFecDoc] varchar(10) NULL, " & vbCrLf & _
'                        "[Asd_cSerieDoc] varchar (5) NULL, " & vbCrLf & "[Asd_cNumDoc] varchar (15) NULL, " & vbCrLf & _
'                        "[Asd_dFecVen] varchar(10) NULL, " & vbCrLf & "[Asd_nMontoInafecto] nvarchar (255) NULL, " & vbCrLf & _
'                        "[Asd_cBaseImp] nvarchar (255) NULL, " & vbCrLf & "[Asd_cProvCanc] char (1) NULL, " & vbCrLf & _
'                        "[Asd_cOperaTC] char (3) NULL, " & vbCrLf & "[Asd_cTipoMoneda] char (3) NULL" & vbCrLf & ")"
'
'            Call cn.Execute(VarSql)
'
'            VarContFila = 0
'            Set Sht = Wb.Worksheets("VENTAS_DET")
'
'            For VarContFila = 2 To 70000
'                If Sht.Range("A" & VarContFila).Value = "" Then
'                    Exit For
'                End If
'            Next VarContFila
'
'            lblAvance.Caption = "IMPORTANDO A TABLA TEMPORAL - VENTASDET"
'            Me.pbAvance.Min = 0
'            Me.pbAvance.Value = 0
'            Me.pbAvance.Max = VarContFila - 1
'
'            For VarContFila = 2 To 70000
'                If Sht.Range("A" & VarContFila).Value = "" Then Exit For
'                cn.BeginTrans
'                pbAvance.Value = pbAvance.Value + 1
'                pbAvance.Refresh
'                DoEvents
'                    '*****SI ENCUENTRA UNA SOLA COMILLA EN LA GLOSA*****
'                    If InStr(Sht.Range("H" & VarContFila).Value, "'") > 0 Then
'                        Sht.Range("H" & VarContFila).Value = Replace(Sht.Range("H" & VarContFila).Value, "'", "")
'                    End If
'
'                    sSql = "Insert Into ZIMP_VENTAS_DET(Ase_cNummov, Pan_cAnio, Per_cPeriodo, "
'                    sSql = sSql & "Lib_cTipoLibro, Ase_nVoucher, Pla_cCuentaContable, Asd_nItem, Asd_cGlosa, Asd_nDebeSoles, Asd_nHaberSoles, Asd_nTipoCambio, "
'                    sSql = sSql & "Asd_nDebeMonExt, Asd_nHaberMonExt, Cos_cCodigo, Ten_cTipoEntidad, Ent_cCodEntidad, Asd_cTipoDoc, Asd_dFecDoc, Asd_cSerieDoc, Asd_cNumDoc, "
'                    sSql = sSql & "Asd_dFecVen, Asd_nMontoInafecto, Asd_cBaseImp, Asd_cProvCanc, Asd_cOperaTC, Asd_cTipoMoneda) Values "
'                    sSql = sSql & "('" & Format(Val(Sht.Range("A" & VarContFila).Value), "0000000000") & "','"
'                    sSql = sSql & Sht.Range("B" & VarContFila).Value & "','"
'                    sSql = sSql & Format(Val(Sht.Range("C" & VarContFila).Value), "00")
'                    sSql = sSql & "','" & Format(Val(Sht.Range("D" & VarContFila).Value), "00") & "','"
'                    sSql = sSql & Format(Val(Sht.Range("E" & VarContFila).Value), "0000000000") & "','"
'                    sSql = sSql & Sht.Range("F" & VarContFila).Value
'                    sSql = sSql & "','" & Sht.Range("G" & VarContFila).Value & "','"
'                    sSql = sSql & Sht.Range("H" & VarContFila).Value & "', "
'                    sSql = sSql & NE(Sht.Range("I" & VarContFila).Value)
'                    sSql = sSql & ", " & NE(Sht.Range("J" & VarContFila).Value) & ", "
'                    sSql = sSql & NE(Sht.Range("K" & VarContFila).Value) & ", "
'                    sSql = sSql & NE(Sht.Range("L" & VarContFila).Value)
'                    sSql = sSql & ", " & NE(Sht.Range("M" & VarContFila).Value) & ", '"
'                    sSql = sSql & Sht.Range("N" & VarContFila).Value & "', '"
'                    sSql = sSql & Sht.Range("O" & VarContFila).Value
'                    sSql = sSql & "', '" & IIf(LTrim(RTrim(Sht.Range("P" & VarContFila).Value)) = "", "", Format(Val(Sht.Range("P" & VarContFila).Value), "00000")) & "', '"
'                    sSql = sSql & IIf(LTrim(RTrim(Sht.Range("Q" & VarContFila).Value)) = "", "", Format(Sht.Range("Q" & VarContFila).Value, "00")) & "', '"
'                    sSql = sSql & IIf(LTrim(RTrim(Sht.Range("R" & VarContFila).Value)) = "", "", Format(CDate(Sht.Range("R" & VarContFila).Value), "dd/mm/yyyy"))
'                    sSql = sSql & "', '" & LTrim(RTrim(Sht.Range("S" & VarContFila).Value)) & "',"
'                    sSql = sSql & "'" & LTrim(RTrim(Sht.Range("T" & VarContFila).Value)) & "',"
'
'                    If LTrim(RTrim(Sht.Range("U" & VarContFila).Value)) = "" Then
'                    sSql = sSql & "''"
'
'                    Else
'                    sSql = sSql & "'" & Format(CDate(Sht.Range("U" & VarContFila).Value), "dd/mm/yyyy") & "'"
'                    End If
'
'
'                    'sSql = sSql & "'" & IIf(LTrim(RTrim(Sht.Range("U" & VarContFila).Value)) = "", "", Format(CDate(Sht.Range("U" & VarContFila).Value), "dd/mm/yyyy"))
'                    'sSql = sSql & "'" & IIf(LTrim(RTrim(Sht.Range("U" & VarContFila).Value)) = "", "", Format(CDate(Sht.Range("U" & VarContFila).Value), "dd/mm/yyyy"))
'                    sSql = sSql & ", '" & Sht.Range("V" & VarContFila).Value & "', '" & Sht.Range("W" & VarContFila).Value & "', '" & Sht.Range("X" & VarContFila).Value
'                    sSql = sSql & "', '" & Sht.Range("Y" & VarContFila).Value & "', '" & Format(Val(Sht.Range("Z" & VarContFila).Value), "000") & "')"
'                    Call cn.Execute(sSql)
'                cn.CommitTrans
'            Next VarContFila
'        End If
'
'        If chkOpcion(5).Value = 1 Then
'            Call cn.Execute("if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_CAJAING_CAB]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf & _
'            "drop table [dbo].[ZIMP_CAJAING_CAB]")
'            VarSql = "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_CAJAING_CAB] (" & vbCrLf & _
'                        "[Ase_cNummov] char (10) NULL, " & vbCrLf & _
'                        "[Pan_cAnio] char (4) NULL, " & vbCrLf & _
'                        "[Per_cPeriodo] char (2) NULL, " & vbCrLf & _
'                        "[Lib_cTipoLibro] char (2) NULL, " & vbCrLf & _
'                        "[Ase_nVoucher] char (10) NULL, " & vbCrLf & _
'                        "[Ase_dFecha] nvarchar (255) NULL, " & vbCrLf & _
'                        "[Ase_cGlosa] nvarchar (255) NULL, " & vbCrLf & _
'                        "[Ase_cTipoMoneda] nvarchar (255) NULL" & vbCrLf & _
'                        ")"
'
'            Call cn.Execute(VarSql)
'
'            VarContFila = 0
'            Set Sht = Wb.Worksheets("CAJAING_CAB")
'
'            For VarContFila = 2 To 70000
'                If Sht.Range("A" & VarContFila).Value = "" Then
'                    Exit For
'                End If
'            Next VarContFila
'
'            lblAvance.Caption = "IMPORTANDO A TABLA TEMPORAL - CAJAINGCAB"
'            Me.pbAvance.Min = 0
'            Me.pbAvance.Value = 0
'            Me.pbAvance.Max = VarContFila - 1
'
'            For VarContFila = 2 To 70000
'                If Sht.Range("A" & VarContFila).Value = "" Then Exit For
'                pbAvance.Value = pbAvance.Value + 1
'                pbAvance.Refresh
'                DoEvents
'                cn.BeginTrans
'                    cn.Execute ("Insert Into ZIMP_CAJAING_CAB(Ase_cNummov, Pan_cAnio, Per_cPeriodo, " & _
'                    "Lib_cTipoLibro, Ase_nVoucher, Ase_dFecha, Ase_cGlosa, Ase_cTipoMoneda) Values " & _
'                    "('" & Format(Val(Sht.Range("A" & VarContFila).Value), "0000000000") & "','" & Sht.Range("B" & VarContFila).Value & "','" & Format(Val(Sht.Range("C" & VarContFila).Value), "00") & _
'                    "','" & Format(Val(Sht.Range("D" & VarContFila).Value), "00") & "','" & Format(Val(Sht.Range("E" & VarContFila).Value), "0000000000") & "','" & IIf(Sht.Range("F" & VarContFila).Value = "", "", Format(Sht.Range("F" & VarContFila).Value, "dd/mm/yyyy")) & _
'                    "','" & Sht.Range("G" & VarContFila).Value & "','" & Format(Val(Sht.Range("H" & VarContFila).Value), "000") & "')")
'                cn.CommitTrans
'            Next VarContFila
'
'            Call cn.Execute("if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_CAJAING_DET]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf & _
'            "drop table [dbo].[ZIMP_CAJAING_DET]")
'
'            VarSql = "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_CAJAING_DET] (" & vbCrLf & _
'            "[Ase_cNummov] char (10) NULL, " & vbCrLf & "[Pan_cAnio] char (4) NULL, " & vbCrLf & _
'            "[Per_cPeriodo] char (2) NULL, " & vbCrLf & "[Lib_cTipoLibro] char (2) NULL, " & vbCrLf & _
'            "[Ase_nVoucher] char (10) NULL, " & vbCrLf & "[Pla_cCuentaContable] varchar (12) NULL, " & vbCrLf & _
'            "[Asd_nItem] INT NULL, " & vbCrLf & "[Asd_cGlosa] nvarchar (255) NULL, " & vbCrLf & _
'            "[Asd_nDebeSoles] NUMERIC (14,3) NULL , " & vbCrLf & "[Asd_nHaberSoles] NUMERIC (14,3) NULL , " & vbCrLf & _
'            "[Asd_nTipoCambio] NUMERIC (14,3) NULL , " & vbCrLf & "[Asd_nDebeMonExt] NUMERIC (14,3) NULL , " & vbCrLf & _
'            "[Asd_nHaberMonExt] NUMERIC (14,3) NULL , " & vbCrLf & "[Cos_cCodigo] varchar (12) NULL, " & vbCrLf & _
'            "[Ten_cTipoEntidad] char (1) NULL, " & vbCrLf & "[Ent_cCodEntidad] char (5) NULL, " & vbCrLf & _
'            "[Asd_cTipoDoc] char (2) NULL, " & vbCrLf & "[Asd_dFecDoc] varchar(10) NULL, " & vbCrLf & _
'            "[Asd_cSerieDoc] varchar (5) NULL, " & vbCrLf & "[Asd_cNumDoc] varchar (15) NULL, " & vbCrLf & _
'            "[Asd_dFecVen] varchar(10) NULL, " & "[Asd_cTipoDocRef] nvarchar(255) NULL, " & vbCrLf & _
'            "[Asd_dFecDocRef] varchar(10) NULL, " & vbCrLf & "[Asd_cSerieDocRef] nvarchar (255) NULL, " & vbCrLf & _
'            "[Asd_cNumDocRef] nvarchar (255) NULL, " & vbCrLf & "[Asd_cRetencion] nvarchar (255) NULL, " & vbCrLf & _
'            "[Asd_cProvCanc] char (1) NULL, " & vbCrLf & "[Asd_cOperaTC] char (3) NULL, " & vbCrLf & _
'            "[Asd_cTipoMoneda] char (3) NULL, " & vbCrLf & "[Tra_cCodigo] nvarchar (255) NULL, " & vbCrLf & _
'            "[Asd_cFormaPago] nvarchar (255) NULL" & vbCrLf & ")"
'
'            Call cn.Execute(VarSql)
'
'            VarContFila = 0
'            Set Sht = Wb.Worksheets("CAJAING_DET")
'
'            For VarContFila = 2 To 70000
'                If Sht.Range("A" & VarContFila).Value = "" Then
'                    Exit For
'                End If
'            Next VarContFila
'
'            lblAvance.Caption = "IMPORTANDO A TABLA TEMPORAL - CAJAINGDET"
'            Me.pbAvance.Min = 0
'            Me.pbAvance.Value = 0
'            Me.pbAvance.Max = VarContFila - 1
'
'            For VarContFila = 2 To 70000
'                If Sht.Range("A" & VarContFila).Value = "" Then Exit For
'                pbAvance.Value = pbAvance.Value + 1
'                pbAvance.Refresh
'                DoEvents
'                cn.BeginTrans
'                    sSql = ""
'                    'cn.Execute (
'                    sSql = sSql & "Insert Into ZIMP_CAJAING_DET(Ase_cNummov, Pan_cAnio, Per_cPeriodo, "
'                    sSql = sSql & "Lib_cTipoLibro, Ase_nVoucher, Pla_cCuentaContable, Asd_nItem, Asd_cGlosa, Asd_nDebeSoles, Asd_nHaberSoles, Asd_nTipoCambio, "
'                    sSql = sSql & "Asd_nDebeMonExt, Asd_nHaberMonExt, Cos_cCodigo, Ten_cTipoEntidad, Ent_cCodEntidad, Asd_cTipoDoc, Asd_dFecDoc, Asd_cSerieDoc, Asd_cNumDoc, "
'                    sSql = sSql & "Asd_dFecVen, Asd_cTipoDocRef, Asd_dFecDocRef, Asd_cSerieDocRef, Asd_cNumDocRef, Asd_cRetencion, Asd_cProvCanc, Asd_cOperaTC, "
'                    sSql = sSql & "Asd_cTipoMoneda, Tra_cCodigo, Asd_cFormaPago) Values "
'                    sSql = sSql & "('" & Format(Val(Sht.Range("A" & VarContFila).Value), "0000000000") & "','" & Sht.Range("B" & VarContFila).Value & "','" & IIf(LTrim(RTrim(Sht.Range("C" & VarContFila).Value)) = "", "", Format(Val(Sht.Range("C" & VarContFila).Value), "00"))
'                    sSql = sSql & "','" & Format(Val(Sht.Range("D" & VarContFila).Value), "00") & "','" & Format(Val(Sht.Range("E" & VarContFila).Value), "0000000000") & "','" & Sht.Range("F" & VarContFila).Value
'                    sSql = sSql & "','" & Sht.Range("G" & VarContFila).Value & "','" & Sht.Range("H" & VarContFila).Value & "', " & NE(Sht.Range("I" & VarContFila).Value)
'                    sSql = sSql & ", " & NE(Sht.Range("J" & VarContFila).Value) & ", " & NE(Sht.Range("K" & VarContFila).Value) & ", " & NE(Sht.Range("L" & VarContFila).Value)
'                    sSql = sSql & ", " & NE(Sht.Range("M" & VarContFila).Value) & ", '" & Sht.Range("N" & VarContFila).Value & "', '" & Sht.Range("O" & VarContFila).Value
'                    sSql = sSql & "', '" & IIf(LTrim(RTrim(Sht.Range("P" & VarContFila).Value)) = "", "", Format(Val(Sht.Range("P" & VarContFila).Value), "00000")) & "', '" & IIf(LTrim(RTrim(Sht.Range("Q" & VarContFila).Value)) = "", "", Format(Val(Sht.Range("Q" & VarContFila).Value), "00")) & "', '" & IIf(LTrim(RTrim(Sht.Range("R" & VarContFila).Value)) = "", "", Format(Sht.Range("R" & VarContFila).Value, "dd/mm/yyyy"))
'                    sSql = sSql & "', '" & IIf(LTrim(RTrim(Sht.Range("S" & VarContFila).Value)) = "", "", Format(Val(Sht.Range("S" & VarContFila).Value), "000")) & "', '" & IIf(LTrim(RTrim(Sht.Range("T" & VarContFila).Value)) = "", "", Sht.Range("T" & VarContFila).Value) & "', '" & IIf(LTrim(RTrim(Sht.Range("U" & VarContFila).Value)) = "", "", Format(Sht.Range("U" & VarContFila).Value, "dd/mm/yyyy"))
'                    sSql = sSql & "', '" & IIf(LTrim(RTrim(Sht.Range("V" & VarContFila).Value)) = "", "", Format(Val(Sht.Range("V" & VarContFila).Value), "00")) & "', '" & IIf(LTrim(RTrim(Sht.Range("W" & VarContFila).Value)) = "", "", Format(Sht.Range("W" & VarContFila).Value, "dd/mm/yyyy")) & "', '" & Sht.Range("X" & VarContFila).Value & "', '" & Sht.Range("Y" & VarContFila).Value
'                    sSql = sSql & "', '" & Sht.Range("Z" & VarContFila).Value & "', '" & Sht.Range("AA" & VarContFila).Value & "', '" & Sht.Range("AB" & VarContFila).Value
'                    sSql = sSql & "', '" & IIf(LTrim(RTrim(Sht.Range("AC" & VarContFila).Value)) = "", "", Format(Val(Sht.Range("AC" & VarContFila).Value), "000")) & "', '" & Sht.Range("AD" & VarContFila).Value & "', '" & Sht.Range("AE" & VarContFila).Value
'                    sSql = sSql & "'"
'                    cn.Execute (sSql)
'                cn.CommitTrans
'            Next VarContFila
'
'        End If
'
'        If chkOpcion(6).Value = 1 Then
'            Call cn.Execute("if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_CAJAEGR_CAB]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf & _
'            "drop table [dbo].[ZIMP_CAJAEGR_CAB]")
'            VarSql = "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_CAJAEGR_CAB] (" & vbCrLf & _
'                    "[Ase_cNummov] char (10) NULL, " & vbCrLf & "[Pan_cAnio] char (4) NULL, " & vbCrLf & _
'                    "[Per_cPeriodo] char (2) NULL, " & vbCrLf & "[Lib_cTipoLibro] char (2) NULL, " & vbCrLf & _
'                    "[Ase_nVoucher] char (10) NULL, " & vbCrLf & "[Ase_dFecha] nvarchar (255) NULL, " & vbCrLf & _
'                    "[Ase_cGlosa] nvarchar (255) NULL, " & vbCrLf & "[Ase_cTipoMoneda] nvarchar (255) NULL" & vbCrLf & ")"
'            Call cn.Execute(VarSql)
'
'            VarContFila = 0
'            Set Sht = Wb.Worksheets("CAJAEGR_CAB")
'
'            For VarContFila = 2 To 70000
'                If Sht.Range("A" & VarContFila).Value = "" Then
'                    Exit For
'                End If
'            Next VarContFila
'
'            lblAvance.Caption = "IMPORTANDO A TABLA TEMPORAL - CAJAEGRESOCAB"
'            Me.pbAvance.Min = 0
'            Me.pbAvance.Value = 0
'            Me.pbAvance.Max = VarContFila - 1
'
'            For VarContFila = 2 To 70000
'                If Sht.Range("A" & VarContFila).Value = "" Then Exit For
'                pbAvance.Value = pbAvance.Value + 1
'                pbAvance.Refresh
'                DoEvents
'                cn.BeginTrans
'                    cn.Execute ("Insert Into ZIMP_CAJAEGR_CAB(Ase_cNummov, Pan_cAnio, Per_cPeriodo, " & _
'                    "Lib_cTipoLibro, Ase_nVoucher, Ase_dFecha, Ase_cGlosa, Ase_cTipoMoneda) Values " & _
'                    "('" & Format(Val(Sht.Range("A" & VarContFila).Value), "0000000000") & "','" & Sht.Range("B" & VarContFila).Value & "','" & Format(Val(Sht.Range("C" & VarContFila).Value), "00") & _
'                    "','" & Format(Sht.Range("D" & VarContFila).Value, "00") & "','" & Format(Val(Sht.Range("E" & VarContFila).Value), "0000000000") & "','" & Sht.Range("F" & VarContFila).Value & _
'                    "','" & Sht.Range("G" & VarContFila).Value & "','" & Format(Val(Sht.Range("H" & VarContFila).Value), "000") & "')")
'                cn.CommitTrans
'            Next VarContFila
'
'            Call cn.Execute("if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_CAJAEGR_DET]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf & _
'            "drop table [dbo].[ZIMP_CAJAEGR_DET]")
'
'            VarSql = "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_CAJAEGR_DET] (" & vbCrLf & _
'            "[Ase_cNummov] char (10) NULL, " & vbCrLf & "[Pan_cAnio] char (4) NULL, " & vbCrLf & _
'            "[Per_cPeriodo] char (2) NULL, " & vbCrLf & "[Lib_cTipoLibro] char (2) NULL, " & vbCrLf & _
'            "[Ase_nVoucher] char (10) NULL, " & vbCrLf & "[Pla_cCuentaContable] varchar (12) NULL, " & vbCrLf & _
'            "[Asd_nItem] INT NULL, " & vbCrLf & "[Asd_cGlosa] nvarchar (255) NULL, " & vbCrLf & _
'            "[Asd_nDebeSoles] NUMERIC (14,3) NULL , " & vbCrLf & "[Asd_nHaberSoles] NUMERIC (14,3) NULL , " & vbCrLf & _
'            "[Asd_nTipoCambio] NUMERIC (14,3) NULL , " & vbCrLf & "[Asd_nDebeMonExt] NUMERIC (14,3) NULL , " & vbCrLf & _
'            "[Asd_nHaberMonExt] NUMERIC (14,3) NULL , " & vbCrLf & "[Cos_cCodigo] varchar (12) NULL, " & vbCrLf & _
'            "[Ten_cTipoEntidad] char (1) NULL, " & vbCrLf & "[Ent_cCodEntidad] char (5) NULL, " & vbCrLf & _
'            "[Asd_cTipoDoc] char (2) NULL, " & vbCrLf & "[Asd_dFecDoc] datetime NULL, " & vbCrLf & _
'            "[Asd_cSerieDoc] varchar (5) NULL, " & vbCrLf & "[Asd_cNumDoc] varchar (15) NULL, " & vbCrLf & _
'            "[Asd_dFecVen] datetime NULL, " & vbCrLf & "[Asd_cTipoDocRef] nvarchar (255) NULL, " & vbCrLf & _
'            "[Asd_dFecDocRef] datetime NULL, " & vbCrLf & "[Asd_cSerieDocRef] nvarchar (255) NULL, " & vbCrLf & _
'            "[Asd_cNumDocRef] nvarchar (255) NULL, " & vbCrLf & "[Asd_cRetencion] nvarchar (255) NULL, " & vbCrLf & _
'            "[Asd_cProvCanc] char (1) NULL, " & vbCrLf & "[Asd_cOperaTC] char (3) NULL, " & vbCrLf & _
'            "[Asd_cTipoMoneda] char (3) NULL, " & vbCrLf & "[Tra_cCodigo] nvarchar (255) NULL, " & vbCrLf & _
'            "[Asd_cFormaPago] nvarchar (255) NULL" & vbCrLf & ")"
'            Call cn.Execute(VarSql)
'
'            VarContFila = 0
'            Set Sht = Wb.Worksheets("CAJAEGR_DET")
'
'            For VarContFila = 2 To 70000
'                If Sht.Range("A" & VarContFila).Value = "" Then
'                    Exit For
'                End If
'            Next VarContFila
'
'            lblAvance.Caption = "IMPORTANDO A TABLA TEMPORAL - CAJAEGRESODET"
'            Me.pbAvance.Min = 0
'            Me.pbAvance.Value = 0
'            Me.pbAvance.Max = VarContFila - 1
'
'            For VarContFila = 2 To 70000
'                If Sht.Range("A" & VarContFila).Value = "" Then Exit For
'                pbAvance.Value = pbAvance.Value + 1
'                pbAvance.Refresh
'                DoEvents
'                cn.BeginTrans
'                    cn.Execute ("Insert Into ZIMP_CAJAEGR_DET(Ase_cNummov, Pan_cAnio, Per_cPeriodo, " & _
'                    "Lib_cTipoLibro, Ase_nVoucher, Pla_cCuentaContable, Asd_nItem, Asd_cGlosa, Asd_nDebeSoles, Asd_nHaberSoles, Asd_nTipoCambio, " & _
'                    "Asd_nDebeMonExt, Asd_nHaberMonExt, Cos_cCodigo, Ten_cTipoEntidad, Ent_cCodEntidad, Asd_cTipoDoc, Asd_dFecDoc, Asd_cSerieDoc, Asd_cNumDoc, " & _
'                    "Asd_dFecVen, Asd_cTipoDocRef, Asd_dFecDocRef, Asd_cSerieDocRef, Asd_cNumDocRef, Asd_cRetencion, Asd_cProvCanc, Asd_cOperaTC, " & _
'                    "Asd_cTipoMoneda, Tra_cCodigo, Asd_cFormaPago) Values " & _
'                    "('" & Format(Val(Sht.Range("A" & VarContFila).Value), "0000000000") & "','" & Sht.Range("B" & VarContFila).Value & "','" & Format(Val(Sht.Range("C" & VarContFila).Value), "00") & _
'                    "','" & Format(Val(Sht.Range("D" & VarContFila).Value), "00") & "','" & Format(Val(Sht.Range("E" & VarContFila).Value), "0000000000") & "','" & Sht.Range("F" & VarContFila).Value & _
'                    "','" & Sht.Range("G" & VarContFila).Value & "','" & Sht.Range("H" & VarContFila).Value & "', " & NE(Sht.Range("I" & VarContFila).Value) & _
'                    ", " & NE(Sht.Range("J" & VarContFila).Value) & ", " & NE(Sht.Range("K" & VarContFila).Value) & ", " & NE(Sht.Range("L" & VarContFila).Value) & _
'                    ", " & NE(Sht.Range("M" & VarContFila).Value) & ", '" & Sht.Range("N" & VarContFila).Value & "', '" & Sht.Range("O" & VarContFila).Value & _
'                    "', '" & IIf(LTrim(RTrim(Sht.Range("P" & VarContFila).Value)) = "", "", Format(Val(Sht.Range("P" & VarContFila).Value), "00000")) & "', '" & IIf(LTrim(RTrim(Sht.Range("Q" & VarContFila).Value)) = "", "", Format(Val(Sht.Range("Q" & VarContFila).Value), "00")) & "', '" & IIf(LTrim(RTrim(Sht.Range("R" & VarContFila).Value)) = "", "", Format(Sht.Range("R" & VarContFila).Value, "dd/mm/yyyy")) & _
'                    "', '" & IIf(LTrim(RTrim(Sht.Range("S" & VarContFila).Value)) = "", "", Format(Val(Sht.Range("S" & VarContFila).Value), "000")) & "', '" & IIf(Sht.Range("T" & VarContFila).Value = "", "", Sht.Range("T" & VarContFila).Value) & "', '" & Sht.Range("U" & VarContFila).Value & _
'                    "', '" & Sht.Range("V" & VarContFila).Value & "', '" & Sht.Range("W" & VarContFila).Value & "', '" & Sht.Range("X" & VarContFila).Value & _
'                    "', '" & Sht.Range("Y" & VarContFila).Value & "', '" & Sht.Range("Z" & VarContFila).Value & "', '" & Sht.Range("AA" & VarContFila).Value & _
'                    "', '" & Sht.Range("AB" & VarContFila).Value & "', '" & Format(Val(Sht.Range("AC" & VarContFila).Value), "000") & "', '" & Sht.Range("AD" & VarContFila).Value & _
'                    "', '" & Sht.Range("AE" & VarContFila).Value & "')")
'                cn.CommitTrans
'            Next VarContFila
'        End If
'
'        If chkOpcion(7).Value = 1 Then
'            Call cn.Execute("if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_PLAN_CAB]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf & _
'            "drop table [dbo].[ZIMP_PLAN_CAB]")
'
'            VarSql = "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_PLAN_CAB] (" & vbCrLf & _
'                        "[Ase_cNummov] char (10) NULL, " & vbCrLf & _
'                        "[Pan_cAnio] char (4) NULL, " & vbCrLf & _
'                        "[Per_cPeriodo] char (2) NULL, " & vbCrLf & _
'                        "[Lib_cTipoLibro] char (2) NULL, " & vbCrLf & _
'                        "[Ase_nVoucher] char (10) NULL, " & vbCrLf & _
'                        "[Ase_dFecha] nvarchar (255) NULL, " & vbCrLf & _
'                        "[Ase_cGlosa] nvarchar (255) NULL, " & vbCrLf & _
'                        "[Ase_cTipoMoneda] nvarchar (255) NULL" & vbCrLf & _
'                        ")"
'            Call cn.Execute(VarSql)
'
'            VarContFila = 0
'            Set Sht = Wb.Worksheets("DIARIO_CAB")
'
'            For VarContFila = 2 To 70000
'                If Sht.Range("A" & VarContFila).Value = "" Then
'                    Exit For
'                End If
'            Next VarContFila
'
'            lblAvance.Caption = "IMPORTANDO A TABLA TEMPORAL - DIARIOCAB"
'            Me.pbAvance.Min = 0
'            Me.pbAvance.Value = 0
'            Me.pbAvance.Max = VarContFila - 1
'
'            For VarContFila = 2 To 70000
'                If Sht.Range("A" & VarContFila).Value = "" Then Exit For
'                pbAvance.Value = pbAvance.Value + 1
'                pbAvance.Refresh
'                DoEvents
'                cn.BeginTrans
'                    cn.Execute ("Insert Into ZIMP_PLAN_CAB(Ase_cNummov, Pan_cAnio, Per_cPeriodo, " & _
'                    "Lib_cTipoLibro, Ase_nVoucher, Ase_dFecha, Ase_cGlosa, Ase_cTipoMoneda) Values " & _
'                    "('" & Sht.Range("A" & VarContFila).Value & "','" & Sht.Range("B" & VarContFila).Value & "','" & Sht.Range("C" & VarContFila).Value & _
'                    "','" & Sht.Range("D" & VarContFila).Value & "','" & Sht.Range("E" & VarContFila).Value & "','" & Sht.Range("F" & VarContFila).Value & _
'                    "','" & Sht.Range("G" & VarContFila).Value & "','" & Sht.Range("H" & VarContFila).Value & "')")
'                cn.CommitTrans
'            Next VarContFila
'
'            Call cn.Execute("if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ZIMP_PLAN_DET]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf & _
'            "drop table [dbo].[ZIMP_PLAN_DET]")
'
'            VarSql = "CREATE TABLE [" & sBaseDatos & "].[dbo].[ZIMP_PLAN_DET] (" & vbCrLf & _
'                        "[Ase_cNummov] char (10) NULL, " & vbCrLf & "[Pan_cAnio] char (4) NULL, " & vbCrLf & _
'                        "[Per_cPeriodo] char (2) NULL, " & vbCrLf & "[Lib_cTipoLibro] char (2) NULL, " & vbCrLf & _
'                        "[Ase_nVoucher] char (10) NULL, " & vbCrLf & "[Pla_cCuentaContable] varchar (12) NULL, " & vbCrLf & _
'                        "[Asd_nItem] INT NULL, " & vbCrLf & "[Asd_cGlosa] nvarchar (255) NULL, " & vbCrLf & _
'                        "[Asd_nDebeSoles] NUMERIC (14,3) NULL , " & vbCrLf & "[Asd_nHaberSoles] NUMERIC (14,3) NULL , " & vbCrLf & _
'                        "[Asd_nTipoCambio] NUMERIC (14,3) NULL , " & vbCrLf & "[Asd_nDebeMonExt] NUMERIC (14,3) NULL , " & vbCrLf & _
'                        "[Asd_nHaberMonExt] NUMERIC (14,3) NULL , " & vbCrLf & "[Cos_cCodigo] varchar (12) NULL, " & vbCrLf & _
'                        "[Ten_cTipoEntidad] char (1) NULL, " & vbCrLf & "[Ent_cCodEntidad] char (5) NULL, " & vbCrLf & _
'                        "[Asd_cTipoDoc] char (2) NULL, " & vbCrLf & "[Asd_dFecDoc] datetime null , " & vbCrLf & _
'                        "[Asd_cSerieDoc] varchar (5) NULL, " & vbCrLf & "[Asd_cNumDoc] varchar (15) NULL, " & vbCrLf & _
'                        "[Asd_dFecVen] datetime NULL, " & vbCrLf & "[Asd_cProvCanc] char (1) NULL, " & vbCrLf & _
'                        "[Asd_cOperaTC] char (3) NULL, " & vbCrLf & "[Asd_cTipoMoneda] char (3) NULL" & vbCrLf & ")"
'            Call cn.Execute(VarSql)
'
'            VarContFila = 0
'            Set Sht = Wb.Worksheets("DIARIO_DET")
'
'            For VarContFila = 2 To 70000
'                If Sht.Range("A" & VarContFila).Value = "" Then
'                    Exit For
'                End If
'            Next VarContFila
'
'            lblAvance.Caption = "IMPORTANDO A TABLA TEMPORAL - DIARIODET"
'            Me.pbAvance.Min = 0
'            Me.pbAvance.Value = 0
'            Me.pbAvance.Max = VarContFila - 1
'
'            For VarContFila = 2 To 70000
'                If Sht.Range("A" & VarContFila).Value = "" Then Exit For
'                pbAvance.Value = pbAvance.Value + 1
'                pbAvance.Refresh
'                DoEvents
'                cn.BeginTrans
'
'
'                cn.Execute ("Insert Into ZIMP_PLAN_DET(Ase_cNummov, Pan_cAnio, Per_cPeriodo, " & _
'                "Lib_cTipoLibro, Ase_nVoucher, Pla_cCuentaContable, Asd_nItem, Asd_cGlosa, Asd_nDebeSoles, " & _
'                "Asd_nHaberSoles, Asd_nTipoCambio, " & _
'                "Asd_nDebeMonExt, Asd_nHaberMonExt, Cos_cCodigo, Ten_cTipoEntidad, Ent_cCodEntidad, " & _
'                "Asd_cTipoDoc, Asd_dFecDoc, Asd_cSerieDoc, " & _
'                "Asd_cNumDoc, Asd_dFecVen, Asd_cProvCanc, Asd_cOperaTC, " & _
'                "Asd_cTipoMoneda) Values " & _
'                "('" & Sht.Range("A" & VarContFila).Value & "','" & Sht.Range("B" & VarContFila).Value & "','" & Sht.Range("C" & VarContFila).Value & _
'                "','" & Sht.Range("D" & VarContFila).Value & "','" & Sht.Range("E" & VarContFila).Value & "','" & Sht.Range("F" & VarContFila).Value & _
'                "','" & Sht.Range("G" & VarContFila).Value & "','" & Sht.Range("H" & VarContFila).Value & "', " & NE(Sht.Range("I" & VarContFila).Value) & _
'                ", " & NE(Sht.Range("J" & VarContFila).Value) & ", " & NE(Sht.Range("K" & VarContFila).Value) & ", " & NE(Sht.Range("L" & VarContFila).Value) & _
'                ", " & NE(Sht.Range("M" & VarContFila).Value) & ", '" & Sht.Range("N" & VarContFila).Value & "', '" & Sht.Range("O" & VarContFila).Value & _
'                "', '" & Sht.Range("P" & VarContFila).Value & "', '" & Sht.Range("Q" & VarContFila).Value & "', '" & Sht.Range("R" & VarContFila).Value & _
'                "', '" & Sht.Range("S" & VarContFila).Value & "', '" & Sht.Range("T" & VarContFila).Value & "', '" & Sht.Range("U" & VarContFila).Value & _
'                "', '" & Sht.Range("V" & VarContFila).Value & "', '" & Sht.Range("W" & VarContFila).Value & "', '" & Sht.Range("X" & VarContFila).Value & "')")
'                cn.CommitTrans
'            Next VarContFila
'        End If
'
'        Ex.Save
'        Ex.Quit
'        cn.Close
'        Set cn = Nothing
'        Set Sht = Nothing
'        Set Wb = Nothing
'        Set Ex = Nothing
'
'        If ProcesaTablas Then
'            DoEvents
'            If Me.chkOpcion(2).Value = 1 Or chkOpcion(3).Value = 1 Or chkOpcion(4).Value = 1 Or _
'              chkOpcion(5).Value = 1 Or chkOpcion(6).Value = 1 Or chkOpcion(7).Value = 1 Then
'                If Mensajes("Desea actualizar los saldos ahora", vbQuestion + vbYesNo) = vbYes Then
'                    Call ActualizaSaldos
'                End If
'            End If
'            DoEvents
'        End If
'    'End If
'    lblAvance.Caption = ""
'    pbAvance.Value = 0
''    pbAvance.Max = 0
'    pbAvance.Refresh
'    DoEvents
'
'    cmdImportarDatos.Enabled = True
'    cmdSeleccionar.Enabled = True
'    cmd_salir.Enabled = True
'    cmdImprimir.Enabled = True
'    Me.MousePointer = vbNormal
'Exit Sub
'Control:
'
''MsgBox "Error en la importación", vbCritical + vbSystemModal, "Mensaje: Sistema"
'MsgBox Err.Description & vbCrLf & sSql & vbcrfl & "Fila: " & VarContFila, vbCritical + vbSystemModal, "Mensaje: Sistema"
''Resume
'        Ex.Quit
''        cn.Close
'        Set cn = Nothing
'        Set Sht = Nothing
'        Set Wb = Nothing
'        Set Ex = Nothing
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
    With Me.dlgAbrirArchivo
        .DialogTitle = "Archivo de Datos de Asientos"
        .InitDir = "C:"
        .Filter = "Archivos de Datos(*.xls)| *.xls"
        .CancelError = True
        .ShowOpen
        If .filename = "" Then
            Mensajes "Selecciones un archivo", vbInformation
        Else
            tdbtArchivo = .filename
        End If
    End With
  
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


Private Function ProcesaTablas() As Boolean
'    Dim lArrMnt(11) As Variant
'    Dim i As Integer
'    Dim Sql As String
'    Dim scadena As String
'    Dim entro As Boolean
'    Dim clsMante2 As clsMantoTablas
'
'    Call EscribirLog("Iniciando importacion de plantilla XLS de la empresa " & gsEmpresa & " " & gsEmpresaNom, Me.Name)
'
'    On Local Error GoTo ErrorEjecucion
'    Set clsMante2 = New clsMantoTablas
'
'    ProcesaTablas = True
'    entro = False
'
'
'    Screen.MousePointer = vbHourglass
'    '-------------------------------------------'
'    lArrMnt(0) = "ELIMINACION"
'    lArrMnt(1) = gsEmpresa
'    lArrMnt(2) = gsAnio
'    lArrMnt(11) = gsUsuario
'
'    clsMante2.InicializaClase
'    clsMante2.BeginTrans
'
'    If clsMante2.MantenimientoDeTablas(gsCadenaConexion, "spCn_ReprocesoImportacionXLSv2", lArrMnt(), False) = False Then
'        gsImportacion = False
'        ProcesaTablas = False
'    End If
'
'    clsMante2.CommitTrans
'    clsMante2.FinalizaClase
'    '-------------------------------------------'
'    lArrMnt(0) = "IMPORTACION"
'    lArrMnt(1) = gsEmpresa
'    lArrMnt(2) = gsAnio
'    lArrMnt(11) = gsUsuario
'
'    For i = 0 To 7
'        If chkOpcion(i).Value = vbChecked Then
'            entro = True
'
'            lArrMnt(3) = IIf(i = 0, "1", "0") 'ENTIDAD
'            lArrMnt(4) = IIf(i = 1, "1", "0") 'TIPO DE CAMBIO
'            lArrMnt(5) = IIf(i = 2, "1", "0") 'APERTURA
'            lArrMnt(6) = IIf(i = 3, "1", "0") 'COMPRAS
'            lArrMnt(7) = IIf(i = 4, "1", "0") 'VENTAS
'            lArrMnt(8) = IIf(i = 5, "1", "0") 'CAJA INGRESO
'            lArrMnt(9) = IIf(i = 6, "1", "0") 'CAJA EGRESO
'            lArrMnt(10) = IIf(i = 7, "1", "0") 'PLANILLA
'
'            gsImportacion = True
'
'            DoEvents
'
'            Select Case i
'                Case 0: scadena = "ENTIDAD"
'                Case 1: scadena = "TIPO DE CAMBIO"
'                Case 2: scadena = "APERTURA"
'                Case 3: scadena = "COMPRAS"
'                Case 4: scadena = "VENTAS"
'                Case 5: scadena = "CAJA INGRESO"
'                Case 6: scadena = "CAJA EGRESO"
'                Case 7: scadena = "PLANILLA"
'            End Select
'
'            Me.lblAvance.Caption = "PROCESANDO -> " & scadena
'
'            lblAvance.Caption = "MIGRANDO A LAS TABLAS DE LA BD DEL PROCESO -> " & scadena
'            DoEvents
'            Me.pbAvance.Min = 0
'            Me.pbAvance.Value = 0
'            Me.pbAvance.Max = 30
'
'            If pbAvance.Value + 1 > pbAvance.Max Then pbAvance.Max = pbAvance.Max + 1
'
'            pbAvance.Value = pbAvance.Value + 1
'            pbAvance.Refresh
'            lblAvance.Caption = "Importando -> " & archivo
'            lblAvance.Refresh
'
'            clsMante2.InicializaClase
'            clsMante2.BeginTrans
'
'            If clsMante2.MantenimientoDeTablas(gsCadenaConexion, "spCn_ReprocesoImportacionXLSv2", lArrMnt(), False) = False Then
'                gsImportacion = False
'                ProcesaTablas = False
'            End If
'
'            pbAvance.Value = pbAvance.Max
'
'            lblAvance.Caption = "Proceso terminado ... "
'            lblAvance.Caption = ""
'            lblAvance.Refresh
'
'            clsMante2.CommitTrans
'            clsMante2.FinalizaClase
'
'            DoEvents
'        End If
'    Next i
'
'    gsImportacion = False
'    Set clsMante2 = Nothing
'
'
'    '-----------------------------------
'    Sql = "select top 1 emp_ccodigo from CNT_REPORTE_IMPORTACION where emp_ccodigo='" & gsEmpresa & "' "
'    scadena = CE(fRetornaValor(Sql))
'
'    If scadena <> "" Then
'        Call Mensajes("Error en los datos de importación, revise el reporte de errores")
'        ProcesaTablas = False
'    Else
'        If entro = False Then
'            Call Mensajes("Seleccione una opcion")
'        Else
'            Call Mensajes("Proceso terminado")
'            Call EscribirLog("Finalizo la importacion de plantilla XLS de la empresa " & gsEmpresa & " " & gsEmpresaNom, Me.Name)
'        End If
'
'    End If
'
'    If entro = False Then ProcesaTablas = False
'    '-----------------------------------
'    Screen.MousePointer = vbNormal
'    Exit Function
'ErrorEjecucion:
'    Call EscribirLog("Error de importacion de plantilla XLS [" & Err.Description & "] de la empresa " & gsEmpresa & " " & gsEmpresaNom, Me.Name)
'    Mensajes Str(Err.Number) & " " & Err.Description, vbInformation
'    Resume
'    Set clsMante2 = Nothing
'    gsImportacion = False
'    Screen.MousePointer = vbNormal
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
      '  Call CentrarTitulo(lblTitulo, fraTodo, Me)
    End If
Exit Sub
errHand:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))

End Sub

