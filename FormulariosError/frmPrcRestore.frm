VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmPrcRestore 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Restaurar Base de Datos"
   ClientHeight    =   4545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6285
   Icon            =   "frmPrcRestore.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4545
   ScaleWidth      =   6285
   Begin VB.Timer Timer1 
      Left            =   5715
      Top             =   180
   End
   Begin MSComDlg.CommonDialog dlgAbrirArchivo 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraTodo 
      Height          =   4455
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   6180
      Begin VB.TextBox txtRutaBackup 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "\\SERVIDOR"
         Top             =   1305
         Width           =   4245
      End
      Begin VB.FileListBox File1 
         Height          =   1260
         Left            =   240
         Pattern         =   "*.bak"
         TabIndex        =   0
         Top             =   2490
         Width           =   5655
      End
      Begin MSForms.CommandButton cmdSalir 
         Height          =   435
         Left            =   3195
         TabIndex        =   9
         Top             =   3870
         Width           =   1665
         Caption         =   " Salir"
         PicturePosition =   327683
         Size            =   "2937;767"
         Picture         =   "frmPrcRestore.frx":0ECA
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdBackup 
         Height          =   435
         Left            =   1305
         TabIndex        =   8
         Top             =   3870
         Width           =   1665
         Caption         =   " Restaurar"
         PicturePosition =   327683
         Size            =   "2937;767"
         Picture         =   "frmPrcRestore.frx":1464
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label Label4 
         Caption         =   "NOTA: ESTA OPERACION REEMPLAZARA TODOS LOS DATOS DE LA BASE DE DATOS ACTUAL CON LOS DATOS DE BACKUP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   420
         Left            =   270
         TabIndex        =   7
         Top             =   2025
         Width           =   5640
      End
      Begin VB.Label lblRutaBackup 
         Caption         =   "Ruta Backup:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   270
         TabIndex        =   6
         Top             =   1350
         Width           =   1230
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   $"frmPrcRestore.frx":19FE
         Height          =   450
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   5655
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "RESTAURAR BASE DE DATOS"
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
         Left            =   1845
         TabIndex        =   3
         Top             =   330
         Width           =   2820
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Seleccione el archivo del backup a restaurar"
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
         Left            =   270
         TabIndex        =   2
         Top             =   1680
         Width           =   5670
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
Attribute VB_Name = "frmPrcRestore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cSql As String
Dim cBat As String
Dim cSW As String
Dim cSql_ant As String
Dim cRuta As String
Dim gsGrupo As String

Private Sub Elimina()
    If ExisteArchivo(cSql_ant) Then Call EliminaArchivo(cSql_ant, False)
    If ExisteArchivo(cSql) Then Call EliminaArchivo(cSql, False)
    If ExisteArchivo(cBat) Then Call EliminaArchivo(cBat, False)
    If ExisteArchivo(cSW) Then Call EliminaArchivo(cSW, False)
End Sub

Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

Private Sub cmdBackup_Click()

    Call Elimina
    
    Dim fso As New Scripting.filesystemobject
    Dim archivo As String
    
    If CE(gsRutaBackup) = "" Then
        Mensajes "Configure la ruta del backup en el archivo de configuraciones del sistema", vbOKOnly + vbInformation
        Exit Sub
    End If
    
    If ExisteArchivo(gsRutaBackup & "\OSQL.exe") = False Then
        Mensajes "Copie el archivo OSQL.EXE a este directorio/carpeta"
        Exit Sub
    End If
    
    On Error GoTo NoSePuedeCopiar
    
    If fso.FolderExists(gsRutaBackup) = False Then
        Mensajes "La ruta asignada para el backup no existe" & Salto(1) & "Cambie la ruta en el archivo de configuraciones del sistema seccion" & Salto(1) & " BACKUP=RUTA", vbOKOnly + vbInformation
        Exit Sub
    End If

    Set fso = Nothing
    
    Dim respuesta As String
    Dim RetVal

    
    ' *** Verificar si se ha seleccionado la base de datos
    If File1.Selected(File1.ListIndex) = False Then
        Mensajes "Seleccione el backup a restaurar", vbInformation
        Exit Sub
    End If
    ' *** Preguntar si se quiere restaurar la base de datos
    respuesta = MsgBox("Desea restaurar la base de Datos con el Backup Seleccionado", vbYesNo + vbQuestion + vbDefaultButton2, "Confirmar Restaurar Base de Datos")
    If respuesta = vbNo Then
        Exit Sub
    End If
    ' *** Generar el archivo sql en la carpeta c
    Screen.MousePointer = vbHourglass
    CrearSqlInicial
    Call CrearSql
    
    ' *** Crear y Ejecutar el archivo bat
    Call CrearBat
    
    Call EscribirLog("Iniciando restauracion de backup [" & gsRutaBackup & "\" & File1.List(File1.ListIndex) & "]", Me.Name)
    RetVal = Shell(cBat, vbNormalFocus)
    Timer1.Interval = 10
    cmdBackup.Enabled = False
    Exit Sub
NoSePuedeCopiar:
    Call EscribirLog("Error a restaurar backup ", Me.Name)
End Sub

Private Sub cmdsalir_Click()
    Screen.MousePointer = 0
    Unload Me
End Sub

Private Sub CrearSqlInicial()
    ' *** Verificando si existe el archivo sql
    If ExisteArchivo(cSql_ant) Then EliminaArchivo (cSql_ant)
    Open cSql_ant For Output As #1 Len = 220
    Print #1, "USE master"
    Print #1, "BACKUP DATABASE " & gsBD & " TO DISK = '" & gsRutaBackup & "\" & gsBD & "_" & Format(Date, "yyyyMMdd") & "_" & Format(Time, "HHMMSS") & "_" & gsUsuario & "_AUTOMATICO.BAK'"
    Print #1, "with init"
    Close #1
    ' ***
End Sub

Private Sub CrearSql()
    ' *** Verificando si existe el archivo sql
    If ExisteArchivo(cSql) Then EliminaArchivo (cSql)
    Open cSql For Output As #1 Len = 220
    Print #1, "USE master"
    Print #1, "ALTER DATABASE " & gsBD
    Print #1, "SET SINGLE_USER"
    Print #1, "RESTORE DATABASE " & gsBD & " FROM DISK = '" & gsRutaBackup & "\" & File1.List(File1.ListIndex) & "'"
    Print #1, "With RECOVERY"
    Print #1, "ALTER DATABASE " & gsBD
    Print #1, "SET MULTI_USER"
    Close #1
    ' ***
End Sub

Private Sub CrearBat()
    
    If ExisteArchivo(cBat) Then EliminaArchivo (cBat)
    If ExisteArchivo(cSW) Then EliminaArchivo (cSW)
    
    Open cBat For Output As #1 Len = 220
    
    Print #1, gsRutaBackup; "\osql -U " & gsBDUS & " -P " & gsBDPW & " -S" & gsServidor & " -i " & cSql_ant
    
    If gsInstancia <> "" Then
        Print #1, "net stop MSSQL$" & gsInstancia
        Print #1, "net start MSSQL$" & gsInstancia
    Else
        Print #1, "net stop MSSQLSERVER "
        Print #1, "net start MSSQLSERVER "
    End If
    
    Print #1, gsRutaBackup; "\osql -U " & gsBDUS & " -P " & gsBDPW & " -S" & gsServidor & " -i " & cSql
    
    If gsInstancia <> "" Then
        Print #1, "net stop MSSQL$" & gsInstancia
        Print #1, "net start MSSQL$" & gsInstancia
    Else
        Print #1, "net stop MSSQLSERVER "
        Print #1, "net start MSSQLSERVER "
    End If
    
    Print #1, "copy " & cSql & " " & cSW
    Close #1
    
End Sub

Private Sub Form_Load()
    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))
    
    
    Call Centrar_form(Me)
    
    If InStr(1, gsServidor, "\") > 0 Then
        gsInstancia = Mid(gsServidor, InStrRev(gsServidor, "\") + 1)
    End If
    
    
    cRuta = App.Path
    If Right(cRuta, 1) = "\" Then cRuta = Mid(cRuta, 1, Len(cRuta) - 1)
    
    cSql_ant = cRuta & "\BackupIni.sql"
    cSql = cRuta & "\Backup.sql"
    cBat = cRuta & "\Backup.bat"
    cSW = cRuta & "\sw.sw"
    
    
    If Mid(gsGrupo, 3, 1) <> "1" And gsGrupo <> gsPrivilegioAdmin Then
        Me.cmdBackup.Enabled = False
        'Me.cmdSeleccionar.Enabled = False
    Else
        Me.cmdBackup.Enabled = True
        'Me.cmdSeleccionar.Enabled = True
    End If
    
    If CE(gsRutaBackup) = "" Then
        Mensajes "Configure la ruta del backup en el archivo de configuraciones del sistema", vbOKOnly + vbInformation
        Exit Sub
    End If
    
    On Error GoTo NoSePuedeCopiar
    Dim fso As New Scripting.filesystemobject
    If fso.FolderExists(gsRutaBackup) = False Then
        Mensajes "La ruta asignada para el backup no existe" & Salto(1) & "Cambie la ruta en el archivo de configuraciones del sistema seccion" & Salto(1) & " BACKUP=RUTA", vbOKOnly + vbInformation
        Exit Sub
    End If
    Set fso = Nothing
    Me.File1.Path = gsRutaBackup
    txtRutaBackup.Text = gsRutaBackup
    
    Exit Sub

NoSePuedeCopiar:
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
    On Error GoTo serror
    ' *** Eliminar el archivo sql y bat de la carpeta c
    Call Elimina
    Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))
    
    Exit Sub
serror:
End Sub

Private Sub Timer1_Timer()
On Local Error GoTo ErrorEjecucion

If ExisteArchivo(cSW) Then
            DoEvents
            Call Elimina
            Mensajes "Se termino de realizar la restauracion de la Base de Datos", vbInformation
            Call Elimina
            Screen.MousePointer = vbNormal
            Timer1.Enabled = False
            cmdBackup.Enabled = True
            Call EscribirLog("Finalizo la restauracion de backup [" & gsRutaBackup & "\" & File1.List(File1.ListIndex) & "]", Me.Name)
End If

ErrorEjecucion:
End Sub

