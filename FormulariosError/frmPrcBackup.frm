VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmPrcBackup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backup de la Base de Datos"
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7140
   Icon            =   "frmPrcBackup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   7140
   Begin VB.Frame fraTodo 
      Height          =   3300
      Left            =   45
      TabIndex        =   3
      Top             =   0
      Width           =   7035
      Begin VB.Frame Frame2 
         Height          =   1935
         Left            =   405
         TabIndex        =   4
         Top             =   705
         Width           =   6180
         Begin VB.TextBox txtRutaBackup 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1620
            Locked          =   -1  'True
            TabIndex        =   0
            Text            =   "\\SERVIDOR"
            Top             =   900
            Width           =   4245
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   $"frmPrcBackup.frx":0ECA
            Height          =   510
            Left            =   360
            TabIndex        =   7
            Top             =   360
            Width           =   5655
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
            Left            =   360
            TabIndex        =   6
            Top             =   945
            Width           =   1230
         End
         Begin VB.Label lblNOTAESTA 
            Alignment       =   2  'Center
            Caption         =   "NOTA: ES RECOMENDABLE SACAR MINIMO UN BACKUP A LA SEMANA O DIARIO SEGUN SUS TRANSACCIONES"
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
            Left            =   360
            TabIndex        =   5
            Top             =   1350
            Width           =   5640
         End
      End
      Begin VB.Timer Timer1 
         Left            =   6405
         Top             =   225
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "GENERAR BACKUP (COPIA DE SEGURIDAD)"
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
         Index           =   0
         Left            =   1845
         TabIndex        =   8
         Top             =   345
         Width           =   3360
      End
      Begin MSForms.CommandButton cmdBackup 
         Height          =   435
         Left            =   1725
         TabIndex        =   1
         Top             =   2745
         Width           =   1665
         Caption         =   " Backup"
         PicturePosition =   327683
         Size            =   "2937;767"
         Picture         =   "frmPrcBackup.frx":0F60
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdSalir 
         Height          =   435
         Left            =   3660
         TabIndex        =   2
         Top             =   2745
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
      TabIndex        =   9
      Top             =   0
      Width           =   4365
   End
End
Attribute VB_Name = "frmPrcBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gsGrupo As String
Dim RetVal
Dim cSql As String
Dim cBat As String
Dim cSW As String
Dim cRuta As String

Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

Private Sub Form_Load()
    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))
    
    
    Call Centrar_form(Me)
    
    cRuta = App.Path
    If Right(cRuta, 1) = "\" Then cRuta = Mid(cRuta, 1, Len(cRuta) - 1)
    
    cSql = cRuta & "\Backup.sql"
    cBat = cRuta & "\Backup.bat"
    cSW = cRuta & "\sw.sw"
        
    
    If Mid(gsGrupo, 3, 1) <> "1" And gsGrupo <> gsPrivilegioAdmin Then
        Me.cmdBackup.Enabled = False

    Else
        Me.cmdBackup.Enabled = True

    End If
    
    txtRutaBackup.Text = gsRutaBackup
End Sub

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
    
    ' *** Generar el archivo sql en la carpeta c
    Screen.MousePointer = vbHourglass
    Call CrearSql
    ' *** Crear y Ejecutar el archivo bat
    Call CrearBat
    
    Call EscribirLog("Iniciando copia de seguridad", Me.Name)
    
    RetVal = Shell(cBat, vbNormalFocus)
    cmdBackup.Enabled = False
    Timer1.Interval = 10
    Screen.MousePointer = vbNormal
    Timer1.Enabled = True
    pSetFocus txtRutaBackup
    Exit Sub
    
NoSePuedeCopiar:
    Call EscribirLog("Error a sacar copia de seguridad, [" & Err.Description & "]", Me.Name)
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub CrearSql()
    ' *** Verificando si existe el archivo sql
    If ExisteArchivo(cSql) Then EliminaArchivo (cSql)
    
    Open cSql For Output As #1 Len = 220
    Print #1, "USE master"
    Print #1, "BACKUP DATABASE " & gsBD & " TO DISK = '" & gsRutaBackup & "\" & gsBD & "_" & Format(Date, "yyyyMMdd") & "_" & Format(Time, "HHMMSS") & "_" & gsUsuario & ".BAK'"
    Print #1, "with init"
    Close #1
    ' ***
End Sub

Private Sub CrearBat()
    ' *** Verificando si existe el archivo sql
    If ExisteArchivo(cBat) Then EliminaArchivo (cBat)
    If ExisteArchivo(cSW) Then EliminaArchivo (cSW)
    '-----------------
    Open cBat For Output As #1 Len = 220
    Print #1, gsRutaBackup; "\osql -U " & gsBDUS & " -P " & gsBDPW & " -S" & gsServidor & " -i " & cSql
    Print #1, "copy " & cSql & " " & cSW
    Close #1
    ' ***
End Sub

Private Function ExtraeCarpeta(cadena As String)
    Dim i As Integer
    ExtraeCarpeta = cadena
    If Right(cadena, 1) = "\" Then ExtraeCarpeta = Mid$(cadena, 1, Len(cadena) - 1)
End Function

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
    Call Elimina
    Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))

End Sub

Private Sub Elimina()
    If ExisteArchivo(cSql) Then Call EliminaArchivo(cSql, False)
    If ExisteArchivo(cBat) Then Call EliminaArchivo(cBat, False)
    If ExisteArchivo(cSW) Then Call EliminaArchivo(cSW, False)
End Sub

Private Sub Timer1_Timer()
On Local Error GoTo ErrorEjecucion
Dim cadena As String
If IsNull(RetVal) = False Then
    If ExisteArchivo(cSW) Then
        DoEvents
        Call Elimina
        Mensajes "Se termino de realizar el backup a la Base de Datos", vbInformation
        Call Elimina
        Screen.MousePointer = vbNormal
        Timer1.Enabled = False
        cmdBackup.Enabled = True
        Call EscribirLog("Finalizando copia de seguridad", Me.Name)
    End If
End If
ErrorEjecucion:
End Sub

