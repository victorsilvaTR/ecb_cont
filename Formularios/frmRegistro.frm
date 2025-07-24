VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{879115B9-8D7C-43CA-ADFE-8B489017BF42}#1.0#0"; "activelock1884.ocx"
Begin VB.Form frmRegistro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ECB-Registro"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   6135
   Icon            =   "frmRegistro.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   6135
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      ForeColor       =   &H80000008&
      Height          =   6585
      Left            =   0
      Picture         =   "frmRegistro.frx":0ECA
      ScaleHeight     =   6555
      ScaleWidth      =   10155
      TabIndex        =   0
      Top             =   -360
      Width           =   10185
      Begin VB.TextBox txtSerie 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   2520
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   2205
         Width           =   2940
      End
      Begin VB.TextBox txtCodigo 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1665
         Width           =   2940
      End
      Begin activelock1884.ActiveLock ActiveLock1 
         Left            =   150
         Top             =   2655
         _ExtentX        =   847
         _ExtentY        =   820
         SoftwareName    =   ""
         SoftwarePassword=   ""
         LiberationKeyLength=   16
         SoftwareCodeLength=   16
         LockToHardDrive =   0   'False
         LockToWindowsSerial=   0   'False
         LockToRandomNumber=   0   'False
         LockToComputerName=   0   'False
         LockToMACAddress=   0   'False
         UseDataLock     =   0   'False
         RegistryPath    =   ""
         RegistryKey     =   ""
         RegistryName    =   ""
         RegistryHive    =   "HKLM"
         LockToCustomString=   ""
         HashAlgorithm   =   1
         RegCounterKey   =   ""
         RegLiberationKey=   ""
         RegLastRunDateKey=   ""
         RegInitialRunDateKey=   ""
         RegRandomKey    =   ""
         EncKey          =   "Default"
         RegEncKey       =   -1  'True
      End
      Begin MSForms.CommandButton cmdRegistrar 
         Height          =   390
         Left            =   2520
         TabIndex        =   7
         Top             =   2760
         Width           =   1530
         Caption         =   "  Registrar"
         PicturePosition =   327683
         Size            =   "2699;688"
         Picture         =   "frmRegistro.frx":4DBD4
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código de Licencia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   6
         Left            =   270
         TabIndex        =   6
         Top             =   2250
         Width           =   2010
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0C0C0&
         Height          =   1410
         Left            =   225
         Top             =   4995
         Width           =   9555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código Generado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   5
         Left            =   270
         TabIndex        =   4
         Top             =   1710
         Width           =   1890
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "¡¡¡ SISTEMA NO REGISTRADO !!!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   4
         Left            =   90
         TabIndex        =   2
         Top             =   450
         Width           =   5865
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Para poder ejecutar el programa comuniquese con su proveedor del sistema para adquirir un codigo de licencia."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   915
         Index           =   1
         Left            =   0
         TabIndex        =   1
         Top             =   945
         Width           =   6045
      End
   End
End
Attribute VB_Name = "frmRegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nContador As Integer


Private Sub cmdRegistrar_Click()
    ActiveLock1.Register (txtSerie.Text)
End Sub

Private Sub Form_Activate()

    If ActiveLock1.RegisteredUser = True Then
        Unload Me
        Call LlamaFormulario
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo serror
    
    ActiveLock1.SoftwareName = oRegistroLock.SoftwareName
    ActiveLock1.SoftwarePassword = oRegistroLock.SoftwarePassword
    ActiveLock1.RegistryHive = oRegistroLock.RegistryHive
    ActiveLock1.RegistryKey = oRegistroLock.RegistryKey
    ActiveLock1.RegistryName = oRegistroLock.RegistryName
    ActiveLock1.RegistryPath = oRegistroLock.RegistryPath
    ActiveLock1.LiberationKeyLength = oRegistroLock.LiberationKeyLength
    ActiveLock1.SoftwareCodeLength = oRegistroLock.SoftwareCodeLength
    ActiveLock1.LockToHardDrive = oRegistroLock.LockToHardDrive
    ActiveLock1.LockToWindowsSerial = oRegistroLock.LockToWindowsSerial
    ActiveLock1.LockToComputerName = oRegistroLock.LockToComputerName
    ActiveLock1.LockToRandomNumber = oRegistroLock.LockToRandomNumber
    ActiveLock1.LockToMACAddress = oRegistroLock.LockToMACAddress
    ActiveLock1.LockToCustomString = oRegistroLock.LockToCustomString
    ActiveLock1.HashAlgorithm = oRegistroLock.HashAlgorithm
    ActiveLock1.UseDataLock = oRegistroLock.UseDataLock
    ActiveLock1.RegCounterKey = oRegistroLock.RegCounterKey
    ActiveLock1.RegLastRunDateKey = oRegistroLock.RegLastRunDateKey
    ActiveLock1.RegRandomKey = oRegistroLock.RegRandomKey
    ActiveLock1.RegLiberationKey = oRegistroLock.RegLiberationKey
    ActiveLock1.RegInitialRunDateKey = oRegistroLock.RegInitialRunDateKey
    
    nContador = ActiveLock1.Counter
    
    'GetMACAddress("-") '
    txtCodigo.Text = ActiveLock1.SoftwareCode
     
   ' txtCodPC.Text = GetMACAddress("-") 'ActiveLock1.LockToCustomString

serror:
    Me.Height = 3300
End Sub


Private Sub ActiveLock1_Registration(WasSuccessful As Boolean)
    If WasSuccessful Then
        Mensajes "Gracias por Registrar su sistema!", vbInformation
        Unload Me
        Call LlamaFormulario
    Else
        Mensajes "Serie no valida, ingresela nuevamente.", vbExclamation

        If ActiveLock1.Counter > nContador + 3 Then
            Mensajes "El registro fue bloqueado, reinicie el sistema para continuar.", vbExclamation
        End If
    End If
End Sub

Private Sub LlamaFormulario()
    frmPrcIngresoSistema.Show
End Sub

