VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmAcercaDe 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acerca de ..."
   ClientHeight    =   4275
   ClientLeft      =   2625
   ClientTop       =   2820
   ClientWidth     =   7245
   Icon            =   "frmAbout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4275
   ScaleWidth      =   7245
   Begin VB.Frame fraTodo 
      BackColor       =   &H00FFFFFF&
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   -45
      Width           =   7260
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1200
         Left            =   75
         Picture         =   "frmAbout.frx":0ECA
         ScaleHeight     =   1200
         ScaleWidth      =   7095
         TabIndex        =   3
         Top             =   180
         Width           =   7095
         Begin VB.Label lblTitulo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "EcbCont v1.0"
            BeginProperty Font 
               Name            =   "Lucida Sans Unicode"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   330
            Left            =   1395
            TabIndex        =   5
            Top             =   180
            Width           =   5610
         End
         Begin VB.Label lblVersion 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Sistema de Contabilidad General"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   285
            Left            =   1530
            TabIndex        =   4
            Top             =   540
            Width           =   5415
         End
      End
      Begin VB.Timer Timer1 
         Interval        =   50
         Left            =   6705
         Top             =   2445
      End
      Begin VB.PictureBox picAbout 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1665
         Left            =   75
         ScaleHeight     =   1665
         ScaleWidth      =   7095
         TabIndex        =   1
         Top             =   1380
         Width           =   7095
         Begin VB.Label lblNormal 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "SISTEMA DE REPORTES Y CONSULTAS"
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
            Height          =   195
            Left            =   -60
            TabIndex        =   2
            Top             =   45
            Width           =   7065
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   $"frmAbout.frx":1757
         Height          =   1065
         Left            =   45
         TabIndex        =   8
         Top             =   3165
         Width           =   4950
      End
      Begin MSForms.CommandButton Command1 
         Height          =   420
         Left            =   5175
         TabIndex        =   7
         Top             =   3210
         Width           =   1905
         VariousPropertyBits=   19
         Caption         =   "  Aceptar"
         PicturePosition =   327683
         Size            =   "3360;741"
         Picture         =   "frmAbout.frx":186B
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton Command2 
         Height          =   420
         Left            =   5175
         TabIndex        =   6
         Top             =   3705
         Width           =   1905
         VariousPropertyBits=   19
         Caption         =   "  Info. del Sistema"
         PicturePosition =   327683
         Size            =   "3360;741"
         Picture         =   "frmAbout.frx":1E05
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
End
Attribute VB_Name = "frmAcercaDe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : frmAcercaDe
'    Project    : Contabilidad
'
'    Description: Formulario de acerda del sistema ECBCont
'--------------------------------------------------------------------------------


Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1
Const REG_DWORD = 4

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
 
'Propiedades
Private propPrograma As String

'Fin de Propiedades
 
Dim nItemTop As Integer
Dim nPos As Long
Dim nLineas As Integer
Dim aMensaje(13) As String

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Form_KeyDown
' Description:       Evento generado al presionar una tecla
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 And Shift = 1 Then
        On Error Resume Next
        Mensajes gsVersion & " " & Left(FileDateTime(App.Path & "\Ecb-Cont.exe"), 19)
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Form_KeyPress
' Description:       Evento generado al presionar una tecla
'
' Parameters :       KeyAscii (Integer)
'--------------------------------------------------------------------------------
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Form_Load
' Description:       Evento generado al iniciar el formulario
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub Form_Load()
    
    Call Centrar_form(Me)
    
    propPrograma = gsNombreModulo & " " & gsVersion
    lblTitulo.FontSize = 12
    lblVersion.FontSize = 12
    lblTitulo.Caption = "ECB-Cont " & gsVersion
    nPos = lblNormal.Top
    aMensaje(0) = "" & vbNewLine
    aMensaje(1) = "ESTUDIO CABALLERO BUSTAMANTE" & vbNewLine
    aMensaje(4) = "" & vbNewLine
    aMensaje(5) = propPrograma & vbNewLine
    aMensaje(6) = ""
    aMensaje(7) = "" & vbNewLine
    aMensaje(8) = "Av. San Borja Sur 1170" & vbNewLine
    aMensaje(9) = "Telf. 710-7100"
    aMensaje(10) = "" & vbNewLine
    aMensaje(12) = "LIMA - PERU" & vbNewLine
    aMensaje(13) = "2008"
    nItemTop = 0
    lblNormal = ""
    lblNormal.Height = 5000
    
    For nLineas = 0 To UBound(aMensaje)
        lblNormal = lblNormal & aMensaje(nLineas)
    Next
    
    
    
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Command1_Click
' Description:       Evento generado al hacer clic el el boton salir
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub Command1_Click()
    Timer1.Enabled = False
    Unload Me
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Command2_Click
' Description:       Evento generado al hacer clic en el boton de informacion del sistema
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub Command2_Click()
    Call StartSysInfo
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       StartSysInfo
' Description:       Procedimiento que ejecuta el EXE de informacion del sistema del sistema operativo
'
' Parameters :
'--------------------------------------------------------------------------------
Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
    Dim rc As Long
    Dim SysInfoPath As String
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
        Else
            GoTo SysInfoErr
        End If
    Else
        GoTo SysInfoErr
    End If
    Call Shell(SysInfoPath, vbNormalFocus)
    Exit Sub
SysInfoErr:
    Mensajes "La Información del Sistema no está disponible en este momento", vbOKOnly + vbInformation
End Sub


'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       GetKeyValue
' Description:       Procedimiento que obtiene el valor de la llave del registro del windows
'
' Parameters :       KeyRoot (Long)
'                    KeyName (String)
'                    SubKeyRef (String)
'                    KeyVal (String)
'--------------------------------------------------------------------------------
Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long
    Dim rc As Long
    Dim hKey As Long
    Dim hDepth As Long                                     '
    Dim KeyValType As Long
    Dim tmpVal As String
    Dim KeyValSize As Long
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey)
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Error
    tmpVal = String$(1024, 0)
    KeyValSize = 1024
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    tmpVal = VBA.Left(tmpVal, InStr(tmpVal, VBA.Chr(0)) - 1)
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
            KeyVal = Format$("&h" + KeyVal)                 ' Convert Double Word To String
    End Select
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
GetKeyError:                                                ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Image1_DblClick
' Description:       Evento que activa el timer del formulario
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub Image1_DblClick()
    Timer1.Enabled = Not Timer1.Enabled
End Sub

Private Sub Form_Resize()
On Error GoTo errHand
    If Me.WindowState <> vbMinimized Then
        Call SeteaFondoForm(Me)
        Call Centrar_Objeto(fratodo, Me)

    End If
Exit Sub
errHand:

End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Timer1_Timer
' Description:       Evento del control del Timer
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub Timer1_Timer()
    If nPos < 0 - lblNormal.Height Then
        nPos = picAbout.Height
    Else
        nPos = nPos - 15
    End If
    lblNormal.Top = nPos
End Sub

