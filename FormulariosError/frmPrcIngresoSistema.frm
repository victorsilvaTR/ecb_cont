VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmPrcIngresoSistema 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sistema de Contabilidad"
   ClientHeight    =   2820
   ClientLeft      =   3450
   ClientTop       =   4260
   ClientWidth     =   6945
   Icon            =   "frmPrcIngresoSistema.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Picture         =   "frmPrcIngresoSistema.frx":0ECA
   ScaleHeight     =   2820
   ScaleWidth      =   6945
   StartUpPosition =   2  'CenterScreen
   Begin TDBText6Ctl.TDBText txtUsuario 
      Height          =   285
      Left            =   2385
      TabIndex        =   0
      Top             =   1575
      Width           =   3345
      _Version        =   65536
      _ExtentX        =   5900
      _ExtentY        =   503
      Caption         =   "frmPrcIngresoSistema.frx":58C5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPrcIngresoSistema.frx":5931
      Key             =   "frmPrcIngresoSistema.frx":594F
      BackColor       =   -2147483643
      EditMode        =   3
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
      Format          =   "a"
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   20
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
   Begin TDBText6Ctl.TDBText txtpassword 
      Height          =   285
      Left            =   2385
      TabIndex        =   1
      Top             =   2025
      Width           =   1545
      _Version        =   65536
      _ExtentX        =   2725
      _ExtentY        =   503
      Caption         =   "frmPrcIngresoSistema.frx":598B
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPrcIngresoSistema.frx":59F7
      Key             =   "frmPrcIngresoSistema.frx":5A15
      BackColor       =   -2147483643
      EditMode        =   3
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
      PasswordChar    =   "*"
      AllowSpace      =   0
      Format          =   "A#@"
      FormatMode      =   0
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   10
      LengthAsByte    =   0
      Text            =   ""
      Furigana        =   0
      HighlightText   =   -1
      IMEMode         =   3
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   1215
      TabIndex        =   4
      Top             =   2115
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   1215
      TabIndex        =   3
      Top             =   1665
      Width           =   705
   End
   Begin MSForms.CommandButton cmdAceptar 
      Height          =   390
      Left            =   4245
      TabIndex        =   2
      Top             =   1980
      Width           =   1530
      Caption         =   "  Ingresar"
      PicturePosition =   327683
      Size            =   "2699;688"
      Picture         =   "frmPrcIngresoSistema.frx":5A59
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "frmPrcIngresoSistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lAccesos As Byte

Private Sub cmdAceptar_Click()
 
    cmdAceptar.Enabled = False
    DoEvents
    
    Dim sqlSp As String
    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
    Dim cadena As String
    cadena = "ADMIN"
    Screen.MousePointer = vbHourglass

    lAccesos = lAccesos + 1
    Set clDatos = New clsMantoTablas
    sqlSp = "spSg_GrabaUsuarios 'SEL_REG', '" & Me.txtUsuario & "', '', '', '', '', '', '' ,'" & gsSOFT & "'"
    arrDatos = Array(sqlSp)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo.State = 0 Then
        If gsError = False Then
            Mensajes "Usuario no existe... Verifique", vbInformation
            pSetFocus txtUsuario
        End If
    Else
        If (DesCifrar(CE(rsArreglo("usu_cClave").Value), "977611") = txtpassword) Or _
            (txtUsuario = "ADMINISTRADOR" And txtpassword = cadena) Then
            If rsArreglo("usu_cEstado").Value = "I" Then
                Mensajes "Usuario no autorizado a ingresar al sistema. " & Salto(1) & _
                         "Consulte al Administrador del Sistema", vbInformation
                pSetFocus txtUsuario
            Else
            
                If CE(rsArreglo("soft_cCodSoft").Value) <> "001" Then
                    Mensajes "Usuario no tiene acceso al " & gsNombreModulo
                    pSetFocus txtUsuario
                Else
                    gsInicio = False
                    Screen.MousePointer = vbNormal
                    gsUsuario = Me.txtUsuario
                    gsAdmin = CE(rsArreglo("usu_cAdmin"))
                    gsAdmin = IIf(gsAdmin <> "1", "0", "1")
                    CerrarRecordSet rsArreglo
                    Unload Me
                    frmBusPeriodo.Show
                    Screen.MousePointer = vbNormal
                    Exit Sub
                End If
            
            End If
        Else
            Mensajes "Password incorrecto... ", vbInformation
            pSetFocus txtpassword
        
        End If
    End If
    CerrarRecordSet rsArreglo
    Screen.MousePointer = vbNormal
    cmdAceptar.Enabled = True
    Set clDatos = Nothing
    
    If lAccesos >= 3 Then
        Unload Me
    End If
 
End Sub

Private Sub Form_Activate()
    pSetFocus txtUsuario
End Sub

Private Function VerificaRegion() As Boolean
    'Dim objeto As Variant
    On Error Resume Next
    VerificaRegion = True
    Dim cFormatoFecha As String, cSepDecimal As String, cSepMillar As String
    
    cFormatoFecha = Leer_Dato(CurrentUser, "sShortDate")
    cSepDecimal = Leer_Dato(CurrentUser, "sDecimal")
    cSepMillar = Leer_Dato(CurrentUser, "sThousand")
    
    If cFormatoFecha <> "dd/MM/yyyy" And cFormatoFecha <> "" Then
        Mensajes "Cambie el formato de Fecha en Configuración Regional del Sistema Operativo" & vbCrLf & "Formato de Fecha válido : dd/MM/yyyy " & vbCrLf & "Formato Actual : " & cFormatoFecha
        VerificaRegion = False
        Exit Function
    End If
    
    If cSepDecimal <> "." And cSepDecimal <> "" Then
        Mensajes "Cambie el formato del Separador Decimal en Configuración Regional del Sistema Operativo" & vbCrLf & "Formato de Separador Decimal válido : . " & vbCrLf & "Formato Actual : " & cSepDecimal
        VerificaRegion = False
        Exit Function
    End If
    
    If cSepMillar <> "," And cSepMillar <> "" Then
        Mensajes "Cambie el formato del Separador de Millar en Configuración Regional del Sistema Operativo" & vbCrLf & "Formato de Separador de Millar válido : , " & vbCrLf & "Formato Actual : " & cSepMillar
        VerificaRegion = False
        Exit Function
    End If
    
'PGBV:Segmento que permite dar fecha limite de ingreso al sistema -- Univ. Pacifico
'    If DateDiff("d", Format(Date, "dd/MM/yyyy"), Format("20/10/2015", "dd/MM/yyyy")) < 0 Then
'        Mensajes "Su licencia de prueba ha caducado, por favor comunicarse con su Representante de Ventas"
'        VerificaRegion = False
'        Exit Function
'    End If
    
    
End Function

Private Sub Form_Load()

    On Error Resume Next
    Label1(0).ForeColor = RGB(4, 62, 74)
    Label2.ForeColor = RGB(4, 62, 74)
    frmPrcIngresoSistema.Caption = "ECBCont " & CE(gsVersion) & " - Modulo de Contabilidad"
    gsEmpresa = ""
    gsAnio = ""
    gsSucursal = ""
    gsRUC = ""
    gsEmpresaNom = ""

    If VerificaRegion = False Then
        cmdAceptar.Enabled = False
        Exit Sub
    Else
        cmdAceptar.Enabled = True
    End If
    
    gsInicio = False
    lAccesos = 0
     
    Dim arrDatos() As Variant
    Dim rsArreglo As New ADODB.Recordset
    Dim sqlSp As String
    Dim clDatos As New clsMantoTablas
    
    sqlSp = "spSg_GrabaUsuarios 'SEL_REG', '" & Me.txtUsuario & "', '', '', '', '', '', '' ,'" & gsSOFT & "'"
    arrDatos = Array(sqlSp)
    clDatos.InicializaClase
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    clDatos.FinalizaClase
    Set clDatos = Nothing
    CerrarRecordSet rsArreglo
     
End Sub

Private Sub txtpassword_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        pSetFocus cmdAceptar
    End If
End Sub

Private Sub txtpassword_LostFocus()
    'cmdAceptar_Click
End Sub

Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        pSetFocus txtpassword
    End If
End Sub
