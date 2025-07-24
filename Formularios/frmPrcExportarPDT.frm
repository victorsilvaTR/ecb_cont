VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.Ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPrcExportarPDT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportar Datos PDT 0601 - PLAME"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   Icon            =   "frmPrcExportarPDT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   5910
   Begin VB.TextBox txtPeriodo 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3420
      TabIndex        =   18
      Top             =   1080
      Width           =   1005
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5865
      Begin VB.Image Image3 
         Height          =   570
         Left            =   0
         Picture         =   "frmPrcExportarPDT.frx":0ECA
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         Caption         =   "Importación de Recibos por Honorarios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   750
         TabIndex        =   5
         Top             =   30
         Width           =   4110
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000D&
         Caption         =   "Programa de Declaración Telemática - PDT - PLAME (Interfase)"
         ForeColor       =   &H80000004&
         Height          =   285
         Left            =   750
         TabIndex        =   4
         Top             =   300
         Width           =   4890
      End
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar"
      Height          =   375
      Left            =   3435
      TabIndex        =   2
      Top             =   2355
      Width           =   1065
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4650
      TabIndex        =   1
      Top             =   2355
      Width           =   1065
   End
   Begin VB.CommandButton cmdLimpia 
      Height          =   585
      Left            =   150
      Picture         =   "frmPrcExportarPDT.frx":2E44
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Limpia la pantalla de presentación de Archivo"
      Top             =   2235
      Width           =   615
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1725
      Left            =   135
      TabIndex        =   16
      Top             =   2880
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   3043
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      RightMargin     =   60000
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmPrcExportarPDT.frx":314E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtpPeriodo 
      Height          =   315
      Left            =   3420
      TabIndex        =   17
      Top             =   1065
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      _Version        =   393216
      Format          =   101056515
      CurrentDate     =   37847
   End
   Begin VB.ListBox lstEmpresas 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   3420
      TabIndex        =   6
      Top             =   990
      Visible         =   0   'False
      Width           =   2295
   End
   Begin RichTextLib.RichTextBox RichTextBox2 
      Height          =   1365
      Left            =   120
      TabIndex        =   19
      Top             =   4920
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   2408
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      RightMargin     =   60000
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmPrcExportarPDT.frx":31CE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label9 
      Caption         =   ".PS4"
      Height          =   195
      Left            =   3000
      TabIndex        =   21
      Top             =   4710
      Width           =   375
   End
   Begin VB.Label lblArchivoPs4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "062196868068012200101"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   120
      TabIndex        =   20
      Top             =   4680
      Width           =   2835
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Prefijo de Archivo :"
      Height          =   195
      Left            =   90
      TabIndex        =   15
      Top             =   750
      Width           =   1335
   End
   Begin VB.Label lblPrefijo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0601"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1980
      TabIndex        =   14
      Top             =   720
      Width           =   825
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "RUC de Empresa :"
      Height          =   195
      Left            =   90
      TabIndex        =   12
      Top             =   1140
      Width           =   1320
   End
   Begin VB.Label lblRuc 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1234567890123"
      Height          =   285
      Left            =   1980
      TabIndex        =   11
      Top             =   1080
      Width           =   1275
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Periodo:"
      Height          =   195
      Left            =   3420
      TabIndex        =   10
      Top             =   765
      Width           =   585
   End
   Begin VB.Label lblArchivo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "062196868068012200101"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   90
      TabIndex        =   9
      Top             =   1845
      Width           =   2835
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Archivo Generado"
      Height          =   195
      Left            =   90
      TabIndex        =   8
      Top             =   1575
      Width           =   1290
   End
   Begin VB.Label Label2 
      Caption         =   ".TXT"
      Height          =   195
      Left            =   2970
      TabIndex        =   7
      Top             =   1875
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Seleccione una Empresa"
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   3420
      TabIndex        =   13
      Top             =   690
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "frmPrcExportarPDT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gsGrupo As String
Dim sSql As String
Dim lsSwich As Boolean
Dim RstDetalle As New ADODB.Recordset '****** NUEVO REGISTRO INGRESADO EL DIA 11/07/2013 - PAUL CUEVA
Dim RstPS4 As New ADODB.Recordset

Public Property Let Grupo(ByVal Grupo As String)
 gsGrupo = Grupo
End Property

Private Sub cmdLimpia_Click()
 RichTextBox1.filename = ""
 RichTextBox2.filename = ""
End Sub

Private Sub cmdProcesar_Click()
'On Error GoTo Control

  Screen.MousePointer = vbHourglass
  
    Dim NombreArchivo As String
    If Not ExistenDatos Then Exit Sub
    
'    NombreArchivo = "C:\" & lblArchivo & ".4ta"
  
    '************************ NUEVO REGISTRO INGRESADO **** MODIFICADO EL DIA 11/07/2013 - PAUL CUEVA
     Dim NombreArchivoPS4 As String
     Dim ruta As String
     Dim fso As Object
     Dim Mes As Integer
     Dim Mess As String
    
     Set fso = CreateObject("Scripting.FileSystemObject")
     
     Mes = CInt(Trim(Right$(txtPeriodo, 2)))
     If Mes = 1 Then Mess = "Enero"
     If Mes = 2 Then Mess = "Febrero"
     If Mes = 3 Then Mess = "Marzo"
     If Mes = 4 Then Mess = "Abril"
     If Mes = 5 Then Mess = "Mayo"
     If Mes = 6 Then Mess = "Junio"
     If Mes = 7 Then Mess = "Julio"
     If Mes = 8 Then Mess = "Agosto"
     If Mes = 9 Then Mess = "Setiembre"
     If Mes = 10 Then Mess = "Octubre"
     If Mes = 11 Then Mess = "Noviembre"
     If Mes = 12 Then Mess = "Diciembre"
    
     'creo la carpeta recibo por honorarios
     ruta = App.Path + "\Recibo_honorarios\"
     If fso.FolderExists(ruta) = False Then
         fso.CreateFolder (ruta)
     End If
    
    'creo la carpeta por mes
    
     ruta = App.Path + "\Recibo_honorarios\" + Mess + "\"
     If fso.FolderExists(ruta) = False Then
     fso.CreateFolder (ruta)
     End If
    
     NombreArchivo = ruta & lblArchivo & ".4ta"
     NombreArchivoPS4 = ruta & lblArchivo & ".ps4"
     
     lblArchivoPs4.Caption = lblArchivo
     
    '****************************************FIN DEL REGISTRO
        
    If MsgBox("¿Está seguro de Procesar el Archivo" & vbCrLf & vbCrLf & NombreArchivo & "?", vbYesNo + vbDefaultButton1 + vbQuestion, App.Title) = vbNo Then
       Exit Sub
    End If

    Open NombreArchivo For Output Shared As #1
    With RstDetalle
       If .RecordCount > 0 Then
          While Not .EOF
             Print #1, !registros
             .MoveNext
          Wend
       End If
    End With
    Close #1
    RichTextBox1.filename = NombreArchivo
   '******************************************** NUEVO REGISTRO INGRESADO **** MODIFICADO EL DIA 11/07/2013 - PAUL CUEVA
   
    If MsgBox("¿Está seguro de Procesar el Archivo" & vbCrLf & vbCrLf & NombreArchivoPS4 & "?", vbYesNo + vbDefaultButton1 + vbQuestion, App.Title) = vbNo Then
       Exit Sub
    End If
        
    Open NombreArchivoPS4 For Output Shared As #2
    
    If Not RstPS4.EOF Then
        With RstPS4
           If .RecordCount > 0 Then
              While Not .EOF
                 Print #2, !registros
                 .MoveNext
              Wend
           End If
        End With
    End If
    
    
    Close #2
    RichTextBox2.filename = NombreArchivoPS4
     '************************************* FIN DEL REGISTRO
  Screen.MousePointer = vbDefault
  
'Exit Sub
'
'Control:
' MsgBox Err.Description
' Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSalir_Click()
 Unload Me
End Sub
Private Sub dtpPeriodo_Change()
  txtPeriodo = Format$(dtpPeriodo.Value, "yyyyMM")
  lblArchivo = lblPrefijo & Trim(txtPeriodo) & lblRuc
  lblArchivoPs4 = lblPrefijo & Trim(txtPeriodo) & lblRuc
End Sub

Private Sub Form_Load()
 Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))
 ConfigForm Me, 6000, 6950
 
 lsSwich = False
 LLenarVariables
 lsSwich = True
 GeneraArchivo
 dtpPeriodo.Value = Date
End Sub
Private Sub GeneraArchivo()
 lblArchivo = lblPrefijo & Trim(txtPeriodo) & lblRuc
 lblArchivoPs4 = lblPrefijo & Trim(txtPeriodo) & lblRuc
End Sub
Private Sub Form_Unload(Cancel As Integer)
 Set frmPrcExportarPDT = Nothing
 Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))
End Sub
Private Sub LLenarVariables()
On Error GoTo Control

   ' Llenar lista
   Dim arrDatos() As Variant
   Dim clDatos As clsMantoTablas
   Dim RstEmpresas As ADODB.Recordset
   
   Set RstEmpresas = New ADODB.Recordset
   sSql = "spCn_ConsultaDscEmpresa '" & gsEmpresa & "'"
   
   Set clDatos = New clsMantoTablas
   arrDatos = Array(sSql)
   Set RstEmpresas = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
   
   With RstEmpresas
      If .RecordCount > 0 Then
         .MoveFirst
         While Not .EOF
            lstEmpresas.AddItem !empresa
            .MoveNext
         Wend
         lstEmpresas.ListIndex = 0
         lblRuc = Trim$(Right$(lstEmpresas.Text, 20))
      Else
         lblRuc = ""
      End If
   End With
   'dtpPeriodo = Format$(fechaServidor, "dd/MM/yyyy")
   txtPeriodo = Format$(FechaServidor, "YYYYMM")
   
Exit Sub

Control:
 MsgBox Err.Description, vbCritical, App.Title
End Sub
Private Function ExistenDatos() As Boolean
On Error GoTo Err_Data

  Dim clDatos As clsMantoTablas
  Dim arrDatos() As Variant
  
  ExistenDatos = False
    
  sSql = "spCn_ConsultaImportacionPDT0601 '" & gsEmpresa & "', '" & Left(txtPeriodo, 4) & "','" & Right(txtPeriodo, 2) & "'"

  Set clDatos = New clsMantoTablas
  arrDatos = Array(sSql)
  Set RstDetalle = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())

  'Cargar el registro para el .PS4 **************** NUEVO REGISTRO INGRESADO 11/07/2013 - PAUL CUEVA
  
   'sSql = "spCn_GrabaEntidad 'BUSCARPLAME', '" & gsEmpresa & "', '', '', '', '', '', '', '', '', '', '', '','', '', '', '','','" & txtPeriodo.Text & "' "
'   MsgBox Right(Me.txtPeriodo.Text, 2)
  sSql = "spCn_GrabaEntidad 'BUSCARPLAME', '" & gsEmpresa & "', '', '', '', '', '', '', '', '', '', '', '','', '', '', '','', '" & Right(Me.txtPeriodo.Text, 2) & "'"
'  arrDatos = Array(sSql)
  
'  Set RstPS4 = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
  
  Dim ObjFuncion As ClsFuncionesExecute
  Set ObjFuncion = New ClsFuncionesExecute
  Set RstPS4 = ObjFuncion.fRetornaRS(sSql)
  

  '********************************************** FIN DEL NUEVO REGISTRO
        
    
  With RstDetalle
  If .State <> 0 Then
   If .RecordCount = 0 Then
    MsgBox "No existen Registros en el Periodo señalado.", vbInformation, App.Title
    Exit Function
   End If
  Else
    MsgBox "No existen Registros en el Periodo señalado.", vbInformation, App.Title
    Screen.MousePointer = vbDefault
    Exit Function
  End If
  End With

  ExistenDatos = True
  
Exit Function
   
Err_Data:
   Screen.MousePointer = vbDefault
   MsgBox Err.Description, vbCritical, App.Title
End Function

Private Sub txtPeriodo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.lblArchivo = lblPrefijo & Trim(txtPeriodo) & lblRuc
    End If
End Sub
