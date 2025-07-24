VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFCImpresion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Impresión"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   Icon            =   "frmFCImpresion.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   6480
   Begin VB.CommandButton cmdGeneraReporte 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Generar Reporte"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   45
      TabIndex        =   14
      Top             =   4755
      Width           =   2000
   End
   Begin VB.CommandButton cmdCancelar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4425
      TabIndex        =   20
      Top             =   4755
      Width           =   2000
   End
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   45
      TabIndex        =   11
      Top             =   0
      Width           =   6375
      Begin VB.TextBox txtPrincipal 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5700
         TabIndex        =   23
         Text            =   "1"
         Top             =   225
         Width           =   525
      End
      Begin VB.TextBox txtHasta 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5700
         TabIndex        =   21
         Text            =   "0"
         Top             =   630
         Width           =   525
      End
      Begin VB.ComboBox List_Destino 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1845
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   420
         Width           =   1740
      End
      Begin VB.TextBox Num_copies 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5715
         MaxLength       =   2
         TabIndex        =   15
         Text            =   "1"
         Top             =   1140
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.CheckBox chkCarroAncho 
         Caption         =   "Impresora de Carro Anc&ho"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   120
         TabIndex        =   12
         Top             =   675
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   2715
      End
      Begin VB.TextBox txtDesde 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4575
         TabIndex        =   13
         Text            =   "0"
         Top             =   630
         Width           =   525
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Página Principal :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   3960
         TabIndex        =   24
         Top             =   225
         Width           =   1380
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Fin :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   5265
         TabIndex        =   22
         Top             =   660
         Width           =   330
      End
      Begin VB.Label lblDestino 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         BackStyle       =   0  'Transparent
         Caption         =   "&Destino del Reporte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   150
         TabIndex        =   19
         Top             =   465
         Width           =   1635
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Nro. Copias"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   4440
         TabIndex        =   18
         Top             =   1140
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Inicio :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   3960
         TabIndex        =   17
         Top             =   660
         Width           =   525
      End
   End
   Begin VB.Frame fraDisco 
      Height          =   3030
      Left            =   45
      TabIndex        =   5
      Top             =   1080
      Width           =   6375
      Begin VB.TextBox txtPath 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   150
         TabIndex        =   9
         Text            =   "Repor01.txt"
         Top             =   435
         Width           =   6060
      End
      Begin VB.FileListBox File1 
         Height          =   2040
         Left            =   3210
         Pattern         =   "*.txt"
         TabIndex        =   8
         Top             =   810
         Width           =   3015
      End
      Begin VB.DirListBox Dir1 
         Height          =   1665
         Left            =   150
         TabIndex        =   7
         Top             =   1200
         Width           =   2895
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   135
         TabIndex        =   6
         Top             =   825
         Width           =   2910
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre del Archivo (Incluyendo Directorio)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   165
         TabIndex        =   10
         Top             =   165
         Width           =   3555
      End
   End
   Begin VB.CommandButton CmdSetupPrinter 
      Caption         =   "Configurar &Impresora"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2160
      TabIndex        =   4
      Top             =   4755
      Width           =   2130
   End
   Begin VB.Frame fraFile 
      Height          =   600
      Left            =   45
      TabIndex        =   1
      Top             =   4110
      Width           =   6360
      Begin VB.TextBox txtNameFile 
         Height          =   315
         Left            =   1830
         TabIndex        =   2
         Top             =   180
         Width           =   4380
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre del &Archivo "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   165
         TabIndex        =   3
         Top             =   255
         Width           =   1680
      End
   End
   Begin VB.TextBox OutputFileName 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   5625
      Visible         =   0   'False
      Width           =   6210
   End
   Begin VB.Timer Timer1 
      Left            =   5955
      Top             =   4170
   End
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   60
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
End
Attribute VB_Name = "frmFCImpresion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdGeneraReporte_Click()
Dim xValida As Boolean
gsTipoImp = False
xValida = False
giCopias = Val(Num_copies.Text)

'If InStr(1, impresora(), "LQ-2180") <> 0 Then gsNomTipoImp = True
 If CDbl(txtDesde) > CDbl(txtHasta) Then
  MsgBox "La Página Final debe ser mayor a la Inicial...", vbInformation, "Verifique"
  txtHasta.SetFocus
  Exit Sub
 ElseIf CDbl(txtDesde) = 0 And CDbl(txtHasta) > CDbl(txtDesde) Then
  MsgBox "La Página Inicial debe ser diferente de cero...", vbInformation, "Verifique"
  txtDesde.SetFocus
  Exit Sub
 ElseIf CDbl(txtHasta) = 0 And CDbl(txtDesde) > CDbl(txtHasta) Then
  MsgBox "La Página Final debe ser diferente de cero...", vbInformation, "Verifique"
  txtHasta.SetFocus
  Exit Sub
 End If

 If gsAccionRep <> 1 Then
  Label4.Visible = False: txtPrincipal.Visible = False
  txtPrincipal.Text = 0
  If CDbl(txtDesde) <> 0 And CDbl(txtHasta) <> 0 Then txtPrincipal.Text = 1
 End If

If txtDesde.Text = "" Then txtDesde.Text = "0"
If txtHasta.Text = "" Then txtHasta.Text = "0"
If txtPrincipal.Text = "" Then txtPrincipal.Text = "1"

 xGs_DesdePag = CDbl(txtDesde)
 xGs_HastaPag = CDbl(txtHasta)
 xGs_Principal = CDbl(txtPrincipal)
  ' --------------------
  ' Destino del reporte
  ' --------------------
  
  If List_Destino.Text = "Archivo" Then
     
     On Error GoTo ErrorFileExist
     
     If Len(Trim(txtNameFile.Text)) = 0 Then
        MsgBox "Seleccione un Archivo", vbExclamation, App.Title
        fraDisco.Enabled = True
        txtNameFile.Enabled = True
        txtNameFile.SelStart = Len(OutputFileName.Text)
        txtNameFile.SetFocus
     Else
        'OutputFileName.Text = Trim(txtPath) & Trim(txtNameFile) & ".txt"
        OutputFileName.Text = Trim(txtPath) & Trim(txtNameFile) & ".txt"
        'Dim fso As New FileSystemObject

        ' Intento abrir el archivo
        Open OutputFileName For Output Shared As #1
        Close #1
        
        File1.Refresh
        xValida = True
     End If
     ' Valido
  Else
     On Error GoTo ErrorInPrint
        If Val(Trim$(Num_copies)) >= 1 Then
                xValida = True
                OutputFileName.Text = Trim(txtPath) & Trim(txtNameFile) & ".txt"
                ' Intento abrir el archivo
                Open OutputFileName For Output Shared As #1
                Close #1
                File1.Refresh
                xValida = True
               'Printer.FontName = "Draft 17cpi"
               'Printer.FontSize = 10
               'Printer.PaperSize = Gs_TamPapel
               'Printer.Orientation = 1
        Else
          MsgBox "Verifique el Número de Copias.", 48, App.Title
          Num_copies.SetFocus
          Num_copies.Text = "1"
          Num_copies.SetFocus
        End If
        
  End If
  
  If xValida Then
    Dim textoTmp As String
    Dim rutaImp As String
    Dim impini As String
    
    Gi_FlagImpresion = 1
          frmMDIConta.stbMdi.Panels(6).Text = "Generando Informe..."
          
          Call CallReports
          
          Dim Indice As Integer
        
'          If InStr(1, textoTmp, "\\") Then
'            rutaImp = InStrRev(textoTmp, "\")
'            rutaImp = Left(textoTmp, rutaImp)
'          Else
'            rutaImp = "\\" + Trim(NomPC) + "\" '+ textoTmp
'          End If
'          Guardar_Archivo "C:\ECBWIN\imprimir.bat", TextBat
'          textoTmp = Abrir_ArchivoBat("C:\ECBWIN\imprimir.bat")
'          TextBat = "TYPE " & Chr(34) & txtPath.Text & txtNameFile & ".txt" & Chr(34) & "> " & rutaImp & impini 'Mid(textoTmp, Indice, Len(textoTmp))
'          Guardar_Archivo "C:\ECBWIN\imprimir.bat", TextBat
         
          Screen.MousePointer = vbNormal
          rutaImp = ""
          Exit Sub
    End If
'End If
  
ErrorFileExist:
   If Err > 0 Then
      Select Case Err
         Case 55
            MsgBox "El Archivo ya esta abierto.", vbExclamation, App.Title
            Close #1
         Case 75
            MsgBox "Error en la selección del Directorio o del Archivo", vbCritical, App.Title
         Case 94
            MsgBox Err.Description, vbCritical, App.Title
         Case 20545
'                 Call ERRORGRAL.Mensajes("4087")
         
            Case 482
'                 Call ERRORGRAL.Mensajes("4088")
            Case 380
'                 Call ERRORGRAL.Mensajes("4087")
            Case Else
                 MsgBox "Error al intentar procesar el archivo temporal", vbCritical, App.Title
                 'MsgBox Str(Err) + " " + Error$
     End Select
     Screen.MousePointer = vbNormal
     frmMDIConta.stbMdi.Panels(6).Text = ""
     Exit Sub
  End If

ErrorInPrint:
  If Err > 0 Then
        Select Case Err
              Case 380
'                     Call ERRORGRAL.Mensajes("4090")
                  MsgBox "Seleccione un impresora de Carro Ancho.", vbCritical, "Contabilidad"
               Case 482
                    MsgBox "Error en Configuración de Impresora", vbCritical, App.Title
               Case Else
                    MsgBox "Error general de Impresión", vbCritical, App.Title
        End Select
        Screen.MousePointer = vbNormal
  End If
 
End Sub


Private Sub CmdSetupPrinter_Click()
'FrmConfigImp.Show 1
'Exit Sub

If Val(Trim$(Num_copies)) >= 1 Then
Else
  MsgBox "Verifique el Número de Copias.", 48, App.Title
  Num_copies.SetFocus
  Num_copies.Text = "1"
  Num_copies.SetFocus
  Exit Sub
End If
'  '----------------------------------------
'  ' Rutina de configuración de impresora
'  '----------------------------------------
   On Error Resume Next
   
   Dim P As New clsPrintDialog
   
   P.Flags = cdlPDPrintSetup
   P.ShowPrinter
   
   'Debug.Print P.
   Err = 0
   Set P = Nothing
   
End Sub
Private Sub Dir1_Change()
   File1.Path = Dir1.Path
   txtNameFile = ""
   File1_Click
End Sub

Private Sub Drive1_Change()
On Error GoTo ErrorDrv
   Dir1.Path = Drive1.Drive
   txtNameFile.Text = ""
   Exit Sub
   
ErrorDrv:
Select Case Err.Number
Case 61       'diskette lleno
'   Call ERRORGRAL.Mensajes("4085")
   Exit Sub
Case 68  ' no hay diskette
'   Call ERRORGRAL.Mensajes("4084")
   Drive1.Drive = "c"
   Exit Sub
Case 70  ' no hay privilegios de acceso, escritura
'   Call ERRORGRAL.Mensajes("4086")
   Exit Sub
End Select
End Sub
Private Sub Drive1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then Dir1.Path = Drive1.Drive
End Sub
Private Sub File1_Click()

txtPath = File1.Path
If Mid(Trim(File1.Path), Len(Trim(File1.Path)), 1) <> "\" Then txtPath = File1.Path & "\"

If Len(File1.filename) > 0 Then
   txtNameFile = Mid(File1.filename, 1, Len(File1.filename) - 4)
   txtNameFile.Enabled = True
   txtNameFile.SelStart = Len(txtNameFile.Text)
   txtNameFile.SetFocus
End If
End Sub

Private Sub Form_Activate()

Me.Icon = frmMDIConta.Icon
frmMDIConta.stbMdi.Panels(6).Text = Me.Caption

If gsNombreVista = "" Then
 txtNameFile.Text = "Diario Simplificado"
Else
 txtNameFile.Text = gsNombreVista
End If

Dim PauseTime, Start, Finish, TotalTime
   
End Sub
Private Sub Form_Load()
Timer1.Interval = 250
  
Screen.MousePointer = vbHourglass

  Me.Move (frmMDIConta.Width - Me.Width) \ 2 - 100, (frmMDIConta.Height - Me.Height) \ 2 + 300
  frmMDIConta.stbMdi.Panels(6).Text = Me.Caption
  
  ' Inicialización de Variables
  '-----------------------------
  
  If gsAccionRep <> 1 Then
   Label4.Visible = False: txtPrincipal.Visible = False
   txtPrincipal.Text = 0
   If CDbl(txtDesde) <> 0 And CDbl(txtHasta) <> 0 Then txtPrincipal.Text = 1
  End If
  
  If gsAccionRep = 1 Then txtPrincipal.Text = 0
  
  Select Case gsAccionRep
   Case "1": OutputFileName = "Diario Simplificado.txt"
   Case "2": OutputFileName = "Libro Banco Detalle Efectivo.txt"
   Case "3": OutputFileName = "Libro Banco Detalle Movimientos Cta.Cte.txt"
   Case "4": OutputFileName = "Libro Diario.txt"
   Case "5": OutputFileName = "Libro Mayor General.txt"
   Case "6": OutputFileName = "Libro Mayor Analitico.txt"
   Case Else: OutputFileName = "Reporte01.txt"
  End Select
    
  Drive1.Drive = "C:"
'  ExisteFile ("C:\ECBWIN\Rpts Formato Matricial")
'Dir1.Path = "C:\ECBWIN\Rpts Formato Matricial"
Dir1.Path = "C:\"

  File1.Path = Dir1.Path
  Num_copies = 1
  fraDisco.Enabled = True 'FALSE
  fraFile.Enabled = True 'FALSE
  Num_copies.Enabled = True
  giBold% = 0
  Gi_FlagImpresion = 0
  Gs_TamPapel = Printer.PaperSize
  List_Destino.AddItem "Impresora"
  List_Destino.AddItem "Archivo"
  List_Destino.ListIndex = 1 'Impresora 1  'ARCHIVO
  Screen.MousePointer = vbNormal
  If Gs_TamPapel = 7 Then
     chkCarroAncho.Value = 1
  Else
     chkCarroAncho.Value = 0
  End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
If Not lsLibShow Then
   frmMDIConta.stbMdi.Panels(6).Text = ""
   Set frmFCImpresion = Nothing
End If
End Sub


Private Sub List_Destino_Click()
Select Case List_Destino.Text
Case "Impresora"
   OutputFileName.Enabled = False
   fraDisco.Enabled = False
   fraFile.Enabled = False
   'txtNameFile.Text = ""
   If File1.ListIndex >= 0 Then File1.Selected(File1.ListIndex) = False
   Num_copies.Enabled = True
   CmdSetupPrinter.Enabled = True
   chkCarroAncho.Enabled = True
   chkCarroAncho.Value = vbChecked
   GsDestino = "Impresora"
   txtDesde.Enabled = True
   'txtDesde = 1
   txtHasta.Enabled = True
   'txtHasta = 1
Case "Archivo"
   Num_copies = 1
   fraDisco.Enabled = True
   fraFile.Enabled = True
   txtNameFile.Enabled = True

   txtNameFile.SelStart = Len(txtNameFile.Text)
   If File1.ListIndex > 0 Then File1.Selected(File1.ListIndex) = True
   Num_copies.Enabled = True 'false
   CmdSetupPrinter.Enabled = True 'false
   chkCarroAncho.Enabled = True 'false
   GsDestino = "Archivo"
   txtDesde.Enabled = True
   'txtDesde = 0
   txtHasta.Enabled = True
   'txtHasta = 0
End Select
End Sub

Private Sub Num_copies_GotFocus()
  Num_copies.SelStart = 0
  Num_copies.SelLength = Len(Num_copies)
End Sub

Private Sub Num_copies_KeyPress(KeyAscii As Integer)
If KeyAscii = "8" Then
   If Len(Num_copies.Text) > 0 Then
      Num_copies.Text = Mid$(Num_copies.Text, 1, Len(Num_copies.Text) - 1)
      Num_copies.SelStart = Len(Num_copies.Text)
      Num_copies.SetFocus
   Else
      Num_copies.Text = ""
      Num_copies.SetFocus
   End If
End If
If Not (Chr(KeyAscii) >= "0" And Chr(KeyAscii) <= "9") Then
   KeyAscii = 0
End If
End Sub

Private Sub Num_copies_LostFocus()
Dim iCopias As Integer
iCopias = Val(Num_copies.Text)
End Sub


Private Sub Timer1_Timer()
 Timer1.Enabled = False
End Sub
Private Sub txtDesde_Change()
 If txtDesde = "" Then txtDesde = 0
End Sub

Private Sub txtDesde_GotFocus()
 SelectedText txtDesde
End Sub

Private Sub txtDesde_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      'SendKeys "{tab}"
      txtHasta.SetFocus
   Else
      If Not KeyAscii = 45 Then
          If Not ValidaSoloNumeros(KeyAscii) Then
             KeyAscii = nGetIniValueAscii
          End If
      End If
   End If
End Sub
Private Sub txtHasta_Change()
 If txtHasta = "" Then txtHasta = 0
End Sub

Private Sub txtHasta_GotFocus()
 SelectedText txtHasta
End Sub
Private Sub txtHasta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
      'SendKeys "{tab}"
      cmdGeneraReporte.SetFocus
   Else
      If Not KeyAscii = 45 Then
          If Not ValidaSoloNumeros(KeyAscii) Then
             KeyAscii = nGetIniValueAscii
          End If
      End If
   End If
End Sub

Private Sub txtNameFile_KeyPress(KeyAscii As Integer)
Dim bTecla As Boolean
    bTecla = False
    If KeyAscii = 95 Then bTecla = True
    If KeyAscii >= 48 And KeyAscii <= 57 Then
      bTecla = True
    ElseIf KeyAscii >= 65 And KeyAscii <= 90 Then
       bTecla = True
    ElseIf KeyAscii >= 97 And KeyAscii <= 122 Then
       bTecla = True
    ElseIf KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 243 Or KeyAscii = 225 Or _
       KeyAscii = 233 Or KeyAscii = 241 Or KeyAscii = 211 Or _
       KeyAscii = 237 Or KeyAscii = 45 Or KeyAscii = 180 Or KeyAscii = 209 Or _
       KeyAscii = 250 Or KeyAscii = 22 Then
       bTecla = True
    End If
    If bTecla = False Then
      KeyAscii = 0
    End If
End Sub
Private Sub txtPrincipal_Change()
 If txtDesde = "" Then txtDesde = 0
End Sub
Private Sub txtPrincipal_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      'SendKeys "{tab}"
      txtDesde.SetFocus
   Else
      If Not KeyAscii = 45 Then
          If Not ValidaSoloNumeros(KeyAscii) Then
             KeyAscii = nGetIniValueAscii
          End If
      End If
   End If
End Sub
