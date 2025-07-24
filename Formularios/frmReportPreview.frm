VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "CRViewer.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmReportPreview 
   ClientHeight    =   5370
   ClientLeft      =   2700
   ClientTop       =   4530
   ClientWidth     =   11025
   Icon            =   "frmReportPreview.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5370
   ScaleWidth      =   11025
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.ComboBox cboZoom 
      Height          =   315
      ItemData        =   "frmReportPreview.frx":0ECA
      Left            =   6570
      List            =   "frmReportPreview.frx":0EE6
      TabIndex        =   2
      Text            =   "cboZoom"
      Top             =   0
      Width           =   960
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   135
      Picture         =   "frmReportPreview.frx":0F10
      ScaleHeight     =   285
      ScaleWidth      =   330
      TabIndex        =   1
      ToolTipText     =   " Configurar Impresora "
      Top             =   45
      Width           =   330
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   3000
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   1845
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11130
      lastProp        =   500
      _cx             =   5080
      _cy             =   5080
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   0   'False
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   5400
      Picture         =   "frmReportPreview.frx":149A
      Top             =   360
      Width           =   240
   End
End
Attribute VB_Name = "frmReportPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oReporte As New CRAXDRT.Report
Public Orientacion As Orientacion_Pagina
Public TipoPagina As Tipo_Pagina

Private Sub Picture1_Click()
    
    oReporte.PrinterSetup Me.hwnd
    CRViewer1.ReportSource = oReporte
    CRViewer1.ViewReport
End Sub

Private Sub CRViewer1_PrintButtonClicked(UseDefault As Boolean)


'    CommonDialog.Flags = cdlPDPrintSetup Or cdlPDReturnIC
    ''CommonDialog.Flags = cdlPDPrintSetup
    
   '' CommonDialog.ShowPrinter
End Sub

Private Sub Form_Activate()
cboZoom.Text = "120"
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        CRViewer1.EnableGroupTree = False
        CRViewer1.EnableAnimationCtrl = False
        CRViewer1.EnablePopupMenu = True
        CRViewer1.EnableProgressControl = False
        CRViewer1.EnableZoomControl = False
    
        CRViewer1.Top = 0
        CRViewer1.Left = 0
        CRViewer1.Height = ScaleHeight
        CRViewer1.Width = ScaleWidth
    End If
End Sub

Private Function fnZoom(cad As String) As Integer
    If NE(cboZoom.Text) > 400 Then cboZoom.Text = "100"
    fnZoom = Round(NE(cboZoom.Text), 0) 'CE(Round(Val(NE(Left(cad, Len(cad) - 1))), 0))
End Function

Private Sub cboZoom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cboZoom.Text = CE(fnZoom(cboZoom.Text))
        Call cboZoom_Click
    End If
End Sub

Private Sub cboZoom_Click()
    Dim nZoom  As Integer
    nZoom = fnZoom(cboZoom.Text)
    
    If nZoom < 50 Or nZoom > 400 Then
        cboZoom.Text = "100"
        nZoom = 100
    End If
    
    CRViewer1.ReportSource = oReporte
    CRViewer1.ViewReport
    CRViewer1.Zoom nZoom
End Sub

Public Sub SetReporte(rptReporteCrystal As CRAXDRT.Report)
On Error GoTo Control
    Screen.MousePointer = vbHourglass
    Set oReporte = rptReporteCrystal
             
    Dim rpt As New CRAXDRT.Report
    
    Set rpt = rptReporteCrystal
    
    Select Case Orientacion
        Case Orientacion_Pagina.Horizontal: oReporte.PaperOrientation = crLandscape
        Case Orientacion_Pagina.Vertical: oReporte.PaperOrientation = crPortrait
    End Select
    
    
    Select Case TipoPagina
        Case Tipo_Pagina.A4: oReporte.PaperSize = crPaperA4
        Case Tipo_Pagina.CARTA: oReporte.PaperSize = crPaperLetter
        Case Tipo_Pagina.OFICIO: oReporte.PaperSize = crPaperLegal
        Case Tipo_Pagina.USA: oReporte.PaperSize = crPaperFanfoldUS
    End Select
    
    CRViewer1.ReportSource = rpt
              
    
    
    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Set clDatos = New clsMantoTablas
    Dim arrDatos() As Variant
    Dim sqlSp  As String
    Set clDatos = New clsMantoTablas
    sqlSp = "SELECT Count(1) AS Pla_cNoTit8_ResultEjer FROM CND_CONFIG_OPERA WITH(READUNCOMMITTED) WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & _
            "' AND Cop_cCodigo = '035'"
    arrDatos = Array(sqlSp)

    Call CerrarRecordSet(rsArreglo)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo.State = 0 Then     ' *** Si no tiene datos
        Set rsArreglo = Nothing
    Else
    
     If NE(rsArreglo!Pla_cNoTit8_ResultEjer) = 0 And NombreReporte = "F0301" Then
        MsgBox "Falta configurar la Cuenta de Resultado del Ejercicio dentro del Plan de Cuentas, Active la casilla: " & Chr(10) + Chr(13) & _
               "Resultado del Ejercicio para la Cta. que corresponda.", vbInformation, App.Title
        Screen.MousePointer = vbDefault
        Set rsArreglo = Nothing
        Exit Sub
     End If
    End If
        
    CRViewer1.ViewReport
    CRViewer1.Zoom 150
    
    DoEvents
    If oReporte.PrintingStatus.NumberOfRecordRead = 0 Then
    '    Unload Me
        Mensajes "No existen datos a imprimir."
    End If
    'Else
        Me.WindowState = vbNormal

        Me.Visible = True
        Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))
        'CRViewer1.EnableAnimationCtrl = True
    'End If
    
    Screen.MousePointer = vbDefault
Exit Sub
Control:
 MsgBox Err.Description
 Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))
    'Set CRViewer1 = Nothing
    Set oReporte = Nothing
End Sub

