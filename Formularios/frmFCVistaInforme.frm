VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.Ocx"
Begin VB.Form frmFCVistaInforme 
   Caption         =   "Informe"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15105
   Icon            =   "frmFCVistaInforme.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8910
   ScaleWidth      =   15105
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12270
      TabIndex        =   1
      Top             =   8430
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13680
      TabIndex        =   0
      Top             =   8460
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   795
      Top             =   9570
   End
   Begin RichTextLib.RichTextBox txtInforme 
      Height          =   8130
      Left            =   90
      TabIndex        =   2
      Top             =   90
      Width           =   14745
      _ExtentX        =   26009
      _ExtentY        =   14340
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      RightMargin     =   60000
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmFCVistaInforme.frx":0ECA
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
   Begin MSComDlg.CommonDialog cmdialImpresion 
      Left            =   2115
      Top             =   8520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      PrinterDefault  =   0   'False
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   135
      Top             =   9540
      Width           =   480
   End
End
Attribute VB_Name = "frmFCVistaInforme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nImagen As Integer
Dim ContVeces As Integer

'constantes para el redimensionado de los controles
Const MARGEN As Single = 80
Const ANCHO_BOTON As Single = 2000
Const ALTO_BOTON As Single = 380

Public Sub Imprimir_DS(Path As String)
On Error Resume Next
Dim Free_File As Integer, Datos As String, Pos As Integer, L As String, Palabra As String, k As Byte, VarTexto As String
Open Path For Input As #1
Dim Linea As String, Total As String
Dim Cont As Integer
Cont = 1
L = 0
k = 0
        
Do Until EOF(1)
    Line Input #1, Linea
    
    If Left(LTrim(RTrim(Linea)), 6) = "VAN..." Or InStr(1, Linea, "TOTAL ACUM") <> 0 Then
       If Cont > 75 Then
         k = 1
       End If
    End If
    
    If Cont >= 14 And Linea <> "" Then
        Printer.Print ; Linea '; Printer.hDC
        'Debug.Print
        'If k = 1 Then k = 0: 'Printer.NewPage
        If L = 1 And k = 1 Then
            Printer.NewPage: Cont = 0
        End If
        If L = 1 Then k = 0: L = 0: 'Printer.NewPage
        If k = 1 Then L = 1
    End If
    If Cont < 14 Then
        Printer.Print ; Linea '; Printer.hDC
        'Debug.Print
        'If k = 1 Then k = 0: 'Printer.NewPage
        If L = 1 And k = 1 Then
            Printer.NewPage: Cont = 0
        End If
        If L = 1 Then k = 0: L = 0: 'Printer.NewPage
        If k = 1 Then L = 1
    End If
    Cont = Cont + 1
Loop

Printer.EndDoc
Screen.MousePointer = vbNormal


Close #1
Exit Sub

     ' número de archivo libre
     Free_File = FreeFile

     ' abre el archivo para leerlo
     Open Path For Input As Free_File

     ' Almacena los datos del archivo en la variable
     Datos = Input(LOF(Free_File), Free_File)
     ' cierra el archivo
     
     Close Free_File

     Do While Len(Datos) > 0
        Pos = InStr(Datos, vbCrLf)
        If Pos = 0 Then
            L = Datos
            Datos = ""
        Else
                 ' linea
            L = Left$(Datos, 232)
            Datos = Mid$(Datos, 233 + 1)
        End If
        
        If (Printer.CurrentX + Printer.TextWidth(L)) <= Printer.ScaleWidth Then
        ' imprime la palabra
            If Left(L, 13) = "   TOTAL ACUM" Then
                k = 1
            End If

            Printer.Print L
            If k = 1 Then k = 0: Printer.NewPage
            'Printer.Print Palabra;
            ' si no imprime en la siguiente linea
        End If
     Loop

Printer.EndDoc
Screen.MousePointer = vbNormal

End Sub

Public Sub Imprimir_DM(Path As String)
Dim Free_File As Integer, Datos As String, Pos As Integer, L As String, Palabra As String, k As Byte, VarTexto As String
Open Path For Input As #1
Dim Linea As String, Total As String
Dim Cont As Integer
Cont = 1
L = 0
k = 0


Do Until EOF(1)
    Line Input #1, Linea

    If Left(LTrim(RTrim(Linea)), 6) = "VAN..." Or InStr(1, Linea, "TOTALES") <> 0 Then
        k = 1
    End If
    Printer.Print ; Linea
    
    'Salto de Linea si encuentra la palabra VAN o TOTALES
'    If (InStr(Linea, "VAN") > 0 Or InStr(Linea, "TOTALES") > 0) Then
'        Printer.Print vbCrLf
'    End If
    
    'Debug.Print Linea
    'If k = 1 Then k = 0: 'Printer.NewPage
    If Cont >= 76 Or k = 1 Then Printer.NewPage: Cont = 0: k = 0 'Printer.NewPage: cont = 0
    'If L = 1 Then k = 0: L = 0: 'Printer.NewPage
    'If k = 1 Then L = 1
'    End If
    Cont = Cont + 1
Loop

Printer.EndDoc
Screen.MousePointer = vbNormal
Close #1
End Sub

Public Sub Imprimir(Path As String)
Dim Free_File As Integer, Datos As String, Pos As Integer, L As String, Palabra As String, k As Byte, VarTexto As String

Open Path For Input As #1
Dim Linea As String, Total As String
Dim Cont As Integer
Cont = 1
L = 0
k = 0
Do Until EOF(1)
    Line Input #1, Linea
    
    If Left(LTrim(RTrim(Linea)), 6) = "VAN..." Then
        k = 1
    End If
    
    Printer.Print ; Linea
    'If k = 1 Then k = 0: 'Printer.NewPage
    If Cont >= 76 And k = 1 Then Printer.NewPage: Cont = 0
    If L = 1 Then k = 0: L = 0: 'Printer.NewPage
    If k = 1 Then L = 1
'    End If
    Cont = Cont + 1
Loop

Printer.EndDoc
Screen.MousePointer = vbNormal
Close #1
Exit Sub

     ' número de archivo libre
     Free_File = FreeFile

     ' abre el archivo para leerlo
     Open Path For Input As Free_File

     ' Almacena los datos del archivo en la variable
     Datos = Input(LOF(Free_File), Free_File)
     ' cierra el archivo
     
     Close Free_File
     Do While Len(Datos) > 0

        Pos = InStr(Datos, vbCrLf)
        If Pos = 0 Then
            L = Datos
            Datos = ""
        Else
                 ' linea
        L = Left$(Datos, Pos - 1)
        
        Datos = Mid$(Datos, Pos + 2)
        End If
                 ' palabras
                 Do While Len(L) > 0
                     ' posición para extraer la palabra
                     Pos = InStr(L, " ")
                    If Pos = 0 Then
                        Palabra = L
                        L = ""
                    Else
                        Palabra = Left$(L, Pos)
                        L = Mid$(L, Pos + 1)
                    End If
                         ' verifica que no se pase del ancho de la hoja
                    If (Printer.CurrentX + Printer.TextWidth(Palabra)) <= Printer.ScaleWidth Then
                             ' imprime la palabra
                             If Left(LTrim(RTrim(Palabra)), 6) = "VAN..." Then
                                k = 1
                            End If
                             VarTexto = VarTexto & IIf(Palabra = "", " ", Palabra)
                    End If
                Loop
                Printer.Print VarTexto
                VarTexto = ""
                
                If k = 1 Then k = 0: Printer.NewPage
             Loop
             ' Fin. Manda a imprimir
            Printer.EndDoc
    Screen.MousePointer = vbNormal
 End Sub

Public Sub cmdImprimir_Click()


On Error GoTo ErrorImp
Dim loImpresion As ExeCmdDos
Dim sCadImp As String




'      sCadImp = "C:\ECBWIN\imprimir.bat"
'      Set loImpresion = New ExeCmdDos
'      loImpresion.ExecCmd (sCadImp)
'      Set loImpresion = Nothing

   Image1.Visible = True
   gsPagina = 0
   Screen.MousePointer = vbHourglass
   On Error GoTo ErrorImp

   'Printer.CurrentX = 0.5 * 2000
   Printer.FontName = "Draft 17cpi"
   Printer.FontSize = 10
   'Printer.PaperSize = Gs_TamPapel
   Printer.PaperSize = crPaperLetter  'JCS Octalia 05-04-18
   Printer.ScaleMode = 0
   'Printer.Orientation = 1
   Printer.Orientation = 2  'JCS Octalia 05-04-18
   frmFCImpresion.List_Destino.Text = "Impresora"

   If Me.Caption = "Informe Predefinido" Then
     Printer.FontName = "Draft 17cpi"
     Printer.FontSize = 10
     Printer.PaperSize = Gs_TamPapel
     GsDestino = "Impresora"
     gsAccionRep = 1
     Screen.MousePointer = vbNormal
   Else
        If gsAccionRep = 5 Or gsAccionRep = 4 Then
            Call Imprimir(txtInforme.filename)
        ElseIf gsAccionRep = 1 Then
            Call Imprimir_DS(txtInforme.filename)
        Else
            Imprimir_DM (txtInforme.filename)
            'Call CallReports
        End If
   End If
   Exit Sub

ErrorImp:
   If Err > 0 Then
        Select Case Err.Number
          Case 380
                 MsgBox "Seleccione una impresora de carro ancho.", vbCritical, "Contabilidad"
                 Screen.MousePointer = vbNormal
               Exit Sub
          Case 482
                 MsgBox "Error en Configuración de Impresora", vbCritical, App.Title
                 Screen.MousePointer = vbNormal
               Exit Sub
          Case Else
                 MsgBox "Error general de Impresión", vbCritical, App.Title
                 Screen.MousePointer = vbNormal
               Exit Sub
       End Select
   End If

   
   
   
   
'   Image1.Visible = True
'   gsPagina = 0
'   Screen.MousePointer = vbHourglass
'   On Error GoTo ErrorImp
'
'   Printer.FontName = "Draft 17cpi"
'   Printer.FontSize = 10
'   Printer.PaperSize = Gs_TamPapel
'   Printer.ScaleMode = 0
'   Printer.Orientation = 1
'   frmFCImpresion.List_Destino.Text = "Impresora"
'
'   If Me.Caption = "Informe Predefinido" Then
'     Printer.FontName = "Draft 17cpi"
'     Printer.FontSize = 10
'     Printer.PaperSize = Gs_TamPapel
'     GsDestino = "Impresora"
'     gsAccionRep = 1
'     Screen.MousePointer = vbNormal
'   Else
'        If gsAccionRep = 5 Or gsAccionRep = 4 Then
'            Call Imprimir(txtInforme.FileName)
'        ElseIf gsAccionRep = 1 Then
'            Call Imprimir_DS(txtInforme.FileName)
'        Else
'            Call CallReports
'        End If
'   End If
End Sub

Private Sub cmdSalir_Click()
   If Me.Caption = "Informe Predefinido" Then
     Unload Me
   End If
   Unload Me
End Sub


Private Sub Form_Activate()
   frmMDIConta.stbMdi.Panels(6).Text = Me.Caption
   Screen.MousePointer = vbNormal
End Sub
Private Sub Form_Load()
   Me.Icon = frmMDIConta.Icon
   frmMDIConta.stbMdi.Panels(6).Text = Me.Caption
   Me.Move (frmMDIConta.Width - Me.Width) \ 2 - 100, (frmMDIConta.Height - Me.Height) \ 2 + 300
   ContVeces = 0
   Image1.Visible = False
   Screen.MousePointer = vbNormal
End Sub
Private Sub Form_Resize()
On Local Error Resume Next

Me.AutoRedraw = True

Me.CurrentX = MARGEN
Me.CurrentY = MARGEN

Me.CurrentX = MARGEN
Me.CurrentY = 550

cmdImprimir.Move txtInforme.Left, (Me.ScaleHeight - cmdImprimir.Height - MARGEN), _
                                ANCHO_BOTON - MARGEN * 6, ALTO_BOTON

cmdSalir.Move (cmdImprimir.Left + cmdImprimir.Width) + 50, _
                         (Me.ScaleHeight - cmdImprimir.Height - MARGEN), _
                          ANCHO_BOTON - MARGEN * 6, ALTO_BOTON
                          
txtInforme.Move txtInforme.Left, txtInforme.Top, _
                        Me.ScaleWidth - MARGEN * 2, _
                        (Me.ScaleHeight - cmdImprimir.Height) - (MARGEN * 3)

End Sub
Private Sub Form_Unload(Cancel As Integer)
 If Forms.Count > 2 Then Exit Sub
 frmMDIConta.stbMdi.Panels(6).Text = App.Title
End Sub

