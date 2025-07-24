VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmExportPLe 
   Caption         =   " ::: Exportar PLE :::"
   ClientHeight    =   6264
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   7380
   Icon            =   "FrmExportPLe.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6264
   ScaleWidth      =   7380
   Begin VB.Frame FraTodo 
      Caption         =   "Libros Exportables"
      Height          =   6090
      Left            =   150
      TabIndex        =   0
      Top             =   90
      Width           =   7110
      Begin VB.CheckBox ChkLibCjaBan 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Libro de Caja y Bancos"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   900
         TabIndex        =   1
         Top             =   420
         Width           =   2505
      End
      Begin MSForms.CommandButton cmdImprimir 
         Height          =   435
         Left            =   3840
         TabIndex        =   4
         Top             =   5520
         Width           =   1530
         Caption         =   " Exportar"
         PicturePosition =   327683
         Size            =   "2699;767"
         Picture         =   "FrmExportPLe.frx":0ECA
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdSalir 
         Height          =   435
         Left            =   5355
         TabIndex        =   3
         Top             =   5520
         Width           =   1665
         Caption         =   " Salir"
         PicturePosition =   327683
         Size            =   "2937;767"
         Picture         =   "FrmExportPLe.frx":1264
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
         Size            =   15.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   330
      TabIndex        =   2
      Top             =   180
      Width           =   4365
   End
End
Attribute VB_Name = "FrmExportPLe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Me.Width = 7500
Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))

Me.Caption = Titulo(Me.Caption, "")

Call Centrar_form(Me)
End Sub

Private Sub Form_Resize()
On Error GoTo errHand
    If Me.WindowState <> vbMinimized Then
        Call SeteaFondoForm(Me)
        Call Centrar_Objeto(FraTodo, Me)
        Call CentrarTitulo(lblTitulo, FraTodo, Me)
    End If
Exit Sub
errHand:
End Sub

