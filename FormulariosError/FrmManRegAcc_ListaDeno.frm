VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmManRegAcc_ListaDeno 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ::: Seleccione :::"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3885
   Icon            =   "FrmManRegAcc_ListaDeno.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   3885
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Lista"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3510
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   3645
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MshDeno 
         Height          =   3120
         Left            =   120
         TabIndex        =   1
         Top             =   255
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   5503
         _Version        =   393216
         BackColor       =   8454143
         ForeColor       =   128
         BackColorSel    =   128
         ForeColorSel    =   8454143
         WordWrap        =   -1  'True
         FocusRect       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   0
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
End
Attribute VB_Name = "FrmManRegAcc_ListaDeno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sTabla As String

Private Sub Form_Load()
    MshDeno.ColWidth(0) = 150
    MshDeno.ColWidth(1) = 700
    MshDeno.ColWidth(2) = 2400
If sCuenta = "31" And gsTipoPlan = "1" And sNroCol = 3 Then
    MshDeno.AddItem ("")
     MshDeno.TextMatrix(0, 0) = ""
    MshDeno.TextMatrix(0, 1) = "Codigo"
    MshDeno.TextMatrix(0, 2) = "Descripcion"
    MshDeno.TextMatrix(1, 0) = ""
    MshDeno.TextMatrix(1, 1) = "01"
    MshDeno.TextMatrix(1, 2) = "Valor Razonable"
    MshDeno.TextMatrix(2, 0) = ""
    MshDeno.TextMatrix(2, 1) = "02"
    MshDeno.TextMatrix(2, 2) = "Modelo del Costo"
Else
    Set MshDeno.DataSource = Fct_Listar_Denominaciones(sTabla, gsEmpresa)
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
sNroCol = 0
End Sub

Private Sub MshDeno_DblClick()
    Call MshDeno_KeyDown(13, 1)
End Sub

Private Sub MshDeno_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Or KeyCode = 13 Then
    If KeyCode = 27 Then Tag = "0" Else Tag = "1"
    sNroCol = 0
    Hide
End If
End Sub
