VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Begin VB.Form FrmManRegAcc_IngNum 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ::: Ingresar Dato :::"
   ClientHeight    =   795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2460
   Icon            =   "FrmManRegAcc_IngNum.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   795
   ScaleWidth      =   2460
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   105
      TabIndex        =   0
      Top             =   15
      Width           =   2265
      Begin TDBDate6Ctl.TDBDate TDBFecDoc 
         Height          =   315
         Left            =   720
         TabIndex        =   3
         Top             =   225
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   556
         Calendar        =   "FrmManRegAcc_IngNum.frx":0ECA
         Caption         =   "FrmManRegAcc_IngNum.frx":0FE2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmManRegAcc_IngNum.frx":104E
         Keys            =   "FrmManRegAcc_IngNum.frx":106C
         Spin            =   "FrmManRegAcc_IngNum.frx":10CA
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "dd/mm/yyyy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   "dd/mm/yyyy"
         HighlightText   =   0
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxDate         =   2958465
         MinDate         =   -657434
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "06/12/2011"
         ValidateMode    =   0
         ValueVT         =   1766195207
         Value           =   40883
         CenturyMode     =   0
      End
      Begin VB.TextBox TxtNum 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   735
         TabIndex        =   2
         Text            =   "0.00"
         Top             =   225
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dato :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   255
         Width           =   450
      End
   End
End
Attribute VB_Name = "FrmManRegAcc_IngNum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub txtBMed_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 27 Then
'    Unload Me
'ElseIf KeyCode = 13 Then
'    FrmManRegAcc.VarBMed = txtBMed.Text
'    Unload Me
'End If
'End Sub
'
'Private Sub TDBFecDoc_Change()
'
'End Sub

Private Sub TDBFecDoc_GotFocus()
Dim lenf As Integer
lenf = Len(TDBFecDoc.Text)
TDBFecDoc.SelStart = 0
TDBFecDoc.SelLength = lenf
End Sub

Private Sub TDBFecDoc_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 27 Then
    Unload Me
ElseIf KeyCode = 13 Then
    FrmManRegAcc.VarBMed = TDBFecDoc.Value
    Unload Me
End If
End Sub

Private Sub TxtNum_GotFocus()
Dim lenN As Integer
lenN = Len(TxtNum.Text)
TxtNum.SelStart = 0
TxtNum.SelLength = lenN
End Sub

Private Sub TxtNum_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
    FrmManRegAcc.VarMonto = 0
ElseIf KeyCode = 13 Then
    FrmManRegAcc.VarMonto = Val(TxtNum.Text)
    Unload Me
ElseIf Not IsNumeric(Chr$(KeyCode)) Then
KeyCode = 0
End If

End Sub
