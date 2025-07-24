VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmPrcCentroCosto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualizar Estructura Centro de Costo"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7665
   Icon            =   "frmPrcCentroCosto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   7665
   Begin VB.Frame Frame1 
      Height          =   1230
      Index           =   0
      Left            =   45
      TabIndex        =   1
      Top             =   990
      Width           =   7560
      Begin MSComctlLib.ProgressBar pgbAvance 
         Height          =   195
         Left            =   90
         TabIndex        =   2
         Top             =   180
         Width           =   7395
         _ExtentX        =   13044
         _ExtentY        =   344
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin MSForms.CommandButton cmdProcesar 
         Height          =   435
         Left            =   2745
         TabIndex        =   3
         Top             =   540
         Width           =   1665
         Caption         =   "Procesar"
         PicturePosition =   327683
         Size            =   "2937;767"
         Picture         =   "frmPrcCentroCosto.frx":0ECA
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   $"frmPrcCentroCosto.frx":285C
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   825
      Index           =   1
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   7545
   End
End
Attribute VB_Name = "frmPrcCentroCosto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gsGrupo As String
Dim i As Integer, Existe As Integer
Dim lArrMnt(3) As Variant
Dim RstVerMig As ADODB.Recordset

Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

Private Sub cmdProcesar_Click()
On Error GoTo Control
 Procesar
 
Exit Sub
Control:
 MsgBox Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo Control
ConectarAdvance
 
Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))
Call Centrar_form(Me)
 
Desconectar

Set RstVerMig = Nothing
 
Exit Sub
Control:
 Set RstVerMig = Nothing
 Desconectar
 MsgBox Err.Description
End Sub
Private Sub Form_Unload(Cancel As Integer)
 Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))
End Sub
Sub Procesar()
On Error GoTo Control
'Dim Res As Integer
' If MsgBox("Antes de realizar esta operación es recomendable realizar un BAKUP DE LA BASE DE DATOS...", 4 + 32, App.Title) = 6 Then
'  cmdProcesar.Enabled = False
'  Exit Sub
' Else
'  cmdProcesar.Enabled = True
' End If
 
Screen.MousePointer = vbHourglass
 Set RstVerMig = New ADODB.Recordset
 
 RstVerMig.Open "spCNT_VerificaCentCosto '" & gsEmpresa & "','" & gsAnio & "'", gcnSistemaAdv, adOpenDynamic, adLockOptimistic
 If RstVerMig.State > 0 Then Existe = RstVerMig.Fields(0)
 If Existe <> 0 Then
  MsgBox "La Estructura actual del Centro de Costo esta Actualizada o no es necesario realizar este Proceso...", vbInformation, App.Title
  Screen.MousePointer = vbDefault
  Exit Sub
 End If
        
 For i = 1 To 1000
  pgbAvance.Min = 0
  pgbAvance.Max = 1000
  
  If i = 500 Then
    gcnSistemaAdv.Execute "spCNT_MigraCentCosto '" & gsEmpresa & "','" & gsAnio & "'"
  End If
  pgbAvance.Value = i
  DoEvents
 Next i
  
Screen.MousePointer = vbDefault
pgbAvance.Value = 0
MsgBox "Proceso Concluido satisfactoriamente...", vbInformation, App.Title

Exit Sub
Control:
 MsgBox Err.Description
 Screen.MousePointer = vbDefault
End Sub
