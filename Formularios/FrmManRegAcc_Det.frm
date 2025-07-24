VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Begin VB.Form FrmManRegAcc_Det 
   Caption         =   " ::: Seleccione la Información :::"
   ClientHeight    =   5625
   ClientLeft      =   4335
   ClientTop       =   4185
   ClientWidth     =   9750
   Icon            =   "FrmManRegAcc_Det.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5625
   ScaleWidth      =   9750
   Begin VB.Frame Frame1 
      Caption         =   "Seleccione la Cuenta a Buscar Información"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   135
      TabIndex        =   0
      Top             =   45
      Visible         =   0   'False
      Width           =   7425
      Begin VB.ComboBox CbCuenta 
         BackColor       =   &H00FCFCFC&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "FrmManRegAcc_Det.frx":0ECA
         Left            =   1020
         List            =   "FrmManRegAcc_Det.frx":0ECC
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   330
         Width           =   6270
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta :"
         Height          =   195
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   600
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MshMov 
      Height          =   4140
      Left            =   135
      TabIndex        =   3
      Top             =   960
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   7303
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
   End
   Begin MSForms.CommandButton CmdQuitarSel 
      Height          =   375
      Left            =   7725
      TabIndex        =   6
      ToolTipText     =   "Insertar el movimientos seleccionado"
      Top             =   510
      Width           =   1890
      Caption         =   "  Quitar Seleccion"
      PicturePosition =   327683
      Size            =   "3334;661"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CmdSelTodo 
      Height          =   375
      Left            =   7725
      TabIndex        =   5
      ToolTipText     =   "Insertar el movimientos seleccionado"
      Top             =   120
      Width           =   1890
      Caption         =   "Seleccionar Todo"
      PicturePosition =   327683
      Size            =   "3334;661"
      Picture         =   "FrmManRegAcc_Det.frx":0ECE
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdInsertarItem 
      Height          =   375
      Left            =   7845
      TabIndex        =   4
      ToolTipText     =   "Insertar el movimientos seleccionado"
      Top             =   5175
      Width           =   1845
      Caption         =   "  Insertar Selección"
      PicturePosition =   327683
      Size            =   "3254;661"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "FrmManRegAcc_Det"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VarPeriodo As String
Public sCuenta As String
Public bEntro As Boolean

Private Sub CbCuenta_Click()
If CbCuenta.ListIndex >= 0 Then
    Dim rs As New ADODB.Recordset, VarNumFila As Integer

    Set rs = Fct_Obtener_Lista_de_Movimiento(gsEmpresa, gsAnio, Left(FrmManRegAcc.CbCuenta.Text, 2), VarPeriodo, VarNumFila)

    If VarNumFila < 0 Then
        Set MshMov.DataSource = rs
    Else
        If MshMov.TextMatrix(1, 1) <> "" Then
            Call Sb_Limpiar_Grilla(MshMov)
        End If
    End If

    MshMov.ColWidth(0) = 150
    MshMov.ColWidth(1) = 1050
    MshMov.ColWidth(2) = 4500
    MshMov.ColWidth(3) = 1650
    MshMov.ColWidth(4) = 1325
    MshMov.ColWidth(5) = 450
    MshMov.ColWidth(6) = 0
    MshMov.ColWidth(7) = 0
    MshMov.ColWidth(8) = 0
    MshMov.ColWidth(9) = 0
    If MshMov.COLS > 2 Then
        On Error Resume Next
        MshMov.ColAlignment(4) = flexAlignRightTop
        MshMov.ColAlignment(5) = flexAlignCenterCenter
    End If
End If
End Sub

Private Sub cmdInsertarItem_Click()
Dim VarContador As Integer, x As Integer

If MshMov.TextMatrix(1, 1) = "" Then Exit Sub

With FrmManRegAcc.MshMontos '.MshEmp
    For VarContador = 1 To MshMov.Rows - 1
        
        If MshMov.TextMatrix(VarContador, 5) = "*" Then
'            x = x + 1
            x = x + FrmManRegAcc.MshMontos.Row
                '>
            If x = 1 Then
'                .AddItem ""
                FrmManRegAcc.MshMontos.AddItem ""
                x = x + 1
            End If
            
            If sCuenta = "30" And gsTipoPlan = "1" Then
                .TextMatrix(x, 1) = MshMov.TextMatrix(VarContador, 1) 'Voucher
                .TextMatrix(x, 2) = MshMov.TextMatrix(VarContador, 6) 'Tipo entidad
                .TextMatrix(x, 3) = MshMov.TextMatrix(VarContador, 7) 'Cod Entidad
                .TextMatrix(x, 4) = MshMov.TextMatrix(VarContador, 3) 'Doc. Identidad
                .TextMatrix(x, 5) = MshMov.TextMatrix(VarContador, 2) 'Razon Social
                .TextMatrix(x, 6) = MshMov.TextMatrix(VarContador, 10) 'Denominacion
'                .TextMatrix(x, 7) = 0
'                .TextMatrix(x, 8) = 0
'                .TextMatrix(x, 9) = 0
'                .TextMatrix(x, 5) = MshMov.TextMatrix(VarContador, 7)
                .TextMatrix(x, 13) = Format(MshMov.TextMatrix(VarContador, 9), "dd/mm/yyyy") 'Fecha
                .TextMatrix(x, 14) = sCuenta
                .TextMatrix(x, 15) = MshMov.TextMatrix(VarContador, 8)
            Else
                .TextMatrix(x, 1) = MshMov.TextMatrix(VarContador, 1) 'Voucher
                .TextMatrix(x, 2) = MshMov.TextMatrix(VarContador, 2) 'Razon Social
                .TextMatrix(x, 3) = MshMov.TextMatrix(VarContador, 3) 'Ruc
                .TextMatrix(x, 4) = MshMov.TextMatrix(VarContador, 6)
                .TextMatrix(x, 5) = MshMov.TextMatrix(VarContador, 7)
                .TextMatrix(x, 6) = MshMov.TextMatrix(VarContador, 8)
            End If
            
            With FrmManRegAcc.MshMontos
                Select Case sCuenta
                    Case "11" ', "30" And gsTipoPlan = "1"
                        .TextMatrix(x, 2) = "1.00"
                        .TextMatrix(x, 3) = LTrim(RTrim(Format(Val(MshMov.TextMatrix(VarContador, 4)), "### ##0.00")))
                        .TextMatrix(x, 4) = LTrim(RTrim(Format(Val(MshMov.TextMatrix(VarContador, 4)), "### ##0.00")))
                        .TextMatrix(x, 5) = "0.00"
                        .TextMatrix(x, 6) = "0.00"
                        .TextMatrix(x, 7) = "0.00"
                        .TextMatrix(x, 8) = "0.00"
                        .TextMatrix(x, 10) = sCuenta
                    Case "31" And gsTipoPlan = "0"
                        .TextMatrix(x, 1) = ""
                        .TextMatrix(x, 2) = "1.00"
                        .TextMatrix(x, 3) = "1.00"
                        .TextMatrix(x, 4) = LTrim(RTrim(Format(Val(MshMov.TextMatrix(VarContador, 4)), "### ##0.00")))
                        .TextMatrix(x, 5) = "0.00"
                        .TextMatrix(x, 6) = "0.00"
                    Case "31" And gsTipoPlan = "1"
                        .TextMatrix(x, 1) = ""
                        .TextMatrix(x, 2) = ""
                        .TextMatrix(x, 3) = ""
                        .TextMatrix(x, 4) = Format(MshMov.TextMatrix(VarContador, 9), "dd/mm/yyyy")
                        .TextMatrix(x, 5) = LTrim(RTrim(Format(Val(MshMov.TextMatrix(VarContador, 4)), "### ##0.00")))
                        .TextMatrix(x, 6) = "0.00"
                        .TextMatrix(x, 7) = "0.00"
                        .TextMatrix(x, 10) = MshMov.TextMatrix(VarContador, 1)
                End Select
            End With
        End If
    Next VarContador
    If sCuenta <> "30" And gsTipoPlan <> "1" Then Call FrmManRegAcc.Sb_Sumar(0)
End With
Unload Me
End Sub

Private Sub CmdQuitarSel_Click()
With MshMov
    If .TextMatrix(1, 1) = "" Then Exit Sub
    Dim VarCont As Integer
    On Error GoTo Control
    For VarCont = 1 To .Rows - 1
        .Row = VarCont
        .Col = 5
        If .TextMatrix(.Row, .Col) = "*" Then
            Call MshMov_Click
        End If
    Next VarCont
Control:
    .Tag = ""
End With
End Sub

Private Sub CmdSelTodo_Click()
With MshMov
    If .TextMatrix(1, 1) = "" Then Exit Sub
    Dim VarCont As Integer
    .Tag = "1"
    On Error GoTo Control
    For VarCont = 1 To .Rows - 1
        .Row = VarCont
        .Col = 5
        Call MshMov_Click
    Next VarCont
Control:
    .Tag = ""
End With
End Sub

Private Sub Form_Activate()
On Error GoTo MIERROR
Dim x As Integer

Dim rs As New ADODB.Recordset, VarNumFila As Integer
If Not bEntro Then
    Exit Sub
End If

Set rs = Fct_Obtener_Lista_de_Movimiento(gsEmpresa, gsAnio, sCuenta, VarPeriodo, VarNumFila)

With MshMov
    If VarNumFila < 0 Then
        Set .DataSource = rs
    Else
        If .TextMatrix(1, 1) <> "" Then
            Call Sb_Limpiar_Grilla(MshMov)
        End If
    End If
    .ColWidth(0) = 150
    .ColWidth(1) = 1050
    .ColWidth(2) = 4500
    .ColWidth(3) = 1650
    .ColWidth(4) = 1325
    .ColWidth(5) = 450
    .ColWidth(6) = 0
    .ColWidth(7) = 0
    .ColWidth(8) = 0
    .ColWidth(9) = 0
    If .COLS > 2 Then
        On Error Resume Next
        .ColAlignment(4) = flexAlignRightTop
        .ColAlignment(5) = flexAlignCenterCenter
    End If
End With

Me.Height = 6195
Me.Width = 9990
Call Centrar_form(Me)

If CbCuenta.ListCount > 0 Then CbCuenta.ListIndex = 0
bEntro = False
Exit Sub
MIERROR:

End Sub

Private Sub Form_Load()
'sCuenta = Left(FrmManRegAcc.CbCuenta.Text, 2)
End Sub

Private Sub MshMov_Click()
With MshMov
    If .Col = 5 Then
        If .Tag = "1" Then GoTo ir
        If LTrim(RTrim(.TextMatrix(.Row, .Col))) <> "" Then
            Dim i As Byte
            .TextMatrix(.Row, .Col) = ""
            For i = 1 To .COLS - 1
                .Col = i
                .CellBackColor = &H80FFFF
            Next i
        Else
ir:
            .TextMatrix(.Row, .Col) = "*"
            For i = 1 To .COLS - 1
                .Col = i
                .CellBackColor = &HFFFF80
            Next i
        End If
    End If
End With
End Sub

Private Sub MshMov_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 32 Or KeyCode = 13 Then Call MshMov_Click
End Sub
'Private Sub Form_Load()
'
'End Sub

