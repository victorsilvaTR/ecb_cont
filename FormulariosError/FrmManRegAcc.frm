VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Begin VB.Form FrmManRegAcc 
   Caption         =   " ::: Registro de Inversiones :::"
   ClientHeight    =   4245
   ClientLeft      =   3165
   ClientTop       =   4575
   ClientWidth     =   13065
   Icon            =   "FrmManRegAcc.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4245
   ScaleWidth      =   13065
   Begin VB.Frame Frame2 
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
      Height          =   795
      Left            =   3630
      TabIndex        =   11
      Top             =   45
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
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   330
         Width           =   6270
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta :"
         Height          =   195
         Left            =   360
         TabIndex        =   13
         Top             =   360
         Width           =   600
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccione"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   135
      TabIndex        =   0
      Top             =   45
      Width           =   3465
      Begin VB.ComboBox CbPeriodo 
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
         Left            =   1125
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   330
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Periodo :"
         Height          =   195
         Left            =   435
         TabIndex        =   2
         Top             =   360
         Width           =   630
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MshMontos 
      Height          =   2790
      Left            =   1545
      TabIndex        =   3
      Top             =   840
      Width           =   11430
      _ExtentX        =   20161
      _ExtentY        =   4921
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
   Begin VB.Label LblCuenta 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3810
      TabIndex        =   10
      Top             =   330
      Width           =   75
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   5430
      X2              =   5445
      Y1              =   855
      Y2              =   3650
   End
   Begin MSForms.CommandButton cmdEliminarTodo 
      Height          =   375
      Left            =   90
      TabIndex        =   9
      ToolTipText     =   "Eliminar todos los movimientos del libro y mes seleccionado"
      Top             =   2790
      Width           =   1380
      Caption         =   " Eliminar Todo"
      PicturePosition =   327683
      Size            =   "2434;661"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdEliminaItem 
      Height          =   375
      Left            =   90
      TabIndex        =   8
      ToolTipText     =   "Eliminar el movimientos seleccionado"
      Top             =   2325
      Width           =   1380
      Caption         =   " Eliminar Item"
      PicturePosition =   327683
      Size            =   "2434;661"
      Picture         =   "FrmManRegAcc.frx":0ECA
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdInsertarItem 
      Height          =   375
      Left            =   105
      TabIndex        =   7
      ToolTipText     =   "Insertar el movimientos seleccionado"
      Top             =   1860
      Width           =   1380
      Caption         =   " Insertar Mov."
      PicturePosition =   327683
      Size            =   "2434;661"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdGrabar 
      Height          =   375
      Left            =   90
      TabIndex        =   6
      ToolTipText     =   "Grabar modificaciones"
      Top             =   1395
      Width           =   1380
      Caption         =   " Grabar"
      PicturePosition =   327683
      Size            =   "2434;661"
      Picture         =   "FrmManRegAcc.frx":1464
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdRefresh 
      Height          =   375
      Left            =   90
      TabIndex        =   5
      ToolTipText     =   " Vuelve a cargar los datos almacenados "
      Top             =   915
      Width           =   1380
      Caption         =   " Listar"
      PicturePosition =   327683
      Size            =   "2434;661"
      Picture         =   "FrmManRegAcc.frx":19FE
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Height          =   390
      Left            =   11625
      TabIndex        =   4
      Top             =   3720
      Width           =   1350
      Caption         =   "   Salir"
      PicturePosition =   327683
      Size            =   "2381;688"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "FrmManRegAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VarMonto As Double
Public VarBMed As String
Dim gsGrupo As String

Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

Public Sub Sb_Formato()
'    With MshEmp
'        .FormatString = "|Voucher|Apellidos y Nombres, Razon Social|R.U.C.|||"
'        .COLS = 7
'        .ColWidth(0) = 150
'        .ColWidth(1) = 1000
'        .ColWidth(2) = 2000
'        .ColWidth(3) = 1700
'        .ColWidth(4) = 0
'        .ColWidth(5) = 0
'        .ColWidth(6) = 0
'
'        .ColAlignment(3) = flexAlignLeftBottom
'    End With
'    Call Sb_Formatear_Grilla(MshEmp)
    With MshMontos
        If sCuenta = "11" Then
            .FormatString = "|Apellidos y Nombres, Razon Social|R.U.C.|Denominación|Valor Nominal Unitario|Cantidad|Costo Total|Provisión Total|Otros Costos|Total Neto|CodEntidad|Voucher|TipoEntidad|Item"
        ElseIf sCuenta = "30" Then
'            .FormatString = "|Apellidos y Nombres, Razon Social|R.U.C.|Denominación|Valor Nominal Unitario|Cantidad|Costo Total|Provisión Total|Ajuste Part. Patrimonial|Total Neto||||"
            .FormatString = "|Nº. Voucher|Tipo Entidad|Cod Entidad|Doc. Ident.|Ape. y Nom, Denominacion o Razon Social|Denominacion|V.N.U|Cantidad|Costo Total|Prov. Total|Ajuste Part.|Total|Fecha|Cuenta|Item"
        ElseIf sCuenta = "31" And gsTipoPlan = "0" Then
            .FormatString = "|Apellidos y Nombres, Razon Social|R.U.C.|Denominación|Valor Nominal Unitario|Cantidad|Costo Total|Provisión Total|Total Neto|||||||||||||||||"
        ElseIf sCuenta = "31" And gsTipoPlan = "1" Then
            .FormatString = "|Cuenta|Tipo de Inversion|Base de Medicion|Fecha de Adquisición|Costo de Adquisicion|Depreciacion Acumulada|Valor Neto|Valor razonable al final del ejercicio|Variacion en los resultados|||||||||"
        End If
        .COLS = 16
        Select Case sCuenta
            Case "11" And gsTipoPlan = "1"
                .ColWidth(0) = 150
                .ColWidth(1) = 2600
                .ColWidth(2) = 1700
                .ColWidth(3) = 2500
                .ColWidth(4) = 950
                .ColWidth(5) = 1050
                .ColWidth(6) = 1050
                .ColWidth(7) = 1050
                .ColWidth(8) = 1050
                .ColWidth(9) = 1050
                .ColWidth(10) = 0
                .ColWidth(11) = 0
                .ColWidth(12) = 0
                .ColWidth(13) = 900
                .ColWidth(14) = 900
                .ColWidth(15) = 900
                .ColWidth(16) = 900
            Case "30" And gsTipoPlan = "1"
                .ColWidth(0) = 150
                .ColWidth(1) = 0 'Voucher
                .ColWidth(2) = 0  'Tipo Entidad
                .ColWidth(3) = 0  'Cod Entidad
                .ColWidth(4) = 1000 'Doc. Ident.
                .ColWidth(5) = 2500 'Razon Social
                .ColWidth(6) = 1700 'Denominacion
                .ColWidth(7) = 1000 'V.N.U
                .ColWidth(8) = 1000 'Cantidad
                .ColWidth(9) = 1000 'Costo
                .ColWidth(10) = 1000 'Provision
                .ColWidth(11) = 1000 'Ajuste
                .ColWidth(12) = 1000 'Total
                .ColWidth(13) = 0 'Fecha
                .ColWidth(14) = 0 'Cuenta
                .ColWidth(15) = 0 'Item
            Case "31" And gsTipoPlan = "1"
                .ColWidth(0) = 150
                .ColWidth(1) = 1050
                .ColWidth(2) = 1550
                .ColWidth(3) = 1550
                .ColWidth(4) = 1550
                .ColWidth(5) = 1550
                .ColWidth(6) = 1550
                .ColWidth(7) = 1550
                .ColWidth(8) = 1550
                .ColWidth(9) = 1550
                .ColWidth(10) = 0
                .ColWidth(11) = 0
                .ColWidth(12) = 0
                .ColWidth(13) = 0
                .ColWidth(15) = 0
                .ColWidth(15) = 0
                '.ColWidth(16) = 0
            Case "31" And gsTipoPlan = "0"
                .ColWidth(0) = 150
                .ColWidth(1) = 2600
                .ColWidth(2) = 1700
                .ColWidth(3) = 2500
                .ColWidth(4) = 950
                .ColWidth(5) = 1050
                .ColWidth(6) = 1050
                .ColWidth(7) = 1050
                .ColWidth(8) = 1050
                .ColWidth(9) = 0
                .ColWidth(10) = 0
                .ColWidth(11) = 0
                .ColWidth(12) = 0
                .ColWidth(13) = 0
                .ColWidth(14) = 0
                .ColWidth(15) = 0
        End Select
    End With
    Call Sb_Formatear_Grilla(MshMontos)
End Sub

Public Sub Sb_Sumar(ByRef Par_Fila As Integer)
    Dim i As Integer
    With MshMontos
        For i = IIf(Par_Fila = 0, 1, Par_Fila) To IIf(Par_Fila = 0, MshMontos.Rows - 1, Par_Fila)
            Select Case sCuenta
                Case "11" ', "30"
'                    If i = 1 Then .TextMatrix(i, 8) = "0"
'                    .TextMatrix(i, 4) = Format(Val(.TextMatrix(i, 2)) * Val(.TextMatrix(i, 4)), "0.00")
                    .TextMatrix(i, 9) = Format((Val(.TextMatrix(i, 6)) + Val(.TextMatrix(i, 8)) - Val(.TextMatrix(i, 7))), "0.00")
'                    Val(.TextMatrix(i, 7))) - Val(.TextMatrix(i, 5))), "0.00")
                Case "30" And gsTipoPlan = "1"
                    
                    If .TextMatrix(i, 7) <> vbNullString And .TextMatrix(i, 8) <> vbNullString Then
                        .TextMatrix(i, 9) = Format(Val(.TextMatrix(i, 7)) * Val(.TextMatrix(i, 8)), "0.00")
                        .TextMatrix(i, 12) = Format(Val(.TextMatrix(i, 9)), "0.00")
                    End If
                    
                    If .TextMatrix(i, 11) <> vbNullString And .TextMatrix(i, 9) <> vbNullString Then
                        .TextMatrix(i, 12) = Format(Val(.TextMatrix(i, 9)) + Val(.TextMatrix(i, 11)), "0.00")
                    Else
                        .TextMatrix(i, 12) = Format(Val(.TextMatrix(i, 9)), "0.00")
                    End If
                    
                Case "31" And gsTipoPlan = "0"
                    .TextMatrix(i, 8) = Format(Val(.TextMatrix(i, 6)) - Val(.TextMatrix(i, 7)), "0.00")
                Case "31" And gsTipoPlan = "1"
                    .TextMatrix(i, 7) = Format(Val(.TextMatrix(i, 5)) - Val(.TextMatrix(i, 6)), "0.00")
                    .TextMatrix(i, 9) = Format(Val(.TextMatrix(i, 8)) - Val(.TextMatrix(i, 7)), "0.00")
            End Select
        Next i
    End With
End Sub

Private Sub CbCuenta_Click()
    sCuenta = Left(CbCuenta.Text, 2)
    CbPeriodo_Click
End Sub

Private Sub CbPeriodo_Click()
Dim rs As New ADODB.Recordset, VarNumFila As Integer
MshMontos.Rows = 2
Call Sb_Formato
'Set rs = Fct_Listar_C_Valores_Detalle(gsEmpresa, gsAnio, Format(CbPeriodo.ListIndex, "00"), VarNumFila, "C", Me.CbCuenta.Text)
'MshMontos.Rows = 2
'MshEmp.Rows = 2
'If VarNumFila < 0 Then
'    Set MshEmp.DataSource = rs
'Else
'    Call Sb_Limpiar_Grilla(MshEmp)
'End If
'Set rs = Nothing
'Set rs = New ADODB.Recordset
'
Set rs = Fct_Listar_C_Valores_Detalle(gsEmpresa, gsAnio, Format(CbPeriodo.ListIndex, "00"), VarNumFila, "D", Me.CbCuenta.Text)

If VarNumFila < 0 Then
Dim x As Integer
x = 0

Do While Not rs.EOF
    x = x + 1
    If x <> 1 Then MshMontos.AddItem ""

If sCuenta = "31" And gsTipoPlan = "1" Then
    MshMontos.TextMatrix(x, 1) = rs(19)
    MshMontos.TextMatrix(x, 2) = rs(2)
    MshMontos.TextMatrix(x, 3) = rs(5)
    MshMontos.TextMatrix(x, 4) = rs(18)
    MshMontos.TextMatrix(x, 5) = rs(6)
    MshMontos.TextMatrix(x, 6) = rs(7)
    MshMontos.TextMatrix(x, 7) = rs(8)
    MshMontos.TextMatrix(x, 15) = rs(9)
    MshMontos.TextMatrix(x, 14) = rs(10)
    MshMontos.TextMatrix(x, 13) = rs(17)
ElseIf sCuenta = "31" And gsTipoPlan = "0" Then
    MshMontos.TextMatrix(x, 1) = rs(0)
    MshMontos.TextMatrix(x, 2) = rs(1)
    MshMontos.TextMatrix(x, 3) = rs(2)
    MshMontos.TextMatrix(x, 4) = rs(3)
    MshMontos.TextMatrix(x, 5) = rs(4)
    MshMontos.TextMatrix(x, 6) = rs(6)
    MshMontos.TextMatrix(x, 7) = rs(7)
    MshMontos.TextMatrix(x, 8) = rs(8)
    MshMontos.TextMatrix(x, 15) = rs(9)
    MshMontos.TextMatrix(x, 14) = rs(10)
    MshMontos.TextMatrix(x, 11) = rs(11)
    MshMontos.TextMatrix(x, 12) = rs(12)
ElseIf sCuenta = "30" And gsTipoPlan = "1" Then
    MshMontos.TextMatrix(x, 5) = rs(0) 'Razon Social
    MshMontos.TextMatrix(x, 4) = rs(1) 'Doc. Identidad
    MshMontos.TextMatrix(x, 7) = rs(3) 'V.N.U.
    MshMontos.TextMatrix(x, 8) = rs!Val_nCantidad 'Cantidad
    MshMontos.TextMatrix(x, 9) = rs(6)
    MshMontos.TextMatrix(x, 10) = rs(7)
    MshMontos.TextMatrix(x, 12) = rs(8)
    MshMontos.TextMatrix(x, 15) = rs(10)
    MshMontos.TextMatrix(x, 3) = rs(12)
    MshMontos.TextMatrix(x, 2) = rs(11)
    MshMontos.TextMatrix(x, 13) = rs(18)
    MshMontos.TextMatrix(x, 14) = rs(19)
    MshMontos.TextMatrix(x, 11) = rs(21)
    MshMontos.TextMatrix(x, 6) = rs(22) 'Denominacion
Else
    MshMontos.TextMatrix(x, 1) = IIf(IsNull(rs(0)), vbNullString, rs(0))
    MshMontos.TextMatrix(x, 2) = IIf(IsNull(rs(1)), vbNullString, rs(1))
    MshMontos.TextMatrix(x, 3) = rs(2)
    MshMontos.TextMatrix(x, 4) = rs(3)
    MshMontos.TextMatrix(x, 5) = rs(4)
    MshMontos.TextMatrix(x, 6) = rs(6)
    MshMontos.TextMatrix(x, 7) = rs(7)
    MshMontos.TextMatrix(x, 8) = rs(20)
    MshMontos.TextMatrix(x, 9) = rs(8)
    MshMontos.TextMatrix(x, 15) = rs(9)
    MshMontos.TextMatrix(x, 14) = rs(10)
    MshMontos.TextMatrix(x, 11) = rs(11)
    MshMontos.TextMatrix(x, 12) = rs(12)
'    MshMontos.TextMatrix(x, 12) = rs(12)
End If
rs.MoveNext
Loop
Else
    Call Sb_Limpiar_Grilla(MshMontos)
End If
Call Sb_Formato
MshMontos.Col = 1
End Sub

Private Sub cmdEliminaItem_Click()

If MshMontos.TextMatrix(MshMontos.Row, 1) = "" Then Exit Sub
If MsgBox("Desea eliminar item seleccionado", vbYesNo + vbInformation, "ECB-CONT") = vbYes Then


Dim iColumnTitulo As Integer
Dim icolumnBaseMedicion As Integer
Dim ContRow As Integer
Dim VarContador As Integer
Select Case Left(Me.CbCuenta.Text, 2)
    Case "11", "30"
        iColumnTitulo = 15
        icolumnBaseMedicion = 13
    Case "31" And gsTipoPlan = "0"
        iColumnTitulo = 15
        icolumnBaseMedicion = 11
    Case "31" And gsTipoPlan = "1"
        iColumnTitulo = 15
        icolumnBaseMedicion = 13
End Select
VarContador = 0
If sCuenta = "30" And gsTipoPlan = "1" Then
'    For VarContador = 1 To Me.MshMontos.Rows - 1
'        Call Sb_Borrar_Valores_Detalle("ELIMINAR_TODO", gsEmpresa, gsAnio, Format(Me.CbPeriodo.ListIndex, "00"), Me.MshMontos.TextMatrix(VarContador, 1), _
'                                       Me.MshMontos.TextMatrix(VarContador, 15), Me.MshMontos.TextMatrix(VarContador, 2), Me.MshMontos.TextMatrix(VarContador, 3), _
'                                       "", "", Me.MshMontos.TextMatrix(VarContador, 14))
'    Next VarContador
    Call Sb_Borrar_Valores_Detalle("ELIMINAR_TODO", gsEmpresa, gsAnio, Format(Me.CbPeriodo.ListIndex, "00"), Me.MshMontos.TextMatrix(VarContador, 1), _
                                       Me.MshMontos.TextMatrix(VarContador, 15), Me.MshMontos.TextMatrix(VarContador, 2), Me.MshMontos.TextMatrix(VarContador, 3), _
                                       "", "", Me.MshMontos.TextMatrix(VarContador, 14))
Else
    Call Sb_Borrar_Valores_Detalle("ELIMINAR_ITEM", gsEmpresa, gsAnio, Format(CbPeriodo.ListIndex, "00"), _
    MshMontos.TextMatrix(MshMontos.Row, 0), MshMontos.TextMatrix(MshMontos.Row, 14), MshMontos.TextMatrix(MshMontos.Row, 12), _
    MshMontos.TextMatrix(MshMontos.Row, 11), MshMontos.TextMatrix(MshMontos.Row, iColumnTitulo), _
    IIf(sCuenta = "31" And gsTipoPlan = "0", "", MshMontos.TextMatrix(MshMontos.Row, icolumnBaseMedicion)), sCuenta)
End If
If MshMontos.Rows = 2 Then
    Call Sb_Limpiar_Grilla(MshMontos)
Else
    CbPeriodo_Click
End If
VarContador = 0
 MsgBox "El item selecionado se elimino correctamente", vbInformation + vbSystemModal, "ECB-CONT"
Else
Exit Sub
End If
End Sub

Private Sub cmdEliminarTodo_Click()

If MshMontos.TextMatrix(1, 1) = "" Then Exit Sub
If MsgBox("Desea eliminar todos los items", vbYesNo + vbInformation, "ECB-CONT") = vbYes Then


Dim VarContador As Integer
Dim iColumnTitulo As Integer
Dim icolumnBaseMedicion As Integer


Select Case Left(Me.CbCuenta.Text, 2)
    Case "11", "30"
        iColumnTitulo = 15
        icolumnBaseMedicion = 11
    Case "31"
        iColumnTitulo = 15
        icolumnBaseMedicion = 11
End Select

If sCuenta = "30" And gsTipoPlan = "1" Then
    For VarContador = 1 To Me.MshMontos.Rows - 1
'        Call Sb_Borrar_Valores_Detalle("ELIMINAR_TODO", gsEmpresa, gsAnio, Format(Me.CbPeriodo.ListIndex, "00"), Me.MshMontos.TextMatrix(VarContador, 1), _
'                                       Me.MshMontos.TextMatrix(VarContador, 15), Me.MshMontos.TextMatrix(VarContador, 2), Me.MshMontos.TextMatrix(VarContador, 3), _
'                                       "", "", Me.MshMontos.TextMatrix(VarContador, 14))
        Call Sb_Borrar_Valores_Detalle("ELIMINAR_TODO", gsEmpresa, gsAnio, Format(Me.CbPeriodo.ListIndex, "00"), Me.MshMontos.TextMatrix(VarContador, 1), _
                                       Me.MshMontos.TextMatrix(VarContador, 15), Me.MshMontos.TextMatrix(VarContador, 2), Me.MshMontos.TextMatrix(VarContador, 3), _
                                       "", "", Me.MshMontos.TextMatrix(VarContador, 14))
    Next VarContador
Else
    For VarContador = 1 To MshMontos.Rows - 1
        Call Sb_Borrar_Valores_Detalle("ELIMINAR_TODO", gsEmpresa, gsAnio, Format(CbPeriodo.ListIndex, "00"), _
        MshMontos.TextMatrix(VarContador, 0), MshMontos.TextMatrix(VarContador, 14), MshMontos.TextMatrix(VarContador, 12), _
        MshMontos.TextMatrix(VarContador, 11), MshMontos.TextMatrix(VarContador, iColumnTitulo), IIf(sCuenta = "31" And gsTipoPlan = "0", "", MshMontos.TextMatrix(VarContador, icolumnBaseMedicion)), sCuenta)
    Next VarContador
End If
'Call Sb_Limpiar_Grilla(MshEmp)
MshMontos.Rows = 2
Call Sb_Limpiar_Grilla(MshMontos)
MsgBox "Se eliminaron todos los items", vbInformation + vbSystemModal, "ECB-CONT"
Else

Exit Sub
End If
End Sub

Private Sub cmdGrabar_Click()
Dim VarI As Integer
Dim bGrabo As Boolean
Dim CampTemp As String
'If MshMontos.TextMatrix(1, 1) = "" Then Exit Sub
If MsgBox("Desea grabar el registro", vbYesNo + vbInformation, "ECB-CONT") = vbYes Then


bGrabo = True
Dim VarContador As Integer

If sCuenta = "30" And gsTipoPlan = "1" Then
    For VarContador = 1 To Me.MshMontos.Rows - 1
        If MshMontos.TextMatrix(VarContador, 0) <> vbNullString Then
            Call Sb_Borrar_Valores_Detalle("ELIMINAR_TODO", gsEmpresa, gsAnio, Format(Me.CbPeriodo.ListIndex, "00"), Me.MshMontos.TextMatrix(VarContador, 1), _
                                           Me.MshMontos.TextMatrix(VarContador, 15), Me.MshMontos.TextMatrix(VarContador, 2), Me.MshMontos.TextMatrix(VarContador, 3), _
                                           "", "", Me.MshMontos.TextMatrix(VarContador, 14))
        End If
    Next VarContador
Else
    For VarContador = 1 To MshMontos.Rows - 1
        If MshMontos.TextMatrix(VarContador, 0) <> vbNullString Then
            Call Sb_Borrar_Valores_Detalle("ELIMINAR_TODO", gsEmpresa, gsAnio, Format(CbPeriodo.ListIndex, "00"), _
            MshMontos.TextMatrix(VarContador, 0), MshMontos.TextMatrix(VarContador, 14), MshMontos.TextMatrix(VarContador, 12), _
            MshMontos.TextMatrix(VarContador, 11), MshMontos.TextMatrix(VarContador, 15), IIf(sCuenta = "31" And gsTipoPlan = "0", "", MshMontos.TextMatrix(VarContador, icolumnBaseMedicion)), sCuenta)
        End If
    Next VarContador
End If


If sCuenta = "31" And gsTipoPlan = "1" Then
For VarI = 1 To MshMontos.Rows - 1
    If MshMontos.TextMatrix(VarI, 4) <> vbNullString Then
        If Not Sb_Grabar_Valores_Detalle(gsEmpresa, gsAnio, Format(CbPeriodo.ListIndex, "00"), _
            MshMontos.TextMatrix(VarI, 10), Val(MshMontos.TextMatrix(VarI, 14)), MshMontos.TextMatrix(VarI, 12), _
            IIf(sCuenta = "31" And gsTipoPlan = "1", "", MshMontos.TextMatrix(VarI, 11)), _
            IIf(sCuenta < "31", MshMontos.TextMatrix(VarI, 15), Trim$(MshMontos.TextMatrix(VarI, 15))), _
            IIf(sCuenta = "31" And gsTipoPlan = "1", MshMontos.TextMatrix(VarI, 2), MshMontos.TextMatrix(VarI, 3)), _
           IIf(sCuenta < "31", "", IIf(sCuenta = "31" And gsTipoPlan = "1", MshMontos.TextMatrix(VarI, 13), "")), _
            IIf(sCuenta < "31", "", IIf(sCuenta = "31" And gsTipoPlan = "1", MshMontos.TextMatrix(VarI, 3), "")), _
            IIf(sCuenta < "31", Val(MshMontos.TextMatrix(VarI, 4)), IIf(sCuenta = "31" And gsTipoPlan = "0", Val(MshMontos.TextMatrix(VarI, 4)), "0")), _
            IIf(sCuenta < "31", Val(MshMontos.TextMatrix(VarI, 5)), IIf(sCuenta = "31" And gsTipoPlan = "0", Val(MshMontos.TextMatrix(VarI, 5)), "0")), _
            IIf(sCuenta < "31", Val(MshMontos.TextMatrix(VarI, 6)), IIf(sCuenta = "31" And gsTipoPlan = "0", Val(MshMontos.TextMatrix(VarI, 6)), Val(MshMontos.TextMatrix(VarI, 5)))), _
            IIf(sCuenta < "31", Val(MshMontos.TextMatrix(VarI, 7)), IIf(sCuenta = "31" And gsTipoPlan = "0", Val(MshMontos.TextMatrix(VarI, 7)), Val(MshMontos.TextMatrix(VarI, 6)))), _
            IIf(sCuenta < "31", Val(MshMontos.TextMatrix(VarI, 9)), IIf(sCuenta = "31" And gsTipoPlan = "0", Val(MshMontos.TextMatrix(VarI, 8)), Val(MshMontos.TextMatrix(VarI, 7)))), _
            IIf(sCuenta < "31", 0, Val(MshMontos.TextMatrix(VarI, 8))), _
            IIf(sCuenta < "31", Val(MshMontos.TextMatrix(VarI, 8)), 0), "I", _
            IIf(sCuenta = "31" And gsTipoPlan = "1", MshMontos.TextMatrix(VarI, 1), sCuenta), MshMontos.TextMatrix(VarI, 4)) Then
                bGrabo = False
                Exit For
        End If
    End If
Next VarI
Else
    If sCuenta = "30" And gsTipoPlan = "1" Then
        For VarI = 1 To Me.MshMontos.Rows - 1
            If MshMontos.TextMatrix(VarI, 4) <> vbNullString Then
                If Not Sb_Grabar_Valores_Detalle(gsEmpresa, _
                                                 gsAnio, _
                                                 Format(Me.CbPeriodo.ListIndex, "00"), _
                                                 Me.MshMontos.TextMatrix(VarI, 1), _
                                                 Me.MshMontos.TextMatrix(VarI, 15), _
                                                 Me.MshMontos.TextMatrix(VarI, 2), _
                                                 Me.MshMontos.TextMatrix(VarI, 3), _
                                                 "", _
                                                 "", _
                                                 "", _
                                                 "", _
                                                 IIf(Me.MshMontos.TextMatrix(VarI, 7) = vbNullString, 0, Val(Me.MshMontos.TextMatrix(VarI, 7))), _
                                                 IIf(Me.MshMontos.TextMatrix(VarI, 8) = vbNullString, 0, Val(Me.MshMontos.TextMatrix(VarI, 8))), _
                                                 IIf(Me.MshMontos.TextMatrix(VarI, 9) = vbNullString, 0, Val(Me.MshMontos.TextMatrix(VarI, 9))), _
                                                 IIf(Me.MshMontos.TextMatrix(VarI, 10) = vbNullString, 0, Val(Me.MshMontos.TextMatrix(VarI, 10))), _
                                                 IIf(Me.MshMontos.TextMatrix(VarI, 12) = vbNullString, 0, Val(Me.MshMontos.TextMatrix(VarI, 12))), _
                                                 IIf(Me.MshMontos.TextMatrix(VarI, 11) = vbNullString, 0, Val(Me.MshMontos.TextMatrix(VarI, 11))), _
                                                 0, "I", Me.MshMontos.TextMatrix(VarI, 14), Me.MshMontos.TextMatrix(VarI, 13)) Then
                    bGrabo = False
                    Exit For
                End If
            End If
        Next VarI
    Else
        If sCuenta <> "11" Then
           For VarI = 1 To MshMontos.Rows - 1
            If MshMontos.TextMatrix(VarI, 4) <> vbNullString Then
                If Not Sb_Grabar_Valores_Detalle(gsEmpresa, gsAnio, Format(CbPeriodo.ListIndex, "00"), _
                    MshMontos.TextMatrix(VarI, 0), Val(MshMontos.TextMatrix(VarI, 14)), MshMontos.TextMatrix(VarI, 12), _
                    IIf(sCuenta = "31" And gsTipoPlan = "1", "", MshMontos.TextMatrix(VarI, 11)), _
                    IIf(sCuenta < "31", MshMontos.TextMatrix(VarI, 15), Trim$(MshMontos.TextMatrix(VarI, 15))), _
                    IIf(sCuenta = "31" And gsTipoPlan = "1", MshMontos.TextMatrix(VarI, 2), MshMontos.TextMatrix(VarI, 3)), _
                   IIf(sCuenta < "31", "", IIf(sCuenta = "31" And gsTipoPlan = "1", MshMontos.TextMatrix(VarI, 13), "")), _
                    IIf(sCuenta < "31", "", IIf(sCuenta = "31" And gsTipoPlan = "1", MshMontos.TextMatrix(VarI, 3), "")), _
                    IIf(sCuenta < "31", Val(MshMontos.TextMatrix(VarI, 4)), IIf(sCuenta = "31" And gsTipoPlan = "0", Val(MshMontos.TextMatrix(VarI, 4)), "0")), _
                    IIf(sCuenta < "31", Val(MshMontos.TextMatrix(VarI, 5)), IIf(sCuenta = "31" And gsTipoPlan = "0", Val(MshMontos.TextMatrix(VarI, 5)), "0")), _
                    IIf(sCuenta < "31", Val(MshMontos.TextMatrix(VarI, 6)), IIf(sCuenta = "31" And gsTipoPlan = "0", Val(MshMontos.TextMatrix(VarI, 6)), Val(MshMontos.TextMatrix(VarI, 5)))), _
                    IIf(sCuenta < "31", Val(MshMontos.TextMatrix(VarI, 7)), IIf(sCuenta = "31" And gsTipoPlan = "0", Val(MshMontos.TextMatrix(VarI, 7)), Val(MshMontos.TextMatrix(VarI, 6)))), _
                    IIf(sCuenta < "31", Val(MshMontos.TextMatrix(VarI, 9)), IIf(sCuenta = "31" And gsTipoPlan = "0", Val(MshMontos.TextMatrix(VarI, 8)), Val(MshMontos.TextMatrix(VarI, 7)))), _
                    IIf(sCuenta < "31", 0, 0), _
                    IIf(sCuenta < "31", Val(MshMontos.TextMatrix(VarI, 8)), 0), "I", _
                    IIf(sCuenta = "31" And gsTipoPlan = "1", MshMontos.TextMatrix(VarI, 1), sCuenta)) Then
                        bGrabo = False
                        Exit For
                End If
            End If
        Next VarI
      Else
        For VarI = 1 To MshMontos.Rows - 1
            If MshMontos.TextMatrix(VarI, 4) <> vbNullString Then
                If Not Sb_Grabar_Valores_Detalle(gsEmpresa, gsAnio, Format(CbPeriodo.ListIndex, "00"), _
                    MshMontos.TextMatrix(VarI, 14), Val(MshMontos.TextMatrix(VarI, 16)), MshMontos.TextMatrix(VarI, 15), _
                    MshMontos.TextMatrix(VarI, 13), _
                    IIf(sCuenta < "31", MshMontos.TextMatrix(VarI, 15), Trim$(MshMontos.TextMatrix(VarI, 15))), _
                    IIf(sCuenta = "31" And gsTipoPlan = "1", MshMontos.TextMatrix(VarI, 2), MshMontos.TextMatrix(VarI, 3)), _
                   IIf(sCuenta < "31", "", IIf(sCuenta = "31" And gsTipoPlan = "1", MshMontos.TextMatrix(VarI, 13), "")), _
                    IIf(sCuenta < "31", "", IIf(sCuenta = "31" And gsTipoPlan = "1", MshMontos.TextMatrix(VarI, 3), "")), _
                    IIf(sCuenta < "31", Val(MshMontos.TextMatrix(VarI, 4)), IIf(sCuenta = "31" And gsTipoPlan = "0", Val(MshMontos.TextMatrix(VarI, 4)), "0")), _
                    IIf(sCuenta < "31", Val(MshMontos.TextMatrix(VarI, 5)), IIf(sCuenta = "31" And gsTipoPlan = "0", Val(MshMontos.TextMatrix(VarI, 5)), "0")), _
                    IIf(sCuenta < "31", Val(MshMontos.TextMatrix(VarI, 6)), IIf(sCuenta = "31" And gsTipoPlan = "0", Val(MshMontos.TextMatrix(VarI, 6)), Val(MshMontos.TextMatrix(VarI, 5)))), _
                    IIf(sCuenta < "31", Val(MshMontos.TextMatrix(VarI, 7)), IIf(sCuenta = "31" And gsTipoPlan = "0", Val(MshMontos.TextMatrix(VarI, 7)), Val(MshMontos.TextMatrix(VarI, 6)))), _
                    IIf(sCuenta < "31", Val(MshMontos.TextMatrix(VarI, 9)), IIf(sCuenta = "31" And gsTipoPlan = "0", Val(MshMontos.TextMatrix(VarI, 8)), Val(MshMontos.TextMatrix(VarI, 7)))), _
                    IIf(sCuenta < "31", 0, 0), _
                    IIf(sCuenta < "31", Val(MshMontos.TextMatrix(VarI, 8)), 0), "I", _
                    IIf(sCuenta = "31" And gsTipoPlan = "1", MshMontos.TextMatrix(VarI, 1), sCuenta)) Then
                        bGrabo = False
                        Exit For
                End If
            End If
        Next VarI
      End If
    End If
End If
Else
Exit Sub
End If
If bGrabo Then
    MsgBox "Asientos registrados, Correctamente", vbInformation + vbSystemModal, "ECB-CONT"
End If
End Sub

Private Sub cmdInsertarItem_Click()


If MsgBox("Desea agregar un movimiento", vbYesNo + vbInformation, "ECB-CONT") = vbYes Then

Dim nroRegAct As Integer
Dim nroRegIni As Integer
FrmManRegAcc_Det.Show
FrmManRegAcc_Det.sCuenta = Left(CbCuenta.Text, 2)
FrmManRegAcc_Det.VarPeriodo = Format(CbPeriodo.ListIndex, "00")
FrmManRegAcc_Det.bEntro = True
pSetFocus FrmManRegAcc_Det
Exit Sub
MshMontos.Col = 1
nroRegIni = MshMontos.Rows - 1

MshMontos.SetFocus
MshMontos.AddItem ""
If MshMontos.Rows > 1 Then
    nroRegAct = MshMontos.Rows - 1
    If (FrmManRegAcc_Det.MshMov.TextMatrix(nroRegAct, 5) <> "*") Then Exit Sub
End If

MshMontos.TextMatrix(nroRegAct, 1) = ""
MshMontos.TextMatrix(nroRegAct, 2) = ""
MshMontos.TextMatrix(nroRegAct, 3) = ""
MshMontos.TextMatrix(nroRegAct, 4) = ""
MshMontos.TextMatrix(nroRegAct, 5) = ""
MshMontos.TextMatrix(nroRegAct, 6) = ""
MshMontos.TextMatrix(nroRegAct, 7) = ""
MshMontos.TextMatrix(nroRegAct, 8) = ""
MshMontos.TextMatrix(nroRegAct, 9) = ""
MshMontos.TextMatrix(nroRegAct, 14) = ""
MshMontos.TextMatrix(nroRegAct, 11) = ""
MshMontos.TextMatrix(nroRegAct, 12) = ""

MshMontos.TextMatrix(nroRegAct, 14) = Val(MshMontos.TextMatrix(nroRegIni, 14)) + 1

MshMontos.Row = nroRegAct

MshMontos.Col = 1
Else
Exit Sub
End If

End Sub

Private Sub cmdRefresh_Click()
Call CbPeriodo_Click
End Sub

Private Sub cmdsalir_Click()
If MsgBox("Desea cerrar la ventana", vbYesNo + vbInformation, "ECB-CONT") = vbYes Then
Unload Me
End If
End Sub

Private Sub cmdVerificar_Click()
If MshMontos.TextMatrix(1, 1) = "" Then Exit Sub
MsgBox "Información Verificada, Satisfactoriamente", vbInformation + vbSystemModal, "Mensaje: Sistema"
End Sub

Public Sub RecibirDatos(ByVal lControl As String, ByVal param0 As String, ByVal param1 As String, ByVal param2 As String, Optional ByVal param3 As String)
   ' *** Dependiendo del control
'    Select Case Control
If gsTipoPlan = "0" Or (sCuenta = "30" Or sCuenta = "11") Then
    MshMontos.TextMatrix(MshMontos.Row, 11) = param0
    MshMontos.TextMatrix(MshMontos.Row, 1) = param1
    MshMontos.TextMatrix(MshMontos.Row, 2) = param2
    MshMontos.TextMatrix(MshMontos.Row, 12) = param3
Else
    MshMontos.TextMatrix(MshMontos.Row, 1) = param0
End If


'    Case "tdbtCuentaDesde" ' *** Caso Desde
''        tdbtCuentaDesde = Trim(param0)
''        tdbtDescripcionDesde = Trim(param1)
''        Unload frmBuscador
''        pSetFocus tdbtCuentaDesde
'    Case "tdbtCodigo" '     *** Caso Codigp
''        tdbtCodigo = Trim(param0)
''        Unload frmBuscador
''        pSetFocus tdbtCodigo
'    Case "tdbtCuentaHasta" ' *** Caso Desde
'        tdbtCuentaHasta = Trim(param0)
'        tdbtDescripcionHasta = Trim(param1)
'        Unload frmBuscador
'        pSetFocus tdbtCuentaDesde
        
'    End Select
End Sub

Private Sub Form_Activate()
MshMontos.SetFocus
End Sub

Private Sub Form_Load()
Call Sb_Formato
Dim rs As New ADODB.Recordset
Set rs = Fct_Obtener_Periodos_Sg_Anio_Empresa(gsEmpresa, gsAnio)

Do While rs.EOF = False
    CbPeriodo.AddItem rs(1).Value
    CbPeriodo.itemData(CbPeriodo.NewIndex) = rs(0).Value
    rs.MoveNext
Loop
rs.Close

If CbPeriodo.ListCount > 0 Then CbPeriodo.ListIndex = 0

Set rs = Fct_Obtener_Cuentas_Inversion(gsEmpresa, gsAnio)

Do While rs.EOF = False
    CbCuenta.AddItem rs(0).Value & " - " & rs(1).Value
    rs.MoveNext
Loop
rs.Close

If CbCuenta.ListCount > 0 Then CbCuenta.ListIndex = 0
sCuenta = Left(CbCuenta.Text, 2)

Me.Height = 4770
Me.Width = 14300
Call Centrar_form(Me)
Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))
MshMontos.TextMatrix(MshMontos.Row, 14) = 1
End Sub

Public Sub Sb_Formatear_Grilla(ByRef Par_MshEmp As MSHFlexGrid)
Par_MshEmp.Row = 0
On Error GoTo Control
For nCol = 0 To Par_MshEmp.COLS - 1
    Par_MshEmp.Col = nCol
    Par_MshEmp.CellAlignment = flexAlignCenterCenter
Next nCol
Control:
Par_MshEmp.Row = 1
Par_MshEmp.FixedRows = 1 ' Una linea fija para titulos
Par_MshEmp.RowHeight(0) = Me.TextHeight("M") * 4
End Sub

Private Sub Form_Resize()
On Error GoTo errHand
    If Me.WindowState <> vbMinimized Then
        'Call SeteaFondoForm(Me)
'        MshMontos.Left = (MshMontos.Left + MshMontos.Width) + 60
        MshMontos.Height = (Me.Height - 1600) - (cmdSalir.Height + 20)
        MshMontos.Width = (Me.Width - MshMontos.Left) - 170
        Line1.Y2 = (MshMontos.Height + MshMontos.Top) - 40
        
        MshMontos.Height = MshMontos.Height
        cmdSalir.Top = (MshMontos.Height + MshMontos.Top) + 60
        cmdSalir.Left = Me.Width - (cmdSalir.Width + 400)
    End If
Exit Sub
errHand:
End Sub

Private Sub Form_Unload(Cancel As Integer)

Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))
End Sub

Private Sub MshEmp_Click()
MshMontos.Row = MshEmp.Row
End Sub

Private Sub MshMontos_Click()
'MshEmp.Row = MshMontos.Row
End Sub
Public Sub Busqueda()
Dim sTabla As String
If MshMontos.TextMatrix(1, 1) = "" And gsTipoPlan = "0" Then Exit Sub
If MshMontos.Col = 3 Or (MshMontos.Col = 2 And gsTipoPlan = "1" And sCuenta <> "30" And sCuenta <> "11") Then
    Select Case sCuenta
        Case "11", "30"
            sTabla = "DENOMINACIONES"
        Case "31" And MshMontos.Col = 2 And gsTipoPlan = "1"
            sTabla = "TIPO_INVERSION"
        Case "31" And MshMontos.Col = 3 And gsTipoPlan = "1"
            sNroCol = MshMontos.Col
            sTabla = "BASE_MEDICION"
    End Select
    FrmManRegAcc_ListaDeno.sTabla = sTabla
    FrmManRegAcc_ListaDeno.Show 1
    If FrmManRegAcc_ListaDeno.Tag = "1" Then
        MshMontos.TextMatrix(MshMontos.Row, IIf(sCuenta = "31" And gsTipoPlan = "0", 3, IIf((sCuenta = "31" Or sCuenta = "30" Or sCuenta = "11") And gsTipoPlan = "1" And MshMontos.Col = 3, 3, 2))) = _
        FrmManRegAcc_ListaDeno.MshDeno.TextMatrix(FrmManRegAcc_ListaDeno.MshDeno.Row, 2)
        MshMontos.TextMatrix(MshMontos.Row, IIf(sCuenta < "31", 15, IIf(sCuenta = "31" And MshMontos.Col = 2, 15, 13))) = FrmManRegAcc_ListaDeno.MshDeno.TextMatrix(FrmManRegAcc_ListaDeno.MshDeno.Row, 1)
'        If sCuenta = "31" And gsTipoPlan = "0" Then
'            MshMontos.TextMatrix(MshMontos.Row, 11) = FrmManRegAcc_ListaDeno.MshDeno.TextMatrix(FrmManRegAcc_ListaDeno.MshDeno.Row, 1)
'        End If
    End If
    Unload FrmManRegAcc_ListaDeno
    
ElseIf gsTipoPlan = "0" And sCuenta = "31" And (MshMontos.Col = 5 Or MshMontos.Col = 6 Or MshMontos.Col = 7) Then
Exit Sub
'    Select Case sCuenta
'        Case "31"
'            sTabla = "BASE_MEDICION"
'    End Select
'    sNroCol = MshMontos.Col
'    FrmManRegAcc_ListaDeno.sTabla = sTabla
'    FrmManRegAcc_ListaDeno.Show 1
'    If FrmManRegAcc_ListaDeno.Tag = "1" Then
'        MshMontos.TextMatrix(MshMontos.Row, 2) = _
'        FrmManRegAcc_ListaDeno.MshDeno.TextMatrix(FrmManRegAcc_ListaDeno.MshDeno.Row, 2)
'        MshMontos.TextMatrix(MshMontos.Row, 11) = _
'        FrmManRegAcc_ListaDeno.MshDeno.TextMatrix(FrmManRegAcc_ListaDeno.MshDeno.Row, 1)
'
'    End If
'    Unload FrmManRegAcc_ListaDeno
    
ElseIf gsTipoPlan = "1" And (MshMontos.Col = 3 And sCuenta = "11") Then
    If gsTipoPlan = "1" And sCuenta = "31" Then Exit Sub
    VarMonto = 0
    FrmManRegAcc_IngNum.TxtNum.Text = Val(MshMontos.TextMatrix(MshMontos.Row, MshMontos.Col))
    FrmManRegAcc_IngNum.TxtNum.SelStart = 0
    FrmManRegAcc_IngNum.TxtNum.SelLength = Len(MshMontos.TextMatrix(MshMontos.Row, MshMontos.Col))
    FrmManRegAcc_IngNum.Show 1
    If MshMontos.Col = 3 Then MshMontos.TextMatrix(MshMontos.Row, MshMontos.Col) = Val(VarMonto): Exit Sub
    
    If MshMontos.Col = 5 Then
        MshMontos.TextMatrix(MshMontos.Row, MshMontos.Col) = Val(VarMonto)
    Else
        If Val(VarMonto) > 0 Then
        MshMontos.TextMatrix(MshMontos.Row, 6) = Format(Val(MshMontos.TextMatrix(MshMontos.Row, 5)) - Val(VarMonto), "0.00")

        End If
    End If
    Call Sb_Sumar(MshMontos.Row)
    
End If
End Sub

Private Sub MshMontos_DblClick()
Dim sTabla As String
If MshMontos.TextMatrix(1, 1) = "" And gsTipoPlan = "0" Then Exit Sub
If MshMontos.Col = 3 Or (MshMontos.Col = 2 And gsTipoPlan = "1" And sCuenta <> "30" And sCuenta <> "11") Then
    Select Case sCuenta
        Case "11", "30"
            sTabla = "DENOMINACIONES"
        Case "31" And MshMontos.Col = 2 And gsTipoPlan = "1"
            sTabla = "TIPO_INVERSION"
        Case "31" And MshMontos.Col = 3 And gsTipoPlan = "1"
            sNroCol = MshMontos.Col
            sTabla = "BASE_MEDICION"
    End Select
    FrmManRegAcc_ListaDeno.sTabla = sTabla
    FrmManRegAcc_ListaDeno.Show 1
    If FrmManRegAcc_ListaDeno.Tag = "1" Then
        MshMontos.TextMatrix(MshMontos.Row, IIf(sCuenta = "31" And gsTipoPlan = "0", 3, IIf((sCuenta = "31" Or sCuenta = "30" Or sCuenta = "11") And gsTipoPlan = "1" And MshMontos.Col = 3, 3, 2))) = _
        FrmManRegAcc_ListaDeno.MshDeno.TextMatrix(FrmManRegAcc_ListaDeno.MshDeno.Row, 2)
        MshMontos.TextMatrix(MshMontos.Row, IIf(sCuenta < "31", 15, IIf(sCuenta = "31" And MshMontos.Col = 2, 15, 13))) = FrmManRegAcc_ListaDeno.MshDeno.TextMatrix(FrmManRegAcc_ListaDeno.MshDeno.Row, 1)
'        If sCuenta = "31" And gsTipoPlan = "0" Then
'            MshMontos.TextMatrix(MshMontos.Row, 11) = FrmManRegAcc_ListaDeno.MshDeno.TextMatrix(FrmManRegAcc_ListaDeno.MshDeno.Row, 1)
'        End If
    End If
    Unload FrmManRegAcc_ListaDeno
    
ElseIf gsTipoPlan = "0" And sCuenta = "31" And (MshMontos.Col = 5 Or MshMontos.Col = 6 Or MshMontos.Col = 7) Then
Exit Sub
'    Select Case sCuenta
'        Case "31"
'            sTabla = "BASE_MEDICION"
'    End Select
'    sNroCol = MshMontos.Col
'    FrmManRegAcc_ListaDeno.sTabla = sTabla
'    FrmManRegAcc_ListaDeno.Show 1
'    If FrmManRegAcc_ListaDeno.Tag = "1" Then
'        MshMontos.TextMatrix(MshMontos.Row, 2) = _
'        FrmManRegAcc_ListaDeno.MshDeno.TextMatrix(FrmManRegAcc_ListaDeno.MshDeno.Row, 2)
'        MshMontos.TextMatrix(MshMontos.Row, 11) = _
'        FrmManRegAcc_ListaDeno.MshDeno.TextMatrix(FrmManRegAcc_ListaDeno.MshDeno.Row, 1)
'
'    End If
'    Unload FrmManRegAcc_ListaDeno
    
ElseIf gsTipoPlan = "1" And (MshMontos.Col = 3 And sCuenta = "11") Then
    If gsTipoPlan = "1" And sCuenta = "31" Then Exit Sub
    VarMonto = 0
    FrmManRegAcc_IngNum.TxtNum.Text = Val(MshMontos.TextMatrix(MshMontos.Row, MshMontos.Col))
    FrmManRegAcc_IngNum.TxtNum.SelStart = 0
    FrmManRegAcc_IngNum.TxtNum.SelLength = Len(MshMontos.TextMatrix(MshMontos.Row, MshMontos.Col))
    FrmManRegAcc_IngNum.Show 1
    If MshMontos.Col = 3 Then MshMontos.TextMatrix(MshMontos.Row, MshMontos.Col) = Val(VarMonto): Exit Sub
    
    If MshMontos.Col = 5 Then
        MshMontos.TextMatrix(MshMontos.Row, MshMontos.Col) = Val(VarMonto)
    Else
        If Val(VarMonto) > 0 Then
        MshMontos.TextMatrix(MshMontos.Row, 6) = Format(Val(MshMontos.TextMatrix(MshMontos.Row, 5)) - Val(VarMonto), "0.00")

        End If
    End If
    Call Sb_Sumar(MshMontos.Row)
    
End If
End Sub

Private Sub MshMontos_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 And MshMontos.Col = "1" Then
    If gsTipoPlan = "0" Or (sCuenta = "30" Or sCuenta = "11") Then
        Call LlamaBuscar(frmBuscador, "Entidades", "", "EntidadesR", Me, "", "")
    Else
        Call LlamaBuscar(frmBuscador, "Cuentas", "", "Cuentas0-31", Me, "", "")
    End If
    Unload FrmManRegAcc_ListaDeno
End If

If KeyCode = 112 And MshMontos.Col > "3" Then
Call SetMonto(KeyCode)
ElseIf KeyCode = 112 And MshMontos.Col <> "1" Then
Call Busqueda
End If


If MshMontos.Col = 1 And KeyCode = vbKeyF2 And MshMontos.TextMatrix(MshMontos.Row, 1) <> "" Then
    Call MshMontos_DblClick
End If
End Sub

Private Sub MshMontos_KeyPress(KeyAscii As Integer)
VarMonto = 0
If KeyAscii = 13 And MshMontos.Col < MshMontos.COLS - 1 And MshMontos.TextMatrix(MshMontos.Row, 1) <> "" Then
    If MshMontos.Col = MshMontos.Col + 1 < MshMontos.COLS Then
        MshMontos.Col = MshMontos.Col + 1
    Else
        MshMontos.Col = 1
        If MshMontos.Rows < MshMontos.Row = MshMontos.Row + 1 Then
            MshMontos.Row = MshMontos.Row + 1
        Else
            Exit Sub
        End If
    End If
    Exit Sub
End If

Select Case sCuenta
    Case "11", "30"
        If (IsNumeric(Chr$(KeyAscii)) Or Chr$(KeyAscii) = "-") And MshMontos.Col = 6 Then
            FrmManRegAcc_IngNum.TxtNum.Text = Chr$(KeyAscii)
            FrmManRegAcc_IngNum.TxtNum.SelStart = 1
        '    FrmManRegAcc_IngNum.TxtNum.SelLength = Len(MshMontos.TextMatrix(MshMontos.Row, MshMontos.Col))
            FrmManRegAcc_IngNum.TxtNum.Visible = True
            FrmManRegAcc_IngNum.TDBFecDoc.Visible = False
            FrmManRegAcc_IngNum.Show 1
            MshMontos.TextMatrix(MshMontos.Row, MshMontos.Col) = Val(VarMonto)
            Call Sb_Sumar(MshMontos.Row)
        ElseIf IsNumeric(Chr$(KeyAscii)) And MshMontos.Col > 1 And MshMontos.TextMatrix(MshMontos.Row, 1) <> "" And MshMontos.Col <> 9 And MshMontos.Col <> 2 And MshMontos.Col <> 3 Then
            FrmManRegAcc_IngNum.TxtNum.Text = Val(Chr$(KeyAscii))
            FrmManRegAcc_IngNum.TxtNum.SelStart = 1
            FrmManRegAcc_IngNum.TxtNum.Visible = True
            FrmManRegAcc_IngNum.TDBFecDoc.Visible = False
            
        '    FrmManRegAcc_IngNum.TxtNum.SelLength = Len(MshMontos.TextMatrix(MshMontos.Row, MshMontos.Col))
            FrmManRegAcc_IngNum.Show 1
'            If Val(VarMonto) > 0 Then
                MshMontos.TextMatrix(MshMontos.Row, MshMontos.Col) = Format(Val(VarMonto), "0.00")
            
'                If (Val(VarMonto) * IIf(MshMontos.Col = 5, -1, 1)) + Val(MshMontos.TextMatrix(MshMontos.Row, 8)) > 0 Then
'
'                Else
'                    MsgBox "El monto ingresado al procesarlo, devuelve el total un valor negativo, Verifique por favor...", vbCritical + vbSystemModal, "Mensaje: Sistema"
'                End If
                Call MshMontos_KeyPress(13)
'            End If
            Call Sb_Sumar(MshMontos.Row)
        End If
    Case "31" And gsTipoPlan = "0"
    If IsNumeric(Chr$(KeyAscii)) And MshMontos.Col > 1 And MshMontos.TextMatrix(MshMontos.Row, 1) <> "" And MshMontos.Col <> 8 Then
            FrmManRegAcc_IngNum.TxtNum.Text = Val(Chr$(KeyAscii))
            FrmManRegAcc_IngNum.TxtNum.SelStart = 1
        '    FrmManRegAcc_IngNum.TxtNum.SelLength = Len(MshMontos.TextMatrix(MshMontos.Row, MshMontos.Col))
            FrmManRegAcc_IngNum.TxtNum.Visible = True
            FrmManRegAcc_IngNum.TDBFecDoc.Visible = False
            FrmManRegAcc_IngNum.Show 1
'            If MshMontos.Col = 4 Or MshMontos.Col = 5 Then
'                MshMontos.TextMatrix(MshMontos.Row, MshMontos.Col) = Val(VarMonto)
'            Else
            If Val(VarMonto) > 0 Then
                MshMontos.TextMatrix(MshMontos.Row, MshMontos.Col) = Format(Val(VarMonto), "0.00")
                Call MshMontos_KeyPress(13)
            End If
'            End If
            Call Sb_Sumar(MshMontos.Row)
        End If
    Case "31" And gsTipoPlan = "1"
'        If MshMontos.Col = 4 Then Exit Sub
        If IsNumeric(Chr$(KeyAscii)) And MshMontos.Col > 1 And MshMontos.TextMatrix(MshMontos.Row, 1) <> "" And MshMontos.Col <> 3 And MshMontos.Col <> 2 Then
        '    FrmManRegAcc_IngNum.TxtNum.SelLength = Len(MshMontos.TextMatrix(MshMontos.Row, MshMontos.Col))
        
        If MshMontos.Col = 5 Or MshMontos.Col = 6 Or MshMontos.Col = 7 Or MshMontos.Col = 8 Or MshMontos.Col = 9 Then
            FrmManRegAcc_IngNum.TxtNum.Text = Val(Chr$(KeyAscii))
            FrmManRegAcc_IngNum.TxtNum.SelStart = 1
            FrmManRegAcc_IngNum.TxtNum.Visible = True
            FrmManRegAcc_IngNum.TDBFecDoc.Visible = False
             FrmManRegAcc_IngNum.Show 1
                MshMontos.TextMatrix(MshMontos.Row, MshMontos.Col) = Val(VarMonto)
        Else
            FrmManRegAcc_IngNum.TxtNum.Visible = False
            FrmManRegAcc_IngNum.TDBFecDoc.Visible = True
            If Not IsNull(VarBMed) Then
            FrmManRegAcc_IngNum.TDBFecDoc.Value = Format(Date, "dd/mm/yyyy")
            FrmManRegAcc_IngNum.Show 1
            FrmManRegAcc_IngNum.TDBFecDoc.SelStart = 1
            FrmManRegAcc_IngNum.TxtNum.Visible = False
            FrmManRegAcc_IngNum.TDBFecDoc.Visible = True
                MshMontos.TextMatrix(MshMontos.Row, MshMontos.Col) = VarBMed
                Call MshMontos_KeyPress(13)
        End If
            FrmManRegAcc_IngNum.TxtNum.Visible = True
            FrmManRegAcc_IngNum.TDBFecDoc.Visible = False
            End If
            Call Sb_Sumar(MshMontos.Row)
        End If
End Select
End Sub

Public Sub SetMonto(key As Integer)
VarMonto = 0
If key = 13 And MshMontos.Col < MshMontos.COLS - 1 And MshMontos.TextMatrix(MshMontos.Row, 1) <> "" Then
    If MshMontos.Col = MshMontos.Col + 1 < MshMontos.COLS Then
        MshMontos.Col = MshMontos.Col + 1
    Else
        MshMontos.Col = 1
        If MshMontos.Rows < MshMontos.Row = MshMontos.Row + 1 Then
            MshMontos.Row = MshMontos.Row + 1
        Else
            Exit Sub
        End If
    End If
    Exit Sub
End If

Select Case sCuenta
    Case "11", "30"
        If (IsNumeric(Chr$(key)) Or Chr$(key) = "-") And MshMontos.Col = 6 Then
            FrmManRegAcc_IngNum.TxtNum.Text = Chr$(key)
            FrmManRegAcc_IngNum.TxtNum.SelStart = 1
        '    FrmManRegAcc_IngNum.TxtNum.SelLength = Len(MshMontos.TextMatrix(MshMontos.Row, MshMontos.Col))
            FrmManRegAcc_IngNum.TxtNum.Visible = True
            FrmManRegAcc_IngNum.TDBFecDoc.Visible = False
            FrmManRegAcc_IngNum.Show 1
            MshMontos.TextMatrix(MshMontos.Row, MshMontos.Col) = Val(VarMonto)
            Call Sb_Sumar(MshMontos.Row)
        ElseIf IsNumeric(Chr$(key)) And MshMontos.Col > 1 And MshMontos.TextMatrix(MshMontos.Row, 1) <> "" And MshMontos.Col <> 9 And MshMontos.Col <> 2 And MshMontos.Col <> 3 Then
            FrmManRegAcc_IngNum.TxtNum.Text = Val(Chr$(key))
            FrmManRegAcc_IngNum.TxtNum.SelStart = 1
            FrmManRegAcc_IngNum.TxtNum.Visible = True
            FrmManRegAcc_IngNum.TDBFecDoc.Visible = False
            
        '    FrmManRegAcc_IngNum.TxtNum.SelLength = Len(MshMontos.TextMatrix(MshMontos.Row, MshMontos.Col))
            FrmManRegAcc_IngNum.Show 1
'            If Val(VarMonto) > 0 Then
                MshMontos.TextMatrix(MshMontos.Row, MshMontos.Col) = Format(Val(VarMonto), "0.00")
            
'                If (Val(VarMonto) * IIf(MshMontos.Col = 5, -1, 1)) + Val(MshMontos.TextMatrix(MshMontos.Row, 8)) > 0 Then
'
'                Else
'                    MsgBox "El monto ingresado al procesarlo, devuelve el total un valor negativo, Verifique por favor...", vbCritical + vbSystemModal, "Mensaje: Sistema"
'                End If
                Call MshMontos_KeyPress(13)
'            End If
            Call Sb_Sumar(MshMontos.Row)
        End If
    Case "31" And gsTipoPlan = "0"
    If IsNumeric(Chr$(key)) And MshMontos.Col > 1 And MshMontos.TextMatrix(MshMontos.Row, 1) <> "" And MshMontos.Col <> 8 Then
            FrmManRegAcc_IngNum.TxtNum.Text = Val(Chr$(key))
            FrmManRegAcc_IngNum.TxtNum.SelStart = 1
        '    FrmManRegAcc_IngNum.TxtNum.SelLength = Len(MshMontos.TextMatrix(MshMontos.Row, MshMontos.Col))
            FrmManRegAcc_IngNum.TxtNum.Visible = True
            FrmManRegAcc_IngNum.TDBFecDoc.Visible = False
            FrmManRegAcc_IngNum.Show 1
'            If MshMontos.Col = 4 Or MshMontos.Col = 5 Then
'                MshMontos.TextMatrix(MshMontos.Row, MshMontos.Col) = Val(VarMonto)
'            Else
            If Val(VarMonto) > 0 Then
                MshMontos.TextMatrix(MshMontos.Row, MshMontos.Col) = Format(Val(VarMonto), "0.00")
                Call MshMontos_KeyPress(13)
            End If
'            End If
            Call Sb_Sumar(MshMontos.Row)
        End If
    Case "31" And gsTipoPlan = "1"
'        If MshMontos.Col = 4 Then Exit Sub
        If IsNumeric(Chr$(key)) And MshMontos.Col > 1 And MshMontos.TextMatrix(MshMontos.Row, 1) <> "" And MshMontos.Col <> 3 And MshMontos.Col <> 2 Then
        '    FrmManRegAcc_IngNum.TxtNum.SelLength = Len(MshMontos.TextMatrix(MshMontos.Row, MshMontos.Col))
        
        If MshMontos.Col = 5 Or MshMontos.Col = 6 Or MshMontos.Col = 7 Or MshMontos.Col = 8 Or MshMontos.Col = 9 Then
            FrmManRegAcc_IngNum.TxtNum.Text = Val(Chr$(key))
            FrmManRegAcc_IngNum.TxtNum.SelStart = 1
            FrmManRegAcc_IngNum.TxtNum.Visible = True
            FrmManRegAcc_IngNum.TDBFecDoc.Visible = False
             FrmManRegAcc_IngNum.Show 1
                MshMontos.TextMatrix(MshMontos.Row, MshMontos.Col) = Val(VarMonto)
        Else
            FrmManRegAcc_IngNum.TxtNum.Visible = False
            FrmManRegAcc_IngNum.TDBFecDoc.Visible = True
            If Not IsNull(VarBMed) Then
            FrmManRegAcc_IngNum.TDBFecDoc.Value = Format(Date, "dd/mm/yyyy")
            FrmManRegAcc_IngNum.Show 1
            FrmManRegAcc_IngNum.TDBFecDoc.SelStart = 1
            FrmManRegAcc_IngNum.TxtNum.Visible = False
            FrmManRegAcc_IngNum.TDBFecDoc.Visible = True
                MshMontos.TextMatrix(MshMontos.Row, MshMontos.Col) = VarBMed
                Call MshMontos_KeyPress(13)
        End If
            FrmManRegAcc_IngNum.TxtNum.Visible = True
            FrmManRegAcc_IngNum.TDBFecDoc.Visible = False
            End If
            Call Sb_Sumar(MshMontos.Row)
        End If
End Select

End Sub

