VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmPrcExportarDatos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportar Datos del Sistema"
   ClientHeight    =   6120
   ClientLeft      =   2340
   ClientTop       =   4200
   ClientWidth     =   11295
   Icon            =   "frmPrcExportarDatos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6120
   ScaleWidth      =   11295
   Begin VB.Frame fraTodo 
      Height          =   6045
      Left            =   45
      TabIndex        =   3
      Top             =   0
      Width           =   11175
      Begin VB.Frame fraDirectorios 
         Caption         =   "DESTINO DE ARCHIVOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5310
         Left            =   180
         TabIndex        =   40
         Top             =   135
         Width           =   3945
         Begin VB.DriveListBox Drive1 
            Height          =   315
            Left            =   180
            TabIndex        =   42
            Top             =   315
            Width           =   3210
         End
         Begin VB.DirListBox Dir1 
            Height          =   3465
            Left            =   180
            TabIndex        =   41
            Top             =   675
            Width           =   3600
         End
         Begin TDBText6Ctl.TDBText tdbtDirectorio 
            Height          =   780
            Left            =   180
            TabIndex        =   43
            Top             =   4185
            Width           =   3615
            _Version        =   65536
            _ExtentX        =   6376
            _ExtentY        =   1376
            Caption         =   "frmPrcExportarDatos.frx":0ECA
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmPrcExportarDatos.frx":0F36
            Key             =   "frmPrcExportarDatos.frx":0F54
            BackColor       =   -2147483633
            EditMode        =   0
            ForeColor       =   -2147483640
            ReadOnly        =   -1
            ShowContextMenu =   -1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MarginBottom    =   1
            Enabled         =   -1
            MousePointer    =   0
            Appearance      =   0
            BorderStyle     =   1
            AlignHorizontal =   0
            AlignVertical   =   0
            MultiLine       =   -1
            ScrollBars      =   0
            PasswordChar    =   ""
            AllowSpace      =   -1
            Format          =   ""
            FormatMode      =   1
            AutoConvert     =   -1
            ErrorBeep       =   0
            MaxLength       =   0
            LengthAsByte    =   0
            Text            =   ""
            Furigana        =   0
            HighlightText   =   0
            IMEMode         =   0
            IMEStatus       =   0
            DropWndWidth    =   0
            DropWndHeight   =   0
            ScrollBarMode   =   0
            MoveOnLRKey     =   0
            OLEDragMode     =   0
            OLEDropMode     =   0
         End
         Begin MSForms.CommandButton cmdRefresh 
            Height          =   345
            Left            =   3420
            TabIndex        =   44
            ToolTipText     =   "Cargar Lista"
            Top             =   315
            Width           =   405
            PicturePosition =   327683
            Size            =   "714;609"
            Picture         =   "frmPrcExportarDatos.frx":0F98
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
      End
      Begin VB.Frame fraOpciones 
         Caption         =   "DATOS A EXPORTAR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3570
         Left            =   4230
         TabIndex        =   9
         Top             =   135
         Width           =   6720
         Begin VB.CheckBox chkMercaderias 
            Caption         =   "Mercaderias"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   4590
            TabIndex        =   47
            Top             =   3150
            Value           =   1  'Checked
            Width           =   1635
         End
         Begin VB.CheckBox chkMovim 
            Caption         =   "Movimientos"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   4590
            TabIndex        =   34
            Top             =   630
            Value           =   1  'Checked
            Width           =   1635
         End
         Begin VB.CheckBox chkCtaCte 
            Caption         =   "Cuenta corriente"
            Height          =   285
            Left            =   2340
            TabIndex        =   33
            Top             =   2385
            Value           =   1  'Checked
            Width           =   1905
         End
         Begin VB.CheckBox chkBancos 
            Caption         =   "Bancos"
            Height          =   285
            Left            =   2340
            TabIndex        =   32
            Top             =   2070
            Value           =   1  'Checked
            Width           =   1905
         End
         Begin VB.CheckBox chkConfigOp 
            Caption         =   "Config. de operaciones"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   2340
            TabIndex        =   31
            Top             =   945
            Value           =   1  'Checked
            Width           =   1995
         End
         Begin VB.CheckBox chkParamIni 
            Caption         =   "Parametros iniciales"
            Height          =   285
            Left            =   2340
            TabIndex        =   30
            Top             =   630
            Value           =   1  'Checked
            Width           =   1905
         End
         Begin VB.CheckBox chkPlanCONA 
            Caption         =   "Plantilla EEFF"
            Height          =   285
            Left            =   180
            TabIndex        =   29
            Top             =   1755
            Value           =   1  'Checked
            Width           =   1905
         End
         Begin VB.CheckBox chkTipAsto 
            Caption         =   "Tipo de asiento"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   2340
            TabIndex        =   28
            Top             =   1260
            Value           =   1  'Checked
            Width           =   1905
         End
         Begin VB.Frame Frame4 
            Height          =   3120
            Index           =   0
            Left            =   2115
            TabIndex        =   27
            Top             =   360
            Width           =   60
         End
         Begin VB.CheckBox chkTipMon 
            Caption         =   "Tipo de moneda"
            Height          =   285
            Left            =   180
            TabIndex        =   26
            Top             =   2070
            Value           =   1  'Checked
            Width           =   1905
         End
         Begin VB.CheckBox chkTabSec 
            Caption         =   "Tablas secundarias"
            Height          =   285
            Left            =   180
            TabIndex        =   25
            Top             =   2385
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   1905
         End
         Begin VB.CheckBox chkTipDoc 
            Caption         =   "Tipo de documentos"
            Height          =   285
            Left            =   180
            TabIndex        =   24
            Top             =   1440
            Value           =   1  'Checked
            Width           =   1905
         End
         Begin VB.CheckBox chkEntidades 
            Caption         =   "Entidades"
            Height          =   285
            Left            =   180
            TabIndex        =   23
            Top             =   1170
            Value           =   1  'Checked
            Width           =   1905
         End
         Begin VB.CheckBox chkLibros 
            Caption         =   "Libros"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   180
            TabIndex        =   22
            Top             =   900
            Value           =   1  'Checked
            Width           =   1905
         End
         Begin VB.CheckBox chkPlan 
            Caption         =   "Plan de cuentas"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   180
            TabIndex        =   21
            Top             =   630
            Value           =   1  'Checked
            Width           =   1905
         End
         Begin VB.CheckBox chkCenCos 
            Caption         =   "Centro de costo"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   2340
            TabIndex        =   20
            Top             =   2700
            Value           =   1  'Checked
            Width           =   1950
         End
         Begin VB.Frame Frame4 
            Height          =   3120
            Index           =   1
            Left            =   4365
            TabIndex        =   19
            Top             =   360
            Width           =   60
         End
         Begin VB.CheckBox chkTipoCbio 
            Caption         =   "Tipo de cambio"
            Height          =   285
            Left            =   180
            TabIndex        =   18
            Top             =   3060
            Value           =   1  'Checked
            Width           =   1500
         End
         Begin VB.CheckBox chkPatrimonio 
            Caption         =   "Patrimonio Neto"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   4590
            TabIndex        =   17
            Top             =   945
            Value           =   1  'Checked
            Width           =   1635
         End
         Begin VB.CheckBox chkFlujo 
            Caption         =   "Flujo de Efectivo"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   4590
            TabIndex        =   16
            Top             =   1260
            Value           =   1  'Checked
            Width           =   1635
         End
         Begin VB.CheckBox chkPresup 
            Caption         =   "Presupuesto"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   4590
            TabIndex        =   15
            Top             =   1575
            Value           =   1  'Checked
            Width           =   1635
         End
         Begin VB.CheckBox chkCapital 
            Caption         =   "Capital"
            CausesValidation=   0   'False
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   4590
            TabIndex        =   14
            Top             =   1890
            Value           =   1  'Checked
            Width           =   1635
         End
         Begin VB.CheckBox chkPDB 
            Caption         =   "PDB"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   4590
            TabIndex        =   13
            Top             =   2205
            Value           =   1  'Checked
            Width           =   1635
         End
         Begin VB.CheckBox chkRatios 
            Caption         =   "Ratios"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   2340
            TabIndex        =   12
            Top             =   3015
            Value           =   1  'Checked
            Width           =   1905
         End
         Begin VB.CheckBox chkValores 
            Caption         =   "Valores"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   4590
            TabIndex        =   11
            Top             =   2520
            Value           =   1  'Checked
            Width           =   1635
         End
         Begin VB.CheckBox chkCostos 
            Caption         =   "Costos Producción"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   4590
            TabIndex        =   10
            Top             =   2835
            Value           =   1  'Checked
            Width           =   1635
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "OPCIONALES"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   13
            Left            =   2340
            TabIndex        =   39
            Top             =   1710
            Width           =   1125
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "DEPENDIENTES"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   10
            Left            =   2340
            TabIndex        =   38
            Top             =   360
            Width           =   1290
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "PRINCIPALES"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   9
            Left            =   180
            TabIndex        =   37
            Top             =   315
            Width           =   1140
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "MOVIMIENTOS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   4590
            TabIndex        =   36
            Top             =   360
            Width           =   1230
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "TIPOS DE CAMBIO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   3
            Left            =   180
            TabIndex        =   35
            Top             =   2790
            Width           =   1530
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " PROCESO DE EXPORTACION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1425
         Index           =   0
         Left            =   4230
         TabIndex        =   4
         Top             =   4005
         Width           =   6735
         Begin MSComctlLib.ProgressBar pgbAvance 
            Height          =   195
            Left            =   180
            TabIndex        =   5
            Top             =   540
            Width           =   6360
            _ExtentX        =   11218
            _ExtentY        =   344
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
            Scrolling       =   1
         End
         Begin MSComctlLib.ProgressBar pgbAvanceTotal 
            Height          =   195
            Left            =   210
            TabIndex        =   6
            Top             =   1125
            Width           =   6360
            _ExtentX        =   11218
            _ExtentY        =   344
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
            Scrolling       =   1
         End
         Begin VB.Label lblAvance 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   135
            TabIndex        =   8
            Top             =   225
            Width           =   6360
         End
         Begin VB.Label lblAvanceTotal 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   210
            TabIndex        =   7
            Top             =   765
            Width           =   6360
         End
      End
      Begin VB.Label lblDesactivarTodo 
         AutoSize        =   -1  'True
         Caption         =   "Desactivar Todo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   9720
         TabIndex        =   49
         Top             =   3735
         Width           =   1185
      End
      Begin VB.Label lblSeleccionarTodo 
         AutoSize        =   -1  'True
         Caption         =   "Seleccionar Todo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   8190
         TabIndex        =   48
         Top             =   3735
         Width           =   1260
      End
      Begin MSForms.CommandButton cmdExportar 
         Height          =   435
         Left            =   4590
         TabIndex        =   45
         Top             =   5490
         Width           =   2115
         Caption         =   "   Exportar Datos"
         PicturePosition =   327683
         Size            =   "3731;767"
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
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   0
      TabIndex        =   46
      Top             =   0
      Width           =   4365
   End
   Begin VB.Label lblNomSucursal 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2160
      TabIndex        =   2
      Top             =   450
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Label lblCodSucursal 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1470
      TabIndex        =   1
      Top             =   630
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Sucursal :"
      Height          =   195
      Left            =   600
      TabIndex        =   0
      Top             =   675
      Visible         =   0   'False
      Width           =   705
   End
End
Attribute VB_Name = "frmPrcExportarDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lBolActivado As Boolean
Dim lsLibroCom As String
Dim lsLibroVen As String
Dim lsLibroHon As String
Dim lsLibroCajIng As String
Dim lsLibroCajEgr As String
Dim lsLibroDiario As String

Dim gsGrupo As String
Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property


Private Function ValidaFecha() As Boolean
    ValidaFecha = True
End Function


Private Sub chkCostos_Click()
    If chkCostos.Value = vbUnchecked Then
    Else
        chkMovim.Value = vbChecked
    End If

End Sub

Private Sub chkMercaderias_Click()
    If chkMercaderias.Value = vbUnchecked Then
    Else
        chkMovim.Value = vbChecked
    End If

End Sub

Private Sub chkValores_Click()
    If chkValores.Value = vbUnchecked Then
    Else
        chkMovim.Value = vbChecked
    End If

End Sub

Private Sub cmdExportar_Click()
    fraDirectorios.Enabled = False
    fraOpciones.Enabled = False
    DoEvents
    
    Call IniciarExportacion
    
    DoEvents
    fraDirectorios.Enabled = True
    fraOpciones.Enabled = True
    
End Sub

Private Sub IniciarExportacion()
    Dim respuesta As String
    Dim sqlExp As String
    Dim ruta As String
    Dim rsExp As ADODB.Recordset
    Dim varSuc As String
    
    respuesta = MsgBox("Desea exportar la información seleccionada", vbYesNo + vbQuestion, "Confirmar Exportación de Datos")
    If respuesta = vbNo Then Exit Sub

    cmdExportar.Enabled = False
    DoEvents
    
    Screen.MousePointer = vbHourglass
    ruta = Trim(tdbtDirectorio)
    If Not Right(ruta, 1) = "\" Then ruta = ruta & "\"
    
    varSuc = CodSucursal
    '------------------------ Configurar progess bar ------------------------
    pgbAvanceTotal.Min = 0
    pgbAvanceTotal.Max = 30
    pgbAvanceTotal.Value = 0
    lblAvanceTotal.Caption = "Iniciando proceso ..."
    
    
    Call EscribirLog("Iniciando exportacion de datos de la empresa " & gsEmpresa & " " & gsEmpresaNom, Me.Name)
    '------------------ CABECERA DE MOVIMIENTO ------------------------'
    If Me.chkMovim.Value = vbChecked Then
        Dim nConta As Long, cCade As String
        cCade = ""
        sqlExp = "SELECT Top 1 * From CNC_ASIENTO_VOUCHER"
        Set rsExp = ConsultarDatosRs(sqlExp)
        If Not rsExp Is Nothing Then
            For nConta = 0 To rsExp.Fields.Count - 1
                If Trim(cCade) <> "" Then cCade = cCade & ", "
                cCade = cCade & rsExp.Fields(nConta).Name
            Next
        End If
        Call CerrarRecordSet(rsExp)
        '------------------------------------------------------------------------------------'
        sqlExp = "Select " & cCade & " From CNC_ASIENTO_VOUCHER " & _
                 "Where Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' and " & _
                 " " & _
                 "(Lib_cTipoLibro='" & lsLibroCom & "' OR Lib_cTipoLibro='" & lsLibroVen & "'  OR Lib_cTipoLibro='" & lsLibroDiario & "' ) " & _
                 "ORDER BY Lib_cTipoLibro, Ase_nVoucher"
        Set rsExp = ConsultarDatosRs(sqlExp)
        If Not rsExp Is Nothing Then
            EliminaArchivo ruta & "AsientosC01.exp"
            rsExp.Save ruta & "AsientosC01.exp"
        End If
        '------------------------------------------------------------------------------------'
        sqlExp = "Select " & cCade & " From CNC_ASIENTO_VOUCHER " & _
                 "Where Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' and " & _
                 " " & _
                 "Lib_cTipoLibro<>'" & lsLibroCom & "' AND Lib_cTipoLibro<>'" & lsLibroVen & "' AND Lib_cTipoLibro<>'" & lsLibroDiario & "' " & _
                 "ORDER BY Lib_cTipoLibro, Ase_nVoucher"
        Set rsExp = ConsultarDatosRs(sqlExp)
        If Not rsExp Is Nothing Then
            EliminaArchivo ruta & "AsientosC02.exp"
            rsExp.Save ruta & "AsientosC02.exp"
        End If
        '------------------------------------------------------------------------------------'
        Call CerrarRecordSet(rsExp)
    End If
    
    '------------------ DETALLE DE MOVIMIENTO ------------------------'
    If Me.chkMovim.Value = vbChecked Then

        
        cCade = " A.Ase_cNummov,A.Emp_cCodigo,A.Pan_cAnio,A.Per_cPeriodo,A.Lib_cTipoLibro,A.Ase_nVoucher,A.Pla_cCuentaContable,A.Asd_nItem," & _
                "A.Asd_cGlosa,A.Asd_nDebeSoles,A.Asd_nHaberSoles,A.Asd_nTipoCambio,A.Asd_nDebeMonExt,A.Asd_nHaberMonExt,A.Cos_cCodigo," & _
                "A.Ten_cTipoEntidad,A.Ent_cCodEntidad,A.Asd_cTipoDoc,A.Asd_dFecDoc,A.Asd_cSerieDoc,A.Asd_cNumDoc,CASE WHEN YEAR(A.Asd_dFecVen) = 1900 THEN NULL ELSE A.Asd_dFecVen END AS Asd_dFecVen," & _
                "A.Asd_cTipoDocRef,A.Asd_dFecDocRef,A.Asd_cSerieDocRef,A.Asd_cNumDocRef,A.Com_cTipoIgv,A.Ecp_cOperacion,A.Asd_nMontoInafecto," & _
                "A.Asd_cBaseImp,A.Asd_cRetencion,A.Asd_cFlgSpot,A.Asd_dFechaSpot,A.Asd_cNumSpot,A.Asd_cDestino,A.Asd_nCorre,A.Asd_cProvCanc," & _
                "A.Imp_nPorcentaje,A.Asd_nPorcenDist,A.Asd_cOperaTC,A.Asd_cTipoMoneda,A.Asd_cMonedaCalculo,A.Asd_cEstado,A.Asd_cDeleted," & _
                "A.Asd_cUserCrea,A.Asd_dFechaCrea,A.Asd_cUserModifica,A.Asd_dFechaModifica,A.Asd_cEquipoUser,A.Tra_cCodigo,A.Asd_cFormaPago," & _
                "A.Asd_cMonAdic , A.Asd_cImpAdic, A.Asd_cComprobante, A.Asd_cProceso, A.Asd_cRegAux, A.Asd_cRegAuxDet, A.Asd_cConvMon, A.Asd_cManual, A.Asd_cGrupo, A.Id_Exoneracion, A.Id_Tipo_Renta, A.Id_Modalidad, A.Id_Aduana, A.Id_Clasific_Servicio "

        
       '------------------------------------------------------------------------------------'
        sqlExp = "Select " & cCade & " From CND_ASIENTO_VOUCHER A Inner Join CNC_ASIENTO_VOUCHER B On " & _
                 "A.Ase_cNummov = B.Ase_cNummov AND A.Emp_cCodigo = B.Emp_cCodigo AND " & _
                 "A.Pan_cAnio = B.Pan_cAnio AND A.Per_cPeriodo = B.Per_cPeriodo AND " & _
                 "A.Lib_cTipoLibro = B.Lib_cTipoLibro AND A.Ase_nVoucher = B.Ase_nVoucher " & _
                 "Where A.Emp_cCodigo = '" & gsEmpresa & "' AND A.Pan_cAnio = '" & gsAnio & "' and " & _
                 "(A.Lib_cTipoLibro='" & lsLibroCom & "' OR A.Lib_cTipoLibro='" & lsLibroVen & "'  OR A.Lib_cTipoLibro='" & lsLibroDiario & "' ) AND " & _
                 "B.Ase_cDeleted <> '*'  " & _
                 "ORDER BY A.Lib_cTipoLibro, A.Ase_nVoucher, A.Asd_nItem"
        Set rsExp = ConsultarDatosRs(sqlExp)
        If Not rsExp Is Nothing Then
            EliminaArchivo ruta & "AsientosD01.exp"
            rsExp.Save ruta & "AsientosD01.exp"
        End If
        '------------------------------------------------------------------------------------'
        sqlExp = "Select " & cCade & " From CND_ASIENTO_VOUCHER A Inner Join CNC_ASIENTO_VOUCHER B On " & _
                 "A.Ase_cNummov = B.Ase_cNummov AND A.Emp_cCodigo = B.Emp_cCodigo AND " & _
                 "A.Pan_cAnio = B.Pan_cAnio AND A.Per_cPeriodo = B.Per_cPeriodo AND " & _
                 "A.Lib_cTipoLibro = B.Lib_cTipoLibro AND A.Ase_nVoucher = B.Ase_nVoucher " & _
                 "Where A.Emp_cCodigo = '" & gsEmpresa & "' AND A.Pan_cAnio = '" & gsAnio & "' and " & _
                 "A.Lib_cTipoLibro<>'" & lsLibroCom & "' AND A.Lib_cTipoLibro<>'" & lsLibroVen & "' AND A.Lib_cTipoLibro<>'" & lsLibroDiario & "' AND " & _
                 "B.Ase_cDeleted <> '*'  " & _
                 "ORDER BY A.Lib_cTipoLibro, A.Ase_nVoucher, A.Asd_nItem"
        Set rsExp = ConsultarDatosRs(sqlExp)
        If Not rsExp Is Nothing Then
            EliminaArchivo ruta & "AsientosD02.exp"
            rsExp.Save ruta & "AsientosD02.exp"
        End If
        '------------------------------------------------------------------------------------'
        Call CerrarRecordSet(rsExp)
    End If
    
    '------------------ PLAN DE CUENTA ------------------'
    If Me.chkPlan.Value = vbChecked Then
        GeneraDAT "CCuentas", ruta & "CCuentas.exp", varSuc
    End If
    '------------------ PLAN DE CUENTA DETALLE ------------------'
    If Me.chkPlan.Value = vbChecked Then
        GeneraDAT "CCuentasDet", ruta & "CCuentasDet.exp", varSuc
    End If
    '------------------------ LIBROS ------------------------
    If chkLibros.Value = vbChecked Then
        GeneraDAT "CLibros", ruta & "CLibros.exp", varSuc
    End If
    '------------------ENTIDADES ------------------'
    If Me.chkEntidades.Value = vbChecked Then
        GeneraDAT "CEntidades", ruta & "Entidades.exp", varSuc
    End If
    If Me.chkEntidades.Value = vbChecked Then
        GeneraDAT "CEntidades2", ruta & "Entidades2.exp", varSuc
    End If
    If Me.chkEntidades.Value = vbChecked Then
        GeneraDAT "CEntidades3", ruta & "Entidades3.exp", varSuc
    End If
    '------------------ TIPOS DE DOCS ------------------'
    If Me.chkTipDoc.Value = vbChecked Then
        GeneraDAT "CDocumentos", ruta & "CDocumentos.exp", varSuc
    End If
    '------------------ CONASEV ------------------'
    If Me.chkPlanCONA.Value = vbChecked Then
        GeneraDAT "CConasev", ruta & "CConasev.exp", varSuc
    End If
    If Me.chkPlanCONA.Value = vbChecked Then
        GeneraDAT "CConasev2", ruta & "CConasev2.exp", varSuc
    End If
    '------------------ MONEDA ------------------'
    If Me.chkTipMon.Value = vbChecked Then
        GeneraDAT "CMoneda", ruta & "CMoneda.exp", varSuc
    End If
    '------------------ SECUNDARIAS ------------------'
    If Me.chkTabSec.Value = vbChecked Then
        GeneraDAT "CSec", ruta & "CSec.exp", varSuc
    End If
    '------------------ PARAMETROS ------------------'
    If Me.chkParamIni.Value = vbChecked Then
        GeneraDAT "CParam", ruta & "CParam.exp", varSuc
    End If
    If Me.chkParamIni.Value = vbChecked Then
        GeneraDAT "CParam2", ruta & "CParam2.exp", varSuc
    End If
    
    '------------------ CONF OPERACIONES ------------------'
    If Me.chkConfigOp.Value = vbChecked Then
        GeneraDAT "COpera", ruta & "COpera.exp", varSuc
    End If
    If Me.chkConfigOp.Value = vbChecked Then
        GeneraDAT "COpera2", ruta & "COpera2.exp", varSuc
    End If
    If Me.chkConfigOp.Value = vbChecked Then
        GeneraDAT "COpera3", ruta & "COpera3.exp", varSuc
    End If
    
    '------------------ TIPO DE ASIENTO ------------------'
    If Me.chkTipAsto.Value = vbChecked Then
        GeneraDAT "CTipoasiento", ruta & "CTipoasiento.exp", varSuc
    End If
    
    '------------------ BANCOS ------------------'
    If Me.chkBancos.Value = vbChecked Then
        GeneraDAT "CBancos", ruta & "CBancos.exp", varSuc
    End If
    
    '------------------ CUENTA CORRIENTE  ------------------'
    If Me.chkCtaCte.Value = vbChecked Then
        GeneraDAT "CCtacte", ruta & "CCtacte.exp", varSuc
    End If
    
    '------------------ CENTRO DE COSTO ------------------'
    If Me.chkCenCos.Value = vbChecked Then
        GeneraDAT "CCosto", ruta & "CCosto.exp", varSuc
    End If
    
    '------------------ TIPO DE CAMBIO ------------------'
    If Me.chkTipoCbio.Value = vbChecked Then
        GeneraDAT "CTc", ruta & "CTc.exp", varSuc
    End If
    If Me.chkTipoCbio.Value = vbChecked Then
        GeneraDAT "CTc2", ruta & "CTc2.exp", varSuc
    End If
    
    '------------------ PATRIMONIO NETO ------------------'
    If Me.chkPatrimonio.Value = vbChecked Then
        GeneraDAT "CPatrim", ruta & "CPatrim.exp", varSuc
    End If
   
    '------------------ FLUJO EJECTIVO ------------------'
    If Me.chkFlujo.Value = vbChecked Then
        GeneraDAT "CFlujoPro", ruta & "CFlujoPro.exp", varSuc
    End If
   
    If Me.chkFlujo.Value = vbChecked Then
        GeneraDAT "CFlujoRep", ruta & "CFlujoRep.exp", varSuc
    End If
   
    If Me.chkFlujo.Value = vbChecked Then
        GeneraDAT "CFlujoSal", ruta & "CFlujoSal.exp", varSuc
    End If
   
    If Me.chkFlujo.Value = vbChecked Then
        GeneraDAT "CFlujoCta", ruta & "CFlujoCta.exp", varSuc
    End If
   
    '------------------ PRESUPUESTO ------------------'
    If Me.chkPresup.Value = vbChecked Then
        GeneraDAT "CPresup", ruta & "CPresup.exp", varSuc
    End If
   
    '------------------ CAPITAL ------------------'
    If Me.chkCapital.Value = vbChecked Then
        GeneraDAT "CCapital", ruta & "CCapital.exp", varSuc
    End If
   
    '------------------ MERCADERIAS ------------------'
    If Me.chkMercaderias.Value = vbChecked Then
        GeneraDAT "CMercaderias", ruta & "CMercaderias.exp", varSuc
    End If
    
    '------------------ VALORES ------------------'
    If Me.chkValores.Value = vbChecked Then
        GeneraDAT "CVal", ruta & "CVal.exp", varSuc
    End If
   
    '------------------ REG AUXILIARES ------------------'
    If Me.chkMovim.Value = vbChecked Then
        GeneraDAT "CRegAux", ruta & "CRegAux.exp", varSuc
    End If
    
    '------------------ COSTOS ------------------'
    If Me.chkCostos.Value = vbChecked Then
        GeneraDAT "CCos", ruta & "CCos.exp", varSuc
    End If
   
    If Me.chkCostos.Value = vbChecked Then
        GeneraDAT "CCosInv", ruta & "CCosInv.exp", varSuc
    End If
   
    '------------------ PDB ------------------'
    If Me.chkPDB.Value = vbChecked Then
        GeneraDAT "CPDB", ruta & "CPDB.exp", varSuc
    End If
   
    '------------------ RATIOS ------------------'
    If Me.chkRatios.Value = vbChecked Then
        GeneraDAT "CRatios", ruta & "CRatios.exp", varSuc
    End If
    
    
    '------------------------ Leer Archivos dat para CONCIL BANCARIA ------------------------

    If chkCtaCte.Value = vbChecked Then
        GeneraDAT "CConcil", ruta & "CConcil.exp", varSuc
    End If
    '-----------------------------------------------------'
    
    pgbAvanceTotal.Value = pgbAvanceTotal.Max
    lblAvanceTotal.Caption = "Proceso terminado ... "
    lblAvanceTotal.Refresh
    
    Call EscribirLog("Finalizando exportacion de datos de la empresa " & gsEmpresa & " " & gsEmpresaNom, Me.Name)
    
    Screen.MousePointer = 0
    
    Mensajes "Proceso de exportación de datos terminado", vbOKOnly + vbInformation
    cmdExportar.Enabled = True
End Sub

Private Function GeneraDAT(tabla As String, archivo As String, Sucursal As String)
    Dim sqlExp As String
    Dim ruta As String
    Dim rsExp As ADODB.Recordset
    On Error GoTo ERROR
    If pgbAvanceTotal.Value + 1 > pgbAvanceTotal.Max Then pgbAvanceTotal.Max = pgbAvanceTotal.Max + 1

    pgbAvanceTotal.Value = pgbAvanceTotal.Value + 1
    pgbAvanceTotal.Refresh
    lblAvanceTotal.Caption = "Importando -> " & archivo
    lblAvanceTotal.Refresh
    DoEvents

    sqlExp = ExtraeCadenaSQL(tabla, Sucursal)
    
    Set rsExp = ConsultarDatosRs(sqlExp)
    If Not rsExp Is Nothing Then
        EliminaArchivo archivo
        rsExp.Save archivo
    End If
    Exit Function
ERROR:
    Mensajes Err.Description, vbOKOnly + vbInformation
End Function

Private Sub cmdRefresh_Click()
    Drive1.Refresh
    Dir1.Refresh
    
    
End Sub

Private Sub Drive1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pSetFocus Dir1
End If
End Sub

Private Sub Form_Activate()
    If Not lBolActivado Then
       If Trim(CodSucursal) = "" Then
          Mensajes "No definió la sucursal origen, para realizar la exportación de Datos", vbInformation
          Unload Me: Exit Sub
       End If
       
       lBolActivado = True
    End If
End Sub

'Private Sub pCargaCfgLibro()
'    Dim clDatos As New clsMantoTablas
'    Dim rsArreglo As ADODB.Recordset
'    Dim arrDatos() As Variant, sqlver As String
'
'    sqlver = "SELECT * From CNT_CONFIG_LIBROS WHERE Emp_cCodigo = '" & gsEmpresa & "'"
'    arrDatos = Array(sqlver)
'    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
'    If Not rsArreglo Is Nothing Then
'       lsLibroCom = CE(rsArreglo("Cfl_cCompras"))
'       lsLibroVen = CE(rsArreglo("Cfl_cVentas"))
'       lsLibroHon = CE(rsArreglo("Cfl_cHonorarios"))
'       lsLibroDiario = CE(rsArreglo("Cfl_cDiario"))
'
'       If Len(Trim(CE(rsArreglo("Cfl_cCaja")))) > 0 Then
'          lsLibroCajIng = CE(rsArreglo("Cfl_cCaja"))
'          lsLibroCajEgr = CE(rsArreglo("Cfl_cCaja"))
'       Else
'          lsLibroCajIng = CE(rsArreglo("Cfl_cCajaIngresos"))
'          lsLibroCajEgr = CE(rsArreglo("Cfl_cCajaEgresos"))
'       End If
'    End If
'
'    Call CerrarRecordSet(rsArreglo)
'End Sub

Private Sub Form_Load()
    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))
    
    
    Dim sqlver As String
    pCargaCfgLibro
    Call Centrar_form(Me)
        
    lBolActivado = False
    lblCodSucursal = CodSucursal
    
    sqlver = "SELECT Tab_cDescripCampo FROM Tabla Where Emp_cCodigo='" & gsEmpresa & "' And Tab_cTabla='031' And Tab_cCodigo='" & lblCodSucursal & "'"
    lblNomSucursal = ExtraeDescripcion(sqlver)
    
    Drive1.Drive = "C:\"
    Dir1.Path = "C:\"
    Dir1.Refresh
    
    If Mid(gsGrupo, 3, 1) <> "1" And gsGrupo <> gsPrivilegioAdmin Then
        Me.cmdExportar.Enabled = False
        
    Else
        Me.cmdExportar.Enabled = True
        
    End If
    
End Sub

Private Sub Dir1_Change()
    tdbtDirectorio.Text = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error Resume Next
    Dir1.Path = Drive1.Drive
    Dir1.Refresh
    tdbtDirectorio.Text = Dir1.Path
    If Err.Number > 0 Then
        Mensajes "Error Nro. " & Err.Number & " " & Err.Description, vbCritical
    End If
End Sub

Private Function CodSucursal() As String
    Dim sqlDatos As String
    Dim arrDatos() As Variant
    Dim rsArreglo As ADODB.Recordset
    Dim clDatos As clsMantoTablas
    
    Set clDatos = New clsMantoTablas
    CodSucursal = ""
    sqlDatos = "SELECT Emp_cCodSuc FROM EMPRESA Where Emp_cCodigo='" & gsEmpresa & "'"
    arrDatos = Array(sqlDatos)
    
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo Is Nothing Then Exit Function
    CodSucursal = NuloText(rsArreglo(0))
    
    Call CerrarRecordSet(rsArreglo)
End Function

Private Function ExtraeCadenaSQL(tabla As String, varSuc As String) As String
    Dim sql As String
    Select Case tabla
        Case "CRegAux":
        '                  "Blta_nBaseImpInafD,'" & varSuc & "' As Sucursal "
            sql = "select Asd_cNumDocORIGEN, Asd_cSerieDocORIGEN,Ase_cNummov, Ase_nVoucher,  Emp_cCodigo, Pan_cAnio, Per_cPeriodo, " & _
                  "Blta_cFlagLibro,Tdo_cCodigo, Blta_Correlativo, Asd_cSerieDoc, Asd_cNumDoc,Asd_dFecha, Blta_nBaseImp, Blta_nIGV," & _
                  "Blta_nTotal, Mon_cCodigo, Blta_nTipoCambio, Tca_nAuxiliar, Blta_nBaseImpEXT, Blta_nIGVEXT, Blta_nTotalEXT, " & _
                  "Blta_nBaseImpInaf, Blta_nOtros, Blta_nOtrosD, Blta_cInafecto,Tab_cTabla, Blta_num_doc, Blta_cNombres, " & _
                  "Blta_cApellidos, Blta_cDNI, Blta_cEstado, " & _
                  "Blta_nBaseImpInafD " & _
                  "from CNT_REG_BOLETAS " & _
                  "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' "
        
        Case "CConcil":
        '"Che_cUserCrea, Che_dFechaCrea, Che_cUserModifica, Che_dFechaModifica, Che_cEquipoUser,'" & varSuc & "' As Sucursal "
            sql = "select Emp_cCodigo, Pan_cAnio, Per_cPeriodo, Che_nVoucherPago, Che_cNummovVoucher, Che_nItemVoucher, Ban_cCodigo, " & _
                  "Cue_cNumCuenta, Che_cTipoMov, Che_cTipoDoc, Che_cOperaCheque, Che_dFechaCheque, Che_nTipoCambio, " & _
                  "Che_nMontoS, Che_nMontoD, Che_dFechaOpera, Che_cObservacion, Che_cGlosa, Che_cEstado, Che_cDeleted," & _
                  "Che_cUserCrea, Che_dFechaCrea, Che_cUserModifica, Che_dFechaModifica, Che_cEquipoUser " & _
                  "FROM CNM_MOV_CHEQUE " & _
                  "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' "
        
        
        Case "CVal":
        '"Val_cDesTitulo, Val_nValorNom, Val_nCantidad, Val_nCostoTot, Val_nProvTot, Val_nTotalNeto,'" & varSuc & "' As Sucursal "
            sql = "SELECT Emp_cCodigo, Pan_cAnio, Per_cPeriodo, Ase_nVoucher, Asd_nItem, Ten_cTipoEntidad, Ent_cCodentidad, Val_cTitulo, " & _
                  "Val_cDesTitulo, Val_nValorNom, Val_nCantidad, Val_nCostoTot, Val_nProvTot, Val_nTotalNeto " & _
                  "FROM CND_VALORES_DETALLE " & _
                  "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' "
        
        Case "CCos":
            'Sql = "SELECT Emp_cCodigo, Pan_cAnio, Cos_cCodigo, '" & varSuc & "' As Sucursal "
            sql = "SELECT Emp_cCodigo, Pan_cAnio, Cos_cCodigo " & _
                  "FROM CNT_COSTOS " & _
                  "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' "
        
        Case "CCosInv":
            '"pro_ncol05, pro_ncol06, pro_ncol07, pro_ncol08, pro_ncol09, pro_ncol10, pro_ncol11, pro_ncol12, pro_ncol13,'" & varSuc & "' As Sucursal "
            sql = "SELECT emp_ccodigo, pan_canio, per_ctipo, pro_cproceso, pro_ccodigo, pro_ncol01, pro_ncol02, pro_ncol03, pro_ncol04, " & _
                  "pro_ncol05, pro_ncol06, pro_ncol07, pro_ncol08, pro_ncol09, pro_ncol10, pro_ncol11, pro_ncol12, pro_ncol13 " & _
                  "FROM CND_COSTOS_SALDOS " & _
                  "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' "
        
        Case "CFlujoPro":
'                  "Pro_cFormulaD, Pro_cDetalleD, Pro_cTipoH, Pro_cFormulaH, Pro_cDetalleH, Pro_cMetodo ,'" & varSuc & "' As Sucursal "
            sql = "SELECT Emp_cCodigo, Pan_cAnio, Per_cPeriodo, Pro_cTipoCta, Pro_cCuenta, Pro_cActividad, Pro_cTipoD, " & _
                  "Pro_cFormulaD, Pro_cDetalleD, Pro_cTipoH, Pro_cFormulaH, Pro_cDetalleH, Pro_cMetodo " & _
                  "FROM CNT_FLUJO_PROCESO " & _
                  "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' "
        
        Case "CFlujoRep":
                  '"Rep_cCodTipo, Rep_cFormula, Rep_cValor, Pro_cMetodo ,'" & varSuc & "' As Sucursal "
            sql = "SELECT Emp_cCodigo, Pan_cAnio, Per_cPeriodo, Rep_cCuenta,  " & _
                  "Rep_cCodTipo, Rep_cFormula, Rep_cValor, Pro_cMetodo " & _
                  "FROM CNT_FLUJO_REPORTE " & _
                  "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' "
        
        Case "CFlujoSal":
            'Sql = "SELECT Emp_cCodigo, Pan_cAnio, Per_cPeriodo, Sal_cCodigo, Sal_nSaldo ,'" & varSuc & "' As Sucursal  "
            sql = "SELECT Emp_cCodigo, Pan_cAnio, Per_cPeriodo, Sal_cCodigo, Sal_nSaldo " & _
                  "FROM CNT_FLUJO_SALDOINI " & _
                  "WHERE Emp_cCodigo = '" & gsEmpresa & "' "
        
        Case "CFlujoCta":
            'Sql = "SELECT Emp_cCodigo, Pan_cAnio, Per_cPeriodo, Pro_cCuenta ,'" & varSuc & "' As Sucursal  "
            sql = "SELECT Emp_cCodigo, Pan_cAnio, Per_cPeriodo, Pro_cCuenta " & _
                  "FROM CNT_FLUJO_CUENTAS " & _
                  "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' "
        
        Case "CPatrim":
        '"Pat_cCol03, Pat_cCol04, Pat_cCol05, Pat_cCol06, Pat_cCol07,'" & varSuc & "' As Sucursal  "
            sql = "SELECT Emp_cCodigo, Pan_cAnio, Pat_cCodigo, Pat_cCol01, Pat_cCol02, " & _
                  "Pat_cCol03, Pat_cCol04, Pat_cCol05, Pat_cCol06, Pat_cCol07 " & _
                  "FROM CNM_PATRIMONIO " & _
                  "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' "
        
        Case "CCapital":
        '"Cap_nAPagadas, Cap_nAcciones, Cap_nPorcent,'" & varSuc & "' As Sucursal "
            sql = "SELECT  Emp_cCodigo, Pan_cAnio, Per_cPeriodo, Ase_nVoucher, Asd_nItem, Ten_cTipoEntidad, " & _
                  "Ent_cCodentidad, Cap_cAcciones, Cap_nImportes, Cap_nValorNom, Cap_nASuscritas, " & _
                  "Cap_nAPagadas, Cap_nAcciones, Cap_nPorcent " & _
                  "FROM CND_CAPITAL_DETALLE " & _
                  "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' "
        
        Case "CPresup":
        '"Prm_cUserModifica, Prm_dFechaModifica, Prm_cEquipoUser,'" & varSuc & "' As Sucursal "
            sql = "SELECT  Emp_cCodigo, Pan_cAnio, Per_cPeriodo, Prm_cTipo, Cos_cCodigo, Mon_cCodigo, " & _
                  "Prm_nMontoPreS, Prm_nTipoCambioPres, Prm_nMontoPreD, Prm_dFechaPres, " & _
                  "Prm_cObserva, Prm_cEstado, Prm_cDeleted, Prm_cUserCrea, Pm_dFechaCrea, " & _
                  "Prm_cUserModifica, Prm_dFechaModifica, Prm_cEquipoUser " & _
                  "FROM PRM_MARCO_PRES " & _
                  "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' "
        
        Case "CPDB":
        '"'" & varSuc & "' As Sucursal "
            sql = "SELECT Ase_cNummov,Ase_nVoucher,Emp_cCodigo,Pan_cAnio,Per_cPeriodo,Lib_cTipoLibro,Dco_cTipoPDB,Dco_nItem,Dco_cTipoComVen," & _
                  "Dco_cTipoComprob,Dco_cFecha,Dco_cSerie,Dco_cNumero,Dco_cTipoPer,Dco_cTipoDocId,Dco_cNumDocId,Dco_cNombre,Dco_cApePat," & _
                  "Dco_cApeMat,Dco_cNombre1,Dco_cNombre2,Dco_cCodMon,Dco_cCodDest,Dco_cNunDest,Dco_nBaseImp,Dco_nISC,Dco_nIGV,Dco_nOtros," & _
                  "Dco_cIndDetra,Dco_cCodTasaDetra,Dco_cNumDetra,Dco_cIndRete,Dco_cRefTipoComp,Dco_cRefSerieComp,Dco_cRefNumComp,Dco_cRefFechaEmi," & _
                  "Dco_nRefBaseImp,Dco_nRefIGV,Dco_cMedPago,Dco_cCodBaco,Dco_cNumOp,Dco_dFechaOp,Dco_cMontoOp,Dco_cEstado,Dco_cDeleted,Dco_cUserCrea," & _
                  "Dco_dFechaCrea,Dco_cUserModifica,Dco_dFechaModifica,Dco_cEquipoUser " & _
                  "FROM CND_ASIENTO_PDB " & _
                  "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' "
        
                 
        'Case "CCuentas":
        '    'Sql = "SELECT *,'" & varSuc & "' As Sucursal FROM CNM_PLAN_CTA  "
        '    Sql = "SELECT Emp_cCodigo, Pan_cAnio, Pla_cCuentaContable, Pla_cNombreCuenta, Pla_cTitulo, " &
        '             "Ten_cTipoEntidad, Pla_cCentroCosto, Pla_cProvision, Pla_cDifCambio, Pla_cOperaTC, " &
        '             "Pla_cRedondeo, Pla_cDocumento, Pla_cTipoCta, Pla_cCptoBG, Pla_cCptoBGDual, Pla_cCptoResFun, " &
        '             "Pla_cCptoResNat, Pla_cTipoAfect, Pla_cDetraccion, Pla_cRetencion, Pla_cPercepcion, " &
        '             "Pla_cCtaPresup, Pla_cEstado, Pla_cDeleted, Pla_cUserCrea, Pla_dFechaCrea, Pla_cUserModifica, " &
        '             "Pla_dFechaModifica , Pla_cEquipoUser FROM CNM_PLAN_CTA  " &
        '          "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' " &
        '          "ORDER BY Pla_cCuentaContable"
                 
        'Case "CCuentasDet"
        '    'Sql = "SELECT *,'" & varSuc & "' As Sucursal FROM CND_CUENTA_DIST "
        '     Sql = "SELECT * FROM CND_CUENTA_DIST " &
        '          "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' " &
        '          "ORDER BY Pla_cCuentaContable, Per_cPeriodo "
                  
        Case "CCuentas":
            sql = "SELECT *,'" & varSuc & "' As Sucursal FROM CNM_PLAN_CTA  " & _
                  "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' " & _
                  "ORDER BY Pla_cCuentaContable"
                 
        Case "CCuentasDet"
            sql = "SELECT *,'" & varSuc & "' As Sucursal FROM CND_CUENTA_DIST " & _
                  "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' " & _
                  "ORDER BY Pla_cCuentaContable, Per_cPeriodo "
                                   
        Case "CCosto":
        '"Cos_dFechaModifica, Cos_cEquipoUser,'" & varSuc & "' As Sucursal "
            sql = "SELECT Emp_cCodigo, Pan_cAnio, Cos_cCodigo, Cos_cDescripcion, Cos_cTitulo, Cos_cSTitulo, " & _
                  "Cos_cEstado, Cos_cDeleted,    Cos_cUserCrea, Cos_dFechaCrea, Cos_cUserModifica,  " & _
                  "Cos_dFechaModifica, Cos_cEquipoUser " & _
                  "FROM CNT_CENTRO_COSTO " & _
                  "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' "
                  
        Case "CLibros"
            '"Lib_cUserCrea, Lib_dFechaCrea, Lib_cUserModifica, Lib_dFechaModifica, Lib_cEquipoUser,'" & varSuc & "' As Sucursal, Lib_cCodSunat  "
            sql = "SELECT Emp_cCodigo, Pan_cAnio, Lib_cTipoLibro, Lib_cDescripcion, Lib_cTipOpe, " & _
                  "Lib_cFlagDocRef,Lib_cFlagAdelIgv, Lib_cFlagInafecto, Lib_cEstado, Lib_cDeleted, " & _
                  "Lib_cUserCrea, Lib_dFechaCrea, Lib_cUserModifica, Lib_dFechaModifica, Lib_cEquipoUser, Lib_cCodSunat  " & _
                  "FROM CNT_LIBRO_OPERA " & _
                  "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' "
    
        Case "CEntidades":
        '"Ten_cUserCrea, Ten_dFechaCrea, Ten_cUserModifica, Ten_dFechaModifica, Ten_cEquipoUser,'" & varSuc & "' As Sucursal "
            sql = "SELECT Emp_cCodigo, Ten_cTipoEntidad, Ten_cNombreEntidad, Ten_cEstado, Ten_cDeleted, " & _
                  "Ten_cUserCrea, Ten_dFechaCrea, Ten_cUserModifica, Ten_dFechaModifica, Ten_cEquipoUser " & _
                  "FROM CNT_ENTIDAD " & _
                  "WHERE Emp_cCodigo = '" & gsEmpresa & "'"
    
        Case "CEntidades2":
        '"Edoc_cUserCrea, Edoc_dFechaCrea, Edoc_cUserModifica, Edoc_dFechaModifica, Edoc_cEquipoUser,'" & varSuc & "' As Sucursal "
            sql = "SELECT Emp_cCodigo, Ten_cTipoEntidad, Edoc_cTipoPersona, Edoc_cTipoDoc, Edoc_cEstado, Edoc_cDeleted, " & _
                  "Edoc_cUserCrea, Edoc_dFechaCrea, Edoc_cUserModifica, Edoc_dFechaModifica, Edoc_cEquipoUser " & _
                  "FROM  CNT_ENTIDAD_DOCU " & _
                  "WHERE Emp_cCodigo = '" & gsEmpresa & "' "
    
        Case "CEntidades3":
'"Ent_cUserCrea, Ent_dFechaCrea, Ent_cUserModifica, Ent_dFechaModifica, Ent_cEquipoUser,'" & varSuc & "' As Sucursal "
            sql = "SELECT Emp_cCodigo, Ent_cCodEntidad, Ten_cTipoEntidad, Ent_cPersona, Ent_cDireccion, Ent_nRuc, " & _
                  "Ent_cRepresentante, Ent_cTipoDoc, Ent_cFlagPersona, Ent_cEstadoEntidad, Ent_cEstado, Ent_cDeleted, " & _
                  "Ent_cUserCrea, Ent_dFechaCrea, Ent_cUserModifica, Ent_dFechaModifica, Ent_cEquipoUser, Ent_cFlagDomiciliado, Id_Pais, Id_Vinculo_Economico, Id_Convenio, PorcentajeSunat " & _
                  "FROM  CNM_ENTIDAD " & _
                  "WHERE Emp_cCodigo = '" & gsEmpresa & "' "
    
        Case "CDocumentos":
        '"Tdo_cUserCrea, Tdo_dFechaCrea, Tdo_cUserModifica, Tdo_dFechaModifica, Tdo_cEquipoUser,'" & varSuc & "' As Sucursal "
            sql = "SELECT Emp_cCodigo, Tdo_cCodigo, Tdo_cNombreLargo, Tdo_cNombreCorto, " & _
                  "Tdo_cEstado, Tdo_cDaot, Tdo_cNatDaot, Tdo_cDeleted, " & _
                  "Tdo_cUserCrea, Tdo_dFechaCrea, Tdo_cUserModifica, Tdo_dFechaModifica, Tdo_cEquipoUser " & _
                  "FROM  CNT_TIPODOC " & _
                  "WHERE Emp_cCodigo = '" & gsEmpresa & "' "
        
        Case "CSec":
        '"Tab_cUserCrea, Tab_dFechaCrea, Tab_cUserModifica, Tab_dFechaModifica, Tab_cEquipoUser,Tab_cCodSunat ,'" & varSuc & "' As Sucursal "
            sql = "SELECT Emp_cCodigo, Tab_cTabla, Tab_cCodigo, Tab_cDescripCampo, Tab_cDescripTabla," & _
                  "Tab_nLongitud, Tab_cEstado, Tab_cMod01,Tab_cMod02,Tab_cMod03,Tab_cMod04,Tab_cMod05,Tab_cMod06,Tab_cMod07,Tab_cMod08,Tab_cMod09, Tab_cMod10, Tab_cDeleted," & _
                  "Tab_cUserCrea, Tab_dFechaCrea, Tab_cUserModifica, Tab_dFechaModifica, Tab_cEquipoUser,Tab_cCodSunat " & _
                  "FROM  TABLA " & _
                  "WHERE Emp_cCodigo = '" & gsEmpresa & "' "
            
        Case "CMoneda":
        '"Mon_cUserCrea, Mon_dFechaCrea, Mon_cUserModifica, Mon_dFechaModifica, Mon_cEquipoUser,Mon_cCodSunat, '" & varSuc & "' As Sucursal "
            sql = "SELECT Emp_cCodigo, Mon_cCodigo, Mon_cNombreLargo, Mon_cNombreCorto," & _
                  "Mon_cMNac, Mon_cMExt, Mon_cEstado, Mon_cDeleted," & _
                  "Mon_cUserCrea, Mon_dFechaCrea, Mon_cUserModifica, Mon_dFechaModifica, Mon_cEquipoUser,Mon_cCodSunat " & _
                  "FROM CNT_TIPO_MONEDA " & _
                  "WHERE Emp_cCodigo = '" & gsEmpresa & "' "
            
        Case "CTc":
        '"Tca_cEquipoUser,'" & varSuc & "' As Sucursal "
            sql = "SELECT Emp_cCodigo, Tca_dFecha, Tca_cCodigoOrigen, Tca_cCodigoDestino," & _
                  "Tca_nCompra, Tca_nVenta, Tca_nCompraP, Tca_nVentaP, Tca_cEstado," & _
                  "Tca_cDeleted, Tca_cUserCrea, Tca_dFechaCrea, Tca_cUserModifica, Tca_dFechaModifica," & _
                  "Tca_cEquipoUser " & _
                  "FROM CNT_TIPO_CAMBIO " & _
                  "WHERE Emp_cCodigo = '" & gsEmpresa & "' "


        Case "CTc2":
        '"Tca_cNov, Tca_cDic,'" & varSuc & "' As Sucursal "
            sql = "SELECT Emp_cCodigo, Pan_cAnio, Tca_cTipo, Tca_cMoneda, Tca_cEne, Tca_cFeb," & _
                  "Tca_cMar, Tca_cAbr, Tca_cMay, Tca_cJun , Tca_cJul, Tca_cAgo, Tca_cSet, Tca_cOct, " & _
                  "Tca_cNov, Tca_cDic " & _
                  "FROM CNT_TIPO_CAMBIO_MENSUAL " & _
                  "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' "

        Case "CBancos":
        '"Ban_cUserCrea, Ban_dFechaCrea, Ban_cUserModifica, Ban_dFechaModifica, Ban_cEquipoUser,'" & varSuc & "' As Sucursal, Ban_ccodSunat "
            sql = "SELECT Emp_cCodigo, Ban_cCodigo, Ban_cNombre, Ban_cEstado, Ban_cDeleted," & _
                  "Ban_cUserCrea, Ban_dFechaCrea, Ban_cUserModifica, Ban_dFechaModifica, Ban_cEquipoUser , Ban_ccodSunat " & _
                  "FROM CNT_BANCO " & _
                  "WHERE Emp_cCodigo = '" & gsEmpresa & "' "
                  
        Case "CCtacte":
        '"Cue_cUserCrea, Cue_dFechaCrea, Cue_cUserModifica, Cue_dFechaModifica, Cue_cEquipoUser,'" & varSuc & "' As Sucursal "
            sql = "SELECT Emp_cCodigo, Ban_cCodigo, Cue_cNumCuenta, Cue_cCuentaContable, Mon_cCodigo," & _
                  "Cue_dFechaApertura, Cue_dFechaCierre, Cue_cObservaCierre, Cue_nMonto," & _
                  "Cue_nNumChequeFin, Cue_nNumChequeIni, Cue_cEstado, Cue_cDeleted," & _
                  "Cue_cUserCrea, Cue_dFechaCrea, Cue_cUserModifica, Cue_dFechaModifica, Cue_cEquipoUser,Pan_cAnio " & _
                  "FROM CNM_CUENTA_BANCO " & _
                  "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' "
            
        Case "CTipoasiento":
        '"Asl_cUserCrea, Asl_dFechaCrea, Asl_cUserModifica, Asl_dFechaModifica, Asl_cEquipoUser,'" & varSuc & "' As Sucursal "
            sql = "SELECT Emp_cCodigo, Pan_cAnio, Lib_cTipoLibro, Asl_cOperacion, Asl_nSecuencia," & _
                  "Asl_cDescripcion, AslTipoMov, Asl_cCuenta, Asl_nPorcen, Asl_cEstado, Asl_cDeleted," & _
                  "Asl_cUserCrea, Asl_dFechaCrea, Asl_cUserModifica, Asl_dFechaModifica, Asl_cEquipoUser " & _
                  "FROM CNT_ASIENTO_LIBRO " & _
                  "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' "
                        
        Case "CConasev":
            Dim clsMante As clsMantoTablas
            Set clsMante = New clsMantoTablas
            
            ReDim lArrMnt(2) As Variant
            lArrMnt(0) = gsEmpresa           ' Empresa
            lArrMnt(1) = gsAnio              ' Anio
            
            If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_UpdPeriodoPlantillaEEFF ", lArrMnt(), True) = False Then
             Debug.Print "No se actualizo..."
            End If
                
                  '"Ppa_cUserCrea, Ppa_dFechaCrea, Ppa_cUseModifica, Ppa_dFechaModifica, Ppa_cEquipoUser, '" & varSuc & "' As Sucursal, Pan_cAnio "
                  
            sql = "SELECT Emp_cCodigo, Ppa_cTipoPlantilla, Ppa_cNumPlantilla, Ppa_cNombre," & _
                  "Ppa_cTitulo, Ppa_cCodigoRef, Ppa_cResult, Ppa_cEstado, Ppa_cDeleted," & _
                  "Ppa_cUserCrea, Ppa_dFechaCrea, Ppa_cUseModifica, Ppa_dFechaModifica, Ppa_cEquipoUser, Pan_cAnio " & _
                  "FROM CNA_TIPO_PLANTILLA " & _
                  "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "'"
            
        Case "CConasev2":
            '"Ecp_cUserCrea, Ecp_dFechaCrea, Ecp_cUserModifica, Ecp_dFechaModifica, Ecp_cEquipoUser,'" & varSuc & "' As Sucursal "
            sql = "SELECT Emp_cCodigo, Pan_cAnio, Ecp_cOperacion, Ecp_cDescripcion, Ecp_cEstado, Ecp_cDeleted," & _
                  "Ecp_cUserCrea, Ecp_dFechaCrea, Ecp_cUserModifica, Ecp_dFechaModifica, Ecp_cEquipoUser " & _
                  "FROM CNT_OPERA_ESTADO " & _
                  "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' "
            
        Case "CParam":
            '"Cfl_cUserCrea, Cfl_dFechaCrea, Cfl_cUserModifica, Cfl_dFechaModifica, Cfl_cEquipoUser,'" & varSuc & "' As Sucursal "
            sql = "SELECT Emp_cCodigo, Cfl_cCompras, Cfl_cVentas, Cfl_cCaja, Cfl_cCajaIngresos," & _
                  "Cfl_cCajaEgresos, Cfl_cHonorarios, Cfl_cPercepcion, Cfl_cRetencion," & _
                  "Cfl_nPorcIGV, Cfl_cEstado, Cfl_cDeleted, Cfl_cDiario,Cfl_cDifCam, Cfl_cCierre,Cfl_cNivelCC,Cfl_cMesCompras,Cfl_cNDigCtas,Cfl_cApertura,Cfl_cBaseDefCom, " & _
                  "Cfl_cUserCrea, Cfl_dFechaCrea, Cfl_cUserModifica, Cfl_dFechaModifica, Cfl_cEquipoUser, Pan_cAnio, Cfl_cTransAutomatico, Cfl_cLEVenta, Cfl_cTransferencia, Cfl_cLECompra, Cfl_cAjusteNIF, Cfl_cVersionLE " & _
                  "FROM CNT_CONFIG_LIBROS " & _
                  "WHERE Emp_cCodigo = '" & gsEmpresa & "' and Pan_cAnio='" & gsAnio & "'"
                  
        Case "CParam2":
        '"Opd_cUserCrea, Opd_dFechaCrea, Opd_cUserModifica, Opd_dFechaModifica, Opd_cEquipoUser,'" & varSuc & "' As Sucursal "
            sql = "SELECT Emp_cCodigo, Pan_cAnio, Lib_cTipoLibro, Tdo_cCodigo, Opd_cEstado, Opd_cDeleted," & _
                  "Opd_cUserCrea, Opd_dFechaCrea, Opd_cUserModifica, Opd_dFechaModifica, Opd_cEquipoUser " & _
                  "FROM CNT_LIBRO_TIPODOC " & _
                  "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' "
            
        Case "COpera":
        '"Cop_cUserCrea, Cop_dFechaCrea, Cop_cUserModifica, Cop_dFechaModifica, Cop_cEquipoUser,'" & varSuc & "' As Sucursal "
            sql = "SELECT Emp_cCodigo, Cop_cCodigo, Cop_cDescripcion, Cop_cTipo, Cop_cEstado, Cop_cDeleted," & _
                  "Cop_cUserCrea, Cop_dFechaCrea, Cop_cUserModifica, Cop_dFechaModifica, Cop_cEquipoUser " & _
                  "FROM CNT_CONFIG_OPERA " & _
                  "WHERE Emp_cCodigo = '" & gsEmpresa & "' "
            
        Case "COpera2":
'"Cod_cUserCrea, Cod_dFechaCrea, Cod_cUserModifica, Cod_dFechaModifica, Cod_cEquipoUser,'" & varSuc & "' As Sucursal "
            sql = "SELECT Emp_cCodigo, Pan_cAnio, Cop_cCodigo, Cod_cValorParam, Cod_nIgvPorc, Cod_cEstado, Cod_cDeleted," & _
                  "Cod_cUserCrea, Cod_dFechaCrea, Cod_cUserModifica, Cod_dFechaModifica, Cod_cEquipoUser " & _
                  "FROM CND_CONFIG_OPERA " & _
                  "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' "
            
        Case "COpera3":
'"Lib_cUserCrea, Lib_dFechaCrea, Lib_cUserModifica, Lib_dFechaModifica, Lib_cEquipoUser,'" & varSuc & "' As Sucursal "
            sql = "SELECT Emp_cCodigo, Pan_cAnio, Lib_cTipoLibro, Pla_cCuentaContable, Lib_cEstado, Lib_cDeleted," & _
                  "Lib_cUserCrea, Lib_dFechaCrea, Lib_cUserModifica, Lib_dFechaModifica, Lib_cEquipoUser " & _
                  "FROM CNM_LIBRO_CTA " & _
                  "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' "
            
        Case "CRatios":
'"Ind_cUserCrea, Ind_dFechaCrea, Ind_cUserModifica, Ind_dFechaModifica, Ind_cEquipoUser,'" & varSuc & "' As Sucursal "
            sql = "SELECT Emp_cCodigo, Ind_cCodCuenta, Ind_cDescripcion, Ind_cEstado, Ind_cDeleted," & _
                  "Ind_cUserCrea, Ind_dFechaCrea, Ind_cUserModifica, Ind_dFechaModifica, Ind_cEquipoUser " & _
                  "FROM CNT_CUENTA_INDI " & _
                  "WHERE Emp_cCodigo = '" & gsEmpresa & "' "
            
        Case "CMercaderias":
'"Mer_cDeleted,Mer_cUserCrea,Mer_dFechaCrea,Mer_cUserModifica,Mer_dFechaModifica,Mer_cEquipoUser,'" & varSuc & "' As Sucursal "
            sql = "SELECT Emp_cCodigo,Pan_cAnio,Per_cPeriodo,Mer_cMetodo,Mer_cItem,Mer_cCodigo,Mer_cTipo,Mer_cDescrip,Mer_cMedida,Mer_nCantidad,Mer_nCosto,Mer_nTotal,Pla_cCuentaContable,Mer_cEstado," & _
                  "Mer_cDeleted,Mer_cUserCrea,Mer_dFechaCrea,Mer_cUserModifica,Mer_dFechaModifica,Mer_cEquipoUser " & _
                  "FROM CNT_MERCADERIAS " & _
                  "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' "
            
    End Select
    
    ExtraeCadenaSQL = sql
End Function
Private Sub chkBancos_Click()
    If chkBancos.Value = vbUnchecked Then
        chkCtaCte.Value = vbUnchecked
    Else
        'chkCtaCte.Value = vbChecked
    End If
End Sub

Private Sub chkCapital_Click()
    If chkCapital.Value = vbUnchecked Then
    Else
        chkMovim.Value = vbChecked
    End If
End Sub

Private Sub chkCenCos_Click()
    If chkCenCos.Value = vbUnchecked Then
        chkPresup.Value = vbUnchecked
    End If
End Sub

Private Sub chkPDB_Click()
    If chkPDB.Value = vbUnchecked Then
    Else
        chkMovim.Value = vbChecked
    End If
    
End Sub

Private Sub chkConfigOp_Click()
    If chkConfigOp.Value = vbUnchecked Then
    Else
        chkPlan.Value = vbChecked
        chkLibros.Value = vbChecked
    End If
End Sub

Private Sub chkCtaCte_Click()
    If chkCtaCte.Value = vbUnchecked Then
    Else
        chkBancos.Value = vbChecked
    End If
End Sub

Private Sub chkFlujo_Click()
'    If chkFlujo.Value = vbUnchecked Then
'    Else
'        'chkPlanCONA.Value = vbChecked
'    End If
End Sub

Private Sub chkLibros_Click()
    If chkLibros.Value = vbUnchecked Then
        chkConfigOp.Value = vbUnchecked
        chkTipAsto.Value = vbUnchecked
    End If
End Sub

Private Sub chkMovim_Click()
    If chkMovim.Value = vbChecked Then
        chkPlan.Value = vbChecked
    Else
        chkPDB.Value = vbUnchecked
        chkCapital.Value = vbUnchecked
    End If
    
End Sub

Private Sub chkPatrimonio_Click()
    If chkPatrimonio.Value = vbChecked Then
        chkPlan.Value = vbChecked
    End If
End Sub

Private Sub chkPlan_Click()
    If chkPlan.Value = vbUnchecked Then
        chkMovim.Value = vbUnchecked
        chkCapital.Value = vbUnchecked
        chkPDB.Value = vbUnchecked
        chkPatrimonio.Value = vbUnchecked
        chkConfigOp.Value = vbUnchecked
        chkTipAsto.Value = vbUnchecked
    End If
    If chkPlan.Value = vbChecked Then
        chkConfigOp.Value = vbChecked
        chkPlanCONA.Value = vbChecked
    End If
End Sub

Private Sub chkPlanCONA_Click()
    If chkPlanCONA.Value = vbUnchecked Then
        'chkFlujo.Value = vbUnchecked
        chkRatios.Value = vbUnchecked
    Else
    End If
End Sub

Private Sub chkPresup_Click()
    If chkPresup.Value = vbChecked Then
        chkCenCos.Value = vbChecked
    End If
End Sub

Private Sub chkRatios_Click()
    If chkRatios.Value = vbUnchecked Then
    Else
        chkPlanCONA.Value = vbChecked
    End If
End Sub

Private Sub chkTipAsto_Click()
    If chkTipAsto.Value = vbUnchecked Then
    Else
        chkPlan.Value = vbChecked
        chkLibros.Value = vbChecked
    End If
End Sub

Private Sub chkTipMon_Click()
    If chkTipMon.Value = vbUnchecked Then
        chkTipoCbio.Value = vbUnchecked
    End If
End Sub

Private Sub chkTipoCbio_Click()
    If chkTipoCbio.Value = vbUnchecked Then
        chkTipMon.Value = vbUnchecked
    Else
        chkTipMon.Value = vbChecked
    End If
End Sub

Private Sub Form_Resize()
On Error GoTo errHand
    If Me.WindowState <> vbMinimized Then
        Call SeteaFondoForm(Me)
        Call Centrar_Objeto(fratodo, Me)
        Call CentrarTitulo(lblTitulo, fratodo, Me)
    End If
Exit Sub
errHand:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))

End Sub

Private Sub lblDesactivarTodo_Click()
    Call SeleccionarChecks(False, Me)
End Sub

Private Sub lblSeleccionarTodo_Click()
    Call SeleccionarChecks(True, Me)
End Sub

