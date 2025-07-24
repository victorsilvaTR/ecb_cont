VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmPrcActualizaSaldos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generar Actualizacion de Saldos"
   ClientHeight    =   3435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6630
   Icon            =   "frmPrcActualizaSaldos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3435
   ScaleWidth      =   6630
   Begin VB.Frame fraTodo 
      Height          =   3345
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   6525
      Begin VB.CheckBox chkMes 
         Caption         =   "Hasta el mes seleccionado"
         Height          =   285
         Left            =   1575
         TabIndex        =   3
         Top             =   900
         Width           =   3300
      End
      Begin TrueOleDBList70.TDBCombo tdbcMes 
         Height          =   300
         Left            =   1530
         TabIndex        =   1
         Top             =   1350
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   529
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         _DropdownWidth  =   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "codigo"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "descripcion"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   2
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).AllowRowSizing=   0   'False
         Splits(0).DividerStyle=   2
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=661"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=582"
         Splits(0)._ColumnProps(4)=   "Column(0)._VertColor=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=688"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=609"
         Splits(0)._ColumnProps(10)=   "Column(1)._VertColor=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits.Count    =   1
         Appearance      =   1
         BorderStyle     =   1
         ComboStyle      =   2
         AutoCompletion  =   0   'False
         LimitToList     =   0   'False
         ColumnHeaders   =   0   'False
         ColumnFooters   =   0   'False
         DataMode        =   5
         DefColWidth     =   0
         Enabled         =   -1  'True
         HeadLines       =   1
         FootLines       =   1
         RowDividerStyle =   1
         Caption         =   ""
         EditFont        =   "Size=6.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         LayoutName      =   ""
         LayoutFileName  =   ""
         MultipleLines   =   0
         EmptyRows       =   -1  'True
         CellTips        =   0
         EditHeight      =   299.906
         AutoSize        =   -1  'True
         GapHeight       =   30.047
         ListField       =   "descripcion"
         BoundColumn     =   ""
         IntegralHeight  =   0   'False
         CellTipsWidth   =   0
         CellTipsDelay   =   1000
         AutoDropdown    =   0   'False
         RowTracking     =   -1  'True
         RightToLeft     =   0   'False
         RowMember       =   ""
         MouseIcon       =   0
         MouseIcon.vt    =   3
         MousePointer    =   0
         MatchEntryTimeout=   2000
         OLEDragMode     =   0
         OLEDropMode     =   0
         AnimateWindow   =   0
         AnimateWindowDirection=   0
         AnimateWindowTime=   200
         AnimateWindowClose=   0
         DropdownPosition=   0
         Locked          =   0   'False
         ScrollTrack     =   0   'False
         RowDividerColor =   12632256
         RowSubDividerColor=   12632256
         AddItemSeparator=   ";"
         _PropDict       =   $"frmPrcActualizaSaldos.frx":0ECA
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000005&,.bold=0,.fontsize=675"
         _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Arial"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(20)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(26)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(40)  =   "Named:id=33:Normal"
         _StyleDefs(41)  =   ":id=33,.parent=0"
         _StyleDefs(42)  =   "Named:id=34:Heading"
         _StyleDefs(43)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(44)  =   ":id=34,.wraptext=-1"
         _StyleDefs(45)  =   "Named:id=35:Footing"
         _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(47)  =   "Named:id=36:Selected"
         _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(49)  =   "Named:id=37:Caption"
         _StyleDefs(50)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(51)  =   "Named:id=38:HighlightRow"
         _StyleDefs(52)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(53)  =   "Named:id=39:EvenRow"
         _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(55)  =   "Named:id=40:OddRow"
         _StyleDefs(56)  =   ":id=40,.parent=33"
         _StyleDefs(57)  =   "Named:id=41:RecordSelector"
         _StyleDefs(58)  =   ":id=41,.parent=34"
         _StyleDefs(59)  =   "Named:id=42:FilterBar"
         _StyleDefs(60)  =   ":id=42,.parent=33"
      End
      Begin MSComctlLib.ProgressBar pgbAvance 
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   2430
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblNroItem 
         Alignment       =   2  'Center
         Caption         =   "0"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1935
         TabIndex        =   16
         Top             =   3735
         Width           =   1455
      End
      Begin VB.Label lblNroVoucher 
         Alignment       =   2  'Center
         Caption         =   "0"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5070
         TabIndex        =   15
         Top             =   2925
         Width           =   1335
      End
      Begin VB.Label lblTitNroItem 
         Caption         =   "Nro. de Item :"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   540
         TabIndex        =   14
         Top             =   3690
         Width           =   1155
      End
      Begin VB.Label lblTitNroVoucher 
         Caption         =   "Nro. de Voucher :"
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   3750
         TabIndex        =   13
         Top             =   2910
         Width           =   1380
      End
      Begin VB.Label Label4 
         Caption         =   "Ubicación de Proceso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   180
         TabIndex        =   12
         Top             =   2880
         Width           =   1635
      End
      Begin VB.Label lblTitTotalOper 
         Caption         =   "Total de Operaciones en General :"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   180
         TabIndex        =   11
         Top             =   4185
         Width           =   2505
      End
      Begin VB.Label lblTitEnProc 
         AutoSize        =   -1  'True
         Caption         =   "Van en Proceso :"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   1455
         TabIndex        =   10
         Top             =   4590
         Width           =   1230
      End
      Begin VB.Label lblNumReg 
         Alignment       =   2  'Center
         Caption         =   "0"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2775
         TabIndex        =   9
         Top             =   4230
         Width           =   1410
      End
      Begin VB.Label lblNumRegProcesado 
         Alignment       =   2  'Center
         Caption         =   "0"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2775
         TabIndex        =   8
         Top             =   4590
         Width           =   1410
      End
      Begin VB.Label lblMes 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   465
         Left            =   765
         TabIndex        =   7
         Top             =   1845
         Width           =   5160
      End
      Begin VB.Label Label1 
         Caption         =   "NO REGISTRAR ningun movimiento  del mes seleccionado al ejecutar el Proceso."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   1575
         TabIndex        =   4
         Top             =   225
         Width           =   3315
      End
      Begin MSForms.CommandButton cmdProcesar 
         Height          =   435
         Left            =   2040
         TabIndex        =   2
         Top             =   2790
         Width           =   1665
         Caption         =   "Procesar"
         PicturePosition =   327683
         Size            =   "2937;767"
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
      TabIndex        =   5
      Top             =   0
      Width           =   4365
   End
End
Attribute VB_Name = "frmPrcActualizaSaldos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gsGrupo As String
Dim clsMante As clsMantoTablas
Public gsMensaje As String
Dim rsArreglo  As ADODB.Recordset
Dim rstDifCam As ADODB.Recordset
Dim rstCorrelativo As ADODB.Recordset
Dim rstSumCancelProv As ADODB.Recordset
Dim sSql As String
Dim nCont As Long
Dim Asd_nCorre As Long
Dim DifCam As String, DebeHaber As String
Dim CancelNac As Double, CancelExt As Double

Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

Private Sub cmdProcesar_Click()
    cmdProcesar.Enabled = False
    Call Procesar
    cmdProcesar.Enabled = True
End Sub

Public Sub Cerrar()
    Unload Me
End Sub

Public Sub Procesar()
    Dim i As Integer
    Dim Mes As String
    Dim Ultimo As Integer
    Dim inicio As Integer
    Dim Retorno As Boolean, RetornoProv As Boolean
    
    On Error GoTo serror
    Call EscribirLog("Iniciando la actuializacion de saldos de la empresa " & gsEmpresa & " " & gsEmpresaNom, Me.Name)
    
    lblMES.Caption = "INICIANDO ..."
    Screen.MousePointer = vbHourglass
    
    Ultimo = Val(tdbcMes.BoundText)
    
    pgbAvance.Min = 0
    If Ultimo = 0 Then
        pgbAvance.Max = 1
    Else
        pgbAvance.Max = Ultimo
    End If
    
    pgbAvance.Value = 0

    If chkMes.Value = vbUnchecked Then
        inicio = Ultimo
    Else
        inicio = 0
    End If
    
    For i = inicio To Ultimo
        Mes = Right("00" & i, 2)
        lblMES.Caption = "PROCES. " & NombreMes(Mes)
        lblMES.Refresh
        pgbAvance.Value = i
        pgbAvance.Refresh
        DoEvents
        Retorno = Proceso(Mes)
        lblMES.Refresh
        If Retorno = False Then Exit For
    Next i
    
    For i = inicio To Ultimo
        Mes = Right("00" & i, 2)
        lblMES.Caption = "PROCES.CTA. TITULOS DE " & NombreMes(Mes)
        lblMES.Refresh
        pgbAvance.Value = i
        pgbAvance.Refresh
        DoEvents
    
        Call ActualizaSaldosSp(Mes)
    Next i
    
    'frmPrcActualizaSaldos.Height = 6165
    'fraTodo.Height = 5685
    'cmdProcesar.Top = 5085

    Label4.Visible = True: lblTitNroVoucher.Visible = True: lblNroVoucher.Visible = True

    ConectarAdvance

    Set rstDifCam = New ADODB.Recordset
    rstDifCam.Open "spCn_ObtieneConfigLibro '" & gsEmpresa & "'", gcnSistemaAdv, adOpenDynamic, adLockOptimistic
    If rstDifCam.State > 0 Then DifCam = rstDifCam.Fields(0)

    gcnSistemaAdv.BeginTrans
    For i = inicio To Ultimo
        Mes = Right("00" & i, 2)
        lblMES.Caption = "ACTUALIZA SALDOS DE PROVISIONES DEL PERIODO : " & NombreMes(Mes)
        lblMES.Refresh
'        pgbAvance.Value = i
'        pgbAvance.Refresh
'        DoEvents
'
'        RetornoProv = ProcesarSaldosProv(Mes)
'        DoEvents
    
      If Not ExistenDatos(Mes) Then
       lblMES.Caption = "NO EXISTEN SALDO DE PROVISIONES EN EL PERIODO : " & NombreMes(Mes)
      End If

      nCont = 0
        If Not rsArreglo.EOF Then
          With rsArreglo
           .MoveFirst
           pgbAvance.Min = 0
           pgbAvance.Max = .RecordCount
           lblNumReg = .RecordCount
           Do While Not .EOF
            lblNroVoucher = IIf(IsNull(!Ase_nVoucher), "", !Ase_nVoucher)
            lblNroItem = IIf(IsNull(!asd_nitem), "", !asd_nitem)
            nCont = nCont + 1

'            Movimientos
            If (CDbl(!Asd_nHaberSoles) - CDbl(!Asd_nDebeSoles)) > 0 Then DebeHaber = "H" Else DebeHaber = "D"
              Set rstCorrelativo = New ADODB.Recordset
              rstCorrelativo.Open "spCn_ExtraeMaxCorreProvisiones", gcnSistemaAdv, adOpenDynamic, adLockOptimistic
              If rstCorrelativo.State > 0 Then Asd_nCorre = rstCorrelativo.Fields(0)

              CancelNac = 0: CancelExt = 0

              If Trim(!Lib_cTipoLibro) <> Trim(DifCam) Then
               Set rstSumCancelProv = New ADODB.Recordset
               rstSumCancelProv.Open "spCn_TotalizaCancelMNME '" & gsEmpresa & "','" & gsAnio & "','" & Trim(Mes) & "','" & _
                                     Trim(!Ent_cCodEntidad) & "','" & Trim(!Asd_cTipoDoc) & "','" & Trim(!Asd_cSerieDoc) & _
                                     "','" & Trim(!Asd_cNumDoc) & "','" & Trim(!Pla_cCuentaContable) & "'", gcnSistemaAdv, adOpenDynamic, adLockOptimistic
               If rstSumCancelProv.State > 0 Then CancelNac = IIf(IsNull(rstSumCancelProv.Fields(0)), 0, rstSumCancelProv.Fields(0)): CancelExt = IIf(IsNull(rstSumCancelProv.Fields(1)), 0, rstSumCancelProv.Fields(1))
              End If

              gcnSistemaAdv.Execute "spCn_InsertaMovAsientosProv " & Asd_nCorre & ",'" & Trim(!Ase_cNummov) & "','" & gsEmpresa & "','" & gsAnio & "','" & _
                                    Trim(Mes) & "','" & Trim(!Lib_cTipoLibro) & "','" & Trim(!Ase_nVoucher) & "'," & !asd_nitem & "," & !Asd_nHaberSoles & _
                                    "," & !Asd_nDebeSoles & "," & !Asd_nHaberMonExt & "," & !Asd_nDebeMonExt & "," & CancelNac & "," & CancelExt & ",'" & _
                                    DebeHaber & "'"
'              -------------------------------------
              gcnSistemaAdv.Execute "spCn_UpdateMovAsientoProv " & Asd_nCorre & ",'" & gsEmpresa & "','" & gsAnio & "','" & Trim(Mes) & _
                                    "'," & !asd_nitem & ",'" & Trim(!Ase_nVoucher) & "','" & Trim(!Ase_cNummov) & "'"
'              -------------------------------------
              gcnSistemaAdv.Execute "spCn_UpdateMovAsientoProvFin " & Asd_nCorre & ",'" & gsEmpresa & "','" & gsAnio & "','" & Trim(Mes) & _
                                    "'," & Trim(!Ent_cCodEntidad) & ",'" & Trim(!Pla_cCuentaContable) & "','" & Trim(!Asd_cTipoDoc) & _
                                    "','" & Trim(!Asd_cSerieDoc) & "','" & Trim(!Asd_cNumDoc) & "'"

              pgbAvance.Value = nCont
              lblNumRegProcesado = nCont
              DoEvents
              pgbAvance.Refresh
              .MoveNext
              If .EOF Then Exit Do
           Loop
          End With
        End If
    Next i

    gcnSistemaAdv.CommitTrans
    Desconectar

    'frmPrcActualizaSaldos.Height = 3810
    'fraTodo.Height = 3345
    'cmdProcesar.Top = 2790
    Label4.Visible = False: lblTitNroVoucher.Visible = False: lblNroVoucher.Visible = False
        
    Set rsArreglo = Nothing
    Set rstDifCam = Nothing
    Set rstCorrelativo = Nothing
    Set rstSumCancelProv = Nothing
        
    lblMES.Caption = "PROCESO TERMINADO ..."
    
    Call EscribirLog("Finalizo la Actualizacion de Saldos de la Empresa " & gsEmpresa & " " & gsEmpresaNom, Me.Name)
    
    If gsMensaje = True Then Mensajes "Proceso ha terminado con exito", vbInformation + vbOKOnly
    pgbAvance.Value = 0
    Screen.MousePointer = vbNormal
            
    Exit Sub
serror:
    Screen.MousePointer = vbNormal
    MsgBox Err.Description
'    Resume
    'gcnSistemaAdv.RollbackTrans
    Call EscribirLog("error de actualizacion de saldos , [" & Err.Description & "]  de la empresa " & gsEmpresa & " " & gsEmpresaNom, Me.Name)
End Sub

Private Sub Form_Activate()
    tdbcMes.BoundText = gsPeriodo
End Sub

Private Sub Form_Load()
    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))
    
    gsMensaje = True
    
    Call Centrar_form(Me)
    
    Call LlenaComboMesApeAddItem(tdbcMes)
    tdbcMes.ReBind
    tdbcMes.BoundText = gsPeriodo
    Me.chkMes.Value = vbChecked
    
    If Mid(gsGrupo, 3, 1) <> "1" And gsGrupo <> gsPrivilegioAdmin Then
        Me.cmdProcesar.Enabled = False
    Else
        Me.cmdProcesar.Enabled = True
    End If
    
    lblMES.Caption = ""
    pgbAvance.Value = 0
    Label4.Visible = False: lblTitNroVoucher.Visible = False: lblNroVoucher.Visible = False
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
Private Sub tdbcMes_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then pSendKeys "{tab}"
End Sub
Private Function ExistenDatos(Mes As String) As Boolean
On Error GoTo Error_cmd

Set rsArreglo = New ADODB.Recordset

 sSql = "spCn_ConsultaProvReproceso '" & gsEmpresa & "','" & gsAnio & "','" & Trim(Mes) & "'"
 rsArreglo.Open sSql, gcnSistemaAdv, adOpenDynamic, adLockOptimistic

ExistenDatos = IIf(rsArreglo.State > 0 And Not rsArreglo.EOF, True, False)

Exit Function
    
Error_cmd:
    ExistenDatos = False
    MsgBox Err.Description, vbInformation, App.Title
End Function


