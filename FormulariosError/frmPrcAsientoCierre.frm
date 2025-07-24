VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmPrcAsientoCierre 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generar Asiento de Cierre"
   ClientHeight    =   3555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5100
   Icon            =   "frmPrcAsientoCierre.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3555
   ScaleWidth      =   5100
   Begin VB.Frame fraTodo 
      Height          =   3480
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   5010
      Begin VB.Frame Frame2 
         Height          =   2640
         Left            =   135
         TabIndex        =   1
         Top             =   135
         Width           =   4725
         Begin TrueOleDBList70.TDBCombo tdbcLibro 
            Height          =   300
            Left            =   765
            TabIndex        =   2
            Tag             =   "enabled"
            Top             =   1665
            Width           =   3660
            _ExtentX        =   6456
            _ExtentY        =   529
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            _DropdownWidth  =   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).DataField=   ""
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   3
            Splits(0)._UserFlags=   0
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).AllowRowSizing=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=3"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=609"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=529"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=847"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=767"
            Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=1138"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1058"
            Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(15)=   "Column(2).Visible=0"
            Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
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
            RowDividerStyle =   0
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
            ListField       =   ""
            BoundColumn     =   ""
            IntegralHeight  =   0   'False
            CellTipsWidth   =   0
            CellTipsDelay   =   1000
            AutoDropdown    =   -1  'True
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
            _PropDict       =   $"frmPrcAsientoCierre.frx":0ECA
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=675,.italic=0"
            _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
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
            _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
            _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
            _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
            _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
            _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Named:id=33:Normal"
            _StyleDefs(45)  =   ":id=33,.parent=0"
            _StyleDefs(46)  =   "Named:id=34:Heading"
            _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(48)  =   ":id=34,.wraptext=-1"
            _StyleDefs(49)  =   "Named:id=35:Footing"
            _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(51)  =   "Named:id=36:Selected"
            _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(53)  =   "Named:id=37:Caption"
            _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(55)  =   "Named:id=38:HighlightRow"
            _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(57)  =   "Named:id=39:EvenRow"
            _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(59)  =   "Named:id=40:OddRow"
            _StyleDefs(60)  =   ":id=40,.parent=33"
            _StyleDefs(61)  =   "Named:id=41:RecordSelector"
            _StyleDefs(62)  =   ":id=41,.parent=34"
            _StyleDefs(63)  =   "Named:id=42:FilterBar"
            _StyleDefs(64)  =   ":id=42,.parent=33"
         End
         Begin TDBDate6Ctl.TDBDate dtpFechaCierre 
            Height          =   300
            Left            =   2835
            TabIndex        =   3
            Top             =   2160
            Width           =   1545
            _Version        =   65536
            _ExtentX        =   2725
            _ExtentY        =   529
            Calendar        =   "frmPrcAsientoCierre.frx":0F51
            Caption         =   "frmPrcAsientoCierre.frx":1053
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmPrcAsientoCierre.frx":10B7
            Keys            =   "frmPrcAsientoCierre.frx":10D5
            Spin            =   "frmPrcAsientoCierre.frx":1141
            AlignHorizontal =   0
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
            FirstMonth      =   1
            ForeColor       =   -2147483640
            Format          =   "dd/mm/yyyy"
            HighlightText   =   0
            IMEMode         =   3
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxDate         =   73415
            MinDate         =   2
            MousePointer    =   0
            MoveOnLRKey     =   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            PromptChar      =   "_"
            ReadOnly        =   0
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__/__/____"
            ValidateMode    =   0
            ValueVT         =   2010185729
            Value           =   38202
            CenturyMode     =   0
         End
         Begin VB.Label lblMensaje 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "GENERAR ASIENTO DE CIERRE FINAL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   765
            TabIndex        =   6
            Top             =   450
            Width           =   3660
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Seleccione el Libro donde se generará el asiento automático, el asiento se creara en el mes de CIERRE"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Index           =   0
            Left            =   765
            TabIndex        =   5
            Top             =   855
            Width           =   3660
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha de cierre"
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
            Left            =   765
            TabIndex        =   4
            Top             =   2205
            Width           =   1545
         End
      End
      Begin MSForms.CommandButton cmd_salir 
         Height          =   435
         Left            =   2625
         TabIndex        =   8
         Top             =   2895
         Width           =   1665
         Caption         =   " Salir"
         PicturePosition =   327683
         Size            =   "2937;767"
         Picture         =   "frmPrcAsientoCierre.frx":1169
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdGenerar 
         Height          =   435
         Left            =   735
         TabIndex        =   7
         Top             =   2895
         Width           =   1665
         Caption         =   " Generar"
         PicturePosition =   327683
         Size            =   "2937;767"
         Picture         =   "frmPrcAsientoCierre.frx":1703
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
      TabIndex        =   9
      Top             =   0
      Width           =   4365
   End
End
Attribute VB_Name = "frmPrcAsientoCierre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gsGrupo As String
Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

Private Sub cmd_salir_Click()
    Unload Me
End Sub

Private Function EliminarAsientosCierre()
    'Dim mensaje
    'mensaje = MsgBox("Desea eliminar todos los asientos creados en el periodo de cierre, del Año Actual", vbYesNoCancel + vbQuestion, "Confirmación de eliminación de asientos")
    'If mensaje = vbYes Then
            Screen.MousePointer = vbHourglass
            Dim clsMante As clsMantoTablas
            Dim lArrMnt(10) As Variant
            Set clsMante = New clsMantoTablas
            lArrMnt(0) = "ELIMINACIERRE"
            lArrMnt(1) = ""
            lArrMnt(2) = gsEmpresa
            lArrMnt(3) = gsAnio
            lArrMnt(4) = "14" 'mes de cierre
            lArrMnt(5) = tdbcLibro.BoundText
            lArrMnt(6) = ""
            lArrMnt(7) = ""
            lArrMnt(8) = ""
            lArrMnt(9) = ""
            lArrMnt(10) = ""
            
            If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_ConsultaAsientos", lArrMnt(), True) = False Then
                Mensajes "El proceso no ha concluido. Verificar...", vbInformation
                Screen.MousePointer = vbDefault
                'EliminarAsientosCierre = False
                Set clsMante = Nothing
                Exit Function
            End If
    
            Set clsMante = Nothing
            
            Screen.MousePointer = vbDefault
            'Mensajes "Se eliminaron los asientos del periodo de cierre, con exito", vbInformation
            'EliminarAsientosCierre = True

    'End If
    
    'If mensaje = vbNo Then
    '    EliminarAsientosCierre = True
    'End If
    
    'If mensaje = vbCancel Then
    '    EliminarAsientosCierre = False
    'End If
    
End Function

Private Function ValidaResEjercicio() As Boolean
    On Error GoTo serror
    Dim rsCta As ADODB.Recordset
    ValidaResEjercicio = False
    Dim scadena As String
    scadena = BuscaValorEnOp("035")
        
        
    If scadena <> "" Then
        ValidaResEjercicio = True
    Else
        ValidaResEjercicio = False
    End If
    
            
    Exit Function
serror:
    ValidaResEjercicio = False
End Function

Private Sub cmdGenerar_Click()
    Dim respuesta As String
    Dim sql As String
    
    If CE(tdbcLibro.Text) = "" Then
        Mensajes "Debe crear el libro de diferencia en cambio y configurelo en Parametros iniciales", vbOKOnly + vbInformation
        Exit Sub
    End If
    
    If IsDate(dtpFechaCierre.Value) = False Or dtpFechaCierre.Text = "__/__/____" Then
        Mensajes "Ingrese una Fecha de Cierre válida"
        pSetFocus dtpFechaCierre
        Exit Sub
    End If
   
    If IsDate(dtpFechaCierre.Value) = True Then
        If dtpFechaCierre.Year <> gsAnio Then
        Mensajes "El año de la Fecha de Cierre, es invalido" & Salto(1) & "Debe ser igual al año del sistema " & gsAnio
        pSetFocus dtpFechaCierre
        Exit Sub
        End If
    End If
    
'    If CierreMes("14") Then
'        Mensajes "El mes de CIERRE esta bloqueado no se puede generar los asientos automaticos"
'        Exit Sub
'    End If
    
    sql = "select * from CNT_CIERRE where Emp_cCodigo = '" & gsEmpresa & "' and Pan_cAnio = '" & gsAnio & "' and TipoLibro = '08' and Per_cPeriodo = '14' and Cic_cEstado = 'I'"
    If ExisteDato(sql) = True Then Mensajes "No se puede procesar el Cierre, debido a que el periodo se encuentra bloqueado", vbInformation: Exit Sub
    
    sql = "select * from CNT_CIERRE where Emp_cCodigo = '" & gsEmpresa & "' and Pan_cAnio = '" & gsAnio & "' and TipoLibro = '08' and Per_cPeriodo = '14'"
    If ExisteDato(sql) = True Then
        Mensajes "Esta corrección modificará los datos ingresados, la misma que será informada a la SUNAT en el período " + UCase(MonthName(Month(lsFecha))) + " del ejercicio " + Str(Year(lsFecha)) + "."
        If MsgBox("¿Desea continuar?", vbQuestion + vbOKCancel, gsNombreModulo) = vbCancel Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If
    
    Dim rsCta As ADODB.Recordset
    ' *** Validar tipo de cambio
    'If NumeroLleno(tdbnTipoCambio, "Tipo de Cambio") = False Then Exit Sub
    Dim CtaRes As String
    respuesta = MsgBox("Desea Generar el Asiento de Cierre del Año Actual" & Salto(1) & "En el periodo de CIERRE", vbYesNo + vbQuestion, "Confirmar Generar Asiento de Cierre")
    If respuesta = vbYes Then
        If ValidaResEjercicio = False Then
            Mensajes "Configure la cuenta de RESULTADO DEL EJERCICIO en el plan de cuentas"
            cmdGenerar.Enabled = True
            Screen.MousePointer = vbDefault
            
            Exit Sub
        End If
    
        ' verifica el saldo de la cuenta 89 en el periodo de ajuste
        If BuscaSaldoCuenta89 = True Then
            cmdGenerar.Enabled = True
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    
        ' elimina los asientos del periodo cierre
        'If EliminarAsientosCierre = False Then
        '    cmdGenerar.Enabled = True
        '    Screen.MousePointer = vbDefault
        '    Exit Sub
        'End If
        
        EliminarAsientosCierre
        
        Screen.MousePointer = vbHourglass
        Dim clsMante As clsMantoTablas
        Dim lArrMnt(7) As Variant
        
        Set clsMante = New clsMantoTablas
        ' *** Generando el Asiento
        lArrMnt(0) = gsEmpresa
        lArrMnt(1) = gsAnio
        lArrMnt(2) = Me.tdbcLibro.BoundText
        lArrMnt(3) = 1
        lArrMnt(4) = gsUsuario
        lArrMnt(5) = "COM"
        lArrMnt(6) = gintBiMoneda
        lArrMnt(7) = dtpFechaCierre.Value
        
        cmd_salir.Enabled = False
        cmdGenerar.Enabled = False
        DoEvents
        
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GeneraAsientoCierre", lArrMnt(), True) = False Then
            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
            Screen.MousePointer = vbDefault
            
            cmd_salir.Enabled = True
            cmdGenerar.Enabled = True
            
            Exit Sub
        End If
        
        
        ' ***
        Call ActualizaSaldos
        
        Screen.MousePointer = vbDefault
        Mensajes "El Asiento se ha creado correctamente" & Salto(1) & "en el periodo de Cierre", vbInformation
        
        cmd_salir.Enabled = True
        cmdGenerar.Enabled = True
    End If
End Sub

Private Sub ActualizaSaldos()
      
        'Mensajes "SE INICIARA LA ACTUALIZACION DE SALDOS", vbOKOnly + vbExclamation
        frmPrcActualizaSaldos.Show
        frmPrcActualizaSaldos.cmdProcesar.Visible = False
        DoEvents
        frmPrcActualizaSaldos.chkMes.Value = vbChecked
        frmPrcActualizaSaldos.chkMes.Enabled = False
        DoEvents
        frmPrcActualizaSaldos.tdbcMes.BoundText = "14"
        DoEvents
        frmPrcActualizaSaldos.gsMensaje = False
        frmPrcActualizaSaldos.Procesar
        DoEvents
        frmPrcActualizaSaldos.Cerrar

End Sub

Private Function BuscaSaldoCuenta89() As Boolean
    On Error GoTo serror
    Dim sql As String
    Dim arrDatos() As Variant
    Dim rsAddItem As ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Dim SaldoMN As Double, SaldoME As Double
    
    sql = "spCn_ConsultaCuentas 'BUSCA_SALDO89', '" & gsEmpresa & "', '" & gsAnio & "'"
    
    Set clDatos = New clsMantoTablas
    arrDatos = Array(sql)
    Set rsAddItem = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())

    SaldoMN = 0
    SaldoME = 0
    
    If Not rsAddItem Is Nothing Then
        Do While Not rsAddItem.EOF
            SaldoMN = SaldoMN + NE(rsAddItem.Fields("MontoSoles"))
            SaldoME = SaldoME + NE(rsAddItem.Fields("MontoDolares"))
            rsAddItem.MoveNext
        Loop
    End If
    
    
    If gsByMoneda = 0 Then
        SaldoME = SaldoMN 'si es solo moneda nacional entonces valida solo SaldoMN
    End If

    If SaldoMN <> 0 Or SaldoME <> 0 Then
        If gsByMoneda = 0 Then
            Mensajes "Verifique el saldo de la cuenta 89" & Salto(2) & "Saldo Nac: " & CE(SaldoMN)
        Else
            Mensajes "Verifique el saldo de la cuenta 89" & Salto(2) & "Saldo Nac: " & CE(SaldoMN) & Salto(1) & "Saldo Ext: " & CE(SaldoME)
        End If
        BuscaSaldoCuenta89 = True
    Else
        BuscaSaldoCuenta89 = False
    End If
    
    Call CerrarRecordSet(rsAddItem)
    Set clDatos = Nothing
    
    Exit Function
serror:
    BuscaSaldoCuenta89 = False
End Function

Private Sub Form_Load()
    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))
    
    
    Dim sqlcombos As String
    pCargaCfgLibro
    DoEvents
    
    Call Centrar_form(Me)
    
    ' *** Llenando los libros
    sqlcombos = "SELECT LIB_CTIPOLIBRO, LIB_CDESCRIPCION FROM CNT_LIBRO_OPERA " & _
                "WHERE EMP_CCODIGO =  '" & gsEmpresa & "' AND LIB_CTIPOLIBRO ='" & lsLibroCierre & "' AND Pan_cAnio = '" & gsAnio & "' ORDER BY LIB_CDESCRIPCION "
    LlenarComboAddItem tdbcLibro, sqlcombos
    
    DoEvents
    
    If Mid(gsGrupo, 3, 1) <> "1" And gsGrupo <> gsPrivilegioAdmin Then
        Me.cmdGenerar.Enabled = False
    Else
        Me.cmdGenerar.Enabled = True
    End If
    
    tdbcLibro.Bookmark = 0
    
    On Error Resume Next
    dtpFechaCierre.Value = UltimoDiaMes("12", gsAnio)
End Sub

Private Sub Form_Resize()
On Error GoTo errHand
    If Me.WindowState <> vbMinimized Then
        Call SeteaFondoForm(Me)
        Call Centrar_Objeto(fraTodo, Me)
        Call CentrarTitulo(lblTitulo, fraTodo, Me)
    End If
Exit Sub
errHand:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))

End Sub

Private Sub tdbcLibro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then pSendKeys "{tab}"
End Sub


