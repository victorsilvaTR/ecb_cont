VERSION 5.00
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmPrcAsientoApertura 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generar Asiento de Apertura"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7905
   Icon            =   "frmPrcAsientoApertura.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3735
   ScaleWidth      =   7905
   Begin VB.Frame fraTodo 
      Height          =   3705
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   7800
      Begin VB.Frame Frame2 
         Height          =   3270
         Left            =   2820
         TabIndex        =   4
         Top             =   240
         Width           =   4725
         Begin TrueOleDBList70.TDBCombo tdbcLibro 
            Height          =   300
            Left            =   1215
            TabIndex        =   5
            Tag             =   "enabled"
            Top             =   1395
            Width           =   2940
            _ExtentX        =   5186
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
            _PropDict       =   $"frmPrcAsientoApertura.frx":0ECA
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "LIBRO DONDE SE CREARA EL ASIENTO"
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
            Index           =   1
            Left            =   990
            TabIndex        =   9
            Top             =   855
            Width           =   3240
         End
         Begin MSForms.CommandButton cmdGenerar 
            Height          =   435
            Left            =   990
            TabIndex        =   8
            Top             =   2475
            Width           =   1665
            Caption         =   " Generar"
            PicturePosition =   327683
            Size            =   "2937;767"
            Picture         =   "frmPrcAsientoApertura.frx":0F51
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmd_salir 
            Height          =   435
            Left            =   2835
            TabIndex        =   7
            Top             =   2475
            Width           =   1665
            Caption         =   " Salir"
            PicturePosition =   327683
            Size            =   "2937;767"
            Picture         =   "frmPrcAsientoApertura.frx":14EB
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "El asiento se creara en el mes de APERTURA"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   3
            Left            =   720
            TabIndex        =   6
            Top             =   1935
            Width           =   3780
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3270
         Left            =   225
         TabIndex        =   1
         Top             =   225
         Width           =   2535
         Begin VB.Label Label1 
            Caption         =   $"frmPrcAsientoApertura.frx":1A85
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1725
            Index           =   0
            Left            =   135
            TabIndex        =   3
            Top             =   1395
            Width           =   2100
         End
         Begin VB.Label Label1 
            Caption         =   "NOTA: CAMBIE EL AÑO DEL SISTEMA AL NUEVO AÑO CONTABLE GENERADO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   960
            Index           =   2
            Left            =   180
            TabIndex        =   2
            Top             =   450
            Width           =   2190
         End
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
      TabIndex        =   10
      Top             =   0
      Width           =   4365
   End
End
Attribute VB_Name = "frmPrcAsientoApertura"
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

Private Sub cmdGenerar_Click()
    Dim respuesta As String
    Dim sql As String
    If CE(tdbcLibro.Text) = "" Then
        Mensajes "Cree el libro diario y configurelo en Parametros iniciales", vbOKOnly + vbInformation
        Exit Sub
    End If
    
    If CierreMes("00") Then
        Mensajes "El mes APERTURA esta bloqueado no se puede generar los asientos automaticos"
        Exit Sub
    End If
    
    sql = "select * from CNT_CIERRE where Emp_cCodigo = '" & gsEmpresa & "' and Pan_cAnio = '" & gsAnio & "' and TipoLibro = '01' and Per_cPeriodo = '00' and Cic_cEstado = 'I'"
    If ExisteDato(sql) = True Then Mensajes "Si desea volver a procesar el Asiento de Apertura, elimine el Libro Electrónico generado." & Salto(2) & "Si desea modificar, " & _
    "dirijase al menú Procesos, luego a la opción Bloquear/Desbloquear Meses y proceda a desbloquear el Libro y Periodo", vbInformation: Exit Sub

    sql = "select * from CNT_CIERRE where Emp_cCodigo = '" & gsEmpresa & "' and Pan_cAnio = '" & gsAnio & "' and TipoLibro = '01'"
    If ExisteDato(sql) = True Then
        sql = "select * from CNT_lIBROSGENERADOS where Emp_cCodigo = '" & gsEmpresa & "' and Pan_cAnio = '" & gsAnio & "' and Per_cPeriodo = '01' and Lib_cTipoLibro = '03' and Estado ='A'"
        If ExisteDato(sql) = True Then
            Mensajes "No se puede procesar el Asiento de Apertura, debido a que el Libro Electrónico del periodo de Enero del " + gsAnio + " se encuentra generado." & Salto(2) & "Elimine el libro para procesar el Asiento de Apertura.", vbInformation
                Exit Sub
        End If
    End If
    
    respuesta = MsgBox("PARA GENERAR EL ASIENTO DE APERTURA" & Salto(1) & "SE ELIMINARAN TODOS LOS VOUCHERS DEL MES DE APERTURA DEL " & Salto(2) & "AÑO " & gsAnio & Salto(2) & "Desea continuar...", vbYesNo + vbQuestion, "Confirmar Generar Asiento de Apertura")
    respuesta = MsgBox("SE ELIMINARAN LOS VOUCHERS DEL " & Salto(2) & "MES : APERTURA " & Salto(2) & "AÑO :" & gsAnio & Salto(2) & "Desea continuar...", vbYesNo + vbQuestion, "Confirmar Generar Asiento de Apertura")

    If respuesta = vbYes Then
        Screen.MousePointer = vbHourglass
        Dim clsMante As clsMantoTablas
        Dim lArrMnt(6) As Variant
        
        Set clsMante = New clsMantoTablas
        ' *** Generando el Asiento
        lArrMnt(0) = gsEmpresa
        lArrMnt(1) = gsAnio
        lArrMnt(2) = Me.tdbcLibro.BoundText
        lArrMnt(3) = 1
        lArrMnt(4) = gsUsuario
        lArrMnt(5) = "COM"
        lArrMnt(6) = gsByMoneda
        
        cmd_salir.Enabled = False
        cmdGenerar.Enabled = False
        DoEvents
                
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GeneraAsientoApertura", lArrMnt(), True) = False Then
            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
            Screen.MousePointer = vbDefault
            
            cmd_salir.Enabled = True
            cmdGenerar.Enabled = True
            Exit Sub
        End If
        ' ***
        ActualizaSaldos
        DoEvents
        Screen.MousePointer = vbDefault
        
        Mensajes "El Asiento de Apertura se ha realizado correctamente", vbInformation
        
        cmd_salir.Enabled = True
        cmdGenerar.Enabled = True
        
    End If
End Sub

Private Sub ActualizaSaldos()

        'Mensajes "SE INICIARA LA ACTUALIZACION DE SALDOS", vbOKOnly + vbExclamation
        frmPrcActualizaSaldos.Show
        DoEvents
        frmPrcActualizaSaldos.chkMes.Value = vbChecked
        frmPrcActualizaSaldos.tdbcMes.BoundText = "00"
        DoEvents
        frmPrcActualizaSaldos.gsMensaje = False
        frmPrcActualizaSaldos.Procesar
        DoEvents
        frmPrcActualizaSaldos.Cerrar

End Sub

Private Sub Form_Load()
    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))
    
    
    Dim sqlcombos As String
    Dim registros As Integer
    pCargaCfgLibro
    DoEvents
    
    Call Centrar_form(Me)
    
    ' *** Llenando los libros
    sqlcombos = "SELECT LIB_CTIPOLIBRO, LIB_CDESCRIPCION FROM CNT_LIBRO_OPERA " & _
                "WHERE EMP_CCODIGO =  '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' AND (LIB_CTIPOLIBRO <>'" & lsLibroCom & "' AND LIB_CTIPOLIBRO <>'" & lsLibroVen & "' AND LIB_CTIPOLIBRO <>'" & lsLibroCajEgr & "' AND LIB_CTIPOLIBRO <>'" & lsLibroCajIng & "' AND LIB_CTIPOLIBRO <>'" & lsLibroCierre & "' AND LIB_CTIPOLIBRO <>'" & lsLibroDif & "' )  ORDER BY LIB_CDESCRIPCION "
    registros = LlenarComboAddItem(tdbcLibro, sqlcombos)
    '
    
    If registros <= 0 Then
        Mensajes "Cree el libro diario y configurelo en Parametros iniciales", vbOKOnly + vbInformation
    End If
    
    
    DoEvents
    
    If Mid(gsGrupo, 3, 1) <> "1" And gsGrupo <> gsPrivilegioAdmin Then
        Me.cmdGenerar.Enabled = False
    Else
        Me.cmdGenerar.Enabled = True
    End If
    
    tdbcLibro.Bookmark = 0

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

Private Sub tdbcLibro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then pSendKeys "{tab}"
End Sub

