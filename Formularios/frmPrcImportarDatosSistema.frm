VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmPrcImportarDatosSistema 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importar Datos del Sistema"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10890
   Icon            =   "frmPrcImportarDatosSistema.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6045
   ScaleWidth      =   10890
   Begin VB.Frame fraTodo 
      Height          =   6045
      Left            =   45
      TabIndex        =   0
      Top             =   -45
      Width           =   10770
      Begin VB.Frame fraDirectorios 
         Caption         =   "DIRECTORIO DE ORIGEN DE ARCHIVOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5280
         Left            =   135
         TabIndex        =   37
         Top             =   135
         Width           =   3705
         Begin VB.DirListBox Dir1 
            Height          =   2115
            Left            =   225
            TabIndex        =   42
            Top             =   645
            Width           =   3345
         End
         Begin VB.DriveListBox Drive1 
            Height          =   315
            Left            =   225
            TabIndex        =   41
            Top             =   270
            Width           =   2850
         End
         Begin VB.FileListBox File1 
            Height          =   675
            Left            =   225
            Pattern         =   "*.exp"
            TabIndex        =   40
            Top             =   2880
            Width           =   3345
         End
         Begin VB.CheckBox chkEliminar 
            Caption         =   "Eliminar los datos de la empresa"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   135
            TabIndex        =   39
            Top             =   4410
            Width           =   3210
         End
         Begin VB.CheckBox chkEliminarTodosAnios 
            Caption         =   "Eliminar todos los años de la empresa"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   135
            TabIndex        =   38
            Top             =   4680
            Width           =   3525
         End
         Begin TDBText6Ctl.TDBText tdbtDirectorio 
            Height          =   780
            Left            =   225
            TabIndex        =   43
            Top             =   3600
            Width           =   3345
            _Version        =   65536
            _ExtentX        =   5900
            _ExtentY        =   1376
            Caption         =   "frmPrcImportarDatosSistema.frx":0ECA
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmPrcImportarDatosSistema.frx":0F36
            Key             =   "frmPrcImportarDatosSistema.frx":0F54
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
         Begin TrueOleDBList70.TDBCombo tdbcSucursal 
            Height          =   300
            Left            =   765
            TabIndex        =   44
            Tag             =   "_"
            Top             =   1620
            Visible         =   0   'False
            Width           =   2550
            _ExtentX        =   4498
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
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).DataField=   ""
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   4
            Splits(0)._UserFlags=   0
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).AllowRowSizing=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=4"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=609"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=529"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=370"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=291"
            Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=847"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=767"
            Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(15)=   "Column(2).Visible=0"
            Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(17)=   "Column(3).Width=1376"
            Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=1296"
            Splits(0)._ColumnProps(20)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(21)=   "Column(3).Visible=0"
            Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
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
            Locked          =   -1  'True
            ScrollTrack     =   0   'False
            RowDividerColor =   12632256
            RowSubDividerColor=   12632256
            AddItemSeparator=   ";"
            _PropDict       =   $"frmPrcImportarDatosSistema.frx":0F98
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
            _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
            _StyleDefs(48)  =   "Named:id=33:Normal"
            _StyleDefs(49)  =   ":id=33,.parent=0"
            _StyleDefs(50)  =   "Named:id=34:Heading"
            _StyleDefs(51)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(52)  =   ":id=34,.wraptext=-1"
            _StyleDefs(53)  =   "Named:id=35:Footing"
            _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(55)  =   "Named:id=36:Selected"
            _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(57)  =   "Named:id=37:Caption"
            _StyleDefs(58)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(59)  =   "Named:id=38:HighlightRow"
            _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(61)  =   "Named:id=39:EvenRow"
            _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(63)  =   "Named:id=40:OddRow"
            _StyleDefs(64)  =   ":id=40,.parent=33"
            _StyleDefs(65)  =   "Named:id=41:RecordSelector"
            _StyleDefs(66)  =   ":id=41,.parent=34"
            _StyleDefs(67)  =   "Named:id=42:FilterBar"
            _StyleDefs(68)  =   ":id=42,.parent=33"
         End
         Begin MSForms.CommandButton cmdRefresh 
            Height          =   345
            Left            =   3150
            TabIndex        =   45
            ToolTipText     =   "Cargar Lista"
            Top             =   270
            Width           =   405
            PicturePosition =   327683
            Size            =   "714;609"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
      End
      Begin VB.Frame fraOpciones 
         Caption         =   "DATOS A IMPORTAR"
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
         Left            =   3945
         TabIndex        =   6
         Top             =   135
         Width           =   6720
         Begin VB.CheckBox chkMercaderias 
            Caption         =   "Mercaderias"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   4590
            TabIndex        =   48
            Top             =   3150
            Value           =   1  'Checked
            Width           =   1635
         End
         Begin VB.CheckBox chkMovim 
            Caption         =   "Movimientos"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   4590
            TabIndex        =   31
            Top             =   630
            Value           =   1  'Checked
            Width           =   1635
         End
         Begin VB.CheckBox chkPDB 
            Caption         =   "PDB"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   4590
            TabIndex        =   30
            Top             =   2205
            Value           =   1  'Checked
            Width           =   1635
         End
         Begin VB.CheckBox chkCapital 
            Caption         =   "Capital"
            CausesValidation=   0   'False
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   4590
            TabIndex        =   29
            Top             =   1890
            Value           =   1  'Checked
            Width           =   1635
         End
         Begin VB.CheckBox chkPresup 
            Caption         =   "Presupuesto"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   4590
            TabIndex        =   28
            Top             =   1575
            Value           =   1  'Checked
            Width           =   1635
         End
         Begin VB.CheckBox chkFlujo 
            Caption         =   "Flujo de Efectivo"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   4590
            TabIndex        =   27
            Top             =   1260
            Value           =   1  'Checked
            Width           =   1635
         End
         Begin VB.CheckBox chkPatrimonio 
            Caption         =   "Patrimonio Neto"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   4590
            TabIndex        =   26
            Top             =   945
            Value           =   1  'Checked
            Width           =   1635
         End
         Begin VB.CheckBox chkTipoCbio 
            Caption         =   "Tipo de cambio"
            Height          =   285
            Left            =   180
            TabIndex        =   25
            Top             =   3060
            Value           =   1  'Checked
            Width           =   1500
         End
         Begin VB.Frame Frame4 
            Height          =   3120
            Index           =   1
            Left            =   4365
            TabIndex        =   24
            Top             =   360
            Width           =   60
         End
         Begin VB.CheckBox chkCenCos 
            Caption         =   "Centro de costo"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   2340
            TabIndex        =   23
            Top             =   2700
            Value           =   1  'Checked
            Width           =   1770
         End
         Begin VB.CheckBox chkPlan 
            Caption         =   "Plan de cuentas"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   180
            TabIndex        =   22
            Top             =   630
            Value           =   1  'Checked
            Width           =   1905
         End
         Begin VB.CheckBox chkLibros 
            Caption         =   "Libros"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   180
            TabIndex        =   21
            Top             =   900
            Value           =   1  'Checked
            Width           =   1905
         End
         Begin VB.CheckBox chkEntidades 
            Caption         =   "Entidades"
            Height          =   285
            Left            =   180
            TabIndex        =   20
            Top             =   1170
            Value           =   1  'Checked
            Width           =   1905
         End
         Begin VB.CheckBox chkTipDoc 
            Caption         =   "Tipo de documentos"
            Height          =   285
            Left            =   180
            TabIndex        =   19
            Top             =   1440
            Value           =   1  'Checked
            Width           =   1905
         End
         Begin VB.CheckBox chkTabSec 
            Caption         =   "Tablas secundarias"
            Height          =   285
            Left            =   180
            TabIndex        =   18
            Top             =   2385
            Value           =   1  'Checked
            Width           =   1905
         End
         Begin VB.CheckBox chkTipMon 
            Caption         =   "Tipo de moneda"
            Height          =   285
            Left            =   180
            TabIndex        =   17
            Top             =   2070
            Value           =   1  'Checked
            Width           =   1905
         End
         Begin VB.Frame Frame4 
            Height          =   3120
            Index           =   0
            Left            =   2115
            TabIndex        =   16
            Top             =   360
            Width           =   60
         End
         Begin VB.CheckBox chkTipAsto 
            Caption         =   "Tipo de asiento"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   2340
            TabIndex        =   15
            Top             =   1260
            Value           =   1  'Checked
            Width           =   1905
         End
         Begin VB.CheckBox chkPlanCONA 
            Caption         =   "Plantilla EEFF"
            Height          =   285
            Left            =   180
            TabIndex        =   14
            Top             =   1755
            Value           =   1  'Checked
            Width           =   1905
         End
         Begin VB.CheckBox chkParamIni 
            Caption         =   "Parametros iniciales"
            Height          =   285
            Left            =   2340
            TabIndex        =   13
            Top             =   630
            Value           =   1  'Checked
            Width           =   1905
         End
         Begin VB.CheckBox chkConfigOp 
            Caption         =   "Config. de operaciones"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   2340
            TabIndex        =   12
            Top             =   945
            Value           =   1  'Checked
            Width           =   1995
         End
         Begin VB.CheckBox chkBancos 
            Caption         =   "Bancos"
            Height          =   285
            Left            =   2340
            TabIndex        =   11
            Top             =   2070
            Value           =   1  'Checked
            Width           =   1905
         End
         Begin VB.CheckBox chkCtaCte 
            Caption         =   "Cuenta corriente"
            Height          =   285
            Left            =   2340
            TabIndex        =   10
            Top             =   2385
            Value           =   1  'Checked
            Width           =   1905
         End
         Begin VB.CheckBox chkRatios 
            Caption         =   "Ratios"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   2340
            TabIndex        =   9
            Top             =   3015
            Value           =   1  'Checked
            Width           =   1905
         End
         Begin VB.CheckBox chkCostos 
            Caption         =   "Costos Producción"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   4590
            TabIndex        =   8
            Top             =   2835
            Value           =   1  'Checked
            Width           =   1635
         End
         Begin VB.CheckBox chkValores 
            Caption         =   "Valores"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   4590
            TabIndex        =   7
            Top             =   2520
            Value           =   1  'Checked
            Width           =   1635
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
            TabIndex        =   36
            Top             =   2790
            Width           =   1530
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
            TabIndex        =   35
            Top             =   360
            Width           =   1230
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
            TabIndex        =   34
            Top             =   315
            Width           =   1140
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
            TabIndex        =   33
            Top             =   360
            Width           =   1290
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
            TabIndex        =   32
            Top             =   1710
            Width           =   1125
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " PROCESO DE IMPORTACION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1380
         Index           =   0
         Left            =   3945
         TabIndex        =   1
         Top             =   4035
         Width           =   6735
         Begin MSComctlLib.ProgressBar pgbAvance 
            Height          =   195
            Left            =   180
            TabIndex        =   2
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
            TabIndex        =   3
            Top             =   1080
            Width           =   6360
            _ExtentX        =   11218
            _ExtentY        =   344
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
            Scrolling       =   1
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
            TabIndex        =   5
            Top             =   765
            Width           =   6360
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
            TabIndex        =   4
            Top             =   180
            Width           =   6360
         End
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
         Left            =   7920
         TabIndex        =   50
         Top             =   3780
         Width           =   1260
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
         Left            =   9450
         TabIndex        =   49
         Top             =   3780
         Width           =   1185
      End
      Begin MSForms.CommandButton cmdImportarDatos 
         Height          =   435
         Left            =   4410
         TabIndex        =   46
         Top             =   5475
         Width           =   2115
         Caption         =   "   Importar Datos"
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
      TabIndex        =   47
      Top             =   0
      Width           =   4365
   End
End
Attribute VB_Name = "frmPrcImportarDatosSistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lArrVou As New XArrayDB
Dim lArrMnt() As Variant        ' *** Arreglo para los mantenimientos
Dim lArrMnt2() As Variant        ' *** Arreglo para los mantenimientos 2
Dim gsGrupo As String
Dim ProcesoTerminado As Boolean

Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

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
'        chkPlanCONA.Value = vbChecked
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
        
        chkEliminar.Value = vbChecked
        chkEliminar.Enabled = False
        
        
    Else
        
        chkEliminar.Enabled = True
        
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

Private Sub cmdRefresh_Click()
    Drive1.Refresh
    Dir1.Refresh
    File1.Refresh
End Sub

Private Sub Dir1_Change()
    tdbtDirectorio.Text = Dir1.Path
    File1.Path = Dir1.Path   ' Establece la ruta del archivo.
End Sub

Private Sub Drive1_Change()
    On Error GoTo serror
    Dir1.Path = Drive1.Drive
    Dir1.Refresh
    tdbtDirectorio.Text = Dir1.Path
    Exit Sub
serror:
    Mensajes Err.Description
End Sub

Private Sub PGBAvanceConfigura(inicio As Long, Fin As Long, Descripcion As String)
On Error GoTo serror
    pgbAvance.Min = inicio
    pgbAvance.Max = Fin
    pgbAvance.Value = 0
    lblAvance.Caption = Descripcion
    Exit Sub
serror:
    
End Sub

Private Sub PGBAvanceActualiza(Valor As Long, Descripcion As String)
On Error GoTo serror
    pgbAvance.Value = Valor
    'pgbAvance.Refresh
    lblAvance.Caption = "Tabla : " & Descripcion
    'lblAvance.Refresh
    DoEvents
    Exit Sub
serror:
    
End Sub

Private Function EliminaDatosactualesDeEmpresa() As Boolean
    EliminaDatosactualesDeEmpresa = False
    Dim respuesta As String
    
    respuesta = MsgBox("Se ELIMINARA TODA LA INFORMACION existente de " & Salto(1) & "EMPRESA:  " & gsEmpresaNom & " AÑO " & gsAnio & Salto(2) & "Para reemplazala con los datos con la importación, desea continuar", vbYesNo + vbQuestion, "Confirmar Eliminacion de Datos")
    
    If respuesta = vbYes Then
        CargaArregloReplica
        Dim clsMante As New clsMantoTablas
        Screen.MousePointer = vbHourglass
        
        gsImportacion = True
                
        ReDim lArrMnt2(2) As Variant
        lArrMnt2(0) = gsEmpresa           ' Empresa
        lArrMnt2(1) = gsAnio              ' Anio
            
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_UpdPeriodoPlantillaEEFF ", lArrMnt(), True) = False Then
         Debug.Print "No se actualizo..."
        End If
        
        Set clsMante = New clsMantoTablas
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_EliminarTablasImportacion", lArrMnt(), True) = False Then
            Screen.MousePointer = vbNormal
            EliminaDatosactualesDeEmpresa = False
            Set clsMante = Nothing
            gsImportacion = False
            Exit Function
        Else
            Screen.MousePointer = vbNormal
            'Mensajes "Se eliminaron los datos con exito, se iniciará la importacion de datos .", vbOKOnly + vbInformation
            EliminaDatosactualesDeEmpresa = True
            Set clsMante = Nothing
        End If
    End If
    
    If respuesta = vbNo Then
       EliminaDatosactualesDeEmpresa = False
    End If
    
    gsImportacion = False
End Function

Private Sub ErrorImp(ByVal cCadena As String)
    Mensajes "Error de importación de datos -> " & cCadena
    DoEvents
    cmdImportarDatos.Enabled = True
    Screen.MousePointer = vbNormal
End Sub

Private Sub cmdImportarDatos_Click()
    fraDirectorios.Enabled = False
    fraOpciones.Enabled = False
    DoEvents
    
    Call IniciarImportacion
    
    DoEvents
    fraDirectorios.Enabled = True
    fraOpciones.Enabled = True

End Sub

Private Sub IniciarImportacion()
    If File1.ListCount < 1 Then
        Mensajes "Seleccione una carpeta con los archivos de importacion", vbOKOnly + vbInformation
        Exit Sub
    End If
    
    Mensajes "ANTES DE REALIZAR ESTE PROCESO SAQUE UN BACKUP TOTAL DEL SISTEMA COMO CONTINGENCIA", vbOKOnly + vbExclamation
    
    Dim respuesta As String
    Dim ruta As String
    Dim archivo As String
    
    respuesta = MsgBox("Desea IMPORTAR la información seleccionada a la " & Salto(2) & "ACTUAL EMPRESA: " & gsEmpresaNom & " AÑO " & gsAnio, vbYesNo + vbQuestion, "Confirmar Importar Datos")
    If respuesta = vbNo Then Exit Sub
    ruta = Trim(tdbtDirectorio)
    If Not Right(ruta, 1) = "\" Then ruta = ruta & "\"
    '------------------------ Configurar progess bar ------------------------
    pgbAvanceTotal.Min = 0
    pgbAvanceTotal.Max = 37
    pgbAvanceTotal.Value = 0
    lblAvanceTotal.Caption = "Iniciando proceso ..."
    Call EscribirLog("Iniciando importacion de datos de la empresa " & gsEmpresa & " " & gsEmpresaNom, Me.Name)
    
    If chkEliminar.Value = vbChecked Then
        If EliminaDatosactualesDeEmpresa = False Then
           lblAvanceTotal.Caption = ""
           Exit Sub
        End If
    End If
    
    lblAvanceTotal.Caption = ""
    cmdImportarDatos.Enabled = False
    DoEvents
    Screen.MousePointer = vbHourglass
    '------------------------ Leer Archivos dat para el Plan de Cuentas ------------------------
    archivo = ruta & "CCuentas.exp"
    If ExisteArchivo(archivo) = True And chkPlan.Value = vbChecked Then
        pgbAvanceTotal.Value = pgbAvanceTotal.Value + 1
        pgbAvanceTotal.Refresh
        lblAvanceTotal.Caption = "Importando -> " & archivo
        lblAvanceTotal.Refresh
        DoEvents
        If importaCuentas(archivo) = False Then Call ErrorImp(archivo): Exit Sub

        Me.File1.Refresh
    End If

    archivo = ruta & "CCuentasDet.exp"
    If ExisteArchivo(archivo) = True And chkPlan.Value = vbChecked Then
        pgbAvanceTotal.Value = pgbAvanceTotal.Value + 1
        pgbAvanceTotal.Refresh
        lblAvanceTotal.Caption = "Importando -> " & archivo
        lblAvanceTotal.Refresh
        DoEvents
        If importaCuentasDet(archivo) = False Then Call ErrorImp(archivo): Exit Sub
        
        Me.File1.Refresh
    End If

    '------------------------ Leer Archivos dat para los libros ------------------------
    archivo = ruta & "CLibros.exp"
    If ExisteArchivo(archivo) = True And chkLibros.Value = vbChecked Then
        pgbAvanceTotal.Value = pgbAvanceTotal.Value + 1
        pgbAvanceTotal.Refresh
        lblAvanceTotal.Caption = "Importando -> " & archivo
        lblAvanceTotal.Refresh
        DoEvents
        If Importar("CLibros", "Libros", archivo) = False Then Call ErrorImp(archivo): Exit Sub
        Me.File1.Refresh
    End If


    '------------------------ Leer Archivos dat para Tipos documentos------------------------
    archivo = ruta & "CDocumentos.exp"
    If ExisteArchivo(archivo) = True And chkTipDoc.Value = vbChecked Then
        pgbAvanceTotal.Value = pgbAvanceTotal.Value + 1
        pgbAvanceTotal.Refresh
        lblAvanceTotal.Caption = "Importando -> " & archivo
        lblAvanceTotal.Refresh
        DoEvents
        Importar "CDocumentos", "Documentos", archivo
        Me.File1.Refresh
    End If

    '------------------------ Leer Archivos dat para plantillas CONASEV ------------------------
    archivo = ruta & "CConasev.exp"
    If ExisteArchivo(archivo) = True And chkPlanCONA.Value = vbChecked Then
        pgbAvanceTotal.Value = pgbAvanceTotal.Value + 1
        pgbAvanceTotal.Refresh
        lblAvanceTotal.Caption = "Importando -> " & archivo
        lblAvanceTotal.Refresh
        DoEvents
        If Importar("CConasev", "Conasev", archivo) = False Then
            Call ErrorImp(archivo)
            Exit Sub
        End If
        Me.File1.Refresh
    End If

    archivo = ruta & "CConasev2.exp"
    If ExisteArchivo(archivo) = True And chkPlanCONA.Value = vbChecked Then
        pgbAvanceTotal.Value = pgbAvanceTotal.Value + 1
        pgbAvanceTotal.Refresh
        lblAvanceTotal.Caption = "Importando -> " & archivo
        lblAvanceTotal.Refresh
        DoEvents
        If Importar("CConasev2", "Conasev", archivo) = False Then
            Call ErrorImp(archivo)
            Exit Sub
        End If
        Me.File1.Refresh
    End If
    '------------------------ Leer Archivos dat para tipo de moneda ------------------------
    archivo = ruta & "CMoneda.exp"
    If ExisteArchivo(archivo) = True And chkTipMon.Value = vbChecked Then
        pgbAvanceTotal.Value = pgbAvanceTotal.Value + 1
        pgbAvanceTotal.Refresh
        lblAvanceTotal.Caption = "Importando -> " & archivo
        lblAvanceTotal.Refresh
        DoEvents
        If Importar("CMoneda", "Moneda", archivo) = False Then
            Call ErrorImp(archivo)
            Exit Sub
        End If
        Me.File1.Refresh
    End If

    '------------------------ Leer Archivos dat para tablas secuendarias ------------------------
     archivo = ruta & "CSec.exp"
    If ExisteArchivo(archivo) = True And chkTabSec.Value = vbChecked Then
        pgbAvanceTotal.Value = pgbAvanceTotal.Value + 1
        pgbAvanceTotal.Refresh
        lblAvanceTotal.Caption = "Importando -> " & archivo
        lblAvanceTotal.Refresh
        DoEvents
        If Importar("CSec", "Tab. Secundarias", archivo) = False Then
            Call ErrorImp(archivo)
            Exit Sub
        End If
        Me.File1.Refresh
    End If

    '------------------------ Leer Archivos dat para parametros iniciales ------------------------
    archivo = ruta & "CParam.exp"
    If ExisteArchivo(archivo) = True And chkParamIni.Value = vbChecked Then
        pgbAvanceTotal.Value = pgbAvanceTotal.Value + 1
        pgbAvanceTotal.Refresh
        lblAvanceTotal.Caption = "Importando -> " & archivo
        lblAvanceTotal.Refresh
        DoEvents
        If Importar("CParam", "Parametros Iniciales", archivo) = False Then
            Call ErrorImp(archivo)
            Exit Sub
        End If
        Me.File1.Refresh
    End If

    archivo = ruta & "CParam2.exp"
    If ExisteArchivo(archivo) = True And chkParamIni.Value = vbChecked Then
        pgbAvanceTotal.Value = pgbAvanceTotal.Value + 1
        pgbAvanceTotal.Refresh
        lblAvanceTotal.Caption = "Importando -> " & archivo
        lblAvanceTotal.Refresh
        DoEvents
        If Importar("CParam2", "Parametros Iniciales", archivo) = False Then
            Call ErrorImp(archivo)
            Exit Sub
        End If
       
        Me.File1.Refresh
    End If

    If chkTabSec.Value = vbUnchecked Then
        Call ActualizaTablaParamIni
    End If
    
    '------------------------ Leer Archivos dat para config operaciones ------------------------
    archivo = ruta & "COpera.exp"
    If ExisteArchivo(archivo) = True And chkConfigOp.Value = vbChecked Then
        pgbAvanceTotal.Value = pgbAvanceTotal.Value + 1
        pgbAvanceTotal.Refresh
        lblAvanceTotal.Caption = "Importando -> " & archivo
        lblAvanceTotal.Refresh
        DoEvents
        If Importar("COpera", "Conf. Operaciones", archivo) = False Then
            Call ErrorImp(archivo)
            Exit Sub
        End If
        
        Me.File1.Refresh
    End If

    archivo = ruta & "COpera2.exp"
    If ExisteArchivo(archivo) = True And chkConfigOp.Value = vbChecked Then
        pgbAvanceTotal.Value = pgbAvanceTotal.Value + 1
        pgbAvanceTotal.Refresh
        lblAvanceTotal.Caption = "Importando -> " & archivo
        lblAvanceTotal.Refresh
        DoEvents
        If Importar("COpera2", "Conf. Operaciones", archivo) = False Then
            Call ErrorImp(archivo)
            Exit Sub
        End If
        
        Me.File1.Refresh
    End If

    archivo = ruta & "COpera3.exp"
    If ExisteArchivo(archivo) = True And chkConfigOp.Value = vbChecked Then
        pgbAvanceTotal.Value = pgbAvanceTotal.Value + 1
        pgbAvanceTotal.Refresh
        lblAvanceTotal.Caption = "Importando -> " & archivo
        lblAvanceTotal.Refresh
        DoEvents
        If Importar("COpera3", "Conf. Operaciones", archivo) = False Then
            Call ErrorImp(archivo)
            Exit Sub
        End If
        
        Me.File1.Refresh
    End If

    '------------------------ Leer Archivos dat para tipo de asientos ------------------------
    archivo = ruta & "CTipoasiento.exp"
    If ExisteArchivo(archivo) = True And chkTipAsto.Value = vbChecked Then
        pgbAvanceTotal.Value = pgbAvanceTotal.Value + 1
        pgbAvanceTotal.Refresh
        lblAvanceTotal.Caption = "Importando -> " & archivo
        lblAvanceTotal.Refresh
        DoEvents
        If Importar("CTipoasiento", "Tipos asientos", archivo) = False Then
            Call ErrorImp(archivo)
            Exit Sub
        End If
        
        Me.File1.Refresh
    End If

    '------------------------ Leer Archivos dat para bancos ------------------------
    archivo = ruta & "CBancos.exp"
    If ExisteArchivo(archivo) = True And chkBancos.Value = vbChecked Then
        pgbAvanceTotal.Value = pgbAvanceTotal.Value + 1
        pgbAvanceTotal.Refresh
        lblAvanceTotal.Caption = "Importando -> " & archivo
        lblAvanceTotal.Refresh
        DoEvents
        If Importar("CBancos", "Bancos", archivo) = False Then
            Call ErrorImp(archivo)
            Exit Sub
        End If
        
        Me.File1.Refresh
    End If

    '------------------------ Leer Archivos dat para cta cte ------------------------
    archivo = ruta & "CCtacte.exp"
    If ExisteArchivo(archivo) = True And chkCtaCte.Value = vbChecked Then
        pgbAvanceTotal.Value = pgbAvanceTotal.Value + 1
        pgbAvanceTotal.Refresh
        lblAvanceTotal.Caption = "Importando -> " & archivo
        lblAvanceTotal.Refresh
        DoEvents
        Call Importar("CCtacte", "Cuenta corriente", archivo)

        Me.File1.Refresh
    End If

    '------------------------ Leer Archivos dat para el Centro de Costo ------------------------
    archivo = ruta & "CCosto.exp"
    If ExisteArchivo(archivo) = True And chkCenCos.Value = vbChecked Then
        pgbAvanceTotal.Value = pgbAvanceTotal.Value + 1
        pgbAvanceTotal.Refresh
        lblAvanceTotal.Caption = "Importando -> " & archivo
        lblAvanceTotal.Refresh
        DoEvents
        If Importar("CCosto", "Centro Costo", archivo) = False Then
            Call ErrorImp(archivo)
            Exit Sub
        End If
        
        Me.File1.Refresh
    End If

    '------------------------ Leer Archivos dat para TC ------------------------
    archivo = ruta & "CTc.exp"
    If ExisteArchivo(archivo) = True And chkTipoCbio.Value = vbChecked Then
        pgbAvanceTotal.Value = pgbAvanceTotal.Value + 1
        pgbAvanceTotal.Refresh
        lblAvanceTotal.Caption = "Importando -> " & archivo
        lblAvanceTotal.Refresh
        DoEvents
        If Importar("CTc", "Tipo Cambio", archivo) = False Then
            Call ErrorImp(archivo)
            Exit Sub
        End If
        
        Me.File1.Refresh
    End If

    archivo = ruta & "CTc2.exp"
    If ExisteArchivo(archivo) = True And chkTipoCbio.Value = vbChecked Then
        pgbAvanceTotal.Value = pgbAvanceTotal.Value + 1
        pgbAvanceTotal.Refresh
        lblAvanceTotal.Caption = "Importando -> " & archivo
        lblAvanceTotal.Refresh
        DoEvents
        If Importar("CTc2", "Tipo Cambio", archivo) = False Then
            Call ErrorImp(archivo)
            Exit Sub
        End If
        
        Me.File1.Refresh
    End If
    
    '------------------------ Leer Archivos dat para las Entidades ------------------------
    archivo = ruta & "Entidades.exp"
    If ExisteArchivo(archivo) = True And chkEntidades.Value = vbChecked Then
        pgbAvanceTotal.Value = pgbAvanceTotal.Value + 1
        pgbAvanceTotal.Refresh
        lblAvanceTotal.Caption = "Importando -> " & archivo
        lblAvanceTotal.Refresh
        DoEvents
        If Importar("CEntidades", "Entidades", archivo) = False Then
            Call ErrorImp(archivo)
            Exit Sub
        End If
        
        Me.File1.Refresh
    End If

    archivo = ruta & "Entidades2.exp"
    If ExisteArchivo(archivo) = True And chkEntidades.Value = vbChecked Then
        pgbAvanceTotal.Value = pgbAvanceTotal.Value + 1
        pgbAvanceTotal.Refresh
        lblAvanceTotal.Caption = "Importando -> " & archivo
        lblAvanceTotal.Refresh
        DoEvents
        If Importar("CEntidades2", "Entidades", archivo) = False Then
            Call ErrorImp(archivo)
            Exit Sub
        End If
        
        Me.File1.Refresh
    End If
    
    archivo = ruta & "Entidades3.exp"
    If ExisteArchivo(archivo) = True And chkEntidades.Value = vbChecked Then
        pgbAvanceTotal.Value = pgbAvanceTotal.Value + 1
        pgbAvanceTotal.Refresh
        lblAvanceTotal.Caption = "Importando -> " & archivo
        lblAvanceTotal.Refresh
        DoEvents
        If Importar("CEntidades3", "Entidades", archivo) = False Then
            Call ErrorImp(archivo)
            Exit Sub
        End If
        
        Me.File1.Refresh
    End If
    
    archivo = ruta & "CRatios.exp"
    If ExisteArchivo(archivo) = True And chkRatios.Value = vbChecked Then
        pgbAvanceTotal.Value = pgbAvanceTotal.Value + 1
        pgbAvanceTotal.Refresh
        lblAvanceTotal.Caption = "Importando -> " & archivo
        lblAvanceTotal.Refresh
        DoEvents
        If Importar("CRatios", "Ratios", archivo) = False Then
            Call ErrorImp(archivo)
            Exit Sub
        End If
        
        Me.File1.Refresh
    End If
    
    archivo = ruta & "CPatrim.exp"
    If ExisteArchivo(archivo) = True And chkPatrimonio.Value = vbChecked Then
        pgbAvanceTotal.Value = pgbAvanceTotal.Value + 1
        pgbAvanceTotal.Refresh
        lblAvanceTotal.Caption = "Importando -> " & archivo
        lblAvanceTotal.Refresh
        DoEvents
        If Importar("CPatrim", "Patrimonio", archivo) = False Then
            Call ErrorImp(archivo)
            Exit Sub
        End If
        Me.File1.Refresh
    End If
    
    archivo = ruta & "CFlujoPro.exp"
    If ExisteArchivo(archivo) = True And chkFlujo.Value = vbChecked Then
        pgbAvanceTotal.Value = pgbAvanceTotal.Value + 1
        pgbAvanceTotal.Refresh
        lblAvanceTotal.Caption = "Importando -> " & archivo
        lblAvanceTotal.Refresh
        DoEvents
        If Importar("CFlujoPro", "Flujo de Efectivo - Proceso", archivo) = False Then
            Call ErrorImp(archivo)
            Exit Sub
        End If
        Me.File1.Refresh
    End If
    
    archivo = ruta & "CFlujoRep.exp"
    If ExisteArchivo(archivo) = True And chkFlujo.Value = vbChecked Then
        pgbAvanceTotal.Value = pgbAvanceTotal.Value + 1
        pgbAvanceTotal.Refresh
        lblAvanceTotal.Caption = "Importando -> " & archivo
        lblAvanceTotal.Refresh
        DoEvents
        If Importar("CFlujoRep", "Flujo de Efectivo - Reporte", archivo) = False Then
            Call ErrorImp(archivo)
            Exit Sub
        End If
        Me.File1.Refresh
    End If
    
    archivo = ruta & "CFlujoSal.exp"
    If ExisteArchivo(archivo) = True And chkFlujo.Value = vbChecked Then
        pgbAvanceTotal.Value = pgbAvanceTotal.Value + 1
        pgbAvanceTotal.Refresh
        lblAvanceTotal.Caption = "Importando -> " & archivo
        lblAvanceTotal.Refresh
        DoEvents
        If Importar("CFlujoSal", "Flujo de Efectivo - Saldos", archivo) = False Then
            Call ErrorImp(archivo)
            Exit Sub
        End If
        Me.File1.Refresh
    End If
   
    archivo = ruta & "CFlujoCta.exp"
    If ExisteArchivo(archivo) = True And chkFlujo.Value = vbChecked Then
        pgbAvanceTotal.Value = pgbAvanceTotal.Value + 1
        pgbAvanceTotal.Refresh
        lblAvanceTotal.Caption = "Importando -> " & archivo
        lblAvanceTotal.Refresh
        DoEvents
        If Importar("CFlujoCta", "Flujo de Efectivo - Cuentas", archivo) = False Then
            Call ErrorImp(archivo)
            Exit Sub
        End If
        Me.File1.Refresh
    End If
   
    archivo = ruta & "CPresup.exp"
    If ExisteArchivo(archivo) = True And chkPresup.Value = vbChecked Then
        pgbAvanceTotal.Value = pgbAvanceTotal.Value + 1
        pgbAvanceTotal.Refresh
        lblAvanceTotal.Caption = "Importando -> " & archivo
        lblAvanceTotal.Refresh
        DoEvents
        If Importar("CPresup", "Presupuesto", archivo) = False Then
            Call ErrorImp(archivo)
            Exit Sub
        End If
        Me.File1.Refresh
    End If
   
    archivo = ruta & "CCapital.exp"
    If ExisteArchivo(archivo) = True And chkCapital.Value = vbChecked Then
        pgbAvanceTotal.Value = pgbAvanceTotal.Value + 1
        pgbAvanceTotal.Refresh
        lblAvanceTotal.Caption = "Importando -> " & archivo
        lblAvanceTotal.Refresh
        DoEvents
        If Importar("CCapital", "Capital", archivo) = False Then
            Call ErrorImp(archivo)
            Exit Sub
        End If
        Me.File1.Refresh
    End If
    '------------------ REG AUXILIARES ------------------'
    archivo = ruta & "CRegAux.exp"
    If ExisteArchivo(archivo) = True And chkMovim.Value = vbChecked Then
        pgbAvanceTotal.Value = pgbAvanceTotal.Value + 1
        pgbAvanceTotal.Refresh
        lblAvanceTotal.Caption = "Importando -> " & archivo
        lblAvanceTotal.Refresh
        DoEvents
        If Importar("CRegAux", "Registros Auxiliares", archivo) = False Then
            Call ErrorImp(archivo)
            Exit Sub
        End If
        Me.File1.Refresh
    End If
   
    '------------------ VALORES ------------------'
    archivo = ruta & "CVal.exp"
    If ExisteArchivo(archivo) = True And chkValores.Value = vbChecked Then
        pgbAvanceTotal.Value = pgbAvanceTotal.Value + 1
        pgbAvanceTotal.Refresh
        lblAvanceTotal.Caption = "Importando -> " & archivo
        lblAvanceTotal.Refresh
        DoEvents
        If Importar("CVal", "Valores", archivo) = False Then
            Call ErrorImp(archivo)
            Exit Sub
        End If
        Me.File1.Refresh
    End If
    
    '------------------ MERCADERIAS ------------------'
    archivo = ruta & "CMercaderias.exp"
    If ExisteArchivo(archivo) = True And chkMercaderias.Value = vbChecked Then
        pgbAvanceTotal.Value = pgbAvanceTotal.Value + 1
        pgbAvanceTotal.Refresh
        lblAvanceTotal.Caption = "Importando -> " & archivo
        lblAvanceTotal.Refresh
        DoEvents
        If Importar("CMercaderias", "Mercaderias", archivo) = False Then
            Call ErrorImp(archivo)
            Exit Sub
        End If
        Me.File1.Refresh
    End If
    
    '------------------ COSTOS ------------------'
    archivo = ruta & "CCos.exp"
    If ExisteArchivo(archivo) = True And chkCostos.Value = vbChecked Then
        pgbAvanceTotal.Value = pgbAvanceTotal.Value + 1
        pgbAvanceTotal.Refresh
        lblAvanceTotal.Caption = "Importando -> " & archivo
        lblAvanceTotal.Refresh
        DoEvents
        If Importar("CCos", "Costos", archivo) = False Then
            Call ErrorImp(archivo)
            Exit Sub
        End If
        Me.File1.Refresh
    End If
    
    archivo = ruta & "CCosInv.exp"
    If ExisteArchivo(archivo) = True And chkCostos.Value = vbChecked Then
        pgbAvanceTotal.Value = pgbAvanceTotal.Value + 1
        pgbAvanceTotal.Refresh
        lblAvanceTotal.Caption = "Importando -> " & archivo
        lblAvanceTotal.Refresh
        DoEvents
        If Importar("CCosInv", "Costos Inventarios", archivo) = False Then
            Call ErrorImp(archivo)
            Exit Sub
        End If
        Me.File1.Refresh
    End If
 
    '------------------------ Leer Archivos dat para los Asientos ------------------------
    archivo = ruta & "AsientosC01.exp"
    If ExisteArchivo(ruta & "AsientosC01.exp") = True And chkMovim.Value = vbChecked Then
        pgbAvanceTotal.Value = pgbAvanceTotal.Value + 1
        pgbAvanceTotal.Refresh
        lblAvanceTotal.Caption = "Importando -> " & archivo
        lblAvanceTotal.Refresh
        DoEvents
        If ImportaAsientos(archivo) = False Then
            Call ErrorImp(archivo)
            Exit Sub
        End If
        Me.File1.Refresh
    End If
    
    archivo = ruta & "AsientosD01.exp"
    If ExisteArchivo(ruta & "AsientosD01.exp") = True And chkMovim.Value = vbChecked Then
        pgbAvanceTotal.Value = pgbAvanceTotal.Value + 1
        pgbAvanceTotal.Refresh
        lblAvanceTotal.Caption = "Importando -> " & archivo
        lblAvanceTotal.Refresh
        DoEvents
        If ImportaAsientosDet(archivo) = False Then
            Call ErrorImp(archivo)
            Exit Sub
        End If
        Me.File1.Refresh
    End If
        
    archivo = ruta & "AsientosC02.exp"
    If ExisteArchivo(ruta & "AsientosC02.exp") = True And chkMovim.Value = vbChecked Then
        pgbAvanceTotal.Value = pgbAvanceTotal.Value + 1
        pgbAvanceTotal.Refresh
        lblAvanceTotal.Caption = "Importando -> " & archivo
        lblAvanceTotal.Refresh
        DoEvents
        If ImportaAsientos(archivo) = False Then
            Call ErrorImp(archivo)
            Exit Sub
        End If
        Me.File1.Refresh
    End If
    
    archivo = ruta & "AsientosD02.exp"
    If ExisteArchivo(ruta & "AsientosD02.exp") = True And chkMovim.Value = vbChecked Then
        pgbAvanceTotal.Value = pgbAvanceTotal.Value + 1
        pgbAvanceTotal.Refresh
        lblAvanceTotal.Caption = "Importando -> " & archivo
        lblAvanceTotal.Refresh
        DoEvents
        If ImportaAsientosDet(archivo) = False Then
            Call ErrorImp(archivo)
            Exit Sub
        End If
        Me.File1.Refresh
    End If
        
    '--------------------------------------------------------------------------------------'
    
    archivo = ruta & "CPDB.exp"
    If ExisteArchivo(archivo) = True And chkPDB.Value = vbChecked Then
        pgbAvanceTotal.Value = pgbAvanceTotal.Value + 1
        pgbAvanceTotal.Refresh
        lblAvanceTotal.Caption = "Importando -> " & archivo
        lblAvanceTotal.Refresh
        DoEvents
        If Importar("CPDB", "PDB", archivo) = False Then
            Call ErrorImp(archivo)
            Exit Sub
        End If
        Me.File1.Refresh
    End If
   
'    '------------------------ Leer Archivos dat para CONCIL BANCARIA ------------------------
    archivo = ruta & "CConcil.exp"
    If ExisteArchivo(archivo) = True And chkCtaCte.Value = vbChecked Then
        pgbAvanceTotal.Value = pgbAvanceTotal.Value + 1
        pgbAvanceTotal.Refresh
        lblAvanceTotal.Caption = "Importando -> " & archivo
        lblAvanceTotal.Refresh
        DoEvents
        If Importar("CConcil", "Conciliacion Bancaria", archivo) = False Then
            Call ErrorImp(archivo)
            Exit Sub
        End If
        Me.File1.Refresh
    End If

    
    pgbAvanceTotal.Value = pgbAvanceTotal.Max
    
    lblAvanceTotal.Caption = "Proceso terminado ... "
    lblAvanceTotal.Caption = ""
    lblAvanceTotal.Refresh
    
    Call EscribirLog("Finalizando importación de datos de la empresa " & gsEmpresa & " " & gsEmpresaNom, Me.Name)
    
    Screen.MousePointer = 0
    '*********************
    
    If ProcesoTerminado = True Then
        ActualizaSaldos
    End If
    
    Mensajes "Proceso terminado, exitosamente.", vbOKOnly + vbInformation
    
    DoEvents
    cmdImportarDatos.Enabled = True
    Screen.MousePointer = vbNormal
    
End Sub

Private Sub ActualizaSaldos()
        'Mensajes "SE INICIARA LA ACTUALIZACION DE CUENTAS DE DESTINO", vbOKOnly + vbExclamation
        frmPrcActualizaDestino.Show
        frmPrcActualizaDestino.cmdProcesar.Visible = False
        DoEvents
        frmPrcActualizaDestino.chkMes.Value = vbChecked
        frmPrcActualizaDestino.tdbcMes.BoundText = "14"
        DoEvents
        frmPrcActualizaDestino.gsMensaje = False
        frmPrcActualizaDestino.Procesar
        DoEvents
        frmPrcActualizaDestino.Cerrar
        
        
        
        'Mensajes "SE INICIARA LA ACTUALIZACION DE SALDOS", vbOKOnly + vbExclamation
        frmPrcActualizaSaldos.Show
        frmPrcActualizaSaldos.cmdProcesar.Visible = False
        DoEvents
        frmPrcActualizaSaldos.chkMes.Value = vbChecked
        frmPrcActualizaSaldos.tdbcMes.BoundText = "14"
        DoEvents
        frmPrcActualizaSaldos.gsMensaje = False
        frmPrcActualizaSaldos.Procesar
        DoEvents
        frmPrcActualizaSaldos.Cerrar

End Sub

Private Function ActualizaTablaParamIni() As Boolean

    Dim sql As String
    Dim Moneda As String
    
    sql = "INSERT INTO TABLA ( " & _
          "Emp_cCodigo,Tab_cTabla, Tab_cCodigo, Tab_nLongCod, Tab_cDescripCampo, Tab_cDescripTabla, Tab_nLongitud, " & _
          "Tab_cCodSunat, Tab_cEstado, Tab_cMod01, Tab_cMod02, Tab_cMod03, Tab_cMod04, Tab_cMod05, Tab_cMod06, " & _
          "Tab_cMod07, Tab_cMod08, Tab_cMod09, Tab_cMod10, Tab_cDeleted, Tab_cUserCrea, Tab_dFechaCrea, " & _
          "Tab_cUserModifica , Tab_dFechaModifica, Tab_cEquipoUser) " & _
          "SELECT '" & gsEmpresa & "', " & _
          "Tab_cTabla, Tab_cCodigo, Tab_nLongCod, Tab_cDescripCampo, Tab_cDescripTabla, Tab_nLongitud, " & _
          "Tab_cCodSunat, Tab_cEstado, Tab_cMod01, Tab_cMod02, Tab_cMod03, Tab_cMod04, Tab_cMod05, Tab_cMod06, " & _
          "Tab_cMod07, Tab_cMod08, Tab_cMod09, Tab_cMod10, Tab_cDeleted, Tab_cUserCrea, Tab_dFechaCrea, " & _
          "Tab_cUserModifica , Tab_dFechaModifica, Tab_cEquipoUser " & _
          "From tabla WHERE EMP_CCODIGO='000' "
    
    If EjecutaQuery(sql) < 1 Then
        'Mensajes "El proceso no ha concluido. Verificar...", vbInformation
        ActualizaTablaParamIni = False
        Exit Function
    End If

    ActualizaTablaParamIni = True
    
End Function

Private Function importaCuentas(tabla As String) As Boolean
    Dim rsExp As New ADODB.Recordset
    Dim sqlExp As String
    Dim j As Integer

    On Error GoTo MsgError
    rsExp.Open tabla
    PGBAvanceConfigura 0, rsExp.RecordCount, "Iniciando proceso ..."

    Call Conectar
    Call gcnSistema.Execute("DELETE FROM CND_CUENTA_DIST WHERE Pan_cAnio = '" & gsAnio & "' And Emp_cCodigo = '" & gsEmpresa & "'")
    Call gcnSistema.Execute("DELETE FROM CNM_PLAN_CTA WHERE Emp_cCodigo + Pan_cAnio + Pla_cCuentaContable In(Select Emp_cCodigo + Pan_cAnio + Pla_cCuentaContable From CNM_PLAN_CTA WHERE Pan_cAnio = '" & gsAnio & "' And Emp_cCodigo = '" & gsEmpresa & "')")
    Do While Not rsExp.EOF
        If ExisteCtaImp(rsExp("Pla_cCuentaContable").Value) = False Then
            sqlExp = "set dateformat dmy " & vbCrLf & _
                    "INSERT INTO CNM_PLAN_CTA(Emp_cCodigo, Pan_cAnio, Pla_cCuentaContable, Pla_cNombreCuenta, Pla_cTitulo, " & _
                     "Ten_cTipoEntidad, Pla_cCentroCosto, Pla_cProvision, Pla_cDifCambio, Pla_cOperaTC, " & _
                     "Pla_cRedondeo, Pla_cDocumento, Pla_cTipoCta, Pla_cCptoBG, Pla_cCptoBGDual, Pla_cCptoResFun, " & _
                     "Pla_cCptoResNat, Pla_cTipoAfect, Pla_cDetraccion, Pla_cRetencion, Pla_cPercepcion, " & _
                     "Pla_cCtaPresup, Pla_cEstado, Pla_cDeleted, Pla_cUserCrea, Pla_dFechaCrea, Pla_cUserModifica, " & _
                     "Pla_dFechaModifica , Pla_cEquipoUser, Pla_cNCND, Pla_cConsPDT, Pla_dEstadoO, Pla_dEstadoD, Pla_cCuentaCosVenta, Pla_cVariacionProduccion, Pla_cCostoProduccion " & _
                     ") VALUES( "
                     '
                                        ' , pla_OrdCentCom, pla_OrdCentVta "Pla_dFechaModifica , Pla_cEquipoUser, Pla_cNCND "
            For j = 0 To rsExp.Fields.Count - 2
                If CE(rsExp.Fields(j).Name) = "Emp_cCodigo" Then
                    sqlExp = sqlExp & " '" & gsEmpresa & "', "
                ElseIf CE(rsExp.Fields(j).Name) = "Pan_cAnio" Then
                    sqlExp = sqlExp & " '" & gsAnio & "', "
                ElseIf CE(rsExp.Fields(j).Name) = "Pla_cNombreCuenta" Then
                    sqlExp = sqlExp & " '" & CC(rsExp(rsExp.Fields(j).Name).Value) & "', "
                Else
                    If j = rsExp.Fields.Count - 2 Then
                        sqlExp = sqlExp + " '" & CE(rsExp(rsExp.Fields(j).Name).Value) & "') "
                    Else
                        If j = 25 Or j = 27 Then 'FECHAS
                            sqlExp = sqlExp & " '" & CE(Format(rsExp(rsExp.Fields(j).Name).Value, "dd/MM/yyyy")) & "', "
                        Else
                            sqlExp = sqlExp & " '" & CE(rsExp(rsExp.Fields(j).Name).Value) & "', "
                        End If
                    End If
                End If
            Next
            gcnSistema.Execute sqlExp
        End If

        PGBAvanceActualiza rsExp.AbsolutePosition, "Plan de cuentas"
        rsExp.MoveNext
    Loop
    Call Desconectar

    CerrarRecordSet rsExp
    'EliminaArchivo tabla
    PGBAvanceConfigura 0, 1, ""
    importaCuentas = True
    Exit Function
MsgError:
MsgBox Err.Description
    Call Desconectar
    PGBAvanceConfigura 0, 1, ""
    CerrarRecordSet rsExp
    Mensajes Err.Description, vbOKOnly + vbInformation
    importaCuentas = False
End Function

Private Function importaCuentasDet(tabla As String) As Boolean
    Dim rsExp As New ADODB.Recordset
    Dim sqlExp As String
    Dim j As Integer
    On Error GoTo MsgError
    rsExp.Open tabla

    PGBAvanceConfigura 0, rsExp.RecordCount, "Iniciando proceso ..."
    Call Conectar
    Do While Not rsExp.EOF
        ' *** Eliminando
        sqlExp = "DELETE CND_CUENTA_DIST " & _
                 " WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' " & _
                 " AND Pla_cCuentaContable = '" & rsExp("Pla_cCuentaContable").Value & "' AND Per_cPeriodo = '" & rsExp("Per_cPeriodo").Value & "' " & _
                 " AND Dis_cSecuencia = '" & rsExp("Dis_cSecuencia").Value & "'"
        gcnSistema.Execute sqlExp

        ' *** Insertando
        sqlExp = "set dateformat dmy " & vbCrLf & _
                 "INSERT INTO CND_CUENTA_DIST (Emp_cCodigo, Pan_cAnio, Per_cPeriodo, Pla_cCuentaContable, Dis_cSecuencia, " & _
                 "Dis_cDestinoDebe, Dis_cDestinoHaber, Dis_nPorcentaje, Dis_cEstado, " & _
                 "Dis_cDeleted, Dis_cUserCrea, Dis_dFechaCrea, Dis_cUserModifica, " & _
                 "Dis_dFechaModifica , Dis_cUserEquipo" & _
                 ")VALUES( "

        For j = 0 To rsExp.Fields.Count - 2
            If CE(rsExp.Fields(j).Name) = "Emp_cCodigo" Then
                sqlExp = sqlExp & " '" & gsEmpresa & "', "
            ElseIf CE(rsExp.Fields(j).Name) = "Pan_cAnio" Then
                sqlExp = sqlExp & " '" & gsAnio & "', "
            Else
                If j = rsExp.Fields.Count - 2 Then
                    sqlExp = sqlExp + " '" & CE(rsExp(rsExp.Fields(j).Name).Value) & "') "
                Else
                        If j = 11 Or j = 13 Then 'FECHAS
                            sqlExp = sqlExp & " '" & CE(Format(rsExp(rsExp.Fields(j).Name).Value, "dd/MM/yyyy")) & "', "
                        Else
                            sqlExp = sqlExp & " '" & CE(rsExp(rsExp.Fields(j).Name).Value) & "', "
                        End If
                End If
            End If
        Next
        On Error Resume Next
        gcnSistema.Execute sqlExp
        PGBAvanceActualiza rsExp.AbsolutePosition, "Plan de cuentas destino"
        rsExp.MoveNext
    Loop
    Call Desconectar
    CerrarRecordSet rsExp
    'EliminaArchivo tabla
    PGBAvanceConfigura 0, 1, ""
    importaCuentasDet = True
    Exit Function
MsgError:
    MsgBox (sqlExp)
    Call Desconectar
    PGBAvanceConfigura 0, 1, ""
    CerrarRecordSet rsExp
    Mensajes Err.Description, vbOKOnly + vbInformation
    importaCuentasDet = False

End Function

Private Function ImportaAsientos(tabla As String) As Boolean
    Dim rsExp As New ADODB.Recordset
    Dim rsExpDet As New ADODB.Recordset
    Dim sqlExp As String
    Dim auxInt As String
    Dim auxVou As String
    Dim lArrCab() As Variant    ' *** Variable de arreglo
    Dim auxEmp As String
    Dim auxAnio As String
    Dim auxPer As String
    Dim auxLib As String
    Dim auxUsu As String
    Dim i As Long
    Dim dFechaCrea As String
    Dim dFechaModif As String
    Dim clsMante As clsMantoTablas
    
    On Error GoTo MsgError
    rsExp.Open tabla

    PGBAvanceConfigura 0, rsExp.RecordCount, "Iniciando proceso ..."
    i = 0
    'lArrVou.Clear
    'lArrVou.ReDim 0, 0, 0, 3
    Call Conectar
    Do While Not rsExp.EOF
        ' *** Verificar q no haya sido registrado anteriormente
        'If ExisteAsientoSuc("SEL_REG", gsEmpresa, gsAnio, CE(rsExp("Ase_nVoucher")), _
        '    rsExp("Per_cPeriodo").Value, rsExp("Lib_cTipoLibro").Value, Me.tdbcSucursal.BoundText) = False Then
            
            ' *** Hallar el numero interno y el numero de voucher
            'lArrVou.ReDim 0, i + 1, 0, 3
            
            auxInt = CE(rsExp("Ase_cNummov").Value) 'numeroVoucher("INTERNO", gsAnio, "", "")
            auxVou = CE(rsExp("Ase_nVoucher").Value) 'numeroVoucher("VOUCHER", gsAnio, rsExp("Per_cPeriodo").Value, rsExp("Lib_cTipoLibro").Value)
            
            '-----------------------------'
            'lArrVou(i, 0) = auxInt
            'lArrVou(i, 1) = auxVou
            'lArrVou(i, 2) = CE(rsExp("Ase_cNummov").Value)
            'lArrVou(i, 3) = CE(rsExp("Ase_nVoucher").Value)
            '-----------------------------'
            
            'Debug.Print lArrVou(i, 0) & " - " & lArrVou(i, 1) & " - " & lArrVou(i, 2) & " - " & lArrVou(i, 3)
            dFechaCrea = Left(CE(rsExp("Ase_dFechaCrea").Value), 19)
            dFechaModif = Left(CE(rsExp("Ase_dFechaModifica").Value), 19)
            ' *** Grabando la cabecera
            sqlExp = "set dateformat dmy " & vbCrLf & _
                     "INSERT INTO CNC_ASIENTO_VOUCHER " & _
                     "([Ase_cNummov], [Emp_cCodigo], [Pan_cAnio], [Per_cPeriodo], [Lib_cTipoLibro], [Ase_nVoucher], " & _
                     "[Ase_dFecha], [Ase_cGlosa], [Ase_cTipoMoneda], [Ase_nTipoCambio], " & _
                     "[Ase_cNumMovTra], [Ase_cCodSucursal], [Ase_cOperaTC], [Ase_cOperaCaja],[Ase_cEstado], " & _
                     "[Ase_cDeleted], [Ase_cUserCrea], [Ase_dFechaCrea], [Ase_cUserModifica], [Ase_dFechaModifica],  " & _
                     "[Ase_cEquipoUser], [Ase_cCuadreManual],[Asd_cEstadoO], [Asd_cEstadoD], [CreditoFiscal], [MaterialConstruccion]) " & _
                     "VALUES ( '" & auxInt & "', '" & gsEmpresa & "', '" & gsAnio & "', " & _
                     "'" & rsExp("Per_cPeriodo").Value & "', '" & rsExp("Lib_cTipoLibro").Value & "', '" & auxVou & "', " & _
                     " cast(convert(varchar(10), '" & rsExp("Ase_dFecha").Value & "') as DateTime), '" & CC(rsExp("Ase_cGlosa").Value) & "', " & _
                     "'" & rsExp("Ase_cTipoMoneda").Value & "', '" & rsExp("Ase_nTipoCambio").Value & "', " & _
                     "'" & auxInt & "', '" & rsExp("Ase_cCodSucursal").Value & "', " & _
                     "'" & rsExp("Ase_cOperaTC").Value & "', '" & rsExp("Ase_cOperaCaja").Value & "', " & _
                     "'" & rsExp("Ase_cEstado").Value & "', '" & rsExp("Ase_cDeleted").Value & "', " & _
                     "'" & rsExp("Ase_cUserCrea").Value & "',cast(convert(varchar(10),'" & dFechaCrea & "') as DateTime), " & _
                     "'" & rsExp("Ase_cUserModifica").Value & "',cast(convert(varchar(10), '" & dFechaModif & "') as DateTime), '" & rsExp("Ase_cEquipoUser").Value & "'," & _
                     "'" & rsExp("Ase_cCuadreManual").Value & "', '" & rsExp("Asd_cEstadoO").Value & "', '" & rsExp("Asd_cEstadoD").Value & "', '" & IIf(RTrim(LTrim(rsExp("CreditoFiscal").Value)) = vbNullString, 0, rsExp("CreditoFiscal").Value) & "','" & IIf(RTrim(LTrim(rsExp("MaterialConstruccion").Value)) = vbNullString, 0, rsExp("MaterialConstruccion").Value) & "') "
            
            Debug.Print sqlExp
            
            gcnSistema.Execute sqlExp
            
        'End If
        
        PGBAvanceActualiza rsExp.AbsolutePosition, "Cabecera de Voucher"
        rsExp.MoveNext
        i = i + 1
    Loop
    Call Desconectar
    CerrarRecordSet rsExp
    'EliminaArchivo tabla
    PGBAvanceConfigura 0, 1, ""
    ImportaAsientos = True
    Exit Function

MsgError:
    Call Desconectar
    CerrarRecordSet rsExp
    Mensajes Err.Description, vbOKOnly + vbInformation
    ImportaAsientos = False
End Function

'********************************************************************
' DESCRIPCION : Cambia todas las comillas simples por su secuencia de escape
'               para evitar que la consulta se corte inesperadamente
' PARAMETROS :  Cadena  = cadena a cambiar
'               Extremos= Si debe devolver los extremos con las comillas simples sin cambiar
Public Function CC(ByVal cadena As String, Optional ByVal Extremos As Boolean = False)
    ' Si la cadena debe devolver los extremos con comillas
    If Extremos Then cadena = Mid(cadena, 2, Len(cadena) - 2)
    ' Elimina todos los espacios en blanco de los extremos
    cadena = Trim(cadena)
    ' Cambia todas las comillas simples por la Secuencia de Escape
    cadena = Replace(cadena, "'", "''")
    CC = IIf(Extremos, "'" & cadena & "'", cadena)
End Function

Private Sub RetornaNumMov(Nummov As String, Voucher As String, ByRef rNummov As String, ByRef rVoucher As String)
    Dim Fila As Long
    Dim i As Integer
    
    For i = 0 To lArrVou.Count(1) - 1
        If lArrVou(i, 2) = Nummov And lArrVou(i, 3) = Voucher Then
            rNummov = lArrVou(i, 0)
            rVoucher = lArrVou(i, 1)
            Exit For
        End If
    Next i
        
End Sub

Private Function ImportaAsientosDet(tabla As String) As Boolean
    Dim rsExp As New ADODB.Recordset
    Dim rsExpDet As New ADODB.Recordset
    Dim sqlExp As String
    Dim auxInt As String
    Dim auxVou As String
    Dim lArrCab() As Variant    ' *** Variable de arreglo
    Dim auxPer As String
    Dim auxLib As String
    Dim auxUsu As String
    Dim auxIntOri As String
    Dim auxIntOri2 As String
    'Dim clsMante As clsMantoTablas

    
    On Error GoTo MsgError
    rsExp.Open tabla
    PGBAvanceConfigura 0, rsExp.RecordCount, "Iniciando proceso ..."
    
    Dim nNumMov As String
    'auxIntOri = "0"
    'auxIntOri2 = ""
    'Set clsMante = New clsMantoTablas
    
    Dim nItem As Long
    '55
    ReDim lArrCab(60) As Variant

    Call Conectar
    
    
''    clsMante.InicializaClase
''    clsMante.BeginTrans
    Do While Not rsExp.EOF
           
        nNumMov = CE(rsExp("Ase_cNummov").Value)
        auxVou = CE(rsExp("Ase_nVoucher").Value)  'auxVou
        nItem = CE(rsExp("Asd_nItem").Value)
           
        Call LlenaArregloDet(nNumMov, auxVou, nItem, rsExp, lArrCab)
        Dim sql As String
        sql = "set dateformat dmy " & vbCrLf & _
        "INSERT Into CND_ASIENTO_VOUCHER  WITH ( ROWLOCK ) " & _
        "(Emp_cCodigo, Pan_cAnio, Per_cPeriodo, Lib_cTipoLibro, Ase_nVoucher, Asd_nItem, Pla_cCuentaContable, Asd_cGlosa, " & _
        "Asd_nDebeSoles, Asd_nDebeMonExt, Asd_nHaberSoles, Asd_nHaberMonExt, Asd_nTipoCambio, Cos_cCodigo, Imp_nPorcentaje, " & _
        "Asd_cOperaTC, Asd_cTipoMoneda, Asd_cMonedaCalculo,Ten_cTipoEntidad, Ent_cCodEntidad, Asd_cTipoDoc, Asd_cSerieDoc, Asd_cNumDoc, Asd_cTipoDocRef, Asd_cSerieDocRef, " & _
        "Asd_cNumDocRef, Asd_nMontoInafecto, Asd_cRetencion, Asd_dFechaSpot, Asd_cNumSpot, Asd_cDestino, Asd_nCorre, Asd_cEstado, " & _
        "Asd_cProvCanc, Asd_cFlgSpot,Asd_cDeleted, Asd_cUserCrea, Asd_dFechaCrea, Asd_cUserModifica, Asd_dFechaModifica, Asd_cEquipoUser, " & _
        "Asd_dFecDoc, Asd_dFecDocRef, Ase_cNummov,Com_cTipoIgv,Asd_dFecVen, " & _
        "Tra_cCodigo,Asd_cFormaPago, Asd_cBaseImp, Asd_cMonAdic, Asd_cImpAdic, Asd_cComprobante, Asd_cProceso, " & _
        "ECP_COPERACION, Asd_cRegAux, Asd_cConvMon, Asd_cManual , Asd_cRegAuxDet , Asd_cGrupo, Id_Exoneracion, Id_Tipo_Renta, Id_Modalidad, Id_Aduana, Id_Clasific_Servicio ) " & _
        "Values " & _
        "('" & CE(lArrCab(2)) & "','" & CE(lArrCab(3)) & "','" & CE(lArrCab(4)) & "','" & CE(lArrCab(5)) & "','" & CE(lArrCab(6)) & "'," & NE(lArrCab(7)) & ",'" & CE(lArrCab(8)) & "','" & CE(lArrCab(9)) & "'," & _
        "" & NE(lArrCab(10)) & "," & NE(lArrCab(11)) & "," & NE(lArrCab(12)) & "," & NE(lArrCab(13)) & "," & NE(lArrCab(14)) & ",'" & CE(lArrCab(15)) & "'," & NE(lArrCab(39)) & "," & _
        "'" & CE(lArrCab(36)) & "','" & CE(lArrCab(37)) & "','" & CE(lArrCab(38)) & "','" & CE(lArrCab(16)) & "','" & CE(lArrCab(17)) & "','" & CE(lArrCab(18)) & "','" & CE(lArrCab(19)) & "','" & CE(lArrCab(20)) & "','" & CE(lArrCab(22)) & "','" & CE(lArrCab(23)) & "'," & _
        "'" & CE(lArrCab(24)) & "'," & NE(lArrCab(26)) & ",'" & CE(lArrCab(27)) & "',convert(DateTime, '" & Left(CE(lArrCab(29)), 20) & "'),'" & CE(lArrCab(30)) & "'," & CE(lArrCab(31)) & ",'" & NE(lArrCab(32)) & "','" & CE(lArrCab(34)) & "'," & _
        "'" & CE(lArrCab(35)) & "','','','" & CE(lArrCab(33)) & "',GETDATE(),'" & CE(lArrCab(33)) & "',GETDATE(),HOST_NAME(),convert(DateTime, '" & Left(CE(lArrCab(21)), 20) & "'),convert(DateTime, '" & Left(CE(lArrCab(25)), 20) & "')," & _
        "'" & CE(lArrCab(1)) & "','',convert(DateTime, '" & Left(CE(lArrCab(41)), 20) & "')," & _
        "'" & CE(lArrCab(42)) & "','" & CE(lArrCab(43)) & "','" & CE(lArrCab(44)) & "','" & CE(lArrCab(45)) & "'," & NE(lArrCab(46)) & ",'" & CE(lArrCab(47)) & "','" & CE(lArrCab(48)) & "'," & _
        "'" & CE(lArrCab(49)) & "','" & CE(lArrCab(50)) & "','" & CE(lArrCab(51)) & "','" & CE(lArrCab(52)) & "','" & CE(lArrCab(53)) & "' , '', " & CE(lArrCab(56)) & ", " & CE(lArrCab(57)) & ", " & CE(lArrCab(58)) & ", " & CE(lArrCab(59)) & ", " & CE(lArrCab(60)) & ")"

         gcnSistema.Execute sql
           
        Debug.Print sql
        
''        If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaAsientoDet", lArrCab(), False) = False Then
''            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
''            CerrarRecordSet rsExp
''
''            clsMante.CancelTrans
''            clsMante.FinalizaClase
''
''            ProcesoTerminado = False
''            Set clsMante = Nothing
''            PGBAvanceConfigura 0, 1, ""
''            'ImportaAsientosDet = False
''
''            clsMante.CancelTrans
''            clsMante.FinalizaClase
''
''            Exit Function
''        End If
''
''        clsMante.CommitTrans
        
        PGBAvanceActualiza rsExp.AbsolutePosition, "Detalle de voucher"
        rsExp.MoveNext
        
    Loop
    
    Call Desconectar
    
''    clsMante.CommitTrans
''    clsMante.FinalizaClase
    
''    Set clsMante = Nothing
    
    CerrarRecordSet rsExp
    PGBAvanceConfigura 0, 1, ""
    ProcesoTerminado = True
    ImportaAsientosDet = True
    Exit Function

MsgError:
    Call Desconectar

    PGBAvanceConfigura 0, 1, ""
    CerrarRecordSet rsExp

    Mensajes Err.Description, vbOKOnly + vbInformation
    ImportaAsientosDet = False
    
''    clsMante.CancelTrans
''    clsMante.FinalizaClase
End Function

Private Sub LlenaArregloDet(ByRef nNumMov As String, ByRef auxVou As String, ByRef nItem As Long, ByRef rsExp As ADODB.Recordset, ByRef lArrCab As Variant)
On Error GoTo serror
    ' *** Cargar los datos a grabar en un arreglo
    lArrCab(0) = "INSERTAR_IMPORTACION"                 ' Accion
    lArrCab(1) = nNumMov                                ' Numero Interno
    lArrCab(2) = gsEmpresa                              ' Empresa
    lArrCab(3) = gsAnio                                 ' Año
    lArrCab(4) = CE(rsExp("Per_cPeriodo").Value)        ' Periodo
    lArrCab(5) = CE(rsExp("Lib_cTipoLibro").Value)      ' Libro
    lArrCab(6) = auxVou                                 ' Voucher
    lArrCab(7) = nItem 'NE(rsExp("Asd_nItem").Value)               ' item
    lArrCab(8) = CE(rsExp("Pla_cCuentaContable").Value) ' Plan Cuenta
    lArrCab(9) = CC(CE(rsExp("Asd_cGlosa").Value))          ' Glosa
    
    lArrCab(10) = NE(rsExp("Asd_nDebeSoles").Value)     ' DebeSoles
    lArrCab(11) = NE(rsExp("Asd_nDebeMonExt").Value)    ' DebeMonExt
    lArrCab(12) = NE(rsExp("Asd_nHaberSoles").Value)    ' HaberSoles
    lArrCab(13) = NE(rsExp("Asd_nHaberMonExt").Value)   ' HaberMonExt
    lArrCab(14) = NE(rsExp("Asd_nTipoCambio").Value)    ' Tipo de Cambio
    
    lArrCab(15) = CE(rsExp("Cos_cCodigo").Value)        ' CCosto
    lArrCab(16) = CE(rsExp("Ten_cTipoEntidad").Value)   ' Tipo Entidad
    lArrCab(17) = CE(rsExp("Ent_cCodEntidad").Value)    ' Codigo Entidad
    lArrCab(18) = CE(rsExp("Asd_cTipoDoc").Value)       ' Tipo Doc
    lArrCab(19) = CE(rsExp("Asd_cSerieDoc").Value)      ' Serie
    lArrCab(20) = CE(rsExp("Asd_cNumDoc").Value)        ' Numero Doc
    lArrCab(21) = FechaNula(rsExp("Asd_dFecDoc").Value)        ' Fecha Doc
    lArrCab(22) = CE(rsExp("Asd_cTipoDocRef").Value)    ' TipoDoc Ref
    lArrCab(23) = CE(rsExp("Asd_cSerieDocRef").Value)   ' SerieDoc Ref
    lArrCab(24) = CE(rsExp("Asd_cNumDocRef").Value)     ' NumeroDoc Ref
    lArrCab(25) = FechaNula(rsExp("Asd_dFecDocRef").Value)     ' FechaDoc Ref
    lArrCab(26) = NE(rsExp("Asd_nMontoInafecto").Value) ' MontoInafecto
    lArrCab(27) = CE(rsExp("Asd_cRetencion").Value)     ' Retencion
    lArrCab(28) = ""
    lArrCab(29) = FechaNula(rsExp("Asd_dFechaSpot").Value)     ' Fecha Spot
    lArrCab(30) = CE(rsExp("Asd_cNumSpot").Value)       ' NumSpot
    lArrCab(31) = CE(rsExp("Asd_cDestino").Value)       ' Destino
    lArrCab(32) = NE(rsExp("Asd_nCorre").Value)         ' Correlativo Provision
    lArrCab(33) = CE(rsExp("Asd_cUserCrea").Value)      ' Usuario
    lArrCab(34) = CE(rsExp("Asd_cEstado").Value)        ' Estado
    lArrCab(35) = CE(rsExp("Asd_cProvCanc").Value)      ' Indica Si es Provision/Cancelacion o Ninguno
    
    lArrCab(36) = CE(rsExp("Asd_cOperaTC").Value)       ' TipoCompra (A B C)
    lArrCab(37) = CE(rsExp("Asd_cTipoMoneda").Value)       ' Tipo PatrimonioNeto
    lArrCab(38) = CE(rsExp("Asd_cMonedaCalculo").Value)
    lArrCab(39) = NE(rsExp("Imp_nPorcentaje").Value)
    lArrCab(40) = ""
    lArrCab(41) = FechaNula(rsExp("Asd_dFecVen").Value)
    lArrCab(42) = CE(rsExp("Tra_cCodigo").Value)
    lArrCab(43) = CE(rsExp("Asd_cFormaPago").Value)
    lArrCab(44) = CE(rsExp("Asd_cBaseImp").Value)
    lArrCab(45) = CE(rsExp("Asd_cMonAdic").Value)
    lArrCab(46) = NE(rsExp("Asd_cImpAdic").Value)
    
    lArrCab(47) = CE(rsExp("Asd_cComprobante").Value)
    lArrCab(48) = CE(rsExp("Asd_cProceso").Value)
    
    lArrCab(49) = CE(rsExp("ECP_COPERACION").Value)
    lArrCab(50) = CE(rsExp("Asd_cRegAux").Value)
    lArrCab(51) = CE(rsExp("Asd_cConvMon").Value)
    
    lArrCab(52) = CE(rsExp("Asd_cManual").Value)
    lArrCab(53) = CE(rsExp("Asd_cRegAuxDet").Value)
    lArrCab(54) = ""
    lArrCab(55) = ""
    
    lArrCab(56) = IIf(CE(rsExp("Id_Exoneracion").Value) = "", "Null", "'" & CE(rsExp("Id_Exoneracion").Value) & "'")
    lArrCab(57) = IIf(CE(rsExp("Id_Tipo_Renta").Value) = "", "Null", "'" & CE(rsExp("Id_Tipo_Renta").Value) & "'")
    lArrCab(58) = IIf(CE(rsExp("Id_Modalidad").Value) = "", "Null", "'" & CE(rsExp("Id_Modalidad").Value) & "'")
    lArrCab(59) = IIf(CE(rsExp("Id_Aduana").Value) = "", "Null", "'" & CE(rsExp("Id_Aduana").Value) & "'")
    lArrCab(60) = IIf(CE(rsExp("Id_Clasific_Servicio").Value) = "", "Null", "'" & CE(rsExp("Id_Clasific_Servicio").Value) & "'")
    Exit Sub
serror:
    Mensajes Err.Description
End Sub

Private Function FechaNula(fecha As Variant)
    If IsNull(fecha) Then fecha = ""
    If Not IsDate(fecha) Then fecha = ""
    If CE(fecha) = "" Then
        FechaNula = Null
    Else
        FechaNula = CE(fecha)
    End If
End Function

Private Sub Form_Load()
    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))
    
    
    Dim sqlcombos As String
    
    Call Centrar_form(Me)
    
    
    Drive1.Drive = "C:\"
    Dir1.Path = "C:\"
    Dir1.Refresh
    
    If chkMovim.Value = vbChecked Then
        chkEliminar.Value = vbChecked
        chkEliminar.Enabled = False
        
    Else
        chkEliminar.Enabled = True
        
    End If
        
    
    If Mid(gsGrupo, 3, 1) <> "1" And gsGrupo <> gsPrivilegioAdmin Then
        Me.cmdImportarDatos.Enabled = False
        
    Else
        Me.cmdImportarDatos.Enabled = True
        
    End If
    
End Sub

Private Function ExisteRucTd(Tipo As String, td As String, Ruc As String) As Boolean
    Dim rsCosto As New ADODB.Recordset
    Dim sqlver As String
    
    ExisteRucTd = False

    sqlver = "select Ent_cCodEntidad from dbo.CNM_ENTIDAD WHERE Emp_cCodigo = '" & gsEmpresa & "' " & _
             " AND Ten_cTipoEntidad =  '" & Tipo & "' AND Ent_cTipoDoc =  '" & td & "' " & _
             " AND Ent_nRuc = '" & Ruc & "' "
    Call Conectar
    rsCosto.Open sqlver, gcnSistema
    If Not rsCosto.EOF And Not rsCosto.BOF Then ExisteRucTd = True
    Call CerrarRecordSet(rsCosto)
    Call Desconectar
End Function

Private Function ExisteEntidad(Tipo As String, Codigo As String) As Boolean
    Dim rsCosto As New ADODB.Recordset
    Dim sqlver As String
    
    ExisteEntidad = False

    sqlver = "select Ent_cCodEntidad from dbo.CNM_ENTIDAD WHERE Emp_cCodigo = '" & gsEmpresa & "' " & _
             " AND Ten_cTipoEntidad =  '" & Tipo & "' AND Ent_cCodEntidad =  '" & Codigo & "' "
    Call Conectar
    rsCosto.Open sqlver, gcnSistema
    If Not rsCosto.EOF And Not rsCosto.BOF Then ExisteEntidad = True
    Call CerrarRecordSet(rsCosto)
    Call Desconectar
End Function

Private Function ExisteCtaImp(valorCta As String) As Boolean
    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
    
    ' *** Verificando q cuenta exista
    Dim sqlSp As String
    
    Set clDatos = New clsMantoTablas
    'sqlSp = "spCn_ConsultaCuentas 'SEL_REG', '" & gsEmpresa & "', '" & gsAnio & "', '" & valorCta & "'"
    
    sqlSp = "SELECT Emp_cCodigo, Pan_cAnio, Pla_cCuentaContable, Pla_cNombreCuenta, Pla_cTitulo, " & _
    "Ten_cTipoEntidad, Pla_cCentroCosto, Pla_cProvision, Pla_cDifCambio, Pla_cOperaTC, Pla_cRedondeo, Pla_cDocumento, Pla_cTipoCta, " & _
    "Pla_cCptoBG, Pla_cCptoBGDual, Pla_cCptoResFun, Pla_cCptoResNat, Pla_cTipoAfect, Pla_cDetraccion, Pla_cRetencion, Pla_cPercepcion, Pla_cCtaPresup, " & _
    "Pla_cEstado , Pla_cDeleted, Pla_cUserCrea, Pla_dFechaCrea, Pla_cUserModifica, Pla_dFechaModifica, Pla_cEquipoUser " & _
    " From CNM_PLAN_CTA WITH(READUNCOMMITTED) WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' AND Pla_cCuentaContable = '" & valorCta & "'"

    arrDatos = Array(sqlSp)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo.State = 0 Then
        ExisteCtaImp = False
    Else
        ExisteCtaImp = True
    End If
    
    Call CerrarRecordSet(rsArreglo)
    Set clDatos = Nothing
    
End Function

Private Function BuscaNumMov(Tipo As String, empresa As String, Anio As String, numTrans As String, periodo As String, Libro As String, Sucursal As String) As String
    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
    
    ' *** Verificando si asiento existe
    Dim sqlSp As String
    
    Set clDatos = New clsMantoTablas
    sqlSp = "spCn_ConsultaAsientosSuc '" & Tipo & " ', '" & empresa & "', '" & Anio & "', '" & periodo & "', '" & Libro & "', '" & numTrans & "', '" & Sucursal & "'"
    arrDatos = Array(sqlSp)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo.State = 0 Then
        BuscaNumMov = "0000000000"
    Else
        BuscaNumMov = Right("0000000000" & CE(rsArreglo!Ase_cNummov), 10)
    End If
    Call CerrarRecordSet(rsArreglo)
    Set clDatos = Nothing
End Function

Private Function ExisteAsientoSuc(Tipo As String, empresa As String, Anio As String, numTrans As String, periodo As String, Libro As String, Sucursal As String) As Boolean
    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
    
    ' *** Verificando si asiento existe
    Dim sqlSp As String
    
    Set clDatos = New clsMantoTablas
    sqlSp = "spCn_ConsultaAsientosSuc '" & Tipo & " ', '" & empresa & "', '" & Anio & "', '" & periodo & "', '" & Libro & "', '" & numTrans & "', '" & Sucursal & "'"
    arrDatos = Array(sqlSp)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo.State = 0 Then
        ExisteAsientoSuc = False
    Else
        ExisteAsientoSuc = True
    End If
    
    Call CerrarRecordSet(rsArreglo)
    Set clDatos = Nothing
End Function

Private Sub CargaArregloReplica()
    ReDim lArrMnt(32) As Variant
    lArrMnt(0) = gsEmpresa
    lArrMnt(1) = gsAnio
    lArrMnt(2) = gsEmpresa
    lArrMnt(3) = gsAnio
    lArrMnt(4) = gsUsuario
    
    lArrMnt(5) = CE(chkPlan.Value)
    lArrMnt(6) = CE(chkCenCos.Value)
    lArrMnt(7) = CE(chkLibros.Value)
    lArrMnt(8) = CE(chkEntidades.Value)
    lArrMnt(9) = CE(chkEntidades.Value)
    
    lArrMnt(10) = CE(chkEntidades.Value)
    lArrMnt(11) = CE(chkTipDoc.Value)
    lArrMnt(12) = CE(chkTabSec.Value)
    lArrMnt(13) = CE(chkTipMon.Value)
    lArrMnt(14) = CE(chkTipoCbio.Value)
    
    lArrMnt(15) = CE(chkBancos.Value)
    lArrMnt(16) = CE(chkCtaCte.Value)
    lArrMnt(17) = CE(chkTipAsto.Value)
    lArrMnt(18) = CE(chkPlanCONA.Value)
    lArrMnt(19) = CE(chkParamIni.Value)
    
    lArrMnt(20) = CE(chkParamIni.Value)
    lArrMnt(21) = CE(chkConfigOp.Value)
    lArrMnt(22) = CE(chkConfigOp.Value)
    lArrMnt(23) = CE(chkRatios.Value)
    lArrMnt(24) = CE(chkMovim.Value)
    
    lArrMnt(25) = CE(chkPatrimonio.Value)
    lArrMnt(26) = CE(chkFlujo.Value)
    lArrMnt(27) = CE(chkPresup.Value)
    lArrMnt(28) = CE(chkCapital.Value)
    lArrMnt(29) = CE(chkPDB.Value)
    
    lArrMnt(30) = CE(chkValores.Value)
    lArrMnt(31) = CE(chkCostos.Value)
    lArrMnt(32) = CE(chkEliminarTodosAnios.Value)
    
End Sub

Private Function Importar(tabla As String, Titulo As String, archivo As String) As Boolean
    Importar = False
    Dim rsExp As New ADODB.Recordset
    Dim sqlExp As String
    Dim j As Integer
    Dim nNumMov As String
    Dim auxVou As String
    Dim Indice As Long
    
    On Error GoTo MsgError
    Call Conectar
    
    rsExp.Open archivo
    PGBAvanceConfigura 0, rsExp.RecordCount, "Iniciando proceso ..."
    
    If tabla = "CConasev" Then
        Call gcnSistema.Execute("DELETE FROM CNA_TIPO_PLANTILLA Where Emp_cCodigo = '" & gsEmpresa & "' And Pan_cAnio = '" & gsAnio & "'")
    End If
    
    If tabla = "CConasev2" Then
        Call gcnSistema.Execute("DELETE FROM CNT_OPERA_ESTADO Where Emp_cCodigo = '" & gsEmpresa & "' And Pan_cAnio = '" & gsAnio & "'")
    End If
    'Entidades
    Dim strIdPais, strIdVinculo, strIdConvenio As String
    'Cabecera y Detalle Voucher
    Dim strIdExoneracion, strIdTipoRenta, strIdModalidad, strIdAduana, strIdClasificServicio As String
    
    
    Do While Not rsExp.EOF
            sqlExp = ExtraeCadenaSQL(tabla)

            For j = 0 To rsExp.Fields.Count - 1
                strIdPais = vbNullString: strIdVinculo = vbNullString: strIdConvenio = vbNullString
                strIdExoneracion = vbNullString: strIdTipoRenta = vbNullString: strIdModalidad = vbNullString: strIdAduana = vbNullString: strIdClasificServicio = vbNullString
                If j = rsExp.Fields.Count - 1 Then
                    If rsExp.Fields(j).Type = adNumeric Or rsExp.Fields(j).Type = adCurrency Or rsExp.Fields(j).Type = adInteger Then
                       sqlExp = sqlExp & " " & NE(rsExp(rsExp.Fields(j).Name).Value) & ") "
                    Else
                     If CE(rsExp.Fields(j).Name) = "Pan_cAnio" Then
                       sqlExp = sqlExp & " '" & gsAnio & "') "
                     Else
                        If tabla = "CParam" And CE(rsExp.Fields(j).Name) = "Cfl_cEquipoUser" Then
                            sqlExp = sqlExp & " '" & CE(rsExp(rsExp.Fields(j).Name).Value) & "', "
                            sqlExp = sqlExp & " '" & gsAnio & "')"
                        Else
                            If rsExp.Fields(j).Name = "Id_Convenio" Then
                                strIdVinculo = IIf(CE(CM(rsExp(rsExp.Fields(j).Name).Value)) = vbNullString, "Null", "'" & CE(CM(rsExp(rsExp.Fields(j).Name).Value)) & "'")
                                sqlExp = sqlExp & " " & strIdVinculo & ") "
                            Else
                                sqlExp = sqlExp & " '" & CE(rsExp(rsExp.Fields(j).Name).Value) & "') "
                            End If
                            
                        End If
                     End If
                    End If
                Else
                    If CE(rsExp.Fields(j).Name) = "Blta_cDNI" Then 'And CE(rsExp(rsExp.Fields(j).Name).Value) = "Blta_nBaseImp" Then
                        sqlExp = sqlExp
                    End If
                'Blta_nBaseImpInafD
                    'If tabla = "CRegAux" Then
                    '    sqlExp = sqlExp
                    'End If
                    
                    If UCase(CE(rsExp.Fields(j).Name)) = "EMP_CCODIGO" Then
                        sqlExp = sqlExp & " '" & gsEmpresa & "', "
                    ElseIf UCase(CE(rsExp.Fields(j).Name)) = "PAN_CANIO" And tabla <> "CFlujoSal" Then
                        sqlExp = sqlExp & " '" & gsAnio & "', "
                        
                    'ElseIf UCase(CE(rsExp.Fields(j).Name)) = "ASE_CNUMMOV" And tabla = "CCoa" Then

                        
                    '    Indice = lArrVou.Find(0, 1, CE(rsExp.Fields("Ase_nVoucher")))
                    '    nNumMov = lArrVou(Indice, 0)
                        
                    
                    '    sqlExp = sqlExp & " '" & nNumMov & "', "
                    
                    Else
                        If IsDate(rsExp(rsExp.Fields(j).Name).Value) Then
                            sqlExp = sqlExp & " '" & CE(Format(rsExp(rsExp.Fields(j).Name).Value, "dd/MM/yyyy")) & "', "
                        Else
                            If UCase(CE(rsExp.Fields(j).Name)) = "TAB_NLONGITUD" Then
                                sqlExp = sqlExp & " '" & NE(rsExp(rsExp.Fields(j).Name).Value) & "', "
                            ElseIf rsExp.Fields(j).Type = adNumeric Or rsExp.Fields(j).Type = adCurrency Or rsExp.Fields(j).Type = adInteger Then
                                sqlExp = sqlExp & " " & NE(rsExp(rsExp.Fields(j).Name).Value) & ", "
                            Else
                                
                                If rsExp.Fields(j).Name = "Id_Pais" Then
                                    strIdPais = IIf(CE(CM(rsExp(rsExp.Fields(j).Name).Value)) = vbNullString, "Null", "'" & CE(CM(rsExp(rsExp.Fields(j).Name).Value)) & "'")
                                    sqlExp = sqlExp & " " & strIdPais & ", "
                                ElseIf rsExp.Fields(j).Name = "Id_Vinculo_Economico" Then
                                    strIdVinculo = IIf(CE(CM(rsExp(rsExp.Fields(j).Name).Value)) = vbNullString, "Null", "'" & CE(CM(rsExp(rsExp.Fields(j).Name).Value)) & "'")
                                    sqlExp = sqlExp & " " & strIdVinculo & ", "
                                ElseIf rsExp.Fields(j).Name = "Id_Convenio" Then
                                    strIdConvenio = IIf(CE(CM(rsExp(rsExp.Fields(j).Name).Value)) = vbNullString, "Null", "'" & CE(CM(rsExp(rsExp.Fields(j).Name).Value)) & "'")
                                    sqlExp = sqlExp & " " & strIdConvenio & ", "
                                Else
                                    sqlExp = sqlExp & " '" & CE(CM(rsExp(rsExp.Fields(j).Name).Value)) & "', "
                                End If
                                
                            End If
                        End If
                    End If
                End If
            Next
            
            Debug.Print sqlExp
            On Error Resume Next
            gcnSistema.Execute "set dateformat dmy " & vbCrLf & " " & sqlExp

           If Err.Number <> 0 Then
                If Err.Number = -2147217873 Then
                    'no insertar y dejar pasar al siguiente
                    Debug.Print "Registro duplicado no se insertará"
                Else
                    GoTo MsgError
                End If
           End If
            

        PGBAvanceActualiza rsExp.AbsolutePosition, Titulo
        rsExp.MoveNext
    Loop
    Call Desconectar
    CerrarRecordSet rsExp
    PGBAvanceConfigura 0, 1, ""
    Importar = True
    Exit Function
MsgError:
    MsgBox (sqlExp)
    Call Desconectar
    
    PGBAvanceConfigura 0, 1, ""
    CerrarRecordSet rsExp
    'Mensajes Err.Description, vbOKOnly + vbInformation
    Importar = False
End Function

Private Function CM(cadena As Variant) As String
    If IsNull(cadena) Then
        CM = ""
    Else
        CM = Replace(cadena, "'", "")
    End If
End Function

Private Function ExtraeCadenaSQL(tabla As String) As String
    Dim sql As String
    Select Case tabla
        Case "CRegAux":
            sql = "INSERT INTO CNT_REG_BOLETAS (Asd_cNumDocORIGEN, Asd_cSerieDocORIGEN, Ase_cNummov, Ase_nVoucher,  Emp_cCodigo, Pan_cAnio, Per_cPeriodo, Blta_cFlagLibro," & _
                  "Tdo_cCodigo, Blta_Correlativo, Asd_cSerieDoc, Asd_cNumDoc, Asd_dFecha, Blta_nBaseImp, Blta_nIGV, Blta_nTotal, Mon_cCodigo, " & _
                  "Blta_nTipoCambio, Tca_nAuxiliar, Blta_nBaseImpEXT, Blta_nIGVEXT, Blta_nTotalEXT, Blta_nBaseImpInaf, Blta_nOtros, Blta_nOtrosD, Blta_cInafecto," & _
                  "Tab_cTabla, Blta_num_doc, Blta_cNombres, Blta_cApellidos, Blta_cDNI, Blta_cEstado, " & _
                  "Blta_nBaseImpInafD ) VALUES("
        
        Case "CConcil":
            sql = "INSERT INTO CNM_MOV_CHEQUE (Emp_cCodigo, Pan_cAnio, Per_cPeriodo, Che_nVoucherPago, Che_cNummovVoucher, Che_nItemVoucher, Ban_cCodigo, " & _
                  "Cue_cNumCuenta, Che_cTipoMov, Che_cTipoDoc, Che_cOperaCheque, Che_dFechaCheque, Che_nTipoCambio, " & _
                  "Che_nMontoS, Che_nMontoD, Che_dFechaOpera, Che_cObservacion, Che_cGlosa, Che_cEstado, Che_cDeleted," & _
                  "Che_cUserCrea, Che_dFechaCrea, Che_cUserModifica, Che_dFechaModifica, Che_cEquipoUser) VALUES("
        
        Case "CVal":
            sql = "INSERT INTO CND_VALORES_DETALLE (Emp_cCodigo, Pan_cAnio, Per_cPeriodo, Ase_nVoucher, Asd_nItem, Ten_cTipoEntidad, Ent_cCodentidad, Val_cTitulo, " & _
                  "Val_cDesTitulo, Val_nValorNom, Val_nCantidad, Val_nCostoTot, Val_nProvTot, Val_nTotalNeto) VALUES("
        
        Case "CCos":
            sql = "INSERT INTO CNT_COSTOS (Emp_cCodigo, Pan_cAnio, Cos_cCodigo) VALUES("
        
        Case "CCosInv":
            sql = "INSERT INTO CND_COSTOS_SALDOS (Emp_cCodigo, Pan_cAnio, per_ctipo, pro_cproceso, pro_ccodigo, pro_ncol01, pro_ncol02, pro_ncol03, pro_ncol04, " & _
                  "pro_ncol05, pro_ncol06, pro_ncol07, pro_ncol08, pro_ncol09, pro_ncol10, pro_ncol11, pro_ncol12, pro_ncol13) VALUES("
            
        Case "CFlujoPro":
            sql = "INSERT INTO CNT_FLUJO_PROCESO (Emp_cCodigo, Pan_cAnio, Per_cPeriodo, Pro_cTipoCta, Pro_cCuenta, Pro_cActividad, Pro_cTipoD, " & _
                  "Pro_cFormulaD, Pro_cDetalleD, Pro_cTipoH, Pro_cFormulaH, Pro_cDetalleH, Pro_cMetodo) VALUES("
        
        Case "CFlujoRep":
            sql = "INSERT INTO CNT_FLUJO_REPORTE (Emp_cCodigo, Pan_cAnio, Per_cPeriodo, Rep_cCuenta, " & _
                  "Rep_cCodTipo, Rep_cFormula, Rep_cValor, Pro_cMetodo ) VALUES("
        
        Case "CFlujoSal":
            sql = "INSERT INTO CNT_FLUJO_SALDOINI (Emp_cCodigo, Pan_cAnio, Per_cPeriodo, Sal_cCodigo, Sal_nSaldo" & _
                  ") VALUES("
        
        Case "CFlujoCta":
            sql = "INSERT INTO CNT_FLUJO_CUENTAS (Emp_cCodigo, Pan_cAnio, Per_cPeriodo, Pro_cCuenta" & _
                  ") VALUES("
        
        Case "CPatrim":
            sql = "INSERT INTO CNM_PATRIMONIO (Emp_cCodigo, Pan_cAnio, Pat_cCodigo, Pat_cCol01, Pat_cCol02, " & _
                  "Pat_cCol03, Pat_cCol04, Pat_cCol05, Pat_cCol06, Pat_cCol07" & _
                  ") VALUES("
        
        Case "CCapital":
            sql = "INSERT INTO CND_CAPITAL_DETALLE (Emp_cCodigo, Pan_cAnio, Per_cPeriodo, Ase_nVoucher, Asd_nItem, Ten_cTipoEntidad, " & _
                  "Ent_cCodentidad, Cap_cAcciones, Cap_nImportes, Cap_nValorNom, Cap_nASuscritas, " & _
                  "Cap_nAPagadas, Cap_nAcciones, Cap_nPorcent) VALUES("
        
        Case "CPresup":
            sql = "INSERT INTO PRM_MARCO_PRES (Emp_cCodigo, Pan_cAnio, Per_cPeriodo, Prm_cTipo, Cos_cCodigo, Mon_cCodigo, " & _
                  "Prm_nMontoPreS, Prm_nTipoCambioPres, Prm_nMontoPreD, Prm_dFechaPres, " & _
                  "Prm_cObserva, Prm_cEstado, Prm_cDeleted, Prm_cUserCrea, Pm_dFechaCrea, " & _
                  "Prm_cUserModifica, Prm_dFechaModifica, Prm_cEquipoUser) VALUES("
        
        Case "CPDB":
            sql = "INSERT INTO CND_ASIENTO_PDB (Ase_cNummov,Ase_nVoucher,Emp_cCodigo,Pan_cAnio,Per_cPeriodo,Lib_cTipoLibro,Dco_cTipoPDB,Dco_nItem,Dco_cTipoComVen," & _
                  "Dco_cTipoComprob,Dco_cFecha,Dco_cSerie,Dco_cNumero,Dco_cTipoPer,Dco_cTipoDocId,Dco_cNumDocId,Dco_cNombre,Dco_cApePat," & _
                  "Dco_cApeMat,Dco_cNombre1,Dco_cNombre2,Dco_cCodMon,Dco_cCodDest,Dco_cNunDest,Dco_nBaseImp,Dco_nISC,Dco_nIGV,Dco_nOtros," & _
                  "Dco_cIndDetra,Dco_cCodTasaDetra,Dco_cNumDetra,Dco_cIndRete,Dco_cRefTipoComp,Dco_cRefSerieComp,Dco_cRefNumComp,Dco_cRefFechaEmi," & _
                  "Dco_nRefBaseImp,Dco_nRefIGV,Dco_cMedPago,Dco_cCodBaco,Dco_cNumOp,Dco_dFechaOp,Dco_cMontoOp,Dco_cEstado,Dco_cDeleted,Dco_cUserCrea," & _
                  "Dco_dFechaCrea,Dco_cUserModifica,Dco_dFechaModifica,Dco_cEquipoUser) VALUES("
        
        Case "CCosto":
            sql = "INSERT INTO CNT_CENTRO_COSTO( " & _
                  "Emp_cCodigo, Pan_cAnio, Cos_cCodigo, Cos_cDescripcion, Cos_cTitulo, Cos_cSTitulo,  " & _
                  "Cos_cEstado, Cos_cDeleted, Cos_cUserCrea, Cos_dFechaCrea, Cos_cUserModifica, " & _
                  "Cos_dFechaModifica, Cos_cEquipoUser) VALUES("
                  
        Case "CLibros"
            sql = "INSERT INTO CNT_LIBRO_OPERA( " & _
                  "Emp_cCodigo, Pan_cAnio, Lib_cTipoLibro, Lib_cDescripcion, Lib_cTipOpe, " & _
                  "Lib_cFlagDocRef,Lib_cFlagAdelIgv, Lib_cFlagInafecto, Lib_cEstado, Lib_cDeleted, " & _
                  "Lib_cUserCrea, Lib_dFechaCrea, Lib_cUserModifica, Lib_dFechaModifica, Lib_cEquipoUser, Lib_cCodsunat) VALUES("
    
        Case "CEntidades":
            sql = "INSERT INTO CNT_ENTIDAD( " & _
                  "Emp_cCodigo, Ten_cTipoEntidad, Ten_cNombreEntidad, Ten_cEstado, Ten_cDeleted, " & _
                  "Ten_cUserCrea, Ten_dFechaCrea, Ten_cUserModifica, Ten_dFechaModifica, Ten_cEquipoUser) VALUES("
    
        Case "CEntidades2":
            sql = "INSERT INTO CNT_ENTIDAD_DOCU( " & _
                  "Emp_cCodigo, Ten_cTipoEntidad, Edoc_cTipoPersona, Edoc_cTipoDoc, Edoc_cEstado, Edoc_cDeleted, " & _
                  "Edoc_cUserCrea, Edoc_dFechaCrea, Edoc_cUserModifica, Edoc_dFechaModifica, Edoc_cEquipoUser) VALUES("
    
        Case "CEntidades3":
            sql = "INSERT INTO CNM_ENTIDAD( " & _
                  "Emp_cCodigo, Ent_cCodEntidad, Ten_cTipoEntidad, Ent_cPersona, Ent_cDireccion, Ent_nRuc, " & _
                  "Ent_cRepresentante, Ent_cTipoDoc, Ent_cFlagPersona, Ent_cEstadoEntidad, Ent_cEstado, Ent_cDeleted, " & _
                  "Ent_cUserCrea, Ent_dFechaCrea, Ent_cUserModifica, Ent_dFechaModifica, Ent_cEquipoUser, Ent_cFlagDomiciliado, Id_Pais, Id_Vinculo_Economico, Id_Convenio, PorcentajeSunat) VALUES("
    
        Case "CDocumentos":
            sql = "INSERT INTO CNT_TIPODOC( " & _
                  "Emp_cCodigo, Tdo_cCodigo, Tdo_cNombreLargo, Tdo_cNombreCorto, " & _
                  "Tdo_cEstado, Tdo_cDaot, Tdo_cNatDaot, Tdo_cDeleted, " & _
                  "Tdo_cUserCrea, Tdo_dFechaCrea, Tdo_cUserModifica, Tdo_dFechaModifica, Tdo_cEquipoUser) VALUES("
        
        Case "CSec":
            sql = "INSERT INTO TABLA( " & _
                  "Emp_cCodigo, Tab_cTabla, Tab_cCodigo, Tab_cDescripCampo, Tab_cDescripTabla," & _
                  "Tab_nLongitud, Tab_cEstado, Tab_cMod01,Tab_cMod02,Tab_cMod03,Tab_cMod04,Tab_cMod05 , Tab_cMod06,Tab_cMod07,Tab_cMod08,Tab_cMod09, Tab_cMod10, Tab_cDeleted," & _
                  "Tab_cUserCrea, Tab_dFechaCrea, Tab_cUserModifica, Tab_dFechaModifica, Tab_cEquipoUser, Tab_cCodSunat  ) VALUES("
            
        Case "CMoneda":
            sql = "INSERT INTO CNT_TIPO_MONEDA(" & _
                  "Emp_cCodigo, Mon_cCodigo, Mon_cNombreLargo, Mon_cNombreCorto," & _
                  "Mon_cMNac, Mon_cMExt, Mon_cEstado, Mon_cDeleted," & _
                  "Mon_cUserCrea, Mon_dFechaCrea, Mon_cUserModifica, Mon_dFechaModifica, Mon_cEquipoUser, Mon_cCodsunat) VALUES("
            
        Case "CTc":
            sql = "INSERT INTO CNT_TIPO_CAMBIO(" & _
                  "Emp_cCodigo, Tca_dFecha, Tca_cCodigoOrigen, Tca_cCodigoDestino," & _
                  "Tca_nCompra, Tca_nVenta, Tca_nCompraP, Tca_nVentaP, Tca_cEstado," & _
                  "Tca_cDeleted, Tca_cUserCrea, Tca_dFechaCrea, Tca_cUserModifica, Tca_dFechaModifica, Tca_cEquipoUser) VALUES("
            
        Case "CTc2":
            sql = "INSERT INTO CNT_TIPO_CAMBIO_MENSUAL(" & _
                  "Emp_cCodigo, Pan_cAnio, Tca_cTipo, Tca_cMoneda, Tca_cEne, Tca_cFeb," & _
                  "Tca_cMar, Tca_cAbr, Tca_cMay, Tca_cJun , Tca_cJul, Tca_cAgo, Tca_cSet, Tca_cOct, Tca_cNov, Tca_cDic) VALUES("

        Case "CBancos":
            sql = "INSERT INTO CNT_BANCO(" & _
                  "Emp_cCodigo, Ban_cCodigo, Ban_cNombre, Ban_cEstado, Ban_cDeleted," & _
                  "Ban_cUserCrea, Ban_dFechaCrea, Ban_cUserModifica, Ban_dFechaModifica, Ban_cEquipoUser, Ban_cCodsunat) VALUES("
            
        Case "CCtacte":
            sql = "INSERT INTO CNM_CUENTA_BANCO(" & _
                  "Emp_cCodigo, Ban_cCodigo, Cue_cNumCuenta, Cue_cCuentaContable, Mon_cCodigo," & _
                  "Cue_dFechaApertura, Cue_dFechaCierre, Cue_cObservaCierre, Cue_nMonto," & _
                  "Cue_nNumChequeFin, Cue_nNumChequeIni, Cue_cEstado, Cue_cDeleted," & _
                  "Cue_cUserCrea, Cue_dFechaCrea, Cue_cUserModifica, Cue_dFechaModifica, Cue_cEquipoUser,Pan_cAnio) VALUES("
            
        Case "CTipoasiento":
            sql = "INSERT INTO CNT_ASIENTO_LIBRO(" & _
                  "Emp_cCodigo, Pan_cAnio, Lib_cTipoLibro, Asl_cOperacion, Asl_nSecuencia," & _
                  "Asl_cDescripcion, AslTipoMov, Asl_cCuenta, Asl_nPorcen, Asl_cEstado, Asl_cDeleted," & _
                  "Asl_cUserCrea, Asl_dFechaCrea, Asl_cUserModifica, Asl_dFechaModifica, Asl_cEquipoUser) VALUES("
                        
        Case "CConasev":
        
            sql = "INSERT INTO CNA_TIPO_PLANTILLA(" & _
                  "Emp_cCodigo, Ppa_cTipoPlantilla, Ppa_cNumPlantilla, Ppa_cNombre," & _
                  "Ppa_cTitulo, Ppa_cCodigoRef, Ppa_cResult, Ppa_cEstado, Ppa_cDeleted," & _
                  "Ppa_cUserCrea, Ppa_dFechaCrea, Ppa_cUseModifica, Ppa_dFechaModifica, Ppa_cEquipoUser, Pan_cAnio) VALUES("
            
        Case "CConasev2":
            sql = "INSERT INTO CNT_OPERA_ESTADO(" & _
                  "Emp_cCodigo, Pan_cAnio, Ecp_cOperacion, Ecp_cDescripcion, Ecp_cEstado, Ecp_cDeleted," & _
                  "Ecp_cUserCrea, Ecp_dFechaCrea, Ecp_cUserModifica, Ecp_dFechaModifica, Ecp_cEquipoUser) VALUES("
            
        Case "CParam":
            sql = "INSERT INTO CNT_CONFIG_LIBROS(" & _
                  "Emp_cCodigo, Cfl_cCompras, Cfl_cVentas, Cfl_cCaja, Cfl_cCajaIngresos," & _
                  "Cfl_cCajaEgresos, Cfl_cHonorarios, Cfl_cPercepcion, Cfl_cRetencion," & _
                  "Cfl_nPorcIGV, Cfl_cEstado, Cfl_cDeleted, Cfl_cDiario, Cfl_cDifCam, Cfl_cCierre,Cfl_cNivelCC,Cfl_cMesCompras,Cfl_cNDigCtas,Cfl_cApertura,Cfl_cBaseDefCom, " & _
                  "Cfl_cUserCrea, Cfl_dFechaCrea, Cfl_cUserModifica, Cfl_dFechaModifica, Cfl_cEquipoUser, Pan_cAnio, Cfl_cTransAutomatico, Cfl_cLEVenta, Cfl_cTransferencia, Cfl_cLECompra, Cfl_cAjusteNIF, Cfl_cVersionLE) VALUES("
                  
        Case "CParam2":
            sql = "INSERT INTO CNT_LIBRO_TIPODOC(" & _
                  "Emp_cCodigo, Pan_cAnio, Lib_cTipoLibro, Tdo_cCodigo, Opd_cEstado, Opd_cDeleted," & _
                  "Opd_cUserCrea, Opd_dFechaCrea, Opd_cUserModifica, Opd_dFechaModifica, Opd_cEquipoUser) VALUES("
            
        Case "COpera":
            sql = "INSERT INTO CNT_CONFIG_OPERA(" & _
                  "Emp_cCodigo, Cop_cCodigo, Cop_cDescripcion, Cop_cTipo, Cop_cEstado, Cop_cDeleted," & _
                  "Cop_cUserCrea, Cop_dFechaCrea, Cop_cUserModifica, Cop_dFechaModifica, Cop_cEquipoUser) VALUES("
            
        Case "COpera2":
            sql = "INSERT INTO CND_CONFIG_OPERA(" & _
                  "Emp_cCodigo, Pan_cAnio, Cop_cCodigo, Cod_cValorParam, Cod_nIgvPorc, Cod_cEstado, Cod_cDeleted," & _
                  "Cod_cUserCrea, Cod_dFechaCrea, Cod_cUserModifica, Cod_dFechaModifica, Cod_cEquipoUser) VALUES("
            
        Case "COpera3":
            sql = "INSERT INTO CNM_LIBRO_CTA(" & _
                  "Emp_cCodigo, Pan_cAnio, Lib_cTipoLibro, Pla_cCuentaContable, Lib_cEstado, Lib_cDeleted," & _
                  "Lib_cUserCrea, Lib_dFechaCrea, Lib_cUserModifica, Lib_dFechaModifica, Lib_cEquipoUser) VALUES("
            
        Case "CRatios":
            sql = "INSERT INTO CNT_CUENTA_INDI(" & _
                  "Emp_cCodigo, Ind_cCodCuenta, Ind_cDescripcion, Ind_cEstado, Ind_cDeleted," & _
                  "Ind_cUserCrea, Ind_dFechaCrea, Ind_cUserModifica, Ind_dFechaModifica, Ind_cEquipoUser) VALUES("
            
        Case "CMercaderias":
            sql = "INSERT INTO CNT_MERCADERIAS(" & _
                  "Emp_cCodigo,Pan_cAnio,Per_cPeriodo,Mer_cMetodo,Mer_cItem,Mer_cCodigo,Mer_cTipo,Mer_cDescrip,Mer_cMedida,Mer_nCantidad,Mer_nCosto,Mer_nTotal,Pla_cCuentaContable,Mer_cEstado," & _
                  "Mer_cDeleted,Mer_cUserCrea,Mer_dFechaCrea,Mer_cUserModifica,Mer_dFechaModifica,Mer_cEquipoUser) VALUES("
            
    End Select
    
    ExtraeCadenaSQL = sql
End Function

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

