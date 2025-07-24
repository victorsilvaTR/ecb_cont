VERSION 5.00
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmRepAnexoInvBalance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anexo al Libro de Inventarios y Balances"
   ClientHeight    =   3156
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   6300
   Icon            =   "frmRepAnexoInvBalance.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3156
   ScaleWidth      =   6300
   Begin VB.Frame fraTodo 
      Height          =   2955
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6120
      Begin VB.Frame fraVentas 
         Height          =   1095
         Left            =   1200
         TabIndex        =   9
         Top             =   1200
         Width           =   3825
         Begin VB.CheckBox chkReducidoVen 
            Caption         =   "Formato Reducido (*)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   15
            Top             =   600
            Width           =   3150
         End
         Begin VB.CheckBox chkFecVcmto 
            Caption         =   "Mostrar Fecha de Vencimiento"
            Height          =   285
            Left            =   240
            TabIndex        =   10
            Top             =   240
            Width           =   3120
         End
      End
      Begin VB.Frame fraCompras 
         Height          =   1095
         Left            =   1080
         TabIndex        =   5
         Top             =   1200
         Visible         =   0   'False
         Width           =   4065
         Begin VB.CheckBox chkReintegro 
            Caption         =   "Mostrar Columna de Reintegro"
            Height          =   195
            Left            =   240
            TabIndex        =   6
            Top             =   240
            Width           =   3630
         End
         Begin VB.CheckBox chkReducido 
            Caption         =   "Formato Reducido (*)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   14
            Top             =   720
            Width           =   3630
         End
         Begin VB.CheckBox chkComprobante 
            Caption         =   "Mostrar Columna de Comprob. No domiciliado"
            Height          =   195
            Left            =   240
            TabIndex        =   7
            Top             =   480
            Width           =   3630
         End
      End
      Begin VB.CheckBox ChkVerMes 
         Caption         =   "Ver del Mes"
         Height          =   330
         Left            =   1365
         TabIndex        =   24
         Top             =   1965
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Frame fraSimplificado 
         Height          =   735
         Left            =   1320
         TabIndex        =   21
         Top             =   1320
         Visible         =   0   'False
         Width           =   3645
         Begin TrueOleDBList70.TDBCombo tdbcTipoCta 
            Height          =   300
            Left            =   225
            TabIndex        =   22
            Tag             =   "_"
            Top             =   945
            Visible         =   0   'False
            Width           =   3330
            _ExtentX        =   5884
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=593"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=529"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=360"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=296"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(12)=   "Column(2).Width=826"
            Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=762"
            Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(16)=   "Column(2).Visible=0"
            Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(18)=   "Column(3).Width=1355"
            Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=1291"
            Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(22)=   "Column(3).Visible=0"
            Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
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
            EditFont        =   "Size=6.6,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
            LayoutName      =   ""
            LayoutFileName  =   ""
            MultipleLines   =   0
            EmptyRows       =   -1  'True
            CellTips        =   0
            EditHeight      =   276.095
            AutoSize        =   -1  'True
            GapHeight       =   36.283
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
            _PropDict       =   $"frmRepAnexoInvBalance.frx":0ECA
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
         Begin MSForms.CheckBox chkImpLibCajaResumido 
            Height          =   330
            Left            =   225
            TabIndex        =   23
            Top             =   270
            Width           =   3075
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "5424;582"
            Value           =   "0"
            Caption         =   "Libro Caja Resumido/Detallado"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
      End
      Begin VB.Frame fraExistencias 
         BorderStyle     =   0  'None
         Height          =   600
         Left            =   2640
         TabIndex        =   11
         Top             =   1440
         Visible         =   0   'False
         Width           =   3300
         Begin TrueOleDBList70.TDBCombo tdbcMetodo 
            Height          =   300
            Left            =   1200
            TabIndex        =   18
            Top             =   120
            Width           =   1935
            _ExtentX        =   3408
            _ExtentY        =   529
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            _DropdownWidth  =   4382
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=783"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=720"
            Splits(0)._ColumnProps(4)=   "Column(0)._VertColor=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=677"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=614"
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
            EditFont        =   "Size=6.6,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
            LayoutName      =   ""
            LayoutFileName  =   ""
            MultipleLines   =   0
            EmptyRows       =   -1  'True
            CellTips        =   0
            EditHeight      =   276.095
            AutoSize        =   -1  'True
            GapHeight       =   36.283
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
            _PropDict       =   $"frmRepAnexoInvBalance.frx":0F51
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
         Begin VB.Label lblMetodo 
            AutoSize        =   -1  'True
            Caption         =   "METODO"
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
            Left            =   135
            TabIndex        =   19
            Top             =   135
            Width           =   765
         End
      End
      Begin VB.CheckBox chkAnexos 
         Caption         =   "Imprimir Anexos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2655
         TabIndex        =   4
         Top             =   1200
         Visible         =   0   'False
         Width           =   2520
      End
      Begin TrueOleDBList70.TDBCombo tdbcMes 
         Height          =   300
         Left            =   2655
         TabIndex        =   0
         Top             =   405
         Width           =   1935
         _ExtentX        =   3408
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=635"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=572"
         Splits(0)._ColumnProps(4)=   "Column(0)._VertColor=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=677"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=614"
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
         EditFont        =   "Size=6.6,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         LayoutName      =   ""
         LayoutFileName  =   ""
         MultipleLines   =   0
         EmptyRows       =   -1  'True
         CellTips        =   0
         EditHeight      =   276.095
         AutoSize        =   -1  'True
         GapHeight       =   36.283
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
         _PropDict       =   $"frmRepAnexoInvBalance.frx":0FD8
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
      Begin TrueOleDBList70.TDBCombo tdbcMoneda 
         Height          =   300
         Left            =   2640
         TabIndex        =   25
         Tag             =   "_"
         Top             =   840
         Width           =   1935
         _ExtentX        =   3408
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=593"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=529"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=360"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=296"
         Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=826"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=762"
         Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(15)=   "Column(2).Visible=0"
         Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(17)=   "Column(3).Width=1355"
         Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=1291"
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
         EditFont        =   "Size=6.6,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         LayoutName      =   ""
         LayoutFileName  =   ""
         MultipleLines   =   0
         EmptyRows       =   -1  'True
         CellTips        =   0
         EditHeight      =   276.095
         AutoSize        =   -1  'True
         GapHeight       =   36.283
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
         _PropDict       =   $"frmRepAnexoInvBalance.frx":105F
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
      Begin VB.Frame fraHastaMes 
         Height          =   1095
         Left            =   1200
         TabIndex        =   12
         Top             =   1200
         Visible         =   0   'False
         Width           =   3705
         Begin VB.CheckBox chkHastaMes 
            Caption         =   "Hasta el mes seleccionado"
            Height          =   285
            Left            =   675
            TabIndex        =   13
            Top             =   450
            Width           =   2535
         End
      End
      Begin VB.Label Label1 
         Caption         =   "MONEDA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   26
         Top             =   880
         Width           =   855
      End
      Begin MSForms.CommandButton cmdSalir 
         Height          =   435
         Left            =   3165
         TabIndex        =   17
         Top             =   2385
         Width           =   1665
         Caption         =   " Salir"
         PicturePosition =   327683
         Size            =   "2937;767"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdImprimir 
         Height          =   435
         Left            =   1350
         TabIndex        =   16
         Top             =   2385
         Width           =   1665
         Caption         =   " Vista Previa"
         PicturePosition =   327683
         Size            =   "2937;767"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label lblRetenciones 
         AutoSize        =   -1  'True
         Caption         =   "Libro Retenciones Inc. E) y F) Art. 34° Ley de Imp. a la Renta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   465
         TabIndex        =   8
         Top             =   2070
         Visible         =   0   'False
         Width           =   5265
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "MES"
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
         Left            =   1590
         TabIndex        =   2
         Top             =   450
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
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
         Index           =   1
         Left            =   1590
         TabIndex        =   3
         Top             =   1035
         Width           =   45
      End
   End
   Begin VB.Label lblTitulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BALANCE DE COMPROBACION"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   4365
   End
End
Attribute VB_Name = "frmRepAnexoInvBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ReporteSunat As String
Public TituloSunat As String
Public CuentaInvBal As String

Dim gsGrupo As String
Dim iReport As Integer
Dim nAncho As Integer
Public rsArreglo  As New ADODB.Recordset
Dim rsArregloTotales  As New ADODB.Recordset
Dim NtoColumnas As Integer
Dim ColInicio As Integer
Dim i As Integer
Dim xCuentas, xLineasCuentas As String
Public ColAlmc, UltCol, Cont, PosReg As Integer
Public ContLineas As Integer

Dim ArrayCuentas(10) As String
Dim ArrayCuentasUS(18) As String

Dim Indice, j As Integer
Dim xCuentaTotal As String

Dim ImpMes As String * 13
Dim ImpAcum As String * 13
Dim LineaMes, LineaTotales As String

Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

Public Sub CerrarForm()
    Unload Me
End Sub

Private Sub chkReducido_Click()
    If chkReducido.Value = vbChecked Then
        chkReintegro.Value = vbUnchecked
        chkComprobante.Value = vbUnchecked
        ActivarControl chkReintegro, False, &H8000000F
        ActivarControl chkComprobante, False, &H8000000F
    Else
        ActivarControl chkReintegro, True, &H8000000F
        ActivarControl chkComprobante, True, &H8000000F
    End If
        
End Sub

Private Sub chkReducidoVen_Click()
    If chkReducidoVen.Value = vbChecked Then
        chkFecVcmto.Value = vbUnchecked
        ActivarControl chkFecVcmto, False, &H8000000F
    Else
        ActivarControl chkFecVcmto, True, &H8000000F
    End If
End Sub

Private Function ValidaCampos() As Boolean
    ValidaCampos = False

    If ReporteSunat = "F0502" And tdbcTipoCta.BoundText = "" Then
        Call Mensajes("Seleccione un tipo de cuenta")
        pSetFocus tdbcTipoCta
        Exit Function
    End If
    
    ValidaCampos = True
End Function
Private Sub cmdImprimir_Click()
    Dim matriz_fecha(25) As Variant
    Dim Tipo As String
    Dim formulas(0) As Variant
    
     'If ValidaCampos = False Then Exit Sub
'    App.P
    Screen.MousePointer = vbHourglass
    cmdImprimir.Enabled = False
    
    DoEvents
    
    If ReporteSunat = "" Then
        matriz_fecha(0) = "@Emp_cCodigo;" & gsEmpresa & ";True"
        matriz_fecha(1) = "@Pan_cAnio;" & gsAnio & ";True"
        matriz_fecha(2) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
        matriz_fecha(3) = "@moneda;" & tdbcMoneda.BoundText & ";True"
        matriz_fecha(4) = "@EMPRESA;" & gsEmpresaNom & ";True"
        matriz_fecha(5) = "@RUC;" & gsRUC & ";True"
        matriz_fecha(6) = "@Cta_BalInv;" & "" & ";True"
        matriz_fecha(7) = "@NOMMONEDA;" & tdbcMoneda.Text & ";True"
        matriz_fecha(8) = "@TITULO;" & CuentaInvBal & " - " & BuscaNombreCuenta(CuentaInvBal, False) & ";True"
        AbreReporteParam gsDSN, Me, rutaReportes & "RptAnexoLibroInvBalance.rpt", crptToWindow, "Anexo al Libro de Inventarios y Balances", "", matriz_fecha(), formulas()
    
    ElseIf ReporteSunat = "Detrac" Then
        matriz_fecha(0) = "@Tipo;CTAS_DETRAC;True"
        matriz_fecha(1) = "@Emp_cCodigo;" & gsEmpresa & ";True"
        matriz_fecha(2) = "@Pla_cAnioPlan;" & gsAnio & ";True"
        matriz_fecha(3) = "@Pla_cCuentaContable1;" & "" & ";True"
        matriz_fecha(4) = "@Pla_cCuentaContable2;" & "" & ";True"
        matriz_fecha(5) = "@Periodo;" & tdbcMes.BoundText & ";True"
        matriz_fecha(6) = "@Digitos;" & "0" & ";True"
        matriz_fecha(7) = "@HastaMes;" & chkHastaMes.Value & ";True"
        matriz_fecha(8) = "@EMPRESA;" & gsEmpresaNom & ";True"
        matriz_fecha(9) = "@RUC;" & gsRUC & ";True"
        matriz_fecha(10) = "@NOMMONEDA;" & tdbcMoneda.Text & ";True"
        matriz_fecha(11) = "@TITULO;" & CuentaInvBal & " - " & BuscaNombreCuenta(CuentaInvBal, False) & ";True"
        AbreReporteParam gsDSN, Me, rutaReportes & "RptCuentasDetraccion.rpt", crptToWindow, "Reporte de Detracciones", "", matriz_fecha(), formulas()
    
    ElseIf ReporteSunat = "F0301" Then
        '------------------------------
        nContadorProc = nContadorProc + 1
        '------------------------------
        matriz_fecha(0) = "@Tipo;BGE;True"
        matriz_fecha(1) = "@Emp_cCodigo;" & gsEmpresa & ";True"
        matriz_fecha(2) = "@Pan_cAnio;" & gsAnio & ";True"
        matriz_fecha(3) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
        matriz_fecha(4) = "@moneda;" & tdbcMoneda.BoundText & ";True"
        matriz_fecha(5) = "@EMPRESA;" & gsEmpresaNom & ";True"
        matriz_fecha(6) = "@RUC;" & gsRUC & ";True"
        matriz_fecha(7) = "@Reporte;;True"
        matriz_fecha(8) = "@NOMMONEDA;" & tdbcMoneda.Text & ";True"
        matriz_fecha(9) = "@TITULO;" & CuentaInvBal & " - " & BuscaNombreCuenta(CuentaInvBal, False) & ";True"
        matriz_fecha(10) = "@Mes;" & IIf(ChkVerMes.Value = 1, 1, 0) & ";True"
        
        If chkAnexos.Value = "1" Then
            NombreReporte = "F0301"
            matriz_fecha(7) = "@Reporte;DETALLE;True"
            AbreReporteParam gsDSN, Me, rutaReportes & "RptBalanceGeneralConasevDetalle.rpt", crptToWindow, "Anexo al Libro de Inventarios y Balances", "", matriz_fecha(), formulas()
        Else
            matriz_fecha(7) = "@Reporte;RESUMEN;True"
            AbreReporteParam gsDSN, Me, rutaReportes & "RptFormato_0301.rpt", crptToWindow, "Anexo al Libro de Inventarios y Balances", "", matriz_fecha(), formulas()
        End If
        NombreReporte = ""
    ElseIf ReporteSunat = "mnuF0307_2021" Then
        
        matriz_fecha(0) = "@Accion;REPORTE_REVISADO;True"
        matriz_fecha(1) = "@Emp_cCodigo;" & gsEmpresa & ";True"
        matriz_fecha(2) = "@Pan_cAnio;" & gsAnio & ";True"
        matriz_fecha(3) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
        matriz_fecha(4) = "@Mer_cMetodo;" & tdbcMetodo.BoundText & ";True"
        matriz_fecha(5) = "@EMPRESA;" & gsEmpresaNom & ";True"
        matriz_fecha(6) = "@RUC;" & gsRUC & ";True"
        matriz_fecha(7) = "@NOMMONEDA;" & tdbcMoneda.Text & ";True"
        matriz_fecha(8) = "@TITULO;" & CuentaInvBal & " - " & BuscaNombreCuenta(CuentaInvBal, False) & ";True"
        'AbreReporteParam gsDSN, Me, rutaReportes & "RptFormato_PCGE2X.rpt", crptToWindow, "Anexo al Libro de Inventarios y Balances", "", matriz_fecha(), formulas()
        AbreReporteParam gsDSN, Me, rutaReportes & "RptFormato_0307_2021.rpt", crptToWindow, "Anexo al Libro de Inventarios y Balances", "", matriz_fecha(), formulas()

    ElseIf ReporteSunat = "F0307_2" Then
        
        matriz_fecha(0) = "@Emp_cCodigo;" & gsEmpresa & ";True"
        matriz_fecha(1) = "@Pan_cAnio;" & gsAnio & ";True"
        matriz_fecha(2) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
        matriz_fecha(3) = "@moneda;" & tdbcMoneda.BoundText & ";True"
        matriz_fecha(4) = "@EMPRESA;" & gsEmpresaNom & ";True"
        matriz_fecha(5) = "@RUC;" & gsRUC & ";True"
        matriz_fecha(6) = "@Cta_BalInv;" & CuentaInvBal & ";True"
        matriz_fecha(7) = "@NOMMONEDA;" & tdbcMoneda.Text & ";True"
        matriz_fecha(8) = "@TITULO;" & CuentaInvBal & " - " & BuscaNombreCuenta(CuentaInvBal, False) & ";True"
        
        AbreReporteParam gsDSN, Me, rutaReportes & "RptFormato_0307_" & CuentaInvBal & ".rpt", crptToWindow, "Anexo al Libro de Inventarios y Balances", "", matriz_fecha(), formulas()
    ElseIf ReporteSunat = "F0031" Then
        matriz_fecha(0) = "@Tipo;REPORTE;True"
        matriz_fecha(1) = "@Emp_cCodigo;" & gsEmpresa & ";True"
        matriz_fecha(2) = "@Pan_cAnio;" & gsAnio & ";True"
        matriz_fecha(3) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
        matriz_fecha(4) = "@EMPRESA;" & gsEmpresaNom & ";True"
        matriz_fecha(5) = "@RUC;" & gsRUC & ";True"
        matriz_fecha(6) = "@Ase_nVoucher;;True"
        matriz_fecha(7) = "@Asd_nItem;;True"
        matriz_fecha(8) = "@Ten_cTipoEntidad;;True"
        matriz_fecha(9) = "@Ent_cCodentidad;;True"
        matriz_fecha(10) = "@Val_cTitulo;;True"
        matriz_fecha(11) = "@Val_cDesTitulo;;True"
        matriz_fecha(12) = "@Val_nValorNom;0;True"
        matriz_fecha(13) = "@Val_nCantidad;0;True"
        matriz_fecha(14) = "@Val_nCostoTot;0;True"
        matriz_fecha(15) = "@Val_nProvTot;True"
        matriz_fecha(16) = "@Val_nTotalNeto;0;True"
        matriz_fecha(17) = "@NOMMONEDA;" & tdbcMoneda.Text & ";True"
        matriz_fecha(18) = "@TITULO;" & CuentaInvBal & " - " & BuscaNombreCuenta(CuentaInvBal, False) & ";True"
        matriz_fecha(19) = "@Cuenta;" & Right(ReporteSunat, 2) & ";True"
        AbreReporteParam gsDSN, Me, rutaReportes & "RptFormato_0031.rpt", crptToWindow, "Anexo al Libro de Inventarios y Balances", "", matriz_fecha(), formulas()

    ElseIf Len(ReporteSunat) = 5 And ReporteSunat >= "F0302" And ReporteSunat <= "F0315" And ReporteSunat <> "F0308" And ReporteSunat <> "F0310" Then
        matriz_fecha(0) = "@Emp_cCodigo;" & gsEmpresa & ";True"
        matriz_fecha(1) = "@Pan_cAnio;" & gsAnio & ";True"
        matriz_fecha(2) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
        matriz_fecha(3) = "@moneda;" & tdbcMoneda.BoundText & ";True"
        matriz_fecha(4) = "@EMPRESA;" & gsEmpresaNom & ";True"
        matriz_fecha(5) = "@RUC;" & gsRUC & ";True"
        matriz_fecha(6) = "@Cta_BalInv;" & CuentaInvBal & ";True"
        matriz_fecha(7) = "@NOMMONEDA;" & tdbcMoneda.Text & ";True"
        matriz_fecha(8) = "@TITULO;" & CuentaInvBal & " - " & BuscaNombreCuenta(CuentaInvBal, False) & ";True"
        AbreReporteParam gsDSN, Me, rutaReportes & "RptFormato_" & Right(ReporteSunat, 4) & ".rpt", crptToWindow, "Anexo al Libro de Inventarios y Balances", "", matriz_fecha(), formulas()

    ElseIf ReporteSunat = "F0309_33" Then
        matriz_fecha(0) = "@Accion;REPORTE;True"
        matriz_fecha(1) = "@Emp_cCodigo;" & gsEmpresa & ";True"
        matriz_fecha(2) = "@Pan_cAnio;" & gsAnio & ";True"
        matriz_fecha(3) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
        matriz_fecha(4) = "@EMPRESA;" & gsEmpresaNom & ";True"
        matriz_fecha(5) = "@RUC;" & gsRUC & ";True"
        matriz_fecha(6) = "@Cta_BalInv;" & CuentaInvBal & ";True"
        matriz_fecha(7) = "@moneda;" & tdbcMoneda.BoundText & ";True"
        matriz_fecha(8) = "@NOMMONEDA;" & tdbcMoneda.Text & ";True"
        matriz_fecha(9) = "@TITULO;" & CuentaInvBal & " - " & BuscaNombreCuenta(CuentaInvBal, False) & ";True"
        AbreReporteParam gsDSN, Me, rutaReportes & "RptFormato_0309_33.rpt", crptToWindow, "Anexo al Libro de Inventarios y Balances", "", matriz_fecha(), formulas()
    ElseIf ReporteSunat = "F0309_34" Then
        matriz_fecha(0) = "@Accion;REPORTE;True"
        matriz_fecha(1) = "@Emp_cCodigo;" & gsEmpresa & ";True"
        matriz_fecha(2) = "@Pan_cAnio;" & gsAnio & ";True"
        matriz_fecha(3) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
        matriz_fecha(4) = "@EMPRESA;" & gsEmpresaNom & ";True"
        matriz_fecha(5) = "@RUC;" & gsRUC & ";True"
        matriz_fecha(6) = "@NOMMONEDA;" & tdbcMoneda.Text & ";True"
        matriz_fecha(7) = "@TITULO;" & CuentaInvBal & " - " & BuscaNombreCuenta(CuentaInvBal, False) & ";True"
        AbreReporteParam gsDSN, Me, rutaReportes & "RptFormato_0309_34.rpt", crptToWindow, "Anexo al Libro de Inventarios y Balances", "", matriz_fecha(), formulas()
        
    ElseIf ReporteSunat = "F0309_35" Then
        matriz_fecha(0) = "@Accion;REPORTE;True"
        matriz_fecha(1) = "@Emp_cCodigo;" & gsEmpresa & ";True"
        matriz_fecha(2) = "@Pan_cAnio;" & gsAnio & ";True"
        matriz_fecha(3) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
        matriz_fecha(4) = "@moneda;" & tdbcMoneda.BoundText & ";True"
        matriz_fecha(5) = "@EMPRESA;" & gsEmpresaNom & ";True"
        matriz_fecha(6) = "@RUC;" & gsRUC & ";True"
        matriz_fecha(7) = "@Cta_BalInv;35;True"
        matriz_fecha(8) = "@NOMMONEDA;" & tdbcMoneda.Text & ";True"
        matriz_fecha(9) = "@TITULO;" & CuentaInvBal & " - " & BuscaNombreCuenta(CuentaInvBal, False) & ";True"
        AbreReporteParam gsDSN, Me, rutaReportes & "RptFormato_0309_35.rpt", crptToWindow, "Anexo al Libro de Inventarios y Balances", "", matriz_fecha(), formulas()
        
    ElseIf ReporteSunat >= "F0309_38" And ReporteSunat <= "F0309_45" Then
        matriz_fecha(0) = "@Emp_cCodigo;" & gsEmpresa & ";True"
        matriz_fecha(1) = "@Pan_cAnio;" & gsAnio & ";True"
        matriz_fecha(2) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
        matriz_fecha(3) = "@moneda;" & tdbcMoneda.BoundText & ";True"
        matriz_fecha(4) = "@EMPRESA;" & gsEmpresaNom & ";True"
        matriz_fecha(5) = "@RUC;" & gsRUC & ";True"
        matriz_fecha(6) = "@Cta_BalInv;" & CuentaInvBal & ";True"
        matriz_fecha(7) = "@NOMMONEDA;" & tdbcMoneda.Text & ";True"
        matriz_fecha(8) = "@TITULO;" & CuentaInvBal & " - " & BuscaNombreCuenta(CuentaInvBal, False) & ";True"
        AbreReporteParam gsDSN, Me, rutaReportes & "RptFormato_0309_" & CuentaInvBal & ".rpt", crptToWindow, "Anexo al Libro de Inventarios y Balances", "", matriz_fecha(), formulas()
    
    ElseIf ReporteSunat = "F0310" Then
        matriz_fecha(0) = "@Emp_cCodigo;" & gsEmpresa & ";True"
        matriz_fecha(1) = "@Pan_cAnio;" & gsAnio & ";True"
        matriz_fecha(2) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
        matriz_fecha(3) = "@CuentaInicio;40;True"
        matriz_fecha(4) = "@CuentaFin;40999999;True"
        matriz_fecha(5) = "@Moneda;" & tdbcMoneda.BoundText & ";True"
        matriz_fecha(6) = "@EMPRESA;" & gsEmpresaNom & ";True"
        matriz_fecha(7) = "@RUC;" & gsRUC & ";True"
        matriz_fecha(8) = "@NOMMONEDA;" & tdbcMoneda.Text & ";True"
        matriz_fecha(9) = "@TITULO;" & CuentaInvBal & " - " & BuscaNombreCuenta(CuentaInvBal, False) & ";True"
        AbreReporteParam gsDSN, Me, rutaReportes & "RptFormato_0310.rpt", crptToWindow, "Anexo al Libro de Inventarios y Balances", "", matriz_fecha(), formulas()
    
    ElseIf ReporteSunat = "F0316" Then
        matriz_fecha(0) = "@Emp_cCodigo;" & gsEmpresa & ";True"
        matriz_fecha(1) = "@Pan_cAnio;" & gsAnio & ";True"
        matriz_fecha(2) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
        matriz_fecha(3) = "@ULTIMODIA;" & UltimoDiaMes(tdbcMes.BoundText, gsAnio) & ";True"
        matriz_fecha(4) = "@EMPRESA;" & gsEmpresaNom & ";True"
        matriz_fecha(5) = "@RUC;" & gsRUC & ";True"
        matriz_fecha(6) = "@Tipo;REPORTE;True"
        matriz_fecha(7) = "@Ase_nVoucher;;True"
        matriz_fecha(8) = "@Asd_nItem;;True"
        matriz_fecha(9) = "@Ten_cTipoEntidad;;True"
        matriz_fecha(10) = "@Ent_cCodentidad;;True"
        matriz_fecha(11) = "@Cap_cAcciones;;True"
        matriz_fecha(12) = "@Cap_nImportes;0;True"
        matriz_fecha(13) = "@Cap_nValorNom;0;True"
        matriz_fecha(14) = "@Cap_nASuscritas;0;True"
        matriz_fecha(15) = "@Cap_nAPagadas;0;True"
        matriz_fecha(16) = "@Cap_nAcciones;0;True"
        matriz_fecha(17) = "@Cap_nPorcent;0;True"
        matriz_fecha(18) = "@NOMMONEDA;" & tdbcMoneda.Text & ";True"
        matriz_fecha(19) = "@TITULO;" & CuentaInvBal & " - " & BuscaNombreCuenta(CuentaInvBal, False) & ";True"
        
        AbreReporteParam gsDSN, Me, rutaReportes & "RptFormato_0316.rpt", crptToWindow, "Anexo al Libro de Inventarios y Balances", "", matriz_fecha(), formulas()
    
    ElseIf ReporteSunat = "F0320" Then
        matriz_fecha(0) = "@Tipo;FUN;True"
        matriz_fecha(1) = "@Emp_cCodigo;" & gsEmpresa & ";True"
        matriz_fecha(2) = "@Pan_cAnio;" & gsAnio & ";True"
        matriz_fecha(3) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
        matriz_fecha(4) = "@moneda;" & tdbcMoneda.BoundText & ";True"
        matriz_fecha(5) = "@EMPRESA;" & gsEmpresaNom & ";True"
        matriz_fecha(6) = "@RUC;" & gsRUC & ";True"
        matriz_fecha(7) = "@Reporte;;True"
        matriz_fecha(8) = "@NOMMONEDA;" & tdbcMoneda.Text & ";True"
        matriz_fecha(9) = "@UltimoDia;" & UltimoDiaMes(tdbcMes.BoundText, gsAnio) & ";True"
        matriz_fecha(10) = "@TITULO;" & CuentaInvBal & " - " & BuscaNombreCuenta(CuentaInvBal, False) & ";True"
        matriz_fecha(11) = "@Mes;" & IIf(ChkVerMes.Value = 1, 1, 0) & ";True"
        
        If chkAnexos.Value = "1" Then
            matriz_fecha(7) = "@Reporte;DETALLE;True"
            AbreReporteParam gsDSN, Me, rutaReportes & "RptBalanceGeneralConasevDetalle.rpt", crptToWindow, "Anexo al Libro de Inventarios y Balances", "", matriz_fecha(), formulas()
        Else
            matriz_fecha(7) = "@Reporte;RESUMEN;True"
            AbreReporteParam gsDSN, Me, rutaReportes & "RptFormato_0320.rpt", crptToWindow, "Anexo al Libro de Inventarios y Balances", "", matriz_fecha(), formulas()
        End If
    
    ElseIf ReporteSunat = "F0321" Then
        matriz_fecha(0) = "@Tipo;NAT;True"
        matriz_fecha(1) = "@Emp_cCodigo;" & gsEmpresa & ";True"
        matriz_fecha(2) = "@Pan_cAnio;" & gsAnio & ";True"
        matriz_fecha(3) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
        matriz_fecha(4) = "@moneda;" & tdbcMoneda.BoundText & ";True"
        matriz_fecha(5) = "@EMPRESA;" & gsEmpresaNom & ";True"
        matriz_fecha(6) = "@RUC;" & gsRUC & ";True"
        matriz_fecha(7) = "@Reporte;;True"
        matriz_fecha(8) = "@NOMMONEDA;" & tdbcMoneda.Text & ";True"
        matriz_fecha(9) = "@UltimoDia;" & UltimoDiaMes(tdbcMes.BoundText, gsAnio) & ";True"
        matriz_fecha(10) = "@TITULO;" & CuentaInvBal & " - " & BuscaNombreCuenta(CuentaInvBal, False) & ";True"
        matriz_fecha(11) = "@Mes;" & IIf(ChkVerMes.Value = 1, 1, 0) & ";True"
        
        If chkAnexos.Value = "1" Then
            matriz_fecha(7) = "@Reporte;DETALLE;True"
            AbreReporteParam gsDSN, Me, rutaReportes & "RptBalanceGeneralConasevDetalle.rpt", crptToWindow, "Anexo al Libro de Inventarios y Balances", "", matriz_fecha(), formulas()
        Else
            matriz_fecha(7) = "@Reporte;RESUMEN;True"
            AbreReporteParam gsDSN, Me, rutaReportes & "RptFormato_0321.rpt", crptToWindow, "Anexo al Libro de Inventarios y Balances", "", matriz_fecha(), formulas()
        End If
        
    ElseIf ReporteSunat >= "F0350" And ReporteSunat <= "F0359" Then
        matriz_fecha(0) = "@Emp_cCodigo;" & gsEmpresa & ";True"
        matriz_fecha(1) = "@Pan_cAnio;" & gsAnio & ";True"
        matriz_fecha(2) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
        matriz_fecha(3) = "@moneda;" & tdbcMoneda.BoundText & ";True"
        matriz_fecha(4) = "@EMPRESA;" & gsEmpresaNom & ";True"
        matriz_fecha(5) = "@RUC;" & gsRUC & ";True"
        matriz_fecha(6) = "@Cta_BalInv;" & CuentaInvBal & ";True"
        matriz_fecha(7) = "@NOMMONEDA;" & tdbcMoneda.Text & ";True"
        matriz_fecha(8) = "@TITULO;" & CuentaInvBal & " - " & BuscaNombreCuenta(CuentaInvBal, False) & ";True"
        
        AbreReporteParam gsDSN, Me, rutaReportes & "RptFormato_0309_" & CuentaInvBal & ".rpt", crptToWindow, "Anexo al Libro de Inventarios y Balances", "", matriz_fecha(), formulas()
    
    ElseIf ReporteSunat = "F0401" Then
        matriz_fecha(0) = "@Emp_cCodigo;" & gsEmpresa & ";True"
        matriz_fecha(1) = "@Pan_cAnio;" & gsAnio & ";True"
        matriz_fecha(2) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
        matriz_fecha(3) = "@moneda;" & tdbcMoneda.BoundText & ";True"
        matriz_fecha(4) = "@EMPRESA;" & gsEmpresaNom & ";True"
        matriz_fecha(5) = "@RUC;" & gsRUC & ";True"
        matriz_fecha(6) = "@Cta_BalInv;" & CuentaInvBal & ";True"
        matriz_fecha(7) = "@NOMMONEDA;" & tdbcMoneda.Text & ";True"
        matriz_fecha(8) = "@TITULO;" & CuentaInvBal & " - " & BuscaNombreCuenta(CuentaInvBal, False) & ";True"
        
        AbreReporteParam gsDSN, Me, rutaReportes & "RptFormato_0401.rpt", crptToWindow, "Libro de Retenciones", "", matriz_fecha(), formulas()
    
    ElseIf ReporteSunat = "F0502" Then
        gsNombreVista = "Diario Simplificado"
        If ValidaParametrosDiario Then Screen.MousePointer = vbDefault: cmdImprimir.Enabled = True: Exit Sub
'Tata-009 Cambio de reporte a Cristal Report
        matriz_fecha(0) = "@Emp_cCodigo;" & gsEmpresa & ";True"
        matriz_fecha(1) = "@Pan_cAnio;" & gsAnio & ";True"
        matriz_fecha(2) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
        matriz_fecha(3) = "@moneda;" & tdbcMoneda.BoundText & ";True"
        matriz_fecha(4) = "@EMPRESA;" & gsEmpresaNom & ";True"
        matriz_fecha(5) = "@RUC;" & gsRUC & ";True"
        matriz_fecha(6) = "@Cta_BalInv;" & CuentaInvBal & ";True"
        matriz_fecha(7) = "@NOMMONEDA;" & tdbcMoneda.Text & ";True"
        matriz_fecha(8) = "@TITULO;" & CuentaInvBal & " - " & BuscaNombreCuenta(CuentaInvBal, False) & ";True"
        sSql = "spCn_ReprocesoDiarioSimp_2D '" & gsEmpresa & "','" & gsAnio & "','" & tdbcMes.BoundText & "',1," & gsNumDigDiarioSimpRep & _
         ",'" & tdbcMoneda.Text & "','" & gsTipoImp & "'"
        gcnSistema.ConnectionTimeout = 0
        gcnSistema.CommandTimeout = 0
        Call Conectar
        gcnSistema.Execute sSql
        Call Desconectar
        AbreReporteParam gsDSN, Me, rutaReportes & "RptDiarioSimplificado.rpt", crptToWindow, "Libro Diario Simplificado", "", matriz_fecha(), formulas()

'Si se quiere volver al informe antiguo se debe comentar desde tata-009 y descomentar Imprimir
'        Imprimir
'Fin Tata-009
    ElseIf ReporteSunat = "F0801" Then
        matriz_fecha(0) = "@Emp_cCodigo;" & gsEmpresa & ";True"
        matriz_fecha(1) = "@Pan_cAnio;" & gsAnio & ";True"
        matriz_fecha(2) = "@desde;" & PrimerDiaMes(tdbcMes.BoundText, gsAnio) & ";True"
        matriz_fecha(3) = "@hasta;" & UltimoDiaMes(tdbcMes.BoundText, gsAnio) & ";True"
        matriz_fecha(4) = "@moneda;" & tdbcMoneda.BoundText & ";True"
        matriz_fecha(5) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
        matriz_fecha(6) = "@RUC;" & gsRUC & ";True"
        matriz_fecha(7) = "@EMPRESA;" & gsEmpresaNom & ";True"
        matriz_fecha(8) = "@NOMMONEDA;" & tdbcMoneda.Text & ";True"
        matriz_fecha(9) = "@TITULO;" & CuentaInvBal & " - " & BuscaNombreCuenta(CuentaInvBal, False) & ";True"

        If chkReintegro.Value = vbChecked And chkComprobante.Value = vbUnchecked Then
            AbreReporteParam gsDSN, Me, rutaReportes & "RptFormato_0801_F02.rpt", crptToWindow, "Registro de Compras", "", matriz_fecha(), formulas(), Orientacion_Pagina.Vertical, Tipo_Pagina.USA
        ElseIf chkReintegro.Value = vbUnchecked And chkComprobante.Value = vbChecked Then
            AbreReporteParam gsDSN, Me, rutaReportes & "RptFormato_0801_F01.rpt", crptToWindow, "Registro de Compras", "", matriz_fecha(), formulas(), Orientacion_Pagina.Vertical, Tipo_Pagina.USA
        ElseIf chkReintegro.Value = vbUnchecked And chkComprobante.Value = vbUnchecked And chkReducido.Value = vbUnchecked Then
            AbreReporteParam gsDSN, Me, rutaReportes & "RptFormato_0801_F03.rpt", crptToWindow, "Registro de Compras", "", matriz_fecha(), formulas(), Orientacion_Pagina.Vertical, Tipo_Pagina.USA
        ElseIf chkReducido.Value = vbChecked Then
            AbreReporteParam gsDSN, Me, rutaReportes & "RptFormato_0801_F04.rpt", crptToWindow, "Registro de Compras", "", matriz_fecha(), formulas(), Orientacion_Pagina.Horizontal, Tipo_Pagina.A4
        Else
            AbreReporteParam gsDSN, Me, rutaReportes & "RptFormato_0801.rpt", crptToWindow, "Registro de Compras", "", matriz_fecha(), formulas(), Orientacion_Pagina.Vertical, Tipo_Pagina.USA
        End If
    ElseIf ReporteSunat = "F1001" Then
        matriz_fecha(0) = "@Emp_cCodigo;" & gsEmpresa & ";True"
        matriz_fecha(1) = "@Pan_cAnio;" & gsAnio & ";True"
        matriz_fecha(2) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
        matriz_fecha(3) = "@Moneda;" & tdbcMoneda.BoundText & ";True"
        matriz_fecha(4) = "@EMPRESA;" & gsEmpresaNom & ";True"
        matriz_fecha(5) = "@RUC;" & gsRUC & ";True"
        matriz_fecha(6) = "@NOMMONEDA;" & tdbcMoneda.Text & ";True"
        matriz_fecha(7) = "@TITULO;" & CuentaInvBal & " - " & BuscaNombreCuenta(CuentaInvBal, False) & ";True"
        AbreReporteParam gsDSN, Me, rutaReportes & "RptFormato_1001.rpt", crptToWindow, "Anexo al Libro de Inventarios y Balances", "", matriz_fecha(), formulas()
    ElseIf ReporteSunat = "F1002" Then
        matriz_fecha(0) = "@Emp_cCodigo;" & gsEmpresa & ";True"
        matriz_fecha(1) = "@Pan_cAnio;" & gsAnio & ";True"
        matriz_fecha(2) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
        matriz_fecha(3) = "@Moneda;" & tdbcMoneda.BoundText & ";True"
        matriz_fecha(4) = "@EMPRESA;" & gsEmpresaNom & ";True"
        matriz_fecha(5) = "@RUC;" & gsRUC & ";True"
        matriz_fecha(6) = "@NOMMONEDA;" & tdbcMoneda.Text & ";True"
        matriz_fecha(7) = "@TITULO;" & CuentaInvBal & " - " & BuscaNombreCuenta(CuentaInvBal, False) & ";True"
        AbreReporteParam gsDSN, Me, rutaReportes & "RptFormato_1002.rpt", crptToWindow, "Anexo al Libro de Inventarios y Balances", "", matriz_fecha(), formulas()
    ElseIf ReporteSunat = "F1003" Then
        matriz_fecha(0) = "@Emp_cCodigo;" & gsEmpresa & ";True"
        matriz_fecha(1) = "@Pan_cAnio;" & gsAnio & ";True"
        matriz_fecha(2) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
        matriz_fecha(3) = "@Moneda;" & tdbcMoneda.BoundText & ";True"
        matriz_fecha(4) = "@EMPRESA;" & gsEmpresaNom & ";True"
        matriz_fecha(5) = "@RUC;" & gsRUC & ";True"
        matriz_fecha(6) = "@NOMMONEDA;" & tdbcMoneda.Text & ";True"
        matriz_fecha(7) = "@TITULO;" & CuentaInvBal & " - " & BuscaNombreCuenta(CuentaInvBal, False) & ";True"
        
        Dim sql As String
        Dim rs As ADODB.Recordset
        sql = "SELECT Tab_cCodigo, Tab_cDescripCampo FROM TABLA " & _
                    "WHERE Emp_cCodigo =  '" & gsEmpresa & "' AND Tab_cTabla = '068' ORDER BY Tab_cCodigo "
        LlenarRecordSet sql, rs
        
        If Not rs Is Nothing Then
            If rs.RecordCount < 9 Then
                AbreReporteParam gsDSN, Me, rutaReportes & "RptFormato_1003v.rpt", crptToWindow, "Anexo al Libro de Inventarios y Balances", "", matriz_fecha(), formulas()
            Else
                AbreReporteParam gsDSN, Me, rutaReportes & "RptFormato_1003.rpt", crptToWindow, "Anexo al Libro de Inventarios y Balances", "", matriz_fecha(), formulas()
            End If
        End If
        
    ElseIf ReporteSunat = "F1401" Then
        matriz_fecha(0) = "@Emp_cCodigo;" & gsEmpresa & ";True"
        matriz_fecha(1) = "@Pan_cAnio;" & gsAnio & ";True"
        matriz_fecha(2) = "@desde;" & PrimerDiaMes(tdbcMes.BoundText, gsAnio) & ";True"
        matriz_fecha(3) = "@hasta;" & UltimoDiaMes(tdbcMes.BoundText, gsAnio) & ";True"
        matriz_fecha(4) = "@moneda;" & tdbcMoneda.BoundText & ";True"
        matriz_fecha(5) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
        matriz_fecha(6) = "@RUC;" & gsRUC & ";True"
        matriz_fecha(7) = "@EMPRESA;" & gsEmpresaNom & ";True"
        matriz_fecha(8) = "@NOMMONEDA;" & tdbcMoneda.Text & ";True"
        matriz_fecha(9) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
        matriz_fecha(10) = "@TITULO;" & CuentaInvBal & " - " & BuscaNombreCuenta(CuentaInvBal, False) & ";True"
        
        If chkFecVcmto.Value = vbUnchecked And chkReducidoVen.Value = vbUnchecked Then
            AbreReporteParam gsDSN, Me, rutaReportes & "RptFormato_1401_F01.rpt", crptToWindow, "Registro de Ventas", "", matriz_fecha(), formulas(), Orientacion_Pagina.Vertical, Tipo_Pagina.USA
        End If
        
        If chkFecVcmto.Value = vbUnchecked And chkReducidoVen.Value = vbChecked Then
'            AbreReporteParam gsDSN, Me, rutaReportes & "RptFormato_1401_F02.rpt", crptToWindow, "Registro de Ventas", "", matriz_fecha(), formulas(), Orientacion_Pagina.Vertical, Tipo_Pagina.A4
            AbreReporteParam gsDSN, Me, rutaReportes & "RptFormato_1401_F02.rpt", crptToWindow, "Registro de Ventas", "", matriz_fecha(), formulas(), Orientacion_Pagina.Horizontal, Tipo_Pagina.A4
        End If
        
        If chkFecVcmto.Value = vbChecked And chkReducidoVen.Value = vbUnchecked Then
            AbreReporteParam gsDSN, Me, rutaReportes & "RptFormato_1401.rpt", crptToWindow, "Registro de Ventas", "", matriz_fecha(), formulas(), Orientacion_Pagina.Vertical, Tipo_Pagina.USA
        End If
    ElseIf ReporteSunat = "EFE" Then 'frt_efe
        '------------------------------
        nContadorProc = nContadorProc + 1
        '------------------------------
        matriz_fecha(0) = "@Tipo;EFE;True"
        matriz_fecha(1) = "@Emp_cCodigo;" & gsEmpresa & ";True"
        matriz_fecha(2) = "@Pan_cAnio;" & gsAnio & ";True"
        matriz_fecha(3) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
        matriz_fecha(4) = "@moneda;" & tdbcMoneda.BoundText & ";True"
        matriz_fecha(5) = "@EMPRESA;" & gsEmpresaNom & ";True"
        matriz_fecha(6) = "@RUC;" & gsRUC & ";True"
        matriz_fecha(7) = "@Reporte;;True"
        matriz_fecha(8) = "@NOMMONEDA;" & tdbcMoneda.Text & ";True"
        matriz_fecha(9) = "@TITULO;" & CuentaInvBal & " - " & BuscaNombreCuenta(CuentaInvBal, False) & ";True"
        matriz_fecha(10) = "@Mes;" & IIf(ChkVerMes.Value = 1, 1, 0) & ";True"
        
        If chkAnexos.Value = "1" Then
            NombreReporte = "F0301"
            matriz_fecha(7) = "@Reporte;DETALLE;True"
            AbreReporteParam gsDSN, Me, rutaReportes & "RptFlujoEfectivoDetalle.rpt", crptToWindow, "Anexo Estado de Flujo de Efectivo", "", matriz_fecha(), formulas()
        Else
            matriz_fecha(7) = "@Reporte;RESUMEN;True"
            AbreReporteParam gsDSN, Me, rutaReportes & "RptFlujoEfectivoResumen.rpt", crptToWindow, "Estado de Flujo de Efectivo", "", matriz_fecha(), formulas()
        End If
        NombreReporte = ""
    End If
    Screen.MousePointer = vbDefault
    cmdImprimir.Enabled = True
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    Me.Caption = Titulo(Me.Caption, TituloSunat)
    
    DoEvents
    
    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))
    
    chkAnexos.Visible = False
    lblRetenciones.Visible = False
    
    Call Centrar_form(Me)

    Call LlenaCombos
    Call BuscarMonedaNacional
    
    If InStr(1, TituloSunat, "3.8") > 0 Or InStr(1, TituloSunat, "3.16") > 0 Or InStr(1, TituloSunat, "10.1") > 0 Then
        tdbcMoneda.Locked = True
    End If
    
    If ReporteSunat = "mnuF0307_2021" Then
        fraExistencias.Visible = True
    Else
        fraExistencias.Visible = False
    End If
        
    If ReporteSunat = "F0502" Then
      '  fraSimplificado.Visible = True ' para el diario simplificado detalle se reviso que no lo consideraba
    Else
        fraSimplificado.Visible = False
    End If
    
    If ReporteSunat = "F0801" Then
        fraCompras.Visible = True
    Else
        fraCompras.Visible = False
    End If
    
    If ReporteSunat = "F1401" Then
        fraVentas.Visible = True
    Else
        fraVentas.Visible = False
    End If
    
    If ReporteSunat = "F0301" Or ReporteSunat = "F0320" Or ReporteSunat = "F0321" Or ReporteSunat = "EFE" Then 'frt_efe
        chkAnexos.Visible = True
    Else
        chkAnexos.Visible = False
    End If
    
    If ReporteSunat = "F0502" Then chkImpLibCajaResumido.Value = 0
    tdbcMoneda.BoundText = 0
    
End Sub

Private Sub LlenaCombos()
    Dim sqlcombos As String
    Dim entro As Boolean
    
    entro = False
    
    If ReporteSunat = "F0801" Or ReporteSunat = "F1401" Then
        Call LlenaComboMesAddItem(tdbcMes)
        entro = True
    Else
        Call LlenaComboMesApeAddItem(tdbcMes)
    End If
    
    
    DoEvents
    tdbcMes.ReBind
    
    If entro = False Then
        If gsPeriodo = "" Then
            tdbcMes.BoundText = "00"
        Else
            tdbcMes.BoundText = gsPeriodo
        End If
    Else
        If gsPeriodo > "00" And gsPeriodo < "13" Then
            tdbcMes.BoundText = gsPeriodo
        Else
            tdbcMes.BoundText = "01"
        End If
    
    End If
    
    
    
    DoEvents
    ' *** Llenando el tipo de Moneda
    sqlcombos = "SELECT Mon_cCodigo, Mon_cNombreLargo, Mon_cMNac, Mon_cMExt From CNT_TIPO_MONEDA " & _
                "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND (Mon_cMNac = '1' or Mon_cMExt = '1') " & _
                "ORDER BY Mon_cNombreLargo"
    LlenarComboAddItem tdbcMoneda, sqlcombos
    
    '---------------------------
    
    sqlcombos = "SELECT Tab_cCodigo, Tab_cDescripCampo FROM TABLA WITH(NOLOCK) " & _
                "WHERE Emp_cCodigo='" & gsEmpresa & "' AND Tab_cTabla = '085' " & _
                "ORDER BY Tab_cCodigo"
                
    LlenarComboAddItem tdbcMetodo, sqlcombos
    
    '---------------------------
    
    tdbcTipoCta.Clear
    tdbcTipoCta.AddItem "" + ";" + "<Seleccione tipo de Cuenta>"
    tdbcTipoCta.AddItem "1;ACTIVOS"
    tdbcTipoCta.AddItem "2;PASIVOS"
    tdbcTipoCta.AddItem "3;PATRIMONIO"
    tdbcTipoCta.AddItem "4;GASTOS"
    tdbcTipoCta.AddItem "5;INGRESOS"
    tdbcTipoCta.AddItem "6;GESTION"
    tdbcTipoCta.AddItem "7;FUNCION"
    tdbcTipoCta.AddItem "8;ORDEN"
    tdbcTipoCta.Bookmark = 0
    tdbcTipoCta.ListField = "column1"
    tdbcTipoCta.BoundColumn = "column0"
    tdbcTipoCta.ReBind
    
End Sub

Private Sub BuscarMonedaNacional()
    Dim i As Integer
    For i = 0 To tdbcMoneda.ListCount - 1
        tdbcMoneda.Row = i
        If tdbcMoneda.Columns(2).Value = "1" Then
            tdbcMoneda.Bookmark = i
            Exit Sub
        End If
    Next
    tdbcMoneda.Bookmark = tdbcMoneda.Bookmark
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
    Set frmRepAnexoInvBalance = Nothing
    Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))
End Sub

Private Sub tdbcMes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pSetFocus cmdImprimir
End If
End Sub
Private Sub tdbcMesFin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pSetFocus tdbcMoneda
End If
End Sub
Private Sub tdbcMoneda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then pSendKeys "{tab}"
End Sub

Sub Imprimir()
On Error GoTo Control

 gsAccionRep = 1
 gsCodMoneda = tdbcMoneda.BoundText
 gsDiarioPeriodo = tdbcMes.BoundText
 gsTipoImp = IIf(chkImpLibCajaResumido.Value, "1", "0")
 
 frmFCImpresion.Show
 
Exit Sub
Control:
 MsgBox Err.Description
End Sub
Private Function ExistenDatos() As Boolean
On Error GoTo Error_cmd

Dim sSql As String
Dim clsMante As New clsMantoTablas
Dim arrDatos() As Variant

Screen.MousePointer = vbHourglass

sSql = "spCn_ReprocesoDiarioSimp_2D '" & gsEmpresa & "','" & gsAnio & "','" & gsDiarioPeriodo & "',1," & gsNumDigDiarioSimpRep & _
         ",'" & gsCodMoneda & "','" & gsTipoImp & "'"
         
arrDatos = Array(sSql)
Set rsArreglo = clsMante.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())

Screen.MousePointer = vbNormal
ExistenDatos = IIf(rsArreglo.State > 0, True, False)
If rsArreglo.State > 0 Then If Not rsArreglo.EOF Then rsArreglo.MoveFirst

Exit Function
    
Error_cmd:
    Screen.MousePointer = vbNormal
    ExistenDatos = False
    MsgBox Err.Description, vbInformation, App.Title
End Function
Private Function ExistenDatosTotales() As Boolean
On Error GoTo Error_cmd

Dim sSql As String
Dim clDatos As clsMantoTablas
Dim arrDatos() As Variant

Screen.MousePointer = vbHourglass

Set clDatos = New clsMantoTablas
sSql = "spCn_ReprocesoDiarioSimp_2D_Totales "
arrDatos = Array(sSql)
Set rsArregloTotales = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())

Screen.MousePointer = vbNormal
ExistenDatosTotales = IIf(rsArregloTotales.State > 0, True, False)
If rsArregloTotales.State > 0 Then If Not rsArregloTotales.EOF Then rsArregloTotales.MoveFirst

Exit Function
    
Error_cmd:
    Screen.MousePointer = vbNormal
    ExistenDatosTotales = False
    MsgBox Err.Description, vbInformation, App.Title
End Function

'HT : 20091111
Public Sub ReporteDiarioSimplificado()

Dim CodAlmacen As String
Dim CodProducto As String
Dim sSaldoanterior As String * 10
Dim sSaldoFinAlm As String * 10
Dim vSaldoAnterior As Double
Dim iContCopias As Integer
Dim nFilRes As Integer
Dim vienen As String

On Error Resume Next
Screen.MousePointer = vbHourglass

If Not ExistenDatos() And Not ExistenDatosTotales() Then
 MsgBox "No existen datos para Imprimir el Reporte.", vbExclamation, App.Title
 Exit Sub
End If

If frmFCImpresion.List_Destino.Text = "Archivo" Then
   Open frmFCImpresion.OutputFileName For Output Shared As #1
   gsPagina = 0
ElseIf frmFCImpresion.List_Destino.Text = "Impresora" Then
    Open frmFCImpresion.OutputFileName For Output Shared As #1
End If

giLineas = 0
giEspacios = 60

If iReport = 1 Then rsArreglo.MoveFirst

If Not rsArreglo.EOF Then
   iReport = 1

gsConTotalPaginas = 0
Dim VarCont As Byte
Dim VarCuentas As String
Dim VarNumColIni As Integer
Dim VarLin As Byte
'gsPaginaPrincipal = Val(frmFCImpresion.txtPrincipal.Text)
gsPagina = gsPagina + 1
VarNumColIni = 3

With rsArreglo
            
    Dim NumFila As Long
    Dim NumCol As Long
'    sPag = Space(4)
    
    'Print #1,  "LIBRO DIARIO FORMATO SIMPLIFICADO", "Fecha :  " & Format(fechaServidor, "dd/MM/yyyy"))
    If Gs_TamPapel = 39 Then NumCol = 13 Else NumCol = 6
    Dim VarTotal(0 To 1000) As Double, x As Byte, mov As Byte, VarPosIni As Long, VarLinea As String, VarContador As Byte, Col As Integer, z As Integer
    
    Dim VarTotal_TM(0 To 1000) As Double
    Dim VarPagina As Integer, k As Byte
    Dim VarPagPrin As Integer
    
    VarPagPrin = xGs_Principal
    xGs_Principal = 0
    
    VarPagina = 1
    VarCont = 2
    
'    Print #1, Chr(27) & Chr(64); 'Inicializa
'    Print #1, Chr(27) & Chr(120) & Chr(0); 'Draft
'    Print #1, Chr(27) & Chr(15); 'Comprimido
'    ' Print #1, Chr(27) & Chr(137); '15cpi
'    Print #1, Chr(27) & Chr(51) & Chr(29) 'Entre lineas 29/180
    
    Col = 1
    
    VarPosIni = .AbsolutePosition
    
    Do While .EOF = False
    
'    VarPagPrin = xGs_Principal
        
        VarCuentas = ""
        x = 0
        
        Do While x < NumCol
            If VarCont + x >= rsArreglo.Fields.Count - 1 Then Exit Do
            x = x + 1
            VarCuentas = VarCuentas & String(6, " ") & .Fields(Val(VarCont + x)).Name & "" & String(7 - Len(.Fields(Val(VarCont + x)).Name), " ")
        Loop
        
        'Debug.Print VarCuentas
        
        If Val(frmFCImpresion.txtDesde.Text) <= 0 And xGs_Principal = 0 Then GoTo entrar_1
        
        If VarPagina > Val(frmFCImpresion.txtHasta.Text) And Val(frmFCImpresion.txtHasta.Text) > 0 And (k = 0) Then
            GoTo SALTAR
        End If
            'If Val(frmFCImpresion.txtHasta.Text) = 0 And Val(frmFCImpresion.txtDesde.Text) = 0 And (k = 0) Then GoTo SALTAR
        '*--*----*--*-**-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
            If xGs_Principal = 0 Then
                xGs_Principal = 1
            End If

            'If VarPagPrin = 1 Then GoTo entrar_3
'            If xGs_Principal > VarPagPrin And VarPagPrin = 0 Then
'                GoTo SALTAR
'            End If
            
        '*--*----*--*-**-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
        
        If VarPagina < Val(frmFCImpresion.txtDesde.Text) Then
            k = 1
            GoTo Control
        Else
entrar_1:
            If xGs_Principal = 0 Then
                xGs_Principal = 1
            End If
             
            'If VarPagPrin = 1 Then GoTo entrar_3
            If xGs_Principal = VarPagPrin Or VarPagPrin = 0 Then
entrar_3:
                k = 0
            Else
                k = 1
                GoTo Control
            End If
        End If

Aqui:
        If Gs_TamPapel = 39 Then ' 242 ' 33 ' 19
            
            'Chr (15) + Chr(27) + "P"

            Print #1, Space(3) & "FORMATO 5.2: LIBRO DIARIO FORMATO SIMPLIFICADO" & String(162, " ") & "Fecha :  " & Format(Now, "dd/MM/yyyy")
            Print #1, Space(3) & "PERIODO/EJERCICIO    : " & NombreMes(gsDiarioPeriodo) & " " & gsAnio & String(175, " ")
                        '23                                 8           13
            Print #1, Space(3) & "RUC                  : " & gsRUC & String(185 - Len(gsRUC), " ") & "Pagina:      " & Format$(xGs_Principal, "###") & " de " & Format$(VarPagina, "###")
            Print #1, Space(3) & "APELLIDOS Y NOMBRES,"
            Print #1, Space(3) & "DENOMINACIÓN O"
            Print #1, Space(3) & "RAZON SOCIAL         : " & gsEmpresaNom & String(185 - Len(gsEmpresaNom), " ")
            
            Print #1, Space(3) & "MONEDA               : " & IIf(gsCodMoneda = gsMonedaNac, gsNombreMonedaNac, gsNombreMonedaExt) & _
            String(185 - Len(IIf(gsCodMoneda = gsMonedaNac, gsNombreMonedaNac, gsNombreMonedaExt)), " ") ' & "Usuario: " & gsUsuario
            
            Print #1, ""
            Print #1, ""
            'Print #1, ""
            
            Print #1, Space(3) & String(12, "-") & " " & String(14, "-") & " " & String(30, "-") & " " & String(Len(VarCuentas), "-")
            ' 14 Lineas por Cta.
            Print #1, Space(3) & "Nro CORREL. " & " " & "FECHA/PERIODO " & " " & String(7, " ") & "GLOSA O DESCRIPCION"
            Print #1, Space(3) & "  O C.U.O   " & " " & "  DE OPER.    " & " " & String(7, " ") & "  DE LA OPERACION  " & "    " & VarCuentas
            Print #1, Space(3) & String(12, "-") & " " & String(14, "-") & " " & String(30, "-") & " " & String(Len(VarCuentas), "-")
            
        Else
            
'''''''            Print #1, ""
            Print #1, Space(3) & "FORMATO 5.2: LIBRO DIARIO FORMATO SIMPLIFICADO" & String(72, " ") & "Fecha :  " & Format(Now, "dd/MM/yyyy")
            Print #1, Space(3) & "PERIODO/EJERCICIO    : " & NombreMes(gsDiarioPeriodo) & " " & gsAnio & String(85, " ")
                        '23                                 8           13
            Print #1, Space(3) & "RUC                  : " & gsRUC & String(95 - Len(gsRUC), " ") & "Pagina:      " & Format$(xGs_Principal, "###") & " de " & Format$(VarPagina, "###")
            Print #1, Space(3) & "APELLIDOS Y NOMBRES,"
            Print #1, Space(3) & "DENOMINACIÓN O"
            Print #1, Space(3) & "RAZON SOCIAL         : " & gsEmpresaNom & String(95 - Len(gsEmpresaNom), " ")
            
            Print #1, Space(3) & "MONEDA               : " & IIf(gsCodMoneda = gsMonedaNac, gsNombreMonedaNac, gsNombreMonedaExt) & _
            String(95 - Len(IIf(gsCodMoneda = gsMonedaNac, gsNombreMonedaNac, gsNombreMonedaExt)), " ") '& "Usuario: " & gsUsuario
            
            Print #1, ""
            Print #1, ""
            'Print #1, ""
            
            Print #1, Space(3) & String(12, "-") & " " & String(14, "-") & " " & String(30, "-") & " " & String(Len(VarCuentas), "-")
            ' 14 Lineas por Cta.
            Print #1, Space(3) & "Nro CORREL. " & " " & "FECHA/PERIODO " & " " & String(7, " ") & "GLOSA O DESCRIPCION"
            Print #1, Space(3) & "  O C.U.O   " & " " & "  DE OPER.    " & " " & String(7, " ") & "  DE LA OPERACION  " & "    " & VarCuentas
            Print #1, Space(3) & String(12, "-") & " " & String(14, "-") & " " & String(30, "-") & " " & String(Len(VarCuentas), "-")
            
            
        End If
        
Control:
        
        NumFila = 14
        Dim sw As Boolean
        sw = False
        'VarLinea = ""
        
        Do While NumFila <= 72
            NumFila = NumFila + 1
            VarLinea = ""
         If vienen <> "" Then Print #1, Space(3) & vienen: sw = True: vienen = ""
            VarLinea = String(12 - Len(.Fields("Ase_NVoucher").Value), " ") & .Fields("Ase_NVoucher").Value & " " & _
            "   " & String(11 - Len(.Fields("Asd_dFecDoc").Value), " ") & .Fields("Asd_dFecDoc").Value & " " & _
            Left(.Fields("Asl_cDescripcion").Value, 27) & String(27 - Len(Left(.Fields("Asl_cDescripcion").Value, 27)), " ")
            
            For VarContador = 1 To x
                VarLinea = VarLinea & _
                String(13 - Len(LTrim(RTrim(Format(Val(IIf(IsNull(.Fields(Val(VarCont + VarContador)).Value), 0, _
                .Fields(Val(VarCont + VarContador)).Value)), "0.00")))), " ") & Format(Val(IIf(IsNull(.Fields(Val(VarCont + VarContador)).Value), 0, _
                .Fields(Val(VarCont + VarContador)).Value)), "0.00")
            Next VarContador
            
            'Debug.Print VarLinea
            
            If k = 0 Then
                Print #1, Space(3) & VarLinea
            End If
            
            If mov = 1 Then
                z = 0
                Do While z < x
                    z = z + 1
                    VarTotal(VarCont + z) = Val(VarTotal(VarCont + z)) + Val(IIf(IsNull(.Fields(Val(VarCont + z)).Value), 0, .Fields(Val(VarCont + z)).Value))
                    VarTotal_TM(VarCont + z) = Val(VarTotal_TM(VarCont + z)) + Val(IIf(IsNull(.Fields(Val(VarCont + z)).Value), 0, .Fields(Val(VarCont + z)).Value))
                Loop

            Else
                z = 0
                Do While z < x
                    z = z + 1
                    VarTotal_TM(VarCont + z) = Val(VarTotal_TM(VarCont + z)) + Val(IIf(IsNull(.Fields(VarCont + z).Value), 0, .Fields(VarCont + z).Value))
                Loop
            End If
            
            mov = 1
            
            If NumFila <= 72 Then
                .MoveNext
            End If
            Dim TotLineas As Integer
            If .EOF Then
                If sw = True Then TotLineas = 71 Else TotLineas = 72
                   For nFilRes = 1 To (TotLineas - NumFila)
                   If k = 0 Then Print #1, ""
                   Next nFilRes
                    Exit Do
                    sw = False
                End If
                'MsgBox gsPagina
        Loop
        
        If .EOF = False Then
            If k = 0 Then
                Print #1, Space(3) & String(13, "-") & String(15, "-") & String(31, "-") & String(Len(VarCuentas), "-")
            End If
            
            VarLinea = "                                             VAN...    "
            vienen = "                                            VIENEN...  "
        
            z = 0
            Do While z < x
                z = z + 1
                VarLinea = VarLinea & String(13 - Len(LTrim(RTrim(Format(Val(VarTotal_TM(VarCont + z)), "0.00")))), " ") & Format(Val(VarTotal_TM(VarCont + z)), "0.00")
                vienen = vienen & String(13 - Len(LTrim(RTrim(Format(Val(VarTotal_TM(VarCont + z)), "0.00")))), " ") & Format(Val(VarTotal_TM(VarCont + z)), "0.00")
            Loop
            
            'Debug.Print VarLinea
        
            If k = 0 Then
                Print #1, Space(3) & VarLinea
                Print #1, Space(3) & String(13, "-") & String(15, "-") & String(31, "-") & String(Len(VarCuentas), "-")
                'Print #1, ""
                'Print #1, Chr(27) & Chr(12)
'                If Not gsNomTipoImp Then
                'Print #1, "": Print #1, "": Print #1, ""

            End If
            
            If VarCont < rsArreglo.Fields.Count - 1 Then
                'If VarPosIni = 1 Then VarPosIni = VarPosIni + 1
'                .AbsolutePosition = VarPosIni
            Else
                .MoveNext
                VarPosIni = .AbsolutePosition
                
                VarPagina = 0
                VarCont = 2
            End If
            .MoveNext
            If xGs_Principal = VarPagPrin And VarPagina = Val(frmFCImpresion.txtHasta.Text) Then
                GoTo SALTAR
            End If

            If k = 0 Then VarPagina = VarPagina + 1: GoTo Aqui:
        Else
            If k = 0 Then
                Print #1, Space(3) & String(13, "-") & String(15, "-") & String(31, "-") & String(Len(VarCuentas), "-")
'               Print #1, Chr(27) & Chr(12)
            End If
            
            VarLinea = "   TOTAL MOVIMIENTOS DEL MES" & "                           "
    
            z = 0
            Do While z < x
                z = z + 1
                VarLinea = VarLinea & String(13 - Len(LTrim(RTrim(Format(Val(VarTotal(VarCont + z)), "0.00")))), " ") & Format(Val(VarTotal(VarCont + z)), "0.00")
            Loop
            
            If k = 0 Then
                Print #1, Space(3) & VarLinea
'                Print #1, Chr(27) & Chr(12)
            End If
            
            VarLinea = "   TOTAL ACUMULADO" & "                                     "
    
            z = 0
            Do While z < x
                z = z + 1
               VarLinea = VarLinea & String(13 - Len(LTrim(RTrim(Format(Val(VarTotal_TM(VarCont + z)), "0.00")))), " ") & Format(Val(VarTotal_TM(VarCont + z)), "0.00")
            Loop
            

            VarCont = (VarCont + x)
            
            If k = 0 Then
                Print #1, Space(3) & VarLinea
                Print #1, Space(3) & String(13, "-") & String(15, "-") & String(31, "-") & String(Len(VarCuentas), "-")
               ' If VarPagPrin = 0 Then VarPagina = 0 ': VarCont = 2
            End If
            
            If VarPagPrin = 0 Or VarPagPrin > 1 Then xGs_Principal = xGs_Principal + 1: VarPagina = 0
            
        ' xGs_Principal = xGs_Principal + 1:
            If VarCont >= rsArreglo.Fields.Count - 1 Then
                Exit Do
            Else
                .AbsolutePosition = VarPosIni
            End If
        End If
        vienen = ""
        VarPagina = VarPagina + 1

    Loop
SALTAR:
 Close #1
 
    Open frmFCImpresion.OutputFileName For Input As #1

    If Len(Input(LOF(1), 1)) = 0 Then
        MsgBox "No existen datos para Imprimir, vefique.", vbExclamation, App.Title
    Else
        
        frmFCVistaInforme.txtInforme.LoadFile frmFCImpresion.OutputFileName, rtfText
    End If
Close #1
    
End With

End If

End Sub
'HT : 20091111
Public Sub CabeceraDiarioSimplificado()

Dim sPag As String
Dim Anio As String
Dim Mes As String
Dim sUSUARIO As String * 10

On Error GoTo ERROR
Gs_HoraServ = DevuelveHoraServidor
sUSUARIO = gsUsuario

If Gs_TamPapel = 39 Then nAncho = 232 Else nAncho = 142
gsConTotalPaginas = gsConTotalPaginas + 1
gsPagina = gsPagina + 1

'gsPagina = gsPagina + 1
'If Not gsControlPag Then
'  gsPaginaPrincipal = 1
' Else
'  gsPaginaPrincipal = gsPaginaPrincipal + 1: gsPagina = 0: gsPagina = gsPagina + 1
' End If
If frmFCImpresion.txtHasta.Text = 0 Then GoTo Salto
If (Val(frmFCImpresion.txtHasta.Text) - Val(frmFCImpresion.txtDesde.Text)) + 1 < gsConTotalPaginas Then
    VarGsIndDS = False
    Exit Sub
Else
Salto:
    VarGsIndDS = True
End If

printl ("")
sPag = Space(4)

RSet sPag = Format(gsPagina + 1, "####")
giLineas = 0

Call AlinearDosTextos(nAncho, "LIBRO DIARIO FORMATO SIMPLIFICADO", "Fecha :  " & Format(FechaServidor, "dd/MM/yyyy"))
Call AlinearDosTextos(nAncho, "PERIODO/EJERCICIO    : " & NombreMes(gsDiarioPeriodo) & " " & gsAnio, "")
Call AlinearDosTextos(nAncho, "RUC                  : " & gsRUC, "Hora  : " & Gs_HoraServ)
Call AlinearDosTextos(nAncho, "RAZON SOCIAL         : " & gsEmpresaNom, "Pagina:       " & Format$(xGs_Principal, "####") & " de " & Format$(gsPagina, "####"))
Call AlinearDosTextos(nAncho, "MONEDA               : " & IIf(gsCodMoneda = gsMonedaNac, gsNombreMonedaNac, gsNombreMonedaExt), "") ', "Usuario: " & sUSUARIO)

Exit Sub
With rsArreglo
 If Gs_TamPapel = 39 Then
    For i = 0 To 12
     If ColAlmc = rsArreglo.Fields.Count - 1 And xCuentas = "" Then ColAlmc = ColInicio
     If ColAlmc < rsArreglo.Fields.Count - 1 Then
      If UltCol > ColInicio Then
        xCuentas = xCuentas + .Fields(UltCol + i).Name + "          "
        ColAlmc = UltCol + i
        Cont = Cont + 1
      Else
        xCuentas = xCuentas + .Fields(ColInicio + i).Name + "          "
        ColAlmc = ColInicio + i
        Cont = Cont + 1
      End If
     Else
       Exit For
     End If
    Next i
    UltCol = ColAlmc + 1
    'if Cont = 13 Then Cont = 0
Else
    For i = 0 To 5
     If ColAlmc = rsArreglo.Fields.Count - 1 And xCuentas = "" Then ColAlmc = ColInicio
     If ColAlmc < rsArreglo.Fields.Count - 1 Then
      If UltCol > ColInicio Then
       xCuentas = xCuentas + .Fields(UltCol + i).Name + "          "
       ColAlmc = UltCol + i
       Cont = Cont + 1
      Else
       xCuentas = xCuentas + .Fields(ColInicio + i).Name + "          "
       ColAlmc = ColInicio + i
      End If
     Else
       Exit For
     End If
    Next i
    UltCol = ColAlmc + 1
    If Cont = 6 Then Cont = 0
    'Cont = 0
End If
End With

    printl ("")
    printl ("")
    printl ("")
    'printl ("------------ -------------- ----------------------------------- " & String(Len(xCuentas), "-"))
    'printl ("Nro CORREL.  FECHA/PERIODO         GLOSA O DESCRIPCION   ")
    'printl ("  O C.U.O      DE OPER.              DE LA OPERACIÓN                    " & xCuentas)
    'printl ("------------ -------------- ----------------------------------- " & String(Len(xCuentas), "-"))

    printl ("------------ -------------- ----------------------------------- " & String(Len(xCuentas) - 4, "-"))
    printl ("Nro CORREL.  FECHA/PERIODO         GLOSA O DESCRIPCION   ")
    printl ("  O C.U.O      DE OPER.              DE LA OPERACION                  " & xCuentas)
    printl ("------------ -------------- ----------------------------------- " & String(Len(xCuentas) - 4, "-"))
    
xLineasCuentas = xCuentas
xCuentas = ""
Exit Sub

ERROR:
 MsgBox Err.Description, vbCritical, App.Title
 Resume
 Exit Sub
End Sub

Public Sub ImprimeDetalle()
On Error GoTo Control

Dim sNroVoucher As String * 10
Dim sFechaDoc As String * 10
'Dim sGlosa As String * 35
Dim sglosa As String * 33
Dim sImporte As String * 13
Dim sImporteTotal As String
Dim i As Integer

With rsArreglo
   'If !Ase_nVoucher = "0403000006" Then MsgBox "STOP.....PLAYER"
   sNroVoucher = "" & !Ase_nVoucher
   sFechaDoc = "" & !Asd_dFecDoc
   sglosa = "" & !Asl_cDescripcion
'If ColAlmc = 0 Then Exit Sub
If ColAlmc <> rsArreglo.Fields.Count - 1 Then
 If Gs_TamPapel = 39 Then 'US Standar
   For i = ColAlmc - 12 To UltCol - 1
    RSet sImporte = Format$(IIf(IsNull(.Fields(i)), 0, .Fields(i)), "#,###,###,##0.00;(#,###,###,##0.00)")
    sImporteTotal = sImporteTotal & sImporte
   Next i
 Else 'Carta
    
    If ColAlmc > 0 Then
        For i = ColAlmc - 5 To UltCol - 1
            RSet sImporte = Format$(IIf(IsNull(.Fields(i)), 0, .Fields(i)), "#,###,###,##0.00;(#,###,###,##0.00)")
            sImporteTotal = sImporteTotal & sImporte
        Next i
    End If
 End If
Else 'Igual al Nro total de columnas
 If Gs_TamPapel = 39 Then 'US Standar
  ContLineas = 0
  Call NroColumnas
    'For I = (ColAlmc - Cont + 1) To UltCol - 1
    For i = (ColAlmc - ContLineas + 1) To UltCol - 1
   'For i = (ColAlmc - 12 + Cont + ColInicio) To UltCol - 1
    RSet sImporte = Format$(IIf(IsNull(.Fields(i)), 0, .Fields(i)), "#,###,###,##0.00;(#,###,###,##0.00)")
    sImporteTotal = sImporteTotal & sImporte
   Next i
 Else 'Carta
  ContLineas = 0
  Call NroColumnas
   For i = (ColAlmc - ContLineas + 1) To UltCol - 1
    RSet sImporte = Format$(IIf(IsNull(.Fields(i)), 0, .Fields(i)), "#,###,###,##0.00;(#,###,###,##0.00)")
    sImporteTotal = sImporteTotal & sImporte
   Next i
 End If
End If
End With
'If sNroVoucher = "0204000343" Then MsgBox "vdv"
Debug.Print gsLinea
gsLinea = sNroVoucher & "    " & sFechaDoc & "    " & sglosa & " " & sImporteTotal
printl gsLinea

Exit Sub

Control:
 MsgBox Err.Description
 Resume
End Sub
Sub ImprimeTotales()
On Error GoTo Control
    For j = 1 To Len(xLineasCuentas)
     If Mid(xLineasCuentas, j, 1) <> " " Then
      xCuentaTotal = xCuentaTotal & Mid(xLineasCuentas, j, 1)
      If xCuentaTotal = "401" Or xCuentaTotal = "403" Then
       j = j + 1
       xCuentaTotal = xCuentaTotal & Mid(xLineasCuentas, j, 1)
      End If
      If Len(xCuentaTotal) = 3 Then
        Indice = Indice + 1
        If Gs_TamPapel = "1" Then ArrayCuentas(Indice) = Trim(xCuentaTotal) Else ArrayCuentasUS(Indice) = Trim(xCuentaTotal)
        xCuentaTotal = ""
      ElseIf Len(xCuentaTotal) = 4 Then
        Indice = Indice + 1
        If Gs_TamPapel = "1" Then ArrayCuentas(Indice) = Trim(xCuentaTotal) Else ArrayCuentasUS(Indice) = Trim(xCuentaTotal)
        xCuentaTotal = ""
      End If
     End If
    Next j

    Do While Not rsArregloTotales.EOF
     For j = 1 To Indice
     
        If Gs_TamPapel = "1" Then
            rsArregloTotales.Find rsArregloTotales.Fields(1).Name & "='" & Trim(ArrayCuentas(j)) & "'"
        Else
            rsArregloTotales.Find rsArregloTotales.Fields(1).Name & "='" & Trim(ArrayCuentasUS(j)) & "'"
        End If
        
        If rsArregloTotales.EOF = False Then
            RSet ImpMes = Format(rsArregloTotales.Fields(2), "#,###,###,##0.00;(#,###,###,##0.00)")
            LineaMes = LineaMes & ImpMes
            RSet ImpAcum = Format(rsArregloTotales.Fields(2) + rsArregloTotales.Fields(3), "#,###,###,##0.00;(#,###,###,##0.00)")
            LineaTotales = LineaTotales & ImpAcum
        End If
        
     Next j
     If j - 1 = Indice Then rsArregloTotales.MoveFirst: Exit Do
    Loop

    If Gs_TamPapel = "1" Then Erase ArrayCuentas() Else Erase ArrayCuentasUS()
    xCuentaTotal = ""
    Indice = 0

    'printl (String(Len("  O C.U.O      DE OPER.              DE LA OPERACIÓN                    " & xLineasCuentas) - 8, "-"))
    'printl ("   TOTAL MOVIMIENTOS DEL MES" & "                                    " & LineaMes)
    'printl ("   TOTAL ACUMULADO" & "                                              " & LineaTotales)
    'printl (String(Len("  O C.U.O      DE OPER.              DE LA OPERACIÓN                    " & xLineasCuentas) - 8, "-"))

    printl (String(Len("  O C.U.O      DE OPER.              DE LA OPERACION                    " & xLineasCuentas) - 10, "-"))
    printl ("   TOTAL MOVIMIENTOS DEL MES" & "                                  " & LineaMes)
    printl ("   TOTAL ACUMULADO" & "                                            " & LineaTotales)
    printl (String(Len("  O C.U.O      DE OPER.              DE LA OPERACION                    " & xLineasCuentas) - 10, "-"))
    
    LineaMes = "": ImpMes = ""
    LineaTotales = "": ImpAcum = ""
Exit Sub
Control:
 MsgBox Err.Description
 'Resume
End Sub
Sub NroColumnas()
Dim i As Integer
Dim Aux_xLineasCuentas As String

 For i = 1 To Len(xLineasCuentas)
  If Trim(Mid(xLineasCuentas, i, 3)) <> "" Then
   ContLineas = ContLineas + 1
   i = i + 12
  End If
 Next i
End Sub
Private Function ValidaParametrosDiario()
On Error GoTo Control

 Dim clDatos As clsMantoTablas
 Dim arrDatos() As Variant
 Dim sqlSp As String
 Dim rs As ADODB.Recordset
ValidaParametrosDiario = False

 Set clDatos = New clsMantoTablas
 sqlSp = "spCNT_SelParamDiario_CONFIG_LIBROS '" & gsEmpresa & "', '" & gsAnio & "'"
 arrDatos = Array(sqlSp)
 Set rs = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
 If rs Is Nothing Then
  Exit Function
 Else
  If rs!Cfl_cDiarioSimplificado = "0" Or rs!Cfl_nDigDiarioRep = 0 Then
   ValidaParametrosDiario = True
   MsgBox "Debe Configurar los Parámetros Iniciales para Imprimir el Reporte." & Chr(10) + Chr(13) & _
          " 1.- Dirigase a la Opción Configuración del Menú Principal." & Chr(10) + Chr(13) & _
          " 2.- Seleccione la Opción Parametros Iniciales." & Chr(10) + Chr(13) & _
          " 3.- Active la Casilla 'Habilitar' de la Sección 'Diario Formato Simplificado' " & Chr(10) + Chr(13) & _
          " 4.- Ingrese el Número de Dígitos del Reporte.", vbInformation, App.Title
   Exit Function
  Else
   ValidaParametrosDiario = False
  End If
 End If
 
Exit Function
Control:
 ValidaParametrosDiario = False
 MsgBox Err.Description
End Function
