VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmRepLibroMayorAnalitico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mayor General Analitico"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7905
   Icon            =   "frmRepLibroMayorAnalitico.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6555
   ScaleWidth      =   7905
   Begin VB.Frame fraTodo 
      Height          =   6405
      Left            =   135
      TabIndex        =   7
      Top             =   45
      Width           =   7620
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   780
         TabIndex        =   34
         Top             =   4680
         Width           =   5865
         Begin VB.OptionButton OptTIpo 
            Caption         =   "Reporte de Operaciones Individuales"
            Enabled         =   0   'False
            Height          =   555
            Index           =   1
            Left            =   3480
            TabIndex        =   37
            Top             =   240
            Width           =   2295
         End
         Begin VB.OptionButton OptTIpo 
            Caption         =   "Detalle de Centralización de Operaciones"
            Enabled         =   0   'False
            Height          =   435
            Index           =   0
            Left            =   240
            TabIndex        =   36
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1935
         Left            =   405
         TabIndex        =   21
         Top             =   120
         Width           =   6795
         Begin VB.OptionButton Option1 
            Caption         =   "PERIODO"
            Height          =   255
            Left            =   1080
            TabIndex        =   33
            Top             =   240
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton Option2 
            Caption         =   "RANGO DE FECHAS"
            Height          =   255
            Left            =   4320
            TabIndex        =   32
            Top             =   240
            Width           =   1935
         End
         Begin TrueOleDBList70.TDBCombo tdbcDelMes 
            Height          =   300
            Left            =   1200
            TabIndex        =   22
            Top             =   600
            Width           =   2400
            _ExtentX        =   4233
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
            _PropDict       =   $"frmRepLibroMayorAnalitico.frx":0ECA
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
            Left            =   3000
            TabIndex        =   23
            Tag             =   "_"
            Top             =   1470
            Width           =   2400
            _ExtentX        =   4233
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
            Locked          =   0   'False
            ScrollTrack     =   0   'False
            RowDividerColor =   12632256
            RowSubDividerColor=   12632256
            AddItemSeparator=   ";"
            _PropDict       =   $"frmRepLibroMayorAnalitico.frx":0F51
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
         Begin TrueOleDBList70.TDBCombo tdbcAlMes 
            Height          =   300
            Left            =   1200
            TabIndex        =   24
            Top             =   1005
            Width           =   2400
            _ExtentX        =   4233
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
            _PropDict       =   $"frmRepLibroMayorAnalitico.frx":0FD8
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
         Begin TDBDate6Ctl.TDBDate dtpDesde 
            Height          =   300
            Left            =   4800
            TabIndex        =   28
            Tag             =   "enabled"
            Top             =   600
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   529
            Calendar        =   "frmRepLibroMayorAnalitico.frx":105F
            Caption         =   "frmRepLibroMayorAnalitico.frx":1161
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmRepLibroMayorAnalitico.frx":11C5
            Keys            =   "frmRepLibroMayorAnalitico.frx":11E3
            Spin            =   "frmRepLibroMayorAnalitico.frx":124F
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
            Text            =   "03/08/2004"
            ValidateMode    =   0
            ValueVT         =   2010185735
            Value           =   38202
            CenturyMode     =   0
         End
         Begin TDBDate6Ctl.TDBDate dtpHasta 
            Height          =   300
            Left            =   4800
            TabIndex        =   29
            Tag             =   "enabled"
            Top             =   1005
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   529
            Calendar        =   "frmRepLibroMayorAnalitico.frx":1277
            Caption         =   "frmRepLibroMayorAnalitico.frx":1379
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmRepLibroMayorAnalitico.frx":13DD
            Keys            =   "frmRepLibroMayorAnalitico.frx":13FB
            Spin            =   "frmRepLibroMayorAnalitico.frx":1467
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
            Text            =   "03/08/2004"
            ValidateMode    =   0
            ValueVT         =   2010185735
            Value           =   38202
            CenturyMode     =   0
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "HASTA"
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
            Index           =   0
            Left            =   3900
            TabIndex        =   31
            Top             =   1035
            Width           =   570
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "DESDE"
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
            Left            =   3900
            TabIndex        =   30
            Top             =   645
            Width           =   630
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "DEL MES"
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
            Left            =   240
            TabIndex        =   27
            Top             =   645
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "AL MES"
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
            Left            =   240
            TabIndex        =   26
            Top             =   1035
            Width           =   630
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
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
            Height          =   225
            Index           =   1
            Left            =   2040
            TabIndex        =   25
            Top             =   1480
            Width           =   765
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   420
         TabIndex        =   14
         Top             =   3900
         Width           =   6795
         Begin MSForms.OptionButton OptImpresion 
            Height          =   390
            Index           =   4
            Left            =   5160
            TabIndex        =   35
            Top             =   240
            Width           =   1455
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   5
            Size            =   "2566;688"
            Value           =   "0"
            Caption         =   "Reporte PLE"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.OptionButton OptImpresion 
            Height          =   510
            Index           =   1
            Left            =   2625
            TabIndex        =   16
            Top             =   165
            Width           =   2175
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   5
            Size            =   "3836;900"
            Value           =   "1"
            Caption         =   "Impresión Formato Láser"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.OptionButton OptImpresion 
            Height          =   420
            Index           =   0
            Left            =   60
            TabIndex        =   15
            Top             =   210
            Width           =   2310
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   5
            Size            =   "4075;741"
            Value           =   "0"
            Caption         =   "Impresión Formato Matricial"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1005
         Left            =   780
         ScaleHeight     =   1005
         ScaleWidth      =   6510
         TabIndex        =   17
         Top             =   4395
         Visible         =   0   'False
         Width           =   6510
         Begin VB.Frame Frame2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   0
            TabIndex        =   18
            Top             =   240
            Width           =   5865
            Begin VB.OptionButton OptForma 
               Caption         =   "Centralizado"
               Height          =   255
               Index           =   1
               Left            =   3345
               TabIndex        =   20
               Top             =   285
               Width           =   1650
            End
            Begin VB.OptionButton OptForma 
               Caption         =   "Detallado"
               Height          =   255
               Index           =   0
               Left            =   465
               TabIndex        =   19
               Top             =   285
               Width           =   1650
            End
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1455
         Left            =   405
         TabIndex        =   8
         Top             =   2400
         Width           =   6795
         Begin VB.OptionButton optTodos 
            Caption         =   "Todas las Cuentas"
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
            Left            =   360
            TabIndex        =   1
            Top             =   240
            Value           =   -1  'True
            Width           =   2535
         End
         Begin VB.OptionButton optCuentas 
            Caption         =   "Por Rango de Cuentas"
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
            Left            =   3000
            TabIndex        =   2
            Top             =   240
            Width           =   2775
         End
         Begin TDBText6Ctl.TDBText tdbtCuentaDesde 
            Height          =   315
            Left            =   1200
            TabIndex        =   3
            Tag             =   "_"
            Top             =   600
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   556
            Caption         =   "frmRepLibroMayorAnalitico.frx":148F
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmRepLibroMayorAnalitico.frx":14FB
            Key             =   "frmRepLibroMayorAnalitico.frx":1519
            BackColor       =   -2147483643
            EditMode        =   0
            ForeColor       =   -2147483640
            ReadOnly        =   0
            ShowContextMenu =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MarginBottom    =   1
            Enabled         =   0
            MousePointer    =   0
            Appearance      =   1
            BorderStyle     =   1
            AlignHorizontal =   0
            AlignVertical   =   0
            MultiLine       =   0
            ScrollBars      =   0
            PasswordChar    =   ""
            AllowSpace      =   -1
            Format          =   "aA"
            FormatMode      =   1
            AutoConvert     =   -1
            ErrorBeep       =   0
            MaxLength       =   12
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
         Begin TDBText6Ctl.TDBText tdbtDescripcionDesde 
            Height          =   315
            Left            =   2550
            TabIndex        =   9
            Tag             =   "_"
            Top             =   600
            Width           =   3705
            _Version        =   65536
            _ExtentX        =   6535
            _ExtentY        =   556
            Caption         =   "frmRepLibroMayorAnalitico.frx":156B
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmRepLibroMayorAnalitico.frx":15D7
            Key             =   "frmRepLibroMayorAnalitico.frx":15F5
            BackColor       =   -2147483643
            EditMode        =   0
            ForeColor       =   -2147483640
            ReadOnly        =   1
            ShowContextMenu =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MarginBottom    =   1
            Enabled         =   -1
            MousePointer    =   0
            Appearance      =   1
            BorderStyle     =   1
            AlignHorizontal =   0
            AlignVertical   =   0
            MultiLine       =   0
            ScrollBars      =   0
            PasswordChar    =   ""
            AllowSpace      =   -1
            Format          =   "a"
            FormatMode      =   1
            AutoConvert     =   -1
            ErrorBeep       =   0
            MaxLength       =   120
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
         Begin TDBText6Ctl.TDBText tdbtCuentaHasta 
            Height          =   315
            Left            =   1200
            TabIndex        =   4
            Tag             =   "_"
            Top             =   960
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   556
            Caption         =   "frmRepLibroMayorAnalitico.frx":1647
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmRepLibroMayorAnalitico.frx":16B3
            Key             =   "frmRepLibroMayorAnalitico.frx":16D1
            BackColor       =   -2147483643
            EditMode        =   0
            ForeColor       =   -2147483640
            ReadOnly        =   0
            ShowContextMenu =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MarginBottom    =   1
            Enabled         =   0
            MousePointer    =   0
            Appearance      =   1
            BorderStyle     =   1
            AlignHorizontal =   0
            AlignVertical   =   0
            MultiLine       =   0
            ScrollBars      =   0
            PasswordChar    =   ""
            AllowSpace      =   -1
            Format          =   "9"
            FormatMode      =   0
            AutoConvert     =   -1
            ErrorBeep       =   0
            MaxLength       =   12
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
         Begin TDBText6Ctl.TDBText tdbtDescripcionHasta 
            Height          =   315
            Left            =   2550
            TabIndex        =   10
            Tag             =   "_"
            Top             =   960
            Width           =   3705
            _Version        =   65536
            _ExtentX        =   6535
            _ExtentY        =   556
            Caption         =   "frmRepLibroMayorAnalitico.frx":1723
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmRepLibroMayorAnalitico.frx":178F
            Key             =   "frmRepLibroMayorAnalitico.frx":17AD
            BackColor       =   -2147483643
            EditMode        =   0
            ForeColor       =   -2147483640
            ReadOnly        =   1
            ShowContextMenu =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MarginBottom    =   1
            Enabled         =   -1
            MousePointer    =   0
            Appearance      =   1
            BorderStyle     =   1
            AlignHorizontal =   0
            AlignVertical   =   0
            MultiLine       =   0
            ScrollBars      =   0
            PasswordChar    =   ""
            AllowSpace      =   -1
            Format          =   "a"
            FormatMode      =   1
            AutoConvert     =   -1
            ErrorBeep       =   0
            MaxLength       =   120
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
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
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
            Left            =   375
            TabIndex        =   12
            Top             =   645
            Width           =   555
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
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
            Left            =   375
            TabIndex        =   11
            Top             =   1005
            Width           =   495
         End
      End
      Begin VB.CheckBox chkMayorAnalitico 
         Caption         =   "LIBRO MAYOR ANALITICO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2640
         TabIndex        =   0
         Top             =   2160
         Width           =   2745
      End
      Begin MSForms.CommandButton cmdSalir 
         Height          =   435
         Left            =   4050
         TabIndex        =   6
         Top             =   5805
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
         Left            =   2145
         TabIndex        =   5
         Top             =   5805
         Width           =   1665
         Caption         =   " Vista Previa"
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
      TabIndex        =   13
      Top             =   0
      Width           =   4365
   End
End
Attribute VB_Name = "frmRepLibroMayorAnalitico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Control As String
Dim MesDesde As String, MesHasta As String
Dim gsGrupo As String
Dim iReport, nAncho As Integer
Dim EntroCabecera As Boolean

Dim pVanDebe As String * 13
Dim pVanHaber As String * 13
Dim pVienenDebe As String * 13
Dim pVienenHaber As String * 13

    Dim sDebeBC As String * 13
    Dim sHaberBC As String * 13

Dim sDebe As String * 13
Dim sHaber As String * 13

Public ReporteSunat As String
Public TituloSunat As String
Public rsArreglo  As New ADODB.Recordset

Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

Public Sub CerrarForm()
Unload Me
End Sub

Private Function Validacion() As Boolean
    Validacion = False
    If tdbcDelMes.BoundText > tdbcAlMes.BoundText Then
        Mensajes "Mes inicio no puede ser mayor que el mes final.", vbInformation
        pSetFocus tdbcDelMes
        Exit Function
    End If
    
    If tdbcAlMes.BoundText < tdbcDelMes.BoundText Then
        Mensajes "Mes final no puede ser menor que mes inicio.", vbInformation
        pSetFocus tdbcDelMes
        Exit Function
    End If
    
    If dtpHasta.Value < dtpDesde.Value And Option2.Value Then
        Mensajes "La fecha HASTA no puede ser menor que la fecha DESDE.", vbInformation
        pSetFocus dtpHasta
        Exit Function
    End If

    If tdbtCuentaDesde.Text > tdbtCuentaHasta.Text And optTodos.Value = vbUnchecked Then
        Mensajes "La cuenta final no puede ser menor que la cuenta de inicio.", vbInformation
        pSetFocus tdbtCuentaHasta
        Exit Function
    End If

    If CE(tdbtCuentaDesde.Text) = "" And optCuentas.Value = True Then
        Mensajes "Ingrese la cuenta inicial para el reporte.", vbInformation
        pSetFocus tdbtCuentaHasta
        Exit Function
    End If

    If CE(tdbtCuentaHasta.Text) = "" And optCuentas.Value = True Then
        Mensajes "Ingrese la cuenta final para el reporte.", vbInformation
        pSetFocus tdbtCuentaHasta
        Exit Function
    End If
    Validacion = True
End Function

Private Sub cmdImprimir_Click()

    Dim matriz_fecha(18) As Variant
    Dim Tipo As String

    If Validacion = False Then Exit Sub

    cmdImprimir.Enabled = False

    Screen.MousePointer = vbHourglass

    nContadorProc = nContadorProc + 1

'    If nContadorProc = 1 Then
'        Call ProcesarSaldos(tdbcAlMes.BoundText)
'    End If

    Screen.MousePointer = vbNormal
    DoEvents
    Screen.MousePointer = vbHourglass
    
    If Me.optCuentas.Value = True Then
        Tipo = "CUENTAS"
    Else
        Tipo = "TODOS"
    End If
    
    If OptImpresion(4).Value Then
        matriz_fecha(0) = "@Emp_cCodigo;" & gsEmpresa & ";True"
        matriz_fecha(1) = "@Pan_cAnio;" & gsAnio & ";True"
        matriz_fecha(2) = "@Per_cPeriodoDesde;" & tdbcDelMes.BoundText & ";True"
        matriz_fecha(3) = "@Per_cPeriodoHasta;" & tdbcAlMes.BoundText & ";True"
        matriz_fecha(4) = "@moneda;" & tdbcMoneda.BoundText & ";True"
        matriz_fecha(5) = "@CtaDesde;" & "" & ";True"
        matriz_fecha(6) = "@CtaHasta;" & "" & ";True"
        matriz_fecha(7) = "@TipoRPTPLE;" & IIf(optTipo(0).Value = True, 1, 0) & ";True"
        matriz_fecha(8) = "@Per_Fechadesde;" & PrimerDiaMes(tdbcDelMes.BoundText, gsAnio) & ";True"
        matriz_fecha(9) = "@Per_Fechahasta;" & PrimerDiaMes(tdbcAlMes.BoundText, gsAnio) & ";True"
        GoTo DCOROI
    End If
    
    
    Call Acumulado
    matriz_fecha(0) = "@Tipo;" & Tipo & ";True"
    matriz_fecha(1) = "@Emp_cCodigo;" & gsEmpresa & ";True"
    matriz_fecha(2) = "@Pan_cAnio;" & gsAnio & ";True"
    
    If Option1.Value Then
        matriz_fecha(3) = "@Per_cPeriodoDesde;" & tdbcDelMes.BoundText & ";True"
        matriz_fecha(4) = "@Per_cPeriodoHasta;" & tdbcAlMes.BoundText & ";True"
    Else
        matriz_fecha(3) = "@Per_cPeriodoDesde;" & "" & ";True"
        matriz_fecha(4) = "@Per_cPeriodoHasta;" & "" & ";True"
    End If
    
    matriz_fecha(5) = "@moneda;" & tdbcMoneda.BoundText & ";True"
    matriz_fecha(6) = "@CtaDesde;" & tdbtCuentaDesde & ";True"
    matriz_fecha(7) = "@CtaHasta;" & tdbtCuentaHasta & ";True"
    matriz_fecha(8) = "@Proceso;;True"
    matriz_fecha(9) = "@EMPRESA;" & gsEmpresaNom & ";True"
    matriz_fecha(10) = "@RUC;" & gsRUC & ";True"
    matriz_fecha(11) = "@NOMMONEDA;" & tdbcMoneda.Text & ";True"
    

    
    matriz_fecha(14) = "@Acumulado_DebeSoles;" & CDbl(sDebeBC) & ";True"
    matriz_fecha(15) = "@Acumulado_HaberSoles;" & CDbl(sHaberBC) & ";True"
    
    Dim sTipo_Lib_Mayor As String
    
    If Me.OptForma(0).Value Then
        sTipo_Lib_Mayor = "0"
    ElseIf Me.OptForma(1).Value Then
        sTipo_Lib_Mayor = "1"
    End If
    
    matriz_fecha(16) = "@Tipo_Lib_Mayor;" & sTipo_Lib_Mayor & ";True"
    
   If Option2.Value Then
        matriz_fecha(12) = "@NOMBREMESINICIAL;" & NombreMes(Format(dtpDesde.Text, "mm")) & ";True"
        matriz_fecha(13) = "@NOMBREMESFINAL;" & NombreMes(Format(dtpHasta.Text, "mm")) & ";True"
        matriz_fecha(17) = "@Per_Fechadesde;" & Format(dtpDesde.Text, "dd/mm/yyyy") & ";True"
        matriz_fecha(18) = "@Per_Fechahasta;" & Format(dtpHasta.Text, "dd/mm/yyyy") & ";True"
    Else
        matriz_fecha(12) = "@NOMBREMESINICIAL;" & tdbcDelMes.Text & ";True"
        matriz_fecha(13) = "@NOMBREMESFINAL;" & tdbcAlMes.Text & ";True"
        matriz_fecha(17) = "@Per_Fechadesde;" & "" & ";True"
        matriz_fecha(18) = "@Per_Fechahasta;" & "" & ";True"
    End If
    
'    matriz_fecha(14) = "@Emp_cCodigo_BC;" & gsEmpresa & ";True"
'    matriz_fecha(15) = "@Pan_cAnio_BC;" & gsAnio & ";True"
'    matriz_fecha(16) = "@Per_cPeriodo_BC;" & tdbcAlMes.BoundText & ";True"
'    matriz_fecha(17) = "@moneda_BC;" & tdbcMoneda.BoundText & ";True"
'    matriz_fecha(18) = "@tipo_BC;ACUMULADO;True"
'    matriz_fecha(19) = "@NroDigitos_BC;2;True"

    cmdImprimir.Enabled = False
    Dim formulas(0) As Variant
    
DCOROI:
    If optTipo(0).Value Then
        AbreReporteParam gsDSN, Me, rutaReportes & "RptMayorElectronicoDCO.rpt", crptToWindow, "Libro Mayor - Detalle Centralizado Operaciones", "", matriz_fecha(), formulas()
    ElseIf optTipo(1).Value Then
        AbreReporteParam gsDSN, Me, rutaReportes & "RptMayorElectronicoROI.rpt", crptToWindow, "Libro Mayor - Reporte de Operaciones Individuales", "", matriz_fecha(), formulas()
    End If
    
    
    If chkMayorAnalitico.Value = vbChecked Then
     If OptImpresion(0).Value Then 'Matricial
       'gsNombreVista = "Libro Mayor Analitico"
       'ImpMatMayAnalitico
     ElseIf OptImpresion(1).Value Then 'Laser
       AbreReporteParam gsDSN, Me, rutaReportes & "RptLibroMayorAnaliticoAnual.rpt", crptToWindow, "Libro Mayor General Analitico", "", matriz_fecha(), formulas()
     End If
    ElseIf ReporteSunat = "F0601" Then
     If OptImpresion(0).Value Then 'Matricial
       gsNombreVista = "Libro Mayor General"
       ImpMatMayGeneral
     ElseIf OptImpresion(1).Value Then 'Laser
        AbreReporteParam gsDSN, Me, rutaReportes & "RptFormato_0601.rpt", crptToWindow, "Anexo al Libro de Inventarios y Balances", "", matriz_fecha(), formulas()
     End If
    
    End If
    Screen.MousePointer = vbDefault
    cmdImprimir.Enabled = True

    ' ***
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))

    Me.Caption = Titulo(Me.Caption, TituloSunat)
    Call Centrar_form(Me)
    dtpHasta = FechaServidor

    If (Mid(dtpHasta, 4, 2) = "02") Then
        Dim strBisiesto As String
        strBisiesto = Right(dtpHasta, 4)
        
        If ((Right(gsAnio, 2) = "00" Or (Right(gsAnio, 2) Mod 4) = 0) And (Right(gsAnio, 2) <> "00") Or (Right(gsAnio, 2) Mod 4) = 0) Then
            dtpHasta = "29/02/" & gsAnio
        Else
            dtpHasta = "28/02/" & gsAnio
        End If
        
'        If (gsAnio <> strBisiesto) Then
'            dtpHasta = "28/02/" & strBisiesto
'        End If
        
    End If
    
    If Year(dtpHasta) <> gsAnio Then dtpHasta = Mid(dtpHasta, 1, 6) & gsAnio
    dtpDesde = dtpHasta
    Call LlenaCombos
    tdbcDelMes.BoundText = gsPeriodo
    tdbcAlMes.BoundText = gsPeriodo  'Format(Month(fechaServidor), "00")
'    tdbcMes.BoundText = gsPeriodo
    
    tdbcMoneda.BoundText = gsMonedaNac
    
    tdbcAlMes.ReBind
    tdbcDelMes.ReBind
    tdbcMoneda.ReBind
    
    Option1_Click
    Option2_Click
    optTodos_Click

    On Error Resume Next
    
    dtpDesde.Value = "01/" & gsPeriodo & "/" & gsAnio
    dtpHasta.Value = "01/" & gsPeriodo & "/" & gsAnio
    dtpDesde.MaxDate = "31/12/" & gsAnio
    dtpHasta.MaxDate = "31/12/" & gsAnio
    dtpDesde.MinDate = "01/01/" & gsAnio
    dtpHasta.MinDate = "01/01/" & gsAnio
    
    dtpDesde.Enabled = False
    dtpHasta.Enabled = False
        
    tdbcDelMes.BoundText = gsPeriodo
    
    tdbcDelMes.ReBind
    tdbcAlMes.ReBind
    
    tdbcMoneda.ReBind
End Sub

Private Sub LlenaCombos()
    Dim sqlcombos As String
    Call LlenaComboMesApeAddItem(tdbcDelMes)
    Call LlenaComboMesApeAddItem(tdbcAlMes)
    ' *** Llenando el tipo de Moneda
    sqlcombos = "SELECT Mon_cCodigo, Mon_cNombreLargo, Mon_cMNac, Mon_cMExt From CNT_TIPO_MONEDA " & _
                "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND (Mon_cMNac = '1' or Mon_cMExt = '1') " & _
                "ORDER BY Mon_cNombreLargo"
    LlenarComboAddItem tdbcMoneda, sqlcombos
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
        Call Centrar_Objeto(fratodo, Me)
        Call CentrarTitulo(lblTitulo, fratodo, Me)
    End If
Exit Sub
errHand:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmRepLibroMayorAnalitico = Nothing
    Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))

End Sub


Private Sub optCuentas_Click()
    tdbtCuentaDesde.Enabled = True
    tdbtCuentaHasta.Enabled = True
    pSetFocus tdbtCuentaDesde
    
    tdbtDescripcionDesde.BackColor = gsColorDesactivado
    tdbtDescripcionHasta.BackColor = gsColorDesactivado
    tdbtCuentaDesde.BackColor = gsColorActivado
    tdbtCuentaHasta.BackColor = gsColorActivado
    
End Sub

Private Sub OptImpresion_Click(Index As Integer)
If OptImpresion(0).Value Or OptImpresion(1).Value Then
    optTipo(0).Enabled = False
    optTipo(1).Enabled = False
    optTipo(0).Value = False
    optTipo(1).Value = False
ElseIf OptImpresion(4).Value Then
    optTipo(0).Enabled = True
    optTipo(1).Enabled = True
End If
End Sub

Private Sub Option1_Click()
If Option1.Value Then
    ActivarControl dtpDesde, False
    ActivarControl dtpHasta, False
    ActivarControl tdbcDelMes, True
    ActivarControl tdbcAlMes, True

Else
    ActivarControl dtpDesde, True
    ActivarControl dtpHasta, True
    ActivarControl tdbcDelMes, False
    ActivarControl tdbcAlMes, False
End If
End Sub

Private Sub Option2_Click()
If Option2.Value Then
    ActivarControl dtpDesde, True
    ActivarControl dtpHasta, True
    ActivarControl tdbcDelMes, False
    ActivarControl tdbcAlMes, False
    Dim Mes As String
    Mes = gsPeriodo
    If Mes = "00" Then Mes = "01"
    If Mes > "12" Then Mes = "12"
    On Error Resume Next
    dtpHasta = UltimoDiaMes(Mes, gsAnio)
    OptImpresion(4).Enabled = False
Else
    ActivarControl dtpDesde, False
    ActivarControl dtpHasta, False
    ActivarControl tdbcDelMes, True
    ActivarControl tdbcAlMes, True
End If
End Sub

Private Sub optTipo_Click(Index As Integer)
Frame5.Enabled = True
End Sub

Private Sub optTodos_Click()
    tdbtCuentaDesde.Text = ""
    tdbtCuentaHasta.Text = ""
    tdbtDescripcionDesde.Text = ""
    tdbtDescripcionHasta.Text = ""
    
    tdbtCuentaDesde.Enabled = False
    tdbtCuentaHasta.Enabled = False
    
    tdbtDescripcionDesde.BackColor = gsColorDesactivado
    tdbtDescripcionHasta.BackColor = gsColorDesactivado
    tdbtCuentaDesde.BackColor = gsColorDesactivado
    tdbtCuentaHasta.BackColor = gsColorDesactivado
End Sub

Private Sub tdbcAlMes_ItemChange()
If (tdbcDelMes.BoundText = "00" Or tdbcDelMes.BoundText = "13" Or tdbcDelMes.BoundText = "14") And _
        (tdbcAlMes.BoundText = "00" Or tdbcAlMes.BoundText = "13" Or tdbcAlMes.BoundText = "14") Then
    OptImpresion(4).Enabled = False
    OptImpresion(4).Value = False
    optTipo(0).Enabled = False
    optTipo(1).Enabled = False
    optTipo(0).Value = False
    optTipo(1).Value = False
Else
    If tdbcAlMes.Text = tdbcDelMes.Text Or ((tdbcDelMes.BoundText = "00" And tdbcAlMes.BoundText = "01") Or (tdbcDelMes.BoundText = "12" And (tdbcAlMes.BoundText = "13" Or tdbcAlMes.BoundText = "14"))) Then
        OptImpresion(4).Enabled = True
    Else
        OptImpresion(4).Enabled = False
        OptImpresion(4).Value = False
        optTipo(0).Enabled = False
        optTipo(1).Enabled = False
        optTipo(0).Value = False
        optTipo(1).Value = False
    End If
End If
End Sub

Private Sub tdbcAlMes_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then pSendKeys "{tab}"
End Sub

Private Sub tdbcDelMes_ItemChange()
If (tdbcDelMes.BoundText = "00" Or tdbcDelMes.BoundText = "13" Or tdbcDelMes.BoundText = "14") And _
        (tdbcAlMes.BoundText = "00" Or tdbcAlMes.BoundText = "13" Or tdbcAlMes.BoundText = "14") Then
    OptImpresion(4).Enabled = False
    OptImpresion(4).Value = False
    optTipo(0).Enabled = False
    optTipo(1).Enabled = False
    optTipo(0).Value = False
    optTipo(1).Value = False
Else
    If tdbcAlMes.Text = tdbcDelMes.Text Or ((tdbcDelMes.BoundText = "00" And tdbcAlMes.BoundText = "01") Or (tdbcDelMes.BoundText = "12" And (tdbcAlMes.BoundText = "13" Or tdbcAlMes.BoundText = "14"))) Then
        OptImpresion(4).Enabled = True
    Else
        OptImpresion(4).Enabled = False
        OptImpresion(4).Value = False
        optTipo(0).Enabled = False
        optTipo(1).Enabled = False
        optTipo(0).Value = False
        optTipo(1).Value = False
    End If
End If
End Sub

Private Sub tdbcDelMes_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then pSendKeys "{tab}"
End Sub

Private Sub tdbcMoneda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then pSendKeys "{tab}"
End Sub

Private Sub tdbtCuentaDesde_Change()
 If CE(tdbtCuentaDesde) = "" Then tdbtDescripcionDesde = ""
End Sub

Private Sub tdbtCuentaDesde_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then
        MesDesde = tdbcDelMes.BoundText
        MesHasta = tdbcAlMes.BoundText
        Call LlamaBuscar(frmBuscador, Me.tdbtCuentaDesde.Name, Control, "Cuentas", Me, gsPeriodo, Me.tdbtCuentaDesde.Text)
        
    End If
End Sub

Private Sub tdbtCuentaDesde_LostFocus()
    ' *** Verificando q cuenta exista y q sea de no titulo
    If tdbtCuentaDesde <> "" And Me.Enabled = True Then
        tdbtDescripcionDesde = ExisteCtaNoTitulo(tdbtCuentaDesde, "")
        If tdbtDescripcionDesde = "" Then pSetFocus tdbtCuentaDesde
    End If
End Sub

Private Sub tdbtCuentaHasta_Change()
    If CE(tdbtCuentaHasta) = "" Then tdbtDescripcionHasta = ""
End Sub

Private Sub tdbtCuentaHasta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then Call LlamaBuscar(frmBuscador, Me.tdbtCuentaHasta.Name, Control, "Cuentas", Me, gsPeriodo, Me.tdbtCuentaHasta.Text)
End Sub

Private Sub tdbtCuentaHasta_LostFocus()
    ' *** Verificando q cuenta exista y q sea de no titulo
    If tdbtCuentaHasta <> "" And Me.Enabled = True Then
        tdbtDescripcionHasta = ExisteCtaNoTitulo(tdbtCuentaHasta, "")
        If tdbtDescripcionHasta = "" Then pSetFocus tdbtCuentaHasta
    End If
End Sub

Public Sub RecibirDatos(ByVal lControl As String, ByVal param0 As String, ByVal param1 As String, ByVal param2 As String, Optional ByVal param3 As String)
   ' *** Dependiendo del control
    Select Case Control
    Case "tdbtCuentaDesde" ' *** Caso Desde
        tdbtCuentaDesde = Trim(param0)
        tdbtDescripcionDesde = Trim(param1)
        Unload frmBuscador
        pSetFocus tdbtCuentaDesde
    Case "tdbtCuentaHasta" ' *** Caso Hasta
        tdbtCuentaHasta = Trim(param0)
        tdbtDescripcionHasta = Trim(param1)
        Unload frmBuscador
        pSetFocus tdbtCuentaHasta
    End Select
    DoEvents
    'tdbcDelMes.BoundText = MesDesde
    'tdbcAlMes.BoundText = MesHasta
    
End Sub
Sub ImpMatMayGeneral()
On Error Resume Next

 gsAccionRep = 5

 gsCodMoneda = tdbcMoneda.BoundText
 gsLdMesIni = tdbcDelMes.BoundText: gsLdMesFin = tdbcAlMes.BoundText

 If optTodos.Value Then
  gsCtaIni = "": gsCtaFin = ""
 ElseIf optCuentas.Value Then
  gsCtaIni = tdbtCuentaDesde.Text: gsCtaFin = tdbtCuentaHasta.Text
 End If
 
 frmFCImpresion.Show

End Sub

Public Sub ReporteMatMayGeneral()
Dim pPla_cCuentaContable As String * 12
Dim pAsd_dFecDoc As String * 10
Dim pAse_nVoucher As String * 10
Dim pAsd_cGlosa As String * 39
Dim pAsd_cTipoDoc As String * 2
Dim pAsd_cSerieDoc As String * 3
Dim pAsd_cNumDoc As String * 8
Dim pPla_cNombreCuenta As String * 16
Dim pD3, pD2 As String * 3
Dim pD3Ant, pD2Ant As String * 3
Dim pPer_cPeriodo As String * 3
Dim SpaceMes, SpaceCta, Cont As Long
Dim SpaceMayCta As Integer
Dim SpaceSldo As Integer
Dim SpaceSumasMeses As Integer

Dim pPla_cNombreCuentaD3, pPla_cNombreCuentaD2 As String * 16

Dim pDebeDet As String * 12
Dim pHaberDet As String * 12
Dim pSumDebeDet As String * 12
Dim pSumHaberDet As String * 12
Dim pSumSldoFinDebeDet As String * 12
Dim pSumSldoFinHaberDet As String * 12
Dim pDebeFin As String * 12
Dim pHaberFin As String * 12

Dim pSaldoDebeD3 As String * 12
Dim pxSaldoDebeD3 As String * 12
Dim pSaldoHaberD3 As String * 12
Dim pxSaldoHaberD3 As String * 12

Dim pSaldoDebeD2 As String * 12
Dim pxSaldoDebeD2 As String * 12
Dim pSaldoHaberD2 As String * 12
Dim pxSaldoHaberD2 As String * 12

Dim pSumMayDebeDet As String * 12
Dim pSumMayHaberDet As String * 12
Dim pSumasMesesDebe As String * 12
Dim pSumasMesesHaber As String * 12

 pPla_cCuentaContable = ""
 
 'On Error GoTo ERROR
 Screen.MousePointer = vbHourglass
 
 If Not ExistenDatos() Then
  MsgBox "No existen Datos para Imprimir el Reporte.", vbExclamation, App.Title
  Exit Sub
 End If
 
 If frmFCImpresion.List_Destino.Text = "Archivo" Then
   Open frmFCImpresion.OutputFileName For Output Shared As #1
   gsPagina = 0
 End If
 
'Print #1, Chr(27) & Chr(64); 'Inicializa
'Print #1, Chr(27) & Chr(120) & Chr(0); 'Draft
'Print #1, Chr(27) & Chr(15); 'Comprimido
  ' Print #1, Chr(27) & Chr(77); '12cpi
'Print #1, Chr(27) & Chr(51) & Chr(29) 'Entre lineas 29/180

 
 giLineas = 0
 giEspacios = 60
 
 sDebe = 0
 sHaber = 0
     
 Dim VarTotal As Double
 Dim VartotalH As Double
 If Not rsArreglo.EOF Then iReport = 1
 'If iReport = 1 Then rsArreglo.Sort = "Asd_dFecDoc": rsArreglo.MoveFirst
 Dim conta As Long
 conta = 0
 With rsArreglo
    If .RecordCount > 0 Then
       iReport = 1
       Debug.Print conta
       gsPaginaPrincipal = 1
       Call CabeceraLibMayorGeneral
       EntroCabecera = False
       .MoveFirst
        
       pSaldoHaberD3 = 0: pxSaldoHaberD3 = 0
       pSaldoDebeD3 = 0: pxSaldoDebeD3 = 0
        
       pSaldoHaberD2 = 0: pxSaldoHaberD2 = 0
       pSaldoDebeD2 = 0: pxSaldoDebeD2 = 0
        
       pVanDebe = 0: pVanHaber = 0
       pVienenDebe = 0: pVienenHaber = 0
        
       Do While Not .EOF
       conta = conta + 1
          pPla_cCuentaContable = !Pla_cCuentaContable
          pD3 = Trim(!d3)
          pD2 = Trim(!d2)
          'Debug.Print "LINEA 1"
          pSumDebeDet = 0: pSumHaberDet = 0
          pSumSldoFinDebeDet = 0: pSumSldoFinHaberDet = 0
          pDebeFin = 0: pHaberFin = 0
          'Debug.Print "LINEA 2"
          pSumasMesesDebe = 0: pSumasMesesHaber = 0
          'Debug.Print "LINEA 3"
          If Trim(!d3) <> Trim(pD3Ant) Then
            pSaldoHaberD3 = 0: pxSaldoHaberD3 = 0
            pSaldoDebeD3 = 0: pxSaldoDebeD3 = 0
          End If
          'Debug.Print "LINEA 4"
          If Trim(!d2) <> Trim(pD2Ant) Then
            pSaldoHaberD2 = 0: pxSaldoHaberD2 = 0
            pSaldoDebeD2 = 0: pxSaldoDebeD2 = 0
            pSumMayDebeDet = 0: pSumMayHaberDet = 0
          End If
          'Debug.Print "LINEA 4"
          Do While Not .EOF
'          If !Pla_cCuentaContable = "10201001" Then MsgBox "cvd"
            If Trim$(!Pla_cCuentaContable) = Trim$(pPla_cCuentaContable) And Trim$(!d3) = Trim$(pD3) Then
                'Debug.Print "LINEA 5"
                 If gsPaginaPrincipal = 1 And tdbcDelMes.BoundText > "00" Then
                    Call ImprimeVanVienen(1, tdbcDelMes.BoundText, Trim$(pPla_cCuentaContable))
                    gsPaginaPrincipal = 9
                 End If
                 EntroCabecera = False
                 pPla_cNombreCuenta = !Pla_cNombreCuenta
                 pAsd_dFecDoc = IIf(IsNull(!Asd_dFecDoc) = True, "", !Asd_dFecDoc)
                 pAse_nVoucher = !Ase_nVoucher
                 If IsNull(!Asd_cGlosa) Then
                    pAsd_cGlosa = ""
                Else
                    pAsd_cGlosa = (!Asd_cGlosa)
                End If
                 pAsd_cTipoDoc = !Asd_cTipoDoc
                 pAsd_cSerieDoc = !Asd_cSerieDoc
                 pAsd_cNumDoc = !Asd_cNumDoc
                 'Debug.Print "LINEA 6"
                 If Trim(!Lib_cTipoLibro) = "XX" Then
                    If Trim(!Mon_cMNac) = "1" Then
                        If !SaldoAntMonNac > 0 Then RSet pDebeDet = Format(Abs(!SaldoAntMonNac), "#,###,###,##0.00;(#,###,###,##0.00)") Else RSet pDebeDet = Format(0, "#,###,###,##0.00;(#,###,###,##0.00)")
                    Else
                        If !SaldoAntMonExt > 0 Then RSet pDebeDet = Format(Abs(!SaldoAntMonExt), "#,###,###,##0.00;(#,###,###,##0.00)") Else RSet pDebeDet = Format(0, "#,###,###,##0.00;(#,###,###,##0.00)")
                    End If
                 Else
                    If Trim(!Mon_cMNac) = "1" Then RSet pDebeDet = Format(!Asd_nDebeSoles, "#,###,###,##0.00;(#,###,###,##0.00)") Else RSet pDebeDet = Format(!Asd_nDebeMonExt, "#,###,###,##0.00;(#,###,###,##0.00)")
                    End If
                 'Debug.Print "LINEA 6-1"
                 If Trim(!Lib_cTipoLibro) = "XX" Then
                  If Trim(!Mon_cMNac) = "1" Then
                   If !SaldoAntMonNac < 0 Then RSet pHaberDet = Format(Abs(!SaldoAntMonNac), "#,###,###,##0.00;(#,###,###,##0.00)") Else RSet pHaberDet = Format(0, "#,###,###,##0.00;(#,###,###,##0.00)")
                  Else
                   If !SaldoAntMonExt < 0 Then RSet pHaberDet = Format(Abs(!SaldoAntMonExt), "#,###,###,##0.00;(#,###,###,##0.00)") Else RSet pHaberDet = Format(0, "#,###,###,##0.00;(#,###,###,##0.00)")
                  End If
                 Else
                  If Trim(!Mon_cMNac) = "1" Then RSet pHaberDet = Format(!Asd_nHaberSoles, "#,###,###,##0.00;(#,###,###,##0.00)") Else RSet pHaberDet = Format(!Asd_nHaberMonExt, "#,###,###,##0.00;(#,###,###,##0.00)")
                 End If
                 'Debug.Print "LINEA 6-2"
                 If Trim(pPer_cPeriodo) = Trim(!Per_cPeriodo) And Trim(!Per_cPeriodo) <> "" Then
                    If Trim(!Lib_cTipoLibro) <> "XX" And Trim(!Per_cPeriodo) = "" Then
                     pPla_cCuentaContable = "": pPla_cNombreCuenta = ""
                    'Debug.Print "LINEA 6-2-1"
                     pSumDebeDet = CDbl(pSumDebeDet) + CDbl(pDebeDet)
                     pSumHaberDet = CDbl(pSumHaberDet) + CDbl(pHaberDet)
                    'Debug.Print "LINEA 6-2-2"
                    Else
                     If Trim(!Lib_cTipoLibro) <> "XX" And Trim(!Per_cPeriodo) <> "" Then
                        If Cont >= 1 Then pPla_cCuentaContable = "": pPla_cNombreCuenta = ""
                           'Debug.Print "LINEA 6-2-3"

                            pSumDebeDet = CDbl(pSumDebeDet) + CDbl(pDebeDet)
                            pSumHaberDet = CDbl(pSumHaberDet) + CDbl(pHaberDet)
    
                           'If f = 199 Then MsgBox "dvd"
                            Cont = Cont + 1
                        End If
                    End If
                    
                    If Err.Description <> "" Then MsgBox "ERROR: GG"
                 Else
                 'Debug.Print "LINEA 6-3"
                 ImprimeVanVienen
                  '  Debug.Print "LINEA 7"
                    If Trim(pPer_cPeriodo) <> Trim(!Per_cPeriodo) And Trim(!Per_cPeriodo) <> "" And Trim(pPer_cPeriodo) <> "" Then
                     If CDbl(pSumDebeDet) <> 0 Or CDbl(pSumHaberDet) <> 0 Then printl Space(3) & (Space(104) & "------------    ------------")
                       If Len(NombreMes(Trim(pPer_cPeriodo))) > Len("ENERO") Then
                        SpaceMes = 41 - (Len(NombreMes(Trim(pPer_cPeriodo))) - Len("ENERO"))
                       ElseIf Len(NombreMes(Trim(pPer_cPeriodo))) < Len("ENERO") Then
                        SpaceMes = 43 - (Len("ENERO") - Len(NombreMes(Trim(pPer_cPeriodo))))
                       ElseIf Len(NombreMes(Trim(pPer_cPeriodo))) = Len("ENERO") Then
                        SpaceMes = 41
                       End If
                       RSet pSumDebeDet = Format(pSumDebeDet, "#,###,###,#0.00;(#,###,###,#0.00)")
                       RSet pSumHaberDet = Format(pSumHaberDet, "#,###,###,#0.00;(#,###,###,#0.00)")
                       
'                       sDebe = cdbl(sDebe) + cdbl(pSumDebeDet)
'                       sHaber = cdbl(sHaber) + cdbl(pSumHaberDet)
                       
                       If CDbl(pSumDebeDet) <> 0 Or CDbl(pSumHaberDet) <> 0 Then
                        printl (Space(3) & Space(48) & "PERIODO : " & NombreMes(Trim(pPer_cPeriodo)) & Space(SpaceMes) & pSumDebeDet & Space(4) & pSumHaberDet)
                        sDebe = CDbl(sDebe) + CDbl(pSumDebeDet)
                        sHaber = CDbl(sHaber) + CDbl(pSumHaberDet)
                       End If
                       
                    pSumasMesesDebe = CDbl(pSumasMesesDebe) + CDbl(pSumDebeDet): RSet pSumasMesesDebe = Format(pSumasMesesDebe, "#,###,###,##0.00;(#,###,###,##0.00)")
                    pSumasMesesHaber = CDbl(pSumasMesesHaber) + CDbl(pSumHaberDet): RSet pSumasMesesHaber = Format(pSumasMesesHaber, "#,###,###,##0.00;(#,###,###,##0.00)")
                    'Debug.Print "LINEA 8"
                    
                    pSumDebeDet = 0: pSumHaberDet = 0
                    End If
                 End If
                 pSumSldoFinDebeDet = CDbl(pSumSldoFinDebeDet) + CDbl(pDebeDet)
                 pSumSldoFinHaberDet = CDbl(pSumSldoFinHaberDet) + CDbl(pHaberDet)
                 
                 If Trim(!d2) = Trim(pD2) Then
                        pSumMayDebeDet = CDbl(pSumMayDebeDet) + CDbl(pDebeDet)
                        pSumMayHaberDet = CDbl(pSumMayHaberDet) + CDbl(pHaberDet)
                   ' If LTrim(RTrim(pAsd_dFecDoc)) <> "" Then
'                        pSumMayDebeDet = CDbl(pSumMayDebeDet) + CDbl(pDebeDet)
'                        pSumMayHaberDet = CDbl(pSumMayHaberDet) + CDbl(pHaberDet)
                        
                        pxSaldoDebeD2 = CDbl(pxSaldoDebeD2) + pDebeDet
                        pxSaldoHaberD2 = CDbl(pxSaldoHaberD2) + pHaberDet
                        
                        pSaldoDebeD2 = CDbl(pxSaldoDebeD2) - CDbl(pxSaldoHaberD2)
                        pSaldoDebeD2 = CDbl(pSaldoDebeD2) * -1
                        If CDbl(pSaldoDebeD2) < 0 Then RSet pSaldoDebeD2 = Format(Abs(pSaldoDebeD2), "#,###,###,##0.00;(#,###,###,##0.00)") Else RSet pSaldoDebeD2 = Format(0, "#,###,###,##0.00;(#,###,###,##0.00)")
                        
                        pSaldoHaberD2 = CDbl(pxSaldoDebeD2) - CDbl(pxSaldoHaberD2)
                        pSaldoHaberD2 = CDbl(pSaldoHaberD2) * -1
                        If CDbl(pSaldoHaberD2) > 0 Then RSet pSaldoHaberD2 = Format(Abs(pSaldoHaberD2), "#,###,###,##0.00;(#,###,###,##0.00)") Else RSet pSaldoHaberD2 = Format(0, "#,###,###,##0.00;(#,###,###,##0.00)")
                        
                   ' End If
                  'Debug.Print "LINEA 9"
                 End If
                                                                                                                       
                 If Trim(pPer_cPeriodo) <> Trim(!Per_cPeriodo) And Trim(!Per_cPeriodo) <> "" Then
                    If Trim(!Lib_cTipoLibro) <> "XX" And Trim(!Per_cPeriodo) = "" Then
                     pPla_cCuentaContable = "": pPla_cNombreCuenta = ""
                     
                     pSumDebeDet = CDbl(pSumDebeDet) + CDbl(pDebeDet)
                     pSumHaberDet = CDbl(pSumHaberDet) + CDbl(pHaberDet)

                    Else
                     If Trim(!Lib_cTipoLibro) <> "XX" And Trim(!Per_cPeriodo) <> "" Then
                       If Cont >= 1 Then pPla_cCuentaContable = "": pPla_cNombreCuenta = ""
                       
                        pSumDebeDet = CDbl(pSumDebeDet) + CDbl(pDebeDet)
                        pSumHaberDet = CDbl(pSumHaberDet) + CDbl(pHaberDet)
                       
                       Cont = Cont + 1
                     End If
                    End If
                 End If
                
                If gsPaginaPrincipal = 9 Then
                    Dim rst As New ADODB.Recordset
                    Set rst = Fct_Obt_Sumas_DH_Sg_Periodo(gsEmpresa, gsAnio, tdbcDelMes.BoundText, Trim$(pPla_cCuentaContable))
                    If rst.EOF = False Then
                        pVanDebe = Format(Val(rst("Debe").Value), "#,###,###,##0.00;(#,###,###,##0.00)")
                        pVanHaber = Format(Val(rst("Haber").Value), "#,###,###,##0.00;(#,###,###,##0.00)")
                    End If
                    gsPaginaPrincipal = 0
                Else
                    If LTrim(RTrim(pAsd_dFecDoc)) <> "" Then
                        pVanDebe = CDbl(pVanDebe) + CDbl(pDebeDet): RSet pVanDebe = Format$(pVanDebe, "#,###,###,##0.00;(#,###,###,##0.00)")
                        pVanHaber = CDbl(pVanHaber) + CDbl(pHaberDet): RSet pVanHaber = Format$(pVanHaber, "#,###,###,##0.00;(#,###,###,##0.00)")
                    End If
                End If
                
                 ImprimeVanVienen
                 
                 'If pPla_cCuentaContable = "95109030" Then MsgBox "vdvd"
                 Dim iLongitudGlosa As Integer
                 iLongitudGlosa = 35
                 gsLinea = (Space(3) & pPla_cCuentaContable & Space(1) & pPla_cNombreCuenta & Space(1) & pAsd_dFecDoc & Space(1) & pAse_nVoucher & Space(1) & Left(pAsd_cGlosa, iLongitudGlosa) & Space(1) & pAsd_cTipoDoc & Space(1) & pAsd_cSerieDoc & Space(1) & pAsd_cNumDoc & Space(1) & pDebeDet & Space(4) & pHaberDet)
                 'Imprime la primera linea del libro
                 printl gsLinea
                 Debug.Print gsLinea
                 'If Trim(!d3) = Trim(pD3) Then

                  pxSaldoDebeD3 = CDbl(pxSaldoDebeD3) + pDebeDet
                  pxSaldoHaberD3 = CDbl(pxSaldoHaberD3) + pHaberDet
                 'End If
                  pSaldoDebeD3 = CDbl(pxSaldoDebeD3) - CDbl(pxSaldoHaberD3)
                  pSaldoDebeD3 = CDbl(pSaldoDebeD3) * -1

                  If CDbl(pSaldoDebeD3) < 0 Then RSet pSaldoDebeD3 = Format(Abs(pSaldoDebeD3), "#,###,###,##0.00;(#,###,###,##0.00)") Else RSet pSaldoDebeD3 = Format(0, "#,###,###,##0.00;(#,###,###,##0.00)")

                  pSaldoHaberD3 = CDbl(pxSaldoDebeD3) - CDbl(pxSaldoHaberD3)
                  pSaldoHaberD3 = CDbl(pSaldoHaberD3) * -1
                  If CDbl(pSaldoHaberD3) > 0 Then RSet pSaldoHaberD3 = Format(Abs(pSaldoHaberD3), "#,###,###,##0.00;(#,###,###,##0.00)") Else RSet pSaldoHaberD3 = Format(0, "#,###,###,##0.00;(#,###,###,##0.00)")

                 pPla_cCuentaContable = Trim(!Pla_cCuentaContable)
                 
                 If Trim(!Lib_cTipoLibro) <> "XX" Then
                  pPer_cPeriodo = Trim(!Per_cPeriodo)
                 Else
                  .MoveNext
                  
                  If .EOF Then 'Fin del Archivo
                    If CDbl(pSumDebeDet) <> 0 Or CDbl(pSumHaberDet) <> 0 Then printl (Space(3) & Space(104) & "------------    ------------")
                    
                     If Len(NombreMes(Trim(pPer_cPeriodo))) > Len("ENERO") Then
                      SpaceMes = 46 - Len(NombreMes(Trim(pPer_cPeriodo)))
                     ElseIf Len(NombreMes(Trim(pPer_cPeriodo))) < Len("ENERO") Then
                      SpaceMes = 44 - (Len("ENERO") - Len(NombreMes(Trim(pPer_cPeriodo))))
                     ElseIf Len(NombreMes(Trim(pPer_cPeriodo))) = Len("ENERO") Then
                      SpaceMes = 41
                     End If
                     
                     RSet pSumDebeDet = Format(pSumDebeDet, "#,###,###,##0.00;(#,###,###,##0.00)")
                     RSet pSumHaberDet = Format(pSumHaberDet, "#,###,###,##0.00;(#,###,###,##0.00)")
                     
                     RSet pSumMayDebeDet = Format(pSumMayDebeDet, "#,###,###,##0.00;(#,###,###,##0.00)")
                     RSet pSumMayHaberDet = Format(pSumMayHaberDet, "#,###,###,##0.00;(#,###,###,##0.00)")
                     
                     pDebeFin = CDbl(pSumSldoFinDebeDet) - CDbl(pSumSldoFinHaberDet)
                     pDebeFin = CDbl(pDebeFin) * -1
                     If CDbl(pDebeFin) < 0 Then RSet pDebeFin = Format(Abs(pDebeFin), "#,###,###,##0.00;(#,###,###,##0.00)") Else RSet pDebeFin = Format(0, "#,###,###,##0.00;(#,###,###,##0.00)")
                     
                     pHaberFin = CDbl(pSumSldoFinDebeDet) - CDbl(pSumSldoFinHaberDet)
                     pHaberFin = CDbl(pHaberFin) * -1
                     If CDbl(pHaberFin) > 0 Then RSet pHaberFin = Format(Abs(pHaberFin), "#,###,###,##0.00;(#,###,###,##0.00)") Else RSet pHaberFin = Format(0, "#,###,###,##0.00;(#,###,###,##0.00)")
                     
                     If CDbl(pSumDebeDet) <> 0 Or CDbl(pSumHaberDet) <> 0 Then
                        printl (Space(3) & Space(48) & "PERIODO : " & NombreMes(Trim(pPer_cPeriodo)) & Space(SpaceMes) & pSumDebeDet & Space(4) & pSumHaberDet)
                        sDebe = CDbl(sDebe) + CDbl(pSumDebeDet)
                        sHaber = CDbl(sHaber) + CDbl(pSumHaberDet)
                     End If
    
                     pSumasMesesDebe = CDbl(pSumasMesesDebe) + CDbl(pSumDebeDet): RSet pSumasMesesDebe = Format(pSumasMesesDebe, "#,###,###,##0.00;(#,###,###,##0.00)")
                     pSumasMesesHaber = CDbl(pSumasMesesHaber) + CDbl(pSumHaberDet): RSet pSumasMesesHaber = Format(pSumasMesesHaber, "#,###,###,##0.00;(#,###,###,##0.00)")
                     
                     SpaceSumasMeses = 47 - Len(NombreMes(Trim(gsLdMesIni)) & " A " & NombreMes(Trim(gsLdMesFin)))
                     If Trim(gsLdMesIni) <> Trim(gsLdMesFin) And Trim(pAse_nVoucher) <> "" Then
                        printl (Space(3) & Space(48) & "SUMAS DE " & NombreMes(Trim(gsLdMesIni)) & " A " & NombreMes(Trim(gsLdMesFin)) & Space(SpaceSumasMeses) & pSumasMesesDebe & Space(4) & pSumasMesesHaber)
                     End If
        
                     If Trim(gsLdMesIni) = Trim(gsLdMesFin) And Len(NombreMes(Trim(pPer_cPeriodo))) = 0 Then pPer_cPeriodo = Trim(gsLdMesIni)
                     SpaceSldo = 42 - Len(NombreMes(Trim(pPer_cPeriodo)))
                 '    Debug.Print "LINEA 12"
                     
                     'printl (pPla_cCuentaContable & Space(1) & pPla_cNombreCuenta & Space(23) & "SALDO FINAL A" & Space(43) & pDebeFin & Space(1) & pHaberFin)
                     If Trim(gsLdMesIni) = Trim(gsLdMesFin) And Len(NombreMes(Trim(pPer_cPeriodo))) = 0 Then pPer_cPeriodo = Trim(gsLdMesIni)
                     printl (pPla_cCuentaContable & Space(1) & pPla_cNombreCuenta & Space(23) & "SALDO FINAL A " & NombreMes(Trim(pPer_cPeriodo)) & Space(SpaceSldo) & pDebeFin & Space(1) & pHaberFin)
                     
                     SpaceCta = 95 - Len(Trim(pPla_cNombreCuentaD3))
                     SpaceMayCta = 95 - Len(Trim(pPla_cNombreCuentaD2))
                     
                      printl (Space(3) & "------------------------------------------------------------------------------------------------------------------------------------")
                      printl (Space(3) & Trim(pD3) & Space(6) & Trim(pPla_cNombreCuentaD3) & Space(SpaceCta) & pSaldoDebeD3 & Space(4) & pSaldoHaberD3)
                      printl (Space(3) & "------------------------------------------------------------------------------------------------------------------------------------")
                      printl (Space(3) & Space(48) & "SUMAS DEL MAYOR" & Space(41) & pSumMayDebeDet & Space(4) & pSumMayHaberDet)
                      VarTotal = CDbl(VarTotal) + CDbl(pSumMayDebeDet)
                      VartotalH = CDbl(VartotalH) + CDbl(pSumMayHaberDet)
                      printl (Space(3) & Trim(pD2) & Space(7) & Trim(pPla_cNombreCuentaD2) & Space(SpaceMayCta) & pSaldoDebeD2 & Space(4) & pSaldoHaberD2)
                      printl (Space(3) & "------------------------------------------------------------------------------------------------------------------------------------")
                      'Vartotal = Vartotal + pSumMayDebeDet
                    Exit Do
                  End If
                  pPer_cPeriodo = Trim(!Per_cPeriodo)
                  
                  .MovePrevious
                 End If
'                 Debug.Print "LINEA 13"
                 pPla_cNombreCuenta = Trim(!Pla_cNombreCuenta)
                 pD3 = Trim(!d3)
                 If Trim(IsNull(!Pla_cNombreCuentaD3)) Then
                        pPla_cNombreCuentaD3 = ""
                Else
                    pPla_cNombreCuentaD3 = !Pla_cNombreCuentaD3
                End If
                 pPla_cNombreCuentaD2 = Trim(!Pla_cNombreCuentaD2)
                 ImprimeVanVienen
                 'Debug.Print "LINEA 14"
            Else
                 Cont = 0
                 ImprimeVanVienen
                 'Debug.Print "LINEA 15"
                 If CDbl(pSumDebeDet) <> 0 Or CDbl(pSumHaberDet) <> 0 Then printl (Space(3) & Space(104) & "------------    ------------")
                  
                  If Len(NombreMes(Trim(pPer_cPeriodo))) > Len("ENERO") Then
                   SpaceMes = 46 - Len(NombreMes(Trim(pPer_cPeriodo)))
                  ElseIf Len(NombreMes(Trim(pPer_cPeriodo))) < Len("ENERO") Then
                   SpaceMes = 43 - (Len("ENERO") - Len(NombreMes(Trim(pPer_cPeriodo))))
                  ElseIf Len(NombreMes(Trim(pPer_cPeriodo))) = Len("ENERO") Then
                   SpaceMes = 41
                  End If
                  
                  RSet pSumDebeDet = Format(pSumDebeDet, "#,###,###,##0.00;(#,###,###,##0.00)")
                  RSet pSumHaberDet = Format(pSumHaberDet, "#,###,###,##0.00;(#,###,###,##0.00)")
                  'Debug.Print "LINEA 16"
                  RSet pSumMayDebeDet = Format(pSumMayDebeDet, "#,###,###,##0.00;(#,###,###,##0.00)")
                  RSet pSumMayHaberDet = Format(pSumMayHaberDet, "#,###,###,##0.00;(#,###,###,##0.00)")
                  
                  RSet pSaldoDebeD2 = Format(pSaldoDebeD2, "#,###,###,##0.00;(#,###,###,##0.00)")
                  RSet pSaldoHaberD2 = Format(pSaldoHaberD2, "#,###,###,##0.00;(#,###,###,##0.00)")
                                
                  pDebeFin = CDbl(pSumSldoFinDebeDet) - CDbl(pSumSldoFinHaberDet)
                  pDebeFin = CDbl(pDebeFin) * -1
                  If CDbl(pDebeFin) < 0 Then RSet pDebeFin = Format(Abs(pDebeFin), "#,###,###,##0.00;(#,###,###,##0.00)") Else RSet pDebeFin = Format(0, "#,###,###,##0.00;(#,###,###,##0.00)")
                  
                  pHaberFin = CDbl(pSumSldoFinDebeDet) - CDbl(pSumSldoFinHaberDet)
                  pHaberFin = CDbl(pHaberFin) * -1
                  If CDbl(pHaberFin) > 0 Then RSet pHaberFin = Format(Abs(pHaberFin), "#,###,###,##0.00;(#,###,###,##0.00)") Else RSet pHaberFin = Format(0, "#,###,###,##0.00;(#,###,###,##0.00)")
                  
                  ImprimeVanVienen
                  If CDbl(pSumDebeDet) <> 0 Or CDbl(pSumHaberDet) <> 0 Then
                    printl (Space(3) & Space(48) & "PERIODO : " & NombreMes(Trim(pPer_cPeriodo)) & Space(SpaceMes) & pSumDebeDet & Space(4) & pSumHaberDet)
                    
                    
                    sDebe = CDbl(sDebe) + CDbl(pSumDebeDet)
                    sHaber = CDbl(sHaber) + CDbl(pSumHaberDet)
                  End If
                                
                     pSumasMesesDebe = CDbl(pSumasMesesDebe) + CDbl(pSumDebeDet): RSet pSumasMesesDebe = Format(pSumasMesesDebe, "#,###,###,##0.00;(#,###,###,##0.00)")
                     pSumasMesesHaber = CDbl(pSumasMesesHaber) + CDbl(pSumHaberDet): RSet pSumasMesesHaber = Format(pSumasMesesHaber, "#,###,###,##0.00;(#,###,###,##0.00)")
                                              
                  SpaceSumasMeses = 47 - Len(NombreMes(Trim(gsLdMesIni)) & " A " & NombreMes(Trim(gsLdMesFin)))
                                
                  ImprimeVanVienen
                  If Trim(gsLdMesIni) <> Trim(gsLdMesFin) And Trim(pAse_nVoucher) <> "" Then
                    printl (Space(3) & Space(48) & "SUMAS DE " & NombreMes(Trim(gsLdMesIni)) & " A " & NombreMes(Trim(gsLdMesFin)) & Space(SpaceSumasMeses) & pSumasMesesDebe & Space(4) & pSumasMesesHaber)
                  End If
                  
                  If Trim(gsLdMesIni) = Trim(gsLdMesFin) And Len(NombreMes(Trim(pPer_cPeriodo))) = 0 Then pPer_cPeriodo = Trim(gsLdMesIni)
                  SpaceSldo = 38 - Len(NombreMes(Trim(pPer_cPeriodo)))
                  
                  ImprimeVanVienen
                  'printl (pPla_cCuentaContable & Space(1) & pPla_cNombreCuenta & Space(23) & "SALDO FINAL A" & Space(43) & pDebeFin & Space(1) & pHaberFin)
                  If Trim(gsLdMesIni) = Trim(gsLdMesFin) And Len(NombreMes(Trim(pPer_cPeriodo))) = 0 Then pPer_cPeriodo = Trim(gsLdMesIni)
                  printl (Space(3) & pPla_cCuentaContable & Space(1) & pPla_cNombreCuenta & Space(23) & "SALDO FINAL A " & NombreMes(Trim(pPer_cPeriodo)) & Space(SpaceSldo) & pDebeFin & Space(4) & pHaberFin)
                  
                  SpaceCta = 95 - Len(Trim(pPla_cNombreCuentaD3))
                  SpaceMayCta = 95 - Len(Trim(pPla_cNombreCuentaD2))
                  
                  ImprimeVanVienen

                  If Trim(!d3) <> Trim(pD3) Then
                   ImprimeVanVienen
                   printl (Space(3) & "------------------------------------------------------------------------------------------------------------------------------------")
                   ImprimeVanVienen
                   printl (Space(3) & Trim(pD3) & Space(6) & Trim(pPla_cNombreCuentaD3) & Space(SpaceCta) & pSaldoDebeD3 & Space(4) & pSaldoHaberD3)
                   ImprimeVanVienen
                   printl (Space(3) & "------------------------------------------------------------------------------------------------------------------------------------")
                  End If
                  
                  ImprimeVanVienen
                  If Trim(!d2) <> Trim(pD2) Then
                   ImprimeVanVienen
                   VarTotal = VarTotal + pSumMayDebeDet
                   VartotalH = VartotalH + pSumMayHaberDet
                   printl (Space(3) & Space(48) & "SUMAS DEL MAYOR" & Space(41) & pSumMayDebeDet & Space(4) & pSumMayHaberDet)
                   ImprimeVanVienen
                   
                   printl (Space(3) & Trim(pD2) & Space(7) & Trim(pPla_cNombreCuentaD2) & Space(SpaceMayCta) & pSaldoDebeD2 & Space(4) & pSaldoHaberD2)
                   ImprimeVanVienen
                   printl (Space(3) & "------------------------------------------------------------------------------------------------------------------------------------")
                  End If
                  
                  pD3Ant = pD3
                  pD2Ant = pD2
                  
                 Exit Do
                End If
             .MoveNext
             'Debug.Print "LINEA 17"
             
             If .EOF Then
                If CDbl(pSumDebeDet) <> 0 Or CDbl(pSumHaberDet) <> 0 Then printl (Space(3) & Space(104) & "------------    ------------")
                
                 If Len(NombreMes(Trim(pPer_cPeriodo))) > Len("ENERO") Then
                  SpaceMes = 46 - Len(NombreMes(Trim(pPer_cPeriodo)))
                 ElseIf Len(NombreMes(Trim(pPer_cPeriodo))) < Len("ENERO") Then
                  SpaceMes = 43 - (Len("ENERO") - Len(NombreMes(Trim(pPer_cPeriodo))))
                 ElseIf Len(NombreMes(Trim(pPer_cPeriodo))) = Len("ENERO") Then
                  SpaceMes = 41
                 End If
                 
                 RSet pSumDebeDet = Format(pSumDebeDet, "#,###,###,##0.00;(#,###,###,##0.00)")
                 RSet pSumHaberDet = Format(pSumHaberDet, "#,###,###,##0.00;(#,###,###,##0.00)")
                 
                 RSet pSumMayDebeDet = Format(pSumMayDebeDet, "#,###,###,##0.00;(#,###,###,##0.00)")
                 RSet pSumMayHaberDet = Format(pSumMayHaberDet, "#,###,###,##0.00;(#,###,###,##0.00)")
                                  
                 pDebeFin = CDbl(pSumSldoFinDebeDet) - CDbl(pSumSldoFinHaberDet)
                 pDebeFin = CDbl(pDebeFin) * -1
                 If CDbl(pDebeFin) < 0 Then RSet pDebeFin = Format(Abs(pDebeFin), "#,###,###,##0.00;(#,###,###,##0.00)") Else RSet pDebeFin = Format(0, "#,###,###,##0.00;(#,###,###,##0.00)")
                 
                 pHaberFin = CDbl(pSumSldoFinDebeDet) - CDbl(pSumSldoFinHaberDet)
                 pHaberFin = CDbl(pHaberFin) * -1
                 If CDbl(pHaberFin) > 0 Then RSet pHaberFin = Format(Abs(pHaberFin), "#,###,###,##0.00;(#,###,###,##0.00)") Else RSet pHaberFin = Format(0, "#,###,###,##0.00;(#,###,###,##0.00)")
                 
                 If CDbl(pSumDebeDet) <> 0 Or CDbl(pSumHaberDet) <> 0 Then
                    printl (Space(3) & Space(48) & "PERIODO : " & NombreMes(Trim(pPer_cPeriodo)) & Space(SpaceMes) & pSumDebeDet & Space(4) & pSumHaberDet)
                    sDebe = CDbl(sDebe) + CDbl(pSumDebeDet)
                    sHaber = CDbl(sHaber) + CDbl(pSumHaberDet)
                 End If

                 pSumasMesesDebe = CDbl(pSumasMesesDebe) + CDbl(pSumDebeDet): RSet pSumasMesesDebe = Format(pSumasMesesDebe, "#,###,###,##0.00;(#,###,###,##0.00)")
                 pSumasMesesHaber = CDbl(pSumasMesesHaber) + CDbl(pSumHaberDet): RSet pSumasMesesHaber = Format(pSumasMesesHaber, "#,###,###,##0.00;(#,###,###,##0.00)")
                 
                 SpaceSumasMeses = 47 - Len(NombreMes(Trim(gsLdMesIni)) & " A " & NombreMes(Trim(gsLdMesFin)))
                 If Trim(gsLdMesIni) <> Trim(gsLdMesFin) And Trim(pAse_nVoucher) <> "" Then
                  printl (Space(3) & Space(48) & "SUMAS DE " & NombreMes(Trim(gsLdMesIni)) & " A " & NombreMes(Trim(gsLdMesFin)) & Space(SpaceSumasMeses) & pSumasMesesDebe & Space(4) & pSumasMesesHaber)
                 End If
                 
                 If Trim(gsLdMesIni) = Trim(gsLdMesFin) And Len(NombreMes(Trim(pPer_cPeriodo))) = 0 Then pPer_cPeriodo = Trim(gsLdMesIni)
                 SpaceSldo = 38 - Len(NombreMes(Trim(pPer_cPeriodo)))
              
                 'printl (pPla_cCuentaContable & Space(1) & pPla_cNombreCuenta & Space(23) & "SALDO FINAL A" & Space(43) & pDebeFin & Space(1) & pHaberFin)
                 If Trim(gsLdMesIni) = Trim(gsLdMesFin) And Len(NombreMes(Trim(pPer_cPeriodo))) = 0 Then pPer_cPeriodo = Trim(gsLdMesIni)
                 'Debug.Print pPla_cCuentaContable & Space(1) & pPla_cNombreCuenta & Space(23) & "SALDO FINAL A " & NombreMes(Trim(pPer_cPeriodo)) & Space(SpaceSldo) & pDebeFin & Space(1) & pHaberFin
                 ImprimeVanVienen
                 printl (Space(3) & pPla_cCuentaContable & Space(1) & pPla_cNombreCuenta & Space(23) & "SALDO FINAL A " & NombreMes(Trim(pPer_cPeriodo)) & Space(SpaceSldo) & pDebeFin & Space(4) & pHaberFin)
                 SpaceCta = 95 - Len(Trim(pPla_cNombreCuentaD3))
                 SpaceMayCta = 95 - Len(Trim(pPla_cNombreCuentaD2))
                 
                 If Err.Number <> 0 Then MsgBox "ERROR: GG"
                 
                  printl (Space(3) & "------------------------------------------------------------------------------------------------------------------------------------")
                  printl (Space(3) & Trim(pD3) & Space(6) & Trim(pPla_cNombreCuentaD3) & Space(SpaceCta) & pSaldoDebeD3 & Space(4) & pSaldoHaberD3)
                  printl (Space(3) & "------------------------------------------------------------------------------------------------------------------------------------")
                  printl (Space(3) & Space(48) & "SUMAS DEL MAYOR" & Space(41) & pSumMayDebeDet & Space(4) & pSumMayHaberDet)
                  VarTotal = VarTotal + pSumMayDebeDet
                  VartotalH = VartotalH + pSumMayHaberDet
                  printl (Space(3) & Trim(pD2) & Space(7) & Trim(pPla_cNombreCuentaD2) & Space(SpaceMayCta) & pSaldoDebeD2 & Space(4) & pSaldoHaberD2)
                  printl (Space(3) & "------------------------------------------------------------------------------------------------------------------------------------")
                  'Debug.Print (Space(48) & "SUMAS DEL MAYOR" & Space(41) & pSumMayDebeDet & Space(1) & pSumMayHaberDet)
             End If
              If Err.Number <> 0 Then MsgBox "vd"
          Loop
'          .MoveNext
       Loop
    End If
 End With
 
'MsgBox Vartotal
Dim VarTempD As Double
Dim VarTempH As Double
VarTempD = pVanDebe
VarTempH = pVanHaber

pVanDebe = Format(VarTotal, "###,###,##0.00;(#,###,###,##0.00)")
pVanHaber = Format(VartotalH, "###,###,##0.00;(#,###,###,##0.00)")
giLineas = 1000
ImprimeVanVienen

pVanDebe = Format(VarTempD, "###,###,##0.00;(#,###,###,##0.00)")
pVanHaber = Format(VarTempH, "###,###,##0.00;(#,###,###,##0.00)")
giLineas = -6
Call ImprimeVanVienen

If frmFCImpresion.List_Destino.Text = "Archivo" Then
    Print #1, Chr(27) & Chr(18)
   Close #1
   frmFCVistaInforme.Caption = "Detalle de los Movimientos en Efectivo"
   frmFCVistaInforme.txtInforme.filename = frmFCImpresion.OutputFileName
   frmFCVistaInforme.Show
Else
   giLineas = 0
   Printer.FontName = "Draft 17cpi"
   Printer.FontSize = 10
   Printer.EndDoc
End If

Screen.MousePointer = vbNormal
DoEvents
'Exit Sub
'
'ERROR:
'MsgBox Err.Description, vbCritical, App.Title
'Resume
End Sub

Private Function ExistenDatos() As Boolean
On Error GoTo Error_cmd

Dim sSql As String

    Dim sTipo_Lib_Mayor As String
    
    If Me.OptForma(0).Value Then
        sTipo_Lib_Mayor = "0"
    ElseIf Me.OptForma(1).Value Then
        sTipo_Lib_Mayor = "1"
    End If

 Set rsArreglo = New ADODB.Recordset
 sSql = "spCn_RptFormato0601 'TODOS','" & gsEmpresa & "','" & gsAnio & "','" & gsLdMesIni & "','" & gsLdMesFin & "','" & gsCodMoneda & "','" & _
         gsCtaIni & "','" & gsCtaFin & "','" & sTipo_Lib_Mayor & "'"
' sSql = "spCn_RptFormato0601 'TODOS','" & gsEmpresa & "','" & gsAnio & "','" & gsLdMesIni & "','" & gsLdMesFin & "','" & gsCodMoneda & "','" & _
'         gsCtaIni & "','" & gsCtaFin & "', NULL" '& IIf(OptForma(0).Value, "0", "1") & "'"
 ConectarAdvance
  rsArreglo.Open sSql, gcnSistemaAdv, adOpenDynamic, adLockOptimistic
 Desconectar

Screen.MousePointer = vbNormal
ExistenDatos = IIf(rsArreglo.RecordCount > 0, True, False)
If Not rsArreglo.EOF Then rsArreglo.MoveFirst

Exit Function

Error_cmd:
    Screen.MousePointer = vbNormal
    ExistenDatos = False
    Desconectar
    MsgBox Err.Description, vbInformation, App.Title
End Function
Public Sub CabeceraLibMayorGeneral()
 EntroCabecera = True

 Dim sPag As String
 Dim Anio As String
 Dim Mes As String
 Dim sUSUARIO As String * 10

 Gs_HoraServ = DevuelveHoraServidor
 LSet sUSUARIO = gsUsuario

 If Gs_TamPapel = 39 Then nAncho = 232 Else nAncho = 142
 sPag = Space(4)
    
  gsConTotalPaginas = gsConTotalPaginas + 1
 gsPagina = gsPagina + 1

 RSet sPag = Format(gsPagina + 1, "####")
 giLineas = 0

 'Call AlinearDosTextos(nAncho - 6, Space(3) & "Formato 6.1: LIBRO MAYOR", "Fecha :  " & Format(FechaServidor, "dd/MM/yyyy"))
 printl (Space(3) & "Formato 6.1: LIBRO MAYOR                                                                                         Fecha :  " & Format(FechaServidor, "dd/MM/yyyy"))
 
 If gsLdMesIni = gsLdMesFin Then Mes = NombreMes(gsLdMesIni) & " " & gsAnio Else Mes = NombreMes(gsLdMesIni) & " A " & NombreMes(gsLdMesFin) & " " & gsAnio
 
 Call AlinearDosTextos(nAncho - 4, Space(3) & "EJERCICIO/PERIODO    : " & Mes, "")
 Dim xgsPagina As String * 4
 RSet xgsPagina = Format(CStr(gsPagina), "####")
 
 'Call AlinearDosTextos(nAncho - 13, Space(3) & "RUC                  : " & gsRUC, "Pagina: " & xgsPagina)
 printl (Space(3) & "RUC                  : " & gsRUC & "                                                                               Pagina: " & xgsPagina)
 
 'Call AlinearDosTextos(nAncho - 10, "RAZON SOCIAL         : " & gsEmpresaNom, "Página:       " & Format$(gsPaginaPrincipal, "####") & " de " & Format$(gsPagina, "####"))

Call AlinearDosTextos(nAncho - 13, Space(3) & "APELLIDOS Y NOMBRES,", "")
Call AlinearDosTextos(nAncho - 13, Space(3) & "DENOMINACIÓN O", "")
Call AlinearDosTextos(nAncho - 13, Space(3) & "RAZON SOCIAL         : " & gsEmpresaNom, "")
 
 Call AlinearDosTextos(nAncho, Space(3) & "MONEDA               : " & IIf(gsCodMoneda = gsMonedaNac, gsNombreMonedaNac, gsNombreMonedaExt), "")
 
 printl ("")
 printl (Space(3) & "------------------------------------------------------------------------------------------------------------------------------------")
 printl (Space(3) & "     CUENTA CONTABLE            FECHA      NUMERO   DESCRIPCION O GLOSA DE LA OPERACION    DOCUMENTO         SALDOS Y MOVIMIENTOS")
 printl (Space(3) & "CODIGO       DENOMINACION       OPER.      CORREL                                       TD SER NUMERO       DEUDOR          ACREEDOR")
 printl (Space(3) & "------------------------------------------------------------------------------------------------------------------------------------")
 Exit Sub
Control:
 MsgBox Err.Description
End Sub
Sub ImprimeVanVienen(Optional VarInd As Byte = 0, Optional Par_Periodo As String = "", Optional Par_Pla_cCuentaContable As String = "")
        On Error GoTo Control
    
    If VarInd = 1 Then GoTo Pintar
    
    Select Case giLineas
    Case 2000
Pintar:
    Case 1000
        printl ""
'        printl (Space(56) & "TOTAL MOVIMIENTOS DEL PERIODO " & Space(18) & pVanDebe & Space(1) & pVanHaber)
        RSet sDebe = Format(sDebe, "#,###,###,##0.00;(#,###,###,##0.00)")
        RSet sHaber = Format(sHaber, "#,###,###,##0.00;(#,###,###,##0.00)")
        printl (Space(3) & Space(54) & "TOTAL MOVIMIENTOS DEL PERIODO " & Space(19) & sDebe & Space(3) & sHaber)
        giLineas = 0
    Case -6
        printl ""
        Call Acumulado
'        printl (Space(56) & "TOTAL GENERAL ACUMULADO " & Space(23) & pVanDebe & Space(1) & pVanHaber)
        printl (Space(3) & Space(54) & "TOTAL GENERAL ACUMULADO " & Space(25) & sDebeBC & Space(3) & sHaberBC)
        giLineas = 0
    End Select
    
    Exit Sub
Control:
    MsgBox Err.Description
'    Resume
End Sub

Function Acumulado()
    On Error GoTo MIERROR
    Dim arrDatos() As Variant
    Dim AdoRsBC As ADODB.Recordset
    Dim sSql As String
    Dim clDatos As clsMantoTablas
    
    Set AdoRsBC = New ADODB.Recordset
    
    sSql = "spCn_RptBalanceComprobacion '" & gsEmpresa & "','" & gsAnio & "','" & _
    tdbcAlMes.BoundText & "','038','ACUMULADO',2"
    
    ConectarAdvance
    AdoRsBC.Open sSql, gcnSistemaAdv, adOpenDynamic, adLockOptimistic
    Desconectar
     
    sDebeBC = 0
    sHaberBC = 0
    
    If AdoRsBC.RecordCount > 0 Then
    
        Do While Not AdoRsBC.EOF
        
            sDebeBC = CDbl(sDebeBC) + CDbl(AdoRsBC("Sal_nMontoDebeS"))
            sHaberBC = CDbl(sHaberBC) + CDbl(AdoRsBC("Sal_nMontoHaberS"))
            
            AdoRsBC.MoveNext
        Loop
    
    End If
    
    RSet sDebeBC = Format(sDebeBC, "#,###,###,##0.00;(#,###,###,##0.00)")
    RSet sHaberBC = Format(sHaberBC, "#,###,###,##0.00;(#,###,###,##0.00)")
    Exit Function
MIERROR:
    
End Function
