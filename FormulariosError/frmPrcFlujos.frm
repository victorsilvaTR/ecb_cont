VERSION 5.00
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmPrcFlujos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Flujos de Efectivo"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7425
   Icon            =   "frmPrcFlujos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4995
   ScaleWidth      =   7425
   Begin VB.Frame fraTodo 
      Height          =   4920
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   7305
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3525
         Left            =   135
         TabIndex        =   1
         Top             =   225
         Width           =   7050
         Begin VB.TextBox tdbtAnio 
            Height          =   300
            Left            =   4500
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   2250
            Width           =   1545
         End
         Begin VB.CheckBox chkSaldos 
            Caption         =   "Utilizar saldos iniciales"
            Height          =   420
            Left            =   1125
            TabIndex        =   3
            Top             =   1305
            Width           =   1950
         End
         Begin VB.ComboBox cboClase 
            Height          =   315
            ItemData        =   "frmPrcFlujos.frx":0ECA
            Left            =   1125
            List            =   "frmPrcFlujos.frx":0ED4
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   3015
            Width           =   1050
         End
         Begin TrueOleDBList70.TDBCombo tdbcMes 
            Height          =   300
            Left            =   1110
            TabIndex        =   5
            Top             =   900
            Width           =   3300
            _ExtentX        =   5821
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=741"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=661"
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
            _PropDict       =   $"frmPrcFlujos.frx":0EDE
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
         Begin TrueOleDBList70.TDBCombo tdbcAnio 
            Height          =   300
            Left            =   4500
            TabIndex        =   6
            Top             =   900
            Width           =   1545
            _ExtentX        =   2725
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
            Columns.Count   =   2
            Splits(0)._UserFlags=   0
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).AllowRowSizing=   0   'False
            Splits(0).DividerStyle=   2
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2196"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2117"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=2196"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2117"
            Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(10)=   "Column(1).AllowSizing=0"
            Splits(0)._ColumnProps(11)=   "Column(1).Visible=0"
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
            _PropDict       =   $"frmPrcFlujos.frx":0F65
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
         Begin TrueOleDBList70.TDBCombo tdbcMesFinal 
            Height          =   300
            Left            =   1110
            TabIndex        =   7
            Top             =   2250
            Width           =   3300
            _ExtentX        =   5821
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=820"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=741"
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
            _PropDict       =   $"frmPrcFlujos.frx":0FEC
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
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Periodo Inicial"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   1110
            TabIndex        =   12
            Top             =   630
            Width           =   1200
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "FLUJOS DE EFECTIVO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1080
            TabIndex        =   11
            Top             =   270
            Width           =   5265
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Periodo Final"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   1110
            TabIndex        =   10
            Top             =   1935
            Width           =   1095
         End
         Begin MSForms.CommandButton cmdSaldos 
            Height          =   435
            Left            =   4500
            TabIndex        =   9
            ToolTipText     =   " Muestra los saldos iniciales ingresados manualmente "
            Top             =   1305
            Width           =   1530
            Caption         =   " Saldos"
            PicturePosition =   327683
            Size            =   "2699;767"
            Picture         =   "frmPrcFlujos.frx":1073
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Clase"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   1080
            TabIndex        =   8
            Top             =   2745
            Width           =   480
         End
      End
      Begin MSForms.CommandButton cmdSalir 
         Height          =   435
         Left            =   2805
         TabIndex        =   17
         Top             =   4365
         Width           =   1665
         Caption         =   " Salir"
         PicturePosition =   327683
         Size            =   "2937;767"
         Picture         =   "frmPrcFlujos.frx":160D
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdReporte 
         Height          =   435
         Left            =   5370
         TabIndex        =   16
         ToolTipText     =   " Impreime el flujo de efectivo "
         Top             =   3825
         Width           =   1665
         Caption         =   "Flujo Efectivo"
         PicturePosition =   327683
         Size            =   "2937;767"
         Picture         =   "frmPrcFlujos.frx":1BA7
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdGenerar 
         Height          =   435
         Left            =   240
         TabIndex        =   15
         ToolTipText     =   " Calcula el flujo de efectivo en el periodo seleccionado "
         Top             =   3825
         Width           =   1665
         Caption         =   " Metodo Directo"
         PicturePosition =   327683
         Size            =   "2937;767"
         Picture         =   "frmPrcFlujos.frx":2141
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdHojasTrabajo 
         Height          =   435
         Left            =   3660
         TabIndex        =   14
         ToolTipText     =   " Imprime las hojas de trabajo del flujo de efectivo "
         Top             =   3825
         Width           =   1665
         Caption         =   " Hojas Trabajo"
         PicturePosition =   327683
         Size            =   "2937;767"
         Picture         =   "frmPrcFlujos.frx":26DB
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdIndirecto 
         Height          =   435
         Left            =   1950
         TabIndex        =   13
         ToolTipText     =   " Calcula el flujo de efectivo en el periodo seleccionado "
         Top             =   3825
         Width           =   1665
         Caption         =   " Metodo Indirecto"
         PicturePosition =   327683
         Size            =   "2937;767"
         Picture         =   "frmPrcFlujos.frx":2C75
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
      TabIndex        =   18
      Top             =   0
      Width           =   4365
   End
End
Attribute VB_Name = "frmPrcFlujos"
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

Private Sub Procesar(Metodo As String)
    cmdHojasTrabajo.Enabled = False
    cmdGenerar.Enabled = False
    cmdsalir.Enabled = False
    cmdReporte.Enabled = False
    cmdIndirecto.Enabled = False
    DoEvents


    Screen.MousePointer = vbHourglass
    Dim clsMante As clsMantoTablas
    Dim lArrMnt(7) As Variant
    Dim Store As String
    Set clsMante = New clsMantoTablas

    lArrMnt(0) = gsEmpresa
    lArrMnt(1) = tdbtAnio.Text 'final
    lArrMnt(2) = tdbcAnio.BoundText 'inicial
    lArrMnt(3) = tdbcMesFinal.BoundText 'final
    lArrMnt(4) = tdbcMes.BoundText 'inicial
    lArrMnt(5) = gsUsuario
    lArrMnt(6) = IIf(chkSaldos.Value = vbChecked, "S", "N")
    lArrMnt(7) = Metodo
    
    
    If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaTemporalFlujo", lArrMnt(), True) = False Then
        Mensajes "El proceso no ha concluido. Verificar...", vbInformation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Set clsMante = Nothing
    
    Screen.MousePointer = vbDefault
    
    Mensajes "Proceso de calculo Terminado"
    
    cmdGenerar.Enabled = True
    cmdsalir.Enabled = True
    cmdReporte.Enabled = True
    cmdHojasTrabajo.Enabled = True
    cmdIndirecto.Enabled = True
End Sub

Private Sub cmdGenerar_Click()
    Call Procesar("DIRECTO")
End Sub

Private Sub cmdIndirecto_Click()
    Call Procesar("INDIRECTO")
End Sub


Private Sub cmdHojasTrabajo_Click()
    
    cmdHojasTrabajo.Enabled = False
    
    DoEvents
    Screen.MousePointer = vbHourglass
    
        
        Dim matriz_fecha(9) As Variant
        Dim Tipo As String
        
        matriz_fecha(0) = "@Tipo;HOJA;True"
        matriz_fecha(1) = "@EMPRESA;" & gsEmpresaNom & ";True"
        matriz_fecha(2) = "@RUC;" & gsRUC & ";True"
        matriz_fecha(3) = "@Emp_cCodigo;" & gsEmpresa & ";True"
        matriz_fecha(4) = "@Pan_cAnio;" & gsAnio & ";True"
        matriz_fecha(5) = "@Per_cPeriodo;" & tdbcMesFinal.BoundText & ";True"
        matriz_fecha(6) = "@Rep_cClase;" & cboClase.Text & ";True"
        matriz_fecha(7) = "@Usuario;" & gsUsuario & ";True"
        matriz_fecha(8) = "@FECHAINI;" & CE(UltimoDiaMes(tdbcMes.BoundText, tdbcAnio.BoundText)) & ";True"
        matriz_fecha(9) = "@FECHAFIN;" & CE(UltimoDiaMes(tdbcMesFinal.BoundText, tdbtAnio.Text)) & ";True"

        Dim formulas(0) As Variant
       
        AbreReporteParam gsDSN, Me, rutaReportes & "RptFlujoEfectivoHoja.rpt", crptToWindow, "Reporte de Flujos de Efectivo", "", matriz_fecha(), formulas()
    
    
    
    Screen.MousePointer = vbDefault
    
    cmdHojasTrabajo.Enabled = True
End Sub


Private Sub cmdReporte_Click()
    
    cmdReporte.Enabled = False
    
    Screen.MousePointer = vbHourglass
    
        
        Dim matriz_fecha(7) As Variant
        Dim Tipo As String
        
        matriz_fecha(0) = "@Tipo;FLUJO;True"
        matriz_fecha(1) = "@EMPRESA;" & gsEmpresaNom & ";True"
        matriz_fecha(2) = "@RUC;" & gsRUC & ";True"
        matriz_fecha(3) = "@Emp_cCodigo;" & gsEmpresa & ";True"
        matriz_fecha(4) = "@Pan_cAnio;" & gsAnio & ";True"
        matriz_fecha(5) = "@Per_cPeriodo;" & tdbcMesFinal.BoundText & ";True"
        matriz_fecha(6) = "@Rep_cClase;" & cboClase.Text & ";True"
        
        matriz_fecha(7) = "@Usuario;" & gsUsuario & ";True"
    
        Dim formulas(0) As Variant
        
        AbreReporteParam gsDSN, Me, rutaReportes & "RptFlujoEfectivo.rpt", crptToWindow, "Reporte de Flujos de Efectivo", "", matriz_fecha(), formulas()
        
    
    Screen.MousePointer = vbDefault
    
    cmdReporte.Enabled = True
End Sub

Private Sub cmdSaldos_Click()
    frmManFlujoSaldos.Grupo = BuscaArray("mnuFlujoEjectivo")
    frmManFlujoSaldos.Show
    
    frmManFlujoSaldos.tdbcAnio.BoundText = tdbcAnio.BoundText
    DoEvents
    frmManFlujoSaldos.tdbcMes.BoundText = tdbcMes.BoundText
    DoEvents
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub Form_Load()
    Call Centrar_form(Me)
    
    Call LlenaCombos
    Call LlenaComboMesApeAddItem(tdbcMes)
    Call LlenaComboMesApeAddItem(tdbcMesFinal)
    
    DoEvents
    
    If gsPeriodo = "00" Then
        tdbcMes.BoundText = "01"
    ElseIf gsPeriodo = "13" Or gsPeriodo = "14" Then
        tdbcMes.BoundText = "12"
    Else
        tdbcMes.BoundText = gsPeriodo
    End If
    
    tdbtAnio.Text = gsAnio
    tdbcMesFinal.BoundText = gsPeriodo
    
    tdbcMes.ReBind
    tdbcMesFinal.ReBind
    cboClase.ListIndex = 1
    
End Sub

Private Sub LlenaCombos()
    Dim i As Integer
    tdbcAnio.Clear
    
    For i = 2000 To NE(gsAnio) - 1
        tdbcAnio.AddItem CE(i) & ";" & CE(i)
    Next i
    

    tdbcAnio.Bookmark = 0
    tdbcAnio.ListField = "column1"
    tdbcAnio.BoundColumn = "column1"
    tdbcAnio.BoundText = gsAnio - 1
    tdbcAnio.ReBind
    
End Sub

Private Sub Form_Resize()
On Error GoTo errHand
    If Me.WindowState <> vbMinimized Then
        
        Call Centrar_Objeto(fraTodo, Me)
        Call CentrarTitulo(lblTitulo, fraTodo, Me)
    End If
Exit Sub
errHand:
End Sub

Private Sub tdbcMes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pSetFocus cmdReporte
End If
End Sub

'Private Sub tdbtClase_LostFocus()
'    If tdbtClase.Text <> "6" And tdbtClase.Text <> "9" Then
'       Mensajes "La clase debe ser 6 o 9"
'       tdbtClase.Text = ""
'       pSetFocus tdbtClase
'    End If
'End Sub


