VERSION 5.00
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmRepAnexoInvBalanceEmp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anexo al Libro de Inventarios y Balances"
   ClientHeight    =   2565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6285
   Icon            =   "frmRepAnexoInvBalanceEmp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   2565
   ScaleWidth      =   6285
   Begin VB.PictureBox picForm 
      BorderStyle     =   0  'None
      Height          =   2580
      Left            =   -90
      ScaleHeight     =   2580
      ScaleWidth      =   6360
      TabIndex        =   5
      Top             =   0
      Width           =   6360
      Begin VB.Frame fraTodo 
         Height          =   2460
         Left            =   135
         TabIndex        =   6
         Top             =   0
         Width           =   6120
         Begin TrueOleDBList70.TDBCombo tdbcMes 
            Height          =   300
            Left            =   2752
            TabIndex        =   0
            Top             =   405
            Width           =   1935
            _ExtentX        =   3413
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
            _PropDict       =   $"frmRepAnexoInvBalanceEmp.frx":0ECA
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
            Left            =   2752
            TabIndex        =   1
            Tag             =   "_"
            Top             =   900
            Width           =   1935
            _ExtentX        =   3413
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
            _PropDict       =   $"frmRepAnexoInvBalanceEmp.frx":0F51
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
         Begin TrueOleDBList70.TDBCombo tdbcMetodo 
            Height          =   300
            Left            =   2752
            TabIndex        =   2
            Top             =   1350
            Visible         =   0   'False
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   529
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            _DropdownWidth  =   4392
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=794"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=714"
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
            _PropDict       =   $"frmRepAnexoInvBalanceEmp.frx":0FD8
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
            Left            =   1672
            TabIndex        =   9
            Top             =   1350
            Visible         =   0   'False
            Width           =   765
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
            Left            =   1672
            TabIndex        =   8
            Top             =   450
            Width           =   375
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
            Left            =   1672
            TabIndex        =   7
            Top             =   945
            Width           =   765
         End
         Begin MSForms.CommandButton cmdImprimir 
            Height          =   435
            Left            =   1402
            TabIndex        =   3
            Top             =   1890
            Width           =   1665
            Caption         =   " Vista Previa"
            PicturePosition =   327683
            Size            =   "2937;767"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdSalir 
            Height          =   435
            Left            =   3217
            TabIndex        =   4
            Top             =   1890
            Width           =   1665
            Caption         =   " Salir"
            PicturePosition =   327683
            Size            =   "2937;767"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
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
Attribute VB_Name = "frmRepAnexoInvBalanceEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ReporteSunat As String
Public TituloSunat As String
Public CuentaInvBal As String

Dim gsGrupo As String
Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

Public Sub CerrarForm(cReporteSunat As String)
    If cReporteSunat = ReporteSunat Then
        Unload Me
    End If
End Sub

Private Sub cmdImprimir_Click()
    Dim matriz_fecha(25) As Variant
    Dim formulas(0) As Variant
    Dim cRpt As String
    Dim i As Integer
    Screen.MousePointer = vbHourglass
    cmdImprimir.Enabled = False
    
    For i = 0 To 25
        matriz_fecha(i) = ""
    Next i
        
    cRpt = ReporteSunat
    
    DoEvents
    
    If ReporteSunat = "PCGE10" Then
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
        
    
    ElseIf ReporteSunat = "PCGE31" Then
'        GoTo entrar
        matriz_fecha(0) = "@Tipo;REPORTE;True"
        matriz_fecha(1) = "@Emp_cCodigo;" & gsEmpresa & ";True"
        matriz_fecha(2) = "@Pan_cAnio;" & gsAnio & ";True"
        matriz_fecha(3) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
        matriz_fecha(4) = "@EMPRESA;" & gsEmpresaNom & ";True"
        matriz_fecha(5) = "@RUC;" & gsRUC & ";True"
        matriz_fecha(6) = "@Ase_nVoucher;;True"
        matriz_fecha(7) = "@Asd_nItem;0;True"
        matriz_fecha(8) = "@Ten_cTipoEntidad;;True"
        matriz_fecha(9) = "@Ent_cCodentidad;;True"
        matriz_fecha(10) = "@Val_cTitulo;;True"
        matriz_fecha(11) = "@Val_cDesTitulo;;True"
        matriz_fecha(12) = "@Val_nValorNom;0;True"
        matriz_fecha(13) = "@Val_nCantidad;0;True"
        matriz_fecha(14) = "@Val_nCostoTot;0;True"
        matriz_fecha(15) = "@Val_nProvTot;0;True"
        matriz_fecha(16) = "@Val_nTotalNeto;0;True"
        matriz_fecha(17) = "@NOMMONEDA;" & tdbcMoneda.Text & ";True"
        matriz_fecha(18) = "@TITULO;" & CuentaInvBal & " - " & BuscaNombreCuenta(CuentaInvBal, False) & ";True"
        matriz_fecha(19) = "@Cuenta;" & Right(ReporteSunat, 2) & ";True"
        
'    ElseIf ReporteSunat = "PCGE18" Then
'        matriz_fecha(0) = "@Tipo;TODOS;True"
'        matriz_fecha(1) = "@Emp_cCodigo;" & gsEmpresa & ";True"
'        matriz_fecha(2) = "@Pan_cAnio;" & gsAnio & ";True"
'        matriz_fecha(3) = "@Per_cPeriodoDesde;" & tdbcMes.BoundText & ";True"
'        matriz_fecha(4) = "@Per_cPeriodoHasta;" & tdbcMes.BoundText & ";True"
'        matriz_fecha(5) = "@moneda;" & tdbcMoneda.BoundText & ";True"
'        matriz_fecha(6) = "@CtaDesde;18;True"
'        matriz_fecha(7) = "@CtaHasta;18;True"
'        matriz_fecha(8) = "@EMPRESA;" & gsEmpresaNom & ";True"
'        matriz_fecha(9) = "@RUC;" & gsRUC & ";True"
'        matriz_fecha(10) = "@NOMMONEDA;" & tdbcMoneda.Text & ";True"
'        matriz_fecha(11) = "@TITULO;" & CuentaInvBal & " - " & BuscaNombreCuenta(CuentaInvBal, False) & ";True"
'
    ElseIf (ReporteSunat >= "PCGE12" And ReporteSunat <= "PCGE19") Then
        matriz_fecha(0) = "@Emp_cCodigo;" & gsEmpresa & ";True"
        matriz_fecha(1) = "@Pan_cAnio;" & gsAnio & ";True"
        matriz_fecha(2) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
        matriz_fecha(3) = "@moneda;" & tdbcMoneda.BoundText & ";True"
        matriz_fecha(4) = "@EMPRESA;" & gsEmpresaNom & ";True"
        matriz_fecha(5) = "@RUC;" & gsRUC & ";True"
        matriz_fecha(6) = "@Cta_BalInv;" & CuentaInvBal & ";True"
        matriz_fecha(7) = "@NOMMONEDA;" & tdbcMoneda.Text & ";True"
        matriz_fecha(8) = "@TITULO;" & CuentaInvBal & " - " & BuscaNombreCuenta(CuentaInvBal, False) & ";True"
        
    ElseIf (ReporteSunat >= "PCGE20" And ReporteSunat <= "PCGE29") Then
        matriz_fecha(0) = "@Accion;REPORTE;True"
        matriz_fecha(1) = "@Emp_cCodigo;" & gsEmpresa & ";True"
        matriz_fecha(2) = "@Pan_cAnio;" & gsAnio & ";True"
        matriz_fecha(3) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
        matriz_fecha(4) = "@Mer_cMetodo;" & tdbcMetodo.BoundText & ";True"
        matriz_fecha(5) = "@Pla_cCuentaContable;" & CuentaInvBal & ";True"
        matriz_fecha(6) = "@EMPRESA;" & gsEmpresaNom & ";True"
        matriz_fecha(7) = "@RUC;" & gsRUC & ";True"
        matriz_fecha(8) = "@NOMMONEDA;" & tdbcMoneda.Text & ";True"
        matriz_fecha(9) = "@TITULO;" & CuentaInvBal & " - " & BuscaNombreCuenta(CuentaInvBal, False) & ";True"
        
        cRpt = "PCGE2X"
    ElseIf ReporteSunat = "PCGE42" Then
        matriz_fecha(0) = "@Emp_cCodigo;" & gsEmpresa & ";True"
        matriz_fecha(1) = "@Pan_cAnio;" & gsAnio & ";True"
        matriz_fecha(2) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
        matriz_fecha(3) = "@moneda;" & tdbcMoneda.BoundText & ";True"
        matriz_fecha(4) = "@EMPRESA;" & gsEmpresaNom & ";True"
        matriz_fecha(5) = "@RUC;" & gsRUC & ";True"
        matriz_fecha(6) = "@Cta_BalInv;" & CuentaInvBal & ";True"
        matriz_fecha(7) = "@NOMMONEDA;" & tdbcMoneda.Text & ";True"
        matriz_fecha(8) = "@TITULO;" & CuentaInvBal & " - " & BuscaNombreCuenta(CuentaInvBal, False) & ";True"
    ElseIf ReporteSunat = "PCGE43" Then
        matriz_fecha(0) = "@Emp_cCodigo;" & gsEmpresa & ";True"
        matriz_fecha(1) = "@Pan_cAnio;" & gsAnio & ";True"
        matriz_fecha(2) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
        matriz_fecha(3) = "@moneda;" & tdbcMoneda.BoundText & ";True"
        matriz_fecha(4) = "@EMPRESA;" & gsEmpresaNom & ";True"
        matriz_fecha(5) = "@RUC;" & gsRUC & ";True"
        matriz_fecha(6) = "@Cta_BalInv;" & CuentaInvBal & ";True"
        matriz_fecha(7) = "@NOMMONEDA;" & tdbcMoneda.Text & ";True"
        matriz_fecha(8) = "@TITULO;" & CuentaInvBal & " - " & BuscaNombreCuenta(CuentaInvBal, False) & ";True"
    ElseIf ReporteSunat = "PCGE40" Then
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
    ElseIf ReporteSunat = "PCGE41" Or ReporteSunat = "PCGE44" Or _
    ReporteSunat = "PCGE45" Or ReporteSunat = "PCGE46" Or _
    ReporteSunat = "PCGE47" Or ReporteSunat = "PCGE48" Then
        matriz_fecha(0) = "@Emp_cCodigo;" & gsEmpresa & ";True"
        matriz_fecha(1) = "@Pan_cAnio;" & gsAnio & ";True"
        matriz_fecha(2) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
        matriz_fecha(3) = "@moneda;" & tdbcMoneda.BoundText & ";True"
        matriz_fecha(4) = "@EMPRESA;" & gsEmpresaNom & ";True"
        matriz_fecha(5) = "@RUC;" & gsRUC & ";True"
        matriz_fecha(6) = "@Cta_BalInv;" & CuentaInvBal & ";True"
        matriz_fecha(7) = "@NOMMONEDA;" & tdbcMoneda.Text & ";True"
        matriz_fecha(8) = "@TITULO;" & CuentaInvBal & " - " & BuscaNombreCuenta(CuentaInvBal, False) & ";True"
    ElseIf ReporteSunat = "PCGE49" Then
        matriz_fecha(0) = "@Emp_cCodigo;" & gsEmpresa & ";True"
        matriz_fecha(1) = "@Pan_cAnio;" & gsAnio & ";True"
        matriz_fecha(2) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
        matriz_fecha(3) = "@CuentaInicio;49;True"
        matriz_fecha(4) = "@CuentaFin;49999999;True"
        matriz_fecha(5) = "@Moneda;" & tdbcMoneda.BoundText & ";True"
        matriz_fecha(6) = "@EMPRESA;" & gsEmpresaNom & ";True"
        matriz_fecha(7) = "@RUC;" & gsRUC & ";True"
        matriz_fecha(8) = "@NOMMONEDA;" & tdbcMoneda.Text & ";True"
        matriz_fecha(9) = "@TITULO;" & CuentaInvBal & " - " & BuscaNombreCuenta(CuentaInvBal, False) & ";True"
'        AbreReporteParam gsDSN, Me, rutaReportes & "RptFormato_0310.rpt", crptToWindow, "Anexo al Libro de Inventarios y Balances", "", matriz_fecha(), formulas()
    ElseIf ReporteSunat = "PCGE50_0" Then
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
        
    ElseIf ReporteSunat = "PCGE50_1" Or ReporteSunat = "PCGE51" Or _
    ReporteSunat = "PCGE52" Or ReporteSunat = "PCGE56" Or _
    ReporteSunat = "PCGE57" Or ReporteSunat = "PCGE58" Then
        matriz_fecha(0) = "@Emp_cCodigo;" & gsEmpresa & ";True"
        matriz_fecha(1) = "@Pan_cAnio;" & gsAnio & ";True"
        matriz_fecha(2) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
        matriz_fecha(3) = "@moneda;" & tdbcMoneda.BoundText & ";True"
        matriz_fecha(4) = "@EMPRESA;" & gsEmpresaNom & ";True"
        matriz_fecha(5) = "@RUC;" & gsRUC & ";True"
        matriz_fecha(6) = "@Cta_BalInv;" & CuentaInvBal & ";True"
        matriz_fecha(7) = "@NOMMONEDA;" & tdbcMoneda.Text & ";True"
        matriz_fecha(8) = "@TITULO;" & CuentaInvBal & " - " & BuscaNombreCuenta(CuentaInvBal, False) & ";True"
'        matriz_fecha(0) = "@Tipo;FUN;True"
'        matriz_fecha(1) = "@Emp_cCodigo;" & gsEmpresa & ";True"
'        matriz_fecha(2) = "@Pan_cAnio;" & gsAnio & ";True"
'        matriz_fecha(3) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
'        matriz_fecha(4) = "@moneda;" & tdbcMoneda.BoundText & ";True"
'        matriz_fecha(5) = "@EMPRESA;" & gsEmpresaNom & ";True"
'        matriz_fecha(6) = "@RUC;" & gsRUC & ";True"
'        matriz_fecha(7) = "@Reporte;;True"
'        matriz_fecha(8) = "@NOMMONEDA;" & tdbcMoneda.Text & ";True"
'        matriz_fecha(9) = "@UltimoDia;" & UltimoDiaMes(tdbcMes.BoundText, gsAnio) & ";True"
'        matriz_fecha(10) = "@TITULO;" & CuentaInvBal & " - " & BuscaNombreCuenta(CuentaInvBal, False) & ";True"
''        matriz_fecha(11) = "@Mes;" & IIf(ChkVerMes.Value = 1, 1, 0) & ";True"
'
''        If chkAnexos.Value = "1" Then
''            matriz_fecha(7) = "@Reporte;DETALLE;True"
''            AbreReporteParam gsDSN, Me, rutaReportes & "RptBalanceGeneralConasevDetalle.rpt", crptToWindow, "Anexo al Libro de Inventarios y Balances", "", matriz_fecha(), formulas()
''        Else
'            matriz_fecha(7) = "@Reporte;RESUMEN;True"
''            AbreReporteParam gsDSN, Me, rutaReportes & "RptFormato_0320.rpt", crptToWindow, "Anexo al Libro de Inventarios y Balances", "", matriz_fecha(), formulas()
''        End If
    
    ElseIf ReporteSunat = "PCGE59" Then
        matriz_fecha(0) = "@Emp_cCodigo;" & gsEmpresa & ";True"
        matriz_fecha(1) = "@Pan_cAnio;" & gsAnio & ";True"
        matriz_fecha(2) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
        matriz_fecha(3) = "@moneda;" & tdbcMoneda.BoundText & ";True"
        matriz_fecha(4) = "@EMPRESA;" & gsEmpresaNom & ";True"
        matriz_fecha(5) = "@RUC;" & gsRUC & ";True"
        matriz_fecha(6) = "@Cta_BalInv;" & CuentaInvBal & ";True"
        matriz_fecha(7) = "@NOMMONEDA;" & tdbcMoneda.Text & ";True"
        matriz_fecha(8) = "@TITULO;" & CuentaInvBal & " - " & BuscaNombreCuenta(CuentaInvBal, False) & ";True"
   ElseIf ReporteSunat = "PCGE33" Or ReporteSunat = "PCGE35" Or _
   ReporteSunat = "PCGE36" Or ReporteSunat = "PCGE37" Or ReporteSunat = "PCGE38" Then
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
        
'    ElseIf ReporteSunat = "PCGE35" Then
'        matriz_fecha(0) = "@EMPRESA;" & gsEmpresaNom & ";True"
'        matriz_fecha(1) = "@RUC;" & gsRUC & ";True"
'        matriz_fecha(2) = "@NOMMONEDA;" & tdbcMoneda.Text & ";True"
'        matriz_fecha(3) = "@TITULO;" & CuentaInvBal & " - " & BuscaNombreCuenta(CuentaInvBal, False) & ";True"
''        AbreReporteParam gsDSN, Me, rutaReportes & "RptFormato_" & cRpt & ".rpt", crptToWindow, "Saldos de Cuentas de Balance", "", matriz_fecha(), formulas()
''        Exit Sub
'    ElseIf ReporteSunat = "PCGE36" Then
'        matriz_fecha(0) = "@EMPRESA;" & gsEmpresaNom & ";True"
'        matriz_fecha(1) = "@RUC;" & gsRUC & ";True"
'        matriz_fecha(2) = "@NOMMONEDA;" & tdbcMoneda.Text & ";True"
'        matriz_fecha(3) = "@TITULO;" & CuentaInvBal & " - " & BuscaNombreCuenta(CuentaInvBal, False) & ";True"
''        AbreReporteParam gsDSN, Me, rutaReportes & "RptFormato_" & cRpt & ".rpt", crptToWindow, "Saldos de Cuentas de Balance", "", matriz_fecha(), formulas()
''        Exit Sub
    ElseIf ReporteSunat = "PCGE34" Then
        matriz_fecha(0) = "@Accion;REPORTE;True"
        matriz_fecha(1) = "@Emp_cCodigo;" & gsEmpresa & ";True"
        matriz_fecha(2) = "@Pan_cAnio;" & gsAnio & ";True"
        matriz_fecha(3) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
        matriz_fecha(4) = "@EMPRESA;" & gsEmpresaNom & ";True"
        matriz_fecha(5) = "@RUC;" & gsRUC & ";True"
        matriz_fecha(6) = "@NOMMONEDA;" & tdbcMoneda.Text & ";True"
        matriz_fecha(7) = "@TITULO;" & CuentaInvBal & " - " & BuscaNombreCuenta(CuentaInvBal, False) & ";True"
'        AbreReporteParam gsDSN, Me, rutaReportes & "RptFormato_0309_34.rpt", crptToWindow, "Anexo al Libro de Inventarios y Balances", "", matriz_fecha(), formulas()
    
    ElseIf ReporteSunat = "PCGE39" Then
        matriz_fecha(0) = "@Emp_cCodigo;" & gsEmpresa & ";True"
        matriz_fecha(1) = "@Pan_cAnio;" & gsAnio & ";True"
        matriz_fecha(2) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
        matriz_fecha(3) = "@moneda;" & tdbcMoneda.BoundText & ";True"
        matriz_fecha(4) = "@EMPRESA;" & gsEmpresaNom & ";True"
        matriz_fecha(5) = "@RUC;" & gsRUC & ";True"
        matriz_fecha(6) = "@Cta_BalInv;" & CuentaInvBal & ";True"
        matriz_fecha(7) = "@NOMMONEDA;" & tdbcMoneda.Text & ";True"
        matriz_fecha(8) = "@TITULO;" & CuentaInvBal & " - " & BuscaNombreCuenta(CuentaInvBal, False) & ";True"
    ElseIf ReporteSunat = "PCGE11" Or ReporteSunat = "PCGE30" Then
entrar:
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
        matriz_fecha(15) = "@Val_nProvTot;0;True"
        matriz_fecha(16) = "@Val_nTotalNeto;0;True"
        matriz_fecha(17) = "@NOMMONEDA;" & tdbcMoneda.Text & ";True"
        matriz_fecha(18) = "@TITULO;" & CuentaInvBal & " - " & BuscaNombreCuenta(CuentaInvBal, False) & ";True"
        matriz_fecha(19) = "@Cuenta;" & Right(ReporteSunat, 2) & ";True"
        
        AbreReporteParam gsDSN, Me, rutaReportes & "RptFormato_0308_PCGE.rpt", crptToWindow, "Anexo al Libro de Inventarios y Balances", "", matriz_fecha(), formulas()
        Screen.MousePointer = vbDefault
        cmdImprimir.Enabled = True
        Exit Sub
    ElseIf ReporteSunat = "PCGE32" Then
        matriz_fecha(0) = "@Tipo;LISTAR;True"
        matriz_fecha(1) = "@Emp_cCodigo;" & gsEmpresa & ";True"
        matriz_fecha(2) = "@Pan_cAnio;" & gsAnio & ";True"
        matriz_fecha(3) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
        matriz_fecha(4) = "@nombreMes;" & tdbcMes.Text & ";True"
'        matriz_fecha(3) = "@moneda;" & tdbcMoneda.BoundText & ";True"
        matriz_fecha(5) = "@EMPRESA;" & gsEmpresaNom & ";True"
        matriz_fecha(6) = "@RUC;" & gsRUC & ";True"
'        matriz_fecha(6) = "@Cta_BalInv;" & CuentaInvBal & ";True"
        matriz_fecha(7) = "@NOMMONEDA;" & tdbcMoneda.Text & ";True"
        matriz_fecha(8) = "@TITULO;" & CuentaInvBal & " - " & BuscaNombreCuenta(CuentaInvBal, False) & ";True"
    End If
    
    If CE(matriz_fecha(0)) = "" Then
        Mensajes "Reporte no encontrado, en el directorio de reportes del sistema", vbExclamation + vbOKOnly
    Else
        AbreReporteParam gsDSN, Me, rutaReportes & "RptFormato_" & cRpt & ".rpt", crptToWindow, "Saldos de Cuentas de Balance", "", matriz_fecha(), formulas()
    End If
    Screen.MousePointer = vbDefault
    cmdImprimir.Enabled = True
     
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    
    Me.Caption = Titulo(Me.Caption, TituloSunat)
    
    Call Centrar_form(Me)

    Call LlenaCombos
    Call BuscarMonedaNacional
    
    If Left(ReporteSunat, 5) = "PCGE2" Then
        Call VisibleMetodo(True)
    Else
        Call VisibleMetodo(False)
    End If
    
End Sub

Private Sub VisibleMetodo(bValor As Boolean)
    lblMetodo.Visible = bValor
    tdbcMetodo.Visible = bValor
End Sub

Private Sub LlenaCombos()
    Dim sqlcombos As String
    Dim entro As Boolean
    
    entro = False
    
    'If InStr(1, TituloSunat, "8.1") > 0 Or InStr(1, TituloSunat, "14.1") > 0 Then
        Call LlenaComboMesApeAddItem(tdbcMes)
        
        entro = True
    'Else
    '    Call LlenaComboMesApeAddItem(tdbcMes)
    'End If
    
    
    DoEvents
    tdbcMes.ReBind
    
    If entro = False Then
        If gsPeriodo = "" Then
            tdbcMes.BoundText = "00"
        Else
            tdbcMes.BoundText = gsPeriodo
        End If
    Else
        If gsPeriodo >= "00" And gsPeriodo < "13" Then
            tdbcMes.BoundText = gsPeriodo
        Else
            tdbcMes.BoundText = "01"
        End If
    
    End If
    
    DoEvents
    
    sqlcombos = "SELECT Mon_cCodigo, Mon_cNombreLargo, Mon_cMNac, Mon_cMExt From CNT_TIPO_MONEDA " & _
                "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND (Mon_cMNac = '1' or Mon_cMExt = '1') " & _
                "ORDER BY Mon_cNombreLargo"
    LlenarComboAddItem tdbcMoneda, sqlcombos
    
    '---------------------------
    
    sqlcombos = "SELECT Tab_cCodigo, Tab_cDescripCampo FROM TABLA WITH(NOLOCK) " & _
                "WHERE Emp_cCodigo='" & gsEmpresa & "' AND Tab_cTabla = '085' " & _
                "ORDER BY Tab_cCodigo"
                
    LlenarComboAddItem tdbcMetodo, sqlcombos
    
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
        Call Centrar_Objeto(picForm, Me, -50)

        Call CentrarTitulo(lblTitulo, fraTodo, Me)
    End If
Exit Sub
errHand:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmRepAnexoInvBalance = Nothing
End Sub

Private Sub tdbcMes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pSetFocus tdbcMoneda
End If
End Sub

Private Sub tdbcMoneda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then pSendKeys "{tab}"
End Sub

