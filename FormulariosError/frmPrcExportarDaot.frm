VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmPrcExportarDaot 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportar Datos DAOT"
   ClientHeight    =   4005
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   8235
   Icon            =   "frmPrcExportarDaot.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4005
   ScaleWidth      =   8235
   Begin VB.Frame fraTodo 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3990
      Left            =   45
      TabIndex        =   3
      Top             =   0
      Width           =   8145
      Begin VB.Frame Frame3 
         Height          =   3990
         Left            =   4095
         TabIndex        =   6
         Top             =   0
         Width           =   4065
         Begin TDBNumber6Ctl.TDBNumber tdbnMonto 
            Height          =   300
            Left            =   1830
            TabIndex        =   7
            Top             =   1170
            Width           =   1935
            _Version        =   65536
            _ExtentX        =   3413
            _ExtentY        =   529
            Calculator      =   "frmPrcExportarDaot.frx":0ECA
            Caption         =   "frmPrcExportarDaot.frx":0EEA
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmPrcExportarDaot.frx":0F56
            Keys            =   "frmPrcExportarDaot.frx":0F74
            Spin            =   "frmPrcExportarDaot.frx":0FCC
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###,###,##0.00"
            EditMode        =   3
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "###,###,##0.00"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   999999999
            MinValue        =   0
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   0
            MaxValueVT      =   1802698757
            MinValueVT      =   1769209861
         End
         Begin TDBDate6Ctl.TDBDate dtpDesde 
            Height          =   300
            Left            =   1830
            TabIndex        =   8
            Tag             =   "enabled"
            Top             =   1575
            Width           =   1935
            _Version        =   65536
            _ExtentX        =   3413
            _ExtentY        =   529
            Calendar        =   "frmPrcExportarDaot.frx":0FF4
            Caption         =   "frmPrcExportarDaot.frx":10F6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmPrcExportarDaot.frx":115A
            Keys            =   "frmPrcExportarDaot.frx":1178
            Spin            =   "frmPrcExportarDaot.frx":11E4
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
            Left            =   1830
            TabIndex        =   9
            Tag             =   "enabled"
            Top             =   1980
            Width           =   1935
            _Version        =   65536
            _ExtentX        =   3413
            _ExtentY        =   529
            Calendar        =   "frmPrcExportarDaot.frx":120C
            Caption         =   "frmPrcExportarDaot.frx":130E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmPrcExportarDaot.frx":1372
            Keys            =   "frmPrcExportarDaot.frx":1390
            Spin            =   "frmPrcExportarDaot.frx":13FC
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
         Begin TrueOleDBList70.TDBCombo tdbcMoneda 
            Height          =   300
            Left            =   1830
            TabIndex        =   10
            Tag             =   "_"
            Top             =   2370
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
            _PropDict       =   $"frmPrcExportarDaot.frx":1424
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
         Begin TrueOleDBList70.TDBCombo tdbcEntidad 
            Height          =   300
            Left            =   1845
            TabIndex        =   11
            Tag             =   "_"
            Top             =   765
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
            _PropDict       =   $"frmPrcExportarDaot.frx":14AB
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
         Begin MSForms.CommandButton cmdExportar 
            Height          =   480
            Left            =   990
            TabIndex        =   18
            Top             =   3150
            Width           =   2115
            Caption         =   "   Exportar Datos"
            PicturePosition =   327683
            Size            =   "3731;847"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "MONTO MAYOR A "
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
            Left            =   90
            TabIndex        =   17
            Top             =   1200
            Width           =   1500
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
            Index           =   0
            Left            =   90
            TabIndex        =   16
            Top             =   2415
            Width           =   765
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
            Index           =   2
            Left            =   90
            TabIndex        =   15
            Top             =   2010
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
            Index           =   0
            Left            =   90
            TabIndex        =   14
            Top             =   1605
            Width           =   630
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "DATOS DEL DAOT"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   1035
            TabIndex        =   13
            Top             =   225
            Width           =   2205
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "ENTIDAD"
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
            Index           =   2
            Left            =   90
            TabIndex        =   12
            Top             =   765
            Width           =   780
         End
      End
      Begin VB.TextBox tdbtDirectorio 
         Height          =   315
         Left            =   270
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   3465
         Width           =   3600
      End
      Begin VB.DirListBox Dir1 
         Height          =   1890
         Left            =   315
         TabIndex        =   1
         Top             =   1035
         Width           =   3570
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   315
         TabIndex        =   0
         Top             =   675
         Width           =   3570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Directorio seleccionado:"
         DragMode        =   1  'Automatic
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
         Left            =   315
         TabIndex        =   5
         Top             =   3150
         Width           =   2130
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "DESTINO DE ARCHIVOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   1035
         TabIndex        =   4
         Top             =   270
         Width           =   2205
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
      TabIndex        =   19
      Top             =   0
      Width           =   4365
   End
End
Attribute VB_Name = "frmPrcExportarDaot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit

Dim gsGrupo As String
Dim ArchivoDest As String

Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

Public Function CopiaArchivosDAOT(ruta As String) As Boolean

    Dim fso As New Scripting.filesystemobject
    Dim archivo As String
    
    archivo = "DAOT.dbf"
        
    If tdbcEntidad.BoundText = "C" Then
     ArchivoDest = "Ingresos.dbf"
    End If
    
    If tdbcEntidad.BoundText = "P" Then
     ArchivoDest = "Costos.dbf"
    End If
        
    Dim RutaDaot As String
    RutaDaot = ruta
    
    If Right(CE(RutaDaot), 1) <> "\" Then RutaDaot = RutaDaot & "\"
    If fso.FileExists(RutaDaot & archivo) Then fso.DeleteFile RutaDaot & archivo
    If fso.FileExists(RutaDaot & ArchivoDest) Then fso.DeleteFile RutaDaot & ArchivoDest
    On Error GoTo NoSePuedeCopiar
    
    Call EscribirLog("Iniciando exportacion de datos DAOT de la empresa " & gsEmpresa & " " & gsEmpresaNom, Me.Name)
    
    'fso.CopyFile App.Path & "\DBF\" & archivo, RutaDaot
    CopyFile App.Path & "\DBF\" & archivo, RutaDaot & ArchivoDest, 0
        
    Call EscribirLog("Finalizando exportacion de datos DAOT de la empresa " & gsEmpresa & " " & gsEmpresaNom, Me.Name)
    
    CopiaArchivosDAOT = True
    Exit Function
NoSePuedeCopiar:
    Mensajes "No se puede copiar el archivo : " & Salto(1) & Err.Description, vbOKOnly + vbInformation
    CopiaArchivosDAOT = False
    Call EscribirLog("Error al exportacion de datos DAOT, [" & Err.Description & "] de la empresa " & gsEmpresa & " " & gsEmpresaNom, Me.Name)
End Function



Private Sub cmdExportar_Click()
    On Error GoTo serror
    Dim respuesta As String
    
   Screen.MousePointer = vbNormal
    If UCase(Dir1.Path) = UCase(App.Path & "\DBF") Then
        Mensajes "Seleccione otro directorio para la exportación" & Salto(1) & "Este es un directorio de sistema reservado"
        Exit Sub
    End If
    
    
    If VerificaTcDAOT(dtpDesde.Value, dtpHasta.Value, tdbnMonto.Value, tdbcMoneda.BoundText, tdbcEntidad.BoundText) = False Then
        Exit Sub
    End If
   
    If tdbcEntidad.Text = "" Then
        Mensajes "Seleccione El tipo de entidad" & Salto(2) & "Si la lista esta vacia active las casilla " & Salto(1) & "incluir en DAOT en el formulario, Tipo de entidad"
        pSetFocus tdbcEntidad
        Exit Sub
    End If
    

    ' *** Confirmando exportación
    respuesta = MsgBox("Desea exportar la información seleccionada", vbYesNo + vbQuestion, "Confirmar Exportar Datos")
    If respuesta = vbNo Then Exit Sub
    
    
    Screen.MousePointer = vbHourglass
    
    If CopiaArchivosDAOT(CE(tdbtDirectorio)) = True Then
        Exportar
    Else
        Mensajes "No se termino la exportacion correctamente", vbInformation
    End If
    
    
    
    Screen.MousePointer = vbNormal
    
    Exit Sub
serror:
    Mensajes Err.Description
    
End Sub

Private Sub Form_Load()
    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))
    
    
    Call Centrar_form(Me)
    
    Me.Height = 4515
    Me.Width = 8415
    
    'Call LlenaComboMesAddItem(tdbcMes)
    Drive1.Drive = "C:\"
    Dir1.Path = "C:\"
    Dir1.Refresh
    
    tdbtDirectorio.Text = Dir1.Path
    
    dtpHasta = "31/12/" & gsAnio       'fechaServidor
    dtpDesde = "01/01/" & gsAnio        'dtpHasta
    
        
    If Mid(gsGrupo, 3, 1) <> "1" And gsGrupo <> gsPrivilegioAdmin Then
        Me.cmdExportar.Enabled = False

    Else
        Me.cmdExportar.Enabled = True
        
    End If
    
    LlenaCombos
   
End Sub

Private Sub LlenaCombos()
    Dim sqlcombos As String
    ' *** Llenando el tipo de Moneda
    sqlcombos = "SELECT Mon_cCodigo, Mon_cNombreLargo, Mon_cMNac, Mon_cMExt From CNT_TIPO_MONEDA " & _
                "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND (Mon_cMNac = '1' ) " & _
                "ORDER BY Mon_cNombreLargo"
    LlenarComboAddItem tdbcMoneda, sqlcombos
    tdbcMoneda.BoundText = gsMonedaNac
    
    sqlcombos = "select ten_ctipoentidad, ten_cnombreentidad  from CNT_ENTIDAD  " & _
                "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Ten_cDaot='1' "
    LlenarComboAddItem tdbcEntidad, sqlcombos


    
End Sub

Private Sub Dir1_Change()
    tdbtDirectorio.Text = Dir1.Path
End Sub

Private Sub Drive1_Change()
    On Error GoTo serror
    Dir1.Path = Drive1.Drive
    Dir1.Refresh
    tdbtDirectorio.Text = Dir1.Path
    Exit Sub
serror:
    Drive1.ListIndex = 1
    Dir1.Path = "C:\"
    Dir1.Refresh
    tdbtDirectorio.Text = Dir1.Path
    
End Sub

Private Function BuscaDatosDAOT(FechaIni As String, FechaFin As String, Monto As Double, Moneda As String, TCAnual As Double) As ADODB.Recordset
    On Error GoTo serror
    
    Dim lrsAsientos As ADODB.Recordset
    Dim sqlSp As String
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
    
    Set clDatos = New clsMantoTablas
    Set lrsAsientos = New ADODB.Recordset

   
    sqlSp = " spCn_RptDaot 'EXPORTACION', '" & gsEmpresa & "', '" & gsAnio & "', '" & FechaIni & "', '" & FechaFin & "'," & Monto & ", '" & Moneda & "','" & tdbcEntidad.BoundText & "'," & NE(TCAnual)
    arrDatos = Array(sqlSp)
    Set lrsAsientos = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    
    If Not lrsAsientos Is Nothing Then
        If lrsAsientos.State = 0 Then
            Set BuscaDatosDAOT = Nothing
        Else
            Set BuscaDatosDAOT = lrsAsientos
        End If
    Else
        Set BuscaDatosDAOT = Nothing
    End If
    
    
    Set lrsAsientos = Nothing
    Set clDatos = Nothing
    
    Exit Function
serror:
    Mensajes Err.Description
End Function

Private Sub Exportar()

    Dim fso As New Scripting.filesystemobject
    Dim archivo As String
    
    'archivo = "DAOT.dbf"

    'If fso.FileExists(tdbtDirectorio & "\" & archivo) = False Then
    '    Mensajes "No se encuentra copio el archivo al directorio seleccionado", vbOKOnly + vbInformation
    '    Exit Sub
    'End If


    If fso.FileExists(tdbtDirectorio & "\" & ArchivoDest) = False Then
        Mensajes "No se encuentra copio el archivo al directorio seleccionado", vbOKOnly + vbInformation
        Exit Sub
    End If


    Dim i As Integer
    Dim ruta As String
    Dim rutaExp As String
    Dim sqlEnt As String
    Dim cnExp As New ADODB.Connection
    Dim rsDatos As ADODB.Recordset
    On Error GoTo ERROR
    Set rsDatos = BuscaDatosDAOT(dtpDesde, dtpHasta, tdbnMonto.Text, tdbcMoneda.BoundText, 0)
    
    If Not rsDatos Is Nothing Then
        ruta = Trim(tdbtDirectorio)
        rutaExp = "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=Driver={Microsoft Visual FoxPro Driver};UID=;SourceDB=" & ruta & ";SourceType=DBF;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;Null=Yes;Deleted=Yes;"
        cnExp.ConnectionString = rutaExp
        cnExp.ConnectionTimeout = 0
        'cnExp.ConnectionTimeout = 30
        cnExp.Open
        i = 0
    
        Dim cTotal As String
        
        Do While Not rsDatos.EOF
            i = i + 1
            'cTotal = CE(Int(NE(rsDatos("Stotal")) * 100))
            cTotal = CE(Round(rsDatos("Stotal"), 0))
            'cTotal = CE(rsDatos("Stotal"))
            cTotal = Replace(cTotal, ",", "")
            cTotal = Replace(cTotal, ".", "")
            rsDatos("ent_cpersona") = Replace(rsDatos("ent_cpersona").Value, "'", "")
            
            cTotal = cTotal
            Select Case Trim(tdbcEntidad.BoundText)
             Case "C"
                sqlEnt = " Insert into Ingresos (contador, d_tipodoc, d_numdoc, periodo, tipo_per, tipo_doc , " & _
                         "num_doc , importe, ap_pater, ap_mater, nombre1, nombre2, razon_soc) " & _
                         "Values ( '" & i & "', " & _
                         "'" & CE(rsDatos("TipoDec")) & "'," & _
                         "'" & CE(rsDatos("RUCDec")) & "'," & _
                         "'" & CE(rsDatos("Asd_cPeriodo")) & "'," & _
                         "'" & CE(rsDatos("Ten_ctipoentidad")) & "'," & _
                         "'" & CE(rsDatos("Asd_cTipoDoc")) & "'," & _
                         "'" & CE(rsDatos("Ent_nRuc")) & "'," & _
                         "'" & cTotal & "'," & _
                         "'" & PrimerApellido((rsDatos("ent_cpersona")), CE(rsDatos("Ent_nRuc"))) & "'," & _
                         "'" & SegundoApellido((rsDatos("ent_cpersona")), CE(rsDatos("Ent_nRuc"))) & "'," & _
                         "'" & PrimerNombre((rsDatos("ent_cpersona")), CE(rsDatos("Ent_nRuc"))) & "'," & _
                         "'" & SegundoNombre((rsDatos("ent_cpersona")), CE(rsDatos("Ent_nRuc"))) & "'," & _
                         "'" & RazonSocial(CE(rsDatos(Trim("ent_cpersona"))), CE(rsDatos("Ent_nRuc"))) & "')"
                cnExp.Execute sqlEnt
             Case "P"
                sqlEnt = " Insert into Costos (contador, d_tipodoc, d_numdoc, periodo, tipo_per, tipo_doc , " & _
                         "num_doc , importe, ap_pater, ap_mater, nombre1, nombre2, razon_soc) " & _
                         "Values ( '" & i & "', " & _
                         "'" & CE(rsDatos("TipoDec")) & "'," & _
                         "'" & CE(rsDatos("RUCDec")) & "'," & _
                         "'" & CE(rsDatos("Asd_cPeriodo")) & "'," & _
                         "'" & CE(rsDatos("Ten_ctipoentidad")) & "'," & _
                         "'" & CE(rsDatos("Asd_cTipoDoc")) & "'," & _
                         "'" & CE(rsDatos("Ent_nRuc")) & "'," & _
                         "'" & cTotal & "'," & _
                         "'" & PrimerApellido((rsDatos("ent_cpersona")), CE(rsDatos("Ent_nRuc"))) & "'," & _
                         "'" & SegundoApellido((rsDatos("ent_cpersona")), CE(rsDatos("Ent_nRuc"))) & "'," & _
                         "'" & PrimerNombre((rsDatos("ent_cpersona")), CE(rsDatos("Ent_nRuc"))) & "'," & _
                         "'" & SegundoNombre((rsDatos("ent_cpersona")), CE(rsDatos("Ent_nRuc"))) & "'," & _
                         "'" & RazonSocial(CE(rsDatos("ent_cpersona")), CE(rsDatos("Ent_nRuc"))) & "')"
                cnExp.Execute sqlEnt
            End Select
            rsDatos.MoveNext
        Loop
        
        'Dim sNewArchivo
        
        'ruta = tdbtDirectorio
        
        'If Right(ruta, 1) <> "\" Then ruta = ruta & "\"

        'Select Case tdbcEntidad.BoundText
        '       Case "C":
        '            sNewArchivo = "Ingresos.dbf"
        '            fso.CopyFile ruta & archivo, ruta & sNewArchivo
        '       Case "P":
        '            sNewArchivo = "Costos.dbf"
        '            fso.CopyFile ruta & archivo, ruta & sNewArchivo
        'End Select

        Mensajes "Los Datos se exportaron correctamente", vbInformation

    Else
        Mensajes "No se encontraron registros.", vbInformation
    End If
    
    Call CerrarRecordSet(rsDatos)
    If cnExp.State = adStateOpen Then cnExp.Close
    Set cnExp = Nothing
    
    Exit Sub
    
ERROR:
    Mensajes Err.Description, vbOKOnly + vbInformation
    
End Sub

Private Function RazonSocial(cadena As String, Ruc As String)
    If Left(Ruc, 1) = "2" Then
'        RazonSocial = ""
'    Else
        RazonSocial = cadena
    Else
        RazonSocial = ""
    End If
End Function

Private Function PrimerNombre(cadena As String, Ruc As String)
    On Error GoTo ERROR
    PrimerNombre = ""
'    If Left(Ruc, 1) = "1" Then
    If Left(Ruc, 1) <> "2" And (Len(Ruc) = 11 Or Len(Ruc) = 8) Then
        Dim NuevaCad As String
        Dim PosIni As Integer
        Dim PosFin As Integer
        NuevaCad = cadena & " "
        PosIni = 1
        If Len(cadena) > 0 Then
           PosIni = InStr(1, NuevaCad, " ")
           NuevaCad = Mid(NuevaCad, PosIni + 1, Len(NuevaCad))
        
           PosIni = InStr(1, NuevaCad, " ")
           NuevaCad = Mid(NuevaCad, PosIni + 1, Len(NuevaCad))
        
           PosIni = 1
           PosFin = InStr(1, NuevaCad, " ")
           
           If PosFin = 0 Then PosFin = Len(NuevaCad)
           
           NuevaCad = Mid(NuevaCad, PosIni, PosFin)
           PrimerNombre = NuevaCad
        End If
    End If
    If PrimerNombre = "" Then PrimerNombre = PrimerApellido(cadena, Ruc)
    Exit Function
    
ERROR:
    PrimerNombre = ""
End Function

Private Function SegundoNombre(cadena As String, Ruc As String)
    On Error GoTo ERROR
    SegundoNombre = ""
'    If Left(Ruc, 1) = "1" Then
    If Left(Ruc, 1) <> "2" And (Len(Ruc) = 11 Or Len(Ruc) = 8) Then
        Dim NuevaCad As String
        Dim PosIni As Integer
        Dim PosFin As Integer
        NuevaCad = cadena & " "
        PosIni = 1
        If Len(cadena) > 0 Then
           PosIni = InStr(1, NuevaCad, " ")
           NuevaCad = Mid(NuevaCad, PosIni + 1, Len(NuevaCad))
        
           PosIni = InStr(1, NuevaCad, " ")
           NuevaCad = Mid(NuevaCad, PosIni + 1, Len(NuevaCad))
        
           PosIni = InStr(1, NuevaCad, " ")
           NuevaCad = Mid(NuevaCad, PosIni + 1, Len(NuevaCad))
        
           PosIni = 1
           
           NuevaCad = Mid(NuevaCad, PosIni, Len(NuevaCad))
           SegundoNombre = NuevaCad
        End If
    End If
    Exit Function
ERROR:
    SegundoNombre = ""
End Function

Private Function PrimerApellido(cadena As String, Ruc As String)
    On Error GoTo ERROR
    PrimerApellido = ""
'    If Left(Ruc, 1) = "1" Then
    If Left(Ruc, 1) <> "2" And (Len(Ruc) = 11 Or Len(Ruc) = 8) Then
        Dim NuevaCad As String
        Dim PosIni As Integer
        Dim PosFin As Integer
        NuevaCad = cadena & " "
        PosIni = 1
        If Len(cadena) > 0 Then
           PosFin = InStr(1, NuevaCad, " ")
           
           If PosFin = 0 Then PosFin = Len(NuevaCad)
           
           NuevaCad = Mid(NuevaCad, PosIni, PosFin)
           PrimerApellido = NuevaCad
        End If
    End If
    Exit Function
ERROR:
    PrimerApellido = ""
End Function

Private Function SegundoApellido(cadena As String, Ruc As String)
    On Error GoTo ERROR
    SegundoApellido = ""
'    If Left(Ruc, 1) = "1" Then
    If Left(Ruc, 1) <> "2" And (Len(Ruc) = 11 Or Len(Ruc) = 8) Then
        Dim NuevaCad As String
        Dim PosIni As Integer
        Dim PosFin As Integer
        NuevaCad = cadena & " "
        PosIni = 1
        If Len(cadena) > 0 Then
           PosIni = InStr(1, NuevaCad, " ")
           NuevaCad = Mid(NuevaCad, PosIni + 1, Len(NuevaCad))
        
           PosIni = 1
           PosFin = InStr(1, NuevaCad, " ")
           
           If PosFin = 0 Then PosFin = Len(NuevaCad)
           
           NuevaCad = Mid(NuevaCad, PosIni, PosFin)
           SegundoApellido = NuevaCad
        End If
    End If
    Exit Function
ERROR:
    SegundoApellido = ""
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
Private Sub tdbnMonto_GotFocus()
    tdbnMonto.SelStart = 0
    tdbnMonto.SelLength = Len(tdbnMonto.Text)
End Sub
