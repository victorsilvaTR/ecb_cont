VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmRepLibroBancos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libro Caja Bancos"
   ClientHeight    =   5160
   ClientLeft      =   3705
   ClientTop       =   3165
   ClientWidth     =   7005
   Icon            =   "frmRepLibroBancos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5160
   ScaleWidth      =   7005
   Begin VB.Frame fraTodo 
      Height          =   5010
      Left            =   45
      TabIndex        =   8
      Top             =   45
      Width           =   6900
      Begin VB.Frame Frame4 
         Height          =   1575
         Left            =   225
         TabIndex        =   9
         Top             =   1800
         Width           =   6495
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
            TabIndex        =   3
            Top             =   240
            Width           =   2775
         End
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
            TabIndex        =   2
            Top             =   240
            Value           =   -1  'True
            Width           =   2535
         End
         Begin TDBText6Ctl.TDBText tdbtCuentaDesde 
            Height          =   315
            Left            =   1125
            TabIndex        =   4
            Tag             =   "_"
            Top             =   585
            Width           =   1965
            _Version        =   65536
            _ExtentX        =   3466
            _ExtentY        =   556
            Caption         =   "frmRepLibroBancos.frx":0ECA
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmRepLibroBancos.frx":0F36
            Key             =   "frmRepLibroBancos.frx":0F54
            BackColor       =   -2147483643
            EditMode        =   0
            ForeColor       =   -2147483640
            ReadOnly        =   0
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
            Left            =   3135
            TabIndex        =   10
            Tag             =   "_"
            Top             =   585
            Width           =   3045
            _Version        =   65536
            _ExtentX        =   5371
            _ExtentY        =   556
            Caption         =   "frmRepLibroBancos.frx":0F88
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmRepLibroBancos.frx":0FF4
            Key             =   "frmRepLibroBancos.frx":1012
            BackColor       =   -2147483643
            EditMode        =   0
            ForeColor       =   -2147483640
            ReadOnly        =   -1
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
            Format          =   ""
            FormatMode      =   1
            AutoConvert     =   -1
            ErrorBeep       =   0
            MaxLength       =   200
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
            Left            =   1125
            TabIndex        =   5
            Tag             =   "_"
            Top             =   945
            Width           =   1965
            _Version        =   65536
            _ExtentX        =   3466
            _ExtentY        =   556
            Caption         =   "frmRepLibroBancos.frx":1046
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmRepLibroBancos.frx":10B2
            Key             =   "frmRepLibroBancos.frx":10D0
            BackColor       =   -2147483643
            EditMode        =   0
            ForeColor       =   -2147483640
            ReadOnly        =   0
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
         Begin TDBText6Ctl.TDBText tdbtDescripcionHasta 
            Height          =   315
            Left            =   3135
            TabIndex        =   11
            Tag             =   "_"
            Top             =   945
            Width           =   3045
            _Version        =   65536
            _ExtentX        =   5371
            _ExtentY        =   556
            Caption         =   "frmRepLibroBancos.frx":1104
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmRepLibroBancos.frx":1170
            Key             =   "frmRepLibroBancos.frx":118E
            BackColor       =   -2147483643
            EditMode        =   0
            ForeColor       =   -2147483640
            ReadOnly        =   -1
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
            Format          =   ""
            FormatMode      =   1
            AutoConvert     =   -1
            ErrorBeep       =   0
            MaxLength       =   200
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
            TabIndex        =   13
            Top             =   1005
            Width           =   495
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
            Index           =   2
            Left            =   375
            TabIndex        =   12
            Top             =   645
            Width           =   555
         End
      End
      Begin TrueOleDBList70.TDBCombo tdbcMes 
         Height          =   300
         Left            =   3075
         TabIndex        =   0
         Top             =   810
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
         _PropDict       =   $"frmRepLibroBancos.frx":11CA
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
         Left            =   3075
         TabIndex        =   1
         Tag             =   "_"
         Top             =   1170
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
         _PropDict       =   $"frmRepLibroBancos.frx":1251
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
         Height          =   990
         Left            =   225
         TabIndex        =   17
         Top             =   3330
         Width           =   6495
         Begin MSForms.OptionButton OptImpresion 
            Height          =   510
            Index           =   1
            Left            =   3465
            TabIndex        =   19
            Top             =   270
            Width           =   2400
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   5
            Size            =   "4233;900"
            Value           =   "0"
            Caption         =   "Impresión Formato Láser"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.OptionButton OptImpresion 
            Height          =   420
            Index           =   0
            Left            =   585
            TabIndex        =   18
            Top             =   315
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
      Begin MSForms.CommandButton cmdImprimir 
         Height          =   435
         Left            =   1620
         TabIndex        =   6
         Top             =   4410
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
         Left            =   3465
         TabIndex        =   7
         Top             =   4410
         Width           =   1665
         Caption         =   " Salir"
         PicturePosition =   327683
         Size            =   "2937;767"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
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
         Left            =   1995
         TabIndex        =   15
         Top             =   810
         Width           =   375
      End
      Begin VB.Label lblMoneda 
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
         Left            =   1995
         TabIndex        =   14
         Top             =   1170
         Width           =   765
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
      Height          =   435
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   4365
   End
End
Attribute VB_Name = "frmRepLibroBancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ReporteSunat As String
Public TituloSunat As String
Public rsArreglo  As ADODB.Recordset

Dim EntroCabecera As Boolean
Dim Control As String
Dim iReport, nAncho As Integer
Dim gsGrupo, Termino1, Termino2 As String

Dim pVanDebe As String * 15
Dim pVanHaber As String * 13
Dim pVienenDebe As String * 12
Dim pVienenHaber As String * 13
Dim pVanHaberME As String * 12
Dim pVienenHaberME As String * 12
Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

Public Sub CerrarForm()
    Unload Me
End Sub
Private Sub cmdImprimir_Click()
    If CE(tdbcMoneda.Text) = "" Then
        Mensajes "seleccione una moneda"
        pSetFocus tdbcMoneda
        Exit Sub
    End If
    
    If optCuentas.Value = True Then
        If TextoLleno2(tdbtCuentaDesde, "Ingrese la cuenta Inicial") = False Then Exit Sub
        If TextoLleno2(tdbtCuentaHasta, "Ingrese la cuenta Final") = False Then Exit Sub
        
        If CE(tdbtCuentaDesde.Text) > CE(tdbtCuentaHasta.Text) Then
           Mensajes "La cuenta inicial debe ser menor o igual a la cuenta final"
           Exit Sub
        End If
        
    End If
    
    Dim matriz_fecha(8) As Variant
    ' ***
    Screen.MousePointer = vbHourglass
    matriz_fecha(0) = "@Emp_cCodigo;" & gsEmpresa & ";True"
    matriz_fecha(1) = "@Pan_cAnio;" & gsAnio & ";True"
    matriz_fecha(2) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
    matriz_fecha(3) = "@Pla_cCtaDesde;" & tdbtCuentaDesde & ";True"
    matriz_fecha(4) = "@Pla_cCtaHasta;" & tdbtCuentaHasta & ";True"
    matriz_fecha(5) = "@moneda;" & tdbcMoneda.BoundText & ";True"
    
    matriz_fecha(6) = "@EMPRESA;" & gsEmpresaNom & ";True"
    matriz_fecha(7) = "@RUC;" & gsRUC & ";True"
    matriz_fecha(8) = "@NOMMONEDA;" & tdbcMoneda.Text & ";True"
    
    cmdImprimir.Enabled = False
    Dim formulas(0) As Variant
        
    If ReporteSunat = "" Then
        AbreReporteParam gsDSN, Me, rutaReportes & "RptLibroCajaBancos.rpt", crptToWindow, "Libro Caja Bancos", "", matriz_fecha(), formulas()
    
    'ElseIf ReporteSunat = "F0101" Then 'Detalle Movimiento en Efectivo (Laser)
    '    AbreReporteParam gsDSN, Me, rutaReportes & "RptFormato_0101.rpt", crptToWindow, "Libro Caja Bancos", "", matriz_fecha(), formulas()
    
    ElseIf ReporteSunat = "F0101" And OptImpresion(1).Value Then 'Detalle Movimiento en Efectivo (Laser)
        AbreReporteParam gsDSN, Me, rutaReportes & "RptFormato_0101.rpt", crptToWindow, "Libro Caja Bancos", "", matriz_fecha(), formulas()
    ElseIf ReporteSunat = "F0101" And OptImpresion(0).Value Then 'Detalle Movimiento en Efectivo (Matricial)
        gsNombreVista = "Libro Banco Detalle Efectivo"
        ImpDetEfMat
        
    'ElseIf ReporteSunat = "F0102" Then
    '    AbreReporteParam gsDSN, Me, rutaReportes & "RptFormato_0102.rpt", crptToWindow, "Libro Caja Bancos", "", matriz_fecha(), formulas()
    
    ElseIf ReporteSunat = "F0102" And OptImpresion(1).Value Then 'Detalle Movimiento en Efectivo (Laser)
        AbreReporteParam gsDSN, Me, rutaReportes & "RptFormato_0102.rpt", crptToWindow, "Libro Caja Bancos", "", matriz_fecha(), formulas()
    ElseIf ReporteSunat = "F0102" And OptImpresion(0).Value Then 'Detalle Movimiento en Efectivo (Matricial)
        gsNombreVista = "Libro Banco Detalle Movimientos Cta.Cte"
        ImpDetCtaCteMat
    End If

    Screen.MousePointer = vbNormal
    cmdImprimir.Enabled = True
    'gsNombreVista = ""
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub


Private Sub Form_Activate()
    'tdbtCuentaDesde_LostFocus
        
End Sub

Private Sub Form_Load()
    Me.Caption = Titulo(Me.Caption, TituloSunat)
    
    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))

    
    Call Centrar_form(Me)

    Call LlenaCombos
    
    optTodos_Click
    
    DoEvents
    pSetFocus tdbtCuentaDesde
    
    tdbcMes.BoundText = gsPeriodo
    tdbcMoneda.BoundText = gsMonedaNac
    
    tdbcMes.ReBind
    tdbcMoneda.ReBind
    

End Sub

Private Sub LlenaCombos()
    Dim sqlcombos As String
    
    Call LlenaComboMesApeAddItem(tdbcMes)
    ' *** Llenando el tipo de Moneda
    sqlcombos = "SELECT Mon_cCodigo, Mon_cNombreLargo, Mon_cMNac, Mon_cMExt From CNT_TIPO_MONEDA " & _
                "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND (Mon_cMNac = '1' or Mon_cMExt = '1') " & _
                "ORDER BY Mon_cNombreLargo"
    LlenarComboAddItem tdbcMoneda, sqlcombos
    tdbcMes.ReBind
    tdbcMoneda.ReBind
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
    Set frmRepLibroBancos = Nothing
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

Private Sub tdbcMes_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then pSendKeys "{tab}"
End Sub

Private Sub tdbcMoneda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then pSendKeys "{tab}"
End Sub

Private Sub tdbtCuentaDesde_Change()
    If CE(tdbtCuentaDesde) = "" Then tdbtDescripcionDesde.Text = ""
End Sub

Private Sub tdbtCuentaDesde_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 112 Then Call LlamaBuscar(frmBuscador, Me.tdbtCuentaDesde.Name, Control, "CuentasFiltCaja", Me, tdbcMes.BoundText, Me.tdbtCuentaDesde.Text)
    
    If KeyCode = 13 Then
        If Left(CE(tdbtCuentaDesde.Text), 2) <> "10" Then
            Mensajes "Ingrese una cuenta de caja"
            tdbtCuentaDesde.Text = ""
            tdbtDescripcionDesde.Text = ""
            pSetFocus tdbtCuentaDesde
        Else
            pSetFocus tdbtCuentaHasta
        End If
    End If
End Sub

Private Sub tdbtCuentaDesde_LostFocus()
'    If ReporteSunat = "F0101" Then
'        If CE(tdbtCuentaDesde.Text) <> "" And Left(CE(tdbtCuentaDesde.Text), 3) > "102" Then
'            tdbtCuentaDesde.Text = ""
'            Mensajes "Cuenta no valida para este reporte"
'            pSetFocus tdbtCuentaDesde
'        End If
'    Else
'        If CE(tdbtCuentaDesde.Text) <> "" And CE(tdbtCuentaDesde.Text) <> "10" And Left(CE(tdbtCuentaDesde.Text), 3) < "103" Then
'            tdbtCuentaDesde.Text = ""
'            Mensajes "Cuenta no valida para este reporte"
'            pSetFocus tdbtCuentaDesde
'        End If
'    End If

    If CE(tdbtCuentaDesde.Text) <> "" And Me.Enabled = True Then
        tdbtDescripcionDesde.Text = ExisteCtaNoTitulo(tdbtCuentaDesde.Text, "")
        If tdbtDescripcionDesde = "" Then pSetFocus tdbtCuentaDesde
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
End Sub



Private Sub tdbtCuentaHasta_Change()
    If CE(tdbtCuentaHasta) = "" Then tdbtDescripcionHasta.Text = ""
End Sub

Private Sub tdbtCuentaHasta_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 112 Then Call LlamaBuscar(frmBuscador, Me.tdbtCuentaHasta.Name, Control, "CuentasFiltCaja", Me, tdbcMes.BoundText, Me.tdbtCuentaHasta.Text)
    
    If KeyCode = 13 Then
        If Left(CE(tdbtCuentaHasta.Text), 2) <> "10" Then
            Mensajes "Ingrese una cuenta de caja"
            tdbtCuentaHasta.Text = ""
            tdbtDescripcionHasta.Text = ""
            pSetFocus tdbtCuentaHasta
        Else
            pSetFocus cmdImprimir
        End If
    End If
    
End Sub

Private Sub tdbtCuentaHasta_LostFocus()
'    If ReporteSunat = "F0101" Then
'        If CE(tdbtCuentaHasta.Text) <> "" And Left(CE(tdbtCuentaHasta.Text), 3) > "102" Then
'            tdbtCuentaHasta.Text = ""
'            Mensajes "Cuenta no valida para este reporte"
'            pSetFocus tdbtCuentaHasta
'        End If
'    Else
'        If CE(tdbtCuentaHasta.Text) <> "" And CE(tdbtCuentaHasta.Text) <> "10" And Left(CE(tdbtCuentaHasta.Text), 3) < "103" Then
'            tdbtCuentaHasta.Text = ""
'            Mensajes "Cuenta no valida para este reporte"
'            pSetFocus tdbtCuentaHasta
'        End If
'    End If

    If CE(tdbtCuentaHasta.Text) <> "" And Me.Enabled = True Then
        tdbtDescripcionHasta.Text = ExisteCtaNoTitulo(tdbtCuentaHasta.Text, "")
        If CE(tdbtDescripcionHasta.Text) = "" Then pSetFocus tdbtCuentaHasta
    End If

End Sub
Sub ImpDetCtaCteMat()
On Error GoTo Control
 gsAccionRep = 3
 gsMesRep = tdbcMes.BoundText
 gsCodMoneda = tdbcMoneda.BoundText
 
 If optTodos.Value Then
  gsCtaIni = "": gsCtaFin = ""
 ElseIf optCuentas.Value Then
  gsCtaIni = Trim(tdbtCuentaDesde.Text)
  gsCtaFin = Trim(tdbtCuentaHasta.Text)
 End If
 
 frmFCImpresion.Show
Exit Sub
Control:
 MsgBox Err.Description
End Sub
Sub ImpDetEfMat()
On Error GoTo Control
 gsAccionRep = 2
 gsMesRep = tdbcMes.BoundText
 gsCodMoneda = tdbcMoneda.BoundText
 
 If optTodos.Value Then
  gsCtaIni = "": gsCtaFin = ""
 ElseIf optCuentas.Value Then
  gsCtaIni = Trim(tdbtCuentaDesde.Text)
  gsCtaFin = Trim(tdbtCuentaHasta.Text)
 End If
 
 frmFCImpresion.Show
  
Exit Sub
Control:
 MsgBox Err.Description
Exit Sub
End Sub
Public Sub ReporteLibBancosDetMovEfectivo()
Dim iContCopias As Integer

Dim pPla_cCuentaContable As String * 12
Dim pPla_cNombreCuenta As String * 21

Dim pAsd_nTipoCambio As String * 5
Dim pSaldoIniME As String * 15
Dim pSaldoIniDeudor As String * 15
Dim pSaldoInicial As String * 13

Dim pAse_nVoucher As String * 11
Dim pAsd_dFecDoc As String * 10
Dim pAsd_cGlosa As String * 31
Dim pDebe As String * 15
Dim pHaber As String * 13
Dim pHaberME As String * 15
Dim pSaldoTotalEfectivo As String * 15

Dim VistapSumDebe As String * 15
Dim VistapSumHaber As String * 15
Dim VistapSaldoAcumDeudor As String * 15
Dim VistapSaldoAcumulado As String * 13
Dim VistapSldoTotalDebe As String * 15
Dim VistapSldoTotalHaber As String * 15
Dim VistapSaldoIniDeudor As String * 15
Dim VistapSaldoInicial As String * 13

Dim VistapFinalDeudor As String * 15
'Dim VistapFinalAcreedor As String * 15
Dim VistapTotalFinDeudor As String * 13
Dim VistapTotalFinAcreedor As String * 13
Dim PlanContable As String * 15
Dim MovMesActual As Boolean
Dim Cont As Integer

Dim pSumHaberME As String * 15: Dim pSumDebe As String * 15: Dim pSumHaber As String * 13
Dim pSaldoAcumME As String * 15: Dim pSaldoAcumDeudor As String * 15: Dim pSaldoAcumulado As String * 13
Dim pTotalHaberME As String * 15: Dim pTotalDebe As String * 15: Dim pTotalHaber As String * 13
Dim pSldoTotalHaberME As String * 15: Dim pSldoTotalDebe As String * 15: Dim pSldoTotalHaber As String * 13

pPla_cCuentaContable = ""
RSet pSaldoTotalEfectivo = Format$(0, "#,###,###,##0.00;-#,###,###,##0.00")

RSet VistapSumDebe = Format$(0, "#,###,###,##0.00;-#,###,###,##0.00")
RSet VistapSumHaber = Format$(0, "#,###,###,##0.00;-#,###,###,##0.00")
RSet VistapSaldoAcumDeudor = Format$(0, "#,###,###,##0.00;-#,###,###,##0.00")
RSet VistapSaldoAcumulado = Format$(0, "#,###,###,##0.00;-#,###,###,##0.00")

RSet VistapSldoTotalDebe = Format$(0, "#,###,###,##0.00;-#,###,###,##0.00")
RSet VistapSldoTotalHaber = Format$(0, "#,###,###,##0.00;-#,###,###,##0.00")

RSet VistapTotalFinDeudor = Format$(0, "#,###,###,##0.00;-#,###,###,##0.00")
RSet VistapTotalFinAcreedor = Format$(0, "#,###,###,##0.00;-#,###,###,##0.00")

RSet VistapFinalDeudor = Format$(0, "#,###,###,##0.00;-#,###,###,##0.00")
'RSet VistapFinalAcreedor = Format$(0, "#,###,###,##0.00;-#,###,###,##0.00")

MovMesActual = False
Cont = 0

On Error GoTo ERROR
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
'  ' Print #1, Chr(27) & Chr(77); '12cpi
'Print #1, Chr(27) & Chr(51) & Chr(29) 'Entre lineas 29/180

giLineas = 0
giEspacios = 60

If Not rsArreglo.EOF Then iReport = 1
If iReport = 1 Then rsArreglo.Sort = "Pla_cCuentaContable": rsArreglo.MoveFirst

gsConTotalPaginas = 0
pSaldoIniDeudor = CDbl(0): pSaldoInicial = CDbl(0)

With rsArreglo
    If .RecordCount > 0 Then
       iReport = 1
       SwHoja = False
       gsPaginaPrincipal = 1
       Call CabeceraDetEfectivo
       .MoveFirst
       
       pTotalHaberME = 0: pTotalDebe = 0: pTotalHaber = 0
       pSldoTotalHaberME = 0: pSldoTotalDebe = 0: pSldoTotalHaber = 0
       
       pVanDebe = 0: pVanHaber = 0
       pVienenDebe = 0: pVienenHaber = 0
       pVanHaberME = 0: pVienenHaberME = 0
       
       Do While Not .EOF
          pPla_cCuentaContable = !Pla_cCuentaContable
          pPla_cNombreCuenta = !Pla_cNombreCuenta
          
          RSet pAsd_nTipoCambio = Format$(IIf(IsNull(!Asd_nTipoCambio), 0, !Asd_nTipoCambio), "#,###0.000;-#,###0.000")
          If Trim(!MNac) <> "1" Then
           RSet pSaldoIniME = Format$(IIf(IsNull(!SaldoSoles), 0, !SaldoSoles), "#,###,###,##0.00;-#,###,###,##0.00")
          Else
           RSet pSaldoIniME = Format$(IIf(IsNull(!SaldoDolares), 0, !SaldoDolares), "#,###,###,##0.00;-#,###,###,##0.00")
          End If
          
          If Trim(!MNac) = "1" Then
           If !SaldoSoles > 0 Then RSet pSaldoIniDeudor = Format$(IIf(IsNull(!SaldoSoles), 0, !SaldoSoles), "#,###,###,##0.00;-#,###,###,##0.00") Else RSet pSaldoIniDeudor = Format$(0, "#,###,###,##0.00;-#,###,###,##0.00")
           If !SaldoSoles < 0 Then RSet pSaldoInicial = Format$(IIf(IsNull(!SaldoSoles), 0, !SaldoSoles), "#,###,###,##0.00;-#,###,###,##0.00") Else pSaldoInicial = Format$(0, "#,###,###,##0.00;-#,###,###,##0.00")
          Else
           If !SaldoDolares > 0 Then RSet pSaldoIniDeudor = Format$(IIf(IsNull(!SaldoDolares), 0, !SaldoDolares), "#,###,###,##0.00;-#,###,###,##0.00") Else RSet pSaldoIniDeudor = Format$(0, "#,###,###,##0.00;-#,###,###,##0.00")
           If !SaldoDolares < 0 Then RSet pSaldoInicial = Format$(IIf(IsNull(!SaldoDolares), 0, !SaldoDolares), "#,###,###,##0.00;-#,###,###,##0.00") Else pSaldoInicial = Format$(0, "#,###,###,##0.00;-#,###,###,##0.00")
          End If
          
          If CDbl(pSaldoInicial) = 0 Then RSet pSaldoInicial = Format$(pSaldoInicial, "#,###,###,##0.00")
          
          If EntroCabecera = False Then
            If pSaldoIniDeudor < 0 Then VistapSaldoIniDeudor = -(pSaldoIniDeudor) Else VistapSaldoIniDeudor = pSaldoIniDeudor
            If pSaldoInicial < 0 Then VistapSaldoInicial = -(pSaldoInicial): RSet VistapSaldoInicial = Format$(VistapSaldoInicial, "#,###,###,##0.00") Else VistapSaldoInicial = pSaldoInicial
          
            VistapFinalDeudor = CDbl(VistapFinalDeudor) + CDbl(VistapSaldoIniDeudor)
            'VistapFinalAcreedor = CDbl(VistapFinalAcreedor) + CDbl(VistapSaldoInicial)
            
            VistapTotalFinDeudor = CDbl(VistapTotalFinDeudor) + CDbl(VistapSaldoIniDeudor)
            VistapTotalFinAcreedor = CDbl(VistapTotalFinAcreedor) + CDbl(VistapSaldoInicial)
            
             printl (Space(2) & Space(22) & "SALDO INICIAL" & Space(19) & pPla_cCuentaContable & Space(1) & pPla_cNombreCuenta & Space(2) & pAsd_nTipoCambio & pSaldoIniME & VistapSaldoIniDeudor & VistapSaldoInicial)
          End If
          MovMesActual = False
          pSumHaberME = 0: pSumDebe = 0: pSumHaber = 0
          pSaldoAcumME = 0: pSaldoAcumDeudor = 0: pSaldoAcumulado = 0
    
          Do While Not .EOF
            If Trim(!Pla_cCuentaContable) = Trim(pPla_cCuentaContable) Then
                 EntroCabecera = False
                 If Not IsNull(!Asd_dFecDoc) And Not IsNull(!Asd_cGlosa) Then
                    pAse_nVoucher = !Ase_nVoucher
                    pAsd_dFecDoc = !Asd_dFecDoc
                    pAsd_cGlosa = !Asd_cGlosa
                    
                    If Trim(!MNac) = "1" Then RSet pDebe = Format$(IIf(IsNull(!Asd_nDebeSoles), 0, !Asd_nDebeSoles), "#,###,###,##0.00;-#,###,###,##0.00") Else RSet pDebe = Format$(IIf(IsNull(!Asd_nDebeMonExt), 0, !Asd_nDebeMonExt), "#,###,###,##0.00;-#,###,###,##0.00")
                    If Trim(!MNac) = "1" Then RSet pHaber = Format$(IIf(IsNull(!Asd_nHaberSoles), 0, !Asd_nHaberSoles), "#,###,###,##0.00;-#,###,###,##0.00") Else RSet pHaber = Format$(IIf(IsNull(!Asd_nHaberMonExt), 0, !Asd_nHaberMonExt), "#,###,###,##0.00;-#,###,###,##0.00")
                    If Trim(!MNac <> "1") Then RSet pHaberME = Format$(IIf(IsNull(!Asd_nDebeSoles), 0, !Asd_nDebeSoles) - IIf(IsNull(!Asd_nHaberSoles), 0, !Asd_nHaberSoles), "#,###,###,##0.00;-#,###,###,##0.00") Else RSet pHaberME = Format$(IIf(IsNull(!Asd_nDebeMonExt), 0, !Asd_nDebeMonExt) - IIf(IsNull(!Asd_nHaberMonExt), 0, !Asd_nHaberMonExt), "#,###,###,##0.00;-#,###,###,##0.00")
                    
                    pVanDebe = CDbl(pVanDebe) + CDbl(pDebe): RSet pVanDebe = Format$(pVanDebe, "#,###,###,##0.00;-#,###,###,##0.00")
                    pVanHaber = CDbl(pVanHaber) + CDbl(pHaber): RSet pVanHaber = Format$(pVanHaber, "#,###,###,##0.00;-#,###,###,##0.00")
                    pVanHaberME = CDbl(pVanHaberME) + CDbl(pHaberME): RSet pVanHaberME = Format$(pVanHaberME, "#,###,###,##0.00;-#,###,###,##0.00")
                    
                    RSet pSumHaberME = Format$(CDbl(pSumHaberME) + CDbl(pHaberME), "#,###,###,##0.00;-#,###,###,##0.00")
                    RSet pSumDebe = Format$(CDbl(pSumDebe) + CDbl(pDebe), "#,###,###,##0.00;-#,###,###,##0.00")
                    RSet pSumHaber = Format$(CDbl(pSumHaber) + CDbl(pHaber), "#,###,###,##0.00;-#,###,###,##0.00")
                    
                    RSet pAsd_nTipoCambio = Format$(IIf(IsNull(!Asd_nTipoCambio), 0, !Asd_nTipoCambio), "#,###0.000;-#,###0.000")
                                       
                    gsLinea = (pAse_nVoucher & pAsd_dFecDoc & Space(1) & pAsd_cGlosa & Space(1) & pPla_cCuentaContable & Space(1) & pPla_cNombreCuenta & Space(2) & pAsd_nTipoCambio & pHaberME & pDebe & pHaber)
                    printl Space(2) & gsLinea
                    PlanContable = pPla_cCuentaContable
                    MovMesActual = True

                    If giLineas = 74 Then
                        ImprimeVanVienen
                    End If
                    
                    VistapFinalDeudor = CDbl(VistapFinalDeudor) + CDbl(pDebe)
                    'VistapFinalAcreedor = CDbl(VistapFinalAcreedor) + CDbl(pHaber)
                Else
                 pPla_cCuentaContable = ""
                 Cont = Cont + 1
                 .MoveNext
                 If .EOF Then Exit Do
                 GoTo a:
                End If
           Else
             Exit Do
           End If
                .MoveNext
          Loop
                    
a:          If Not .EOF Then
            RSet pSaldoAcumME = Format$(CDbl(pSaldoIniME) + CDbl(pSumHaberME), "#,###,###,##0.00;-#,###,###,##0.00")
            
            If ((CDbl(pSaldoIniDeudor) + CDbl(pSaldoInicial) + CDbl(pSumDebe)) - CDbl(pSumHaber)) > 0 Then
             RSet pSaldoAcumDeudor = Format$((CDbl(pSaldoIniDeudor) + CDbl(pSaldoInicial) + CDbl(pSumDebe)) - CDbl(pSumHaber), "#,###,###,##0.00;-#,###,###,##0.00")
            Else
             RSet pSaldoAcumDeudor = Format$(0, "#,###,###,##0.00;-#,###,###,##0.00")
            End If
            
            If ((CDbl(pSaldoIniDeudor) + CDbl(pSaldoInicial) + CDbl(pSumDebe)) - CDbl(pSumHaber)) < 0 Then
                pSaldoAcumulado = (CDbl(pSaldoIniDeudor) + CDbl(pSaldoInicial) + CDbl(pSumDebe)) - CDbl(pSumHaber)
            Else
             RSet pSaldoAcumulado = Format$(0, "#,###,###,##0.00;-#,###,###,##0.00")
            End If
            
            pPla_cCuentaContable = ""
            
            If pSumDebe < 0 Then VistapSumDebe = -(pSumDebe) Else VistapSumDebe = pSumDebe
            If pSumHaber < 0 Then VistapSumHaber = -(pSumHaber) Else VistapSumHaber = pSumHaber

               If (pSaldoAcumDeudor) < 0 Then
                    VistapSaldoAcumulado = -(pSaldoAcumDeudor)
                    RSet VistapSaldoAcumulado = Format$(CDbl(VistapSaldoAcumulado), "#,###,###,##0.00;-#,###,###,##0.00")
                Else
                    RSet VistapSaldoAcumulado = Format$(CDbl(pSaldoAcumDeudor), "#,###,###,##0.00;-#,###,###,##0.00")
                End If
                If CDbl(pSaldoAcumulado) < 0 Then
                    VistapSaldoAcumDeudor = -(pSaldoAcumulado)
                    RSet VistapSaldoAcumDeudor = Format$(CDbl(VistapSaldoAcumDeudor), "#,###,###,##0.00;-#,###,###,##0.00")
                Else
                    RSet VistapSaldoAcumDeudor = Format$(CDbl(pSaldoAcumulado), "#,###,###,##0.00;-#,###,###,##0.00")
                End If

            VistapTotalFinDeudor = CDbl(VistapTotalFinDeudor) + CDbl(VistapSumDebe)
            VistapTotalFinAcreedor = CDbl(VistapTotalFinAcreedor) + CDbl(VistapSumHaber)
            If giLineas = 74 Then ImprimeVanVienen
            printl (Space(2) & Space(98) & "-------------- --------------- ----------")
            If giLineas = 74 Then ImprimeVanVienen
            If MovMesActual Then printl (Space(2) & Space(22) & "TOTAL MOVIMIENTO CTA." & PlanContable & Space(37) & pSumHaberME & VistapSumDebe & VistapSumHaber)
            If giLineas = 74 Then ImprimeVanVienen
            printl (Space(2) & Space(22) & "SALDO FINAL" & Space(62) & pSaldoAcumME & VistapSaldoAcumDeudor & VistapSaldoAcumulado)
            If giLineas = 74 Then ImprimeVanVienen
'            printl ("")
'            If giLineas = 74 Then ImprimeVanVienen
            RSet pSaldoTotalEfectivo = Format$(CDbl(pSaldoTotalEfectivo) + CDbl(pSaldoAcumDeudor) + CDbl(pSaldoAcumulado), "#,###,###,##0.00;-#,###,###,##0.00")
            
            pTotalHaberME = CDbl(pTotalHaberME) + CDbl(pSumHaberME): RSet pTotalHaberME = Format$(pTotalHaberME, "#,###,###,##0.00;-#,###,###,##0.00")
            pTotalDebe = CDbl(pTotalDebe) + CDbl(pSumDebe): RSet pTotalDebe = Format$(pTotalDebe, "#,###,###,##0.00;-#,###,###,##0.00")
            pTotalHaber = CDbl(pTotalHaber) + CDbl(pSumHaber): RSet pTotalHaber = Format$(pTotalHaber, "#,###,###,##0.00;-#,###,###,##0.00")
            
            pSldoTotalHaberME = CDbl(pSldoTotalHaberME) + CDbl(pSaldoAcumME)
            pSldoTotalDebe = CDbl(pSldoTotalDebe) + CDbl(pSaldoAcumDeudor)
            pSldoTotalHaber = CDbl(pSldoTotalHaber) + CDbl(pSaldoAcumulado)
                                                            
          ElseIf .EOF And gsCtaIni = "" And gsCtaFin = "" Then 'Imprime Total General
            RSet pSaldoAcumME = Format$(CDbl(pSaldoIniME) + CDbl(pSumHaberME), "#,###,###,##0.00;-#,###,###,##0.00")
            
            If ((CDbl(pSaldoIniDeudor) + CDbl(pSaldoInicial) + CDbl(pSumDebe)) - CDbl(pSumHaber)) > 0 Then
             RSet pSaldoAcumDeudor = Format$((CDbl(pSaldoIniDeudor) + CDbl(pSaldoInicial) + CDbl(pSumDebe)) - CDbl(pSumHaber), "#,###,###,##0.00;-#,###,###,##0.00")
            Else
             RSet pSaldoAcumDeudor = Format$(0, "#,###,###,##0.00;-#,###,###,##0.00")
            End If
            
            If ((CDbl(pSaldoIniDeudor) + CDbl(pSaldoInicial) + CDbl(pSumDebe)) - CDbl(pSumHaber)) < 0 Then
                pSaldoAcumulado = (CDbl(pSaldoIniDeudor) + CDbl(pSaldoInicial) + CDbl(pSumDebe)) - CDbl(pSumHaber)
            Else
             RSet pSaldoAcumulado = Format$(0, "#,###,###,##0.00;-#,###,###,##0.00")
            End If
            
            pPla_cCuentaContable = ""
            
            If pSumDebe < 0 Then VistapSumDebe = -(pSumDebe) Else VistapSumDebe = pSumDebe
            If pSumHaber < 0 Then VistapSumHaber = -(pSumHaber) Else VistapSumHaber = pSumHaber
            
               If pSaldoAcumDeudor < 0 Then
                    VistapSaldoAcumulado = -(pSaldoAcumDeudor)
                    RSet VistapSaldoAcumulado = Format$(CDbl(VistapSaldoAcumulado), "#,###,###,##0.00;-#,###,###,##0.00")
                Else
                    RSet VistapSaldoAcumulado = Format$(CDbl(pSaldoAcumDeudor), "#,###,###,##0.00;-#,###,###,##0.00")
                End If
                If pSaldoAcumulado < 0 Then
                    VistapSaldoAcumDeudor = -(pSaldoAcumulado)
                    RSet VistapSaldoAcumDeudor = Format$(CDbl(VistapSaldoAcumDeudor), "#,###,###,##0.00;-#,###,###,##0.00")
                Else
                    RSet VistapSaldoAcumDeudor = Format$(CDbl(pSaldoAcumulado), "#,###,###,##0.00;-#,###,###,##0.00")
                End If
                
            VistapTotalFinDeudor = CDbl(VistapTotalFinDeudor) + CDbl(VistapSumDebe)
            VistapTotalFinAcreedor = CDbl(VistapTotalFinAcreedor) + CDbl(VistapSumHaber)

            printl (Space(2) & Space(98) & "-------------- --------------- ----------")
            If giLineas = 74 Then ImprimeVanVienen
            If .EOF And MovMesActual Then printl (Space(2) & Space(22) & "TOTAL MOVIMIENTO CTA." & PlanContable & Space(37) & pSumHaberME & VistapSumDebe & VistapSumHaber)
            If giLineas = 74 Then ImprimeVanVienen
            printl (Space(2) & Space(22) & "SALDO FINAL" & Space(62) & pSaldoAcumME & VistapSaldoAcumDeudor & VistapSaldoAcumulado)
'            If giLineas = 74 Then ImprimeVanVienen
'            printl ("")
            If giLineas = 74 Then ImprimeVanVienen
            RSet pSaldoTotalEfectivo = Format$(CDbl(pSaldoTotalEfectivo) + CDbl(pSaldoAcumDeudor) + CDbl(pSaldoAcumulado), "#,###,###,##0.00;-#,###,###,##0.00")
            
            pTotalHaberME = CDbl(pTotalHaberME) + CDbl(pSumHaberME): RSet pTotalHaberME = Format$(pTotalHaberME, "#,###,###,##0.00;-#,###,###,##0.00")
            pTotalDebe = CDbl(pTotalDebe) + CDbl(pSumDebe): RSet pTotalDebe = Format$(pTotalDebe, "#,###,###,##0.00;-#,###,###,##0.00")
            pTotalHaber = CDbl(pTotalHaber) + CDbl(pSumHaber): RSet pTotalHaber = Format$(pTotalHaber, "#,###,###,##0.00;-#,###,###,##0.00")
            
            If .EOF And Trim(pPla_cCuentaContable) = "" Then 'And gsPagina <= 1 Then 'El Inicial al no tener Movimientos es el Final
             pSldoTotalHaberME = CDbl(pSldoTotalHaberME) + CDbl(pSaldoAcumME): RSet pSldoTotalHaberME = Format$(pSldoTotalHaberME, "#,###,###,##0.00;-#,###,###,##0.00")

                If CDbl(VistapTotalFinDeudor) > CDbl(VistapTotalFinAcreedor) Then
                       If CDbl(pSaldoTotalEfectivo) < 0 Then
                           pSldoTotalHaber = -(pSaldoTotalEfectivo)
                           RSet pSldoTotalHaber = Format$(CDbl(pSldoTotalHaber), "#,###,###,##0.00")
                       Else
                           RSet pSldoTotalHaber = Format$(CDbl(pSaldoTotalEfectivo), "#,###,###,##0.00")
                       End If
'                       VistapFinalDeudor = CDbl(VistapTotalFinAcreedor) + CDbl(pSaldoTotalEfectivo)
                       RSet pSldoTotalDebe = Format$(0, "#,###,###,##0.00")
                Else
                       If CDbl(pSaldoTotalEfectivo) < 0 Then
                           pSldoTotalDebe = -(pSaldoTotalEfectivo)
                           RSet pSldoTotalDebe = Format$(CDbl(pSldoTotalDebe), "#,###,###,##0.00")
                       Else
                           RSet pSldoTotalDebe = Format$(CDbl(pSaldoTotalEfectivo), "#,###,###,##0.00")
                       End If
'                       VistapTotalFinAcreedor = CDbl(VistapTotalFinAcreedor) + CDbl(pSaldoTotalEfectivo)
                   RSet pSldoTotalHaber = Format$(0, "#,###,###,##0.00")
                End If
            End If
            

          VistapFinalDeudor = CDbl(VistapFinalDeudor) + CDbl(pSldoTotalDebe)
          VistapTotalFinAcreedor = CDbl(VistapTotalFinAcreedor) + CDbl(pSldoTotalHaber)
        RSet VistapTotalFinAcreedor = Format$(VistapTotalFinAcreedor, "#,###,###,##0.00")
        RSet VistapFinalDeudor = Format$(VistapFinalDeudor, "#,###,###,##0.00")
            RSet VistapFinalDeudor = Format$(VistapFinalDeudor, "#,###,###,##0.00")
            RSet VistapTotalFinAcreedor = Format$(VistapTotalFinAcreedor, "#,###,###,##0.00")
            If giLineas = 74 Then ImprimeVanVienen
            printl (Space(2) & Space(22) & "---------------------------------------------------------------------------------------------------------------------")
            If giLineas = 74 Then ImprimeVanVienen
            printl (Space(2) & Space(22) & "TOTAL MOVIMIENTO DEL EFECTIVO" & Space(44) & pTotalHaberME & pTotalDebe & pTotalHaber)
            If giLineas = 74 Then ImprimeVanVienen
            printl (Space(2) & Space(22) & "SALDO TOTAL EFECTIVO" & Space(53) & pSldoTotalHaberME & pSldoTotalDebe & pSldoTotalHaber)
'            If giLineas = 74 Then ImprimeVanVienen
'            printl ("")
            If giLineas = 74 Then ImprimeVanVienen
            printl (Space(2) & Space(98) & "-------------- --------------- ----------")
            printl (Space(2) & Space(22) & "TOTALES" & Space(66) & pSldoTotalHaberME & VistapFinalDeudor & VistapTotalFinAcreedor)
            If giLineas = 74 Then ImprimeVanVienen
          ElseIf .EOF And gsCtaIni <> "" And gsCtaFin <> "" Then 'Imprime Total General y SubTotal
          

'PGBV -- Se descomenta esta linea para imprimir totales

          
            RSet pSaldoAcumME = Format$(CDbl(pSaldoIniME) + CDbl(pSumHaberME), "#,###,###,##0.00;-#,###,###,##0.00")
            pPla_cCuentaContable = ""
            If pSumDebe < 0 Then VistapSumDebe = -(pSumDebe) Else VistapSumDebe = pSumDebe
                If pSumHaber < 0 Then VistapSumHaber = -(pSumHaber) Else VistapSumHaber = pSumHaber

               If (CDbl(pSumDebe) + CDbl(pSaldoAcumDeudor)) > (CDbl(pSumHaber) + CDbl(pSaldoAcumulado)) Then
                   If (pSaldoAcumDeudor) < 0 Then
                        VistapSaldoAcumulado = -(pSaldoAcumDeudor)
                        RSet VistapSaldoAcumulado = Format$(CDbl(VistapSaldoAcumulado), "#,###,###,##0.00;-#,###,###,##0.00")
                    Else
                        RSet VistapSaldoAcumulado = Format$(CDbl(pSaldoAcumDeudor), "#,###,###,##0.00;-#,###,###,##0.00")
                    End If
                    RSet VistapSaldoAcumDeudor = Format$(0, "#,###,###,##0.00")
                Else
                    If CDbl(pSaldoAcumulado) < 0 Then
                        VistapSaldoAcumDeudor = -(pSaldoAcumulado)
                        RSet VistapSaldoAcumDeudor = Format$(CDbl(VistapSaldoAcumDeudor), "#,###,###,##0.00;-#,###,###,##0.00")
                    Else
                        RSet VistapSaldoAcumDeudor = Format$(CDbl(pSaldoAcumulado), "#,###,###,##0.00;-#,###,###,##0.00")
                    End If
                    RSet VistapSaldoAcumulado = Format$(0, "#,###,###,##0.00")
                 End If
            VistapTotalFinAcreedor = CDbl(VistapTotalFinAcreedor) + CDbl(VistapSumHaber)

            printl (Space(98) & "-------------- --------------- ----------")
            If giLineas = 74 Then ImprimeVanVienen
            If MovMesActual Then printl (Space(22) & "TOTAL MOVIMIENTO CTA." & PlanContable & Space(37) & pSumHaberME & VistapSumDebe & VistapSumHaber)
            If giLineas = 74 Then ImprimeVanVienen
            printl (Space(22) & "SALDO FINAL" & Space(62) & pSaldoAcumME & VistapSaldoAcumDeudor & VistapSaldoAcumulado)
            If giLineas = 74 Then ImprimeVanVienen
            printl ("")
            If giLineas = 74 Then ImprimeVanVienen
            RSet pSaldoTotalEfectivo = Format$(CDbl(pSaldoTotalEfectivo) + CDbl(pSaldoAcumDeudor) + CDbl(pSaldoAcumulado), "#,###,###,##0.00;-#,###,###,##0.00")

            pTotalHaberME = CDbl(pTotalHaberME) + CDbl(pSumHaberME): RSet pTotalHaberME = Format$(pTotalHaberME, "#,###,###,##0.00;-#,###,###,##0.00")
            pTotalDebe = CDbl(pTotalDebe) + CDbl(pSumDebe): RSet pTotalDebe = Format$(pTotalDebe, "#,###,###,##0.00;-#,###,###,##0.00")
            pTotalHaber = CDbl(pTotalHaber) + CDbl(pSumHaber): RSet pTotalHaber = Format$(pTotalHaber, "#,###,###,##0.00;-#,###,###,##0.00")

            pSldoTotalHaberME = CDbl(pSldoTotalHaberME) + CDbl(pSaldoAcumME): RSet pSldoTotalHaberME = Format$(pSldoTotalHaberME, "#,###,###,##0.00;-#,###,###,##0.00")
            pSldoTotalDebe = CDbl(pSldoTotalDebe) + CDbl(pSaldoAcumDeudor): RSet pSldoTotalDebe = Format$(pSldoTotalDebe, "#,###,###,##0.00;-#,###,###,##0.00")
            pSldoTotalHaber = CDbl(pSldoTotalHaber) + CDbl(pSaldoAcumulado): RSet pSldoTotalHaber = Format$(pSldoTotalHaber, "#,###,###,##0.00;-#,###,###,##0.00")

            If CDbl(VistapFinalDeudor) > CDbl(VistapTotalFinAcreedor) Then
               If CDbl(pSaldoTotalEfectivo) < 0 Then
                   pSldoTotalHaber = -(pSaldoTotalEfectivo)
                   RSet pSldoTotalHaber = Format$(CDbl(pSldoTotalHaber), "#,###,###,##0.00")
               Else
                   RSet pSldoTotalHaber = Format$(CDbl(pSaldoTotalEfectivo), "#,###,###,##0.00")
               End If
               RSet pSldoTotalDebe = Format$(0, "#,###,###,##0.00")
            Else
               If CDbl(pSaldoTotalEfectivo) < 0 Then
                   pSldoTotalDebe = -(pSaldoTotalEfectivo)
                   RSet pSldoTotalDebe = Format$(CDbl(pSldoTotalDebe), "#,###,###,##0.00")
               Else
                   RSet pSldoTotalDebe = Format$(CDbl(pSaldoTotalEfectivo), "#,###,###,##0.00")
               End If
                RSet pSldoTotalHaber = Format$(0, "#,###,###,##0.00")
            End If

            If giLineas = 74 Then ImprimeVanVienen
            printl (Space(22) & "---------------------------------------------------------------------------------------------------------------------")
            If giLineas = 74 Then ImprimeVanVienen
            printl (Space(22) & "TOTAL MOVIMIENTO DEL EFECTIVO" & Space(43) & pTotalHaberME & Space(2) & pTotalDebe & pTotalHaber)
            If giLineas = 74 Then ImprimeVanVienen
            printl (Space(22) & "SALDO TOTAL EFECTIVO" & Space(52) & pSldoTotalHaberME & Space(2) & pSldoTotalDebe & pSldoTotalHaber)
            If giLineas = 74 Then ImprimeVanVienen
            printl ("")
            If giLineas = 74 Then ImprimeVanVienen
            printl (Space(98) & "-------------- --------------- ----------")
            If giLineas = 74 Then ImprimeVanVienen
            printl (Space(22) & "TOTALES" & Space(66) & pSldoTotalHaberME & VistapFinalDeudor & VistapTotalFinAcreedor)
            If giLineas = 74 Then ImprimeVanVienen

'PGBV FIN  --  Se finaliza con segmento comentado

          End If
          'If Cont = 2 Then Cont = 0: .MoveNext
          
          If .EOF Then Exit Do
       Loop
    End If
End With

If frmFCImpresion.List_Destino.Text = "Archivo" Then
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

Exit Sub

ERROR:
 MsgBox Err.Description, vbCritical, App.Title
End Sub
Private Function ExistenDatos() As Boolean
On Error GoTo Error_cmd

Dim sSql As String

Select Case gsAccionRep
Case 2
    Set rsArreglo = New ADODB.Recordset
    sSql = "spCn_RptFormato0101 '" & gsEmpresa & "','" & gsAnio & "','" & gsMesRep & "','" & gsCtaIni & "','" & gsCtaFin & "','" & gsCodMoneda & "'"
    
    ConectarAdvance
      rsArreglo.Open sSql, gcnSistemaAdv, adOpenDynamic, adLockOptimistic
    Desconectar
Case 3
    Set rsArreglo = New ADODB.Recordset
    sSql = "spCn_RptFormato0102 '" & gsEmpresa & "','" & gsAnio & "','" & gsMesRep & "','" & gsCtaIni & "','" & gsCtaFin & "','" & gsCodMoneda & "'"
    
    ConectarAdvance
      rsArreglo.Open sSql, gcnSistemaAdv, adOpenDynamic, adLockOptimistic
    Desconectar
End Select

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
Public Sub CabeceraDetEfectivo()

EntroCabecera = True

Dim sPag As String
Dim Anio As String
Dim Mes As String
Dim sUSUARIO As String * 10

On Error GoTo ERROR
 Gs_HoraServ = DevuelveHoraServidor
 LSet sUSUARIO = gsUsuario

 If Gs_TamPapel = 39 Then nAncho = 232 Else nAncho = 142
' If Not gsNomTipoImp Then printl ("")
 sPag = Space(4)
  
 gsConTotalPaginas = gsConTotalPaginas + 1
 gsPagina = gsPagina + 1

 RSet sPag = Format(gsPagina + 1, "####")
 giLineas = 0
   
 If gsCodMoneda = gsMonedaNac Then Termino1 = "SALDOS Y MOV M.EXT.   " Else Termino1 = "SALDOS Y MOV M.NAC."
 If gsCodMoneda = gsMonedaNac Then Termino2 = "SALDOS Y MOV. M.NAC.    " Else Termino2 = "SALDOS Y MOV. M.EXT."
 
    printl (Space(2) & "Formato 1.1: LIBRO CAJA Y BANCOS                                                                                      Fecha :  " & Format(FechaServidor, "dd/MM/yyyy"))
    'Call AlinearDosTextos(nAncho - 6, Space(2) & "Formato 1.1: LIBRO CAJA Y BANCOS", "Fecha :  " & Format(FechaServidor, "dd/MM/yyyy"))
    Call AlinearDosTextos(nAncho - 4, Space(2) & "DETALLE DE LOS MOVIMIENTOS DEL EFECTIVO", "")
    Dim xgsPagina As String * 4
    RSet xgsPagina = Format(CStr(gsPagina), "####")
    'Call AlinearDosTextos(nAncho - 13, Space(2) & "EJERCICIO/PERIODO    : " & NombreMes(gsMesRep) & " " & gsAnio, "Pagina: " & xgsPagina)
    printl (Space(2) & "EJERCICIO/PERIODO    : " & NombreMes(gsMesRep) & " " & gsAnio & "                                                                                  Pagina: " & xgsPagina)
    Call AlinearDosTextos(nAncho, Space(2) & "RUC                  : " & gsRUC, "")
    'Call AlinearDosTextos(nAncho - 4, "RAZON SOCIAL         : " & gsEmpresaNom, "Página:       " & Format$(gsPaginaPrincipal, "####") & " de " & Format$(gsPagina, "####"))
    Call AlinearDosTextos(nAncho, Space(2) & "APELLIDOS Y NOMBRES,", "")
    Call AlinearDosTextos(nAncho, Space(2) & "DENOMINACIÓN O", "")
    Call AlinearDosTextos(nAncho, Space(2) & "RAZON SOCIAL         : " & gsEmpresaNom, "")
    Call AlinearDosTextos(nAncho, Space(2) & "MONEDA               : " & IIf(gsCodMoneda = gsMonedaNac, gsNombreMonedaNac, gsNombreMonedaExt), "")
 
    printl ("")
    printl ("")
    printl ("")
    printl (Space(2) & "---------- ---------- ------------------------------- ------------------------------ --------------------------- --------------------------")
    printl (Space(2) & "   NUM.       FECHA         DESCRIPCION DE LA              SUBDIVISIONARIA ASOCIADA       " & Termino1 & "     " & Termino2)
    printl (Space(2) & "  CORREL    OPERACION          OPERACION                 CODIGO      DENOMINACION         T.C.        IMPORTE        DEUDOR      ACREEDOR")
    printl (Space(2) & " COD.OPER.  ")
    printl (Space(2) & "---------- ---------- ------------------------------- ------------------------------ --------------------------- --------------------------")
EntroCabecera = False
 Exit Sub
ERROR:
 MsgBox Err.Description, vbCritical, App.Title
End Sub
Public Sub ReporteLibBancosDetMovCtaCte()
On Error GoTo Control
 
Dim pBan_cCodigo As String * 3
Dim pPla_cCuentaContable As String * 12
Dim pSaldoIniME As String * 13
Dim pSaldoIniDeudor As String * 13
Dim pSaldoIniAcreedor As String * 13
Dim pSaldoInicial As Double
Dim pAse_nVoucher As String * 10
Dim pAse_dFecDoc As String * 10
Dim pTra_cCodigo As String * 3
Dim pAsd_cGlosa As String * 29
Dim pEnt_cPersona As String * 24
Dim pAsd_cNumDoc As String * 12
Dim pAsd_nTipoCambio As String * 5
Dim pHaberME As String * 12
Dim pDebe As String * 12
Dim pHaber As String * 12

Dim pTotalHaberME As String * 12
Dim pTotalDebe As String * 12
Dim pTotalHaber As String * 12
Dim pSaldoAcumME As String * 12
Dim pSaldoAcumDeudor As String * 12
Dim pSaldoAcumAcreedor As String * 12

Dim VistapSaldoAcumDeudor As String * 12
Dim VistapSaldoAcumAcreedor As String * 12

Dim VistapTotalDeudor As String * 12
Dim VistapTotalAcreedor As String * 12
Dim pSaldoTotalEfectivo As String * 15
Dim VistapFinalDeudor As String * 15

RSet VistapTotalDeudor = Format$(0, "#,###,###,##0.00;-#,###,###,##0.00")
RSet VistapTotalAcreedor = Format$(0, "#,###,###,##0.00;-#,###,###,##0.00")
RSet pSaldoTotalEfectivo = Format$(0, "#,###,###,##0.00;-#,###,###,##0.00")
RSet VistapFinalDeudor = Format$(0, "#,###,###,##0.00;-#,###,###,##0.00")

Dim swSldoIni As Boolean
Dim i, NroLinIni, NroLinFin As Integer
Dim Count As Integer

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
'  ' Print #1, Chr(27) & Chr(77); '12cpi
'Print #1, Chr(27) & Chr(51) & Chr(29) 'Entre lineas 29/180

giLineas = 0
giEspacios = 60

If Not rsArreglo.EOF Then iReport = 1
If iReport = 1 Then rsArreglo.Sort = "Pla_cCuentaContable": rsArreglo.MoveFirst

gsConTotalPaginas = 0

With rsArreglo
   If .RecordCount > 0 Then
   SwHoja = False
    iReport = 1
    gsPaginaPrincipal = 1
    Call CabeceraDetMovCtaCte
    
    .MoveFirst
    
    pVanDebe = 0: pVanHaber = 0
    pVienenDebe = 0: pVienenHaber = 0
    pVanHaberME = 0: pVienenHaberME = 0
    
    Do While Not .EOF
     pBan_cCodigo = IIf(IsNull(!Ban_cCodigo), "", !Ban_cCodigo)
      Do While Trim(IIf(IsNull(!Ban_cCodigo), "", !Ban_cCodigo)) = Trim(pBan_cCodigo)
       'EntroCabecera = False
       pPla_cCuentaContable = !Pla_cCuentaContable
'       EntroCabecera = False
       
       pTotalHaberME = 0: pTotalDebe = 0: pTotalHaber = 0
       pSaldoAcumME = 0: pSaldoAcumDeudor = 0: pSaldoAcumAcreedor = 0
       VistapSaldoAcumDeudor = 0: VistapSaldoAcumAcreedor = 0

       Do While Trim(IIf(IsNull(!Ban_cCodigo), "", !Ban_cCodigo)) = Trim(pBan_cCodigo) And Trim(!Pla_cCuentaContable) = Trim(pPla_cCuentaContable)
       
        If Not swSldoIni Then
         If Trim(!MNac) <> "1" Then
          RSet pSaldoIniME = Format$(IIf(IsNull(!SaldoSoles), 0, !SaldoSoles), "#,###,###,##0.00;-#,###,###,##0.00")
         Else
          RSet pSaldoIniME = Format$(IIf(IsNull(!SaldoDolares), 0, !SaldoDolares), "#,###,###,##0.00;-#,###,###,##0.00")
         End If
         If Trim(!MNac) = "1" Then pSaldoInicial = IIf(IsNull(!SaldoSoles), 0, !SaldoSoles) Else pSaldoInicial = IIf(IsNull(!SaldoDolares), 0, !SaldoDolares)
         If pSaldoInicial > 0 Then pSaldoIniDeudor = Abs(pSaldoInicial): RSet pSaldoIniDeudor = Format$(pSaldoIniDeudor, "#,###,###,##0.00;-#,###,###,##0.00") Else pSaldoIniDeudor = 0: RSet pSaldoIniDeudor = Format$(pSaldoIniDeudor, "#,###,###,##0.00;-#,###,###,##0.00")
         If pSaldoInicial < 0 Then pSaldoIniAcreedor = Abs(pSaldoInicial): RSet pSaldoIniAcreedor = Format$(pSaldoIniAcreedor, "#,###,###,##0.00;-#,###,###,##0.00") Else pSaldoIniAcreedor = 0: RSet pSaldoIniAcreedor = Format$(pSaldoIniAcreedor, "#,###,###,##0.00;-#,###,###,##0.00")
         
         If EntroCabecera = False Then
          swSldoIni = True
          pSaldoIniDeudor = Right(pSaldoIniDeudor, Len(pSaldoIniDeudor) - 1)
          pSaldoIniAcreedor = Right(pSaldoIniAcreedor, Len(pSaldoIniAcreedor) - 1)
          
          VistapTotalDeudor = CDbl(VistapTotalDeudor) + CDbl(pSaldoIniDeudor)
          VistapTotalAcreedor = CDbl(VistapTotalAcreedor) + CDbl(pSaldoIniAcreedor)
          
          printl (Space(2) & Space(26) & "SALDO INICIAL" & Space(61) & pSaldoIniME & Space(1) & pSaldoIniDeudor & pSaldoIniAcreedor)
         End If
        End If
        
        pAse_nVoucher = IIf(IsNull(!Ase_nVoucher), "", !Ase_nVoucher)
        pAse_dFecDoc = Format$(!Asd_dFecDoc, "dd/mm/yyyy")
        pTra_cCodigo = IIf(IsNull(!Tra_cCodigo), "", !Tra_cCodigo)
        If Trim(IsNull(!Asd_cGlosa)) Then
            pAsd_cGlosa = ""
        Else
            pAsd_cGlosa = !Asd_cGlosa
        End If
        pEnt_cPersona = IIf(IsNull(!Ent_cPersona), "", !Ent_cPersona)
        pAsd_cNumDoc = IIf(IsNull(!Asd_cNumDoc), "", !Asd_cNumDoc)
        pAsd_nTipoCambio = Format$(!Asd_nTipoCambio, "#,###0.000;-#,###0.000")
        
        If Trim(!MNac) <> "1" Then RSet pHaberME = Format$((!Asd_nDebeSoles - !Asd_nHaberSoles), "#,###,###,##0.00;-#,###,###,##0.00") Else RSet pHaberME = Format$((!Asd_nDebeMonExt - !Asd_nHaberMonExt), "#,###,###,##0.00;-#,###,###,##0.00")
        If Trim(!MNac) = "1" Then RSet pDebe = Format$(!Asd_nDebeSoles, "#,###,###,##0.00;-#,###,###,##0.00") Else RSet pDebe = Format$(!Asd_nDebeMonExt, "#,###,###,##0.00;-#,###,###,##0.00")
        If Trim(!MNac) = "1" Then RSet pHaber = Format$(!Asd_nHaberSoles, "#,###,###,##0.00;-#,###,###,##0.00") Else RSet pHaber = Format$(!Asd_nHaberMonExt, "#,###,###,##0.00;-#,###,###,##0.00")
        
        pVanDebe = CDbl(pVanDebe) + CDbl(pDebe)
        RSet pVanDebe = Format$(pVanDebe, "#,###,###,##0.00;-#,###,###,##0.00")
        pVanHaber = CDbl(pVanHaber) + CDbl(pHaber)
        RSet pVanHaber = Format$(pVanHaber, "#,###,###,##0.00;-#,###,###,##0.00")
        
        Dim cSpace As Integer
        If InStr(1, pHaberME, "(") Then
            pHaberME = Format$(CDbl(pHaberME), "#,###,###,##0.00;-#,###,###,##0.00")
            cSpace = Len(pHaberME) - Len(Replace(pHaberME, " ", ""))
            pHaberME = Space(cSpace) + pHaberME
        End If
        
        pVanHaberME = CDbl(pVanHaberME) + CDbl(pHaberME)
        RSet pVanHaberME = Format$(pVanHaberME, "#,###,###,##0.00;-#,###,###,##0.00")
        
        pTotalHaberME = CDbl(pTotalHaberME) + CDbl(pHaberME)
        RSet pTotalHaberME = Format$(pTotalHaberME, "#,###,###,##0.00;-#,###,###,##0.00")
        
       
        pTotalDebe = CDbl(pTotalDebe) + CDbl(pDebe)
        RSet pTotalDebe = Format$(pTotalDebe, "#,###,###,##0.00;-#,###,###,##0.00")
        
        pTotalHaber = CDbl(pTotalHaber) + CDbl(pHaber)
        RSet pTotalHaber = Format$(pTotalHaber, "#,###,###,##0.00;-#,###,###,##0.00-")
                       
        RSet pSaldoAcumME = Format$(CDbl(pSaldoIniME) + CDbl(pTotalHaberME), "#,###,###,##0.00;-#,###,###,##0.00")
        If (CDbl(pSaldoInicial) + (CDbl(pTotalDebe) - CDbl(pTotalHaber))) > 0 Then
         pSaldoAcumDeudor = Abs((CDbl(pSaldoInicial) + (CDbl(pTotalDebe) - CDbl(pTotalHaber)))): RSet pSaldoAcumDeudor = Format$(pSaldoAcumDeudor, "#,###,###,##0.00;-#,###,###,##0.00")
        Else
         pSaldoAcumDeudor = 0: RSet pSaldoAcumDeudor = Format$(pSaldoAcumDeudor, "#,###,###,##0.00;-#,###,###,##0.00")
        End If
        If (CDbl(pSaldoInicial) + (CDbl(pTotalDebe) - CDbl(pTotalHaber))) < 0 Then
         pSaldoAcumAcreedor = Abs((CDbl(pSaldoInicial) + (CDbl(pTotalDebe) - CDbl(pTotalHaber)))): RSet pSaldoAcumAcreedor = Format$(pSaldoAcumAcreedor, "#,###,###,##0.00;-#,###,###,##0.00")
        Else
         pSaldoAcumAcreedor = 0: RSet pSaldoAcumAcreedor = Format$(pSaldoAcumAcreedor, "#,###,###,##0.00;-#,###,###,##0.00")
        End If
        
        gsLinea = (pAse_nVoucher & Space(1) & pAse_dFecDoc & Space(1) & pTra_cCodigo & Space(1) & pAsd_cGlosa & Space(1) & pEnt_cPersona & Space(1) & pAsd_cNumDoc & Space(2) & pAsd_nTipoCambio & Space(1) & pHaberME & Space(1) & pDebe & Space(1) & pHaber)
        printl Space(2) & gsLinea
'_________________inicio esta linea es para el tema de replica activarla si es necesaria
'        If giLineas = 68 And .EOF = False Then
'            'ImprimeVanVienen
'            giLineas = 74
'        End If
'_________________fin esta linea es para el tema de replica activarla si es necesaria

        If giLineas = 74 Then
            ImprimeVanVienen
        End If
        .MoveNext
        If .EOF Then
     
        printl (Space(2) & Space(99) & "--------------  ------------ ------------")
        If CDbl(pSaldoAcumDeudor) < 0 Then
             VistapSaldoAcumAcreedor = -(pSaldoAcumDeudor)
             RSet VistapSaldoAcumAcreedor = Format$(CDbl(VistapSaldoAcumAcreedor), "#,###,###,##0.00;-#,###,###,##0.00")
         Else
             RSet VistapSaldoAcumAcreedor = Format$(CDbl(pSaldoAcumDeudor), "#,###,###,##0.00;-#,###,###,##0.00")
         End If
         If CDbl(pSaldoAcumAcreedor) < 0 Then
             VistapSaldoAcumDeudor = -(pSaldoAcumAcreedor)
             RSet VistapSaldoAcumDeudor = Format$(CDbl(VistapSaldoAcumDeudor), "#,###,###,##0.00;-#,###,###,##0.00")
         Else
             RSet VistapSaldoAcumDeudor = Format$(CDbl(pSaldoAcumAcreedor), "#,###,###,##0.00;-#,###,###,##0.00")
         End If
          VistapTotalDeudor = CDbl(VistapTotalDeudor) + CDbl(pTotalDebe)
          VistapTotalAcreedor = CDbl(VistapTotalAcreedor) + CDbl(pTotalHaber)
          
        'printl (Space(82) & "TOTAL" & Space(11) & pTotalHaberME & Space(2) & pTotalDebe & Space(1) & pTotalHaber)
        printl (Space(2) & Space(29) & "SUMAS DEL PERIODO" & Space(55) & pTotalHaberME & Space(1) & pTotalDebe & Space(1) & pTotalHaber)
        printl (Space(2) & Space(29) & "SALDO FINAL" & Space(61) & pSaldoAcumME & Space(1) & VistapSaldoAcumDeudor & Space(1) & VistapSaldoAcumAcreedor)
        printl (Space(2) & Space(99) & "")
        printl (Space(2) & Space(99) & "--------------  ------------ ------------")
        
        If CDbl(VistapTotalDeudor) > CDbl(VistapTotalAcreedor) Then
               VistapTotalAcreedor = CDbl(VistapTotalAcreedor) + CDbl(VistapSaldoAcumAcreedor)
        Else
               VistapTotalDeudor = CDbl(VistapTotalDeudor) + CDbl(VistapSaldoAcumDeudor)
        End If
        RSet VistapTotalAcreedor = Format$(VistapTotalAcreedor, "#,###,###,##0.00")
        RSet VistapTotalDeudor = Format$(VistapTotalDeudor, "#,###,###,##0.00")
        
         printl (Space(2) & Space(29) & "TOTALES" & Space(65) & pSaldoAcumME & Space(1) & VistapTotalDeudor & Space(1) & VistapTotalAcreedor)
         
        swSldoIni = False

         Exit Do
        End If
       Loop
       
       If .EOF Then Exit Do
       If Trim(!Ban_cCodigo) = Trim(pBan_cCodigo) And Trim(!Pla_cCuentaContable) <> Trim(pPla_cCuentaContable) Then 'Fin de Cuenta Contable - Calculos
        printl (Space(2) & Space(99) & "--------------  ------------ ------------")
       
        If giLineas = 74 Then ImprimeVanVienen
        If CDbl(pSaldoAcumDeudor) < 0 Then
             VistapSaldoAcumAcreedor = -(pSaldoAcumDeudor)
             RSet VistapSaldoAcumAcreedor = Format$(CDbl(VistapSaldoAcumAcreedor), "#,###,###,##0.00;-#,###,###,##0.00")
         Else
             RSet VistapSaldoAcumAcreedor = Format$(CDbl(pSaldoAcumDeudor), "#,###,###,##0.00;-#,###,###,##0.00")
         End If
         If CDbl(pSaldoAcumAcreedor) < 0 Then
             VistapSaldoAcumDeudor = -(pSaldoAcumAcreedor)
             RSet VistapSaldoAcumDeudor = Format$(CDbl(VistapSaldoAcumDeudor), "#,###,###,##0.00;-#,###,###,##0.00")
         Else
             RSet VistapSaldoAcumDeudor = Format$(CDbl(pSaldoAcumAcreedor), "#,###,###,##0.00;-#,###,###,##0.00")
         End If
          VistapTotalDeudor = CDbl(VistapTotalDeudor) + CDbl(pTotalDebe)
          VistapTotalAcreedor = CDbl(VistapTotalAcreedor) + CDbl(pTotalHaber)
        printl (Space(2) & Space(29) & "SUMAS DEL PERIODO" & Space(55) & pTotalHaberME & Space(1) & pTotalDebe & Space(1) & pTotalHaber)
        If giLineas = 74 Then ImprimeVanVienen
        printl (Space(2) & Space(29) & "SALDO FINAL" & Space(61) & pSaldoAcumME & Space(1) & VistapSaldoAcumDeudor & Space(1) & VistapSaldoAcumAcreedor)
        If giLineas = 74 Then ImprimeVanVienen
        printl (Space(2) & Space(99) & "")
        If giLineas = 74 Then ImprimeVanVienen
        printl (Space(2) & Space(99) & "--------------  ------------ ------------")
        If giLineas = 74 Then ImprimeVanVienen
        
        If CDbl(VistapTotalDeudor) > CDbl(VistapTotalAcreedor) Then
               VistapTotalAcreedor = CDbl(VistapTotalAcreedor) + CDbl(VistapSaldoAcumAcreedor)
        Else
               VistapTotalDeudor = CDbl(VistapTotalDeudor) + CDbl(VistapSaldoAcumDeudor)
        End If
        RSet VistapTotalAcreedor = Format$(VistapTotalAcreedor, "#,###,###,##0.00")
        RSet VistapTotalDeudor = Format$(VistapTotalDeudor, "#,###,###,##0.00")
        
        If giLineas = 74 Then ImprimeVanVienen
         printl (Space(2) & Space(29) & "TOTALES" & Space(65) & pSaldoAcumME & Space(1) & VistapTotalDeudor & Space(1) & VistapTotalAcreedor)
        If giLineas = 74 Then
        ImprimeVanVienen
        'giLineas = 72
        End If
        Call CabeceraDetMovCtaCte
        swSldoIni = False
       ElseIf Trim(IIf(IsNull(!Ban_cCodigo), "", !Ban_cCodigo)) <> Trim(pBan_cCodigo) And Trim(!Pla_cCuentaContable) <> Trim(pPla_cCuentaContable) Then 'Cambio de Cta.Cte y Cuenta Contable
        
        printl (Space(2) & Space(99) & "--------------  ------------ ------------")
        If giLineas = 74 Then ImprimeVanVienen
        If CDbl(pSaldoAcumDeudor) < 0 Then
             VistapSaldoAcumAcreedor = -(pSaldoAcumDeudor)
             RSet VistapSaldoAcumAcreedor = Format$(CDbl(VistapSaldoAcumAcreedor), "#,###,###,##0.00;-#,###,###,##0.00")
         Else
             RSet VistapSaldoAcumAcreedor = Format$(CDbl(pSaldoAcumDeudor), "#,###,###,##0.00;-#,###,###,##0.00")
         End If
         If CDbl(pSaldoAcumAcreedor) < 0 Then
             VistapSaldoAcumDeudor = -(pSaldoAcumAcreedor)
             RSet VistapSaldoAcumDeudor = Format$(CDbl(VistapSaldoAcumDeudor), "#,###,###,##0.00;-#,###,###,##0.00")
         Else
             RSet VistapSaldoAcumDeudor = Format$(CDbl(pSaldoAcumAcreedor), "#,###,###,##0.00;-#,###,###,##0.00")
         End If
          VistapTotalDeudor = CDbl(VistapTotalDeudor) + CDbl(pTotalDebe)
          VistapTotalAcreedor = CDbl(VistapTotalAcreedor) + CDbl(pTotalHaber)
        printl (Space(2) & Space(29) & "SUMAS DEL PERIODO" & Space(55) & pTotalHaberME & Space(1) & pTotalDebe & Space(1) & pTotalHaber)
        If giLineas = 74 Then ImprimeVanVienen
        printl (Space(2) & Space(29) & "SALDO FINAL" & Space(61) & pSaldoAcumME & Space(1) & VistapSaldoAcumDeudor & Space(1) & VistapSaldoAcumAcreedor)
        If giLineas = 74 Then ImprimeVanVienen
        printl (Space(2) & Space(99) & "")
        If giLineas = 74 Then ImprimeVanVienen
        printl (Space(2) & Space(99) & "--------------  ------------ ------------")
        
        If CDbl(VistapTotalDeudor) > CDbl(VistapTotalAcreedor) Then
               VistapTotalAcreedor = CDbl(VistapTotalAcreedor) + CDbl(VistapSaldoAcumAcreedor)
        Else
               VistapTotalDeudor = CDbl(VistapTotalDeudor) + CDbl(VistapSaldoAcumDeudor)
        End If
        RSet VistapTotalAcreedor = Format$(VistapTotalAcreedor, "#,###,###,##0.00")
        RSet VistapTotalDeudor = Format$(VistapTotalDeudor, "#,###,###,##0.00")
        If giLineas = 74 Then ImprimeVanVienen
         printl (Space(2) & Space(29) & "TOTALES" & Space(65) & pSaldoAcumME & Space(1) & VistapTotalDeudor & Space(1) & VistapTotalAcreedor)
        If giLineas = 74 Then ImprimeVanVienen
        Call CabeceraDetMovCtaCte
        swSldoIni = False
       ElseIf IsNull(!Ban_cCodigo) And Trim(!Pla_cCuentaContable) <> Trim(pPla_cCuentaContable) Then 'Cambio de Cta.Cte y Cuenta Contable
        printl (Space(2) & Space(99) & "--------------  ------------ ------------")
        If giLineas = 74 Then ImprimeVanVienen
        If CDbl(pSaldoAcumDeudor) < 0 Then
             VistapSaldoAcumAcreedor = -(pSaldoAcumDeudor)
             RSet VistapSaldoAcumAcreedor = Format$(CDbl(VistapSaldoAcumAcreedor), "#,###,###,##0.00;-#,###,###,##0.00")
         Else
             RSet VistapSaldoAcumAcreedor = Format$(CDbl(pSaldoAcumDeudor), "#,###,###,##0.00;-#,###,###,##0.00")
         End If
         If CDbl(pSaldoAcumAcreedor) < 0 Then
             VistapSaldoAcumDeudor = -(pSaldoAcumAcreedor)
             RSet VistapSaldoAcumDeudor = Format$(CDbl(VistapSaldoAcumDeudor), "#,###,###,##0.00;-#,###,###,##0.00")
         Else
             RSet VistapSaldoAcumDeudor = Format$(CDbl(pSaldoAcumAcreedor), "#,###,###,##0.00;-#,###,###,##0.00")
         End If
        
          VistapTotalDeudor = CDbl(VistapTotalDeudor) + CDbl(pTotalDebe)
          VistapTotalAcreedor = CDbl(VistapTotalAcreedor) + CDbl(pTotalHaber)
        printl (Space(2) & Space(29) & "SUMAS DEL PERIODO" & Space(55) & pTotalHaberME & Space(1) & pTotalDebe & Space(1) & pTotalHaber)
        If giLineas = 74 Then ImprimeVanVienen
        printl (Space(2) & Space(29) & "SALDO FINAL" & Space(61) & pSaldoAcumME & Space(1) & VistapSaldoAcumDeudor & Space(1) & VistapSaldoAcumAcreedor)
        If giLineas = 74 Then ImprimeVanVienen
        printl (Space(2) & Space(99) & "")
        If giLineas = 74 Then ImprimeVanVienen
        printl (Space(2) & Space(99) & "--------------  ------------ ------------")
        If giLineas = 74 Then ImprimeVanVienen
         VistapTotalAcreedor = Format$(VistapTotalAcreedor, "#,###,###,##0.00")
         VistapTotalDeudor = Format$(VistapTotalDeudor, "#,###,###,##0.00")
        
        If CDbl(VistapTotalDeudor) > CDbl(VistapTotalAcreedor) Then
               VistapTotalAcreedor = CDbl(VistapTotalAcreedor) + CDbl(VistapSaldoAcumAcreedor)
        Else
               VistapTotalDeudor = CDbl(VistapTotalDeudor) + CDbl(VistapSaldoAcumDeudor)
        End If
        
        RSet VistapTotalAcreedor = Format$(VistapTotalAcreedor, "#,###,###,##0.00")
        RSet VistapTotalDeudor = Format$(VistapTotalDeudor, "#,###,###,##0.00")
        printl (Space(2) & Space(29) & "TOTALES" & Space(65) & pSaldoAcumME & Space(1) & VistapTotalDeudor & Space(1) & VistapTotalAcreedor)
        If giLineas = 74 Then ImprimeVanVienen
        Call CabeceraDetMovCtaCte
        swSldoIni = False
        printl ""
        
        RSet VistapTotalAcreedor = Format$(0, "#,###,###,##0.00")
        RSet VistapTotalDeudor = Format$(0, "#,###,###,##0.00")
       End If
      Loop
    Loop
    
   End If
End With

If frmFCImpresion.List_Destino.Text = "Archivo" Then
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
 
Exit Sub
Control:
 Screen.MousePointer = vbDefault
 MsgBox Err.Description
 Resume
End Sub
Public Sub CabeceraDetMovCtaCte()
On Error GoTo Control
EntroCabecera = True

Dim sPag As String
Dim Anio As String
Dim Mes As String
Dim sUSUARIO As String * 10

 Gs_HoraServ = DevuelveHoraServidor
 LSet sUSUARIO = gsUsuario

 If Gs_TamPapel = 39 Then nAncho = 232 Else nAncho = 142
' If Not gsNomTipoImp Then printl ("")
 sPag = Space(3)
  
 gsConTotalPaginas = gsConTotalPaginas + 1
 gsPagina = gsPagina + 1

 RSet sPag = Format(gsPagina + 1, "####")
 giLineas = 0
  
 If gsCodMoneda = gsMonedaNac Then Termino1 = "SALDOS Y MOV.M.EXT.   " Else Termino1 = "SALDOS Y MOV.M.NAC."
 If gsCodMoneda = gsMonedaNac Then Termino2 = "SALDOS Y MOV.M.NAC.    " Else Termino2 = "SALDOS Y MOV.M.EXT."
 
    printl (Space(2) & "Formato 1.2: LIBRO CAJA Y BANCOS                                                                                      Fecha :  " & Format(FechaServidor, "dd/MM/yyyy"))
    'Call AlinearDosTextos(nAncho - 6, Space(2) & "Formato 1.2: LIBRO CAJA Y BANCOS", "Fecha :  " & Format(FechaServidor, "dd/MM/yyyy"))
    Call AlinearDosTextos(nAncho - 4, Space(2) & "DETALLE DE LOS MOVIMIENTOS DE LA CUENTA CORRIENTE", "")
    Dim xgsPagina As String * 4
    RSet xgsPagina = Format(CStr(gsPagina), "####")
    'Call AlinearDosTextos(nAncho - 13, Space(2) & "EJERCICIO/PERIODO    : " & NombreMes(gsMesRep) & " " & gsAnio, "Pagina: " & xgsPagina)
    printl (Space(2) & "EJERCICIO/PERIODO    : " & NombreMes(gsMesRep) & " " & gsAnio & "                                                                                  Pagina: " & xgsPagina)
    Call AlinearDosTextos(nAncho, Space(2) & "RUC                  : " & gsRUC, "")
    'Call AlinearDosTextos(nAncho - 4, "RAZON SOCIAL         : " & gsEmpresaNom, "Página:       " & Format$(gsPaginaPrincipal, "####") & " de " & Format$(gsPagina, "####"))
   
    Call AlinearDosTextos(nAncho - 13, Space(2) & "APELLIDOS Y NOMBRES,", "")
    Call AlinearDosTextos(nAncho - 13, Space(2) & "DENOMINACIÓN O", "")
    Call AlinearDosTextos(nAncho - 13, Space(2) & "RAZON SOCIAL         : " & gsEmpresaNom, "")
    
    Call AlinearDosTextos(nAncho, Space(2) & "MONEDA               : " & IIf(gsCodMoneda = gsMonedaNac, gsNombreMonedaNac, gsNombreMonedaExt), "")
    
    printl ("")
    
    With rsArreglo
     Call AlinearDosTextos(nAncho, Space(2) & "ENTIDAD FINANCIERA   : " & !Ban_cNombre, "")
     Call AlinearDosTextos(nAncho, Space(2) & "CODIGO CTA CORRIENTE : " & !Cue_cNumCuenta & "   " & "MONEDA DE ORIGEN : " & !DesMonOrigen, "")
     Call AlinearDosTextos(nAncho, Space(2) & "CUENTA CONTABLE ASOC.: " & !Pla_cCuentaContable & "   " & !Pla_cNombreCuenta, "")
    End With
    
    printl ("")
    printl (Space(2) & "----------- --------- ------------------------------------------------------------------------ ----------------- -------------------------")
    printl (Space(2) & "   NUM.       FECHA                            OPERACIONES BANCARIAS                           " & Termino1 & "  " & Termino2)
    printl (Space(2) & "  CORREL    OPERACION MEDIO DESCRIPCION DE LA OPERACION  APELLIDOS Y NOMBRES,  NRO.TRANS BANC. T.C.       IMPORTE       DEUDOR     ACREEDOR")
    printl (Space(2) & " COD.OPER.            PAGO                               DENOM.O RAZON SOCIAL     O DOC.     ")
    printl (Space(2) & "----------- ---------- ------------------------------------------------------------------------ -----------------  ------------------------")
 EntroCabecera = False
Exit Sub
Control:
 MsgBox Err.Description
End Sub
Sub ImprimeVanVienen()
Dim pVanDebeCta As String * 12
Dim pVienenDebeCta As String * 12
On Error GoTo Control
 Select Case giLineas
 Case 74
  printl ""
  RSet pVanHaberME = Format$(pVanHaberME, "#,###,##0.00;-#,###,##0.00")
  RSet pVanDebe = Format$(pVanDebe, "#,###,##0.00;-#,###,##0.00")
  RSet pVanHaber = Format$(pVanHaber, "#,###,##0.00;-#,###,##0.00")
    RSet pVanDebeCta = Format$(pVanDebe, "#,###,##0.00;-#,###,##0.00")
  SwHoja = True
  'MsgBox Str(Len(pVanHaberME)) + " " + Str(Len(pVanDebe)) + " " + Str(Len(pVanHaber))
  If gsAccionRep = 2 Then printl (Space(2) & Space(63) & "VAN..." & Space(29) & pVanHaberME & pVanDebe & pVanHaber)
 
  If gsAccionRep = 3 Then printl (Space(2) & Space(66) & "VAN..." & Space(29) & pVanHaberME & Space(1) & pVanDebeCta & pVanHaber)
  
  pVienenDebe = pVanDebe
  pVienenHaber = pVanHaber
  pVienenHaberME = pVanHaberME

  RSet pVienenDebe = Format$(pVienenDebe, "#,###,##0.00;-#,###,##0.00")
  RSet pVienenHaber = Format$(pVienenHaber, "#,###,##0.00;-#,###,##0.00")
  RSet pVienenHaberME = Format$(pVienenHaberME, "#,###,##0.00;-#,###,##0.00")
  RSet pVienenDebeCta = Format$(pVienenDebe, "#,###,##0.00;-#,###,##0.00")
  
  
 If gsAccionRep = 2 Then printl (Space(2) & Space(63) & "VIENEN..." & Space(26) & pVienenHaberME & Space(3) & pVienenDebe & pVienenHaber)
 If gsAccionRep = 3 Then printl (Space(2) & Space(66) & "VIENEN..." & Space(26) & pVienenHaberME & Space(1) & pVienenDebeCta & pVienenHaber)
End Select
Exit Sub
Control:
 MsgBox Err.Description
End Sub
