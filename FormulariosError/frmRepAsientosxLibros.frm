VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmRepAsientosxLibros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Analisis por Libros"
   ClientHeight    =   7905
   ClientLeft      =   2865
   ClientTop       =   3735
   ClientWidth     =   7320
   Icon            =   "frmRepAsientosxLibros.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7905
   ScaleWidth      =   7320
   Begin VB.Frame fraTodo 
      Height          =   7845
      Left            =   45
      TabIndex        =   28
      Top             =   -15
      Width           =   7170
      Begin VB.Frame FraENTIDADES 
         Caption         =   " ENTIDADES"
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
         Left            =   135
         TabIndex        =   54
         Top             =   5850
         Width           =   4965
         Begin TrueOleDBList70.TDBCombo tdbcTipoEntidad 
            Height          =   300
            Left            =   1140
            TabIndex        =   23
            Tag             =   "_"
            Top             =   270
            Width           =   2460
            _ExtentX        =   4339
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
            _PropDict       =   $"frmRepAsientosxLibros.frx":0ECA
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
         Begin TDBText6Ctl.TDBText tdbtCodigo 
            Height          =   315
            Left            =   3660
            TabIndex        =   24
            Top             =   270
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   556
            Caption         =   "frmRepAsientosxLibros.frx":0F51
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmRepAsientosxLibros.frx":0FBD
            Key             =   "frmRepAsientosxLibros.frx":0FDB
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
            Format          =   "a"
            FormatMode      =   1
            AutoConvert     =   -1
            ErrorBeep       =   0
            MaxLength       =   5
            LengthAsByte    =   0
            Text            =   ""
            Furigana        =   0
            HighlightText   =   -1
            IMEMode         =   0
            IMEStatus       =   0
            DropWndWidth    =   0
            DropWndHeight   =   0
            ScrollBarMode   =   0
            MoveOnLRKey     =   0
            OLEDragMode     =   0
            OLEDropMode     =   0
         End
         Begin VB.Label lblEntidad 
            AutoSize        =   -1  'True
            Caption         =   "Entidad"
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
            Left            =   180
            TabIndex        =   55
            Top             =   315
            Width           =   630
         End
      End
      Begin VB.Frame Frame7 
         Height          =   825
         Left            =   4995
         TabIndex        =   49
         Top             =   4935
         Width           =   2025
         Begin TDBText6Ctl.TDBText tdbtTipoDoc 
            Height          =   315
            Left            =   1065
            TabIndex        =   22
            Tag             =   "_"
            Top             =   330
            Width           =   615
            _Version        =   65536
            _ExtentX        =   1085
            _ExtentY        =   556
            Caption         =   "frmRepAsientosxLibros.frx":102D
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmRepAsientosxLibros.frx":1099
            Key             =   "frmRepAsientosxLibros.frx":10B7
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
            MaxLength       =   2
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
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Codigo"
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
            Left            =   210
            TabIndex        =   51
            Top             =   360
            Width           =   600
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "TIPO DOCUM."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   180
            TabIndex        =   50
            Top             =   0
            Width           =   1305
         End
      End
      Begin VB.Frame fraVouchers 
         Height          =   825
         Left            =   135
         TabIndex        =   45
         Top             =   4935
         Width           =   4815
         Begin TDBText6Ctl.TDBText tdbtVoucherIni 
            Height          =   315
            Left            =   915
            TabIndex        =   20
            Tag             =   "_"
            Top             =   360
            Width           =   1290
            _Version        =   65536
            _ExtentX        =   2275
            _ExtentY        =   556
            Caption         =   "frmRepAsientosxLibros.frx":10F9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmRepAsientosxLibros.frx":1165
            Key             =   "frmRepAsientosxLibros.frx":1183
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
         Begin TDBText6Ctl.TDBText tdbtVoucherFin 
            Height          =   315
            Left            =   3255
            TabIndex        =   21
            Tag             =   "_"
            Top             =   360
            Width           =   1290
            _Version        =   65536
            _ExtentX        =   2275
            _ExtentY        =   556
            Caption         =   "frmRepAsientosxLibros.frx":11C5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmRepAsientosxLibros.frx":1231
            Key             =   "frmRepAsientosxLibros.frx":124F
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
         Begin VB.Label Label10 
            Caption         =   "RANGO DE VOUCHERS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   180
            TabIndex        =   48
            Top             =   0
            Width           =   2265
         End
         Begin VB.Label Label9 
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
            Left            =   180
            TabIndex        =   47
            Top             =   360
            Width           =   555
         End
         Begin VB.Label Label8 
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
            Left            =   2520
            TabIndex        =   46
            Top             =   405
            Width           =   555
         End
      End
      Begin VB.Frame Frame6 
         Height          =   1185
         Left            =   120
         TabIndex        =   40
         Top             =   3720
         Width           =   6900
         Begin VB.CheckBox chkCuentas 
            Caption         =   "POR RANGO DE CUENTAS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   0
            Value           =   1  'Checked
            Width           =   2520
         End
         Begin TDBText6Ctl.TDBText TDBInicio 
            Height          =   315
            Left            =   915
            TabIndex        =   18
            Tag             =   "_"
            Top             =   330
            Width           =   1290
            _Version        =   65536
            _ExtentX        =   2275
            _ExtentY        =   556
            Caption         =   "frmRepAsientosxLibros.frx":1291
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmRepAsientosxLibros.frx":12FD
            Key             =   "frmRepAsientosxLibros.frx":131B
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
            Left            =   2235
            TabIndex        =   43
            Tag             =   "_"
            Top             =   330
            Width           =   4305
            _Version        =   65536
            _ExtentX        =   7594
            _ExtentY        =   556
            Caption         =   "frmRepAsientosxLibros.frx":135D
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmRepAsientosxLibros.frx":13C9
            Key             =   "frmRepAsientosxLibros.frx":13E7
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
         Begin TDBText6Ctl.TDBText TDBFinal 
            Height          =   315
            Left            =   915
            TabIndex        =   19
            Tag             =   "_"
            Top             =   690
            Width           =   1290
            _Version        =   65536
            _ExtentX        =   2275
            _ExtentY        =   556
            Caption         =   "frmRepAsientosxLibros.frx":1429
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmRepAsientosxLibros.frx":1495
            Key             =   "frmRepAsientosxLibros.frx":14B3
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
            Left            =   2235
            TabIndex        =   44
            Tag             =   "_"
            Top             =   690
            Width           =   4305
            _Version        =   65536
            _ExtentX        =   7594
            _ExtentY        =   556
            Caption         =   "frmRepAsientosxLibros.frx":14F5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmRepAsientosxLibros.frx":1561
            Key             =   "frmRepAsientosxLibros.frx":157F
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
         Begin VB.Label Label7 
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
            Left            =   180
            TabIndex        =   42
            Top             =   720
            Width           =   555
         End
         Begin VB.Label Label6 
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
            Left            =   180
            TabIndex        =   41
            Top             =   360
            Width           =   555
         End
      End
      Begin VB.CheckBox chkPeriodo 
         Caption         =   "POR RANGO DE PERIODO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   225
         TabIndex        =   5
         Top             =   2475
         Width           =   2475
      End
      Begin VB.CheckBox chkFechas 
         Caption         =   "POR RANGO DE FECHAS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   225
         TabIndex        =   2
         Top             =   1305
         Width           =   2385
      End
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
         Height          =   1230
         Left            =   135
         TabIndex        =   27
         Top             =   2430
         Width           =   3525
         Begin TrueOleDBList70.TDBCombo tdbcMes 
            Height          =   300
            Left            =   1305
            TabIndex        =   6
            Top             =   315
            Width           =   1815
            _ExtentX        =   3201
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
            _PropDict       =   $"frmRepAsientosxLibros.frx":15B3
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
         Begin TrueOleDBList70.TDBCombo tdbcMesFin 
            Height          =   300
            Left            =   1305
            TabIndex        =   7
            Top             =   765
            Width           =   1815
            _ExtentX        =   3201
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
            _PropDict       =   $"frmRepAsientosxLibros.frx":163A
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
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Fin"
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
            Left            =   495
            TabIndex        =   39
            Top             =   810
            Width           =   255
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Inicio"
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
            Left            =   495
            TabIndex        =   38
            Top             =   360
            Width           =   435
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "RANGO DE MONTOS"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1230
         Left            =   3825
         TabIndex        =   35
         Top             =   2430
         Width           =   3210
         Begin TDBNumber6Ctl.TDBNumber tdbnDesde 
            Height          =   300
            Left            =   1260
            TabIndex        =   15
            Tag             =   "enabled"
            Top             =   360
            Width           =   1560
            _Version        =   65536
            _ExtentX        =   2752
            _ExtentY        =   529
            Calculator      =   "frmRepAsientosxLibros.frx":16C1
            Caption         =   "frmRepAsientosxLibros.frx":16E1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmRepAsientosxLibros.frx":174D
            Keys            =   "frmRepAsientosxLibros.frx":176B
            Spin            =   "frmRepAsientosxLibros.frx":17B3
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "##,###,##0.00"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "##,###,##0.00"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   999999999
            MinValue        =   -999999999
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
            MaxValueVT      =   1380909061
            MinValueVT      =   1162608645
         End
         Begin TDBNumber6Ctl.TDBNumber tdbnHasta 
            Height          =   300
            Left            =   1260
            TabIndex        =   16
            Tag             =   "enabled"
            Top             =   735
            Width           =   1560
            _Version        =   65536
            _ExtentX        =   2752
            _ExtentY        =   529
            Calculator      =   "frmRepAsientosxLibros.frx":17DB
            Caption         =   "frmRepAsientosxLibros.frx":17FB
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmRepAsientosxLibros.frx":1867
            Keys            =   "frmRepAsientosxLibros.frx":1885
            Spin            =   "frmRepAsientosxLibros.frx":18BF
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "##,###,##0.00"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "##,###,##0.00"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   999999999
            MinValue        =   -999999999
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
            MaxValueVT      =   1380909061
            MinValueVT      =   1162608645
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            Height          =   195
            Index           =   0
            Left            =   435
            TabIndex        =   37
            Top             =   750
            Width           =   420
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            Height          =   195
            Left            =   435
            TabIndex        =   36
            Top             =   390
            Width           =   465
         End
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1125
         Left            =   135
         TabIndex        =   32
         Top             =   1290
         Width           =   3525
         Begin TDBDate6Ctl.TDBDate dtpDesde 
            Height          =   300
            Left            =   1320
            TabIndex        =   3
            Tag             =   "enabled"
            Top             =   270
            Width           =   1530
            _Version        =   65536
            _ExtentX        =   2699
            _ExtentY        =   529
            Calendar        =   "frmRepAsientosxLibros.frx":18E7
            Caption         =   "frmRepAsientosxLibros.frx":19E9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmRepAsientosxLibros.frx":1A4D
            Keys            =   "frmRepAsientosxLibros.frx":1A6B
            Spin            =   "frmRepAsientosxLibros.frx":1ABF
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
            Left            =   1305
            TabIndex        =   4
            Tag             =   "enabled"
            Top             =   690
            Width           =   1530
            _Version        =   65536
            _ExtentX        =   2699
            _ExtentY        =   529
            Calendar        =   "frmRepAsientosxLibros.frx":1AE7
            Caption         =   "frmRepAsientosxLibros.frx":1BE9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmRepAsientosxLibros.frx":1C4D
            Keys            =   "frmRepAsientosxLibros.frx":1C6B
            Spin            =   "frmRepAsientosxLibros.frx":1CBF
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
         Begin VB.Label Label11 
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
            Left            =   465
            TabIndex        =   34
            Top             =   315
            Width           =   555
         End
         Begin VB.Label Label3 
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
            Index           =   0
            Left            =   465
            TabIndex        =   33
            Top             =   735
            Width           =   495
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "OPCIONES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2205
         Left            =   3840
         TabIndex        =   31
         Top             =   240
         Width           =   3225
         Begin VB.OptionButton Option1 
            Caption         =   "Solo asientos descuadrados"
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
            Left            =   45
            TabIndex        =   14
            Top             =   1845
            Width           =   2775
         End
         Begin VB.CheckBox chkFiltroRango 
            Caption         =   "Filtrar por Rango de Montos"
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
            TabIndex        =   10
            Top             =   930
            Width           =   2655
         End
         Begin VB.CheckBox chkDestino 
            Caption         =   "Incluir Asientos por Destino"
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
            Left            =   375
            TabIndex        =   9
            Top             =   615
            Width           =   2655
         End
         Begin VB.OptionButton optOpciones 
            Caption         =   "Reporte de cuentas sin amarre"
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
            Index           =   1
            Left            =   75
            TabIndex        =   11
            Top             =   1215
            Width           =   2970
         End
         Begin VB.OptionButton optOpciones 
            Caption         =   "Reporte de Consistencias"
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
            Left            =   75
            TabIndex        =   8
            Top             =   315
            Value           =   -1  'True
            Width           =   3060
         End
         Begin VB.CheckBox chkClase9 
            Caption         =   "Clase 9"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1860
            TabIndex        =   13
            Top             =   1530
            Width           =   1005
         End
         Begin VB.CheckBox chkClase6 
            Caption         =   "Clase 6"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   375
            TabIndex        =   12
            Top             =   1530
            Width           =   1125
         End
      End
      Begin TrueOleDBList70.TDBCombo tdbcMoneda 
         Height          =   300
         Left            =   1125
         TabIndex        =   1
         Tag             =   "_"
         Top             =   765
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
         Locked          =   0   'False
         ScrollTrack     =   0   'False
         RowDividerColor =   12632256
         RowSubDividerColor=   12632256
         AddItemSeparator=   ";"
         _PropDict       =   $"frmRepAsientosxLibros.frx":1CE7
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
      Begin TrueOleDBList70.TDBCombo tdbcLibro 
         Height          =   300
         Left            =   1140
         TabIndex        =   0
         Tag             =   "enabled"
         Top             =   270
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
         _PropDict       =   $"frmRepAsientosxLibros.frx":1D6E
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
      Begin MSForms.CommandButton cmdSalir 
         Height          =   435
         Left            =   3645
         TabIndex        =   26
         Top             =   7200
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
         Left            =   1800
         TabIndex        =   25
         Top             =   7200
         Width           =   1665
         Caption         =   " Imprimir"
         PicturePosition =   327683
         Size            =   "2937;767"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label lblMensaje 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "(PRESIONE F1 PARA BUSCAR VOUCHERS Y TIPO DE DOCUMENTO)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   240
         TabIndex        =   52
         Top             =   6750
         Visible         =   0   'False
         Width           =   6690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Libro"
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
         Left            =   270
         TabIndex        =   30
         Top             =   315
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Moneda"
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
         Left            =   270
         TabIndex        =   29
         Top             =   795
         Width           =   660
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
      TabIndex        =   53
      Top             =   0
      Width           =   4365
   End
End
Attribute VB_Name = "frmRepAsientosxLibros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Control As String
Dim gsGrupo As String

Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

Private Sub chkClase6_Click()
    optOpciones(1).Value = True
    optOpciones_Click (1)
End Sub

Private Sub chkClase6_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    pSetFocus chkClase9
End If
End Sub

Private Sub chkClase9_Click()
    optOpciones(1).Value = True
    optOpciones_Click (1)
End Sub

Private Sub chkClase9_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    pSetFocus cmdImprimir
End If

End Sub

Private Sub chkDestino_Click()
    optOpciones(0).Value = True
    optOpciones_Click (0)
End Sub

Private Sub chkDestino_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pSetFocus chkFiltroRango
End If
End Sub


Private Sub chkFiltroRango_Click()
    optOpciones(0).Value = True
    optOpciones_Click (0)
    Me.Frame4.Enabled = True
    If chkFiltroRango.Value = vbChecked Then
        ActivarControl tdbnDesde, True
        ActivarControl tdbnHasta, True
    Else
        tdbnDesde.Value = 0
        tdbnHasta.Value = 0
        ActivarControl tdbnDesde, False
        ActivarControl tdbnHasta, False
    End If
    
    'Frame4.Enabled = Not (Frame4.Enabled)
End Sub

Private Sub chkFiltroRango_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pSetFocus optOpciones(1)
End If
End Sub

Private Function Valida() As Boolean
    If CE(dtpDesde) = "" Then
        Mensajes "Ingrese una fecha de inicio", vbOKOnly + vbInformation
        Valida = False
        Exit Function
    End If
    
    If CE(dtpHasta) = "" Then
        Mensajes "Ingrese una fecha final", vbOKOnly + vbInformation
        Valida = False
        Exit Function
    End If
    
    If chkPeriodo.Value = vbUnchecked And chkFechas.Value = vbUnchecked Then
        Mensajes "Seleccione la opción : " & Salto(2) & "RANGO DE FECHAS o RANGO DE PERIODO", vbOKOnly + vbInformation
        Valida = False
        Exit Function
    
    End If

    If chkPeriodo.Value = vbChecked And CE(tdbcMes.Text) = "" Then
        Mensajes "Seleccione el periodo de inicio ", vbOKOnly + vbInformation
        Valida = False
        Exit Function
    End If
    
    If chkPeriodo.Value = vbChecked And CE(tdbcMesFin.Text) = "" Then
        Mensajes "Seleccione el periodo de final", vbOKOnly + vbInformation
        Valida = False
        Exit Function
    End If
    
    If chkPeriodo.Value = vbChecked And CE(tdbcMes.BoundText) > CE(tdbcMesFin.BoundText) Then
        Mensajes "El periodo final no debe ser mayor al inicial", vbOKOnly + vbInformation
        Valida = False
        Exit Function
    End If
    
    If chkFechas.Value = vbChecked And Format(dtpDesde.Value, "yyyyMMdd") > Format(dtpHasta.Value, "yyyyMMdd") Then
        Mensajes "La fecha final no debe ser mayor a la fecha inicial", vbOKOnly + vbInformation
        Valida = False
        Exit Function
    End If
    
    If chkCuentas.Value = vbChecked Then
        If CE(TDBInicio.Text) = "" Then
            Mensajes "Ingrese la cuenta de inicio"
            Valida = False
            Exit Function
        End If
        
        If CE(TDBFinal.Text) = "" Then
            Mensajes "Ingrese la cuenta final"
            Valida = False
            Exit Function
        End If
        
        If CE(TDBInicio.Text) > CE(TDBFinal.Text) Then
            Mensajes "La cuenta de inicio debe ser menor a la cuenta final"
            Valida = False
            Exit Function
        End If
    End If
    
    If chkFiltroRango.Value = vbChecked Then
        
        If CE(tdbnDesde.Text) > CE(tdbnHasta.Text) Then
            Mensajes "El monto inicial no debe ser mayor al monto final"
            Valida = False
            Exit Function
        End If
    End If
    
    If CE(tdbtVoucherIni.Text) <> "" Or CE(tdbtVoucherFin.Text) <> "" Then
        If CE(tdbtVoucherIni.Text) = "" Then
            Mensajes "Ingrese el voucher inicial"
            Valida = False
            Exit Function
        End If
    
        If CE(tdbtVoucherFin.Text) = "" Then
            Mensajes "Ingrese el voucher final"
            Valida = False
            Exit Function
        End If
        
        If CE(tdbtVoucherIni.Text) > CE(tdbtVoucherFin.Text) Then
            Mensajes "El voucher final no debe ser menor que el voucher inicial"
            Valida = False
            Exit Function
        End If
        
    End If
    
    
    Valida = True
End Function

Private Sub cmdImprimir_Click()

Dim Nombre_Rep As String
    Dim formulas(0) As Variant
    
    If Valida = False Then Exit Sub
    Screen.MousePointer = vbHourglass
    

    '------------------- REPORTE DE CUENTAS POR LIBRO --------------------------------'
    If optOpciones(0).Value Then
       ' *** Abrir el reporte y enviar los parametros
       Dim matriz_fecha(20) As Variant
       
       If chkDestino.Value = "1" Then
          formulas(0) = "conDestino = '*'"
       Else
          formulas(0) = "conDestino = '0'"
       End If
       matriz_fecha(0) = "@Emp_cCodigo;" & gsEmpresa & ";True"
       matriz_fecha(1) = "@Pan_cAnio;" & gsAnio & ";True"
       '------------------------------------------------
       If chkFechas.Value = vbChecked Then
            matriz_fecha(2) = "@Per_cPeriodo;" & Format(Month(dtpDesde), "00") & ";True"
       End If
       
       If chkPeriodo.Value = vbChecked Then
            On Error Resume Next
            'dtpDesde = "01/" & tdbcMes.BoundText & "/" & gsAnio
            matriz_fecha(2) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
       End If
       
       If chkFechas.Value = vbChecked Then
            matriz_fecha(11) = "@Per_cPeriodoFin;" & Format(Month(dtpDesde), "00") & ";True"
       End If
       
       If chkPeriodo.Value = vbChecked Then
            On Error Resume Next
            'dtpHasta = UltimoDiaMes(tdbcMesFin.BoundText, gsAnio)
            matriz_fecha(11) = "@Per_cPeriodoFin;" & tdbcMesFin.BoundText & ";True"
       End If
       
       '------------------------------------------------
       If tdbcLibro.BoundText = "" Then
          matriz_fecha(3) = "@Lib_cTipoLibro;%%;True"
       Else
          matriz_fecha(3) = "@Lib_cTipoLibro;" & tdbcLibro.BoundText & ";True"
       End If
       matriz_fecha(4) = "@desde;" & dtpDesde.Text & ";True"
       matriz_fecha(5) = "@hasta;" & dtpHasta.Text & ";True"
       matriz_fecha(6) = "@moneda;" & tdbcMoneda.BoundText & ";True"
       matriz_fecha(7) = "@cFiltroRan;" & chkFiltroRango.Value & ";True"
       matriz_fecha(8) = "@MontoIni;" & tdbnDesde.Value & ";True"
       matriz_fecha(9) = "@MontoFin;" & tdbnHasta.Value & ";True"
       matriz_fecha(10) = "@Accion;TODOS;True"
       
       
       matriz_fecha(12) = "@CuentaIni;" & TDBInicio.Text & ";True"
       matriz_fecha(13) = "@CuentaFin;" & TDBFinal.Text & ";True"
       matriz_fecha(14) = "@VoucherIni;" & tdbtVoucherIni.Text & ";True"
       matriz_fecha(15) = "@VoucherFin;" & tdbtVoucherFin.Text & ";True"
       matriz_fecha(16) = "@TipoDoc;" & tdbtTipoDoc.Text & ";True"
       
       matriz_fecha(17) = "@cTipoEntidad;" & tdbcTipoEntidad.BoundText & ";True"
       matriz_fecha(18) = "@cCodEntidad;" & tdbtCodigo.Text & ";True"
              
       matriz_fecha(19) = "@EMPRESA;" & gsEmpresaNom & ";True"
       matriz_fecha(20) = "@RUC;" & "RUC : " & gsRUC & ";True"
       
       
       AbreReporteParam gsDSN, Me, rutaReportes & "RptAsientosxLibro.RPT", crptToWindow, "Reporte de Asientos por Libros Contables", "", matriz_fecha(), formulas()
      
    '------------------- REPORTE DE LIBROS DESCUADRADOS --------------------------------'
    ElseIf Option1.Value = True Then
       Dim matriz_fecha3(14) As Variant
       
       matriz_fecha3(0) = "@Emp_cCodigo;" & gsEmpresa & ";True"
       matriz_fecha3(1) = "@Pan_cAnio;" & gsAnio & ";True"
       
       matriz_fecha3(2) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
       matriz_fecha3(11) = "@Per_cPeriodoFin;" & tdbcMesFin.BoundText & ";True"
       
       If tdbcLibro.BoundText = "" Then
          matriz_fecha3(3) = "@Lib_cTipoLibro;;True"
       Else
          matriz_fecha3(3) = "@Lib_cTipoLibro;" & tdbcLibro.BoundText & ";True"
       End If
       matriz_fecha3(4) = "@desde;" & dtpDesde.Text & ";True"
       matriz_fecha3(5) = "@hasta;" & dtpHasta.Text & ";True"
       matriz_fecha3(6) = "@moneda;" & tdbcMoneda.BoundText & ";True"
       matriz_fecha3(7) = "@cFiltroRan;" & chkFiltroRango.Value & ";True"
       matriz_fecha3(8) = "@MontoIni;" & tdbnDesde.Value & ";True"
       matriz_fecha3(9) = "@MontoFin;" & tdbnHasta.Value & ";True"
       matriz_fecha3(10) = "@Accion;SALDO;True"
       
       If chkFechas.Value = vbChecked Then
            matriz_fecha3(4) = "@desde;" & dtpDesde.Text & ";True"
            matriz_fecha3(5) = "@hasta;" & dtpHasta.Text & ";True"
       Else
            matriz_fecha3(4) = "@desde;;True"
            matriz_fecha3(5) = "@hasta;;True"
       End If
       
       
       If chkPeriodo.Value = vbChecked Then
            matriz_fecha3(2) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
            matriz_fecha3(11) = "@Per_cPeriodoFin;" & tdbcMesFin.BoundText & ";True"
       Else
            matriz_fecha3(2) = "@Per_cPeriodo;;True"
            matriz_fecha3(11) = "@Per_cPeriodoFin;;True"
       End If
       
       matriz_fecha3(12) = "NombreEmp;" & gsEmpresaNom & ";True"
       matriz_fecha3(13) = "@EMPRESA;" & gsEmpresaNom & ";True"
       matriz_fecha3(14) = "@RUC;" & "RUC : " & gsRUC & ";True"
        
       AbreReporteParam gsDSN, Me, rutaReportes & "RptAsientosxLibroDescuad.RPT", crptToWindow, "Reporte de Asientos por Libros Contables", "", matriz_fecha3(), formulas()
       
       
    '------------------- REPORTE DE CUENTAS SIN AMARRE --------------------------------'
    Else
    
       If chkClase6.Value = 0 And chkClase9.Value = 0 Then
          Screen.MousePointer = vbNormal
          Mensajes "Seleccione una de las clases 6 o 9...", vbInformation
          
          Exit Sub
       End If
    
       ' *** Abrir el reporte y enviar los parametros
       Dim matriz_fecha2(7) As Variant
       Dim cClaseCta As String
    
       matriz_fecha2(0) = "@Emp_cCodigo;" & gsEmpresa & ";True"
       If tdbcLibro.BoundText = "" Then
          matriz_fecha2(1) = "@Lib_cTipoLibro;%%;True"
       Else
          matriz_fecha2(1) = "@Lib_cTipoLibro;" & tdbcLibro.BoundText & ";True"
       End If
       matriz_fecha2(2) = "@desde;" & dtpDesde.Text & ";True"
       matriz_fecha2(3) = "@hasta;" & dtpHasta.Text & ";True"
       matriz_fecha2(4) = "@moneda;" & tdbcMoneda.BoundText & ";True"
       
       cClaseCta = ""
       If chkClase6.Value = 1 Then
          cClaseCta = "6"
       End If
       If chkClase9.Value = 1 Then
          If Len(Trim(cClaseCta)) > 0 Then
             cClaseCta = "[69]"
          Else
             cClaseCta = "9"
          End If
       End If
       matriz_fecha2(5) = "@condicta69;" & cClaseCta & ";True"
       matriz_fecha2(6) = "@EMPRESA;" & gsEmpresaNom & ";True"
       matriz_fecha2(7) = "@RUC;" & "RUC : " & gsRUC & ";True"
    
       AbreReporteParam gsDSN, Me, rutaReportes & "RptCuentaSinAmarre.rpt", crptToWindow, "Reporte de Cuentas Sin Amarre", "", matriz_fecha2(), formulas()
       
       ' ***
    End If
    Screen.MousePointer = vbNormal

End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub dtpDesde_LostFocus()
    On Error Resume Next
    If CE(dtpDesde) = "" Then dtpDesde = Date
    tdbcMes.BoundText = Right("00" & dtpDesde.Month, 2)

End Sub

Private Sub dtpHasta_LostFocus()
    On Error Resume Next
    If CE(dtpHasta) = "" Then dtpHasta = Date
    tdbcMesFin.BoundText = Right("00" & dtpHasta.Month, 2)
End Sub

Private Sub Form_Load()
    Dim VarMes As String
    Dim sqlcombos As String
        
    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))
    
    
    Call Centrar_form(Me)
    
    If gsPeriodo = "00" Then
        VarMes = "01/01/" + gsAnio
        dtpHasta = UltimoDiaMes("01", gsAnio)
    ElseIf gsPeriodo = "13" Or gsPeriodo = "14" Then
        VarMes = "01/12/" + gsAnio
        dtpHasta = UltimoDiaMes("12", gsAnio)
    Else
        VarMes = "01/" & gsPeriodo & "/" + gsAnio
        dtpHasta = UltimoDiaMes(gsPeriodo, gsAnio)
    End If
    
    dtpDesde = VarMes
    
    Call LlenaComboMesApeAddItem(tdbcMes)
    Call LlenaComboMesApeAddItem(tdbcMesFin)
    tdbcMes.ReBind
    tdbcMesFin.ReBind
    
    tdbcMes.BoundText = gsPeriodo
    tdbcMesFin.BoundText = gsPeriodo
    
    Call LlenaCombos
    Call BuscarMonedaNacional
    
    chkFechas.Value = vbChecked
    
    chkFiltroRango_Click
    
    chkCuentas.Value = vbUnchecked
    chkCuentas_Click
    
    ' *** Llenando el tipo de Entidad
    sqlcombos = "SELECT Ten_cTipoEntidad, Ten_cNombreEntidad From CNT_ENTIDAD "
    sqlcombos = sqlcombos + "WHERE Emp_cCodigo = '" & gsEmpresa & "' ORDER BY Ten_cNombreEntidad"
    LlenarComboAddItem tdbcTipoEntidad, sqlcombos, True
    
    tdbcMoneda.Bookmark = 0
    
End Sub

Private Sub LlenaCombos()
    Dim sqlcombos As String
    ' *** Llenando los libros
    sqlcombos = "SELECT LIB_CTIPOLIBRO, LIB_CDESCRIPCION FROM CNT_LIBRO_OPERA " & _
                "WHERE EMP_CCODIGO =  '" & gsEmpresa & "' AND Pan_cAnio = '" & gsAnio & "' ORDER BY LIB_CDESCRIPCION "
    LlenarComboAddItem tdbcLibro, sqlcombos
    tdbcLibro.AddItem ";TODOS"
    tdbcLibro.ReBind
    
    ' *** Llenando el tipo de Moneda
    
    sqlcombos = "SELECT Mon_cCodigo, Mon_cNombreLargo, Mon_cMNac, Mon_cMExt From CNT_TIPO_MONEDA " & _
                "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND (Mon_cMNac = '1' or Mon_cMExt = '1') " & _
                "ORDER BY Mon_cCodigo"
    
'    sqlcombos = "SELECT Mon_cCodigo, Mon_cNombreLargo, Mon_cMNac, Mon_cMExt From CNT_TIPO_MONEDA " & _
'                "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND (Mon_cMNac = '1' Or Mon_cMExt = '1') " & _
'                "ORDER BY Mon_cCodigo"

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
        Call Centrar_Objeto(fraTodo, Me)

        Call CentrarTitulo(lblTitulo, fraTodo, Me)
    End If
Exit Sub
errHand:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmRepAsientosxLibros = Nothing
    Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))

End Sub

Private Sub Option1_Click()
        chkClase6.Value = vbUnchecked
        chkClase9.Value = vbUnchecked
        chkDestino.Value = vbUnchecked
        chkFiltroRango.Value = vbUnchecked

End Sub

Private Sub optOpciones_Click(Index As Integer)
    If Index = 0 Then
        chkClase6.Value = vbUnchecked
        chkClase9.Value = vbUnchecked
    Else
        chkDestino.Value = vbUnchecked
        chkFiltroRango.Value = vbUnchecked
    End If
End Sub

Private Sub optOpciones_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    pSetFocus chkClase6
End If
End Sub

Private Sub tdbcLibro_ItemChange()
    If tdbcLibro.BoundText = "" Then
        fraVouchers.Enabled = False
        ActivarControl tdbtVoucherIni, False, gsColorDesactivado
        ActivarControl tdbtVoucherFin, False, gsColorDesactivado
        ActivarControl tdbtTipoDoc, False, gsColorDesactivado
        
        
    Else
        fraVouchers.Enabled = True
        ActivarControl tdbtVoucherIni, True, gsColorActivado
        ActivarControl tdbtVoucherFin, True, gsColorActivado
        ActivarControl tdbtTipoDoc, True, gsColorActivado
    End If
    
    tdbtVoucherIni.Text = ""
    tdbtVoucherFin.Text = ""
    tdbtTipoDoc.Text = ""
    
End Sub

Private Sub tdbcLibro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then pSendKeys "{tab}"
End Sub

Private Sub tdbcMes_SelChange(Cancel As Integer)
    Dim periodo As String
    periodo = tdbcMes.BoundText
    chkPeriodo.Value = vbChecked
    If chkPeriodo.Value = vbChecked Then
        If periodo = "00" Then periodo = "01"
        dtpDesde = "01/" & periodo & "/" & gsAnio
    End If

End Sub

Private Sub tdbcMesFin_SelChange(Cancel As Integer)
    Dim periodo As String
    chkPeriodo.Value = vbChecked
    periodo = tdbcMesFin.BoundText
    If chkPeriodo.Value = vbChecked Then
        If periodo = "00" Then periodo = "01"
        dtpHasta = UltimoDiaMes(periodo, gsAnio)
    End If

End Sub

Private Sub tdbcMoneda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then pSendKeys "{tab}"
End Sub

Private Sub chkPeriodo_Click()
    chkFechas.Value = vbUnchecked
    
    If chkFechas.Value = vbUnchecked Then
        ActivarControl dtpDesde, False
        ActivarControl dtpHasta, False
    End If
    
    
    ActivarControl tdbcMes, True
    ActivarControl tdbcMesFin, True

    Dim periodo As String
    periodo = tdbcMes.BoundText

    If chkPeriodo.Value = vbChecked Then
        If periodo = "00" Then periodo = "01"
        dtpDesde = "01/" & periodo & "/" & gsAnio
        
        periodo = tdbcMesFin.BoundText
        If periodo = "00" Then periodo = "01"
        
        dtpHasta = UltimoDiaMes(periodo, gsAnio)
    Else
    
        ActivarControl tdbcMes, False
        ActivarControl tdbcMesFin, False
        
    End If
End Sub

Private Sub chkFechas_Click()
    chkPeriodo.Value = vbUnchecked
    
    ActivarControl dtpDesde, True
    ActivarControl dtpHasta, True
    ActivarControl tdbcMes, False
    ActivarControl tdbcMesFin, False
    
    If chkFechas.Value = vbUnchecked Then
        ActivarControl dtpDesde, False
        ActivarControl dtpHasta, False
    End If
    
End Sub

Private Sub chkCuentas_Click()
    If chkCuentas.Value = 1 Then
        ActivarControl TDBInicio, True
        ActivarControl TDBFinal, True
        ActivarControl tdbtDescripcionDesde, False
        ActivarControl tdbtDescripcionHasta, False
    

    Else
    
        ActivarControl TDBInicio, False
        ActivarControl TDBFinal, False
        ActivarControl tdbtDescripcionDesde, False
        ActivarControl tdbtDescripcionHasta, False
        
        TDBInicio.Text = ""
        TDBFinal.Text = ""
        tdbtDescripcionDesde.Text = ""
        tdbtDescripcionHasta.Text = ""
        
    End If
End Sub
Private Sub tdbcTipoEntidad_ItemChange()
    If tdbcTipoEntidad.BoundText = "" Then
        tdbtCodigo = ""
        ActivarControl tdbtCodigo, False
    Else
        ActivarControl tdbtCodigo, True
    End If
End Sub
Private Sub tdbcTipoEntidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pSetFocus tdbtCodigo
    If tdbtCodigo.Enabled = True Then
       pSetFocus tdbtCodigo
    Else
       pSetFocus cmdImprimir
    End If
End If
End Sub

Private Sub TDBFinal_Change()
    If CE(TDBFinal) = "" Then tdbtDescripcionHasta = ""
End Sub

Private Sub TDBFinal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then
        Control = "TDBFinal"
        Call LlamaBuscar(frmBuscador, Me.TDBFinal.Name, Control, "Cuentas", Me, gsPeriodo, Me.TDBFinal.Text)
    End If
End Sub

Private Sub TDBFinal_LostFocus()
    If TDBFinal <> "" And Me.Enabled = True Then
        tdbtDescripcionHasta = ExisteCtaNoTitulo(TDBFinal, "")
        If tdbtDescripcionHasta = "" Then pSetFocus TDBFinal
    End If
End Sub

Private Sub TDBInicio_Change()
    If CE(TDBInicio) = "" Then tdbtDescripcionDesde = ""
End Sub

Private Sub TDBInicio_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then
        Control = "TDBInicio"
        Call LlamaBuscar(frmBuscador, Me.TDBInicio.Name, Control, "Cuentas", Me, gsPeriodo, Me.TDBInicio.Text)
    End If
End Sub

Private Sub TDBInicio_LostFocus()
    If TDBInicio <> "" And Me.Enabled = True Then
        tdbtDescripcionDesde = ExisteCtaNoTitulo(TDBInicio, "")
        If tdbtDescripcionDesde = "" Then pSetFocus TDBInicio
    End If

End Sub


Public Sub RecibirDatos(ByVal lControl As String, ByVal param0 As String, ByVal param1 As String, ByVal param2 As String, Optional ByVal param3 As String)
   ' *** Dependiendo del control
    Select Case Control
    Case "TDBInicio" ' *** Caso Desde
        TDBInicio = Trim(param0)
        tdbtDescripcionDesde.Text = Trim(param1)
        Unload frmBuscador
        pSetFocus TDBInicio
    Case "TDBFinal" ' *** Caso Hasta
        TDBFinal = Trim(param0)
        tdbtDescripcionHasta.Text = Trim(param1)
        Unload frmBuscador
        pSetFocus TDBFinal
    Case "tdbtCodigo" '     *** Caso Codigp
        tdbtCodigo = Trim(param0)
        Unload frmBuscador
        pSetFocus tdbtCodigo
    End Select
End Sub

Private Sub tdbnDesde_GotFocus()
    tdbnDesde.SelStart = 0
    tdbnDesde.SelLength = Len(tdbnDesde.Text)

End Sub

Private Sub tdbnHasta_GotFocus()
    tdbnHasta.SelStart = 0
    tdbnHasta.SelLength = Len(tdbnHasta.Text)
End Sub

Private Sub tdbnHasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        pSetFocus cmdImprimir
    End If
End Sub
'Private Sub tdbtCodigo_GotFocus()
'    Call VerMensaje(True, "PRESIONE F1 PARA BUSCAR EL TIPO DE ENTIDAD")
'End Sub

Private Sub tdbtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(tdbcTipoEntidad.BoundText) = "" And tdbcTipoEntidad.BoundText <> "" Then
       Mensajes "Debe ingresar el Tipo de Entidad, verificar...", vbInformation
       pSetFocus tdbtCodigo
       Exit Sub
    End If

    If KeyCode = 112 Then Call LlamaBuscar(frmBuscador, Me.tdbtCodigo.Name, Control, "Entidad", Me, gsPeriodo, tdbcTipoEntidad.BoundText)
End Sub
Private Sub tdbtCodigo_LostFocus()
    Dim valorDato As String, sqlver As String
        
    If tdbtCodigo <> "" And Me.Enabled = True Then
        tdbtCodigo.Text = Right("00000" & CE(tdbtCodigo.Text), 5)
        If Trim(tdbcTipoEntidad.BoundText) = "" Then
           Mensajes "Debe ingresar el Tipo de Entidad, verificar...", vbInformation
           pSetFocus tdbcTipoEntidad
           Exit Sub
        End If
            
        sqlver = "SELECT ENT_CPERSONA From CNM_ENTIDAD "
        sqlver = sqlver + "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND  ent_cEstado = 'A'"
        sqlver = sqlver + "AND Ent_cCodEntidad = '" & tdbtCodigo.Text & "' "
        sqlver = sqlver + "AND Ten_CTipoEntidad = '" & Trim(tdbcTipoEntidad.BoundText) & "' "
        
        valorDato = ExtraeDescripcion(sqlver)
        If valorDato = "" Then
           Mensajes "Codigo de entidad no existe, verificar...", vbInformation
           tdbtCodigo.Text = ""
           pSetFocus tdbtCodigo
           
        End If
    End If
End Sub

Private Sub tdbtVoucherIni_LostFocus()
    Call VerMensaje(False, "")
End Sub

Private Sub tdbtTipoDoc_LostFocus()
    Call VerMensaje(False, "")
End Sub

Private Sub tdbtVoucherFin_LostFocus()
    Call VerMensaje(False, "")
End Sub

Private Sub tdbtTipoDoc_GotFocus()
    Call VerMensaje(True, "PRESIONE F1 PARA BUSCAR EL TIPO DE DOCUMENTO")
End Sub

Private Sub tdbtVoucherFin_GotFocus()
    Call VerMensaje(True, "PRESIONE F1 PARA BUSCAR EL VOUCHER FINAL")
End Sub

Private Sub tdbtVoucherIni_GotFocus()
    Call VerMensaje(True, "PRESIONE F1 PARA BUSCAR EL VOUCHER INICIAL")
End Sub

Private Sub VerMensaje(bValor As Boolean, cMensaje As String)
    lblMensaje.Visible = bValor
    lblMensaje.Caption = cMensaje
    
End Sub


Private Sub tdbtVoucherIni_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim CodOld As String, CodNew As String
    If KeyCode = vbKeyF1 Then
        CodOld = tdbtVoucherIni.Text
        tdbtVoucherIni.Text = ""
        CodNew = BuscarVoucher(Me, tdbcMes.BoundText, tdbcMesFin.BoundText, tdbcLibro.BoundText)
        tdbtVoucherIni.Text = IIf(CodNew <> "", CodNew, CodOld)
    End If
    
    If KeyCode <> vbKeyDelete And KeyCode <> vbKeyF1 Then
        KeyCode = 0
    End If

    If KeyCode = 8 Or KeyCode = vbKeyDelete Then
        tdbtVoucherIni.Text = ""
    End If

End Sub

Private Sub tdbtVoucherIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        tdbtVoucherIni.Text = ""
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub tdbtVoucherFin_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim CodOld As String, CodNew As String
    If KeyCode = vbKeyF1 Then
        CodOld = tdbtVoucherFin.Text
        tdbtVoucherFin.Text = ""
        CodNew = BuscarVoucher(Me, tdbcMes.BoundText, tdbcMesFin.BoundText, tdbcLibro.BoundText)
        tdbtVoucherFin.Text = IIf(CodNew <> "", CodNew, CodOld)
    End If
    
    If KeyCode <> vbKeyDelete And KeyCode <> vbKeyF1 Then
        KeyCode = 0
    End If

    If KeyCode = 8 Or KeyCode = vbKeyDelete Then
        tdbtVoucherFin.Text = ""
    End If

End Sub

Private Sub tdbtVoucherFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        tdbtVoucherFin.Text = ""
    Else
        KeyAscii = 0
    End If
End Sub


Private Sub tdbtTipoDoc_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim CodOld As String, CodNew As String
    
    If KeyCode = vbKeyF1 Then
        CodOld = tdbtTipoDoc.Text
        tdbtTipoDoc.Text = ""
        CodNew = BuscarTipoDoc(Me, "")
        tdbtTipoDoc.Text = IIf(CodNew <> "", CodNew, CodOld)
    End If
    
    If KeyCode <> vbKeyDelete And KeyCode <> vbKeyF1 Then
        KeyCode = 0
    End If

    If KeyCode = 8 Or KeyCode = vbKeyDelete Then
        tdbtTipoDoc.Text = ""
    End If

End Sub

Private Sub tdbtTipoDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        tdbtTipoDoc.Text = ""
    Else
        KeyAscii = 0
    End If
End Sub


