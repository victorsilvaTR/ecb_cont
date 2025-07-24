VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmRepLibroDiario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libro Diario"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8490
   Icon            =   "frmRepLibroDiario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5520
   ScaleWidth      =   8490
   Begin VB.Frame fraTodo 
      Height          =   5370
      Left            =   120
      TabIndex        =   12
      Top             =   45
      Width           =   8220
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
         Height          =   855
         Left            =   1170
         TabIndex        =   33
         Top             =   3405
         Visible         =   0   'False
         Width           =   5865
         Begin MSForms.OptionButton OptImpresion 
            Height          =   510
            Index           =   3
            Left            =   3390
            TabIndex        =   31
            Top             =   210
            Width           =   2175
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   5
            Size            =   "3836;900"
            Value           =   "0"
            Caption         =   "Reporte de Operaciones Individuales"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.OptionButton OptImpresion 
            Height          =   540
            Index           =   2
            Left            =   445
            TabIndex        =   30
            Top             =   210
            Width           =   2310
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   5
            Size            =   "4075;952"
            Value           =   "1"
            Caption         =   "Detalle de Centralización de Operaciones"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   645
         ScaleHeight     =   405
         ScaleWidth      =   7290
         TabIndex        =   27
         Top             =   3075
         Width           =   7290
         Begin VB.OptionButton OptTIpo 
            Caption         =   "Reporte PLE"
            Enabled         =   0   'False
            Height          =   195
            Index           =   2
            Left            =   5160
            TabIndex        =   32
            Top             =   30
            Width           =   1335
         End
         Begin VB.OptionButton OptTIpo 
            Caption         =   "Centralizado"
            Height          =   195
            Index           =   1
            Left            =   2805
            TabIndex        =   29
            Top             =   30
            Width           =   1335
         End
         Begin VB.OptionButton OptTIpo 
            Caption         =   "Detallado"
            Height          =   195
            Index           =   0
            Left            =   540
            TabIndex        =   28
            Top             =   30
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Por Rango de Cuentas"
         Height          =   255
         Left            =   330
         TabIndex        =   7
         Top             =   1710
         Width           =   1935
      End
      Begin VB.Frame Frame3 
         Height          =   1140
         Left            =   150
         TabIndex        =   17
         Top             =   1740
         Width           =   7845
         Begin VB.TextBox Text1 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1320
            TabIndex        =   8
            Top             =   285
            Width           =   1515
         End
         Begin VB.TextBox Text2 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1305
            TabIndex        =   9
            Top             =   645
            Width           =   1545
         End
         Begin TDBText6Ctl.TDBText tdbtDescripcionDesde 
            Height          =   315
            Left            =   2850
            TabIndex        =   20
            Tag             =   "_"
            Top             =   270
            Width           =   4590
            _Version        =   65536
            _ExtentX        =   8096
            _ExtentY        =   556
            Caption         =   "frmRepLibroDiario.frx":0ECA
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmRepLibroDiario.frx":0F36
            Key             =   "frmRepLibroDiario.frx":0F54
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
         Begin TDBText6Ctl.TDBText tdbtDescripcionHasta 
            Height          =   315
            Left            =   2850
            TabIndex        =   21
            Tag             =   "_"
            Top             =   630
            Width           =   4590
            _Version        =   65536
            _ExtentX        =   8096
            _ExtentY        =   556
            Caption         =   "frmRepLibroDiario.frx":0FA6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmRepLibroDiario.frx":1012
            Key             =   "frmRepLibroDiario.frx":1030
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
         Begin VB.Label Label2 
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
            Left            =   360
            TabIndex        =   19
            Top             =   285
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
            Index           =   2
            Left            =   360
            TabIndex        =   18
            Top             =   675
            Width           =   495
         End
      End
      Begin VB.OptionButton Option2 
         Caption         =   "RANGO DE FECHAS"
         Height          =   255
         Left            =   4305
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "PERIODO"
         Height          =   255
         Left            =   1335
         TabIndex        =   0
         Top             =   255
         Value           =   -1  'True
         Width           =   1575
      End
      Begin TDBDate6Ctl.TDBDate dtpDesde 
         Height          =   300
         Left            =   5175
         TabIndex        =   4
         Tag             =   "enabled"
         Top             =   600
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   529
         Calendar        =   "frmRepLibroDiario.frx":1074
         Caption         =   "frmRepLibroDiario.frx":1176
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmRepLibroDiario.frx":11DA
         Keys            =   "frmRepLibroDiario.frx":11F8
         Spin            =   "frmRepLibroDiario.frx":1264
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
         Left            =   5175
         TabIndex        =   5
         Tag             =   "enabled"
         Top             =   1005
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   529
         Calendar        =   "frmRepLibroDiario.frx":128C
         Caption         =   "frmRepLibroDiario.frx":138E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmRepLibroDiario.frx":13F2
         Keys            =   "frmRepLibroDiario.frx":1410
         Spin            =   "frmRepLibroDiario.frx":147C
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
         Left            =   5175
         TabIndex        =   6
         Tag             =   "_"
         Top             =   1410
         Width           =   1815
         _ExtentX        =   3201
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
         _PropDict       =   $"frmRepLibroDiario.frx":14A4
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
      Begin TrueOleDBList70.TDBCombo tdbcMes 
         Height          =   300
         Left            =   2070
         TabIndex        =   2
         Top             =   630
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
         _PropDict       =   $"frmRepLibroDiario.frx":152B
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
         Left            =   2070
         TabIndex        =   3
         Top             =   1005
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
         _PropDict       =   $"frmRepLibroDiario.frx":15B2
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
         Height          =   855
         Left            =   1170
         TabIndex        =   24
         Top             =   3405
         Width           =   5865
         Begin MSForms.OptionButton OptImpresion 
            Height          =   420
            Index           =   0
            Left            =   465
            TabIndex        =   26
            Top             =   255
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
         Begin MSForms.OptionButton OptImpresion 
            Height          =   510
            Index           =   1
            Left            =   3390
            TabIndex        =   25
            Top             =   210
            Width           =   2175
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   5
            Size            =   "3836;900"
            Value           =   "0"
            Caption         =   "Impresión Formato Láser"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "FIN"
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
         Left            =   1305
         TabIndex        =   22
         Top             =   1050
         Width           =   285
      End
      Begin MSForms.CommandButton cmdSalir 
         Height          =   435
         Left            =   4065
         TabIndex        =   11
         Top             =   4650
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
         Left            =   2175
         TabIndex        =   10
         Top             =   4650
         Width           =   1665
         Caption         =   " Vista Previa"
         PicturePosition =   327683
         Size            =   "2937;767"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "INICIO"
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
         Left            =   1305
         TabIndex        =   16
         Top             =   690
         Width           =   540
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
         Left            =   4275
         TabIndex        =   15
         Top             =   645
         Width           =   630
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
         Left            =   4275
         TabIndex        =   14
         Top             =   1035
         Width           =   570
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
         Left            =   4275
         TabIndex        =   13
         Top             =   1455
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
      Height          =   345
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   4365
   End
End
Attribute VB_Name = "frmRepLibroDiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Control As String
Public ReporteSunat As String
Public TituloSunat As String
Public rsArreglo  As ADODB.Recordset
Public rsSumPeriIg  As ADODB.Recordset

Dim Tipo As String
Dim EntroCabecera As Boolean
Dim iReport, nAncho As Integer
Dim gsGrupo As String
Dim swAsientoIgual As Boolean
Dim pPer_cPeriodo As String, xPer_cPeriodo As String
Dim SumPerAnt_DebeSoles As Double
Dim SumPerAnt_HaberSoles As Double
Dim sSql As String

Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

Public Sub CerrarForm()
Unload Me
End Sub

Private Sub Check1_Click()
If Check1.Value Then
    ActivarControl Text1, True
    ActivarControl Text2, True
    
    ActivarControl tdbtDescripcionDesde, False
    ActivarControl tdbtDescripcionHasta, False
    
    pSetFocus Text1
Else
    Text1.Text = ""
    Text2.Text = ""
    
    tdbtDescripcionDesde.Text = ""
    tdbtDescripcionHasta.Text = ""
    
    ActivarControl Text1, False
    ActivarControl Text2, False
    
    ActivarControl tdbtDescripcionDesde, False
    ActivarControl tdbtDescripcionHasta, False
    
End If
End Sub

Private Function Validacion() As Boolean
    Validacion = False
    If Check1.Value Then
        If Text1.Text = "" Then
            Mensajes "Tiene que ingresar el número de cuenta de inicio de la consulta", vbInformation
            pSetFocus Text1
            Exit Function
        ElseIf Text2.Text = "" Then
            Mensajes "Tiene que ingresar el número de cuenta del final de la consulta", vbInformation
            pSetFocus Text2
            Exit Function
        ElseIf Text1.Text > Text2.Text Then
            Mensajes "La cuenta final no debe ser mayo a la cuenta inicial", vbInformation
            pSetFocus Text2
            Exit Function
        End If
    End If
    Validacion = True
End Function

Private Sub cmdImprimir_Click()
    Dim matriz_fecha(21) As Variant

    If Validacion = False Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    If OptTipo(2).Value Then
        matriz_fecha(0) = "@Emp_cCodigo;" & gsEmpresa & ";True"
        matriz_fecha(1) = "@Pan_cAnio;" & gsAnio & ";True"
        matriz_fecha(2) = "@Per_cPeriodo;" & Me.tdbcMes.BoundText & ";True"
        matriz_fecha(3) = "@desde;" & PrimerDiaMes(tdbcMes.BoundText, Year(dtpDesde.Text)) & ";True"
        matriz_fecha(4) = "@hasta;" & UltimoDiaMes(tdbcMesFin.BoundText, Year(dtpHasta.Text)) & ";True"
        matriz_fecha(5) = "@NOMMONEDA;" & tdbcMoneda.Text & ";True"
        matriz_fecha(6) = "@Per_cPeriodoFin;" & Me.tdbcMesFin.BoundText & ";True"
        matriz_fecha(7) = "@TipoRPTPLE;" & IIf(OptImpresion(2).Value = True, 1, 0) & ";True"
        GoTo DCOROI
    End If
    
    Tipo = "DIARIO"
    matriz_fecha(0) = "@Tipo;" & Tipo & ";True"
    matriz_fecha(1) = "@Emp_cCodigo;" & gsEmpresa & ";True"
    matriz_fecha(2) = "@Pan_cAnio;" & gsAnio & ";True"
    matriz_fecha(3) = "@Per_cPeriodo;" & Me.tdbcMes.BoundText & ";True"
    matriz_fecha(19) = "@Per_cPeriodo_Nombre;" & Me.tdbcMes.Text & ";True"
    matriz_fecha(4) = "@Lib_cTipoLibro;%;True"
    matriz_fecha(5) = "@Ase_nVoucher;;True"
    matriz_fecha(6) = "@desde;" & Format(dtpDesde.Text, "yyyy-mm-dd") & ";True"
    matriz_fecha(7) = "@hasta;" & Format(dtpHasta.Text, "yyyy-mm-dd") & ";True"
    matriz_fecha(8) = "@moneda;" & tdbcMoneda.BoundText & ";True"
    matriz_fecha(9) = "@ctaini;" & Text1.Text & ";True"
    matriz_fecha(10) = "@ctafin;" & Text2.Text & ";True"
    matriz_fecha(11) = "@Per_cPeriodoFin;" & Me.tdbcMesFin.BoundText & ";True"
    matriz_fecha(20) = "@Per_cPeriodoFin_Nombre;" & Me.tdbcMesFin.Text & ";True"
    matriz_fecha(12) = "@TipoLibDi;" & IIf(OptTipo(0).Value = True, 0, 1) & ";True"
    matriz_fecha(13) = "@EMPRESA;" & gsEmpresaNom & ";True"
    matriz_fecha(14) = "@RUC;" & gsRUC & ";True"
    matriz_fecha(15) = "@NOMMONEDA;" & tdbcMoneda.Text & ";True"

    Dim VarRst As New ADODB.Recordset
    
    Set VarRst = Fct_Listar_Saldo_Anterior_LibDiario(gsEmpresa, gsAnio, Me.tdbcMes.BoundText, _
    IIf(Text1.Text = "", "10", Text1.Text), IIf(Text2.Text = "", "99999999999", Text2.Text))

    If VarRst.EOF = False Then
        matriz_fecha(17) = "TotalAnt;" & VarRst(0).Value & ";True"
    Else
        matriz_fecha(17) = "TotalAnt;0;True"
    End If
    
    If VarRst.EOF = False Then
        matriz_fecha(18) = "TotalAntH;" & VarRst(1).Value & ";True"
    Else
        matriz_fecha(18) = "TotalAntH;0;True"
    End If
    
    Dim formulas(2) As Variant
    If Option2.Value Then
        Tipo = "DIARIOANUAL"
        matriz_fecha(0) = "@Tipo;" & Tipo & ";True"
    End If
    
    If Check1.Value Then
        formulas(0) = "Desde = " & Text1.Text
        formulas(1) = "Hasta = " & Text2.Text
        'formulas(2) = "PeriodoTexto = " & tdbcMes.Text & " A " & tdbcMesFin.Text & " DEL " & gsAnio
    Else
        formulas(0) = "Desde = ''"
        formulas(1) = "Hasta = ''"
        'formulas(2) = "PeriodoTexto = " & tdbcMes.Text & " A " & tdbcMesFin.Text & " DEL " & gsAnio
    End If
    
DCOROI:
    cmdImprimir.Enabled = False
    
    If OptImpresion(2).Value Then
        AbreReporteParam gsDSN, Me, rutaReportes & "RptDiarioElectronicoDCO.rpt", crptToWindow, "Libro Diario - Detalle Centralizado Operaciones", "", matriz_fecha(), formulas()
    ElseIf OptImpresion(3).Value Then
        AbreReporteParam gsDSN, Me, rutaReportes & "RptDiarioElectronicoROI.rpt", crptToWindow, "Libro Diario - Reporte de Operaciones Individuales", "", matriz_fecha(), formulas()
    End If
    If ReporteSunat = "F0501" And OptImpresion(1).Value Then 'frt_rvie
        AbreReporteParam gsDSN, Me, rutaReportes & "RptFormato_0501" & IIf(OptTipo(0).Value, "r", "") & ".rpt", crptToWindow, "Libro Diario", "", matriz_fecha(), formulas()
    ElseIf ReporteSunat = "F0501" And OptImpresion(0).Value Then
        gsNombreVista = "Libro Diario"
        ImpLibroDiario
    End If
    
    Screen.MousePointer = vbNormal
    cmdImprimir.Enabled = True
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub dtpDesde_LostFocus()
    ' *** Validar el año de trabajo
    Dim fecha As String
    fecha = FechaServidor
    If Year(dtpDesde) <> gsAnio Then
        Mensajes "Año de la fecha debe ser igual al año de trabajo", vbInformation
        'pSetFocus dtpDesde
    End If
End Sub

Private Sub dtpHasta_LostFocus()
    ' *** Validar el año de trabajo
    Dim fecha As String
    fecha = FechaServidor
    If Year(dtpHasta) <> gsAnio Then
        Mensajes "Año de la fecha debe ser igual al año de trabajo", vbInformation
        'pSetFocus dtpHasta
        Exit Sub
    End If
    If Format(dtpHasta, "yyyyMMdd") < Format(dtpDesde, "yyyyMMdd") Then
        Mensajes "Fecha Final no puede ser menor que la Fecha de inicio", vbInformation
    End If
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
    Call BuscarMonedaNacional
    tdbcMes.BoundText = gsPeriodo
    tdbcMesFin.BoundText = gsPeriodo
    
    'ActivarControl Text1, False
    'ActivarControl Text2, False

    Option1_Click
    Option2_Click
    Check1_Click
    
    On Error Resume Next
    
    dtpDesde.Value = "01/" & gsPeriodo & "/" & gsAnio
    dtpHasta.Value = "01/" & gsPeriodo & "/" & gsAnio
    dtpDesde.MaxDate = "31/12/" & gsAnio
    dtpHasta.MaxDate = "31/12/" & gsAnio
    dtpDesde.MinDate = "01/01/" & gsAnio
    dtpHasta.MinDate = "01/01/" & gsAnio
    
    dtpDesde.Enabled = False
    dtpHasta.Enabled = False
    
    tdbcMes.BoundText = gsPeriodo
    
    tdbcMes.ReBind
    tdbcMesFin.ReBind
    
    tdbcMoneda.ReBind
    
End Sub

Private Sub LlenaCombos()
    Dim sqlcombos As String
    
    Call LlenaComboMesApeAddItem(tdbcMes)
    Call LlenaComboMesApeAddItem(tdbcMesFin)
    
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
    Set frmRepLibroDiario = Nothing
    Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))

End Sub

Private Sub Option1_Click()
If Option1.Value Then
    ActivarControl dtpDesde, False
    ActivarControl dtpHasta, False
    ActivarControl tdbcMes, True
    ActivarControl tdbcMesFin, True

Else
    ActivarControl dtpDesde, True
    ActivarControl dtpHasta, True
    ActivarControl tdbcMes, False
    ActivarControl tdbcMesFin, False
End If
End Sub

Private Sub Option2_Click()
If Option2.Value Then
    ActivarControl dtpDesde, True
    ActivarControl dtpHasta, True
    ActivarControl tdbcMes, False
    ActivarControl tdbcMesFin, False
    
    Dim Mes As String
    Mes = gsPeriodo
    If Mes = "00" Then Mes = "01"
    If Mes > "12" Then Mes = "12"
    On Error Resume Next
    dtpHasta = UltimoDiaMes(Mes, gsAnio)
    OptTipo(2).Enabled = False
Else
    ActivarControl dtpDesde, False
    ActivarControl dtpHasta, False
    ActivarControl tdbcMes, True
    ActivarControl tdbcMesFin, True
End If
End Sub

Private Sub optTipo_Click(Index As Integer)
If Index = 2 Then
    Frame1.Visible = False
    Frame2.Visible = True
    OptImpresion(2).Value = True
Else
    Frame1.Visible = True
    Frame2.Visible = False
    If (Not OptImpresion(0).Value And Not OptImpresion(1).Value) Then OptImpresion(0).Value = True
End If
End Sub

Private Sub tdbcMes_ItemChange()
If (tdbcMes.BoundText = "00" Or tdbcMes.BoundText = "13" Or tdbcMes.BoundText = "14") And _
        (tdbcMesFin.BoundText = "00" Or tdbcMesFin.BoundText = "13" Or tdbcMesFin.BoundText = "14") Then
    OptTipo(2).Enabled = False
    OptTipo(2).Value = False
    OptImpresion(1).Value = True
Else
    If tdbcMesFin.Text = tdbcMes.Text Or ((tdbcMes.BoundText = "00" And tdbcMesFin.BoundText = "01") Or (tdbcMes.BoundText = "12" And (tdbcMesFin.BoundText = "13" Or tdbcMesFin.BoundText = "14"))) Then
        OptTipo(2).Enabled = True
    Else
        OptTipo(2).Enabled = False
        OptTipo(0).Value = True
        OptImpresion(1).Value = True
    End If
End If
End Sub

Private Sub tdbcMes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pSetFocus tdbcMoneda
End If
End Sub

Private Sub tdbcMesFin_ItemChange()
If (tdbcMes.BoundText = "00" Or tdbcMes.BoundText = "13" Or tdbcMes.BoundText = "14") And _
        (tdbcMesFin.BoundText = "00" Or tdbcMesFin.BoundText = "13" Or tdbcMesFin.BoundText = "14") Then
    OptTipo(2).Enabled = False
    OptTipo(2).Value = False
Else
    If tdbcMesFin.Text = tdbcMes.Text Or ((tdbcMes.BoundText = "00" And tdbcMesFin.BoundText = "01") Or (tdbcMes.BoundText = "12" And (tdbcMesFin.BoundText = "13" Or tdbcMesFin.BoundText = "14"))) Then
        OptTipo(2).Enabled = True
    Else
        OptTipo(2).Enabled = False
        OptTipo(0).Value = True
    End If
End If
End Sub

Private Sub tdbcMoneda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pSetFocus cmdImprimir
End If

End Sub

Private Sub tdbtDescripcionHasta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then pSendKeys "{tab}"
End Sub

Private Sub Text1_Change()
    If CE(Text1) = "" Then tdbtDescripcionDesde = ""
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Control = "tdbtCuentaDesde"
If KeyCode = 112 Then Call LlamaBuscar(frmBuscador, Name, Text1, "Cuentas", Me, gsPeriodo, Text1.Text)
If KeyCode = 13 Then pSendKeys "{tab}"
End Sub

Private Sub Text1_LostFocus()
    If Text1 <> "" And Me.Enabled = True Then
        tdbtDescripcionDesde = ExisteCtaNoTitulo(Text1, "")
        If Text1 = "" Then pSetFocus Text1
    End If
End Sub

Private Sub Text2_Change()
    If CE(Text2) = "" Then tdbtDescripcionHasta = ""
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
Control = "tdbtCuentaHasta"
If KeyCode = 112 Then Call LlamaBuscar(frmBuscador, Name, Text2, "Cuentas", Me, gsPeriodo, Text2.Text)
If KeyCode = 13 Then pSendKeys "{tab}"
End Sub

Public Sub RecibirDatos(ByVal lControl As String, ByVal param0 As String, ByVal param1 As String, ByVal param2 As String, Optional ByVal param3 As String)
   ' *** Dependiendo del control
    Select Case Control
    Case "tdbtCuentaDesde" ' *** Caso Desde
        Text1.Text = Trim(param0)
        'tdbtDescripcionDesde = Trim(frmBuscador.TDBGTabla.Columns(1).Value)
        Unload frmBuscador
        pSetFocus Text1
    Case "tdbtCuentaHasta" ' *** Caso Hasta
        Text2.Text = Trim(param0)
        'tdbtDescripcionHasta = Trim(frmBuscador.TDBGTabla.Columns(1).Value)
        Unload frmBuscador
        pSetFocus Text2
    End Select
End Sub

Private Sub Text2_LostFocus()
    If Text2 <> "" And Me.Enabled = True Then
        tdbtDescripcionHasta = ExisteCtaNoTitulo(Text2, "")
        If Text2 = "" Then pSetFocus Text2
    End If
End Sub

Sub ImpLibroDiario()
On Error GoTo Control

 gsAccionRep = IIf(OptTipo(0).Value, 7, 4) 'frt_rvie
 gsCodMoneda = tdbcMoneda.BoundText
 
 If Not Option1.Value Then gsLdMesIni = "": gsLdMesFin = "" Else gsLdMesIni = tdbcMes.BoundText: gsLdMesFin = tdbcMesFin.BoundText
 If Not Option2.Value Then gsLdFechIni = "": gsLdFechFin = "" Else gsLdFechIni = Format(dtpDesde.Text, "yyyy-mm-dd"): gsLdFechFin = Format(dtpHasta.Text, "yyyy-mm-dd")
 If Check1.Value = "0" Then gsCtaIni = "": gsCtaFin = "" Else gsCtaIni = Trim(Text1.Text): gsCtaFin = Trim(Text2.Text)
 
 frmFCImpresion.Show
 
Exit Sub
Control:
 MsgBox Err.Description
End Sub

Public Sub ReporteLibDiario(Optional nTipo As Integer = 0) 'frt_rvie
On Error GoTo Control

Dim pAse_nVoucher As String * 10
Dim pAse_dFecha As String * 10
Dim pCAR As String * 28 'frt_rvie
Dim pAsd_cGlosa As String * 29
Dim pAsd_cTipoDoc As String * 3
Dim pAsd_cSerieDoc As String * 5
Dim pAsd_cNumDoc As String * 12
Dim pEnt_nRuc As String * 12
Dim pPla_cCuentaContable As String * 12
Dim pPla_cNombreCuenta As String * 21
Dim pDebe As String * 12
Dim pHaber As String * 12
Dim pSumDebe As String * 14
Dim pSumHaber As String * 14

Dim pSumDebeXA As String * 14
Dim pSumHaberXA As String * 14

Dim pTotalMesDebe As String * 14
Dim pTotalMesHaber As String * 14
Dim pTotalGenDebe As String * 14
Dim pTotalGenHaber As String * 14

Dim pVanDebe As String * 13
Dim pVanHaber As String * 13

Dim pVienenDebe As String * 13
Dim pVienenHaber As String * 13

Dim SpaceMes, giLineasIni, NroSpacesMVan, i As Integer

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
 If iReport = 1 Then rsArreglo.MoveFirst

 gsConTotalPaginas = 0
 
ConectarAdvance
 With rsArreglo
   If .RecordCount > 0 Then
    iReport = 1
    gsPaginaPrincipal = 1
    Call CabeceraLibroDiario(nTipo) 'frt_rvie
    EntroCabecera = False
    .MoveFirst
    pTotalMesDebe = 0: pTotalMesHaber = 0
    pVanDebe = 0: pVanHaber = 0
    pVienenDebe = 0: pVienenHaber = 0
    pTotalGenDebe = 0: pTotalGenHaber = 0
    
    Do While Not .EOF
     If Trim(!Per_cPeriodo) <> Trim(pPer_cPeriodo) And Trim(pPer_cPeriodo) <> "" And giLineas = 11 Then
      RSet pVienenDebe = Format$(pVanDebe, "#,###,###,##0.00;(#,###,###,##0.00)")
      RSet pVienenHaber = Format$(pVanHaber, "#,###,###,##0.00;(#,###,###,##0.00)")
           
      printl (Space(60 + nTipo) & "VIENEN..." & Space(23) & pVienenDebe & Space(2) & pVienenHaber)
     End If
     pPer_cPeriodo = !Per_cPeriodo

      Do While Trim(!Per_cPeriodo) = Trim(pPer_cPeriodo)
       EntroCabecera = False
       pAse_nVoucher = !Ase_nVoucher
       EntroCabecera = False
       
       pSumDebe = 0: pSumHaber = 0
       pSumDebeXA = 0: pSumHaberXA = 0

       Do While Trim(!Per_cPeriodo) = Trim(pPer_cPeriodo) And Trim(!Ase_nVoucher) = Trim(pAse_nVoucher)
       
        If !Mon_cMNac = "1" Then RSet pDebe = Format$(!Asd_nDebeSoles, "#,###,###,##0.00;(#,###,###,##0.00)") Else RSet pDebe = Format$(!Asd_nDebeMonExt, "#,###,###,##0.00;(#,###,###,##0.00)")
        If !Mon_cMNac = "1" Then RSet pHaber = Format$(!Asd_nHaberSoles, "#,###,###,##0.00;(#,###,###,##0.00)") Else RSet pHaber = Format$(!Asd_nHaberMonExt, "#,###,###,##0.00;(#,###,###,##0.00)")

        pAse_dFecha = Format(!Ase_dFecha, "dd/mm/yyyy")
        If Trim(IsNull(!Asd_cGlosa)) Then
            pAsd_cGlosa = ""
        Else
            pAsd_cGlosa = !Asd_cGlosa
        End If
        pAsd_cTipoDoc = !Asd_cTipoDoc
        pAsd_cSerieDoc = !Asd_cSerieDoc
        pAsd_cNumDoc = !Asd_cNumDoc
        
        pCAR = " " 'frt_rvie
        If nTipo > 0 And !Lib_cTipoLibro = "05" Then
            i = InStr(pAsd_cNumDoc, "-")
            If i > 0 Then
                pAsd_cNumDoc = Left(pAsd_cNumDoc, i - 1)
            Else
                i = InStr(pAsd_cNumDoc, "/")
                If i > 0 Then pAsd_cNumDoc = Left(pAsd_cNumDoc, i - 1)
            End If
            RSet pCAR = Right("0000000000" + Trim(!pEnt_nRuc), 11) & Left(pAsd_cTipoDoc, 2) & Left(pAsd_cSerieDoc, 4) & Right("0000000000" + Trim(pAsd_cNumDoc), 10) & " "
            pAsd_cNumDoc = !Asd_cNumDoc
        End If
        
        pPla_cCuentaContable = !Pla_cCuentaContable
        pPla_cNombreCuenta = !Pla_cNombreCuenta
        
        pSumDebe = CDbl(pSumDebe) + CDbl(pDebe): RSet pSumDebe = Format$(pSumDebe, "#,###,###,##0.00;(#,###,###,##0.00)")
        pSumHaber = CDbl(pSumHaber) + CDbl(pHaber): RSet pSumHaber = Format$(pSumHaber, "#,###,###,##0.00;(#,###,###,##0.00)")
        
        pSumDebeXA = CDbl(pSumDebeXA) + CDbl(pDebe): RSet pSumDebeXA = Format$(pSumDebeXA, "#,###,###,##0.00;(#,###,###,##0.00)")
        pSumHaberXA = CDbl(pSumHaberXA) + CDbl(pHaber): RSet pSumHaberXA = Format$(pSumHaberXA, "#,###,###,##0.00;(#,###,###,##0.00)")
        
        pTotalMesDebe = CDbl(pTotalMesDebe) + CDbl(pDebe): RSet pTotalMesDebe = Format$(pTotalMesDebe, "#,###,###,##0.00;(#,###,###,##0.00)")
        pTotalMesHaber = CDbl(pTotalMesHaber) + CDbl(pHaber): RSet pTotalMesHaber = Format$(pTotalMesHaber, "#,###,###,##0.00;(#,###,###,##0.00)")
        If Trim(pPer_cPeriodo) = "00" Then
            pVanDebe = CDbl(pVanDebe) + CDbl(pDebe): RSet pVanDebe = Format$(pVanDebe, "#,###,###,##0.00;(#,###,###,##0.00)")
            pVanHaber = CDbl(pVanHaber) + CDbl(pHaber): RSet pVanHaber = Format$(pVanHaber, "#,###,###,##0.00;(#,###,###,##0.00)")
        Else
            pVanDebe = CDbl(pVanDebe) + CDbl(pDebe): RSet pVanDebe = Format$(pVanDebe, "#,###,###,##0.00;(#,###,###,##0.00)")
            pVanHaber = CDbl(pVanHaber) + CDbl(pHaber): RSet pVanHaber = Format$(pVanHaber, "#,###,###,##0.00;(#,###,###,##0.00)")
        End If
         
         gsLinea = (Space(3) & pAse_nVoucher & Space(1) & IIf(nTipo = 0, "", pCAR) & pAse_dFecha & Space(1) & pAsd_cGlosa & Space(3) & pAsd_cTipoDoc & pAsd_cSerieDoc & pAsd_cNumDoc & Space(1) & pPla_cCuentaContable & Space(3) & pDebe & Space(3) & pHaber)
         printl gsLinea
         
         If giLineas <= 73 Then .MoveNext
         
        If giLineas >= 73 Then '71
            printl ""
            
            printl (Space(60 + nTipo) & "VAN..." & Space(26) & pVanDebe & Space(2) & pVanHaber)
            If giLineas = 75 Then printl ""
            
            pVienenDebe = pVanDebe
            pVienenHaber = pVanHaber
            
            RSet pVienenDebe = Format$(pVienenDebe, "#,###,###,##0.00;(#,###,###,##0.00)")
            RSet pVienenHaber = Format$(pVienenHaber, "#,###,###,##0.00;(#,###,###,##0.00)")
                       
            printl (Space(60 + nTipo) & "VIENEN..." & Space(23) & pVienenDebe & Space(2) & pVienenHaber)
        End If

        If Not .EOF Then
            If Trim(!Per_cPeriodo) = Trim(pPer_cPeriodo) And Trim(!Ase_nVoucher) <> Trim(pAse_nVoucher) Then
                printl (Space(92 + nTipo) & "------------- --------------")
                printl (Space(91 + nTipo) & pSumDebeXA & Space(1) & pSumHaberXA)
                
                pSumDebeXA = 0
                pSumHaberXA = 0
                
                If giLineas >= 73 Then '72
                    printl ""
                    printl (Space(60 + nTipo) & "VAN..." & Space(26) & pVanDebe & Space(2) & pVanHaber)
                    If giLineas = 75 Then printl ""
                    pVienenDebe = pVanDebe
                    pVienenHaber = pVanHaber
                    
                    RSet pVienenDebe = Format$(pVienenDebe, "#,###,###,##0.00;(#,###,###,##0.00)")
                    RSet pVienenHaber = Format$(pVienenHaber, "#,###,###,##0.00;(#,###,###,##0.00)")
                    
                    printl (Space(60 + nTipo) & "VIENEN..." & Space(23) & pVienenDebe & Space(2) & pVienenHaber)
                End If
            End If
        Else
            printl (Space(92 + nTipo) & "------------- --------------")
            printl (Space(91 + nTipo) & pSumDebe & Space(1) & pSumHaber)
            pVanDebe = CDbl(pVanDebe) + CDbl(pSumDebe): RSet pVanDebe = Format$(pVanDebe, "#,###,###,##0.00;(#,###,###,##0.00)")
            pVanHaber = CDbl(pVanHaber) + CDbl(pSumHaber): RSet pVanHaber = Format$(pVanHaber, "#,###,###,##0.00;(#,###,###,##0.00)")
            Exit Do
        End If
       Loop
       
       If .EOF Then
            printl (Space(92 + nTipo) & "------------- --------------")
            If Trim(Len(NombreMes(pPer_cPeriodo))) > Len("ENERO") Then
                SpaceMes = 22 - (Len(NombreMes(pPer_cPeriodo)) - Len("ENERO"))
            ElseIf Trim(Len(NombreMes(pPer_cPeriodo))) < Len("ENERO") Then
                SpaceMes = 22 - (Len("ENERO") - Len(NombreMes(pPer_cPeriodo)))
            ElseIf Trim(Len(NombreMes(pPer_cPeriodo))) = Len("ENERO") Then
                SpaceMes = 22
            End If
        pTotalGenDebe = CDbl(pTotalGenDebe) + CDbl(pTotalMesDebe): RSet pTotalGenDebe = Format$(pTotalGenDebe, "#,###,###,##0.00;(#,###,###,##0.00)")
        pTotalGenHaber = CDbl(pTotalGenHaber) + CDbl(pTotalMesHaber): RSet pTotalGenHaber = Format$(pTotalGenHaber, "#,###,###,##0.00;(#,###,###,##0.00)")
        If LTrim(RTrim(NombreMes(pPer_cPeriodo))) <> "" Then
            printl (Space(57 + nTipo) & "TOTAL  " & NombreMes(pPer_cPeriodo) & Space(SpaceMes) & pTotalMesDebe & Space(1) & pTotalMesHaber)
        Else
            printl (Space(57 + nTipo) & "TOTAL  " & NombreMes(pPer_cPeriodo) & Space(SpaceMes) & pTotalMesDebe & Space(1) & pTotalMesHaber)
        End If
        
        Exit Do
       ElseIf !Per_cPeriodo <> Trim(pPer_cPeriodo) Then
            printl (Space(92 + nTipo) & "------------- --------------")
            printl (Space(91 + nTipo) & pSumDebe & Space(1) & pSumHaber)
            printl ("")
            If Trim(Len(NombreMes(pPer_cPeriodo))) > Len("ENERO") Then
                SpaceMes = 22 - (Len(NombreMes(pPer_cPeriodo)) - Len("ENERO"))
            ElseIf Trim(Len(NombreMes(pPer_cPeriodo))) < Len("ENERO") Then
                SpaceMes = 22 - (Len("ENERO") - Len(NombreMes(pPer_cPeriodo)))
            ElseIf Trim(Len(NombreMes(pPer_cPeriodo))) = Len("ENERO") Then
                SpaceMes = 22
            End If
        pTotalGenDebe = CDbl(pTotalGenDebe) + CDbl(pTotalMesDebe): RSet pTotalGenDebe = Format$(pTotalGenDebe, "#,###,###,##0.00;(#,###,###,##0.00)")
        pTotalGenHaber = CDbl(pTotalGenHaber) + CDbl(pTotalMesHaber): RSet pTotalGenHaber = Format$(pTotalGenHaber, "#,###,###,##0.00;(#,###,###,##0.00)")
        
        If LTrim(RTrim(NombreMes(pPer_cPeriodo))) = "" Then
            printl (Space(91 + nTipo) & pTotalMesDebe & Space(1) & pTotalMesHaber)
        Else
            printl (Space(57 + nTipo) & "TOTAL " & NombreMes(pPer_cPeriodo) & Space(SpaceMes + 1) & pTotalMesDebe & Space(1) & pTotalMesHaber)
            pTotalMesDebe = 0
            pTotalMesHaber = 0
        End If
        pSumDebeXA = 0
        pSumHaberXA = 0
        NroSpacesMVan = 74 - giLineas
        giLineasIni = giLineas
        For i = giLineas To giLineasIni + (NroSpacesMVan - 3)
          printl ""
          If i = 72 And Trim(pPer_cPeriodo) <> "00" Then
          
           printl (Space(60 + nTipo) & "VAN..." & Space(26) & pVanDebe & Space(2) & pVanHaber)
           printl ""
          Else
           pVanDebe = CDbl(pTotalMesDebe): pVanHaber = CDbl(pTotalMesHaber)
           RSet pVanDebe = Format$(pVanDebe, "#,###,###,##0.00;(#,###,###,##0.00)")
           RSet pVanHaber = Format$(pVanHaber, "#,###,###,##0.00;(#,###,###,##0.00)")
           
            If Trim(gsLdMesIni) <> Trim(gsLdMesFin) And Trim(pPer_cPeriodo) <> Trim(gsLdMesFin) And Trim(pPer_cPeriodo) <> "00" And pPer_cPeriodo <> "" Then
              If (CDbl(pPer_cPeriodo) - 1) < 10 Then
                    xPer_cPeriodo = "0" & Trim(Str(CDbl(pPer_cPeriodo) - 1))
              Else
                    xPer_cPeriodo = Str(CDbl(pPer_cPeriodo) - 1)
              End If
              
              Set rsSumPeriIg = New ADODB.Recordset
              sSql = "spCn_RptFormato0501_SumPeriodos '" & Tipo & "','" & gsEmpresa & "','" & gsAnio & "','00','%','','',''," & _
                     "'" & gsCodMoneda & "','" & gsCtaIni & "','" & gsCtaFin & "','" & xPer_cPeriodo & "'"
              rsSumPeriIg.Open sSql, gcnSistemaAdv, adOpenDynamic, adLockOptimistic
              If rsSumPeriIg.State <> 0 Then SumPerAnt_DebeSoles = rsSumPeriIg.Fields(0): SumPerAnt_HaberSoles = rsSumPeriIg.Fields(1)
            End If
           Exit For
          End If
        Next i
        Debug.Print "5"
       End If
      Loop
      
      If tdbcMes.BoundText = "00" Then
        SumPerAnt_DebeSoles = 0
        SumPerAnt_HaberSoles = 0
      End If
      
      If .EOF Then
        printl (Space(92 + nTipo) & "------------- --------------")
        sSql = "spCn_RptFormato0501_SumPeriodos '" & Tipo & "','" & gsEmpresa & "','" & gsAnio & "','" & gsLdMesIni & "','%','','" & gsLdFechIni & "','" & _
        gsLdFechFin & "','" & gsCodMoneda & "','" & gsCtaIni & "','" & gsCtaFin & "','" & Format(Val(pPer_cPeriodo) - 1, "00") & "'"
        
        If rsSumPeriIg.State = 1 Then rsSumPeriIg.Close
        rsSumPeriIg.Open sSql, gcnSistemaAdv, adOpenDynamic, adLockOptimistic
        If rsSumPeriIg.State <> 0 Then
            SumPerAnt_DebeSoles = IIf(IsNull(rsSumPeriIg.Fields(0)), 0, rsSumPeriIg.Fields(0))
            SumPerAnt_HaberSoles = IIf(IsNull(rsSumPeriIg.Fields(1)), 0, rsSumPeriIg.Fields(1))
        End If
        
        RSet pTotalGenDebe = Format$(CStr(CDbl(pTotalMesDebe) + CDbl(SumPerAnt_DebeSoles)), "#,###,###,##0.00;(#,###,###,##0.00)")
        RSet pTotalGenHaber = Format$(CStr(CDbl(pTotalMesHaber) + CDbl(SumPerAnt_HaberSoles)), "#,###,###,##0.00;(#,###,###,##0.00)")
        printl (Space(57 + nTipo) & "TOTALES " & Space(26) & pTotalGenDebe & Space(1) & pTotalGenHaber)
        
        Exit Do
      End If
    Loop
    pPer_cPeriodo = ""
   End If
 End With
Desconectar
 Screen.MousePointer = vbDefault

If frmFCImpresion.List_Destino.Text = "Archivo" Then
   Close #1
   frmFCVistaInforme.Caption = "Libro Diario"
   frmFCVistaInforme.txtInforme.filename = frmFCImpresion.OutputFileName

   frmFCVistaInforme.Show
Else
   giLineas = 0
   Printer.FontName = "Draft 17cpi"
   Printer.FontSize = 10
   Printer.EndDoc
End If


Exit Sub

Control:
 Screen.MousePointer = vbDefault
 Desconectar
 MsgBox Err.Description
 Resume
End Sub
Private Function ExistenDatos() As Boolean
On Error GoTo Error_cmd

Dim sSql As String

 Set rsArreglo = New ADODB.Recordset
 Set rsSumPeriIg = New ADODB.Recordset
 
 SumPerAnt_DebeSoles = 0: SumPerAnt_HaberSoles = 0
 
 ConectarAdvance
 'Cambiar antes de compilar
 sSql = "spCn_RptFormato0501 '" & Tipo & "','" & gsEmpresa & "','" & gsAnio & "','" & gsLdMesIni & "','%','','" & gsLdFechIni & "','" & _
        gsLdFechFin & "','" & gsCodMoneda & "','" & gsCtaIni & "','" & gsCtaFin & "','" & gsLdMesFin & "', '" & IIf(OptTipo(0).Value, 0, 1) & "'"

 rsArreglo.Open sSql, gcnSistemaAdv, adOpenDynamic, adLockOptimistic
 SumPerAnt_DebeSoles = 0
 SumPerAnt_HaberSoles = 0
 If Trim(gsLdMesIni) = Trim(gsLdMesFin) And Trim(gsLdMesIni) <> "00" Then
  If (CDbl(gsLdMesFin) - 1) < 10 Then
   xPer_cPeriodo = "0" & Trim(Str(CDbl(gsLdMesFin) - 1))
  Else
   xPer_cPeriodo = Str(CDbl(gsLdMesFin) - 1)
  End If
  sSql = "spCn_RptFormato0501_SumPeriodos '" & Tipo & "','" & gsEmpresa & "','" & gsAnio & "','" & gsLdMesIni & "','%','','" & gsLdFechIni & "','" & _
         gsLdFechFin & "','" & gsCodMoneda & "','" & gsCtaIni & "','" & gsCtaFin & "','" & LTrim(RTrim(xPer_cPeriodo)) & "'"
         
  rsSumPeriIg.Open sSql, gcnSistemaAdv, adOpenDynamic, adLockOptimistic
  If rsSumPeriIg.State <> 0 Then SumPerAnt_DebeSoles = IIf(IsNull(rsSumPeriIg.Fields(0)), 0, rsSumPeriIg.Fields(0)): SumPerAnt_HaberSoles = IIf(IsNull(rsSumPeriIg.Fields(1)), 0, rsSumPeriIg.Fields(1))
  
 End If
 
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

Public Sub CabeceraLibroDiario(nTipo As Integer)
On Error GoTo Control
EntroCabecera = True

Dim sPag As String
Dim Anio As String
Dim Mes As String
Dim sUSUARIO As String * 10
Dim VarTitulo As String

 Gs_HoraServ = DevuelveHoraServidor
 LSet sUSUARIO = gsUsuario

 If Gs_TamPapel = 39 Then nAncho = 204 + nTipo Else nAncho = 131 + nTipo '142
 sPag = Space(4)
  
 gsConTotalPaginas = gsConTotalPaginas + 1
 gsPagina = gsPagina + 1
 RSet sPag = Format(gsPagina + 1, "####")
 giLineas = 0
 
    Call AlinearDosTextos(nAncho - 11, "   Formato 5.1: LIBRO DIARIO", "Fecha : " & Format(FechaServidor, "dd/MM/yyyy"))
    'printl (Space(3) & "Formato 5.1: LIBRO DIARIO" & Space(74 + 28) & "Fecha :  " & Format(FechaServidor, "dd/MM/yyyy"))
    
    If gsLdMesIni = gsLdMesFin Then Mes = NombreMes(gsLdMesIni) Else Mes = NombreMes(gsLdMesIni) & " A " & NombreMes(gsLdMesFin)
    If rsArreglo.EOF Then
        Call AlinearDosTextos(nAncho, "   EJERCICIO/PERIODO    : " & NombreMes(pPer_cPeriodo) & " " & gsAnio, "")
    Else
        If tdbcMes.BoundText = "00" Then
            VarTitulo = "APERTURA"
        ElseIf tdbcMes.BoundText = "13" Then
            VarTitulo = "AJUSTE"
        ElseIf tdbcMes.BoundText = "14" Then
            VarTitulo = "CIERRE"
        Else
            VarTitulo = NombreMes(tdbcMes.BoundText)
        End If
        
        If tdbcMes.BoundText <> tdbcMesFin.BoundText Then
            If tdbcMesFin.BoundText = "00" Then
                VarTitulo = VarTitulo & " A APERTURA"
            ElseIf tdbcMesFin.BoundText = "13" Then
                VarTitulo = VarTitulo & " A AJUSTE"
            ElseIf tdbcMesFin.BoundText = "14" Then
                VarTitulo = VarTitulo & " A CIERRE"
            Else
                VarTitulo = VarTitulo & " A " & NombreMes(tdbcMesFin.BoundText)
            End If
        End If
        
        VarTitulo = VarTitulo & " DEL " & gsAnio
        Call AlinearDosTextos(nAncho - 9, "   EJERCICIO/PERIODO    : " & VarTitulo, "")
        
    End If
    Dim xgsPagina As String * 4
    RSet xgsPagina = Format(CStr(gsPagina), "####")
    
    Call AlinearDosTextos(nAncho - 17, "   RUC                  : " & gsRUC, "Pagina: " & xgsPagina)
    'printl (Space(3) & "RUC                  : " & gsRUC & Space(65 + 28) & "Pagina: " & xgsPagina)

    Call AlinearDosTextos(nAncho - 18, "   APELLIDOS Y NOMBRES,", "")
    Call AlinearDosTextos(nAncho - 18, "   DENOMINACIÓN O", "")
    Call AlinearDosTextos(nAncho - 18, "   RAZON SOCIAL         : " & gsEmpresaNom, "")

    Call AlinearDosTextos(nAncho, "   MONEDA               : " & IIf(gsCodMoneda = gsMonedaNac, gsNombreMonedaNac, gsNombreMonedaExt), "")
    printl ("")
    If nTipo = 0 Then
    printl ("   ---------- ---------- ------------------------------- -------------------- ------------ -----------------------------")
    printl ("      NUM.       FECHA              GLOSA O                 DOCUMENTO         CTA.CONTABLE               MOVIMIENTO")
    printl ("     CORREL       DE             DESCRIPCION DE          TD SER. NUMERO       CODIGO                 DEBE          HABER")
    printl ("    COD.OPER.  OPERACION          LA OPERACION")
    printl ("   ---------- ---------- ------------------------------- -- ---- ------------ ------------ -------------- --------------")
    '           0609000003.03/09/2020.COMPRA OTROS                   .01.F032.00004455    .61325301    .          0.00         175.00
    Else
    printl ("   ---------- --------------------------- ---------- ------------------------------- -------------------- ------------ -----------------------------")
    printl ("      NUM.         CODIGO DE ANOTACION       FECHA              GLOSA O                 DOCUMENTO         CTA.CONTABLE               MOVIMIENTO")
    printl ("     CORREL            DE REGISTRO            DE             DESCRIPCION DE          TD SER. NUMERO       CODIGO                 DEBE          HABER")
    printl ("    COD.OPER.                              OPERACION          LA OPERACION")
    printl ("   ---------- --------------------------- ---------- ------------------------------- -- ---- ------------ ------------ -------------- --------------")
    '           0609000003.2010134030001F0320000004455.03/09/2020.COMPRA OTROS                   .01.F032.00004455    .61325301    .          0.00         175.00
    End If
Exit Sub

Control:
 MsgBox Err.Description
 Resume
End Sub
