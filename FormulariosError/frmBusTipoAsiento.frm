VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmBusTipoAsiento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo de Asientos"
   ClientHeight    =   6390
   ClientLeft      =   2025
   ClientTop       =   4305
   ClientWidth     =   8400
   Icon            =   "frmBusTipoAsiento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6390
   ScaleWidth      =   8400
   Begin VB.Frame fraAsientoTipo 
      Height          =   6300
      Left            =   45
      TabIndex        =   25
      Top             =   0
      Width           =   8280
      Begin VB.Frame fra02 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Datos Complementarios"
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
         Height          =   1950
         Left            =   90
         TabIndex        =   46
         Top             =   2700
         Width           =   8070
         Begin VB.CheckBox ChkRegAux 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            Caption         =   "Reg. Aux."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   6645
            TabIndex        =   57
            Top             =   375
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   1185
         End
         Begin TDBText6Ctl.TDBText tdbtSerDocRef 
            Height          =   285
            Left            =   2160
            TabIndex        =   13
            Tag             =   "enabled"
            ToolTipText     =   " Serie del documento de referencia "
            Top             =   1110
            Width           =   555
            _Version        =   65536
            _ExtentX        =   979
            _ExtentY        =   503
            Caption         =   "frmBusTipoAsiento.frx":0ECA
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmBusTipoAsiento.frx":0F36
            Key             =   "frmBusTipoAsiento.frx":0F54
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
            Format          =   "a@"
            FormatMode      =   1
            AutoConvert     =   -1
            ErrorBeep       =   0
            MaxLength       =   20
            LengthAsByte    =   0
            Text            =   "12345"
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
         Begin TDBText6Ctl.TDBText tdbtNroDocRef 
            Height          =   285
            Left            =   3555
            TabIndex        =   14
            Tag             =   "enabled"
            ToolTipText     =   " Número del documento de referencia "
            Top             =   1110
            Width           =   990
            _Version        =   65536
            _ExtentX        =   1746
            _ExtentY        =   503
            Caption         =   "frmBusTipoAsiento.frx":0F96
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmBusTipoAsiento.frx":1002
            Key             =   "frmBusTipoAsiento.frx":1020
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
            Format          =   "a@"
            FormatMode      =   1
            AutoConvert     =   -1
            ErrorBeep       =   0
            MaxLength       =   25
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
         Begin TDBDate6Ctl.TDBDate dtpFecDocRef 
            Height          =   300
            Left            =   5580
            TabIndex        =   15
            Tag             =   "enabled"
            ToolTipText     =   " Fecha del documento"
            Top             =   1110
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   529
            Calendar        =   "frmBusTipoAsiento.frx":1062
            Caption         =   "frmBusTipoAsiento.frx":1164
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmBusTipoAsiento.frx":11C8
            Keys            =   "frmBusTipoAsiento.frx":11E6
            Spin            =   "frmBusTipoAsiento.frx":123A
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
            Enabled         =   0
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
            Text            =   "__/__/____"
            ValidateMode    =   0
            ValueVT         =   1
            Value           =   38718
            CenturyMode     =   0
         End
         Begin TDBText6Ctl.TDBText tdbtNroConsta 
            Height          =   285
            Left            =   3555
            TabIndex        =   17
            Tag             =   "enabled"
            ToolTipText     =   " Número del documento de depósito de detracción "
            Top             =   1485
            Width           =   990
            _Version        =   65536
            _ExtentX        =   1746
            _ExtentY        =   503
            Caption         =   "frmBusTipoAsiento.frx":1262
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmBusTipoAsiento.frx":12CE
            Key             =   "frmBusTipoAsiento.frx":12EC
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
            Format          =   "a@"
            FormatMode      =   1
            AutoConvert     =   -1
            ErrorBeep       =   0
            MaxLength       =   20
            LengthAsByte    =   0
            Text            =   "1234567890"
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
         Begin TDBDate6Ctl.TDBDate dtpFecDeposito 
            Height          =   300
            Left            =   5580
            TabIndex        =   18
            Tag             =   "enabled"
            ToolTipText     =   " Fecha de depósito de detracción "
            Top             =   1485
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   529
            Calendar        =   "frmBusTipoAsiento.frx":132E
            Caption         =   "frmBusTipoAsiento.frx":1430
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmBusTipoAsiento.frx":1494
            Keys            =   "frmBusTipoAsiento.frx":14B2
            Spin            =   "frmBusTipoAsiento.frx":1506
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   16777215
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            CursorPosition  =   0
            DataProperty    =   0
            DisplayFormat   =   "dd/mm/yyyy"
            EditMode        =   0
            Enabled         =   0
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
            Text            =   "__/__/____"
            ValidateMode    =   0
            ValueVT         =   1
            Value           =   38974
            CenturyMode     =   0
         End
         Begin TDBText6Ctl.TDBText tdbtTDRef 
            Height          =   285
            Left            =   1260
            TabIndex        =   12
            Tag             =   "enabled"
            ToolTipText     =   " Tipo del documento de referencia "
            Top             =   1110
            Width           =   270
            _Version        =   65536
            _ExtentX        =   476
            _ExtentY        =   503
            Caption         =   "frmBusTipoAsiento.frx":152E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmBusTipoAsiento.frx":159A
            Key             =   "frmBusTipoAsiento.frx":15B8
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
            AlignHorizontal =   2
            AlignVertical   =   0
            MultiLine       =   0
            ScrollBars      =   0
            PasswordChar    =   ""
            AllowSpace      =   -1
            Format          =   "a"
            FormatMode      =   1
            AutoConvert     =   -1
            ErrorBeep       =   0
            MaxLength       =   2
            LengthAsByte    =   0
            Text            =   "12"
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
         Begin TDBText6Ctl.TDBText tdbtComprobante 
            Height          =   285
            Left            =   1260
            TabIndex        =   16
            Tag             =   "enabled"
            ToolTipText     =   " Número de comprobante de pago emitido por sujeto no domiciliado "
            Top             =   1485
            Width           =   990
            _Version        =   65536
            _ExtentX        =   1746
            _ExtentY        =   503
            Caption         =   "frmBusTipoAsiento.frx":15FA
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmBusTipoAsiento.frx":1666
            Key             =   "frmBusTipoAsiento.frx":1684
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
            Format          =   "a@"
            FormatMode      =   1
            AutoConvert     =   -1
            ErrorBeep       =   0
            MaxLength       =   25
            LengthAsByte    =   0
            Text            =   "1234567890"
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
         Begin MSDataListLib.DataCombo tdbcBaseImponible 
            Height          =   300
            Left            =   1260
            TabIndex        =   10
            Top             =   360
            Width           =   4320
            _ExtentX        =   7620
            _ExtentY        =   529
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin TrueOleDBList70.TDBCombo tdbcPagoIGV 
            Height          =   300
            Left            =   1575
            TabIndex        =   11
            Tag             =   "_"
            Top             =   765
            Width           =   4005
            _ExtentX        =   7064
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
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).DataField=   ""
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   5
            Splits(0)._UserFlags=   0
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).AllowRowSizing=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=5"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=609"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=529"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=370"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=291"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(12)=   "Column(2).Width=847"
            Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=767"
            Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(16)=   "Column(2).Visible=0"
            Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(18)=   "Column(3).Width=1376"
            Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=1296"
            Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(22)=   "Column(3).Visible=0"
            Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(24)=   "Column(4).Width=2196"
            Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=2117"
            Splits(0)._ColumnProps(27)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(28)=   "Column(4).Visible=0"
            Splits(0)._ColumnProps(29)=   "Column(4).Order=5"
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
            Enabled         =   0   'False
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
            _PropDict       =   $"frmBusTipoAsiento.frx":16C6
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
            _StyleDefs(48)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(52)  =   "Named:id=33:Normal"
            _StyleDefs(53)  =   ":id=33,.parent=0"
            _StyleDefs(54)  =   "Named:id=34:Heading"
            _StyleDefs(55)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(56)  =   ":id=34,.wraptext=-1"
            _StyleDefs(57)  =   "Named:id=35:Footing"
            _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(59)  =   "Named:id=36:Selected"
            _StyleDefs(60)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(61)  =   "Named:id=37:Caption"
            _StyleDefs(62)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(63)  =   "Named:id=38:HighlightRow"
            _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(65)  =   "Named:id=39:EvenRow"
            _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(67)  =   "Named:id=40:OddRow"
            _StyleDefs(68)  =   ":id=40,.parent=33"
            _StyleDefs(69)  =   "Named:id=41:RecordSelector"
            _StyleDefs(70)  =   ":id=41,.parent=34"
            _StyleDefs(71)  =   "Named:id=42:FilterBar"
            _StyleDefs(72)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Doc. Ref :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   90
            TabIndex        =   55
            Top             =   1140
            Width           =   750
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sistema/Regimen"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   75
            TabIndex        =   54
            Top             =   780
            Width           =   1455
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Doc. Dep:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   2790
            TabIndex        =   53
            Top             =   1530
            Width           =   750
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Dep:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   4635
            TabIndex        =   52
            Top             =   1530
            Width           =   885
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Serie:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   0
            Left            =   1620
            TabIndex        =   51
            Top             =   1140
            Width           =   480
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Doc:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   4635
            TabIndex        =   50
            Top             =   1140
            Width           =   870
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Número:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   1
            Left            =   2790
            TabIndex        =   49
            Top             =   1140
            Width           =   705
         End
         Begin VB.Label lblBaseImponible 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Base :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   90
            TabIndex        =   48
            ToolTipText     =   " Tipo  de Base "
            Top             =   405
            Width           =   495
         End
         Begin VB.Label lblComp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nro.Comp:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   90
            TabIndex        =   47
            Top             =   1530
            Width           =   870
         End
      End
      Begin VB.Frame fraTC 
         Caption         =   "Tipos de Cambio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   90
         TabIndex        =   35
         Top             =   4770
         Width           =   8055
         Begin VB.Frame fraTCTrib 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   705
            Left            =   4830
            TabIndex        =   38
            Top             =   180
            Width           =   1815
            Begin TDBNumber6Ctl.TDBNumber tdbnTCCompra 
               Height          =   300
               Left            =   840
               TabIndex        =   39
               Top             =   30
               Width           =   900
               _Version        =   65536
               _ExtentX        =   1587
               _ExtentY        =   529
               Calculator      =   "frmBusTipoAsiento.frx":174D
               Caption         =   "frmBusTipoAsiento.frx":176D
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmBusTipoAsiento.frx":17D9
               Keys            =   "frmBusTipoAsiento.frx":17F7
               Spin            =   "frmBusTipoAsiento.frx":183F
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483624
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "##,###,##0.000"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "##,###,##0.000"
               HighlightText   =   1
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99999999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   -1
               Separator       =   ","
               ShowContextMenu =   1
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber tdbnTCVenta 
               Height          =   300
               Left            =   840
               TabIndex        =   40
               Top             =   360
               Width           =   900
               _Version        =   65536
               _ExtentX        =   1587
               _ExtentY        =   529
               Calculator      =   "frmBusTipoAsiento.frx":1867
               Caption         =   "frmBusTipoAsiento.frx":1887
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmBusTipoAsiento.frx":18F3
               Keys            =   "frmBusTipoAsiento.frx":1911
               Spin            =   "frmBusTipoAsiento.frx":1959
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483624
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "##,###,##0.000"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "##,###,##0.000"
               HighlightText   =   1
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99999999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   -1
               Separator       =   ","
               ShowContextMenu =   1
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Compra"
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
               Index           =   8
               Left            =   30
               TabIndex        =   42
               Top             =   60
               Width           =   675
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Venta"
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
               Index           =   9
               Left            =   60
               TabIndex        =   41
               Top             =   390
               Width           =   465
            End
         End
         Begin VB.CheckBox chkTipoCambio 
            Alignment       =   1  'Right Justify
            Caption         =   "Utilizar TC. Tributario"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000001&
            Height          =   285
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Value           =   1  'Checked
            Width           =   2595
         End
         Begin TDBNumber6Ctl.TDBNumber tdbnTCVentaP 
            Height          =   300
            Left            =   3810
            TabIndex        =   37
            Top             =   210
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   529
            Calculator      =   "frmBusTipoAsiento.frx":1981
            Caption         =   "frmBusTipoAsiento.frx":19A1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmBusTipoAsiento.frx":1A0D
            Keys            =   "frmBusTipoAsiento.frx":1A2B
            Spin            =   "frmBusTipoAsiento.frx":1A73
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483624
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "##,###,##0.000"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "##,###,##0.000"
            HighlightText   =   1
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   99999999
            MinValue        =   0
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   1
            ValueVT         =   5
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin MSForms.CommandButton cmdTC 
            Height          =   375
            Left            =   6780
            TabIndex        =   23
            ToolTipText     =   " Insertar Item"
            Top             =   180
            Width           =   1215
            Caption         =   " Leer TC"
            PicturePosition =   327683
            Size            =   "2143;661"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Venta P."
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
            Index           =   7
            Left            =   3030
            TabIndex        =   36
            Top             =   270
            Width           =   675
         End
      End
      Begin TDBNumber6Ctl.TDBNumber tdbnMonto 
         Height          =   300
         Left            =   1365
         TabIndex        =   9
         Top             =   2235
         Width           =   1440
         _Version        =   65536
         _ExtentX        =   2540
         _ExtentY        =   529
         Calculator      =   "frmBusTipoAsiento.frx":1A9B
         Caption         =   "frmBusTipoAsiento.frx":1ABB
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmBusTipoAsiento.frx":1B27
         Keys            =   "frmBusTipoAsiento.frx":1B45
         Spin            =   "frmBusTipoAsiento.frx":1B8D
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
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99999999
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
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TrueOleDBList70.TDBCombo tdbcTipoEntidad 
         Height          =   300
         Left            =   1365
         TabIndex        =   3
         Tag             =   "_"
         Top             =   1455
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _LayoutType     =   0
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         _DropdownWidth  =   5821
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
         _PropDict       =   $"frmBusTipoAsiento.frx":1BB5
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
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
         _StyleDefs(20)  =   "Splits(0).Style:id=43,.parent=1"
         _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=52,.parent=4"
         _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=44,.parent=2"
         _StyleDefs(23)  =   "Splits(0).FooterStyle:id=45,.parent=3"
         _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=46,.parent=5"
         _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=48,.parent=6"
         _StyleDefs(26)  =   "Splits(0).EditorStyle:id=47,.parent=7"
         _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=49,.parent=8"
         _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=50,.parent=9"
         _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=51,.parent=10"
         _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=53,.parent=11"
         _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=54,.parent=12"
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=43"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=44"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=45"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=47"
         _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=43"
         _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=44"
         _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=45"
         _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=47"
         _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=58,.parent=43"
         _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=44"
         _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=45"
         _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=47"
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
      Begin TDBText6Ctl.TDBText tdbtNombreEntidad 
         Height          =   300
         Left            =   4725
         TabIndex        =   44
         Top             =   1440
         Width           =   3405
         _Version        =   65536
         _ExtentX        =   6006
         _ExtentY        =   529
         Caption         =   "frmBusTipoAsiento.frx":1C3C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmBusTipoAsiento.frx":1CA8
         Key             =   "frmBusTipoAsiento.frx":1CC6
         BackColor       =   16773345
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
         Format          =   ""
         FormatMode      =   1
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   120
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
      Begin TDBText6Ctl.TDBText tdbtSerie 
         Height          =   300
         Left            =   3690
         TabIndex        =   6
         Top             =   1845
         Width           =   630
         _Version        =   65536
         _ExtentX        =   1111
         _ExtentY        =   529
         Caption         =   "frmBusTipoAsiento.frx":1D08
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmBusTipoAsiento.frx":1D74
         Key             =   "frmBusTipoAsiento.frx":1D92
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
         Format          =   "a@"
         FormatMode      =   1
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   20
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
      Begin TDBText6Ctl.TDBText tdbtNumero 
         Height          =   300
         Left            =   4365
         TabIndex        =   7
         Top             =   1845
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   529
         Caption         =   "frmBusTipoAsiento.frx":1DD4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmBusTipoAsiento.frx":1E40
         Key             =   "frmBusTipoAsiento.frx":1E5E
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
         Format          =   "a@#"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   25
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
      Begin TDBText6Ctl.TDBText tdbtEntidad 
         Height          =   300
         Left            =   3720
         TabIndex        =   4
         Top             =   1440
         Width           =   960
         _Version        =   65536
         _ExtentX        =   1693
         _ExtentY        =   529
         Caption         =   "frmBusTipoAsiento.frx":1EA0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmBusTipoAsiento.frx":1F0C
         Key             =   "frmBusTipoAsiento.frx":1F2A
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
         Format          =   "@"
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
      Begin TDBText6Ctl.TDBText tdbtGlosa 
         Height          =   300
         Left            =   1365
         TabIndex        =   1
         Top             =   645
         Width           =   6720
         _Version        =   65536
         _ExtentX        =   11853
         _ExtentY        =   529
         Caption         =   "frmBusTipoAsiento.frx":1F6C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmBusTipoAsiento.frx":1FD8
         Key             =   "frmBusTipoAsiento.frx":1FF6
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
         MaxLength       =   150
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
      Begin TrueOleDBList70.TDBCombo tdbcTD 
         Height          =   300
         Left            =   4545
         TabIndex        =   24
         Tag             =   "_"
         Top             =   2295
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   529
         _LayoutType     =   0
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         _DropdownWidth  =   9710
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
         _PropDict       =   $"frmBusTipoAsiento.frx":2038
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HC0C0FF&,.bold=0,.fontsize=825"
         _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
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
         _StyleDefs(20)  =   "Splits(0).Style:id=43,.parent=1"
         _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=52,.parent=4"
         _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=44,.parent=2"
         _StyleDefs(23)  =   "Splits(0).FooterStyle:id=45,.parent=3"
         _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=46,.parent=5"
         _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=48,.parent=6"
         _StyleDefs(26)  =   "Splits(0).EditorStyle:id=47,.parent=7"
         _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=49,.parent=8"
         _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=50,.parent=9"
         _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=51,.parent=10"
         _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=53,.parent=11"
         _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=54,.parent=12"
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=43"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=44"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=45"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=47"
         _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=43"
         _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=44"
         _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=45"
         _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=47"
         _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=58,.parent=43"
         _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=44"
         _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=45"
         _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=47"
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
      Begin TDBText6Ctl.TDBText tdbtTipDocRef 
         Height          =   285
         Left            =   1365
         TabIndex        =   5
         Tag             =   "enabled"
         ToolTipText     =   " Tipo del documento de referencia "
         Top             =   1860
         Width           =   450
         _Version        =   65536
         _ExtentX        =   794
         _ExtentY        =   503
         Caption         =   "frmBusTipoAsiento.frx":20BF
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmBusTipoAsiento.frx":212B
         Key             =   "frmBusTipoAsiento.frx":2149
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
         AlignHorizontal =   2
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   "a"
         FormatMode      =   1
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   2
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
      Begin TDBText6Ctl.TDBText tdbtDesCCosto 
         Height          =   285
         Left            =   2610
         TabIndex        =   43
         Top             =   1035
         Width           =   5505
         _Version        =   65536
         _ExtentX        =   9710
         _ExtentY        =   503
         Caption         =   "frmBusTipoAsiento.frx":218B
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmBusTipoAsiento.frx":21F7
         Key             =   "frmBusTipoAsiento.frx":2215
         BackColor       =   16773345
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   0
         ShowContextMenu =   -1
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
      Begin TDBText6Ctl.TDBText tdbtCCosto 
         Height          =   285
         Left            =   1365
         TabIndex        =   2
         Top             =   1035
         Width           =   1230
         _Version        =   65536
         _ExtentX        =   2170
         _ExtentY        =   503
         Caption         =   "frmBusTipoAsiento.frx":2259
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmBusTipoAsiento.frx":22C5
         Key             =   "frmBusTipoAsiento.frx":22E3
         BackColor       =   16777215
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   0
         ShowContextMenu =   -1
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
         MaxLength       =   12
         LengthAsByte    =   0
         Text            =   "123456789012"
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
      Begin TDBText6Ctl.TDBText tdbtCodTipoAsiento 
         Height          =   285
         Left            =   1365
         TabIndex        =   0
         Top             =   270
         Width           =   1230
         _Version        =   65536
         _ExtentX        =   2170
         _ExtentY        =   503
         Caption         =   "frmBusTipoAsiento.frx":2327
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmBusTipoAsiento.frx":2393
         Key             =   "frmBusTipoAsiento.frx":23B1
         BackColor       =   16777215
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   0
         ShowContextMenu =   -1
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
         MaxLength       =   12
         LengthAsByte    =   0
         Text            =   "123456789012"
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
      Begin TDBText6Ctl.TDBText tdbtDesAsientoTipo 
         Height          =   285
         Left            =   2610
         TabIndex        =   45
         Top             =   270
         Width           =   5505
         _Version        =   65536
         _ExtentX        =   9710
         _ExtentY        =   503
         Caption         =   "frmBusTipoAsiento.frx":23F5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmBusTipoAsiento.frx":2461
         Key             =   "frmBusTipoAsiento.frx":247F
         BackColor       =   16773345
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   0
         ShowContextMenu =   -1
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
      Begin TrueOleDBList70.TDBCombo tdbcTDRef 
         Height          =   300
         Left            =   5625
         TabIndex        =   56
         Tag             =   "_"
         Top             =   2295
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   529
         _LayoutType     =   0
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         _DropdownWidth  =   9710
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
         _PropDict       =   $"frmBusTipoAsiento.frx":24C3
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HC0C0FF&,.bold=0,.fontsize=825"
         _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
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
         _StyleDefs(20)  =   "Splits(0).Style:id=43,.parent=1"
         _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=52,.parent=4"
         _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=44,.parent=2"
         _StyleDefs(23)  =   "Splits(0).FooterStyle:id=45,.parent=3"
         _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=46,.parent=5"
         _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=48,.parent=6"
         _StyleDefs(26)  =   "Splits(0).EditorStyle:id=47,.parent=7"
         _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=49,.parent=8"
         _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=50,.parent=9"
         _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=51,.parent=10"
         _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=53,.parent=11"
         _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=54,.parent=12"
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=43"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=44"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=45"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=47"
         _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=43"
         _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=44"
         _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=45"
         _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=47"
         _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=58,.parent=43"
         _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=44"
         _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=45"
         _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=47"
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
      Begin TDBDate6Ctl.TDBDate dtpFecha 
         Height          =   300
         Left            =   6840
         TabIndex        =   8
         Tag             =   "enabled"
         Top             =   1845
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   529
         Calendar        =   "frmBusTipoAsiento.frx":254A
         Caption         =   "frmBusTipoAsiento.frx":264C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmBusTipoAsiento.frx":26B0
         Keys            =   "frmBusTipoAsiento.frx":26CE
         Spin            =   "frmBusTipoAsiento.frx":273A
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
      Begin MSForms.CommandButton cmdCrearEntidad 
         Height          =   390
         Left            =   5430
         TabIndex        =   20
         ToolTipText     =   "Permite ingresar o actualizar el maestro de entidades"
         Top             =   5805
         Width           =   1305
         Caption         =   " Entidad"
         PicturePosition =   327683
         Size            =   "2302;688"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdInsertar 
         Height          =   390
         Left            =   4050
         TabIndex        =   19
         ToolTipText     =   " Insertar Item"
         Top             =   5805
         Width           =   1305
         Caption         =   " Insertar"
         PicturePosition =   327683
         Size            =   "2302;688"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdSalir 
         Height          =   390
         Left            =   6810
         TabIndex        =   21
         ToolTipText     =   "Permite ingresar o actualizar el maestro de entidades"
         Top             =   5805
         Width           =   1305
         Caption         =   " Salir"
         PicturePosition =   327683
         Size            =   "2302;688"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Centro Costo"
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
         Left            =   105
         TabIndex        =   34
         Top             =   1095
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Asiento Tipo"
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
         Left            =   105
         TabIndex        =   33
         Top             =   285
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Monto"
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
         Left            =   135
         TabIndex        =   32
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Doc"
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
         Left            =   105
         TabIndex        =   31
         Top             =   1890
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Numero Doc."
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
         Left            =   2520
         TabIndex        =   30
         Top             =   1890
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Entidad"
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
         Index           =   4
         Left            =   105
         TabIndex        =   29
         Top             =   1485
         Width           =   1035
      End
      Begin VB.Label Label2 
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
         Index           =   5
         Left            =   2910
         TabIndex        =   28
         Top             =   1485
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Doc"
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
         Left            =   5805
         TabIndex        =   27
         Top             =   1875
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Glosa"
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
         Index           =   6
         Left            =   105
         TabIndex        =   26
         Top             =   670
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmBusTipoAsiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : frmBusTipoAsiento
'    Project    : Contabilidad
'
'    Description: Formulario de ingreso de asientos tipos al mantenimiento del voucher
'--------------------------------------------------------------------------------
Option Explicit
Dim lrsProvision As ADODB.Recordset
Public frmOrigen As Form
Public tabla As String
Public auxiliar As String
Public enUso As Boolean
Public nDigitos As Integer

Dim lControl As String

Public NombreOrigen As String
Public NombreBuscador As String
Public VarLibCom As Boolean
Public FechaReg As Date
Public VarRst As New ADODB.Recordset
Dim lrsDocBolVta As ADODB.Recordset

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       validarDatos
' Description:       Valida los datos antes de enviar al mantenimiento de voucher
'
' Parameters :
'--------------------------------------------------------------------------------
Private Function LlenaRsTDBolVenta() As Boolean
    LlenaRsTDBolVenta = False
    
    Dim clDatos As clsMantoTablas
    Set clDatos = New clsMantoTablas
    
    Dim arrDatos() As Variant
    Dim sqlSp As String
    sqlSp = "SELECT COD_CVALORPARAM FROM CND_CONFIG_OPERA WHERE EMP_CCODIGO= '" & gsEmpresa & "' AND PAN_CANIO='" & gsAnio & "' AND COP_CCODIGO='028'"
    arrDatos = Array(sqlSp)
    
    
    If Not lrsDocBolVta Is Nothing Then
        If lrsDocBolVta.State = 1 Then
            lrsDocBolVta.Close
        End If
    End If
    
    Set lrsDocBolVta = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    
    If Not lrsDocBolVta Is Nothing Then
        If lrsDocBolVta.RecordCount > 0 Then LlenaRsTDBolVenta = True
    End If
    
    Set clDatos = Nothing
    
End Function

Private Function BuscaRsBolVenta(td As String) As Boolean
    BuscaRsBolVenta = False
    
    If td = "03" Then GoTo ir
    If CE(td) <> "" Then
        If Not lrsDocBolVta Is Nothing Then
            If Not (lrsDocBolVta.EOF And lrsDocBolVta.BOF) Then
                lrsDocBolVta.MoveFirst
                lrsDocBolVta.Find "COD_CVALORPARAM = " & "'" & td & "'"
                
                If lrsDocBolVta.EOF Then
                    BuscaRsBolVenta = False
                Else
ir:
                    BuscaRsBolVenta = True
                End If
            End If
        Else
            BuscaRsBolVenta = True
        End If
    End If
End Function

Private Function validarDatos() As Boolean
    validarDatos = False
    
     If tdbtCodTipoAsiento.Text = "" Then
        Mensajes "Ingrese el codigo del asiento tipo", vbOKOnly + vbInformation
        pSetFocus tdbtCodTipoAsiento
        Exit Function
    End If
    
    If tdbtCCosto.Enabled And tdbtCCosto.Text = "" Then
        Mensajes "Seleccione un centro de costo", vbOKOnly + vbInformation
        pSetFocus tdbtCCosto
        Exit Function
    End If
    
    If Not IsDate(dtpFecha.Value) Then
        Mensajes "Fecha invalida", vbOKOnly + vbInformation
        pSetFocus dtpFecha
        Exit Function
    End If
    
    
    If CE(tdbtGlosa) = "" Then
        Mensajes "Ingrese una glosa al tipo de asiento seleccionado", vbOKOnly + vbInformation
        pSetFocus tdbtGlosa
        Exit Function
    End If
    
    If CE(tdbcTipoEntidad.BoundText) = "" Then
        Mensajes "Seleccione un tipo de entidad", vbOKOnly + vbInformation
        pSetFocus tdbcTipoEntidad
        Exit Function
    End If
    
    
    If CE(tdbcTipoEntidad.BoundText) = "" And CE(tdbtNombreEntidad.Text) <> "" Then
        Mensajes "Seleccione un tipo de entidad", vbOKOnly + vbInformation
        pSetFocus tdbcTipoEntidad
        Exit Function
    End If
    
    
    If CE(tdbcTipoEntidad.BoundText) <> "" And CE(tdbtNombreEntidad.Text) = "" Then
        Mensajes "Ingrese un codigo de entidad", vbOKOnly + vbInformation
        pSetFocus tdbtEntidad
        Exit Function
    End If
    
    
    If CE(tdbtTipDocRef.Text) = "" Then
        Mensajes "Ingrese el tipo de documento", vbOKOnly + vbInformation
        pSetFocus tdbtTipDocRef
        Exit Function
    End If
    
    If CE(tdbtSerie.Text) = "" Then
        Mensajes "Ingrese la serie del documento ", vbOKOnly + vbInformation
        pSetFocus tdbtSerie
        Exit Function
    End If
    
    If CE(tdbtNumero.Text) = "" Then
        Mensajes "Ingrese el número del documento ", vbOKOnly + vbInformation
        pSetFocus tdbtNumero
        Exit Function
    End If
    

    

    
    If tdbtCodTipoAsiento.Text = "" Then
        Mensajes "Ingrese un tipo de asiento automatico", vbOKOnly + vbInformation
        pSetFocus tdbtCodTipoAsiento
        Exit Function
    End If
    
    If Not fValidEntidad Then Exit Function

    If ExisteDocumentoProvi(tdbcTipoEntidad.BoundText, tdbtEntidad.Text, _
                            tdbcTD.BoundText, tdbtSerie.Text, tdbtNumero.Text) = True Then
                            
        Mensajes "Este documento ha sido provisionado, " & Salto(1) & "ingreselo directamente en el voucher" & Salto(1) & "para cancelarlo", vbOKOnly + vbInformation
        tdbtNumero.Text = ""
        pSetFocus tdbtNumero
        Exit Function

    End If


   If frmManAsientosContables.BuscaCorrelProvisionDetalle("", tdbtEntidad.Text, tdbtTipDocRef.Text, tdbtSerie.Text, tdbtNumero.Text) = True Then
        Mensajes "Este documento ya esta siendo utilizado en el voucher"
        tdbtNumero.Text = ""
        pSetFocus tdbtNumero
        Exit Function

   End If
    
    If tdbcPagoIGV.BoundText = "" Then
        '
    ElseIf tdbcPagoIGV.BoundText = "D" Then
        'If CE(tdbtNroConsta.Text) = "" Then Mensajes "Ingrese el numero de constancia de deposito de detraccion": pSetFocus tdbtNroConsta: Exit Function
        'If IsNull(FE(dtpFecDeposito.Value)) Then Mensajes "Ingrese la fecha de deposito de detraccion": pSetFocus dtpFecDeposito: Exit Function
    Else
        If CE(tdbtTDRef.Text) = "" Then Mensajes "Ingrese el tipo de documento de referencia": pSetFocus tdbtTDRef: Exit Function
        If CE(tdbtSerDocRef.Text) = "" Then Mensajes "Ingrese la serie del documento de referencia": pSetFocus tdbtSerDocRef: Exit Function
        If CE(tdbtNroDocRef.Text) = "" Then Mensajes "Ingrese el numero de documento de referencia": pSetFocus tdbtNroDocRef: Exit Function
        If IsNull(FE(dtpFecDocRef.Value)) Then Mensajes "Ingrese la fecha del documento de referencia": pSetFocus dtpFecDocRef: Exit Function
    End If
    
    validarDatos = True
End Function

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       CargaFormularioTC
' Description:       Muestra el formulario de tipos de cambio
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub CargaFormularioTC()
    Mensajes "No se ingreso tipo de cambio para la fecha " & dtpFecha.Value & ",  Ingrese Tipo de Cambio de la fecha...", vbInformation
    
    Me.Enabled = False 'desactivo el formulario de reg auxiliares
    frmManTipoCambio.Show
    DoEvents
    frmManTipoCambio.SSTCentroCosto.Tab = 0
    frmManTipoCambio.tdbcMes(0).BoundText = Right("00" & CStr(Month(dtpFecha.Value)), 2)
    
    DoEvents
    frmManTipoCambio.ConfigurarControlFecha (0)
    DoEvents
    
    frmManTipoCambio.dtpFechaBus(0).Value = dtpFecha.Value
    frmManTipoCambio.chkFecha(0).Value = vbChecked
    DoEvents
    frmManTipoCambio.ManNuevo
    frmManTipoCambio.dtpFecha.Value = dtpFecha.Value
    pSetFocus frmManTipoCambio.tdbnCompra
    frmManTipoCambio.Asientos = True

End Sub


'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       chkTipoCambio_Click
' Description:       Evento que se ejecuta al hacer clic en el check de tipo de cambio
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub chkTipoCambio_Click()
    If chkTipoCambio.Value = vbChecked Then
        fraTCTrib.Visible = True
    Else
        fraTCTrib.Visible = False
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       cmdCrearEntidad_Click
' Description:       Evento que se ejecuta al hacer clic en el boton de entidad
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub cmdCrearEntidad_Click()
    ' *** Llamar al formulario de entidad; para registrarlo
    cmdCrearEntidad.Enabled = False
    DoEvents
    Me.Enabled = False
    frmManEntidades.Show
    frmManEntidades.automatico = True
    pSetFocus frmManEntidades
    pSendKeys "{F2}"
    frmManEntidades.tdbcTipoEntidad.BoundText = tdbcTipoEntidad.BoundText
    pSendKeys "{Enter}"
    ' ***
    cmdCrearEntidad.Enabled = True
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Insertar
' Description:       Procedimiento de enviar los datos ingresados al mantenimiento de voucher
'
' Parameters :
'--------------------------------------------------------------------------------
Public Sub Insertar()
    Call cmdInsertar_Click
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       ValidaTC
' Description:       Procedimiento que valida el tipo de cambio si el movimiento es de moneda extranjera
'
' Parameters :
'--------------------------------------------------------------------------------
Private Function ValidaTC() As Boolean
    ValidaTC = False
    Dim nTC As Double
    nTC = tdbnTCCompra + tdbnTCVenta + tdbnTCVentaP
    
    If nTC = 0 Then
        If MsgBox("No hay tipos de cambio, desea continuar", vbYesNo + vbQuestion) = vbYes Then
            ValidaTC = True
        End If
    Else
        ValidaTC = True
    End If
    
End Function

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       cmdInsertar_Click
' Description:       Evento que se ejecuta alhacer clic en enviar datos
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub cmdInsertar_Click()


    If validarDatos = False Then
        cmdInsertar.Enabled = True
        Exit Sub
    End If
    
    If fraTC.Visible Then
        If ValidaTC = False Then Exit Sub
    End If
    
    cmdInsertar.Enabled = False
    DoEvents
    
    
    If NombreOrigen = "frmManAsientosContables" Then
        frmManAsientosContables.Enabled = True
        frmManAsientosContables.RecibirDatos "TipoAsiento", "", "", ""
'        frmManAsientosContables.tdbgDetalle.Row = 3
'        frmManAsientosContables.tdbgDetalle.Col = 11
        cmdInsertar.Enabled = True
    
        frmManAsientosContables.Enabled = True
    Else
        frmOrigen.Enabled = True
    End If
    
    DoEvents
    
    Unload Me
    
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       cmdSalir_Click
' Description:       Evento que se ejecuta al hacer clic en salir delformulario
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub cmdsalir_Click()
Unload Me
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       cmdTC_Click
' Description:       Evento que se ejecuta al hacer clic en tipo de cambio
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub cmdTC_Click()
    cmdTC.Enabled = False
    DoEvents
    
    Me.Enabled = False
    frmManTipoCambio.Asientos = True
    
    frmManTipoCambio.Show
    frmManTipoCambio.Asientos = True
    frmManTipoCambio.tdbcMes(0).BoundText = gsPeriodo
    frmManTipoCambio.dtpFechaBus(0).Value = dtpFecha.Value
    frmManTipoCambio.tdbcMes(1).BoundText = gsPeriodo
    frmManTipoCambio.dtpFechaBus(1).Value = dtpFecha.Value
    
    Dim nTC As Double
    Dim sMoneda As String
    nTC = tdbnTCCompra + tdbnTCVenta + tdbnTCVentaP
    sMoneda = frmManAsientosContables.tdbcMoneda.BoundText

    
    If nTC = 0 Then
    
        If sMoneda = gsMonedaExt Then
            frmManTipoCambio.SSTCentroCosto.Tab = 0
        End If
        
        If sMoneda <> gsMonedaNac And sMoneda <> gsMonedaExt Then
            frmManTipoCambio.SSTCentroCosto.Tab = 1
            frmManTipoCambio.tdbcMonedaAdic.BoundText = sMoneda
        End If
        
        frmManTipoCambio.ManNuevo
        frmManTipoCambio.dtpFecha.Value = dtpFecha.Value
        
    End If
    
    cmdTC.Enabled = True
End Sub


Private Sub dtpFecDocRef_LostFocus()
    If dtpFecDocRef.Value > dtpFecha.Value And (tdbtTipDocRef.Text = "07" Or tdbtTipDocRef.Text = "08") Then
        Mensajes "¡La fecha del documento de referencia no puede ser mayor a la fecha del documento origen!", vbExclamation
        dtpFecDocRef.Value = dtpFecha.Value
        dtpFecDocRef.SetFocus
        Exit Sub
    End If
    If dtpFecDocRef.Enabled = True Then fechDocRef = IIf(IsNull(dtpFecDocRef.Value), "__/__/____", dtpFecDocRef.Value)
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       dtpFecha_KeyDown
' Description:       Evento que se ejecuta al presioanr una tecla en la fecha de documento
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
Private Sub dtpFecha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        tdbnMonto.SetFocus
    End If
    
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
       KeyCode = 0
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       dtpFecha_LostFocus
' Description:       Evento que se ejecuta al perder el enfoque en la fecha del documento
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub dtpFecha_LostFocus()

'    If VarLibCom Then
'        Dim dFechaMin As Date
'        Dim gsMesesAnteriores  As Long
'        gsMesesAnteriores = NE(BuscaConfigAnual("Cfl_cMesCompras"))
'
'        dFechaMin = DateAdd("m", gsMesesAnteriores * -1, FechaReg)
'
'        If dtpFecha.Value < dFechaMin Then
'           Mensajes "La fecha ingresada " & Format(dtpFecha.Value, "dd/MM/yyyy") & " no debe ser menor que " & CE(dFechaMin) & Salto(1) & "el rango permitido es de " & CE(gsMesesAnteriores) & " meses anteriores a la fecha del voucher"
'           dtpFecha.Value = FechaReg
'           Exit Sub
'        End If
'    End If
    
       If VarLibCom Then
           Dim dFechaAct As Date
           Dim dFechaMin As Date
        Dim gsMesesAnteriores  As Long
        gsMesesAnteriores = NE(BuscaConfigAnual("Cfl_cMesCompras"))
           
           dFechaAct = dtpFecha.Value
           dFechaMin = DateAdd("m", gsMesesAnteriores * -1, FechaReg)
           
           'PGBV - 04012013
            'If EstadoDes = "" Then
                If dtpFecha.Value >= dFechaMin Then
                    'If Month(tdbgDetalle.Columns(nColAsd_dFecDoc)) < Month(lsFecha) Or Year(tdbgDetalle.Columns(nColAsd_dFecDoc)) < Year(lsFecha) And Year(dtpFecha.Value) <= Year(lsFecha) Then
                        EstadoOri = "6"
                        If FechaReg = dtpFecha.Value Then
                            EstadoOri = "1"
                        End If
                    'End If
                ElseIf dtpFecha.Value < dFechaMin Then
                    Mensajes "El comprobante ingresado supera los " & CE(gsMesesAnteriores) & " meses (Revise parametros iniciales), se informará a SUNAT lo que implicaría una infracción tributaria. "
                    If MsgBox("Desea continuar..?", vbQuestion + vbOKCancel, gsNombreModulo) = vbOK Then
                       EstadoOri = "7"
                    Else
                        dtpFecha.Value = FechaReg
                        dtpFecha.SetFocus
                        Exit Sub
                    End If
                End If
        Else
            If Month(dtpFecha.Value) <> Month(FechaReg) And Year(dtpFecha.Value) <> Year(FechaReg) Then   'And Year(dtpFecha.Value) = Year(lsFecha) Then
                Mensajes "El comprobante ingresado corresponde a un período anterior, se informará a SUNAT lo que implicaría una infracción tributaria. "
                If MsgBox("Desea continuar..?", vbQuestion + vbOKCancel, gsNombreModulo) = vbOK Then
                    EstadoDes = "8"
                Else
                        dtpFecha.Value = FechaReg
                        dtpFecha.SetFocus
                        Exit Sub
                End If
            End If
        End If

    If IsDate(dtpFecha) Then
        CargaTC
        If tdbnMonto.Enabled = True Then tdbnMonto.SetFocus
    Else
        tdbnTCCompra = 0
        tdbnTCVenta = 0
        tdbnTCVentaP = 0
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Form_KeyDown
' Description:       Evento que se ejecuta al preisonar una tecla en el formulario
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 27 Then Unload Me
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       CargaTC
' Description:       Procedimiento que busca los tipos de cambios
'
' Parameters :
'--------------------------------------------------------------------------------
Public Sub CargaTC()
    'If fraTC.Visible Then
    Dim sqlSp As String
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
    Set clDatos = New clsMantoTablas
    Dim lrsTabla  As ADODB.Recordset
    Set lrsTabla = New ADODB.Recordset
    
    'Carga el Tipo de Cambio de Dolares
    If gintBiMoneda = 1 Then
        sqlSp = "spCn_HallarTC '" & gsEmpresa & "', '" & dtpFecha.Value & "', ''"
    Else
        sqlSp = "spCn_HallarTC '" & gsEmpresa & "', '" & dtpFecha.Value & "', '" & _
            frmManAsientosContables.tdbcMoneda.BoundText & "'"
    End If
    
    
            
    arrDatos = Array(sqlSp)
    
    Set lrsTabla = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    
    tdbnTCCompra = 0
    tdbnTCVenta = 0
    tdbnTCVentaP = 0
    
    If Not lrsTabla Is Nothing Then
        tdbnTCCompra = NE(lrsTabla("Tca_nCompra"))
        tdbnTCVenta = NE(lrsTabla("Tca_nVenta"))
        tdbnTCVentaP = NE(lrsTabla("Tca_nVentaP"))
    End If
    
    Call CerrarRecordSet(lrsTabla)
    Set clDatos = Nothing
    'End If
End Sub



Private Sub Form_Resize()

On Error GoTo errHand
    If Me.WindowState <> vbMinimized Then
        
        Call Centrar_Objeto(fraAsientoTipo, Me)

    End If
Exit Sub
errHand:

End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Form_Unload
' Description:       Evento que se ejecuta al cerrar el formulario
'
' Parameters :       Cancel (Integer)
'--------------------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
    
    If NombreOrigen = "frmManAsientosContables" Then
        frmManAsientosContables.Enabled = True
    Else
        If Not frmOrigen Is Nothing Then
            frmOrigen.Enabled = True
        End If
    End If
    
    
    
    Set frmOrigen = Nothing
    Set frmBusTipoAsiento = Nothing
    
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbcBaseImponible_KeyDown
' Description:       Evento que se ejecuta al presionar unatecla en el combo de base imponible
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbcBaseImponible_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        pSetFocus tdbcPagoIGV
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbcPagoIGV_ItemChange
' Description:       Evento que se ejecuta al cambiar la seleccion enel combo de pago de IGV
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbcPagoIGV_ItemChange()
    If tdbcPagoIGV.BoundText = "D" Then
        Call ActivarCamposDetraccion(True)
        Call LimpiarCamposReferencia
        Call ActivarCamposReferencia(False)
        
    Else
        Call LimpiarCamposDetraccion
        Call ActivarCamposDetraccion(False)
        Call ActivarCamposReferencia(True)
    End If
    
    If tdbcPagoIGV.BoundText = "" Then
        Call LimpiarCamposDetraccion
        Call ActivarCamposDetraccion(False)
        Call LimpiarCamposReferencia
        Call ActivarCamposReferencia(False)
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbcPagoIGV_KeyDown
' Description:       Evento que se ejecuta al presionar una tecla en el combo de PAgo de IGV
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbcPagoIGV_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If tdbcPagoIGV.BoundText = "D" Then
            pSetFocus tdbtComprobante
        Else
            Call LimpiarCamposDetraccion
            Call ActivarCamposDetraccion(False)
            Call ActivarCamposReferencia(True)
            
            If tdbtTDRef.Enabled Then
                pSetFocus tdbtTDRef
            Else
                pSetFocus tdbtNroConsta
            End If
        End If
        
        If tdbcPagoIGV.BoundText = "" Then
            Call LimpiarCamposDetraccion
            Call ActivarCamposDetraccion(False)
            Call LimpiarCamposReferencia
            Call ActivarCamposReferencia(False)
            
            pSetFocus tdbtComprobante
        End If
    End If
    
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbcTD_KeyDown
' Description:       Evento que se ejecuta al presioanr una tecla en el combo de tipo de documento
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbcTD_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then pSendKeys "{tab}"
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbcTD_SelChange
' Description:       Evento que se ejecuta al cambiar la seleccion del tipo de documento
'
' Parameters :       Cancel (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbcTD_SelChange(Cancel As Integer)
    If CE(tdbcTD.BoundText) = "" Then
        tdbtSerie.Text = ""
        tdbtNumero.Text = ""
    End If

End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbcTipoEntidad_KeyDown
' Description:       Evento que se ejecuta al presionar una tecla en el tipo de entidad
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbcTipoEntidad_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then pSendKeys "{tab}"
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbcTipoEntidad_SelChange
' Description:       Evento que se ejecuta al cambiar el tipo de entidad
'
' Parameters :       Cancel (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbcTipoEntidad_SelChange(Cancel As Integer)
    If CE(tdbcTipoEntidad.BoundText) = "" Then
        tdbtEntidad.Text = ""
        tdbtNombreEntidad.Text = ""
    End If

End Sub


'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtEntidad_Change
' Description:       Evento que se ejecuta al cambiar el tipo de entidad
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbtEntidad_Change()
    If CE(tdbtEntidad.Text) = "" Then
        tdbtNombreEntidad.Text = ""
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtEntidad_KeyDown
' Description:       Evento que se ejecuta al presionar una tecla en el tipo de entidad
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbtEntidad_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then
        If tdbcTipoEntidad.BoundText = "" Then
            Mensajes "Seleccione una entidad", vbOKOnly + vbInformation
            pSetFocus tdbcTipoEntidad
            Exit Sub
        End If
    
        Call LlamaBuscar(frmBuscador, Me.tdbtEntidad.Name, lControl, "Entidad", Me, gsPeriodo, Me.tdbcTipoEntidad.BoundText)
    End If
    
    
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       RecibirDatos
' Description:       Procedimiento que recibe los datos de codigo de entidades y tipos de documento
'
' Parameters :       lControl (String)
'                    param0 (String)
'                    param1 (String)
'                    param2 (String)
'--------------------------------------------------------------------------------
Public Sub RecibirDatos(ByVal lControl As String, ByVal param0 As String, ByVal param1 As String, ByVal param2 As String, Optional ByVal param3 As String)
   ' *** Dependiendo del control
    Select Case lControl
           Case "tdbtEntidad", "Entidad" ': *** Caso Desde
                tdbtEntidad = Trim(param0)
                tdbtNombreEntidad = Trim(param1)
                Unload frmBuscador
                pSetFocus tdbtEntidad
           Case "TipoDocumento"
                tdbtTipDocRef = Trim(param0)
    End Select
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtEntidad_LostFocus
' Description:       Evento que se ejecuta al perder el enfoque el codigo de entidad
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbtEntidad_LostFocus()
    If CE(tdbtEntidad.Text) = "" Then
        tdbtNombreEntidad.Text = ""
    End If

    fValidEntidad
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       fValidEntidad
' Description:       Funcion que validad el codigo de entidad digitado
'
' Parameters :
'--------------------------------------------------------------------------------
Private Function fValidEntidad() As Boolean
    Dim sqlver As String
    Dim valorDato As String
    
    fValidEntidad = False
    
    If tdbtEntidad <> "" And Me.Enabled = True Then
        sqlver = "SELECT ENT_CPERSONA From CNM_ENTIDAD " & _
                 "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND  ent_cEstado = 'A'" & _
                 "AND Ent_cCodEntidad = '" & tdbtEntidad & "' " & _
                 "AND Ten_CTipoEntidad = '" & Me.tdbcTipoEntidad.BoundText & "' "
        valorDato = ExtraeDescripcion(sqlver)
        If valorDato = "" Then
            Mensajes "Codigo no existe, verificar...", vbInformation
            tdbtEntidad.Text = ""
            pSetFocus tdbtEntidad
            Exit Function
        Else
            Me.tdbtNombreEntidad = valorDato
            'pSetFocus tdbtTipDocRef

        End If
    End If
    
    fValidEntidad = True
End Function

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtGlosa_Change
' Description:       Evento que se ejecuta al cambiar la glosa
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbtGlosa_Change()
    If gsKey = 219 Then
       tdbtGlosa = Replace(tdbtGlosa, "'", "")
       tdbtGlosa.SelStart = Len(tdbtGlosa)
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtGlosa_KeyDown
' Description:       Evento que se ejecuta al presioanr una tecla en el campo de la glosa
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbtGlosa_KeyDown(KeyCode As Integer, Shift As Integer)
    gsKey = KeyCode
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtNroDocRef_LostFocus
' Description:       Evento que se ejecuta al perder el enfoque el numero dedocumento de referencia
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbtNroDocRef_LostFocus()
    If CE(tdbtNroDocRef.Text) <> "" Then
            If Len(CE(tdbtNroDocRef.Text)) > 8 Then
                tdbtNroDocRef.Text = CE(tdbtNroDocRef.Text)
            Else
                tdbtNroDocRef.Text = Right("00000000" & tdbtNroDocRef.Text, 8)
            End If
            NumDocRef = tdbtNroDocRef.Text
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtNumero_KeyDown
' Description:       Evento que se ejecuta al presionar una tecla en el numero d documento
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbtNumero_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
       tdbtNumero_LostFocus
    End If

End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtNumero_LostFocus
' Description:       Evento que se ejecuta al perder el enfoque el numero de documento
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbtNumero_LostFocus()
    If CE(tdbtNumero) <> "" Then
            If Len(CE(tdbtNumero.Text)) > 8 Then
                tdbtNumero.Text = CE(tdbtNumero.Text)
            Else
                tdbtNumero.Text = Right("00000000" & CE(tdbtNumero.Text), 8)
            End If
    End If
    
    If CE(tdbtNumero) <> "" Then
       fValidEntidad
    End If
End Sub



'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtSerDocRef_LostFocus
' Description:       Evento que se ejecuta al perder el enfoque el numero de documento de referencia
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbtSerDocRef_LostFocus()
    If CE(tdbtSerDocRef.Text) <> "" Then
        If Len(CE(tdbtSerDocRef.Text)) > 3 Then
            tdbtSerDocRef.Text = CE(tdbtSerDocRef.Text)
        Else
            tdbtSerDocRef.Text = Right("0000" & tdbtSerDocRef.Text, 4)
        End If
        SerieDocRef = tdbtSerDocRef.Text
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtSerie_KeyDown
' Description:       Evento que se ejecuta al hacer clic en el campo de serie de documento
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbtSerie_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And CE(tdbtSerie.Text) <> "" Then
            tdbtSerie_LostFocus
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtSerie_LostFocus
' Description:       Evento que se ejecuta al perder el enfoque la serie del documento
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbtSerie_LostFocus()
    If CE(tdbtSerie) <> "" Then
            If Len(CE(tdbtSerie.Text)) > 3 Then
                tdbtSerie.Text = CE(tdbtSerie.Text)
            Else
                tdbtSerie.Text = Right("0000" & CE(tdbtSerie.Text), 4)
            End If

    End If

End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtTDRef_KeyDown
' Description:       Evento que se ejecuta al presionar una tecla el tipo de documento de referencia
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbtTDRef_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
       Call LlamaBuscar(frmBuscador, "AsientoTipoDocumento", lControl, "TipoDocumento", Me, gsPeriodo, "")
    End If

End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtTDRef_LostFocus
' Description:       Evento que se ejecuta al perder el enfoque el tipo de documento de referencia
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbtTDRef_LostFocus()
    Dim sqlver As String, valorDato As String
    
    If CE(tdbtTDRef.Text) <> "" Then tdbtTDRef.Text = Right("00" & CE(tdbtTDRef.Text), 2): tipoDocRef = tdbtTDRef.Text
    
    If Len(CE(tdbtTDRef.Text)) > 0 Then
        tdbcTDRef.BoundText = tdbtTDRef.Text
        
        If tdbcTDRef.BoundText = "" Then
           Mensajes "Codigo de Documento no existe, verificar..."
           
           tdbtTDRef.Text = ""
           
           pSetFocus tdbtTDRef
           Exit Sub
        End If
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtTipDocRef_KeyDown
' Description:       Evento que se ejecuta al presionar una tecla en el tipo de documento de referencia
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbtTipDocRef_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
       Call LlamaBuscar(frmBuscador, "AsientoTipoDocumento", lControl, "TipoDocumento", Me, gsPeriodo, "")
    End If
End Sub



'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtTipDocRef_LostFocus
' Description:       Evento que se ejecuta al perder el enfoque en el tipo de documento de referencia
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbtTipDocRef_LostFocus()

    Dim sqlver As String, valorDato As String
    
    If CE(tdbtTipDocRef.Text) <> "" Then tdbtTipDocRef.Text = Right("00" & CE(tdbtTipDocRef.Text), 2)
    
    If Len(CE(tdbtTipDocRef.Text)) > 0 Then
        tdbcTD.BoundText = tdbtTipDocRef.Text
        
        If tdbcTD.BoundText = "" Then
           Mensajes "Codigo de Documento no existe, verificar..."
           
           tdbtTipDocRef.Text = ""
           
           pSetFocus tdbtTipDocRef
           Exit Sub
        End If
    End If
    
    If CE(tdbtTipDocRef.Text) = "07" Or CE(tdbtTipDocRef.Text) = "08" Then
        tdbcPagoIGV.Bookmark = 1
    End If

    
    If BuscaRsBolVenta(tdbtTipDocRef.Text) = True Then
        ChkRegAux.Visible = True
        ChkRegAux.Value = vbChecked
    Else
        ChkRegAux.Value = vbUnchecked
        ChkRegAux.Visible = False
    End If
    
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtCCosto_Change
' Description:       Evento que se ejecuta al cambiar el codigo de centro decosto
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbtCCosto_Change()
    If CE(tdbtCCosto.Text) = "" Then
       tdbtDesCCosto.Text = ""
    End If

End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtCCosto_KeyDown
' Description:       Evento que se ejecuta al presionar el codigo de centro de costo
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbtCCosto_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim CodOld As String, CodNew As String
    If KeyCode = vbKeyF1 Then
        CodOld = tdbtCCosto.Text
        tdbtCCosto.Text = ""
        CodNew = BuscarCentroCosto(Me, tdbtCCosto.Text)
        tdbtCCosto.Text = IIf(CodNew <> "", CodNew, CodOld)
    End If

End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtCCosto_LostFocus
' Description:       Evento que se ejecuta al perder el enfoque del codigo del centro de costo
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbtCCosto_LostFocus()
    If CE(tdbtCCosto.Text) <> "" Then
        tdbtDesCCosto.Text = BuscaNombreCC(tdbtCCosto.Text, True)
        
        If CE(tdbtDesCCosto.Text) = "" Then
            tdbtCCosto.Text = ""
            pSetFocus tdbtCCosto
        End If
        
    End If

End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtCodTipoAsiento_Change
' Description:       Evento que se ejecuta al cambiar el codigo del tipo de asiento
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbtCodTipoAsiento_Change()
    If CE(tdbtCodTipoAsiento.Text) = "" Then
       tdbtDesAsientoTipo.Text = ""
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtCodTipoAsiento_KeyDown
' Description:       Evento que se ejecuta al presioanr una tecla en el codigo del tipo de asiento
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbtCodTipoAsiento_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim CodOld As String, CodNew As String
    Dim sqlSp  As String
    Dim sTipoLibro  As String
    Dim cTipoEnt As String
    
    If KeyCode = vbKeyF1 Then
        CodOld = tdbtCodTipoAsiento.Text
        tdbtCodTipoAsiento.Text = ""
        
        
        sTipoLibro = frmManAsientosContables.tdbcLibro.BoundText
        
        CodNew = BuscarAsientoTipo(sTipoLibro, Me, tdbtCodTipoAsiento.Text)
        tdbtCodTipoAsiento.Text = IIf(CodNew <> "", CodNew, CodOld)
        
'-------
        If gsCampo4 = "1" Then
            ActivarControl tdbtCCosto, True, gsColorActivado
        Else
            ActivarControl tdbtCCosto, False, gsColorDesactivado
            tdbtCCosto.Text = ""
        End If

        If gsCampo5 <> "" Then
            cTipoEnt = "and Ten_cTipoEntidad = '" & gsCampo5 & "'"
            
            sqlSp = "SELECT Ten_cTipoEntidad, Ten_cNombreEntidad From CNT_ENTIDAD " & _
                    "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND Ten_cEstado='A' " & cTipoEnt & " ORDER BY Ten_cNombreEntidad"
            LlenarComboAddItem tdbcTipoEntidad, sqlSp, True
        Else
            tdbcTipoEntidad.Clear
        End If
'-------
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtCodTipoAsiento_LostFocus
' Description:       Evento que se ejecuta al perder el enfoque del codigo de tipo de asiento
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbtCodTipoAsiento_LostFocus()
    Dim sTipoLibro  As String
       
    sTipoLibro = frmManAsientosContables.tdbcLibro.BoundText
        
    If CE(tdbtCodTipoAsiento.Text) <> "" Then
        tdbtCodTipoAsiento.Text = Right("000" + CE(tdbtCodTipoAsiento.Text), 3)
        
        tdbtDesAsientoTipo.Text = BuscaNombreAsientoTipo(sTipoLibro, tdbtCodTipoAsiento.Text)
        If CE(tdbtDesAsientoTipo.Text) = "" Then
            tdbtCodTipoAsiento.Text = ""
            pSetFocus tdbtCodTipoAsiento
        End If
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       ActivarCamposReferencia
' Description:       Procedimiento que activa los campos de referencia
'
' Parameters :       bValor (Boolean)
'--------------------------------------------------------------------------------
Private Sub ActivarCamposReferencia(bValor As Boolean)
    Dim nColor As OLE_COLOR
    If bValor = True Then
        nColor = gsColorActivado
    Else
        nColor = gsColorDesactivado
    End If
    
    ActivarControl tdbtTDRef, bValor, nColor
    ActivarControl tdbtSerDocRef, bValor, nColor
    ActivarControl tdbtNroDocRef, bValor, nColor
    ActivarControl dtpFecDocRef, bValor, nColor
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       ActivarCamposDetraccion
' Description:       Procedimiento de activacion de los campos de detraccion
'
' Parameters :       bValor (Boolean)
'--------------------------------------------------------------------------------
Private Sub ActivarCamposDetraccion(bValor As Boolean)
    Dim nColor As OLE_COLOR
    If bValor = True Then
        nColor = gsColorActivado
    Else
        nColor = gsColorDesactivado
    End If
    
    ActivarControl tdbtNroConsta, bValor, nColor
    ActivarControl dtpFecDeposito, bValor, nColor
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       ActivarNroComp
' Description:       Procedimiento que cambia de color a los campos de numero de comprobante
'
' Parameters :       bValor (Boolean)
'--------------------------------------------------------------------------------
Private Sub ActivarNroComp(bValor As Boolean)
    Dim nColor As OLE_COLOR
    If bValor = True Then
        nColor = gsColorActivado
    Else
        nColor = gsColorDesactivado
    End If
    
    ActivarControl tdbtComprobante, bValor, nColor
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       ActivarSistemaRegimen
' Description:       Procedimiento que cambia el color del combo de PAgo IGV
'
' Parameters :       bValor (Boolean)
'--------------------------------------------------------------------------------
Private Sub ActivarSistemaRegimen(bValor As Boolean)
    Dim nColor As OLE_COLOR
    If bValor = True Then
        nColor = gsColorActivado
    Else
        nColor = gsColorDesactivado
    End If
    
    ActivarControl tdbcPagoIGV, bValor, nColor
End Sub



'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       LimpiarCamposReferencia
' Description:       Procedimiento de limpiar campos de referencia
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub LimpiarCamposReferencia()
    tdbtTDRef.Text = ""
    tdbtSerDocRef.Text = ""
    tdbtNroDocRef.Text = ""
    dtpFecDocRef.Text = ""
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       LimpiarCamposDetraccion
' Description:       Procedimiento de limpiar campos de detraccion
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub LimpiarCamposDetraccion()
    tdbtNroConsta.Text = ""
    dtpFecDeposito.Text = "__/__/____"
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       LimpiarDatosComplementarios
' Description:       Procedimiento de limpiar los campos complementarios
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub LimpiarDatosComplementarios()
    tdbcBaseImponible.BoundText = ""
    tdbcPagoIGV.BoundText = ""
    Call LimpiarCamposReferencia
    Call LimpiarCamposDetraccion
    tdbtComprobante.Text = ""
    
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       ActivarDatosComplementarios
' Description:       Procedimiento de activacion de campos complementarios
'
' Parameters :       bValor (Boolean)
'--------------------------------------------------------------------------------
Private Sub ActivarDatosComplementarios(bValor As Boolean)
    Call ActivarCamposReferencia(bValor)
    Call ActivarCamposDetraccion(bValor)
    Call ActivarNroComp(bValor)
    Call ActivarSistemaRegimen(bValor)
    
    Dim nColor As OLE_COLOR
    If bValor = True Then
        nColor = gsColorActivado
    Else
        nColor = gsColorDesactivado
    End If
    
    ActivarControl tdbcBaseImponible, bValor, nColor
    ActivarControl tdbcPagoIGV, bValor, nColor
    ActivarControl tdbtComprobante, bValor, nColor
    ActivarControl tdbcBaseImponible, bValor, nColor
    
'    tdbcBaseImponible.Enabled = bValor
'    tdbcPagoIGV.Enabled = bValor
'    tdbcBaseImponible.Locked = Not bValor
'    tdbcPagoIGV.Locked = Not bValor
    
    
End Sub


'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Form_Load
' Description:       Evento que se ejecuta al iniciar el formulario
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub Form_Load()
    
    
    Call Centrar_form(Me)
    
    tdbtGlosa.Text = ""
    tdbtCCosto.Text = ""
    tdbtCodTipoAsiento.Text = ""

    Me.Caption = "Tipo de Asientos en Libros de Operaciones"
    
    Call LimpiarDatosComplementarios
    Call ActivarDatosComplementarios(True)
    Call LlenaCombos
    
    Call SeteaFechas
    Call LlenaRsTDBolVenta
    
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       SeteaFechas
' Description:       Procedimiento de seteo inicial de fechas
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub SeteaFechas()
    'On Error Resume Next
    dtpFecDeposito.MinDate = "01/01/" & CE(NE(gsAnio) - 1)
    dtpFecDeposito.MaxDate = "31/12/" & CE(NE(gsAnio) + 1)
    
    dtpFecDocRef.MinDate = "01/01/" & CE(NE(gsAnio) - 1)
    dtpFecDocRef.MaxDate = "31/12/" & CE(NE(gsAnio) + 1)
    
    'dtpFecha.MinDate = "01/01/" & CE(NE(gsAnio) - 1)
    'dtpFecha.MaxDate = "31/12/" & CE(NE(gsAnio) + 1)

End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       LlenaCombos
' Description:       Procedimiento de llenado de combos
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub LlenaCombos()
    Dim sqlSp As String

    sqlSp = "spCn_ConsultaTipDocsLibro 'SEL_DOCS_ALL_LIBRO', '" & gsEmpresa & "', '" & gsAnio & "', '" & tabla & "', '' "
    LlenarComboAddItem tdbcTD, sqlSp, True
    LlenarComboAddItem tdbcTDRef, sqlSp, True
    '--------------------------------------------------
    
    Call LlenaComboBaseImponible
    
    If frmManAsientosContables.tdbcLibro.BoundText = lsLibroCom Then
        tdbcBaseImponible.BoundText = gsBaseImpDefCom
    End If
    
    If frmManAsientosContables.tdbcLibro.BoundText = lsLibroVen Then
        tdbcBaseImponible.BoundText = "002"
    End If
    
    '--------------------------------------------------
    Call LlenaComboIGVPublico(tdbcPagoIGV, "NDPR")
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       LlenaComboBaseImponible
' Description:       Procedimiento de llenado de combo de base imponible
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub LlenaComboBaseImponible()
    Dim lrsTipo As New ADODB.Recordset
    With lrsTipo
        .CursorLocation = TEL_CURSOR_LOCATION.TEL_CURSOR_CLIENT
        .CursorType = TEL_CURSOR_TYPE.TEL_TYPE_DYNAMIC
        .LockType = TEL_LOCK_TYPES.TEL_LOCK_BATCH_OPTIMISTIC
        .Fields.Append "CODIGO", adChar, 3
        .Fields.Append "DESCRIPCION", adVarChar, 50
        .Open
        
        If frmManAsientosContables.tdbcLibro.BoundText = lsLibroCom Then
            .AddNew: lrsTipo.Fields("CODIGO") = "   ": .Fields("DESCRIPCION") = "<Seleccione un tipo de Base Imponible>"
            .AddNew: lrsTipo.Fields("CODIGO") = "006": .Fields("DESCRIPCION") = "(A) DEST. A OP.GRAV Y/O EXPORTACION"
            .AddNew: lrsTipo.Fields("CODIGO") = "007": .Fields("DESCRIPCION") = "(B) DEST. A OP.GRAV Y/O EXP. Y NO GRAV."
            .AddNew: lrsTipo.Fields("CODIGO") = "008": .Fields("DESCRIPCION") = "(C) DEST. A OP. NO GRAVADAS"
            .AddNew: lrsTipo.Fields("CODIGO") = "999": .Fields("DESCRIPCION") = " VALOR DE ADQUISICION NO GRAVADO"
            .AddNew: lrsTipo.Fields("CODIGO") = "024": .Fields("DESCRIPCION") = " OTROS"
        Else
            .AddNew: lrsTipo.Fields("CODIGO") = "   ": .Fields("DESCRIPCION") = "<Seleccione un tipo de Base Imponible de Ventas>"
            .AddNew: lrsTipo.Fields("CODIGO") = "002": .Fields("DESCRIPCION") = "GRAVABLE VENTAS"
            .AddNew: lrsTipo.Fields("CODIGO") = "021": .Fields("DESCRIPCION") = "EXPORTACIONES"
            .AddNew: lrsTipo.Fields("CODIGO") = "998": .Fields("DESCRIPCION") = "EXONERADA"
            .AddNew: lrsTipo.Fields("CODIGO") = "999": .Fields("DESCRIPCION") = "INAFECTO"
        End If
        .Update
    End With
    
    Set tdbcBaseImponible.RowSource = lrsTipo
    tdbcBaseImponible.ListField = "DESCRIPCION"
    tdbcBaseImponible.BoundColumn = "CODIGO"
End Sub
