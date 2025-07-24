VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmRepLibroMayor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libro Mayor General"
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6705
   Icon            =   "frmRepLibroMayor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3750
   ScaleWidth      =   6705
   Begin VB.CheckBox chkCuentas 
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
      Height          =   255
      Left            =   270
      TabIndex        =   4
      Top             =   1890
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Height          =   1185
      Left            =   90
      TabIndex        =   11
      Top             =   1935
      Width           =   6540
      Begin TDBText6Ctl.TDBText tdbtCuentaDesde 
         Height          =   315
         Left            =   915
         TabIndex        =   5
         Tag             =   "_"
         Top             =   330
         Width           =   1290
         _Version        =   65536
         _ExtentX        =   2275
         _ExtentY        =   556
         Caption         =   "frmRepLibroMayor.frx":1982
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmRepLibroMayor.frx":19EE
         Key             =   "frmRepLibroMayor.frx":1A0C
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
         TabIndex        =   12
         Tag             =   "_"
         Top             =   315
         Width           =   4095
         _Version        =   65536
         _ExtentX        =   7223
         _ExtentY        =   556
         Caption         =   "frmRepLibroMayor.frx":1A5E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmRepLibroMayor.frx":1ACA
         Key             =   "frmRepLibroMayor.frx":1AE8
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
         Left            =   915
         TabIndex        =   6
         Tag             =   "_"
         Top             =   690
         Width           =   1290
         _Version        =   65536
         _ExtentX        =   2275
         _ExtentY        =   556
         Caption         =   "frmRepLibroMayor.frx":1B3A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmRepLibroMayor.frx":1BA6
         Key             =   "frmRepLibroMayor.frx":1BC4
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
         TabIndex        =   13
         Tag             =   "_"
         Top             =   675
         Width           =   4095
         _Version        =   65536
         _ExtentX        =   7223
         _ExtentY        =   556
         Caption         =   "frmRepLibroMayor.frx":1C16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmRepLibroMayor.frx":1C82
         Key             =   "frmRepLibroMayor.frx":1CA0
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
         Left            =   180
         TabIndex        =   15
         Top             =   720
         Width           =   555
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
         Left            =   180
         TabIndex        =   14
         Top             =   360
         Width           =   555
      End
   End
   Begin VB.CheckBox chkAcumulado 
      Alignment       =   1  'Right Justify
      Caption         =   "MOSTRAR SALDOS ANTERIORES"
      Height          =   285
      Left            =   3420
      TabIndex        =   0
      Top             =   135
      Value           =   1  'Checked
      Width           =   3165
   End
   Begin TrueOleDBList70.TDBCombo tdbcMes 
      Height          =   300
      Left            =   4680
      TabIndex        =   1
      Top             =   540
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
      _PropDict       =   $"frmRepLibroMayor.frx":1CE4
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
      Left            =   4680
      TabIndex        =   2
      Tag             =   "_"
      Top             =   945
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
      _PropDict       =   $"frmRepLibroMayor.frx":1D6B
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
   Begin TDBText6Ctl.TDBText tdbnDigitos 
      Height          =   315
      Left            =   4680
      TabIndex        =   3
      Tag             =   "_"
      Top             =   1350
      Width           =   570
      _Version        =   65536
      _ExtentX        =   1005
      _ExtentY        =   556
      Caption         =   "frmRepLibroMayor.frx":1DF2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmRepLibroMayor.frx":1E5E
      Key             =   "frmRepLibroMayor.frx":1E7C
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
      Format          =   "aA"
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   12
      LengthAsByte    =   0
      Text            =   "12"
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
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1485
      Left            =   180
      Picture         =   "frmRepLibroMayor.frx":1ECE
      Stretch         =   -1  'True
      Top             =   90
      Width           =   1935
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "DIGITOS"
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
      Left            =   3375
      TabIndex        =   16
      Top             =   1395
      Width           =   720
   End
   Begin MSForms.CommandButton cmdSalir 
      Height          =   435
      Left            =   3510
      TabIndex        =   8
      Top             =   3240
      Width           =   1665
      Caption         =   " Salir"
      PicturePosition =   327683
      Size            =   "2937;767"
      Picture         =   "frmRepLibroMayor.frx":61CC
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdImprimir 
      Height          =   435
      Left            =   1575
      TabIndex        =   7
      Top             =   3240
      Width           =   1665
      Caption         =   " Imprimir"
      PicturePosition =   327683
      Size            =   "2937;767"
      Picture         =   "frmRepLibroMayor.frx":6766
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
      Left            =   3420
      TabIndex        =   10
      Top             =   585
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
      Left            =   3420
      TabIndex        =   9
      Top             =   975
      Width           =   765
   End
End
Attribute VB_Name = "frmRepLibroMayor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Control As String
Public ReporteSunat As String
Public TituloSunat As String

Dim gsGrupo As String
Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

Private Sub chkCuentas_Click()
    If chkCuentas.Value = 1 Then
        ActivarControl tdbtCuentaDesde, True
        ActivarControl tdbtCuentaHasta, True
        ActivarControl tdbtDescripcionDesde, True
        ActivarControl tdbtDescripcionHasta, True
    
        pSetFocus tdbtCuentaDesde
    Else
    
        ActivarControl tdbtCuentaDesde, False
        ActivarControl tdbtCuentaHasta, False
        ActivarControl tdbtDescripcionDesde, False
        ActivarControl tdbtDescripcionHasta, False
        
        tdbtCuentaDesde.Text = ""
        tdbtCuentaHasta.Text = ""
        tdbtDescripcionDesde.Text = ""
        tdbtDescripcionHasta.Text = ""
        
    End If
End Sub

Private Sub chkCuentas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pSetFocus cmdImprimir
End If
End Sub

Private Sub chkDigitos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pSetFocus tdbnDigitos
End If
End Sub

Private Function Valida() As Boolean
    Valida = False
    
    If CE(tdbtCuentaDesde) = "" And chkCuentas.Value = vbChecked Then
        Mensajes "Ingrese la cuenta de inicio", vbOKOnly + vbInformation
        Valida = False
        pSetFocus tdbtCuentaDesde
        Exit Function
    End If

    If CE(tdbtCuentaHasta) = "" And chkCuentas.Value = vbChecked Then
        Mensajes "Ingrese la cuenta final", vbOKOnly + vbInformation
        pSetFocus tdbtCuentaHasta
        Valida = False
        Exit Function
    End If

    If CE(tdbtCuentaHasta) < CE(tdbtCuentaDesde) And chkCuentas.Value = vbChecked Then
        Mensajes "La cuenta final debe ser mayor a la cuenta de inicio", vbOKOnly + vbInformation
        Valida = False
        pSetFocus tdbtCuentaHasta
        Exit Function
    End If


    Valida = True
End Function
Private Sub cmdImprimir_Click()

    If Valida = False Then Exit Sub

    Dim matriz_fecha(11) As Variant

    
    Dim Tipo As String
    
    Screen.MousePointer = vbHourglass
    '--'If CierreMes(tdbcMes.BoundText) = False Then Call ActualizaSaldosSp(tdbcMes.BoundText)
    
    Screen.MousePointer = vbHourglass
    
    If chkCuentas.Value = 1 Then
        Tipo = "CUENTAS"
    Else
        Tipo = "TODOS"
    End If
    
    matriz_fecha(0) = "@Tipo;" & Tipo & ";true"
    matriz_fecha(1) = "@Emp_cCodigo;" & gsEmpresa & ";true"
    matriz_fecha(2) = "@Pan_cAnio;" & gsAnio & ";true"
    matriz_fecha(4) = "@moneda;" & tdbcMoneda.BoundText & ";true"
    matriz_fecha(5) = "@CtaDesde;" & tdbtCuentaDesde & ";true"
    matriz_fecha(6) = "@CtaHasta;" & tdbtCuentaHasta & ";true"
    
    If Me.chkAcumulado.Value = vbChecked Then
        matriz_fecha(3) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";true"
    Else
        matriz_fecha(3) = "@Per_cPeriodo;;true"
    End If
    
    matriz_fecha(7) = "@Per_cPeriodoFIN;" & tdbcMes.BoundText & ";true"
    
    matriz_fecha(8) = "@MostrarSaldo;" & chkAcumulado.Value & ";true"
    matriz_fecha(9) = "@Digitos;" & tdbnDigitos.Text & ";true"
    
    matriz_fecha(10) = "@EMPRESA;" & gsEmpresaNom & ";true"
    matriz_fecha(11) = "@RUC;" & "RUC : " & gsRUC & ";true"
    
    Dim formulas(0) As Variant
    cmdImprimir.Enabled = False
    
    If ReporteSunat = "" Then
        AbreReporteParam gsDSN, Me, rutaReportes & "RptLibroMayorGeneral.rpt", crptToWindow, "Libro Mayor General", "", matriz_fecha(), formulas()
    
    ElseIf ReporteSunat = "F0601" Then
        AbreReporteParam gsDSN, Me, rutaReportes & "RptFormato_0601.rpt", crptToWindow, "Anexo al Libro de Inventarios y Balances", "", matriz_fecha(), formulas()
    
    End If
    
    Screen.MousePointer = vbDefault
    cmdImprimir.Enabled = True
    ' ***
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = Titulo(Me.Caption, TituloSunat)
    
    Call Centrar_form(Me)

    Call LlenaCombos

    'tdbcAlMes.BoundText = gsPeriodo
    tdbcMes.BoundText = gsPeriodo
    tdbcMoneda.BoundText = gsMonedaNac
    
    tdbcMes.ReBind
    tdbcMoneda.ReBind
    
    ActivarControl tdbtCuentaDesde, False
    ActivarControl tdbtCuentaHasta, False
    ActivarControl tdbtDescripcionDesde, False
    ActivarControl tdbtDescripcionHasta, False
    
'    SSTab1.Tab = 0
    Me.tdbnDigitos = 2
    Me.Tag = Val(tdbnDigitos.Text)
End Sub

Private Sub LlenaCombos()
    Dim sqlcombos As String
    Call LlenaComboMesApeAddItem(tdbcMes)
    'Call LlenaComboMesApeAddItem(tdbcAlMes)
    
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
'borrar
    Call SeteaFondoForm(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmRepLibroMayor = Nothing
End Sub


Private Sub tdbcMes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pSetFocus tdbcMoneda
End If
End Sub

Private Sub tdbcMoneda_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = 13 Then pSendKeys "{tab}"
End Sub

Private Sub tdbcMoneda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pSetFocus chkCuentas
End If
End Sub

Private Sub tdbnDigitos_Change()
    'If Val(tdbnDigitos.Text) < 2 Then tdbnDigitos.Text = 2
    
    tdbtCuentaDesde.MaxLength = Val(tdbnDigitos.Text)
    tdbtCuentaHasta.MaxLength = Val(tdbnDigitos.Text)

    tdbtCuentaDesde.Text = ""
    tdbtCuentaHasta.Text = ""

    'Me.Tag = Val(tdbnDigitos.Text)
    
End Sub

Private Sub tdbnDigitos_LostFocus()
        If ValidaDigitos(tdbnDigitos.Text) = False Then
            tdbnDigitos.Text = "2"
            pSetFocus tdbnDigitos
        End If
        
        Me.Tag = Val(tdbnDigitos.Text)
End Sub

Private Sub tdbtCuentaDesde_Change()
    If CE(tdbtCuentaDesde) = "" Then tdbtDescripcionDesde = ""
End Sub

Private Sub tdbtCuentaDesde_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then
        Call LlamaBuscar(frmBuscador, Me.tdbtCuentaDesde.Name, Control, "CuentasND", Me, gsPeriodo, Me.tdbtCuentaDesde.Text)
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
    If KeyCode = 112 Then
        Call LlamaBuscar(frmBuscador, Me.tdbtCuentaHasta.Name, Control, "CuentasND", Me, gsPeriodo, Me.tdbtCuentaHasta.Text)
    End If
End Sub

Private Sub tdbtCuentaHasta_LostFocus()
    ' *** Verificando q cuenta exista y q sea de no titulo
    If tdbtCuentaHasta <> "" And Me.Enabled = True Then
        tdbtDescripcionHasta = ExisteCtaNoTitulo(tdbtCuentaHasta, "")
        If tdbtDescripcionHasta = "" Then pSetFocus tdbtCuentaHasta
    End If
End Sub

Public Sub RecibirDatos(ByVal lControl As String, ByVal param0 As String, ByVal param1 As String, ByVal param2 As String)
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

'Private Function ExisteCodigo(valor As String) As Boolean
'    ' *** Verificar q codigo exista
'    Dim rsArreglo As New ADODB.Recordset
'    Dim clDatos As clsMantoTablas
'    Dim arrDatos() As Variant
'    ' *** Cargando Datos de la Cuenta
'    Dim sqlSp As String
'    ExisteCodigo = False
'    Set clDatos = New clsMantoTablas
'    sqlSp = "spCn_ConsultaCuentas 'SEL_REG', '" & gsEmpresa & "', '" & gsAnio & "', '" & valor & "'"
'    arrDatos = Array(sqlSp)
'    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
'    If rsArreglo.State <> 0 Then
'        ExisteCodigo = True
'    End If
'    Call CerrarRecordSet(rsArreglo)
'    ' ***
'End Function

Private Sub tdbtDescripcionHasta_Change()

End Sub

Private Sub tdbtDescripcionHasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        pSetFocus cmdImprimir
    End If
End Sub
