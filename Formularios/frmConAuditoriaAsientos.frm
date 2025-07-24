VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmConAuditoriaAsientos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Auditoria de Asientos Contables"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11655
   Icon            =   "frmConAuditoriaAsientos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6375
   ScaleWidth      =   11655
   Begin VB.Frame fraCabecera 
      Height          =   1860
      Left            =   120
      TabIndex        =   10
      Top             =   60
      Width           =   11415
      Begin VB.OptionButton optTipo 
         Caption         =   "Solo Asientos Eliminados"
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
         Left            =   8220
         TabIndex        =   8
         Top             =   1530
         Width           =   2745
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "Solo Asientos Activos"
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
         Left            =   8220
         TabIndex        =   7
         Top             =   1185
         Width           =   2760
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "Todos los Asientos"
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
         Left            =   8220
         TabIndex        =   6
         Top             =   870
         Value           =   -1  'True
         Width           =   2610
      End
      Begin TrueOleDBList70.TDBCombo tdbcMes 
         Height          =   300
         Left            =   9225
         TabIndex        =   0
         Top             =   270
         Width           =   2040
         _ExtentX        =   3598
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
         _PropDict       =   $"frmConAuditoriaAsientos.frx":0ECA
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
      Begin TDBText6Ctl.TDBText tdbtLibroDesc 
         Height          =   300
         Left            =   4005
         TabIndex        =   12
         Top             =   270
         Width           =   3870
         _Version        =   65536
         _ExtentX        =   6826
         _ExtentY        =   529
         Caption         =   "frmConAuditoriaAsientos.frx":0F51
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmConAuditoriaAsientos.frx":0FB5
         Key             =   "frmConAuditoriaAsientos.frx":0FD3
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
         MaxLength       =   100
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
      Begin TDBText6Ctl.TDBText tdbtLibroCod 
         Height          =   300
         Left            =   1575
         TabIndex        =   1
         Top             =   270
         Width           =   510
         _Version        =   65536
         _ExtentX        =   900
         _ExtentY        =   529
         Caption         =   "frmConAuditoriaAsientos.frx":1017
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmConAuditoriaAsientos.frx":107B
         Key             =   "frmConAuditoriaAsientos.frx":1099
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
      Begin TDBText6Ctl.TDBText tdbtNumVoucherBus 
         Height          =   300
         Left            =   1575
         TabIndex        =   2
         Top             =   630
         Width           =   1410
         _Version        =   65536
         _ExtentX        =   2487
         _ExtentY        =   529
         Caption         =   "frmConAuditoriaAsientos.frx":10DD
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmConAuditoriaAsientos.frx":1141
         Key             =   "frmConAuditoriaAsientos.frx":115F
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
         MaxLength       =   10
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
      Begin TDBText6Ctl.TDBText tdbtMonedaBus 
         Height          =   300
         Left            =   4005
         TabIndex        =   5
         Top             =   630
         Width           =   3870
         _Version        =   65536
         _ExtentX        =   6826
         _ExtentY        =   529
         Caption         =   "frmConAuditoriaAsientos.frx":11A3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmConAuditoriaAsientos.frx":1207
         Key             =   "frmConAuditoriaAsientos.frx":1225
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
         MaxLength       =   50
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
      Begin TDBText6Ctl.TDBText tdbtUsuarioCrea 
         Height          =   300
         Left            =   1575
         TabIndex        =   3
         Top             =   990
         Width           =   1410
         _Version        =   65536
         _ExtentX        =   2487
         _ExtentY        =   529
         Caption         =   "frmConAuditoriaAsientos.frx":1269
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmConAuditoriaAsientos.frx":12CD
         Key             =   "frmConAuditoriaAsientos.frx":12EB
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
         MaxLength       =   3
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
      Begin TDBText6Ctl.TDBText tdbtUsuarioModi 
         Height          =   300
         Left            =   1575
         TabIndex        =   4
         Top             =   1350
         Width           =   1410
         _Version        =   65536
         _ExtentX        =   2487
         _ExtentY        =   529
         Caption         =   "frmConAuditoriaAsientos.frx":132F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmConAuditoriaAsientos.frx":1393
         Key             =   "frmConAuditoriaAsientos.frx":13B1
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
         MaxLength       =   3
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Descripción Libro"
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
         Left            =   2295
         TabIndex        =   20
         Top             =   315
         Width           =   1455
      End
      Begin MSForms.CommandButton cmdPrint 
         Height          =   435
         Left            =   3060
         TabIndex        =   19
         ToolTipText     =   " Borrar los asientos eliminados del AÑO ACTUAL DEL SISTEMA (PACK)"
         Top             =   1260
         Width           =   1485
         Caption         =   " Imprimir"
         PicturePosition =   327683
         Size            =   "2619;767"
         Picture         =   "frmConAuditoriaAsientos.frx":13F5
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdPack 
         Height          =   435
         Left            =   4635
         TabIndex        =   18
         ToolTipText     =   " Borrar los asientos eliminados del AÑO ACTUAL DEL SISTEMA (PACK)"
         Top             =   1260
         Visible         =   0   'False
         Width           =   3285
         Caption         =   " Borrar los asientos eliminados"
         PicturePosition =   327683
         Size            =   "5794;767"
         Picture         =   "frmConAuditoriaAsientos.frx":198F
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Usuario Modif"
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
         Left            =   150
         TabIndex        =   17
         Top             =   1395
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Usuario Crea"
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
         Left            =   150
         TabIndex        =   16
         Top             =   1020
         Width           =   1110
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Cod. Libro"
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
         Left            =   165
         TabIndex        =   15
         Top             =   330
         Width           =   855
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Por Voucher"
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
         Left            =   165
         TabIndex        =   14
         Top             =   675
         Width           =   1005
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Por Glosa"
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
         Left            =   3075
         TabIndex        =   13
         Top             =   675
         Width           =   825
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Periodo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8280
         TabIndex        =   11
         Top             =   330
         Width           =   645
      End
   End
   Begin TrueOleDBGrid70.TDBGrid tdbgAsientos 
      Height          =   4335
      Left            =   135
      TabIndex        =   9
      Top             =   1980
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   7646
      _LayoutType     =   4
      _RowHeight      =   18
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Interno"
      Columns(0).DataField=   "Ase_cNummov"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Año"
      Columns(1).DataField=   "PAN_CANIO"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Mes"
      Columns(2).DataField=   "Per_cPeriodo"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "CodMoneda"
      Columns(3).DataField=   "Ase_cTipoMoneda"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Cod."
      Columns(4).DataField=   "Lib_cTipoLibro"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Libro"
      Columns(5).DataField=   "Lib_cDescripcion"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Voucher"
      Columns(6).DataField=   "Ase_nVoucher"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Fecha"
      Columns(7).DataField=   "Ase_dFecha"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Glosa"
      Columns(8).DataField=   "Ase_cGlosa"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Fecha Crea"
      Columns(9).DataField=   "Ase_dFechaCrea"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Usuario Crea"
      Columns(10).DataField=   "Ase_cUserCrea"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "Fecha Ult Mod."
      Columns(11).DataField=   "Ase_dFechaModifica"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "Usuario Ult Mod."
      Columns(12).DataField=   "Ase_cUserModifica"
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "Maquina Ult Mod."
      Columns(13).DataField=   "Ase_cEquipoUser"
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   16
      Columns(14)._MaxComboItems=   5
      Columns(14).ValueItems(0)._DefaultItem=   0
      Columns(14).ValueItems(0).Value=   "1"
      Columns(14).ValueItems(0).Value.vt=   8
      Columns(14).ValueItems(0).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(14).ValueItems(0).DisplayValue(0)=   "bHQAAOYEAABCTeYEAAAAAAAANgAAACgAAAAUAAAAFAAAAAEAGAAAAAAAsAQAAAAAAAAAAAAAAAAA"
      Columns(14).ValueItems(0).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
      Columns(14).ValueItems(0).DisplayValue(2)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(14).ValueItems(0).DisplayValue(3)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(14).ValueItems(0).DisplayValue(4)=   "//////////////////////////////////////////////////////////8AAP8AAP//////////"
      Columns(14).ValueItems(0).DisplayValue(5)=   "//////////8AAP////////////////////////////////////////////////8AAP8AAP8AAP//"
      Columns(14).ValueItems(0).DisplayValue(6)=   "//////////8AAP////////////////////////////////////////////////////////8AAP8A"
      Columns(14).ValueItems(0).DisplayValue(7)=   "AP////////8AAP8AAP////////////////////8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
      Columns(14).ValueItems(0).DisplayValue(8)=   "AAAAAAAAAP8AAP8AAP////////////////////////8AAAD/////////////////////////////"
      Columns(14).ValueItems(0).DisplayValue(9)=   "//////8AAP8AAP8AAP////////////////////////////8AAAD///8AAAAAAAD///8AAAAAAAD/"
      Columns(14).ValueItems(0).DisplayValue(10)=   "//8AAADAwMAAAP8AAP8AAP8AAP////////////////////////8AAAD/////////////////////"
      Columns(14).ValueItems(0).DisplayValue(11)=   "//////////8AAP8AAP/AwMDAwMAAAP8AAP////////////////////8AAAD/AAD/AAD/AAD/AAD/"
      Columns(14).ValueItems(0).DisplayValue(12)=   "AAD/AACAgIAAAP//AAD/AAD/AAD/AAD///////8AAP//////////////////////AAD/////////"
      Columns(14).ValueItems(0).DisplayValue(13)=   "//////////8AAP+AgID/////////////AAD/////////////////////////////////AAD///8A"
      Columns(14).ValueItems(0).DisplayValue(14)=   "AAAAAAD///8AAAAAAAD///8AAAAAAAD/////AAD/////////////////////////////////AAD/"
      Columns(14).ValueItems(0).DisplayValue(15)=   "////////////////////////////////////////AAD/////////////////////////////////"
      Columns(14).ValueItems(0).DisplayValue(16)=   "AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/////////////////////////////"
      Columns(14).ValueItems(0).DisplayValue(17)=   "////AADAwMD/AAD/AADAwMD/AAD/AADAwMD/AAD/AADAwMD/AAD/////////////////////////"
      Columns(14).ValueItems(0).DisplayValue(18)=   "////////AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/////////////////////"
      Columns(14).ValueItems(0).DisplayValue(19)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(14).ValueItems(0).DisplayValue(20)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(14).ValueItems(0).DisplayValue(21)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(14).ValueItems(0).DisplayValue(22)=   "//////////8="
      Columns(14).ValueItems(0).DisplayValue.vt=   9
      Columns(14).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(14).ValueItems(1)._DefaultItem=   0
      Columns(14).ValueItems(1).Value=   "0"
      Columns(14).ValueItems(1).Value.vt=   8
      Columns(14).ValueItems(1).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(14).ValueItems(1).DisplayValue(0)=   "bHQAAAYIAABCTQYIAAAAAAAANgAAACgAAAAaAAAAGQAAAAEAGAAAAAAA0AcAAAAAAAAAAAAAAAAA"
      Columns(14).ValueItems(1).DisplayValue(1)=   "AAAAAADx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx"
      Columns(14).ValueItems(1).DisplayValue(2)=   "7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+sAAPHv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv"
      Columns(14).ValueItems(1).DisplayValue(3)=   "6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6wAA8e/r8e/r"
      Columns(14).ValueItems(1).DisplayValue(4)=   "8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r"
      Columns(14).ValueItems(1).DisplayValue(5)=   "8e/r8e/r8e/r8e/r8e/rAADx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx"
      Columns(14).ValueItems(1).DisplayValue(6)=   "7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+sAAPHv6/Hv6/Hv6/Hv6/Hv6/Hv"
      Columns(14).ValueItems(1).DisplayValue(7)=   "6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv"
      Columns(14).ValueItems(1).DisplayValue(8)=   "6/Hv6wAA8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r"
      Columns(14).ValueItems(1).DisplayValue(9)=   "8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/rAADx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx"
      Columns(14).ValueItems(1).DisplayValue(10)=   "7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+sAAPHv6/Hv"
      Columns(14).ValueItems(1).DisplayValue(11)=   "6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv"
      Columns(14).ValueItems(1).DisplayValue(12)=   "6/Hv6/Hv6/Hv6/Hv6/Hv6wAA8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r"
      Columns(14).ValueItems(1).DisplayValue(13)=   "8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/rAADx7+vx7+vx7+vx7+vx7+vx"
      Columns(14).ValueItems(1).DisplayValue(14)=   "7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx"
      Columns(14).ValueItems(1).DisplayValue(15)=   "7+vx7+sAAPHv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv"
      Columns(14).ValueItems(1).DisplayValue(16)=   "6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6wAA8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r"
      Columns(14).ValueItems(1).DisplayValue(17)=   "8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/rAADx7+vx"
      Columns(14).ValueItems(1).DisplayValue(18)=   "7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx"
      Columns(14).ValueItems(1).DisplayValue(19)=   "7+vx7+vx7+vx7+vx7+vx7+sAAPHv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv"
      Columns(14).ValueItems(1).DisplayValue(20)=   "6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6wAA8e/r8e/r8e/r8e/r8e/r"
      Columns(14).ValueItems(1).DisplayValue(21)=   "8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r"
      Columns(14).ValueItems(1).DisplayValue(22)=   "8e/r8e/rAADx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx"
      Columns(14).ValueItems(1).DisplayValue(23)=   "7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+sAAPHv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv"
      Columns(14).ValueItems(1).DisplayValue(24)=   "6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6wAA8e/r"
      Columns(14).ValueItems(1).DisplayValue(25)=   "8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r"
      Columns(14).ValueItems(1).DisplayValue(26)=   "8e/r8e/r8e/r8e/r8e/r8e/rAADx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx"
      Columns(14).ValueItems(1).DisplayValue(27)=   "7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+sAAPHv6/Hv6/Hv6/Hv6/Hv"
      Columns(14).ValueItems(1).DisplayValue(28)=   "6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv"
      Columns(14).ValueItems(1).DisplayValue(29)=   "6/Hv6/Hv6wAA8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r"
      Columns(14).ValueItems(1).DisplayValue(30)=   "8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/rAADx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx"
      Columns(14).ValueItems(1).DisplayValue(31)=   "7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+sAAPHv"
      Columns(14).ValueItems(1).DisplayValue(32)=   "6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv"
      Columns(14).ValueItems(1).DisplayValue(33)=   "6/Hv6/Hv6/Hv6/Hv6/Hv6/Hv6wAA8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r"
      Columns(14).ValueItems(1).DisplayValue(34)=   "8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/r8e/rAADx7+vx7+vx7+vx7+vx"
      Columns(14).ValueItems(1).DisplayValue(35)=   "7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx7+vx"
      Columns(14).ValueItems(1).DisplayValue(36)=   "7+vx7+vx7+sAAA=="
      Columns(14).ValueItems(1).DisplayValue.vt=   9
      Columns(14).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(14).ValueItems.Count=   2
      Columns(14).Caption=   "Estado"
      Columns(14).DataField=   "Ase_cDeleted"
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   15
      Splits(0)._UserFlags=   0
      Splits(0).Locked=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=15"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=532"
      Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=532"
      Splits(0)._ColumnProps(13)=   "Column(1).Visible=0"
      Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(15)=   "Column(2).Width=688"
      Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=609"
      Splits(0)._ColumnProps(18)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(19)=   "Column(2)._ColStyle=532"
      Splits(0)._ColumnProps(20)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(21)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(22)=   "Column(3).Width=1535"
      Splits(0)._ColumnProps(23)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(3)._WidthInPix=1455"
      Splits(0)._ColumnProps(25)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(26)=   "Column(3)._ColStyle=532"
      Splits(0)._ColumnProps(27)=   "Column(3).Visible=0"
      Splits(0)._ColumnProps(28)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(29)=   "Column(4).Width=688"
      Splits(0)._ColumnProps(30)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(4)._WidthInPix=609"
      Splits(0)._ColumnProps(32)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(33)=   "Column(4)._ColStyle=532"
      Splits(0)._ColumnProps(34)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(35)=   "Column(4).Merge=1"
      Splits(0)._ColumnProps(36)=   "Column(5).Width=2831"
      Splits(0)._ColumnProps(37)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(5)._WidthInPix=2752"
      Splits(0)._ColumnProps(39)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(40)=   "Column(5)._ColStyle=532"
      Splits(0)._ColumnProps(41)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(42)=   "Column(5).Merge=1"
      Splits(0)._ColumnProps(43)=   "Column(6).Width=1931"
      Splits(0)._ColumnProps(44)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(45)=   "Column(6)._WidthInPix=1852"
      Splits(0)._ColumnProps(46)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(47)=   "Column(6)._ColStyle=529"
      Splits(0)._ColumnProps(48)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(49)=   "Column(7).Width=1746"
      Splits(0)._ColumnProps(50)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(51)=   "Column(7)._WidthInPix=1667"
      Splits(0)._ColumnProps(52)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(53)=   "Column(7)._ColStyle=532"
      Splits(0)._ColumnProps(54)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(55)=   "Column(8).Width=6165"
      Splits(0)._ColumnProps(56)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(57)=   "Column(8)._WidthInPix=6085"
      Splits(0)._ColumnProps(58)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(59)=   "Column(8)._ColStyle=532"
      Splits(0)._ColumnProps(60)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(61)=   "Column(9).Width=3704"
      Splits(0)._ColumnProps(62)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(63)=   "Column(9)._WidthInPix=3625"
      Splits(0)._ColumnProps(64)=   "Column(9)._EditAlways=0"
      Splits(0)._ColumnProps(65)=   "Column(9)._ColStyle=532"
      Splits(0)._ColumnProps(66)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(67)=   "Column(10).Width=2699"
      Splits(0)._ColumnProps(68)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(69)=   "Column(10)._WidthInPix=2619"
      Splits(0)._ColumnProps(70)=   "Column(10)._EditAlways=0"
      Splits(0)._ColumnProps(71)=   "Column(10)._ColStyle=532"
      Splits(0)._ColumnProps(72)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(73)=   "Column(11).Width=3863"
      Splits(0)._ColumnProps(74)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(75)=   "Column(11)._WidthInPix=3784"
      Splits(0)._ColumnProps(76)=   "Column(11)._EditAlways=0"
      Splits(0)._ColumnProps(77)=   "Column(11)._ColStyle=532"
      Splits(0)._ColumnProps(78)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(79)=   "Column(12).Width=2196"
      Splits(0)._ColumnProps(80)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(81)=   "Column(12)._WidthInPix=2117"
      Splits(0)._ColumnProps(82)=   "Column(12)._EditAlways=0"
      Splits(0)._ColumnProps(83)=   "Column(12)._ColStyle=532"
      Splits(0)._ColumnProps(84)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(85)=   "Column(13).Width=2778"
      Splits(0)._ColumnProps(86)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(87)=   "Column(13)._WidthInPix=2699"
      Splits(0)._ColumnProps(88)=   "Column(13)._EditAlways=0"
      Splits(0)._ColumnProps(89)=   "Column(13)._ColStyle=532"
      Splits(0)._ColumnProps(90)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(91)=   "Column(14).Width=820"
      Splits(0)._ColumnProps(92)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(93)=   "Column(14)._WidthInPix=741"
      Splits(0)._ColumnProps(94)=   "Column(14)._EditAlways=0"
      Splits(0)._ColumnProps(95)=   "Column(14)._ColStyle=529"
      Splits(0)._ColumnProps(96)=   "Column(14).Order=15"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      EmptyRows       =   -1  'True
      CellTipsWidth   =   0
      DeadAreaBackColor=   16777215
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=192,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.fgcolor=&H80000008&,.bold=0,.fontsize=825"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Arial"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bgcolor=&HCA570B&"
      _StyleDefs(11)  =   ":id=2,.fgcolor=&H80000014&,.bold=0,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(12)  =   ":id=2,.strikethrough=0,.charset=0"
      _StyleDefs(13)  =   ":id=2,.fontname=Arial"
      _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(15)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(16)  =   ":id=3,.fontname=Arial"
      _StyleDefs(17)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(18)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1"
      _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1,.valignment=2,.bgcolor=&HFFFFFF&"
      _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(28)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(29)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(30)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(31)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(32)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(33)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(34)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(35)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(36)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=74,.parent=13"
      _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=71,.parent=14"
      _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=72,.parent=15"
      _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=73,.parent=17"
      _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=70,.parent=13"
      _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=67,.parent=14"
      _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=68,.parent=15"
      _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=69,.parent=17"
      _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=54,.parent=13"
      _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
      _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
      _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
      _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=62,.parent=13"
      _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=14"
      _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=15"
      _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=17"
      _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=46,.parent=13"
      _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
      _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
      _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
      _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
      _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
      _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
      _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
      _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=28,.parent=13,.alignment=2"
      _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=25,.parent=14"
      _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=26,.parent=15"
      _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=27,.parent=17"
      _StyleDefs(65)  =   "Splits(0).Columns(7).Style:id=50,.parent=13"
      _StyleDefs(66)  =   "Splits(0).Columns(7).HeadingStyle:id=47,.parent=14"
      _StyleDefs(67)  =   "Splits(0).Columns(7).FooterStyle:id=48,.parent=15"
      _StyleDefs(68)  =   "Splits(0).Columns(7).EditorStyle:id=49,.parent=17"
      _StyleDefs(69)  =   "Splits(0).Columns(8).Style:id=78,.parent=13"
      _StyleDefs(70)  =   "Splits(0).Columns(8).HeadingStyle:id=75,.parent=14"
      _StyleDefs(71)  =   "Splits(0).Columns(8).FooterStyle:id=76,.parent=15"
      _StyleDefs(72)  =   "Splits(0).Columns(8).EditorStyle:id=77,.parent=17"
      _StyleDefs(73)  =   "Splits(0).Columns(9).Style:id=32,.parent=13"
      _StyleDefs(74)  =   "Splits(0).Columns(9).HeadingStyle:id=29,.parent=14"
      _StyleDefs(75)  =   "Splits(0).Columns(9).FooterStyle:id=30,.parent=15"
      _StyleDefs(76)  =   "Splits(0).Columns(9).EditorStyle:id=31,.parent=17"
      _StyleDefs(77)  =   "Splits(0).Columns(10).Style:id=66,.parent=13"
      _StyleDefs(78)  =   "Splits(0).Columns(10).HeadingStyle:id=63,.parent=14"
      _StyleDefs(79)  =   "Splits(0).Columns(10).FooterStyle:id=64,.parent=15"
      _StyleDefs(80)  =   "Splits(0).Columns(10).EditorStyle:id=65,.parent=17"
      _StyleDefs(81)  =   "Splits(0).Columns(11).Style:id=82,.parent=13"
      _StyleDefs(82)  =   "Splits(0).Columns(11).HeadingStyle:id=79,.parent=14"
      _StyleDefs(83)  =   "Splits(0).Columns(11).FooterStyle:id=80,.parent=15"
      _StyleDefs(84)  =   "Splits(0).Columns(11).EditorStyle:id=81,.parent=17"
      _StyleDefs(85)  =   "Splits(0).Columns(12).Style:id=86,.parent=13"
      _StyleDefs(86)  =   "Splits(0).Columns(12).HeadingStyle:id=83,.parent=14"
      _StyleDefs(87)  =   "Splits(0).Columns(12).FooterStyle:id=84,.parent=15"
      _StyleDefs(88)  =   "Splits(0).Columns(12).EditorStyle:id=85,.parent=17"
      _StyleDefs(89)  =   "Splits(0).Columns(13).Style:id=90,.parent=13"
      _StyleDefs(90)  =   "Splits(0).Columns(13).HeadingStyle:id=87,.parent=14"
      _StyleDefs(91)  =   "Splits(0).Columns(13).FooterStyle:id=88,.parent=15"
      _StyleDefs(92)  =   "Splits(0).Columns(13).EditorStyle:id=89,.parent=17"
      _StyleDefs(93)  =   "Splits(0).Columns(14).Style:id=94,.parent=13,.alignment=2"
      _StyleDefs(94)  =   "Splits(0).Columns(14).HeadingStyle:id=91,.parent=14"
      _StyleDefs(95)  =   "Splits(0).Columns(14).FooterStyle:id=92,.parent=15"
      _StyleDefs(96)  =   "Splits(0).Columns(14).EditorStyle:id=93,.parent=17"
      _StyleDefs(97)  =   "Named:id=33:Normal"
      _StyleDefs(98)  =   ":id=33,.parent=0"
      _StyleDefs(99)  =   "Named:id=34:Heading"
      _StyleDefs(100) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(101) =   ":id=34,.wraptext=-1"
      _StyleDefs(102) =   "Named:id=35:Footing"
      _StyleDefs(103) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(104) =   "Named:id=36:Selected"
      _StyleDefs(105) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(106) =   "Named:id=37:Caption"
      _StyleDefs(107) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(108) =   "Named:id=38:HighlightRow"
      _StyleDefs(109) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(110) =   "Named:id=39:EvenRow"
      _StyleDefs(111) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(112) =   "Named:id=40:OddRow"
      _StyleDefs(113) =   ":id=40,.parent=33"
      _StyleDefs(114) =   "Named:id=41:RecordSelector"
      _StyleDefs(115) =   ":id=41,.parent=34"
      _StyleDefs(116) =   "Named:id=42:FilterBar"
      _StyleDefs(117) =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "frmConAuditoriaAsientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : frmConAuditoriaAsientos
'    Project    : Contabilidad
'
'    Description: Formulario de auditoria de asientos contables
'--------------------------------------------------------------------------------
Option Explicit
Dim lrsAsientos As ADODB.Recordset

Dim gsGrupo As String
'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Grupo
' Description:       Propiedad de grupo de asientos contables
'
' Parameters :       Grupo (String)
'--------------------------------------------------------------------------------
Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       cmdPack_Click
' Description:       Evento que se ejecuta al hacer clic en eliminar asientos temporales
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub cmdPack_Click()
    If MsgBox("Desea borrar los asientos eliminados que fueron almacenados para auditoria " & Salto(1) & "del ejercicio contable " & gsAnio & " de la empresa " & gsEmpresaNom, vbQuestion + vbYesNo) = vbYes Then
        Call Pack
        
        DoEvents
        
        Call tdbcMes_ItemChange
    End If
    
    
End Sub


'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Pack
' Description:       Procedimiento de eliminacion de asientos temporales
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub Pack()
    Dim clsMante As clsMantoTablas
    Dim lArrMnt(7) As Variant
    
    Set clsMante = New clsMantoTablas
    lArrMnt(0) = "PACK_CONTA"
    lArrMnt(1) = gsEmpresa
    lArrMnt(2) = Null
    lArrMnt(3) = Null
    lArrMnt(4) = Null
    lArrMnt(5) = Null
    lArrMnt(6) = Null
    lArrMnt(7) = gsAnio
    
    If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaEmpresa", lArrMnt(), True) = False Then
        Mensajes "El proceso no ha concluido. Verificar...", vbInformation
        Exit Sub
    End If
    Mensajes "Se borraron los movimientos eliminados con exito (son movimientos que quedan almacenados para auditoria) ...", vbInformation
    
    Set clsMante = Nothing
    Exit Sub
ErrorEjecucion:
    Mensajes Str(Err.Number) & Err.Description, vbInformation
End Sub


'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       cmdPrint_Click
' Description:       Evento que se ejecuta al hacer clic en imprimir
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub cmdPrint_Click()
    Dim cIndice As Integer
    Dim cCadena As String
    If optTipo(0).Value = True Then cIndice = 0: cCadena = "Todos los asientos"
    If optTipo(1).Value = True Then cIndice = 1: cCadena = "Asientos activos"
    If optTipo(2).Value = True Then cIndice = 2: cCadena = "Asientos eliminados"
    
    cmdPrint.Enabled = False
    DoEvents
    
    Dim matriz(7) As Variant
    Dim Titulo As String
    Titulo = "Reporte de Auditoria - " & cCadena
    Titulo = UCase(Titulo)
    matriz(0) = "@Tipo;" & "AUDITORIA" & ";True"
    matriz(1) = "@Emp_cCodigo;" & gsEmpresa & ";True"
    matriz(2) = "@Pan_cAnio;" & gsAnio & ";True"
    matriz(3) = "@Per_cPeriodo;" & tdbcMes.BoundText & ";True"
    matriz(4) = "@Estado;" & cIndice & ";True"
    matriz(5) = "@EMPRESA;" & gsEmpresaNom & ";True"
    matriz(6) = "@RUC;" & "RUC : " & gsRUC & ";True"
    matriz(7) = "@Titulo00;" & Titulo & ";True"
    
    
    Dim formulas(0) As Variant
    AbreReporteParam gsDSN, Me, rutaReportes & "RptAuditoriaAsientos.rpt", crptToWindow, Titulo, "", matriz(), formulas()
    
    cmdPrint.Enabled = True
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Form_Load
' Description:       Evento que se ejecuta al iniciar el formulario
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub Form_Load()
    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))
    
    Call Centrar_form(Me)
    Call LlenaComboMesApeAddItem(tdbcMes)
    tdbcMes.Bookmark = 1
    Call CargaTabla(0)
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       CargaTabla
' Description:       Procedimiento que carga la lista de asientos a auditar
'
' Parameters :       cIndice (Integer)
'--------------------------------------------------------------------------------
Private Sub CargaTabla(cIndice As Integer)
    Dim sqlSp As String
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
    Set clDatos = New clsMantoTablas
    Set lrsAsientos = New ADODB.Recordset
    Set lrsAsientos.DataSource = Nothing
    tdbgAsientos.DataSource = lrsAsientos
    tdbgAsientos.ReBind
    
    sqlSp = "spCn_ConsultaAsientosAuditoria 'AUDITORIA', '" & gsEmpresa & "', '" & gsAnio & "' , '" & Me.tdbcMes.BoundText & "','" & cIndice & "'"
    
    arrDatos = Array(sqlSp)
    Set lrsAsientos = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If Not lrsAsientos Is Nothing Then
        lrsAsientos.Sort = "Per_cPeriodo, Lib_cDescripcion, Ase_nVoucher"
        tdbgAsientos.DataSource = lrsAsientos
        tdbgAsientos.ReBind
        Exit Sub
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Form_Resize
' Description:       Evento que se ejecuta al cambiar de tamaño el formulario
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
       On Error GoTo serror
       tdbgAsientos.Width = Me.Width - 300
       fraCabecera.Width = tdbgAsientos.Width
       tdbgAsientos.Height = Me.Height - tdbgAsientos.Top - 600
    End If
    Exit Sub
serror:
    'Mensajes Err.Description
    
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       Form_Unload
' Description:       Evento que se ejecuta al cerrar el formulario
'
' Parameters :       Cancel (Integer)
'--------------------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
    Set lrsAsientos = Nothing
    Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))

End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       optEliminados_Click
' Description:       Evento que se ejecuta al cambiar la opcion de mostrar lo vouchers eliminados
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub optEliminados_Click()
    Call FiltrarRecordSet
    cmdPack.Visible = True
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       optNoEliminados_Click
' Description:       Evento que se ejecuta al cambiar la opcion de mostrar los vouchers no eliminados
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub optNoEliminados_Click()
    Call FiltrarRecordSet
    cmdPack.Visible = False
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       optTodos_Click
' Description:       Evento que se ejecuta al cambiar la opcion de mostrar todos los vouchers
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub optTodos_Click()
    Call FiltrarRecordSet
    cmdPack.Visible = False
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       optTipo_Click
' Description:       Evento que se ejecuta al cambiar el tipo de vouchers a mostrar ocultando el boton PACK
'
' Parameters :       Index (Integer)
'--------------------------------------------------------------------------------
Private Sub optTipo_Click(Index As Integer)
    Call CargaTabla(Index)
    
    If Index = 2 Then
        cmdPack.Visible = True
    Else
        cmdPack.Visible = False
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbcMes_ItemChange
' Description:       Evento que se ejecuta al cambiar el combo del mes
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbcMes_ItemChange()
    Dim cIndice As Integer
    If optTipo(0).Value = True Then cIndice = 0
    If optTipo(1).Value = True Then cIndice = 1
    If optTipo(2).Value = True Then cIndice = 2
    
    Call optTipo_Click(cIndice)
    Call FiltrarRecordSet
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbcMes_KeyPress
' Description:       Evento que se ejecuta al presionar la tecla en el combo delmes
'
' Parameters :       KeyAscii (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbcMes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pSetFocus tdbtLibroCod
End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbgAsientos_GotFocus
' Description:       Evento que se ejecuta al recibirel enfoque la grilla
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbgAsientos_GotFocus()
tdbgAsientos.HighlightRowStyle = "HighlightRow"
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbgAsientos_LostFocus
' Description:       Evento que se ejecuta al perder el enfoquela grilla
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbgAsientos_LostFocus()
tdbgAsientos.HighlightRowStyle = ""
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtLibroCod_Change
' Description:       Evento que se ejecuta al cambiar eltexto del codigo de libro
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbtLibroCod_Change()
    Call FiltrarRecordSet
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtLibroCod_KeyPress
' Description:       Evento que se ejecuta al presionar una tecla en el codigo del libro
'
' Parameters :       KeyAscii (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbtLibroCod_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pSetFocus tdbtNumVoucherBus
End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtLibroDesc_Change
' Description:       Evento que se ejecuta al cambiar la descripcion del codigo de libro
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbtLibroDesc_Change()
    If gsKey = 219 Then
       tdbtLibroDesc = Replace(tdbtLibroDesc, "'", "")
       tdbtLibroDesc.SelStart = Len(tdbtLibroDesc)
    End If

    Call FiltrarRecordSet
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       FiltrarRecordSet
' Description:       Procedimiento de filtradode la grilla segun los campos digitados
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub FiltrarRecordSet()
    ' *** Filtrar segun los textos indicados
    Dim cadena As String
    Dim cadElim As String
    Dim filtros(6) As String
    Dim i As Integer
    If lrsAsientos Is Nothing Then Exit Sub
    cadena = ""
    If Trim(Me.tdbtLibroCod) <> "" Then filtros(0) = "Lib_cTipoLibro like '*" & tdbtLibroCod & "*'"
    If Trim(Me.tdbtLibroDesc) <> "" Then filtros(1) = "Lib_cDescripcion like '*" & tdbtLibroDesc & "*'"
    If Trim(Me.tdbtNumVoucherBus) <> "" Then filtros(2) = "Ase_nVoucher like '*" & tdbtNumVoucherBus & "*'"
    If Trim(Me.tdbtMonedaBus) <> "" Then filtros(3) = "Ase_cGlosa like '*" & tdbtMonedaBus & "*'"
    If Trim(Me.tdbtUsuarioCrea) <> "" Then filtros(4) = "Ase_cUserCrea like '*" & tdbtUsuarioCrea & "*'"
    If Trim(Me.tdbtUsuarioModi) <> "" Then filtros(5) = "Ase_cUserModifica like '*" & tdbtUsuarioModi & "*'"
    ' *** Ver el filtro de los eliminados
    'If Me.optEliminados.Value = True Then filtros(6) = "Ase_cDeleted like '1'"
    'If Me.optNoEliminados.Value = True Then filtros(6) = "Ase_cDeleted like '0'"
    ' ***
    For i = 0 To 6
        If filtros(i) <> "" Then
            If cadena = "" Then
                cadena = cadena + filtros(i)
            Else
                cadena = cadena + " and " + filtros(i)
            End If
        End If
    Next
    ' *** Filtrando segun campos
    If Trim(cadena) <> "" Then
        lrsAsientos.Filter = cadena
    Else
        lrsAsientos.Filter = 0
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtLibroDesc_KeyDown
' Description:       Evento que se ejecuta al presionar la tecla en el campo de descripcion del libro
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbtLibroDesc_KeyDown(KeyCode As Integer, Shift As Integer)
    gsKey = KeyCode
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtMonedaBus_Change
' Description:       Evento que se ejecuta al cambiar ek codigo de moneda
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbtMonedaBus_Change()
    If gsKey = 219 Then
       tdbtMonedaBus = Replace(tdbtMonedaBus, "'", "")
       tdbtMonedaBus.SelStart = Len(tdbtMonedaBus)
    End If

    Call FiltrarRecordSet
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtMonedaBus_KeyDown
' Description:       Evento que se ejecuta al presionar una tecla en el codigo de moneda
'
' Parameters :       KeyCode (Integer)
'                    Shift (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbtMonedaBus_KeyDown(KeyCode As Integer, Shift As Integer)
    gsKey = KeyCode
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtNumVoucherBus_Change
' Description:       Evento que se ejecuta al cmabiar el numero de ovucher
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbtNumVoucherBus_Change()
    Call FiltrarRecordSet
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtNumVoucherBus_KeyPress
' Description:       Evento que se ejecuta al presionar una tecla en el numero de voucher
'
' Parameters :       KeyAscii (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbtNumVoucherBus_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pSetFocus tdbtUsuarioCrea
End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtUsuarioCrea_Change
' Description:       Evento que se ejecuta al cambiar el usuario de creacion
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbtUsuarioCrea_Change()
    Call FiltrarRecordSet
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtUsuarioCrea_KeyPress
' Description:       Evento que se ejecuta al presionar una tecla en el usuario de creacion
'
' Parameters :       KeyAscii (Integer)
'--------------------------------------------------------------------------------
Private Sub tdbtUsuarioCrea_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pSetFocus tdbtUsuarioModi
End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       Contabilidad
' Procedure  :       tdbtUsuarioModi_Change
' Description:       Evento que se ejecuta al cambiar el usuario de modificacion
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub tdbtUsuarioModi_Change()
    Call FiltrarRecordSet
End Sub
