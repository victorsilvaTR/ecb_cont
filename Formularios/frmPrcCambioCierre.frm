VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmPrcCambioCierre 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Diferencia por Tipo de Cambio Mensual"
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5685
   Icon            =   "frmPrcCambioCierre.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3015
   ScaleWidth      =   5685
   Begin VB.Frame fraTodo 
      Height          =   2940
      Left            =   45
      TabIndex        =   7
      Top             =   30
      Width           =   5610
      Begin TrueOleDBList70.TDBCombo tdbcMes 
         Height          =   300
         Left            =   1230
         TabIndex        =   1
         Top             =   840
         Width           =   4050
         _ExtentX        =   7144
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
         _PropDict       =   $"frmPrcCambioCierre.frx":0ECA
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
      Begin TDBNumber6Ctl.TDBNumber tdbnTipoCambioC 
         Height          =   300
         Left            =   3240
         TabIndex        =   2
         Tag             =   "enabled"
         Top             =   1380
         Width           =   765
         _Version        =   65536
         _ExtentX        =   1349
         _ExtentY        =   529
         Calculator      =   "frmPrcCambioCierre.frx":0F51
         Caption         =   "frmPrcCambioCierre.frx":0F71
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPrcCambioCierre.frx":0FDD
         Keys            =   "frmPrcCambioCierre.frx":0FFB
         Spin            =   "frmPrcCambioCierre.frx":1053
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "####0.000"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0.000"
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99999
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
         MaxValueVT      =   1380909061
         MinValueVT      =   1162608645
      End
      Begin TrueOleDBList70.TDBCombo tdbcLibro 
         Height          =   300
         Left            =   1230
         TabIndex        =   0
         Tag             =   "enabled"
         Top             =   345
         Width           =   4020
         _ExtentX        =   7091
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
         _PropDict       =   $"frmPrcCambioCierre.frx":107B
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
      Begin TDBNumber6Ctl.TDBNumber tdbnTipoCambioV 
         Height          =   300
         Left            =   3240
         TabIndex        =   3
         Tag             =   "enabled"
         Top             =   1770
         Width           =   765
         _Version        =   65536
         _ExtentX        =   1349
         _ExtentY        =   529
         Calculator      =   "frmPrcCambioCierre.frx":1102
         Caption         =   "frmPrcCambioCierre.frx":1122
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmPrcCambioCierre.frx":118E
         Keys            =   "frmPrcCambioCierre.frx":11AC
         Spin            =   "frmPrcCambioCierre.frx":1204
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "####0.000"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0.000"
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99999
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
         MaxValueVT      =   1380909061
         MinValueVT      =   1162608645
      End
      Begin MSForms.CommandButton cmdEliminar 
         Height          =   435
         Left            =   1935
         TabIndex        =   5
         Top             =   2370
         Width           =   1665
         Caption         =   " Eliminar"
         PicturePosition =   327683
         Size            =   "2937;767"
         Picture         =   "frmPrcCambioCierre.frx":122C
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "T. CAMBIO VENTA"
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
         Left            =   1260
         TabIndex        =   11
         Top             =   1800
         Width           =   1485
      End
      Begin MSForms.CommandButton cmd_salir 
         Height          =   435
         Left            =   3690
         TabIndex        =   6
         Top             =   2370
         Width           =   1665
         Caption         =   " Salir"
         PicturePosition =   327683
         Size            =   "2937;767"
         Picture         =   "frmPrcCambioCierre.frx":17C6
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdGenerar 
         Height          =   435
         Left            =   180
         TabIndex        =   4
         Top             =   2370
         Width           =   1665
         Caption         =   " Generar"
         PicturePosition =   327683
         Size            =   "2937;767"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "LIBRO"
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
         Left            =   195
         TabIndex        =   10
         Top             =   390
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "T. CAMBIO COMPRA"
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
         Left            =   1230
         TabIndex        =   9
         Top             =   1380
         Width           =   1680
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
         Left            =   195
         TabIndex        =   8
         Top             =   900
         Width           =   375
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
      TabIndex        =   12
      Top             =   0
      Width           =   4365
   End
End
Attribute VB_Name = "frmPrcCambioCierre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Control As String
Dim lsLibroDif  As String
Dim gsGrupo As String
Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

Private Sub cmd_salir_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If CierreMes(tdbcMes.BoundText) Then
        Mensajes "El mes esta bloqueado, no se puede eliminar los asientos del libro diferencia de cambio"
        Exit Sub
    End If
    
    If MsgBox("Deseas eliminar los asientos de " & Salto(1) & "Dif. de Cambio a partir del mes de " & tdbcMes.Text, vbYesNo + vbQuestion) = vbYes Then
       Call EliminaAsientosdifCabio
    End If
End Sub

Private Sub EliminaAsientosdifCabio()
    Dim lArrMnt() As Variant
    Dim Mes As Integer
    Dim MesInicio As Integer
    Dim Cerrado As Boolean
    Dim pMes As String
    
    Screen.MousePointer = vbHourglass
    cmdEliminar.Enabled = False
    DoEvents
    
    Dim clsMante As clsMantoTablas
    Set clsMante = New clsMantoTablas
    ReDim lArrMnt(5) As Variant
    On Local Error GoTo ErrorEjecucion
    MesInicio = NE(tdbcMes.BoundText)
    lArrMnt(0) = gsEmpresa
    lArrMnt(1) = gsAnio
    lArrMnt(2) = tdbcMes.BoundText
    lArrMnt(3) = tdbcMes.BoundText '"12"
    lArrMnt(4) = tdbcLibro.BoundText
    lArrMnt(5) = gsUsuario
    
    Cerrado = False
    
    For Mes = MesInicio To 12
        pMes = Right("00" & CE(Mes), 2)
        If CierreMes(pMes) Then
            Cerrado = True
            Mensajes "El mes de " & NombreMes(pMes) & " esta cerrado, no se puede continuar"
            Exit For
        End If
    Next Mes
    
    If Cerrado = False Then
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_EliminaDifCambio", lArrMnt(), True) = False Then
            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
            Screen.MousePointer = vbDefault
        End If
    End If
    
    Set clsMante = Nothing
    
    Screen.MousePointer = vbNormal
    cmdEliminar.Enabled = True
    Mensajes "Proceso terminado ...", vbInformation
    Exit Sub
ErrorEjecucion:
    Mensajes Str(Err.Number) & " " & Err.Description, vbInformation
    Set clsMante = Nothing
    Screen.MousePointer = vbNormal
    cmdEliminar.Enabled = True
    
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
    Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))
End Sub

Private Sub tdbcMes_ItemChange()
    tdbnTipoCambioC.Value = CargaTCxMeses(1)
    tdbnTipoCambioV.Value = CargaTCxMeses(2)
End Sub

Private Function CargaTCxMeses(nTipoCambio As Integer) As Double
    Dim sql As String
    Dim arrDatos() As Variant
    Dim rsAddItem As ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Dim vMes(11) As Double
    
    sql = "select * from CNT_TIPO_CAMBIO_MENSUAL " & _
          "where emp_ccodigo='" & gsEmpresa & "' and pan_canio='" & gsAnio & "' and " & _
          "tca_cmoneda='" & gsMonedaExt & "' and tca_ctipo='" & nTipoCambio & "'"
    
    Set clDatos = New clsMantoTablas
    arrDatos = Array(sql)
    Set rsAddItem = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())

    If Not rsAddItem Is Nothing Then
        Do While Not rsAddItem.EOF
            vMes(0) = NE(rsAddItem!Tca_cEne)
            vMes(1) = NE(rsAddItem!Tca_cFeb)
            vMes(2) = NE(rsAddItem!Tca_cMar)
            vMes(3) = NE(rsAddItem!Tca_cAbr)
            vMes(4) = NE(rsAddItem!Tca_cMay)
            vMes(5) = NE(rsAddItem!Tca_cJun)
            vMes(6) = NE(rsAddItem!Tca_cJul)
            vMes(7) = NE(rsAddItem!Tca_cAgo)
            vMes(8) = NE(rsAddItem!Tca_cSet)
            vMes(9) = NE(rsAddItem!Tca_cOct)
            vMes(10) = NE(rsAddItem!Tca_cNov)
            vMes(11) = NE(rsAddItem!Tca_cDic)
            
            CargaTCxMeses = vMes(NE(Val(tdbcMes.BoundText) - 1))

            rsAddItem.MoveNext
        Loop
    Else
        'Mensajes "Ingrese los tipos de cambio mensuales del tipo " & tdbcTipoMensual, vbOKOnly + vbInformation
        'tdbnTipoCambio.Value = 0
        CargaTCxMeses = 0
    End If
    
    Call CerrarRecordSet(rsAddItem)
    Set clDatos = Nothing
End Function

Private Function Validar() As Boolean
    Validar = False
    
    If CE(tdbcLibro.Text) = "" Then
        Mensajes "Cree el libro de diferencia en cambio y configurelo en Parametros iniciales", vbOKOnly + vbInformation
        Exit Function
    End If
    
'    If Left(CE(tdbtCuentaDesde.Text), 2) > "59" Then
'        Mensajes "La cuenta de inicio debe ser una cuenta de balance", vbOKOnly + vbInformation
'        Exit Function
'    End If
'
'    If Left(CE(tdbtCuentaHasta.Text), 2) > "59" Then
'        Mensajes "La cuenta final debe ser una cuenta de balance", vbOKOnly + vbInformation
'        Exit Function
'    End If
    
    If CE(tdbcMes.Text) = "" Then
        Mensajes "Seleccione el mes del Tipo del rpoceso"
        Exit Function
    End If
    
    ' *** Validar tipo de cambio
    If NumeroLleno(tdbnTipoCambioC, "Tipo de Cambio Compras") = False Then Exit Function
    If NumeroLleno(tdbnTipoCambioV, "Tipo de Cambio Ventas") = False Then Exit Function
    
'    If tdbtCuentaHasta.Text < tdbtCuentaDesde Then
'        Mensajes "La cuenta inicial debe ser menor o igual al la cuenta mayor", vbOKOnly + vbInformation
'        Exit Function
'    End If
'
'    If CE(tdbtCuentaHasta.Text) = "" Or CE(tdbtCuentaDesde) = "" Then
'        Mensajes "Ingrese los numeros de cuentas iniciales y finales", vbOKOnly + vbInformation
'        Exit Function
'    End If

    If CierreMes(tdbcMes.BoundText) Then
        Mensajes "El mes esta bloqueado no se puede generar los asientos automaticos"
        Exit Function
    End If
    
    Validar = True
End Function

Private Sub cmdGenerar_Click()
 Dim respuesta  As String
 Dim sql As String
    If ActualizaTCC(False) = False Then Exit Sub
    If ActualizaTCV(False) = False Then Exit Sub
    
    If Validar = False Then Exit Sub
    
    'If tdbcLibro.BoundText = "01" Or tdbcLibro.BoundText = "02" Or tdbcLibro.BoundText = "04" Or tdbcLibro.BoundText = "07" Or tdbcLibro.BoundText = "08" Then
    sql = "select * from CNT_CIERRE where Emp_cCodigo = '" & gsEmpresa & "' and Pan_cAnio = '" & gsAnio & "' and TipoLibro = '07' and Per_cPeriodo = '" & Me.tdbcMes.BoundText & "' and Cic_cEstado = 'I'"
    If ExisteDato(sql) = True Then Mensajes "No se puede procesar la Dif. De Cambio, debido a que el periodo se encuentra bloqueado", vbInformation: Exit Sub

    sql = "select * from CNT_CIERRE where Emp_cCodigo = '" & gsEmpresa & "' and Pan_cAnio = '" & gsAnio & "' and TipoLibro = '07'"
    If ExisteDato(sql) = True Then
        sql = "select * from CNT_lIBROSGENERADOS where Emp_cCodigo = '" & gsEmpresa & "' and Pan_cAnio = '" & gsAnio & "' and Per_cPeriodo = '" & Me.tdbcMes.BoundText & "' and Lib_cTipoLibro = '03' and Estado ='A'"
        If ExisteDato(sql) = True Then
            Mensajes "Esta corrección modificará los datos ingresados, la misma que será informada a la SUNAT en el período " + UCase(MonthName(Month(lsFecha))) + " del ejercicio " + Str(Year(lsFecha)) + "."
            If MsgBox("Desea continuar..?", vbQuestion + vbOKCancel, gsNombreModulo) = vbCancel Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
    End If
    
    respuesta = MsgBox("Desea Generar Asiento por Tipo de Cambio del mes seleccionado", vbYesNo + vbQuestion, "Confirmar Generar Asiento")
    If respuesta = vbYes Then
        Screen.MousePointer = vbHourglass
        Dim clsMante As clsMantoTablas
        Dim lArrMnt(10) As Variant
        
        Set clsMante = New clsMantoTablas
        ' *** Generando el Asiento
        lArrMnt(0) = gsEmpresa
        lArrMnt(1) = gsAnio
        lArrMnt(2) = Me.tdbcMes.BoundText
        lArrMnt(3) = Me.tdbcLibro.BoundText
        lArrMnt(4) = Me.tdbnTipoCambioC.Value
        lArrMnt(5) = gsUsuario
        lArrMnt(6) = ""
        lArrMnt(7) = ""
        lArrMnt(8) = ""
        lArrMnt(9) = Me.tdbnTipoCambioV.Value
        lArrMnt(10) = ""
                
        If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GeneraDifCambio", lArrMnt(), False) = False Then
            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        
        lArrMnt(10) = "009" 'cuentas por cobrar

        If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GeneraDcmtosDifCambio", lArrMnt(), False) = False Then
            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        lArrMnt(10) = "010" 'cuentas por pagar

        If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GeneraDcmtosDifCambio", lArrMnt(), True) = False Then
            Mensajes "El proceso no ha concluido. Verificar...", vbInformation
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        Set clsMante = Nothing
        Screen.MousePointer = vbDefault
        Mensajes "El proceso se ejecuto con exito", vbInformation
        
    End If
End Sub

Private Sub pCargaCfgLibro()
    Dim clDatos As New clsMantoTablas
    Dim rsArreglo As ADODB.Recordset
    Dim arrDatos() As Variant, sqlver As String

    sqlver = "SELECT * From CNT_CONFIG_LIBROS WHERE Emp_cCodigo = '" & gsEmpresa & "' and Pan_cAnio='" & gsAnio & "'"
    arrDatos = Array(sqlver)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If Not rsArreglo Is Nothing Then
       lsLibroDif = CE(rsArreglo("Cfl_cDifCam"))
       
    End If
    
    Call CerrarRecordSet(rsArreglo)
End Sub

Private Sub Form_Load()
    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))
    
    
    On Error GoTo serror
    Dim sqlcombos As String
    Call pCargaCfgLibro
    
    Call Centrar_form(Me)
    Call LlenaComboMesAddItem(tdbcMes)
    
    
    ' *** Llenando los libros
    sqlcombos = "SELECT LIB_CTIPOLIBRO, LIB_CDESCRIPCION FROM CNT_LIBRO_OPERA " & _
                "WHERE EMP_CCODIGO =  '" & gsEmpresa & "' AND LIB_CTIPOLIBRO IN ('" & lsLibroDif & "') AND Pan_cAnio = '" & gsAnio & "' ORDER BY LIB_CDESCRIPCION "
    LlenarComboAddItem tdbcLibro, sqlcombos
    
    
    DoEvents
    
    tdbnTipoCambioC.Value = CargaTCxMeses(1)
    tdbnTipoCambioV.Value = CargaTCxMeses(2)
  
    
    Dim Mes As String
    If gsPeriodo <> "" Then
        
        If gsPeriodo = "00" Then
        Mes = "01"
        End If
        If gsPeriodo > "12" Then
        Mes = "12"
        End If
        If gsPeriodo > "00" And gsPeriodo < "13" Then
        Mes = gsPeriodo
        End If
        tdbcMes.BoundText = Mes
        'tdbcMes.Row = NE(Mes) - 1
        'tdbcMes.Bookmark = NE(Mes) - 1
        'pSetFocus tdbcMes
        'tdbcMes.Refresh
    End If
    
    
    If Mid(gsGrupo, 3, 1) <> "1" And gsGrupo <> gsPrivilegioAdmin Then
        Me.cmdGenerar.Enabled = False
    Else
        Me.cmdGenerar.Enabled = True
    End If
    
    On Error Resume Next
    tdbcLibro.Bookmark = 0

    
    tdbcMes.ReBind
    tdbcLibro.ReBind

    'tdbnTipoCambio.ReadOnly = True
    Exit Sub
serror:
    'Mensajes Err.Description
End Sub

Private Sub tdbcLibro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then pSendKeys "{tab}"
End Sub

Private Sub tdbcMes_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        
        If CE(tdbcMes.Text) = "" Then
            Mensajes "Seleccione el mes del tipo de cambio", vbOKOnly + vbInformation
            tdbcMes.Bookmark = 0
            pSetFocus tdbcMes
            Exit Sub
        End If
               
    End If

End Sub

Private Sub ActualizaTCMensual(periodo As String, nTipo As Integer, nvalor As Double)
    If Val(periodo) < 1 And Val(periodo) > 12 Then Exit Sub
    
    Dim sql As String
    Dim arrDatos()  As Variant
    Dim clDatos As clsMantoTablas
    Dim clDatosEx As New ClsFuncionesExecute
    Dim rsAddItem  As New ADODB.Recordset
    
    On Error GoTo strError
    
    sql = "select * from CNT_TIPO_CAMBIO_MENSUAL " & _
          "where emp_ccodigo='" & gsEmpresa & "' and pan_canio='" & gsAnio & "' and " & _
          "tca_cmoneda='" & gsMonedaExt & "' and tca_ctipo='" & nTipo & "'"
    
    Set clDatos = New clsMantoTablas
    arrDatos = Array(sql)
    Set rsAddItem = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())

    sql = "insert into CNT_TIPO_CAMBIO_MENSUAL(emp_ccodigo, pan_canio, tca_ctipo, tca_cmoneda, tca_c" & Left(NombreMes(periodo), 3) & ") values (" & _
          "'" & gsEmpresa & "','" & gsAnio & "','" & nTipo & "','" & gsMonedaExt & "'," & nvalor & ") "

    If Not rsAddItem Is Nothing Then
        If rsAddItem.State = adStateOpen Then
            If Not (rsAddItem.BOF And rsAddItem.EOF) Then
                If rsAddItem.RecordCount > 0 Then
                    sql = "Update CNT_TIPO_CAMBIO_MENSUAL set Tca_c" & Left(NombreMes(periodo), 3) & "=" & nvalor & " " & _
                          "where emp_ccodigo='" & gsEmpresa & "' and pan_canio='" & gsAnio & "' and " & _
                          "tca_cmoneda='" & gsMonedaExt & "' and tca_ctipo='" & nTipo & "'"
                End If
            End If
        End If
    End If
    
    clDatosEx.pEjecutaSQL (sql)
    Set clDatosEx = Nothing
    Mensajes "Fue actualizado el tipo de cambio de cierre de este mes", vbOKOnly + vbInformation
    
    
    CerrarRecordSet rsAddItem
    Set clDatos = Nothing
    
    Exit Sub
strError:
    Mensajes "No se actualizo el tipo de cambio de cierre de este mes", vbOKOnly + vbInformation
    CerrarRecordSet rsAddItem
    Set clDatosEx = Nothing
    Set clDatos = Nothing
End Sub

Private Function ActualizaTCC(bMensaje As Boolean) As Boolean
    ActualizaTCC = False
    Dim TipoCambio As Double
    
    If NE(tdbnTipoCambioC.Value) = 0 Then
        Mensajes "Ingrese un tipo de cambio COMPRA valido", vbOKOnly + vbInformation
        pSetFocus tdbnTipoCambioC
        Exit Function
    Else
        TipoCambio = CargaTCxMeses(1)
        If TipoCambio = 0 Then
            
            ActualizaTCMensual tdbcMes.BoundText, 1, NE(tdbnTipoCambioC.Value)
        Else
            If NE(Me.tdbnTipoCambioC.Text) <> NE(TipoCambio) Then
                Mensajes "El tipo de cambio COMPRA se actualizara al tipo de cambio ingresado para este mes", vbOKOnly + vbInformation
                ActualizaTCMensual tdbcMes.BoundText, 1, NE(tdbnTipoCambioC.Value)
            Else
                Me.tdbnTipoCambioC.Value = TipoCambio
            End If
                        
        End If
    End If

    ActualizaTCC = True
End Function

Private Sub tdbnTipoCambioC_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 Then
        If ActualizaTCC(True) = False Then KeyCode = 0
    End If
End Sub

Private Function ActualizaTCV(bMensaje As Boolean) As Boolean
    ActualizaTCV = False
    Dim TipoCambio As Double
    
    If NE(tdbnTipoCambioV.Value) = 0 Then
        Mensajes "Ingrese un tipo de cambio VENTA valido", vbOKOnly + vbInformation
        pSetFocus tdbnTipoCambioV
        Exit Function
    Else
        TipoCambio = CargaTCxMeses(2)
        If TipoCambio = 0 Then
            
            ActualizaTCMensual tdbcMes.BoundText, 2, NE(tdbnTipoCambioV.Value)
        Else
            If NE(Me.tdbnTipoCambioV.Text) <> NE(TipoCambio) Then
                Mensajes "El tipo de cambio VENTA se actualizara al tipo de cambio ingresado para este mes", vbOKOnly + vbInformation
                ActualizaTCMensual tdbcMes.BoundText, 2, NE(tdbnTipoCambioV.Value)
            Else
                Me.tdbnTipoCambioV.Value = TipoCambio
            End If
            
        End If
    End If
    
    ActualizaTCV = True
End Function

Private Sub tdbnTipoCambioV_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim TipoCambio As Double
    If KeyCode = 13 Then
        If ActualizaTCV(True) = False Then KeyCode = 0
    End If
End Sub
