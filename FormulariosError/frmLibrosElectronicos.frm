VERSION 5.00
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmLibrosElectronicos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libros Electrónicos"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5490
   Icon            =   "frmLibrosElectronicos.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   5490
   Begin VB.CommandButton cmdAyuda 
      Height          =   735
      Left            =   4560
      Picture         =   "frmLibrosElectronicos.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin TrueOleDBList70.TDBCombo tdbcMes 
      Height          =   300
      Left            =   2265
      TabIndex        =   1
      Top             =   840
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
      _PropDict       =   $"frmLibrosElectronicos.frx":12DB
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
   Begin TrueOleDBList70.TDBCombo TDBLibro 
      Height          =   300
      Left            =   2265
      TabIndex        =   0
      Top             =   360
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
      _PropDict       =   $"frmLibrosElectronicos.frx":1362
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
      Left            =   5760
      TabIndex        =   6
      Tag             =   "_"
      Top             =   1200
      Visible         =   0   'False
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
      _PropDict       =   $"frmLibrosElectronicos.frx":13E9
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "LIBRO:"
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
      Left            =   1200
      TabIndex        =   5
      Top             =   405
      Width           =   585
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "MES:"
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
      Left            =   1200
      TabIndex        =   4
      Top             =   885
      Width           =   420
   End
   Begin MSForms.CommandButton cmdGenerar 
      Height          =   435
      Left            =   720
      TabIndex        =   2
      Top             =   1620
      Width           =   2025
      Caption         =   "Generar Libro Electrónico"
      PicturePosition =   327683
      Size            =   "3572;767"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdSalir 
      Height          =   435
      Left            =   2895
      TabIndex        =   3
      Top             =   1620
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
Attribute VB_Name = "frmLibrosElectronicos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sSql As String
Dim RstDetalle As New ADODB.Recordset

Private Sub cmdAyuda_Click()
ShowSearch 2
 '   ShowTopicID 1, 2
End Sub

Private Sub cmdGenerar_Click()
On Error GoTo Control
Dim clDatos As New ClsFuncionesExecute
Dim RsDetalle As ADODB.Recordset
  Screen.MousePointer = vbHourglass

    Dim NombreArchivo As String
    Dim NombreArchivo2 As String
    Dim NombreArchivo3 As String
    Dim ruta As String
    Dim Ruta2 As String
    Dim fso As Object
    Dim i As Integer

    i = 0

    Dim sql As String

    Set fso = CreateObject("Scripting.FileSystemObject")

    Ruta1 = App.Path + "\Libros_Electronicos\"
    Ruta2 = App.Path + "\Backup_LE\"

    If Int(tdbcMes.BoundText) > Month(Date) And gsAnio >= Year(Date) Then Mensajes "No se puede generar el archivo con período posterior a la fecha, Verifique!", vbInformation: Screen.MousePointer = vbDefault: Exit Sub

    sql = "select * from CNT_lIBROSGENERADOS where Emp_cCodigo = '" & gsEmpresa & "' and Pan_cAnio = '" & gsAnio & "' and Per_cPeriodo = '" & tdbcMes.BoundText & "' and Lib_cTipoLibro = '" & TDBLibro.BoundText & "' and Estado ='A'"
    If ExisteDato(sql) = True Then
        MsgBox "El Libro solicitado ya ha sido generado. Consulte en la opción Libros enviados.", vbInformation + vbOKOnly, gsNombreModulo
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    If TDBLibro.BoundText = "DS" Then
        sql = "select * from CNT_lIBROSGENERADOS where Emp_cCodigo = '" & gsEmpresa & "' and Pan_cAnio = '" & gsAnio & "' and Per_cPeriodo = '" & tdbcMes.BoundText & "' and Lib_cTipoLibro = '03' and Estado ='A'"
        If ExisteDato(sql) = True Then
            MsgBox "Usted ya ha generado el Libro Diario. Usted debe de eliminarlo para poder generar el Libro Diario Simplificado.", vbInformation + vbOKOnly, gsNombreModulo
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    ElseIf TDBLibro.BoundText = "03" Then
        sql = "select * from CNT_lIBROSGENERADOS where Emp_cCodigo = '" & gsEmpresa & "' and Pan_cAnio = '" & gsAnio & "' and Per_cPeriodo = '" & tdbcMes.BoundText & "' and Lib_cTipoLibro = 'DS' and Estado ='A'"
        If ExisteDato(sql) = True Then
            MsgBox "Usted ya ha generado el Libro Diario Simplificado. Usted debe de eliminarlo para poder generar el Libro Diario.", vbInformation + vbOKOnly, gsNombreModulo
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If

    
    If fso.FolderExists(Ruta1) = False Then
        If TDBLibro.BoundText = lsLibroCom Then
            fso.CreateFolder (Ruta1)
            Ruta1 = Ruta1 & gsRUC & "-" & gsEmpresa & "\"
            If fso.FolderExists(Ruta1) = False Then
                fso.CreateFolder (Ruta1)
            End If
            Ruta1 = Ruta1 & "Compras\"
            If fso.FolderExists(Ruta1) = False Then
                fso.CreateFolder (Ruta1)
            End If
        ElseIf TDBLibro.BoundText = lsLibroVen Then
            fso.CreateFolder (Ruta1)
            Ruta1 = Ruta1 & gsRUC & "-" & gsEmpresa & "\"
            If fso.FolderExists(Ruta1) = False Then
                fso.CreateFolder (Ruta1)
            End If
            Ruta1 = Ruta1 & "Ventas\"
            If fso.FolderExists(Ruta1) = False Then
                fso.CreateFolder (Ruta1)
            End If
        ElseIf TDBLibro.BoundText = lsLibroDiario And gsDiarioSimplificado = 0 Then
            fso.CreateFolder (Ruta1)
            Ruta1 = Ruta1 & gsRUC & "-" & gsEmpresa & "\"
            If fso.FolderExists(Ruta1) = False Then
                fso.CreateFolder (Ruta1)
            End If
            Ruta1 = Ruta1 & "Diario\"
            If fso.FolderExists(Ruta1) = False Then
                fso.CreateFolder (Ruta1)
            End If
        ElseIf TDBLibro.BoundText = "04" And gsDiarioSimplificado = 0 Then   'Mayor
            fso.CreateFolder (Ruta1)
            Ruta1 = Ruta1 & gsRUC & "-" & gsEmpresa & "\"
            If fso.FolderExists(Ruta1) = False Then
                fso.CreateFolder (Ruta1)
            End If
            Ruta1 = Ruta1 & "Mayor\"
            If fso.FolderExists(Ruta1) = False Then
                fso.CreateFolder (Ruta1)
            End If
        ElseIf TDBLibro.BoundText = "DS" And gsDiarioSimplificado = 1 Then   'Diario Simplificado
            fso.CreateFolder (Ruta1)
            Ruta1 = Ruta1 & gsRUC & "-" & gsEmpresa & "\"
            If fso.FolderExists(Ruta1) = False Then
                fso.CreateFolder (Ruta1)
            End If
            Ruta1 = Ruta1 & "Diario_Simplificado\"
            If fso.FolderExists(Ruta1) = False Then
                fso.CreateFolder (Ruta1)
            End If
        End If
    Else
         If TDBLibro.BoundText = lsLibroCom Then
            Ruta1 = Ruta1 & gsRUC & "-" & gsEmpresa & "\"
            If fso.FolderExists(Ruta1) = False Then
                fso.CreateFolder (Ruta1)
            End If
            Ruta1 = Ruta1 & "Compras\"
            If fso.FolderExists(Ruta1) = False Then
                fso.CreateFolder (Ruta1)
            End If
        ElseIf TDBLibro.BoundText = lsLibroVen Then
            Ruta1 = Ruta1 & gsRUC & "-" & gsEmpresa & "\"
            If fso.FolderExists(Ruta1) = False Then
                fso.CreateFolder (Ruta1)
            End If
            Ruta1 = Ruta1 & "Ventas\"
            If fso.FolderExists(Ruta1) = False Then
                fso.CreateFolder (Ruta1)
            End If
        ElseIf TDBLibro.BoundText = lsLibroDiario And gsDiarioSimplificado = 0 Then
            Ruta1 = Ruta1 & gsRUC & "-" & gsEmpresa & "\"
            If fso.FolderExists(Ruta1) = False Then
                fso.CreateFolder (Ruta1)
            End If
            Ruta1 = Ruta1 & "Diario\"
            If fso.FolderExists(Ruta1) = False Then
                fso.CreateFolder (Ruta1)
            End If
        ElseIf TDBLibro.BoundText = "04" And gsDiarioSimplificado = 0 Then   'Mayor
            Ruta1 = Ruta1 & gsRUC & "-" & gsEmpresa & "\"
            If fso.FolderExists(Ruta1) = False Then
                fso.CreateFolder (Ruta1)
            End If
            Ruta1 = Ruta1 & "Mayor\"
            If fso.FolderExists(Ruta1) = False Then
                fso.CreateFolder (Ruta1)
            End If
        ElseIf TDBLibro.BoundText = "DS" And gsDiarioSimplificado = 1 Then   'Diario Simplificado
            'fso.CreateFolder (ruta1)
            Ruta1 = Ruta1 & gsRUC & "-" & gsEmpresa & "\"
            If fso.FolderExists(Ruta1) = False Then
                fso.CreateFolder (Ruta1)
            End If
            Ruta1 = Ruta1 & "Diario_Simplificado\"
            If fso.FolderExists(Ruta1) = False Then
                fso.CreateFolder (Ruta1)
            End If
        End If
    End If
    
    If fso.FolderExists(Ruta2) = False Then
        If TDBLibro.BoundText = lsLibroCom Then
            fso.CreateFolder (Ruta2)
            Ruta2 = Ruta2 & gsRUC & "-" & gsEmpresa & "\"
            If fso.FolderExists(Ruta2) = False Then
                fso.CreateFolder (Ruta2)
            End If
            
            Ruta2 = Ruta2 & "Compras\"
            If fso.FolderExists(Ruta2) = False Then
                fso.CreateFolder (Ruta2)
            End If
        ElseIf TDBLibro.BoundText = lsLibroVen Then
            fso.CreateFolder (Ruta2)
            Ruta2 = Ruta2 & gsRUC & "-" & gsEmpresa & "\"
            If fso.FolderExists(Ruta2) = False Then
                fso.CreateFolder (Ruta2)
            End If
            Ruta2 = Ruta2 & "Ventas\"
            If fso.FolderExists(Ruta2) = False Then
                fso.CreateFolder (Ruta2)
            End If
        ElseIf TDBLibro.BoundText = lsLibroDiario And gsDiarioSimplificado = 0 Then
            fso.CreateFolder (Ruta2)
            Ruta2 = Ruta2 & gsRUC & "-" & gsEmpresa & "\"
            If fso.FolderExists(Ruta2) = False Then
                fso.CreateFolder (Ruta2)
            End If
            Ruta2 = Ruta2 & "Diario\"
            If fso.FolderExists(Ruta2) = False Then
                fso.CreateFolder (Ruta2)
            End If
        ElseIf TDBLibro.BoundText = "04" And gsDiarioSimplificado = 0 Then 'Mayor
            fso.CreateFolder (Ruta2)
            Ruta2 = Ruta2 & gsRUC & "-" & gsEmpresa & "\"
            If fso.FolderExists(Ruta2) = False Then
                fso.CreateFolder (Ruta2)
            End If
            Ruta2 = Ruta2 & "Mayor\"
            If fso.FolderExists(Ruta2) = False Then
                fso.CreateFolder (Ruta2)
            End If
        ElseIf TDBLibro.BoundText = "DS" And gsDiarioSimplificado = 1 Then   'Diario Simplificado
            fso.CreateFolder (Ruta2)
            Ruta2 = Ruta2 & gsRUC & "-" & gsEmpresa & "\"
            If fso.FolderExists(Ruta2) = False Then
                fso.CreateFolder (Ruta2)
            End If
            Ruta1 = Ruta1 & "Diario_Simplificado\"
            If fso.FolderExists(Ruta2) = False Then
                fso.CreateFolder (Ruta2)
            End If
        End If
    Else
         If TDBLibro.BoundText = lsLibroCom Then
            Ruta2 = Ruta2 & gsRUC & "-" & gsEmpresa & "\"
            If fso.FolderExists(Ruta2) = False Then
                fso.CreateFolder (Ruta2)
            End If
            Ruta2 = Ruta2 & "Compras\"
            If fso.FolderExists(Ruta2) = False Then
                fso.CreateFolder (Ruta2)
            End If
        ElseIf TDBLibro.BoundText = lsLibroVen Then
            Ruta2 = Ruta2 & gsRUC & "-" & gsEmpresa & "\"
            If fso.FolderExists(Ruta2) = False Then
                fso.CreateFolder (Ruta2)
            End If
            Ruta2 = Ruta2 & "Ventas\"
            If fso.FolderExists(Ruta2) = False Then
                fso.CreateFolder (Ruta2)
            End If
        ElseIf TDBLibro.BoundText = lsLibroDiario And gsDiarioSimplificado = 0 Then
            Ruta2 = Ruta2 & gsRUC & "-" & gsEmpresa & "\"
            If fso.FolderExists(Ruta2) = False Then
                fso.CreateFolder (Ruta2)
            End If
            Ruta2 = Ruta2 & "Diario\"
            If fso.FolderExists(Ruta2) = False Then
                fso.CreateFolder (Ruta2)
            End If
        ElseIf TDBLibro.BoundText = "04" And gsDiarioSimplificado = 0 Then 'Mayor
            Ruta2 = Ruta2 & gsRUC & "-" & gsEmpresa & "\"
            If fso.FolderExists(Ruta2) = False Then
                fso.CreateFolder (Ruta2)
            End If
            Ruta2 = Ruta2 & "Mayor\"
            If fso.FolderExists(Ruta2) = False Then
                fso.CreateFolder (Ruta2)
            End If
        ElseIf TDBLibro.BoundText = "DS" And gsDiarioSimplificado = 1 Then   'Diario Simplificado
            'fso.CreateFolder (Ruta2)
            Ruta2 = Ruta2 & gsRUC & "-" & gsEmpresa & "\"
            If fso.FolderExists(Ruta2) = False Then
                fso.CreateFolder (Ruta2)
            End If
            Ruta2 = Ruta2 & "Diario_Simplificado\"
            If fso.FolderExists(Ruta2) = False Then
                fso.CreateFolder (Ruta2)
            End If
        End If
    End If
    
       
    If Not ExistenDatos Then
            If TDBLibro.BoundText = lsLibroCom Then
            NombreArchivo = Ruta1 & "LE" & gsRUC & gsAnio & tdbcMes.BoundText & "00080100001011.txt"
            NombreArchivo2 = Ruta2 & "LE" & gsRUC & gsAnio & tdbcMes.BoundText & "00080100001011.txt"
        ElseIf TDBLibro.BoundText = lsLibroVen Then
            NombreArchivo = Ruta1 & "LE" & gsRUC & gsAnio & tdbcMes.BoundText & "00140100001011.txt"
            NombreArchivo2 = Ruta2 & "LE" & gsRUC & gsAnio & tdbcMes.BoundText & "00140100001011.txt"
        ElseIf TDBLibro.BoundText = lsLibroDiario And gsDiarioSimplificado = 0 Then 'Diario
            NombreArchivo = Ruta1 & "LE" & gsRUC & gsAnio & tdbcMes.BoundText & "00050100001011.txt"
            NombreArchivo2 = Ruta2 & "LE" & gsRUC & gsAnio & tdbcMes.BoundText & "00050100001011.txt"
        ElseIf TDBLibro.BoundText = "04" And gsDiarioSimplificado = 0 Then 'Mayor
            NombreArchivo = Ruta1 & "LE" & gsRUC & gsAnio & tdbcMes.BoundText & "00060100001011.txt"
            NombreArchivo2 = Ruta2 & "LE" & gsRUC & gsAnio & tdbcMes.BoundText & "00060100001011.txt"
        ElseIf TDBLibro.BoundText = "DS" And gsDiarioSimplificado = 1 Then 'Diario Simplificado
            NombreArchivo = Ruta1 & "LE" & gsRUC & gsAnio & tdbcMes.BoundText & "00050200001011.txt"
            NombreArchivo2 = Ruta2 & "LE" & gsRUC & gsAnio & tdbcMes.BoundText & "00050200001011.txt"
        End If

        If MsgBox("¿Está seguro de Procesar el Archivo?", vbYesNo + vbDefaultButton1 + vbQuestion, App.Title) = vbNo Then
           Screen.MousePointer = vbDefault
           Exit Sub
        Else
                CrearFileVacio (NombreArchivo)
                CrearFileVacio (NombreArchivo2)
                
  Dim clDatosLD As clsMantoTablas
  Dim arrDatos() As Variant

  sSql = "select * from CNT_lIBROSGENERADOS where Emp_cCodigo = '" & gsEmpresa & "' and Pan_cAnio = '" & gsAnio & "' and Lib_cTipoLibro = 'LD'"

  Set clDatosLD = New clsMantoTablas
  arrDatos = Array(sSql)
  Set RsDetalle = clDatosLD.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())

   If RsDetalle Is Nothing Then
    EstadoLDOri = "1"
    Else
    EstadoLDDes = "8"
   End If

    If TDBLibro.BoundText = lsLibroDiario And gsDiarioSimplificado = 0 Then 'Diario Detalle de Plan de Cuentas - Diario
            NombreArchivo3 = Ruta1 & "LE" & gsRUC & gsAnio & tdbcMes.BoundText & "00050300001011.txt"
            CrearFileVacio (NombreArchivo3)
            NombreArchivo3 = Ruta2 & "LE" & gsRUC & gsAnio & tdbcMes.BoundText & "00050300001011.txt"
            CrearFileVacio (NombreArchivo3)
            GoTo Reg_LibroElectonico
        End If
    End If
    
    If TDBLibro.BoundText = "DS" And gsDiarioSimplificado = 1 Then 'Diario Detalle de Plan de Cuentas - DS
            NombreArchivo3 = Ruta1 & "LE" & gsRUC & gsAnio & tdbcMes.BoundText & "00050400001011.txt"
            CrearFileVacio (NombreArchivo3)
            NombreArchivo3 = Ruta2 & "LE" & gsRUC & gsAnio & tdbcMes.BoundText & "00050400001011.txt"
            CrearFileVacio (NombreArchivo3)
            GoTo Reg_LibroElectonico
        End If
    End If
    
    If NombreArchivo1 <> "" Or NombreArchivo2 <> "" Or NombreArchivo3 <> "" Then
        GoTo Reg_LibroElectonico
    End If
    
    i = 0
    If MsgBox("¿Está seguro de Procesar el Archivo?", vbYesNo + vbDefaultButton1 + vbQuestion, App.Title) = vbYes Then
        If TDBLibro.BoundText = lsLibroCom Then
            NombreArchivo = Ruta1 & "LE" & gsRUC & gsAnio & tdbcMes.BoundText & "00080100001111.txt"
            NombreArchivo2 = Ruta2 & "LE" & gsRUC & gsAnio & tdbcMes.BoundText & "00080100001111.txt"
        ElseIf TDBLibro.BoundText = lsLibroVen Then
            NombreArchivo = Ruta1 & "LE" & gsRUC & gsAnio & tdbcMes.BoundText & "00140100001111.txt"
            NombreArchivo2 = Ruta2 & "LE" & gsRUC & gsAnio & tdbcMes.BoundText & "00140100001111.txt"
        ElseIf TDBLibro.BoundText = lsLibroDiario And gsDiarioSimplificado = 0 Then 'Diario
            NombreArchivo = Ruta1 & "LE" & gsRUC & gsAnio & tdbcMes.BoundText & "00050100001111.txt"
            NombreArchivo2 = Ruta2 & "LE" & gsRUC & gsAnio & tdbcMes.BoundText & "00050100001111.txt"

        ElseIf TDBLibro.BoundText = "04" And gsDiarioSimplificado = 0 Then  'Mayor
            NombreArchivo = Ruta1 & "LE" & gsRUC & gsAnio & tdbcMes.BoundText & "00060100001111.txt"
            NombreArchivo2 = Ruta2 & "LE" & gsRUC & gsAnio & tdbcMes.BoundText & "00060100001111.txt"
        ElseIf TDBLibro.BoundText = "DS" And gsDiarioSimplificado = 1 Then 'Diario Simplificado
            NombreArchivo = Ruta1 & "LE" & gsRUC & gsAnio & tdbcMes.BoundText & "00050200001111.txt"
            NombreArchivo2 = Ruta2 & "LE" & gsRUC & gsAnio & tdbcMes.BoundText & "00050200001111.txt"
        End If
    Else
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    CrearFileLleno (NombreArchivo)
    CrearFileLleno (NombreArchivo2)
    
    If TDBLibro.BoundText = lsLibroDiario And gsDiarioSimplificado = 0 Then 'Diario Detalle de Plan de Cuentas
        If ExistenDatos(1) Then
            NombreArchivo3 = Ruta1 & "LE" & gsRUC & gsAnio & tdbcMes.BoundText & "00050300001111.txt"
            CrearFileLleno (NombreArchivo3)
            NombreArchivo3 = Ruta2 & "LE" & gsRUC & gsAnio & tdbcMes.BoundText & "00050300001111.txt"
            CrearFileLleno (NombreArchivo3)
        Else
            NombreArchivo3 = Ruta1 & "LE" & gsRUC & gsAnio & tdbcMes.BoundText & "00050300001011.txt"
            CrearFileVacio (NombreArchivo3)
            NombreArchivo3 = Ruta2 & "LE" & gsRUC & gsAnio & tdbcMes.BoundText & "00050300001011.txt"
            CrearFileVacio (NombreArchivo3)
        End If
    End If
    
    If TDBLibro.BoundText = "DS" And gsDiarioSimplificado = 1 Then  'Diario Detalle de Plan de Cuentas
        If ExistenDatos(1) Then
            NombreArchivo3 = Ruta1 & "LE" & gsRUC & gsAnio & tdbcMes.BoundText & "00050400001111.txt"
            CrearFileLleno (NombreArchivo3)
            NombreArchivo3 = Ruta2 & "LE" & gsRUC & gsAnio & tdbcMes.BoundText & "00050400001111.txt"
            CrearFileLleno (NombreArchivo3)
        Else
            NombreArchivo3 = Ruta1 & "LE" & gsRUC & gsAnio & tdbcMes.BoundText & "00050400001011.txt"
            CrearFileVacio (NombreArchivo3)
            NombreArchivo3 = Ruta2 & "LE" & gsRUC & gsAnio & tdbcMes.BoundText & "00050400001011.txt"
            CrearFileVacio (NombreArchivo3)
        End If
    End If

Reg_LibroElectonico:
    sql = "insert into CNT_lIBROSGENERADOS(Emp_cCodigo,Pan_cAnio,Per_cPeriodo,Lib_cTipoLibro,FecCrea,Estado) values('" & gsEmpresa & "', '" & gsAnio & "', '" & tdbcMes.BoundText & "', '" & TDBLibro.BoundText & "', getdate(),'A')"
    clDatos.pEjecutaSQL (sql)
    If TDBLibro.BoundText = "03" Then
        sql = "insert into CNT_lIBROSGENERADOS(Emp_cCodigo,Pan_cAnio,Per_cPeriodo,Lib_cTipoLibro,FecCrea,Estado) values('" & gsEmpresa & "', '" & gsAnio & "', '" & tdbcMes.BoundText & "', 'LD', getdate(),'A')"
        clDatos.pEjecutaSQL (sql)
    End If
  
    If TDBLibro.BoundText <> "04" Or TDBLibro.BoundText <> "DS" Then
        Call GrabarCierre(tdbcMes.BoundText)
    End If
    
    Mensajes "El archivo se ha generado con éxito, en la ruta:" & vbCrLf & vbCrLf & NombreArchivo & vbCrLf & vbCrLf & "Y la copia de seguridad en la ruta:" & vbCrLf & vbCrLf & NombreArchivo2, vbInformation
    
  Screen.MousePointer = vbDefault
  
Exit Sub

Control:
    Set fso = Nothing
  
 MsgBox Err.Description
 Screen.MousePointer = vbDefault
End Sub

Sub CrearFileVacio(nomfile As String)
Open nomfile For Output Shared As #1
Close #1
End Sub

Sub CrearFileLleno(nomfile As String)
        Open nomfile For Output Shared As #1
        With RstDetalle
           If .RecordCount > 0 Then
              While Not .EOF
                 Print #1, !Registro
                 .MoveNext
              Wend
           End If
        End With
        RstDetalle.MoveFirst
        Close #1
End Sub


Private Function GrabarCierre(cPeriodo As String) As Boolean
    ' *** Generar el cierre
    Dim clsMante As clsMantoTablas
    Dim lArrMnt() As Variant
    Dim cMensaje As String
    Dim sql As String
    Set clsMante = New clsMantoTablas
    ' *** Grabando Centro de Costo
    Screen.MousePointer = vbHourglass
    On Local Error GoTo ErrorEjecucion
    ReDim lArrMnt(6) As Variant
    lArrMnt(0) = "INSERTAR"     ' Accion
    lArrMnt(1) = gsEmpresa
    lArrMnt(2) = gsAnio
    lArrMnt(3) = cPeriodo
    lArrMnt(4) = "I"
    lArrMnt(5) = gsUsuario
    lArrMnt(6) = TDBLibro.BoundText
    
    If TDBLibro.BoundText = "03" Then
        sql = "SELECT * FROM CNT_CIERRE WITH(READUNCOMMITTED) WHERE Emp_cCodigo = '" & gsEmpresa & "' and Pan_cAnio = '" & gsAnio & "' and TipoLibro = '01'"
        If ExisteDato(sql) = False Then
            If TDBLibro.BoundText = "03" And cPeriodo <> "12" Then
                If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaCierre", lArrMnt(), True) = False Then
                    GoTo mensaje
                End If
                  
                lArrMnt(6) = "02" 'Ingreso
                If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaCierre", lArrMnt(), True) = False Then
                    GoTo mensaje
                End If
                
                lArrMnt(6) = "04" 'Egresos
                If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaCierre", lArrMnt(), True) = False Then
                    GoTo mensaje
                End If
                
                lArrMnt(6) = "07" 'Dif. Cambio
                If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaCierre", lArrMnt(), True) = False Then
                    GoTo mensaje
                End If
                
                lArrMnt(6) = "01" 'Apertura
                lArrMnt(3) = "00"
                If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaCierre", lArrMnt(), True) = False Then
                    GoTo mensaje
                End If
            End If
        Else
                If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaCierre", lArrMnt(), True) = False Then
                    GoTo mensaje
                End If
                
                lArrMnt(6) = "02" 'Ingreso
                If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaCierre", lArrMnt(), True) = False Then
                    GoTo mensaje
                End If
                
                lArrMnt(6) = "04" 'Egresos
                If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaCierre", lArrMnt(), True) = False Then
                    GoTo mensaje
                End If
                
                lArrMnt(6) = "07" 'Dif. Cambio
                If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaCierre", lArrMnt(), True) = False Then
                    GoTo mensaje
                End If
                
        End If
            If TDBLibro.BoundText = "03" And cPeriodo = "12" Then
                If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaCierre", lArrMnt(), True) = False Then
                    GoTo mensaje
                End If

                lArrMnt(6) = "02" 'Ingreso
                If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaCierre", lArrMnt(), True) = False Then
                    GoTo mensaje
                End If
                
                lArrMnt(6) = "04" 'Egresos
                If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaCierre", lArrMnt(), True) = False Then
                    GoTo mensaje
                End If
                
                lArrMnt(6) = "07" 'Dif. Cambio
                If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaCierre", lArrMnt(), True) = False Then
                    GoTo mensaje
                End If
                
                lArrMnt(6) = "08" 'Cierre
                lArrMnt(3) = "13" 'Ajuste
                If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaCierre", lArrMnt(), True) = False Then
                    GoTo mensaje
                End If
                
                lArrMnt(3) = "14" 'Cierre
                If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaCierre", lArrMnt(), True) = False Then
                    GoTo mensaje
                End If
            End If
        Else
                If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaCierre", lArrMnt(), True) = False Then
                    GoTo mensaje
                End If
    End If
        Exit Function
    
mensaje:
    Mensajes "El proceso no ha concluido. Verificar...", vbInformation
    Screen.MousePointer = vbDefault
    Exit Function
    
'    'PGBV - 02012013
'    Dim Sql As String
'
'    Sql = "SELECT * FROM CNT_CIERRE WITH(READUNCOMMITTED) WHERE Emp_cCodigo = '" & gsEmpresa & "' and Pan_cAnio = '" & gsAnio & "' and Per_cPeriodo = '" & gsPeriodo & "' and TipoLibro = '" & TDBLibro.BoundText & "'"
'    If ExisteDato(Sql) = True Then
'        lArrMnt(0) = "EDITAR"     ' Accion
'    Else
'        If optCerrar.Value = True Then
'            lArrMnt(0) = "INSERTAR"     ' Accion
'        End If
'    End If
   
ErrorEjecucion:
    Mensajes Str(Err.Number) & " " & Err.Description, vbInformation
End Function

Private Function ExistenDatos(Optional LEDD As Integer) As Boolean
On Error GoTo Err_Data

  Dim clDatos As clsMantoTablas
  Dim arrDatos() As Variant
  
  ExistenDatos = False
    
    If TDBLibro.BoundText = lsLibroCom Then
        sSql = "spCn_LibroElectronicoCompras2'" & gsEmpresa & "', '" & gsAnio & "','" & PrimerDiaMes(tdbcMes.BoundText, gsAnio) & "','" & UltimoDiaMes(tdbcMes.BoundText, gsAnio) & "','" & tdbcMoneda.BoundText & "','" & tdbcMes.BoundText & "'"
'Nuevo Formato del Libro Compras
'        sSql = "spCn_LibroElectronicoCompras3 '" & gsEmpresa & "', '" & gsAnio & "', '" & tdbcMoneda.BoundText & "', '" & tdbcMes.BoundText & "', 1"
    ElseIf TDBLibro.BoundText = lsLibroVen Then
        sSql = "spCn_LibroElectronicoVentas2'" & gsEmpresa & "', '" & gsAnio & "','" & PrimerDiaMes(tdbcMes.BoundText, gsAnio) & "','" & UltimoDiaMes(tdbcMes.BoundText, gsAnio) & "','" & tdbcMoneda.BoundText & "','" & tdbcMes.BoundText & "'"
        'Nuevo Formato de Libro Ventas
'        sSql = "spCn_LibroElectronicoVentas3 '" & gsEmpresa & "', '" & gsAnio & "', '" & tdbcMoneda.BoundText & "', '" & tdbcMes.BoundText & "'"
    ElseIf TDBLibro.BoundText = lsLibroDiario And gsDiarioSimplificado = 0 And LEDD = 0 Then    'Diario 1
        sSql = "spCn_RptDiarioElectronico2'" & gsEmpresa & "', '" & gsAnio & "','" & tdbcMes.BoundText & "','" & PrimerDiaMes(tdbcMes.BoundText, gsAnio) & "','" & UltimoDiaMes(tdbcMes.BoundText, gsAnio) & "','" & tdbcMoneda.BoundText & "','" & tdbcMes.BoundText & "','1'"
'        sSql = "spCn_RptDiarioElectronico3'" & gsEmpresa & "', '" & gsAnio & "','" & tdbcMes.BoundText & "','" & PrimerDiaMes(tdbcMes.BoundText, gsAnio) & "','" & UltimoDiaMes(tdbcMes.BoundText, gsAnio) & "','" & tdbcMoneda.BoundText & "','" & tdbcMes.BoundText & "','1'"
     ElseIf LEDD = 1 Then
        sSql = "DiarioDetalleElectronico'" & gsEmpresa & "', '" & gsAnio & "','" & tdbcMes.BoundText & "','LD'"
'        sSql = "DiarioDetalleElectronico2'" & gsEmpresa & "', '" & gsAnio & "','" & tdbcMes.BoundText & "','LD'"
    ElseIf TDBLibro.BoundText = "04" And gsDiarioSimplificado = 0 Then  'Mayor
        sSql = "spCn_RptMayorElectronico2'" & gsEmpresa & "', '" & gsAnio & "','" & tdbcMes.BoundText & "','" & tdbcMes.BoundText & "','" & tdbcMoneda.BoundText & "','','','','',''"
    ElseIf TDBLibro.BoundText = "DS" And gsDiarioSimplificado = 1 Then  'Diario Simplificado
        sSql = "spCn_RptDiarioElectronico2'" & gsEmpresa & "', '" & gsAnio & "','" & tdbcMes.BoundText & "','" & PrimerDiaMes(tdbcMes.BoundText, gsAnio) & "','" & UltimoDiaMes(tdbcMes.BoundText, gsAnio) & "','" & tdbcMoneda.BoundText & "','" & tdbcMes.BoundText & "','1'"
'        sSql = "spCn_RptDiarioElectronico3'" & gsEmpresa & "', '" & gsAnio & "','" & tdbcMes.BoundText & "','" & PrimerDiaMes(tdbcMes.BoundText, gsAnio) & "','" & UltimoDiaMes(tdbcMes.BoundText, gsAnio) & "','" & tdbcMoneda.BoundText & "','" & tdbcMes.BoundText & "','1'"
    End If

  Set clDatos = New clsMantoTablas
  arrDatos = Array(sSql)
  Set RstDetalle = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
  With RstDetalle
  If .State <> 0 Then
   If .RecordCount = 0 And LEDD <> 1 Then
    MsgBox "No existen Registros en el Periodo señalado.", vbInformation, App.Title
    Screen.MousePointer = vbDefault
    Exit Function
   End If
  ElseIf LEDD <> 1 Then
    MsgBox "No existen Registros en el Periodo señalado.", vbInformation, App.Title
    Screen.MousePointer = vbDefault
    Exit Function
  Else
    Screen.MousePointer = vbDefault
    Exit Function
  End If
  End With

  ExistenDatos = True
  
Exit Function
   
Err_Data:
   Screen.MousePointer = vbDefault
   MsgBox Err.Description, vbCritical, App.Title
End Function

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Dim sqlcombos As String
    Call Centrar_form(Me)
    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))
    ' *** Llenando el tipo de Moneda
    sqlcombos = "SELECT Mon_cCodigo, Mon_cNombreLargo, Mon_cMNac, Mon_cMExt From CNT_TIPO_MONEDA " & _
                "WHERE Emp_cCodigo = '" & gsEmpresa & "' AND (Mon_cMNac = '1' or Mon_cMExt = '1') " & _
                "ORDER BY Mon_cNombreLargo"
    LlenarComboAddItem tdbcMoneda, sqlcombos
    
    Call BuscarMonedaNacional
    Call LlenarPeriodo
    Call Llenarlibros
    tdbcMes.BoundText = gsPeriodo
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
End Sub

Private Sub LlenarPeriodo()
    Dim i As Integer
    Dim cadena As String
    
    For i = 0 To 11
        tdbcMes.AddItem Format(i + 1, "00") & ";" & UCase(MonthName(i + 1))
    Next
    tdbcMes.Bookmark = 0
    tdbcMes.ListField = "column1"
    tdbcMes.BoundColumn = "column0"
End Sub

Private Sub Llenarlibros()
'    TDBLibro.AddItem "01" & ";" & "APERTURA"
'    TDBLibro.AddItem "02" & ";" & "CAJA INGRESOS"
    If gsDiarioSimplificado = 0 Then
        TDBLibro.AddItem "03" & ";" & "DIARIO"
        TDBLibro.AddItem "04" & ";" & "MAYOR"
    End If
    TDBLibro.AddItem "05" & ";" & "VENTAS"
    TDBLibro.AddItem "06" & ";" & "COMPRAS"
    If gsDiarioSimplificado = 1 Then TDBLibro.AddItem "DS" & ";" & "DIARIO SIMPLIFICADO"
'    TDBLibro.AddItem "07" & ";" & "DIFERENCIA DE CAMBIO"
'    TDBLibro.AddItem "08" & ";" & "CIERRE"
    TDBLibro.Bookmark = 0
    TDBLibro.ListField = "column1"
    TDBLibro.BoundColumn = "column0"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))
End Sub
