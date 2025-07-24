VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmManRelacionCuentaRatio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relacion de Cuenta de Ratio con Plan de Cuenta"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8565
   Icon            =   "frmManRelacionCuentaRatio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5580
   ScaleWidth      =   8565
   Begin VB.CheckBox chkNegativo 
      Caption         =   "Saldo Negativo"
      Height          =   285
      Left            =   2190
      TabIndex        =   9
      Top             =   1725
      Value           =   1  'Checked
      Width           =   1470
   End
   Begin VB.CheckBox chkPositivo 
      Caption         =   "Saldo Positivo"
      Height          =   285
      Left            =   345
      TabIndex        =   8
      Top             =   1740
      Value           =   1  'Checked
      Width           =   1470
   End
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   300
      TabIndex        =   3
      Top             =   2070
      Width           =   7830
      Begin TrueOleDBGrid70.TDBGrid tdbgDatos 
         Height          =   2565
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   4524
         _LayoutType     =   4
         _RowHeight      =   19
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Codigo"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Cta Ratio"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Cuenta"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Nombre de Cuenta"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   20
         Columns(4)._MaxComboItems=   5
         Columns(4).ValueItems(0)._DefaultItem=   0
         Columns(4).ValueItems(0).Value=   "1"
         Columns(4).ValueItems(0).Value.vt=   8
         Columns(4).ValueItems(0).DisplayValue=   "-1"
         Columns(4).ValueItems(0).DisplayValue.vt=   8
         Columns(4).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
         Columns(4).ValueItems(1)._DefaultItem=   0
         Columns(4).ValueItems(1).Value=   "0"
         Columns(4).ValueItems(1).Value.vt=   8
         Columns(4).ValueItems(1).DisplayValue=   "0"
         Columns(4).ValueItems(1).DisplayValue.vt=   8
         Columns(4).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
         Columns(4).ValueItems.Count=   2
         Columns(4).Caption=   "Saldo +"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   20
         Columns(5)._MaxComboItems=   5
         Columns(5).ValueItems(0)._DefaultItem=   0
         Columns(5).ValueItems(0).Value=   "1"
         Columns(5).ValueItems(0).Value.vt=   8
         Columns(5).ValueItems(0).DisplayValue=   "-1"
         Columns(5).ValueItems(0).DisplayValue.vt=   8
         Columns(5).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
         Columns(5).ValueItems(1)._DefaultItem=   0
         Columns(5).ValueItems(1).Value=   "0"
         Columns(5).ValueItems(1).Value.vt=   8
         Columns(5).ValueItems(1).DisplayValue=   "0"
         Columns(5).ValueItems(1).DisplayValue.vt=   8
         Columns(5).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
         Columns(5).ValueItems.Count=   2
         Columns(5).Caption=   "Saldo -"
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   6
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=6"
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
         Splits(0)._ColumnProps(15)=   "Column(2).Width=2302"
         Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=2223"
         Splits(0)._ColumnProps(18)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(19)=   "Column(2)._ColStyle=532"
         Splits(0)._ColumnProps(20)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(21)=   "Column(3).Width=7382"
         Splits(0)._ColumnProps(22)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(3)._WidthInPix=7303"
         Splits(0)._ColumnProps(24)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(25)=   "Column(3)._ColStyle=532"
         Splits(0)._ColumnProps(26)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(27)=   "Column(4).Width=1085"
         Splits(0)._ColumnProps(28)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(29)=   "Column(4)._WidthInPix=1005"
         Splits(0)._ColumnProps(30)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(31)=   "Column(4)._ColStyle=529"
         Splits(0)._ColumnProps(32)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(33)=   "Column(5).Width=873"
         Splits(0)._ColumnProps(34)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(35)=   "Column(5)._WidthInPix=794"
         Splits(0)._ColumnProps(36)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(37)=   "Column(5)._ColStyle=529"
         Splits(0)._ColumnProps(38)=   "Column(5).Order=6"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         EmptyRows       =   -1  'True
         CellTipsWidth   =   0
         DeadAreaBackColor=   12632256
         RowDividerColor =   12632256
         RowSubDividerColor=   12632256
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.fgcolor=&H80000008&,.bold=0,.fontsize=825"
         _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Arial"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&HCA570B&"
         _StyleDefs(11)  =   ":id=2,.fgcolor=&H80000014&,.bold=0,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(12)  =   ":id=2,.strikethrough=0,.charset=0"
         _StyleDefs(13)  =   ":id=2,.fontname=Arial"
         _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(15)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(16)  =   ":id=3,.fontname=Arial"
         _StyleDefs(17)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(18)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1,.fgcolor=&H800000&"
         _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1"
         _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1,.valignment=2,.bgcolor=&HF1EFEB&"
         _StyleDefs(26)  =   ":id=13,.fgcolor=&H80000008&"
         _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.alignment=2"
         _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=62,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=59,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=60,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=61,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=46,.parent=13,.alignment=2"
         _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
         _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=54,.parent=13,.alignment=2"
         _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
         _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
         _StyleDefs(62)  =   "Named:id=33:Normal"
         _StyleDefs(63)  =   ":id=33,.parent=0"
         _StyleDefs(64)  =   "Named:id=34:Heading"
         _StyleDefs(65)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(66)  =   ":id=34,.wraptext=-1"
         _StyleDefs(67)  =   "Named:id=35:Footing"
         _StyleDefs(68)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(69)  =   "Named:id=36:Selected"
         _StyleDefs(70)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(71)  =   "Named:id=37:Caption"
         _StyleDefs(72)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(73)  =   "Named:id=38:HighlightRow"
         _StyleDefs(74)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(75)  =   "Named:id=39:EvenRow"
         _StyleDefs(76)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(77)  =   "Named:id=40:OddRow"
         _StyleDefs(78)  =   ":id=40,.parent=33"
         _StyleDefs(79)  =   "Named:id=41:RecordSelector"
         _StyleDefs(80)  =   ":id=41,.parent=34"
         _StyleDefs(81)  =   "Named:id=42:FilterBar"
         _StyleDefs(82)  =   ":id=42,.parent=33"
      End
   End
   Begin TrueOleDBList70.TDBCombo tdbcReporte 
      Height          =   300
      Left            =   1860
      TabIndex        =   0
      Tag             =   "enabled"
      Top             =   885
      Width           =   3405
      _ExtentX        =   6006
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=688"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=609"
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
      _PropDict       =   $"frmManRelacionCuentaRatio.frx":1982
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
   Begin TDBText6Ctl.TDBText tdbtCtaDestino 
      Height          =   315
      Left            =   1860
      TabIndex        =   1
      Tag             =   "_"
      Top             =   1230
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   556
      Caption         =   "frmManRelacionCuentaRatio.frx":1A09
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmManRelacionCuentaRatio.frx":1A75
      Key             =   "frmManRelacionCuentaRatio.frx":1A93
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
   Begin TDBText6Ctl.TDBText tdbtNombreDestino 
      Height          =   315
      Left            =   3405
      TabIndex        =   2
      Top             =   1230
      Width           =   4755
      _Version        =   65536
      _ExtentX        =   8387
      _ExtentY        =   556
      Caption         =   "frmManRelacionCuentaRatio.frx":1AE5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmManRelacionCuentaRatio.frx":1B51
      Key             =   "frmManRelacionCuentaRatio.frx":1B6F
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
      Format          =   ""
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
   Begin MSForms.CommandButton cmdInsertarItem 
      Height          =   390
      Left            =   4680
      TabIndex        =   10
      Top             =   1665
      Width           =   1665
      Caption         =   " Insertar Item"
      PicturePosition =   327683
      Size            =   "2937;688"
      Picture         =   "frmManRelacionCuentaRatio.frx":1BC1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdEliminaItem 
      Height          =   390
      Left            =   6480
      TabIndex        =   11
      Top             =   1665
      Width           =   1665
      Caption         =   " Eliminar Item"
      PicturePosition =   327683
      Size            =   "2937;688"
      Picture         =   "frmManRelacionCuentaRatio.frx":215B
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Cuenta Contable"
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
      Left            =   300
      TabIndex        =   7
      Top             =   1230
      Width           =   1395
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Cuenta de Ratio"
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
      Left            =   300
      TabIndex        =   6
      Top             =   870
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   300
      Picture         =   "frmManRelacionCuentaRatio.frx":26F5
      Top             =   270
      Width           =   480
   End
   Begin VB.Label Label4 
      Caption         =   "Para relacionar una cuenta del Plan Contable a una Cuenta de Ratio. Seleccionelos de las Listas y presione el Boton Insertar Item"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   1380
      TabIndex        =   5
      Top             =   270
      Width           =   6135
   End
End
Attribute VB_Name = "frmManRelacionCuentaRatio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lControl As String
Dim lArrDatos As New XArrayDB
Dim lArrMnt() As Variant

Dim gsGrupo As String
Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

Private Sub cmdEliminaItem_Click()
    ' *** Eliminar los datos; segun el q esta seleccionado
    Dim respuesta As String
    If Trim(tdbgDatos.Columns(0).Value) <> "" Then
        respuesta = MsgBox("Desea eliminar definitivamente el registro Seleccionado", vbYesNo + vbQuestion + vbDefaultButton2, "Confirmar Eliminar Registro")
        If respuesta = vbYes Then
            Dim clsMante As clsMantoTablas
            Set clsMante = New clsMantoTablas
            ' *** Eliminando la Cuenta
            Screen.MousePointer = vbHourglass
            ReDim lArrMnt(8) As Variant
            lArrMnt(0) = "ELIMINAR"             ' Accion
            lArrMnt(1) = gsEmpresa              ' Empresa
            lArrMnt(2) = gsAnio                 ' Año de Trabajo
            lArrMnt(3) = tdbgDatos.Columns(2).Value  ' Codigo Reporte
            lArrMnt(4) = tdbgDatos.Columns(0).Value  ' Cuenta
        '    lArrMnt(3) = tdbtCtaDestino         ' Cuenta
        '    lArrMnt(4) = tdbcReporte.BoundText  ' Codigo Reporte
            If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaPlanGestion", lArrMnt(), True) = False Then
                Mensajes "El proceso no se ha realizado. Verificar...", vbInformation
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            Call CargaTabla
            Screen.MousePointer = vbDefault
            Mensajes "Registro ha sido eliminado", vbInformation
            tdbgDatos.ReBind
        End If
    Else
        Mensajes "Debe Seleccionar el registro a eliminar...", vbInformation
    End If
    ' ***
End Sub

Private Sub cmdInsertarItem_Click()
    Dim clsMante As clsMantoTablas
    Dim condicion As Boolean
    ' *** Graba la Relacion establecida en el año establecido
    Set clsMante = New clsMantoTablas
    On Local Error GoTo ErrorEjecucion
    ReDim lArrMnt(8) As Variant
    lArrMnt(0) = "INSERTAR"             ' Accion
    lArrMnt(1) = gsEmpresa              ' Empresa
    lArrMnt(2) = gsAnio                 ' Año de Trabajo
    lArrMnt(3) = tdbtCtaDestino         ' Cuenta
    lArrMnt(4) = tdbcReporte.BoundText  ' Codigo Reporte
    lArrMnt(5) = chkPositivo.Value      ' Codigo Reporte
    lArrMnt(6) = chkNegativo.Value      ' Codigo Reporte
    lArrMnt(7) = "A"                    ' Estado
    lArrMnt(8) = gsUsuario              ' Usuario
    If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaPlanGestion", lArrMnt(), True) = False Then
        Mensajes "El proceso no ha concluido. Verificar...", vbInformation
        Exit Sub
    End If
    CargaTabla
    ' ***
    Mensajes "La relacion se inserto con exito...", vbInformation
    pSetFocus tdbtCtaDestino
    tdbgDatos.ReBind
    chkPositivo.Value = 1
    chkNegativo.Value = 1
    Exit Sub
ErrorEjecucion:
    Mensajes Str(Err.Number) & Err.Description, vbInformation
End Sub

Private Sub Form_Load()
    Me.Top = (frmMDIConta.ScaleHeight - Me.Height) / 2
    Me.Left = (frmMDIConta.ScaleWidth - Me.Width) / 2
    
    ' *** Llenando el tipo de Reportes
    Call LlenaComboRatio
    
    If Mid(gsGrupo, 3, 1) <> "1" And gsGrupo <> gsPrivilegioAdmin Then
        Me.cmdEliminaItem.Enabled = False
        Me.cmdInsertarItem.Enabled = False
    Else
        Me.cmdEliminaItem.Enabled = True
        Me.cmdInsertarItem.Enabled = True
    End If

End Sub

Private Sub LlenaComboRatio()
    Dim sqlcombos As String
    ' *** Llenando los libros
    sqlcombos = "spCn_GrabaCuentaRatio 'SEL_COMBO', '" & gsEmpresa & "', '', '', '', '' "
    LlenarComboAddItem tdbcReporte, sqlcombos
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set lArrDatos = Nothing
End Sub

Private Sub tdbcReporte_ItemChange()
    CargaTabla
    chkPositivo.Value = 1
    chkNegativo.Value = 1
End Sub

Private Sub CargaTabla()
    Dim sqlcombos As String
    ' *** Llenando las cuentas del Reporte
    sqlcombos = "spCn_GrabaPlanGestion 'SEL_ALLRAT', '" & gsEmpresa & "', '" & gsAnio & "', '', '" & tdbcReporte.BoundText & "', '', '', '', '' "
    Call GridArreglo(lArrDatos, tdbgDatos, sqlcombos)
    ' ***
End Sub

Private Sub tdbcReporte_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then pSendKeys "{tab}"
End Sub

Private Sub tdbgDatos_GotFocus()
tdbgDatos.HighlightRowStyle = "HighlightRow"
End Sub

Private Sub tdbgDatos_LostFocus()
tdbgDatos.HighlightRowStyle = ""
End Sub

Private Sub tdbtCtaDestino_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then Call LlamaBuscar(frmBuscador, "tdbtCtaDestino", lControl, "Cuentas", Me, gsPeriodo, tdbtCtaDestino.Text)
End Sub

Private Sub tdbtCtaDestino_LostFocus()
    ' *** Verificando q cuenta exista y q sea de no titulo
    If tdbtCtaDestino <> "" And Me.Enabled = True Then
        tdbtNombreDestino = ExisteCtaNoTitulo(tdbtCtaDestino, "")
        If tdbtNombreDestino = "" Then pSetFocus tdbtCtaDestino
    End If
End Sub

Public Sub RecibirDatos(ByVal lControl As String, ByVal param0 As String, ByVal param1 As String, ByVal param2 As String)
   ' *** Dependiendo del control
    Select Case lControl
    Case "tdbtCtaDestino", "Cuentas"  ' *** Caso de cliente
        tdbtCtaDestino.Text = Trim(param0)
        Me.tdbtNombreDestino.Text = Trim(param1)
        Unload frmBuscador
        pSetFocus tdbtCtaDestino
    End Select
End Sub


