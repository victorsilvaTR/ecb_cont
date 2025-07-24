VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmManTipoMoneda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Tipo de Moneda"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7995
   Icon            =   "frmManTipoMoneda.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5505
   ScaleWidth      =   7995
   Begin TabDlg.SSTab SSTCentroCosto 
      Height          =   5100
      Left            =   45
      TabIndex        =   11
      Top             =   405
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   8996
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tipos de Moneda"
      TabPicture(0)   =   "frmManTipoMoneda.frx":0ECA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Mantenimiento de Tipo de Moneda"
      TabPicture(1)   =   "frmManTipoMoneda.frx":0EE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   4455
         Left            =   240
         TabIndex        =   12
         Top             =   555
         Width           =   7215
         Begin TrueOleDBGrid70.TDBGrid tdbgCostos 
            Height          =   3555
            Left            =   360
            TabIndex        =   2
            Top             =   720
            Width           =   6510
            _ExtentX        =   11483
            _ExtentY        =   6271
            _LayoutType     =   4
            _RowHeight      =   19
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Codigo "
            Columns(0).DataField=   "Mon_cCodigo"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Descripción"
            Columns(1).DataField=   "Mon_cNombreLargo"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Abrev."
            Columns(2).DataField=   "Mon_cNombreCorto"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   16
            Columns(3)._MaxComboItems=   5
            Columns(3).ValueItems(0)._DefaultItem=   0
            Columns(3).ValueItems(0).Value=   "1"
            Columns(3).ValueItems(0).Value.vt=   8
            Columns(3).ValueItems(0).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
            Columns(3).ValueItems(0).DisplayValue(0)=   "bHQAAGYDAABCTWYDAAAAAAAANgAAACgAAAAQAAAAEQAAAAEAGAAAAAAAMAMAAAAAAAAAAAAAAAAA"
            Columns(3).ValueItems(0).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
            Columns(3).ValueItems(0).DisplayValue(2)=   "//////////////////9SpUoAlAhrtWP/////////////////////////////////////////////"
            Columns(3).ValueItems(0).DisplayValue(3)=   "//////8YtSkAvSEAlACMvXv///////////////////////////////////////////9rtWMAvSEA"
            Columns(3).ValueItems(0).DisplayValue(4)=   "xikApQAxnDH///////////////////////////////////////////8AnBAAzjEAxikArRAAlACl"
            Columns(3).ValueItems(0).DisplayValue(5)=   "xpT///////////////////////////////////9SpUoAzjEAxikA/2MAzjEAnAAAjAD/////////"
            Columns(3).ValueItems(0).DisplayValue(6)=   "//////////////////////////8YtSkpzloA/2MA/2MAvSEAxikAlACMvXv/////////////////"
            Columns(3).ValueItems(0).DisplayValue(7)=   "//////////////8YxkIA/2MA/2NSpUpSpUoAxikApQAxnDH/////////////////////////////"
            Columns(3).ValueItems(0).DisplayValue(8)=   "//////8ArSEArSH///////8ArRgAxikAlAClxpT/////////////////////////////////////"
            Columns(3).ValueItems(0).DisplayValue(9)=   "//////////8xtUIAxikAnAAAjAD/////////////////////////////////////////////////"
            Columns(3).ValueItems(0).DisplayValue(10)=   "//8AtSEAxikAlACMvXv///////////////////////////////////////////////9SpUoAxikp"
            Columns(3).ValueItems(0).DisplayValue(11)=   "rTkxtUL///////////////////////////////////////////////////8prUpa56UprTmMvXv/"
            Columns(3).ValueItems(0).DisplayValue(12)=   "//////////////////////////////////////////////8xtUIA1kKMvXv/////////////////"
            Columns(3).ValueItems(0).DisplayValue(13)=   "//////////////////////////////////////+lxpT/////////////////////////////////"
            Columns(3).ValueItems(0).DisplayValue(14)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(3).ValueItems(0).DisplayValue(15)=   "//////////////////////////////8="
            Columns(3).ValueItems(0).DisplayValue.vt=   9
            Columns(3).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
            Columns(3).ValueItems(1)._DefaultItem=   0
            Columns(3).ValueItems(1).Value=   "0"
            Columns(3).ValueItems(1).Value.vt=   8
            Columns(3).ValueItems(1).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
            Columns(3).ValueItems(1).DisplayValue(0)=   "bHQAAGYDAABCTWYDAAAAAAAANgAAACgAAAAQAAAAEQAAAAEAGAAAAAAAMAMAAAAAAAAAAAAAAAAA"
            Columns(3).ValueItems(1).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
            Columns(3).ValueItems(1).DisplayValue(2)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(3).ValueItems(1).DisplayValue(3)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(3).ValueItems(1).DisplayValue(4)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(3).ValueItems(1).DisplayValue(5)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(3).ValueItems(1).DisplayValue(6)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(3).ValueItems(1).DisplayValue(7)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(3).ValueItems(1).DisplayValue(8)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(3).ValueItems(1).DisplayValue(9)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(3).ValueItems(1).DisplayValue(10)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(3).ValueItems(1).DisplayValue(11)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(3).ValueItems(1).DisplayValue(12)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(3).ValueItems(1).DisplayValue(13)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(3).ValueItems(1).DisplayValue(14)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(3).ValueItems(1).DisplayValue(15)=   "//////////////////////////////8="
            Columns(3).ValueItems(1).DisplayValue.vt=   9
            Columns(3).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
            Columns(3).ValueItems.Count=   2
            Columns(3).Caption=   "Nac."
            Columns(3).DataField=   "Mon_cMNac"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   16
            Columns(4)._MaxComboItems=   5
            Columns(4).ValueItems(0)._DefaultItem=   0
            Columns(4).ValueItems(0).Value=   "1"
            Columns(4).ValueItems(0).Value.vt=   8
            Columns(4).ValueItems(0).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
            Columns(4).ValueItems(0).DisplayValue(0)=   "bHQAAGYDAABCTWYDAAAAAAAANgAAACgAAAAQAAAAEQAAAAEAGAAAAAAAMAMAAAAAAAAAAAAAAAAA"
            Columns(4).ValueItems(0).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
            Columns(4).ValueItems(0).DisplayValue(2)=   "//////////////////9SpUoAlAhrtWP/////////////////////////////////////////////"
            Columns(4).ValueItems(0).DisplayValue(3)=   "//////8YtSkAvSEAlACMvXv///////////////////////////////////////////9rtWMAvSEA"
            Columns(4).ValueItems(0).DisplayValue(4)=   "xikApQAxnDH///////////////////////////////////////////8AnBAAzjEAxikArRAAlACl"
            Columns(4).ValueItems(0).DisplayValue(5)=   "xpT///////////////////////////////////9SpUoAzjEAxikA/2MAzjEAnAAAjAD/////////"
            Columns(4).ValueItems(0).DisplayValue(6)=   "//////////////////////////8YtSkpzloA/2MA/2MAvSEAxikAlACMvXv/////////////////"
            Columns(4).ValueItems(0).DisplayValue(7)=   "//////////////8YxkIA/2MA/2NSpUpSpUoAxikApQAxnDH/////////////////////////////"
            Columns(4).ValueItems(0).DisplayValue(8)=   "//////8ArSEArSH///////8ArRgAxikAlAClxpT/////////////////////////////////////"
            Columns(4).ValueItems(0).DisplayValue(9)=   "//////////8xtUIAxikAnAAAjAD/////////////////////////////////////////////////"
            Columns(4).ValueItems(0).DisplayValue(10)=   "//8AtSEAxikAlACMvXv///////////////////////////////////////////////9SpUoAxikp"
            Columns(4).ValueItems(0).DisplayValue(11)=   "rTkxtUL///////////////////////////////////////////////////8prUpa56UprTmMvXv/"
            Columns(4).ValueItems(0).DisplayValue(12)=   "//////////////////////////////////////////////8xtUIA1kKMvXv/////////////////"
            Columns(4).ValueItems(0).DisplayValue(13)=   "//////////////////////////////////////+lxpT/////////////////////////////////"
            Columns(4).ValueItems(0).DisplayValue(14)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(4).ValueItems(0).DisplayValue(15)=   "//////////////////////////////8="
            Columns(4).ValueItems(0).DisplayValue.vt=   9
            Columns(4).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
            Columns(4).ValueItems(1)._DefaultItem=   0
            Columns(4).ValueItems(1).Value=   "0"
            Columns(4).ValueItems(1).Value.vt=   8
            Columns(4).ValueItems(1).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
            Columns(4).ValueItems(1).DisplayValue(0)=   "bHQAAGYDAABCTWYDAAAAAAAANgAAACgAAAAQAAAAEQAAAAEAGAAAAAAAMAMAAAAAAAAAAAAAAAAA"
            Columns(4).ValueItems(1).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
            Columns(4).ValueItems(1).DisplayValue(2)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(4).ValueItems(1).DisplayValue(3)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(4).ValueItems(1).DisplayValue(4)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(4).ValueItems(1).DisplayValue(5)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(4).ValueItems(1).DisplayValue(6)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(4).ValueItems(1).DisplayValue(7)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(4).ValueItems(1).DisplayValue(8)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(4).ValueItems(1).DisplayValue(9)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(4).ValueItems(1).DisplayValue(10)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(4).ValueItems(1).DisplayValue(11)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(4).ValueItems(1).DisplayValue(12)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(4).ValueItems(1).DisplayValue(13)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(4).ValueItems(1).DisplayValue(14)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(4).ValueItems(1).DisplayValue(15)=   "//////////////////////////////8="
            Columns(4).ValueItems(1).DisplayValue.vt=   9
            Columns(4).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
            Columns(4).ValueItems.Count=   2
            Columns(4).Caption=   "Ext."
            Columns(4).DataField=   "Mon_cMExt"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Cod. SUNAT"
            Columns(5).DataField=   "Mon_cCodSunat"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   6
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=6"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1085"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1005"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=532"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=3704"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=3625"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=532"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=1164"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1085"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=532"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=1032"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=953"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=529"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=1164"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=1085"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=529"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(31)=   "Column(5).Width=2328"
            Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=2249"
            Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=532"
            Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
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
            _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1,.valignment=2,.bgcolor=&HFFFFFF&"
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
            _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=62,.parent=13"
            _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
            _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
            _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
            _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=32,.parent=13,.alignment=2"
            _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
            _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
            _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
            _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
            _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
            _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=54,.parent=13"
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
         Begin TDBText6Ctl.TDBText tdbtDescripcionBus 
            Height          =   315
            Left            =   1800
            TabIndex        =   9
            Top             =   285
            Width           =   4215
            _Version        =   65536
            _ExtentX        =   7435
            _ExtentY        =   556
            Caption         =   "frmManTipoMoneda.frx":0F02
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManTipoMoneda.frx":0F6E
            Key             =   "frmManTipoMoneda.frx":0F8C
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
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Descripción :"
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
            Left            =   360
            TabIndex        =   13
            Top             =   330
            Width           =   1080
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1215
         Left            =   270
         TabIndex        =   18
         Top             =   480
         Visible         =   0   'False
         Width           =   7200
         Begin VB.OptionButton optContabilidad 
            Caption         =   "Contabilidad Bimoneda"
            Height          =   375
            Index           =   1
            Left            =   2340
            TabIndex        =   1
            Top             =   645
            Width           =   3105
         End
         Begin VB.OptionButton optContabilidad 
            Caption         =   "Contabilidad en Moneda Nacional"
            Height          =   375
            Index           =   0
            Left            =   2340
            TabIndex        =   0
            Top             =   285
            Value           =   -1  'True
            Width           =   3105
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3975
         Left            =   -74775
         TabIndex        =   10
         Top             =   495
         Width           =   6315
         Begin VB.CheckBox chkExtranjera 
            Caption         =   "Moneda Extranjera"
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
            Left            =   1800
            TabIndex        =   8
            Top             =   3000
            Width           =   2055
         End
         Begin VB.CheckBox chkNacional 
            Caption         =   "Moneda Nacional"
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
            Left            =   1800
            TabIndex        =   7
            Top             =   2640
            Width           =   1935
         End
         Begin TDBText6Ctl.TDBText tdbtCodigo 
            Height          =   315
            Left            =   1830
            TabIndex        =   3
            Tag             =   "_"
            Top             =   840
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   556
            Caption         =   "frmManTipoMoneda.frx":0FDE
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManTipoMoneda.frx":104A
            Key             =   "frmManTipoMoneda.frx":1068
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
            AllowSpace      =   0
            Format          =   "a"
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
         Begin TDBText6Ctl.TDBText tdbtDescripcion 
            Height          =   315
            Left            =   1830
            TabIndex        =   5
            Tag             =   "_"
            Top             =   1320
            Width           =   4095
            _Version        =   65536
            _ExtentX        =   7223
            _ExtentY        =   556
            Caption         =   "frmManTipoMoneda.frx":10BA
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManTipoMoneda.frx":1126
            Key             =   "frmManTipoMoneda.frx":1144
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
            Format          =   "a"
            FormatMode      =   1
            AutoConvert     =   -1
            ErrorBeep       =   0
            MaxLength       =   60
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
         Begin TDBText6Ctl.TDBText tdbtDescripcionCorta 
            Height          =   315
            Left            =   1845
            TabIndex        =   6
            Tag             =   "_"
            Top             =   1800
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   556
            Caption         =   "frmManTipoMoneda.frx":1196
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManTipoMoneda.frx":1202
            Key             =   "frmManTipoMoneda.frx":1220
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
            Format          =   "a"
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
         Begin TDBText6Ctl.TDBText tdbtCodSunat 
            Height          =   315
            Left            =   5175
            TabIndex        =   4
            Tag             =   "_"
            Top             =   855
            Width           =   780
            _Version        =   65536
            _ExtentX        =   1376
            _ExtentY        =   556
            Caption         =   "frmManTipoMoneda.frx":1272
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManTipoMoneda.frx":12DE
            Key             =   "frmManTipoMoneda.frx":12FC
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
            AllowSpace      =   0
            Format          =   "a"
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
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Código SUNAT"
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
            Left            =   3825
            TabIndex        =   19
            Top             =   900
            Width           =   1245
         End
         Begin VB.Label lblMante 
            AutoSize        =   -1  'True
            Caption         =   "NUEVO REGISTRO"
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
            Left            =   240
            TabIndex        =   17
            Top             =   360
            Width           =   1515
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Código"
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
            Left            =   240
            TabIndex        =   16
            Top             =   840
            Width           =   600
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Descripción"
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
            Left            =   240
            TabIndex        =   15
            Top             =   1320
            Width           =   990
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Abreviatura"
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
            Left            =   240
            TabIndex        =   14
            Top             =   1800
            Width           =   915
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   80
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoMoneda.frx":134E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoMoneda.frx":1728
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoMoneda.frx":1B02
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoMoneda.frx":1EDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoMoneda.frx":22B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoMoneda.frx":2690
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoMoneda.frx":2A6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoMoneda.frx":2E44
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imglstdisabled 
      Left            =   11025
      Top             =   2475
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoMoneda.frx":3E5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoMoneda.frx":3FB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoMoneda.frx":4112
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoMoneda.frx":426C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoMoneda.frx":43C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoMoneda.frx":4520
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoMoneda.frx":467A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoMoneda.frx":47D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoMoneda.frx":492E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imglstTool 
      Left            =   11025
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoMoneda.frx":4A88
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoMoneda.frx":5022
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoMoneda.frx":55BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoMoneda.frx":5B56
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoMoneda.frx":60F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoMoneda.frx":668A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoMoneda.frx":6C24
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoMoneda.frx":71BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoMoneda.frx":7758
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrOpciones 
      Height          =   330
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   3585
      _ExtentX        =   6324
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "imglstTool"
      DisabledImageList=   "imglstdisabled"
      HotImageList    =   "imglstTool"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo F2"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Ver Datos F3"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Grabar F4"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar F5"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Editar F6"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir F7"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cancelar o Salir ESC"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmManTipoMoneda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lArrMnt() As Variant        ' *** Arreglo para los mantenimientos
Dim lTipoMnt As String          ' *** Tipo de Mantenimiento (insert/update/delete)
Dim lrsTabla As Recordset
Dim lRegElim As Boolean         ' *** Verifica si registro ha sido eliminado desde otra sesion
Dim MonActNac  As String
Dim MonActExt  As String
Dim gsGrupo As String
Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

Private Function ValidaProcesoConversion() As Boolean
    ValidaProcesoConversion = False
    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Set clDatos = New clsMantoTablas
    Dim arrDatos() As Variant
    ' *** Cargando Datos de la Cuenta
    Dim sqlSp As String
    sqlSp = "spCn_ConsultaCuentas 'BUSCAMOVOTROS',  '" & gsEmpresa & "', '" & gsAnio & "'"
    arrDatos = Array(sqlSp)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If Not rsArreglo Is Nothing Then
        If rsArreglo.State = adStateOpen Then
            If NE(rsArreglo.RecordCount) = 1 Then
                ValidaProcesoConversion = True
            Else
                ValidaProcesoConversion = False
            End If
        End If
    End If
    
    Call CerrarRecordSet(rsArreglo)
    Set clDatos = Nothing
    
End Function

Private Sub Form_Resize()
On Error GoTo errHand
    If Me.WindowState <> vbMinimized Then
        
        '*** REDIMENSIONAR SST
        With SSTCentroCosto
            .Width = Me.Width - .Left - 200
            .Height = Me.Height - .Top - 400
            '*** REDIMENSIONAR FRAME PRINCIPAL
            Frame1.Width = .Width - IIf(.TabOrientation = ssTabOrientationLeft, .TabHeight, 0) - 400
            Frame1.Height = .Height - IIf(.TabOrientation = ssTabOrientationTop, .TabHeight, 0) - 400
        End With
       
        With tdbgCostos
            '*** REDIMENSIONAR CUADRICULA DE LISTADO
            .Width = Frame1.Width - .Left - 500
            .Height = Frame1.Height - .Top - 200
        End With
        
        '*** REDIMENSIONAR DETALLE
        Frame2.Height = Frame1.Height
        Frame2.Width = Frame1.Width
        
        tbrOpciones.Width = Me.Width
    End If
Exit Sub
errHand:
End Sub

Private Sub optContabilidad_Click(Index As Integer)

    If gsMonedaNac = "" Then
        MsgBox "Seleccione una moneda Nacional", vbOKOnly + vbInformation, "Aviso..."
        Exit Sub
    End If

    If Index = 0 Then
'        If gsMonedaExt <> "" Then
'            MsgBox "Primero desactive la moneda Extranjera", vbOKOnly + vbInformation, "Aviso..."
'            optContabilidad(1).Value = True
'            'Cancel = 1
'        End If
    Else
        If gsMonedaNac = "" Or gsMonedaExt = "" Then
            MsgBox "Primero seleccione una moneda Extranjera", vbOKOnly + vbInformation, "Aviso..."
            optContabilidad(0).Value = True
            'Cancel = 1
        End If
    End If
    
    If ValidaProcesoConversion = True And Index = 0 Then
        Mensajes "No se puede cambiar el tipo de contabilidad," & Salto(1) & "por que se ha realizado el proceso de conversion anteriormente"
        optContabilidad(1).Value = True
    End If

End Sub

Private Sub chkExtranjera_Click()
    If chkExtranjera.Value = vbChecked Then
        chkNacional.Value = vbUnchecked
        'Mensajes "Este registro será la Moneda Extranjera por defecto.", vbInformation
    End If
End Sub

Private Sub chkExtranjera_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then pSendKeys "{TAB}"
    If KeyAscii = 48 Then chkExtranjera.Value = 0
    If KeyAscii = 49 Then chkExtranjera.Value = 1
End Sub

Private Sub chkNacional_Click()
    If chkNacional.Value = vbChecked Then
        chkExtranjera.Value = vbUnchecked
        'Mensajes "Este registro será la Moneda Nacional por defecto.", vbInformation
    End If
End Sub

Private Sub chkNacional_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then pSendKeys "{TAB}"
    If KeyAscii = 48 Then chkNacional.Value = 0
    If KeyAscii = 49 Then chkNacional.Value = 1
End Sub



Private Sub SSTCentroCosto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        pSetFocus tdbtCodigo
    End If
End Sub

Private Sub Imprimir()
    Dim matriz(15) As Variant
    Dim Titulo As String
    Titulo = "Reporte de Tipo de Moneda"
    Titulo = UCase(Titulo)
    matriz(0) = "@NomEmp;" & gsEmpresaNom & ";True"
    matriz(1) = "@Titulo00;" & Titulo & ";True"
    matriz(2) = "@Titulo01;;True"
    matriz(3) = "@Titulo02;CODIGO;True"
    matriz(4) = "@Titulo03;DESCRIPCION;True"
    matriz(5) = "@Titulo04;ABREV;True"
    matriz(6) = "@Titulo05;NAC;True"
    matriz(7) = "@Titulo06;EXT;True"
    matriz(8) = "@Titulo07;COD.SUNAT;True"
    matriz(9) = "@Tipo;TIPOS_MON;True"
    matriz(10) = "@Emp_cCodigo;" & gsEmpresa & ";True"
    matriz(11) = "@Pan_cAnio;" & gsAnio & ";True"
    matriz(12) = "@Per_cPeriodo;" & gsPeriodo & ";True"
    matriz(13) = "@Aux;;True"
    matriz(14) = "@EMPRESA;" & gsEmpresaNom & ";True"
    matriz(15) = "@RUC;" & "RUC : " & gsRUC & ";True"
    
    Dim formulas(0) As Variant
    AbreReporteParam gsDSN, Me, rutaReportes & "RptEstandar.rpt", crptToWindow, Titulo, "", matriz(), formulas()
    
End Sub
Private Sub tbrOpciones_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim respuesta As String
    Select Case Button.Index
        Case 1: ManNuevo
        Case 2: VerDatos
        Case 3: Grabar
                SeteaBarraHerramientas Me.tbrOpciones, gsGrupo
                
        Case 4: Borrar
        Case 5: Editar
        Case 6: Imprimir
        Case 7
            If SSTCentroCosto.TabEnabled(1) = False Then ' *** Grabar
'                respuesta = MsgBox("Esta seguro que desea salir del formulario actual", vbYesNo+ vbQuestion , "Confirmar Salir")
'                If respuesta = vbYes Then Unload Me
                Unload Me
            Else
                respuesta = MsgBox("Desea cancelar la siguiente operación", vbYesNo + vbQuestion, "Confirmar Cancelar")
                If respuesta = vbYes Then
                    Call Cancelar
                    SeteaBarraHerramientas Me.tbrOpciones, gsGrupo
                End If
                    
            End If
    End Select
End Sub

Private Sub Borrar()
    ' *** Eliminar los datos; segun el q esta seleccionado
    Dim respuesta As String
    If Trim(tdbgCostos.Columns(0).Value) <> "" Then
        ' ***
        respuesta = MsgBox("Desea eliminar definitivamente el registro Seleccionado", vbYesNo + vbQuestion + vbDefaultButton2, "Confirmar Eliminar Registro")
        If respuesta = vbYes Then
            Dim clsMante As clsMantoTablas
            Set clsMante = New clsMantoTablas
            ' *** Eliminando la Cuenta
            Screen.MousePointer = vbHourglass
            Call CargaArregloMnt
            lArrMnt(0) = "ELIMINAR"                     ' Accion
            lArrMnt(2) = tdbgCostos.Columns(0).Value    ' Codigo
            If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaTipoMoneda", lArrMnt(), True) = False Then
                Mensajes "El proceso no se ha realizado. Verificar...", vbInformation
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            Call CargaTabla
            Screen.MousePointer = vbDefault
            Mensajes "Registro ha sido eliminado", vbInformation
        End If
    Else
        Mensajes "Debe Seleccionar el registro a eliminar...", vbInformation
    End If
    ' ***
End Sub

Private Sub VerDatos()
    Call CargaDatosRegistro
    If lRegElim = False Then
        lblMante = "VER REGISTRO"
        SSTCentroCosto.TabEnabled(1) = True
        SSTCentroCosto.TabEnabled(0) = False
        SSTCentroCosto.Tab = 1
        tbrOpciones.Buttons(4).Enabled = False  ' *** Borrar
        tbrOpciones.Buttons(7).Image = 8
        lTipoMnt = "EDITAR"
        Call AseguraControl(Me, True)
    Else
        lRegElim = False
    End If
End Sub

Private Sub Editar()
    MonActNac = tdbgCostos.Columns(3).Value
    MonActExt = tdbgCostos.Columns(4).Value

    Call CargaDatosRegistro
    If lRegElim = False Then
        
        lTipoMnt = "EDITAR"
        If Me.lblMante = "VER REGISTRO" Then Call AseguraControl(Me, False)
        Call HabilitaControl(Me)
        lblMante = "MODIFICANDO REGISTRO"
        Call TabMantenimiento(True)
        
        '---------------- CAMBIA LOS CHECKS DE TIPO MONEDA -------------------'
        ActivaCheckMonedas
        '----------------------------------------------------------------------'
        pSetFocus tdbtCodSunat
        
    Else
        lRegElim = False
    End If
End Sub

Private Sub ActivaCheckMonedas()
        If CE(gsMonedaNac) = "" Then
            Me.chkNacional.Enabled = True
        Else
            Me.chkNacional.Enabled = False
        End If
        
        If CE(gsMonedaExt) = "" Then
            Me.chkExtranjera.Enabled = True
        Else
            Me.chkExtranjera.Enabled = False
        End If
        
        If tdbtCodigo.Text = gsMonedaNac Then
            Me.chkNacional.Enabled = True
        End If
        
        If tdbtCodigo.Text = gsMonedaExt Then
            Me.chkExtranjera.Enabled = True
        End If
End Sub
Private Sub ManNuevo()
    MonActNac = "0"
    MonActExt = "0"

    lTipoMnt = "INSERTAR"
    Call LimpiaTexto(Me)
    Call HabilitaControl(Me)
    ' ***
    lblMante = "NUEVO REGISTRO"
    Call TabMantenimiento(True)
    pSetFocus tdbtCodigo
    chkExtranjera.Value = 0
    chkNacional.Value = 0
    ActivaCheckMonedas
End Sub

Private Sub TabMantenimiento(Valor As Boolean)
    SSTCentroCosto.TabEnabled(1) = Valor
    SSTCentroCosto.TabEnabled(0) = Not Valor
    If Valor = True Then SSTCentroCosto.Tab = 1
    If Valor = False Then SSTCentroCosto.Tab = 0
    tbrOpciones.Buttons(1).Enabled = Not Valor  ' *** Nuevo
    tbrOpciones.Buttons(2).Enabled = Not Valor  ' *** Buscar
    'tbrOpciones.Buttons(3).Enabled = valor      ' *** Grabar
    tbrOpciones.Buttons(4).Enabled = Not Valor  ' *** Borrar
    tbrOpciones.Buttons(5).Enabled = Not Valor  ' *** Editar
    If Valor = True Then
        tbrOpciones.Buttons(7).Image = 8
    Else
        tbrOpciones.Buttons(7).Image = 7
    End If
End Sub

Private Sub Cancelar()
    If Me.lblMante = "VER REGISTRO" Then
        Call AseguraControl(Me, False)
    Else
        Call HabilitaControl(Me)
    End If
    Call TabMantenimiento(False)
    pSetFocus tdbgCostos
End Sub

Private Function ConsultaAño(Tipo As String, año As String, TipoMon As String) As Boolean
    Dim rsDatos As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Dim sqlDatos As String
    Dim arrDatos() As Variant
    
    Set clDatos = New clsMantoTablas
    sqlDatos = "spCn_GrabaAnio '" & Tipo & "', '" & gsEmpresa & "', '" & gsAnio & "', '" & TipoMon & "', ''"
    arrDatos = Array(sqlDatos)
    Set rsDatos = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If NE(rsDatos(0).Value) = 0 Then
        ConsultaAño = False
    Else
        ConsultaAño = True
    End If
    Call CerrarRecordSet(rsDatos)
    Set clDatos = Nothing
End Function

Private Sub GrabarMoneda()
    Dim clsMante As clsMantoTablas
    Dim i As Integer
    Dim condicion As Boolean
    If validarDatos = False Then Exit Sub
    Set clsMante = New clsMantoTablas
    
    On Local Error GoTo ErrorEjecucion
    
    If ValidaCodSunat = False Then Exit Sub
    
    CargaArregloMnt
    'MONEDA NACIONAL
    If MonActNac = "1" Then
        If Me.chkNacional.Value = vbUnchecked Then
            'BUSCA MOVIMIENTO
            If ConsultaAño("EXISTEMOVMON", gsAnio, "N") = True Then
                Mensajes "No se puede cambiar el valor de la moneda NACIONAL ya que tiene movimiento en el año.", vbOKOnly + vbInformation
                Me.chkNacional.Value = vbChecked
                Exit Sub
            End If
            
        End If
    End If
    
    'MONEDA EXTRANJERA
    If MonActExt = "1" Then
        If Me.chkExtranjera.Value = vbUnchecked Then
            'BUSCA MOVIMIENTO
            If ConsultaAño("EXISTEMOVMON", gsAnio, "E") = True Then
                Mensajes "No se puede cambiar el valor de la moneda EXTRANJERA ya que tiene movimiento en el año.", vbOKOnly + vbInformation
                Me.chkExtranjera.Value = vbChecked
                Exit Sub
            End If
            
        End If
    End If
    
    If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaTipoMoneda", lArrMnt(), True) = False Then
        Mensajes "El proceso no ha concluido. Verificar...", vbInformation
        Exit Sub
    End If
    
    Call Cancelar
    CargaTabla
    
    '------------- GRABA E IDENTIFICA CUAL ES LA MON NAC Y EXT ------------'
    If Not lrsTabla Is Nothing Then
        gsMonedaNac = ""
        gsMonedaExt = ""
        lrsTabla.MoveFirst
        Do While Not lrsTabla.EOF
            If lrsTabla!Mon_cMNac = "1" Then
                gsMonedaNac = lrsTabla!Mon_cCodigo
                gsNombreMonedaNac = lrsTabla!Mon_cNombreLargo
            End If
            If lrsTabla!Mon_cMExt = "1" Then
                gsMonedaExt = lrsTabla!Mon_cCodigo
                gsNombreMonedaExt = lrsTabla!Mon_cNombreLargo
            End If

            lrsTabla.MoveNext
        Loop
    End If
    
    '----------- Buscar la moneda creada y posicionarse ---------------------'
    Dim Valor As Integer
    Valor = BuscarCadRs(tdbtCodigo, lrsTabla, 1)
    If Valor = 0 Then lrsTabla.MoveFirst
    '------------------------------------------------------------------------'
    'Mensajes "Los datos se grabaron con exito...", vbInformation
    pSetFocus tdbgCostos
    Exit Sub
ErrorEjecucion:
    Mensajes Str(Err.Number) & Err.Description, vbInformation
End Sub

Private Sub GrabarTipoContabilidad()
    '------------- GRABA SI LA CONTABILIDAD ES BIMONEDA  ------------'
    Dim sql As String
    Dim Moneda As String
    
    If optContabilidad(0).Value = True Then Moneda = 0
    If optContabilidad(1).Value = True Then Moneda = 1
    
    
    sql = "UPDATE EMPRESA SET emp_bymoneda='" & Moneda & "', Emp_cUserModifica='" & gsUsuario & "' " & _
          "Where EMP_CCODIGO ='" & gsEmpresa & "' and Emp_cCodSuc ='" & gsSucursal & "'"
    
    If EjecutaQuery(sql) < 1 Then
        Mensajes "El proceso no ha concluido. Verificar...", vbInformation
    Else
        Mensajes "Se grabo el tipo de contabilidad. Correctamente...", vbInformation
        If optContabilidad(0).Value = True Then gsByMoneda = 0
        If optContabilidad(1).Value = True Then gsByMoneda = 1
        
    End If

End Sub

Private Sub Grabar()
    Select Case SSTCentroCosto.Tab
           Case 0: GrabarTipoContabilidad
           Case 1: GrabarMoneda
    End Select
End Sub

Private Function validarDatos() As Boolean
    validarDatos = False
    ' *** Validar q los datos necesarios esten ingresados
    If TextoLleno2(Me.tdbtCodigo, "Codigo") = False Then Exit Function
    If TextoLleno2(Me.tdbtCodSunat, "Codigo Sunat") = False Then Exit Function
    If TextoLleno2(Me.tdbtDescripcion, "Descripcion") = False Then Exit Function
    If chkNacional.Value = 1 And chkExtranjera.Value = 1 Then
        Mensajes "Solo debe elegir Moneda Nacional o Moneda Extranjera. No ambas. ", vbInformation
        pSetFocus chkNacional
        Exit Function
    End If
    ' ***
    validarDatos = True
End Function

Private Sub CargaArregloMnt()
    ' *** Cargar los datos a grabar en un arreglo
    ReDim lArrMnt(9) As Variant
    lArrMnt(0) = lTipoMnt                ' Accion
    lArrMnt(1) = gsEmpresa               ' Empresa
    lArrMnt(2) = tdbtCodigo              ' Codigo
    lArrMnt(3) = tdbtDescripcion         ' Nombre
    lArrMnt(4) = tdbtDescripcionCorta    ' Nombre
    ' *** Si es Nacional o Extranjera=
    lArrMnt(5) = IIf(chkNacional.Value = 1, "1", "0")
    lArrMnt(6) = IIf(chkExtranjera.Value = 1, "1", "0")
    lArrMnt(7) = "A"                     ' Estado
    lArrMnt(8) = gsUsuario               ' Usuario
    lArrMnt(9) = CE(tdbtCodSunat.Text)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim respuesta As String
    Select Case KeyCode
        Case 27:
            If SSTCentroCosto.TabEnabled(1) = False Then ' *** Grabar
'                respuesta = MsgBox("Esta seguro que desea salir del formulario actual", vbYesNo+ vbQuestion , "Confirmar Salir")
'                If respuesta = vbYes Then Unload Me
                Unload Me
            Else
                respuesta = MsgBox("Desea cancelar la siguiente operación", vbYesNo + vbQuestion, "Confirmar Cancelar")
                If respuesta = vbYes Then Call Cancelar
            End If
        Case 113: If tbrOpciones.Buttons(1).Enabled Then ManNuevo
        Case 114: If tbrOpciones.Buttons(2).Enabled Then VerDatos
        Case 115: If tbrOpciones.Buttons(3).Enabled Then Grabar
        Case 116: If tbrOpciones.Buttons(4).Enabled Then Borrar
        Case 117: If tbrOpciones.Buttons(5).Enabled Then Editar
        'Case 118: If tbrOpciones.Buttons(5).Enabled Then 'Imprimir
    End Select
    ' ***
End Sub

Private Sub Form_Load()
    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))

    SSTCentroCosto.Tab = 0
    
    Call Centrar_form(Me)
    
    ' *** Llenando las grillas y los combos
    Call CargaTabla
    ' *** Seteando algunas variables
    lRegElim = False
    lTipoMnt = "INSERTAR"
    SSTCentroCosto.TabEnabled(1) = False
    
    SeteaBarraHerramientas Me.tbrOpciones, gsGrupo
    
    
    Call CargaTipodecontabilidad
    
End Sub

Private Sub CargaTipodecontabilidad()
    Dim sql As String
    Dim valorDato As String

    sql = "select emp_bymoneda from EMPRESA " & _
          "Where EMP_CCODIGO ='" & gsEmpresa & "' and Emp_cCodSuc ='" & gsSucursal & "'"
    
    valorDato = ExtraeDescripcion(sql)
    
    gsByMoneda = NE(valorDato)

    If NE(valorDato) = 1 Then
        optContabilidad(0).Value = False
        optContabilidad(1).Value = True
    Else
        optContabilidad(0).Value = True
        optContabilidad(1).Value = False
    End If

    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CerrarRecordSet(lrsTabla)
    Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))
    
End Sub

Private Function BuscaMonNacExt() As Boolean
    Dim Contador As Integer
    Contador = 0
    If Not lrsTabla Is Nothing Then
        lrsTabla.MoveFirst
        Do While Not lrsTabla.EOF
            If CE(lrsTabla!Mon_cMNac) = "1" Then Contador = Contador + 1
            If CE(lrsTabla!Mon_cMExt) = "1" Then Contador = Contador + 1
            lrsTabla.MoveNext
        Loop
    End If
    If Contador > 1 Then
        BuscaMonNacExt = True
    Else
        BuscaMonNacExt = False
    End If
    
    Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))
    
End Function
Private Sub CargaTabla()
    Dim sqlSp As String
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
    Set clDatos = New clsMantoTablas
    Set lrsTabla = New ADODB.Recordset
    
    Set tdbgCostos.DataSource = Nothing
    
    sqlSp = "spCn_GrabaTipoMoneda 'SEL_ALL', '" & gsEmpresa & "', '', '', '', '', '', '', '' "
    arrDatos = Array(sqlSp)
    Set lrsTabla = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If Not lrsTabla Is Nothing Then

        lrsTabla.Sort = "Mon_cNombreLargo"
        tdbgCostos.DataSource = lrsTabla
        
    End If
End Sub

Private Sub CargaDatosRegistro()
    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Set clDatos = New clsMantoTablas
    Dim arrDatos() As Variant
    ' *** Cargando Datos de la Cuenta
    Dim sqlSp As String
    sqlSp = "spCn_GrabaTipoMoneda 'SEL_REG', '" & gsEmpresa & "', '" & tdbgCostos.Columns(0).Value & "', '', '', '', '', '', '' "
    arrDatos = Array(sqlSp)
    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If rsArreglo.State = 0 Then
        lRegElim = True
        Mensajes "No se encontro el registro. Probablemente eliminado desde otra sesion", vbInformation
        Set rsArreglo = Nothing
        Call CargaTabla
        Exit Sub
    End If
    ' *** Asignando Datos de la Cuenta
    tdbtCodigo = CE(rsArreglo!Mon_cCodigo)
    tdbtDescripcion = CE(rsArreglo!Mon_cNombreLargo)
    tdbtDescripcionCorta = CE(rsArreglo!Mon_cNombreCorto)
    chkNacional.Value = IIf(CE(rsArreglo!Mon_cMNac) = "1", 1, 0)
    chkExtranjera.Value = IIf(CE(rsArreglo!Mon_cMExt) = "1", 1, 0)
    tdbtCodSunat.Text = CE(rsArreglo!Mon_cCodSunat)
    Call CerrarRecordSet(rsArreglo)
    ' ***
End Sub

Private Sub FiltrarRecordSet()
    ' *** Filtrar segun los textos indicados
    Dim cadena As String
    Dim filtros(0) As String
    Dim i As Integer
    If lrsTabla Is Nothing Then Exit Sub
    cadena = ""
    If Trim(tdbtDescripcionBus) <> "" Then filtros(0) = "Mon_cNombreLargo like '*" & tdbtDescripcionBus & "*'"
    For i = 0 To 0
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
        lrsTabla.Filter = cadena
    Else
        lrsTabla.Filter = 0
    End If
End Sub

Private Sub tdbgCostos_GotFocus()
tdbgCostos.HighlightRowStyle = "HighlightRow"
End Sub

Private Sub tdbgCostos_HeadClick(ByVal ColIndex As Integer)
If Not lrsTabla Is Nothing Then
    If lrsTabla.RecordCount > 0 Then

        lrsTabla.Sort = tdbgCostos.Columns(ColIndex).DataField
        tdbgCostos.DataSource = lrsTabla
    End If
End If
End Sub

Private Sub tdbgCostos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Editar
End If
End Sub

Private Sub tdbgCostos_LostFocus()
tdbgCostos.HighlightRowStyle = ""
End Sub

Private Sub tdbtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    gsKey = KeyCode
    If gsKey = 219 Then
       tdbtCodigo = Replace(tdbtCodigo, "'", "")
       tdbtCodigo.SelStart = Len(tdbtCodigo)
    End If
End Sub

Private Function ValidaCodSunat() As Boolean
    If SSTCentroCosto.TabEnabled(1) = True And CE(tdbtCodSunat.Text) <> "" And (CE(tdbtCodSunat.Text) <> "9") Then
        If ExisteDato("select * from CNT_TIPO_MONEDA where  mon_ccodigo<> '" & tdbtCodigo.Text & "' and emp_ccodigo='" & gsEmpresa & "' and mon_ccodsunat='" & tdbtCodSunat.Text & "'") Then
            Mensajes "Codigo Sunat ya existe. Verifique...", vbInformation
            tdbtCodSunat.Text = ""
            pSetFocus tdbtCodSunat
            ValidaCodSunat = False
            Exit Function
        End If
    End If
    ValidaCodSunat = True
End Function

Private Sub tdbtCodSunat_LostFocus()
    ValidaCodSunat
End Sub

Private Sub tdbtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
    gsKey = KeyCode
    If gsKey = 219 Then
       tdbtDescripcion = Replace(tdbtDescripcion, "'", "")
       tdbtDescripcion.SelStart = Len(tdbtDescripcion)
    End If
End Sub

Private Sub tdbtDescripcionBus_Change()
    
    If gsKey = 219 Then
       tdbtDescripcionBus = Replace(tdbtDescripcionBus, "'", "")
       tdbtDescripcionBus.SelStart = Len(tdbtDescripcionBus)
    End If
    
    Call FiltrarRecordSet
End Sub

Private Sub tdbtCodigo_GotFocus()
    ' *** Si es modificar no editar
    If lTipoMnt = "EDITAR" Then
        tdbtCodigo.ReadOnly = True
    Else
        tdbtCodigo.ReadOnly = False
    End If
    ' ***
End Sub

Private Sub tdbtCodigo_LostFocus()
    ' *** Verificar q codigo exista
    If SSTCentroCosto.TabEnabled(1) = True And lTipoMnt = "INSERTAR" Then
        'If ExisteCodigo(tdbtCodigo) = True Then
        If ExisteRegistro(tdbtCodigo, "spCn_GrabaTipoMoneda", False) = True Then
            Mensajes "Codigo ya existe. Verifique...", vbInformation
            tdbtCodigo.Text = ""
            pSetFocus tdbtCodigo
        End If
    End If
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
'    sqlSp = "spCn_GrabaTipoMoneda 'SEL_REG', '" & gsEmpresa & "', '" & valor & "', '', '', '', '', '', '' "
'    arrDatos = Array(sqlSp)
'    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
'    If rsArreglo.State <> 0 Then
'        ExisteCodigo = True
'    End If
'    Call CerrarRecordSet(rsArreglo)
'    ' ***
'End Function

Private Sub tdbtDescripcionBus_KeyDown(KeyCode As Integer, Shift As Integer)
    gsKey = KeyCode
End Sub

Private Sub tdbtDescripcionCorta_KeyDown(KeyCode As Integer, Shift As Integer)
    gsKey = KeyCode
    If gsKey = 219 Then
       tdbtDescripcionCorta = Replace(tdbtDescripcionCorta, "'", "")
       tdbtDescripcionCorta.SelStart = Len(tdbtDescripcionCorta)
    End If
End Sub
