VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CB36-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todl7.ocx"
Begin VB.Form frmManTipoDocumento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Tipo de Documentos"
   ClientHeight    =   6150
   ClientLeft      =   4515
   ClientTop       =   2970
   ClientWidth     =   7845
   Icon            =   "frmManTipoDocumento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6150
   ScaleWidth      =   7845
   Begin TabDlg.SSTab SSTCentroCosto 
      Height          =   5595
      Left            =   240
      TabIndex        =   5
      Top             =   495
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   9869
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Consultas de Tipo de Documentos"
      TabPicture(0)   =   "frmManTipoDocumento.frx":0ECA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Mantenimiento de Tipo de Documentos"
      TabPicture(1)   =   "frmManTipoDocumento.frx":0EE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   5055
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   6945
         Begin TrueOleDBGrid70.TDBGrid tdbgCostos 
            Height          =   3795
            Left            =   225
            TabIndex        =   2
            Top             =   1080
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   6694
            _LayoutType     =   4
            _RowHeight      =   19
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Codigo "
            Columns(0).DataField=   "Tdo_cCodigo"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Nombre Largo"
            Columns(1).DataField=   "Tdo_cNombreLargo"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   16
            Columns(2)._MaxComboItems=   5
            Columns(2).ValueItems(0)._DefaultItem=   0
            Columns(2).ValueItems(0).Value=   "1"
            Columns(2).ValueItems(0).Value.vt=   8
            Columns(2).ValueItems(0).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
            Columns(2).ValueItems(0).DisplayValue(0)=   "bHQAAGYDAABCTWYDAAAAAAAANgAAACgAAAAQAAAAEQAAAAEAGAAAAAAAMAMAAAAAAAAAAAAAAAAA"
            Columns(2).ValueItems(0).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
            Columns(2).ValueItems(0).DisplayValue(2)=   "//////////////////9SpUoAlAhrtWP/////////////////////////////////////////////"
            Columns(2).ValueItems(0).DisplayValue(3)=   "//////8YtSkAvSEAlACMvXv///////////////////////////////////////////9rtWMAvSEA"
            Columns(2).ValueItems(0).DisplayValue(4)=   "xikApQAxnDH///////////////////////////////////////////8AnBAAzjEAxikArRAAlACl"
            Columns(2).ValueItems(0).DisplayValue(5)=   "xpT///////////////////////////////////9SpUoAzjEAxikA/2MAzjEAnAAAjAD/////////"
            Columns(2).ValueItems(0).DisplayValue(6)=   "//////////////////////////8YtSkpzloA/2MA/2MAvSEAxikAlACMvXv/////////////////"
            Columns(2).ValueItems(0).DisplayValue(7)=   "//////////////8YxkIA/2MA/2NSpUpSpUoAxikApQAxnDH/////////////////////////////"
            Columns(2).ValueItems(0).DisplayValue(8)=   "//////8ArSEArSH///////8ArRgAxikAlAClxpT/////////////////////////////////////"
            Columns(2).ValueItems(0).DisplayValue(9)=   "//////////8xtUIAxikAnAAAjAD/////////////////////////////////////////////////"
            Columns(2).ValueItems(0).DisplayValue(10)=   "//8AtSEAxikAlACMvXv///////////////////////////////////////////////9SpUoAxikp"
            Columns(2).ValueItems(0).DisplayValue(11)=   "rTkxtUL///////////////////////////////////////////////////8prUpa56UprTmMvXv/"
            Columns(2).ValueItems(0).DisplayValue(12)=   "//////////////////////////////////////////////8xtUIA1kKMvXv/////////////////"
            Columns(2).ValueItems(0).DisplayValue(13)=   "//////////////////////////////////////+lxpT/////////////////////////////////"
            Columns(2).ValueItems(0).DisplayValue(14)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(2).ValueItems(0).DisplayValue(15)=   "//////////////////////////////8="
            Columns(2).ValueItems(0).DisplayValue.vt=   9
            Columns(2).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
            Columns(2).ValueItems(1)._DefaultItem=   0
            Columns(2).ValueItems(1).Value=   "0"
            Columns(2).ValueItems(1).Value.vt=   8
            Columns(2).ValueItems(1).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
            Columns(2).ValueItems(1).DisplayValue(0)=   "bHQAAGYDAABCTWYDAAAAAAAANgAAACgAAAAQAAAAEQAAAAEAGAAAAAAAMAMAAAAAAAAAAAAAAAAA"
            Columns(2).ValueItems(1).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
            Columns(2).ValueItems(1).DisplayValue(2)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(2).ValueItems(1).DisplayValue(3)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(2).ValueItems(1).DisplayValue(4)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(2).ValueItems(1).DisplayValue(5)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(2).ValueItems(1).DisplayValue(6)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(2).ValueItems(1).DisplayValue(7)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(2).ValueItems(1).DisplayValue(8)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(2).ValueItems(1).DisplayValue(9)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(2).ValueItems(1).DisplayValue(10)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(2).ValueItems(1).DisplayValue(11)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(2).ValueItems(1).DisplayValue(12)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(2).ValueItems(1).DisplayValue(13)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(2).ValueItems(1).DisplayValue(14)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(2).ValueItems(1).DisplayValue(15)=   "//////////////////////////////8="
            Columns(2).ValueItems(1).DisplayValue.vt=   9
            Columns(2).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
            Columns(2).ValueItems.Count=   2
            Columns(2).Caption=   "Daot"
            Columns(2).DataField=   "Tdo_cDaot"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   16
            Columns(3)._MaxComboItems=   5
            Columns(3).ValueItems(0)._DefaultItem=   0
            Columns(3).ValueItems(0).Value=   "+"
            Columns(3).ValueItems(0).Value.vt=   8
            Columns(3).ValueItems(0).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
            Columns(3).ValueItems(0).DisplayValue(0)=   "bHQAADYDAABCTTYDAAAAAAAANgAAACgAAAAQAAAAEAAAAAEAGAAAAAAAAAMAAAAAAAAAAAAAAAAA"
            Columns(3).ValueItems(0).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
            Columns(3).ValueItems(0).DisplayValue(2)=   "//////////////////9rrYQhhCkhhClrrYT/////////////////////////////////////////"
            Columns(3).ValueItems(0).DisplayValue(3)=   "//////9jpWOU3ow5tVIhhCn///////////////////////////////////////////////9jpWOU"
            Columns(3).ValueItems(0).DisplayValue(4)=   "3ow5tVIhhCn///////////////////////////////////////////////9jpWOU3ow5tVIhhCn/"
            Columns(3).ValueItems(0).DisplayValue(5)=   "//////////////////////////////////////////////9jpWOU3ow5tVIhhCn/////////////"
            Columns(3).ValueItems(0).DisplayValue(6)=   "//////////////9rrYQhhCkhhCkhhCkhhCkhhCmU3ow5tVIhhCkhhCkhhCkhhCkhhClrrYT/////"
            Columns(3).ValueItems(0).DisplayValue(7)=   "//9jpWM5tVI5tVI5tVI5tVI5tVI5tVI5tVI5tVI5tVI5tVI5tVI5tVIhhCn///////9jpWOU3oyU"
            Columns(3).ValueItems(0).DisplayValue(8)=   "3oyU3oyU3oyU3oyU3ow5tVKU3oyU3oyU3oyU3oyU3owhhCn///////9rrYRjpWNjpWNjpWNjpWNj"
            Columns(3).ValueItems(0).DisplayValue(9)=   "pWOU3ow5tVIhhCkhhCkhhCkhhCkhhClrrYT///////////////////////////9jpWOU3ow5tVIh"
            Columns(3).ValueItems(0).DisplayValue(10)=   "hCn///////////////////////////////////////////////9jpWOU3ow5tVIhhCn/////////"
            Columns(3).ValueItems(0).DisplayValue(11)=   "//////////////////////////////////////9jpWOU3ow5tVIhhCn/////////////////////"
            Columns(3).ValueItems(0).DisplayValue(12)=   "//////////////////////////9jpWOU3ow5tVIhhCn/////////////////////////////////"
            Columns(3).ValueItems(0).DisplayValue(13)=   "//////////////9rrYRjpWNjpWNrrYT/////////////////////////////////////////////"
            Columns(3).ValueItems(0).DisplayValue(14)=   "//////////////////////////////////////////8="
            Columns(3).ValueItems(0).DisplayValue.vt=   9
            Columns(3).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
            Columns(3).ValueItems(1)._DefaultItem=   0
            Columns(3).ValueItems(1).Value=   "-"
            Columns(3).ValueItems(1).Value.vt=   8
            Columns(3).ValueItems(1).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
            Columns(3).ValueItems(1).DisplayValue(0)=   "bHQAADYDAABCTTYDAAAAAAAANgAAACgAAAAQAAAAEAAAAAEAGAAAAAAAAAMAAAAAAAAAAAAAAAAA"
            Columns(3).ValueItems(1).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
            Columns(3).ValueItems(1).DisplayValue(2)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(3).ValueItems(1).DisplayValue(3)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(3).ValueItems(1).DisplayValue(4)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(3).ValueItems(1).DisplayValue(5)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(3).ValueItems(1).DisplayValue(6)=   "//////////////9rhMYAIaUAIaUAIaUAIaUAIaUAIaUAIaUAIaUAIaUAIaUAIaUAIaVrhMb/////"
            Columns(3).ValueItems(1).DisplayValue(7)=   "//8AIaWUlPcAKecAKecAKecAKecAKecAKecAKecAKecAKecAKecAKecAIaX///////8AIaW1xv+c"
            Columns(3).ValueItems(1).DisplayValue(8)=   "vf+cvf+ctf+ctf+ctf9jjPdjjPdjjPdjjPdSa/dSa/cAIaX///////9rhMYAIaUAIaUAIaUAIaUA"
            Columns(3).ValueItems(1).DisplayValue(9)=   "IaUAIaUAIaUAIaUAIaUAIaUAIaUAIaVrhMb/////////////////////////////////////////"
            Columns(3).ValueItems(1).DisplayValue(10)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(3).ValueItems(1).DisplayValue(11)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(3).ValueItems(1).DisplayValue(12)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(3).ValueItems(1).DisplayValue(13)=   "////////////////////////////////////////////////////////////////////////////"
            Columns(3).ValueItems(1).DisplayValue(14)=   "//////////////////////////////////////////8="
            Columns(3).ValueItems(1).DisplayValue.vt=   9
            Columns(3).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
            Columns(3).ValueItems.Count=   2
            Columns(3).Caption=   "Ope"
            Columns(3).DataField=   "Tdo_cNatDaot"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Nombre Corto"
            Columns(4).DataField=   "Tdo_cNombreCorto"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   5
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=5"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1085"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1005"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=532"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=5821"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=5741"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=532"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=1588"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1508"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=529"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=1852"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1773"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=532"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=1879"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=1799"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=532"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
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
            _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=2,.bgcolor=&HFFFFFF&"
            _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.bgcolor=&HFFFFFF&"
            _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
            _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
            _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=28,.parent=13"
            _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=14"
            _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=15"
            _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=17"
            _StyleDefs(58)  =   "Named:id=33:Normal"
            _StyleDefs(59)  =   ":id=33,.parent=0"
            _StyleDefs(60)  =   "Named:id=34:Heading"
            _StyleDefs(61)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(62)  =   ":id=34,.wraptext=-1"
            _StyleDefs(63)  =   "Named:id=35:Footing"
            _StyleDefs(64)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(65)  =   "Named:id=36:Selected"
            _StyleDefs(66)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(67)  =   "Named:id=37:Caption"
            _StyleDefs(68)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(69)  =   "Named:id=38:HighlightRow"
            _StyleDefs(70)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(71)  =   "Named:id=39:EvenRow"
            _StyleDefs(72)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(73)  =   "Named:id=40:OddRow"
            _StyleDefs(74)  =   ":id=40,.parent=33"
            _StyleDefs(75)  =   "Named:id=41:RecordSelector"
            _StyleDefs(76)  =   ":id=41,.parent=34"
            _StyleDefs(77)  =   "Named:id=42:FilterBar"
            _StyleDefs(78)  =   ":id=42,.parent=33"
         End
         Begin TDBText6Ctl.TDBText tdbtDescripcionBus 
            Height          =   315
            Left            =   1830
            TabIndex        =   0
            Top             =   585
            Width           =   4860
            _Version        =   65536
            _ExtentX        =   8572
            _ExtentY        =   556
            Caption         =   "frmManTipoDocumento.frx":0F02
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManTipoDocumento.frx":0F6E
            Key             =   "frmManTipoDocumento.frx":0F8C
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
            Caption         =   "Nombre Largo"
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
            Left            =   285
            TabIndex        =   12
            Top             =   630
            Width           =   1200
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Filtrar Datos"
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
            Index           =   6
            Left            =   285
            TabIndex        =   11
            Top             =   330
            Width           =   1035
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3945
         Left            =   -74760
         TabIndex        =   6
         Top             =   360
         Width           =   6780
         Begin VB.Frame Frame3 
            Caption         =   "DAOT"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1410
            Left            =   225
            TabIndex        =   14
            Top             =   2250
            Width           =   6090
            Begin VB.CheckBox chkDaot 
               Alignment       =   1  'Right Justify
               Caption         =   "Incluir en DAOT"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1485
               TabIndex        =   15
               Tag             =   "_"
               Top             =   315
               Width           =   2430
            End
            Begin TrueOleDBList70.TDBCombo tdbcDAOT 
               Height          =   300
               Left            =   3690
               TabIndex        =   17
               Tag             =   "_"
               Top             =   765
               Width           =   1620
               _ExtentX        =   2858
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
               _PropDict       =   $"frmManTipoDocumento.frx":0FDE
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
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Tipo de Operacion DAOT"
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
               Left            =   1485
               TabIndex        =   16
               Top             =   825
               Width           =   2040
            End
         End
         Begin TDBText6Ctl.TDBText tdbtCodigo 
            Height          =   315
            Left            =   1755
            TabIndex        =   1
            Tag             =   "_"
            Top             =   825
            Width           =   735
            _Version        =   65536
            _ExtentX        =   1296
            _ExtentY        =   556
            Caption         =   "frmManTipoDocumento.frx":1065
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManTipoDocumento.frx":10D1
            Key             =   "frmManTipoDocumento.frx":10EF
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
         Begin TDBText6Ctl.TDBText tdbtDescripcion 
            Height          =   315
            Left            =   1755
            TabIndex        =   3
            Tag             =   "_"
            Top             =   1275
            Width           =   4530
            _Version        =   65536
            _ExtentX        =   7990
            _ExtentY        =   556
            Caption         =   "frmManTipoDocumento.frx":1141
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManTipoDocumento.frx":11AD
            Key             =   "frmManTipoDocumento.frx":11CB
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
            Left            =   1755
            TabIndex        =   4
            Tag             =   "_"
            Top             =   1725
            Width           =   4530
            _Version        =   65536
            _ExtentX        =   7990
            _ExtentY        =   556
            Caption         =   "frmManTipoDocumento.frx":121D
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmManTipoDocumento.frx":1289
            Key             =   "frmManTipoDocumento.frx":12A7
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
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Nombre Corto"
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
            TabIndex        =   13
            Top             =   1725
            Width           =   1170
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Nombre Largo"
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
            TabIndex        =   9
            Top             =   1275
            Width           =   1200
         End
         Begin VB.Label Label2 
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
            Index           =   8
            Left            =   240
            TabIndex        =   8
            Top             =   825
            Width           =   600
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
            TabIndex        =   7
            Top             =   360
            Width           =   1515
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
            Picture         =   "frmManTipoDocumento.frx":12F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoDocumento.frx":16D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoDocumento.frx":1AAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoDocumento.frx":1E87
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoDocumento.frx":2261
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoDocumento.frx":263B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoDocumento.frx":2A15
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoDocumento.frx":2DEF
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
            Picture         =   "frmManTipoDocumento.frx":3E09
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoDocumento.frx":3F63
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoDocumento.frx":40BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoDocumento.frx":4217
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoDocumento.frx":4371
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoDocumento.frx":44CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoDocumento.frx":4625
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoDocumento.frx":477F
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoDocumento.frx":48D9
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
            Picture         =   "frmManTipoDocumento.frx":4A33
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoDocumento.frx":4FCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoDocumento.frx":5567
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoDocumento.frx":5B01
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoDocumento.frx":609B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoDocumento.frx":6635
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoDocumento.frx":6BCF
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoDocumento.frx":7169
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManTipoDocumento.frx":7703
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrOpciones 
      Height          =   330
      Left            =   0
      TabIndex        =   18
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
Attribute VB_Name = "frmManTipoDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lArrMnt() As Variant        ' *** Arreglo para los mantenimientos
Dim lTipoMnt As String          ' *** Tipo de Mantenimiento (insert/update/delete)
Dim lrsTabla As Recordset
Dim lRegElim As Boolean         ' *** Verifica si registro ha sido eliminado desde otra sesion

Dim gsGrupo As String
Public Property Let Grupo(ByVal Grupo As String)
     gsGrupo = Grupo
End Property

Private Sub chkDaot_Click()
    CambiaValorDaot
End Sub

Private Sub CambiaValorDaot()
    If chkDaot.Value = vbChecked Then
        'tdbcDAOT.BoundText = "+"
        tdbcDAOT.Enabled = True
    Else
        tdbcDAOT.BoundText = ""
        tdbcDAOT.Enabled = False
    End If
End Sub

Private Sub Form_Resize()
On Error GoTo errHand
    If Me.WindowState <> vbMinimized Then
        
        '*** REDIMENSIONAR SST
        With SSTCentroCosto
            .Width = Me.Width - .Left - 300
            .Height = Me.Height - .Top - 400
            '*** REDIMENSIONAR FRAME PRINCIPAL
            Frame1.Width = .Width - IIf(.TabOrientation = ssTabOrientationLeft, .TabHeight, 0) - 500
            Frame1.Height = .Height - IIf(.TabOrientation = ssTabOrientationTop, .TabHeight, 0) - 300
        End With
       
        With tdbgCostos
            '*** REDIMENSIONAR CUADRICULA DE LISTADO
            .Width = Frame1.Width - .Left - 500
            .Height = Frame1.Height - .Top - 200
        End With
        
        '*** REDIMENSIONAR DETALLE
        Frame2.Top = 360
        Frame2.Left = 240
        Frame2.Height = Frame1.Height
        Frame2.Width = Frame1.Width
        
        tbrOpciones.Width = Me.Width
    End If
Exit Sub
errHand:
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
            If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaTipoDocumento", lArrMnt(), True) = False Then
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
        CambiaValorDaot
    Else
        lRegElim = False
    End If
End Sub

Private Sub Editar()
    Call CargaDatosRegistro
    If lRegElim = False Then
        lTipoMnt = "EDITAR"
        If Me.lblMante = "VER REGISTRO" Then Call AseguraControl(Me, False)
        Call HabilitaControl(Me)
        lblMante = "MODIFICANDO REGISTRO"
        Call TabMantenimiento(True)
        pSetFocus tdbtDescripcion
        CambiaValorDaot
    Else
        lRegElim = False
    End If
End Sub

Private Sub ManNuevo()
    lTipoMnt = "INSERTAR"
    Call LimpiaTexto(Me)
    Call HabilitaControl(Me)
    ' ***
    lblMante = "NUEVO REGISTRO"
    Call TabMantenimiento(True)
    pSetFocus tdbtCodigo
    chkDaot.Value = 0
    CambiaValorDaot
End Sub

Private Sub TabMantenimiento(Valor As Boolean)
    SSTCentroCosto.TabEnabled(1) = Valor
    SSTCentroCosto.TabEnabled(0) = Not Valor
    If Valor = True Then SSTCentroCosto.Tab = 1
    If Valor = False Then SSTCentroCosto.Tab = 0
    tbrOpciones.Buttons(1).Enabled = Not Valor  ' *** Nuevo
    tbrOpciones.Buttons(2).Enabled = Not Valor  ' *** Buscar
    tbrOpciones.Buttons(3).Enabled = Valor      ' *** Grabar
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

Private Sub Imprimir()
    Dim matriz(10) As Variant
    Dim Tipo As String
            Tipo = "IMPRIMIR"
            matriz(0) = "@Accion;" & Tipo & ";True"
            matriz(1) = "@Emp_cCodigo;" & gsEmpresa & ";True"
            matriz(2) = "@Tdo_cCodigo;;True"
            matriz(3) = "@Tdo_cNombreLargo;;True"
            matriz(4) = "@Tdo_cNombreCorto;;True"
            matriz(5) = "@Tdo_cEstado;;True"
            matriz(6) = "@Tdo_cDaot;;True"
            matriz(7) = "@Tdo_cNatDaot;;True"
            matriz(8) = "@Tdo_cUserCrea;;True"
       
            matriz(9) = "@EMPRESA;" & gsEmpresaNom & ";True"
            matriz(10) = "@RUC;" & "RUC : " & gsRUC & ";True"
            
    Dim formulas(0) As Variant
       AbreReporteParam gsDSN, Me, rutaReportes & "RptTipoDocumento.rpt", crptToWindow, "Tipos de Comprobantes", "", matriz(), formulas()
       
End Sub

Private Sub Grabar()
    Dim clsMante As clsMantoTablas
    Dim i As Integer
    Dim condicion As Boolean
    If validarDatos = False Then Exit Sub
    Set clsMante = New clsMantoTablas
    ' *** Grabando Centro de Costo
    On Local Error GoTo ErrorEjecucion
    CargaArregloMnt
    If clsMante.MantenimientoDeTablas(gsCadenaConexion, "spCn_GrabaTipoDocumento", lArrMnt(), True) = False Then
        Mensajes "El proceso no ha concluido. Verificar...", vbInformation
        Exit Sub
    End If
    Call Cancelar
    CargaTabla
    ' *** Buscar el Costo creado y posicionarse alli
    Dim Valor As Integer
    Valor = BuscarCadRs(tdbtCodigo, lrsTabla, 1)
    If Valor = 0 Then lrsTabla.MoveFirst
    ' ***
    Mensajes "Los datos se grabaron con exito...", vbInformation
    tdbgCostos.HighlightRowStyle = "HighlightRow"
    pSetFocus tdbgCostos
    Exit Sub
ErrorEjecucion:
    Mensajes Str(Err.Number) & Err.Description, vbInformation
End Sub

Private Function validarDatos() As Boolean
    validarDatos = False
    ' *** Validar q los datos necesarios esten ingresados
    If TextoLleno2(Me.tdbtCodigo, "Codigo") = False Then Exit Function
    If TextoLleno2(Me.tdbtDescripcion, "Descripcion") = False Then Exit Function
    ' ***
    If chkDaot.Value = vbChecked And tdbcDAOT.Text = "" Then
        Mensajes "Seleccione el operador a utilizar para el tipo de documento seleccionado"
        pSetFocus tdbcDAOT
        Exit Function
    End If
    
    validarDatos = True
End Function

Private Sub CargaArregloMnt()
    ' *** Cargar los datos a grabar en un arreglo
    ReDim lArrMnt(8) As Variant
    lArrMnt(0) = lTipoMnt           ' Accion
    lArrMnt(1) = gsEmpresa          ' Empresa
    lArrMnt(2) = tdbtCodigo         ' Codigo
    lArrMnt(3) = tdbtDescripcion    ' Nombre
    lArrMnt(4) = tdbtDescripcionCorta    ' Nombre
    lArrMnt(5) = "A"                ' Nombre Plantilla
    lArrMnt(6) = chkDaot.Value      ' Daot
    lArrMnt(7) = tdbcDAOT.BoundText    ' Operacion Daot
    lArrMnt(8) = gsUsuario          ' Usuario
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim respuesta As String
    Select Case KeyCode
        Case vbKeyEscape:
            If SSTCentroCosto.TabEnabled(1) = False Then ' *** Grabar
                Unload Me
            Else
                respuesta = MsgBox("Desea cancelar la siguiente operación", vbYesNo + vbQuestion, "Confirmar Cancelar")
                If respuesta = vbYes Then Call Cancelar
            End If
        Case vbKeyF2: If tbrOpciones.Buttons(1).Enabled Then ManNuevo
        Case vbKeyF3: If tbrOpciones.Buttons(2).Enabled Then VerDatos
        Case vbKeyF4: If tbrOpciones.Buttons(3).Enabled Then Grabar
        Case vbKeyF5: If tbrOpciones.Buttons(4).Enabled Then Borrar
        Case vbKeyF6: If tbrOpciones.Buttons(5).Enabled Then Editar
        Case vbKeyF7: If tbrOpciones.Buttons(5).Enabled Then Imprimir
    End Select
    ' ***
End Sub

Private Sub Form_Load()
    Call frmMDIConta.TabForm_CrearN(NE(Me.hwnd))

    Call Centrar_form(Me)

    ' *** Llenando las grillas y los combos
    Call LlenacomboDaot
    Call CargaTabla
    ' *** Seteando algunas variables
    lRegElim = False
    lTipoMnt = "INSERTAR"
    SSTCentroCosto.TabEnabled(1) = False
    tdbcDAOT.BoundText = ""

    SeteaBarraHerramientas Me.tbrOpciones, gsGrupo
    SSTCentroCosto.Tab = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CerrarRecordSet(lrsTabla)
    Call frmMDIConta.TabForm_CerrarN(NE(Me.hwnd))
End Sub

Private Sub CargaTabla()
    Dim sqlSp As String
    Dim clDatos As clsMantoTablas
    Dim arrDatos() As Variant
    Set clDatos = New clsMantoTablas
    Set lrsTabla = New ADODB.Recordset
    Set tdbgCostos.DataSource = Nothing
    sqlSp = "spCn_GrabaTipoDocumento 'SEL_ALL', '" & gsEmpresa & "', '', '', '', '', '', '', '' "
    arrDatos = Array(sqlSp)
    Set lrsTabla = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
    If Not lrsTabla Is Nothing Then
        ' *** Llenar grilla con el RecordSet
        lrsTabla.Sort = "Tdo_cCodigo"
        tdbgCostos.DataSource = lrsTabla
        ' ***
        Exit Sub
    End If
End Sub

Private Sub CargaDatosRegistro()
    Dim rsArreglo As New ADODB.Recordset
    Dim clDatos As clsMantoTablas
    Set clDatos = New clsMantoTablas
    Dim arrDatos() As Variant
    ' *** Cargando Datos de la Cuenta
    Dim sqlSp As String
    sqlSp = "spCn_GrabaTipoDocumento 'SEL_REG', '" & gsEmpresa & "', '" & tdbgCostos.Columns(0).Value & "', '', '', '', '', '', '' "
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
    tdbtCodigo = Trim(CE(rsArreglo!Tdo_cCodigo))
    tdbtDescripcion = CE(rsArreglo!Tdo_cNombreLargo)
    tdbtDescripcionCorta = CE(rsArreglo!Tdo_cNombreCorto)
    If CE(rsArreglo!Tdo_cDaot) = "" Then
        chkDaot = "0"
    Else
        chkDaot = NE(rsArreglo!Tdo_cDaot)
    End If
    tdbcDAOT.BoundText = ""
    If CE(rsArreglo!Tdo_cNatDaot) <> "" Then
        tdbcDAOT.BoundText = CE(rsArreglo!Tdo_cNatDaot)
        'If CE(rsArreglo!Tdo_cNatDaot) = "+" Then
        '    cmbOperacion.ListIndex = 0
        'Else
        '    cmbOperacion.ListIndex = 1
        'End If
    End If
    tdbcDAOT.ReBind
    Call CerrarRecordSet(rsArreglo)
    ' ***
End Sub

Private Sub LlenacomboDaot()
    tdbcDAOT.Clear
    
        tdbcDAOT.AddItem "+" & ";" & "(+) POSITIVO"
        tdbcDAOT.AddItem "-" & ";" & "(-) NEGATIVO"
        tdbcDAOT.Bookmark = 0
        tdbcDAOT.ListField = "column1"
        tdbcDAOT.BoundColumn = "column0"
    
End Sub

Private Sub FiltrarRecordSet()
    ' *** Filtrar segun los textos indicados
    Dim cadena As String
    Dim filtros(0) As String
    Dim i As Integer
    If lrsTabla Is Nothing Then Exit Sub
    cadena = ""
    If Trim(tdbtDescripcionBus) <> "" Then filtros(0) = "Tdo_cNombreLargo like '*" & tdbtDescripcionBus & "*'"
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

Private Sub tdbgCostos_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    'CODIGO 44 ES EL ULTIMO CODIGO DE DOC POR LA SUNAT
    'SI AUMENTA CAMBIAR ESTE NUMERO
    
    If NE(tdbgCostos.Columns(0).Value) <= 44 Or NE(tdbgCostos.Columns(0).Value) = 98 Or NE(tdbgCostos.Columns(0).Value) = 99 Then
        If IsNumeric(CE(tdbgCostos.Columns(0).Value)) Then
           tbrOpciones.Buttons(4).Enabled = False
        Else
           tbrOpciones.Buttons(4).Enabled = True
        End If
    Else
        tbrOpciones.Buttons(4).Enabled = True
    End If
    
End Sub

Private Sub tdbtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    gsKey = KeyCode
    If gsKey = 219 Then
       tdbtCodigo = Replace(tdbtCodigo, "'", "")
       tdbtCodigo.SelStart = Len(tdbtCodigo)
    End If
    
If KeyCode = 13 Then
    pSetFocus tdbtDescripcion
End If
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
        If ExisteRegistro(tdbtCodigo, "spCn_GrabaTipoDocumento", False) = True Then
            Mensajes "Codigo ya existe. Verifique...", vbInformation
            pSetFocus tdbtCodigo
        End If
'        If ExisteCodigo(tdbtCodigo) = True Then
'            Mensajes "Codigo ya existe. Verifique...", vbInformation
'            pSetFocus tdbtCodigo
'        End If
    End If
End Sub

Private Sub tdbtDescripcionBus_KeyDown(KeyCode As Integer, Shift As Integer)
    gsKey = KeyCode
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
'    sqlSp = "spCn_GrabaTipoDocumento 'SEL_REG', '" & gsEmpresa & "', '" & valor & "', '', '', '', '', '', '' "
'    arrDatos = Array(sqlSp)
'    Set rsArreglo = clDatos.ConsultaDatosTabla(gsCadenaConexion, "pSTD_EjecutaQuery", arrDatos())
'    If rsArreglo.State <> 0 Then
'        ExisteCodigo = True
'    End If
'    Call CerrarRecordSet(rsArreglo)
'    ' ***
'End Function
Private Sub tdbtDescripcionBus_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pSetFocus tdbgCostos
End If
End Sub

Private Sub tdbtDescripcionCorta_KeyDown(KeyCode As Integer, Shift As Integer)
    gsKey = KeyCode
    If gsKey = 219 Then
       tdbtDescripcionCorta = Replace(tdbtDescripcionCorta, "'", "")
       tdbtDescripcionCorta.SelStart = Len(tdbtDescripcionCorta)
    End If
End Sub
